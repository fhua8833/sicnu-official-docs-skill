#!/usr/bin/env python3
"""Deterministic official-document builders and fail-closed delivery gate.

Build only from the documented JSON contract. Never infer missing facts, fonts,
sort orders, or a human review. `repair` means regenerate from that JSON.
"""
from __future__ import annotations
import argparse
import copy
import hashlib
import json
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
import unicodedata
import uuid
import xml.etree.ElementTree as E
from datetime import date
from decimal import Decimal
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from bundled_fonts import bundle_status, font_environment, bundled_glyph_issues, bundled_pdf_font_issues

ROOT = Path(__file__).resolve().parents[1]
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
X = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
NS = {'w': W, 'x': X}
MARGINS = {'top':3.7, 'bottom':3.5, 'left':2.8, 'right':2.6, 'footer':2.75}
RULES = {
 'Title':('方正小标宋简体',22,35,'center',0),
 'Subtitle':('楷体_GB2312',18,35,'center',0),
 'Meta':('楷体_GB2312',16,28,'center',0),
 'h1':('黑体',16,28,'both',2), 'h2':('楷体_GB2312',16,28,'both',2),
 'h3':('仿宋_GB2312',16,28,'both',2), 'h4':('仿宋_GB2312',16,28,'both',2),
 'body':('仿宋_GB2312',16,28,'both',2),
 'Signature':('仿宋_GB2312',16,28,'center',0),
 'TableTitle':('方正小标宋简体',22,None,'center',0),
 'TableHead':('黑体',16,None,'center',0),
 'TableBody':('仿宋_GB2312',16,None,'center',0),
 'TableFooter':('仿宋_GB2312',14,None,'right',0),
 'Caption':('仿宋_GB2312',10.5,12,'center',0),
 'Attachment':('黑体',16,28,'left',0),
 'AttachmentList':('仿宋_GB2312',16,28,'left',0),
 'Secret':('黑体',16,28,'left',0), 'Scope':('黑体',15,28,'left',0),
 'Blank':('仿宋_GB2312',16,28,'left',0),
 'Image':('仿宋_GB2312',16,None,'center',0),
 'Page':('Times New Roman',14,None,'right',0),
}

class Blocked(ValueError):
    pass

def need(ok, message):
    if not ok:
        raise Blocked(message)

def dump(path, data):
    Path(path).write_text(json.dumps(data,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')

def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()

def twips(cm):
    return round(Decimal(str(cm))/Decimal('2.54')*1440)

def tag(ns, name):
    return '{'+ns+'}'+name

def text_of(value):
    if isinstance(value,list):
        return ''.join(x['text'] for x in value)
    return str(value)

def width_pt(text, size):
    # Conservative planning estimate. Final layout still needs rendered review.
    return sum(size if unicodedata.east_asian_width(c) in 'WF' else size*.55 for c in str(text))

def is_number(v):
    return isinstance(v,(int,float)) and not isinstance(v,bool)

def normalized_table(table):
    t=copy.deepcopy(table)
    need(isinstance(t.get('title'),str) and t['title'].strip(),'表格缺少标题')
    hs=t.get('headers'); rows=t.get('rows'); widths=t.get('widths_mm')
    need(isinstance(hs,list) and 1<=len(hs)<=24 and all(isinstance(v,str) and v for v in hs),'表头须为非空字符串数组')
    need(isinstance(rows,list) and len(rows)>0,'表格没有数据；空表单须由用户明确提供空行')
    need(all(isinstance(r,list) and len(r)==len(hs) for r in rows),'表格行列数不一致')
    need(isinstance(widths,list) and len(widths)==len(hs) and all(is_number(v) and math.isfinite(v) and v>0 for v in widths),'缺少有效widths_mm')
    need(sum(widths)<=156.001,'表格总宽度超过竖向A4版心156毫米；不得自动缩字或改横排')
    need(isinstance(t.get('organization'),str) and t['organization'].strip(),'表格缺少真实主责组织')
    ids=t.get('identifier_columns',[])
    need(all(type(i) is int and 0<=i<len(hs) for i in ids),'identifier_columns越界')
    for r in rows:
        for i,v in enumerate(r):
            need(isinstance(v,(str,int,float,bool,dict)) or v is None,'不支持的单元格类型')
            if is_number(v): need(math.isfinite(v),'拒绝NaN/Infinity')
            if i in ids: need(isinstance(v,str),'学号、电话等标识必须在JSON中用字符串，不能从数字猜回前导零')
            if i in ids: need(width_pt(v,16)<=widths[i]*72/25.4-8,'标识字段列宽不足，需调整列宽或按规范登记局部例外，不能裁切/自动缩字')
            if isinstance(v,dict):
                need(set(v)=={'formula','expected'} and isinstance(v['formula'],str) and v['formula'].startswith('='),'公式必须提供formula及独立核算expected')
                need(is_number(v['expected']) and math.isfinite(v['expected']),'示例计算器仅支持数值型公式结果')
    sorts=t.get('sort',[])
    if not sorts: need(t.get('ordering')=='source-confirmed','必须提供sort或确认ordering为source-confirmed；不猜组织顺序')
    if sorts:
        need(not any(isinstance(v,dict) for r in rows for v in r),'含公式表格不可直接重排；先按最终行序提供公式和expected')
        for s in reversed(sorts):
            i=s.get('column');need(type(i) is int and 0<=i<len(hs),'排序列越界')
            if 'order' in s:
                order=s['order'];need(isinstance(order,list) and len(order)==len(set(order)),'组织顺序重复或无效')
                need(all(r[i] in order for r in rows),'存在未列入组织排序依据的值')
                rows.sort(key=lambda r:order.index(r[i]))
            elif s.get('numeric'):
                need(all(re.fullmatch(r'[0-9]+',str(r[i])) for r in rows),'数字排序列含空值或非数字')
                rows.sort(key=lambda r:int(r[i]))
            else: rows.sort(key=lambda r:str(r[i]))
    merges=t.get('merge_columns',[])
    need(all(type(i) is int and 0<=i<len(hs) and i not in ids for i in merges),'不能合并标识列或越界列')
    need(not any(isinstance(r[i],dict) for r in rows for i in merges),'不能合并公式列')
    t['rows']=rows
    return t

def validate(spec):
    need(isinstance(spec,dict),'输入必须是JSON对象')
    need(spec.get('kind') in ('docx','xlsx'),'kind只能是docx或xlsx')
    need(spec.get('orientation','portrait')=='portrait','横向装订版式须走专门复核，标准生成器不冒充已支持')
    if spec['kind']=='xlsx':
        need(isinstance(spec.get('sheets'),list) and spec['sheets'],'缺少sheets')
        names=[]
        for t in spec['sheets']:
            n=t.get('name','表格');need(isinstance(n,str) and 0<len(n)<=31 and not re.search(r'[\[\]:*?/\\]',n),'无效工作表名')
            names.append(n.casefold());normalized_table(t)
        need(len(names)==len(set(names)),'工作表名重复')
        return
    need(isinstance(spec.get('title'),str) and spec['title'].strip(),'缺少公文标题')
    need(isinstance(spec.get('organization'),str) and spec['organization'].strip(),'缺少署名组织')
    need(isinstance(spec.get('signature',True),bool),'signature须为布尔值')
    try: date.fromisoformat(spec.get('date',''))
    except (ValueError,TypeError): raise Blocked('date必须是真实ISO日期，例如2026-10-18')
    if spec.get('confidential'):
        need(spec['confidential'] in ('内部资料','保密资料') and spec.get('scope'),'保密标识缺少可见范围')
        need(not spec.get('attachments'),'保密标识与附件首页规则重叠，需参照原规范专门复核；标准生成器不擅自解释')
    need(not spec.get('cover') and not spec.get('toc'),'封面/目录使用扩展流程；不可把普通通知冒充策划案成品')
    blocks=list(spec.get('blocks',[]))
    for a in spec.get('attachments',[]):
        need(isinstance(a.get('name'),str) and a['name'] and a.get('blocks'),'附件缺名称或内容')
        need(not re.search(r'[。；，、：.!?;,:]$',a['name']),'附件名称末尾不得加标点')
        blocks.extend(a['blocks'])
    need(blocks,'没有正文')
    for b in blocks:
        role=b.get('role');need(role in ('body','h1','h2','h3','h4','table','image','pagebreak'),'未知block.role，不能默认为正文')
        need(isinstance(b.get('keep_next',False),bool),'keep_next须为布尔值')
        if role=='table': normalized_table(b)
        elif role=='image': need(Path(b.get('path','')).is_file() and b.get('caption'),'图片缺真实文件或说明')
        elif role!='pagebreak':
            value=b.get('text');need(isinstance(value,(str,list)) and text_of(value),'段落内容为空')
            if isinstance(value,list): need(all(isinstance(x,dict) and isinstance(x.get('text'),str) for x in value),'分段文本须含text')
            pat={'h1':r'^[一二三四五六七八九十百]+、','h2':r'^（[一二三四五六七八九十百]+）','h3':r'^\d+\.','h4':r'^（\d+）'}.get(role)
            if pat: need(re.match(pat,text_of(value)),'标题编号与层级不匹配：'+role)

def write_font(run, chinese, size, bold=False):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt, RGBColor
    run.font.name='Times New Roman';run.font.size=Pt(size);run.font.bold=bold;run.font.color.rgb=RGBColor(0,0,0)
    rp=run._r.get_or_add_rPr();rf=rp.find(qn('w:rFonts'))
    if rf is None: rf=OxmlElement('w:rFonts');rp.insert(0,rf)
    for k,v in [('eastAsia',chinese),('ascii','Times New Roman'),('hAnsi','Times New Roman'),('cs','Times New Roman')]: rf.set(qn('w:'+k),v)
    for k in list(rf.attrib):
        if k.endswith('Theme'): del rf.attrib[k]

def paragraph(doc, value='', role='body', **overrides):
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
    from docx.oxml.ns import qn
    p=doc.add_paragraph(style='Title' if role=='Title' else 'SICSU_'+role)
    f,size,line,align,indent=RULES[role]
    pf=p.paragraph_format;pf.space_before=Pt(0);pf.space_after=Pt(0)
    pf.alignment={'center':WD_ALIGN_PARAGRAPH.CENTER,'both':WD_ALIGN_PARAGRAPH.JUSTIFY,'left':WD_ALIGN_PARAGRAPH.LEFT,'right':WD_ALIGN_PARAGRAPH.RIGHT}[align]
    if line is None: pf.line_spacing=1.0
    else: pf.line_spacing=Pt(line);pf.line_spacing_rule=WD_LINE_SPACING.EXACTLY
    pf.left_indent=Pt(0);pf.right_indent=Pt(0);pf.first_line_indent=Pt(size*indent)
    ind=p._p.get_or_add_pPr().get_or_add_ind();ind.set(qn('w:firstLineChars'),str(indent*100))
    pf.keep_with_next=role in ('Title','Subtitle','Meta','h1','h2','h3','h4','TableTitle','Attachment','Secret','Scope')
    pf.keep_together=role not in ('body','Blank')
    items=value if isinstance(value,list) else [{'text':str(value)}]
    for item in items: write_font(p.add_run(item['text']),'黑体' if item.get('emphasis') else f,size,bool(item.get('bold',False)))
    return p

def insert_ordered(parent, child, successors):
    """Insert an OOXML child at its schema position, not blindly at the end."""
    parent.insert_element_before(child, *successors)
    return child

def border(p, edge):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    bs=OxmlElement('w:pBdr');b=OxmlElement('w:'+edge)
    for k,v in [('val','single'),('sz','4'),('color','000000'),('space','2')]: b.set(qn('w:'+k),v)
    bs.append(b)
    insert_ordered(p._p.get_or_add_pPr(),bs,(
        'w:shd','w:tabs','w:suppressAutoHyphens','w:kinsoku','w:wordWrap',
        'w:overflowPunct','w:topLinePunct','w:autoSpaceDE','w:autoSpaceDN','w:bidi',
        'w:adjustRightInd','w:snapToGrid','w:spacing','w:ind','w:contextualSpacing',
        'w:mirrorIndents','w:suppressOverlap','w:jc','w:textDirection',
        'w:textAlignment','w:textboxTightWrap','w:outlineLvl','w:divId',
        'w:cnfStyle','w:rPr','w:sectPr','w:pPrChange'))

def docx_table(doc, t):
    from docx.shared import Mm, Pt
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
    t=normalized_table(t)
    caption=paragraph(doc,t['title'],'TableTitle');border(caption,'bottom')
    inset=max(0,(442.2-sum(t['widths_mm'])*72/25.4)/2)
    caption.paragraph_format.left_indent=Pt(inset);caption.paragraph_format.right_indent=Pt(inset)
    table=doc.add_table(rows=1,cols=len(t['headers']));table.autofit=False;table.alignment=WD_TABLE_ALIGNMENT.CENTER
    for c,mm in zip(table.columns,t['widths_mm']): c.width=Mm(mm)
    bs=OxmlElement('w:tblBorders')
    for edge in ['top','left','bottom','right','insideH','insideV']:
        b=OxmlElement('w:'+edge)
        for k,v in [('val','single'),('sz','4'),('color','000000')]: b.set(qn('w:'+k),v)
        bs.append(b)
    table._tbl.tblPr.insert_element_before(bs,
        'w:shd','w:tblLayout','w:tblCellMar','w:tblLook',
        'w:tblCaption','w:tblDescription','w:tblPrChange')
    repeat=OxmlElement('w:tblHeader');repeat.set(qn('w:val'),'true')
    insert_ordered(table.rows[0]._tr.get_or_add_trPr(),repeat,
        ('w:tblCellSpacing','w:jc','w:hidden','w:ins','w:del','w:trPrChange'))
    for idx,values in enumerate([t['headers']]+t['rows']):
        row=table.rows[0] if idx==0 else table.add_row()
        no_split=OxmlElement('w:cantSplit')
        insert_ordered(row._tr.get_or_add_trPr(),no_split,
            ('w:trHeight','w:tblHeader','w:tblCellSpacing','w:jc',
             'w:hidden','w:ins','w:del','w:trPrChange'))
        for c,mm,v in zip(row.cells,t['widths_mm'],values):
            c.width=Mm(mm);c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
            c._tc.remove(c.paragraphs[0]._p)
            need(not isinstance(v,dict),'Word表格不接受Excel公式；请提供已核算的显示值')
            paragraph(c,'' if v is None else str(v),'TableHead' if idx==0 else 'TableBody')
    footer=paragraph(doc,t['organization']+' 制','TableFooter');border(footer,'top')
    footer.paragraph_format.left_indent=Pt(inset);footer.paragraph_format.right_indent=Pt(inset)

def create_docx(spec, path):
    from docx import Document
    from docx.enum.style import WD_STYLE_TYPE
    from docx.shared import Cm, Pt
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    doc=Document();sec=doc.sections[0];sec.page_width=Cm(21);sec.page_height=Cm(29.7)
    for key in ['top','bottom','left','right']: setattr(sec,key+'_margin',Cm(MARGINS[key]))
    sec.footer_distance=Cm(MARGINS['footer']);doc.settings.odd_and_even_pages_header_footer=True
    for role in RULES:
        name='Title' if role=='Title' else 'SICSU_'+role
        if name not in doc.styles: doc.styles.add_style(name,WD_STYLE_TYPE.PARAGRAPH)
        style=doc.styles[name];style.base_style=doc.styles['Normal']
        for prop in ['pPr','rPr']:
            node=getattr(style.element,prop,None)
            if node is not None:style.element.remove(node)
        style.font.name='Times New Roman';style.font.size=Pt(RULES[role][1])
    for footer,alignment in [(sec.footer,WD_ALIGN_PARAGRAPH.RIGHT),(sec.even_page_footer,WD_ALIGN_PARAGRAPH.LEFT)]:
        p=footer.paragraphs[0];p.style=doc.styles['SICSU_Page'];p.alignment=alignment
        p.paragraph_format.space_before=Pt(0);p.paragraph_format.space_after=Pt(0);p.paragraph_format.line_spacing=1
        write_font(p.add_run('— '),'Times New Roman',14)
        run=p.add_run();write_font(run,'Times New Roman',14)
        fld=OxmlElement('w:fldSimple');fld.set(qn('w:instr'),'PAGE');r=OxmlElement('w:r');rp=copy.deepcopy(run._r.rPr);r.append(rp);txt=OxmlElement('w:t');txt.text='1';r.append(txt);fld.append(r);p._p.append(fld)
        write_font(p.add_run(' —'),'Times New Roman',14)
    def blocks(items):
        for b in items:
            role=b['role']
            if role=='table': docx_table(doc,b)
            elif role=='pagebreak': paragraph(doc,'','Blank').paragraph_format.page_break_before=True
            elif role=='image':
                from PIL import Image
                with Image.open(b['path']) as im: ratio=im.width/im.height
                width=float(b.get('width_mm',120));need(0<width<=156,'图片宽度超版心')
                need(abs(ratio-4/3)<.04 or abs(ratio-16/9)<.04 or b.get('ratio_exception'),'图片非4:3/16:9，需明确ratio_exception并登记')
                p=paragraph(doc,'','Image');p.add_run().add_picture(b['path'],width=Cm(width/10));paragraph(doc,b['caption'],'Caption')
            else:
                p=paragraph(doc,b['text'],role)
                if b.get('keep_next'):p.paragraph_format.keep_with_next=True
    if spec.get('confidential'):
        paragraph(doc,spec['confidential'],'Secret');paragraph(doc,spec['scope'],'Scope');paragraph(doc,'','Blank')
    paragraph(doc,spec['title'],'Title')
    if spec.get('subtitle'): paragraph(doc,spec['subtitle'],'Subtitle')
    if spec.get('meta_date'): paragraph(doc,'（'+spec['meta_date']+'）','Meta')
    if spec.get('author'): paragraph(doc,spec['author'],'Meta')
    paragraph(doc,'','Blank');blocks(spec['blocks'])
    attachments=spec.get('attachments',[])
    if attachments:
        paragraph(doc,'','Blank')
        multi=len(attachments)>1
        for i,a in enumerate(attachments,1):
            # Keep the required space after the dot: “附件：1. 名称”.
            # Item 2 and later align under item 1's “1.”; wrapped lines align
            # with the name's first character.
            prefix='附件：' if i==1 else ''
            number=(str(i)+'. ') if multi else ''
            text=prefix+number+a['name']
            p=paragraph(doc,text,'AttachmentList');pf=p.paragraph_format
            if i==1:
                hanging=width_pt(prefix+number,16)
                base=32
            else:
                hanging=width_pt(number,16)
                base=32+width_pt('附件：',16)
            pf.left_indent=Pt(base+hanging);pf.first_line_indent=Pt(-hanging)
            p._p.pPr.ind.attrib.pop(qn('w:firstLineChars'),None)
    if spec.get('signature',True):
        for _ in range(2): paragraph(doc,'','Blank')
        dt=date.fromisoformat(spec['date']);ds=f'{dt.year}年{dt.month}月{dt.day}日'
        width=min(442.2,max(width_pt(spec['organization'],16),width_pt(ds,16))+12)
        for value in [spec['organization'],ds]:
            p=paragraph(doc,value,'Signature');p.paragraph_format.left_indent=Pt(442.2-width)
    for i,a in enumerate(attachments,1):
        p=paragraph(doc,'附件'+(' '+str(i) if len(attachments)>1 else ''),'Attachment');p.paragraph_format.page_break_before=True
        paragraph(doc,'','Blank')
        same_table_title=len(a['blocks'])==1 and a['blocks'][0].get('role')=='table' and a['blocks'][0].get('title')==a['name']
        if not same_table_title:paragraph(doc,a['name'],'Title');paragraph(doc,'','Blank')
        blocks(a['blocks'])
    update=OxmlElement('w:updateFields');update.set(qn('w:val'),'true');doc.settings.element.append(update)
    doc.save(path)

def xml_update(path, updates):
    with ZipFile(path) as z: data={n:z.read(n) for n in z.namelist()}
    for n,root in updates.items(): data[n]=E.tostring(root,encoding='utf-8',xml_declaration=True)
    with ZipFile(path,'w',ZIP_DEFLATED) as z:
        for n,b in data.items():z.writestr(n,b)

def patch_xlsx(spec, path):
    """Add print settings and script-specific rich text unavailable in quick API."""
    with ZipFile(path) as z:
        styles=E.fromstring(z.read('xl/styles.xml'));wb=E.fromstring(z.read('xl/workbook.xml'))
        sheets=[E.fromstring(z.read(f'xl/worksheets/sheet{i+1}.xml')) for i in range(len(spec['sheets']))]
    fonts=styles.find('x:fonts',NS);xfs=styles.find('x:cellXfs',NS);borders=styles.find('x:borders',NS)
    def newfont(name,size):
        f=E.SubElement(fonts,tag(X,'font'));E.SubElement(f,tag(X,'sz'),val=str(size));E.SubElement(f,tag(X,'color'),rgb='FF000000');E.SubElement(f,tag(X,'name'),val=name)
        return len(fonts)-1
    def newborder(edges):
        b=E.SubElement(borders,tag(X,'border'))
        for edge in ['left','right','top','bottom','diagonal']:
            el=E.SubElement(b,tag(X,edge),**({'style':'thin'} if edge in edges else {}))
            if edge in edges:E.SubElement(el,tag(X,'color'),rgb='FF000000')
        return len(borders)-1
    allb=newborder(['left','right','top','bottom']);bottom=newborder(['bottom']);top=newborder(['top'])
    sty={}
    for role,f,size,b,align in [('title','方正小标宋简体',22,bottom,'center'),('header','黑体',16,allb,'center'),('data','仿宋_GB2312',16,allb,'center'),('footer','仿宋_GB2312',14,top,'right'),('attachment','黑体',16,0,'left'),('blank','仿宋_GB2312',16,0,'left')]:
        for latin in [False,True]:
            fi=newfont('Times New Roman' if latin else f,size)
            xf=E.SubElement(xfs,tag(X,'xf'),numFmtId='0',fontId=str(fi),fillId='0',borderId=str(b),xfId='0',applyFont='1',applyBorder='1',applyAlignment='1')
            E.SubElement(xf,tag(X,'alignment'),horizontal=align,vertical='center',wrapText='1' if role=='title' else '0',shrinkToFit='0')
            sty[(role,latin)]=str(len(xfs)-1)
    fonts.set('count',str(len(fonts)));xfs.set('count',str(len(xfs)));borders.set('count',str(len(borders)))
    defs=wb.find('x:definedNames',NS)
    if defs is None:defs=E.SubElement(wb,tag(X,'definedNames'))
    for el in list(defs):
        if el.get('name') in ('_xlnm.Print_Area','_xlnm.Print_Titles'):defs.remove(el)
    updates={'xl/styles.xml':styles,'xl/workbook.xml':wb}
    for si,(raw,root) in enumerate(zip(spec['sheets'],sheets)):
        t=normalized_table(raw);offset=2 if t.get('attachment_label') else 0;hr=2+offset;end=hr+len(t['rows'])+1;n=len(t['headers']);last=colname(n)
        sd=root.find('x:sheetData',NS)
        for row in sd:
            rn=int(row.get('r'));role='title' if rn==1+offset else 'header' if rn==hr else 'footer' if rn==end else 'attachment' if rn==1 and offset else 'blank' if rn==2 and offset else 'data'
            for c in row:
                value=cell_value(c,None)
                ci=colindex(re.match(r'[A-Z]+',c.get('r')).group())-1
                if rn==1+offset:value=t['title'] if ci==0 else ''
                elif rn==hr:value=t['headers'][ci]
                elif hr<rn<end:value=t['rows'][rn-hr-1][ci]
                elif rn==end:value=t['organization']+' 制' if ci==0 else ''
                elif rn==1 and offset:value=t['attachment_label'] if ci==0 else ''
                else:value=''
                c.set('s',sty[(role,is_number(value) or isinstance(value,dict))])
                if isinstance(value,str):
                    for el in list(c):c.remove(el)
                    c.set('t','inlineStr');inline=E.SubElement(c,tag(X,'is'))
                    font,size={'title':('方正小标宋简体',22),'header':('黑体',16),'footer':('仿宋_GB2312',14),'attachment':('黑体',16)}.get(role,('仿宋_GB2312',16))
                    for part in re.findall(r'[A-Za-z0-9]+|[^A-Za-z0-9]+',value):
                        r=E.SubElement(inline,tag(X,'r'));rp=E.SubElement(r,tag(X,'rPr'));E.SubElement(rp,tag(X,'rFont'),val='Times New Roman' if re.fullmatch('[A-Za-z0-9]+',part) else font);E.SubElement(rp,tag(X,'sz'),val=str(size));E.SubElement(rp,tag(X,'color'),rgb='FF000000');text=E.SubElement(r,tag(X,'t'));text.set('{http://www.w3.org/XML/1998/namespace}space','preserve');text.text=part
        # Worksheet children have a schema-defined order: append print nodes before extLst.
        for k in ['printOptions','pageMargins','pageSetup','headerFooter']:
            old=root.find('x:'+k,NS)
            if old is not None:root.remove(old)
        E.SubElement(root,tag(X,'printOptions'),horizontalCentered='1',gridLines='0',headings='0')
        E.SubElement(root,tag(X,'pageMargins'),**{k:str(v/2.54) for k,v in MARGINS.items()},header='0')
        E.SubElement(root,tag(X,'pageSetup'),paperSize='9',orientation='portrait',scale='100',fitToWidth='0',fitToHeight='0')
        hf=E.SubElement(root,tag(X,'headerFooter'),differentOddEven='1',differentFirst='0',scaleWithDoc='0',alignWithMargins='1')
        E.SubElement(hf,tag(X,'oddFooter')).text='&R&"Times New Roman,Regular"&14— &P —'
        E.SubElement(hf,tag(X,'evenFooter')).text='&L&"Times New Roman,Regular"&14— &P —'
        # Explicit row breaks keep at least one data row with the final maker line.
        old=root.find('x:rowBreaks',NS)
        if old is not None:root.remove(old)
        title_row=next(row for row in sd if int(row.get('r'))==1+offset)
        title_height=float(title_row.get('ht','35'))
        usable=(29.7-MARGINS['top']-MARGINS['bottom'])*72/2.54
        overhead=title_height+26+offset*26;used=0;remaining=len(t['rows']);breaks=[]
        while remaining:
            capacity=int((usable-overhead)//26)
            need(capacity>=2,'标题过长，当前页无法容纳有效表格，需专门排版')
            if remaining+1<=capacity:break
            take=remaining-1 if remaining<=capacity else capacity
            used+=take;remaining-=take;breaks.append(hr+used);overhead=26
        if breaks:
            br=E.SubElement(root,tag(X,'rowBreaks'),count=str(len(breaks)),manualBreakCount=str(len(breaks)))
            for rownum in breaks:E.SubElement(br,tag(X,'brk'),id=str(rownum),min='0',max='16383',man='1')
        sn=t.get('name','表格').replace("'","''")
        E.SubElement(defs,tag(X,'definedName'),name='_xlnm.Print_Area',localSheetId=str(si)).text=f"'{sn}'!$A$1:${last}${end}"
        E.SubElement(defs,tag(X,'definedName'),name='_xlnm.Print_Titles',localSheetId=str(si)).text=f"'{sn}'!${hr}:${hr}"
        updates[f'xl/worksheets/sheet{si+1}.xml']=root
    xml_update(path,updates)

def colname(n):
    s=''
    while n:n,r=divmod(n-1,26);s=chr(65+r)+s
    return s

def colindex(s):
    n=0
    for c in s:n=n*26+ord(c)-64
    return n

def cell_value(c, strings):
    kind=c.get('t');v=c.find('x:v',NS)
    if kind=='inlineStr':return ''.join(c.itertext())
    if kind=='s':return strings[int(v.text)] if strings is not None and v is not None else ''
    if c.find('x:f',NS) is not None:return {'formula':'='+c.find('x:f',NS).text,'cached':v.text if v is not None else None}
    if v is None:return ''
    if kind=='b':return v.text=='1'
    if kind in ('str','e'):return v.text or ''
    try:return float(v.text)
    except (ValueError,TypeError):return v.text

def create_xlsx(spec,path,work):
    node=os.environ.get('CODEX_PRIMARY_RUNTIME_NODE') or shutil.which('node')
    need(node,'缺少Node运行时')
    modules=os.environ.get('CODEX_PRIMARY_RUNTIME_NODE_MODULES')
    if modules and not (work/'node_modules').exists(): (work/'node_modules').symlink_to(modules,target_is_directory=True)
    builder=work/'xlsx_builder.mjs';shutil.copyfile(ROOT/'scripts/xlsx_builder.mjs',builder)
    data=work/'normalized.json';dump(data,{'kind':'xlsx','sheets':[normalized_table(t) for t in spec['sheets']]})
    p=subprocess.run([node,str(builder),str(data),str(path)],cwd=work,capture_output=True,text=True,timeout=180)
    need(p.returncode==0,'Excel标准生成失败：'+p.stderr[-2500:])
    patch_xlsx(spec,path)

def font_inventory(env=None):
    result={}
    if shutil.which('fc-list'):
        p=subprocess.run(['fc-list','--format=%{family}\n'],capture_output=True,text=True,timeout=20,env=env)
        for line in p.stdout.splitlines():
            for family in line.split(','):result[family.strip().casefold()]=family.strip()
    return result

def required_fonts(spec):
    fs={'Times New Roman','方正小标宋简体','仿宋_GB2312','黑体'}
    if spec['kind']=='docx':
        fs.add('楷体_GB2312')
    return sorted(fs)

def doctor(spec):
    bundle=bundle_status()
    with tempfile.TemporaryDirectory(prefix='sicsu-fontcheck-',dir=os.environ.get('TMPDIR') or Path.cwd()) as work:
        inv=font_inventory(font_environment(work))
    missing=[f for f in required_fonts(spec) if f.casefold() not in inv]
    return {'missing_fonts':missing,'soffice':shutil.which('soffice') or shutil.which('libreoffice'),'font_check':'exact internal families; bundled fonts privately loaded, no fallback aliases','bundled_fonts':[f['family'] for f in bundle['fonts']],'bundle_issues':bundle['issues'],'font_bundle_sha256':bundle['manifest_sha256']}

def export_pdf(source,out,work):
    engine=shutil.which('soffice') or shutil.which('libreoffice');need(engine,'缺少LibreOffice/soffice，不能交付单独源文件')
    out.mkdir(parents=True,exist_ok=True);profile=work/'lo-profile';profile.mkdir(exist_ok=True)
    env=font_environment(work/'fonts');env.update({'TMPDIR':str(work),'TMP':str(work),'TEMP':str(work),'XDG_CACHE_HOME':str(work/'cache'),'XDG_CONFIG_HOME':str(work/'config')})
    pdf=out/(source.stem+'.pdf')
    if pdf.exists():pdf.unlink()
    p=subprocess.run([engine,'-env:UserInstallation='+profile.as_uri(),'--headless','--convert-to','pdf','--outdir',str(out),str(source)],capture_output=True,text=True,env=env,timeout=90)
    need(p.returncode==0 and pdf.exists() and pdf.stat().st_size>100,'PDF导出失败：'+p.stderr[-1200:]+p.stdout[-1200:])
    return pdf

def audit_source(source,spec):
    """Independent structural audit, plus content/formula comparisons to frozen input."""
    issues=[]
    def check(ok,msg):
        if not ok:issues.append(msg)
    def check_order(parent,before,after,label):
        if parent is None:return
        names=[c.tag.split('}')[-1] for c in parent]
        if before in names and after in names:
            check(names.index(before)<names.index(after),f'OOXML元素顺序错误：{label} {before} 应在 {after} 之前')
    with ZipFile(source) as z:
        if source.suffix=='.docx':
            root=E.fromstring(z.read('word/document.xml'))
            styles_root=E.fromstring(z.read('word/styles.xml'))
            for s in styles_root.findall('w:style',NS):
                if s.get(tag(W,'styleId'))=='Title':check(s.find('w:pPr/w:pBdr',NS) is None,'标题继承了非规范装饰边框')
            for sec in root.findall('.//w:sectPr',NS):
                pg=sec.find('w:pgSz',NS);mg=sec.find('w:pgMar',NS)
                check(pg is not None and abs(int(pg.get(tag(W,'w'),'0'))-twips(21))<=1 and abs(int(pg.get(tag(W,'h'),'0'))-twips(29.7))<=1,'Word纸型不是A4')
                for key,value in MARGINS.items():check(mg is not None and abs(int(mg.get(tag(W,key),'0'))-twips(value))<=1,'Word页边距/页脚错误：'+key)
            for name in [n for n in z.namelist() if n=='word/document.xml' or re.match(r'word/footer\d+\.xml',n)]:
                rr=E.fromstring(z.read(name))
                for p in rr.findall('.//w:p',NS):
                    ps=p.find('w:pPr/w:pStyle',NS);style=ps.get(tag(W,'val'),'') if ps is not None else ''
                    role='Title' if style=='Title' else style.replace('SICSU_','')
                    check(role in RULES,'未知段落样式：'+style)
                    ppr=p.find('w:pPr',NS)
                    for later in ('spacing','ind','jc'):
                        check_order(ppr,'pBdr',later,'段落')
                    if role not in RULES:continue
                    font,size,line,align,indent=RULES[role]
                    spacing=p.find('w:pPr/w:spacing',NS)
                    if role!='Page':
                        check(spacing is not None and spacing.get(tag(W,'before'),'0')=='0' and spacing.get(tag(W,'after'),'0')=='0','段前后间距错误：'+role)
                        check(spacing is not None and spacing.get(tag(W,'line'))==str(round(line*20) if line else 240) and spacing.get(tag(W,'lineRule'))==('exact' if line else 'auto'),'行距错误：'+role)
                        jc=p.find('w:pPr/w:jc',NS);check(jc is not None and jc.get(tag(W,'val'))==align,'段落对齐错误：'+role)
                    if indent:
                        ind=p.find('w:pPr/w:ind',NS);check(ind is not None and ind.get(tag(W,'firstLineChars'))==str(indent*100),'首行缩进字符数错误：'+role)
                        check(ind is not None and ind.get(tag(W,'firstLine'))==str(round(size*indent*20)),'首行缩进实际值错误：'+role)
                    for run in p.findall('.//w:r',NS):
                        if not ''.join(run.itertext()):continue
                        rf=run.find('w:rPr/w:rFonts',NS);sz=run.find('w:rPr/w:sz',NS);color=run.find('w:rPr/w:color',NS)
                        check(rf is not None and rf.get(tag(W,'ascii'))=='Times New Roman' and rf.get(tag(W,'hAnsi'))=='Times New Roman','西文字体错误：'+role)
                        run_text=''.join(run.itertext())
                        emphasis=[]
                        for block in spec.get('blocks',[]):
                            if isinstance(block.get('text'),list):emphasis.extend(x['text'] for x in block['text'] if x.get('emphasis'))
                        allowed=(font,'黑体') if role=='body' and run_text in emphasis else (font,)
                        check(rf is not None and rf.get(tag(W,'eastAsia')) in allowed,'中文字体错误：'+role)
                        check(sz is not None and float(sz.get(tag(W,'val'),'0'))==size*2,'字号错误：'+role)
                        check(color is not None and color.get(tag(W,'val'))=='000000','颜色错误：'+role)
            settings=E.fromstring(z.read('word/settings.xml'));check(settings.find('w:evenAndOddHeaders',NS) is not None,'没有启用奇偶页脚')
            for ref in root.findall('.//w:sectPr/w:footerReference',NS):
                # Our standard builder writes two footer parts; check actual field content below.
                check(ref.get(tag(W,'type')) in ('default','even'),'非标准首页页脚覆盖')
            for footer_name in [n for n in z.namelist() if re.match(r'word/footer\d+\.xml',n)]:
                footer=E.fromstring(z.read(footer_name))
                check(any((f.get(tag(W,'instr')) or '').strip()=='PAGE' for f in footer.findall('.//w:fldSimple',NS)),'Word页码不是动态PAGE域')
            for t in root.findall('.//w:tbl',NS):
                check(t.find('w:tr/w:trPr/w:tblHeader',NS) is not None,'Word续表缺重复表头')
                tp=t.find('w:tblPr',NS)
                check_order(tp,'tblBorders','tblLayout','表格')
                check_order(tp,'tblBorders','tblLook','表格')
                for trpr in t.findall('w:tr/w:trPr',NS):
                    check_order(trpr,'cantSplit','tblHeader','表格行')
            for p in root.findall('.//w:p',NS):
                style=p.find('w:pPr/w:pStyle',NS)
                if style is not None and style.get(tag(W,'val')) in ('SICSU_TableTitle','SICSU_TableFooter'):
                    edge='bottom' if style.get(tag(W,'val'))=='SICSU_TableTitle' else 'top';check(p.find('w:pPr/w:pBdr/w:'+edge,NS) is not None,'表格标题/落款边框缺失')
            got=''.join(root.itertext())
            expected=[spec['title']]+([spec['organization']] if spec.get('signature',True) else [])
            for b in spec.get('blocks',[]):
                if b.get('role') in ('body','h1','h2','h3','h4'):expected.append(text_of(b['text']))
            for value in expected:check(value.replace('\n','') in got,'Word内容缺失：'+value[:30])
        else:
            styles=E.fromstring(z.read('xl/styles.xml'));xfs=list(styles.find('x:cellXfs',NS));fonts=list(styles.find('x:fonts',NS));borders=list(styles.find('x:borders',NS))
            wb=E.fromstring(z.read('xl/workbook.xml'));strings=[]
            if 'xl/sharedStrings.xml' in z.namelist():strings=[''.join(x.itertext()) for x in E.fromstring(z.read('xl/sharedStrings.xml'))]
            expected_count=len(spec['sheets']);check(len(wb.find('x:sheets',NS))==expected_count,'工作表数不匹配')
            for si,raw in enumerate(spec['sheets']):
                t=normalized_table(raw);root=E.fromstring(z.read(f'xl/worksheets/sheet{si+1}.xml'));mg=root.find('x:pageMargins',NS);ps=root.find('x:pageSetup',NS);hf=root.find('x:headerFooter',NS)
                for k,v in MARGINS.items():check(mg is not None and abs(float(mg.get(k,'0'))*2.54-v)<.002,'Excel页边距/页脚错误：'+k)
                check(ps is not None and ps.get('paperSize')=='9' and ps.get('scale')=='100','Excel纸型或缩放错误')
                check(hf is not None and hf.get('differentOddEven')=='1' and hf.get('scaleWithDoc')=='0','Excel奇偶页脚/缩放错误')
                for k,side in [('oddFooter','&R'),('evenFooter','&L')]:check(hf is not None and (hf.find('x:'+k,NS).text or '').startswith(side) and '&14— &P —' in hf.find('x:'+k,NS).text,'Excel页码格式错误：'+k)
                offset=2 if t.get('attachment_label') else 0;hr=2+offset;end=hr+len(t['rows'])+1;n=len(t['headers']);cells={c.get('r'):c for c in root.findall('.//x:sheetData/x:row/x:c',NS)}
                check(f'A{1+offset}' in cells,'Excel表格标题单元格缺失')
                check(f'A{end}' in cells,'Excel表格落款单元格缺失')
                defs=[d for d in wb.findall('x:definedNames/x:definedName',NS) if d.get('localSheetId')==str(si)]
                check(any(d.get('name')=='_xlnm.Print_Titles' and d.text.endswith(f'${hr}:${hr}') for d in defs),'Excel重复表头错误')
                check(any(d.get('name')=='_xlnm.Print_Area' and d.text.endswith(f'${colname(n)}${end}') for d in defs),'Excel打印区域漏数据')
                for ri,row in enumerate([t['headers']]+t['rows']):
                    for ci,v in enumerate(row):
                        addr=f'{colname(ci+1)}{hr+ri}';c=cells.get(addr)
                        if c is None:check(False,'Excel缺单元格：'+addr);continue
                        got=cell_value(c,strings)
                        if isinstance(v,dict):check(isinstance(got,dict) and got['formula']==v['formula'] and got['cached'] is not None and abs(float(got['cached'])-v['expected'])<1e-8,'Excel公式/独立期望值不匹配：'+addr)
                        else:check((got==v) or (got=='' and v is None),'Excel内容不匹配：'+addr)
                        xf=xfs[int(c.get('s','0'))];font=fonts[int(xf.get('fontId','0'))];bid=int(xf.get('borderId','0'));align=xf.find('x:alignment',NS)
                        check(font.find('x:sz',NS) is not None and float(font.find('x:sz',NS).get('val'))==16,'Excel字号错误：'+addr)
                        check(align is not None and align.get('horizontal')=='center' and align.get('shrinkToFit','0')=='0','Excel对齐/缩字错误：'+addr)
                        check(all(borders[bid].find('x:'+e,NS) is not None and borders[bid].find('x:'+e,NS).get('style') for e in ['left','right','top','bottom']),'Excel边框错误：'+addr)
                        expected_cn='黑体' if ri==0 else '仿宋_GB2312'
                        for run in c.findall('x:is/x:r',NS):
                            tx=''.join(run.find('x:t',NS).itertext());expected_font='Times New Roman' if re.fullmatch('[A-Za-z0-9]+',tx) else expected_cn
                            rf=run.find('x:rPr/x:rFont',NS);check(rf is not None and rf.get('val')==expected_font,'Excel富文本字体错误：'+addr)
                # Independently check title/footer plus effective run size/color and numeric font.
                for addr,c in cells.items():
                    rn=int(re.search(r'\d+',addr).group());ci=colindex(re.match(r'[A-Z]+',addr).group())-1
                    role='title' if rn==1+offset else 'header' if rn==hr else 'footer' if rn==end else 'attachment' if rn==1 and offset else 'blank' if rn==2 and offset else 'data'
                    cn,size,edges={'title':('方正小标宋简体',22,['bottom']),'header':('黑体',16,['left','right','top','bottom']),'footer':('仿宋_GB2312',14,['top']),'attachment':('黑体',16,[]),'blank':('仿宋_GB2312',16,[])}.get(role,('仿宋_GB2312',16,['left','right','top','bottom']))
                    if role=='title' and ci==0:check(cell_value(c,strings)==t['title'],'Excel表格标题内容错误')
                    if role=='footer' and ci==0:check(cell_value(c,strings)==t['organization']+' 制','Excel表格落款内容错误')
                    xf=xfs[int(c.get('s','0'))];font=fonts[int(xf.get('fontId','0'))];bd=borders[int(xf.get('borderId','0'))]
                    alignment=xf.find('x:alignment',NS);want_align='right' if role=='footer' else 'left' if role in ('attachment','blank') else 'center'
                    check(alignment is not None and alignment.get('horizontal')==want_align and alignment.get('vertical')=='center','Excel要素对齐错误：'+addr)
                    val=cell_value(c,strings);latin=is_number(val) or isinstance(val,dict)
                    check(font.find('x:name',NS) is not None and font.find('x:name',NS).get('val')==('Times New Roman' if latin else cn),'Excel单元格字体错误：'+addr)
                    check(font.find('x:sz',NS) is not None and float(font.find('x:sz',NS).get('val'))==size,'Excel要素字号错误：'+addr)
                    check(font.find('x:color',NS) is not None and font.find('x:color',NS).get('rgb')=='FF000000','Excel颜色错误：'+addr)
                    check(all(bool(bd.find('x:'+e,NS) is not None and bd.find('x:'+e,NS).get('style'))==(e in edges) for e in ['left','right','top','bottom']),'Excel要素边框错误：'+addr)
                    for run in c.findall('x:is/x:r',NS):
                        sz=run.find('x:rPr/x:sz',NS);co=run.find('x:rPr/x:color',NS);rf=run.find('x:rPr/x:rFont',NS);tx=''.join(run.find('x:t',NS).itertext())
                        check(sz is not None and float(sz.get('val'))==size and co is not None and co.get('rgb')=='FF000000','Excel富文本字号/颜色错误：'+addr)
                        check(rf is not None and rf.get('val')==('Times New Roman' if re.fullmatch('[A-Za-z0-9]+',tx) else cn),'Excel要素富文本字体错误：'+addr)
    return sorted(set(issues))

def audit_pdf(pdf,preview=None):
    import fitz
    issues=[];pages=[]
    if preview is not None:preview.mkdir(parents=True,exist_ok=True)
    with fitz.open(pdf) as doc:
        for i,p in enumerate(doc):
            if abs(p.rect.width-595.276)>1 or abs(p.rect.height-841.89)>1:issues.append(f'PDF第{i+1}页非竖向A4')
            text=p.get_text();needle=f'— {i+1} —'
            boxes=p.search_for(needle)
            if not boxes:issues.append(f'PDF第{i+1}页页码缺失或不连续')
            elif not any(b.y0>p.rect.height*.75 and ((b.x0>p.rect.width/2) if (i+1)%2 else (b.x1<p.rect.width/2)) for b in boxes):issues.append(f'PDF第{i+1}页奇偶位置错误')
            if '\ufffd' in text or '\x00' in text:issues.append(f'PDF第{i+1}页含缺字标记')
            item={'page':i+1,'fonts':[f[3] for f in p.get_fonts()]}
            if preview is not None:
                p.get_pixmap(matrix=fitz.Matrix(1.3,1.3)).save(preview/f'page-{i+1:03}.png')
                item.update(image=f'page-{i+1:03}.png',image_sha256=sha(preview/f'page-{i+1:03}.png'))
            pages.append(item)
    need(pages,'PDF没有页面')
    return issues,pages

def build(args):
    out=Path(args.out).resolve();out.mkdir(parents=True,exist_ok=True)
    run_id=uuid.uuid4().hex
    dump(out/'attempt.json',{'run_id':run_id,'status':'BUILDING','requested_input':str(Path(args.input).resolve())})
    if (out/'delivery').exists():
        history=out/'.previous-delivery';history.mkdir(exist_ok=True);shutil.move(str(out/'delivery'),str(history/run_id))
    spec=json.loads(Path(args.input).read_text(encoding='utf-8'));validate(spec)
    d=doctor(spec);dump(out/'preflight.json',d)
    need(d['soffice'],'缺少PDF排版引擎；请先配置soffice')
    need(not d.get('bundle_issues') or args.draft,'包内字体损坏：'+'；'.join(d.get('bundle_issues',[])))
    need(not d['missing_fonts'] or args.draft,'缺少指定字体：'+', '.join(d['missing_fonts'])+'；禁止发布，可用--draft仅测试代码链路')
    suffix='_公文' if spec['kind']=='docx' else '_表格';stem=args.name+suffix+('_待确认草稿' if args.draft else '')
    need(not re.search(r'[<>:"/\\|?*\x00-\x1f]',stem) and stem not in ('.','..'),'无效输出名称')
    source=out/(stem+'.'+spec['kind']);work=out/'.work';work.mkdir(exist_ok=True)
    if spec['kind']=='docx':create_docx(spec,source)
    else:create_xlsx(spec,source,work)
    issues=audit_source(source,spec);need(not issues,'源文件机检失败：'+'；'.join(issues))
    glyph_issues=bundled_glyph_issues(source)
    need(not glyph_issues or args.draft,'字体字形检查失败：'+'；'.join(glyph_issues))
    pdf=export_pdf(source,out,work);pdfissues,pages=audit_pdf(pdf,out/'preview')
    pdfissues.extend(glyph_issues)
    pdfissues.extend(bundled_pdf_font_issues(source,pdf))
    snapshot=out/'input.snapshot.json';dump(snapshot,spec)
    manifest={'run_id':run_id,'status':'DRAFT' if args.draft else 'REVIEW_REQUIRED','source':source.name,'source_sha256':sha(source),'pdf':pdf.name,'pdf_sha256':sha(pdf),'input':snapshot.name,'input_sha256':sha(snapshot),'engine':d['soffice'],'missing_fonts':d['missing_fonts'],'machine_issues':pdfissues,'pages':pages,'interpretation':'正文按具体条款两端对齐；范围行沿用黑体；详见format-standard.md'}
    manifest['font_bundle_sha256']=d.get('font_bundle_sha256')
    dump(out/'manifest.json',manifest)
    dump(out/'attempt.json',{'run_id':run_id,'status':'BUILT','requested_input':str(Path(args.input).resolve())})
    print(json.dumps({'status':manifest['status'],'manifest':str(out/'manifest.json'),'machine_issues':pdfissues},ensure_ascii=False))
    return 2 if pdfissues else 0

def verify(folder):
    folder=Path(folder).resolve();m=json.loads((folder/'manifest.json').read_text());issues=[]
    attempt=json.loads((folder/'attempt.json').read_text()) if (folder/'attempt.json').exists() else {}
    if attempt.get('status')!='BUILT' or attempt.get('run_id')!=m.get('run_id'):return m,['本次生成未完成或与旧版本不同；旧验收不得沿用']
    for key in ['source','pdf','input']:
        p=folder/m[key]
        if not p.is_file() or sha(p)!=m[key+'_sha256']:issues.append(key+'文件缺失或已改变，必须重新生成及审查')
    if issues:return m,issues
    spec=json.loads((folder/m['input']).read_text());validate(spec)
    issues.extend(audit_source(folder/m['source'],spec));issues.extend(m['machine_issues'])
    current_pdf_issues,current_pages=audit_pdf(folder/m['pdf'])
    issues.extend(current_pdf_issues)
    if m.get('font_bundle_sha256'):issues.extend(bundled_pdf_font_issues(folder/m['source'],folder/m['pdf']))
    if len(m['pages'])!=len(current_pages) or [x['page'] for x in m['pages']]!=[x['page'] for x in current_pages]:issues.append('PDF实际页数与审查清单不匹配')
    for p in m['pages']:
        image=folder/'preview'/p['image']
        if not image.is_file() or sha(image)!=p['image_sha256']:issues.append('逐页预览已缺失或改变')
    if m['missing_fonts']:issues.append('生成时缺少指定字体')
    if m.get('font_bundle_sha256'):
        bundle=bundle_status();issues.extend(bundle['issues'])
        if bundle['manifest_sha256']!=m['font_bundle_sha256']:issues.append('包内字体清单已变化，须重新生成及审查')
    if m['status']=='DRAFT':issues.append('草稿不能发布；补齐环境后不带--draft重新生成')
    return m,sorted(set(issues))

def release(args):
    folder=Path(args.folder).resolve();m,issues=verify(folder)
    review=json.loads(Path(args.review).read_text(encoding='utf-8'))
    need(review.get('source_sha256')==m['source_sha256'] and review.get('pdf_sha256')==m['pdf_sha256'],'审查记录绑定的文件版本不一致')
    need(review.get('reviewer') and review.get('facts_checked') is True and review.get('fonts_verified') is True and review.get('interpretation_accepted') is True,'缺少事实校核/字体核对/执行约定确认/审查者')
    checks=review.get('pages',[])
    need(len(checks)==len(m['pages']) and {r.get('page') for r in checks}==set(range(1,len(m['pages'])+1)),'审查记录未覆盖全部页')
    need(all(r.get('passed') is True and isinstance(r.get('evidence'),str) and r['evidence'].strip() for r in checks),'不得伪填通过；每页须有真实检查证据')
    need(not review.get('unresolved_issues'),'审查仍有未解决问题')
    need(not issues,'发布阻断：'+'；'.join(issues))
    target=folder/'delivery'/m['source_sha256'][:16];target.mkdir(parents=True,exist_ok=True)
    for key in ['source','pdf']:shutil.copy2(folder/m[key],target/m[key])
    dump(target/'audit.json',{'status':'按本SOP执行约定通过','manifest':m,'review':review})
    print(str(target));return 0

def main():
    p=argparse.ArgumentParser(description=__doc__);sub=p.add_subparsers(dest='command',required=True)
    for cmd in ['build','repair']:
        b=sub.add_parser(cmd);b.add_argument('input');b.add_argument('--out',required=True);b.add_argument('--name',default='公文');b.add_argument('--draft',action='store_true')
    b=sub.add_parser('doctor');b.add_argument('input')
    b=sub.add_parser('audit');b.add_argument('folder')
    b=sub.add_parser('review-template');b.add_argument('folder')
    b=sub.add_parser('release');b.add_argument('folder');b.add_argument('--review',required=True)
    args=p.parse_args()
    try:
        if args.command in ('build','repair'):return build(args)
        if args.command=='doctor':
            s=json.loads(Path(args.input).read_text());validate(s);d=doctor(s);print(json.dumps(d,ensure_ascii=False,indent=2));return 2 if d['missing_fonts'] or d.get('bundle_issues') or not d['soffice'] else 0
        if args.command=='audit':
            m,issues=verify(args.folder);print(json.dumps({'status':'BLOCKED' if issues else 'REVIEW_REQUIRED','issues':issues},ensure_ascii=False,indent=2));return 2 if issues else 0
        if args.command=='review-template':
            f=Path(args.folder);m=json.loads((f/'manifest.json').read_text());dump(f/'review.json',{'reviewer':'','source_sha256':m['source_sha256'],'pdf_sha256':m['pdf_sha256'],'facts_checked':False,'fonts_verified':False,'interpretation_accepted':False,'pages':[{'page':x['page'],'passed':False,'evidence':''} for x in m['pages']],'unresolved_issues':[]});return 0
        return release(args)
    except (Blocked,KeyError,TypeError,ValueError,OSError,subprocess.TimeoutExpired) as e:
        print('BLOCKED: '+str(e),file=sys.stderr);return 2

if __name__=='__main__':sys.exit(main())
