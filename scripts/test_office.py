#!/usr/bin/env python3
"""Run with: python scripts/test_office.py. Tests never approve a visual review."""
import copy
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest
import xml.etree.ElementTree as E
from unittest.mock import patch
import office as o
import bundled_fonts as bf
import add_fonts as af

def table():
    return {'name':'Roster','title':'报名表2026','headers':['学院','姓名','学号'],
            'rows':[['文学院','乙','00102'],['文学院','甲','00101'],['外国语学院','丙','00103']],
            'widths_mm':[52,30,74],'identifier_columns':[2],'organization':'示例组织',
            'sort':[{'column':0,'order':['文学院','外国语学院']},{'column':2,'numeric':True}]}

def doc():
    return {'kind':'docx','title':'Test notice 2026','organization':'Example Organization','date':'2026-10-18',
            'blocks':[{'role':'h1','text':'一、活动安排'},{'role':'h2','text':'（一）说明'},
                      {'role':'h3','text':'1.Example'},{'role':'h4','text':'（1）Example'},
                      {'role':'body','text':[{'text':'Example 2026 '},{'text':'重点','emphasis':True}]}]}

class Tests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory(dir=os.environ.get('TMPDIR') or '.')
        self.work=Path(self.tmp.name).resolve()
    def tearDown(self):self.tmp.cleanup()
    def test_margin_roundtrip(self):
        for cm in list(o.MARGINS.values())+[21,29.7]:self.assertLess(abs(o.twips(cm)/1440*2.54-cm),.001)
        self.assertAlmostEqual(21-o.MARGINS['left']-o.MARGINS['right'],15.6)
    def test_sort_and_identifiers(self):
        t=table();r=o.normalized_table(t)
        self.assertEqual([x[2] for x in r['rows']],['00101','00102','00103'])
        self.assertEqual(t['rows'][0][2],'00102')
    def test_numeric_identifier_rejected(self):
        t=table();t['rows'][0][2]=102
        with self.assertRaises(o.Blocked):o.normalized_table(t)
    def test_unknown_organization_rejected(self):
        t=table();t['rows'][0][0]='未知学院'
        with self.assertRaises(o.Blocked):o.normalized_table(t)
    def test_bad_dimensions_empty_and_nan(self):
        for mutation in [lambda t:t.update(rows=[]),lambda t:t.update(widths_mm=[100,100,100]),lambda t:t['rows'][0].pop(),lambda t:t['rows'][0].__setitem__(0,float('nan'))]:
            t=table();mutation(t)
            with self.assertRaises(o.Blocked):o.normalized_table(t)
    def test_bad_date_heading_and_roles(self):
        for mutate in [lambda s:s.update(date='2026-02-30'),lambda s:s['blocks'][0].update(text='1.错误'),lambda s:s['blocks'][0].update(role='heading9')]:
            s=doc();mutate(s)
            with self.assertRaises(o.Blocked):o.validate(s)
    def test_missing_cover_or_confidential_scope(self):
        s=doc();s.update(cover=True)
        with self.assertRaises(o.Blocked):o.validate(s)
        s=doc();s.update(confidential='内部资料')
        with self.assertRaises(o.Blocked):o.validate(s)
    def test_formula_reorder_blocked(self):
        t=table();t['rows'][0][1]={'formula':'=1+2','expected':3}
        with self.assertRaises(o.Blocked):o.normalized_table(t)
    def test_docx_structural_and_mutation(self):
        s=doc();p=self.work/'notice.docx';o.create_docx(s,p)
        self.assertEqual(o.audit_source(p,s),[])
        with o.ZipFile(p) as z:r=E.fromstring(z.read('word/document.xml'))
        r.find('.//w:pgMar',o.NS).set(o.tag(o.W,'left'),'1');o.xml_update(p,{'word/document.xml':r})
        self.assertTrue(any('页边距' in x for x in o.audit_source(p,s)))
    def test_docx_table_and_attachment(self):
        s=doc();t=table();t['role']='table';t['rows']=[['文学院','甲','00101']]
        s['attachments']=[{'name':'信息表','blocks':[t]}]
        p=self.work/'attachment.docx';o.create_docx(s,p)
        self.assertEqual(o.audit_source(p,s),[])
    def test_script_without_signature_block(self):
        s=doc();s['signature']=False;p=self.work/'script.docx';o.create_docx(s,p)
        self.assertEqual(o.audit_source(p,s),[])
        with o.ZipFile(p) as z:r=E.fromstring(z.read('word/document.xml'))
        self.assertFalse(any(x.find('w:pPr/w:pStyle',o.NS) is not None
                             and x.find('w:pPr/w:pStyle',o.NS).get(o.tag(o.W,'val'))=='SICSU_Signature'
                             for x in r.findall('.//w:p',o.NS)))
        bad=doc();bad['signature']='no'
        with self.assertRaises(o.Blocked):o.validate(bad)
    def test_multiline_title_breaks_inside_title_paragraph(self):
        s=doc();s['title']='测试公文标题\n第二行保持词意完整'
        p=self.work/'title.docx';o.create_docx(s,p)
        self.assertEqual(o.audit_source(p,s),[])
        with o.ZipFile(p) as z:r=E.fromstring(z.read('word/document.xml'))
        first=r.find('.//w:body/w:p',o.NS)
        self.assertEqual(first.find('w:pPr/w:pStyle',o.NS).get(o.tag(o.W,'val')),'Title')
        self.assertEqual(len(first.findall('.//w:br',o.NS)),1)
    def test_keep_next_block_stays_with_following(self):
        s=doc();s['blocks']=[{'role':'body','text':'成绩公布：','keep_next':True},{'role':'body','text':'1号成绩为：'}]
        p=self.work/'keepnext.docx';o.create_docx(s,p)
        self.assertEqual(o.audit_source(p,s),[])
        from docx import Document
        paras=[x for x in Document(p).paragraphs if x.text.strip()]
        self.assertTrue(paras[0].paragraph_format.keep_with_next)
        self.assertFalse(paras[-1].paragraph_format.keep_with_next)
        bad=doc();bad['blocks']=[{'role':'body','text':'示例','keep_next':'yes'}]
        with self.assertRaises(o.Blocked):o.validate(bad)
    def test_multiple_attachments_keep_required_space_and_alignment(self):
        s=doc();s['attachments']=[
            {'name':'报名表','blocks':[{'role':'body','text':'附件一正文'}]},
            {'name':'汇总表','blocks':[{'role':'body','text':'附件二正文'}]}]
        p=self.work/'attachments.docx';o.create_docx(s,p)
        with o.ZipFile(p) as z:r=E.fromstring(z.read('word/document.xml'))
        paras=[( ''.join(x.itertext()),x.find('w:pPr/w:ind',o.NS)) for x in r.findall('.//w:p',o.NS)]
        first=next(x for x in paras if x[0].startswith('附件：'))
        second=next(x for x in paras if x[0].startswith('2.'))
        self.assertEqual(first[0],'附件：1. 报名表')
        self.assertEqual(second[0],'2. 汇总表')
        # Item 2's “2.” starts under item 1's “1.”; wrapped lines align with names.
        self.assertEqual(first[1].get(o.tag(o.W,'left')),second[1].get(o.tag(o.W,'left')))
        self.assertNotEqual(first[1].get(o.tag(o.W,'hanging')),second[1].get(o.tag(o.W,'hanging')))
    def test_ooxml_child_order_is_schema_compliant(self):
        s=doc();t=table();t['role']='table';t['rows']=[['文学院','甲','00101']]
        s['blocks'].append(t);p=self.work/'order.docx';o.create_docx(s,p)
        with o.ZipFile(p) as z:r=E.fromstring(z.read('word/document.xml'))
        for ppr in r.findall('.//w:pPr',o.NS):
            names=[c.tag.split('}')[-1] for c in ppr]
            if 'pBdr' in names:
                for later in ('spacing','ind','jc'):
                    if later in names:self.assertLess(names.index('pBdr'),names.index(later))
        for tp in r.findall('.//w:tblPr',o.NS):
            names=[c.tag.split('}')[-1] for c in tp]
            for later in ('tblLayout','tblLook'):
                if 'tblBorders' in names and later in names:self.assertLess(names.index('tblBorders'),names.index(later))
        for tr in r.findall('.//w:trPr',o.NS):
            names=[c.tag.split('}')[-1] for c in tr]
            if 'cantSplit' in names and 'tblHeader' in names:self.assertLess(names.index('cantSplit'),names.index('tblHeader'))
        self.assertEqual(o.audit_source(p,s),[])
    def test_single_attachment_table_title_is_not_duplicated(self):
        s=doc();t=table();t['role']='table';t['title']='信息表'
        s['attachments']=[{'name':'信息表','blocks':[t]}];p=self.work/'attachment.docx';o.create_docx(s,p)
        with o.ZipFile(p) as z:r=E.fromstring(z.read('word/document.xml'))
        self.assertEqual(''.join(r.itertext()).count('信息表'),2)
    def test_strict_missing_font_blocks(self):
        f=self.work/'input.json';o.dump(f,doc())
        args=type('Args',(),dict(input=str(f),out=str(self.work/'out'),name='test',draft=False))()
        with patch.object(o,'doctor',return_value={'missing_fonts':['仿宋_GB2312'],'soffice':'mock'}):
            with self.assertRaises(o.Blocked):o.build(args)
        self.assertFalse(list((self.work/'out').glob('*.docx')))
    def test_failed_rebuild_invalidates_old_manifest(self):
        out=self.work/'out';out.mkdir();o.dump(out/'manifest.json',{'run_id':'old'})
        o.dump(out/'attempt.json',{'run_id':'old','status':'BUILT'});(out/'delivery').mkdir();(out/'delivery'/'old.txt').write_text('old')
        f=self.work/'input.json';o.dump(f,doc());a=type('Args',(),dict(input=str(f),out=str(out),name='test',draft=False))()
        with patch.object(o,'doctor',return_value={'missing_fonts':['仿宋_GB2312'],'soffice':'mock'}):
            with self.assertRaises(o.Blocked):o.build(a)
        self.assertTrue(any('本次生成未完成' in x for x in o.verify(out)[1]));self.assertFalse((out/'delivery').exists())
    def test_pdf_and_stale_pair_gate(self):
        import fitz
        f=self.work/'test.pdf';pdf=fitz.open();p=pdf.new_page(width=595.276,height=841.89)
        # Use built-in CJK font for the em dash in a test fixture, never an official output.
        p.insert_text((450,770),'— 1 —',fontname='china-s',fontsize=14);pdf.save(f);pdf.close()
        issues,pages=o.audit_pdf(f,self.work/'preview');self.assertEqual(issues,[])
        src=self.work/'notice.docx';s=doc();o.create_docx(s,src);o.dump(self.work/'input.snapshot.json',s)
        m={'run_id':'fixture','status':'DRAFT','source':src.name,'source_sha256':o.sha(src),'pdf':f.name,'pdf_sha256':o.sha(f),'input':'input.snapshot.json','input_sha256':o.sha(self.work/'input.snapshot.json'),'missing_fonts':[],'machine_issues':[],'pages':pages}
        o.dump(self.work/'attempt.json',{'run_id':'fixture','status':'BUILT'})
        o.dump(self.work/'manifest.json',m);self.assertTrue(any('草稿' in x for x in o.verify(self.work)[1]))
        altered=copy.deepcopy(m);altered['pages']=[];o.dump(self.work/'manifest.json',altered)
        self.assertTrue(any('实际页数' in x for x in o.verify(self.work)[1]));o.dump(self.work/'manifest.json',m)
        src.write_bytes(src.read_bytes()+b'changed');self.assertTrue(any('已改变' in x for x in o.verify(self.work)[1]))
    def test_release_rejects_incomplete_or_wrong_version_review(self):
        manifest={'source_sha256':'source','pdf_sha256':'pdf','pages':[{'page':1}]}
        review={'reviewer':'Tester','source_sha256':'source','pdf_sha256':'wrong','facts_checked':True,'fonts_verified':True,'interpretation_accepted':True,'pages':[{'page':1,'passed':True,'evidence':'fixture'}]}
        f=self.work/'review.json';o.dump(f,review)
        args=type('Args',(),dict(folder=str(self.work),review=str(f)))()
        with patch.object(o,'verify',return_value=(manifest,[])):
            with self.assertRaises(o.Blocked):o.release(args)
            review['pdf_sha256']='pdf';review['pages']=[];o.dump(f,review)
            with self.assertRaises(o.Blocked):o.release(args)
            review['pages']=[{'page':1,'passed':True,'evidence':''}];o.dump(f,review)
            with self.assertRaises(o.Blocked):o.release(args)
            review['pages'][0]['evidence']='fixture';review['fonts_verified']=False;o.dump(f,review)
            with self.assertRaises(o.Blocked):o.release(args)
    def test_draft_cannot_release_with_review(self):
        m={'source_sha256':'source','pdf_sha256':'pdf','pages':[{'page':1}]}
        review={'reviewer':'Tester','source_sha256':'source','pdf_sha256':'pdf','facts_checked':True,'fonts_verified':True,'interpretation_accepted':True,'pages':[{'page':1,'passed':True,'evidence':'fixture only'}]}
        f=self.work/'review.json';o.dump(f,review);a=type('Args',(),dict(folder=str(self.work),review=str(f)))()
        with patch.object(o,'verify',return_value=(m,['草稿不能发布'])):
            with self.assertRaises(o.Blocked):o.release(a)
    def test_release_copies_only_bound_pair(self):
        # Gate unit test only: synthetic bytes and mocked verification, not official files.
        src=self.work/'unit.docx';pdf=self.work/'unit.pdf';src.write_bytes(b'unit source');pdf.write_bytes(b'unit pdf')
        m={'source':src.name,'pdf':pdf.name,'source_sha256':o.sha(src),'pdf_sha256':o.sha(pdf),'pages':[{'page':1}]}
        review={'reviewer':'Unit test','source_sha256':m['source_sha256'],'pdf_sha256':m['pdf_sha256'],'facts_checked':True,'fonts_verified':True,'interpretation_accepted':True,'pages':[{'page':1,'passed':True,'evidence':'synthetic unit fixture'}]}
        f=self.work/'review.json';o.dump(f,review);a=type('Args',(),dict(folder=str(self.work),review=str(f)))()
        with patch.object(o,'verify',return_value=(m,[])):self.assertEqual(o.release(a),0)
        delivered=self.work/'delivery'/m['source_sha256'][:16]
        self.assertEqual({p.name for p in delivered.iterdir()},{'unit.docx','unit.pdf','audit.json'})
        self.assertEqual(o.sha(delivered/src.name),o.sha(src));self.assertEqual(o.sha(delivered/pdf.name),o.sha(pdf))
    def test_doctor_failure_exit_code(self):
        f=self.work/'input.json';o.dump(f,doc())
        with patch.object(sys,'argv',['office.py','doctor',str(f)]),patch.object(o,'doctor',return_value={'missing_fonts':['黑体'],'soffice':'mock'}):
            self.assertEqual(o.main(),2)
    def test_supplied_font_bundle_integrity_and_names(self):
        status=bf.bundle_status();self.assertEqual(status['issues'],[])
        self.assertTrue({'楷体_GB2312','方正小标宋简体','仿宋_GB2312','Times New Roman'} <= {x['family'] for x in status['fonts']})
    def test_private_font_environment_with_spaces(self):
        old=os.environ.get('FONTCONFIG_FILE');path=self.work/'font config with spaces'
        env=bf.font_environment(path)
        self.assertEqual(os.environ.get('FONTCONFIG_FILE'),old)
        self.assertTrue(Path(env['FONTCONFIG_FILE']).is_file())
        if shutil.which('fc-list'):
            visible=o.font_inventory(env)
            for f in bf.bundle_status()['fonts']:self.assertIn(f['family'].casefold(),visible)
    def test_font_corruption_is_blocked(self):
        target=self.work/'copied fonts';shutil.copytree(bf.FONT_DIR,target)
        font=target/'times-new-roman.ttf';font.write_bytes(font.read_bytes()+b'changed')
        with patch.object(bf,'FONT_DIR',target):self.assertTrue(any('校验失败' in x for x in bf.bundle_status()['issues']))
    def test_bundled_font_missing_glyph_is_reported(self):
        s=doc();s['blocks']=[{'role':'body','text':'正文🚀'}];p=self.work/'glyph.docx';o.create_docx(s,p)
        self.assertTrue(any('仿宋_GB2312缺少字形' in x for x in bf.bundled_glyph_issues(p)))
    def test_font_add_preserves_bytes_and_reports_missing_family(self):
        target=self.work/'added fonts';src=bf.FONT_DIR/'kaiti-gb2312.ttf';result=af.add_fonts([src],target)
        self.assertEqual(o.sha(src),o.sha(target/'kaiti-gb2312.ttf'))
        self.assertIn('黑体',result['not_bundled']);self.assertEqual(result['fonts'][0]['family'],'楷体_GB2312')
    def test_pdf_font_substitution_is_reported(self):
        import fitz
        s=doc();source=self.work/'font.docx';o.create_docx(s,source)
        pdf=fitz.open();page=pdf.new_page();page.insert_text((40,40),'Replacement font',fontname='helv')
        path=self.work/'replacement.pdf';pdf.save(path);pdf.close()
        self.assertTrue(any('未使用目标包内字体' in x for x in bf.bundled_pdf_font_issues(source,path)))
    def test_footer_field_mutation_detected(self):
        p=self.work/'notice.docx';s=doc();o.create_docx(s,p)
        with o.ZipFile(p) as z:r=E.fromstring(z.read('word/footer1.xml'))
        r.find('.//w:fldSimple',o.NS).set(o.tag(o.W,'instr'),'1');o.xml_update(p,{'word/footer1.xml':r})
        self.assertTrue(any('动态PAGE' in x for x in o.audit_source(p,s)))
    @unittest.skipUnless(os.environ.get('CODEX_PRIMARY_RUNTIME_NODE_MODULES'),'requires Work Mode Artifact Tool')
    def test_xlsx_roundtrip_formulas_and_mutation(self):
        t=table();t['merge_columns']=[0]
        spec={'kind':'xlsx','sheets':[t]};p=self.work/'roster.xlsx';o.create_xlsx(spec,p,self.work)
        self.assertEqual(o.audit_source(p,spec),[])
        t2={'name':'Formula','title':'计算表','headers':['值','加倍'], 'rows':[[0,{'formula':'=A3*2','expected':0}],[3,{'formula':'=A4*2','expected':6}]],'widths_mm':[78,78],'organization':'示例组织','ordering':'source-confirmed'}
        s2={'kind':'xlsx','sheets':[t2]};q=self.work/'formula.xlsx';o.create_xlsx(s2,q,self.work)
        self.assertEqual(o.audit_source(q,s2),[])
        with o.ZipFile(q) as z:r=E.fromstring(z.read('xl/worksheets/sheet1.xml'))
        row=r.find('x:sheetData/x:row',o.NS);r.find('x:sheetData',o.NS).remove(row);o.xml_update(q,{'xl/worksheets/sheet1.xml':r})
        self.assertTrue(any('标题单元格缺失' in x for x in o.audit_source(q,s2)))
        o.create_xlsx(s2,q,self.work)
        with o.ZipFile(q) as z:r=E.fromstring(z.read('xl/worksheets/sheet1.xml'))
        title=r.find('.//x:c[@r="A1"]/x:is/x:r/x:t',o.NS);title.text='错误标题';o.xml_update(q,{'xl/worksheets/sheet1.xml':r})
        self.assertTrue(any('标题内容' in x for x in o.audit_source(q,s2)))
        o.create_xlsx(s2,q,self.work)
        with o.ZipFile(q) as z:r=E.fromstring(z.read('xl/worksheets/sheet1.xml'))
        r.find('x:pageMargins',o.NS).set('left','0');o.xml_update(q,{'xl/worksheets/sheet1.xml':r})
        self.assertTrue(any('页边距' in x for x in o.audit_source(q,s2)))

if __name__=='__main__':unittest.main(verbosity=2)
