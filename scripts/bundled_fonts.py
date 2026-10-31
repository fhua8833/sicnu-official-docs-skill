"""Load the supplied fonts privately for Linux/fontconfig subprocesses.

Never rename a font's internal family or pretend a fallback is the target font.
The returned environment must be passed to BOTH fc-list and the PDF engine.
"""
from __future__ import annotations
import hashlib
import json
import os
from pathlib import Path
import xml.etree.ElementTree as E

FONT_DIR=Path(__file__).resolve().parents[1]/'assets'/'fonts'

def bundle_status():
    manifest=FONT_DIR/'manifest.json'
    if not manifest.is_file():return {'fonts':[],'issues':['包内字体清单缺失'],'manifest_sha256':None}
    try:data=json.loads(manifest.read_text(encoding='utf-8'))
    except (OSError,ValueError) as exc:return {'fonts':[],'issues':['字体清单损坏：'+str(exc)],'manifest_sha256':None}
    fonts=[];issues=[]
    for item in data.get('fonts',[]):
        name=item.get('file','');path=FONT_DIR/name
        if Path(name).name!=name or not name:issues.append('无效字体路径');continue
        if not path.is_file():issues.append('包内字体文件缺失：'+name);continue
        digest=hashlib.sha256(path.read_bytes()).hexdigest()
        if digest!=item.get('sha256'):issues.append('包内字体文件校验失败：'+name);continue
        # Metadata integrity is checked against the actual supplied font, not its filename.
        try:
            from fontTools.ttLib import TTFont
            with TTFont(path,lazy=True) as font:
                names={n.toUnicode() for n in font['name'].names if n.nameID in (1,16)}
            if item.get('family') not in names:issues.append('字体内部名称不匹配：'+name);continue
        except (OSError,ValueError,ImportError) as exc:
            issues.append('字体读取失败：'+name+'：'+str(exc));continue
        fonts.append({**item,'path':str(path.resolve())})
    return {'fonts':fonts,'issues':issues,'manifest_sha256':hashlib.sha256(manifest.read_bytes()).hexdigest()}

def font_environment(work):
    """Return an isolated environment; do not install fonts or alter system config."""
    work=Path(work).resolve();work.mkdir(parents=True,exist_ok=True)
    cache=work/'font-cache';cache.mkdir(exist_ok=True)
    config=work/'fonts.conf';root=E.Element('fontconfig')
    includes=[]
    for name in [os.environ.get('FONTCONFIG_FILE'),'/etc/fonts/fonts.conf']:
        if name and Path(name).is_file() and Path(name).resolve()!=config and name not in includes:
            E.SubElement(root,'include',ignore_missing='yes').text=str(Path(name).resolve());includes.append(name)
    directories=[FONT_DIR,Path('/usr/share/fonts'),Path('/usr/local/share/fonts')]
    runtime=os.environ.get('CODEX_PRIMARY_RUNTIME_ROOT')
    if runtime:directories.append(Path(runtime)/'dependencies/native/libreoffice-headless/libreoffice/share/fonts/truetype')
    for d in directories:
        if d.is_dir():E.SubElement(root,'dir').text=str(d.resolve())
    E.SubElement(root,'cachedir').text=str(cache)
    config.write_bytes(b'<?xml version="1.0"?>\n<!DOCTYPE fontconfig SYSTEM "urn:fontconfig:fonts.dtd">\n'+E.tostring(root,encoding='utf-8'))
    env=os.environ.copy();env['FONTCONFIG_FILE']=str(config)
    return env

def assigned_glyphs(source):
    """Collect actual text assigned to each font family in our generated OOXML."""
    from zipfile import ZipFile
    from collections import defaultdict
    assigned=defaultdict(set)
    w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    x='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    ns={'w':w,'x':x}
    with ZipFile(source) as z:
        if Path(source).suffix=='.docx':
            for name in z.namelist():
                if name=='word/document.xml' or (name.startswith('word/footer') and name.endswith('.xml')):
                    root=E.fromstring(z.read(name))
                    for run in root.findall('.//w:r',ns):
                        rf=run.find('w:rPr/w:rFonts',ns)
                        if rf is None:continue
                        for char in ''.join(t.text or '' for t in run.findall('w:t',ns)):
                            family=rf.get('{'+w+'}ascii') if char.isascii() else rf.get('{'+w+'}eastAsia')
                            if family:assigned[family].add(char)
        else:
            styles=E.fromstring(z.read('xl/styles.xml'));fonts=list(styles.find('x:fonts',ns));xfs=list(styles.find('x:cellXfs',ns))
            for name in z.namelist():
                if not (name.startswith('xl/worksheets/sheet') and name.endswith('.xml')):continue
                root=E.fromstring(z.read(name))
                for run in root.findall('.//x:is/x:r',ns):
                    f=run.find('x:rPr/x:rFont',ns)
                    if f is not None:assigned[f.get('val')].update(''.join(run.find('x:t',ns).itertext()))
                for c in root.findall('.//x:sheetData/x:row/x:c',ns):
                    if c.find('x:is',ns) is not None:continue
                    f=fonts[int(xfs[int(c.get('s','0'))].get('fontId','0'))].find('x:name',ns)
                    v=c.find('x:v',ns)
                    if f is not None and v is not None:assigned[f.get('val')].update(v.text or '')
            assigned['Times New Roman'].update('— 0123456789')
    return assigned

def bundled_glyph_issues(source):
    """Check assigned text against bundled cmap; return missing glyphs, not substitutes."""
    from fontTools.ttLib import TTFont
    status=bundle_status();issues=list(status['issues']);assigned=assigned_glyphs(source)
    for item in status['fonts']:
        with TTFont(item['path'],lazy=True) as font:cmap=font.getBestCmap() or {}
        missing=sorted(c for c in assigned[item['family']] if not c.isspace() and ord(c) not in cmap)
        if missing:issues.append(item['family']+'缺少字形：'+''.join(missing))
    return issues

def bundled_pdf_font_issues(source,pdf):
    """Require used bundled font faces to appear and be embedded in the PDF.

    This face-level check supplements cmap and visual checks; it does not prove
    per-character layout or support an unbundled family such as the missing 黑体.
    """
    import fitz
    import re
    status=bundle_status();assigned=assigned_glyphs(source);issues=[]
    def normalized(name):
        return ''.join(c.lower() for c in re.sub(r'^[A-Z]{6}\+','',name) if c.isalnum())
    with fitz.open(pdf) as doc:
        actual={}
        for page in doc:
            for f in page.get_fonts():actual.setdefault(normalized(f[3]),set()).add(f[0])
        for item in status['fonts']:
            if not any(not c.isspace() for c in assigned[item['family']]):continue
            names={normalized(n) for n in item['postscript_names']+item['internal_families']}
            xrefs={r for n,refs in actual.items() if n in names for r in refs}
            if not xrefs:issues.append('PDF未使用目标包内字体（可能被替换）：'+item['family'])
            elif not any(r>0 and bool(doc.extract_font(r)[3]) for r in xrefs):issues.append('PDF目标字体未嵌入：'+item['family'])
    return issues
