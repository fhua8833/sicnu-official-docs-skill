#!/usr/bin/env python3
"""Copy supplied fonts unchanged into this skill and regenerate integrity metadata."""
import argparse
import hashlib
import json
import shutil
from pathlib import Path
from fontTools.ttLib import TTFont
from bundled_fonts import FONT_DIR

FILES={'楷体_GB2312':'kaiti-gb2312.ttf','方正小标宋简体':'fzxiaobiaosong.ttf',
       '仿宋_GB2312':'fangsong-gb2312.ttf','Times New Roman':'times-new-roman.ttf','黑体':'simhei.ttf'}

def inspect(path):
    path=Path(path)
    with TTFont(path,lazy=True) as font:
        names=sorted({n.toUnicode() for n in font['name'].names if n.nameID in (1,16)})
        matched=[f for f in FILES if f in names]
        if len(matched)!=1:raise ValueError('不是规范要求的明确字体内部名称：'+str(path))
        family=matched[0]
        return {'file':FILES[family],'family':family,'internal_families':names,
                'sha256':hashlib.sha256(path.read_bytes()).hexdigest(),'size_bytes':path.stat().st_size,
                'embedding_flags':font['OS/2'].fsType,
                'postscript_names':sorted({n.toUnicode() for n in font['name'].names if n.nameID==6})}

def add_fonts(paths,directory=FONT_DIR):
    items=[(Path(p),inspect(p)) for p in paths]
    if len({x['family'] for _,x in items})!=len(items):raise ValueError('同一批输入字体重复')
    directory=Path(directory);directory.mkdir(parents=True,exist_ok=True)
    for src,item in items:
        dst=directory/item['file']
        if src.resolve()!=dst.resolve():shutil.copyfile(src,dst)
    fonts=[inspect(directory/n) for n in FILES.values() if (directory/n).is_file()]
    present={f['family'] for f in fonts}
    result={'fonts':fonts,'not_bundled':[f for f in FILES if f not in present]}
    (directory/'manifest.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    return result

if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__);parser.add_argument('fonts',nargs='+');args=parser.parse_args()
    print(json.dumps(add_fonts(args.fonts),ensure_ascii=False,indent=2))
