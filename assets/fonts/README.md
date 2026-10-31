# 字体目录

本目录随仓库提供 5 个字体文件和 `manifest.json`：

- `kaiti-gb2312.ttf` → 楷体_GB2312
- `fzxiaobiaosong.ttf` → 方正小标宋简体
- `fangsong-gb2312.ttf` → 仿宋_GB2312
- `times-new-roman.ttf` → Times New Roman
- `simhei.ttf` → 黑体（SimHei）

`manifest.json` 记录每个文件的 SHA-256、内部 family 和 embedding flags；`scripts/bundled_fonts.py` 会逐一核对，不会把改名文件当成目标字体。

字体权利归各自厂商，来源和授权提示见根目录 `README.md` 与 `references/fonts.md`。公开再分发前请确认你已获得相应授权。
