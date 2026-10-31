# 字体目录

本目录默认只保留 `manifest.json`。字体二进制不随公开仓库分发，原因和授权提示见根目录 `README.md` 与 `references/fonts.md`。

取得授权后，把文件按下列名称放入本目录：

- `kaiti-gb2312.ttf` → 楷体_GB2312
- `fzxiaobiaosong.ttf` → 方正小标宋简体
- `fangsong-gb2312.ttf` → 仿宋_GB2312
- `times-new-roman.ttf` → Times New Roman
- `simhei.ttf` → 黑体（SimHei）

`manifest.json` 记录了每个文件的 SHA-256、内部 family 和 embedding flags；`scripts/bundled_fonts.py` 会逐一核对，不会把改名文件当成目标字体。
