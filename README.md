# 川师学生会公文规范 Skill

按《四川师范大学学生会公文排版格式标准（2026年版）》生成、审校 Word / Excel，并强制同源导出 PDF。仓库内容是可执行的 Skill：格式规则、生成器、机检脚本、示例和回归测试。

> 这是按标准固化的排版工具，不是发布单位的正式解释。出现对齐等歧义时，以发布单位解释为准；本仓库沿用 `references/format-standard.md` 记录的执行约定。

## 包含什么

- `SKILL.md`：Skill 入口、固定 SOP 和交付门禁。
- `references/format-standard.md`：固化的格式参数和审查清单。
- `references/fonts.md`：字体名称、来源、哈希和加载方式。
- `references/input-contract.md`、`references/extensions.md`：输入 JSON 契约和扩展场景。
- `scripts/office.py`：DOCX/XLSX 生成、机检、PDF 导出和发布门禁。
- `scripts/test_office.py`：回归测试。
- `examples/`：通知、名单两个输入示例。

## 字体随仓库分发

本仓库直接包含 `assets/fonts/` 下的 5 个字体文件。`doctor` 会核对文件名、SHA-256 和内部 family，Linux/Work 环境通过独立 Fontconfig 私有加载，不需要安装到系统：

| 文件 | 字体内部名称 |
|---|---|
| `kaiti-gb2312.ttf` | 楷体_GB2312 |
| `fzxiaobiaosong.ttf` | 方正小标宋简体 |
| `fangsong-gb2312.ttf` | 仿宋_GB2312 |
| `times-new-roman.ttf` | Times New Roman |
| `simhei.ttf` | 黑体（SimHei） |

字体权利归各自厂商。`references/fonts.md` 记录了来源和授权提示；公开再分发前请确认你已获得相应授权，尤其是方正字体和 Times New Roman。

原始标准 PDF 未放入仓库：如需原文，把有公开权限的版本放到 `assets/source-standard-2026.pdf`；代码不依赖该文件，`references/format-standard.md` 已固化执行参数。

## 快速开始

```bash
python -m pip install -r requirements.txt

python scripts/office.py doctor examples/notice.json
python scripts/office.py build examples/notice.json --out output/notice --name 通知
python scripts/office.py audit output/notice
python scripts/office.py review-template output/notice
python scripts/office.py release output/notice --review output/notice/review.json
```

Excel 示例依赖 Work Mode 的 `@oai/artifact-tool`；普通环境没有该包时，Excel 用例会按设计跳过，不会用未经验证的库硬拼。

## 测试

```bash
python scripts/test_office.py
```

测试验证代码行为，不代表具体输入的事实、视觉效果或字体已通过审查。完整测试需要放入上述字体文件；缺字体时与字体相关的用例会失败或跳过。

## 权利说明

代码按原样提供，未附开源许可证；如需 MIT/Apache 等许可证请先确认。规范文本、字体和原始标准 PDF 的权利归各自权利人，不在本仓库的授权范围内。使用前请确认你所在单位允许公开这些内容，并遵守字体厂商的授权条款。
