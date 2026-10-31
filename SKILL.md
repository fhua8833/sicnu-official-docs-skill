---
name: sicsu-official-docs
description: 按四川师范大学学生会2026年公文标准制作、审校Word和Excel；调用随附标准代码生成固定样式、执行自动检查、导出同源PDF，逐页审查后成对交付。适用于公文、通知、报名表、名单及附件；必须同时交付可编辑文件与对应PDF。
---

# 川师学生会公文制作

## 固定入口

先读 `references/format-standard.md`，需要原文时查 `assets/source-standard-2026.pdf`。再按 `references/input-contract.md` 填JSON；复制 `examples/notice.json` 或 `examples/roster.json` 修改内容，勿凭记忆重写生成器。

使用本环境指定Python（Work Mode优先 `$CODEX_PRIMARY_RUNTIME_PYTHON`；普通Linux使用已安装依赖的Python）。命令在技能根目录运行，路径有空格须加引号。

本项目使用楷体_GB2312、方正小标宋简体、仿宋_GB2312、Times New Roman和黑体。公开仓库不附带字体二进制；首次运行前须按`assets/fonts/README.md`放入已授权字体（文件名和内部名称须与`assets/fonts/manifest.json`一致），否则doctor会以缺字体阻断。Linux/Work环境的doctor与PDF导出会在本地私有加载，无须预先安装到系统；详见`references/fonts.md`。黑体为SimHei（GB18030，28,562字形，可编辑嵌入），取自浙江师范大学外国语学院公开字体包；不得用其他字体改名冒充。替换字体时运行`python scripts/add_fonts.py "黑体文件路径"`后重新预检。

```bash
python scripts/office.py doctor examples/notice.json
python scripts/office.py build examples/notice.json --out output/notice --name 通知
python scripts/office.py audit output/notice
python scripts/office.py review-template output/notice
python scripts/office.py release output/notice --review output/notice/review.json
```

制作Excel时，使用下面这一整组命令，后续检查始终指向本次同一输出目录：

```bash
python scripts/office.py doctor examples/roster.json
python scripts/office.py build examples/roster.json --out output/roster --name 名单
python scripts/office.py audit output/roster
python scripts/office.py review-template output/roster
python scripts/office.py release output/roster --review output/roster/review.json
```

## 固定SOP

1. 确认真实内容、署名日期、排序依据、附件关系和所需类型；缺关键事实时只询问该缺项。读取所用Word/Excel/PDF制作技能，遵守其工具要求。
2. 填输入JSON并运行doctor。自动校验字体目录中的字体的原始文件哈希和内部名称，并检查引擎实际可见字体；源文件生成后再核对字体目录中的字体字形覆盖。缺指定字体、字形或PDF引擎时阻断定稿，绝不整篇替换字体或修改字体内部名称。
3. 运行build。常见通知、正文、四级标题、署名、Word表格/附件/图片、Excel名单与数值公式由代码生成；页面、字号、行距、中西文字体、边框和页脚固定。复杂封面目录、横排装订及保密附件冲突须按扩展流程处理，不能静默省略。
4. 查看manifest.json和audit结果。DRAFT、BLOCKED、REVIEW_REQUIRED均不是验收通过。自动检查覆盖结构、数据、公式期望值、打印范围、PDF尺寸和页码；不替代事实及视觉审查。
5. 打开preview全部逐页图片，并与源文件打印预览对照；按规范检查标题、空行、续表、图片、署名、字体替换、缺字、裁切、留白、数据及全部工作表。真实确认每页后，填写review.json的逐页证据、事实核对、审查者和执行约定接受情况。禁止脚本批量把false改true冒充审查。
6. 发现格式问题优先修JSON，运行同参数repair重新生成；禁止只改PDF。源码只在确认生成器缺陷或扩展能力时修改，修改后运行test_office.py。所有源文件修改均须重导出并重审。
7. 运行release；只交付其返回目录中的DOCX/XLSX和同名PDF，可附audit.json。Word和Excel同主题时用不同后缀避免PDF覆盖。不得把工作目录草稿误当定稿。

## 强制门禁

- 正文对齐冲突按format-standard.md的执行约定处理，结论写“按本SOP执行约定通过”，不得冒充发布单位正式解释。
- build只直接导出最终源文件；manifest分别绑定源文件、PDF、输入快照和逐页图片的SHA-256。每次生成先记录新attempt；失败或未完成会阻断旧版本发布，并将旧delivery归档。后续改变任何受检文件，旧审查失效。
- release拦截缺字体、缺PDF、版本变化、机器错误、漏审页面、未解决问题、未确认事实或执行约定；不得伪填审查记录绕过。
- --draft仅供明确的开发测试；自动加“待确认草稿”后缀，永远不能release。缺环境条件时说明未完成，不交付只有源文件的“完整成果”。
- 学号/电话使用JSON字符串；不掌握组织顺序不猜；先排序再合并。公式必须给独立计算的expected，含公式时不得由生成器重排。
- 不支持的场景按 `references/extensions.md` 扩展并重审，不把失败降级成自动通过。

## 回归验证

```bash
python scripts/test_office.py
```

测试包含正常生成和异常拦截。它验证代码行为，不代表未来输入的事实、视觉效果或字体已通过审查。
