# 输入契约与运行要求

优先复制examples中的JSON。JSON中不得写注释，不得用猜测补全事实。示例内容不是实际活动通知。

## 运行依赖

- 已测试目标：ChatGPT Work的Linux运行环境；Python依赖见requirements.txt。
- Excel生成使用Work内置的`@oai/artifact-tool`和Node；脚本自动读取`CODEX_PRIMARY_RUNTIME_NODE`及`CODEX_PRIMARY_RUNTIME_NODE_MODULES`。其他环境须先使Node能解析该包；缺少时明确阻断，不另换库拼出未经验证的文件。
- PDF使用PATH中的LibreOffice/soffice；逐页图片使用PyMuPDF。字体预检依赖fontconfig的fc-list。`assets/fonts/`中的字体通过独立FONTCONFIG_FILE提供给预检和排版引擎，不改系统字体配置。Windows/macOS请安装字体目录中的字体并使用经验证的本机排版流程，不声称自动加载已经跨平台测试。
- 本项目使用方正小标宋简体、仿宋_GB2312、楷体_GB2312、Times New Roman和黑体（SimHei）；公开仓库不附带字体二进制，须自行取得授权后放入`assets/fonts/`。不得把其他字体或伪造别名当成目标字体；包内文件、内部名称和字形均须校验；详见fonts.md。黑体来源为非商业公开字体包，商业分发前须核实授权。
- Work Mode执行命令时将示例中的python替换为`$CODEX_PRIMARY_RUNTIME_PYTHON`。首次实际制作前，执行本环境Word/Excel技能要求的制作标记；本包不硬编码平台内部标记路径。

## DOCX

顶层必填：`kind:"docx"`、`title`、`organization`、`date`、`blocks`。date为真实ISO日期，如`2026-10-18`。

可选：`subtitle`、`meta_date`（显示用中文日期）、`author`、`confidential`（内部资料或保密资料）、`scope`、`attachments`。

blocks支持：

| role | 必填字段 | 说明 |
|---|---|---|
| body、h1、h2、h3、h4 | text | 层级标题必须带对应编号；拒绝错号 |
| table | 下述表格字段 | Word中拒绝Excel公式对象 |
| image | path、caption | path为真实图片绝对路径；可选width_mm，最大156；不变形 |
| pagebreak | 无 | 另起一页；不得用大量空行模拟分页 |

text可为字符串，或`[{"text":"普通文字"},{"text":"强调文字","emphasis":true}]`。emphasis将中文改黑体，中西文字体仍按规范区分。

每个附件为`{"name":"附件名称","blocks":[...]}`，自动生成附件说明和独立附件页。图片默认接受4:3或16:9；其他真实比例须显式给`ratio_exception:true`并在审查记录中说明。

## XLSX与Word表格

XLSX顶层为`{"kind":"xlsx","sheets":[表格对象,...]}`。每个表格必填：

| 字段 | 类型与约束 |
|---|---|
| name | XLSX工作表名，不重复、不超31字符、无Excel非法字符 |
| title | 非空标题 |
| headers | 非空字符串数组 |
| rows | 等宽二维数组；每行长度与headers相同；保留全部记录 |
| widths_mm | 每列实际宽度，正数，合计不超过156毫米 |
| organization | 真实主要责任组织全称 |
| identifier_columns | 从0开始的学号/电话等标识列；其所有值必须为JSON字符串 |
| sort或ordering | 见下文，禁止默认猜测学院、组织顺序 |

- sort示例：`[{"column":0,"order":["文学院","外国语学院"]},{"column":2,"numeric":true}]`。第一项优先级最高，未知组织值会报错；数字排序用整数比较但保留字符串原值。
- 已核实无需重排时，显式写`"ordering":"source-confirmed"`。这表示操作者已确认原行序，不是跳过排序审查的借口。
- 可选`merge_columns:[0]`：只合并同组连续同值组织字段；不得合并标识列或公式列。Word生成器目前保留全部重复值而不自动合并，目检时按规范判断是否需要合并扩展。
- 可选`attachment_label:"附件 1"`：XLSX第一行附件标识、第二行空行、第三行表格标题；其后是表头。只有表头重复打印。
- 单元格允许字符串、有限数值、布尔值、null。空白与数值0不同。
- 公式用`{"formula":"=A3*2","expected":6}`；expected必须由独立计算得出。字符串`"=A3*2"`是原样文本，不执行为公式。公式单元格当前只支持数值结果；错误值或不一致的缓存结果会阻断。
- 含公式的表格必须提供最终行序和正确引用，不允许生成器在含公式时排序。不要把公式变成硬编码的结果值。
- 单个工作簿可以有多表；PDF页码须跨表连续。若目标引擎重置页码，PDF机检会阻断，走扩展流程修复后重导出。

## 输出与状态

- build输出：源文件、同名PDF、input.snapshot.json、manifest.json、preflight.json、preview逐页图片。机器校验通过后仍为REVIEW_REQUIRED，绝不自动生成“通过”的审查记录。
- review-template生成全部为false的review.json；只在真实审查后填写，不复制上一版本审查。
- release只在机器审查、版本绑定、完整逐页证据、事实与字体核对及执行约定确认均通过时，输出`delivery/源文件哈希前缀/`，其中有成对文件和audit.json。
- --draft仅用于测试代码，输出名带“待确认草稿”，状态DRAFT且release永久拒绝。它不证明目标字体、字形或最终版式合规。
- 自动审查不是全功能公文校对器：标题形状、视觉裁切、字形、名称与事实、装订方向等仍须真实逐页复核。字体名称存在也不能证明所有生僻字可显示。
