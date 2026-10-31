# 字体与自动加载

## 预期文件

| 字体内部名称 | 预期文件 | 来源 |
|---|---|---|
| 楷体_GB2312 | assets/fonts/kaiti-gb2312.ttf | 用户上传，原样复制 |
| 方正小标宋简体 | assets/fonts/fzxiaobiaosong.ttf | 用户上传“方正小标宋简.TTF”，内部名称符合规范 |
| 仿宋_GB2312 | assets/fonts/fangsong-gb2312.ttf | 用户上传，原样复制 |
| Times New Roman | assets/fonts/times-new-roman.ttf | 浙江师范大学外国语学院同一公开包，完整版（4,728字形，fsType=8可编辑嵌入） |
| 黑体（SimHei） | assets/fonts/simhei.ttf | 浙江师范大学外国语学院《常用办公字体（字库）》（2024-11-10）公开包；GB18030，28,562字形，fsType=8可编辑嵌入 |

> 公开仓库只保留 `manifest.json`，不包含上表的字体二进制。请自行取得授权后按文件名放入 `assets/fonts/`，否则 `doctor` 会阻断。
>
> 来源页：https://flc.zjnu.edu.cn/2024/1110/c18580a485306/page.htm（2024-11-10）。来源页说明：部分字体仅供非商业用途，商业使用须自行向字体厂商取得授权。本项目按校内非商业公文制作使用；对外商业分发前须重新核实授权。
>
> 当前方正小标宋文件 `embedding_flags=2`（限制嵌入），PDF引擎可嵌入但授权需另行核实；同一来源包另有 `fsType=8` 版本，如替换须重新生成、导出和逐页审查。

## 使用方式

- Linux/当前Work环境：office.py自动调用bundled_fonts.py，生成独立Fontconfig配置。fc-list和LibreOffice使用相同字体目录，不需要管理员权限或全局安装，不影响系统已有字体配置。
- 原始字体文件有SHA-256清单，并核对内部family名称。缺失、损坏或错误文件不会自动改名“修复”。
- 字形检查针对实际分配给各字体的文字；GB2312之外的生僻字、繁体字等可能缺字。遇到缺字按原规范单字替代例外流程处理，不整篇换字库。黑体已覆盖GB18030，但其他字体仍可能缺字。
- PDF自动比对源文件实际用到的字体与PDF中的PostScript名称，并检查对应字库是否嵌入；整款目标字体丢失、被替换或未嵌入时阻断。该检查不等于逐字排版验证，仍需目检；未纳入字体清单的字体不会被自动核验。
- 新增或更换字体后必须重新生成文件、导出PDF、逐页审查；旧字体清单绑定的验收失效。
- 替换黑体示例：`python scripts/add_fonts.py "/实际路径/黑体.ttf"`。脚本读取内部名称后原样复制，并更新完整性清单；随后重新运行doctor。仅把一个文件重命名为simhei.ttf不能替代这一步。
- 当前自动私有加载实现已面向Linux/fontconfig；Windows和macOS可安装assets/fonts内的原始字体并重启Office，但不能把Linux自动加载测试冒称跨平台测试。

## 源文件与PDF

私有字体加载解决制作机器找不到这些字体的问题，不等于把字体自动嵌入DOCX/XLSX。另台机器继续编辑源文件时，仍须加载或安装字体目录中的对应文件。字体原始嵌入标志保持不变；不改标志强制嵌入。PDF须核对实际字体及字形，不能仅因导出成功就宣称显示完全一致。

实现依据：[Fontconfig配置文档](https://fontconfig.pages.freedesktop.org/fontconfig/fontconfig-user.html)。
