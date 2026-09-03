# MDTT — Markdown 与 docx 互转命令行工具

## 触发条件

当用户要求将 Markdown（.md）文件转换为 Word（.docx）文档，或将 Word（.docx）文件转换为 Markdown 时，使用此 Skill。

## 工具位置

全局命令：`MDTT`（通过 `npm link` 注册，可在任意目录下使用）

## 基本用法

```bash
MDTT <文件名>.md [参数]                  Markdown 转 docx
MDTT <文件名>.docx [参数]                docx 转 Markdown
```

输出文件默认与源文件同目录、同名（仅扩展名变化）。

转化的文档是法律相关文档时请添加预设参数：`MDTT notes.md --preset legal`。

## 常用示例

### Markdown → docx

```bash
MDTT notes.md                              # 默认格式（与 legal 排版相同：仿宋四号、首行缩进两字符、行距1.28、H1居中；仅无页眉页脚页码）
MDTT notes.md --preset legal               # 法律文书预设（圣典律师事务所专用排版）
MDTT notes.md -o output/result.docx        # 指定输出路径
MDTT notes.md --overwrite                  # 覆盖已有文件
MDTT notes.md --preset legal --overwrite   # 使用法律文书预设并覆盖
```

### docx → Markdown

```bash
MDTT report.docx                           # 转为同名 .md 文件
MDTT report.docx -o output.md              # 指定输出路径
MDTT report.docx --overwrite               # 覆盖已有文件
```

## 预设方案（仅 md → docx 时有效）

| 预设                 | 说明                                                           |
| ------------------ | ------------------------------------------------------------ |
| `--preset legal`   | 法律文书：A4、宋体标题、仿宋正文四号、首行缩进2字符、页眉（圣典律师事务所 logo + 渐变色带）、页脚页码（五号） |
| `--preset report`  | 报告：A4、微软雅黑11pt、页脚居中页码                                        |
| `--preset compact` | 紧凑排版：五号字、1.5cm页边距、窄段距                                        |
| `--preset cover`   | 封面页：内容垂直居中、无页眉页脚页码                                           |

## 常用参数速查（仅 md → docx 时有效）

| 参数                          | 说明                    |
| --------------------------- | --------------------- |
| `-s A4 / A3 / letter`       | 页面尺寸                  |
| `-m 上,右,下,左`                | 页边距（cm）               |
| `-f 宋体`                     | 正文字体                  |
| `--font-size 四号`            | 正文字号（支持中文字号名）         |
| `--line-height 1.5`         | 行距倍数                  |
| `--indent 2`                | 首行缩进字符数               |
| `-p bottom`                 | 页码位置（top/bottom/none） |
| `--page-num-format 第X页/共Y页` | 页码格式                  |
| `-o 路径`                     | 指定输出文件路径              |
| `--overwrite`               | 覆盖已有输出文件              |
| `--help`                    | 查看全部参数                |

## 注意事项

- 输出文件已存在时必须加 `--overwrite`，否则报错并退出
- 支持的中文字号名：初号、小初、一号、小一、二号、小二、三号、小三、四号、小四、五号、小五、六号、小六
- 图片缺失时以占位文本提示，不中断转换
- docx → md 会保留标题、列表（含嵌套）、表格、加粗/斜体/删除线、超链接；图片提取到输出目录的 `MDPictures/` 文件夹（hash 命名，相对路径引用）；**含合并单元格的表格以 HTML `<table>` 形式嵌入**（GFM 管道表格无法表达 colspan/rowspan，主流渲染器可正常显示）；代码块和引用因 Word 格式限制无法自动识别
- 完整参数列表见 `MDTT --help`
