# MDTT — Markdown Trans Tool

## 一、基础信息

MDTT 是一款 Markdown 与 Word（docx）互转命令行工具。支持双向转换：

- **Markdown → docx**：`MDTT file.md [参数]`，可深度定制版式（页面、字体、页眉页脚等）

- **docx → Markdown**：`MDTT file.docx to md [参数]`，保留标题、列表、表格、加粗/斜体、超链接、图片等结构

### 技术栈

- [markdown-it](https://github.com/markdown-it/markdown-it) — Markdown 解析（md → docx）

- [docx](https://github.com/dolanmiu/docx) — docx 文档生成（md → docx）

- [image-size](https://github.com/image-size/image-size) — 图片尺寸读取（用于按比例缩放）

- [mammoth](https://github.com/mwilliamson/mammoth.js) — docx → HTML 提取（docx → md）

- [node-html-parser](https://github.com/taoqf/node-html-parser) — HTML DOM 预处理（表格结构重建与合并单元格检测）

- [turndown](https://github.com/mixmark-io/turndown) + [turndown-plugin-gfm](https://github.com/mixmark-io/turndown-plugin-gfm) — HTML → Markdown（docx → md）

### 项目结构

```
MDTransTool/
├── .github/workflows/ci.yml  # CI：push/PR 时在 Ubuntu/Windows × Node 18/20/22 上自动测试
├── src/
│   ├── cli.js                # 命令行入口：双向转换路由、位置参数校验
│   ├── converter.js          # md → docx 编排层：组装 Document
│   ├── blocks.js             # md → docx 块级解析（标题/列表/引用/代码块/表格）
│   ├── inline.js             # md → docx 行内解析（强调/链接/图片）与图片加载
│   ├── header-footer.js      # md → docx 页眉页脚（文字/图片布局/页码/渐变色带）
│   ├── styles.js             # md → docx 样式表、字体对象、间距换算、编号
│   ├── page.js               # md → docx 页面属性（尺寸/边距/垂直对齐/页码格式）
│   ├── docx-to-md.js         # docx → md 核心（mammoth 提取 → turndown 转换）
│   ├── args.js               # 参数解析：-- 参数规格表与校验
│   ├── presets.js            # 预设方案（legal / report / compact / cover / default）
│   └── options.js            # 默认配置、单位换算、深合并
└── test/
    ├── sample.md             # 测试样例（覆盖全部支持的语法）
    ├── unit-test.mjs         # 单元测试（纯函数层：换算/字号/合并/参数解析）
    ├── run-test.mjs          # 端到端校验（五组用例，检查 docx 内部结构）
    └── assets/               # 测试图片
```

### 系统要求

- Node.js >= 18

## 二、本地安装

在本项目根目录执行：

```bash
npm install
npm link
```

`npm link` 会将 `MDTT` 命令注册为全局命令，之后即可在任意目录、任意终端中使用。

## 三、使用方法

### 基本用法

```bash
MDTT <文件名>.md [参数]                  Markdown 转 docx
MDTT <文件名>.docx to md [参数]          docx 转 Markdown
```

示例：

```bash
MDTT notes.md                                        # 使用默认格式转为 docx
MDTT notes.md --preset legal                         # 使用法律文书预设转为 docx
MDTT notes.md --font 宋体 -p bottom --page-num-format 第X页
MDTT report.docx to md                               # docx 转为 Markdown
MDTT report.docx to md -o output.md                  # 指定输出路径
```

转换完成后，docx 文件默认输出到源文件所在目录，文件名与源文件相同（扩展名变为 `.docx`）；可用 `-o` 指定其他路径。

**docx → Markdown 的图片处理**：文档中的图片会提取到输出目录下的 `MDPictures/` 文件夹，以内容 hash 命名（如 `img-a1b2c3d4e5.png`），正文中以相对路径引用。相同图片只保存一份，重复转换文件名稳定；该文件夹可手动清理，下次转换会按需重建。

**docx → Markdown 的合并单元格**：GFM 管道表格语法无法表达 colspan/rowspan，因此含合并单元格的表格会以 HTML `<table>` 形式嵌入输出（Markdown 原生支持内嵌 HTML），完整保留合并语义，GitHub / VS Code / Typora 等主流渲染器均可正常显示；表格前附注释说明。无合并的表格仍转换为标准 GFM 管道表格。

查看帮助（含全部参数说明）：

```bash
MDTT --help
```

错误场景（文件不存在、位置参数格式不正确、参数取值非法、互斥参数冲突、输出文件已存在且未加 `--overwrite`）均输出中文提示并以非零状态码退出。

### 支持的 Markdown 语法

#### 块级元素

| 语法                    | 转换效果                                |
| --------------------- | ----------------------------------- |
| 六级标题（`#` \~ `######`） | Word 内置标题样式 Heading 1-6，字号/对齐/加粗可定制 |
| 有序 / 无序列表（含嵌套）        | Word 原生编号列表，层级缩进                    |
| 引用（`>`，含嵌套）           | 左侧竖线 + 左缩进样式                        |
| 围栏代码块                 | 等宽字体 + 灰色底纹（字体字号可定制）                |
| 表格（含对齐方式）             | Word 表格，表头加粗 + 灰色底纹，支持左/中/右对齐       |
| 分隔线（`---`）            | 段落底部横线                              |

#### 行内元素

| 语法                        | 转换效果                    |
| ------------------------- | ----------------------- |
| `**加粗**`、`*斜体*`、`~~删除线~~` | 对应 Word 字体样式（可分别关闭）     |
| `` `行内代码` ``              | 等宽字体 + 灰色底纹             |
| `[链接](url)`               | 可点击的超链接（蓝色下划线）          |
| `![图片](路径)`               | 嵌入图片，自动缩放至不超过页宽，独立成段时居中 |
| 行尾两个空格（硬换行）               | Word 内换行                |

#### 其他特性

- **中英文混排**：中西文字体分别设置（如仿宋 + Times New Roman）

- **图片容错**：图片缺失时以 `[图片缺失: ...]` 文本占位，不中断转换

- **首页差异化**：`--first-*` / `--no-first-*` 系列参数实现首页页眉页脚独立设置

## 四、参数列表

### 页面设置

| 参数              | 别名     | 取值                                              | 说明                                                  |
| --------------- | ------ | ----------------------------------------------- | --------------------------------------------------- |
| `--page-size`   | `-s`   | `A4` / `A3` / `A5` / `letter` / `legal` / `宽,高` | 页面尺寸预设或自定义宽高（cm），如 `-s 21,29.7`                     |
| `--orientation` | <br /> | `portrait` / `landscape`                        | 纵向（默认）/横向                                           |
| `--margin`      | `-m`   | `上,右,下,左` 或单值                                   | 页边距（cm），如 `-m 2.54,3.18,2.54,3.18` 或 `-m 2.5`（四边统一） |
| `--v-align`     | <br /> | `top` / `center` / `bottom`                     | 页面内容垂直对齐（重心），默认 top                                 |

### 字体与字号

| 参数               | 别名     | 取值                   | 说明                         |
| ---------------- | ------ | -------------------- | -------------------------- |
| `--font`         | `-f`   | 字体名                  | 正文字体（中西文统一），如 `-f 宋体`      |
| `--font-heading` | <br /> | 字体名                  | 标题字体（默认跟随 `--font`）        |
| `--font-code`    | <br /> | 字体名                  | 代码块字体（默认 Consolas）         |
| `--font-size`    | <br /> | pt 或中文字号名            | 正文字号，如 `12` 或 `小四`         |
| `--heading-size` | <br /> | `auto` / 单值 / 六级逗号分隔 | 标题字号，如 `22,16,14,14,14,14` |
| `--code-size`    | <br /> | pt 或中文字号名            | 代码块字号                      |
| `--line-height`  | <br /> | 倍数                   | 行距，如 `1.5`                 |

支持的中文字号名：初号、小初、一号、小一、二号、小二、三号、小三、四号、小四、五号、小五、六号、小六。

### 页眉

| 参数                                           | 说明                                     |
| -------------------------------------------- | -------------------------------------- |
| `--header <文本>`                              | 页眉文字（默认居中），文本中的 URL 自动转为超链接，支持 `\n` 多行 |
| `--header-left <文本>` / `--header-right <文本>` | 页眉左右分布（与 `--header` 互斥）                |
| `--header-align <left\|center\|right>`       | 页眉对齐，默认 center                         |
| `--header-font <字体名>`                        | 页眉字体，默认跟随正文                            |
| `--header-size <字号>`                         | 页眉字号，默认 9pt                            |
| `--header-image <路径>`                        | 页眉右侧放置图片（相对 md 文件或绝对路径），文字自动左置         |
| `--header-line`                              | 显示页眉下横线（默认不显示）                         |
| `--first-header <文本>`                        | 首页页眉文字                                 |
| `--no-first-header`                          | 首页不显示页眉                                |

### 页脚

与页眉基本对称：`--footer`、`--footer-left`、`--footer-right`、`--footer-align`、`--footer-font`、`--footer-size`（含页码字号）、`--footer-line`（页脚上横线）、`--first-footer`、`--no-first-footer`。页码位于页脚时使用页脚字体字号，位于页眉（`-p top`）时使用页眉字体字号。

### 页码

| 参数                                       | 别名                                                                                                | 说明           |
| ---------------------------------------- | ------------------------------------------------------------------------------------------------- | ------------ |
| `--page-number <none\|top\|bottom>`      | `-p`                                                                                              | 页码位置，默认 none |
| `--page-num-align <left\|center\|right>` | 页码对齐，默认 center                                                                                    | <br />       |
| `--page-num-format <模板>`                 | 模板中 `X`=当前页码、`Y`=总页数，如 `第X页`、`第X页/共Y页`、`X / Y`；特殊值 `1` / `i` / `I` / `一` 表示数字样式（阿拉伯/小写罗马/大写罗马/中文） | <br />       |
| `--page-num-start <数字>`                  | 起始页码                                                                                              | <br />       |
| `--total-pages`                          | 页码后附加总页数（格式无 `Y` 时追加 `/共Y页`）                                                                      | <br />       |
| `--no-first-page-number`                 | 首页不显示页码                                                                                           | <br />       |

### 正文对齐与间距

| 参数                                       | 说明                                              |
| ---------------------------------------- | ----------------------------------------------- |
| `--align <left\|center\|right\|justify>` | 正文对齐，默认左对齐                                      |
| `--heading-align <left\|center\|right>`  | 标题对齐（各级统一）                                      |
| `--para-spacing <pt>`                    | 段后间距                                            |
| `--heading-spacing <pt>`                 | 标题段前及段后间距                                       |
| `--indent <字符数>`                         | 首行缩进（仅普通正文段落；标题、列表、引用、代码块、表格不缩进），如 `--indent 2` |

### 样式开关

| 参数                    | 说明           |
| --------------------- | ------------ |
| `--no-bold`           | 标题不加粗        |
| `--no-italic`         | 斜体语法按普通文字渲染  |
| `--no-strike`         | 删除线语法按普通文字渲染 |
| `--heading-uppercase` | 标题英文转大写      |

### 输出控制

| 参数              | 别名                 | 说明       |
| --------------- | ------------------ | -------- |
| `--output <路径>` | `-o`               | 指定输出文件路径 |
| `--overwrite`   | 覆盖已存在的输出文件（默认报错提示） | <br />   |

## 五、预设方案

| 预设                 | 说明                                      |
| ------------------ | --------------------------------------- |
| `--preset legal`   | **法律文书**（详见下文完整规范）                      |
| `--preset report`  | 报告：A4、微软雅黑 11pt、1.15 倍行距、页脚居中页码 `X / Y` |
| `--preset compact` | 紧凑排版：五号字、1.5cm 页边距、窄段距，适合长文             |
| `--preset cover`   | 封面页：内容垂直居中，无页眉页脚页码                      |
| `--preset default` | 显式使用默认值                                 |

预设可与单项参数混用，**单项参数覆盖预设中的对应项**：

```bash
MDTT doc.md --preset legal --font 黑体 --header "保密文件"
```

参数优先级：命令行单项参数 > `--preset` 预设 > 默认值。

### legal 预设规范

| 项目   | 设定                                                                                                                                             |
| ---- | ---------------------------------------------------------------------------------------------------------------------------------------------- |
| 纸张   | A4 纵向；上下页边距 2.54cm，左右 3.18cm                                                                                                                   |
| 标题字体 | 中文宋体 / 西文 Times New Roman                                                                                                                      |
| 标题字号 | H1 二号、H2 三号、H3-H6 四号                                                                                                                           |
| 标题对齐 | H1 居中，其余左对齐                                                                                                                                    |
| 标题加粗 | H1-H3 加粗，H4-H6 不加粗                                                                                                                             |
| 标题间距 | 段前 0.5 行，段后 0 行                                                                                                                                |
| 正文字体 | 中文仿宋 / 西文 Times New Roman                                                                                                                      |
| 正文字号 | 四号                                                                                                                                             |
| 正文段落 | 首行缩进 2 字符；段后 0.5 行；行距 1.28 倍                                                                                                                   |
| 页眉   | 三行左对齐：①圣典律师事务所 ②圣典官网：<https://www.sundylawyer.com/（超链接）③总所地址：南京市建邺区奥体大街68号新城科技园4A栋6楼、7楼；右端放置律所> logo；页眉底端红→橙→金渐变色带（以红为主）；仿宋/Times New Roman 小五 |
| 页脚   | 居中页码「第X页/共Y页」，仿宋/Times New Roman 四号                                                                                                            |

## 六、运行测试

```bash
npm test
```

分两层共 127 项校验：

- **单元测试**（`test/unit-test.mjs`，39 项）：单位换算、中文字号解析、深合并、参数解析（别名、互斥、校验规则）

- **端到端校验**（`test/run-test.mjs`，88 项）：默认转换（19）、preset legal（30）、自定义参数组合（10）、错误处理（8）、docx → Markdown（含 MDPictures 图片提取，14）、合并单元格 HTML 回退（7）

期望值由预设声明与换算函数推导（换算函数由单元测试护航），避免测试与实现细节耦合。

### 持续集成

GitHub Actions（[ci.yml](.github/workflows/ci.yml)）在每次 push 与 pull request 时自动于 Ubuntu / Windows × Node 18 / 20 / 22 六个环境组合上运行全部测试。
