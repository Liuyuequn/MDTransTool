#!/usr/bin/env node
// MDTT 命令行入口
// 格式：MDTT <文件>.md [参数]              → md 转 docx
//       MDTT <文件>.docx [参数]            → docx 转 md
// 参数优先级：命令行单项参数 > 预设（--preset）> 默认值（options.js）

import fs from "node:fs";
import path from "node:path";
import { convertMarkdownFile } from "./converter.js";
import { convertDocxFile } from "./docx-to-md.js";
import { defaultOptions, mergeOptions } from "./options.js";
import { presets } from "./presets.js";
import { parseArgs, argsHelpText } from "./args.js";

const USAGE = `MDTT - Markdown Trans Tool

用法:
  MDTT <文件名>.md [参数]                  Markdown 转 docx
  MDTT <文件名>.docx [参数]                docx 转 Markdown

示例:
  MDTT notes.md                            使用默认格式转为 docx
  MDTT notes.md --preset legal             使用法律文书预设转为 docx
  MDTT report.docx                         docx 转为 Markdown
  MDTT report.docx -o output.md            docx 转为 Markdown 并指定输出路径

预设方案（仅 md → docx 时有效）:
  legal    法律文书（A4、宋体标题/仿宋正文、四号、首行缩进两字符、页眉页脚）
  report   报告（A4、微软雅黑、页脚页码）
  compact  紧凑排版（小字号、窄边距）
  cover    封面页（内容垂直居中、无页眉页脚页码）
  default  显式使用默认值

参数列表（仅 md → docx 时有效）:
${argsHelpText()}

说明:
  - 尺寸参数（页边距、自定义页面宽高）为纯数字，单位 cm
  - 字号参数接受数字（pt）或中文字号名（如 四号、小五）
  - 页码格式模板中 X 表示当前页码、Y 表示总页数`;

const argv = process.argv.slice(2);

// 引号使用提示：文件名含空格或特殊符号时，须用英文引号将「文件名+后缀」整体包裹
// （半包裹如 "文件.md".docx 会被 PowerShell 解析为属性访问，整个参数被丢弃，程序收到的将是空参数）
const QUOTE_TIP = '提示: 文件名包含空格或特殊符号时，请用英文引号将「文件名+后缀」整体包裹（后缀也必须在引号内），例如: MDTT "我的 文档.md"';

if (argv.includes("--help") || argv.includes("-h")) {
  console.log(USAGE);
  process.exit(0);
}
if (argv.length === 0) {
  console.error("错误: 请指定要转换的文件");
  console.error(QUOTE_TIP);
  console.log(USAGE);
  process.exit(1);
}

// 位置参数必须在最前，遇到第一个 - 开头的 token 后其余全部视为参数项
const firstOption = argv.findIndex((t) => t.startsWith("-"));
const positional = firstOption === -1 ? argv : argv.slice(0, firstOption);
const rest = firstOption === -1 ? [] : argv.slice(firstOption);

if (positional.length < 1) {
  console.error("错误: 请指定要转换的文件");
  console.error(QUOTE_TIP);
  console.log(USAGE);
  process.exit(1);
}

// 旧语法提示：MDTT <文件>.docx to md → 扩展名即转换方向，无需 to md
if (positional.length === 3 && positional[1] === "to" && ["md", "docx"].includes(positional[2].toLowerCase())) {
  console.error(`错误: 已不再需要「${positional.slice(1).join(" ")}」，扩展名即转换方向，请直接使用: MDTT ${positional[0]}`);
  process.exit(1);
}

const inputPath = path.resolve(positional[0]);

if (!fs.existsSync(inputPath)) {
  console.error(`错误: 找不到文件 ${inputPath}`);
  if (positional.length > 1) {
    // 文件名含空格未加引号时，会被终端拆成多段——重组原始输入并给出正确写法
    console.error(`提示: 文件名疑似包含空格但未加引号，被终端拆成了多段，请用英文引号将「文件名+后缀」整体包裹，如: MDTT "${positional.join(" ")}"`);
  } else {
    console.error("提示: 请检查文件名和路径是否输入正确");
  }
  process.exit(1);
}

const ext = path.extname(inputPath).toLowerCase();

// ============ 位置参数校验：扩展名即转换方向（.md→docx、.docx→md），只允许一个位置参数 ============
if (ext !== ".md" && ext !== ".docx") {
  console.error(`错误: 不支持的文件类型「${ext}」，仅支持 .md（转 docx）与 .docx（转 Markdown）`);
  process.exit(1);
}
if (positional.length !== 1) {
  console.error(`错误: 无法识别的参数「${positional.slice(1).join(" ")}」，用法：MDTT <文件名>${ext} [选项参数]`);
  console.error(QUOTE_TIP);
  process.exit(1);
}

// 解析参数（位置参数已单独取出，参数项从 rest 中解析）
const parsed = parseArgs(rest);
if (parsed.errors.length) {
  for (const err of parsed.errors) console.error(`错误: ${err}`);
  process.exit(1);
}

if (ext === ".docx") {
  // ============ docx → Markdown 模式 ============
  const outputPath = parsed.output
    ? path.resolve(parsed.output)
    : inputPath.replace(/\.docx$/i, ".md");

  if (fs.existsSync(outputPath) && !parsed.overwrite) {
    console.error(`错误: 输出文件已存在 ${outputPath}，如需覆盖请加 --overwrite`);
    process.exit(1);
  }

  try {
    const out = await convertDocxFile(inputPath, outputPath);
    console.log(`转换完成: ${out}`);
  } catch (err) {
    console.error(`转换失败: ${err.message}`);
    process.exit(1);
  }
} else {
  // ============ Markdown → docx 模式 ============
  // 预设校验与合并：单项参数 > 预设 > 默认值
  if (parsed.preset != null && !presets[parsed.preset]) {
    console.error(`错误: 未知预设「${parsed.preset}」，可选：${Object.keys(presets).join(" / ")}`);
    process.exit(1);
  }
  const presetPatch = parsed.preset ? presets[parsed.preset] : {};
  const opts = mergeOptions(defaultOptions, presetPatch, parsed.patch);

  const outputPath = parsed.output
    ? path.resolve(parsed.output)
    : inputPath.replace(/\.md$/i, ".docx");

  if (fs.existsSync(outputPath) && !parsed.overwrite) {
    console.error(`错误: 输出文件已存在 ${outputPath}，如需覆盖请加 --overwrite`);
    process.exit(1);
  }

  try {
    await convertMarkdownFile(inputPath, outputPath, opts);
    console.log(`转换完成: ${outputPath}`);
  } catch (err) {
    console.error(`转换失败: ${err.message}`);
    process.exit(1);
  }
}
