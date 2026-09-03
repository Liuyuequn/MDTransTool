#!/usr/bin/env node
// MDTT 命令行入口
// 格式：MDTT <文件>.md [参数]
// 参数优先级：命令行单项参数 > 预设（--preset）> 默认值（options.js）

import fs from "node:fs";
import path from "node:path";
import { convertMarkdownFile } from "./converter.js";
import { defaultOptions, mergeOptions } from "./options.js";
import { presets } from "./presets.js";
import { parseArgs, argsHelpText } from "./args.js";

const USAGE = `MDTT - Markdown Trans Tool

用法:
  MDTT <文件名>.md [参数]

示例:
  MDTT notes.md                                        使用默认格式转换
  MDTT notes.md --preset legal                         使用法律文书预设
  MDTT notes.md --font 宋体 -p bottom                  自定义字体并添加页码

预设方案:
  legal    法律文书（A4、宋体标题/仿宋正文、四号、首行缩进两字符、页眉页脚）
  report   报告（A4、微软雅黑、页脚页码）
  compact  紧凑排版（小字号、窄边距）
  cover    封面页（内容垂直居中、无页眉页脚页码）
  default  显式使用默认值

参数列表:
${argsHelpText()}

说明:
  - 尺寸参数（页边距、自定义页面宽高）为纯数字，单位 cm
  - 字号参数接受数字（pt）或中文字号名（如 四号、小五）
  - 页码格式模板中 X 表示当前页码、Y 表示总页数`;

const argv = process.argv.slice(2);

if (argv.length === 0 || argv.includes("--help") || argv.includes("-h")) {
  console.log(USAGE);
  process.exit(0);
}

// 位置参数（文件名）必须在最前，遇到第一个 - 开头的 token 后其余全部视为参数项
const firstOption = argv.findIndex((t) => t.startsWith("-"));
const positional = firstOption === -1 ? argv : argv.slice(0, firstOption);
const rest = firstOption === -1 ? [] : argv.slice(firstOption);

if (positional.length < 1) {
  console.error("错误: 请指定要转换的 .md 文件\n");
  console.log(USAGE);
  process.exit(1);
}

const inputPath = path.resolve(positional[0]);

if (!fs.existsSync(inputPath)) {
  console.error(`错误: 找不到文件 ${inputPath}`);
  process.exit(1);
}
if (path.extname(inputPath).toLowerCase() !== ".md") {
  console.error("错误: 仅支持 .md 文件。");
  process.exit(1);
}

// 解析参数（传原始 argv，位置参数已单独取出，参数项从 rest 中解析）
const parsed = parseArgs(rest);
if (parsed.errors.length) {
  for (const err of parsed.errors) console.error(`错误: ${err}`);
  process.exit(1);
}

// 预设校验与合并：单项参数 > 预设 > 默认值
if (parsed.preset != null && !presets[parsed.preset]) {
  console.error(`错误: 未知预设「${parsed.preset}」，可选：${Object.keys(presets).join(" / ")}`);
  process.exit(1);
}
const presetPatch = parsed.preset ? presets[parsed.preset] : {};
const opts = mergeOptions(defaultOptions, presetPatch, parsed.patch);

// 输出路径：默认与源文件同目录同名
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
