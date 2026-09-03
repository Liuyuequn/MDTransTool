// 校验脚本：三组用例（默认转换 / preset legal / 自定义参数），检查生成的 docx 内部结构
import fs from "node:fs";
import path from "node:path";
import { execFileSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const JSZip = require("jszip");

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, "..");
const cli = path.join(root, "src", "cli.js");
const mdPath = path.join(__dirname, "sample.md");

let failed = 0;
function check(name, cond) {
  console.log(`${cond ? "通过" : "失败"} | ${name}`);
  if (!cond) failed++;
}

async function convert(outName, extraArgs = []) {
  const outPath = path.join(__dirname, outName);
  if (fs.existsSync(outPath)) fs.unlinkSync(outPath);
  execFileSync(process.execPath, [cli, mdPath, "to", "docx", "-o", outPath, ...extraArgs], { stdio: "pipe" });
  const zip = await JSZip.loadAsync(fs.readFileSync(outPath));
  const headerFiles = Object.keys(zip.files).filter((f) => /^word\/header\d+\.xml$/.test(f));
  const footerFiles = Object.keys(zip.files).filter((f) => /^word\/footer\d+\.xml$/.test(f));
  const readAll = (files) => files.map((f) => zip.file(f).async("string"));
  return {
    document: await zip.file("word/document.xml").async("string"),
    styles: await zip.file("word/styles.xml").async("string"),
    numbering: await zip.file("word/numbering.xml").async("string"),
    header: headerFiles.length ? (await Promise.all(readAll(headerFiles))).join("\n") : null,
    footer: footerFiles.length ? (await Promise.all(readAll(footerFiles))).join("\n") : null,
  };
}

function styleBlock(stylesXml, styleId) {
  const re = new RegExp(`<w:style [^>]*w:styleId="${styleId}"[\\s\\S]*?</w:style>`);
  const m = stylesXml.match(re);
  return m ? m[0] : "";
}

function docDefaults(stylesXml) {
  const m = stylesXml.match(/<w:docDefaults>[\s\S]*?<\/w:docDefaults>/);
  return m ? m[0] : "";
}

// ============ 用例 1：默认转换 ============
console.log("—— 用例 1：默认转换 ——");
const d = await convert("sample-default.docx");
[
  ["标题样式 Heading1", d.document.includes('<w:pStyle w:val="Heading1"/>')],
  ["六级标题 Heading6", d.document.includes('<w:pStyle w:val="Heading6"/>')],
  ["加粗", d.document.includes("<w:b/>")],
  ["斜体", d.document.includes("<w:i/>")],
  ["删除线", d.document.includes("<w:strike/>")],
  ["行内代码底纹", d.document.includes("EFEFEF")],
  ["代码块底纹", d.document.includes("F2F2F2")],
  ["表格", d.document.includes("<w:tbl>")],
  ["表头底纹", d.document.includes("EFEFEF")],
  ["图片", d.document.includes("<w:drawing>")],
  ["超链接", d.document.includes("<w:hyperlink")],
  ["引用左边框", d.document.includes("w:left")],
  ["有序列表编号", d.numbering.includes("%1.")],
  ["无序列表编号", d.numbering.includes("\u2022")],
  ["A4 页面尺寸", d.document.includes('w:w="11906"')],
  ["默认无页眉", d.header === null],
  ["默认无页脚", d.footer === null],
  ["默认无页码字段", !d.document.includes("PAGE")],
  ["缺失图片容错提示", d.document.includes("图片缺失")],
].forEach(([n, c]) => check(n, c));

// ============ 用例 2：preset legal ============
console.log("\n—— 用例 2：preset legal ——");
const l = await convert("sample-legal.docx", ["--preset", "legal"]);
const h1 = styleBlock(l.styles, "Heading1");
const h4 = styleBlock(l.styles, "Heading4");
const lDefaults = docDefaults(l.styles);
[
  // 纸张
  ["A4 尺寸", l.document.includes('w:w="11906"')],
  ["上下页边距 2.54cm（1440 twip）", l.document.includes('w:top="1440"') && l.document.includes('w:bottom="1440"')],
  ["左右页边距 3.18cm（1803 twip）", l.document.includes('w:left="1803"') && l.document.includes('w:right="1803"')],
  // 标题
  ["标题宋体", h1.includes('w:eastAsia="宋体"')],
  ["标题西文 Times New Roman", h1.includes('w:ascii="Times New Roman"')],
  ["H1 二号（44 半磅）", h1.includes('w:val="44"')],
  ["H1 居中", h1.includes('<w:jc w:val="center"/>')],
  ["H1 加粗", h1.includes("<w:b/>")],
  ["H4 四号且不加粗", h4.includes('w:val="28"') && !h4.includes("<w:b/>")],
  // 正文
  ["正文仿宋", lDefaults.includes('w:eastAsia="仿宋"')],
  ["正文西文 Times New Roman", lDefaults.includes('w:ascii="Times New Roman"')],
  ["正文四号（28 半磅）", lDefaults.includes('<w:sz w:val="28"/>')],
  ["首行缩进两字符（firstLineChars=200）", l.document.includes('w:firstLineChars="200"')],
  ["行距 1.28（w:line=307 auto）", lDefaults.includes('w:line="307"') && lDefaults.includes('w:lineRule="auto"')],
  // 页眉页脚
  ["页眉三行文字（所名/官网/地址）", l.header && l.header.includes("圣典律师事务所") && l.header.includes("圣典官网") && l.header.includes("总所地址") && l.header.includes("新城科技园4A栋6楼、7楼")],
  ["页眉左对齐", l.header && l.header.includes('<w:jc w:val="left"/>')],
  ["页眉 logo 图片", l.header && l.header.includes("<w:drawing>")],
  ["页眉 URL 为超链接", l.header && l.header.includes("<w:hyperlink")],
  ["页眉小五仿宋（18 半磅）", l.header && l.header.includes('w:val="18"') && l.header.includes('w:eastAsia="仿宋"')],
  ["页眉渐变色带为原生渐变（a:gradFill）", l.header && l.header.includes("<a:gradFill")],
  ["渐变色带起点亮正红 E60012", l.header && l.header.includes('val="E60012"')],
  ["渐变色带终点金色 FFD700", l.header && l.header.includes('val="FFD700"')],
  ["渐变色带含 4 个颜色停靠点", l.header && (l.header.match(/<a:gs /g) || []).length >= 4],
  ["渐变色带橙金占比（45% 起橙红、70% 起纯橙）", l.header && l.header.includes('pos="45000"') && l.header.includes('pos="70000"')],
  ["页脚“第X页/共Y页”模板", l.footer && l.footer.includes("第") && l.footer.includes("页/共")],
  ["页脚页码五号（21 半磅）", l.footer && l.footer.includes('w:val="21"')],
  ["页脚页码仿宋", l.footer && l.footer.includes('w:eastAsia="仿宋"')],
  ["页脚含页码字段", l.footer && l.footer.includes("PAGE")],
  ["页脚含总页数字段", l.footer && l.footer.includes("NUMPAGES")],
  ["无 titlePage（首页不差异化）", !l.document.includes("<w:titlePg/>")],
].forEach(([n, c]) => check(n, c));

// ============ 用例 3：自定义参数 ============
console.log("\n—— 用例 3：自定义参数 ——");
const c = await convert("sample-custom.docx", [
  "--v-align", "center",
  "-p", "bottom", "--page-num-format", "第X页", "--page-num-start", "3",
  "--header-left", "MDTT 测试", "--header-right", "2026-09-03",
  "--font", "黑体", "--font-size", "小四",
  "--indent", "2", "--align", "justify", "--no-italic",
  "--margin", "2,2,3,2",
]);
[
  ["页面垂直居中（v-align center）", c.document.includes('w:vAlign w:val="center"') || c.document.includes('w:val="center"')],
  ["自定义页边距（上 2cm=1134）", c.document.includes('w:top="1134"')],
  ["左分布页眉（tab 右对齐）", c.header && c.header.includes("MDTT 测试") && c.header.includes("2026-09-03")],
  ["页码格式“第X页”", c.footer && c.footer.includes("第") && c.footer.includes("PAGE")],
  ["起始页码 3", c.document.includes('w:start="3"')],
  ["正文黑体", c.styles.includes('w:eastAsia="黑体"')],
  ["正文小四（24 半磅）", c.styles.includes('w:val="24"')],
  ["首行缩进两字符", c.document.includes('w:firstLineChars="200"')],
  ["正文两端对齐", c.document.includes('w:val="both"')],
  ["--no-italic 生效（无斜体）", !c.document.includes("<w:i/>")],
].forEach(([n, cond]) => check(n, cond));

// ============ 用例 4：错误处理 ============
console.log("\n—— 用例 4：错误处理 ——");
function expectError(name, args) {
  try {
    execFileSync(process.execPath, [cli, ...args], { stdio: "pipe" });
    check(name, false);
  } catch (e) {
    check(name, e.status === 1 && e.stderr.length > 0);
  }
}
expectError("未知参数报错", [mdPath, "to", "docx", "--no-such-param"]);
expectError("互斥参数报错（--header 与 --header-left）", [mdPath, "to", "docx", "--header", "A", "--header-left", "B"]);
expectError("未知预设报错", [mdPath, "to", "docx", "--preset", "nope"]);
expectError("页边距格式错误报错", [mdPath, "to", "docx", "-m", "abc"]);

console.log(failed ? `\n${failed} 项校验未通过` : "\n全部校验通过");
process.exit(failed ? 1 : 0);
