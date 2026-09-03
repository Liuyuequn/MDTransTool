// 校验脚本：三组用例（默认转换 / preset legal / 自定义参数），检查生成的 docx 内部结构
import fs from "node:fs";
import path from "node:path";
import { execFileSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { createRequire } from "node:module";

// 期望值统一由“配置声明 + 换算函数”推导，避免魔法数与实现耦合；
// 换算函数本身由 unit-test.mjs 覆盖
import { cmToTwip, ptToHalfPoint, PAGE_SIZES, CHINESE_FONT_SIZES, defaultOptions } from "../src/options.js";
import { presets } from "../src/presets.js";

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
  execFileSync(process.execPath, [cli, mdPath, "-o", outPath, ...extraArgs], { stdio: "pipe" });
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
// 约定：默认输出 = legal 排版 − 页眉页脚页码（期望值从 defaultOptions 推导）
const DF = defaultOptions;
const dfLineTwip = String(Math.round(DF.paragraph.line * 240));
const dfIndentChars = String(Math.round(DF.paragraph.firstLineChars * 100));
const dfAfterTwip = String(Math.round(DF.paragraph.afterLines * DF.sizes.body * DF.paragraph.line * 20));
const dH1 = styleBlock(d.styles, "Heading1");
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
  [`默认首行缩进两字符（firstLineChars=${dfIndentChars}）`, d.document.includes(`w:firstLineChars="${dfIndentChars}"`)],
  [`默认行距 ${DF.paragraph.line}（w:line=${dfLineTwip}）`, docDefaults(d.styles).includes(`w:line="${dfLineTwip}"`)],
  [`默认段后 ${DF.paragraph.afterLines} 行（w:after=${dfAfterTwip}）`, docDefaults(d.styles).includes(`w:after="${dfAfterTwip}"`)],
  ["默认 H1 居中", dH1.includes('<w:jc w:val="center"/>')],
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
// 期望值全部从 legal 预设声明推导
const LG = presets.legal;
const lgA4w = String(PAGE_SIZES[LG.page.size].width);
const lgMarginTop = String(cmToTwip(LG.page.margin.top));
const lgMarginLeft = String(cmToTwip(LG.page.margin.left));
const lgH1Half = String(ptToHalfPoint(LG.sizes.heading[0]));
const lgH4Half = String(ptToHalfPoint(LG.sizes.heading[3]));
const lgBodyHalf = String(ptToHalfPoint(LG.sizes.body));
const lgHeaderHalf = String(ptToHalfPoint(LG.sizes.header));
const lgFooterHalf = String(ptToHalfPoint(LG.sizes.footer));
const lgLineTwip = String(Math.round(LG.paragraph.line * 240));
const lgIndentChars = String(Math.round(LG.paragraph.firstLineChars * 100));
[
  // 纸张
  ["A4 尺寸", l.document.includes(`w:w="${lgA4w}"`)],
  ["上下页边距按预设（twip）", l.document.includes(`w:top="${lgMarginTop}"`) && l.document.includes(`w:bottom="${lgMarginTop}"`)],
  ["左右页边距按预设（twip）", l.document.includes(`w:left="${lgMarginLeft}"`) && l.document.includes(`w:right="${lgMarginLeft}"`)],
  // 标题
  ["标题宋体", h1.includes('w:eastAsia="宋体"')],
  ["标题西文 Times New Roman", h1.includes('w:ascii="Times New Roman"')],
  ["H1 字号按预设（半磅）", h1.includes(`w:val="${lgH1Half}"`)],
  ["H1 居中", h1.includes('<w:jc w:val="center"/>')],
  ["H1 加粗", h1.includes("<w:b/>")],
  ["H4 字号按预设且不加粗", h4.includes(`w:val="${lgH4Half}"`) && !h4.includes("<w:b/>")],
  // 正文
  ["正文仿宋", lDefaults.includes('w:eastAsia="仿宋"')],
  ["正文西文 Times New Roman", lDefaults.includes('w:ascii="Times New Roman"')],
  ["正文字号按预设（半磅）", lDefaults.includes(`<w:sz w:val="${lgBodyHalf}"/>`)],
  [`首行缩进两字符（firstLineChars=${lgIndentChars}）`, l.document.includes(`w:firstLineChars="${lgIndentChars}"`)],
  [`行距按预设（w:line=${lgLineTwip} auto）`, lDefaults.includes(`w:line="${lgLineTwip}"`) && lDefaults.includes('w:lineRule="auto"')],
  // 页眉页脚
  ["页眉三行文字（所名/官网/地址）", l.header && l.header.includes("圣典律师事务所") && l.header.includes("圣典官网") && l.header.includes("总所地址") && l.header.includes("新城科技园4A栋6楼、7楼")],
  ["页眉左对齐", l.header && l.header.includes('<w:jc w:val="left"/>')],
  ["页眉 logo 图片", l.header && l.header.includes("<w:drawing>")],
  ["页眉 URL 为超链接", l.header && l.header.includes("<w:hyperlink")],
  ["页眉字号字体按预设", l.header && l.header.includes(`w:val="${lgHeaderHalf}"`) && l.header.includes('w:eastAsia="仿宋"')],
  ["页眉渐变色带为原生渐变（a:gradFill）", l.header && l.header.includes("<a:gradFill")],
  ["渐变色带起点亮正红 E60012", l.header && l.header.includes('val="E60012"')],
  ["渐变色带终点金色 FFD700", l.header && l.header.includes('val="FFD700"')],
  ["渐变色带含 4 个颜色停靠点", l.header && (l.header.match(/<a:gs /g) || []).length >= 4],
  ["渐变色带橙金占比（45% 起橙红、70% 起纯橙）", l.header && l.header.includes('pos="45000"') && l.header.includes('pos="70000"')],
  ["页脚“第X页/共Y页”模板", l.footer && l.footer.includes("第") && l.footer.includes("页/共")],
  ["页脚页码字号按预设（半磅）", l.footer && l.footer.includes(`w:val="${lgFooterHalf}"`)],
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
  [`自定义页边距（上 2cm=${cmToTwip(2)} twip）`, c.document.includes(`w:top="${cmToTwip(2)}"`)],
  ["左分布页眉（tab 右对齐）", c.header && c.header.includes("MDTT 测试") && c.header.includes("2026-09-03")],
  ["页码格式“第X页”", c.footer && c.footer.includes("第") && c.footer.includes("PAGE")],
  ["起始页码 3", c.document.includes('w:start="3"')],
  ["正文黑体", c.styles.includes('w:eastAsia="黑体"')],
  [`正文小四（${ptToHalfPoint(CHINESE_FONT_SIZES["小四"])} 半磅）`, c.styles.includes(`w:val="${ptToHalfPoint(CHINESE_FONT_SIZES["小四"])}"`)],
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
// 在 expectError 基础上，额外校验 stderr 含指定提示文本（用于验证错误反馈的引导信息）
function expectErrorWith(name, args, text) {
  try {
    execFileSync(process.execPath, [cli, ...args], { stdio: "pipe" });
    check(name, false);
  } catch (e) {
    check(name, e.status === 1 && e.stderr.toString().includes(text));
  }
}
expectError("未知参数报错", [mdPath, "--no-such-param"]);
expectError("互斥参数报错（--header 与 --header-left）", [mdPath, "--header", "A", "--header-left", "B"]);
expectError("未知预设报错", [mdPath, "--preset", "nope"]);
expectError("页边距格式错误报错", [mdPath, "-m", "abc"]);
expectErrorWith("旧语法 md to docx 提示直接使用", [mdPath, "to", "docx"], "请直接使用");
expectErrorWith("旧语法 docx to md 提示直接使用", [path.join(__dirname, "sample-default.docx"), "to", "md"], "请直接使用");
expectError("多余位置参数报错", [mdPath, "垃圾参数"]);
expectErrorWith("无参数报错并提示引号包裹", [], "英文引号");
expectErrorWith("仅选项参数无文件报错", ["--preset", "legal"], "请指定要转换的文件");
expectErrorWith("找不到文件提示检查文件名", ["no-such-file.md"], "请检查文件名");
expectErrorWith(
  "文件名含空格未加引号提示整体包裹",
  ["no-such-dir", "我的", "文件.md"],
  'MDTT "no-such-dir 我的 文件.md"',
);
{
  // 不支持的扩展名：先造一个存在的 .txt 文件，确保走到扩展名校验分支而非“找不到文件”
  const txtPath = path.join(__dirname, "unsupported.txt");
  fs.writeFileSync(txtPath, "test", "utf-8");
  expectError("不支持的扩展名报错", [txtPath]);
  fs.unlinkSync(txtPath);
}
{
  // 无后缀文件存在：提示写明后缀（而非报“不支持的文件类型「」”）
  const noExtPath = path.join(__dirname, "no-extension-file");
  fs.writeFileSync(noExtPath, "test", "utf-8");
  expectErrorWith("无后缀文件提示写明后缀", [noExtPath], "请写明文件后缀");
  fs.unlinkSync(noExtPath);
}
{
  // 后缀自动匹配：仅 .md 存在 → 自动补全并转换成功；.md 与 .docx 共存 → 报错要求写明
  const stem = path.join(__dirname, "auto-ext-stem");
  const mdPath2 = `${stem}.md`;
  const docxPath2 = `${stem}.docx`;
  // 清理残留后仅创建 .md
  for (const p of [mdPath2, docxPath2, `${stem}-out.docx`]) { if (fs.existsSync(p)) fs.unlinkSync(p); }
  fs.writeFileSync(mdPath2, "# 自动匹配测试\n\n正文段落。\n", "utf-8");
  const stdout = execFileSync(process.execPath, [cli, stem, "-o", docxPath2], { stdio: "pipe", encoding: "utf-8" });
  check("省略后缀自动匹配 .md 并转换", stdout.includes("已自动匹配") && fs.existsSync(docxPath2));
  // 此时 .md 与 .docx 共存，再省略后缀应报错
  try {
    execFileSync(process.execPath, [cli, stem, "-o", `${stem}-out.docx`], { stdio: "pipe" });
    check("双后缀共存时报错要求写明", false);
  } catch (e) {
    check("双后缀共存时报错要求写明", e.status === 1 && e.stderr.toString().includes("写明后缀"));
  }
  // 无任何匹配 → 仍走“找不到文件”
  expectErrorWith("无匹配后缀提示找不到文件", [path.join(__dirname, "no-such-stem")], "找不到文件");
  fs.unlinkSync(mdPath2);
  fs.unlinkSync(docxPath2);
}

// ============ 用例 5：docx → Markdown ============
console.log("\n—— 用例 5：docx → Markdown ——");
{
  // 先用 sample.md 生成一个 docx，再转回 md
  const docxPath = path.join(__dirname, "docx2md-test.docx");
  const mdOutPath = path.join(__dirname, "docx2md-test.md");
  const picDir = path.join(__dirname, "MDPictures");
  // 清理残留
  for (const p of [docxPath, mdOutPath]) { if (fs.existsSync(p)) fs.unlinkSync(p); }
  if (fs.existsSync(picDir)) fs.rmSync(picDir, { recursive: true });
  // md → docx
  execFileSync(process.execPath, [cli, mdPath, "-o", docxPath], { stdio: "pipe" });
  check("docx→md 前置：生成测试 docx", fs.existsSync(docxPath));
  // docx → md
  execFileSync(process.execPath, [cli, docxPath, "-o", mdOutPath], { stdio: "pipe" });
  const mdOut = fs.readFileSync(mdOutPath, "utf-8");
  const pics = fs.existsSync(picDir) ? fs.readdirSync(picDir) : [];
  check("docx→md 输出文件存在", fs.existsSync(mdOutPath));
  check("docx→md 含一级标题", mdOut.includes("# "));
  check("docx→md 含二级标题", mdOut.includes("## "));
  check("docx→md 含加粗", mdOut.includes("**"));
  check("docx→md 含斜体", mdOut.includes("*"));
  check("docx→md 含超链接", mdOut.includes("](http"));
  check("docx→md 含图片引用", mdOut.includes("!["));
  check("docx→md 图片提取到 MDPictures 文件夹", pics.some((f) => f.startsWith("img-") && /\.(png|jpg|gif|bmp)$/.test(f)));
  check("docx→md 图片为相对路径引用", mdOut.includes("](MDPictures/img-"));
  check("docx→md 无 data URI 内嵌图片", !mdOut.includes("data:image"));
  check("docx→md 含列表项", mdOut.includes("- "));
  check("docx→md 含表格", mdOut.includes("| ---"));
  check("docx→md 无连续三个空行", !mdOut.includes("\n\n\n\n"));
  // 清理
  for (const p of [docxPath, mdOutPath]) { if (fs.existsSync(p)) fs.unlinkSync(p); }
  if (fs.existsSync(picDir)) fs.rmSync(picDir, { recursive: true });
}

// ============ 用例 6：docx → md 合并单元格（HTML 回退） ============
console.log("\n—— 用例 6：docx → md 合并单元格（HTML 回退）——");
{
  // 用 docx 库生成含横跨/纵跨合并单元格的测试文档
  const { Document, Packer, Paragraph, Table, TableRow, TableCell, WidthType } = await import("docx");
  const docxPath = path.join(__dirname, "merged-cells.docx");
  const mdOutPath = path.join(__dirname, "merged-cells.md");
  for (const p of [docxPath, mdOutPath]) { if (fs.existsSync(p)) fs.unlinkSync(p); }
  const doc = new Document({
    sections: [{
      children: [
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [
            new TableRow({
              children: [
                new TableCell({ children: [new Paragraph("横跨两列")], columnSpan: 2 }),
                new TableCell({ children: [new Paragraph("C1")] }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({ children: [new Paragraph("纵跨两行")], rowSpan: 2 }),
                new TableCell({ children: [new Paragraph("B2")] }),
                new TableCell({ children: [new Paragraph("C2")] }),
              ],
            }),
            new TableRow({
              children: [
                new TableCell({ children: [new Paragraph("B3")] }),
                new TableCell({ children: [new Paragraph("C3")] }),
              ],
            }),
          ],
        }),
      ],
    }],
  });
  fs.writeFileSync(docxPath, await Packer.toBuffer(doc));
  execFileSync(process.execPath, [cli, docxPath, "-o", mdOutPath], { stdio: "pipe" });
  const out = fs.readFileSync(mdOutPath, "utf-8");
  check("合并表格：输出文件存在", fs.existsSync(mdOutPath));
  check("合并表格：保留 colspan 属性", out.includes('colspan="2"'));
  check("合并表格：保留 rowspan 属性", out.includes('rowspan="2"'));
  check("合并表格：HTML 回退含 <table>", out.includes("<table"));
  check("合并表格：含说明注释", out.includes("合并单元格"));
  check("合并表格：内容完整（横跨两列/纵跨两行/B3）", out.includes("横跨两列") && out.includes("纵跨两行") && out.includes("B3"));
  check("合并表格：未误生成管道表格", !out.includes("| ---"));
  for (const p of [docxPath, mdOutPath]) { if (fs.existsSync(p)) fs.unlinkSync(p); }
}

console.log(failed ? `\n${failed} 项校验未通过` : "\n全部校验通过");
process.exit(failed ? 1 : 0);
