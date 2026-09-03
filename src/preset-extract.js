// 从 docx 提取版式为 MDTT 预设：解包 OOXML 原始 XML，逆向映射为与 presets.js 同形的补丁
// 入口（cli.js 路由）：MDTT <文件>.docx --save-preset <预设名>
// 自定义预设保存于 ~/.mdtt/presets/<名>.json；MDTT_HOME 环境变量可重定向（供测试隔离）

import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import JSZip from "jszip";
import { xml2js } from "xml-js";
import { PAGE_SIZES } from "./options.js";

const TWIP_PER_CM = 566.9291339;

// pgNumType w:fmt → MDTT 页码格式特殊值（纯数字样式）
const NUM_STYLE_FROM = {
  lowerRoman: "i",
  upperRoman: "I",
  chineseCounting: "一",
  chineseCountingThousand: "一",
  chineseLegalSimplified: "一",
};

// ================= xml-js 导航辅助 =================

function parseXml(xml) {
  return xml2js(xml, { compact: false });
}
function rootEl(doc) {
  return (doc.elements || []).find((e) => e.type === "element") ?? null;
}
function kids(el, name) {
  if (!el || !Array.isArray(el.elements)) return [];
  return el.elements.filter((e) => e.type === "element" && e.name === name);
}
function kid(el, name) {
  return kids(el, name)[0] ?? null;
}
function lastKid(el, name) {
  const all = kids(el, name);
  return all[all.length - 1] ?? null;
}
function attr(el, key) {
  return el?.attributes?.[key] ?? null;
}
function deepAll(el, name, out = []) {
  for (const c of el?.elements || []) {
    if (c.type !== "element") continue;
    if (c.name === name) out.push(c);
    deepAll(c, name, out);
  }
  return out;
}
function textOf(el) {
  return (el?.elements || []).filter((e) => e.type === "text").map((e) => e.text).join("");
}
const round2 = (v) => Math.round(v * 100) / 100;
const round1 = (v) => Math.round(v * 10) / 10;

// ================= rPr / pPr 读取 =================

/** rPr → 字体（中文 eastAsia / 西文 ascii） */
function fontFromRPr(rPr) {
  const rf = kid(rPr, "w:rFonts");
  if (!rf) return null;
  const eastAsia = attr(rf, "w:eastAsia");
  const ascii = attr(rf, "w:ascii") ?? attr(rf, "w:hAnsi");
  if (!eastAsia && !ascii) return null;
  return { eastAsia: eastAsia ?? ascii, ascii: ascii ?? eastAsia };
}
/** rPr → 字号（半磅值 → pt） */
function halfPtFromRPr(rPr) {
  const sz = attr(kid(rPr, "w:sz"), "w:val");
  return sz != null ? Number(sz) / 2 : null;
}
/** rPr → 是否加粗（w:val 为 false/0/none 视为不加粗） */
function isBold(rPr) {
  const b = kid(rPr, "w:b");
  if (!b) return false;
  return !["false", "0", "none"].includes(attr(b, "w:val"));
}
/** pPr → 间距（before/after/line twip 与 lineRule） */
function spacingOf(pPr) {
  const sp = kid(pPr, "w:spacing");
  if (!sp) return null;
  const num = (k) => (attr(sp, `w:${k}`) != null ? Number(attr(sp, `w:${k}`)) : null);
  return { before: num("before"), after: num("after"), line: num("line"), lineRule: attr(sp, "w:lineRule") };
}

// ================= 自定义预设存取 =================

/** 自定义预设目录：MDTT_HOME 可重定向（测试用），默认 ~/.mdtt/presets/ */
export function presetsDir() {
  return process.env.MDTT_HOME
    ? path.join(process.env.MDTT_HOME, "presets")
    : path.join(os.homedir(), ".mdtt", "presets");
}

/** 预设名校验：中英文/数字/下划线/连字符，不以连字符开头，≤64 字符 */
export function validPresetName(name) {
  return /^[A-Za-z0-9_\u4e00-\u9fff][\w\u4e00-\u9fff-]{0,63}$/.test(String(name ?? ""));
}

export function customPresetPath(name) {
  return path.join(presetsDir(), `${name}.json`);
}

/** 列出自定义预设名（按名称排序）；目录不存在返回空数组 */
export function listCustomPresets() {
  const dir = presetsDir();
  if (!fs.existsSync(dir)) return [];
  return fs.readdirSync(dir).filter((f) => f.endsWith(".json")).map((f) => f.slice(0, -5)).sort();
}

/** 读取自定义预设的 options 补丁；不存在返回 null，文件损坏抛错 */
export function loadCustomPreset(name) {
  const file = customPresetPath(name);
  if (!fs.existsSync(file)) return null;
  let data;
  try {
    data = JSON.parse(fs.readFileSync(file, "utf-8"));
  } catch (e) {
    throw new Error(`自定义预设文件损坏「${file}」: ${e.message}`);
  }
  if (!data || typeof data !== "object" || !data.options || typeof data.options !== "object") {
    throw new Error(`自定义预设文件格式不正确「${file}」（缺少 options 字段）`);
  }
  return data.options;
}

// ================= 提取核心 =================

/**
 * 从 docx 提取版式，返回 { options, notes }：
 * options 为与 presets.js 条目同形的补丁（经 mergeOptions 深合并后生效）
 * notes 为无法映射项的中文说明
 */
export async function extractPresetOptions(docxPath) {
  const zip = await JSZip.loadAsync(fs.readFileSync(docxPath));
  const docXml = await zip.file("word/document.xml")?.async("string");
  if (!docXml) throw new Error("不是有效的 docx 文件（缺少 word/document.xml）");
  const stylesXml = (await zip.file("word/styles.xml")?.async("string")) ?? "";
  const relsXml = (await zip.file("word/_rels/document.xml.rels")?.async("string")) ?? "";
  const rels = parseRels(relsXml);

  const docRoot = rootEl(parseXml(docXml));
  const body = kid(docRoot, "w:body");
  const sectPr = body ? lastKid(body, "w:sectPr") : null;
  const notes = [];
  if (!sectPr) notes.push("未找到页面设置（sectPr），页面参数未提取");

  const opts = {};

  // ---- 页面：尺寸 / 方向 / 边距 / 垂直对齐 ----
  if (sectPr) {
    const pgSz = kid(sectPr, "w:pgSz");
    if (pgSz) Object.assign(opts, { page: { ...opts.page, ...extractPageSize(pgSz) } });
    const pgMar = kid(sectPr, "w:pgMar");
    if (pgMar) {
      const cm = (k) => round2(Number(attr(pgMar, `w:${k}`)) / TWIP_PER_CM);
      opts.page = {
        ...opts.page,
        margin: { top: cm("top"), right: cm("right"), bottom: cm("bottom"), left: cm("left") },
      };
    }
    const vAlign = attr(kid(sectPr, "w:vAlign"), "w:val");
    if (vAlign === "center" || vAlign === "bottom") opts.page = { ...opts.page, vAlign };
    if (kid(sectPr, "w:titlePg")) notes.push("文档启用了首页不同（titlePg），首页页眉页脚差异未提取");
  }

  // ---- 正文默认样式：docDefaults 优先，Normal 样式兜底 ----
  const stylesRoot = stylesXml ? rootEl(parseXml(stylesXml)) : null;
  const docDefaults = stylesRoot ? kid(stylesRoot, "w:docDefaults") : null;
  const normalStyle = stylesRoot
    ? kids(stylesRoot, "w:style").find((s) => attr(s, "w:styleId") === "Normal") ?? null
    : null;
  const ddRPr = kid(kid(docDefaults, "w:rPrDefault"), "w:rPr");
  const ddPPr = kid(kid(docDefaults, "w:pPrDefault"), "w:pPr");
  const nRPr = normalStyle ? kid(normalStyle, "w:rPr") : null;
  const nPPr = normalStyle ? kid(normalStyle, "w:pPr") : null;

  const bodyFont = fontFromRPr(nRPr) ?? fontFromRPr(ddRPr);
  const bodySize = halfPtFromRPr(nRPr) ?? halfPtFromRPr(ddRPr);
  const bodySpacing = spacingOf(nPPr) ?? spacingOf(ddPPr);
  const bodyLine =
    bodySpacing?.lineRule === "auto" && bodySpacing.line != null
      ? round2(bodySpacing.line / 240)
      : null;
  if (bodySpacing?.line != null && bodyLine == null) {
    notes.push(`行距为固定值规则（lineRule=${bodySpacing.lineRule}），MDTT 仅支持倍数行距，未提取`);
  }

  if (bodyFont) opts.fonts = { body: bodyFont };
  if (bodySize != null) opts.sizes = { body: bodySize };

  // ---- 标题样式：六级字体 / 字号 / 加粗 / 对齐 / 间距 ----
  const hStyles = new Map();
  for (const st of stylesRoot ? kids(stylesRoot, "w:style") : []) {
    const m = /^Heading([1-6])$/.exec(attr(st, "w:styleId") ?? "");
    if (m) hStyles.set(Number(m[1]), st);
  }
  const headingSizes = [];
  const headingBold = [];
  const headingAlign = [];
  const hFonts = [];
  const beforeTw = [];
  const afterTw = [];
  for (let i = 1; i <= 6; i++) {
    const st = hStyles.get(i) ?? null;
    const rPr = st ? kid(st, "w:rPr") : null;
    const pPr = st ? kid(st, "w:pPr") : null;
    hFonts.push(fontFromRPr(rPr) ?? bodyFont ?? null);
    headingSizes.push(halfPtFromRPr(rPr) ?? bodySize ?? null);
    headingBold.push(st ? isBold(rPr) : false);
    const jc = attr(kid(pPr, "w:jc"), "w:val");
    headingAlign.push(jc === "center" ? "center" : jc === "right" ? "right" : null);
    const sp = spacingOf(pPr);
    beforeTw.push(sp?.before ?? null);
    afterTw.push(sp?.after ?? null);
  }
  if (hStyles.size) {
    const firstFont = hFonts.find(Boolean);
    if (firstFont) {
      const key = (f) => `${f?.eastAsia}/${f?.ascii}`;
      if (!hFonts.every((f) => f && key(f) === key(firstFont))) {
        notes.push("各级标题字体不一致，已采用一级标题字体");
      }
      opts.fonts = { ...opts.fonts, heading: firstFont };
    }
    opts.sizes = { ...opts.sizes, heading: headingSizes };
    opts.heading = {
      bold: headingBold,
      align: headingAlign,
      spacing: {
        beforeLines: null,
        afterLines: null,
        ...spacingPatch(beforeTw, headingSizes, bodyLine, "标题段前距", notes, "before"),
        ...spacingPatch(afterTw, headingSizes, bodyLine, "标题段后距", notes, "after"),
      },
    };
  }

  // ---- 段落：首行缩进 / 行距 / 段后距 / 对齐 ----
  const para = {};
  if (bodyLine != null) para.line = bodyLine;
  if (bodySpacing?.after != null) {
    if (bodyLine != null && bodySize != null) {
      para.afterLines = round2(bodySpacing.after / (bodySize * bodyLine * 20));
    } else {
      para.afterLines = null;
      para.afterPt = round2(bodySpacing.after / 20);
    }
  } else {
    para.afterLines = null;
    para.afterPt = 0; // 未声明段后距 = 无段后距
  }
  para.firstLineChars = extractFirstLineChars(body, bodySize);
  const bodyAlign = extractBodyAlign(body, nPPr, ddPPr);
  if (bodyAlign) para.align = bodyAlign;
  opts.paragraph = para;

  // ---- 页眉页脚：文字 / 对齐 / 字体字号 / 页码字段 ----
  const hdr = await parseHeaderFooter(zip, sectPr, rels, "header", notes);
  const ftr = await parseHeaderFooter(zip, sectPr, rels, "footer", notes);
  if (hdr?.text) opts.header = { text: hdr.text, align: hdr.align ?? "left" };
  if (hdr?.font) opts.fonts = { ...opts.fonts, header: hdr.font };
  if (hdr?.size != null) opts.sizes = { ...opts.sizes, header: hdr.size };
  if (ftr?.text) opts.footer = { text: ftr.text, align: ftr.align ?? "left" };
  if (ftr?.font) opts.fonts = { ...opts.fonts, footer: ftr.font };
  if (ftr?.size != null) opts.sizes = { ...opts.sizes, footer: ftr.size };

  // ---- 页码：字段所在位置（页眉/页脚）+ 模板 + 数字样式 + 起始页 ----
  const pnFrom = ftr?.pn
    ? { ...ftr.pn, pos: "bottom" }
    : hdr?.pn
      ? { ...hdr.pn, pos: "top" }
      : null;
  const pgNumType = sectPr ? kid(sectPr, "w:pgNumType") : null;
  const styleVal = NUM_STYLE_FROM[attr(pgNumType, "w:fmt")] ?? null;
  const startVal = attr(pgNumType, "w:start");
  if (pnFrom) {
    const pn = { pos: pnFrom.pos, align: pnFrom.align ?? "left", format: pnFrom.template || "X" };
    if (styleVal) {
      if (pnFrom.template && pnFrom.template !== "X") {
        notes.push("页码数字样式与文字模板无法并存，已保留数字样式");
      }
      pn.format = styleVal;
    }
    if (startVal != null) pn.start = Number(startVal);
    opts.pageNumber = pn;
  } else if (styleVal != null || startVal != null) {
    notes.push("检测到页码设置（数字样式/起始页）但未找到页码字段，未提取");
  }

  return { options: opts, notes };
}

/**
 * 间距提取：优先换算为"行"单位（docx 的 before/after twip ÷ 字号×行距×20），
 * 各级一致时返回 { xxxLines }；不一致或无法换算时退回 pt（取一级标题值）。
 * prefix 为 "before" / "after"。
 */
function spacingPatch(twArr, sizeArr, line, label, notes, prefix) {
  const firstIdx = twArr.findIndex((t) => t != null);
  if (firstIdx === -1) return {};
  const key = `${prefix}Lines`;
  if (line != null) {
    const asLines = twArr.map((tw, i) =>
      tw != null && sizeArr[i] ? tw / (sizeArr[i] * line * 20) : null
    );
    if (asLines.every((v) => v != null && Math.abs(v - asLines[firstIdx]) < 0.05)) {
      return { [key]: round2(asLines[firstIdx]) };
    }
  }
  const present = twArr.filter((t) => t != null);
  if (!present.every((t) => t === twArr[firstIdx])) {
    notes.push(`${label}各级不一致，已采用一级标题的值`);
  }
  return { [key]: null, [`${prefix}Pt`]: round2(twArr[firstIdx] / 20) };
}

/** 页面尺寸：匹配 PAGE_SIZES 命名尺寸；横向时先换算为纵向等效尺寸再匹配 */
function extractPageSize(pgSz) {
  const w = Number(attr(pgSz, "w:w"));
  const h = Number(attr(pgSz, "w:h"));
  if (!w || !h) return {};
  const orientation = attr(pgSz, "w:orient") ?? (w > h ? "landscape" : "portrait");
  const isLandscape = orientation === "landscape";
  const pw = isLandscape ? h : w;
  const ph = isLandscape ? w : h;
  for (const [name, dim] of Object.entries(PAGE_SIZES)) {
    if (dim.width === pw && dim.height === ph) return { size: name, orientation };
  }
  // 自定义尺寸："W,H"（cm，纵向等效）
  return { size: `${round1(pw / TWIP_PER_CM)},${round1(ph / TWIP_PER_CM)}`, orientation };
}

/** 正文首行缩进：统计正文顶层段落（跳过列表/样式/引用/代码块）最常见的缩进值 */
function extractFirstLineChars(body, bodySizePt) {
  const counts = new Map();
  const bump = (v) => counts.set(v, (counts.get(v) ?? 0) + 1);
  for (const p of kids(body ?? null, "w:p")) {
    const pPr = kid(p, "w:pPr");
    if (!pPr) { bump(0); continue; }
    // 列表（numPr）、标题/引用样式（pStyle）、代码块（shd 底纹）不计
    if (kid(pPr, "w:numPr") || kid(pPr, "w:pStyle") || kid(pPr, "w:shd")) continue;
    const ind = kid(pPr, "w:ind");
    if (!ind) { bump(0); continue; }
    const flc = attr(ind, "w:firstLineChars");
    if (flc) { bump(Number(flc) / 100); continue; }
    const fl = attr(ind, "w:firstLine");
    if (fl != null) { bump(bodySizePt ? round2(Number(fl) / (bodySizePt * 20)) : 0); continue; }
    // 仅有 left/right/hanging 缩进而无首行缩进：引用、代码块等 → 不计
  }
  if (!counts.size) return 0;
  let best = 0;
  let bestN = -1;
  for (const [v, n] of counts) {
    if (n > bestN) { best = v; bestN = n; }
  }
  return round2(best);
}

/** 正文对齐：顶层段落最常见 jc（需有过半数且至少 3 段），否则看 Normal/docDefaults */
function extractBodyAlign(body, nPPr, ddPPr) {
  const counts = new Map();
  for (const p of kids(body ?? null, "w:p")) {
    const pPr = kid(p, "w:pPr");
    if (!pPr || kid(pPr, "w:numPr") || kid(pPr, "w:pStyle")) continue;
    const jc = attr(kid(pPr, "w:jc"), "w:val");
    if (jc) counts.set(jc, (counts.get(jc) ?? 0) + 1);
  }
  let best = null;
  let total = 0;
  for (const [k, n] of counts) {
    total += n;
    if (!best || n > counts.get(best)) best = k;
  }
  const jc = total >= 3 && counts.get(best) >= total / 2 ? best
    : attr(kid(nPPr, "w:jc"), "w:val") ?? attr(kid(ddPPr, "w:jc"), "w:val");
  if (jc === "center" || jc === "right") return jc;
  if (jc === "both" || jc === "distribute") return "justify";
  return null; // left / 缺省 = 默认
}

/** rels：rId → Target（相对 word/ 的路径） */
function parseRels(relsXml) {
  const map = new Map();
  if (!relsXml) return map;
  const root = rootEl(parseXml(relsXml));
  for (const rel of kids(root, "Relationship")) {
    map.set(attr(rel, "Id"), attr(rel, "Target"));
  }
  return map;
}

/**
 * 段落内容分段：文字段与页码字段段（X=当前页，Y=总页数）。
 * 字段缓存文本（separate 与 end 之间的 w:t）跳过，避免把 Word 缓存的页码值混入模板。
 */
function segmentsOf(p) {
  const segs = [];
  let fieldState = "outside"; // outside → instr → cached
  for (const c of p.elements || []) {
    if (c.type !== "element") continue;
    if (c.name === "w:r") {
      const instr = kid(c, "w:instrText");
      if (instr) {
        const t = textOf(instr).toUpperCase();
        if (t.includes("NUMPAGES")) segs.push({ field: "Y" });
        else if (t.includes("PAGE")) segs.push({ field: "X" });
        continue;
      }
      if (kid(c, "w:fldChar")) {
        const ty = attr(kid(c, "w:fldChar"), "w:fldCharType");
        if (ty === "begin") fieldState = "instr";
        else if (ty === "separate") fieldState = "cached";
        else if (ty === "end") fieldState = "outside";
        continue;
      }
      if (kid(c, "w:tab")) { segs.push({ text: "\t" }); continue; }
      if (fieldState === "cached") continue;
      for (const t of kids(c, "w:t")) segs.push({ text: textOf(t) });
    } else if (c.name === "w:hyperlink") {
      for (const t of deepAll(c, "w:t")) segs.push({ text: textOf(t) });
    } else if (c.name === "w:fldSimple") {
      const instr = (attr(c, "w:instr") ?? "").toUpperCase();
      if (instr.includes("NUMPAGES")) segs.push({ field: "Y" });
      else if (instr.includes("PAGE")) segs.push({ field: "X" });
    }
  }
  return segs;
}

/** 解析页眉/页脚 XML：文字行 / 对齐 / 字体字号 / 页码字段段落 */
async function parseHeaderFooter(zip, sectPr, rels, kind, notes) {
  if (!sectPr) return null;
  const refs = kids(sectPr, `w:${kind}Reference`);
  const ref = refs.find((r) => attr(r, "w:type") === "default") ?? refs[0];
  if (!ref) return null;
  const target = rels.get(attr(ref, "r:id"));
  if (!target) return null;
  const xml = await zip.file(`word/${target.replace(/^\//, "")}`)?.async("string");
  if (!xml) return null;
  const root = rootEl(parseXml(xml));
  const label = kind === "header" ? "页眉" : "页脚";

  if (deepAll(root, "w:drawing").length || deepAll(root, "w:pict").length) {
    notes.push(`${label}含图片/图形，未提取`);
  }
  if (deepAll(root, "w:tbl").length) {
    notes.push(`${label}含表格布局，仅提取文字部分`);
  }

  const lines = [];
  let align = null;
  let font = null;
  let sizePt = null;
  let pn = null;
  for (const p of deepAll(root, "w:p")) {
    const segs = segmentsOf(p);
    if (segs.some((s) => s.field)) {
      const jc = attr(kid(kid(p, "w:pPr"), "w:jc"), "w:val");
      pn = { template: segs.map((s) => s.field ?? s.text).join(""), align: jcToAlign(jc) };
      font ??= firstFontIn(p);
      sizePt ??= firstSizeIn(p);
      continue;
    }
    const text = segs.map((s) => s.text).join("");
    if (!text.trim()) continue;
    lines.push(text);
    const jc = attr(kid(kid(p, "w:pPr"), "w:jc"), "w:val");
    align ??= jcToAlign(jc);
    font ??= firstFontIn(p);
    sizePt ??= firstSizeIn(p);
  }
  return {
    text: lines.length ? lines.join("\n") : null,
    align,
    font,
    size: sizePt,
    pn,
  };
}

/** 段内第一个含 rFonts 的 rPr → 字体 */
function firstFontIn(p) {
  const rPr = deepAll(p, "w:rPr").find((r) => kid(r, "w:rFonts"));
  return rPr ? fontFromRPr(rPr) : null;
}
/** 段内第一个含 sz 的 rPr → 字号 */
function firstSizeIn(p) {
  const rPr = deepAll(p, "w:rPr").find((r) => kid(r, "w:sz"));
  return rPr ? halfPtFromRPr(rPr) : null;
}
/** jc → MDTT 对齐值；both/distribute/缺省视为默认（左） */
function jcToAlign(jc) {
  return jc === "center" || jc === "right" || jc === "left" ? jc : null;
}

// ================= 保存与摘要 =================

/** 提取并保存自定义预设；返回 { file, options, notes } */
export async function extractAndSavePreset(docxPath, name, { overwrite = false } = {}) {
  if (!validPresetName(name)) {
    throw new Error(`预设名「${name}」不合法（仅限中英文、数字、下划线、连字符，且不以连字符开头）`);
  }
  const file = customPresetPath(name);
  if (fs.existsSync(file) && !overwrite) {
    throw new Error(`预设已存在「${file}」，如需覆盖请加 --overwrite`);
  }
  const { options, notes } = await extractPresetOptions(docxPath);
  const data = {
    name,
    source: path.resolve(docxPath),
    extractedAt: new Date().toISOString(),
    notes,
    options,
  };
  fs.mkdirSync(presetsDir(), { recursive: true });
  fs.writeFileSync(file, JSON.stringify(data, null, 2) + "\n", "utf-8");
  return { file, options, notes };
}

/** 提取结果的人读摘要（逐行） */
export function presetSummaryLines(options) {
  const L = [];
  const cn = { left: "左", center: "居中", right: "右", justify: "两端" };
  if (options.page) {
    const p = options.page;
    const m = p.margin
      ? `，边距 上${p.margin.top}/右${p.margin.right}/下${p.margin.bottom}/左${p.margin.left}cm`
      : "";
    const v = p.vAlign ? `，内容${cn[p.vAlign] ?? p.vAlign}对齐` : "";
    L.push(`页面    ${p.size} ${p.orientation === "landscape" ? "横向" : "纵向"}${m}${v}`);
  }
  if (options.fonts?.body || options.sizes?.body || options.paragraph) {
    const parts = [];
    if (options.fonts?.body) parts.push(`${options.fonts.body.eastAsia} / ${options.fonts.body.ascii}`);
    if (options.sizes?.body) parts.push(`${options.sizes.body}pt`);
    const pa = options.paragraph ?? {};
    if (pa.firstLineChars != null) parts.push(`首行缩进${pa.firstLineChars}字符`);
    if (pa.line != null) parts.push(`行距${pa.line}`);
    if (pa.afterLines != null) parts.push(`段后${pa.afterLines}行`);
    else if (pa.afterPt != null) parts.push(`段后${pa.afterPt}pt`);
    if (pa.align) parts.push(`${cn[pa.align] ?? pa.align}对齐`);
    if (parts.length) L.push(`正文    ${parts.join("，")}`);
  }
  if (options.fonts?.heading || options.sizes?.heading || options.heading) {
    const parts = [];
    if (options.fonts?.heading) parts.push(`${options.fonts.heading.eastAsia} / ${options.fonts.heading.ascii}`);
    if (options.sizes?.heading) parts.push(`H1-H6 字号 ${options.sizes.heading.map((s) => s ?? "?").join("/")}pt`);
    if (options.heading?.align) parts.push(`对齐 ${options.heading.align.map((a) => cn[a] ?? "左").join("/")}`);
    if (options.heading?.bold) parts.push(`加粗 ${options.heading.bold.map((b) => (b ? "√" : "×")).join("")}`);
    if (parts.length) L.push(`标题    ${parts.join("，")}`);
  }
  if (options.header?.text) {
    const extra = [
      `${cn[options.header.align] ?? "居中"}对齐`,
      options.fonts?.header ? options.fonts.header.eastAsia : null,
      options.sizes?.header ? `${options.sizes.header}pt` : null,
    ].filter(Boolean);
    L.push(`页眉    ${options.header.text.split("\n").length}行文字，${extra.join(" ")}`);
  }
  if (options.footer?.text) {
    L.push(`页脚    ${options.footer.text.split("\n").length}行文字，${cn[options.footer.align] ?? "居中"}对齐`);
  }
  if (options.pageNumber) {
    const pn = options.pageNumber;
    L.push(
      `页码    ${pn.pos === "bottom" ? "页脚" : "页眉"}${pn.align ? `${cn[pn.align] ?? ""}对齐` : ""}，格式「${pn.format}」${pn.start != null ? `，起始页码 ${pn.start}` : ""}`
    );
  }
  return L;
}
