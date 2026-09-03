// 页眉页脚：文字/左右分布/图片布局/页码/首页差异化/渐变色带

import path from "node:path";
import {
  AlignmentType,
  BorderStyle,
  ExternalHyperlink,
  Footer,
  Header,
  ImageRun,
  LineRuleType,
  PageNumber,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  VerticalAlignTable,
  WidthType,
  convertToXmlComponent,
} from "docx";
import { xml2js } from "xml-js";
import { cmToTwip, ptToHalfPoint } from "./options.js";
import { ALIGN_MAP, fontObj } from "./styles.js";
import { contentWidthTwip } from "./page.js";
import { loadImage } from "./inline.js";

/** 构建页眉页脚（含首页差异化与页码） */
export function buildHeaderFooter(opts, basePath) {
  const pn = opts.pageNumber;
  const defaultHeaderParas = headerFooterParas(opts.header, opts, "header", basePath);
  const defaultFooterParas = headerFooterParas(opts.footer, opts, "footer", basePath);

  if (pn.pos === "top") defaultHeaderParas.push(...pageNumberParas(pn, opts));
  if (pn.pos === "bottom") defaultFooterParas.push(...pageNumberParas(pn, opts));

  // 首页差异化：--first-*、--no-first-*、--no-first-page-number 任一生效时启用 titlePage
  let firstHeaderParas = firstVariantParas(opts.header, defaultHeaderParas, opts, "header");
  let firstFooterParas = firstVariantParas(opts.footer, defaultFooterParas, opts, "footer");
  if (!pn.showOnFirstPage) {
    // 首页不显示页码：复制一份并剔除页码段落（避免改动共享的默认页眉/页脚）
    const numCount = pageNumberParas(pn, opts).length;
    if (pn.pos === "bottom") firstFooterParas = firstFooterParas.slice(0, firstFooterParas.length - numCount);
    if (pn.pos === "top") firstHeaderParas = firstHeaderParas.slice(0, firstHeaderParas.length - numCount);
  }
  const titlePage =
    firstHeaderParas !== defaultHeaderParas ||
    firstFooterParas !== defaultFooterParas ||
    !pn.showOnFirstPage;

  const headers = {};
  const footers = {};
  if (defaultHeaderParas.length) headers.default = new Header({ children: defaultHeaderParas });
  if (defaultFooterParas.length) footers.default = new Footer({ children: defaultFooterParas });
  if (titlePage) {
    if (firstHeaderParas.length) headers.first = new Header({ children: firstHeaderParas });
    if (firstFooterParas.length) footers.first = new Footer({ children: firstFooterParas });
  }
  return { headers, footers, titlePage };
}

/** 页眉（或页脚）的默认内容列表（Paragraph 或 Table） */
function headerFooterParas(cfg, opts, which, basePath) {
  const paras = [];
  const font = which === "header" ? opts.fonts.header : opts.fonts.footer;
  const size = which === "header" ? opts.sizes.header : opts.sizes.footer;
  const align = ALIGN_MAP[cfg.align] || AlignmentType.CENTER;
  const { text, left, right, image, gradientLine } = cfg;
  const widthTwip = contentWidthTwip(opts);
  if (image) {
    // 左文右图布局：无边框表格，左列文字右列图片；渐变色带为独立原生渐变图形
    const lines = text != null ? splitLines(text) : [];
    paras.push(...buildHeaderImageBlock(lines, align, font, size, image, basePath, gradientLine, widthTwip));
    return paras;
  }
  if (left != null || right != null) {
    paras.push(new Paragraph({
      tabStops: [{ type: "right", position: contentWidthTwip(opts) }],
      children: [
        ...textRunsWithLinks(left || "", font, size),
        new TextRun({ text: "\t" }),
        ...textRunsWithLinks(right || "", font, size),
      ],
    }));
  } else if (text != null) {
    const tight = { after: 0, line: 240, lineRule: LineRuleType.AUTO };
    for (const line of splitLines(text)) {
      paras.push(new Paragraph({ alignment: align, spacing: tight, children: textRunsWithLinks(line, font, size) }));
    }
  }
  if (gradientLine && paras.length) {
    paras.push(buildGradientLineShape(gradientLine, widthTwip));
  }
  return paras;
}

/** 多行文本拆分：支持真实换行符与字面 \n */
function splitLines(text) {
  return String(text).split(/\r?\n|\\n/);
}

/** 首页变体段落：无差异时返回原数组引用 */
function firstVariantParas(cfg, defaultParas, opts, which) {
  if (cfg.noFirst) return [];
  if (cfg.firstText == null) return defaultParas;
  const font = which === "header" ? opts.fonts.header : opts.fonts.footer;
  const size = which === "header" ? opts.sizes.header : opts.sizes.footer;
  return [new Paragraph({
    alignment: ALIGN_MAP[cfg.align] || AlignmentType.CENTER,
    children: textRunsWithLinks(cfg.firstText, font, size),
  })];
}

/** 页码段落：模板中 X=当前页码字段，Y=总页数字段 */
export function pageNumberParas(pn, opts) {
  if (pn.pos === "none") return [];
  const font = pn.pos === "top" ? opts.fonts.header : opts.fonts.footer;
  const size = pn.pos === "top" ? opts.sizes.header : opts.sizes.footer;
  const template = NUM_TEMPLATE(pn);
  const parts = template.split(/([XY])/);
  const children = [];
  for (const part of parts) {
    if (part === "X") children.push(new TextRun({ children: [PageNumber.CURRENT], font: fontObj(font), size: ptToHalfPoint(size) }));
    else if (part === "Y") children.push(new TextRun({ children: [PageNumber.TOTAL_PAGES], font: fontObj(font), size: ptToHalfPoint(size) }));
    else if (part) children.push(new TextRun({ text: part, font: fontObj(font), size: ptToHalfPoint(size) }));
  }
  return [new Paragraph({ alignment: ALIGN_MAP[pn.align] || AlignmentType.CENTER, children })];
}

/** 页码模板：数字样式特殊值（1/i/I/一）按纯页码字段处理 */
function NUM_TEMPLATE(pn) {
  return ["1", "i", "I", "一"].includes(pn.format) ? "X" : pn.format;
}

/** 文本中的 URL 自动转为超链接 */
function textRunsWithLinks(text, font, size) {
  const f = fontObj(font);
  const half = ptToHalfPoint(size);
  const parts = String(text).split(/(https?:\/\/[^\s，。"'\）)]+)/);
  const runs = [];
  for (const part of parts) {
    if (!part) continue;
    if (/^https?:\/\//.test(part)) {
      runs.push(new ExternalHyperlink({
        link: part,
        children: [new TextRun({ text: part, font: f, size: half, style: "Hyperlink" })],
      }));
    } else {
      runs.push(new TextRun({ text: part, font: f, size: half }));
    }
  }
  return runs;
}

/** 页眉左文右图：无边框两列表格，右列图片垂直居中；渐变色带为独立原生渐变图形 */
function buildHeaderImageBlock(lines, align, font, size, image, basePath, gradientLine, widthTwip) {
  const tight = { after: 0, line: 240, lineRule: LineRuleType.AUTO };
  const textParas = lines.map((line) =>
    new Paragraph({ alignment: align, spacing: tight, children: textRunsWithLinks(line, font, size) })
  );
  let imageRun;
  try {
    imageRun = makeHeaderImageRun(image, basePath);
  } catch (e) {
    console.warn(`警告: 页眉图片加载失败，已跳过（${e.message}）`);
    if (!textParas.length) return [];
    return gradientLine
      ? [wrapNoBorderTable(textParas.length ? textParas : []), buildGradientLineShape(gradientLine, widthTwip)]
      : textParas;
  }
  const none = { style: BorderStyle.NONE, size: 0, color: "auto" };
  const zeroMargins = { marginUnitType: WidthType.DXA, top: 0, bottom: 0, left: 0, right: 0 };
  const blocks = [
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: { top: none, bottom: none, left: none, right: none, insideHorizontal: none, insideVertical: none },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: textParas.length ? textParas : [new Paragraph({ children: [] })],
              width: { size: 84, type: WidthType.PERCENTAGE },
              margins: zeroMargins,
            }),
            new TableCell({
              children: [new Paragraph({ alignment: AlignmentType.RIGHT, spacing: tight, children: [imageRun] })],
              width: { size: 16, type: WidthType.PERCENTAGE },
              verticalAlign: VerticalAlignTable.CENTER,
              margins: zeroMargins,
            }),
          ],
        }),
      ],
    }),
  ];
  // 渐变条用 Word 原生 DrawingML 渐变实现（真渐变，非拼格）
  if (gradientLine) blocks.push(buildGradientLineShape(gradientLine, widthTwip));
  return blocks;
}

/** 无边框单列表格（用于纯文字页眉 + 渐变条的组合，保持宽度一致） */
function wrapNoBorderTable(paras) {
  const none = { style: BorderStyle.NONE, size: 0, color: "auto" };
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: { top: none, bottom: none, left: none, right: none, insideHorizontal: none, insideVertical: none },
    rows: [new TableRow({
      children: [new TableCell({
        children: paras,
        width: { size: 100, type: WidthType.PERCENTAGE },
        margins: { marginUnitType: WidthType.DXA, top: 0, bottom: 0, left: 0, right: 0 },
      })],
    })],
  });
}

/** 页眉图片：按 heightCm 等比缩放 */
function makeHeaderImageRun(image, basePath) {
  const absPath = path.isAbsolute(image.path) ? image.path : path.resolve(basePath, image.path);
  const { data, dim, type } = loadImage(absPath);
  const heightCm = image.heightCm || 1.2;
  const heightPx = Math.round((heightCm / 2.54) * 96);
  const scale = heightPx / dim.height;
  return new ImageRun({
    type,
    data,
    transformation: { width: Math.round(dim.width * scale), height: heightPx },
  });
}

/**
 * 渐变色带：用 Word 原生 DrawingML 线性渐变（a:gradFill）实现真实连续渐变。
 * 相比"多个单元格拼色"，单个渐变图形在 Word 中渲染稳定、过渡平滑，且停靠点可任意控制。
 * widthTwip 为内容区宽度，用于让色带铺满页眉；返回一段可直接放入页眉 children 的组件。
 */
function buildGradientLineShape(gradientLine, widthTwip) {
  const { stops, heightCm = 0.06, gapCm = 0.2 } = gradientLine;
  const cx = Math.round(widthTwip * 635);            // 1 twip = 635 EMU
  const cy = Math.round(cmToTwip(heightCm) * 635);
  const beforeTwip = Math.round(cmToTwip(gapCm));    // 色带段前距，拉开与上方文字间距
  const gsXml = stops
    .map(([pos, color]) => `<a:gs pos="${Math.round(pos * 100000)}"><a:srgbClr val="${color}"/></a:gs>`)
    .join("");
  const xml =
    `<w:p><w:pPr><w:spacing w:before="${beforeTwip}" w:after="0" w:line="14" w:lineRule="exact"/></w:pPr>` +
    `<w:r><w:drawing>` +
    `<wp:inline distT="0" distB="0" distL="0" distR="0">` +
    `<wp:extent cx="${cx}" cy="${cy}"/>` +
    `<wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    `<wp:docPr id="1000" name="GradientLine"/>` +
    `<wp:cNvGraphicFramePr/>` +
    `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
    `<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr>` +
    `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
    `<a:gradFill rotWithShape="1"><a:gsLst>${gsXml}</a:gsLst><a:lin ang="0" scaled="1"/></a:gradFill>` +
    `<a:ln><a:noFill/></a:ln>` +
    `</wps:spPr><wps:bodyPr/></wps:wsp>` +
    `</a:graphicData></a:graphic>` +
    `</wp:inline></w:drawing></w:r></w:p>`;
  // 注意：docx 的 ImportedXmlComponent.fromXmlString 会把根元素名丢失（产出 <undefined> 标签），
  // 因此改用 convertToXmlComponent 直接解析单一根元素，确保结构正确。
  return convertToXmlComponent(xml2js(xml, { compact: false }).elements[0]);
}
