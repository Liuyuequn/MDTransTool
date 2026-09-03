import fs from "node:fs";
import path from "node:path";
import MarkdownIt from "markdown-it";
import { imageSize } from "image-size";
import {
  AlignmentType,
  BorderStyle,
  convertToXmlComponent,
  Document,
  ExternalHyperlink,
  Footer,
  Header,
  HeadingLevel,
  ImageRun,
  LevelFormat,
  LineRuleType,
  NumberFormat,
  Packer,
  PageNumber,
  PageOrientation,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  UnderlineType,
  VerticalAlignSection,
  VerticalAlignTable,
  WidthType,
} from "docx";
import { xml2js } from "xml-js";
import { PAGE_SIZES, cmToTwip, ptToHalfPoint, ptToTwip } from "./options.js";

const HEADINGS = [
  HeadingLevel.HEADING_1,
  HeadingLevel.HEADING_2,
  HeadingLevel.HEADING_3,
  HeadingLevel.HEADING_4,
  HeadingLevel.HEADING_5,
  HeadingLevel.HEADING_6,
];

const BULLET_CHARS = ["\u2022", "\u25E6", "\u2022", "\u25E6", "\u2022", "\u25E6"];
const IMAGE_TYPES = { ".png": "png", ".jpg": "jpg", ".jpeg": "jpg", ".gif": "gif", ".bmp": "bmp" };
const MAX_IMAGE_WIDTH = 600; // 像素，超出按比例缩小

const ALIGN_MAP = {
  left: AlignmentType.LEFT,
  center: AlignmentType.CENTER,
  right: AlignmentType.RIGHT,
  justify: AlignmentType.JUSTIFIED,
};

const NUM_FORMAT_MAP = {
  "1": NumberFormat.DECIMAL,
  i: NumberFormat.LOWER_ROMAN,
  I: NumberFormat.UPPER_ROMAN,
  一: NumberFormat.CHINESE_COUNTING,
};

const md = new MarkdownIt({ html: false, linkify: true, typographer: true });

/** 将 markdown 文件转换为 docx 文件（opts 为合并完成的完整配置） */
export async function convertMarkdownFile(inputPath, outputPath, opts) {
  const mdText = fs.readFileSync(inputPath, "utf-8");
  const doc = buildDocument(mdText, opts, path.dirname(path.resolve(inputPath)));
  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(outputPath, buffer);
}

function buildDocument(mdText, opts, basePath) {
  const tokens = md.parse(mdText, {});
  const ctx = makeCtx(opts, basePath);
  const state = { orderedInstance: 0 };
  const blocks = parseBlocks(tokens, 0, tokens.length, ctx, state);

  const headerFooter = buildHeaderFooter(opts, basePath);

  return new Document({
    styles: buildStyles(opts),
    numbering: { config: [buildBulletNumbering(), buildOrderedNumbering()] },
    sections: [
      {
        properties: {
          page: {
            size: buildPageSize(opts),
            margin: {
              top: cmToTwip(opts.page.margin.top),
              right: cmToTwip(opts.page.margin.right),
              bottom: cmToTwip(opts.page.margin.bottom),
              left: cmToTwip(opts.page.margin.left),
            },
            pageNumbers: buildPageNumbers(opts),
          },
          verticalAlign: ALIGN_MAP[opts.page.vAlign] ? toSectionVAlign(opts.page.vAlign) : undefined,
          titlePage: headerFooter.titlePage || undefined,
        },
        headers: headerFooter.headers,
        footers: headerFooter.footers,
        children: blocks.length ? blocks : [new Paragraph({ children: [] })],
      },
    ],
  });
}

function toSectionVAlign(v) {
  return v === "center" ? VerticalAlignSection.CENTER : v === "bottom" ? VerticalAlignSection.BOTTOM : VerticalAlignSection.TOP;
}

/** 页面尺寸：预设名或 "W,H"（cm），横向时自动交换宽高 */
function buildPageSize(opts) {
  let width;
  let height;
  if (PAGE_SIZES[opts.page.size]) {
    ({ width, height } = PAGE_SIZES[opts.page.size]);
  } else {
    const [w, h] = opts.page.size.split(",").map(Number);
    if (!w || !h || Number.isNaN(w) || Number.isNaN(h)) {
      throw new Error(`无法识别的页面尺寸「${opts.page.size}」`);
    }
    width = cmToTwip(w);
    height = cmToTwip(h);
  }
  const landscape = opts.page.orientation === "landscape";
  return {
    width: landscape ? height : width,
    height: landscape ? width : height,
    orientation: landscape ? PageOrientation.LANDSCAPE : PageOrientation.PORTRAIT,
  };
}

/** 起始页码与页码数字样式 */
function buildPageNumbers(opts) {
  const pn = opts.pageNumber;
  const result = {};
  if (pn.start != null) result.start = pn.start;
  const fmt = NUM_FORMAT_MAP[pn.format];
  if (fmt) result.formatType = fmt;
  return Object.keys(result).length ? result : undefined;
}

// ============ 样式 ============

function makeCtx(opts, basePath) {
  return {
    basePath,
    opts,
    listLevel: -1,
    listType: null,
    instance: 0,
    quote: false,
    forceBold: false,
  };
}

function buildStyles(opts) {
  const headingDefaults = (level) => {
    const i = level - 1;
    return {
      run: {
        font: headingFont(opts),
        size: ptToHalfPoint(opts.sizes.heading[i]),
        bold: opts.heading.bold[i] || undefined,
        color: "000000",
      },
      paragraph: {
        alignment: opts.heading.align[i] ? ALIGN_MAP[opts.heading.align[i]] : AlignmentType.LEFT,
        spacing: resolveSpacing(opts.heading.spacing, opts.sizes.heading[i], opts.paragraph.line),
      },
    };
  };
  return {
    default: {
      document: {
        run: { font: fontObj(opts.fonts.body), size: ptToHalfPoint(opts.sizes.body) },
        paragraph: {
          spacing: resolveSpacing(
            { beforePt: 0, afterPt: opts.paragraph.afterPt, afterLines: opts.paragraph.afterLines },
            opts.sizes.body,
            opts.paragraph.line
          ),
        },
      },
      heading1: headingDefaults(1),
      heading2: headingDefaults(2),
      heading3: headingDefaults(3),
      heading4: headingDefaults(4),
      heading5: headingDefaults(5),
      heading6: headingDefaults(6),
    },
    characterStyles: [
      {
        id: "Hyperlink",
        name: "Hyperlink",
        run: { color: "0563C1", underline: { type: UnderlineType.SINGLE } },
      },
    ],
  };
}

function headingFont(opts) {
  return fontObj(opts.fonts.heading || opts.fonts.body);
}

function fontObj(f) {
  return { ascii: f.ascii, hAnsi: f.ascii, eastAsia: f.eastAsia };
}

/**
 * 解析间距：lines 字段优先（按 字号×行距倍数 换算 twip），否则用 pt；
 * line 为行距倍数（w:line = 倍数×240，lineRule=auto）
 */
function resolveSpacing(spacing, sizePt, lineMultiplier) {
  const out = {};
  if (spacing.beforePt != null) out.before = ptToTwip(spacing.beforePt);
  if (spacing.beforeLines != null) out.before = linesToTwip(spacing.beforeLines, sizePt, lineMultiplier);
  if (spacing.afterPt != null) out.after = ptToTwip(spacing.afterPt);
  if (spacing.afterLines != null) out.after = linesToTwip(spacing.afterLines, sizePt, lineMultiplier);
  if (lineMultiplier) {
    out.line = Math.round(lineMultiplier * 240);
    out.lineRule = LineRuleType.AUTO;
  }
  return out;
}

/** “行”单位 → twip：1 行 ≈ 字号 × 行距倍数 */
function linesToTwip(lines, sizePt, lineMultiplier = 1.28) {
  return Math.round(lines * sizePt * lineMultiplier * 20);
}

function buildBulletNumbering() {
  return {
    reference: "md-bullet",
    levels: Array.from({ length: 6 }, (_, l) => ({
      level: l,
      format: LevelFormat.BULLET,
      text: BULLET_CHARS[l],
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 720 * (l + 1), hanging: 360 } } },
    })),
  };
}

function buildOrderedNumbering() {
  return {
    reference: "md-ordered",
    levels: Array.from({ length: 6 }, (_, l) => ({
      level: l,
      format: LevelFormat.DECIMAL,
      text: `%${l + 1}.`,
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 720 * (l + 1), hanging: 360 } } },
    })),
  };
}

// ============ 页眉页脚 ============

/** 构建页眉页脚（含首页差异化与页码） */
function buildHeaderFooter(opts, basePath) {
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

/** 页面内容区宽度（twip）：用于左右分布页眉的右对齐制表位 */
function contentWidthTwip(opts) {
  const size = buildPageSize(opts);
  return size.width - cmToTwip(opts.page.margin.left) - cmToTwip(opts.page.margin.right);
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

/**
 * 渐变色带：用 Word 原生 DrawingML 线性渐变（a:gradFill）实现真实连续渐变。
 * 相比“多个单元格拼色”，单个渐变图形在 Word 中渲染稳定、过渡平滑，且停靠点可任意控制。
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

/** 页眉图片：按 heightCm 等比缩放 */
function makeHeaderImageRun(image, basePath) {
  const absPath = path.isAbsolute(image.path) ? image.path : path.resolve(basePath, image.path);
  if (!fs.existsSync(absPath)) throw new Error(`找不到页眉图片: ${absPath}`);
  const data = fs.readFileSync(absPath);
  const dim = imageSize(data);
  const type = IMAGE_TYPES[path.extname(absPath).toLowerCase()];
  if (!type) throw new Error(`不支持的图片格式: ${path.extname(absPath)}`);
  const heightCm = image.heightCm || 1.2;
  const heightPx = Math.round((heightCm / 2.54) * 96);
  const scale = heightPx / dim.height;
  return new ImageRun({
    type,
    data,
    transformation: { width: Math.round(dim.width * scale), height: heightPx },
  });
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

/** 页码段落：模板中 X=当前页码字段，Y=总页数字段 */
function pageNumberParas(pn, opts) {
  if (pn.pos === "none") return [];
  const font = pn.pos === "top" ? opts.fonts.header : opts.fonts.footer;
  const size = pn.pos === "top" ? opts.sizes.header : opts.sizes.footer;
  const template = NUM_FORMAT_MAP[pn.format] ? "X" : pn.format;
  const parts = template.split(/([XY])/);
  const children = [];
  for (const part of parts) {
    if (part === "X") children.push(new TextRun({ children: [PageNumber.CURRENT], font: fontObj(font), size: ptToHalfPoint(size) }));
    else if (part === "Y") children.push(new TextRun({ children: [PageNumber.TOTAL_PAGES], font: fontObj(font), size: ptToHalfPoint(size) }));
    else if (part) children.push(new TextRun({ text: part, font: fontObj(font), size: ptToHalfPoint(size) }));
  }
  return [new Paragraph({ alignment: ALIGN_MAP[pn.align] || AlignmentType.CENTER, children })];
}

// ============ 块级元素解析 ============

/** 在 [start, end) 范围内查找与 openType 配对的 close token 下标 */
function findClose(tokens, start, end, openType, closeType) {
  let depth = 0;
  for (let j = start; j < end; j++) {
    if (tokens[j].type === openType) depth++;
    else if (tokens[j].type === closeType && --depth === 0) return j;
  }
  return -1;
}

function getAttr(token, name) {
  const attr = (token.attrs || []).find(([k]) => k === name);
  return attr ? attr[1] : null;
}

function parseBlocks(tokens, start, end, ctx, state) {
  const out = [];
  let i = start;
  while (i < end) {
    const t = tokens[i];
    switch (t.type) {
      case "heading_open": {
        const level = Number(t.tag.slice(1));
        const hCtx = { ...ctx, uppercase: ctx.opts.styles.headingUppercase };
        out.push(
          new Paragraph({
            heading: HEADINGS[level - 1],
            children: parseInline(tokens[i + 1].children, hCtx),
          })
        );
        i += 3;
        break;
      }
      case "paragraph_open": {
        const inlineTok = tokens[i + 1];
        if (isImageOnly(inlineTok.children)) {
          out.push(
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { before: 120, after: 120 },
              children: parseInline(inlineTok.children, ctx),
            })
          );
        } else {
          out.push(makeParagraph(parseInline(inlineTok.children, ctx), ctx));
        }
        i += 3;
        break;
      }
      case "bullet_list_open":
      case "ordered_list_open": {
        const closeType = t.type.replace("_open", "_close");
        const close = findClose(tokens, i, end, t.type, closeType);
        if (close === -1) { i++; break; }
        const isOrdered = t.type === "ordered_list_open";
        const itemCtx = {
          ...ctx,
          listLevel: ctx.listLevel + 1,
          listType: isOrdered ? "md-ordered" : "md-bullet",
          instance: isOrdered ? ++state.orderedInstance : 0,
        };
        out.push(...parseBlocks(tokens, i + 1, close, itemCtx, state));
        i = close + 1;
        break;
      }
      case "list_item_open":
      case "list_item_close":
      case "space":
        i++;
        break;
      case "fence":
      case "code_block": {
        out.push(...makeCodeBlock(t.content, ctx));
        i++;
        break;
      }
      case "blockquote_open": {
        const close = findClose(tokens, i, end, "blockquote_open", "blockquote_close");
        if (close === -1) { i++; break; }
        out.push(...parseBlocks(tokens, i + 1, close, { ...ctx, quote: true }, state));
        i = close + 1;
        break;
      }
      case "hr": {
        out.push(
          new Paragraph({
            border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "AAAAAA", space: 1 } },
            spacing: { before: 240, after: 240 },
          })
        );
        i++;
        break;
      }
      case "table_open": {
        const close = findClose(tokens, i, end, "table_open", "table_close");
        if (close === -1) { i++; break; }
        out.push(parseTable(tokens, i + 1, close, ctx));
        out.push(new Paragraph({ spacing: { after: 120 }, children: [] }));
        i = close + 1;
        break;
      }
      default:
        i++;
        break;
    }
  }
  return out;
}

/** 普通段落构造：应用正文对齐、首行缩进（仅普通段落）、列表编号、引用样式 */
function makeParagraph(runs, ctx, extra = {}) {
  const { opts } = ctx;
  const props = { children: runs };
  const isNormal = ctx.listLevel < 0 && !ctx.quote;
  if (isNormal && opts.paragraph.align) props.alignment = ALIGN_MAP[opts.paragraph.align];
  if (isNormal && opts.paragraph.firstLineChars > 0) {
    props.indent = { firstLineChars: Math.round(opts.paragraph.firstLineChars * 100) };
  }
  if (ctx.listLevel >= 0) {
    props.numbering = {
      reference: ctx.listType,
      level: Math.min(ctx.listLevel, 5),
      instance: ctx.instance,
    };
  }
  if (ctx.quote) {
    props.indent = { left: 480 };
    props.border = { left: { style: BorderStyle.SINGLE, size: 18, color: "CCCCCC", space: 10 } };
  }
  return new Paragraph({ ...props, ...extra });
}

/** 判断段落是否只包含图片（图片独立成段并居中，不参与首行缩进） */
function isImageOnly(children) {
  const meaningful = (children || []).filter((tok) => tok.type !== "text" || tok.content.trim() !== "");
  return meaningful.length === 1 && meaningful[0].type === "image";
}

function makeCodeBlock(content, ctx) {
  const { opts } = ctx;
  const lines = content.replace(/\n$/, "").split("\n");
  return lines.map((line, idx) =>
    new Paragraph({
      children: [
        new TextRun({
          text: line.length ? line : " ",
          font: fontObj(opts.fonts.code),
          size: ptToHalfPoint(opts.sizes.code),
        }),
      ],
      shading: { type: ShadingType.CLEAR, fill: "F2F2F2" },
      spacing: { after: idx === lines.length - 1 ? 200 : 0 },
      indent: { left: 120, right: 120 },
    })
  );
}

function parseTable(tokens, start, end, ctx) {
  const rows = [];
  let isHeader = true;
  let colCount = 0;
  let j = start;
  while (j < end) {
    if (tokens[j].type !== "tr_open") { j++; continue; }
    const trClose = findClose(tokens, j, end, "tr_open", "tr_close");
    if (colCount === 0) {
      for (let k = j + 1; k < trClose; k++) {
        if (tokens[k].type === "th_open" || tokens[k].type === "td_open") colCount++;
      }
    }
    const cells = [];
    let k = j + 1;
    while (k < trClose) {
      const type = tokens[k].type;
      if (type === "th_open" || type === "td_open") {
        const closeType = type.replace("_open", "_close");
        const cellEnd = findClose(tokens, k, trClose, type, closeType);
        const inlineTok = tokens.slice(k + 1, cellEnd).find((x) => x.type === "inline");
        const cellCtx = isHeader ? { ...ctx, forceBold: true } : ctx;
        const runs = parseInline(inlineTok ? inlineTok.children : [], cellCtx);
        const style = getAttr(tokens[k], "style");
        // 水平对齐：表头居中；数据行按 markdown 语法标注（无标注则默认左对齐）
        let alignment;
        if (isHeader) {
          alignment = AlignmentType.CENTER;
        } else if (style) {
          if (/text-align:\s*center/.test(style)) alignment = AlignmentType.CENTER;
          else if (/text-align:\s*(right|end)/.test(style)) alignment = AlignmentType.RIGHT;
        }
        cells.push(
          new TableCell({
            children: [new Paragraph({ children: runs, alignment })],
            verticalAlign: VerticalAlignTable.CENTER,
            width: colCount ? { size: Math.floor(1000 / colCount) / 10, type: WidthType.PERCENTAGE } : undefined,
            shading: isHeader ? { type: ShadingType.CLEAR, fill: "EFEFEF" } : undefined,
          })
        );
        k = cellEnd + 1;
      } else {
        k++;
      }
    }
    rows.push(new TableRow({ children: cells, tableHeader: isHeader }));
    if (isHeader) isHeader = false;
    j = trClose + 1;
  }
  return new Table({ rows, width: { size: 100, type: WidthType.PERCENTAGE } });
}

// ============ 行内元素解析 ============

function parseInline(children, ctx) {
  const runs = [];
  if (!children) return runs;
  let bold = false;
  let italic = false;
  let strike = false;
  let link = null;
  for (const t of children) {
    switch (t.type) {
      case "text":
      case "text_special":
        runs.push(makeRun(ctx.uppercase ? t.content.toUpperCase() : t.content, { bold, italic, strike, link, ctx }));
        break;
      case "strong_open":
        bold = true;
        break;
      case "strong_close":
        bold = false;
        break;
      case "em_open":
        italic = true;
        break;
      case "em_close":
        italic = false;
        break;
      case "s_open":
        strike = true;
        break;
      case "s_close":
        strike = false;
        break;
      case "code_inline":
        runs.push(
          new TextRun({
            text: t.content,
            font: fontObj(ctx.opts.fonts.code),
            size: ptToHalfPoint(ctx.opts.sizes.code),
            shading: { type: ShadingType.CLEAR, fill: "EFEFEF" },
          })
        );
        break;
      case "link_open":
        link = getAttr(t, "href");
        break;
      case "link_close":
        link = null;
        break;
      case "image":
        try {
          runs.push(makeImageRun(t, ctx));
        } catch {
          runs.push(makeRun(`[图片缺失: ${t.content || getAttr(t, "src") || "未知来源"}]`, { ctx }));
        }
        break;
      case "hardbreak":
        runs.push(new TextRun({ break: 1 }));
        break;
      case "softbreak":
        runs.push(new TextRun({ text: " " }));
        break;
      default:
        break;
    }
  }
  return runs;
}

function makeRun(text, { bold, italic, strike, link, ctx }) {
  const italicOn = italic && ctx.opts.styles.italic;
  const strikeOn = strike && ctx.opts.styles.strike;
  const runProps = {
    text,
    bold: bold || ctx.forceBold || undefined,
    italics: italicOn || undefined,
    strike: strikeOn || undefined,
  };
  if (link) {
    return new ExternalHyperlink({
      link,
      children: [new TextRun({ ...runProps, style: "Hyperlink" })],
    });
  }
  return new TextRun(runProps);
}

function makeImageRun(token, ctx) {
  const src = getAttr(token, "src");
  if (!src) throw new Error("图片缺少 src 属性");
  const absPath = path.resolve(ctx.basePath, decodeURIComponent(src.replace(/^file:\/\//, "")));
  if (!fs.existsSync(absPath)) throw new Error(`找不到图片文件: ${absPath}`);
  const data = fs.readFileSync(absPath);
  const dim = imageSize(data);
  const type = IMAGE_TYPES[path.extname(absPath).toLowerCase()];
  if (!type) throw new Error(`不支持的图片格式: ${path.extname(absPath)}`);
  const scale = Math.min(1, MAX_IMAGE_WIDTH / dim.width);
  return new ImageRun({
    type,
    data,
    transformation: {
      width: Math.round(dim.width * scale),
      height: Math.round(dim.height * scale),
    },
  });
}
