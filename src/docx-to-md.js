// docx → Markdown 转换模块
// 流程：mammoth 提取 HTML（图片落盘到 MDPictures/）→ DOM 预处理（表格重建/合并检测）
//       → turndown 转换（含合并单元格的表格以 HTML 形式回填）→ 后处理清洗

import crypto from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import mammoth from "mammoth";
import TurndownService from "turndown";
import { gfm } from "turndown-plugin-gfm";
import { parse as parseHtml } from "node-html-parser";

/** 图片输出目录名（位于输出 md 文件同目录下） */
export const PICTURE_DIR = "MDPictures";

/** 含合并单元格表格的回退说明注释 */
const MERGED_TABLE_NOTE = "<!-- 该表格含合并单元格，Markdown 表格无法表达，以 HTML 形式保留 -->";

/** contentType → 扩展名 */
function extFromContentType(contentType) {
  const map = {
    "image/png": "png",
    "image/jpeg": "jpg",
    "image/gif": "gif",
    "image/bmp": "bmp",
  };
  if (map[contentType]) return map[contentType];
  // 未知类型：取子类型部分并仅保留字母数字
  const sub = String(contentType).split("/")[1] || "img";
  return sub.replace(/[^a-z0-9]/gi, "") || "img";
}

/**
 * 将 docx 文件转换为 Markdown 字符串
 * 图片以内容 hash 命名写入 <outputDir>/MDPictures/，正文中以相对路径引用
 * @param {string} docxPath - docx 文件路径
 * @param {string} outputDir - 输出 md 所在目录（MDPictures 的父目录）
 * @returns {Promise<string>} Markdown 内容
 */
export async function convertDocxToMarkdown(docxPath, outputDir) {
  if (!fs.existsSync(docxPath)) {
    throw new Error(`找不到文件 ${docxPath}`);
  }

  const picDir = path.join(outputDir, PICTURE_DIR);

  // 1. mammoth: docx → HTML，保留语义结构
  const { value: rawHtml } = await mammoth.convertToHtml(
    { path: docxPath },
    {
      styleMap: [
        "p[style-name='Heading 1'] => h1:fresh",
        "p[style-name='Heading 2'] => h2:fresh",
        "p[style-name='Heading 3'] => h3:fresh",
        "p[style-name='Heading 4'] => h4:fresh",
        "p[style-name='Heading 5'] => h5:fresh",
        "p[style-name='Heading 6'] => h6:fresh",
        "p[style-name='标题 1'] => h1:fresh",
        "p[style-name='标题 2'] => h2:fresh",
        "p[style-name='标题 3'] => h3:fresh",
        "p[style-name='标题 4'] => h4:fresh",
        "p[style-name='标题 5'] => h5:fresh",
        "p[style-name='标题 6'] => h6:fresh",
        "p[style-name='Quote'] => blockquote",
      ],
      convertImage: mammoth.images.imgElement(async (image) => {
        const buffer = await image.read();
        const ext = extFromContentType(image.contentType);
        // 内容 hash 命名：重复图片只落盘一次，重复转换文件名稳定
        const hash = crypto.createHash("md5").update(buffer).digest("hex").slice(0, 10);
        const name = `img-${hash}.${ext}`;
        fs.mkdirSync(picDir, { recursive: true });
        fs.writeFileSync(path.join(picDir, name), buffer);
        return { src: `${PICTURE_DIR}/${name}` }; // md 中统一用正斜杠相对路径
      }),
    }
  );

  // 2. DOM 预处理：表格结构重建；含合并单元格的表格摘出待回填
  const { html, htmlTables } = preprocessHtml(rawHtml);

  // 3. turndown: HTML → Markdown
  const td = createTurndownService();
  let md = td.turndown(html);

  // 4. HTML 回退表格回填（含合并单元格的表格以 HTML 形式保留）
  //    替换回调形式避免 replacement 字符串中 $& 等特殊序列被解释
  htmlTables.forEach((tableHtml, i) => {
    md = md.replace(`MDTTHTMLTABLE${i}`, () => `${MERGED_TABLE_NOTE}\n\n${tableHtml}`);
  });

  // 5. 后处理
  md = postProcess(md);

  return md;
}

/**
 * DOM 预处理：修复 mammoth 输出的结构性问题
 * - 普通表格：重组为规范 thead/tbody 结构（首行 th、数据行 td、剥离单元格内 <p> 包裹），
 *   供 turndown 生成 GFM 管道表格
 * - 含合并单元格（colspan/rowspan）的表格：GFM 管道语法无法表达合并，
 *   清理结构后整体摘出（占位符替换），转换完成后以 HTML 形式回填，完整保留合并语义
 */
function preprocessHtml(rawHtml) {
  const root = parseHtml(rawHtml);
  const htmlTables = [];
  for (const table of root.querySelectorAll("table")) {
    const rows = directRows(table);
    if (rows.length === 0) continue;
    const hasMerge = rows.some((tr) =>
      directCells(tr).some((c) => MERGE_ATTR_RE.test(c.rawAttrs))
    );
    if (hasMerge) {
      htmlTables.push(buildCleanTable(rows, { keepAttrs: true }));
      table.insertAdjacentHTML("beforebegin", `<p>MDTTHTMLTABLE${htmlTables.length - 1}</p>`);
      table.remove();
    } else {
      table.insertAdjacentHTML("beforebegin", buildCleanTable(rows, { keepAttrs: false }));
      table.remove();
    }
  }
  // 列表项内部的段落：<li><p>text</p></li> → <li>text</li>
  let h = root.toString();
  h = h.replace(/<li>\s*<p>([\s\S]*?)<\/p>\s*<\/li>/g, "<li>$1</li>");
  return { html: h, htmlTables };
}

const MERGE_ATTR_RE = /\b(?:colspan|rowspan)\s*=/i;

/** 表格的直接行：兼容 table>tr 与 table>thead/tbody>tr 两种 mammoth 输出结构，不深入嵌套表格 */
function directRows(table) {
  const rows = [];
  for (const child of table.childNodes) {
    if (child.nodeType !== 1) continue;
    if (child.tagName === "TR") rows.push(child);
    else if (child.tagName === "THEAD" || child.tagName === "TBODY" || child.tagName === "TFOOT") {
      for (const gc of child.childNodes) {
        if (gc.nodeType === 1 && gc.tagName === "TR") rows.push(gc);
      }
    }
  }
  return rows;
}

/** 行的直接单元格（th/td），不深入嵌套内容 */
function directCells(tr) {
  return tr.childNodes.filter((n) => n.nodeType === 1 && (n.tagName === "TD" || n.tagName === "TH"));
}

/**
 * 重建干净的表格 HTML 字符串：统一 thead（首行）+ tbody（其余行），剥离单元格内 <p> 包裹
 * keepAttrs=true（HTML 回退）：保留 colspan/rowspan 等全部属性与原单元格标签
 * keepAttrs=false（GFM 路径）：丢弃属性，首行强制 th、数据行强制 td（mammoth 可能把所有行都标成 th）
 */
function buildCleanTable(rows, { keepAttrs }) {
  const trs = rows.map((tr, i) => {
    const isHeader = i === 0;
    const cells = directCells(tr).map((c) => {
      const tag = keepAttrs ? c.tagName.toLowerCase() : isHeader ? "th" : "td";
      const attrs = keepAttrs && c.rawAttrs ? ` ${c.rawAttrs}` : "";
      const content = c.innerHTML.replace(/<\/?p>/g, "").trim();
      return `<${tag}${attrs}>${content}</${tag}>`;
    });
    return `<tr>${cells.join("")}</tr>`;
  });
  const thead = `<thead>${trs[0]}</thead>`;
  const tbody = trs.length > 1 ? `<tbody>${trs.slice(1).join("")}</tbody>` : "";
  return `<table>${thead}${tbody}</table>`;
}

/**
 * 创建 turndown 服务实例
 */
function createTurndownService() {
  const td = new TurndownService({
    headingStyle: "atx",
    codeBlockStyle: "fenced",
    bulletListMarker: "-",
    emDelimiter: "*",
    strongDelimiter: "**",
  });

  // 移除空段落
  td.addRule("emptyParagraph", {
    filter: (node) =>
      node.nodeName === "P" && !node.textContent.trim() && !node.querySelector("img"),
    replacement: () => "",
  });

  // 启用 GFM 插件（表格、删除线、任务列表）
  td.use(gfm);

  return td;
}

/**
 * 后处理：清洗转换后的 Markdown
 */
function postProcess(md) {
  return md
    // 连续空行合并为最多两个换行
    .replace(/\n{3,}/g, "\n\n")
    // 去除行尾多余空格
    .replace(/(?<! ) {3,}$/gm, "")
    // 去除文档开头的空行
    .replace(/^\n+/, "")
    // 确保文档末尾有且仅有一个换行
    .replace(/\n*$/, "\n");
}

/**
 * 将 docx 转换为 Markdown 并写入文件
 * @param {string} docxPath - 输入 docx 文件路径
 * @param {string} outputPath - 输出 md 文件路径
 * @returns {Promise<string>} 输出文件路径
 */
export async function convertDocxFile(docxPath, outputPath) {
  const out = outputPath || docxPath.replace(/\.docx$/i, ".md");
  const md = await convertDocxToMarkdown(docxPath, path.dirname(path.resolve(out)));
  fs.writeFileSync(out, md, "utf-8");
  return out;
}
