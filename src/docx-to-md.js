// docx → Markdown 转换模块
// 流程：mammoth 提取 HTML → 预处理修复结构 → turndown 转换为 Markdown → 后处理清洗

import mammoth from "mammoth";
import TurndownService from "turndown";
import { gfm } from "turndown-plugin-gfm";
import fs from "node:fs";

/**
 * 将 docx 文件转换为 Markdown 字符串
 * @param {string} docxPath - docx 文件路径
 * @returns {Promise<string>} Markdown 内容
 */
export async function convertDocxToMarkdown(docxPath) {
  if (!fs.existsSync(docxPath)) {
    throw new Error(`找不到文件 ${docxPath}`);
  }

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
        const base64 = buffer.toString("base64");
        const src = `data:${image.contentType};base64,${base64}`;
        return { src };
      }),
    }
  );

  // 2. 预处理：修复 mammoth 输出的结构性问题
  const html = preprocessHtml(rawHtml);

  // 3. turndown: HTML → Markdown
  const td = createTurndownService();
  let md = td.turndown(html);

  // 4. 后处理
  md = postProcess(md);

  return md;
}

/**
 * 预处理 HTML：修复 mammoth 常见的结构问题
 */
function preprocessHtml(html) {
  let h = html;

  // 修复表格：mammoth 把所有行都放在 <thead> 里，且数据行也用 <th>，
  // 需要：1) 数据行移到 <tbody>；2) 数据行的 <th> 改为 <td>；3) 去掉单元格内的 <p> 包裹
  h = h.replace(
    /<table>([\s\S]*?)<\/table>/g,
    (match, inner) => {
      const trMatches = [...inner.matchAll(/<tr>([\s\S]*?)<\/tr>/g)];
      if (trMatches.length === 0) return match;
      const makeRow = (trContent, isHeader) => {
        const cellTag = isHeader ? "th" : "td";
        const srcTag = isHeader ? "th" : "(?:th|td)";
        const cells = [...trContent.matchAll(new RegExp(`<${srcTag}>([\\s\\S]*?)<\\/${srcTag}>`, "g"))];
        const row = cells.map((m) => `<${cellTag}>${m[1].replace(/<\/?p>/g, "").trim()}</${cellTag}>`).join("");
        return `<tr>${row}</tr>`;
      };
      const headerRow = makeRow(trMatches[0][1], true);
      const bodyRows = trMatches.slice(1).map((m) => makeRow(m[1], false)).join("");
      return `<table><thead>${headerRow}</thead><tbody>${bodyRows}</tbody></table>`;
    }
  );

  // 列表项内部的段落：<li><p>text</p></li> → <li>text</li>
  h = h.replace(/<li>\s*<p>([\s\S]*?)<\/p>\s*<\/li>/g, "<li>$1</li>");

  return h;
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
  const md = await convertDocxToMarkdown(docxPath);
  const out = outputPath || docxPath.replace(/\.docx$/i, ".md");
  fs.writeFileSync(out, md, "utf-8");
  return out;
}
