// 行内元素解析：加粗/斜体/删除线/行内代码/超链接/图片/换行
// 以及图片文件加载（正文图片与页眉 logo 共用）

import fs from "node:fs";
import path from "node:path";
import { imageSize } from "image-size";
import {
  ExternalHyperlink,
  ImageRun,
  ShadingType,
  TextRun,
} from "docx";
import { ptToHalfPoint } from "./options.js";
import { fontObj } from "./styles.js";

const IMAGE_TYPES = { ".png": "png", ".jpg": "jpg", ".jpeg": "jpg", ".gif": "gif", ".bmp": "bmp" };
const MAX_IMAGE_WIDTH = 600; // 像素，超出按比例缩小

/** 读取 markdown-it token 的属性 */
export function getAttr(token, name) {
  const attr = (token.attrs || []).find(([k]) => k === name);
  return attr ? attr[1] : null;
}

/** 加载并校验图片文件：返回 { data, dim, type } */
export function loadImage(absPath) {
  if (!fs.existsSync(absPath)) throw new Error(`找不到图片文件: ${absPath}`);
  const data = fs.readFileSync(absPath);
  const dim = imageSize(data);
  const type = IMAGE_TYPES[path.extname(absPath).toLowerCase()];
  if (!type) throw new Error(`不支持的图片格式: ${path.extname(absPath)}`);
  return { data, dim, type };
}

/** 解析行内 token 序列为 docx run 序列 */
export function parseInline(children, ctx) {
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

/** 正文图片：按 MAX_IMAGE_WIDTH 等比缩放 */
function makeImageRun(token, ctx) {
  const src = getAttr(token, "src");
  if (!src) throw new Error("图片缺少 src 属性");
  const absPath = path.resolve(ctx.basePath, decodeURIComponent(src.replace(/^file:\/\//, "")));
  const { data, dim, type } = loadImage(absPath);
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
