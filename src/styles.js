// 样式与编号：docx 样式表、字体对象、间距换算、列表编号配置
// 被 converter（组装 Document）、blocks（代码块字体）、header-footer（页眉页脚字体）共用

import {
  AlignmentType,
  LevelFormat,
  LineRuleType,
  UnderlineType,
} from "docx";
import { ptToHalfPoint, ptToTwip } from "./options.js";

const BULLET_CHARS = ["\u2022", "\u25E6", "\u2022", "\u25E6", "\u2022", "\u25E6"];

/** 对齐方式字符串 → docx 枚举 */
export const ALIGN_MAP = {
  left: AlignmentType.LEFT,
  center: AlignmentType.CENTER,
  right: AlignmentType.RIGHT,
  justify: AlignmentType.JUSTIFIED,
};

/** docx 字体对象：中文 eastAsia / 西文 ascii */
export function fontObj(f) {
  return { ascii: f.ascii, hAnsi: f.ascii, eastAsia: f.eastAsia };
}

/** 标题字体：未单独指定时跟随正文 */
export function headingFont(opts) {
  return fontObj(opts.fonts.heading || opts.fonts.body);
}

/**
 * 解析间距：lines 字段优先（按 字号×行距倍数 换算 twip），否则用 pt；
 * line 为行距倍数（w:line = 倍数×240，lineRule=auto）
 */
export function resolveSpacing(spacing, sizePt, lineMultiplier) {
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

/** "行"单位 → twip：1 行 ≈ 字号 × 行距倍数 */
export function linesToTwip(lines, sizePt, lineMultiplier = 1.28) {
  return Math.round(lines * sizePt * lineMultiplier * 20);
}

/** docx 样式表：正文默认 + 六级标题 + 超链接字符样式 */
export function buildStyles(opts) {
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

export function buildBulletNumbering() {
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

export function buildOrderedNumbering() {
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
