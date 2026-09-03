// 页面属性：尺寸、内容区宽度、垂直对齐、页码数字格式
// 被 converter（文档组装）与 header-footer（内容区宽度）共用

import { PageOrientation, NumberFormat, VerticalAlignSection } from "docx";
import { PAGE_SIZES, cmToTwip } from "./options.js";

/** 页码格式特殊值 → docx 数字样式 */
export const NUM_FORMAT_MAP = {
  "1": NumberFormat.DECIMAL,
  i: NumberFormat.LOWER_ROMAN,
  I: NumberFormat.UPPER_ROMAN,
  一: NumberFormat.CHINESE_COUNTING,
};

/** 页面尺寸：预设名或 "W,H"（cm），横向时自动交换宽高 */
export function buildPageSize(opts) {
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

/** 页面内容区宽度（twip）：用于左右分布页眉的右对齐制表位、页眉渐变色带铺满 */
export function contentWidthTwip(opts) {
  const size = buildPageSize(opts);
  return size.width - cmToTwip(opts.page.margin.left) - cmToTwip(opts.page.margin.right);
}

export function toSectionVAlign(v) {
  return v === "center" ? VerticalAlignSection.CENTER : v === "bottom" ? VerticalAlignSection.BOTTOM : VerticalAlignSection.TOP;
}

/** 起始页码与页码数字样式 */
export function buildPageNumbers(opts) {
  const pn = opts.pageNumber;
  const result = {};
  if (pn.start != null) result.start = pn.start;
  const fmt = NUM_FORMAT_MAP[pn.format];
  if (fmt) result.formatType = fmt;
  return Object.keys(result).length ? result : undefined;
}
