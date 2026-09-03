// 预设方案：每个预设是与 defaultOptions 同形的补丁对象
// legal 为已锁定的法律文书规范；其余预设为初版，后续可按需精调

import path from "node:path";
import { fileURLToPath } from "node:url";

const pkgRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

export const presets = {
  // 法律文书（已锁定）
  legal: {
    page: {
      size: "A4",
      orientation: "portrait",
      margin: { top: 2.54, right: 3.18, bottom: 2.54, left: 3.18 },
    },
    fonts: {
      body: { eastAsia: "仿宋", ascii: "Times New Roman" },
      heading: { eastAsia: "宋体", ascii: "Times New Roman" },
      header: { eastAsia: "仿宋", ascii: "Times New Roman" },
      footer: { eastAsia: "仿宋", ascii: "Times New Roman" },
    },
    sizes: {
      body: 14,                          // 四号
      heading: [22, 16, 14, 14, 14, 14], // 二号/三号/四号×4
      header: 9,                         // 页眉小五
      footer: 10.5,                       // 页脚页码五号
    },
    heading: {
      bold: [true, true, true, false, false, false],
      align: ["center", "left", "left", "left", "left", "left"],
      spacing: { beforeLines: 0.5, afterLines: 0 },
    },
    paragraph: {
      firstLineChars: 2,              // 首行缩进两字符
      afterLines: 0.5,                // 段间距 0.5 行
      line: 1.28,                     // 行距 1.28 行
    },
    header: {
      text: "圣典律师事务所\n圣典官网：https://www.sundylawyer.com/\n总所地址：南京市建邺区奥体大街68号新城科技园4A栋6楼、7楼",
      align: "left",
      image: { path: path.join(pkgRoot, "reference", "圣典律师.png"), heightCm: 1.2 },
      // 页眉底端渐变色带：亮正红 → 红橙 → 橙 → 金，橙金从 45% 起即明显（占比约 55%）
      gradientLine: {
        stops: [[0, "E60012"], [0.45, "FF5A00"], [0.7, "FFA500"], [1, "FFD700"]],
        heightCm: 0.06,
        gapCm: 0.2, // 色带与上方页眉文字之间的间距
      },
    },
    pageNumber: {
      pos: "bottom",
      align: "center",
      format: "第X页/共Y页",
    },
  },

  // 报告（初版）：A4、微软雅黑、页脚居中页码
  report: {
    page: {
      size: "A4",
      orientation: "portrait",
      margin: { top: 2.54, right: 2.54, bottom: 2.54, left: 2.54 },
    },
    fonts: {
      body: { eastAsia: "微软雅黑", ascii: "Calibri" },
      heading: { eastAsia: "微软雅黑", ascii: "Calibri" },
    },
    sizes: { body: 11, heading: [22, 16, 14, 12, 11, 11] },
    paragraph: { afterPt: 8, line: 1.15 },
    pageNumber: { pos: "bottom", align: "center", format: "X / Y" },
  },

  // 紧凑排版（初版）：小字号、窄边距，适合长文
  compact: {
    page: {
      size: "A4",
      margin: { top: 1.5, right: 1.5, bottom: 1.5, left: 1.5 },
    },
    fonts: {
      body: { eastAsia: "等线", ascii: "Calibri" },
    },
    sizes: { body: 10.5, heading: [16, 14, 12, 11, 10.5, 10.5], code: 8.5 },
    paragraph: { afterPt: 4, line: 1.15 },
  },

  // 封面页（初版）：内容垂直居中，无页眉页脚页码
  cover: {
    page: { vAlign: "center" },
    header: { text: null, left: null, right: null },
    footer: { text: null, left: null, right: null },
    pageNumber: { pos: "none" },
  },

  // 显式指定默认值
  default: {},
};
