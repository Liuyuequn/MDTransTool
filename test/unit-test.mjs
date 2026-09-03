// 单元测试：纯函数层（单位换算、字号解析、深合并、参数解析）
// 运行：node test/unit-test.mjs（npm test 会先跑本文件再跑端到端校验）

import { test } from "node:test";
import assert from "node:assert/strict";
import {
  cmToTwip,
  ptToTwip,
  ptToHalfPoint,
  parseFontSize,
  mergeOptions,
} from "../src/options.js";
import { parseArgs, parseMargin, parseHeadingSize } from "../src/args.js";

// ============ options.js：单位换算 ============

test("cmToTwip: 2.54cm = 1440 twip（1 英寸）", () => {
  assert.equal(cmToTwip(2.54), 1440);
});

test("cmToTwip: 0 = 0", () => {
  assert.equal(cmToTwip(0), 0);
});

test("cmToTwip: 3.18cm 四舍五入为 1803", () => {
  assert.equal(cmToTwip(3.18), 1803);
});

test("ptToTwip: 1pt = 20 twip", () => {
  assert.equal(ptToTwip(12), 240);
});

test("ptToHalfPoint: 四号 14pt = 28 半磅", () => {
  assert.equal(ptToHalfPoint(14), 28);
});

test("ptToHalfPoint: 五号 10.5pt = 21 半磅", () => {
  assert.equal(ptToHalfPoint(10.5), 21);
});

// ============ options.js：字号解析 ============

test("parseFontSize: 数字字符串解析为 pt", () => {
  assert.equal(parseFontSize("14"), 14);
});

test("parseFontSize: 纯数字直接返回", () => {
  assert.equal(parseFontSize(12), 12);
});

test("parseFontSize: 中文字号名映射（四号=14）", () => {
  assert.equal(parseFontSize("四号"), 14);
});

test("parseFontSize: 中文字号名映射（五号=10.5）", () => {
  assert.equal(parseFontSize("五号"), 10.5);
});

test("parseFontSize: 容忍前后空白", () => {
  assert.equal(parseFontSize("  小五 "), 9);
});

test("parseFontSize: 无法识别时抛错", () => {
  assert.throws(() => parseFontSize("胡说"), /无法识别的字号/);
});

// ============ options.js：深合并 ============

test("mergeOptions: 嵌套对象递归合并", () => {
  const base = { page: { size: "A4", margin: { top: 1, left: 1 } } };
  const patch = { page: { margin: { top: 2 } } };
  const merged = mergeOptions(base, patch);
  assert.deepEqual(merged, { page: { size: "A4", margin: { top: 2, left: 1 } } });
});

test("mergeOptions: 数组整体覆盖而非逐项合并", () => {
  const base = { sizes: { heading: [22, 16, 14, 12, 11, 11] } };
  const patch = { sizes: { heading: [22, 16, 14, 14, 14, 14] } };
  const merged = mergeOptions(base, patch);
  assert.deepEqual(merged.sizes.heading, [22, 16, 14, 14, 14, 14]);
});

test("mergeOptions: undefined 字段不覆盖已有值", () => {
  const base = { header: { text: "A" } };
  const patch = { header: { text: undefined, align: "left" } };
  const merged = mergeOptions(base, patch);
  assert.equal(merged.header.text, "A");
  assert.equal(merged.header.align, "left");
});

test("mergeOptions: 多个补丁按顺序叠加（后者胜）", () => {
  const base = { fonts: { body: { eastAsia: "等线" } } };
  const p1 = { fonts: { body: { eastAsia: "仿宋" } } };
  const p2 = { fonts: { body: { eastAsia: "黑体" } } };
  assert.equal(mergeOptions(base, p1, p2).fonts.body.eastAsia, "黑体");
});

test("mergeOptions: null 值直接覆盖（用于显式清空语义）", () => {
  const base = { header: { text: "A" } };
  const patch = { header: { text: null } };
  assert.equal(mergeOptions(base, patch).header.text, null);
});

// ============ args.js：parseMargin ============

test("parseMargin: 单值四边统一", () => {
  const patch = { page: {} };
  parseMargin("2.5", patch);
  assert.deepEqual(patch.page.margin, { top: 2.5, right: 2.5, bottom: 2.5, left: 2.5 });
});

test("parseMargin: 四值按 上,右,下,左 顺序", () => {
  const patch = { page: {} };
  parseMargin("2.54,3.18,2.54,3.18", patch);
  assert.deepEqual(patch.page.margin, { top: 2.54, right: 3.18, bottom: 2.54, left: 3.18 });
});

test("parseMargin: 容忍数值间空白", () => {
  const patch = { page: {} };
  parseMargin("1, 2, 3, 4", patch);
  assert.deepEqual(patch.page.margin, { top: 1, right: 2, bottom: 3, left: 4 });
});

test("parseMargin: 非数字抛错", () => {
  assert.throws(() => parseMargin("abc", { page: {} }), /页边距格式错误/);
});

test("parseMargin: 负数抛错", () => {
  assert.throws(() => parseMargin("-1", { page: {} }), /页边距格式错误/);
});

test("parseMargin: 三个值抛错（应为 1 或 4 个）", () => {
  assert.throws(() => parseMargin("1,2,3", { page: {} }), /1 个或 4 个/);
});

// ============ args.js：parseHeadingSize ============

test("parseHeadingSize: auto 不写入字号", () => {
  const patch = { heading: {}, sizes: {} };
  parseHeadingSize("auto", patch);
  assert.equal(patch.sizes.heading, undefined);
});

test("parseHeadingSize: 单值扩展为六级相同", () => {
  const patch = { heading: {}, sizes: {} };
  parseHeadingSize("22", patch);
  assert.deepEqual(patch.sizes.heading, [22, 22, 22, 22, 22, 22]);
});

test("parseHeadingSize: 六级逗号分隔原样生效", () => {
  const patch = { heading: {}, sizes: {} };
  parseHeadingSize("22,16,14,14,14,14", patch);
  assert.deepEqual(patch.sizes.heading, [22, 16, 14, 14, 14, 14]);
});

test("parseHeadingSize: 中文字号名可用于标题", () => {
  const patch = { heading: {}, sizes: {} };
  parseHeadingSize("二号", patch);
  assert.deepEqual(patch.sizes.heading, [22, 22, 22, 22, 22, 22]);
});

test("parseHeadingSize: 错误级数抛错", () => {
  assert.throws(() => parseHeadingSize("1,2,3", { heading: {}, sizes: {} }), /单一值或逗号分隔的六级值/);
});

// ============ args.js：parseArgs ============

test("parseArgs: 别名与全名等价（-f 与 --font）", () => {
  const a = parseArgs(["-f", "宋体"]);
  const b = parseArgs(["--font", "宋体"]);
  assert.deepEqual(a.patch.fonts.body, b.patch.fonts.body);
  assert.equal(a.patch.fonts.body.eastAsia, "宋体");
});

test("parseArgs: --name=value 等价于 --name value", () => {
  const a = parseArgs(["--font=宋体"]);
  const b = parseArgs(["--font", "宋体"]);
  assert.deepEqual(a.patch, b.patch);
});

test("parseArgs: flag 参数无需取值", () => {
  const r = parseArgs(["--no-bold"]);
  assert.deepEqual(r.patch.heading.bold, [false, false, false, false, false, false]);
});

test("parseArgs: 输出控制参数单独归位（output/overwrite/preset）", () => {
  const r = parseArgs(["-o", "out.docx", "--overwrite", "--preset", "legal"]);
  assert.equal(r.output, "out.docx");
  assert.equal(r.overwrite, true);
  assert.equal(r.preset, "legal");
});

test("parseArgs: 未知参数记入 errors", () => {
  const r = parseArgs(["--no-such-param"]);
  assert.equal(r.errors.length, 1);
  assert.match(r.errors[0], /未知参数/);
});

test("parseArgs: 参数缺少取值记入 errors", () => {
  const r = parseArgs(["--font"]);
  assert.equal(r.errors.length, 1);
  assert.match(r.errors[0], /缺少取值/);
});

test("parseArgs: 取值非法记入 errors（校验函数）", () => {
  const r = parseArgs(["--orientation", "diagonal"]);
  assert.equal(r.errors.length, 1);
  assert.match(r.errors[0], /portrait 或 landscape/);
});

test("parseArgs: 互斥页眉参数记入 errors", () => {
  const r = parseArgs(["--header", "A", "--header-left", "B"]);
  assert.ok(r.errors.some((e) => e.includes("互斥")));
});

test("parseArgs: --total-pages 在无 Y 模板上追加 /共Y页", () => {
  const r = parseArgs(["-p", "bottom", "--total-pages"]);
  assert.equal(r.patch.pageNumber.format, "X/共Y页");
});

test("parseArgs: --total-pages 不改动已含 Y 的模板", () => {
  const r = parseArgs(["--page-num-format", "第X页/共Y页", "--total-pages"]);
  assert.equal(r.patch.pageNumber.format, "第X页/共Y页");
});

test("parseArgs: 非选项 token 记为多余位置参数", () => {
  const r = parseArgs(["to"]);
  assert.equal(r.errors.length, 1);
  assert.match(r.errors[0], /多余的位置参数/);
});
