// build_template.js — recreate a vertical flowchart image as ONE editable PPTX slide.
//
// This is a working, heavily-commented starting point (the medical-research
// flowchart it produces is a real example). To reuse:
//   1. Replace the TEXT CONTENT with your transcribed + OCR-cleaned strings.
//   2. Adjust the PALETTE block to colors sampled from your source image.
//   3. Adjust layout height to match your image aspect: h ≈ 10 * (origH / origW).
//   4. Set OUT to an absolute output path.
//   5. Run from a dir where pptxgenjs is installed:  node build_template.js
//
// Then render + eyeball with scripts/render_qa.sh and iterate.

const pptxgen = require("pptxgenjs");
const p = new pptxgen();

// ---- OUTPUT PATH (edit me) ----
const OUT = "/Users/you/Desktop/output.pptx";

// ---- LAYOUT: portrait, matched to source aspect ratio ----
p.defineLayout({ name: "FLOW", width: 10, height: 16.4 });
p.layout = "FLOW";

const FONT = "PingFang SC"; // CJK face — required for Chinese text

// ---- PALETTE (sample these from the source image) ----
const NAVY = "1B2A6B", NAVYTX = "1E2761";
const PURPLE = "7A4AA6", PURPLE_D = "5B2C87", DARKP = "5A2A86";
const LAV1 = "F3EEFA", LAV2 = "EBE0F5", LAVLINE = "D9C9EC";
const INK = "2E2A45", GRAY = "6E6A85", GRAYLN = "9A94AE";
const WHITE = "FFFFFF", ACCENT = "C4A6E8";

const s = p.addSlide();
s.background = { color: WHITE };

const CX = 5.0;                         // horizontal center
const LX = 0.5, RW = 4.35, RX = 5.15;   // two-column geometry
const LCX = LX + RW / 2, RCX = RX + RW / 2;

// gray down-arrow connector
function arrow(x, y, h) {
  s.addShape(p.shapes.LINE, { x, y, w: 0, h,
    line: { color: GRAYLN, width: 2.5, endArrowType: "triangle" } });
}

// ---- TITLE + SUBTITLE ----
s.addText("研究内容四：靶向胃肠激素网络的术后低血压干预与因果验证", {
  x: 0.4, y: 0.35, w: 9.2, h: 0.9, fontFace: FONT, fontSize: 23, color: NAVYTX,
  bold: true, align: "center", valign: "middle", lineSpacingMultiple: 1.05 });
s.addText([
  { text: "随机双盲安慰剂对照交叉设计" },
  { text: "  ·  ", options: { color: PURPLE } },
  { text: "分层决策药物选择" },
  { text: "  ·  ", options: { color: PURPLE } },
  { text: "四层递进因果验证" },
], { x: 0.5, y: 1.28, w: 9.0, h: 0.4, fontFace: FONT, fontSize: 13, color: GRAY, align: "center" });

// ---- BOX with header sitting on its top edge (flanking lines) ----
s.addShape(p.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.8, w: 9.0, h: 2.15,
  fill: { color: "FBFAFE" }, line: { color: "C9C6D6", width: 1.25 }, rectRadius: 0.1 });
s.addShape(p.shapes.LINE, { x: 2.3, y: 2.18, w: 1.25, h: 0, line: { color: GRAYLN, width: 1.5 } });
s.addShape(p.shapes.LINE, { x: 6.45, y: 2.18, w: 1.25, h: 0, line: { color: GRAYLN, width: 1.5 } });
s.addText("研究设计与人群", { x: 3.0, y: 1.93, w: 4.0, h: 0.5, fontFace: FONT,
  fontSize: 18, color: NAVYTX, bold: true, align: "center", valign: "middle" });
s.addText([
  { text: "研究设计：", options: { bold: true, color: NAVYTX } },
  { text: "随机、双盲、安慰剂对照、两阶段交叉试验", options: { color: INK } },
], { x: 0.6, y: 2.5, w: 8.8, h: 0.45, fontFace: FONT, fontSize: 14, align: "center", valign: "middle" });
s.addText([
  { text: "受试者：", options: { bold: true, color: NAVYTX } },
  { text: "术后症状性低血压患者 ", options: { color: INK } },
  { text: "20", options: { color: PURPLE, bold: true, fontSize: 17 } },
  { text: " 例", options: { color: INK } },
], { x: 0.6, y: 3.02, w: 8.8, h: 0.45, fontFace: FONT, fontSize: 14, align: "center", valign: "middle" });
s.addText([
  { text: "A组：", options: { bold: true, color: PURPLE } },
  { text: "药物 → 洗脱 ≥ 7天 → 安慰剂", options: { color: INK } },
  { text: "　｜　", options: { color: GRAYLN } },
  { text: "B组：", options: { bold: true, color: NAVYTX } },
  { text: "安慰剂 → 洗脱 ≥ 7天 → 药物", options: { color: INK } },
], { x: 0.6, y: 3.5, w: 8.8, h: 0.42, fontFace: FONT, fontSize: 13.5, align: "center", valign: "middle" });

arrow(CX, 3.98, 0.42);

// ---- single emphasized bar ----
s.addShape(p.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 4.45, w: 9.0, h: 0.8,
  fill: { color: WHITE }, line: { color: NAVYTX, width: 1.5 }, rectRadius: 0.1 });
s.addText("靶向药物分层决策（基于前期核心发现）", { x: 0.5, y: 4.45, w: 9.0, h: 0.8,
  fontFace: FONT, fontSize: 19, color: NAVYTX, bold: true, align: "center", valign: "middle" });

arrow(CX, 5.28, 0.42);

// ---- two-column cards ----
function planCard(x, titleTop, big, note, headColor) {
  s.addShape(p.shapes.ROUNDED_RECTANGLE, { x, y: 5.75, w: RW, h: 1.62,
    fill: { color: LAV1 }, line: { color: LAVLINE, width: 1.25 }, rectRadius: 0.1 });
  s.addText(titleTop, { x: x + 0.15, y: 5.9, w: RW - 0.3, h: 0.5, fontFace: FONT,
    fontSize: 17, color: headColor, bold: true, align: "center", valign: "middle" });
  s.addText(big, { x: x + 0.15, y: 6.42, w: RW - 0.3, h: 0.45, fontFace: FONT,
    fontSize: 16, color: PURPLE_D, bold: true, align: "center", valign: "middle" });
  s.addText(note, { x: x + 0.15, y: 6.92, w: RW - 0.3, h: 0.38, fontFace: FONT,
    fontSize: 12.5, color: GRAY, align: "center", valign: "middle" });
}
planCard(LX, "方案 A（单一激素主导时）", "高选择性靶点拮抗剂", "特异性阻断单一核心激素", NAVYTX);
planCard(RX, "方案 B（共享模块主导时）", "奥曲肽广谱抑制方案", "同步抑制共享模块，默认方案", PURPLE_D);

// converging arrows + centered label between the columns
arrow(LCX, 7.4, 0.62);
arrow(RCX, 7.4, 0.62);
s.addText("最终决策以网络分析结果为依据", { x: 3.0, y: 7.55, w: 4.0, h: 0.4,
  fontFace: FONT, fontSize: 13.5, color: GRAY, align: "center", valign: "middle" });

// ---- numbered validation grid; one card highlighted (dark) ----
function vCard(x, y, num, title, desc, tag, dark) {
  s.addShape(p.shapes.ROUNDED_RECTANGLE, { x, y, w: RW, h: 1.5,
    fill: { color: dark ? DARKP : LAV1 },
    line: { color: dark ? DARKP : LAVLINE, width: 1.25 }, rectRadius: 0.1 });
  s.addText([
    { text: num + "  ", options: { color: dark ? ACCENT : PURPLE, bold: true } },
    { text: title, options: { color: dark ? WHITE : NAVYTX, bold: true } },
  ], { x: x + 0.2, y: y + 0.14, w: RW - 0.4, h: 0.5, fontFace: FONT, fontSize: 17, align: "center", valign: "middle" });
  s.addText(desc, { x: x + 0.2, y: y + 0.66, w: RW - 0.4, h: 0.45, fontFace: FONT,
    fontSize: 13, color: dark ? "F0E8FA" : INK, bold: dark, align: "center", valign: "middle" });
  s.addText(tag, { x: x + 0.2, y: y + 1.1, w: RW - 0.4, h: 0.34, fontFace: FONT,
    fontSize: 12, color: dark ? ACCENT : GRAY, align: "center", valign: "middle" });
}
vCard(LX, 8.1, "①", "靶点结合验证", "目标激素 AUC 显著下降，非靶激素无变化？", "验证药物对预设靶点的调控特性", false);
vCard(RX, 8.1, "②", "通路逆转验证", "下游 NOx / 醛固酮通路指标同步逆转", "验证「激素 → 通路」的因果方向", false);
vCard(LX, 9.7, "③", "表型逆转验证", "餐后血压下降幅度减小，临床症状改善", "验证「通路 → 表型」的因果方向", false);
vCard(RX, 9.7, "④", "中介效应验证", "药物 → Δ激素 → Δ通路 → Δ血压  链式中介", "间接路径显著，闭合完整因果链", true);

arrow(CX, 11.25, 0.42);

// ---- dark conclusion bar ----
s.addShape(p.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 11.72, w: 9.0, h: 1.15,
  fill: { color: NAVY }, line: { color: NAVY, width: 1 }, rectRadius: 0.08 });
s.addText("四层验证全部成立 → 因果证据链闭合", { x: 0.6, y: 11.85, w: 8.8, h: 0.55,
  fontFace: FONT, fontSize: 20, color: WHITE, bold: true, align: "center", valign: "middle" });
s.addText("反向验证核心假说：胃肠激素网络调控术后血压", { x: 0.6, y: 12.42, w: 8.8, h: 0.38,
  fontFace: FONT, fontSize: 13.5, color: ACCENT, align: "center", valign: "middle" });

arrow(CX, 12.93, 0.42);

// ---- label + output box ----
s.addText("研究产出", { x: 0.5, y: 13.38, w: 9.0, h: 0.42, fontFace: FONT,
  fontSize: 17, color: NAVYTX, bold: true, align: "center", valign: "middle" });
s.addShape(p.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 13.85, w: 9.0, h: 0.78,
  fill: { color: "FBFAFE" }, line: { color: "C9C6D6", width: 1.25 }, rectRadius: 0.1 });
s.addText([
  { text: "产出：", options: { bold: true, color: NAVYTX } },
  { text: "术后低血压靶向干预方案", options: { color: INK } },
  { text: "  ·  ", options: { color: PURPLE } },
  { text: "胃肠激素调控血压因果证据", options: { color: INK } },
  { text: "  ·  ", options: { color: PURPLE } },
  { text: "后续大样本试验方法学依据", options: { color: INK } },
], { x: 0.6, y: 13.85, w: 8.8, h: 0.78, fontFace: FONT, fontSize: 13.5, align: "center", valign: "middle" });

arrow(CX, 14.66, 0.4);

// ---- final requirements box ----
s.addShape(p.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 15.08, w: 9.0, h: 0.92,
  fill: { color: WHITE }, line: { color: "C9C6D6", width: 1.25 }, rectRadius: 0.1 });
s.addText("【因果样式要求】", { x: 0.5, y: 15.16, w: 9.0, h: 0.42, fontFace: FONT,
  fontSize: 16, color: NAVYTX, bold: true, align: "center", valign: "middle" });
s.addText([
  { text: "因果验证逻辑：", options: { bold: true, color: NAVYTX } },
  { text: "靶点结合 → 通路逆转 → 表型改善 → 中介闭合", options: { color: INK } },
  { text: "　·　四层递进　·　环环相扣", options: { color: PURPLE, bold: true } },
], { x: 0.6, y: 15.56, w: 8.8, h: 0.4, fontFace: FONT, fontSize: 13, align: "center", valign: "middle" });

// ---- footer ----
s.addText("字体：无衬线字体　·　标题 16pt　·　正文 10–11pt　·　小文字 8–9pt", {
  x: 0.5, y: 16.05, w: 9.0, h: 0.3, fontFace: FONT, fontSize: 10, color: "A6A2B5", align: "center", valign: "middle" });

p.writeFile({ fileName: OUT }).then(f => console.log("SAVED:", f));
