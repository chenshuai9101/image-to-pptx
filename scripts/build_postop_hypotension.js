// build_postop_hypotension.js - pptxgenjs v4 API
const pptxgen = require("pptxgenjs");
const p = new pptxgen();

const OUT = "/Users/muyunye/Desktop/术后低血压研究设计图_可编辑版.pptx";

p.defineLayout({ name: "FLOW", width: 10, height: 15 });
p.layout = "FLOW";

const S = p.shapes;
// Slide helpers
const s = p.addSlide();
s.background = { fill: "FFFFFF" };
const FONT = "PingFang SC";

// Colors
const NAVY = "1B2A4A";
const BLUE = "2C6FAE", BLUEL = "D6E4F0";
const TEAL = "219E8C", TEALL = "E0F2F1";
const ORANGE = "F89C2C", ORANGEL = "FDF3E0";
const RED = "E53E3E", REDL = "FDE6E6";
const GRY = "95A5A6", LGRY = "BDC3C7", DGRY = "B0B0B0";
const WHITE = "FFFFFF", BLACK = "333333";

function T(x, y, w, h, text, opts = {}) {
    const opt = {
        x, y, w, h, fontFace: FONT,
        fontSize: opts.sz || 10, bold: opts.b || false,
        color: opts.c || BLACK, align: opts.a || "left",
        valign: opts.v || "top",
    };
    if (Array.isArray(text)) {
        opt.paraSpaceAfter = 2;
        return s.addText(text.map(t => t.text ? t : { text: "", options: { fontSize: 6 } }), opt);
    }
    return s.addText(text, opt);
}

function shape(type, x, y, w, h, fill, stroke) {
    const opt = { x, y, w, h, fill: { color: fill } };
    if (stroke) { opt.line = { color: stroke, width: 1.5 }; } else { opt.line = { width: 0 }; }
    if (type === S.ROUNDED_RECTANGLE) opt.rectRadius = 0.08;
    s.addShape(type, opt);
}

// ══════ BUILD ══════

// 1. Top navy bar
shape(S.RECTANGLE, 0, 0, 10, 0.3, NAVY);
T(0.4, 0.08, 9.2, 0.25, "研究内容四：靶向胃肠激素网络的术后低血压干预与因果验证",
  { sz: 13, b: true, c: WHITE });

// 2. Blue subtitle
T(0.4, 0.4, 9.2, 0.18, "随机双盲安慰剂对照两阶段交叉设计  ·  分层决策药物选择  ·  四层递进因果验证",
  { sz: 8, c: "4A90D9" });

// 3. Blue bar + label
shape(S.RECTANGLE, 0.4, 0.6, 0.06, 0.16, BLUE);
T(0.52, 0.6, 3, 0.16, "研究背景与方法", { sz: 8, b: true, c: BLUE });

// 4. Left box: Study design
shape(S.ROUNDED_RECTANGLE, 0.4, 0.88, 4.4, 1.1, BLUEL, BLUE);
T(0.55, 0.9, 4, 0.2, "1. 研究设计与人群", { sz: 10, b: true, c: NAVY });
T(0.55, 1.12, 4, 0.8, [
    { text: "研究设计：随机、双盲、安慰剂对照、两阶段交叉试验", options: { fontSize: 8 } },
    { text: "受试者：术后症状性低血压患者 20 例", options: { fontSize: 8 } },
    {},
    { text: "A 组：药物 → 洗脱 ≥7 天 → 安慰剂", options: { fontSize: 8, bold: true, color: NAVY } },
    { text: "B 组：安慰剂 → 洗脱 ≥7 天 → 药物", options: { fontSize: 8, bold: true, color: NAVY } },
]);

// 5. Right box: Drug strategy
shape(S.ROUNDED_RECTANGLE, 5.2, 0.88, 4.4, 1.1, TEALL, TEAL);
T(5.35, 0.9, 4, 0.2, "2. 靶向药物分层决策（基于前期核心发现）", { sz: 8, b: true, c: TEAL });
T(5.35, 1.12, 4, 0.8, [
    { text: "方案 A（单一激素主导时）：", options: { fontSize: 8, bold: true, color: NAVY } },
    { text: "  高选择性靶点拮抗剂", options: { fontSize: 8 } },
    { text: "  特异性阻断单一核心激素信号通路", options: { fontSize: 8, color: GRY } },
    {},
    { text: "方案 B（共分泌模块主导时）：", options: { fontSize: 8, bold: true, color: NAVY } },
    { text: "  奥曲肽广谱抑制方案（默认方案）", options: { fontSize: 8 } },
    { text: "  同步抑制共分泌模块，覆盖多激素", options: { fontSize: 8, color: GRY } },
]);

// 6. Down connectors + decision badge
shape(S.RECTANGLE, 2.6, 2.0, 0.04, 0.18, LGRY);
shape(S.RECTANGLE, 7.4, 2.0, 0.04, 0.18, LGRY);

shape(S.ROUNDED_RECTANGLE, 2.5, 2.25, 1.1, 0.2, ORANGE);
T(2.5, 2.25, 1.1, 0.2, "最终决策", { sz: 7, b: true, c: WHITE, a: "center", v: "middle" });
T(3.7, 2.25, 2.5, 0.2, "网络分析结果", { sz: 8, b: true, c: ORANGE, v: "middle" });

T(4.5, 2.5, 1, 0.2, "▼", { sz: 12, c: ORANGE, a: "center", v: "middle" });

// 7. Divider
shape(S.RECTANGLE, 0.4, 2.78, 9.2, 0.025, ORANGE);

// 8. Section title
T(0.4, 2.85, 5, 0.22, "3. 四层递进因果验证", { sz: 11, b: true, c: NAVY });

// 9. Vertical spine
shape(S.RECTANGLE, 0.7, 3.15, 0.035, 4.2, LGRY);

// 10. Four layers
const layers = [
    { n: "1", t: "① 靶点结合验证",  c: BLUE,  bg: BLUEL,
      q: "核心问题：目标激素 AUC 显著下降，非靶激素无变化？",
      p: "验证目的：验证药物对预设靶点的调控特异性" },
    { n: "2", t: "② 通路逆转验证",  c: TEAL,  bg: TEALL,
      q: "核心问题：下游 NOx / 醛固酮通路指标同步逆转？",
      p: "验证目的：验证靶点-表型间的因果通路连接" },
    { n: "3", t: "③ 网络重编程验证", c: ORANGE, bg: ORANGEL,
      q: "核心问题：激素网络拓扑结构重构？正常化趋势？",
      p: "验证目的：验证干预在系统层面的整体效应" },
    { n: "4", t: "④ 非靶向因果排除", c: RED,   bg: REDL,
      q: "核心问题：排除非特异性效应 / 安慰剂效应混杂影响？",
      p: "验证目的：确证因果效应的特异性与排他性" },
];

layers.forEach((ly, i) => {
    const yb = 3.15 + i * 1.05;
    
    // Number circle
    shape(S.OVAL, 0.58, yb + 0.06, 0.28, 0.22, ly.c);
    T(0.58, yb + 0.06, 0.28, 0.22, ly.n,
      { sz: 8, b: true, c: WHITE, a: "center", v: "middle" });
    
    // Dash line to title
    shape(S.RECTANGLE, 0.86, yb + 0.14, 0.15, 0.03, DGRY);
    
    // Title
    T(1.05, yb, 2.2, 0.25, ly.t, { sz: 9, b: true, c: ly.c, v: "middle" });
    
    // Content box
    shape(S.ROUNDED_RECTANGLE, 3.4, yb - 0.05, 6.2, 0.7, ly.bg, ly.c);
    T(3.6, yb + 0.02, 5.8, 0.3, "• " + ly.q, { sz: 7.5, c: BLACK, v: "middle" });
    T(3.6, yb + 0.32, 5.8, 0.3, "• " + ly.p, { sz: 7.5, c: GRY, v: "middle" });
    
    // Down arrow
    if (i < 3) T(0.65, yb + 0.65, 0.25, 0.2, "▼", { sz: 7, c: LGRY, a: "center", v: "middle" });
});

// 11. Right markers
[3.4, 4.45, 5.5].forEach((x, i) => {
    T(x, 3.35 + i*1.05, 0.3, 0.12, "▸", { sz: 6, c: LGRY });
});

// 12. Summary
shape(S.RECTANGLE, 0.4, 5.55, 3, 0.02, LGRY);
T(0.4, 5.6, 9.2, 0.6, [
    { text: "研究总结", options: { fontSize: 8, bold: true, color: NAVY } },
    {},
    { text: "本研究采用随机双盲两阶段交叉设计，20 例术后症状性低血压患者入组。基于前期核心发现制定靶向药物分层决策（单一激素主导→高选择性拮抗剂；共分泌模块主导→奥曲肽方案）。通过四层递进因果验证框架，从靶点结合验证、通路逆转验证、网络重编程验证到非靶向因果排除，逐层确证因果效应。",
      options: { fontSize: 6.5, color: BLACK } },
]);

// 13. Bottom bar
shape(S.RECTANGLE, 0, 6.35, 10, 0.2, NAVY);
T(0.4, 6.35, 9.2, 0.2, "研究内容四：靶向胃肠激素网络的术后低血压干预与因果验证",
  { sz: 7, c: WHITE, v: "middle" });

p.writeFile({ fileName: OUT }).then(() => {
    console.log("✅ PPT saved → " + OUT);
}).catch(err => console.error("❌", err));
