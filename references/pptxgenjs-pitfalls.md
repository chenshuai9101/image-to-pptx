# pptxgenjs — pitfalls & essential techniques

These are the mistakes that silently corrupt files or produce ugly output.
Read before writing generation code.

## File-corrupting mistakes (avoid absolutely)

1. **No `#` in hex colors.** `color: "FF0000"` ✅  ·  `color: "#FF0000"` ❌ corrupts the file.
2. **Never encode opacity in an 8-char hex string.** `color: "00000020"` ❌.
   Use the `opacity` property: `shadow: { type:"outer", color:"000000", opacity:0.15 }`.
3. **Don't reuse an options object across calls.** pptxgenjs mutates objects
   in place (e.g. shadow → EMU). Sharing one shadow object between two shapes
   corrupts the second. Use a factory:
   ```js
   const shadow = () => ({ type:"outer", color:"6B3FA0", blur:8, offset:3, angle:90, opacity:0.18 });
   addShape(..., { shadow: shadow() }); addShape(..., { shadow: shadow() });
   ```
4. **Shadow `offset` must be ≥ 0.** Negative offsets corrupt the file; to throw a
   shadow upward use `angle: 270` with a positive offset.
5. **Fresh `new pptxgen()` per file.** Don't reuse the instance.

## Layout & shapes

- Custom size: `p.defineLayout({ name:"FLOW", width:10, height:16.4 }); p.layout="FLOW";`
  (units are inches). Built-ins: `LAYOUT_WIDE` = 13.3×7.5, `LAYOUT_16x9` = 10×5.625.
- `rectRadius` only works on `ROUNDED_RECTANGLE`, not `RECTANGLE`.
- Don't overlay a rectangular accent bar on a `ROUNDED_RECTANGLE` — it won't
  cover the rounded corner. Use `RECTANGLE` for the base if you need edge bars.

## Arrows (flowchart connectors)

A vertical connector arrow = a zero-width LINE with an arrowhead:
```js
s.addShape(p.shapes.LINE, { x:5, y:4, w:0, h:0.42,
  line: { color:"9A94AE", width:2.5, endArrowType:"triangle" } });
```
`endArrowType`/`beginArrowType` accept `"triangle"`, `"arrow"`, `"stealth"`,
`"diamond"`, `"oval"`. The arrow points along the line's direction. A horizontal
arrow is a zero-**height** LINE (`w:0.36, h:0`).

## Connector TOPOLOGY — match the source's branch style (critical for fidelity)

A single center arrow between rows is almost never what the source draws. When
one node connects to several (or several to one), the source uses a specific
*shape* of connector. Reproducing the wrong shape is the most common "not 1:1"
complaint. **Before coding, look at each fan-out/fan-in and classify it as fork,
merge, or bridge — they are not interchangeable.**

Build them from the primitives below (drop these helpers into every flowchart
script). `CX` = canvas center x; `centers` = array of the target/source column
center x-coordinates.

**1. Fork / comb (top-down: ONE label → MANY boxes).** A stub down from the
source, a horizontal *distribution* bus, then a drop arrow into each box top.
Used when a stage label splits into columns (e.g. a "分析方法" label above 3
method boxes).
```js
function fork(s, centers, busY, dropTop, srcTopY, srcX = CX, c = "86909C") {
  if (srcTopY != null && busY > srcTopY)              // stub from source down to bus
    s.addShape(S.LINE, { x:srcX, y:srcTopY, w:0, h:busY-srcTopY, line:{color:c,width:2} });
  const xs = centers.concat([srcX]), minX = Math.min(...xs), maxX = Math.max(...xs);
  s.addShape(S.LINE, { x:minX, y:busY, w:maxX-minX, h:0, line:{color:c,width:2} }); // bus
  centers.forEach(cx =>                               // drop ARROWS into each box top
    s.addShape(S.LINE, { x:cx, y:busY, w:0, h:dropTop-busY, line:{color:c,width:2,endArrowType:"triangle"} }));
}
```

**2. Merge (bottom-up: MANY boxes → ONE label).** The mirror image: a line up
out of each box bottom, a horizontal *collection* bus, then a single arrow down
from center into the next label. Used when a row of boxes converges before the
next stage (e.g. two pathway boxes joining down to "独立性确认").
```js
function merge(s, centers, boxBottomY, busY, arrowEndY, dstX = CX, c = "86909C") {
  centers.forEach(cx =>                               // riser from each box bottom to bus
    s.addShape(S.LINE, { x:cx, y:boxBottomY, w:0, h:busY-boxBottomY, line:{color:c,width:2} }));
  const xs = centers.concat([dstX]), minX = Math.min(...xs), maxX = Math.max(...xs);
  s.addShape(S.LINE, { x:minX, y:busY, w:maxX-minX, h:0, line:{color:c,width:2} }); // bus
  s.addShape(S.LINE, { x:dstX, y:busY, w:0, h:arrowEndY-busY, line:{color:c,width:2,endArrowType:"triangle"} });
}
```

**3. Bridge (sibling → sibling).** A short horizontal arrow in the gap between
two adjacent boxes (e.g. 发现集 → 验证集 → 全队列). Draw it at the boxes' vertical
midline: `s.addShape(S.LINE, { x:boxRightEdge+0.02, y:rowMidY, w:gap-0.04, h:0, line:{...endArrowType:"triangle"} })`.

**How to tell fork from merge — look at where the arrowHEADS are:**
- Arrowheads point *down into the box tops* → the entry is a **fork**; the single
  incoming arrow ends at the *label* above the boxes (no arrow between label and box tops).
- Arrowheads point *down into the next label* while the lines *emerge from the box
  bottoms* → it's a **merge**.
- One diagram can mix styles per row. Don't assume all rows use the same one, and
  don't assume a whole deck is uniform — e.g. one slide may fork every stage while
  a sibling slide merges every stage. Classify each transition from the image.

## Framed sections with straddling pills & dividers

Some posters wrap the whole flow in one outer rectangle split into horizontal
*zones* by full-width divider lines, with a section-label "pill" (rounded
capsule) sitting *on* each divider, and faint thin rules between rows inside a
multi-row box. To reproduce faithfully:
- Outer frame: a `RECTANGLE` with `fill:{type:"none"}` and a thin border.
- Zone divider: a full-width very-thin `RECTANGLE` (`h:0.014`, mid-gray fill),
  drawn edge-to-edge across the frame's inner width.
- Straddling pill: draw the divider first, then a `ROUNDED_RECTANGLE` pill
  (opaque fill) centered on that y so it covers the line's middle; then its text.
- Row divider inside a box: a fainter, slightly-inset thin `RECTANGLE`
  (`h:0.01`, light-blue fill).
- Respect the source's box **grouping**: rows that share one bordered box vs.
  rows in separate boxes are a real distinction — don't collapse three separate
  boxes into one, or split one into three. Transcribe the box boundaries.

## Text

- **CJK**: always set `fontFace: "PingFang SC"` (macOS) or another installed CJK
  face, or Chinese renders in a substituted font.
- **Multi-line**: use the array form with `breakLine: true`, or `\n` with
  `lineSpacingMultiple`.
- **Mixed styling in one line**: pass an array of runs, each with its own
  `options` (bold / color / fontSize):
  ```js
  s.addText([
    { text:"研究设计：", options:{ bold:true, color:"1E2761" } },
    { text:"随机、双盲、安慰剂对照", options:{ color:"2E2A45" } },
  ], { x:0.6, y:2.5, w:8.8, h:0.45, fontFace:"PingFang SC", fontSize:14, align:"center" });
  ```
- Set `margin: 0` on a text box when you need the text to align precisely with a
  shape edge (text boxes have internal padding by default).
- Bullets: `bullet: true` — never a literal `"•"` (creates double bullets).

## Rendering note

LibreOffice's PDF preview sometimes substitutes a serif/kai font for CJK even
when the .pptx correctly specifies PingFang SC. That is a *preview-only*
artifact — PowerPoint/Keynote on macOS will show the specified font. Judge font
choice from the actual app, not the QA JPG, but still use the JPGs for layout,
overflow, alignment, and contrast checks.
