---
name: image-to-pptx
description: >-
  Faithfully recreate a source image — especially a diagram, flowchart, org
  chart, poster, or infographic (Chinese research/medical slides are a common
  case) — as an editable PowerPoint (.pptx) using pptxgenjs. Use this whenever
  the user shares an image and asks to "把这张图做成 PPT", "turn this diagram/chart
  into slides", "recreate this figure as a PPT", "照着这张图做 PPT", reproduce a
  flowchart's full layout on one slide, or otherwise convert a picture of a
  layout into native, editable PowerPoint shapes and text (not just an embedded
  screenshot). Also triggers when redrawing/cleaning up a messy or low-res
  diagram into a polished deck. Do NOT use for editing an existing .pptx that
  already exists, or for plain text-to-slides with no source image.
---

# Image → Editable PPTX

Recreate what a source image *shows* as **native, editable** PowerPoint objects
(boxes, arrows, text runs) — never by just pasting the image onto a slide. The
goal is a deck the user can open and edit, that reads as a faithful, cleaned-up
version of the original.

This is the workflow that reliably produces good results. Follow it in order.

## 1. Read the image carefully and transcribe ALL content

Look at the image and write down every text block, its hierarchy (title →
section → body → caption), and the connections between blocks (arrows, columns,
nesting). Missing content is the most common failure — capture footers, tiny
captions, side labels, and connector labels too.

**Fix OCR / source garbling.** Photographed or AI-generated diagrams often
contain nonsense characters. Repair them into the sentence the author clearly
meant, using domain context. Examples seen in practice:

- `共分必模块` → `共享模块`   ·   `表路` → `通路`   ·   `控控` → `调控`
- `广频羋抑制` → `广谱抑制`   ·   `单一主导主导时` → `单一激素主导时`

When a word is ambiguous, prefer the reading that is medically/technically
coherent. If genuinely unsure and it matters, ask the user.

## 2. Decide layout: match the source's structure

- **A single tall flowchart / poster** → reproduce it on **one slide** using a
  custom portrait layout so the vertical top-to-bottom flow is preserved. This
  is usually what "照着原图完整样式做" means.
- **A wide dashboard or multi-topic figure** → one landscape slide, or split
  into a few slides only if the user wants a presentation rather than a replica.
- When the user says "完整样式" / "按原图" / "same layout", **replicate the
  original geometry** (columns, arrow directions, which box is highlighted) — do
  not redesign it into a generic bullet deck.
- **Classify every connector before coding.** Each fan-out/fan-in has a specific
  shape — top-down **fork/comb** (one label → many box tops), bottom-up **merge**
  (many box bottoms → one label), or a sibling **bridge** arrow — and they are
  NOT interchangeable. Look at where the arrowheads land to tell them apart.
  Styles can differ per row and per slide, so classify each transition from the
  image rather than defaulting to a single center arrow. Also honor box
  **grouping** (rows sharing one bordered box vs. separate boxes) and any
  outer-frame **zones with straddling section pills / divider rules**. See
  `references/pptxgenjs-pitfalls.md` → "Connector TOPOLOGY" and "Framed sections"
  for ready-to-paste `fork()` / `merge()` / bridge / divider helpers.

For a portrait flowchart, match the source aspect ratio:
```js
p.defineLayout({ name: "FLOW", width: 10, height: 16.4 }); // h ≈ 10 * origH/origW
p.layout = "FLOW";
```

## 3. Pull the palette from the image

Sample the source's own colors so the result feels like the same document. The
example flowchart used lavender-purple + deep navy:

```
NAVY 1B2A6B · PURPLE 7A4AA6 · PURPLE_D 5B2C87 · DARK ④ 5A2A86
LAV1 F3EEFA · LAV2 EBE0F5 · border D9C9EC · gray connectors 9A94AE
ink 2E2A45 · muted 6E6A85 · accent C4A6E8
```
One color should dominate; reserve a darker/heavier fill for the block the
source emphasizes (e.g. the "④ 中介效应验证" and the conclusion bar were dark).

## 4. Build with pptxgenjs

Use `scripts/build_template.js` as a starting point — it renders a full
vertical medical-research flowchart (title → boxed sections → two-column
cards → numbered validation grid → dark conclusion bar → outputs → footer)
with gray connector arrows, and is heavily commented. Copy it, swap in the
transcribed content, layout, and palette.

Key techniques it demonstrates:
- **CJK font**: set `fontFace: "PingFang SC"` on every text object (macOS). This
  is essential or Chinese renders in a fallback font.
- **Connector arrows**: a zero-width LINE with an arrowhead —
  `addShape(shapes.LINE, { x, y, w:0, h:0.42, line:{ color, width:2.5, endArrowType:"triangle" }})`.
- **Section header sitting on a box's top edge**: draw the box, then two short
  flanking `LINE`s, then centered bold title text.
- **Rich text runs** for mixed emphasis in one line (bold label + normal body +
  colored accent) via the array form of `addText`.
- **Highlighted card**: give one card a dark fill with light text to mirror the
  source's emphasis.

See `references/pptxgenjs-pitfalls.md` before writing code — a few mistakes
silently corrupt the file (e.g. `#` in hex, 8-char hex opacity, reusing a
shadow object across shapes).

## 5. Always run a visual QA pass — assume there are bugs

Render to images and actually look. Your first render is rarely perfect.

```bash
scripts/render_qa.sh "/path/to/output.pptx"   # → qa-N.jpg next to the pptx
```

Then Read the `qa-*.jpg` files and hunt for: missing/garbled text, overflow past
box edges, overlapping elements, low-contrast text (light text on light fill, or
dark caption on a dark card), uneven gaps, arrows not lining up between boxes.
Fix, re-render, re-check until a full pass is clean. Only then declare done.

Also verify text content:
```bash
python3 -m markitdown "/path/to/output.pptx"   # confirm nothing is missing
```

## 6. Save where asked and clean up

Write the `.pptx` to the path/desktop the user requested (use the descriptive
Chinese title as the filename when the source is Chinese). Remove temp build
scripts and QA artifacts (`*.pdf`, `qa-*.jpg`, the throwaway build `.js`) unless
the user wants them kept.

## Setup / dependencies

```bash
npm install pptxgenjs          # in the dir you run node from (e.g. ~)
# LibreOffice `soffice` + Poppler `pdftoppm` for QA rendering (usually present)
pip install "markitdown[pptx]" # optional, for text-content QA
```
Node resolves `node_modules` by walking **up** from the build script's
location, so place the script somewhere under the directory that has
`node_modules` (e.g. keep the build `.js` inside or below `~` after
`npm install pptxgenjs` in `~`). Write the output to an absolute path.
