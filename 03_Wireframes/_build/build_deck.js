// Builds wireframes.pptx in the same Class 4/5 "Usability and Design Principles" style.
// To rebuild: cd 03_Wireframes/_build, run "npm install" once, then "npm run build".
// The Harbourside wireframe pictures are drawn from wireframe.js (the same layout as the .drawio file),
// so edit wireframe.js and re-run to update both.
const pptxgen = require("pptxgenjs");
const path = require("path");
const WF = require("./wireframe");

const REPO = path.join(__dirname, ".."); // the 03_Wireframes folder
const SCR = __dirname;
const OUT = path.join(REPO, "wireframes.pptx");

const FONT = "Aptos";
const MONO = "Consolas";
const C = {
  dark: "262626", light: "F2F2F2", black: "000000", white: "FFFFFF",
  orange: "E97132", purple: "A02B93", blue: "0F9ED5", navy: "12304A",
  teal: "1B5E82", grey: "595959", mid: "7F7F7F", actBg: "E1E8EB",
  green: "2E7D32", red: "C62828", lmBlue: "1E88E5",
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "Class 6: Wireframes";
pres.author = "Jamie Symonds";

const W = 13.333, H = 7.5;
const PANEL_W = 4.1;
const CX = 5.2, CW = 7.55; // content column on panel slides

// ---------- helpers (same as the Class 5 deck) ----------
function txt(slide, text, opts) {
  slide.addText(text, Object.assign({ isTextBox: true, fontFace: FONT, margin: 0 }, opts));
}

function panelSlide(title, notes) {
  const s = pres.addSlide();
  s.background = { color: C.light };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: PANEL_W, h: H, fill: { color: C.dark }, line: { color: C.dark } });
  txt(s, "Usability and Design Principles", { x: 0.95, y: 0.6, w: 2.8, h: 1.4, fontSize: 22, color: C.light, valign: "top" });
  s.addShape(pres.shapes.RECTANGLE, { x: CX, y: 0.72, w: 0.5, h: 0.05, fill: { color: C.black }, line: { color: C.black } });
  txt(s, title, { x: CX, y: 0.9, w: CW, h: 0.95, fontSize: 28, bold: true, color: C.black, valign: "top" });
  if (notes) s.addNotes(notes);
  return s;
}

function labelBullets(s, items, { y = 1.95, h = 5.1, size = 17, sub = 15, x = CX, w = CW } = {}) {
  const runs = [];
  items.forEach((it, i) => {
    runs.push({ text: it.label, options: { bold: true, bullet: true, fontSize: size, paraSpaceBefore: i ? 8 : 0, breakLine: true } });
    (Array.isArray(it.text) ? it.text : [it.text]).forEach((t) => {
      runs.push({ text: t, options: { bullet: true, indentLevel: 1, fontSize: sub, paraSpaceBefore: 3, breakLine: true } });
    });
  });
  delete runs[runs.length - 1].options.breakLine;
  txt(s, runs, { x, y, w, h, color: C.black, valign: "top" });
}

// Rich bullet: array of [text, style] segments per bullet (style: b = bold, c = code, i = italic)
function richBullets(s, bullets, { x, y, w, h, size = 15, color = C.dark, space = 9, fontFace = FONT, bullet = true, valign = "top" }) {
  const runs = [];
  bullets.forEach((segs, bi) => {
    segs.forEach((seg, si) => {
      const [t, style] = typeof seg === "string" ? [seg, ""] : seg;
      const o = { fontSize: size, fontFace: style.includes("c") ? MONO : fontFace, bold: style.includes("b"), italic: style.includes("i") };
      if (style.includes("c")) o.color = "8A2B0B";
      if (si === 0) runs.push({ text: "​", options: { fontSize: size, fontFace, color, bullet, paraSpaceBefore: bi ? space : 0 } });
      if (si === segs.length - 1 && bi < bullets.length - 1) o.breakLine = true;
      runs.push({ text: t, options: o });
    });
  });
  txt(s, runs, { x, y, w, h, color, valign });
}

function dividerSlide(title, sub, notes) {
  const s = pres.addSlide();
  s.background = { path: path.join(SCR, "grad_full.png") };
  txt(s, title, { x: 3.2, y: 0.7, w: 9.2, h: 4.6, fontSize: 60, color: C.white, align: "right", valign: "middle" });
  if (sub) txt(s, sub, { x: 1.55, y: 6.55, w: 11, h: 0.45, fontSize: 18, color: C.white });
  s.addShape(pres.shapes.LINE, { x: 1.0, y: 4.2, w: 0, h: 3.3, line: { color: C.white, width: 2 } });
  [["+", 4.15, 0.55, 20], ["•", 4.55, 0.8, 20], ["○", 4.1, 1.0, 16], ["+", 12.6, 6.3, 20], ["•", 12.35, 6.85, 20], ["○", 13.0, 6.85, 14]]
    .forEach(([t, x, y, f]) => txt(s, t, { x, y, w: 0.4, h: 0.4, fontSize: f, color: C.white, align: "center", valign: "middle" }));
  if (notes) s.addNotes(notes);
  return s;
}

function activitySlide(title, notes) {
  const s = pres.addSlide();
  s.background = { color: C.actBg };
  txt(s, title, { x: 0.75, y: 0.5, w: 11.8, h: 0.7, fontSize: 28, color: C.black });
  if (notes) s.addNotes(notes);
  return s;
}

function card(s, { x, y, w, h, fill = C.white, border = "D0D7DB", bw = 0.75 }) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: fill }, line: { color: border, width: bw },
    shadow: { type: "outer", blur: 6, offset: 2, angle: 90, color: "000000", opacity: 0.12 } });
}

function numCircle(s, n, x, y, color, d = 0.5) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color }, line: { color } });
  txt(s, String(n), { x, y, w: d, h: d, fontSize: d > 0.45 ? 18 : 13, bold: true, color: C.white, align: "center", valign: "middle" });
}

function pill(s, text, x, y, color, w) {
  w = w || 0.3 + text.length * 0.1;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.34, rectRadius: 0.17, fill: { color }, line: { color } });
  txt(s, text, { x, y, w, h: 0.34, fontSize: 11, bold: true, color: C.white, align: "center", valign: "middle", charSpacing: 1 });
  return w;
}

function table(s, rows, { x, y, w, colW, size = 14, rowH = 0.5 }) {
  const data = rows.map((r, ri) => r.map((cell) => ({ text: cell, options: {
    fontFace: FONT, fontSize: size, color: ri === 0 ? C.white : C.black, bold: ri === 0, valign: "middle",
    fill: { color: ri === 0 ? C.dark : (ri % 2 ? C.white : "E4E4E4") }, margin: [4, 8, 4, 8] } })));
  s.addTable(data, { x, y, w, colW, rowH, border: { type: "solid", pt: 0.5, color: "C8C8C8" } });
}

// ---------- wireframe renderer: draws wireframe.js cells as native shapes ----------
const hex = (c) => (c || "").replace("#", "").toUpperCase();
function parseStyle(style) {
  const o = {};
  style.split(";").filter(Boolean).forEach((kv) => {
    const i = kv.indexOf("=");
    if (i < 0) o[kv] = true; else o[kv.slice(0, i)] = kv.slice(i + 1);
  });
  return o;
}

// crop = px region {x, y, w, h}; box = slide region in inches. Returns a px → inches mapper.
function drawWire(s, cells, box, { crop, show = () => true, align = "center" }) {
  const k = Math.min(box.w / crop.w, box.h / crop.h);
  const ox = box.x + (box.w - crop.w * k) / 2;
  const oy = align === "top" ? box.y : box.y + (box.h - crop.h * k) / 2;
  const X = (px) => ox + (px - crop.x) * k, Y = (py) => oy + (py - crop.y) * k;
  const pt = (px) => Math.max(1, Math.round(px * k * 72 * 2) / 2);
  const byId = {};
  cells.forEach((c) => { byId[c.id] = c; });

  cells.filter(show).forEach((c) => {
    const st = parseStyle(c.style);
    const stroke = hex(st.strokeColor);
    const lw = Math.max(0.5, (+st.strokeWidth || 1) * k * 72);
    const line = !st.strokeColor || st.strokeColor === "none" ? { type: "none" }
      : { color: stroke, width: lw, dashType: st.dashed === "1" ? "dash" : "solid" };

    if (c.edge) {
      const a = byId[c.source], b = byId[c.target];
      const sx = a.x + a.w, sy = a.y + a.h / 2, tx = b.x, ty = b.y + b.h * 0.8, mx = (sx + tx) / 2;
      const seg = (x1, y1, x2, y2, last) => s.addShape(pres.shapes.LINE, {
        x: X(Math.min(x1, x2)), y: Y(Math.min(y1, y2)), w: Math.abs(x2 - x1) * k, h: Math.abs(y2 - y1) * k,
        line: Object.assign({ color: stroke, width: lw, dashType: "dash" }, last ? { endArrowType: "triangle" } : {}) });
      seg(sx, sy, mx, sy); seg(mx, ty, mx, sy); seg(mx, ty, tx, ty, true);
      return;
    }

    const geo = { x: X(c.x), y: Y(c.y), w: c.w * k, h: c.h * k };
    const fill = !st.fillColor || st.fillColor === "none" ? { type: "none" } : { color: hex(st.fillColor) };
    if (!st.text) {
      if (st.shape === "mxgraph.mockup.graphics.simpleIcon") {
        s.addShape(pres.shapes.RECTANGLE, Object.assign({}, geo, { fill, line }));
        s.addShape(pres.shapes.LINE, Object.assign({}, geo, { line: { color: stroke, width: lw } }));
        s.addShape(pres.shapes.LINE, Object.assign({}, geo, { flipV: true, line: { color: stroke, width: lw } }));
      } else if (st.ellipse) {
        s.addShape(pres.shapes.OVAL, Object.assign({}, geo, { fill, line }));
      } else if (st.rounded === "1") {
        const r = Math.min(c.w, c.h) * k * (+st.arcSize || 15) / 100;
        s.addShape(pres.shapes.ROUNDED_RECTANGLE, Object.assign({}, geo, { fill, line, rectRadius: r }));
      } else {
        s.addShape(pres.shapes.RECTANGLE, Object.assign({}, geo, { fill, line }));
      }
    }
    if (c.value) {
      const fs = +st.fontStyle || 0;
      const sp = st.spacing !== undefined ? +st.spacing : 2;
      const ml = st.spacingLeft ? +st.spacingLeft : sp;
      txt(s, c.value, Object.assign({}, geo, {
        fontFace: st.fontFamily || "Arial", fontSize: pt(+st.fontSize || 12), color: hex(st.fontColor) || "333333",
        bold: !!(fs & 1), italic: !!(fs & 2), underline: fs & 4 ? { style: "sng" } : undefined,
        align: st.align || "center", valign: st.verticalAlign === "top" ? "top" : "middle",
        margin: [pt(sp), pt(sp), pt(sp), pt(ml)], lineSpacingMultiple: 0.95,
      }));
    }
  });
  return { X, Y, k };
}

const DESK_CROP = { x: 30, y: 60, w: 1220, h: 1080 };     // browser frame only
const DESK_NOTES_CROP = { x: 30, y: 60, w: 1450, h: 1080 }; // frame + yellow notes
const MOB_CROP = { x: 15, y: 60, w: 1130, h: 1170 };       // both phones, no notes
const noMeta = (c) => c.step !== "meta" && c.step !== "note";

function highlight(s, map, r) {
  const pad = 10;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: map.X(r.x - pad), y: map.Y(r.y - pad), w: (r.w + 2 * pad) * map.k, h: (r.h + 2 * pad) * map.k,
    rectRadius: 0.08, fill: { type: "none" }, line: { color: C.orange, width: 3.5 } });
}

// ---------- guided-example step slide ----------
function stepSlide({ n, title, story, region, bullets, ask, notes }) {
  const s = pres.addSlide();
  s.background = { color: C.light };
  const map = drawWire(s, WF.desktop, { x: 0.35, y: 0.5, w: 7.3, h: 6.5 }, { crop: DESK_CROP, show: (c) => typeof c.step === "number" && c.step <= n });
  highlight(s, map, region);
  const RX = 8.0, RW = W - RX - 0.45;
  txt(s, `GUIDED EXAMPLE · STEP ${n}`, { x: RX, y: 0.5, w: RW, h: 0.35, fontSize: 13, color: C.grey, bold: true, charSpacing: 1 });
  txt(s, title, { x: RX, y: 0.85, w: RW, h: 0.65, fontSize: 28, color: C.black, valign: "top" });
  let by = 1.6;
  if (story) { pill(s, story, RX, 1.6, C.orange); by = 2.15; }
  richBullets(s, bullets, { x: RX, y: by, w: RW, h: 5.15 - by, size: 14.5, space: 8 });
  card(s, { x: RX, y: 5.3, w: RW, h: 1.65 });
  numCircle(s, "?", RX + 0.2, 5.5, C.purple, 0.45);
  txt(s, "Ask the class", { x: RX + 0.8, y: 5.5, w: RW - 1, h: 0.45, fontSize: 15, bold: true, color: C.black, valign: "middle" });
  txt(s, ask, { x: RX + 0.25, y: 6.05, w: RW - 0.45, h: 0.8, fontSize: 14, italic: true, color: C.dark, valign: "top" });
  if (notes) s.addNotes(notes);
  return s;
}

// =====================================================================
// 1. Title
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  txt(s, "WEBD 1000", { x: 0.3, y: 0.3, w: 3, h: 0.4, fontSize: 16, color: C.black });
  s.addShape(pres.shapes.OVAL, { x: 0.55, y: 1.1, w: 5.6, h: 5.6, fill: { color: C.orange }, line: { color: C.orange } });
  txt(s, "Usability and Design Principles", { x: 1.45, y: 2.4, w: 4.2, h: 2.9, fontSize: 46, color: C.white, valign: "middle" });
  s.addShape(pres.shapes.OVAL, { x: 1.0, y: 5.75, w: 0.62, h: 0.62, fill: { color: C.purple }, line: { color: C.purple } });
  s.addShape(pres.shapes.ARC, { x: 9.6, y: 0.75, w: 3.3, h: 3.3, angleRange: [200, 20], line: { color: C.blue, width: 7, dashType: "dash" } });
  txt(s, [
    { text: "Class 6:", options: { bullet: true, fontSize: 30, breakLine: true } },
    { text: "Wireframes: planning a page before we build it.", options: { bullet: { code: "25CB" }, indentLevel: 1, fontSize: 26, paraSpaceBefore: 6 } },
  ], { x: 6.85, y: 1.9, w: 5.6, h: 2.2, color: C.black, valign: "top" });
  s.addNotes("Welcome back. So far we've written user stories (what users need to do) and semantic HTML (how the page is structured). Today sits between the two: before we write any HTML, we plan the page with a wireframe.\n\nWe'll use draw.io (app.diagrams.net). It's free, runs in the browser, and needs no account.");
}

// 2. Attendance
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  txt(s, "Attendance", { x: 1.1, y: 0.75, w: 8, h: 0.9, fontSize: 44, color: C.black });
  s.addNotes("Take attendance.");
}

// 3. Overview
{
  const s = panelSlide("Class 6: Class Overview",
    "Agenda for today. Roughly: 10 min on what wireframes are, 35 min guided example (build it live in draw.io while students follow along), 50 min for their own wireframe plus Teams posting and feedback, 5 min wrap-up.\n\nStudents need a browser only. draw.io works without signing in.");
  txt(s, [
    { text: "Title: ", options: { bold: true, bullet: true, fontSize: 22 } },
    { text: "Wireframes with draw.io", options: { fontSize: 22, breakLine: true } },
    { text: " ", options: { fontSize: 10, breakLine: true } },
    { text: "Agenda:", options: { bold: true, bullet: true, fontSize: 22, breakLine: true } },
    { text: "What a wireframe is, and why we make one first", options: { bullet: true, indentLevel: 1, fontSize: 18, paraSpaceBefore: 4, breakLine: true } },
    { text: "The low-fi toolkit, and turning user stories into boxes", options: { bullet: true, indentLevel: 1, fontSize: 18, paraSpaceBefore: 4, breakLine: true } },
    { text: "Guided example: Harbourside Fish & Chips (desktop + mobile)", options: { bullet: true, indentLevel: 1, fontSize: 18, paraSpaceBefore: 4, breakLine: true } },
    { text: "Your turn: wireframe a page of your own", options: { bullet: true, indentLevel: 1, fontSize: 18, paraSpaceBefore: 4, breakLine: true } },
    { text: "Share on Teams and give feedback", options: { bullet: true, indentLevel: 1, fontSize: 18, paraSpaceBefore: 4 } },
  ], { x: CX, y: 1.95, w: CW, h: 4.8, color: C.black, valign: "top" });
}

// 4. What is a wireframe? (fidelity spectrum)
{
  const s = panelSlide("What Is a Wireframe?",
    "A wireframe is like a floor plan for a house: it shows where the rooms go, not the paint colours.\n\nWalk through the spectrum. Sketches are fastest (paper, whiteboard). Wireframes are what we do today: tidy boxes, still no colour. Mockups add the real look (colours, fonts, images). Prototypes are clickable. Each step costs more time to change, so we make the big layout decisions early, while they're cheap.\n\nIn industry, tools like Figma are common for mockups and prototypes. draw.io is great for wireframes because it's free and quick to learn.");
  richBullets(s, [["A wireframe is a simple plan of a web page. It shows ", ["what goes where", "b"], ", and nothing about how it looks."]],
    { x: CX, y: 1.85, w: CW, h: 0.8, size: 17, bullet: false, color: C.black });
  const cw = 1.7, gap = (CW - 4 * cw) / 3, y = 2.85, ch = 2.75;
  const stages = [
    ["Sketch", "Paper or whiteboard. Fast and messy."],
    ["Wireframe", "Tidy boxes and labels. No colour."],
    ["Mockup", "Real colours, fonts and images."],
    ["Prototype", "Clickable. Feels like the real site."],
  ];
  stages.forEach(([name, d], i) => {
    const x = CX + i * (cw + gap);
    const today = i === 1;
    card(s, { x, y, w: cw, h: ch, border: today ? C.orange : "D0D7DB", bw: today ? 2.5 : 0.75 });
    // mini illustration (0.15" inset, 1.4 x 0.95)
    const ix = x + 0.15, iy = y + 0.2, iw = cw - 0.3, ih = 0.95;
    if (i === 0) {
      s.addShape(pres.shapes.RECTANGLE, { x: ix, y: iy, w: iw, h: ih, fill: { type: "none" }, line: { color: C.mid, width: 1.25, dashType: "sysDash" } });
      s.addShape(pres.shapes.LINE, { x: ix + 0.1, y: iy + 0.2, w: iw - 0.2, h: 0.04, line: { color: C.mid, width: 1.25 } });
      s.addShape(pres.shapes.OVAL, { x: ix + 0.12, y: iy + 0.4, w: 0.45, h: 0.4, fill: { type: "none" }, line: { color: C.mid, width: 1.25 } });
      s.addShape(pres.shapes.LINE, { x: ix + 0.7, y: iy + 0.5, w: 0.55, h: 0.03, line: { color: C.mid, width: 1.25 } });
      s.addShape(pres.shapes.LINE, { x: ix + 0.7, y: iy + 0.68, w: 0.4, h: 0.02, flipV: true, line: { color: C.mid, width: 1.25 } });
    } else {
      const pal = i === 1 ? ["FFFFFF", "999999", "DDDDDD", "555555"] : ["FFF7F2", C.teal, "F2C6AC", C.orange];
      s.addShape(pres.shapes.RECTANGLE, { x: ix, y: iy, w: iw, h: ih, fill: { color: pal[0] }, line: { color: "777777", width: 1 } });
      s.addShape(pres.shapes.RECTANGLE, { x: ix, y: iy, w: iw, h: 0.18, fill: { color: i === 1 ? "EEEEEE" : C.teal }, line: { color: "777777", width: 1 } });
      s.addShape(pres.shapes.RECTANGLE, { x: ix + 0.1, y: iy + 0.3, w: 0.5, h: 0.5, fill: { color: i === 1 ? "FFFFFF" : "F2C6AC" }, line: { color: pal[1], width: 0.75 } });
      if (i === 1) {
        s.addShape(pres.shapes.LINE, { x: ix + 0.1, y: iy + 0.3, w: 0.5, h: 0.5, line: { color: "999999", width: 0.75 } });
        s.addShape(pres.shapes.LINE, { x: ix + 0.1, y: iy + 0.3, w: 0.5, h: 0.5, flipV: true, line: { color: "999999", width: 0.75 } });
      }
      s.addShape(pres.shapes.RECTANGLE, { x: ix + 0.7, y: iy + 0.33, w: 0.55, h: 0.06, fill: { color: pal[2] }, line: { type: "none" } });
      s.addShape(pres.shapes.RECTANGLE, { x: ix + 0.7, y: iy + 0.46, w: 0.4, h: 0.06, fill: { color: pal[2] }, line: { type: "none" } });
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: ix + 0.7, y: iy + 0.62, w: 0.5, h: 0.18, rectRadius: 0.04, fill: { color: pal[3] }, line: { type: "none" } });
      if (i === 3) txt(s, "👆", { x: ix + 1.02, y: iy + 0.58, w: 0.35, h: 0.35, fontSize: 16, fontFace: "Segoe UI Emoji" });
    }
    txt(s, name, { x: x + 0.15, y: y + 1.3, w: cw - 0.3, h: 0.4, fontSize: 16, bold: true, color: C.black });
    txt(s, d, { x: x + 0.15, y: y + 1.72, w: cw - 0.3, h: 0.9, fontSize: 12.5, color: C.grey, valign: "top" });
    if (today) pill(s, "TODAY", x + cw / 2 - 0.45, y - 0.17, C.orange, 0.9);
  });
  s.addShape(pres.shapes.LINE, { x: CX + 0.1, y: 5.95, w: CW - 0.2, h: 0, line: { color: C.mid, width: 1.5, beginArrowType: "triangle", endArrowType: "triangle" } });
  txt(s, "Low fidelity: quick to change", { x: CX, y: 6.1, w: 3.5, h: 0.35, fontSize: 12.5, color: C.grey });
  txt(s, "High fidelity: slow to change", { x: CX + CW - 3.5, y: 6.1, w: 3.5, h: 0.35, fontSize: 12.5, color: C.grey, align: "right" });
}

// 5. Why wireframe first?
{
  const s = panelSlide("Why Wireframe First?",
    "Ask the class before revealing: why not just start coding?\n\nThe main point is cost of change. Moving a box in draw.io takes seconds. Moving a section of a finished, styled, coded page can take hours, and may break other things.\n\nPoint 4 links back to Class 5: every region we draw today becomes a landmark element when we build it.");
  txt(s, "Before we write any HTML, a wireframe lets us check the plan.", { x: CX, y: 1.85, w: CW, h: 0.5, fontSize: 16, color: C.dark });
  const cards = [
    ["Cheap to change", "Moving a box takes seconds. Moving a finished, coded page takes hours.", C.orange],
    ["Layout, not looks", "With no colours or fonts to discuss, feedback is about what goes where.", C.teal],
    ["Easy to show a client", "They can point at a box and say \"that's not what I meant\" before any code exists.", C.purple],
    ["A plan for the HTML", "Each region becomes a landmark: <header>, <nav>, <main>, <footer>.", C.blue],
  ];
  const cw = 3.65, ch = 2.1;
  cards.forEach(([t, d, col], i) => {
    const x = CX + (i % 2) * (cw + 0.25), y = 2.55 + Math.floor(i / 2) * (ch + 0.25);
    card(s, { x, y, w: cw, h: ch });
    numCircle(s, i + 1, x + 0.2, y + 0.22, col, 0.55);
    txt(s, t, { x: x + 0.9, y: y + 0.22, w: cw - 1.0, h: 0.55, fontSize: 18, bold: true, color: C.black, valign: "middle" });
    txt(s, d, { x: x + 0.25, y: y + 0.95, w: cw - 0.45, h: 1.0, fontSize: 14, color: C.grey, valign: "top" });
  });
}

// 6. The low-fi toolkit
{
  const s = panelSlide("The Low-Fi Toolkit",
    "These are the only building blocks students need today. Everything on the Harbourside wireframe is made from these six.\n\nIn draw.io: the X box is Mockups → Graphics → Image (or draw a rectangle with two diagonal lines). Buttons are rounded rectangles. Grey bars are thin rectangles with a light grey fill and no border.\n\nThe blue landmark boxes are notes to ourselves about the HTML. They wouldn't be shown to a client as part of the design.");
  const cw = 2.35, ch = 1.95, gap = (CW - 3 * cw) / 2;
  const tiles = [
    ["Image", "Logo or photo. Never a real picture yet."],
    ["Heading", "Big, bold text for real headings."],
    ["Body text", "Grey bars instead of real paragraphs."],
    ["Main button", "The main action. Filled, and usually only one per screen."],
    ["Other buttons", "Other choices. Outlined."],
    ["Landmark label", "A note for the HTML, not part of the design."],
  ];
  tiles.forEach(([t, d], i) => {
    const x = CX + (i % 3) * (cw + gap), y = 1.9 + Math.floor(i / 3) * (ch + 0.25);
    card(s, { x, y, w: cw, h: ch });
    const ix = x + 0.25, iy = y + 0.2, iw = cw - 0.5, ih = 0.7;
    if (i === 0) {
      const g = { x: ix + 0.35, y: iy, w: iw - 0.7, h: ih };
      s.addShape(pres.shapes.RECTANGLE, Object.assign({}, g, { fill: { color: C.white }, line: { color: "999999", width: 1 } }));
      s.addShape(pres.shapes.LINE, Object.assign({}, g, { line: { color: "999999", width: 1 } }));
      s.addShape(pres.shapes.LINE, Object.assign({}, g, { flipV: true, line: { color: "999999", width: 1 } }));
    } else if (i === 1) {
      txt(s, "Fish & Chips", { x: ix, y: iy, w: iw, h: ih, fontSize: 20, bold: true, color: "333333", fontFace: "Arial", valign: "middle" });
    } else if (i === 2) {
      [1, 0.85, 0.6].forEach((f, j) => s.addShape(pres.shapes.RECTANGLE, { x: ix, y: iy + 0.12 + j * 0.2, w: iw * f, h: 0.09, fill: { color: "DDDDDD" }, line: { type: "none" } }));
    } else if (i === 3 || i === 4) {
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: ix + 0.2, y: iy + 0.12, w: iw - 0.4, h: 0.46, rectRadius: 0.07,
        fill: { color: i === 3 ? "555555" : C.white }, line: { color: "555555", width: 1.25 } });
      txt(s, i === 3 ? "Checkout" : "Delivery", { x: ix + 0.2, y: iy + 0.12, w: iw - 0.4, h: 0.46, fontSize: 14, bold: true, fontFace: "Arial",
        color: i === 3 ? C.white : "333333", align: "center", valign: "middle" });
    } else {
      s.addShape(pres.shapes.RECTANGLE, { x: ix, y: iy + 0.15, w: iw, h: ih - 0.15, fill: { type: "none" }, line: { color: C.lmBlue, width: 1.75, dashType: "dash" } });
      s.addShape(pres.shapes.RECTANGLE, { x: ix, y: iy - 0.05, w: 0.72, h: 0.22, fill: { color: C.lmBlue }, line: { type: "none" } });
      txt(s, "<main>", { x: ix, y: iy - 0.05, w: 0.72, h: 0.22, fontSize: 10.5, bold: true, fontFace: "Courier New", color: C.white, align: "center", valign: "middle" });
    }
    txt(s, t, { x: x + 0.2, y: y + 1.0, w: cw - 0.4, h: 0.35, fontSize: 15, bold: true, color: C.black });
    txt(s, d, { x: x + 0.2, y: y + 1.35, w: cw - 0.4, h: 0.55, fontSize: 12, color: C.grey, valign: "top" });
  });
  richBullets(s, [[["Rules for today: ", "b"], "greyscale only · real words on buttons and links · no fonts, colours or photos yet."]],
    { x: CX, y: 6.35, w: CW, h: 0.5, size: 14, bullet: false, color: C.dark });
}

// 7. From user stories to wireframe
{
  const s = panelSlide("From User Stories to Boxes",
    "This connects to the requirements and user stories class. A user story says what someone needs to do. The wireframe decides WHERE on the page they do it. The landmark says what HTML element that region becomes.\n\nGood habit for today: number your user stories and put the matching number on the wireframe. If a story has no number on the page, the design is missing something. If a big box has no story, ask whether it belongs.");
  const cols = [["User story", 3.3], ["Region on the page", 2.5], ["HTML", 1.15]];
  const aw = (CW - cols.reduce((a, c) => a + c[1], 0)) / 2;
  let hx = CX;
  cols.forEach(([t, w]) => { txt(s, t.toUpperCase(), { x: hx, y: 1.9, w, h: 0.35, fontSize: 12, bold: true, color: C.grey, charSpacing: 1 }); hx += w + aw; });
  const rows = [
    [[["As a ", ""], ["customer", "b"], ", I want to ", ["see my total while I'm still browsing", "b"], " so that I don't go over budget."], "Order summary panel beside the menu", "<aside>"],
    [[["As a ", ""], ["hungry customer", "b"], ", I want to ", ["browse the menu by category", "b"], " so that I can find what I want quickly."], "Category tabs and a grid of menu cards", "<article>\nper card"],
  ];
  rows.forEach(([story, region, html], ri) => {
    const y = 2.35 + ri * 1.65, h = 1.35;
    let x = CX;
    card(s, { x, y, w: 3.3, h });
    richBullets(s, [story], { x: x + 0.2, y: y + 0.15, w: 2.9, h: h - 0.3, size: 13.5, bullet: false, color: C.dark });
    x += 3.3;
    txt(s, "→", { x, y, w: aw, h, fontSize: 22, color: C.orange, bold: true, align: "center", valign: "middle" });
    x += aw;
    card(s, { x, y, w: 2.5, h, fill: "FBE9DF", border: "F2C6AC" });
    txt(s, region, { x: x + 0.2, y, w: 2.1, h, fontSize: 14, bold: true, color: C.black, valign: "middle" });
    x += 2.5;
    txt(s, "→", { x, y, w: aw, h, fontSize: 22, color: C.orange, bold: true, align: "center", valign: "middle" });
    x += aw;
    card(s, { x, y, w: 1.15, h, fill: "E3F0FC", border: C.lmBlue });
    txt(s, html, { x, y, w: 1.15, h, fontSize: 12.5, bold: true, fontFace: MONO, color: "0D47A1", align: "center", valign: "middle" });
  });
  card(s, { x: CX, y: 5.8, w: CW, h: 0.95, fill: C.dark, border: C.dark });
  txt(s, "Every story needs a place on the page. Number your stories and put the numbers on your wireframe.",
    { x: CX + 0.3, y: 5.8, w: CW - 0.6, h: 0.95, fontSize: 15, color: C.white, valign: "middle" });
}

// 8. Divider: guided example
dividerSlide("Guided Example", "Harbourside Fish & Chips   ·   online ordering page   ·   desktop + mobile",
  "Guided example, about 35 minutes. Build it live in draw.io from a blank file while students follow along on their own laptops. Go slowly on Step 0 and Step 1; after that students usually pick it up.\n\nThe finished version is in 03_Wireframes/harbourside_ordering_wireframe.drawio if you need to jump ahead or someone falls behind.");

// 9. The client and user stories
{
  const s = panelSlide("The Client: Harbourside Fish & Chips",
    "Put these stories on the board (or leave this slide up) before opening draw.io. Keep asking 'which story is this part of the page for?' as you build.\n\nStory 5 is really a non-functional requirement written as a story (it's about how the site works on a phone, not a feature). That's the one that drives the mobile layout in Step 6.");
  txt(s, "A small takeout on the Yarmouth waterfront. They want customers to order online, from a laptop or a phone.",
    { x: CX, y: 1.85, w: CW, h: 0.6, fontSize: 15, color: C.dark });
  const stories = [
    ["hungry customer", "browse the menu by category", "I can find what I want quickly"],
    ["customer", "choose pickup or delivery and a time", "my food is ready when I get there"],
    ["customer", "see my order and total while I'm still browsing", "I don't go over budget"],
    ["customer", "add special instructions", "I get my fish the way I like it"],
    ["phone user", "order with one thumb", "I can order while I'm on the go"],
  ];
  stories.forEach(([who, what, why], i) => {
    const y = 2.6 + i * 0.86;
    card(s, { x: CX, y, w: CW, h: 0.72 });
    numCircle(s, i + 1, CX + 0.15, y + 0.13, C.orange, 0.46);
    richBullets(s, [["As a ", [who, "b"], ", I want to ", [what, "b"], " so that " + why + "."]],
      { x: CX + 0.8, y, w: CW - 1.0, h: 0.72, size: 14.5, bullet: false, color: C.dark, valign: "middle" });
  });
}

// 10. Step 0: set up draw.io
{
  const s = panelSlide("Step 0: Set Up draw.io",
    "Do this together and wait until everyone has the Mockups library showing on the left before moving on.\n\nSaving: 'Device' downloads a .drawio file; OneDrive keeps it in the cloud. Either is fine. Remind students to save often (Ctrl + S).\n\nThe shortcuts on the right come up again in Step 3 when we duplicate the menu cards.");
  const steps = [
    ["Open draw.io", "Go to app.diagrams.net → Create New Diagram → Blank Diagram."],
    ["Save it", "Name it harbourside.drawio and save it to OneDrive or your device."],
    ["Turn on Mockups", "+ More Shapes (bottom-left) → tick Mockups → Apply."],
    ["Name the page", "Double-click the page tab at the bottom and call it Desktop."],
  ];
  steps.forEach(([t, d], i) => {
    const y = 1.95 + i * 1.2;
    numCircle(s, i + 1, CX, y, C.orange, 0.48);
    txt(s, t, { x: CX + 0.7, y: y - 0.02, w: 3.8, h: 0.35, fontSize: 17, bold: true, color: C.black });
    txt(s, d, { x: CX + 0.7, y: y + 0.35, w: 3.8, h: 0.7, fontSize: 14, color: C.dark, valign: "top" });
  });
  const kx = CX + 4.75, kw = CW - 4.75;
  card(s, { x: kx, y: 1.9, w: kw, h: 4.85, fill: C.dark, border: C.dark });
  txt(s, "Handy in draw.io", { x: kx + 0.25, y: 2.1, w: kw - 0.5, h: 0.4, fontSize: 16, bold: true, color: C.white });
  const keys = [
    ["Arrange tab", "set an exact width and height"],
    ["Ctrl + G", "group shapes together"],
    ["Ctrl + D", "duplicate"],
    ["Arrange → Align", "line things up"],
    ["View → Grid", "turn the grid on or off"],
    ["Ctrl + S", "save often!"],
  ];
  keys.forEach(([k, d], i) => {
    const y = 2.65 + i * 0.66;
    txt(s, k, { x: kx + 0.25, y, w: kw - 0.5, h: 0.3, fontSize: 13.5, bold: true, color: "F2C6AC", fontFace: MONO });
    txt(s, d, { x: kx + 0.25, y: y + 0.28, w: kw - 0.5, h: 0.3, fontSize: 13, color: "D9D9D9" });
  });
}

// 11–15. Guided steps (desktop build-up)
stepSlide({
  n: 1, title: "Frame, Header and Nav", region: { x: 40, y: 70, w: 1200, h: 116 },
  bullets: [
    ["Draw a ", ["1200 × 1060", "b"], " rectangle (Arrange tab → Width / Height). This is the browser window."],
    ["Add a grey bar across the top with ", ["harbourside.ca/order", "c"], "."],
    ["Logo: ", ["Mockups → Graphics → Image", "b"], " (the X box), then the restaurant name."],
    ["Nav: Menu, ", ["Order Online", "b"], ", Hours & Location, Contact. Bold the current page."],
    ["A ", ["Cart (3)", "b"], " button on the right."],
  ],
  ask: "What landmarks will this become? Why is \"Order Online\" in bold?",
  notes: "Why 1200 wide? It's a common desktop layout width. Mobile will be 375.\n\nAnswers: the whole strip is a <header> with a <nav> inside it. Bold marks the current page, so users know where they are. It's the visual version of aria-current=\"page\" from Class 5.\n\nThe cart stays visible on every page because it's the thing an ordering customer needs most.",
});
stepSlide({
  n: 2, title: "Pickup or Delivery", story: "STORY 2", region: { x: 60, y: 202, w: 1160, h: 80 },
  bullets: [
    ["A light grey strip under the header."],
    ["Heading: ", ["Order for pickup or delivery", "b"], "."],
    ["Two buttons side by side: ", ["Pickup", "b"], " (filled = selected) and ", ["Delivery", "b"], " (outlined)."],
    ["A dropdown: ", ["Pickup time: ASAP ▾", "b"], "."],
  ],
  ask: "Why does this go at the top, before the menu?",
  notes: "Answer: the customer has to decide it before they order anything. It can change what's available, the wait time, and the fees.\n\nPoint out the filled / outlined convention from the toolkit slide: filled = selected or main action.",
});
stepSlide({
  n: 3, title: "Menu Tabs and Cards", story: "STORY 1", region: { x: 60, y: 302, w: 760, h: 640 },
  bullets: [
    ["A row of tabs: ", ["Fish", "b"], " (filled = selected), Combos, Sides, Drinks, Desserts."],
    ["Build ", ["one", "b"], " card: image, name, two grey bars, price, and an ", ["Add +", "b"], " button."],
    ["Select the card → ", ["Ctrl + G", "b"], " to group it → ", ["Ctrl + D", "b"], " to duplicate."],
    ["Arrange → Align / Distribute into a ", ["3 × 2 grid", "b"], "."],
  ],
  ask: "What HTML element is one card? Why grey bars instead of the real descriptions?",
  notes: "Answers: each card makes sense on its own and repeats, so it's an <article>. Grey bars keep it low-fi: we're deciding layout, not writing the menu.\n\nMention that repeating grids like this are what CSS Grid is for, which comes later in the course.\n\nThis is the slowest step. Make sure everyone has used group + duplicate at least once.",
});
stepSlide({
  n: 4, title: "Your Order Summary", story: "STORY 3", region: { x: 850, y: 302, w: 370, h: 440 },
  bullets: [
    ["A grey panel on the right: ", ["Your Order", "b"], "."],
    ["Three items, each with ", ["− 1 +", "b"], " and a price."],
    ["Subtotal, HST (14%), and a bold ", ["Total", "b"], "."],
    ["A big ", ["Checkout", "b"], " button. It's the only filled button in the panel."],
  ],
  ask: "Is this panel an <aside> or a <section>?",
  notes: "Let students argue it out; both can be defended. <aside> = related to the main content but separate from it (like a sidebar). <section> with a heading also works. The wireframe uses <aside>.\n\nThis panel is what makes story 3 work: the total is visible the whole time the customer is adding items. Hold on to that; it matters when we go mobile.",
});
stepSlide({
  n: 5, title: "Footer and Landmark Labels", region: { x: 40, y: 986, w: 1200, h: 144 },
  bullets: [
    ["Footer: address, phone, hours, social icons, ©."],
    ["Draw a ", ["blue dashed box", "b"], " around each region, with a small tag on top."],
    ["Tags: ", ["<header>", "c"], " ", ["<nav>", "c"], " ", ["<main>", "c"], " ", ["<article>", "c"], " ", ["<aside>", "c"], " ", ["<footer>", "c"], "."],
  ],
  ask: "Compare this to the landmarks example from Class 5. What's the same?",
  notes: "The structure is the same as 02_landmarks_example.html from Class 5. The wireframe is the plan for the HTML.\n\nTip for the tags: make one blue tag, then Ctrl + D and change the text.\n\nThe desktop page is now done. The next slide shows the finished version with the design notes.",
});

// 16. Finished desktop
{
  const s = pres.addSlide();
  s.background = { color: C.light };
  txt(s, "Finished: Desktop (1200px)", { x: 0.6, y: 0.35, w: 9, h: 0.6, fontSize: 28, color: C.black });
  drawWire(s, WF.desktop, { x: 0.4, y: 1.05, w: 12.55, h: 6.25 }, { crop: DESK_NOTES_CROP, show: (c) => c.step !== "meta" });
  s.addNotes("The finished desktop wireframe, with the yellow design notes. This is page 1 (Desktop) of harbourside_ordering_wireframe.drawio.\n\nRead the notes with the class. Each one is a design decision with a reason, which is what we want students to write in their Teams post.");
}

// 17. Step 6: going mobile
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  txt(s, "GUIDED EXAMPLE · STEP 6", { x: 0.6, y: 0.45, w: 8, h: 0.35, fontSize: 13, color: C.grey, bold: true, charSpacing: 1 });
  txt(s, "Going Mobile (375px)", { x: 0.6, y: 0.8, w: 8, h: 0.65, fontSize: 32, color: C.black });
  pill(s, "STORY 5", 9.2, 0.95, C.orange);
  richBullets(s, [["Right-click the page tab → ", ["Duplicate", "b"], ", rename it ", ["Mobile", "b"], ", and change the frame to ", ["375 wide", "b"], ". Before each change, ask: ", ["what has to change, and why?", "i"]]],
    { x: 0.6, y: 1.6, w: 12.1, h: 0.6, size: 15, bullet: false, color: C.dark });
  table(s, [
    ["Desktop", "Mobile", "Why"],
    ["Nav links across the header", "☰ button. Cart stays visible", "Not enough width. Keep the most important action on screen"],
    ["Tabs in one row", "Tabs scroll sideways", "Wrapping onto two rows pushes the menu down"],
    ["3-column card grid", "1-column list: small image left, + button right", "One thumb, one column"],
    ["Order <aside> beside the menu", "A bar that stays at the bottom: View order (3) · $33.03 →", "No room for a sidebar, but the total is still always visible (story 3)"],
    ["Order panel", "A separate Your Order screen with Special instructions and Checkout", "Room for the form (story 4). Checkout is easy to reach with a thumb"],
  ], { x: 0.6, y: 2.35, w: 12.1, colW: [3.2, 4.4, 4.5], size: 14, rowH: 0.6 });
  const py = 6.35;
  s.addShape(pres.shapes.LINE, { x: 0.6, y: py + 0.2, w: 0.6, h: 0, line: { color: "E53935", width: 2.5, dashType: "dash" } });
  txt(s, "Red dashed line at 812px = the bottom of the first screen. What does the user see before scrolling?", { x: 1.35, y: py, w: 6.2, h: 0.45, fontSize: 13.5, color: C.dark, valign: "middle" });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.9, y: py + 0.02, w: 0.42, h: 0.42, rectRadius: 0.06, fill: { color: "555555" }, line: { type: "none" } });
  txt(s, "+", { x: 7.9, y: py + 0.02, w: 0.42, h: 0.42, fontSize: 18, bold: true, color: C.white, align: "center", valign: "middle" });
  txt(s, "Tap targets at least 44 × 44px, so they're easy to hit with a thumb.", { x: 8.5, y: py, w: 4.3, h: 0.45, fontSize: 13.5, color: C.dark, valign: "middle" });
  s.addNotes("Ask 'what has to change?' BEFORE revealing each row. Students usually get the nav and the grid themselves. The order summary is the interesting one: there's no room for a sidebar, but story 3 still says the total must be visible while browsing. The answer is the bar that stays at the bottom of the screen.\n\nIn draw.io: you can also start from Mockups → Containers for a phone frame instead of a plain rectangle.\n\n375px is a common phone width in CSS pixels. 812px tall is a typical first screen. 44×44px is the tap-target size Apple recommends, and WCAG 2.1's AAA target size criterion (2.5.5) uses the same number.");
}

// 18. Finished mobile
{
  const s = pres.addSlide();
  s.background = { color: C.light };
  const map = drawWire(s, WF.mobile, { x: 0.35, y: 0.3, w: 6.9, h: 6.9 }, { crop: MOB_CROP, show: noMeta });
  const badges = [[452, 122], [472, 342], [452, 484], [452, 790], [722, 504], [722, 830]];
  const bd = 0.36;
  badges.forEach(([px, py], i) => numCircle(s, i + 1, map.X(px) - bd / 2, map.Y(py) - bd / 2, C.orange, bd));
  const RX = 7.55, RW = W - RX - 0.45;
  txt(s, "Finished: Mobile (375px)", { x: RX, y: 0.45, w: RW, h: 0.6, fontSize: 28, color: C.black });
  const notes = [
    ["Nav moves behind ☰", "Cart stays visible: it's what users need most here."],
    ["Tabs scroll sideways", "Instead of wrapping onto a second row."],
    ["Grid → 1 column", "+ buttons are 48 × 40, easy to tap with a thumb."],
    ["Sidebar → bottom bar", "The order bar stays on screen. Tap it to open Your Order."],
    ["A new screen", "A real form with a <label> for Special instructions."],
    ["Checkout at the bottom", "Easy to reach with a thumb."],
  ];
  notes.forEach(([t, d], i) => {
    const y = 1.3 + i * 0.95;
    numCircle(s, i + 1, RX, y, C.orange, 0.42);
    txt(s, t, { x: RX + 0.6, y: y - 0.03, w: RW - 0.6, h: 0.35, fontSize: 16, bold: true, color: C.black });
    txt(s, d, { x: RX + 0.6, y: y + 0.32, w: RW - 0.6, h: 0.5, fontSize: 13.5, color: C.dark, valign: "top" });
  });
  s.addNotes("The finished mobile wireframe (page 2 of the .drawio file). The left phone shows the full length of the menu page; the red line is where the first screen ends. The right phone is the Your Order screen that opens when you tap the bar.\n\nNote 5 ties back to Class 5 Example 10 (accessible forms). Note 3: 44px is the usual minimum tap target.\n\nAsk: 'What's the first thing a phone user sees? Is it the right thing?' (The pickup/delivery choice and the first few menu items, plus the order bar.)");
}

// 19. Step 7: export and post
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  txt(s, "GUIDED EXAMPLE · STEP 7", { x: 0.6, y: 0.45, w: 6, h: 0.35, fontSize: 13, color: C.grey, bold: true, charSpacing: 1 });
  txt(s, "Export and Post to Teams", { x: 0.6, y: 0.8, w: 6, h: 0.65, fontSize: 32, color: C.black });
  const steps = [
    ["Export", "File → Export as → PNG."],
    ["Tick \"Include a copy of my diagram\"", "The PNG then opens again in draw.io and you can still edit it. You only need to share one file per page."],
    ["Do both pages", "Switch to the Mobile tab and export again."],
    ["Post", "In the class Teams channel: both images, your name and client, your user stories, and one sentence about a mobile change."],
  ];
  steps.forEach(([t, d], i) => {
    const y = 1.8 + i * 1.3;
    numCircle(s, i + 1, 0.6, y, C.orange, 0.48);
    txt(s, t, { x: 1.3, y: y - 0.02, w: 5.2, h: 0.35, fontSize: 17, bold: true, color: C.black });
    txt(s, d, { x: 1.3, y: y + 0.36, w: 5.2, h: 0.85, fontSize: 14, color: C.dark, valign: "top" });
  });
  // mock Teams post
  const px = 7.0, pw = 5.8;
  card(s, { x: px, y: 0.6, w: pw, h: 6.3 });
  s.addShape(pres.shapes.OVAL, { x: px + 0.25, y: 0.82, w: 0.5, h: 0.5, fill: { color: C.teal }, line: { color: C.teal } });
  txt(s, "JS", { x: px + 0.25, y: 0.82, w: 0.5, h: 0.5, fontSize: 13, bold: true, color: C.white, align: "center", valign: "middle" });
  txt(s, [{ text: "Jamie Symonds", options: { bold: true, color: C.black } }, { text: "   10:42 AM", options: { color: C.mid, fontSize: 11 } }],
    { x: px + 0.9, y: 0.85, w: pw - 1.1, h: 0.4, fontSize: 13 });
  txt(s, "Harbourside Fish & Chips: online ordering page", { x: px + 0.9, y: 1.35, w: pw - 1.1, h: 0.35, fontSize: 13.5, bold: true, color: C.black });
  txt(s, "Stories: browse by category · pickup or delivery · see my total · special instructions · order with one thumb",
    { x: px + 0.9, y: 1.72, w: pw - 1.1, h: 0.55, fontSize: 12, color: C.dark, valign: "top" });
  txt(s, "On mobile I replaced the order sidebar with a bar at the bottom of the screen, because there's no room for a sidebar but users still need to see their total.",
    { x: px + 0.9, y: 2.3, w: pw - 1.1, h: 0.75, fontSize: 12, color: C.dark, valign: "top" });
  // thumbnails
  const tY = 3.2, tH = 1.75;
  s.addShape(pres.shapes.RECTANGLE, { x: px + 0.9, y: tY, w: 2.75, h: tH, fill: { color: C.light }, line: { color: "D0D7DB", width: 0.75 } });
  drawWire(s, WF.desktop, { x: px + 0.95, y: tY + 0.05, w: 2.65, h: tH - 0.1 }, { crop: DESK_CROP, show: noMeta });
  s.addShape(pres.shapes.RECTANGLE, { x: px + 3.8, y: tY, w: 1.75, h: tH, fill: { color: C.light }, line: { color: "D0D7DB", width: 0.75 } });
  drawWire(s, WF.mobile, { x: px + 3.85, y: tY + 0.05, w: 1.65, h: tH - 0.1 }, { crop: MOB_CROP, show: noMeta });
  // reply
  s.addShape(pres.shapes.LINE, { x: px + 0.25, y: 5.2, w: pw - 0.5, h: 0, line: { color: "E0E0E0", width: 1 } });
  s.addShape(pres.shapes.OVAL, { x: px + 0.9, y: 5.38, w: 0.4, h: 0.4, fill: { color: C.purple }, line: { color: C.purple } });
  txt(s, "AB", { x: px + 0.9, y: 5.38, w: 0.4, h: 0.4, fontSize: 10.5, bold: true, color: C.white, align: "center", valign: "middle" });
  txt(s, [
    { text: "✅ Works: ", options: { bold: true } }, { text: "I could see my total the whole time, even on the phone.", options: { breakLine: true } },
    { text: "❓ Wonder: ", options: { bold: true } }, { text: "where would I type my address for delivery?" },
  ], { x: px + 1.45, y: 5.35, w: pw - 1.7, h: 1.3, fontSize: 12, color: C.dark, valign: "top", paraSpaceAfter: 4 });
  s.addNotes("Do this live: export both pages and post them to the class Teams channel. That gives students a real example of the posting format before they do their own.\n\n'Include a copy of my diagram' puts the draw.io data inside the PNG. If you open that PNG in draw.io, it comes back as an editable diagram, so the image is the source file too.\n\nThe reply shown is an example of the feedback format they'll use in Part 3. The 'wonder' is a fair point too: delivery would need an address field, which this wireframe doesn't show yet.");
}

// 20. Divider: your turn
dividerSlide("Your Turn", "Wireframe one page of your own   ·   desktop + mobile   ·   share it on Teams",
  "About 50 minutes: 10 to pick a client and write stories, 30 to wireframe, 10 to post and reply.\n\nWalk around during Part 2. Common problems are on the 'Watch Out For' slide.");

// 21. Part 1: pick a client, write stories
{
  const s = activitySlide("Your Turn – Part 1: Pick a Client, Write Stories (10 min)",
    "Students choose one of these or pitch their own (check it quickly: it needs one key page with something to DO, not just information).\n\nThe key page for each: bakery = pre-order form, band = shows and tickets, groomer = booking, hockey = schedule and registration, food truck = today's location and menu, campground = site picker and dates.\n\nStories must follow the format from last week. 3–4 is plenty.");
  const clients = [
    ["🥐", "Bakery", "Pre-order a birthday cake"],
    ["🎸", "Local band", "Upcoming shows and tickets"],
    ["🐾", "Dog groomer", "Book an appointment"],
    ["🏒", "Minor hockey", "Team schedule and registration"],
    ["🚚", "Food truck", "Today's location and menu"],
    ["🏕️", "Campground", "Pick a site, reserve dates"],
  ];
  const cw = 2.37, ch = 1.7, gap = (7.6 - 3 * cw) / 2;
  clients.forEach(([e, t, d], i) => {
    const x = 0.75 + (i % 3) * (cw + gap), y = 1.45 + Math.floor(i / 3) * (ch + 0.22);
    card(s, { x, y, w: cw, h: ch });
    txt(s, e, { x: x + 0.2, y: y + 0.15, w: 0.6, h: 0.55, fontSize: 26, fontFace: "Segoe UI Emoji" });
    txt(s, t, { x: x + 0.2, y: y + 0.75, w: cw - 0.4, h: 0.38, fontSize: 16, bold: true, color: C.black });
    txt(s, d, { x: x + 0.2, y: y + 1.12, w: cw - 0.4, h: 0.5, fontSize: 13, color: C.grey, valign: "top" });
  });
  card(s, { x: 0.75, y: 5.35, w: 7.6, h: 1.1, fill: "FBE9DF", border: "F2C6AC" });
  txt(s, [{ text: "💡 Or pitch your own. ", options: { bold: true } }, { text: "Check with me first. It needs one main page where the user does something, not just reads." }],
    { x: 1.0, y: 5.35, w: 7.1, h: 1.1, fontSize: 14, color: C.dark, valign: "middle" });
  card(s, { x: 8.75, y: 1.45, w: 3.85, h: 5.0, fill: C.dark, border: C.dark });
  txt(s, "Then write 3–4 user stories", { x: 9.0, y: 1.7, w: 3.4, h: 0.7, fontSize: 18, bold: true, color: C.white, valign: "top" });
  txt(s, "As a ___, I want to ___ so that ___.", { x: 9.0, y: 2.5, w: 3.4, h: 0.6, fontSize: 15, italic: true, color: "F2C6AC", valign: "top" });
  txt(s, [
    { text: "Example", options: { bold: true, color: C.white, breakLine: true } },
    { text: "As a dog owner, I want to see open appointment times so that I can book without calling.", options: { color: "D9D9D9", breakLine: true } },
    { text: " ", options: { fontSize: 8, breakLine: true } },
    { text: "Include at least one story about using a phone.", options: { color: "D9D9D9", bold: true } },
  ], { x: 9.0, y: 3.3, w: 3.4, h: 2.8, fontSize: 14, valign: "top" });
}

// 22. Part 2: wireframe it (checklist)
{
  const s = activitySlide("Your Turn – Part 2: Wireframe It (30 min)",
    "Circulate. The most common problems are on the next-but-one slide (Watch Out For).\n\nIf someone finishes early: add a third frame for a tablet (about 768px), or wireframe a second screen (e.g. the confirmation page after checkout).");
  card(s, { x: 0.75, y: 1.45, w: 7.6, h: 5.5 });
  txt(s, "Your wireframe must have…", { x: 1.05, y: 1.65, w: 7, h: 0.4, fontSize: 18, bold: true, color: C.black });
  const checks = [
    [["Your ", ["user stories", "b"], " written at the top of the page"]],
    [["A ", ["Desktop", "b"], " page tab (1200px wide) and a ", ["Mobile", "b"], " page tab (375px wide)"]],
    [["Low-fi only: greyscale, X boxes for images, grey bars for text"]],
    [["Every story numbered ", ["① ② ③", "b"], " on the part of the page that handles it"]],
    [["Landmark labels: ", ["<header>", "c"], " ", ["<nav>", "c"], " ", ["<main>", "c"], " ", ["<footer>", "c"], " + one more"]],
    [["At least ", ["two things rearranged", "b"], " for mobile, not just made narrower"]],
    [["Mobile buttons that look ", ["44px or bigger", "b"]]],
  ];
  checks.forEach((segs, i) => {
    const y = 2.25 + i * 0.64;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.05, y: y + 0.05, w: 0.3, h: 0.3, rectRadius: 0.05, fill: { color: C.white }, line: { color: C.teal, width: 1.5 } });
    richBullets(s, segs, { x: 1.55, y, w: 6.6, h: 0.45, size: 14.5, bullet: false, color: C.dark });
  });
  // frame sizes visual
  card(s, { x: 8.75, y: 1.45, w: 3.85, h: 2.65 });
  s.addShape(pres.shapes.RECTANGLE, { x: 9.05, y: 1.8, w: 2.0, h: 1.4, fill: { color: C.white }, line: { color: "333333", width: 1.25 } });
  s.addShape(pres.shapes.RECTANGLE, { x: 9.05, y: 1.8, w: 2.0, h: 0.14, fill: { color: "EEEEEE" }, line: { color: "333333", width: 1.25 } });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 11.5, y: 1.8, w: 0.72, h: 1.4, rectRadius: 0.08, fill: { color: C.white }, line: { color: "333333", width: 1.25 } });
  txt(s, "1200px", { x: 9.05, y: 3.3, w: 2.0, h: 0.3, fontSize: 13, bold: true, color: C.black, align: "center", fontFace: MONO });
  txt(s, "375px", { x: 11.26, y: 3.3, w: 1.2, h: 0.3, fontSize: 13, bold: true, color: C.black, align: "center", fontFace: MONO });
  txt(s, "Two page tabs, one file", { x: 9.0, y: 3.62, w: 3.4, h: 0.35, fontSize: 13, color: C.grey, align: "center" });
  card(s, { x: 8.75, y: 4.3, w: 3.85, h: 2.65, fill: C.dark, border: C.dark });
  txt(s, [
    { text: "Stuck?", options: { bold: true, fontSize: 17, color: C.white, breakLine: true } },
    { text: "Start with the header and footer. Every page has them. Then add one box for each user story, and move them around until it makes sense.", options: { fontSize: 14, color: "D9D9D9" } },
  ], { x: 9.0, y: 4.5, w: 3.4, h: 2.3, valign: "top", paraSpaceAfter: 6 });
}

// 23. Part 3: post and feedback
{
  const s = activitySlide("Your Turn – Part 3: Post and Give Feedback (10 min)",
    "One post per student in the class channel, so each wireframe gets its own reply thread.\n\nThe 'fewer than 2 replies' rule spreads feedback around so nobody is left out. If you want to check participation later, each student should have one post and two replies.\n\nGood feedback is specific: a Works should name a user story, and a Wonder should name a place on the page.");
  card(s, { x: 0.75, y: 1.45, w: 5.9, h: 5.5 });
  txt(s, "Your post", { x: 1.05, y: 1.65, w: 5.3, h: 0.4, fontSize: 18, bold: true, color: C.black });
  const post = [
    ["Two PNGs", "Desktop and Mobile, exported with Include a copy of my diagram."],
    ["Name and client", "e.g. \"Jordan: Paws & Suds dog grooming, booking page\""],
    ["Your user stories", "Paste them in, numbered."],
    ["One sentence", "Something you changed for mobile, and why."],
  ];
  post.forEach(([t, d], i) => {
    const y = 2.25 + i * 1.15;
    numCircle(s, i + 1, 1.05, y, C.orange, 0.46);
    txt(s, t, { x: 1.7, y: y - 0.02, w: 4.7, h: 0.35, fontSize: 16, bold: true, color: C.black });
    txt(s, d, { x: 1.7, y: y + 0.33, w: 4.7, h: 0.65, fontSize: 13.5, color: C.dark, valign: "top" });
  });
  txt(s, "Then reply to 2 classmates", { x: 7.05, y: 1.5, w: 5.5, h: 0.45, fontSize: 18, bold: true, color: C.black });
  const fb = [
    ["✓", "Works", C.green, "Name one user story you could clearly complete, and how.", "\"I could book an appointment in two taps from the home page.\""],
    ["?", "Wonder", C.purple, "One place a user might get stuck, or something you'd move.", "\"I wonder if people will find the price list. It's below the fold on mobile.\""],
  ];
  fb.forEach(([sym, t, col, d, ex], i) => {
    const y = 2.1 + i * 2.05;
    card(s, { x: 7.05, y, w: 5.55, h: 1.85 });
    numCircle(s, sym, 7.3, y + 0.2, col, 0.46);
    txt(s, t, { x: 7.95, y: y + 0.2, w: 4.4, h: 0.46, fontSize: 17, bold: true, color: C.black, valign: "middle" });
    txt(s, d, { x: 7.3, y: y + 0.78, w: 5.1, h: 0.4, fontSize: 13.5, color: C.dark });
    txt(s, ex, { x: 7.3, y: y + 1.2, w: 5.1, h: 0.55, fontSize: 13, italic: true, color: C.grey, valign: "top" });
  });
  txt(s, "Pick posts with fewer than 2 replies, so everyone gets feedback.", { x: 7.05, y: 6.3, w: 5.55, h: 0.5, fontSize: 13.5, bold: true, color: C.teal, valign: "middle" });
}

// 24. Watch out for
{
  const s = panelSlide("Watch Out For",
    "These are the mistakes to look for while walking around during Part 2. Show this slide if you see the same one several times.");
  const rows = [
    ["Colours, fonts and photos", "Stay greyscale. Use X boxes for images."],
    ["Mobile = desktop, only narrower", "Rearrange it: stack columns, collapse the nav, move the sidebar."],
    ["A story with nowhere to happen", "Number the stories on the wireframe. Every number needs a place."],
    ["No landmarks, or everything is a <div>", "Label header, nav, main and footer at least."],
    ["Wonky alignment", "Arrange → Align, and turn on View → Grid."],
  ];
  rows.forEach(([t, d], i) => {
    const y = 1.95 + i * 0.98;
    card(s, { x: CX, y, w: CW, h: 0.82 });
    numCircle(s, "✕", CX + 0.18, y + 0.18, C.red, 0.46);
    txt(s, t, { x: CX + 0.85, y: y + 0.08, w: CW - 1.05, h: 0.35, fontSize: 16, bold: true, color: C.black });
    txt(s, d, { x: CX + 0.85, y: y + 0.43, w: CW - 1.05, h: 0.32, fontSize: 13.5, color: C.dark });
  });
}

// 25. Wrap-up
{
  const s = panelSlide("Wrap-up",
    "Put two or three Teams posts on the projector and ask these questions about each one.\n\nThe first question is the bridge to building: the wireframe's labels should turn straight into HTML landmarks.");
  labelBullets(s, [
    { label: "Recap:", text: [
      "A wireframe is a low-fi plan of a page: what goes where, not how it looks.",
      "Every user story needs a place on the page.",
      "Mobile isn't a smaller desktop. Rearrange the content for one thumb and one column.",
      "The regions you draw become the landmarks you code.",
    ] },
    { label: "Let's look at a few Teams posts:", text: [
      "If you had to write the HTML tomorrow, what's the landmark structure?",
      "What's the first thing a phone user sees before scrolling? Is it the right thing?",
    ] },
  ], { size: 20, sub: 16 });
}

pres.writeFile({ fileName: OUT }).then((f) => console.log("wrote " + f));
