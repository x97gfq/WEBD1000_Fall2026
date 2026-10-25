// Builds two decks, one per class, in the same style as the Class 5–11 decks:
//   Flexbox.pptx (Week 8, Class 1) and Grid_and_Float.pptx (Week 8, Class 2).
// Both are sprinkled with interactive slides: pop quizzes (the answer is on the next slide),
// "predict the layout", "spot the bug", a sorting game and a quick-fire exit ticket.
// To rebuild: cd 08_Flex_Grid_CSS/_build, run "npm install" once, then "npm run build".
// The "Result" pictures are screenshots of the example pages, taken with Microsoft Edge (headless) each build.
// Code slides read the example files directly, so edit those files and re-run to update the slides.
const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");
const { execFileSync } = require("child_process");

const REPO = path.join(__dirname, ".."); // the 08_Flex_Grid_CSS folder
const SCR = __dirname;
const SHOTS = path.join(SCR, "shots");

const FONT = "Aptos";
const MONO = "Consolas";
const UNIT = "Week 8: Responsive Web Design";
const C = {
  dark: "262626", light: "F2F2F2", black: "000000", white: "FFFFFF",
  orange: "E97132", purple: "A02B93", blue: "0F9ED5", navy: "12304A",
  teal: "1B5E82", grey: "595959", mid: "7F7F7F", actBg: "E1E8EB",
  green: "2E7D32", red: "C62828", ring: "B7CFC6",
};

// ---------- screenshots of the example pages ----------
const EDGE = ["C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe", "C:/Program Files/Microsoft/Edge/Application/msedge.exe"].find((p) => fs.existsSync(p));
if (!EDGE) throw new Error("Microsoft Edge not found: it's needed to screenshot the example pages.");
fs.mkdirSync(SHOTS, { recursive: true });
const PROFILE = path.join(require("os").tmpdir(), "webd1000-edge-shots");
// bars: keep the scrollbars (for the responsive table)
function shot(file, name, w = 600, h = 400, scale = 2, bars = false) {
  const out = path.join(SHOTS, name + ".png");
  const url = "file:///" + path.join(REPO, file).replace(/\\/g, "/");
  // Its own profile folder, so it never attaches to an Edge window you already have open
  const args = ["--headless=new", "--disable-gpu", ...(bars ? [] : ["--hide-scrollbars"]), "--no-first-run", `--user-data-dir=${PROFILE}`,
    `--window-size=${w},${h}`, `--force-device-scale-factor=${scale}`, `--screenshot=${out}`, url];
  for (let tries = 0; ; tries++) {
    try { execFileSync(EDGE, args, { stdio: "ignore", timeout: 60000 }); break; }
    catch (e) { if (tries >= 4) throw e; }
  }
  return { path: out, ratio: h / w };
}

// One deck per class: deck() saves the one in progress (if any) and starts the next
let pres = null, OUT = null;
function save() { if (pres) { const f = OUT; pres.writeFile({ fileName: f }).then(() => console.log("wrote " + f)); } }
function deck(title, file) {
  save();
  pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
  pres.title = title;
  pres.author = "Jamie Symonds";
  OUT = path.join(REPO, file);
}
function newSlide() { return pres.addSlide(); }

const W = 13.333, H = 7.5;
const PANEL_W = 4.1;
const CX = 5.2, CW = 7.55; // content column on panel slides

// ---------- helpers (same as the Class 5–11 decks) ----------
function txt(slide, text, opts) {
  slide.addText(text, Object.assign({ isTextBox: true, fontFace: FONT, margin: 0 }, opts));
}

function panelSlide(title, notes) {
  const s = newSlide();
  s.background = { color: C.light };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: PANEL_W, h: H, fill: { color: C.dark }, line: { color: C.dark } });
  txt(s, UNIT, { x: 0.95, y: 0.6, w: 2.8, h: 1.4, fontSize: 22, color: C.light, valign: "top" });
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
// codeColor: pass a light colour on dark cards
function richBullets(s, bullets, { x, y, w, h, size = 15, color = C.dark, space = 9, fontFace = FONT, bullet = true, valign = "top", codeColor = "8A2B0B" }) {
  const runs = [];
  bullets.forEach((segs, bi) => {
    segs.forEach((seg, si) => {
      const [t, style] = typeof seg === "string" ? [seg, ""] : seg;
      const o = { fontSize: size, fontFace: style.includes("c") ? MONO : fontFace, bold: style.includes("b"), italic: style.includes("i") };
      if (style.includes("c")) o.color = codeColor;
      if (si === 0) runs.push({ text: "​", options: { fontSize: size, fontFace, color, bullet, paraSpaceBefore: bi ? space : 0 } });
      if (si === segs.length - 1 && bi < bullets.length - 1) o.breakLine = true;
      runs.push({ text: t, options: o });
    });
  });
  txt(s, runs, { x, y, w, h, color, valign });
}

// ---- syntax highlighting (VS Code Dark+ colours) ----
const HL = { text: "D4D4D4", punct: "808080", tag: "569CD6", attr: "9CDCFE", val: "CE9178", comment: "6A9955", ent: "D7BA7D",
  sel: "D7BA7D", prop: "9CDCFE", num: "B5CEA8", brace: "D4D4D4", at: "C586C0" };

function tokPusher(toks) {
  return (t, c) => { if (!t) return; const last = toks[toks.length - 1]; if (last && last.c === c) last.t += t; else toks.push({ t, c }); };
}

function highlightHTML(lines) {
  let inComment = false, inTag = false, inVal = null;
  return lines.map((line) => {
    const toks = [], push = tokPusher(toks);
    let i = 0;
    while (i < line.length) {
      if (inComment) {
        const e = line.indexOf("-->", i);
        if (e < 0) { push(line.slice(i), HL.comment); i = line.length; }
        else { push(line.slice(i, e + 3), HL.comment); i = e + 3; inComment = false; }
      } else if (inVal) {
        const e = line.indexOf(inVal, i);
        if (e < 0) { push(line.slice(i), HL.val); i = line.length; }
        else { push(line.slice(i, e + 1), HL.val); i = e + 1; inVal = null; }
      } else if (inTag) {
        const ch = line[i];
        if (ch === ">") { push(">", HL.punct); i++; inTag = false; }
        else if (ch === "/" && line[i + 1] === ">") { push("/>", HL.punct); i += 2; inTag = false; }
        else if (ch === '"' || ch === "'") { inVal = ch; push(ch, HL.val); i++; }
        else if (ch === "=") { push("=", HL.text); i++; }
        else if (/\s/.test(ch)) { push(ch, HL.text); i++; }
        else { const m = line.slice(i).match(/^[^\s=>"']+/); push(m[0], HL.attr); i += m[0].length; }
      } else {
        if (line.startsWith("<!--", i)) { inComment = true; continue; }
        const m = line.slice(i).match(/^<(\/?)([a-zA-Z!][a-zA-Z0-9]*)/);
        if (m) { push("<" + m[1], HL.punct); push(m[2], HL.tag); i += m[0].length; inTag = true; continue; }
        const e = line.slice(i).match(/^&[#a-zA-Z0-9]+;/);
        if (e) { push(e[0], HL.ent); i += e[0].length; continue; }
        push(line[i], HL.text); i++;
      }
    }
    return toks;
  });
}

function highlightCSS(lines) {
  let inComment = false, depth = 0, inValue = false, selDepth = 0, lastSel = "";
  return lines.map((line) => {
    const toks = [], push = tokPusher(toks);
    let i = 0;
    while (i < line.length) {
      if (inComment) {
        const e = line.indexOf("*/", i);
        if (e < 0) { push(line.slice(i), HL.comment); i = line.length; }
        else { push(line.slice(i, e + 2), HL.comment); i = e + 2; inComment = false; }
        continue;
      }
      if (line.startsWith("/*", i)) { inComment = true; continue; }
      const ch = line[i];
      // Inside @keyframes and @media, the next level down holds selectors (or 0%, from, to), not declarations
      if (ch === "{") { push(ch, HL.brace); if (depth === selDepth && lastSel.trim().startsWith("@")) selDepth++; depth++; inValue = false; lastSel = ""; i++; continue; }
      if (ch === "}") { push(ch, HL.brace); depth = Math.max(0, depth - 1); selDepth = Math.min(selDepth, depth); inValue = false; i++; continue; }
      if (depth === selDepth) {
        const at = line.slice(i).match(/^@[a-z-]+/);
        if (at) { push(at[0], HL.at); lastSel += at[0]; i += at[0].length; continue; }
        const m = line.slice(i).match(/^[^{/]+/);
        if (m) { push(m[0], /^\s+$/.test(m[0]) ? HL.text : HL.sel); lastSel += m[0]; i += m[0].length; } else { push(ch, HL.sel); i++; }
        continue;
      }
      if (!inValue) {
        if (ch === ":") { push(":", HL.text); inValue = true; i++; continue; }
        const m = line.slice(i).match(/^[a-zA-Z-]+/);
        if (m) { push(m[0], HL.prop); i += m[0].length; continue; }
        push(ch, HL.text); i++; continue;
      }
      if (ch === ";") { push(";", HL.text); inValue = false; i++; continue; }
      const num = line.slice(i).match(/^-?\d+(\.\d+)?(px|rem|em|%|deg|ms|s|vw|vh)?/);
      if (num && !/[#\w]/.test(line[i - 1] || "")) { push(num[0], HL.num); i += num[0].length; continue; }
      const word = line.slice(i).match(/^[^;\s/}]+/);
      if (word) { push(word[0], HL.val); i += word[0].length; continue; }
      push(ch, HL.text); i++;
    }
    return toks;
  });
}

// lang: "html", "css", or "mixed" (CSS inside <style>…</style>, and lines from a .css file after a "/* in style.css" comment)
function highlight(lines, lang) {
  if (lang === "html") return highlightHTML(lines);
  if (lang === "css") return highlightCSS(lines);
  const out = [];
  let chunk = [], mode = "html";
  const flush = () => { if (chunk.length) out.push(...(mode === "css" ? highlightCSS(chunk) : highlightHTML(chunk))); chunk = []; };
  lines.forEach((l) => {
    const t = l.trim();
    if (mode === "html" && (t.startsWith("/*") || t === "<style>")) {
      if (t === "<style>") { chunk.push(l); flush(); mode = "css"; return; }
      flush(); mode = "css"; chunk.push(l); return;
    }
    if (mode === "css" && t.startsWith("<")) { flush(); mode = "html"; chunk.push(l); return; }
    chunk.push(l);
  });
  flush();
  return out;
}

const tabs = (l) => l.replace(/\t/g, "   ");
function readLines(file) { return fs.readFileSync(path.join(REPO, file), "utf8").split("\n").map((l) => l.replace(/\r$/, "")); }
function trim(b) { while (b.length && !b[0].trim()) b.shift(); while (b.length && !b[b.length - 1].trim()) b.pop(); return b; }
function dedent(b, n) { const re = new RegExp("^\\t{0," + n + "}"); return b.map((l) => l.replace(re, "")); }
// Every example keeps its CSS in its own stylesheet: 01_x_example.html uses 01_x_example.css
function cssOf(file) { return trim(readLines(file.replace(/\.html$/, ".css"))).map(tabs); }
function bodyOf(file) {
  const L = readLines(file);
  const a = L.findIndex((l) => l.trim() === "<body>"), b = L.findIndex((l) => l.trim() === "</body>");
  return trim(dedent(L.slice(a + 1, b), 2)).map(tabs);
}

// mark: line numbers (from 0) to flag with an orange bar in the margin
function codeBox(s, lines, { x, y, w, h, maxSize = 13, fitHeight = false, lang = "html", title, mark = [] }) {
  const top = title ? 0.38 : 0;
  const pad = Math.min(0.25, h * 0.18);
  const n = lines.length, maxLen = Math.max(...lines.map((l) => l.length));
  const byH = ((h - top - 2 * pad) * 72) / (1.18 * n);
  const byW = ((w - 2 * pad) * 72) / (0.56 * maxLen);
  const size = Math.max(9, Math.min(maxSize, byH, byW));
  if (fitHeight) {
    const nh = Math.min(h, (n * size * 1.2) / 72 + 2 * pad + top + 0.1);
    y = y + (h - nh) / 2; h = nh;
  }
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: "1E1E1E" }, line: { color: "1E1E1E" },
    shadow: { type: "outer", blur: 8, offset: 3, angle: 90, color: "000000", opacity: 0.35 } });
  if (title) txt(s, title, { x: x + pad, y: y + 0.12, w: w - 2 * pad, h: 0.3, fontSize: 11, bold: true, color: "F2C6AC", charSpacing: 1 });
  const runs = [];
  const hl = highlight(lines, lang);
  hl.forEach((toks, li) => {
    if (!toks.length) toks = [{ t: " ", c: HL.text }];
    toks.forEach((tk, ti) => {
      const o = { color: tk.c, fontFace: MONO, fontSize: Math.round(size * 2) / 2 };
      if (ti === toks.length - 1 && li < hl.length - 1) o.breakLine = true;
      runs.push({ text: tk.t, options: o });
    });
  });
  txt(s, runs, { x: x + pad, y: y + pad + top, w: w - 2 * pad, h: h - 2 * pad - top, valign: "top", lineSpacingMultiple: 1.0 });
  const lh = (Math.round(size * 2) / 2) * 1.2 / 72;
  mark.forEach((li) => s.addShape(pres.shapes.RECTANGLE, { x: x + pad * 0.3, y: y + pad + top + li * lh + lh * 0.1, w: 0.06, h: lh * 0.8,
    fill: { color: C.orange }, line: { color: C.orange } }));
}

function pill(s, text, x, y, color, w) {
  w = w || 0.3 + text.length * 0.1;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.34, rectRadius: 0.17, fill: { color }, line: { color } });
  txt(s, text, { x, y, w, h: 0.34, fontSize: 10.5, bold: true, color: C.white, align: "center", valign: "middle", charSpacing: 1 });
  return w;
}

function framedImage(s, img, x, y, w, label) {
  const h = w * img.ratio;
  s.addImage({ path: img.path, x, y, w, h });
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { type: "none" }, line: { color: "B0B8BC", width: 1 } });
  if (label) txt(s, label, { x, y: y - 0.32, w, h: 0.28, fontSize: 11, bold: true, color: C.teal, charSpacing: 1 });
  return h;
}

// Code slide: gradient left + code, title, bullets and the rendered result on the right
const CODE_PANEL = 7.0;
// frames: a grid of { img, label } screenshots in place of one result
// placeholder: a screenshot whose size to copy as an empty dashed box (a "predict the layout" question)
function codeSlide({ kicker, kickerColor = C.grey, title, bullets, file, lines, lang = "css", result, resultLabel = "RESULT", placeholder, frames, cols, maxFrameH = 3.3, notes, size = 14, codeTitle, mark }) {
  const s = newSlide();
  s.background = { color: C.white };
  s.addImage({ path: path.join(SCR, "grad_left.png"), x: 0, y: 0, w: CODE_PANEL, h: H });
  codeBox(s, lines, { x: 0.3, y: 0.4, w: CODE_PANEL - 0.6, h: H - 0.8, maxSize: 15, fitHeight: true, lang, title: codeTitle, mark });
  const RX = CODE_PANEL + 0.45, RW = W - RX - 0.45;
  txt(s, kicker, { x: RX, y: 0.45, w: RW, h: 0.35, fontSize: 14, color: kickerColor, bold: true, charSpacing: 1 });
  const long = title.length > 22;
  txt(s, title, { x: RX, y: 0.8, w: RW, h: long ? 1.05 : 0.7, fontSize: long ? 26 : 32, color: C.black, valign: "top" });
  const by = long ? 1.95 : 1.65;
  let bh = 6.85 - by;
  const shown = result || placeholder;
  if (shown) {
    // as wide as the column, but never taller than 3.05", so the bullets keep their room
    const ih = Math.min(RW * shown.ratio, 3.05), iw = ih / shown.ratio;
    const iy = 6.85 - ih - 0.05;
    if (result) framedImage(s, result, RX, iy, iw, resultLabel);
    else {
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: RX, y: iy, w: iw, h: ih, rectRadius: 0.1, fill: { color: "F7F2F8" }, line: { color: C.purple, width: 1.5, dashType: "dash" } });
      txt(s, "?", { x: RX, y: iy, w: iw, h: ih * 0.7, fontSize: 60, bold: true, color: "D9B8D5", align: "center", valign: "middle" });
      txt(s, "Sketch it on paper first", { x: RX, y: iy + ih * 0.62, w: iw, h: 0.4, fontSize: 13, italic: true, color: C.purple, align: "center" });
      txt(s, "YOUR PREDICTION", { x: RX, y: iy - 0.32, w: iw, h: 0.28, fontSize: 11, bold: true, color: C.purple, charSpacing: 1 });
    }
    bh = iy - 0.32 - by - 0.1;
  }
  if (frames) {
    // every frame in a row is the same width; each row is as tall as its tallest picture
    cols = cols || Math.min(frames.length, 2);
    const rows = Math.ceil(frames.length / cols), gap = 0.15, lab = 0.32;
    const rowRatio = [...Array(rows)].map((_, r) => Math.max(...frames.slice(r * cols, r * cols + cols).map((f) => f.img.ratio)));
    const sumRatio = rowRatio.reduce((a, b) => a + b, 0);
    let fw = (RW - gap * (cols - 1)) / cols;
    const gridH = (w) => w * sumRatio + rows * lab + (rows - 1) * gap;
    if (gridH(fw) > maxFrameH) fw = (maxFrameH - (rows - 1) * gap - rows * lab) / sumRatio;
    let y = 6.85 - gridH(fw);
    const top = y;
    rowRatio.forEach((rr, r) => {
      frames.slice(r * cols, r * cols + cols).forEach(({ img, label }, c) => framedImage(s, img, RX + c * (fw + gap), y + lab, fw, label));
      y += lab + fw * rr + gap;
    });
    bh = top - by - 0.1;
  }
  richBullets(s, bullets, { x: RX, y: by, w: RW, h: bh, size });
  if (file) txt(s, file, { x: RX, y: 6.95, w: RW, h: 0.3, fontSize: 11, color: C.mid, fontFace: MONO });
  if (notes) s.addNotes(notes);
  return s;
}

function dividerSlide(title, sub, notes) {
  const s = newSlide();
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
  const s = newSlide();
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

function table(s, rows, { x = CX, y = 1.95, w = CW, colW, size = 13, rowH = 0.4, codeCols = [] }) {
  const data = rows.map((r, ri) => r.map((cell, ci) => ({ text: cell, options: {
    fontFace: codeCols.includes(ci) && ri > 0 ? MONO : FONT, fontSize: codeCols.includes(ci) && ri > 0 ? size - 1 : size,
    color: ri === 0 ? C.white : (codeCols.includes(ci) ? "8A2B0B" : C.black), bold: ri === 0, valign: "middle",
    fill: { color: ri === 0 ? C.dark : (ri % 2 ? C.white : "E4E4E4") }, margin: [3, 6, 3, 6] } })));
  s.addTable(data, { x, y, w, colW, rowH, border: { type: "solid", pt: 0.5, color: "C8C8C8" } });
}

// "Try it" slide: numbered steps on the left, the code to type in a code box on the right
function trySlide({ n, title, time, steps, code, lang = "css", codeTitle = "IN STYLE.CSS", tip, notes, stepGap }) {
  const s = activitySlide(`Try It ${n}: ${title} (${time})`, notes);
  const LW = 6.55;
  card(s, { x: 0.75, y: 1.45, w: LW, h: 5.5 });
  stepGap = stepGap || Math.min(1.0, 5.0 / steps.length);
  steps.forEach(([t, d], i) => {
    const y = 1.72 + i * stepGap;
    numCircle(s, i + 1, 1.0, y, C.orange, 0.44);
    txt(s, t, { x: 1.6, y: y - 0.02, w: LW - 1.05, h: 0.35, fontSize: 15.5, bold: true, color: C.black });
    if (d) richBullets(s, [typeof d === "string" ? [d] : d], { x: 1.6, y: y + 0.32, w: LW - 1.05, h: stepGap - 0.36, size: 13, bullet: false, color: C.dark });
  });
  const RX = 0.75 + LW + 0.35, RW = W - RX - 0.6;
  const tipH = tip ? 0.95 : 0;
  codeBox(s, code, { x: RX, y: 1.45, w: RW, h: 5.5 - (tip ? tipH + 0.2 : 0), maxSize: 15, lang, title: codeTitle });
  if (tip) {
    card(s, { x: RX, y: 6.95 - tipH, w: RW, h: tipH, fill: "FBE9DF", border: "F2C6AC" });
    richBullets(s, tip, { x: RX + 0.25, y: 6.95 - tipH, w: RW - 0.5, h: tipH, size: 13, bullet: false, valign: "middle" });
  }
  return s;
}


// A copy of an example page, with a <base> so its links and images still work, and extra code injected before </body>:
//   hover: a selector; every :hover rule is applied to the elements it matches, as if the mouse were over them
//   js:    any other script to run first (e.g. scroll to a section)
function frame(file, name, { hover = null, js = "", w = 600, h = 400, scale = 2, bars = false } = {}) {
  const abs = path.join(REPO, file);
  const base = "file:///" + path.dirname(abs).replace(/\\/g, "/") + "/";
  let inject = "";
  if (hover) inject += `<style>* { transition: none !important; }</style><script>
    for (const sheet of document.styleSheets) for (const r of [...sheet.cssRules])
      if (r.selectorText && r.selectorText.includes(":hover"))
        sheet.insertRule(r.cssText.replace(/:hover/g, ".sim-hover"), sheet.cssRules.length);
    document.querySelectorAll(${JSON.stringify(hover)}).forEach((e) => e.classList.add("sim-hover"));
  </script>`;
  if (js) inject += `<script>${js}</script>`;
  // Linked stylesheets are copied in as <style>: on a file:// page, scripts can't read a linked sheet's rules
  const src = fs.readFileSync(abs, "utf8")
    .replace(/<link rel="stylesheet" href="([^"]+)">/g, (m, href) => `<style>${fs.readFileSync(path.join(path.dirname(abs), href), "utf8")}</style>`)
    .replace("<head>", `<head><base href="${base}">`).replace("</body>", inject + "</body>");
  const tmp = path.join(SHOTS, name + ".html");
  fs.writeFileSync(tmp, src);
  return shot(path.relative(REPO, tmp), name, w, h, scale, bars);
}

// The lines from the one that starts with `from` up to (not including) the one that starts with `to`
function between(lines, from, to) {
  const a = lines.findIndex((l) => l.trim().startsWith(from));
  const b = to ? lines.findIndex((l, i) => i > a && l.trim().startsWith(to)) : lines.length;
  if (a < 0 || b < 0) throw new Error(`between(): couldn't find "${from}" / "${to}"`);
  return trim(lines.slice(a, b));
}

// A true 360 px phone width: headless Edge won't make a window that narrow, so the page goes in a 360 px iframe
function phone(file, name, w = 360, h = 640) {
  const src = "file:///" + path.join(REPO, file).replace(/\\/g, "/");
  fs.writeFileSync(path.join(SHOTS, name + ".html"), `<!DOCTYPE html><html><body style="margin: 0"><iframe src="${src}" width="${w}" height="${h}" style="border: 0; display: block"></iframe></body></html>`);
  return shot(path.relative(REPO, path.join(SHOTS, name + ".html")), name, w, h);
}


function titleSlide(cls, sub, notes) {
  const s = newSlide();
  s.background = { color: C.white };
  txt(s, "WEBD 1000", { x: 0.3, y: 0.3, w: 3, h: 0.4, fontSize: 16, color: C.black });
  s.addShape(pres.shapes.OVAL, { x: 0.55, y: 1.1, w: 5.6, h: 5.6, fill: { color: C.orange }, line: { color: C.orange } });
  txt(s, UNIT, { x: 1.45, y: 2.4, w: 4.2, h: 2.9, fontSize: 46, color: C.white, valign: "middle" });
  s.addShape(pres.shapes.OVAL, { x: 1.0, y: 5.75, w: 0.62, h: 0.62, fill: { color: C.purple }, line: { color: C.purple } });
  s.addShape(pres.shapes.ARC, { x: 9.6, y: 0.75, w: 3.3, h: 3.3, angleRange: [200, 20], line: { color: C.blue, width: 7, dashType: "dash" } });
  txt(s, [
    { text: cls + ":", options: { bullet: true, fontSize: 30, breakLine: true } },
    { text: sub, options: { bullet: { code: "25CB" }, indentLevel: 1, fontSize: 26, paraSpaceBefore: 6 } },
  ], { x: 6.85, y: 1.9, w: 5.6, h: 2.4, color: C.black, valign: "top" });
  s.addNotes(notes);
}

function attendanceSlide(notes) {
  const s = newSlide();
  s.background = { color: C.white };
  txt(s, "Attendance", { x: 1.1, y: 0.75, w: 8, h: 0.9, fontSize: 44, color: C.black });
  s.addNotes(notes);
}

// Overview: title, agenda lines, and three step cards along the bottom (the middle one orange)
function overviewSlide(heading, title, agenda, steps, notes) {
  const s = panelSlide(heading, notes);
  txt(s, [
    { text: "Title: ", options: { bold: true, bullet: true, fontSize: 22 } },
    { text: title, options: { fontSize: 22, breakLine: true } },
    { text: " ", options: { fontSize: 10, breakLine: true } },
    { text: "Agenda:", options: { bold: true, bullet: true, fontSize: 22, breakLine: true } },
    ...agenda.map((a, i) => ({ text: a, options: { bullet: true, indentLevel: 1, fontSize: 18, paraSpaceBefore: 4, breakLine: i < agenda.length - 1 } })),
  ], { x: CX, y: 1.95, w: CW, h: 3.6, color: C.black, valign: "top" });
  steps.forEach(([t, d], i) => {
    const x = CX + i * 2.6;
    card(s, { x, y: 5.6, w: 2.2, h: 1.15, fill: i === 1 ? C.orange : C.white, border: i === 1 ? C.orange : "D0D7DB" });
    txt(s, t, { x: x + 0.2, y: 5.7, w: 1.8, h: 0.45, fontSize: 17, bold: true, color: i === 1 ? C.white : C.black });
    txt(s, d, { x: x + 0.2, y: 6.15, w: 1.8, h: 0.5, fontSize: 12.5, color: i === 1 ? C.white : C.grey, valign: "top" });
    if (i < 2) txt(s, "→", { x: x + 2.2, y: 5.6, w: 0.4, h: 1.15, fontSize: 22, bold: true, color: C.orange, align: "center", valign: "middle" });
  });
}

// ---------- new for Week 8: interactive slides ----------

// "Hello `code` world" → richBullets segments, with the `backticked` parts in code font
function segs(t) { return t.split("`").map((p, i) => [p, i % 2 ? "c" : ""]).filter(([p]) => p); }

function checkInSlide(notes) {
  const s = newSlide();
  s.background = { color: C.white };
  txt(s, "Attendance and a Quick Check-in", { x: 0.75, y: 0.5, w: 11.8, h: 0.7, fontSize: 28, bold: true, color: C.black });
  const items = [["OUTLOOK", "Are you checking your school Outlook email?"], ["TEAMS", "Are you signed into Teams?"],
    ["TEAMS", "Are you following our Teams channel and General chat?"], ["BRIGHTSPACE", "Are you following the Content on Brightspace?"]];
  const cw = 5.75, ch = 2.0;
  items.forEach(([k, q], i) => {
    const x = 0.75 + (i % 2) * (cw + 0.35), y = 1.6 + Math.floor(i / 2) * (ch + 0.35);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: cw, h: ch, rectRadius: 0.1, fill: { color: "E4EFF3" }, line: { color: "E4EFF3" } });
    numCircle(s, "✓", x + 0.35, y + (ch - 0.75) / 2, C.teal, 0.75);
    txt(s, k, { x: x + 1.35, y: y + 0.35, w: cw - 1.6, h: 0.3, fontSize: 12, bold: true, color: C.teal, charSpacing: 1 });
    txt(s, q, { x: x + 1.35, y: y + 0.65, w: cw - 1.6, h: 1.1, fontSize: 19, bold: true, color: C.navy, valign: "top" });
  });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 6.6, w: 12.1, h: 0.28, fill: { color: C.teal }, line: { color: C.teal } });
  s.addNotes(notes);
}

// A slide that opens a short video (the hook at the start of each class)
function videoSlide({ title, label, url, watchFor, notes }) {
  const s = panelSlide(title, notes);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: CX, y: 1.95, w: CW, h: 2.3, rectRadius: 0.12, fill: { color: "1E1E1E" }, line: { color: "1E1E1E" },
    hyperlink: { url, tooltip: label } });
  s.addShape(pres.shapes.OVAL, { x: CX + 0.45, y: 2.55, w: 1.1, h: 1.1, fill: { color: C.red }, line: { color: C.red } });
  s.addShape(pres.shapes.ISOSCELES_TRIANGLE, { x: CX + 0.83, y: 2.83, w: 0.45, h: 0.55, rotate: 90, fill: { color: C.white }, line: { color: C.white } });
  txt(s, label, { x: CX + 1.85, y: 2.35, w: CW - 2.1, h: 0.6, fontSize: 24, bold: true, color: C.white });
  txt(s, "Fireship on YouTube · about 2 minutes", { x: CX + 1.85, y: 2.95, w: CW - 2.1, h: 0.4, fontSize: 14, color: "BFBFBF" });
  txt(s, [{ text: url, options: { hyperlink: { url }, color: "F2C6AC" } }], { x: CX + 1.85, y: 3.45, w: CW - 2.1, h: 0.4, fontSize: 14, fontFace: MONO });
  txt(s, "Words you'll know by the end of class:", { x: CX, y: 4.55, w: CW, h: 0.4, fontSize: 16, bold: true, color: C.black });
  watchFor.forEach((w, i) => pill(s, w, CX + (i % 3) * 2.5, 5.1 + Math.floor(i / 3) * 0.5, [C.orange, C.teal, C.purple][i % 3], 2.3));
}

// A pop quiz is two slides: the question (everyone votes), then the same slide with the answer marked and a "why".
// options: 4 strings (`backticks` = code font). answer: 0–3. code: optional lines shown on the left.
let quizNo = 0;
const OPT = [C.orange, C.blue, C.purple, C.green];
// revealImg: on the answer slide, a screenshot of what the browser does takes the code's place
function quiz({ kicker, q, options, answer, why, code, lang = "css", codeTitle, revealImg, notes, revealNotes }) {
  kicker = kicker || `POP QUIZ ${++quizNo}`;
  [false, true].forEach((reveal) => {
    const s = newSlide();
    s.background = { color: C.white };
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 0.14, fill: { color: C.purple }, line: { color: C.purple } });
    pill(s, reveal ? kicker + ": ANSWER" : kicker, 0.75, 0.4, reveal ? C.green : C.purple);
    txt(s, "?", { x: 11.6, y: 0.2, w: 1.2, h: 1.4, fontSize: 80, bold: true, color: "EFE3EE", align: "center", valign: "middle" });
    richBullets(s, [segs(q).map(([t, st]) => [t, st + "b"])], { x: 0.75, y: 0.9, w: 10.9, h: 1.05, size: 23, bullet: false, color: C.black });
    let ox, ow, ch, gap, cols;
    const oy = 2.15;
    if (code) {
      if (reveal && revealImg) {
        const iw = Math.min(5.7, 4.4 / revealImg.ratio);
        framedImage(s, revealImg, 0.75, 2.5, iw, "WHAT THE BROWSER DOES");
      } else codeBox(s, code, { x: 0.75, y: 2.15, w: 5.7, h: 4.7, maxSize: 15, fitHeight: true, lang, title: codeTitle });
      ox = 6.75; ow = W - 0.75 - ox; cols = 1; ch = reveal ? 0.62 : 0.95; gap = reveal ? 0.12 : 0.18;
    } else {
      ox = 0.75; ow = W - 1.5; cols = 2; ch = reveal ? 1.0 : 1.45; gap = 0.22;
    }
    const cw = cols === 2 ? (ow - gap) / 2 : ow;
    options.forEach((o, i) => {
      const x = ox + (cols === 2 ? (i % 2) * (cw + gap) : 0), y = oy + (cols === 2 ? Math.floor(i / 2) : i) * (ch + gap);
      const right = i === answer, faded = reveal && !right;
      card(s, { x, y, w: cw, h: ch, fill: reveal ? (right ? "E8F5E9" : "F5F5F5") : C.white,
        border: reveal ? (right ? C.green : "E0E0E0") : OPT[i], bw: reveal && right ? 2.5 : 1.25 });
      const d = Math.min(0.5, ch - 0.16);
      numCircle(s, reveal && right ? "✓" : "ABCD"[i], x + 0.18, y + (ch - d) / 2, faded ? "BDBDBD" : reveal ? C.green : OPT[i], d);
      richBullets(s, [segs(o)], { x: x + d + 0.35, y, w: cw - d - 0.5, h: ch, size: code ? (reveal ? 13.5 : 15) : (reveal ? 15 : 17), bullet: false,
        color: faded ? "8C8C8C" : C.black, valign: "middle", codeColor: faded ? "B08A7E" : "8A2B0B" });
    });
    const bottom = oy + (cols === 2 ? 2 : 4) * (ch + gap) - gap;
    if (reveal) {
      const wy = bottom + 0.22, wx = ox, wh = 6.9 - wy;
      card(s, { x: wx, y: wy, w: ow, h: wh, fill: C.dark, border: C.dark });
      txt(s, "Why?", { x: wx + 0.25, y: wy + 0.12, w: 1.5, h: 0.35, fontSize: 14, bold: true, color: "F2C6AC" });
      richBullets(s, why.map((w) => (typeof w === "string" ? segs(w) : w)), { x: wx + 0.25, y: wy + 0.45, w: ow - 0.5, h: wh - 0.55, size: code ? 13 : 14.5,
        color: "E6E6E6", codeColor: "F2C6AC", space: 4 });
      s.addNotes(revealNotes || "");
    } else {
      txt(s, [{ text: "Vote: ", options: { bold: true, color: C.purple } }, { text: "hold up 1, 2, 3 or 4 fingers (A to D), or type your letter in the Teams chat. 30 seconds!" }],
        { x: 0.75, y: 6.95, w: 11.8, h: 0.35, fontSize: 13, color: C.grey });
      s.addNotes(notes || "");
    }
  });
}

// "Predict the layout": the code and an empty box, then the same code with the real screenshot
function predict({ kicker = "PREDICT THE LAYOUT", title, lines, lang = "mixed", img, prompt, bullets, file, notes, revealNotes, size = 14 }) {
  codeSlide({ kicker, kickerColor: C.purple, title, lines, lang, placeholder: img, bullets: prompt, file, notes, size });
  codeSlide({ kicker: kicker + ": ANSWER", kickerColor: C.green, title, lines, lang, result: img, resultLabel: "WHAT THE BROWSER DOES", bullets, file, notes: revealNotes, size });
}

// Six cards: a job on top, the answer (code) underneath. reveal = false shows "?" boxes instead
function quickFire(title, items, reveal, notes) {
  const s = activitySlide(title, notes);
  const cw = 3.82, ch = 2.45, gap = 0.27;
  const cols = [C.orange, C.blue, C.purple, C.green, C.teal, C.red];
  items.forEach(([q, a], i) => {
    const x = 0.75 + (i % 3) * (cw + gap), y = 1.5 + Math.floor(i / 3) * (ch + gap);
    card(s, { x, y, w: cw, h: ch });
    numCircle(s, i + 1, x + 0.25, y + 0.25, cols[i], 0.48);
    txt(s, q, { x: x + 0.9, y: y + 0.2, w: cw - 1.1, h: 0.9, fontSize: 15.5, bold: true, color: C.black, valign: "middle" });
    const bx = { x: x + 0.25, y: y + 1.3, w: cw - 0.5, h: 0.85 };
    if (reveal) codeBox(s, [a], Object.assign({ maxSize: 15, lang: "css" }, bx));
    else {
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, Object.assign({ rectRadius: 0.08, fill: { color: "F4F8FA" }, line: { color: "B0B8BC", width: 1, dashType: "dash" } }, bx));
      txt(s, "?", Object.assign({ fontSize: 28, bold: true, color: "B0B8BC", align: "center", valign: "middle" }, bx));
    }
  });
  return s;
}

// The sorting game: which tool fits each job? items: [job, "Grid" | "Flexbox" | "Float", why]
const TOOL = { Grid: C.teal, Flexbox: C.orange, Float: C.purple };
function sortGame(title, items, reveal, notes) {
  const s = activitySlide(title, notes);
  const cw = 3.82, ch = 2.45, gap = 0.27;
  if (!reveal) Object.keys(TOOL).forEach((t, i) => pill(s, t, 8.55 + i * 1.35, 0.68, TOOL[t], 1.2));
  items.forEach(([job, tool, why], i) => {
    const x = 0.75 + (i % 3) * (cw + gap), y = 1.5 + Math.floor(i / 3) * (ch + gap);
    card(s, { x, y, w: cw, h: ch, border: reveal ? TOOL[tool] : "D0D7DB", bw: reveal ? 2 : 0.75 });
    numCircle(s, i + 1, x + 0.25, y + 0.25, C.dark, 0.45);
    txt(s, job, { x: x + 0.85, y: y + 0.15, w: cw - 1.05, h: 1.0, fontSize: 15, bold: true, color: C.black, valign: "middle" });
    if (reveal) {
      pill(s, tool, x + 0.25, y + 1.3, TOOL[tool], 1.3);
      txt(s, why, { x: x + 0.25, y: y + 1.7, w: cw - 0.5, h: 0.65, fontSize: 12, color: C.dark, valign: "top" });
    } else {
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.25, y: y + 1.3, w: cw - 0.5, h: 0.85, rectRadius: 0.08, fill: { color: "F4F8FA" }, line: { color: "B0B8BC", width: 1, dashType: "dash" } });
      txt(s, "Grid?  Flexbox?  Float?", { x: x + 0.25, y: y + 1.3, w: cw - 0.5, h: 0.85, fontSize: 14, bold: true, color: "9AA5AB", align: "center", valign: "middle" });
    }
  });
}

// An online game: what it is, where, which levels, and how the game's words map to our CSS
function gameSlide({ title, name, url, aim, how, words, notes }) {
  const s = activitySlide(title, notes);
  const LW = 6.2;
  card(s, { x: 0.75, y: 1.45, w: LW, h: 5.5, fill: C.dark, border: C.dark });
  txt(s, name, { x: 1.05, y: 1.75, w: LW - 0.6, h: 0.8, fontSize: 36, bold: true, color: C.white });
  txt(s, [{ text: url, options: { hyperlink: { url } } }], { x: 1.05, y: 2.6, w: LW - 0.6, h: 0.45, fontSize: 18, fontFace: MONO, color: "F2C6AC" });
  txt(s, aim, { x: 1.05, y: 3.3, w: LW - 0.6, h: 0.9, fontSize: 18, bold: true, color: C.white, valign: "top" });
  richBullets(s, how.map(segs), { x: 1.05, y: 4.3, w: LW - 0.6, h: 2.5, size: 14, color: "D9D9D9", codeColor: "F2C6AC", space: 6 });
  const x = 0.75 + LW + 0.3, w = W - x - 0.6;
  card(s, { x, y: 1.45, w, h: 5.5 });
  txt(s, "It's today's CSS", { x: x + 0.25, y: 1.65, w: w - 0.5, h: 0.45, fontSize: 18, bold: true, color: C.black });
  table(s, [["In the game", "What it does"], ...words], { x: x + 0.25, y: 2.2, w: w - 0.5, colW: [2.6, w - 3.1], size: 12.5, rowH: 0.42, codeCols: [0] });
}

// W3Schools work time: the instructions on the left, the links on the right
function workTimeSlide({ title, links, notes }) {
  const s = activitySlide(title, notes);
  const LW = 5.7;
  card(s, { x: 0.75, y: 1.45, w: LW, h: 5.5 });
  txt(s, "What to do", { x: 1.0, y: 1.65, w: LW - 0.5, h: 0.45, fontSize: 18, bold: true, color: C.black });
  const steps = [
    ["Open a page on the right", "Read it, top to bottom."],
    ["Click every \"Try it Yourself »\"", "The code opens beside the result."],
    ["Change something, then Run", "A different value, a new property, more boxes."],
    ["Screenshot it", ["Code and result both showing: ", ["Win + Shift + S", "b"], "."]],
  ];
  steps.forEach(([t, d], i) => {
    const y = 2.3 + i * 0.95;
    numCircle(s, i + 1, 1.0, y, C.orange, 0.42);
    txt(s, t, { x: 1.6, y: y - 0.03, w: LW - 1.1, h: 0.35, fontSize: 15, bold: true, color: C.black });
    richBullets(s, [typeof d === "string" ? [d] : d], { x: 1.6, y: y + 0.32, w: LW - 1.1, h: 0.5, size: 13, bullet: false });
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.0, y: 6.1, w: LW - 0.5, h: 0.65, rectRadius: 0.08, fill: { color: "FBE9DF" }, line: { color: "F2C6AC" } });
  txt(s, "Submit at least 5 screenshots (both classes together) to Brightspace.", { x: 1.15, y: 6.1, w: LW - 0.8, h: 0.65, fontSize: 13, bold: true, color: C.dark, valign: "middle" });
  const x = 0.75 + LW + 0.3, w = W - x - 0.6;
  card(s, { x, y: 1.45, w, h: 5.5 });
  txt(s, "W3Schools tutorials", { x: x + 0.25, y: 1.65, w: w - 0.5, h: 0.45, fontSize: 18, bold: true, color: C.black });
  const rh = Math.min(1.1, 4.6 / links.length);
  links.forEach(([name, url], i) => {
    const y = 2.25 + i * rh;
    txt(s, name, { x: x + 0.25, y, w: w - 0.5, h: 0.38, fontSize: 15, bold: true, color: C.navy });
    txt(s, [{ text: url, options: { hyperlink: { url } } }], { x: x + 0.25, y: y + 0.36, w: w - 0.5, h: 0.32, fontSize: 11.5, fontFace: MONO, color: C.teal });
  });
}

// A throwaway page made from some HTML and CSS, screenshotted (for the predict and spot-the-bug answers)
function snippet(name, html, css, w = 600, h = 300) {
  const page = `<!DOCTYPE html><html><head><meta charset="UTF-8"><style>body { font-family: "Segoe UI", Arial, sans-serif; margin: 16px; }\n${css}</style></head><body>${html}</body></html>`;
  fs.writeFileSync(path.join(SHOTS, name + ".html"), page);
  return shot(path.relative(REPO, path.join(SHOTS, name + ".html")), name, w, h);
}

// The Harbourside answer key with its style.css cut off just before the heading `upTo` (null = all of it):
// the page as it looks after a given Try It
const ANSWER_CSS = fs.readFileSync(path.join(REPO, "activity/answer/style.css"), "utf8").replace(/\r/g, "");
function harbour(name, upTo, { w = 1000, h = 700, js = "", scale = 1 } = {}) {
  const css = upTo ? ANSWER_CSS.slice(0, ANSWER_CSS.indexOf(upTo)) : ANSWER_CSS;
  const base = "file:///" + path.join(REPO, "activity/answer/").replace(/\\/g, "/");
  const src = fs.readFileSync(path.join(REPO, "activity/answer/index.html"), "utf8")
    .replace('<link rel="stylesheet" href="style.css">', `<style>${css}</style>`)
    .replace("<head>", `<head><base href="${base}">`).replace("</body>", (js ? `<script>${js}</script>` : "") + "</body>");
  fs.writeFileSync(path.join(SHOTS, name + ".html"), src);
  return shot(path.relative(REPO, path.join(SHOTS, name + ".html")), name, w, h, scale);
}

// The code for Try It n, straight from the answer key, without its comments
function tryCode(n) {
  const L = ANSWER_CSS.split("\n").map(tabs);
  const a = L.findIndex((l) => l.startsWith(`/* ---------- Try It ${n}:`));
  let b = L.findIndex((l, i) => i > a && /^\/\* (-{10}|={10})/.test(l));
  if (b < 0) b = L.length;
  const out = [];
  let inComment = false;
  L.slice(a + 1, b).forEach((l) => {
    const t = l.trim();
    if (inComment) { if (t.includes("*/")) inComment = false; return; }
    if (t.startsWith("/*")) { if (!t.includes("*/")) inComment = true; return; }
    if (!t && (!out.length || !out[out.length - 1].trim())) return; // no double blank lines
    out.push(l);
  });
  return trim(out);
}

// ---------- screenshots ----------
const SH = {
  ex01: shot("01_flex_container_example.html", "ex01", 600, 480),
  ex02: shot("02_flex_direction_example.html", "ex02", 960, 260),
  ex03: shot("03_justify_content_example.html", "ex03", 600, 570),
  ex04: shot("04_align_items_example.html", "ex04", 980, 250),
  ex05: shot("05_centering_example.html", "ex05", 600, 300),
  ex06: shot("06_flex_wrap_example.html", "ex06", 600, 420),
  ex07: shot("07_flex_grow_example.html", "ex07", 600, 520),
  ex08: shot("08_flex_navbar_example.html", "ex08", 800, 150),
  ex09: shot("09_flex_cards_example.html", "ex09", 1000, 330),
  ex09n: phone("09_flex_cards_example.html", "ex09n", 700, 640),
  ex10: shot("10_grid_columns_example.html", "ex10", 600, 560),
  ex11: shot("11_grid_span_example.html", "ex11", 600, 380),
  ex12: shot("12_grid_layout_example.html", "ex12", 600, 340),
  ex13: shot("13_grid_areas_example.html", "ex13", 600, 290),
  ex14: shot("14_grid_auto_fit_example.html", "ex14", 900, 330),
  ex14p: phone("14_grid_auto_fit_example.html", "ex14p", 400, 560),
  ex15: shot("15_float_text_wrap_example.html", "ex15", 800, 470),
  ex16: shot("16_float_clearfix_example.html", "ex16", 600, 400),
  ex17: shot("17_responsive_layout_example.html", "ex17", 1000, 380),
  ex17p: phone("17_responsive_layout_example.html", "ex17p", 360, 660),
  // The playgrounds with a few values picked, as if someone had been playing
  playFlex: frame("playground/flex_playground.html", "playFlex", { w: 1100, h: 700, scale: 1, js: `
    const set = (p, v) => { document.querySelector('[data-prop="' + p + '"]').value = v; };
    set("justify-content", "space-between"); set("align-items", "center"); set("gap", "16px"); set("order", "-1");
    document.querySelector(".controls").dispatchEvent(new Event("input"));` }),
  playGrid: frame("playground/grid_playground.html", "playGrid", { w: 1100, h: 700, scale: 1, js: `
    const set = (p, v) => { document.querySelector('[data-prop="' + p + '"]').value = v; };
    set("grid-template-columns", "repeat(4, 1fr)"); set("gap", "8px"); set("grid-column", "span 2"); set("grid-row", "span 2");
    document.querySelector(".controls").dispatchEvent(new Event("input"));` }),
  start: shot("activity/index.html", "start", 1000, 700, 1),
  c1: harbour("c1", "/* ========== Week 8, Class 2", { w: 1000, h: 700, js: "document.querySelector('.intro').style.display = 'none';" }),
  ti5: harbour("ti5", "/* ---------- Try It 6", { w: 1200, h: 820, js: "document.querySelector('.intro').style.display = 'none';" }),
  c2: harbour("c2", null, { w: 1200, h: 1180 }),
  c2p: phone("activity/answer/index.html", "c2p", 360, 740),
};

// Two little diagrams: the flexbox axes, and grid lines
function flexAxesSlide(notes) {
  const s = panelSlide("Flexbox: The Key Ideas", notes);
  const gx = CX + 0.2, gy = 2.25, gw = 4.6, gh = 1.75;
  txt(s, "FLEX CONTAINER (the parent)", { x: gx, y: gy - 0.35, w: gw, h: 0.3, fontSize: 11, bold: true, color: C.navy, charSpacing: 1 });
  s.addShape(pres.shapes.RECTANGLE, { x: gx, y: gy, w: gw, h: gh, fill: { color: C.actBg }, line: { color: C.navy, width: 1.5, dashType: "dash" } });
  [1, 2, 3].forEach((n, i) => {
    s.addShape(pres.shapes.RECTANGLE, { x: gx + 0.2 + i * 1.15, y: gy + 0.2, w: 1.0, h: gh - 0.4, fill: { color: C.navy }, line: { color: C.navy } });
    txt(s, "item " + n, { x: gx + 0.2 + i * 1.15, y: gy + 0.2, w: 1.0, h: gh - 0.4, fontSize: 13, bold: true, color: C.white, align: "center", valign: "middle" });
  });
  s.addShape(pres.shapes.LINE, { x: gx, y: gy + gh + 0.25, w: gw, h: 0, line: { color: C.orange, width: 3, endArrowType: "triangle" } });
  txt(s, "MAIN AXIS  (flex-direction: row)", { x: gx, y: gy + gh + 0.32, w: gw, h: 0.3, fontSize: 12, bold: true, color: C.orange, align: "center" });
  s.addShape(pres.shapes.LINE, { x: gx + gw + 0.3, y: gy, w: 0, h: gh, line: { color: C.teal, width: 3, endArrowType: "triangle" } });
  txt(s, "CROSS\nAXIS", { x: gx + gw + 0.45, y: gy + gh / 2 - 0.3, w: 1.0, h: 0.6, fontSize: 12, bold: true, color: C.teal });
  const parts = [
    ["Flex container", "The parent: the element with display: flex"],
    ["Flex items", "Its direct children (only those, not grandchildren)"],
    ["Main axis", "The way the items line up. flex-direction sets it"],
    ["Cross axis", "At a right angle to the main axis"],
  ];
  parts.forEach(([t, d], i) => {
    const y = 4.8 + i * 0.38;
    txt(s, t, { x: CX, y, w: 2.0, h: 0.36, fontSize: 14, bold: true, color: C.black, valign: "middle" });
    txt(s, d, { x: CX + 2.0, y, w: CW - 2.0, h: 0.36, fontSize: 14, color: C.dark, valign: "middle" });
  });
  pill(s, "justify-content → main axis", CX, 6.4, C.orange, 3.3);
  pill(s, "align-items → cross axis", CX + 3.5, 6.4, C.teal, 3.0);
}

function gridLinesSlide(notes) {
  const s = panelSlide("Grid: The Key Ideas", notes);
  const gx = CX + 0.5, gy = 2.35, cw = 1.3, rh = 0.75;
  for (let r = 0; r < 2; r++) for (let c = 0; c < 3; c++) {
    s.addShape(pres.shapes.RECTANGLE, { x: gx + c * cw, y: gy + r * rh, w: cw, h: rh, fill: { color: C.actBg }, line: { color: "B0B8BC", width: 0.75 } });
  }
  for (let l = 0; l < 4; l++) {
    s.addShape(pres.shapes.LINE, { x: gx + l * cw, y: gy - 0.12, w: 0, h: 2 * rh + 0.24, line: { color: C.navy, width: 2 } });
    numCircle(s, l + 1, gx + l * cw - 0.17, gy - 0.5, C.navy, 0.34);
  }
  // an item from column line 2 to line 4 (drawn over the lines)
  s.addShape(pres.shapes.RECTANGLE, { x: gx + cw + 0.06, y: gy + 0.06, w: 2 * cw - 0.12, h: rh - 0.12, fill: { color: "B5451B" }, line: { color: "B5451B" } });
  txt(s, "grid-column: 2 / 4", { x: gx + cw, y: gy, w: 2 * cw, h: rh, fontSize: 13, bold: true, fontFace: MONO, color: C.white, align: "center", valign: "middle" });
  txt(s, "a cell", { x: gx, y: gy + rh, w: cw, h: rh, fontSize: 12, italic: true, color: C.grey, align: "center", valign: "middle" });
  for (let l = 0; l < 3; l++) numCircle(s, l + 1, gx - 0.55, gy + l * rh - 0.17, C.teal, 0.34);
  txt(s, "column lines", { x: gx + 3 * cw + 0.3, y: gy - 0.52, w: 1.6, h: 0.35, fontSize: 12, bold: true, color: C.navy });
  txt(s, "row lines", { x: gx - 0.6, y: gy + 2 * rh + 0.25, w: 1.6, h: 0.35, fontSize: 12, bold: true, color: C.teal });
  const parts = [
    ["Grid container", "The parent: the element with display: grid"],
    ["Lines", "The edges between columns and rows, numbered from 1"],
    ["Tracks", "A column or a row. grid-template-columns sets them"],
    ["Cells and areas", "One box of the grid, or a rectangle of them"],
  ];
  parts.forEach(([t, d], i) => {
    const y = 4.8 + i * 0.38;
    txt(s, t, { x: CX, y, w: 2.0, h: 0.36, fontSize: 14, bold: true, color: C.black, valign: "middle" });
    txt(s, d, { x: CX + 2.0, y, w: CW - 2.0, h: 0.36, fontSize: 14, color: C.dark, valign: "middle" });
  });
  pill(s, "Flexbox: a row OR a column", CX, 6.4, C.orange, 3.3);
  pill(s, "Grid: rows AND columns", CX + 3.5, 6.4, C.teal, 3.0);
}

// The playground slide: a screenshot on the left, four challenges on the right
function playgroundSlide({ title, img, file, challenges, notes }) {
  const s = activitySlide(title, notes);
  const iw = 6.7;
  framedImage(s, img, 0.75, 1.75, iw, "OPEN " + file.toUpperCase());
  const x = 0.75 + iw + 0.3, w = W - x - 0.6;
  card(s, { x, y: 1.45, w, h: 5.5 });
  txt(s, "Can you…", { x: x + 0.25, y: 1.65, w: w - 0.5, h: 0.45, fontSize: 18, bold: true, color: C.black });
  challenges.forEach((c, i) => {
    const y = 2.3 + i * 1.12;
    numCircle(s, i + 1, x + 0.25, y, [C.orange, C.blue, C.purple, C.green][i], 0.42);
    richBullets(s, [segs(c)], { x: x + 0.85, y: y - 0.1, w: w - 1.1, h: 0.95, size: 14, bullet: false, valign: "top" });
  });
}

// The activity at a glance: three pictures, the eight Try Its, and what to submit
function activityIntroSlide(notes) {
  const s = activitySlide("Harbourside, Part 3: This Week's Activity", notes);
  const ih = 2.55;
  let x = 0.75;
  [[SH.start, "START"], [SH.c1, "END OF CLASS 1"], [SH.c2, "END OF CLASS 2"]].forEach(([img, label]) => {
    const w = ih / img.ratio;
    framedImage(s, img, x, 1.75, w, label);
    x += w + 0.25;
  });
  const y0 = 1.75 + ih + 0.3, bh = 6.95 - y0, LW = 7.3;
  card(s, { x: 0.75, y: y0, w: LW, h: bh });
  const tis = [["1", "The header"], ["2", "Menu cards side by side"], ["3", "Line up the prices"], ["4", "Footer + special first"],
    ["5", "A page layout with grid"], ["6", "Menu cards with grid"], ["7", "Float the photo"], ["8", "On a phone"]];
  tis.forEach(([n, t], i) => {
    const cx = 1.0 + Math.floor(i / 4) * 3.55, cy = y0 + 0.15 + (i % 4) * ((bh - 0.25) / 4);
    numCircle(s, n, cx, cy, i < 4 ? C.orange : C.teal, 0.36);
    txt(s, t, { x: cx + 0.48, y: cy - 0.04, w: 2.95, h: 0.44, fontSize: 13.5, bold: true, color: C.black, valign: "middle" });
  });
  const rx = 0.75 + LW + 0.25, rw = W - rx - 0.6;
  card(s, { x: rx, y: y0, w: rw, h: bh, fill: C.dark, border: C.dark });
  txt(s, "Submit after Class 2", { x: rx + 0.25, y: y0 + 0.12, w: rw - 0.5, h: 0.4, fontSize: 17, bold: true, color: C.white });
  txt(s, [
    { text: "Harbourside: ", options: { bold: true, color: "F2C6AC" } },
    { text: "two screenshots, wide and phone-sized.", options: { breakLine: true } },
    { text: "W3Schools: ", options: { bold: true, color: "F2C6AC" } },
    { text: "5+ \"Try it Yourself\" screenshots.", options: { breakLine: true } },
    { text: "Win + Shift + S", options: { bold: true, color: "F2C6AC", fontFace: MONO } },
    { text: " takes a screenshot." },
  ], { x: rx + 0.25, y: y0 + 0.55, w: rw - 0.5, h: bh - 0.65, fontSize: 13.5, color: "D9D9D9", valign: "top", paraSpaceBefore: 4 });
}

// =====================================================================
// WEEK 8, CLASS 1: Flexbox
deck("Week 8, Class 1: Flexbox", "Flexbox.pptx");

titleSlide("Week 8, Class 1", "Flexbox: line things up in a row or a column, and share out the space.",
  "Welcome back! Last week: navigation menus, tables, and the W3C validators.\n\nThis week is all about layout: how to put things side by side, in columns, and in the middle of the screen. Today is Flexbox, the tool for one row or one column of things. Next class is CSS Grid (rows and columns together) and float (text wrapping around a photo).\n\nToday is more interactive than usual: there are pop quizzes all the way through (vote with your fingers, or in the Teams chat), a 'predict the layout' slide, a 'spot the bug', a game, and a playground page where you can change values live.");

checkInSlide("Take attendance while everyone checks these four things. Ask for a show of hands on each one; anyone who isn't on Teams or Brightspace yet should sort it out at the break.\n\nWhile you do, have students git pull and open the new 08_Flex_Grid_CSS folder in VS Code.");

overviewSlide("Week 8, Class 1: Overview", "Flexbox", [
  "Flex containers, flex items, and the two axes",
  "Lining things up: justify-content, align-items, centring",
  "Wrapping, sizing and order, then two real-world layouts",
  "Harbourside, Part 3: Try Its 1 to 4, Flexbox Froggy and W3Schools",
], [["Learn", "Containers, items, axes"], ["Try", "Harbourside, Part 3"], ["Play", "Quizzes, a game, a playground"]],
  "Rough plan for the two hours:\n- 10 min: check-in, get the code, the 100-second video\n- 25 min: the key ideas and Examples 01 to 05, with Pop Quizzes 1 to 3 and the 'predict' slide\n- 10 min: the playground (the class drives) and Try It 1\n- 20 min: wrapping, sizing and order (Examples 06 and 07), spot the bug, Pop Quiz 4\n- 15 min: Examples 08 and 09 (Pop Quiz 5 in between), then Try Its 2 to 4\n- 10 min: Flexbox Froggy\n- 15 min: W3Schools work time\n- 5 min: exit ticket and recap\nThat's about 110 minutes. If time is short, Froggy can be homework.\n\nPop quizzes: give everyone 30 seconds to think, then count down and have them hold up 1 to 4 fingers (A to D) all at once, so nobody just copies a neighbour. Or use the Teams chat. Then show the answer slide and ask someone who got it right to explain. Nobody is marked on these: they're there to check understanding and keep energy up.");

{
  const s = panelSlide("Before We Start: Get the Code",
    "One new folder: 08_Flex_Grid_CSS. Examples 01 to 09 are Flexbox (today); 10 to 17 are Grid, float and a responsive layout (next class).\n\nNew this week: every example is two files, an HTML page and its own stylesheet with the same name. That's how real sites are built (one CSS file linked from the page, like Assignment 2), and it means students can change the CSS without scrolling past the HTML.\n\nThe playground folder has two interactive pages: the Flexbox Playground and the Grid Playground. They use a little JavaScript, which students don't need to read.\n\nThe activity folder is the Harbourside page from last week (with the answer key's nav, dropdown and hours table), plus three new wrapper elements. This week students only edit style.css. Ask them not to open activity/answer until they're stuck.");
  txt(s, [{ text: "1.  Pull this week's examples", options: { bold: true, fontSize: 18 } }], { x: CX, y: 1.95, w: CW, h: 0.4 });
  codeBox(s, ["git pull"], { x: CX, y: 2.45, w: 5.2, h: 0.7, maxSize: 16 });
  txt(s, "New folder: 08_Flex_Grid_CSS (examples 01–17, playground, activity).", { x: CX, y: 3.25, w: CW, h: 0.5, fontSize: 15, color: C.dark });
  txt(s, [{ text: "2.  Every example is now two files", options: { bold: true, fontSize: 18 } }], { x: CX, y: 3.95, w: CW, h: 0.4 });
  richBullets(s, [[["01_flex_container_example", "c"], [".html", "b"], " + ", ["01_flex_container_example", "c"], [".css", "b"], ". Open the HTML with Live Server, change the CSS, save, watch."]],
    { x: CX, y: 4.4, w: CW, h: 0.8, size: 15, bullet: false });
  txt(s, [{ text: "3.  The activity: Harbourside, Part 3", options: { bold: true, fontSize: 18 } }], { x: CX, y: 5.4, w: CW, h: 0.4 });
  txt(s, "activity/index.html and activity/style.css. Last week's page, ready for a real layout.", { x: CX, y: 5.85, w: CW, h: 0.8, fontSize: 15, color: C.dark });
}

videoSlide({ title: "Flexbox in 100 Seconds", label: "CSS Flexbox in 100 Seconds", url: "https://youtu.be/K74l26pE4YA",
  watchFor: ["display: flex", "main axis", "cross axis", "justify-content", "align-items", "flex-wrap"],
  notes: "A fast, fun overview to start: click the dark box (or the link) to open the video. It's about two minutes.\n\nDon't worry if it goes by too fast. Tell students the goal is just to hear the words once; by the end of class every one of them will make sense, and they'll have used them on a real page.\n\nThe Grid video in the same series opens next class." });

dividerSlide("Flexbox", "One row or one column   ·   the container and its items   ·   two axes",
  "Before Flexbox (around 2015), putting three boxes side by side, or centring something up and down, took a pile of tricks: floats, clearfixes, negative margins, tables. Flexbox made it a few lines of CSS, and every browser supports it.\n\nWe've already used it once: the horizontal menu last week was display: flex and gap. Today we find out what else it can do.");

flexAxesSlide("Four words to learn, and everything else follows.\n\nThe flex container is the parent: the element you put display: flex on. Its direct children become flex items. Only the children: grandchildren are not affected (unless their own parent is a flex container too).\n\nThe main axis is the direction the items line up in. By default it runs left to right (flex-direction: row). The cross axis is at a right angle to it.\n\nThe two pills at the bottom are the most useful thing on this slide: justify-content works along the main axis, and align-items works along the cross axis. When students mix those up (everyone does), send them back to the axes.");

codeSlide({
  kicker: "FLEXBOX 1 · EXAMPLE 01", title: "display: flex", file: "01_flex_container_example.css",
  lines: between(cssOf("01_flex_container_example.html"), "/* The container"), result: SH.ex01, size: 14,
  bullets: [
    ["Divs normally ", ["stack", "b"], ": each one takes a whole line."],
    [["display: flex", "c"], " on the ", ["parent", "b"], " puts its children side by side."],
    ["Each item is only as wide as its content. They ", ["stretch", "b"], " to the same height."],
  ],
  notes: "Open 01_flex_container_example.html with Live Server, and 01_flex_container_example.css beside it.\n\nThe two containers have exactly the same HTML. The only difference is the second one also has the class flex, and .flex has one line: display: flex.\n\nPoint out what changed:\n- The items sit in a row.\n- Each one shrinks to fit its content ('Number three' is wider than '4').\n- They're all the same height. That's align-items: stretch, the default (Example 04).\n\nTry it live: delete display: flex and save. Back to a stack. Put it back.",
});

quiz({
  q: "You want the four cards inside `.menu` to sit side by side. Where does `display: flex` go?",
  options: ["On each card: `.menu article`", "On `.menu`, the parent of the cards", "On the `body`", "On the `h2` above the cards"], answer: 1,
  why: ["`display: flex` goes on the container: the parent.", "Its direct children (the cards) become flex items and line up in a row.",
    "On each card, it would line up what's INSIDE each card instead. That's useful too (Example 09), but it doesn't put the cards side by side."],
  notes: "First pop quiz: explain the routine. 30 seconds to think, then on 3 everyone holds up 1, 2, 3 or 4 fingers at once (A to D). Or type the letter in the Teams chat.\n\nThe answer is B. The most common wrong answer is A, which is worth talking about.",
  revealNotes: "Answer: B.\n\nThis is the single most common Flexbox mistake: putting display: flex on the things you want to move, instead of on their parent. If it's not working, ask 'which element is the parent?'\n\nOn the body (C), the flex items would be the body's children: the header, main and footer. They'd sit side by side, which is not what anyone wants.",
});

codeSlide({
  kicker: "FLEXBOX 2 · EXAMPLE 02", title: "flex-direction", file: "02_flex_direction_example.css",
  lines: between(cssOf("02_flex_direction_example.html"), "/* The only difference"), result: SH.ex02, size: 14,
  bullets: [
    [["flex-direction", "c"], " sets the ", ["main axis", "b"], ": which way the items line up."],
    [["row", "c"], " is the default. ", ["column", "c"], " stacks them, but they're still flex items."],
    [["-reverse", "c"], " flips the order on screen. (The HTML order stays the same.)"],
  ],
  notes: "Open 02_flex_direction_example.html.\n\nFour containers, same three items, one line different each time.\n\nWhy would you want column? A plain div stack is already a column, after all. Because a flex column can do flex things: centre the items, push one to the bottom with an auto margin (Example 09), or switch to row on a wider screen with one line in a media query (next class).\n\nNotice the column items stretch to the full width. That's align-items: stretch again: in a column, the cross axis is left to right, so stretch makes them as wide as the container.\n\nThe page itself uses flexbox too: the four demos sit side by side because .demos has display: flex and flex-wrap: wrap.",
});

quiz({
  q: "A container has `flex-direction: column`. Which way does the main axis run now?",
  options: ["Left to right", "Top to bottom", "Bottom to top", "There's no main axis in a column"], answer: 1,
  why: ["The main axis always follows `flex-direction`. Column = top to bottom.", "So in a column, `justify-content` moves things up and down, and `align-items` moves them left and right.",
    "(`column-reverse` would be bottom to top.)"],
  notes: "The trap: people learn 'justify-content is horizontal' and 'align-items is vertical'. That's only true for rows.",
  revealNotes: "Answer: B.\n\nThe point to land: the axes turn with flex-direction. justify-content is always the main axis, whichever way that runs. If you set column and justify-content: center doesn't seem to do anything, the container is probably only as tall as its content, so there's no spare room to move in. Give it a height to see it work.",
});

codeSlide({
  kicker: "FLEXBOX 3 · EXAMPLE 03", title: "justify-content", file: "03_justify_content_example.css",
  lines: between(cssOf("03_justify_content_example.html"), "/* Where does"), result: SH.ex03, size: 13,
  bullets: [
    ["Shares out the ", ["spare space", "b"], " along the ", ["main axis", "b"], "."],
    [["space-between", "c"], ": the ends touch the edges. Great for a logo left, nav right."],
  ],
  notes: "Open 03_justify_content_example.html.\n\nThe items don't change size: justify-content only decides where the spare space goes.\n\nThe three 'space' values are easy to mix up:\n- space-between: no space at the ends; equal space between items.\n- space-around: every item gets the same space on each side, so the ends get half as much as the gaps between.\n- space-evenly: every gap is exactly the same, the ends too.\n\nspace-between is the one you'll use most: logo on the left, menu on the right (Try It 1).\n\nThe gap of 6px is on every container. gap is extra space between items that justify-content doesn't touch.",
});

codeSlide({
  kicker: "FLEXBOX 4 · EXAMPLE 04", title: "align-items", file: "04_align_items_example.css",
  lines: between(cssOf("04_align_items_example.html"), "/* align-items"), result: SH.ex04, size: 14,
  bullets: [
    ["Places the items on the ", ["cross axis", "b"], " (up and down, in a row)."],
    [["stretch", "c"], " is the default: that's why flex items are all the same height."],
    [["center", "c"], " lines things up down the middle: icons beside text, a logo beside a menu."],
  ],
  notes: "Open 04_align_items_example.html. Each container is 140px tall, so the items have room to move up and down.\n\nstretch (the default) explains something students saw in Example 01: flex items in a row are all as tall as the tallest one. Handy for cards of different lengths (Try It 3).\n\nbaseline lines up the bottoms of the letters, not the boxes. Look at the big 2: its text sits on the same line as the 1 and 3. Useful when a big heading sits beside small text.\n\nTo move just one item, use align-self on that item. It takes the same values. It's in the playground.",
});

codeSlide({
  kicker: "FLEXBOX 5 · EXAMPLE 05", title: "Centre Anything", file: "05_centering_example.css",
  lines: between(cssOf("05_centering_example.html"), "/* The container"), result: SH.ex05, size: 14,
  bullets: [
    [["justify-content: center", "c"], " + ", ["align-items: center", "c"], " = the exact middle."],
    ["The container needs ", ["height", "b"], " to centre up and down: ", ["min-height: 100vh", "c"], " is the whole window."],
  ],
  notes: "Open 05_centering_example.html and resize the window: the card stays in the middle.\n\nCentring something vertically used to be a famous CSS headache (there are jokes about it). With flexbox it's these two lines.\n\nvh = viewport height: 100vh is the full height of the browser window. min-height rather than height, so the hero can still grow if the content is taller than the window (on a phone, say).\n\nTry it live: remove min-height. The navy box shrinks to fit the card, so there's no spare room, and nothing visibly centres up and down. Centring needs space to centre in.",
});

quiz({
  q: "Which two lines centre a box both ways inside a flex container?",
  options: ["`text-align: center;` and `vertical-align: middle;`", "`justify-content: center;` and `align-items: center;`",
    "`margin: center;` and `padding: center;`", "`align-content: middle;` and `justify-items: middle;`"], answer: 1,
  why: ["`justify-content` centres along the main axis (across), `align-items` along the cross axis (down).",
    "The container needs some height for the up-and-down part: Example 05 uses `min-height: 100vh`.",
    "`middle` isn't a value for either property, and `margin: center` doesn't exist."],
  notes: "A quick one, straight after Example 05. Most people should get this.",
  revealNotes: "Answer: B.\n\nA is a trap: text-align: center does centre text (and inline things) across, and vertical-align works in table cells and inline content. Neither centres a box inside a flex container.\n\nWorth knowing for next class: grid can do it in one line too, with place-items: center.",
});

{
  const html = ["<div class=\"bar\">", "   <div>A</div>", "   <div class=\"tall\">B</div>", "   <div>C</div>", "</div>", ""];
  const css = [".bar {", "   display: flex;", "   justify-content: space-between;", "   align-items: center;", "   height: 120px;", "}", "", ".tall {", "   height: 80px;", "}"];
  const img = snippet("predictBar", "<div class=\"bar\"><div>A</div><div class=\"tall\">B</div><div>C</div></div>",
    ".bar { display: flex; justify-content: space-between; align-items: center; height: 120px; background: #E1E8EB; border: 2px dashed #12304A; padding: 6px; }\n.bar div { background: #12304A; color: white; font-weight: bold; font-size: 1.4rem; padding: 8px 24px; }\n.tall { height: 80px; }",
    600, 175);
  predict({
    title: "Three Boxes in a Bar", lines: html.concat(css), img,
    prompt: [["Where do A, B and C go, left to right?"], ["And up and down? B is ", ["taller", "b"], " than the others."], ["Sketch it, then compare with a neighbour."]],
    bullets: [[["space-between", "c"], ": A at the left edge, C at the right, B in the middle."], [["align-items: center", "c"], ": all three centred up and down, so A and C line up with B's middle."]],
    notes: "Give everyone a minute to sketch the result on paper (or in a notebook app), then compare with a neighbour. Ask a couple of people to describe theirs before you show the answer.\n\nThe things to predict: the horizontal positions (justify-content: space-between) and the vertical ones (align-items: center), with one item taller than the others.",
    revealNotes: "A sits against the left edge, C against the right edge, and B exactly halfway: space-between puts all the spare space in the gaps.\n\nAll three are centred vertically in the 120px bar, so the short boxes line up with the middle of the tall one.\n\nWhat if align-items were left out? Then it's stretch, but B has its own height, so A and C stretch to the full 120px and B stays at 80px at the top. Worth trying in the playground next.",
  });
}

playgroundSlide({ title: "Your Turn to Drive: The Flexbox Playground", img: SH.playFlex, file: "playground/flex_playground.html",
  challenges: ["Put all the boxes in the exact middle of the stage.", "Move box 2 to the front, without touching the HTML.",
    "Make box 2 take up all the spare room.", "Add boxes until they run out of room. Then make them wrap onto a second line."],
  notes: "Open playground/flex_playground.html on the projector (students can open it too). The class calls out values and you change them. Or give everyone 5 minutes to try the four challenges themselves.\n\nAnswers:\n1. justify-content: center and align-items: center.\n2. Item 2: order: -1.\n3. Item 2: flex: 1 (it gets all the spare space, because the others are 0 1 auto).\n4. Click 'Add an item' several times: with nowrap they squash. Then flex-wrap: wrap.\n\nThe code box only shows properties that are not on their default value, so it shows exactly the CSS you'd need to write. Good to point out." });

activityIntroSlide("This week's hands-on part: Harbourside, Part 3. The same page as last week, now with a real layout. Four Try Its today (flexbox, in orange), four next class (grid, float and a phone layout, in blue).\n\nThe HTML has three new wrappers: a div class=\"brand\" around the logo and the h1, a class=\"intro\" on the first section, and a div class=\"cards\" around the four menu articles. Students don't change the HTML at all this week. Everything goes in style.css, under the heading for each Try It.\n\nSubmission (after Class 2): two screenshots of the finished Harbourside page (one at full width and one at phone size), plus at least five screenshots of W3Schools 'Try it Yourself' pages they changed.");

trySlide({
  n: 1, title: "The Header", time: "8 min",
  steps: [
    ["Open activity/index.html with Live Server", "Last week's page. Find the new div.brand around the logo and h1."],
    ["Type the CSS on the right", ["In style.css, under ", ["Try It 1", "b"], ". Save."]],
    ["Two header rules?", "That's fine: the later one adds to the first (and wins any ties)."],
    ["Play with it", ["Try ", ["justify-content: center", "c"], ", then ", ["flex-end", "c"], ". Make the window narrow."]],
  ],
  code: tryCode(1),
  tip: [["Two flex containers: ", [".brand", "c"], " lines up the logo and the name; ", ["header", "c"], " lines up the brand and the nav."]],
  notes: "The first Try It: the logo and name side by side, and the menu on the right of the same bar.\n\n.brand { display: flex } puts the logo and h1 in a row; align-items: center lines the logo up with the middle of the text.\n\nheader { display: flex } makes the brand div and the nav into flex items. space-between pushes them to opposite ends. flex-wrap: wrap lets the nav drop below the brand when the window is too narrow (we'll come back to that on a phone in Try It 8).\n\nThe padding changes from '20px 20px 0' to '12px 20px' so the nav sits nicely in the middle of the bar. The later rule wins, because it comes later in the file and has the same specificity (the cascade, from Class 8).\n\nThe dropdown still works: it's positioned against its own li, so the flex layout doesn't affect it.",
});

codeSlide({
  kicker: "FLEXBOX 6 · EXAMPLE 06", title: "flex-wrap and gap", file: "06_flex_wrap_example.css",
  lines: between(cssOf("06_flex_wrap_example.html"), "/* gap"), result: SH.ex06, size: 13.5,
  bullets: [
    ["By default flex keeps everything on ", ["one line", "b"], ", and squashes the items to fit."],
    [["flex-wrap: wrap", "c"], ": items keep their width and start a new line."],
    [["gap", "c"], " spaces rows and columns, never the outside edges."],
  ],
  notes: "Open 06_flex_wrap_example.html and make the window narrower and wider.\n\nEvery card asks to be 140px wide. In the first container (nowrap, the default), seven cards don't fit, so flexbox shrinks them all. Make the window very narrow and they'll even poke out of the container once they can't shrink any further than their text.\n\nIn the second container, flex-wrap: wrap lets them keep their 140px and flow onto new lines, like words in a paragraph.\n\ngap: 12px is the space between the cards, both across and down. Before gap worked in flexbox (around 2021), people used margins, which also added space at the edges and needed fixing. Use gap.",
});

quiz({
  kicker: "SPOT THE BUG",
  q: "The photos squash onto one line instead of wrapping. What's wrong?",
  code: [".gallery {", "   display: flex;", "   gap: 12px;", "}", "", ".gallery img {", "   width: 200px;", "   flex-wrap: wrap;", "}"],
  options: ["`width` has to be a percentage", "`flex-wrap` is on the items. It belongs on the container",
    "`wrap` isn't a real value: it's `flex-wrap: yes`", "Images can't be flex items"], answer: 1,
  why: ["Container properties go on the parent: `display`, `flex-direction`, `flex-wrap`, `justify-content`, `align-items`, `gap`.",
    "Move `flex-wrap: wrap;` into `.gallery` and the photos wrap.", "On an item, `flex-wrap` is silently ignored. No error!"],
  notes: "Spot the bug: same voting routine. Give them a little longer (45 seconds): they have to read the code.",
  revealNotes: "Answer: B.\n\nA useful habit: sort flexbox properties into 'container' (the parent) and 'item' (the children). Container: display, flex-direction, flex-wrap, justify-content, align-items, gap. Item: flex (grow, shrink, basis), order, align-self.\n\nThe CSS validator wouldn't catch this one either: flex-wrap is a real property with a real value. It's just on the wrong element. That's why DevTools is worth knowing: it shows a flex-wrap on an img as having no effect (some browsers grey it out with a little 'i' explaining why).",
});

codeSlide({
  kicker: "FLEXBOX 7 · EXAMPLE 07", title: "Sizing Items with flex", file: "07_flex_grow_example.css",
  lines: between(cssOf("07_flex_grow_example.html"), "/* 1."), result: SH.ex07, size: 13,
  bullets: [
    [["flex: 1", "c"], " = take ", ["one share", "b"], " of the space. ", ["flex: 2", "c"], " = two shares."],
    [["flex: 0 0 160px", "c"], " = grow 0, shrink 0, start at 160px: always exactly 160px."],
    [["order: -1", "c"], " moves an item first. (Only the look: Tab still follows the HTML.)"],
  ],
  notes: "Open 07_flex_grow_example.html. Four demos, top to bottom.\n\nflex is shorthand for three properties: flex-grow, flex-shrink and flex-basis.\n- flex-grow: how many shares of the spare space this item gets.\n- flex-shrink: how much it gives up when space is short (0 = never shrink).\n- flex-basis: its starting size, before growing or shrinking.\n\nflex: 1 is short for 1 1 0%: start from nothing, then share out ALL the space equally. That's why three flex: 1 items are exactly equal, whatever their text.\n\nDemo 3 is the classic sidebar layout: the sidebar is fixed at 160px, and main has flex: 1 so it takes whatever is left. Resize the window: only main changes size.\n\nDemo 4: order. Every item starts at order: 0; lower numbers go first. Accessibility warning: screen readers and the Tab key follow the HTML order, not the visual order. Fine for moving a card around; don't use it to rearrange things people read or Tab through in order.",
});

quiz({
  q: "Three items: `flex: 1`, `flex: 2` and `flex: 1`. The container is 800px wide, with no gap or padding. How wide is the middle item?",
  options: ["200px", "266px", "400px", "800px"], answer: 2,
  why: ["1 + 2 + 1 = 4 shares. 800px ÷ 4 = 200px per share.", "The middle item has 2 shares: 400px. The other two get 200px each.",
    "(`flex: 1` starts every item from 0, so ALL the space is shared out. Add padding and the maths gets fuzzier.)"],
  notes: "A little maths. 266px (B) is what you get if you think the middle item is 'one third, plus a bit', and 800px (D) is a guess that 2 means 'twice the container'.",
  revealNotes: "Answer: C, 400px.\n\nIf you have time, show it in the playground: set item 2 to flex: 2, and give the others flex: 1 in DevTools (or just eyeball it). Or open Example 07 and use DevTools to measure: hover the middle item in the Elements panel and the width shows on the page.",
});

codeSlide({
  kicker: "FLEXBOX 8 · EXAMPLE 08", title: "A Flexbox Header", file: "08_flex_navbar_example.css",
  lines: between(cssOf("08_flex_navbar_example.html"), "/* 1.", ".site-header a {"), result: SH.ex08, size: 14,
  bullets: [
    ["Three items: the logo, the nav and the button, lined up with ", ["align-items: center", "c"], "."],
    [["margin-left: auto", "c"], " on the nav soaks up ", ["all", "b"], " the free space, so it's pushed right."],
    ["Flex ", ["inside", "b"], " flex: the ", ["ul", "c"], " is a flex container too."],
  ],
  notes: "Open 08_flex_navbar_example.html. Almost every website header looks like this.\n\nThe new trick is margin-left: auto. In a flex container, an auto margin takes up all the free space on that side. So the nav (and the button after it) get pushed all the way to the right, and the logo stays on the left.\n\nWhy not justify-content: space-between? With three items, space-between would put the nav in the middle, halfway between the logo and the button. margin-left: auto keeps the nav and the button together on the right.\n\nThe ul inside the nav has its own display: flex: an element can be a flex item (of the header) and a flex container (for its li's) at the same time.",
});

quiz({
  q: "`margin-left: auto` pushed the nav to the right. In a card with `flex-direction: column`, what will `margin-top: auto` do to the card's button?",
  options: ["Centre it in the card", "Push it down to the bottom of the card", "Remove the space above it", "Nothing: auto margins only work sideways"], answer: 1,
  why: ["An auto margin soaks up all the free space on its side.", "In a column, the free space is up and down, so `margin-top: auto` pushes the button to the bottom.",
    "Every card's button lines up, however long its text is. That's the next example."],
  notes: "This one asks students to reason from what they just saw, rather than remember. Give them the full 30 seconds.",
  revealNotes: "Answer: B.\n\nAuto margins work along both axes in flexbox. That's the whole trick in Example 09 and Try It 3.\n\n(D is how auto margins behave in normal block layout, where margin-top: auto does nothing. Flexbox is different, which is part of why it's so useful.)",
});

codeSlide({
  kicker: "FLEXBOX 9 · EXAMPLE 09", title: "Flexbox Cards", file: "09_flex_cards_example.css",
  lines: between(cssOf("09_flex_cards_example.html"), "/* 1.", ".card h2"), size: 13,
  frames: [{ img: SH.ex09, label: "WIDE: FOUR ACROSS" }, { img: SH.ex09n, label: "NARROWER: WATCH THE LAST CARD" }], cols: 2, maxFrameH: 2.9,
  bullets: [
    [["flex: 1 1 200px", "c"], ": start at 200px, then grow to fill the row. Fewer fit on a phone."],
    ["Each card is a ", ["column", "b"], " flexbox, and ", ["margin-top: auto", "c"], " puts every button at the bottom."],
  ],
  notes: "Open 09_flex_cards_example.html and resize the window slowly.\n\n.cards is a wrapping flex container. Each card has flex: 1 1 200px: try to be 200px, grow to fill the row, shrink if needed. On a laptop four fit; narrower, three; then two; on a phone, one. No media queries.\n\nEach card is also a flex container with flex-direction: column, and the button has margin-top: auto: the answer to the last quiz. The text is different lengths, but every button lines up along the bottom.\n\nLook at the narrower screenshot: three cards fit on the first line, and the Kids card is alone on the second line, stretched all the way across. That's flexbox being one-dimensional: each line is sized on its own, and flex-grow lets the lonely card fill its line. Sometimes that's what you want; often it isn't. CSS Grid fixes it, next class (and Try It 6).",
});

trySlide({
  n: 2, title: "Menu Cards Side by Side", time: "6 min",
  steps: [
    ["Find div.cards", "In index.html: a new div around the four menu articles."],
    ["Type the CSS on the right", ["In style.css, under ", ["Try It 2", "b"], ". Save."]],
    ["Make the window narrow", "The cards wrap: four across, then fewer."],
    ["Play with it", ["Change ", ["180px", "c"], " to ", ["300px", "c"], ". How many fit now?"]],
  ],
  code: tryCode(2),
  tip: [[["flex: 1 1 180px", "c"], " = grow, shrink, start at 180px. ", ["gap", "c"], " does the spacing now, so ", ["margin-bottom: 0", "c"], "."]],
  notes: "The cards from last week finally sit side by side.\n\n.cards is the container, the articles are the items. The article rule already exists in the Class 8 CSS (#menu article, with the white background and border); this new rule adds the flex sizing and removes the old bottom margin, because gap spaces them now.\n\nAt full width (900px main column), four cards fit. Make the window narrower to see them wrap.\n\nAsk: are the prices lined up? No! The descriptions are different lengths, so each price sits at a different height. That's Try It 3.",
});

trySlide({
  n: 3, title: "Line Up the Prices", time: "5 min",
  steps: [
    ["Look at the prices", "They sit at different heights: the text above each one is a different length."],
    ["Type the CSS on the right", ["Under ", ["Try It 3", "b"], ". Save."]],
    ["Every price on one line", "The cards were already the same height. Now each price sits at the bottom."],
  ],
  code: tryCode(3),
  tip: [["Same trick as Example 09's button: a ", ["column", "b"], " flexbox, then ", ["margin-top: auto", "c"], "."]],
  notes: "Short and satisfying.\n\nThe cards are already the same height, because flex items stretch (align-items: stretch, the default). But inside each card, the content sits at the top, so the prices land wherever the text ends.\n\nMaking each card a column flexbox and giving the price margin-top: auto pushes it to the bottom of its card. All four prices line up.\n\nA second '.cards article' rule: that's fine, the browser combines it with the one from Try It 2.",
});

trySlide({
  n: 4, title: "The Footer, and Today's Special First", time: "5 min",
  steps: [
    ["Type the CSS on the right", ["Under ", ["Try It 4", "b"], ". Save."]],
    ["Scroll to the footer", "Phone number on the left, copyright on the right."],
    ["Look at the menu", ["The Lobster Roll special jumped to the front: ", ["order: -1", "c"], "."]],
    ["Play with it", ["Give ", ["#haddock", "c"], " an ", ["order: 1", "c"], ". Where does it go?"]],
  ],
  code: tryCode(4),
  tip: [[["order", "c"], " only changes the ", ["look", "b"], ". Screen readers and the Tab key still follow the HTML."]],
  notes: "Two quick ones.\n\nThe footer: the same space-between pattern as the header, with flex-wrap so the two lines stack on a narrow phone.\n\norder: -1 puts the special first, because every other card has the default order: 0. In 'play with it', #haddock { order: 1 } sends Haddock to the end: it's now the only card with a higher number than 0.\n\nThe accessibility point again: order is fine here because the cards don't have to be read in a particular order. Don't use it for things like the steps of a recipe or form fields.\n\nCheckpoint: that's all four of today's Try Its. Students keep this page for next class (Try Its 5 to 8 build on it).",
});

gameSlide({
  title: "Game Time: Flexbox Froggy (10 min)", name: "Flexbox Froggy", url: "https://flexboxfroggy.com/",
  aim: "Get as far as you can in 10 minutes. Then race a neighbour.",
  how: ["Type CSS to move each frog onto the lily pad of its colour.", "It starts with `justify-content` and `align-items`, then adds `flex-direction`, `order`, `align-self` and `flex-wrap`.",
    "Stuck? Click a property name in the instructions to see its values."],
  words: [["justify-content", "Moves frogs along the main axis"], ["align-items", "Moves them on the cross axis"], ["flex-direction", "Which way the row runs"],
    ["order", "Changes one frog's place"], ["align-self", "Moves just one frog"], ["flex-wrap", "Lets frogs wrap onto new lines"]],
  notes: "Flexbox Froggy is a free browser game (no sign-up) by Codepip. Each level gives you a pond with frogs and lily pads, and you type the CSS to put each frog on its lily pad. It uses exactly today's properties.\n\nSuggestion: 10 minutes, then ask who got furthest, and have them show the level they're on. The later levels combine properties (flex-direction with justify-content and align-items) and get tricky.\n\nIt's also good homework, and good revision before the next quiz. There's a matching game for Grid next class (Grid Garden), from the same people.",
});

workTimeSlide({
  title: "Work Time: W3Schools Flexbox (15 min)",
  links: [["CSS Flexbox", "https://www.w3schools.com/css/css3_flexbox.asp"], ["CSS Flex Container", "https://www.w3schools.com/css/css3_flexbox_container.asp"],
    ["CSS Flex Items", "https://www.w3schools.com/css/css3_flexbox_items.asp"], ["CSS Flex Responsive", "https://www.w3schools.com/css/css3_flexbox_responsive.asp"]],
  notes: "From the original Week 8 deck: work through the W3Schools Flexbox pages, and do as many 'Try it Yourself' examples as possible. In each one, change the CSS (a different value, a new property, more boxes) and click Run.\n\nDo one together first, on the projector: open CSS Flexbox, click Try it Yourself, change flex-wrap: nowrap to wrap (or change the background colour), click Run, and take a screenshot that shows both the code and the result.\n\nSubmission: at least 5 screenshots in total, across this class and next class's Grid pages, uploaded to Brightspace after Class 2. Each screenshot should show code the student changed, not the original.",
});

{
  const s = panelSlide("Cheat Sheet: Flexbox", "Everything from today in one place. It's in the README too.");
  table(s, [
    ["Code", "What it does"],
    ["display: flex", "The parent becomes a flex container"],
    ["flex-direction: row | column", "Sets the main axis"],
    ["justify-content: center | space-between", "Places items along the main axis"],
    ["align-items: stretch | center", "Places items on the cross axis"],
    ["flex-wrap: wrap", "Items can move onto new lines"],
    ["gap: 16px", "Space between items (not the edges)"],
    ["flex: 1", "An item takes a share of the spare space"],
    ["flex: 0 0 200px", "An item stays exactly 200px"],
    ["order: -1", "Moves an item first (the look only)"],
    ["margin-left: auto / margin-top: auto", "Pushes an item right / down"],
  ], { x: CX, y: 1.9, w: CW, colW: [4.15, 3.4], size: 12.5, rowH: 0.44, codeCols: [0] });
}

const EXIT1 = [
  ["Turn a parent into a flex container", "display: flex;"],
  ["Spread items to the two ends", "justify-content: space-between;"],
  ["Line items up down the middle", "align-items: center;"],
  ["Let items wrap onto a new line", "flex-wrap: wrap;"],
  ["Space between items, not the edges", "gap: 16px;"],
  ["Stack the items top to bottom", "flex-direction: column;"],
];
quickFire("Exit Ticket: Quick-fire Flexbox (5 min)", EXIT1, false,
  "A quick-fire round to finish. Teams of 2 or 3 (or on your own), 3 minutes, write the CSS for each job on paper or in the Teams chat. One point per card. Answers on the next slide.\n\nAnswers:\n1. display: flex;\n2. justify-content: space-between;\n3. align-items: center;\n4. flex-wrap: wrap;\n5. gap: 16px; (any size)\n6. flex-direction: column;");
quickFire("Exit Ticket: Answers", EXIT1, true,
  "Go through each one quickly. #2 also accepts space-around or space-evenly if they explain it, and #5 accepts any size.\n\nAnyone who got 6 out of 6 is ready for Grid.");

{
  const s = panelSlide("Class 1 Recap and Next Steps",
    "Next class: CSS Grid (rows and columns at the same time), float (text that wraps around a photo), how to choose between them, and a first look at media queries. Try Its 5 to 8 finish the Harbourside layout.\n\nRemind students to keep their Harbourside page, and to keep their W3Schools screenshots: everything is submitted together after next class.");
  labelBullets(s, [
    { label: "Recap:", text: [
      "display: flex goes on the parent. Its children line up along the main axis.",
      "justify-content works on the main axis, align-items on the cross axis. Both centre = the exact middle.",
      "flex-wrap and gap for rows of cards. flex: 1 shares out the space. order changes the look only.",
      "An auto margin pushes an item away: right in a row, down in a column.",
    ] },
    { label: "Before next class:", text: [
      "Keep your Harbourside page and your W3Schools screenshots. Finish Froggy if you like!",
      "Assignment 3 (10%) is due Sunday, November 1: start sketching your \"My Portfolio\" page.",
    ] },
    { label: "Next class:", text: [
      "CSS Grid, float, picking the right tool, and a page that works on a phone.",
    ] },
  ], { size: 20, sub: 15.5 });
}

// =====================================================================
// WEEK 8, CLASS 2: Grid, float and choosing a layout tool
deck("Week 8, Class 2: CSS Grid and Float", "Grid_and_Float.pptx");
quizNo = 0;

titleSlide("Week 8, Class 2", "CSS Grid, float, and picking the right layout tool for the job.",
  "Welcome back! Last class: Flexbox, for lining things up in one row or one column.\n\nToday: CSS Grid, for rows AND columns at the same time (whole page layouts, galleries), then float, the old layout tool that's still the best way to wrap text around a photo. Then we sort out which tool to use when, and finish the Harbourside page so it works on a phone.\n\nSame routine as last class: pop quizzes all the way through, a predict slide, a spot-the-bug, a sorting game and an exit ticket.");

checkInSlide("Take attendance while everyone checks these four things.\n\nWhile you do, have students git pull (in case anything changed), open their Harbourside page from last class with Live Server, and open playground/grid_playground.html.");

overviewSlide("Week 8, Class 2: Overview", "CSS Grid and Float", [
  "CSS Grid: columns, fr, spanning, named areas, auto-fit",
  "Float: wrapping text around a photo, and the clearfix",
  "Grid, flexbox or float? Picking the right tool",
  "Harbourside, Part 3: Try Its 5 to 8 (a phone layout too), then submit",
], [["Grid", "Rows and columns"], ["Float", "Text around a photo"], ["Choose", "The right tool for the job"]],
  "Rough plan for the two hours:\n- 10 min: check-in, the warm-up quiz, the 100-second video\n- 25 min: grid ideas and Examples 10 to 13, with Pop Quizzes 2 and 3, the predict slide and spot the bug\n- 10 min: Try It 5, and what happened to the Kids card\n- 10 min: auto-fit (Example 14), Pop Quiz 4, Try It 6\n- 15 min: float (Examples 15 and 16), Pop Quiz 5, Try It 7\n- 10 min: comparing the three, and the sorting game\n- 15 min: a responsive layout (Example 17) and Try It 8\n- 15 min: Grid Garden or W3Schools grid work time, then submit\n- 5 min: exit ticket and recap\nAbout 115 minutes, so Grid Garden can be homework if you're short of time.\n\nAnyone who missed Class 1 can copy Try Its 1 to 4 from activity/answer/style.css first (the CSS between the Try It 1 and Try It 5 headings).");

{
  const s = panelSlide("Before We Start",
    "Students carry on with the page they built last class. If someone missed it, they can copy the Try It 1 to 4 CSS out of activity/answer/style.css, under the same headings in their own style.css.\n\nThe Grid Playground works just like the Flexbox one, with one extra: the grid-template-columns box is a text box, so students can type any value they like (try 100px 1fr 100px).");
  txt(s, [{ text: "1.  Pull, just in case", options: { bold: true, fontSize: 18 } }], { x: CX, y: 1.95, w: CW, h: 0.4 });
  codeBox(s, ["git pull"], { x: CX, y: 2.45, w: 5.2, h: 0.7, maxSize: 16 });
  txt(s, "Today: examples 10–17, and playground/grid_playground.html.", { x: CX, y: 3.25, w: CW, h: 0.5, fontSize: 15, color: C.dark });
  txt(s, [{ text: "2.  Open your Harbourside page from Class 1", options: { bold: true, fontSize: 18 } }], { x: CX, y: 3.95, w: CW, h: 0.4 });
  txt(s, "Try Its 5 to 8 build on it. Missed Class 1? Copy Try Its 1–4 from activity/answer/style.css.", { x: CX, y: 4.4, w: CW, h: 0.8, fontSize: 15, color: C.dark });
  txt(s, [{ text: "3.  Have your W3Schools screenshots handy", options: { bold: true, fontSize: 18 } }], { x: CX, y: 5.4, w: CW, h: 0.4 });
  txt(s, "You'll add the grid ones today and submit everything at the end.", { x: CX, y: 5.85, w: CW, h: 0.8, fontSize: 15, color: C.dark });
}

quiz({
  kicker: "WARM-UP",
  q: "Flexbox is called \"one-dimensional\". What does that mean?",
  options: ["It only works on one element per page", "It lays items out along one axis at a time: a row OR a column", "It only works left to right", "Only one item can grow"], answer: 1,
  why: ["Flexbox handles one line of items at a time. When they wrap, each new line is sized on its own.",
    "Remember the Kids card in Example 09, alone on its line and stretched across?",
    "Grid is two-dimensional: rows AND columns line up together. That's today."],
  notes: "A warm-up from last class, and a lead-in to today. Same routine: 30 seconds, then fingers or the Teams chat.",
  revealNotes: "Answer: B.\n\nThis is the key difference between flexbox and grid, and it's today's big idea. Flexbox: content first, one line at a time; each item sizes itself and the line shares out the space. Grid: layout first; you draw the rows and columns, then place the items in them.",
});

videoSlide({ title: "Grid in 100 Seconds", label: "CSS Grid in 100 Seconds", url: "https://youtu.be/uuOXPWCh-6o",
  watchFor: ["display: grid", "grid-template-columns", "fr", "gap", "grid-column", "grid-template-areas"],
  notes: "The matching video from last class's series. Click the dark box (or the link) to open it. About two minutes.\n\nAgain, it's fast. The goal is to hear the words once; by the end of class they'll all be familiar." });

dividerSlide("CSS Grid", "Rows and columns   ·   fr   ·   spanning   ·   named areas   ·   auto-fit",
  "CSS Grid is the newest of the three layout tools (all browsers since 2017). It's the first CSS tool that was actually designed for page layout: you draw a grid of rows and columns, then place things in it.\n\nWhere flexbox starts from the content ('here are some items, share out the space'), grid starts from the layout ('here are my columns, put the items in them').");

gridLinesSlide("Grid has a few more words than flexbox.\n\nThe grid container is the parent with display: grid; its children are grid items, just like flexbox.\n\nLines are the most important idea. The edges between columns (and rows) are numbered from 1, starting at the left (and the top). A grid with 3 columns has 4 column lines. Placing items uses line numbers, so the orange item here goes from column line 2 to column line 4: two columns wide.\n\nA track is a single column or row. grid-template-columns sets the column tracks; rows are usually made automatically as needed.\n\nA cell is one box; an area is any rectangle of cells (one or more). Named areas are coming in Example 13.");

codeSlide({
  kicker: "GRID 1 · EXAMPLE 10", title: "grid-template-columns", file: "10_grid_columns_example.css",
  lines: between(cssOf("10_grid_columns_example.html"), "/* The grid container"), result: SH.ex10, size: 13,
  bullets: [
    ["You set the ", ["columns", "b"], "; the grid adds ", ["rows", "b"], " as it needs them."],
    [["fr", "c"], " = a fraction (share) of the free space, like ", ["flex: 1", "c"], "."],
    [["repeat(4, 1fr)", "c"], " is short for ", ["1fr 1fr 1fr 1fr", "c"], "."],
  ],
  notes: "Open 10_grid_columns_example.html.\n\nThree grids, five items each, only grid-template-columns changes.\n\n1. Fixed 150px columns: they never change size, so there's spare space at the right. Items 4 and 5 go onto a second row, automatically.\n2. 1fr 2fr 1fr: the space is shared out in 4 parts, the middle column gets 2. Same maths as flex: 2 last class. Resize the window: the columns grow and shrink together, and they stay lined up from row to row (unlike flexbox!).\n3. repeat(4, 1fr): four equal columns.\n\ngap works exactly as in flexbox, between rows and columns. You'll still see grid-gap in older code: it's the old name for the same thing.",
});

quiz({
  q: "`grid-template-columns: repeat(3, 1fr)` and the grid has 7 items. How many rows does it make?",
  options: ["1", "2", "3", "7"], answer: 2,
  why: ["You set the columns; the grid makes as many rows as it needs.", "7 items, 3 per row: two full rows, and item 7 on its own in row 3.",
    "Item 7 sits in the first column of row 3, just like items 4 and 5 in Example 10."],
  notes: "A quick check of 'you set the columns, the rows come for free'.",
  revealNotes: "Answer: C, 3 rows.\n\nThe rows the grid makes on its own are called the implicit grid. You can set their height with grid-auto-rows (Example 11 uses grid-auto-rows: 64px), or write grid-template-rows to set them yourself.",
});

codeSlide({
  kicker: "GRID 2 · EXAMPLE 11", title: "Spanning Rows and Columns", file: "11_grid_span_example.css",
  lines: between(cssOf("11_grid_span_example.html"), "/* Four columns"), result: SH.ex11, size: 13,
  bullets: [
    [["span 2", "c"], " = two tracks wide (or tall), from wherever the item lands."],
    [["2 / 4", "c"], " = from line 2 to line 4. ", ["1 / -1", "c"], " = the whole width."],
  ],
  notes: "Open 11_grid_span_example.html.\n\nTwo ways to make an item bigger:\n- span: 'take up 2 columns (or rows), starting wherever you'd normally go'. A and C.\n- line numbers: 'start at line 2 and end at line 4'. E. Notice the grid fills in around it: D slots into column 1 of row 2, before E.\n\n-1 means the last line, counting from the end. So grid-column: 1 / -1 is the full width, however many columns there are. Very handy for a heading or a banner across a grid (F).\n\nC spans two rows, so the grid flows the other items around it. That's two dimensions at once, something flexbox can't do.\n\nTry it live: change A to span 3 and watch the others move to make room.",
});

quiz({
  q: "An item has `grid-column: 1 / 3`. How many columns wide is it?",
  options: ["1", "2", "3", "All of them"], answer: 1,
  why: ["The numbers are LINES, not columns. Line 1 is the left edge of column 1.", "From line 1 to line 3 crosses 2 columns.",
    "The same as `grid-column: span 2`, when the item starts in column 1."],
  notes: "THE classic grid trap. Expect a lot of 3s.",
  revealNotes: "Answer: B, 2 columns.\n\nGo back to the Key Ideas diagram if people are unsure: lines are the edges, numbered from 1, and a 3-column grid has 4 lines.\n\nIt's also the first game mechanic in Grid Garden, later today.",
});

predict({
  title: "Where Does the Sidebar Go?", file: "12_grid_layout_example.css", size: 13,
  lines: between(cssOf("12_grid_layout_example.html"), "/* Two columns").concat(["", "<div class=\"grid-container\">", "   <header>Header</header>", "   <nav>…</nav>", "   <main>Main Content</main>", "   <aside>Sidebar</aside>", "   <footer>Footer</footer>", "</div>"]),
  img: SH.ex12,
  prompt: [["The \"Practical Example\" from last year's slides: ", ["two columns", "b"], ", five items."], ["Where do the nav, main, ", ["aside", "c"], " and footer land? Sketch it!"]],
  bullets: [["Items fill the grid ", ["in order", "b"], ", one free cell at a time."], ["Nav and main fill row 2, so the ", ["aside", "c"], " starts row 3, in column 1."], ["The fix: name the areas (next)."]],
  notes: "This is the code from the Practical Example slide in the original Week 8 deck, word for word (gap was grid-gap). Ask students to sketch where each part ends up. Most people expect a nice three-part layout with the sidebar on the right.\n\nGive them a minute, have a couple of people describe their sketch, then show the answer.",
  revealNotes: "The sidebar ends up on its own row, under the nav, with an empty space beside it.\n\nWhy: the grid has 2 columns. The header spans both (1 / 3). Nav and main fill the two cells of row 2. The aside is next in the HTML, so it goes in the next free cell: row 3, column 1. Then the footer spans both columns in row 4.\n\nThis is auto-placement: without instructions, grid items go into the next free cell, in HTML order. The fix is to tell each part where to go. The neatest way is named areas: Example 13.\n\nOpen 12_grid_layout_example.html to see it in the browser.",
});

codeSlide({
  kicker: "GRID 3 · EXAMPLE 13", title: "Named Areas", file: "13_grid_areas_example.css",
  lines: between(cssOf("13_grid_areas_example.html"), "/* 1.", "/* Flexbox"), result: SH.ex13, size: 13.5,
  bullets: [
    [["grid-template-areas", "c"], ": draw the layout with names, one string per row."],
    [["grid-area: main", "c"], " puts ", ["main", "c"], " in the area called main."],
    ["Same HTML as Example 12. Only the CSS changed."],
  ],
  notes: "Open 13_grid_areas_example.html beside Example 12. The HTML is identical.\n\ngrid-template-areas is like drawing the layout in your CSS file. Each string is a row; each word is a column. Writing a name twice (header header header) makes that area span those cells. The extra spaces in 'nav    main   aside' are just to line the words up for us humans.\n\nThen each element is placed with grid-area and its name. The names are made up by you: they don't have to match the element names, but it's easier if they do.\n\nThe rules: every row needs the same number of words, and each area has to be a rectangle (no L shapes). Use a . for an empty cell.\n\nThis is how most real page layouts are done today, and it makes the phone layout (Example 17) easy: just redraw the areas.",
});

{
  const img = snippet("bugAreas",
    "<div class=\"page\"><header>header</header><nav>nav</nav><main>main</main><footer>footer</footer></div>",
    ".page { display: grid; grid-template-columns: 1fr 3fr; grid-template-areas: \"header header\" \"nav main\" \"footer footer\"; gap: 6px; }\n.page > * { background: #E1E8EB; border: 2px solid #12304A; padding: 10px; font-weight: bold; }\nmain { background: #FBE9DF; border-color: #B5451B; }\nheader { grid-area: header; } nav { grid-area: nav; } main { grid-area: content; } footer { grid-area: footer; }",
    600, 240);
  quiz({
    kicker: "SPOT THE BUG",
    q: "The main content ends up in a strange spot, outside the layout. Why?",
    code: [".page {", "   display: grid;", "   grid-template-columns: 1fr 3fr;", "   grid-template-areas:", "      \"header header\"", "      \"nav    main\"", "      \"footer footer\";", "}", "", "header { grid-area: header; }", "nav    { grid-area: nav; }", "main   { grid-area: content; }", "footer { grid-area: footer; }"],
    revealImg: img,
    options: ["The rows need commas between them", "`content` doesn't match any name in the template", "The template needs a third column", "`main` can't be a grid item"], answer: 1,
    why: ["You make up the area names, and they have to match exactly.", "There's no area called `content`, so the browser adds an extra row and column and puts `main` there.",
      "Fix: `grid-area: main;` (or rename the area in the template)."],
    notes: "Spot the bug: give them 45 seconds to read the code.",
    revealNotes: "Answer: B.\n\nThe screenshot is what the browser does: it can't find an area called content, so it puts main in a new row and column after the ones you drew, bottom right, and the empty 'main' cell is left blank.\n\nNo error message, of course. DevTools helps here: in the Elements panel, click the 'grid' badge next to the container and the browser draws the grid lines and area names over the page.",
  });
}

trySlide({
  n: 5, title: "A Page Layout with Grid", time: "8 min",
  steps: [
    ["Open your Harbourside page", "From last class, with Live Server."],
    ["Type the CSS on the right", ["In style.css, under ", ["Try It 5", "b"], ". Save."]],
    ["Two columns!", "Menu on the left, hours on the right. Make the window wide to see it."],
    ["Look at the menu cards…", "Something odd happened to one of them. Don't fix it yet!"],
  ],
  code: tryCode(5),
  tip: [[["align-items: start", "c"], " stops the hours box stretching as tall as the menu."]],
  notes: "main becomes a grid with two columns (2fr and 1fr) and three named areas. The intro spans both columns; the menu and the hours sit side by side underneath.\n\nmax-width goes up from 900px to 1100px so there's room for two columns.\n\nalign-items: start: grid items stretch to the height of their row by default (like flex items). Without this, the grey hours box would be as tall as the whole menu column. Try taking it out.\n\nThe odd thing: the menu column is narrower now, so only three cards fit on a line, and the Kids card is alone on the second line, stretched all the way across. Next slide.",
});

{
  const s = activitySlide("Pause and Think: What Happened to the Kids Card?",
    "Ask the class first: why is the Kids card so wide? Give them a minute with a neighbour.\n\nThe explanation: the menu column is narrower now (two thirds of the page), so only three 180px cards fit on a line. The Kids card wraps onto a line of its own. Flexbox sizes each line separately, and flex: 1 1 180px says 'grow to fill the line', so it grows to fill the whole line.\n\nThat's flexbox being one-dimensional (the warm-up quiz). The fix is to use a two-dimensional tool for the cards: a grid, where every row uses the same columns. That's Try It 6, and it's only two lines.");
  const iw = 6.9;
  framedImage(s, SH.ti5, 0.75, 1.75, iw, "AFTER TRY IT 5");
  const x = 0.75 + iw + 0.3, w = W - x - 0.6;
  card(s, { x, y: 1.45, w, h: 5.5 });
  txt(s, "Why so wide?", { x: x + 0.25, y: 1.65, w: w - 0.5, h: 0.45, fontSize: 18, bold: true, color: C.black });
  richBullets(s, [
    ["The menu column is narrower now: only ", ["three", "b"], " cards fit on a line."],
    ["Flexbox sizes ", ["each line on its own", "b"], ". The Kids card is alone on line 2..."],
    ["...and ", ["flex: 1 1 180px", "c"], " says \"grow to fill the line\". So it does."],
    ["We want columns that line up on ", ["every", "b"], " row. That's a grid!"],
  ], { x: x + 0.25, y: 2.25, w: w - 0.5, h: 4.5, size: 15, space: 14 });
}

codeSlide({
  kicker: "GRID 4 · EXAMPLE 14", title: "auto-fit: No Media Query Needed", file: "14_grid_auto_fit_example.css",
  lines: between(cssOf("14_grid_auto_fit_example.html"), "/* As many"), size: 13.5,
  frames: [{ img: SH.ex14, label: "900PX WIDE: FIVE COLUMNS" }, { img: SH.ex14p, label: "400PX: TWO" }], cols: 2, maxFrameH: 3.6,
  bullets: [
    [["minmax(160px, 1fr)", "c"], ": every column is at least 160px, and they share any spare space."],
    [["auto-fit", "c"], ": as many columns as fit. The browser does the maths."],
  ],
  notes: "Open 14_grid_auto_fit_example.html and resize the window slowly. The columns drop from five to four to three to two to one, and every tile in every row stays lined up.\n\nRead the line from the inside out:\n- minmax(160px, 1fr): a column can't be narrower than 160px, and can grow to a 1fr share.\n- repeat(auto-fit, ...): make as many of those columns as fit in the width.\n\nCompare with Example 09 (flexbox cards): there, a card alone on the last line stretched across. Here, every row has the same columns, so the last row's tiles stay the same size. This one line is probably the most useful thing in today's class.\n\n(auto-fill is almost the same: the difference only shows when there are fewer items than columns. auto-fit stretches the items to fill the row; auto-fill keeps empty columns.)",
});

quiz({
  q: "`repeat(auto-fit, minmax(200px, 1fr))` in a grid 700px wide, with no gap. How many columns?",
  options: ["2", "3", "4", "7"], answer: 1,
  why: ["How many 200px columns fit in 700px? Three (3 × 200 = 600). A fourth would need 800px.", "Then `1fr` shares out the spare 100px: three columns of about 233px.",
    "Make the window 800px or wider and a fourth column appears."],
  notes: "More maths, but easy maths. 4 (C) is what you get by rounding 3.5 up.",
  revealNotes: "Answer: B, 3 columns.\n\nWith a gap, there's a little less room: three 200px columns and two 16px gaps need 632px, still fine in 700px. The browser always does this sum for you, every time the window changes size.",
});

trySlide({
  n: 6, title: "Menu Cards with Grid", time: "5 min",
  steps: [
    ["Type the CSS on the right", ["Under ", ["Try It 6", "b"], ". Save."]],
    ["The Kids card is fixed", "It stays in the first column, the same width as the others."],
    ["Still lined up?", ["The prices still line up, and the special is still first: ", ["order", "c"], " works in grid too."]],
    ["Make the window narrow", "3, 2, then 1 column. No media query."],
  ],
  code: tryCode(6),
  tip: [["Try It 2's ", ["flex: 1 1 180px", "c"], " is now ignored: grid items are sized by the grid, not by ", ["flex", "c"], "."]],
  notes: "Two lines replace the flexbox from Try It 2. The later .cards rule wins for display (grid beats flex because it comes later), and gap: 16px from Try It 2 still applies, because gap works in both.\n\nNow every row uses the same three columns, so the Kids card sits in column 1 at the same width as the others.\n\nWhat still works from Class 1:\n- Try It 3: each card is still a column flexbox with the price at the bottom. Grid items stretch to the height of their row, like flex items.\n- Try It 4: order: -1 still puts the special first. order works on grid items too.\n\nThe flex: 1 1 180px line from Try It 2 is now ignored, because the articles aren't flex items any more. Students can leave it or delete it.",
});

dividerSlide("Float", "Text around a photo   ·   the clearfix   ·   the old way to make columns",
  "Float is the oldest of the three. It was made for one job: letting text flow around an image, like in a newspaper or a magazine. In the 2000s and early 2010s, people also used it to build whole page layouts, because there wasn't anything better. That took a lot of hacks.\n\nToday, use grid and flexbox for layout, and float for its original job. But you'll still meet floats in older code, so it's worth knowing how they behave.");

codeSlide({
  kicker: "FLOAT 1 · EXAMPLE 15", title: "Wrapping Text Around a Photo", file: "15_float_text_wrap_example.css",
  lines: between(cssOf("15_float_text_wrap_example.html"), "/* 1."), result: SH.ex15, size: 13.5,
  bullets: [
    [["float: left", "c"], ": the photo moves to the left, and the text flows around its right side."],
    ["Give it a ", ["width", "b"], " and a ", ["margin", "b"], " on the side the text is on."],
  ],
  notes: "Open 15_float_text_wrap_example.html and resize the window.\n\nThe photo floats to the left of its paragraph; the text wraps around it, and when the text gets below the photo it goes back to full width. That's the one thing flexbox and grid can't do: they put things side by side in boxes, they don't wrap text around things.\n\nThe pull quote floats to the right, the same way. Pull quotes are a magazine idea: a short, eye-catching quote from the article.\n\nThe margins matter: margin-right on the left-floated photo (and margin-left on the right-floated quote) keeps the text from touching them.\n\nWidth as a percentage keeps it responsive: the photo is always 45% of the article.",
});

codeSlide({
  kicker: "FLOAT 2 · EXAMPLE 16", title: "Floats and the Clearfix", file: "16_float_clearfix_example.css",
  lines: between(cssOf("16_float_clearfix_example.html"), "/* 1."), result: SH.ex16, size: 12.5,
  bullets: [
    ["Floats don't count towards their parent's height. Only floats inside? The box ", ["collapses", "b"], "."],
    ["The clearfix hack, or today's one-liner: ", ["display: flow-root", "c"], "."],
  ],
  notes: "Open 16_float_clearfix_example.html. This is the 'Float for layouts' and 'Clearfix hack' code from the original deck, side by side.\n\nDemo 1: two floated 50% columns inside a box with an orange border. The box has nothing else inside, and floats are taken out of the normal flow, so the box has a height of 0. Its border is the thin orange line at the top, and the columns hang out below it. The next heading even tries to wrap around them.\n\nDemo 2: the clearfix. ::after adds an empty element at the end of the box (a pseudo-element from Class 9), and clear: both makes it go below both floats. The box has to stretch down to hold it, so it wraps the columns. You'll see this in a lot of older code.\n\nDemo 3: display: flow-root does the same job in one line, with no hack. Use this one.\n\nAnd the third option: don't use floats for columns at all. Two columns are one line of grid or flexbox, and those containers always wrap their items.\n\nbox-sizing: border-box at the top: without it, 50% plus the padding is more than half, and the right column drops below the left one.",
});

quiz({
  q: "A `div` holds two floated columns and nothing else. Its border shows up as a thin line above them. Why?",
  options: ["The border needs a width", "Floats are taken out of the normal flow, so the div thinks it's empty", "The floats need a z-index", "The columns need display: block"], answer: 1,
  why: ["A floated element doesn't add to its parent's height. Only floats inside = a height of 0.", "Fix it with `display: flow-root` on the parent (or the older clearfix hack).",
    "Flex and grid containers don't have this problem: they always wrap their items."],
  notes: "Straight after Example 16, so this should go well. It checks that students know WHY the clearfix is needed, not just what it is.",
  revealNotes: "Answer: B.\n\nThis is the number one reason people moved from floats to flexbox and grid for layout. And floats don't mix with them: float: left on a flex or grid item does nothing at all.",
});

trySlide({
  n: 7, title: "Float the Photo", time: "5 min",
  steps: [
    ["Type the CSS on the right", ["Under ", ["Try It 7", "b"], ". Save."]],
    ["The text wraps", "The intro paragraphs flow around the left side of the photo."],
    ["Play with it", ["Try ", ["float: left", "c"], " (and swap the margins). Or change the width to ", ["30%", "c"], "."]],
  ],
  code: tryCode(7),
  tip: [["No clearfix needed: a grid item always wraps its own floats. On a page with no grid, you'd add ", ["display: flow-root", "c"], "."]],
  notes: "The photo in the intro floats right, at 45% of the intro's width, and the paragraphs wrap around it.\n\nThe figure is floated, not the img, so the caption goes with the photo.\n\nWhy no collapsing box, after all that? The intro section is a grid item (Try It 5), and a grid item always contains its own floats. If the text is shorter than the photo, the section just grows to fit the photo.\n\nIf you swap to float: left, swap the margin too: margin: 0 20px 12px 0, so the space is on the right, where the text is.",
});

{
  const s = panelSlide("Grid, Flexbox or Float?",
    "This is the comparison from the original deck (Comparing CSS Grid, Flexbox and Float, and the Flexbox and Grid advantages slide), as one picture.\n\nFlexbox: one dimension. Great for a row of things that share out space: nav bars, toolbars, buttons, the inside of a card, and centring.\nGrid: two dimensions. Great when things need to line up in rows AND columns: page layouts, galleries, card grids, forms.\nFloat: text wrapping around something. And legacy code.\n\nThey're not rivals: real pages use all three. Example 17 is grid for the page, flexbox for the nav, and float for the photo.\n\nOne rule from the best practices: don't mix float with flex or grid on the same items. Floats are ignored on flex and grid items, which is confusing to debug.\n\nPositioning (position: absolute) is the fourth tool, from the dropdown menu last week. It's for things that sit on top of the page (dropdowns, badges, pop-ups), not for layout: absolute elements don't take up space, so they don't adapt to the screen.");
  const tools = [
    ["Flexbox", C.orange, "One dimension: a row OR a column", "Nav bars, buttons, centring, the inside of a card"],
    ["Grid", C.teal, "Two dimensions: rows AND columns", "Page layouts, galleries, card grids"],
    ["Float", C.purple, "Text wraps around it", "A photo in an article. Old layouts"],
  ];
  const cw = 2.4, gap = 0.17;
  tools.forEach(([t, col, d, use], i) => {
    const x = CX + i * (cw + gap), y = 1.95;
    card(s, { x, y, w: cw, h: 3.9, border: col, bw: 1.5 });
    const mx = x + 0.25, my = y + 0.25, mw = cw - 0.5;
    if (i === 0) {
      [0, 1, 2].forEach((j) => s.addShape(pres.shapes.RECTANGLE, { x: mx + j * 0.65, y: my + 0.15, w: 0.55, h: 0.75, fill: { color: col }, line: { color: col } }));
      s.addShape(pres.shapes.LINE, { x: mx, y: my + 1.1, w: mw, h: 0, line: { color: C.dark, width: 2, endArrowType: "triangle" } });
    } else if (i === 1) {
      for (let r = 0; r < 2; r++) for (let c = 0; c < 3; c++)
        s.addShape(pres.shapes.RECTANGLE, { x: mx + c * 0.65, y: my + r * 0.6, w: 0.55, h: 0.5, fill: { color: col }, line: { color: col } });
    } else {
      s.addShape(pres.shapes.RECTANGLE, { x: mx, y: my, w: 0.75, h: 0.65, fill: { color: col }, line: { color: col } });
      [0, 1, 2, 3, 4].forEach((j) => {
        const lx = j < 3 ? mx + 0.9 : mx, lw = j < 3 ? mw - 0.9 : mw;
        s.addShape(pres.shapes.RECTANGLE, { x: lx, y: my + 0.05 + j * 0.24, w: lw, h: 0.1, fill: { color: "B0B8BC" }, line: { color: "B0B8BC" } });
      });
    }
    txt(s, t, { x: x + 0.2, y: y + 1.55, w: cw - 0.4, h: 0.45, fontSize: 19, bold: true, color: col });
    txt(s, d, { x: x + 0.2, y: y + 2.0, w: cw - 0.4, h: 0.75, fontSize: 13.5, bold: true, color: C.black, valign: "top" });
    txt(s, "Use for: " + use, { x: x + 0.2, y: y + 2.8, w: cw - 0.4, h: 0.95, fontSize: 12.5, italic: true, color: C.grey, valign: "top" });
  });
  richBullets(s, [["They work together: ", ["grid", "b"], " for the page, ", ["flexbox", "b"], " inside its parts, ", ["float", "b"], " for a photo in the text."],
    ["Don't mix float with flex or grid on the same items: floats are ignored on flex and grid items."]],
    { x: CX, y: 6.0, w: CW, h: 1.0, size: 14 });
}

const SORT = [
  ["Logo on the left, links on the right, in a header", "Flexbox", "One row: space-between, or margin-left: auto."],
  ["A whole page: header, sidebar, main and footer", "Grid", "Rows and columns together: grid-template-areas."],
  ["Text that wraps around a photo in an article", "Float", "The one job float is still the best at."],
  ["A photo gallery where the columns line up", "Grid", "Every row has the same columns: repeat(auto-fit, …)."],
  ["A message in the exact middle of a box", "Flexbox", "justify-content + align-items: center. (Grid can too.)"],
  ["Each card's button at the bottom of the card", "Flexbox", "A column flexbox + margin-top: auto."],
];
sortGame("Game: Grid, Flexbox or Float? (5 min)", SORT, false,
  "Teams of 2 or 3. For each job, pick the best tool: Grid, Flexbox or Float. Write down six answers (1 to 6), or post them in the Teams chat. 3 minutes, then the answers are on the next slide. One point each.\n\nAnswers:\n1. Flexbox\n2. Grid\n3. Float\n4. Grid\n5. Flexbox (grid also works: display: grid; place-items: center. Give the point.)\n6. Flexbox");
sortGame("Grid, Flexbox or Float: Answers", SORT, true,
  "Go through them and accept a good argument for another answer. Number 5 especially: grid's place-items: center centres too. And a header could be a grid with two columns (1fr auto), though flexbox is more usual.\n\nThe rule of thumb: one direction → flexbox. Two directions, or things that must line up → grid. Text wrapping → float.");

codeSlide({
  kicker: "LAYOUT · EXAMPLE 17", title: "A Layout for Every Screen", file: "17_responsive_layout_example.css",
  lines: between(cssOf("17_responsive_layout_example.html"), "/* 1. Phones", "header { grid-area").concat([""], between(cssOf("17_responsive_layout_example.html"), "/* 3.")), size: 12,
  frames: [{ img: SH.ex17, label: "WIDE" }, { img: SH.ex17p, label: "A PHONE (360PX)" }], cols: 2, maxFrameH: 3.75,
  bullets: [
    ["Phones first: ", ["one column", "b"], ". Then ", ["@media (min-width: 700px)", "c"], " redraws the grid."],
    ["Grid for the page, flexbox for the links, float for the photo."],
  ],
  notes: "Open 17_responsive_layout_example.html and make the window narrow, then wide. At 700px the whole layout changes.\n\nA media query is a set of rules that only apply when a condition is true. @media (min-width: 700px) { ... } means 'only when the window is at least 700px wide'.\n\nThe idea is 'mobile first': write the simple one-column layout for phones first, then add the extra columns for bigger screens. Inside the media query, the grid gets three columns and new areas, the links turn into a column, and the photo floats.\n\nThe HTML never changes: only the CSS. That's the point of keeping them separate.\n\nThis is the start of responsive design, which is next class's topic: more about media queries, breakpoints and images.\n\nThe code is long, so it's in a small font: students should open the .css file to read it properly.",
});

trySlide({
  n: 8, title: "On a Phone", time: "8 min",
  steps: [
    ["Type the CSS on the right", ["At the ", ["very end", "b"], " of style.css, under Try It 8. Save."]],
    ["Be a phone", ["F12, then ", ["Ctrl + Shift + M", "b"], ". Pick a phone, or drag to 360px."]],
    ["One column", "Intro, menu, hours stack up. The photo stops floating."],
    ["Take your two screenshots", "One full width, one phone-sized. Then submit!"],
  ],
  code: tryCode(8),
  tip: [["Why last? If two rules tie, the ", ["later", "b"], " one wins. A media query doesn't make a rule stronger."]],
  notes: "This one uses max-width instead of Example 17's min-width: the page is already built for wide screens, so we add rules for narrow ones. Both are fine; next class covers the difference properly.\n\nInside the media query, at 700px wide or less:\n- main goes back to one column, with the three areas stacked.\n- The photo stops floating and goes full width (width: auto).\n- The big h1 gets smaller, so 'Harbourside Fish & Chips' fits.\n- The nav links wrap as whole links instead of squashing.\n\nWhy it has to be last in the file: the media query's rules have the same specificity as the rules they replace (main, .intro figure ...). If two rules tie, the later one wins. If the media query were above Try It 5, Try It 5's rules would win, even on a phone. A media query doesn't add any specificity.\n\nThe DevTools device toolbar (Ctrl + Shift + M) is the easiest way to test. Pick 'iPhone SE' or similar, or drag the width.",
});

{
  const s = panelSlide("Best Practices for Layout",
    "From the original deck's Best Practices slide, with a little more detail.\n\nConsistency: the same layout (header, nav, footer) on every page, from one shared stylesheet. That's Assignment 2's lesson.\n\nResponsiveness: always check at phone width (360px) in the device toolbar. A layout that only works on your laptop isn't finished.\n\nThe right tool: flexbox for one direction, grid for two, float only for wrapping text. Don't put float on flex or grid items: it's ignored.\n\nComments: a short comment above each layout rule ('Two columns: menu and hours') saves a lot of time later, for you and your teammates. Every example this week has them.\n\nAnd separate files: HTML for content, one CSS file for the look. Every example this week does that.");
  labelBullets(s, [
    { label: "Consistency:", text: "The same layout on every page, from one shared stylesheet." },
    { label: "Responsiveness:", text: "Check every page at phone width (Ctrl + Shift + M). Media queries do the rest." },
    { label: "One job, one tool:", text: "Flexbox in one direction, grid in two, float for text wrapping. Don't mix float with flex or grid on the same items." },
    { label: "Comments:", text: "Say what each layout rule is for: /* Two columns: menu and hours */" },
    { label: "Separate files:", text: "HTML for the content, an external stylesheet for the look." },
  ], { size: 18, sub: 15 });
}

playgroundSlide({ title: "Your Turn to Drive: The Grid Playground", img: SH.playGrid, file: "playground/grid_playground.html",
  challenges: ["Make four equal columns.", "Make box 2 go all the way across.", "Make box 2 two rows tall, and two columns wide.", "Type your own columns: try `100px 1fr 100px`. What happens when you resize?"],
  notes: "If there's time before work time, do this with the class driving (or let students try it themselves).\n\nAnswers:\n1. grid-template-columns: repeat(4, 1fr).\n2. Item 2: grid-column: 1 / -1.\n3. Item 2: grid-column: span 2 and grid-row: span 2 (the screenshot shows this, with repeat(4, 1fr)).\n4. The two 100px columns stay the same and the middle one stretches: the classic 'sidebars and content' layout.\n\nThe columns box is a text box, so anything goes: try repeat(auto-fit, minmax(150px, 1fr)) and resize the window, or a typo, to see that the browser just ignores a value it doesn't understand." });

gameSlide({
  title: "Game Time: Grid Garden (10 min)", name: "Grid Garden", url: "https://cssgridgarden.com/",
  aim: "Water the garden: get as far as you can in 10 minutes.",
  how: ["Write CSS to water the right patch of the garden (and nothing else).", "It starts with `grid-column-start` and `grid-column-end`, then `grid-column`, `grid-row` and `grid-area`.",
    "Line numbers count from 1, and negative numbers count from the end: just like Example 11."],
  words: [["grid-column-start", "The line a patch starts on"], ["grid-column-end", "The line it ends on"], ["grid-column: 2 / 4", "Both at once"],
    ["span 2", "Two tracks wide"], ["grid-row", "The same, for rows"], ["grid-area", "Rows and columns in one line"]],
  notes: "Grid Garden is Flexbox Froggy's sibling, from the same people (Codepip). Free, no sign-up. Each level is a garden with a patch to water; you write grid placement CSS to cover exactly that patch.\n\nThe early levels are all about line numbers: a great follow-up to Pop Quiz 3 (grid-column: 1 / 3 is two columns, not three).\n\nIf time is short, make it homework and go straight to work time.",
});

workTimeSlide({
  title: "Work Time: W3Schools Grid (15 min)",
  links: [["CSS Grid Layout Module", "https://www.w3schools.com/css/css_grid.asp"], ["CSS Grid Container", "https://www.w3schools.com/css/css_grid_container.asp"],
    ["CSS Grid Item", "https://www.w3schools.com/css/css_grid_item.asp"]],
  notes: "Same as last class, now with the Grid pages: read each page, open the 'Try it Yourself' examples, change something, Run, and screenshot the code and the result.\n\nStudents need at least 5 screenshots in total across both classes (flexbox and grid together).\n\nGood ones to change: the grid-template-columns values, gap, and the grid-column / grid-row of one item.",
});

{
  const s = activitySlide("Submit: Harbourside, Part 3 and Your W3Schools Screenshots",
    "Everything from both classes goes on Brightspace in one submission (participation):\n1. A screenshot of the finished Harbourside page at full width.\n2. A screenshot at phone width (the device toolbar, Ctrl + Shift + M).\n3. At least five W3Schools 'Try it Yourself' screenshots, each one with code the student changed.\n\nThe Harbourside screenshots are this week's version of the original deck's class activity ('create a web page layout using CSS Grid or Flexbox, and experiment with float'): the page uses all three.\n\nStudents who finish early: the 'Make it yours' ideas. A different auto-fit minimum (try 240px), a third column for the hours box on very wide screens, the intro photo floated left, or rounded corners on the hours table.");
  const ph = 4.6;
  framedImage(s, SH.c2, 0.75, 1.75, ph / SH.c2.ratio, "1. FULL WIDTH");
  const pw = ph / SH.c2p.ratio;
  framedImage(s, SH.c2p, 0.75 + ph / SH.c2.ratio + 0.25, 1.75, pw, "2. PHONE");
  const x = 0.75 + ph / SH.c2.ratio + 0.25 + pw + 0.3, w = W - x - 0.6;
  card(s, { x, y: 1.45, w, h: 5.5, fill: C.dark, border: C.dark });
  txt(s, "On Brightspace", { x: x + 0.25, y: 1.7, w: w - 0.5, h: 0.45, fontSize: 18, bold: true, color: C.white });
  txt(s, [
    { text: "1. ", options: { bold: true, color: "F2C6AC" } }, { text: "Harbourside, full width", options: { breakLine: true } },
    { text: "2. ", options: { bold: true, color: "F2C6AC" } }, { text: "Harbourside, phone width (Ctrl + Shift + M)", options: { breakLine: true } },
    { text: "3. ", options: { bold: true, color: "F2C6AC" } }, { text: "5+ W3Schools \"Try it Yourself\" screenshots, with your changes", options: { breakLine: true } },
    { text: " ", options: { fontSize: 8, breakLine: true } },
    { text: "Win + Shift + S", options: { bold: true, color: "F2C6AC", fontFace: MONO } }, { text: " takes a screenshot.", options: { breakLine: true } },
    { text: " ", options: { fontSize: 8, breakLine: true } },
    { text: "Stuck? ", options: { bold: true, color: "F2C6AC" } }, { text: "activity/answer/" },
  ], { x: x + 0.25, y: 2.25, w: w - 0.5, h: 4.5, fontSize: 14, color: "D9D9D9", valign: "top", paraSpaceBefore: 6 });
}

{
  const s = panelSlide("Assignment 3: Show Off Your Layout",
    "Assignment 3 (10%, due Sunday, November 1 at 11:59 pm): a small page, written by hand (no WYSIWYG editor), that shows they can lay things out with Flexbox or Grid. A single page is fine. The suggestion is a 'My Portfolio' page that starts with the courses they're taking this term; they can reuse content from Assignment 2. Hand-in: the HTML and CSS in a ZIP on Brightspace.\n\nThe table maps each rubric row (all four are out of 4) to what we did this week:\n- HTML Structure: semantic elements (header, nav, main, section or article, footer), like Examples 12, 13 and 17. Run the W3C validator (Class 11).\n- CSS Usage: flexbox or grid doing a real layout job. A nav bar (Example 08), cards (09, 14) or a page layout (13). 'Effective' means the layout does something useful, not just display: flex on one div.\n- Responsiveness: 'works well on all devices'. flex-wrap or auto-fit, images with max-width: 100%, the viewport meta tag (it's in every example's head), and ideally one media query (Example 17, Try It 8). Check it at 360px with Ctrl + Shift + M.\n- Code Quality: a comment on each layout rule (like every example this week), one external style.css, and both validators.\n\nA nice idea for the course list: one card per course in a repeat(auto-fit, minmax(200px, 1fr)) grid, with each card a column flexbox. That covers grid, flexbox and responsiveness in about 15 lines.");
  table(s, [
    ["Rubric row", "For a 4", "Like this week's"],
    ["HTML Structure", "Semantic: header, nav, main, section, footer. Valid HTML.", "Examples 12, 13, 17"],
    ["CSS Usage", "Flexbox or grid doing a real job: a nav, cards, a page layout", "Ex 08, 09, 13, 14 · Try Its 1–6"],
    ["Responsiveness", "Wraps or reflows at 360px: auto-fit, flex-wrap, a media query, viewport meta", "Ex 14, 17 · Try It 8"],
    ["Code Quality", "A comment on each layout rule, one external style.css, both validators", "Every example"],
  ], { x: CX, y: 1.9, w: CW, colW: [1.7, 3.65, 2.2], size: 12.5, rowH: 0.62 });
  card(s, { x: CX, y: 5.35, w: CW, h: 1.5, fill: C.dark, border: C.dark });
  richBullets(s, [
    [["Due Sunday, November 1, 11:59 pm", "b"], " · 10% · a ZIP with your HTML and CSS (and images) on Brightspace."],
    ["Idea: a ", ["My Portfolio", "b"], " page with a card for each of your courses, in a ", ["repeat(auto-fit, minmax(200px, 1fr))", "c"], " grid."],
  ], { x: CX + 0.25, y: 5.45, w: CW - 0.5, h: 1.3, size: 14, color: "E6E6E6", codeColor: "F2C6AC", space: 8, valign: "middle" });
}

{
  const s = panelSlide("Cheat Sheet: Grid and Float", "Everything from today in one place. It's in the README too.");
  table(s, [
    ["Code", "What it does"],
    ["display: grid", "The parent becomes a grid container"],
    ["grid-template-columns: 1fr 2fr", "The columns (fr = a share of the space)"],
    ["repeat(3, 1fr)", "Three equal columns"],
    ["repeat(auto-fit, minmax(200px, 1fr))", "As many columns as fit"],
    ["grid-column: 1 / 3   or   span 2", "Line 1 to line 3: two columns wide"],
    ["grid-template-areas + grid-area", "Draw the layout with names"],
    ["float: left | right", "Text wraps around the element"],
    ["display: flow-root", "The parent wraps its floats"],
    ["@media (max-width: 700px) { … }", "Rules for narrow screens only"],
  ], { x: CX, y: 1.9, w: CW, colW: [4.15, 3.4], size: 12.5, rowH: 0.46, codeCols: [0] });
}

const EXIT2 = [
  ["Three equal columns", "grid-template-columns: repeat(3, 1fr);"],
  ["An item all the way across", "grid-column: 1 / -1;"],
  ["Put main in the area called \"main\"", "grid-area: main;"],
  ["As many 200px columns as fit", "repeat(auto-fit, minmax(200px, 1fr))"],
  ["Wrap text around a photo", "float: left;"],
  ["Stop a box collapsing around floats", "display: flow-root;"],
];
quickFire("Exit Ticket: Quick-fire Grid and Float (5 min)", EXIT2, false,
  "Same as last class: teams of 2 or 3, 3 minutes, write the CSS for each job. One point per card.\n\nAnswers:\n1. grid-template-columns: repeat(3, 1fr); (1fr 1fr 1fr is fine too)\n2. grid-column: 1 / -1;\n3. grid-area: main;\n4. grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));\n5. float: left; (or right)\n6. display: flow-root; (the clearfix is fine too)");
quickFire("Exit Ticket: Answers", EXIT2, true,
  "Go through them quickly and accept the alternatives in the notes on the previous slide.");

{
  const s = panelSlide("Class 2 Recap and Next Steps",
    "Next class: advanced responsive design techniques. More media queries, breakpoints, mobile-first CSS, and responsive images. Example 17 and Try It 8 were a first taste.\n\nMake sure the Harbourside screenshots (wide and phone) and the five or more W3Schools screenshots are on Brightspace.");
  labelBullets(s, [
    { label: "Recap:", text: [
      "Grid is two-dimensional: you set the columns, the rows come as needed. fr shares out the space.",
      "Lines are numbered from 1: grid-column: 1 / 3 is two columns. Name areas to draw a layout.",
      "repeat(auto-fit, minmax(…)) makes a responsive grid with no media query.",
      "Float wraps text around a photo. A parent with only floats collapses: display: flow-root.",
      "Flexbox for one direction, grid for two, float for text. They work together.",
    ] },
    { label: "This week:", text: ["Submit the Harbourside screenshots (wide + phone) and 5+ W3Schools screenshots.", "Assignment 3 (10%): due Sunday, November 1 at 11:59 pm."] },
    { label: "Next class:", text: "Advanced responsive design: media queries, breakpoints and images." },
  ], { size: 20, sub: 15.5 });
}

save();
