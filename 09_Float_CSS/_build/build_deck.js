// Builds Float.pptx (Week 9, Class 1), in the same style as the Class 5–11 and Week 8 decks.
// This is the float section of Grid_and_Float.pptx, carried forward from Week 8, Class 2.
// The examples and Harbourside (activity/) are in this folder; Try Its 1 to 6 come over from 08_Flex_Grid_CSS already done.
// To rebuild: cd 09_Float_CSS/_build, run "npm install" once, then "npm run build".
// The "Result" pictures are screenshots of the example pages, taken with Microsoft Edge (headless) each build.
// Code slides read the example files directly, so edit those files and re-run to update the slides.
const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");
const { execFileSync } = require("child_process");

const REPO = path.join(__dirname, "..");
const SCR = __dirname;
const SHOTS = path.join(SCR, "shots");

const FONT = "Aptos";
const MONO = "Consolas";
const UNIT = "Week 9: Responsive Web Design";
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

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "Week 9, Class 1: Float";
pres.author = "Jamie Symonds";
const OUT = path.join(__dirname, "..", "Float.pptx");
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

// The lines from the one that starts with `from` up to (not including) the one that starts with `to`
function between(lines, from, to) {
  const a = lines.findIndex((l) => l.trim().startsWith(from));
  const b = to ? lines.findIndex((l, i) => i > a && l.trim().startsWith(to)) : lines.length;
  if (a < 0 || b < 0) throw new Error(`between(): couldn't find "${from}" / "${to}"`);
  return trim(lines.slice(a, b));
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

const ANSWER_CSS = fs.readFileSync(path.join(REPO, "activity/answer/style.css"), "utf8").replace(/\r/g, "");

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
  ex15: shot("15_float_text_wrap_example.html", "ex15", 800, 470),
  ex16: shot("16_float_clearfix_example.html", "ex16", 600, 400),
};

// =====================================================================
// WEEK 9, CLASS 1: Float (carried forward from Week 8, Class 2)



titleSlide("Week 9, Class 1", "Float: wrapping text around a photo, and the clearfix.",
  "Welcome back! We didn't get to float last class, so we're picking it up today.\n\nFloat is the old layout tool that's still the best way to wrap text around a photo. It's the third layout tool after flexbox and grid.");

checkInSlide("Take attendance while everyone checks these four things.\n\nWhile you do, have students git pull, open activity/index.html in 09_Float_CSS with Live Server (Try Its 1 to 6 are already done), and open 15_float_text_wrap_example.html in the same folder.");

overviewSlide("Week 9, Class 1: Overview", "Float", [
  "Float: wrapping text around a photo",
  "Collapsing parents: the clearfix and display: flow-root",
  "Harbourside, Part 3: Try It 7, float the photo",
], [["Wrap", "Text around a photo"], ["Collapse", "Why the box disappears"], ["Fix", "display: flow-root"]],
  "Carried forward from Week 8, Class 2. Rough plan for this part: about 15 minutes.\n- Examples 15 and 16\n- Pop Quiz 1\n- Try It 7 (5 min)\n\nThe examples and the Harbourside page are in the 09_Float_CSS folder.\n\nTry It 7 builds on Try It 5 (the page layout with grid). The Harbourside page in activity/ has Try Its 1 to 6 already done, so everyone starts from the same place.");

dividerSlide("Float", "Text around a photo   ·   the clearfix   ·   the old way to make columns",
  "Float is the oldest of the three. It was made for one job: letting text flow around an image, like in a newspaper or a magazine. In the 2000s and early 2010s, people also used it to build whole page layouts, because there wasn't anything better. That took a lot of hacks.\n\nToday, use grid and flexbox for layout, and float for its original job. But you'll still meet floats in older code, so it's worth knowing how they behave.");

codeSlide({
  kicker: "FLOAT 1 · EXAMPLE 15", title: "Wrapping Text Around a Photo", file: "09_Float_CSS/15_float_text_wrap_example.css",
  lines: between(cssOf("15_float_text_wrap_example.html"), "/* 1."), result: SH.ex15, size: 13.5,
  bullets: [
    [["float: left", "c"], ": the photo moves to the left, and the text flows around its right side."],
    ["Give it a ", ["width", "b"], " and a ", ["margin", "b"], " on the side the text is on."],
  ],
  notes: "Open 15_float_text_wrap_example.html and resize the window.\n\nThe photo floats to the left of its paragraph; the text wraps around it, and when the text gets below the photo it goes back to full width. That's the one thing flexbox and grid can't do: they put things side by side in boxes, they don't wrap text around things.\n\nThe pull quote floats to the right, the same way. Pull quotes are a magazine idea: a short, eye-catching quote from the article.\n\nThe margins matter: margin-right on the left-floated photo (and margin-left on the right-floated quote) keeps the text from touching them.\n\nWidth as a percentage keeps it responsive: the photo is always 45% of the article.",
});

codeSlide({
  kicker: "FLOAT 2 · EXAMPLE 16", title: "Floats and the Clearfix", file: "09_Float_CSS/16_float_clearfix_example.css",
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
  const s = panelSlide("Float Recap",
    "The two things to remember: float is for wrapping text around something, and a parent with only floats inside collapses unless you give it display: flow-root.\n\nFor page layout, keep using flexbox and grid.");
  labelBullets(s, [
    { label: "Recap:", text: [
      "float: left or right moves an element to the side, and the text flows around it.",
      "Give a float a width, and a margin on the side the text is on.",
      "A parent with only floats collapses. Fix it with display: flow-root (or the older clearfix).",
      "Floats are ignored on flex and grid items. Use flexbox and grid for layout, and float for text.",
    ] },
  ], { size: 20, sub: 15.5 });
}

pres.writeFile({ fileName: OUT }).then(() => console.log("wrote " + OUT));
