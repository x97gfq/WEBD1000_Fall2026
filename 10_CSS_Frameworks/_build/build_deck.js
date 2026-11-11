// Builds CSS_Frameworks_Bootstrap.pptx (Week 10, Classes 1 and 2), in the same style as the Week 8 and Week 9 decks.
// To rebuild: cd 10_CSS_Frameworks/_build, run "npm install" once, then "npm run build".
// The "Result" pictures are screenshots of the example pages, taken with Microsoft Edge (headless) each build.
// Bootstrap loads from a CDN, so the build needs the internet.
// Customizing, Tailwind and Harbourside moved to Week 11: 11_Customizing_Bootstrap_Tailwind/_build/build_deck.js.
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
const UNIT = "Week 10: CSS Frameworks";
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
pres.title = "Week 10: CSS Frameworks and Bootstrap";
pres.author = "Jamie Symonds";
const OUT = path.join(__dirname, "..", "CSS_Frameworks_Bootstrap.pptx");
// One PowerPoint section per class
let SECTION = null;
function section(title) { pres.addSection({ title }); SECTION = title; }
function newSlide() { return SECTION ? pres.addSlide({ sectionTitle: SECTION }) : pres.addSlide(); }

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
// frames: { img, label } screenshots in place of one result: one row, or an array of rows
// placeholder: a screenshot whose size to copy as an empty dashed box (a "predict the layout" question)
function codeSlide({ kicker, kickerColor = C.grey, title, bullets, file, lines, lang = "css", result, resultLabel = "RESULT", placeholder, frames, maxFrameH = 3.3, notes, size = 14, codeTitle, mark }) {
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
    // every picture in a row is the same height, and the row fills the column (scaled down to fit maxFrameH)
    const rows = Array.isArray(frames[0]) ? frames : [frames];
    const gap = 0.15, lab = 0.32;
    const rowH = rows.map((r) => (RW - gap * (r.length - 1)) / r.reduce((a, f) => a + 1 / f.img.ratio, 0));
    const total = rowH.reduce((a, b) => a + b, 0);
    const k = Math.min(1, (maxFrameH - rows.length * lab - (rows.length - 1) * gap) / total);
    let y = 6.85 - (total * k + rows.length * lab + (rows.length - 1) * gap);
    const top = y;
    rows.forEach((r, ri) => {
      const h = rowH[ri] * k;
      let x = RX;
      r.forEach(({ img, label }) => { const w = h / img.ratio; framedImage(s, img, x, y + lab, w, label); x += w + gap; });
      y += lab + h + gap;
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
// label: a title that isn't a numbered Try It (the warm-up, and Your Layout parts 2 and 3)
function trySlide({ n, title, label, time, steps, code, lang = "css", codeTitle = "IN STYLE.CSS", tip, notes, stepGap, lw }) {
  const s = activitySlide(`${label || `Try It ${n}: ${title}`} (${time})`, notes);
  const LW = lw || 6.55;
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
  codeBox(s, code, { x: RX, y: 1.45, w: RW, h: 5.5 - (tip ? tipH + 0.2 : 0), maxSize: 16, fitHeight: true, lang, title: codeTitle });
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

function setupSlide(title, items, notes) {
  const s = newSlide();
  s.background = { color: C.white };
  txt(s, title, { x: 0.75, y: 0.5, w: 11.8, h: 0.7, fontSize: 28, bold: true, color: C.black });
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

// A true 360 px phone width: headless Edge won't make a window that narrow, so the page goes in a 360 px iframe
function phone(file, name, w = 360, h = 640) {
  const src = "file:///" + path.join(REPO, file).replace(/\\/g, "/");
  fs.writeFileSync(path.join(SHOTS, name + ".html"), `<!DOCTYPE html><html><body style="margin: 0"><iframe src="${src}" width="${w}" height="${h}" style="border: 0; display: block"></iframe></body></html>`);
  return shot(path.relative(REPO, path.join(SHOTS, name + ".html")), name, w, h);
}

// What a phone does with no viewport tag: it lays the page out about 980 px wide, then shrinks it to fit.
// A copy of Example 02 with the tag removed, screenshotted at 980 px and shown at phone size.
function noViewport(file, name, w = 980, h = 1742) {
  const base = "file:///" + REPO.replace(/\\/g, "/") + "/";
  const html = fs.readFileSync(path.join(REPO, file), "utf8").replace(/\s*<meta name="viewport"[^>]*>/, "").replace("<head>", `<head><base href="${base}">`);
  fs.writeFileSync(path.join(SHOTS, name + ".html"), html);
  return shot(path.relative(REPO, path.join(SHOTS, name + ".html")), name, w, h, 1);
}

function videoSlide({ title, label, source, url, watchFor, notes }) {
  const s = panelSlide(title, notes);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: CX, y: 1.95, w: CW, h: 2.3, rectRadius: 0.12, fill: { color: "1E1E1E" }, line: { color: "1E1E1E" },
    hyperlink: { url, tooltip: label } });
  s.addShape(pres.shapes.OVAL, { x: CX + 0.45, y: 2.55, w: 1.1, h: 1.1, fill: { color: C.red }, line: { color: C.red } });
  s.addShape(pres.shapes.ISOSCELES_TRIANGLE, { x: CX + 0.83, y: 2.83, w: 0.45, h: 0.55, rotate: 90, fill: { color: C.white }, line: { color: C.white } });
  txt(s, label, { x: CX + 1.85, y: 2.2, w: CW - 2.1, h: 0.8, fontSize: 22, bold: true, color: C.white });
  txt(s, source, { x: CX + 1.85, y: 3.0, w: CW - 2.1, h: 0.4, fontSize: 14, color: "BFBFBF" });
  txt(s, [{ text: url, options: { hyperlink: { url }, color: "F2C6AC" } }], { x: CX + 1.85, y: 3.45, w: CW - 2.1, h: 0.4, fontSize: 11.5, fontFace: MONO });
  txt(s, "Listen for:", { x: CX, y: 4.55, w: CW, h: 0.4, fontSize: 16, bold: true, color: C.black });
  watchFor.forEach((w, i) => pill(s, w, CX + (i % 3) * 2.5, 5.1 + Math.floor(i / 3) * 0.5, [C.orange, C.teal, C.purple][i % 3], 2.3));
}

// A phone drawn around a screenshot
function phoneFrame(s, img, x, y, w, label, labelColor) {
  const bezel = 0.1, sw = w - 2 * bezel, sh = sw * img.ratio;
  txt(s, label, { x, y: y - 0.32, w, h: 0.28, fontSize: 11, bold: true, color: labelColor, charSpacing: 1, align: "center" });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: sh + 2 * bezel + 0.14, rectRadius: 0.14, fill: { color: "1E1E1E" }, line: { color: "1E1E1E" } });
  s.addImage({ path: img.path, x: x + bezel, y: y + bezel + 0.07, w: sw, h: sh });
}

// A copy of an example page with extra CSS, to show just one part of it in a screenshot
function variant(file, name, css) {
  const base = "file:///" + REPO.replace(/\\/g, "/") + "/";
  const html = fs.readFileSync(path.join(REPO, file), "utf8").replace("<head>", `<head><base href="${base}">`).replace("</head>", `<style>${css}</style></head>`);
  fs.writeFileSync(path.join(SHOTS, name + "_src.html"), html);
  return path.relative(REPO, path.join(SHOTS, name + "_src.html"));
}

// ---------- helpers for this deck ----------

// Take away the indent that every line shares
function undent(lines) {
  const n = Math.min(...lines.filter((l) => l.trim()).map((l) => l.match(/^ */)[0].length));
  return lines.map((l) => l.slice(n));
}
const fileLines = (file) => readLines(file).map(tabs);
// The lines from the one that contains `from` (searching from line `start`), up to and including
// the first one after it that contains `to` (skipping the first `skip` lines after `from`)
function upTo(lines, from, to, { start = 0, skip = 0 } = {}) {
  const a = lines.findIndex((l, i) => i >= start && l.includes(from));
  const b = lines.findIndex((l, i) => i > a + skip && l.includes(to));
  if (a < 0 || b < 0) throw new Error(`upTo(): couldn't find "${from}" / "${to}"`);
  return undent(lines.slice(a, b + 1));
}

// A white slide with a bold title, like the setup slides
function plainSlide(title, notes) {
  const s = newSlide();
  s.background = { color: C.white };
  txt(s, title, { x: 0.75, y: 0.5, w: 11.8, h: 0.7, fontSize: 28, bold: true, color: C.black });
  if (notes) s.addNotes(notes);
  return s;
}

// Recall: a drawing on the left, four numbered points on the right
function recallSlide(title, draw, items, notes) {
  const s = plainSlide(title, notes);
  draw(s);
  const rx = 6.4, rw = W - rx - 0.75, rh = 1.1;
  items.forEach(([t, d], i) => {
    const y = 1.6 + i * (rh + 0.08);
    numCircle(s, i + 1, rx, y + 0.05, C.orange, 0.5);
    txt(s, t, { x: rx + 0.7, y, w: rw - 0.7, h: 0.38, fontSize: 18, bold: true, color: C.black });
    richBullets(s, [d], { x: rx + 0.7, y: y + 0.38, w: rw - 0.7, h: rh - 0.38, size: 15, bullet: false });
  });
}

function strip(s, x, y, w, h, color, label, size = 13) {
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color }, line: { color: C.white, width: 1.5 } });
  if (label) txt(s, label, { x, y, w, h, fontSize: size, fontFace: MONO, bold: true, color: C.white, align: "center", valign: "middle" });
}

// "What you should see": screenshots in one row, all the same height, filling the slide's width
function resultSlide(title, frames, notes, maxH = 5.0, caption) {
  const s = activitySlide(title, notes);
  const gap = 0.35, w = W - 1.5;
  const h = Math.min(maxH, (w - gap * (frames.length - 1)) / frames.reduce((a, f) => a + 1 / f.img.ratio, 0));
  let x = 0.75 + (w - frames.reduce((a, f) => a + h / f.img.ratio, 0) - gap * (frames.length - 1)) / 2;
  frames.forEach(({ img, label }) => { framedImage(s, img, x, 1.85, h / img.ratio, label); x += h / img.ratio + gap; });
  if (caption) {
    const cy = 1.85 + h + 0.4;
    card(s, { x: 0.75, y: cy, w, h: 0.75, fill: C.white });
    richBullets(s, [caption], { x: 1.05, y: cy, w: w - 0.6, h: 0.75, size: 16, bullet: false, valign: "middle" });
  }
  return s;
}

// Bootstrap's own colours, for slides that point at text-bg-* classes
const BS = { primary: "0D6EFD", warning: "FFC107", info: "0DCAF0", dark: "212529" };
const DOCS = "https://getbootstrap.com/docs/5.3/components/";
const W3 = "https://www.w3schools.com/bootstrap5/bootstrap_";

// A page of component tiles, 4 across and 3 down. Each: [name, what it is, docs page, W3Schools page or null, needs JS?]
function componentSlide(title, items, notes) {
  const s = plainSlide(title, notes);
  const cols = 4, gap = 0.2, tw = (W - 1.5 - gap * (cols - 1)) / cols, th = 1.45;
  items.forEach(([name, d, doc, w3, js], i) => {
    const x = 0.75 + (i % cols) * (tw + gap), y = 1.4 + Math.floor(i / cols) * (th + 0.17);
    const color = js ? C.purple : C.teal;
    card(s, { x, y, w: tw, h: th });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.1, h: th, fill: { color }, line: { color } });
    txt(s, [{ text: name, options: { hyperlink: { url: DOCS + doc + "/", tooltip: name + " in Bootstrap's docs" } } }],
      { x: x + 0.25, y: y + 0.1, w: tw - 1.2, h: 0.38, fontSize: 15.5, bold: true, color: C.navy });
    if (js) pill(s, "NEEDS JS", x + tw - 1.0, y + 0.12, C.purple, 0.88);
    txt(s, d, { x: x + 0.25, y: y + 0.48, w: tw - 0.4, h: 0.58, fontSize: 11.5, color: C.dark, valign: "top" });
    const links = [{ text: "Docs", options: { hyperlink: { url: DOCS + doc + "/" } } }];
    if (w3) links.push({ text: "   ·   ", options: { color: C.mid } }, { text: "W3Schools", options: { hyperlink: { url: W3 + w3 + ".php" } } });
    txt(s, links, { x: x + 0.25, y: y + th - 0.36, w: tw - 0.4, h: 0.28, fontSize: 11, bold: true, color: C.teal });
  });
  card(s, { x: 0.75, y: 6.38, w: W - 1.5, h: 0.62, fill: "FBE9DF", border: "F2C6AC" });
  richBullets(s, [[["Your turn: ", "b"], "pick one, open its docs page, and copy the example into your layout. Change the words. ", ["Click a name to open its page.", "i"]]],
    { x: 1.0, y: 6.38, w: W - 2.0, h: 0.62, size: 14, bullet: false, valign: "middle" });
  return s;
}

// A link card: a coloured bar, the name (a link), the address, and a line about it
function linkCard(s, { x, y, w, h, name, url, show, d, color, video }) {
  card(s, { x, y, w, h });
  s.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.1, h, fill: { color }, line: { color } });
  let tx = x + 0.3;
  if (video) {
    s.addShape(pres.shapes.OVAL, { x: x + 0.28, y: y + 0.17, w: 0.42, h: 0.42, fill: { color: C.red }, line: { color: C.red }, hyperlink: { url, tooltip: name } });
    s.addShape(pres.shapes.ISOSCELES_TRIANGLE, { x: x + 0.43, y: y + 0.28, w: 0.17, h: 0.2, rotate: 90, fill: { color: C.white }, line: { color: C.white } });
    tx = x + 0.85;
  }
  txt(s, [{ text: name, options: { hyperlink: { url, tooltip: name } } }], { x: tx, y: y + 0.1, w: x + w - tx - 0.2, h: 0.36, fontSize: 15.5, bold: true, color: C.navy });
  txt(s, [{ text: show, options: { hyperlink: { url } } }], { x: tx, y: y + 0.46, w: x + w - tx - 0.2, h: 0.26, fontSize: 10.5, fontFace: MONO, color: C.teal });
  txt(s, d, { x: tx, y: y + 0.73, w: x + w - tx - 0.2, h: h - 0.8, fontSize: 12, color: C.dark, valign: "top" });
}

// ---------- screenshots ----------
const SH = {
  ex01: shot("01_bootstrap_starter_example.html", "ex01", 900, 300),
  ex02: shot("02_grid_example.html", "ex02", 1000, 430),
  ex03d: shot("03_responsive_grid_example.html", "ex03d", 1100, 270),
  ex03t: shot("03_responsive_grid_example.html", "ex03t", 800, 400),
  ex03p: phone("03_responsive_grid_example.html", "ex03p", 360, 640),
  ex04: shot("04_utilities_example.html", "ex04", 1100, 420),
  ex05: shot("05_components_example.html", "ex05", 1000, 500),
  ex06d: shot("06_navbar_example.html", "ex06d", 1000, 230),
  ex06p: phone("06_navbar_example.html", "ex06p", 360, 330),
  try1: shot("try_it_1_answer.html", "try1", 1000, 330),
  try1p: phone("try_it_1_answer.html", "try1p", 360, 640),
  // Your Layout: the warm-up, then the navbar (Part 2) and cards (Part 3)
  wire: shot("layout/00_wireframe.svg", "wire", 1200, 720, 1),
  lay1: shot("layout/01_layout_answer.html", "lay1", 1000, 350),
  lay1p: phone("layout/01_layout_answer.html", "lay1p", 360, 760),
  lay2: shot("layout/02_navbar_answer.html", "lay2", 1000, 330),
  // The phone menu, open: the same page with the collapse forced open
  lay2p: phone(variant("layout/02_navbar_answer.html", "lay2open", ".navbar-collapse.collapse { display: block !important; }"), "lay2p", 360, 470),
  lay3: shot("layout/03_cards_answer.html", "lay3", 1000, 660),
  lay3p: phone("layout/03_cards_answer.html", "lay3p", 360, 1080),
};


// =====================================================================
// WEEK 10, CLASS 1: CSS Frameworks and the Bootstrap Grid
section("Class 1 - CSS Frameworks and the Bootstrap Grid");

titleSlide("Week 10, Class 1", "CSS frameworks, and Bootstrap's 12-column grid.",
  "Welcome back! For the last few weeks we've written every line of CSS ourselves: flexbox, grid, float, media queries. This week we look at CSS frameworks: CSS that someone else has already written, that we link to and use with class names.\n\nBootstrap is the main one, this week and next. Next week we also look at Tailwind CSS, and rebuild Harbourside with Bootstrap.\n\nSame routine as before: pop quizzes all the way through, and Try Its.");

setupSlide("Attendance and a Quick Check-in", [
  ["OUTLOOK", "Are you checking your school Outlook email?"], ["TEAMS", "Are you signed into Teams?"],
  ["TEAMS", "Are you following our Teams channel and General chat?"], ["BRIGHTSPACE", "Are you following the Content on Brightspace?"],
], "Take attendance while everyone checks these four things.\n\nRemind students that announcements go to their school email and the Teams channel. If anyone hasn't been checking, now is the time.");

setupSlide("Is Your Environment Set Up?", [
  ["GIT", "Did you run git pull? You should see a 10_CSS_Frameworks folder."],
  ["VS CODE", "Is the 10_CSS_Frameworks folder open in VS Code?"],
  ["LIVE SERVER", "Does Example 01 (the starter) open with Go Live?"],
  ["INTERNET", "Bootstrap loads from the web today. Does Example 01 have a blue button?"],
], "Help anyone who's stuck now, so they aren't lost during the Try Its.\n\nNew this week: the examples link to Bootstrap on a CDN (a server on the internet). If Example 01 shows plain black text and a grey button, the Bootstrap file didn't load. Check the Wi-Fi, then check the link for typos.\n\ngit pull: if it complains about local changes, they can commit their changes first, or copy their work somewhere safe and run git stash.");

recallSlide("Recall: What Did We Do Last Class?", (s) => {
  // A phone, a tablet and a laptop, with 1, 2 and 3 columns
  const devs = [[0.75, 2.3, 1.1, 2.6, 1], [2.05, 1.95, 1.75, 2.95, 2], [4.0, 1.6, 2.15, 1.8, 3]];
  devs.forEach(([x, y, w, h, n]) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.1, fill: { color: "1E1E1E" }, line: { color: "1E1E1E" } });
    const ix = x + 0.1, iy = y + 0.12, iw = w - 0.2, ih = h - 0.24;
    s.addShape(pres.shapes.RECTANGLE, { x: ix, y: iy, w: iw, h: ih, fill: { color: C.white }, line: { color: C.white } });
    s.addShape(pres.shapes.RECTANGLE, { x: ix, y: iy, w: iw, h: 0.25, fill: { color: C.navy }, line: { color: C.navy } });
    const cw = (iw - 0.1 - (n - 1) * 0.06) / n;
    for (let r = 0; r < Math.ceil(3 / n); r++) for (let c = 0; c < n && r * n + c < 3; c++) {
      const ch = n === 1 ? 0.55 : n === 2 ? 0.95 : ih - 0.45;
      s.addShape(pres.shapes.RECTANGLE, { x: ix + 0.05 + c * (cw + 0.06), y: iy + 0.33 + r * (ch + 0.08), w: cw, h: ch, fill: { color: [C.orange, C.teal, C.purple][r * n + c] }, line: { type: "none" } });
    }
  });
  txt(s, "One page, three screens: 1, 2 and 3 columns.", { x: 0.75, y: 5.85, w: 5.2, h: 0.4, fontSize: 13, italic: true, color: C.grey, align: "center" });
}, [
  ["The viewport meta tag", ["Goes in every page, or phones lay it out 980px wide."]],
  ["Media queries", [["@media (min-width: 768px)", "c"], ": phones first, then add to it."]],
  ["Fluid layouts", ["Percentages, ", ["max-width", "c"], ", and ", ["box-sizing: border-box", "c"], "."]],
  ["Images and text", [["img { max-width: 100%; }", "c"], " and ", ["clamp()", "c"], "."]],
], "Before you read anything out, ask: what did we do last class? Give everyone 30 seconds to tell a neighbour, then take a few answers.\n\nThen go through the four points.\n\nThe link to today: Bootstrap is built on exactly these things. Its grid is flexbox plus min-width media queries, and it needs the viewport tag. Students who understood last week will find Bootstrap much less magical.");

overviewSlide("Week 10, Class 1: Overview", "CSS Frameworks and the Bootstrap Grid", [
  "What a CSS framework is, and what it costs",
  "The popular frameworks, and the two kinds",
  "Adding Bootstrap to a page",
  "Containers, rows and the 12-column grid",
  "Breakpoints, and colour and spacing classes",
], [["Why", "Frameworks"], ["Link", "Bootstrap"], ["Grid", "12 columns"]],
  "Rough plan for the class:\n- 10 min: check-in, environment, recall, overview, warm-up\n- 15 min: frameworks, the two kinds, Pop Quiz 1\n- 15 min: Bootstrap, Example 01, Pop Quiz 2\n- 25 min: the grid, breakpoints, Examples 02 and 03, Pop Quizzes 3 to 5\n- 10 min: utilities, Example 04, Pop Quiz 6\n- 15 min: Try It 1, together\n- 5 min: recap\n\nPop quizzes: give everyone 30 seconds to think, then count down and have them hold up 1 to 4 fingers (A to D) all at once. Or use the Teams chat. Nobody is marked on these, but the Week 10 quiz on Brightspace uses the same questions.");

{
  const s = activitySlide("Warm-up: What Does \"Bootstrap\" Mean? (1 min)",
    "Ask before saying anything about the framework. One minute with a neighbour, then take a few answers.\n\nThe phrase: to 'pull yourself up by your bootstraps' means to get going with no outside help. (Bootstraps are the loops at the back of a boot. Taken literally it's impossible, which was the original joke.)\n\nComputers: 'booting' a computer is short for bootstrapping: a tiny program that loads a bigger one, which loads the operating system.\n\nBusiness: a 'bootstrapped' startup grows on its own money, with no investors.\n\nThe framework: Bootstrap gets a site up and running quickly, from almost nothing. It started at Twitter in 2011 as a way to make their internal tools look consistent.");
  [["\"To bootstrap\"", "What does it mean when a person or a business does it?"], ["\"Booting\" a computer", "Where do you think that word comes from?"]].forEach(([t, d], i) => {
    const x = 0.75 + i * 6.1;
    card(s, { x, y: 1.75, w: 5.75, h: 4.2 });
    txt(s, "?", { x, y: 2.0, w: 5.75, h: 1.4, fontSize: 80, bold: true, color: i ? C.purple : C.orange, align: "center", valign: "middle" });
    txt(s, t, { x: x + 0.35, y: 3.6, w: 5.05, h: 0.9, fontSize: 24, bold: true, color: C.black, align: "center", valign: "middle" });
    txt(s, d, { x: x + 0.35, y: 4.6, w: 5.05, h: 0.8, fontSize: 16, color: C.grey, align: "center", valign: "top" });
  });
}

dividerSlide("CSS Frameworks", "What they are   ·   what they cost   ·   the popular ones",
  "So far we've written all of our CSS by hand. A framework is CSS someone else has already written and tested, which you use by putting class names on your HTML.");

{
  const s = panelSlide("What Is a CSS Framework?",
    "Read the definition. The important words: already written, and class names.\n\nThe four benefits from last year's slides:\n- Consistency: every button, form and table on every page looks the same, without you having to remember the exact padding.\n- Speed: a navbar that folds into a menu button on phones takes a few minutes, not an afternoon.\n- Responsive: the grid and components are built phones first, with media queries like last week's.\n- Cross-browser: the framework's authors have already tested it in every current browser.\n\nThe dark box is the point to come back to all week: a framework is not a different language. When something looks wrong, it's still the cascade, specificity, the box model and flexbox underneath. DevTools still works.");
  txt(s, "A library of CSS (and sometimes JavaScript) that's already written. You use it by adding class names to your HTML.", { x: CX, y: 1.95, w: CW, h: 1.05, fontSize: 19, color: C.black, valign: "top" });
  [["Consistency", "Every page looks the same"], ["Speed", "Ready-made parts"], ["Responsive", "Phones first, built in"], ["Cross-browser", "Already tested"]].forEach(([t, d], i) => {
    const w = 1.74, x = CX + i * (w + 0.197);
    card(s, { x, y: 3.15, w, h: 2.1 });
    numCircle(s, i + 1, x + 0.18, 3.3, [C.orange, C.teal, C.purple, C.navy][i], 0.48);
    txt(s, t, { x: x + 0.18, y: 3.9, w: w - 0.3, h: 0.4, fontSize: 15, bold: true, color: C.black });
    txt(s, d, { x: x + 0.18, y: 4.3, w: w - 0.3, h: 0.8, fontSize: 13.5, color: C.grey, valign: "top" });
  });
  card(s, { x: CX, y: 5.5, w: CW, h: 1.15, fill: C.dark, border: C.dark });
  richBullets(s, [[["It's still just CSS. ", "b"], "Selectors, the cascade, flexbox and media queries: everything you've learned is underneath."]], { x: CX + 0.3, y: 5.5, w: CW - 0.6, h: 1.15, size: 16, bullet: false, color: C.white, valign: "middle" });
}

{
  const s = panelSlide("What a Framework Costs",
    "Last year's slides only listed the benefits. Students should hear the other side too, so they can choose.\n\n- Every site looks the same: the default Bootstrap look (that blue button, that navbar) is very recognizable. You need your own CSS on top to look like yourself. That's next week.\n- Size: bootstrap.min.css is about 230 KB, and the JavaScript bundle about 80 KB. Compressed over the network they're about 31 KB and 24 KB. Most pages use a small part of it. Show this in the DevTools Network tab on Example 01 (F12, Network, reload).\n- Lots of classes: the HTML gets busy, and there are hundreds of class names to learn. The docs are open on every developer's second monitor.\n- You still need CSS: to change the look, and to fix things when they go wrong. Which is why we learned CSS first.\n\nOn balance, frameworks are very widely used, especially for prototypes, admin pages and internal tools, where speed matters more than a unique look.");
  [["Sites look alike", "The \"Bootstrap look\" is easy to spot. You'll want your own CSS on top."],
    ["A big file", "About 230 KB of CSS (31 KB compressed), and most pages use a small part."],
    ["Lots of classes", "Busy HTML, and hundreds of class names to look up."],
    ["You still need CSS", "To make it yours, and to fix it when it goes wrong."]].forEach(([t, d], i) => {
    const w = 3.65, h = 2.2, x = CX + (i % 2) * (w + 0.25), y = 1.95 + Math.floor(i / 2) * (h + 0.25);
    card(s, { x, y, w, h });
    txt(s, t, { x: x + 0.25, y: y + 0.25, w: w - 0.5, h: 0.45, fontSize: 19, bold: true, color: C.red });
    txt(s, d, { x: x + 0.25, y: y + 0.8, w: w - 0.5, h: 1.2, fontSize: 16, color: C.dark, valign: "top" });
  });
}

{
  const s = panelSlide("Two Kinds of Framework",
    "This is the main idea for comparing frameworks, and it's why we look at Tailwind next week.\n\nComponent frameworks (Bootstrap, Foundation, Bulma) give you finished parts with names: a button, a card, a navbar. You write class=\"btn btn-primary\" and get a complete, styled button.\n\nUtility-first frameworks (Tailwind) give you small classes that each do about one CSS property: bg-blue-600 is a background colour, px-4 is left and right padding, rounded is a border radius. You build the button yourself, out of those.\n\nBoth buttons here look about the same in a browser. Ask: which HTML would be easier to read? Which would be easier to make look different from every other site?\n\nBootstrap has utility classes too (we'll see them today), so the line isn't perfectly sharp. It's about what each one leads with.");
  [["COMPONENT · BOOTSTRAP", "Ready-made parts with names.", ['<button class="btn btn-primary">', "   Order", "</button>"], C.teal],
    ["UTILITY-FIRST · TAILWIND", "Small classes you put together.", ['<button class="bg-blue-600 text-white', '               px-4 py-2 rounded">', "   Order", "</button>"], C.purple]].forEach(([k, d, code, c], i) => {
    const y = 1.95 + i * 2.4;
    txt(s, k, { x: CX, y, w: CW, h: 0.32, fontSize: 13, bold: true, color: c, charSpacing: 1 });
    txt(s, d, { x: CX, y: y + 0.32, w: CW, h: 0.4, fontSize: 17, color: C.black });
    codeBox(s, code, { x: CX, y: y + 0.8, w: CW, h: 1.35, maxSize: 15, lang: "html", fitHeight: true });
  });
}

{
  const s = panelSlide("Popular CSS Frameworks",
    "From last year's slides, updated.\n\n- Bootstrap: on more existing websites than any other CSS framework. Version 5.3 is current, and version 6 is in alpha (a test version, not for real sites yet).\n- Tailwind CSS: the most popular choice for new projects in developer surveys over the last few years. Version 4 is current.\n- Foundation, by ZURB: very customizable, and was big for a while, but it gets few updates now.\n- Bulma: CSS only (no JavaScript), built on flexbox, with readable class names like 'button is-primary'.\n\nWhy we demo Bootstrap and Tailwind: they're the two students are most likely to meet in a job, and they're one of each kind. If you learn those two, the others will look familiar.");
  const rows = [["Bootstrap", "Component", "The most used on existing sites. Grid, components, JavaScript."],
    ["Tailwind CSS", "Utility-first", "The most popular for new projects. Small classes for everything."],
    ["Foundation", "Component", "Very customizable. Few updates now."],
    ["Bulma", "Component", "Flexbox based, CSS only, readable class names."]];
  const hdr = ["Framework", "Kind", "Known for"].map((t) => ({ text: t, options: { bold: true, color: C.white, fill: { color: C.navy } } }));
  s.addTable([hdr, ...rows.map((r, i) => r.map((t, j) => ({ text: t, options: { bold: j === 0, fill: { color: i % 2 ? "F7F9FA" : C.white } } })))],
    { x: CX, y: 1.95, w: CW, colW: [1.85, 1.6, 4.1], fontFace: FONT, fontSize: 15, color: C.black, border: { type: "solid", pt: 0.75, color: "D0D7DB" }, rowH: 0.75, valign: "middle", margin: 0.1 });
  card(s, { x: CX, y: 5.95, w: CW, h: 0.75, fill: "FBE9DF", border: "F2C6AC" });
  richBullets(s, [[["Bootstrap this week, ", "b"], "and a look at Tailwind next week: one of each kind."]], { x: CX + 0.25, y: 5.95, w: CW - 0.5, h: 0.75, size: 15, bullet: false, valign: "middle" });
}

quiz({
  q: "Which of these frameworks is utility-first?",
  options: ["Bootstrap", "Bulma", "Foundation", "Tailwind CSS"], answer: 3,
  why: ["Utility-first means lots of small classes, each doing about one thing: `px-4`, `bg-blue-600`, `rounded`.", "Bootstrap, Bulma and Foundation lead with ready-made components, like `btn` and `card`."],
  notes: "A quick check on the last two slides.",
  revealNotes: "Answer: D.\n\nIf anyone said Bootstrap: it does have utility classes (we'll use them today), but it's known for its components.",
});

dividerSlide("Bootstrap", "What it is   ·   adding it to a page   ·   the grid",
  "Bootstrap is the framework for the rest of the week.");

{
  const s = panelSlide("What Is Bootstrap?",
    "Bootstrap is free and open source (MIT licence), so you can use it on any site, including for a client.\n\nHistory: Mark Otto and Jacob Thornton made it at Twitter in 2011, to make Twitter's internal tools consistent. It was released to everyone later that year and became the most used CSS framework on the web.\n\nVersions matter: a lot of tutorials online are for Bootstrap 3 or 4, and their code doesn't always work in 5. When you search for help, put 'Bootstrap 5' in the search. Last year's slides used Bootstrap 4 code in places, and we'll see what goes wrong with that next class.\n\nThe four parts, from last year's slides: today is the grid and the utilities. Next class is components and JavaScript.");
  txt(s, "A free, open-source framework of CSS and JavaScript, built phones first. Made at Twitter in 2011. We use version 5.3.", { x: CX, y: 1.95, w: CW, h: 1.05, fontSize: 19, color: C.black, valign: "top" });
  [["Grid", "Rows and 12 columns", "Today"], ["Utilities", "Colour, spacing, display", "Today"], ["Components", "Buttons, cards, navbars", "Next class"], ["JavaScript", "Menus, dropdowns, modals", "Next class"]].forEach(([t, d, when], i) => {
    const w = 3.65, h = 1.25, x = CX + (i % 2) * (w + 0.25), y = 3.15 + Math.floor(i / 2) * (h + 0.15);
    card(s, { x, y, w, h, fill: i < 2 ? C.white : "F7F9FA" });
    txt(s, t, { x: x + 0.25, y: y + 0.18, w: w - 1.6, h: 0.4, fontSize: 18, bold: true, color: C.teal });
    pill(s, when.toUpperCase(), x + w - 1.45, y + 0.2, i < 2 ? C.orange : C.mid, 1.25);
    txt(s, d, { x: x + 0.25, y: y + 0.7, w: w - 0.5, h: 0.5, fontSize: 15, color: C.dark, valign: "top" });
  });
  richBullets(s, [[["Searching for help? ", "b"], "Add \"Bootstrap 5\". Code for 3 and 4 doesn't always work."]], { x: CX, y: 6.0, w: CW, h: 0.4, size: 14, bullet: false, color: C.grey });
}

codeSlide({
  kicker: "LINKING BOOTSTRAP · EXAMPLE 01", title: "Adding Bootstrap to a Page", file: "10_CSS_Frameworks/01_bootstrap_starter_example.html",
  lines: [...upTo(fileLines("01_bootstrap_starter_example.html"), "<!-- 1.", "</head>"), "<body>", "   ...", ...upTo(fileLines("01_bootstrap_starter_example.html"), "<script", "</body>")],
  lang: "html", result: SH.ex01, size: 14, mark: [1, 3, 7],
  bullets: [
    [["1. ", "b"], "Bootstrap's CSS in the ", ["head", "c"], "."],
    [["2. ", "b"], "Your own CSS ", ["after", "b"], " it."],
    [["3. ", "b"], "Bootstrap's JavaScript just before ", ["</body>", "c"], "."],
  ],
  notes: "Open 01_bootstrap_starter_example.html with Live Server.\n\nThree lines do everything (the orange marks):\n1. The CSS link. Without it, nothing on the page is styled by Bootstrap. Try it: comment it out and save. The button goes grey and the font changes back.\n2. Our own stylesheet, after Bootstrap's. Our h1 rule turns the heading navy. That's the next quiz.\n3. The JavaScript bundle. Only the interactive parts need it (navbars on phones, dropdowns, modals, carousels). 'Bundle' means it includes Popper, which positions dropdowns and tooltips.\n\nThe viewport meta tag from last week is here too. Bootstrap is built phones first and won't work properly on a phone without it.\n\nThe classes: container centres the page, py-4 adds padding on the top and bottom, lead makes a bigger paragraph, and btn btn-primary is a button. More on all of these soon.\n\nBootstrap's docs also show integrity=\"sha384-...\" and crossorigin attributes on these tags. They're a safety check that the CDN's file hasn't been changed. Fine to copy them from the docs exactly; leave them out if you type the link by hand.",
});

{
  const s = panelSlide("A CDN, or Download It?",
    "Option 1, the CDN, is what we use in class. A CDN (content delivery network) is a group of servers around the world that hand out popular files. jsDelivr is the one Bootstrap's own docs use.\n\nThe version is in the URL: bootstrap@5.3.8. Keep it, so the site doesn't change when Bootstrap releases a new version.\n\nLast year's slides used a stackpath.bootstrapcdn.com link for Bootstrap 4.5.2. Don't use it: it's an old version, from an old CDN.\n\nOption 2, downloading: put bootstrap.min.css and bootstrap.bundle.min.js in a folder in your site and link to them with a relative path, like any other stylesheet. It works offline, and nothing outside your site can break it.\n\nOption 3, npm, is what professional projects do (npm install bootstrap). It also lets you customize Bootstrap with Sass before it's built. That's beyond this course.");
  [["1. A CDN", "One link to a copy on the web. Fast, and nothing to download. Needs the internet.", ['<link rel="stylesheet"', '  href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/css/bootstrap.min.css">'], C.orange],
    ["2. Download it", "Put the files in your site's folder, and link to them like your own CSS. Works offline.", ['<link rel="stylesheet" href="css/bootstrap.min.css">'], C.teal]].forEach(([t, d, code, c], i) => {
    const y = 1.95 + i * 2.15;
    txt(s, t, { x: CX, y, w: CW, h: 0.4, fontSize: 19, bold: true, color: c });
    txt(s, d, { x: CX, y: y + 0.42, w: CW, h: 0.45, fontSize: 15, color: C.dark });
    codeBox(s, code, { x: CX, y: y + 0.95, w: CW, h: code.length > 1 ? 1.0 : 0.7, maxSize: 13, lang: "html" });
  });
  richBullets(s, [[["3. npm", "b"], " (", ["npm install bootstrap", "c"], "): what big projects use. Not in this course."]], { x: CX, y: 6.25, w: CW, h: 0.45, size: 15, bullet: false, color: C.grey });
}

quiz({
  q: "You want your own CSS to change some of Bootstrap's styles. Where does the `<link>` to your stylesheet go?",
  options: ["Before Bootstrap's `<link>`", "At the very end of the `<body>`", "After Bootstrap's `<link>`", "Anywhere: the order doesn't matter"], answer: 2,
  why: ["Bootstrap's rules and yours often have the same specificity, like `h1` and `h1`.", "When rules tie, the one that comes later wins. So link yours second."],
  notes: "The cascade, from Week 5, coming back.",
  revealNotes: "Answer: C.\n\nDemo it in Example 01: swap the two link lines and save. The heading goes from navy back to Bootstrap's near-black.\n\nIt's the same reason media queries go last in a stylesheet.",
});

{
  const s = panelSlide("Container, Row, Columns",
    "Last year's slides put it like this: in the minds of Bootstrap users, a page has one container, the container has rows, and each row has 12 columns. It's a lot like thinking in tables, but with divs and classes.\n\n- container: centres the page and adds padding at the sides. It has a max-width that steps up at each breakpoint (up to 1320px). container-fluid is always the full width of the window.\n- row: a flex container (display: flex; flex-wrap: wrap). Everything in it is a column.\n- col-4: this column is 4 of the 12 columns wide, so a third.\n\nWhy 12? It divides evenly into halves (6), thirds (4), quarters (3) and sixths (2).\n\nThe numbers in a row should add up to 12. If they add up to more, the extra wraps to a new line: that's a pop quiz soon.");
  const x0 = CX, y0 = 1.95, w0 = CW;
  s.addShape(pres.shapes.RECTANGLE, { x: x0, y: y0, w: w0, h: 4.1, fill: { color: C.white }, line: { color: C.red, width: 2 } });
  txt(s, '.container', { x: x0 + 0.15, y: y0 + 0.08, w: 3, h: 0.3, fontSize: 13, fontFace: MONO, bold: true, color: C.red });
  const rx = x0 + 0.3, rw = w0 - 0.6, cw = rw / 12;
  const rows = [[[1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1], "12 columns"], [[6, 6], "6 + 6"], [[4, 4, 4], "4 + 4 + 4"], [[3, 9], "3 + 9"]];
  rows.forEach(([spans, lab], ri) => {
    const ry = y0 + 0.5 + ri * 0.88;
    s.addShape(pres.shapes.RECTANGLE, { x: rx, y: ry, w: rw, h: 0.72, fill: { color: "E4EFF3" }, line: { color: C.blue, width: 1.5 } });
    txt(s, ".row", { x: rx + 0.05, y: ry - 0.01, w: 1, h: 0.22, fontSize: 9.5, fontFace: MONO, bold: true, color: C.blue });
    let cx = rx;
    spans.forEach((n, i) => {
      strip(s, cx + 0.03, ry + 0.22, n * cw - 0.06, 0.44, [C.green, C.teal, C.navy][i % 3], n > 1 ? `col-${n}` : "1", n > 1 ? 13 : 10);
      cx += n * cw;
    });
  });
  richBullets(s, [[["Why 12? ", "b"], "It splits evenly into halves, thirds, quarters and sixths."]], { x: CX, y: 6.2, w: CW, h: 0.45, size: 15, bullet: false });
}

codeSlide({
  kicker: "THE GRID · EXAMPLE 02", title: "Columns Add Up to 12", file: "10_CSS_Frameworks/02_grid_example.html",
  lines: upTo(bodyOf("02_grid_example.html"), "<!-- Halves", "(wraps)").concat(["</div>"]),
  lang: "html", result: SH.ex02, size: 14,
  bullets: [
    ["A ", ["row", "c"], " holds columns. ", ["col-4", "c"], " is ", ["4 of the 12", "b"], " columns wide."],
    ["Make each row add up to 12."],
  ],
  notes: "Open 02_grid_example.html and walk down the rows: 12 ones, halves, thirds, then a 3 + 9 sidebar layout.\n\nThe last row on the page adds up to 14. Don't explain it yet: it's the next quiz.\n\nTry it live: change col-3 and col-9 to col-2 and col-10, save, and look.\n\nThe CSS file only colours the columns so you can see them. Every width comes from Bootstrap.\n\nThe white gaps between columns are Bootstrap's 'gutters'. The g-3 class (Example 03) changes their size.",
});

quiz({
  q: "What does this row look like?",
  code: ['<div class="row">', '   <div class="col-6">A</div>', '   <div class="col-4">B</div>', '   <div class="col-4">C</div>', "</div>"], lang: "html",
  options: ["All three squeeze onto one line", "C wraps onto a new line", "Bootstrap ignores C", "The browser shows an error"], answer: 1,
  why: ["6 + 4 + 4 = 14, and a row only has 12.", "A row is a flexbox with `flex-wrap: wrap`, so C moves down to the next line, still 4 columns wide."],
  revealImg: SH.ex02,
  notes: "Let them add it up first.",
  revealNotes: "Answer: B.\n\nThe picture is Example 02. The bottom row is this exact code.\n\nSometimes wrapping is what you want: a row of 8 cards that are col-3 each makes two lines of 4. That's how Bootstrap card grids usually work.",
});

{
  const s = panelSlide("Bootstrap's Breakpoints",
    "These are Bootstrap's breakpoints. They're close to the ones we used last week, with a few more.\n\nEvery one of them is a min-width media query, so Bootstrap is phones first, like Example 04 last week. col-md-4 means: from 768px and up, be 4 columns wide. Below 768px, the class does nothing, and a column with no other class is the full width of the row.\n\nRead the class in three parts: col, the breakpoint, the number. No breakpoint (col-4) means every size, from the smallest phone up.\n\nLast year's slide listed col-sm-, col-md- and col-lg-. There are also col-xl- and col-xxl-, and the classes add up: col-12 col-md-6 col-lg-3 on one div. That's the next example.");
  const bps = [["xs", "under 576", "(none)"], ["sm", "576px +", "col-sm-"], ["md", "768px +", "col-md-"], ["lg", "992px +", "col-lg-"], ["xl", "1200px +", "col-xl-"], ["xxl", "1400px +", "col-xxl-"]];
  const sw = CW / 6, cols = [C.teal, "1F7A8C", C.purple, C.navy, C.dark, "000000"];
  bps.forEach(([t, d, cl], i) => {
    s.addShape(pres.shapes.RECTANGLE, { x: CX + i * sw, y: 1.95, w: sw, h: 1.35, fill: { color: cols[i] }, line: { color: C.light, width: 1.5 } });
    txt(s, t, { x: CX + i * sw, y: 2.02, w: sw, h: 0.42, fontSize: 18, bold: true, color: C.white, align: "center" });
    txt(s, d, { x: CX + i * sw, y: 2.45, w: sw, h: 0.35, fontSize: 12.5, color: C.white, align: "center" });
    txt(s, cl, { x: CX + i * sw, y: 2.82, w: sw, h: 0.35, fontSize: 11.5, fontFace: MONO, color: "F2C6AC", align: "center" });
  });
  txt(s, "Reading a column class:", { x: CX, y: 3.6, w: CW, h: 0.4, fontSize: 17, bold: true, color: C.black });
  [["col", "A grid column.", C.green], ["md", ["From the ", ["md", "b"], " breakpoint (768px) ", ["and up", "b"], "."], C.purple], ["4", "4 of the 12 columns wide.", C.orange]].forEach(([p, d, c], i) => {
    const y = 4.1 + i * 0.62;
    pill(s, p, CX, y + 0.05, c, 1.2);
    richBullets(s, [typeof d === "string" ? [d] : d], { x: CX + 1.4, y, w: CW - 1.4, h: 0.45, size: 15, bullet: false, valign: "middle" });
  });
  card(s, { x: CX, y: 6.0, w: CW, h: 0.7, fill: C.dark, border: C.dark });
  richBullets(s, [["Each one is a ", ["min-width", "c"], " media query: ", ["phones first", "b"], ", like last week."]], { x: CX + 0.25, y: 6.0, w: CW - 0.5, h: 0.7, size: 15, bullet: false, color: C.white, valign: "middle", codeColor: "F2C6AC" });
}

codeSlide({
  kicker: "THE GRID · EXAMPLE 03", title: "A Responsive Grid", file: "10_CSS_Frameworks/03_responsive_grid_example.html",
  lines: upTo(bodyOf("03_responsive_grid_example.html"), "<!-- 1 across", '<div class="box">4').concat(["   </div>", "</div>"]),
  lang: "html", frames: [[{ img: SH.ex03d, label: "1100PX: 4 ACROSS" }], [{ img: SH.ex03p, label: "360PX" }, { img: SH.ex03t, label: "800PX: 2 ACROSS" }]], maxFrameH: 4.0, size: 14,
  bullets: [
    ["Classes add up: ", ["1", "b"], " across on a phone, ", ["2", "b"], " from md, ", ["4", "b"], " from lg."],
  ],
  notes: "Open 03_responsive_grid_example.html and drag the window from wide to narrow. The badge at the top says which breakpoint you're at.\n\nRead the class from left to right, smallest screen first:\n- col-12: on the smallest screens, 12 of 12 columns, so full width. (You could leave this one out: a column with no class for that size is full width anyway. It's there to make it obvious.)\n- col-md-6: from 768px, half width, so 2 across.\n- col-lg-3: from 992px, a quarter, so 4 across.\n\ng-3 on the row sets the gap ('gutter') between the boxes, across and down.\n\nThe badges use display classes: d-none hides one, d-md-inline shows it again from md up. More on those next class.",
});

quiz({
  q: "On a 600px-wide phone, how do these three columns look?",
  code: ['<div class="row">', '   <div class="col-md-4">A</div>', '   <div class="col-md-4">B</div>', '   <div class="col-md-4">C</div>', "</div>"], lang: "html",
  options: ["Stacked, each one the full width", "Side by side, a third each", "Two on the first line, one on the next", "Hidden until the window is 768px"], answer: 0,
  why: ["`md` means 768px and up, and 600px is smaller, so `col-md-4` does nothing yet.", "Below that, each column is the full width of the row. That's what you want on a phone."],
  notes: "This is how the activity from last year's slides is written: 'three columns that take up equal space on medium and larger screens'.",
  revealNotes: "Answer: A.\n\nIf they said B: that would be col-4, with no breakpoint, which is a third at every size, even 360px. Usually too narrow on a phone.\n\nDemo: Example 03 with the device toolbar.",
});

{
  const s = panelSlide("Look Inside a Bootstrap Class",
    "This links Bootstrap back to Weeks 8 and 9, and gets students using DevTools on it.\n\nDemo in Example 03: right-click a box, Inspect, then click its parent div (the col-12 col-md-6 col-lg-3 one). In the Styles panel you'll see the .col-md-6 and .col-lg-3 rules, each inside an @media (min-width: ...) block. Drag the window narrower and watch them get crossed out or disappear.\n\nThen click the row and find display: flex and flex-wrap: wrap. Chrome and Edge show a little 'flex' badge on the element in the Elements panel; click it to see the flex overlay.\n\nThis is the actual CSS from bootstrap.css 5.3.8, trimmed. 33.33333333% is a third.\n\nThe point: there's nothing magic in a framework. It's flexbox and media queries, which they already know.");
  codeBox(s, [".row {", "   display: flex;", "   flex-wrap: wrap;", "}", "", "@media (min-width: 768px) {", "   .col-md-4 {", "      flex: 0 0 auto;", "      width: 33.33333333%;", "   }", "}"],
    { x: CX, y: 1.95, w: 4.1, h: 3.9, maxSize: 15, lang: "css", title: "FROM BOOTSTRAP.CSS" });
  const tx = CX + 4.4, tw = CW - 4.4;
  txt(s, "Try it in DevTools:", { x: tx, y: 1.95, w: tw, h: 0.4, fontSize: 16, bold: true, color: C.black });
  [["Right-click a column", "Inspect"], ["Styles panel", "Find the @media rule"], ["Drag the window", "Watch it switch off"]].forEach(([t, d], i) => {
    const y = 2.5 + i * 1.05;
    numCircle(s, i + 1, tx, y, C.orange, 0.42);
    txt(s, t, { x: tx + 0.55, y: y - 0.03, w: tw - 0.55, h: 0.35, fontSize: 15, bold: true, color: C.black });
    txt(s, d, { x: tx + 0.55, y: y + 0.32, w: tw - 0.55, h: 0.35, fontSize: 13.5, color: C.grey });
  });
  card(s, { x: CX, y: 6.05, w: CW, h: 0.65, fill: C.dark, border: C.dark });
  richBullets(s, [[["No magic: ", "b"], "flexbox (Week 8) and a min-width media query (Week 9)."]], { x: CX + 0.25, y: 6.05, w: CW - 0.5, h: 0.65, size: 15, bullet: false, color: C.white, valign: "middle" });
}

quiz({
  q: "What does Bootstrap's grid use to put columns side by side?",
  options: ["Float", "HTML tables", "Flexbox", "JavaScript"], answer: 2,
  why: ["`.row` is `display: flex` with `flex-wrap: wrap`. Each column sets its own width.", "Bootstrap 3 used float. Since Bootstrap 4, it's flexbox. No JavaScript needed: the grid is pure CSS."],
  notes: "Straight after the DevTools slide, so most people should get it.",
  revealNotes: "Answer: C.\n\nA bit of history: Bootstrap 3 (2013) used floats and clearfixes, like our float examples from Week 9. Bootstrap 4 (2018) switched to flexbox. Bootstrap 5 also has an optional CSS Grid version, but flexbox is still the default.",
});

codeSlide({
  kicker: "UTILITIES · EXAMPLE 04", title: "Colour and Spacing Classes", file: "10_CSS_Frameworks/04_utilities_example.html",
  lines: upTo(bodyOf("04_utilities_example.html"), "<h2 class=\"h5\">Backgrounds", ".px-5 .py-1"),
  lang: "html", result: SH.ex04, size: 14,
  bullets: [
    [["text-bg-primary", "c"], ": a background ", ["and", "b"], " a text colour you can read on it."],
    [["p-3", "c"], " is padding, ", ["mb-2", "c"], " is margin on the bottom."],
  ],
  notes: "Open 04_utilities_example.html. There's no CSS file of our own: every colour and space is a Bootstrap class.\n\nColours (from last year's colour slides, updated for Bootstrap 5.3):\n- text-primary, text-success, text-danger: text colour.\n- bg-primary and friends: background colour only.\n- text-bg-primary: background plus a matching readable text colour. Use this one for coloured boxes. On yellow (warning) it picks black text, which bg-warning with text-white would get wrong.\n- text-muted and text-black-50 from last year's slides still work, but are deprecated in 5.3. Use text-body-secondary.\n\nThe names (primary, success, danger, warning, info, light, dark) describe a purpose, not a colour. 'danger' is red, but if you change the theme, it changes everywhere.\n\nSpacing is on the next slide.",
});

{
  const s = panelSlide("Reading a Spacing Class",
    "Every spacing class is built from three parts. Once students can read one, they can read them all.\n\n- m for margin, p for padding.\n- The side: t (top), b (bottom), s (start, the left in English), e (end, the right), x (left and right), y (top and bottom). Leave it out for all four sides.\n- The size: 0 to 5. With Bootstrap's defaults: 0 = 0, 1 = 0.25rem (4px), 2 = 0.5rem (8px), 3 = 1rem (16px), 4 = 1.5rem (24px), 5 = 3rem (48px). Plus auto for margins.\n\nmx-auto is the old margin: 0 auto centring trick.\n\nBootstrap 4 used l and r (ml-3, pr-2). Bootstrap 5 changed them to s and e (ms-3, pe-2), so that pages in right-to-left languages flip properly. Old tutorials will have ml- and mr-, which do nothing in Bootstrap 5.\n\nThey can also take a breakpoint: p-2 p-md-5 is small padding on phones, big from md up.");
  codeBox(s, ['<div class="mt-3 px-5 mb-0">'], { x: CX, y: 1.95, w: CW, h: 0.7, maxSize: 17, lang: "html" });
  [["m  or  p", ["Margin or padding."], C.purple], ["t b s e x y", [["t", "b"], "op, ", ["b", "b"], "ottom, ", ["s", "b"], "tart (left), ", ["e", "b"], "nd (right), ", ["x", "b"], " = left and right, ", ["y", "b"], " = top and bottom. None = all four."], C.teal], ["0 to 5", [["0", "c"], " none, ", ["1", "c"], " 4px, ", ["2", "c"], " 8px, ", ["3", "c"], " 16px, ", ["4", "c"], " 24px, ", ["5", "c"], " 48px, or ", ["auto", "c"], "."], C.orange]].forEach(([p, d, c], i) => {
    const y = 2.95 + i * 0.95;
    pill(s, p, CX, y + 0.12, c, 1.9);
    richBullets(s, [d], { x: CX + 2.1, y, w: CW - 2.1, h: 0.8, size: 15, bullet: false, valign: "middle" });
  });
  card(s, { x: CX, y: 5.85, w: CW, h: 0.85, fill: "FBE9DF", border: "F2C6AC" });
  richBullets(s, [[["Bootstrap 4 used ", "b"], ["ml-", "c"], [" and ", "b"], ["mr-", "c"], [". Version 5 uses ", "b"], ["ms-", "c"], [" and ", "b"], ["me-", "c"], [".", "b"], " Old tutorials won't work."]], { x: CX + 0.25, y: 5.85, w: CW - 0.5, h: 0.85, size: 14, bullet: false, valign: "middle" });
}

quiz({
  q: "What does `class=\"px-3 mb-0\"` do?",
  options: ["3px of padding on the left and right, no bottom margin", "1rem of padding on the top and bottom, no bottom margin", "1rem of margin on the left and right, no bottom padding", "1rem of padding on the left and right, no bottom margin"], answer: 3,
  why: ["`p` is padding, `x` is left and right, and size `3` is 1rem (16px), not 3px.", "`m` is margin, `b` is bottom, and `0` is none. Handy for taking the space off the last paragraph in a box."],
  notes: "Two traps: 3 is a size step, not pixels, and x is left and right (y is top and bottom).",
  revealNotes: "Answer: D.\n\nCheck in DevTools: inspect a box with px-3 and the Computed tab shows padding-left: 16px and padding-right: 16px.",
});

trySlide({
  n: 1, title: "A Colour Grid, Together", time: "15 min", codeTitle: "INSIDE <BODY>, IN A NEW FILE", lang: "html",
  steps: [
    ["Make a new file", ["Save Example 01 as ", ["my_grid.html", "c"], ". Keep the head. Empty the body."]],
    ["A container and a row", ["A ", ["container", "c"], ", then a ", ["row", "c"], " inside it."]],
    ["Three columns", ["Equal widths from md up: ", ["col-md-4", "c"], ". Stacked on a phone."]],
    ["Colour each one", [["text-bg-primary", "c"], ", ", ["-success", "c"], ", ", ["-warning", "c"], ", plus ", ["p-4", "c"], "."]],
    ["Add a heading and a button", ["Then drag the window narrower. Do they stack?"]],
  ],
  lw: 5.4,
  code: upTo(bodyOf("try_it_1_answer.html"), '<div class="row g-3">', "A light button").map((l) => l.replace(/<p>[^<]*<\/p>/, "<p>...</p>")).concat(["      </div>", "   </div>", "   <!-- ...two more columns, the same way -->", "</div>"]),
  tip: [[["Why ", "b"], ["text-bg-", "c"], [" and not ", "b"], ["bg-", "c"], ["? ", "b"], "It picks a text colour you can read, even on yellow."]],
  notes: "This is the 'Activity - together' from last year's slides. Do it together: you type, they type along, one step at a time.\n\nOne change from the original: it asked for bg-primary, bg-success and bg-warning. Use text-bg-* instead, so the text is readable (white on blue and green, black on yellow).\n\nWhy the colour goes on an inner div, not on the column: the column has padding for the gutter, and a coloured column would make the boxes touch. g-3 on the row adds space between them.\n\nThe finished page is try_it_1_answer.html.\n\nFast finishers: make the middle column wider (col-md-6, with col-md-3 on each side), or add a fourth column that goes 2 across on md and 4 across on lg.",
});

{
  const s = panelSlide("Class 1 Recap",
    "If students remember three things: link your CSS after Bootstrap's, rows have 12 columns, and col-md-4 means 'from 768px up'.\n\nNext class: a wireframe as a Bootstrap grid, Bootstrap's components, and a layout of our own.\n\nIf there's time, have students start reading through getbootstrap.com/docs/5.3/layout/grid/.");
  labelBullets(s, [
    { label: "What we covered:", text: [
      "A framework is CSS someone else wrote. You use it with class names.",
      "Two kinds: components (Bootstrap) and utility-first (Tailwind).",
      "Bootstrap's CSS in the head, yours after it, the JavaScript at the end.",
      "A container holds rows. Each row has 12 columns.",
      "col-md-4: a third of the row, from 768px and up.",
      "text-bg-*, p-3, mb-2: colour and spacing utilities.",
    ] },
    { label: "Next class:", text: ["A wireframe as a Bootstrap grid, then Bootstrap's components."] },
  ], { size: 20, sub: 16 });
}

// =====================================================================
// WEEK 10, CLASS 2: Components and Your Layout
section("Class 2 - Components and Your Layout");

dividerSlide("Week 10, Class 2", "A wireframe as a grid   ·   components   ·   your layout   ·   the component library",
  "Welcome back. Last class was the grid and the utilities. Today: a warm-up that turns a wireframe into a Bootstrap grid, Bootstrap's ready-made components, and our own layout with a navbar and cards.\n\nNext week: making Bootstrap look like your site, a look at Tailwind, and Harbourside rebuilt with Bootstrap.");

setupSlide("Is Your Environment Set Up?", [
  ["GIT", "Did you run git pull since last class?"],
  ["VS CODE", "Is the 10_CSS_Frameworks folder open in VS Code?"],
  ["LIVE SERVER", "Does layout/01_start.html open with Go Live? (It's blank.)"],
  ["OUTLOOK AND TEAMS", "Have you checked your school email and the Teams channel?"],
], "Take attendance while everyone checks these.\n\nlayout/01_start.html is the warm-up's starting page: Bootstrap is linked, but the body is empty, so the page is blank. That's correct.\n\nToday needs the internet again, for Bootstrap's CDN.");

recallSlide("Recall: What Did We Do Last Class?", (s) => {
  const x = 0.75, w = 5.2, cw = w / 12;
  card(s, { x: x - 0.15, y: 1.6, w: w + 0.3, h: 3.85, border: C.teal, bw: 1.5 });
  for (let i = 0; i < 12; i++) strip(s, x + i * cw, 1.85, cw, 0.5, i % 2 ? C.teal : "1F7A8C", String(i + 1), 11);
  [[[4, 8]], [[6, 6]], [[12]]].forEach(([spans], r) => {
    let cx = x;
    spans.forEach((n, i) => { strip(s, cx, 2.6 + r * 0.75, n * cw, 0.6, [C.orange, C.purple][i], `col-md-${n}`, 13); cx += n * cw; });
  });
  txt(s, "Every row has 12 columns. The numbers add up to 12.", { x: x - 0.15, y: 4.95, w: w + 0.3, h: 0.4, fontSize: 13, italic: true, color: C.grey, align: "center" });
}, [
  ["Linking Bootstrap", ["Its CSS in the head, ", ["yours after it", "b"], ", its JavaScript at the end."]],
  ["Container, row, columns", ["A ", ["row", "c"], " is a flexbox with 12 columns."]],
  ["Breakpoints", [["col-md-4", "c"], ": a third of the row ", ["from 768px up", "b"], ". Stacked under that."]],
  ["Utilities", [["text-bg-primary", "c"], ", ", ["p-3", "c"], ", ", ["mb-2", "c"], ": colour and spacing with classes."]],
], "Same as always: 30 seconds with a neighbour, then take answers, then go through the four points.\n\nQuick check while you're here: on a phone, what does col-md-4 do? (Nothing: the column is full width.)");

// ---- Warm-up: a wireframe, as a Bootstrap grid (layout/) ----
// The body of each layout page, as students' code (2 tabs of indent taken off)
const LAY1 = bodyOf("layout/01_layout_answer.html");
const LAY2 = bodyOf("layout/02_navbar_answer.html");
const LAY3 = bodyOf("layout/03_cards_answer.html");

{
  const s = activitySlide("Warm-up: From Wireframe to Grid (2 min)",
    "The warm-up. Students saw wireframes back in Week 3; this is the step from a wireframe to code.\n\nGive them two minutes with a neighbour to answer the four questions from the picture. The faint orange stripes are Bootstrap's 12 columns, so they can count.\n\nAnswers:\n1. Three rows: the header, then the sidebar and main together, then the footer.\n2. The sidebar covers 3 of the 12 columns.\n3. The main area covers the other 9. 3 + 9 = 12.\n4. On a phone everything is full width, stacked in order: header, sidebar, main, footer. So the columns need a breakpoint: col-md-3 and col-md-9, not col-3 and col-9 (which would stay side by side on a phone, squashed).\n\nThe wireframe is layout/00_wireframe.svg. It opens in any browser.");
  framedImage(s, SH.wire, 0.75, 1.75, 7.9, "THE WIREFRAME · LAYOUT/00_WIREFRAME.SVG");
  const rx = 9.0, rw = W - rx - 0.6;
  card(s, { x: rx, y: 1.45, w: rw, h: 5.5 });
  txt(s, "Count the columns:", { x: rx + 0.3, y: 1.65, w: rw - 0.6, h: 0.4, fontSize: 18, bold: true, color: C.black });
  [["How many rows?", "The bands that go all the way across."],
    ["The sidebar?", "How many of the 12 columns?"],
    ["The main area?", "It has to add up to 12."],
    ["On a phone?", "What happens, and which class does that?"]].forEach(([t, d], i) => {
    const y = 2.3 + i * 1.12;
    numCircle(s, i + 1, rx + 0.3, y, C.orange, 0.44);
    txt(s, t, { x: rx + 0.9, y: y - 0.02, w: rw - 1.1, h: 0.36, fontSize: 16, bold: true, color: C.black });
    txt(s, d, { x: rx + 0.9, y: y + 0.34, w: rw - 1.1, h: 0.6, fontSize: 13.5, color: C.grey, valign: "top" });
  });
}

trySlide({
  label: "Warm-up: Build It, Together", time: "15 min", lw: 5.5, codeTitle: "INSIDE <BODY>, ABOVE THE SCRIPT", lang: "html",
  steps: [
    ["Open the starter", ["Save ", ["layout/01_start.html", "c"], " as ", ["my_layout.html", "c"], ". The links and the viewport tag are done."]],
    ["One container", ["A ", ["<div class=\"container\">", "c"], " around the whole page."]],
    ["Row 1: the header", ["A ", ["row", "c"], " with one ", ["col-12", "c"], " in it. Put the h1 inside."]],
    ["Row 2: sidebar and main", [["col-md-3", "c"], " + ", ["col-md-9", "c"], " = 12. Use ", ["<aside>", "c"], " and ", ["<main>", "c"], "."]],
    ["Row 3: the footer", ["Another ", ["col-12", "c"], "."]],
    ["Colour every part", [["text-bg-primary", "c"], ", ", ["-warning", "c"], ", ", ["-info", "c"], ", ", ["-dark", "c"], ", and ", ["p-3", "c"], ". Then drag the window narrower."]],
  ],
  code: between(LAY1, '<div class="container">', "<!-- Bootstrap's JavaScript")
    .filter((l) => !l.includes("Link three") && !l.includes("<p>Make the window"))
    .map((l) => l.replace(/<p>[^<]*<\/p>/, "<p>...</p>")),
  notes: "Do this together: you type, they type along, one row at a time. Save after each row and look.\n\nIt's meant to look like a wireframe, not a finished site: the bright colours are only there so you can see where each part is. Don't let anyone spend time making it pretty yet.\n\nThings to point out as you go:\n- The semantic elements are still here. header, aside, main and footer can carry Bootstrap's classes themselves: <header class=\"row\">, <aside class=\"col-md-3\">.\n- Every row adds up to 12: 12, then 3 + 9, then 12.\n- text-bg-* sets a background and a readable text colour (Class 1's colour grid). p-3 adds padding so the words aren't on the edge.\n- The colour is on the columns, so the parts touch. That's fine for a layout. (On the colour grid we put it on an inner div, so the gutters showed.)\n- Drag the window under 768px: everything stacks, like the phone wireframe.\n\nThe finished page is layout/01_layout_answer.html.\n\nFast finishers: make the sidebar narrower on large screens (col-md-3 col-lg-2, and col-md-9 col-lg-10). Or put the sidebar on the right: swap the aside and main.",
});

{
  const s = activitySlide("Warm-up: The Wireframe, in Bootstrap",
    "The answer, side by side with the wireframe. Every box in the wireframe became a column, and every band across became a row.\n\nOpen layout/01_layout_answer.html and drag the window narrower: at 767px it changes to the phone wireframe, everything stacked.\n\nThis is the plan for the rest of the day's 'Your Layout' work: next we replace the h1 with a real navbar (Part 2), then put cards in the main area (Part 3). Students who didn't finish can start Part 2 from layout/02_navbar_start.html, which is this page.");
  const gap = 0.35, w = W - 1.5;
  const h = Math.min(3.45, (w - gap) / (1 / SH.wire.ratio + 1 / SH.lay1.ratio));
  framedImage(s, SH.wire, 0.75, 1.75, h / SH.wire.ratio, "THE WIREFRAME");
  framedImage(s, SH.lay1, 0.75 + h / SH.wire.ratio + gap, 1.75, h / SH.lay1.ratio, "IN BOOTSTRAP · 1000PX");
  const items = [["<header class=\"row\">", "col-12", "All 12 columns.", BS.primary], ["<aside>", "col-md-3", "3 of 12, from 768px up.", BS.warning],
    ["<main>", "col-md-9", "The other 9. 3 + 9 = 12.", BS.info], ["<footer class=\"row\">", "col-12", "All 12 columns.", BS.dark]];
  const iy = 1.75 + h + 0.3, iw = (w - 3 * 0.2) / 4, ih = 1.45;
  items.forEach(([el, cls, d, c], i) => {
    const x = 0.75 + i * (iw + 0.2);
    card(s, { x, y: iy, w: iw, h: ih });
    s.addShape(pres.shapes.RECTANGLE, { x, y: iy, w: iw, h: 0.12, fill: { color: c }, line: { color: c } });
    txt(s, el, { x: x + 0.2, y: iy + 0.22, w: iw - 0.4, h: 0.3, fontSize: 12, fontFace: MONO, color: C.grey });
    txt(s, cls, { x: x + 0.2, y: iy + 0.52, w: iw - 0.4, h: 0.4, fontSize: 19, fontFace: MONO, bold: true, color: "8A2B0B" });
    txt(s, d, { x: x + 0.2, y: iy + 0.95, w: iw - 0.4, h: 0.35, fontSize: 13, color: C.dark });
  });
  card(s, { x: 0.75, y: iy + ih + 0.2, w, h: 0.6, fill: C.dark, border: C.dark });
  richBullets(s, [[["Under 768px: ", "b"], "the ", ["md", "c"], " classes switch off, and every part is full width, like the phone wireframe."]],
    { x: 1.05, y: iy + ih + 0.2, w: w - 0.6, h: 0.6, size: 15, bullet: false, color: C.white, valign: "middle", codeColor: "F2C6AC" });
}

overviewSlide("Week 10, Class 2: Overview", "Components and Your Layout", [
  "Warm-up: a wireframe, as a Bootstrap grid",
  "Bootstrap's components: buttons, alerts, cards, the navbar",
  "Your layout: add a navbar, then cards",
  "The rest of Bootstrap's components, and where to look them up",
], [["Plan", "A wireframe"], ["Use", "Components"], ["Build", "Your layout"]],
  "Rough plan:\n- 5 min: environment, recall\n- 15 min: warm-up, the wireframe as a grid\n- 5 min: overview\n- 15 min: components, Examples 05 and 06, Pop Quiz 7\n- 20 min: your layout, Part 2 (navbar) and Part 3 (cards)\n- 10 min: the component library: fast finishers add one to their layout\n- 10 min: show and hide, where to look things up, resources\n- 5 min: recap and the one-sentence summary\n\nAbout 85 minutes. If there's time left, students start the component library's 'Your turn' or look at Assignment 4.\n\nCustomizing, Tailwind and Harbourside moved to Week 11 (11_Customizing_Bootstrap_Tailwind).");

dividerSlide("Components", "Buttons   ·   alerts   ·   cards   ·   the navbar",
  "Components are Bootstrap's ready-made parts. You copy the HTML from the docs, then change the words.");

codeSlide({
  kicker: "COMPONENTS · EXAMPLE 05", title: "Buttons and Alerts", file: "10_CSS_Frameworks/05_components_example.html",
  lines: upTo(bodyOf("05_components_example.html"), "<!-- 1. Buttons", "</div>"), lang: "html", result: SH.ex05, size: 14,
  bullets: [
    [["btn", "c"], " plus a colour: ", ["btn-primary", "c"], ", ", ["btn-outline-danger", "c"], ", ", ["btn-lg", "c"], "."],
    ["Works on a ", ["<button>", "c"], " or a link."],
  ],
  notes: "Open 05_components_example.html.\n\nButtons: always two classes, btn and a style. btn alone gives you the shape and padding with no colour. Add btn-lg or btn-sm for size.\n\nUse a <button> for actions on the page and an <a> for going somewhere. The classes look the same on both, so pick the right element for the job (Week 2: semantic HTML).\n\nAlerts: alert plus a colour. role=\"alert\" makes screen readers announce it straight away, so only use it for messages that matter right now.\n\nThe colour names are the same ones as last class's text-bg-* classes.",
});

codeSlide({
  kicker: "COMPONENTS · EXAMPLE 05", title: "Cards", file: "10_CSS_Frameworks/05_components_example.html",
  lines: upTo(bodyOf("05_components_example.html"), '<div class="row g-3">', "</div>", { skip: 11 }).concat(["   <!-- ...and the Lobster Roll card -->", "</div>"]),
  lang: "html", result: SH.ex05, size: 14,
  bullets: [
    [["card", "c"], " is the box. ", ["card-body", "c"], " adds the padding inside."],
    ["Put cards ", ["in grid columns", "b"], " to line them up."],
    [["badge", "c"], " is the small label."],
  ],
  notes: "A card is a box with a border and rounded corners. Inside: card-body for padding, card-title and card-text for spacing. It can also have card-header, card-footer and card-img-top for a picture at the top.\n\nCards don't set their own width. They fill whatever column they're in. So the pattern is always: row, then col-md-6 (or whatever), then the card inside it. That's how Harbourside's menu will work next week.\n\nLast year's slides listed 'wells'. Wells were Bootstrap 3, and cards replaced them in Bootstrap 4.\n\nThe headings use class=\"h5\": it makes an h3 look like an h5, so you can keep the right heading order for screen readers and still pick the size.",
});

codeSlide({
  kicker: "COMPONENTS · EXAMPLE 06", title: "The Navbar", file: "10_CSS_Frameworks/06_navbar_example.html",
  lines: (() => { const L = upTo(bodyOf("06_navbar_example.html"), "<!-- navbar-expand-md", "</nav>"); const i = L.findIndex((l) => l.includes("aria-current")); L.splice(i - 1, 3); return L; })(),
  lang: "html", frames: [[{ img: SH.ex06d, label: "1000PX" }], [{ img: SH.ex06p, label: "360PX" }]], maxFrameH: 3.6, size: 14,
  bullets: [
    [["navbar-expand-md", "c"], ": the full menu from md up."],
    ["Under that, a menu button opens it. That needs ", ["Bootstrap's JavaScript", "b"], "."],
  ],
  notes: "This is the navbar from last year's slides, updated for Bootstrap 5. (The slide skips some of the list items to fit; the file has them all.)\n\nOpen 06_navbar_example.html and drag the window narrower. At 767px the links disappear and the menu button (the 'hamburger') shows up. Click it and the links slide open.\n\nThe parts:\n- navbar-expand-md: when to show the full menu. Change it to -lg and the menu button stays until 992px.\n- navbar-brand: the site name or logo.\n- navbar-toggler: the menu button. data-bs-toggle=\"collapse\" says what it does, and data-bs-target=\"#mainNav\" says which element it opens. That's the id on the collapse div.\n- aria-controls, aria-expanded and aria-label: for screen readers. Bootstrap's JavaScript flips aria-expanded between false and true for you.\n- ms-auto on the ul pushes the links to the right (margin-start: auto, the flexbox trick from Week 8).\n\nComment out the script tag at the bottom and try the button again: nothing happens.",
});

quiz({
  q: "You copied this menu button from an old tutorial. On a phone, clicking it does nothing. Why?",
  code: ['<button class="navbar-toggler" type="button"', '   data-toggle="collapse"', '   data-target="#mainNav">', '   <span class="navbar-toggler-icon"></span>', "</button>"], lang: "html",
  options: ["Bootstrap 5 renamed them: `data-bs-toggle` and `data-bs-target`", "A button needs `type=\"submit\"` to open a menu", "The navbar needs its own CSS file", "Menus only open on a real phone, not in DevTools"], answer: 0,
  why: ["This is Bootstrap 4 code, from last year's slides. Bootstrap 5's JavaScript only looks for `data-bs-`.", "The other usual cause: the `bootstrap.bundle.min.js` script is missing."],
  notes: "The navbar code on last year's slides was Bootstrap 4. Students will meet this constantly when they copy from Stack Overflow or old tutorials.",
  revealNotes: "Answer: A.\n\nDemo it: in Example 06, change data-bs-toggle to data-toggle and save. The button stops working, and there's no error in the Console, which is what makes it hard to find.\n\nThe lesson: check the version of anything you copy.",
});

// ---- Your layout, parts 2 and 3: the warm-up page gets a navbar, then cards ----
trySlide({
  label: "Your Layout, Part 2: Add a Navbar", time: "10 min", lw: 4.9, codeTitle: "THE HEADER ROW · LAYOUT/02_NAVBAR_ANSWER.HTML", lang: "html",
  steps: [
    ["Open the finished warm-up", [["layout/02_navbar_start.html", "c"], ", or keep going in your own file."]],
    ["Empty the header's column", ["Delete the ", ["<h1>", "c"], ". Change the column's classes to ", ["col-12 p-0", "c"], "."]],
    ["Paste a navbar", ["From Example 06, or type it from the slide. ", ["container-fluid", "c"], " inside it."]],
    ["Make it yours", ["Your site's name in the ", ["navbar-brand", "c"], ", and your own links."]],
    ["Test it", ["Narrow the window. Does the menu button open the links?"]],
  ],
  code: upTo(LAY2, '<header class="row">', "</header>").filter((l) => !l.includes("<!-- p-0")).map((l) => l.replace(' aria-current="page"', "")),
  notes: "Students who didn't finish the warm-up start from layout/02_navbar_start.html: it's the finished warm-up, with a comment in the header saying where the navbar goes.\n\nThe navbar is Example 06's, with three changes:\n- The column it's in gets p-0 (no padding), so the navbar fills it edge to edge. The h1 and text-bg-primary go: the navbar has its own colour, bg-primary, and data-bs-theme=\"dark\" makes its links and menu button light.\n- container-fluid inside the navbar, not container: we're already inside the page's container, and a second fixed-width one would be narrower than the column.\n- navbar-brand holds the site name, the job the h1 had.\n\n(The slide leaves aria-current=\"page\" off the Home link to fit. The answer file has it: it tells screen readers which page you're on.)\n\nIf the menu button does nothing: the script is missing, or data-bs-target doesn't match the id. Pop Quiz 7 again.\n\nFinished page: layout/02_navbar_answer.html.\n\nFast finishers: add a dropdown to the navbar (Bootstrap's docs, Navbar, 'Supported content' has one), or make the sidebar a list group (list-group and list-group-item on the ul and li).",
});

resultSlide("Part 2: What You Should See", [{ img: SH.lay2, label: "1000PX" }, { img: SH.lay2p, label: "360PX, MENU OPEN" }],
  "Left: wide, the links sit on the right (ms-auto). Right: on a phone, the links fold into the menu button. This picture shows it after clicking.\n\nCheck a few screens around the room before moving on. The usual problems:\n- The navbar is narrower than the yellow and blue parts: they used container instead of container-fluid, or left the padding on the column (p-0 fixes it).\n- The menu button shows, but nothing happens: the script, or data-bs-toggle and data-bs-target.\n- Dark links on blue: data-bs-theme=\"dark\" is missing.", 5.0,
  [["Wide: ", "b"], "the links sit on the right (", ["ms-auto", "c"], "). ", ["Narrow: ", "b"], "they fold into the menu button, and it opens."]);

trySlide({
  label: "Your Layout, Part 3: Add Cards", time: "10 min", lw: 5.3, codeTitle: "IN <MAIN>, UNDER THE PARAGRAPHS", lang: "html",
  steps: [
    ["A row inside main", ["Under the paragraphs, a ", ["<div class=\"row g-3\">", "c"], ". Rows can go inside columns."]],
    ["Two columns", [["col-sm-6", "c"], " each: side by side from 576px up."]],
    ["A card in each", [["card", "c"], ", then ", ["card-body", "c"], " with a ", ["card-title", "c"], ", ", ["card-text", "c"], " and a button."]],
    ["A picture on one", ["An ", ["<img class=\"card-img-top\">", "c"], " above the card-body. Try ", ["images/harbour.jpg", "c"], "."]],
    ["Line them up", [["h-100", "c"], " on each card: the same height, even with different words."]],
  ],
  code: upTo(LAY3, '<div class="row g-3">', "</div>", { skip: 8 }).concat([
    '   <div class="col-sm-6">', '      <div class="card h-100">', '         <img src="images/harbour.jpg"', '            class="card-img-top" alt="...">',
    "         <!-- a card-body, like card one -->", "      </div>", "   </div>", "</div>"]),
  notes: "The new idea: a row inside a column. The main column is 9 of 12 wide, and the row inside it has its own 12 columns. So col-sm-6 is half of main, not half of the page.\n\nThe card is from Example 05:\n- card: the box with a border and rounded corners. It's white, so it shows up on the blue main area.\n- card-body: the padding inside. card-title and card-text space the words.\n- card-img-top: a picture across the top, with the top corners rounded to match. It goes inside the card, before card-body.\n- h-100: height 100%, so both cards are as tall as the column. Without it, the card with less in it is shorter.\n\ng-3 on the row puts a gap between the two cards.\n\nalt=\"...\" on the slide is only to fit: write real alt text. The answer file has it.\n\nFinished page: layout/03_cards_answer.html.\n\nFast finishers: a third card, with col-sm-6 col-lg-4 on all three (2 across, then 3 across). Or a card-header and card-footer on a card.",
});

resultSlide("Part 3: What You Should See", [{ img: SH.lay3, label: "1000PX: TWO CARDS IN MAIN" }, { img: SH.lay3p, label: "360PX" }],
  "Left: wide, the two cards share the main column. Right: on a phone, everything is one column again, the cards too. (The phone picture is the whole page, scaled down.)\n\nAsk: between 576px and 767px, what happens? (The sidebar and main stack, because they're md, but the cards are still side by side, because they're sm.) Drag the window to check.\n\nThat's the whole layout: a wireframe, the grid, a navbar and cards. Next: what else Bootstrap has.", 5.1);

componentSlide("Bootstrap's Components (1 of 2): Content and Style", [
  ["Alerts", "A coloured message box: success, warning, danger.", "alerts", "alerts"],
  ["Badge", "A small label or count: New, 3, Today's special.", "badge", "badges"],
  ["Breadcrumb", "Where you are: Home / Menu / Lunch.", "breadcrumb", "pagination"],
  ["Buttons", "btn plus a colour, an outline or a size.", "buttons", "buttons"],
  ["Button group", "Buttons joined into one bar.", "button-group", "button_groups"],
  ["Card", "A box with a body, and a header, footer or picture.", "card", "cards"],
  ["List group", "A list of items or links, in a bordered box.", "list-group", "list_groups"],
  ["Navs and tabs", "Links styled as tabs or pills.", "navs-tabs", "navs"],
  ["Pagination", "Page links: Previous, 1, 2, 3, Next.", "pagination", "pagination"],
  ["Progress", "A progress bar, in any colour, with stripes.", "progress", "progressbars"],
  ["Spinners", "A \"loading\" circle, or growing dots.", "spinners", "spinners"],
  ["Placeholders", "Grey blocks that stand in while content loads.", "placeholders", null],
], "Bootstrap 5.3 has 24 components. These 12 are CSS only: copy the HTML and they work.\n\nThe names are links. Click one to open its page in Bootstrap's docs, and the W3Schools link under it opens W3Schools' version (breadcrumbs are on W3Schools' pagination page; placeholders aren't on W3Schools at all).\n\nShow how to read a docs page: an example at the top, the HTML under it with a copy button, then the variations (colours, sizes) further down. Copy, paste, change the words.\n\nSo far today: alerts, badges, buttons and cards (Example 05) and the navbar (Example 06).\n\nThe 'Your turn' line: students who've finished Part 3 pick one and add it to their layout. Good first choices: a list group in the sidebar, a breadcrumb above the main heading, an alert at the top of main, or a progress bar.");

componentSlide("Bootstrap's Components (2 of 2): The Ones That Move", [
  ["Navbar", "The site menu. Folds into a menu button on phones.", "navbar", "navbar", true],
  ["Dropdowns", "A button or link that opens a list of links.", "dropdowns", "dropdowns", true],
  ["Collapse", "A button that shows and hides a section.", "collapse", "collapse", true],
  ["Accordion", "Stacked sections that open one at a time.", "accordion", "collapse", true],
  ["Carousel", "A slideshow of pictures, with arrows.", "carousel", "carousel", true],
  ["Modal", "A pop-up dialog box over the page.", "modal", "modal", true],
  ["Offcanvas", "A panel that slides in from the side.", "offcanvas", "offcanvas", true],
  ["Tooltips", "A small hint on hover. One line of JS turns them on.", "tooltips", "tooltip", true],
  ["Popovers", "A bigger box on click. One line of JS turns them on.", "popovers", "popover", true],
  ["Toasts", "A small notice in a corner, like a phone's.", "toasts", "toast", true],
  ["Scrollspy", "Lights up the menu link for the part you're reading.", "scrollspy", "scrollspy", true],
  ["Close button", "The × that closes an alert, a modal or a toast.", "close-button", null, false],
], "These need Bootstrap's JavaScript: the bootstrap.bundle.min.js script at the end of the body. Without it they show up but don't do anything, like the navbar's menu button in Pop Quiz 7.\n\nMost of them work with data-bs- attributes only, no JavaScript of our own: data-bs-toggle=\"modal\", data-bs-toggle=\"collapse\", and so on. Tooltips and popovers are the exception: the docs page has one line of JavaScript to turn them on, because they're off by default (they slow down big pages).\n\nGood ones to try in the layout: an accordion in the main area, a modal opened by one of the card buttons, or an offcanvas menu for the sidebar on phones.\n\nThe close button isn't moving on its own: it's the × used inside alerts, modals, toasts and offcanvas, and their JavaScript makes it work.\n\nThe accordion is on W3Schools' Collapse page.");

{
  const s = panelSlide("Show and Hide by Screen Size",
    "From last year's 'Responsive Utilities' slide.\n\nThe display classes: d-none is display: none. d-block, d-inline, d-flex and so on set display. Add a breakpoint to make them responsive, and like the grid, each one applies from that size up.\n\nSo d-none d-md-block reads: hidden, and then from 768px up, shown as a block. Example 03's badges use this to show one badge at each size.\n\nThe opposite, d-md-none: shown on phones, hidden from md up. Handy for a phone-only call button.\n\nBe careful hiding content: if something matters, phone users need it too. Hiding is better for different versions of the same thing.\n\ncontainer-fluid is from the same slide: always the full width, where container stops growing at each breakpoint (max 1320px).");
  codeBox(s, ['<div class="d-none d-md-block">', "   Only on screens 768px and wider", "</div>", "", '<div class="d-md-none">', "   Only on screens under 768px", "</div>"], { x: CX, y: 1.95, w: CW, h: 2.6, maxSize: 15, lang: "html" });
  [["d-none", "Hidden (display: none)."], ["d-md-block", "From 768px up: display: block."], ["container-fluid", "Always the full width. container stops growing."]].forEach(([p, d], i) => {
    const y = 4.8 + i * 0.62;
    pill(s, p, CX, y + 0.05, [C.red, C.green, C.navy][i], 2.3);
    txt(s, d, { x: CX + 2.5, y, w: CW - 2.5, h: 0.45, fontSize: 15, color: C.dark, valign: "middle" });
  });
}

{
  const s = panelSlide("Where to Look Things Up",
    "Nobody memorizes Bootstrap. Professionals have the docs open all the time. Show the docs site quickly: the left menu has Layout (the grid), Content (tables, images), Components and Utilities.\n\nThe Examples page has full page templates (Album, Pricing, Checkout, Product, Cover, Carousel...). Last year's activity was to download one and customize it. It's a fast-finisher option next week, after Harbourside.\n\nThe list from last year's slides, updated:\n- Grids, tables, images, alerts, buttons, navbar, carousel, modal, tooltip and popover are all still in Bootstrap 5.\n- Wells and Glyphicons were Bootstrap 3, and are gone. Use cards instead of wells, and Bootstrap Icons (a separate download, at icons.getbootstrap.com) instead of Glyphicons.\n- Tooltips and popovers need one line of JavaScript to turn them on. The docs show it.\n\nW3Schools has a Bootstrap 5 tutorial too. Check the URL says bootstrap5, not bootstrap (which is version 3).");
  [["Bootstrap's docs", "getbootstrap.com/docs/5.3", "Layout, Components, Utilities. Copy the HTML, change the words.", C.purple],
    ["Bootstrap's examples", "getbootstrap.com/docs/5.3/examples", "Whole pages to start from: Album, Pricing, Product, Cover...", C.orange],
    ["W3Schools", "w3schools.com/bootstrap5", "Short lessons with Try It editors. Check it says bootstrap5.", C.teal]].forEach(([t, u, d, c], i) => {
    const y = 1.95 + i * 1.2;
    card(s, { x: CX, y, w: CW, h: 1.05 });
    s.addShape(pres.shapes.RECTANGLE, { x: CX, y, w: 0.12, h: 1.05, fill: { color: c }, line: { color: c } });
    txt(s, t, { x: CX + 0.35, y: y + 0.1, w: 3, h: 0.4, fontSize: 17, bold: true, color: C.black });
    txt(s, [{ text: u, options: { hyperlink: { url: "https://" + u + "/" } } }], { x: CX + 3.3, y: y + 0.12, w: CW - 3.5, h: 0.35, fontSize: 13, fontFace: MONO, color: C.teal });
    txt(s, d, { x: CX + 0.35, y: y + 0.55, w: CW - 0.6, h: 0.4, fontSize: 14, color: C.dark });
  });
  card(s, { x: CX, y: 5.65, w: CW, h: 1.05, fill: "FBE9DF", border: "F2C6AC" });
  richBullets(s, [[["Gone since Bootstrap 3: ", "b"], "wells (use ", ["card", "c"], ") and Glyphicons (use Bootstrap Icons)."]], { x: CX + 0.25, y: 5.65, w: CW - 0.5, h: 1.05, size: 15, bullet: false, valign: "middle" });
}

{
  const s = plainSlide("Resources",
    "Everything on this slide is a link: the names, the addresses and the red play buttons.\n\nUseful links:\n- Bootstrap's docs: the left menu has Layout (the grid), Content, Forms, Components, Helpers and Utilities. Layout, Components and Utilities are the ones we used.\n- Bootstrap's Examples page: whole page templates (Album, Pricing, Checkout, Dashboard...). Download them all as one zip, or view the source of any of them.\n- W3Schools: short lessons, each with a Try It editor. Make sure the address says bootstrap5. W3Schools' plain /bootstrap/ tutorial is Bootstrap 3.\n- Tailwind's docs, for next week's Tailwind examples.\n\nVideos:\n- thenewboston (Bucky Roberts): a beginner series, short videos. It's Bootstrap 3, from 2015. The grid works the same way (container, row, 12 columns), but some class names have changed: col-xs-6 is col-6 now, and navbar-default, panels and glyphicons are gone. A good test of what we've learned: spot what's changed.\n- Web Dev Simplified (Kyle Cook): the Bootstrap 5 Crash Course covers the grid, six components and the utility classes in about 70 minutes. It's the best match for this week.\n- Web Dev Simplified, Should You Use Bootstrap?: the pros and cons, like our 'Bootstrap, Tailwind, or Plain CSS?' slide.\n- Web Dev Simplified, Bootstrap 5 Will Change The Web: what changed from Bootstrap 4 to 5. Made in 2020, before 5 was finished, but it explains why old code like data-toggle stops working.");
  const cw = (W - 1.5 - 0.35) / 2, ch = 1.18, gap = 0.14;
  [["Useful links", C.teal, false, [
    ["Bootstrap 5.3 documentation", "https://getbootstrap.com/docs/5.3/", "getbootstrap.com/docs/5.3", "Layout, Components and Utilities are the sections we use."],
    ["Bootstrap examples", "https://getbootstrap.com/docs/5.3/examples/", "getbootstrap.com/docs/5.3/examples", "Whole page templates to start from."],
    ["W3Schools Bootstrap 5 tutorial", "https://www.w3schools.com/bootstrap5/", "w3schools.com/bootstrap5", "Short lessons with Try It editors. Check it says bootstrap5."],
    ["Tailwind CSS documentation", "https://tailwindcss.com/docs", "tailwindcss.com/docs", "Every utility class, with what CSS it writes."],
  ]], ["Videos", C.red, true, [
    ["Bootstrap Tutorials for Beginners", "https://www.youtube.com/playlist?list=PL6gx4Cwl9DGBPw1sFodruZUPheWVKchlM", "thenewboston · YouTube playlist", "A playlist of short videos. Bootstrap 3: the grid's the same, some class names aren't."],
    ["Bootstrap 5 Crash Course", "https://www.youtube.com/watch?v=Jyvffr3aCp0", "Web Dev Simplified · YouTube", "The grid, six components and the utilities, in about 70 minutes."],
    ["Should You Use Bootstrap?", "https://www.youtube.com/watch?v=svPDhmXY1Yg", "Web Dev Simplified · YouTube", "When a framework helps, and when it gets in the way."],
    ["Bootstrap 5 Will Change The Web", "https://www.youtube.com/watch?v=pKw-fBvsyYE", "Web Dev Simplified · YouTube", "What changed from Bootstrap 4 to 5, and why old code breaks."],
  ]]].forEach(([head, color, video, links], c) => {
    const x = 0.75 + c * (cw + 0.35);
    txt(s, head.toUpperCase(), { x, y: 1.35, w: cw, h: 0.35, fontSize: 13, bold: true, color, charSpacing: 1 });
    links.forEach(([name, url, show, d], i) => linkCard(s, { x, y: 1.8 + i * (ch + gap), w: cw, h: ch, name, url, show, d, color, video }));
  });
}

{
  const s = panelSlide("Week 10 Recap",
    "If students remember three things from this week: a framework is still CSS underneath, Bootstrap's grid is 12 columns with min-width breakpoints, and a wireframe turns into rows and columns.\n\nThe Week 10 quiz on Brightspace is 7 questions, built from today's and last class's pop quizzes (1 to 7).\n\nThen the last slide: the one-sentence summary.");
  labelBullets(s, [
    { label: "Class 1:", text: [
      "A framework is CSS someone else wrote. It's still CSS underneath.",
      "Bootstrap's CSS first, yours after it. Rows have 12 columns.",
      "col-md-4: a third, from 768px up. text-bg-*, p-3, mb-2.",
    ] },
    { label: "Class 2:", text: [
      "A wireframe becomes rows and columns. Rows can go inside columns.",
      "Components: btn, alert, card, navbar (which needs the JavaScript).",
      "24 components: find one in the docs, copy it, change the words.",
    ] },
    { label: "Coming up:", text: ["The Week 10 quiz on Brightspace: 7 questions, like the pop quizzes.", "Next week: customizing Bootstrap, Tailwind CSS, and Harbourside."] },
  ], { size: 19, sub: 15.5 });
}

dividerSlide("Write a one-sentence summary of what we did this week.", "One sentence, in your own words.",
  "Give everyone a minute. They can write it on paper or post it in the Teams chat. Ask two or three people to read theirs out.\n\nListen for whether anyone mentions that Bootstrap is built on flexbox and media queries. If nobody does, say it: it's the main idea that connects this week to the last three.");

pres.writeFile({ fileName: OUT }).then(() => console.log("wrote " + OUT));
