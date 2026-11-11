// Builds Customizing_Bootstrap_Tailwind.pptx (Week 11, Class 1), in the same style as the Week 10 deck.
// To rebuild: cd 11_Customizing_Bootstrap_Tailwind/_build, run "npm install" once, then "npm run build".
// The helpers are the Week 10 deck's (10_CSS_Frameworks/_build/build_deck.js); the slides moved here from its Class 2.
// The "Result" pictures are screenshots of the example pages, taken with Microsoft Edge (headless) each build.
// Bootstrap and Tailwind load from a CDN, so the build needs the internet.
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
const UNIT = "Week 11: Making Bootstrap Yours";
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
pres.title = "Week 11: Customizing Bootstrap, Tailwind CSS and Harbourside";
pres.author = "Jamie Symonds";
const OUT = path.join(__dirname, "..", "Customizing_Bootstrap_Tailwind.pptx");
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
  ex01: shot("01_customizing_example.html", "ex01", 900, 200),
  // :hover can't be screenshotted, so this copy applies Bootstrap's own .btn:hover rule to both buttons
  ex01h: shot(variant("01_customizing_example.html", "ex01h",
    ".btn.old-way, .btn.new-way { background-color: var(--bs-btn-hover-bg); border-color: var(--bs-btn-hover-border-color); }"), "ex01h", 900, 200),
  ex02: shot("02_bootstrap_cards_example.html", "ex02", 1100, 300),
  ex03: shot("03_tailwind_cards_example.html", "ex03", 1100, 300),
  act0: shot("activity/index.html", "act0", 900, 700),
  actD: shot("activity/answer/index.html", "actD", 1100, 1500),
  actP: phone("activity/answer/index.html", "actP", 360, 700),
};

// Students' index.html is in activity/, not activity/answer/, so the image paths lose the ../
const ANS = fileLines("activity/answer/index.html").map((l) => l.replace("../images/", "images/"));
const ANSCSS = fileLines("activity/answer/style.css");

// =====================================================================
// WEEK 11, CLASS 1: Customizing Bootstrap, Tailwind and Harbourside
section("Class 1 - Customizing Bootstrap, Tailwind and Harbourside");

titleSlide("Week 11, Class 1", "Making Bootstrap look like your site, a look at Tailwind, and Harbourside in Bootstrap.",
  "Welcome back. Last week was Bootstrap's grid and its components: we turned a wireframe into a layout, then gave it a navbar and cards.\n\nThe problem with what we built: it looks like every other Bootstrap site. Today is about fixing that. First, how to change Bootstrap's look with our own CSS. Then a look at Tailwind, the other kind of framework, built for exactly this. Then the big activity: Harbourside, rebuilt with Bootstrap and made to look like Harbourside.\n\nAll of it feeds straight into Assignment 4, which needs a Bootstrap site with your own styling.");

setupSlide("Attendance and a Quick Check-in", [
  ["OUTLOOK", "Are you checking your school Outlook email?"], ["TEAMS", "Are you signed into Teams?"],
  ["TEAMS", "Are you following our Teams channel and General chat?"], ["BRIGHTSPACE", "Have you read Assignment 4 on Brightspace?"],
], "Take attendance while everyone checks these four things.\n\nAssignment 4 is a two-page Bootstrap site. Today's customizing section and Try It 6 are its 'your own styling' requirement, so point that out when you get there.");

setupSlide("Is Your Environment Set Up?", [
  ["GIT", "Did you run git pull? You should see an 11_Customizing_Bootstrap_Tailwind folder."],
  ["VS CODE", "Is the 11_Customizing_Bootstrap_Tailwind folder open in VS Code?"],
  ["LIVE SERVER", "Does activity/index.html open with Go Live? (It has no CSS yet.)"],
  ["INTERNET", "Bootstrap and Tailwind load from the web. Does Example 01 have orange buttons?"],
], "Help anyone who's stuck now, so they aren't lost during the Try Its.\n\nThe Harbourside activity and the Tailwind examples moved here from the Week 10 folder. If someone's git pull complains, they probably edited a file in 10_CSS_Frameworks/activity: they can copy their work somewhere safe and run git stash, then pull again.\n\nactivity/index.html is the Harbourside page with no CSS at all, so it should look plain: big logo, bullet-point menu, Times New Roman. That's correct.");

recallSlide("Recall: What Did We Do Last Week?", (s) => {
  // Last week's layout: header, sidebar, main with two cards, footer
  const x = 0.75, w = 5.2, y = 1.75;
  const bs = { primary: "0D6EFD", warning: "FFC107", info: "0DCAF0", dark: "212529" };
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h: 0.6, fill: { color: bs.primary }, line: { type: "none" } });
  txt(s, "My Site", { x: x + 0.15, y, w: 2, h: 0.6, fontSize: 14, color: C.white, valign: "middle" });
  txt(s, "Home   About   Contact", { x: x + 2.4, y, w: 2.65, h: 0.6, fontSize: 11, color: "DCE8FF", align: "right", valign: "middle" });
  s.addShape(pres.shapes.RECTANGLE, { x, y: y + 0.6, w: w * 0.25, h: 2.45, fill: { color: bs.warning }, line: { type: "none" } });
  txt(s, "aside", { x: x + 0.1, y: y + 0.7, w: 1.1, h: 0.3, fontSize: 11, fontFace: MONO, color: C.black });
  s.addShape(pres.shapes.RECTANGLE, { x: x + w * 0.25, y: y + 0.6, w: w * 0.75, h: 2.45, fill: { color: bs.info }, line: { type: "none" } });
  txt(s, "main", { x: x + w * 0.25 + 0.1, y: y + 0.7, w: 1.1, h: 0.3, fontSize: 11, fontFace: MONO, color: C.black });
  [0, 1].forEach((i) => {
    const cx = x + w * 0.25 + 0.15 + i * 1.88;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cx, y: y + 1.15, w: 1.75, h: 1.7, rectRadius: 0.06, fill: { color: C.white }, line: { color: "D0D7DB" } });
    txt(s, "card", { x: cx + 0.12, y: y + 1.22, w: 1.5, h: 0.3, fontSize: 11, fontFace: MONO, color: C.grey });
  });
  s.addShape(pres.shapes.RECTANGLE, { x, y: y + 3.05, w, h: 0.5, fill: { color: bs.dark }, line: { type: "none" } });
  txt(s, "Your layout: a wireframe, as a grid, with a navbar and cards.", { x: x - 0.15, y: 5.5, w: w + 0.3, h: 0.4, fontSize: 13, italic: true, color: C.grey, align: "center" });
}, [
  ["Linking Bootstrap", ["Its CSS in the head, ", ["yours after it", "b"], ", its JavaScript at the end."]],
  ["The grid", ["A ", ["row", "c"], " has 12 columns. ", ["col-md-3", "c"], " + ", ["col-md-9", "c"], ": side by side from 768px up."]],
  ["Components", [["card", "c"], ", ", ["btn", "c"], ", ", ["alert", "c"], ", and the ", ["navbar", "c"], ", which needs the JavaScript."]],
  ["The docs", ["Find the component, copy the HTML, change the words."]],
], "Same as always: 30 seconds with a neighbour, then take answers, then go through the four points.\n\nThen ask: does anyone's layout from last week look like their own site? (No: it looks like Bootstrap, in very loud colours.) That's today's problem.");

overviewSlide("Week 11, Class 1: Overview", "Making Bootstrap Yours", [
  "Customizing Bootstrap: your CSS after theirs, and its variables",
  "Tailwind CSS: the same cards, built the other way",
  "Bootstrap, Tailwind, or plain CSS?",
  "Rebuild Harbourside with Bootstrap, make it look like Harbourside, then submit",
], [["Customize", "Your own CSS"], ["Compare", "Tailwind"], ["Build", "Harbourside"]],
  "Rough plan:\n- 10 min: check-in, environment, recall, overview\n- 10 min: customizing, Example 01, Pop Quiz 1\n- 20 min: Tailwind, Examples 02 and 03, Pop Quizzes 2 and 3\n- 50 min: Harbourside, Try Its 1 to 6, then submit\n- 5 min: recap and the one-sentence summary\n\nThe activity is the priority. If time is short, demo Tailwind quickly and skip the 'Which one?' slide.\n\nPop quizzes: same as always. 30 seconds to think, then hold up 1 to 4 fingers (A to D), or the Teams chat.");

dividerSlide("Customizing Bootstrap", "Your CSS after theirs   ·   Bootstrap 5's variables",
  "Every Bootstrap site looks like Bootstrap until you change it. The question is how.");

codeSlide({
  kicker: "CUSTOMIZING · EXAMPLE 01", title: "Changing a Button's Colour", file: "11_Customizing_Bootstrap_Tailwind/01_customizing_example.css",
  lines: cssOf("01_customizing_example.html").slice(2), lang: "css", frames: [[{ img: SH.ex01, label: "NORMAL" }], [{ img: SH.ex01h, label: "WHILE YOU HOVER" }]], maxFrameH: 2.9, size: 14,
  bullets: [
    ["Button 1 changes ", ["background-color", "c"], ". Hover, and it goes ", ["blue", "b"], "."],
    ["Button 2 changes Bootstrap's ", ["variables", "b"], ". Every state uses them."],
  ],
  notes: "Open 01_customizing_example.html and hover over each button.\n\nPart 0: overriding any rule works the normal way. Bootstrap's body rule and ours tie, ours is later, ours wins. The font is Georgia.\n\nPart 1 is the way from last year's slides (and lots of tutorials). It looks right until you hover: Bootstrap has a .btn:hover rule, which is more specific than ours, and it sets the background to var(--bs-btn-hover-bg). That's still blue. The bottom picture applies that exact rule to both buttons.\n\nPart 2 is how Bootstrap 5.3 wants it done. A Bootstrap button reads all of its colours from variables (custom properties, the --name ones). Set those, and the normal, hover and active states all follow.\n\nWhere do the variable names come from? Inspect a button in DevTools: the .btn-primary rule lists them all.\n\nFor bigger changes (every primary colour on the site), professional projects use Bootstrap's Sass source. That's beyond this course. Ready-made Bootstrap themes are another option (last year's slides mention buying one); Bootswatch has free ones.",
});

quiz({
  q: "This is in your CSS, after Bootstrap's. What colour is the button while you hover over it?",
  code: [".btn-primary {", "   background-color: #ff5722;", "   border-color: #ff5722;", "}"], codeTitle: "LAST YEAR'S SLIDE",
  options: ["Orange, the same as always", "Bootstrap's blue", "A slightly darker orange", "See-through: it has no colour"], answer: 1,
  why: ["Bootstrap's `.btn:hover` (two classes' worth) beats `.btn-primary` (one), so its hover rule wins.", "That rule uses `var(--bs-btn-hover-bg)`, which is still blue. Change the variables instead: Example 01, part 2."],
  revealImg: SH.ex01h,
  notes: "The code is word for word from last year's Customizing Bootstrap slide. Specificity from Week 6.",
  revealNotes: "Answer: B.\n\nThe picture is Example 01 with Bootstrap's hover rule applied: button 1 is blue.\n\nSpecificity: .btn:hover is a class plus a pseudo-class, so (0,2,0). .btn-primary is (0,1,0). Ours comes later, but specificity is checked first.",
});

dividerSlide("Tailwind CSS", "Utility-first   ·   the same cards, two ways   ·   which one should you use?",
  "Tailwind is the other kind of framework: no ready-made components, just small classes you combine.");

{
  const s = plainSlide("The Same Cards, Two Ways",
    "Examples 02 and 03 are the same three menu cards. Open them side by side.\n\nBootstrap (left): the card is card, card-body, card-title. Someone else decided what a card looks like, and you get it with a few names.\n\nTailwind (right): there's no 'card' class. The article gets bg-white (background), rounded-lg (border-radius), shadow (box-shadow) and p-5 (padding). You decide every part.\n\nThey look nearly the same. Ask: which HTML would you rather read? Which would you rather change to a totally different look?\n\nTailwind's HTML is longer, but you rarely need a CSS file. Bootstrap's HTML is shorter, but making it look different takes CSS, like Example 01.");
  [["BOOTSTRAP · EXAMPLE 02", "02_bootstrap_cards_example.html", SH.ex02, ['<article class="card h-100 shadow-sm">', '   <div class="card-body">', '      <h2 class="card-title h5">...</h2>', '      <p class="card-text">...</p>', '      <a class="btn btn-primary">Order</a>', "   </div>", "</article>"], C.teal],
    ["TAILWIND · EXAMPLE 03", "03_tailwind_cards_example.html", SH.ex03, ['<article class="bg-white rounded-lg shadow p-5">', '   <h2 class="text-lg font-semibold">...</h2>', '   <p class="mt-2">...</p>', '   <a class="inline-block mt-3 px-4 py-2', '      rounded bg-blue-600 text-white', '      hover:bg-blue-700">Order</a>', "</article>"], C.purple]].forEach(([k, f, img, code, c], i) => {
    const x = 0.75 + i * 6.05, w = 5.8;
    txt(s, k, { x, y: 1.35, w, h: 0.32, fontSize: 13, bold: true, color: c, charSpacing: 1 });
    codeBox(s, code, { x, y: 1.75, w, h: 2.55, maxSize: 13, lang: "html" });
    framedImage(s, img, x, 4.75, w, "RESULT");
    txt(s, "11_Customizing_Bootstrap_Tailwind/" + f, { x, y: 4.75 + w * img.ratio + 0.08, w, h: 0.3, fontSize: 11, color: C.mid, fontFace: MONO });
  });
}

codeSlide({
  kicker: "TAILWIND · EXAMPLE 03", title: "Reading Tailwind Classes", file: "11_Customizing_Bootstrap_Tailwind/03_tailwind_cards_example.html",
  lines: [...upTo(bodyOf("03_tailwind_cards_example.html"), '<div class="grid', "</article>")],
  lang: "html", result: SH.ex03, size: 14,
  bullets: [
    ["One ", ["<script>", "c"], " in the head loads it (the \"Play CDN\")."],
    ["About one CSS property per class: ", ["p-5", "c"], " is padding, ", ["rounded-lg", "c"], " is border-radius."],
    [["md:grid-cols-3", "c"], ": a prefix is a ", ["media query", "b"], ". ", ["hover:", "c"], " is ", [":hover", "c"], "."],
  ],
  notes: "Open 03_tailwind_cards_example.html.\n\nThe script tag: Tailwind's 'Play CDN'. It reads the classes on the page and writes the CSS in the browser, as the page loads. It's great for learning and trying ideas. Real projects run Tailwind as a build step, which writes a small CSS file with only the classes the site uses. That's why Tailwind sites end up smaller than Bootstrap ones.\n\nRead the grid line: grid (display: grid), gap-4 (a 1rem gap), md:grid-cols-3 (three columns, from 768px up). Tailwind's md is 768px, same as Bootstrap's, and also min-width: phones first. Below md there's no grid-cols, so it's one column.\n\nOn the button: hover:bg-blue-700 is a darker blue on hover. Tailwind puts the state in the class name, which is why there's no hover problem like Example 01's.\n\nThe colours are a palette: blue-600 is a mid blue, blue-700 a bit darker. 50 is the lightest, 950 the darkest.\n\nTailwind's spacing numbers are different from Bootstrap's: p-5 here is 1.25rem (each step is 0.25rem). Same idea, different scale, which is a good reason to check the docs.",
});

quiz({
  q: "In Tailwind, what does `md:grid-cols-3` mean?",
  options: ["Three columns, but only on medium-sized screens", "Three columns, on screens up to 768px wide", "Three columns of medium width", "Three columns, on screens 768px wide and wider"], answer: 3,
  why: ["Like Bootstrap, Tailwind is phones first: `md:` is `@media (min-width: 768px)`.", "It applies from 768px up, including big screens, until a bigger prefix like `lg:` overrides it."],
  notes: "Students who got Week 10's Pop Quiz 4 (col-md-4 on a phone) should get this. A is the common wrong answer.",
  revealNotes: "Answer: D.\n\nCheck in DevTools: inspect the grid in Example 03 and the rule is inside @media (width >= 48rem). 48rem is 768px.",
});

{
  const s = panelSlide("Bootstrap, Tailwind, or Plain CSS?",
    "There's no single right answer, and students will hear strong opinions online.\n\n- Bootstrap: when you need something that works and looks decent quickly, especially with menus, modals and forms. Very common on existing sites, admin pages and internal tools, so you'll likely maintain some at work.\n- Tailwind: when you have a design to build, and want it to look like no one else's. Very popular for new projects, and often used with JavaScript frameworks like React. Its build step makes small CSS files.\n- Plain CSS: small sites, and anywhere you want full control with no extra files. Also: everything you learn in plain CSS makes both frameworks easier.\n\nWhy we demoed these two: they're the most common, they're one of each kind, and between them students will recognize almost any framework they meet.");
  [["Bootstrap", "Pick it when you need it working fast: menus, forms and pages that look good with little CSS.", C.teal],
    ["Tailwind", "Pick it when you have your own design to build, and want it to look like nobody else's.", C.purple],
    ["Plain CSS", "Pick it for small sites, or full control with no extra files. And it's what both are made of.", C.orange]].forEach(([t, d, c], i) => {
    const y = 1.95 + i * 1.45;
    card(s, { x: CX, y, w: CW, h: 1.25 });
    s.addShape(pres.shapes.RECTANGLE, { x: CX, y, w: 0.12, h: 1.25, fill: { color: c }, line: { color: c } });
    txt(s, t, { x: CX + 0.35, y: y + 0.12, w: 3, h: 0.4, fontSize: 19, bold: true, color: c });
    txt(s, d, { x: CX + 0.35, y: y + 0.55, w: CW - 0.6, h: 0.6, fontSize: 15, color: C.dark, valign: "top" });
  });
  card(s, { x: CX, y: 6.35 - 0.05, w: CW, h: 0.45, fill: C.light, border: C.light });
  txt(s, "Learn one of each kind, and the rest will look familiar.", { x: CX, y: 6.3, w: CW, h: 0.45, fontSize: 15, italic: true, color: C.grey });
}

quiz({
  q: "Which of these was written with Tailwind CSS?",
  options: ['`class="bg-white rounded-lg shadow p-5"`', '`class="card h-100 shadow-sm"`', '`class="btn btn-primary"`', '`class="alert alert-warning"`'], answer: 0,
  why: ["Each Tailwind class does about one thing: a background, a radius, a shadow, padding.", "`card`, `btn` and `alert` are Bootstrap's named components. (`shadow-sm` is a Bootstrap utility, but `card` gives it away.)"],
  notes: "A last check before the activity.",
  revealNotes: "Answer: A.\n\nIt's the card from Example 03. B is the card from Example 02.",
});

dividerSlide("Harbourside in Bootstrap", "Same page, same words   ·   Bootstrap does the layout   ·   Try Its 1 to 6",
  "The big activity. This replaces last year's 'download an example and customize it', which is now a fast-finisher option.\n\nStudents take the Harbourside page from Week 7, with all the CSS taken away, and rebuild it with Bootstrap. Each Try It adds one part.");

{
  const s = activitySlide("Harbourside, Part 4: Before and After",
    "Open activity/index.html with Go Live. It's our Harbourside page with no CSS at all: that's the left picture.\n\nBy the end of class it looks like the right pictures, and on a phone the menu folds into a menu button.\n\nThe plan:\n- Try It 1: link Bootstrap and add a container.\n- Try It 2: the navbar.\n- Try It 3: a banner, and the intro in two columns.\n- Try It 4: the menu as cards.\n- Try It 5: the hours table and the footer.\n- Try It 6: make it look like Harbourside, in style.css.\n\nEvery Try It's code is on its slide, and the finished page is in activity/answer/. The comments in index.html say where each Try It goes.\n\nThe content and the semantic HTML (header, nav, main, section, article, aside, footer) all stay. Bootstrap only adds classes and a few wrapper divs.");
  const ih = 5.3;
  framedImage(s, SH.act0, 0.75, 1.75, Math.min(4.2, ih / SH.act0.ratio), "BEFORE: NO CSS");
  const crop = { path: SH.actD.path, ratio: SH.actD.ratio };
  framedImage(s, crop, 5.35, 1.75, ih / crop.ratio, "AFTER: BOOTSTRAP");
  phoneFrame(s, SH.actP, 5.35 + ih / crop.ratio + 0.35, 1.75, 2.15, "PHONE", C.green);
}

trySlide({
  n: 1, title: "Link Bootstrap", time: "4 min", lw: 5.0, codeTitle: "IN THE HEAD, WHERE THE TRY IT 2 COMMENT IS", lang: "html",
  steps: [
    ["Open the activity", [["activity/index.html", "c"], " with Go Live, and in VS Code."]],
    ["Add the two links", "Bootstrap's first, then style.css. Save: the fonts change right away."],
    ["Add a container", [["<main class=\"container my-5\">", "c"], ": centred, with space above and below."]],
  ],
  code: upTo(ANS, '<link rel="stylesheet" href="https://cdn', 'href="style.css"').concat(["", "<!-- further down -->", '<main class="container my-5">']),
  tip: [[["Watch the page as you save: ", "b"], "it changes before you've added a single other class."]],
  notes: "Bootstrap's base styles (its 'Reboot') change the page as soon as it's linked: a modern font, no margin on the body, nicer tables. Point that out.\n\nmy-5: margin on the top and bottom (y), size 5 (3rem).\n\nCheck everyone has the link right before moving on: if Bootstrap didn't load, nothing else today works.",
});

trySlide({
  n: 2, title: "The Navbar", time: "10 min", codeTitle: "REPLACES THE LOGO AND THE OLD <NAV>", lang: "html",
  steps: [
    ["Delete the old parts", ["The ", ["<img>", "c"], " logo and the whole ", ["<nav>", "c"], ". Keep the ", ["<h1>", "c"], " for now."]],
    ["Paste the navbar", "In the header, where they were."],
    ["Add the script", ["Bootstrap's JavaScript just before ", ["</body>", "c"], ", like Week 10's Example 01."]],
    ["Test it", "Make the window narrow. Does the menu button open? Does the Menu dropdown work?"],
  ],
  lw: 4.0,
  code: [
    '<nav class="navbar navbar-expand-md bg-harbour"',
    '     data-bs-theme="dark" aria-label="Main">',
    '  <div class="container">',
    '    <a class="navbar-brand d-flex align-items-center gap-2" href="#">',
    '      <img src="images/logo.svg" width="40" height="40" alt="">',
    '      Harbourside',
    '    </a>',
    '    <button class="navbar-toggler" type="button"',
    '        data-bs-toggle="collapse" data-bs-target="#mainNav"',
    '        aria-controls="mainNav" aria-expanded="false"',
    '        aria-label="Toggle navigation">',
    '      <span class="navbar-toggler-icon"></span>',
    '    </button>',
    '    <div class="collapse navbar-collapse" id="mainNav">',
    '      <ul class="navbar-nav ms-auto">',
    '        <li class="nav-item dropdown">',
    '          <a class="nav-link dropdown-toggle" href="#menu" role="button"',
    '              data-bs-toggle="dropdown" aria-expanded="false">Menu</a>',
    '          <ul class="dropdown-menu">',
    '            <li><a class="dropdown-item"',
    '                href="#haddock">Haddock &amp; Chips</a></li>',
    '            <!-- ...and the other three, the same way -->',
    '          </ul>',
    '        </li>',
    '        <li class="nav-item"><a class="nav-link" href="#hours">Hours &amp; Location</a></li>',
    '        <li class="nav-item"><a class="nav-link" href="#contact">Contact</a></li>',
    '      </ul>',
    '    </div>',
    '  </div>',
    '</nav>',
  ],
  notes: "The code is on the slide, but it's long: students can copy it from activity/answer/index.html. The slide leaves out three of the dropdown items to fit; the answer file has them all. Careful: the answer file's image paths start with ../ because it's one folder deeper. In activity/index.html it's images/logo.svg.\n\nNew things since Week 10's Example 06 and your layout's navbar:\n- bg-harbour is our own class (it's already in style.css): navy.\n- data-bs-theme=\"dark\" tells Bootstrap the background is dark, so the links and the menu button turn light.\n- The logo goes inside the navbar-brand. d-flex align-items-center gap-2 is flexbox from Week 8, as classes.\n- The Menu item is a dropdown: dropdown-toggle on the link, dropdown-menu on the list, dropdown-item on each link. data-bs-toggle=\"dropdown\" makes it open. It needs the script too.\n\nThe usual problems: the script is missing (nothing opens), or the id on the collapse div doesn't match data-bs-target.",
});

trySlide({
  n: 3, title: "A Banner and Two Columns", time: "8 min", codeTitle: "THE H1 BECOMES A BANNER; THE INTRO BECOMES A ROW", lang: "html",
  steps: [
    ["Make the banner", ["Replace the ", ["<h1>", "c"], " with the banner div, under the navbar."]],
    ["Make the intro a row", ["Add ", ["row g-4 align-items-center", "c"], " to the intro section."]],
    ["Two columns", ["The words in one ", ["col-md-6", "c"], ", the figure in another."]],
    ["Fit the photo", [["img-fluid", "c"], " is ", ["max-width: 100%", "c"], " from last week."]],
  ],
  lw: 5.2,
  code: [...upTo(ANS, '<div class="bg-harbour text-white', "</div>"), "", ...upTo(ANS, '<section class="intro', '<div class="col-md-6">'), "      <h2>Fresh from the Wharf</h2>", "      <p>...</p>  <p>...</p>", "   </div>", ...upTo(ANS, '<div class="col-md-6">', "</section>", { start: ANS.findIndex((l) => l.includes("<figure")) - 1 }).filter((l) => !l.includes("alt=") && !l.includes("figcaption"))],
  notes: "The banner: display-5 is one of Bootstrap's big heading sizes, lead is the bigger paragraph from Week 10's Example 01. btn-harbour is a class we make in Try It 6. Until then the button has no colour, which is fine.\n\nThe intro: the section itself becomes the row. Two col-md-6 divs: words first, photo second. On a phone they stack, with the words on top.\n\nalign-items-center lines the two columns up in the middle (flexbox, Week 8). g-4 is the gap.\n\nimg-fluid on the photo: it's Bootstrap's version of last week's img { max-width: 100%; height: auto; }. figure, figure-img and figure-caption style the figure.\n\nThe slide skips the alt text and the figcaption to fit: keep them, from the original.",
});

trySlide({
  n: 4, title: "The Menu as Cards", time: "10 min", codeTitle: "INSIDE <SECTION ID=\"MENU\">, UNDER THE H2", lang: "html",
  steps: [
    ["Wrap the cards in a row", [["<div class=\"row g-4\">", "c"], " around all four articles."]],
    ["A column for each card", [["col-sm-6 col-lg-3", "c"], ": 1 across, 2 from sm, 4 from lg."]],
    ["Make each one a card", ["Add ", ["card", "c"], " to the article, and a ", ["card-body", "c"], " div inside it."]],
    ["Today's special", ["Add the ", ["badge", "c"], " to the lobster roll."]],
  ],
  lw: 5.2,
  code: upTo(ANS, '<div class="row g-4">', "</div>", { skip: 8 }).map((l) => l.replace(/(card-text">)[^<]*/, "$1...")).concat(["", "   <!-- ...the same for the other three -->", "   <!-- In the lobster roll's card-body: -->", '   <span class="badge text-bg-warning mb-2">Today\'s special</span>', "</div>"]),
  notes: "Each article keeps its id (for the dropdown links) and its class=\"special\". h-100 makes every card the full height of its column, so they line up even with different amounts of text. shadow-sm adds a light shadow.\n\nclass=\"card-title h5\" on the h3: the right heading level, at a smaller size.\n\nThe column classes: count them at each size. Under 576px, no class applies, so full width. sm (576px+): 6 of 12, so 2 across. lg (992px+): 3 of 12, so 4 across.\n\nCheck with the device toolbar at 360px, 700px and 1100px.",
});

trySlide({
  n: 5, title: "The Table and the Footer", time: "6 min", lw: 5.0, codeTitle: "THE ASIDE, THEN THE FOOTER", lang: "html",
  steps: [
    ["Make the aside a row", ["The address in ", ["col-md-5", "c"], ", the table in ", ["col-md-7", "c"], "."]],
    ["Style the table", [["table table-striped", "c"], ", inside a ", ["table-responsive", "c"], " div."]],
    ["Style the footer", ["Navy, white text, centred, with padding."]],
  ],
  code: [...upTo(ANS, '<aside id="hours"', '<table class="table'), "            ...", "         </table>", "      </div>", "   </div>", "</aside>", "", ...upTo(ANS, '<footer id="contact"', "</footer>")],
  notes: "table: Bootstrap's table style (padding, lines between rows). table-striped: every second row shaded. Others to try: table-hover, table-bordered, table-sm.\n\ntable-responsive: if the table is wider than the screen, it scrolls sideways inside its own box, instead of the whole page scrolling. We did this by hand in Week 7 with the table-wrap div.\n\nThe footer reuses bg-harbour. text-center centres the text, py-4 is padding on the top and bottom. mb-1 and mb-0 tidy up the paragraph margins.\n\nThe caption shows under the table: Bootstrap puts captions at the bottom. Add caption-top to the table to move it.",
});

trySlide({
  n: 6, title: "Make It Yours", time: "8 min", lw: 5.6, codeTitle: "ADD TO STYLE.CSS, UNDER THE TRY IT 7 HEADING",
  steps: [
    ["Open style.css", ["It's linked after Bootstrap, so a tie goes to us."]],
    ["Add the rules on the right", "Georgia headings, orange-red prices, the special, our button."],
    ["Hover the banner button", ["It should stay orange-red: we set the ", ["variables", "b"], ", like Example 01."]],
    ["Make it yours", "Change a colour, a font, or a Bootstrap class. Anything you like."],
  ],
  code: upTo(ANSCSS, "/* 1. Headings", "font-weight: bold;").concat(["}", "", "/* 3. and 4.: the special, and .btn-harbour", "   (copy them from answer/style.css) */"]),
  tip: [[["Bootstrap's rule still winning? ", "b"], "Inspect it in DevTools. It may be more specific, like ", [".btn:hover", "c"], "."]],
  notes: "This is what turns a Bootstrap page into Harbourside. The whole file is about 40 lines, compared with about 300 lines in Week 9's style.css. That's the trade-off in one number.\n\nThe special: .special sets background-color and border-color. Bootstrap's .card rule uses the same specificity (one class), so ours wins because it's later.\n\n.btn-harbour sets the button variables, so the hover and active colours are ours too. Orange-red (#B5451B) and not the logo's lighter orange (#E97132): white text on the lighter orange is too low contrast to read easily.\n\nFast finishers: make the navbar stick to the top (sticky-top on the nav), add a card-img-top photo to a card, or swap btn-harbour for a btn-outline-light.",
});

{
  const s = activitySlide("Class Activity: Submit Harbourside",
    "This replaces last year's activity (download a Bootstrap example and customize it), which is now the fast-finisher option.\n\nStudents take the screenshots with the device toolbar, so the phone width shows. Its menu (the three dots on the toolbar) has 'Capture screenshot', and 'Capture full size screenshot' gets the whole page.\n\nCheck that the Week 11 Brightspace submission box is set up before class.\n\nTailwind version: for students who want a challenge, rebuild the menu cards with Tailwind's Play CDN, using Example 03 as a guide. Don't link Bootstrap and Tailwind in the same page: their classes clash (both have p-3, for example).");
  const LW = 6.05;
  card(s, { x: 0.75, y: 1.45, w: LW, h: 5.5 });
  txt(s, "Your page should have:", { x: 1.05, y: 1.65, w: LW - 0.6, h: 0.4, fontSize: 18, bold: true, color: C.black });
  ["A navbar that folds into a menu button", "The menu as cards in a responsive grid", "Your own colours in style.css"].forEach((t, i) => {
    const y = 2.2 + i * 0.6;
    numCircle(s, "✓", 1.05, y, C.teal, 0.42);
    txt(s, t, { x: 1.65, y, w: LW - 1.2, h: 0.42, fontSize: 16, color: C.black, valign: "middle" });
  });
  txt(s, "Then:", { x: 1.05, y: 4.1, w: LW - 0.6, h: 0.4, fontSize: 18, bold: true, color: C.black });
  [[["Take two screenshots: wide, and at phone width (", ["Ctrl + Shift + M", "b"], ")."]], [["Drop them in the box on Week 11 Brightspace."]]].forEach(([d], i) => {
    const y = 4.6 + i * 0.95;
    numCircle(s, i + 1, 1.05, y, C.orange, 0.42);
    richBullets(s, [d], { x: 1.65, y: y - 0.03, w: LW - 1.2, h: 0.8, size: 16, bullet: false, color: C.black });
  });
  const rx = 0.75 + LW + 0.35, rw = W - rx - 0.6;
  txt(s, "Finished early?", { x: rx, y: 1.45, w: rw, h: 0.4, fontSize: 18, bold: true, color: C.black });
  [["Last year's activity", ["Download one of the Bootstrap examples (", ["getbootstrap.com/docs/5.3/examples", "c"], ") and make it yours: a heading, some text, an image."]],
    ["Try Tailwind", ["Rebuild the menu cards in a new file with Tailwind, using Example 03 as a guide."]],
    ["Stuck?", [["activity/answer/", "c"], " has the finished page."]]].forEach(([t, d], i) => {
    const y = 2.0 + i * 1.6;
    card(s, { x: rx, y, w: rw, h: 1.4 });
    txt(s, t, { x: rx + 0.25, y: y + 0.12, w: rw - 0.5, h: 0.38, fontSize: 16, bold: true, color: [C.purple, C.teal, C.orange][i] });
    richBullets(s, [d], { x: rx + 0.25, y: y + 0.5, w: rw - 0.5, h: 0.85, size: 14, bullet: false });
  });
}

{
  const s = plainSlide("Resources",
    "Everything on this slide is a link: the names, the addresses and the red play buttons.\n\nUseful links:\n- Bootstrap's docs: the left menu has Layout (the grid), Content, Forms, Components, Helpers and Utilities. Layout, Components and Utilities are the ones we used.\n- Bootstrap's Examples page: whole page templates (Album, Pricing, Checkout, Dashboard...). Download them all as one zip, or view the source of any of them.\n- W3Schools: short lessons, each with a Try It editor. Make sure the address says bootstrap5. W3Schools' plain /bootstrap/ tutorial is Bootstrap 3.\n- Tailwind's docs, for Examples 02 and 03.\n\nVideos:\n- thenewboston (Bucky Roberts): a beginner series, short videos. It's Bootstrap 3, from 2015. The grid works the same way (container, row, 12 columns), but some class names have changed: col-xs-6 is col-6 now, and navbar-default, panels and glyphicons are gone. A good test of what we've learned: spot what's changed.\n- Web Dev Simplified (Kyle Cook): the Bootstrap 5 Crash Course covers the grid, six components and the utility classes in about 70 minutes. It's the best match for this week.\n- Web Dev Simplified, Should You Use Bootstrap?: the pros and cons, like our 'Bootstrap, Tailwind, or Plain CSS?' slide.\n- Web Dev Simplified, Bootstrap 5 Will Change The Web: what changed from Bootstrap 4 to 5. Made in 2020, before 5 was finished, but it explains why old code like data-toggle stops working.");
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
  const s = panelSlide("Week 11 Recap",
    "If students remember three things: put your CSS after Bootstrap's, change a button's --bs-btn- variables (not its background-color) so hover follows, and Tailwind builds the same things from small classes.\n\nAssignment 4 needs all of this week and last: a navbar, the grid, components, and your own styling in style.css. Harbourside is a good model to copy from.\n\nThen the last slide: the one-sentence summary.");
  labelBullets(s, [
    { label: "Customizing Bootstrap:", text: [
      "Your style.css after Bootstrap's: when rules tie, yours wins.",
      "Buttons: set the --bs-btn- variables, so hover and active follow.",
      "Still losing? Inspect it: Bootstrap's rule may be more specific.",
    ] },
    { label: "Tailwind:", text: [
      "Small utility classes, about one CSS property each.",
      "md: and hover: prefixes: media queries and states in the class name.",
    ] },
    { label: "Coming up:", text: ["Assignment 4: your own two-page Bootstrap site, with your own styling."] },
  ], { size: 19, sub: 15.5 });
}

dividerSlide("Write a one-sentence summary of what we did this week.", "One sentence, in your own words.",
  "Give everyone a minute. They can write it on paper or post it in the Teams chat. Ask two or three people to read theirs out.\n\nListen for whether anyone mentions that Bootstrap is built on flexbox and media queries. If nobody does, say it: it's the main idea that connects this week to the last three.");

pres.writeFile({ fileName: OUT }).then(() => console.log("wrote " + OUT));
