// Builds Responsive_Web_Design.pptx (Week 9, Class 2), in the same style as the Class 5-11, Week 8 and Week 9, Class 1 decks.
// To rebuild: cd 09_Responsive/_build, run "npm install" once, then "npm run build".
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
pres.title = "Week 9, Class 2: Responsive Web Design";
pres.author = "Jamie Symonds";
const OUT = path.join(__dirname, "..", "Responsive_Web_Design.pptx");
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

// ---------- screenshots ----------
const SH = {
  ex01: shot("01_fixed_layout_example.html", "ex01", 1040, 760),
  ex01p: phone("01_fixed_layout_example.html", "ex01p", 360, 640),
  ex02no: noViewport("02_viewport_example.html", "ex02no"),
  ex02yes: phone("02_viewport_example.html", "ex02yes", 360, 640),
  ex03: shot("03_media_query_example.html", "ex03", 800, 270),
  ex03n: shot("03_media_query_example.html", "ex03n", 560, 240),
  ex04p: phone("04_breakpoints_example.html", "ex04p", 360, 480),
  ex04t: shot("04_breakpoints_example.html", "ex04t", 640, 320),
  ex04d: shot("04_breakpoints_example.html", "ex04d", 1000, 220),
  ex05: shot(variant("05_fluid_layout_example.html", "ex05", "h2, .frame { display: none; }"), "ex05", 800, 560),
  ex05p: phone(variant("05_fluid_layout_example.html", "ex05p", ".fixed, .container { display: none; } h2 { margin-top: 0; }"), "ex05p", 360, 210),
  ex06: shot("06_responsive_media_example.html", "ex06", 1000, 300),
  ex07: shot("07_fluid_type_example.html", "ex07", 1000, 330),
  ex07a: shot("07_fluid_type_example.html", "ex07a", 768, 300),
  ex07b: shot("07_fluid_type_example.html", "ex07b", 769, 300),
  ex08: shot("08_responsive_page_example.html", "ex08", 1040, 760),
  ex08p: phone("08_responsive_page_example.html", "ex08p", 360, 760),
};

const CSS08 = cssOf("08_responsive_page_example.html");

// =====================================================================
// WEEK 9, CLASS 2: Responsive Web Design

titleSlide("Week 9, Class 2", "Media queries, fluid layouts, and images and text that fit any screen.",
  "Welcome back! Last class we finished our layout tools with float. Today we make layouts work on any screen, from a phone to a big monitor.\n\nThis is one of the most useful classes in the course: most people will look at the sites students build on a phone.\n\nSame routine as before: pop quizzes all the way through, then students make a page responsive themselves.");

setupSlide("Is Your Environment Set Up?", [
  ["GIT", "Did you run git pull? You should see a 09_Responsive folder."],
  ["VS CODE", "Is the 09_Responsive folder open in VS Code?"],
  ["LIVE SERVER", "Does 01_fixed_layout_example.html open with Go Live?"],
  ["DEVTOOLS", "In Edge or Chrome: F12, then Ctrl + Shift + M. Can you pick a phone?"],
], "Take attendance while everyone checks these four things, and help anyone who's stuck now, so they aren't lost during the Try Its.\n\nThe device toolbar is the new one, and we'll use it all class. F12 opens DevTools. Ctrl + Shift + M turns on the device toolbar, where you can choose a phone from the list at the top or drag the page to any width. Press Ctrl + Shift + M again to turn it off.\n\nStudents last used it in Class 11, to check their Assignment 2 at phone width.");

{
  const s = newSlide();
  s.background = { color: C.white };
  s.addNotes("Before you read anything out, ask: what did we do last class? Give everyone 30 seconds to tell a neighbour, then take a few answers. Trying to remember it helps more than hearing it again.\n\nThen go through the four points.\n\nThe link to today: float, flexbox and grid arrange things on the page. Today's question is what happens to those arrangements when the screen gets small.\n\nThe drawing on the left shows both ideas: the text wrapping around a floated photo, and a box that holds its float because of display: flow-root.");
  txt(s, "Recall: What Did We Do Last Class?", { x: 0.75, y: 0.5, w: 11.8, h: 0.7, fontSize: 28, bold: true, color: C.black });
  // a floated photo with text wrapping around it, drawn with shapes
  const bx = 0.75, by = 1.6, bw = 5.2, bh = 4.1;
  card(s, { x: bx, y: by, w: bw, h: bh, border: C.teal, bw: 1.5 });
  s.addShape(pres.shapes.RECTANGLE, { x: bx + 0.3, y: by + 0.3, w: 2.2, h: 1.75, fill: { color: C.teal }, line: { color: C.teal } });
  txt(s, "float: left", { x: bx + 0.3, y: by + 0.3, w: 2.2, h: 1.75, fontFace: MONO, fontSize: 15, bold: true, color: C.white, align: "center", valign: "middle" });
  [0, 1, 2, 3, 4].forEach((i) => s.addShape(pres.shapes.RECTANGLE, { x: bx + 2.75, y: by + 0.38 + i * 0.36, w: 2.1 - (i % 2) * 0.3, h: 0.13, fill: { color: "BFC8CD" }, line: { color: "BFC8CD" } }));
  [0, 1, 2, 3].forEach((i) => s.addShape(pres.shapes.RECTANGLE, { x: bx + 0.3, y: by + 2.28 + i * 0.36, w: 4.55 - (i % 3) * 0.4, h: 0.13, fill: { color: "BFC8CD" }, line: { color: "BFC8CD" } }));
  txt(s, "The text flows around the photo, and the box still holds it.", { x: bx, y: by + bh + 0.12, w: bw, h: 0.4, fontSize: 13, italic: true, color: C.grey, align: "center" });
  const items = [
    ["Float", [["float: left", "c"], " or ", ["right", "c"], " moves an element to the side, and text flows around it."]],
    ["Width and margin", ["Give a float a width, and a margin on the side the text is on."]],
    ["Collapsing parents", ["A parent with only floats inside collapses. The fix: ", ["display: flow-root", "c"], "."]],
    ["The right tool", ["Flexbox and grid for layout. Float for wrapping text around something."]],
  ];
  const rx = 6.4, rw = W - rx - 0.75, rh = 1.1;
  items.forEach(([t, d], i) => {
    const y = 1.6 + i * (rh + 0.08);
    numCircle(s, i + 1, rx, y + 0.05, C.orange, 0.5);
    txt(s, t, { x: rx + 0.7, y, w: rw - 0.7, h: 0.38, fontSize: 18, bold: true, color: C.black });
    richBullets(s, [d], { x: rx + 0.7, y: y + 0.38, w: rw - 0.7, h: rh - 0.38, size: 15, bullet: false });
  });
}

overviewSlide("Week 9, Class 2: Overview", "Responsive Web Design", [
  "What responsive design is, and why it matters",
  "Media queries and breakpoints",
  "Fluid layouts, with percentages",
  "Responsive images, video and text",
  "Make a page responsive, then submit it",
], [["Why", "Every screen size"], ["How", "Queries, fluid units"], ["Build", "Your own page"]],
  "Rough plan for the two hours:\n- 10 min: environment check, recall, overview\n- 15 min: what and why, Examples 01 and 02, Pop Quiz 1\n- 25 min: the video, media queries, breakpoints, Examples 03 and 04, Pop Quizzes 2 and 3\n- 15 min: fluid layouts, Example 05, Pop Quiz 4\n- 20 min: images, video and text, Examples 06 and 07, Pop Quiz 5\n- 5 min: the full page, Example 08\n- 25 min: Try Its 1 to 3, then submit\n- 5 min: recap and the one-sentence summary\nThat's about two hours. If time is short, skip the old padding trick for video and leave clamp() as a demo.\n\nPop quizzes: give everyone 30 seconds to think, then count down and have them hold up 1 to 4 fingers (A to D) all at once, so nobody just copies a neighbour. Or use the Teams chat. Nobody is marked on these.");

dividerSlide("Responsive Web Design", "What it is   ·   why it matters   ·   what goes wrong without it",
  "Today's sites get viewed on everything from a small phone to a huge monitor. In the early days of phones, companies often built a second, separate site (m.example.com). Responsive design replaced that with one site that adapts.");

{
  const s = activitySlide("Two Words to Define (1 min)",
    "Ask before defining anything. One minute with a neighbour, then take a few answers.\n\nResponsive: in everyday English, something that reacts. Here, the page reacts to the space it's given.\n\nMedia query: a query is a question. 'Media' means the kind of device or screen. So a media query is CSS asking a question about the screen: are you 768 pixels wide or narrower?\n\nThe definitions are on the next slides, so don't spend long on it.");
  [["What does RESPONSIVE mean?", "On a web page, what is it responding to?"], ["What does MEDIA QUERY mean?", "A query is a question. What is the CSS asking?"]].forEach(([t, d], i) => {
    const x = 0.75 + i * 6.1;
    card(s, { x, y: 1.75, w: 5.75, h: 4.2 });
    txt(s, "?", { x, y: 2.0, w: 5.75, h: 1.4, fontSize: 80, bold: true, color: i ? C.purple : C.orange, align: "center", valign: "middle" });
    txt(s, t, { x: x + 0.35, y: 3.6, w: 5.05, h: 0.9, fontSize: 24, bold: true, color: C.black, align: "center", valign: "middle" });
    txt(s, d, { x: x + 0.35, y: 4.6, w: 5.05, h: 0.8, fontSize: 16, color: C.grey, align: "center", valign: "top" });
  });
}

{
  const s = panelSlide("What is Responsive Web Design?",
    "Read the definition, then the three ingredients. Each one gets its own section today.\n\n1. Fluid layouts: sizes in percentages and other flexible units instead of fixed pixels. Flexbox and grid from last week already help here.\n2. Flexible media: images and videos that shrink to fit.\n3. Media queries: different CSS for different screen sizes.\n\nThe goal isn't to make a site look identical on every device. It's to make it work well on each one.");
  txt(s, "An approach to web design that makes a site look good and work well on devices of all sizes, from desktops to phones.", { x: CX, y: 1.95, w: CW, h: 1.05, fontSize: 19, color: C.black, valign: "top" });
  [["Fluid layouts", "Adapt to the screen size"], ["Flexible media", "Images and video that fit"], ["Media queries", "Different CSS for different screens"]].forEach(([t, d], i) => {
    const w = 2.35, x = CX + i * (w + 0.25);
    card(s, { x, y: 3.15, w, h: 2.2 });
    numCircle(s, i + 1, x + 0.2, 3.32, [C.orange, C.teal, C.purple][i], 0.5);
    txt(s, t, { x: x + 0.2, y: 3.92, w: w - 0.35, h: 0.4, fontSize: 16, bold: true, color: C.black });
    txt(s, d, { x: x + 0.2, y: 4.32, w: w - 0.35, h: 0.85, fontSize: 14, color: C.grey, valign: "top" });
  });
  card(s, { x: CX, y: 5.6, w: CW, h: 1.05, fill: C.dark, border: C.dark });
  richBullets(s, [[["The goal: ", "b"], "the best experience on each device, not the same look on all of them."]], { x: CX + 0.3, y: 5.6, w: CW - 0.6, h: 1.05, size: 16, bullet: false, color: C.white, valign: "middle" });
}

{
  const s = panelSlide("Why Responsive Design Matters",
    "User experience: nobody wants to zoom in and scroll sideways to read a menu. Ask who has given up on a site like that on their phone.\n\nSearch: Google uses the mobile version of a site for indexing and ranking, so what's on the phone version is what counts.\n\nCost: one site means one set of files to update, instead of a separate desktop site and mobile site.\n\nAccessibility: someone who zooms in to 400% on a laptop is effectively looking at a phone-width page. If the layout reflows, they can read it without scrolling sideways. That's a WCAG requirement (Reflow, 1.4.10).");
  [["User experience", "No zooming in, no scrolling sideways, on any device."], ["Search", "Google looks at the mobile version of your site first."],
    ["Cost", "One site to update, not a desktop site and a mobile site."], ["Accessibility", "Text stays readable when people zoom in, because the page reflows."]].forEach(([t, d], i) => {
    const w = 3.65, h = 2.2, x = CX + (i % 2) * (w + 0.25), y = 1.95 + Math.floor(i / 2) * (h + 0.25);
    card(s, { x, y, w, h });
    txt(s, t, { x: x + 0.25, y: y + 0.25, w: w - 0.5, h: 0.45, fontSize: 19, bold: true, color: C.teal });
    txt(s, d, { x: x + 0.25, y: y + 0.8, w: w - 0.5, h: 1.2, fontSize: 16, color: C.dark, valign: "top" });
  });
}

codeSlide({
  kicker: "THE PROBLEM · EXAMPLE 01", title: "A Page That Doesn't Fit", file: "09_Responsive/01_fixed_layout_example.css",
  lines: between(cssOf("01_fixed_layout_example.html"), "/* 1.", "/* 4."), frames: [{ img: SH.ex01, label: "WIDE" }, { img: SH.ex01p, label: "PHONE, 360PX" }], maxFrameH: 3.3, size: 14,
  bullets: [
    ["Every size here is in ", ["pixels", "b"], ", so nothing can shrink."],
    ["On a phone you only see part of the page, and you have to ", ["scroll sideways", "b"], "."],
  ],
  notes: "Open 01_fixed_layout_example.html with Live Server. It looks fine on the classroom screen.\n\nNow drag the window narrower, or use the device toolbar (Ctrl + Shift + M) and pick a phone. The right side of the page disappears and a sideways scrollbar shows up.\n\nThree fixed things cause this: the 960px container, the 800px photo, and the three 300px columns. We'll fix each of them today. Students fix this very file in the Try Its at the end of class, so leave it broken.",
});

{
  const s = panelSlide("First: the Viewport Meta Tag",
    "This one line has been in the head of our example pages for a while. Today we find out what it does.\n\nOpen 02_viewport_example.html, turn on the device toolbar and pick a phone. Then comment out the meta tag, save and reload. Everything shrinks to tiny text.\n\nWhy: the first smartphones had to show sites built for desktops, so a phone pretends to be about 980 pixels wide and shrinks the page to fit. This tag says: this page is ready for phones, so use the device's real width.\n\nwidth=device-width: make the layout as wide as the screen. initial-scale=1.0: don't zoom in or out when the page loads.\n\nWithout it, media queries don't work as you'd expect either: a phone that thinks it's 980 pixels wide never matches max-width: 768px.\n\nThe left picture is a true 980-pixel layout, shown at phone size, to match what a phone does.\n\nNever add maximum-scale=1 or user-scalable=no: they stop people zooming in.");
  codeBox(s, ['<meta name="viewport"', '      content="width=device-width, initial-scale=1.0">'], { x: CX, y: 1.9, w: CW, h: 0.95, maxSize: 15, lang: "html" });
  const py = 3.3, pw = 2.0;
  phoneFrame(s, SH.ex02no, CX, py, pw, "WITHOUT IT", C.red);
  phoneFrame(s, SH.ex02yes, CX + pw + 0.3, py, pw, "WITH IT", C.green);
  const tx = CX + 2 * pw + 0.65, tw = CW - (2 * pw + 0.65);
  richBullets(s, [
    [["Without it: ", "b"], "the phone lays the page out about 980px wide, then shrinks it."],
    [["With it: ", "b"], "the page uses the phone's real width."],
    ["In every page you make. ", ["Example 02", "b"], "."],
  ], { x: tx, y: py - 0.1, w: tw, h: 3.6, size: 14, bullet: false, space: 12 });
}

quiz({
  q: "You forgot the viewport meta tag. What happens when someone opens your page on a phone?",
  options: ["The browser shows an error", "The page lays out about 980px wide and shrinks to fit, so the text is tiny", "The page looks the same, because the tag only affects search engines", "All of your CSS is ignored"], answer: 1,
  why: ["Phones do this so older desktop sites can still be seen in full.", "`width=device-width` tells the phone that the page is ready for its real width."],
  notes: "A quick check straight after the demo. Option C is the common guess, because it's in the head with the other 'meta' tags.",
  revealNotes: "Answer: B.\n\nNo error, and the CSS still works. Which is why it's easy to miss: the page looks fine on a laptop, and it's only on a real phone or in the device toolbar that you notice.",
});

dividerSlide("Media Queries", "Asking the screen a question   ·   breakpoints   ·   the order of your CSS",
  "A media query is CSS that only applies when a condition is true. This is the part of responsive design people usually think of first.");

videoSlide({ title: "Media Queries in 7 Minutes", label: "Learn CSS Media Query In 7 Minutes", source: "Web Dev Simplified on YouTube", url: "https://www.youtube.com/watch?v=yU7jJ3NbPdA",
  watchFor: ["@media", "max-width", "min-width"],
  notes: "Click the dark box (or the link) to open the video. It's the same one as in last year's slides.\n\nTell students not to take notes yet: the next few slides go through the same ideas, slowly, and they'll write their own in the Try Its." });

{
  const s = panelSlide("Reading a Media Query",
    "Read it out as a sentence: 'If the window is 768 pixels wide or narrower, make the body text 14 pixels.'\n\nInside the curly braces it's just normal CSS rules, as many as you like. That's why there are two levels of braces, and why people forget to close the outer one.\n\nA media query can also ask about height and orientation (portrait or landscape), or whether the page is being printed (@media print). Width is the one students will use nearly every time.\n\nThe width it checks is the browser window's, not the size of the device. That's why dragging the window on a laptop works for testing.");
  codeBox(s, ["@media (max-width: 768px) {", "   body {", "      font-size: 14px;", "   }", "}"], { x: CX, y: 1.9, w: CW, h: 2.05, maxSize: 17, lang: "css" });
  [["@media", "Starts a media query.", C.purple], ["(max-width: 768px)", ["The question: is the window ", ["768px or narrower", "b"], "?"], C.orange], ["{ ... }", "Normal CSS rules, used only when the answer is yes.", C.teal]].forEach(([p, d, c], i) => {
    const y = 4.2 + i * 0.68;
    pill(s, p, CX, y + 0.05, c, 2.3);
    richBullets(s, [typeof d === "string" ? [d] : d], { x: CX + 2.5, y, w: CW - 2.5, h: 0.45, size: 15, bullet: false, valign: "middle" });
  });
  txt(s, "It can also ask about height, orientation, or printing. Width is the one you'll use most.", { x: CX, y: 6.3, w: CW, h: 0.4, fontSize: 13, italic: true, color: C.grey });
}

codeSlide({
  kicker: "MEDIA QUERIES 1 · EXAMPLE 03", title: "A First Media Query", file: "09_Responsive/03_media_query_example.css",
  lines: between(cssOf("03_media_query_example.html"), "/* 1."), frames: [[{ img: SH.ex03, label: "800PX WIDE" }], [{ img: SH.ex03n, label: "560PX WIDE" }]], maxFrameH: 4.2, size: 14,
  bullets: [
    ["The normal rules apply at every width. The query's rules ", ["win", "b"], " when it's true."],
  ],
  notes: "Open 03_media_query_example.html and slowly drag the window narrower. At 768px the background and the text size change, all at once.\n\nThe label on the page changes too: that's the ::before pseudo-element from Class 9, with different content in each state.\n\nTry it live: change 768px to 500px and save. Now the change happens later. Then change the colour inside the media query, so students can see it's just normal CSS.\n\nThe device toolbar shows the exact width at the top of the page, which helps when you're hunting for a breakpoint.",
});

quiz({
  q: "At which window widths does this background turn pink?",
  code: ["@media (max-width: 768px) {", "   body {", "      background: pink;", "   }", "}"],
  options: ["Only at exactly 768px", "At 768px and narrower", "At 768px and wider", "At every width"], answer: 1,
  why: ["`max-width: 768px` means the window can be, at most, 768px.", "So the rule applies at 768px, and at everything narrower."],
  notes: "'Max' trips people up, because it sounds like big screens. Say it this way: the biggest window this applies to is 768px.",
  revealNotes: "Answer: B.\n\nAnd the opposite: min-width: 768px means at least 768px, so 768px and wider. Next, breakpoints use both.",
});

{
  const s = panelSlide("Breakpoints",
    "A breakpoint is a width where the layout changes.\n\nThese are the common values from last year's slides. Treat them as starting points: there are hundreds of screen sizes, and they keep changing. A better habit: drag the window, and add a breakpoint where your own layout starts to look wrong.\n\nTwo ways to write them:\n- max-width (desktop first): write the big-screen layout, then change things as the window gets smaller.\n- min-width (mobile first): write the phone layout, then add to it as the window gets bigger.\n\nBoth work. Phone CSS tends to be simpler, so mobile first usually means less code, and that's how Example 04 and Week 8's Example 17 are written. The page students fix today uses max-width, because they're starting from a page that's already a desktop layout.");
  txt(s, "A breakpoint is a screen width where your layout changes.", { x: CX, y: 1.9, w: CW, h: 0.45, fontSize: 18, color: C.black });
  const segs4 = [["Phone", "up to 480px", C.teal], ["Tablet", "481px to 768px", C.purple], ["Desktop", "769px and up", C.navy], ["Big screen", "1200px and up", C.dark]];
  const sw = CW / 4;
  segs4.forEach(([t, d, c], i) => {
    s.addShape(pres.shapes.RECTANGLE, { x: CX + i * sw, y: 2.55, w: sw, h: 1.15, fill: { color: c }, line: { color: C.light, width: 1.5 } });
    txt(s, t, { x: CX + i * sw, y: 2.65, w: sw, h: 0.45, fontSize: 16, bold: true, color: C.white, align: "center" });
    txt(s, d, { x: CX + i * sw, y: 3.1, w: sw, h: 0.4, fontSize: 13, color: C.white, align: "center" });
  });
  [["max-width", "Desktop first", "Start with the big layout. Change things as the window gets smaller."], ["min-width", "Phones first", "Start with the phone layout. Add to it as the window gets bigger."]].forEach(([c, t, d], i) => {
    const w = 3.65, x = CX + i * (w + 0.25);
    card(s, { x, y: 3.95, w, h: 1.85 });
    txt(s, c, { x: x + 0.22, y: 4.07, w: w - 0.44, h: 0.35, fontSize: 15, fontFace: MONO, bold: true, color: "8A2B0B" });
    txt(s, t, { x: x + 0.22, y: 4.42, w: w - 0.44, h: 0.35, fontSize: 16, bold: true, color: C.black });
    txt(s, d, { x: x + 0.22, y: 4.8, w: w - 0.44, h: 0.9, fontSize: 14, color: C.dark, valign: "top" });
  });
  richBullets(s, [[["Tip: ", "b"], "these are starting points. Add a breakpoint where ", ["your", "i"], " layout starts to look wrong."]], { x: CX, y: 6.0, w: CW, h: 0.65, size: 15, bullet: false });
}

codeSlide({
  kicker: "MEDIA QUERIES 2 · EXAMPLE 04", title: "Phones First", file: "09_Responsive/04_breakpoints_example.css",
  lines: between(cssOf("04_breakpoints_example.html"), "/* 1.", "/* 4."), frames: [[{ img: SH.ex04d, label: "1000PX" }], [{ img: SH.ex04p, label: "360PX" }, { img: SH.ex04t, label: "640PX" }]], maxFrameH: 3.75, size: 14,
  bullets: [
    ["The phone layout needs ", ["no media query", "b"], ". It's the starting point."],
    ["Each ", ["min-width", "c"], " query adds to the one before it."],
    ["The 1200px rule is in the file too: it stops the page growing."],
  ],
  notes: "Open 04_breakpoints_example.html and drag the window from narrow to wide: one column, two columns, four columns. The line under the heading says which rule is on.\n\nRead the CSS from the top down. No media query means it applies everywhere, so that's the phone layout. At 481px and up, the second block also applies and overrides the columns. At 769px and up, the third block also applies, and it overrides the second one because it comes later.\n\nPart 4 of the file (not on the slide) sets width: 1170px from 1200px up, with margin: 0 auto. That's the 'big screen' case: long lines of text and huge cards get hard to read.\n\nThese blocks only work in this order, smallest first. That's the next quiz.",
});

quiz({
  q: "You check this page on a phone. How big is the `h1`?",
  code: ["@media (max-width: 768px) {", "   h1 {", "      font-size: 20px;", "   }", "}", "", "h1 {", "   font-size: 40px;", "}"],
  options: ["20px, because the media query matches", "40px, because the last rule wins", "30px, halfway between the two", "It depends on the phone"], answer: 1,
  why: ["A media query adds no specificity, so both rules tie.", "When rules tie, the later one wins. Keep media queries after the rules they change."],
  notes: "This is the most common media query bug. The query is correct, and DevTools shows it matching, but the rule is crossed out.\n\nIt's the cascade at work.",
  revealNotes: "Answer: B.\n\nDemo it: in any example, move a media query above the rule it changes, and watch it stop working. In DevTools, the Styles panel shows the 20px rule crossed out.\n\nThe habit: put media queries at the bottom of the stylesheet.",
});

dividerSlide("Fluid Layouts", "Percentages, not pixels   ·   the padding trap",
  "Media queries change the layout at a few set widths. Fluid layouts take care of all the widths in between, so the page keeps fitting as the window changes.");

codeSlide({
  kicker: "FLUID 1 · EXAMPLE 05", title: "Percentages, Not Pixels", file: "09_Responsive/05_fluid_layout_example.css",
  lines: between(cssOf("05_fluid_layout_example.html"), "/* 1.", "/* 4."), result: SH.ex05, size: 14,
  bullets: [
    ["A ", ["percentage", "b"], " width resizes with the window, so there's no sideways scrolling."],
    [["margin: auto", "c"], " on the sides centres it."],
  ],
  notes: "Open 05_fluid_layout_example.html and drag the window narrower (the picture leaves out the padding trap at the bottom of the page, which is coming up next). The grey fixed box stays at 700px and starts sticking out. The fluid box keeps shrinking.\n\nA percentage is a percentage of the parent. Here the parent is the body, so 80% of the body's width.\n\nPractical tips from this one:\n- Use percentages for containers.\n- Keep images inside them with max-width: 100% (more on that soon).\n- Grid's fr units and flexbox from last week are fluid too.\n\nmax-width is a useful partner: width: 80% plus max-width: 1170px means 'fluid, but stop growing at some point'. Example 08 does that.",
});

quiz({
  q: "On a 360px phone, how wide is this `.container` box?",
  code: [".container {", "   width: 100%;", "   padding: 0 15px;", "}"],
  options: ["330px", "345px", "360px", "390px"], answer: 3,
  why: ["By default, `width` doesn't include padding. So it's 360 + 15 + 15 = 390px.", "That's 30px too wide, and now the page scrolls sideways."],
  revealImg: SH.ex05p, notes: "The box model, coming back at exactly the moment it matters. Most students will say 360px. Let them vote before you say anything.\n\nThis is the code from last year's slides, and it's in a lot of real sites. It's only a problem without box-sizing.",
  revealNotes: "Answer: D.\n\nThe screenshot is the bottom of Example 05 at 360px. The red box has width: 100% and padding, and it sticks out of the dashed box. The green one has box-sizing: border-box as well, and fits.\n\nThe fix is on the next slide.",
});

codeSlide({
  kicker: "FLUID 2 · EXAMPLE 05", title: "The Padding Trap", file: "09_Responsive/05_fluid_layout_example.css",
  lines: between(cssOf("05_fluid_layout_example.html"), "/* 4."), result: SH.ex05p, resultLabel: "A PHONE (360PX)", size: 14,
  bullets: [
    [["box-sizing: border-box", "c"], " keeps padding and borders ", ["inside", "b"], " the width."],
  ],
  notes: "This is the bottom of 05_fluid_layout_example.html at phone width. Same width, same padding. The only difference is one line.\n\nWith border-box, width: 100% means the whole box, padding included, is 100%. The browser takes the padding out of the content area.\n\nMany developers put this at the top of every stylesheet:\n* { box-sizing: border-box; }\nThat's fine for students to do too.",
});

dividerSlide("Images, Video and Text", "max-width: 100%   ·   aspect-ratio   ·   vw and clamp()",
  "The layout fits now. What's left is what goes inside it: images, video and text.");

codeSlide({
  kicker: "MEDIA 1 · EXAMPLE 06", title: "Responsive Images", file: "09_Responsive/06_responsive_media_example.css",
  lines: between(cssOf("06_responsive_media_example.html"), "/* 1.", "/* 2."), result: SH.ex06, size: 14,
  bullets: [
    [["max-width: 100%", "c"], ": never wider than the container."],
    [["height: auto", "c"], ": keep the shape, so it isn't squashed."],
    ["Keep the HTML's ", ["width", "c"], " and ", ["height", "c"], ", so the page doesn't jump as it loads."],
  ],
  notes: "Open 06_responsive_media_example.html and drag the window narrower. The photo shrinks with its column.\n\nWhy max-width and not width? width: 100% would stretch a small image to fill the container, and it would look blurry. max-width: 100% lets it shrink but never grow past its real size.\n\nWithout height: auto, the height from the HTML attribute (458) would stay put while the width shrinks, and the photo would look squashed. Try it live: delete height: auto.\n\nThe width and height attributes on the img tag tell the browser the shape before the file arrives, so it can save the right amount of space.",
});

codeSlide({
  kicker: "MEDIA 2 · EXAMPLE 06", title: "Responsive Video", file: "09_Responsive/06_responsive_media_example.css",
  lines: between(cssOf("06_responsive_media_example.html"), "/* 2."), result: SH.ex06, size: 14,
  bullets: [
    [["aspect-ratio", "c"], " does the job in one line. Use this one."],
    ["You'll still see the old ", ["padding-bottom", "c"], " trick in older code, and in code copied from the web."],
  ],
  notes: "Videos and iframes don't keep their shape the way images do, so they need help.\n\nToday's way: width: 100% and aspect-ratio: 16 / 9. The browser works out the height from the width. Every current browser supports it. It works on a YouTube iframe too.\n\nThe old way, from last year's slides, is worth recognizing:\n- padding-bottom: 56.25%: percentage padding is based on the parent's WIDTH, and 9 divided by 16 is 56.25%. So the padding makes a box that's always 16:9.\n- height: 0, so the padding is the only height there is.\n- The video is positioned absolutely, to fill that padding.\nIt's clever, and it isn't needed any more. For another shape, like 4:3, you'd use 3 / 4 = 75%.\n\nThe videos in this example have a poster image and no video file, so the repo stays small and nothing needs the internet. The comment in the HTML shows how to add a real one.",
});

codeSlide({
  kicker: "TEXT 1 · EXAMPLE 07", title: "Fluid Typography", file: "09_Responsive/07_fluid_type_example.css",
  lines: between(cssOf("07_fluid_type_example.html"), "/* 1.", "/* 2."), result: SH.ex07, resultLabel: "1000PX WIDE", size: 14,
  bullets: [
    [["1vw", "c"], " is ", ["1% of the window's width", "b"], ": at 1000px, ", ["5vw", "c"], " is 50px."],
    ["Text scales with the window, with no media query."],
    ["But it has ", ["no limits", "b"], ": tiny on a phone, huge on a big screen."],
  ],
  notes: "Open 07_fluid_type_example.html and look at the top line while you drag the window. It scales smoothly.\n\nThe arithmetic: 5vw on a 1000px window is 50px. On a 360px phone it's 18px. On a 2000px monitor it's 100px.\n\nA more serious problem: text sized only in vw doesn't get bigger when you zoom the browser (try Ctrl and +), because the window's width in CSS pixels shrinks as you zoom. People who need to zoom in can't. That's an accessibility failure, WCAG 1.4.4 (Resize Text).\n\nSo plain vw isn't enough. Next, two ways to control it.",
});

quiz({
  q: "You drag the window from 768px to 769px, just 1px wider. What happens to the `h1`?",
  code: ["@media (max-width: 768px) {", "   h1 {", "      font-size: 6vw;", "   }", "}", "", "@media (min-width: 769px) {", "   h1 {", "      font-size: 3vw;", "   }", "}"],
  options: ["It stays about the same", "It gets a tiny bit bigger", "It suddenly gets about half as big", "It suddenly gets about twice as big"], answer: 2,
  why: ["At 768px: 6vw is about 46px. At 769px: 3vw is about 23px.", "The window got bigger, and the heading got smaller! The next slide has a fix."],
  notes: "Last year's slides use this exact code to combine vw with media queries, and it does work to set a different size for phones and desktops. But work out the arithmetic at the breakpoint with the class.\n\nThey can use a calculator. 768 x 6% and 769 x 3%.",
  revealNotes: "Answer: C.\n\nDemo it in Example 07: watch the middle line while you drag slowly across 768px. It's a real jump. It also means the heading is smaller on a 1000px laptop (30px) than on a 768px tablet (46px).\n\nYou can fix that by picking better numbers, or use clamp().",
});

codeSlide({
  kicker: "TEXT 2 · EXAMPLE 07", title: "clamp(): Fluid, With Limits", file: "09_Responsive/07_fluid_type_example.css",
  lines: between(cssOf("07_fluid_type_example.html"), "/* 3."), frames: [[{ img: SH.ex07a, label: "768PX" }], [{ img: SH.ex07b, label: "769PX" }]], maxFrameH: 3.45, size: 14,
  bullets: [
    [["clamp(smallest, fluid, biggest)", "c"], " sets a floor and a ceiling."],
    ["Mix in ", ["rem", "c"], " so the text still grows when people zoom."],
    ["Look at the pictures: only the middle line jumps."],
  ],
  notes: "clamp takes three values: a minimum, the size you'd like, and a maximum.\n\nclamp(1.5rem, 1rem + 3vw, 3rem):\n- never smaller than 1.5rem (24px by default)\n- never bigger than 3rem (48px)\n- in between, 1rem + 3vw. You can do arithmetic inside clamp, without calc().\n\nAt 768px: 16 + 23 = 39px. At 360px: 16 + 11 = 27px. At 2000px: it would be 76px, but it stops at 48px.\n\nThe 1rem in the middle is what helps zooming, because rem grows with the browser's zoom and text size settings.\n\nThe pictures are Example 07 at 768px and 769px, one pixel apart. The top and bottom lines barely change, and the middle one halves.\n\nStudents don't need to memorize the numbers: 1rem + 3vw is a fine default to copy for headings. Body text can stay in rem with no vw at all.",
});

dividerSlide("Putting It Together", "Example 01, made responsive   ·   then your turn",
  "Everything from today, on one page. It's the same page as Example 01, with the same HTML, and it's what students are about to build.");

codeSlide({
  kicker: "PRACTICAL EXAMPLE · EXAMPLE 08", title: "Same HTML as Example 01", file: "09_Responsive/08_responsive_page_example.html",
  lines: between(bodyOf("08_responsive_page_example.html"), "<div", '<section class="cards">').concat(["   ...", "</div>"]), lang: "html", frames: [{ img: SH.ex08, label: "WIDE" }, { img: SH.ex08p, label: "PHONE, 360PX" }], maxFrameH: 3.3, size: 14,
  bullets: [
    ["A simple page in a ", ["container", "c"], ". Only the CSS changed."],
    ["Good HTML doesn't need to change for a phone."],
  ],
  notes: "Open 08_responsive_page_example.html next to 01_fixed_layout_example.html, and drag both narrower.\n\nThe point of this slide: the HTML files are the same, apart from the title, the stylesheet link and one comment. Responsive design is nearly all CSS, as long as the HTML is clean and in a sensible order. That's a good reminder of why semantic HTML, in reading order, was worth doing.",
});

codeSlide({
  kicker: "PRACTICAL EXAMPLE · EXAMPLE 08", title: "All the CSS It Took", file: "09_Responsive/08_responsive_page_example.css",
  lines: between(CSS08, "/* 1."), result: SH.ex08p, resultLabel: "A PHONE (360PX)", size: 14,
  bullets: [
    ["Everything from today: fluid width, a fitted photo, ", ["clamp()", "c"], ", and one media query, ", ["last", "b"], "."],
  ],
  notes: "Walk through the numbered comments, and link each one back to its section:\n1. Fluid container, with a max-width and box-sizing. 80% on a big screen, 100% on a small one.\n2. The photo.\n3. Three columns with fr instead of pixels.\n4. clamp() for the heading.\n5. One media query at the very end.\n\nThe media query includes width: 100% with padding, which is the padding trap, and box-sizing: border-box in part 1 is what makes it safe.\n\nThat's about 30 lines to fix the whole page. Don't dwell here: students are about to type it.",
});

trySlide({
  n: 1, title: "A Fluid Page", time: "8 min", codeTitle: "REPLACE PARTS 1 TO 3 IN THE CSS FILE",
  steps: [
    ["Open Example 01", ["Open the ", ["html", "c"], " with Go Live, and the ", ["css", "c"], " beside it. Using your own page is fine too."]],
    ["Be a phone", ["F12, then ", ["Ctrl + Shift + M", "b"], ". Scroll sideways: that's the problem."]],
    ["Change the CSS", "Replace parts 1, 2 and 3 with the code on the right. Save."],
    ["Check it", "The photo and the container now fit, but the three columns get squished. That's next."],
  ],
  code: between(CSS08, "/* 1.", "/* 4."),
  notes: "Students work directly in 01_fixed_layout_example.css. The finished version is Example 08, so anyone who's stuck can look there.\n\nThe old activity was 'create a simple web page', so anyone who'd rather use their own page can do that. They'll need a fluid layout, a responsive image and a media query.\n\nExpect this: after Try It 1, at phone width, there are still three squished columns. That's right, and the media query in Try It 3 fixes it.",
});

trySlide({
  n: 2, title: "A Fluid Heading", time: "3 min", codeTitle: "REPLACE PART 4 IN THE CSS FILE",
  steps: [
    ["Change the CSS", "Replace part 4 with the code on the right. Save."],
    ["Drag the window", "The heading grows and shrinks smoothly."],
    ["Play with it", ["Change ", ["3vw", "c"], " to ", ["5vw", "c"], ". Or change the limits. When does it stop growing?"]],
  ],
  code: between(CSS08, "/* 4.", "/* 5."),
  notes: "A quick one. It's fine to copy the formula.\n\nTo see the limits: drag the window to full width (it stops at 3rem) and then to a phone's width (it stops at 1.5rem).",
});

trySlide({
  n: 3, title: "A Media Query", time: "8 min", codeTitle: "ADD TO THE VERY END OF THE CSS FILE",
  steps: [
    ["Type the CSS on the right", ["At the ", ["very end", "b"], " of the file. Save."]],
    ["Drag past 768px", "One column at 768px and narrower. Three columns above it."],
    ["Check for sideways scrolling", "At 360px there should be none. The page should still have side padding."],
    ["Make it yours", "Add one more change to the query: a colour, a font size, or a smaller gap."],
  ],
  code: between(CSS08, "/* 5."),
  tip: [["Not working? Check the two levels of ", ["}", "c"], ", and that the query comes ", ["after", "b"], " the rules it changes."]],
  notes: "The two usual problems:\n- A missing closing brace. Everything after it stops working. The CSS validator will catch it.\n- The query isn't last. That's Pop Quiz 3.\n\nIf the page scrolls sideways at 360px, check that box-sizing: border-box from Try It 1 is on .container.\n\nFast finishers: add a second breakpoint, with two columns between 481px and 768px. Or go back to Harbourside and do Try It 8 ('On a phone') from Grid_and_Float.pptx, in 08_Flex_Grid_CSS/activity.",
});

{
  const s = activitySlide("Class Activity: Submit Your Page",
    "This is the activity from last year's slides: a simple page with a fluid layout, responsive images and media queries.\n\nStudents take the screenshots with the device toolbar, so the phone width shows. Its menu (the three dots on the toolbar) also has 'Capture screenshot'.\n\nCheck that the Brightspace box is set up before class.");
  const LW = 6.05;
  card(s, { x: 0.75, y: 1.45, w: LW, h: 5.5 });
  txt(s, "Your page should have:", { x: 1.05, y: 1.65, w: LW - 0.6, h: 0.4, fontSize: 18, bold: true, color: C.black });
  ["A fluid layout", "Responsive images", "Media queries for different screen sizes"].forEach((t, i) => {
    const y = 2.2 + i * 0.6;
    numCircle(s, "✓", 1.05, y, C.teal, 0.42);
    txt(s, t, { x: 1.65, y, w: LW - 1.2, h: 0.42, fontSize: 16, color: C.black, valign: "middle" });
  });
  txt(s, "Then:", { x: 1.05, y: 4.1, w: LW - 0.6, h: 0.4, fontSize: 18, bold: true, color: C.black });
  [[["Take two screenshots: wide, and at phone width (", ["Ctrl + Shift + M", "b"], ")."]], [["Drop them in the box on Week 9 Brightspace."]]].forEach(([d], i) => {
    const y = 4.6 + i * 0.95;
    numCircle(s, i + 1, 1.05, y, C.orange, 0.42);
    richBullets(s, [d], { x: 1.65, y: y - 0.03, w: LW - 1.2, h: 0.8, size: 16, bullet: false, color: C.black });
  });
  const rx = 0.75 + LW + 0.35, ih = 3.0;
  framedImage(s, SH.ex08, rx, 1.8, ih / SH.ex08.ratio > 3.95 ? 3.95 : ih / SH.ex08.ratio, "WIDE");
  const ph = 2.7, pw = ph / SH.ex08p.ratio;
  framedImage(s, SH.ex08p, rx + 4.1, 1.8, Math.min(1.65, pw), "PHONE");
  richBullets(s, [[["Stuck? ", "b"], ["08_responsive_page_example", "c"], " is the finished page."]], { x: rx, y: 5.5, w: W - rx - 0.6, h: 0.8, size: 15, bullet: false });
}

{
  const s = panelSlide("Class Recap",
    "Go through these quickly. If students remember three things: the viewport tag goes on every page, media queries go last, and use max-width: 100% on images.\n\nThen the last slide: the one-sentence summary.");
  labelBullets(s, [
    { label: "What we covered:", text: [
      "Responsive design: one site that works well on any screen.",
      "The viewport meta tag goes on every page.",
      "Media queries change the CSS at breakpoints. Keep them last.",
      "Fluid layouts use percentages, with box-sizing: border-box.",
      "img { max-width: 100%; height: auto; } and aspect-ratio for video.",
      "clamp() gives fluid text, with limits.",
    ] },
  ], { size: 20, sub: 16 });
}

dividerSlide("Write a one-sentence summary of what we did in today's class.", "One sentence, in your own words.",
  "Give everyone a minute. They can write it on paper or post it in the Teams chat. Ask two or three people to read theirs out.\n\nIf the sentences mention media queries but nothing else, that's normal. Bring up fluid layouts next time.");

pres.writeFile({ fileName: OUT }).then(() => console.log("wrote " + OUT));
