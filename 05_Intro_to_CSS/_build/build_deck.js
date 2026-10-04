// Builds Class8_Intro_to_CSS.pptx in the same style as the Class 5, 6 and 7 decks.
// To rebuild: cd 05_Intro_to_CSS/_build, run "npm install" once, then "npm run build".
// The "Result" pictures are screenshots of the example pages, taken with Microsoft Edge (headless) each build.
// Code slides read the example files directly, so edit those files and re-run to update the slides.
const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");
const { execFileSync } = require("child_process");

const REPO = path.join(__dirname, ".."); // the 05_Intro_to_CSS folder
const SCR = __dirname;
const SHOTS = path.join(SCR, "shots");
const OUT = path.join(REPO, "Class8_Intro_to_CSS.pptx");

const FONT = "Aptos";
const MONO = "Consolas";
const UNIT = "HTML/CSS Techniques";
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
function shot(file, name, w = 600, h = 400, scale = 2) {
  const out = path.join(SHOTS, name + ".png");
  const url = "file:///" + path.join(REPO, file).replace(/\\/g, "/");
  // Its own profile folder, so it never attaches to an Edge window you already have open
  const args = ["--headless=new", "--disable-gpu", "--hide-scrollbars", "--no-first-run", `--user-data-dir=${PROFILE}`,
    `--window-size=${w},${h}`, `--force-device-scale-factor=${scale}`, `--screenshot=${out}`, url];
  for (let tries = 0; ; tries++) {
    try { execFileSync(EDGE, args, { stdio: "ignore", timeout: 30000 }); break; }
    catch (e) { if (tries >= 2) throw e; }
  }
  return { path: out, ratio: h / w };
}

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "Class 8: Intro to CSS";
pres.author = "Jamie Symonds";

const W = 13.333, H = 7.5;
const PANEL_W = 4.1;
const CX = 5.2, CW = 7.55; // content column on panel slides

// ---------- helpers (same as the Class 5–7 decks) ----------
function txt(slide, text, opts) {
  slide.addText(text, Object.assign({ isTextBox: true, fontFace: FONT, margin: 0 }, opts));
}

function panelSlide(title, notes) {
  const s = pres.addSlide();
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

// ---- syntax highlighting (VS Code Dark+ colours) ----
const HL = { text: "D4D4D4", punct: "808080", tag: "569CD6", attr: "9CDCFE", val: "CE9178", comment: "6A9955", ent: "D7BA7D",
  sel: "D7BA7D", prop: "9CDCFE", num: "B5CEA8", brace: "D4D4D4" };

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
  let inComment = false, depth = 0, inValue = false;
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
      if (ch === "{") { push(ch, HL.brace); depth++; inValue = false; i++; continue; }
      if (ch === "}") { push(ch, HL.brace); depth = Math.max(0, depth - 1); inValue = false; i++; continue; }
      if (depth === 0) {
        const m = line.slice(i).match(/^[^{/]+/);
        if (m) { push(m[0], /^\s+$/.test(m[0]) ? HL.text : HL.sel); i += m[0].length; } else { push(ch, HL.sel); i++; }
        continue;
      }
      if (!inValue) {
        if (ch === ":") { push(":", HL.text); inValue = true; i++; continue; }
        const m = line.slice(i).match(/^[a-zA-Z-]+/);
        if (m) { push(m[0], HL.prop); i += m[0].length; continue; }
        push(ch, HL.text); i++; continue;
      }
      if (ch === ";") { push(";", HL.text); inValue = false; i++; continue; }
      const num = line.slice(i).match(/^-?\d+(\.\d+)?(px|rem|em|%|deg|s)?/);
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
// The CSS inside a page's <style> block
function styleOf(file) {
  const L = readLines(file);
  const a = L.findIndex((l) => l.trim() === "<style>"), b = L.findIndex((l) => l.trim() === "</style>");
  return trim(dedent(L.slice(a + 1, b), 3)).map(tabs);
}
function bodyOf(file) {
  const L = readLines(file);
  const a = L.findIndex((l) => l.trim() === "<body>"), b = L.findIndex((l) => l.trim() === "</body>");
  return trim(dedent(L.slice(a + 1, b), 2)).map(tabs);
}

function codeBox(s, lines, { x, y, w, h, maxSize = 13, fitHeight = false, lang = "html", title }) {
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
function codeSlide({ kicker, title, bullets, file, lines, lang = "css", result, notes, size = 14, codeTitle }) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  s.addImage({ path: path.join(SCR, "grad_left.png"), x: 0, y: 0, w: CODE_PANEL, h: H });
  codeBox(s, lines, { x: 0.3, y: 0.4, w: CODE_PANEL - 0.6, h: H - 0.8, maxSize: 15, fitHeight: true, lang, title: codeTitle });
  const RX = CODE_PANEL + 0.45, RW = W - RX - 0.45;
  txt(s, kicker, { x: RX, y: 0.45, w: RW, h: 0.35, fontSize: 14, color: C.grey, bold: true, charSpacing: 1 });
  const long = title.length > 26;
  txt(s, title, { x: RX, y: 0.8, w: RW, h: long ? 1.05 : 0.7, fontSize: long ? 26 : 32, color: C.black, valign: "top" });
  const by = long ? 1.95 : 1.65;
  let bh = 6.85 - by;
  if (result) {
    // as wide as the column, but never taller than 3.05", so the bullets keep their room
    const ih = Math.min(RW * result.ratio, 3.05), iw = ih / result.ratio;
    const iy = 6.85 - ih - 0.05;
    framedImage(s, result, RX, iy, iw, "RESULT");
    bh = iy - 0.32 - by - 0.1;
  }
  richBullets(s, bullets, { x: RX, y: by, w: RW, h: bh, size });
  if (file) txt(s, file, { x: RX, y: 6.95, w: RW, h: 0.3, fontSize: 11, color: C.mid, fontFace: MONO });
  if (notes) s.addNotes(notes);
  return s;
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

// ---------- screenshots ----------
const SH = {
  before: shot("activity/index.html", "before", 1100, 800, 1),
  after: shot("activity/answer/index.html", "after", 1100, 800, 1),
  ex01: shot("01_three_ways_example.html", "ex01", 600, 290),
  ex02: shot("02_selectors_example.html", "ex02", 600, 500),
  ex03: shot("03_colour_example.html", "ex03", 600, 260),
  ex04: shot("04_text_fonts_example.html", "ex04", 600, 470),
  ex05: shot("05_box_model_example.html", "ex05", 600, 520),
  ex06: shot("06_styling_images_example.html", "ex06", 600, 600),
  ex07: shot("07_cascade_example.html", "ex07", 600, 330),
};

// =====================================================================
// 1. Title
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  txt(s, "WEBD 1000", { x: 0.3, y: 0.3, w: 3, h: 0.4, fontSize: 16, color: C.black });
  s.addShape(pres.shapes.OVAL, { x: 0.55, y: 1.1, w: 5.6, h: 5.6, fill: { color: C.orange }, line: { color: C.orange } });
  txt(s, UNIT, { x: 1.45, y: 2.4, w: 4.2, h: 2.9, fontSize: 46, color: C.white, valign: "middle" });
  s.addShape(pres.shapes.OVAL, { x: 1.0, y: 5.75, w: 0.62, h: 0.62, fill: { color: C.purple }, line: { color: C.purple } });
  s.addShape(pres.shapes.ARC, { x: 9.6, y: 0.75, w: 3.3, h: 3.3, angleRange: [200, 20], line: { color: C.blue, width: 7, dashType: "dash" } });
  txt(s, [
    { text: "Class 8:", options: { bullet: true, fontSize: 30, breakLine: true } },
    { text: "Intro to CSS: making our HTML look good.", options: { bullet: { code: "25CB" }, indentLevel: 1, fontSize: 26, paraSpaceBefore: 6 } },
  ], { x: 6.85, y: 1.9, w: 5.6, h: 2.2, color: C.black, valign: "top" });
  s.addNotes("Welcome back. So far: user stories, semantic HTML, a wireframe for Harbourside Fish & Chips, and last class, optimized images. Today we finally make it look like a real website, with CSS.\n\nFormat for today: a short concept slide, then you try it straight away on the Harbourside page. Seven rounds of that. By the end of class everyone has a styled page.");
}

// 2. Attendance
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  txt(s, "Attendance", { x: 1.1, y: 0.75, w: 8, h: 0.9, fontSize: 44, color: C.black });
  s.addNotes("Take attendance. While you do, have students git pull and open the 05_Intro_to_CSS folder in VS Code.");
}

// 3. Overview
{
  const s = panelSlide("Class 8: Class Overview",
    "Each round is about 5 minutes of concept and 8–12 minutes of hands-on time. Rough plan: 15 min intro, then seven rounds (~85 min), and 'Make It Yours' if there's time (otherwise it's homework).\n\nEverything students write goes into one file, activity/style.css, under a heading for each activity. If anyone falls behind, they can copy that section from activity/answer/style.css and keep going.");
  txt(s, [
    { text: "Title: ", options: { bold: true, bullet: true, fontSize: 22 } },
    { text: "Intro to CSS", options: { fontSize: 22, breakLine: true } },
    { text: " ", options: { fontSize: 10, breakLine: true } },
    { text: "Agenda:", options: { bold: true, bullet: true, fontSize: 22, breakLine: true } },
    { text: "What CSS is, and how a rule is written", options: { bullet: true, indentLevel: 1, fontSize: 18, paraSpaceBefore: 4, breakLine: true } },
    { text: "Seven rounds of concept → try it:", options: { bullet: true, indentLevel: 1, fontSize: 18, paraSpaceBefore: 4, breakLine: true } },
    { text: "linking CSS · selectors · colour · text · the box model · images · the cascade and DevTools", options: { bullet: true, indentLevel: 2, fontSize: 15, paraSpaceBefore: 2, breakLine: true, color: C.grey } },
    { text: "Make it yours, and submit", options: { bullet: true, indentLevel: 1, fontSize: 18, paraSpaceBefore: 4 } },
  ], { x: CX, y: 1.95, w: CW, h: 3.6, color: C.black, valign: "top" });
  const steps = [["Concept", "One idea, one example"], ["Try it", "Add it to the Harbourside page"], ["Repeat", "Seven times"]];
  steps.forEach(([t, d], i) => {
    const x = CX + i * 2.6;
    card(s, { x, y: 5.6, w: 2.2, h: 1.15, fill: i === 1 ? C.orange : C.white, border: i === 1 ? C.orange : "D0D7DB" });
    txt(s, t, { x: x + 0.2, y: 5.7, w: 1.8, h: 0.45, fontSize: 17, bold: true, color: i === 1 ? C.white : C.black });
    txt(s, d, { x: x + 0.2, y: 6.15, w: 1.8, h: 0.5, fontSize: 12.5, color: i === 1 ? C.white : C.grey, valign: "top" });
    if (i < 2) txt(s, "→", { x: x + 2.2, y: 5.6, w: 0.4, h: 1.15, fontSize: 22, bold: true, color: C.orange, align: "center", valign: "middle" });
  });
}

// 4. Get the code
{
  const s = panelSlide("Before We Start: Get the Code",
    "The activity folder is the Harbourside Fish & Chips page from the wireframe class, written in semantic HTML, with no CSS yet.\n\nAsk students not to open activity/answer until they're stuck. It has the finished style.css with a section for each activity.\n\nThey'll also want their my_images folder from last class for Try It 6.");
  txt(s, [{ text: "1.  Pull this week's examples", options: { bold: true, fontSize: 18 } }], { x: CX, y: 1.95, w: CW, h: 0.4 });
  codeBox(s, ["git pull"], { x: CX, y: 2.45, w: 5.2, h: 0.7, maxSize: 16 });
  txt(s, "New folder: 05_Intro_to_CSS (examples 01–07, and the activity folder).", { x: CX, y: 3.25, w: CW, h: 0.5, fontSize: 15, color: C.dark });
  txt(s, [{ text: "2.  Open the activity with Live Server", options: { bold: true, fontSize: 18 } }], { x: CX, y: 3.95, w: CW, h: 0.4 });
  txt(s, "Right-click activity/index.html → Open with Live Server. Put VS Code and the browser side by side (Win + ← and Win + →).", { x: CX, y: 4.4, w: CW, h: 0.8, fontSize: 15, color: C.dark });
  txt(s, [{ text: "3.  Have last class's photo handy", options: { bold: true, fontSize: 18 } }], { x: CX, y: 5.4, w: CW, h: 0.4 });
  txt(s, "04_Images_for_the_Web/my_images/photo_1200.jpg. We'll put it on the page in Try It 6.", { x: CX, y: 5.85, w: CW, h: 0.5, fontSize: 15, color: C.dark });
}

// 5. What is CSS?
{
  const s = panelSlide("What Is CSS?",
    "CSS = Cascading Style Sheets. HTML says what each thing IS (a heading, a list, a navigation area). CSS says how it LOOKS: colours, fonts, spacing, layout.\n\nThese two screenshots are the SAME index.html file. The only difference is one linked stylesheet. That's where we're going today.\n\nFamous demo: csszengarden.com. Hundreds of designs, all using exactly the same HTML file. Click through a few designs if there's time.\n\nWhy keep them separate? Change one .css file and every page on the site updates. And the HTML stays clean and accessible, which is what screen readers and search engines read.");
  const iw = 3.6;
  const h1 = framedImage(s, SH.before, CX, 2.25, iw, "HTML ONLY");
  framedImage(s, SH.after, CX + iw + 0.35, 2.25, iw, "HTML + CSS");
  richBullets(s, [
    [["CSS", "b"], " = Cascading Style Sheets. ", ["HTML", "b"], " says what things ", ["are", "i"], ". ", ["CSS", "b"], " says how they ", ["look", "i"], "."],
    ["Same index.html in both pictures. The only difference is ", ["one linked .css file", "b"], "."],
    ["One stylesheet can style a whole site. Change it once, every page updates."],
  ], { x: CX, y: 2.25 + h1 + 0.25, w: CW, h: 6.95 - (2.25 + h1 + 0.25), size: 15 });
}

// 6. Anatomy of a rule
{
  const s = panelSlide("Anatomy of a CSS Rule",
    "Every CSS rule has the same shape. Read it out loud: 'For every h1: make the colour navy, and the font size 2rem.'\n\nSelector = WHO. Declarations = WHAT. Each declaration is property: value; and you can have as many as you like inside the braces.\n\nThe colours on the right match the colours VS Code uses, so students can spot each part in their own editor.");
  codeBox(s, ["h1 {", "   color: navy;", "   font-size: 2rem;", "}"], { x: CX, y: 1.95, w: 3.6, h: 2.3, maxSize: 28, lang: "css" });
  const parts = [
    ["Selector", HL.sel, "Which elements to style", "h1"],
    ["Property", HL.prop, "What to change", "color"],
    ["Value", HL.val, "What to change it to", "navy"],
    ["Declaration", "D4D4D4", "One instruction, ending in ;", "color: navy;"],
    ["Braces", "D4D4D4", "Wrap all the declarations", "{ }"],
  ];
  parts.forEach(([t, col, d, ex], i) => {
    const y = 1.95 + i * 0.95, x = CX + 3.9, w = CW - 3.9;
    card(s, { x, y, w, h: 0.8, fill: "1E1E1E", border: "1E1E1E" });
    txt(s, ex, { x: x + 0.2, y, w: 1.35, h: 0.8, fontSize: 13, bold: true, fontFace: MONO, color: col, valign: "middle" });
    txt(s, [{ text: t, options: { bold: true, color: C.white, fontSize: 14.5, breakLine: true } }, { text: d, options: { color: "BFBFBF", fontSize: 12 } }],
      { x: x + 1.55, y, w: w - 1.7, h: 0.8, valign: "middle" });
  });
  card(s, { x: CX, y: 4.5, w: 3.6, h: 2.25, fill: C.white });
  txt(s, "Read it out loud", { x: CX + 0.25, y: 4.65, w: 3.1, h: 0.4, fontSize: 15, bold: true, color: C.black });
  txt(s, "\"For every h1: make the colour navy, and the font size 2rem.\"", { x: CX + 0.25, y: 5.1, w: 3.1, h: 1.5, fontSize: 15, italic: true, color: C.dark, valign: "top" });
}

// 7. Writing CSS: gotchas
{
  const s = panelSlide("Writing CSS: Watch Out For…",
    "These are the mistakes beginners make most. The big problem: CSS fails SILENTLY. A typo doesn't give an error message. The browser skips that line and carries on, so the page just doesn't change.\n\nThe Canadian one: it's color, not colour. CSS only understands American spelling. Same with 'center', not 'centre'.\n\nDevTools helps: an invalid declaration shows with a yellow warning triangle and a line through it.");
  codeBox(s, ["/* Broken: nothing happens */", "h1 {", "   colour: navy", "   font-size: 2 rem;", "   text-align = centre;", "}"],
    { x: CX, y: 1.95, w: 3.6, h: 2.5, maxSize: 15, lang: "css" });
  codeBox(s, ["/* Fixed */", "h1 {", "   color: navy;", "   font-size: 2rem;", "   text-align: center;", "}"],
    { x: CX + 3.95, y: 1.95, w: 3.6, h: 2.5, maxSize: 15, lang: "css" });
  richBullets(s, [
    [["American spelling: ", "b"], ["color", "c"], " and ", ["center", "c"], ", not colour and centre."],
    [["Every declaration ends in ", "b"], [";", "c"], ". Missing one breaks the next line too."],
    [["Colon, not equals: ", "b"], ["color: navy", "c"], ". (HTML attributes use ", ["=", "c"], ".)"],
    [["No space before the unit: ", "b"], ["2rem", "c"], ", not ", ["2 rem", "c"], "."],
    [["Comments are ", "b"], ["/* … */", "c"], " in CSS, not ", ["<!-- -->", "c"], "."],
  ], { x: CX, y: 4.7, w: CW, h: 2.3, size: 14.5, space: 5 });
}

// 8. Divider
dividerSlide("Concept → Try It", "Seven rounds   ·   one restaurant website   ·   all in activity/style.css",
  "From here on it's a rhythm: a concept slide, then a Try It slide. Students add each piece to activity/style.css under the matching heading.\n\nKeep the browser and VS Code side by side. Live Server reloads every time they save.");

// 9. Three ways to add CSS
{
  const s = panelSlide("Concept 1: Three Ways to Add CSS",
    "Inline styles are handy for a quick test but a nightmare to maintain: to change the colour of every price, you'd have to edit every element. They also beat almost every other rule, which makes debugging hard.\n\nInternal (<style> in the head) is fine for a single page or a demo. That's what our numbered examples use, so everything is in one file.\n\nExternal is what real sites use: one style.css linked from every page. The browser also downloads it once and caches it for the other pages.");
  const ways = [
    ["Inline", "style=\"…\" on one element", "One element only. Hard to maintain. Avoid.", C.red, "<p style=\"color: red;\">"],
    ["Internal", "<style> in the <head>", "One page only. Fine for demos.", C.orange, "<style> p { … } </style>"],
    ["External", "<link> to a .css file", "Every page shares one file. Use this.", C.green, "<link rel=\"stylesheet\" href=\"style.css\">"],
  ];
  ways.forEach(([t, how, d, col, ex], i) => {
    const y = 1.95 + i * 1.62;
    card(s, { x: CX, y, w: CW, h: 1.42 });
    numCircle(s, i + 1, CX + 0.2, y + 0.2, col, 0.5);
    txt(s, t, { x: CX + 0.9, y: y + 0.2, w: 2.0, h: 0.5, fontSize: 18, bold: true, color: C.black, valign: "middle" });
    txt(s, how, { x: CX + 2.8, y: y + 0.2, w: CW - 3.0, h: 0.5, fontSize: 14, color: C.grey, valign: "middle" });
    txt(s, d, { x: CX + 0.9, y: y + 0.72, w: 3.0, h: 0.6, fontSize: 13.5, color: C.dark, valign: "top" });
    codeBox(s, [ex], { x: CX + 3.95, y: y + 0.72, w: CW - 4.15, h: 0.5, maxSize: 12, lang: "html" });
  });
}

// 10. Example 01
codeSlide({
  kicker: "EXAMPLE 01", title: "All Three on One Page", file: "01_three_ways_example.html + 01_external.css", lang: "mixed", result: SH.ex01,
  lines: [...trim(readLines("01_three_ways_example.html").slice(7, 16)).map((l) => tabs(l.replace(/^\t\t/, ""))), "", ...bodyOf("01_three_ways_example.html")],
  bullets: [
    [["<link>", "c"], " goes in the ", ["<head>", "c"], ". ", ["href", "c"], " is the path to the .css file, just like an ", ["<a href>", "c"], "."],
    ["Open ", ["01_external.css", "c"], ": no ", ["<style>", "c"], " tags, just rules."],
    ["The inline style only affects that one paragraph."],
  ],
  notes: "Open 01_three_ways_example.html with Live Server, then open 01_external.css next to it.\n\nDemo: change navy to darkred in 01_external.css and save. The h1 changes. If ten pages linked this file, all ten would change. That's the reason external stylesheets win.\n\nCommon bug for Try It 1: the href path is wrong (e.g. href=\"css/style.css\" when the file isn't in a css folder). The browser fails silently. DevTools → Console shows a 404 for the missing file.",
});

// 11. Try it 1
trySlide({
  n: 1, title: "Link Your Stylesheet", time: "5 min",
  steps: [
    ["Open activity/index.html with Live Server", "Black text on white: the browser's default styles."],
    ["Open activity/style.css", "Only comments so far. It isn't linked, so the browser ignores it."],
    ["Add the <link> in index.html", ["Inside the ", ["<head>", "c"], ", where the comment says ", ["Activity 1", "b"], "."]],
    ["Add the body rule under Activity 1", "Save both files. The background should turn a very light blue."],
    ["Nothing happened?", "Check: both files saved? href spelled exactly style.css? Both files in the activity folder?"],
  ],
  code: ["<!-- in index.html, inside <head> -->", "<link rel=\"stylesheet\" href=\"style.css\">", "", "/* in style.css */", "body {", "   background-color: #F4F8FA;", "}"],
  lang: "mixed", codeTitle: "TYPE THIS",
  tip: [["The colour is subtle on purpose. If you can't tell it worked, try ", ["hotpink", "c"], " first, then change it back."]],
  notes: "Walk around and make sure EVERYONE has a working link before moving on. Everything else today depends on it.\n\nThe most common problems: forgot to save index.html; typed href=\"styles.css\" (with an s); put the <link> inside <body>; or opened index.html directly from File Explorer instead of with Live Server (it still works, it just doesn't auto-reload).",
});

// 12. Selectors (Example 02)
codeSlide({
  kicker: "CONCEPT 2", title: "Selectors: Choosing What to Style", file: "02_selectors_example.html", lines: styleOf("02_selectors_example.html"), result: SH.ex02, size: 13.5,
  bullets: [
    [["p", "c"], " element: every one. ", [".price", "c"], " class: every element with ", ["class=\"price\"", "c"], ". ", ["#specials", "c"], " id: one element only."],
    [["h1, h2", "c"], " comma = both. ", ["nav a", "c"], " space = an ", ["a", "c"], " somewhere inside a ", ["nav", "c"], "."],
  ],
  notes: "Walk through each selector and find what it styles in the result.\n\nClass vs id: a class can be used on as many elements as you like, and one element can have several classes (class=\"menu-item special\"). An id must be unique on the page. We've already used ids for links (href=\"#menu\") and aria-labelledby.\n\nNaming: lower case, words joined with hyphens: menu-item, not MenuItem or menu item (a space would make it TWO classes).\n\nIn the HTML it's class=\"price\" (no dot); in the CSS it's .price (with a dot). That's the most common mistake in Try It 2.\n\nNote the 'Contact us' link at the bottom stays blue. It isn't inside a nav, so 'nav a' doesn't match it.",
});

// 13. Try it 2
trySlide({
  n: 2, title: "Selectors", time: "8 min",
  steps: [
    ["Group: make h1 and h2 navy", ["One rule with a comma: ", ["h1, h2", "c"], "."]],
    ["Add class=\"price\" to all four prices", ["In index.html: ", ["<p class=\"price\">$16.99</p>", "c"], " and so on."]],
    ["Add class=\"special\" to the Lobster Roll", ["On its ", ["<article>", "c"], " tag."]],
    ["Write the .price, .special and nav a rules", "Copy the code on the right into style.css under Activity 2."],
    ["Check", "All four prices bold? Only the Lobster Roll highlighted?"],
  ],
  code: ["h1, h2 {", "   color: #12304A;", "}", "", ".price {", "   font-weight: bold;", "}", "", ".special {", "   background-color: #FBE9DF;", "}", "", "nav a {", "   color: #12304A;", "}"],
  tip: [["Typed ", [".price", "c"], " in the HTML? The dot only goes in the CSS."]],
  notes: "Two kinds of change in this one: HTML (adding class attributes) and CSS. That's the normal workflow: the HTML says what something is (this is a price), the CSS says how prices look.\n\nIf the Lobster Roll isn't highlighted: check class=\"special\" is on the <article>, not the <h3>.\n\nThe nav links turning navy matters in Try It 3: on a navy header they'll disappear, and we'll fix that.",
});

// 14. Colour (Example 03)
codeSlide({
  kicker: "CONCEPT 3", title: "Colour and Contrast", file: "03_colour_example.html", lines: styleOf("03_colour_example.html"), result: SH.ex03, size: 13.5,
  bullets: [
    [["color", "c"], " = text. ", ["background-color", "c"], " = behind it. Write colours as a ", ["name", "b"], ", ", ["hex", "b"], " (", ["#12304A", "c"], ") or ", ["rgb", "b"], "."],
    [["Contrast", "b"], ": WCAG AA needs ", ["4.5 : 1", "b"], " for normal text (3 : 1 for large headings)."],
    ["DevTools shows the ratio: inspect the text → click the colour square in Styles."],
  ],
  notes: "Hex: #RRGGBB, two digits each for red, green and blue, from 00 (none) to FF (full). #000000 is black, #FFFFFF is white, #FF0000 is red. rgb(255, 0, 0) is the same red written in decimal.\n\nColour names: there are about 140 (navy, teal, darkorange, hotpink…). Handy for testing; real designs use hex.\n\nContrast ties back to Class 4 and 5 (Perceivable). Lighthouse flags low contrast. The interesting one here: the brand orange from our logo fails for body text (3.1 : 1). It's fine for big bold headings or decoration, but prices need the darker orange.\n\nDevTools: right-click text → Inspect → in Styles, click the small colour square next to color. The picker shows 'Contrast ratio' with a ✓ or ✗ for AA and AAA.",
});

// 15. Try it 3
trySlide({
  n: 3, title: "Colour", time: "8 min",
  steps: [
    ["Header: navy background, white text", ["Add the ", ["header", "c"], " rule."]],
    ["Uh-oh: the h1 is still navy on navy!", ["Why? The ", ["h1, h2", "c"], " rule targets the h1 directly. Fix it with ", ["header h1", "c"], "."]],
    ["White nav links, dark orange prices, dark footer", "Add the rest of the rules on the right."],
    ["Check the contrast in DevTools", "Inspect a price → click its colour square → contrast ratio. Is it at least 4.5?"],
    ["Try your own colours", "Change any colour you like, as long as it still passes 4.5 : 1."],
  ],
  code: ["header {", "   background-color: #12304A;", "   color: white;", "}", "", "header h1 {", "   color: white;", "}", "", "nav a {", "   color: white;", "}", "", ".price {", "   color: #B5451B;", "}", "", "footer {", "   background-color: #262626;", "   color: #F2F2F2;", "}"],
  notes: "Step 2 is a planned 'bug'. Text colour is inherited from the header, BUT the h1 has its own rule (h1, h2 { color: navy }) and a rule aimed at the element itself always beats an inherited value. 'header h1' targets the h1 more specifically, so it wins. Full explanation in Concept 7.\n\nColour picking: coolors.co and Adobe Color are good palette generators. WebAIM's contrast checker (webaim.org/resources/contrastchecker) is another way to check a pair of colours.",
});

// 16. Text and fonts (Example 04)
codeSlide({
  kicker: "CONCEPT 4", title: "Text and Fonts", file: "04_text_fonts_example.html", lines: styleOf("04_text_fonts_example.html"), result: SH.ex04, size: 13.5,
  bullets: [
    [["Font stack", "b"], ": a list of fonts. The browser uses the first one it has. End with ", ["serif", "c"], " or ", ["sans-serif", "c"], "."],
    ["Size text in ", ["rem", "c"], ", not ", ["px", "c"], ": rem respects the user's own font-size setting."],
    [["line-height: 1.5–1.7", "c"], " makes paragraphs much easier to read."],
  ],
  notes: "Font stacks: we can only use fonts the visitor's computer has (or ones we download, like Google Fonts). 'Segoe UI' is on Windows; Mac users fall back to Arial; and if all else fails, any sans-serif. Font names with spaces need quotes.\n\nrem: 1rem = the browser's default font size, usually 16px. People with low vision often turn that up in their browser settings. Text sized in rem grows with it; text sized in px doesn't. That's a WCAG point (Resize Text, 1.4.4).\n\ntext-transform: uppercase vs typing in capitals: with CSS, the HTML still says 'Our Story', so screen readers read it normally (some read typed ALL CAPS letter by letter).\n\nGoogle Fonts (fonts.google.com): pick a font, copy the <link> into your <head>, then use the font name in font-family. That's the bonus in Try It 4.",
});

// 17. Try it 4
trySlide({
  n: 4, title: "Text and Fonts", time: "8 min",
  steps: [
    ["Body: a font stack, line height and text colour", ["Everything inherits it, so one rule changes the whole page."]],
    ["h1: a serif font", [["Georgia", "c"], " looks good for a restaurant name."]],
    ["h2: small capitals with spacing", [["text-transform", "c"], " and ", ["letter-spacing", "c"], "."]],
    ["Nav links: no underline, bold", ["Add to your existing ", ["nav a", "c"], " rule, or write a new one below it."]],
    ["Bonus: a Google Font", ["fonts.google.com → pick one → copy the ", ["<link>", "c"], " into your head → use its name in ", ["font-family", "c"], "."]],
  ],
  code: ["body {", "   font-family: \"Segoe UI\", Arial, sans-serif;", "   line-height: 1.6;", "   color: #262626;", "}", "", "h1 {", "   font-family: Georgia, serif;", "   font-size: 2.2rem;", "   font-weight: normal;", "}", "", "h2 {", "   font-size: 1.2rem;", "   text-transform: uppercase;", "   letter-spacing: 2px;", "}", "", "nav a {", "   text-decoration: none;", "   font-weight: bold;", "}"],
  notes: "Point out that body now has two rules (Activity 1's background, and this one). That's fine: the browser combines them. The answer key keeps them separate so you can see what each activity added.\n\nFor the Google Font bonus, 'Pacifico' or 'Lobster' are fun for the h1 of a fish-and-chip shop; 'Open Sans' or 'Lato' for body text. Remind them to keep a fallback: font-family: \"Lobster\", Georgia, serif;",
});

// 18. Box model diagram
{
  const s = panelSlide("Concept 5: The Box Model",
    "Every element on the page is a rectangular box, even text and images. The box has four layers, from the inside out: content, padding, border, margin.\n\nPadding is INSIDE the border, so it gets the background colour. Margin is OUTSIDE the border; it's always transparent, it just pushes other boxes away.\n\nShorthand: 'padding: 10px 20px' = top & bottom 10, left & right 20. Four values go clockwise from the top: top, right, bottom, left (remember 'TRouBLe').\n\nTurn on DevTools → Elements → Computed to see this exact diagram for any element.");
  const layers = [
    ["margin", "F9CC9D", "Space OUTSIDE the border. Always see-through. Pushes other boxes away."],
    ["border", "FDDD9B", "A line around the padding: width, style, colour."],
    ["padding", "C3D08B", "Space INSIDE the border. Gets the background colour."],
    ["content", "8CB6C0", "The text or image itself: width and height."],
  ];
  const bx = CX, by = 1.95, bw = 4.1, bh = 3.6, step = 0.42;
  layers.forEach(([name, col], i) => {
    s.addShape(pres.shapes.RECTANGLE, { x: bx + i * step, y: by + i * step, w: bw - 2 * i * step, h: bh - 2 * i * step,
      fill: { color: col }, line: i === 1 ? { color: "333333", width: 2 } : { color: "333333", width: 0.75, dashType: i === 0 ? "dash" : "solid" } });
    txt(s, name, { x: bx + i * step + 0.1, y: by + i * step + 0.06, w: 1.5, h: 0.3, fontSize: 12, bold: true, color: "333333", fontFace: MONO });
  });
  txt(s, "Haddock & Chips", { x: bx + 3 * step, y: by + 3 * step + 0.3, w: bw - 6 * step, h: bh - 6 * step - 0.3, fontSize: 15, bold: true, color: "1E1E1E", align: "center", valign: "middle" });
  layers.forEach(([name, col, d], i) => {
    const y = 1.95 + i * 0.9, x = CX + 4.4;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: y + 0.05, w: 0.35, h: 0.35, rectRadius: 0.05, fill: { color: col }, line: { color: "333333", width: 0.75 } });
    txt(s, name, { x: x + 0.5, y, w: 2.6, h: 0.35, fontSize: 15, bold: true, color: C.black, fontFace: MONO });
    txt(s, d, { x: x + 0.5, y: y + 0.36, w: CW - 4.9, h: 0.5, fontSize: 12, color: C.dark, valign: "top" });
  });
  richBullets(s, [
    [["padding: 10px 20px;", "c"], " = top/bottom 10px, left/right 20px. Four values go clockwise: top, right, bottom, left."],
    [["margin: 0 auto;", "c"], " + a ", ["max-width", "c"], " = a centred column."],
  ], { x: CX, y: 5.85, w: CW, h: 1.2, size: 14.5, space: 6 });
}

// 19. Example 05
codeSlide({
  kicker: "EXAMPLE 05", title: "Padding, Border, Margin", file: "05_box_model_example.html", lines: styleOf("05_box_model_example.html"), result: SH.ex05, size: 13.5,
  bullets: [
    ["Both boxes say ", ["width: 300px", "c"], ", but by default ", ["width", "c"], " is only the ", ["content", "b"], ". Padding and border are added on top: 346px."],
    [["box-sizing: border-box", "c"], " makes width include padding and border. Much easier to plan."],
  ],
  notes: "Open 05_box_model_example.html. The two boxes have identical width, padding and border, but the first is visibly wider.\n\nMost developers put this at the top of every stylesheet:\n* { box-sizing: border-box; }\nThe * is the universal selector: every element. We'll do that in Try It 5.\n\nDemo in DevTools: inspect the first box, open Computed, and hover over the box diagram. The page highlights margin (orange), border (yellow), padding (green) and content (blue). Those are the colours on the previous slide.",
});

// 20. Try it 5
trySlide({
  n: 5, title: "The Box Model", time: "12 min",
  steps: [
    ["border-box for everything, and no gap around the page", [["* { box-sizing: border-box; }", "c"], " and ", ["body { margin: 0; }", "c"], "."]],
    ["Centre the main content", [["main", "c"], ": a ", ["max-width", "c"], " and ", ["margin: 0 auto", "c"], "."]],
    ["Turn each menu item into a card", [["#menu article", "c"], ": white background, border, rounded corners, padding, margin."]],
    ["Padding for the header, footer and aside", [["header, footer { padding: 20px; }", "c"], " Then try a background on the ", ["aside", "c"], "."]],
    ["Nav links side by side (sneak peek)", ["Remove the bullets, then ", ["display: inline", "c"], " on each ", ["li", "c"], "."]],
    ["Hmm… did something disappear?", "Look at the Lobster Roll. Hold that thought until Try It 7!"],
  ],
  code: ["* {", "   box-sizing: border-box;", "}", "body {", "   margin: 0;", "}", "main {", "   max-width: 900px;", "   margin: 0 auto;", "   padding: 20px;", "}", "#menu article {", "   background-color: white;", "   border: 1px solid #D0D7DB;", "   border-radius: 8px;", "   padding: 4px 20px;", "   margin-bottom: 16px;", "}", "nav ul {", "   list-style: none;", "   padding: 0;", "}", "nav li {", "   display: inline;", "   margin-right: 24px;", "}"],
  notes: "The biggest visual change of the day. Resize the browser: the main column stays centred and never gets wider than 900px.\n\ndisplay is a big topic (block vs inline, then flexbox and grid). Here it's a sneak peek so the nav looks right. We'll cover it properly when we do layout.\n\nStep 6 is deliberate: the Lobster Roll's .special highlight from Try It 2 is gone, because '#menu article' (white background) beats '.special'. Don't explain it yet. Students solve it with DevTools in Try It 7.",
});

// 21. Images (Example 06)
codeSlide({
  kicker: "CONCEPT 6", title: "Styling Images", file: "06_styling_images_example.html", lines: styleOf("06_styling_images_example.html"), result: SH.ex06, size: 13.5,
  bullets: [
    [["max-width: 100%", "c"], " + ", ["height: auto", "c"], ": the image shrinks on small screens and keeps its shape. Use it on every site."],
    [["border-radius", "c"], " rounds corners. ", ["50%", "c"], " on a square = a circle. ", ["object-fit: cover", "c"], " crops to fill the box."],
  ],
  notes: "max-width: 100% is the single most useful image rule. Without it, an 800px image on a 375px phone makes the whole page scroll sideways. With it, the image is never wider than its container. height: auto keeps the proportions (the HTML height attribute would otherwise stretch it).\n\nThis connects to last class: keep the width and height attributes in the HTML (so the browser reserves space), and let CSS decide the displayed size.\n\nCentring an image: images are inline by default (they sit in a line of text like a big letter), so margin: auto doesn't centre them. display: block makes it a box of its own, then margin: 0 auto centres it.\n\nobject-fit: cover is how 'profile picture' circles are made from rectangular photos.",
});

// 22. Try it 6
trySlide({
  n: 6, title: "Images", time: "8 min",
  steps: [
    ["Make every image flexible", ["Then make the browser narrow. Does the photo shrink?"]],
    ["Round the corners of the main photo", [["figure img", "c"], ": ", ["display: block", "c"], " and a ", ["border-radius", "c"], "."]],
    ["Use YOUR photo from last class", ["Copy ", ["photo_1200.jpg", "c"], " into activity/images. Update the ", ["src", "c"], ", ", ["width", "c"], ", ", ["height", "c"], " and ", ["alt", "c"], "."]],
    ["Check the download size", "DevTools → Network → Img → reload. Is it under 150 KB?"],
    ["Bonus: a round thumbnail", [["object-fit: cover", "c"], " and ", ["border-radius: 50%", "c"], ", like Example 06."]],
  ],
  code: ["img {", "   max-width: 100%;", "   height: auto;", "}", "", "figure {", "   margin: 0 0 16px;", "}", "", "figure img {", "   display: block;", "   border-radius: 12px;", "}"],
  tip: [["Your photo's width and height are in File Explorer → Properties → Details (Lab 1)."]],
  notes: "This ties the two classes together: the optimized image from Class 7 goes onto the page from Class 6's wireframe, styled with today's CSS.\n\nTesting flexible images: drag the browser narrower, or use DevTools device mode (Ctrl + Shift + M) and pick a phone.\n\nIf a student's photo is portrait (taller than wide), it will be very tall. That's a good conversation: choose a landscape crop for a banner photo, or limit it with max-height and object-fit: cover.",
});

// 23. The cascade
{
  const s = panelSlide("Concept 7: The Cascade — Who Wins?",
    "When two rules set the same property on the same element, the browser has to choose. That decision is 'the cascade': the C in CSS.\n\n1. Specificity: the more specific selector wins. Inline style beats an id, an id beats a class, a class beats an element. (A combination like '#menu article' counts the id AND the element, so it beats '.special'.)\n2. If they're equally specific, the rule that comes LATER in the file wins.\n3. Inheritance: some properties (color, font-family, line-height) pass down from parent to child, unless the child has its own rule. Others (background, border, padding, margin) don't.\n\nBefore showing the next slide, play 'Who wins?': show the CSS from Example 07 and have students predict the colour of each line.");
  const ladder = [
    ["Inline style", "style=\"…\"", "", C.red],
    ["ID", "#special", "", C.purple],
    ["Class", ".notice", "", C.orange],
    ["Element", "p", "", C.teal],
  ];
  ladder.forEach(([t, ex, tag, col], i) => {
    const y = 1.95 + i * 0.82, w = 3.6 - i * 0.35;
    card(s, { x: CX, y, w, h: 0.68, fill: col, border: col });
    txt(s, t, { x: CX + 0.2, y, w: 1.6, h: 0.68, fontSize: 16, bold: true, color: C.white, valign: "middle" });
    txt(s, ex, { x: CX + 1.75, y, w: w - 1.85, h: 0.68, fontSize: 13, bold: true, color: C.white, fontFace: MONO, valign: "middle" });
  });
  const rx = CX + 4.0, rw = CW - 4.0;
  const rules = [
    ["1", "More specific wins", "Inline beats id beats class beats element."],
    ["2", "A tie? The later rule wins", "Same specificity: whichever comes last in the file."],
    ["3", "Inheritance", "Text styles like color and font-family pass down to children that don't have their own rule."],
  ];
  rules.forEach(([n, t, d], i) => {
    const y = 1.95 + i * 1.1;
    numCircle(s, n, rx, y, C.dark, 0.42);
    txt(s, t, { x: rx + 0.55, y: y - 0.02, w: rw - 0.55, h: 0.35, fontSize: 15, bold: true, color: C.black });
    txt(s, d, { x: rx + 0.55, y: y + 0.33, w: rw - 0.55, h: 0.7, fontSize: 12.5, color: C.dark, valign: "top" });
  });
  card(s, { x: CX, y: 5.45, w: CW, h: 1.35, fill: C.dark, border: C.dark });
  richBullets(s, [
    [["Try It 3 mystery solved: ", "b"], "header passed white down to the h1, but the ", ["h1, h2", "b"], " rule targets the h1 ", ["directly", "i"], ". A rule aimed at the element always beats an inherited value."],
  ], { x: CX + 0.3, y: 5.45, w: CW - 0.6, h: 1.35, size: 14, bullet: false, color: C.white, valign: "middle" });
}

// 24. Example 07 (the reveal)
codeSlide({
  kicker: "EXAMPLE 07", title: "Who Wins? Check Your Predictions", file: "07_cascade_example.html", lines: [...styleOf("07_cascade_example.html"), "", ...bodyOf("07_cascade_example.html")], lang: "mixed", result: SH.ex07, size: 13,
  bullets: [
    ["1 blue: the later ", ["p", "c"], " rule wins the tie. 2 orange: class beats element. 3 purple: id beats class. 4 red: inline beats everything."],
    ["5 and the h1: no rule of their own, so they ", ["inherit", "b"], " gray from body."],
  ],
  notes: "Show the CSS first (cover the result, or show the previous slide), and get predictions for each line. Then reveal.\n\nThen open it in DevTools: inspect line 3. In Styles, #special is at the top; .notice and both p rules are listed below it with color crossed out. Inherited styles appear under 'Inherited from body'.\n\nThere's one level stronger than inline: !important. Mention that it exists and that students should avoid it: it's usually a sign the selectors need fixing.",
});

// 25. DevTools Styles pane
{
  const s = panelSlide("Your Best Friend: DevTools Styles",
    "Right-click anything → Inspect. The Styles pane (Edge/Chrome; Firefox calls it Rules) shows every rule that applies to the selected element, most specific first.\n\n- Crossed-out declarations lost the cascade to a rule higher up.\n- A yellow warning triangle means the browser didn't understand the line (typo, colour vs color, missing unit).\n- Checkboxes appear on hover: untick to switch a declaration off.\n- Click any value to edit it live. Click a colour square for the colour picker and contrast ratio.\n- The link on the right (style.css:42) jumps to the line in the Sources panel.\n\nIMPORTANT: DevTools changes vanish on reload. When you find something you like, copy it into style.css.");
  const px = CX, pw = 4.45, py = 1.95, ph = 4.9;
  s.addShape(pres.shapes.RECTANGLE, { x: px, y: py, w: pw, h: ph, fill: { color: "FFFFFF" }, line: { color: "B0B8BC", width: 1 } });
  s.addShape(pres.shapes.RECTANGLE, { x: px, y: py, w: pw, h: 0.36, fill: { color: "F1F3F4" }, line: { color: "B0B8BC", width: 1 } });
  txt(s, "Styles     Computed     Layout", { x: px + 0.15, y: py, w: pw - 0.3, h: 0.36, fontSize: 11, color: "333333", valign: "middle" });
  s.addShape(pres.shapes.LINE, { x: px + 0.12, y: py + 0.34, w: 0.5, h: 0, line: { color: "1A73E8", width: 2 } });
  const rows = [
    ["#menu .special {", "sel", "style.css:150"], ["  background-color: ■ #FBE9DF;", "decl"], ["  border-color: ■ #E97132;", "decl"], ["}", "sel"], ["", ""],
    ["#menu article {", "sel", "style.css:124"], ["  background-color: ■ white;", "strike"], ["  border: 1px solid ■ #D0D7DB;", "decl"], ["  padding: 4px 20px;", "decl"], ["}", "sel"], ["", ""],
    [".special {", "sel", "style.css:23"], ["  background-color: ■ #FBE9DF;", "strike"], ["  colour: red;", "warn"], ["}", "sel"],
  ];
  rows.forEach(([t, k, link], i) => {
    const y = py + 0.5 + i * 0.285;
    const strike = k === "strike";
    txt(s, t, { x: px + 0.2, y, w: pw - 1.4, h: 0.28, fontSize: 10.5, fontFace: MONO, color: k === "sel" ? "881280" : (strike ? "9AA0A6" : (k === "warn" ? "9AA0A6" : "1A1AA6")),
      strike: strike || k === "warn" ? "sngStrike" : undefined });
    if (link) txt(s, link, { x: px + pw - 1.25, y, w: 1.1, h: 0.28, fontSize: 9.5, color: "5F6368", align: "right", underline: { style: "sng" } });
    if (k === "warn") txt(s, "⚠", { x: px + 1.75, y: y - 0.02, w: 0.3, h: 0.3, fontSize: 12, color: "E37400", fontFace: "Segoe UI Symbol" });
  });
  const notes = [
    ["Crossed out", "Lost to a stronger rule above it.", C.mid],
    ["⚠ Warning", "The browser didn't understand it (here: colour).", "E37400"],
    ["Click to edit", "Change any value live. Untick to switch it off.", C.blue],
    ["style.css:124", "Which file and line the rule came from.", C.teal],
  ];
  notes.forEach(([t, d, col], i) => {
    const y = 1.95 + i * 1.0, x = CX + 4.75, w = CW - 4.75;
    txt(s, t, { x, y, w, h: 0.35, fontSize: 15, bold: true, color: col });
    txt(s, d, { x, y: y + 0.35, w, h: 0.6, fontSize: 12.5, color: C.dark, valign: "top" });
  });
  card(s, { x: CX + 4.75, y: 6.05, w: CW - 4.75, h: 0.8, fill: "FBE9DF", border: "F2C6AC" });
  txt(s, "DevTools changes vanish on reload. Copy the good ones to style.css!", { x: CX + 4.9, y: 6.05, w: CW - 5.05, h: 0.8, fontSize: 12.5, bold: true, color: C.dark, valign: "middle" });
}

// 26. Try it 7
trySlide({
  n: 7, title: "DevTools Detective", time: "10 min",
  steps: [
    ["The mystery: where did the Lobster Roll's highlight go?", "It was peach after Try It 2. Now it's white."],
    ["Right-click the Lobster Roll card → Inspect", ["Click the ", ["<article class=\"special\">", "c"], " line in Elements."]],
    ["Find .special in Styles", ["Its background is ", ["crossed out", "b"], ". Which rule beat it, and why?"]],
    ["Fix it in style.css", [["#menu article", "c"], " has an id, so it beats ", [".special", "c"], ". Add the id to your rule too (on the right)."]],
    ["Play with it", "Click values in Styles and change them live. Untick some. Then reload: what happens?"],
  ],
  code: ["/* An id + a class beats an id + an element */", "#menu .special {", "   background-color: #FBE9DF;", "   border-color: #E97132;", "}"],
  tip: [["Open the ", ["Computed", "b"], " tab too: it shows the box model diagram for the card, with its real padding, border and margin."]],
  notes: "Answer: '#menu article' (one id + one element) beats '.special' (one class). It doesn't matter that .special came first or last; specificity is checked before order.\n\nThe fix: '#menu .special' (one id + one class) beats '#menu article' (one id + one element). Something like 'article.special' isn't enough, because it has no id. This is why many developers avoid id selectors in CSS: they're so strong they're hard to override. Classes for styling, ids for links and labels is a good habit.\n\nStep 5: changes made in DevTools disappear on reload. It's a sandbox for trying things, and then you copy what works into the file.",
});

// 27. Make it yours
{
  const s = activitySlide("Make It Yours, Then Submit (10 min, or homework)",
    "If time is short, students submit what they have and finish 'make it yours' as homework.\n\nWhat to collect on Brightspace: a screenshot of their page in the browser, and their style.css file. Checking style.css tells you whether they actually wrote the rules (look for their own colours and at least two changes beyond the answer key).");
  card(s, { x: 0.75, y: 1.45, w: 7.6, h: 5.5 });
  txt(s, "Change at least three things. Ideas:", { x: 1.05, y: 1.65, w: 7, h: 0.4, fontSize: 18, bold: true, color: C.black });
  const ideas = [
    [["Your own ", ["colour palette", "b"], " (coolors.co). Check every pair passes 4.5 : 1."]],
    [["A ", ["Google Font", "b"], " for the h1 (fonts.google.com)."]],
    [["A ", ["hover", "b"], " style for links: ", ["nav a:hover { text-decoration: underline; }", "c"]]],
    [["Style the ", ["aside", "b"], " (hours): background, border, rounded corners."]],
    [["Make the ", ["footer", "b"], " centred, with smaller text."]],
    [["Make the prices bigger, or put a ", ["border-bottom", "c"], " under each h2."]],
  ];
  ideas.forEach((segs, i) => {
    const y = 2.25 + i * 0.68;
    numCircle(s, i + 1, 1.05, y, C.orange, 0.4);
    richBullets(s, segs, { x: 1.65, y, w: 6.5, h: 0.6, size: 14.5, bullet: false, color: C.dark, valign: "middle" });
  });
  card(s, { x: 8.75, y: 1.45, w: 3.85, h: 5.5, fill: C.dark, border: C.dark });
  txt(s, "Submit to Brightspace", { x: 9.0, y: 1.7, w: 3.4, h: 0.45, fontSize: 18, bold: true, color: C.white });
  txt(s, [
    { text: "1. A screenshot of your page in the browser (Win + Shift + S)", options: { breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "2. Your style.css file", options: { breakLine: true } },
    { text: " ", options: { fontSize: 10, breakLine: true } },
    { text: "Stuck? ", options: { bold: true, color: "F2C6AC" } },
    { text: "Compare with activity/answer/style.css, one section at a time. Then use DevTools to find what's different." },
  ], { x: 9.0, y: 2.3, w: 3.4, h: 4.4, fontSize: 14, color: "D9D9D9", valign: "top" });
}

// 28. Discussion
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  s.addShape(pres.shapes.OVAL, { x: 7.6, y: 0.9, w: 5.4, h: 5.4, fill: { type: "none" }, line: { color: C.ring, width: 28 } });
  txt(s, "How did you make out?", { x: 0.8, y: 0.8, w: 7, h: 0.9, fontSize: 40, color: C.black });
  richBullets(s, [
    ["Which change made the ", ["biggest difference", "b"], " to how the page looks?"],
    ["What was your most annoying bug? How did you find it?"],
    ["Why use a class for the prices instead of styling each one inline?"],
    ["Did any of your colour choices fail the contrast check?"],
  ], { x: 0.8, y: 2.0, w: 6.5, h: 4.5, size: 20, space: 16 });
  txt(s, "HTML = what it is\nCSS = how it looks", { x: 8.4, y: 2.8, w: 3.8, h: 1.6, fontSize: 22, bold: true, color: C.teal, align: "center", valign: "middle" });
  s.addNotes("Usual answers: the box model round (centring main, cards) makes the biggest visual jump. Common bugs: a missing semicolon or brace, 'colour', a class typed with a dot in the HTML, and the specificity mystery.\n\nWhy classes: change one rule and every price updates; the HTML stays clean; and inline styles are so specific they're hard to override later.");
}

// 29. Cheat sheet
{
  const s = panelSlide("Today's CSS Cheat Sheet",
    "Everything we used today, in one place. It's also in the README. Students can keep this slide open while they finish 'make it yours'.");
  table(s, [
    ["Property", "Example", "What it does"],
    ["color", "color: #12304A;", "Text colour"],
    ["background-color", "background-color: white;", "Colour behind the element"],
    ["font-family", "font-family: Georgia, serif;", "The font (with fallbacks)"],
    ["font-size", "font-size: 1.2rem;", "Text size (prefer rem)"],
    ["font-weight / font-style", "font-weight: bold;", "Bold / italic"],
    ["line-height", "line-height: 1.6;", "Space between lines"],
    ["text-align / text-transform", "text-align: center;", "Alignment / capitals"],
    ["text-decoration", "text-decoration: none;", "Remove or add underlines"],
    ["padding / border / margin", "padding: 10px 20px;", "The box model layers"],
    ["max-width + margin: auto", "margin: 0 auto;", "A centred column"],
    ["border-radius", "border-radius: 8px;", "Rounded corners (50% = circle)"],
    ["box-sizing", "box-sizing: border-box;", "Width includes padding and border"],
  ], { x: CX, y: 1.9, w: CW, colW: [2.45, 2.85, 2.25], size: 12, rowH: 0.38, codeCols: [0, 1] });
}

// 30. Recap
{
  const s = panelSlide("Class Recap and Next Steps",
    "Next: more selectors (pseudo-classes like :hover and :focus), styling navigation menus and tables, and checking our code with the W3C validator. Then layout with flexbox.");
  labelBullets(s, [
    { label: "Recap:", text: [
      "CSS rules: selector { property: value; }. Link one external style.css to every page.",
      "Select by element, .class, #id, groups (a, b) and descendants (nav a).",
      "Colour with contrast of 4.5 : 1 or better. Fonts in stacks, sizes in rem.",
      "Every element is a box: content, padding, border, margin.",
      "img { max-width: 100%; height: auto; } on every site.",
      "When rules clash: more specific wins, then the later one. DevTools shows you which.",
    ] },
    { label: "Next class:", text: [
      "More selectors (:hover, :focus), navigation menus, tables, and validating our code.",
    ] },
  ], { size: 20, sub: 15.5 });
}

pres.writeFile({ fileName: OUT }).then((f) => console.log("wrote " + f));
