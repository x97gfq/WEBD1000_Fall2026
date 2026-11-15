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

// JavaScript, for Class 1's tooltip and toast slides. Dark+ colours again:
// blue keywords, yellow function names, light-blue variables and properties, orange strings.
const JSC = { kw: "569CD6", fn: "DCDCAA", vari: "9CDCFE" };
const JSKW = /^(const|let|var|function|return|new|if|else|for|while|true|false|null|undefined|this)\b/;
function highlightJS(lines) {
  let inBlock = false;
  return lines.map((line) => {
    const toks = [], push = tokPusher(toks);
    let i = 0;
    while (i < line.length) {
      if (inBlock) {
        const e = line.indexOf("*/", i);
        if (e < 0) { push(line.slice(i), HL.comment); i = line.length; }
        else { push(line.slice(i, e + 2), HL.comment); i = e + 2; inBlock = false; }
        continue;
      }
      if (line.startsWith("/*", i)) { inBlock = true; continue; }
      if (line.startsWith("//", i)) { push(line.slice(i), HL.comment); break; }
      const ch = line[i], prev = line[i - 1] || "";
      if (ch === '"' || ch === "'" || ch === "`") {
        let j = i + 1;
        while (j < line.length && line[j] !== ch) { if (line[j] === "\\") j++; j++; }
        push(line.slice(i, Math.min(j + 1, line.length)), HL.val); i = j + 1; continue;
      }
      const kw = line.slice(i).match(JSKW);
      if (kw && !/[\w$.]/.test(prev)) { push(kw[0], JSC.kw); i += kw[0].length; continue; }
      const num = line.slice(i).match(/^\d+(\.\d+)?/);
      if (num && !/[\w$]/.test(prev)) { push(num[0], HL.num); i += num[0].length; continue; }
      const id = line.slice(i).match(/^[A-Za-z_$][\w$]*/);
      if (id) {
        const call = /^\s*\(/.test(line.slice(i + id[0].length));
        push(id[0], call ? JSC.fn : prev === "." ? HL.prop : JSC.vari);
        i += id[0].length; continue;
      }
      push(ch, HL.punct); i++;
    }
    return toks;
  });
}

// lang: "html", "css", "js", or "mixed" (CSS inside <style>…</style>, and lines from a .css file after a "/* in style.css" comment)
function highlight(lines, lang) {
  if (lang === "html") return highlightHTML(lines);
  if (lang === "css") return highlightCSS(lines);
  if (lang === "js") return highlightJS(lines);
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
// trivia: a question about where all this came from, rather than about code. It still counts as
// a numbered pop quiz, because every one of them becomes a question on the Week 11 quiz.
function quiz({ kicker, trivia, q, options, answer, why, code, lang = "css", codeTitle, revealImg, notes, revealNotes }) {
  kicker = kicker || `POP QUIZ ${++quizNo}` + (trivia ? "  ·  TRIVIA" : "");
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

// A copy of an example page with extra CSS, and optionally JavaScript that runs once the page
// has loaded. The JavaScript is how the JS-component slides get their pictures: a screenshot
// can't click a button, so the build opens the modal, the toast or the offcanvas itself, and
// scrolls to the part of the page the slide is about.
// Transitions and animations are turned off, so the shot catches the finished state, not a fade.
// The <base> is the example's own folder, so its images and stylesheet still load.
function variant(file, name, css, js) {
  const base = "file:///" + path.dirname(path.join(REPO, file)).replace(/\\/g, "/") + "/";
  let html = fs.readFileSync(path.join(REPO, file), "utf8")
    .replace("<head>", `<head><base href="${base}">`)
    .replace("</head>", `<style>*, *::before, *::after { transition: none !important; animation: none !important }\n${css || ""}</style></head>`);
  if (js) html = html.replace("</body>", `<script>window.addEventListener("load", function () { ${js} });</script></body>`);
  fs.writeFileSync(path.join(SHOTS, name + "_src.html"), html);
  return path.relative(REPO, path.join(SHOTS, name + "_src.html"));
}

// Snippets for variant()'s `js` argument
const OPEN = {
  modal: (id) => `bootstrap.Modal.getOrCreateInstance(document.getElementById("${id}")).show();`,
  collapse: (id) => `bootstrap.Collapse.getOrCreateInstance(document.getElementById("${id}")).show();`,
  offcanvas: (id) => `bootstrap.Offcanvas.getOrCreateInstance(document.getElementById("${id}")).show();`,
  toast: (id) => `bootstrap.Toast.getOrCreateInstance(document.getElementById("${id}")).show();`,
  tooltip: (sel) => `bootstrap.Tooltip.getOrCreateInstance(document.querySelector('${sel}')).show();`,
  slide: (id, n) => `bootstrap.Carousel.getOrCreateInstance(document.getElementById("${id}")).to(${n});`,
};
// Start the shot at this element: headless Edge screenshots the top of the page whatever the
// scroll position, so to frame a component further down we delete what comes before it.
const onlyFrom = (sel) => `var t = document.querySelector('${sel}'); var p = t.parentElement;
  while (p.firstElementChild && p.firstElementChild !== t) p.firstElementChild.remove();`;

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

// A page of component tiles for the JavaScript components, 4 across and 3 down.
// Each: [name, what it is, docs page, W3Schools page or null, "today" or null]
// The six we actually build today are purple and badged; the rest are there to look up.
function jsComponentSlide(title, items, footer, notes) {
  const s = plainSlide(title, notes);
  const cols = 4, gap = 0.2, tw = (W - 1.5 - gap * (cols - 1)) / cols, th = 1.45;
  items.forEach(([name, d, doc, w3, today], i) => {
    const x = 0.75 + (i % cols) * (tw + gap), y = 1.4 + Math.floor(i / cols) * (th + 0.17);
    const color = today ? C.purple : C.teal;
    card(s, { x, y, w: tw, h: th, fill: today ? "FAF5FA" : C.white, border: today ? "E2CCE0" : "D0D7DB" });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.1, h: th, fill: { color }, line: { color } });
    txt(s, [{ text: name, options: { hyperlink: { url: DOCS + doc + "/", tooltip: name + " in Bootstrap's docs" } } }],
      { x: x + 0.25, y: y + 0.1, w: tw - (today ? 1.05 : 0.4), h: 0.38, fontSize: 15, bold: true, color: C.navy });
    if (today) pill(s, "TODAY", x + tw - 0.95, y + 0.12, C.purple, 0.8);
    txt(s, d, { x: x + 0.25, y: y + 0.48, w: tw - 0.4, h: 0.58, fontSize: 11.5, color: C.dark, valign: "top" });
    const links = [{ text: "Docs", options: { hyperlink: { url: DOCS + doc + "/" } } }];
    if (w3) links.push({ text: "   ·   ", options: { color: C.mid } }, { text: "W3Schools", options: { hyperlink: { url: W3 + w3 + ".php" } } });
    txt(s, links, { x: x + 0.25, y: y + th - 0.36, w: tw - 0.4, h: 0.28, fontSize: 11, bold: true, color: C.teal });
  });
  card(s, { x: 0.75, y: 6.38, w: W - 1.5, h: 0.62, fill: "FBE9DF", border: "F2C6AC" });
  richBullets(s, [footer], { x: 1.0, y: 6.38, w: W - 2.0, h: 0.62, size: 14, bullet: false, valign: "middle" });
  return s;
}

// The three levels of JavaScript this week: none, one line, a few lines.
// Each band: a number, a title, a sentence, the components it covers, and the code.
function ladderSlide(title, levels, notes) {
  const s = plainSlide(title, notes);
  const bh = 1.72, lw = 6.0;
  levels.forEach(([head, d, who, code, color, lang], i) => {
    const y = 1.4 + i * (bh + 0.16);
    card(s, { x: 0.75, y, w: W - 1.5, h: bh });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.75, y, w: 0.12, h: bh, fill: { color }, line: { color } });
    numCircle(s, i + 1, 1.05, y + 0.26, color, 0.5);
    txt(s, head, { x: 1.72, y: y + 0.2, w: lw - 1.0, h: 0.4, fontSize: 19, bold: true, color });
    richBullets(s, [d], { x: 1.72, y: y + 0.62, w: lw - 1.0, h: 0.5, size: 13.5, bullet: false, color: C.dark });
    txt(s, who, { x: 1.72, y: y + 1.18, w: lw - 1.0, h: 0.42, fontSize: 12, italic: true, color: C.grey, valign: "top" });
    codeBox(s, code, { x: 0.75 + lw + 0.3, y: y + 0.16, w: W - 1.5 - lw - 0.6, h: bh - 0.32, maxSize: 13.5, lang: lang || "js" });
  });
  return s;
}

// One line of JavaScript, pulled apart: the line in a code box, then a row per piece.
function anatomySlide(title, code, parts, notes) {
  const s = plainSlide(title, notes);
  codeBox(s, code, { x: 0.75, y: 1.35, w: W - 1.5, h: 1.35, maxSize: 17, fitHeight: true, lang: "js" });
  const ry = 2.95, rh = (6.95 - ry - 0.1 * (parts.length - 1)) / parts.length;
  parts.forEach(([name, d], i) => {
    const y = ry + i * (rh + 0.1);
    card(s, { x: 0.75, y, w: W - 1.5, h: rh });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.75, y, w: 0.1, h: rh, fill: { color: C.orange }, line: { color: C.orange } });
    txt(s, name, { x: 1.0, y, w: 4.5, h: rh, fontSize: 14, fontFace: MONO, bold: true, color: "8A2B0B", valign: "middle" });
    richBullets(s, [d], { x: 5.7, y, w: W - 6.45, h: rh, size: 14.5, bullet: false, valign: "middle" });
  });
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
// A screenshot can't click a button, so the JS-component shots are taken from a copy of the
// page with a line of JavaScript injected: open this modal, show that toast, scroll to there.
const SH = {
  // Customizing (Class 1)
  ex01: shot("01_customizing_example.html", "ex01", 900, 200),
  // :hover can't be screenshotted, so this copy applies Bootstrap's own .btn:hover rule to both buttons
  ex01h: shot(variant("01_customizing_example.html", "ex01h",
    ".btn.old-way, .btn.new-way { background-color: var(--bs-btn-hover-bg); border-color: var(--bs-btn-hover-border-color); }"), "ex01h", 900, 200),

  // Bootstrap's JavaScript components (Class 1)
  nojs: shot("js/00_no_script_example.html", "nojs", 900, 540),
  modal: shot(variant("js/01_modal_example.html", "modal", "", OPEN.modal("orderModal")), "modal", 780, 470),
  accordion: shot(variant("js/02_accordion_example.html", "acc", "", onlyFrom(".accordion")), "acc", 880, 260),
  carousel: shot(variant("js/03_carousel_example.html", "car", "", onlyFrom(".carousel")), "car", 880, 330),
  offcanvas: shot(variant("js/04_offcanvas_example.html", "off", "", OPEN.offcanvas("sideMenu")), "off", 700, 440),
  tooltip: shot(variant("js/05_tooltip_example.html", "tip", "", OPEN.tooltip('[data-bs-placement="bottom"]')), "tip", 780, 250),
  toast: shot(variant("js/06_toast_example.html", "toast", "", OPEN.toast("orderToast")), "toast", 700, 300),
  allD: shot("js/07_all_together_example.html", "allD", 1000, 1450),
  allP: phone("js/07_all_together_example.html", "allP", 360, 700),
  allT: shot(variant("js/07_all_together_example.html", "allT", "", OPEN.toast("orderToast") + onlyFrom("#menu")), "allT", 900, 430),

  // Tailwind and Harbourside (Class 2)
  ex02: shot("02_bootstrap_cards_example.html", "ex02", 1100, 300),
  ex03: shot("03_tailwind_cards_example.html", "ex03", 1100, 300),
  act0: shot("activity/index.html", "act0", 900, 700),
  actD: shot("activity/answer/index.html", "actD", 1100, 1500),
  actP: phone("activity/answer/index.html", "actP", 360, 700),
};

// The example files, for the code on the slides
const JS01 = fileLines("js/01_modal_example.html");
const JS02 = fileLines("js/02_accordion_example.html");
const JS03 = fileLines("js/03_carousel_example.html");
const JS05 = fileLines("js/05_tooltip_example.html");
const JS06 = fileLines("js/06_toast_example.html");
// Students' index.html is in activity/, not activity/answer/, so the image paths lose the ../
const ANS = fileLines("activity/answer/index.html").map((l) => l.replace("../images/", "images/"));
const ANSCSS = fileLines("activity/answer/style.css");

// =====================================================================
// WEEK 11, CLASS 1: Customizing Bootstrap, and the parts that move
// =====================================================================
section("Class 1 - Customizing Bootstrap, and the Parts That Move");

titleSlide("Week 11, Class 1", "Making Bootstrap look like your site, and making it move.",
  "Welcome back. Last week was Bootstrap's grid and its components: we turned a wireframe into a layout, then gave it a navbar and cards.\n\nTwo things were missing. First, it looks like every other Bootstrap site. Second, almost nothing on it moves. Today is both.\n\nFirst half: customizing, so it looks like yours. Second half: the components that need JavaScript - a modal, an accordion, a carousel, a toast - and the first few lines of JavaScript you have ever written.\n\nSay this out loud early: you have not learned JavaScript, and you are not expected to today. Most of what we do needs none at all. The last two components need one line and three lines, and we read them word by word.\n\nNext class is Tailwind and the Harbourside rebuild. The week after that, JavaScript properly, starting with form validation.");

setupSlide("Attendance and a Quick Check-in", [
  ["OUTLOOK", "Are you checking your school Outlook email?"], ["TEAMS", "Are you signed into Teams?"],
  ["TEAMS", "Are you following our Teams channel and General chat?"], ["BRIGHTSPACE", "Have you read Assignment 4 on Brightspace?"],
], "Take attendance while everyone checks these four things.\n\nAssignment 4 is a two-page Bootstrap site. Today's customizing section is its 'your own styling' requirement, and any one of today's interactive components is an easy way to make it stand out. Point that out as you go.");

setupSlide("Is Your Environment Set Up?", [
  ["GIT", "Did you run git pull? You should see a new Week 11 folder."],
  ["VS CODE", "Is the Week 11 folder open in VS Code? Look for the js folder in it."],
  ["LIVE SERVER", "Does js/01_modal_example.html open with Go Live, and does the button open a box?"],
  ["INTERNET", "Bootstrap loads from the web. No blue button means no Bootstrap."],
], "Help anyone who's stuck now, so they aren't lost during the Try Its. Everything today is in the js folder, and every page in it opens on its own.\n\nIf someone's git pull complains, they probably edited a file in 10_CSS_Frameworks: they can copy their work somewhere safe and run git stash, then pull again.\n\nIf js/01_modal_example.html opens but the button does nothing, their internet or the CDN is the problem, not their code. That is worth saying out loud, because 'the button does nothing' is today's theme.");

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
], "Same as always: 30 seconds with a neighbour, then take answers, then go through the four points.\n\nLean on point 1 (ours goes after theirs) and point 3 (the navbar needs the JavaScript). Those are the two doors into today.\n\nThen ask two questions. Does anyone's layout from last week look like their own site? (No: it looks like Bootstrap, in loud colours.) And does anything on it move, apart from the menu button? (No.) That's today.");

overviewSlide("Week 11, Class 1: Overview", "Customizing, and the Parts That Move", [
  "Customizing Bootstrap: your CSS after theirs, and its variables",
  "Java is not JavaScript, and what Bootstrap's script actually does",
  "No JavaScript at all: modal, accordion, carousel, offcanvas",
  "Your first lines of JavaScript: a tooltip, then a toast",
], [["Customize", "Your own CSS"], ["Click it", "Six components"], ["Write it", "Four lines of JS"]],
  "Rough plan:\n- 10 min: check-in, environment, recall, overview\n- 10 min: customizing, Example 01, Pop Quiz 1\n- 16 min: Java vs JavaScript, the three levels, the component list, Pop Quiz 2\n- 30 min: modal, accordion, carousel, offcanvas, Try Its 1 to 3, Pop Quizzes 3 and 4\n- 28 min: reading a line of JavaScript, tooltips and toasts, Try Its 4 and 5, Pop Quizzes 5 and 6\n- 5 min: the all-together page, and the recap\n\nIf you are running behind: cut Try It 3 (the carousel) and the offcanvas slide, and make the all-together page a link rather than a demo. Do not cut the toast: it is the one that makes JavaScript click for people.\n\nSix pop quizzes today and four next class. All ten are on the Week 11 quiz, so tell them that at the first one.");

// ---------------------------------------------------------------- customizing
dividerSlide("Customizing Bootstrap", "Your CSS after theirs   ·   Bootstrap 5's variables",
  "Every Bootstrap site looks like Bootstrap until you change it. The question is how.");

codeSlide({
  kicker: "CUSTOMIZING · EXAMPLE 01", title: "Changing a Button's Colour", file: "11_Customizing_Bootstrap_Tailwind/01_customizing_example.css",
  lines: cssOf("01_customizing_example.html").slice(2), lang: "css", frames: [[{ img: SH.ex01, label: "NORMAL" }], [{ img: SH.ex01h, label: "WHILE YOU HOVER" }]], maxFrameH: 2.9, size: 14,
  bullets: [
    ["Button 1 changes ", ["background-color", "c"], ". Hover, and it goes ", ["blue", "b"], "."],
    ["Button 2 changes Bootstrap's ", ["variables", "b"], ". Every state uses them."],
  ],
  notes: "Open 01_customizing_example.html and hover over each button.\n\nPart 0: overriding any rule works the normal way. Bootstrap's body rule and ours tie, ours is later, ours wins. The font is Georgia.\n\nPart 1 is the way from last year's slides (and lots of tutorials). It looks right until you hover: Bootstrap has a .btn:hover rule, which is more specific than ours, and it sets the background to var(--bs-btn-hover-bg). That's still blue. The bottom picture applies that exact rule to both buttons.\n\nPart 2 is how Bootstrap 5.3 wants it done. A Bootstrap button reads all of its colours from variables (custom properties, the --name ones). Set those, and the normal, hover and active states all follow.\n\nWhere do the variable names come from? Inspect a button in DevTools: the .btn-primary rule lists them all.\n\nFor bigger changes (every primary colour on the site), professional projects use Bootstrap's Sass source. That's beyond this course. Ready-made Bootstrap themes are another option; Bootswatch has free ones.",
});

quiz({
  q: "This is in your CSS, after Bootstrap's. What colour is the button while you hover over it?",
  code: [".btn-primary {", "   background-color: #ff5722;", "   border-color: #ff5722;", "}"], codeTitle: "LAST YEAR'S SLIDE",
  options: ["Orange, the same as always", "Bootstrap's blue", "A slightly darker orange", "See-through: it has no colour"], answer: 1,
  why: ["Bootstrap's `.btn:hover` (two classes' worth) beats `.btn-primary` (one), so its hover rule wins.", "That rule uses `var(--bs-btn-hover-bg)`, which is still blue. Change the variables instead: Example 01, part 2."],
  revealImg: SH.ex01h,
  notes: "The code is word for word from last year's Customizing Bootstrap slide. Specificity from Week 6.\n\nThis is Pop Quiz 1 of 10. Tell them now: every pop quiz this week turns into a question on the Week 11 quiz, so write the answers down.",
  revealNotes: "Answer: B.\n\nThe picture is Example 01 with Bootstrap's hover rule applied: button 1 is blue.\n\nSpecificity: .btn:hover is a class plus a pseudo-class, so (0,2,0). .btn-primary is (0,1,0). Ours comes later, but specificity is checked first.\n\nKeep this one in your pocket: Pop Quiz 10 next class is the same question where the rules DO tie, and then being later is enough.",
});

// ---------------------------------------------------------------- the parts that move
dividerSlide("The Parts That Move", "Bootstrap's JavaScript   ·   six components   ·   four lines of your own",
  "Second half of the class. Everything from here needs Bootstrap's script, and the last two need a little of ours.\n\nSet it up like this: last week we built things that sit there. Today we build things that do something when you click them.");

dividerSlide("Java is to JavaScript as ham is to hamster.", "An old web development joke, and the most useful thing you will hear today.",
  "Say it, then pause and let them read it.\n\nThis is the single most common mix-up for new developers, and it is not their fault: the names really are that confusing. Ask who has heard of Java (most will have). Ask whether they think JavaScript is the browser version of it. Then the next slide.\n\nIf anyone asks whether they need to know Java for this course: no. There is no Java anywhere in this program's web stream.");

{
  const s = plainSlide("Two Different Languages",
    "Go through the two columns, then the timeline at the bottom.\n\nThe story: in 1995 Netscape wanted a small scripting language in their browser. Brendan Eich wrote the first version in about ten days. It was called Mocha, then LiveScript. In December 1995, Netscape and Sun Microsystems (who owned Java) announced a deal, and LiveScript was renamed JavaScript. Java was the hottest thing in software that year, and the name rode on that. That is the whole reason. It is a marketing decision from 1995 that we are still explaining thirty years later.\n\nThe name is also why the standard is called ECMAScript: Sun owned the Java trademark, so the standards body could not use it.\n\nWhat matters for today: JavaScript is the language the browser speaks. HTML is the content, CSS is the look, JavaScript is the behaviour. You have done two of the three. Today is a toe in the water of the third.\n\nDon't oversell it. Four lines today, and we read every word.");
  const cw = 5.8;
  [["Java", C.teal, [
    "Made at Sun Microsystems, 1995.",
    "Compiled, and runs almost anywhere: servers, Android apps, big business systems.",
    "You will meet it in PROG courses, not in a browser.",
  ]], ["JavaScript", C.purple, [
    "Made at Netscape, 1995. First version written in about ten days.",
    "Runs in the browser, on the page you are looking at right now.",
    "The third language of the web: HTML, CSS, JavaScript.",
  ]]].forEach(([t, c, items], i) => {
    const x = 0.75 + i * (cw + 0.35);
    card(s, { x, y: 1.4, w: cw, h: 3.5 });
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.4, w: cw, h: 0.62, fill: { color: c }, line: { color: c } });
    txt(s, t, { x: x + 0.3, y: 1.4, w: cw - 0.6, h: 0.62, fontSize: 22, bold: true, color: C.white, valign: "middle" });
    richBullets(s, items.map((it) => [it]), { x: x + 0.3, y: 2.2, w: cw - 0.6, h: 2.55, size: 14.5, space: 12 });
  });
  card(s, { x: 0.75, y: 5.15, w: W - 1.5, h: 1.75, fill: "F7F2F8", border: "E2CCE0" });
  txt(s, "So why the name?", { x: 1.05, y: 5.3, w: 3.4, h: 0.4, fontSize: 17, bold: true, color: C.purple });
  const tx = 1.05, tw = W - 2.1;
  [["Mocha", "the first name, 1995"], ["LiveScript", "the second name"], ["JavaScript", "renamed in a deal with Sun"]].forEach(([n, d], i) => {
    const x = tx + i * (tw / 3);
    numCircle(s, i + 1, x, 5.85, C.purple, 0.44);
    txt(s, n, { x: x + 0.58, y: 5.82, w: tw / 3 - 0.75, h: 0.32, fontSize: 15.5, bold: true, color: C.black });
    txt(s, d, { x: x + 0.58, y: 6.14, w: tw / 3 - 0.75, h: 0.45, fontSize: 12.5, color: C.grey, valign: "top" });
    if (i < 2) txt(s, "→", { x: x + tw / 3 - 0.5, y: 5.8, w: 0.4, h: 0.5, fontSize: 20, bold: true, color: C.purple, align: "center" });
  });
  txt(s, "Java was the hottest language of 1995. The name was marketing, and we have been explaining it ever since.",
    { x: 1.05, y: 6.52, w: tw, h: 0.35, fontSize: 13, italic: true, color: C.grey });
}

quiz({
  trivia: true,
  q: "Why is JavaScript called JavaScript?",
  options: [
    "It is the browser version of Java",
    "Java and JavaScript were written by the same team",
    "It was renamed as part of a 1995 marketing deal with Sun, who owned Java",
    "JavaScript gets turned into Java before it runs",
  ], answer: 2,
  why: ["It was called Mocha, then LiveScript. The name changed in December 1995, when Netscape and Sun announced a deal.",
    "The two languages have almost nothing to do with each other. Ham and hamster."],
  notes: "Trivia, and a free mark on the quiz for anyone who was listening. Expect a lot of votes for A: that is the whole point of the joke.",
  revealNotes: "Answer: C.\n\nIf someone asks: the standard is called ECMAScript for the same reason. Sun owned the Java trademark, so the standards body (Ecma International) could not use the name.",
});

ladderSlide("How Much JavaScript Does Today Need?", [
  ["None at all", ["Put two attributes on a button. Bootstrap's script is already listening for the click."],
    "Modal · Collapse · Accordion · Carousel · Offcanvas · Dropdown · the navbar's menu button",
    ['<button data-bs-toggle="modal"', '        data-bs-target="#orderModal">', "   Order", "</button>"], C.teal, "html"],
  ["One line", ["Some components are ", ["opt in", "b"], ": you tell Bootstrap which elements to set up."],
    "Tooltip · Popover",
    ["new bootstrap.Tooltip(el);"], C.orange, "js"],
  ["Three lines", ["When something happens, do something. This is what JavaScript ", ["is", "i"], "."],
    "Toast · and everything you write from next week on",
    ['button.addEventListener("click", function () {', "   toast.show();", "});"], C.purple, "js"],
], "This slide is the map for the rest of the class, and the thing to point back at every time someone gets lost.\n\nLevel 1 is most of Bootstrap. The script is loaded, it is watching the whole page, and when you click something with data-bs-toggle it goes looking for the matching data-bs-target. You write HTML. That is all.\n\nLevel 2 exists for one reason: speed. A page could have hundreds of elements with a title. Setting up a tooltip on every one of them would be slow, so Bootstrap makes you ask for the ones you want.\n\nLevel 3 is the real thing, and it is four English words: when this, do that. Say that phrase a few times today. Next week it is 'when the form is submitted, check the email address'.\n\nNobody has to memorize any of this. They have to recognize it.");

jsComponentSlide("The Ones That Move: Which Are Worth Your Time?", [
  ["Modal", "A box over the page. The one everybody wants first.", "modal", "modal", true],
  ["Collapse", "A button that shows and hides a section.", "collapse", "collapse", true],
  ["Accordion", "Stacked sections that open one at a time. Great for FAQs.", "accordion", "collapse", true],
  ["Carousel", "A slideshow with arrows and dots.", "carousel", "carousel", true],
  ["Offcanvas", "A panel that slides in from the side. A good phone menu.", "offcanvas", "offcanvas", true],
  ["Toasts", "A small notice in the corner: \"saved\", \"added to order\".", "toasts", "toast", true],
  ["Tooltips", "A small hint on hover. One line of JS turns them on.", "tooltips", "tooltip", true],
  ["Dropdowns", "A button or link that opens a list. You used one last week.", "dropdowns", "dropdowns", false],
  ["Popovers", "A tooltip with a heading, that stays open. Same one line.", "popovers", "popover", false],
  ["Navs and tabs", "Tabbed panels: one section visible at a time.", "navs-tabs", "navs", false],
  ["Alerts", "Add a close button and it dismisses itself.", "alerts", "alerts", false],
  ["Scrollspy", "Lights up the menu link for the part you're reading.", "scrollspy", "scrollspy", false],
], [["The six in purple are today. ", "b"], "The rest work the same way, so once you can read one you can read all of them. ", ["Click a name to open its docs page.", "i"]],
  "This is Week 10's 'The Ones That Move' slide again, now with a plan attached.\n\nWhy these six: modal, collapse/accordion, carousel and offcanvas are the four you can build with no JavaScript at all, and they cover most of what a small business site actually needs. Tooltip and toast are here because they are the gentlest possible introduction to writing JavaScript: one line, then three.\n\nWhy not the others today: dropdown they already have (the Harbourside navbar next class). Popover is a tooltip with a hat on. Tabs, alerts and scrollspy are all copy-the-docs jobs once you know the pattern.\n\nFor Assignment 4: any one of these is an easy way to make a page feel finished. A modal on a 'Book a table' button, or an accordion of FAQs, takes about five minutes.\n\nEvery name on this slide is a link. Open one or two live so they see the docs page has a copy-paste example at the top.");

// ---------------------------------------------------------------- modal
codeSlide({
  kicker: "NO JAVASCRIPT · EXAMPLE 01", title: "A Modal", file: "11_Customizing_Bootstrap_Tailwind/js/01_modal_example.html",
  lines: [
    '<!-- 1. The button: WHAT to open, and WHICH one -->',
    '<button type="button" class="btn btn-primary btn-lg"',
    '   data-bs-toggle="modal" data-bs-target="#orderModal">',
    '   Order the lobster roll',
    '</button>',
    '',
    '<!-- 2. The box. Its id matches the target. -->',
    '<div class="modal fade" id="orderModal" tabindex="-1">',
    '   <div class="modal-dialog">',
    '      <div class="modal-content">',
    '         <div class="modal-header">',
    '            <h2 class="modal-title h5">Lobster Roll</h2>',
    '            <button class="btn-close"',
    '               data-bs-dismiss="modal"></button>',
    '         </div>',
    '         <div class="modal-body">...</div>',
    '      </div>',
    '   </div>',
    '</div>',
  ],
  lang: "html", result: SH.modal, size: 14, mark: [2, 7],
  bullets: [
    [["data-bs-toggle", "c"], " = what. ", ["data-bs-target", "c"], " = which one, by ", ["id", "c"], "."],
    ["The ", ["#", "c"], " is there because it is a selector, same as CSS."],
    [["data-bs-dismiss", "c"], " closes whatever it is inside."],
  ],
  notes: "Open js/01_modal_example.html and click the button. Then close it three ways: the X, the dark area outside, and the Esc key. All three are free.\n\nWalk the two orange marks: the target on the button, the id on the box. Say it plainly - these two have to match, character for character, and the # belongs on the target only.\n\nThe modal div sits in the HTML the whole time. Bootstrap's CSS has .modal { display: none }, and the script adds a class to show it. Prove it: open DevTools, find the div, click the button, and watch the class attribute change.\n\nmodal-dialog, modal-content, modal-header, modal-body, modal-footer: five nested divs, which looks like a lot. Nobody types this from memory. Copy it from the docs.\n\ntabindex=\"-1\" is so the Esc key works. Leave it in.\n\nWhat if two modals have the same id? The first one wins, every time. Ids are unique on a page.",
});

trySlide({
  n: 1, title: "Open It, Then Break It", time: "5 min", lw: 5.6, codeTitle: "THE TWO LINES THAT HAVE TO MATCH", lang: "html",
  steps: [
    ["Open the example", [["js/01_modal_example.html", "c"], " with Go Live. Click the button."]],
    ["Close it three ways", "The X, then the dark area, then the Esc key."],
    ["Change the words", "The title, the body, and both footer buttons. Save and reload."],
    ["Break it on purpose", ["Change ", ["#orderModal", "c"], " to ", ["#orderModel", "c"], " on the button. Click. Nothing."]],
    ["Fix it", "Then add a second modal, with its own id and its own button."],
  ],
  code: [
    '<button data-bs-toggle="modal"',
    '        data-bs-target="#orderModal">',
    '',
    '<div class="modal fade" id="orderModal">',
    '',
    '',
    '<!-- A second one needs a NEW id: -->',
    '<button data-bs-toggle="modal"',
    '        data-bs-target="#hoursModal">',
    '',
    '<div class="modal fade" id="hoursModal">',
  ],
  tip: [[["Step 4 is the lesson: ", "b"], "a typo in an id does not give you an error. The button just stops working."]],
  notes: "Walk around. Step 4 is the one that matters: everyone should see, once, that a broken Bootstrap component is completely silent. No red text, no warning, nothing in the Console.\n\nAsk afterwards: how would you find that bug in your own page? Answer: check the target against the id, then check the script tag. In that order, every time.\n\nFast finishers: give the modal a bigger size with modal-lg on the modal-dialog div, or make it slide down from the top with modal-dialog-centered. Both are one class, both are in the docs.\n\nCommon problems: they change the id on the box but not on the button (or the other way round); they paste the second modal inside the first one's div.",
});

quiz({
  q: "What goes on the button to open the box below it?",
  code: ['<button type="button" class="btn btn-primary">', "   Order the lobster roll", "</button>", "", '<div class="modal fade" id="orderModal">', "   ...", "</div>"], lang: "html", codeTitle: "ONE BUTTON, ONE MODAL",
  options: ['`data-bs-toggle="modal" data-bs-target="#orderModal"`', '`data-toggle="modal" data-target="#orderModal"`', '`data-bs-toggle="#orderModal"`', '`onclick="openModal()"`'], answer: 0,
  why: ["`data-bs-toggle` is what kind of thing to open; `data-bs-target` is which one, by its id, with the `#`.",
    "B is Bootstrap 4 and does nothing in Bootstrap 5 - that was Week 10's Pop Quiz 7. D would need JavaScript you write yourself."],
  notes: "Pop Quiz 3. B is the trap, and it is the same trap as last week's navbar question, so it is worth seeing twice.",
  revealNotes: "Answer: A.\n\nIf anyone picked B: Bootstrap 5's script only ever looks for data-bs-. Bootstrap 4 code found online is the single most common reason copied snippets do nothing.\n\nC is wrong because toggle takes the kind of component, not a selector.",
});

// ---------------------------------------------------------------- collapse and accordion
codeSlide({
  kicker: "NO JAVASCRIPT · EXAMPLE 02", title: "Collapse, and an Accordion", file: "11_Customizing_Bootstrap_Tailwind/js/02_accordion_example.html",
  lines: [
    '<!-- One collapse: the same two attributes again -->',
    '<button class="btn btn-outline-primary"',
    '   data-bs-toggle="collapse" data-bs-target="#allergens">',
    '   Allergen information',
    '</button>',
    '',
    '<div class="collapse" id="allergens">',
    '   <div class="card card-body">Our fryers are shared...</div>',
    '</div>',
    '',
    '',
    '<!-- An accordion: collapses that know about each other -->',
    '<div class="accordion" id="faq">',
    '   <div class="accordion-item">',
    '      <h3 class="accordion-header">',
    '         <button class="accordion-button" type="button"',
    '            data-bs-toggle="collapse" data-bs-target="#q1"',
    '            aria-expanded="true" aria-controls="q1">',
    '            Do you take reservations?',
    '         </button>',
    '      </h3>',
    '      <div class="accordion-collapse collapse show"',
    '         id="q1" data-bs-parent="#faq">',
    '         <div class="accordion-body">Not in the summer...</div>',
    '      </div>',
    '   </div>',
    '   <!-- ...and two more items, the same way -->',
    '</div>',
  ],
  lang: "html", result: SH.accordion, size: 13.5, mark: [22],
  bullets: [
    ["Same pattern, new word: ", ["data-bs-toggle=\"collapse\"", "c"], "."],
    [["data-bs-parent=\"#faq\"", "c"], " is what makes it an ", ["accordion", "b"], ": open one, the open one closes."],
    [["show", "c"], " on an answer means it starts open."],
  ],
  notes: "Open js/02_accordion_example.html. Click the allergen button first: one collapse on its own, nothing fancy. Then say the useful thing - this is exactly what the navbar's menu button does under 768px. They have been using a collapse since last week without knowing its name.\n\nThen the accordion. Open each question and let them watch the previous one close.\n\nThe orange mark is the whole lesson: data-bs-parent. Take it away and you have three independent collapses that all stay open. That is Try It 2, step 3, and it is worth doing live if there is time.\n\nThe heading structure matters for accessibility: the button lives inside a real heading (h3 here), so a screen reader can jump between questions. Bootstrap's docs use h2; use whatever level fits your page.\n\naria-expanded and aria-controls are on the buttons (slide and file). aria-controls names the answer it opens, so it changes along with data-bs-target. Bootstrap updates aria-expanded as you click, so a screen reader announces 'expanded' or 'collapsed'. Point at it, don't dwell.\n\nFAQ accordions are everywhere. If anyone is short of ideas for Assignment 4, this is a five-minute win.",
});

trySlide({
  n: 2, title: "Add a Question, Then Remove the Parent", time: "5 min", lw: 5.6, codeTitle: "ONE ACCORDION ITEM", lang: "html",
  steps: [
    ["Open the example", [["js/02_accordion_example.html", "c"], ". Open all three questions."]],
    ["Add a fourth", ["Copy an ", ["accordion-item", "c"], ". Change ", ["q3", "c"], " to ", ["q4", "c"], " in ", ["all three", "b"], " places."]],
    ["Take the parent out", ["Delete ", ["data-bs-parent=\"#faq\"", "c"], " from all four. Reload. What changed?"]],
    ["Put it back", "One attribute is the difference between a list and an accordion."],
  ],
  code: [
    '<div class="accordion-item">',
    '   <h3 class="accordion-header">',
    '      <button class="accordion-button collapsed"',
    '         type="button" data-bs-toggle="collapse"',
    '         data-bs-target="#q4"',
    '         aria-expanded="false" aria-controls="q4">',
    '         Do you deliver?',
    '      </button>',
    '   </h3>',
    '   <div class="accordion-collapse collapse"',
    '      id="q4" data-bs-parent="#faq">',
    '      <div class="accordion-body">',
    '         Not yet. Takeout only.',
    '      </div>',
    '   </div>',
    '</div>',
  ],
  tip: [[["Three places, every time: ", "b"], "the button's ", ["data-bs-target", "c"], " and ", ["aria-controls", "c"], ", and the ", ["id", "c"], " on the answer below it."]],
  notes: "Step 3 is the point of the Try It. Ask for the answer out loud before they put it back: with no parent, every answer stays open, because nothing tells one collapse about the others.\n\nThe copied item needs collapsed on the button and no show on the answer, or it starts open and the arrow points the wrong way. That is the most common thing to fix while walking around.\n\nFast finishers: give the accordion flush (no outer border and rounded corners), or make the whole thing start closed by taking show off the first answer and adding collapsed to its button.",
});

// ---------------------------------------------------------------- carousel
codeSlide({
  kicker: "NO JAVASCRIPT · EXAMPLE 03", title: "A Carousel", file: "11_Customizing_Bootstrap_Tailwind/js/03_carousel_example.html",
  lines: [
    '<div id="specials" class="carousel slide"',
    '   data-bs-ride="carousel">',
    '',
    '   <div class="carousel-inner">',
    '      <div class="carousel-item active" data-bs-interval="4000">',
    '         <img src="images/slide1.svg" class="d-block w-100" alt="...">',
    '         <div class="carousel-caption d-none d-md-block">',
    '            <h2 class="h4">Haddock &amp; Chips</h2>',
    '         </div>',
    '      </div>',
    '      <!-- ...two more carousel-items, without "active" -->',
    '   </div>',
    '',
    '   <button class="carousel-control-next" type="button"',
    '      data-bs-target="#specials" data-bs-slide="next">',
    '      <span class="carousel-control-next-icon"></span>',
    '   </button>',
    '</div>',
  ],
  lang: "html", result: SH.carousel, resultLabel: "SLIDE 1 OF 3", size: 13.5, mark: [4],
  bullets: [
    [["Exactly one", "b"], " slide gets ", ["active", "c"], ". Two, or none, and nothing shows."],
    [["data-bs-slide", "c"], " on the arrows, ", ["data-bs-slide-to", "c"], " on the dots."],
    [["data-bs-interval=\"4000\"", "c"], ": four seconds, in milliseconds."],
  ],
  notes: "Open js/03_carousel_example.html and let it run for a few seconds while you talk.\n\nThe orange mark: active. This is the one thing people get wrong. Copy a slide, forget to take active off the copy, and you get two 'current' slides and a carousel that sits there jittering. Delete active from all of them and you get a blank box.\n\nThree new attributes, all still data-bs-: ride (start it moving), interval (how long, in milliseconds - 4000 is four seconds), slide and slide-to (the arrows and the dots).\n\nThe honest bit, and it is worth saying: most real sites should not auto-play a carousel. People do not read them, they move away before you finish a sentence, and they are hard work for anyone using a screen reader or who finds motion difficult. Take data-bs-ride off and let people click. Bootstrap's own docs say the same thing.\n\nOur three slides are SVG files in js/images. If they swap in photos, all three have to be the same size or the page jumps as it moves.\n\nThe captions use d-none d-md-block, from Week 10: hidden under 768px, because they do not fit on a phone.",
});

trySlide({
  n: 3, title: "Make the Carousel Behave", time: "5 min", lw: 5.6, codeTitle: "WHAT TO CHANGE", lang: "html",
  steps: [
    ["Open the example", [["js/03_carousel_example.html", "c"], ". Watch it go round once."]],
    ["Make it painful", ["Set every ", ["data-bs-interval", "c"], " to ", ["1000", "c"], ". Reload. Could you read that?"]],
    ["Make it polite", ["Set them to ", ["8000", "c"], ", or delete ", ["data-bs-ride", "c"], " so it only moves when clicked."]],
    ["Break it, then fix it", ["Put ", ["active", "c"], " on two slides. Reload. Now take it off one."]],
  ],
  code: [
    '<!-- Step 2 and 3: on every carousel-item -->',
    'data-bs-interval="1000"   <!-- one second -->',
    'data-bs-interval="8000"   <!-- eight seconds -->',
    '',
    '<!-- Step 3: no autoplay at all -->',
    '<div id="specials" class="carousel slide">',
    '',
    '',
    '<!-- Step 4: two actives = broken -->',
    '<div class="carousel-item active">',
    '<div class="carousel-item active">',
  ],
  tip: [[["The accessibility answer: ", "b"], "a carousel nobody can read is a carousel nobody reads. Let people click."]],
  notes: "This Try It is the first one you cut if you are behind; the carousel is the least useful of today's six.\n\nStep 2 is deliberately annoying, and the annoyance is the lesson. Ask the room after step 3 which version they would put on a real site.\n\nStep 4: with two actives the carousel usually shows both stacked or refuses to move. With none it shows an empty box. Either way, 'exactly one active' is the thing to remember.\n\nFast finishers: add a fourth slide. They need a fourth carousel-item (no active) and a fourth dot with data-bs-slide-to=\"3\" - the dots count from 0.",
});

// ---------------------------------------------------------------- offcanvas and the pattern
{
  const s = plainSlide("Offcanvas, and the Pattern Behind All of It",
    "Open js/04_offcanvas_example.html and click Menu. A panel slides in from the left, the page behind it dims, and Esc closes it. You have seen all of that before - it is a modal that comes from the side.\n\nThe point of this slide is the box in the middle. Three different components, three different looks, one pattern: data-bs-toggle says what kind, data-bs-target says which one. Learn that once and the docs become copy-and-paste.\n\nWhere offcanvas earns its place: as a phone menu. A navbar whose toggler opens an offcanvas instead of a collapse, so the menu slides over the page instead of pushing everything down. Bootstrap's docs call it 'Navbar with offcanvas', and Example 07 has one.\n\nThe variants: offcanvas-start (left), -end (right), -top, -bottom.\n\nIf you are short of time, this is the second slide to cut. The pattern box is the only thing on it they must keep.");
  codeBox(s, [
    'data-bs-toggle="modal"       data-bs-target="#orderModal"',
    'data-bs-toggle="collapse"    data-bs-target="#allergens"',
    'data-bs-toggle="offcanvas"   data-bs-target="#sideMenu"',
    'data-bs-toggle="dropdown"',
    'data-bs-toggle="tooltip"',
  ], { x: 0.75, y: 1.4, w: 7.0, h: 2.0, maxSize: 15, fitHeight: true, lang: "html", title: "THE SAME TWO ATTRIBUTES, EVERY TIME" });
  card(s, { x: 0.75, y: 3.7, w: 7.0, h: 1.35, fill: "FBE9DF", border: "F2C6AC" });
  richBullets(s, [[["What", "b"], " to open, and ", ["which one", "b"], ", by its ", ["id", "c"], ". Only the word after ", ["toggle", "c"], " changes."],
    ["Learn the pattern once, and every component in the docs is copy, paste, change the words."]],
    { x: 1.05, y: 3.75, w: 6.4, h: 1.25, size: 14.5, bullet: false, valign: "middle" });
  [["offcanvas-start", "slides in from the left"], ["offcanvas-end", "from the right"], ["offcanvas-top", "from the top (and -bottom)"]].forEach(([n, d], i) => {
    const y = 5.3 + i * 0.55;
    pill(s, n, 0.75, y, [C.purple, C.teal, C.orange][i], 2.0);
    txt(s, d, { x: 2.95, y: y - 0.04, w: 4.8, h: 0.42, fontSize: 14, color: C.dark, valign: "middle" });
  });
  framedImage(s, SH.offcanvas, 8.15, 1.75, W - 8.15 - 0.75, "EXAMPLE 04 · THE PANEL OPEN");
  txt(s, "11_Customizing_Bootstrap_Tailwind/js/04_offcanvas_example.html", { x: 8.15, y: 6.6, w: W - 8.15 - 0.75, h: 0.3, fontSize: 10.5, color: C.mid, fontFace: MONO });
}

quiz({
  q: "You delete Bootstrap's `<script>` tag from a page that was working. What happens?",
  options: [
    "Everything breaks: the page loses its colours and layout too",
    "The modal, carousel and menu button stop. The grid, cards and colours are fine",
    "Nothing: the `data-bs-` attributes do the work on their own",
    "You get a red error in the Console telling you what is wrong",
  ], answer: 1,
  why: ["Bootstrap is two downloads: a CSS file for how things look, and a script for the things that move.",
    ["And there is no error. Nothing was loaded, so nothing complained. ",
      ["A button that does nothing, with an empty Console, is almost always a missing script.", "b"]]],
  revealImg: SH.nojs,
  notes: "Pop Quiz 4, and the most practically useful question of the day. D catches people who assume the browser will tell them.\n\nThe picture on the answer slide is js/00_no_script_example.html, which is Example 01 with the script line deleted. Open it live if you can: the buttons at the top do nothing, the cards at the bottom look perfect.",
  revealNotes: "Answer: B.\n\nOpen js/00_no_script_example.html and click the two buttons at the top. Nothing. Then scroll down: the card, the alert and the coloured grid are all fine, because those are pure CSS.\n\nThe missing line, and it goes just before </body>:\nhttps://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js\n\nWhy 'bundle'? That file is Bootstrap's own JavaScript plus Popper, the little library that works out where a dropdown or tooltip should sit so it doesn't fall off the screen. Take the bundle and you never have to think about it.",
});

// ---------------------------------------------------------------- actually writing some
dividerSlide("Now Some Actual JavaScript", "Four lines   ·   read word by word   ·   nobody is expected to know this yet",
  "Change of gear. Everything so far was HTML attributes. The next two components need code.\n\nSay it again before you start: you have not learned JavaScript. You are not expected to write this from memory. You are expected to be able to read it and change the words in it. That is a completely reasonable thing to ask of you today.");

anatomySlide("Reading a Line of JavaScript",
  ['document.getElementById("orderButton")', '   .addEventListener("click", function () {', '      toast.show();', "   });"],
  [
    ["document", [["The page itself. Everything in browser JavaScript starts from here.", ""]]],
    [".getElementById(\"orderButton\")", [["Find the one element with ", ""], ["id=\"orderButton\"", "c"], [". Same idea as ", ""], ["#orderButton", "c"], [" in CSS - but with ", ""], ["no #", "b"], [".", ""]]],
    [".addEventListener(\"click\", ...)", [["\"When this gets clicked, run that.\" The code does ", ""], ["not", "bi"], [" run now. It waits.", ""]]],
    ["function () { ... }", [["The \"that\": the code to run when the click happens. Everything between the braces.", ""]]],
    ["toast.show()", [["Show the toast. The round brackets mean ", ""], ["do it", "b"], [", right now.", ""]]],
  ],
  "Put this on the board and leave it up. Read it out as English, left to right, top to bottom:\n\n'In the page, find the element with id orderButton. When it is clicked, run this: show the toast.'\n\nFour things to land:\n\n1. document is the page. Everything in browser JavaScript starts from there.\n\n2. getElementById takes the id WITHOUT the hash. The hash is a CSS thing. This catches everyone once.\n\n3. addEventListener is the heart of it. The code inside does not run when the page loads; it sits there waiting. If a student only takes one idea out of today, make it this one, because next week's form validation is 'when the form is submitted, check it'.\n\n4. The round brackets on show() mean do it now. toast.show without brackets does nothing at all, silently.\n\nAsk them to say the whole thing in plain English before you move on. Someone will say 'when the button is clicked, show the toast', and that is the win.");

codeSlide({
  kicker: "ONE LINE · EXAMPLE 05", title: "Tooltips: Opt In", file: "11_Customizing_Bootstrap_Tailwind/js/05_tooltip_example.html",
  lines: [
    '<button type="button" class="btn btn-primary"',
    '   data-bs-toggle="tooltip"',
    '   data-bs-title="Comes with fries and coleslaw">',
    '   Haddock &amp; Chips',
    '</button>',
    '',
    '',
    "<!-- and at the end of the page, after Bootstrap's script: -->",
    '<script>',
    '   document.querySelectorAll(\'[data-bs-toggle="tooltip"]\').forEach(function (el) {',
    '      new bootstrap.Tooltip(el);',
    '   });',
    '</script>',
  ],
  lang: "html", result: SH.tooltip, size: 14, mark: [9, 10],
  bullets: [
    ["The attribute is not enough. Tooltips are ", ["opt in", "b"], "."],
    [["querySelectorAll", "c"], " takes a ", ["CSS selector", "b"], ": find them all."],
    [["forEach", "c"], ": do this once for each one you found."],
  ],
  notes: "Open js/05_tooltip_example.html and hover a button. Then comment the script out live (Ctrl + /) and reload: the tooltips are gone, and nothing else changes.\n\nWhy opt in? Performance. A real page might have hundreds of elements with a title. Setting up a tooltip object for every one of them on every page load would be slow, so Bootstrap makes you say which.\n\nRead the two marked lines as English: 'find everything on the page with that attribute; for each one, make a tooltip out of it.'\n\nThe selector in the brackets is a CSS attribute selector from Week 6. They have written these before, in a stylesheet. Same selector, different place.\n\nnew bootstrap.Tooltip(el) - 'bootstrap' here is a name the bundle script puts on the page for us. Don't go further than that today.\n\ndata-bs-placement=\"bottom\" or \"left\" or \"right\" moves it. data-bs-title is the text; in older code you will see plain title instead, which also works.\n\nPopovers are the same two steps, with Popover instead of Tooltip.",
});

trySlide({
  n: 4, title: "Your First Line of JavaScript", time: "5 min", lw: 5.6, codeTitle: "THE LINE THAT TURNS THEM ON", lang: "js",
  steps: [
    ["Open the example", [["js/05_tooltip_example.html", "c"], ". Hover each of the three buttons."]],
    ["Switch it off", ["Select the script and press ", ["Ctrl + /", "b"], " to comment it out. Reload: no tooltips."]],
    ["Switch it back on", ["Same keys. They are back. ", ["You just made JavaScript run.", "b"]]],
    ["Add your own", ["Put the two attributes on something else on the page."]],
    ["Move it", [["data-bs-placement=\"left\"", "c"], ", or ", ["\"right\"", "c"], ", or ", ["\"bottom\"", "c"], "."]],
  ],
  code: upTo(JS05, "// Find every tooltip", "});"),
  tip: [[["Nothing happening? ", "b"], "Press ", ["F12", "b"], " and look at the Console. Unlike HTML, a typo in JavaScript turns up there in red."]],
  notes: "Steps 2 and 3 are the whole point: they turn a piece of JavaScript off and on and watch the page change. That is a first, and it is worth naming out loud.\n\nStep 4: the easiest target is the h1 or a word in a paragraph wrapped in a span. Remind them the element needs both data-bs-toggle=\"tooltip\" and data-bs-title=\"...\".\n\nThe Console is the new tool today. Show it once on the projector: F12, Console tab. Make a deliberate typo (bootstrap.Tooltipp) and read the red line together. The message is ugly but the file and line number are gold.\n\nFast finishers: make one a popover instead. data-bs-toggle=\"popover\", add data-bs-content=\"...\", and a second line with new bootstrap.Popover(el). It opens on click instead of hover.",
});

quiz({
  q: "Your button has `data-bs-toggle=\"tooltip\"`, Bootstrap's script is loaded, and the page has no errors. You hover. Nothing. Why?",
  options: [
    "Tooltips only work on links, not on buttons",
    "You have to use `title`, not `data-bs-title`",
    "Tooltips are opt in: one line of JavaScript has to turn them on",
    "Tooltips need their own separate script file",
  ], answer: 2,
  why: ["Most components work from the attributes alone. Tooltips and popovers are the exception, for speed: a page could have hundreds of them.",
    "The line: find every element with that attribute, and make a `new bootstrap.Tooltip` out of each one."],
  notes: "Pop Quiz 5. Ask for a show of hands on who got caught by this in Try It 4 before you reveal.",
  revealNotes: "Answer: C.\n\nThis is the one Bootstrap behaviour that genuinely surprises people, because every other component just works.\n\nB is a half-truth worth clearing up: plain title does work, and a lot of older code uses it. But data-bs-title is what Bootstrap 5.3's docs use now, and it avoids the browser showing its own grey tooltip as well.",
});

codeSlide({
  kicker: "THREE LINES · EXAMPLE 06", title: "A Toast: When This, Do That", file: "11_Customizing_Bootstrap_Tailwind/js/06_toast_example.html",
  lines: upTo(JS06, "// 1. Get the toast ready", "});"),
  lang: "js", result: SH.toast, size: 14, codeTitle: "AT THE END OF THE PAGE",
  bullets: [
    ["There is no ", ["data-bs-toggle=\"toast\"", "c"], ". A toast only appears when ", ["we", "b"], " show it."],
    ["Line 1 gets it ready. It does ", ["not", "bi"], " show it."],
    ["Then: ", ["when", "b"], " the button is clicked, ", ["show", "b"], " the toast."],
  ],
  notes: "Open js/06_toast_example.html and click the button. Do it two or three times; it is satisfying, and that matters at this point in the class.\n\nThis is the payoff slide. Read the code as the sentence from the anatomy slide: 'find the button; when it is clicked, show the toast.'\n\nLine 1 deserves a beat. new bootstrap.Toast(...) builds the toast and hands it back; we keep it in a box called toast so we can use it on line 3. const means 'a name for something that will not change'. It does NOT show anything. That is Pop Quiz 6.\n\nWhy no data-bs-toggle for toasts? Because a toast almost never responds to a click on a button next to it. In a real app it fires when a form saves, or a file uploads, or a message arrives. There is no attribute for 'when the save worked'.\n\nThe toast-container with position-fixed bottom-0 end-0 p-3 is what pins it to the bottom-right corner. Those are Week 10 utility classes, nothing new.\n\nIt hides itself after five seconds: data-bs-autohide=\"false\" stops that, data-bs-delay=\"10000\" makes it ten.\n\nFinish on this: four English words. When this, do that. Next week it is 'when the form is submitted, check the email address'. Same shape.",
});

trySlide({
  n: 5, title: "When This, Do That", time: "6 min", lw: 5.6, codeTitle: "THE THREE LINES", lang: "js",
  steps: [
    ["Open the example", [["js/06_toast_example.html", "c"], ". Click the button a few times."]],
    ["Change the message", ["The words in ", ["toast-body", "c"], ", and the name in ", ["toast-header", "c"], "."]],
    ["Break it", ["Comment out ", ["toast.show();", "c"], ". Click. Nothing, and no error."]],
    ["Fix it, then make it stay", ["Add ", ["data-bs-autohide=\"false\"", "c"], " to the toast."]],
    ["Change the trigger", ["Point ", ["getElementById", "c"], " at a different button on the page."]],
  ],
  code: upTo(JS06, "// 1. Get the toast ready", "});"),
  tip: [[["Case matters: ", "b"], [" getElementById(\"orderButton\")", "c"], " will not find ", ["id=\"orderbutton\"", "c"], "."]],
  notes: "The last Try It, and the best one. Walk around and ask people to read their own line out as a sentence.\n\nStep 3 is the setup for Pop Quiz 6, so make sure everyone actually does it.\n\nStep 5 is the real test of understanding: they have to change the id in the JavaScript to match an id in the HTML. If they can do that, they have understood the idea that the two files point at each other.\n\nCommon problems: capital letters in the id; the script placed above the HTML it talks about (ours is at the end of the body for exactly that reason); a missing semicolon, which usually still works and is worth not making a fuss about today.\n\nFast finishers: make a second button show the same toast, or change the toast's colours with text-bg-success on the toast div.",
});

quiz({
  q: "The button is clicked. What happens?",
  code: ['const toast = new bootstrap.Toast(', '   document.getElementById("orderToast"));', "", 'document.getElementById("orderButton")', '   .addEventListener("click", function () {', '      console.log("clicked");', "   });"], lang: "js", codeTitle: "SPOT WHAT IS MISSING",
  options: ["The toast appears", "Nothing at all happens", "\"clicked\" shows up in the Console, but the toast never appears", "The toast appears once, then stops working"], answer: 2,
  why: [["Line 1 gets the toast ", ["ready", "bi"], ". It does not show it. A toast stays hidden until something calls ", ["toast.show()", "c"], "."],
    "The listener works fine, so the Console message appears. Add `toast.show();` inside the braces and it works."],
  notes: "Pop Quiz 6, the last one today. Everyone who did Try It 5 step 3 has already seen this.\n\nconsole.log is new on this slide. Explain it in one line: it writes a message into the Console so you can see your code ran. It is the single most useful debugging tool there is, and they will use it constantly from next week.",
  revealNotes: "Answer: C.\n\nThe lesson underneath: 'my code ran but nothing happened' and 'my code never ran' are two different bugs, and console.log is how you tell them apart.\n\nThat is also why the empty Console in Pop Quiz 4 was such a good clue.",
});

{
  const s = activitySlide("Everything Today, On One Page",
    "js/07_all_together_example.html. Open it on the projector and click around: the carousel, the allergen tooltip, Order (a modal), Yes add it (which closes the modal and pops a toast), the FAQ accordion. Then make the window narrow and open the phone menu, which is an offcanvas.\n\nIts whole script is five lines, and they have seen every one of them today.\n\nIt is also deliberately a Harbourside page with .bg-harbour and .btn-harbour in a <style> block, so it ties back to the customizing half of the class.\n\nThis is the slide to point at when someone asks what to do for Assignment 4. Anything on this page is fair game, and most of it is a copy-paste plus changing the words.\n\nIf you are out of time: skip the demo, tell them the file is there, and move to the recap.");
  const ih = 4.6;
  framedImage(s, SH.allT, 0.75, 1.9, 6.1, "ORDER → MODAL → TOAST");
  const dw = Math.min(3.1, ih / SH.allD.ratio);
  framedImage(s, SH.allD, 7.15, 1.9, dw, "THE WHOLE PAGE");
  phoneFrame(s, SH.allP, 7.15 + dw + 0.4, 1.9, 2.1, "PHONE: OFFCANVAS MENU", C.green);
  card(s, { x: 0.75, y: 5.3, w: 6.1, h: 1.6, fill: C.white });
  txt(s, "Steal any of this for Assignment 4:", { x: 1.0, y: 5.42, w: 5.6, h: 0.35, fontSize: 15, bold: true, color: C.black });
  richBullets(s, [
    ["A modal on a \"Book a table\" button, an accordion of FAQs, a toast when something is added."],
    ["Five lines of JavaScript in total, and you have read all of them today."],
  ], { x: 1.0, y: 5.78, w: 5.6, h: 1.05, size: 13.5 });
}

{
  const s = panelSlide("Class 1 Recap",
    "If they remember three things: put your CSS after Bootstrap's and set the --bs-btn- variables; data-bs-toggle plus data-bs-target opens almost anything, with no JavaScript from you; and 'when this, do that' is what JavaScript is.\n\nAnd the practical one, which will save them hours: a Bootstrap component that does nothing is silent. Check the target against the id, then check the script tag.\n\nNext class: Tailwind, then the big Harbourside rebuild. The navbar in that activity is a collapse and a dropdown, so today was the groundwork for it.\n\nThe week after: JavaScript properly, starting with custom form validation. Tell them the toast line is the shape of everything that is coming.");
  labelBullets(s, [
    { label: "Customizing:", text: [
      "Your style.css after Bootstrap's. Buttons: set the --bs-btn- variables, so hover follows.",
    ] },
    { label: "No JavaScript needed:", text: [
      "data-bs-toggle = what to open. data-bs-target = which one, by its id.",
      "Modal, collapse, accordion, carousel, offcanvas, dropdown. One pattern, six components.",
      "No script tag = nothing moves, and no error message. Check that first.",
    ] },
    { label: "Your first JavaScript:", text: [
      "Tooltips are opt in: one line turns them on.",
      "Toasts: when the button is clicked, show the toast. Four English words.",
    ] },
    { label: "Next class:", text: ["Tailwind CSS, then Harbourside rebuilt in Bootstrap."] },
  ], { size: 17, sub: 14 });
}

// =====================================================================
// WEEK 11, CLASS 2: Tailwind, and Harbourside in Bootstrap
// =====================================================================
section("Class 2 - Tailwind, and Harbourside in Bootstrap");

dividerSlide("Week 11, Class 2", "Tailwind CSS   ·   Harbourside, rebuilt in Bootstrap   ·   Try Its 1 to 6",
  "Class 2. Two halves: a quick look at the other kind of framework, then the big activity.\n\nThe activity is the priority. If Tailwind runs long, cut the 'Which one?' slide and get started on Harbourside.");

setupSlide("Is Your Environment Set Up?", [
  ["GIT", "Did you run git pull?"],
  ["VS CODE", "Is the Week 11 folder open in VS Code?"],
  ["LIVE SERVER", "Does activity/index.html open with Go Live? (It has no CSS yet.)"],
  ["INTERNET", "Bootstrap and Tailwind both load from the web."],
], "Help anyone who's stuck now, so they aren't lost during the Try Its.\n\nactivity/index.html is the Harbourside page with no CSS at all, so it should look plain: big logo, bullet-point menu, Times New Roman. That's correct, and it is where we start.\n\nThe finished page is in activity/answer/ if anyone falls badly behind.");

recallSlide("Recall: What Did We Do Last Class?", (s) => {
  // The data-bs- pattern, drawn: a button pointing at a box
  const x = 0.75, w = 5.2;
  card(s, { x, y: 1.75, w, h: 1.05, fill: C.white });
  txt(s, "<button", { x: x + 0.25, y: 1.85, w: w - 0.5, h: 0.3, fontSize: 13, fontFace: MONO, color: C.grey });
  txt(s, 'data-bs-toggle="modal"', { x: x + 0.6, y: 2.12, w: w - 0.8, h: 0.3, fontSize: 13.5, fontFace: MONO, bold: true, color: "8A2B0B" });
  txt(s, 'data-bs-target="#order"', { x: x + 0.6, y: 2.4, w: w - 0.8, h: 0.3, fontSize: 13.5, fontFace: MONO, bold: true, color: C.purple });
  s.addShape(pres.shapes.LINE, { x: x + 3.4, y: 2.55, w: 0, h: 0.75, line: { color: C.purple, width: 2.5, endArrowType: "triangle" } });
  card(s, { x, y: 3.35, w, h: 0.95, fill: "F7F2F8", border: "E2CCE0" });
  txt(s, '<div class="modal" id="order">', { x: x + 0.25, y: 3.5, w: w - 0.5, h: 0.3, fontSize: 13.5, fontFace: MONO, bold: true, color: C.purple });
  txt(s, "   the box that opens", { x: x + 0.25, y: 3.8, w: w - 0.5, h: 0.3, fontSize: 13, fontFace: MONO, color: C.grey });
  card(s, { x, y: 4.6, w, h: 0.95, fill: "1E1E1E", border: "1E1E1E" });
  txt(s, "and at the end of the body:", { x: x + 0.25, y: 4.68, w: w - 0.5, h: 0.28, fontSize: 11.5, color: "F2C6AC" });
  txt(s, "<script src=\"...bootstrap.bundle.min.js\">", { x: x + 0.25, y: 4.98, w: w - 0.5, h: 0.4, fontSize: 12, fontFace: MONO, color: "9CDCFE" });
  txt(s, "What opens what, and the script that does it.", { x: x - 0.15, y: 5.75, w: w + 0.3, h: 0.4, fontSize: 13, italic: true, color: C.grey, align: "center" });
}, [
  ["The pattern", [["data-bs-toggle", "c"], " = what kind. ", ["data-bs-target", "c"], " = which one, by ", ["id", "c"], "."]],
  ["Six components", ["Modal, collapse, accordion, carousel, offcanvas, dropdown. No JavaScript from you."]],
  ["The script", ["No ", ["bootstrap.bundle.min.js", "c"], ", nothing moves, ", ["and no error", "b"], "."]],
  ["Your own JavaScript", ["Tooltips need one line. A toast needs ", ["when this, do that", "b"], "."]],
], "30 seconds with a neighbour, then take answers.\n\nThe one to press on is point 3, because today's Try It 2 (the navbar) needs the script, and half the room will forget it.\n\nThen point at the drawing: today's navbar is the same pattern again, with 'collapse' and 'dropdown' instead of 'modal'. They have already done the hard part.");

overviewSlide("Week 11, Class 2: Overview", "Tailwind, and Harbourside", [
  "Tailwind CSS: the same cards, built the other way",
  "Bootstrap, Tailwind, or plain CSS?",
  "Rebuild Harbourside with Bootstrap: Try Its 1 to 6",
  "Make it look like Harbourside, then submit two screenshots",
], [["Compare", "Tailwind"], ["Build", "Harbourside"], ["Submit", "Two screenshots"]],
  "Rough plan:\n- 10 min: check-in, environment, recall, overview\n- 22 min: Tailwind, Examples 02 and 03, Pop Quizzes 7 and 8\n- 50 min: Harbourside, Try Its 1 to 6, Pop Quizzes 9 and 10\n- 8 min: submit, resources, recap, the one-sentence summary\n\nThe activity is the priority. If time is short, demo Tailwind quickly and skip the 'Which one?' slide.\n\nFour pop quizzes today, 7 to 10. With last class's six that is the whole Week 11 quiz.");

dividerSlide("Tailwind CSS", "Utility-first   ·   the same cards, two ways   ·   which one should you use?",
  "Tailwind is the other kind of framework: no ready-made components, just small classes you combine.");

{
  const s = plainSlide("The Same Cards, Two Ways",
    "Examples 02 and 03 are the same three menu cards. Open them side by side.\n\nBootstrap (left): the card is card, card-body, card-title. Someone else decided what a card looks like, and you get it with a few names.\n\nTailwind (right): there's no 'card' class. The article gets bg-white (background), rounded-lg (border-radius), shadow (box-shadow) and p-5 (padding). You decide every part.\n\nThey look nearly the same. Ask: which HTML would you rather read? Which would you rather change to a totally different look?\n\nTailwind's HTML is longer, but you rarely need a CSS file. Bootstrap's HTML is shorter, but making it look different takes CSS, like Example 01.\n\nOne more thing worth saying now that they have done last class: Tailwind gives you none of the interactive components. No modal, no carousel, no accordion. If you want those with Tailwind you write the JavaScript yourself or add another library. That is a real part of the trade-off.");
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
  notes: "Pop Quiz 7. Students who got Week 10's Pop Quiz 4 (col-md-4 on a phone) should get this. A is the common wrong answer.",
  revealNotes: "Answer: D.\n\nCheck in DevTools: inspect the grid in Example 03 and the rule is inside @media (width >= 48rem). 48rem is 768px.",
});

{
  const s = panelSlide("Bootstrap, Tailwind, or Plain CSS?",
    "There's no single right answer, and students will hear strong opinions online.\n\n- Bootstrap: when you need something that works and looks decent quickly, especially with menus, modals and forms. Very common on existing sites, admin pages and internal tools, so you'll likely maintain some at work. And it brings the interactive components with it, which is what last class was about.\n- Tailwind: when you have a design to build, and want it to look like no one else's. Very popular for new projects, and often used with JavaScript frameworks like React. Its build step makes small CSS files. It gives you no components, so a modal is your problem.\n- Plain CSS: small sites, and anywhere you want full control with no extra files. Also: everything you learn in plain CSS makes both frameworks easier.\n\nWhy we demoed these two: they're the most common, they're one of each kind, and between them students will recognize almost any framework they meet.");
  [["Bootstrap", "Pick it when you need it working fast: menus, modals, forms and pages that look good with little CSS.", C.teal],
    ["Tailwind", "Pick it when you have your own design to build. No ready-made components: a modal is your own problem.", C.purple],
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
  notes: "Pop Quiz 8, a last check before the activity.",
  revealNotes: "Answer: A.\n\nIt's the card from Example 03. B is the card from Example 02.",
});

// ---------------------------------------------------------------- Harbourside
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
  n: 1, title: "Link Bootstrap", time: "4 min", lw: 5.0, codeTitle: "IN THE HEAD, WHERE THE TRY IT 1 COMMENT IS", lang: "html",
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
    ["Add the script", ["Bootstrap's JavaScript just before ", ["</body>", "c"], ". ", ["Last class: no script, nothing moves.", "b"]]],
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
  notes: "The code is on the slide, but it's long: students can copy it from activity/answer/index.html. The slide leaves out three of the dropdown items to fit; the answer file has them all. Careful: the answer file's image paths start with ../ because it's one folder deeper. In activity/index.html it's images/logo.svg.\n\nThis is last class's pattern twice over, so say so: data-bs-toggle=\"collapse\" on the menu button is Example 02's collapse, and data-bs-toggle=\"dropdown\" on the Menu link is the dropdown from the component list. Same two attributes, same matching ids.\n\nOther new things since Week 10's Example 06:\n- bg-harbour is our own class (it's already in style.css): navy.\n- data-bs-theme=\"dark\" tells Bootstrap the background is dark, so the links and the menu button turn light.\n- The logo goes inside the navbar-brand. d-flex align-items-center gap-2 is flexbox from Week 8, as classes.\n\nThe usual problems, and they are exactly last class's two: the script is missing (nothing opens, no error), or the id on the collapse div doesn't match data-bs-target.",
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
  notes: "Each article keeps its id (for the dropdown links) and its class=\"special\". h-100 makes every card the full height of its column, so they line up even with different amounts of text. shadow-sm adds a light shadow.\n\nclass=\"card-title h5\" on the h3: the right heading level, at a smaller size.\n\nThe column classes: count them at each size. Under 576px, no class applies, so full width. sm (576px+): 6 of 12, so 2 across. lg (992px+): 3 of 12, so 4 across.\n\nCheck with the device toolbar at 360px, 700px and 1100px.\n\nFast finishers who want more: put an Order button on each card that opens a modal, straight from last class's Example 01.",
});

trySlide({
  n: 5, title: "The Table and the Footer", time: "6 min", lw: 5.0, codeTitle: "THE ASIDE, THEN THE FOOTER", lang: "html",
  steps: [
    ["Make the aside a row", ["The address in ", ["col-md-5", "c"], ", the table in ", ["col-md-7", "c"], "."]],
    ["Style the table", [["table table-striped", "c"], ", inside a ", ["table-responsive", "c"], " div."]],
    ["Style the footer", ["Navy, white text, centred, with padding."]],
  ],
  code: [...upTo(ANS, '<aside id="hours"', '<table class="table'), "            ...", "         </table>", "      </div>", "   </div>", "</aside>", "", ...upTo(ANS, '<footer id="contact"', "</footer>")],
  notes: "table: Bootstrap's table style (padding, lines between rows). table-striped: every second row shaded. Others to try: table-hover, table-bordered, table-sm.\n\ntable-responsive: if the table is wider than the screen, it scrolls sideways inside its own box, instead of the whole page scrolling. We did this by hand in Week 7 with the table-wrap div. That's Pop Quiz 9, coming up next.\n\nThe footer reuses bg-harbour. text-center centres the text, py-4 is padding on the top and bottom. mb-1 and mb-0 tidy up the paragraph margins.\n\nThe caption shows under the table: Bootstrap puts captions at the bottom. Add caption-top to the table to move it.",
});

quiz({
  q: "Your hours table is wider than a phone screen. What does `table-responsive` do about it?",
  options: ["Shrinks the text until the table fits", "Stacks the columns on top of each other", "Lets the table scroll sideways inside its own box", "Hides the columns that don't fit"], answer: 2,
  why: ["The `table-responsive` div gets `overflow-x: auto`, so only the table scrolls and the rest of the page stays still.",
    "We did exactly this by hand in Week 7, with a `table-wrap` div. Bootstrap has a class for it."],
  notes: "Pop Quiz 9. Ask them to test it: device toolbar at 360px, then drag the table sideways with the mouse.",
  revealNotes: "Answer: C.\n\nWithout it, a wide table makes the whole page scroll sideways, which is the thing we spent Week 9 avoiding.\n\nWorth adding: table-responsive-md only does it below 768px, and leaves the table alone above that.",
});

trySlide({
  n: 6, title: "Make It Yours", time: "8 min", lw: 5.6, codeTitle: "ADD TO STYLE.CSS, UNDER THE TRY IT 6 HEADING",
  steps: [
    ["Open style.css", ["It's linked after Bootstrap, so a tie goes to us."]],
    ["Add the rules on the right", "Georgia headings, orange-red prices, the special, our button."],
    ["Hover the banner button", ["It should stay orange-red: we set the ", ["variables", "b"], ", like Example 01."]],
    ["Make it yours", "Change a colour, a font, or a Bootstrap class. Anything you like."],
  ],
  code: upTo(ANSCSS, "/* 1. Headings", "font-weight: bold;").concat(["}", "", "/* 3. and 4.: the special, and .btn-harbour", "   (copy them from answer/style.css) */"]),
  tip: [[["Bootstrap's rule still winning? ", "b"], "Inspect it in DevTools. It may be more specific, like ", [".btn:hover", "c"], "."]],
  notes: "This is what turns a Bootstrap page into Harbourside. The whole file is about 40 lines, compared with about 300 lines in Week 9's style.css. That's the trade-off in one number.\n\nThe special: .special sets background-color and border-color. Bootstrap's .card rule uses the same specificity (one class), so ours wins because it's later. That is Pop Quiz 10, next.\n\n.btn-harbour sets the button variables, so the hover and active colours are ours too. Orange-red (#B5451B) and not the logo's lighter orange (#E97132): white text on the lighter orange is too low contrast to read easily.\n\nFast finishers: make the navbar stick to the top (sticky-top on the nav), add a card-img-top photo to a card, or put an accordion of FAQs under the menu, from last class's Example 02.",
});

quiz({
  q: "Both rules are one class, so they tie on specificity. Which background does the special card get?",
  code: ["/* bootstrap.min.css, linked FIRST */", ".card {", "   background-color: #fff;", "}", "", "/* your style.css, linked SECOND */", ".special {", "   background-color: #FFF8E1;", "}"], lang: "css", codeTitle: "ONE ELEMENT, class=\"card special\"",
  options: ["White: `.card` is Bootstrap's own class, so it wins", "Cream: they tie on specificity, and ours is read last", "White: Bootstrap's file is bigger and loads first", "Neither: you need `!important` to beat a framework"], answer: 1,
  why: [["Two single-class selectors tie. When rules tie, ", ["the one the browser reads last wins", "b"],
    " - which is the whole reason ", ["style.css", "c"], " goes after Bootstrap's link."],
  ["This only works because they tie. Pop Quiz 1's ", [".btn:hover", "c"], " did ", ["not", "bi"], " tie, and being later did not help at all."]],
  notes: "Pop Quiz 10, the last one of the week, and deliberately the bookend to Pop Quiz 1.\n\nIf anyone says 'but Pop Quiz 1 said ours loses' - that is exactly the right confusion to have, and the answer slide clears it up: specificity first, then order. Order only decides a tie.",
  revealNotes: "Answer: B.\n\nSay the rule as two steps, in order:\n1. Is one selector more specific? If yes, it wins, and nothing else matters. (Pop Quiz 1: .btn:hover beat .btn-primary.)\n2. If they tie, the last one the browser read wins. (Here: ours.)\n\nThat is why the order of the two <link> tags matters, and why ours goes second.\n\nD is worth a word: !important does work, and it is also how stylesheets become unmaintainable. Specificity and order first; !important as a last resort.",
});

{
  const s = activitySlide("Class Activity: Submit Harbourside",
    "This replaces last year's activity (download a Bootstrap example and customize it), which is now the fast-finisher option.\n\nStudents take the screenshots with the device toolbar, so the phone width shows. Its menu (the three dots on the toolbar) has 'Capture screenshot', and 'Capture full size screenshot' gets the whole page.\n\nCheck that the Week 11 Brightspace submission box is set up before class.\n\nThe interactive extra is optional, but it is five minutes and it is a straight lift from last class's examples. An Order button that opens a modal, or an accordion of FAQs under the menu.\n\nTailwind version: for students who want a challenge, rebuild the menu cards with Tailwind's Play CDN, using Example 03 as a guide. Don't link Bootstrap and Tailwind in the same page: their classes clash (both have p-3, for example).");
  const LW = 6.05;
  card(s, { x: 0.75, y: 1.45, w: LW, h: 5.5 });
  txt(s, "Your page should have:", { x: 1.05, y: 1.65, w: LW - 0.6, h: 0.4, fontSize: 18, bold: true, color: C.black });
  ["A navbar that folds into a menu button", "The menu as cards in a responsive grid", "Your own colours in style.css"].forEach((t, i) => {
    const y = 2.2 + i * 0.6;
    numCircle(s, "✓", 1.05, y, C.teal, 0.42);
    txt(s, t, { x: 1.65, y, w: LW - 1.2, h: 0.42, fontSize: 16, color: C.black, valign: "middle" });
  });
  numCircle(s, "+", 1.05, 4.0, C.purple, 0.42);
  richBullets(s, [[["Optional: ", "b"], "one thing from last class. A modal, or an accordion."]],
    { x: 1.65, y: 4.0, w: LW - 1.2, h: 0.42, size: 16, bullet: false, valign: "middle" });
  txt(s, "Then:", { x: 1.05, y: 4.55, w: LW - 0.6, h: 0.4, fontSize: 18, bold: true, color: C.black });
  [[["Take two screenshots: wide, and at phone width (", ["Ctrl + Shift + M", "b"], ")."]], [["Drop them in the box on Week 11 Brightspace."]]].forEach(([d], i) => {
    const y = 5.05 + i * 0.85;
    numCircle(s, i + 1, 1.05, y, C.orange, 0.42);
    richBullets(s, [d], { x: 1.65, y: y - 0.03, w: LW - 1.2, h: 0.8, size: 16, bullet: false, color: C.black });
  });
  const rx = 0.75 + LW + 0.35, rw = W - rx - 0.6;
  txt(s, "Finished early?", { x: rx, y: 1.45, w: rw, h: 0.4, fontSize: 18, bold: true, color: C.black });
  [["Add something that moves", ["An Order button that opens a modal (", ["js/01", "c"], "), or an FAQ accordion (", ["js/02", "c"], ") under the menu."]],
    ["Last year's activity", ["Download one of the Bootstrap examples (", ["getbootstrap.com/docs/5.3/examples", "c"], ") and make it yours."]],
    ["Try Tailwind", ["Rebuild the menu cards in a new file with Tailwind, using Example 03 as a guide."]]].forEach(([t, d], i) => {
    const y = 2.0 + i * 1.6;
    card(s, { x: rx, y, w: rw, h: 1.4 });
    txt(s, t, { x: rx + 0.25, y: y + 0.12, w: rw - 0.5, h: 0.38, fontSize: 16, bold: true, color: [C.purple, C.teal, C.orange][i] });
    richBullets(s, [d], { x: rx + 0.25, y: y + 0.5, w: rw - 0.5, h: 0.85, size: 14, bullet: false });
  });
}

{
  const s = plainSlide("Resources",
    "Everything on this slide is a link: the names, the addresses and the red play buttons.\n\nUseful links:\n- Bootstrap's docs: the left menu has Layout (the grid), Content, Forms, Components, Helpers and Utilities.\n- Bootstrap's JavaScript page: what the bundle actually is, and the list of which components need it. Worth a look after last class.\n- W3Schools: short lessons, each with a Try It editor. Make sure the address says bootstrap5.\n- Tailwind's docs, for Examples 02 and 03.\n\nVideos:\n- Web Dev Simplified (Kyle Cook): the Bootstrap 5 Crash Course covers the grid, six components and the utility classes in about 70 minutes. It's the best match for this week.\n- thenewboston (Bucky Roberts): a beginner series, short videos. It's Bootstrap 3, from 2015. The grid works the same way, but some class names have changed. A good test of what we've learned: spot what's changed.\n- Web Dev Simplified, Should You Use Bootstrap?: the pros and cons, like our 'Bootstrap, Tailwind, or Plain CSS?' slide.\n- Web Dev Simplified, Learn JavaScript in 12 Minutes: for anyone who got curious last class and wants a head start before we start properly.");
  const cw = (W - 1.5 - 0.35) / 2, ch = 1.18, gap = 0.14;
  [["Useful links", C.teal, false, [
    ["Bootstrap 5.3 documentation", "https://getbootstrap.com/docs/5.3/", "getbootstrap.com/docs/5.3", "Layout, Components and Utilities are the sections we use."],
    ["Bootstrap's JavaScript", "https://getbootstrap.com/docs/5.3/getting-started/javascript/", "getbootstrap.com/docs/5.3/getting-started/javascript", "What the bundle is, and which components need it."],
    ["W3Schools Bootstrap 5 tutorial", "https://www.w3schools.com/bootstrap5/", "w3schools.com/bootstrap5", "Short lessons with Try It editors. Check it says bootstrap5."],
    ["Tailwind CSS documentation", "https://tailwindcss.com/docs", "tailwindcss.com/docs", "Every utility class, with what CSS it writes."],
  ]], ["Videos", C.red, true, [
    ["Bootstrap 5 Crash Course", "https://www.youtube.com/watch?v=Jyvffr3aCp0", "Web Dev Simplified · YouTube", "The grid, six components and the utilities, in about 70 minutes."],
    ["Bootstrap Tutorials for Beginners", "https://www.youtube.com/playlist?list=PL6gx4Cwl9DGBPw1sFodruZUPheWVKchlM", "thenewboston · YouTube playlist", "Short videos. Bootstrap 3: the grid's the same, some class names aren't."],
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
    "If students remember four things: put your CSS after Bootstrap's and set the --bs-btn- variables; data-bs-toggle plus data-bs-target opens almost anything with no JavaScript from you; a silent component means a missing script or a mismatched id; and 'when this, do that' is what JavaScript is.\n\nAssignment 4 needs all of this week and last: a navbar, the grid, components, and your own styling in style.css. Harbourside is a good model to copy from, and one interactive component makes it stand out.\n\nThe Week 11 quiz is the ten pop quizzes from these two classes, word for word in meaning if not in wording. Say so.\n\nThen the last slide: the one-sentence summary.");
  labelBullets(s, [
    { label: "Customizing Bootstrap:", text: [
      "Your style.css after Bootstrap's. Specificity first, then order breaks the tie.",
      "Buttons: set the --bs-btn- variables, so hover and active follow.",
    ] },
    { label: "The parts that move:", text: [
      "data-bs-toggle + data-bs-target, and the bundle script at the end of the body.",
      "Tooltips are opt in. A toast is: when this, do that.",
    ] },
    { label: "Tailwind:", text: ["Small utility classes, about one property each. No ready-made components."] },
    { label: "Coming up:", text: ["Assignment 4, then JavaScript properly: custom form validation."] },
  ], { size: 17, sub: 14 });
}

dividerSlide("Write a one-sentence summary of what we did this week.", "One sentence, in your own words.",
  "Give everyone a minute. They can write it on paper or post it in the Teams chat. Ask two or three people to read theirs out.\n\nListen for whether anyone says that Bootstrap's components are just HTML plus a script that watches for clicks. If nobody does, say it: that is the idea that carries into every JavaScript class from here on.");

pres.writeFile({ fileName: OUT }).then(() => console.log("wrote " + OUT));
