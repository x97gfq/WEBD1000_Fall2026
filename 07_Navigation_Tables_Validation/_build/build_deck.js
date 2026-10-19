// Builds Navigation_Tables_Validation.pptx (Class 10) in the same style as the Class 5–9 decks.
// To rebuild: cd 07_Navigation_Tables_Validation/_build, run "npm install" once, then "npm run build".
// The "Result" pictures are screenshots of the example pages, taken with Microsoft Edge (headless) each build.
// Code slides read the example files directly, so edit those files and re-run to update the slides.
const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");
const { execFileSync } = require("child_process");

const REPO = path.join(__dirname, ".."); // the 07_Navigation_Tables_Validation folder
const SCR = __dirname;
const SHOTS = path.join(SCR, "shots");
const OUT = path.join(REPO, "Navigation_Tables_Validation.pptx");

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
pres.title = "Class 10: Navigation Menus, Tables and HTML Validation";
pres.author = "Jamie Symonds";

const W = 13.333, H = 7.5;
const PANEL_W = 4.1;
const CX = 5.2, CW = 7.55; // content column on panel slides

// ---------- helpers (same as the Class 5–9 decks) ----------
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
// frames: a grid of { img, label } screenshots (an animation paused at different times) in place of one result
function codeSlide({ kicker, title, bullets, file, lines, lang = "css", result, frames, cols, maxFrameH = 3.3, notes, size = 14, codeTitle }) {
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
  if (frames) {
    cols = cols || Math.min(frames.length, 2);
    const rows = Math.ceil(frames.length / cols), gap = 0.15, lab = 0.32, ratio = frames[0].img.ratio;
    let fw = (RW - gap * (cols - 1)) / cols;
    const gridH = (w) => rows * (w * ratio + lab) + (rows - 1) * gap;
    if (gridH(fw) > maxFrameH) fw = (maxFrameH - (rows - 1) * gap - rows * lab) / rows / ratio;
    const fh = fw * ratio, top = 6.85 - gridH(fw);
    frames.forEach(({ img, label }, i) => {
      framedImage(s, img, RX + (i % cols) * (fw + gap), top + lab + Math.floor(i / cols) * (fh + lab + gap), fw, label);
    });
    bh = top - by - 0.1;
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

// ---------- screenshots ----------
const SH = {
  ex01: shot("01_nav_structure_example.html", "ex01", 600, 250),
  ex02: shot("02_nav_horizontal_example.html", "ex02", 600, 110),
  ex02h: frame("02_nav_horizontal_example.html", "ex02h", { hover: "nav li:nth-child(3) a", w: 600, h: 110 }),
  ex03: shot("03_nav_vertical_example.html", "ex03", 400, 330),
  ex04b: frame("04_nav_dropdown_example.html", "ex04b", { hover: "nav > ul > li:nth-child(2), nav ul ul li:nth-child(2) a", w: 600, h: 230 }),
  ex05: shot("05_table_structure_example.html", "ex05", 500, 200),
  ex06: shot("06_table_styled_example.html", "ex06", 600, 240),
  ex07: frame("07_table_zebra_example.html", "ex07", { hover: "tbody tr:nth-child(3)", w: 600, h: 330 }),
  ex08: frame("08_table_responsive_example.html", "ex08", { js: "document.querySelector('.table-wrap').scrollLeft = 140;", w: 380, h: 270, bars: true }),
  ex09: shot("09_validation_broken_example.html", "ex09", 600, 280),
  ex11: shot("11_css_validation_broken_example.html", "ex11", 600, 260),
  ex12: shot("12_css_validation_fixed_example.html", "ex12", 600, 260),
  before: shot("activity/index.html", "before", 900, 640, 1),
  afterNav: frame("activity/answer/index.html", "afterNav", { hover: "header nav > ul > li:first-child, nav ul ul li:nth-child(3) a", w: 900, h: 640, scale: 1 }),
  afterTable: frame("activity/answer/index.html", "afterTable", { js: "document.querySelectorAll('main > section').forEach((e) => e.style.display = 'none');", w: 900, h: 640, scale: 1 }),
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
    { text: "Class 10:", options: { bullet: true, fontSize: 30, breakLine: true } },
    { text: "Navigation menus, tables, and checking your HTML with the W3C validator.", options: { bullet: { code: "25CB" }, indentLevel: 1, fontSize: 26, paraSpaceBefore: 6 } },
  ], { x: 6.85, y: 1.9, w: 5.6, h: 2.4, color: C.black, valign: "top" });
  s.addNotes("Welcome back! Last class: advanced selectors and CSS animation.\n\nToday is about the furniture every real site has: a navigation menu to get around, tables for data like prices and opening hours, and a way to check our HTML for the mistakes the browser quietly hides from us.\n\nWe're back on the Harbourside page from Class 8. By the end of class it will have a real navigation bar with a dropdown, an opening-hours table, and HTML that passes the W3C validator with zero errors.\n\nReminder: no class on Wednesday, October 21 (Assessment & Evaluation day, college open). Today is our only class this week.");
}

// 2. Attendance
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  txt(s, "Attendance", { x: 1.1, y: 0.75, w: 8, h: 0.9, fontSize: 44, color: C.black });
  s.addNotes("Take attendance. While you do, have students git pull and open the 07_Navigation_Tables_Validation folder in VS Code.");
}

// 3. Overview
{
  const s = panelSlide("Class 10: Class Overview",
    "Rough plan for the two hours:\n- 8 min: intro and getting the code\n- 40 min: navigation menus (four examples, then Try Its 1 to 4)\n- 30 min: tables (four examples, then Try Its 5 and 6)\n- 22 min: HTML validation (what, why, how, then Try It 7)\n- 10 min: CSS validation (Example 11, then Try It 8)\n- 7 min: make it yours and submit\nThat's about 117 minutes, so there's very little slack. If time runs short, Try It 8 can be finished as homework.\n\nEvery Try It adds to the same Harbourside page, so it builds up across the class, like Class 8. The submission is three screenshots on Brightspace (participation).\n\nNo class Wednesday, October 21: Assessment & Evaluation day. The college is open, there are just no classes.");
  txt(s, [
    { text: "Title: ", options: { bold: true, bullet: true, fontSize: 22 } },
    { text: "Navigation Menus, Tables and HTML Validation", options: { fontSize: 22, breakLine: true } },
    { text: " ", options: { fontSize: 10, breakLine: true } },
    { text: "Agenda:", options: { bold: true, bullet: true, fontSize: 22, breakLine: true } },
    { text: "Navigation menus: horizontal, vertical and dropdown", options: { bullet: true, indentLevel: 1, fontSize: 18, paraSpaceBefore: 4, breakLine: true } },
    { text: "Tables for data: structure, styling, zebra stripes, small screens", options: { bullet: true, indentLevel: 1, fontSize: 18, paraSpaceBefore: 4, breakLine: true } },
    { text: "Validation with the W3C HTML and CSS validators", options: { bullet: true, indentLevel: 1, fontSize: 18, paraSpaceBefore: 4, breakLine: true } },
    { text: "Harbourside, Part 2: eight Try Its, then submit", options: { bullet: true, indentLevel: 1, fontSize: 18, paraSpaceBefore: 4, breakLine: true } },
    { text: "No class Wednesday, October 21 (Assessment & Evaluation day)", options: { bullet: true, indentLevel: 2, fontSize: 14, paraSpaceBefore: 6, color: C.grey } },
  ], { x: CX, y: 1.95, w: CW, h: 3.6, color: C.black, valign: "top" });
  const steps = [["Navigate", "A menu to get around"], ["Tabulate", "Data in rows and columns"], ["Validate", "Find the hidden mistakes"]];
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
    "Examples 01–04 are navigation, 05–08 are tables, 09–10 are HTML validation and 11–12 are CSS validation. Every example is a single file with its CSS in a <style> block.\n\nThe activity folder is the Harbourside page from Class 8, with our finished Class 8 CSS already in style.css (tidied up). There's no nav yet, and the hours are still a bulleted list.\n\nThe starter index.html has 5 deliberate mistakes in it, and the Class 8 part of style.css has 3. Don't mention them yet! They're the surprises in Try Its 7 and 8.\n\nAsk students not to open activity/answer until they're stuck.");
  txt(s, [{ text: "1.  Pull this week's examples", options: { bold: true, fontSize: 18 } }], { x: CX, y: 1.95, w: CW, h: 0.4 });
  codeBox(s, ["git pull"], { x: CX, y: 2.45, w: 5.2, h: 0.7, maxSize: 16 });
  txt(s, "New folder: 07_Navigation_Tables_Validation (examples 01–12, and the activity folder).", { x: CX, y: 3.25, w: CW, h: 0.5, fontSize: 15, color: C.dark });
  txt(s, [{ text: "2.  Open each example with Live Server as we go", options: { bold: true, fontSize: 18 } }], { x: CX, y: 3.95, w: CW, h: 0.4 });
  txt(s, "Hover the menus, press Tab, shrink the window. Change a value, save, and see what happens.", { x: CX, y: 4.4, w: CW, h: 0.8, fontSize: 15, color: C.dark });
  txt(s, [{ text: "3.  The activity: Harbourside, Part 2", options: { bold: true, fontSize: 18 } }], { x: CX, y: 5.4, w: CW, h: 0.4 });
  txt(s, "activity/index.html and activity/style.css. Our Class 8 page, ready for a nav and a table.", { x: CX, y: 5.85, w: CW, h: 0.8, fontSize: 15, color: C.dark });
}

// 5. Divider
dividerSlide("Navigation Menus", "Every menu is a list   ·   horizontal   ·   vertical   ·   dropdown",
  "Every site you use has a navigation menu: the set of links that gets you around. It's often the first thing people look for, so it has to be easy to find, easy to read and easy to click (or tap, or Tab to).\n\nThe good news: every menu, however fancy, starts as the same plain HTML. The difference is all CSS.");

// 6. Three kinds of menu
{
  const s = panelSlide("Three Kinds of Menu",
    "Horizontal: the classic bar across the top. Best for a handful of main links (about 3 to 7).\nVertical: a stack of links down the side. Good for longer lists: sections of a menu, account settings, documentation.\nDropdown: a horizontal bar where some items open a smaller list. Good for grouping related pages without making the bar too long.\n\nOn phones, a horizontal menu usually collapses into a 'hamburger' button (☰). That needs a little JavaScript, which we'll meet later in the program.\n\nThe key point is the bottom of the slide: all three are the same HTML structure. nav, then ul, then li, then a.");
  const kinds = [
    ["Horizontal", "Side by side across the top", "A few main links"],
    ["Vertical", "Stacked down the side", "Longer lists, sections"],
    ["Dropdown", "A list inside a list", "Grouping related pages"],
  ];
  const cw = 2.4, gap = 0.18;
  kinds.forEach(([t, d, use], i) => {
    const x = CX + i * (cw + gap), y = 1.95;
    card(s, { x, y, w: cw, h: 3.35 });
    // a little drawing of each menu
    const mx = x + 0.2, my = y + 0.25, mw = cw - 0.4;
    if (i === 0) {
      s.addShape(pres.shapes.RECTANGLE, { x: mx, y: my + 0.3, w: mw, h: 0.42, fill: { color: C.navy }, line: { color: C.navy } });
      ["Home", "Menu", "Contact"].forEach((l, j) => txt(s, l, { x: mx + 0.08 + j * 0.65, y: my + 0.3, w: 0.62, h: 0.42, fontSize: 10, bold: true, color: C.white, valign: "middle" }));
      s.addShape(pres.shapes.RECTANGLE, { x: mx + 0.08, y: my + 0.68, w: 0.42, h: 0.04, fill: { color: C.orange }, line: { color: C.orange } });
    } else if (i === 1) {
      ["Starters", "Fish & Chips", "Lobster", "Drinks"].forEach((l, j) => {
        const on = j === 1;
        s.addShape(pres.shapes.RECTANGLE, { x: mx, y: my + j * 0.33, w: 1.35, h: 0.28, fill: { color: on ? C.navy : C.actBg }, line: { color: on ? C.navy : C.actBg } });
        txt(s, l, { x: mx + 0.08, y: my + j * 0.33, w: 1.25, h: 0.28, fontSize: 9, color: on ? C.white : C.navy, valign: "middle" });
      });
    } else {
      s.addShape(pres.shapes.RECTANGLE, { x: mx, y: my, w: mw, h: 0.42, fill: { color: C.navy }, line: { color: C.navy } });
      ["Home", "Menu ▾", "Contact"].forEach((l, j) => txt(s, l, { x: mx + 0.08 + j * 0.65, y: my, w: 0.62, h: 0.42, fontSize: 10, bold: true, color: C.white, valign: "middle" }));
      s.addShape(pres.shapes.RECTANGLE, { x: mx + 0.66, y: my + 0.42, w: 1.15, h: 0.84, fill: { color: C.white }, line: { color: "B0B8BC", width: 0.75 },
        shadow: { type: "outer", blur: 4, offset: 2, angle: 90, color: "000000", opacity: 0.25 } });
      ["Fish & Chips", "Lobster", "Kids' Menu"].forEach((l, j) => txt(s, l, { x: mx + 0.74, y: my + 0.45 + j * 0.26, w: 1.05, h: 0.26, fontSize: 9, color: C.navy, valign: "middle" }));
    }
    txt(s, t, { x: x + 0.2, y: y + 1.75, w: cw - 0.4, h: 0.4, fontSize: 17, bold: true, color: C.black });
    txt(s, d, { x: x + 0.2, y: y + 2.15, w: cw - 0.4, h: 0.5, fontSize: 13, color: C.dark, valign: "top" });
    txt(s, "Use for: " + use, { x: x + 0.2, y: y + 2.7, w: cw - 0.4, h: 0.5, fontSize: 11.5, italic: true, color: C.grey, valign: "top" });
  });
  txt(s, "Same HTML every time:", { x: CX, y: 5.6, w: CW, h: 0.4, fontSize: 15, bold: true, color: C.black });
  [["<nav>", C.navy], ["<ul>", C.teal], ["<li>", C.purple], ["<a>", C.orange]].forEach(([t, col], i) => {
    pill(s, t, CX + i * 1.55, 6.1, col, 1.1);
    if (i < 3) txt(s, "→", { x: CX + 1.1 + i * 1.55, y: 6.07, w: 0.45, h: 0.4, fontSize: 16, bold: true, color: C.grey, align: "center" });
  });
}

// 7. Example 01: structure
codeSlide({
  kicker: "NAVIGATION 1 · EXAMPLE 01", title: "Every Menu Is a List", file: "01_nav_structure_example.html", lines: bodyOf("01_nav_structure_example.html"), lang: "html", result: SH.ex01, size: 14,
  bullets: [
    [["<nav>", "c"], " says ", ["\"this is the main way around the site\"", "b"], ". Screen readers can jump straight to it."],
    ["A ", ["<ul>", "c"], " of links, because a menu ", ["is", "b"], " a list. Screen readers say \"list, 4 items\"."],
    [["aria-label", "c"], " names it, if a page has more than one nav."],
  ],
  notes: "Open 01_nav_structure_example.html. With no CSS it's just a bulleted list of links. That's fine! It works, it's accessible, and it's the starting point for every menu we'll build.\n\nWhy a list? Because it is one: a group of related links. A screen reader announces 'navigation, list, 4 items', so the user knows how big the menu is before they start. A row of bare links or a pile of divs doesn't tell them that.\n\nWe met <nav> as a landmark in Class 4. Only use it for major navigation blocks (the main menu, maybe a table of contents), not for every group of links on the page. A link list in the footer doesn't need one.\n\nOn a real site the href values would be page names like menu.html or contact.html. Here they're # links so nothing breaks when you click.",
});

// 8. Example 02: horizontal
codeSlide({
  kicker: "NAVIGATION 2 · EXAMPLE 02", title: "A Horizontal Menu", file: "02_nav_horizontal_example.html",
  lines: between(styleOf("02_nav_horizontal_example.html"), "/* 1."), size: 13,
  frames: [{ img: SH.ex02, label: "NORMAL" }, { img: SH.ex02h, label: "HOVER ON \"ABOUT\"" }], cols: 1, maxFrameH: 2.2,
  bullets: [
    [["list-style: none", "c"], " and ", ["padding: 0", "c"], " remove the bullets and the indent."],
    [["display: flex", "c"], " puts the items side by side; ", ["gap", "c"], " spaces them. (A sneak peek: more on Flexbox next class.)"],
    [["display: block", "c"], " + ", ["padding", "c"], " on the ", ["a", "b"], " makes the whole box clickable, not just the word."],
  ],
  notes: "Open 02_nav_horizontal_example.html. Walk through the four numbered comments; each one is one step from Example 01 to a real menu bar.\n\n1. The bullets and indent go. (Last year's version of this slide used display: inline-block on each li, which also works. Flexbox is the modern way, and it lines things up more reliably.)\n2. display: flex on the ul lays its children, the li's, out in a row. gap adds space between them, but not at the ends. We'll do Flexbox properly next class; today we just need these two lines.\n3. Put the padding on the a, not the li, and make the a display: block. Then the whole coloured box is the link: a much bigger target for a mouse, and especially for a finger on a phone. WCAG 2.5.8 asks for targets of at least 24 by 24 pixels.\n4. aria-current=\"page\" tells screen readers 'you are here', and the attribute selector from Class 9 lets us style it too. One attribute does both jobs.\n\nTry in class: comment out display: flex and save. The menu goes back to a stacked list.",
});

// 9. Example 03: vertical
codeSlide({
  kicker: "NAVIGATION 3 · EXAMPLE 03", title: "A Vertical Menu", file: "03_nav_vertical_example.html",
  lines: between(styleOf("03_nav_vertical_example.html"), "nav ul {"), result: SH.ex03, size: 14,
  bullets: [
    ["A list is ", ["already vertical", "b"], ". No flex needed: just remove the bullets and style the links."],
    [["border-left", "c"], " that's ", ["transparent", "c"], " until you hover: the text doesn't jump sideways."],
  ],
  notes: "Open 03_nav_vertical_example.html. This is the kind of menu you see down the side of a page: menu sections, account settings, a course's list of weeks in Brightspace.\n\nSame starting point (remove the bullets), but no display: flex this time, because a list stacks vertically on its own.\n\nThe transparent border trick: the left border is always there, just invisible. On hover it changes colour. If the border only appeared on hover, the box would get 4px wider and the text would shift. Little details like this make a menu feel solid.\n\nThe 'current page' style is the strongest one (navy with white text), so you can see where you are at a glance.\n\nThe width: 220px on the nav keeps it narrow, like a sidebar. Placing it beside the page content is a layout job (Flexbox), coming next class.",
});

// 10. Example 04: dropdown
codeSlide({
  kicker: "NAVIGATION 4 · EXAMPLE 04", title: "A Dropdown Menu", file: "04_nav_dropdown_example.html",
  lines: between(styleOf("04_nav_dropdown_example.html"), "/* 1.", "nav ul ul a {"), size: 13,
  frames: [{ img: SH.ex04b, label: "MOUSE OVER \"MENU\"" }], cols: 1, maxFrameH: 2.6,
  bullets: [
    ["The dropdown is a ", ["list inside a list item", "b"], ": a ", ["<ul>", "c"], " inside the Menu ", ["<li>", "c"], "."],
    [["position: absolute", "c"], " floats it over the page; ", ["position: relative", "c"], " on the ", ["li", "b"], " keeps it under its own item."],
    [[":focus-within", "c"], " opens it for ", ["keyboard", "b"], " users too, not just the mouse."],
  ],
  notes: "Open 04_nav_dropdown_example.html. Hover Menu: the list drops down. Now click in the page and press Tab: when you reach Menu, the dropdown opens too, and you can Tab through its links. That's :focus-within. It matches an element when it, or anything inside it, has focus.\n\nLast year's version of this slide only used :hover, which means keyboard users could never open the dropdown. That's a WCAG failure (2.1.1 Keyboard). One extra selector fixes it.\n\nThe HTML: look at the body of the file. The Menu li contains a link AND a whole second <ul>.\n\nThe CSS, step by step:\n1. position: relative on each li. On its own it does nothing visible. It just says 'measure absolute things inside me from here'.\n2. The inner list is hidden (display: none) and positioned absolutely: top: 100% means 'just below the bottom of my li', left: 0 lines it up with the li's left edge. Absolute elements don't take up space, so the dropdown floats over the paragraph instead of pushing it down.\n3. When the li is hovered or has focus inside it, show the list.\n\nThe > in 'li:hover > ul' is the child combinator from Class 9: only the li's own list, not lists deeper down.\n\nTry: remove position: relative from nav li. The dropdown jumps to the left edge of the page, because it's now measured from the page itself.\n\nReal-world note: big sites usually open dropdowns with a button and a few lines of JavaScript, so they work on touch screens too. The CSS version is the foundation.",
});

// 11. Activity intro
{
  const s = activitySlide("Harbourside, Part 2: Today's Activity",
    "This is the plan for the hands-on part. Same Harbourside page as Class 8, with eight Try Its spread through the class: four for the nav, two for the table, and two for validation (HTML, then CSS).\n\nThe left picture is the starter: our Class 8 page, no nav, hours as a bullet list. (Look closely: the prices are plain black. That's one of the hidden CSS mistakes Try It 8 finds.) The right pictures are where we'll be by the end.\n\nThe starter style.css has the Class 8 CSS at the top, then an empty heading for each Try It. Everything we type goes under those headings.\n\nSubmission (participation): three screenshots on Brightspace: the page with the dropdown open, the HTML validator saying 'No errors', and the CSS validator saying 'No Error Found'.");
  const iw = (W - 1.5 - 0.4) / 3;
  [[SH.before, "START"], [SH.afterNav, "END: THE NAV"], [SH.afterTable, "END: THE TABLE"]].forEach(([img, label], i) => {
    framedImage(s, img, 0.75 + i * (iw + 0.2), 1.75, iw, label);
  });
  const y0 = 1.75 + iw * SH.before.ratio + 0.25, bh = 6.95 - y0, LW = 7.3;
  card(s, { x: 0.75, y: y0, w: LW, h: bh });
  const tis = [["1", "Nav HTML + horizontal"], ["2", "Style the links"], ["3", "A dropdown"], ["4", "Make it readable"], ["5", "Hours table HTML"], ["6", "Style the table"], ["7", "Validate the HTML"], ["8", "Validate the CSS"]];
  const colrs = [C.orange, C.orange, C.orange, C.orange, C.blue, C.blue, C.green, C.green];
  tis.forEach(([n, t], i) => {
    const x = 1.0 + Math.floor(i / 4) * 3.55, y = y0 + 0.18 + (i % 4) * ((bh - 0.3) / 4);
    numCircle(s, n, x, y, colrs[i], 0.36);
    txt(s, t, { x: x + 0.48, y: y - 0.04, w: 2.9, h: 0.44, fontSize: 13.5, bold: true, color: C.black, valign: "middle" });
  });
  const x = 0.75 + LW + 0.25, w = W - x - 0.6;
  card(s, { x, y: y0, w, h: bh, fill: C.dark, border: C.dark });
  txt(s, "Submit to Brightspace", { x: x + 0.25, y: y0 + 0.15, w: w - 0.5, h: 0.4, fontSize: 17, bold: true, color: C.white });
  txt(s, [
    { text: "Three screenshots:", options: { bold: true, breakLine: true } },
    { text: "1. Your page, with the dropdown open", options: { breakLine: true } },
    { text: "2. The HTML validator: \"No errors\"", options: { breakLine: true } },
    { text: "3. The CSS validator: \"No Error Found\"", options: { breakLine: true } },
    { text: "Win + Shift + S", options: { bold: true, color: "F2C6AC", fontFace: MONO } },
    { text: " takes a screenshot." },
  ], { x: x + 0.25, y: y0 + 0.6, w: w - 0.5, h: bh - 0.75, fontSize: 13.5, color: "D9D9D9", valign: "top", paraSpaceBefore: 4 });
}

// 12. Try It 1: nav HTML + horizontal
trySlide({
  n: 1, title: "Add the Nav", time: "8 min",
  steps: [
    ["Open activity/index.html with Live Server", "Our Class 8 page, with no menu yet."],
    ["Type the nav HTML", ["In index.html, inside ", ["<header>", "c"], ", under the ", ["h1", "c"], ". Save: a bulleted list appears."]],
    ["Type the CSS", ["In style.css, under ", ["Try It 1", "b"], ". Save: the links line up in a row."]],
    ["Play with it", ["Change ", ["gap: 4px", "c"], " to ", ["40px", "c"], ". Then take out ", ["display: flex", "c"], " and put it back."]],
  ],
  codeTitle: "INDEX.HTML, THEN STYLE.CSS", lang: "mixed",
  code: ["<nav aria-label=\"Main\">", "   <ul>", "      <li><a href=\"#menu\" aria-current=\"page\">Menu</a></li>", "      <li><a href=\"#hours\">Hours &amp; Location</a></li>", "      <li><a href=\"#contact\">Contact</a></li>", "   </ul>", "</nav>", "",
    "/* style.css, Try It 1 */", "nav ul {", "   list-style: none;", "   margin: 0;", "   padding: 0;", "   display: flex;", "   gap: 4px;", "}"],
  tip: [["The links go to the ", ["ids", "b"], " already on the page: ", ["#menu", "c"], ", ", ["#hours", "c"], " and ", ["#contact", "c"], ". Click them!"]],
  notes: "Type it together. Two files this time: the HTML in index.html (inside the header, under the h1, where the 'Try It 1' comment is), then the CSS in style.css.\n\nSave after the HTML first, so they see the plain bulleted list (Example 01). Then save after the CSS and watch it become a row (Example 02). Seeing both steps makes the point that CSS is doing all the work.\n\n&amp; is the entity for & from Class 1. A bare & usually works, but the entity is the safe habit.\n\nThe links jump down the page to the sections with those ids. That's an in-page link: href=\"#hours\" goes to id=\"hours\".\n\nUntil Try It 2, the links are the browser's default blue and purple on navy: hard to read. That's fixed next.",
});

// 13. Try It 2: style the links
trySlide({
  n: 2, title: "Style the Links", time: "6 min",
  steps: [
    ["Type the CSS on the right", ["Under ", ["Try It 2", "b"], ". Save."]],
    ["Hover each link", "The whole box lights up, not just the word."],
    ["Press Tab", ["Click in the page, then Tab: the ", [":focus-visible", "c"], " style follows the keyboard."]],
    ["Play with it", ["Change the hover colour. Try ", ["border-bottom: 4px dotted", "c"], " on the current page."]],
  ],
  code: ["nav a {", "   display: block;", "   padding: 12px 16px;", "   color: white;", "   text-decoration: none;", "   font-weight: bold;", "}", "", "nav a:hover,", "nav a:focus-visible {", "   background-color: #1B5E82;", "}", "", "nav a[aria-current=\"page\"] {", "   border-bottom: 4px solid #E97132;", "}"],
  tip: [[["aria-current=\"page\"", "c"], " tells screen readers \"you are here\". The attribute selector (Class 9) styles it."]],
  notes: "Now it looks like a real menu bar.\n\nWhite on navy is 13.6 : 1 contrast (from Class 8). White on the hover teal #1B5E82 is about 7 : 1, still well over 4.5 : 1.\n\nThe orange underline is decoration, not text, so the 3 : 1 rule for non-text elements applies, and #E97132 on navy passes that.\n\nWhy hover AND focus-visible in one rule: anything that reacts to the mouse should also react to the keyboard. Grouping them means you can't forget one.",
});

// 14. Try It 3: dropdown
trySlide({
  n: 3, title: "A Dropdown", time: "10 min",
  steps: [
    ["Replace the Menu li", ["In index.html. The new ", ["<ul>", "c"], " goes ", ["inside", "b"], " the ", ["<li>", "c"], ", after the link."]],
    ["Type the CSS", ["Under ", ["Try It 3", "b"], ". Save and hover Menu."]],
    ["Can't read it? Good!", "The links are white on a pale page. Try It 4 fixes it."],
    ["Press Tab", ["When you reach Menu, it opens. That's ", [":focus-within", "c"], "."]],
  ],
  codeTitle: "INDEX.HTML, THEN STYLE.CSS", lang: "mixed",
  code: ["<li>", "   <a href=\"#menu\" aria-current=\"page\">Menu &#9662;</a>", "   <ul>", "      <li><a href=\"#haddock\">Haddock &amp; Chips</a></li>", "      <li><a href=\"#clams\">Clam Strips</a></li>", "      <li><a href=\"#lobster\">Lobster Roll</a></li>", "      <li><a href=\"#kids\">Kids' Fish Bites</a></li>", "   </ul>", "</li>", "",
    "/* style.css, Try It 3 */", "nav li { position: relative; }", "", "nav ul ul {", "   display: none;", "   position: absolute;", "   top: 100%;", "   left: 0;", "}", "", "nav li:hover > ul,", "nav li:focus-within > ul {", "   display: block;", "}"],
  notes: "The trickiest Try It, so go slowly and check the HTML first. The most common mistake: putting the new <ul> after the </li> instead of inside it. Then it's a sibling, not a child, and the CSS never finds it.\n\n&#9662; is a numbered entity for the small down-triangle ▾ (Class 1 entities again). It tells sighted users 'this one opens'.\n\nnav li { position: relative; } is written on one line to save space. It's exactly the same as three lines.\n\nThe planned 'bug': when the dropdown opens, the links are white (from Try It 2) on a transparent background, so they almost disappear against the pale page. Let students discover this, then fix it in Try It 4. A good moment to point out that a white-on-white menu passes no contrast test.\n\nWhy does the dropdown stack vertically when the main menu is a row? 'nav ul ul' is more specific than 'nav ul', so its display wins: none, then block. The cascade again!\n\nOne more thing to notice: the Kids' Fish Bites link won't jump anywhere when clicked. There's a reason, and the validator finds it in Try It 7. Don't give it away.",
});

// 15. Try It 4: readable dropdown
trySlide({
  n: 4, title: "Make the Dropdown Readable", time: "4 min",
  steps: [
    ["Type the CSS on the right", ["Under ", ["Try It 4", "b"], ". Save and hover Menu."]],
    ["Two nav ul ul rules?", "That's fine: the browser combines them."],
    ["Play with it", ["Make the shadow bigger: ", ["0 8px 24px", "c"], ". Or change ", ["min-width", "c"], "."]],
  ],
  code: ["nav ul ul {", "   min-width: 220px;", "   background-color: white;", "   box-shadow: 0 4px 12px rgba(0, 0, 0, 0.25);", "}", "", "nav ul ul a {", "   color: #12304A;", "   font-weight: normal;", "}", "", "nav ul ul a:hover,", "nav ul ul a:focus-visible {", "   background-color: #FBE9DF;", "}"],
  tip: [[["box-shadow", "c"], ": across, down, blur, colour. ", ["rgba", "c"], " = a colour you can see through (0.25 = 25%)."]],
  notes: "Short and satisfying: the dropdown becomes a proper white panel with a shadow, and navy links.\n\nbox-shadow: 0 4px 12px rgba(0, 0, 0, 0.25) means: 0 across, 4px down, 12px of blur, in black at 25% opacity. A soft shadow just below makes it look like it's floating above the page.\n\nnav ul ul a is more specific than nav a, so the dropdown links are navy, not white. The hover rule overrides the teal from Try It 2 with peach.\n\nCheckpoint: everyone should now have a working menu bar with a dropdown. That's half of today's activity. Have them take screenshot 1 now (the dropdown open), while they're here.",
});

// 16. Divider: tables
dividerSlide("Tables", "For data, not layout   ·   structure   ·   styling   ·   zebra stripes   ·   small screens",
  "Tables are for data that has rows AND columns: prices, schedules, results, timetables. If you'd put it in a spreadsheet, it's probably a table.\n\nIn the 1990s, whole websites were laid out with tables. Don't do that! It's a nightmare for screen readers and phones. Today, CSS does layout, and tables are only for tabular data.");

// 17. Table parts
{
  const s = panelSlide("The Parts of a Table",
    "We met accessible tables in Class 4 (caption, th, scope). Here they are all together, because these are exactly the tags we'll style.\n\nThe nesting: a table contains rows (tr), and rows contain cells (th or td). There's no 'column' tag: columns are made by lining cells up, row after row. So every row needs the same number of cells.\n\nthead and tbody split the header rows from the data. They matter for CSS: 'tbody tr:nth-child(even)' stripes the data rows only.\n\nscope=\"col\" says 'this heading labels the column below it'. scope=\"row\" says 'this heading labels the row to its right'. With these, a screen reader can say 'Halibut, Price per lb, $24.50' instead of just '$24.50'.\n\ncaption is the table's name, and it's the first thing a screen reader announces.");
  // the diagram
  const gx = CX, gy = 2.35, cw = 1.35, rh = 0.48;
  txt(s, "Fresh Fish Prices", { x: gx, y: 1.95, w: 3 * cw, h: 0.38, fontSize: 14, bold: true, color: C.black });
  const rows = [["Fish", "Caught in", "Price / lb"], ["Haddock", "Yarmouth", "$11.99"], ["Halibut", "Digby", "$24.50"], ["Scallops", "Digby", "$29.99"]];
  rows.forEach((r, ri) => r.forEach((c, ci) => {
    const head = ri === 0, rowHead = ci === 0 && ri > 0;
    const fill = head ? C.navy : rowHead ? "F4F8FA" : C.white;
    s.addShape(pres.shapes.RECTANGLE, { x: gx + ci * cw, y: gy + ri * rh + (ri ? 0.1 : 0), w: cw, h: rh, fill: { color: fill }, line: { color: "B0B8BC", width: 0.75 } });
    txt(s, c, { x: gx + ci * cw + 0.1, y: gy + ri * rh + (ri ? 0.1 : 0), w: cw - 0.2, h: rh, fontSize: 12, bold: head || rowHead, color: head ? C.white : C.black, valign: "middle" });
  }));
  // brackets for thead / tbody
  const bx = gx + 3 * cw + 0.12;
  [[gy, rh, "<thead>", C.navy], [gy + rh + 0.1, 3 * rh, "<tbody>", C.teal]].forEach(([y, h, t, col]) => {
    s.addShape(pres.shapes.RECTANGLE, { x: bx, y: y + 0.04, w: 0.05, h: h - 0.08, fill: { color: col }, line: { color: col } });
    txt(s, t, { x: bx + 0.12, y, w: 0.9, h, fontSize: 12, bold: true, fontFace: MONO, color: col, valign: "middle" });
  });
  const parts = [
    ["<caption>", "The table's name: read out first"],
    ["<thead> <tbody>", "The header rows, and the data rows"],
    ["<tr>", "A table row"],
    ["<th scope=\"col\">", "A heading for the column below"],
    ["<th scope=\"row\">", "A heading for the row beside it"],
    ["<td>", "A table data cell: the actual data"],
  ];
  parts.forEach(([t, d], i) => {
    const y = 4.6 + i * 0.4;
    txt(s, t, { x: CX, y, w: 2.5, h: 0.36, fontSize: 13, bold: true, fontFace: MONO, color: "8A2B0B", valign: "middle" });
    txt(s, d, { x: CX + 2.55, y, w: CW - 2.55, h: 0.36, fontSize: 13.5, color: C.dark, valign: "middle" });
  });
}

// 18. Example 05: structure
{
  const L = bodyOf("05_table_structure_example.html");
  const third = L.reduce((acc, l, i) => (l.trim() === "</tr>" ? acc.concat(i) : acc), [])[2];
  const lines = L.slice(0, third + 1).concat(["      <!-- …two more rows… -->", "   </tbody>", "</table>"]);
  codeSlide({
    kicker: "TABLES 1 · EXAMPLE 05", title: "A Table, No CSS", file: "05_table_structure_example.html", lines, lang: "html", result: SH.ex05, size: 14,
    bullets: [
      ["The browser bolds and centres every ", ["th", "c"], ". There are no borders yet."],
      ["Write the ", ["structure first", "b"], ": caption, thead, tbody, th with scope. Then style it."],
    ],
    notes: "Open 05_table_structure_example.html. It works, and it's accessible, but it's hard to read: no lines, the columns are squashed together, and the headings are centred.\n\nNotice the browser already treats the th cells differently: bold, and centred in the header row. The first cell of each data row is a th too (scope=\"row\"), so the fish names are bold.\n\nAll of the next three examples use exactly this HTML. Only the CSS changes. That's the point of separating structure (HTML) from presentation (CSS).",
  });
}

// 19. Example 06: styled
codeSlide({
  kicker: "TABLES 2 · EXAMPLE 06", title: "Styling a Table", file: "06_table_styled_example.html",
  lines: between(styleOf("06_table_styled_example.html"), "/* collapse"), result: SH.ex06, size: 13.5,
  bullets: [
    [["border-collapse: collapse", "c"], ": one line between cells, not two."],
    ["Borders and padding go on the ", ["cells", "b"], " (", ["th, td", "c"], "), not the table."],
    ["Line numbers up on the ", ["right", "b"], ": ", ["td:last-child", "c"], " (a Class 9 selector)."],
  ],
  notes: "Open 06_table_styled_example.html, then comment out border-collapse: collapse and save. Every cell gets its own border with a gap in between: double lines everywhere. That's the browser default, and almost nobody wants it.\n\nPadding goes on th and td. Padding on the table itself only adds space around the outside.\n\ntext-align: left on th undoes the browser's centred headings, so headings line up with their data.\n\nRight-aligning prices is a typography habit: the decimal points line up, so it's easy to compare numbers down a column. th:last-child, td:last-child picks the last cell in every row, no classes needed.\n\nwidth: 100% makes the table fill its container. Without it, a table is only as wide as its content.",
});

// 20. Example 07: zebra
codeSlide({
  kicker: "TABLES 3 · EXAMPLE 07", title: "Zebra Stripes", file: "07_table_zebra_example.html",
  lines: between(styleOf("07_table_zebra_example.html"), "/* Lines between"), size: 13.5,
  frames: [{ img: SH.ex07, label: "RESULT (MOUSE OVER SCALLOPS)" }], cols: 1, maxFrameH: 3.5,
  bullets: [
    [["tbody tr:nth-child(even)", "c"], ": every second data row. Add a row and the stripes still work."],
    [["tbody", "c"], " so the header row isn't counted. ", ["tr:hover", "c"], " helps the eye follow a row."],
  ],
  notes: "Open 07_table_zebra_example.html and hover the rows.\n\nZebra stripes (alternating row colours) help your eye stay on the same row as it moves across a wide table. nth-child from Class 9 does it automatically: no classes on rows, and if you add or delete a row, the stripes fix themselves.\n\nWhy tbody tr and not just tr? Plain tr:nth-child(even) would also work here, because each tr is counted inside its own parent (thead or tbody). But writing tbody makes your intent clear, and keeps your header style safe if the stripe rule ever changes to odd.\n\nAlso new: borders on the bottom only. A full grid can look busy; lines between rows are cleaner and very common on real sites.\n\nKeep the stripe colour gentle. The text still needs 4.5 : 1 contrast on both colours.",
});

// 21. Example 08: responsive
{
  const css = between(styleOf("08_table_responsive_example.html"), "/* The wrapper", "th, td");
  const lines = css.concat(["", "<div class=\"table-wrap\" role=\"region\"", "     aria-label=\"Weekly specials\" tabindex=\"0\">", "   <table>", "      …", "   </table>", "</div>"]);
  codeSlide({
    kicker: "TABLES 4 · EXAMPLE 08", title: "Tables on Small Screens", file: "08_table_responsive_example.html", lines, lang: "mixed", size: 13.5,
    frames: [{ img: SH.ex08, label: "ON A PHONE: IT SCROLLS SIDEWAYS" }], cols: 1, maxFrameH: 3.2,
    bullets: [
      ["Wrap the table in a ", ["div", "c"], " with ", ["overflow-x: auto", "c"], ". The ", ["box", "b"], " scrolls, not the whole page."],
      [["tabindex=\"0\"", "c"], " + a label: keyboard users can scroll it with the arrow keys."],
      ["Or show it another way on phones, like cards (later, with media queries)."],
    ],
    notes: "Open 08_table_responsive_example.html and make the browser narrow (or use the DevTools device toolbar: Ctrl + Shift + M). The table keeps its width and the wrapper scrolls sideways. The rest of the page stays put.\n\nWithout the wrapper, a wide table pokes out of the page, and the WHOLE page scrolls sideways on a phone. That's WCAG 1.4.10 Reflow: content shouldn't need two-way scrolling at 320px wide (data tables are an allowed exception, but they should only scroll themselves).\n\nLast year's slide put overflow-x: auto and display: block on the table itself. That works visually, but changing a table's display can make some screen readers stop treating it as a table. The wrapper div is the safer way.\n\nrole=\"region\" + aria-label + tabindex=\"0\": the scrolling box can get keyboard focus (so the arrow keys scroll it), and a screen reader announces its name. Nice to have, not required today.\n\nmin-width: 700px stops the browser from squashing the columns down to nothing before it scrolls.\n\nThe other approach, turning each row into a 'card' on small screens, needs media queries, which come with responsive design.",
  });
}

// 22. Try It 5: hours table HTML
trySlide({
  n: 5, title: "An Hours Table", time: "10 min",
  steps: [
    ["Find the hours list", ["In index.html, in ", ["<aside id=\"hours\">", "c"], ". Delete the ", ["<ul>", "c"], "."]],
    ["Type the table on the right", "In its place. Save: an unstyled table appears."],
    ["Copy the row, six times", ["Select the Monday ", ["<tr>", "c"], " (5 lines), then ", ["Shift + Alt + ↓", "b"], " copies it."]],
    ["Fill in the days", "Mon–Thu: 11am to 8pm. Fri–Sat: 11am to 9pm. Sun: Noon to 7pm."],
  ],
  codeTitle: "IN INDEX.HTML", lang: "html",
  code: ["<div class=\"table-wrap\">", "   <table>", "      <caption>Opening hours</caption>", "      <thead>", "         <tr>", "            <th scope=\"col\">Day</th>", "            <th scope=\"col\">Open</th>", "            <th scope=\"col\">Close</th>", "         </tr>", "      </thead>", "      <tbody>", "         <tr>", "            <th scope=\"row\">Monday</th>", "            <td>11am</td>", "            <td>8pm</td>", "         </tr>", "         <!-- Tuesday to Sunday -->", "      </tbody>", "   </table>", "</div>"],
  notes: "The div with class table-wrap is the wrapper from Example 08. Our table is narrow, so it may never actually scroll, but wrapping every table is a good habit.\n\nVS Code shortcut: select the five lines of the Monday row, then Shift + Alt + Down Arrow copies them below. Press it six times, then change the days and times. Much faster than typing, and no typos in the tags.\n\nCommon mistakes:\n- A row with only two cells: the columns go out of line.\n- <th> for the day but forgetting scope=\"row\".\n- Rows typed outside the </tbody>.\n\nDon't remove the 'Try It 5' comment if you don't want to. Comments are fine.",
});

// 23. Try It 6: style the table
trySlide({
  n: 6, title: "Style the Table", time: "8 min",
  steps: [
    ["Type the CSS on the right", ["In style.css, under ", ["Try It 6", "b"], ". Save."]],
    ["Find the stripes", ["Tuesday, Thursday and Saturday: ", ["even", "c"], " rows."]],
    ["Play with it", ["Change ", ["even", "c"], " to ", ["odd", "c"], ". Then take out ", ["border-collapse", "c"], " and look closely."]],
  ],
  code: [".table-wrap { overflow-x: auto; }", "", "table {", "   width: 100%;", "   border-collapse: collapse;", "   background-color: white;", "}", "", "caption { font-weight: bold; text-align: left; }", "", "th, td {", "   padding: 8px 12px;", "   border-bottom: 1px solid #D0D7DB;", "   text-align: left;", "}", "", "thead th {", "   background-color: #12304A;", "   color: white;", "}", "", "tbody tr:nth-child(even) {", "   background-color: #E9EFF2;", "}"],
  notes: "Same recipe as Examples 06 and 07: collapse the borders, padding and lines on the cells, a navy header row, zebra stripes on the data rows.\n\nthead th, not just th: the day names are th too (scope=\"row\"). Plain th would make every day name white on navy as well.\n\nThe stripes land on Tuesday, Thursday and Saturday, rows 2, 4 and 6 of the tbody.\n\nCheckpoint: everyone should have a styled nav with a dropdown and a styled hours table. Next: is our HTML actually correct?",
});

// 24. Divider: validation
dividerSlide("Validation", "Browsers forgive   ·   validators don't   ·   HTML and CSS",
  "Here's a secret: browsers are very forgiving. If your HTML has a mistake, the browser quietly guesses what you meant and shows the page anyway. Most of the time it guesses right, so you never notice.\n\nThat's great for people browsing the web, but terrible for finding your own mistakes. A validator is a tool that doesn't forgive: it reads your HTML and lists every place it breaks the rules.");

// 25. What and why
{
  const s = panelSlide("What Is Validation, and Why Bother?",
    "What: checking your HTML against the official rules (the HTML standard, maintained by WHATWG and W3C). The W3C validator is a free website that does this.\n\nWhy, honestly:\n- The browser's guess isn't always the same guess. Different browsers, and especially screen readers, can repair broken HTML differently. Valid HTML means everyone gets the same page.\n- It finds bugs you can't see. A duplicate id breaks in-page links and form labels. A wrongly closed tag can swallow half your page.\n- Accessibility: missing alt text and broken nesting confuse assistive technology.\n- It's a professional habit. Clients and employers notice, and it's part of how your projects are marked in this program.\n\nA myth from last year's slide: validation doesn't directly change your Google ranking. Search engines read broken HTML fine. But a clean, well-structured page helps everything that reads it, search engines included.");
  const cw = (CW - 0.3) / 2;
  card(s, { x: CX, y: 1.95, w: cw, h: 2.1 });
  txt(s, "The browser forgives", { x: CX + 0.2, y: 2.05, w: cw - 0.4, h: 0.4, fontSize: 15, bold: true, color: C.black });
  codeBox(s, ["<h2>Lobster Roll</h3>"], { x: CX + 0.2, y: 2.55, w: cw - 0.4, h: 0.6, maxSize: 14 });
  txt(s, "Shows up fine. You never notice.", { x: CX + 0.2, y: 3.3, w: cw - 0.4, h: 0.6, fontSize: 13, color: C.grey, valign: "top" });
  const x2 = CX + cw + 0.3;
  card(s, { x: x2, y: 1.95, w: cw, h: 2.1, fill: "FDECEA", border: "F4B4AE" });
  txt(s, "The validator doesn't", { x: x2 + 0.2, y: 2.05, w: cw - 0.4, h: 0.4, fontSize: 15, bold: true, color: C.black });
  pill(s, "ERROR", x2 + 0.2, 2.6, C.red, 0.9);
  txt(s, "End tag “h3” seen, but there were open elements.", { x: x2 + 0.2, y: 3.05, w: cw - 0.4, h: 0.85, fontSize: 13, color: C.dark, valign: "top" });
  labelBullets(s, [
    { label: "Why bother, if it looks fine?", text: [
      "Browsers and screen readers can guess differently. Valid = the same page for everyone.",
      "It finds bugs you can't see: a duplicate id quietly breaks a link.",
      "Accessibility: missing alt text, tags in the wrong order.",
      "It's a professional habit, and part of how your projects are marked.",
    ] },
  ], { y: 4.3, h: 2.7, size: 17, sub: 14.5 });
}

// 26. How to use the validator
{
  const s = panelSlide("How to Use the W3C Validator",
    "validator.w3.org has three tabs across the top.\n\nValidate by URI (address): for pages that are already online. Ours aren't yet, so this one's for later in the program.\nValidate by File Upload: pick the .html file from your computer.\nValidate by Direct Input: paste the code into a box. This is the quickest today: in VS Code, Ctrl + A then Ctrl + C, and paste.\n\nThe results are a numbered list. Red 'Error' messages are rule-breakers: fix them. 'Warning' and 'Info' messages are advice: read them, and fix them if you agree. For example, a slash at the end of <img ... /> gets an info message. It's harmless.\n\nThere's a separate CSS validator at jigsaw.w3.org/css-validator, with the same three tabs. We'll use it after Try It 7.");
  codeBox(s, ["validator.w3.org"], { x: CX, y: 1.95, w: 4.2, h: 0.65, maxSize: 18 });
  const opts = [
    ["By URI", "A page that's already online", "Later in the program", C.mid],
    ["By File Upload", "Pick your .html file", "Works today", C.teal],
    ["By Direct Input", "Paste your code", "Quickest today!", C.orange],
  ];
  const cw = (CW - 0.4) / 3;
  opts.forEach(([t, d, n, col], i) => {
    const x = CX + i * (cw + 0.2), y = 2.85;
    card(s, { x, y, w: cw, h: 1.85, border: i === 2 ? C.orange : "D0D7DB", bw: i === 2 ? 2 : 0.75 });
    numCircle(s, i + 1, x + 0.2, y + 0.2, col, 0.42);
    txt(s, t, { x: x + 0.72, y: y + 0.18, w: cw - 0.85, h: 0.45, fontSize: 14.5, bold: true, color: C.black, valign: "middle" });
    txt(s, d, { x: x + 0.2, y: y + 0.8, w: cw - 0.4, h: 0.5, fontSize: 13, color: C.dark, valign: "top" });
    txt(s, n, { x: x + 0.2, y: y + 1.3, w: cw - 0.4, h: 0.4, fontSize: 12, italic: true, bold: i === 2, color: i === 2 ? "B5451B" : C.grey });
  });
  card(s, { x: CX, y: 4.95, w: CW, h: 1.55 });
  richBullets(s, [
    [["Error", "b"], " = breaks the rules. ", ["Fix it.", "b"]],
    [["Warning / Info", "b"], " = advice. Read it; fix it if you agree."],
    ["CSS has its own: ", ["jigsaw.w3.org/css-validator", "c"], " (after Try It 7)."],
  ], { x: CX + 0.25, y: 5.08, w: CW - 0.5, h: 1.35, size: 15, space: 10 });
}

// 27. Example 09: looks fine
codeSlide({
  kicker: "VALIDATION · EXAMPLE 09", title: "Looks Fine. Isn't.", file: "09_validation_broken_example.html", lines: bodyOf("09_validation_broken_example.html"), lang: "html", result: SH.ex09, size: 14,
  bullets: [
    ["This page has ", ["4 mistakes", "b"], ". Can you spot them before the validator does?"],
    ["The browser shows it ", ["perfectly", "b"], ". That's the problem!"],
  ],
  notes: "Open 09_validation_broken_example.html in the browser: it looks completely normal.\n\nGive the class a minute to spot the four mistakes in the code (don't reveal yet):\n1. The img has no alt attribute.\n2. <h2> is closed with </h3>.\n3. color=\"red\" isn't an HTML attribute (and the text isn't even red: the browser ignored it).\n4. <strong><em> … </strong></em>: closed in the wrong order.\n\nThen demo it live: validator.w3.org → Validate by Direct Input → paste the whole file → Check. The next slide shows what comes back.",
});

// 28. Reading the results
{
  const s = panelSlide("Reading the Results",
    "This is the real output for Example 09 (copied from validator.w3.org). Four mistakes, but FIVE error messages: the wrong nesting on line 21 creates two. One mistake often causes a little chain of messages.\n\nSo the method is:\n1. Read the FIRST error.\n2. Go to that line in VS Code (Ctrl + G, then the line number).\n3. Fix it, save, and validate again.\n4. Repeat until it says 'Document checking completed. No errors or warnings to show.'\n\nOften fixing one error makes others disappear too.\n\nThe messages use the official names, so they can sound scary. 'End tag h3 seen, but there were open elements' just means 'you closed an h3, but the h2 you opened is still open'.\n\nExample 10 is the fixed version, with a comment on each fix.");
  const msgs = [
    ["16", "An “img” element must have an “alt” attribute…", "Add alt=\"\" (decoration) or alt=\"description\""],
    ["19", "End tag “h3” seen, but there were open elements.", "Close it with the same tag: </h2>"],
    ["20", "Attribute “color” not allowed on element “p” at this point.", "Not an HTML attribute: do it in CSS"],
    ["21", "End tag “strong” violates nesting rules.", "Last opened, first closed:"],
    ["21", "No “em” element in scope but a “em” end tag seen.", "<strong><em>…</em></strong>"],
  ];
  msgs.forEach(([line, m, fix], i) => {
    const y = 1.95 + i * 0.82;
    card(s, { x: CX, y, w: CW, h: 0.72, fill: i >= 3 ? "FFF7F6" : C.white });
    pill(s, "ERROR", CX + 0.15, y + 0.1, C.red, 0.85);
    txt(s, "line " + line, { x: CX + 0.15, y: y + 0.44, w: 0.85, h: 0.25, fontSize: 10.5, fontFace: MONO, color: C.grey, align: "center" });
    txt(s, m, { x: CX + 1.15, y, w: 3.7, h: 0.72, fontSize: 12, color: C.black, valign: "middle" });
    txt(s, fix, { x: CX + 4.95, y, w: CW - 5.05, h: 0.72, fontSize: 12, color: C.teal, bold: true, valign: "middle", fontFace: fix.startsWith("<") ? MONO : FONT });
  });
  txt(s, "↑ one mistake on line 21, two messages", { x: CX + 1.15, y: 1.95 + 5 * 0.82 - 0.08, w: 3.7, h: 0.3, fontSize: 11, italic: true, color: C.red });
  card(s, { x: CX, y: 6.3, w: CW, h: 0.65, fill: C.dark, border: C.dark });
  richBullets(s, [[["Fix the first error, save, check again.", "b"], " Repeat until: “No errors or warnings to show.”"]],
    { x: CX + 0.25, y: 6.3, w: CW - 0.5, h: 0.65, size: 14, bullet: false, color: C.white, valign: "middle" });
}

// 29. Try It 7: validate Harbourside
{
  const s = activitySlide("Try It 7: Validate the HTML (12 min)",
    "The surprise: the starter index.html had five mistakes in it from the start. Not in anything students typed today (hopefully!), but in the page they were given.\n\nThe five, for you (all fixed and commented in activity/answer/index.html):\n1. Line ~15: the header logo <img> has no alt. It's decoration (the h1 says the name), so alt=\"\".\n2. Line ~30: <strong>…<em>boats</strong></em>, closed in the wrong order (two messages).\n3. Line ~44: <h3>Clam Strips</h2>.\n4. Line ~55: Duplicate ID “clams”. The Kids' Fish Bites article is id=\"clams\" (copy-and-paste!). It should be id=\"kids\", and that's why the dropdown's Kids link didn't work in Try It 3.\n5. Line ~77: <p font=\"small\">. Not an attribute: use class=\"small\" and CSS.\nLine numbers move a little, depending on what each student typed.\n\nIf students get errors in their own nav or table code, even better: they fix those too. Common ones: a missing </li>, a <ul> outside its <li>, or a missing </tr>.\n\nThe HTML validator doesn't open style.css, so it can't see the CSS mistakes. That's next, in Try It 8.");
  const LW = 6.55;
  card(s, { x: 0.75, y: 1.45, w: LW, h: 5.5 });
  const steps = [
    ["Open validator.w3.org", ["Choose ", ["Validate by Direct Input", "b"], "."]],
    ["Copy all of index.html", ["In VS Code: ", ["Ctrl + A", "b"], ", ", ["Ctrl + C", "b"], ". Paste it in, then ", ["Check", "b"], "."]],
    ["Fix the first error", ["Ctrl + G jumps to the line. Fix it and save."]],
    ["Copy, paste, check again", "Repeat until there are no errors left."],
    ["Screenshot the result", ["“Document checking completed. ", ["No errors or warnings to show.", "b"], "”"]],
  ];
  steps.forEach(([t, d], i) => {
    const y = 1.72 + i * 1.0;
    numCircle(s, i + 1, 1.0, y, C.orange, 0.44);
    txt(s, t, { x: 1.6, y: y - 0.02, w: LW - 1.05, h: 0.35, fontSize: 15.5, bold: true, color: C.black });
    richBullets(s, [typeof d === "string" ? [d] : d], { x: 1.6, y: y + 0.32, w: LW - 1.05, h: 0.64, size: 13, bullet: false, color: C.dark });
  });
  const RX = 0.75 + LW + 0.35, RW = W - RX - 0.6;
  card(s, { x: RX, y: 1.45, w: RW, h: 3.9, fill: C.dark, border: C.dark });
  txt(s, "The starter page had 5 hidden mistakes!", { x: RX + 0.3, y: 1.6, w: RW - 0.6, h: 0.5, fontSize: 17, bold: true, color: C.white });
  txt(s, "Clues:", { x: RX + 0.3, y: 2.1, w: RW - 0.6, h: 0.35, fontSize: 13, color: "F2C6AC", bold: true });
  const clues = [
    "An image that says nothing about itself",
    "Two tags closed in the wrong order",
    "A closing tag that doesn't match its opening tag",
    "Two of something there should only be one of",
    "An attribute HTML has never heard of",
  ];
  clues.forEach((c, i) => {
    const y = 2.5 + i * 0.55;
    numCircle(s, i + 1, RX + 0.3, y, C.orange, 0.36);
    txt(s, c, { x: RX + 0.8, y: y - 0.06, w: RW - 1.1, h: 0.48, fontSize: 13.5, color: "D9D9D9", valign: "middle" });
  });
  card(s, { x: RX, y: 5.55, w: RW, h: 1.4, fill: "FBE9DF", border: "F2C6AC" });
  richBullets(s, [
    ["Clue 4 explains why the ", ["Kids' Fish Bites", "b"], " link in your dropdown doesn't go anywhere!"],
    ["Errors in the code ", ["you", "b"], " typed? Fix those too."],
  ], { x: RX + 0.25, y: 5.6, w: RW - 0.5, h: 1.3, size: 13, bullet: false, valign: "middle", space: 6 });
}

// 30. Example 11: CSS fails quietly
codeSlide({
  kicker: "CSS VALIDATION · EXAMPLE 11", title: "CSS Fails Quietly", file: "11_css_validation_broken_example.html",
  lines: between(styleOf("11_css_validation_broken_example.html"), "h1 {"), size: 13.5,
  frames: [{ img: SH.ex11, label: "WHAT YOU GET" }, { img: SH.ex12, label: "WHAT YOU MEANT (EXAMPLE 12)" }], cols: 2, maxFrameH: 2.4,
  bullets: [
    ["The browser ", ["skips", "b"], " any CSS it doesn't understand. No error, no warning: it just doesn't happen."],
    ["4 mistakes: the heading isn't navy, the text isn't bigger, the price isn't orange, the note isn't centred."],
    ["The HTML validator checks a ", ["<style>", "c"], " block, but ", ["never opens", "b"], " your linked ", ["style.css", "c"], "."],
  ],
  notes: "Open 11_css_validation_broken_example.html. It looks like a plain, unstyled page, and nothing tells you why. Open DevTools (F12) and look at the Styles panel: the broken declarations are crossed out, with a little warning triangle. That's the browser's only hint.\n\nGive the class a minute to spot the four mistakes:\n1. colour (British spelling). CSS only knows color.\n2. font-size: 18 with no unit. Only 0 can go without a unit.\n3. A missing semicolon after font-weight: bold. The browser reads 'bold color: #B5451B' as one broken value, so BOTH lines are lost: the price is neither bold nor orange.\n4. text-align: middle. The word is center.\n\nThen demo it live: jigsaw.w3.org/css-validator → By direct input → paste the whole file → Check. The CSS validator accepts a whole HTML file and checks its <style> block. The next slide shows what comes back.\n\nWhy we need a second validator: validator.w3.org does check CSS inside a <style> block (paste Example 11 into it and you get the same four errors). But it never downloads linked stylesheets. Our Harbourside CSS lives in style.css, so only the CSS validator can check it.\n\nRight-hand picture: Example 12, the fixed version (each fix has a comment).",
});

// 31. Reading the CSS results
{
  const s = panelSlide("The W3C CSS Validator",
    "This is the real output for Example 11 (from jigsaw.w3.org/css-validator, October 2026). One message per mistake this time, and they're friendlier than the HTML ones: the first even suggests the right spelling.\n\nThe page has the same three tabs as the HTML validator: By URI, By file upload, By direct input. Direct input is quickest: Ctrl + A, Ctrl + C in style.css, paste, Check.\n\nEach message gives the line number and the selector it's in (h1, p, .price, .note), which makes it easy to find.\n\nWhen it's clean, the result says 'Congratulations! No Error Found.' (with a green banner). That's screenshot 3.\n\nWarnings: the CSS validator sometimes warns about things that are fine, like vendor prefixes (-webkit-…). Errors matter; warnings are advice, same as with HTML.");
  codeBox(s, ["jigsaw.w3.org/css-validator"], { x: CX, y: 1.95, w: 4.9, h: 0.6, maxSize: 17 });
  txt(s, "→ By direct input", { x: CX + 5.05, y: 1.95, w: CW - 5.05, h: 0.6, fontSize: 15, bold: true, color: C.teal, valign: "middle" });
  const msgs = [
    ["14", "h1", "Property “colour” doesn't exist. The closest matching property name is “color”", "Spell it color"],
    ["18", "p", "only “0” can be a “unit”. You must put a unit after your number", "18px (or 1.1rem)"],
    ["23", ".price", "Missing a semicolon before the property name “color”", "bold;  ← it broke both lines"],
    ["27", ".note", "“middle” is not a “text-align” value", "text-align: center"],
  ];
  msgs.forEach(([line, sel, m, fix], i) => {
    const y = 2.75 + i * 0.86;
    card(s, { x: CX, y, w: CW, h: 0.76 });
    pill(s, "ERROR", CX + 0.15, y + 0.1, C.red, 0.85);
    txt(s, "line " + line, { x: CX + 0.15, y: y + 0.46, w: 0.85, h: 0.25, fontSize: 10.5, fontFace: MONO, color: C.grey, align: "center" });
    txt(s, [{ text: sel + "  ", options: { bold: true, fontFace: MONO, color: "8A2B0B" } }, { text: m }], { x: CX + 1.15, y, w: 4.0, h: 0.76, fontSize: 11.5, color: C.black, valign: "middle" });
    txt(s, fix, { x: CX + 5.3, y, w: CW - 5.4, h: 0.76, fontSize: 12, color: C.teal, bold: true, valign: "middle" });
  });
  card(s, { x: CX, y: 6.3, w: CW, h: 0.65, fill: C.green, border: C.green });
  richBullets(s, [[["When it's clean: ", "b"], "“Congratulations! No Error Found.”"]],
    { x: CX + 0.25, y: 6.3, w: CW - 0.5, h: 0.65, size: 14, bullet: false, color: C.white, valign: "middle" });
}

// 32. Try It 8: validate the CSS
{
  const s = activitySlide("Try It 8: Validate the CSS (6 min)",
    "The second surprise: the Class 8 part of the starter style.css has three mistakes in it. They've been there all class, and nobody noticed, because the browser just skipped them.\n\nThe three, for you (fixed and commented as CSS Fix 1–3 in activity/answer/style.css):\n1. Line ~50: main { padding: 20; }, no unit, so main has no padding at all.\n2. Line ~70: figcaption { colour: #595959; }, so the caption is black instead of grey.\n3. Line ~74: .price { font-weight: bold  (no semicolon) color: #B5451B; }, so the prices are neither bold nor orange.\nLine numbers move depending on what each student typed below.\n\nAfter fixing, the page visibly changes: orange bold prices, a grey caption, and the main column gets its padding back. Ask: 'Did anyone notice the prices were black all class?'\n\nAlso possible: errors in the CSS students typed today, e.g. a missing } or a typo in a property. They fix those too.\n\nIf time is short, this one is fine as homework.");
  const LW = 6.55;
  card(s, { x: 0.75, y: 1.45, w: LW, h: 5.5 });
  const steps = [
    ["Open jigsaw.w3.org/css-validator", ["Choose ", ["By direct input", "b"], "."]],
    ["Copy all of style.css", ["In VS Code: ", ["Ctrl + A", "b"], ", ", ["Ctrl + C", "b"], ". Paste it in, then ", ["Check", "b"], "."]],
    ["Fix each error", "The message names the line and the selector."],
    ["Copy, paste, check again", "Until there are no errors left."],
    ["Screenshot the result", ["“Congratulations! ", ["No Error Found.", "b"], "”"]],
  ];
  steps.forEach(([t, d], i) => {
    const y = 1.72 + i * 1.0;
    numCircle(s, i + 1, 1.0, y, C.green, 0.44);
    txt(s, t, { x: 1.6, y: y - 0.02, w: LW - 1.05, h: 0.35, fontSize: 15.5, bold: true, color: C.black });
    richBullets(s, [typeof d === "string" ? [d] : d], { x: 1.6, y: y + 0.32, w: LW - 1.05, h: 0.64, size: 13, bullet: false, color: C.dark });
  });
  const RX = 0.75 + LW + 0.35, RW = W - RX - 0.6;
  card(s, { x: RX, y: 1.45, w: RW, h: 3.3, fill: C.dark, border: C.dark });
  txt(s, "Our Class 8 CSS had 3 hidden mistakes!", { x: RX + 0.3, y: 1.6, w: RW - 0.6, h: 0.5, fontSize: 17, bold: true, color: C.white });
  txt(s, "Clues:", { x: RX + 0.3, y: 2.1, w: RW - 0.6, h: 0.35, fontSize: 13, color: "F2C6AC", bold: true });
  ["A number with no unit", "A word CSS spells differently", "A missing semicolon (it breaks two lines!)"].forEach((c, i) => {
    const y = 2.55 + i * 0.6;
    numCircle(s, i + 1, RX + 0.3, y, C.green, 0.36);
    txt(s, c, { x: RX + 0.8, y: y - 0.06, w: RW - 1.1, h: 0.48, fontSize: 13.5, color: "D9D9D9", valign: "middle" });
  });
  card(s, { x: RX, y: 4.95, w: RW, h: 2.0, fill: "FBE9DF", border: "F2C6AC" });
  richBullets(s, [
    [["Watch the page as you fix them.", "b"], " Something changes every time."],
    ["Did you notice the prices were ", ["plain black", "b"], " all class?"],
    ["The HTML validator ", ["can't", "b"], " find these: it never opens style.css."],
  ], { x: RX + 0.25, y: 5.0, w: RW - 0.5, h: 1.9, size: 13, bullet: false, valign: "middle", space: 6 });
}

// 33. Make it yours + submit
{
  const s = activitySlide("Make It Yours, Then Submit",
    "Anyone who has a working dropdown, a styled table and two clean validator results has done the job. The ideas are for fast finishers, or homework.\n\nRe-validate after making changes! It's easy to break something while adding to it.\n\nWhat to collect on Brightspace (participation): three screenshots. 1: the page in the browser with the Menu dropdown open (hover it, then Win + Shift + S. The Snipping Tool waits for you to pick an area, so the dropdown stays open). 2: the HTML validator showing no errors. 3: the CSS validator showing 'No Error Found'.");
  card(s, { x: 0.75, y: 1.45, w: 7.4, h: 5.5 });
  txt(s, "Make it yours", { x: 1.0, y: 1.65, w: 6.9, h: 0.45, fontSize: 18, bold: true, color: C.black });
  const ideas = [
    [["Highlight the row under the mouse: ", ["tbody tr:hover", "c"]]],
    [["Bold Friday to Sunday: ", ["tbody tr:nth-child(n+5)", "c"]]],
    [["Give the dropdown ", ["rounded corners", "b"], ": ", ["border-radius", "c"], " (try ", ["0 0 8px 8px", "c"], ")"]],
    [["A new colour for the ", ["current page", "b"], " underline"]],
    [["Line up the Open and Close times on the right: ", ["td", "c"], " + ", ["text-align", "c"]]],
  ];
  ideas.forEach((segs, i) => {
    const y = 2.35 + i * 0.85;
    numCircle(s, i + 1, 1.0, y, C.orange, 0.4);
    richBullets(s, segs, { x: 1.55, y: y - 0.12, w: 6.4, h: 0.65, size: 14.5, bullet: false, color: C.dark, valign: "middle" });
  });
  const x = 8.4, w = W - x - 0.6;
  card(s, { x, y: 1.45, w, h: 5.5, fill: C.dark, border: C.dark });
  txt(s, "Submit to Brightspace", { x: x + 0.25, y: 1.7, w: w - 0.5, h: 0.75, fontSize: 18, bold: true, color: C.white, valign: "top" });
  txt(s, [
    { text: "Three screenshots:", options: { bold: true, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "1. Your page in the browser, with the Menu dropdown open", options: { breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "2. The HTML validator: “No errors or warnings to show.”", options: { breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "3. The CSS validator: “No Error Found.”", options: { breakLine: true } },
    { text: " ", options: { fontSize: 10, breakLine: true } },
    { text: "Win + Shift + S", options: { bold: true, color: "F2C6AC", fontFace: MONO } },
    { text: ": hover the menu first, then pick the area.", options: { breakLine: true } },
    { text: " ", options: { fontSize: 10, breakLine: true } },
    { text: "Stuck? ", options: { bold: true, color: "F2C6AC" } },
    { text: "Peek at activity/answer/. Every fix has a comment." },
  ], { x: x + 0.25, y: 2.5, w: w - 0.5, h: 4.3, fontSize: 13.5, color: "D9D9D9", valign: "top" });
}

// 31. Cheat sheet: nav and tables
{
  const s = panelSlide("Cheat Sheet: Menus and Tables",
    "Everything from today's CSS in one place. It's in the README too.");
  table(s, [
    ["Code", "What it does"],
    ["nav ul { list-style: none; padding: 0; }", "Remove the bullets and the indent"],
    ["nav ul { display: flex; gap: 4px; }", "Links side by side, with space between"],
    ["nav a { display: block; padding: … }", "The whole box is clickable"],
    ["a[aria-current=\"page\"]", "Style the \"you are here\" link"],
    ["li { position: relative; }", "Dropdowns are placed against their li"],
    ["ul ul { display: none; position: absolute; }", "Hide the dropdown; float it over the page"],
    ["li:hover > ul, li:focus-within > ul", "Show it for the mouse and the keyboard"],
    ["border-collapse: collapse", "One line between table cells, not two"],
    ["tbody tr:nth-child(even)", "Zebra stripes on the data rows"],
    [".table-wrap { overflow-x: auto; }", "A wide table scrolls in its own box"],
  ], { x: CX, y: 1.9, w: CW, colW: [4.15, 3.4], size: 12.5, rowH: 0.44, codeCols: [0] });
}

// 32. Cheat sheet: validation
{
  const s = panelSlide("Cheat Sheet: Validator Messages",
    "The messages students will see most often this term, and what they really mean. Top: validator.w3.org (HTML). Bottom: jigsaw.w3.org/css-validator (CSS).\n\nTwo more HTML ones worth knowing: 'Element li not allowed as child of element nav' (an li must be inside a ul or ol) and 'Stray end tag div' (a closing tag with no opening tag).");
  table(s, [
    ["The HTML validator says…", "It means… / the fix"],
    ["An “img” element must have an “alt” attribute", "Add alt=\"description\", or alt=\"\" for decoration"],
    ["End tag “h3” seen, but there were open elements", "The closing tag doesn't match the opening tag"],
    ["Unclosed element “section”", "A closing tag is missing"],
    ["End tag “strong” violates nesting rules", "Last opened, first closed: <b><i>…</i></b>"],
    ["Duplicate ID “clams”", "Every id must be unique on the page"],
    ["Attribute “color” not allowed on element “p”", "Not an HTML attribute: do it in CSS"],
  ], { x: CX, y: 1.9, w: CW, colW: [4.0, 3.55], size: 12, rowH: 0.4, codeCols: [] });
  table(s, [
    ["The CSS validator says…", "It means… / the fix"],
    ["Property “colour” doesn't exist", "Spelling: CSS uses color (and center)"],
    ["only “0” can be a “unit”…", "Add a unit: 20px, 1.2rem"],
    ["Missing a semicolon before the property name…", "End every declaration with ;"],
    ["“middle” is not a “text-align” value", "Not an allowed value: check MDN"],
  ], { x: CX, y: 5.0, w: CW, colW: [4.0, 3.55], size: 12, rowH: 0.4, codeCols: [] });
}

// 33. Recap
{
  const s = panelSlide("Class Recap and Next Steps",
    "Next class: page layout. Today's display: flex was a sneak peek; next time we look at display, position and Flexbox properly, and use them to put the Harbourside menu cards side by side.\n\nReminder: no class Wednesday, October 21 (Assessment & Evaluation day). Make sure the three screenshots are on Brightspace.");
  labelBullets(s, [
    { label: "Recap:", text: [
      "Every menu is nav → ul → li → a. CSS makes it horizontal, vertical or a dropdown.",
      "Make the whole link clickable (display: block + padding), and open dropdowns with :focus-within too.",
      "Tables are for data: caption, thead, tbody, th with scope, td.",
      "border-collapse, padding on cells, zebra stripes with nth-child, and a scrolling wrapper for small screens.",
      "Browsers forgive broken HTML. The W3C validator doesn't: fix the first error, then check again.",
      "Browsers silently skip broken CSS. The W3C CSS validator checks style.css (the HTML one can't).",
    ] },
    { label: "This week:", text: [
      "Submit your three screenshots to Brightspace. No class Wednesday, October 21.",
    ] },
    { label: "Next class:", text: [
      "Page layout: display, position, and Flexbox.",
    ] },
  ], { size: 20, sub: 15.5 });
}

pres.writeFile({ fileName: OUT }).then((f) => console.log("wrote " + f));
