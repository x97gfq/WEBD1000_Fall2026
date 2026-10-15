// Builds Selectors_and_Animation.pptx (Class 9) in the same style as the Class 5–8 decks.
// To rebuild: cd 06_Selectors_and_Animation/_build, run "npm install" once, then "npm run build".
// The "Result" pictures are screenshots of the example pages, taken with Microsoft Edge (headless) each build.
// Code slides read the example files directly, so edit those files and re-run to update the slides.
const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");
const { execFileSync } = require("child_process");

const REPO = path.join(__dirname, ".."); // the 06_Selectors_and_Animation folder
const SCR = __dirname;
const SHOTS = path.join(SCR, "shots");
const OUT = path.join(REPO, "Selectors_and_Animation.pptx");

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
pres.title = "Class 9: Advanced Selectors and CSS Animation";
pres.author = "Jamie Symonds";

const W = 13.333, H = 7.5;
const PANEL_W = 4.1;
const CX = 5.2, CW = 7.55; // content column on panel slides

// ---------- helpers (same as the Class 5–8 decks) ----------
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


// A copy of an example page, with a <base> so its links still work, and extra code injected before </body>:
//   t: pause every animation at t seconds (Web Animations API), so a screenshot shows one frame
//   hover: a selector; its :hover rules are applied as if the mouse were over it
function frame(file, name, { t = null, hover = null, w = 600, h = 400, scale = 2 } = {}) {
  const abs = path.join(REPO, file);
  const base = "file:///" + path.dirname(abs).replace(/\\/g, "/") + "/";
  let inject = "";
  if (hover) inject += `<style>* { transition: none !important; }</style><script>
    for (const sheet of document.styleSheets) for (const r of [...sheet.cssRules])
      if (r.selectorText && r.selectorText.replace(/:hover/g, "") === ${JSON.stringify(hover)} && r.selectorText.includes(":hover"))
        sheet.insertRule(r.cssText.replace(/:hover/g, ".sim-hover"), sheet.cssRules.length);
    document.querySelectorAll(${JSON.stringify(hover)}).forEach((e) => e.classList.add("sim-hover"));
  </script>`;
  if (t !== null) inject += `<script>document.getAnimations().forEach((a) => { a.pause(); a.currentTime = ${t * 1000}; });</script>`;
  const src = fs.readFileSync(abs, "utf8").replace("<head>", `<head><base href="${base}">`).replace("</body>", inject + "</body>");
  const tmp = path.join(SHOTS, name + ".html");
  fs.writeFileSync(tmp, src);
  return shot(path.relative(REPO, tmp), name, w, h, scale);
}

// ---------- screenshots ----------
const SH = {
  ex01: shot("01_pseudo_classes_example.html", "ex01", 600, 170),
  ex02: shot("02_structural_example.html", "ex02", 600, 400),
  ex03: shot("03_attribute_example.html", "ex03", 600, 280),
  ex04: shot("04_combinators_example.html", "ex04", 600, 400),
  ex05: shot("05_pseudo_elements_example.html", "ex05", 600, 250),
  ex06a: frame("06_transition_example.html", "ex06a", { w: 500, h: 230 }),
  ex06b: frame("06_transition_example.html", "ex06b", { hover: ".card", w: 500, h: 230 }),
  ex07: [0, 1, 2, 3].map((t) => frame("07_ball_around_example.html", "ex07_" + t, { t, w: 600, h: 400, scale: 1 })),
  ex08: frame("08_timing_race_example.html", "ex08", { t: 1, w: 600, h: 440 }),
  ex09: [0, 0.4, 0.6, 0.85].map((t, i) => frame("09_bounce_example.html", "ex09_" + i, { t, w: 500, h: 430, scale: 1 })),
  ex10: [0, 0.75, 1.5, 2.25].map((t, i) => frame("10_morph_square_example.html", "ex10_" + i, { t, w: 500, h: 400, scale: 1 })),
  ex11: [0.9, 1.6, 2.4].map((t, i) => frame("11_car_example.html", "ex11_" + i, { t, w: 600, h: 190, scale: 1 })),
  ex12: frame("12_real_world_example.html", "ex12", { t: 0.5, w: 600, h: 250 }),
  playground: frame("activity/answer/index.html", "playground", { t: 0.6, w: 700, h: 1000, scale: 1 }),
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
    { text: "Class 9:", options: { bullet: true, fontSize: 30, breakLine: true } },
    { text: "Advanced selectors, and making things move with CSS animation.", options: { bullet: { code: "25CB" }, indentLevel: 1, fontSize: 26, paraSpaceBefore: 6 } },
  ], { x: 6.85, y: 1.9, w: 5.6, h: 2.4, color: C.black, valign: "top" });
  s.addNotes("Welcome back! Last class we styled the Harbourside page from scratch: selectors, colour, fonts, the box model and the cascade.\n\nToday has two halves. First, some more powerful selectors, so we can pick out exactly the element we want (and play a quick game). Then the fun part: CSS animation. By the end of class, everyone will have a page full of rolling balls, spinning pinwheels and pulsing boxes.");
}

// 2. Attendance
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  txt(s, "Attendance", { x: 1.1, y: 0.75, w: 8, h: 0.9, fontSize: 44, color: C.black });
  s.addNotes("Take attendance. While you do, have students git pull and open the 06_Selectors_and_Animation folder in VS Code.");
}

// 3. Overview
{
  const s = panelSlide("Class 9: Class Overview",
    "Rough plan: 5 min intro, about 35 min of selectors (five quick examples and a game), about 30 min on animation concepts with demos, then about 40 min in the Animation Playground, and 10 min of show and tell at the end.\n\nThe selector half is quicker and lighter than last class: no Try Its, just examples to open and poke at in DevTools. Save the energy for the playground.");
  txt(s, [
    { text: "Title: ", options: { bold: true, bullet: true, fontSize: 22 } },
    { text: "Advanced Selectors and CSS Animation", options: { fontSize: 22, breakLine: true } },
    { text: " ", options: { fontSize: 10, breakLine: true } },
    { text: "Agenda:", options: { bold: true, bullet: true, fontSize: 22, breakLine: true } },
    { text: "Advanced selectors, and a game: Name That Selector!", options: { bullet: true, indentLevel: 1, fontSize: 18, paraSpaceBefore: 4, breakLine: true } },
    { text: "states · positions · attributes · relationships · parts of an element", options: { bullet: true, indentLevel: 2, fontSize: 15, paraSpaceBefore: 2, breakLine: true, color: C.grey } },
    { text: "CSS animation: transitions, @keyframes, and lots of demos", options: { bullet: true, indentLevel: 1, fontSize: 18, paraSpaceBefore: 4, breakLine: true } },
    { text: "The Animation Playground: four challenges", options: { bullet: true, indentLevel: 1, fontSize: 18, paraSpaceBefore: 4, breakLine: true } },
    { text: "Show and tell", options: { bullet: true, indentLevel: 1, fontSize: 18, paraSpaceBefore: 4 } },
  ], { x: CX, y: 1.95, w: CW, h: 3.6, color: C.black, valign: "top" });
  const steps = [["Select", "Pick exactly the element you want"], ["Animate", "Make it move with @keyframes"], ["Play", "Four challenges, lots of motion"]];
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
    "Examples 01–05 are selectors, 06–12 are animation. Every example is a single file with its CSS in a <style> block, so students can open one and change things straight away.\n\nThe activity folder has two CSS files. stage.css draws the stages and objects (they shouldn't need to touch it). style.css is theirs, with a heading for each challenge.\n\nAsk students not to open activity/answer until they're stuck.");
  txt(s, [{ text: "1.  Pull this week's examples", options: { bold: true, fontSize: 18 } }], { x: CX, y: 1.95, w: CW, h: 0.4 });
  codeBox(s, ["git pull"], { x: CX, y: 2.45, w: 5.2, h: 0.7, maxSize: 16 });
  txt(s, "New folder: 06_Selectors_and_Animation (examples 01–12, and the activity folder).", { x: CX, y: 3.25, w: CW, h: 0.5, fontSize: 15, color: C.dark });
  txt(s, [{ text: "2.  Open each example with Live Server as we go", options: { bold: true, fontSize: 18 } }], { x: CX, y: 3.95, w: CW, h: 0.4 });
  txt(s, "Animations only move in a real browser, so open them! Change a number, save, and watch what happens.", { x: CX, y: 4.4, w: CW, h: 0.8, fontSize: 15, color: C.dark });
  txt(s, [{ text: "3.  Keep DevTools handy", options: { bold: true, fontSize: 18 } }], { x: CX, y: 5.4, w: CW, h: 0.4 });
  txt(s, "F12 → Elements → Styles shows which rule picked each element. There's an Animations panel too (later).", { x: CX, y: 5.85, w: CW, h: 0.8, fontSize: 15, color: C.dark });
}

// 5. Divider
dividerSlide("Advanced Selectors", "States   ·   positions   ·   attributes   ·   relationships   ·   parts",
  "Last class's selectors (element, .class, #id, groups and descendants) get you a long way. Today's selectors let you pick things out without adding a class to every element: the third row of a table, every link to a PDF, the paragraph right after a heading.\n\nThe pace here is quick: one example per selector family, then the game.");

// 6. Selector families map
{
  const s = panelSlide("Five New Kinds of Selector",
    "This is the map for the next five slides. Each family has a symbol to look for.\n\n: (one colon) = a pseudo-class. It matches an element in a certain state or position.\n[ ] = an attribute selector. It matches by what's inside the HTML tag.\n> + ~ = combinators. They match by how elements are related in the HTML.\n:: (two colons) = a pseudo-element. It styles a PART of an element, or adds content that isn't in the HTML at all.\n\nYou don't need to memorize them. You need to know they exist, so you can look them up (MDN is the best reference).");
  txt(s, "Last class:", { x: CX, y: 1.9, w: 1.6, h: 0.4, fontSize: 15, bold: true, color: C.grey });
  ["p", ".price", "#menu", "h1, h2", "nav a"].forEach((t, i) => {
    pill(s, t, CX + 1.65 + i * 1.15, 1.93, C.mid, 1.0);
  });
  const fams = [
    [":", "State", "a:hover   button:focus-visible", "What's happening to it right now", C.orange],
    [":nth", "Position", "tr:nth-child(even)   li:not(.sold-out)", "Where it sits inside its parent", C.blue],
    ["[ ]", "Attribute", "a[href$=\".pdf\"]   input[type=\"email\"]", "What's written inside its tag", C.purple],
    ["> + ~", "Relationship", "ul > li   h2 + p   h3 ~ p", "Who its parent or neighbours are", C.green],
    ["::", "Part", "p::first-letter   .new::before", "Part of an element, or added content", C.red],
  ];
  fams.forEach(([sym, t, ex, d, col], i) => {
    const y = 2.55 + i * 0.88;
    card(s, { x: CX, y, w: CW, h: 0.76 });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: CX + 0.12, y: y + 0.11, w: 0.95, h: 0.54, rectRadius: 0.08, fill: { color: col }, line: { color: col } });
    txt(s, sym, { x: CX + 0.12, y: y + 0.11, w: 0.95, h: 0.54, fontSize: 15, bold: true, fontFace: MONO, color: C.white, align: "center", valign: "middle" });
    txt(s, t, { x: CX + 1.25, y: y + 0.06, w: 1.6, h: 0.36, fontSize: 15.5, bold: true, color: C.black });
    txt(s, d, { x: CX + 1.25, y: y + 0.4, w: 2.7, h: 0.32, fontSize: 11.5, color: C.grey });
    txt(s, ex, { x: CX + 3.95, y, w: CW - 4.05, h: 0.76, fontSize: 11.5, fontFace: MONO, color: "8A2B0B", valign: "middle" });
  });
}

// 7. Pseudo-classes (Example 01)
codeSlide({
  kicker: "SELECTORS 1 · STATE", title: "Pseudo-classes: :hover, :focus-visible, :active", file: "01_pseudo_classes_example.html", lines: styleOf("01_pseudo_classes_example.html"), result: SH.ex01, size: 13.5,
  bullets: [
    ["A ", ["pseudo-class", "b"], " (one colon) matches an element in a ", ["state", "b"], ": ", [":hover", "c"], " mouse over, ", [":active", "c"], " being clicked, ", [":visited", "c"], " link already seen."],
    [[":focus-visible", "c"], " shows where keyboard users are. ", ["Never remove the focus outline", "b"], " without replacing it."],
  ],
  notes: "Open 01_pseudo_classes_example.html. Hover the links and buttons, click and hold a button (:active), then click in the page and press Tab. The orange outline follows the keyboard.\n\nAccessibility: people who can't use a mouse move around with Tab. If the focus outline is removed (outline: none, sadly common), they have no idea where they are on the page. That fails WCAG 2.4.7 Focus Visible. :focus-visible only shows the outline for keyboard users, so mouse users don't see it on every click. Best of both.\n\nLink order matters (the cascade again): :link, :visited, :hover, :active. Memory trick: LoVe HAte. If :hover comes before :visited, visited links never show their hover colour.\n\n:visited is limited for privacy: browsers only let you change colours, so sites can't detect your browsing history.\n\nDevTools: in Styles, the ':hov' button lets you force a state (hover, focus, active) so you can inspect it without holding the mouse still.",
});

// 8. Structural (Example 02)
codeSlide({
  kicker: "SELECTORS 2 · POSITION", title: "Structural Pseudo-classes", file: "02_structural_example.html", lines: styleOf("02_structural_example.html"), result: SH.ex02, size: 13.5,
  bullets: [
    [[":first-child", "c"], " / ", [":last-child", "c"], ": the first or last inside its parent."],
    [[":nth-child(even)", "c"], " zebra stripes. Also ", ["odd", "c"], ", ", ["3", "c"], " (the third) or ", ["3n", "c"], " (every third)."],
    [[":not(.sold-out)", "c"], ": everything ", ["except", "b"], " what's in the brackets."],
  ],
  notes: "None of these need a class in the HTML. The browser works out the position for you, so if you add a row to the table, the stripes still alternate correctly. With classes, you'd have to renumber every row by hand.\n\nnth-child counts from 1, not 0 (unlike JavaScript arrays, which students will meet later).\n\n'tr:first-child' targets the header row here because it's the first tr in the table. In a real table you'd use <thead> and <th> (Class 4's accessible table), but this shows the idea.\n\nTry in class: change even to odd, then to 3n, and save. Ask: what would li:nth-child(2) pick?\n\n:not() is handy with lists: 'li:not(:last-child)' adds a border between items but not after the last one.",
});

// 9. Attribute (Example 03)
codeSlide({
  kicker: "SELECTORS 3 · ATTRIBUTE", title: "Attribute Selectors", file: "03_attribute_example.html", lines: styleOf("03_attribute_example.html"), result: SH.ex03, size: 13.5,
  bullets: [
    ["Square brackets match ", ["what's inside the tag", "b"], ": ", ["[type=\"email\"]", "c"], " is an exact match."],
    [["^=", "c"], " starts with, ", ["$=", "c"], " ends with, ", ["*=", "c"], " contains. ", ["[target]", "c"], " just has the attribute."],
    ["Great for links: show which ones go to another site, or download a PDF."],
  ],
  notes: "These read the HTML attributes we already write, so again, no extra classes.\n\nMemory trick: ^ is the start of a line and $ is the end in 'regular expressions', which students will see again in programming courses. * means 'anywhere'.\n\nReal uses: a little icon after every external link or PDF link (with ::after, coming up), styling form fields by type, and [aria-current=\"page\"] to highlight the current page in the nav. Our Harbourside nav already has aria-current=\"page\" on the Menu link from Class 4!\n\nWhy tell people a link is a PDF? It's an accessibility and usability point: a PDF can be big, and it opens differently from a web page.",
});

// 10. Combinators (Example 04)
codeSlide({
  kicker: "SELECTORS 4 · RELATIONSHIP", title: "Combinators: space, >, +, ~", file: "04_combinators_example.html", lines: styleOf("04_combinators_example.html"), result: SH.ex04, size: 13.5,
  bullets: [
    [["article p", "c"], " (space): any p inside, ", ["any depth", "b"], ". ", ["article > p", "c"], ": ", ["direct children", "b"], " only."],
    [["h2 + p", "c"], ": the p ", ["right after", "b"], " an h2. ", ["h3 ~ p", "c"], ": ", ["every", "b"], " p after an h3 (same parent)."],
  ],
  notes: "Family-tree language: the article is the PARENT of the h2 and the two p's directly inside it. They're CHILDREN of the article and SIBLINGS of each other. The p inside the div is a grandchild (a descendant, but not a child).\n\nWalk through the result:\n- All three p's in the article are blue (descendant: any depth).\n- Only the first two have the left border. The nested one is inside a div, so it isn't a direct child.\n- 'Fresh lobster…' is big and bold: it comes straight after the h2.\n- Both p's after the h3 are orange: ~ picks every later sibling, not just the next one.\n\nVS Code tip: hover over any selector in a .css file and it shows a tiny HTML example of what it matches, and its specificity.",
});

// 11. Pseudo-elements (Example 05)
codeSlide({
  kicker: "SELECTORS 5 · PART", title: "Pseudo-elements: ::before, ::after", file: "05_pseudo_elements_example.html", lines: styleOf("05_pseudo_elements_example.html"), result: SH.ex05, size: 13.5,
  bullets: [
    [["Two colons", "b"], " = part of an element: ", ["::first-letter", "c"], ", ", ["::first-line", "c"], ", ", ["::selection", "c"], "."],
    [["::before", "c"], " and ", ["::after", "c"], " add content with CSS. They ", ["need", "b"], " a ", ["content", "c"], " property, even if it's empty: ", ["content: \"\"", "c"], "."],
  ],
  notes: "Open the example and highlight some text with the mouse to see ::selection.\n\n::before and ::after create a little 'fake' element inside the real one, at the start or end. Without the content property, nothing appears at all. That's the number one bug with these.\n\nWhat to use them for: decoration. Badges, icons, quote marks, arrows on links. Don't put important information in them: some screen readers read the content, some don't, and it can't be selected or copied.\n\nThey're also how a lot of fancy CSS shapes and effects are made, with content: \"\" and a background colour.\n\nOne colon vs two: : is a state (pseudo-class), :: is a part (pseudo-element). Older code sometimes uses one colon for :before. Browsers still accept it, but two is correct.",
});

// 12-13. Game: Name That Selector
const PUZZLES = [
  ["Every odd row of a table", "tr:nth-child(odd)"],
  ["A button while the mouse is over it", "button:hover"],
  ["Every link to a Word file (.docx)", "a[href$=\".docx\"]"],
  ["Only the li's directly inside .menu", ".menu > li"],
  ["The first p right after each h2", "h2 + p"],
  ["A ★ before every .favourite", ".favourite::before"],
];
function puzzleSlide(title, reveal, notes) {
  const s = activitySlide(title, notes);
  const cw = 3.82, ch = 2.45, gap = 0.27;
  const cols = [C.orange, C.blue, C.purple, C.green, C.teal, C.red];
  PUZZLES.forEach(([q, a], i) => {
    const x = 0.75 + (i % 3) * (cw + gap), y = 1.5 + Math.floor(i / 3) * (ch + gap);
    card(s, { x, y, w: cw, h: ch });
    numCircle(s, i + 1, x + 0.25, y + 0.25, cols[i], 0.48);
    txt(s, q, { x: x + 0.9, y: y + 0.2, w: cw - 1.1, h: 0.9, fontSize: 16, bold: true, color: C.black, valign: "middle" });
    const bx = { x: x + 0.25, y: y + 1.3, w: cw - 0.5, h: 0.85 };
    if (reveal) codeBox(s, [a], Object.assign({ maxSize: 16 }, bx));
    else {
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, Object.assign({ rectRadius: 0.08, fill: { color: "F4F8FA" }, line: { color: "B0B8BC", width: 1, dashType: "dash" } }, bx));
      txt(s, "?", Object.assign({ fontSize: 28, bold: true, color: "B0B8BC", align: "center", valign: "middle" }, bx));
    }
  });
  return s;
}
puzzleSlide("Game: Name That Selector! (10 min)", false,
  "Make it a game: teams of 2–3, a minute or two per card (or all six at once with a 6-minute timer). Teams write their answers on paper or in a Teams chat, then reveal on the next slide. One point per correct selector, a bonus point for #6 if they also write the content line.\n\nAnswers (also on the next slide):\n1. tr:nth-child(odd)\n2. button:hover\n3. a[href$=\".docx\"]\n4. .menu > li\n5. h2 + p\n6. .favourite::before { content: \"★ \"; }\n\nTiebreaker: 'Every button except .cancel' → button:not(.cancel)\n\nAccept reasonable alternatives: e.g. #1 could be tr:nth-child(2n+1), and #2 could include :focus-visible too (a nice touch).");
puzzleSlide("Name That Selector: Answers", true,
  "Go through each one and ask a team that got it right to explain it.\n\n#6 needs the content property to actually show the star:\n.favourite::before { content: \"★ \"; }\n\nTiebreaker answer: button:not(.cancel).\n\nCelebrate the winning team, then move on to the fun part: animation.");

// 14. Divider
dividerSlide("Let's Make Things Move!", "transform   ·   transitions   ·   @keyframes   ·   a ball, a race and a car",
  "Animation is the fun part of CSS, and it's used everywhere: buttons that react when you hover, loading spinners, things that fade in as the page loads.\n\nThe plan: first the transform property (how to move, spin and resize things), then transitions (a smooth change on hover), then @keyframes (animations that run on their own). Lots of demos, then the Animation Playground.");

// 15. Transform toolkit
{
  const s = panelSlide("Your Moving Toolkit: transform",
    "Almost every animation uses transform. It moves, spins or resizes an element without changing the page layout: the other elements don't get pushed around, the element is just drawn somewhere else.\n\ntranslate(x, y) moves. Positive x goes right, positive y goes DOWN (screen coordinates start at the top left, unlike in maths class).\nrotate() spins, in degrees: 360deg is one full turn. Positive is clockwise.\nscale() resizes: 1 is normal, 2 is double, 0.5 is half. scale(1.3, 0.7) stretches wide and squashes flat. That's how cartoon bounces are made.\n\nCombine them in one transform, separated by spaces. Two transform declarations on the same element don't add up: the second replaces the first (the cascade again).\n\nWhy transform and not margin or width? The browser can animate transform and opacity very smoothly (on the graphics card). Animating margin, width or top makes the browser redo the whole page layout every frame, which can stutter.");
  const tools = [
    ["translate(x, y)", "Move it", "translateX(60px)", { dx: 0.62 }, C.orange],
    ["rotate(deg)", "Spin it", "rotate(45deg)", { rot: 45 }, C.blue],
    ["scale(n)", "Resize it", "scale(1.5)", { sc: 1.45 }, C.purple],
    ["Combine them", "spaces between", "translateX(40px) rotate(20deg) scale(1.2)", { dx: 0.45, rot: 20, sc: 1.2 }, C.green],
  ];
  const cw = 3.65, ch = 2.4;
  tools.forEach(([fn, t, ex, tr, col], i) => {
    if (i === 3) [fn, t] = [t, fn];
    const x = CX + (i % 2) * (cw + 0.25), y = 1.95 + Math.floor(i / 2) * (ch + 0.22);
    card(s, { x, y, w: cw, h: ch });
    txt(s, t, { x: x + 0.2, y: y + 0.12, w: cw - 0.4, h: 0.38, fontSize: 16, bold: true, color: C.black });
    txt(s, fn, { x: x + 0.2, y: y + 0.48, w: cw - 0.4, h: 0.3, fontSize: 12.5, bold: true, fontFace: MONO, color: "8A2B0B" });
    // the "before" (dashed ghost) and "after" squares
    const d = 0.55, gx = x + 0.55, gy = y + 1.12;
    s.addShape(pres.shapes.RECTANGLE, { x: gx, y: gy, w: d, h: d, fill: { type: "none" }, line: { color: "9AA0A6", width: 1, dashType: "dash" } });
    const sc = tr.sc || 1, nd = d * sc;
    s.addShape(pres.shapes.RECTANGLE, { x: gx + (tr.dx || 0) - (nd - d) / 2, y: gy - (nd - d) / 2, w: nd, h: nd, rotate: tr.rot || 0,
      fill: { color: col, transparency: 15 }, line: { color: col } });
    if (tr.dx) txt(s, "→", { x: gx + d - 0.05, y: gy - 0.38, w: 0.5, h: 0.35, fontSize: 16, bold: true, color: col, align: "center" });
    txt(s, ex, { x: x + 0.2, y: y + ch - 0.62, w: cw - 0.4, h: 0.5, fontSize: 10.5, fontFace: MONO, color: C.grey, valign: "bottom" });
  });
}

// 16. Transitions (Example 06)
codeSlide({
  kicker: "ANIMATION 1 · TRANSITIONS", title: "Transitions: Smooth Changes", file: "06_transition_example.html", lines: styleOf("06_transition_example.html"), size: 13.5,
  frames: [{ img: SH.ex06a, label: "NORMAL" }, { img: SH.ex06b, label: ":HOVER" }],
  bullets: [
    ["Without ", ["transition", "c"], ", a ", [":hover", "c"], " change snaps instantly. With it, the browser ", ["fills in the in-between", "b"], "."],
    ["Put ", ["transition", "c"], " on the ", ["normal", "b"], " rule, so it animates both ways: on and off."],
    ["A transition needs a ", ["trigger", "b"], " (like hover) and has only two states: start and end."],
  ],
  notes: "Open 06_transition_example.html and hover the card. Then comment out the transition line, save, and hover again: it snaps. That's the whole difference.\n\ntransition: [property] [duration] [timing];\n'transition: all 0.3s' works too, but naming the properties is clearer and avoids surprises.\n\nWhere to put it: on the normal rule (.card), not on .card:hover. If it's only on the hover rule, it animates when you mouse over, but snaps back when you leave.\n\nGood durations for UI: 0.15s to 0.4s. Longer feels sluggish.\n\nThis ties back to Selectors 1: transitions are what make :hover and :focus feel polished. Next: what if we want something to move all by itself, or with more than two steps? That's @keyframes.",
});

// 17. @keyframes anatomy
{
  const s = panelSlide("@keyframes: Describe It, Then Play It",
    "Two steps, every time.\n1. Describe the animation with @keyframes and give it a name. from = 0% (the start), to = 100% (the end). Inside each step, ordinary CSS declarations.\n2. Play it: put the animation property on the element, using the same name and a duration.\n\nRead it out loud: 'Slide the ball from 0 to 300 pixels, over 2 seconds, easing in and out, forever.'\n\nTransition vs animation:\n- A transition needs a trigger (hover, focus) and goes from A to B.\n- An animation runs by itself as soon as the page loads, can have as many steps as you like (0%, 25%, 50%…), and can loop forever.\n\nThe name is yours to choose: lower case and hyphens, like a class name. The most common bug is a typo, so the name in animation doesn't match the name in @keyframes. CSS fails silently: nothing moves.");
  codeBox(s, ["@keyframes slide {", "   from { transform: translateX(0); }", "   to   { transform: translateX(300px); }", "}", "", ".ball {", "   animation: slide 2s ease-in-out infinite;", "}"],
    { x: CX, y: 1.95, w: 4.3, h: 2.75, maxSize: 15, lang: "css" });
  const parts = [
    ["@keyframes slide", HL.at, "1. Describe it", "and give it a name"],
    ["from / to", HL.sel, "The start and end", "(or 0% … 100%)"],
    ["transform", HL.prop, "What changes", "any CSS that has numbers or colours"],
    ["animation: slide", HL.prop, "2. Play it", "on an element, using the same name"],
    ["2s", HL.num, "How long", "one run takes (or 500ms)"],
    ["infinite", HL.val, "How many times", "1 (default), 3, or forever"],
  ];
  parts.forEach(([ex, col, t, d], i) => {
    const y = 1.95 + i * 0.83, x = CX + 4.55, w = CW - 4.55;
    card(s, { x, y, w, h: 0.72, fill: "1E1E1E", border: "1E1E1E" });
    txt(s, ex, { x: x + 0.15, y: y + 0.04, w: w - 0.3, h: 0.3, fontSize: 11.5, bold: true, fontFace: MONO, color: col });
    txt(s, [{ text: t + " ", options: { bold: true, color: C.white } }, { text: d, options: { color: "BFBFBF" } }],
      { x: x + 0.15, y: y + 0.34, w: w - 0.3, h: 0.32, fontSize: 11.5 });
  });
  card(s, { x: CX, y: 4.95, w: 4.3, h: 1.95, fill: C.white });
  txt(s, "Read it out loud", { x: CX + 0.25, y: 5.08, w: 3.8, h: 0.4, fontSize: 15, bold: true, color: C.black });
  txt(s, "\"Slide the ball from 0 to 300px, over 2 seconds, easing in and out, forever.\"", { x: CX + 0.25, y: 5.5, w: 3.8, h: 1.3, fontSize: 15, italic: true, color: C.dark, valign: "top" });
}

// 18. Ball around the browser (Example 07)
codeSlide({
  kicker: "ANIMATION 2 · EXAMPLE 07", title: "A Ball Around the Browser", file: "07_ball_around_example.html", lines: styleOf("07_ball_around_example.html"), size: 13.5,
  frames: SH.ex07.map((img, i) => ({ img, label: `${i} s · ${["0%", "25%", "50%", "75%"][i]}` })),
  bullets: [
    [["border-radius: 50%", "c"], " turns the square into a ball."],
    ["Percentages are ", ["checkpoints", "b"], ": 25% of 4s = 1s. The browser animates between them."],
    [["100vw", "c"], " / ", ["100vh", "c"], " = the browser's full width / height. Minus 80px, so the ball stays on screen."],
  ],
  notes: "Open 07_ball_around_example.html and resize the browser while it runs: the ball still finds the corners, because the keyframes use vw and vh.\n\nThe pictures are the same page, paused at 0, 1, 2 and 3 seconds: the four corners. The colour changes too: you can animate colours as well as positions.\n\nWhy 'minus 80px'? translate moves the ball's top-left corner. If it moved the full 100vw, the ball would end up just past the right edge, out of sight.\n\nTry in class (pick one or two):\n- Change 4s to 1s. Whoa.\n- Change linear to ease-in-out: now it pauses at every corner.\n- Delete the 50% step: it cuts diagonally across the screen.\n- Add a rotate: transform: translate(…) rotate(360deg) at 100%.\n\nDevTools: More tools → Animations shows a timeline you can slow down to 10% and scrub through.",
});

// 19. Animation properties
{
  const s = panelSlide("The animation Properties",
    "The shorthand on the left is what we'll use most of the time. Every part has a long-hand property too, which is handy when you only want to change one thing (like Example 08, where only the timing function differs).\n\nThe order in the shorthand is flexible, with one rule: the first time value is the duration, and the second is the delay.\n\niteration-count: a number or infinite.\ndirection: alternate plays forwards then backwards (ping-pong). reverse plays backwards.\nfill-mode: forwards keeps the last keyframe when the animation finishes, instead of jumping back to the start.\ndelay: a negative delay starts the animation part-way through, handy for making several copies out of step with each other.");
  codeBox(s, ["animation: around 4s ease-in-out 1s infinite alternate;"], { x: CX, y: 1.95, w: CW, h: 0.6, maxSize: 15, lang: "css", });
  txt(s, "name · duration · timing-function · delay · iteration-count · direction", { x: CX, y: 2.65, w: CW, h: 0.35, fontSize: 12, color: C.grey, fontFace: MONO, align: "center" });
  table(s, [
    ["Property", "Values", "What it does"],
    ["animation-name", "around", "Which @keyframes to play"],
    ["animation-duration", "4s   500ms", "How long one run takes"],
    ["animation-timing-function", "linear   ease-in-out", "The speed curve (next slide!)"],
    ["animation-delay", "1s", "Wait before starting"],
    ["animation-iteration-count", "3   infinite", "How many times"],
    ["animation-direction", "alternate   reverse", "Ping-pong, or backwards"],
    ["animation-fill-mode", "forwards", "Stay on the last frame at the end"],
  ], { x: CX, y: 3.3, w: CW, colW: [2.75, 2.15, 2.65], size: 12, rowH: 0.43, codeCols: [0, 1] });
}

// 20. Timing race (Example 08)
codeSlide({
  kicker: "ANIMATION 3 · EXAMPLE 08", title: "The Timing Function Race", file: "08_timing_race_example.html", lines: styleOf("08_timing_race_example.html"), size: 13,
  frames: [{ img: SH.ex08, label: "PAUSED AT 1 SECOND" }], cols: 1, maxFrameH: 3.45,
  bullets: [
    ["Same distance, same 3 seconds. ", ["They all tie!", "b"], " Only the speed along the way is different."],
    [["ease-in", "c"], " starts slow, ", ["ease-out", "c"], " ends slow, ", ["linear", "c"], " never changes. ", ["steps()", "c"], " jumps."],
  ],
  notes: "Before opening it: 'Six balls, same race, same 3 seconds. Which one wins?' Take bets. Then open 08_timing_race_example.html. Spoiler: it's always a tie. They all start and finish together.\n\nThe picture is paused at 1 second: ease-out (green) is way ahead because it starts fast; ease-in (purple) is still near the start.\n\nWhen to use which:\n- ease-out: things arriving on screen (a menu sliding in). Feels responsive.\n- ease-in: things leaving, and things falling (gravity speeds them up).\n- ease-in-out: things going back and forth.\n- linear: spinning (so each turn joins smoothly onto the next) and constant motion like a car or conveyor belt.\n- steps(): ticking clocks, sprite animations, typewriter effects.\n\nease is the default if you don't say anything.\n\nBonus: cubic-bezier.com lets you design your own curve, including ones that overshoot and bounce back.",
});

// 21. Bounce (Example 09)
codeSlide({
  kicker: "ANIMATION 4 · EXAMPLE 09", title: "A Bouncing Ball (with Squash!)", file: "09_bounce_example.html", lines: styleOf("09_bounce_example.html"), size: 13,
  frames: SH.ex09.map((img, i) => ({ img, label: ["0 s", "0.4 s", "0.6 s", "0.85 s"][i] })), cols: 4,
  bullets: [
    ["A timing function ", ["inside", "b"], " a keyframe applies until the ", ["next", "b"], " keyframe: ", ["ease-in", "c"], " falling, ", ["ease-out", "c"], " rising."],
    [["scale(1.3, 0.7)", "c"], " squashes it: wider and flatter. ", ["transform-origin", "c"], " keeps it on the floor."],
  ],
  notes: "Open 09_bounce_example.html. This is the most 'real' looking of the demos, and it's because of two tricks animators have used since the first cartoons:\n\n1. Timing. Real things speed up as they fall (ease-in) and slow down as they rise (ease-out). With linear, it looks robotic. Try it: change both to linear and watch.\n2. Squash and stretch. For a split second at the bottom, the ball squashes flat. It's only 10% of the animation and you barely notice it, but delete the 50% step and it suddenly feels like a hard plastic ball.\n\ntransform-origin: center bottom means 'squash towards the floor'. Without it, the ball squashes towards its middle and floats above the floor for a moment.\n\nTry: change 1.2s to 0.6s (a super ball) or 2.5s (on the moon).",
});

// 22. Shape-shifting square (Example 10)
codeSlide({
  kicker: "ANIMATION 5 · EXAMPLE 10", title: "A Shape-shifting Square", file: "10_morph_square_example.html", lines: styleOf("10_morph_square_example.html"), size: 14,
  frames: SH.ex10.map((img, i) => ({ img, label: ["0 s", "0.75 s", "1.5 s", "2.25 s"][i] })),
  bullets: [
    ["Only ", ["50%", "c"], " and ", ["100%", "c"], "? The browser uses the element's ", ["normal style", "b"], " for the missing steps."],
    ["Animate several things at once: ", ["transform", "c"], ", ", ["border-radius", "c"], " and ", ["background-color", "c"], "."],
  ],
  notes: "Open 10_morph_square_example.html.\n\nThe morph keyframes have no 0%, so the animation starts from the square's normal style: purple, square corners, no rotation. Halfway through it's a big blue circle. At 100% only transform is listed, so border-radius and colour go back to normal, and it has spun one full turn: rotate(360deg) looks exactly like the start, so the loop is seamless.\n\nTry: change scale(1.5) to scale(0.2), or rotate(180deg) to rotate(720deg) at 50%. Swap ease-in-out for linear and see how the 'breathing' feel disappears.",
});

// 23. Car (Example 11)
codeSlide({
  kicker: "ANIMATION 6 · EXAMPLE 11", title: "A Car Made of Boxes", file: "11_car_example.html", lines: styleOf("11_car_example.html"), size: 13.5,
  frames: SH.ex11.map((img, i) => ({ img, label: ["0.9 s", "1.6 s", "2.4 s"][i] })), cols: 1,
  bullets: [
    ["Four boxes make a car. The car ", ["drives", "b"], " and each wheel ", ["spins", "b"], ": two animations at once."],
    ["The ", ["dashed border", "b"], " is what makes the spin visible. A plain circle spinning looks like it's standing still!"],
  ],
  notes: "Open 11_car_example.html.\n\nThe .car element drives across the screen, from off the left edge (-200px) to off the right edge (100vw). Each wheel has its own spin animation, and because the wheels are inside the car, they turn AND travel with it.\n\nposition: relative / absolute is a sneak peek: it lets us put the roof, body and wheels exactly where we want them inside the car (left, bottom). We'll cover position properly with layout.\n\nTry: make the drive 2s (faster) without changing the spin. It looks like the wheels are skidding! Real animators match the two speeds.\n\nChallenge for fast finishers later: add a sun, a cloud that drifts slower than the car, or headlights with ::before.",
});

// 24. Real world + accessibility (Example 12)
codeSlide({
  kicker: "ANIMATION 7 · EXAMPLE 12", title: "Real Websites, and Being Kind", file: "12_real_world_example.html", lines: styleOf("12_real_world_example.html").slice(4), size: 13,
  frames: [{ img: SH.ex12, label: "PAUSED AT 0.5 SECONDS" }], cols: 1,
  bullets: [
    ["Real sites use small animations: ", ["fade in", "b"], ", ", ["loading spinners", "b"], ", a ", ["pulse", "b"], " to catch your eye."],
    [["prefers-reduced-motion", "c"], ": some people get dizzy or sick from motion. Turn it off for them."],
    ["Animate ", ["transform", "c"], " and ", ["opacity", "c"], " (smooth). Never flash more than 3 times a second."],
  ],
  notes: "Open 12_real_world_example.html (reload to see the heading fade in).\n\nprefers-reduced-motion: people with vestibular (inner-ear) disorders, migraines or some kinds of epilepsy can feel sick from big moving things. They switch animations off in their operating system: Windows Settings → Accessibility → Visual effects → Animation effects. The @media rule asks 'has this person turned motion off?' and if so, switches our animations off. It's WCAG 2.3.3.\n\nDemo it: turn Animation effects off in Windows, then reload the example: nothing moves. Turn it back on.\n\nIMPORTANT for the playground: if a student's animations don't work AT ALL, check that setting first. Some lab images have it turned off.\n\nFlashing: WCAG 2.3.1 says nothing should flash more than 3 times per second. It can trigger seizures. So no strobing backgrounds, however fun it seems.\n\nGood taste: animation should help the user (show something is loading, draw the eye to a sale) and not get in the way. If everything on a page moves, nothing stands out.",
});

// 25. Playground intro
{
  const s = activitySlide("The Animation Playground (40 min)",
    "This is the main event, and it's meant to be fun and very doable. Four stages, each with an object that doesn't move yet. Each challenge gives students the full code to type, then asks them to change one thing and watch what happens.\n\nEvery challenge uses the same recipe: @keyframes with from and to, then an animation line. By the fourth one, it should feel easy. That's the point.\n\nFast finishers: the Bonus (Around the Block) uses percentages like Example 07, then 'Make it yours'.\n\nWalk around and celebrate the first person to get each challenge working. The moment something first moves, people cheer.");
  const ih = 5.5, img = SH.playground, iw = ih / img.ratio;
  framedImage(s, img, 0.75, 1.45, iw);
  const x = 0.75 + iw + 0.45, w = W - x - 0.6;
  card(s, { x, y: 1.45, w, h: 5.5 });
  txt(s, "Four stages. Make them all move!", { x: x + 0.3, y: 1.6, w: w - 0.6, h: 0.45, fontSize: 18, bold: true, color: C.black });
  const list = [["Roll", "a ball rolls across"], ["Ping-pong", "back and forth, forever"], ["Pinwheel", "spin, spin, spin"], ["Pulse", "grow and change colour"], ["Bonus: Around the Block", "if you finish early"]];
  const colrs = [C.orange, C.blue, C.teal, C.purple, C.green];
  list.forEach(([t, d], i) => {
    const y = 2.25 + i * 0.52;
    numCircle(s, i < 4 ? i + 1 : "★", x + 0.3, y, colrs[i], 0.4);
    txt(s, t, { x: x + 0.85, y: y - 0.05, w: 3.4, h: 0.5, fontSize: 15.5, bold: true, color: C.black, valign: "middle" });
    txt(s, d, { x: x + 4.0, y: y - 0.05, w: w - 4.2, h: 0.5, fontSize: 13, color: C.grey, valign: "middle" });
  });
  // the recipe
  card(s, { x: x + 0.3, y: 5.0, w: w - 0.6, h: 1.75, fill: C.dark, border: C.dark });
  txt(s, "Every challenge, the same two steps:", { x: x + 0.55, y: 5.1, w: w - 1.1, h: 0.4, fontSize: 14, bold: true, color: C.white });
  richBullets(s, [
    [["1. @keyframes", "b"], " describes the move: from … to …"],
    [["2. animation", "b"], " plays it on the object."],
  ], { x: x + 0.55, y: 5.55, w: w - 1.1, h: 1.1, size: 14, bullet: false, color: "D9D9D9", space: 6 });
}

// 26. Try It 1: Roll
trySlide({
  n: 1, title: "Roll", time: "8 min",
  steps: [
    ["Open activity/index.html with Live Server", "Four stages. Nothing moves… yet!"],
    ["Open activity/style.css", ["Find the ", ["Challenge 1", "b"], " heading."]],
    ["Type the code on the right", "Save. The ball rolls!"],
    ["Play with it", ["Change ", ["2s", "c"], " to ", ["0.5s", "c"], ". Then try ", ["5s", "c"], ". Which do you like?"]],
  ],
  code: ["@keyframes roll {", "   from { transform: translateX(0); }", "   to   { transform: translateX(540px); }", "}", "", "#roll {", "   animation: roll 2s infinite;", "}"],
  tip: [["Nothing moving? Check that ", ["roll", "c"], " is spelled the same in both places, and that ", ["2s", "c"], " has its ", ["s", "c"], "."]],
  notes: "Do this one together, typing it live. Make sure everyone has a rolling ball before moving on.\n\nRead it out loud: 'Roll from 0 to 540 pixels, over 2 seconds, forever.' Why 540? The stage is 600px wide and the ball is 60px wide, so 540 puts it right at the edge.\n\nThe usual problems:\n- The name is spelled differently in the two places (roll vs Roll).\n- No unit on the duration: '2' does nothing; it must be 2s.\n- Nothing moves at all, anywhere, even in the examples: check Windows Settings → Accessibility → Visual effects → Animation effects is On.",
});

// 27. Try It 2: Ping-pong
trySlide({
  n: 2, title: "Ping-pong", time: "5 min",
  steps: [
    ["No new @keyframes needed!", ["Ping-pong uses the same ", ["roll", "c"], " you wrote in Challenge 1."]],
    ["Type the code on the right", ["Just one rule, under the ", ["Challenge 2", "b"], " heading."]],
    ["Watch the difference", [["alternate", "c"], " rolls there… then back. No more jumping!"]],
    ["Play with it", ["Take out ", ["alternate", "c"], " and save. Put it back. Try ", ["0.3s", "c"], " for a fast rally!"]],
  ],
  code: ["#ping-pong {", "   animation: roll 1s infinite alternate;", "}"],
  tip: [["One @keyframes can be used by ", ["as many elements as you like", "b"], "."]],
  notes: "The shortest challenge, and a nice confidence boost.\n\nCompare the two balls: Roll jumps back to the start at the end of each run. Ping-pong rolls back smoothly, because alternate plays the animation forwards, then backwards.\n\nIf anyone asks why it slows down at each end: that's the default timing function, ease. Try linear for a robotic ping-pong.",
});

// 28. Try It 3: Pinwheel
trySlide({
  n: 3, title: "Pinwheel", time: "8 min",
  steps: [
    ["New move: rotate", [["rotate(360deg)", "c"], " is one full turn."]],
    ["Type the code on the right", ["Under the ", ["Challenge 3", "b"], " heading. Save. It spins!"]],
    ["Why linear?", ["It keeps the same speed all the way round. Take it out and see it slow down every turn."]],
    ["Play with it", ["Try ", ["720deg", "c"], ", or a faster ", ["0.5s", "c"], ". Whee!"]],
  ],
  code: ["@keyframes spin {", "   from { transform: rotate(0deg); }", "   to   { transform: rotate(360deg); }", "}", "", "#pinwheel {", "   animation: spin 2s linear infinite;", "}"],
  tip: [["The pinwheel is just a box with ", ["four thick coloured borders", "b"], " and no size. Look in stage.css!"]],
  notes: "Same recipe, a new kind of move. Instead of translateX (move), it's rotate (spin).\n\nThe pinwheel trick: a box with width 0 and height 0, and a thick border in four different colours. The borders meet at the corners in diagonal lines, which makes four triangles. Show stage.css if anyone's curious.\n\nlinear vs ease: with ease (the default), the pinwheel slows down at the end of every turn and speeds up again, so it looks like it's stuttering. linear keeps it smooth.",
});

// 29. Try It 4: Pulse
trySlide({
  n: 4, title: "Pulse", time: "8 min",
  steps: [
    ["New move: scale", [["scale(1.5)", "c"], " makes it one and a half times bigger."]],
    ["Change two things at once", ["Size ", ["and", "b"], " colour, in the same from and to."]],
    ["Type the code on the right", ["Under the ", ["Challenge 4", "b"], " heading. Save. It breathes!"]],
    ["Play with it", ["Pick your own colours. Add ", ["border-radius: 50%;", "c"], " to the to step. What happens?"]],
  ],
  code: ["@keyframes pulse {", "   from {", "      transform: scale(1);", "      background-color: #A02B93;", "   }", "   to {", "      transform: scale(1.5);", "      background-color: #E97132;", "   }", "}", "", "#pulse {", "   animation: pulse 1s infinite alternate;", "}"],
  tip: [["You can animate ", ["colours", "b"], " too, not just movement."]],
  notes: "The last core challenge. By now the recipe should feel familiar: keyframes with from and to, then an animation line.\n\nNew ideas: scale (resize), and changing two properties in the same keyframe. alternate again, so it grows and shrinks smoothly.\n\nThe border-radius suggestion turns the square into a circle at the big end. Students love it.\n\nAt this point everyone has four things moving on their page. Celebrate that!",
});

// 30. Bonus + make it yours + submit
{
  const s = activitySlide("Bonus, Make It Yours, Then Submit",
    "For anyone who finishes early, or as homework.\n\nThe bonus is Example 07 at playground size: 600 − 60 = 540 across, 200 − 60 = 140 down. Percentages are checkpoints along the way.\n\nWhat to collect on Brightspace: their style.css, and a screenshot (or a short screen recording: Windows 11's Snipping Tool can record the screen). Anyone who has the four challenges moving has done the job.");
  card(s, { x: 0.75, y: 1.45, w: 5.6, h: 5.5 });
  numCircle(s, "★", 1.0, 1.7, C.green, 0.44);
  txt(s, "Bonus: Around the Block", { x: 1.6, y: 1.67, w: 4.6, h: 0.5, fontSize: 16, bold: true, color: C.black, valign: "middle" });
  txt(s, "Copy it in. Percentages are checkpoints: the four corners.", { x: 1.0, y: 2.25, w: 5.1, h: 0.6, fontSize: 13, color: C.dark, valign: "top" });
  codeBox(s, ["@keyframes around {", "   0%   { transform: translate(0, 0); }", "   25%  { transform: translate(540px, 0); }", "   50%  { transform: translate(540px, 140px); }", "   75%  { transform: translate(0, 140px); }", "   100% { transform: translate(0, 0); }", "}", "", "#around {", "   animation: around 4s linear infinite;", "}"],
    { x: 0.95, y: 2.8, w: 5.2, h: 3.95, maxSize: 14, lang: "css" });
  card(s, { x: 6.55, y: 1.45, w: 3.25, h: 5.5 });
  txt(s, "Make it yours", { x: 6.8, y: 1.65, w: 2.8, h: 0.45, fontSize: 18, bold: true, color: C.black });
  const ideas = [
    [["Change the ", ["colours", "b"]]],
    [["Make something ", ["faster", "b"], " or ", ["slower", "b"]]],
    [["Spin the pinwheel ", ["backwards", "b"], ": add ", ["reverse", "c"]]],
    [["Turn the pulse into a ", ["circle", "b"]]],
    [["Start the pulse ", ["tiny", "b"], ": ", ["scale(0.5)", "c"]]],
  ];
  ideas.forEach((segs, i) => {
    const y = 2.3 + i * 0.85;
    numCircle(s, i + 1, 6.8, y, C.orange, 0.38);
    richBullets(s, segs, { x: 7.3, y: y - 0.1, w: 2.4, h: 0.6, size: 14, bullet: false, color: C.dark, valign: "middle" });
  });
  card(s, { x: 10.0, y: 1.45, w: 2.6, h: 5.5, fill: C.dark, border: C.dark });
  txt(s, "Submit to Brightspace", { x: 10.2, y: 1.7, w: 2.2, h: 0.75, fontSize: 18, bold: true, color: C.white, valign: "top" });
  txt(s, [
    { text: "1. Your style.css file", options: { breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "2. A screenshot of your playground", options: { breakLine: true } },
    { text: " ", options: { fontSize: 10, breakLine: true } },
    { text: "Stuck? ", options: { bold: true, color: "F2C6AC" } },
    { text: "Peek at activity/answer/style.css. Then check the names match and every time has an s." },
  ], { x: 10.2, y: 2.55, w: 2.25, h: 4.2, fontSize: 13.5, color: "D9D9D9", valign: "top" });
}

// 31. Show and tell
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  s.addShape(pres.shapes.OVAL, { x: 7.6, y: 0.9, w: 5.4, h: 5.4, fill: { type: "none" }, line: { color: C.ring, width: 28 } });
  txt(s, "Show and Tell!", { x: 0.8, y: 0.8, w: 7, h: 0.9, fontSize: 40, color: C.black });
  richBullets(s, [
    ["Who wants to ", ["show off", "b"], " their playground?"],
    ["Which animation was your ", ["favourite", "b"], "?"],
    ["What was your funniest ", ["bug", "b"], "?"],
    ["Where have you seen animation on a ", ["real website", "b"], " this week?"],
  ], { x: 0.8, y: 2.0, w: 6.5, h: 4.5, size: 20, space: 16 });
  txt(s, "@keyframes = describe it\nanimation = play it", { x: 8.3, y: 2.8, w: 4.0, h: 1.6, fontSize: 21, bold: true, color: C.teal, align: "center", valign: "middle" });
  s.addNotes("Get two or three students to put their page on the projector. Applaud generously!\n\nCommon bugs: the name doesn't match between animation and @keyframes, a missing s on the duration, and @keyframes typed inside another rule.\n\nReal-world animation they've probably seen: loading spinners, the 'typing…' dots in chat apps, notification badges that pulse, buttons that change on hover, menus that slide in on phones.");
}

// 32. Selector cheat sheet
{
  const s = panelSlide("Cheat Sheet: Selectors",
    "Everything from the first half, in one place. It's in the README too.");
  table(s, [
    ["Selector", "Matches"],
    ["a:hover   button:active", "While the mouse is over it / while it's clicked"],
    [":focus-visible", "Focused with the keyboard (keep an outline!)"],
    [":first-child   :last-child", "The first / last inside its parent"],
    ["tr:nth-child(even)", "Every even row (also odd, 3, 3n)"],
    ["li:not(.sold-out)", "Every li except .sold-out"],
    ["input[type=\"email\"]", "An attribute with an exact value"],
    ["a[href^=\"https\"]   a[href$=\".pdf\"]", "Starts with / ends with"],
    ["article > p", "Direct children only"],
    ["h2 + p   h3 ~ p", "The very next sibling / every later sibling"],
    ["p::first-letter   ::selection", "Part of an element"],
    [".new::before   .price::after", "Content added by CSS (needs content: \"\")"],
  ], { x: CX, y: 1.9, w: CW, colW: [3.5, 4.05], size: 12.5, rowH: 0.41, codeCols: [0] });
}

// 33. Animation cheat sheet
{
  const s = panelSlide("Cheat Sheet: Animation",
    "Everything from the second half, in one place. Students can keep this slide open while they finish the playground.");
  table(s, [
    ["Code", "What it does"],
    ["transform: translate(x, y)", "Move (also translateX, translateY)"],
    ["transform: rotate(45deg)", "Spin (360deg = one full turn)"],
    ["transform: scale(1.5)", "Resize (scale(1.3, 0.7) = squash)"],
    ["transition: transform 0.3s ease;", "Smooth change between two states (e.g. :hover)"],
    ["@keyframes name { from {…} to {…} }", "Describe an animation (or 0%, 50%, 100%)"],
    ["animation: name 2s;", "Play it: name + duration"],
    ["linear   ease-in   ease-out   ease-in-out", "Timing function: the speed curve"],
    ["infinite   alternate", "Loop forever / go there and back"],
    ["forwards", "Stay on the last keyframe at the end"],
    ["@media (prefers-reduced-motion: reduce)", "Turn animations off for those who ask"],
  ], { x: CX, y: 1.9, w: CW, colW: [4.0, 3.55], size: 12.5, rowH: 0.44, codeCols: [0] });
}

// 34. Recap
{
  const s = panelSlide("Class Recap and Next Steps",
    "Next class: page layout. We've had sneak peeks of display and position today and last class; next time we look at them properly, then Flexbox for laying out navigation bars and menu cards side by side.");
  labelBullets(s, [
    { label: "Recap:", text: [
      "Pseudo-classes (:) match a state or position: :hover, :focus-visible, :nth-child(), :not().",
      "Attribute selectors [ ] and combinators (> + ~) pick elements without extra classes.",
      "Pseudo-elements (::) style part of an element, or add content with ::before / ::after.",
      "transform moves, spins and resizes. transition smooths a change between two states.",
      "@keyframes describes an animation; the animation property plays it.",
      "Animate with care: transform and opacity, no flashing, respect prefers-reduced-motion.",
    ] },
    { label: "Next class:", text: [
      "Page layout: display, position, and Flexbox.",
    ] },
  ], { size: 20, sub: 15.5 });
}

pres.writeFile({ fileName: OUT }).then((f) => console.log("wrote " + f));
