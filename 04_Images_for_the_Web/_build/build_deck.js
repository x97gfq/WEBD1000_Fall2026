// Builds Class7_Images_for_the_Web.pptx in the same style as the Class 5 and 6 decks.
// To rebuild: cd 04_Images_for_the_Web/_build, run "npm install" once, then "npm run build".
// make_samples.js runs first: it makes ../samples and sizes.json, so the sizes on the slides are real.
// Code slides read the example .html files directly, so edit those files and re-run to update the slides.
const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const REPO = path.join(__dirname, ".."); // the 04_Images_for_the_Web folder
const SCR = __dirname;
const OUT = path.join(REPO, "Class7_Images_for_the_Web.pptx");
const S = JSON.parse(fs.readFileSync(path.join(SCR, "sizes.json"), "utf8"));
const KB = (n) => (n >= 1000 ? (n / 1024).toFixed(1) + " MB" : (n < 10 ? n.toFixed(1) : Math.round(n)) + " KB");

const FONT = "Aptos";
const MONO = "Consolas";
const UNIT = "HTML/CSS Techniques";
const C = {
  dark: "262626", light: "F2F2F2", black: "000000", white: "FFFFFF",
  orange: "E97132", purple: "A02B93", blue: "0F9ED5", navy: "12304A",
  teal: "1B5E82", grey: "595959", mid: "7F7F7F", actBg: "E1E8EB",
  green: "2E7D32", red: "C62828", ring: "B7CFC6",
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "Class 7: Images for the Web";
pres.author = "Jamie Symonds";

const W = 13.333, H = 7.5;
const PANEL_W = 4.1;
const CX = 5.2, CW = 7.55; // content column on panel slides

// ---------- helpers (same as the Class 5 and 6 decks) ----------
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

// ---- HTML syntax highlighting (VS Code Dark+ colours) ----
const HL = { text: "D4D4D4", punct: "808080", tag: "569CD6", attr: "9CDCFE", val: "CE9178", comment: "6A9955", ent: "D7BA7D" };
function highlight(lines) {
  let inComment = false, inTag = false, inVal = null;
  const out = [];
  for (const line of lines) {
    const toks = [];
    let i = 0;
    const push = (t, c) => { if (!t) return; const last = toks[toks.length - 1]; if (last && last.c === c) last.t += t; else toks.push({ t, c }); };
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
    out.push(toks);
  }
  return out;
}

function bodyOf(file) {
  const s = fs.readFileSync(path.join(REPO, file), "utf8");
  let b = s.split("<body>")[1].split("</body>")[0].split("\n").map((l) => l.replace(/\r$/, ""));
  b = b.map((l) => (l.startsWith("\t\t") ? l.slice(2) : l));
  while (b.length && !b[0].trim()) b.shift();
  while (b.length && !b[b.length - 1].trim()) b.pop();
  return b.map((l) => l.replace(/\t/g, "   "));
}

function codeBox(s, lines, { x, y, w, h, maxSize = 13, fitHeight = false }) {
  const pad = Math.min(0.25, h * 0.18);
  const n = lines.length, maxLen = Math.max(...lines.map((l) => l.length));
  const byH = ((h - 2 * pad) * 72) / (1.18 * n);
  const byW = ((w - 2 * pad) * 72) / (0.6 * maxLen);
  const size = Math.max(9, Math.min(maxSize, byH, byW));
  if (fitHeight) {
    const nh = Math.min(h, (n * size * 1.2) / 72 + 2 * pad + 0.1);
    y = y + (h - nh) / 2; h = nh;
  }
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: "1E1E1E" }, line: { color: "1E1E1E" },
    shadow: { type: "outer", blur: 8, offset: 3, angle: 90, color: "000000", opacity: 0.35 } });
  const runs = [];
  const hl = highlight(lines);
  hl.forEach((toks, li) => {
    if (!toks.length) toks = [{ t: " ", c: HL.text }];
    toks.forEach((tk, ti) => {
      const o = { color: tk.c, fontFace: MONO, fontSize: Math.round(size * 2) / 2 };
      if (ti === toks.length - 1 && li < hl.length - 1) o.breakLine = true;
      runs.push({ text: tk.t, options: o });
    });
  });
  txt(s, runs, { x: x + pad, y: y + pad, w: w - 2 * pad, h: h - 2 * pad, valign: "top", lineSpacingMultiple: 1.0 });
}

function pill(s, text, x, y, color, w) {
  w = w || 0.3 + text.length * 0.1;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.34, rectRadius: 0.17, fill: { color }, line: { color } });
  txt(s, text, { x, y, w, h: 0.34, fontSize: 10.5, bold: true, color: C.white, align: "center", valign: "middle", charSpacing: 1 });
  return w;
}

// Code slide: gradient left + code, title & description right
const CODE_PANEL = 7.0;
function codeSlide({ kicker, title, tags = [], bullets, file, lines, notes, size = 14.5 }) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  s.addImage({ path: path.join(SCR, "grad_left.png"), x: 0, y: 0, w: CODE_PANEL, h: H });
  codeBox(s, lines, { x: 0.3, y: 0.4, w: CODE_PANEL - 0.6, h: H - 0.8, maxSize: 15, fitHeight: true });
  const RX = CODE_PANEL + 0.45, RW = W - RX - 0.45;
  txt(s, kicker, { x: RX, y: 0.45, w: RW, h: 0.35, fontSize: 14, color: C.grey, bold: true, charSpacing: 1 });
  const long = title.length > 26;
  txt(s, title, { x: RX, y: 0.8, w: RW, h: long ? 1.05 : 0.7, fontSize: long ? 26 : 32, color: C.black, valign: "top" });
  const ty = long ? 1.98 : 1.65;
  let px = RX;
  tags.forEach(([t, col]) => { px += pill(s, t, px, ty, col) + 0.12; });
  const by = tags.length ? ty + 0.55 : ty;
  richBullets(s, bullets, { x: RX, y: by, w: RW, h: 6.85 - by, size });
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
    fontFace: codeCols.includes(ci) && ri > 0 ? MONO : FONT, fontSize: size, color: ri === 0 ? C.white : C.black, bold: ri === 0 || (ci === 0 && ri > 0), valign: "middle",
    fill: { color: ri === 0 ? C.dark : (ri % 2 ? C.white : "E4E4E4") }, margin: [3, 6, 3, 6] } })));
  s.addTable(data, { x, y, w, colW, rowH, border: { type: "solid", pt: 0.5, color: "C8C8C8" } });
}

// Lab slide: numbered steps in a white card on the left, a dark "record / look for" card on the right
function labSlide({ title, steps, sideTitle, side, notes, stepGap = 0.98 }) {
  const s = activitySlide(title, notes);
  card(s, { x: 0.75, y: 1.45, w: 7.6, h: 5.5 });
  steps.forEach(([t, d], i) => {
    const y = 1.75 + i * stepGap;
    numCircle(s, i + 1, 1.05, y, C.orange, 0.46);
    txt(s, t, { x: 1.7, y: y - 0.02, w: 6.4, h: 0.35, fontSize: 16, bold: true, color: C.black });
    if (d) richBullets(s, [typeof d === "string" ? [d] : d], { x: 1.7, y: y + 0.33, w: 6.4, h: stepGap - 0.38, size: 13.5, bullet: false, color: C.dark });
  });
  card(s, { x: 8.75, y: 1.45, w: 3.85, h: 5.5, fill: C.dark, border: C.dark });
  txt(s, sideTitle, { x: 9.0, y: 1.7, w: 3.4, h: 0.45, fontSize: 18, bold: true, color: C.white });
  const runs = [];
  side.forEach((l, i) => {
    const [t, st] = typeof l === "string" ? [l, ""] : l;
    runs.push({ text: t, options: { bullet: st !== "h", bold: st === "b" || st === "h", color: st === "h" ? "F2C6AC" : "D9D9D9",
      fontFace: st === "c" ? MONO : FONT, paraSpaceBefore: i ? 7 : 0, breakLine: i < side.length - 1 } });
  });
  txt(s, runs, { x: 9.0, y: 2.25, w: 3.4, h: 4.5, fontSize: 13.5, valign: "top" });
  return s;
}

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
    { text: "Class 7:", options: { bullet: true, fontSize: 30, breakLine: true } },
    { text: "Images for the web: file types, colour depth, size and compression.", options: { bullet: { code: "25CB" }, indentLevel: 1, fontSize: 26, paraSpaceBefore: 6 } },
  ], { x: 6.85, y: 1.9, w: 5.6, h: 2.6, color: C.black, valign: "top" });
  s.addNotes("Welcome back. Last class we planned a page with a wireframe, and every wireframe was full of X boxes: images. Today we learn how to make those images fast.\n\nToday is mostly hands-on. There are six short labs. Students need a large photo (from their phone is best) and a browser for pixlr.com.");
}

// 2. Attendance
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  txt(s, "Attendance", { x: 1.1, y: 0.75, w: 8, h: 0.9, fontSize: 44, color: C.black });
  s.addNotes("Take attendance. While you do, ask everyone to get a photo from their phone onto their laptop (OneDrive, email it to yourself, or a USB cable). They'll need it for Lab 1.");
}

// 3. Overview
{
  const s = panelSlide("Class 7: Class Overview",
    "Roughly: 10 min pixels and colour depth + Lab 1, 15 min compression and formats, then Labs 2–5 in Pixlr (about 45 min, with a short concept slide before each), 15 min putting images on a page (Lab 6), 5 min wrap-up.\n\nEverything students make today gets used next class, when we style images with CSS.");
  txt(s, [
    { text: "Title: ", options: { bold: true, bullet: true, fontSize: 22 } },
    { text: "Images for the Web", options: { fontSize: 22, breakLine: true } },
    { text: " ", options: { fontSize: 10, breakLine: true } },
    { text: "Agenda:", options: { bold: true, bullet: true, fontSize: 22, breakLine: true } },
    { text: "Pixels, dimensions and colour depth", options: { bullet: true, indentLevel: 1, fontSize: 18, paraSpaceBefore: 4, breakLine: true } },
    { text: "Compression: lossless vs lossy", options: { bullet: true, indentLevel: 1, fontSize: 18, paraSpaceBefore: 4, breakLine: true } },
    { text: "File formats: JPG, PNG, GIF, WebP, AVIF, SVG", options: { bullet: true, indentLevel: 1, fontSize: 18, paraSpaceBefore: 4, breakLine: true } },
    { text: "Six labs: test it yourself in Pixlr and File Explorer", options: { bullet: true, indentLevel: 1, fontSize: 18, paraSpaceBefore: 4, breakLine: true } },
    { text: "Images in HTML: width, height, lazy loading, srcset", options: { bullet: true, indentLevel: 1, fontSize: 18, paraSpaceBefore: 4, breakLine: true } },
    { text: "Choosing the right format: a checklist", options: { bullet: true, indentLevel: 1, fontSize: 18, paraSpaceBefore: 4 } },
  ], { x: CX, y: 1.95, w: CW, h: 4.8, color: C.black, valign: "top" });
}

// 4. Get ready
{
  const s = panelSlide("Before We Start: Get Ready",
    "The best test photo is a straight-off-the-phone photo: 3,000–4,000+ pixels wide and 2–5 MB. That's what makes the size differences dramatic.\n\nIf a student has no photo: unsplash.com or pexels.com have free photos. Download the largest ('Original') size.\n\nPixlr: pixlr.com, then Pixlr Editor (the full editor). No account needed to start, but CHECK BEFORE CLASS: Pixlr's free tier has at times limited how many saves you get per day. If students hit a limit, use squoosh.app instead (free, no account, made by Google, shows the file size live).");
  txt(s, [{ text: "1.  Pull this week's examples", options: { bold: true, fontSize: 18 } }], { x: CX, y: 1.95, w: CW, h: 0.4 });
  codeBox(s, ["git pull"], { x: CX, y: 2.45, w: 5.2, h: 0.7, maxSize: 16 });
  txt(s, "New folder: 04_Images_for_the_Web (examples 01–05, samples/, image_lab.html, my_images/).", { x: CX, y: 3.25, w: CW, h: 0.5, fontSize: 15, color: C.dark });
  txt(s, [{ text: "2.  Get a big photo onto your laptop", options: { bold: true, fontSize: 18 } }], { x: CX, y: 3.95, w: CW, h: 0.4 });
  txt(s, "A photo from your phone is best (OneDrive, email, or a cable). No photo? Download a free one at full size from unsplash.com. Save it as my_images/original.jpg.", { x: CX, y: 4.4, w: CW, h: 0.8, fontSize: 15, color: C.dark });
  txt(s, [{ text: "3.  Open Pixlr", options: { bold: true, fontSize: 18 } }], { x: CX, y: 5.4, w: CW, h: 0.4 });
  txt(s, "pixlr.com → Pixlr Editor. Backup if Pixlr won't save: squoosh.app", { x: CX, y: 5.85, w: CW, h: 0.5, fontSize: 15, color: C.dark });
}

// 5. Why images matter
{
  const s = panelSlide("Why Do Images Matter So Much?",
    "Images are usually the heaviest part of a web page. The HTML for a whole page might be 20–50 KB. One unoptimized phone photo can be 3–5 MB, about 100 times more.\n\nGoogle measures 'Largest Contentful Paint' (LCP): how long until the biggest thing on screen appears. On most pages the biggest thing is an image. Slow LCP hurts the user and can hurt search ranking.\n\nConnect to non-functional requirements from earlier: 'Pages load in under 2 seconds on a typical connection.' Image choices are how you meet that requirement.");
  const cards = [
    ["Speed", "Images are usually the heaviest part of a page. Smaller images = faster pages.", C.orange],
    ["Mobile data", "A 5 MB photo on a phone plan costs the visitor real money and time.", C.teal],
    ["Search ranking", "Google measures how fast the biggest image appears (LCP).", C.purple],
    ["Requirements", "\"Pages load in under 2 seconds\" [NF]. Images are how you meet it.", C.blue],
  ];
  const cw = 3.65, ch = 2.1;
  cards.forEach(([t, d, col], i) => {
    const x = CX + (i % 2) * (cw + 0.25), y = 2.0 + Math.floor(i / 2) * (ch + 0.25);
    card(s, { x, y, w: cw, h: ch });
    numCircle(s, i + 1, x + 0.2, y + 0.22, col, 0.55);
    txt(s, t, { x: x + 0.9, y: y + 0.22, w: cw - 1.0, h: 0.55, fontSize: 18, bold: true, color: C.black, valign: "middle" });
    txt(s, d, { x: x + 0.25, y: y + 0.95, w: cw - 0.45, h: 1.0, fontSize: 14, color: C.grey, valign: "top" });
  });
  card(s, { x: CX, y: 6.5, w: CW, h: 0.6, fill: C.dark, border: C.dark });
  txt(s, "Goal: the smallest file that still looks good at the size it's shown.", { x: CX + 0.3, y: 6.5, w: CW - 0.6, h: 0.6, fontSize: 15, bold: true, color: C.white, valign: "middle" });
}

// 6. Pixels and dimensions
{
  const s = panelSlide("Pixels and Dimensions",
    "Zoom into any photo far enough and you see it's a grid of coloured squares: pixels. A 'raster' image (JPG, PNG, GIF, WebP, AVIF) is exactly that grid, stored as numbers.\n\nDimensions = width × height in pixels. A 12-megapixel phone photo is about 4032 × 3024 = 12.2 million pixels.\n\nDPI / PPI: you'll see '72 dpi' or '300 dpi' in Properties. That number matters for PRINTING. On the web it's ignored: a 1200-pixel-wide image is 1200 pixels wide whether it says 72 or 300 dpi. Only the pixel dimensions matter.");
  s.addImage({ path: path.join(SCR, "zoom_q90.png"), x: CX, y: 1.95, w: 3.4, h: 2.27 });
  s.addShape(pres.shapes.RECTANGLE, { x: CX, y: 1.95, w: 3.4, h: 2.27, fill: { type: "none" }, line: { color: C.mid, width: 1 } });
  txt(s, "Zoomed in 5×: every square is one pixel.", { x: CX, y: 4.3, w: 3.4, h: 0.35, fontSize: 12.5, italic: true, color: C.grey });
  richBullets(s, [
    [["Raster", "b"], " images are a grid of ", ["pixels", "b"], ". Each pixel is one colour."],
    [["Dimensions", "b"], " = width × height in pixels, e.g. ", ["4032 × 3024", "b"], "."],
    [["Megapixels", "b"], " = width × height ÷ 1 million. That one is 12 MP."],
  ], { x: CX + 3.7, y: 1.95, w: CW - 3.7, h: 2.7, size: 14.5 });
  card(s, { x: CX, y: 4.9, w: CW, h: 1.85 });
  numCircle(s, "!", CX + 0.2, 5.1, C.purple, 0.45);
  txt(s, "What about DPI?", { x: CX + 0.8, y: 5.1, w: CW - 1, h: 0.45, fontSize: 16, bold: true, color: C.black, valign: "middle" });
  richBullets(s, [["DPI (dots per inch) only matters for ", ["printing", "b"], ". Browsers ignore it. A 1200-pixel-wide image is 1200 pixels wide at 72 dpi or 300 dpi. ", ["On the web, only pixels count.", "b"]]],
    { x: CX + 0.25, y: 5.65, w: CW - 0.5, h: 1.0, size: 14, bullet: false });
}

// 7. Colour depth
{
  const s = panelSlide("Colour Depth (Bit Depth)",
    "Colour depth = how many bits are used to store the colour of ONE pixel. More bits = more possible colours = bigger files (before compression).\n\n8 bits per channel × 3 channels (red, green, blue) = 24-bit colour: 256 × 256 × 256 = about 16.7 million colours. That's what JPG uses and what screens show.\n\n32-bit = 24-bit colour + 8 bits of 'alpha' (transparency): each pixel can be anywhere from fully see-through to solid.\n\nThe images on the slide are the same photo: 24-bit, then GIF with 256 colours (look at the sky: banding), then only 16 colours. In File Explorer → Properties → Details this is called 'Bit depth'.");
  const imgs = [["depth_beach_original.png", "24-bit: 16.7 million colours"], ["depth_beach.png", "8-bit: 256 colours (GIF)"], ["depth_beach_16_colours.png", "4-bit: 16 colours"]];
  const iw = 2.4, ih = iw * 321 / 560, gap = (CW - 3 * iw) / 2;
  imgs.forEach(([f, cap], i) => {
    const x = CX + i * (iw + gap);
    s.addImage({ path: path.join(SCR, f), x, y: 1.95, w: iw, h: ih });
    txt(s, cap, { x, y: 1.95 + ih + 0.07, w: iw, h: 0.3, fontSize: 12, bold: true, color: C.dark, align: "center" });
  });
  table(s, [
    ["Bits per pixel", "Colours", "Where you see it"],
    ["1-bit", "2", "Black and white only"],
    ["8-bit", "256", "GIF, PNG-8"],
    ["24-bit", "16.7 million", "JPG, PNG-24, WebP: photos"],
    ["32-bit", "16.7 million + transparency", "PNG-32, WebP, AVIF with alpha"],
  ], { x: CX, y: 3.75, w: CW, colW: [1.7, 2.75, 3.1], size: 13.5, rowH: 0.42 });
  txt(s, "Alpha channel = the extra 8 bits that store how see-through each pixel is.", { x: CX, y: 6.0, w: CW, h: 0.4, fontSize: 14, italic: true, color: C.grey });
}

// 8. Do the math
{
  const s = panelSlide("Do the Math: Why We Need Compression",
    "Uncompressed size = width × height × bytes per pixel. 24-bit colour = 3 bytes per pixel.\n\nOur sample: 1118 × 640 × 3 = about 2 MB uncompressed, but the JPG file is only about 134 KB. Compression made it ~15× smaller.\n\nA phone photo: 4032 × 3024 × 3 = 36.6 MB uncompressed. The phone's JPG is usually 2–5 MB. Still WAY too big for a web page, which is why we'll resize too.\n\n1 KB = 1024 bytes, 1 MB = 1024 KB. (Windows shows sizes this way; macOS uses 1000. Close enough for us.)");
  txt(s, "Uncompressed = width × height × 3 bytes", { x: CX, y: 1.95, w: CW, h: 0.55, fontSize: 19, bold: true, color: C.black, fontFace: MONO });
  txt(s, "(24-bit colour = 3 bytes per pixel: one each for red, green and blue)", { x: CX, y: 2.5, w: CW, h: 0.35, fontSize: 13.5, color: C.grey });
  const rows = [
    ["Our sample beach photo", `${S.width} × ${S.height} × 3`, KB(S.uncompressed), KB(S.original), C.teal],
    ["A 12 MP phone photo", "4032 × 3024 × 3", "36.6 MB", "about 3 MB", C.orange],
  ];
  rows.forEach(([name, calc, raw, file, col], i) => {
    const y = 3.1 + i * 1.6, h = 1.35;
    card(s, { x: CX, y, w: CW, h });
    s.addShape(pres.shapes.RECTANGLE, { x: CX, y: y + 0.12, w: 0.1, h: h - 0.24, fill: { color: col }, line: { type: "none" } });
    txt(s, name, { x: CX + 0.3, y: y + 0.15, w: 3.2, h: 0.4, fontSize: 16, bold: true, color: C.black });
    txt(s, calc, { x: CX + 0.3, y: y + 0.6, w: 3.2, h: 0.4, fontSize: 14, color: C.dark, fontFace: MONO });
    txt(s, [{ text: "Uncompressed", options: { fontSize: 11, color: C.grey, breakLine: true } }, { text: raw, options: { fontSize: 22, bold: true, color: C.black } }],
      { x: CX + 3.6, y: y + 0.2, w: 1.8, h: 0.95, valign: "middle" });
    txt(s, "→", { x: CX + 5.25, y, w: 0.4, h, fontSize: 22, bold: true, color: C.orange, align: "center", valign: "middle" });
    txt(s, [{ text: "Actual JPG file", options: { fontSize: 11, color: C.grey, breakLine: true } }, { text: file, options: { fontSize: 22, bold: true, color: col } }],
      { x: CX + 5.75, y: y + 0.2, w: 1.75, h: 0.95, valign: "middle" });
  });
  txt(s, "Compression makes files much smaller. But 3 MB is still far too big for one image on a web page.", { x: CX, y: 6.35, w: CW, h: 0.6, fontSize: 15, bold: true, color: C.teal });
}

// 9. Lab 1
labSlide({
  title: "Lab 1: Get to Know Your Photo (5 min)",
  steps: [
    ["Save your photo as my_images/original.jpg", "Inside the 04_Images_for_the_Web folder."],
    ["Right-click it → Properties → Details tab", "Scroll to Image: find Dimensions and Bit depth. File size is on the General tab."],
    ["Do the math", [["width × height × 3", "c"], " = the uncompressed size in bytes. Divide by 1,048,576 for MB."]],
    ["Fill in Lab 1 in image_lab.html", "How many times smaller is the real file than the uncompressed size?"],
    ["Bonus: set up a better File Explorer view", [["View → Details", "b"], ", then right-click a column heading → More… → tick ", ["Dimensions", "b"], " and ", ["Bit depth", "b"], ". You'll use this all class."]],
  ],
  sideTitle: "Write down",
  side: [
    "Dimensions (W × H)", "Bit depth", "File size", "Uncompressed size",
    ["", "h"],
    ["Expect roughly", "h"], "Phone: 4000 × 3000, 24-bit, 2–5 MB", "Uncompressed: 30–50 MB",
  ],
  notes: "Walk around. Common snags:\n- iPhone photos may be .HEIC, not .jpg. Fix: on the phone, Settings → Camera → Formats → Most Compatible, or open the HEIC in Pixlr and save as JPG first. Windows may need the free HEIF extension to show HEIC details.\n- Photos sent through some chat apps arrive already shrunk (e.g. 1600 px). That's fine, but the differences will be smaller.\n\nThe File Explorer Details view with Dimensions and Bit depth columns is the fastest way to compare files for the rest of the class. Show it on the projector.\n\nAsk a couple of students for their compression ratio. Usually 10–20×.",
});

// 10. Divider: compression
dividerSlide("Compression", "Lossless vs lossy   ·   what you lose and what you save",
  "Next: how files get smaller. Two kinds of compression, and the difference decides which file type to use.");

// 11. Lossless vs lossy
{
  const s = panelSlide("Lossless vs Lossy",
    "Lossless: like a ZIP file. The file is packed more cleverly (e.g. 'the next 200 pixels are all this exact blue' instead of listing each one), and when you open it you get back EXACTLY the original pixels. PNG and GIF are lossless. Works great on flat colour; not much help on photos where every pixel is slightly different.\n\nLossy: throws away detail your eye is unlikely to notice: tiny colour differences, fine texture. You choose how much with the quality slider. You can never get that detail back, which is why you always keep the original and export copies.\n\nGeneration loss: open a JPG, save it as JPG, repeat. Each save loses a bit more. Edit from the original.");
  const cards = [
    ["Lossless", C.teal, "Packs the data more cleverly. Open it and every pixel is exactly the same as before.",
      ["Like a ZIP file", "PNG, GIF, lossless WebP", "Great for logos, icons, screenshots, flat colour", "Poor for photos (big files)"]],
    ["Lossy", C.orange, "Throws away detail your eye probably won't notice. You choose how much.",
      ["Like summarizing a book", "JPG, WebP, AVIF", "Great for photos (tiny files)", "Too much = blurry, blocky pictures. You can't undo it"]],
  ];
  const cw = 3.65;
  cards.forEach(([t, col, d, pts], i) => {
    const x = CX + i * (cw + 0.25), y = 1.95;
    card(s, { x, y, w: cw, h: 4.4 });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: cw, h: 0.65, fill: { color: col }, line: { type: "none" } });
    txt(s, t, { x: x + 0.25, y, w: cw - 0.5, h: 0.65, fontSize: 20, bold: true, color: C.white, valign: "middle" });
    txt(s, d, { x: x + 0.25, y: y + 0.85, w: cw - 0.5, h: 1.0, fontSize: 14, color: C.dark, valign: "top" });
    richBullets(s, pts.map((p) => [p]), { x: x + 0.25, y: y + 1.95, w: cw - 0.5, h: 2.3, size: 13.5, space: 7 });
  });
  card(s, { x: CX, y: 6.5, w: CW, h: 0.6, fill: C.dark, border: C.dark });
  txt(s, "Always keep your original. Export copies. Never edit and re-save the same JPG over and over.", { x: CX + 0.3, y: 6.5, w: CW - 0.6, h: 0.6, fontSize: 14, bold: true, color: C.white, valign: "middle" });
}

// 12. Seeing lossy compression
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  txt(s, "What Lossy Compression Looks Like", { x: 0.6, y: 0.4, w: 12, h: 0.65, fontSize: 32, color: C.black });
  txt(s, "The same JPG at three quality settings, zoomed in 5×.", { x: 0.6, y: 1.05, w: 12, h: 0.4, fontSize: 16, color: C.grey });
  const shots = [["zoom_q90.png", "Quality 90", S.q90, C.green], ["zoom_q50.png", "Quality 50", S.q50, C.orange], ["zoom_q10.png", "Quality 10", S.q10, C.red]];
  const iw = 3.85, ih = iw * 400 / 600, gap = (12.1 - 3 * iw) / 2;
  shots.forEach(([f, t, kb, col], i) => {
    const x = 0.6 + i * (iw + gap);
    s.addImage({ path: path.join(SCR, f), x, y: 1.7, w: iw, h: ih });
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.7, w: iw, h: ih, fill: { type: "none" }, line: { color: C.mid, width: 1 } });
    txt(s, t, { x, y: 1.75 + ih + 0.1, w: 2.5, h: 0.45, fontSize: 20, bold: true, color: C.black });
    txt(s, KB(kb), { x: x + iw - 1.6, y: 1.75 + ih + 0.1, w: 1.6, h: 0.45, fontSize: 20, bold: true, color: col, align: "right" });
  });
  richBullets(s, [
    [["Blocks: ", "b"], "JPG works in 8 × 8 pixel squares. Push the quality too low and you can see them."],
    [["Smudges and noise ", "b"], "appear around sharp edges, and smooth areas like sky turn patchy."],
    [["Quality 90 is bigger than the original file (", "b"], ["" + KB(S.original), "b"], [")! ", "b"], "Re-saving at a higher quality can't bring back detail. It only adds bytes."],
  ], { x: 0.6, y: 5.0, w: 12.1, h: 2.1, size: 15.5 });
  s.addNotes(`These are real crops from samples/beach_q90.jpg, beach_q50.jpg and beach_q10.jpg (Example 02).\n\nAt normal size, quality 50 looks fine to most people. Zoomed in, you start to see smudging. Quality 10 is obviously blocky even at normal size.\n\nThe surprising one: the original file is ${KB(S.original)}, and re-saving it at quality 90 made it ${KB(S.q90)}. Bigger, not better, because the detail was already gone. Lesson: the quality number isn't a promise of quality; it's how much MORE you're willing to throw away.\n\nMost photos look good at quality 60–80. That's where we'll aim in Lab 3.`);
}

// 13. Divider: formats
dividerSlide("File Formats", "JPG · PNG · GIF · WebP · AVIF · SVG",
  "Six formats. Students only really need four day to day: JPG or WebP for photos, PNG for screenshots and graphics with transparency, SVG for logos and icons.");

// 14. The six formats
{
  const s = panelSlide("The Six Formats You'll Meet",
    "Quick tour. Dates help explain why things are the way they are:\n- GIF (1987): only 256 colours. Still around because of animated memes, but a video file is much smaller for animation.\n- JPEG (1992): the classic photo format. Every device and program supports it.\n- PNG (1996): created as a free replacement for GIF. Lossless, full colour, real transparency.\n- SVG (2001): vector graphics written as XML code. Can even be written by hand or styled with CSS.\n- WebP (Google, 2010): lossy or lossless, transparency, animation. Works in all current browsers.\n- AVIF (2019): newest, usually the smallest for photos. Works in all current major browsers; some older software (and some image editors) can't open it yet.");
  const fmts = [
    [".jpg", "JPEG", "Photos. Lossy, millions of colours, no transparency.", C.orange],
    [".png", "PNG", "Screenshots, logos, anything needing transparency. Lossless.", C.teal],
    [".gif", "GIF", "Simple animations. Only 256 colours. Mostly old-fashioned now.", C.mid],
    [".webp", "WebP", "Modern photos AND graphics. Smaller than JPG/PNG. Transparency too.", C.purple],
    [".avif", "AVIF", "Newest. Usually the smallest photos. Some older apps can't open it.", C.navy],
    [".svg", "SVG", "Logos, icons. Vector: shapes in code, sharp at any size.", C.blue],
  ];
  const cw = 3.65, ch = 1.45;
  fmts.forEach(([ext, name, d, col], i) => {
    const x = CX + (i % 2) * (cw + 0.25), y = 1.95 + Math.floor(i / 2) * (ch + 0.18);
    card(s, { x, y, w: cw, h: ch });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.2, y: y + 0.22, w: 0.95, h: 0.42, rectRadius: 0.08, fill: { color: col }, line: { type: "none" } });
    txt(s, ext, { x: x + 0.2, y: y + 0.22, w: 0.95, h: 0.42, fontSize: 13, bold: true, color: C.white, align: "center", valign: "middle", fontFace: MONO });
    txt(s, name, { x: x + 1.3, y: y + 0.22, w: cw - 1.5, h: 0.42, fontSize: 17, bold: true, color: C.black, valign: "middle" });
    txt(s, d, { x: x + 0.2, y: y + 0.75, w: cw - 0.4, h: 0.65, fontSize: 12.5, color: C.grey, valign: "top" });
  });
}

// 15. Comparison table
{
  const s = panelSlide("Formats Side by Side",
    `Sizes in the last column are our sample photo (${S.width} × ${S.height}) from Example 01. Same pixels, very different files.\n\nThe headline: PNG is lossless, so for a photo it's huge (${KB(S.png)}). GIF is smaller than PNG only because it threw away colours, and the sky shows it. WebP and AVIF beat JPG at similar visual quality.\n\nTransparency column: JPG can't do it. If you save a logo with a transparent background as JPG, the background turns solid (usually white).`);
  table(s, [
    ["Format", "Compression", "Colours", "Transparency", "Animation", "Our photo"],
    ["JPG", "Lossy", "16.7 M", "✗", "✗", KB(S.q75)],
    ["PNG", "Lossless", "16.7 M", "✓", "✗", KB(S.png)],
    ["GIF", "Lossless", "256", "On / off only", "✓", KB(S.gif)],
    ["WebP", "Both", "16.7 M", "✓", "✓", KB(S.webp)],
    ["AVIF", "Both", "16.7 M+", "✓", "✓", KB(S.avif)],
    ["SVG", "Vector (code)", "Any", "✓", "✓", "n/a"],
  ], { x: CX, y: 1.95, w: CW, colW: [0.95, 1.4, 1.05, 1.55, 1.2, 1.4], size: 13.5, rowH: 0.5 });
  richBullets(s, [
    [["Photo? ", "b"], "JPG, or WebP / AVIF for smaller files."],
    [["Needs transparency? ", "b"], "PNG or WebP. Never JPG."],
    [["Logo or icon? ", "b"], "SVG if you have it."],
  ], { x: CX, y: 5.65, w: CW, h: 1.4, size: 15, space: 5 });
}

// 16. Example 01
codeSlide({
  kicker: "EXAMPLE 01", title: "One Photo, Six Files", file: "01_image_formats_example.html", lines: bodyOf("01_image_formats_example.html"), size: 14,
  bullets: [
    ["The same photo in five raster formats. Open it with Live Server and compare. Most look identical."],
    ["Now open the ", ["samples", "c"], " folder in File Explorer (Details view) and compare the sizes: ", [`${KB(S.avif)} → ${KB(S.png)}`, "b"], "."],
    ["Look closely at the ", ["GIF", "b"], "'s sky. Only 256 colours means smooth gradients turn into bands."],
    [["width", "c"], " and ", ["height", "c"], " match the shape of the photo. More on why later."],
  ],
  notes: "Open 01_image_formats_example.html with Live Server.\n\nAsk: 'Which one is the PNG?' Most students can't tell by looking. Then show the samples folder in File Explorer, sorted by size. That's the 'aha'.\n\nThe GIF is the only one with a visible difference (banding in the sky). Zoom the browser in to make it obvious.\n\nAVIF note: if a student's browser shows a broken image for AVIF, their browser is out of date. All current Chrome, Edge, Firefox and Safari versions support it.",
});

// 17. Raster vs vector
{
  const s = panelSlide("Raster vs Vector",
    `Raster = a grid of pixels (everything we've seen so far). Make it bigger than its pixel size and the browser has to stretch the pixels, so it goes soft or blocky.\n\nVector = a description of shapes: 'circle at 120,120, radius 112, navy'. The browser redraws it at whatever size you ask for, so it's always sharp. SVG is the web's vector format, and it's actually text (open logo.svg in VS Code!).\n\nOur logo: SVG ${KB(S.logoSvg)} vs PNG ${KB(S.logoPng)} vs JPG ${KB(S.logoJpg)}. The SVG is about 30 times smaller and sharper at every size.\n\nVector doesn't work for photos. There are no simple shapes to describe, so photos stay raster.`);
  const iw = 3.6, ih = iw * 440 / 640;
  [["zoom_logo_png.png", "PNG, zoomed: clean edges", C.teal], ["zoom_logo_jpg.png", "JPG, zoomed: smudges around edges", C.orange]].forEach(([f, cap, col], i) => {
    const x = CX + i * (iw + 0.35);
    s.addImage({ path: path.join(SCR, f), x, y: 1.95, w: iw, h: ih });
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.95, w: iw, h: ih, fill: { type: "none" }, line: { color: C.mid, width: 1 } });
    txt(s, cap, { x, y: 1.95 + ih + 0.07, w: iw, h: 0.3, fontSize: 12.5, bold: true, color: col });
  });
  richBullets(s, [
    [["Raster", "b"], " (JPG, PNG, GIF, WebP, AVIF): a grid of pixels. Stretch it bigger and it goes ", ["soft or blocky", "b"], ". Best for photos."],
    [["Vector", "b"], " (SVG): shapes described in code. ", ["Sharp at any size", "b"], ", and tiny. Best for logos, icons and simple drawings."],
    ["Our logo: ", [`SVG ${KB(S.logoSvg)}`, "b"], " · PNG " + KB(S.logoPng) + " · JPG " + KB(S.logoJpg) + " (and the JPG lost its transparency)."],
  ], { x: CX, y: 5.0, w: CW, h: 2.1, size: 15 });
}

// 18. Example 03
codeSlide({
  kicker: "EXAMPLE 03", title: "Logos: SVG vs PNG vs JPG", file: "03_svg_vs_png_example.html", lines: bodyOf("03_svg_vs_png_example.html"), size: 14,
  bullets: [
    ["At 100 pixels wide, all three look about the same."],
    ["At 900 pixels wide: the ", ["SVG", "b"], " is still sharp, the ", ["PNG", "b"], " (600 px) is stretched and soft, and the ", ["JPG", "b"], " has smudgy edges and ", ["white corners", "b"], " (no transparency)."],
    ["An SVG is just text. Open ", ["samples/logo.svg", "c"], " in VS Code: it's a circle, an ellipse and a few paths."],
    ["The ", ["<style>", "c"], " line is a sneak peek at next class: CSS."],
  ],
  notes: "Open 03_svg_vs_png_example.html with Live Server. The page background is light blue-grey so the transparency shows: the SVG and PNG have see-through corners, the JPG has white ones.\n\nThen open samples/logo.svg in VS Code. Point out it's readable code with a <title> (for accessibility). Change a fill colour (e.g. #E97132 to #2E7D32), save, and refresh: the logo changes colour. That's something you can't do with a PNG.\n\nBrowser zoom (Ctrl +) to 300% makes the difference between SVG and PNG very obvious.",
});

// 19. Lab 2: format shootout
labSlide({
  title: "Lab 2: Format Shootout in Pixlr (15 min)",
  steps: [
    ["Open your photo in Pixlr Editor", "pixlr.com → Pixlr Editor → Open image → my_images/original.jpg."],
    ["Save as PNG", [["File → Save", "b"], " (Ctrl + S). Pick ", ["PNG", "b"], ". Name it ", ["photo.png", "c"], " in my_images."]],
    ["Save as JPG, quality 80", ["Same again: ", ["JPG", "b"], ", slide quality to ", ["80", "b"], ". Name it ", ["photo_q80.jpg", "c"], "."]],
    ["Save as WebP, quality 80", ["Name it ", ["photo.webp", "c"], ". Keep the full size for all three."]],
    ["Compare in File Explorer", "Details view, sorted by Size. Fill in the table in image_lab.html."],
  ],
  sideTitle: "Questions",
  side: [
    "Which file is biggest? How many times bigger than the JPG?",
    "Open all three. Can you see a difference?",
    "Which would you put on a website?",
    ["", "h"],
    ["Pixlr won't save?", "h"], "Use squoosh.app: drop in your photo, pick a format on the right, and it shows the new size live.",
  ],
  notes: "Expect: PNG is enormous, often bigger than the original JPG (e.g. a 3 MB JPG becomes a 15–25 MB PNG). JPG at 80 is usually a bit smaller than the phone original. WebP at 80 is usually smaller again.\n\nPixlr's save dialog lets you pick the file type and quality, and usually shows an estimated file size. Menus in Pixlr change from time to time; check them before class.\n\nIf Pixlr has a daily save limit on the free tier when you teach this, switch the class to squoosh.app for the remaining labs. It does format, quality AND resize in one screen.",
});

// 20. Example 02
codeSlide({
  kicker: "EXAMPLE 02", title: "How Low Can You Go?", file: "02_jpeg_quality_example.html", lines: bodyOf("02_jpeg_quality_example.html"), size: 14,
  bullets: [
    ["The quality slider (0–100) decides ", ["how much detail lossy compression throws away", "b"], "."],
    [["60–80", "b"], " is the sweet spot for most photos: much smaller, and you can't see the difference at normal size."],
    ["Below about 40, problems start to show: blocks, smudges and patchy sky."],
    ["Quality 90 here is ", ["bigger", "b"], " than the original. Always start from the original, never a copy you already compressed."],
  ],
  notes: `Open 02_jpeg_quality_example.html. At normal size ask: which ones look bad? Usually only quality 10. Then Ctrl + to zoom the browser and look again.\n\nSizes: q90 ${KB(S.q90)}, q75 ${KB(S.q75)}, q50 ${KB(S.q50)}, q10 ${KB(S.q10)}. Going from 75 to 50 roughly halves the file with little visible change on this photo.\n\nDifferent photos react differently: busy, detailed photos (grass, gravel) hide compression well; smooth areas (sky, skin, gradients) show it first. That's why we test with our own photos next.`,
});

// 21. Lab 3: quality slider
labSlide({
  title: "Lab 3: Find the Sweet Spot (10 min)",
  steps: [
    ["Start from original.jpg again", "Never from a file you already compressed."],
    ["Save as JPG at quality 30", ["Name it ", ["photo_q30.jpg", "c"], "."]],
    ["Save as JPG at quality 10", ["Name it ", ["photo_q10.jpg", "c"], "."]],
    ["Zoom in on all of them", "Open each one and zoom to 200–400%. Look at sky, skin, edges and text."],
    ["Record sizes and \"Looks OK?\"", "In image_lab.html. You already have q80 from Lab 2."],
  ],
  sideTitle: "Look for",
  side: [
    "Blocky squares in smooth areas (sky, walls)",
    "Smudges or halos around edges",
    "Colour bands instead of smooth gradients",
    ["", "h"],
    ["Challenge", "h"],
    "Find the lowest quality where you honestly can't see a difference at normal size. Share your number in the class chat.",
  ],
  notes: "Students often find 60–70 is indistinguishable at normal size. Collect a few numbers in the Teams chat: they'll differ by photo, which is the point. Busy photos survive lower quality; photos with lots of sky or a plain background show problems sooner.\n\nIf someone has time: try WebP at 30 vs JPG at 30. WebP usually holds up better at low quality.",
});

// 22. The right size
{
  const s = panelSlide("Size It Right: Dimensions Matter Most",
    "The biggest win isn't the format or the quality slider. It's not sending pixels nobody will see.\n\nIf your page shows an image 800 pixels wide, a 4032-pixel photo sends about 25 times more pixels than needed (4032×3024 vs 800×600). The browser downloads all of it and then shrinks it.\n\nRule of thumb: export at the largest size it's ever displayed, times 2 for sharp screens (phones, MacBooks and many laptops have 2× 'high-DPI' screens). So an image shown at 600 px wide → export 1200 px.\n\nThe width/file size table is a rule of thumb, not a law. Use it as a starting point.\n\nAlso halving the width quarters the pixels (width AND height both halve), so file size drops roughly 4×.");
  richBullets(s, [
    [["Export at the size it's shown, × 2", "b"], " for sharp phone and laptop screens. Shown at 600 px wide → export about 1200 px."],
    ["Halve the width → ", ["¼ of the pixels", "b"], " (the height halves too) → roughly ¼ of the file size."],
  ], { x: CX, y: 1.9, w: CW, h: 1.35, size: 15.5 });
  table(s, [
    ["Image type", "Export width", "Aim for"],
    ["Full-width banner (hero)", "1600–2000 px", "under 300 KB"],
    ["Content photo", "1000–1200 px", "under 150 KB"],
    ["Card / thumbnail", "400–600 px", "under 50 KB"],
    ["Logo / icon", "SVG (any size)", "a few KB"],
  ], { x: CX, y: 3.35, w: CW, colW: [3.0, 2.3, 2.25], size: 14.5, rowH: 0.48 });
  card(s, { x: CX, y: 5.95, w: CW, h: 0.95, fill: "FBE9DF", border: "F2C6AC" });
  txt(s, [{ text: "Our sample: ", options: { bold: true } }, { text: `1118 px wide = ${KB(S.original)}. Resized to 800 px = ${KB(S.w800)}. Resized to 400 px = ${KB(S.w400)}.` }],
    { x: CX + 0.3, y: 5.95, w: CW - 0.6, h: 0.95, fontSize: 14.5, color: C.dark, valign: "middle" });
}

// 23. Lab 4: resize
labSlide({
  title: "Lab 4: Resize and Compare (10 min)",
  steps: [
    ["Open original.jpg in Pixlr", "Start from the original again."],
    ["Resize to 2000 px wide", ["Find ", ["Resize", "b"], " (Pixlr Editor: the Page menu, or set the size in the Save dialog). Keep ", ["proportions locked", "b"], ". Save as JPG q80: ", ["photo_2000.jpg", "c"], "."]],
    ["Repeat at 1200 px and 600 px", ["Start from the original each time: ", ["photo_1200.jpg", "c"], ", ", ["photo_600.jpg", "c"], "."]],
    ["Compare in File Explorer", "Check the Dimensions and Size columns. Fill in the table."],
    ["Open original.jpg and photo_1200.jpg", "Fit both to your screen. Can you tell them apart?"],
  ],
  sideTitle: "Questions",
  side: [
    "How much smaller is photo_1200 than the original?",
    "Halving the width: did the size drop by about ½ or ¼?",
    "Which size would you use for a content photo on a page?",
    ["", "h"],
    ["Watch out", "h"],
    "Never make a small image bigger. You can't add pixels that aren't there. It just goes blurry.",
  ],
  notes: "Typical result from a 4032 px phone photo: 2000 px ≈ 600–900 KB, 1200 px ≈ 200–350 KB, 600 px ≈ 60–100 KB. That's 10–40× smaller than the original, and at normal viewing size it looks the same.\n\nThe 'lock proportions' icon (chain link) matters: unlock it and the photo gets squashed.\n\nThis is usually the biggest 'aha' of the day: resizing saves far more than switching formats.",
});

// 24. Lab 5: graphics vs photos
labSlide({
  title: "Lab 5: Graphics vs Photos (5–10 min)",
  steps: [
    ["Open samples/logo.png in Pixlr", "A flat-colour logo with a transparent background."],
    ["Save it as JPG, quality 80", ["Name it ", ["my_images/logo_test.jpg", "c"], "."]],
    ["Compare logo.png, logo_test.jpg and logo.svg", "Sizes in File Explorer. Then open 03_svg_vs_png_example.html."],
    ["Zoom in on the edges", "Where the orange fish meets the navy circle. And look at the corners."],
    ["Take a screenshot (Win + Shift + S)", "Paste it into Pixlr. Save it as PNG and as JPG q80. Which is smaller? Which is sharper?"],
  ],
  sideTitle: "What you'll find",
  side: [
    "The JPG lost the transparent corners (now white)",
    "Smudges along every edge",
    "The JPG isn't even smaller",
    "SVG: smallest and sharpest",
    ["", "h"],
    ["The rule", "h"],
    ["Photos → JPG / WebP. Flat colour, text, screenshots → PNG. Logos and icons → SVG.", "b"],
  ],
  notes: `Our results: logo.svg ${KB(S.logoSvg)}, logo.png ${KB(S.logoPng)}, logo.jpg (quality 80) ${KB(S.logoJpg)}. The JPG is no smaller AND it looks worse AND it lost transparency.\n\nWhy: JPG is designed for the gradual changes in photos. Hard edges between flat colours are the worst case for it, so it smudges them. PNG's lossless compression is great at long runs of identical pixels.\n\nThe screenshot step makes the same point with text: text in a JPG screenshot gets fuzzy halos.`,
});

// 25. Which format? decision flow
{
  const s = panelSlide("Which Format Should I Use?",
    "Walk through the decision with two or three examples from the class: a menu photo (JPG/WebP), the restaurant's logo (SVG), a screenshot for a how-to page (PNG), a photo of a product with the background removed (PNG or WebP, because it needs transparency).\n\nAnimated GIF: mention that a short looping MP4 or WebM video in a <video> tag is usually 5–10× smaller than the same GIF. We'll see video later.\n\nWebP vs JPG: both are fine. WebP is usually 25–35% smaller and works in every current browser. Keep a JPG if the image also needs to open in older software (email, Word, etc.).");
  const qs = [
    ["Is it a logo, icon or simple drawing?", "SVG", "(PNG if you only have a raster version)", C.blue],
    ["Is it a photo?", "JPG or WebP", "quality 60–80 · AVIF for even smaller", C.orange],
    ["Does it need a see-through background?", "PNG or WebP", "never JPG", C.teal],
    ["Screenshot, text or flat colours?", "PNG", "lossless keeps edges sharp", C.purple],
    ["Is it animated?", "Video (MP4)", "or animated WebP · GIF is the old way", C.mid],
  ];
  qs.forEach(([q, a, d, col], i) => {
    const y = 1.95 + i * 0.97;
    card(s, { x: CX, y, w: 4.1, h: 0.8 });
    txt(s, q, { x: CX + 0.25, y, w: 3.7, h: 0.8, fontSize: 14.5, bold: true, color: C.black, valign: "middle" });
    txt(s, "→", { x: CX + 4.1, y, w: 0.45, h: 0.8, fontSize: 22, bold: true, color: col, align: "center", valign: "middle" });
    card(s, { x: CX + 4.55, y, w: CW - 4.55, h: 0.8, fill: col, border: col });
    txt(s, [{ text: a, options: { bold: true, fontSize: 16, color: C.white, breakLine: true } }, { text: d, options: { fontSize: 11.5, color: C.white } }],
      { x: CX + 4.75, y, w: CW - 4.95, h: 0.8, valign: "middle" });
  });
  txt(s, "Then: resize to the size it's shown (× 2), and check the file size.", { x: CX, y: 6.85, w: CW, h: 0.4, fontSize: 14.5, bold: true, color: C.teal });
}

// 26. Divider: images in HTML
dividerSlide("Images in HTML", "width and height · lazy loading · srcset · picture",
  "Now that students know how to make a good image file, a few HTML attributes help the browser load it well.");

// 27. Example 04
codeSlide({
  kicker: "EXAMPLE 04", title: "width, height and loading=\"lazy\"", file: "04_width_height_lazy_example.html", lines: bodyOf("04_width_height_lazy_example.html"), size: 14,
  bullets: [
    [["width", "c"], " and ", ["height", "c"], " tell the browser the image's shape ", ["before", "b"], " it downloads, so it saves the space and the text below doesn't jump around."],
    ["They don't resize the file. A 4000 px photo with ", ["width=\"800\"", "c"], " still downloads all 4000 px. Resize the file itself."],
    [["loading=\"lazy\"", "c"], " waits until the user scrolls near the image before downloading it."],
    ["Don't lazy-load the first big image at the top. The user needs that one right away."],
  ],
  notes: "Demo the layout jump: in DevTools → Network, set throttling to 'Slow 4G', then reload a page with and without width/height. Without them, the text appears first and then gets pushed down when the image arrives. Google calls this Cumulative Layout Shift (CLS) and it's one of its page-speed scores.\n\nLazy loading demo: DevTools → Network → filter 'Img', reload. Only beach_800.jpg loads. Scroll down: beach_400.jpg appears in the list when you get close.\n\nThe inline style on the tall paragraph is just to make room for scrolling. We'll do this properly with CSS next class.",
});

// 28. Example 05
codeSlide({
  kicker: "EXAMPLE 05", title: "srcset and picture", file: "05_srcset_picture_example.html", lines: bodyOf("05_srcset_picture_example.html"), size: 14,
  bullets: [
    [["srcset", "c"], ": offer the ", ["same image at several widths", "b"], ". The browser picks the smallest one that will still look sharp on that screen."],
    ["The ", ["400w", "c"], " / ", ["800w", "c"], " tell the browser how wide each file really is."],
    [["<picture>", "c"], ": offer ", ["different formats", "b"], ". The browser uses the first one it supports; older browsers use the ", ["<img>", "c"], "."],
    ["Alt goes on the ", ["<img>", "c"], ", once. It covers every version."],
  ],
  notes: "This is a preview; students don't need to memorize the syntax. The idea is what matters: you can give the browser options and let it choose.\n\nDemo srcset: open DevTools, turn on the device toolbar (Ctrl+Shift+M), pick a phone, and reload with the Network tab open. You'll often see beach_400.jpg or beach_800.jpg load instead of the original. (Browsers also consider screen density and what's cached, so results vary. Use a private window to test.)\n\nThe sizes attribute is the advanced part of srcset; we skip it today. Without sizes, the browser assumes the image is the full width of the screen.",
});

// 29. Lab 6: put it on a page
labSlide({
  title: "Lab 6: Put It on a Page (15 min)",
  steps: [
    ["Open image_lab.html in VS Code", "Your name in the <title> and <h1>. Finish the Lab 1 list and the table."],
    ["Fix the alt text", ["Replace each ", ["TODO", "c"], " with a real description of your photo."]],
    ["Open it with Live Server", "Right-click → Open with Live Server. All four images are shown 600 px wide."],
    ["DevTools → Network → Img → reload", ["Ctrl + Shift + I. Compare the ", ["Size", "b"], " column for each image. Try throttling to ", ["Slow 4G", "b"], "."]],
    ["Fill in \"What I Noticed\" and submit", "Screenshot the page (and the Network tab), then upload to Brightspace for participation."],
  ],
  sideTitle: "Look for",
  side: [
    "All four look the same size on screen. Are they the same download size?",
    "On Slow 4G, which image appears last?",
    "Which one would you actually use?",
    ["", "h"],
    ["Bonus", "h"],
    "Add loading=\"lazy\" to the last two images, and width/height to all of them. Reload and watch the Network tab.",
  ],
  notes: "This brings it all together: the same picture, shown at the same size, with very different download costs. On Slow 4G the original phone photo can take many seconds to appear while the 600 px version is almost instant.\n\nThe table in image_lab.html uses <caption>, <thead> and scope from the Semantic HTML class. Point that out.\n\nCollect: screenshot of the page + Network tab, uploaded to Brightspace (participation). Students keep their my_images folder: next class we'll style these images with CSS.",
});

// 30. Discussion
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  s.addShape(pres.shapes.OVAL, { x: 7.6, y: 0.9, w: 5.4, h: 5.4, fill: { type: "none" }, line: { color: C.ring, width: 28 } });
  txt(s, "How did you make out?", { x: 0.8, y: 0.8, w: 7, h: 0.9, fontSize: 40, color: C.black });
  richBullets(s, [
    ["What was the ", ["biggest saving", "b"], ": format, quality, or size?"],
    ["What was the lowest JPG quality that still looked OK on ", ["your", "i"], " photo?"],
    ["Why might the same quality number look fine on one photo and bad on another?"],
    ["When would you choose PNG over JPG? SVG over PNG?"],
  ], { x: 0.8, y: 2.0, w: 6.5, h: 4.5, size: 20, space: 16 });
  txt(s, "Smallest file that still looks good", { x: 8.4, y: 2.9, w: 3.8, h: 1.4, fontSize: 22, bold: true, color: C.teal, align: "center", valign: "middle" });
  s.addNotes("Usual answers: size (resizing) is by far the biggest saving, often 10–40×. Format is next (PNG → JPG for a photo is huge; JPG → WebP is a nice extra 25–35%). Quality 80 → 60 is a smaller extra saving.\n\nSame quality, different results: busy photos hide compression; smooth areas (sky, skin, plain walls) show it.\n\nPNG over JPG: transparency, screenshots, text, flat colour. SVG over PNG: logos and icons, anything that needs to be sharp at many sizes.");
}

// 31. Checklist
{
  const s = panelSlide("Web Image Checklist",
    "This is the checklist to use on every image from now on, including the assignment. It's also in the README.");
  const checks = [
    [["Keep the original. Always export ", ["copies", "b"], "."]],
    [["Right ", ["format", "b"], ": photo → JPG / WebP · transparency → PNG / WebP · logo → SVG."]],
    [["Right ", ["size", "b"], ": the width it's shown at × 2, no bigger."]],
    [["Right ", ["quality", "b"], ": 60–80 for photos. Zoom in and check."]],
    [["Check the ", ["file size", "b"], ": content photos under ~150 KB, banners under ~300 KB."]],
    [["In HTML: ", ["alt", "c"], " text, ", ["width", "c"], " and ", ["height", "c"], ", and ", ["loading=\"lazy\"", "c"], " below the fold."]],
    [["Use sensible lower-case file names: ", ["harbour-view.jpg", "c"], ", not ", ["IMG_4032.JPG", "c"], "."]],
  ];
  checks.forEach((segs, i) => {
    const y = 1.95 + i * 0.7;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: CX, y: y + 0.05, w: 0.32, h: 0.32, rectRadius: 0.05, fill: { color: C.white }, line: { color: C.teal, width: 1.5 } });
    txt(s, "✓", { x: CX, y: y + 0.05, w: 0.32, h: 0.32, fontSize: 14, bold: true, color: C.teal, align: "center", valign: "middle" });
    richBullets(s, segs, { x: CX + 0.5, y, w: CW - 0.5, h: 0.6, size: 15, bullet: false, color: C.dark });
  });
}

// 32. Recap
{
  const s = panelSlide("Class Recap and Next Steps",
    "Next class is CSS. Ask students to keep their my_images folder and image_lab.html: we'll use their optimized photo when we style images with CSS.");
  labelBullets(s, [
    { label: "Recap:", text: [
      "Raster images are grids of pixels. Dimensions and colour depth set the raw size.",
      "Lossless keeps every pixel (PNG, GIF). Lossy throws away detail you won't miss (JPG, WebP, AVIF).",
      "Photos → JPG/WebP. Transparency → PNG/WebP. Logos → SVG.",
      "The biggest saving is usually resizing to the size the image is shown at.",
    ] },
    { label: "Next class: Intro to CSS", text: [
      "Selectors, colour, text, the box model, and styling images.",
      "Bring your my_images folder. We'll style your photo.",
    ] },
  ], { size: 20, sub: 16 });
}

pres.writeFile({ fileName: OUT }).then((f) => console.log("wrote " + f));
