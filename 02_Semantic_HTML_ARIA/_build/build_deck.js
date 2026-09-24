// Builds Class5_Semantic_HTML_ARIA.pptx in the Class 4 "Usability and Design Principles" style.
// To rebuild: cd 02_Semantic_HTML_ARIA/_build, run "npm install" once, then "npm run build".
// Code slides read the example .html files directly, so edit those files and re-run to update the slides.
const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const REPO = path.join(__dirname, ".."); // the 02_Semantic_HTML_ARIA folder
const SCR = __dirname;
const OUT = path.join(REPO, "Class5_Semantic_HTML_ARIA.pptx");

const FONT = "Aptos";
const MONO = "Consolas";
const C = {
  dark: "262626", light: "F2F2F2", black: "000000", white: "FFFFFF",
  orange: "E97132", purple: "A02B93", blue: "0F9ED5", navy: "12304A",
  teal: "1B5E82", grey: "595959", mid: "7F7F7F", actBg: "E1E8EB",
  ring: "B7CFC6",
};
const POUR = {
  P: { name: "Perceivable", color: "C2571F" },
  O: { name: "Operable", color: "A02B93" },
  U: { name: "Understandable", color: "0B7A9E" },
  R: { name: "Robust", color: "1B4F72" },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "Class 5: Accessibility in Code";
pres.author = "Jamie Symonds";

const W = 13.333, H = 7.5;
const PANEL_W = 4.1;
const CX = 5.2, CW = 7.55; // content column on panel slides

// ---------- helpers ----------
function txt(slide, text, opts) {
  slide.addText(text, Object.assign({ isTextBox: true, fontFace: FONT, margin: 0 }, opts));
}

function panelSlide(title, notes) {
  const s = pres.addSlide();
  s.background = { color: C.light };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: PANEL_W, h: H, fill: { color: C.dark }, line: { color: C.dark } });
  txt(s, "Usability and Design Principles", { x: 0.95, y: 0.6, w: 2.8, h: 1.4, fontSize: 22, color: C.light, valign: "top" });
  s.addShape(pres.shapes.RECTANGLE, { x: CX, y: 0.72, w: 0.5, h: 0.05, fill: { color: C.black }, line: { color: C.black } });
  txt(s, title, { x: CX, y: 0.9, w: CW, h: 0.95, fontSize: 28, bold: true, color: C.black, valign: "top" });
  if (notes) s.addNotes(notes);
  return s;
}

// Bold-label bullets with sub-bullets, Class 4 style
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

function simpleBullets(s, items, { x = CX, y = 1.95, w = CW, h = 5, size = 17, color = C.black, space = 10 } = {}) {
  const runs = items.map((t, i) => {
    const o = { bullet: true, fontSize: size, paraSpaceBefore: i ? space : 0, breakLine: i < items.length - 1 };
    if (Array.isArray(t)) return t; // pre-built run array (unused)
    return { text: t, options: o };
  });
  txt(s, runs, { x, y, w, h, color, valign: "top" });
}

// Rich bullet: array of [text, bold?] segments per bullet
function richBullets(s, bullets, { x, y, w, h, size = 15, color = C.dark, space = 9, fontFace = FONT }) {
  const runs = [];
  bullets.forEach((segs, bi) => {
    segs.forEach((seg, si) => {
      const [t, style] = typeof seg === "string" ? [seg, ""] : seg;
      const o = { fontSize: size, fontFace: style.includes("c") ? MONO : fontFace, bold: style.includes("b"), italic: style.includes("i") };
      if (style.includes("c")) o.color = "8A2B0B";
      if (si === 0) runs.push({ text: "​", options: { fontSize: size, fontFace, color, bullet: true, paraSpaceBefore: bi ? space : 0 } });
      if (si === segs.length - 1 && bi < bullets.length - 1) o.breakLine = true;
      runs.push({ text: t, options: o });
    });
  });
  txt(s, runs, { x, y, w, h, color, valign: "top" });
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
        if (ch === ">" ) { push(">", HL.punct); i++; inTag = false; }
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

function pourTags(s, tags, x, y) {
  let cx = x;
  tags.forEach((t) => {
    const p = POUR[t];
    const w = 0.28 + p.name.length * 0.095;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cx, y, w, h: 0.34, rectRadius: 0.17, fill: { color: p.color }, line: { color: p.color } });
    txt(s, p.name.toUpperCase(), { x: cx, y, w, h: 0.34, fontSize: 10.5, bold: true, color: C.white, align: "center", valign: "middle", charSpacing: 1 });
    cx += w + 0.12;
  });
}

// Code slide: gradient left + code, title & description right (Class 4 "Semantic HTML" slide)
const CODE_PANEL = 7.0;
function codeSlide({ kicker, title, tags, bullets, file, lines, notes, size = 14.5 }) {
  const s = pres.addSlide();
  s.background = { color: C.white };
  s.addImage({ path: path.join(SCR, "grad_left.png"), x: 0, y: 0, w: CODE_PANEL, h: H });
  codeBox(s, lines, { x: 0.3, y: 0.4, w: CODE_PANEL - 0.6, h: H - 0.8, maxSize: 15, fitHeight: true });
  const RX = CODE_PANEL + 0.45, RW = W - RX - 0.45;
  txt(s, kicker, { x: RX, y: 0.45, w: RW, h: 0.35, fontSize: 14, color: C.grey, bold: true, charSpacing: 1 });
  const long = title.length > 26;
  txt(s, title, { x: RX, y: 0.8, w: RW, h: long ? 1.05 : 0.7, fontSize: long ? 26 : 32, color: C.black, valign: "top" });
  const ty = long ? 1.98 : 1.65;
  if (tags) pourTags(s, tags, RX, ty);
  richBullets(s, bullets, { x: RX, y: ty + 0.55, w: RW, h: 6.85 - (ty + 0.55), size });
  if (file) txt(s, file, { x: RX, y: 6.95, w: RW, h: 0.3, fontSize: 11, color: C.mid, fontFace: MONO });
  if (notes) s.addNotes(notes);
  return s;
}

function dividerSlide(title, sub, notes) {
  const s = pres.addSlide();
  s.background = { path: path.join(SCR, "grad_full.png") };
  txt(s, title, { x: 3.2, y: 0.7, w: 9.2, h: 4.6, fontSize: 60, color: C.white, align: "right", valign: "middle" });
  if (sub) txt(s, sub, { x: 1.55, y: 6.55, w: 9, h: 0.45, fontSize: 18, color: C.white });
  s.addShape(pres.shapes.LINE, { x: 1.0, y: 4.2, w: 0, h: 3.3, line: { color: C.white, width: 2 } });
  // Class 4 sparkle decorations
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

function card(s, { x, y, w, h, fill = C.white }) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: fill }, line: { color: "D0D7DB", width: 0.75 },
    shadow: { type: "outer", blur: 6, offset: 2, angle: 90, color: "000000", opacity: 0.12 } });
}

function numCircle(s, n, x, y, color, d = 0.5) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color }, line: { color } });
  txt(s, String(n), { x, y, w: d, h: d, fontSize: d > 0.45 ? 18 : 14, bold: true, color: C.white, align: "center", valign: "middle" });
}

function table(s, rows, { x = CX, y = 1.95, w = CW, colW, size = 13, codeCols = [], rowH = 0.36 }) {
  const data = rows.map((r, ri) => r.map((cell, ci) => {
    const o = { fontFace: codeCols.includes(ci) && ri > 0 ? MONO : FONT, fontSize: ri === 0 ? size : (codeCols.includes(ci) ? size - 1 : size),
      color: ri === 0 ? C.white : C.black, bold: ri === 0, valign: "middle",
      fill: { color: ri === 0 ? C.dark : (ri % 2 ? C.white : "E4E4E4") }, margin: [3, 6, 3, 6] };
    return { text: cell, options: o };
  }));
  s.addTable(data, { x, y, w, colW, rowH, border: { type: "solid", pt: 0.5, color: "C8C8C8" } });
}

// =====================================================================
// 1. Title (Class 4 orange-circle title slide)
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  txt(s, "WEBD 1000", { x: 0.3, y: 0.3, w: 3, h: 0.4, fontSize: 16, color: C.black });
  s.addShape(pres.shapes.OVAL, { x: 0.55, y: 1.1, w: 5.6, h: 5.6, fill: { color: C.orange }, line: { color: C.orange } });
  txt(s, "Usability and Design Principles", { x: 1.45, y: 2.4, w: 4.2, h: 2.9, fontSize: 46, color: C.white, valign: "middle" });
  s.addShape(pres.shapes.OVAL, { x: 1.0, y: 5.75, w: 0.62, h: 0.62, fill: { color: C.purple }, line: { color: C.purple } });
  s.addShape(pres.shapes.ARC, { x: 9.6, y: 0.75, w: 3.3, h: 3.3, angleRange: [200, 20], line: { color: C.blue, width: 7, dashType: "dash" } });
  txt(s, [
    { text: "Class 5:", options: { bullet: true, fontSize: 30, breakLine: true } },
    { text: "Accessibility in code: semantic HTML5 and ARIA.", options: { bullet: { code: "25CB" }, indentLevel: 1, fontSize: 26, paraSpaceBefore: 6 } },
  ], { x: 6.85, y: 1.9, w: 5.6, h: 2.2, color: C.black, valign: "top" });
  s.addNotes("Welcome back. Last class (Class 4) was the WHY of accessibility: who our users are, WCAG 2.1 AA, POUR, and testing tools like Lighthouse and Read Aloud.\n\nToday is the HOW: what accessible code actually looks like. We'll look at 10 code examples (semantic HTML5 first, then ARIA), and finish with an activity where we test a broken page, fix it, and test it again.");
}

// 2. Attendance
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  txt(s, "Attendance", { x: 1.1, y: 0.75, w: 8, h: 0.9, fontSize: 44, color: C.black });
  s.addNotes("Take attendance.");
}

// 3. Overview
{
  const s = panelSlide("Class 5: Class Overview",
    "Agenda for today. Roughly: 10 min recap and setup, 30 min semantic HTML examples, 35 min ARIA, 35 min activity, 5 min wrap-up.\n\nEmphasize that everything today connects back to POUR from Class 4. Each code example slide has a coloured tag showing which WCAG principle it supports.");
  txt(s, [
    { text: "Title: ", options: { bold: true, bullet: true, fontSize: 22 } },
    { text: "Accessibility in Code", options: { fontSize: 22, breakLine: true } },
    { text: " ", options: { fontSize: 10, breakLine: true } },
    { text: "Agenda:", options: { bold: true, bullet: true, fontSize: 22, breakLine: true } },
    { text: "Recap: WCAG and POUR", options: { bullet: true, indentLevel: 1, fontSize: 18, paraSpaceBefore: 4, breakLine: true } },
    { text: "Part 1: Semantic HTML5 (Examples 01–06)", options: { bullet: true, indentLevel: 1, fontSize: 18, paraSpaceBefore: 4, breakLine: true } },
    { text: "Part 2: ARIA: what it is, the rules, and Examples 07–10", options: { bullet: true, indentLevel: 1, fontSize: 18, paraSpaceBefore: 4, breakLine: true } },
    { text: "Class Activity: test → fix → re-test", options: { bullet: true, indentLevel: 1, fontSize: 18, paraSpaceBefore: 4, breakLine: true } },
    { text: "Accessibility checklist and next steps", options: { bullet: true, indentLevel: 1, fontSize: 18, paraSpaceBefore: 4 } },
  ], { x: CX, y: 1.95, w: CW, h: 4.8, color: C.black, valign: "top" });
}

// 4. Get the code
{
  const s = panelSlide("Before We Start: Get the Code",
    "Everyone should already have the repo cloned from the HTML Examples class. Just pull.\n\nThe new folder is 02_Semantic_HTML_ARIA. It has 10 numbered examples, activity.html (the broken page), and activity_answer.html. Ask students NOT to open activity_answer.html until the end.\n\nLive Server: Lighthouse sometimes refuses to audit a page opened as a file:// URL, so serve it with Live Server (http://127.0.0.1:5500/...).");
  txt(s, [{ text: "1.  Pull this week's examples", options: { bold: true, fontSize: 18 } }], { x: CX, y: 1.95, w: CW, h: 0.4 });
  codeBox(s, ["git pull"], { x: CX, y: 2.45, w: 5.2, h: 0.75, maxSize: 16 });
  txt(s, "A new folder appears: 02_Semantic_HTML_ARIA (examples 01–10, activity.html, activity_answer.html).", { x: CX, y: 3.35, w: CW, h: 0.6, fontSize: 15, color: C.dark });
  txt(s, [{ text: "2.  Open pages with Live Server", options: { bold: true, fontSize: 18 } }], { x: CX, y: 4.2, w: CW, h: 0.4 });
  txt(s, "Right-click any .html file in VS Code → Open with Live Server. Lighthouse works best on a served page (http://127.0.0.1:5500/…) rather than one opened straight from a file.", { x: CX, y: 4.65, w: CW, h: 0.9, fontSize: 15, color: C.dark });
  txt(s, [{ text: "3.  Keep DevTools handy", options: { bold: true, fontSize: 18 } }], { x: CX, y: 5.75, w: CW, h: 0.4 });
  txt(s, "Ctrl + Shift + I. We'll use the Lighthouse tab and the Accessibility pane today.", { x: CX, y: 6.2, w: CW, h: 0.5, fontSize: 15, color: C.dark });
}

// 5. Recap: POUR cards
{
  const s = panelSlide("Class 4 Recap: WCAG 2.1 AA and POUR",
    "Quick recap. In Canada the Government of Canada recommends WCAG 2.1 Level AA (a11y.canada.ca). WCAG is organized around four principles, POUR.\n\nAsk the class: which of these did Lighthouse check last class? (Mostly Perceivable: alt text and contrast. Some Robust: names and labels.) Which can't a tool check well? (Understandable: clear language. Operable: real keyboard use.)\n\nToday's examples each carry a coloured tag for the principle they support.");
  txt(s, "The standard to follow in Canada is WCAG 2.1 Level AA. It is organized around four principles:", { x: CX, y: 1.85, w: CW, h: 0.6, fontSize: 15, color: C.dark });
  const cards = [
    ["P", "Can users perceive it?", "Text alternatives (alt), captions, colour contrast, resizable text"],
    ["O", "Can users operate it?", "Keyboard access, visible focus, headings and landmarks to navigate"],
    ["U", "Can users understand it?", "Clear language, consistent navigation, helpful error messages"],
    ["R", "Does it work with assistive tech?", "Semantic HTML and ARIA, so screen readers can interpret it"],
  ];
  const cw = 3.65, ch = 2.1;
  cards.forEach(([k, q, d], i) => {
    const x = CX + (i % 2) * (cw + 0.25), y = 2.6 + Math.floor(i / 2) * (ch + 0.25);
    card(s, { x, y, w: cw, h: ch });
    numCircle(s, k, x + 0.2, y + 0.22, POUR[k].color, 0.55);
    txt(s, POUR[k].name, { x: x + 0.9, y: y + 0.22, w: cw - 1.0, h: 0.55, fontSize: 18, bold: true, color: C.black, valign: "middle" });
    txt(s, q, { x: x + 0.25, y: y + 0.9, w: cw - 0.45, h: 0.35, fontSize: 14, italic: true, color: C.dark });
    txt(s, d, { x: x + 0.25, y: y + 1.25, w: cw - 0.45, h: 0.75, fontSize: 13, color: C.grey, valign: "top" });
  });
}

// 6. Bridge: Class 4 semantic HTML code, corrected
codeSlide({
  kicker: "FROM CLASS 4 → TODAY",
  title: "Semantic HTML",
  tags: ["R"],
  file: "The code from the Class 4 slide, corrected",
  lines: [
    "<!DOCTYPE html>", '<html lang="en">', "<head>", '   <meta charset="UTF-8">', "   <title>My Page</title>", "</head>", "<body>", "",
    "   <header>", "      <h1>My Website</h1>", "   </header>", "",
    "   <nav>", '      <a href="#">Home</a>', '      <a href="#">About</a>', "   </nav>", "",
    "   <main>", "      <section>", "         <h2>Welcome</h2>", "         <p>This is the main content.</p>", "      </section>", "   </main>", "",
    "   <footer>", "      <p>&copy; 2026 Jamie</p>", "   </footer>", "", "</body>", "</html>",
  ],
  bullets: [
    ["Semantic HTML uses elements such as ", ["<header>", "c"], ", ", ["<nav>", "c"], ", ", ["<main>", "c"], ", ", ["<section>", "c"], " and ", ["<footer>", "c"], " to describe the ", ["purpose", "b"], " of each part of a page."],
    ["Instead of generic ", ["<div>", "c"], " boxes, the elements themselves say what the content is. That's easier for developers to read, for screen readers to navigate, and for search engines to understand."],
    ["Last class this was the ", ["why", "b"], ". Today we write it, one element at a time, and then add ARIA where HTML alone isn't enough."],
    [["Corrected: ", "b"], "the nav links are real ", ["<a href>", "c"], " tags, and ", ["lang=\"en\"", "c"], " is on ", ["<html>", "c"], " so screen readers use the right voice."],
  ],
  notes: "This is the Semantic HTML slide from Class 4, with the code corrected. In the Class 4 screenshot the nav links had lost their <a href=\"#\"> opening tags. Here they're proper links. I also added lang=\"en\" (screen readers use it to pick the pronunciation voice, and Lighthouse checks for it) and a charset meta tag.\n\nWalk through the structure: header holds the site title, nav holds the site links, main holds this page's unique content, section groups a themed part of main, footer holds the copyright. Today we go element by element.",
});

// 7. Divider: Part 1
dividerSlide("Part 1: Semantic HTML5", "Examples 01–06", "Part 1: semantic HTML. Six examples. Have students open each file with Live Server as we go.");

// 8. Why semantic HTML matters
{
  const s = panelSlide("Who Reads Your HTML?",
    "The key idea: your HTML has many readers, not just the browser's paint engine.\n\n- Screen readers build a list of headings, landmarks, links and form fields from your HTML.\n- Reader View (F9 in Edge, from Class 4) looks for <main>, <article> and headings to decide what's real content.\n- Search engines weigh headings and structure (the SEO benefit from Class 4).\n- Keyboard and voice-control users rely on real buttons and links being focusable.\n- Developers: <nav> is self-documenting; <div class=\"menu\"> isn't.\n\nThe browser draws a <div> and a <nav> the same way. Everyone else treats them differently.");
  labelBullets(s, [
    { label: "Screen readers", text: "Announce landmarks, headings, lists, tables and buttons, and let users jump between them." },
    { label: "Reader View and Read Aloud (Class 4)", text: "Use <main>, <article> and headings to find the real content on the page." },
    { label: "Search engines", text: "Use headings and structure to understand what a page is about (the SEO benefit)." },
    { label: "Keyboard and voice-control users", text: "Real <button> and <a> elements are focusable and clickable by voice automatically." },
    { label: "Other developers (and future you)", text: "<nav> explains itself. <div class=\"menu\"> doesn't." },
  ], { size: 17, sub: 15 });
}

// 9. How screen reader users navigate
{
  const s = panelSlide("How Screen Reader Users Navigate",
    "Most screen reader users don't listen to a page top to bottom, any more than sighted users read every word. They skim using shortcut keys.\n\nIn NVDA (free, Windows): H jumps to the next heading, 1–6 jump to headings of that level, D jumps to the next landmark, K to the next link, and Insert+F7 opens a list of all headings, links or landmarks.\n\nWebAIM's screen reader user survey has consistently found that navigating by headings is the most common way users find information on a long page.\n\nConnect back to Class 4: when Read Aloud struggled on some pages, it was usually because the structure wasn't in the HTML.");
  labelBullets(s, [
    { label: "They skim, just like sighted users", text: "Very few people listen to a whole page from top to bottom." },
    { label: "By headings", text: ["Jump heading to heading (NVDA: H key, or 1–6 for a level).", "WebAIM's screen reader survey: headings are the #1 way users find information on a long page."] },
    { label: "By landmarks", text: "Jump straight to navigation, main or footer (NVDA: D key)." },
    { label: "By lists of links", text: "Pull up every link on the page, so link text like \"click here\" ×10 is useless." },
    { label: "Takeaway", text: "No headings + no landmarks = nothing to skim. That's what made Read Aloud struggle in Class 4." },
  ], { size: 17, sub: 15 });
}

// 10–15 Examples 01–06
codeSlide({
  kicker: "EXAMPLE 01", title: "Div Soup: the \"Before\"", tags: ["R"], file: "01_div_soup_example.html", lines: bodyOf("01_div_soup_example.html"),
  bullets: [
    [["<div>", "c"], " is a generic box with ", ["no meaning", "b"], ". To a screen reader, \"Burridge Campus\" is just more text, not a heading."],
    ["There are no landmarks to jump to and no headings to skim, so a screen reader user has to listen to everything, in order."],
    ["Once CSS is added it looks perfectly fine on screen. That's why this problem is so common, and so easy to miss."],
    [["Try it: ", "b"], "DevTools → Elements → Accessibility pane. Inspect the title div: its role is just ", ["generic", "c"], "."],
  ],
  notes: "Open 01_div_soup_example.html with Live Server. It looks like a normal (unstyled) page.\n\nNow open DevTools (Ctrl+Shift+I), select the 'Burridge Campus' div in Elements, and open the Accessibility pane (in Edge/Chrome it's a tab next to Styles; you may need to click >>). The role says 'generic'. There's no heading, no navigation, no main.\n\nOptional: turn on Windows Narrator (Win+Ctrl+Enter) and press H to jump to headings. Narrator says there are no headings.\n\n'Div soup' is a real industry term for pages built only from divs.",
});
codeSlide({
  kicker: "EXAMPLE 02", title: "Landmarks and a Skip Link", tags: ["R", "O"], file: "02_landmarks_example.html", lines: bodyOf("02_landmarks_example.html"),
  bullets: [
    [["<header>", "c"], ", ", ["<nav>", "c"], ", ", ["<main>", "c"], " and ", ["<footer>", "c"], " create ", ["landmarks", "b"], ": named regions a screen reader lists and can jump between."],
    ["The site title is now a real ", ["<h1>", "c"], ", so it shows up in the headings list too. Links are in a list, so a screen reader says \"list, 2 items\"."],
    ["Use exactly ", ["one <main>", "b"], " per page. It holds the content that's unique to this page."],
    ["The ", ["skip link", "b"], " lets keyboard users jump past the menu instead of tabbing through every link, on every page."],
  ],
  notes: "Same content as Example 01, rewritten with semantic elements.\n\nLandmark names a screen reader announces: <header> = banner, <nav> = navigation, <main> = main, <footer> = content info. We'll see the full mapping on the ARIA slides.\n\nSkip link demo: load the page, click in the address bar, press Tab once. The first thing focused is 'Skip to main content'. Press Enter and focus jumps to main. Real sites usually hide the skip link with CSS until it gets focus. We'll do that when we get to CSS.\n\nThis connects to Class 4's Operable slide: 'Navigable Content: use logical and clear navigation structures, such as headings and landmarks.'",
});
codeSlide({
  kicker: "EXAMPLE 03", title: "article, section and aside", tags: ["R"], file: "03_article_section_aside_example.html", lines: bodyOf("03_article_section_aside_example.html"),
  bullets: [
    [["<article>", "c"], ": complete ", ["on its own", "b"], ". You could move it to another page and it still makes sense (news item, blog post, product card)."],
    [["<section>", "c"], ": a themed ", ["part", "b"], " of something bigger. It should almost always start with a heading."],
    [["<aside>", "c"], ": ", ["related but not essential", "b"], ". A sidebar, a \"did you know?\" box, or related links."],
    ["Headings nest: ", ["h2", "c"], " for the article, ", ["h3", "c"], " for its sections. Never skip a level just to get smaller text. That's CSS's job."],
  ],
  notes: "The most common question: article vs section? Test: 'Would this make sense on its own if I shared just this part?' If yes, article. If it's a chapter or part of something bigger, section.\n\nA section without a heading is usually a sign you wanted a <div>. <div> is still fine for grouping purely for styling; it just adds no meaning.\n\nReader View (F9 in Edge) looks for <article> to decide what to show. Try Reader View on this page vs. Example 01.\n\n<aside> is announced as 'complementary' by screen readers.",
});
codeSlide({
  kicker: "EXAMPLE 04", title: "figure and figcaption", tags: ["P"], file: "04_figure_figcaption_example.html", lines: bodyOf("04_figure_figcaption_example.html"),
  bullets: [
    ["Builds on last class's ", ["07_picture_example", "c"], "."],
    [["alt", "c"], " is text that ", ["replaces", "b"], " the image for people who can't see it. It's read by screen readers and shown if the image fails to load."],
    [["<figcaption>", "c"], " is a ", ["visible", "b"], " caption everyone sees. ", ["<figure>", "c"], " ties the image and its caption together."],
    ["They do different jobs, so don't copy the caption into the alt. Alt describes ", ["what's in", "i"], " the picture; the caption adds ", ["context", "i"], "."],
  ],
  notes: "Class 4 Perceivable slide: 'Text Alternatives: provide text alternatives (e.g., ALT attributes) for non-text content such as images.'\n\n<figure> isn't only for images. It can also wrap a chart, diagram, code sample or quote that's referenced from the main text.\n\nDemo: rename beach.JPG temporarily (or change src to beachX.JPG) and reload. The browser shows the alt text in place of the image. That's exactly what a screen reader user 'sees'.",
});
codeSlide({
  kicker: "EXAMPLE 05", title: "An Accessible Table", tags: ["P", "R"], file: "05_accessible_table_example.html", lines: bodyOf("05_accessible_table_example.html"),
  bullets: [
    ["Builds on last class's ", ["06_table_example", "c"], "."],
    [["<caption>", "c"], " gives the table a name, announced before the table is read."],
    [["<th>", "c"], " marks header cells. ", ["scope=\"col\"", "c"], " or ", ["scope=\"row\"", "c"], " says which direction each header applies."],
    ["Result: a screen reader says \"", ["August, Temperature, +20 degrees", "b"], "\" instead of just \"+20\"."],
    [["border=\"1\"", "c"], " is gone. How a table ", ["looks", "i"], " belongs in CSS, not HTML."],
  ],
  notes: "Screen reader users move around a table cell by cell (NVDA: Ctrl+Alt+arrow keys). Without <th> and scope, moving to '+20' gives no context. With them, the screen reader reads the column and row headers as you move.\n\n<thead> and <tbody> group the rows. They're also handy for styling later.\n\nReminder from the first class: tables are for tabular data only, never for page layout.",
});
codeSlide({
  kicker: "EXAMPLE 06", title: "Alt Text: Good, Bad and Decorative", tags: ["P"], file: "06_alt_text_example.html", lines: bodyOf("06_alt_text_example.html"),
  bullets: [
    [["No alt at all: ", "b"], "many screen readers read the file name instead (\"beach dot J P G\"). Lighthouse flags this."],
    [["alt=\"image\"", "c"], " or ", ["\"photo\"", "c"], ": useless. The screen reader already says \"graphic\"."],
    [["Good alt", "b"], " says what the image ", ["communicates", "i"], ", in a short sentence. Ask yourself: \"If I described this over the phone, what would I say?\""],
    [["Decorative", "b"], " images (dividers, flourishes) get ", ["alt=\"\"", "c"], ": empty but present, so screen readers skip them."],
  ],
  notes: "Rules of thumb for alt text:\n- Don't start with 'image of' or 'picture of'. The screen reader already announces it's a graphic.\n- Keep it short: one sentence, usually under ~125 characters.\n- Context matters. The same beach photo on a travel site vs. a weather site might need different alt text.\n- If the image is a link or a button, the alt should describe the destination or action, not the picture.\n- Images of text: the alt should contain the same text.\n\nalt=\"\" (empty) is NOT the same as leaving alt out. Empty means 'decorative, skip me'. Missing means 'I forgot', and screen readers guess (often reading the file name).\n\nRun Lighthouse on this page: it flags exactly one image, the one with no alt.",
});

// 16. Divider: Part 2
dividerSlide("Part 2: ARIA", "Accessible Rich Internet Applications. Examples 07–10",
  "Part 2: ARIA. Class 4's Robust slide said 'Implement ARIA roles to enhance accessibility for assistive technologies.' That's true, with one big caveat, which is rule #1 of ARIA. The next few slides explain what ARIA is before we look at code.");

// 17. What is ARIA?
{
  const s = panelSlide("What is ARIA?",
    "WAI-ARIA 1.0 became a W3C Recommendation in 2014; the current version is ARIA 1.2. It's published by WAI, the same W3C group that publishes WCAG.\n\nThe most important sentence on this slide: ARIA changes what assistive technology is told. It changes nothing else. Putting role=\"button\" on a div doesn't make it focusable, doesn't make Enter work, and doesn't make it look like a button. It only makes the screen reader SAY 'button'. That's a promise the developer then has to keep with JavaScript.\n\nWhy it exists: before HTML5, there were no elements for tabs, menus, dialogs, live updates, etc. JavaScript apps built them out of divs, and ARIA was created so those widgets could describe themselves.");
  labelBullets(s, [
    { label: "WAI-ARIA", text: "Web Accessibility Initiative – Accessible Rich Internet Applications. A W3C standard from the same group that publishes WCAG." },
    { label: "A set of HTML attributes", text: ["role=\"…\" plus attributes that start with aria-, e.g. aria-label, aria-hidden.", "They describe elements to assistive technology: screen readers, voice control, braille displays."] },
    { label: "It only changes what is announced", text: "No visual change. No keyboard support. No click behaviour. It changes the label, not the thing." },
    { label: "Why it exists", text: "To describe things HTML has no element for: custom widgets (tabs, menus, pop-ups) and content that updates without a page reload." },
  ], { size: 17, sub: 15 });
}

// 18. The accessibility tree (diagram)
{
  const s = panelSlide("The Accessibility Tree",
    "The browser builds two trees from your HTML: the DOM (everything) and the accessibility tree (a simplified version with only what matters to assistive tech). Each node in the accessibility tree has a role, a name, a state, and sometimes a description.\n\nScreen readers never see your pixels or your CSS classes. They read this tree.\n\nDemo: in Edge/Chrome DevTools, Elements tab → Accessibility pane shows the computed Role and Name for the selected element. Edge and Chrome also have a 'full-page accessibility tree' toggle (the person icon in the Elements panel) that swaps the DOM view for the accessibility tree view.\n\nSemantic HTML fills this tree in for free. ARIA is how you edit it when HTML can't.");
  const bx = [
    { t: "Your HTML", d: "The DOM: every element, class and attribute", fill: C.white, tc: C.black },
    { t: "Accessibility tree", d: "Role · Name · State · Description", fill: C.dark, tc: C.white },
    { t: "Assistive tech", d: "Screen reader, voice control, braille display", fill: C.white, tc: C.black },
  ];
  const bw = 2.15, gap = 0.55, by = 2.0, bh = 1.55;
  bx.forEach((b, i) => {
    const x = CX + i * (bw + gap);
    card(s, { x, y: by, w: bw, h: bh, fill: b.fill });
    txt(s, b.t, { x: x + 0.15, y: by + 0.2, w: bw - 0.3, h: 0.45, fontSize: 17, bold: true, color: b.tc, align: "center" });
    txt(s, b.d, { x: x + 0.15, y: by + 0.7, w: bw - 0.3, h: 0.75, fontSize: 13, color: b.tc === C.white ? "E0E0E0" : C.grey, align: "center", valign: "top" });
    if (i < 2) s.addShape(pres.shapes.RIGHT_ARROW, { x: x + bw + 0.1, y: by + bh / 2 - 0.17, w: gap - 0.2, h: 0.34, fill: { color: C.orange }, line: { color: C.orange } });
  });
  // worked example
  txt(s, "Example", { x: CX, y: 3.9, w: 3, h: 0.35, fontSize: 15, bold: true, color: C.black });
  codeBox(s, ["<button>Save</button>"], { x: CX, y: 4.3, w: bw, h: 0.6, maxSize: 13 });
  card(s, { x: CX + bw + gap, y: 4.3, w: bw, h: 0.6, fill: C.dark });
  txt(s, "role: button · name: \"Save\"", { x: CX + bw + gap, y: 4.3, w: bw, h: 0.6, fontSize: 12.5, color: C.white, align: "center", valign: "middle" });
  card(s, { x: CX + 2 * (bw + gap), y: 4.3, w: bw, h: 0.6 });
  txt(s, "\"Save, button\"", { x: CX + 2 * (bw + gap), y: 4.3, w: bw, h: 0.6, fontSize: 14, italic: true, color: C.black, align: "center", valign: "middle" });
  simpleBullets(s, [
    "Screen readers read this tree, not your pixels or your CSS.",
    "Semantic HTML fills in the role and name for free. ARIA lets you edit the tree when HTML can't.",
    "See it yourself: DevTools → Elements → Accessibility pane.",
  ], { y: 5.25, h: 1.8, size: 15, space: 6 });
}

// 19. Roles, properties, states
{
  const s = panelSlide("ARIA Roles, Properties and States",
    "Every ARIA attribute is one of three kinds.\n\nRoles: what the element IS. Almost every HTML element already has a role, called its implicit role (next slides). You rarely need to set one.\n\nProperties: characteristics that generally don't change: its name (aria-label, aria-labelledby), its description (aria-describedby), whether it's required.\n\nStates: things that change as the user interacts: is the menu open (aria-expanded), is this the current page (aria-current), is this field invalid (aria-invalid), is this hidden from AT (aria-hidden). JavaScript usually updates states.\n\nBottom example: a menu toggle button. When JS opens the menu, it must change aria-expanded to \"true\", and the screen reader then says 'expanded'.");
  const cols = [
    ["Role", "What is it?", ["role=\"button\"", "role=\"navigation\"", "role=\"alert\""], "Most HTML elements already have one (their implicit role)."],
    ["Property", "What is it called or described as?", ["aria-label", "aria-labelledby", "aria-describedby"], "Usually set once and doesn't change."],
    ["State", "What's happening right now?", ["aria-expanded", "aria-current", "aria-invalid", "aria-hidden"], "Changes as the user interacts, usually updated by JavaScript."],
  ];
  const cw = 2.35, gap = 0.25, y = 1.95, h = 3.45;
  cols.forEach(([t, q, ex, d], i) => {
    const x = CX + i * (cw + gap);
    card(s, { x, y, w: cw, h });
    txt(s, t, { x: x + 0.2, y: y + 0.15, w: cw - 0.4, h: 0.45, fontSize: 20, bold: true, color: [C.orange, C.purple, C.teal][i] });
    txt(s, q, { x: x + 0.2, y: y + 0.6, w: cw - 0.4, h: 0.55, fontSize: 13, italic: true, color: C.dark, valign: "top" });
    txt(s, ex.map((e, j) => ({ text: e, options: { breakLine: j < ex.length - 1 } })), { x: x + 0.2, y: y + 1.2, w: cw - 0.4, h: 1.2, fontSize: 12, fontFace: MONO, color: "8A2B0B", valign: "top" });
    txt(s, d, { x: x + 0.2, y: y + 2.45, w: cw - 0.4, h: 0.9, fontSize: 12, color: C.grey, valign: "top" });
  });
  txt(s, "All three together:", { x: CX, y: 5.65, w: 4, h: 0.35, fontSize: 15, bold: true, color: C.black });
  codeBox(s, ['<button aria-expanded="false">Menu</button>'], { x: CX, y: 6.05, w: 4.6, h: 0.6, maxSize: 12 });
  txt(s, "→  \"Menu, button, collapsed\"", { x: CX + 4.75, y: 6.05, w: 2.8, h: 0.6, fontSize: 15, italic: true, color: C.black, valign: "middle" });
}

// 20. Five rules
{
  const s = panelSlide("The Five Rules of ARIA",
    "These come from the W3C document 'Using ARIA'.\n\n1. Native first: <button> not <div role=button>; <nav> not <div role=navigation>; <input type=checkbox> not a div with role=checkbox.\n2. Don't override meaning: <h2 role=\"button\"> destroys the heading. Put a <button> inside the <h2> instead.\n3. If you do build a custom widget with ARIA, you owe the keyboard support yourself: tabindex, Enter and Space handling, arrow keys for menus, etc. That's a lot of JavaScript, which is why rule 1 exists.\n4. aria-hidden on a focusable element creates a 'ghost': keyboard focus lands on something the screen reader says doesn't exist.\n5. Every button, link and form field needs a name, from its text, a <label>, alt, or aria-label.\n\n'No ARIA is better than bad ARIA': WebAIM's yearly 'Million' report (automated tests of the top 1,000,000 home pages) consistently finds that pages using ARIA have more detected errors on average than pages without it.");
  const rules = [
    ["Use native HTML first.", "If an HTML element already does the job, use it instead of adding ARIA."],
    ["Don't change native meaning.", "No role=\"button\" on an <h2>. Put a <button> inside the heading instead."],
    ["Interactive ARIA must work with a keyboard.", "A role=\"button\" must also be focusable and respond to Enter and Space."],
    ["Don't hide focusable elements.", "Never put aria-hidden=\"true\" on something a user can Tab to."],
    ["Every interactive element needs a name.", "Buttons, links and form fields must have a name a screen reader can say."],
  ];
  rules.forEach(([t, d], i) => {
    const y = 1.95 + i * 0.85;
    numCircle(s, i + 1, CX, y + 0.05, i === 0 ? C.orange : C.dark, 0.5);
    txt(s, t, { x: CX + 0.7, y, w: CW - 0.7, h: 0.35, fontSize: 17, bold: true, color: C.black, valign: "top" });
    txt(s, d, { x: CX + 0.7, y: y + 0.36, w: CW - 0.7, h: 0.45, fontSize: 14, color: C.dark, valign: "top" });
  });
  card(s, { x: CX, y: 6.3, w: CW, h: 0.7, fill: C.dark });
  txt(s, [{ text: "\"No ARIA is better than bad ARIA.\"", options: { italic: true, bold: true } }, { text: "  WebAIM finds pages using ARIA average more errors, not fewer.", options: {} }],
    { x: CX + 0.25, y: 6.3, w: CW - 0.5, h: 0.7, fontSize: 14, color: C.white, valign: "middle" });
}

// 21. Implicit roles
{
  const s = panelSlide("Semantic HTML = Built-in ARIA",
    "This is the answer to Class 4's 'ARIA Landmarks' bullet. With HTML5, you get the landmark roles automatically by using the right element. No role attribute needed.\n\nNotes:\n- <header> and <footer> only become banner / contentinfo when they're for the whole page (not inside an <article> or <section>).\n- <section> only becomes a 'region' landmark when it has an accessible name (aria-label or aria-labelledby). Otherwise it's just a group.\n- You'll still see <nav role=\"navigation\"> in older tutorials. That was a workaround for very old browsers (IE) and isn't needed today. It's harmless but redundant.\n\nTry it in DevTools: inspect <nav> in Example 02. The Accessibility pane shows role: navigation.");
  txt(s, "Every one of these elements comes with an ARIA role already built in, so you don't need to add role=\"…\" yourself.", { x: CX, y: 1.85, w: CW, h: 0.55, fontSize: 14, color: C.dark });
  table(s, [
    ["HTML element", "Implicit ARIA role", "Screen reader says…"],
    ["<header> (page level)", "banner", "\"banner\""],
    ["<nav>", "navigation", "\"navigation\""],
    ["<main>", "main", "\"main\""],
    ["<aside>", "complementary", "\"complementary\""],
    ["<footer> (page level)", "contentinfo", "\"content info\""],
    ["<section> with a name", "region", "\"Campus News, region\""],
    ["<article>", "article", "\"article\""],
    ["<button>", "button", "\"Save, button\""],
    ["<a href=\"…\">", "link", "\"Home, link\""],
    ["<h2>", "heading, level 2", "\"heading level 2\""],
    ["<img alt=\"\">", "none (ignored)", "(skipped)"],
  ], { y: 2.45, colW: [2.6, 2.25, 2.7], size: 12.5, codeCols: [0, 1], rowH: 0.33 });
  txt(s, [{ text: "So ", options: {} }, { text: "<nav role=\"navigation\">", options: { fontFace: MONO, color: "8A2B0B" } }, { text: " is redundant: the <nav> already has that role.", options: {} }],
    { x: CX, y: 6.6, w: CW, h: 0.4, fontSize: 14, color: C.dark });
}

// 22. Example 07
codeSlide({
  kicker: "EXAMPLE 07", title: "The First Rule of ARIA: Use a Real <button>", tags: ["O", "R"], file: "07_aria_first_rule_example.html", lines: bodyOf("07_aria_first_rule_example.html"),
  bullets: [
    [["1. Plain <div>: ", "b"], "works with a mouse only. Tab skips it, and a screen reader doesn't know it's clickable."],
    [["2. <div role=\"button\">: ", "b"], "the screen reader now ", ["says", "i"], " \"button\", but Tab still skips it and Enter does nothing. ARIA changed the label, not the behaviour."],
    [["3. <button>: ", "b"], "focusable, Enter and Space work, visible focus ring, announced as a button. All for free."],
    [["Try it: ", "b"], "open the file and press Tab. Only one of the three \"buttons\" ever gets focus."],
  ],
  notes: "This is THE most important demo of the day.\n\nOpen 07_aria_first_rule_example.html. Click each one with the mouse: all three show the alert. Now put the mouse down and press Tab. Only the real <button> receives focus (you'll see the focus ring). Press Enter or Space and it works.\n\nTo make version 2 actually work, you'd need tabindex=\"0\" to make it focusable, plus JavaScript that listens for Enter AND Space keydown events, plus CSS for a focus style. That's rule 3 of ARIA. Or you could just... use <button>. That's rule 1.\n\nConnects to Class 4's Operable slide: 'Keyboard Accessibility: ensure all functionality is accessible via keyboard' and 'Focus Indicators'.\n\nRule of thumb: if it goes somewhere, use <a href>. If it does something, use <button>.",
});

// 23. Accessible names
{
  const s = panelSlide("Accessible Names: Three Ways to Label",
    "Every interactive element needs an accessible name (rule 5). Normally it comes from the element's text, a <label>, or alt. When it can't, use one of these.\n\naria-label: a string you type in. Invisible on screen. Good for icon-only buttons and for telling two navs apart.\naria-labelledby: points at the id of visible text that already names the thing. Preferred when that text exists, because what you see and what you hear stay in sync.\naria-describedby: NOT a name. It's extra information read after the name and role, for hints, formats and error messages.\n\nPriority if several exist: aria-labelledby wins, then aria-label, then the native label (label element, alt, text content), then title as a last resort.\n\nWCAG 2.5.3 'Label in Name': if a button shows the text 'Search', its accessible name must contain 'Search'. Voice-control users say what they see ('click Search'). aria-label=\"Find stuff\" on a button that shows 'Search' breaks that.");
  table(s, [
    ["Attribute", "Use it when…", "Example"],
    ["aria-label=\"…\"", "There's no visible text to use: an icon-only button, or two navs to tell apart", "<button aria-label=\"Close menu\">×</button>"],
    ["aria-labelledby=\"id\"", "Visible text on the page already names it (e.g. a heading)", "<section aria-labelledby=\"news-heading\">"],
    ["aria-describedby=\"id\"", "You want extra info read after the name: a hint or an error", "<input aria-describedby=\"email-error\">"],
  ], { y: 1.95, colW: [2.35, 2.6, 2.6], size: 13, codeCols: [0, 2], rowH: [0.4, 0.8, 0.65, 0.65] });
  labelBullets(s, [
    { label: "Priority", text: "aria-labelledby → aria-label → native label (<label>, alt, the element's text)." },
    { label: "aria-label replaces the visible text for screen readers", text: "Keep the on-screen words in the name. Voice-control users say what they see (\"click Search\")." },
    { label: "For form fields, a real <label> beats any ARIA", text: "It's visible, clickable and works everywhere." },
  ], { y: 4.8, h: 2.3, size: 15, sub: 13.5 });
}

// 24–26 Examples 08–10
codeSlide({
  kicker: "EXAMPLE 08", title: "aria-label, aria-labelledby and aria-current", tags: ["R"], file: "08_aria_labels_example.html", lines: bodyOf("08_aria_labels_example.html"),
  bullets: [
    [["aria-label=\"Main\" / \"Footer\": ", "b"], "two ", ["<nav>", "c"], "s would both just be called \"navigation\". The labels tell them apart."],
    [["aria-current=\"page\": ", "b"], "announced as \"Home, link, current page\". It's the screen reader version of highlighting the active menu item."],
    [["aria-labelledby: ", "b"], "names the section using the heading that's already on screen, and turns it into a \"Campus News, region\" landmark."],
    [["aria-label on ×: ", "b"], "without it, a screen reader says \"times, button\". With it: \"Close menu, button\"."],
  ],
  notes: "Walk through each attribute.\n\n- Two navs: a screen reader's landmark list would show 'navigation, navigation'. With labels: 'Main navigation, Footer navigation'. Don't include the word 'navigation' in the label; the role already says it.\n- aria-current=\"page\" goes on the link to the page you're currently on. It's a state, and it's usually set by the server or template. Later, CSS can use it for styling too: a[aria-current=\"page\"] { font-weight: bold; }\n- aria-labelledby uses the id of the visible h2. If the heading text changes, the name updates automatically.\n- The × character is read literally ('times' or 'multiplication sign'). aria-label gives it a real name.\n\nDemo in DevTools: inspect the × button. The Accessibility pane's Name field shows 'Close menu'.",
});
codeSlide({
  kicker: "EXAMPLE 09", title: "aria-hidden and aria-live", tags: ["P", "R"], file: "09_aria_hidden_live_example.html", lines: bodyOf("09_aria_hidden_live_example.html"),
  bullets: [
    [["aria-hidden=\"true\"", "c"], " removes an element from the accessibility tree. Use it for decoration: \"black star, Favourite\" becomes just \"Favourite\"."],
    ["Never put it on anything ", ["focusable", "b"], " (rule 4). Keyboard users would land on something the screen reader says isn't there."],
    [["aria-live=\"polite\"", "c"], ": when the text inside changes, the screen reader announces it once it finishes what it's saying. Good for \"Saved!\" messages."],
    ["The live region must already be on the page (empty) ", ["before", "i"], " the text changes. ", ["\"assertive\"", "c"], " interrupts immediately, so save it for urgent errors."],
  ],
  notes: "aria-hidden: common uses are icon fonts, emoji and decorative SVGs next to text that already says the same thing. Remember: aria-hidden hides from assistive tech only. It's still visible on screen.\n\naria-live: without it, when JavaScript changes text on the page, a screen reader user has no idea anything happened. They pressed Save and heard... nothing. With aria-live=\"polite\", the new text is announced.\n\nDemo with Narrator (Win+Ctrl+Enter): Tab to Save, press Enter, and Narrator says 'Saved!'. Remove aria-live and try again: silence.\n\nThe one line of JavaScript just sets the paragraph's text. We'll learn JS later; for now it's 'this changes the text'.\n\nrole=\"status\" and role=\"alert\" are shortcuts that include live-region behaviour (polite and assertive respectively).",
});
codeSlide({
  kicker: "EXAMPLE 10", title: "An Accessible Form", tags: ["U"], file: "10_accessible_form_example.html", lines: bodyOf("10_accessible_form_example.html"),
  bullets: [
    ["Every input needs a ", ["<label>", "c"], ". ", ["for=\"email\"", "c"], " matches ", ["id=\"email\"", "c"], ", so the label is announced with the field, and clicking the label focuses it."],
    ["Placeholder text is ", ["not", "b"], " a label. It disappears as you type and is usually low-contrast."],
    [["aria-describedby", "c"], " connects the hint and the error to the field. They're read ", ["after", "i"], " the label."],
    [["aria-invalid=\"true\"", "c"], " marks the field as having an error. Normally JavaScript sets it after checking the input; here it's hard-coded so you can hear it."],
    ["Error text says what's wrong ", ["and how to fix it", "b"], "."],
  ], size: 14,
  notes: "Class 4 Understandable slide: 'Error Identification and Suggestions: provide helpful error messages and suggestions for fixing errors in forms.'\n\nWhat a screen reader says when you Tab into the field: 'Email address, edit, required, invalid entry. We'll only use this to send your timetable. Error: enter an email like name@example.ca'.\n\naria-describedby can list several ids separated by spaces; they're read in that order.\n\nThe 'required' attribute is native HTML: the browser announces it AND blocks submission if it's empty. Prefer it over aria-required (rule 1 again).\n\nClicking the label text focuses the input. That's a bigger click target for people with motor impairments, and it's free.\n\nIn real forms, the error message would only appear (and aria-invalid would only be set) after the user submits or leaves the field. That needs JavaScript, which comes later.",
});

// 27. Cheat sheet
{
  const s = panelSlide("ARIA Cheat Sheet",
    "Reference slide. Students can screenshot this.\n\nThe ones we used today: aria-label, aria-labelledby, aria-describedby, aria-current, aria-hidden, aria-live, aria-invalid.\n\naria-expanded is what you'll use most when we build menus later (hamburger menu buttons). aria-required exists, but the native required attribute is better because it also adds browser validation.\n\nFull reference: MDN Web Docs → ARIA (developer.mozilla.org/en-US/docs/Web/Accessibility/ARIA).");
  table(s, [
    ["Attribute", "What it does", "Example"],
    ["aria-label", "Gives a name when there's no visible text", "aria-label=\"Close menu\""],
    ["aria-labelledby", "Uses visible text (by id) as the name", "aria-labelledby=\"news-heading\""],
    ["aria-describedby", "Adds a hint or error, read after the name", "aria-describedby=\"email-error\""],
    ["aria-current", "Marks the current item in a set", "aria-current=\"page\""],
    ["aria-hidden", "Hides decoration from assistive tech", "aria-hidden=\"true\""],
    ["aria-live", "Announces content changes", "aria-live=\"polite\""],
    ["aria-expanded", "Says if a menu or section is open", "aria-expanded=\"false\""],
    ["aria-invalid", "Marks a field with an error", "aria-invalid=\"true\""],
    ["aria-required", "Marks a field as required (prefer required)", "aria-required=\"true\""],
  ], { y: 1.95, colW: [2.0, 3.0, 2.55], size: 13, codeCols: [0, 2], rowH: 0.42 });
  txt(s, "Full reference: MDN Web Docs → ARIA", { x: CX, y: 6.6, w: CW, h: 0.4, fontSize: 13, color: C.grey, italic: true });
}

// 28. Checking your work
{
  const s = panelSlide("Checking Your Work",
    "Four ways to check, from quickest to most thorough. Class 4 listed automated tools (WAVE, axe, Lighthouse) and manual testing (keyboard, screen reader). Automated tools typically catch only a minority of accessibility problems. They can check that alt exists, but not whether it's any good. They can't tell that a div is pretending to be a button.\n\nNarrator: Win+Ctrl+Enter toggles it. Caps Lock is the 'Narrator key'. Useful: Narrator+F7 (Caps Lock + F7) opens a list of links/headings/landmarks. NVDA (free from nvaccess.org) is the most popular free screen reader on Windows. On a Mac, VoiceOver is Cmd+F5.\n\nWarn students that Narrator talks a lot. Have them turn their volume down before trying it in class!");
  const rows = [
    ["⌨", "Keyboard", "Tab / Shift+Tab to move, Enter / Space to activate. Can you reach and use everything? Can you always see where focus is?", C.orange],
    ["🔍", "DevTools Accessibility pane", "Elements → Accessibility. Inspect any element to see its computed Role and Name.", C.purple],
    ["🔊", "Windows Narrator / NVDA", "Narrator is built in: Win + Ctrl + Enter to toggle. NVDA is free (nvaccess.org). Try H for headings.", C.blue],
    ["💡", "Lighthouse", "A good first pass, but automated tools only catch some problems. Always test by hand too.", C.teal],
  ];
  rows.forEach(([ic, t, d, col], i) => {
    const y = 1.95 + i * 1.25;
    numCircle(s, i + 1, CX, y + 0.05, col, 0.62);
    txt(s, t, { x: CX + 0.95, y, w: CW - 0.95, h: 0.4, fontSize: 18, bold: true, color: C.black });
    txt(s, d, { x: CX + 0.95, y: y + 0.4, w: CW - 0.95, h: 0.75, fontSize: 14, color: C.dark, valign: "top" });
  });
}

// 29. Divider: activity
dividerSlide("Class Activity", "Test → Fix → Re-test   ·   02_Semantic_HTML_ARIA/activity.html",
  "Activity time, about 35 minutes. Same tools as Class 4 (Lighthouse, Read Aloud), but this time on a page the students fix themselves.\n\nactivity.html has these planted problems: no lang attribute; divs everywhere with no landmarks; fake headings (styled divs) and a skipped heading level (h4 with no h2/h3); an image with no alt; low-contrast grey text (#bbbbbb); an icon-only × button with no name; an email input with no label; a table with no caption, th or scope; and a <div onclick> fake Subscribe button.");

// 30. Activity part 1
{
  const s = activitySlide("Activity – Part 1: Test the broken page (5 min)",
    "Expected Lighthouse results for activity.html (Accessibility category): a score well below 100, with failures such as: '<html> element does not have a [lang] attribute', 'Image elements do not have [alt] attributes', 'Buttons do not have an accessible name', 'Form elements do not have associated labels', 'Background and foreground colors do not have a sufficient contrast ratio', and 'Heading elements are not in a sequentially-descending order'. Exact wording and score vary by Lighthouse version.\n\nKeyboard test: Tab goes × button → Home → About → Programs → email input, and then leaves the page. 'Subscribe' can never be reached with the keyboard.");
  card(s, { x: 0.75, y: 1.45, w: 7.6, h: 5.5 });
  const steps = [
    ["Open it", "Right-click 02_Semantic_HTML_ARIA/activity.html → Open with Live Server."],
    ["Run Lighthouse", "Ctrl + Shift + I → Lighthouse tab → tick only Accessibility → Analyze page load."],
    ["Write it down", "Your score, and every failed audit Lighthouse lists."],
    ["Keyboard test", "Click in the page, then press Tab over and over. Can you reach Subscribe? Can you see where you are?"],
    ["Listen (optional)", "F9 for Reader View → Read Aloud, like in Class 4. What does it read? What does it skip?"],
  ];
  steps.forEach(([t, d], i) => {
    const y = 1.75 + i * 1.02;
    numCircle(s, i + 1, 1.05, y, C.orange, 0.48);
    txt(s, t, { x: 1.75, y: y - 0.02, w: 6.4, h: 0.35, fontSize: 17, bold: true, color: C.black });
    txt(s, d, { x: 1.75, y: y + 0.33, w: 6.4, h: 0.6, fontSize: 14, color: C.dark, valign: "top" });
  });
  card(s, { x: 8.75, y: 1.45, w: 3.85, h: 2.6, fill: C.dark });
  txt(s, "My score before", { x: 9.0, y: 1.65, w: 3.4, h: 0.4, fontSize: 16, color: "D9D9D9" });
  txt(s, "____ / 100", { x: 9.0, y: 2.25, w: 3.4, h: 1.0, fontSize: 40, bold: true, color: C.white });
  card(s, { x: 8.75, y: 4.35, w: 3.85, h: 2.6 });
  txt(s, [
    { text: "Hint", options: { bold: true, fontSize: 16, breakLine: true } },
    { text: "The page has at least 10 problems. Lighthouse will only find some of them. Keep your list; you'll need it in Part 3.", options: { fontSize: 14 } },
  ], { x: 9.0, y: 4.55, w: 3.4, h: 2.2, color: C.dark, valign: "top", paraSpaceAfter: 6 });
}

// 31. Activity part 2
{
  const s = activitySlide("Activity – Part 2: Fix it (25 min)",
    "Circulate while students work. Common sticking points:\n- Forgetting that heading levels must go h1 → h2 → h3 with no skips (the page jumps straight to h4).\n- Putting alt text on the figcaption instead of the img.\n- Contrast: any dark grey works. #595959 on white is about 7:1, well above the 4.5:1 AA minimum for normal text. #767676 is roughly the lightest grey that passes AA on white.\n- Wrapping the input in <label> is also valid: <label>Email <input type=\"email\"></label>.\n- The skip link's href must match the id on <main>.\n\nThe full answer key is in the notes of the next slide and in activity_answer.html.");
  const tasks = [
    ["Landmarks", "Replace the divs with header, nav, main, article, aside and footer. Add a skip link."],
    ["Headings", "Turn the big-text divs into real headings, in order: h1 → h2 → h3. No skipped levels."],
    ["Image", "Wrap it in <figure>, add good alt text, and make the caption a <figcaption>."],
    ["Language and contrast", "Add lang=\"en\" to <html>. Change the light grey #bbbbbb text to a darker grey."],
    ["Table", "Add a <caption>, <thead>/<tbody>, and <th> cells with scope."],
    ["Buttons", "Turn the Subscribe div into a real <button>. Give the × button an aria-label."],
    ["Form", "Give the email input a real <label> (for + id)."],
    ["Navigation", "Put the links in a list, add aria-label=\"Main\" to the nav and aria-current=\"page\" to Home."],
  ];
  const colW = 5.8, rowH = 1.28;
  tasks.forEach(([t, d], i) => {
    const col = Math.floor(i / 4), row = i % 4;
    const x = 0.75 + col * (colW + 0.25), y = 1.45 + row * (rowH + 0.1);
    card(s, { x, y, w: colW, h: rowH });
    numCircle(s, i + 1, x + 0.2, y + 0.22, col === 0 ? C.orange : C.teal, 0.46);
    txt(s, t, { x: x + 0.85, y: y + 0.17, w: colW - 1.05, h: 0.35, fontSize: 16, bold: true, color: C.black });
    txt(s, d, { x: x + 0.85, y: y + 0.52, w: colW - 1.05, h: 0.7, fontSize: 13, color: C.dark, valign: "top" });
  });
}

// 32. Activity part 3 (answer key in notes)
{
  const answer = fs.readFileSync(path.join(REPO, "activity_answer.html"), "utf8").replace(/\r/g, "");
  const s = activitySlide("Activity – Part 3: Test again (5 min)",
    "ANSWER KEY (activity_answer.html). For instructor reference only, not shown on the slide:\n\n" + answer);
  card(s, { x: 0.75, y: 1.45, w: 7.6, h: 5.5 });
  const steps = [
    ["Re-run Lighthouse", "Accessibility only, same as before. Write down your new score."],
    ["Re-do the keyboard test", "Tab through the page. Can you reach Subscribe now? Does the skip link appear first?"],
    ["Inspect in DevTools", "Select the × button → Accessibility pane. What's its Name now?"],
    ["Compare with a neighbour", "Did you fix the same things? Did anyone use a different element for the same content?"],
  ];
  steps.forEach(([t, d], i) => {
    const y = 1.8 + i * 1.25;
    numCircle(s, i + 1, 1.05, y, C.teal, 0.48);
    txt(s, t, { x: 1.75, y: y - 0.02, w: 6.4, h: 0.35, fontSize: 17, bold: true, color: C.black });
    txt(s, d, { x: 1.75, y: y + 0.33, w: 6.4, h: 0.7, fontSize: 14, color: C.dark, valign: "top" });
  });
  card(s, { x: 8.75, y: 1.45, w: 3.85, h: 2.6, fill: C.dark });
  txt(s, "My score after", { x: 9.0, y: 1.65, w: 3.4, h: 0.4, fontSize: 16, color: "D9D9D9" });
  txt(s, "____ / 100", { x: 9.0, y: 2.25, w: 3.4, h: 1.0, fontSize: 40, bold: true, color: C.white });
  card(s, { x: 8.75, y: 4.35, w: 3.85, h: 2.6 });
  txt(s, [
    { text: "Stuck?", options: { bold: true, fontSize: 16, breakLine: true } },
    { text: "Compare with activity_answer.html. Every change is marked with a FIX comment.", options: { fontSize: 14 } },
  ], { x: 9.0, y: 4.55, w: 3.4, h: 2.2, color: C.dark, valign: "top", paraSpaceAfter: 6 });
}

// 33. Discussion ring
{
  const s = pres.addSlide();
  s.background = { color: C.white };
  s.addShape(pres.shapes.OVAL, { x: -1.3, y: 0.35, w: 7.4, h: 6.9, fill: { type: "none" }, line: { color: C.ring, width: 34 } });
  s.addShape(pres.shapes.OVAL, { x: -1.05, y: 0.6, w: 6.9, h: 6.4, fill: { type: "none" }, line: { color: "D8E6E0", width: 8 } });
  txt(s, "How did you make out?", { x: 0.85, y: 2.9, w: 4.2, h: 1.6, fontSize: 40, color: C.navy, valign: "middle" });
  txt(s, [
    { text: "What was your Lighthouse score before and after?", options: { bullet: true, breakLine: true } },
    { text: "Lighthouse says 100. Is the page done?", options: { bullet: true, paraSpaceBefore: 14, breakLine: true } },
    { text: "Yes/no?", options: { bullet: { code: "25CB" }, indentLevel: 1, paraSpaceBefore: 6, breakLine: true } },
    { text: "What did Lighthouse not catch? (Hint: press Tab.)", options: { bullet: { code: "25CB" }, indentLevel: 1, paraSpaceBefore: 6, breakLine: true } },
    { text: "Which fix made the biggest difference to how the page sounds in Read Aloud or Narrator?", options: { bullet: true, paraSpaceBefore: 14 } },
  ], { x: 7.1, y: 2.3, w: 5.6, h: 3.2, fontSize: 20, color: C.navy, valign: "top" });
  s.addNotes("Discussion. The key teaching moment:\n\nLighthouse did NOT catch:\n- The <div onclick> fake Subscribe button (only a keyboard test finds it: Tab never reaches it).\n- The table with no <th>, caption or scope (Lighthouse only checks tables that already have headers).\n- Missing landmarks (divs instead of header/nav/main/footer). Lighthouse doesn't require landmarks.\n- Whether the alt text is actually GOOD. It only checks that alt exists; alt=\"image\" would pass.\n\nSo a page can score 100 and still be unusable with a keyboard. This is why Class 4's tools slide listed Manual Testing (keyboard and screen reader) alongside the automated tools. Automated tools catch only a fraction of real accessibility issues.");
}

// 34. Checklist
{
  const s = panelSlide("Accessibility Checklist",
    "From now on, every assignment in this course should pass this checklist. It's in the README for this class too.\n\nEncourage students to run Lighthouse AND do a keyboard test before submitting anything.");
  const items = [
    ["lang on <html> and a real <title>", "U"],
    ["One <h1>, and headings in order (no skipped levels)", "O"],
    ["<header>, <nav>, one <main>, <footer>", "R"],
    ["Every <img> has alt (empty alt=\"\" if decorative)", "P"],
    ["Text has enough contrast with its background", "P"],
    ["Every form field has a <label>; errors say how to fix them", "U"],
    ["Clickable things are <a> (goes somewhere) or <button> (does something)", "O"],
    ["ARIA only when HTML can't do the job", "R"],
    ["Tested with Lighthouse and with the keyboard", "O"],
  ];
  items.forEach(([t, k], i) => {
    const y = 1.95 + i * 0.56;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: CX, y: y + 0.06, w: 0.32, h: 0.32, rectRadius: 0.05, fill: { color: C.white }, line: { color: C.dark, width: 1.5 } });
    txt(s, t, { x: CX + 0.55, y, w: 5.6, h: 0.45, fontSize: 15.5, color: C.black, valign: "middle" });
    const p = POUR[k];
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: CX + 6.3, y: y + 0.07, w: 1.25, h: 0.3, rectRadius: 0.15, fill: { color: p.color }, line: { color: p.color } });
    txt(s, p.name.toUpperCase(), { x: CX + 6.3, y: y + 0.07, w: 1.25, h: 0.3, fontSize: 8.5, bold: true, color: C.white, align: "center", valign: "middle" });
  });
}

// 35. Recap
{
  const s = panelSlide("Class Recap and Next Steps",
    "Recap the day. Next class: Advanced HTML/CSS techniques and responsive design, as previewed at the end of Class 4. Several things from today come back with CSS: hiding the skip link until it's focused, styling focus indicators, colour contrast, and aria-current for styling the active nav link.");
  labelBullets(s, [
    { label: "Recap:", text: [
      "Semantic HTML5 (header, nav, main, article, section, aside, footer, figure, accessible tables) gives the page structure screen readers can navigate.",
      "ARIA edits the accessibility tree when HTML can't. Native HTML first; no ARIA is better than bad ARIA.",
      "Automated tools like Lighthouse are a start. Always test with the keyboard too.",
    ] },
    { label: "Next Class Preview:", text: "Deep dive into Advanced HTML/CSS Techniques and Responsive Design, including styling focus indicators and the skip link from today." },
  ], { size: 20, sub: 17 });
}

pres.writeFile({ fileName: OUT }).then((f) => console.log("wrote " + f));
