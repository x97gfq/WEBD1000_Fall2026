// Harbourside Fish & Chips wireframe: one layout, two outputs.
//   node wireframe.js        → writes ../harbourside_ordering_wireframe.drawio
//   require("./wireframe")   → { desktop, mobile } cell lists, used by build_deck.js to draw the slides
// Each cell has a `step` (the guided-example step that adds it) so the deck can show the page being built up.
// step "meta" = page title text, step "note" = yellow teaching notes.
const fs = require("fs");
const path = require("path");

const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
  .replace(/"/g, "&quot;").replace(/\n/g, "&#xa;");

function page(name) {
  const cells = [];
  let n = 0, OX = 0, OY = 0;
  const P = { name, cells, step: 1 };
  P.origin = (x, y) => { OX = x; OY = y; };
  P.shape = (x, y, w, h, style, value = "", kind = "") => {
    const id = `${name.toLowerCase()}-${++n}`;
    cells.push({ id, x: x + OX, y: y + OY, w, h, style, value, kind, step: P.step });
    return id;
  };
  P.edge = (source, target, style) => {
    cells.push({ id: `${name.toLowerCase()}-${++n}`, edge: true, source, target, style, value: "", step: P.step });
  };
  P.xml = () => {
    const body = cells.map((c) => c.edge
      ? `<mxCell id="${c.id}" value="" style="${c.style}" edge="1" parent="1" source="${c.source}" target="${c.target}"><mxGeometry relative="1" as="geometry"/></mxCell>`
      : `<mxCell id="${c.id}" value="${esc(c.value)}" style="${c.style}" vertex="1" parent="1"><mxGeometry x="${c.x}" y="${c.y}" width="${c.w}" height="${c.h}" as="geometry"/></mxCell>`).join("");
    return `<diagram id="${name.toLowerCase()}" name="${name}"><mxGraphModel dx="1400" dy="900" grid="1" gridSize="10" guides="1" tooltips="1" connect="1" arrows="1" fold="1" page="0" pageScale="1" pageWidth="1600" pageHeight="1200" math="0" shadow="0"><root><mxCell id="0"/><mxCell id="1" parent="0"/>${body}</root></mxGraphModel></diagram>`;
  };
  // building blocks
  const T = (o = {}) => `text;html=0;whiteSpace=wrap;strokeColor=none;fillColor=none;align=${o.align || "left"};verticalAlign=middle;fontSize=${o.size || 14};fontColor=${o.color || "#333333"};fontStyle=${o.style || 0};spacing=0;`;
  P.text = (x, y, w, h, v, o) => P.shape(x, y, w, h, T(o), v);
  P.box = (x, y, w, h, fill = "#FFFFFF", stroke = "#999999", extra = "") => P.shape(x, y, w, h, `rounded=0;whiteSpace=wrap;fillColor=${fill};strokeColor=${stroke};${extra}`);
  P.img = (x, y, w, h) => P.shape(x, y, w, h, "shape=mxgraph.mockup.graphics.simpleIcon;html=0;strokeColor=#999999;fillColor=#FFFFFF;");
  P.btn = (x, y, w, h, v, filled = true, size = 14) => P.shape(x, y, w, h,
    `rounded=1;arcSize=15;whiteSpace=wrap;html=0;fontSize=${size};fontStyle=1;` +
    (filled ? "fillColor=#555555;strokeColor=#555555;fontColor=#FFFFFF;" : "fillColor=#FFFFFF;strokeColor=#555555;fontColor=#333333;"), v);
  P.line = (x, y, w) => P.box(x, y, w, 1, "#BBBBBB", "none");
  P.lorem = (x, y, w) => P.box(x, y, w, 8, "#DDDDDD", "none");
  // Landmark annotation: blue dashed outline + tag
  P.landmark = (x, y, w, h, tag) => {
    P.shape(x, y, w, h, "rounded=0;fillColor=none;dashed=1;strokeColor=#1E88E5;strokeWidth=2;", "", "landmark");
    return P.shape(x, y - 20, tag.length * 8 + 14, 20, "rounded=0;fillColor=#1E88E5;strokeColor=none;fontColor=#FFFFFF;fontStyle=1;fontSize=12;fontFamily=Courier New;", tag, "landmark");
  };
  P.note = (x, y, w, h, v) => P.shape(x, y, w, h, "shape=note;size=14;whiteSpace=wrap;html=0;fillColor=#FFF2CC;strokeColor=#D6B656;fontColor=#333333;fontSize=13;align=left;verticalAlign=top;spacing=10;", v, "note");
  return P;
}

const items = [["Haddock & Chips (1 pc)", "$14.99"], ["Haddock & Chips (2 pc)", "$19.99"], ["Fish Tacos (2)", "$13.49"],
  ["Clam Strips & Chips", "$16.99"], ["Scallops & Chips", "$24.99"], ["Lobster Roll", "$22.99"]];
const order = [["Haddock & Chips (2 pc)", "$19.99"], ["Fries (large)", "$5.49"], ["Coleslaw", "$3.49"]];
const totals = [["Subtotal", "$28.97", 0], ["HST (14%)", "$4.06", 0], ["Total", "$33.03", 1]];

/* ================= DESKTOP ================= */
const D = page("Desktop");
D.step = "meta";
D.text(40, -10, 900, 30, "Harbourside Fish & Chips: Online Ordering (Desktop, 1200px wide)", { size: 22, style: 1 });
D.text(40, 20, 900, 20, "Low-fidelity wireframe. Blue dashed boxes = the HTML landmark each region becomes (not part of the design).", { size: 13, color: "#1E88E5" });
D.origin(40, 70);

// Step 1: browser frame, header and nav
D.step = 1;
D.box(0, 0, 1200, 1060, "#FFFFFF", "#333333", "strokeWidth=2;");
D.box(0, 0, 1200, 36, "#EEEEEE", "#333333", "strokeWidth=2;");
D.shape(100, 6, 500, 24, "rounded=1;arcSize=50;fillColor=#FFFFFF;strokeColor=#BBBBBB;fontSize=12;fontColor=#666666;align=left;spacingLeft=10;", "harbourside.ca/order");
[14, 34, 54].forEach((x) => D.shape(x, 12, 12, 12, "ellipse;fillColor=#CCCCCC;strokeColor=none;"));
D.img(20, 51, 50, 50);
D.text(80, 56, 280, 40, "Harbourside Fish & Chips", { size: 20, style: 1 });
[["Menu", 480, 0], ["Order Online", 560, 5], ["Hours & Location", 690, 0], ["Contact", 840, 0]]
  .forEach(([t, x, s]) => D.text(x, 64, 130, 24, t, { size: 15, style: s }));
D.btn(1060, 58, 120, 36, "Cart (3)", false);
D.line(0, 116, 1200);

// Step 2: pickup/delivery banner
D.step = 2;
D.box(20, 132, 1160, 80, "#F5F5F5", "#BBBBBB");
D.text(40, 142, 450, 30, "Order for pickup or delivery", { size: 22, style: 1 });
D.text(40, 172, 450, 24, "Ready in about 20 minutes", { size: 14, color: "#777777" });
D.btn(660, 152, 120, 40, "Pickup", true);
D.btn(780, 152, 120, 40, "Delivery", false);
D.shape(920, 152, 240, 40, "rounded=1;arcSize=15;fillColor=#FFFFFF;strokeColor=#555555;fontSize=14;align=left;spacingLeft=12;", "Pickup time: ASAP        ▾");

// Step 3: category tabs and menu cards
D.step = 3;
[["Fish", 1], ["Combos", 0], ["Sides", 0], ["Drinks", 0], ["Desserts", 0]]
  .forEach(([t, f], i) => D.btn(20 + i * 120, 232, 110, 40, t, !!f));
D.text(20, 290, 300, 32, "Fish", { size: 22, style: 1 });
items.forEach(([name, price], i) => {
  const x = 20 + (i % 3) * 260, y = 332 + Math.floor(i / 3) * 280;
  D.box(x, y, 240, 260, "#FFFFFF", "#999999");
  D.img(x + 10, y + 10, 220, 120);
  D.text(x + 10, y + 138, 220, 24, name, { size: 16, style: 1 });
  D.lorem(x + 10, y + 170, 200);
  D.lorem(x + 10, y + 186, 150);
  D.text(x + 10, y + 212, 90, 36, price, { size: 16, style: 1 });
  D.btn(x + 130, y + 212, 100, 36, "Add +", true);
});

// Step 4: order summary
D.step = 4;
const ax = 810, ay = 232;
D.box(ax, ay, 370, 440, "#F5F5F5", "#999999");
D.text(ax + 20, ay + 16, 250, 30, "Your Order", { size: 20, style: 1 });
order.forEach(([name, price], i) => {
  const y = ay + 64 + i * 50;
  D.text(ax + 20, y, 180, 28, name, { size: 14 });
  D.btn(ax + 206, y, 28, 28, "−", false);
  D.text(ax + 234, y, 24, 28, "1", { align: "center" });
  D.btn(ax + 258, y, 28, 28, "+", false);
  D.text(ax + 290, y, 60, 28, price, { align: "right" });
});
D.line(ax + 20, ay + 220, 330);
totals.forEach(([l, v, b], i) => {
  const y = ay + 232 + i * 30;
  D.text(ax + 20, y, 200, 26, l, { size: b ? 17 : 14, style: b });
  D.text(ax + 250, y, 100, 26, v, { size: b ? 17 : 14, style: b, align: "right" });
});
D.btn(ax + 20, ay + 336, 330, 48, "Checkout", true, 17);
D.text(ax + 20, ay + 394, 330, 24, "Pickup · ASAP", { size: 13, color: "#777777", align: "center" });

// Step 5: footer, then landmark labels over everything
D.step = 5;
D.line(0, 916, 1200);
D.text(20, 932, 360, 80, "Harbourside Fish & Chips\n12 Water St, Yarmouth NS\n(902) 555-0142", { size: 14 });
D.text(420, 932, 360, 80, "Hours\nTue–Sun 11 am – 8 pm\nClosed Mondays", { size: 14 });
D.text(820, 932, 200, 24, "Follow us", { size: 14, style: 1 });
[820, 860, 900].forEach((x) => D.shape(x, 964, 30, 30, "ellipse;fillColor=#FFFFFF;strokeColor=#999999;"));
D.text(20, 1020, 1160, 24, "© 2026 Harbourside Fish & Chips", { size: 12, color: "#777777", align: "center" });
D.landmark(8, 44, 1184, 66, "<header>");
D.landmark(470, 56, 470, 40, "<nav>");
D.landmark(274, 326, 252, 272, "<article>");
D.landmark(ax - 8, ay - 8, 386, 456, "<aside>");
D.landmark(8, 124, 1184, 776, "<main>");
D.landmark(8, 924, 1184, 128, "<footer>");

// teaching notes
D.step = "note";
D.note(1230, 44, 240, 80, "① The current page (\"Order Online\") is bold in the nav so users know where they are.");
D.note(1230, 140, 240, 80, "② The two choices users make first (pickup/delivery and time) go at the top of the page.");
D.note(1230, 232, 240, 100, "③ The order summary stays next to the menu, so users can see their total while they add items.");
D.note(1230, 346, 240, 100, "④ Every menu card uses the same layout (image, name, description, price, button), so it becomes one repeated <article>.");

/* ================= MOBILE ================= */
const M = page("Mobile");
M.step = "meta";
M.text(40, -10, 1100, 30, "Harbourside Fish & Chips: Online Ordering (Mobile, 375px wide)", { size: 22, style: 1 });
M.text(40, 20, 1100, 20, "Same content as desktop, reorganised for a phone. Yellow notes = what changed and why.", { size: 13, color: "#1E88E5" });
M.step = 6;

// Phone A: menu (full scroll length)
M.origin(40, 70);
M.shape(0, 0, 375, 1150, "rounded=1;arcSize=3;fillColor=#FFFFFF;strokeColor=#333333;strokeWidth=2;");
M.box(3, 3, 369, 21, "#EEEEEE", "none");
M.text(160, 0, 55, 24, "9:41", { size: 12, style: 1, align: "center" });
M.img(16, 32, 40, 40);
M.text(64, 36, 170, 32, "Harbourside", { size: 18, style: 1 });
M.btn(245, 34, 70, 36, "Cart (3)", false, 12);
M.btn(323, 34, 40, 36, "☰", false, 18);
M.line(0, 84, 375);
M.landmark(4, 28, 367, 52, "<header>");

M.box(16, 96, 343, 140, "#F5F5F5", "#BBBBBB");
M.text(28, 104, 320, 28, "Order for pickup or delivery", { size: 17, style: 1 });
M.btn(28, 138, 157, 40, "Pickup", true);
M.btn(185, 138, 157, 40, "Delivery", false);
M.shape(28, 186, 314, 40, "rounded=1;arcSize=15;fillColor=#FFFFFF;strokeColor=#555555;fontSize=14;align=left;spacingLeft=12;", "Pickup time: ASAP                     ▾");

[["Fish", 1], ["Combos", 0], ["Sides", 0], ["Drinks", 0]].forEach(([t, f], i) => M.btn(16 + i * 98, 252, 90, 40, t, !!f));
M.text(16, 306, 200, 28, "Fish", { size: 20, style: 1 });

items.forEach(([name, price], i) => {
  const y = 344 + i * 108;
  M.img(16, y + 8, 80, 80);
  M.text(108, y + 8, 190, 22, name, { size: 15, style: 1 });
  M.lorem(108, y + 38, 180);
  M.text(108, y + 56, 100, 26, price, { size: 15, style: 1 });
  M.btn(311, y + 50, 48, 40, "+", true, 20);
  M.line(16, y + 100, 343);
});
M.landmark(4, 90, 367, 904, "<main>");

// fold marker
M.shape(-20, 812, 415, 1, "rounded=0;fillColor=none;strokeColor=#E53935;dashed=1;strokeWidth=2;");
M.text(-20, 816, 200, 18, "bottom of the first screen (812px)", { size: 11, color: "#E53935" });

// sticky order bar (drawn at bottom of first screen)
const bar = M.shape(12, 740, 351, 60, "rounded=1;arcSize=20;fillColor=#333333;strokeColor=none;fontColor=#FFFFFF;fontSize=16;fontStyle=1;shadow=1;", "View order (3) · $33.03   →");

M.line(0, 1000, 375);
M.text(16, 1012, 343, 60, "Harbourside Fish & Chips\n12 Water St, Yarmouth NS · (902) 555-0142", { size: 13 });
M.text(16, 1072, 343, 40, "Tue–Sun 11 am – 8 pm · Closed Mondays", { size: 13 });
M.text(16, 1112, 343, 24, "© 2026 Harbourside Fish & Chips", { size: 11, color: "#777777" });
M.landmark(4, 1006, 367, 136, "<footer>");

// Phone B: Your Order screen
M.origin(760, 70);
const phoneB = M.shape(0, 0, 375, 812, "rounded=1;arcSize=4;fillColor=#FFFFFF;strokeColor=#333333;strokeWidth=2;");
M.box(3, 3, 369, 21, "#EEEEEE", "none");
M.text(160, 0, 55, 24, "9:41", { size: 12, style: 1, align: "center" });
M.btn(16, 34, 84, 36, "← Menu", false, 13);
M.text(110, 36, 155, 32, "Your Order", { size: 18, style: 1, align: "center" });
M.line(0, 84, 375);
M.landmark(4, 28, 367, 52, "<header>");

M.box(16, 96, 343, 56, "#F5F5F5", "#BBBBBB");
M.text(28, 108, 250, 32, "Pickup · ASAP (about 20 min)", { size: 14 });
M.text(290, 108, 60, 32, "Change", { size: 14, style: 4, align: "right" });

order.forEach(([name, price], i) => {
  const y = 172 + i * 64;
  M.text(16, y, 220, 24, name, { size: 15, style: 1 });
  M.text(16, y + 26, 100, 22, price, { size: 14, color: "#777777" });
  M.btn(247, y + 4, 40, 44, "−", false, 18);
  M.text(287, y + 4, 28, 44, "1", { align: "center", size: 16 });
  M.btn(315, y + 4, 44, 44, "+", false, 18);
  M.line(16, y + 56, 343);
});

M.text(16, 372, 250, 22, "Special instructions", { size: 14, style: 1 });
M.shape(16, 398, 343, 72, "rounded=1;arcSize=8;fillColor=#FFFFFF;strokeColor=#999999;fontSize=13;fontColor=#AAAAAA;align=left;verticalAlign=top;spacing=10;", "e.g. extra tartar sauce, no vinegar");

totals.forEach(([l, v, b], i) => {
  const y = 492 + i * 30;
  M.text(16, y, 200, 26, l, { size: b ? 17 : 14, style: b });
  M.text(259, y, 100, 26, v, { size: b ? 17 : 14, style: b, align: "right" });
});
M.btn(16, 732, 343, 56, "Checkout · $33.03", true, 17);
M.landmark(4, 90, 367, 710, "<main>");

M.edge(bar, phoneB, "edgeStyle=orthogonalEdgeStyle;endArrow=block;endFill=1;dashed=1;strokeColor=#1E88E5;strokeWidth=2;exitX=1;exitY=0.5;entryX=0;entryY=0.8;");

// teaching notes (between phones)
M.origin(0, 0);
M.step = "note";
M.note(440, 90, 290, 90, "① The nav links move behind the ☰ button. Cart stays visible because it is the main thing users need on this page.");
M.note(440, 310, 290, 80, "② The category tabs scroll sideways instead of wrapping onto a second row.");
M.note(440, 420, 290, 110, "③ The 3-column card grid becomes a 1-column list. The + buttons are 48×40 so they are easy to tap with a thumb (aim for at least 44px).");
M.note(440, 790, 290, 110, "④ The desktop order <aside> is replaced by a bar that stays at the bottom of the screen. Tapping it opens the \"Your Order\" screen →");
M.note(1160, 90, 260, 120, "⑤ New screen, not a popup. It uses a normal <form> with a <label> for Special instructions (Class 5, Example 10).");
M.note(1160, 802, 260, 90, "⑥ Checkout sits at the bottom of the screen, within easy reach of a thumb.");

module.exports = { desktop: D.cells, mobile: M.cells };

if (require.main === module) {
  const out = `<mxfile host="app.diagrams.net" agent="WEBD1000 generator" version="24.0.0">${D.xml()}${M.xml()}</mxfile>\n`;
  const file = path.join(__dirname, "..", "harbourside_ordering_wireframe.drawio");
  fs.writeFileSync(file, out);
  console.log("wrote " + file);
}
