# Class 8: Intro to CSS

> This README is the source for:
> 1. `05_Intro_to_CSS/*.html`: numbered code samples, each with its CSS in a `<style>` block (Example 01 also uses `01_external.css`)
> 2. `activity/`: the Harbourside Fish & Chips page students style during class, plus `activity/answer/`
> 3. `Class8_Intro_to_CSS.pptx`: the class deck (Class 5–7 style; every slide has speaker notes)
>
> Rebuild the deck: `cd _build`, `npm install` once, then `npm run build`. The build takes a screenshot of each example with Microsoft Edge (headless) for the "Result" pictures, so edit an example and rebuild to update its slide.

---

## Format: concept → try it

After a short intro (what CSS is, how a rule is written, common mistakes), the class runs in **seven rounds**. Each round is one concept slide (code + bullets + the rendered result) followed by one **Try It** slide (steps on the left, the code to type on the right). Every Try It adds to the same file, `activity/style.css`, under a heading for that activity, so by the end each student has a fully styled Harbourside page.

| Round | Concept | Example | Try It (in activity/) | Time |
|---|---|---|---|---|
| 1 | Three ways to add CSS (inline, internal, external) | 01 + 01_external.css | Link style.css, set the body background | 5 min |
| 2 | Selectors: element, `.class`, `#id`, group, descendant | 02 | Add `class="price"` and `class="special"`; style them | 8 min |
| 3 | Colour (name, hex, rgb) and contrast (4.5 : 1) | 03 | Navy header, white nav, dark orange prices; check contrast in DevTools | 8 min |
| 4 | Text and fonts: font stacks, `rem`, line-height, text-transform | 04 | Fonts and text styles; bonus Google Font | 8 min |
| 5 | The box model: padding, border, margin, box-sizing, `margin: 0 auto` | 05 | Centred main column, menu cards, side-by-side nav | 12 min |
| 6 | Styling images: `max-width: 100%`, border-radius, object-fit | 06 | Flexible, rounded image; **swap in their own photo from Class 7** | 8 min |
| 7 | The cascade: specificity, order, inheritance, DevTools Styles | 07 | **DevTools detective**: why did the Lobster Roll lose its highlight? | 10 min |
| — | Make it yours (at least three changes), submit screenshot + style.css | | | 10 min / homework |

**Timing:** about 15 min of intro, then about 35 min of concept slides and about 60 min of Try Its.

### Two planned "bugs"
- **Try It 3:** after the header gets `color: white`, the `h1` stays navy (from the `h1, h2` rule in Try It 2). A rule aimed at an element beats an inherited value. Fixed with `header h1`. Explained properly in Concept 7.
- **Try It 5:** `#menu article { background-color: white; }` wipes out the `.special` highlight from Try It 2, because an id beats a class. Don't explain it then. Students solve it in **Try It 7** with DevTools, and fix it with `#menu .special`.

---

## Folder

```
05_Intro_to_CSS/
├── README.md
├── 01_three_ways_example.html + 01_external.css
├── 02_selectors_example.html
├── 03_colour_example.html            ← contrast: 1.9 : 1 fail, 3.1 : 1 fail (logo orange), 5.5 : 1 pass
├── 04_text_fonts_example.html
├── 05_box_model_example.html         ← two "300px" boxes: 346px vs 300px
├── 06_styling_images_example.html
├── 07_cascade_example.html           ← "Who wins?" prediction game
├── images/                           ← harbour.jpg (800px, 45 KB) and logo.svg, from Class 7's samples
├── activity/
│   ├── index.html                    ← starter: semantic HTML, no CSS linked
│   ├── style.css                     ← starter: empty, with a heading per activity
│   ├── images/
│   └── answer/
│       ├── index.html                ← with <link>, class="price", class="special"
│       └── style.css                 ← complete, organized by activity, with comments
├── Class8_Intro_to_CSS.pptx
└── _build/                           ← build_deck.js (Edge screenshots go in _build/shots, git-ignored)
```

## Colours used (all pass WCAG AA)

| Use | Colour | Contrast |
|---|---|---|
| Header background / headings | `#12304A` navy | white on navy 13.6 : 1 |
| Prices | `#B5451B` dark orange | 5.5 : 1 on white, 4.65 : 1 on the peach highlight |
| Special highlight | `#FBE9DF` peach | |
| Footer | `#262626` on `#F2F2F2` | 13.5 : 1 |
| Logo orange (decoration only) | `#E97132` | 3.1 : 1. **Fails** for body text (used as a teaching point) |

---

## Deck outline (30 slides)

| # | Slide |
|---|---|
| 1–4 | Title · Attendance · Overview (concept → try it) · Get the code |
| 5–7 | What is CSS? (before/after screenshots) · Anatomy of a rule · Watch out for… (color not colour, semicolons, units) |
| 8 | *Concept → Try It* divider |
| 9–11 | Three ways to add CSS · Example 01 · **Try It 1** |
| 12–13 | Selectors (Example 02) · **Try It 2** |
| 14–15 | Colour and contrast (Example 03) · **Try It 3** |
| 16–17 | Text and fonts (Example 04) · **Try It 4** |
| 18–20 | The box model diagram · Example 05 · **Try It 5** |
| 21–22 | Styling images (Example 06) · **Try It 6** |
| 23–26 | The cascade: who wins? · Example 07 (the reveal) · DevTools Styles pane · **Try It 7** |
| 27 | Make it yours, then submit |
| 28–30 | How did you make out? · CSS cheat sheet · Recap and next steps |

## CSS cheat sheet (slide 29)

| Property | Example | What it does |
|---|---|---|
| `color` | `color: #12304A;` | Text colour |
| `background-color` | `background-color: white;` | Colour behind the element |
| `font-family` | `font-family: Georgia, serif;` | The font (with fallbacks) |
| `font-size` | `font-size: 1.2rem;` | Text size (prefer rem) |
| `font-weight` / `font-style` | `font-weight: bold;` | Bold / italic |
| `line-height` | `line-height: 1.6;` | Space between lines |
| `text-align` / `text-transform` | `text-align: center;` | Alignment / capitals |
| `text-decoration` | `text-decoration: none;` | Remove or add underlines |
| `padding` / `border` / `margin` | `padding: 10px 20px;` | The box model layers |
| `max-width` + `margin: auto` | `margin: 0 auto;` | A centred column |
| `border-radius` | `border-radius: 8px;` | Rounded corners (50% = circle) |
| `box-sizing` | `box-sizing: border-box;` | Width includes padding and border |

**Submit (participation):** a screenshot of the styled page, plus their `style.css`, uploaded to Brightspace.
