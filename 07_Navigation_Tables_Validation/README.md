# Class 10: Navigation Menus, Tables and HTML Validation

> This README is the source for:
> 1. `07_Navigation_Tables_Validation/*.html`: twelve numbered code samples, each with its CSS in a `<style>` block (01–04 navigation, 05–08 tables, 09–10 HTML validation, 11–12 CSS validation)
> 2. `activity/`: Harbourside, Part 2 (eight Try Its on the Class 8 page), plus `activity/answer/`
> 3. `Navigation_Tables_Validation.pptx`: the class deck (Class 5–9 style; every slide has speaker notes)
>
> Rebuild the deck: `cd _build`, `npm install` once, then `npm run build`. The build screenshots each example with Microsoft Edge (headless). For hover states (the open dropdown, a hovered row) it copies every `:hover` rule onto a class, so the slides show the hovered look.

Week 7 has **one class**: Wednesday, October 21 is Assessment & Evaluation day (college open, no classes).

---

## Format: concept → try it, three parts

Every Try It adds to the same Harbourside page, like Class 8. The starter `style.css` already holds the Class 8 CSS (tidied up, plus 3 planted mistakes for Try It 8), with an empty heading for each Try It underneath.

| Part | Concept | Example | Try It (in activity/) | Time |
|---|---|---|---|---|
| — | Intro, get the code | | | 8 min |
| Nav | Every menu is a list: `nav > ul > li > a` | 01 | | |
| Nav | Horizontal: `list-style: none`, `display: flex` + `gap`, `display: block` links, `aria-current` | 02 | **1** Nav HTML + horizontal · **2** Style the links | 6 + 8 + 6 min |
| Nav | Vertical: transparent border, current page | 03 | | 3 min |
| Nav | Dropdown: nested `ul`, `position: relative / absolute`, `:hover` **and** `:focus-within` | 04 | **3** A dropdown · **4** Make it readable | 5 + 10 + 4 min |
| Tables | Parts: caption, thead, tbody, tr, th scope, td | 05 | | 4 min |
| Tables | `border-collapse`, padding on cells, right-aligned numbers | 06 | | 3 min |
| Tables | Zebra stripes `tbody tr:nth-child(even)`, row hover | 07 | **5** Hours table HTML | 3 + 10 min |
| Tables | Small screens: a scrolling wrapper `div` | 08 | **6** Style the table | 3 + 8 min |
| HTML validation | Browsers forgive, validators don't · the three ways in · reading results | 09, 10 | **7** Validate the HTML: fix the 5 hidden mistakes | 10 + 12 min |
| CSS validation | Browsers skip broken CSS · jigsaw.w3.org/css-validator · reading results | 11, 12 | **8** Validate the CSS: fix the 3 hidden mistakes | 4 + 6 min |
| — | Make it yours, submit | | | 7 min |

That's about 117 minutes, so there's very little slack. If time runs short, Try It 8 is fine as homework.

### Planned surprises
- **Try It 3:** the dropdown opens, but its links are white (from Try It 2) on a transparent background, so they almost vanish. Try It 4 gives it a white panel and navy links.
- **Try It 3:** the dropdown's *Kids' Fish Bites* link goes nowhere. Don't explain it. In Try It 7 the validator reports `Duplicate ID "clams"`: the Kids article was copy-and-pasted as `id="clams"`.
- **Try It 7:** the starter `index.html` has five deliberate mistakes in code students didn't type:

| # | Line (approx.) | Mistake | Validator says | Fix |
|---|---|---|---|---|
| 1 | 15 | Header logo `<img>` has no `alt` | An "img" element must have an "alt" attribute | `alt=""` (decoration: the h1 says the name) |
| 2 | 30 | `<strong>…<em>boats</strong></em>` | End tag "strong" violates nesting rules (+ a second message) | `</em></strong>` |
| 3 | 44 | `<h3>Clam Strips</h2>` | End tag "h2" seen, but there were open elements | `</h3>` |
| 4 | 55 | Kids article is `id="clams"` too | Duplicate ID "clams" | `id="kids"` |
| 5 | 77 | `<p font="small">` | Attribute "font" not allowed on element "p" | `class="small"` + CSS |

Line numbers move depending on what each student typed. Five mistakes give six error messages. All five are fixed, with a comment on each, in `activity/answer/index.html`.
- **Try It 8:** the Class 8 part of the starter `style.css` has three deliberate mistakes. The browser skips them silently, and the HTML validator can't see them because it never opens a linked stylesheet:

| # | Line (approx.) | Mistake | CSS validator says | What it breaks |
|---|---|---|---|---|
| 1 | 50 | `main { padding: 20; }` | only "0" can be a "unit"… | `main` has no padding |
| 2 | 70 | `figcaption { colour: #595959; }` | Property "colour" doesn't exist… | caption is black, not grey |
| 3 | 74 | `font-weight: bold` with no `;` | Missing a semicolon before the property name "color" | prices are neither bold nor orange |

Fixed and commented (CSS Fix 1–3) in `activity/answer/style.css`.

**Validator gotchas found while building this** (checked against validator.w3.org, Oct 2026):
- `align="center"` now gets only a *warning* ("obsolete"), not an error, so it's no good as a planted mistake.
- An `<img>` with no `alt` inside a `<figure>` that has a `<figcaption>` is *not* reported. That's why mistake 1 is on the header logo.
- validator.w3.org *does* report CSS errors inside a `<style>` block (Example 11 gets the same four errors there), but it never downloads a linked `style.css`. That's the case for the CSS validator.
- jigsaw.w3.org/css-validator's "By direct input" accepts a whole HTML file and checks its `<style>` block; line numbers match the file.

**Submit (participation):** three screenshots on Brightspace: the page with the Menu dropdown open, the HTML validator's "No errors or warnings to show.", and the CSS validator's "Congratulations! No Error Found."

### Changes from last year's deck
- Horizontal menu uses `display: flex` + `gap` instead of `inline-block` (a lead-in to next class's Flexbox).
- Dropdown adds `position: relative` on the `li` (missing before) and `:focus-within`, so keyboard users can open it (WCAG 2.1.1).
- Responsive tables use a wrapper `div` with `overflow-x: auto` instead of `display: block` on the `<table>`, which can strip table semantics for screen readers.
- Zebra stripes target `tbody tr`, and every table uses caption, thead, tbody and `scope` (from Class 4).
- Validation slide drops the "improves search engine rankings" claim. The speaker notes explain why.
- Validation is hands-on: broken examples (09 HTML, 11 CSS) demoed live, then planted mistakes in the activity (5 HTML, 3 CSS).
- Adds the W3C CSS validator, with its own example, Try It and submission screenshot.

---

## Folder

```
07_Navigation_Tables_Validation/
├── README.md
├── 01_nav_structure_example.html     ← an unstyled nav: it's just a list
├── 02_nav_horizontal_example.html
├── 03_nav_vertical_example.html
├── 04_nav_dropdown_example.html      ← hover OR Tab opens it (:focus-within)
├── 05_table_structure_example.html   ← caption, thead, tbody, scope, no CSS
├── 06_table_styled_example.html
├── 07_table_zebra_example.html
├── 08_table_responsive_example.html  ← wide table in a scrolling wrapper
├── 09_validation_broken_example.html ← looks fine, 4 mistakes, 5 error messages
├── 10_validation_fixed_example.html  ← each fix commented
├── 11_css_validation_broken_example.html ← CSS the browser silently skips: 4 errors
├── 12_css_validation_fixed_example.html  ← each fix commented
├── images/                           ← logo.svg (used by 09 and 10), harbour.jpg
├── activity/
│   ├── index.html                    ← starter: Class 8 page, no nav, 5 hidden mistakes
│   ├── style.css                     ← starter: Class 8 CSS (3 hidden mistakes) + a heading per Try It
│   ├── images/
│   └── answer/
│       ├── index.html                ← nav, dropdown, hours table; Fix 1–5 commented
│       └── style.css                 ← complete, organized by Try It; CSS Fix 1–3 commented
├── Navigation_Tables_Validation.pptx
└── _build/                           ← build_deck.js (Edge screenshots go in _build/shots, git-ignored)
```

## Deck outline (36 slides)

| # | Slide |
|---|---|
| 1–4 | Title · Attendance · Overview · Get the code |
| 5–6 | *Navigation Menus* divider · Three kinds of menu (same HTML every time) |
| 7–10 | Examples 01–04 |
| 11 | Harbourside, Part 2: before/after, the eight Try Its, what to submit |
| 12–15 | Try It 1 Add the nav · 2 Style the links · 3 A dropdown · 4 Make it readable |
| 16–17 | *Tables* divider · The parts of a table |
| 18–21 | Examples 05–08 |
| 22–23 | Try It 5 An hours table · 6 Style the table |
| 24–26 | *Validation* divider · What and why · How to use the W3C validator |
| 27–28 | Example 09 (looks fine, isn't) · Reading the results (real validator output) |
| 29 | Try It 7: Validate the HTML (clues for the 5 mistakes) |
| 30–31 | Example 11 (CSS fails quietly, beside Example 12) · The W3C CSS validator (real output) |
| 32 | Try It 8: Validate the CSS (clues for the 3 mistakes) |
| 33 | Make it yours, then submit |
| 34–36 | Cheat sheet: menus and tables · Cheat sheet: validator messages (HTML and CSS) · Recap and next steps |

## Colours used (all pass WCAG AA)

Same palette as Class 8: navy `#12304A` (white on navy 13.6 : 1), hover teal `#1B5E82` (white on teal about 7 : 1), current-page underline `#E97132` (decoration, 3 : 1 non-text), dropdown hover peach `#FBE9DF`, table stripe `#E9EFF2`.
