# Week 8: Flexbox, CSS Grid and Float (two classes)

> This README is the source for:
> 1. `08_Flex_Grid_CSS/*.html` + `*.css`: seventeen numbered examples. **Each example is two files:** an HTML page and its own external stylesheet with the same name (01–09 flexbox, 10–14 grid, 15–16 float, 17 a responsive layout that uses all three)
> 2. `playground/`: the Flexbox Playground and the Grid Playground (pick values, watch the boxes move, see the CSS you'd write)
> 3. `activity/`: Harbourside, Part 3 (eight Try Its on the Week 7 page), plus `activity/answer/`
> 4. Two decks, one per class (same style as Classes 5–11; every slide has speaker notes): `Flexbox.pptx` (Week 8, Class 1) and `Grid_and_Float.pptx` (Week 8, Class 2)
>
> Rebuild the decks: `cd _build`, `npm install` once, then `npm run build` (it writes both). The build screenshots every example with Microsoft Edge (headless). Code on the slides is read from the example `.css` files, and Try It code from `activity/answer/style.css`, so edit those files and rebuild.

Split from the original Week 8 deck (32.108_PPT_Wk08): Class 1 is Flexbox; Class 2 is Grid, float, "when to use which", and the best-practices slide. The original's W3Schools work time, screenshot submission, Fireship videos and the class activity are all kept (the activity is now Harbourside, Part 3).

---

## Interactive parts

Both decks are built to keep everyone involved, with something every 5 to 10 minutes:

| Kind | How it works | Class 1 | Class 2 |
|---|---|---|---|
| **Pop quiz** | Question slide, everyone votes at once (1–4 fingers for A–D, or the Teams chat), then the answer slide with a "Why?" | 5 | Warm-up + 4 |
| **Predict the layout** | Code and an empty "your prediction" box; sketch it on paper; the next slide shows the real screenshot | Three boxes in a bar | Where does the sidebar go? (the original deck's Practical Example, which really does put the sidebar on its own row) |
| **Spot the bug** | Code that silently fails; pick the reason | `flex-wrap` on the items | `grid-area: content` with no area called content (the answer shows what the browser does) |
| **Playground** | The class calls out values, you drive `playground/*.html` (four challenges on the slide) | Flexbox | Grid |
| **Game** | Free browser games, 10 min (or homework) | Flexbox Froggy | Grid Garden |
| **Sorting game** | Teams pick Grid, Flexbox or Float for six jobs | | ✓ |
| **Pause and think** | Why did the Kids card stretch across its own row? (flexbox sizes each line on its own) | | After Try It 5 |
| **Exit ticket** | Six quick-fire "write the CSS" cards, answers on the next slide | ✓ | ✓ |

Quiz answers and common wrong answers are in the speaker notes.

## Class 1: Flexbox (`Flexbox.pptx`, 42 slides, about 110 min)

| Part | Topic | Example | Interactive / Try It | Time |
|---|---|---|---|---|
| — | Title, check-in, overview, get the code, *CSS Flexbox in 100 Seconds* | | | 10 min |
| F0 | Key ideas: container, items, main axis, cross axis | | Pop Quiz 1 (where does `display: flex` go?) | 5 min |
| F1 | `display: flex` | 01 | | 3 min |
| F2 | `flex-direction` | 02 | Pop Quiz 2 (column = top to bottom) | 5 min |
| F3–F4 | `justify-content`, `align-items` | 03, 04 | | 8 min |
| F5 | Centre anything | 05 | Pop Quiz 3, Predict | 8 min |
| — | Flexbox Playground, Harbourside intro | | **Try It 1** The header | 15 min |
| F6 | `flex-wrap` and `gap` | 06 | Spot the bug | 6 min |
| F7 | `flex: 1`, `flex: 2`, a fixed sidebar, `order` | 07 | Pop Quiz 4 (flex 1 / 2 / 1 in 800px) | 8 min |
| F8 | A header: `margin-left: auto` | 08 | Pop Quiz 5 (what does `margin-top: auto` do?) | 5 min |
| F9 | Cards: wrap + `flex: 1 1 200px`, column cards, button at the bottom | 09 | **Try Its 2–4** | 20 min |
| — | Flexbox Froggy, W3Schools work time, cheat sheet, exit ticket, recap | | | 30 min |

## Class 2: Grid and Float (`Grid_and_Float.pptx`, 47 slides, about 115 min)

| Part | Topic | Example | Interactive / Try It | Time |
|---|---|---|---|---|
| — | Title, check-in, overview, warm-up quiz, *CSS Grid in 100 Seconds* | | Warm-up (flexbox is one-dimensional) | 10 min |
| G0 | Key ideas: lines, tracks, cells, areas | | | 3 min |
| G1 | `grid-template-columns`, `fr`, `repeat()` | 10 | Pop Quiz 1 (7 items, 3 columns: rows?) | 6 min |
| G2 | Spanning: `span 2`, `2 / 4`, `1 / -1` | 11 | Pop Quiz 2 (`1 / 3` = two columns) | 6 min |
| G3 | Named areas | 12, 13 | Predict (12), spot the bug | 12 min |
| — | | | **Try It 5** Page layout, then Pause and think | 10 min |
| G4 | `repeat(auto-fit, minmax(…))` | 14 | Pop Quiz 3, **Try It 6** | 10 min |
| F1–F2 | Float: text wrap, clearfix, `display: flow-root` | 15, 16 | Pop Quiz 4, **Try It 7** | 15 min |
| — | Grid, Flexbox or Float? | | Sorting game | 10 min |
| L1 | A responsive layout: mobile first + `@media` | 17 | **Try It 8** On a phone | 15 min |
| — | Best practices, Grid Playground, Grid Garden, W3Schools work time, submit | | | 15 min |
| — | **Assignment 3** (due Sun, Nov 1): each rubric row mapped to this week's examples, then cheat sheet, exit ticket, recap | | | 8 min |

## Harbourside, Part 3

`activity/index.html` is the Week 7 answer key (nav, dropdown, hours table) with three new wrappers: `div.brand` (logo + h1), `section.intro`, and `div.cards` around the menu articles. **Students only edit `style.css` this week.** The starter `style.css` has the Class 8 and Class 10 CSS tidied up at the top, then an empty heading for each Try It.

| # | Try It | Class | New idea |
|---|---|---|---|
| 1 | The header | 1 | `.brand` and `header` as flex containers, `space-between`, `align-items: center` |
| 2 | Menu cards side by side | 1 | `flex-wrap`, `gap`, `flex: 1 1 180px` |
| 3 | Line up the prices | 1 | Each card a column flexbox, `margin-top: auto` |
| 4 | The footer, and today's special first | 1 | `space-between`, `order: -1` |
| 5 | A page layout with grid | 2 | `grid-template-areas` on `main`, `align-items: start` |
| 6 | Menu cards with grid | 2 | `repeat(auto-fit, minmax(180px, 1fr))` replaces Try It 2's flex |
| 7 | Float the photo | 2 | `float: right` on the intro figure |
| 8 | On a phone | 2 | `@media (max-width: 700px)`: one column, no float, smaller h1, wrapping nav |

**Planned surprise (Try It 5):** the menu column gets narrower, so only three cards fit per line and the Kids card stretches across a line of its own. The "Pause and think" slide explains why (flexbox sizes each line separately), and Try It 6 fixes it with grid.

**Things to know:**
- Try It 7 needs no clearfix: `.intro` is a grid item, and grid items always contain their floats.
- Try It 8 must be **last** in the file: media queries add no specificity, so a later rule would win.
- `order` (Try It 4) still works after Try It 6, because grid items have an order too. Try It 2's `flex: 1 1 180px` is then ignored.

**Submit (participation, after Class 2):** Harbourside at full width, Harbourside at phone width (DevTools, Ctrl + Shift + M), and at least 5 W3Schools "Try it Yourself" screenshots with the student's own changes (flexbox and grid together).

## Changes from the original Week 8 deck
- Split over two classes; every example is a separate HTML page + external stylesheet.
- `grid-gap` → `gap` (the notes mention the old name).
- The Practical Example is kept word for word as a "predict" slide. Its auto-placement puts the sidebar on its own row, which leads into named areas (Example 13).
- The clearfix slide adds `display: flow-root`, the modern one-line fix.
- `order` comes with an accessibility note: it changes the look, not the reading or Tab order.
- The "Quick check-in" slide (Outlook, Teams, Brightspace) is now the attendance slide in both decks.
- The two Fireship videos open each class; Flexbox Froggy and Grid Garden are added as games.
- Assignment 3 (10%, due Sunday, November 1): a reminder in both recaps, and a Class 2 slide that maps each rubric row (HTML Structure, CSS Usage, Responsiveness, Code Quality) to this week's examples.

## Folder

```
08_Flex_Grid_CSS/
├── README.md
├── 01_flex_container_example.html / .css
├── 02_flex_direction_example.html / .css
├── 03_justify_content_example.html / .css
├── 04_align_items_example.html / .css
├── 05_centering_example.html / .css
├── 06_flex_wrap_example.html / .css
├── 07_flex_grow_example.html / .css        ← flex: 1 / 2, fixed sidebar, order
├── 08_flex_navbar_example.html / .css      ← margin-left: auto
├── 09_flex_cards_example.html / .css       ← column cards, button at the bottom
├── 10_grid_columns_example.html / .css
├── 11_grid_span_example.html / .css
├── 12_grid_layout_example.html / .css      ← the original deck's Practical Example (sidebar surprise)
├── 13_grid_areas_example.html / .css       ← same HTML, fixed with named areas
├── 14_grid_auto_fit_example.html / .css
├── 15_float_text_wrap_example.html / .css
├── 16_float_clearfix_example.html / .css   ← collapse, clearfix, flow-root
├── 17_responsive_layout_example.html / .css
├── images/harbour.jpg
├── playground/                             ← flex_playground.html, grid_playground.html, playground.css, playground.js
├── activity/                               ← index.html, style.css, images/, answer/
├── Flexbox.pptx                            ← Week 8, Class 1
├── Grid_and_Float.pptx                     ← Week 8, Class 2
└── _build/                                 ← build_deck.js (screenshots go in _build/shots, git-ignored)
```

## Colours (WCAG AA)

The Class 8 palette: navy `#12304A`, teal `#1B5E82` (white text on both), and dark orange `#B5451B` for highlighted boxes (white text, 5.5 : 1). The bright orange `#E97132` is only used for borders and underlines (white or navy text on it is under 4.5 : 1). The header button in Example 08 is `#F4A06B` with navy text (6.5 : 1).
