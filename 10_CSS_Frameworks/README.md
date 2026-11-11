# Week 10: CSS Frameworks (Bootstrap)

> This README is the source for:
> 1. `10_CSS_Frameworks/*.html` (+ `.css` where an example needs its own): six numbered examples. Bootstrap loads from a CDN, so **the examples need the internet**. `try_it_1_answer.html` is Class 1's colour grid.
> 2. `layout/`: Your Layout, Class 2's warm-up and its two follow-ups. A wireframe (`00_wireframe.svg`) becomes a Bootstrap grid (`01_start.html` → `01_layout_answer.html`), then gets a navbar (`02_navbar_start.html` → `02_navbar_answer.html`) and cards (`03_cards_answer.html`). Bootstrap classes only, no CSS of our own.
> 3. `CSS_Frameworks_Bootstrap.pptx` (60 slides, two PowerPoint sections, one per class; same style as the Week 8 and 9 decks; every slide has speaker notes).
>
> Customizing Bootstrap, Tailwind and the Harbourside activity moved to Week 11: `../11_Customizing_Bootstrap_Tailwind/`.
>
> Rebuild the deck: `cd _build`, `npm install` once, then `npm run build`. The build screenshots the examples with Microsoft Edge (headless). Code on the slides is read from the example files and from `layout/`, so edit those files and rebuild.
>
> No Node.js on this PC? VS Code's own Node works. In PowerShell, from `_build`: `$env:ELECTRON_RUN_AS_NODE = "1"; & "C:\Program Files\Microsoft VS Code\Code.exe" build_deck.js` (it prints nothing; check the .pptx's time).

## Class 1: CSS Frameworks and the Bootstrap Grid (slides 1–36)

| Slides | Topic | Example | Interactive / Try It |
|---|---|---|---|
| 1–6 | Title, check-in (email, Teams, Brightspace), *Is your environment set up?*, *Recall*, overview | | Warm-up: what does "bootstrap" mean? |
| 7–13 | What a framework is, what it costs, the two kinds, the popular ones | | Pop Quiz 1 (D) |
| 14–19 | What Bootstrap is, linking it, CDN or download | 01 | Pop Quiz 2 (C) |
| 20–23 | Container, row, 12 columns | 02 | Pop Quiz 3 (B: 6 + 4 + 4 wraps) |
| 24–27 | Breakpoints, a responsive grid | 03 | Pop Quiz 4 (A: `col-md-4` at 600px) |
| 28–30 | Look inside a Bootstrap class with DevTools | | Pop Quiz 5 (C: flexbox) |
| 31–34 | Colour and spacing utilities | 04 | Pop Quiz 6 (D: `px-3 mb-0`) |
| 35–36 | A colour grid, together; recap | | **Try It 1** |

## Class 2: Components and Your Layout (slides 37–60)

| Slides | Topic | Example | Interactive / Try It |
|---|---|---|---|
| 37–39 | Divider, environment, *Recall* | | |
| 40–42 | Warm-up: count the wireframe's columns, build it together, the wireframe beside the result | layout 00, 01 | **Warm-up** |
| 43 | Overview | | |
| 44–49 | Buttons, alerts, cards, the navbar | 05, 06 | Pop Quiz 7 (A: `data-bs-toggle`) |
| 50–53 | Your layout: add a navbar, add cards, and what each should look like | layout 02, 03 | **Your Layout, Parts 2 and 3** |
| 54–55 | Bootstrap's 24 components, each linked to its docs page (and W3Schools' where there is one) | | Fast finishers: add one |
| 56–57 | Show/hide classes; where to look things up | | |
| 58 | Resources: docs, examples, W3Schools, Tailwind; videos (thenewboston, Web Dev Simplified) | | |
| 59–60 | Week recap, *write a one-sentence summary* | | |

**Your Layout** (added after Class 1 ran a bit dry): Class 2 opens with a warm-up. Students count the columns on a wireframe, then build it from `layout/01_start.html` (a blank page with the viewport tag and Bootstrap's links): a container and three rows (header `col-12`; `aside` `col-md-3` + `main` `col-md-9`; footer `col-12`), each part in a loud `text-bg-*` colour so the layout is obvious. After the navbar example and Pop Quiz 7, they add a navbar to it (Part 2; `02_navbar_start.html` is the finished warm-up), then two cards in a row inside `main` (Part 3). About 85 minutes in all.

**Brightspace quiz:** 7 questions, one per pop quiz (1 to 7). Pop Quizzes 8 to 10 moved to Week 11 as its Pop Quizzes 1 to 3.

## Changes from the original deck

- **Bootstrap 5.3.8 from jsDelivr**, not Bootstrap 4.5.2 from `stackpath.bootstrapcdn.com`. Bootstrap 6 is in alpha; the deck says so.
- **The navbar** is Bootstrap 5 code: `data-bs-toggle` / `data-bs-target` (Bootstrap 4's `data-toggle` does nothing in 5; now Pop Quiz 7), `container` inside the navbar, `ms-auto`.
- **The colour slides** used Bootstrap 4.1 docs. `text-muted` and `text-black-50` are deprecated in 5.3 (use `text-body-secondary`), and `text-bg-*` replaces `bg-*` with `text-white`, so text on yellow stays readable. Try It 1 uses it.
- **"Wells" and "Glyphicons"** were Bootstrap 3 and are gone: cards and Bootstrap Icons.
- **Spacing**: `ml-`/`mr-` (Bootstrap 4) are `ms-`/`me-` in 5.
- "Each row has 12 cells" is now 12 columns. The "together" activity asked for equal columns on medium and larger screens with `col-sm-4` on an earlier slide; it's `col-md-4` throughout now.
- Adds what a framework **costs** (file size: about 230 KB of CSS, 31 KB compressed; the "Bootstrap look"), and a DevTools slide showing `.col-md-4` is a min-width media query on a flexbox.
- The original Class 14 (browser DevTools: JavaScript debugging, Network, Performance) isn't in this deck. Students have used the Elements panel and device toolbar since Week 9; the DevTools slide and the notes (Network tab on Example 01) cover what's relevant here.

## Folder

```
10_CSS_Frameworks/
├── README.md
├── 01_bootstrap_starter_example.html / .css  ← the two links and the script
├── 02_grid_example.html / .css               ← 12 columns; the last row adds up to 14
├── 03_responsive_grid_example.html / .css    ← col-12 col-md-6 col-lg-3, and a breakpoint badge
├── 04_utilities_example.html                 ← colours and spacing, no CSS of our own
├── 05_components_example.html                ← buttons, an alert, cards
├── 06_navbar_example.html                    ← folds into a menu button under 768px
├── try_it_1_answer.html                      ← Class 1's colour grid
├── layout/         ← Your Layout: wireframe, blank start, warm-up answer, navbar start/answer, cards answer
├── CSS_Frameworks_Bootstrap.pptx
└── _build/         ← build_deck.js (screenshots go in _build/shots, git-ignored)
```
