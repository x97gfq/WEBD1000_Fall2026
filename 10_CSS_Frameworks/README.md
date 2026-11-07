# Week 10: CSS Frameworks (Bootstrap, and a look at Tailwind)

> This README is the source for:
> 1. `10_CSS_Frameworks/*.html` (+ `.css` where an example needs its own): nine numbered examples. Bootstrap and Tailwind load from a CDN, so **the examples need the internet**.
> 2. `activity/`: Harbourside, Part 4. The Week 7 page with no CSS, rebuilt with Bootstrap in Try Its 2 to 7. `activity/answer/` is the key. `activity/try_it_1_answer.html` is Class 1's colour grid.
> 3. `CSS_Frameworks_Bootstrap.pptx` (71 slides, two PowerPoint sections, one per class; same style as the Week 8 and 9 decks; every slide has speaker notes).
>
> Rebuild the deck: `cd _build`, `npm install` once, then `npm run build`. The build screenshots the examples with Microsoft Edge (headless). Code on the slides is read from the example files and from `activity/answer/`, so edit those files and rebuild.
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

## Class 2: Components, Tailwind and Harbourside (slides 37–71)

| Slides | Topic | Example | Interactive / Try It |
|---|---|---|---|
| 37–40 | Divider, environment, *Recall*, overview | | |
| 41–47 | Buttons, alerts, cards, the navbar, show/hide classes | 05, 06 | Pop Quiz 7 (A: `data-bs-toggle`) |
| 48–52 | Customizing: your CSS after theirs, button variables; where to look things up | 07 | Pop Quiz 8 (B: hover goes blue) |
| 53–60 | Tailwind: the same cards two ways, reading classes, which to use | 08, 09 | Pop Quizzes 9 (D) and 10 (A) |
| 61–69 | Harbourside in Bootstrap: before and after, then submit | activity | **Try Its 2–7** |
| 70–71 | Week recap, *write a one-sentence summary* | | |

**The activity** (replaces last year's "download a Bootstrap example and customize it", which is now a fast-finisher option): students open `activity/index.html` (no CSS) and rebuild it with Bootstrap: links and container (2), navbar (3), banner and two-column intro (4), menu cards (5), table and footer (6), Harbourside colours in `style.css` (7). The comments in `index.html` say where each Try It goes.

**Submit (slide 69):** two screenshots, wide and at phone width, in the box on Week 10 Brightspace.

**Brightspace quiz:** 10 questions, one per pop quiz, in `../temp/WEBD1000_QUIZ_10_CSS_Frameworks_Bootstrap.csv` (git-ignored, so students don't get it with `git pull`).

## Changes from the original deck

- **Bootstrap 5.3.8 from jsDelivr**, not Bootstrap 4.5.2 from `stackpath.bootstrapcdn.com`. Bootstrap 6 is in alpha; the deck says so.
- **The navbar** is Bootstrap 5 code: `data-bs-toggle` / `data-bs-target` (Bootstrap 4's `data-toggle` does nothing in 5; now Pop Quiz 7), `container` inside the navbar, `ms-auto`.
- **"Customizing Bootstrap"**: the slide's `.btn-primary { background-color: #ff5722; }` goes back to blue on hover, because Bootstrap's `.btn:hover` is more specific. It's now Pop Quiz 8, and the fix (the `--bs-btn-*` variables) is Example 07.
- **The colour slides** used Bootstrap 4.1 docs. `text-muted` and `text-black-50` are deprecated in 5.3 (use `text-body-secondary`), and `text-bg-*` replaces `bg-*` with `text-white`, so text on yellow stays readable. Try It 1 uses it.
- **"Wells" and "Glyphicons"** were Bootstrap 3 and are gone: cards and Bootstrap Icons.
- **Spacing**: `ml-`/`mr-` (Bootstrap 4) are `ms-`/`me-` in 5.
- "Each row has 12 cells" is now 12 columns. The "together" activity asked for equal columns on medium and larger screens with `col-sm-4` on an earlier slide; it's `col-md-4` throughout now.
- Adds what a framework **costs** (file size: about 230 KB of CSS, 31 KB compressed; the "Bootstrap look"), and a DevTools slide showing `.col-md-4` is a min-width media query on a flexbox.
- **Tailwind** (Examples 08 and 09, the same cards two ways) replaces Foundation and Bulma as the second demo. They stay on the comparison table.
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
├── 07_customizing_example.html / .css        ← hover each button: the old way vs the variables
├── 08_bootstrap_cards_example.html           ← the same three cards...
├── 09_tailwind_cards_example.html            ← ...in Tailwind (Play CDN)
├── activity/       ← Harbourside, Part 4 (index.html + style.css to do; answer/ is the key)
├── CSS_Frameworks_Bootstrap.pptx
└── _build/         ← build_deck.js (screenshots go in _build/shots, git-ignored)
```
