# Week 11: Making Bootstrap Yours (customizing, its JavaScript components, Tailwind, Harbourside)

> This README is the source for:
> 1. `11_Customizing_Bootstrap_Tailwind/*.html` (+ `.css`): three numbered examples (customizing, and the same cards in Bootstrap and Tailwind).
> 2. `js/`: eight more example pages, one per JavaScript-powered component, each one openable on its own. Class 1's "try it" material.
> 3. `activity/`: Harbourside, Part 4. The Week 7 page with no CSS, rebuilt with Bootstrap in Try Its 1 to 6. `activity/answer/` is the key.
> 4. `WEBD1000_Quiz_Week_11.md`: the 10-question Brightspace quiz. One question per pop quiz. `WEBD1000_Quiz_Week_11.csv` is the same quiz for Brightspace's question import (Quizzes → Question Library → Import → Upload a File). The `.md` is the source: after editing it, run `python _build/quiz_to_csv.py WEBD1000_Quiz_Week_11.md WEBD1000_Quiz_Week_11.csv`. Code in the questions is flattened onto one line after "Code:", and the csv module does the quoting, so a `"` in the JavaScript comes out as `""`.
> 5. `Customizing_Bootstrap_Tailwind.pptx` (67 slides, two PowerPoint sections, one per class; same style as the Week 10 deck; every slide has speaker notes).
>
> Bootstrap and Tailwind load from a CDN, so **every example needs the internet**.
>
> Rebuild the deck: `cd _build`, `npm install` once, then `npm run build`. The build screenshots the examples with Microsoft Edge (headless). Code on the slides is read from the example files where it fits on a slide, and abridged by hand where it doesn't (the modal, accordion and carousel HTML), so check the slide as well as the file when you change one.
>
> No Node.js on this PC? VS Code's own Node works. In PowerShell, from `_build`: `$env:ELECTRON_RUN_AS_NODE = "1"; & "C:\Program Files\Microsoft VS Code\Code.exe" build_deck.js` (run it with `Start-Process -Wait` if you want to know when it has finished; it prints the path it wrote).

## Class 1: Customizing, and the Parts That Move (slides 1–39)

| Slides | Topic | Example | Interactive / Try It |
|---|---|---|---|
| 1–5 | Title, check-in (incl. Assignment 4), *Is your environment set up?*, *Recall* (last week's layout), overview | | |
| 6–9 | Customizing: your CSS after theirs, button variables | 01 | Pop Quiz 1 (B: hover goes blue) |
| 10–14 | Divider, **"Java is to JavaScript as ham is to hamster"**, two different languages, Mocha → LiveScript → JavaScript | | Pop Quiz 2 (C, trivia) |
| 15–16 | How much JavaScript today needs (none / one line / three lines), and which components are worth their time | | |
| 17–20 | A modal | `js/01` | **Try It 1**, Pop Quiz 3 (A) |
| 21–22 | Collapse, and an accordion | `js/02` | **Try It 2** |
| 23–24 | A carousel | `js/03` | **Try It 3** |
| 25–27 | Offcanvas, and the `data-bs-toggle` / `data-bs-target` pattern behind all of it | `js/04`, `js/00` | Pop Quiz 4 (B) |
| 28–29 | Divider, **reading a line of JavaScript** word by word | | |
| 30–33 | Tooltips: opt in, and the one line that turns them on | `js/05` | **Try It 4**, Pop Quiz 5 (C) |
| 34–37 | A toast: when this, do that | `js/06` | **Try It 5**, Pop Quiz 6 (C) |
| 38–39 | Everything on one page, and the Class 1 recap | `js/07` | Fast finishers |

About 100 minutes. **If you are behind:** cut Try It 3 (the carousel, slide 24) and the offcanvas
slide (25), and make slide 38 a link rather than a demo. Don't cut the toast — it is the one that
makes JavaScript click for people.

## Class 2: Tailwind, and Harbourside in Bootstrap (slides 40–67)

| Slides | Topic | Example | Interactive / Try It |
|---|---|---|---|
| 40–43 | Divider, environment, *Recall* (last class's pattern), overview | | |
| 44–51 | Tailwind: the same cards two ways, reading classes, which to use | 02, 03 | Pop Quizzes 7 (D) and 8 (A) |
| 52–53 | Harbourside: before and after | activity | |
| 54–61 | Rebuild it: links and container, navbar, banner and two columns, menu cards, table and footer, then your own colours | activity | **Try Its 1–6**, Pop Quiz 9 (C) |
| 62–64 | Pop Quiz 10 (B), then submit | | |
| 65–67 | Resources, week recap, *write a one-sentence summary* | | |

About 95 minutes, then any time left is Assignment 4 work time.

**Why this order:** Class 1 is all Bootstrap, going deeper — first make it look like yours
(customizing), then make it *do* something (the JavaScript components). Class 2 is the contrast
(Tailwind) and the build (Harbourside). The Harbourside navbar in Class 2's Try It 2 is a collapse
plus a dropdown, so Class 1 is the groundwork for it: by then students have already seen
`data-bs-toggle` five times and know that a missing script breaks everything silently.

## Easing into JavaScript

Students have **not** learned JavaScript, and the deck says so out loud more than once. Class 1
climbs three steps, and slide 15 is the map:

| Level | What the student writes | Components |
|---|---|---|
| None at all | Two HTML attributes | Modal, collapse, accordion, carousel, offcanvas, dropdown, navbar toggle |
| One line | `new bootstrap.Tooltip(el)` | Tooltip, popover |
| Three lines | `button.addEventListener("click", … )` | Toast, and everything from next week on |

Slide 29 pulls one line apart (`document` → `.getElementById()` → `.addEventListener()` →
`function () { … }` → `toast.show()`), and the phrase to repeat all class is **"when this, do
that"** — which is the same shape as next week's "when the form is submitted, check the email
address".

**Which components, and why these:** modal, collapse/accordion, carousel and offcanvas are the four
that need no JavaScript at all and cover most of what a small business site needs; the accordion is
the easiest five-minute win for Assignment 4. Tooltip and toast come last because they are the
gentlest possible introduction to actually writing JavaScript — one line, then three. Dropdown,
popover, tabs, alerts and scrollspy are on the component slide as "look them up": once you can read
one, you can read all of them.

## The activity

Students open `activity/index.html` (no CSS) and rebuild it with Bootstrap: links and container (1),
navbar (2), banner and two-column intro (3), menu cards (4), table and footer (5), Harbourside
colours in `style.css` (6). The comments in `index.html` say where each Try It goes.

**Submit (slide 64):** two screenshots, wide and at phone width, in the box on Week 11 Brightspace.
One interactive component from Class 1 is an optional extra.

## Brightspace quiz

**10 questions**, one per pop quiz, in `WEBD1000_Quiz_Week_11.md`. Six come from Class 1 and four
from Class 2. Tell students at Pop Quiz 1 that every pop quiz this week is on the quiz, so they
write the answers down.

## Changes from the original deck

- **"Customizing Bootstrap"**: the old slide's `.btn-primary { background-color: #ff5722; }` goes back to blue on hover, because Bootstrap's `.btn:hover` is more specific. It's now Pop Quiz 1, and the fix (the `--bs-btn-*` variables) is Example 01.
- **Tailwind** (Examples 02 and 03, the same cards two ways) replaces Foundation and Bulma as the second demo.
- **Bootstrap's JavaScript components** are new this year, and are most of Class 1. Week 10 ends with a slide listing the components that move; this week picks six of them up and builds them.
- **Pop Quizzes 1 and 10 are a deliberate pair**: the first is a case where specificity beats source order, the last is a case where the rules tie and order decides. Students who are confused by the second one are confused in exactly the right way.

## Folder

```
11_Customizing_Bootstrap_Tailwind/
├── README.md
├── WEBD1000_Quiz_Week_11.md             ← the 10-question Brightspace quiz (the source)
├── WEBD1000_Quiz_Week_11.csv            ← ...as a Brightspace question-import CSV (_build/quiz_to_csv.py)
├── 01_customizing_example.html / .css   ← hover each button: the old way vs the variables
├── 02_bootstrap_cards_example.html      ← the same three cards...
├── 03_tailwind_cards_example.html       ← ...in Tailwind (Play CDN)
├── js/                                  ← Class 1: the components that move
│   ├── 00_no_script_example.html        ← broken on purpose: the bundle script deleted
│   ├── 01_modal_example.html            ← a modal, no JavaScript of our own
│   ├── 02_accordion_example.html        ← one collapse, then an accordion (data-bs-parent)
│   ├── 03_carousel_example.html         ← a slideshow; interval, autoplay, exactly one "active"
│   ├── 04_offcanvas_example.html        ← a side panel, and the pattern behind all of them
│   ├── 05_tooltip_example.html          ← opt in: the first line of JavaScript
│   ├── 06_toast_example.html            ← when this, do that: the first event listener
│   ├── 07_all_together_example.html     ← all six on one Harbourside page, five lines of JS
│   └── images/                          ← slide1–3.svg, the carousel's pictures
├── activity/       ← Harbourside, Part 4 (index.html + style.css to do; answer/ is the key)
├── Customizing_Bootstrap_Tailwind.pptx
└── _build/         ← build_deck.js (screenshots go in _build/shots, git-ignored)
```
