# Week 11: Making Bootstrap Yours (customizing, Tailwind, Harbourside)

> This README is the source for:
> 1. `11_Customizing_Bootstrap_Tailwind/*.html` (+ `.css`): three numbered examples. Bootstrap and Tailwind load from a CDN, so **the examples need the internet**.
> 2. `activity/`: Harbourside, Part 4. The Week 7 page with no CSS, rebuilt with Bootstrap in Try Its 1 to 6. `activity/answer/` is the key.
> 3. `Customizing_Bootstrap_Tailwind.pptx` (29 slides, one PowerPoint section for Class 1; same style as the Week 10 deck; every slide has speaker notes).
>
> These moved here from Week 10's Class 2 (`../10_CSS_Frameworks/`), which now ends with the component library. Examples 07, 08 and 09 are now 01, 02 and 03; Try Its 2 to 7 are now 1 to 6; Pop Quizzes 8 to 10 are now 1 to 3.
>
> Rebuild the deck: `cd _build`, `npm install` once, then `npm run build`. The build screenshots the examples with Microsoft Edge (headless). Code on the slides is read from the example files and from `activity/answer/`, so edit those files and rebuild. The build's helpers are a copy of Week 10's.
>
> No Node.js on this PC? VS Code's own Node works. In PowerShell, from `_build`: `$env:ELECTRON_RUN_AS_NODE = "1"; & "C:\Program Files\Microsoft VS Code\Code.exe" build_deck.js` (it prints nothing; check the .pptx's time).

## Class 1: Customizing Bootstrap, Tailwind and Harbourside (slides 1–29)

| Slides | Topic | Example | Interactive / Try It |
|---|---|---|---|
| 1–5 | Title, check-in (incl. Assignment 4), *Is your environment set up?*, *Recall* (last week's layout), overview | | |
| 6–9 | Customizing: your CSS after theirs, button variables | 01 | Pop Quiz 1 (B: hover goes blue) |
| 10–17 | Tailwind: the same cards two ways, reading classes, which to use | 02, 03 | Pop Quizzes 2 (D) and 3 (A) |
| 18–26 | Harbourside in Bootstrap: before and after, then submit | activity | **Try Its 1–6** |
| 27 | Resources: docs, examples, W3Schools, Tailwind; videos (thenewboston, Web Dev Simplified) | | |
| 28–29 | Week recap, *write a one-sentence summary* | | |

About 95 minutes. Class 2 of Week 11 isn't in this deck: it's free for Assignment 4 work time, or anything else.

**Why this order:** customizing first (Example 01's button variables), then Tailwind as the other answer to "every Bootstrap site looks the same", then Harbourside, whose last Try It (6) puts the customizing into practice. That's also Assignment 4's "your own styling" requirement.

**The activity** (replaces last year's "download a Bootstrap example and customize it", which is now a fast-finisher option): students open `activity/index.html` (no CSS) and rebuild it with Bootstrap: links and container (1), navbar (2), banner and two-column intro (3), menu cards (4), table and footer (5), Harbourside colours in `style.css` (6). The comments in `index.html` say where each Try It goes.

**Submit (slide 26):** two screenshots, wide and at phone width, in the box on Week 11 Brightspace.

**Brightspace quiz:** 3 questions, one per pop quiz.

## Changes from the original deck

- **"Customizing Bootstrap"**: the old slide's `.btn-primary { background-color: #ff5722; }` goes back to blue on hover, because Bootstrap's `.btn:hover` is more specific. It's now Pop Quiz 1, and the fix (the `--bs-btn-*` variables) is Example 01.
- **Tailwind** (Examples 02 and 03, the same cards two ways) replaces Foundation and Bulma as the second demo.

## Folder

```
11_Customizing_Bootstrap_Tailwind/
├── README.md
├── 01_customizing_example.html / .css   ← hover each button: the old way vs the variables
├── 02_bootstrap_cards_example.html      ← the same three cards...
├── 03_tailwind_cards_example.html       ← ...in Tailwind (Play CDN)
├── activity/       ← Harbourside, Part 4 (index.html + style.css to do; answer/ is the key)
├── Customizing_Bootstrap_Tailwind.pptx
└── _build/         ← build_deck.js (screenshots go in _build/shots, git-ignored)
```
