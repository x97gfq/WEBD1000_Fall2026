# Week 9, Class 1: Float

> This README is the source for `Float.pptx` (Week 9, Class 1; same style as the Week 8 decks; every slide has speaker notes).
>
> Rebuild the deck: `cd _build`, `npm install` once, then `npm run build`. The build screenshots Examples 15 and 16 with Microsoft Edge (headless). Code on the slides is read from the example `.css` files and from `activity/answer/style.css`, so edit those files and rebuild.

Carried forward from Week 8, Class 2 (`Grid_and_Float.pptx`, slides 27 to 32), because class ran out of time before float. The examples and a copy of the Harbourside page are in this folder, with Try Its 1 to 6 already done, so students can start Try It 7 here. `08_Flex_Grid_CSS/activity` keeps all eight Try Its for Week 8.

## The deck (`Float.pptx`, 10 slides, about 15 min of float)

| Slide | Topic | Example | Interactive / Try It |
|---|---|---|---|
| 1–3 | Title, check-in, overview | | |
| 4 | Float divider | | |
| 5 | Wrapping text around a photo | 15 | |
| 6 | Floats and the clearfix, `display: flow-root` | 16 | |
| 7–8 | Why does the box collapse? | | Pop Quiz 1 (Pop Quiz 4 in the Week 8 deck) |
| 9 | Float the photo | Harbourside | **Try It 7** |
| 10 | Float recap | | |

**Things to know:**
- Try It 7 needs Try It 5 on the page: `.intro` is a grid item, and grid items always contain their floats. It's already in `activity/style.css`, along with Try Its 1 to 4 and 6.

## Not carried forward

Still only in `Grid_and_Float.pptx`: Grid, Flexbox or Float? and the sorting game (slides 33–35), the responsive layout and Try It 8 (36–37), best practices, the Grid Playground, Grid Garden, W3Schools work time, submit, Assignment 3, the cheat sheet and the exit ticket (38–47).

## Folder

```
09_Float_CSS/
├── README.md
├── 15_float_text_wrap_example.html / .css
├── 16_float_clearfix_example.html / .css
├── images/
├── activity/       ← Harbourside: Try Its 1 to 6 done, Try It 7 to do (answer/ is the key)
├── Float.pptx      ← Week 9, Class 1
└── _build/         ← build_deck.js (screenshots go in _build/shots, git-ignored)
```
