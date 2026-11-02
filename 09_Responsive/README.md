# Week 9, Class 2: Responsive Web Design

> This README is the source for:
> 1. `09_Responsive/*.html` + `*.css`: eight numbered examples. **Each example is two files:** an HTML page and its own external stylesheet with the same name
> 2. `Responsive_Web_Design.pptx` (43 slides, same style as the Week 8 and Week 9, Class 1 decks; every slide has speaker notes)
>
> Rebuild the deck: `cd _build`, `npm install` once, then `npm run build`. The build screenshots the examples with Microsoft Edge (headless). Code on the slides is read from the example files, so edit those files and rebuild.

## The class (about two hours)

| Slides | Topic | Example | Interactive / Try It | Time |
|---|---|---|---|---|
| 1–4 | Title, *Is your environment set up?*, *Recall*, overview | | | 10 min |
| 5–9 | What it is, why it matters, a page that doesn't fit | 01 | Two words to define | 10 min |
| 10–12 | The viewport meta tag | 02 | Pop Quiz 1 | 5 min |
| 13–18 | The video, reading a media query, a first media query | 03 | Pop Quiz 2 (`max-width`) | 15 min |
| 19–22 | Breakpoints, `max-width` and `min-width`, phones first | 04 | Pop Quiz 3 (the last rule wins) | 10 min |
| 23–27 | Fluid layouts, the padding trap | 05 | Pop Quiz 4 (360 + 30 = 390px) | 15 min |
| 28–30 | Responsive images and video | 06 | | 10 min |
| 31–34 | Fluid type: `vw`, the jump at the breakpoint, `clamp()` | 07 | Pop Quiz 5 (46px to 23px) | 10 min |
| 35–37 | Putting it together | 08 | | 5 min |
| 38–41 | Make Example 01 responsive, then submit | 01 → 08 | **Try Its 1–3**, class activity | 25 min |
| 42–43 | Recap, *write a one-sentence summary* | | | 5 min |

**The activity:** students edit `01_fixed_layout_example.css` until it works on a phone. Examples 01 and 08 have the same HTML, so Example 08 is the answer key. Anyone who'd rather use their own page can.

**Submit (Slide 41):** two screenshots, wide and at phone width, in the box on Week 9 Brightspace.

## Changes from the original deck

- One class instead of two, with no Class 12 slide, no duplicate overview, and no dates.
- **Adds the viewport meta tag.** Without it, phones don't match the media queries at all.
- **`width: 100%` plus `padding`** (the original's breakpoint and practical example slides) overflows by the padding, which causes sideways scrolling. It's now Pop Quiz 4, and the fix is `box-sizing: border-box`.
- **`6vw` up to 768px, then `3vw`** (the original's font sizing slide) makes the heading drop from 46px to 23px at 769px. It's now Pop Quiz 5, and `clamp()` is the fix. Plain `vw` also doesn't grow when you zoom the browser (WCAG 1.4.4).
- **Video:** `aspect-ratio` comes first. The original's `padding-bottom: 56.25%` trick stays in Example 06, as what students will see in older code.
- **One set of breakpoints.** The original had 320/768/1024 on one slide and 480/768 on another.
- **Adds the order of media queries**, and `min-width` (phones first) next to `max-width`.
- The search slide says Google uses the mobile version of a site, instead of "favors mobile-friendly websites".

## Folder

```
09_Responsive/
├── README.md
├── 01_fixed_layout_example.html / .css     ← the problem; students fix this one
├── 02_viewport_example.html / .css         ← comment out the meta tag to see the difference
├── 03_media_query_example.html / .css
├── 04_breakpoints_example.html / .css      ← phones first, with min-width
├── 05_fluid_layout_example.html / .css     ← percentages, and the padding trap
├── 06_responsive_media_example.html / .css ← images and video (poster only, no video file)
├── 07_fluid_type_example.html / .css       ← vw, the jump, clamp()
├── 08_responsive_page_example.html / .css  ← Example 01, made responsive
├── images/                                 ← harbour.jpg
├── Responsive_Web_Design.pptx
└── _build/                                 ← build_deck.js (screenshots go in _build/shots, git-ignored)
```
