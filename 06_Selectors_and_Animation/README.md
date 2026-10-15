# Class 9: Advanced Selectors and CSS Animation

> This README is the source for:
> 1. `06_Selectors_and_Animation/*.html`: twelve numbered code samples, each with its CSS in a `<style>` block (01–05 selectors, 06–12 animation)
> 2. `activity/`: the Animation Playground (four challenges and a bonus), plus `activity/answer/`
> 3. `Selectors_and_Animation.pptx`: the class deck (Class 5–8 style; every slide has speaker notes)
>
> Rebuild the deck: `cd _build`, `npm install` once, then `npm run build`. The build screenshots each example with Microsoft Edge (headless). For the animated examples it pauses every animation at set times (Web Animations API), so the slides show frames of the motion, not just the starting position.

---

## Format: two halves

**Part 1: Advanced selectors (~35 min).** It moves quickly, with no Try Its: one example per selector family, opened and changed live, then a team game.

**Part 2: CSS animation (~30 min of concepts and demos, then ~50 min of hands-on, including show and tell).** transform → transitions → `@keyframes` → the animation properties → four "fun" demos → real-world use and accessibility. Then the **Animation Playground**, and show and tell.

| # | Topic | Example | Time |
|---|---|---|---|
| S1 | State: `:hover`, `:focus-visible`, `:active`, `:visited` (LoVe HAte order) | 01 | 5 min |
| S2 | Position: `:first-child`, `:last-child`, `:nth-child()`, `:not()` | 02 | 5 min |
| S3 | Attribute: `[type="email"]`, `^=`, `$=`, `[target]` | 03 | 5 min |
| S4 | Relationship: space, `>`, `+`, `~` | 04 | 5 min |
| S5 | Part: `::first-letter`, `::before`, `::after`, `::selection` | 05 | 5 min |
| — | **Game: Name That Selector!** (6 cards + tiebreaker, answers on the next slide) | | 10 min |
| A1 | Transitions on `:hover` | 06 | 5 min |
| A2 | `@keyframes` + percentages: **a ball around the browser** (`100vw` / `100vh`) | 07 | 5 min |
| A3 | Timing functions: **the race** (all six tie!) | 08 | 5 min |
| A4 | **Bouncing ball** with ease-in/ease-out and squash | 09 | 5 min |
| A5 | **Shape-shifting square** (missing keyframes use the normal style) | 10 | 3 min |
| A6 | **A car made of boxes**: two animations at once | 11 | 3 min |
| A7 | Real-world use (fade-in, spinner, pulse) + `prefers-reduced-motion` | 12 | 5 min |
| — | **Animation Playground**: four challenges + a bonus | activity/ | 40 min |
| — | Make it yours + show and tell | | 10 min / homework |

## The Animation Playground

`activity/index.html` has four "stages" (600 × 200px) and a bonus, each with an object that doesn't move. `stage.css` draws them (students don't touch it). Students write everything in `style.css`, under a heading for each challenge.

It's meant to be fun and very doable. Every challenge is a Try It slide with **the full code to type**, then a "play with it" step: change one number and watch what happens. Every challenge uses the same recipe: `@keyframes` with `from` / `to`, then an `animation` line.

| # | Challenge | New idea |
|---|---|---|
| 1 | Roll (typed together) | `translateX(540px)`, `animation: roll 2s infinite` |
| 2 | Ping-pong | Reuses `roll`: just add `alternate` |
| 3 | Pinwheel (a zero-size box with four coloured borders) | `rotate(360deg)`, `linear` |
| 4 | Pulse | `scale(1.5)` and a colour change together |
| ★ | Bonus: Around the Block (fast finishers) | Percentages as checkpoints (code given) |

**Make it yours:** change the colours or speed, spin the pinwheel backwards with `reverse`, turn the pulse into a circle, or start it tiny with `scale(0.5)`.

**Submit (participation):** `style.css` plus a screenshot of the playground, uploaded to Brightspace.

### Watch out for
- **Nothing moves anywhere, even in the examples:** Windows Settings → Accessibility → Visual effects → **Animation effects** is off, so the browser reports `prefers-reduced-motion: reduce`. Example 12 and the answer key respect that setting on purpose.
- The name in `animation:` doesn't match the name in `@keyframes`, or the duration has no unit (`2` instead of `2s`).
- `@keyframes` written *inside* another rule.

---

## Folder

```
06_Selectors_and_Animation/
├── README.md
├── 01_pseudo_classes_example.html
├── 02_structural_example.html
├── 03_attribute_example.html
├── 04_combinators_example.html
├── 05_pseudo_elements_example.html
├── 06_transition_example.html
├── 07_ball_around_example.html       ← a ball travels round all four edges of the browser
├── 08_timing_race_example.html       ← six balls, six timing functions, one tie
├── 09_bounce_example.html            ← gravity + squash
├── 10_morph_square_example.html      ← square → big blue circle → square, spinning
├── 11_car_example.html               ← drives across the screen while the wheels spin
├── 12_real_world_example.html        ← fade-in, spinner, pulse, prefers-reduced-motion
├── activity/
│   ├── index.html                    ← four stages + a bonus (don't edit)
│   ├── stage.css                     ← draws the stages and objects (don't edit)
│   ├── style.css                     ← starter: a heading for each challenge
│   └── answer/
│       ├── index.html                ← links ../stage.css
│       └── style.css                 ← every challenge solved, with comments
├── Selectors_and_Animation.pptx
└── _build/                           ← build_deck.js (Edge screenshots go in _build/shots, git-ignored)
```

## Deck outline (34 slides)

| # | Slide |
|---|---|
| 1–4 | Title · Attendance · Overview · Get the code |
| 5–6 | *Advanced Selectors* divider · Five new kinds of selector (the map) |
| 7–11 | Examples 01–05 |
| 12–13 | Game: Name That Selector! · Answers |
| 14–15 | *Let's Make Things Move!* divider · The transform toolkit |
| 16 | Transitions (Example 06, normal vs :hover) |
| 17 | @keyframes: describe it, then play it |
| 18 | A ball around the browser (Example 07, four frames) |
| 19 | The animation properties |
| 20–24 | The race · Bounce · Shape-shifting square · Car · Real websites and being kind (Examples 08–12) |
| 25 | Animation Playground (the four stages, and the two-step recipe) |
| 26–29 | Try It 1 Roll · Try It 2 Ping-pong · Try It 3 Pinwheel · Try It 4 Pulse |
| 30–31 | Bonus, make it yours, submit · Show and tell |
| 32–34 | Cheat sheet: selectors · Cheat sheet: animation · Recap and next steps |

## Game answers (slide 12–13)

1. `tr:nth-child(odd)` · 2. `button:hover` · 3. `a[href$=".docx"]` · 4. `.menu > li` · 5. `h2 + p` · 6. `.favourite::before { content: "★ "; }`
Tiebreaker: every button except `.cancel` → `button:not(.cancel)`
