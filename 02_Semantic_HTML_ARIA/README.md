# Class 5: Accessibility in Code (Semantic HTML5 & ARIA)

> **Status:** Approved. This README is the source for:
> 1. `02_Semantic_HTML_ARIA/*.html`: numbered code samples (same style as `01_HTML_Examples`)
> 2. `02_Semantic_HTML_ARIA/Class5_Semantic_HTML_ARIA.pptx`: the class deck, styled like the Class 4 deck (code slide per example, activity slides with the answer key in the speaker notes)
>
> **Decisions:** Class 4 slide style · labelled **Class 5** · all 10 examples · extra ARIA explanation slides (see *ARIA background* below) · students already have Live Server · the Class 4 "Semantic HTML" code is reused, corrected, as the bridge slide.

---

## Where we left off (Class 4 recap)

Class 4 was the **why** of accessibility. Today is the **how**: what it looks like in the HTML we write.

| Class 4 covered… | Today we write the code for it |
|---|---|
| Visitors have different visual, auditory, physical, speech, cognitive, and neurological abilities | Screen readers and keyboards only work well if our HTML describes the page |
| **WCAG 2.1 Level AA** is the standard in Canada (a11y.canada.ca) | Examples are tagged with the WCAG principle they support |
| **POUR**: Perceivable, Operable, Understandable, Robust | Each example is labelled **P / O / U / R** |
| *Robust:* "Use semantic HTML (`<header>`, `<nav>`, `<main>`)" and "ARIA landmarks" | Examples 01–05 (semantic HTML), 07–09 (ARIA) |
| *Perceivable:* text alternatives (alt) | Examples 04, 06 |
| *Operable:* keyboard access, focus, headings and landmarks | Examples 02, 07 |
| *Understandable:* error identification in forms | Example 10 |
| Edge Reader View (F9), Read Aloud, Lighthouse (Ctrl+Shift+I) | Used again in today's activity, on **our own** page |

**Today's big idea:** last class we ran Lighthouse and Read Aloud on nscc.ca and asked *"why did it do well or badly?"* The answer is mostly the HTML underneath. Today we write a bad page, test it, fix it, and test it again.

## Learning outcomes

By the end of class, students can:
1. Explain how semantic HTML supports the WCAG **Robust** principle (and helps SEO, as mentioned in Class 4).
2. Structure a page with `<header>`, `<nav>`, `<main>`, `<article>`, `<section>`, `<aside>`, `<footer>`, and a proper heading order.
3. Write useful `alt` text (including `alt=""`), and use `<figure>`/`<figcaption>` and accessible tables.
4. State the **first rule of ARIA** and use `aria-label`, `aria-labelledby`, `aria-current`, `aria-hidden`, `aria-live`, `aria-describedby`.
5. Build a form with proper `<label>`s and an accessible error message.
6. Use Lighthouse **and** a keyboard test to check a page, and explain why automated tools alone aren't enough.

---

## Folder & file plan

```
02_Semantic_HTML_ARIA/
├── README.md
├── 01_div_soup_example.html            ← the "before" (R)
├── 02_landmarks_example.html           ← header / nav / main / footer + skip link (R, O)
├── 03_article_section_aside_example.html  (R)
├── 04_figure_figcaption_example.html   ← builds on last week's 07_picture (P)
├── 05_accessible_table_example.html    ← builds on 06_table (P, R)
├── 06_alt_text_example.html            ← good / bad / decorative (P)
├── 07_aria_first_rule_example.html     ← <div onclick> vs <button>, keyboard (O, R)
├── 08_aria_labels_example.html         ← aria-label / labelledby / current (R)
├── 09_aria_hidden_live_example.html    ← aria-hidden / aria-live (P, R)
├── 10_accessible_form_example.html     ← label, required, aria-describedby error (U)
├── activity.html                       ← inaccessible "Burridge Campus" page to fix
├── activity_answer.html                ← answer key (also in pptx speaker notes)
├── beach.JPG                           ← copied from 01_HTML_Examples
├── wave-divider.png                    ← decorative image for Example 06 (alt="")
├── Class5_Semantic_HTML_ARIA.pptx
└── _build/                             ← deck generator: npm install, then npm run build
```

All samples use last week's boilerplate (`<!DOCTYPE html>`, `lang="en"`, `charset`, `viewport`, `<title>`), tab indentation, and `<!-- comments -->` on the key lines.

---

## Deck outline (≈35 slides, Class 4 visual style)

| # | Slide | Layout (from Class 4) |
|---|---|---|
| 1 | Title: Class 5, Accessibility in Code | Orange circle title slide |
| 2 | Attendance | Plain |
| 3 | Class 5 overview / agenda | Dark left panel |
| 4 | Get the code (`git pull`, Live Server) | Dark left panel |
| 5 | Class 4 recap: WCAG 2.1 AA and POUR | Dark left panel |
| 6 | Semantic HTML: the bridge (Class 4 code, corrected) | Gradient + code |
| 7 | **Part 1: Semantic HTML5** | Gradient divider |
| 8 | Why semantic HTML matters: who reads your HTML | Dark left panel |
| 9 | How screen reader users actually navigate | Dark left panel |
| 10–15 | Examples 01–06 | Gradient + code |
| 16 | **Part 2: ARIA** | Gradient divider |
| 17 | What is ARIA? | Dark left panel |
| 18 | The accessibility tree (diagram) | Dark left panel |
| 19 | Roles, properties, and states | Dark left panel |
| 20 | The five rules of ARIA | Dark left panel |
| 21 | Semantic HTML = built-in ARIA (implicit roles table) | Dark left panel |
| 22 | Example 07: first rule / keyboard | Gradient + code |
| 23 | Accessible names: label vs labelledby vs describedby | Dark left panel |
| 24–26 | Examples 08–10 | Gradient + code |
| 27 | ARIA cheat sheet | Dark left panel |
| 28 | Checking your work: DevTools accessibility pane, Narrator, Tab | Dark left panel |
| 29 | **Class Activity** | Gradient divider |
| 30–32 | Activity parts 1–3 (test → fix → re-test) | Class 4 "Activity" slide |
| 33 | "How did you make out?" Is Lighthouse 100 = done? | Ring slide |
| 34 | Accessibility checklist | Dark left panel |
| 35 | Class recap and next steps | Dark left panel |

Every slide has detailed speaker notes.

---

## ARIA background (the descriptive text for slides 17–23, 27–28)

### What is ARIA?
- **WAI-ARIA** stands for *Web Accessibility Initiative – Accessible Rich Internet Applications*. It's a W3C specification from the same group (WAI) that publishes WCAG.
- ARIA is a set of **HTML attributes**: `role="…"` plus attributes starting with `aria-`. They add to, or change, what assistive technology (screen readers, voice control, braille displays) is told about an element.
- ARIA changes **only what assistive technology hears**. It does **not** change how the page looks, and it does **not** add behaviour: no keyboard support, no focus, no click handling. That's the most common misunderstanding.
- ARIA exists to fill gaps: custom widgets (tabs, menus, pop-ups) and content that updates without a page reload. Those didn't have native HTML elements when ARIA was created.

### The accessibility tree
- The browser turns your HTML into the **DOM**. From the DOM it builds a second, simpler tree: the **accessibility tree**.
- Each node in that tree has a **role** (what it is), a **name** (what it's called), a **state** (e.g. checked, expanded, current), and optionally a **description**.
- Screen readers read the accessibility tree, not your pixels.
- Semantic HTML fills in role and name automatically. ARIA lets you edit the tree when HTML can't.
- You can see it in DevTools: Elements → Accessibility pane (Chrome/Edge), or the Accessibility Inspector (Firefox).

### Roles, properties, and states
- **Roles** say what something *is*: `role="button"`, `role="navigation"`, `role="alert"`. HTML elements already have roles (implicit roles).
- **Properties** describe characteristics that usually don't change: `aria-label`, `aria-labelledby`, `aria-describedby`, `aria-required`.
- **States** describe conditions that change as the user interacts: `aria-expanded`, `aria-checked`, `aria-current`, `aria-invalid`, `aria-hidden`.

### The five rules of ARIA (from W3C's *Using ARIA*)
1. **Use native HTML first.** If an HTML element or attribute already has the meaning and behaviour you need, use it instead of adding ARIA.
2. **Don't change native semantics** unless you really have to. For example, don't put `role="button"` on an `<h2>`. Wrap a `<button>` inside the heading instead.
3. **All interactive ARIA controls must work with a keyboard.** If you make a `role="button"`, you must also make it focusable and respond to Enter and Space.
4. **Don't hide focusable elements.** Never put `aria-hidden="true"` or `role="presentation"` on something a user can Tab to.
5. **All interactive elements need an accessible name.** Every button, link, and form field must have a name a screen reader can say.

*"No ARIA is better than bad ARIA."* WebAIM's yearly analysis of the top million home pages consistently finds that pages using ARIA have **more** detected accessibility errors on average than pages without it. ARIA is powerful, but misused ARIA makes pages worse.

### Semantic HTML = built-in ARIA (implicit roles)
Class 4 said to use *"ARIA landmarks"*. With HTML5 you get those landmarks for free by using the right elements:

| HTML element | Implicit ARIA role | Screen reader says |
|---|---|---|
| `<header>` (top-level) | `banner` | "banner" |
| `<nav>` | `navigation` | "navigation" |
| `<main>` | `main` | "main" |
| `<aside>` | `complementary` | "complementary" |
| `<footer>` (top-level) | `contentinfo` | "content info" |
| `<section>` with a name | `region` | "region" |
| `<article>` | `article` | "article" |
| `<button>` | `button` | "button" |
| `<a href>` | `link` | "link" |
| `<h1>`–`<h6>` | `heading` (level 1–6) | "heading level 2" |
| `<img alt="…">` / `alt=""` | `img` / none (ignored) | "graphic" / *(skipped)* |

So `<nav role="navigation">` is **redundant**. That was a workaround for very old browsers and isn't needed today.

### Accessible names: three ways to label
| Attribute | Use it when | Example |
|---|---|---|
| `aria-label="…"` | There's no visible text to use, e.g. an icon-only button or two navs | `<button aria-label="Close menu">&times;</button>` |
| `aria-labelledby="id"` | Visible text already on the page names the thing | `<section aria-labelledby="news-heading">` |
| `aria-describedby="id"` | You want *extra* info read after the name, e.g. a hint or error | `<input aria-describedby="email-error">` |

- If several are present, the browser uses this priority: **aria-labelledby → aria-label → native label** (`<label>`, `alt`, the element's text).
- `aria-label` **replaces** the visible text for screen readers, so make sure the name includes the words that are on screen. Voice-control users will say those visible words.
- Prefer a real `<label>` for form fields over any ARIA.

### Cheat sheet
`aria-label`, `aria-labelledby`, `aria-describedby`, `aria-current="page"`, `aria-hidden="true"`, `aria-live="polite" | "assertive"`, `aria-expanded="true|false"`, `aria-invalid="true"`, `aria-required="true"` (prefer the HTML `required` attribute).

### Checking your work
- **Keyboard:** Tab / Shift+Tab to move, Enter / Space to activate. Can you reach and use everything? Can you see where focus is?
- **DevTools accessibility pane:** inspect an element to see its computed **Role** and **Name**.
- **Windows Narrator:** built into Windows, toggle with **Win + Ctrl + Enter**. **NVDA** is a free alternative (nvaccess.org).
- **Lighthouse:** a good first pass, but it only catches *some* problems.

---

## Examples (draft code)

### 01 — Div soup (the "before") · *Robust*
**Point:** This page *looks* fine, but to a screen reader it's one long blob. There are no landmarks to jump to and no headings to skim. This is why Read Aloud struggles on some sites.

```html
<div class="header">
	<div class="title">Burridge Campus</div>
</div>
<div class="menu">
	<div><a href="index.html">Home</a></div>
	<div><a href="programs.html">Programs</a></div>
</div>
<div class="content">
	<div class="big">Welcome</div>
	<div>Welcoming campus with big school spirit...</div>
</div>
<div class="footer">&copy; 2026 NSCC</div>
```

### 02 — Landmarks + skip link · *Robust, Operable*
**Point:** The same page with meaningful elements. Screen readers list these as landmarks users can jump between (the "Navigable Content" bullet from Class 4). Use only **one `<main>`** per page. The skip link lets keyboard users jump past the menu.

```html
<a href="#main">Skip to main content</a>

<header>
	<h1>Burridge Campus</h1>
</header>

<nav>
	<ul>
		<li><a href="index.html">Home</a></li>
		<li><a href="programs.html">Programs</a></li>
	</ul>
</nav>

<main id="main">
	<h2>Welcome</h2>
	<p>Welcoming campus with big school spirit...</p>
</main>

<footer>
	<p>&copy; 2026 Nova Scotia Community College</p>
</footer>
```

### 03 — article, section, aside · *Robust*
**Point:**
- `<article>` is content that makes sense **on its own**, like a news item or a blog post.
- `<section>` is a themed **part** of something, and usually has a heading.
- `<aside>` is related content that isn't essential to the main content.

Reader View (F9) uses `<article>` to decide what to show.

```html
<main>
	<article>
		<h2>Campus Cinnamon Bun Day</h2>
		<p>Every Wednesday the cafeteria...</p>

		<section>
			<h3>When</h3>
			<p>Wednesdays, starting at 10am.</p>
		</section>

		<section>
			<h3>Where</h3>
			<p>The main cafeteria.</p>
		</section>
	</article>

	<aside>
		<h2>Did you know?</h2>
		<p>Burridge hosts one of the Cube locations for entrepreneurs.</p>
	</aside>
</main>
```

### 04 — figure & figcaption · *Perceivable*
**Point:** This builds on last week's `07_picture_example`. `alt` replaces the image for people who can't see it. `<figcaption>` is a **visible** caption for everyone.

```html
<figure>
	<img src="beach.JPG" alt="Sandy beach with waves under a clear blue sky">
	<figcaption>Mavillette Beach, a short drive from the Yarmouth campus.</figcaption>
</figure>
```

### 05 — Accessible table · *Perceivable, Robust*
**Point:** This builds on last week's `06_table_example`.
- `<caption>` names the table.
- `<thead>`/`<tbody>` separate the headers from the data.
- `scope` tells a screen reader which cells each header belongs to, so it reads "August, Temperature, +20°C" instead of just "+20".

```html
<table>
	<caption>Average Temperature in Yarmouth</caption>
	<thead>
		<tr>
			<th scope="col">Month</th>
			<th scope="col">Temperature</th>
		</tr>
	</thead>
	<tbody>
		<tr>
			<th scope="row">January</th>
			<td>-10&deg;C</td>
		</tr>
		<tr>
			<th scope="row">August</th>
			<td>+20&deg;C</td>
		</tr>
	</tbody>
</table>
```

### 06 — Alt text: good, bad, decorative · *Perceivable*
**Point:** This is the "Text Alternatives" bullet from Class 4.
- `alt` should say what the image *communicates*.
- Decorative images get `alt=""` so screen readers skip them.
- A missing `alt` is the worst case: many screen readers read out the file name instead. Lighthouse flags this.

```html
<!-- Bad: missing alt (Lighthouse will flag this) -->
<img src="beach.JPG">

<!-- Bad: not helpful -->
<img src="beach.JPG" alt="image">

<!-- Good: describes the content -->
<img src="beach.JPG" alt="Sandy beach with waves under a clear blue sky">

<!-- Decorative: empty alt, screen readers skip it -->
<img src="divider.png" alt="">
```

### 07 — The first rule of ARIA / keyboard access · *Operable, Robust*
**Point:** Class 4 said *"implement ARIA roles"*. Here's the catch: **if a native HTML element can do the job, use it instead of ARIA.**
- The fake button can't be reached with **Tab** or activated with **Enter**.
- The real `<button>` gets both, plus a visible focus ring, for free.

Students try it by pressing Tab through the page.

```html
<!-- Avoid: looks like a button, but a keyboard can't reach it -->
<div onclick="alert('Subscribed!')">Subscribe</div>

<!-- ARIA "fix": a screen reader now says "button"... but Tab still skips it -->
<div role="button" onclick="alert('Subscribed!')">Subscribe</div>

<!-- Prefer: a real button (keyboard, focus, and screen reader support built in) -->
<button type="button" onclick="alert('Subscribed!')">Subscribe</button>
```

### 08 — aria-label, aria-labelledby, aria-current · *Robust*
**Point:** Use ARIA *labels* when the visible text isn't enough, for example:
- two `<nav>`s on the same page
- an icon-only button

`aria-current="page"` tells screen readers which link is the page you're on (the "Consistent Navigation" bullet from Class 4).

```html
<nav aria-label="Main">
	<ul>
		<li><a href="index.html" aria-current="page">Home</a></li>
		<li><a href="programs.html">Programs</a></li>
	</ul>
</nav>

<section aria-labelledby="news-heading">
	<h2 id="news-heading">Campus News</h2>
	<p>...</p>
</section>

<nav aria-label="Footer">
	<a href="privacy.html">Privacy</a>
</nav>

<!-- Icon-only button: without aria-label a screen reader just says "times, button" -->
<button type="button" aria-label="Close menu">&times;</button>
```

### 09 — aria-hidden and aria-live · *Perceivable, Robust*
**Point:**
- `aria-hidden="true"` hides decoration (like an icon next to text that already says the same thing) from screen readers.
- `aria-live="polite"` announces content that changes *after* the page loads.

```html
<!-- The star is decoration; the text already says "Favourite" -->
<button type="button"><span aria-hidden="true">&#9733;</span> Favourite</button>

<!-- Screen readers announce changes to this area -->
<p id="status" aria-live="polite"></p>
<button type="button" onclick="document.getElementById('status').textContent = 'Saved!'">
	Save
</button>
```

### 10 — Accessible form · *Understandable*
**Point:** This is the "Error Identification and Suggestions" bullet from Class 4.
- Every input needs a `<label>`; clicking the label also focuses the field.
- `aria-describedby` connects a hint or error message to its field, so it's read aloud when the field gets focus.
- `aria-invalid` marks the field as having an error.

```html
<form>
	<label for="email">Email address</label>
	<input type="email" id="email" name="email" required
		aria-describedby="email-hint email-error" aria-invalid="true">
	<p id="email-hint">We'll only use this to send your timetable.</p>
	<p id="email-error">Error: please enter an email like name@example.ca</p>

	<button type="submit">Sign up</button>
</form>
```

---

## Class Activity: test → fix → re-test

This uses the same tools as Class 4 (Lighthouse, Read Aloud), but on a page students fix themselves.

**`activity.html`** is a "Burridge Campus" page (using the campus text from `page1.html`/`page2.html`), deliberately broken so that Lighthouse and a keyboard test both find problems:

| Problem planted | POUR | Caught by Lighthouse? |
|---|---|---|
| All `<div>`s, no landmarks | R / O | Partly |
| "Headings" are styled divs; heading levels skip | O | Partly |
| `<img>` with no `alt` | P | ✅ Yes |
| Light-grey text on white (low contrast) | P | ✅ Yes |
| `<html>` missing `lang` | U | ✅ Yes |
| Icon-only `<button>&times;</button>` with no name | R | ✅ Yes |
| Form input with no `<label>` | U | ✅ Yes |
| Table with no `caption`/`th`/`scope` | P | ❌ No |
| `<div onclick>` fake "Subscribe" button | O | ❌ **No, only a keyboard test finds it** |

**Part 1: Test (5 min).**
1. Open `activity.html`.
2. Run Lighthouse and choose **Accessibility** only. Write down the score.
3. Press **Tab** through the page. Can you reach "Subscribe"?

**Part 2: Fix (25 min).**
1. Replace the divs with `header`/`nav`/`main`/`article`/`section`/`aside`/`footer`, and add a skip link.
2. Use real headings, in order (`h1` → `h2` → `h3`).
3. Wrap the image in `<figure>`, and add good `alt` text and a `<figcaption>`.
4. Add `lang="en"` and fix the contrast.
5. Make the table accessible.
6. Turn the fake button into a real `<button>`, and give the × button an `aria-label`.
7. Label the form field, and connect the error message with `aria-describedby`.
8. Add `aria-label="Main"` to the nav and `aria-current="page"` to the Home link.

**Part 3: Re-test (5 min).**
1. Run Lighthouse again and compare the scores.
2. Try Read Aloud.
3. Discussion: *"Lighthouse gave us 100. Is the page done?"* This leads into the table and the fake button, which Lighthouse doesn't check, and why Class 4 listed **manual testing** alongside the automated tools.

> **Setup note:** Open the page with the VS Code **Live Server** extension (right-click the file, then "Open with Live Server"). Lighthouse may not run on a plain `file://` page.

**`activity_answer.html`** is the fixed page. It also goes in the speaker notes of the activity slides.

## Wrap-up checklist (final slide)
- [ ] `lang` on `<html>`, a real `<title>`
- [ ] One `<h1>`, headings in order
- [ ] `<header>`, `<nav>`, `<main>` (one), `<footer>`
- [ ] Every `<img>` has `alt` (empty if decorative); good contrast
- [ ] Every input has a `<label>`; errors say what to fix
- [ ] Clickable = `<a>` (goes somewhere) or `<button>` (does something); ARIA only when HTML can't do it
- [ ] Test with Lighthouse **and** your keyboard

---
