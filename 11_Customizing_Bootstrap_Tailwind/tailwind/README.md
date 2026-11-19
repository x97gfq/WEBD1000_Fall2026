# Tailwind Examples (Class 2)

Five short pages that cover the basics of Tailwind CSS. Each one opens on its own in a browser.

Every page loads Tailwind from the **Play CDN** (`@tailwindcss/browser@4`), so **you need the internet**. The Play CDN reads the classes in your HTML and writes the CSS while the page loads. That's fine for learning. Real projects build the CSS file ahead of time.

Each page ends with a yellow **Try this** box: a small change to make, save, and refresh.

## The examples

### 01 – Spacing, colour and text
[01_spacing_colour_text_example.html](01_spacing_colour_text_example.html)

The three scales behind most Tailwind classes:

- **Spacing:** `p-` (padding) and `m-` (margin). Each step is 4px, so `p-4` is 16px. Add a side with `px`, `py`, `pt`, and so on.
- **Colour:** a colour name and a shade from 50 (nearly white) to 950 (nearly black), e.g. `bg-orange-500`. The same names work with `bg-`, `text-` and `border-`.
- **Text:** sizes (`text-sm`, `text-base`, `text-xl`, `text-3xl`…) and weights (`font-semibold`, `font-bold`).

The page ends with a "Today's special" card that uses all three.

### 02 – Hover and focus
[02_hover_and_focus_example.html](02_hover_and_focus_example.html)

In plain CSS you'd write a second rule for `:hover`. In Tailwind you put a **prefix** in front of a class:

- `hover:bg-blue-800`: only applies while the mouse is over the element.
- `transition` and `duration-300`: make the change fade in instead of jumping.
- `focus:ring-4`: shows a ring when an input is clicked or reached with Tab. Keyboard users need this.
- `disabled:bg-slate-300`: applies when the element has the `disabled` attribute.

### 03 – Responsive (sm, md, lg)
[03_responsive_example.html](03_responsive_example.html)

Tailwind is **mobile first**. A class with no prefix applies at every width. `md:` means "from 768px and wider, also do this".

| Prefix | From |
|---|---|
| `sm:` | 640px |
| `md:` | 768px |
| `lg:` | 1024px |
| `xl:` | 1280px |

A badge at the top shows which size you're at, so resize the window and watch it change. The page shows one, two, then four columns (`grid-cols-1 sm:grid-cols-2 lg:grid-cols-4`), a card that stacks on a phone and sits side by side from `md`, and text and padding that grow.

### 04 – A header with flexbox
[04_flex_header_example.html](04_flex_header_example.html)

Tailwind has no ready-made navbar the way Bootstrap does. You build one from flexbox classes, and each class is one CSS property you already know from Week 8:

| Class | CSS |
|---|---|
| `flex` | `display: flex` |
| `justify-between` | `justify-content: space-between` |
| `items-center` | `align-items: center` |
| `gap-6` | `gap: 1.5rem` |

The page also puts `justify-start`, `justify-center` and `justify-between` side by side, and builds a price list with the name on the left and the price on the right. Tailwind includes no JavaScript, so a menu that collapses on a phone is not something it can do on its own. That's one reason people still choose Bootstrap.

### 05 – Your own colours with @theme
[05_theme_colours_example.html](05_theme_colours_example.html)

This is the Tailwind answer to "how do I use **our** colours?", the same question as Example 01 in Bootstrap. A `<style type="text/tailwindcss">` block adds Harbourside's colours to Tailwind's theme:

```css
@theme {
	--color-harbour: #12304A;
	--color-ember: #B5451B;
}
```

Tailwind then creates `bg-harbour`, `text-ember`, `border-ember`, `hover:bg-ember`, and so on. Bootstrap's button went back to blue on hover. Here `hover:bg-ember-dark` is right in the class list, so there's nothing hidden to override.

## Bootstrap or Tailwind?

| | Bootstrap | Tailwind |
|---|---|---|
| You get | Ready-made components (`.card`, `.navbar`, `.btn`) | Small one-job classes you combine yourself |
| Looks like | Bootstrap, unless you customize it | Whatever you build |
| JavaScript | Included (modals, collapse, carousel…) | None |
| HTML | Short class lists | Long class lists |
