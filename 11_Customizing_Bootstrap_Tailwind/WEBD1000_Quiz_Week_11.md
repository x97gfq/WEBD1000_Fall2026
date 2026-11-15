# WEBD 1000 — Week 11 Quiz: Customizing Bootstrap, Its JavaScript Components, and Tailwind

**10 multiple-choice questions.** Every question comes from a pop quiz on the Week 11 slides, so a
student who was in both classes and wrote the answers down has already seen all ten.

## Suggested Brightspace settings

| Setting | Value |
|---|---|
| Questions | 10, one point each |
| Shuffle | Shuffle questions **and** answers |
| Attempts | 2, highest attempt counts |
| Time limit | None (or 20 minutes) |
| Opens | After Class 2 |
| Closes | End of the week |
| Feedback | Show the answer and the feedback after the attempt is submitted |

## Pasteable description

> This quiz covers Week 11: customizing Bootstrap with your own CSS, Bootstrap's JavaScript
> components (modal, collapse, accordion, carousel, offcanvas, tooltip and toast), your first few
> lines of JavaScript, and a first look at Tailwind CSS.
>
> Every question is one of the ten pop quizzes from this week's two classes. If you want to review,
> the pop quiz slides and their answers are in the Week 11 deck, and all of the example pages are
> in the `11_Customizing_Bootstrap_Tailwind` folder. You have two attempts, and your best one counts.

## Where each question comes from

| Quiz | Pop quiz | Class | Topic |
|---|---|---|---|
| 1 | Pop Quiz 1 | 1 | Specificity: `.btn:hover` beats `.btn-primary` |
| 2 | Pop Quiz 2 (trivia) | 1 | Why JavaScript is called JavaScript |
| 3 | Pop Quiz 3 | 1 | `data-bs-toggle` / `data-bs-target` |
| 4 | Pop Quiz 4 | 1 | What breaks without the bundle script |
| 5 | Pop Quiz 5 | 1 | Tooltips are opt in |
| 6 | Pop Quiz 6 | 1 | A toast needs `.show()` |
| 7 | Pop Quiz 7 | 2 | Tailwind's `md:` prefix |
| 8 | Pop Quiz 8 | 2 | Recognizing Tailwind classes |
| 9 | Pop Quiz 9 | 2 | `table-responsive` |
| 10 | Pop Quiz 10 | 2 | When rules tie, the later one wins |

---

## Question 1

Your `style.css` is linked after Bootstrap's and contains this rule. What colour is the button
**while you hover over it**?

```css
.btn-primary {
	background-color: #ff5722;
	border-color: #ff5722;
}
```

- a) Bootstrap's blue **(correct)**
- b) Orange, the same as when you are not hovering
- c) A slightly darker orange
- d) See-through, because the rule removes the colour

**Feedback:** Bootstrap has a `.btn:hover` rule. That is a class plus a pseudo-class, which is more
specific than your single-class `.btn-primary`, so it wins no matter which file came last. Its hover
colour comes from `var(--bs-btn-hover-bg)`, which is still blue. The fix is to set the variables
instead: `--bs-btn-bg`, `--bs-btn-hover-bg`, and so on. See Example 01.

## Question 2

JavaScript and Java are two completely unrelated languages. So why is JavaScript called JavaScript?

- a) It was renamed in a 1995 marketing deal with Sun Microsystems, who owned Java **(correct)**
- b) It is the version of Java that runs inside a web browser
- c) Java and JavaScript were designed by the same team at Sun
- d) JavaScript is converted into Java before the browser runs it

**Feedback:** The language was called Mocha, then LiveScript. In December 1995, Netscape and Sun
announced a deal and LiveScript was renamed JavaScript, because Java was the most talked-about
language of that year. That is the entire reason. Java is to JavaScript as ham is to hamster.

## Question 3

You want this button to open the modal below it. What has to go on the button?

```html
<button type="button" class="btn btn-primary">
	Order the lobster roll
</button>

<div class="modal fade" id="orderModal">
	...
</div>
```

- a) `data-bs-toggle="modal"` and `data-bs-target="#orderModal"` **(correct)**
- b) `data-toggle="modal"` and `data-target="#orderModal"`
- c) `data-bs-toggle="#orderModal"` on its own
- d) `onclick="openModal()"`

**Feedback:** `data-bs-toggle` says **what kind** of thing to open, and `data-bs-target` says
**which one**, by its `id`, with a `#` in front because it is a CSS selector. Option b is
Bootstrap 4 code: Bootstrap 5's JavaScript only ever looks for `data-bs-`, so those attributes do
nothing at all. Option d would need JavaScript you write yourself.

## Question 4

A page is working correctly. You delete this line and reload:

```html
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.8/dist/js/bootstrap.bundle.min.js"></script>
```

What happens?

- a) The modal, the carousel and the navbar's menu button stop working; the grid, cards and
  colours are unaffected **(correct)**
- b) The whole page loses its styling, because that file contains Bootstrap
- c) Nothing: the `data-bs-` attributes do the work on their own
- d) The browser shows an error in the Console explaining what is missing

**Feedback:** Bootstrap arrives in two separate downloads. The CSS file controls how things look;
the script controls the things that move. Take the script away and everything visual is fine while
every interactive component silently stops. And there is no error — nothing was loaded, so nothing
complained. A button that does nothing, with an empty Console, is almost always a missing script.

## Question 5

Your button has `data-bs-toggle="tooltip"` and `data-bs-title="..."`, Bootstrap's script is loaded,
and the Console shows no errors. You hover over the button and no tooltip appears. Why?

- a) Tooltips are opt in: a line of JavaScript has to turn them on **(correct)**
- b) Tooltips only work on links, not on buttons
- c) You have to use the plain `title` attribute instead of `data-bs-title`
- d) Tooltips come in a separate file that has to be linked as well

**Feedback:** Tooltips and popovers are the one exception to "the attributes are enough", and the
reason is speed: a page could have hundreds of elements with a title, and setting all of them up on
every page load would be wasteful. So you say which ones you want:

```js
document.querySelectorAll('[data-bs-toggle="tooltip"]')
	.forEach(function (el) {
		new bootstrap.Tooltip(el);
	});
```

## Question 6

The button with `id="orderButton"` is clicked. What happens?

```js
const toast = new bootstrap.Toast(
	document.getElementById("orderToast"));

document.getElementById("orderButton")
	.addEventListener("click", function () {
		console.log("clicked");
	});
```

- a) "clicked" appears in the Console, but the toast never appears **(correct)**
- b) The toast appears
- c) Nothing happens at all
- d) The toast appears once and then stops working

**Feedback:** The first line gets the toast **ready**; it does not show it. A toast stays hidden
until something calls `toast.show()`. The click listener is fine, so the Console message does
appear. Adding `toast.show();` inside the braces makes it work. "My code ran but nothing happened"
and "my code never ran" are two different bugs, and `console.log` is how you tell them apart.

## Question 7

In Tailwind CSS, what does the class `md:grid-cols-3` mean?

- a) Three columns, on screens 768px wide and wider **(correct)**
- b) Three columns, but only on medium-sized screens
- c) Three columns, on screens up to 768px wide
- d) Three columns, each of medium width

**Feedback:** Like Bootstrap, Tailwind is written phones-first, so `md:` means
`@media (min-width: 768px)`. The rule applies from 768px upwards, including large screens, until a
bigger prefix such as `lg:` overrides it. Below 768px there is no `grid-cols` at all, so the items
stack in one column.

## Question 8

Which of these `class` attributes was written with Tailwind CSS rather than Bootstrap?

- a) `class="bg-white rounded-lg shadow p-5"` **(correct)**
- b) `class="card h-100 shadow-sm"`
- c) `class="btn btn-primary"`
- d) `class="alert alert-warning"`

**Feedback:** Each Tailwind class does roughly one thing: a background, a border radius, a shadow,
some padding. You build the look out of small pieces. `card`, `btn` and `alert` are Bootstrap's
ready-made components, where one name gives you a whole designed thing.

## Question 9

Your opening-hours table is wider than a phone screen. You wrap it in
`<div class="table-responsive">`. What does that do?

- a) Lets the table scroll sideways inside its own box **(correct)**
- b) Shrinks the text until the table fits
- c) Stacks the table's columns on top of each other
- d) Hides the columns that do not fit

**Feedback:** The `table-responsive` div gets `overflow-x: auto`, so the table gets its own
sideways scrollbar and the rest of the page stays still. Without it, a wide table makes the whole
page scroll sideways — exactly the problem we spent Week 9 avoiding. We did this by hand in Week 7
with a `table-wrap` div; Bootstrap just has a class for it.

## Question 10

An element has `class="card special"`. Both rules below use a single class, so they tie on
specificity. What colour is its background?

```css
/* bootstrap.min.css, linked FIRST */
.card {
	background-color: #fff;
}

/* your style.css, linked SECOND */
.special {
	background-color: #FFF8E1;
}
```

- a) Cream (`#FFF8E1`): the rules tie, and yours is read last **(correct)**
- b) White (`#fff`): `.card` is Bootstrap's own class, so Bootstrap wins
- c) White (`#fff`): Bootstrap's file is larger and loads first
- d) Neither: you need `!important` to override a framework

**Feedback:** The browser checks specificity first. Only when two rules tie does the order decide,
and then the one read last wins — which is the whole reason your `<link>` to `style.css` goes
*after* Bootstrap's. Compare this with Question 1, where the rules did **not** tie: there
`.btn:hover` was more specific, and coming later did not help at all. Specificity first, then order.
