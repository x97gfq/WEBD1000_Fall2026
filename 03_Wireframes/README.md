# Class 6: Wireframes with draw.io

**Plan for today:** a guided example you build live (Harbourside Fish & Chips, online ordering page), then students wireframe a page of their own and post it to the class Teams channel.

**Files**
- `harbourside_ordering_wireframe.drawio`: the finished guided example (page tabs: **Desktop** and **Mobile**). Open it in draw.io to see the end result. Build yours from a blank file in class so students can follow each step.
- `_build/wireframe.js`: the Harbourside layout. It writes the `.drawio` file. To regenerate: `cd _build`, then `node wireframe.js`.

---

## Where this fits

| Earlier classes | Today |
|---|---|
| Functional vs non-functional requirements, user stories ("As a ___, I want to ___ so that ___") | Each user story needs a place on the page where the user can do it |
| Semantic landmarks: `<header>`, `<nav>`, `<main>`, `<article>`, `<aside>`, `<footer>` | Each region of the wireframe is labelled with the landmark it will become |
| "The site works on phones, tablets and desktops" [NF] | Every wireframe is drawn at **desktop and mobile** sizes |

**Big idea:** a wireframe is a plan, not a finished design. It decides **what goes where** before anyone argues about colours and fonts.

## Learning outcomes

By the end of class, students can:
1. Explain what a low-fidelity wireframe is for, and why it has no colour or real images.
2. Use draw.io to build a wireframe with frames, shapes, text and mockup elements.
3. Turn user stories into regions of a page.
4. Rearrange the same content for desktop (1200px) and mobile (375px), and explain the choices.
5. Label a wireframe with the semantic HTML landmarks it will use.

---

## Part 1: Guided example, Harbourside Fish & Chips (~30 min)

**The client:** Harbourside Fish & Chips, a small takeout restaurant on the Yarmouth waterfront. They want customers to be able to order online.

**User stories** (put these on the board before opening draw.io):
1. As a **hungry customer**, I want to **browse the menu by category** so that **I can find what I want quickly**.
2. As a **customer**, I want to **choose pickup or delivery and a time** so that **my food is ready when I get there**.
3. As a **customer**, I want to **see my order and total while I'm still browsing** so that **I don't go over budget**.
4. As a **customer**, I want to **add special instructions** so that **I get my fish the way I like it**.
5. As a **phone user**, I want to **order with one thumb** so that **I can order while I'm on the go**.

Ask the class: *"Which story does each part of the page handle?"* Keep asking it as you build.

### Step 0: Set up draw.io (3 min)
- Go to **app.diagrams.net**, then **Create New Diagram → Blank Diagram**. Save it to OneDrive or your device as `harbourside.drawio`.
- Bottom-left **+ More Shapes**: turn on **Mockups**. Leave **General** on.
- Rename the page tab at the bottom to **Desktop**.

### Step 1: The frame (2 min)
- Draw a rectangle **1200 × 1060**. Set exact sizes in the right panel under **Arrange**. This is the browser window.
- Add a thin grey bar across the top with a rounded box inside that says `harbourside.ca/order`.
- 💬 *Why 1200?* It's a common width for desktop layouts. Mobile will be 375.

### Step 2: Header and nav (4 min)
- Add a logo placeholder (Mockups → **Graphics → Image**, the box with an X) and the text **Harbourside Fish & Chips**.
- Add four nav links: Menu, **Order Online** (in bold, because it's the current page), Hours & Location, Contact.
- On the right, add a **Cart (3)** button.
- 💬 Ask: *"What landmark is this? And the links?"* → `<header>` containing a `<nav>`. (Bolding the current link is the visual version of `aria-current="page"` from Class 5.)

### Step 3: Pickup/delivery banner (4 min) → story 2
- Add a light grey strip with the heading **Order for pickup or delivery**, two buttons side by side (**Pickup** filled, **Delivery** outlined), and a dropdown **Pickup time: ASAP ▾**.
- 💬 *Why is this at the top?* The customer has to decide it before they can order anything.

### Step 4: Category tabs and menu cards (7 min) → story 1
- Add a row of tab buttons: **Fish** (filled = selected), Combos, Sides, Drinks, Desserts.
- Build **one** menu card: image X box, item name, two grey bars for the description, price, and an **Add +** button.
- Select the whole card and press **Ctrl+D** to duplicate it. Use **Arrange → Align/Distribute** to make a 3 × 2 grid.
- 💬 *What landmark is one card?* A card makes sense on its own and repeats, so it's an `<article>`. Mention that this repeating pattern is what CSS Grid is for (later in the course).
- 💬 *Why grey bars instead of real text?* Low-fi on purpose. We're deciding layout here, not writing the menu.

### Step 5: Order summary (5 min) → story 3
- Draw a grey panel on the right: **Your Order**, three line items each with **− 1 +** and a price, then Subtotal / HST (14%) / **Total**, and a big **Checkout** button.
- 💬 *`<aside>` or `<section>`?* Either is fine to discuss. It's related to the main content but separate from it. Let students argue it out.

### Step 6: Footer and landmark labels (3 min)
- Footer: address, phone, hours, social icons, copyright.
- Draw **blue dashed boxes** around each region and add small tags: `<header>`, `<nav>`, `<main>`, `<article>`, `<aside>`, `<footer>`.
- 💬 Point out that this is Class 5's landmark structure. The wireframe is the plan for the HTML.

### Step 7: Mobile (8 min) → story 5
- **Right-click the page tab → Duplicate**, rename the copy **Mobile**, and change the frame to **375 wide**. (Or start with Mockups → Containers → a phone frame.)
- Then ask the class *"What has to change?"* before each move:

| Desktop | Mobile | Why |
|---|---|---|
| Nav links across the header | **☰ button**, with Cart still visible | Not enough width. Keep the most important action on screen |
| Tabs in one row | Tabs **scroll sideways** | Wrapping onto two rows pushes the menu down |
| 3-column card grid | **1-column list**: small image on the left, **+** button on the right | One thumb, one column |
| Order `<aside>` beside the menu | A **bar at the bottom** that stays on screen: *View order (3) · $33.03 →* | Nowhere to put a sidebar. The total is still always visible (story 3) |
| Order panel | A separate **Your Order screen** with **Special instructions** and Checkout at the bottom | Room for the form (story 4). Checkout sits within easy reach of a thumb |

- Draw a red dashed line at **812px** labelled *bottom of the first screen*. 💬 *What does the user see before they scroll?*
- Buttons need to be at least **44 × 44px** to tap easily. Size them to match.

### Step 8: Export and post (2 min)
- **File → Export as → PNG**, tick **Include a copy of my diagram**, and export **each page** (Desktop, Mobile).
- 💬 A PNG with the diagram included opens again in draw.io and can still be edited, so you only need to share the one file.
- Post it to the Teams channel to show students the posting format they'll use in Part 2.

---

## Part 2: Your turn (~40 min + Teams feedback)

Students wireframe **one key page** of a small site of their own, at **desktop and mobile** sizes.

**Choose a client** (or pitch your own):
- 🥐 A bakery: pre-order a birthday cake
- 🎸 A local band: upcoming shows and buying tickets
- 🐾 A dog groomer: book an appointment
- 🏒 A minor hockey association: team schedule and registration
- 🚚 A food truck: today's location and menu
- 🏕️ A campground: pick a site and reserve dates

**Requirements checklist**
- [ ] 3–4 user stories written at the top of the page (*As a ___, I want to ___ so that ___*)
- [ ] **Desktop** frame (1200px wide) on one page tab, **Mobile** frame (375px wide) on another
- [ ] Low-fi only: greyscale, X boxes for images, grey bars for body text
- [ ] Every user story can be done somewhere on the page. Number them on the wireframe: ① ② ③
- [ ] Blue dashed landmark labels: `<header>`, `<nav>`, `<main>`, `<footer>`, plus at least one of `<article>`, `<section>`, `<aside>`
- [ ] At least **two** things are rearranged for mobile, not just made narrower
- [ ] Tap targets on mobile look about 44px or larger

**Post to Teams** (in the class channel, one post per student):
1. Your two PNGs (exported with **Include a copy of my diagram**)
2. Your name and client, e.g. *"Jordan: Paws & Suds dog grooming, booking page"*
3. Your user stories
4. **One sentence** about something you changed for mobile, and why

**Give feedback** by replying to **two** classmates' posts (spread out: reply to posts that have fewer than two replies):
- ✅ **Works:** name one user story you could clearly complete, and how
- ❓ **Wonder:** one place where a user might get stuck, or something you'd move

---

## Wrap-up (5 min)

Put two or three Teams posts on the projector. For each one, ask:
- *"If you had to write the HTML tomorrow, what's the landmark structure?"*
- *"What's the first thing a phone user sees before scrolling? Is it the right thing?"*

## Common mistakes to look for while circulating
- Colours, fonts and stock photos. Send them back to greyscale.
- Mobile is the desktop layout made smaller (three tiny columns).
- A user story with nowhere on the page to do it (or a big feature with no story behind it).
- Everything labelled `<div>`, or no `<main>`.
- Things not aligned. Show them **Arrange → Align** and **View → Grid**.
