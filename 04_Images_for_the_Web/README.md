# Class 7: Images for the Web

> This README is the source for:
> 1. `04_Images_for_the_Web/*.html`: numbered code samples (same style as earlier weeks)
> 2. `Class7_Images_for_the_Web.pptx`: the class deck (Class 5/6 style; every slide has speaker notes)
>
> Rebuild the deck and sample images: `cd _build`, `npm install` once, then `npm run build`.
> `make_samples.js` regenerates `samples/` from `../02_Semantic_HTML_ARIA/beach.JPG` and writes `sizes.json`, so the file sizes on the slides are real.

---

## Where this fits

| Earlier classes | Today |
|---|---|
| `<img>`, `alt`, `<figure>`/`<figcaption>` | What goes *inside* the image file: format, size, compression |
| "Pages load in under 2 seconds" [NF] | Images are usually the heaviest part of a page. This is how you meet that requirement |
| Wireframes full of X boxes | Today we make the real images that replace them |

**Big idea:** the goal is *the smallest file that still looks good at the size it's shown.*

## Learning outcomes

By the end of class, students can:
1. Explain pixels, dimensions, megapixels and colour (bit) depth, and why DPI doesn't matter on the web.
2. Explain lossless vs lossy compression, and what lossy artifacts look like.
3. Choose between JPG, PNG, GIF, WebP, AVIF and SVG for a given image.
4. Resize and export an image at a sensible size and quality, and check the result in File Explorer.
5. Use `width`/`height`, `loading="lazy"`, `srcset` and `<picture>` in HTML.

---

## Folder

```
04_Images_for_the_Web/
├── README.md
├── 01_image_formats_example.html      ← one photo as JPG, PNG, GIF, WebP, AVIF
├── 02_jpeg_quality_example.html       ← quality 90 / 75 / 50 / 10
├── 03_svg_vs_png_example.html         ← logo at 100px and 900px; transparency
├── 04_width_height_lazy_example.html  ← layout shift, loading="lazy"
├── 05_srcset_picture_example.html     ← srcset widths, <picture> formats
├── image_lab.html                     ← students' results page (Labs 1–6), submitted to Brightspace
├── my_images/                         ← students save their Pixlr exports here (see README.txt)
├── samples/                           ← generated sample images (beach + Harbourside logo)
├── Class7_Images_for_the_Web.pptx
└── _build/                            ← make_samples.js, build_deck.js
```

## Sample file sizes (from `sizes.json`)

Beach photo, 1118 × 640 (uncompressed: 2.0 MB):

| File | Size | Note |
|---|---|---|
| Original JPG | 134 KB | |
| JPG q90 | 140 KB | **Bigger** than the original: re-saving can't add detail back |
| JPG q75 / q50 / q10 | 93 / 51 / 12 KB | q10 shows 8 × 8 blocks |
| PNG | 978 KB | Lossless, so photos get huge |
| GIF (256 colours) | 407 KB | Banding in the sky |
| WebP q75 | 91 KB | |
| AVIF q50 | 69 KB | |
| JPG at 800 / 400 px wide | 45 / 13 KB | Resizing is the biggest saving |

Logo: SVG 0.7 KB · PNG 21 KB · JPG (q80) 21 KB. The JPG is no smaller, it's blurrier, and it loses the transparency.

---

## Deck outline (32 slides)

| # | Slide |
|---|---|
| 1–4 | Title · Attendance · Overview · Get ready (git pull, a big photo, Pixlr) |
| 5–8 | Why images matter · Pixels and dimensions (DPI note) · Colour depth · Do the math (uncompressed size) |
| 9 | **Lab 1:** get to know your photo (File Explorer → Properties → Details) |
| 10–12 | *Compression* divider · Lossless vs lossy · What lossy looks like (real zoomed crops) |
| 13–18 | *File formats* divider · Six formats · Comparison table · Example 01 · Raster vs vector · Example 03 |
| 19 | **Lab 2:** format shootout in Pixlr (PNG vs JPG vs WebP) |
| 20–21 | Example 02 · **Lab 3:** find the quality sweet spot |
| 22–23 | Size it right (display width × 2, target sizes) · **Lab 4:** resize and compare |
| 24–25 | **Lab 5:** graphics vs photos (logo) · Which format? decision chart |
| 26–28 | *Images in HTML* divider · Example 04 · Example 05 |
| 29 | **Lab 6:** put it on a page (image_lab.html, DevTools Network, submit) |
| 30–32 | How did you make out? · Web image checklist · Recap |

**Timing (about 2 hours):** about 40 min of concept slides, plus about 65 min across the six labs (5 + 15 + 10 + 10 + 5–10 + 15).

---

## The labs

Students save their exports into `my_images/` using the names in `my_images/README.txt`, and record results in `image_lab.html`.

| Lab | What students do | What they should find |
|---|---|---|
| 1 | Properties → Details for their phone photo; compute width × height × 3 | A 2–5 MB file vs 30–50 MB uncompressed (10–20× compression) |
| 2 | Save the same photo as PNG, JPG q80, WebP q80 | PNG is enormous (often bigger than the original); WebP is smallest |
| 3 | JPG at q30 and q10, zoom to 200–400% | 60–80 looks the same as the original at normal size; low quality shows blocks |
| 4 | Resize to 2000 / 1200 / 600 px (from the original each time) | Halving the width ≈ ¼ the file. 1200 px is often 10× smaller than the original |
| 5 | logo.png → JPG; a screenshot as PNG vs JPG | JPG loses transparency, smudges edges, and isn't even smaller |
| 6 | Finish image_lab.html, view with Live Server, DevTools → Network (Slow 4G) | Images that look the same on screen cost very different amounts to download |

**Submit (participation):** a screenshot of image_lab.html and the Network tab, uploaded to Brightspace.

### Before class: check Pixlr

- Use **pixlr.com → Pixlr Editor**. The Save dialog (Ctrl + S) lets you choose the format (JPG / PNG / WebP) and the quality, and you can set the size there too.
- Pixlr's menus change from time to time, and the free tier has at times **limited the number of saves per day**. Do a test run with a free account before class.
- **Backup:** [squoosh.app](https://squoosh.app). It's free, needs no account, and shows the new file size live; it handles format, quality and resize on one screen.
- iPhone photos may be **.HEIC**. Open them in Pixlr and save as JPG first, or set the phone to Settings → Camera → Formats → Most Compatible.
- Students with no photo can download a free full-size one from unsplash.com or pexels.com.

---

## Web image checklist (also the second-last slide)

- [ ] Keep the original. Always export copies.
- [ ] Right **format**: photo → JPG / WebP · transparency → PNG / WebP · logo → SVG
- [ ] Right **size**: the width it's shown at × 2, no bigger
- [ ] Right **quality**: 60–80 for photos. Zoom in and check
- [ ] Check the **file size**: content photos under ~150 KB, banners under ~300 KB
- [ ] In HTML: `alt`, `width` and `height`, and `loading="lazy"` below the fold
- [ ] Sensible lower-case file names: `harbour-view.jpg`, not `IMG_4032.JPG`
