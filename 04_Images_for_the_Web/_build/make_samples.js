// Makes the sample images in ../samples (used by the example pages and the slides)
// and writes sizes.json so the slides show the real file sizes.
// To rebuild: cd 04_Images_for_the_Web/_build, run "npm install" once, then "npm run build".
const sharp = require("sharp");
const fs = require("fs");
const path = require("path");

const REPO = path.join(__dirname, "..");
const OUT = path.join(REPO, "samples");
const SCR = __dirname;
const SRC = path.join(REPO, "..", "02_Semantic_HTML_ARIA", "beach.JPG");

const LOGO_SVG = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 240 240" width="240" height="240">
	<!-- Harbourside Fish & Chips logo: a few flat shapes, so it's tiny as an SVG -->
	<title>Harbourside Fish &amp; Chips</title>
	<circle cx="120" cy="120" r="112" fill="#12304A"/>
	<ellipse cx="108" cy="104" rx="52" ry="30" fill="#E97132"/>
	<path d="M152 104 L196 74 L196 134 Z" fill="#E97132"/>
	<circle cx="82" cy="96" r="7" fill="#12304A"/>
	<path d="M38 166 q 22 -16 44 0 t 44 0 t 44 0 t 32 -6" fill="none" stroke="#0F9ED5" stroke-width="10" stroke-linecap="round"/>
	<path d="M58 196 q 20 -12 40 0 t 40 0 t 40 0" fill="none" stroke="#FFFFFF" stroke-width="6" stroke-linecap="round"/>
</svg>
`;

const kb = (f) => Math.round(fs.statSync(path.join(OUT, f)).size / 102.4) / 10;

async function main() {
	fs.mkdirSync(OUT, { recursive: true });
	const src = sharp(SRC);
	const meta = await src.metadata();

	// --- the beach photo in every format ---
	fs.copyFileSync(SRC, path.join(OUT, "beach_original.jpg"));
	await sharp(SRC).jpeg({ quality: 90, mozjpeg: true }).toFile(path.join(OUT, "beach_q90.jpg"));
	await sharp(SRC).jpeg({ quality: 75, mozjpeg: true }).toFile(path.join(OUT, "beach_q75.jpg"));
	await sharp(SRC).jpeg({ quality: 50, mozjpeg: true }).toFile(path.join(OUT, "beach_q50.jpg"));
	await sharp(SRC).jpeg({ quality: 10, mozjpeg: true }).toFile(path.join(OUT, "beach_q10.jpg"));
	await sharp(SRC).png({ compressionLevel: 9 }).toFile(path.join(OUT, "beach.png"));
	await sharp(SRC).gif({ colours: 256, dither: 0 }).toFile(path.join(OUT, "beach.gif"));
	await sharp(SRC).gif({ colours: 16, dither: 0 }).toFile(path.join(OUT, "beach_16_colours.gif"));
	await sharp(SRC).webp({ quality: 75 }).toFile(path.join(OUT, "beach.webp"));
	await sharp(SRC).avif({ quality: 50 }).toFile(path.join(OUT, "beach.avif"));

	// --- smaller copies for the srcset / resize examples ---
	for (const w of [800, 400]) {
		await sharp(SRC).resize({ width: w }).jpeg({ quality: 75, mozjpeg: true }).toFile(path.join(OUT, `beach_${w}.jpg`));
		await sharp(SRC).resize({ width: w }).webp({ quality: 75 }).toFile(path.join(OUT, `beach_${w}.webp`));
	}

	// --- the logo: SVG vs PNG vs JPG ---
	fs.writeFileSync(path.join(OUT, "logo.svg"), LOGO_SVG);
	await sharp(Buffer.from(LOGO_SVG), { density: 72 * 600 / 240 }).resize(600, 600).png({ compressionLevel: 9 }).toFile(path.join(OUT, "logo.png"));
	await sharp(path.join(OUT, "logo.png")).flatten({ background: "#ffffff" }).jpeg({ quality: 80 }).toFile(path.join(OUT, "logo.jpg"));

	// --- zoomed crops for the slides (nearest neighbour so you see the real pixels) ---
	const crop = { left: 520, top: 300, width: 120, height: 80 };
	const zoom = (file, out) => sharp(path.join(OUT, file)).extract(crop).resize(crop.width * 5, crop.height * 5, { kernel: "nearest" }).png().toFile(path.join(SCR, out));
	await zoom("beach_q90.jpg", "zoom_q90.png");
	await zoom("beach_q50.jpg", "zoom_q50.png");
	await zoom("beach_q10.jpg", "zoom_q10.png");
	const lcrop = { left: 140, top: 140, width: 160, height: 110 };
	const lzoom = (file, out) => sharp(path.join(OUT, file)).extract(lcrop).resize(lcrop.width * 4, lcrop.height * 4, { kernel: "nearest" })
		.flatten({ background: "#ffffff" }).png().toFile(path.join(SCR, out));
	await lzoom("logo.png", "zoom_logo_png.png");
	await lzoom("logo.jpg", "zoom_logo_jpg.png");
	// colour depth strip: 24-bit vs 256 vs 16 colours
	for (const f of ["beach_original.jpg", "beach.gif", "beach_16_colours.gif"]) {
		await sharp(path.join(OUT, f)).resize({ width: 560 }).png().toFile(path.join(SCR, "depth_" + f.replace(/\.\w+$/, ".png")));
	}

	const sizes = {
		width: meta.width, height: meta.height,
		original: kb("beach_original.jpg"),
		q90: kb("beach_q90.jpg"), q75: kb("beach_q75.jpg"), q50: kb("beach_q50.jpg"), q10: kb("beach_q10.jpg"),
		png: kb("beach.png"), gif: kb("beach.gif"), gif16: kb("beach_16_colours.gif"), webp: kb("beach.webp"), avif: kb("beach.avif"),
		w800: kb("beach_800.jpg"), w400: kb("beach_400.jpg"), w800webp: kb("beach_800.webp"), w400webp: kb("beach_400.webp"),
		logoSvg: kb("logo.svg"), logoPng: kb("logo.png"), logoJpg: kb("logo.jpg"),
		uncompressed: Math.round(meta.width * meta.height * 3 / 102.4) / 10,
	};
	fs.writeFileSync(path.join(SCR, "sizes.json"), JSON.stringify(sizes, null, 2));
	console.log(sizes);
}

main();
