/**
 * generate-previews.js
 *
 * Rebuilds the sample-page PNG for every workbook from its generated PDF.
 * The images are referenced by the book table in the root README.md and by
 * each project's README.md, so GitHub visitors can see a page without opening
 * the PDF.
 *
 * Uses pdf.js for rasterising and @napi-rs/canvas for the 2D context. Both are
 * pure npm packages with prebuilt binaries - no ImageMagick, Ghostscript or
 * Python needed.
 *
 * Usage (from the repository root, or anywhere):
 *   node scripts/generate-previews.js
 *   node scripts/generate-previews.js --page=4 --scale=2
 *   node scripts/generate-previews.js --book=Arabic_Stories_Workbook
 *
 * Options:
 *   --page=N    PDF page to capture (1-based, default: 3 - after cover + title)
 *   --scale=N   Render scale (default: 2 -> ~840x1191 px for an A5 page)
 *   --book=DIR  Limit generation to one project directory
 */

const fs = require("fs");
const path = require("path");
const { createCanvas } = require("@napi-rs/canvas");
const pdfjs = require("pdfjs-dist/legacy/build/pdf");

const ROOT = path.resolve(__dirname, "..");

function option(name, fallback) {
  const hit = process.argv.find(a => a.startsWith("--" + name + "="));
  return hit ? hit.slice(name.length + 3) : fallback;
}

const PAGE = Math.max(1, parseInt(option("page", "3"), 10) || 3);
const SCALE = Math.max(1, parseFloat(option("scale", "2")) || 2);
const ONLY = option("book", "");

/**
 * pdf.js expects a canvas factory with create/reset/destroy. The default
 * Node factory loads the native `canvas` package; this one uses @napi-rs/canvas
 * instead, which ships platform binaries inside the npm package.
 */
class NapiCanvasFactory {
  create(width, height) {
    if (width <= 0 || height <= 0) throw new Error("Invalid canvas size");
    const canvas = createCanvas(width, height);
    return { canvas, context: canvas.getContext("2d") };
  }
  reset(entry, width, height) {
    if (!entry.canvas) throw new Error("Canvas is not specified");
    entry.canvas.width = width;
    entry.canvas.height = height;
  }
  destroy(entry) {
    if (!entry.canvas) return;
    entry.canvas.width = 0;
    entry.canvas.height = 0;
    entry.canvas = null;
    entry.context = null;
  }
}

async function renderPreview(dir) {
  const pdfPath = path.join(ROOT, dir, dir + ".pdf");
  const outPath = path.join(ROOT, dir, "preview.png");

  if (!fs.existsSync(pdfPath)) {
    throw new Error("PDF not found: " + path.relative(ROOT, pdfPath));
  }

  const data = new Uint8Array(fs.readFileSync(pdfPath));
  const doc = await pdfjs.getDocument({ data, useSystemFonts: true }).promise;
  if (PAGE > doc.numPages) {
    throw new Error("page " + PAGE + " is out of range (" + doc.numPages + " pages)");
  }

  const page = await doc.getPage(PAGE);
  const viewport = page.getViewport({ scale: SCALE });
  const width = Math.round(viewport.width);
  const height = Math.round(viewport.height);
  const factory = new NapiCanvasFactory();
  const { canvas, context } = factory.create(width, height);

  await page.render({ canvasContext: context, viewport, canvasFactory: factory }).promise;
  fs.writeFileSync(outPath, canvas.toBuffer("image/png"));
  await doc.destroy();

  return {
    out: path.relative(ROOT, outPath),
    width,
    height,
    kb: Math.round(fs.statSync(outPath).size / 1024)
  };
}

(async () => {
  const books = JSON.parse(fs.readFileSync(path.join(ROOT, "books.json"), "utf8")).books;
  const targets = ONLY ? books.filter(b => b.dir === ONLY) : books;

  if (!targets.length) {
    console.error("No book matched --book=" + ONLY);
    process.exit(1);
  }

  console.log("جاري توليد الصور النموذجية (صفحة " + PAGE + "، مقياس " + SCALE + ") ...\n");
  let failed = 0;
  for (const book of targets) {
    try {
      const r = await renderPreview(book.dir);
      console.log("  [OK]   " + r.out + "  (" + r.width + "x" + r.height + ", " + r.kb + " KB)");
    } catch (e) {
      failed++;
      console.error("  [FAIL] " + book.dir + ": " + (e && e.message || e));
    }
  }
  if (failed) {
    console.error("\n" + failed + " preview(s) failed.");
    process.exit(1);
  }
  console.log("\nتم تحديث " + targets.length + " صورة نموذجية.");
})().catch(e => {
  console.error(e);
  process.exit(1);
});
