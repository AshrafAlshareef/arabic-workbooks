/*
    images-to-book.js
    Builds a book out of every image in the images/ folder sitting next to this
    script: one full-bleed A5 page per image, then the same Word conversion the
    four workbook projects use.

    Usage:
        node images-to-book.js [--size a5|a4] [--no-pdf] [--images <dir>] [--out <dir>]

    Defaults keep the original workflow: images/ next to this script, output
    written next to this script. --images and --out let the GUI point the tool
    at another folder without touching the defaults.

    Writes:
        images-to-book.docx
        images-to-book.pdf      (unless --no-pdf)
*/

const docx = require("docx");
const fs = require("fs");
const path = require("path");
const { spawnSync } = require("child_process");

const {
  Document,
  Packer,
  Paragraph,
  ImageRun,
  AlignmentType,
  PageOrientation,
  convertInchesToTwip,
} = docx;

// ---------------------------------------------------------------------------
// Paths
// ---------------------------------------------------------------------------
const ROOT = __dirname;
let IMAGES_DIR = path.join(ROOT, "images");
let OUTPUT_DOCX = path.join(ROOT, "images-to-book.docx");
let OUTPUT_PDF = path.join(ROOT, "images-to-book.pdf");

// A relative --images / --out is relative to the folder the user ran the tool
// from, not to this folder: images-to-book.bat changes into its own directory
// before starting node, so it passes the caller's folder in this variable.
const USER_CWD = process.env.IMAGES_TO_BOOK_CWD || process.cwd();

/** Path as shown in logs: relative when inside the repo, absolute otherwise. */
function displayPath(target) {
  const relative = path.relative(ROOT, target);
  return relative.startsWith("..") || path.isAbsolute(relative) ? target : relative;
}

// The converter is the one the repo already ships, so this tool never grows a
// second copy of the Word COM logic.
const CONVERTER = path.join(ROOT, "..", "convert-one.ps1");

// ---------------------------------------------------------------------------
// Page geometry
// ---------------------------------------------------------------------------

// Word lays an image out in pixels, and 1 px at 96 dpi is 15 twips, so
// twips/15 is the pixel width of the paper. Same divisor the book projects use,
// which is what makes a cover here print identically to a cover there.
const TWIPS_PER_PX = 15;

const SIZES = {
  a5: { width: 5.83, height: 8.27 },
  a4: { width: 8.27, height: 11.69 },
};

const EXTENSIONS = new Set([".png", ".jpg", ".jpeg"]);

/**
 * How much of a cover image may be cropped before the tool stops bleeding it.
 *
 * The threshold is on the cropped FRACTION, not on the aspect ratio, because
 * the two are not the same question. A 1024x1536 cover is 5.4% off the A5 page
 * ratio and the book covers deliberately bleed (crop) exactly that much, so a
 * plain "is the ratio close enough" test would have rejected the very images
 * this tool exists to reproduce. A 1536x1024 landscape image would instead lose
 * 53% of its width, so anything past a small crop has to be fitted whole.
 */
const MAX_BLEED_CROP = 0.08;

// ---------------------------------------------------------------------------
// Arguments
// ---------------------------------------------------------------------------
function parseArgs(argv) {
  const options = { size: "a5", pdf: true, images: null, out: null };

  for (let i = 0; i < argv.length; i += 1) {
    const arg = String(argv[i]);
    const equals = arg.indexOf("=");
    const flag = equals === -1 ? arg : arg.slice(0, equals);
    const inlineValue = equals === -1 ? null : arg.slice(equals + 1);

    const takeValue = () => {
      if (inlineValue !== null) return inlineValue;
      i += 1;
      if (i >= argv.length) throw new Error(`${flag} needs a value.`);
      return String(argv[i]);
    };

    if (flag === "--no-pdf") {
      options.pdf = false;
    } else if (flag === "--size") {
      const value = takeValue().toLowerCase();
      if (!SIZES[value]) {
        throw new Error(
          `--size must be one of: ${Object.keys(SIZES).join(", ")} (got "${value}")`
        );
      }
      options.size = value;
    } else if (flag === "--images") {
      options.images = takeValue();
    } else if (flag === "--out") {
      options.out = takeValue();
    } else if (flag === "--help" || flag === "-h") {
      console.log(
        "Usage: node images-to-book.js [--size a5|a4] [--no-pdf] [--images <dir>] [--out <dir>]"
      );
      process.exit(0);
    } else {
      throw new Error(`Unknown option: ${arg}`);
    }
  }

  return options;
}

// ---------------------------------------------------------------------------
// Collecting the images
// ---------------------------------------------------------------------------
function toPosix(relative) {
  return relative.split(path.sep).join("/");
}

/**
 * Every image under `dir`, subfolders included, in reading order.
 *
 * Subfolders are walked rather than skipped: dropping a batch of covers into
 * images/2026/ should not need a different command. The sort is a natural
 * collation on the relative path, so 1- comes before 2- before 10- the way a
 * file listing should read, matching the 01-asad.jpg naming already used by the
 * stories project.
 */
function collectImages(dir) {
  const found = [];

  const walk = (current) => {
    const entries = fs.readdirSync(current, { withFileTypes: true });
    for (const entry of entries) {
      const full = path.join(current, entry.name);
      if (entry.isDirectory()) {
        walk(full);
      } else if (entry.isFile() && EXTENSIONS.has(path.extname(entry.name).toLowerCase())) {
        found.push(full);
      }
    }
  };

  walk(dir);

  const collator = new Intl.Collator("en", { numeric: true, sensitivity: "base" });
  return found.sort((a, b) =>
    collator.compare(toPosix(path.relative(dir, a)), toPosix(path.relative(dir, b)))
  );
}

// ---------------------------------------------------------------------------
// Reading image dimensions
//
// No image library is installed, so each format is measured from its header.
// ---------------------------------------------------------------------------

/** PNG width/height, from the IHDR chunk at bytes 16-23. */
function pngDimensions(buffer) {
  const SIGNATURE = "89504e470d0a1a0a";
  if (buffer.length < 24 || buffer.toString("hex", 0, 8) !== SIGNATURE) {
    throw new Error("not a PNG (bad signature)");
  }
  const width = buffer.readUInt32BE(16);
  const height = buffer.readUInt32BE(20);
  if (width === 0 || height === 0) {
    throw new Error("PNG reports a zero dimension");
  }
  return { width, height };
}

/**
 * JPEG width/height, from the first SOF (start of frame) marker.
 *
 * The marker payload is precision, height, width, so height sits 5 bytes in and
 * width 7 bytes in. Every other segment is skipped using its own 2-byte length,
 * and the entropy-coded scan data cannot be walked, which is why the search
 * stops at the first SOF rather than looking further.
 */
const SOF_MARKERS = new Set([
  0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf,
]);

function jpegDimensions(buffer) {
  if (buffer.length < 4 || buffer[0] !== 0xff || buffer[1] !== 0xd8) {
    throw new Error("not a JPEG (missing SOI)");
  }

  let offset = 2;
  while (offset + 3 < buffer.length) {
    if (buffer[offset] !== 0xff) {
      offset += 1;
      continue;
    }
    const marker = buffer[offset + 1];
    // 0xFF fill bytes repeat before a real marker, and the standalone markers
    // (SOI, TEM, RSTn) carry no payload at all.
    if (marker === 0xff) {
      offset += 1;
      continue;
    }
    if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) {
      offset += 2;
      continue;
    }
    const length = buffer.readUInt16BE(offset + 2);
    if (SOF_MARKERS.has(marker)) {
      const height = buffer.readUInt16BE(offset + 5);
      const width = buffer.readUInt16BE(offset + 7);
      if (width === 0 || height === 0) {
        throw new Error("JPEG reports a zero dimension");
      }
      return { width, height };
    }
    offset += 2 + length;
  }

  throw new Error("no SOF marker found");
}

function measure(file, buffer) {
  return path.extname(file).toLowerCase() === ".png"
    ? pngDimensions(buffer)
    : jpegDimensions(buffer);
}

// ---------------------------------------------------------------------------
// Fitting an image to the page
// ---------------------------------------------------------------------------

/**
 * Pick between bleeding and fitting, then size the image for that choice.
 *
 * "bleed" repeats the book cover maths exactly: scale by the larger of the two
 * ratios so the picture overspills all four edges and nothing is left showing.
 * "fit" scales by the smaller ratio so the whole picture survives, and the caller
 * centres it on the leftover paper.
 */
function fitImage(image, page) {
  const imageRatio = image.width / image.height;
  const pageRatio = page.widthPx / page.heightPx;
  const crop = 1 - Math.min(1, imageRatio / pageRatio, pageRatio / imageRatio);

  if (crop <= MAX_BLEED_CROP) {
    const scale = Math.max(page.widthPx / image.width, page.heightPx / image.height);
    return {
      mode: "bleed",
      widthPx: image.width * scale,
      heightPx: image.height * scale,
      crop,
    };
  }

  const scale = Math.min(page.widthPx / image.width, page.heightPx / image.height);
  return { mode: "fit", widthPx: image.width * scale, heightPx: image.height * scale, crop };
}

/**
 * One image, one section, one page.
 *
 * All four margins must be zero: a bled image is deliberately oversized and
 * centred so it runs off every edge, and a leftover side margin would shift it
 * sideways and leave a white band. The page frame is skipped too, so the page
 * stays a pure full-bleed image with nothing else on it.
 */
function createImagePage(page, fit, buffer, type) {
  // A bled image fills the paper, so it starts at the very top. A fitted one
  // leaves bands above and below, and gets half the slack as space before the
  // paragraph so it sits centred instead of hanging off the top edge.
  const beforeTwips =
    fit.mode === "fit"
      ? Math.round(((page.heightPx - fit.heightPx) / 2) * TWIPS_PER_PX)
      : 0;

  return {
    properties: {
      page: {
        size: {
          width: page.widthTwips,
          height: page.heightTwips,
          orientation: PageOrientation.PORTRAIT,
        },
        margin: { top: 0, right: 0, bottom: 0, left: 0 },
      },
      direction: "RTL",
    },
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: beforeTwips, after: 0 },
        children: [
          new ImageRun({
            data: buffer,
            transformation: {
              width: Math.round(fit.widthPx),
              height: Math.round(fit.heightPx),
            },
            type,
          }),
        ],
      }),
    ],
  };
}

// ---------------------------------------------------------------------------
// PDF conversion
// ---------------------------------------------------------------------------
function formatBytes(bytes) {
  return `${(bytes / 1024 / 1024).toFixed(2)} MB`;
}

function convertToPdf() {
  if (!fs.existsSync(CONVERTER)) {
    throw new Error(
      `converter not found: ${path.relative(ROOT, CONVERTER)} - it ships one folder up`
    );
  }

  console.log("جاري تحويل الكتاب إلى PDF عبر Microsoft Word...");
  const started = Date.now();

  const result = spawnSync(
    "powershell.exe",
    [
      "-NoProfile",
      "-ExecutionPolicy",
      "Bypass",
      "-File",
      CONVERTER,
      "-Docx",
      OUTPUT_DOCX,
      "-Pdf",
      OUTPUT_PDF,
    ],
    { stdio: "inherit" }
  );

  if (result.error) {
    throw new Error(`cannot launch powershell.exe: ${result.error.message}`);
  }

  const seconds = ((Date.now() - started) / 1000).toFixed(1);
  if (result.status === 2) {
    throw new Error("Microsoft Word is not available for automation.");
  }
  if (result.status !== 0 || !fs.existsSync(OUTPUT_PDF)) {
    throw new Error(`PDF conversion failed (exit code ${result.status}). Close Word and retry.`);
  }

  console.log(
    `  [OK]   ${path.basename(OUTPUT_PDF)} (${formatBytes(fs.statSync(OUTPUT_PDF).size)}, ${seconds}s)`
  );
}

// ---------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------
async function buildBook(options) {
  const size = SIZES[options.size];
  const page = {
    widthTwips: convertInchesToTwip(size.width),
    heightTwips: convertInchesToTwip(size.height),
  };
  page.widthPx = page.widthTwips / TWIPS_PER_PX;
  page.heightPx = page.heightTwips / TWIPS_PER_PX;
  page.label = options.size.toUpperCase();

  if (!fs.existsSync(IMAGES_DIR)) {
    throw new Error(`no images folder: ${displayPath(IMAGES_DIR)}`);
  }

  const files = collectImages(IMAGES_DIR);
  if (files.length === 0) {
    throw new Error(
      `no .png/.jpg/.jpeg images in ${displayPath(IMAGES_DIR)} (subfolders are included)`
    );
  }

  console.log(`\nمجلد الصور: ${displayPath(IMAGES_DIR)}`);
  console.log(`  عدد الصور: ${files.length}`);
  console.log(
    `  الصفحة: ${page.label} (${Math.round(page.widthPx)} x ${Math.round(page.heightPx)} بكسل)\n`
  );

  const doc = new Document({
    creator: "image_2_cover_style",
    title: "كتاب الصور",
    description: "كتاب مُولَّد من مجلد images بأسلوب صفحات الأغلفة",
    styles: {
      default: {
        document: {
          run: { font: "Traditional Arabic", size: 28 },
        },
      },
    },
    sections: [],
  });

  let added = 0;
  let skipped = 0;

  for (const [index, file] of files.entries()) {
    const name = toPosix(path.relative(IMAGES_DIR, file));
    const number = String(index + 1).padStart(String(files.length).length, " ");
    let buffer;
    let image;
    try {
      buffer = fs.readFileSync(file);
      image = measure(file, buffer);
    } catch (err) {
      skipped += 1;
      console.log(`  [SKIP] ${number}/${files.length}  ${name}  - ${err.message}`);
      continue;
    }

    const fit = fitImage(image, page);
    const type = path.extname(file).toLowerCase() === ".png" ? "png" : "jpg";

    doc.addSection(createImagePage(page, fit, buffer, type));
    added += 1;

    const printed =
      `${Math.round(fit.widthPx)}x${Math.round(fit.heightPx)} بكسل` +
      (fit.mode === "bleed" ? "  (full-bleed)" : `  (مُحتوى كامل، بدون قص)`);
    console.log(`  [OK]   ${number}/${files.length}  ${name}  ${image.width}x${image.height}  ->  ${printed}`);
  }

  if (added === 0) {
    throw new Error("every image failed to load, so there is nothing to build.");
  }
  if (skipped > 0) {
    console.log(`\n  تم تخطي ${skipped} صورة تعذّرت قراءتها.`);
  }

  console.log("\nجاري حفظ الملف...");
  const bytes = await Packer.toBuffer(doc);
  fs.mkdirSync(path.dirname(OUTPUT_DOCX), { recursive: true });
  fs.writeFileSync(OUTPUT_DOCX, bytes);
  console.log(`تم إنشاء الملف بنجاح: ${path.basename(OUTPUT_DOCX)} (${formatBytes(bytes.length)})`);
  console.log(`عدد الصفحات: ${added} صفحة`);

  if (options.pdf) {
    convertToPdf();
  }

  console.log(`\nتم: ${added} صورة  ->  ${path.basename(OUTPUT_DOCX)}${options.pdf ? ` + ${path.basename(OUTPUT_PDF)}` : ""}\n`);
}

if (require.main === module) {
  let options;
  try {
    options = parseArgs(process.argv.slice(2));
  } catch (err) {
    console.error(`[FAIL] ${err.message}`);
    console.error(
      "الاستخدام: images-to-book.bat [--size a5|a4] [--no-pdf] [--images <dir>] [--out <dir>]"
    );
    process.exit(1);
  }

  if (options.images) IMAGES_DIR = path.resolve(USER_CWD, options.images);
  if (options.out) {
    const outDir = path.resolve(USER_CWD, options.out);
    OUTPUT_DOCX = path.join(outDir, "images-to-book.docx");
    OUTPUT_PDF = path.join(outDir, "images-to-book.pdf");
  }

  buildBook(options).catch((err) => {
    console.error(`\n[FAIL] ${err.message}`);
    process.exitCode = 1;
  });
}
