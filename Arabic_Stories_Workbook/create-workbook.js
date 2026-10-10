#!/usr/bin/env node
/**
 * create-workbook.js
 *
 * Builds `Arabic_Stories_Workbook.docx` from the assets in `assets/` only --
 * it never reads source.docx, and no story text is hard-coded here.
 *
 *   assets/stories.json    ordered story index
 *   assets/texts/*.txt     line 1 = title, remaining lines = the sentences
 *   assets/images/*.jpg    one 3:2 picture per story
 *   assets/coverpage.png   full-bleed cover
 *
 * Usage: node create-workbook.js
 */

const fs = require("fs");
const path = require("path");
const {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  ImageRun,
  LineRuleType,
  Packer,
  PageNumber,
  PageOrientation,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  VerticalAlign,
  WidthType,
  convertInchesToTwip,
} = require("docx");

// Set to false to hide the page number footer (Arabic_Alphabet_Workbook style).
const SHOW_PAGE_NUMBERS = false;

// A5 dimensions in twips (14.8cm x 21cm)
const A5_WIDTH = convertInchesToTwip(5.83);
const A5_HEIGHT = convertInchesToTwip(8.27);

const MARGIN_TOP = convertInchesToTwip(0.35);
const MARGIN_BOTTOM = convertInchesToTwip(0.35);
const MARGIN_SIDE = convertInchesToTwip(0.5);

const CONTENT_WIDTH = A5_WIDTH - MARGIN_SIDE * 2; // 6955 twips (~122.6 mm)

const ROOT = __dirname;
const ASSETS = path.join(ROOT, "assets");
const MANIFEST = path.join(ASSETS, "stories.json");
const OUTPUT = path.join(ROOT, "Arabic_Stories_Workbook.docx");

const FONT = "Traditional Arabic";

const COLOR_TITLE = "2E86AB";
const COLOR_MAIN = "1B4965";
const COLOR_BORDER = "34495E";
const COLOR_MUTED = "7F8C8D";

// docx font sizes are in HALF-POINTS, so 48 = 24pt, not 48pt.
const TITLE_SIZE = 48; // 24pt - titles run 3-5 words
const SENTENCE_SIZE = 40; // 20pt
const SENTENCE_LINE = 480; // twips, AT_LEAST (24pt)
const SENTENCE_AFTER = 120; // twips
const TITLE_AFTER = 200; // twips

// Story picture: contained inside a 3:2 box, 112 mm wide.
const MM_TO_PX = 96 / 25.4;
const TWIPS_PER_PX = 15; // 1px at 96dpi = 1/96 inch = 15 twips
const IMAGE_BOX_PX = Math.round(112 * MM_TO_PX);
const IMAGE_BOX_RATIO = 3 / 2;
const IMAGE_FRAME_PADDING_PX = 10;
const BORDER_SIZE = 8;

const TITLE_PAGE_SIZE = 96; // 48pt

// ---------------------------------------------------------------------------
// Image dimensions
// ---------------------------------------------------------------------------

/** Read width/height from a PNG buffer (IHDR at bytes 16-23). */
function getPngDimensions(buffer) {
  if (buffer.length < 24 || buffer.slice(1, 4).toString() !== "PNG") {
    throw new Error("Not a PNG buffer");
  }
  return { width: buffer.readUInt32BE(16), height: buffer.readUInt32BE(20) };
}

/**
 * Read width/height from a JPEG buffer by walking the marker segments to the
 * first SOF frame header. Needed because the story pictures are JPEG; the
 * siblings only ever embedded PNGs and so only ship getPngDimensions().
 */
function getJpegDimensions(buffer) {
  if (buffer[0] !== 0xff || buffer[1] !== 0xd8) {
    throw new Error("Not a JPEG buffer");
  }

  let offset = 2;
  while (offset + 9 < buffer.length) {
    if (buffer[offset] !== 0xff) {
      offset++;
      continue;
    }
    const marker = buffer[offset + 1];

    // Standalone markers carry no length field.
    if (marker === 0x01 || (marker >= 0xd0 && marker <= 0xd9)) {
      offset += 2;
      continue;
    }
    // SOF0-SOF15, minus DHT (c4), JPGA (c8) and DAC (cc).
    if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
      return { height: buffer.readUInt16BE(offset + 5), width: buffer.readUInt16BE(offset + 7) };
    }
    offset += 2 + buffer.readUInt16BE(offset + 2);
  }
  throw new Error("JPEG has no SOF marker");
}

/** Scale to fit inside a box, preserving aspect ratio. Never upscales. */
function fitImage(buffer, maxWidth, maxHeight) {
  const dims = getJpegDimensions(buffer);
  const scale = Math.min(maxWidth / dims.width, maxHeight / dims.height, 1);
  return {
    width: Math.max(1, Math.round(dims.width * scale)),
    height: Math.max(1, Math.round(dims.height * scale)),
  };
}

function formatBytes(bytes) {
  return bytes >= 1024 * 1024
    ? `${(bytes / 1024 / 1024).toFixed(2)} MB`
    : `${(bytes / 1024).toFixed(0)} KB`;
}

// ---------------------------------------------------------------------------
// Asset loading
// ---------------------------------------------------------------------------

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

/**
 * Read a story text file: line 1 is the title, the rest are the sentences.
 * A leading BOM would otherwise corrupt the title, so strip it defensively.
 */
function readStoryText(relativePath) {
  const absolute = path.join(ASSETS, relativePath);
  assert(fs.existsSync(absolute), `Missing text asset: assets/${relativePath}`);

  const lines = fs
    .readFileSync(absolute, "utf-8")
    .replace(/^\uFEFF/, "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0);

  assert(lines.length >= 2, `assets/${relativePath} has no sentences`);
  return { title: lines[0], sentences: lines.slice(1) };
}

function loadManifest() {
  assert(fs.existsSync(MANIFEST), "Missing assets/stories.json - run: npm run extract");

  const manifest = JSON.parse(fs.readFileSync(MANIFEST, "utf-8").replace(/^\uFEFF/, ""));
  assert(manifest.stories.length > 0, "assets/stories.json lists no stories");
  return manifest;
}

function loadStories(manifest) {
  return manifest.stories.map((story, index) => {
    const text = readStoryText(story.text);
    const imagePath = path.join(ASSETS, story.image);
    assert(fs.existsSync(imagePath), `Missing image asset: assets/${story.image}`);

    const image = fs.readFileSync(imagePath);
    return {
      id: story.id || index + 1,
      word: story.word,
      title: text.title,
      sentences: text.sentences,
      image,
      imageBytes: image.length,
    };
  });
}

/** assets/coverpage.png first, then the shared cover in covers-and-photos/ as a fallback. */
function loadCover() {
  const candidates = [
    path.join(ASSETS, "coverpage.png"),
    path.join(ROOT, "..", "covers-and-photos", "cover-arabic.png"),
  ];
  const coverPath = candidates.find((candidate) => fs.existsSync(candidate));
  assert(coverPath, "No cover found - expected assets/coverpage.png");

  const buffer = fs.readFileSync(coverPath);
  return {
    buffer,
    isPlaceholder: path.basename(coverPath) === "coverpage.png",
    path: path.relative(ROOT, coverPath),
  };
}

// ---------------------------------------------------------------------------
// Page builders
// ---------------------------------------------------------------------------

function createPageFooter() {
  if (!SHOW_PAGE_NUMBERS) {
    return null;
  }
  return new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        bidirectional: true,
        children: [
          new TextRun({
            children: [PageNumber.CURRENT],
            font: FONT,
            size: 20,
            color: COLOR_BORDER,
          }),
        ],
      }),
    ],
  });
}

function a5SectionProperties(margins) {
  // NB: docx 9.7.x has no section-level `direction` option, so section RTL is
  // set per paragraph with `bidirectional: true` (which emits <w:bidi/>).
  return {
    page: {
      size: {
        width: A5_WIDTH,
        height: A5_HEIGHT,
        orientation: PageOrientation.PORTRAIT,
      },
      margin: margins,
    },
  };
}

const PAGE_MARGINS = {
  top: MARGIN_TOP,
  right: MARGIN_SIDE,
  bottom: MARGIN_BOTTOM,
  left: MARGIN_SIDE,
};

const FULL_BLEED_MARGINS = { top: 0, right: 0, bottom: 0, left: 0 };

/** Full A5 page, cover-scaled so the picture overspills all four edges. */
function createCoverPage(cover) {
  const dims = getPngDimensions(cover.buffer);

  // Full A5 page in pixels at 96dpi (no margins for full-bleed cover)
  const pageWidthPx = A5_WIDTH / TWIPS_PER_PX;
  const pageHeightPx = A5_HEIGHT / TWIPS_PER_PX;

  const scale = Math.max(pageWidthPx / dims.width, pageHeightPx / dims.height);
  const widthPx = Math.round(dims.width * scale);
  const heightPx = Math.round(dims.height * scale);

  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 0 },
      children: [
        new ImageRun({
          data: cover.buffer,
          transformation: { width: widthPx, height: heightPx },
          type: "png",
        }),
      ],
    }),
  ];
}

function createTitlePage(book) {
  const rule = {
    bottom: { style: BorderStyle.SINGLE, size: 12, color: COLOR_BORDER, space: 8 },
  };

  return [
    new Paragraph({ spacing: { before: 2400, after: 0 }, children: [] }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      bidirectional: true,
      spacing: { before: 0, after: 160 },
      border: rule,
      children: [
        new TextRun({
          text: book.title,
          font: FONT,
          size: TITLE_PAGE_SIZE,
          color: COLOR_MAIN,
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      bidirectional: true,
      spacing: { before: 320, after: 0 },
      children: [
        new TextRun({
          text: book.subtitle,
          font: FONT,
          size: 48,
          color: COLOR_TITLE,
        }),
      ],
    }),
  ];
}

/** Story picture, contained in a 3:2 box inside a thin ruled frame. */
function createImageFrame(jpeg) {
  const size = fitImage(jpeg, IMAGE_BOX_PX, Math.round(IMAGE_BOX_PX / IMAGE_BOX_RATIO));
  const cellWidth = (size.width + IMAGE_FRAME_PADDING_PX * 2) * TWIPS_PER_PX;

  const border = { style: BorderStyle.SINGLE, size: BORDER_SIZE, color: COLOR_BORDER };

  return new Table({
    width: { size: cellWidth, type: WidthType.DXA },
    columnWidths: [cellWidth],
    alignment: AlignmentType.CENTER,
    borders: { top: border, bottom: border, left: border, right: border },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: cellWidth, type: WidthType.DXA },
            verticalAlign: VerticalAlign.CENTER,
            margins: { top: 60, bottom: 60, left: 60, right: 60 },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 0, after: 0 },
                bidirectional: true,
                children: [
                  new ImageRun({
                    data: jpeg,
                    transformation: { width: size.width, height: size.height },
                    type: "jpg",
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

/**
 * One story per page: title, then the framed picture, then the sentences.
 * Nothing here is bold.
 */
function createStoryPage(story) {
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      bidirectional: true,
      spacing: { before: 0, after: TITLE_AFTER, line: 520, lineRule: LineRuleType.AT_LEAST },
      children: [
        new TextRun({
          text: story.title,
          font: FONT,
          size: TITLE_SIZE,
          color: COLOR_MAIN,
        }),
      ],
    }),
    createImageFrame(story.image),
    new Paragraph({ spacing: { before: 0, after: 0 }, children: [] }),
    ...story.sentences.map(
      (sentence) =>
        new Paragraph({
          alignment: AlignmentType.CENTER,
          bidirectional: true,
          spacing: {
            before: 0,
            after: SENTENCE_AFTER,
            line: SENTENCE_LINE,
            lineRule: LineRuleType.AT_LEAST,
          },
          children: [
            new TextRun({
              text: sentence,
              font: FONT,
              size: SENTENCE_SIZE,
              color: COLOR_MAIN,
            }),
          ],
        })
    ),
  ];
}

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function createWorkbook() {
  console.log("جاري إنشاء كتاب القصص المصورة...");

  const manifest = loadManifest();
  const stories = loadStories(manifest);
  const cover = loadCover();

  if (cover.isPlaceholder) {
    console.log(`Warning: using placeholder cover (${cover.path}) - replace it with the real cover.`);
  }
  console.log(`cover : ${cover.path} (${formatBytes(cover.buffer.length)})`);
  console.log(`stories: ${stories.length}`);

  const doc = new Document({
    creator: "Arabic Learning Workbook",
    title: manifest.book.title,
    description: manifest.book.subtitle,
    styles: {
      default: {
        document: {
          run: { font: FONT, size: 28 },
        },
      },
    },
    sections: [],
  });

  // Cover: full-bleed, no footer or page number.
  doc.addSection({
    properties: a5SectionProperties(FULL_BLEED_MARGINS),
    children: createCoverPage(cover),
  });

  doc.addSection({
    properties: a5SectionProperties(PAGE_MARGINS),
    ...(SHOW_PAGE_NUMBERS ? { footers: { default: createPageFooter() } } : {}),
    children: createTitlePage(manifest.book),
  });

  const footers = SHOW_PAGE_NUMBERS ? { footers: { default: createPageFooter() } } : {};

  for (const story of stories) {
    console.log(`جاري إنشاء صفحة: ${story.title}`);
    doc.addSection({
      properties: a5SectionProperties(PAGE_MARGINS),
      ...footers,
      children: createStoryPage(story),
    });
  }

  console.log("جاري حفظ الملف...");

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(OUTPUT, buffer);

  const imageBytes = stories.reduce((total, story) => total + story.imageBytes, 0);

  console.log("تم إنشاء الملف بنجاح!");
  console.log("الملف: Arabic_Stories_Workbook.docx");
  console.log(`عدد الصفحات: ${stories.length + 2} صفحة`);
  console.log(`عدد القصص: ${stories.length}`);
  console.log(`حجم الصور: ${formatBytes(imageBytes)}`);
  console.log(`حجم الملف: ${formatBytes(buffer.length)}`);
  console.log("حجم الصفحة: A5");
  console.log(`أرقام الصفحات: ${SHOW_PAGE_NUMBERS ? "معروضة" : "مخفية"}`);
}

createWorkbook().catch((error) => {
  console.error(`create-workbook failed: ${error.message}`);
  process.exit(1);
});
