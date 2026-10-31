const docx = require("docx");
const fs = require("fs");
const path = require("path");
const svg2img = require("svg2img");

const { parseBookArgs, USAGE } = require("../build-options.js");
const { sortByLetter } = require("../letter-order.js");

const CLI = parseBookArgs(process.argv.slice(2));
if (CLI.help) {
  console.log(USAGE);
  process.exit(0);
}

const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  HeightRule,
  AlignmentType,
  BorderStyle,
  PageOrientation,
  convertInchesToTwip,
  VerticalAlign,
  ImageRun,
  PageNumber,
  Footer,
  PageBorderDisplay,
  PageBorderOffsetFrom,
  PageBorderZOrder,
} = docx;

// ---------------------------------------------------------------------------
// Page geometry (A5 portrait) - twips
// ---------------------------------------------------------------------------
const A5_WIDTH = convertInchesToTwip(5.83); // 14.8 cm
const A5_HEIGHT = convertInchesToTwip(8.27); // 21 cm

const MARGIN_TOP = convertInchesToTwip(0.35);
const MARGIN_BOTTOM = convertInchesToTwip(0.35);
const MARGIN_SIDE = convertInchesToTwip(0.5);

const CONTENT_WIDTH = A5_WIDTH - MARGIN_SIDE * 2; // 6955 twips (~122.6 mm)
const CONTENT_HEIGHT = A5_HEIGHT - MARGIN_TOP - MARGIN_BOTTOM; // 10901 twips (~192.2 mm)

// ---------------------------------------------------------------------------
// Shared constants
// ---------------------------------------------------------------------------
const FONT = "Traditional Arabic";

// Set to false to hide the page number footer
const SHOW_PAGE_NUMBERS = false;

const COLOR_TITLE = "2E86AB";
const COLOR_MAIN = "1B4965";
const COLOR_LETTER = "E74C3C";
const COLOR_BORDER = "34495E";
const COLOR_LINE = "2C3E50";

// Simple single-line frame around every content page. The offset is measured
// from the paper edge (18pt), so the frame always sits outside the text area
// regardless of how tight a section's margins are.
const PAGE_FRAME = {
  pageBorders: {
    display: PageBorderDisplay.ALL_PAGES,
    offsetFrom: PageBorderOffsetFrom.PAGE,
    zOrder: PageBorderZOrder.FRONT,
  },
  pageBorderTop: { style: BorderStyle.SINGLE, size: 6, color: COLOR_BORDER, space: 18 },
  pageBorderRight: { style: BorderStyle.SINGLE, size: 6, color: COLOR_BORDER, space: 18 },
  pageBorderBottom: { style: BorderStyle.SINGLE, size: 6, color: COLOR_BORDER, space: 18 },
  pageBorderLeft: { style: BorderStyle.SINGLE, size: 6, color: COLOR_BORDER, space: 18 },
};

// A letter box is not square any more: a 28pt Arabic glyph needs more vertical
// room than a 14pt one, so width and height are separate. Five boxes (the
// longest word, صندوق) must fit across the boxes column, so BOX_W is the
// binding constraint on that column's width.
const BOX_W = 780; // twips (13.75 mm)
const BOX_H = 950; // twips (16.8 mm)
const BOX_BORDER_SIZE = 8;

// Image box in mm, converted to pixels at 96 dpi
const MM_TO_PX = 96 / 25.4;
const IMAGE_BOX_MM = 22;
const IMAGE_BOX_PX = Math.round(IMAGE_BOX_MM * MM_TO_PX);
const SVG_RENDER_WIDTH = 400; // px, same value the Alphabet book renders at

// One exercise row height (a floor; the real height comes from the content).
// 5 stacked rows on a 545pt text area: the measured cliff is at ~92pt (5 rows
// no longer fit and Word spills one onto a new page), so stay at 89pt.
const ROW_HEIGHT = 1780; // twips (~89 pt)

// Gap between exercises: an exact 1pt line plus 8pt of space after, so the
// gap stays deterministic instead of inheriting the body font size.
const GAP_LINE = 20; // twips
const GAP_AFTER = 160; // twips

// Page header, kept small so five rows always fit
const HEADER_SIZE = 36; // 18pt
const HEADER_AFTER = 60; // twips

// Word label and letters.
// NOTE: docx font sizes are in HALF-POINTS, so 56 = 28pt, not 28pt.
const WORD_SIZE = 56; // 28pt
const WORD_LINE = 950; // twips, at-least so Arabic is never clipped
const LETTER_SIZE = 56; // 28pt, same as the word name in the split section
const LETTER_LINE = 950; // twips, at-least

// Writing line
const WRITE_LINE_HEIGHT = 560; // twips
const WRITE_LINE_INNER = 500; // twips, exact (cell is empty)

// A 1pt paragraph used to keep two adjacent tables from merging into one
const EXERCISES_PER_PAGE = 5;

const ROOT = __dirname;
const ASSETS = path.join(ROOT, "assets");
const WORDS_FILE = path.join(ASSETS, "words.json");
const EXERCISES_FILE = path.join(ASSETS, "exercise-types.json");
// Named after the project folder, so a copy of this project in a differently
// named folder writes to its own name instead of overwriting another book.
const OUT_DIR = CLI.out || ROOT;
const OUTPUT = path.join(OUT_DIR, `${path.basename(ROOT)}.docx`);

// All exercises are stacked in one vertical column, one full-width row each.
// Every row is a single table of three side-by-side cells, listed
// right-to-left because the table is visuallyRightToLeft, so index 0 is the
// right-most cell on the page:
//
//   split     : صورة | كلمة | مربعات
//   contextual: صورة | مربعات | خط
//   compose   : صورة | مربعات | خط
//
// The boxes column must stay wide enough for the longest word. The longest
// words (زرافة, ليمون, صندوق) need 5 boxes = 5 * 780 = 3900 twips, which is
// 56% of CONTENT_WIDTH, so 58% is the floor and going narrower clips the strip.
//
// The image column shrank to 19% (image is 22 mm) and the word/line column grew
// to 23%, because a 28pt word name no longer fits in the old 18%.
const LAYOUTS = {
  split: [
    { kind: "image", width: 0.19 },
    { kind: "word", width: 0.23 },
    { kind: "boxes", width: 0.58 },
  ],
  contextual: [
    { kind: "image", width: 0.19 },
    { kind: "boxes", width: 0.58 },
    { kind: "line", width: 0.23 },
  ],
  compose: [
    { kind: "image", width: 0.19 },
    { kind: "boxes", width: 0.58 },
    { kind: "line", width: 0.23 },
  ],
};

// The exercise sections and their text live in assets/exercise-types.json.
// LAYOUTS above defines a row shape for each supported type; the book's
// `order` array selects which of them this project actually prints, so a
// project with order: ["split"] builds a split-only book even though the
// other two shapes are still defined here:
//   split     : الصورة + الاسم + مربعات فارغة        (LAYOUTS.split)
//   contextual: الصورة + الحروف بأشكالها السياقية + خط  (LAYOUTS.contextual)
//   compose   : الصورة + الحروف مرتبة + خط للكتابة عليه (LAYOUTS.compose)
// ملاحظة: لا يُطبع اسم الكلمة في قسمي contextual وcompose — المطلوب أن يقرأه
// الطفل من الأشكال أو أن يكوّنه من الحروف.

// ---------------------------------------------------------------------------
// Content
//
// The book text lives in assets/*.json and is checked here before anything is
// laid out, so a typo in the data fails the build instead of producing a page
// with a missing word, a missing image or an empty section.
// ---------------------------------------------------------------------------

function readJson(file) {
  if (!fs.existsSync(file)) {
    throw new Error(`Missing data file: ${path.relative(ROOT, file)}`);
  }
  const raw = fs.readFileSync(file, "utf-8").replace(/^\uFEFF/, "");
  let parsed;
  try {
    parsed = JSON.parse(raw);
  } catch (err) {
    throw new Error(`${path.relative(ROOT, file)} is not valid JSON: ${err.message}`);
  }
  if (parsed === null || typeof parsed !== "object" || Array.isArray(parsed)) {
    throw new Error(`${path.relative(ROOT, file)} must contain a JSON object.`);
  }
  return parsed;
}

function requireText(value, where) {
  if (typeof value !== "string" || value.trim() === "") {
    throw new Error(`${where} must be a non-empty string.`);
  }
  return value;
}

function requireBoolean(value, where) {
  if (typeof value !== "boolean") {
    throw new Error(`${where} must be true or false.`);
  }
  return value;
}

function requireDirList(value, where) {
  if (!Array.isArray(value) || value.length === 0) {
    throw new Error(`${where} must be a non-empty array of directories.`);
  }
  for (const [index, dir] of value.entries()) {
    requireText(dir, `${where}[${index}]`);
  }
  return value;
}

/** First existing `dir`/file from a list of directories, or null. */
function resolveInDirs(dirs, file) {
  for (const dir of dirs) {
    const full = path.resolve(ROOT, dir, file);
    if (fs.existsSync(full)) return full;
  }
  return null;
}

/** First existing path from a list of paths relative to the project, or null. */
function resolveRelativePath(paths) {
  for (const relative of paths) {
    const full = path.resolve(ROOT, relative);
    if (fs.existsSync(full)) return full;
  }
  return null;
}

function loadContent() {
  const wordsFile = readJson(WORDS_FILE);
  const exercisesFile = readJson(EXERCISES_FILE);

  // -- words.json ----------------------------------------------------------
  for (const key of ["creator", "title", "description"]) {
    requireText(wordsFile.book && wordsFile.book[key], `book.${key}`);
  }
  if (!Array.isArray(wordsFile.words) || wordsFile.words.length === 0) {
    throw new Error('words.json must contain a non-empty "words" array.');
  }
  requireDirList(wordsFile.imageDirs, "imageDirs");

  const seen = new Set();
  for (const [index, entry] of wordsFile.words.entries()) {
    const at = `words[${index}]`;
    requireText(entry.word, `${at}.word`);
    requireText(entry.image, `${at}.image`);
    if (seen.has(entry.word)) {
      throw new Error(`${at}: word "${entry.word}" is listed more than once.`);
    }
    seen.add(entry.word);
    // Contextual shapes come from FORMS, which is keyed by code point, so every
    // letter of every word has to be one the shaping tables know about.
    for (const ch of entry.word) {
      if (ch === SHADDA) continue;
      if (!FORMS[ch.codePointAt(0)]) {
        throw new Error(`${at}: "${ch}" (U+${ch.codePointAt(0).toString(16).toUpperCase()}) is not in FORMS.`);
      }
    }
    if (!resolveInDirs(wordsFile.imageDirs, entry.image)) {
      throw new Error(`${at}: image "${entry.image}" is not in ${wordsFile.imageDirs.join(", ")}.`);
    }
  }

  requireText(wordsFile.cover && wordsFile.cover.file, "cover.file");
  requireDirList(wordsFile.cover && wordsFile.cover.candidates, "cover.candidates");

  // -- exercise-types.json -------------------------------------------------
  const order = exercisesFile.order;
  if (!Array.isArray(order) || order.length === 0) {
    throw new Error('exercise-types.json must contain a non-empty "order" array.');
  }
  const types = exercisesFile.types;
  if (!types || typeof types !== "object" || Array.isArray(types)) {
    throw new Error('exercise-types.json is missing the "types" section.');
  }
  if (order.length !== Object.keys(types).length) {
    throw new Error("order and types must list the same number of exercise types.");
  }

  const exerciseTypes = [];
  for (const [index, id] of order.entries()) {
    const at = `order[${index}]`;
    requireText(id, at);
    if (order.indexOf(id) !== index) {
      throw new Error(`${at}: "${id}" is listed more than once.`);
    }
    const type = types[id];
    if (!type) {
      throw new Error(`${at}: "${id}" has no entry in types.`);
    }
    // The layout is code, not data, so both sides have to agree on the set.
    if (!LAYOUTS[id]) {
      throw new Error(`${at}: "${id}" has no layout in LAYOUTS.`);
    }
    for (const key of ["title", "pageHeader", "subtitle"]) {
      requireText(type[key], `types.${id}.${key}`);
    }
    requireBoolean(type.lettersInBoxes, `types.${id}.lettersInBoxes`);
    // contextualForms only has an effect when the boxes hold letters.
    if (type.contextualForms !== undefined) {
      requireBoolean(type.contextualForms, `types.${id}.contextualForms`);
      if (type.contextualForms && !type.lettersInBoxes) {
        throw new Error(`types.${id}.contextualForms needs lettersInBoxes to be true.`);
      }
    }
    if (!LAYOUTS[id].some((slot) => slot.kind === "boxes")) {
      throw new Error(`LAYOUTS.${id} must have a boxes column.`);
    }
    exerciseTypes.push({ id, ...type });
  }
  for (const id of Object.keys(types)) {
    if (!order.includes(id)) {
      throw new Error(`types.${id} is not listed in order.`);
    }
    if (!LAYOUTS[id]) {
      throw new Error(`LAYOUTS.${id} is missing for the type in types.`);
    }
  }

  // The page order is the array order, so --order only has to re-sort this copy.
  const words = sortByLetter(wordsFile.words, CLI.order, {
    key: (entry) => entry.word,
    label: "words.json words",
  });

  return {
    book: wordsFile.book,
    imageDirs: wordsFile.imageDirs,
    cover: wordsFile.cover,
    words,
    exerciseTypes,
  };
}

// ---------------------------------------------------------------------------
// Image helpers
// ---------------------------------------------------------------------------
const imageCache = new Map();

// Read width/height from a PNG buffer (IHDR at bytes 16-23)
function getPngDimensions(buffer) {
  return {
    width: buffer.readUInt32BE(16),
    height: buffer.readUInt32BE(20),
  };
}

function svgToBuffer(svgContent, width) {
  return new Promise((resolve, reject) => {
    svg2img(
      svgContent,
      { resvg: { fitTo: { mode: "width", value: width } } },
      (error, buffer) => {
        if (error) reject(error);
        else resolve(buffer);
      }
    );
  });
}

// Load an image as a PNG buffer. SVG is rasterised through resvg.
// Results are cached so each image is converted only once per build.
// loadContent() has already checked that every image exists in imageDirs.
async function loadImage(imageName) {
  if (imageCache.has(imageName)) return imageCache.get(imageName);

  const imgPath = resolveInDirs(content.imageDirs, imageName);
  if (!imgPath) {
    throw new Error(`Image not found: ${imageName} (looked in ${content.imageDirs.join(", ")})`);
  }

  let buffer = null;
  try {
    if (path.extname(imageName).toLowerCase() === ".png") {
      buffer = fs.readFileSync(imgPath);
    } else {
      const svgContent = fs.readFileSync(imgPath, "utf-8");
      buffer = await svgToBuffer(svgContent, SVG_RENDER_WIDTH);
    }
  } catch (error) {
    console.log(`  Warning: could not load image ${imageName}: ${error.message}`);
  }

  imageCache.set(imageName, buffer);
  return buffer;
}

// Contain-fit the image inside a square box, preserving aspect ratio
function fitImage(pngBuffer, maxSize) {
  const dims = getPngDimensions(pngBuffer);
  const scale = Math.min(maxSize / dims.width, maxSize / dims.height);
  return {
    width: Math.max(1, Math.round(dims.width * scale)),
    height: Math.max(1, Math.round(dims.height * scale)),
  };
}

// ---------------------------------------------------------------------------
// Letter helpers
// ---------------------------------------------------------------------------
const SHADDA = "\u0651";

/**
 * Split a word into letter units, in reading order.
 * A shadda is attached to the letter it sits on rather than becoming
 * its own unit, so that a shadda can be counted as an extra box.
 */
function getLetterUnits(word) {
  const units = [];
  for (const ch of word) {
    if (ch === SHADDA) {
      if (units.length > 0) units[units.length - 1].shadda = true;
      continue;
    }
    units.push({ letter: ch, shadda: false });
  }
  return units;
}

/**
 * One box per letter, plus one extra box per shadda.
 * The 28 current words contain no shadda, so this returns the letter
 * count today; the rule is in place for words added later.
 */
function getBoxCount(units) {
  return units.reduce((total, unit) => total + (unit.shadda ? 2 : 1), 0);
}

// ---------------------------------------------------------------------------
// Contextual letter forms
//
// Each letter is printed on its own, so nothing joins it to a neighbour and
// the font always draws the isolated shape. To show the shape the letter
// really has inside its word we have to print a pre-shaped glyph. Two
// mechanisms were measured against the font (see README "Contextual forms"):
//
//   * Presentation Forms (U+FB50..U+FEFF) are standalone pre-shaped glyphs
//     that take no part in contextual shaping, so emitting the right codepoint
//     yields exactly the wanted form. Traditional Arabic ships every form
//     except all four teh marbuta ones.
//   * TATWEEL + letter forces the final form in the font's own glyph, so it
//     covers teh marbuta. Validated against heh, which does have a real final
//     form: tatweel+heh draws the identical glyph and identical height, with
//     only a short connecting stub added.
//
// ZWJ is NOT usable for medial/final (ZWJ+letter+ZWJ merely repeats the
// initial form) and ZWNJ merely blocks joining, i.e. it yields isolated.
// Emitting FB9F must be avoided: the font lacks it, so Word substitutes a
// different font and draws a glyph about three times too large.
// ---------------------------------------------------------------------------
const TATWEEL = "\u0640";
const ZWJ = "\u200D";

// Letters that join on both sides, so they can be initial, medial or final.
const DUAL = new Set(Array.from(
  "\u0628\u062A\u062B\u062C\u062D\u062E\u0633\u0634\u0635\u0636\u0637\u0638" +
  "\u0639\u063A\u0641\u0642\u0643\u0644\u0645\u0646\u0647\u064A\u0626"
));
// Letters that only join to the letter on their right: they accept a
// connection from the previous letter but never start one towards the next.
const RIGHT_JOINING = new Set(Array.from(
  "\u0622\u0623\u0624\u0625\u0627\u062F\u0630\u0631\u0632\u0648\u0649\u0629"
));

// base codepoint -> { i: isolated, f: final, n: initial, m: medial }.
// The Presentation Forms block is ordered ISOLATED, FINAL, INITIAL, MEDIAL.
// 0 means the letter has no such form.

// Presentation forms that Traditional Arabic does NOT ship. Verified by
// reading the font cmap: every other form in the table below is present.
const MISSING_IN_FONT = new Set([0xFB9E, 0xFB9F, 0xFBA0, 0xFBA1]);

const FORMS = {
  0x0621: { i: 0xFE80, f: 0, n: 0, m: 0 }, // ء hamza
  0x0623: { i: 0xFE83, f: 0xFE84, n: 0, m: 0 }, // أ
  0x0626: { i: 0xFE89, f: 0xFE8A, n: 0xFE8B, m: 0xFE8C }, // ئ
  0x0627: { i: 0xFE8D, f: 0xFE8E, n: 0, m: 0 }, // ا
  0x0628: { i: 0xFE8F, f: 0xFE90, n: 0xFE91, m: 0xFE92 }, // ب
  0x062A: { i: 0xFE95, f: 0xFE96, n: 0xFE97, m: 0xFE98 }, // ت
  0x062B: { i: 0xFE99, f: 0xFE9A, n: 0xFE9B, m: 0xFE9C }, // ث
  0x062C: { i: 0xFE9D, f: 0xFE9E, n: 0xFE9F, m: 0xFEA0 }, // ج
  0x062D: { i: 0xFEA1, f: 0xFEA2, n: 0xFEA3, m: 0xFEA4 }, // ح
  0x062E: { i: 0xFEA5, f: 0xFEA6, n: 0xFEA7, m: 0xFEA8 }, // خ
  0x062F: { i: 0xFEA9, f: 0xFEAA, n: 0, m: 0 }, // د
  0x0630: { i: 0xFEAB, f: 0xFEAC, n: 0, m: 0 }, // ذ
  0x0631: { i: 0xFEAD, f: 0xFEAE, n: 0, m: 0 }, // ر
  0x0632: { i: 0xFEAF, f: 0xFEB0, n: 0, m: 0 }, // ز
  0x0633: { i: 0xFEB1, f: 0xFEB2, n: 0xFEB3, m: 0xFEB4 }, // س
  0x0634: { i: 0xFEB5, f: 0xFEB6, n: 0xFEB7, m: 0xFEB8 }, // ش
  0x0635: { i: 0xFEB9, f: 0xFEBA, n: 0xFEBB, m: 0xFEBC }, // ص
  0x0636: { i: 0xFEBD, f: 0xFEBE, n: 0xFEBF, m: 0xFEC0 }, // ض
  0x0637: { i: 0xFEC1, f: 0xFEC2, n: 0xFEC3, m: 0xFEC4 }, // ط
  0x0638: { i: 0xFEC5, f: 0xFEC6, n: 0xFEC7, m: 0xFEC8 }, // ظ
  0x0639: { i: 0xFEC9, f: 0xFECA, n: 0xFECB, m: 0xFECC }, // ع
  0x063A: { i: 0xFECD, f: 0xFECE, n: 0xFECF, m: 0xFED0 }, // غ
  0x0641: { i: 0xFED1, f: 0xFED2, n: 0xFED3, m: 0xFED4 }, // ف
  0x0642: { i: 0xFED5, f: 0xFED6, n: 0xFED7, m: 0xFED8 }, // ق
  0x0643: { i: 0xFED9, f: 0xFEDA, n: 0xFEDB, m: 0xFEDC }, // ك
  0x0644: { i: 0xFEDD, f: 0xFEDE, n: 0xFEDF, m: 0xFEE0 }, // ل
  0x0645: { i: 0xFEE1, f: 0xFEE2, n: 0xFEE3, m: 0xFEE4 }, // م
  0x0646: { i: 0xFEE5, f: 0xFEE6, n: 0xFEE7, m: 0xFEE8 }, // ن
  0x0647: { i: 0xFEE9, f: 0xFEEA, n: 0xFEEB, m: 0xFEEC }, // ه
  0x0648: { i: 0xFEED, f: 0xFEEE, n: 0, m: 0 }, // و
  0x064A: { i: 0xFEF1, f: 0xFEF2, n: 0xFEF3, m: 0xFEF4 }, // ي
  0x0629: { i: 0xFB9E, f: 0xFB9F, n: 0xFBA0, m: 0xFBA1 }, // ة
};

const PF_AVAILABLE = new Set(
  Object.keys(FORMS)
    .flatMap((k) => {
      const f = FORMS[k];
      return [f.i, f.f, f.n, f.m];
    })
    .filter((cp) => cp !== 0 && !MISSING_IN_FONT.has(cp))
);

/**
 * Which contextual form does `letter` take at `index` of `letters`?
 * A connection across a boundary needs both sides to agree: the right-hand
 * letter must join forward and the left-hand letter must join backward, so
 * this is not simply a property of the neighbour.
 */
function contextualForm(letter, index, letters) {
  const prev = index > 0 ? letters[index - 1] : null;
  const next = index < letters.length - 1 ? letters[index + 1] : null;
  const dual = (ch) => ch !== null && DUAL.has(ch);
  const right = (ch) => ch !== null && RIGHT_JOINING.has(ch);

  const joinsPrev = prev !== null && dual(prev) && (dual(letter) || right(letter));
  const joinsNext = next !== null && dual(letter) && (dual(next) || right(next));

  if (joinsPrev && joinsNext) return "m";
  if (!joinsPrev && joinsNext) return "n";
  if (joinsPrev && !joinsNext) return "f";
  return "i";
}

/** The exact string to print in a box for this letter in this form. */
function contextualGlyph(letter, form) {
  const forms = FORMS[letter.codePointAt(0)];
  if (forms) {
    const cp = forms[form];
    if (cp && PF_AVAILABLE.has(cp)) return String.fromCodePoint(cp);
  }
  // No presentation form in this font: force the shape in the font itself.
  if (form === "f") return TATWEEL + letter;
  if (form === "n") return letter + ZWJ;
  if (form === "m") return TATWEEL + letter + TATWEEL;
  return letter;
}

/** Contextual shape of every letter of a word, in reading order. */
function contextualGlyphs(word) {
  const letters = Array.from(word);
  return letters.map((ch, i) => contextualGlyph(ch, contextualForm(ch, i, letters)));
}

// ---------------------------------------------------------------------------
// Building blocks
// ---------------------------------------------------------------------------
// Loaded once, after the shaping tables above exist because the validation
// checks every word against FORMS. module.exports exposes this same data.
const content = loadContent();
const words = content.words;

const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const NO_BORDERS = {
  top: noBorder,
  bottom: noBorder,
  left: noBorder,
  right: noBorder,
  insideHorizontal: noBorder,
  insideVertical: noBorder,
};

const BOX_BORDERS = {
  top: { style: BorderStyle.SINGLE, size: BOX_BORDER_SIZE, color: COLOR_BORDER },
  bottom: { style: BorderStyle.SINGLE, size: BOX_BORDER_SIZE, color: COLOR_BORDER },
  left: { style: BorderStyle.SINGLE, size: BOX_BORDER_SIZE, color: COLOR_BORDER },
  right: { style: BorderStyle.SINGLE, size: BOX_BORDER_SIZE, color: COLOR_BORDER },
  insideHorizontal: { style: BorderStyle.SINGLE, size: BOX_BORDER_SIZE, color: COLOR_BORDER },
  insideVertical: { style: BorderStyle.SINGLE, size: BOX_BORDER_SIZE, color: COLOR_BORDER },
};

function emptyParagraph() {
  return new Paragraph({
    spacing: { before: 0, after: 0 },
    direction: "RTL",
    children: [],
  });
}

/**
 * A strip of adjacent equal-sized boxes.
 * `letterTexts` fills the boxes, one string per unit, in order. Pass null for
 * empty boxes. A string may be a whole contextual shape (a presentation form
 * or TATWEEL+letter), not just a single character.
 */
function createBoxesRow(units, letterTexts) {
  const cells = [];
  for (let i = 0; i < units.length; i++) {
    const unit = units[i];
    const runs = [];
    if (letterTexts) {
      runs.push(
        new TextRun({
          text: letterTexts[i],
          font: FONT,
          size: LETTER_SIZE,
          bold: true,
          color: COLOR_LETTER,
        })
      );
    }
    cells.push(
      new TableCell({
        width: { size: BOX_W, type: WidthType.DXA },
        verticalAlign: VerticalAlign.CENTER,
        borders: BOX_BORDERS,
        margins: { top: 0, bottom: 0, left: 20, right: 20 },
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: {
              before: 0,
              after: 0,
              line: LETTER_LINE,
              lineRule: docx.LineRuleType.AT_LEAST,
            },
            direction: "RTL",
            children: runs,
          }),
        ],
      })
    );

    // A shadda gets its own extra box, so the child can place it too.
    if (unit.shadda) {
      cells.push(
        new TableCell({
          width: { size: BOX_W, type: WidthType.DXA },
          verticalAlign: VerticalAlign.CENTER,
          borders: BOX_BORDERS,
          margins: { top: 0, bottom: 0, left: 20, right: 20 },
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { before: 0, after: 0 },
              direction: "RTL",
            children: letterTexts
              ? [
                    new TextRun({
                      text: SHADDA,
                      font: FONT,
                      size: LETTER_SIZE,
                      bold: true,
                      color: COLOR_LETTER,
                    }),
                  ]
                : [],
            }),
          ],
        })
      );
    }
  }

  return new Table({
    width: { size: cells.length * BOX_W, type: WidthType.DXA },
    alignment: AlignmentType.CENTER,
    layout: docx.TableLayoutType.FIXED,
    borders: BOX_BORDERS,
    rows: [
      new TableRow({
        height: { value: BOX_H, rule: HeightRule.ATLEAST },
        // Reverse so the first letter of the word sits in the rightmost box
        // and the word reads right-to-left. Done by reversing the cells
        // rather than with visuallyRightToLeft, because bidiVisual makes
        // Word mirror the grid and push the strip outside the page margins.
        children: cells.slice().reverse(),
      }),
    ],
  });
}

/** A single continuous horizontal line for writing the whole word. */
function createWriteLine() {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: docx.TableLayoutType.FIXED,
    borders: {
      top: noBorder,
      left: noBorder,
      right: noBorder,
      insideHorizontal: noBorder,
      insideVertical: noBorder,
      bottom: { style: BorderStyle.SINGLE, size: 12, color: COLOR_LINE },
    },
    rows: [
      new TableRow({
        height: { value: WRITE_LINE_HEIGHT, rule: HeightRule.ATLEAST },
        children: [
          new TableCell({
            width: { size: 100, type: WidthType.PERCENTAGE },
            verticalAlign: VerticalAlign.BOTTOM,
            borders: {
              top: noBorder,
              left: noBorder,
              right: noBorder,
              bottom: { style: BorderStyle.SINGLE, size: 12, color: COLOR_LINE },
            },
            margins: { top: 0, bottom: 0, left: 0, right: 0 },
            children: [
              new Paragraph({
                spacing: {
                  before: 0,
                  after: 0,
                  line: WRITE_LINE_INNER,
                  lineRule: docx.LineRuleType.EXACT,
                },
                direction: "RTL",
                children: [],
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

// ---------------------------------------------------------------------------
// Exercise
// ---------------------------------------------------------------------------
function createImageCell(widthPercent, pngBuffer) {
  const children = [];

  if (pngBuffer) {
    const size = fitImage(pngBuffer, IMAGE_BOX_PX);
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 0 },
        direction: "RTL",
        children: [
          new ImageRun({
            data: pngBuffer,
            transformation: { width: size.width, height: size.height },
            type: "png",
          }),
        ],
      })
    );
  } else {
    children.push(emptyParagraph());
  }

  return new TableCell({
    width: { size: Math.round(CONTENT_WIDTH * widthPercent), type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER,
    borders: NO_BORDERS,
    margins: { top: 0, bottom: 0, left: 0, right: 0 },
    children,
  });
}

/** Shared shell for every exercise cell: fixed DXA width, no borders. */
function createCell(widthPercent, children) {
  return new TableCell({
    width: { size: Math.round(CONTENT_WIDTH * widthPercent), type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER,
    borders: NO_BORDERS,
    margins: { top: 0, bottom: 0, left: 0, right: 0 },
    children,
  });
}

/** خلية اسم الكلمة (القسم الأول فقط) */
function createWordCell(widthPercent, units) {
  return createCell(widthPercent, [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: {
        before: 0,
        after: 0,
        line: WORD_LINE,
        lineRule: docx.LineRuleType.AT_LEAST,
      },
      direction: "RTL",
      children: [
        new TextRun({
          text: units.map((u) => u.letter).join(""),
          font: FONT,
          size: WORD_SIZE,
          bold: true,
          color: COLOR_MAIN,
        }),
      ],
    }),
  ]);
}

/**
 * خلية المربعات — فارغة في التحليل، ومطبوعة الحروف في التجميع،
 * وفي القسم الجديد تُطبع الحروف بشكلها السياقي داخل الكلمة.
 */
function createBoxesCell(widthPercent, units, type) {
  let letterTexts = null;
  if (type.lettersInBoxes) {
    letterTexts = type.contextualForms
      ? contextualGlyphs(Array.from(units.map((u) => u.letter).join("")))
      : units.map((u) => u.letter);
  }
  return createCell(widthPercent, [createBoxesRow(units, letterTexts)]);
}

/** خلية خط الكتابة */
function createLineCell(widthPercent) {
  return createCell(widthPercent, [createWriteLine()]);
}

/** One exercise as a single full-width row in the vertical stack. */
function createExerciseRow(entry, type, pngBuffer) {
  const units = getLetterUnits(entry.word);
  const layout = LAYOUTS[type.id];
  const cells = [];

  for (const slot of layout) {
    if (slot.kind === "image") {
      cells.push(createImageCell(slot.width, pngBuffer));
    } else if (slot.kind === "word") {
      cells.push(createWordCell(slot.width, units));
    } else if (slot.kind === "boxes") {
      cells.push(createBoxesCell(slot.width, units, type));
    } else if (slot.kind === "line") {
      cells.push(createLineCell(slot.width));
    }
  }

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    alignment: AlignmentType.CENTER,
    columnWidths: layout.map((s) => Math.round(CONTENT_WIDTH * s.width)),
    visuallyRightToLeft: true,
    layout: docx.TableLayoutType.FIXED,
    borders: NO_BORDERS,
    rows: [
      new TableRow({
        height: { value: ROW_HEIGHT, rule: HeightRule.ATLEAST },
        cantSplit: true,
        children: cells,
      }),
    ],
  });
}

// ---------------------------------------------------------------------------
// Pages
// ---------------------------------------------------------------------------
function createPageHeader(type) {
  return new Paragraph({
    // Right-aligned, not centred: the sentence anchors to the same right edge
    // as the image column of every exercise below it.
    alignment: AlignmentType.RIGHT,
    spacing: { before: 0, after: HEADER_AFTER },
    direction: "RTL",
        children: [
          new TextRun({
            text: type.pageHeader,
            font: FONT,
            size: HEADER_SIZE,
            bold: true,
            color: COLOR_TITLE,
          }),
        ],
  });
}

/** A page of 5 exercises: 2 top, 1 middle, 2 bottom. */
function createExercisePage(entries, type) {
  const children = [createPageHeader(type)];

  entries.forEach((entry, index) => {
    children.push(createExerciseRow(entry, type, imageCache.get(entry.image)));

    if (index < entries.length - 1) {
      children.push(
        new Paragraph({
          spacing: {
            before: 0,
            after: GAP_AFTER,
            line: GAP_LINE,
            lineRule: docx.LineRuleType.EXACT,
          },
          direction: "RTL",
          children: [],
        })
      );
    }
  });

  return children;
}

function createSectionTitlePage(type) {
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 1600, after: 400 },
      direction: "RTL",
      children: [
        new TextRun({
          text: type.title,
          font: FONT,
          size: 72, // 36pt
          bold: true,
          color: COLOR_TITLE,
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 400 },
      direction: "RTL",
      children: [
        new TextRun({
          text: type.subtitle,
          font: FONT,
          size: 48, // 24pt
          bold: true,
          color: COLOR_MAIN,
        }),
      ],
    }),
  ];
}

/** Full-bleed cover, looked up in the order given by words.json. */
function createCoverPage() {
  const coverPath = resolveRelativePath(content.cover.candidates);
  if (!coverPath) {
    console.log("Warning: cover-arabic.png not found, skipping cover page");
    return null;
  }

  const coverBuffer = fs.readFileSync(coverPath);
  const dims = getPngDimensions(coverBuffer);

  // Full A5 page in pixels at 96 dpi
  const pageWidthPx = A5_WIDTH / 15;
  const pageHeightPx = A5_HEIGHT / 15;

  const scale = Math.max(pageWidthPx / dims.width, pageHeightPx / dims.height);

  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 0 },
      children: [
        new ImageRun({
          data: coverBuffer,
          transformation: {
            width: Math.round(dims.width * scale),
            height: Math.round(dims.height * scale),
          },
          type: "png",
        }),
      ],
    }),
  ];
}

// Create footer with page number
function createPageFooter() {
  if (!SHOW_PAGE_NUMBERS) {
    return null;
  }
  return new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        direction: "RTL",
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

function pageProperties(marginTop, marginBottom, marginSide = MARGIN_SIDE, framed = true) {
  return {
    page: {
      size: {
        width: A5_WIDTH,
        height: A5_HEIGHT,
        orientation: PageOrientation.PORTRAIT,
      },
      margin: {
        top: marginTop,
        right: marginSide,
        bottom: marginBottom,
        left: marginSide,
      },
      ...(framed ? { borders: PAGE_FRAME } : {}),
    },
    direction: "RTL",
  };
}

// ---------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------
async function createWorkbook() {
  console.log(`جاري إنشاء ${content.book.title}...`);

  // Rasterise every image once, up front
  for (const entry of words) {
    const buffer = await loadImage(entry.image);
    if (!buffer) console.log(`  Warning: missing image for ${entry.word}`);
  }

  const doc = new Document({
    creator: content.book.creator,
    title: content.book.title,
    description: content.book.description,
    styles: {
      default: {
        document: {
          run: {
            font: FONT,
            size: 28, // 14pt
          },
        },
      },
    },
    sections: [],
  });

  // Cover
  const coverContent = createCoverPage();
  if (coverContent) {
    doc.addSection({
      // All four margins must be zero: the cover image is deliberately
      // oversized and centred so it bleeds off every edge. Leaving the 0.5in
      // side margin here shifts the image right and leaves a white band.
      // The page frame is skipped too, so the cover stays a full-bleed image.
      properties: pageProperties(0, 0, 0, false),
      children: coverContent,
    });
  }

  const typeList = content.exerciseTypes;
  const pageCountByType = {};
  const footers = SHOW_PAGE_NUMBERS ? { footers: { default: createPageFooter() } } : {};

  for (const type of typeList) {
    console.log(`جاري إنشاء قسم: ${type.title}`);

    doc.addSection({
      properties: pageProperties(MARGIN_TOP, MARGIN_BOTTOM),
      ...footers,
      children: createSectionTitlePage(type),
    });

    let pages = 0;
    for (let i = 0; i < words.length; i += EXERCISES_PER_PAGE) {
      const entries = words.slice(i, i + EXERCISES_PER_PAGE);
      pages += 1;
      console.log(
        `  صفحة ${pages}: ${entries.map((e) => e.word).join("، ")}`
      );

      doc.addSection({
        properties: pageProperties(MARGIN_TOP, MARGIN_BOTTOM),
        ...footers,
        children: createExercisePage(entries, type),
      });
    }

    pageCountByType[type.id] = pages;
  }

  console.log("جاري حفظ الملف...");

  const buffer = await Packer.toBuffer(doc);
  fs.mkdirSync(path.dirname(OUTPUT), { recursive: true });
  fs.writeFileSync(OUTPUT, buffer);

  console.log("تم إنشاء الملف بنجاح!");
  console.log(`الملف: ${OUTPUT}`);
  console.log(`الترتيب: ${CLI.order}`);
  for (const type of typeList) {
    console.log(
      `عدد صفحات «${type.title}»: ${pageCountByType[type.id]} صفحة تمارين + 1 عنوان`
    );
  }
  console.log(
    `الإجمالي: ${(coverContent ? 1 : 0) + typeList.length +
      typeList.reduce((n, t) => n + pageCountByType[t.id], 0)} صفحة`
  );
  console.log("حجم الصفحة: A5");
}

if (require.main === module) {
  createWorkbook().catch((err) => {
    console.error(`[FAIL] ${err.message}`);
    process.exitCode = 1;
  });
}

// Exported so the shaping logic can be checked against the real implementation
// instead of a copy that could drift.
module.exports = {
  FORMS,
  MISSING_IN_FONT,
  PF_AVAILABLE,
  DUAL,
  RIGHT_JOINING,
  contextualForm,
  contextualGlyph,
  contextualGlyphs,
  words,
  LAYOUTS,
  EXERCISE_TYPES: content.exerciseTypes,
};
