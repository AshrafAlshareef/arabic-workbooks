#!/usr/bin/env node
/**
 * create-workbook.js
 *
 * Builds `Arabic_Alphabet_Workbook.docx` from the assets in `assets/` only --
 * no letter, word, shape or caption is hard-coded here.
 *
 *   assets/letters.json        book metadata, page titles, table config and the
 *                              ordered letter list (letter, name, word, image)
 *   assets/letter-shapes.json  the four positional shapes of every letter
 *   assets/images              the picture for each letter
 *
 * Usage: node create-workbook.js
 */

const docx = require("docx");
const fs = require("fs");
const path = require("path");
const svg2img = require("svg2img");

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
  ShadingType,
  ImageRun,
  PageNumber,
  Footer,
  PageBorderDisplay,
  PageBorderOffsetFrom,
  PageBorderZOrder,
} = docx;

const ROOT = __dirname;
const ASSETS = path.join(ROOT, "assets");
const LETTERS_FILE = path.join(ASSETS, "letters.json");
const SHAPES_FILE = path.join(ASSETS, "letter-shapes.json");
const OUTPUT = path.join(ROOT, "Arabic_Alphabet_Workbook.docx");

// Set to false to hide the page number footer
const SHOW_PAGE_NUMBERS = false;

const FONT = "Traditional Arabic";
const FONT_DOTTED = "KFGQPC Alphabet Dotted";

// Width/height every letter picture is rasterised to before embedding.
const IMAGE_RENDER_WIDTH = 400;
const IMAGE_TRANSFORM = { width: 250, height: 250 };

// Every positional shape a table column can show.
const SHAPE_KEYS = ["isolated", "beginning", "middle", "end"];

// Simple single-line frame around every content page. Offset is measured from
// the paper edge, so the frame sits outside the text area on every page.
const PAGE_FRAME = {
  pageBorders: {
    display: PageBorderDisplay.ALL_PAGES,
    offsetFrom: PageBorderOffsetFrom.PAGE,
    zOrder: PageBorderZOrder.FRONT,
  },
  pageBorderTop: { style: BorderStyle.SINGLE, size: 6, color: "34495E", space: 18 },
  pageBorderRight: { style: BorderStyle.SINGLE, size: 6, color: "34495E", space: 18 },
  pageBorderBottom: { style: BorderStyle.SINGLE, size: 6, color: "34495E", space: 18 },
  pageBorderLeft: { style: BorderStyle.SINGLE, size: 6, color: "34495E", space: 18 },
};

// A5 dimensions in twips (14.8cm x 21cm)
const A5_WIDTH = convertInchesToTwip(5.83);  // 14.8cm
const A5_HEIGHT = convertInchesToTwip(8.27);  // 21cm

const CONTENT_MARGIN = convertInchesToTwip(0.5);

// Read width/height from a PNG buffer (IHDR at bytes 16-23)
function getPngDimensions(buffer) {
  return {
    width: buffer.readUInt32BE(16),
    height: buffer.readUInt32BE(20),
  };
}

// ---------------------------------------------------------------------------
// Content
// ---------------------------------------------------------------------------

function readJson(file) {
  if (!fs.existsSync(file)) {
    throw new Error(`Missing data file: ${path.relative(ROOT, file)}`);
  }
  // The BOM strip keeps the file readable after a Windows editor adds one.
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

function requireTitle(value, where) {
  requireText(value, where);
  if (!value.includes("{name}")) {
    throw new Error(`${where} must contain the {name} placeholder.`);
  }
  return value;
}

function requireCount(value, where) {
  if (!Number.isInteger(value) || value < 0) {
    throw new Error(`${where} must be a non-negative whole number.`);
  }
  return value;
}

/** First existing file for `file` inside `dirs`, or null. Dirs are relative to ROOT. */
function resolveAsset(dirs, file) {
  for (const dir of dirs) {
    const candidate = path.join(ROOT, dir, file);
    if (fs.existsSync(candidate)) return candidate;
  }
  return null;
}

/**
 * Read both data files and check everything the build relies on, so a typo in
 * the JSON fails here instead of silently producing a wrong book.
 * Returns the letters plus the resolved picture path of each letter.
 */
function loadContent() {
  const manifest = readJson(LETTERS_FILE);
  const allShapes = readJson(SHAPES_FILE);

  for (const key of ["book", "pages", "table"]) {
    if (!manifest[key] || typeof manifest[key] !== "object") {
      throw new Error(`letters.json is missing the "${key}" section.`);
    }
  }
  for (const key of ["creator", "title", "description"]) {
    requireText(manifest.book[key], `book.${key}`);
  }
  for (const key of ["letterTitle", "shapesTitle"]) {
    requireTitle(manifest.pages[key], `pages.${key}`);
  }

  if (!Array.isArray(manifest.imageDirs) || manifest.imageDirs.length === 0) {
    throw new Error("imageDirs must be a non-empty array of directories.");
  }
  if (!Array.isArray(manifest.letters) || manifest.letters.length === 0) {
    throw new Error("letters.json must contain a non-empty \"letters\" array.");
  }

  // -- table columns -------------------------------------------------------
  const columns = manifest.table.columns;
  if (!Array.isArray(columns) || columns.length === 0) {
    throw new Error("table.columns must be a non-empty array.");
  }
  for (const column of columns) {
    if (!SHAPE_KEYS.includes(column.shape)) {
      throw new Error(
        `table.columns: "${column.shape}" is not one of ${SHAPE_KEYS.join(", ")}.`
      );
    }
    requireText(column.header, `table.columns header for "${column.shape}"`);
  }
  requireCount(manifest.table.dottedRows, "table.dottedRows");
  requireCount(manifest.table.emptyRows, "table.emptyRows");
  requireCount(manifest.table.dottedRowHeight, "table.dottedRowHeight");
  requireCount(manifest.table.emptyRowHeight, "table.emptyRowHeight");

  // -- letters -------------------------------------------------------------
  const seen = new Set();
  const images = new Map();
  const shapesByLetter = new Map();

  for (const [index, entry] of manifest.letters.entries()) {
    const at = `letters[${index}]`;
    requireText(entry.letter, `${at}.letter`);
    requireText(entry.name, `${at}.name`);
    requireText(entry.word, `${at}.word`);
    requireText(entry.image, `${at}.image`);
    if (entry.mainForm !== undefined) requireText(entry.mainForm, `${at}.mainForm`);

    if (seen.has(entry.letter)) {
      throw new Error(`${at}: letter "${entry.letter}" is listed more than once.`);
    }
    seen.add(entry.letter);

    if (!entry.word.startsWith(entry.letter)) {
      throw new Error(
        `${at}: word "${entry.word}" does not start with its letter "${entry.letter}".`
      );
    }

    const shapes = allShapes[entry.letter];
    if (!shapes) {
      throw new Error(`${at}: letter-shapes.json has no entry for "${entry.letter}".`);
    }
    for (const key of SHAPE_KEYS) {
      requireText(shapes[key], `letter-shapes.json["${entry.letter}"].${key}`);
    }
    shapesByLetter.set(entry.letter, shapes);

    const imagePath = resolveAsset(manifest.imageDirs, entry.image);
    if (!imagePath) {
      throw new Error(
        `${at}: image "${entry.image}" not found. Searched: ${manifest.imageDirs.join(", ")}`
      );
    }
    images.set(entry.letter, imagePath);
  }

  // A shape table nobody uses is a trap for the next person who edits it.
  for (const letter of Object.keys(allShapes)) {
    if (!seen.has(letter)) {
      throw new Error(
        `letter-shapes.json has an entry for "${letter}", which is not in letters.json.`
      );
    }
  }

  return { manifest, letters: manifest.letters, columns, images, shapesByLetter };
}

function svgToBuffer(svgContent, width, height) {
  return new Promise((resolve, reject) => {
    svg2img(svgContent, { resvg: { fitTo: { mode: "width", value: width } } }, (error, buffer) => {
      if (error) reject(error);
      else resolve(buffer);
    });
  });
}

/** Read a letter picture as a PNG buffer. SVG is rasterised through resvg. */
async function loadImage(imagePath, letter) {
  try {
    if (path.extname(imagePath).toLowerCase() === ".png") {
      return fs.readFileSync(imagePath);
    }
    const svgContent = fs.readFileSync(imagePath, "utf-8");
    return await svgToBuffer(svgContent, IMAGE_RENDER_WIDTH, IMAGE_RENDER_WIDTH);
  } catch (error) {
    console.log(`  Warning: could not load image for ${letter}: ${error.message}`);
    return null;
  }
}

// ---------------------------------------------------------------------------
// Page builders
// ---------------------------------------------------------------------------

function createLetterPage(letterData, imageBuffer) {
  const mainForm = letterData.mainForm || letterData.letter;

  const children = [
    // Title
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 400 },
      direction: "RTL",
      children: [
        new TextRun({
          text: content.manifest.pages.letterTitle.replace("{name}", letterData.name),
          font: FONT,
          size: 72,
          bold: true,
          color: "2E86AB",
        }),
      ],
    }),

    // Large letter
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 400 },
      direction: "RTL",
      children: [
        new TextRun({
          text: mainForm,
          font: FONT,
          size: 144,
          bold: true,
          color: "1B4965",
        }),
      ],
    }),
  ];

  if (imageBuffer) {
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 },
        direction: "RTL",
        children: [
          new ImageRun({
            data: imageBuffer,
            transformation: IMAGE_TRANSFORM,
            type: "png",
          }),
        ],
      })
    );
  }

  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
      direction: "RTL",
      children: [
        new TextRun({
          text: letterData.letter,
          font: FONT,
          size: 96,
          bold: true,
          color: "E74C3C",
        }),
        new TextRun({
          text: letterData.word.substring(1),
          font: FONT,
          size: 96,
          bold: true,
          color: "2C3E50",
        }),
      ],
    })
  );

  return children;
}

function createShapeCell(text, fontName, isBold, color) {
  return new TableCell({
    width: { size: 25, type: WidthType.PERCENTAGE },
    verticalAlign: VerticalAlign.CENTER,
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        direction: "RTL",
        children: [
          new TextRun({
            text,
            font: fontName,
            size: 48, // 24pt
            bold: isBold,
            color,
          }),
        ],
      }),
    ],
  });
}

function createTablePage(letterData) {
  const shapes = content.shapesByLetter.get(letterData.letter);
  const table = content.manifest.table;

  // One row of the four shapes in the given font (RTL order: right to left)
  const createShapeRow = (fontName, isBold = false, color = "000000", height = null) => {
    return new TableRow({
      ...(height && { height: { value: height, rule: HeightRule.EXACT } }),
      children: content.columns.map((column) =>
        createShapeCell(shapes[column.shape], fontName, isBold, color)
      ),
    });
  };

  // One blank practice row. Kept as its own builder so the empty run carries no
  // bold or color at all, exactly like the hand-written version did.
  const createEmptyRow = () => {
    return new TableRow({
      height: { value: table.emptyRowHeight, rule: HeightRule.EXACT },
      children: content.columns.map(
        () =>
          new TableCell({
            width: { size: 25, type: WidthType.PERCENTAGE },
            verticalAlign: VerticalAlign.CENTER,
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                direction: "RTL",
                children: [
                  new TextRun({
                    text: "",
                    font: FONT,
                    size: 28,
                  }),
                ],
              }),
            ],
          })
      ),
    });
  };

  const headerCell = (header) =>
    new TableCell({
      width: { size: 25, type: WidthType.PERCENTAGE },
      verticalAlign: VerticalAlign.CENTER,
      shading: { type: ShadingType.SOLID, color: "2E86AB" },
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          direction: "RTL",
          children: [
            new TextRun({
              text: header,
              font: FONT,
              size: 28,
              bold: true,
              color: "FFFFFF",
            }),
          ],
        }),
      ],
    });

  const tablePage = [
    // Title
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 400 },
      direction: "RTL",
      children: [
        new TextRun({
          text: content.manifest.pages.shapesTitle.replace("{name}", letterData.name),
          font: FONT,
          size: 72, // 36pt
          bold: true,
          color: "2E86AB",
        }),
      ],
    }),

    // Table
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 2, color: "34495E" },
        bottom: { style: BorderStyle.SINGLE, size: 2, color: "34495E" },
        left: { style: BorderStyle.SINGLE, size: 2, color: "34495E" },
        right: { style: BorderStyle.SINGLE, size: 2, color: "34495E" },
        insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: "34495E" },
        insideVertical: { style: BorderStyle.SINGLE, size: 2, color: "34495E" },
      },
      rows: [
        // Header row (RTL order: right to left)
        new TableRow({
          children: content.columns.map((column) => headerCell(column.header)),
        }),

        // Connected script row (Traditional Arabic)
        createShapeRow(FONT, false, "1B4965"),

        // Disconnected script rows (Dotted font for children)
        ...Array.from({ length: table.dottedRows }, () =>
          createShapeRow(FONT_DOTTED, false, "000000", table.dottedRowHeight)
        ),

        // Empty practice rows
        ...Array.from({ length: table.emptyRows }, () => createEmptyRow()),
      ],
    }),
  ];

  return tablePage;
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
            color: "34495E",
          }),
        ],
      }),
    ],
  });
}

// Create cover page with the cover image
function createCoverPage() {
  const cover = content.manifest.cover;
  const coverPath = cover.candidates
    .map((candidate) => path.join(ROOT, candidate))
    .find((candidate) => fs.existsSync(candidate));
  if (!coverPath) {
    console.log(`Warning: ${cover.file} not found, skipping cover page`);
    return null;
  }

  const coverBuffer = fs.readFileSync(coverPath);
  const imgDims = getPngDimensions(coverBuffer);

  // Full A5 page in pixels at 96dpi (no margins for full-bleed cover)
  const pageWidthPx = A5_WIDTH / 15;
  const pageHeightPx = A5_HEIGHT / 15;

  // Scale image to cover the whole page while preserving aspect ratio
  const scale = Math.max(pageWidthPx / imgDims.width, pageHeightPx / imgDims.height);
  const widthPx = Math.round(imgDims.width * scale);
  const heightPx = Math.round(imgDims.height * scale);

  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 0 },
      children: [
        new ImageRun({
          data: coverBuffer,
          transformation: {
            width: widthPx,
            height: heightPx,
          },
          type: "png",
        }),
      ],
    }),
  ];
}

// ---------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------

let content = null;

async function createWorkbook() {
  console.log("جاري إنشاء كتاب التمارين...");

  content = loadContent();

  const doc = new Document({
    creator: content.manifest.book.creator,
    title: content.manifest.book.title,
    description: content.manifest.book.description,

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

  // Create cover page section (full-bleed, no footer/page number)
  const coverContent = createCoverPage();
  if (coverContent) {
    doc.addSection({
      properties: {
        page: {
          size: {
            width: A5_WIDTH,
            height: A5_HEIGHT,
            orientation: PageOrientation.PORTRAIT,
          },
          margin: {
            top: 0,
            right: 0,
            bottom: 0,
            left: 0,
          },
        },
        direction: "RTL",
      },
      children: coverContent,
    });
  }

  // Create pages for each letter
  for (const letterData of content.letters) {
    console.log(`جاري إنشاء صفحة حرف ${letterData.name}...`);

    const imageBuffer = await loadImage(content.images.get(letterData.letter), letterData.letter);

    // Letter page
    const letterPageContent = createLetterPage(letterData, imageBuffer);

    // Table page
    const tablePageContent = createTablePage(letterData);

    // A fresh properties object per section: the docx library owns what it is
    // handed, so a shared one could be normalised twice.
    const contentProperties = () => ({
      page: {
        size: {
          width: A5_WIDTH,
          height: A5_HEIGHT,
          orientation: PageOrientation.PORTRAIT,
        },
        margin: {
          top: CONTENT_MARGIN,
          right: CONTENT_MARGIN,
          bottom: CONTENT_MARGIN,
          left: CONTENT_MARGIN,
        },
        borders: PAGE_FRAME,
      },
      direction: "RTL",
    });
    const footers = SHOW_PAGE_NUMBERS ? { footers: { default: createPageFooter() } } : {};

    doc.addSection({ properties: contentProperties(), ...footers, children: letterPageContent });
    doc.addSection({ properties: contentProperties(), ...footers, children: tablePageContent });
  }

  console.log("جاري حفظ الملف...");

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(OUTPUT, buffer);

  console.log("تم إنشاء الملف بنجاح!");
  console.log(`الملف: ${path.basename(OUTPUT)}`);
  console.log(`عدد الصفحات: ${content.letters.length * 2 + (coverContent ? 1 : 0)} صفحة`);
  console.log("حجم الصفحة: A5");
}

if (require.main === module) {
  createWorkbook().catch((error) => {
    console.error(`[FAIL] ${error.message}`);
    process.exit(1);
  });
}
