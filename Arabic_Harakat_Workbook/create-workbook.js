const docx = require("docx");
const fs = require("fs");
const path = require("path");

const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  ImageRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  HeightType,
  HeightRule,
  AlignmentType,
  BorderStyle,
  PageOrientation,
  convertInchesToTwip,
  VerticalAlign,
  ShadingType,
  PageBorderDisplay,
  PageBorderOffsetFrom,
  PageBorderZOrder,
} = docx;

// A5 dimensions in twips (14.8cm x 21cm)
const A5_WIDTH = convertInchesToTwip(5.83);
const A5_HEIGHT = convertInchesToTwip(8.27);

const MARGIN_TOP = convertInchesToTwip(0.3);
const MARGIN_BOTTOM = convertInchesToTwip(0.3);
const MARGIN_SIDE = convertInchesToTwip(0.5);

// Simple single-line frame around every page. The offset is measured from the
// paper edge, so the frame sits outside the text area.
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

const ROOT = __dirname;
const ASSETS = path.join(ROOT, "assets");
const LETTERS_FILE = path.join(ASSETS, "letters.json");
const HARAKAT_FILE = path.join(ASSETS, "harakat.json");
// A named folder writes to its own name instead of overwriting another book.
const OUTPUT = path.join(ROOT, `${path.basename(ROOT)}.docx`);

// The two page groups this generator can print. `sections` in letters.json picks
// which ones a project actually builds, so the letter tables and the syllable
// analysis can live in separate books without forking this script.
const SECTIONS = ["letters", "syllable"];

// ---------------------------------------------------------------------------
// Content
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

/** A page title has to carry the letter name placeholder. */
function requireTitle(value, where) {
  requireText(value, where);
  if (!value.includes("{name}")) {
    throw new Error(`${where} must contain the {name} placeholder.`);
  }
  return value;
}

function requireSymbolList(value, where) {
  if (!Array.isArray(value) || value.length === 0) {
    throw new Error(`${where} must be a non-empty array.`);
  }
  for (const [index, symbol] of value.entries()) {
    requireText(symbol, `${where}[${index}]`);
  }
  return value;
}

function requirePathList(value, where) {
  if (!Array.isArray(value) || value.length === 0) {
    throw new Error(`${where} must be a non-empty array.`);
  }
  for (const [index, entry] of value.entries()) {
    requireText(entry, `${where}[${index}]`);
  }
  return value;
}

/** First existing path from a list of paths relative to the project, or null. */
function resolveRelativePath(paths) {
  for (const relative of paths) {
    const full = path.resolve(ROOT, relative);
    if (fs.existsSync(full)) return full;
  }
  return null;
}

/** Read width/height from a PNG buffer (IHDR at bytes 16-23). */
function getPngDimensions(buffer) {
  return {
    width: buffer.readUInt32BE(16),
    height: buffer.readUInt32BE(20),
  };
}

/** Name of the haraka carrying `symbol`, used to label the hamza table. */
function harakaNameFor(symbol) {
  return content.harakat.find((haraka) => haraka.symbol === symbol).name;
}

/**
 * Read both data files and check everything the build depends on, so a typo in
 * the JSON fails here instead of quietly producing a wrong book.
 */
function loadContent() {
  const lettersFile = readJson(LETTERS_FILE);
  const harakatFile = readJson(HARAKAT_FILE);

  // -- letters.json --------------------------------------------------------
  for (const key of ["creator", "title", "description"]) {
    requireText(lettersFile.book && lettersFile.book[key], `book.${key}`);
  }
  const sections = lettersFile.sections;
  if (!Array.isArray(sections) || sections.length === 0) {
    throw new Error('letters.json must contain a non-empty "sections" array.');
  }
  const seenSections = new Set();
  for (const [index, section] of sections.entries()) {
    const at = `sections[${index}]`;
    requireText(section, at);
    if (!SECTIONS.includes(section)) {
      throw new Error(`${at}: "${section}" is not a known section (${SECTIONS.join(", ")}).`);
    }
    if (seenSections.has(section)) {
      throw new Error(`${at}: section "${section}" is listed more than once.`);
    }
    seenSections.add(section);
  }
  requireText(lettersFile.cover && lettersFile.cover.file, "cover.file");
  requirePathList(lettersFile.cover && lettersFile.cover.candidates, "cover.candidates");
  if (!Array.isArray(lettersFile.letters) || lettersFile.letters.length === 0) {
    throw new Error('letters.json must contain a non-empty "letters" array.');
  }
  const seen = new Set();
  for (const [index, entry] of lettersFile.letters.entries()) {
    const at = `letters[${index}]`;
    requireText(entry.letter, `${at}.letter`);
    requireText(entry.name, `${at}.name`);
    if (seen.has(entry.letter)) {
      throw new Error(`${at}: letter "${entry.letter}" is listed more than once.`);
    }
    seen.add(entry.letter);
  }
  if (!Array.isArray(lettersFile.syllableExcludedLetters)) {
    throw new Error("syllableExcludedLetters must be an array of letters.");
  }
  for (const letter of lettersFile.syllableExcludedLetters) {
    if (!seen.has(letter)) {
      throw new Error(`syllableExcludedLetters: "${letter}" is not in letters.json.`);
    }
  }

  // -- harakat.json --------------------------------------------------------
  if (!Array.isArray(harakatFile.harakat) || harakatFile.harakat.length === 0) {
    throw new Error('harakat.json must contain a non-empty "harakat" array.');
  }
  const symbols = new Set();
  for (const [index, haraka] of harakatFile.harakat.entries()) {
    const at = `harakat[${index}]`;
    requireText(haraka.name, `${at}.name`);
    requireText(haraka.symbol, `${at}.symbol`);
    requireText(haraka.description, `${at}.description`);
    if (symbols.has(haraka.symbol)) {
      throw new Error(`${at}: symbol "${haraka.symbol}" is listed more than once.`);
    }
    symbols.add(haraka.symbol);
    // A haraka either has a long vowel or it has none, never just one of them.
    const longVowel = haraka.longVowel ?? null;
    const longSymbol = haraka.longSymbol ?? null;
    if ((longVowel === null) !== (longSymbol === null)) {
      throw new Error(`${at}: longVowel and longSymbol must both be set or both be null.`);
    }
    if (longVowel !== null) {
      requireText(longVowel, `${at}.longVowel`);
      requireText(longSymbol, `${at}.longSymbol`);
    }
  }

  if (!Array.isArray(harakatFile.baseLetterOverrides)) {
    throw new Error("baseLetterOverrides must be an array.");
  }
  for (const [index, override] of harakatFile.baseLetterOverrides.entries()) {
    const at = `baseLetterOverrides[${index}]`;
    requireText(override.letter, `${at}.letter`);
    requireText(override.haraka, `${at}.haraka`);
    requireText(override.base, `${at}.base`);
    if (!seen.has(override.letter)) {
      throw new Error(`${at}: letter "${override.letter}" is not in letters.json.`);
    }
    if (!symbols.has(override.haraka)) {
      throw new Error(`${at}: haraka "${override.haraka}" is not in harakat.`);
    }
  }

  const display = harakatFile.displayOrder;
  if (!display || typeof display !== "object") {
    throw new Error('harakat.json is missing the "displayOrder" section.');
  }
  for (const key of ["hamzaTableSymbols", "harakaSymbols"]) {
    requireSymbolList(display[key], `displayOrder.${key}`);
    for (const symbol of display[key]) {
      if (!symbols.has(symbol)) {
        throw new Error(`displayOrder.${key}: "${symbol}" is not a haraka symbol.`);
      }
    }
  }
  requireSymbolList(display.longVowelDottedSymbols, "displayOrder.longVowelDottedSymbols");
  requireSymbolList(display.syllableVowels, "displayOrder.syllableVowels");

  const texts = harakatFile.texts;
  if (!texts || typeof texts !== "object") {
    throw new Error('harakat.json is missing the "texts" section.');
  }
  for (const key of ["letterTitle", "practiceTitle"]) {
    requireTitle(texts[key], `texts.${key}`);
  }
  for (const key of ["letterSubtitle", "syllableSectionTitle", "syllableSectionSubtitle", "hamzaLetter"]) {
    requireText(texts[key], `texts.${key}`);
  }
  const headers = texts.tableHeaders;
  if (!headers || typeof headers !== "object") {
    throw new Error('harakat.json is missing the "texts.tableHeaders" section.');
  }
  for (const key of ["letterWithLongVowel", "longVowel", "letterWithHaraka", "haraka"]) {
    requireText(headers[key], `texts.tableHeaders.${key}`);
  }

  return {
    book: lettersFile.book,
    sections,
    cover: lettersFile.cover,
    letters: lettersFile.letters,
    syllableExcludedLetters: lettersFile.syllableExcludedLetters,
    harakat: harakatFile.harakat,
    baseLetterOverrides: harakatFile.baseLetterOverrides,
    displayOrder: display,
    texts,
  };
}

let content = null;

// For the hamza letter (أ) the kasra is written as إ (hamza below the alef).
// Which letter replaces which base comes from baseLetterOverrides in the data.
function getHarakaBaseLetter(letter, harakaSymbol) {
  const override = content.baseLetterOverrides.find(
    (entry) => entry.letter === letter && entry.haraka === harakaSymbol
  );
  return override ? override.base : letter;
}

// Create display page (page 1 for each letter)
function createDisplayPage(letterData) {
  const letter = letterData.letter;
  
  const children = [
    // Title
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
      direction: "RTL",
      children: [
        new TextRun({
          text: content.texts.letterTitle.replace("{name}", letterData.name),
          font: "Traditional Arabic",
          size: 72,
          bold: true,
          color: "2E86AB",
        }),
      ],
    }),

    // Large letter
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
      direction: "RTL",
      children: [
        new TextRun({
          text: letter,
          font: "Traditional Arabic",
          size: 144,
          bold: true,
          color: "1B4965",
        }),
      ],
    }),

    // Subtitle
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
      direction: "RTL",
      children: [
        new TextRun({
          text: content.texts.letterSubtitle,
          font: "Traditional Arabic",
          size: 48,
          bold: true,
          color: "2E86AB",
        }),
      ],
    }),
  ];

  // Create table
  const isHamza = letter === "أ";
  const tableRows = [];

  if (isHamza) {
    // Hamza: transposed table (2 rows x 4 columns)
    // Row 1: haraka names, Row 2: letters with harakat
    const hamzaSymbols = content.displayOrder.hamzaTableSymbols;
    const hamzaNames = hamzaSymbols.map((symbol) => harakaNameFor(symbol));
    const hamzaColors = ["E74C3C", "3498DB", "27AE60", "9B59B6"];

    tableRows.push(
      new TableRow({
        children: hamzaNames.map((name) =>
          new TableCell({
            width: { size: 25, type: WidthType.PERCENTAGE },
            verticalAlign: VerticalAlign.CENTER,
            height: { size: 800, rule: HeightRule.ATLEAST },
            shading: { type: ShadingType.SOLID, color: "F5F5F5" },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                direction: "RTL",
                children: [
                  new TextRun({
                    text: name,
                    font: "Traditional Arabic",
                    size: 32,
                    bold: true,
                    color: "1B4965",
                  }),
                ],
              }),
            ],
          })
        ),
      })
    );

    tableRows.push(
      new TableRow({
        children: hamzaSymbols.map((symbol, index) =>
          new TableCell({
            width: { size: 25, type: WidthType.PERCENTAGE },
            verticalAlign: VerticalAlign.CENTER,
            height: { size: 800, rule: HeightRule.ATLEAST },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                direction: "RTL",
                children: [
                  new TextRun({
                    text: getHarakaBaseLetter(letter, symbol) + symbol,
                    font: "Traditional Arabic",
                    size: 56,
                    bold: true,
                    color: hamzaColors[index],
                  }),
                ],
              }),
            ],
          })
        ),
      })
    );
  } else {
  // Header row (RTL order: right to left)
  tableRows.push(
    new TableRow({
      children: [
        // الحرف بحرف المد - rightmost
        new TableCell({
          width: { size: 40, type: WidthType.PERCENTAGE },
          verticalAlign: VerticalAlign.CENTER,
          shading: { type: ShadingType.SOLID, color: "2E86AB" },
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              direction: "RTL",
              children: [
                new TextRun({
                  text: content.texts.tableHeaders.letterWithLongVowel,
                  font: "Traditional Arabic",
                  size: 28,
                  bold: true,
                  color: "FFFFFF",
                }),
              ],
            }),
          ],
        }),
        // حرف المد
        new TableCell({
          width: { size: 15, type: WidthType.PERCENTAGE },
          verticalAlign: VerticalAlign.CENTER,
          shading: { type: ShadingType.SOLID, color: "2E86AB" },
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              direction: "RTL",
              children: [
                new TextRun({
                  text: content.texts.tableHeaders.longVowel,
                  font: "Traditional Arabic",
                  size: 28,
                  bold: true,
                  color: "FFFFFF",
                }),
              ],
            }),
          ],
        }),
        // الحرف بالحركة
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
                  text: content.texts.tableHeaders.letterWithHaraka,
                  font: "Traditional Arabic",
                  size: 28,
                  bold: true,
                  color: "FFFFFF",
                }),
              ],
            }),
          ],
        }),
        // الحركة - leftmost
        new TableCell({
          width: { size: 20, type: WidthType.PERCENTAGE },
          verticalAlign: VerticalAlign.CENTER,
          shading: { type: ShadingType.SOLID, color: "2E86AB" },
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              direction: "RTL",
              children: [
                new TextRun({
                  text: content.texts.tableHeaders.haraka,
                  font: "Traditional Arabic",
                  size: 28,
                  bold: true,
                  color: "FFFFFF",
                }),
              ],
            }),
          ],
        }),
      ],
    })
  );

  // Data rows for each haraka (RTL order: right to left)
    content.harakat.forEach((haraka) => {
    const letterWithHaraka = getHarakaBaseLetter(letter, haraka.symbol) + haraka.symbol;
    // For hamza, we don't show long vowels (المدود)
    const letterWithLongVowel = isHamza ? "-" : (haraka.longSymbol ? letter + haraka.longSymbol : "-");
    const longVowelName = isHamza ? "-" : (haraka.longVowel || "-");
    

    tableRows.push(
      new TableRow({
        children: [
          // الحرف بحرف المد - rightmost
          new TableCell({
            width: { size: 40, type: WidthType.PERCENTAGE },
            verticalAlign: VerticalAlign.CENTER,
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                direction: "RTL",
                children: [
                  new TextRun({
                    text: letterWithLongVowel,
                    font: "Traditional Arabic",
                    size: 56,
                    bold: true,
                    color: "8E44AD",
                  }),
                ],
              }),
            ],
          }),
          // حرف المد
          new TableCell({
            width: { size: 15, type: WidthType.PERCENTAGE },
            verticalAlign: VerticalAlign.CENTER,
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                direction: "RTL",
                children: [
                  new TextRun({
                    text: longVowelName,
                    font: "Traditional Arabic",
                    size: 32,
                    bold: true,
                    color: "27AE60",
                  }),
                ],
              }),
            ],
          }),
          // الحرف بالحركة
          new TableCell({
            width: { size: 25, type: WidthType.PERCENTAGE },
            verticalAlign: VerticalAlign.CENTER,
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                direction: "RTL",
                children: [
                  new TextRun({
                    text: letterWithHaraka,
                    font: "Traditional Arabic",
                    size: 56,
                    bold: true,
                    color: "E74C3C",
                  }),
                ],
              }),
            ],
          }),
          // الحركة - leftmost
          new TableCell({
            width: { size: 20, type: WidthType.PERCENTAGE },
            verticalAlign: VerticalAlign.CENTER,
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                direction: "RTL",
                children: [
                  new TextRun({
                    text: haraka.name,
                    font: "Traditional Arabic",
                    size: 32,
                    bold: true,
                    color: "1B4965",
                  }),
                ],
              }),
            ],
          }),
        ],
      })
    );
  });
  }

  children.push(
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
      rows: tableRows,
    })
  );

  return children;
}

// Transpose a grid of cells: grid[row][col] becomes a row where the old rightmost
// column is the LAST row, so the new top-to-bottom order is فتحة، ضمة، كسرة، سكون
// (and ألف، واو، ياء for the long-vowel table).
// grid[row][col] keeps col 0 = rightmost, so the table that receives the result
// must be created with visuallyRightToLeft: true to read as متصل، منقّط، فارغ، فارغ.
function transposeGrid(grid) {
  const cols = grid[0].length;
  const rows = [];
  for (let c = cols - 1; c >= 0; c--) {
    rows.push(
      new TableRow({
        height: { value: 800, rule: HeightRule.ATLEAST },
        children: grid.map((gridRow) => gridRow[c]),
      })
    );
  }
  return rows;
}

// Create practice page (page 2 for each letter)
function createPracticePage(letterData) {
  const letter = letterData.letter;
  
  const children = [
    // Title
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
      direction: "RTL",
      children: [
        new TextRun({
          text: content.texts.practiceTitle.replace("{name}", letterData.name),
          font: "Traditional Arabic",
          size: 72,
          bold: true,
          color: "2E86AB",
        }),
      ],
    }),
  ];

  // Table borders
  const tableBorders = {
    top: { style: BorderStyle.SINGLE, size: 2, color: "34495E" },
    bottom: { style: BorderStyle.SINGLE, size: 2, color: "34495E" },
    left: { style: BorderStyle.SINGLE, size: 2, color: "34495E" },
    right: { style: BorderStyle.SINGLE, size: 2, color: "34495E" },
    insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: "34495E" },
    insideVertical: { style: BorderStyle.SINGLE, size: 2, color: "34495E" },
  };

  // ===== Table 1: Harakat - after transposing: 4 rows (one per haraka) x 4 columns =====
  const harakatGrid = []; // grid[row][col], RTL: col 0 = rightmost

  // Grid row 0: letter + harakat in the solid font (RTL cells: sukun, kasra, damma, fatha)
  harakatGrid.push([
    // بْ - rightmost
    new TableCell({
      width: { size: 25, type: WidthType.PERCENTAGE },
      verticalAlign: VerticalAlign.CENTER,
      shading: { type: ShadingType.SOLID, color: "F5F5F5" },
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          direction: "RTL",
          children: [
            new TextRun({
              text: getHarakaBaseLetter(letter, "\u0652") + "\u0652",
              font: "Traditional Arabic",
              size: 56,
              bold: true,
              color: "9B59B6",
            }),
          ],
        }),
      ],
    }),
    // بِ
    new TableCell({
      width: { size: 25, type: WidthType.PERCENTAGE },
      verticalAlign: VerticalAlign.CENTER,
      shading: { type: ShadingType.SOLID, color: "F5F5F5" },
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          direction: "RTL",
          children: [
            new TextRun({
              text: getHarakaBaseLetter(letter, "\u0650") + "\u0650",
              font: "Traditional Arabic",
              size: 56,
              bold: true,
              color: "27AE60",
            }),
          ],
        }),
      ],
    }),
    // بُ
    new TableCell({
      width: { size: 25, type: WidthType.PERCENTAGE },
      verticalAlign: VerticalAlign.CENTER,
      shading: { type: ShadingType.SOLID, color: "F5F5F5" },
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          direction: "RTL",
          children: [
            new TextRun({
              text: getHarakaBaseLetter(letter, "\u064F") + "\u064F",
              font: "Traditional Arabic",
              size: 56,
              bold: true,
              color: "3498DB",
            }),
          ],
        }),
      ],
    }),
    // بَ - leftmost
    new TableCell({
      width: { size: 25, type: WidthType.PERCENTAGE },
      verticalAlign: VerticalAlign.CENTER,
      shading: { type: ShadingType.SOLID, color: "F5F5F5" },
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          direction: "RTL",
          children: [
            new TextRun({
              text: getHarakaBaseLetter(letter, "\u064E") + "\u064E",
              font: "Traditional Arabic",
              size: 56,
              bold: true,
              color: "E74C3C",
            }),
          ],
        }),
      ],
    }),
  ]);

  // Grid row 1: the same 4 harakat in the dotted font (RTL cells: sukun, kasra, damma, fatha)
  const harakaColors = ["E74C3C", "3498DB", "27AE60", "9B59B6"];
  const harakaSymbols = content.displayOrder.harakaSymbols; // بْ, بِ, بُ, بَ
  
  harakatGrid.push([
    // بْ (dotted) - rightmost
    new TableCell({
      width: { size: 25, type: WidthType.PERCENTAGE },
      verticalAlign: VerticalAlign.CENTER,
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          direction: "RTL",
          children: [
            new TextRun({
              text: getHarakaBaseLetter(letter, harakaSymbols[0]) + harakaSymbols[0],
              font: "KFGQPC Alphabet Dotted",
              size: 56,
              color: harakaColors[3],
            }),
          ],
        }),
      ],
    }),
    // بِ (dotted)
    new TableCell({
      width: { size: 25, type: WidthType.PERCENTAGE },
      verticalAlign: VerticalAlign.CENTER,
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          direction: "RTL",
          children: [
            new TextRun({
              text: getHarakaBaseLetter(letter, harakaSymbols[1]) + harakaSymbols[1],
              font: "KFGQPC Alphabet Dotted",
              size: 56,
              color: harakaColors[2],
            }),
          ],
        }),
      ],
    }),
    // بُ (dotted)
    new TableCell({
      width: { size: 25, type: WidthType.PERCENTAGE },
      verticalAlign: VerticalAlign.CENTER,
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          direction: "RTL",
          children: [
            new TextRun({
              text: getHarakaBaseLetter(letter, harakaSymbols[2]) + harakaSymbols[2],
              font: "KFGQPC Alphabet Dotted",
              size: 56,
              color: harakaColors[1],
            }),
          ],
        }),
      ],
    }),
    // بَ (dotted) - leftmost
    new TableCell({
      width: { size: 25, type: WidthType.PERCENTAGE },
      verticalAlign: VerticalAlign.CENTER,
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          direction: "RTL",
          children: [
            new TextRun({
              text: getHarakaBaseLetter(letter, harakaSymbols[3]) + harakaSymbols[3],
              font: "KFGQPC Alphabet Dotted",
              size: 56,
              color: harakaColors[0],
            }),
          ],
        }),
      ],
    }),
  ]);

  // Two empty writing rows - they become the last two columns after transposing
  for (let i = 0; i < 2; i++) {
    harakatGrid.push([
    new TableCell({
      width: { size: 25, type: WidthType.PERCENTAGE },
      verticalAlign: VerticalAlign.CENTER,
      children: [new Paragraph({ direction: "RTL", children: [] })],
    }),
    new TableCell({
      width: { size: 25, type: WidthType.PERCENTAGE },
      verticalAlign: VerticalAlign.CENTER,
      children: [new Paragraph({ direction: "RTL", children: [] })],
    }),
    new TableCell({
      width: { size: 25, type: WidthType.PERCENTAGE },
      verticalAlign: VerticalAlign.CENTER,
      children: [new Paragraph({ direction: "RTL", children: [] })],
    }),
    new TableCell({
      width: { size: 25, type: WidthType.PERCENTAGE },
      verticalAlign: VerticalAlign.CENTER,
      children: [new Paragraph({ direction: "RTL", children: [] })],
    }),
  ]);
  }

  // Add first table
  children.push(
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: tableBorders,
      rows: transposeGrid(harakatGrid),
      visuallyRightToLeft: true,
    })
  );

  // Spacing between tables
  children.push(
    new Paragraph({
      spacing: { after: 200 },
      direction: "RTL",
      children: [],
    })
  );

  // ===== Table 2: Long vowels - after transposing: 3 rows (one per long vowel) x 4 columns =====
  // For the hamza letter, use a table of the hamza on the line (ء) with the 4 harakat instead
  const isHamzaLetter = letter === "أ";
  const longVowelGrid = []; // grid[row][col], RTL: col 0 = rightmost

  if (isHamzaLetter) {
    // Hamza on the line (ء) with the 4 harakat
    // Grid row 0: hamza on the line + harakat, solid font (RTL: sukun, kasra, damma, fatha)
    const hamzaHarakaSymbols = content.displayOrder.harakaSymbols;
  const hamzaLetter = content.texts.hamzaLetter; // ءْ, ءِ, ءُ, ءَ
    const hamzaHarakaColors = ["9B59B6", "27AE60", "3498DB", "E74C3C"];

    longVowelGrid.push([
      // ءْ - rightmost
      new TableCell({
        width: { size: 25, type: WidthType.PERCENTAGE },
        verticalAlign: VerticalAlign.CENTER,
        shading: { type: ShadingType.SOLID, color: "F5F5F5" },
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            direction: "RTL",
            children: [
              new TextRun({
                text: hamzaLetter + hamzaHarakaSymbols[0],
                font: "Traditional Arabic",
                size: 56,
                bold: true,
                color: hamzaHarakaColors[0],
              }),
            ],
          }),
        ],
      }),
      // ءِ
      new TableCell({
        width: { size: 25, type: WidthType.PERCENTAGE },
        verticalAlign: VerticalAlign.CENTER,
        shading: { type: ShadingType.SOLID, color: "F5F5F5" },
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            direction: "RTL",
            children: [
              new TextRun({
                text: hamzaLetter + hamzaHarakaSymbols[1],
                font: "Traditional Arabic",
                size: 56,
                bold: true,
                color: hamzaHarakaColors[1],
              }),
            ],
          }),
        ],
      }),
      // ءُ
      new TableCell({
        width: { size: 25, type: WidthType.PERCENTAGE },
        verticalAlign: VerticalAlign.CENTER,
        shading: { type: ShadingType.SOLID, color: "F5F5F5" },
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            direction: "RTL",
            children: [
              new TextRun({
                text: hamzaLetter + hamzaHarakaSymbols[2],
                font: "Traditional Arabic",
                size: 56,
                bold: true,
                color: hamzaHarakaColors[2],
              }),
            ],
          }),
        ],
      }),
      // ءَ - leftmost
      new TableCell({
        width: { size: 25, type: WidthType.PERCENTAGE },
        verticalAlign: VerticalAlign.CENTER,
        shading: { type: ShadingType.SOLID, color: "F5F5F5" },
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            direction: "RTL",
            children: [
              new TextRun({
                text: hamzaLetter + hamzaHarakaSymbols[3],
                font: "Traditional Arabic",
                size: 56,
                bold: true,
                color: hamzaHarakaColors[3],
              }),
            ],
          }),
        ],
      }),
    ]);

    // Grid row 1: the same 4 harakat in the dotted font
    longVowelGrid.push([
      new TableCell({
        width: { size: 25, type: WidthType.PERCENTAGE },
        verticalAlign: VerticalAlign.CENTER,
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            direction: "RTL",
            children: [
              new TextRun({
                text: hamzaLetter + hamzaHarakaSymbols[0],
                font: "KFGQPC Alphabet Dotted",
                size: 56,
                color: hamzaHarakaColors[0],
              }),
            ],
          }),
        ],
      }),
      new TableCell({
        width: { size: 25, type: WidthType.PERCENTAGE },
        verticalAlign: VerticalAlign.CENTER,
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            direction: "RTL",
            children: [
              new TextRun({
                text: hamzaLetter + hamzaHarakaSymbols[1],
                font: "KFGQPC Alphabet Dotted",
                size: 56,
                color: hamzaHarakaColors[1],
              }),
            ],
          }),
        ],
      }),
      new TableCell({
        width: { size: 25, type: WidthType.PERCENTAGE },
        verticalAlign: VerticalAlign.CENTER,
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            direction: "RTL",
            children: [
              new TextRun({
                text: hamzaLetter + hamzaHarakaSymbols[2],
                font: "KFGQPC Alphabet Dotted",
                size: 56,
                color: hamzaHarakaColors[2],
              }),
            ],
          }),
        ],
      }),
      new TableCell({
        width: { size: 25, type: WidthType.PERCENTAGE },
        verticalAlign: VerticalAlign.CENTER,
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            direction: "RTL",
            children: [
              new TextRun({
                text: hamzaLetter + hamzaHarakaSymbols[3],
                font: "KFGQPC Alphabet Dotted",
                size: 56,
                color: hamzaHarakaColors[3],
              }),
            ],
          }),
        ],
      }),
    ]);

    // Two empty writing rows - they become the last two columns after transposing
    for (let i = 0; i < 2; i++) {
      longVowelGrid.push([
      new TableCell({
        width: { size: 25, type: WidthType.PERCENTAGE },
        verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({ direction: "RTL", children: [] })],
      }),
      new TableCell({
        width: { size: 25, type: WidthType.PERCENTAGE },
        verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({ direction: "RTL", children: [] })],
      }),
      new TableCell({
        width: { size: 25, type: WidthType.PERCENTAGE },
        verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({ direction: "RTL", children: [] })],
      }),
      new TableCell({
        width: { size: 25, type: WidthType.PERCENTAGE },
        verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({ direction: "RTL", children: [] })],
      }),
    ]);
    }
  } else {
    // Grid row 0: letter + long vowel, solid font (RTL cells: yeh, waw, alef)
    longVowelGrid.push([
    // بِي - rightmost
    new TableCell({
      width: { size: 25, type: WidthType.PERCENTAGE },
      verticalAlign: VerticalAlign.CENTER,
      shading: { type: ShadingType.SOLID, color: "F5F5F5" },
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          direction: "RTL",
          children: [
            new TextRun({
              text: letter + "\u064A",
              font: "Traditional Arabic",
              size: 56,
              bold: true,
              color: "27AE60",
            }),
          ],
        }),
      ],
    }),
    // بُو
    new TableCell({
      width: { size: 25, type: WidthType.PERCENTAGE },
      verticalAlign: VerticalAlign.CENTER,
      shading: { type: ShadingType.SOLID, color: "F5F5F5" },
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          direction: "RTL",
          children: [
            new TextRun({
              text: letter + "\u0648",
              font: "Traditional Arabic",
              size: 56,
              bold: true,
              color: "3498DB",
            }),
          ],
        }),
      ],
    }),
    // بَا - leftmost
    new TableCell({
      width: { size: 25, type: WidthType.PERCENTAGE },
      verticalAlign: VerticalAlign.CENTER,
      shading: { type: ShadingType.SOLID, color: "F5F5F5" },
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          direction: "RTL",
          children: [
            new TextRun({
              text: letter + "\u0627",
              font: "Traditional Arabic",
              size: 56,
              bold: true,
              color: "E74C3C",
            }),
          ],
        }),
      ],
    }),
  ]);

  // Grid row 1: the same 3 long vowels in the dotted font (RTL cells: yeh, waw, alef)
  const longVowelSymbols = content.displayOrder.longVowelDottedSymbols; // بِي, بُو, بَا
  const longVowelColors = ["27AE60", "3498DB", "E74C3C"];
  
  longVowelGrid.push([
    // بِي (dotted) - rightmost
    new TableCell({
      width: { size: 25, type: WidthType.PERCENTAGE },
      verticalAlign: VerticalAlign.CENTER,
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          direction: "RTL",
          children: [
            new TextRun({
              text: letter + longVowelSymbols[0],
              font: "KFGQPC Alphabet Dotted",
              size: 56,
              color: longVowelColors[0],
            }),
          ],
        }),
      ],
    }),
    // بُو (dotted)
    new TableCell({
      width: { size: 25, type: WidthType.PERCENTAGE },
      verticalAlign: VerticalAlign.CENTER,
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          direction: "RTL",
          children: [
            new TextRun({
              text: letter + longVowelSymbols[1],
              font: "KFGQPC Alphabet Dotted",
              size: 56,
              color: longVowelColors[1],
            }),
          ],
        }),
      ],
    }),
    // بَا (dotted) - leftmost
    new TableCell({
      width: { size: 25, type: WidthType.PERCENTAGE },
      verticalAlign: VerticalAlign.CENTER,
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          direction: "RTL",
          children: [
            new TextRun({
              text: letter + longVowelSymbols[2],
              font: "KFGQPC Alphabet Dotted",
              size: 56,
              color: longVowelColors[2],
            }),
          ],
        }),
      ],
    }),
  ]);

  // Two empty writing rows - they become the last two columns after transposing
  for (let i = 0; i < 2; i++) {
    longVowelGrid.push([
    new TableCell({
      width: { size: 25, type: WidthType.PERCENTAGE },
      verticalAlign: VerticalAlign.CENTER,
      children: [new Paragraph({ direction: "RTL", children: [] })],
    }),
    new TableCell({
      width: { size: 25, type: WidthType.PERCENTAGE },
      verticalAlign: VerticalAlign.CENTER,
      children: [new Paragraph({ direction: "RTL", children: [] })],
    }),
    new TableCell({
      width: { size: 25, type: WidthType.PERCENTAGE },
      verticalAlign: VerticalAlign.CENTER,
      children: [new Paragraph({ direction: "RTL", children: [] })],
    }),
    new TableCell({
      width: { size: 25, type: WidthType.PERCENTAGE },
      verticalAlign: VerticalAlign.CENTER,
      children: [new Paragraph({ direction: "RTL", children: [] })],
    }),
  ]);
  }
  }

  // Add second table (shared by both branches: hamza-on-the-line or long vowels)
  children.push(
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: tableBorders,
      rows: transposeGrid(longVowelGrid),
      visuallyRightToLeft: true,
    })
  );

  return children;
}

// Create syllable analysis exercise for a single letter
function createSyllableExercise(letterData) {
  const letter = letterData.letter;
  const children = [];

  // Exercise title
  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 100 },
      direction: "RTL",
      children: [
        new TextRun({
          text: content.texts.letterTitle.replace("{name}", letterData.name),
          font: "Traditional Arabic",
          size: 36,
          bold: true,
          color: "2E86AB",
        }),
      ],
    })
  );

  // All syllables on one line
  const longVowelExamples = content.displayOrder.syllableVowels; // ا, و, ي
  const exerciseColors = ["E74C3C", "3498DB", "27AE60"];
  
  const syllables = [
    { text: letter + longVowelExamples[2], color: exerciseColors[2] }, // بِي (leftmost)
    { text: letter + longVowelExamples[1], color: exerciseColors[1] }, // بُو (middle)
    { text: letter + longVowelExamples[0], color: exerciseColors[0] }, // بَا (rightmost)
  ];

  const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  const noBorders = {
    top: noBorder, bottom: noBorder,
    left: noBorder, right: noBorder,
    insideHorizontal: noBorder, insideVertical: noBorder,
  };

  const tableBorders = {
    top: { style: BorderStyle.SINGLE, size: 2, color: "34495E" },
    bottom: { style: BorderStyle.SINGLE, size: 2, color: "34495E" },
    left: { style: BorderStyle.SINGLE, size: 2, color: "34495E" },
    right: { style: BorderStyle.SINGLE, size: 2, color: "34495E" },
    insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: "34495E" },
    insideVertical: { style: BorderStyle.SINGLE, size: 2, color: "34495E" },
  };

  children.push(
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      direction: "RTL",
      borders: noBorders,
      rows: [
        // Row 1: Syllables (no borders)
        new TableRow({
          children: [
            // بَا (spans 2 columns)
            new TableCell({
              width: { size: 30, type: WidthType.PERCENTAGE },
              verticalAlign: VerticalAlign.CENTER,
              height: { size: 700, rule: HeightRule.ATLEAST },
              columnSpan: 2,
              borders: noBorders,
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  direction: "RTL",
                  children: [
                    new TextRun({
                      text: syllables[0].text,
                      font: "Traditional Arabic",
                      size: 56,
                      bold: true,
                      color: syllables[0].color,
                    }),
                  ],
                }),
              ],
            }),
            // Spacer (no borders)
            new TableCell({
              width: { size: 5, type: WidthType.PERCENTAGE },
              verticalAlign: VerticalAlign.CENTER,
              height: { size: 700, rule: HeightRule.ATLEAST },
              borders: noBorders,
              children: [new Paragraph({ direction: "RTL", children: [] })],
            }),
            // بُو (spans 2 columns)
            new TableCell({
              width: { size: 30, type: WidthType.PERCENTAGE },
              verticalAlign: VerticalAlign.CENTER,
              height: { size: 700, rule: HeightRule.ATLEAST },
              columnSpan: 2,
              borders: noBorders,
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  direction: "RTL",
                  children: [
                    new TextRun({
                      text: syllables[1].text,
                      font: "Traditional Arabic",
                      size: 56,
                      bold: true,
                      color: syllables[1].color,
                    }),
                  ],
                }),
              ],
            }),
            // Spacer (no borders)
            new TableCell({
              width: { size: 5, type: WidthType.PERCENTAGE },
              verticalAlign: VerticalAlign.CENTER,
              height: { size: 700, rule: HeightRule.ATLEAST },
              borders: noBorders,
              children: [new Paragraph({ direction: "RTL", children: [] })],
            }),
            // بِي (spans 2 columns)
            new TableCell({
              width: { size: 30, type: WidthType.PERCENTAGE },
              verticalAlign: VerticalAlign.CENTER,
              height: { size: 700, rule: HeightRule.ATLEAST },
              columnSpan: 2,
              borders: noBorders,
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  direction: "RTL",
                  children: [
                    new TextRun({
                      text: syllables[2].text,
                      font: "Traditional Arabic",
                      size: 56,
                      bold: true,
                      color: syllables[2].color,
                    }),
                  ],
                }),
              ],
            }),
          ],
        }),
        // Row 2: 8 cells (6 with borders + 2 spacers)
        new TableRow({
          children: [
            // 2 cells for بَا
            new TableCell({
              width: { size: 15, type: WidthType.PERCENTAGE },
              verticalAlign: VerticalAlign.CENTER,
              height: { size: 500, rule: HeightRule.ATLEAST },
              borders: tableBorders,
              children: [new Paragraph({ direction: "RTL", children: [] })],
            }),
            new TableCell({
              width: { size: 15, type: WidthType.PERCENTAGE },
              verticalAlign: VerticalAlign.CENTER,
              height: { size: 500, rule: HeightRule.ATLEAST },
              borders: tableBorders,
              children: [new Paragraph({ direction: "RTL", children: [] })],
            }),
            // Spacer
            new TableCell({
              width: { size: 5, type: WidthType.PERCENTAGE },
              verticalAlign: VerticalAlign.CENTER,
              height: { size: 500, rule: HeightRule.ATLEAST },
              borders: noBorders,
              children: [new Paragraph({ direction: "RTL", children: [] })],
            }),
            // 2 cells for بُو
            new TableCell({
              width: { size: 15, type: WidthType.PERCENTAGE },
              verticalAlign: VerticalAlign.CENTER,
              height: { size: 500, rule: HeightRule.ATLEAST },
              borders: tableBorders,
              children: [new Paragraph({ direction: "RTL", children: [] })],
            }),
            new TableCell({
              width: { size: 15, type: WidthType.PERCENTAGE },
              verticalAlign: VerticalAlign.CENTER,
              height: { size: 500, rule: HeightRule.ATLEAST },
              borders: tableBorders,
              children: [new Paragraph({ direction: "RTL", children: [] })],
            }),
            // Spacer
            new TableCell({
              width: { size: 5, type: WidthType.PERCENTAGE },
              verticalAlign: VerticalAlign.CENTER,
              height: { size: 500, rule: HeightRule.ATLEAST },
              borders: noBorders,
              children: [new Paragraph({ direction: "RTL", children: [] })],
            }),
            // 2 cells for بِي
            new TableCell({
              width: { size: 15, type: WidthType.PERCENTAGE },
              verticalAlign: VerticalAlign.CENTER,
              height: { size: 500, rule: HeightRule.ATLEAST },
              borders: tableBorders,
              children: [new Paragraph({ direction: "RTL", children: [] })],
            }),
            new TableCell({
              width: { size: 15, type: WidthType.PERCENTAGE },
              verticalAlign: VerticalAlign.CENTER,
              height: { size: 500, rule: HeightRule.ATLEAST },
              borders: tableBorders,
              children: [new Paragraph({ direction: "RTL", children: [] })],
            }),
          ],
        }),
      ],
    })
  );

  return children;
}

/** Every section in this book is A5 + RTL; only margins and the frame vary. */
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

/**
 * Full-bleed cover, looked up in the order given by letters.json. Returns null
 * when no candidate exists, so a missing cover skips the page instead of
 * failing the build.
 */
function createCoverPage() {
  const coverPath = resolveRelativePath(content.cover.candidates);
  if (!coverPath) {
    console.log(`Warning: ${content.cover.file} not found, skipping cover page`);
    return null;
  }

  const coverBuffer = fs.readFileSync(coverPath);
  const dims = getPngDimensions(coverBuffer);

  // Full A5 page in pixels at 96 dpi (twips / 15 = px).
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

// Main function to create the workbook
async function createWorkbook() {
  content = loadContent();

  console.log(`جاري إنشاء ${content.book.title}...`);

  const doc = new Document({
    creator: content.book.creator,
    title: content.book.title,
    description: content.book.description,

    styles: {
      default: {
        document: {
          run: {
            font: "Traditional Arabic",
            size: 28,
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
      // oversized and centred so it bleeds off every edge. Leaving a side
      // margin here shifts the image and leaves a white band. The page frame
      // is skipped too, so the cover stays a full-bleed image.
      properties: pageProperties(0, 0, 0, false),
      children: coverContent,
    });
  }

  // Create pages for each letter
  if (content.sections.includes("letters")) {
    for (let i = 0; i < content.letters.length; i++) {
      const letterData = content.letters[i];

      console.log(`جاري إنشاء صفحة حرف ${letterData.name}...`);

      // Display page (page 1)
      const displayPageContent = createDisplayPage(letterData);

      // Practice page (page 2)
      const practicePageContent = createPracticePage(letterData);

      // Add section for display page
      doc.addSection({
        properties: pageProperties(MARGIN_TOP, MARGIN_BOTTOM),
        children: displayPageContent,
      });

      // Add section for practice page
      doc.addSection({
        properties: pageProperties(MARGIN_TOP, MARGIN_BOTTOM),
        children: practicePageContent,
      });
    }
  }

  // ===== Add syllable exercises section at the end =====
  const exercisesPerPage = 4;
  const syllableLetters = content.letters.filter(
    (entry) => !content.syllableExcludedLetters.includes(entry.letter)
  );
  let syllablePages = 0;

  if (content.sections.includes("syllable")) {
    console.log("جاري إضافة قسم تدريبات تحليل المقاطع الصوتية...");

    // Title page for syllable exercises
    doc.addSection({
      properties: pageProperties(MARGIN_TOP, MARGIN_BOTTOM),
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 400 },
          direction: "RTL",
          children: [
            new TextRun({
              text: content.texts.syllableSectionTitle,
              font: "Traditional Arabic",
              size: 72,
              bold: true,
              color: "2E86AB",
            }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 400 },
          direction: "RTL",
          children: [
            new TextRun({
              text: content.texts.syllableSectionSubtitle,
              font: "Traditional Arabic",
              size: 48,
              bold: true,
              color: "1B4965",
            }),
          ],
        }),
      ],
    });

    syllablePages = 1;

    for (let i = 0; i < syllableLetters.length; i += exercisesPerPage) {
      const pageContent = [];

      // Add exercises for this page
      const end = Math.min(i + exercisesPerPage, syllableLetters.length);
      for (let j = i; j < end; j++) {
        const letterData = syllableLetters[j];
        console.log(`جاري إضافة تمرين حرف ${letterData.name}...`);

        const exerciseContent = createSyllableExercise(letterData);
        pageContent.push(...exerciseContent);

        // Add spacing between exercises
        if (j < end - 1) {
          pageContent.push(
            new Paragraph({
              spacing: { after: 200 },
              direction: "RTL",
              children: [],
            })
          );
        }
      }

      // Add section for this page of exercises
      doc.addSection({
        properties: pageProperties(MARGIN_TOP, MARGIN_BOTTOM),
        children: pageContent,
      });

      syllablePages += 1;
    }
  }

  console.log("جاري حفظ الملف...");

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(OUTPUT, buffer);

  console.log("تم إنشاء الملف بنجاح!");
  console.log(`الملف: ${path.basename(OUTPUT)}`);
  if (coverContent) {
    console.log("صفحة الغلاف: 1 صفحة");
  }
  if (content.sections.includes("letters")) {
    console.log(`عدد صفحات الحروف: ${content.letters.length * 2} صفحة`);
  }
  if (content.sections.includes("syllable")) {
    console.log(`عدد صفحات تدريبات المقاطع الصوتية: ${syllablePages} صفحة`);
  }
  const total =
    (coverContent ? 1 : 0) +
    (content.sections.includes("letters") ? content.letters.length * 2 : 0) +
    syllablePages;
  console.log(`الإجمالي: ${total} صفحة`);
  console.log("حجم الصفحة: A5");
}

if (require.main === module) {
  createWorkbook().catch((err) => {
    console.error(`[FAIL] ${err.message}`);
    process.exitCode = 1;
  });
}
