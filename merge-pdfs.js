/*
    merge-pdfs.js
    Merges the seven workbook PDFs into one file next to this script.
    Uses pdf-lib (pure JS, no native deps).

    Every book must have a freshly converted PDF. There is no fallback to an
    archived copy: a missing or unreadable book fails the merge, so the merged
    file can never mix new pages with stale ones.
*/

const fs = require("fs");
const path = require("path");
const { PDFDocument } = require("pdf-lib");

const ROOT = __dirname;
const OUTPUT = path.join(ROOT, "Arabic_Workbooks_Complete.pdf");

// Merge order is the reading order of the complete book.
const BOOKS = [
  { name: "Arabic Alphabet Workbook", dir: "Arabic_Alphabet_Workbook" },
  { name: "Arabic Harakat Workbook", dir: "Arabic_Harakat_Workbook" },
  { name: "Arabic Syllable Analysis Workbook", dir: "Arabic_Syllable_Analysis_Workbook" },
  { name: "Arabic Word Analysis Workbook", dir: "Arabic_Word_Analysis_Workbook" },
  { name: "Arabic Letter Shapes Workbook", dir: "Arabic_Letter_Shapes_Workbook" },
  { name: "Arabic Word Composition Workbook", dir: "Arabic_Word_Composition_Workbook" },
  { name: "Arabic Stories Workbook", dir: "Arabic_Stories_Workbook" },
];

const OUT_TITLE = "Arabic Workbooks - Complete Book";
const OUT_AR_TITLE = "كتاب تمارين العربية الكامل";

function sourceFor(book) {
  return path.join(ROOT, book.dir, path.basename(book.dir) + ".pdf");
}

(async () => {
  const merged = await PDFDocument.create();
  merged.setTitle(OUT_TITLE);
  merged.setSubject(OUT_AR_TITLE);
  merged.setCreator("merge-pdfs.js (pdf-lib)");
  merged.setProducer("pdf-lib");
  const now = new Date();
  merged.setCreationDate(now);
  merged.setModificationDate(now);

  // Collect every source first: nothing is written unless all of them are there.
  const sources = [];
  const missing = [];

  for (const book of BOOKS) {
    const src = sourceFor(book);
    if (!fs.existsSync(src)) {
      missing.push(`${book.name} (${path.relative(ROOT, src)})`);
      continue;
    }
    sources.push({ ...book, src });
  }

  if (missing.length > 0) {
    throw new Error(
      `Missing PDF for ${missing.length} of ${BOOKS.length} books:\n` +
        missing.map((m) => `  - ${m}`).join("\n") +
        "\nRun build-all.bat (or docx-to-pdf.ps1) to convert them first."
    );
  }

  const used = [];

  for (const book of sources) {
    const bytes = fs.readFileSync(book.src);
    let srcDoc;
    try {
      srcDoc = await PDFDocument.load(bytes, { ignoreEncryption: true });
    } catch (err) {
      throw new Error(`Cannot read ${path.relative(ROOT, book.src)}: ${err.message}`);
    }

    const indices = srcDoc.getPageIndices();
    const pages = await merged.copyPages(srcDoc, indices);
    pages.forEach((page) => merged.addPage(page));

    const rel = path.relative(ROOT, book.src).split(path.sep).join("/");
    used.push({ name: book.name, file: rel, pages: pages.length });
    console.log(
      `  [OK]   ${book.name.padEnd(30)} ${String(pages.length).padStart(3)} pages  <- ${rel}`
    );
  }

  merged.setProducer(`pdf-lib | ${used.map((u) => u.name).join(" + ")}`);

  const outBytes = await merged.save();
  fs.writeFileSync(OUTPUT, outBytes);

  console.log("");
  console.log(
    `  Merged ${used.length} book(s), ${merged.getPageCount()} pages -> ${path.basename(OUTPUT)}`
  );
  console.log(
    `  Size: ${(fs.statSync(OUTPUT).size / 1024 / 1024).toFixed(2)} MB`
  );
  console.log(`  Path: ${OUTPUT}`);
})().catch((err) => {
  console.error(`  [FAIL] ${err.message}`);
  process.exit(1);
});
