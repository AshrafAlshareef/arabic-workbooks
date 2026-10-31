/*
    merge-pdfs.js
    Merges the workbook PDFs into one file.

    Uses pdf-lib (pure JS, no native deps).

    Usage:
        node merge-pdfs.js [--books=folder1,folder2] [--root=<dir>] [--out=<file>]

        --books  project folders to merge (default: all seven)
        --root   folder holding <Book>\<Book>.pdf (default: this folder)
        --out    output PDF (default: Arabic_Workbooks_Complete.pdf next to this script)

    Every selected book must have a freshly converted PDF. There is no fallback
    to an archived copy: a missing or unreadable book fails the merge, so the
    merged file can never mix new pages with stale ones.
*/

const fs = require("fs");
const path = require("path");
const { PDFDocument } = require("pdf-lib");

const SCRIPT_ROOT = __dirname;
const DEFAULT_OUTPUT = "Arabic_Workbooks_Complete.pdf";

// Merge order is the reading order of the complete book.
const ALL_BOOKS = JSON.parse(
  fs.readFileSync(path.join(SCRIPT_ROOT, "books.json"), "utf-8")
).books;

const OUT_TITLE = "Arabic Workbooks - Complete Book";
const OUT_AR_TITLE = "كتاب تمارين العربية الكامل";

const USAGE = [
  "Usage: node merge-pdfs.js [options]",
  "",
  "  --books=<a,b>  project folders to merge (default: all seven)",
  "  --root=<dir>   folder holding <Book>\\<Book>.pdf (default: this folder)",
  "  --out=<file>   output PDF (default: Arabic_Workbooks_Complete.pdf)",
].join("\n");

/** Parses `--flag=value` and `--flag value`; rejects anything unknown. */
function parseArgs(argv) {
  const options = { books: [], root: SCRIPT_ROOT, out: path.join(SCRIPT_ROOT, DEFAULT_OUTPUT) };

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

    if (flag === "--books") {
      options.books = takeValue()
        .split(",")
        .map((name) => name.trim())
        .filter((name) => name.length > 0);
    } else if (flag === "--root") {
      options.root = path.resolve(takeValue());
    } else if (flag === "--out") {
      options.out = path.resolve(takeValue());
    } else if (flag === "--help" || flag === "-h") {
      console.log(USAGE);
      process.exit(0);
    } else {
      throw new Error(`Unknown argument "${arg}". Try --help.`);
    }
  }

  return options;
}

function sourceFor(root, book) {
  return path.join(root, book.dir, path.basename(book.dir) + ".pdf");
}

(async () => {
  const options = parseArgs(process.argv.slice(2));

  let books = ALL_BOOKS;
  if (options.books.length > 0) {
    const known = ALL_BOOKS.map((book) => book.dir);
    const unknown = options.books.filter((name) => !known.includes(name));
    if (unknown.length > 0) {
      throw new Error(
        `Unknown book(s): ${unknown.join(", ")}\nKnown books: ${known.join(", ")}`
      );
    }
    // Keep the canonical reading order, whatever order the caller listed.
    books = ALL_BOOKS.filter((book) => options.books.includes(book.dir));
  }

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

  for (const book of books) {
    const src = sourceFor(options.root, book);
    if (!fs.existsSync(src)) {
      missing.push(`${book.name} (${path.relative(SCRIPT_ROOT, src)})`);
      continue;
    }
    sources.push({ ...book, src });
  }

  if (missing.length > 0) {
    throw new Error(
      `Missing PDF for ${missing.length} of ${books.length} books:\n` +
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
      throw new Error(`Cannot read ${path.relative(SCRIPT_ROOT, book.src)}: ${err.message}`);
    }

    const indices = srcDoc.getPageIndices();
    const pages = await merged.copyPages(srcDoc, indices);
    pages.forEach((page) => merged.addPage(page));

    const rel = path.relative(SCRIPT_ROOT, book.src).split(path.sep).join("/");
    used.push({ name: book.name, file: rel, pages: pages.length });
    console.log(
      `  [OK]   ${book.name.padEnd(30)} ${String(pages.length).padStart(3)} pages  <- ${rel}`
    );
  }

  merged.setProducer(`pdf-lib | ${used.map((u) => u.name).join(" + ")}`);

  const outBytes = await merged.save();
  fs.mkdirSync(path.dirname(options.out), { recursive: true });
  fs.writeFileSync(options.out, outBytes);

  console.log("");
  console.log(
    `  Merged ${used.length} book(s), ${merged.getPageCount()} pages -> ${path.basename(options.out)}`
  );
  console.log(
    `  Size: ${(fs.statSync(options.out).size / 1024 / 1024).toFixed(2)} MB`
  );
  console.log(`  Path: ${options.out}`);
})().catch((err) => {
  console.error(`  [FAIL] ${err.message}`);
  process.exit(1);
});
