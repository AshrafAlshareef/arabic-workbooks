/*
    merge-pdfs.js
    Merges the workbook PDFs into one file.

    Uses pdf-lib (pure JS, no native deps).

    Usage:
        node scripts/merge-pdfs.js [--books=folder1,folder2] [--root=<dir>] [--out=<file>]

        --books  project folders to merge (default: all seven)
        --root   folder holding <Book>\<Book>.pdf (default: project root)
        --out    output PDF (default: Arabic_Workbooks_Complete.pdf in the project root)

    Every selected book must have a freshly converted PDF. There is no fallback
    to an archived copy: a missing or unreadable book fails the merge, so the
    merged file can never mix new pages with stale ones.
*/

const fs = require("fs");
const path = require("path");
const { PDFDocument } = require("pdf-lib");

// This file lives in scripts/, so the project root is its parent: that is where
// books.json, the book folders and the default merged PDF live.
const SCRIPT_DIR = __dirname;
const ROOT = path.resolve(SCRIPT_DIR, "..");
const DEFAULT_OUTPUT = "Arabic_Workbooks_Complete.pdf";

// Merge order is the reading order of the complete book.
const ALL_BOOKS = JSON.parse(
  fs.readFileSync(path.join(ROOT, "books.json"), "utf-8")
).books;

const OUT_TITLE = "Arabic Workbooks - Complete Book";
const OUT_AR_TITLE = "كتاب تمارين العربية الكامل";

const USAGE = [
  "Usage: node scripts/merge-pdfs.js [options]",
  "",
  "  --books=<a,b>  project folders to merge (default: all seven)",
  "  --root=<dir>   folder holding <Book>\\<Book>.pdf (default: project root)",
  "  --out=<file>   output PDF (default: Arabic_Workbooks_Complete.pdf)",
].join("\n");

/** Parses `--flag=value` and `--flag value`; rejects anything unknown. */
function parseArgs(argv) {
  const options = { books: [], root: ROOT, out: path.join(ROOT, DEFAULT_OUTPUT) };

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
      missing.push(`${book.name} (${path.relative(ROOT, src)})`);
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

  // Renumber footer page numbers consecutively across the entire merged document.
  // Covers remain unnumbered (no footer page number), just like in individual books.
  try {
    const { PDFName, PDFArray, PDFRawStream, PDFString } = require("pdf-lib");
    const zlib = require("zlib");

    const pageCount = merged.getPageCount();
    let pageNum = 0;
    for (let i = 0; i < pageCount; i += 1) {
      const page = merged.getPage(i);
      const contentsRef = page.node.get(PDFName.of("Contents"));
      if (!contentsRef) continue;
      const refs = contentsRef instanceof PDFArray ? contentsRef.asArray() : [contentsRef];
      if (!refs.length) continue;

      // Try to find and replace footer page numbers (digits near bottom of page)
      // We'll look through content streams for page numbers in footer area
      for (let r = 0; r < refs.length; r += 1) {
        const ref = refs[r];
        const stream = merged.context.lookup(ref);
        if (!(stream instanceof PDFRawStream)) continue;

        let raw = Buffer.from(stream.contents);
        let isCompressed = raw.length > 0 && raw[0] === 0x78;
        let content;
        if (isCompressed) {
          try {
            content = zlib.inflateSync(raw).toString("latin1");
          } catch (e) {
            continue;
          }
        } else {
          content = raw.toString("latin1");
        }

        // Check if this page has a footer page number (digit in bottom area)
        // Parse to find y positions
        const lines = content.split(/\r?\n/);
        let tm = [1, 0, 0, 1, 0, 0];
        let tlm = tm.slice();
        const ops = [];
        for (const line of lines) {
          const t = line;
          let m = t.match(/([-.\d]+) ([-.\d]+) ([-.\d]+) ([-.\d]+) ([-.\d]+) ([-.\d]+)\s+Tm\b/);
          if (m) {
            tm = [parseFloat(m[1]), parseFloat(m[2]), parseFloat(m[3]), parseFloat(m[4]), parseFloat(m[5]), parseFloat(m[6])];
            tlm = tm.slice();
            continue;
          }
          m = t.match(/([-.\d]+) ([-.\d]+)\s+Td\b/);
          if (m) {
            const tx = parseFloat(m[1]), ty = parseFloat(m[2]);
            tlm = [tlm[0], tlm[1], tlm[2], tlm[3], tlm[4] + tx * tlm[0] + ty * tlm[2], tlm[5] + tx * tlm[1] + ty * tlm[3]];
            tm = tlm.slice();
            continue;
          }
          if (/\bT\*\b/.test(t)) {
            tm = tlm.slice();
            continue;
          }
          m = t.match(/\(([^()]*)\)\s*Tj\b/);
          if (m) {
            ops.push({ y: tm[5], x: tm[4], type: "Tj", text: m[1], match: m[0], fullLine: t });
            continue;
          }
          m = t.match(/\[([^\]]*)\]\s*TJ\b/);
          if (m) {
            const inner = m[1];
            const texts = [];
            const sm = [...inner.matchAll(/\(([^()]*)\)/g)];
            for (const sx of sm) texts.push(sx[1]);
            const fullText = texts.join("");
            ops.push({ y: tm[5], x: tm[4], type: "TJ", text: fullText, match: m[0], fullLine: t, inner });
          }
        }

        // Find footer page numbers (y < 50 and contains only digits/spaces)
        const candidates = ops.filter((o) => o.y < 50 && /^\s*\d+\s*$/.test(o.text));

        if (candidates.length > 0) {
          pageNum += 1;
          const newNum = String(pageNum);
          let newContent = content;

          // Replace each candidate
          const sorted = candidates.slice().sort((a, b) => b.x - a.x);
          for (const cand of sorted) {
            if (cand.type === "Tj") {
              // Replace just the (number) Tj part
              const oldOp = cand.match;
              // Use function replace to avoid escaping issues
              newContent = newContent.split(oldOp).join(`(${newNum}) Tj`);
            } else if (cand.type === "TJ") {
              const oldOp = cand.match;
              newContent = newContent.split(oldOp).join(`[(${newNum})] TJ`);
            }
          }

          if (newContent !== content) {
            let newRaw;
            if (isCompressed) {
              const compressed = zlib.deflateSync(Buffer.from(newContent, "latin1"));
              newRaw = compressed;
            } else {
              newRaw = Buffer.from(newContent, "latin1");
            }
            stream.contents = newRaw;
          }
        }
        // If no footer numbers found, it's likely a cover - leave it unnumbered
      }
    }
  } catch (e) {
    console.log(`  [WARN] Could not renumber pages in merged PDF: ${e.message}`);
  }

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
