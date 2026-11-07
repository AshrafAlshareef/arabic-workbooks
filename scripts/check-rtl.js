#!/usr/bin/env node
/*
    check-rtl.js
    Verifies that generated .docx files are Arabic right-to-left documents.

    `direction: "RTL"` is not a real OOXML/docx property, so it is silently
    ignored. RTL must be requested with:
        paragraph: `bidirectional: true` -> <w:bidi/>
        text run:  `rightToLeft: true`    -> <w:rtl/>

    Every generator in this project sets those once in the document defaults,
    so every paragraph and run inherits them - including pages added later.
    This checker fails the build if that inheritance is missing or disabled.

    Usage:
        node scripts/check-rtl.js                  # all books (books.json)
        node scripts/check-rtl.js <file.docx>...   # explicit files
        node scripts/check-rtl.js --out=<dir>      # books built into <dir>

    Exit code 0 = every file is RTL, 1 = at least one file is not.
*/

const fs = require("fs");
const path = require("path");
const JSZip = require("jszip");

const ROOT = path.resolve(__dirname, "..");
const BOOKS = JSON.parse(fs.readFileSync(path.join(ROOT, "books.json"), "utf-8")).books;
const EXTRA_BOOKS = ["image_2_cover_style/images-to-book.docx"];

const DISABLED = /<w:(?:rtl|bidi)\b[^>]*\bw:val="(?:0|false|off)"[^>]*\/?>/gi;
const RTL_ELEMENT = /<w:rtl\b[^>]*\/>/i;
const BIDI_ELEMENT = /<w:bidi\b[^>]*\/>/i;
const ARABIC_TEXT = /[\u0600-\u06FF]/;

function bookDocxPaths(outRoot) {
  return [
    ...BOOKS.map((book) => path.join(outRoot, book.dir, `${book.dir}.docx`)),
    ...EXTRA_BOOKS.map((rel) => path.join(outRoot, rel)),
  ];
}

function parseArgs(argv) {
  const options = { files: [], out: ROOT, explicit: false };
  for (const arg of argv) {
    if (arg === "--help" || arg === "-h") {
      console.log("Usage: node scripts/check-rtl.js [--out=<dir>] [file.docx ...]");
      process.exit(0);
    } else if (arg.startsWith("--out=")) {
      options.out = path.resolve(arg.slice("--out=".length));
    } else if (arg.startsWith("--")) {
      throw new Error(`Unknown argument "${arg}". Try --help.`);
    } else {
      options.files.push(path.resolve(arg));
      options.explicit = true;
    }
  }
  if (!options.explicit) options.files = bookDocxPaths(options.out);
  return options;
}

async function checkFile(file) {
  const problems = [];
  const notes = [];

  if (!fs.existsSync(file)) {
    return { file, problems: ["file does not exist"], notes };
  }

  let zip;
  try {
    zip = await JSZip.loadAsync(fs.readFileSync(file));
  } catch (err) {
    return { file, problems: [`not a readable .docx (${err.message})`], notes };
  }

  const stylesEntry = zip.file("word/styles.xml");
  const documentEntry = zip.file("word/document.xml");
  if (!stylesEntry) {
    return { file, problems: ["missing word/styles.xml"], notes };
  }

  const styles = await stylesEntry.async("string");
  const documentXml = documentEntry ? await documentEntry.async("string") : "";

  const rPrDefault = /<w:rPrDefault>([\s\S]*?)<\/w:rPrDefault>/i.exec(styles);
  const pPrDefault = /<w:pPrDefault>([\s\S]*?)<\/w:pPrDefault>/i.exec(styles);

  if (!rPrDefault) {
    problems.push("styles.xml has no <w:rPrDefault> block");
  } else if (!RTL_ELEMENT.test(rPrDefault[1])) {
    problems.push("document run defaults are missing <w:rtl/> (runs are not RTL)");
  }

  if (!pPrDefault) {
    problems.push("styles.xml has no <w:pPrDefault> block");
  } else if (!BIDI_ELEMENT.test(pPrDefault[1])) {
    problems.push("document paragraph defaults are missing <w:bidi/> (paragraphs are not RTL)");
  }

  for (const [name, xml] of [
    ["word/styles.xml", styles],
    ["word/document.xml", documentXml],
  ]) {
    const disabled = xml.match(DISABLED);
    if (disabled) {
      const unique = [...new Set(disabled.map((tag) => tag.replace(/\s+/g, " ")))];
      problems.push(`${name} explicitly disables RTL: ${unique.join(", ")}`);
    }
  }

  if (documentXml && !ARABIC_TEXT.test(documentXml)) {
    notes.push("no Arabic characters in body (image-only book?)");
  }

  return { file, problems, notes };
}

async function main() {
  const options = parseArgs(process.argv.slice(2));
  if (!options.explicit) {
    // Image-tool output is optional: it is gitignored and may not be built yet.
    options.files = options.files.filter((file) => fs.existsSync(file));
  }
  if (options.files.length === 0) {
    console.error("check-rtl: no .docx files found to check.");
    process.exit(1);
  }

  const results = [];
  for (const file of options.files) {
    results.push(await checkFile(file));
  }

  let failed = 0;
  for (const result of results) {
    const rel = path.relative(ROOT, result.file) || result.file;
    if (result.problems.length === 0) {
      console.log(`OK   ${rel}`);
    } else {
      failed += 1;
      console.log(`FAIL ${rel}`);
      for (const problem of result.problems) console.log(`     - ${problem}`);
    }
    for (const note of result.notes) console.log(`     note: ${note}`);
  }

  console.log("");
  console.log(`${results.length - failed}/${results.length} file(s) are Arabic RTL.`);
  process.exit(failed === 0 ? 0 : 1);
}

main().catch((err) => {
  console.error(`check-rtl: ${err.message}`);
  process.exit(1);
});
