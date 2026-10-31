#!/usr/bin/env node
/*
    build.js
    Runs the whole pipeline for a selection of workbooks:

        1. node create-workbook.js   (per selected book)
        2. docx-to-pdf.ps1           (Word COM, selected books only)
        3. merge-pdfs.js             (optional)

    Usage:
        node build.js [--books=folder1,folder2] [--order=abjadi|alphabetical]
                      [--out=<dir>] [--merge] [--merged-name=<file.pdf>]
                      [--no-pdf]

        --books        project folders to build (default: all seven)
        --order        letter order, abjadi (default) or alphabetical
        --out          output folder (default: this folder). Each book gets
                       <out>\<Book>\<Book>.docx|pdf; the merged file goes in <out>
        --merge        merge the selected books into one PDF
        --merged-name  merged file name (default: Arabic_Workbooks_Complete.pdf)
        --no-pdf       DOCX only: skip the Word conversion and the merge

    Progress markers (consumed by the desktop GUI):
        ##STAGE <n>/<total> <name>
        ##BOOK <folder> building|ok|failed
        ##OUTPUT <path>
        ##DONE ok|failed
*/

const fs = require("fs");
const path = require("path");
const { spawnSync } = require("child_process");
const { parseOrder } = require("./letter-order.js");

const ROOT = __dirname;
const BOOKS = JSON.parse(fs.readFileSync(path.join(ROOT, "books.json"), "utf-8")).books;
const DEFAULT_MERGED = "Arabic_Workbooks_Complete.pdf";

const USAGE = [
  "Usage: node build.js [options]",
  "",
  "  --books=<a,b>        project folders to build (default: all seven)",
  "  --order=<name>       abjadi | alphabetical (default: abjadi)",
  "  --out=<dir>          output folder (default: this folder)",
  "  --merge              merge the selected books into one PDF",
  "  --merged-name=<f>    merged file name (default: " + DEFAULT_MERGED + ")",
  "  --no-pdf             DOCX only, no Word conversion and no merge",
  "  --help               show this message",
].join("\n");

function emit(line) {
  console.log(line);
}

function parseArgs(argv) {
  const options = {
    books: [],
    order: "abjadi",
    out: ROOT,
    merge: false,
    mergedName: DEFAULT_MERGED,
    pdf: true,
  };

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
    } else if (flag === "--order") {
      options.order = parseOrder(takeValue());
    } else if (flag === "--out") {
      options.out = path.resolve(takeValue());
    } else if (flag === "--merge") {
      options.merge = true;
    } else if (flag === "--merged-name") {
      const name = path.basename(takeValue().trim());
      if (!name || name === "." || name === "..") {
        throw new Error("--merged-name needs a file name.");
      }
      options.mergedName = name.toLowerCase().endsWith(".pdf") ? name : `${name}.pdf`;
    } else if (flag === "--no-pdf") {
      options.pdf = false;
    } else if (flag === "--help" || flag === "-h") {
      console.log(USAGE);
      process.exit(0);
    } else {
      throw new Error(`Unknown argument "${arg}". Try --help.`);
    }
  }

  return options;
}

function runNode(script, args, cwd) {
  const result = spawnSync(process.execPath, [script, ...args], { cwd, stdio: "inherit" });
  if (result.error) throw new Error(`cannot run ${script}: ${result.error.message}`);
  return result.status;
}

function runPowershell(scriptArgs) {
  const result = spawnSync(
    "powershell.exe",
    ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ...scriptArgs],
    { cwd: ROOT, stdio: "inherit" }
  );
  if (result.error) throw new Error(`cannot run powershell.exe: ${result.error.message}`);
  return result.status;
}

/** The seven projects share the root node_modules; install only what is missing. */
function ensureDependencies() {
  const required = ["docx", "pdf-lib", "svg2img"];
  const missing = required.filter((name) => !fs.existsSync(path.join(ROOT, "node_modules", name)));
  if (missing.length === 0) return;

  emit(`##INFO installing shared dependencies: ${missing.join(", ")}`);
  const npm = process.platform === "win32" ? "npm.cmd" : "npm";
  const result = spawnSync(npm, ["install", "--no-audit", "--no-fund"], {
    cwd: ROOT,
    stdio: "inherit",
    shell: true,
  });
  if (result.error) throw new Error(`cannot run npm: ${result.error.message}`);
  if (result.status !== 0) {
    throw new Error("npm install failed. Check your network connection and Node version.");
  }
}

function main() {
  const options = parseArgs(process.argv.slice(2));

  let selected = BOOKS;
  if (options.books.length > 0) {
    const known = BOOKS.map((book) => book.dir);
    const unknown = options.books.filter((name) => !known.includes(name));
    if (unknown.length > 0) {
      throw new Error(`Unknown book(s): ${unknown.join(", ")}\nKnown books: ${known.join(", ")}`);
    }
    selected = BOOKS.filter((book) => options.books.includes(book.dir));
  }

  if (selected.length === 0) throw new Error("no books selected.");
  if (options.merge && !options.pdf) {
    throw new Error("--merge needs PDFs, so it cannot be combined with --no-pdf.");
  }

  const outRoot = path.resolve(options.out);
  const stages = options.pdf ? 3 : 1;

  emit(
    `##CONFIG order=${options.order} out=${outRoot} books=${selected.length} merge=${options.merge}`
  );

  ensureDependencies();

  // -- 1. DOCX --------------------------------------------------------------
  emit(`##STAGE 1/${stages} build`);
  for (const book of selected) {
    const bookOut = path.join(outRoot, book.dir);
    emit(`##BOOK ${book.dir} building`);
    const status = runNode(
      "create-workbook.js",
      [`--order=${options.order}`, `--out=${bookOut}`],
      path.join(ROOT, book.dir)
    );
    if (status !== 0) {
      emit(`##BOOK ${book.dir} failed`);
      throw new Error(`Build failed in ${book.dir} (exit code ${status}).`);
    }
    emit(`##BOOK ${book.dir} ok`);
    emit(`##OUTPUT ${path.join(bookOut, `${book.dir}.docx`)}`);
  }

  // -- 2. PDF ---------------------------------------------------------------
  if (options.pdf) {
    emit(`##STAGE 2/${stages} convert`);
    const status = runPowershell([
      path.join(ROOT, "docx-to-pdf.ps1"),
      "-OutputRoot",
      outRoot,
      "-Books",
      selected.map((book) => book.dir).join(","),
    ]);
    if (status !== 0) {
      throw new Error(`PDF conversion failed (exit code ${status}). Close Microsoft Word and try again.`);
    }
    for (const book of selected) {
      emit(`##OUTPUT ${path.join(outRoot, book.dir, `${book.dir}.pdf`)}`);
    }
  }

  // -- 3. Merge -------------------------------------------------------------
  if (options.merge) {
    emit(`##STAGE 3/${stages} merge`);
    const mergedPath = path.join(outRoot, options.mergedName);
    const status = runNode(
      "merge-pdfs.js",
      [
        `--books=${selected.map((book) => book.dir).join(",")}`,
        `--root=${outRoot}`,
        `--out=${mergedPath}`,
      ],
      ROOT
    );
    if (status !== 0) throw new Error("Merge failed.");
    emit(`##OUTPUT ${mergedPath}`);
  }

  emit("##DONE ok");
}

try {
  main();
} catch (err) {
  console.error(`[FAIL] ${err.message}`);
  emit("##DONE failed");
  process.exitCode = 1;
}
