/*
    build-options.js
    The command-line options shared by the seven create-workbook.js entry
    points:

        node create-workbook.js [--order=abjadi|alphabetical] [--out=<dir>]

    Defaults reproduce the original behaviour exactly: Abjadi order, DOCX
    written inside the project folder.
*/

const path = require("path");
const { parseOrder, ORDER_NAMES } = require("./letter-order.js");

const USAGE = [
  "Usage: node create-workbook.js [options]",
  "",
  "  --order=<name>   letter order: " + ORDER_NAMES.join(" | ") + " (default: abjadi)",
  "  --out=<dir>      write the .docx into <dir> instead of the project folder",
  "  --help           show this message",
].join("\n");

/**
 * Parses `--flag=value` and `--flag value` pairs. Unknown arguments are
 * rejected instead of ignored, so a typo cannot silently build the wrong book.
 */
function parseBookArgs(argv) {
  const options = { order: "abjadi", out: null, help: false };

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

    if (flag === "--order") {
      options.order = parseOrder(takeValue());
    } else if (flag === "--out") {
      const value = takeValue();
      if (value.trim() === "") throw new Error("--out needs a directory.");
      options.out = path.resolve(value);
    } else if (flag === "--help" || flag === "-h") {
      options.help = true;
    } else {
      throw new Error(`Unknown argument "${arg}". Try --help.`);
    }
  }

  return options;
}

module.exports = { USAGE, parseBookArgs };
