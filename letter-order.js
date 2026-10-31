/*
    letter-order.js
    The two Arabic letter orders the workbooks can be built in, plus the helper
    that re-orders a book's data in memory.

    The committed JSON files stay in Abjadi order. A builder receives --order
    and sorts its own copy, so nothing here (and nothing in the builders)
    writes to the data files.
*/

// Traditional Abjadi order (abjad hawwaz ...): the order the data is authored in.
const ABJADI = [
  "أ", "ب", "ج", "د", "ه", "و", "ز", "ح", "ط", "ي",
  "ك", "ل", "م", "ن", "س", "ع", "ف", "ص", "ق", "ر",
  "ش", "ت", "ث", "خ", "ذ", "ض", "ظ", "غ",
];

// Modern alphabetical / dictionary order (alef, ba, ta, tha, jim, ...).
const ALPHABETICAL = [
  "أ", "ب", "ت", "ث", "ج", "ح", "خ", "د", "ذ", "ر",
  "ز", "س", "ش", "ص", "ض", "ط", "ظ", "ع", "غ", "ف",
  "ق", "ك", "ل", "م", "ن", "ه", "و", "ي",
];

const ORDERS = { abjadi: ABJADI, alphabetical: ALPHABETICAL };
const ORDER_NAMES = Object.keys(ORDERS);

const TATWEEL = /\u0640/g;
// Harakat, tanween, shadda, sukun, superscript alef and Quranic marks.
const HARAKAT = /[\u0610-\u061A\u064B-\u065F\u0670\u06D6-\u06ED]/g;
// Every alef shape looks the same to the order tables.
const ALEF_FORMS = /[\u0622\u0623\u0625\u0671]/g;
const BARE_ALEF = "\u0627";

/** Aliases accepted on the command line, in addition to the canonical names. */
const ALIASES = {
  abjad: "abjadi",
  "أبجدي": "abjadi",
  alphabet: "alphabetical",
  modern: "alphabetical",
  "أبجدي حديث": "alphabetical",
  "معجمي": "alphabetical",
};

/**
 * The base letter an entry sorts by: marks and tatweel are removed, alef
 * shapes are unified, and only the first letter is kept. `"هـ"` -> `"ه"`,
 * `"أَسَد"` -> `"ا"`.
 */
function normalizeLetter(value) {
  if (typeof value !== "string") return "";
  const cleaned = value.replace(HARAKAT, "").replace(TATWEEL, "").trim();
  if (cleaned.length === 0) return "";
  return cleaned[0].replace(ALEF_FORMS, BARE_ALEF);
}

/** Canonical order name, or a clear error. An empty value means the default. */
function parseOrder(value) {
  if (value === undefined || value === null || value === "") return "abjadi";
  const name = String(value).trim().toLowerCase();
  const resolved = ALIASES[name] || (ORDER_NAMES.includes(name) ? name : null);
  if (!resolved) {
    throw new Error(
      `Unknown letter order "${value}". Use one of: ${ORDER_NAMES.join(", ")}.`
    );
  }
  return resolved;
}

const indexCache = new Map();

/** Map of normalized letter -> position, built once per order. */
function orderIndex(orderName) {
  const name = parseOrder(orderName);
  if (!indexCache.has(name)) {
    const index = new Map();
    ORDERS[name].forEach((letter, position) => {
      const key = normalizeLetter(letter);
      if (index.has(key)) {
        throw new Error(`Order "${name}" lists "${letter}" more than once.`);
      }
      index.set(key, position);
    });
    indexCache.set(name, index);
  }
  return indexCache.get(name);
}

/**
 * Stable copy of `items` in `orderName`, sorted by the base letter of
 * `options.key(item)` (default: `item.letter`).
 *
 * Throws before any page is built when an entry has no letter, or its letter
 * is not in the chosen order table - a missing letter would silently land at
 * the end of the book otherwise.
 */
function sortByLetter(items, orderName, options = {}) {
  const name = parseOrder(orderName);
  const index = orderIndex(name);
  const keyOf = options.key || ((item) => item.letter);
  const label = options.label || "items";

  const decorated = items.map((item, position) => {
    const raw = keyOf(item);
    const letter = normalizeLetter(raw);
    if (!letter) {
      throw new Error(`${label}: entry ${position} has no letter to sort by.`);
    }
    if (!index.has(letter)) {
      throw new Error(
        `${label}: "${raw}" starts with "${letter}", which is not in the ${name} order table.`
      );
    }
    return { item, position, rank: index.get(letter) };
  });

  // Explicit position tie-breaker keeps the sort stable on every Node version.
  decorated.sort((a, b) => a.rank - b.rank || a.position - b.position);
  return decorated.map((entry) => entry.item);
}

module.exports = {
  ORDERS,
  ORDER_NAMES,
  normalizeLetter,
  parseOrder,
  sortByLetter,
};
