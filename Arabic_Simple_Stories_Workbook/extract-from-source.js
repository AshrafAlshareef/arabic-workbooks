#!/usr/bin/env node
/**
 * extract-from-source.js
 *
 * Re-runnable importer that turns the simple-stories Word file into the
 * project's data assets:
 *
 *   assets/texts/NN-slug.txt    line 1 = title, remaining lines = 3 sentences
 *   assets/images/NN-slug.svg   the matching picture from Arabic_Alphabet_Workbook
 *   assets/stories.json         ordered index used by create-workbook.js
 *   assets/coverpage.png        copied from Arabic_Stories_Workbook when missing
 *
 * The Word file is big (it embeds the pictures), so it is NOT copied into this
 * project. Extract it from `../temp/stories-simple.docx` by default; pass
 * `--source=<path>` to read a different file.
 *
 * Usage: node extract-from-source.js [--source=<path.docx>]
 */

const fs = require("fs");
const path = require("path");
const JSZip = require("jszip");

const ROOT = __dirname;
const DEFAULT_SOURCE = path.join(ROOT, "..", "temp", "stories-simple.docx");
const ASSETS = path.join(ROOT, "assets");
const TEXTS = path.join(ASSETS, "texts");
const IMAGES = path.join(ASSETS, "images");
const COVER = path.join(ASSETS, "coverpage.png");
const MANIFEST = path.join(ASSETS, "stories.json");
const ALPHABET_IMAGES = path.join(ROOT, "..", "Arabic_Alphabet_Workbook", "assets", "images");
const STORIES_COVER = path.join(ROOT, "..", "Arabic_Stories_Workbook", "assets", "coverpage.png");

const BOOK = {
  title: "كِتَابُ الْقِصَصِ الْبَسِيطَةِ",
  subtitle: "قِصَصٌ سَهْلَةٌ بِجُمَلٍ قَصِيرَةٍ مُشَكَّلَةٍ",
  cover: "coverpage.png",
};

/**
 * The order is the Word file's order (أ ... غ). `letter` mirrors the sibling
 * stories workbook so `sortByLetter` behaves identically; `imageFile` names
 * the same picture in Arabic_Alphabet_Workbook/assets/images.
 */
const STORIES = [
  { slug: "asad", letter: "أ", word: "أسد", imageFile: "lion.svg" },
  { slug: "bab", letter: "ب", word: "باب", imageFile: "door.svg" },
  { slug: "jamal", letter: "ج", word: "جمل", imageFile: "camel.svg" },
  { slug: "dubb", letter: "د", word: "دب", imageFile: "bear.svg" },
  { slug: "hilal", letter: "هـ", word: "هلال", imageFile: "crescent.svg" },
  { slug: "warda", letter: "و", word: "وردة", imageFile: "flower.svg" },
  { slug: "zarafa", letter: "ز", word: "زرافة", imageFile: "giraffe.svg" },
  { slug: "hisan", letter: "ح", word: "حصان", imageFile: "horse.svg" },
  { slug: "taair", letter: "ط", word: "طائر", imageFile: "bird.svg" },
  { slug: "yad", letter: "ي", word: "يد", imageFile: "hand.svg" },
  { slug: "kitab", letter: "ك", word: "كتاب", imageFile: "book.svg" },
  { slug: "laymun", letter: "ل", word: "ليمون", imageFile: "lemon.svg" },
  { slug: "maa", letter: "م", word: "ماء", imageFile: "water.svg" },
  { slug: "najm", letter: "ن", word: "نجم", imageFile: "star.svg" },
  { slug: "samak", letter: "س", word: "سمك", imageFile: "fish.svg" },
  { slug: "ain", letter: "ع", word: "عين", imageFile: "eye.svg" },
  { slug: "fil", letter: "ف", word: "فيل", imageFile: "elephant.svg" },
  { slug: "sanduq", letter: "ص", word: "صندوق", imageFile: "box.svg" },
  { slug: "qalb", letter: "ق", word: "قلب", imageFile: "heart.svg" },
  { slug: "rajul", letter: "ر", word: "رجل", imageFile: "man.svg" },
  { slug: "shams", letter: "ش", word: "شمس", imageFile: "sun.svg" },
  { slug: "tuffah", letter: "ت", word: "تفاح", imageFile: "apple.svg" },
  { slug: "thaalab", letter: "ث", word: "ثعلب", imageFile: "fox.svg" },
  { slug: "khubz", letter: "خ", word: "خبز", imageFile: "bread.svg" },
  { slug: "dhiib", letter: "ذ", word: "ذئب", imageFile: "wolf.svg" },
  { slug: "dafdaa", letter: "ض", word: "ضفدع", imageFile: "frog.svg" },
  { slug: "zarf", letter: "ظ", word: "ظرف", imageFile: "envelope.svg" },
  { slug: "ghazal", letter: "غ", word: "غزال", imageFile: "gazelle.svg" },
];

const EXPECTED_SENTENCES = 3;
const HARAKAT_RE = /[\u0610-\u061A\u064B-\u065F\u0670\u06D6-\u06ED]/g;
const TATWEEL_RE = /\u0640/g;

// Two typos in the Word file, corrected once here so re-running the extractor
// keeps the corrected texts.
const TEXT_FIXES = [["مِنَ اشكال", "مِنْ أَشْكَالِ"]];
const TITLE_FIXES = { rajul: "الرَّجُلُ" };

function fixSentence(sentence) {
  return TEXT_FIXES.reduce(
    (line, [from, to]) => line.split(from).join(to),
    sentence
  );
}

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

function decodeEntities(value) {
  return value
    .replace(/&#x([0-9A-Fa-f]+);/g, (_, hex) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, dec) => String.fromCodePoint(parseInt(dec, 10)))
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

/** Only the visible text runs matter; drawings carry no story text. */
function paragraphText(pXml) {
  const text = [...pXml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)]
    .map((match) => match[1])
    .join("");
  return decodeEntities(text).trim();
}

/** "5. هـ — الْهِلَالُ" -> "الْهِلَالُ"; ordinary titles are already clean. */
function cleanTitle(raw) {
  let title = raw.trim();
  const dash = title.search(/[\u2014\u2013-]/);
  if (dash >= 0) title = title.slice(dash + 1);
  return title.replace(/^\d+\.\s*/, "").trim();
}

function normalizeWord(value) {
  return value
    .replace(HARAKAT_RE, "")
    .replace(TATWEEL_RE, "")
    .replace(/^ال/, "")
    .trim();
}

/** Split the flat document into heading-led stories, in document order. */
function extractStories(documentXml) {
  const stories = [];
  for (const pXml of documentXml.match(/<w:p\b[\s>][\s\S]*?<\/w:p>/g) || []) {
    const line = paragraphText(pXml);
    if (!line) continue;

    if (/[.۔]$/.test(line)) {
      assert(stories.length > 0, `Body text before any title: ${line}`);
      const story = stories[stories.length - 1];
      assert(
        story.sentences.length < EXPECTED_SENTENCES,
        `Story ${stories.length} has more than ${EXPECTED_SENTENCES} sentences`
      );
      story.sentences.push(fixSentence(line));
    } else {
      stories.push({ rawTitle: line, title: cleanTitle(line), sentences: [] });
    }
  }
  return stories;
}

function validate(stories) {
  assert(
    stories.length === STORIES.length,
    `Expected ${STORIES.length} stories, parsed ${stories.length}`
  );
  stories.forEach((story, index) => {
    const meta = STORIES[index];
    assert(
      story.sentences.length === EXPECTED_SENTENCES,
      `Story ${index + 1} (${meta.word}): expected ${EXPECTED_SENTENCES} sentences, got ${story.sentences.length}`
    );
    assert(
      normalizeWord(story.title) === meta.word,
      `Story ${index + 1}: title "${story.title}" does not match "${meta.word}"`
    );
  });
}

function resolveSource(argv) {
  const flag = argv.find((arg) => arg.startsWith("--source="));
  return flag ? path.resolve(flag.slice("--source=".length)) : DEFAULT_SOURCE;
}

/** Keep hand-edited text files when their sentences still match the source. */
function writeStoryText(textPath, title, sentences) {
  if (fs.existsSync(textPath)) {
    const existing = fs
      .readFileSync(textPath, "utf-8")
      .replace(/^\uFEFF/, "")
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter(Boolean);
    const same =
      existing.length === sentences.length + 1 &&
      existing[0] === title &&
      existing.slice(1).every((line, i) => line === sentences[i]);
    if (same) return false;
  }
  fs.writeFileSync(textPath, [title, ...sentences].join("\n") + "\n", "utf8");
  return true;
}

function ensureCover() {
  if (fs.existsSync(COVER)) return "keeping existing";
  assert(fs.existsSync(STORIES_COVER), `Missing cover to copy: ${STORIES_COVER}`);
  fs.copyFileSync(STORIES_COVER, COVER);
  return "copied from Arabic_Stories_Workbook";
}

async function main() {
  const source = resolveSource(process.argv.slice(2));
  assert(fs.existsSync(source), `Missing source file: ${source}`);

  const zip = await JSZip.loadAsync(fs.readFileSync(source));
  const documentXml = await zip.file("word/document.xml").async("string");
  const stories = extractStories(documentXml);
  validate(stories);

  for (const dir of [ASSETS, TEXTS, IMAGES]) fs.mkdirSync(dir, { recursive: true });

  const manifest = { book: BOOK, stories: [] };
  let copiedBytes = 0;

  for (let i = 0; i < stories.length; i++) {
    const meta = STORIES[i];
    const name = `${String(i + 1).padStart(2, "0")}-${meta.slug}`;
    const title = TITLE_FIXES[meta.slug] || stories[i].title;

    writeStoryText(path.join(TEXTS, `${name}.txt`), title, stories[i].sentences);

    const picture = path.join(ALPHABET_IMAGES, meta.imageFile);
    assert(fs.existsSync(picture), `Missing alphabet workbook picture: ${picture}`);
    const destination = path.join(IMAGES, `${name}.svg`);
    fs.copyFileSync(picture, destination);
    copiedBytes += fs.statSync(destination).size;

    manifest.stories.push({
      id: i + 1,
      letter: meta.letter,
      word: title,
      title,
      text: `texts/${name}.txt`,
      image: `images/${name}.svg`,
    });

    console.log(`${name}  ${title}`);
  }

  fs.writeFileSync(MANIFEST, JSON.stringify(manifest, null, 2) + "\n", "utf8");
  console.log(`cover: ${ensureCover()}`);
  console.log(
    `\n${stories.length} stories, ${stories.length * EXPECTED_SENTENCES} sentences, ` +
      `${stories.length} pictures (${(copiedBytes / 1024 / 1024).toFixed(1)} MB copied)`
  );
  console.log("wrote assets/stories.json");
}

main().catch((error) => {
  console.error(`extract-from-source failed: ${error.message}`);
  process.exit(1);
});
