#!/usr/bin/env node
/**
 * extract-from-source.js
 *
 * One-time (but re-runnable) importer that turns `source.docx` into the
 * project's data assets:
 *
 *   assets/texts/NN-slug.txt    line 1 = title, remaining lines = the 3 sentences
 *   assets/images/NN-slug.jpg   story picture, downscaled and re-encoded
 *   assets/stories.json         ordered index used by create-workbook.js
 *   assets/coverpage.png        placeholder cover (only created when missing)
 *
 * `create-workbook.js` reads nothing but these assets, so the workbook can be
 * rebuilt without ever touching source.docx again.
 *
 * Usage: node extract-from-source.js
 */

const fs = require("fs");
const path = require("path");
const JSZip = require("jszip");
const Jimp = require("jimp");

const ROOT = __dirname;
const SOURCE_DOCX = path.join(ROOT, "source.docx");
const ASSETS = path.join(ROOT, "assets");
const TEXTS = path.join(ASSETS, "texts");
const IMAGES = path.join(ASSETS, "images");
const COVER = path.join(ASSETS, "coverpage.png");
const MANIFEST = path.join(ASSETS, "stories.json");

const BOOK = {
  title: "كتاب القصص المصورة",
  subtitle: "هيا بنا نقرأ القصص",
  cover: "coverpage.png",
};

const MAX_IMAGE_WIDTH = 800;
const JPEG_QUALITY = 70;

// A5 aspect ratio at the same pixel size the sibling covers use.
const COVER_WIDTH = 1024;
const COVER_HEIGHT = 1536;
const COVER_COLOR = 0x1b4965ff;

/**
 * The only hand-authored content in the pipeline.
 *
 * source.docx has no story titles -- each story is just "X - word" followed by
 * three sentences -- so the titles below were written to match the scenes. Edit
 * them here and re-run; nothing else needs to change.
 *
 * `letter` and `word` are duplicated from source.docx on purpose: extract()
 * asserts they still match, so this table cannot silently drift out of order.
 */
const STORIES = [
  { slug: "asad",    letter: "أ", word: "أَسَد",   title: "أَسَدٌ في الغَابَةِ" },
  { slug: "bab",     letter: "ب", word: "باب",    title: "بَابٌ أَزْرَقُ في الحَدِيقَةِ" },
  { slug: "jamal",   letter: "ج", word: "جمل",    title: "جَمَلٌ في الصَّحْرَاءِ" },
  { slug: "dubb",    letter: "د", word: "دب",     title: "دُبٌّ صَغِيرٌ في الغَابَةِ" },
  { slug: "hilal",   letter: "هـ", word: "هلال",  title: "هِلالٌ مُضِيءٌ في السَّمَاءِ" },
  { slug: "warda",   letter: "و", word: "وردة",   title: "وَرْدَةٌ حَمْرَاءُ في الحَدِيقَةِ" },
  { slug: "zarafa",  letter: "ز", word: "زرافة",  title: "زَرَافَةٌ طَوِيلَةٌ في الحَدِيقَةِ" },
  { slug: "hisan",   letter: "ح", word: "حصان",   title: "حِصَانٌ جَمِيلٌ في الْحَقْلِ" },
  { slug: "taair",   letter: "ط", word: "طائر",   title: "طَائِرٌ صَغِيرٌ على الشَّجَرَةِ" },
  { slug: "yad",     letter: "ي", word: "يد",     title: "يَدٌ تَعُونُ صَدِيقَتَهَا" },
  { slug: "kitab",   letter: "ك", word: "كتاب",   title: "كِتَابٌ مُلَوَّنٌ على الطَّاوِلَةِ" },
  { slug: "laymun",  letter: "ل", word: "ليمون",  title: "لَيْمُونَةٌ صَفْرَاءُ على الشَّجَرَةِ" },
  { slug: "maa",     letter: "م", word: "ماء",    title: "مَاءٌ بَارِدٌ يُسْعِدُ الصِّغْرَانَ" },
  { slug: "najm",    letter: "ن", word: "نجم",    title: "نَجْمٌ لَامِعٌ في السَّمَاءِ" },
  { slug: "samak",   letter: "س", word: "سمك",    title: "سَمَكٌ مُلَوَّنٌ في الْبَحْرِ" },
  { slug: "ain",     letter: "ع", word: "عين",    title: "عَيْنٌ تَرَى جَمَالَ الْعَالَمِ" },
  { slug: "fil",     letter: "ف", word: "فيل",    title: "فِيلٌ ضَخْمٌ عِنْدَ النَّهْرِ" },
  { slug: "sanduq",  letter: "ص", word: "صندوق",  title: "صَنْدُقٌ صَغِيرٌ تَحْتَ الشَّجَرَةِ" },
  { slug: "qalb",    letter: "ق", word: "قلب",    title: "قَلْبٌ مَلِءٌ بِالْحَبِّ" },
  { slug: "rajul",   letter: "ر", word: "رجل",    title: "رَجُلٌ يَحْمِلُ أَكْيَاسًا" },
  { slug: "shams",   letter: "ش", word: "شمس",    title: "شَمْسٌ مُشْرِقَةٌ مِنَ النَّافِذَةِ" },
  { slug: "tuffah",  letter: "ت", word: "تفاح",   title: "تُفَّاحَةٌ حَمْرَاءُ مِنَ الشَّجَرَةِ" },
  { slug: "thaalab", letter: "ث", word: "ثعلب",   title: "ثَعْلَبٌ صَغِيرٌ قُرْبَ الغَابَةِ" },
  { slug: "khubz",   letter: "خ", word: "خبز",    title: "خُبْزٌ طَازِجٌ عَلَى الْمَائِدَةِ" },
  { slug: "dhiib",   letter: "ذ", word: "ذئب",    title: "ذِئْبٌ بَعِيدٌ عِنْدَ طَرَفِ الغَابَةِ" },
  { slug: "dafdaa",  letter: "ض", word: "ضفدع",   title: "ضِفْدَعٌ أَخْضَرُ قُرْبَ الْبُرْكَةِ" },
  { slug: "zarf",    letter: "ظ", word: "ظرف",    title: "ظَرْفٌ جَمِيلٌ فَوْقَ الطَّاوِلَةِ" },
  { slug: "ghazal",  letter: "غ", word: "غزال",   title: "غَزَالٌ جَمِيلٌ في الْحَقْلِ" },
];

const EXPECTED_SENTENCES = 3;
const MAX_TITLE_WORDS = 5;

// "X - word", tolerating the tatweel in the letter of story 5 (هـ - هلال).
const HEADING_RE = /^(\S+)\s+[\u2014\u2013-]\s+(\S+)$/;

function decodeEntities(s) {
  return s
    .replace(/&#x([0-9A-Fa-f]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(parseInt(d, 10)))
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

/**
 * Flatten a <w:p> to plain text, one line per <w:br/>.
 *
 * <w:drawing> / <w:pict> blocks are dropped first: they are anchored floating
 * images whose child elements carry layout text (wp:align "right",
 * wp:posOffset "561191"), which would otherwise leak into the story text.
 */
function paragraphText(pXml) {
  return decodeEntities(
    pXml
      .replace(/<w:drawing>[\s\S]*?<\/w:drawing>/g, "")
      .replace(/<w:pict>[\s\S]*?<\/w:pict>/g, "")
      .replace(/<w:br\s*\/>/g, "\n")
      .replace(/<w:br\s*[^>]*>/g, "\n")
      .replace(/<w:cr\s*\/>/g, "\n")
      .replace(/<w:tab\s*\/>/g, "\t")
      .replace(/<[^>]+>/g, "")
  );
}

function parseRelationshipTargets(relsXml) {
  const targets = new Map();
  for (const tag of relsXml.match(/<Relationship\b[^>]*>/g) || []) {
    const id = /Id="([^"]+)"/.exec(tag);
    const target = /Target="([^"]+)"/.exec(tag);
    if (id && target) targets.set(id[1], target[1]);
  }
  return targets;
}

/** Split the flat document into heading-led stories, in document order. */
function extractStories(documentXml, relTargets) {
  const paragraphs = (documentXml.match(/<w:p\b[\s>][\s\S]*?<\/w:p>/g) || []).map((xml) => ({
    lines: paragraphText(xml)
      .split("\n")
      .map((s) => s.trim())
      .filter(Boolean),
    images: [...xml.matchAll(/r:embed="([^"]+)"/g)].map((m) => relTargets.get(m[1])),
  }));

  // Images are anchored as floating shapes, so pair them by document order:
  // paragraph 13 (the heading of story 7) holds image 7, and paragraph 41
  // (body of story 20 with story 21's heading fused onto its end) holds image 20.
  const images = [];
  for (const p of paragraphs) images.push(...p.images);

  const stories = [];
  let current = null;

  for (const { lines } of paragraphs) {
    if (lines.length === 0) continue;

    if (lines.length === 1 && HEADING_RE.test(lines[0])) {
      current = { heading: lines[0], sentences: [] };
      stories.push(current);
      continue;
    }

    if (!current) {
      throw new Error(`Body text before any heading: ${JSON.stringify(lines[0])}`);
    }

    // Story 21's heading is glued onto the end of story 20's body paragraph.
    const trailing = lines[lines.length - 1];
    const fused = HEADING_RE.test(trailing);
    const body = fused ? lines.slice(0, -1) : lines;

    for (const sentence of body) current.sentences.push(sentence);

    if (fused) {
      current = { heading: trailing, sentences: [] };
      stories.push(current);
    }
  }

  return { stories, images };
}

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

function validate(stories, images) {
  assert(
    stories.length === STORIES.length,
    `Expected ${STORIES.length} stories, parsed ${stories.length}`
  );
  assert(
    images.length === STORIES.length,
    `Expected ${STORIES.length} images, found ${images.length}`
  );

  stories.forEach((story, i) => {
    const meta = STORIES[i];
    const match = HEADING_RE.exec(story.heading);
    assert(match, `Story ${i + 1}: heading is not "X - word": ${story.heading}`);

    const [, letter, word] = match;
    assert(letter === meta.letter, `Story ${i + 1}: letter "${letter}" != "${meta.letter}"`);
    assert(word === meta.word, `Story ${i + 1}: word "${word}" != "${meta.word}"`);
    assert(
      story.sentences.length === EXPECTED_SENTENCES,
      `Story ${i + 1} (${meta.word}): expected ${EXPECTED_SENTENCES} sentences, got ${story.sentences.length}`
    );
    assert(
      meta.title.split(/\s+/).length <= MAX_TITLE_WORDS,
      `Story ${i + 1} (${meta.word}): title exceeds ${MAX_TITLE_WORDS} words`
    );
  });
}

async function writeStoryImage(sourceBuffer, destination) {
  const image = await Jimp.read(sourceBuffer);
  if (image.bitmap.width > MAX_IMAGE_WIDTH) {
    image.resize(MAX_IMAGE_WIDTH, Jimp.AUTO);
  }
  // .quality() must be set on the instance; passing { quality } to
  // getBufferAsync() is silently ignored and produces ~20x larger files.
  image.quality(JPEG_QUALITY);
  const jpeg = await image.getBufferAsync(Jimp.MIME_JPEG);
  fs.writeFileSync(destination, jpeg);
  return jpeg.length;
}

async function ensureCoverPlaceholder() {
  if (fs.existsSync(COVER)) {
    const { size } = fs.statSync(COVER);
    console.log(`cover: keeping existing assets/coverpage.png (${formatBytes(size)})`);
    return;
  }
  const placeholder = new Jimp(COVER_WIDTH, COVER_HEIGHT, COVER_COLOR);
  const png = await placeholder.getBufferAsync(Jimp.MIME_PNG);
  fs.writeFileSync(COVER, png);
  console.log(
    `cover: wrote PLACEHOLDER assets/coverpage.png ` +
      `(${COVER_WIDTH}x${COVER_HEIGHT} solid #1B4965) - replace it with the real cover`
  );
}

function formatBytes(bytes) {
  return bytes >= 1024 * 1024
    ? `${(bytes / 1024 / 1024).toFixed(2)} MB`
    : `${(bytes / 1024).toFixed(0)} KB`;
}

async function main() {
  assert(fs.existsSync(SOURCE_DOCX), `Missing source file: ${SOURCE_DOCX}`);

  const zip = await JSZip.loadAsync(fs.readFileSync(SOURCE_DOCX));
  const [documentXml, relsXml] = await Promise.all([
    zip.file("word/document.xml").async("string"),
    zip.file("word/_rels/document.xml.rels").async("string"),
  ]);

  const relTargets = parseRelationshipTargets(relsXml);
  const { stories, images } = extractStories(documentXml, relTargets);
  validate(stories, images);

  for (const dir of [ASSETS, TEXTS, IMAGES]) fs.mkdirSync(dir, { recursive: true });

  const manifest = { book: BOOK, stories: [] };
  let totalImageBytes = 0;
  let totalSourceBytes = 0;

  for (let i = 0; i < stories.length; i++) {
    const meta = STORIES[i];
    const name = `${String(i + 1).padStart(2, "0")}-${meta.slug}`;
    const textPath = path.join(TEXTS, `${name}.txt`);
    const imagePath = path.join(IMAGES, `${name}.jpg`);
    // JSZip paths are always POSIX, never path.join()'d on Windows.
    const mediaPath = `word/${images[i]}`;

    const body = [meta.title, ...stories[i].sentences].join("\n") + "\n";
    fs.writeFileSync(textPath, body, "utf8");

    const sourceBuffer = await zip.file(mediaPath).async("nodebuffer");
    totalSourceBytes += sourceBuffer.length;
    totalImageBytes += await writeStoryImage(sourceBuffer, imagePath);

    manifest.stories.push({
      id: i + 1,
      letter: meta.letter,
      word: meta.word,
      title: meta.title,
      text: `texts/${name}.txt`,
      image: `images/${name}.jpg`,
    });

    console.log(`${name}  ${meta.title}`);
  }

  fs.writeFileSync(MANIFEST, JSON.stringify(manifest, null, 2) + "\n", "utf8");
  await ensureCoverPlaceholder();

  console.log(
    `\n${manifest.stories.length} stories, ` +
      `${manifest.stories.length * EXPECTED_SENTENCES} sentences, ` +
      `${manifest.stories.length} images`
  );
  console.log(`images: ${formatBytes(totalSourceBytes)} -> ${formatBytes(totalImageBytes)}`);
  console.log(`wrote assets/stories.json`);
}

main().catch((error) => {
  console.error(`extract-from-source failed: ${error.message}`);
  process.exit(1);
});
