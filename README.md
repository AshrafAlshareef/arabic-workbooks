# مكتبة تمارين العربية

سبعة **كتب تمارين عربية** للطلاب المبتدئين، تُبنى برمجياً من الصفر بصيغة **DOCX**
بخط عربي واتجاه من اليمين لليسار، ثم تُحوَّل إلى **PDF** عبر Microsoft Word وتُدمج
في كتاب واحد مرتَّب حسب ترتيب القراءة.

| # | الكتاب | مجلد المشروع | الصفحات |
|---|--------|--------------|---------|
| 1 | كتاب الحروف العربية | `Arabic_Alphabet_Workbook` | 57 |
| 2 | كتاب الحركات | `Arabic_Harakat_Workbook` | 57 |
| 3 | كتاب تحليل المقاطع الصوتية | `Arabic_Syllable_Analysis_Workbook` | 9 |
| 4 | كتاب تحليل الكلمات إلى أحرف | `Arabic_Word_Analysis_Workbook` | 8 |
| 5 | كتاب أشكال الحروف في الكلمة | `Arabic_Letter_Shapes_Workbook` | 8 |
| 6 | كتاب تجميع الحروف لتكوين كلمة | `Arabic_Word_Composition_Workbook` | 8 |
| 7 | كتاب القصص | `Arabic_Stories_Workbook` | 30 |
| | **الكتاب المدمج الكامل** | `Arabic_Workbooks_Complete.pdf` | **177** |

## التحميل / Download

كل كتاب جاهز للتحميل مباشرة من هذا المستودع بصيغتَي **PDF** و**DOCX** (محرَّر):

| الكتاب | PDF | DOCX |
|--------|-----|------|
| الحروف | [`Arabic_Alphabet_Workbook.pdf`](Arabic_Alphabet_Workbook/Arabic_Alphabet_Workbook.pdf) | [`docx`](Arabic_Alphabet_Workbook/Arabic_Alphabet_Workbook.docx) |
| الحركات | [`Arabic_Harakat_Workbook.pdf`](Arabic_Harakat_Workbook/Arabic_Harakat_Workbook.pdf) | [`docx`](Arabic_Harakat_Workbook/Arabic_Harakat_Workbook.docx) |
| تحليل المقاطع | [`Arabic_Syllable_Analysis_Workbook.pdf`](Arabic_Syllable_Analysis_Workbook/Arabic_Syllable_Analysis_Workbook.pdf) | [`docx`](Arabic_Syllable_Analysis_Workbook/Arabic_Syllable_Analysis_Workbook.docx) |
| تحليل الكلمات | [`Arabic_Word_Analysis_Workbook.pdf`](Arabic_Word_Analysis_Workbook/Arabic_Word_Analysis_Workbook.pdf) | [`docx`](Arabic_Word_Analysis_Workbook/Arabic_Word_Analysis_Workbook.docx) |
| أشكال الحروف | [`Arabic_Letter_Shapes_Workbook.pdf`](Arabic_Letter_Shapes_Workbook/Arabic_Letter_Shapes_Workbook.pdf) | [`docx`](Arabic_Letter_Shapes_Workbook/Arabic_Letter_Shapes_Workbook.docx) |
| تجميع الكلمات | [`Arabic_Word_Composition_Workbook.pdf`](Arabic_Word_Composition_Workbook/Arabic_Word_Composition_Workbook.pdf) | [`docx`](Arabic_Word_Composition_Workbook/Arabic_Word_Composition_Workbook.docx) |
| القصص | [`Arabic_Stories_Workbook.pdf`](Arabic_Stories_Workbook/Arabic_Stories_Workbook.pdf) | [`docx`](Arabic_Stories_Workbook/Arabic_Stories_Workbook.docx) |
| **المدمج (177 صفحة)** | [`Arabic_Workbooks_Complete.pdf`](Arabic_Workbooks_Complete.pdf) | — |

> **تحميل كتاب واحد:** اضغط على رابط الـPDF في صفِّه. **تحميل الكتاب الكامل:**
> اضغط `Arabic_Workbooks_Complete.pdf`.

> الكتب 4 و5 و6 (**كتب الكلمات**) كانت كتاباً واحداً `Arabic_Word_Analysis_Workbook`
> يجمع ثلاثة تمارين في 22 صفحة، وفُصلت الآن كل تمرين في مشروع مستقل يبني كتاباً
> مستقلاً بغلافه. `create-workbook.js` **متطابق في المشاريع الثلاثة** (سوى سطرين)،
> والتمرين المطبوع يحدده الحقل `order` في `assets/exercise-types.json` داخل كل مشروع.

> الكتب 2 و3 (**كتاب الحركات**) كانا كتاباً واحداً `Arabic_Harakat_Workbook`
> يجمع صفحات الحروف (56) وتدريبات المقاطع الصوتية (8) في 64 صفحة، وفُصل الآن
> كل قسم في مشروع مستقل. `create-workbook.js` **متطابق بايت ببايت في المشروعين**،
> والقسم المطبوع يحدده الحقل `sections` في `assets/letters.json` داخل كل مشروع.
> أُضيف للكتابين غلافٌ لم يكن في الكتاب الأصلي، ولذلك صار مجموع الصفحات 177
> بدل 175 مع بقاء المحتوى على حاله.

> كل كتاب من الكتب السبعة **مستقل** بملفاته وصوره ومخرجاته، وتفاصيل كل واحد
> (بنية الصفحات، المقاسات، الخطوط، الصور) موثّقة في `README.md` داخل مجلده.
> هذا الملف يشرح **المكتبة ككل** وطريقة بنائها.

## المتطلبات

| المتطلب | الإصدار | لماذا |
|---------|---------|-------|
| **Node.js** | 16 أو أحدث | بناء ملفات DOCX وتحويل SVG إلى PNG |
| **Microsoft Word** | أي إصدار حديث | الطريقة الوحيدة الموثوقة لتحويل DOCX إلى PDF |
| خط **Traditional Arabic** | مثبَّت على النظام | الخط المتصل لكل الكتب |
| خط **KFGQPC Alphabet Dotted** | مثبَّت على النظام | سطور التدريب المنقّطة (كتاب الحروف والحركات) |

> إنذار `EBADENGINE` أثناء `npm install` بسبب حزمة `nanoid` (تطلب Node 22/24/26)
> **لا يؤثر** على البناء.

## البناء بأمر واحد

من داخل هذا المجلد:

```bat
build-all.bat
```

الأمر ينفّذ ثلاث مراحل بالترتيب، و**يتوقف عند أول خطأ**:

| المرحلة | ماذا تفعل | الأداة |
|---------|-----------|--------|
| **1/3** بناء DOCX | يبني الكتب السبعة بالترتيب (الحروف ← الحركات ← تحليل المقاطع ← تحليل الكلمات ← أشكال الحروف ← تجميع الكلمات ← القصص) | `node create-workbook.js` داخل كل مجلد |
| **2/3** تحويل PDF | يحوّل كل DOCX إلى PDF بجواره عبر Word | `docx-to-pdf.ps1` ← `convert-one.ps1` |
| **3/3** الدمج | يدمج ملفات PDF السبعة في كتاب واحد | `node merge-pdfs.js` |

> **أغلق Microsoft Word قبل البناء.** وجود Word مفتوح يجعل تصدير PDF غير موثوق،
> وقد يفشل الملف عند حذف نسخة PDF قديمة يكون قد فتحها برنامج آخر.

### خيارات البناء

| الأمر | الوظيفة |
|-------|---------|
| `build-all.bat` | بناء كامل: DOCX ← PDF ← دمج |
| `build-all.bat /nobuild` | تخطّي بناء DOCX: والتحويل والدمج فقط |
| `build-all.bat /no-build` | نفس `/nobuild` |
| `build-all.bat /merge` | نفس `/nobuild` |
| `set BUILD_ALL_NO_PAUSE=1` | لا يتوقف البناء بانتظار مفتاح Enter (مفيد في النصوص البرمجية) |

> `/nobuild` لا يصلح بعد تعديل ملفات `assets/*.json`، لأنه يتخطّي مرحلة البناء
> ويستخدم ملفات DOCX الموجودة فقط.

### تثبيت الاعتماديات

الاعتماديات الثلاث (`docx` و`pdf-lib` و`svg2img`) **مشتركة** في `node_modules`
الجذر، فتشمل كتب الحروف والحركات وكتب الكلمات الثلاثة والقصص معاً. يثبّت
`build-all.bat` ما ينقص منها تلقائياً عند أول تشغيل، و**لا يحتاج** إلى `npm install`
يدوي.

> **مجلد `node_modules` واحد في الجذر فقط:** لا يمتلك أي مشروع كتابٍ مجلدَ
> `node_modules` خاصاً به، لأن Node يرتقي تلقائياً في شجرة المجلدات حتى يجد
> الاعتماديات في الجذر. فـ**تثبيت واحد يكفي** للكتب السبعة ولأداة الصور جميعاً.

## أداة الصور (`image_2_cover_style`)

أداة **مستقلة** عن `build-all.bat`: تحوّل كل صورة في مجلد `image_2_cover_style\images\`
إلى صفحة كاملة واحدة بمقاس A5 **بأسلوب الأغلفة** نفسه (صورة تملأ الورقة، بلا هوامش
ولا إطار ولا تذييل)، ثم تدمجها في `images-to-book.docx` و `images-to-book.pdf`.

```bat
image_2_cover_style\images-to-book.bat
```

| الأمر | النتيجة |
|-------|---------|
| `images-to-book.bat` | A5، ويُنتج DOCX و PDF |
| `images-to-book.bat --size a4` | مقاس A4 بدل A5 |
| `images-to-book.bat --no-pdf` | DOCX فقط، بلا حاجة إلى Word |

> مجلد الإدخال والإخراج **ثابتان**: لا يوجد وسيط للمجلدات، فضع الصور في
> `image_2_cover_style\images\` وشغّل. الأداة تستعمل `convert-one.ps1` من الجذر
> لتحويل PDF، وتختار لكل صورة بين القصّ الكامل (full-bleed) والاحتواء الكامل حسب
> نسبة القصّ. التفاصيل في `image_2_cover_style\README.md`.

## الأغلفة والصور الجاهزة (`covers-and-photos`)

مجلد `covers-and-photos/` يجمع **أغلفةً وصوراً جاهزة** يمكن إعادة استخدامها في
الكتب الخاصة: أغلفة (`cover-arabic.png`، `harakat.png`، `tahlel.png`،
`stories cover.png`) وصور تعليمية (`kitab.png`، `way.png`،
`MAP.png`، `modod.png`، `lastpage.png`). بعضها نسخة مطابقة لأغلفة داخل مشاريع
الكتب، جُمعت هنا لتسهيل إعادة الاستخدام.

> هذه الملفات من إنشاء الذكاء الاصطناعي، ومتاحة تحت رخصة المحتوى **CC BY 4.0**
> (انظر `LICENSE-CONTENT`).

## المخرجات

| الملف | المُنشئ |
|-------|--------|
| `Arabic_Alphabet_Workbook\Arabic_Alphabet_Workbook.docx` / `.pdf` | المرحلة 1 و2 |
| `Arabic_Harakat_Workbook\Arabic_Harakat_Workbook.docx` / `.pdf` | المرحلة 1 و2 |
| `Arabic_Syllable_Analysis_Workbook\Arabic_Syllable_Analysis_Workbook.docx` / `.pdf` | المرحلة 1 و2 |
| `Arabic_Word_Analysis_Workbook\Arabic_Word_Analysis_Workbook.docx` / `.pdf` | المرحلة 1 و2 |
| `Arabic_Letter_Shapes_Workbook\Arabic_Letter_Shapes_Workbook.docx` / `.pdf` | المرحلة 1 و2 |
| `Arabic_Word_Composition_Workbook\Arabic_Word_Composition_Workbook.docx` / `.pdf` | المرحلة 1 و2 |
| `Arabic_Stories_Workbook\Arabic_Stories_Workbook.docx` / `.pdf` | المرحلة 1 و2 |
| `Arabic_Workbooks_Complete.pdf` | المرحلة 3 (177 صفحة) |

- كل **DOCX** يُكتب **داخل مجلد مشروعه** (`__dirname`)، فلا يختلف مساره مهما كان
  المجلد الذي شغّلت منه الأمر. واسم الملف **مشتق من اسم مجلد المشروع**، فيحمل كل
  كتاب نسخةً باسم مجلده حتى لو نُسخ المشروع إلى مجلد باسم آخر.
- ملفات **PDF** تُكتب بجوار كل DOCX داخل مجلد المشروع نفسه.
- الدمج **لا يستخدم أي نسخة مؤرشفة**: إن نقص ملف PDF لأي كتاب تتوقف العملية
  بخطأ، فلا يمكن أن يختلط الكتاب الجديد بصفحات قديمة.

## بنية المجلد

```
created-books/
├── build-all.bat                  # نقطة الدخول الوحيدة: docx -> pdf -> دمج
├── docx-to-pdf.ps1                # يحوّل الكتب السبعة، كل كتاب في عملية منفصلة
├── convert-one.ps1                # يحوّل ملف DOCX واحد إلى PDF عبر Word COM
├── merge-pdfs.js                  # يدمج ملفات PDF السبعة (pdf-lib)
├── package.json                   # الاعتماديات المشتركة (node_modules الجذر)
├── README.md                      # هذا الملف
├── LICENSE                        # رخصة الكود (MIT)
├── LICENSE-CONTENT                # رخصة المحتوى والصور (CC BY 4.0)
├── Arabic_Workbooks_Complete.pdf  # ناتج الدمج
├── node_modules/                  # اعتماديات مشتركة (مُولَّدة، غير مُتتبَّعة)
│
├── Arabic_Alphabet_Workbook/      # كتاب الحروف           (57 صفحة)
├── Arabic_Harakat_Workbook/       # كتاب الحركات          (57 صفحة)
├── Arabic_Syllable_Analysis_Workbook/ # كتاب تحليل المقاطع (9 صفحات)
├── Arabic_Word_Analysis_Workbook/ # كتاب تحليل الكلمات   (8 صفحات)
├── Arabic_Letter_Shapes_Workbook/ # كتاب أشكال الحروف     (8 صفحات)
├── Arabic_Word_Composition_Workbook/ # كتاب تجميع الكلمات (8 صفحات)
├── Arabic_Stories_Workbook/       # كتاب القصص            (30 صفحة)
│
├── image_2_cover_style/           # أداة مستقلة: صور مجلد images -> كتاب بأسلوب الأغلفة
│
└── covers-and-photos/             # أغلفة وصور جاهزة لإعادة الاستخدام
```

كل مجلد مشروع يحتوي: `create-workbook.js` (نقطة الدخول) و`package.json`
و`README.md` و`assets/` ومخرجاته.

## تعديل المحتوى

> **القاعدة:** المحتوى التعليمي في ملفات JSON داخل `assets/`، والتنسيق والتخطيط
> والألوان والمقاسات في `create-workbook.js`. لتغيير **ماذا** يُطبع، عدّل JSON؛
> لتغيير **كيف** يُطبع، عدّل الكود.

| الكتاب | ملفات البيانات | أمثلة على ما يمكن تعديله |
|--------|----------------|--------------------------|
| الحروف | `assets/letters.json` · `assets/letter-shapes.json` | الحروف، الأشكال، العناوين، ترتيب الأقسام، بيانات الملف |
| الحركات | `assets/letters.json` · `assets/harakat.json` | الحركات ورموزها وترتيب عرضها، استثناءات الأشكال، النصوص، الغلاف، الأقسام المبنية |
| تحليل المقاطع الصوتية | `assets/letters.json` · `assets/harakat.json` | الحروف المستثناة من التدريبات، نصوص صفحة العنوان، الغلاف |
| تحليل الكلمات | `assets/words.json` · `assets/exercise-types.json` | الكلمات وصورها، الغلاف، عناوين الأقسام، نصوص صفحات التمارين |
| أشكال الحروف | `assets/words.json` · `assets/exercise-types.json` | نفس ما سبق، وجدول `FORMS` في الكود لأشكال الحروف السياقية |
| تجميع الكلمات | `assets/words.json` · `assets/exercise-types.json` | نفس ما سبق |
| القصص | `assets/stories.json` · `assets/texts/*.txt` · `assets/images/*` | عناوين القصص وجملها، الغلاف، ترتيب القصص |

> كتب الكلمات الثلاثة تشارك نفس الـ28 كلمة ونفس الصور ونفس `create-workbook.js`،
> ويفصل بينها `order` في `assets/exercise-types.json` فقط. لتغيير الكلمات في
> الثلاثة معاً عدّل `words.json` في كل مجلد (نسخة مستقلة في كل منها).
>
> كتابا الحركات يشاركان نفس الـ28 حرفاً ونفس `create-workbook.js` (مطابق بايت
> ببايت)، ويفصل بينهما `sections` في `assets/letters.json` فقط: `"letters"`
> لكتاب الحركات و`"syllable"` لكتاب تحليل المقاطع.

## الفحوصات قبل البناء

كل مشروع من مشاريع الكتب السبعة يفحص بياناته **قبل** بناء أي صفحة، ويتوقف
برسالة خطأ دقيقة إذا:

- ملف بيانات مفقود أو ليس JSON صالحاً
- حقل نصي مطلوب فارغ، أو علم منطقي غير `true`/`false`
- قيمة مكرَّرة، أو حرف غير موجود في جدول أشكال الخط، أو مسار صورة غير معرَّف
- `order` و`types` لا يتفقان، أو قسم بلا تخطيط مقابل
- `sections` فارغة، أو فيها قيمة غير معروفة، أو قسم مكرَّر
- `cover.file` أو `cover.candidates` مفقود أو فارغ

وإذا فشل **مسار** صورة يوقف البناء، أما إذا فشل **قراءة** الصورة أو تحويلها فيُبنى
التمرين بدونها مع رسالة تحذير. وكذلك **غياب صورة الغلاف** لا يوقف البناء: تُطبع
تحذير وتُتخطى صفحة الغلاف.

## استكشاف الأخطاء

| الرسالة / العَرَض | السبب | الحل |
|-------------------|-------|------|
| `[ERROR] Node.js not found in PATH` | Node غير مثبَّت | ثبّت Node.js 16+ |
| `Converted 3/7 workbooks. Failed: ...` | Word مفتوح، أو الملف مفتوح في برنامج آخر | أغلق Microsoft Word وكل ملفات الـPDF، ثم أعد البناء |
| `The process cannot access the file because it is being used by another process` | PDF قديم مفتوح | أغلق عارض PDF ثم أعد البناء |
| `Missing PDF for 1 of 7 books` | كتاب لم يُحوَّل | شغّل `build-all.bat` كاملاً (المرحلة 2 لم تكتمل) |
| `Word reported exit code …` مع نجاح الملف | Word أبلغ عن خطأ لكن الـPDF كُتب | تجاهله إن كانت كل صفحاته صحيحة |
| `npm install failed` | انقطاع شبكة أو إصدار Node قديم جداً | تحقق من Node ثم أعد البناء |
| نص يظهر بخط مختلف | خط Traditional Arabic غير مثبَّت | ثبّت الخط على النظام |

> نسخة Word `~$*.docx` ملفات قفل مؤقتة ينظّفها `docx-to-pdf.ps1` تلقائياً.
> نسخ Word الخلفية **غير المرئية** المتبقّية من تشغيل فاشل يوقفها السكربت
> تلقائياً، لكنه **لا يمسّ** أي نافذة Word يراها المستخدم.

**1.2.0** — فصل كتاب الحركات إلى كتابين مستقلين (كتاب الحركات، تحليل المقاطع
الصوتية) بغلاف لكل كتاب، وكتاب مدمج من **177** صفحة.
سبعة كتب. آخر تحديث: **2026-09-29**

---

## English

Seven Arabic workbooks for young beginners. Each book is generated from JSON data by a
Node.js script into a right-to-left **DOCX** (script font *Traditional Arabic*), converted
to **PDF** with Microsoft Word, and the seven PDFs are merged into one complete book.

### Download

- **Combined book (177 pages):** [`Arabic_Workbooks_Complete.pdf`](Arabic_Workbooks_Complete.pdf)
- **Individual books:** the `.pdf` (and editable `.docx`) inside each `Arabic_*_Workbook/` folder — see the table at the top.

### Requirements

| Requirement | Why |
|-------------|-----|
| **Node.js 16+** | build the DOCX files and convert SVG images to PNG |
| **Microsoft Word** | the reliable DOCX → PDF conversion |
| **Traditional Arabic** font | connected script used by every book |
| **KFGQPC Alphabet Dotted** font | dotted practice rows (Alphabet & Harakat books) |

Fonts are **not** redistributed; install them on your system. A missing font makes Word fall back to a substitute.

### Build everything

```bat
build-all.bat
```

Three stages, stopping at the first error:

1. build the seven `.docx` files,
2. convert each to `.pdf` with Word,
3. merge the seven PDFs into `Arabic_Workbooks_Complete.pdf`.

### Dependencies

The shared dependencies (`docx`, `pdf-lib`, `svg2img`) live in the **single root
`node_modules/`**. A book has no `node_modules` of its own — Node resolves each
book's imports by walking up the folder tree to the root, so one install covers all
seven books and the image tool. `build-all.bat` installs anything missing on first
run; no manual `npm install` is required.

### Make your own book

The rule: **content lives in `assets/*.json`, layout/labels live in `create-workbook.js`.**

- Change the printed text by editing the JSON of the book you want, e.g.
  `assets/letters.json`, `assets/words.json`, `assets/harakat.json`,
  `assets/exercise-types.json`, or `assets/stories.json`.
- Replace an image by dropping a file into that book's `assets/images/` and pointing the
  `image` field (or `cover.candidates`) at the new filename.
- Re-run `build-all.bat`. Each book is self-contained in its own folder.
- `image_2_cover_style/` is a standalone helper that turns every image in
  `image_2_cover_style/images/` into one full-bleed A5/A4 page and produces
  `images-to-book.docx` / `.pdf`.
- Reusable covers and photos are collected in `covers-and-photos/` (CC BY 4.0).

### License

- **Code:** MIT — see [`LICENSE`](LICENSE).
- **Content and images:** CC BY 4.0 — see [`LICENSE-CONTENT`](LICENSE-CONTENT). All artwork is AI-generated.

