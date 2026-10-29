<div dir="rtl" align="right">

# كتاب تمارين الحروف العربية

مولّد برمجي لكتاب تمارين على الحروف العربية، مُعدّ للطلاب المبتدئين.
يُبنى الكتاب من الصفر بصيغة **DOCX** بخط عربي واتجاه من اليمين لليسار.

## مخرجات البناء

| الملف | الوصف |
|-------|-------|
| `Arabic_Alphabet_Workbook.docx` | الكتاب النهائي (يُعاد توليده عند كل بناء) |
| `Arabic_Alphabet_Workbook.pdf` | نسخة PDF (تُنتجها أداة البناء في المجلد الأعلى) |

> ملاحظة: يُكتب الملف `Arabic_Alphabet_Workbook.docx` في مجلد المشروع نفسه مهما كان مجلد العمل الحالي (CWD).

## عدد الصفحات

| القسم | العدد |
|-------|-------|
| صفحة الغلاف | 1 |
| صفحة الحرف (صفحة لكل حرف × 28) | 28 |
| صفحة جدول الأشكال (صفحة لكل حرف × 28) | 28 |
| **الإجمالي** | **57 صفحة** |

## بنية الكتاب

### صفحة الغلاف
- صورة `cover-arabic.png` بمقاس صفحة A5 كاملة (**full-bleed**) بلا هوامش
- **بدون** ترقيم pages ولا تذييل
- تُبحث عن الصورة بالترتيب (من `assets/letters.json` ← `cover.candidates`): `../covers-and-photos/cover-arabic.png` ← `assets/images/cover-arabic.png`
- إن لم تُوجد الصورة، تُحذف صفحة الغلاف ويُكمل البناء بدونها (مع رسالة تحذير)

### صفحة الحرف (لكل حرف)
1. عنوان `حرف {اسم الحرف}` — خط Traditional Arabic، 36pt، لون `2E86AB`
2. الحرف بشكله الرئيسي — 72pt، لون `1B4965`
3. صورة كرتونية بحجم 250×250 بكسل
4. الكلمة تبدأ بالحرف — **الحرف باللون الأحمر** `E74C3C` وباقي الكلمة بالأزرق `2C3E50`، 48pt

> حرف الهاء يستخدم الشكل `هـ` في الموضع الرئيسي (الحقل `mainForm`).

### صفحة جدول الأشكال (لكل حرف)
جدول بأربعة أعمدة و**10 صفوف**، يُقرأ من اليمين لليسار:

| نهاية منفصل | نهاية متصل | وسط | بداية |
|------------|------------|-----|-------|
| الشكل المنفصل | شكل نهاية الكلمة | شكل وسط الكلمة | شكل بداية الكلمة |

| الصف | المحتوى | الخط | اللون | الارتفاع |
|------|---------|------|-------|----------|
| 1 | رؤوس الأعمدة (خلفية زرقاء `2E86AB`، نص أبيض) | Traditional Arabic | `FFFFFF` | تلقائي |
| 2 | الأشكال بخط متصل | Traditional Arabic | `1B4965` | تلقائي |
| 3 | الأشكال بخط منقّط (للتعلّم) | KFGQPC Alphabet Dotted | `000000` | 800 تويب ثابت |
| 4 | الأشكال بخط منقّط (للتعلّم) | KFGQPC Alphabet Dotted | `000000` | 800 تويب ثابت |
| 5–10 | صفوف فارغة للتدريب | — | — | 800 تويب ثابت |

- كل الصفوف لها حدود رمادية `34495E`
- **لا يوجد** تذييل برقم الصفحة: الثابت `SHOW_PAGE_NUMBERS = false` في `create-workbook.js` هو ما يتحكم بذلك (اجعله `true` لإظهاره)

## إعدادات الصفحة

| الإعداد | القيمة |
|---------|--------|
| المقاس | A5 (14.8 × 21 سم) |
| الاتجاه | عمودي |
| الهوامش | 0.5 بوصة من الجهات الأربع (عدا الغلاف: صفر) |
| اتجاه المستند | RTL |
| الخط الافتراضي | Traditional Arabic — 14pt |

## الحروف المستخدمة

> المصدر الرسمي لهذه البيانات هو `assets/letters.json`، والجدول التالي للعرض فقط.

الحروف مرتّبة **تعليمياً** في أربع مجموعات (7 + 7 + 6 + 8 = 28 حرفاً)، وليست بترتيب أبجد هوز:

| # | المجموعة | الحروف |
|---|----------|--------|
| 1 | الأولى | أ ب ج د هـ و ز |
| 2 | الثانية | ح ط ي ك ل م ن |
| 3 | الثالثة | س ع ف ص ق ر |
| 4 | الرابعة | ش ت ث خ ذ ض ظ غ |

### بيانات الحروف والكلمات والصور

| الحرف | الاسم | الكلمة | الصورة |
|-------|-------|--------|--------|
| أ | الألف | أسد | `lion.svg` |
| ب | الباء | باب | `door.svg` |
| ج | الجيم | جمل | `camel.svg` |
| د | الدال | دب | `bear.svg` |
| ه | الهاء | هلال | `crescent.svg` |
| و | الواو | وردة | `flower.svg` |
| ز | الزاي | زرافة | `giraffe.svg` |
| ح | الحاء | حصان | `horse.svg` |
| ط | الطاء | طائر | `bird.svg` |
| ي | الياء | يد | `hand.svg` |
| ك | الكاف | كتاب | `book.svg` |
| ل | اللام | ليمون | `lemon.svg` |
| م | الميم | ماء | `water.svg` |
| ن | النون | نجم | `star.svg` |
| س | السين | سمك | `fish.svg` |
| ع | العين | عين | `eye.svg` |
| ف | الفاء | فيل | `elephant.svg` |
| ص | الصاد | صندوق | `box.svg` |
| ق | القاف | قلب | `heart.svg` |
| ر | الراء | رجل | `man.svg` |
| ش | الشين | شمس | `sun.svg` |
| ت | التاء | تفاح | `apple.svg` |
| ث | الثاء | ثعلب | `fox.svg` |
| خ | الخاء | خبز | `bread.svg` |
| ذ | الذال | ذئب | `wolf.svg` |
| ض | الضاد | ضفدع | `frog.svg` |
| ظ | الظاء | ظرف | `envelope.svg` |
| غ | الغين | غزال | `gazelle.svg` |

## الصور

- تُقرأ الصور من `assets/images/` (الترتيب في `assets/letters.json` ← `imageDirs`)
- ملفات **SVG** تُحوَّل إلى PNG عبر `svg2img`/`resvg` بعرض 400 بكسل، ثم تُدرج في المستند
- ملفات **PNG** (مثل `envelope.png`) تُدرج مباشرة
- إن فشلت قراءة الصورة أو تحويلها، تُبنى الصفحة بدونها مع رسالة تحذير
- ملفات الصور **مُخزَّنة داخل المشروع** وتُقرأ كما هي؛ لا يوجد أي سكريبت لتوليدها أو تعديلها
- إن كان مسار أي صورة خاطئاً، **يتوقف البناء برسالة خطأ** قبل إنتاج الكتاب

## ملفات البيانات

كل محتوى الكتاب موجود في `assets/`، و`create-workbook.js` لا يحتوي أي نص أو حرف أو كلمة.

### `assets/letters.json`

| الحقل | الوصف |
|-------|-------|
| `book` | بيانات الملف: `creator` و`title` و`description` |
| `imageDirs` | مجلدات البحث عن صور الحروف، بالترتيب، نسبةً لمجلد المشروع |
| `cover.file` / `cover.candidates` | اسم صورة الغلاف ثم مسارات البحث عنها |
| `pages.letterTitle` / `pages.shapesTitle` | عناوين الصفحات، وتحتويان على `{name}` |
| `table.columns` | أعمدة جدول الأشكال: `shape` (المفتاح في `letter-shapes.json`) و`header` (نص الرأس) |
| `table.dottedRows` / `emptyRows` / `dottedRowHeight` / `emptyRowHeight` | بنية الجدول: عدد الصفوف المنقّطة والفارغة وارتفاعاتها |
| `letters` | قائمة الحروف مرتّبة: `letter` و`name` و`word` و`image`، وحقل `mainForm` اختياري |

### `assets/letter-shapes.json`

كائن واحد لكل حرف، مفتاحه رمز الحرف، وقيمته الأشكال الأربعة:
`isolated` و`beginning` و`middle` و`end`.

### فحوصات ما قبل البناء

يتوقف البناء برسالة خطأ واضحة إذا:

- ملف بيانات مفقود أو ليس JSON صالحاً
- حقل نصي مطلوب فارغ، أو عنوان صفحة بلا `{name}`
- حرف مكرَّر، أو عدد الحروف صفر
- الكلمة لا تبدأ بالحرف المرتبط بها
- حرف بلا entry في `letter-shapes.json`، أو أحد أشكاله الأربعة فارغ
- `entry` في `letter-shapes.json` لحرف غير موجود في `letters.json`
- عمود جدول يشير إلى `shape` غير معروف
- مسار صورة غير موجود في `imageDirs`
- عدد صفوف أو ارتفاع غير رقم صحيح موجب

## البنية العامة

```
Arabic_Alphabet_Workbook/
├── create-workbook.js              # سكريبت بناء الكتاب (نقطة الدخول)
├── package.json                    # إعدادات المشروع والسكربتات
├── package-lock.json
├── README.md                       # هذا الملف
├── Arabic_Alphabet_Workbook.docx   # ناتج البناء
├── Arabic_Alphabet_Workbook.pdf    # نسخة PDF (من أداة البناء بالمجلد الأعلى)
├── node_modules/
└── assets/
    ├── letters.json                # بيانات الحروف والكلمات والصور + بنية الجدول
    ├── letter-shapes.json          # الأشكال الأربعة لكل حرف
    └── images/                     # صور SVG للكلمات + cover-arabic.png
```

## طريقة التشغيل

### المتطلبات
- **Node.js 16** أو أحدث
- خط **Traditional Arabic** مثبَّت على النظام (خط متصل)
- خط **KFGQPC Alphabet Dotted** مثبَّت على النظام (خط منقّط للتدريب)

### الخطوات

```bash
# 1) تثبيت الاعتماديات
npm install

# 2) بناء الكتاب
npm run build
# أو ببساطة:
node create-workbook.js
```

فتح الناتج:
```bash
# Windows
start Arabic_Alphabet_Workbook.docx
# macOS
open Arabic_Alphabet_Workbook.docx
```

### سكربتات npm
| السكربت | الوظيفة |
|---------|---------|
| `npm run build` | بناء الكتاب (نفس `npm start`) |
| `npm start` | بناء الكتاب |

## التكنولوجيا المستخدمة

| التقنية | الاستخدام |
|---------|-----------|
| **Node.js ≥ 16** | بيئة تشغيل JavaScript |
| **docx ^9.7.1** | إنشاء ملفات Word برمجياً (MIT) — دعم RTL، جداول، صور، تذييلات |
| **svg2img ^1.0.0-beta.2** | تحويل SVG إلى PNG عبر resvg |

### الخطوط

| الخط | الاستخدام | الترخيص / المصدر |
|------|-----------|-------------------|
| Traditional Arabic | كل النصوص والأشكال المتصلة | خط نظام |
| KFGQPC Alphabet Dotted | الأشكال المنقّطة للتدريب | مجمع الملك فهد لطباعة المصحف الشريف |

## الترخيص

- الكود: MIT
- خطوط KFGQPC: مجمع الملك فهد لطباعة المصحف الشريف
- الصور: رسومات SVG/PNG — الحقوق محفوظة

## ملاحظات تقنية

### قبل التشغيل
- أغلق Microsoft Word أولاً لتجنّب خطأ «الملف قيد الاستخدام».
- تأكد من وجود خطَّي **Traditional Arabic** و**KFGQPC Alphabet Dotted** على النظام، وإلا ظهر النص بخط بديل.

### التخصيص
- **تعديل الحروف/الكلمات/الصور:** عدّل `assets/letters.json` ← `letters` (الترتيب في القائمة هو ترتيب ظهور الحروف في الكتاب)
- **تعديل أشكال الحروف:** عدّل `assets/letter-shapes.json` (المفتاح = رمز الحرف)
- **تعديل عناوين الصفحات:** `assets/letters.json` ← `pages`
- **تعديل بنية جدول الأشكال ورؤوسه:** `assets/letters.json` ← `table`
- **تعديل مسار البحث عن صورة الغلاف:** `assets/letters.json` ← `cover.candidates`
- **ألوان الجدول:** ثوابت داخل الدوال `createLetterPage` و`createTablePage`
- **أبعاد الصورة داخل الكتاب:** الثابت `IMAGE_TRANSFORM` (`250 × 250`) و`IMAGE_RENDER_WIDTH` (`400`) في `create-workbook.js`

## مثال سريع للتخصيص

لإضافة حرف أو تعديل حرف موجود، عدّل عنصراً واحداً في `assets/letters.json`:

```json
{ "letter": "أ", "name": "الألف", "word": "أسد", "image": "lion.svg" }
```

ثم أضف أشكاله الأربعة في `assets/letter-shapes.json` تحت المفتاح `"أ"`، وضع الصورة
في `assets/images/`، وأعد البناء. ترتيب العناصر في `letters` هو ترتيب ظهور الحروف
في الكتاب؛ لإعادة ترتيب الصفحات أعد ترتيب القائمة فقط.

## الإصدار

**1.0.0** — آخر تحديث: **2026-09-27**

</div>

---

## English

A programmatic generator for an **Arabic alphabet workbook** for young beginners. The
book is built from scratch as a right-to-left **DOCX** (script font *Traditional Arabic*),
then converted to **PDF** with Microsoft Word.

### Output

| File | Description |
|------|-------------|
| `Arabic_Alphabet_Workbook.docx` | the final book (regenerated on every build, written inside the project folder) |
| `Arabic_Alphabet_Workbook.pdf` | PDF version (produced by the root build tool) |

### Page count

| Section | Count |
|---------|-------|
| Cover page | 1 |
| Letter page (one per letter × 28) | 28 |
| Shape-table page (one per letter × 28) | 28 |
| **Total** | **57 pages** |

### Structure

- **Cover:** `cover-arabic.png` full-bleed A5 (no margins, no footer). A missing image skips the cover with a warning.
- **Letter page:** title `حرف {name}` (36pt), the letter (72pt), a 250×250 cartoon image, and the word starting with the letter (first letter red `E74C3C`, rest blue `2C3E50`, 48pt).
- **Shape table:** four RTL columns (isolated / end / middle / beginning) and 10 rows — header, connected forms, two dotted practice rows (*KFGQPC Alphabet Dotted*), and six empty practice rows. No page-number footer (`SHOW_PAGE_NUMBERS = false`).

### Requirements & run

- Node.js 16+
- System fonts: **Traditional Arabic** (connected) and **KFGQPC Alphabet Dotted** (dotted practice)

```bash
npm install      # optional: the root node_modules is shared
npm run build    # or: node create-workbook.js
```

### Customize

- Letters / words / images / titles / table: `assets/letters.json`
- Letter forms: `assets/letter-shapes.json`
- Layout, colors, image sizes: constants in `create-workbook.js`

### License

Code: MIT. Font *KFGQPC Alphabet Dotted* © King Fahd Glorious Quran Printing Complex.
