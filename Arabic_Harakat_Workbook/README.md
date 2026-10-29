<div dir="rtl" align="right">

# كتاب تمارين الحركات وحروف المد

مولّد برمجي لكتاب تمارين على الحركات العربية وحروف المد، مُعدّ للطلاب المبتدئين.
يُبنى الكتاب من الصفر بصيغة **DOCX** بخط عربي واتجاه من اليمين لليسار.

هذا المشروع هو القسم الأول من كتاب الحركات: تدريبات تحليل المقاطع الصوتية
انتقلت إلى مجلده الخاص في
[`Arabic_Syllable_Analysis_Workbook`](../Arabic_Syllable_Analysis_Workbook)،
وبقي هنا صفحات الحروف وحروف المد فقط.

## مخرجات البناء

| الملف | الوصف |
|-------|-------|
| `Arabic_Harakat_Workbook.docx` | الكتاب النهائي (يُعاد توليده عند كل بناء، داخل مجلد المشروع) |
| `Arabic_Harakat_Workbook.pdf` | نسخة PDF (ينشئها `build-all.bat` في المجلد الأعلى) |

> ملاحظة: يُكتب ملف DOCX دائماً في مجلد المشروع، أي بغضّ النظر عن مجلد العمل الحالي (CWD).
> واسم الملف مشتقّ من اسم المجلد، فأي إعادة تسمية للمجلد تُعيد تسمية الناتج معه.

## عدد الصفحات

| القسم | العدد |
|-------|-------|
| صفحة الغلاف | 1 |
| صفحة الحركات (صفحتان لكل حرف × 28) | 56 |
| **الإجمالي** | **57 صفحة** |

## البنية العامة

```
Arabic_Harakat_Workbook/
├── create-workbook.js              # سكريبت بناء الكتاب (نقطة الدخول)
├── package.json                    # إعدادات المشروع والسكربتات
├── package-lock.json
├── README.md                       # هذا الملف
├── Arabic_Harakat_Workbook.docx    # ناتج البناء
├── Arabic_Harakat_Workbook.pdf     # نسخة PDF (ينشئها build-all.bat)
├── node_modules/
└── assets/
    ├── letters.json                # بيانات الكتاب + الأقسام + الغلاف + الحروف المستثناة
    ├── harakat.json                # الحركات + ترتيب العرض + عناوين الجداول والنصوص
    ├── cover-arabic.png            # صورة الغلاف (قابلة للاستبدال)
    └── modod.png                   # صورة محفوظة، غير مستخدمة في البناء حالياً
```

> لا يحتاج هذا المشروع إلى خطوط مخزَّنة داخله؛ الخطوط المستخدمة مثبَّتة على النظام.
> `cover-arabic.png` و`modod.png` ملفات صور فقط؛ `modod.png` غير مذكور في أي
> مكان من الشيفرة ولا يؤثر في البناء.

## بنية الكتاب

### صفحة الغلاف
صورة PNG تملأ الصفحة A5 بالكامل (full-bleed) بلا إطار وبلا هوامش. تُبحث عن
الصورة بالترتيب المذكور في `cover.candidates` داخل `assets/letters.json`، وأول
صورة موجودة هي المستخدمة. إن لم يُعثر على أي صورة، تُطبع تحذير وتُتخطى صفحة
الغلاف بدل أن يفشل البناء.

### صفحة الحركات (صفحة لكل حرف)
1. عنوان `حرف {اسم الحرف}` — 36pt، لون `2E86AB`
2. الحرف كبيراً — 72pt، لون `1B4965`
3. عنوان فرعي `الحركات وحروف المد` — 24pt، لون `2E86AB`
4. جدول الحركات

### صفحة التدريب (صفحة لكل حرف)
جدولان بفواصل، يُقرآن من اليمين لليسار في كل صف:
**متصل ← منقّط ← صف فارغ ← صف فارغ**

- **الجدول الأول — الحركات:** صف لكل حركة (فتحة، ضمة، كسرة، سكون)
- **الجدول الثاني — حروف المد:** صف لكل حرف مد (ألف، واو، ياء)
- الخلية الأولى من كل صف بخلفية رمادية فاتحة `F5F5F5` (نص الخط المتصل)، والثانية بخط منقّط
- ارتفاع الصفوف 800 تويب، والحدود رمادية `34495E`

## اختيار القسم

`assets/letters.json` ← `sections` يحدد أيّ الأقسام يُبنى في هذا المشروع:

```json
"sections": ["letters"]
```

القيم المعروفة: `letters` (صفحات الحروف) و`syllable` (تدريبات المقاطع).
هذا المشروع يبني قسم `letters` فقط، وتدريبات المقاطع الصوتية في
`Arabic_Syllable_Analysis_Workbook`.

> `create-workbook.js` **مطابق بايت ببايت** لنسخته في
> `Arabic_Syllable_Analysis_Workbook`؛ الفرق بين الكتابين في ملف
> `assets/letters.json` فقط (`sections` + بيانات `book`).

## الحركات وحروف المد

> المصدر الرسمي هو `harakat` في `assets/harakat.json`، والجدول التالي للعرض فقط.

| الحركة | الرمز (Unicode) | حرف المد | الرمز | المثال |
|--------|-----------------|----------|-------|--------|
| الفتحة | U+064E | الألف | ا (U+0627) | بَ → بَا |
| الضمة | U+064F | الواو | و (U+0648) | بُ → بُو |
| الكسرة | U+0650 | الياء | ي (U+064A) | بِ → بِي |
| السكون | U+0652 | — | — | بْ |

ألوان الحركات المستخدمة في الجداول:
الفتحة `E74C3C` · الضمة `3498DB` · الكسرة `27AE60` · السكون `9B59B6`

## معالجة خاصة بحرف الهمزة (أ)

الهمزة تحتاج معالجة مختلفة عن باقي الحروف، وهي مُطبَّقة في الدوال التالية. أي حرف آخر يحتاج معالجة مماثلة يمكن إضافته من `baseLetterOverrides` في `assets/harakat.json` دون تعديل الشيفرة:

| الحالة | السلوك |
|--------|--------|
| الهمزة + كسرة | تُكتب **إ** (U+0625) بدل أ (الدالة `getHarakaBaseLetter`) |
| جدول صفحة الحركات | جدول **مُنقل** 2×4: صف بأسماء الحركات وصف بالحروف عليها، بلا حروف مد |
| الجدول الثاني في صفحة التدريب | همزة على السطر **ء** (U+0621) مع الحركات الأربع، بدل حروف المد |

## الحروف المستخدمة

> المصدر الرسمي لهذه البيانات هو `assets/letters.json`، والجدول التالي للعرض فقط.

الحروف مرتّبة **تعليمياً** في أربع مجموعات (7 + 7 + 6 + 8 = 28 حرفاً)، وليست بترتيب أبجد هوز:

| # | المجموعة | الحروف |
|---|----------|--------|
| 1 | الأولى | أ ب ج د هـ و ز |
| 2 | الثانية | ح ط ي ك ل م ن |
| 3 | الثالثة | س ع ف ص ق ر |
| 4 | الرابعة | ش ت ث خ ذ ض ظ غ |

> حرف الهاء يُخزَّن كرمز `هـ` ليُعرض بالشكل الصحيح.

## إعدادات الصفحة

| الإعداد | القيمة |
|---------|--------|
| المقاس | A5 (14.8 × 21 سم) |
| الاتجاه | عمودي |
| الهوامش | 0.3 بوصة أعلى وأسفل · 0.5 بوصة يمين ويسار · 0 لصفحة الغلاف |
| اتجاه المستند | RTL |
| الخط الافتراضي | Traditional Arabic — 14pt |
| الإطار | خط واحد حول كل صفحة عدا صفحة الغلاف |
| تذييل/ترقيم | **لا يوجد** (بعكس كتاب الحروف) |

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
start Arabic_Harakat_Workbook.docx
# macOS
open Arabic_Harakat_Workbook.docx
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
| **docx ^9.7.1** | إنشاء ملفات Word برمجياً (MIT) — دعم RTL وجداول |

### الخطوط

| الخط | الاستخدام | الترخيص / المصدر |
|------|-----------|-------------------|
| Traditional Arabic | كل النصوص والمقاطع | خط نظام |
| KFGQPC Alphabet Dotted | الصفوف المنقّطة للتدريب | مجمع الملك فهد لطباعة المصحف الشريف |

## ملاحظات تقنية

### قبل التشغيل
- تأكد من وجود خطَّي **Traditional Arabic** و**KFGQPC Alphabet Dotted** على النظام، وإلا ظهر النص بخط بديل.

### بنية الشيفرة
| الدالة | المسؤولية |
|--------|-----------|
| `loadContent` | قراءة ملفي البيانات وفحص كل حقل قبل البناء |
| `pageProperties` | خصائص الصفحة الموحّدة (A5 + RTL + الإطار) |
| `createCoverPage` | صفحة الغلاف الممتلئة، أو `null` إن لم توجد صورة |
| `resolveRelativePath` | أول مسار موجود من قائمة مسارات |
| `getPngDimensions` | قراءة أبعاد الصورة من ترويسة PNG |
| `createDisplayPage` | صفحة الحركات (جدول الحركات أو جدول الهمزة المنقول) |
| `createPracticePage` | صفحة التدريب (جدول الحركات + جدول حروف المد/الهمزة) |
| `transposeGrid` | قلب شبكة الخلايا من أعمدة إلى صفوف |
| `getHarakaBaseLetter` | استبدال حرف الحركة حسب `baseLetterOverrides` (أ + كسرة ← إ) |
| `createSyllableExercise` | تمرين تحليل المقطعين لحرف واحد (لا يعمل في هذا المشروع) |
| `createWorkbook` | نقطة الدخول: بناء المستند وحفظه |

> `createSyllableExercise` باقية في الملف رغم أن `sections` لا تحتوي
> `syllable`، حفاظاً على تطابق السكربتين بايت ببايت.

## ملفات البيانات

كل محتوى الكتاب موجود في `assets/`، و`create-workbook.js` لا يحتوي أي نص أو حرف أو كلمة.

### `assets/letters.json`

| الحقل | الوصف |
|-------|-------|
| `book` | بيانات الملف: `creator` و`title` و`description` |
| `sections` | الأقسام التي يبنيها هذا المشروع (`letters` هنا) |
| `cover` | `file` للعرض + `candidates` قائمة مسارات تُبحث بالترتيب |
| `letters` | قائمة الحروف مرتّبة: `letter` و`name` (الترتيب في القائمة هو ترتيب ظهور الحروف في الكتاب) |
| `syllableExcludedLetters` | الحروف المستثناة من تدريبات المقاطع الصوتية (يُستخدم في المشروع الآخر) |

### `assets/harakat.json`

| الحقل | الوصف |
|-------|-------|
| `harakat` | الحركات الأربع: `name` و`symbol` و`longVowel` و`longSymbol` و`description` (`null` لحرفَي المد في السكون) |
| `baseLetterOverrides` | قواعد استبدال الحرف قبل الحركة، مثل `أ` + كسرة ← `إ` |
| `displayOrder` | ترتيب عرض الرموز في الجداول: `hamzaTableSymbols` و`harakaSymbols` و`longVowelDottedSymbols` و`syllableVowels` |
| `texts` | نصوص الكتاب: عناوين الصفحات (تحتوي `{name}`) وعناوين الأعمدة وحرف الهمزة |

### فحوصات ما قبل البناء

يتوقف البناء برسالة خطأ واضحة إذا:

- ملف بيانات مفقود أو ليس JSON صالحاً
- `sections` فارغة، أو تحتوي قيمة غير معروفة، أو تكرّر قسم
- `cover.file` أو `cover.candidates` مفقود أو فارغ
- حقل نصي مطلوب فارغ، أو عنوان صفحة بلا `{name}`
- حرف أو رمز حركة مكرَّر
- `longVowel` و`longSymbol` مضبوطان أحدهما دون الآخر
- حرف أو حركة في `baseLetterOverrides` أو `displayOrder` غير معرَّفة
- عنوان جدول أو نص صفحة مفقود
- حرف مستثنى من تدريبات المقاطع الصوتية غير موجود في `letters.json`

### التخصيص
- **تغيير الأقسام المبنية:** عدّل `assets/letters.json` ← `sections`
- **إضافة/تعديل الحروف:** عدّل `assets/letters.json` ← `letters`
- **تعديل الحركات:** عدّل `assets/harakat.json` ← `harakat`
- **تعديل الحركات المستثناة من تدريبات المقاطع:** عدّل `syllableExcludedLetters`
- **تعديل النصوص والعناوين:** عدّل `assets/harakat.json` ← `texts`
- **تغيير صورة الغلاف:** استبدل `assets/cover-arabic.png`، أو انقلها إلى
  `assets/images/cover-arabic.png` (كلاهما في قائمة `candidates`)
- **عدد تمارين الصفحة:** الثابت `exercisesPerPage` في `createWorkbook`
- **ترتيب الصفوف في جداول التدريب:** يعتمد على `transposeGrid` مع `visuallyRightToLeft: true` — تعديل أحدهما يتطلب تعديل الآخر

## الترخيص

الكود: MIT · خطوط KFGQPC: مجمع الملك فهد لطباعة المصحف الشريف

## مثال سريع للتخصيص

لتعديل حركة أو نصّ، عدّل `assets/harakat.json`:

```json
{
  "name": "الفتحة",
  "symbol": "َ",
  "longVowel": "الألف",
  "longSymbol": "ا",
  "description": "فتحة مع الألف"
}
```

ولتغيير الحروف المبنية أو صورة الغلاف، عدّل `assets/letters.json`
(`letters` و`cover.candidates`)، ثم أعد البناء.

## الإصدار

**1.0.0** — آخر تحديث: **2026-09-27**

</div>

---

## English

A generator for an **Arabic vowel-marks (harakat) and long-vowels workbook**. Built from
scratch as a right-to-left **DOCX** (font *Traditional Arabic*), then converted to **PDF**
with Microsoft Word.

This project is the *letters* half of the original harakat book; the syllable-analysis
exercises moved to [`Arabic_Syllable_Analysis_Workbook`](../Arabic_Syllable_Analysis_Workbook).

### Output

| File | Description |
|------|-------------|
| `Arabic_Harakat_Workbook.docx` | final book, written inside the project folder |
| `Arabic_Harakat_Workbook.pdf` | PDF version (via the root `build-all.bat`) |

### Page count

| Section | Count |
|---------|-------|
| Cover page | 1 |
| Harakat page (two per letter × 28) | 56 |
| **Total** | **57 pages** |

### Structure

- **Cover:** full-bleed A5 `cover-arabic.png` (skipped with a warning if absent).
- **Display page (per letter):** title, the letter (72pt), subtitle `الحركات وحروف المد`, then the harakat table.
- **Practice page (per letter):** two bordered tables read right-to-left — the harakat table (fatha, damma, kasra, sukun) and the long-vowels table (alef, waw, ya). The first cell of each row is shaded `F5F5F5` (connected font); the second uses the dotted font.
- The built section is chosen by `sections` in `assets/letters.json` (`["letters"]` here).
- Special case: hamza (`أ`) + kasra is written as `إ`, and its display table is transposed 2×4. Rules live in `baseLetterOverrides` in `assets/harakat.json`.
- `create-workbook.js` is **byte-identical** to the one in `Arabic_Syllable_Analysis_Workbook`; only `assets/letters.json` (`sections` + `book`) differs.

### Requirements & run

- Node.js 16+
- System fonts: **Traditional Arabic** (connected) and **KFGQPC Alphabet Dotted** (dotted practice)

```bash
npm install      # optional: the root node_modules is shared
npm run build    # or: node create-workbook.js
```

### Customize

- Built sections: `assets/letters.json` → `sections`
- Letters: `assets/letters.json` → `letters`
- Harakat and labels: `assets/harakat.json`
- Cover image: replace `assets/cover-arabic.png`

### License

Code: MIT. Font *KFGQPC Alphabet Dotted* © King Fahd Glorious Quran Printing Complex.
