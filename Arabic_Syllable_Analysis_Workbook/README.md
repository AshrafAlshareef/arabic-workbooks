# كتاب تدريبات تحليل المقاطع الصوتية

مولّد برمجي لكتاب تدريبات تحليل المقاطع الصوتية، مُعدّ للطلاب المبتدئين.
يُبنى الكتاب من الصفر بصيغة **DOCX** بخط عربي واتجاه من اليمين لليسار.

هذا المشروع هو القسم الثاني من كتاب الحركات: أرقام صفحات تدريبات المقاطع
الصوتية انتقلت إلى مجلده الخاص، وبقيت صفحات الحروف وحروف المد في
[`Arabic_Harakat_Workbook`](../Arabic_Harakat_Workbook).

## مخرجات البناء

| الملف | الوصف |
|-------|-------|
| `Arabic_Syllable_Analysis_Workbook.docx` | الكتاب النهائي (يُعاد توليده عند كل بناء، داخل مجلد المشروع) |
| `Arabic_Syllable_Analysis_Workbook.pdf` | نسخة PDF (ينشئها `build-all.bat` في المجلد الأعلى) |

> ملاحظة: يُكتب ملف DOCX دائماً في مجلد المشروع، أي بغضّ النظر عن مجلد العمل الحالي (CWD).
> واسم الملف مشتقّ من اسم المجلد، فأي إعادة تسمية للمجلد تُعيد تسمية الناتج معه.

## عدد الصفحات

| القسم | العدد |
|-------|-------|
| صفحة الغلاف | 1 |
| صفحة عنوان تدريبات تحليل المقاطع الصوتية | 1 |
| صفحات تدريبات تحليل المقاطع الصوتية (27 حرفاً ÷ 4 لكل صفحة) | 7 |
| **الإجمالي** | **9 صفحة** |

## البنية العامة

```
Arabic_Syllable_Analysis_Workbook/
├── create-workbook.js              # سكريبت بناء الكتاب (نقطة الدخول)
├── package.json                    # إعدادات المشروع والسكربتات
├── package-lock.json
├── README.md                       # هذا الملف
├── Arabic_Syllable_Analysis_Workbook.docx    # ناتج البناء
├── Arabic_Syllable_Analysis_Workbook.pdf     # نسخة PDF (ينشئها build-all.bat)
└── assets/
    ├── letters.json                # بيانات الكتاب + الأقسام + الغلاف + الحروف
    ├── harakat.json                # الحركات + ترتيب العرض + عناوين الجداول والنصوص
    ├── cover-arabic.png            # صورة الغلاف (قابلة للاستبدال)
    └── modod.png                   # صورة محفوظة، غير مستخدمة في البناء حالياً
```

> لا يحتاج هذا المشروع إلى `node_modules` خاص به: كل المشاريع تشترك في
> `node_modules` الموجود في المجلد الأعلى، و`build-all.bat` يتولّى التثبيت عند الحاجة.

## بنية الكتاب

### صفحة الغلاف
صورة PNG تملأ الصفحة A5 بالكامل (full-bleed) بلا إطار وبلا هوامش، تماماً
مثل كتب تحليل الكلمات. تُبحث عن الصورة بالترتيب المذكور في
`cover.candidates` داخل `assets/letters.json`، وأول صورة موجودة هي
المستخدمة. إن لم يُعثر على أي صورة، تُطبع تحذير وتُتخطى صفحة الغلاف بدل
أن يفشل البناء.

### صفحة عنوان تدريبات تحليل المقاطع الصوتية
`تدريبات تحليل المقاطع الصوتية` / `جميع الحروف`

### صفحات تدريبات تحليل المقاطع الصوتية
- **4 تمارين في كل صفحة** (حرف لكل تمرين)
- كل تمرين يعرض المقطعين الزائدين فوق ستة صناديق فارغة: صندوان لكل مقطع (2 × 3)
- تُستثنى حروف المد (أ، و، ي) من هذا القسم، و**حرف الهمزة (أ) مستثنى** أيضاً فيُبقَى 27 حرفاً
- الصناديق السفلية لها حدود، والخانات الفاصلة بين المقاطع بلا حدود

## اختيار القسم

`assets/letters.json` ← `sections` يحدد أيّ الأقسام يُبنى في هذا المشروع:

```json
"sections": ["syllable"]
```

القيم المعروفة: `letters` (صفحات الحروف) و`syllable` (تدريبات المقاطع).
هذا المشروع يبني قسم `syllable` فقط.

> `create-workbook.js` **مطابق بايت ببايت** لنسخته في
> `Arabic_Harakat_Workbook`؛ الفرق بين الكتابين في ملف `assets/letters.json`
> فقط (`sections` + بيانات `book`).

## الحروف المستخدمة

> المصدر الرسمي لهذه البيانات هو `assets/letters.json`، والجدول التالي للعرض فقط.

الحروف مرتّبة **تعليمياً** في أربع مجموعات (7 + 7 + 6 + 8 = 28 حرفاً)، وليست بترتيب أبجد هوز:

| # | المجموعة | الحروف |
|---|----------|--------|
| 1 | الأولى | أ ب ج د هـ و ز |
| 2 | الثانية | ح ط ي ك ل م ن |
| 3 | الثالثة | س ع ف ص ق ر |
| 4 | الرابعة | ش ت ث خ ذ ض ظ غ |

> حرف الهمزة (أ) مستثنى من تدريبات المقاطع الصوتية، فيُطبع 27 حرفاً فقط.

## إعدادات الصفحة

| الإعداد | القيمة |
|---------|--------|
| المقاس | A5 (14.8 × 21 سم) |
| الاتجاه | عمودي |
| الهوامش | 0.3 بوصة أعلى وأسفل · 0.5 بوصة يمين ويسار · 0 للغلاف |
| اتجاه المستند | RTL |
| الخط الافتراضي | Traditional Arabic — 14pt |
| الإطار | خط واحد حول كل صفحة عدا صفحة الغلاف |
| تذييل/ترقيم | **لا يوجد** |

## طريقة التشغيل

### المتطلبات
- **Node.js 16** أو أحدث
- خط **Traditional Arabic** مثبَّت على النظام (خط متصل)

### الخطوات

```bash
# من مجلد المشروع
npm install     # اختياري: يمكن الاكتفاء بـ node_modules الموجود في المجلد الأعلى
npm run build   # أو: node create-workbook.js
```

أو عبر سكربت البناء الموحّد من المجلد الأعلى:
```bash
build-all.bat
```

فتح الناتج:
```bash
# Windows
start Arabic_Syllable_Analysis_Workbook.docx
# macOS
open Arabic_Syllable_Analysis_Workbook.docx
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

## ملاحظات تقنية

### بنية الشيفرة
| الدالة | المسؤولية |
|--------|-----------|
| `loadContent` | قراءة ملفي البيانات وفحص كل حقل قبل البناء |
| `pageProperties` | خصائص الصفحة الموحّدة (A5 + RTL + الإطار) |
| `createCoverPage` | صفحة الغلاف الممتلئة، أو `null` إن لم توجد صورة |
| `resolveRelativePath` | أول مسار موجود من قائمة مسارات |
| `getPngDimensions` | قراءة أبعاد الصورة من ترويسة PNG |
| `createSyllableExercise` | تمرين تحليل المقطعين لحرف واحد |
| `createWorkbook` | نقطة الدخول: بناء المستند وحفظه |

الدوال `createDisplayPage` و`createPracticePage` و`transposeGrid` و
`getHarakaBaseLetter` موجودة في هذا الملف لكنها **لا تعمل** في هذا المشروع،
لأن `sections` لا تحتوي `letters`. وهي باقية عمداً للحفاظ على تطابق
السكربتين بايت ببايت.

## ملفات البيانات

كل محتوى الكتاب موجود في `assets/`، و`create-workbook.js` لا يحتوي أي نص أو حرف أو كلمة.

### `assets/letters.json`

| الحقل | الوصف |
|-------|-------|
| `book` | بيانات الملف: `creator` و`title` و`description` |
| `sections` | الأقسام التي يبنيها هذا المشروع (`syllable` هنا) |
| `cover` | `file` للعرض + `candidates` قائمة مسارات تُبحث بالترتيب |
| `syllableExcludedLetters` | الحروف المستثناة من تدريبات المقاطع الصوتية |
| `letters` | قائمة الحروف مرتّبة: `letter` و`name` |

### `assets/harakat.json`

| الحقل | الوصف |
|-------|-------|
| `harakat` | الحركات الأربع: `name` و`symbol` و`longVowel` و`longSymbol` و`description` |
| `baseLetterOverrides` | قواعد استبدال الحرف قبل الحركة، مثل `أ` + كسرة ← `إ` |
| `displayOrder` | ترتيب عرض الرموز في الجداول، و`syllableVowels` المستخدمة في هذا القسم |
| `texts` | نصوص الكتاب: عناوين الصفحات وعناوين الأعمدة وحرف الهمزة |

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

## الترخيص

الكود: MIT

## الإصدار

**1.0.0** — آخر تحديث: **2026-09-29**

---

## English

A generator for a **syllable-analysis practice workbook**. Built from scratch as a
right-to-left **DOCX** (font *Traditional Arabic*), then converted to **PDF** with Word.

This project is the *syllable* half of the original harakat book; the letter and
long-vowel pages stayed in [`Arabic_Harakat_Workbook`](../Arabic_Harakat_Workbook).

### Output

| File | Description |
|------|-------------|
| `Arabic_Syllable_Analysis_Workbook.docx` | final book, written inside the project folder |
| `Arabic_Syllable_Analysis_Workbook.pdf` | PDF version (via the root `build-all.bat`) |

### Page count

| Section | Count |
|---------|-------|
| Cover page | 1 |
| Section title page | 1 |
| Syllable exercises (27 letters ÷ 4 per page) | 7 |
| **Total** | **9 pages** |

### Structure

- **Cover:** full-bleed A5 (skipped with a warning if absent).
- **Exercises:** **4 exercises per page** (one letter each). Each shows two extended syllables above six empty boxes (2 × 3). The long vowels (ا، و، ي) and the hamza (أ) are excluded, leaving **27 letters**.
- The built section is chosen by `sections` in `assets/letters.json` (`["syllable"]` here).
- `create-workbook.js` is **byte-identical** to the one in `Arabic_Harakat_Workbook`; only `assets/letters.json` differs.

### Requirements & run

- Node.js 16+
- System font: **Traditional Arabic**

```bash
npm install      # optional: the root node_modules is shared
npm run build    # or: node create-workbook.js
```

### Customize

- Built sections: `assets/letters.json` → `sections`
- Excluded letters: `syllableExcludedLetters`
- Harakat and labels: `assets/harakat.json`
- Cover image: replace `assets/cover-arabic.png`

### License

Code: MIT.
