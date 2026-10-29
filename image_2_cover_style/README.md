<div dir="rtl" align="right">

# images-to-book — كتاب الصور بأسلوب الأغلفة

أداة مستقلة تحوّل كل صورة في مجلد `images/` إلى **صفحة كاملة واحدة** بمقاس **A5**،
بأسلوب **صفحات الأغلفة** (full-bleed) نفسه المستخدم في أغلفة
`Arabic_Alphabet_Workbook` و`Arabic_Harakat_Workbook` و`Arabic_Stories_Workbook`:
صورة تملأ الورقة من الحافة إلى الحافة، بلا هوامش ولا إطار ولا تذييل ولا ترقيم.

> **مهم:** هذا المشروع **مستقل** ولا يتصل بسكربتات الجذر في `build-all.bat` —
> لا يُبنى ضمن الكتب السبعة. لكنه **يستعمل** `convert-one.ps1` من الجذر لتحويل
> DOCX إلى PDF، ويستعمل `node_modules` المشترك في الجذر، فلا يحتاج تثبيتاً منفصلاً.

## طريقة الاستعمال

ضع صورك في مجلد `images/` ثم شغّل:

```bat
images-to-book.bat
```

لا يوجد أي وسيط للمجلدات: مجلد الإدخال **ثابت** هو `images/` المجاور للسكربت،
ومجلد الإخراج **ثابت** هو المجلد نفسه. لذلك كل ما عليك فعله هو نسخ الصور وتشغيل
الملف.

### الخيارات

| الأمر | النتيجة |
|-------|---------|
| `images-to-book.bat` | A5، ويُنتج DOCX و PDF |
| `images-to-book.bat --size a4` | مقاس A4 بدل A5 |
| `images-to-book.bat --no-pdf` | DOCX فقط، بلا حاجة إلى Word |

> **قبل PDF:** أغلق Microsoft Word، وإلا فشل التحويل برسالة
> `The process cannot access the file...`. وإن أردت DOCX فقط، استخدم `--no-pdf`.

### مثال سريع

```bat
copy "C:\My Covers\*.png" images\
images-to-book.bat --size a4
```

ينشئ هذا كتاباً بمقاس A4 يحتوي كل صورة في `images/`، صفحة لكل صورة، ثم يحوّله إلى
PDF. استخدم `--no-pdf` إن لم يكن Word مثبَّتاً.

## ما الذي يدخل الكتاب

| القاعدة | التفصيل |
|---------|---------|
| الصيغ | `.png` و `.jpg` و `.jpeg` (غير حساسة لحالة الأحرف) |
| المجلدات الفرعية | **تُضمَّن** تلقائياً، فمسار `2026/ملف.png` يصبح صفحة باسم `2026/ملف.png` |
| الترتيب | ترتيب طبيعي على المسار النسبي، فيأتي `1-` قبل `2-` قبل `10-` |
| الملف التالف | يُطبع `SKIP` مع السبب، وتُبنى بقية الصفحات |
| SVG والملفات الأخرى | **تُتجاهل** بلا تحذير |

> أبعاد كل صورة تُقرأ من ترويستها مباشرة (بايتات `IHDR` في PNG، وأول مقطع `SOF`
> في JPEG)، فلا حاجة إلى أي مكتبة صور إضافية.

## بنية كل صفحة

كل صورة **قسم مستقل** (section) بصفحة واحدة، فلا تنتقل صفحة إلى أخرى:

| العنصر | القيمة |
|--------|--------|
| المقاس | A5 (5.83 × 8.27 بوصة) أو A4 (8.27 × 11.69) |
| الهوامش | **صفر** من الجهات الأربع |
| الإطار | **بدون** `PAGE_FRAME` |
| التذييل والترقيم | **بدون** |
| المحتوى | صورة واحدة `ImageRun` في فقرة واحدة في الوسط |

وهذه هي نفس المعادلة المستخدمة في أغلفة `create-workbook.js`: `convertInchesToTwip` للقسم، وتقسيم الناتج على `15` لأن `1` بكسل عند `96` dpi يساوي `15` twip. ولهذا تخرج صورة `1024x1536` بمقاس `560x840` بكسل تماماً كما تخرج في الأغلفة.

## كيف تُوضَع الصورة في الصفحة

الأداة تختار تلقائياً بين وضعين حسب الصورة نفسها، ولا يوجد خيار لتغيير ذلك:

| الوضع | الشرط | ما يحدث |
|-------|-------|---------|
| **full-bleed** | نسبة القصّ ≤ **8%** | تُكبَّر الصورة عمداً لتغطي الورقة كلها، ويخرج ما زاد عن الصفحة على الحواف الأربع |
| **محتوى كامل** | نسبة القصّ > 8% | تُصغَّر الصورة حتى تظهر كاملة، وتُوسَّط رأسياً في الفراغ المتبقي |

### لماذا العتبة على نسبة القصّ لا على نسبة الأبعاد

هذا هو الجزء الوحيد في الأداة الذي يبدو غريباً، وهو مقصود:

- صورة الغلاف `1024x1536` تبعد بنسبة **5.4%** عن نسبة صفحة A5، والأغلفة في هذا
  المشروع **تقطع** هذه النسبة عمداً عند التحويل إلى bleed. فلو كان الشرط «هل
  النسبتان متقاربتان» لَرُفضت هذه الصورة — أي ستُرفض الصور التي وُجدت الأداة لأجلها.
- صورة `1536x1024` عرضية على الصفحة نفسها ستخسر **53%** من عرضها لو قُصّت، وهذا كثير
  جداً، فيجب احتواء الصورة كاملة بدل ذلك.

فالعتبة تُقاس على **كم يُقتطع**، لا على **مدى تقارب النسبتين**، لأن الأول هو السؤال
الذي يهم فعلاً: هل سيُفقد جزء مهم من الصورة؟

### الفراغ في وضع المحتوى الكامل

الورقة 794 بكسل، فإذا كانت الصورة 507 بكسل يتبقى 287، ويذهب نصفها (`143` بكسل)
مسافة قبل الفقرة، فتصير الفراغات أعلى وأسفل متساويين. وصفر المسافة في وضع
full-bleed مقصود: الصورة تملأ الورقة، فلا مكان لمسافة.

## الملفات

```
image_2_cover_style/
├── 📄 images-to-book.bat      # نقطة الدخول: تشغيل بأمر واحد
├── 📄 images-to-book.js       # جمع الصور + قراءة الأبعاد + بناء DOCX
├── 📄 package.json            # بلا تبعيات؛ docx يُحَلّ من node_modules المشترك
├── 📄 README.md               # هذا الملف
├── 📂 images/                 # مجلد الإدخال الثابت
├── 📂 images-to-book.docx     # الناتج (يُنشأ عند التشغيل)
└── 📂 images-to-book.pdf      # الناتج (يتطلب Word)
```

## المتطلبات

| المتطلب | لماذا |
|---------|-------|
| **Node.js** 16 أو أحدث | تشغيل `images-to-book.js` وبناء ملف DOCX |
| **Microsoft Word** | تحويل DOCX إلى PDF عبر `convert-one.ps1` (اختياري مع `--no-pdf`) |
| `docx` في `node_modules` المشترك بالجذر | تثبّته `build-all.bat` تلقائياً، وإن لم يكن موجوداً ثبّته `images-to-book.bat` |

## ملاحظات

- الأداة تختار أفضل وضع لكل صورة على حدة، فيمكن أن يخرج كتاب تخلط فيه صفحات
  full-bleed وصفحات محتوى كامل.
- الأسماء في Windows لا تكفي للترتيب الرقمي (`10-` يسبق `2-`)، ولهذا الترتيب
  هنا **طبيعي** لا أبجدي.
- إن غيّرت `MAX_BLEED_CROP` في `images-to-book.js` فأنت تغيّر نقطة التحوّل بين
  الوضعين لكل الصور معاً.

</div>

---

## English

A standalone helper that turns every image in `images/` into one **full-bleed A5 (or A4)
page** — the same cover style used by the workbooks (image edge to edge, no margins,
frame, footer, or page numbers) — and writes `images-to-book.docx` / `.pdf`.

> Standalone from `build-all.bat`: it is not one of the seven books, but it does use the
> root `convert-one.ps1` and the shared root `node_modules`.

### Usage

Drop your images in `images/` and run:

```bat
images-to-book.bat
```

| Command | Result |
|---------|--------|
| `images-to-book.bat` | A5, produces DOCX and PDF |
| `images-to-book.bat --size a4` | A4 instead of A5 |
| `images-to-book.bat --no-pdf` | DOCX only, no Word needed |

Input and output folders are **fixed** (`images/` next to the script). Accepted formats:
`.png`, `.jpg`, `.jpeg`; subfolders are included; the order is natural (so `1-` before `2-`
before `10-`); broken files are skipped and other file types ignored.

### Page fitting

Each image is auto-fitted: if the crop ratio is ≤ **8%** it is **full-bleed**; otherwise it
is **contained** (scaled to fit and centered). The threshold is on *how much is cropped*,
not on how close the aspect ratios are — so `1024×1536` cover images stay full-bleed while
wide `1536×1024` images are contained.

### Requirements

- Node.js 16+
- Microsoft Word (only to produce the PDF; use `--no-pdf` otherwise)
- `docx` in the shared root `node_modules`

### License

Code: MIT.
