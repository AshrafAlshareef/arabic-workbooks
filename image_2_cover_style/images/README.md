<div dir="rtl" align="right">

# مجلد الإدخال

ضع صورك هنا، ثم شغّل `..\images-to-book.bat` (أو `images-to-book.bat` من داخل هذا
المجلد).

- الصيغ المقبولة: `.png` و `.jpg` و `.jpeg`
- المجلدات الفرعية تُضمَّن تلقائياً
- كل صورة تصبح صفحة واحدة بترتيب المسار النسبي

### مثال

```
images/
├── 1-cover.png
├── 2-page.jpg
└── chapter-1/3-photo.png
```

سيدخل الكتاب الصور الثلاث بهذا الترتيب: `1-cover.png` ثم `2-page.jpg` ثم
`chapter-1/3-photo.png`.

تفاصيل الأداة: [`../README.md`](../README.md)

</div>

---

## English

Drop your images here, then run `..\images-to-book.bat`.

- Accepted formats: `.png`, `.jpg`, `.jpeg`
- Subfolders are included automatically
- Each image becomes one page, ordered by its relative path

Tool details: [`../README.md`](../README.md)
