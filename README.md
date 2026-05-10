# دار الدواء — منصة التحليلات التنفيذية 2026

لوحة قيادة تفاعلية لتحليل بيانات المبيعات والزيارات الميدانية، مبنية بـ HTML/CSS/JS خالص.

## 🚀 الميزات

- **رفع ملف Excel** مباشرة من جهازك
- **ربط GitHub** — تحميل ملف Excel من رابط Raw مباشر
- **فلاتر متقدمة**: فريق، منطقة، تخصص، مندوب، صنف، نطاق زمني
- **مؤشرات KPI** مع أنيميشن عداد
- **مخططات بيانية**: شهري (Bar) + مساهمة الفرق (Doughnut)
- **رؤى SWOT** تلقائية
- **جدول زيارات** مع بحث وشارات أداء
- **تبديل القيمة/الكمية**
- **الوضع الليلي** كامل

## 📁 هيكل الملفات

```
dar-aldawa-dashboard/
├── index.html
├── assets/
│   ├── css/
│   │   └── styles.css
│   └── js/
│       ├── data.js      ← قراءة Excel + معالجة البيانات
│       └── app.js       ← منطق الواجهة والمخططات
├── data/
│   └── sample.xlsx      ← ملف بيانات تجريبي
└── README.md
```

## 📊 هيكل ملف Excel المطلوب

الشيت الأول يجب أن يحتوي على الأعمدة التالية (يمكن أي ترتيب):

| العمود | الاسم المتوقع | ملاحظة |
|--------|--------------|--------|
| المندوب | المندوب / Rep | إلزامي |
| الفريق | الفريق / Team | إلزامي |
| المنطقة | المنطقة / Area | اختياري |
| التخصص | التخصص / Specialty | اختياري |
| الصنف | الصنف / Item | اختياري |
| التاريخ | التاريخ / Date | اختياري |
| القيمة | القيمة / Value | المبيعات بالدينار |
| الكمية | الكمية / Qty | عدد الوحدات |
| الزيارات | الزيارات / Visits | عدد الزيارات |
| الهدف | الهدف / Target | الهدف البيعي |

## 🔗 ربط GitHub (تحديث تلقائي)

1. ارفع ملف Excel إلى مستودعك على GitHub
2. افتح الملف في GitHub → اضغط **Raw**
3. انسخ الرابط (يبدأ بـ `https://raw.githubusercontent.com/...`)
4. في اللوحة: اضغط **GitHub Sync** → الصق الرابط → تحميل

## 🌐 النشر على GitHub Pages

```bash
git init
git add .
git commit -m "initial: dar aldawa dashboard"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/dar-aldawa-dashboard.git
git push -u origin main
```

ثم في إعدادات المستودع: **Settings → Pages → Branch: main → Save**

سيصبح الرابط:
`https://YOUR_USERNAME.github.io/dar-aldawa-dashboard/`

## 🛠 تخصيص أسماء الأعمدة

في `assets/js/data.js`، عدّل `COLUMN_MAP` لمطابقة أسماء أعمدة ملفك:

```js
const COLUMN_MAP = {
  rep:    ['المندوب', 'Rep'],      // اسم العمود في ملف Excel
  team:   ['الفريق', 'Team'],
  value:  ['المبيعات', 'Value'],
  // ...
};
```

---
تم التطوير بواسطة **Basel Ghanem** | دار الدواء © 2026
