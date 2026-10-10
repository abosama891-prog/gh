# نظام مراقبة سيارات شركة أبو غالي

تطبيق ويب لمتابعة أسطول سيارات شركة أبو غالي وسجلات الوقود والمصاريف والعدادات.

- **المستودع:** https://github.com/abosama891-prog/gh
- **الموقع:** https://aboghaly-cars.web.app/
- **التقنيات:** React + TypeScript + Vite + Tailwind + Firebase

## التشغيل

يتطلب المشروع Node.js وnpm.

```powershell
npm install
npm run dev -- --host 127.0.0.1 --port 3001
```

افتح `http://127.0.0.1:3001/`.

## حساب التجربة

- المدير: `admin` / `admin123`

يمكن للمدير إضافة المستخدمين وتحديد أدوارهم من التطبيق. تُحفظ بيانات السيارات والحسابات محليًا في المتصفح، ولا يوجد خادم خلفي أو مزامنة بين الأجهزة.

## البناء

```powershell
npm run build
```

## النشر على Google Firebase Hosting

1. أنشئ مشروع Firebase من حساب Google الذي تريد استخدامه، أو اختر مشروعًا موجودًا.
2. من مجلد المشروع، سجّل الدخول واختر الحساب الآخر في نافذة المتصفح:

```powershell
npx firebase-tools login
npx firebase-tools use --add
```

3. ابنِ التطبيق وانشره:

```powershell
npm run build
npx firebase-tools deploy --only hosting
```

## قاعدة البيانات Firestore

1. من Firebase Console فعّل Firestore Database واختر وضع الإنتاج.
2. من Authentication فعّل مزود Anonymous.
3. انسخ إعدادات تطبيق الويب إلى ملف `.env` باسم المتغيرات الموجودة في `.env.example`.
4. انشر القواعد والاستضافة:

```powershell
npx firebase-tools deploy --only firestore:rules,hosting
```

عند وجود إعدادات Firebase، تُحفظ بيانات السيارات وسجل الوقود وسعر التموين في Firestore داخل المستند `fleet/company`. إذا لم تُضف الإعدادات، يستمر التطبيق في العمل محليًا باستخدام `localStorage`.

إعداد Firebase يعيد توجيه جميع مسارات التطبيق إلى صفحة React، ويمنع تخزين نسخة قديمة من التطبيق أو عامل الخدمة.

> ملاحظة: بيانات الأسطول والحسابات محفوظة محليًا في المتصفح، ولا تُرفع إلى Firebase ولا تتشارك بين المستخدمين أو الأجهزة.