# دليل بناء وتثبيت APK — نظام الحضور

---

## متطلبات النظام (يجب تثبيتها قبل البدء)

| الأداة | الإصدار المطلوب | رابط التحميل |
|--------|-----------------|-------------|
| Node.js | 18 أو أحدث | https://nodejs.org |
| Android Studio | أحدث إصدار مستقر | https://developer.android.com/studio |
| Java JDK | 17 أو أحدث (مضمّن في Android Studio) | — |
| Android SDK | API 34 (مُثبَّت من Android Studio) | — |

> **ملاحظة:** بعد تثبيت Android Studio، افتحه مرة واحدة على الأقل حتى يُكمل تثبيت SDK tools.
> تأكد أن متغير البيئة `ANDROID_HOME` مضبوط. يُضبط تلقائياً عند تثبيت Android Studio.

---

## الخطوة 0 — إعداد GAS Web App قبل أي شيء

**هذه الخطوة إلزامية. بدونها لن يعمل APK.**

1. افتح مشروع Google Apps Script الخاص بك.
2. تأكد أنك أضفت دالة `doPost()` الجديدة (موجودة في نهاية `Attendance.Js`).
3. في محرر GAS، اذهب إلى:
   **Deploy → New deployment**
4. اختر النوع: **Web app**
5. اضبط الإعدادات:
   - Execute as: **Me**
   - Who has access: **Anyone**  ← مهم جداً، وإلا ستفشل طلبات APK
6. انسخ رابط الـ Deployment URL. يبدو هكذا:
   ```
   https://script.google.com/macros/s/AKfycb.../exec
   ```
7. افتح الملف:
   ```
   www\index.html
   ```
8. ابحث عن هذا السطر في أعلى الـ `<script>`:
   ```javascript
   const GAS_WEBAPP_URL = 'https://script.google.com/macros/s/REPLACE_WITH_YOUR_DEPLOYMENT_ID/exec';
   ```
9. **استبدل** `REPLACE_WITH_YOUR_DEPLOYMENT_ID` بالرابط الكامل الذي نسخته.

---

## الخطوة 1 — تثبيت الحزم

افتح Terminal أو PowerShell في مجلد المشروع:

```powershell
cd "e:\كنيسة\شهادات الكورسات\Attendance"
npm install
```

**ما يحدث:** يُثبِّت Capacitor core، Android platform، Filesystem، Share، وBarcodeScanner plugins.

---

## الخطوة 2 — إنشاء مشروع Android

```powershell
npx cap add android
```

**ما يحدث:** ينشئ مجلد `android\` كامل يحتوي على مشروع Gradle الأصلي.

> **تحذير:** نفّذ هذا الأمر مرة واحدة فقط. إذا كان مجلد `android\` موجوداً بالفعل، تخطَّ هذه الخطوة.

---

## الخطوة 3 — نسخ ملفات www إلى مشروع Android

```powershell
npx cap sync android
```

**ما يحدث:**
- ينسخ محتوى `www\` إلى `android\app\src\main\assets\public\`
- يُحدِّث إعدادات plugins في مشروع Android

> **مهم:** نفّذ هذا الأمر في كل مرة تُعدِّل فيها `www\index.html`.

---

## الخطوة 4 — إضافة صلاحيات Android يدوياً

بعد تنفيذ `npx cap add android`، افتح الملف:
```
android\app\src\main\AndroidManifest.xml
```

أضف هذه الأسطر **داخل** عنصر `<manifest>` وقبل `<application>`:

```xml
<!-- صلاحية الكاميرا — مطلوبة لمسح QR -->
<uses-permission android:name="android.permission.CAMERA" />
<uses-feature android:name="android.hardware.camera" android:required="false" />
<uses-feature android:name="android.hardware.camera.autofocus" android:required="false" />

<!-- صلاحية الإنترنت — مطلوبة للاتصال بـ Google Apps Script -->
<uses-permission android:name="android.permission.INTERNET" />

<!-- صلاحية الكتابة في التخزين — مطلوبة لحفظ PDF في Downloads -->
<uses-permission android:name="android.permission.WRITE_EXTERNAL_STORAGE"
    android:maxSdkVersion="28" />
<uses-permission android:name="android.permission.READ_EXTERNAL_STORAGE"
    android:maxSdkVersion="32" />
```

> **ملاحظة:** على Android 10+ (API 29+) لا تحتاج `WRITE_EXTERNAL_STORAGE` للكتابة في
> مجلد Downloads الخاص بالتطبيق، لكن إضافتها مع `maxSdkVersion="28"` لا تضر.

---

## الخطوة 5 — تثبيت BarcodeScanner في مشروع Android

افتح الملف:
```
android\app\build.gradle
```

في قسم `dependencies {}` أضف:

```gradle
implementation "com.github.capacitor-community:barcode-scanner:3.0.3"
```

ثم افتح:
```
android\build.gradle
```

في قسم `allprojects { repositories {} }` تأكد من وجود:

```gradle
maven { url 'https://jitpack.io' }
```

---

## الخطوة 6 — تسجيل BarcodeScanner Plugin في MainActivity

افتح الملف:
```
android\app\src\main\java\com\church\attendance\MainActivity.java
```

استبدله بالكامل بهذا المحتوى:

```java
package com.church.attendance;

import android.os.Bundle;
import com.getcapacitor.BridgeActivity;
import com.capacitorjs.plugins.filesystem.FilesystemPlugin;
import com.capacitorjs.plugins.share.SharePlugin;
import com.capacitor.community.barcodescanner.BarcodeScanner;

public class MainActivity extends BridgeActivity {
    @Override
    public void onCreate(Bundle savedInstanceState) {
        registerPlugin(BarcodeScanner.class);
        registerPlugin(FilesystemPlugin.class);
        registerPlugin(SharePlugin.class);
        super.onCreate(savedInstanceState);
    }
}
```

---

## الخطوة 7 — بناء APK تجريبي (Debug)

### الطريقة الأولى: من Terminal (أسرع)

```powershell
cd "e:\كنيسة\شهادات الكورسات\Attendance\android"
.\gradlew assembleDebug
```

**مكان الـ APK بعد البناء:**
```
android\app\build\outputs\apk\debug\app-debug.apk
```

### الطريقة الثانية: من Android Studio (أسهل)

```powershell
npx cap open android
```

ثم في Android Studio:
- **Build → Build Bundle(s) / APK(s) → Build APK(s)**

---

## الخطوة 8 — تثبيت APK على هاتف Android

### عبر USB (الطريقة الموصى بها للاختبار)

1. على الهاتف: **الإعدادات → خيارات المطور → USB Debugging** (شغّله)
2. وصّل الهاتف بالكمبيوتر عبر USB
3. نفّذ:

```powershell
# تأكد أن الهاتف متصل
adb devices

# ثبّت الـ APK
adb install "e:\كنيسة\شهادات الكورسات\Attendance\android\app\build\outputs\apk\debug\app-debug.apk"
```

### عبر نقل الملف مباشرة

1. انسخ ملف `app-debug.apk` إلى الهاتف
2. افتحه من مدير الملفات
3. اقبل تثبيت التطبيقات من مصادر غير معروفة إذا طُلب منك

---

## الخطوة 9 — اختبار التطبيق

### اختبار الكاميرا و QR

1. افتح التطبيق
2. اختر أي كورس → افتح محاضرة → اضغط **Scan QR**
3. التطبيق سيطلب إذن الكاميرا — اقبله
4. وجّه الكاميرا نحو QR code بالصيغة:
   ```
   AB001--توماس ماجد فطيم عطية
   ```
5. يجب أن يظهر toast: **✓ تم تسجيل الحضور**

### اختبار الشهادة والـ PDF

1. افتح كشف الحضور لشخص لديه شهادة
2. اضغط **عرض الشهادة**
3. انتظر تحميل PDF
4. اضغط زر **المشاركة**
5. تحقق أن اسم الملف في نافذة المشاركة يحتوي على اسم الشخص
   - مثال: `توماس ماجد فطيم عطية.pdf`

### اختبار الاتصال بـ Google Sheets

1. سجّل حضور شخص
2. افتح Google Sheets مباشرة
3. تأكد أن الحضور سُجِّل في العمود الصحيح بالتوقيت الصحيح

---

## الخطوة 10 — بناء APK نهائي (Release)

### إنشاء مفتاح التوقيع (مرة واحدة فقط)

```powershell
keytool -genkey -v -keystore "e:\كنيسة\شهادات الكورسات\Attendance\attendance-release-key.jks" `
  -alias attendance `
  -keyalg RSA `
  -keysize 2048 `
  -validity 10000
```

سيطلب منك كلمة مرور وبيانات المنظمة — احتفظ بها في مكان آمن.

### بناء APK Release

```powershell
cd "e:\كنيسة\شهادات الكورسات\Attendance\android"
.\gradlew assembleRelease
```

### توقيع APK يدوياً

```powershell
jarsigner -verbose -sigalg SHA256withRSA -digestalg SHA-256 `
  -keystore "e:\كنيسة\شهادات الكورسات\Attendance\attendance-release-key.jks" `
  "android\app\build\outputs\apk\release\app-release-unsigned.apk" `
  attendance

zipalign -v 4 `
  "android\app\build\outputs\apk\release\app-release-unsigned.apk" `
  "android\app\build\outputs\apk\release\app-release.apk"
```

---

## تحديث التطبيق (workflow يومي)

في كل مرة تُعدِّل `www\index.html`:

```powershell
cd "e:\كنيسة\شهادات الكورسات\Attendance"
npx cap sync android
cd android
.\gradlew assembleDebug
adb install -r "app\build\outputs\apk\debug\app-debug.apk"
```

الخيار `-r` يعني "استبدال" — لا يحذف بيانات التطبيق.

---

## استكشاف الأخطاء وإصلاحها

### مشكلة: الكاميرا لا تُفتح

**الأسباب المحتملة:**
1. BarcodeScanner Plugin لم يُسجَّل في `MainActivity.java` → راجع الخطوة 6
2. الصلاحية غير موجودة في `AndroidManifest.xml` → راجع الخطوة 4
3. المستخدم رفض الإذن → اطلب منه فتح **إعدادات التطبيق → الأذونات → الكاميرا**

### مشكلة: لا تصل طلبات الشبكة إلى GAS

**الأسباب المحتملة:**
1. `GAS_WEBAPP_URL` في `www\index.html` لم يُعدَّل → راجع الخطوة 0
2. Deployment لم يُنشر بإعداد "Anyone" → أعد النشر
3. في Android 9+ يجب استخدام HTTPS — GAS يستخدمه تلقائياً ✓

**للتصحيح:** شغّل:
```powershell
adb logcat -s "Capacitor" "chromium"
```
وابحث عن أخطاء الشبكة.

### مشكلة: اسم PDF ظهر مشوهاً

تأكد أنك تُنفِّذ `npx cap sync android` بعد كل تعديل على `www\index.html`.
ثم أعد تثبيت APK.

### مشكلة: `npx cap add android` يفشل

```powershell
# تأكد أن Node.js مثبت
node --version   # يجب أن يكون v18+

# تأكد أن ANDROID_HOME مضبوط
echo $env:ANDROID_HOME   # يجب أن يُظهر مسار SDK

# جرب تثبيت Capacitor CLI عالمياً
npm install -g @capacitor/cli@6.1.2
cap add android
```

---

## ملخص الملفات المُعدَّلة

| الملف | النوع | التغيير |
|-------|-------|---------|
| `Attendance.Js` | GAS Backend | أُضيفت دالة `doPost()` في النهاية |
| `www\index.html` | APK Frontend | نسخة معدَّلة للـ APK بثلاثة تغييرات مشروطة |
| `www\stylesheet.css` | APK Assets | نسخة من `Stylesheet.css` للـ APK |
| `package.json` | جديد | تعريف مشروع Capacitor والاعتماديات |
| `capacitor.config.json` | جديد | إعدادات Capacitor |

**الملفات الأصلية التي لم تُمَس:**
- `Index.html` — نسخة GAS تعمل كما كانت بالضبط
- `Stylesheet.css` — لم يتغير
- منطق الحضور في `Attendance.Js` — لم يتغير
- بنية Google Sheets — لم تتغير
- صيغة QR — لم تتغير

---

## قائمة التحقق النهائية

### الكاميرا
- [ ] التطبيق يطلب إذن الكاميرا عند أول تشغيل
- [ ] تفتح كاميرا خلفية للمسح
- [ ] مسح QR يعمل
- [ ] القيمة الممسوحة تصل إلى `handleQrDetected()`

### الحضور
- [ ] يُسجَّل الحضور بنجاح
- [ ] منع التسجيل المكرر يعمل كما كان
- [ ] الإضافة اليدوية تعمل
- [ ] البيانات تظهر في Google Sheets

### الشهادة
- [ ] PDF يُعرض داخل التطبيق
- [ ] زر المشاركة يفتح نافذة مشاركة Android
- [ ] اسم الملف يحتوي على اسم الشخص: `توماس ماجد فطيم عطية.pdf`
- [ ] الاسم العربي غير مشوه
- [ ] الملف يفتح بشكل صحيح في WhatsApp وGoogle Drive
- [ ] زر التنزيل يحفظ الملف في Downloads بالاسم الصحيح

### Google Sheets
- [ ] `getActivities` يُرجع البيانات
- [ ] `recordQrAttendance` يكتب في الخلية الصحيحة
- [ ] `batchSyncAttendance` يعمل بعد انقطاع الإنترنت

### الأوفلاين
- [ ] مسح QR بدون إنترنت يُخزَّن محلياً
- [ ] البانر الأوفلاين يظهر
- [ ] المزامنة تحدث تلقائياً عند عودة الإنترنت
- [ ] السجلات لا تتكرر بعد المزامنة
