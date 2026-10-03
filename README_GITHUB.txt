راهنمای ساخت DoustanTransport.exe با GitHub Actions

۱) آماده‌سازی مخزن
   • فایل ZIP را از حالت فشرده خارج کنید.
   • وارد پوشهٔ DoustanTransport شوید.
   • همهٔ محتویات داخل این پوشه (از جمله app.py، assets، requirements.txt و پوشهٔ .github) را در مخزن GitHub بارگذاری کنید.
     نکته: پوشهٔ مخفی .github را هم حتماً بارگذاری کنید؛ مسیر باید دقیقاً .github/workflows/build.yml باشد.

۲) شروع ساخت خودکار
   • در GitHub مخزن را باز کنید و زبانهٔ Actions را انتخاب کنید.
   • اجرای «Build DoustanTransport for Windows» را باز کنید.
   • اجرای موفق با علامت سبز یعنی فایل اجرایی ویندوز آماده است. پس از بارگذاری/ثبت تغییرات روی main یا master نیز ساخت خودکار آغاز می‌شود.

۳) دانلود مستقیم از Artifact
   • داخل اجرای موفق، تا بخش Artifacts پایین بروید.
   • روی DoustanTransport-Windows کلیک کنید؛ ZIP دانلودشده شامل DoustanTransport.exe است.
   • ZIP را باز و فایل EXE را استخراج کنید.

۴) انتشار در صفحهٔ Releases (اختیاری)
   • برای ساخت Release، یک Tag با نامی مانند v1.0.0 بسازید و آن را به GitHub push کنید.
   • اجرای خودکار، فایل DoustanTransport.exe را به Release همان Tag پیوست می‌کند؛ لینک مستقیم دانلود از صفحهٔ Releases قابل دریافت است.

نمودار مسیر:
  📤 بارگذاری فایل‌ها  →  🪟 ساخت روی Windows  →  📦 Artifact قابل دانلود
                                      └── Tag v* → 🏷️ فایل پیوست Release

توجه: مخزن باید عمومی باشد یا کاربر مجوز دسترسی داشته باشد. برای Artifact وارد حساب GitHub شوید. این بسته فایل‌های ورودی اکسل واقعی را شامل نمی‌شود.
