@echo off
rem ============================================================
rem  بناء EXE — نظام تفريغ البصمات (الواجهة الحديثة v2.0)
rem  يتطلب Python مثبتاً + إنترنت لأول مرة فقط
rem ============================================================
chcp 65001 >nul
cd /d "%~dp0"

echo [1/3] تثبيت المتطلبات...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install pyinstaller

echo [2/3] بناء الملف التنفيذي (يتضمن واجهة الويب داخل EXE)...
python -m PyInstaller --onefile --windowed --name "AttendanceDump" ^
  --collect-all openpyxl ^
  --collect-all webview ^
  --add-data "web;web" ^
  desktop_app.py

echo [3/3] انتهى.
echo الملف في مجلد: dist\AttendanceDump.exe
echo انسخ settings.json بجانب الـ EXE قبل التشغيل (أو دعه ينشئه تلقائياً).
echo.
echo ملاحظة: يحتاج ويندوز 10/11 مكوّن Edge WebView2 (مثبت مسبقاً في الأغلب).
echo إذا لم تفتح النافذة، شغّل:  AttendanceDump.exe --browser
pause
