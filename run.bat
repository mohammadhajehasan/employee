@echo off
rem تشغيل مباشر بالواجهة الحديثة (نافذة أصلية — بديل تلقائي: المتصفح)
chcp 65001 >nul
cd /d "%~dp0"
python -m pip install -r requirements.txt >nul 2>&1
python desktop_app.py
pause
