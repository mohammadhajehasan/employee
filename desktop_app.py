# -*- coding: utf-8 -*-
"""
desktop_app.py — نظام تفريغ البصمات: الواجهة الحديثة (الإصدار 2.0)
=================================================================
نافذة سطح مكتب أصلية تعرض واجهة ويب حديثة (HTML/CSS) عبر محرك النظام:
  • على ويندوز: Edge WebView2 (مدمج في ويندوز 10/11) — خفيف وسريع
  • بديل تلقائي: إذا لم تتوفر مكتبة pywebview يفتح التطبيق في المتصفح الافتراضي

التشغيل:
  python desktop_app.py            ← نافذة أصلية (أو متصفح تلقائياً عند عدم توفرها)
  python desktop_app.py --browser  ← إجبار وضع المتصفح

بناء EXE: استخدم build.bat (يتضمن مجلد web داخل الملف التنفيذي)
"""
from __future__ import annotations

import base64
import json
import os
import re
import shutil
import sys
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, quote, urlparse

# ---------- المسارات (تعمل في التطوير وفي EXE) ----------
if getattr(sys, "frozen", False):          # PyInstaller
    APP_DIR = os.path.dirname(sys.executable)
    WEB_DIR = os.path.join(getattr(sys, "_MEIPASS", APP_DIR), "web")
    _CODE_DIR = getattr(sys, "_MEIPASS", APP_DIR)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))
    WEB_DIR = os.path.join(APP_DIR, "web")
    _CODE_DIR = APP_DIR

sys.path.insert(0, _CODE_DIR)

SETTINGS_PATH = os.path.join(APP_DIR, "settings.json")
OUTPUT_DIR = os.path.join(APP_DIR, "output")
UPLOAD_DIR = os.path.join(APP_DIR, "uploads")
VERSION = "2.5"
MAX_UPLOAD = 40 * 1024 * 1024  # 40MB

os.makedirs(OUTPUT_DIR, exist_ok=True)

MONTH_AR = {1: "يناير", 2: "فبراير", 3: "مارس", 4: "أبريل", 5: "مايو", 6: "يونيو",
            7: "يوليو", 8: "أغسطس", 9: "سبتمبر", 10: "أكتوبر", 11: "نوفمبر", 12: "ديسمبر"}


# ---------- محرك المعالجة (نفس نواة v1) ----------
from app import run_pipeline                      # noqa: E402
from att_parser import parse_attlog               # noqa: E402
from processor import Processor                   # noqa: E402
from excel_writer import export_csv               # noqa: E402
from quality_report import build_report, save_report  # noqa: E402


def load_settings() -> dict:
    if not os.path.exists(SETTINGS_PATH):
        return {"version": 1, "employees": [], "notes": {}}
    try:
        with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {"version": 1, "employees": [], "notes": {}}


class Api:
    """جسر موحّد تستدعيه الواجهة (نفس التوابع تعمل في وضعي النافذة والمتصفح)."""

    def __init__(self):
        self.settings = load_settings()

    # ----- أدوات داخلية -----
    def _save(self):
        with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
            json.dump(self.settings, f, ensure_ascii=False, indent=2)

    def _ym(self, p):
        return int(p.get("year", 2026)), int(p.get("month", 9))

    def _default_out(self, kind: str, year: int, month: int) -> str:
        if kind == "csv":
            return os.path.join(OUTPUT_DIR, f"تحقق_{year}_{month:02d}.csv")
        if kind == "quality":
            return os.path.join(OUTPUT_DIR, f"تقرير_الجودة_{year}_{month:02d}.txt")
        return os.path.join(OUTPUT_DIR, f"دوام_{MONTH_AR.get(month, month)}_{year}.xlsx")

    # ----- معلومات -----
    def app_info(self, p=None):
        mode = "نافذة سطح مكتب (WebView)" if _RUN_MODE == "webview" else "متصفح"
        return {"ok": True, "version": VERSION, "mode": mode,
                "settings_path": SETTINGS_PATH, "output_dir": OUTPUT_DIR,
                "platform": sys.platform}

    def get_state(self, p=None):
        return {"ok": True, "settings": self.settings, "output_dir": OUTPUT_DIR}

    def save_settings(self, p):
        s = p.get("settings")
        if not isinstance(s, dict):
            return {"ok": False, "error": "إعدادات غير صالحة"}
        s.setdefault("employees", [])
        s.setdefault("notes", {})
        self.settings = s
        self._save()
        return {"ok": True}

    # ----- حوارات الملفات -----
    def select_file(self, p):
        """فتح ملف من قرص المستخدم (وضع النافذة الأصلية فقط)."""
        kind = p.get("kind", "attlog")
        if _RUN_MODE != "webview":
            return {"ok": False, "error": "الحوار الأصلي متاح في وضع النافذة فقط"}
        import webview
        types = {
            # «كل الملفات» أولاً حتى لا يظهر أي ملف للمستخدم رمادياً في الحوار
            "attlog": ["كل الملفات (*.*)", "ملفات النص (*.txt;*.log)"],
            "template": ["ملفات Excel (*.xlsx;*.xlsm;*.xls)", "كل الملفات (*.*)"],
        }.get(kind, ["كل الملفات (*.*)"])
        result = webview.windows[0].create_file_dialog(
            webview.OPEN_DIALOG, allow_multiple=False, file_types=types)
        if result:
            path = result[0] if isinstance(result, (list, tuple)) else result
            return {"ok": True, "path": path, "name": os.path.basename(str(path))}
        return {"ok": False, "error": "cancelled"}

    def save_dialog(self, p):
        """حوار حفظ باسم افتراضي (وضع النافذة الأصلية فقط)."""
        if _RUN_MODE != "webview":
            return {"ok": False, "error": "cancelled"}
        import webview
        kinds = {
            "xlsx": ["ملفات Excel (*.xlsx)"],
            "csv": ["ملفات CSV (*.csv)"],
            "quality": ["ملفات النص (*.txt)"],
        }
        result = webview.windows[0].create_file_dialog(
            webview.SAVE_DIALOG,
            save_filename=p.get("default_name", "دوام.xlsx"),
            file_types=kinds.get(p.get("kind", "xlsx"), ["كل الملفات (*.*)"]))
        if result:
            path = result[0] if isinstance(result, (list, tuple)) else result
            return {"ok": True, "path": str(path)}
        return {"ok": False, "error": "cancelled"}

    def ingest_file(self, p):
        """استقبال ملف من الواجهة (base64) وحفظه في uploads — مع تحقق واضح."""
        try:
            os.makedirs(UPLOAD_DIR, exist_ok=True)
            b64s = str(p.get("b64") or "")
            if b64s.startswith("data:") and "," in b64s:      # احتياط: لو أُرسل dataURL كاملاً
                b64s = b64s.split(",", 1)[1]
            b64s = re.sub(r"\s+", "", b64s)
            if not b64s:
                return {"ok": False, "error": "محتوى الملف فارغ — أعد اختيار الملف وتأكد من اكتماله"}
            try:
                data = base64.b64decode(b64s)
            except Exception:
                return {"ok": False, "error": "تعذر فك ترميز الملف — أعد اختياره من جديد"}
            if not data:
                return {"ok": False, "error": "الملف المختار فارغ (0 بايت)"}
            if len(data) > MAX_UPLOAD:
                return {"ok": False, "error": "الملف أكبر من الحد المسموح (40MB)"}
            kind = p.get("kind", "")
            if kind == "template":
                # ملف xlsx حقيقي يبدأ بتوقيع ZIP «PK» — يكشف xls القديم والملفات التالفة
                if len(data) < 100 or data[:2] != b"PK":
                    return {"ok": False, "error":
                            "هذا الملف ليس Excel بصيغة xlsx الحديثة (قد يكون xls قديماً أو تالفاً). "
                            "الحل: افتحه في Excel ثم ملف ← حفظ باسم ← «Excel Workbook (*.xlsx)» "
                            "وأعد اختياره."}
            if kind == "attlog" and data[:2] == b"PK":
                return {"ok": False, "error":
                        "يبدو أنك اخترت ملف Excel لملف البصمات — ملف البصمات attlog ملف نصي "
                        "(TXT) يصدّره جهاز البصمة. اختر الملف النصي الصحيح."}
            raw_name = os.path.basename(str(p.get("name") or "file"))
            safe = "".join(c for c in raw_name if c not in '\\/:*?"<>|').strip() or "file"
            path = os.path.join(UPLOAD_DIR, safe)
            with open(path, "wb") as f:
                f.write(data)
            return {"ok": True, "path": path, "name": safe, "size": len(data)}
        except Exception as e:
            return {"ok": False, "error": f"فشل حفظ الملف المرفوع: {type(e).__name__}: {e}"}

    def download_file(self, p):
        """«حفظ باسم» للملفات المولّدة (وضع النافذة الأصلية): حوار + نسخ."""
        src = p.get("path", "")
        if not src or not os.path.isfile(src):
            return {"ok": False, "error": "الملف غير موجود"}
        if _RUN_MODE != "webview":
            return {"ok": False, "error": "cancelled"}
        import webview
        result = webview.windows[0].create_file_dialog(
            webview.SAVE_DIALOG,
            save_filename=os.path.basename(src) or "output.xlsx",
            file_types=["كل الملفات (*.*)"])
        if not result:
            return {"ok": False, "error": "cancelled"}
        dest = result[0] if isinstance(result, (list, tuple)) else result
        try:
            shutil.copyfile(src, str(dest))
            return {"ok": True, "path": str(dest)}
        except Exception as e:
            return {"ok": False, "error": f"تعذر الحفظ: {e}"}

    def read_output_b64(self, p):
        """قراءة ملف من مجلد المخرجات (اسم فقط، بدون مسارات) كـ base64."""
        name = os.path.basename(p.get("name", ""))
        path = os.path.join(OUTPUT_DIR, name)
        if not name or not os.path.isfile(path):
            return {"ok": False, "error": "الملف غير موجود"}
        with open(path, "rb") as f:
            data = f.read()
        return {"ok": True, "name": name, "size": len(data),
                "b64": base64.b64encode(data).decode("ascii")}

    def template_info(self, p):
        """تحليل قالب مختار: عدد الأسماء، أيام الشهر المكتشفة، وضع التوليد المتوقع."""
        from template_writer import analyze_template
        path = p.get("path") or self.settings.get("template_path") or ""
        if not path or not os.path.exists(path):
            return {"ok": False, "error": "لم يُحدد ملف قالب صالح"}
        try:
            year, month = self._ym(p)
            return analyze_template(path, year, month)
        except Exception as e:
            return {"ok": False, "error": f"تعذر تحليل القالب: {type(e).__name__}: {e}"}

    def open_path(self, p):
        path = p.get("path", "")
        if path == "OUTPUT_DIR":
            path = OUTPUT_DIR
        if not path or not os.path.exists(path):
            return {"ok": False, "error": "المسار غير موجود"}
        try:
            if sys.platform == "win32":
                os.startfile(path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                os.system(f'open "{path}" &')
            else:
                os.system(f'xdg-open "{path}" &')
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    # ----- المعاينة والتوليد -----
    def preview(self, p):
        attlog = p.get("attlog", "")
        year, month = self._ym(p)
        if not attlog or not os.path.exists(attlog):
            return {"ok": False, "error": "مسار attlog غير صحيح"}
        pr = parse_attlog(attlog, int(self.settings.get("dedup_seconds", 120)))
        proc = Processor(self.settings)
        att = proc.process(pr, year, month)
        summ = proc.summary(att)
        employees = []
        for emp_id, s in sorted(summ["per_employee"].items()):
            employees.append({"id": emp_id, "name": s["name"] or f"#{emp_id}",
                              "days": s["days"], "problems": s["problems"]})
        return {"ok": True, "stats": {
            "records": len(pr.records),
            "corrupt": len(pr.corrupt_lines),
            "duplicates": pr.duplicates_removed,
            "encoding": pr.encoding,
            "total_lines": pr.total_lines,
            "unmapped": {str(k): v for k, v in summ["unmapped_ids"].items()},
            "employees": employees,
        }}

    def generate(self, p):
        attlog = p.get("attlog", "")
        year, month = self._ym(p)
        kind = p.get("kind", "xlsx")
        if not attlog or not os.path.exists(attlog):
            return {"ok": False, "error": "مسار attlog غير صحيح"}
        out = p.get("out") or self._default_out("xlsx" if kind == "xlsx" else kind, year, month)
        os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
        logs = []

        # فحص مسبع: إن لم يُفهم أي سطر نوقف التوليد فوراً برسالة واضحة
        # بدل إنتاج ملف Excel فارغ بصمت (شكوى المستخدم: "الملف الناتج لا يجد به بيانات")
        pre = parse_attlog(attlog, int(self.settings.get("dedup_seconds", 120)))
        fmt_desc = ("بأسماء الموظفين" if pre.has_names else
                    "بدون أسماء (صيغة الجهاز: رقم + تاريخ + حالات)")
        if pre.records:
            logs.append(f"صيغة الملف المكتشفة: {fmt_desc} | تاريخ: {pre.date_style or 'غير محددة'}")
        if not pre.records:
            try:
                with open(attlog, "rb") as fh:
                    raw_head = fh.read(400)
                sample = raw_head.decode("utf-8", errors="replace")[:200]
            except Exception:
                sample = ""
            sample = sample.replace("\r", " ").replace("\n", " ⏎ ")
            reasons = "\n".join(f"   • سطر {ln}: {err}" for ln, _tx, err in pre.corrupt_lines[:3])
            return {"ok": False, "error": (
                "فشل التفريغ: لم أفهم أي سطر من الملف النصي "
                f"({pre.total_lines} سطر) — لن أولّد ملفاً فارغاً.\n"
                f"الترميز المكتشف: {pre.encoding}\n"
                f"{('أسباب الاستبعاد:\n' + reasons) if reasons else 'الملف لا يحتوي أسطراً قابلة للقراءة.'}\n"
                f"أول ما في ملفك: «{sample}»\n"
                "تأكد أن الملف هو attlog.txt الصادر من جهاز البصمة (سطر لكل بصمة: رقم الموظف ثم التاريخ والوقت). "
                "إن بقي الخطأ أرسل أول 3 أسطر من الملف لإضافة دعم صيغتك فوراً.")}

        if kind == "csv":
            pr = parse_attlog(attlog, int(self.settings.get("dedup_seconds", 120)))
            proc = Processor(self.settings)
            att = proc.process(pr, year, month)
            export_csv(att, out, year, month)
            logs.append(f"تصدير CSV للتحقق: {len(att.rows)} سطر")
            files = [out]
        else:
            settings = dict(self.settings)
            tpl = p.get("template") or self.settings.get("template_path") or ""
            if tpl:
                settings["template_path"] = tpl
                # فحص مسبع للقالب: نشرح في السجل ما سيحدث قبل التوليد
                try:
                    from template_writer import analyze_template
                    tinfo = analyze_template(tpl, year, month)
                    if tinfo.get("ok") and tinfo.get("mode"):
                        logs.append("القالب: " + tinfo["mode"]
                                    + f" ({len(tinfo.get('month_days_found') or [])} يوماً)")
                    elif tinfo.get("warning"):
                        logs.append("⚠ القالب: " + tinfo["warning"])
                    elif tinfo.get("error"):
                        logs.append("⚠ القالب: " + tinfo["error"] + " — سيُستخدم المولّد المدمج")
                except Exception as e:
                    logs.append(f"⚠ تعذر فحص القالب: {e} — سيُستخدم المولّد المدمج")
            try:
                pr, att, scan, quality = run_pipeline(
                    attlog, year, month, out, settings, verbose=False)
                logs.append(f"[1/4] قراءة attlog: {len(pr.records)} بصمة سليمة | "
                            f"{len(pr.corrupt_lines)} سطر تالف | {pr.duplicates_removed} مكرر محذوف")
                logs.append(f"[2/4] المعالجة: {len(att.rows)} يوم/موظف | "
                            f"{len(att.unmapped_ids)} رقم غير مربوط")
                if scan is not None:
                    logs.append("[3/4] Excel (وضع القالب الأصلي)")
                    logs.extend("      " + m for m in scan.messages)
                else:
                    logs.append("[3/4] Excel (المولّد المدمج المطابق للوصف)")
                files = [out]
            except Exception as e:
                return {"ok": False, "error": f"فشل التوليد: {e}"}
        return {"ok": True, "path": out, "files": files, "logs": logs}

    def quality(self, p):
        attlog = p.get("attlog", "")
        year, month = self._ym(p)
        if not attlog or not os.path.exists(attlog):
            return {"ok": False, "error": "مسار attlog غير صحيح"}
        pr = parse_attlog(attlog, int(self.settings.get("dedup_seconds", 120)))
        proc = Processor(self.settings)
        att = proc.process(pr, year, month)
        return {"ok": True, "lines": build_report(att, pr, self.settings, year, month)}

    def save_quality(self, p):
        attlog = p.get("attlog", "")
        year, month = self._ym(p)
        if not attlog or not os.path.exists(attlog):
            return {"ok": False, "error": "مسار attlog غير صحيح"}
        out = p.get("out") or self._default_out("quality", year, month)
        pr = parse_attlog(attlog, int(self.settings.get("dedup_seconds", 120)))
        proc = Processor(self.settings)
        att = proc.process(pr, year, month)
        save_report(build_report(att, pr, self.settings, year, month), out)
        return {"ok": True, "files": [out]}

    def scan_ids(self, p):
        """فحص attlog وإضافة الأرقام غير المعروفة إلى جدول الربط."""
        attlog = p.get("attlog", "")
        if not attlog or not os.path.exists(attlog):
            return {"ok": False, "error": "مسار attlog غير صحيح"}
        pr = parse_attlog(attlog, int(self.settings.get("dedup_seconds", 120)))
        known = {e.get("id") for e in self.settings.get("employees", [])}
        added = 0
        for emp_id, (dev, _cnt) in pr.ids_found().items():
            if emp_id not in known:
                self.settings.setdefault("employees", []).append(
                    {"id": emp_id, "template_name": "", "device_name": dev, "no_punch": False})
                added += 1
        if added:
            self._save()
        return {"ok": True, "added": added}


API = Api()
_RUN_MODE = "browser"

# الطرق المسموح استدعاؤها من الواجهة عبر HTTP
ALLOWED = {"app_info", "get_state", "save_settings", "select_file", "save_dialog",
           "ingest_file", "open_path", "download_file", "read_output_b64",
           "preview", "generate", "quality", "save_quality", "scan_ids",
           "template_info"}


# ---------- وضع المتصفح الاحتياطي ----------

_MIME = {".html": "text/html; charset=utf-8", ".css": "text/css; charset=utf-8",
         ".js": "application/javascript; charset=utf-8", ".ttf": "font/ttf",
         ".png": "image/png", ".svg": "image/svg+xml", ".ico": "image/x-icon",
         ".woff2": "font/woff2", ".json": "application/json; charset=utf-8"}


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):  # إسكات سجل الكونسول
        pass

    def _send(self, code, body: bytes, ctype: str):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-cache")
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        path = self.path.split("?")[0]
        # ----- تنزيل ملف من مجلد المخرجات (اسم فقط، بدون مسارات) -----
        if path == "/api/download":
            qs = parse_qs(urlparse(self.path).query)
            name = os.path.basename((qs.get("name") or [""])[0])
            full = os.path.join(OUTPUT_DIR, name)
            if not name or not os.path.isfile(full):
                self._send(404, "الملف غير موجود".encode("utf-8"),
                           "text/plain; charset=utf-8")
                return
            with open(full, "rb") as f:
                body = f.read()
            self.send_response(200)
            self.send_header("Content-Type", "application/octet-stream")
            self.send_header("Content-Length", str(len(body)))
            ext = os.path.splitext(full)[1].lower()
            ascii_fallback = "attendance" + (ext if ext in (".xlsx", ".csv", ".txt") else ".bin")
            self.send_header(
                "Content-Disposition",
                f"attachment; filename=\"{ascii_fallback}\"; "
                f"filename*=UTF-8''{quote(name)}")
            self.send_header("Cache-Control", "no-cache")
            self.end_headers()
            self.wfile.write(body)
            return
        if path in ("/", "/index.html"):
            path = "/index.html"
        fname = os.path.normpath(path.lstrip("/")).replace("..", "")
        full = os.path.join(WEB_DIR, fname)
        if not os.path.isfile(full):
            self._send(404, "غير موجود".encode("utf-8"), "text/plain; charset=utf-8")
            return
        ext = os.path.splitext(full)[1].lower()
        with open(full, "rb") as f:
            self._send(200, f.read(), _MIME.get(ext, "application/octet-stream"))

    def do_POST(self):
        if not self.path.startswith("/api/"):
            self._send(404, b"{}", "application/json")
            return
        fn = self.path[len("/api/"):].strip("/")
        try:
            length = int(self.headers.get("Content-Length") or 0)
            payload = json.loads(self.rfile.read(length) or b"{}") if length else {}
        except Exception:
            payload = {}
        if fn not in ALLOWED or not hasattr(API, fn):
            self._send(200, json.dumps({"ok": False, "error": "طريقة غير معروفة"},
                                       ensure_ascii=False).encode("utf-8"), "application/json; charset=utf-8")
            return
        try:
            result = getattr(API, fn)(payload)
        except Exception as e:
            result = {"ok": False, "error": f"{type(e).__name__}: {e}"}
        self._send(200, json.dumps(result, ensure_ascii=False).encode("utf-8"),
                   "application/json; charset=utf-8")


def run_browser_mode(port: int = 0):
    global _RUN_MODE
    _RUN_MODE = "browser"
    # منفذ صريح = وضع خدمة (ربط على كل الواجهات)، منفذ عشوائي = تشغيل محلي
    host = "0.0.0.0" if port else "127.0.0.1"
    server = ThreadingHTTPServer((host, port), Handler)
    bound_port = server.server_address[1]
    url = f"http://127.0.0.1:{bound_port}/"
    print(f"[*] وضع المتصفح: {url}")
    print("[*] أغلق نافذة الطرفية (أو Ctrl+C) لإيقاف التطبيق.")
    if not port:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n[*] تم الإيقاف.")


def run_webview_mode():
    global _RUN_MODE
    import webview  # pywebview
    _RUN_MODE = "webview"
    index = os.path.join(WEB_DIR, "index.html")
    webview.create_window(
        "نظام تفريغ البصمات — جدول الدوام الآلي",
        index,
        js_api=API,
        width=1300, height=860,
        min_size=(1080, 700),
        background_color="#0b1220",
    )
    webview.start(http_server=True)


def main():
    force_browser = "--browser" in sys.argv
    port = 0
    for a in sys.argv[1:]:
        if a.startswith("--port="):
            try:
                port = int(a.split("=", 1)[1])
                force_browser = True
            except ValueError:
                pass
    if not force_browser:
        try:
            run_webview_mode()
            return 0
        except ImportError:
            print("[!] مكتبة pywebview غير مثبتة — التحويل إلى وضع المتصفح.")
            print("    لتشغيل النافذة الأصلية:  pip install pywebview")
        except Exception as e:
            print(f"[!] تعذر تشغيل النافذة الأصلية ({e}) — التحويل إلى وضع المتصفح.")
    run_browser_mode(port)
    return 0


if __name__ == "__main__":
    sys.exit(main())
