# -*- coding: utf-8 -*-
"""
app.py — نظام تفريغ البصمات تلقائياً إلى جدول الدوام
الواجهة الرئيسية: سطر أوامر + تشغيل الواجهة الرسومية

أمثلة:
  python app.py --gui
  python app.py --attlog attlog.txt --year 2026 --month 9 --out out.xlsx
  python app.py --attlog attlog.txt --year 2026 --month 9 --out out.xlsx --csv check.csv --quality quality.txt
"""
from __future__ import annotations

import argparse
import json
import os
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SETTINGS_PATH = os.path.join(BASE_DIR, "settings.json")

MONTH_AR = {1: "يناير", 2: "فبراير", 3: "مارس", 4: "أبريل", 5: "مايو", 6: "يونيو",
            7: "يوليو", 8: "أغسطس", 9: "سبتمبر", 10: "أكتوبر", 11: "نوفمبر", 12: "ديسمبر"}


def load_settings(path: str = SETTINGS_PATH) -> dict:
    if not os.path.exists(path):
        return {"employees": [], "notes": {}}
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_settings(settings: dict, path: str = SETTINGS_PATH):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)


def run_pipeline(attlog: str, year: int, month: int, out_xlsx: str,
                 settings: dict, csv_path: str = None, quality_path: str = None,
                 template_path: str = None, verbose: bool = True):
    """يشغّل خط المعالجة كاملاً ويعيد (parse_result, att, scan, quality_lines)."""
    from att_parser import parse_attlog
    from processor import Processor
    from excel_writer import write_fallback_excel, export_csv
    from quality_report import build_report, save_report

    settings = dict(settings)
    if template_path:
        settings["template_path"] = template_path

    # 1) القراءة
    pr = parse_attlog(attlog, dedup_seconds=int(settings.get("dedup_seconds", 120)))
    if verbose:
        print(f"[1/4] قراءة attlog: {len(pr.records)} بصمة سليمة | "
              f"{len(pr.corrupt_lines)} سطر تالف | {pr.duplicates_removed} مكرر محذوف | "
              f"الترميز: {pr.encoding}")

    # 2) المعالجة
    proc = Processor(settings)
    att = proc.process(pr, year, month)
    if verbose:
        print(f"[2/4] المعالجة: {len(att.rows)} يوم/موظف | "
              f"{len(att.unmapped_ids)} رقم غير مربوط")

    # 3) توليد Excel
    scan = None
    tpl = settings.get("template_path", "")
    used_fallback = True
    if tpl and os.path.exists(tpl):
        try:
            from template_writer import write_month_to_template
            scan = write_month_to_template(tpl, out_xlsx, att,
                                           settings.get("employees", []), year, month)
            used_fallback = False
            if verbose:
                print(f"[3/4] Excel (وضع القالب): {out_xlsx}")
                for m in scan.messages:
                    print("      - " + m)
        except Exception as e:
            if verbose:
                print(f"[3/4] تعذر استخدام القالب ({e}) → المولّد الاحتياطي")
    if used_fallback:
        write_fallback_excel(att, settings, out_xlsx, year, month)
        if verbose:
            print(f"[3/4] Excel (مولّد احتياطي مطابق للوصف): {out_xlsx}")

    # 4) CSV + تقرير الجودة
    if csv_path:
        export_csv(att, csv_path, year, month)
        if verbose:
            print(f"[4/4] CSV للتحقق اليدوي: {csv_path}")
    quality = build_report(att, pr, settings, year, month)
    if quality_path:
        save_report(quality, quality_path)
        if verbose:
            print(f"      تقرير الجودة: {quality_path}")

    return pr, att, scan, quality


def main(argv=None):
    ap = argparse.ArgumentParser(description="تفريغ البصمات تلقائياً إلى جدول الدوام")
    ap.add_argument("--gui", action="store_true", help="تشغيل الواجهة الرسومية")
    ap.add_argument("--attlog", help="مسار ملف attlog.txt")
    ap.add_argument("--year", type=int, default=2026)
    ap.add_argument("--month", type=int, default=9)
    ap.add_argument("--out", help="مسار ملف Excel الناتج")
    ap.add_argument("--csv", help="(اختياري) تصدير CSV للتحقق اليدوي")
    ap.add_argument("--quality", help="(اختياري) حفظ تقرير الجودة")
    ap.add_argument("--template", help="مسار ملف القالب الأصلي (اختياري)")
    ap.add_argument("--settings", default=SETTINGS_PATH)
    args = ap.parse_args(argv)

    if args.gui or not args.attlog:
        from gui import launch
        launch()
        return 0

    settings = load_settings(args.settings)
    out_xlsx = args.out or f"دوام_{args.year}_{args.month:02d}.xlsx"
    pr, att, scan, quality = run_pipeline(
        args.attlog, args.year, args.month, out_xlsx, settings,
        csv_path=args.csv, quality_path=args.quality, template_path=args.template)

    print("\n===== تقرير الجودة =====")
    for line in quality:
        print(line)
    print(f"\nتم: {out_xlsx}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
