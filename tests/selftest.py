# -*- coding: utf-8 -*-
"""
tests/selftest.py — اختبار ذاتي شامل لخط المعالجة كاملاً
يبني ملف attlog تجريبياً لشهر سبتمبر 2026 يحاكي كل الأنماط الموجودة في ملف المستخدم
(بنفس الأسطر الدليلية التي ذكرها في الخطة) ثم يتحقق من النتائج واحدًا واحدًا.
"""
from __future__ import annotations

import calendar
import os
import sys
from datetime import date, datetime, timedelta

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)

SAMPLE_DIR = os.path.join(BASE, "sample")
OUT_DIR = os.path.join(BASE, "output")
os.makedirs(SAMPLE_DIR, exist_ok=True)
os.makedirs(OUT_DIR, exist_ok=True)

ATTLOG = os.path.join(SAMPLE_DIR, "attlog_sample.txt")
OUT_XLSX = os.path.join(OUT_DIR, "دوام_سبتمبر_2026.xlsx")
OUT_CSV = os.path.join(OUT_DIR, "تحقق_سبتمبر_2026.csv")
OUT_Q = os.path.join(OUT_DIR, "تقرير_الجودة.txt")


def hhmm(base_hour, base_min, add_min=0):
    t = datetime(2026, 9, 1, base_hour, 0) + timedelta(minutes=base_min + add_min)
    return t.strftime("%H:%M:%S")


def build_sample_attlog():
    """يبني attlog تجريبياً يغطي كل الحالات."""
    lines = []
    add = lines.append
    workdays = [d for d in (date(2026, 9, x) for x in range(1, 31))
                if d.weekday() in (6, 0, 1, 2, 3)]  # الأحد–الخميس

    roster = [1025, 1000, 1004, 1006, 1017, 1022]
    names = {1025: "Yazeed", 1000: "Dania ma", 1004: "Hazem Go",
             1006: "Baraa B", 1017: "M.Tareq", 1022: "Sallam"}

    k = 0
    for d in workdays:
        for emp in roster:
            k += 1
            jitter = (k % 7)
            # استثناءات الحالات الخاصة (نضيفها يدوياً بعد الحلقة)
            if emp == 1003 or d == date(2026, 9, 20):
                continue
            if emp == 1004 and d == date(2026, 9, 9):
                continue
            if emp == 1006 and d == date(2026, 9, 8):
                continue
            if emp == 1022 and d == date(2026, 9, 7):
                continue
            if emp == 1000 and d == date(2026, 9, 15):
                continue
            if emp == 1022 and d == date(2026, 9, 1):
                continue
            add(f"{emp}\t{names[emp]}\t{d} {hhmm(7, 55 + jitter % 10)}\t0\t0")
            add(f"{emp}\t{names[emp]}\t{d} {hhmm(15, 2 + jitter % 8)}\t1\t1")

    # 1) الأسطر الدليلية من ملف المستخدم (عمود الحالة مضروب عمداً)
    add("1003\tA.Diab\t2026-09-03 08:05:49\t1\t1")      # صباحية مسجلة كخروج → تدخل كدخول
    add("1022\tSallam\t2026-09-01 14:48:18\t1\t0")      # عصرية مسجلة كدخول → تخرج كخروج
    add("1000\tDania ma\t2026-09-15 14:55:06\t1\t0")    # نفس المشكلة

    # 2) 1003 حاضر باقي أيامه (بالأسطر العادية) — نضيف أزواجه لكل أيام العمل ما عدا الجمعة
    for d in workdays:
        if d == date(2026, 9, 3):
            continue
        add(f"1003\tA.Diab\t{d} {hhmm(8, 5, int(d.day) % 9)}\t0\t0")
        add(f"1003\tA.Diab\t{d} {hhmm(15, 10, int(d.day) % 7)}\t1\t1")
    # بصمة يوم الجمعة (عطلة) → تحذير
    add("1003\tA.Diab\t2026-09-04 10:22:00\t0\t0")

    # 3) نقص بصمة خروج: 1004 يوم 9/9 بصمة صباحية واحدة
    add("1004\tHazem Go\t2026-09-09 08:01:00\t0\t0")

    # 4) خروج وعودة (3 بصمات): 1006 يوم 8/9
    add("1006\tBaraa B\t2026-09-08 07:59:00\t0\t0")
    add("1006\tBaraa B\t2026-09-08 12:34:00\t1\t0")
    add("1006\tBaraa B\t2026-09-08 15:07:00\t0\t1")

    # 5) تكرار بأقل من دقيقتين: 1022 يوم 7/9 (بصمة ثانية بعد 90 ثانية)
    add("1022\tSallam\t2026-09-07 07:59:00\t0\t0")
    add("1022\tSallam\t2026-09-07 08:00:30\t0\t0")
    add("1022\tSallam\t2026-09-07 15:03:00\t1\t1")

    # 6) رقم غير مربوط بقالب
    add("1077\tMystery\t2026-09-10 08:12:00\t0\t0")
    add("1077\tMystery\t2026-09-10 15:31:00\t1\t1")

    # 7) أسطر تالفة
    add("1010\tBadGuy\t2026-09-30 25:99:00\t1\t1")   # وقت غير صالح
    add("سطر تالف تماماً بدون أرقام")                  # غير قابل للتحليل
    add("")                                            # سطر فارغ

    with open(ATTLOG, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    return len(lines)


def main():
    total = build_sample_attlog()
    print(f"① بُني ملف attlog تجريبي: {ATTLOG} ({total} سطر)")

    from app import load_settings, run_pipeline
    from att_parser import parse_attlog
    from processor import Processor, fmt_time
    from quality_report import build_report

    settings = load_settings(os.path.join(BASE, "settings.json"))
    # ملاحظة إضافية لاختبار الدمج مع بصمات فعلية
    settings["notes"]["2026-09-06|1017"] = "مأمورية"

    pr = parse_attlog(ATTLOG, int(settings["dedup_seconds"]))
    proc = Processor(settings)
    att = proc.process(pr, 2026, 9)
    quality = build_report(att, pr, settings, 2026, 9)

    # كتابة المخرجات عبر خط الأنابيب الكامل
    run_pipeline(ATTLOG, 2026, 9, OUT_XLSX, settings,
                 csv_path=OUT_CSV, quality_path=OUT_Q, verbose=True)

    checks = []
    def check(name, cond, detail=""):
        checks.append((name, bool(cond), detail))

    # --- تحقق القراءة ---
    check("كشف الترميز utf-8", pr.encoding.startswith("utf-8"), pr.encoding)
    check("حذف البصمة المكررة (1022 يوم 7/9)", pr.duplicates_removed >= 1, str(pr.duplicates_removed))
    check("استبعاد الأسطر التالفة (>=2)", len(pr.corrupt_lines) >= 2,
          "; ".join(f"سطر {n}: {r}" for n, _, r in pr.corrupt_lines))

    def row(emp, iso):
        return att.rows.get((emp, date.fromisoformat(iso)))

    # --- تحقق الاقتران والتجاهل عمود الحالة ---
    r = row(1003, "2026-09-03")
    check("1003 يوم 3/9: الصباحية = دخول 8:05 رغم الحالة 1 1",
          r and r.in_time and fmt_time(r.in_time) == "8:05",
          f"in={r.in_time} out={r.out_time}")

    r = row(1022, "2026-09-01")
    check("1022 يوم 1/9: العصرية = خروج 14:48 + نقص بصمة دخول",
          r and r.out_time and fmt_time(r.out_time) == "14:48" and "نقص بصمة دخول" in r.note,
          f"note={r.note}")

    r = row(1000, "2026-09-15")
    check("1000 يوم 15/9: خروج 14:55 + نقص بصمة دخول",
          r and r.out_time and fmt_time(r.out_time) == "14:55" and "نقص بصمة دخول" in r.note,
          f"note={r.note}")

    r = row(1004, "2026-09-09")
    check("1004 يوم 9/9: دخول فقط + نقص بصمة خروج",
          r and r.in_time and not r.out_time and "نقص بصمة خروج" in r.note,
          f"note={r.note}")

    r = row(1006, "2026-09-08")
    check("1006 يوم 8/9: دخول 7:59 وخروج 15:07 + خروج وعودة 12:34",
          r and fmt_time(r.in_time) == "7:59" and fmt_time(r.out_time) == "15:07"
          and "خروج وعودة: 12:34" in r.note,
          f"in={r.in_time} out={r.out_time} note={r.note}")

    r = row(1022, "2026-09-07")
    check("1022 يوم 7/9: التكرار دُمج (بصمتان فقط بعد الدمج)",
          r and r.punches == 2 and fmt_time(r.in_time) == "7:59",
          f"punches={r.punches} in={r.in_time}")

    # --- تحقق الملاحظات المحفوظة ---
    r = row(1017, "2026-09-06")
    check("1017 يوم 6/9: مأمورية دُمجت مع البصمات",
          r and "مأمورية" in r.note and r.in_time is not None, f"note={r.note}")

    r = row(1017, "2026-09-20")
    check("1017 يوم 20/9: اجازة تظهر رغم عدم وجود بصمات",
          r and r.note == "اجازة" and r.punches == 0, f"note={r.note}")

    # --- تحقق تقرير الجودة ---
    qtext = "\n".join(quality)
    check("تقرير: رقم غير مربوط 1077", "1077" in qtext)
    check("تقرير: بصمة يوم الجمعة (عطلة)", "الجمعة" in qtext)
    check("تقرير: أسطر تالفة", "أسطر تالفة" in qtext)

    # --- تحقق المخرجات ---
    check("ملف Excel وُجد", os.path.exists(OUT_XLSX))
    check("ملف CSV وُجد", os.path.exists(OUT_CSV))
    check("تقرير الجودة حُفظ", os.path.exists(OUT_Q))

    from openpyxl import load_workbook
    wb = load_workbook(OUT_XLSX)
    ws = wb.active
    check("Excel: ورقة RTL", bool(ws.sheet_view.rightToLeft))
    a1 = ws.cell(row=1, column=1).value or ""
    check("Excel: العنوان يحتوي سبتمبر 2026", "سبتمبر" in a1 and "2026" in a1, a1)

    # وقت مكتوب بصيغة h:mm في خلية (ابحث عن 8:05 ليوم 3/9 لـ1003)
    found_805 = False
    for rrow in ws.iter_rows():
        for c in rrow:
            if c.value == "8:5" or c.value == "8:05":
                found_805 = True
    # 1003 يدخل 8:05 + (day%9) → 3/9 = 8:05+3 = 8:8? لا: hhmm(8,5,int(d.day)%9) → دقائق 5+؟
    # في الحقيقة البصمة 3/9 مستثناة (السطر الدليلي 8:05:49) → يجب أن تظهر 8:05
    check("Excel: قيمة 8:05 (1003 يوم 3/9) موجودة", found_805)

    print("\n===== نتائج الاختبار =====")
    passed = 0
    for name, ok, detail in checks:
        mark = "PASS" if ok else "FAIL"
        if ok:
            passed += 1
        print(f"[{mark}] {name}" + (f"  → {detail}" if (detail and not ok) else ""))
    print(f"\n{passed}/{len(checks)} ناجحة")
    return 0 if passed == len(checks) else 1


if __name__ == "__main__":
    sys.exit(main())
