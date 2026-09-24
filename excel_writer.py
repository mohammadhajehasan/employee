# -*- coding: utf-8 -*-
"""
excel_writer.py — محرك Excel (المرحلة 2): المولّد الاحتياطي
يبني جدولاً مطابقاً لوصف القالب:
  - ورقة RTL
  - الأسبوع = شريطان: شريط (الأحد + الاثنين) وشريط (الثلاثاء + الأربعاء + الخميس)
  - كل يوم = 3 أعمدة: ساعة الدخول / ساعة الخروج / ملاحظات
  - الصفوف = قائمة الموظفين الثابتة (29 صفاً) مع صفوف فاصلة لأصحاب المناصب
ملاحظة: عند توفر ملف القالب الأصلي يُستخدم template_writer.py الذي يعبّئ نسخة من
قالبك نفسه (مطابقة تنسيق 100%) — هذا المولّد للعمل بدون قالب.
"""
from __future__ import annotations

import calendar
from datetime import date, datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from processor import MON, TUE, WED, THU, FRI, SAT, SUN, fmt_time

DAY_AR = {SAT: "السبت", SUN: "الأحد", MON: "الاثنين", TUE: "الثلاثاء",
          WED: "الأربعاء", THU: "الخميس", FRI: "الجمعة"}
MONTH_AR = {1: "يناير", 2: "فبراير", 3: "مارس", 4: "أبريل", 5: "مايو", 6: "يونيو",
            7: "يوليو", 8: "أغسطس", 9: "سبتمبر", 10: "أكتوبر", 11: "نوفمبر", 12: "ديسمبر"}

# ألوان الشريطين + الرؤوس
BAND_A = "DCE6F1"   # شريط الأحد + الاثنين (أزرق فاتح)
BAND_B = "E2EFDA"   # شريط الثلاثاء - الخميس (أخضر فاتح)
HDR_GRAY = "D9D9D9"
SEP_FILL = "F2F2F2"

THIN = Side(style="thin", color="808080")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center")


def _month_weeks(year: int, month: int, workdays):
    """يقسّم أيام الشهر إلى أسابيع (الأحد–الخميس) مع دعم الأسابيع الجزئية."""
    ndays = calendar.monthrange(year, month)[1]
    days = [date(year, month, d) for d in range(1, ndays + 1)]
    days = [d for d in days if d.weekday() in workdays]
    weeks = []
    cur = []
    for d in days:
        if cur and d.weekday() == SUN:
            weeks.append(cur)
            cur = []
        cur.append(d)
    if cur:
        weeks.append(cur)
    return weeks


def write_fallback_excel(att, settings: dict, out_path: str, year: int, month: int):
    workdays = [int(x) for x in settings.get("workdays", [SUN, MON, TUE, WED, THU])]
    employees = settings.get("employees", [])

    wb = Workbook()
    ws = wb.active
    ws.title = f"دوام {MONTH_AR.get(month, month)}"
    ws.sheet_view.rightToLeft = True

    weeks = _month_weeks(year, month, set(workdays))
    n_rows = 3 + len(employees)

    # عناوين عامة
    ws.cell(row=1, column=1, value=f"جدول الدوام — شهر {MONTH_AR.get(month, month)} {year}")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=1 + len(weeks) * 15)
    ws.cell(row=1, column=1).font = Font(bold=True, size=14)
    ws.cell(row=1, column=1).alignment = CENTER

    ws.cell(row=2, column=1, value="الاسم")
    ws.cell(row=2, column=1).font = Font(bold=True, size=11)
    ws.cell(row=2, column=1).alignment = CENTER
    ws.cell(row=2, column=1).border = BORDER

    thin_note_font = Font(size=9)

    col = 2
    for w_idx, week in enumerate(weeks, start=1):
        # تجميع أيام الأسبوع إلى شريطين
        strip_a = [d for d in week if d.weekday() in (SUN, MON)]
        strip_b = [d for d in week if d.weekday() in (TUE, WED, THU)]

        start_col = col
        # صف أسماء الشريطين
        if strip_a:
            ws.cell(row=2, column=col, value="الأحد + الاثنين")
            ws.merge_cells(start_row=2, start_column=col, end_row=2, end_column=col + len(strip_a) * 3 - 1)
            for c in range(col, col + len(strip_a) * 3):
                ws.cell(row=2, column=c).fill = PatternFill("solid", fgColor=BAND_A)
                ws.cell(row=2, column=c).border = BORDER
            ws.cell(row=2, column=col).alignment = CENTER
            col += len(strip_a) * 3
        if strip_b:
            ws.cell(row=2, column=col, value="الثلاثاء + الأربعاء + الخميس")
            ws.merge_cells(start_row=2, start_column=col, end_row=2, end_column=col + len(strip_b) * 3 - 1)
            for c in range(col, col + len(strip_b) * 3):
                ws.cell(row=2, column=c).fill = PatternFill("solid", fgColor=BAND_B)
                ws.cell(row=2, column=c).border = BORDER
            ws.cell(row=2, column=col).alignment = CENTER
            col += len(strip_b) * 3

        # صف أسماء الأيام والتواريخ + صف الدخول/الخروج/ملاحظات
        col = start_col
        sub_row = 3
        for d in week:
            band = BAND_A if d.weekday() in (SUN, MON) else BAND_B
            label = f"{DAY_AR[d.weekday()]} {d.day}/{month}"
            ws.cell(row=sub_row, column=col, value=label)
            ws.merge_cells(start_row=sub_row, start_column=col, end_row=sub_row, end_column=col + 2)
            for c in range(col, col + 3):
                ws.cell(row=sub_row, column=c).fill = PatternFill("solid", fgColor=band)
                ws.cell(row=sub_row, column=c).border = BORDER
            ws.cell(row=sub_row, column=col).alignment = CENTER
            ws.cell(row=sub_row, column=col).font = Font(bold=True)

            for off, label3 in enumerate(("ساعة الدخول", "ساعة الخروج", "ملاحظات")):
                cell = ws.cell(row=sub_row + 1, column=col + off, value=label3)
                cell.font = Font(bold=True, size=9)
                cell.alignment = CENTER
                cell.border = BORDER
                cell.fill = PatternFill("solid", fgColor=HDR_GRAY)
            col += 3

        # صفوف الموظفين
        row = 4
        for emp in employees:
            name = emp.get("template_name", "") or (f"(بدون ربط) #{emp.get('id')}" if emp.get("id") else "")
            ws.cell(row=row, column=1, value=name)
            ws.cell(row=row, column=1).alignment = RIGHT
            ws.cell(row=row, column=1).border = BORDER
            if emp.get("no_punch"):
                ws.cell(row=row, column=1).font = Font(bold=True)
                for c in range(2, 1 + (len(weeks) * 15) + 1):
                    ws.cell(row=row, column=c).fill = PatternFill("solid", fgColor=SEP_FILL)
                    ws.cell(row=row, column=c).border = BORDER
            elif emp.get("id"):
                col2 = 2
                for d in week:
                    r2 = att.rows.get((emp["id"], d))
                    if r2:
                        ws.cell(row=row, column=col2, value=fmt_time(r2.in_time) if r2.in_time else "")
                        ws.cell(row=row, column=col2 + 1, value=fmt_time(r2.out_time) if r2.out_time else "")
                        ws.cell(row=row, column=col2 + 2, value=r2.note if r2.note else "")
                    for c in range(col2, col2 + 3):
                        ws.cell(row=row, column=c).border = BORDER
                        ws.cell(row=row, column=c).alignment = CENTER
                    ws.cell(row=row, column=col2 + 2).alignment = RIGHT
                    col2 += 3
            row += 1

        col = start_col + len(week) * 3

    # عرض الأعمدة
    ws.column_dimensions["A"].width = 18
    for c in range(2, 2 + len(weeks) * 15):
        letter = get_column_letter(c)
        ws.column_dimensions[letter].width = 9
    for c in range(2, 2 + len(weeks) * 15, 3):
        letter = get_column_letter(c + 2)
        ws.column_dimensions[letter].width = 14

    # تثبيت أعمدة الأسماء
    ws.freeze_panes = "B4"
    wb.properties.creator = "Z.ai"
    wb.save(out_path)
    return out_path


def export_csv(att, out_path: str, year: int, month: int):
    """تصدير CSV للتحقق اليدوي (بترميز يدعم العربية في Excel)."""
    import csv
    with open(out_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["الرقم الوظيفي", "اسم الجهاز", "الاسم في القالب", "التاريخ", "اليوم",
                    "ساعة الدخول", "ساعة الخروج", "عدد البصمات",
                    "ملاحظة النظام", "ملاحظة محفوظة", "الملاحظة النهائية", "تأخير"])
        for (emp_id, d), r in sorted(att.rows.items(), key=lambda kv: (kv[0][1], kv[0][0])):
            w.writerow([emp_id, r.device_name, r.template_name, d.isoformat(), DAY_AR.get(d.weekday(), ""),
                        fmt_time(r.in_time) if r.in_time else "",
                        fmt_time(r.out_time) if r.out_time else "",
                        r.punches, r.sys_note, r.saved_note, r.note,
                        "نعم" if r.late else ""])
    return out_path
