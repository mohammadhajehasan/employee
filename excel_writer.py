# -*- coding: utf-8 -*-
"""
excel_writer.py — محرك Excel (المرحلة 2): المولّد الاحتياطي
يبني جدولاً مطابقاً لبنية القالب المرجعي (نفس ترتيب قالبكم):
  - ورقة RTL بعنوان «جدول الدوام لشهر X السنة» مدموجاً
  - صف التواريخ: كل يوم دوام مختار في الشهر تسلسلياً (تاريخ حقيقي)
    مدموجاً على 3 خلايا: ساعة الدخول / ساعة الخروج / ملاحظات
  - الأسماء في العمود A ابتداءً من الصف 5
ملاحظة: عند توفر ملف القالب الأصلي يُستخدم template_writer.py الذي يعبّئ نسخة من
قالبك نفسه (مطابقة تنسيق 100%) — هذا المولّد للعمل بدون قالب وبنفس البنية.
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

HDR_GRAY = "D9D9D9"
EMPTY_FILL = "F2F2F2"

THIN = Side(style="thin", color="808080")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center")


def month_workdays(year: int, month: int, workdays):
    """كل أيام الدوام المختارة في الشهر مرتبة تصاعدياً (بنية القالب: شريط واحد)."""
    ndays = calendar.monthrange(year, month)[1]
    wd = {int(x) for x in workdays}
    return [date(year, month, d) for d in range(1, ndays + 1) if date(year, month, d).weekday() in wd]


def write_fallback_excel(att, settings: dict, out_path: str, year: int, month: int):
    workdays = [int(x) for x in settings.get("workdays", [SUN, MON, TUE, WED, THU])]
    employees = settings.get("employees", [])

    wb = Workbook()
    ws = wb.active
    ws.title = f"دوام {MONTH_AR.get(month, month)}"
    ws.sheet_view.rightToLeft = True

    days = month_workdays(year, month, set(workdays))
    last_col = 1 + max(1, len(days)) * 3

    # صف 1: العنوان مدموجاً على كل الأعمدة
    ws.cell(row=1, column=1, value=f"جدول الدوام لشهر {MONTH_AR.get(month, month)} {year}")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    c1 = ws.cell(row=1, column=1)
    c1.font = Font(bold=True, size=14)
    c1.alignment = CENTER

    # صف 2: تاريخ كل يوم مختار (تاريخ حقيقي) مدموجاً على 3 خلايا: B2:D2, E2:G2, ...
    for i, d in enumerate(days):
        col = 2 + i * 3
        ws.cell(row=2, column=col, value=datetime(d.year, d.month, d.day))
        ws.merge_cells(start_row=2, start_column=col, end_row=2, end_column=col + 2)
        cell = ws.cell(row=2, column=col)
        cell.number_format = "yyyy-mm-dd"
        cell.font = Font(bold=True)
        cell.alignment = CENTER

    # صف 3: رؤوس الأعمدة الثلاثة لكل يوم
    for i in range(len(days)):
        col = 2 + i * 3
        for off, label in enumerate(("ساعة الدخول", "ساعة الخروج", "ملاحظات")):
            cell = ws.cell(row=3, column=col + off, value=label)
            cell.font = Font(bold=True, size=9)
            cell.alignment = CENTER
            cell.border = BORDER
            cell.fill = PatternFill("solid", fgColor=HDR_GRAY)

    # الصف 4 فارغ (كما في القالب) — الأسماء من الصف 5
    row = 5
    for emp in employees:
        name = emp.get("template_name", "") or (f"(بدون ربط) #{emp.get('id')}" if emp.get("id") else "")
        ncell = ws.cell(row=row, column=1, value=name)
        ncell.alignment = RIGHT
        ncell.border = BORDER
        ncell.font = Font(bold=True)
        if emp.get("no_punch"):
            for c in range(2, last_col + 1):
                ws.cell(row=row, column=c).fill = PatternFill("solid", fgColor=EMPTY_FILL)
                ws.cell(row=row, column=c).border = BORDER
        elif emp.get("id"):
            for i, d in enumerate(days):
                r2 = att.rows.get((emp["id"], d))
                col = 2 + i * 3
                if r2:
                    if r2.in_time:
                        ws.cell(row=row, column=col, value=fmt_time(r2.in_time))
                    if r2.out_time:
                        ws.cell(row=row, column=col + 1, value=fmt_time(r2.out_time))
                    if r2.note:
                        ws.cell(row=row, column=col + 2, value=r2.note)
                for c in range(col, col + 3):
                    ws.cell(row=row, column=c).border = BORDER
                    ws.cell(row=row, column=c).alignment = CENTER
                ws.cell(row=row, column=col + 2).alignment = RIGHT
        row += 1

    # عرض الأعمدة: A للأسماء، ملاحظات أوسع قليلاً
    ws.column_dimensions["A"].width = 18
    for i in range(len(days)):
        col = 2 + i * 3
        ws.column_dimensions[get_column_letter(col)].width = 10
        ws.column_dimensions[get_column_letter(col + 1)].width = 10
        ws.column_dimensions[get_column_letter(col + 2)].width = 14

    # تثبيت صفوف الرأس وعمود الأسماء
    ws.freeze_panes = "B5"
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
