# -*- coding: utf-8 -*-
"""
template_writer.py — وضع القالب (الأدق): يعبّئ نسخة من ملف قالبك الأصلي نفسه
مما يضمن مطابقة التنسيق 100% (ألوان، دمج، حدود، RTL) دون إعادة بناء.

آلية العمل:
  1) مسح القالب للعثور على خلايا التواريخ الخاصة بشهر الهدف (تواريخ حقيقية في الخلايا)
  2) تحديد مجموعة أعمدة كل يوم (3 أعمدة: دخول / خروج / ملاحظات) عبر دمج الخلية
     أو عبر رؤوس "دخول/خروج/ملاحظات" أسفلها
  3) مطابقة أسماء الموظفين (بمقارنة عربية مطبّعة) لتحديد صف كل موظف
  4) كتابة القيم في الخلايا الصحيحة فقط — كل التنسيق يبقى كما هو
"""
from __future__ import annotations

import calendar
import re
from datetime import date, datetime

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from processor import fmt_time

AR_DIAC = re.compile(r"[\u064B-\u0652\u0640]")  # تشكيل + تطويل


def normalize_ar(s: str) -> str:
    """توحيد النص العربي للمقارنة: بدون تشكيل، أ/إ/آ→ا، ى→ي، ة→ه، مسافات مضغوطة."""
    if s is None:
        return ""
    s = str(s)
    s = AR_DIAC.sub("", s)
    s = s.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
    s = s.replace("ى", "ي").replace("ة", "ه").replace("ؤ", "و").replace("ئ", "ي")
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s


class TemplateScanError(Exception):
    pass


class TemplateScan:
    def __init__(self):
        self.date_anchors = {}      # date -> dict(col_start, col_end, sub_map, row)
        self.emp_rows = {}          # emp_id -> excel_row
        self.unmatched_names = []   # أسماء في القالب لم تُربط
        self.messages = []


def scan_template(wb, year: int, month: int) -> TemplateScan:
    scan = TemplateScan()
    ws = wb.active

    # 1) خلايا التواريخ لشهر الهدف
    anchors = []  # (row, col, d, merge_span)
    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            d = None
            if isinstance(v, datetime):
                d = v.date()
            elif isinstance(v, date):
                d = v
            if d and d.year == year and d.month == month:
                span = (cell.column, cell.column)
                for mr in ws.merged_cells.ranges:
                    if (cell.row, cell.column) in set():
                        pass
                    if mr.min_row <= cell.row <= mr.max_row and mr.min_col <= cell.column <= mr.max_col:
                        span = (mr.min_col, mr.max_col)
                        break
                anchors.append((cell.row, cell.column, d, span))

    if not anchors:
        raise TemplateScanError(
            "لم أجد خلايا تواريخ حقيقية لشهر الهدف في القالب. "
            "تأكد أن خلايا التواريخ في القالب قيم تاريخ وليست نصاً، "
            "أو استخدم المولّد الاحتياطي وأرسل لي نسخة من القالب لمعايرته.")

    for (row, col, d, span) in anchors:
        if d in scan.date_anchors:
            continue
        info = {"row": row, "col_start": span[0], "col_end": span[1], "sub_map": {}}

        # 2) تحديد أعمدة الدخول/الخروج/الملاحظات: ابحث تحت خلية التاريخ حتى 4 صفوف
        labels = {"دخول": None, "خروج": None, "ملاحظات": None}
        for rr in range(row + 1, min(row + 5, ws.max_row + 1)):
            for cc in range(span[0], min(span[1] + 3, ws.max_column + 1)):
                v = ws.cell(row=rr, column=cc).value
                if isinstance(v, str):
                    nv = normalize_ar(v)
                    for lbl in labels:
                        if labels[lbl] is None and lbl in nv:
                            labels[lbl] = cc
            if all(v is not None for v in labels.values()):
                break
        if labels["دخول"] and labels["خروج"]:
            info["sub_map"] = {"in": labels["دخول"], "out": labels["خروج"],
                               "note": labels["ملاحظات"] or (labels["خروج"] + 1)}
        else:
            # افتراضي: التاريخ مدموج على 3 أعمدة بالترتيب
            info["sub_map"] = {"in": span[0], "out": span[0] + 1, "note": span[0] + 2}
        scan.date_anchors[d] = info

    scan.messages.append(f"تم العثور على {len(scan.date_anchors)} مجموعة أيام لشهر الهدف.")

    # 3) مطابقة صفوف الموظفين: افحص أول عمودين
    name_cells = []
    for row in ws.iter_rows(min_col=1, max_col=2):
        for cell in row:
            if isinstance(cell.value, str) and len(cell.value.strip()) > 2:
                name_cells.append((cell.row, cell.value))

    return scan


def match_employee_rows(wb, scan: TemplateScan, employees: list):
    """يربط أسماء settings بصفوف القالب عبر المقارنة المطبّعة."""
    ws = wb.active
    name_cells = []
    for row in ws.iter_rows(min_col=1, max_col=2):
        for cell in row:
            if isinstance(cell.value, str) and len(cell.value.strip()) > 2:
                name_cells.append((cell.row, normalize_ar(cell.value), cell.value))

    for emp in employees:
        tn = emp.get("template_name", "")
        if not tn:
            continue
        ntn = normalize_ar(tn)
        hit = None
        for (r, nv, orig) in name_cells:
            if nv == ntn:
                hit = r
                break
        if hit is None:
            for (r, nv, orig) in name_cells:
                if ntn and (ntn in nv or nv in ntn):
                    hit = r
                    break
        if hit:
            scan.emp_rows[emp["id"]] = hit
        else:
            scan.unmatched_names.append(tn)

    scan.messages.append(f"تم ربط {len(scan.emp_rows)} موظفاً بصفوف القالب.")
    if scan.unmatched_names:
        scan.messages.append("أسماء لم تُعثر على صف لها: " + " ، ".join(scan.unmatched_names))
    return scan


def write_month_to_template(template_path: str, out_path: str, att,
                            employees: list, year: int, month: int,
                            clear_existing: bool = True):
    """يفتح نسخة من القالب، يعبّئه، ويحفظه باسم جديد. يعيد تقرير المسح."""
    wb = load_workbook(template_path)
    scan = scan_template(wb, year, month)
    match_employee_rows(wb, scan, employees)
    ws = wb.active

    written = 0
    for (emp_id, d), r in att.rows.items():
        info = scan.date_anchors.get(d)
        row = scan.emp_rows.get(emp_id)
        if not info or not row:
            continue  # يظهر في تقرير الجودة
        sm = info["sub_map"]
        if r.in_time:
            ws.cell(row=row, column=sm["in"], value=fmt_time(r.in_time))
            written += 1
        if r.out_time:
            ws.cell(row=row, column=sm["out"], value=fmt_time(r.out_time))
            written += 1
        if r.note:
            ws.cell(row=row, column=sm["note"], value=r.note)
            written += 1

    wb.properties.creator = "Z.ai"
    wb.save(out_path)
    scan.messages.append(f"تمت كتابة {written} خلية في نسخة القالب.")
    return scan
