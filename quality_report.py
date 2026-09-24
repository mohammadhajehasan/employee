# -*- coding: utf-8 -*-
"""
quality_report.py — تقرير الجودة (المرحلة 6)
يجمع كل ما يحتاج مراجعة بشرية قبل اعتماد الجدول:
  - أرقام وظيفية في attlog غير مربوطة بأسماء القالب
  - أيام بنقص بصمة (دخول أو خروج)
  - موظفون بلا أي بصمة طوال الشهر
  - بصمات في أيام عطلة (الجمعة/السبت)
  - أسطر تالفة واستُبعدت
  - بصمات مكررة حُذفت
"""
from __future__ import annotations

from processor import MON, TUE, WED, THU, FRI, SAT, SUN


def build_report(att, parse_result, settings: dict, year: int, month: int) -> list:
    """يعيد قائمة سطور جاهزة للعرض/الحفظ."""
    lines = []
    workdays = set(int(x) for x in settings.get("workdays", [SUN, MON, TUE, WED, THU]))
    no_punch_names = set(e.get("template_name", "") for e in settings.get("employees", [])
                         if e.get("no_punch"))
    employees = settings.get("employees", [])

    # 1) أرقام غير مربوطة
    if att.unmapped_ids:
        lines.append("== أرقام وظيفية في attlog غير مربوطة بأسماء القالب (يجب ربطها من شاشة الربط) ==")
        for emp_id, dev in sorted(att.unmapped_ids.items()):
            lines.append(f"  - الرقم {emp_id} (اسم الجهاز: {dev or 'غير معروف'})")
        lines.append("")

    # 2) أيام بنقص بصمة
    missing = [(k, row) for k, row in att.rows.items() if row.sys_note and "خروج وعودة" not in row.sys_note]
    if missing:
        lines.append("== أيام بنقص بصمة ==")
        for (emp_id, d), row in sorted(missing, key=lambda x: (x[0][1], x[0][0])):
            who = row.template_name or f"#{emp_id}"
            lines.append(f"  - {d.isoformat()}: {who} ← {row.sys_note} ({row.in_time and 'دخول ' or ''}"
                         f"{row.out_time and 'خروج ' or ''})")
        lines.append("")

    # 3) خروج وعودة (معلومة)
    mids = [(k, row) for k, row in att.rows.items() if "خروج وعودة" in row.sys_note]
    if mids:
        lines.append("== أيام فيها خروج وعودة (أكثر من بصمتين) ==")
        for (emp_id, d), row in sorted(mids, key=lambda x: (x[0][1], x[0][0])):
            who = row.template_name or f"#{emp_id}"
            lines.append(f"  - {d.isoformat()}: {who} ← {row.sys_note}")
        lines.append("")

    # 4) بصمات في عطلة نهاية الأسبوع
    weekend = [(k, row) for k, row in att.rows.items() if row.d.weekday() not in workdays]
    if weekend:
        lines.append("== بصمات في أيام غير أيام دوام (تظهر كتحذير ولا تُدرج في القالب) ==")
        for (emp_id, d), row in sorted(weekend, key=lambda x: x[0][1]):
            who = row.template_name or f"#{emp_id}"
            day_ar = {SAT: "السبت", SUN: "الأحد", MON: "الاثنين", TUE: "الثلاثاء",
                      WED: "الأربعاء", THU: "الخميس", FRI: "الجمعة"}.get(d.weekday(), "")
            lines.append(f"  - {d.isoformat()} ({day_ar}): {who}")
        lines.append("")

    # 5) موظفو القالب بلا أي بصمة طوال الشهر
    seen_ids = set(emp_id for (emp_id, _) in att.rows.keys())
    silent = [e for e in employees
              if e.get("id") and not e.get("no_punch") and e["id"] not in seen_ids]
    if silent:
        lines.append("== موظفون في القالب بلا أي بصمة طوال الشهر (يبقى صفهم فارغاً) ==")
        for e in silent:
            lines.append(f"  - {e.get('template_name', '?')} (الرقم {e['id']})")
        lines.append("")

    # 6) أسطر تالفة
    if parse_result.corrupt_lines:
        lines.append(f"== أسطر تالفة استُبعدت ({len(parse_result.corrupt_lines)}) ==")
        for line_no, text, reason in parse_result.corrupt_lines[:20]:
            lines.append(f"  - سطر {line_no}: {reason} ← {text}")
        if len(parse_result.corrupt_lines) > 20:
            lines.append(f"  ... و{len(parse_result.corrupt_lines) - 20} أسطر أخرى")
        lines.append("")

    # 7) تكرار محذوف
    if parse_result.duplicates_removed:
        lines.append(f"== بصمات مكررة حُذفت (فارق أقل من دقيقتين): {parse_result.duplicates_removed} ==")
        lines.append("")

    if not lines:
        lines.append("لا توجد ملاحظات جودة — كل شيء سليم.")
    return lines


def save_report(lines: list, path: str):
    with open(path, "w", encoding="utf-8-sig") as f:
        f.write("\n".join(lines) + "\n")
