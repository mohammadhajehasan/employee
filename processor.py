# -*- coding: utf-8 -*-
"""
processor.py — محرك المعالجة (قلب النظام)
الخوارزمية المعتمدة (سؤال 1 من الخطة — معتمدة من المستخدم):
  1. تجميع البصمات: (موظف + تاريخ) ثم ترتيب زمني — عمود الحالة يُتجاهل عند التعارض
  2. اقتران كل يوم:
     - بصمتان أو أكثر → الأولى = دخول، الأخيرة = خروج
     - بصمة واحدة قبل منتصف اليوم → دخول + "نقص بصمة خروج"
     - بصمة واحدة بعد منتصف اليوم → خروج + "نقص بصمة دخول"
     - أكثر من بصمتين → الوسطى تُسجل في الملاحظات (خروج وعودة)
  3. دمج الملاحظات المحفوظة (اجازة/مأمورية) من settings.json
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime, time

# python weekday(): الاثنين=0 ... الأحد=6
MON, TUE, WED, THU, FRI, SAT, SUN = 0, 1, 2, 3, 4, 5, 6

NOTE_MISSING_OUT = "نقص بصمة خروج"
NOTE_MISSING_IN = "نقص بصمة دخول"
NOTE_MID_RETURN = "خروج وعودة"


def fmt_time(dt: datetime) -> str:
    """تنسيق h:mm مثل 8:11 بدون ثوانٍ وبدون صفار أمام الساعة."""
    return f"{dt.hour}:{dt.minute:02d}"


def _parse_hhmm(s: str, default: time) -> time:
    try:
        parts = s.split(":")
        return time(int(parts[0]), int(parts[1]))
    except Exception:
        return default


@dataclass
class DayRow:
    emp_id: int
    template_name: str = ""
    device_name: str = ""
    d: date = None
    in_time: datetime = None
    out_time: datetime = None
    punches: int = 0
    note: str = ""            # الملاحظة النهائية المعروضة في الخلية
    sys_note: str = ""        # ملاحظة النظام (نقص بصمة / خروج وعودة)
    saved_note: str = ""      # ملاحظة محفوظة (اجازة/مأمورية)
    late: bool = False        # تجاوز وقت التأخير (اختياري)


@dataclass
class MonthAttendance:
    year: int
    month: int
    rows: dict = field(default_factory=dict)   # (emp_id, date) -> DayRow
    unmapped_ids: dict = field(default_factory=dict)  # id -> device_name (غير مربوطين)


class Processor:
    def __init__(self, settings: dict):
        self.dedup_seconds = int(settings.get("dedup_seconds", 120))
        self.midday_split = _parse_hhmm(settings.get("midday_split", "12:00"), time(12, 0))
        self.late_after = _parse_hhmm(settings.get("late_after", "08:15"), time(8, 15))
        self.late_enabled = bool(settings.get("late_enabled", False))
        self.workdays = set(int(x) for x in settings.get("workdays", [SUN, MON, TUE, WED, THU]))
        self.notes = dict(settings.get("notes", {}))
        self.employees = {e["id"]: e for e in settings.get("employees", []) if e.get("id")}

    # ---------- بناء شبكة الشهر ----------
    def month_days(self, year: int, month: int):
        """كل أيام الشهر ككائنات date."""
        import calendar
        ndays = calendar.monthrange(year, month)[1]
        return [date(year, month, d) for d in range(1, ndays + 1)]

    def workdays_set(self):
        return set(int(x) for x in self.settings_workdays()) if hasattr(self, "settings_workdays") else None

    # ---------- المعالجة ----------
    def process(self, parse_result, year: int, month: int) -> MonthAttendance:
        att = MonthAttendance(year, month)

        # 1) تجميع (موظف + تاريخ)
        grouped = {}
        name_of = {}
        for r in parse_result.records:
            if r.dt.year != year or r.dt.month != month:
                continue
            grouped.setdefault((r.emp_id, r.dt.date()), []).append(r.dt)
            name_of.setdefault(r.emp_id, r.device_name)

        # 2) اقتران + ملاحظات
        for (emp_id, d), times in sorted(grouped.items()):
            times.sort()
            row = DayRow(emp_id=emp_id, d=d, punches=len(times))
            row.device_name = name_of.get(emp_id, "")

            emp = self.employees.get(emp_id)
            row.template_name = (emp or {}).get("template_name", "")

            # 3) دمج الملاحظات المحفوظة
            key = f"{d.isoformat()}|{emp_id}"
            row.saved_note = self.notes.get(key, "")

            if len(times) >= 2:
                row.in_time = times[0]
                row.out_time = times[-1]
                middles = times[1:-1]
                if middles:
                    row.sys_note = (NOTE_MID_RETURN + ": " +
                                    " ، ".join(fmt_time(t) for t in middles))
            else:
                t = times[0]
                if t.time() < self.midday_split:
                    row.in_time = t
                    row.sys_note = NOTE_MISSING_OUT
                else:
                    row.out_time = t
                    row.sys_note = NOTE_MISSING_IN

            if self.late_enabled and row.in_time and row.in_time.time() > self.late_after:
                row.late = True

            # الملاحظة النهائية: المحفوظة أولاً، ثم ملاحظة النظام
            parts = [p for p in (row.saved_note, row.sys_note) if p]
            row.note = " / ".join(parts)

            att.rows[(emp_id, d)] = row

        # 3.5) الملاحظات المحفوظة بلا أي بصمات (مثل اجازة كاملة): تنشأ خليتها بنفسها
        for key, note in self.notes.items():
            try:
                d_str, emp_str = key.split("|")
                emp_id = int(emp_str)
                d = date.fromisoformat(d_str)
            except Exception:
                continue
            if d.year != year or d.month != month:
                continue
            if (emp_id, d) in att.rows:
                continue
            emp = self.employees.get(emp_id)
            if not emp or d.weekday() not in self.workdays:
                continue
            row = DayRow(emp_id=emp_id, d=d, punches=0)
            row.template_name = emp.get("template_name", "")
            row.saved_note = note
            row.note = note
            att.rows[(emp_id, d)] = row

        # 4) الأرقام غير المربوطة
        for emp_id, dev in name_of.items():
            if emp_id not in self.employees:
                att.unmapped_ids[emp_id] = dev

        return att

    # ---------- ملخص للعرض ----------
    def summary(self, att: MonthAttendance) -> dict:
        days = set(d for (_, d) in att.rows.keys())
        per_emp = {}
        for (emp_id, d), row in att.rows.items():
            s = per_emp.setdefault(emp_id, {"name": row.template_name or row.device_name,
                                            "days": 0, "problems": 0})
            s["days"] += 1
            if row.sys_note:
                s["problems"] += 1
        return {
            "records_in_month": len(att.rows),
            "days_with_data": len(days),
            "employees_seen": len(per_emp),
            "unmapped_ids": dict(att.unmapped_ids),
            "per_employee": per_emp,
        }
