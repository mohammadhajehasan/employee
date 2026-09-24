# -*- coding: utf-8 -*-
"""
att_parser.py — محرك القراءة: قراءة ملف attlog.txt من جهاز البصمة
الصيغة: رقم_الموظف <TAB> الاسم <TAB> التاريخ والوقت <TAB> حالة1 <TAB> حالة2
- كشف الترميز تلقائياً (UTF-8 / UTF-8-BOM / UTF-16 / ANSI-CP1256)
- استبعاد الأسطر التالفة مع تسجيل رقم السطر وسبب الاستبعاد
- إزالة التكرار: بصمتان لنفس الموظف بفارق أقل من دقيقتين = بصمة واحدة (قابلة للتعديل)
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime

# ---------- نماذج البيانات ----------

@dataclass
class Record:
    emp_id: int
    device_name: str
    dt: datetime
    status1: str = ""
    status2: str = ""
    line_no: int = 0


@dataclass
class ParseResult:
    records: list = field(default_factory=list)          # list[Record] بعد إزالة التكرار
    corrupt_lines: list = field(default_factory=list)    # list[(line_no, text, reason)]
    encoding: str = ""
    duplicates_removed: int = 0
    total_lines: int = 0

    def ids_found(self):
        """الأرقام الوظيفية المكتشفة في الملف مع اسم الجهاز الأكثر تكراراً."""
        names = {}
        for r in self.records:
            cur, cnt = names.get(r.emp_id, ("", 0))
            names[r.emp_id] = (r.device_name, cnt + 1)
        return names  # {id: (device_name, count)}


# ---------- كشف الترميز ----------

def _decode(data: bytes) -> tuple:
    """يحاول فك الترميز بالترتيب ويعيد (النص, اسم الترميز)."""
    # BOM checks
    if data.startswith(b"\xff\xfe") or data.startswith(b"\xfe\xff"):
        try:
            return data.decode("utf-16"), "utf-16"
        except UnicodeDecodeError:
            pass
    if data.startswith(b"\xef\xbb\xbf"):
        try:
            return data.decode("utf-8-sig"), "utf-8-sig"
        except UnicodeDecodeError:
            pass
    for enc in ("utf-8", "cp1256", "latin-1"):
        try:
            return data.decode(enc), enc
        except UnicodeDecodeError:
            continue
    return data.decode("latin-1", errors="replace"), "latin-1(replace)"


# ---------- تحليل الأسطر ----------

DT_FORMATS = ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M",
              "%Y/%m/%d %H:%M:%S", "%Y/%m/%d %H:%M",
              "%Y-%m-%d %I:%M:%S %p", "%Y/%m/%d %I:%M %p")

# بديل عند غياب TAB: id name date time s1 [s2]
_NO_TAB_RE = re.compile(
    r"^\s*(\d+)\s+(.+?)\s+(\d{4}[-/]\d{1,2}[-/]\d{1,2})\s+"
    r"(\d{1,2}:\d{2}(?::\d{2})?(?:\s*[AaPp][Mm])?)\s+(\S+)\s*(\S*)\s*$"
)


def _parse_datetime(text: str):
    text = text.strip()
    for fmt in DT_FORMATS:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def _parse_line(line: str, line_no: int):
    """يعيد Record أو (None, سبب)."""
    if "\t" in line:
        fields = [f.strip() for f in line.split("\t") if f.strip() != ""]
        if len(fields) < 3:
            return None, "عدد الحقول أقل من 3"
        emp_part, name_part, dt_part = fields[0], fields[1], fields[2]
        s1 = fields[3] if len(fields) > 3 else ""
        s2 = fields[4] if len(fields) > 4 else ""
        if not emp_part.isdigit():
            return None, "رقم الموظف غير رقمي"
        dt = _parse_datetime(dt_part)
        if dt is None:
            return None, f"تاريخ/وقت غير صالح: '{dt_part}'"
        return Record(int(emp_part), name_part, dt, s1, s2, line_no), None
    else:
        m = _NO_TAB_RE.match(line)
        if not m:
            return None, "سطر غير مطابق للصيغة (لا TAB ولا نمط نصي)"
        emp_id = int(m.group(1))
        name = m.group(2).strip()
        dt = _parse_datetime(m.group(3) + " " + m.group(4))
        if dt is None:
            return None, "تاريخ/وقت غير صالح"
        return Record(emp_id, name, dt, m.group(5), m.group(6), line_no), None


# ---------- إزالة التكرار ----------

def _dedupe(records: list, dedup_seconds: int):
    """بصمتان متتاليتان لنفس الموظف بفارق أقل من dedup_seconds = بصمة واحدة."""
    records.sort(key=lambda r: (r.emp_id, r.dt))
    out = []
    removed = 0
    for r in records:
        if out and out[-1].emp_id == r.emp_id:
            gap = (r.dt - out[-1].dt).total_seconds()
            if 0 <= gap < dedup_seconds:
                removed += 1
                continue
        out.append(r)
    return out, removed


# ---------- الواجهة الرئيسية ----------

def parse_attlog(path: str, dedup_seconds: int = 120) -> ParseResult:
    with open(path, "rb") as f:
        data = f.read()
    text, encoding = _decode(data)

    res = ParseResult(encoding=encoding)
    for line_no, raw in enumerate(text.splitlines(), start=1):
        res.total_lines += 1
        line = raw.strip("\ufeff").strip()
        if not line:
            continue
        rec, err = _parse_line(line, line_no)
        if rec is None:
            res.corrupt_lines.append((line_no, line[:80], err))
        else:
            res.records.append(rec)

    res.records, res.duplicates_removed = _dedupe(res.records, dedup_seconds)
    return res
