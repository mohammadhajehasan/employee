# -*- coding: utf-8 -*-
"""
att_parser.py — محرك القراءة: قراءة ملف attlog.txt من جهاز البصمة
الصيغ المدعومة (كشف تلقائي لكل سطر):
  1) صيغة الجهاز الحقيقية ZKTeco:  رقم_الموظف <TAB> التاريخ_والوقت <TAB> حالة1 <TAB> حالة2...
  2) صيغة بأسماء:                  رقم_الموظف <TAB> الاسم <TAB> التاريخ_والوقت <TAB> حالة1 <TAB> حالة2
  3) مفصولة بفواصل (CSV) بنفس المنطقين أعلاه
  4) سطر نصي بدون TAB: id name date time s1 [s2]
- صيغ التاريخ: YYYY-MM-DD / YYYY/MM/DD / DD-MM-YYYY / DD/MM/YYYY (يوم/شهر أولاً —
  يُعكس تلقائياً إذا كان العنصر الثاني > 12) + 12 ساعة AM/PM + مكوّنات من خانة واحدة (1/9/2026)
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
    has_names: bool = False        # هل الملف يتضمن عمود أسماء؟ (False = صيغة الجهاز بدون أسماء)
    date_style: str = ""           # صيغة التاريخ المكتشفة: year-first / day-first / mixed

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

# صيغ يوم/شهر أولاً: d/m/yyyy أو d-m-yyyy (+ ثوانٍ اختيارية + AM/PM اختياري)
_DAYFIRST_RE = re.compile(
    r"^(\d{1,2})[-/](\d{1,2})[-/](\d{4})\s+"
    r"(\d{1,2}):(\d{2})(?::(\d{2}))?(?:\s*([AaPp])\.?[Mm]\.?)?$"
)


def _apply_ampm(hh: int, ap) -> int:
    if ap:
        ap = ap.lower()
        if ap == "p" and hh < 12:
            hh += 12
        elif ap == "a" and hh == 12:
            hh = 0
    return hh


def _parse_dayfirst(text: str):
    """d/m/yyyy (أو d-m-yyyy) — يُعكس تلقائياً إذا كان العنصر الثاني > 12.
    الغامضة (كلاهما ≤ 12) تُفسر يوم/شهر أولاً (المعتاد عربياً)."""
    m = _DAYFIRST_RE.match(text.strip())
    if not m:
        return None
    a, b, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    hh = _apply_ampm(int(m.group(4)), m.group(7))
    mm, ss = int(m.group(5)), int(m.group(6) or 0)
    if hh > 23 or mm > 59 or ss > 59:
        return None
    # ترتيب المحاولات: (العنصر1=يوم) أولاً، ثم العكس إذا فشل
    for d_, mo_ in ((a, b), (b, a)):
        if 1 <= mo_ <= 12 and 1 <= d_ <= 31:
            try:
                return datetime(y, mo_, d_, hh, mm, ss)
            except ValueError:
                continue
    return None


def _date_style(text: str) -> str:
    """تصنيف صيغة التاريخ للإحصاءات."""
    text = text.strip()
    if re.match(r"^\d{4}[-/]", text):
        return "year-first"
    if re.match(r"^\d{1,2}[-/]\d{1,2}[-/]\d{4}", text):
        return "day-first"
    return "unknown"


def _parse_datetime(text: str):
    text = text.strip()
    for fmt in DT_FORMATS:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return _parse_dayfirst(text)


def _split_fields(line: str):
    """يقسم السطر: TAB أولاً، ثم الفواصل إن وُجدت. يعيد (قائمة, النمط) أو (None, النمط)."""
    if "\t" in line:
        return [f.strip() for f in line.split("\t") if f.strip() != ""], "tab"
    if "," in line:
        return [f.strip() for f in line.split(",") if f.strip() != ""], "comma"
    return None, "none"


# بديل عند غياب TAB والفواصل: id name date time s1 [s2]
_NO_TAB_RE = re.compile(
    r"^\s*(\d+)\s+(.+?)\s+(\d{1,4}[-/]\d{1,2}[-/]\d{1,4})\s+"
    r"(\d{1,2}:\d{2}(?::\d{2})?(?:\s*[AaPp]\.?[Mm]\.?)?)\s+(\S+)\s*(\S*)\s*$"
)


def _parse_line(line: str, line_no: int):
    """يعيد Record أو (None, سبب). يدعم: بأسماء / بدون أسماء (صيغة الجهاز) / فواصل."""
    fields, sep = _split_fields(line)
    if fields is not None:
        if len(fields) < 2:
            return None, f"عدد الحقول ({len(fields)}) أقل من 2"
        emp_part = fields[0]
        if not emp_part.isdigit():
            return None, "رقم الموظف غير رقمي: '" + emp_part[:20] + "'"
        # كشف ذكي: إن كان الحقل الثاني تاريخاً → صيغة الجهاز بدون أسماء
        # (ZKTeco: id<TAB>datetime<TAB>s1<TAB>s2...) وإلا فالحقل الثاني اسم
        if _parse_datetime(fields[1]) is not None:
            dt_part, name = fields[1], ""
            s1 = fields[2] if len(fields) > 2 else ""
            s2 = fields[3] if len(fields) > 3 else ""
        elif len(fields) >= 3 and _parse_datetime(fields[2]) is not None:
            dt_part, name = fields[2], fields[1]
            s1 = fields[3] if len(fields) > 3 else ""
            s2 = fields[4] if len(fields) > 4 else ""
        else:
            return None, (f"لا توجد صيغة تاريخ صالحة في: "
                          f"'{fields[1][:25]}' / '{fields[2][:25] if len(fields) > 2 else ''}'")
        dt = _parse_datetime(dt_part)
        if dt is None:
            return None, f"تاريخ/وقت غير صالح: '{dt_part}'"
        return Record(int(emp_part), name, dt, s1, s2, line_no), None
    else:
        m = _NO_TAB_RE.match(line)
        if not m:
            return None, "سطر غير مطابق للصيغة (لا TAB ولا فاصلة ولا نمط نصي)"
        emp_id = int(m.group(1))
        name = m.group(2).strip()
        dt = _parse_datetime(m.group(3) + " " + m.group(4))
        if dt is None:
            return None, "تاريخ/وقت غير صالح: '" + (m.group(3) + " " + m.group(4))[:40] + "'"
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

_DATE_TOKEN_RE = re.compile(r"\d{1,4}[-/]\d{1,2}[-/]\d{1,4}")


def parse_attlog(path: str, dedup_seconds: int = 120) -> ParseResult:
    with open(path, "rb") as f:
        data = f.read()
    text, encoding = _decode(data)

    res = ParseResult(encoding=encoding)
    styles = set()
    named = 0
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
            if rec.device_name:
                named += 1
            mtok = _DATE_TOKEN_RE.search(line)
            if mtok:
                styles.add(_date_style(mtok.group(0)))

    res.records, res.duplicates_removed = _dedupe(res.records, dedup_seconds)
    res.has_names = named > 0
    if len(styles) == 1:
        res.date_style = next(iter(styles))
    elif len(styles) > 1:
        res.date_style = "mixed"
    return res
