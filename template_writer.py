# -*- coding: utf-8 -*-
"""
template_writer.py — وضع القالب (الأدق): يعبّئ نسخة من ملف قالبك الأصلي نفسه
مما يضمن مطابقة التنسيق 100% (ألوان، دمج، حدود، RTL) دون إعادة بناء.

الصيغ المدعومة لتحديد أعمدة الأيام:
  1) خلايا تواريخ حقيقية (datetime/date) لشهر الهدف — مدموجة أو مفردة
  2) تواريخ نصية شائعة: "2026-09-15"، "15/09/2026"، "15.09.2026"، "15 سبتمبر"، "15 أيلول"
  3) صفوف أرقام الأيام (1..31) كأرقام أو نصوص — الشائع في القوالب العربية،
     حتى بأرقام هندية (٠١٢٣) وعلى صفوف أسبوعية متعددة

آلية العمل:
  1) مسح القالب للعثور على أعمدة كل يوم من أيام الشهر الهدف
  2) تحديد أعمدة الدخول/الخروج/الملاحظات عبر الدمج أو رؤوس "دخول/حضور/خروج/انصراف/ملاحظات"
  3) مطابقة أسماء الموظفين (بمقارنة عربية مطبّعة) لتحديد صف كل موظف
  4) كتابة القيم في الخلايا الصحيحة فقط — كل التنسيق يبقى كما هو
"""
from __future__ import annotations

import re
from datetime import date, datetime

from openpyxl import load_workbook

from processor import fmt_time

AR_DIAC = re.compile(r"[\u064B-\u0652\u0640]")  # تشكيل + تطويل
AR_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")

# أسماء الأشهر: الميلادية + الشامية (أيلول = سبتمبر ...)
MONTH_WORDS = {
    "يناير": 1, "كانون الثاني": 1, "فبراير": 2, "شباط": 2,
    "مارس": 3, "آذار": 3, "اذار": 3, "أبريل": 4, "ابريل": 4, "نيسان": 4,
    "مايو": 5, "أيار": 5, "ايار": 5, "يونيو": 6, "حزيران": 6,
    "يوليو": 7, "تموز": 7, "أغسطس": 8, "اغسطس": 8, "آب": 8,
    "سبتمبر": 9, "أيلول": 9, "ايلول": 9,
    "أكتوبر": 10, "اكتوبر": 10, "تشرين الاول": 10, "تشرين الأول": 10,
    "نوفمبر": 11, "تشرين الثاني": 11, "ديسمبر": 12, "كانون الاول": 12, "كانون الأول": 12,
}

# مرادفات رؤوس الأعمدة تحت كل يوم
LABEL_SYNONYMS = {
    "in": ("دخول", "حضور", "بداية", "الدخول", "الحضور"),
    "out": ("خروج", "انصراف", "انصراف ", "نهاية", "الخروج", "الانصراف"),
    "note": ("ملاحظات", "ملاحظة", "ملحوظة", "الملاحظات"),
}

DAY_LABEL_ROW_SPAN = 4      # البحث عن الرؤوس حتى 4 صفوف تحت خلية اليوم
MIN_DAY_CELLS = 4           # أقل عدد أرقام أيام لنعتبر الصف رأس أيام


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


def _word_month(word: str):
    """إيجاد رقم الشهر من كلمة/كلمات عربية (يدعم تشرين الأول/الثاني)."""
    w = normalize_ar(word)
    if not w:
        return None
    if w in MONTH_WORDS:
        return MONTH_WORDS[w]
    if "تشرين" in w:
        return 10 if ("اول" in w or "أول" in w) else (11 if "ثاني" in w else None)
    if "كانون" in w:
        return 1 if ("اول" in w or "أول" in w) else (12 if "ثاني" in w else None)
    # أول كلمة طويلة مطابقة
    for name, m in MONTH_WORDS.items():
        n = normalize_ar(name)
        if n and (n in w or w in n) and abs(len(n) - len(w)) <= 2:
            return m
    return None


def _to_int(v):
    """قراءة عدد صحيح من رقم/نص (يدعم الأرقام الهندية ٠١٢٣)."""
    if isinstance(v, bool):
        return None
    if isinstance(v, int):
        return v
    if isinstance(v, float):
        return int(v) if v == int(v) else None
    if isinstance(v, str):
        s = v.strip().translate(AR_DIGITS)
        if s.isdigit() and len(s) <= 2:
            return int(s)
    return None


def _parse_text_date(v, year: int, month: int):
    """محاولة قراءة تاريخ من نص. يعيد date إذا كان للشهر الهدف أو لعام الهدف."""
    if not isinstance(v, str):
        return None
    s = v.strip().translate(AR_DIGITS)
    if not s or len(s) > 40 or s.isdigit():
        return None
    s = re.sub(r"^(ال)?(سبت|احد|أحد|اثنين|إثنين|ثلاثاء|أربعاء|اربعاء|خميس|جمعة)[\s،,\-]+", "", s)
    # صيغ كاملة
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%Y/%m/%d",
                "%d.%m.%Y", "%d-%m-%y", "%d/%m/%y", "%d %m %Y"):
        try:
            dt = datetime.strptime(s, fmt)
            return dt.date()
        except ValueError:
            pass
    # "15 سبتمبر" أو "15 سبتمبر 2026"
    m = re.match(r"^(\d{1,2})\s+([\u0621-\u064A][\u0621-\u064A ]*?)(?:\s+(\d{4}))?$", s)
    if m:
        day = int(m.group(1))
        mon = _word_month(m.group(2))
        yr = int(m.group(3)) if m.group(3) else year
        if mon and 1 <= day <= 31:
            try:
                return date(yr, mon, day)
            except ValueError:
                return None
    return None


class TemplateScanError(Exception):
    pass


class TemplateScan:
    def __init__(self):
        self.date_anchors = {}      # date -> dict(col_start, col_end, sub_map, row, combined)
        self.emp_rows = {}          # emp_id -> excel_row
        self.unmatched_names = []   # أسماء في القالب لم تُربط
        self.messages = []
        self.mode = ""


def _merge_span(ws, row, col):
    """امتداد الدمج الأفقي للخلية، وإلا (col, col)."""
    try:
        for mr in ws.merged_cells.ranges:
            if mr.min_row <= row <= mr.max_row and mr.min_col <= col <= mr.max_col:
                return mr.min_col, mr.max_col
    except Exception:
        pass
    return col, col


def _is_title_cell(ws, cell) -> bool:
    """خلية مدموجة عرضياً على أكثر من 3 أعمدة = عنوان وليست اسماً."""
    c0, c1 = _merge_span(ws, cell.row, cell.column)
    return (c1 - c0 + 1) > 3


def _find_labels(ws, row, col_start, col_end):
    """البحث عن رؤوس دخول/خروج/ملاحظات تحت مجموعة يوم (حتى 4 صفوف)."""
    found = {}
    for rr in range(row + 1, min(row + 1 + DAY_LABEL_ROW_SPAN, ws.max_row + 1)):
        for cc in range(col_start, min(col_end + 4, ws.max_column + 1)):
            v = ws.cell(row=rr, column=cc).value
            if not isinstance(v, str):
                continue
            nv = normalize_ar(v)
            if not nv or len(nv) > 30:
                continue
            for key, words in LABEL_SYNONYMS.items():
                if key not in found and any(w in nv for w in words):
                    found[key] = cc
        if "in" in found and "out" in found:
            break
    return found


def _sub_map_for(ws, row, col_start, col_end):
    """تحديد أعمدة in/out/note لمجموعة يوم واحدة."""
    width = col_end - col_start + 1
    labels = _find_labels(ws, row, col_start, col_end)
    if "in" in labels and "out" in labels:
        return {"in": labels["in"], "out": labels["out"],
                "note": labels.get("note", labels["out"] + 1), "combined": False}
    if width == 1:
        return {"in": col_start, "out": col_start, "combined": True}
    if width == 2:
        return {"in": col_start, "out": col_start + 1, "combined": True}
    return {"in": col_start, "out": col_start + 1,
            "note": labels.get("note", col_start + 2), "combined": False}


def _day_cells_of_row(row):
    """كل الخلايا الرقمية (1..31) في صف مع أعمدتها."""
    cells = []
    for cell in row:
        d = _to_int(cell.value)
        if d is not None and 1 <= d <= 31:
            cells.append((cell.column, d))
    return cells


def _increasing(cells):
    return all(cells[i + 1][1] > cells[i][1] for i in range(len(cells) - 1))


def _find_day_header_rows(ws):
    """كل الصفوف وأرقام أيامها (للفحص والاستمرار)."""
    rows = []
    for row in ws.iter_rows():
        cells = _day_cells_of_row(row)
        rows.append((row[0].row, cells))
    return rows


def _anchors_from_day_rows(ws, chosen, year: int, month: int, seen=None):
    """يبني مراسلات التواريخ من صفوف أرقام أيام مختارة."""
    seen = seen if seen is not None else set()
    anchors = {}
    for (rnum, cells) in chosen:
        for i, (col, dnum) in enumerate(cells):
            if dnum in seen:
                continue
            nxt = cells[i + 1][0] if i + 1 < len(cells) else None
            c0, c1 = _merge_span(ws, rnum, col)
            if c1 == c0 and nxt:
                c1 = min(c0 + max(nxt - col, 1) - 1, c0 + 2)
            try:
                d = date(year, month, dnum)
            except ValueError:
                continue
            seen.add(dnum)
            anchors[d] = {"row": rnum, "col_start": c0, "col_end": c1,
                          "sub_map": _sub_map_for(ws, rnum, c0, c1)}
    return anchors


def _collect_anchors(ws, year: int, month: int):
    """يجمع مراسلات (تاريخ → مجموعة أعمدة) بالصيغ الثلاث. يعيد (dict, mode, messages)."""
    messages = []
    anchors = {}

    # ---------- 1) تواريخ حقيقية / نصية ----------
    other_months = set()
    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            d = None
            if isinstance(v, datetime):
                d = v.date()
            elif isinstance(v, date):
                d = v
            if d is None and isinstance(v, str):
                d = _parse_text_date(v, year, month)
            if d is None:
                continue
            if d.year == year and d.month == month:
                c0, c1 = _merge_span(ws, cell.row, cell.column)
                anchors.setdefault(d, {
                    "row": cell.row, "col_start": c0, "col_end": c1,
                    "sub_map": _sub_map_for(ws, cell.row, c0, c1)})
            else:
                other_months.add(f"{d.year}-{d.month:02d}")

    if anchors:
        messages.append(f"تم العثور على {len(anchors)} يوم لشهر الهدف في القالب.")
        return anchors, "تعبئة نسخة القالب مباشرة (تواريخ حقيقية)", messages

    # ---------- 2) صفوف أرقام الأيام ----------
    day_rows = _find_day_header_rows(ws)
    qualifying = [(r, cells) for (r, cells) in day_rows
                  if len(cells) >= MIN_DAY_CELLS and _increasing(cells)]
    if qualifying:
        # صف كامل الشهر (≥15 رقماً)؟ خذ الأفضل. وإلا اجمع الصفوف الأسبوعية.
        full = [rc for rc in qualifying if len(rc[1]) >= 15]
        if full:
            full.sort(key=lambda x: (-len(x[1]), x[0]))
            chosen = [full[0]]
        else:
            chosen = qualifying
        seen = set()
        anchors = _anchors_from_day_rows(ws, chosen, year, month, seen)
        # استكمال: صفوف لاحقة قصيرة (آخر أسبوع مثل 29، 30)
        if anchors:
            last_row = max(info["row"] for info in anchors.values())
            max_day = max(d.day for d in anchors)
            cont = [(r, cells) for (r, cells) in day_rows
                    if r > last_row and cells and _increasing(cells)
                    and max(d for _, d in cells) > max_day]
            if cont:
                extra = _anchors_from_day_rows(ws, cont, year, month, seen)
                for d, info in extra.items():
                    if d.day > max_day and d not in anchors:
                        anchors[d] = info
                if max(anchors).day > max_day:
                    messages.append(f"تم استكمال الأيام من صفوف لاحقة حتى اليوم {max(anchors).day}.")
        if anchors:
            rows_desc = "، ".join(str(r) for r, _ in chosen[:6])
            messages.append(f"تم اكتشاف أرقام أيام الشهر ({len(anchors)} يوماً) في الصف: {rows_desc}.")
            return anchors, "تعبئة نسخة القالب مباشرة (صف أرقام الأيام)", messages

    # ---------- لا شيء ----------
    if other_months:
        mts = "، ".join(sorted(other_months)[:6])
        messages.append(f"القالب يحتوي تواريخ لأشهر أخرى ({mts}) وليس الشهر المطلوب.")
    return anchors, "", messages


def _no_anchor_error(ws, year, month, extra_msgs):
    msg = ("لم أتعرف على أعمدة الأيام في القالب. الصيغ المدعومة: تواريخ حقيقية "
           "لشهر الهدف، أو تواريخ نصية مثل 15/09/2026 أو «15 سبتمبر»، "
           "أو صف أرقام الأيام 1..31. ")
    if extra_msgs:
        msg += " ".join(extra_msgs) + " "
    msg += "يمكنك دائماً التوليد بالمولّد المدمج المطابق لبنية القالب."
    return TemplateScanError(msg)


def scan_template(wb, year: int, month: int) -> TemplateScan:
    scan = TemplateScan()
    ws = wb.active if wb.active is not None else wb.worksheets[0]

    anchors, mode, messages = _collect_anchors(ws, year, month)
    scan.mode = mode
    scan.messages.extend(messages)
    if not anchors:
        raise _no_anchor_error(ws, year, month, messages)
    scan.date_anchors = anchors
    return scan


def match_employee_rows(wb, scan: TemplateScan, employees: list):
    """يربط أسماء settings بصفوف القالب عبر المقارنة المطبّعة."""
    ws = wb.active if wb.active is not None else wb.worksheets[0]
    name_cells = []
    for row in ws.iter_rows(min_col=1, max_col=2):
        for cell in row:
            v = cell.value
            if isinstance(v, str) and 2 < len(v.strip()) <= 60 and not _is_title_cell(ws, cell):
                # استبعد رؤوس الأعمدة مثل «الاسم/الموظف»
                nv = normalize_ar(v)
                if nv in ("الاسم", "الموظف", "اسم الموظف", "الاسم والرقم", "م", "ت"):
                    continue
                name_cells.append((cell.row, nv, v))

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
        scan.messages.append("أسماء لم تُعثر على صف لها: " + " ، ".join(scan.unmatched_names[:12]))
    return scan


def analyze_template(template_path: str, year: int, month: int) -> dict:
    """تحليل قالب بدون تعبئة — للمعاينة في الواجهة (أسماء/أيام/بنية/وضع)."""
    try:
        keep_vba = str(template_path).lower().endswith(".xlsm")
        wb = load_workbook(template_path, read_only=False, keep_vba=keep_vba)
    except Exception as e:
        raw = str(e)
        if "zip" in raw.lower():
            raw = ("الملف ليس بتنسيق xlsx الحديث (قد يكون xls قديم أو تالفاً). "
                   "افتحه في Excel ثم: ملف ← حفظ باسم ← «Excel Workbook (*.xlsx)» وأعد اختياره.")
        elif "Permission" in raw or "Errno 13" in raw:
            raw = "الملف مفتوح في Excel — أغلقه ثم أعد المحاولة."
        return {"ok": False, "error": f"تعذر فتح القالب: {raw}"}
    try:
        ws = wb.active if wb.active is not None else wb.worksheets[0]
        info = {"ok": True, "sheet": ws.title,
                "dimensions": ws.dimensions,
                "merged_count": len(ws.merged_cells.ranges)}
        anchors, mode, messages = _collect_anchors(ws, year, month)
        info["month_days_found"] = sorted(d.day for d in anchors)
        info["messages"] = messages
        names = []
        for row in ws.iter_rows(min_col=1, max_col=2):
            for cell in row:
                v = cell.value
                if isinstance(v, str) and 2 < len(v.strip()) <= 60 and not _is_title_cell(ws, cell):
                    nv = normalize_ar(v)
                    if nv in ("الاسم", "الموظف", "اسم الموظف"):
                        continue
                    names.append(v.strip())
        # إزالة التكرار مع الحفاظ على الترتيب
        seen = set(); uniq = []
        for n in names:
            k = normalize_ar(n)
            if k not in seen:
                seen.add(k); uniq.append(n)
        info["names"] = uniq[:80]
        info["names_count"] = len(uniq)
        if anchors:
            info["mode"] = "✅ " + mode + " — ستُعبّأ نسخة من قالبكم نفسه بكل تنسيقاته"
        else:
            info["warning"] = ("لم أتعرف على أعمدة الأيام في القالب — سيُستخدم المولّد "
                               "المدمج المطابق لبنية القالب. " + " ".join(messages))
        return info
    except Exception as e:
        return {"ok": False, "error": f"تعذر تحليل القالب: {type(e).__name__}: {e}"}
    finally:
        wb.close()


def write_month_to_template(template_path: str, out_path: str, att,
                            employees: list, year: int, month: int,
                            clear_existing: bool = True):
    """يفتح نسخة من القالب، يعبّئه، ويحفظه باسم جديد. يعيد تقرير المسح."""
    keep_vba = str(template_path).lower().endswith(".xlsm")
    wb = load_workbook(template_path, keep_vba=keep_vba)
    scan = scan_template(wb, year, month)
    match_employee_rows(wb, scan, employees)
    ws = wb.active if wb.active is not None else wb.worksheets[0]

    written = 0
    combined_cells = 0
    for (emp_id, d), r in att.rows.items():
        info = scan.date_anchors.get(d)
        row = scan.emp_rows.get(emp_id)
        if not info or not row:
            continue  # يظهر في تقرير الجودة
        sm = info["sub_map"]
        if sm.get("combined"):
            # عمود واحد لكل يوم: اجمع الدخول والخروج في خلية واحدة
            txt = " - ".join(t for t in (fmt_time(r.in_time) if r.in_time else "",
                                         fmt_time(r.out_time) if r.out_time else "") if t)
            if txt:
                ws.cell(row=row, column=sm["in"], value=txt)
                written += 1
                combined_cells += 1
            if r.note and sm.get("note"):
                ws.cell(row=row, column=sm["note"], value=r.note)
                written += 1
            continue
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
    scan.messages.append(f"تمت كتابة {written} خلية في نسخة القالب."
                         + (f" ({combined_cells} خلية مدمجة دخول-خروج)" if combined_cells else ""))
    return scan
