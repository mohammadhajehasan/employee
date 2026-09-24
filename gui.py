# -*- coding: utf-8 -*-
"""
gui.py — الواجهة الرسومية (Tkinter) — المرحلة 3 + 4 + 5
التبويبات:
  1) التوليد: اختيار attlog + الشهر → معاينة ملخص → توليد Excel / CSV / تقرير جودة
  2) الربط: ربط الأرقام الوظيفية بأسماء القالب (يُحفظ دائماً في settings.json)
  3) الملاحظات: إدخال اجازة/مأمورية لكل موظف/يوم وتُدمج عند كل توليد
  4) الإعدادات: القالب، دقيقة التكرار، خط منتصف اليوم، قاعدة التأخير
"""
from __future__ import annotations

import json
import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from app import BASE_DIR, SETTINGS_PATH, load_settings, save_settings, run_pipeline
from att_parser import parse_attlog
from processor import Processor

MONTHS_AR = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو",
             "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"]


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("تفريغ البصمات تلقائياً إلى جدول الدوام")
        self.geometry("980x640")
        self.minsize(860, 560)

        self.settings = load_settings(SETTINGS_PATH)
        self.parse_result = None
        self.att = None
        self.preview_emp = None

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=8, pady=8)

        self.tab_gen = ttk.Frame(nb)
        self.tab_map = ttk.Frame(nb)
        self.tab_notes = ttk.Frame(nb)
        self.tab_cfg = ttk.Frame(nb)
        nb.add(self.tab_gen, text=" التوليد ")
        nb.add(self.tab_map, text=" الربط ")
        nb.add(self.tab_notes, text=" الملاحظات ")
        nb.add(self.tab_cfg, text=" الإعدادات ")

        self._build_gen()
        self._build_map()
        self._build_notes()
        self._build_cfg()

    # ================= 1) التوليد =================
    def _build_gen(self):
        f = self.tab_gen
        top = ttk.LabelFrame(f, text="اختيار البيانات")
        top.pack(fill="x", padx=8, pady=6)

        ttk.Label(top, text="ملف attlog:").grid(row=0, column=0, sticky="e", padx=4, pady=4)
        self.attlog_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.attlog_var, width=60).grid(row=0, column=1, padx=4)
        ttk.Button(top, text="اختيار...", command=self.pick_attlog).grid(row=0, column=2, padx=4)

        ttk.Label(top, text="الشهر:").grid(row=1, column=0, sticky="e", padx=4, pady=4)
        self.month_var = tk.StringVar(value=MONTHS_AR[8])
        ttk.Combobox(top, textvariable=self.month_var, values=MONTHS_AR,
                     state="readonly", width=10).grid(row=1, column=1, sticky="e", padx=4)
        ttk.Label(top, text="السنة:").grid(row=1, column=1, sticky="w", padx=90)
        self.year_var = tk.StringVar(value="2026")
        ttk.Spinbox(top, from_=2020, to=2100, textvariable=self.year_var, width=7).grid(row=1, column=1, padx=150)

        ttk.Button(top, text="قراءة ومعاينة", command=self.do_preview).grid(row=1, column=2, padx=4)

        mid = ttk.LabelFrame(f, text="ملخص المعاينة")
        mid.pack(fill="both", expand=True, padx=8, pady=6)
        self.preview = ttk.Treeview(mid, columns=("name", "days", "problems"), show="headings", height=12)
        self.preview.heading("name", text="الموظف")
        self.preview.heading("days", text="عدد أيام الحضور")
        self.preview.heading("problems", text="أيام بملاحظات")
        self.preview.column("name", width=380, anchor="e")
        self.preview.column("days", width=120, anchor="center")
        self.preview.column("problems", width=120, anchor="center")
        self.preview.pack(side="right", fill="both", expand=True, padx=4, pady=4)

        self.info_text = tk.Text(mid, width=42, height=12, state="disabled", wrap="word")
        self.info_text.pack(side="left", fill="both", expand=True, padx=4, pady=4)

        btns = ttk.Frame(f)
        btns.pack(fill="x", padx=8, pady=6)
        ttk.Button(btns, text="توليد Excel", command=self.do_generate).pack(side="right", padx=4)
        ttk.Button(btns, text="تصدير CSV للتحقق", command=self.do_csv).pack(side="right", padx=4)
        ttk.Button(btns, text="عرض تقرير الجودة", command=self.do_quality).pack(side="right", padx=4)

    def pick_attlog(self):
        p = filedialog.askopenfilename(title="اختر ملف attlog",
                                       filetypes=[("ملفات النص", "*.txt"), ("كل الملفات", "*.*")])
        if p:
            self.attlog_var.set(p)

    def _ym(self):
        return int(self.year_var.get()), MONTHS_AR.index(self.month_var.get()) + 1

    def do_preview(self):
        path = self.attlog_var.get().strip()
        if not os.path.exists(path):
            messagebox.showerror("خطأ", "اختر ملف attlog صحيحاً أولاً")
            return
        year, month = self._ym()
        self.parse_result = parse_attlog(path, int(self.settings.get("dedup_seconds", 120)))
        proc = Processor(self.settings)
        self.att = proc.process(self.parse_result, year, month)
        summary = proc.summary(self.att)

        for i in self.preview.get_children():
            self.preview.delete(i)
        for emp_id, s in sorted(summary["per_employee"].items()):
            self.preview.insert("", "end", values=(s["name"] or f"#{emp_id}", s["days"], s["problems"]))

        self.info_text.configure(state="normal")
        self.info_text.delete("1.0", "end")
        msg = (f"الترميز المكتشف: {self.parse_result.encoding}\n"
               f"إجمالي البصمات السليمة: {len(self.parse_result.records)}\n"
               f"أسطر تالفة استُبعدت: {len(self.parse_result.corrupt_lines)}\n"
               f"بصمات مكررة حُذفت: {self.parse_result.duplicates_removed}\n"
               f"أرقام غير مربوطة: {len(summary['unmapped_ids'])}\n\n"
               "استخدم شاشة (الربط) لربط الأرقام غير المعروفة،\n"
               "وشاشة (الملاحظات) لإدخال الاجازات والمأموريات.")
        self.info_text.insert("1.0", msg)
        self.info_text.configure(state="disabled")

    def _ensure_processed(self):
        if self.att is None:
            self.do_preview()
        return self.att is not None

    def do_generate(self):
        if not self._ensure_processed():
            return
        year, month = self._ym()
        out = filedialog.asksaveasfilename(
            title="احفظ ملف الدوام", defaultextension=".xlsx",
            initialfile=f"دوام {MONTHS_AR[month-1]} {year}.xlsx",
            filetypes=[("Excel", "*.xlsx")])
        if not out:
            return
        self._run_task(lambda: run_pipeline(
            self.attlog_var.get().strip(), year, month, out, self.settings,
            verbose=False), f"تم توليد الملف:\n{out}")

    def do_csv(self):
        if not self._ensure_processed():
            return
        year, month = self._ym()
        out = filedialog.asksaveasfilename(
            title="احفظ ملف CSV", defaultextension=".csv",
            initialfile=f"تحقق_{year}_{month:02d}.csv", filetypes=[("CSV", "*.csv")])
        if not out:
            return
        from excel_writer import export_csv
        self._run_task(lambda: export_csv(self.att, out, year, month),
                       f"تم تصدير ملف التحقق:\n{out}")

    def do_quality(self):
        if not self._ensure_processed():
            return
        from quality_report import build_report
        year, month = self._ym()
        lines = build_report(self.att, self.parse_result, self.settings, year, month)
        win = tk.Toplevel(self)
        win.title("تقرير الجودة")
        win.geometry("760x520")
        txt = tk.Text(win, wrap="word")
        txt.pack(fill="both", expand=True)
        txt.insert("1.0", "\n".join(lines))

    def _run_task(self, fn, success_msg):
        def worker():
            try:
                fn()
                self.after(0, lambda: messagebox.showinfo("تم", success_msg))
            except Exception as e:
                err = str(e)
                self.after(0, lambda: messagebox.showerror("خطأ", err))
        threading.Thread(target=worker, daemon=True).start()

    # ================= 2) الربط =================
    def _build_map(self):
        f = self.tab_map
        ttk.Label(f, text="الربط يتم مرة واحدة فقط ثم يُحفظ دائماً (عبر الرقم الوظيفي لا الاسم)").pack(anchor="e", padx=8, pady=4)

        cols = ("id", "device", "template")
        self.map_tree = ttk.Treeview(f, columns=cols, show="headings", height=14)
        self.map_tree.heading("id", text="الرقم الوظيفي")
        self.map_tree.heading("device", text="اسم الجهاز في attlog")
        self.map_tree.heading("template", text="الاسم العربي في القالب")
        self.map_tree.column("id", width=120, anchor="center")
        self.map_tree.column("device", width=220, anchor="center")
        self.map_tree.column("template", width=380, anchor="center")
        self.map_tree.pack(fill="both", expand=True, padx=8, pady=4)

        bar = ttk.Frame(f)
        bar.pack(fill="x", padx=8, pady=4)
        ttk.Label(bar, text="الاسم في القالب:").pack(side="right", padx=4)
        self.map_name_var = tk.StringVar()
        ttk.Entry(bar, textvariable=self.map_name_var, width=30).pack(side="right", padx=4)
        ttk.Button(bar, text="حفظ الربط للعنصر المحدد", command=self.save_mapping).pack(side="right", padx=4)
        ttk.Button(bar, text="تحديث القائمة من settings", command=self.refresh_mapping).pack(side="right", padx=4)
        ttk.Button(bar, text="فحص attlog لاكتشاف الأرقام", command=self.scan_ids).pack(side="right", padx=4)

    def scan_ids(self):
        path = self.attlog_var.get().strip() if hasattr(self, "attlog_var") else ""
        if path and os.path.exists(path):
            pr = parse_attlog(path, int(self.settings.get("dedup_seconds", 120)))
            for emp_id, (dev, cnt) in pr.ids_found().items():
                known = any(e.get("id") == emp_id for e in self.settings["employees"])
                if not known:
                    self.settings["employees"].append(
                        {"id": emp_id, "template_name": "", "device_name": dev, "no_punch": False})
        self.refresh_mapping()

    def refresh_mapping(self):
        for i in self.map_tree.get_children():
            self.map_tree.delete(i)
        for e in self.settings.get("employees", []):
            self.map_tree.insert("", "end", values=(e.get("id") or "—", e.get("device_name", ""),
                                                    e.get("template_name", "")))

    def save_mapping(self):
        sel = self.map_tree.selection()
        if not sel:
            messagebox.showinfo("تنبيه", "اختر صفاً من القائمة أولاً")
            return
        name = self.map_name_var.get().strip()
        if not name:
            messagebox.showinfo("تنبيه", "اكتب الاسم العربي كما يظهر في القالب")
            return
        values = self.map_tree.item(sel[0], "values")
        emp_id = values[0]
        for e in self.settings["employees"]:
            if str(e.get("id")) == str(emp_id):
                e["template_name"] = name
                break
        save_settings(self.settings, SETTINGS_PATH)
        self.refresh_mapping()
        messagebox.showinfo("تم", f"حُفظ الربط: {emp_id} ← {name}")

    # ================= 3) الملاحظات =================
    def _build_notes(self):
        f = self.tab_notes
        box = ttk.LabelFrame(f, text="إضافة ملاحظة (اجازة / مأمورية / أخرى)")
        box.pack(fill="x", padx=8, pady=6)

        ttk.Label(box, text="التاريخ (YYYY-MM-DD):").grid(row=0, column=0, padx=4, pady=6, sticky="e")
        self.note_date_var = tk.StringVar()
        ttk.Entry(box, textvariable=self.note_date_var, width=14).grid(row=0, column=1, padx=4)

        ttk.Label(box, text="الرقم الوظيفي:").grid(row=0, column=2, padx=4, sticky="e")
        self.note_id_var = tk.StringVar()
        ttk.Entry(box, textvariable=self.note_id_var, width=10).grid(row=0, column=3, padx=4)

        ttk.Label(box, text="الملاحظة:").grid(row=0, column=4, padx=4, sticky="e")
        self.note_text_var = tk.StringVar()
        ttk.Entry(box, textvariable=self.note_text_var, width=24).grid(row=0, column=5, padx=4)
        ttk.Button(box, text="إضافة", command=self.add_note).grid(row=0, column=6, padx=6)

        self.notes_tree = ttk.Treeview(f, columns=("key", "note"), show="headings", height=14)
        self.notes_tree.heading("key", text="التاريخ | الرقم الوظيفي")
        self.notes_tree.heading("note", text="الملاحظة")
        self.notes_tree.column("key", width=280, anchor="center")
        self.notes_tree.column("note", width=420, anchor="center")
        self.notes_tree.pack(fill="both", expand=True, padx=8, pady=4)

        ttk.Button(f, text="حذف الملاحظة المحددة", command=self.del_note).pack(anchor="e", padx=8, pady=4)
        self.refresh_notes()

    def refresh_notes(self):
        for i in self.notes_tree.get_children():
            self.notes_tree.delete(i)
        for key, note in sorted(self.settings.get("notes", {}).items()):
            self.notes_tree.insert("", "end", values=(key, note))

    def add_note(self):
        d = self.note_date_var.get().strip()
        emp = self.note_id_var.get().strip()
        note = self.note_text_var.get().strip()
        if not d or not emp or not note:
            messagebox.showinfo("تنبيه", "أكمل التاريخ والرقم الوظيفي والملاحظة")
            return
        self.settings.setdefault("notes", {})[f"{d}|{emp}"] = note
        save_settings(self.settings, SETTINGS_PATH)
        self.refresh_notes()
        self.note_text_var.set("")

    def del_note(self):
        sel = self.notes_tree.selection()
        if not sel:
            return
        key = self.notes_tree.item(sel[0], "values")[0]
        self.settings.get("notes", {}).pop(key, None)
        save_settings(self.settings, SETTINGS_PATH)
        self.refresh_notes()

    # ================= 4) الإعدادات =================
    def _build_cfg(self):
        f = self.tab_cfg
        box = ttk.LabelFrame(f, text="القالب والإعدادات العامة")
        box.pack(fill="x", padx=8, pady=6)

        ttk.Label(box, text="ملف القالب الأصلي (اختياري):").grid(row=0, column=0, sticky="e", padx=4, pady=4)
        self.tpl_var = tk.StringVar(value=self.settings.get("template_path", ""))
        ttk.Entry(box, textvariable=self.tpl_var, width=55).grid(row=0, column=1, padx=4)
        ttk.Button(box, text="اختيار...", command=self.pick_template).grid(row=0, column=2, padx=4)

        ttk.Label(box, text="حد التكرار (ثانية):").grid(row=1, column=0, sticky="e", padx=4, pady=4)
        self.dedup_var = tk.StringVar(value=str(self.settings.get("dedup_seconds", 120)))
        ttk.Entry(box, textvariable=self.dedup_var, width=8).grid(row=1, column=1, sticky="e", padx=4)

        ttk.Label(box, text="خط منتصف اليوم (لنقص البصمة):").grid(row=2, column=0, sticky="e", padx=4, pady=4)
        self.split_var = tk.StringVar(value=self.settings.get("midday_split", "12:00"))
        ttk.Entry(box, textvariable=self.split_var, width=8).grid(row=2, column=1, sticky="e", padx=4)

        ttk.Label(box, text="تمييز التأخير بعد:").grid(row=3, column=0, sticky="e", padx=4, pady=4)
        self.late_var = tk.StringVar(value=self.settings.get("late_after", "08:15"))
        ttk.Entry(box, textvariable=self.late_var, width=8).grid(row=3, column=1, sticky="e", padx=4)
        self.late_enabled = tk.BooleanVar(value=bool(self.settings.get("late_enabled", False)))
        ttk.Checkbutton(box, text="تفعيل قاعدة التأخير", variable=self.late_enabled).grid(row=3, column=2, padx=4)

        ttk.Button(box, text="حفظ الإعدادات", command=self.save_cfg).grid(row=4, column=1, sticky="e", padx=4, pady=8)

    def pick_template(self):
        p = filedialog.askopenfilename(title="اختر ملف القالب الأصلي", filetypes=[("Excel", "*.xlsx")])
        if p:
            self.tpl_var.set(p)

    def save_cfg(self):
        try:
            self.settings["dedup_seconds"] = int(self.dedup_var.get())
        except ValueError:
            pass
        self.settings["midday_split"] = self.split_var.get().strip()
        self.settings["late_after"] = self.late_var.get().strip()
        self.settings["late_enabled"] = bool(self.late_enabled.get())
        self.settings["template_path"] = self.tpl_var.get().strip()
        save_settings(self.settings, SETTINGS_PATH)
        messagebox.showinfo("تم", "حُفظت الإعدادات")


def launch():
    App().mainloop()


if __name__ == "__main__":
    launch()
