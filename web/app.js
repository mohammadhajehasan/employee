"use strict";
/* ============================================================
   نظام تفريغ البصمات — منطق الواجهة v2.0
   ثلاث طبقات اتصال بنفس الـ API:
   1) pywebview  — نافذة سطح مكتب أصلية
   2) HTTP       — وضع المتصفح الاحتياطي
   3) Demo       — وضع عرض تجريبي ببيانات محاكاة
   ============================================================ */

const $  = (s, r) => (r || document).querySelector(s);
const $$ = (s, r) => [...(r || document).querySelectorAll(s)];

const MONTHS_AR = ["يناير","فبراير","مارس","أبريل","مايو","يونيو","يوليو","أغسطس","سبتمبر","أكتوبر","نوفمبر","ديسمبر"];

const state = {
  settings: null,
  attlogPath: "", attlogName: "",
  tplPath: "",
  preview: null,
  qualityLines: [],
  noteKind: "اجازة",
  ready: false,
};

/* ---------------- طبقة الاتصال ---------------- */

const isWebview = () => !!(window.pywebview && window.pywebview.api);
const isServerMode = () => !isWebview() && location.protocol.startsWith("http");
const isDemo = () => !isWebview() && !isServerMode();

async function httpApi(fn, payload) {
  const res = await fetch("/api/" + fn, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload || {}),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok || data.ok === false) throw new Error(data.error || ("HTTP " + res.status));
  return data;
}

let FORCE_DEMO = false;

async function callApi(fn, payload) {
  if (!FORCE_DEMO) {
    if (isWebview()) {
      const r = await window.pywebview.api[fn](payload || {});
      if (r && r.ok === false) throw new Error(r.error || "خطأ غير معروف");
      return r;
    }
    if (isServerMode()) {
      try { return await httpApi(fn, payload); }
      catch (e) {
        // على استضافة ثابتة (GitHub Pages) لا يوجد محرك — ننتقل للوضع التجريبي
        if (fn === "app_info" || fn === "get_state") FORCE_DEMO = true;
        else throw e;
      }
    }
  }
  return await demoApi(fn, payload || {});
}

function waitWebview(ms) {
  return new Promise(resolve => {
    const t0 = Date.now();
    (function poll() {
      if (window.pywebview && window.pywebview.api) return resolve(true);
      if (Date.now() - t0 > (ms || 2500)) return resolve(false);
      setTimeout(poll, 100);
    })();
  });
}

/* ---------------- وضع العرض التجريبي ---------------- */

const DEMO = {
  employees: [
    { id: 1025, template_name: "يزن أبو بايع",     device_name: "Yazeed",   no_punch: false },
    { id: 1000, template_name: "دانيامحموض",       device_name: "Dania ma", no_punch: false },
    { id: 1004, template_name: "حازم حط",          device_name: "Hazem Go", no_punch: false },
    { id: 1006, template_name: "براءه بطرني",      device_name: "Baraa B",  no_punch: false },
    { id: 1017, template_name: "محمد طارق كركش",   device_name: "M.Tareq",  no_punch: false },
    { id: 1003, template_name: "",                 device_name: "A.Diab",   no_punch: false },
    { id: 1022, template_name: "",                 device_name: "Sallam",   no_punch: false },
  ],
  notes: { "2026-09-20|1017": "اجازة", "2026-09-06|1016": "مأمورية" },
  workdays: [6, 0, 1, 2, 3],
  perEmp: {
    1025: { days: 22, problems: 2 }, 1000: { days: 21, problems: 1 },
    1004: { days: 22, problems: 0 }, 1006: { days: 20, problems: 3 },
    1017: { days: 21, problems: 1 }, 1003: { days: 18, problems: 2 },
    1022: { days: 9,  problems: 1 },
  },
};

const sleep = ms => new Promise(r => setTimeout(r, ms));

async function demoApi(fn, p) {
  switch (fn) {
    case "app_info":
      return { ok: true, version: "2.0", mode: "وضع العرض التجريبي", settings_path: "C:\\Attendance\\settings.json" };

    case "get_state":
      return { ok: true, settings: { version: 1, template_path: "", dedup_seconds: 120, midday_split: "12:00", late_after: "08:15", late_enabled: false, workdays: DEMO.workdays, employees: DEMO.employees, notes: DEMO.notes } };

    case "save_settings":
      await sleep(300);
      return { ok: true };

    case "select_file": {
      await sleep(250);
      if (p.kind === "attlog") return { ok: true, path: "C:\\Attendance\\attlog.txt", name: "attlog.txt" };
      return { ok: true, path: "C:\\Attendance\\دوام_قالب.xlsx", name: "دوام_قالب.xlsx" };
    }

    case "preview": {
      await sleep(900);
      const employees = Object.entries(DEMO.perEmp).map(([id, s]) => {
        const e = DEMO.employees.find(x => x.id === +id) || {};
        return { id: +id, name: e.template_name || e.device_name || ("#" + id), days: s.days, problems: s.problems };
      });
      return { ok: true, stats: { records: 287, corrupt: 2, duplicates: 4, encoding: "utf-8", unmapped: { 1023: "Khaled", 1031: "Omar" }, employees } };
    }

    case "generate": {
      await sleep(1500);
      const isCsv = p.kind === "csv";
      return {
        ok: true,
        path: isCsv ? "C:\\Attendance\\output\\تحقق_2026_09.csv" : "C:\\Attendance\\output\\دوام_سبتمبر_2026.xlsx",
        files: isCsv
          ? ["C:\\Attendance\\output\\تحقق_2026_09.csv"]
          : ["C:\\Attendance\\output\\دوام_سبتمبر_2026.xlsx"],
        logs: isCsv
          ? ["تصدير CSV للتحقق اليدوي: 212 سطر"]
          : ["[1/4] قراءة attlog: 287 بصمة سليمة | 2 سطر تالف | 4 مكرر محذوف", "[2/4] المعالجة: 133 يوم/موظف | 2 رقم غير مربوط", "[3/4] Excel (وضع القالب): دوام_سبتمبر_2026.xlsx", "[4/4] تقرير الجودة جاهز"],
      };
    }

    case "quality":
      await sleep(700);
      return { ok: true, lines: [
        "== أرقام وظيفية في attlog غير مربوطة بأسماء القالب ==",
        "  - الرقم 1023 (اسم الجهاز: Khaled)",
        "  - الرقم 1031 (اسم الجهاز: Omar)", "",
        "== أيام بنقص بصمة ==",
        "  - 2026-09-03: براءه بطرني ← نقص بصمة خروج (دخول)",
        "  - 2026-09-15: A.Diab ← نقص بصمة دخول (خروج)", "",
        "== أيام فيها خروج وعودة (أكثر من بصمتين) ==",
        "  - 2026-09-10: يزن أبو بايع ← خروج وعودة: 12:31", "",
        "== موظفون في القالب بلا أي بصمة طوال الشهر ==",
        "  - مرام الراعي", "", "== بصمات مكررة حُذفت (فارق أقل من دقيقتين): 4 ==",
      ] };

    case "open_path":
      toast("في وضع العرض: " + (p.path || ""), "warn");
      return { ok: true };

    case "scan_ids":
      await sleep(700);
      return { ok: true, added: 2 };

    default:
      return { ok: true };
  }
}

/* ---------------- أدوات واجهة ---------------- */

const ICON_OK = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6 9 17l-5-5"/></svg>';
const ICON_ERR = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><circle cx="12" cy="12" r="10"/><path d="M15 9l-6 6M9 9l6 6"/></svg>';
const ICON_WARN = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0z"/><path d="M12 9v4M12 17h.01"/></svg>';
const ICON_FILE = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/></svg>';

function toast(msg, kind) {
  const el = document.createElement("div");
  el.className = "toast " + (kind === "err" ? "err" : kind === "warn" ? "warn" : "");
  el.innerHTML = (kind === "err" ? ICON_ERR : kind === "warn" ? ICON_WARN : ICON_OK) + "<span></span>";
  el.querySelector("span").textContent = msg;
  $("#toasts").appendChild(el);
  setTimeout(() => { el.classList.add("out"); setTimeout(() => el.remove(), 300); }, 3600);
}

function overlay(show, text) {
  $("#overlayText").textContent = text || "جارٍ المعالجة…";
  $("#overlay").classList.toggle("hidden", !show);
}

async function busy(fn, text) {
  overlay(true, text);
  try { return await fn(); }
  catch (e) { console.error(e); toast(String(e.message || e), "err"); }
  finally { overlay(false); }
}

/* ---------------- التنقل والثيم ---------------- */

const VIEW_TITLES = {
  generate: ["توليد جدول الدوام", "اختر ملف بصمات attlog ثم اقرأ المعاينة وولّد ملف Excel النهائي"],
  mapping:  ["ربط الموظفين", "اربط كل رقم وظيفي في الجهاز باسمه كما يظهر في قالب Excel"],
  notes:    ["الملاحظات الدائمة", "اجازات ومأموريات تُدمج تلقائياً في كل عملية توليد"],
  settings: ["الإعدادات", "القالب وقواعد المعالجة وأيام الدوام الأسبوعية"],
  about:    ["دليل الاستخدام", "خطوات العمل والخوارزمية المعتمدة ومعلومات النظام"],
};

function goto(view) {
  state.view = view;
  $$(".nav-item").forEach(b => b.classList.toggle("active", b.dataset.view === view));
  $$(".view").forEach(v => v.classList.toggle("active", v.id === "view-" + view));
  const [t, s] = VIEW_TITLES[view];
  $("#viewTitle").textContent = t;
  $("#viewSub").textContent = s;
  if (view === "mapping") renderMapping();
  if (view === "notes") renderNotes();
}

function applyTheme(t) {
  document.documentElement.dataset.theme = t;
  $(".ic-moon").classList.toggle("hidden", t === "light");
  $(".ic-sun").classList.toggle("hidden", t !== "light");
  try { localStorage.setItem("att_theme", t); } catch (e) {}
}

/* ---------------- شاشة التوليد ---------------- */

function ym() {
  return { year: +$("#yearSel").value, month: +$("#monthSel").value };
}

function updateAttlogUI() {
  $("#attlogName").textContent = state.attlogName || "لم يتم الاختيار";
  $("#attlogMeta").textContent = state.attlogPath ? state.attlogPath : "ملف TXT من جهاز البصمة";
}

function updateTplUI() {
  $("#tplName").textContent = state.tplPath ? state.tplPath.split(/[\\/]/).pop() : "افتراضي: مولّد النظام";
  $("#clearTpl").classList.toggle("hidden", !state.tplPath);
}

async function pickAttlog() {
  if (isWebview()) {
    const r = await callApi("select_file", { kind: "attlog" });
    if (r && r.path) { state.attlogPath = r.path; state.attlogName = r.name; updateAttlogUI(); }
    return;
  }
  const inp = document.createElement("input");
  inp.type = "file"; inp.accept = ".txt,.log,text/plain";
  inp.onchange = async () => {
    const f = inp.files[0]; if (!f) return;
    if (isServerMode()) {
      const b64 = await new Promise(res => { const rd = new FileReader(); rd.onload = () => res(rd.result.split(",")[1]); rd.readAsDataURL(f); });
      try {
        const r = await callApi("ingest_file", { name: f.name, b64, kind: "attlog" });
        state.attlogPath = r.path; state.attlogName = r.name || f.name;
      } catch (e) { toast(String(e.message || e), "err"); return; }
    } else {
      state.attlogPath = "(ملف تجريبي)"; state.attlogName = f.name;
    }
    updateAttlogUI();
  };
  inp.click();
}

async function pickTpl() {
  if (isWebview()) {
    const r = await callApi("select_file", { kind: "template" });
    if (r && r.path) { state.tplPath = r.path; updateTplUI(); }
    return;
  }
  const inp = document.createElement("input");
  inp.type = "file"; inp.accept = ".xlsx";
  inp.onchange = async () => {
    const f = inp.files[0]; if (!f) return;
    if (isServerMode()) {
      const b64 = await new Promise(res => { const rd = new FileReader(); rd.onload = () => res(rd.result.split(",")[1]); rd.readAsDataURL(f); });
      try {
        const r = await callApi("ingest_file", { name: f.name, b64, kind: "template" });
        state.tplPath = r.path;
      } catch (e) { toast(String(e.message || e), "err"); return; }
    } else { state.tplPath = "(قالب تجريبي).xlsx"; }
    updateTplUI();
  };
  inp.click();
}

async function doPreview() {
  if (!state.attlogPath) { toast("اختر ملف attlog أولاً", "warn"); return; }
  const { year, month } = ym();
  const r = await busy(async () => {
    const res = await callApi("preview", { attlog: state.attlogPath, year, month });
    renderPreview(res, year, month);
  }, "جارٍ قراءة وتحليل البصمات…");
  return r;
}

function renderPreview(res, year, month) {
  state.preview = res;
  const st = res.stats;
  $("#kpiRecords").textContent = st.records;
  $("#kpiCorrupt").textContent = st.corrupt;
  $("#kpiDup").textContent = st.duplicates;
  $("#kpiUnmapped").textContent = Object.keys(st.unmapped || {}).length;
  $("#encChip").textContent = "الترميز: " + (st.encoding || "—") + " • " + MONTHS_AR[month - 1] + " " + year;

  const unmappedCount = Object.keys(st.unmapped || {}).length;
  const banner = $("#unmappedBanner");
  if (unmappedCount) {
    banner.classList.remove("hidden");
    $("#unmappedText").textContent = `يوجد ${unmappedCount} رقم وظيفي غير مربوط بأسماء القالب — سيظهر صفهم فارغاً حتى الربط: ` +
      Object.entries(st.unmapped).map(([id, n]) => `${id} (${n || "?"})`).join(" ، ");
    $("#unmappedBadge").textContent = unmappedCount;
    $("#unmappedBadge").classList.remove("hidden");
  } else {
    banner.classList.add("hidden");
    $("#unmappedBadge").classList.add("hidden");
  }

  const tb = $("#empTable tbody");
  tb.innerHTML = "";
  (st.employees || []).forEach(e => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td><b>${e.name}</b></td><td class="center">${e.days}</td>` +
      `<td class="center">${e.problems ? `<span class="badge warn">${e.problems}</span>` : '<span class="badge ok">سليم</span>'}</td>`;
    tb.appendChild(tr);
  });

  $("#previewArea").classList.remove("hidden");
  $("#genEmpty").classList.add("hidden");
  $("#btnGenerate").disabled = false;
  toast(`تمت قراءة ${st.records} بصمة لـ ${(st.employees || []).length} موظف`);
}

function addOutFiles(files, logs) {
  const box = $("#outLog");
  if (logs) {
    const p = document.createElement("p");
    p.textContent = logs.join("  •  ");
    box.prepend(p);
  }
  (files || []).forEach(f => {
    const line = document.createElement("div");
    line.className = "file-line";
    line.innerHTML = ICON_FILE + '<code title="انقر للفتح"></code>';
    line.querySelector("code").textContent = f;
    line.querySelector("code").onclick = () => openPath(f);
    box.appendChild(line);
  });
}

async function doGenerate(kind) {
  if (!state.attlogPath) { toast("اختر ملف attlog أولاً", "warn"); return; }
  const { year, month } = ym();
  await busy(async () => {
    let out = null;
    if (isWebview()) {
      const def = kind === "csv" ? `تحقق_${year}_${String(month).padStart(2, "0")}.csv` : `دوام_${MONTHS_AR[month - 1]}_${year}.xlsx`;
      const r = await callApi("save_dialog", { kind, default_name: def });
      if (r && r.path) out = r.path;
    }
    const res = await callApi("generate", { attlog: state.attlogPath, year, month, out, kind: kind || "xlsx", template: state.tplPath || null });
    addOutFiles(res.files, res.logs);
    toast(kind === "csv" ? "تم تصدير ملف التحقق CSV" : "تم توليد ملف الدوام بنجاح");
  }, "جارٍ توليد ملف الدوام…");
}

async function showQuality() {
  if (!state.attlogPath) { toast("اختر ملف attlog أولاً", "warn"); return; }
  const { year, month } = ym();
  await busy(async () => {
    const r = await callApi("quality", { attlog: state.attlogPath, year, month });
    state.qualityLines = r.lines || [];
    const pre = document.createElement("pre");
    pre.textContent = state.qualityLines.join("\n");
    const body = $("#modalBody"); body.innerHTML = ""; body.appendChild(pre);
    $("#modalWrap").classList.remove("hidden");
  }, "جارٍ إعداد تقرير الجودة…");
}

function openPath(p) { if (p) callApi("open_path", { path: p }); }

async function saveQualityTxt() {
  if (!state.attlogPath || !state.qualityLines.length) return;
  const { year, month } = ym();
  let out = null;
  if (isWebview()) {
    const r = await callApi("save_dialog", { kind: "quality", default_name: `تقرير_الجودة_${year}_${String(month).padStart(2, "0")}.txt` });
    if (r && r.path) out = r.path;
  }
  await busy(async () => {
    const res = await callApi("save_quality", { attlog: state.attlogPath, year, month, out });
    addOutFiles(res.files, null);
    toast("حُفظ تقرير الجودة كملف نصي");
  }, "جارٍ حفظ التقرير…");
}

/* ---------------- شاشة الربط ---------------- */

function empStatus(e) {
  if (e.no_punch) return ["na", "بلا بصمة"];
  if (e.template_name) return ["ok", "مربوط"];
  if (e.id) return ["warn", "غير مربوط"];
  return ["mute", "اسم فقط"];
}

function renderMapping() {
  if (!state.settings) return;
  const q = ($("#mapSearch").value || "").trim().toLowerCase();
  const tb = $("#mapTable tbody");
  tb.innerHTML = "";
  let shown = 0;
  state.settings.employees.forEach((e, i) => {
    const hay = ((e.template_name || "") + " " + (e.device_name || "") + " " + (e.id || "")).toLowerCase();
    if (q && !hay.includes(q)) return;
    shown++;
    const [cls, label] = empStatus(e);
    const tr = document.createElement("tr");
    tr.innerHTML =
      `<td class="center"><code>${e.id == null ? "—" : e.id}</code></td>` +
      `<td>${e.device_name || "—"}</td>` +
      `<td><input type="text" data-idx="${i}" data-field="template_name" value="" placeholder="اكتب الاسم كما في القالب…"></td>` +
      `<td class="center"><span class="badge ${cls}">${label}</span></td>` +
      `<td class="center"><button class="del-btn" title="حذف">✕</button></td>`;
    tr.querySelector("input").value = e.template_name || "";
    tr.querySelector("input").oninput = ev => { state.settings.employees[i].template_name = ev.target.value.trim(); };
    tr.querySelector(".del-btn").onclick = () => {
      state.settings.employees.splice(i, 1);
      renderMapping();
    };
    tb.appendChild(tr);
  });
  $("#mapCount").textContent = `${shown} موظف`;
}

async function saveEmployees(silent) {
  await callApi("save_settings", { settings: state.settings });
  if (!silent) toast("حُفظ جدول الربط — سيُستخدم في كل التوليدات القادمة");
}

async function scanIds() {
  if (!state.attlogPath) { toast("اختر ملف attlog في شاشة التوليد أولاً", "warn"); return; }
  await busy(async () => {
    const r = await callApi("scan_ids", { attlog: state.attlogPath });
    await loadState(false);
    renderMapping();
    toast(r.added ? `أُضيف ${r.added} رقم جديد من الملف` : "لا توجد أرقام جديدة — الكل موجود");
  }, "جارٍ فحص attlog…");
}

function addEmployee() {
  const id = prompt("الرقم الوظيفي في جهاز البصمة (اتركه فارغاً لموظف بلا بصمة):");
  if (id === null) return;
  const name = prompt("الاسم كما يظهر في القالب:");
  if (!name) return;
  state.settings.employees.push({ id: id ? parseInt(id, 10) : null, template_name: name, device_name: "", no_punch: !id });
  renderMapping();
}

/* ---------------- شاشة الملاحظات ---------------- */

function empLabel(e) {
  return (e.id != null ? "#" + e.id : "بلا رقم") + " — " + (e.template_name || e.device_name || "بدون اسم");
}

function renderNotes() {
  if (!state.settings) return;
  const sel = $("#noteEmp");
  sel.innerHTML = "";
  state.settings.employees.forEach(e => {
    const o = document.createElement("option");
    o.value = e.id == null ? "" : e.id;
    o.textContent = empLabel(e);
    sel.appendChild(o);
  });
  const notes = state.settings.notes || {};
  const tb = $("#notesTable tbody");
  tb.innerHTML = "";
  const keys = Object.keys(notes).sort().reverse();
  keys.forEach(key => {
    const [d, id] = key.split("|");
    const emp = state.settings.employees.find(x => String(x.id) === id);
    const tr = document.createElement("tr");
    tr.innerHTML = `<td class="center"><code>${d}</code></td><td>${emp ? (emp.template_name || emp.device_name || "#" + id) : "#" + id}</td>` +
      `<td><span class="badge mute">${notes[key]}</span></td><td class="center"><button class="del-btn">✕</button></td>`;
    tr.querySelector(".del-btn").onclick = () => {
      delete state.settings.notes[key];
      saveEmployees(true);
      renderNotes();
      toast("حُذفت الملاحظة");
    };
    tb.appendChild(tr);
  });
  $("#notesCount").textContent = keys.length + " ملاحظة";
}

async function addNote() {
  const d = $("#noteDate").value;
  const empId = $("#noteEmp").value;
  const extra = $("#noteExtra").value.trim();
  if (!d || !empId) { toast("اختر التاريخ والموظف", "warn"); return; }
  const note = state.noteKind + (extra ? " — " + extra : "");
  state.settings.notes = state.settings.notes || {};
  state.settings.notes[`${d}|${empId}`] = note;
  $("#noteExtra").value = "";
  saveEmployees(true);
  renderNotes();
  toast(`أُضيفت الملاحظة: ${note}`);
}

/* ---------------- شاشة الإعدادات ---------------- */

function renderSettings() {
  const s = state.settings;
  $("#setTplPath").value = s.template_path || "";
  $("#setDedup").value = s.dedup_seconds ?? 120;
  $("#setMidday").value = s.midday_split || "12:00";
  $("#setLate").value = s.late_after || "08:15";
  $("#setLateOn").checked = !!s.late_enabled;
  const wds = new Set((s.workdays || [6, 0, 1, 2, 3]).map(Number));
  $$("#workdayChips .day-chip").forEach(ch => ch.classList.toggle("on", wds.has(+ch.dataset.wd)));
}

async function saveSettings() {
  const s = state.settings;
  s.template_path = $("#setTplPath").value.trim();
  s.dedup_seconds = Math.max(0, parseInt($("#setDedup").value, 10) || 120);
  s.midday_split = $("#setMidday").value || "12:00";
  s.late_after = $("#setLate").value || "08:15";
  s.late_enabled = $("#setLateOn").checked;
  s.workdays = $$("#workdayChips .day-chip.on").map(ch => +ch.dataset.wd);
  if (state.tplPath) s.template_path = state.tplPath;
  await callApi("save_settings", { settings: s });
  toast("حُفظت الإعدادات بنجاح");
}

/* ---------------- التحميل والتشغيل ---------------- */

async function loadState(notify) {
  const r = await callApi("get_state", {});
  state.settings = r.settings;
  renderSettings();
  renderMapping();
  renderNotes();
  if (notify) toast("تم تحميل الإعدادات");
}

function buildSelects() {
  const ms = $("#monthSel");
  MONTHS_AR.forEach((m, i) => {
    const o = document.createElement("option");
    o.value = i + 1; o.textContent = m;
    if (i === 8) o.selected = true;
    ms.appendChild(o);
  });
  const ys = $("#yearSel");
  for (let y = 2024; y <= 2032; y++) {
    const o = document.createElement("option");
    o.value = y; o.textContent = y;
    if (y === 2026) o.selected = true;
    ys.appendChild(o);
  }
}

function bindUI() {
  $$(".nav-item").forEach(b => b.onclick = () => goto(b.dataset.view));
  $$("[data-goto]").forEach(b => b.onclick = () => goto(b.dataset.goto));
  $("#themeToggle").onclick = () =>
    applyTheme(document.documentElement.dataset.theme === "dark" ? "light" : "dark");

  $("#pickAttlog").onclick = pickAttlog;
  $("#pickTpl").onclick = pickTpl;
  $("#clearTpl").onclick = () => { state.tplPath = ""; updateTplUI(); toast("أُزيل القالب — سيُستخدم مولّد النظام", "warn"); };

  $("#btnPreview").onclick = () => doPreview();
  $("#btnGenerate").onclick = () => doGenerate("xlsx");
  $("#btnGenerate2").onclick = () => doGenerate("xlsx");
  $("#btnCsv").onclick = () => doGenerate("csv");
  $("#btnQuality").onclick = showQuality;
  $("#btnOpenFolder").onclick = () => callApi("open_path", { path: "OUTPUT_DIR" });

  $("#mapSearch").oninput = () => renderMapping();
  $("#btnScanIds").onclick = scanIds;
  $("#btnAddEmp").onclick = addEmployee;
  $("#btnSaveMap").onclick = () => busy(() => saveEmployees(false), "جارٍ الحفظ…");

  $$("#noteKinds .chip-btn").forEach(b => b.onclick = () => {
    $$("#noteKinds .chip-btn").forEach(x => x.classList.remove("active"));
    b.classList.add("active");
    state.noteKind = b.dataset.kind;
  });
  $("#btnAddNote").onclick = () => busy(addNote, "جارٍ الحفظ…");

  $$("#workdayChips .day-chip").forEach(ch => ch.onclick = () => ch.classList.toggle("on"));
  $("#btnSaveSettings").onclick = () => busy(saveSettings, "جارٍ الحفظ…");
  $("#setPickTpl").onclick = pickTpl;

  $("#modalClose").onclick = () => $("#modalWrap").classList.add("hidden");
  $("#modalWrap").onclick = e => { if (e.target === $("#modalWrap")) $("#modalWrap").classList.add("hidden"); };
  $("#modalCopy").onclick = () => {
    navigator.clipboard.writeText(state.qualityLines.join("\n")).then(() => toast("نُسخ التقرير"));
  };
  $("#modalSave").onclick = saveQualityTxt;
}

/* ---------------- الإقلاع ---------------- */

(async function boot() {
  applyTheme((() => { try { return localStorage.getItem("att_theme") || "dark"; } catch (e) { return "dark"; } })());
  buildSelects();
  bindUI();

  // انتظار جسر pywebview عند التشغيل داخل نافذة أصلية
  if (!location.protocol.startsWith("http") && !window.__NO_WEBVIEW__) await waitWebview(2500);

  try {
    await loadState(false);
    const info = await callApi("app_info", {});
    $("#aboutSettings").textContent = info.settings_path || "—";
    $("#aboutMode").textContent = info.mode || "—";
    state.ready = true;
    if (FORCE_DEMO || isDemo()) {
      $("#modeChip").classList.remove("hidden");
      $("#enginePill span").textContent = "وضع العرض";
    } else if (isWebview()) {
      $("#enginePill span").textContent = "متصل بالمحرك";
    } else {
      $("#enginePill span").textContent = "وضع المتصفح";
    }
  } catch (e) {
    toast("تعذر الاتصال بالمحرك: " + String(e.message || e), "err");
  }
})();
