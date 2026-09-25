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
const isServerMode = () => !isWebview() && !FORCE_DEMO && location.protocol.startsWith("http");
const isDemo = () => !isWebview() && (!isServerMode() || FORCE_DEMO);
const isLocalHost = () => ["localhost", "127.0.0.1", "", "0.0.0.0"].includes(location.hostname);

async function httpApi(fn, payload) {
  const res = await fetch("/api/" + fn, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload || {}),
  });
  const ct = res.headers.get("content-type") || "";
  const text = await res.text();
  let data = null;
  try { data = JSON.parse(text); } catch (e) {
    throw new Error("الاستجابة ليست من محرك النظام — تأكد من تشغيل التطبيق من desktop_app.py");
  }
  if (!res.ok || data.ok === false) throw new Error(data.error || ("HTTP " + res.status));
  return data;
}

let FORCE_DEMO = false;

/* قراءة ملف من القرص كـ base64 مع معالجة أخطاء واضحة */
function fileToB64(f) {
  return new Promise((res, rej) => {
    const rd = new FileReader();
    rd.onerror = () => rej(new Error("تعذر قراءة الملف من القرص — تأكد أنه غير محمي وأن لديك صلاحية عليه"));
    rd.onload = () => {
      const s = String(rd.result || "");
      const b64 = s.includes(",") ? s.split(",")[1] : s;
      if (!b64) rej(new Error("الملف فارغ أو تعذر قراءته")); else res(b64);
    };
    rd.readAsDataURL(f);
  });
}

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
        // تأخر جسر pywebview في النافذة الأصلية (تشتغل بخادم داخلي بلا /api)؟
        // ننتظر ظهوره قبل أي تحول للوضع التجريبي — هذا سبب «فشل الرفع» في EXE سابقاً
        const maybeWebview = !!window.pywebview && !window.__NO_WEBVIEW__;
        const localMaybe = isLocalHost() && !window.__NO_WEBVIEW__;
        if (!isWebview() && (maybeWebview || localMaybe)) {
          if (await waitWebview(maybeWebview ? 8000 : 3000)) return callApi(fn, payload);
        }
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

/* انتظار ذكي: داخل النافذة الأصلية قد يتأخر الجسر — ننتظره قبل أي قرار */
async function waitForBridge(ms) {
  if (isWebview()) return true;
  if (window.__NO_WEBVIEW__) return false;
  // إشارات بيئة نافذة سطح مكتب: وجود كائن pywebview أو بروتوكول غير http
  if (window.pywebview || !location.protocol.startsWith("http")) return await waitWebview(ms || 8000);
  return isWebview();
}

/* ---------------- وضع المعالجة المحلية (بدون خادم — محرك حقيقي داخل الصفحة) ----------------
   عبر web/engine.js: يقرأ attlog فعلياً، يحلل القالب، ويولّد ملف Excel حقيقي
   قابل للتنزيل — كل ذلك محلياً على جهاز المستخدم دون إرسال أي ملف لأي خادم. */

const DEMO_STORE = {
  settings: {
    version: 1, template_path: "", dedup_seconds: 120, midday_split: "12:00",
    late_after: "08:15", late_enabled: false,
    workdays: [6, 0, 1, 2, 3],
    employees: [
      { id: 1025, template_name: "يزن أبو بايع",   device_name: "Yazeed",   no_punch: false },
      { id: 1000, template_name: "دانيامحموض",     device_name: "Dania ma", no_punch: false },
      { id: 1004, template_name: "حازم حط",        device_name: "Hazem Go", no_punch: false },
      { id: 1006, template_name: "براءه بطرني",    device_name: "Baraa B",  no_punch: false },
      { id: 1017, template_name: "محمد طارق كركش", device_name: "M.Tareq",  no_punch: false },
      { id: 1003, template_name: "",               device_name: "A.Diab",   no_punch: false },
      { id: 1022, template_name: "",               device_name: "Sallam",   no_punch: false },
    ],
    notes: { "2026-09-20|1017": "اجازة", "2026-09-06|1016": "مأمورية" },
  },
  attlog: null,    // {name, bytes: Uint8Array}
  tpl: null,       // {name, bytes}
  outBlobs: {},    // اسم الملف → Blob (ناتج توليد حقيقي قابل للتنزيل)
};

const sleep = ms => new Promise(r => setTimeout(r, ms));

function b64ToBytes(b64) {
  const bin = atob(String(b64 || "").replace(/\s+/g, ""));
  const u8 = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
  return u8;
}

function bytesToBlob(u8, mime) {
  return new Blob([u8.slice().buffer], { type: mime || "application/octet-stream" });
}

async function demoApi(fn, p) {
  const S = DEMO_STORE.settings;

  switch (fn) {
    case "app_info":
      return { ok: true, version: "2.4", mode: "المعالجة المحلية داخل المتصفح (محرك مدمج)",
               settings_path: "(الوضع المحلي — لا يوجد ملف إعدادات)" };

    case "get_state":
      return { ok: true, settings: S };

    case "save_settings":
      if (p && p.settings) Object.assign(S, p.settings);
      return { ok: true };

    case "select_file":
      return { ok: false, error: "الحوار الأصلي غير متاح هنا — استخدم زر اختيار الملف أو اسحب الملف إلى البطاقة" };

    case "ingest_file": {
      await sleep(200);
      const bytes = b64ToBytes(p.b64);
      if (!bytes.length) return { ok: false, error: "الملف المختار فارغ (0 بايت)" };
      if (bytes.length > 40 * 1024 * 1024) return { ok: false, error: "الملف أكبر من الحد المسموح (40MB)" };
      const kind = p.kind || "";
      if (kind === "template") {
        if (bytes.length < 100 || bytes[0] !== 0x50 || bytes[1] !== 0x4b)
          return { ok: false, error: "هذا الملف ليس Excel بصيغة xlsx الحديثة (قد يكون xls قديماً أو تالفاً). الحل: افتحه في Excel ثم ملف ← حفظ باسم ← «Excel Workbook (*.xlsx)» وأعد اختياره." };
      }
      if (kind === "attlog" && bytes[0] === 0x50 && bytes[1] === 0x4b)
        return { ok: false, error: "يبدو أنك اخترت ملف Excel لملف البصمات — ملف البصمات attlog ملف نصي (TXT) يصدّره جهاز البصمة." };
      const rawName = String(p.name || "file").split(/[\\/]/).pop();
      if (kind === "attlog") { DEMO_STORE.attlog = { name: rawName, bytes }; DEMO_STORE.outBlobs = {}; }
      if (kind === "template") { DEMO_STORE.tpl = { name: rawName, bytes }; S.template_path = "(محلي) " + rawName; }
      return { ok: true, path: "(محلي) " + rawName, name: rawName, size: bytes.length };
    }

    case "template_info": {
      await sleep(150);
      if (!DEMO_STORE.tpl) return { ok: false, error: "اختر ملف القالب أولاً" };
      try {
        const year = p.year || 2026, month = p.month || 9;
        const info = await window.Eng.analyzeTemplate(DEMO_STORE.tpl.bytes, year, month);
        return Object.assign({ ok: true }, info);
      } catch (e) {
        return { ok: false, error: "تعذر تحليل القالب: " + String(e.message || e) };
      }
    }

    case "preview": {
      await sleep(400);
      if (!DEMO_STORE.attlog) {
        // بيانات توضيحية عندما لا يُختار ملف بعد
        const perEmp = { 1025: { days: 22, problems: 2 }, 1000: { days: 21, problems: 1 },
                         1004: { days: 22, problems: 0 }, 1006: { days: 20, problems: 3 },
                         1017: { days: 21, problems: 1 }, 1003: { days: 18, problems: 2 },
                         1022: { days: 9,  problems: 1 } };
        const employees = Object.entries(perEmp).map(([id, s]) => {
          const e = S.employees.find(x => x.id === +id) || {};
          return { id: +id, name: e.template_name || e.device_name || ("#" + id), days: s.days, problems: s.problems };
        });
        return { ok: true, stats: { records: 287, corrupt: 2, duplicates: 4, encoding: "utf-8",
                 unmapped: { 1023: "Khaled", 1031: "Omar" }, employees, demo: true } };
      }
      const year = p.year || 2026, month = p.month || 9;
      const pr = window.Eng.parseAttlog(DEMO_STORE.attlog.bytes, +S.dedup_seconds || 120);
      const att = window.Eng.processMonth(pr, S, year, month);
      const sum = att.summary();
      const employees = Object.entries(sum.perEmp).map(([id, s]) =>
        ({ id: +id, name: s.name || ("#" + id), days: s.days, problems: s.problems }))
        .sort((a, b) => a.id - b.id);
      return { ok: true, stats: { records: pr.records.length, corrupt: pr.corrupt.length,
               duplicates: pr.duplicates, encoding: pr.encoding,
               total_lines: pr.totalLines, unmapped: sum.unmapped, employees } };
    }

    case "generate": {
      await sleep(500);
      const year = p.year || 2026, month = p.month || 9;
      const isCsv = p.kind === "csv";
      if (!DEMO_STORE.attlog) return { ok: false, error: "اختر ملف attlog أولاً — في هذا الوضع يُعالج الملف فعلياً محلياً" };
      const pr = window.Eng.parseAttlog(DEMO_STORE.attlog.bytes, +S.dedup_seconds || 120);
      // حارس: لا تولّد ملفاً فارغاً أبداً — أوقف برسالة تشخيصية واضحة
      if (!pr.records.length) {
        const head = new TextDecoder("utf-8", { fatal: false })
          .decode(DEMO_STORE.attlog.bytes.slice(0, 300)).replace(/\r/g, " ").replace(/\n/g, " ⏎ ");
        const reasons = pr.corrupt.slice(0, 3).map(c => `   • سطر ${c.no}: ${c.reason}`).join("\n");
        return { ok: false, error:
          "فشل التفريغ: لم أفهم أي سطر من الملف النصي (" + pr.totalLines + " سطر) — لن أوّلّد ملفاً فارغاً.\n" +
          "الترميز المكتشف: " + pr.encoding + "\n" +
          (reasons ? "أسباب الاستبعاد:\n" + reasons + "\n" : "") +
          "أول ما في ملفك: «" + head + "»\n" +
          "تأكد أن الملف هو attlog.txt الصادر من جهاز البصمة. إن بقي الخطأ أرسل أول 3 أسطر من الملف." };
      }
      const att = window.Eng.processMonth(pr, S, year, month);
      const sum = att.summary();
      const base = isCsv ? `تحقق_${year}_${String(month).padStart(2, "0")}.csv`
                         : `دوام_${window.Eng.MONTH_AR[month]}_${year}.xlsx`;
      const logs = [
        `صيغة الملف: ${pr.hasNames ? "بأسماء الموظفين" : "بدون أسماء (صيغة الجهاز: رقم + تاريخ + حالات)"}`,
        `[1/4] قراءة attlog: ${pr.records.length} بصمة سليمة | ${pr.corrupt.length} سطر تالف | ${pr.duplicates} مكرر محذوف (ترميز: ${pr.encoding})`,
        `[2/4] المعالجة: ${sum.rowsCount} يوم/موظف | ${Object.keys(sum.unmapped).length} رقم غير مربوط`,
      ];
      try {
        if (isCsv) {
          const csv = "\uFEFF" + window.Eng.buildCsv(att, year, month);
          DEMO_STORE.outBlobs[base] = bytesToBlob(new TextEncoder().encode(csv), "text/csv;charset=utf-8");
          logs.push("[3/4] تصدير CSV للتحقق: " + (att.rows.size) + " سطر");
        } else {
          let outBytes = null;
          const tpl = DEMO_STORE.tpl;
          if (tpl) {
            try {
              const res = await window.Eng.fillTemplateBytes(tpl.bytes, att, S.employees || [], year, month);
              outBytes = res.bytes;
              logs.push("[3/4] Excel (وضع القالب الأصلي — نسخة من قالبكم نفسه)");
              res.messages.forEach(m => logs.push("      " + m));
            } catch (e) {
              logs.push("⚠ القالب: " + String(e.message || e) + " — سيُستخدم المولّد المدمج");
            }
          }
          if (!outBytes) {
            outBytes = window.Eng.buildFallbackBytes(att, S, year, month);
            logs.push("[3/4] Excel (المولّد المدمج المطابق لبنية القالب)" + (tpl ? "" : " — لم يُحدد قالب"));
            logs.push("      ملاحظة: للتعبيئة داخل قالبكم الأصلي بنفس التنسيق استخدم تطبيق سطح المكتب (AttendanceDump.exe)");
          }
          DEMO_STORE.outBlobs[base] = bytesToBlob(outBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        }
      } catch (e) {
        return { ok: false, error: "فشل التوليد: " + String(e.message || e) };
      }
      logs.push("[4/4] جاهز — الملف مُولّد محلياً على جهازك");
      return { ok: true, path: "(محلي) " + base, files: [base], logs };
    }

    case "quality": {
      await sleep(300);
      if (!DEMO_STORE.attlog) return { ok: false, error: "اختر ملف attlog أولاً" };
      const year = p.year || 2026, month = p.month || 9;
      const pr = window.Eng.parseAttlog(DEMO_STORE.attlog.bytes, +S.dedup_seconds || 120);
      const att = window.Eng.processMonth(pr, S, year, month);
      return { ok: true, lines: window.Eng.buildQuality(att, pr, S, year, month) };
    }

    case "open_path":
      toast("في الوضع المحلي: " + (p.path || ""), "warn");
      return { ok: true };

    case "scan_ids": {
      await sleep(300);
      if (!DEMO_STORE.attlog) return { ok: false, error: "اختر ملف attlog في شاشة التوليد أولاً" };
      const pr = window.Eng.parseAttlog(DEMO_STORE.attlog.bytes, +S.dedup_seconds || 120);
      const known = new Set((S.employees || []).map(e => e.id));
      let added = 0;
      for (const id of Object.keys(pr.idsFound)) {
        if (!known.has(+id)) {
          S.employees.push({ id: +id, template_name: "", device_name: pr.idsFound[id], no_punch: false });
          added++;
        }
      }
      return { ok: true, added };
    }

    default:
      return { ok: true };
  }
}

/* ---------------- أدوات واجهة ---------------- */

const ICON_OK = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6 9 17l-5-5"/></svg>';
const ICON_ERR = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><circle cx="12" cy="12" r="10"/><path d="M15 9l-6 6M9 9l6 6"/></svg>';
const ICON_WARN = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0z"/><path d="M12 9v4M12 17h.01"/></svg>';
const ICON_FILE = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/></svg>';
const ICON_DL  = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="M7 10l5 5 5-5"/><path d="M12 15V3"/></svg>';

function toast(msg, kind, ms) {
  const el = document.createElement("div");
  el.className = "toast " + (kind === "err" ? "err" : kind === "warn" ? "warn" : "");
  el.innerHTML = (kind === "err" ? ICON_ERR : kind === "warn" ? ICON_WARN : ICON_OK) + "<span></span>";
  el.querySelector("span").textContent = msg;
  $("#toasts").appendChild(el);
  const dur = ms && ms > 1000 ? ms : 3600;
  setTimeout(() => { el.classList.add("out"); setTimeout(() => el.remove(), 300); }, dur);
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

async function showTplInfo(path) {
  const bar = $("#tplBar");
  if (!path) { bar.classList.add("hidden"); return; }
  const { year, month } = ym();
  try {
    const info = await callApi("template_info", { path, year, month });
    bar.classList.remove("hidden");
    bar.classList.toggle("warn", !!(info.warning || !info.ok));
    if (!info.ok) { $("#tplBarText").textContent = "⚠ " + (info.error || "تعذر تحليل القالب"); return; }
    if (info.warning) {
      $("#tplBarText").textContent = "⚠ " + info.warning;
    } else {
      const dcount = (info.month_days_found || []).length;
      $("#tplBarText").textContent =
        `${info.mode} — ${info.names_count} اسماً في القالب، ${dcount} يوماً لـ ${MONTHS_AR[month - 1]} ${year} (ورقة: ${info.sheet})`;
    }
  } catch (e) {
    bar.classList.remove("hidden");
    bar.classList.add("warn");
    $("#tplBarText").textContent = "⚠ تعذر تحليل القالب: " + String(e.message || e);
  }
}

/* ---------- اختيار الملفات: قنوات متعددة لضمان النجاح دائماً ----------
   1) حوار النظام الأصلي (وضع النافذة)  2) منتقي الملفات المدمج (كل الصيغ)
   3) السحب والإفلات على البطاقة — وبعد الاختيار تحقق واضح برسائل عربية */

function pickViaInput(onFile) {
  // بلا سمة accept إطلاقاً: كل الملفات قابلة للاختيار ولا شيء يظهر رمادياً
  const inp = document.createElement("input");
  inp.type = "file";
  inp.onchange = async () => {
    const f = inp.files && inp.files[0];
    if (f) await onFile(f);
  };
  inp.click();
}

function enableDrop(el, onFile) {
  if (!el) return;
  ["dragenter", "dragover"].forEach(ev =>
    el.addEventListener(ev, e => { e.preventDefault(); el.classList.add("drop-hover"); }));
  ["dragleave", "drop"].forEach(ev =>
    el.addEventListener(ev, e => { e.preventDefault(); el.classList.remove("drop-hover"); }));
  el.addEventListener("drop", e => {
    const f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
    if (f) onFile(f);
  });
}

function isCancelledErr(e) {
  return /cancel|إلغاء|ألغي/i.test(String(e && e.message || e || ""));
}

async function handleAttlogFile(f) {
  try {
    if (isWebview() || isServerMode()) {
      const b64 = await fileToB64(f);
      const r = await callApi("ingest_file", { name: f.name, b64, kind: "attlog" });
      state.attlogPath = r.path || "(ملف مرفوع) " + f.name;
      state.attlogName = r.name || f.name;
      toast("✅ تم استلام ملف البصمات بنجاح — " + (r.name || f.name), "ok");
    } else {
      // وضع المعالجة المحلية: يُرسل الملف كاملاً للمحرك المدمج
      const b64 = await fileToB64(f);
      const r = await callApi("ingest_file", { name: f.name, b64, kind: "attlog" });
      state.attlogPath = r.path || "(محلي) " + f.name;
      state.attlogName = r.name || f.name;
      toast("✅ استُلم الملف وسيُقرأ محلياً على جهازك — " + f.name, "ok");
    }
  } catch (e) { toast(String(e.message || e), "err"); return; }
  updateAttlogUI();
}

async function handleTplFile(f) {
  try {
    if (isWebview() || isServerMode()) {
      const b64 = await fileToB64(f);
      const r = await callApi("ingest_file", { name: f.name, b64, kind: "template" });
      state.tplPath = r.path || "(قالب مرفوع) " + f.name;
    } else {
      const b64 = await fileToB64(f);
      const r = await callApi("ingest_file", { name: f.name, b64, kind: "template" });
      state.tplPath = r.path || "(محلي) " + f.name;
      toast("✅ تم استلام القالب — سيُحلل محلياً — " + (r.name || f.name), "ok");
    }
    state.settings.template_path = state.tplPath;
    callApi("save_settings", { settings: state.settings }).catch(() => {});
    if (!isWebview() && !isServerMode()) { /* التوست أعلاه */ }
    else toast("✅ تم استلام القالب بنجاح — " + f.name, "ok");
  } catch (e) { toast(String(e.message || e), "err"); return; }
  updateTplUI();
  showTplInfo(state.tplPath);
}

async function pickAttlog() {
  try {
    if (isWebview()) {
      try {
        const r = await callApi("select_file", { kind: "attlog" });
        if (r && r.path) { state.attlogPath = r.path; state.attlogName = r.name; updateAttlogUI(); toast("تم اختيار ملف البصمات ✓", "ok"); return; }
        return; // أُلغي الحوار
      } catch (e) {
        if (isCancelledErr(e)) return;
        // فشل حوار النظام — نجرّب منتقي الملفات المدمج داخل النافذة
        toast("تعذر فتح حوار النظام — سيُستخدم منتقي الملفات المدمج", "warn");
      }
    }
    pickViaInput(handleAttlogFile);
  } catch (e) { if (!isCancelledErr(e)) toast(String(e.message || e), "err"); }
}

async function pickTpl() {
  try {
    if (isWebview()) {
      try {
        const r = await callApi("select_file", { kind: "template" });
        if (r && r.path) {
          state.tplPath = r.path;
          state.settings.template_path = r.path;
          callApi("save_settings", { settings: state.settings }).catch(() => {});
          updateTplUI();
          toast("✅ تم استلام القالب بنجاح — " + (r.name || ""), "ok");
          showTplInfo(r.path);
          return;
        }
        return; // أُلغي الحوار
      } catch (e) {
        if (isCancelledErr(e)) return;
        toast("تعذر فتح حوار النظام — سيُستخدم منتقي الملفات المدمج", "warn");
      }
    }
    pickViaInput(handleTplFile);
  } catch (e) { if (!isCancelledErr(e)) toast(String(e.message || e), "err"); }
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
  if (!st.records) {
    toast("⚠ لم تُفهم أي بصمة من الملف! التوليد سيرفض إنتاج ملف فارغ — راجع التشخيص وأرسل أول 3 أسطر من ملفك إن استمر الخطأ", "err", 9000);
  } else {
    toast(`تمت قراءة ${st.records} بصمة لـ ${(st.employees || []).length} موظف` +
      (st.corrupt ? ` | ${st.corrupt} سطر تالف` : ""));
  }
}

function baseName(p) { return String(p).split(/[\\/]/).pop() || "file"; }

async function downloadFile(path) {
  if (!path) return;
  const name = baseName(path);

  // 0) ناتج توليد حقيقي من المحرك المحلي — يُنزّل كما هو
  if (typeof DEMO_STORE !== "undefined" && DEMO_STORE.outBlobs[name]) {
    const url = URL.createObjectURL(DEMO_STORE.outBlobs[name]);
    const a = document.createElement("a");
    a.href = url;
    a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
    toast("نُزّل: " + name);
    return;
  }

  // 1) وضع المعالجة المحلية بلا ناتج مولّد: نولّد CSV حقيقي من بيانات العرض
  if (FORCE_DEMO || isDemo()) {
    const rows = [["الرقم", "الموظف", "أيام الحضور", "أيام بملاحظات"]];
    (state.settings.employees || []).forEach(e => {
      if (e.id == null) return;
      rows.push([e.id, e.template_name || e.device_name || ("#" + e.id), "", ""]);
    });
    const csv = "\uFEFF" + rows.map(r => r.join(",")).join("\r\n");
    const url = URL.createObjectURL(new Blob([csv], { type: "text/csv;charset=utf-8" }));
    const a = document.createElement("a");
    a.href = url;
    a.download = name.replace(/\.(xlsx|csv|txt)$/i, "") + "_تجريبي.csv";
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
    toast("ولّد ملف attlog ثم اضغط «توليد» للحصول على Excel حقيقي", "warn");
    return;
  }

  // 2) وضع المتصفح: نقطة تنزيل من الخادم المحلي
  if (isServerMode()) {
    window.location.href = "/api/download?name=" + encodeURIComponent(name);
    toast("بدأ تنزيل: " + name);
    return;
  }

  // 3) وضع النافذة الأصلية: حوار «حفظ باسم» ثم نسخ الملف
  try {
    const r = await callApi("download_file", { path, default_name: name });
    if (r && r.ok) toast("حُفظ الملف في: " + r.path);
  } catch (e) {
    if (!/cancel/i.test(String(e.message || e))) toast(String(e.message || e), "err");
  }
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
    line.innerHTML = ICON_FILE + '<code title="انقر لفتح الملف"></code>' +
      '<button class="dl-btn" title="تنزيل الملف إلى جهازك">' + ICON_DL + '<span>تنزيل</span></button>';
    const code = line.querySelector("code");
    code.textContent = f;
    code.onclick = () => openPath(f);
    line.querySelector(".dl-btn").onclick = () => downloadFile(f);
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
    // وضع المتصفح/المحلي: بدء التنزيل تلقائياً بعد التوليد
    if (!isWebview()) {
      (res.files || []).forEach((f, i) => setTimeout(() => downloadFile(f), 600 + i * 700));
    }
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

  // اختيار الملفات: زر + النقر على البطاقة كاملة + السحب والإفلات
  $("#pickAttlog").onclick = e => { e.stopPropagation(); pickAttlog(); };
  $("#attlogCard").onclick = () => pickAttlog();
  enableDrop($("#attlogCard"), handleAttlogFile);
  $("#pickTpl").onclick = e => { e.stopPropagation(); pickTpl(); };
  $("#tplCard").onclick = () => pickTpl();
  enableDrop($("#tplCard"), handleTplFile);
  $("#clearTpl").onclick = e => {
    e.stopPropagation();
    state.tplPath = "";
    state.settings.template_path = "";
    callApi("save_settings", { settings: state.settings });
    updateTplUI();
    $("#tplBar").classList.add("hidden");
    toast("أُزيل القالب — سيُستخدم المولّد المطابق لبنية القالب", "warn");
  };

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

  // انتظار جسر pywebview عند التشغيل داخل نافذة أصلية (حتى لو كان العرض عبر خادم داخلي http)
  await waitForBridge(8000);

  try {
    await loadState(false);
    const info = await callApi("app_info", {});
    $("#aboutSettings").textContent = info.settings_path || "—";
    $("#aboutMode").textContent = info.mode || "—";
    state.ready = true;
    // لو كان قالب محفوظاً من جلسة سابقة — أعرض حالته فوراً
    if (state.settings && state.settings.template_path &&
        !String(state.settings.template_path).startsWith("(محلي)")) {
      state.tplPath = state.settings.template_path;
      updateTplUI();
      showTplInfo(state.tplPath);
    }
    if (FORCE_DEMO || isDemo()) {
      $("#modeChip").classList.remove("hidden");
      $("#enginePill span").textContent = "المحرك المحلي جاهز";
    } else if (isWebview()) {
      $("#enginePill span").textContent = "متصل بالمحرك";
    } else {
      $("#enginePill span").textContent = "وضع المتصفح";
    }
  } catch (e) {
    toast("تعذر الاتصال بالمحرك: " + String(e.message || e), "err");
  }
})();
