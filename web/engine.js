"use strict";
/* ============================================================
   engine.js — محرك حقيقي يعمل داخل المتصفح (وضع العرض/الاستضافة الثابتة)
   ============================================================
   منفذ لخوارزميات النظام نفسها (att_parser/processor/template_writer/
   excel_writer/quality_report) بلغة JavaScript كي يعمل رحلة المستخدم
   كاملة — من اختيار ملف attlog حتى تنزيل Excel — حتى بدون خادم:
   • قراءة attlog بأي ترميز (UTF-8 / UTF-16 / CP1256) وإزالة التكرار
   • اقتران أول بصمة دخول وآخر بصمة خروج + ملاحظات النظام والمحفوظة
   • تحليل قالب Excel: تواريخ حقيقية أو نصية أو أرقام أيام (وحتى هندية)
   • توليد xlsx: تعبئة نسخة من قالبكم نفسه (تعديل XML فقط مع الحفاظ
     الكامل على التنسيق) أو المولّد المدمج المطابق لبنية القالب
   • كل شيء يجري محلياً على جهازك — لا يُرسل أي ملف إلى أي خادم
   ============================================================ */
window.Eng = (function () {

  /* ==================== أدوات عامة ==================== */

  const AR_DIGITS = "٠١٢٣٤٥٦٧٨٩";
  const normDigits = s => String(s).replace(/[٠-٩]/g, d => String(AR_DIGITS.indexOf(d)));
  const AR_DIAC_RE = /[\u064B-\u0652\u0640]/g;

  function normalizeAr(s) {
    if (s == null) return "";
    s = String(s).replace(AR_DIAC_RE, "")
      .replace(/[أإآ]/g, "ا").replace(/ى/g, "ي")
      .replace(/ة/g, "ه").replace(/ؤ/g, "و").replace(/ئ/g, "ي");
    return s.replace(/\s+/g, " ").trim().toLowerCase();
  }

  function toInt(v) {
    if (v == null || typeof v === "boolean" || v instanceof Date) return null;
    if (typeof v === "number") return Number.isInteger(v) ? v : null;
    const s = normDigits(String(v)).trim();
    if (/^\d{1,2}$/.test(s)) return +s;
    return null;
  }

  function fmtTime(mins) {
    const h = Math.floor(mins / 60), m = mins % 60;
    return h + ":" + String(m).padStart(2, "0");
  }

  const MONTH_AR = { 1: "يناير", 2: "فبراير", 3: "مارس", 4: "أبريل", 5: "مايو", 6: "يونيو",
                     7: "يوليو", 8: "أغسطس", 9: "سبتمبر", 10: "أكتوبر", 11: "نوفمبر", 12: "ديسمبر" };
  // الأيام بحساب weekday البايثون (الاثنين=0 ... الأحد=6)
  const DAY_AR = { 0: "الاثنين", 1: "الثلاثاء", 2: "الأربعاء", 3: "الخميس",
                   4: "الجمعة", 5: "السبت", 6: "الأحد" };

  function escapeXml(s) {
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;").replace(/'/g, "&apos;");
  }

  function pyWeekday(y, m, d) { // الاثنين=0 ... الأحد=6 (مطابق لبايثون)
    return (new Date(Date.UTC(y, m - 1, d)).getUTCDay() + 6) % 7;
  }

  const dstr = (y, m, d) => y + "-" + String(m).padStart(2, "0") + "-" + String(d).padStart(2, "0");

  function monthDaysCount(y, m) { return new Date(Date.UTC(y, m, 0)).getUTCDate(); }

  function monthWorkdays(year, month, workdays) {
    const wd = new Set((workdays || [6, 0, 1, 2, 3]).map(Number));
    const out = [];
    const n = monthDaysCount(year, month);
    for (let d = 1; d <= n; d++) {
      const pw = pyWeekday(year, month, d);
      if (wd.has(pw)) out.push({ day: d, pyWd: pw, dstr: dstr(year, month, d) });
    }
    return out;
  }

  /* ==================== قراءة attlog ==================== */

  function decodeBytes(u8) {
    // BOM
    if (u8.length >= 2 && ((u8[0] === 0xff && u8[1] === 0xfe) || (u8[0] === 0xfe && u8[1] === 0xff))) {
      try { return { text: new TextDecoder("utf-16" + (u8[0] === 0xff ? "le" : "be")).decode(u8), encoding: "utf-16" }; } catch (e) {}
    }
    if (u8.length >= 3 && u8[0] === 0xef && u8[1] === 0xbb && u8[2] === 0xbf) {
      try { return { text: new TextDecoder("utf-8").decode(u8), encoding: "utf-8-sig" }; } catch (e) {}
    }
    try { return { text: new TextDecoder("utf-8", { fatal: true }).decode(u8), encoding: "utf-8" }; } catch (e) {}
    try { return { text: new TextDecoder("windows-1256").decode(u8), encoding: "cp1256" }; } catch (e) {}
    return { text: new TextDecoder("latin1").decode(u8), encoding: "latin-1" };
  }

  function validYMD(y, m, d) {
    if (m < 1 || m > 12) return false;
    if (d < 1 || d > monthDaysCount(y, m)) return false;
    return true;
  }

  function parseDT(text) {
    const m = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AaPp])\.?[Mm]\.?$/.exec(text.trim());
    if (m) {
      let h = +m[4];
      const ap = m[7] ? m[7].toLowerCase() : null;
      if (ap === "p" && h < 12) h += 12;
      if (ap === "a" && h === 12) h = 0;
      const y = +m[1], mo = +m[2], d = +m[3];
      if (!validYMD(y, mo, d) || h > 23 || +m[5] > 59 || (+m[6] || 0) > 59) return null;
      return { y, m: mo, d, mins: h * 60 + +m[5] };
    }
    const t = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(text.trim());
    if (t) {
      const y = +t[1], mo = +t[2], d = +t[3];
      if (!validYMD(y, mo, d) || +t[4] > 23 || +t[5] > 59 || (+t[6] || 0) > 59) return null;
      return { y, m: mo, d, mins: +t[4] * 60 + +t[5] };
    }
    return null;
  }

  const NO_TAB_RE = /^\s*(\d+)\s+(.+?)\s+(\d{4}[-/]\d{1,2}[-/]\d{1,2})\s+(\d{1,2}:\d{2}(?::\d{2})?(?:\s*[AaPp][Mm])?)\s+(\S+)\s*(\S*)\s*$/;

  function parseLine(line, lineNo) {
    if (line.indexOf("\t") >= 0) {
      const f = line.split("\t").map(x => x.trim()).filter(x => x !== "");
      if (f.length < 3) return { rec: null, reason: "عدد الحقول أقل من 3" };
      const [emp, name, dtPart] = f;
      if (!/^\d+$/.test(emp)) return { rec: null, reason: "رقم الموظف غير رقمي" };
      const dt = parseDT(dtPart);
      if (!dt) return { rec: null, reason: "تاريخ/وقت غير صالح: '" + dtPart + "'" };
      return { rec: { id: +emp, name, ...dt, s1: f[3] || "", s2: f[4] || "", line: lineNo }, reason: null };
    }
    const m = NO_TAB_RE.exec(line);
    if (!m) return { rec: null, reason: "سطر غير مطابق للصيغة (لا TAB ولا نمط نصي)" };
    const dt = parseDT(m[3] + " " + m[4]);
    if (!dt) return { rec: null, reason: "تاريخ/وقت غير صالح" };
    return { rec: { id: +m[1], name: m[2].trim(), ...dt, s1: m[5], s2: m[6] || "", line: lineNo }, reason: null };
  }

  function parseAttlog(bytes, dedupSeconds) {
    dedupSeconds = dedupSeconds == null ? 120 : dedupSeconds;
    const { text, encoding } = decodeBytes(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes));
    const records = [], corrupt = [];
    let total = 0;
    const lines = text.split(/\r\n|\n|\r/);
    for (let i = 0; i < lines.length; i++) {
      total++;
      const line = lines[i].replace(/\uFEFF/g, "").trim();
      if (!line) continue;
      const { rec, reason } = parseLine(line, i + 1);
      if (!rec) corrupt.push({ no: i + 1, text: line.slice(0, 80), reason });
      else records.push(rec);
    }
    // إزالة التكرار: نفس الموظف بفارق أقل من dedupSeconds
    const epoch = r => Date.UTC(r.y, r.m - 1, r.d) / 60000 + r.mins;
    records.sort((a, b) => a.id - b.id || epoch(a) - epoch(b));
    const out = [];
    let removed = 0;
    for (const r of records) {
      const last = out[out.length - 1];
      if (last && last.id === r.id) {
        const gap = epoch(r) - epoch(last);
        if (gap >= 0 && gap < dedupSeconds) { removed++; continue; }
      }
      out.push(r);
    }
    const names = {};
    for (const r of out) names[r.id] = r.name; // آخر اسم (نفس منطق بايثون)
    return { records: out, corrupt, encoding, duplicates: removed, totalLines: total, idsFound: names };
  }

  /* ==================== معالجة الشهر ==================== */

  function parseHHMM(s, def) {
    const m = /^(\d{1,2}):(\d{2})$/.exec(String(s || "").trim());
    return m ? (+m[1]) * 60 + (+m[2]) : def;
  }

  const NOTE_MISSING_OUT = "نقص بصمة خروج";
  const NOTE_MISSING_IN = "نقص بصمة دخول";

  function processMonth(parseRes, settings, year, month) {
    const midday = parseHHMM(settings.midday_split, 720);
    const lateAfter = parseHHMM(settings.late_after, 495);
    const lateOn = !!settings.late_enabled;
    const workdays = new Set((settings.workdays || [6, 0, 1, 2, 3]).map(Number));
    const notes = settings.notes || {};
    const employees = {};
    (settings.employees || []).forEach(e => { if (e.id != null && e.id !== "") employees[e.id] = e; });

    const rows = new Map();          // "id|dstr" → row
    const nameOf = {};
    const grouped = new Map();       // "id|dstr" → [rec]

    for (const r of parseRes.records) {
      if (r.y !== year || r.m !== month) continue;
      const k = r.id + "|" + dstr(r.y, r.m, r.d);
      if (!grouped.has(k)) grouped.set(k, []);
      grouped.get(k).push(r);
      if (nameOf[r.id] == null) nameOf[r.id] = r.name;
    }

    const keys = [...grouped.keys()].sort((a, b) => {
      const [id1, d1] = a.split("|"), [id2, d2] = b.split("|");
      return +id1 - +id2 || d1.localeCompare(d2);
    });

    for (const k of keys) {
      const [idS, ds] = k.split("|");
      const times = grouped.get(k).slice().sort((a, b) => a.mins - b.mins);
      const id = +idS;
      const [y, m, dnum] = ds.split("-").map(Number);
      const row = { empId: id, d: ds, y, m, dnum, punches: times.length,
                    deviceName: nameOf[id] || "", templateName: "",
                    inMins: null, outMins: null, sysNote: "", savedNote: "", note: "", late: false };
      const emp = employees[id];
      row.templateName = (emp && emp.template_name) || "";
      const saved = notes[ds + "|" + id] || "";
      row.savedNote = saved;

      if (times.length >= 2) {
        row.inMins = times[0].mins;
        row.outMins = times[times.length - 1].mins;
        const middles = times.slice(1, -1);
        if (middles.length) row.sysNote = "خروج وعودة: " + middles.map(t => fmtTime(t.mins)).join(" ، ");
      } else {
        const t = times[0];
        if (t.mins < midday) { row.inMins = t.mins; row.sysNote = NOTE_MISSING_OUT; }
        else { row.outMins = t.mins; row.sysNote = NOTE_MISSING_IN; }
      }
      if (lateOn && row.inMins != null && row.inMins > lateAfter) row.late = true;
      row.note = [saved, row.sysNote].filter(Boolean).join(" / ");
      rows.set(k, row);
    }

    // ملاحظات محفوظة بلا أي بصمة (اجازة كاملة مثلاً)
    for (const key of Object.keys(notes)) {
      const parts = key.split("|");
      if (parts.length !== 2) continue;
      const ds = parts[0], id = +parts[1];
      if (!/^\d{4}-\d{2}-\d{2}$/.test(ds) || isNaN(id)) continue;
      const [y, m, dnum] = ds.split("-").map(Number);
      if (y !== year || m !== month) continue;
      if (rows.has(key)) continue;
      const emp = employees[id];
      if (!emp || !workdays.has(pyWeekday(y, m, dnum))) continue;
      const row = { empId: id, d: ds, y, m, dnum, punches: 0, deviceName: "",
                    templateName: emp.template_name || "", inMins: null, outMins: null,
                    sysNote: "", savedNote: notes[key], note: notes[key], late: false };
      rows.set(key, row);
    }

    // الأرقام غير المربوطة
    const unmapped = {};
    for (const id of Object.keys(nameOf)) {
      if (!employees[+id]) unmapped[id] = nameOf[id];
    }

    function summary() {
      const perEmp = {};
      for (const row of rows.values()) {
        const s = perEmp[row.empId] || (perEmp[row.empId] = { name: row.templateName || row.deviceName, days: 0, problems: 0 });
        s.days++;
        if (row.sysNote) s.problems++;
      }
      return { rowsCount: rows.size, perEmp, unmapped };
    }

    return { rows, unmapped, summary };
  }

  /* ==================== ZIP (قراءة) ==================== */

  async function unzip(buf) {
    const u8 = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
    // ابحث عن EOCD من النهاية
    let eocd = -1;
    const from = Math.max(0, u8.length - 66000);
    for (let i = u8.length - 22; i >= from; i--) {
      if (u8[i] === 0x50 && u8[i + 1] === 0x4b && u8[i + 2] === 0x05 && u8[i + 3] === 0x06) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error("الملف ليس بصيغة xlsx الحديثة (أرشيف ZIP غير صالح)");
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    const count = dv.getUint16(eocd + 10, true);
    const cdSize = dv.getUint32(eocd + 12, true);
    const cdOff = dv.getUint32(eocd + 16, true);
    const td = new TextDecoder();
    const files = {};
    let p = cdOff;
    for (let i = 0; i < count; i++) {
      if (dv.getUint32(p, true) !== 0x02014b50) break;
      const method = dv.getUint16(p + 10, true);
      const compSize = dv.getUint32(p + 20, true);
      const nameLen = dv.getUint16(p + 28, true);
      const extraLen = dv.getUint16(p + 30, true);
      const commLen = dv.getUint16(p + 32, true);
      const localOff = dv.getUint32(p + 42, true);
      const name = td.decode(u8.subarray(p + 46, p + 46 + nameLen));
      // رأس الملف المحلي (أطواله قد تختلف عن المركزي)
      const lnLen = dv.getUint16(localOff + 26, true);
      const leLen = dv.getUint16(localOff + 28, true);
      const dataStart = localOff + 30 + lnLen + leLen;
      const comp = u8.subarray(dataStart, dataStart + compSize);
      if (method === 0) {
        files[name] = comp.slice();
      } else if (method === 8) {
        if (typeof DecompressionStream === "undefined")
          throw new Error("المتصفح لا يدعم فك ضغط الملفات — استخدم Chrome/Edge حديثاً أو تطبيق سطح المكتب");
        const ds = new DecompressionStream("deflate-raw");
        const ab = await new Response(new Blob([comp]).stream().pipeThrough(ds)).arrayBuffer();
        files[name] = new Uint8Array(ab);
      } else {
        throw new Error("طريقة ضغط غير مدعومة داخل الملف (" + method + ")");
      }
      p += 46 + nameLen + extraLen + commLen;
    }
    if (!Object.keys(files).length) throw new Error("تعذر قراءة محتويات الملف — قد يكون تالفاً");
    return files;
  }

  /* ==================== XML / جدول البيانات ==================== */

  let _parser = null;
  function getParser() {
    if (!_parser) _parser = new DOMParser();
    return _parser;
  }
  const serializer = () => new XMLSerializer();
  const MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

  function tags(root, local) {
    const out = [];
    const all = root.getElementsByTagName("*");
    for (let i = 0; i < all.length; i++) {
      if (all[i].localName === local) out.push(all[i]);
    }
    return out;
  }

  function parseXml(u8) {
    const text = new TextDecoder("utf-8").decode(u8);
    const doc = getParser().parseFromString(text, "application/xml");
    if (doc.getElementsByTagName("parsererror").length)
      throw new Error("XML تالف داخل الملف");
    return doc;
  }

  function colToRef(c) {
    let s = "";
    while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = (c - 1 - m) / 26 | 0; }
    return s;
  }
  function refToCol(ref) {
    let c = 0;
    for (const ch of ref) {
      if (ch >= "A" && ch <= "Z") c = c * 26 + ch.charCodeAt(0) - 64;
      else if (ch >= "a" && ch <= "z") c = c * 26 + ch.charCodeAt(0) - 96;
      else break;
    }
    return c;
  }
  function splitRef(ref) {
    const m = /^([A-Za-z]+)(\d+)$/.exec(ref || "");
    return m ? { row: +m[2], col: refToCol(m[1]) } : null;
  }
  function parseRange(ref) {
    const parts = String(ref || "").split(":");
    const a = splitRef(parts[0]), b = splitRef(parts[1] || parts[0]);
    return a && b ? { r1: a.row, c1: a.col, r2: b.row, c2: b.col } : null;
  }

  // تاريخ Excel التسلسلي (نظام 1900)
  function serialToDate(serial) {
    return new Date(Math.round((serial - 25569) * 86400000));
  }
  function datePartsToSerial(y, m, d) {
    return (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000;
  }

  function parseSST(u8) {
    if (!u8) return [];
    const doc = parseXml(u8);
    const out = [];
    for (const si of tags(doc, "si")) {
      let s = "";
      for (const t of tags(si, "t")) s += t.textContent || "";
      out.push(s);
    }
    return out;
  }

  function parseDateStyles(u8) {
    const dateStyles = new Set();
    if (!u8) return dateStyles;
    const doc = parseXml(u8);
    const custom = {};
    for (const nf of tags(doc, "numFmt")) {
      custom[+nf.getAttribute("numFmtId")] = nf.getAttribute("formatCode") || "";
    }
    const isDateFmt = id => {
      if (custom[id] != null) return /([dy])|(m{1,2}[^h])|h.*m/i.test(custom[id].replace(/\\.|"[^"]*"/g, "")) && /d|y/i.test(custom[id]);
      return (id >= 14 && id <= 22) || (id >= 27 && id <= 36) || (id >= 45 && id <= 47) || (id >= 50 && id <= 58);
    };
    for (const xfs of tags(doc, "cellXfs")) {
      let i = 0;
      for (const xf of tags(xfs, "xf")) {
        const id = +xf.getAttribute("numFmtId") || 0;
        if (id && isDateFmt(id)) dateStyles.add(i);
        i++;
      }
    }
    return dateStyles;
  }

  function parseSheet(u8, sst, dateStyles) {
    const doc = parseXml(u8);
    const cells = new Map();     // "r:c" → {row, col, v, s}
    const merges = [];
    let maxRow = 0, maxCol = 0;
    for (const mc of tags(doc, "mergeCell")) {
      const r = parseRange(mc.getAttribute("ref"));
      if (r) merges.push(r);
    }
    for (const c of tags(doc, "c")) {
      const pos = splitRef(c.getAttribute("r"));
      if (!pos) continue;
      const t = c.getAttribute("t") || "";
      const s = +(c.getAttribute("s") || 0);
      let v = null;
      if (t === "inlineStr") {
        v = tags(c, "t").map(x => x.textContent || "").join("");
      } else {
        const ve = tags(c, "v")[0];
        const raw = ve ? (ve.textContent || "") : null;
        if (t === "s") v = sst[+raw] != null ? sst[+raw] : "";
        else if (t === "str") v = raw != null ? raw : "";
        else if (t === "b") v = raw === "1";
        else if (raw != null && raw !== "") {
          const num = parseFloat(raw);
          if (!isNaN(num)) {
            if (dateStyles.has(s)) v = serialToDate(num);
            else v = num;
          } else v = raw;
        }
      }
      cells.set(pos.row + ":" + pos.col, { row: pos.row, col: pos.col, v, s });
      if (pos.row > maxRow) maxRow = pos.row;
      if (pos.col > maxCol) maxCol = pos.col;
    }
    // امتداد الدمج الأفقي لكل خلية
    const hspan = new Map();
    for (const m of merges) {
      if (m.r1 === m.r2 && m.c2 > m.c1) hspan.set(m.r1 + ":" + m.c1, { c0: m.c1, c1: m.c2 });
    }
    const getCell = (row, col) => cells.get(row + ":" + col) || null;
    const mergeSpan = (row, col) => hspan.get(row + ":" + col) || { c0: col, c1: col };
    return { doc, cells, merges, hspan, maxRow, maxCol, getCell, mergeSpan,
             sheetName: null, sheetPath: null };
  }

  function cellText(cell) {
    if (cell == null || cell.v == null) return "";
    if (cell.v instanceof Date) {
      return dstr(cell.v.getUTCFullYear(), cell.v.getUTCMonth() + 1, cell.v.getUTCDate());
    }
    return String(cell.v);
  }

  /* ==================== تحليل القالب ==================== */

  const MONTH_WORDS = {
    "يناير": 1, "كانون الثاني": 1, "فبراير": 2, "شباط": 2,
    "مارس": 3, "اذار": 3, "أبريل": 4, "ابريل": 4, "نيسان": 4,
    "مايو": 5, "ايار": 5, "يونيو": 6, "حزيران": 6,
    "يوليو": 7, "تموز": 7, "اغسطس": 8, "اب": 8,
    "سبتمبر": 9, "ايلول": 9, "أكتوبر": 10, "اكتوبر": 10, "تشرين الاول": 10,
    "نوفمبر": 11, "تشرين الثاني": 11, "ديسمبر": 12, "كانون الاول": 12,
  };

  function wordMonth(word) {
    const w = normalizeAr(word);
    if (!w) return null;
    if (MONTH_WORDS[w]) return MONTH_WORDS[w];
    if (w.indexOf("تشرين") >= 0) return w.indexOf("اول") >= 0 ? 10 : (w.indexOf("ثاني") >= 0 ? 11 : null);
    if (w.indexOf("كانون") >= 0) return w.indexOf("اول") >= 0 ? 1 : (w.indexOf("ثاني") >= 0 ? 12 : null);
    for (const name of Object.keys(MONTH_WORDS)) {
      const n = normalizeAr(name);
      if (n && (n.indexOf(w) >= 0 || w.indexOf(n) >= 0) && Math.abs(n.length - w.length) <= 2) return MONTH_WORDS[name];
    }
    return null;
  }

  function parseTextDate(v, year, month) {
    if (typeof v !== "string") return null;
    const s0 = normDigits(v.trim());
    if (!s0 || s0.length > 40 || /^\d+$/.test(s0)) return null;
    const s = s0.replace(/^(ال)?(سبت|احد|أحد|اثنين|إثنين|ثلاثاء|أربعاء|اربعاء|خميس|جمعة)[\s،,\-]+/, "");
    let m = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/.exec(s);
    if (m) return { y: +m[1], m: +m[2], d: +m[3] };
    m = /^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/.exec(s);
    if (m) {
      let y = +m[3]; if (y < 100) y += 2000;
      return { y, m: +m[2], d: +m[1] };
    }
    m = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/.exec(s.replace(/\./g, "/"));
    if (m) return { y: +m[1], m: +m[2], d: +m[3] };
    // «15 سبتمبر» أو «15 سبتمبر 2026»
    m = /^(\d{1,2})\s+([\u0621-\u064A][\u0621-\u064A ]*?)(?:\s+(\d{4}))?$/.exec(s);
    if (m) {
      const day = +m[1];
      const mon = wordMonth(m[2]);
      const yr = m[3] ? +m[3] : year;
      if (mon && day >= 1 && day <= 31) {
        const n = monthDaysCount(yr, mon);
        if (day <= n) return { y: yr, m: mon, d: day };
      }
    }
    return null;
  }

  const LABELS = {
    in: ["دخول", "حضور", "بداية"],
    out: ["خروج", "انصراف", "نهاية"],
    note: ["ملاحظات", "ملاحظة", "ملحوظة"],
  };

  function findLabels(model, row, c0, c1) {
    const found = {};
    for (let rr = row + 1; rr <= Math.min(row + 4, model.maxRow); rr++) {
      for (let cc = c0; cc <= Math.min(c1 + 4, model.maxCol); cc++) {
        const v = cellText(model.getCell(rr, cc));
        const nv = normalizeAr(v);
        if (!nv || nv.length > 30) continue;
        for (const key of Object.keys(LABELS)) {
          if (!(key in found) && LABELS[key].some(w => nv.indexOf(normalizeAr(w)) >= 0)) found[key] = cc;
        }
      }
      if ("in" in found && "out" in found) break;
    }
    return found;
  }

  function subMapFor(model, row, c0, c1) {
    const width = c1 - c0 + 1;
    const labels = findLabels(model, row, c0, c1);
    if ("in" in labels && "out" in labels)
      return { in: labels.in, out: labels.out, note: labels.note != null ? labels.note : labels.out + 1, combined: false };
    if (width === 1) return { in: c0, out: c0, combined: true };
    if (width === 2) return { in: c0, out: c0 + 1, combined: true };
    return { in: c0, out: c0 + 1, note: labels.note != null ? labels.note : c0 + 2, combined: false };
  }

  function dayCellsOfRow(model, row) {
    const cells = [];
    for (const cell of model.cells.values()) {
      if (cell.row !== row) continue;
      const d = toInt(cell.v);
      if (d != null && d >= 1 && d <= 31) cells.push({ col: cell.col, d });
    }
    cells.sort((a, b) => a.col - b.col);
    return cells;
  }

  function increasing(cells) {
    for (let i = 0; i < cells.length - 1; i++) if (cells[i + 1].d <= cells[i].d) return false;
    return true;
  }

  function anchorsFromDayRows(model, chosen, year, month, seen) {
    const anchors = {};
    for (const row of chosen) {
      const cells = dayCellsOfRow(model, row);
      for (let i = 0; i < cells.length; i++) {
        const { col, d } = cells[i];
        if (seen.has(d)) continue;
        const nxt = i + 1 < cells.length ? cells[i + 1].col : null;
        const span = model.mergeSpan(row, col);
        let c0 = span.c0, c1 = span.c1;
        if (c1 === c0 && nxt) c1 = Math.min(c0 + Math.max(nxt - col, 1) - 1, c0 + 2);
        const n = monthDaysCount(year, month);
        if (d > n) continue;
        seen.add(d);
        const ds = dstr(year, month, d);
        anchors[ds] = { row, c0, c1, subMap: subMapFor(model, row, c0, c1) };
      }
    }
    return anchors;
  }

  function collectAnchors(model, year, month) {
    const messages = [];
    const anchors = {};
    const otherMonths = new Set();

    // 1) تواريخ حقيقية / نصية
    for (const cell of model.cells.values()) {
      let dt = null;
      if (cell.v instanceof Date) {
        dt = { y: cell.v.getUTCFullYear(), m: cell.v.getUTCMonth() + 1, d: cell.v.getUTCDate() };
      } else if (typeof cell.v === "string") {
        dt = parseTextDate(cell.v, year, month);
      }
      if (!dt) continue;
      if (dt.y === year && dt.m === month) {
        const span = model.mergeSpan(cell.row, cell.col);
        const ds = dstr(dt.y, dt.m, dt.d);
        if (!anchors[ds])
          anchors[ds] = { row: cell.row, c0: span.c0, c1: span.c1, subMap: subMapFor(model, cell.row, span.c0, span.c1) };
      } else {
        otherMonths.add(dt.y + "-" + String(dt.m).padStart(2, "0"));
      }
    }
    if (Object.keys(anchors).length) {
      messages.push("تم العثور على " + Object.keys(anchors).length + " يوم لشهر الهدف في القالب.");
      return { anchors, mode: "تعبئة نسخة القالب مباشرة (تواريخ حقيقية)", messages };
    }

    // 2) صفوف أرقام الأيام
    const dayRows = [];
    for (let r = 1; r <= model.maxRow; r++) {
      const cells = dayCellsOfRow(model, r);
      if (cells.length) dayRows.push({ row: r, cells });
    }
    const qualifying = dayRows.filter(rc => rc.cells.length >= 4 && increasing(rc.cells));
    if (qualifying.length) {
      const full = qualifying.filter(rc => rc.cells.length >= 15);
      let chosen;
      if (full.length) {
        full.sort((a, b) => b.cells.length - a.cells.length || a.row - b.row);
        chosen = [full[0]];
      } else chosen = qualifying;
      const seen = new Set();
      let found = anchorsFromDayRows(model, chosen.map(rc => rc.row), year, month, seen);
      // استكمال آخر أسبوع من صفوف لاحقة
      if (Object.keys(found).length) {
        const lastRow = Math.max(...Object.values(found).map(a => a.row));
        const maxDay = Math.max(...Object.keys(found).map(ds => +ds.split("-")[2]));
        const cont = dayRows.filter(rc => rc.row > lastRow && rc.cells.length && increasing(rc.cells)
          && Math.max(...rc.cells.map(c => c.d)) > maxDay).map(rc => rc.row);
        if (cont.length) {
          const extra = anchorsFromDayRows(model, cont, year, month, seen);
          for (const ds of Object.keys(extra)) {
            const day = +ds.split("-")[2];
            if (day > maxDay && !found[ds]) found[ds] = extra[ds];
          }
          if (Math.max(...Object.keys(found).map(ds => +ds.split("-")[2])) > maxDay)
            messages.push("تم استكمال الأيام من صفوف لاحقة حتى اليوم " + Math.max(...Object.keys(found).map(ds => +ds.split("-")[2])) + ".");
        }
      }
      if (Object.keys(found).length) {
        for (const ds of Object.keys(found)) anchors[ds] = found[ds];
        messages.push("تم اكتشاف أرقام أيام الشهر (" + Object.keys(anchors).length + " يوماً) في الصف: " + chosen.map(rc => rc.row).slice(0, 6).join("، ") + ".");
        return { anchors, mode: "تعبئة نسخة القالب مباشرة (صف أرقام الأيام)", messages };
      }
    }

    if (otherMonths.size) {
      messages.push("القالب يحتوي تواريخ لأشهر أخرى (" + [...otherMonths].sort().slice(0, 6).join("، ") + ") وليس الشهر المطلوب.");
    }
    return { anchors: {}, mode: "", messages };
  }

  function isTitleCell(model, row, col) {
    const span = model.mergeSpan(row, col);
    return span.c1 - span.c0 + 1 > 3;
  }

  const HEADER_WORDS = ["دخول", "حضور", "انصراف", "خروج", "ملاحظ", "ملحوظ",
                        "نهاية", "بداية", "الاسم", "الموظف", "اليوم", "التاريخ"];

  function isHeaderWord(v) {
    const nv = normalizeAr(v);
    if (nv === "م" || nv === "ت") return true;
    return HEADER_WORDS.some(w => nv.indexOf(normalizeAr(w)) >= 0);
  }

  function findNameCells(model) {
    const out = [];
    for (const cell of model.cells.values()) {
      if (cell.col > 2) continue;
      const v = cell.v;
      if (typeof v !== "string") continue;
      const t = v.trim();
      if (t.length < 3 || t.length > 60) continue;
      if (isTitleCell(model, cell.row, cell.col)) continue;
      if (isHeaderWord(t)) continue;
      out.push({ row: cell.row, nv: normalizeAr(t), orig: t });
    }
    return out;
  }

  function matchEmployeeRows(model, employees) {
    const nameCells = findNameCells(model);
    const empRows = new Map();
    const unmatched = [];
    for (const emp of employees) {
      const tn = emp.template_name || "";
      if (!tn) continue;
      const ntn = normalizeAr(tn);
      let hit = null;
      for (const nc of nameCells) if (nc.nv === ntn) { hit = nc.row; break; }
      if (hit == null) {
        for (const nc of nameCells) if (ntn && (nc.nv.indexOf(ntn) >= 0 || ntn.indexOf(nc.nv) >= 0)) { hit = nc.row; break; }
      }
      if (hit != null) empRows.set(emp.id, hit);
      else unmatched.push(tn);
    }
    return { empRows, unmatched };
  }

  function firstSheetInfo(files) {
    const wbDoc = parseXml(files["xl/workbook.xml"]);
    const sheets = tags(wbDoc, "sheet");
    if (!sheets.length) throw new Error("لا توجد أوراق في ملف القالب");
    const name = sheets[0].getAttribute("name") || "Sheet1";
    let rid = sheets[0].getAttribute("r:id") || sheets[0].getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id") || "";
    let target = "";
    const relsDoc = parseXml(files["xl/_rels/workbook.xml.rels"]);
    for (const rel of tags(relsDoc, "Relationship")) {
      if (rel.getAttribute("Id") === rid) { target = rel.getAttribute("Target") || ""; break; }
    }
    if (!target) throw new Error("تعذر تحديد ورقة العمل الأولى في القالب");
    let path = target.replace(/^\//, "");
    if (!/^xl\//.test(path)) path = "xl/" + path;
    return { name, path, rid };
  }

  /* ==================== ZIP (كتابة — STORE) ==================== */

  const CRC_TABLE = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      t[n] = c >>> 0;
    }
    return t;
  })();

  function crc32(u8) {
    let c = 0xFFFFFFFF;
    for (let i = 0; i < u8.length; i++) c = CRC_TABLE[(c ^ u8[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  function zipStore(files) {
    const enc = new TextEncoder();
    const names = Object.keys(files);
    const parts = [], central = [];
    let offset = 0;
    const dosTime = (9 << 11);                            // 09:00:00
    const dosDate = ((2026 - 1980) << 9) | (9 << 5) | 1;  // 2026-09-01
    for (const name of names) {
      const nb = enc.encode(name);
      const data = files[name] instanceof Uint8Array ? files[name] : enc.encode(String(files[name]));
      const crc = crc32(data);
      const lh = new DataView(new ArrayBuffer(30));
      lh.setUint32(0, 0x04034b50, true);
      lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true);
      lh.setUint16(8, 0, true); lh.setUint16(10, dosTime, true); lh.setUint16(12, dosDate, true);
      lh.setUint32(14, crc, true);
      lh.setUint32(18, data.length, true); lh.setUint32(22, data.length, true);
      lh.setUint16(26, nb.length, true); lh.setUint16(28, 0, true);
      parts.push(new Uint8Array(lh.buffer), nb, data);
      central.push({ nb, crc, size: data.length, offset });
      offset += 30 + nb.length + data.length;
    }
    const cdStart = offset;
    const cdParts = [];
    let cdSize = 0;
    for (const e of central) {
      const ch = new DataView(new ArrayBuffer(46));
      ch.setUint32(0, 0x02014b50, true);
      ch.setUint16(4, 20, true); ch.setUint16(6, 20, true); ch.setUint16(8, 0x0800, true);
      ch.setUint16(10, 0, true); ch.setUint16(12, dosTime, true); ch.setUint16(14, dosDate, true);
      ch.setUint32(16, e.crc, true);
      ch.setUint32(20, e.size, true); ch.setUint32(24, e.size, true);
      ch.setUint16(28, e.nb.length, true);
      ch.setUint32(42, e.offset, true);
      cdParts.push(new Uint8Array(ch.buffer), e.nb);
      cdSize += 46 + e.nb.length;
    }
    const eocd = new DataView(new ArrayBuffer(22));
    eocd.setUint32(0, 0x06054b50, true);
    eocd.setUint16(8, central.length, true);
    eocd.setUint16(10, central.length, true);
    eocd.setUint32(12, cdSize, true);
    eocd.setUint32(16, cdStart, true);
    const total = new Uint8Array(cdStart + cdSize + 22);
    let at = 0;
    for (const part of [...parts, ...cdParts, new Uint8Array(eocd.buffer)]) {
      total.set(part, at);
      at += part.length;
    }
    return total;
  }

  /* ==================== توليد xlsx: تعبئة نسخة القالب ==================== */

  function ensureDecl(xml) {
    return /^<\?xml/.test(xml) ? xml :
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + xml;
  }

  function sheetNsEl(doc, name) {
    return doc.createElementNS(MAIN_NS, name);
  }

  function setCellInDoc(model, row, col, text) {
    // لا نكتب قيماً فارغة
    if (text == null || String(text) === "") return false;
    const doc = model.doc;
    const sheetData = tags(doc, "sheetData")[0];
    if (!sheetData) return false;
    const ref = colToRef(col) + row;

    // صف موجود أو جديد (بترتيب تصاعدي)
    let rowEl = null;
    for (const r of tags(sheetData, "row")) {
      const rn = +(r.getAttribute("r") || 0);
      if (rn === row) { rowEl = r; break; }
      if (rn > row) break;
    }
    if (!rowEl) {
      rowEl = sheetNsEl(doc, "row");
      rowEl.setAttribute("r", String(row));
      let inserted = false;
      for (const r of tags(sheetData, "row")) {
        if (+(r.getAttribute("r") || 0) > row) {
          sheetData.insertBefore(rowEl, r);
          inserted = true;
          break;
        }
      }
      if (!inserted) sheetData.appendChild(rowEl);
    }

    // خلية موجودة أو جديدة (بترتيب الأعمدة)
    let cellEl = null;
    const kids = [];
    for (let i = 0; i < rowEl.childNodes.length; i++) {
      const n = rowEl.childNodes[i];
      if (n.nodeType === 1 && n.localName === "c") kids.push(n);
    }
    for (const c of kids) {
      const pos = splitRef(c.getAttribute("r") || "");
      if (pos && pos.col === col) { cellEl = c; break; }
      if (pos && pos.col > col) break;
    }
    if (!cellEl) {
      cellEl = sheetNsEl(doc, "c");
      cellEl.setAttribute("r", ref);
      let inserted = false;
      for (const c of kids) {
        const pos = splitRef(c.getAttribute("r") || "");
        if (pos && pos.col > col) {
          rowEl.insertBefore(cellEl, c);
          inserted = true;
          break;
        }
      }
      if (!inserted) rowEl.appendChild(cellEl);
    }

    // نص مضمن (inline string) — نحافظ على تنسيق الخلية s كما هو
    while (cellEl.firstChild) cellEl.removeChild(cellEl.firstChild);
    cellEl.setAttribute("t", "inlineStr");
    const isEl = sheetNsEl(doc, "is");
    const tEl = sheetNsEl(doc, "t");
    tEl.setAttribute("xml:space", "preserve");
    tEl.textContent = String(text);
    isEl.appendChild(tEl);
    cellEl.appendChild(isEl);
    return true;
  }

  async function fillTemplateBytes(tplBytes, att, employees, year, month) {
    const files = await unzip(tplBytes);
    const info = firstSheetInfo(files);
    const sst = parseSST(files["xl/sharedStrings.xml"]);
    const dateStyles = parseDateStyles(files["xl/styles.xml"]);
    const model = parseSheet(files[info.path], sst, dateStyles);
    model.sheetName = info.name;

    const res = collectAnchors(model, year, month);
    if (!Object.keys(res.anchors).length) {
      const err = new Error("لم أتعرف على أعمدة الأيام في القالب. " + res.messages.join(" "));
      err.messages = res.messages;
      throw err;
    }
    const match = matchEmployeeRows(model, employees);

    let written = 0, combinedCells = 0;
    for (const row of att.rows.values()) {
      const ds = row.d;
      const anchor = res.anchors[ds];
      const excelRow = match.empRows.get(row.empId);
      if (!anchor || !excelRow) continue;
      const sm = anchor.subMap;
      if (sm.combined) {
        const txt = [row.inMins != null ? fmtTime(row.inMins) : "",
                     row.outMins != null ? fmtTime(row.outMins) : ""].filter(Boolean).join(" - ");
        if (txt) { written += setCellInDoc(model, excelRow, sm.in, txt) ? 1 : 0; combinedCells++; }
        if (row.note && sm.note) written += setCellInDoc(model, excelRow, sm.note, row.note) ? 1 : 0;
        continue;
      }
      if (row.inMins != null) written += setCellInDoc(model, excelRow, sm.in, fmtTime(row.inMins)) ? 1 : 0;
      if (row.outMins != null) written += setCellInDoc(model, excelRow, sm.out, fmtTime(row.outMins)) ? 1 : 0;
      if (row.note && sm.note) written += setCellInDoc(model, excelRow, sm.note, row.note) ? 1 : 0;
    }

    // تسلسل حفظ: تعديل ورقة العمل فقط + إزالة calcChain للحيلولة دون تحذيرات Excel
    const ser = new XMLSerializer().serializeToString(model.doc);
    files[info.path] = new TextEncoder().encode(ensureDecl(ser));
    if (files["xl/calcChain.xml"]) {
      delete files["xl/calcChain.xml"];
      if (files["[Content_Types].xml"]) {
        files["[Content_Types].xml"] = new TextEncoder().encode(
          new TextDecoder().decode(files["[Content_Types].xml"])
            .replace(/<Override[^>]*calcChain[^>]*\/>\s*/g, ""));
      }
      if (files["xl/_rels/workbook.xml.rels"]) {
        files["xl/_rels/workbook.xml.rels"] = new TextEncoder().encode(
          new TextDecoder().decode(files["xl/_rels/workbook.xml.rels"])
            .replace(/<Relationship[^>]*calcChain[^>]*\/>\s*/g, ""));
      }
    }

    const messages = res.messages.slice();
    messages.push("تم ربط " + match.empRows.size + " موظفاً بصفوف القالب.");
    if (match.unmatched.length)
      messages.push("أسماء لم تُعثر على صف لها: " + match.unmatched.slice(0, 12).join(" ، "));
    messages.push("تمت كتابة " + written + " خلية في نسخة القالب." +
      (combinedCells ? " (" + combinedCells + " خلية مدمجة دخول-خروج)" : ""));
    return { bytes: zipStore(files), messages, written, matched: match.empRows.size, mode: res.mode };
  }

  /* ==================== توليد xlsx: المولّد المدمج ==================== */

  function buildFallbackBytes(att, settings, year, month) {
    const workdays = (settings.workdays || [6, 0, 1, 2, 3]).map(Number);
    const employees = settings.employees || [];
    const days = monthWorkdays(year, month, workdays);
    const monthName = MONTH_AR[month] || month;
    const lastCol = 1 + Math.max(1, days.length) * 3;

    const esc = escapeXml;
    let rowsXml = "";
    const merges = [];
    const colRef = c => colToRef(c);

    // صف 1: العنوان
    {
      const cells = '<c r="A1" s="1" t="inlineStr"><is><t xml:space="preserve">' +
        esc("جدول الدوام لشهر " + monthName + " " + year) + '</t></is></c>';
      rowsXml += '<row r="1" ht="22" customHeight="1">' + cells + '</row>';
      merges.push("A1:" + colRef(lastCol) + "1");
    }
    // صف 2: تاريخ كل يوم مختار مدموجاً على 3 خلايا
    {
      let cells = "";
      days.forEach((dw, i) => {
        const col = 2 + i * 3;
        const serial = datePartsToSerial(year, month, dw.day);
        cells += '<c r="' + colRef(col) + '2" s="2"><v>' + serial + '</v></c>';
        merges.push(colRef(col) + "2:" + colRef(col + 2) + "2");
      });
      rowsXml += '<row r="2">' + cells + '</row>';
    }
    // صف 3: الرؤوس
    {
      let cells = "";
      days.forEach((dw, i) => {
        const col = 2 + i * 3;
        ["ساعة الدخول", "ساعة الخروج", "ملاحظات"].forEach((lbl, off) => {
          cells += '<c r="' + colRef(col + off) + '3" s="3" t="inlineStr"><is><t>' + esc(lbl) + '</t></is></c>';
        });
      });
      rowsXml += '<row r="3">' + cells + '</row>';
    }
    // الصفوف 5+: الأسماء والقيم
    let r = 5;
    for (const emp of employees) {
      const name = emp.template_name || (emp.id ? "(بدون ربط) #" + emp.id : "");
      let cells = '<c r="A' + r + '" s="6" t="inlineStr"><is><t xml:space="preserve">' + esc(name) + '</t></is></c>';
      if (emp.no_punch) {
        for (let c = 2; c <= lastCol; c++) cells += '<c r="' + colRef(c) + r + '" s="7"/>';
      } else if (emp.id) {
        days.forEach((dw, i) => {
          const row = att.rows.get(emp.id + "|" + dw.dstr);
          const col = 2 + i * 3;
          if (row) {
            if (row.inMins != null)
              cells += '<c r="' + colRef(col) + r + '" s="4" t="inlineStr"><is><t>' + esc(fmtTime(row.inMins)) + '</t></is></c>';
            if (row.outMins != null)
              cells += '<c r="' + colRef(col + 1) + r + '" s="4" t="inlineStr"><is><t>' + esc(fmtTime(row.outMins)) + '</t></is></c>';
            if (row.note)
              cells += '<c r="' + colRef(col + 2) + r + '" s="5" t="inlineStr"><is><t xml:space="preserve">' + esc(row.note) + '</t></is></c>';
          }
          for (let c = col; c <= col + 2; c++) {
            if (cells.indexOf('r="' + colRef(c) + r + '"') < 0) cells += '<c r="' + colRef(c) + r + '" s="4"/>';
          }
        });
      }
      rowsXml += '<row r="' + r + '">' + cells + '</row>';
      r++;
    }

    // أعمدة
    let colsXml = '<col min="1" max="1" width="18" customWidth="1"/>';
    days.forEach((dw, i) => {
      const col = 2 + i * 3;
      colsXml += '<col min="' + col + '" max="' + (col + 1) + '" width="10" customWidth="1"/>';
      colsXml += '<col min="' + (col + 2) + '" max="' + (col + 2) + '" width="14" customWidth="1"/>';
    });

    const lastColRef = colRef(lastCol);
    const sheetXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
      '<worksheet xmlns="' + MAIN_NS + '">' +
      '<dimension ref="A1:' + lastColRef + r + '"/>' +
      '<sheetViews><sheetView rightToLeft="1" tabSelected="1" workbookViewId="0">' +
      '<pane xSplit="1" ySplit="4" topLeftCell="B5" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>' +
      '<sheetFormatPr defaultRowHeight="15"/>' +
      '<cols>' + colsXml + '</cols>' +
      '<sheetData>' + rowsXml + '</sheetData>' +
      '<mergeCells count="' + merges.length + '">' +
      merges.map(m => '<mergeCell ref="' + m + '"/>').join("") + '</mergeCells>' +
      '</worksheet>';

    const stylesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
      '<styleSheet xmlns="' + MAIN_NS + '">' +
      '<numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy\\-mm\\-dd"/></numFmts>' +
      '<fonts count="3">' +
      '<font><sz val="11"/><color theme="1"/><name val="Calibri"/></font>' +
      '<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/></font>' +
      '<font><b/><sz val="14"/><color theme="1"/><name val="Calibri"/></font>' +
      '</fonts>' +
      '<fills count="4">' +
      '<fill><patternFill patternType="none"/></fill>' +
      '<fill><patternFill patternType="gray125"/></fill>' +
      '<fill><patternFill patternType="solid"><fgColor rgb="FFD9D9D9"/><bgColor indexed="64"/></patternFill></fill>' +
      '<fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/><bgColor indexed="64"/></patternFill></fill>' +
      '</fills>' +
      '<borders count="2">' +
      '<border><left/><right/><top/><bottom/><diagonal/></border>' +
      '<border><left style="thin"><color rgb="FF808080"/></left><right style="thin"><color rgb="FF808080"/></right>' +
      '<top style="thin"><color rgb="FF808080"/></top><bottom style="thin"><color rgb="FF808080"/></bottom><diagonal/></border>' +
      '</borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="8">' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +                                          // 0
      '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1">' +          // 1 عنوان
      '<alignment horizontal="center" vertical="center"/></xf>' +
      '<xf numFmtId="164" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyAlignment="1">' + // 2 تاريخ
      '<alignment horizontal="center" vertical="center"/></xf>' +
      '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">' + // 3 رأس
      '<alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">' +        // 4 خلية
      '<alignment horizontal="center" vertical="center"/></xf>' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">' +        // 5 ملاحظة
      '<alignment horizontal="right" vertical="center"/></xf>' +
      '<xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1">' + // 6 اسم
      '<alignment horizontal="right" vertical="center"/></xf>' +
      '<xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyBorder="1"/>' +            // 7 بلا بصمة
      '</cellXfs>' +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      '</styleSheet>';

    const files = {
      "[Content_Types].xml":
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
        '</Types>',
      "_rels/.rels":
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
        '</Relationships>',
      "xl/workbook.xml":
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
        '<workbook xmlns="' + MAIN_NS + '" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<bookViews><workbookView/></bookViews>' +
        '<sheets><sheet name="' + esc("دوام " + monthName) + '" sheetId="1" r:id="rId1"/></sheets>' +
        '</workbook>',
      "xl/_rels/workbook.xml.rels":
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
        '</Relationships>',
      "xl/styles.xml": stylesXml,
      "xl/worksheets/sheet1.xml": sheetXml,
    };
    const enc = new TextEncoder();
    const out = {};
    for (const k of Object.keys(files)) out[k] = enc.encode(files[k]);
    return zipStore(out);
  }

  /* ==================== CSV وتقرير الجودة ==================== */

  function buildCsv(att, year, month) {
    const lines = [["الرقم الوظيفي", "اسم الجهاز", "الاسم في القالب", "التاريخ", "اليوم",
      "ساعة الدخول", "ساعة الخروج", "عدد البصمات", "ملاحظة النظام", "ملاحظة محفوظة", "الملاحظة النهائية", "تأخير"]];
    const sorted = [...att.rows.values()].sort((a, b) => a.d.localeCompare(b.d) || a.empId - b.empId);
    for (const row of sorted) {
      lines.push([row.empId, row.deviceName, row.templateName, row.d, DAY_AR[pyWeekday(row.y, row.m, row.dnum)] || "",
        row.inMins != null ? fmtTime(row.inMins) : "", row.outMins != null ? fmtTime(row.outMins) : "",
        row.punches, row.sysNote, row.savedNote, row.note, row.late ? "نعم" : ""]);
    }
    return lines.map(r => r.map(v => {
      v = String(v == null ? "" : v);
      return /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v;
    }).join(",")).join("\r\n");
  }

  function buildQuality(att, parseRes, settings, year, month) {
    const lines = [];
    const workdays = new Set((settings.workdays || [6, 0, 1, 2, 3]).map(Number));
    const employees = settings.employees || [];

    if (Object.keys(att.unmapped).length) {
      lines.push("== أرقام وظيفية في attlog غير مربوطة بأسماء القالب (يجب ربطها من شاشة الربط) ==");
      for (const id of Object.keys(att.unmapped).sort((a, b) => a - b))
        lines.push("  - الرقم " + id + " (اسم الجهاز: " + (att.unmapped[id] || "غير معروف") + ")");
      lines.push("");
    }
    const missing = [...att.rows.values()].filter(r => r.sysNote && r.sysNote.indexOf("خروج وعودة") < 0);
    if (missing.length) {
      lines.push("== أيام بنقص بصمة ==");
      for (const row of missing.sort((a, b) => a.d.localeCompare(b.d) || a.empId - b.empId)) {
        const who = row.templateName || ("#" + row.empId);
        const when = (row.inMins != null ? "دخول " : "") + (row.outMins != null ? "خروج " : "");
        lines.push("  - " + row.d + ": " + who + " ← " + row.sysNote + " (" + when + ")");
      }
      lines.push("");
    }
    const mids = [...att.rows.values()].filter(r => r.sysNote.indexOf("خروج وعودة") >= 0);
    if (mids.length) {
      lines.push("== أيام فيها خروج وعودة (أكثر من بصمتين) ==");
      for (const row of mids.sort((a, b) => a.d.localeCompare(b.d) || a.empId - b.empId))
        lines.push("  - " + row.d + ": " + (row.templateName || "#" + row.empId) + " ← " + row.sysNote);
      lines.push("");
    }
    const weekend = [...att.rows.values()].filter(r => !workdays.has(pyWeekday(r.y, r.m, r.dnum)));
    if (weekend.length) {
      lines.push("== بصمات في أيام غير أيام دوام (تظهر كتحذير ولا تُدرج في القالب) ==");
      for (const row of weekend.sort((a, b) => a.d.localeCompare(b.d)))
        lines.push("  - " + row.d + " (" + (DAY_AR[pyWeekday(row.y, row.m, row.dnum)] || "") + "): " + (row.templateName || "#" + row.empId));
      lines.push("");
    }
    const seenIds = new Set([...att.rows.values()].map(r => r.empId));
    const silent = employees.filter(e => e.id && !e.no_punch && !seenIds.has(e.id));
    if (silent.length) {
      lines.push("== موظفون في القالب بلا أي بصمة طوال الشهر (يبقى صفهم فارغاً) ==");
      for (const e of silent) lines.push("  - " + (e.template_name || "?") + " (الرقم " + e.id + ")");
      lines.push("");
    }
    if (parseRes.corrupt.length) {
      lines.push("== أسطر تالفة استُبعدت (" + parseRes.corrupt.length + ") ==");
      for (const c of parseRes.corrupt.slice(0, 20))
        lines.push("  - سطر " + c.no + ": " + c.reason + " ← " + c.text);
      if (parseRes.corrupt.length > 20)
        lines.push("  ... و" + (parseRes.corrupt.length - 20) + " أسطر أخرى");
      lines.push("");
    }
    if (parseRes.duplicates) {
      lines.push("== بصمات مكررة حُذفت (فارق أقل من دقيقتين): " + parseRes.duplicates + " ==");
      lines.push("");
    }
    if (!lines.length) lines.push("لا توجد ملاحظات جودة — كل شيء سليم.");
    return lines;
  }

  /* ==================== الواجهة العامة ==================== */

  function analyzeFromFiles(files, year, month) {
    const info = firstSheetInfo(files);
    const sst = parseSST(files["xl/sharedStrings.xml"]);
    const dateStyles = parseDateStyles(files["xl/styles.xml"]);
    const model = parseSheet(files[info.path], sst, dateStyles);
    model.sheetName = info.name;
    const res = collectAnchors(model, year, month);
    const names = [];
    const seen = new Set();
    for (const nc of findNameCells(model)) {
      if (!seen.has(nc.nv)) { seen.add(nc.nv); names.push(nc.orig); }
    }
    const out = { ok: true, sheet: info.name, names: names.slice(0, 80), names_count: names.length,
                  month_days_found: Object.keys(res.anchors).map(ds => +ds.split("-")[2]).sort((a, b) => a - b),
                  messages: res.messages, mergedCount: model.merges.length };
    if (Object.keys(res.anchors).length) out.mode = "✅ " + res.mode + " — ستُعبّأ نسخة من قالبكم نفسه بكل تنسيقاته";
    else out.warning = "لم أتعرف على أعمدة الأيام في القالب — سيُستخدم المولّد المدمج المطابق لبنية القالب. " + res.messages.join(" ");
    return out;
  }

  async function analyzeTemplate(bytes, year, month) {
    const files = await unzip(bytes);
    return analyzeFromFiles(files, year, month);
  }

  return {
    normalizeAr, toInt, fmtTime, MONTH_AR, DAY_AR, monthWorkdays, pyWeekday,
    parseAttlog, processMonth, unzip, parseSST, parseDateStyles, parseSheet,
    collectAnchors, matchEmployeeRows, analyzeTemplate, analyzeFromFiles,
    cellText, escapeXml, zipStore, crc32, colToRef, splitRef, parseRange,
    serialToDate, datePartsToSerial, dstr, monthDaysCount,
    fillTemplateBytes, buildFallbackBytes, buildCsv, buildQuality,
  };
})();
