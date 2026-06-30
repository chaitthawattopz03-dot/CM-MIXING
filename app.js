/* ============================================================
   IW28 / IW38 Maintenance Dashboard — standalone build
   HTML + CSS + vanilla JS, plus Firebase Firestore for shared
   (cross-device, cross-browser) "latest uploaded file" sync via
   the window.__fb bridge set up by firebase-init.js.
   ============================================================ */

(function () {
  "use strict";

  const STORAGE_DATA_KEY = "iw_dash_static_v1";
  const STORAGE_SETTINGS_KEY = "iw_dash_static_settings_v1";

  const DEFAULT_SETTINGS = { overdueDays: 30, monthsBack: 18, accent: "#2b5cc4" };

  const state = {
    loaded: false,
    parsing: false,
    error: "",
    iw28: null,
    iw38: null,
    name28: "",
    name38: "",
    pgFilter: "ALL",
    statusFilter: "ALL",
    scope: "backlog",
    settingsOpen: false,
    settings: Object.assign({}, DEFAULT_SETTINGS),
    cloudStatus: "connecting", // "connecting" | "synced" | "offline" | "saving"
    updatedAt: null,
  };

  // ---------------------------------------------------------------
  // xlsx parsing (zero-dependency: manual ZIP + sharedStrings + sheet XML)
  // ---------------------------------------------------------------
  function unzip(buf) {
    const dv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
    let eocd = -1;
    for (let i = buf.length - 22; i >= 0; i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error("ไม่ใช่ไฟล์ xlsx ที่ถูกต้อง");
    const cdOffset = dv.getUint32(eocd + 16, true), cdCount = dv.getUint16(eocd + 10, true);
    const files = {};
    let p = cdOffset;
    for (let n = 0; n < cdCount; n++) {
      if (dv.getUint32(p, true) !== 0x02014b50) break;
      const method = dv.getUint16(p + 10, true), compSize = dv.getUint32(p + 20, true);
      const nameLen = dv.getUint16(p + 28, true), extraLen = dv.getUint16(p + 30, true), commentLen = dv.getUint16(p + 32, true);
      const lho = dv.getUint32(p + 42, true);
      const name = new TextDecoder().decode(buf.subarray(p + 46, p + 46 + nameLen));
      const lnameLen = dv.getUint16(lho + 26, true), lextraLen = dv.getUint16(lho + 28, true);
      const ds = lho + 30 + lnameLen + lextraLen;
      files[name] = { method, comp: buf.subarray(ds, ds + compSize) };
      p += 46 + nameLen + extraLen + commentLen;
    }
    return files;
  }
  async function inflateEntry(e) {
    if (e.method === 0) return e.comp;
    const ds = new DecompressionStream("deflate-raw");
    return new Uint8Array(await new Response(new Blob([e.comp]).stream().pipeThrough(ds)).arrayBuffer());
  }
  async function readEntry(files, name) {
    return new TextDecoder().decode(await inflateEntry(files[name]));
  }
  function unesc(s) {
    return s.replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#10;/g, "\n");
  }
  function parseSharedStrings(ss) {
    const out = [];
    const re = /<si>(.*?)<\/si>/gs;
    let m;
    while ((m = re.exec(ss))) {
      const t = [...m[1].matchAll(/<t[^>]*>(.*?)<\/t>/gs)].map((x) => x[1]).join("");
      out.push(unesc(t));
    }
    return out;
  }
  function colToNum(c) {
    let n = 0;
    for (const x of c) n = n * 26 + (x.charCodeAt(0) - 64);
    return n;
  }
  function parseSheet(sheet, strings) {
    const rows = [];
    const rowRe = /<row[^>]*r="(\d+)"[^>]*>(.*?)<\/row>/gs;
    let rm;
    while ((rm = rowRe.exec(sheet))) {
      const cells = {};
      const cellRe = /<c r="([A-Z]+)\d+"(?:[^>]*t="([^"]*)")?[^>]*>(?:<v>(.*?)<\/v>|<is><t[^>]*>(.*?)<\/t><\/is>)?<\/c>/gs;
      let cm;
      while ((cm = cellRe.exec(rm[2]))) {
        const col = colToNum(cm[1]);
        const t = cm[2];
        let val = cm[3] !== undefined ? cm[3] : cm[4];
        if (val === undefined) continue;
        if (t === "s") val = strings[parseInt(val, 10)];
        else val = unesc(val);
        cells[col] = val;
      }
      rows.push(cells);
    }
    return rows;
  }
  function colOf(header, name) {
    for (const k in header) {
      if ((header[k] || "").toString().trim().toLowerCase() === name.toLowerCase()) return +k;
    }
    for (const k in header) {
      if ((header[k] || "").toString().trim().toLowerCase().includes(name.toLowerCase())) return +k;
    }
    return null;
  }
  function numVal(v) {
    const n = parseFloat(v);
    return isFinite(n) ? n : 0;
  }
  async function parseWorkbook(bytes, name) {
    const files = unzip(bytes);
    const strings = files["xl/sharedStrings.xml"] ? parseSharedStrings(await readEntry(files, "xl/sharedStrings.xml")) : [];
    const sheetKey =
      Object.keys(files).find((k) => /worksheets\/sheet1\.xml$/.test(k)) ||
      Object.keys(files).find((k) => /worksheets\/.*\.xml$/.test(k));
    if (!sheetKey) throw new Error("ไม่พบชีตข้อมูลในไฟล์");
    const rows = parseSheet(await readEntry(files, sheetKey), strings);
    const header = rows[0] || {};
    const hvals = Object.values(header).map((v) => (v || "").toString());
    const has = (s) => hvals.some((h) => h.toLowerCase().includes(s.toLowerCase()));
    const co = (n) => colOf(header, n);

    let type = "unknown";
    if (has("TotalPlnndCosts") || has("act.costs")) type = "iw38";
    else if (has("Notifictn") || has("Notif. Date") || has("Reported By")) type = "iw28";

    const records = [];
    if (type === "iw38") {
      const c = {
        pg: co("Planner Group"), otype: co("Order Type"), order: co("Order"), desc: co("Description"),
        plan: co("TotalPlnndCosts"), act: co("act.costs"), start: co("start date"), fin: co("fin. date"), us: co("User status"),
      };
      for (let i = 1; i < rows.length; i++) {
        const r = rows[i];
        if (!r || Object.keys(r).length === 0) continue;
        const ord = (r[c.order] || "").toString().trim();
        if (!ord) continue;
        records.push({
          pg: (r[c.pg] || "").toString().trim() || "(blank)",
          otype: (r[c.otype] || "").toString(),
          order: ord,
          desc: (r[c.desc] || "").toString(),
          plan: numVal(r[c.plan]),
          act: numVal(r[c.act]),
          start: numVal(r[c.start]),
          fin: numVal(r[c.fin]),
          us: (r[c.us] || "").toString().trim(),
        });
      }
    } else if (type === "iw28") {
      const c = {
        notif: co("Notification"), order: co("Order"), date: co("Notif. Date"), desc: co("Description"),
        us: co("User status"), reporter: co("Reported By"), pg: co("Planner Group"), prio: co("Priority"),
      };
      for (let i = 1; i < rows.length; i++) {
        const r = rows[i];
        if (!r || Object.keys(r).length === 0) continue;
        records.push({
          notif: (r[c.notif] || "").toString(),
          order: (r[c.order] || "").toString().trim(),
          date: numVal(r[c.date]),
          desc: (r[c.desc] || "").toString(),
          us: (r[c.us] || "").toString().trim(),
          reporter: (r[c.reporter] || "").toString(),
          pg: (r[c.pg] || "").toString().trim() || "(blank)",
          prio: (r[c.prio] || "").toString(),
        });
      }
    }
    return { type, name, records };
  }

  // ---------------------------------------------------------------
  // formatting helpers
  // ---------------------------------------------------------------
  function fmtInt(n) { return Math.round(n || 0).toLocaleString("en-US"); }
  function fmtM(n) { const v = (n || 0) / 1e6; return "฿" + v.toFixed(2) + "M"; }
  function serialMs(s) { return Date.UTC(1899, 11, 30) + s * 86400000; }
  function fmtDate(s) {
    if (!s) return "—";
    const d = new Date(serialMs(s));
    return String(d.getUTCDate()).padStart(2, "0") + "/" + String(d.getUTCMonth() + 1).padStart(2, "0") + "/" + String(d.getUTCFullYear()).slice(2);
  }
  function ym(ms) { const d = new Date(ms); return d.getUTCFullYear() + "-" + String(d.getUTCMonth() + 1).padStart(2, "0"); }
  function fmtDateTime(ms) {
    const d = new Date(ms);
    return String(d.getDate()).padStart(2, "0") + "/" + String(d.getMonth() + 1).padStart(2, "0") + "/" + d.getFullYear() +
      " " + String(d.getHours()).padStart(2, "0") + ":" + String(d.getMinutes()).padStart(2, "0");
  }
  function renderCloudBadge() {
    const map = {
      connecting: { text: "กำลังเชื่อมต่อคลาวด์…", cls: "badge-wait" },
      saving: { text: "กำลังบันทึกขึ้นคลาวด์…", cls: "badge-wait" },
      synced: { text: "ซิงค์กับคลาวด์แล้ว", cls: "badge-ok" },
      offline: { text: "ออฟไลน์ — เห็นแค่เครื่องนี้", cls: "badge-warn" },
    };
    const m = map[state.cloudStatus] || map.connecting;
    return `<span class="cloud-badge ${m.cls}">${esc(m.text)}</span>`;
  }
  const MONTH_TH = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];
  function mLabel(k) { const [y, m] = k.split("-"); return MONTH_TH[+m - 1] + " " + String(y).slice(2); }
  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  }

  const STATUS_META = {
    NTAP: { label: "NTAP · ค้าง", color: "#c1443a" },
    NTCN: { label: "NTCN · ค้าง", color: "#d4682f" },
    NTAC: { label: "NTAC · ค้าง", color: "#c98a1a" },
    NTMC: { label: "NTMC · ค้าง", color: "#9bbf3a" },
    "(none)": { label: "ไม่มีสถานะ", color: "#b9b6ad" },
    NOCO: { label: "NOCO · เสร็จ", color: "#2e8b6f" },
  };
  function meta(g) { return STATUS_META[g] || { label: g || "—", color: "#5b6b8c" }; }

  // ---------------------------------------------------------------
  // persistence
  // ---------------------------------------------------------------
  // ---------------------------------------------------------------
  // persistence — Firestore is the shared source of truth; localStorage
  // is just an instant-paint cache for this one browser (and an offline
  // fallback if the cloud is briefly unreachable).
  // ---------------------------------------------------------------
  let cloudIW28 = null; // raw {records,name,updatedAt} from Firestore doc "iw28"
  let cloudIW38 = null; // raw {records,name,updatedAt} from Firestore doc "iw38"
  let unsubIW28 = null, unsubIW38 = null;

  function saveCache() {
    try {
      localStorage.setItem(STORAGE_DATA_KEY, JSON.stringify({
        iw28: state.iw28, iw38: state.iw38, name28: state.name28, name38: state.name38, updatedAt: state.updatedAt,
      }));
    } catch (e) { /* ignore quota errors */ }
  }
  function loadCache() {
    try {
      const raw = localStorage.getItem(STORAGE_DATA_KEY);
      if (!raw) return false;
      const d = JSON.parse(raw);
      if (d && d.iw28 && d.iw38) {
        state.loaded = true; state.iw28 = d.iw28; state.iw38 = d.iw38;
        state.name28 = d.name28 || ""; state.name38 = d.name38 || ""; state.updatedAt = d.updatedAt || null;
        return true;
      }
    } catch (e) { /* ignore */ }
    return false;
  }
  function saveSettings() {
    try { localStorage.setItem(STORAGE_SETTINGS_KEY, JSON.stringify(state.settings)); } catch (e) {}
  }
  function loadSettings() {
    try {
      const raw = localStorage.getItem(STORAGE_SETTINGS_KEY);
      if (raw) state.settings = Object.assign({}, DEFAULT_SETTINGS, JSON.parse(raw));
    } catch (e) {}
  }

  /** Once both Firestore docs have reported in at least once, adopt
   *  whichever pair is most current as the displayed dataset. Fires on
   *  every subsequent update too, so all open tabs stay in sync live. */
  function maybeApplyCloud() {
    if (!cloudIW28 || !cloudIW38) return;
    state.loaded = true; state.parsing = false;
    state.iw28 = cloudIW28.records; state.iw38 = cloudIW38.records;
    state.name28 = cloudIW28.name || ""; state.name38 = cloudIW38.name || "";
    state.updatedAt = Math.max(cloudIW28.updatedAt || 0, cloudIW38.updatedAt || 0);
    state.cloudStatus = "synced";
    saveCache();
    render();
  }

  function startCloudSync() {
    if (!window.__fb) return false;
    unsubIW28 = window.__fb.subscribe("iw28", (data, err) => {
      if (err) { state.cloudStatus = "offline"; render(); return; }
      cloudIW28 = data; maybeApplyCloud();
    });
    unsubIW38 = window.__fb.subscribe("iw38", (data, err) => {
      if (err) { state.cloudStatus = "offline"; render(); return; }
      cloudIW38 = data; maybeApplyCloud();
    });
    return true;
  }

  async function pushToCloud(f28, f38) {
    if (!window.__fb) { state.cloudStatus = "offline"; return; }
    state.cloudStatus = "saving"; render();
    const updatedAt = Date.now();
    try {
      await Promise.all([
        window.__fb.saveDoc("iw28", { records: f28.records, name: f28.name, updatedAt }),
        window.__fb.saveDoc("iw38", { records: f38.records, name: f38.name, updatedAt }),
      ]);
      // onSnapshot will fire locally too and set cloudStatus="synced" via maybeApplyCloud
    } catch (e) {
      console.error("[firebase] save failed", e);
      state.cloudStatus = "offline";
      state.error = "บันทึกขึ้นคลาวด์ไม่สำเร็จ (" + (e && e.message ? e.message : String(e)) + ") — เห็นข้อมูลนี้แค่เครื่องนี้จนกว่าจะเชื่อมต่อใหม่ได้";
      render();
    }
  }

  // ---------------------------------------------------------------
  // file ingestion
  // ---------------------------------------------------------------
  async function tryAutoload() {
    try {
      const [r1, r2] = await Promise.all([
        fetch("./uploads/IW28.xlsx").then((r) => (r.ok ? r.arrayBuffer() : null)).catch(() => null),
        fetch("./uploads/IW38.xlsx").then((r) => (r.ok ? r.arrayBuffer() : null)).catch(() => null),
      ]);
      if (!r1 || !r2) return false;
      const f1 = await parseWorkbook(new Uint8Array(r1), "IW28.xlsx");
      const f2 = await parseWorkbook(new Uint8Array(r2), "IW38.xlsx");
      const map = {}; for (const f of [f1, f2]) if (f && f.type !== "unknown") map[f.type] = f;
      if (map.iw28 && map.iw38) { applyFiles(map.iw28, map.iw38); return true; }
    } catch (e) { /* silent — fall back to upload screen */ }
    return false;
  }

  function applyFiles(f28, f38) {
    state.loaded = true; state.parsing = false; state.error = "";
    state.iw28 = f28.records; state.iw38 = f38.records;
    state.name28 = f28.name; state.name38 = f38.name;
    saveCache();
    render();
    pushToCloud(f28, f38);
  }

  async function ingestFiles(fileList) {
    state.parsing = true; state.error = ""; render();
    try {
      let f28 = state.iw28 ? { type: "iw28", records: state.iw28, name: state.name28 } : null;
      let f38 = state.iw38 ? { type: "iw38", records: state.iw38, name: state.name38 } : null;
      for (const file of fileList) {
        const ab = await file.arrayBuffer();
        const pf = await parseWorkbook(new Uint8Array(ab), file.name);
        if (pf.type === "iw28") f28 = pf;
        else if (pf.type === "iw38") f38 = pf;
      }
      if (!f28 || !f38) {
        state.parsing = false;
        state.iw28 = f28 ? f28.records : state.iw28;
        state.iw38 = f38 ? f38.records : state.iw38;
        state.name28 = f28 ? f28.name : state.name28;
        state.name38 = f38 ? f38.name : state.name38;
        state.error = !f28 && !f38
          ? "ไม่พบไฟล์ IW28/IW38 ที่อ่านได้ — ตรวจสอบว่าไฟล์มีหัวตารางที่ถูกต้อง"
          : !f28 ? "ได้ไฟล์ IW38 แล้ว — กรุณาวางไฟล์ IW28 เพิ่ม" : "ได้ไฟล์ IW28 แล้ว — กรุณาวางไฟล์ IW38 เพิ่ม";
        render();
        return;
      }
      applyFiles(f28, f38);
    } catch (err) {
      state.parsing = false;
      state.error = "อ่านไฟล์ไม่สำเร็จ: " + (err && err.message ? err.message : String(err));
      render();
    }
  }

  function resetAll() {
    try { localStorage.removeItem(STORAGE_DATA_KEY); } catch (e) {}
    state.loaded = false; state.iw28 = null; state.iw38 = null; state.name28 = ""; state.name38 = "";
    state.pgFilter = "ALL"; state.statusFilter = "ALL"; state.scope = "backlog"; state.error = "";
    render();
  }

  // ---------------------------------------------------------------
  // compute view-model from current state
  // ---------------------------------------------------------------
  function compute() {
    const overdue = state.settings.overdueDays;
    const monthsBack = state.settings.monthsBack;
    const accent = state.settings.accent;
    const o28 = state.iw28 || [], o38 = state.iw38 || [];
    const now = Date.now();

    const obo = {};
    for (const o of o38) {
      if (!o.order) continue;
      if (!obo[o.order]) obo[o.order] = { plan: 0, act: 0, us: o.us, fin: o.fin };
      obo[o.order].plan += o.plan; obo[o.order].act += o.act;
      if (o.fin) obo[o.order].fin = o.fin;
    }

    const recs = o28.map((n) => {
      const ord = n.order; const o = ord ? obo[ord] : null;
      const grp = n.us || "(none)";
      const openS = n.date || 0;
      const openMs = openS ? serialMs(openS) : null;
      const finMs = o && o.fin ? serialMs(o.fin) : null;
      const isBack = grp !== "NOCO";
      let age = null;
      if (isBack && openMs && openMs <= now) age = Math.floor((now - openMs) / 86400000);
      else if (!isBack && openMs && finMs && finMs >= openMs) age = Math.floor((finMs - openMs) / 86400000);
      return { notif: n.notif, pg: n.pg, order: ord, matched: !!o, desc: n.desc || "", plan: o ? o.plan : 0, act: o ? o.act : 0, grp, us28: n.us || "", openS, openMs, finMs, isBack, age };
    });

    const f = state.pgFilter, sf = state.statusFilter;
    let view = f === "ALL" ? recs.slice() : recs.filter((r) => r.pg === f);
    if (sf !== "ALL") view = view.filter((r) => (sf === "(none)" ? !r.us28 : r.us28 === sf));

    const usCount = {};
    for (const r of recs) { const k = r.us28 || "(none)"; usCount[k] = (usCount[k] || 0) + 1; }
    const statusOptions = [{ value: "ALL", label: "ทุก Status งานแจ้ง" }].concat(
      Object.keys(usCount).sort((a, b) => usCount[b] - usCount[a]).map((k) => ({ value: k, label: (k === "(none)" ? "ไม่มีใบแจ้ง" : k) + " (" + usCount[k] + ")" }))
    );

    const pgMap = {};
    for (const r of recs) {
      const k = r.pg;
      (pgMap[k] = pgMap[k] || { pg: k, total: 0, backlog: 0, plan: 0, act: 0 });
      pgMap[k].total++; if (r.isBack) pgMap[k].backlog++; pgMap[k].plan += r.plan; pgMap[k].act += r.act;
    }
    const pgArr = Object.values(pgMap).sort((a, b) => b.backlog - a.backlog);
    const maxBack = Math.max(1, ...pgArr.map((p) => p.backlog));
    const pgBars = pgArr.map((p) => {
      const active = f === p.pg;
      return { pg: p.pg, backStr: fmtInt(p.backlog), totalStr: fmtInt(p.total), wpct: Math.max(2, (p.backlog / maxBack) * 100), barColor: accent, active };
    });

    const totalOrders = view.length;
    const backlog = view.filter((r) => r.isBack);
    const noco = view.filter((r) => r.grp === "NOCO");
    const seenO = {}; let totalPlan = 0, totalAct = 0, matchedN = 0;
    for (const r of view) { if (r.matched) matchedN++; if (r.order && !seenO[r.order]) { seenO[r.order] = 1; totalPlan += r.plan; totalAct += r.act; } }
    const seenB = {}; let backPlan = 0;
    for (const r of backlog) { if (r.order && !seenB[r.order]) { seenB[r.order] = 1; backPlan += r.plan; } }
    const ages = backlog.map((r) => r.age).filter((a) => a != null);
    const avgAge = ages.length ? Math.round(ages.reduce((s, a) => s + a, 0) / ages.length) : 0;
    const maxAge = ages.length ? Math.max(...ages) : 0;
    const compRate = totalOrders ? Math.round((noco.length / totalOrders) * 100) : 0;
    const variance = totalAct - totalPlan;
    const ink = "#1d1d18";
    const kpis = [
      { label: "ใบแจ้งทั้งหมด (IW28)", value: fmtInt(totalOrders), sub: "จับคู่ Order ได้ " + fmtInt(matchedN) + (totalOrders ? " (" + Math.round((matchedN / totalOrders) * 100) + "%)" : ""), stripe: "#dcdcd6", valColor: ink },
      { label: "งานค้าง (ยังไม่ NOCO)", value: fmtInt(backlog.length), sub: totalOrders ? Math.round((backlog.length / totalOrders) * 100) + "% ของใบแจ้งทั้งหมด" : "—", stripe: accent, valColor: accent },
      { label: "เสร็จแล้ว (NOCO)", value: fmtInt(noco.length), sub: compRate + "% completion rate", stripe: "#2e8b6f", valColor: "#2e8b6f" },
      { label: "อายุงานค้างเฉลี่ย", value: fmtInt(avgAge), sub: "วัน · ค้างนานสุด " + fmtInt(maxAge) + " วัน", stripe: "#c98a1a", valColor: ink },
      { label: "มูลค่า Plan รวม", value: fmtM(totalPlan), sub: fmtInt(totalPlan) + " ฿", stripe: "#dcdcd6", valColor: ink },
      { label: "มูลค่า Actual รวม", value: fmtM(totalAct), sub: fmtInt(totalAct) + " ฿", stripe: "#dcdcd6", valColor: ink },
      { label: "ส่วนต่าง Act − Plan", value: (variance >= 0 ? "+" : "−") + fmtM(Math.abs(variance)), sub: totalPlan ? Math.round((variance / totalPlan) * 100) + "% เทียบแผน" : "—", stripe: variance > 0 ? "#c1443a" : "#2e8b6f", valColor: variance > 0 ? "#c1443a" : "#2e8b6f" },
      { label: "มูลค่างานค้าง (Plan)", value: fmtM(backPlan), sub: fmtInt(backPlan) + " ฿ รอดำเนินการ", stripe: accent, valColor: accent },
    ];

    const order = ["NTAP", "NTCN", "NTAC", "NTMC", "(none)", "NOCO"];
    const stMap = {}; for (const r of view) stMap[r.grp] = (stMap[r.grp] || 0) + 1;
    const stTotal = view.length || 1; let acc = 0; const segs = [];
    const seen = order.reduce((m, k) => ((m[k] = 1), m), {});
    for (const k in stMap) if (!seen[k]) order.push(k);
    const statusRows = order.filter((k) => stMap[k]).map((k) => {
      const cnt = stMap[k], pct = (cnt / stTotal) * 100;
      segs.push({ from: acc, to: acc + pct, color: meta(k).color }); acc += pct;
      return { label: meta(k).label, count: fmtInt(cnt), pct: Math.round(pct), color: meta(k).color };
    });
    const donut = segs.length ? "conic-gradient(" + segs.map((s) => s.color + " " + s.from + "% " + s.to + "%").join(",") + ")" : "#eee";

    const buckets = [
      { label: "0–30 วัน", max: 30, color: "#2e8b6f" }, { label: "31–60", max: 60, color: "#9bbf3a" },
      { label: "61–90", max: 90, color: "#c98a1a" }, { label: "91–180", max: 180, color: "#d4682f" }, { label: "180+ วัน", max: 1e9, color: "#c1443a" },
    ];
    const bc = buckets.map(() => 0);
    for (const r of backlog) { if (r.age == null) continue; for (let i = 0; i < buckets.length; i++) { if (r.age <= buckets[i].max) { bc[i]++; break; } } }
    const maxBkt = Math.max(1, ...bc);
    const agingRows = buckets.map((b, i) => ({ label: b.label, count: fmtInt(bc[i]), hpct: Math.max(3, (bc[i] / maxBkt) * 100), color: b.color }));

    const mOpen = {}, mClose = {};
    for (const r of view) {
      if (r.openMs) { const k = ym(r.openMs); mOpen[k] = (mOpen[k] || 0) + 1; }
      if (r.grp === "NOCO" && r.finMs) { const k = ym(r.finMs); mClose[k] = (mClose[k] || 0) + 1; }
    }
    let keys = Object.keys(Object.assign({}, mOpen, mClose)).filter((k) => k >= "2024-01" && k <= "2027-12").sort();
    keys = keys.slice(-monthsBack);
    const maxM = Math.max(1, ...keys.map((k) => Math.max(mOpen[k] || 0, mClose[k] || 0)));
    const monthRows = keys.map((k) => ({ label: mLabel(k), opened: mOpen[k] || 0, closed: mClose[k] || 0, openedH: Math.max(2, ((mOpen[k] || 0) / maxM) * 100), closedH: Math.max(2, ((mClose[k] || 0) / maxM) * 100) }));

    const yMap = {};
    for (const r of view) {
      if (r.openMs) { const y = new Date(r.openMs).getUTCFullYear(); (yMap[y] = yMap[y] || { year: y, opened: 0, closed: 0, backlog: 0, plan: 0, act: 0 }); yMap[y].opened++; if (r.isBack) yMap[y].backlog++; yMap[y].plan += r.plan; yMap[y].act += r.act; }
      if (r.grp === "NOCO" && r.finMs) { const y = new Date(r.finMs).getUTCFullYear(); (yMap[y] = yMap[y] || { year: y, opened: 0, closed: 0, backlog: 0, plan: 0, act: 0 }); yMap[y].closed++; }
    }
    const yearRows = Object.values(yMap).sort((a, b) => a.year - b.year).map((y) => ({ year: y.year, opened: fmtInt(y.opened), closed: fmtInt(y.closed), backlog: fmtInt(y.backlog), plan: fmtM(y.plan), act: fmtM(y.act) }));

    const maxPA = Math.max(1, ...pgArr.map((p) => Math.max(p.plan, p.act)));
    const paRows = pgArr.slice().sort((a, b) => b.plan - a.plan).map((p) => ({ pg: p.pg, planH: Math.max(2, (p.plan / maxPA) * 100), actH: Math.max(2, (p.act / maxPA) * 100), planStr: fmtM(p.plan), actStr: fmtM(p.act) }));

    const trowsAll = state.scope === "backlog" ? backlog.slice() : view.slice();
    trowsAll.sort((a, b) => (b.age || 0) - (a.age || 0));
    const cap = 120; const tableTotal = trowsAll.length;
    const tableRows = trowsAll.slice(0, cap).map((r) => ({
      order: r.order || "—", desc: (r.desc || "").slice(0, 72), pg: r.pg, statusLabel: meta(r.grp).label, statusColor: meta(r.grp).color,
      openStr: fmtDate(r.openS), age: r.age == null ? "—" : fmtInt(r.age), ageColor: r.age != null && r.age > overdue ? "#c1443a" : "#33332d",
      planStr: r.matched ? fmtInt(r.plan) : "—", actStr: r.matched ? fmtInt(r.act) : "—",
    }));

    const asOf = new Date(now);
    const asOfStr = String(asOf.getDate()).padStart(2, "0") + "/" + String(asOf.getMonth() + 1).padStart(2, "0") + "/" + asOf.getFullYear();

    return {
      kpis, pgBars, statusRows, donut, donutCenter: compRate + "%", agingRows, monthRows, yearRows, paRows, accent,
      tableRows, tableTotal, tableShown: Math.min(cap, tableTotal),
      asOfStr, c28: fmtInt(o28.length), c38: fmtInt(o38.length), statusOptions,
    };
  }

  // ---------------------------------------------------------------
  // rendering
  // ---------------------------------------------------------------
  const app = document.getElementById("app");

  function render() {
    document.documentElement.style.setProperty("--accent", state.settings.accent);
    app.innerHTML = state.loaded ? renderDashboard() : renderUpload();
    wireEvents();
  }

  function renderUpload() {
    return `
      <div class="upload-wrap" id="dropzone">
        <div class="upload-head">
          <div class="upload-title">แดชบอร์ดงานซ่อมบำรุง · IW28 / IW38</div>
          <div class="upload-sub">ลากวางไฟล์ Excel ทั้งสอง แล้วระบบจะเชื่อมข้อมูลด้วยเลข Order สรุปงาน งานค้าง มูลค่า Plan/Actual และกราฟสำหรับ HOD</div>
        </div>
        <div class="upload-box" id="dropbox">
          <div class="upload-types">
            <div class="type-card">
              <div class="type-tag tag-blue">IW28</div>
              <div class="type-desc">ใบแจ้งซ่อม (Notifications)</div>
            </div>
            <div class="type-card">
              <div class="type-tag tag-green">IW38</div>
              <div class="type-desc">ใบสั่งงาน + ค่าใช้จ่าย (Orders)</div>
            </div>
          </div>
          <div class="upload-cta">วางไฟล์ <b>.xlsx</b> ทั้งสองที่นี่</div>
          <div class="upload-hint">ระบบแยกประเภทไฟล์อัตโนมัติจากหัวตาราง</div>
          <label class="pick-btn">เลือกไฟล์จากเครื่อง
            <input type="file" accept=".xlsx" multiple id="filePick" style="display:none;">
          </label>
          ${state.parsing ? `<div class="parsing-row"><span class="spinner"></span>กำลังอ่านและประมวลผลไฟล์…</div>` : ""}
          ${state.error ? `<div class="error-pill">${esc(state.error)}</div>` : ""}
        </div>
      </div>`;
  }

  function renderDashboard() {
    const m = compute();
    return `
    <div class="dash-wrap">
      <div class="dash-header">
        <div>
          <div class="dash-title">แดชบอร์ดงานซ่อมบำรุง · IW28 / IW38</div>
          <div class="dash-meta">ข้อมูล ณ ${m.asOfStr} &nbsp;·&nbsp; IW28 ${m.c28} ใบแจ้ง &nbsp;·&nbsp; IW38 ${m.c38} ใบสั่งงาน${state.updatedAt ? " &nbsp;·&nbsp; อัปเดตล่าสุด " + fmtDateTime(state.updatedAt) : ""}</div>
        </div>
        <div class="iw-noprint header-actions">
          ${renderCloudBadge()}
          <button class="btn-ghost" id="btnSettings">ตั้งค่า</button>
          <button class="btn-ghost" id="btnPrint">พิมพ์ / PDF</button>
          <button class="btn-dark" id="btnReset">เปลี่ยนไฟล์</button>
        </div>
      </div>

      ${state.settingsOpen ? renderSettingsPanel() : ""}

      <div class="iw-noprint filter-row">
        <span class="filter-label">กรองตาม PG:</span>
        <button class="chip ${state.pgFilter === "ALL" ? "chip-active" : ""}" data-pg="ALL">ทุกกลุ่ม</button>
        ${m.pgBars.map((b) => `<button class="chip ${b.active ? "chip-active" : ""}" data-pg="${esc(b.pg)}">${esc(b.pg)}</button>`).join("")}
        <span class="filter-label" style="margin-left:10px;">Status งานแจ้ง:</span>
        <select class="select-status" id="statusSelect">
          ${m.statusOptions.map((o) => `<option value="${esc(o.value)}" ${o.value === state.statusFilter ? "selected" : ""}>${esc(o.label)}</option>`).join("")}
        </select>
      </div>

      <div class="kpi-grid">
        ${m.kpis.map((k) => `
          <div class="kpi-card">
            <div class="kpi-stripe" style="background:${k.stripe}"></div>
            <div class="kpi-label">${esc(k.label)}</div>
            <div class="kpi-value" style="color:${k.valColor}">${esc(k.value)}</div>
            <div class="kpi-sub">${esc(k.sub)}</div>
          </div>`).join("")}
      </div>

      <div class="row-2 row-pg-status">
        <div class="card">
          <div class="card-head"><h3>งานค้างแยกตาม PG Group</h3><span class="hint">ค้าง / ทั้งหมด · คลิกเพื่อกรอง</span></div>
          <div class="pg-bars">
            ${m.pgBars.map((b) => `
              <button class="pg-bar-row ${b.active ? "pg-bar-active" : ""}" data-pg="${esc(b.pg)}">
                <span class="pg-label" style="color:${b.active ? m.accent : "#33332d"}">${esc(b.pg)}</span>
                <span class="pg-track"><span class="pg-fill" style="width:${b.wpct}%;background:${b.barColor}"></span></span>
                <span class="pg-nums"><b>${b.backStr}</b> / ${b.totalStr}</span>
              </button>`).join("")}
          </div>
        </div>

        <div class="card">
          <h3>สัดส่วนตาม Status</h3>
          <div class="donut-row">
            <div class="donut" style="background:${m.donut}">
              <div class="donut-center"><div class="donut-num">${esc(m.donutCenter)}</div><div class="donut-cap">เสร็จ NOCO</div></div>
            </div>
            <div class="status-list">
              ${m.statusRows.map((s) => `
                <div class="status-row">
                  <span class="status-tag"><span class="dot" style="background:${s.color}"></span>${esc(s.label)}</span>
                  <span class="status-num"><b>${s.count}</b> · ${s.pct}%</span>
                </div>`).join("")}
            </div>
          </div>
        </div>
      </div>

      <div class="row-2 row-aging-pa">
        <div class="card">
          <h3 class="tight">อายุงานค้าง (Aging)</h3>
          <div class="hint mb14">นับจากวันเปิดงานถึงปัจจุบัน</div>
          <div class="bars-flex">
            ${m.agingRows.map((a) => `
              <div class="bar-col">
                <div class="bar-num">${a.count}</div>
                <div class="bar-track"><div class="bar-fill" style="height:${a.hpct}%;background:${a.color}"></div></div>
                <div class="bar-label">${esc(a.label)}</div>
              </div>`).join("")}
          </div>
        </div>

        <div class="card">
          <div class="card-head">
            <h3>มูลค่างาน Plan vs Actual แยกตาม PG</h3>
            <span class="legend"><span><i style="background:${m.accent}"></i>Plan</span><span><i style="background:#2e8b6f"></i>Actual</span></span>
          </div>
          <div class="bars-flex pa-flex">
            ${m.paRows.map((p) => `
              <div class="bar-col">
                <div class="pa-pair">
                  <div class="pa-bar" title="Plan ${esc(p.planStr)}" style="height:${p.planH}%;background:${m.accent}"></div>
                  <div class="pa-bar" title="Actual ${esc(p.actStr)}" style="height:${p.actH}%;background:#2e8b6f"></div>
                </div>
                <div class="pa-pg">${esc(p.pg)}</div>
                <div class="pa-vals">${esc(p.planStr)}<br>${esc(p.actStr)}</div>
              </div>`).join("")}
          </div>
        </div>
      </div>

      <div class="card">
        <div class="card-head">
          <h3>แนวโน้มรายเดือน · เปิดงาน (แจ้ง) vs เสร็จ (NOCO)</h3>
          <span class="legend"><span><i style="background:${m.accent}"></i>เปิดงาน</span><span><i style="background:#2e8b6f"></i>ปิดงาน</span></span>
        </div>
        <div class="month-scroll iwscroll">
          ${m.monthRows.map((mo) => `
            <div class="month-col">
              <div class="month-pair">
                <div class="month-bar" title="เปิด ${mo.opened}" style="height:${mo.openedH}%;background:${m.accent}"></div>
                <div class="month-bar" title="ปิด ${mo.closed}" style="height:${mo.closedH}%;background:#2e8b6f"></div>
              </div>
              <div class="month-label">${esc(mo.label)}</div>
            </div>`).join("")}
        </div>
      </div>

      <div class="card">
        <h3>สรุปรายปี</h3>
        <div class="year-table">
          <div class="yt-head">ปี</div><div class="yt-head right">เปิดงาน</div><div class="yt-head right">เสร็จ (NOCO)</div>
          <div class="yt-head right">เปิดแล้วยังค้าง</div><div class="yt-head right">Plan</div><div class="yt-head right">Actual</div>
          ${m.yearRows.map((y) => `
            <div class="yt-cell yt-year">${y.year}</div>
            <div class="yt-cell right">${y.opened}</div>
            <div class="yt-cell right" style="color:#2e8b6f">${y.closed}</div>
            <div class="yt-cell right" style="color:#c1443a">${y.backlog}</div>
            <div class="yt-cell right">${y.plan}</div>
            <div class="yt-cell right">${y.act}</div>`).join("")}
        </div>
      </div>

      <div class="card">
        <div class="card-head wrap-head">
          <h3>รายการงาน (เรียงตามอายุค้างนานสุด)</h3>
          <div class="table-tools">
            <span class="hint">แสดง ${m.tableShown} จาก ${m.tableTotal} รายการ</span>
            <div class="iw-noprint seg">
              <button class="seg-btn ${state.scope === "backlog" ? "seg-active" : ""}" id="scBacklog">งานค้าง</button>
              <button class="seg-btn ${state.scope === "all" ? "seg-active" : ""}" id="scAll">ทั้งหมด</button>
            </div>
          </div>
        </div>
        <div class="iwscroll table-scroll">
          <div class="data-table">
            <div class="dt-head">Order</div><div class="dt-head">รายละเอียด</div><div class="dt-head">PG</div>
            <div class="dt-head">Status</div><div class="dt-head">วันเปิด</div><div class="dt-head right">อายุ (วัน)</div>
            <div class="dt-head right">Plan ฿</div><div class="dt-head right">Actual ฿</div>
            ${m.tableRows.map((r) => `
              <div class="dt-cell mono">${esc(r.order)}</div>
              <div class="dt-cell">${esc(r.desc)}</div>
              <div class="dt-cell strong">${esc(r.pg)}</div>
              <div class="dt-cell"><span class="status-pill" style="color:${r.statusColor}"><span class="dot" style="background:${r.statusColor}"></span>${esc(r.statusLabel)}</span></div>
              <div class="dt-cell mono">${esc(r.openStr)}</div>
              <div class="dt-cell right strong" style="color:${r.ageColor}">${esc(r.age)}</div>
              <div class="dt-cell right">${esc(r.planStr)}</div>
              <div class="dt-cell right">${esc(r.actStr)}</div>`).join("")}
          </div>
        </div>
      </div>
    </div>`;
  }

  function renderSettingsPanel() {
    const s = state.settings;
    return `
      <div class="settings-panel iw-noprint">
        <div class="settings-row">
          <label>เกณฑ์ค้างนาน (วัน): <b>${s.overdueDays}</b></label>
          <input type="range" id="setOverdue" min="7" max="180" step="1" value="${s.overdueDays}">
        </div>
        <div class="settings-row">
          <label>ช่วงแนวโน้มรายเดือน (เดือน): <b>${s.monthsBack}</b></label>
          <input type="range" id="setMonths" min="6" max="24" step="1" value="${s.monthsBack}">
        </div>
        <div class="settings-row">
          <label>สีหลัก</label>
          <div class="swatches">
            ${["#2b5cc4", "#2e7d6b", "#7a4fd0", "#b4532a"].map((c) => `<button class="swatch ${c === s.accent ? "swatch-active" : ""}" data-color="${c}" style="background:${c}"></button>`).join("")}
          </div>
        </div>
      </div>`;
  }

  // ---------------------------------------------------------------
  // event wiring (re-attached after every render since innerHTML is replaced)
  // ---------------------------------------------------------------
  function wireEvents() {
    const dropbox = document.getElementById("dropbox");
    if (dropbox) {
      ["dragover", "dragenter"].forEach((ev) => dropbox.addEventListener(ev, (e) => { e.preventDefault(); dropbox.classList.add("drag-over"); }));
      ["dragleave", "drop"].forEach((ev) => dropbox.addEventListener(ev, (e) => { e.preventDefault(); dropbox.classList.remove("drag-over"); }));
      dropbox.addEventListener("drop", (e) => {
        const files = e.dataTransfer ? [...e.dataTransfer.files] : [];
        ingestFiles(files.filter((f) => /\.xlsx$/i.test(f.name)));
      });
    }
    const filePick = document.getElementById("filePick");
    if (filePick) filePick.addEventListener("change", (e) => ingestFiles([...e.target.files]));

    const btnSettings = document.getElementById("btnSettings");
    if (btnSettings) btnSettings.addEventListener("click", () => { state.settingsOpen = !state.settingsOpen; render(); });
    const btnPrint = document.getElementById("btnPrint");
    if (btnPrint) btnPrint.addEventListener("click", () => window.print());
    const btnReset = document.getElementById("btnReset");
    if (btnReset) btnReset.addEventListener("click", resetAll);

    document.querySelectorAll("[data-pg]").forEach((el) => el.addEventListener("click", () => { state.pgFilter = el.getAttribute("data-pg"); render(); }));
    const statusSelect = document.getElementById("statusSelect");
    if (statusSelect) statusSelect.addEventListener("change", (e) => { state.statusFilter = e.target.value; render(); });

    const scBacklog = document.getElementById("scBacklog");
    if (scBacklog) scBacklog.addEventListener("click", () => { state.scope = "backlog"; render(); });
    const scAll = document.getElementById("scAll");
    if (scAll) scAll.addEventListener("click", () => { state.scope = "all"; render(); });

    const setOverdue = document.getElementById("setOverdue");
    if (setOverdue) setOverdue.addEventListener("input", (e) => { state.settings.overdueDays = +e.target.value; saveSettings(); render(); });
    const setMonths = document.getElementById("setMonths");
    if (setMonths) setMonths.addEventListener("input", (e) => { state.settings.monthsBack = +e.target.value; saveSettings(); render(); });
    document.querySelectorAll("[data-color]").forEach((el) => el.addEventListener("click", () => { state.settings.accent = el.getAttribute("data-color"); saveSettings(); render(); }));
  }

  // ---------------------------------------------------------------
  // boot
  // ---------------------------------------------------------------
  async function boot() {
    loadSettings();

    // 1) Paint instantly from this browser's last-known cache (no flash of
    //    the upload screen if we've shown a dashboard here before).
    const hadCache = loadCache();
    render();

    // 2) Connect to Firestore — whichever data is freshest (cache vs cloud)
    //    will win once the first snapshot arrives, and every later upload
    //    from ANY browser will live-update this page automatically.
    const fbReady = window.__fb ? true : await new Promise((resolve) => {
      const t = setTimeout(() => resolve(false), 6000);
      window.addEventListener("fb-ready", () => { clearTimeout(t); resolve(true); }, { once: true });
    });

    if (fbReady && startCloudSync()) {
      // give the cloud a moment to report in; if it has nothing at all
      // (brand-new project, no uploads yet) and we also have no cache,
      // fall back to the bundled ./uploads/ files or the upload screen.
      setTimeout(async () => {
        if (!state.loaded) {
          const ok = await tryAutoload();
          if (ok) pushToCloud({ records: state.iw28, name: state.name28 }, { records: state.iw38, name: state.name38 });
          else render();
        }
      }, 1500);
    } else {
      state.cloudStatus = "offline";
      if (!hadCache) {
        const ok = await tryAutoload();
        if (!ok) render();
      } else {
        render();
      }
    }
  }

  document.addEventListener("DOMContentLoaded", boot);
})();
