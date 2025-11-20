// Utilities ---------------------------------------------------------------
const $ = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => [...root.querySelectorAll(sel)];
const statusBadge = $("#statusBadge");
const btnBack = $("#btnBack");
const fileInput = $("#fileInput");
const searchInput = $("#searchInput");
const dateColumnSelect = $("#dateColumnSelect");
const timeColumnSelect = $("#timeColumnSelect");
const table = $("#dataTable");
const tHead = table.tHead || table.createTHead();
const tBody = table.tBodies[0] || table.createTBody();
const rowsCount = $("#rowsCount");
const landing = $("#landing");
const appSection = $("#appSection");
const yearLabel = $("#yearLabel");
const courseSection = $("#courseSection");
const sheetSelect = $("#sheetSelect");



/* ======= Mappa foglio → Google Calendar ID (rimane) ======= */
const CALENDAR_BY_SHEET = {
  Fust2:
    "c_013ea419a34139e404c9756601ca3c1e0065cd221281bf919e8d73ccea96dd8d@group.calendar.google.com",
  Frot2:
    "c_198eb75ebc89748a4a8d305a0033c0a23f10f1426c5943c2a461b746768736be@group.calendar.google.com",
  Cyse2:
    "c_18aa898d91b45e39dfbd80900347e13ac24c8bd4ad8f250288099eb51f999f38@group.calendar.google.com",
  Dolc2:
    "c_3be7b7fab10384a7f85430d3a3847f6bc88508746c8a6dee964ad3dee6d3fb5e@group.calendar.google.com",
  AgoD2:
    "c_6c8669cc9b35556376327ae1a269fccc59faf8f0ef9049222b0d1f018835cdae@group.calendar.google.com",
  "Fust A1":
    "c_4974dfb894175cda42b8909491ff216c5e76bda37e5f8f9971dfeb832dac2b44@group.calendar.google.com",
  "Cyse A1":
    "c_645321fed6640203fe366362c39783da363b7cbbff9df294063ee809189e1355@group.calendar.google.com",
  "Arti A1":
    "c_d5059f4709fcf82caa2b8bbbc17a044daeb6c037c667fbfd375025cd5fd1accd@group.calendar.google.com",
  "Syam A1":
    "c_1d561c548bceb07cf6797cd95611e2473fd74645566221d882df29ca053770ac@group.calendar.google.com",
};

/* === Download XLSX locale/statico === */
async function fetchXlsxArrayBuffer(url) {
  const res = await fetch(url, {
    cache: "no-store",
    headers: { "Cache-Control": "no-cache", Pragma: "no-cache" },
  });
  if (!res.ok) throw new Error("HTTP " + res.status);
  return await res.arrayBuffer();
}
function isLikelyXLSXArrayBuffer(buf) {
  if (!buf || !buf.byteLength) return false;
  const u8 = new Uint8Array(buf.slice(0, 4));
  return u8[0] === 0x50 && u8[1] === 0x4b && u8[2] === 0x03 && u8[3] === 0x04;
}

const DEFAULT_XLSX_URL = "data/calendario.xlsx";

async function loadLocalCalendar() {
  try {
    setStatus("Caricando Excel...");
    const buf = await fetchXlsxArrayBuffer(DEFAULT_XLSX_URL);
    if (!isLikelyXLSXArrayBuffer(buf))
      throw new Error("Il file ottenuto non è un .xlsx valido (no firma PK)");
    workbook = XLSX.read(buf, { type: "array" });

    const defaultSheet =
      workbook.SheetNames.find((n) => n.toLowerCase().includes("fust")) ||
      workbook.SheetNames[0];

    sheetSelect.innerHTML = "";
    workbook.SheetNames.forEach((name, i) => {
      const opt = document.createElement("option");
      opt.value = name;
      opt.textContent = `${i + 1}. ${name}`;
      sheetSelect.appendChild(opt);
    });

    sheetSelect.value = defaultSheet;
    loadSheet(defaultSheet);

    const _origLoadSheet = typeof loadSheet === "function" ? loadSheet : null;
    if (_origLoadSheet) {
      window.loadSheet = function (name) {
        _origLoadSheet(name);
        updateAddButtonLink();
      };
    }

    setStatus("Excel Caricato", "ok");
  } catch (err) {
    console.error("Errore nel caricamento calendario:", err);
    setStatus(
      "Errore nel calendario (non è un XLSX oppure URL non raggiungibile).",
      "err"
    );
  }
}
window.addEventListener("DOMContentLoaded", loadLocalCalendar);

let workbook = null;
let baseHeaders = [];
let currentHeaders = [];
let currentData = [];
let allRows = [];
let showAll = false;

let selectedYear = null;
let pendingSheetName = null;

// copertura giornaliera precomputata
let dayCoverage = new Map();
let coverageDateHeader = null;
// Penultima-lezione
let penultimateKeys = new Set();
let moduleHeaderName = null;

// Corsi per anno
const COURSES = {
  1: [
    { key: "fust", label: "Fust", sheet: "Fust1" },
    { key: "cyse", label: "Cyse", sheet: "Cyse1" },
    { key: "arti", label: "Arti", sheet: "Arti1" },
    { key: "syam", label: "Syam", sheet: "Syam1" },
    // --- Nuovi corsi Anno 1 ---
    { key: "enem1", label: "EneM", sheet: "EneM1" },
    { key: "agod1", label: "AgoD", sheet: "AgoD1" },
    { key: "imer1", label: "ImeR", sheet: "ImeR1" },
    { key: "dita1", label: "Dita", sheet: "Dita1" },
  ],
  2: [
    { key: "front", label: "Front", sheet: "Frot2" },
    { key: "cyse", label: "Cyse", sheet: "Cyse2" },
    { key: "dolc", label: "Dolc", sheet: "Dolc2" },
    { key: "fust", label: "Fust", sheet: "Fust2" },
    { key: "ago", label: "Ago", sheet: "AgoD2" },
    // --- Nuovi corsi Anno 2 ---
    { key: "enes2", label: "EneS", sheet: "EneS2" },
    { key: "iota2", label: "IotA", sheet: "IotA2" },
  ],
};

const DROP_HEADER_RE = /^(colonna|giorno|fust2)$/i;

// Avvio: se c'è ?year=1|2 salto la landing, altrimenti mostro la landing
window.addEventListener("DOMContentLoaded", () => {
  const params = new URLSearchParams(window.location.search);
  const forcedYear = params.get("year");

  if (forcedYear === "1" || forcedYear === "2") {
    // URL già completo → vai direttamente alla scelta corso di quell'anno
    document.getElementById("landing")?.classList.add("hidden");
    applyYearChoice(forcedYear);
  } else {
    // Nessun year → mostra la landing (scelta anno) e nascondi il resto
    document.getElementById("landing")?.classList.remove("hidden");
    courseSection.classList.add("hidden");
    appSection.classList.add("hidden");
    yearLabel.textContent = "—";
    btnBack.classList.add("hidden");
  }
});

function isEmptyCell(v) {
  if (v === null || v === undefined) return true;
  if (v instanceof Date) return false;
  if (typeof v === "number") return false;
  const s = String(v).trim();
  return s === "" || s === "-" || s === "—";
}
function shouldDropHeader(h, rows) {
  if (!h) return true;
  if (DROP_HEADER_RE.test(String(h).trim())) return true;
  return rows.every((r) => isEmptyCell(r[h]));
}
function setStatus(text, tone = "info") {
  statusBadge.textContent = text;
  const color =
    tone === "ok"
      ? "var(--accent)"
      : tone === "err"
        ? "var(--danger)"
        : "var(--brand)";
  statusBadge.style.borderColor = "var(--border)";
  statusBadge.style.boxShadow = "inset 0 0 0 1px var(--border)";
  statusBadge.style.color = "#fff";
  statusBadge.style.background = `linear-gradient(180deg, ${color}, ${shade(
    color,
    -20
  )})`;
}
function shade(hex, percent) {
  if (!hex.startsWith("#")) return hex;
  const num = parseInt(hex.slice(1), 16);
  let r = (num >> 16) + Math.round((255 * percent) / 100);
  let g = ((num >> 8) & 0x00ff) + Math.round((255 * percent) / 100);
  let b = (num & 0x0000ff) + Math.round((255 * percent) / 100);
  r = Math.max(Math.min(255, r), 0);
  g = Math.max(Math.min(255, g), 0);
  b = Math.max(Math.min(255, b), 0);
  return `#${(b | (g << 8) | (r << 16)).toString(16).padStart(6, "0")}`;
}
function excelToJson(ws) {
  const aoa = XLSX.utils.sheet_to_json(ws, {
    header: 1,
    raw: true,
    defval: "",
  });
  if (!aoa.length) return { headers: [], rows: [] };
  let headers = aoa[0].map((h) => sanitizeHeader(String(h || "Colonna")));
  if (headers.every((h) => h === "Colonna"))
    headers = aoa[0].map((_, i) => colName(i));
  const rows = aoa.slice(1).map((r) => {
    const obj = {};
    headers.forEach((h, i) => {
      obj[h] = formatCell(r[i]);
    });
    return obj;
  });
  return { headers, rows };
}
function sanitizeHeader(h) {
  return h
    .trim()
    .replace(/\s+/g, " ")
    .replace(/[\n\r]+/g, " ")
    .replace(/[<>"']/g, "")
    .slice(0, 80);
}
function colName(i) {
  let s = "";
  i++;
  while (i > 0) {
    let m = (i - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    i = Math.floor((i - 1) / 26);
  }
  return s;
}
function isExcelDate(v) {
  return typeof v === "number" && v > 59 && v < 60000;
}
function formatExcelDate(v) {
  try {
    const d = XLSX.SSF.parse_date_code(v);
    if (!d) return v;
    return new Date(
      Date.UTC(
        d.y,
        (d.m || 1) - 1,
        d.d || 1,
        d.H || 0,
        d.M || 0,
        Math.floor(d.S || 0)
      )
    );
  } catch (e) {
    return v;
  }
}
function formatCell(v) {
  if (isExcelDate(v)) return formatExcelDate(v);
  return v;
}
function renderOptions(selectEl, options) {
  selectEl.innerHTML = options
    .map((o) => `<option value="${String(o)}">${String(o)}</option>`)
    .join("");
}

// Date/time helpers
const DATE_HEADER_RE = /^(data|date)$/i;
const START_HEADER_RE = /^(dalle|ora ?inizio|inizio|start)$/i;
const END_HEADER_RE = /^(alle|ora ?fine|fine|end)$/i;
const TIME_HEADER_RE = /^(ora|orario|dalle|alle|inizio|fine|start|end)$/i;

function autoDetectDateHeader(headers) {
  const chosen =
    dateColumnSelect &&
      dateColumnSelect.value &&
      dateColumnSelect.value !== "— nessuna —"
      ? dateColumnSelect.value
      : null;
  if (chosen && headers.includes(chosen)) return chosen;
  return headers.find((h) => DATE_HEADER_RE.test(String(h))) || null;
}
function autoDetectStartHeader(headers) {
  return headers.find((h) => START_HEADER_RE.test(String(h))) || null;
}
function autoDetectEndHeader(headers) {
  return headers.find((h) => END_HEADER_RE.test(String(h))) || null;
}

function toMinutes(v) {
  if (v instanceof Date) return v.getUTCHours() * 60 + v.getUTCMinutes();
  if (typeof v === "number") {
    if (v >= 0 && v <= 1) return Math.round(v * 24 * 60);
    if (v > 59 && v < 2400) {
      const hh = Math.floor(v / 100),
        mm = Math.round(v % 100);
      return hh * 60 + mm;
    }
  }
  if (typeof v === "string") {
    const m = v.trim().match(/^(\d{1,2})[:.](\d{2})/);
    if (m) return parseInt(m[1]) * 60 + parseInt(m[2]);
  }
  return null;
}
function mergeIntervals(intervals) {
  const arr = intervals
    .filter(
      (iv) => iv && iv.start != null && iv.end != null && iv.end > iv.start
    )
    .sort((a, b) => a.start - b.start);
  const merged = [];
  for (const iv of arr) {
    if (!merged.length || iv.start > merged[merged.length - 1].end)
      merged.push({ start: iv.start, end: iv.end });
    else
      merged[merged.length - 1].end = Math.max(
        merged[merged.length - 1].end,
        iv.end
      );
  }
  return merged;
}
function coveredMinutesWithinNeeds(intervals) {
  const merged = mergeIntervals(intervals),
    needs = [
      { start: 9 * 60, end: 13 * 60 },
      { start: 14 * 60, end: 18 * 60 },
    ];
  let total = 0;
  for (const need of needs)
    for (const iv of merged) {
      const start = Math.max(need.start, iv.start),
        end = Math.min(need.end, iv.end);
      if (end > start) total += end - start;
    }
  return total;
}
function fmtDateIT(d) {
  try {
    if (!(d instanceof Date)) return String(d);
    const weekday = new Intl.DateTimeFormat("it-IT", { weekday: "short" })
      .format(d)
      .replace(/\.$/, "");
    const datePart = new Intl.DateTimeFormat("it-IT", {
      day: "2-digit",
      month: "2-digit",
      year: "2-digit",
    }).format(d);
    return `${weekday} ${datePart}`;
  } catch (e) {
    if (d instanceof Date) {
      const days = ["dom", "lun", "mar", "mer", "gio", "ven", "sab"];
      const wd = days[d.getDay()];
      const y = String(d.getFullYear()).slice(-2);
      const m = String(d.getMonth() + 1).padStart(2, "0");
      const da = String(d.getDate()).padStart(2, "0");
      return `${wd} ${da}/${m}/${y}`;
    }
    return String(d);
  }
}
function isLikelyTimeHeader(h) {
  return !!h && TIME_HEADER_RE.test(String(h).trim());
}
function fmtTimeFromFraction(fr) {
  const total = Math.round(fr * 24 * 60);
  const hh = Math.floor(total / 60),
    mm = total % 60;
  return String(hh).padStart(2, "0") + ":" + String(mm).padStart(2, "0");
}
function fmtTimeFromDate(d) {
  return (
    String(d.getUTCHours()).padStart(2, "0") +
    ":" +
    String(d.getUTCMinutes()).padStart(2, "0")
  );
}
function prettyValue(v, header) {
  if (v instanceof Date) {
    if (isLikelyTimeHeader(header)) return fmtTimeFromDate(v);
    return fmtDateIT(v);
  }
  if (typeof v === "number") {
    if (v >= 0 && v < 1) return fmtTimeFromFraction(v);
  }
  return String(v);
}
function toText(v) {
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  return String(v).toLowerCase();
}
function dateKeyFromVal(v) {
  const d = v instanceof Date ? v : typeof v === "string" ? new Date(v) : null;
  if (d && !isNaN(d)) return d.toISOString().slice(0, 10);
  return String(v);
}

function renderCourseButtons(year) {
  const area = $("#courseButtons");
  const list = COURSES[String(year)] || [];
  area.innerHTML = list
    .map(
      (c) =>
        `<button class="btn big primary" data-course="${c.key}">${c.label}</button>`
    )
    .join("");
  $$("#courseButtons [data-course]").forEach((btn) => {
    btn.addEventListener("click", () => {
      goToCalendarWithCourse(btn.dataset.course);
    });
  });
}

// Rileva colonna Modulo/UF e penultima lezione
const MODULE_HEADER_RE =
  /^(modulo|uf|unit[aà] ?formativa|materia|insegnamento|argomento)$/i;
function autoDetectModuleHeader(headers) {
  return headers.find((h) => MODULE_HEADER_RE.test(String(h).trim())) || null;
}
function keyForRowPenultimate(row, dateH, startH, moduleH) {
  const mod = row[moduleH] ?? "";
  const dateKey = dateKeyFromVal(row[dateH]) ?? "";
  const startMin = toMinutes(row[startH]);
  const t = startMin != null ? String(startMin) : "";
  return `${mod}__${dateKey}__${t}`;
}
function computePenultimateKeys(headers, rows) {
  penultimateKeys = new Set();
  moduleHeaderName = autoDetectModuleHeader(headers);
  const dateH = autoDetectDateHeader(headers),
    startH = autoDetectStartHeader(headers);
  if (!moduleHeaderName || !dateH) return;
  const groups = new Map();
  for (const r of rows) {
    const mod = String(r[moduleHeaderName] ?? "").trim();
    if (!mod) continue;
    const dVal = r[dateH];
    const d = dVal instanceof Date ? dVal : new Date(dVal);
    if (isNaN(d)) continue;
    const sMin = startH ? toMinutes(r[startH]) : null;
    if (!groups.has(mod)) groups.set(mod, []);
    groups.get(mod).push({ row: r, d, sMin });
  }
  for (const [mod, arr] of groups.entries()) {
    arr.sort((a, b) => {
      const cmpD = a.d - b.d;
      if (cmpD !== 0) return cmpD;
      const aa = a.sMin ?? -1,
        bb = b.sMin ?? -1;
      return aa - bb;
    });
    if (arr.length >= 2) {
      const penult = arr[arr.length - 2];
      const k = keyForRowPenultimate(
        penult.row,
        dateH,
        startH,
        moduleHeaderName
      );
      penultimateKeys.add(k);
    }
  }
}

// Copertura giornaliera
function computeDayCoverage(headers, rows) {
  dayCoverage = new Map();
  coverageDateHeader = null;
  const dateH = autoDetectDateHeader(headers),
    startH = autoDetectStartHeader(headers),
    endH = autoDetectEndHeader(headers);
  if (!(dateH && startH && endH)) return;
  coverageDateHeader = dateH;
  const map = new Map();
  rows.forEach((r) => {
    const key = dateKeyFromVal(r[dateH]);
    const sMin = toMinutes(r[startH]);
    const eMin = toMinutes(r[endH]);
    if (sMin != null && eMin != null && eMin > sMin) {
      if (!map.has(key)) map.set(key, { intervals: [] });
      map.get(key).intervals.push({ start: sMin, end: eMin });
    }
  });
  for (const [k, obj] of map.entries())
    dayCoverage.set(k, coveredMinutesWithinNeeds(obj.intervals));
}

// Rendering tabella
function renderTable(_headersInput, rowsBase) {
  let headers = baseHeaders.slice();
  let rows = rowsBase.slice();
  headers = headers.filter((h) => !shouldDropHeader(h, rows));

  // Mobile: unisci Dalle/Alle in "Orario" e nascondi UF
  const isMobile = window.innerWidth <= 520;
  if (isMobile) {
    const startH = autoDetectStartHeader(headers);
    const endH = autoDetectEndHeader(headers);
    if (startH && endH) {
      const combinedHeader = "Orario";
      const dateH = autoDetectDateHeader(headers);
      headers = headers.filter((h) => h !== startH && h !== endH);
      const datePos = dateH ? headers.indexOf(dateH) : -1;
      const insertPos = datePos >= 0 ? datePos + 1 : 0;
      headers.splice(insertPos, 0, combinedHeader);
      rows = rows.map((r) => {
        const start = prettyValue(r[startH], startH);
        const end = prettyValue(r[endH], endH);
        const startLine = start || "";
        const dashLine = start && end ? "-" : "";
        const endLine = end || "";
        return {
          ...r,
          [combinedHeader]: `${startLine}\n${dashLine}\n${endLine}`,
        };
      });
    }
    headers = headers.filter((h) => !/^uf$/i.test(String(h).trim()));
  }

  // filtro "da oggi"
  const dateHeader = headers.find((h) =>
    /^(data|date)$/i.test(String(h).trim())
  );
  if (dateHeader && !showAll) {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    rows = rows.filter((r) => {
      const v = r[dateHeader];
      const d = v instanceof Date ? v : new Date(v);
      if (isNaN(d)) return false;
      return d >= today;
    });
  }

  currentHeaders = headers;
  currentData = rows;
  const q = searchInput.value.trim().toLowerCase();
  const filtered = !q
    ? rows
    : rows.filter((row) =>
      Object.values(row).some((v) => toText(v).includes(q))
    );

  // header
  tHead.innerHTML = "";
  const trh = document.createElement("tr");
  headers.forEach((h) => {
    const th = document.createElement("th");
    th.textContent = h;
    trh.appendChild(th);
  });
  tHead.appendChild(trh);

  // body
  tBody.innerHTML = "";
  const frag = document.createDocumentFragment();
  filtered.forEach((row) => {
    const tr = document.createElement("tr");
    headers.forEach((h) => {
      const td = document.createElement("td");
      td.textContent = prettyValue(row[h], h);
      tr.appendChild(td);
    });
    // evidenzia penultima
    {
      const dateH = autoDetectDateHeader(baseHeaders);
      const startH = autoDetectStartHeader(baseHeaders);
      if (moduleHeaderName && dateH) {
        const k = keyForRowPenultimate(row, dateH, startH, moduleHeaderName);
        if (penultimateKeys.has(k)) tr.classList.add("row-penultimate");
      }
    }
    frag.appendChild(tr);
  });
  tBody.appendChild(frag);

  // separatore per giorno
  const _dateHeaderForSep = autoDetectDateHeader(headers);
  if (_dateHeaderForSep) {
    const createdRows2 = [...tBody.querySelectorAll("tr")];
    let prevKey = null;
    createdRows2.forEach((tr, idx) => {
      const r = filtered[idx];
      const key = dateKeyFromVal(r[_dateHeaderForSep]);
      if (idx > 0 && key !== prevKey) tr.classList.add("date-sep");
      else tr.classList.remove("date-sep");
      prevKey = key;
    });
  }

  // evidenziazione ore giornaliere
  if (coverageDateHeader) {
    const dateIdxInRendered = headers.indexOf(coverageDateHeader);
    const createdRows = [...tBody.querySelectorAll("tr")];
    createdRows.forEach((tr, idx) => {
      const r = filtered[idx];
      const key = dateKeyFromVal(r[coverageDateHeader]);
      const minutes = dayCoverage.get(key) || 0;
      tr.classList.remove("day-short", "day-very-short");
      if (dateIdxInRendered >= 0)
        tr.children[dateIdxInRendered]?.classList.remove("date-red");
      if (minutes <= 240) tr.classList.add("day-very-short");
      else if (minutes < 480) tr.classList.add("day-short");
    });
  }

  rowsCount.textContent = filtered.length
    ? `${filtered.length} righe visualizzate`
    : "Nessun dato da mostrare";
  renderOptions(dateColumnSelect, ["— nessuna —", ...headers]);
  renderOptions(timeColumnSelect, ["— nessuna —", ...headers]);


}

// Import e caricamento locale
async function handleFile(file) {
  if (!file) return;
  if (file.size > 10 * 1024 * 1024) {
    setStatus("File troppo grande (>10MB)", "err");
    return;
  }
  setStatus("Caricando Excel...");
  const data = await file.arrayBuffer();
  workbook = XLSX.read(data, { type: "array" });

  sheetSelect.innerHTML = "";
  workbook.SheetNames.forEach((name, i) => {
    const opt = document.createElement("option");
    opt.value = name;
    opt.textContent = `${i + 1}. ${name}`;
    sheetSelect.appendChild(opt);
  });

  let defaultSheet = null;
  if (pendingSheetName && workbook.SheetNames.includes(pendingSheetName))
    defaultSheet = pendingSheetName;
  else
    defaultSheet =
      workbook.SheetNames.find((n) => n.toLowerCase().includes("fust")) ||
      workbook.SheetNames[0];

  sheetSelect.value = defaultSheet;
  loadSheet(defaultSheet);
  setStatus("Excel Caricato", "ok");
  fileInput.value = "";
}
function loadSheet(name) {
  const ws = workbook.Sheets[name];
  const { headers, rows } = excelToJson(ws);
  baseHeaders = headers.slice();
  computeDayCoverage(baseHeaders, rows);
  computePenultimateKeys(baseHeaders, rows);
  allRows = rows.slice();
  currentHeaders = baseHeaders.slice();
  showAll = false;
  updateToggleButton();
  renderTable(baseHeaders, allRows);
  updateAddButtonLink();
}

// CLEAR
function clearAll() {
  tHead.innerHTML = "";
  tBody.innerHTML = "";
  rowsCount.textContent = "—";
  currentData = [];
  currentHeaders = [];
  baseHeaders = [];
  allRows = [];
  showAll = false;
  updateToggleButton();
  workbook = null;
  dayCoverage = new Map();
  coverageDateHeader = null;
  sheetSelect.innerHTML = "";
  dateColumnSelect.innerHTML = "";
  timeColumnSelect.innerHTML = "";
  searchInput.value = "";
  fileInput.value = "";
  setStatus("Nessun file");
}

function applyYearChoice(year) {
  selectedYear = String(year);
  localStorage.setItem("cal-anno", selectedYear);

  // Aggiorna l’URL mantenendo il dominio/percorso, aggiungendo ?year=X
  const base = window.location.pathname.replace(/\/+$/, "") || "/";
  window.history.replaceState(null, "", `${base}?year=${selectedYear}`);

  yearLabel.textContent = `Anno ${selectedYear}`;

  const logo = document.getElementById("courseLogo");
  if (logo) {
    logo.textContent = "ITS";
    logo.classList.remove("active-logo");
  }

  // passa dalla landing alla scelta corso
  document.getElementById("landing")?.classList.add("hidden");
  appSection.classList.add("hidden");
  courseSection.classList.remove("hidden");
  btnBack.classList.add("hidden");

  renderCourseButtons(selectedYear);

  const hint = $("#courseHint");
  if (hint) {
    hint.textContent =
      selectedYear === "1"
        ? "Seleziona un corso dell’Anno 1 per aprire il calendario."
        : "Seleziona un corso dell’Anno 2 per aprire il calendario.";
  }
}

function goToCalendarWithCourse(courseKey) {
  const list = COURSES[String(selectedYear)] || [];
  const course = list.find((c) => c.key === courseKey);
  if (!course) return;
  pendingSheetName = course.sheet || null;
  courseSection.classList.add("hidden");
  appSection.classList.remove("hidden");
  btnBack.classList.remove("hidden");
  if (
    workbook &&
    pendingSheetName &&
    workbook.SheetNames.includes(pendingSheetName)
  ) {
    sheetSelect.value = pendingSheetName;
    loadSheet(pendingSheetName);
  }
  const logo = document.getElementById("courseLogo");
  if (logo) {
    logo.textContent = course.label;
    logo.classList.add("active-logo");
  }
}

// Toggle "Mostra tutto"
function updateToggleButton() {
  const btn = $("#btnToggleAll");
  if (!btn) return;
  btn.textContent = showAll ? "📅 Mostra da oggi" : "📅 Mostra tutto";
  btn.title = showAll ? "Mostra solo da oggi" : "Mostra tutto il calendario";
}

// Eventi UI
fileInput?.addEventListener("change", (e) => handleFile(e.target.files[0]));
$("#btnBack")?.addEventListener("click", () => {
  appSection.classList.add("hidden");
  courseSection.classList.remove("hidden");
  btnBack.classList.add("hidden");
  pendingSheetName = null;
  const logo = document.getElementById("courseLogo");
  if (logo) {
    logo.textContent = "ITS";
    logo.classList.remove("active-logo");
  }
  yearLabel.textContent = `Anno ${selectedYear}`;
});
sheetSelect.addEventListener("change", (e) => loadSheet(e.target.value));
searchInput?.addEventListener("input", () => renderTable(baseHeaders, allRows));
dateColumnSelect?.addEventListener("change", () =>
  renderTable(baseHeaders, allRows)
);
timeColumnSelect?.addEventListener("change", () =>
  renderTable(baseHeaders, allRows)
);
$("#btnToggleAll")?.addEventListener("click", () => {
  showAll = !showAll;
  updateToggleButton();
  renderTable(baseHeaders, allRows);
});
$("#btnClearTop")?.addEventListener("click", clearAll);
$("#btnLoadTop")?.addEventListener("click", () => {
  fileInput.value = "";
  fileInput.click();
});
$$(".landing [data-anno]").forEach((btn) => {
  btn.addEventListener("click", () => applyYearChoice(btn.dataset.anno));
});

/* Helper: ID calendario del foglio selezionato */
function getCurrentSheetCalendarId() {
  if (!sheetSelect || !sheetSelect.value) return null;
  const sheetName = sheetSelect.value;
  return CALENDAR_BY_SHEET[sheetName] || null;
}

/* ====== Google Calendar Link ====== */
// Rimossa logica OAuth complessa, manteniamo solo il link diretto


/* Abbonati: apre Google Calendar con il cid del foglio corrente */
$("#btnPushEvents")?.addEventListener("click", () => {
  const calId = getCurrentSheetCalendarId();
  if (!calId)
    return setStatus(
      "Seleziona un corso/foglio con calendario associato.",
      "err"
    );
  const url = `https://calendar.google.com/calendar/u/0/r?cid=${encodeURIComponent(
    calId
  )}`;
  window.open(url, "_blank", "noopener,noreferrer");

});

/* Aggiorna stato/label del bottone “Abbonati…” in base al foglio */
function updateAddButtonLink() {
  const btn = $("#btnPushEvents");
  if (!btn) return;
  const calId = getCurrentSheetCalendarId();
  if (calId) {
    btn.disabled = false;
    btn.dataset.calId = calId;
    btn.textContent = "➕ Sincronizza su Google Calendar";
    btn.title =
      "Apri Google Calendar e aggiungi il calendario del corso attivo";
  } else {
    btn.disabled = true;
    delete btn.dataset.calId;
    btn.textContent = "➕ Sincronizza su Google Calendar";
  }
}

// Init
(function init() {
  localStorage.removeItem("cal-anno");
  const logo = document.getElementById("courseLogo");
  if (logo) {
    logo.textContent = "ITS";
    logo.classList.remove("active-logo");
  }
  setStatus("Carico calendario…");
  updateToggleButton();
})();
