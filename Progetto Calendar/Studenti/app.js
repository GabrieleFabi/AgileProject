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

// --- Helpers per scaricare XLSX da Web App (JSON b64) o file statico ---
async function fetchXlsxArrayBuffer(url) {
  const res = await fetch(url, {
    // evita cache del browser
    cache: 'no-store',
    // alcuni proxy/CDN rispettano anche questi header di richiesta
    headers: { 'Cache-Control': 'no-cache', 'Pragma': 'no-cache' }
  });
  if (!res.ok) throw new Error('HTTP '+res.status);
  return await res.arrayBuffer();
}

function base64ToArrayBuffer(b64) {
  const binary = atob(b64);
  const len = binary.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i++) bytes[i] = binary.charCodeAt(i);
  return bytes.buffer;
}

function isLikelyXLSXArrayBuffer(buf) {
  if (!buf || !buf.byteLength) return false;
  const u8 = new Uint8Array(buf.slice(0, 4));
  return u8[0] === 0x50 && u8[1] === 0x4B && u8[2] === 0x03 && u8[3] === 0x04;
}

// ==========================
// Caricamento automatico XLSX locale (per Surge)
// ==========================

const DEFAULT_XLSX_URL = 'data/calendario.xlsx';

async function loadLocalCalendar() {
  try {
    setStatus("Carico calendario…");

    const buf = await fetchXlsxArrayBuffer(DEFAULT_XLSX_URL);

    if (!isLikelyXLSXArrayBuffer(buf)) {
      throw new Error("Il file ottenuto non è un .xlsx valido (no firma PK)");
    }

    workbook = XLSX.read(buf, { type: "array" });

    // Seleziona automaticamente il foglio predefinito (es. Fust)
    const defaultSheet =
      workbook.SheetNames.find((n) => n.toLowerCase().includes("fust")) ||
      workbook.SheetNames[0];

    // Popola la select e carica la tabella
    sheetSelect.innerHTML = "";
    workbook.SheetNames.forEach((name, i) => {
      const opt = document.createElement("option");
      opt.value = name;
      opt.textContent = `${i + 1}. ${name}`;
      sheetSelect.appendChild(opt);
    });

    sheetSelect.value = defaultSheet;
    loadSheet(defaultSheet);

    const IS_REMOTE = /^https?:\/\//i.test(DEFAULT_XLSX_URL);
    const LABEL_SRC = IS_REMOTE ? "(web app)" : "(locale)";
    setStatus(`Calendario caricato ${LABEL_SRC}`, "ok");
  } catch (err) {
    console.error("Errore nel caricamento calendario:", err);
    setStatus("Errore nel calendario (non è un XLSX oppure URL non raggiungibile).", "err");
  }
}

// Avvia caricamento all’apertura della pagina
window.addEventListener("DOMContentLoaded", loadLocalCalendar);

let workbook = null;
let baseHeaders = []; // headers originali del foglio
let currentHeaders = [];
let currentData = []; // ultimo dataset renderizzato (post filtro)
let allRows = []; // tutte le righe del foglio (sorgente per render)
let showAll = false; // stato toggle "Mostra tutto"

let selectedYear = null;
let pendingSheetName = null;

// copertura giornaliera precomputata
let dayCoverage = new Map(); // key (YYYY-MM-DD) -> minuti
let coverageDateHeader = null; // nome colonna data usato per calcolo
// Penultima-lezione (esame) -----------------------------------------------
let penultimateKeys = new Set(); // Set di chiavi riga da evidenziare
let moduleHeaderName = null; // nome colonna modulo/UF rilevato

// Corsi per anno: key = anno ("1" | "2")
const COURSES = {
  1: [
    { key: "fust", label: "Fust", sheet: "Fust A1" },
    { key: "cyse", label: "Cyse", sheet: "Cyse A1" },
    { key: "arti", label: "Arti", sheet: "Arti A1" },
    { key: "syam", label: "Syam", sheet: "Syam A1" },
  ],
  2: [
    { key: "front", label: "Front", sheet: "Frot2" },
    { key: "cyse", label: "Cyse", sheet: "Cyse2" },
    { key: "dolc", label: "Dolc", sheet: "Dolc2" },
    { key: "fust", label: "Fust", sheet: "Fust2" },
    { key: "ago", label: "Ago", sheet: "AgoD2" },
  ],
};

const DROP_HEADER_RE = /^(colonna|giorno|fust2)$/i;

// --- Avvio automatico in base alla query string ---
window.addEventListener("DOMContentLoaded", () => {
  const params = new URLSearchParams(window.location.search);
  const forcedYear = params.get("year");
  if (forcedYear === "1" || forcedYear === "2") {
    // salta la home e vai subito alla selezione corso
    landing.classList.add("hidden");
    applyYearChoice(forcedYear);
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
  if (!aoa.length) {
    return { headers: [], rows: [] };
  }
  let headers = aoa[0].map((h) => sanitizeHeader(String(h || "Colonna")));
  if (headers.every((h) => h === "Colonna")) {
    headers = aoa[0].map((_, i) => colName(i));
  }
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

// =======================
// Batch + Resume (localStorage) per Add/Delete eventi
// =======================
const LS_KEYS = {
  addQueue: "itsaa_add_queue_v1",
  delQueue: "itsaa_del_queue_v1",
};

function lsGetJSON(key, fallback = []) {
  try {
    const s = localStorage.getItem(key);
    return s ? JSON.parse(s) : fallback;
  } catch { return fallback; }
}
function lsSetJSON(key, val) {
  try { localStorage.setItem(key, JSON.stringify(val)); } catch {}
}
function lsDel(key) {
  try { localStorage.removeItem(key); } catch {}
}

// Stato runtime
let processingAdd = false;
let processingDel = false;

function updateProgressBadge(text) {
  setStatus(text, "info");
}

// ---------- ADD: enqueue + process in batch (20) ----------
function enqueueAddEvents(events) {
  // Normalizza payload (riduciamo dimensione in LS)
  const items = events.map(ev => ({
    uid: ev.extendedProperties?.private?.itsaa_uid || "",
    ev
  }));
  const q = lsGetJSON(LS_KEYS.addQueue, []);
  q.push(...items);
  lsSetJSON(LS_KEYS.addQueue, q);
  return q.length;
}

async function processAddQueue(batchSize = 20) {
  if (processingAdd) return;
  processingAdd = true;

  try {
    let q = lsGetJSON(LS_KEYS.addQueue, []);
    let inserted = 0, skipped = 0, failed = 0;

    while (q.length) {
      const batch = q.splice(0, batchSize);
      updateProgressBadge(`Creo eventi… (restanti: ${q.length + batch.length})`);

      // esegui batch in sequenza (potresti anche parallelizzare con Promise.allSettled)
      for (const item of batch) {
        const uid = item.uid;
        const ev = item.ev;

        try {
          // dedup: se già presente, skip
          const exists = await findByPrivateProp("primary", "itsaa_uid", uid);
          if (exists) { skipped++; continue; }

          await gapi.client.calendar.events.insert({
            calendarId: "primary",
            resource: ev,
            conferenceDataVersion: 0,
            supportsAttachments: false,
          });
          inserted++;
          await delay(50);
        } catch (e) {
          console.warn("Insert fail", e);
          failed++;
        }
      }

      // aggiorna LS con coda restante e UI
      lsSetJSON(LS_KEYS.addQueue, q);
      updateGcalUi();
    }

    if (inserted + skipped + failed === 0) {
      setStatus("Coda vuota: nulla da creare.", "ok");
    } else {
      setStatus(`Fatto: inseriti ${inserted}, già presenti ${skipped}, errori ${failed}.`, "ok");
    }
  } finally {
    processingAdd = false;
    updateGcalUi();
  }
}

function resumeAddQueueIfAny() {
  const q = lsGetJSON(LS_KEYS.addQueue, []);
  if (q.length && GCAL.authed && GCAL.gapiReady) {
    processAddQueue(20);
  }
}

// ---------- DELETE: enqueue + process in batch (20) ----------
function enqueueDeleteByEvents(items) {
  // in coda salviamo solo id evento
  const q = lsGetJSON(LS_KEYS.delQueue, []);
  q.push(...items.map(ev => ({ id: ev.id })));
  lsSetJSON(LS_KEYS.delQueue, q);
  return q.length;
}

async function processDeleteQueue(batchSize = 20) {
  if (processingDel) return;
  processingDel = true;

  try {
    let q = lsGetJSON(LS_KEYS.delQueue, []);
    let deleted = 0, failed = 0;

    while (q.length) {
      const batch = q.splice(0, batchSize);
      updateProgressBadge(`Elimino eventi… (restanti: ${q.length + batch.length})`);

      for (const item of batch) {
        try {
          await gapi.client.calendar.events.delete({
            calendarId: "primary",
            eventId: item.id,
          });
          deleted++;
          await delay(50);
        } catch (e) {
          console.warn("Delete fail", e);
          failed++;
        }
      }

      lsSetJSON(LS_KEYS.delQueue, q);
      updateGcalUi();
    }

    if (deleted + failed === 0) {
      setStatus("Coda vuota: nulla da eliminare.", "ok");
    } else {
      setStatus(`Eliminati ${deleted} eventi. Errori ${failed}.`, "ok");
    }
  } finally {
    processingDel = false;
    updateGcalUi();
  }
}

function resumeDeleteQueueIfAny() {
  const q = lsGetJSON(LS_KEYS.delQueue, []);
  if (q.length && GCAL.authed && GCAL.gapiReady) {
    processDeleteQueue(20);
  }
}




// --- Helpers date/time ---------------------------------------------------
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
    if (m) {
      return parseInt(m[1]) * 60 + parseInt(m[2]);
    }
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
    if (!merged.length || iv.start > merged[merged.length - 1].end) {
      merged.push({ start: iv.start, end: iv.end });
    } else {
      merged[merged.length - 1].end = Math.max(
        merged[merged.length - 1].end,
        iv.end
      );
    }
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
  for (const need of needs) {
    for (const iv of merged) {
      const start = Math.max(need.start, iv.start);
      const end = Math.min(need.end, iv.end);
      if (end > start) total += end - start;
    }
  }
  return total;
}

function fmtDateIT(d) {
  try {
    if (!(d instanceof Date)) return String(d);

    // Giorno della settimana (es. lun, mar, mer, ...)
    const weekday = new Intl.DateTimeFormat("it-IT", { weekday: "short" })
      .format(d)
      .replace(/\.$/, ""); // rimuove eventuale punto finale

    const datePart = new Intl.DateTimeFormat("it-IT", {
      day: "2-digit",
      month: "2-digit",
      year: "2-digit",
    }).format(d);

    // Esempio: lun 27/10/25
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
  const hh = Math.floor(total / 60);
  const mm = total % 60;
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
  if (v instanceof Date) {
    return v.toISOString().slice(0, 10);
  }
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

  // collega i click dei nuovi bottoni
  $$("#courseButtons [data-course]").forEach((btn) => {
    btn.addEventListener("click", () => {
      goToCalendarWithCourse(btn.dataset.course);
    });
  });
}

// Rileva la colonna "Modulo/UF"
const MODULE_HEADER_RE =
  /^(modulo|uf|unit[aà] ?formativa|materia|insegnamento|argomento)$/i;
function autoDetectModuleHeader(headers) {
  return headers.find((h) => MODULE_HEADER_RE.test(String(h).trim())) || null;
}

// Chiave univoca per una lezione (per modulo + data + ora-inizio se c'è)
function keyForRowPenultimate(row, dateH, startH, moduleH) {
  const mod = row[moduleH] ?? "";
  const dateKey = dateKeyFromVal(row[dateH]) ?? "";
  const startMin = toMinutes(row[startH]);
  const t = startMin != null ? String(startMin) : "";
  return `${mod}__${dateKey}__${t}`;
}

// Calcola le penultime lezioni per ogni modulo
function computePenultimateKeys(headers, rows) {
  penultimateKeys = new Set();
  moduleHeaderName = autoDetectModuleHeader(headers);

  const dateH = autoDetectDateHeader(headers);
  const startH = autoDetectStartHeader(headers);

  if (!moduleHeaderName || !dateH) return;

  // Raggruppa per modulo
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

  // Ordina cronologicamente ogni gruppo e segna la penultima
  for (const [mod, arr] of groups.entries()) {
    arr.sort((a, b) => {
      // ordine per data, poi ora-inizio (se presente)
      const cmpD = a.d - b.d;
      if (cmpD !== 0) return cmpD;
      const aa = a.sMin ?? -1;
      const bb = b.sMin ?? -1;
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

// --- PRECOMPUTO copertura -----------------------------------------------
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
  for (const [k, obj] of map.entries()) {
    dayCoverage.set(k, coveredMinutesWithinNeeds(obj.intervals));
  }
}

// Rendering tabella -------------------------------------------------------
function renderTable(_headersInput, rowsBase) {
  // Riparti dagli header originali del foglio
  let headers = baseHeaders.slice();
  let rows = rowsBase.slice();

  headers = headers.filter((h) => !shouldDropHeader(h, rows));

  // --- Vista mobile: combina Dalle+Alle in "Orario" e nasconde UF ---
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

  // --- filtro "da oggi" se non Mostra tutto ---
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

  // --- Header ---
  tHead.innerHTML = "";
  const trh = document.createElement("tr");
  headers.forEach((h) => {
    const th = document.createElement("th");
    // niente classe "sortable", niente freccette e niente listener
    th.textContent = h;
    trh.appendChild(th);
  });
  tHead.appendChild(trh);

  // --- Body ---
  tBody.innerHTML = "";
  const frag = document.createDocumentFragment();
  filtered.forEach((row) => {
    const tr = document.createElement("tr");
    headers.forEach((h) => {
      const td = document.createElement("td");
      td.textContent = prettyValue(row[h], h);
      tr.appendChild(td);
    });
    // Evidenzia "penultima lezione" (esame) in verde
    {
      const dateH = autoDetectDateHeader(baseHeaders);
      const startH = autoDetectStartHeader(baseHeaders);
      if (moduleHeaderName && dateH) {
        const k = keyForRowPenultimate(row, dateH, startH, moduleHeaderName);
        if (penultimateKeys.has(k)) {
          tr.classList.add("row-penultimate");
        }
      }
    }
    frag.appendChild(tr);
  });
  tBody.appendChild(frag);

  // --- Separatore tra date ---
  const _dateHeaderForSep = autoDetectDateHeader(headers);
  if (_dateHeaderForSep) {
    const createdRows2 = [...tBody.querySelectorAll("tr")];
    let prevKey = null;
    createdRows2.forEach((tr, idx) => {
      const r = filtered[idx];
      const key = dateKeyFromVal(r[_dateHeaderForSep]);
      if (idx > 0 && key !== prevKey) {
        tr.classList.add("date-sep");
      } else {
        tr.classList.remove("date-sep");
      }
      prevKey = key;
    });
  }

  // --- Evidenziazione ore giornaliere ---
  if (coverageDateHeader) {
    const dateIdxInRendered = headers.indexOf(coverageDateHeader);
    const createdRows = [...tBody.querySelectorAll("tr")];
    createdRows.forEach((tr, idx) => {
      const r = filtered[idx];
      const key = dateKeyFromVal(r[coverageDateHeader]);
      const minutes = dayCoverage.get(key) || 0;

      // pulizia stati precedenti
      tr.classList.remove("day-short", "day-very-short");
      if (dateIdxInRendered >= 0) {
        tr.children[dateIdxInRendered]?.classList.remove("date-red");
      }

      // <4h = rosso chiaro (riga), 4–8h = arancione (riga)
      if (minutes <= 240) {
        tr.classList.add("day-very-short");
      } else if (minutes < 480) {
        tr.classList.add("day-short");
      }
    });
  }

  rowsCount.textContent = filtered.length
    ? `${filtered.length} righe visualizzate`
    : "Nessun dato da mostrare";
  renderOptions(dateColumnSelect, ["— nessuna —", ...headers]);
  renderOptions(timeColumnSelect, ["— nessuna —", ...headers]);
}


// Import e caricamento ----------------------------------------------------
async function handleFile(file) {
  if (!file) return;
  if (file.size > 10 * 1024 * 1024) {
    setStatus("File troppo grande (>10MB)", "err");
    return;
  }
  setStatus(`Caricamento: ${file.name}…`);
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
  if (pendingSheetName && workbook.SheetNames.includes(pendingSheetName)) {
    defaultSheet = pendingSheetName;
  } else {
    defaultSheet =
      workbook.SheetNames.find((n) => n.toLowerCase().includes("fust")) ||
      workbook.SheetNames[0];
  }
  sheetSelect.value = defaultSheet;
  loadSheet(defaultSheet);

  setStatus(`Pronto: ${file.name}`, "ok");
  fileInput.value = "";
}

function loadSheet(name) {
  const ws = workbook.Sheets[name];
  const { headers, rows } = excelToJson(ws);

  baseHeaders = headers.slice(); // salva headers originali
  computeDayCoverage(baseHeaders, rows);
  computePenultimateKeys(baseHeaders, rows);

  allRows = rows.slice();
  currentHeaders = baseHeaders.slice();
  showAll = false; // si parte da "da oggi"
  updateToggleButton();

  renderTable(baseHeaders, allRows);
}

// CLEAR -------------------------------------------------------------------
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

// Anno / corso ----------------------------------------------------
function applyYearChoice(year) {
  selectedYear = String(year);
  localStorage.setItem("cal-anno", selectedYear);
  yearLabel.textContent = `Anno ${selectedYear}`;

  const logo = document.getElementById("courseLogo");
  if (logo) {
    logo.textContent = "ITS";
    logo.classList.remove("active-logo");
  }

  // mostra la pagina corsi
  landing.classList.add("hidden");
  appSection.classList.add("hidden");
  courseSection.classList.remove("hidden");

  // costruisci i bottoni in base all'anno
  renderCourseButtons(selectedYear);

  // hint
  $("#courseHint").textContent =
    selectedYear === "1"
      ? "Seleziona un corso dell’Anno 1 per aprire il calendario."
      : "Seleziona un corso dell’Anno 2 per aprire il calendario.";
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

  // logo con nome corso
  const logo = document.getElementById("courseLogo");
  if (logo) {
    logo.textContent = course.label;
    logo.classList.add("active-logo");
  }
}

// --- Toggle "Mostra tutto" ----------------------------------------------
function updateToggleButton() {
  const btn = $("#btnToggleAll");
  if (!btn) return;
  btn.textContent = showAll ? "📅 Mostra da oggi" : "📅 Mostra tutto";
  btn.title = showAll ? "Mostra solo da oggi" : "Mostra tutto il calendario";
}

// Eventi UI ---------------------------------------------------------------
fileInput?.addEventListener("change", (e) => handleFile(e.target.files[0]));
$("#btnBack")?.addEventListener("click", () => {
  // torna al menu corsi
  appSection.classList.add("hidden");
  courseSection.classList.remove("hidden");
  btnBack.classList.add("hidden"); // nascondi perché è la prima pagina visibile ora

  pendingSheetName = null;

  // ripristina il logo
  const logo = document.getElementById("courseLogo");
  if (logo) {
    logo.textContent = "ITS";
    logo.classList.remove("active-logo");
  }

  yearLabel.textContent = `Anno ${selectedYear}`;
});

sheetSelect.addEventListener("change", (e) => loadSheet(e.target.value));

// rigenera sempre partendo dagli header grezzi
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

// =======================
// Google Calendar (OAuth + insert events)
// =======================
const GCAL = {
  CLIENT_ID: "835074642817-l007g9fchi8dbqpedev1hrqrkmjkd109.apps.googleusercontent.com",
  SCOPES: "https://www.googleapis.com/auth/calendar.events",
  tokenClient: null,
  gapiReady: false,
  gisReady: false,
  authed: false,
};

// Carica gapi client quando pronto
window.addEventListener("load", () => {
  // gapi
  if (window.gapi?.load) {
    gapi.load("client", async () => {
      await gapi.client.init({
        // API Key non strettamente necessaria con discovery
        discoveryDocs: ["https://www.googleapis.com/discovery/v1/apis/calendar/v3/rest"],
      });
      GCAL.gapiReady = true;
      updateGcalUi();
    });
  }
  // GIS
  if (window.google?.accounts?.oauth2) {
    GCAL.tokenClient = google.accounts.oauth2.initTokenClient({
      client_id: GCAL.CLIENT_ID,
      scope: GCAL.SCOPES,
      callback: (resp) => {
        if (resp && resp.access_token) {
          GCAL.authed = true;
          updateGcalUi();
          resumeAddQueueIfAny();
          resumeDeleteQueueIfAny();
          updateGcalUi();
        }
      },
    });
    GCAL.gisReady = true;
    updateGcalUi();
  }
});

function updateGcalUi() {
  const btnConn = $("#btnGConnect");
  const btnPush = $("#btnPushEvents");
  const btnDel  = $("#btnDeleteEvents");
  const ready = GCAL.gapiReady && GCAL.gisReady;

  if (btnConn) {
    btnConn.disabled = !ready;
    btnConn.textContent = GCAL.authed ? "✅ Connesso a Google" : "🔑 Connetti Google";
  }
  if (btnPush) {
    btnPush.disabled = !(ready && GCAL.authed && currentData?.length);
  }
  if (btnDel) {
    btnDel.disabled = !(ready && GCAL.authed);
  }
}

$("#btnGConnect")?.addEventListener("click", () => {
  if (!GCAL.tokenClient) return;
  // Prompt consenso la prima volta, poi silenzioso
  GCAL.tokenClient.requestAccessToken({ prompt: GCAL.authed ? "" : "consent" });
});

$("#btnPushEvents")?.addEventListener("click", async () => {
  try {
    if (!(GCAL.gapiReady && GCAL.authed)) {
      setStatus("Connetti Google prima di creare eventi.", "err");
      return;
    }
    if (!currentData?.length) {
      setStatus("Nessuna riga da esportare.", "err");
      return;
    }

    setStatus("Preparo eventi…");
    const events = buildEventsFromRows(currentHeaders, currentData);
    if (!events.length) {
      setStatus("Nessun evento valido (manca Data o orari).", "err");
      return;
    }

    const ok = confirm(`Creare ${events.length} eventi in batch sul Calendar? (dedup attivo)`);
    if (!ok) return;

    const totalQueued = enqueueAddEvents(events);
    setStatus(`In coda ${totalQueued} eventi. Avvio creazione…`);
    await processAddQueue(20); // batch da 20
  } catch (e) {
    console.error(e);
    setStatus("Errore durante la creazione degli eventi.", "err");
  } finally {
    updateGcalUi();
  }
});


// --- Pulsante per eliminare eventi creati dallo script ITSAA ---
$("#btnDeleteEvents")?.addEventListener("click", async () => {
  if (!(GCAL.gapiReady && GCAL.authed)) {
    setStatus("Collega prima Google.", "err");
    return;
  }

  const ok = confirm("Vuoi eliminare TUTTI gli eventi ITSAA (proprietà privata) oppure, se non trovati, quelli con titolo 'Lezione —'?");
  if (!ok) return;

  try {
    setStatus("Scansiono calendario per selezionare cosa eliminare…");

    const calendarId = "primary";
    const now = new Date();
    const min = new Date(now); min.setFullYear(now.getFullYear() - 1);
    const max = new Date(now); max.setFullYear(now.getFullYear() + 2);

    let pageToken = null;
    let candidates = [];

    do {
      const resp = await gapi.client.calendar.events.list({
        calendarId,
        timeMin: min.toISOString(),
        timeMax: max.toISOString(),
        singleEvents: true,
        maxResults: 2500,
        orderBy: "startTime",
        pageToken,
      });
      const items = resp.result.items || [];

      // 1) preferisci quelli con extended property our UID
      const withUID = items.filter(ev => ev.extendedProperties?.private?.itsaa_uid);
      candidates.push(...withUID);
      pageToken = resp.result.nextPageToken || null;
    } while (pageToken);

    // 2) Se nulla trovato con UID, fallback per titolo "Lezione —"
    if (!candidates.length) {
      setStatus("Nessun evento con UID. Attivo fallback 'Lezione —'…");
      pageToken = null;
      do {
        const resp = await gapi.client.calendar.events.list({
          calendarId,
          timeMin: min.toISOString(),
          timeMax: max.toISOString(),
          singleEvents: true,
          maxResults: 2500,
          orderBy: "startTime",
          pageToken,
        });
        const items = resp.result.items || [];
        const titled = items.filter(ev => ev.summary?.trim()?.startsWith("Lezione —"));
        candidates.push(...titled);
        pageToken = resp.result.nextPageToken || null;
      } while (pageToken);
    }

    if (!candidates.length) {
      setStatus("Nessun evento trovabile da eliminare.", "ok");
      return;
    }

    const totalQueued = enqueueDeleteByEvents(candidates);
    setStatus(`In coda per eliminazione: ${totalQueued} eventi. Avvio rimozione…`);
    await processDeleteQueue(20);
  } catch (e) {
    console.error(e);
    setStatus("Errore durante la rimozione.", "err");
  }
});




function buildEventsFromRows(headers, rows) {
  const dateH = autoDetectDateHeader(headers);
  const startH = autoDetectStartHeader(headers);
  const endH = autoDetectEndHeader(headers);
  const moduleH = autoDetectModuleHeader(headers);

  if (!dateH) return [];

  // altre colonne utili
  const LOCATION_HEADER_RE = /^(aula|sede|luogo|location)$/i;
  const DESC_HEADER_RE = /^(descrizione|note|argomento|tema|contenuti)$/i;
  const locationH = headers.find(h => LOCATION_HEADER_RE.test(String(h)));
  const descH = headers.find(h => DESC_HEADER_RE.test(String(h)));

  const tz = "Europe/Rome";
  const events = [];

  for (const r of rows) {
    const dVal = r[dateH];
    const d = dVal instanceof Date ? dVal : new Date(dVal);
    if (isNaN(d)) continue;

    const sMin = startH ? toMinutes(r[startH]) : null;
    const eMin = endH ? toMinutes(r[endH]) : null;

    // UID stabile per dedup: modulo + data + start
    const uid = keyForRowPenultimate(r, dateH, startH, moduleH || "");

    // summary
    const mod = moduleH ? String(r[moduleH] ?? "").trim() : "";
    let summary = mod ? `Lezione — ${mod}` : "Lezione";

    const event = {
      summary,
      description: buildDescription(r, headers),
      location: locationH ? String(r[locationH] || "") : undefined,
      extendedProperties: { private: { itsaa_uid: uid } },
    };

    if (sMin != null && eMin != null && eMin > sMin) {
      const date = d.toISOString().slice(0, 10);
      const start = minutesToLocalRfc3339(date, sMin);
      const end   = minutesToLocalRfc3339(date, eMin);
      event.start = { dateTime: start, timeZone: "Europe/Rome" };
      event.end   = { dateTime: end,   timeZone: "Europe/Rome" };
    } else {
      // Evento giornaliero (se mancano orari)
      const date = d.toISOString().slice(0, 10);
      event.start = { date };
      // Google richiede end esclusivo per eventi all-day → aggiungiamo 1 giorno
      const nextDay = new Date(d); nextDay.setDate(d.getDate() + 1);
      event.end = { date: nextDay.toISOString().slice(0, 10) };
    }
    events.push(event);

  }
  return events;
}

function buildDescription(row, headers) {
  // Monta una descrizione leggibile con alcune colonne chiave
  const parts = [];
  const moduleH = autoDetectModuleHeader(headers);
  const dateH   = autoDetectDateHeader(headers);
  const startH  = autoDetectStartHeader(headers);
  const endH    = autoDetectEndHeader(headers);

  if (moduleH) parts.push(`Modulo/UF: ${String(row[moduleH] ?? "")}`);
  if (dateH)   parts.push(`Data: ${prettyValue(row[dateH], dateH)}`);
  if (startH || endH) {
    const s = startH ? prettyValue(row[startH], startH) : "";
    const e = endH ? prettyValue(row[endH], endH) : "";
    parts.push(`Orario: ${s}${s && e ? " - " : ""}${e}`);
  }

  // Aggiunge tutte le altre colonne non vuote (esclude Data/Orari/UF)
  const skip = new Set([moduleH, dateH, startH, endH].filter(Boolean));
  headers.forEach(h => {
    if (!skip.has(h)) {
      const v = row[h];
      if (!isEmptyCell(v)) parts.push(`${h}: ${prettyValue(v, h)}`);
    }
  });

  // UID nascosto di servizio (non necessario, ma utile in debug)
  const uid = keyForRowPenultimate(row, dateH, startH, moduleH || "");
  parts.push(`UID: ${uid}`);

  return parts.join("\n");
}

function minutesToLocalRfc3339(dateYYYYMMDD, minutes) {
  const pad = (n) => String(n).padStart(2, "0");
  const hh = Math.floor(minutes / 60);
  const mm = minutes % 60;
  // RFC3339 senza offset/Z: Google userà il campo timeZone per interpretarlo correttamente
  return `${dateYYYYMMDD}T${pad(hh)}:${pad(mm)}:00`;
}

async function pushEventsDedup(events) {
  const calendarId = "primary";
  let inserted = 0, skipped = 0;

  for (const ev of events) {
    const uid = ev.extendedProperties?.private?.itsaa_uid;
    try {
      // 1) check se già esiste evento con la stessa uid
      const exists = await findByPrivateProp(calendarId, "itsaa_uid", uid);
      if (exists) { skipped++; continue; }
      // 2) insert
      await gapi.client.calendar.events.insert({
        calendarId,
        resource: ev,
        conferenceDataVersion: 0,
        supportsAttachments: false,
      });
      inserted++;
      await delay(120); // piccolo pacing per evitare rate-limit
    } catch (e) {
      console.warn("Insert error for UID", uid, e);
      // se qualcosa va storto non blocchiamo tutta la coda
    }
  }
  return { inserted, skipped };
}

async function findByPrivateProp(calendarId, key, value) {
  if (!(key && value)) return false;
  try {
    // Finestra ampia: oggi-1y → oggi+2y
    const now = new Date();
    const min = new Date(now); min.setFullYear(now.getFullYear() - 1);
    const max = new Date(now); max.setFullYear(now.getFullYear() + 2);

    const resp = await gapi.client.calendar.events.list({
      calendarId,
      privateExtendedProperty: `${key}=${value}`,
      timeMin: min.toISOString(),
      timeMax: max.toISOString(),
      maxResults: 1,
      singleEvents: true,
      orderBy: "startTime",
    });
    return (resp.result?.items || []).length > 0;
  } catch (e) {
    console.warn("findByPrivateProp error", e);
    return false;
  }
}

function delay(ms){ return new Promise(r => setTimeout(r, ms)); }

// Rinfresca lo stato bottoni quando cambia la tabella
const _origRenderTable = renderTable;
renderTable = function(headers, rows) {
  _origRenderTable(headers, rows);
  updateGcalUi();
};




// Init --------------------------------------------------------------------
(function init() {
  // (Rimosso il ramo con window.CALENDAR_XLSX_URL)

  localStorage.removeItem("cal-anno");

  const logo = document.getElementById("courseLogo");
  if (logo) {
    logo.textContent = "ITS";
    logo.classList.remove("active-logo");
  }

  landing.classList.remove("hidden");
  courseSection.classList.add("hidden");
  appSection.classList.add("hidden");
  btnBack.classList.add("hidden");
  yearLabel.textContent = "Scegli un anno per iniziare";

  setStatus("Carico calendario…");
  updateToggleButton();
  // Se l'utente era a metà di un'operazione e torna online, proviamo a riprendere
  resumeAddQueueIfAny();
  resumeDeleteQueueIfAny();
})();