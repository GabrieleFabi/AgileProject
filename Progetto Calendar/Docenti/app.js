// Calendario Docenti — Home + Calendario
// Carica un XLSX e genera i bottoni dei docenti; poi mostra il calendario filtrato per docente.

// --- Helpers DOM ---------------------------------------------------------
const $ = (sel, root = document) => root.querySelector(sel);
const statusBadge = $("#statusBadge");
const pageTitle = $("#pageTitle");
const subLabel = $("#subLabel");
const homeSection = $("#homeSection");
const calendarSection = $("#calendarSection");
const teacherList = $("#teacherList");
const teacherSearch = $("#teacherSearch");
const resetTeacherSearch = $("#resetTeacherSearch");

const dropZone = $("#dropZone");
const pickBtn = $("#pickBtn");
const fileInput = $("#fileInput");

const dateColumnSelect = $("#dateColumnSelect");
const timeColumnSelect = $("#timeColumnSelect");
const table = $("#dataTable");
const tHead = table.tHead || table.createTHead();
const tBody = table.tBodies[0] || table.createTBody();
const rowsCount = $("#rowsCount");
const prevPage = $("#prevPage");
const nextPage = $("#nextPage");
const pageInfo = $("#pageInfo");
const pageSizeSel = $("#pageSize");
const showAllBtn = $("#showAllBtn");
const searchInput = $("#searchInput");
const backHomeBtn = $("#backHomeBtn");

// (opzionale: se usi COURSE_OVERRIDES altrove)
function findCourseOverride(name) {
  const n = String(name || "");
  for (const o of (window.COURSE_OVERRIDES || [])) {
    if (!o || !o.test) continue;
    try { if (o.test.test(n)) return o; } catch {}
  }
  return null;
}

const COURSE_COLORS = new Map();  // cache, NON toccare

// 1) Mappa fissa corso → colore (tutti diversi) (puoi adattare)
const COURSE_FIXED = new Map([
  ["syam1", "#16A34A"], // Verde
  ["arti1", "#F97316"], // Arancione
  ["cyse1", "#FACC15"], // Giallo
  ["fust1", "#3B82F6"], // Blu
  ["agod2", "#8B5CF6"], // Viola
  ["fust2", "#92400E"], // Marrone
  ["dolc2", "#9CA3AF"], // Grigio
  ["cyse2", "#EC4899"], // Rosa
  ["frot2", "#F5DEB3"], // Beige
]);

const DEFAULT_XLSX_URL = 'data/calendario.xlsx';

async function fetchXlsxArrayBuffer(url) {
  const res = await fetch(url, {
    cache: 'no-store',
    headers: { 'Cache-Control': 'no-cache', 'Pragma': 'no-cache' }
  });
  if (!res.ok) throw new Error('HTTP '+res.status);
  return await res.arrayBuffer();
}

function isLikelyXLSXArrayBuffer(buf) {
  if (!buf || !buf.byteLength) return false;
  const u8 = new Uint8Array(buf.slice(0, 4));
  return u8[0] === 0x50 && u8[1] === 0x4B && u8[2] === 0x03 && u8[3] === 0x04; // "PK\003\004"
}

// ==========================
// Caricamento automatico XLSX locale (per Surge)
// =========================
async function loadLocalCalendar() {
  try {
    setStatus("Carico calendario…");
    const buf = await fetchXlsxArrayBuffer(DEFAULT_XLSX_URL);
    if (!isLikelyXLSXArrayBuffer(buf)) throw new Error("Il file ottenuto non è un .xlsx valido (assenza firma PK)");
    workbook = XLSX.read(buf, { type: "array" });
    buildTeacherList();
    const label = /script\.google\.com\/macros\/s\//i.test(DEFAULT_XLSX_URL) ? "(web app GAS)" : "(statico)";
    setStatus(`Calendario caricato ${label}`, "ok");
  } catch (err) {
    console.error("Errore nel caricamento calendario:", err);
    setStatus("Errore nel caricamento del calendario.", "err");
  }
}
window.addEventListener("DOMContentLoaded", loadLocalCalendar);

// --- Colori per corso & legenda -----------------------------------------
function hexToRgba(hex, alpha) {
  if (!hex) return `rgba(0,0,0,${alpha})`;
  let h = String(hex).trim();
  if (h[0] === '#') h = h.slice(1);
  const r = parseInt(h.slice(0,2),16), g = parseInt(h.slice(2,4),16), b = parseInt(h.slice(4,6),16);
  return `rgba(${r},${g},${b},${alpha})`;
}
function courseColor(courseName) {
  if (!courseName) {
    return { swatch:"#94a3b8", bg:"rgba(148,163,184,0.12)", hover:"rgba(148,163,184,0.22)", border:"#64748b" };
  }
  if (COURSE_COLORS.has(courseName)) return COURSE_COLORS.get(courseName);
  const key = String(courseName).trim().toLowerCase().replace(/\s+/g, "").replace(/a1$/, "1");
  const border = COURSE_FIXED.get(key);
  const c = border
    ? { swatch: border, bg: hexToRgba(border, 0.20), hover: hexToRgba(border, 0.32), border }
    : { swatch: "#64748b", bg: "rgba(100,116,139,0.20)", hover: "rgba(100,116,139,0.32)", border: "#64748b" };
  COURSE_COLORS.set(courseName, c);
  return c;
}

function buildLegendFromRows(rows) {
  const byCourse = new Map();
  rows.forEach(r => {
    const cName = r["Corso"];
    if (!cName) return;
    if (!byCourse.has(cName)) byCourse.set(cName, courseColor(cName));
  });

  const legend = document.getElementById("legend");
  if (!legend) return;
  legend.innerHTML = "";

  if (!byCourse.size) {
    legend.innerHTML = "<span class='muted'>Legenda corsi: nessun corso visibile</span>";
    return;
  }

  const frag = document.createDocumentFragment();
  [...byCourse.keys()].sort((a,b) => a.localeCompare(b, "it")).forEach(corso => {
    const c = byCourse.get(corso);
    const chip = document.createElement("button");
    chip.type = "button";
    chip.className = "legend-chip";
    chip.setAttribute("role", "switch");
    const active = selectedCourses.has(corso);
    chip.setAttribute("aria-checked", active ? "true" : "false");
    chip.classList.toggle("is-active", active);

    const sw = document.createElement("span");
    sw.className = "legend-swatch";
    sw.style.background = c.swatch;

    const label = document.createElement("span");
    label.textContent = corso;

    chip.appendChild(sw);
    chip.appendChild(label);

    chip.addEventListener("click", () => {
      if (selectedCourses.has(corso)) selectedCourses.delete(corso);
      else selectedCourses.add(corso);
      rowsForTeacher = applyFilters();
      currentPage = 1;
      renderTable();
    });

    frag.appendChild(chip);
  });
  legend.appendChild(frag);
}

// Stato
let workbook = null;
let headersRef = [];
let allRows = [];
let rowsForTeacher = [];
let teacherHeader = null;
let selectedTeacher = null;

// Paginazione
let currentPage = 1;
let pageSize = 25;

// Filtri
let showAll = false;
let searchQuery = "";
let selectedCourses = new Set();

// --- Progress Overlay (bloccante) ------------------------------------------
const overlayEl = document.getElementById("progressOverlay");
const fillEl = document.getElementById("progressFill");
const titleEl = document.getElementById("progressTitle");
const pctEl = document.getElementById("progressPct");
const countEl = document.getElementById("progressCount");

/* NEW: assicurati che all'avvio sia nascosto */
if (overlayEl) overlayEl.hidden = true;

let _progressTotal = 0;
let _progressDone = 0;

function showProgress({ title = "Elaborazione…", total = 0 } = {}) {
  _progressTotal = Math.max(0, total);
  _progressDone = 0;
  if (titleEl) titleEl.textContent = title;
  updateProgress(0);
  if (overlayEl) {
    overlayEl.hidden = false;
    // blocca scroll sotto
    document.documentElement.style.overflow = "hidden";
  }
}

function updateProgress(done, extraText) {
  _progressDone = Math.min(Math.max(0, done), _progressTotal || done);
  const pct = _progressTotal > 0 ? Math.round((_progressDone / _progressTotal) * 100) : 0;
  if (fillEl) fillEl.style.width = `${pct}%`;
  if (pctEl) pctEl.textContent = `${pct}%`;
  if (countEl) countEl.textContent = `${_progressDone} / ${_progressTotal || "?"}`;
  if (extraText && titleEl) titleEl.textContent = extraText;
}

function hideProgress(finalMessage) {
  if (finalMessage && titleEl) titleEl.textContent = finalMessage;
  if (overlayEl) {
    // piccola pausa per far vedere il 100%
    setTimeout(() => {
      overlayEl.hidden = true;
      document.documentElement.style.overflow = "";
    }, 250);
  }
}


// --- Utility UI -------------------------------------------------------------
function setStatus(text, tone = "info") {
  statusBadge.textContent = text;
  const color = tone === "ok" ? "#10b981" : tone === "err" ? "#ef4444" : "#3b82f6";
  statusBadge.style.borderColor = "var(--border)";
  statusBadge.style.boxShadow = "inset 0 0 0 1px var(--border)";
  statusBadge.style.color = "#fff";
  statusBadge.style.background = `linear-gradient(180deg, ${color}, #1f2937aa)`;
}

function scrollToTop(smooth = true) {
  const el = document.scrollingElement || document.documentElement;
  requestAnimationFrame(() => {
    if (smooth && "scrollBehavior" in document.documentElement.style) {
      el.scrollTo({ top: 0, left: 0, behavior: "smooth" });
    } else {
      window.scrollTo(0, 0);
    }
  });
}

// XLSX utils
function sanitizeHeader(h) {
  return String(h || "").trim().replace(/\s+/g, " ").replace(/[\n\r]+/g, " ").replace(/[<>\"']/g, "").slice(0, 80);
}
function colName(i) {
  let s = "", n = i + 1;
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
function isExcelDate(v) { return typeof v === "number" && v > 59 && v < 60000; }
function formatExcelDate(v) {
  try {
    const d = XLSX.SSF.parse_date_code(v); if (!d) return v;
    return new Date(Date.UTC(d.y, (d.m || 1) - 1, d.d || 1, d.H || 0, d.M || 0, Math.floor(d.S || 0)));
  } catch { return v; }
}
function formatCell(v) { if (isExcelDate(v)) return formatExcelDate(v); return v; }
function excelToJson(ws) {
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" });
  if (!aoa.length) return { headers: [], rows: [] };
  let headers = aoa[0].map((h) => sanitizeHeader(String(h || "Colonna")));
  if (headers.every((h) => h === "Colonna")) headers = aoa[0].map((_, i) => colName(i));
  const rows = aoa.slice(1).map((r) => { const obj = {}; headers.forEach((h, i) => (obj[h] = formatCell(r[i]))); return obj; });
  return { headers, rows };
}

// Data/ora
const DATE_HEADER_RE    = /^(data|date)$/i;
const START_HEADER_RE   = /^(dalle|ora ?inizio|inizio|start)$/i;
const END_HEADER_RE     = /^(alle|ora ?fine|fine|end)$/i;
const TIME_HEADER_RE    = /^(ora|orario|dalle|alle|inizio|fine|start|end)$/i;
const TEACHER_HEADER_RE = /^(docente|insegnante|prof|teacher|formatore)$/i;
// NEW: modulo/UF per il titolo evento
const MODULE_HEADER_RE  = /^(modulo|uf|unit[aà]\s?formativa|materia|insegnamento|argomento)$/i;

function autoDetectDateHeader(headers)    { const chosen = dateColumnSelect?.value && dateColumnSelect.value !== "— nessuna —" ? dateColumnSelect.value : null; if (chosen && headers.includes(chosen)) return chosen; return headers.find((h) => DATE_HEADER_RE.test(String(h))) || null; }
function autoDetectStartHeader(headers)   { return headers.find((h) => START_HEADER_RE.test(String(h))) || null; }
function autoDetectEndHeader(headers)     { return headers.find((h) => END_HEADER_RE.test(String(h))) || null; }
function autoDetectTeacherHeader(headers) { return headers.find((h) => TEACHER_HEADER_RE.test(String(h).trim())) || null; }
function autoDetectModuleHeader(headers)  { return headers.find(h => MODULE_HEADER_RE.test(String(h).trim())) || null; }

function isLikelyTimeHeader(h) { return !!h && TIME_HEADER_RE.test(String(h).trim()); }
function fmtDateIT(d) {
  try {
    if (!(d instanceof Date)) return String(d);
    const weekday = new Intl.DateTimeFormat("it-IT", { weekday: "short" }).format(d).replace(/\.$/, "");
    const datePart = new Intl.DateTimeFormat("it-IT", { day: "2-digit", month: "2-digit", year: "2-digit" }).format(d);
    return `${weekday} ${datePart}`;
  } catch {
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
function fmtTimeFromFraction(fr) { const total = Math.round(fr * 24 * 60); const hh = Math.floor(total / 60); const mm = total % 60; return String(hh).padStart(2, "0") + ":" + String(mm).padStart(2, "0"); }
function fmtTimeFromDate(d) { return String(d.getUTCHours()).padStart(2, "0") + ":" + String(d.getUTCMinutes()).padStart(2, "0"); }
function parseTimeValue(v) {
  if (v instanceof Date) return v.getUTCHours() * 60 + v.getUTCMinutes();
  if (typeof v === "number" && v >= 0 && v < 1) return Math.round(v * 24 * 60);
  const s = String(v || "").trim(); const m = s.match(/(\d{1,2}):(\d{2})/);
  if (m) return parseInt(m[1], 10) * 60 + parseInt(m[2], 10);
  return 0;
}
function prettyValue(v, header) {
  if (v instanceof Date) { if (isLikelyTimeHeader(header)) return fmtTimeFromDate(v); return fmtDateIT(v); }
  if (typeof v === "number" && v >= 0 && v < 1) return fmtTimeFromFraction(v);
  return String(v);
}

// Normalizzazione/pulizia colonne
function normalizeHeaderName(h) { return String(h || "").trim().toLowerCase(); }
function dropUnwantedColumns(headers, rows) {
  const norm = headers.map(normalizeHeaderName);
  const removeByName = new Set();
  const isMobile = window.innerWidth <= 520;

  norm.forEach((h, i) => {
    if (
      h === "giorno" ||
      h === "frot2" ||
      (/^uf$/.test(h) && isMobile) ||
      /^(docente|insegnante|prof|teacher|formatore)$/.test(h)
    ) removeByName.add(headers[i]);
  });

  const isColEmpty = (hdr) => rows.every((r) => { const v = r[hdr]; return v == null || String(v).trim() === ""; });
  headers.forEach((hdr) => { if (isColEmpty(hdr)) removeByName.add(hdr); });
  headers.forEach((hdr) => { if (normalizeHeaderName(hdr) === "colonna" && isColEmpty(hdr)) removeByName.add(hdr); });

  const keptHeaders = headers.filter((h) => h === "Corso" || !removeByName.has(h));
  const cleanedRows = rows.map((row) => { const o = {}; keptHeaders.forEach((h) => (o[h] = row[h])); return o; });
  return { headers: keptHeaders, rows: cleanedRows };
}

function normalizeCourseName(name) { return String(name || "").replace(/\s*A1\b/i, "1"); }

function compareByDateAndStart(a, b, dateH, startH) {
  const da = rowDate(a, dateH), db = rowDate(b, dateH);
  if (da && db) { const diff = da.getTime() - db.getTime(); if (diff !== 0) return diff; }
  else if (da) return -1; else if (db) return 1;
  if (startH) { const ta = parseTimeValue(a[startH]); const tb = parseTimeValue(b[startH]); return ta - tb; }
  return 0;
}

// Caricamento file locale
async function handleFile(file) {
  try {
    setStatus("Leggo il file…");
    const ab = await file.arrayBuffer();
    workbook = XLSX.read(ab, { type: "array" });
    buildTeacherList();
  } catch (e) {
    console.error(e);
    setStatus("Errore lettura file", "err");
  }
}

// Home: docenti unici
function titleCaseName(s) { s = String(s || "").trim().replace(/\s+/g, " "); if (!s) return ""; return s.toLowerCase().split(" ").map(w => w.charAt(0).toUpperCase() + w.slice(1)).join(" "); }
function isGarbageTeacherName(s) {
  const t = String(s || "").trim(); if (!t) return true;
  if (/[!#?]{2,}/.test(t)) return true;
  if (/(^|[\s_-])(err|error|errore|modulo|sheet|test|dummy)([\s_-]|$)/i.test(t)) return true;
  const words = t.split(/\s+/).filter(Boolean); if (words.length < 2) return true;
  const letters = (t.match(/[a-zà-ù]/gi) || []).length; if (letters < 4) return true;
  if (words.some(w => !/[a-zà-ù]/i.test(w))) return true;
  return false;
}

function buildTeacherList() {
  if (!workbook) return;
  const unique = new Map();
  for (const sheetName of workbook.SheetNames) {
    const ws = workbook.Sheets[sheetName];
    const { headers, rows } = excelToJson(ws);
    if (!headers.length || !rows.length) continue;
    const tHeader = autoDetectTeacherHeader(headers);
    if (!tHeader) continue;
    teacherHeader = tHeader;

    for (const r of rows) {
      const raw = String(r[tHeader] || "").trim().replace(/\s+/g, " ");
      if (!raw) continue;
      if (isGarbageTeacherName(raw)) continue;
      const key = raw.toLowerCase();
      if (!unique.has(key)) unique.set(key, titleCaseName(raw));
    }
  }

  const arr = Array.from(unique.values()).sort((a, b) => a.localeCompare(b, "it"));

  teacherList.innerHTML = "";
  if (!arr.length) {
    teacherList.innerHTML = "<div class='muted'>Nessun docente trovato nel file.</div>";
    setStatus("File caricato, ma nessun docente trovato", "err");
    return;
  }
  setStatus("File caricato — Docenti trovati", "ok");

  const frag = document.createDocumentFragment();
  arr.forEach((name) => {
    const btn = document.createElement("button");
    btn.className = "btn";
    btn.textContent = name;
    btn.addEventListener("click", () => openCalendarFor(name));
    frag.appendChild(btn);
  });
  teacherList.appendChild(frag);
  $("#teacherCount").textContent = `${arr.length} docenti`;

  teacherSearch?.addEventListener("input", () => {
    const q = teacherSearch.value.toLowerCase();
    [...teacherList.children].forEach(btn => { btn.style.display = btn.textContent.toLowerCase().includes(q) ? "" : "none"; });
  });
  resetTeacherSearch?.addEventListener("click", () => {
    teacherSearch.value = "";
    teacherSearch.dispatchEvent(new Event("input"));
  });
}

// Passo 2: apri calendario per docente
function openCalendarFor(displayName) {
  selectedTeacher = displayName;
  pageTitle.textContent = "Calendario Docenti";
  subLabel.innerHTML = `Lezioni di <strong>${displayName}</strong>`;
  $("#courseLogo").textContent = "DOCENTE";

  collectRowsForTeacher(displayName);

  homeSection.style.display = "none";
  calendarSection.style.display = "grid";
  scrollToTop();
  location.hash = `#docente=${encodeURIComponent(displayName)}`;
  selectedCourses.clear();
  showAll = false;
  searchQuery = "";
  currentPage = 1;
  renderTable();
  setStatus(rowsForTeacher.length ? "Pronto" : "Nessuna lezione trovata", rowsForTeacher.length ? "ok" : "err");
  updateGcalUi(); // <-- aggiorna stato bottoni Google
}

function collectRowsForTeacher(displayName) {
  headersRef = [];
  allRows = [];

  const teacherKey = String(displayName || "").trim().toLowerCase();
  for (const sheetName of workbook.SheetNames) {
    const ws = workbook.Sheets[sheetName];
    const { headers, rows } = excelToJson(ws);
    if (!headers.length || !rows.length) continue;

    const tHeader = teacherHeader || autoDetectTeacherHeader(headers);
    if (!tHeader) continue;

    const matches = rows.filter((r) => String(r[tHeader] || "").trim().toLowerCase() === teacherKey);
    if (!matches.length) continue;

    matches.forEach((r) => (r["Corso"] = normalizeCourseName(sheetName)));

    // Mobile: combina orari
    const isMobile = window.innerWidth <= 520;
    let processed = matches;
    let finalHeaders = headers.slice();
    if (!finalHeaders.includes("Corso")) finalHeaders.push("Corso");

    if (isMobile) {
      const startH = autoDetectStartHeader(headers);
      const endH = autoDetectEndHeader(headers);
      if (startH && endH) {
        const combinedHeader = "Orario";
        const dateH = autoDetectDateHeader(headers);
        finalHeaders = finalHeaders.filter((h) => h !== startH && h !== endH);
        const datePos = dateH ? finalHeaders.indexOf(dateH) : -1;
        const insertPos = datePos >= 0 ? datePos + 1 : 0;
        finalHeaders.splice(insertPos, 0, combinedHeader);

        processed = matches.map((r) => {
          const start = prettyValue(r[startH], startH);
          const end = prettyValue(r[endH], endH);
          const startLine = start || "";
          const dashLine = start && end ? "-" : "";
          const endLine = end || "";
          return { ...r, [combinedHeader]: `${startLine}\n${dashLine}\n${endLine}` };
        });
      }
    }

    if (!headersRef.length) headersRef = finalHeaders.filter(Boolean);
    allRows.push(...processed);
  }

  const cleaned = dropUnwantedColumns(headersRef || [], allRows);
  const dateH = autoDetectDateHeader(cleaned.headers);
  const startH = autoDetectStartHeader(cleaned.headers);
  if (dateH) cleaned.rows.sort((a, b) => compareByDateAndStart(a, b, dateH, startH));

  headersRef = cleaned.headers;
  allRows = cleaned.rows;
  rowsForTeacher = applyFilters();
}

// Date helpers / filtri
function startOfToday() { const d = new Date(); d.setHours(0,0,0,0); return d; }
function rowDate(row, dateHeader) {
  const v = row?.[dateHeader];
  if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  let s = String(v || "").trim();
  s = s.replace(/^(lun|mar|mer|gio|ven|sab|dom)\.?[ ,]+/i, "");
  const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (m) { const dd=+m[1], mm=+m[2]-1, yy=+m[3]; const yyyy = yy<100 ? 2000+yy : yy; return new Date(yyyy,mm,dd); }
  const dt = new Date(s); return isNaN(+dt) ? null : new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
}
function filterFromToday(rows) {
  const header = autoDetectDateHeader(headersRef) || (rows[0] && autoDetectDateHeader(Object.keys(rows[0])));
  if (!header) return rows;
  const today0 = startOfToday().getTime();
  return rows.filter(r => { const d = rowDate(r, header); return d ? d.getTime() >= today0 : true; });
}
function textMatchRow(row, q, headers) {
  if (!q) return true;
  const needle = q.toLowerCase();
  for (const h of headers) {
    const v = row[h]; if (v==null) continue;
    const s = (v instanceof Date) ? prettyValue(v, h) : String(v);
    if (String(s).toLowerCase().includes(needle)) return true;
  }
  return false;
}
function applyFilters() {
  let rows = showAll ? allRows.slice() : filterFromToday(allRows);
  rows = rows.filter(r => textMatchRow(r, searchQuery, headersRef));
  if (selectedCourses.size > 0) rows = rows.filter(r => selectedCourses.has(String(r["Corso"] || "")));
  const dateH2 = autoDetectDateHeader(headersRef), startH2 = autoDetectStartHeader(headersRef);
  if (dateH2) rows.sort((a, b) => compareByDateAndStart(a, b, dateH2, startH2));
  return rows;
}

// Rendering
function totalPages() { return Math.max(1, Math.ceil(rowsForTeacher.length / pageSize)); }
function getPageSlice() { const start = (currentPage - 1) * pageSize; return rowsForTeacher.slice(start, start + pageSize); }
function updatePagerUI() {
  pageInfo.textContent = `${currentPage} / ${totalPages()}`;
  prevPage.disabled = currentPage <= 1;
  nextPage.disabled = currentPage >= totalPages();
}
function renderOptions(selectEl, options) { if (!selectEl) return; selectEl.innerHTML = options.map((o) => `<option value="${String(o)}">${String(o)}</option>`).join(""); }
function updateShowAllLabel() { if (!showAllBtn) return; showAllBtn.textContent = showAll ? "📅 Mostra da oggi" : "📅 Mostra tutto"; }

function renderTable() {
  const headers = headersRef.slice();
  const rows = getPageSlice();

  tHead.innerHTML = "";
  tBody.innerHTML = "";
  updateShowAllLabel();

  if (!headers.length || !rows.length) {
    rowsCount.textContent = "—";
    renderOptions(dateColumnSelect, ["— nessuna —"]);
    renderOptions(timeColumnSelect, ["— nessuna —"]);
    buildLegendFromRows([]);
    updateGcalUi(); // aggiorna anche i bottoni Google
    return;
  }

  const trh = document.createElement("tr");
  headers.forEach((h) => { const th = document.createElement("th"); th.textContent = h; trh.appendChild(th); });
  tHead.appendChild(trh);

  const frag = document.createDocumentFragment();
  rows.forEach((row) => {
    const tr = document.createElement("tr");
    tr.className = "course-row";
    const corso = row["Corso"];
    if (corso) {
      const c = courseColor(corso);
      tr.dataset.course = corso;
      tr.style.background = c.bg;
      tr.style.borderLeft = `4px solid ${c.border}`;
      tr.addEventListener("mouseenter", () => tr.style.background = c.hover);
      tr.addEventListener("mouseleave", () => tr.style.background = c.bg);
    }
    headers.forEach((h) => { const td=document.createElement("td"); td.textContent = prettyValue(row[h], h); tr.appendChild(td); });
    frag.appendChild(tr);
  });
  tBody.appendChild(frag);

  renderOptions(dateColumnSelect, ["— nessuna —", ...headers]);
  renderOptions(timeColumnSelect, ["— nessuna —", ...headers]);

  const total = allRows.length; const vis = rowsForTeacher.length; rowsCount.textContent = showAll ? `${vis} lezioni totali` : `${vis} da oggi (${total} totali)`;
  updatePagerUI();
  buildLegendFromRows(allRows);
  updateGcalUi(); // <-- aggiorna stato bottoni Google
}

// Eventi UI
dropZone?.addEventListener("click", () => fileInput?.click());
pickBtn?.addEventListener("click", () => fileInput?.click());
fileInput?.addEventListener("change", (e) => { const f = e.target.files?.[0]; if (f) handleFile(f); });

["dragenter", "dragover"].forEach((ev) =>
  dropZone?.addEventListener(ev, (e) => { e.preventDefault(); e.stopPropagation(); dropZone.classList.add("dragover"); setStatus("Rilascia per caricare…"); })
);
["dragleave", "dragend", "drop"].forEach((ev) =>
  dropZone?.addEventListener(ev, (e) => { e.preventDefault(); e.stopPropagation(); dropZone.classList.remove("dragover"); })
);
dropZone?.addEventListener("drop", (e) => { const f = e.dataTransfer?.files?.[0]; if (f) handleFile(f); });

// Accessibilità tastiera
dropZone?.addEventListener("keydown", (e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); fileInput?.click(); } });

// Paginazione
prevPage?.addEventListener("click", () => { if (currentPage > 1) { currentPage--; renderTable(); } });
nextPage?.addEventListener("click", () => { if (currentPage < totalPages()) { currentPage++; renderTable(); } });
pageSizeSel?.addEventListener("change", (e) => { const v = parseInt(e.target.value, 10); pageSize = [25,50,100].includes(v) ? v : 25; currentPage = 1; renderTable(); });

// Search nel calendario
searchInput?.addEventListener("input", (e) => { searchQuery = String(e.target.value || ""); currentPage = 1; rowsForTeacher = applyFilters(); renderTable(); });
searchInput?.addEventListener("keydown", (e) => { if (e.key === "Escape") { searchInput.value = ""; searchQuery = ""; currentPage = 1; rowsForTeacher = applyFilters(); renderTable(); } });

// Toggle mostra da oggi / tutto
showAllBtn?.addEventListener("click", () => { showAll = !showAll; rowsForTeacher = applyFilters(); currentPage = 1; renderTable(); });

// Logo torna home
const courseLogo = $("#courseLogo");
function enableLogoAsHome(enable) {
  if (!courseLogo) return;
  if (enable) { courseLogo.classList.add("is-clickable"); courseLogo.addEventListener("click", goHomeFromLogo); }
  else { courseLogo.classList.remove("is-clickable"); courseLogo.removeEventListener("click", goHomeFromLogo); }
}
function goHomeFromLogo() {
  calendarSection.style.display = "none";
  homeSection.style.display = "grid";
  scrollToTop();
  pageTitle.textContent = "Calendario Docenti";
  subLabel.textContent = "Seleziona un docente per vedere le sue lezioni";
  location.hash = "";
  setStatus("File caricato — scegli un docente", "ok");
  enableLogoAsHome(false);
}
const _openCalendarFor = openCalendarFor;
openCalendarFor = function(displayName) { _openCalendarFor(displayName); enableLogoAsHome(true); };
const _backHomeHandler = () => {
  calendarSection.style.display = "none";
  homeSection.style.display = "grid";
  scrollToTop();
  pageTitle.textContent = "Calendario Docenti";
  subLabel.textContent = "Seleziona un docente per vedere le sue lezioni";
  location.hash = "";
  setStatus("File caricato — scegli un docente", "ok");
  enableLogoAsHome(false);
};
backHomeBtn?.removeEventListener("click", _backHomeHandler);
backHomeBtn?.addEventListener("click", _backHomeHandler);

// --- Init ---------------------------------------------------------------
(function init() {
  setStatus("Carica un file Excel (.xlsx)…");
  window.addEventListener("hashchange", () => {
    const m = location.hash.match(/#docente=([^&]+)/);
    if (m && workbook) { const name = decodeURIComponent(m[1]); openCalendarFor(name); }
  });
})();

// ===================================================
// GOOGLE CALENDAR (OAuth + Create/Delete con dedup)
// ===================================================
const GCAL = {
  CLIENT_ID: "835074642817-l007g9fchi8dbqpedev1hrqrkmjkd109.apps.googleusercontent.com", // ⬅️ SOSTITUISCI
  SCOPES: "https://www.googleapis.com/auth/calendar.events",
  tokenClient: null,
  gapiReady: false,
  gisReady: false,
  authed: false,
};

// Helpers per eventi
function minutesToLocalRfc3339(dateYYYYMMDD, minutes) {
  const pad = (n) => String(n).padStart(2, "0");
  const hh = Math.floor(minutes / 60), mm = minutes % 60;
  return `${dateYYYYMMDD}T${pad(hh)}:${pad(mm)}:00`;
}
function toMinutes(v){
  if (v instanceof Date) return v.getUTCHours()*60+v.getUTCMinutes();
  if (typeof v === "number" && v >= 0 && v < 1) return Math.round(v*24*60);
  const m = String(v||"").match(/^(\d{1,2})[:.](\d{2})$/);
  return m ? (+m[1]*60 + +m[2]) : null;
}
function dateKeyFromVal(v){
  const d = v instanceof Date ? v : new Date(v);
  return isNaN(d) ? String(v) : d.toISOString().slice(0,10);
}
function keyForRow(row, dateH, startH, moduleH){
  const mod = row[moduleH] ?? "";
  const dateKey = dateKeyFromVal(row[dateH]) ?? "";
  const sMin = toMinutes(row[startH]);
  const t = sMin != null ? String(sMin) : "";
  return `${mod}__${dateKey}__${t}`;
}

// Carica gapi/gis
window.addEventListener("load", () => {
  if (window.gapi?.load) {
    gapi.load("client", async () => {
      await gapi.client.init({ discoveryDocs: ["https://www.googleapis.com/discovery/v1/apis/calendar/v3/rest"] });
      GCAL.gapiReady = true; updateGcalUi();
    });
  }
  if (window.google?.accounts?.oauth2) {
    GCAL.tokenClient = google.accounts.oauth2.initTokenClient({
      client_id: GCAL.CLIENT_ID,
      scope: GCAL.SCOPES,
      callback: (resp) => { if (resp?.access_token) { GCAL.authed = true; updateGcalUi(); } },
    });
    GCAL.gisReady = true; updateGcalUi();
  }
});

function updateGcalUi(){
  const ready = GCAL.gapiReady && GCAL.gisReady;
  const btnConn = document.getElementById("btnGConnect");
  const btnPush = document.getElementById("btnPushEvents");
  const btnDel  = document.getElementById("btnDeleteEvents");
  if (btnConn){ btnConn.disabled = !ready; btnConn.textContent = GCAL.authed ? "✅ Connesso a Google" : "🔑 Connetti Google"; }
  if (btnPush){ btnPush.disabled = !(ready && GCAL.authed && rowsForTeacher?.length); }
  if (btnDel) { btnDel.disabled  = !(ready && GCAL.authed); }
}

document.getElementById("btnGConnect")?.addEventListener("click", () => {
  GCAL.tokenClient?.requestAccessToken({ prompt: GCAL.authed ? "" : "consent" });
});

// Costruisce eventi Docenti: summary "(aula) – (modulo)"
function buildEventsFromTeacher(headers, rows){
  const dateH   = autoDetectDateHeader(headers);
  const startH  = autoDetectStartHeader(headers);
  const endH    = autoDetectEndHeader(headers);
  const moduleH = autoDetectModuleHeader(headers);
  if (!dateH) return [];

  const LOCATION_HEADER_RE = /^(aula|sede|luogo|location)$/i;
  const locationH = headers.find(h => LOCATION_HEADER_RE.test(String(h)));
  const tz = "Europe/Rome";
  const out = [];

  for (const r of rows){
    const dVal = r[dateH];
    const d = dVal instanceof Date ? dVal : new Date(dVal);
    if (isNaN(d)) continue;

    const sMin = startH ? toMinutes(r[startH]) : null;
    const eMin = endH ? toMinutes(r[endH]) : null;

    const mod  = moduleH ? String(r[moduleH] ?? "").trim() : "";
    const aula = locationH ? String(r[locationH] || "").trim() : "";
    let summary = (aula && mod) ? `${aula} – ${mod}` : (aula || mod || "Lezione");

    const uid = keyForRow(r, dateH, startH, moduleH || "");

    const ev = {
      summary,
      description: "", // se vuoi: come negli Studenti, costruisci descrizione completa
      location: aula || undefined,
      extendedProperties: { private: { itsaa_uid: uid } },
    };

    const date = d.toISOString().slice(0,10);
    if (sMin != null && eMin != null && eMin > sMin){
      ev.start = { dateTime: minutesToLocalRfc3339(date, sMin), timeZone: tz };
      ev.end   = { dateTime: minutesToLocalRfc3339(date, eMin), timeZone: tz };
    } else {
      ev.start = { date };
      const next = new Date(d); next.setDate(d.getDate()+1);
      ev.end = { date: next.toISOString().slice(0,10) }; // all-day (end esclusivo)
    }
    out.push(ev);
  }
  return out;
}

async function findByPrivateProp(calendarId, key, value){
  if (!(key && value)) return false;
  try {
    const now = new Date();
    const min = new Date(now); min.setFullYear(now.getFullYear()-1);
    const max = new Date(now); max.setFullYear(now.getFullYear()+2);
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
  } catch { return false; }
}
function delay(ms){ return new Promise(r => setTimeout(r, ms)); }

// Aggiungi eventi (dedup)
document.getElementById("btnPushEvents")?.addEventListener("click", async () => {
  try{
    if (!(GCAL.gapiReady && GCAL.authed)) return setStatus("Connetti Google prima.", "err");
    if (!rowsForTeacher?.length)        return setStatus("Nessuna lezione da esportare.", "err");

    const events = buildEventsFromTeacher(headersRef, rowsForTeacher);
    if (!events.length) return setStatus("Nessun evento valido.", "err");

    if (!confirm(`Creare ${events.length} eventi su Calendar (dedup attivo)?`)) return;

    // === PROGRESS: avvia overlay ===
    showProgress({ title: "Aggiungo eventi al tuo Google Calendar…", total: events.length });

    let inserted=0, skipped=0, failed=0, i=0;
    for (const ev of events){
      const uid = ev.extendedProperties.private.itsaa_uid;
      try{
        // Messaggio dinamico
        updateProgress(i, `Verifica duplicati… (${i+1}/${events.length})`);

        if (await findByPrivateProp("primary","itsaa_uid", uid)) { 
          skipped++; 
        } else {
          updateProgress(i, `Creo evento… (${i+1}/${events.length})`);
          await gapi.client.calendar.events.insert({ calendarId:"primary", resource: ev });
          inserted++;
          await delay(40);
        }
      } catch(e){ failed++; }
      i++;
      updateProgress(i);
    }

    hideProgress("Completato.");
    setStatus(`Inseriti ${inserted}, già presenti ${skipped}, errori ${failed}.`, "ok");
  } catch(e){
    console.error(e);
    hideProgress("Errore durante la creazione.");
    setStatus("Errore durante la creazione.", "err");
  } finally { updateGcalUi(); }
});


// Elimina eventi (UID → fallback titolo)
document.getElementById("btnDeleteEvents")?.addEventListener("click", async () => {
  if (!(GCAL.gapiReady && GCAL.authed)) return setStatus("Connetti Google prima.", "err");
  if (!confirm("Eliminare eventi creati (UID) o, se non trovati, quelli con titolo che inizia con 'Lezione —' o che contiene ' – '?")) return;

  try{
    const calendarId = "primary";
    const now = new Date();
    const min = new Date(now); min.setFullYear(now.getFullYear()-1);
    const max = new Date(now); max.setFullYear(now.getFullYear()+2);

    // 1) prova con UID
    let pageToken=null, candidates=[];
    do{
      const resp = await gapi.client.calendar.events.list({
        calendarId, timeMin:min.toISOString(), timeMax:max.toISOString(),
        singleEvents:true, maxResults:2500, orderBy:"startTime", pageToken
      });
      const items = resp.result.items || [];
      candidates.push(...items.filter(ev => ev.extendedProperties?.private?.itsaa_uid));
      pageToken = resp.result.nextPageToken || null;
    } while(pageToken);

    // 2) fallback per titolo
    if (!candidates.length){
      pageToken=null;
      do{
        const resp = await gapi.client.calendar.events.list({
          calendarId, timeMin:min.toISOString(), timeMax:max.toISOString(),
          singleEvents:true, maxResults:2500, orderBy:"startTime", pageToken
        });
        const items = resp.result.items || [];
        const matches = items.filter(ev => {
          const s = ev.summary?.trim() || "";
          return s.startsWith("Lezione —") || s.includes(" – ");
        });
        candidates.push(...matches);
        pageToken = resp.result.nextPageToken || null;
      } while(pageToken);
    }

    if (!candidates.length) {
      setStatus("Nessun evento trovabile da eliminare.", "ok");
      return;
    }

    // === PROGRESS: avvia overlay ===
    showProgress({ title: "Elimino eventi dal tuo Google Calendar…", total: candidates.length });

    let deleted=0, failed=0, i=0;
    for (const ev of candidates){
      try{
        updateProgress(i, `Elimino… (${i+1}/${candidates.length})`);
        await gapi.client.calendar.events.delete({ calendarId, eventId: ev.id });
        deleted++;
        await delay(30);
      } catch(e){ failed++; }
      i++;
      updateProgress(i);
    }

    hideProgress("Completato.");
    setStatus(`Eliminati ${deleted} eventi. Errori ${failed}.`, "ok");
  } catch(e){
    console.error(e);
    hideProgress("Errore durante l'eliminazione.");
    setStatus("Errore durante la rimozione.", "err");
  }
});

