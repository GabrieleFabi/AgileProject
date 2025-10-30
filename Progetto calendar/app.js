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


/// Override facili: colori fissi, uno per corso (match esatto, case-insensitive)
const COURSE_OVERRIDES = [
  { test: /^frot2$/i,      base:"#fecaca", border:"#f87171" }, // rosso tenue
  { test: /^cyse2$/i,      base:"#bbf7d0", border:"#22c55e" }, // verde
  { test: /^dolc2$/i,      base:"#a5f3fc", border:"#06b6d4" }, // ciano/teal
  { test: /^fust2$/i,      base:"#ddd6fe", border:"#8b5cf6" }, // viola
  { test: /^agod2$/i,      base:"#fde68a", border:"#f59e0b" }, // arancio/ambra
  { test: /^fust\s?a1$/i,  base:"#c7d2fe", border:"#6366f1" }, // indaco (spazio opzionale)
  { test: /^cyse\s?a1$/i,  base:"#86efac", border:"#16a34a" }, // verde alternativo
  { test: /^arti\s?a1$/i,  base:"#fbcfe8", border:"#ec4899" }, // rosa
  { test: /^syam\s?a1$/i,  base:"#93c5fd", border:"#3b82f6" }, // blu
];

const COURSE_COLORS = new Map();  // cache, NON toccare

const PALETTE = [
  { base: "#93c5fd", border: "#3b82f6" }, // blu
  { base: "#a5f3fc", border: "#06b6d4" }, // ciano
  { base: "#bbf7d0", border: "#22c55e" }, // verde
  { base: "#fde68a", border: "#f59e0b" }, // arancio
  { base: "#fbcfe8", border: "#ec4899" }, // rosa
  { base: "#ddd6fe", border: "#8b5cf6" }, // viola
  { base: "#fecaca", border: "#f87171" }, // rosso tenue
  { base: "#c7d2fe", border: "#6366f1" }, // indaco
  { base: "#86efac", border: "#22c55e" }, // verde pastello
  { base: "#e9d5ff", border: "#9333ea" }, // viola pastello
];

function hashStr(s) {
  s = String(s || "");
  let h = 2166136261 >>> 0;         // FNV-1a-like
  for (let i = 0; i < s.length; i++) {
    h ^= s.charCodeAt(i);
    h = (h * 16777619) >>> 0;
  }
  return h >>> 0;
}

function findCourseOverride(name) {
  const n = String(name || "");
  for (const o of COURSE_OVERRIDES ) {
    if (!o || !o.test) continue;
    try { if (o.test.test(n)) return o; } catch {}
  }
  return null;
}
// --- Colori per corso & legenda -----------------------------------------
function hashHue(str) {
  str = String(str || "");
  let h = 0;
  for (let i = 0; i < str.length; i++) h = (h * 31 + str.charCodeAt(i)) >>> 0;
  return h % 360;
}
function courseColor(courseName) {
  if (!courseName) {
    return {
      swatch: "#94a3b8",
      bg: "rgba(148,163,184,0.12)",
      hover: "rgba(148,163,184,0.22)",
      border: "#64748b",
    };
  }

  // cache
  if (COURSE_COLORS.has(courseName)) return COURSE_COLORS.get(courseName);

  // 1) override manuale (se definito)
  const ov = findCourseOverride(courseName);
  let colors;
  if (ov) {
    colors = { base: ov.base, border: ov.border };
  } else {
    // 2) fallback: palette con indice deterministico
    const idx = hashStr(courseName) % PALETTE.length;
    colors = PALETTE[idx];
  }

  const c = {
    swatch: colors.border,
    bg: colors.base + "40",    // ~25% alpha
    hover: colors.base + "60", // ~38% alpha
    border: colors.border,
  };
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

    // chip = bottone toggle
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

    // toggle filtro al click
    chip.addEventListener("click", () => {
      if (selectedCourses.has(corso)) selectedCourses.delete(corso);
      else selectedCourses.add(corso);
      rowsForTeacher = applyFilters();
      currentPage = 1;
      renderTable(); // ricostruisce tabella e aggiorna legenda (stato chip)
    });

    frag.appendChild(chip);
  });

  legend.appendChild(frag);
}



// Stato
let workbook = null;
let headersRef = [];
let allRows = [];      // dataset completo filtrabile
let rowsForTeacher = [];
let teacherHeader = null;
let selectedTeacher = null;

// Paginazione
let currentPage = 1;
let pageSize = 25; // default

// Filtri
let showAll = false;   // default: mostra da OGGI in poi
let searchQuery = "";  // testo di ricerca
let selectedCourses = new Set();  // ← nuovo: corsi selezionati dalla legenda

// --- Utility UI -------------------------------------------------------------
function setStatus(text, tone = "info") {
  statusBadge.textContent = text;
  const color =
    tone === "ok"
      ? "#10b981"
      : tone === "err"
      ? "#ef4444"
      : "#3b82f6";
  statusBadge.style.borderColor = "var(--border)";
  statusBadge.style.boxShadow = "inset 0 0 0 1px var(--border)";
  statusBadge.style.color = "#fff";
  statusBadge.style.background = `linear-gradient(180deg, ${color}, #1f2937aa)`;
}

function sanitizeHeader(h) {
  return String(h || "")
    .trim()
    .replace(/\s+/g, " ")
    .replace(/[\n\r]+/g, " ")
    .replace(/[<>\"']/g, "")
    .slice(0, 80);
}

function colName(i) {
  let s = "", n = i + 1;
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
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
    return new Date(Date.UTC(d.y, (d.m || 1) - 1, d.d || 1, d.H || 0, d.M || 0, Math.floor(d.S || 0)));
  } catch {
    return v;
  }
}
function formatCell(v) {
  if (isExcelDate(v)) return formatExcelDate(v);
  return v;
}
function excelToJson(ws) {
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" });
  if (!aoa.length) return { headers: [], rows: [] };
  let headers = aoa[0].map((h) => sanitizeHeader(String(h || "Colonna")));
  if (headers.every((h) => h === "Colonna")) headers = aoa[0].map((_, i) => colName(i));
  const rows = aoa.slice(1).map((r) => {
    const obj = {};
    headers.forEach((h, i) => (obj[h] = formatCell(r[i])));
    return obj;
  });
  return { headers, rows };
}

// Data/ora pretty
const DATE_HEADER_RE = /^(data|date)$/i;
const START_HEADER_RE = /^(dalle|ora ?inizio|inizio|start)$/i;
const END_HEADER_RE = /^(alle|ora ?fine|fine|end)$/i;
const TIME_HEADER_RE = /^(ora|orario|dalle|alle|inizio|fine|start|end)$/i;
const TEACHER_HEADER_RE = /^(docente|insegnante|prof|teacher|formatore)$/i;

function autoDetectDateHeader(headers) {
  const chosen = dateColumnSelect?.value && dateColumnSelect.value !== "— nessuna —" ? dateColumnSelect.value : null;
  if (chosen && headers.includes(chosen)) return chosen;
  return headers.find((h) => DATE_HEADER_RE.test(String(h))) || null;
}
function autoDetectStartHeader(headers) {
  return headers.find((h) => START_HEADER_RE.test(String(h))) || null;
}
function autoDetectEndHeader(headers) {
  return headers.find((h) => END_HEADER_RE.test(String(h))) || null;
}
function autoDetectTeacherHeader(headers) {
  return headers.find((h) => TEACHER_HEADER_RE.test(String(h).trim())) || null;
}
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
function fmtTimeFromFraction(fr) {
  const total = Math.round(fr * 24 * 60);
  const hh = Math.floor(total / 60);
  const mm = total % 60;
  return String(hh).padStart(2, "0") + ":" + String(mm).padStart(2, "0");
}
function fmtTimeFromDate(d) {
  return String(d.getUTCHours()).padStart(2, "0") + ":" + String(d.getUTCMinutes()).padStart(2, "0");
}

function parseTimeValue(v) {
  if (v instanceof Date) return v.getUTCHours() * 60 + v.getUTCMinutes();
  if (typeof v === "number" && v >= 0 && v < 1) return Math.round(v * 24 * 60);
  const s = String(v || "").trim();
  const m = s.match(/(\d{1,2}):(\d{2})/);
  if (m) return parseInt(m[1], 10) * 60 + parseInt(m[2], 10);
  return 0;
}

function prettyValue(v, header) {
  if (v instanceof Date) {
    if (isLikelyTimeHeader(header)) return fmtTimeFromDate(v);
    return fmtDateIT(v);
  }
  if (typeof v === "number" && v >= 0 && v < 1) return fmtTimeFromFraction(v);
  return String(v);
}

// --- Normalizzazione e pulizia colonne ----------------------------------
function normalizeHeaderName(h) { return String(h || "").trim().toLowerCase(); }

function dropUnwantedColumns(headers, rows) {
  const norm = headers.map(normalizeHeaderName);
  const removeByName = new Set();

  norm.forEach((h, i) => {
    if (
      h === "giorno" ||
      h === "frot2" ||
      h === "uf" ||                              // ← nascondi sempre la colonna UF
      /^(docente|insegnante|prof|teacher|formatore)$/.test(h)
    ) removeByName.add(headers[i]);
  });

  const isColEmpty = (hdr) => rows.every((r) => {
    const v = r[hdr];
    return v === null || v === undefined || String(v).trim() === "";
  });
  headers.forEach((hdr) => { if (isColEmpty(hdr)) removeByName.add(hdr); });

  headers.forEach((hdr) => {
    if (normalizeHeaderName(hdr) === "colonna" && isColEmpty(hdr)) removeByName.add(hdr);
  });

  const keptHeaders = headers.filter((h) => h === 'Corso' || !removeByName.has(h));
  const cleanedRows = rows.map((row) => {
    const o = {}; keptHeaders.forEach((h) => (o[h] = row[h])); return o;
  });
  return { headers: keptHeaders, rows: cleanedRows };
}

function normalizeCourseName(name) { return String(name || "").replace(/\s*A1\b/i, "1"); }

function compareByDateAndStart(a, b, dateH, startH) {
  const da = rowDate(a, dateH);
  const db = rowDate(b, dateH);
  if (da && db) {
    const diff = da.getTime() - db.getTime();
    if (diff !== 0) return diff;
  } else if (da) return -1;
  else if (db) return 1;
  if (startH) {
    const ta = parseTimeValue(a[startH]);
    const tb = parseTimeValue(b[startH]);
    return ta - tb;
  }
  return 0;
}

// --- Caricamento da file locale -----------------------------------------
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

// --- Home: estrai docenti unici -----------------------------------------
function titleCaseName(s) {
  s = String(s || "").trim().replace(/\s+/g, " ");
  if (!s) return "";
  return s.toLowerCase().split(" ").map(w => w.charAt(0).toUpperCase() + w.slice(1)).join(" ");
}

function buildTeacherList() {
  if (!workbook) return;
  const unique = new Map(); // key lower -> display
  let foundCount = 0;

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
      // deve avere almeno nome + cognome
      if (raw.split(" ").length < 2) continue;
      const key = raw.toLowerCase();
      if (!unique.has(key)) unique.set(key, titleCaseName(raw));
    }
  }

  const arr = Array.from(unique.values()).sort((a, b) => a.localeCompare(b, "it"));
  foundCount = arr.length;

  // Render
  teacherList.innerHTML = "";
  if (!foundCount) {
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
  $("#teacherCount").textContent = `${foundCount} docenti`;

  // Search on home
  teacherSearch?.addEventListener("input", () => {
    const q = teacherSearch.value.toLowerCase();
    [...teacherList.children].forEach(btn => {
      btn.style.display = btn.textContent.toLowerCase().includes(q) ? "" : "none";
    });
  });
  resetTeacherSearch?.addEventListener("click", () => {
    teacherSearch.value = "";
    teacherSearch.dispatchEvent(new Event("input"));
  });
}

// --- Passo 2: apri calendario per docente -------------------------------
function openCalendarFor(displayName) {
  selectedTeacher = displayName;
  pageTitle.textContent = "Calendario Docenti";
  subLabel.innerHTML = `Lezioni di <strong>${displayName}</strong>`;
  $("#courseLogo").textContent = "DOCENTE";

  // Prepara dataset filtrato da workbook
  collectRowsForTeacher(displayName);

  // Switch view
  homeSection.style.display = "none";
  calendarSection.style.display = "grid";
  // Aggiorna hash per deep-link
  location.hash = `#docente=${encodeURIComponent(displayName)}`;
  selectedCourses.clear(); // azzera eventuali selezioni corsi precedenti
  // Reset UI calendar
  showAll = false;
  searchQuery = "";
  currentPage = 1;
  renderTable();
  setStatus(rowsForTeacher.length ? "Pronto" : "Nessuna lezione trovata", rowsForTeacher.length ? "ok" : "err");
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

    // Mobile: combina orari se possibile
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

  // Pulizia colonne (incluso nascondere UF)
  const cleaned = dropUnwantedColumns(headersRef || [], allRows);
  const dateH = autoDetectDateHeader(cleaned.headers);
  const startH = autoDetectStartHeader(cleaned.headers);
  if (dateH) cleaned.rows.sort((a, b) => compareByDateAndStart(a, b, dateH, startH));

  headersRef = cleaned.headers;
  allRows = cleaned.rows;
  rowsForTeacher = applyFilters();
}

// --- Date helpers --------------------------------------------------------
function startOfToday() { const d = new Date(); d.setHours(0,0,0,0); return d; }
function rowDate(row, dateHeader) {
  const v = row?.[dateHeader];
  if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), v.getDate());

  let s = String(v || "").trim();
  s = s.replace(/^(lun|mar|mer|gio|ven|sab|dom)\.?[ ,]+/i, "");

  const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (m) {
    const dd = parseInt(m[1], 10), mm = parseInt(m[2], 10) - 1, yy = parseInt(m[3], 10);
    const yyyy = yy < 100 ? 2000 + yy : yy;
    return new Date(yyyy, mm, dd);
  }
  const dt = new Date(s);
  return isNaN(+dt) ? null : new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
}

function filterFromToday(rows) {
  const header = autoDetectDateHeader(headersRef) || (rows[0] && autoDetectDateHeader(Object.keys(rows[0])));
  if (!header) return rows;
  const today0 = startOfToday().getTime();
  return rows.filter(r => {
    const d = rowDate(r, header);
    return d ? d.getTime() >= today0 : true;
  });
}

function textMatchRow(row, q, headers) {
  if (!q) return true;
  const needle = q.toLowerCase();
  for (const h of headers) {
    const v = row[h];
    if (v == null) continue;
    const s = (v instanceof Date) ? prettyValue(v, h) : String(v);
    if (String(s).toLowerCase().includes(needle)) return true;
  }
  return false;
}
function applyFilters() {
  let rows = showAll ? allRows.slice() : filterFromToday(allRows);

  // filtro testo
  rows = rows.filter(r => textMatchRow(r, searchQuery, headersRef));

  // ← nuovo: filtro corsi se c’è almeno una selezione
  if (selectedCourses.size > 0) {
    rows = rows.filter(r => selectedCourses.has(String(r["Corso"] || "")));
  }

  const dateH2 = autoDetectDateHeader(headersRef);
  const startH2 = autoDetectStartHeader(headersRef);
  if (dateH2) rows.sort((a, b) => compareByDateAndStart(a, b, dateH2, startH2));
  return rows;
}

// --- Rendering -----------------------------------------------------------
function totalPages() { return Math.max(1, Math.ceil(rowsForTeacher.length / pageSize)); }
function getPageSlice() { const start = (currentPage - 1) * pageSize; return rowsForTeacher.slice(start, start + pageSize); }
function updatePagerUI() {
  pageInfo.textContent = `${currentPage} / ${totalPages()}`;
  prevPage.disabled = currentPage <= 1;
  nextPage.disabled = currentPage >= totalPages();
}
function renderOptions(selectEl, options) {
  if (!selectEl) return;
  selectEl.innerHTML = options.map((o) => `<option value="${String(o)}">${String(o)}</option>`).join("");
}

function updateShowAllLabel() {
  if (!showAllBtn) return;
  showAllBtn.textContent = showAll ? "📅 Mostra da oggi" : "📅 Mostra tutto";
}

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
    // aggiorna legenda (vuota)
    buildLegendFromRows([]);
    return;
  }

  const trh = document.createElement("tr");
  headers.forEach((h) => {
    const th = document.createElement("th");
    th.textContent = h;
    trh.appendChild(th);
  });
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
    headers.forEach((h) => {
      const td = document.createElement("td");
      td.textContent = prettyValue(row[h], h);
      tr.appendChild(td);
    });
    frag.appendChild(tr);
  });
  tBody.appendChild(frag);

  renderOptions(dateColumnSelect, ["— nessuna —", ...headers]);
  renderOptions(timeColumnSelect, ["— nessuna —", ...headers]);

  const total = allRows.length; const vis = rowsForTeacher.length; rowsCount.textContent = showAll ? `${vis} lezioni totali` : `${vis} da oggi (${total} totali)`;
  updatePagerUI();

  // Legenda basata sulle RIGHE FILTRATE (visibili nel calendario)
  buildLegendFromRows(allRows);
}

// --- Eventi UI -----------------------------------------------------------
dropZone?.addEventListener("click", () => fileInput?.click());
pickBtn?.addEventListener("click", () => fileInput?.click());
fileInput?.addEventListener("change", (e) => { const f = e.target.files?.[0]; if (f) handleFile(f); });

["dragenter", "dragover"].forEach((ev) =>
  dropZone?.addEventListener(ev, (e) => {
    e.preventDefault(); e.stopPropagation();
    dropZone.classList.add("dragover");
    setStatus("Rilascia per caricare…");
  })
);
["dragleave", "dragend", "drop"].forEach((ev) =>
  dropZone?.addEventListener(ev, (e) => {
    e.preventDefault(); e.stopPropagation();
    dropZone.classList.remove("dragover");
  })
);
dropZone?.addEventListener("drop", (e) => {
  const f = e.dataTransfer?.files?.[0];
  if (f) handleFile(f);
});

// Accessibilità tastiera
dropZone?.addEventListener("keydown", (e) => {
  if (e.key === "Enter" || e.key === " ") { e.preventDefault(); fileInput?.click(); }
});

// Paginazione
prevPage?.addEventListener("click", () => { if (currentPage > 1) { currentPage--; renderTable(); } });
nextPage?.addEventListener("click", () => { if (currentPage < totalPages()) { currentPage++; renderTable(); } });
pageSizeSel?.addEventListener("change", (e) => {
  const v = parseInt(e.target.value, 10);
  pageSize = [25,50,100].includes(v) ? v : 25;
  currentPage = 1; renderTable();
});

// Search nel calendario
searchInput?.addEventListener("input", (e) => {
  searchQuery = String(e.target.value || ""); currentPage = 1; rowsForTeacher = applyFilters(); renderTable();
});
searchInput?.addEventListener("keydown", (e) => {
  if (e.key === "Escape") { searchInput.value = ""; searchQuery = ""; currentPage = 1; rowsForTeacher = applyFilters(); renderTable(); }
});

// Toggle mostra da oggi / tutto
showAllBtn?.addEventListener("click", () => {
  showAll = !showAll;
  rowsForTeacher = applyFilters();
  currentPage = 1;
  renderTable(); // aggiorna anche il testo del bottone
});

// Torna alla Home
backHomeBtn?.addEventListener("click", () => {
  calendarSection.style.display = "none";
  homeSection.style.display = "grid";
  pageTitle.textContent = "Calendario Docenti";
  subLabel.textContent = "Seleziona un docente per vedere le sue lezioni";
  location.hash = "";
  setStatus("File caricato — scegli un docente", "ok");
});

// --- Init ---------------------------------------------------------------
(function init() {
  setStatus("Carica un file Excel (.xlsx)…");

  // Se l'URL contiene #docente=, mostreremo il calendario dopo il caricamento del file
  window.addEventListener("hashchange", () => {
    const m = location.hash.match(/#docente=([^&]+)/);
    if (m && workbook) {
      const name = decodeURIComponent(m[1]);
      openCalendarFor(name);
    }
  });
})();