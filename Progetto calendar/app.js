// Calendario Docenti — Upload locale per GIACOMAZZI ENRICO
// Carica un XLSX dal computer e mostra SOLO le lezioni del docente richiesto.

// --- Helpers DOM ---------------------------------------------------------
const $ = (sel, root = document) => root.querySelector(sel);
const statusBadge = $("#statusBadge");
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

// Stato
let workbook = null;
let headersRef = [];
let rowsTeacher = [];

// Paginazione
let currentPage = 1;
let pageSize = 25; // default

// Filtri (devono essere globali e definiti prima di init)
let showAll = false;   // default: mostra da OGGI in poi
let allRows = [];      // dataset completo
let searchQuery = "";  // testo di ricerca

// --- Utility -------------------------------------------------------------
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
    .replace(/[<>"']/g, "")
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
function isLikelyTimeHeader(h) {
  return !!h && TIME_HEADER_RE.test(String(h).trim());
}
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
function prettyValue(v, header) {
  if (v instanceof Date) {
    if (isLikelyTimeHeader(header)) return fmtTimeFromDate(v);
    return fmtDateIT(v);
  }
  if (typeof v === "number" && v >= 0 && v < 1) return fmtTimeFromFraction(v);
  return String(v);
}


// --- Post-processing: pulizia colonne -------------------------------
function normalizeHeaderName(h) {
  return String(h || "").trim().toLowerCase();
}

function dropUnwantedColumns(headers, rows) {
  const norm = headers.map(normalizeHeaderName);

  // colonne candidate da rimuovere
  const removeByName = new Set();
  norm.forEach((h, i) => {
    // rimuovi "giorno", "frot2" e la colonna docente (sinonimi)
    if (
      h === "giorno" ||
      h === "frot2" ||
      /^(docente|insegnante|prof|teacher|formatore)$/.test(h)
    ) {
      removeByName.add(headers[i]);
    }
  });

  // rimuovi colonne completamente vuote
  const isColEmpty = (hdr) => rows.every((r) => {
    const v = r[hdr];
    return v === null || v === undefined || String(v).trim() === "";
  });
  headers.forEach((hdr) => {
    if (isColEmpty(hdr)) removeByName.add(hdr);
  });

  // rimuovi colonne chiamate "colonna" SOLO se vuote
  headers.forEach((hdr, i) => {
    if (normalizeHeaderName(hdr) === "colonna" && isColEmpty(hdr)) {
      removeByName.add(hdr);
    }
  });

  // Applica rimozioni
  const keptHeaders = headers.filter((h) => h === 'Corso' || !removeByName.has(h));
  const cleanedRows = rows.map((row) => {
    const o = {};
    keptHeaders.forEach((h) => (o[h] = row[h]));
    return o;
  });
  return { headers: keptHeaders, rows: cleanedRows };
}

// Rileva colonna docente
const TEACHER_HEADER_RE = /^(docente|insegnante|prof|teacher|formatore)$/i;
function autoDetectTeacherHeader(headers) {
  return headers.find((h) => TEACHER_HEADER_RE.test(String(h).trim())) || null;
}


function normalizeCourseName(name) {
  // "Fust A1" -> "Fust1" (spazio opzionale, case-insensitive)
  return String(name || "").replace(/\s*A1\b/i, "1");
}
// --- Caricamento da file locale -----------------------------------------
async function handleFile(file) {
  try {
    setStatus("Leggo il file…");
    const ab = await file.arrayBuffer();
    workbook = XLSX.read(ab, { type: "array" });
    processWorkbook();
  } catch (e) {
    console.error(e);
    setStatus("Errore lettura file", "err");
  }
}

function processWorkbook() {
  const teacher = String(window.TEACHER_NAME || "").trim().toLowerCase();
  const collected = [];
  let headersForRender = null;

  for (const sheetName of workbook.SheetNames) {
    const ws = workbook.Sheets[sheetName];
    const { headers, rows } = excelToJson(ws);
    if (!headers.length || !rows.length) continue;

    const teacherH = autoDetectTeacherHeader(headers);
    if (!teacherH) continue;

    const matches = rows.filter((r) => String(r[teacherH] || "").trim().toLowerCase() === teacher);
    if (!matches.length) continue;

    // Scrivi SUBITO la colonna visibile "Corso" = nome del foglio (normalizzato)
    matches.forEach((r) => (r["Corso"] = normalizeCourseName(sheetName)));

    // Per mobile: se presenti Dalle + Alle, crea colonna "Orario" subito dopo Data
    const isMobile = window.innerWidth <= 520;
    let processed = matches;
    let finalHeaders = headers.slice();
    // Assicura che l'header "Corso" sia presente
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

    if (!headersForRender) headersForRender = finalHeaders.filter(Boolean);
    collected.push(...processed);
  }

  // Pulizia colonne in base alle regole richieste (Corso è già valorizzato)
  const cleaned = dropUnwantedColumns(headersForRender || [], collected);

  headersRef = cleaned.headers;
  allRows = cleaned.rows;
  rowsTeacher = applyFilters();
  currentPage = 1;
  renderTable();

  setStatus(rowsTeacher.length ? "Pronto" : "Nessuna lezione trovata", rowsTeacher.length ? "ok" : "err");
}


function startOfToday() {
  const d = new Date();
  d.setHours(0,0,0,0);
  return d;
}
function rowDate(row, dateHeader) {
  const v = row?.[dateHeader];
  if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  const s = String(v || "").trim();
  const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (m) {
    const dd = parseInt(m[1],10), mm = parseInt(m[2],10)-1, yy = parseInt(m[3],10);
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
  // 1) base set
  let rows = showAll ? allRows.slice() : filterFromToday(allRows);
  // 2) text search over current headers
  rows = rows.filter(r => textMatchRow(r, searchQuery, headersRef));
  return rows;
}

// --- Rendering -----------------------------------------------------------

function totalPages() {
  return Math.max(1, Math.ceil(rowsTeacher.length / pageSize));
}
function getPageSlice() {
  const start = (currentPage - 1) * pageSize;
  return rowsTeacher.slice(start, start + pageSize);
}
function updatePagerUI() {
  pageInfo.textContent = `${currentPage} / ${totalPages()}`;
  prevPage.disabled = currentPage <= 1;
  nextPage.disabled = currentPage >= totalPages();
}

function renderOptions(selectEl, options) {
  if (!selectEl) return;
  selectEl.innerHTML = options.map((o) => `<option value="${String(o)}">${String(o)}</option>`).join("");
}
function renderTable() {
  const headers = headersRef.slice();
  const rows = getPageSlice();

  tHead.innerHTML = "";
  tBody.innerHTML = "";

  if (!headers.length || !rows.length) {
    rowsCount.textContent = "—";
    renderOptions(dateColumnSelect, ["— nessuna —"]);
    renderOptions(timeColumnSelect, ["— nessuna —"]);
    return;
  }

  // Header
  const trh = document.createElement("tr");
  headers.forEach((h) => {
    const th = document.createElement("th");
    th.textContent = h;
    trh.appendChild(th);
  });
  tHead.appendChild(trh);

  // Body
  const frag = document.createDocumentFragment();
  rows.forEach((row) => {
    const tr = document.createElement("tr");
    headers.forEach((h) => {
      const td = document.createElement("td");
      td.textContent = prettyValue(row[h], h);
      tr.appendChild(td);
    });
    frag.appendChild(tr);
  });
  tBody.appendChild(frag);

  // Popola select (usate solo per auto-formattazione data/ora)
  renderOptions(dateColumnSelect, ["— nessuna —", ...headers]);
  renderOptions(timeColumnSelect, ["— nessuna —", ...headers]);

  const total = allRows.length; const vis = rowsTeacher.length; rowsCount.textContent = showAll ? `${vis} lezioni totali` : `${vis} da oggi (${total} totali)`;
  updatePagerUI();
}

// --- Eventi UI -----------------------------------------------------------
dropZone?.addEventListener("click", () => fileInput?.click());
$("#pickBtn")?.addEventListener("click", () => fileInput?.click());

fileInput?.addEventListener("change", (e) => {
  const f = e.target.files?.[0];
  if (f) handleFile(f);
});

// drag & drop
["dragenter", "dragover"].forEach((ev) =>
  dropZone?.addEventListener(ev, (e) => {
    e.preventDefault();
    e.stopPropagation();
    dropZone.classList.add("dragover");
    setStatus("Rilascia per caricare…");
  })
);
["dragleave", "dragend", "drop"].forEach((ev) =>
  dropZone?.addEventListener(ev, (e) => {
    e.preventDefault();
    e.stopPropagation();
    dropZone.classList.remove("dragover");
  })
);
dropZone?.addEventListener("drop", (e) => {
  const f = e.dataTransfer?.files?.[0];
  if (f) handleFile(f);
});

// Accessibilità tastiera
dropZone?.addEventListener("keydown", (e) => {
  if (e.key === "Enter" || e.key === " ") {
    e.preventDefault();
    fileInput?.click();
  }
});

// --- Init ---------------------------------------------------------------
(function init() {
  setStatus("Carica un file Excel (.xlsx)…");
  if (showAllBtn) showAllBtn.textContent = showAll ? "Mostra da oggi" : "Mostra tutto";
})();

// --- Paginazione: eventi -----------------------------------------------
prevPage?.addEventListener("click", () => {
  if (currentPage > 1) { currentPage--; renderTable(); }
});
nextPage?.addEventListener("click", () => {
  if (currentPage < totalPages()) { currentPage++; renderTable(); }
});
pageSizeSel?.addEventListener("change", (e) => {
  const v = parseInt(e.target.value, 10);
  pageSize = [25,50,100].includes(v) ? v : 25;
  currentPage = 1;
  renderTable();
});


// --- Search ---------------------------------------------------------------
searchInput?.addEventListener("input", (e) => {
  searchQuery = String(e.target.value || "");
  currentPage = 1;
  rowsTeacher = applyFilters();
  renderTable();
});
searchInput?.addEventListener("keydown", (e) => {
  if (e.key === "Escape") {
    searchInput.value = "";
    searchQuery = "";
    currentPage = 1;
    rowsTeacher = applyFilters();
    renderTable();
  }
});


// --- Toggle mostra da oggi / tutto --------------------------------------
showAllBtn?.addEventListener("click", () => {
  showAll = !showAll;
  showAllBtn.textContent = showAll ? "Mostra da oggi" : "Mostra tutto";
  rowsTeacher = applyFilters();
  currentPage = 1;
  renderTable();
});
