// Utilities ---------------------------------------------------------------
const $ = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => [...root.querySelectorAll(sel)];
const statusBadge = $("#statusBadge");
const btnBack = $("#btnBack");
const dropZone = $("#dropZone");
const fileInput = $("#fileInput");
const sheetSelect = $("#sheetSelect");
const searchInput = $("#searchInput");
const dateColumnSelect = $("#dateColumnSelect");
const timeColumnSelect = $("#timeColumnSelect");
const table = $("#dataTable");
const tHead = table.tHead || table.createTHead();
const tBody = table.tBodies[0] || table.createTBody();
const rowsCount = $("#rowsCount");
const sortHint = $("#sortHint");
const landing = $("#landing");
const appSection = $("#appSection");
const yearLabel = $("#yearLabel");
const courseSection = $("#courseSection");

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

const COURSE_TO_SHEET = {
  front: "Frot2",
  cyse: "Cyse2",
  dolc: "Dolc2",
  fust: "Fust2",
  ago: "Ago2",
};

const DROP_HEADER_RE = /^(colonna|giorno|fust2)$/i;

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

function escapeHtml(s) {
  return s.replace(
    /[&<>"']/g,
    (m) =>
      ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[
        m
      ])
  );
}

function dateKeyFromVal(v) {
  const d = v instanceof Date ? v : typeof v === "string" ? new Date(v) : null;
  if (d && !isNaN(d)) return d.toISOString().slice(0, 10);
  return String(v);
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

  // Nascondi l'hint di ordinamento
  sortHint?.classList.add("hidden");

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
    logo.textContent = "ITS"; // 👈 mostra ITS nella selezione corso
    logo.classList.remove("active-logo");
  }

  landing.classList.add("hidden");
  appSection.classList.add("hidden");
  courseSection.classList.remove("hidden");
  btnBack.classList.remove("hidden"); // mostra torna indietro quando si entra nella selezione corso


  const isYear2 = selectedYear === "2";
  $("#courseHint").textContent = isYear2
    ? "Scegli un corso per aprire il calendario (foglio pre-selezionato)."
    : "Per l’Anno 1 non è disponibile il calendario: i bottoni non sono attivi.";
  $$("#courseButtons [data-course]").forEach((btn) => {
    btn.disabled = !isYear2;
  });
}

function goToCalendarWithCourse(courseKey) {
  if (selectedYear !== "2") return;
  const wanted = COURSE_TO_SHEET[courseKey];
  pendingSheetName = wanted || null;

  courseSection.classList.add("hidden");
  appSection.classList.remove("hidden");
  btnBack.classList.remove("hidden"); // resta visibile anche nel calendario


  if (
    workbook &&
    pendingSheetName &&
    workbook.SheetNames.includes(pendingSheetName)
  ) {
    sheetSelect.value = pendingSheetName;
    loadSheet(pendingSheetName);
  }

  // Aggiorna logo con il nome del corso
  const logo = document.getElementById("courseLogo");
  if (logo) {
    logo.textContent = courseKey.charAt(0).toUpperCase() + courseKey.slice(1);
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
  localStorage.removeItem("cal-anno");

  landing.classList.remove("hidden");
  appSection.classList.add("hidden");
  courseSection.classList.add("hidden");
  btnBack.classList.add("hidden"); // nascondi quando si torna alla landing


  yearLabel.textContent = "Scegli un anno per iniziare";
  pendingSheetName = null;

  // 👇 resetta anche il logo
  const logo = document.getElementById("courseLogo");
  if (logo) {
    logo.textContent = "ITS";
    logo.classList.remove("active-logo");
  }
});

sheetSelect?.addEventListener("change", (e) => loadSheet(e.target.value));

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
$$("#courseButtons [data-course]").forEach((btn) => {
  btn.addEventListener("click", () => {
    if (selectedYear === "2") {
      goToCalendarWithCourse(btn.dataset.course);
    }
  });
});

// Init --------------------------------------------------------------------
(function init() {
  // Mostra SEMPRE la scelta anno all'avvio
  localStorage.removeItem("cal-anno"); // opzionale ma utile: azzera lo stato salvato

  const logo = document.getElementById("courseLogo");
  if (logo) {
    logo.textContent = "ITS"; // logo iniziale
    logo.classList.remove("active-logo");
  }

  // Stato sezioni: solo landing visibile
  landing.classList.remove("hidden");
  courseSection.classList.add("hidden");
  appSection.classList.add("hidden");
  btnBack.classList.add("hidden"); // all'avvio non visibile
  yearLabel.textContent = "Scegli un anno per iniziare";

  setStatus("Carica un file Excel");
  updateToggleButton();
})();
