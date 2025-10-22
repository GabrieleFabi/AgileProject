// Utilities ---------------------------------------------------------------
const $ = (sel, root=document) => root.querySelector(sel);
const $$ = (sel, root=document) => [...root.querySelectorAll(sel)];
const statusBadge = $('#statusBadge');
const dropZone = $('#dropZone');
const fileInput = $('#fileInput');           // ora è globale (in header)
const sheetSelect = $('#sheetSelect');
const searchInput = $('#searchInput');
const dateColumnSelect = $('#dateColumnSelect');
const timeColumnSelect = $('#timeColumnSelect');
const table = $('#dataTable');
const tHead = table.tHead || table.createTHead();
const tBody = table.tBodies[0] || table.createTBody();
const rowsCount = $('#rowsCount');
const sortHint = $('#sortHint');
const landing = $('#landing');
const appSection = $('#appSection');
const yearLabel = $('#yearLabel');
const courseSection = $('#courseSection');
const courseButtonsWrap = $('#courseButtons');
const courseHint = $('#courseHint');

let workbook = null;
let currentData = [];   // array di oggetti {col:val}
let currentHeaders = [];// array di stringhe
let sortState = {key:null, dir:1};
let selectedYear = null; // "1" | "2"
let pendingSheetName = null; // quando scelgo un corso prima di caricare il file

// Mappa corsi -> fogli anno 2
const COURSE_TO_SHEET = {
  front: 'Frot2',
  cyse:  'Cyse2',
  dolc:  'Dolc2',
  fust:  'Fust2',
  ago:   'Ago2'
};

// --- Filtri colonne da nascondere ---
const DROP_HEADER_RE = /^(colonna|giorno|fust2)$/i;

function isEmptyCell(v){
  if (v === null || v === undefined) return true;
  if (v instanceof Date) return false;
  if (typeof v === 'number') return false;
  const s = String(v).trim();
  return s === '' || s === '-' || s === '—';
}

function shouldDropHeader(h, rows){
  if (!h) return true;
  if (DROP_HEADER_RE.test(String(h).trim())) return true;
  return rows.every(r => isEmptyCell(r[h]));
}

function setStatus(text, tone="info"){
  statusBadge.textContent = text;
  const color = tone === 'ok' ? 'var(--accent)' : tone === 'err' ? 'var(--danger)' : 'var(--brand)';
  statusBadge.style.borderColor = 'var(--border)';
  statusBadge.style.boxShadow = 'inset 0 0 0 1px var(--border)';
  statusBadge.style.color = '#fff';
  statusBadge.style.background = `linear-gradient(180deg, ${color}, ${shade(color, -20)})`;
}
function shade(hex, percent){
  if(!hex.startsWith('#')) return hex;
  const num = parseInt(hex.slice(1), 16);
  let r = (num >> 16) + Math.round(255*percent/100);
  let g = (num >> 8 & 0x00FF) + Math.round(255*percent/100);
  let b = (num & 0x0000FF) + Math.round(255*percent/100);
  r=Math.max(Math.min(255,r),0); g=Math.max(Math.min(255,g),0); b=Math.max(Math.min(255,b),0);
  return `#${(b | (g << 8) | (r << 16)).toString(16).padStart(6,'0')}`
}

function excelToJson(ws){
  const aoa = XLSX.utils.sheet_to_json(ws, {header:1, raw:true, defval:""});
  if(!aoa.length){ return {headers:[], rows:[]} }
  let headers = aoa[0].map(h => sanitizeHeader(String(h || 'Colonna')));
  if(headers.every(h => h === 'Colonna')){
    headers = aoa[0].map((_,i)=> colName(i));
  }
  const rows = aoa.slice(1).map(r => {
    const obj = {};
    headers.forEach((h,i)=>{ obj[h] = formatCell(r[i]); });
    return obj;
  });
  return {headers, rows};
}
function sanitizeHeader(h){
  return h.trim().replace(/\s+/g,' ').replace(/[\n\r]+/g,' ').replace(/[<>"']/g,'').slice(0,80);
}
function colName(i){
  let s=""; i++;
  while(i>0){ let m=(i-1)%26; s=String.fromCharCode(65+m)+s; i=Math.floor((i-1)/26) }
  return s;
}
function isExcelDate(v){ return typeof v === 'number' && v > 59 && v < 60000 }
function formatExcelDate(v){
  try{
    const d = XLSX.SSF.parse_date_code(v);
    if(!d) return v;
    const date = new Date(Date.UTC(d.y, (d.m||1)-1, d.d||1, d.H||0, d.M||0, Math.floor(d.S||0)));
    return date;
  }catch(e){ return v }
}
function formatCell(v){
  if(isExcelDate(v)) return formatExcelDate(v);
  return v;
}
function renderOptions(selectEl, options){
  selectEl.innerHTML = options.map(o => `<option value="${String(o)}">${String(o)}</option>`).join('');
}

// Rendering tabella -------------------------------------------------------
function renderTable(headers, rows){
  headers = headers.filter(h => !shouldDropHeader(h, rows));

  // Mostra solo righe dalla data corrente in poi
  const dateHeader = headers.find(h => /^(data|date)$/i.test(String(h).trim()));
  if (dateHeader) {
    const today = new Date(); today.setHours(0,0,0,0);
    rows = rows.filter(r => {
      const v = r[dateHeader];
      const d = v instanceof Date ? v : new Date(v);
      if (isNaN(d)) return false;
      return d >= today;
    });
  }

  currentHeaders = headers; currentData = rows;

  const q = searchInput.value.trim().toLowerCase();
  const filtered = !q ? rows : rows.filter(row => Object.values(row).some(v => toText(v).includes(q)));

  if(sortState.key){
    filtered.sort((a,b)=> cmp(a[sortState.key], b[sortState.key]) * sortState.dir);
  }

  // Header
  tHead.innerHTML = '';
  const trh = document.createElement('tr');
  headers.forEach(h => {
    const th = document.createElement('th');
    th.className = 'sortable';
    th.innerHTML = `<span>${escapeHtml(h)}</span><span class="chev">${sortIcon(h)}</span>`;
    th.addEventListener('click', ()=>{
      if(sortState.key === h){ sortState.dir *= -1 } else { sortState.key = h; sortState.dir = 1 }
      renderTable(headers, rows);
    });
    trh.appendChild(th);
  });
  tHead.appendChild(trh);
  sortHint.classList.toggle('hidden', headers.length === 0);

  // Body
  tBody.innerHTML = '';
  const frag = document.createDocumentFragment();
  filtered.forEach((row)=>{
    const tr = document.createElement('tr');
    headers.forEach(h=>{
      const td = document.createElement('td');
      const v = row[h];
      td.textContent = prettyValue(v, h);
      tr.appendChild(td);
    });
    frag.appendChild(tr);
  });
  tBody.appendChild(frag);

  // Separatore tra date diverse
  const _dateHeaderForSep = autoDetectDateHeader(headers);
  if(_dateHeaderForSep){
    const createdRows2 = [...tBody.querySelectorAll('tr')];
    let prevKey = null;
    createdRows2.forEach((tr, idx)=>{
      const r = filtered[idx];
      const dVal = r[_dateHeaderForSep];
      const dateObj = (dVal instanceof Date) ? dVal : (typeof dVal === 'string' && dVal) ? new Date(dVal) : null;
      const key = dateObj ? dateObj.toISOString().slice(0,10) : String(dVal);
      if(idx>0 && key !== prevKey){
        tr.classList.add('date-sep');
      }else{
        tr.classList.remove('date-sep');
      }
      prevKey = key;
    });
  }

  // Evidenzia record di giornate con copertura < 8h
  const _dateH = autoDetectDateHeader(headers);
  const _startH = autoDetectStartHeader(headers);
  const _endH = autoDetectEndHeader(headers);
  if(_dateH && _startH && _endH){
    const createdRows = [...tBody.querySelectorAll('tr')];
    const map = new Map();
    createdRows.forEach((tr, idx)=>{
      const r = filtered[idx];
      const dVal = r[_dateH];
      const sVal = r[_startH];
      const eVal = r[_endH];
      const dateObj = (dVal instanceof Date) ? dVal : (typeof dVal === 'string' && dVal) ? new Date(dVal) : null;
      const key = dateObj ? dateObj.toISOString().slice(0,10) : String(dVal);
      if(!map.has(key)) map.set(key, {intervals:[], rows:[]});
      map.get(key).rows.push(tr);
      const sMin = toMinutes(sVal);
      const eMin = toMinutes(eVal);
      if(sMin!=null && eMin!=null && eMin>sMin){
        map.get(key).intervals.push({start:sMin, end:eMin});
      }
    });
    for(const [k, grp] of map.entries()){
      const minutes = coveredMinutesWithinNeeds(grp.intervals);
      if(minutes < 480){
        grp.rows.forEach(tr => tr.classList.add('day-short'));
      }else{
        grp.rows.forEach(tr => tr.classList.remove('day-short'));
      }
    }
  }

  rowsCount.textContent = filtered.length ? `${filtered.length} righe visualizzate` : 'Nessun dato da mostrare';

  // Aggiorna select colonna data/ora
  renderOptions(dateColumnSelect, ['— nessuna —', ...headers]);
  renderOptions(timeColumnSelect, ['— nessuna —', ...headers]);
}

function sortIcon(h){
  if(sortState.key !== h) return '↕';
  return sortState.dir === 1 ? '↑' : '↓'
}
function toText(v){
  if(v instanceof Date){
    return v.toISOString().slice(0,10);
  }
  return String(v).toLowerCase();
}
function cmp(a,b){
  if(a instanceof Date && b instanceof Date) return a - b;
  if(a instanceof Date) return -1;
  if(b instanceof Date) return 1;
  const na = Number(a), nb = Number(b);
  const aNum = !Number.isNaN(na); const bNum = !Number.isNaN(nb);
  if(aNum && bNum) return na - nb;
  return String(a).localeCompare(String(b), 'it', {numeric:true, sensitivity:'base'});
}
function escapeHtml(s){ return s.replace(/[&<>"']/g, m=> ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;','\'':'&#39;'}[m]) ) }

// Valore "carino" per celle (formattazione data/ora)
function prettyValue(v, header){
  if(v instanceof Date){
    if(isLikelyTimeHeader(header)) return fmtTimeFromDate(v);
    return fmtDateIT(v);
  }
  if(typeof v === 'number'){
    if(v >= 0 && v < 1) return fmtTimeFromFraction(v);
  }
  return String(v);
}

const TIME_HEADER_RE = /^(ora|orario|dalle|alle|inizio|fine|start|end)$/i;
function isLikelyTimeHeader(h){ return !!h && TIME_HEADER_RE.test(String(h).trim()); }

function fmtTimeFromFraction(fr){
  const total = Math.round(fr * 24 * 60);
  const hh = Math.floor(total / 60);
  const mm = total % 60;
  return String(hh).padStart(2,'0') + ':' + String(mm).padStart(2,'0');
}
function fmtTimeFromDate(d){
  return String(d.getUTCHours()).padStart(2,'0') + ':' + String(d.getUTCMinutes()).padStart(2,'0');
}
const DATE_HEADER_RE = /^(data|date)$/i;
const START_HEADER_RE = /^(dalle|ora ?inizio|inizio|start)$/i;
const END_HEADER_RE   = /^(alle|ora ?fine|fine|end)$/i;

function autoDetectDateHeader(headers){
  const chosen = (dateColumnSelect && dateColumnSelect.value && dateColumnSelect.value !== '— nessuna —')
      ? dateColumnSelect.value : null;
  if (chosen && headers.includes(chosen)) return chosen;
  return headers.find(h => DATE_HEADER_RE.test(String(h))) || null;
}
function autoDetectStartHeader(headers){
  return headers.find(h => START_HEADER_RE.test(String(h))) || null;
}
function autoDetectEndHeader(headers){
  return headers.find(h => END_HEADER_RE.test(String(h))) || null;
}

function toMinutes(v){
  if (v instanceof Date) return v.getUTCHours()*60 + v.getUTCMinutes();
  if (typeof v === 'number'){
    if (v >= 0 && v <= 1) return Math.round(v*24*60);
    if (v > 59 && v < 2400){
      const hh = Math.floor(v/100), mm = Math.round(v%100);
      return hh*60+mm;
    }
  }
  if (typeof v === 'string'){
    const m = v.trim().match(/^(\d{1,2})[:.](\d{2})/);
    if (m){ return parseInt(m[1])*60 + parseInt(m[2]); }
  }
  return null;
}

function mergeIntervals(intervals){
  const arr = intervals.filter(iv => iv && iv.start!=null && iv.end!=null && iv.end>iv.start)
                       .sort((a,b)=> a.start-b.start);
  const merged = [];
  for(const iv of arr){
    if(!merged.length || iv.start > merged[merged.length-1].end){
      merged.push({start:iv.start, end:iv.end});
    } else {
      merged[merged.length-1].end = Math.max(merged[merged.length-1].end, iv.end);
    }
  }
  return merged;
}

function coveredMinutesWithinNeeds(intervals){
  const merged = mergeIntervals(intervals);
  const needs = [{start:9*60, end:13*60}, {start:14*60, end:18*60}];
  let total = 0;
  for(const need of needs){
    for(const iv of merged){
      const start = Math.max(need.start, iv.start);
      const end = Math.min(need.end, iv.end);
      if(end > start) total += (end - start);
    }
  }
  return total;
}

function fmtDateIT(d){
  try{
    return new Intl.DateTimeFormat('it-IT', {weekday:'short', day:'2-digit', month:'2-digit', year:'numeric'}).format(d);
  }catch(e){ return d.toISOString().slice(0,10) }
}

// Import ---------------------------------------------------------
async function handleFile(file){
  if(!file) return;
  if(file.size > 10*1024*1024){ setStatus('File troppo grande (>10MB)','err'); return; }
  setStatus(`Caricamento: ${file.name}…`);
  const data = await file.arrayBuffer();
  workbook = XLSX.read(data, {type:'array'});

  // Popola select fogli
  sheetSelect.innerHTML = '';
  workbook.SheetNames.forEach((name, i)=>{
    const opt = document.createElement('option');
    opt.value = name;
    opt.textContent = `${i+1}. ${name}`;
    sheetSelect.appendChild(opt);
  });

  // Se ho una richiesta specifica da "selezione corso" e il foglio esiste, usa quella.
  let defaultSheet = null;
  if (pendingSheetName && workbook.SheetNames.includes(pendingSheetName)) {
    defaultSheet = pendingSheetName;
  } else {
    // fallback: un foglio che contiene "fust", altrimenti il primo
    defaultSheet = workbook.SheetNames.find(n => n.toLowerCase().includes('fust'))
                   || workbook.SheetNames[0];
  }
  sheetSelect.value = defaultSheet;
  loadSheet(defaultSheet);

  setStatus(`Pronto: ${file.name}`,'ok');

  fileInput.value = ''; // permette di ricaricare anche lo stesso file dopo
}

function loadSheet(name){
  const ws = workbook.Sheets[name];
  const {headers, rows} = excelToJson(ws);
  renderTable(headers, rows);
}

// --- Navigazione: Anno -> Corso -> Calendario ---------------------------
function applyYearChoice(year){
  selectedYear = String(year);
  localStorage.setItem('cal-anno', selectedYear);
  yearLabel.textContent = `Anno ${selectedYear}`;

  landing.classList.add('hidden');
  appSection.classList.add('hidden');
  courseSection.classList.remove('hidden');

  const isYear2 = selectedYear === '2';
  $('#courseHint').textContent = isYear2
    ? 'Scegli un corso per aprire il calendario (foglio pre-selezionato).'
    : 'Per l’Anno 1 non è disponibile il calendario: i bottoni non sono attivi.';
  $$('#courseButtons [data-course]').forEach(btn=>{
    btn.disabled = !isYear2; // anno 1: disabilitati
  });
}

function goToCalendarWithCourse(courseKey){
  if(selectedYear !== '2') return;
  const wanted = COURSE_TO_SHEET[courseKey];
  pendingSheetName = wanted || null;

  courseSection.classList.add('hidden');
  appSection.classList.remove('hidden');

  // Se il file è già stato caricato dalla landing, seleziona subito il foglio
  if(workbook && pendingSheetName && workbook.SheetNames.includes(pendingSheetName)){
    sheetSelect.value = pendingSheetName;
    loadSheet(pendingSheetName);
  }
}

// Eventi UI ---------------------------------------------------------------
fileInput?.addEventListener('change', e=> handleFile(e.target.files[0]));
$('#btnBack')?.addEventListener('click', () => {
  localStorage.removeItem('cal-anno');
  landing.classList.remove('hidden');
  appSection.classList.add('hidden');
  courseSection.classList.add('hidden');
  yearLabel.textContent = 'Scegli un anno per iniziare';
  pendingSheetName = null;
});
sheetSelect?.addEventListener('change', e=> loadSheet(e.target.value));
searchInput?.addEventListener('input', ()=> renderTable(currentHeaders, currentData));
dateColumnSelect?.addEventListener('change', ()=> renderTable(currentHeaders, currentData));
timeColumnSelect?.addEventListener('change', ()=> renderTable(currentHeaders, currentData));

$('#btnClear')?.addEventListener('click', ()=>{
  tHead.innerHTML = '';
  tBody.innerHTML = '';
  rowsCount.textContent = '—';

  currentData = [];
  currentHeaders = [];
  sortState = { key: null, dir: 1 };
  workbook = null;

  sheetSelect.innerHTML = '';
  dateColumnSelect.innerHTML = '';
  timeColumnSelect.innerHTML = '';
  searchInput.value = '';

  fileInput.value = '';

  setStatus('Nessun file');
});

// Bottone landing (nuovo) per caricare l’Excel
$('#btnLoadTop')?.addEventListener('click', ()=>{
  fileInput.value = '';
  fileInput.click();
});

// Bottone già esistente nella pagina calendario: resta invariato
$('#btnSample')?.addEventListener('click', ()=>{
  fileInput.value = '';
  fileInput.click();
});

// Click su pulsanti Anno 1 / Anno 2
$$('.landing [data-anno]').forEach(btn => {
  btn.addEventListener('click', () => applyYearChoice(btn.dataset.anno));
});

// Click su bottoni corso (solo anno 2 abilita)
$$('#courseButtons [data-course]').forEach(btn=>{
  btn.addEventListener('click', ()=>{
    if (selectedYear === '2') {
      goToCalendarWithCourse(btn.dataset.course);
    }
  });
});

// Init ---------------------------------------------------------------
(function init(){
  const saved = localStorage.getItem('cal-anno');
  if(saved === '1' || saved === '2'){
    applyYearChoice(saved);
  } else {
    yearLabel.textContent = 'Scegli un anno per iniziare';
  }
  setStatus('Carica un file Excel');
})();
