/**********************************************************
 * PLANNER DE JORNADAS · APP.JS
 * Mejorado para claridad, experiencia y sincronización
 * automática con la fecha actual
 **********************************************************/

/**********************************************************
 * CONFIG
 **********************************************************/
const TSV_URL =
  "https://docs.google.com/spreadsheets/d/e/2PACX-1vRVAZZw2JvNB1DHJh42Y_6KaNjRmg_MT9gul_heM_XNlj-6Du6aB3E6TihcGQ5DvJHM45xief44MY9u/pub?gid=0&single=true&output=tsv";

const FETCH_TIMEOUT_MS = 15000;

const STORAGE_KEYS = {
  month: "planner_current_month",
  week: "planner_current_week"
};

const DAYS = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"];
const MONTHS = [
  "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
  "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
];

/**********************************************************
 * DOM
 **********************************************************/
const byId = (id) => document.getElementById(id);

/**********************************************************
 * STATE
 **********************************************************/
let allDays = [];
let currentMonth = 0;
let currentWeekIndex = 0;
let monthWeeks = [];
let currentYear = new Date().getFullYear();

/**
 * Índices para búsquedas rápidas
 */
let dayMap = new Map(); // yyyy-mm-dd => day record
let monthBuckets = Array.from({ length: 12 }, () => []);

/**********************************************************
 * HELPERS GENERALES
 **********************************************************/
function tsvToRows(tsv){
  return String(tsv || "")
    .replace(/\r/g, "")
    .split("\n")
    .filter(line => line.trim())
    .map(line => line.split("\t").map(v => (v ?? "").trim()));
}

function parseDMY(str){
  if(!str) return null;

  const parts = String(str).trim().split("/");
  if(parts.length !== 3) return null;

  const [d, m, y] = parts.map(Number);
  if(!d || !m || !y) return null;

  const date = new Date(y, m - 1, d);

  if(
    Number.isNaN(date.getTime()) ||
    date.getFullYear() !== y ||
    date.getMonth() !== (m - 1) ||
    date.getDate() !== d
  ){
    return null;
  }

  return date;
}

function parseTime(str){
  if(str === null || str === undefined) return null;

  const raw = String(str).trim();
  if(!raw || raw === "-") return null;

  const s = raw.toLowerCase();
  const match = s.match(/^(\d{1,2})(?::(\d{1,2}))?\s*(am|pm)?$/i);
  if(!match) return null;

  let h = Number(match[1]);
  let m = Number(match[2] || 0);
  const meridiem = match[3] || "";

  if(Number.isNaN(h) || Number.isNaN(m) || h < 0 || m < 0 || m > 59){
    return null;
  }

  if(meridiem){
    if(h < 1 || h > 12) return null;
    if(meridiem === "pm" && h !== 12) h += 12;
    if(meridiem === "am" && h === 12) h = 0;
  } else if(h > 23){
    return null;
  }

  return (h * 60) + m;
}

function minToHHMM(min){
  if(min === null || min === undefined || Number.isNaN(min)) return "--:--";
  const h = Math.floor(min / 60);
  const m = min % 60;
  return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`;
}

function fmtHours(hours){
  return `${Number(hours || 0).toFixed(1)}h`;
}

function sameDay(a, b){
  return a.getFullYear() === b.getFullYear()
    && a.getMonth() === b.getMonth()
    && a.getDate() === b.getDate();
}

function todayAtLocal(){
  const now = new Date();
  return new Date(now.getFullYear(), now.getMonth(), now.getDate());
}

function isoKeyFromDate(date){
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function startOfWeekMonday(date){
  const d = new Date(date);
  const wd = (d.getDay() + 6) % 7; // lun = 0
  d.setDate(d.getDate() - wd);
  d.setHours(0, 0, 0, 0);
  return d;
}

function endOfWeekSunday(date){
  const s = startOfWeekMonday(date);
  const e = new Date(s);
  e.setDate(e.getDate() + 6);
  e.setHours(23, 59, 59, 999);
  return e;
}

function getWeeksForMonth(year, month){
  const first = new Date(year, month, 1);
  const last = new Date(year, month + 1, 0);

  const weeks = [];
  let cursor = startOfWeekMonday(first);

  while(cursor <= last){
    const start = new Date(cursor);
    const end = endOfWeekSunday(cursor);

    if(end >= first && start <= last){
      weeks.push({ start, end });
    }

    cursor.setDate(cursor.getDate() + 7);
  }

  return weeks;
}

function getWeeksForYear(year){
  const first = new Date(year, 0, 1);
  const last = new Date(year, 11, 31);

  const weeks = [];
  let cursor = startOfWeekMonday(first);

  while(cursor <= last){
    const start = new Date(cursor);
    const end = endOfWeekSunday(cursor);
    weeks.push({ start, end });
    cursor.setDate(cursor.getDate() + 7);
  }

  return weeks;
}

function inRange(date, start, end){
  return date >= start && date <= end;
}

function fmtDM(date){
  const d = String(date.getDate()).padStart(2, "0");
  const m = String(date.getMonth() + 1).padStart(2, "0");
  return `${d}/${m}`;
}

function fmtDMY(date){
  const d = String(date.getDate()).padStart(2, "0");
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const y = date.getFullYear();
  return `${d}/${m}/${y}`;
}

function clamp(n, a, b){
  return Math.max(a, Math.min(b, n));
}

function isMobileLayout(){
  return window.matchMedia("(max-width: 720px)").matches;
}

function safeSetText(id, text){
  const el = byId(id);
  if(el) el.textContent = text;
}

function safeSetHTML(id, html){
  const el = byId(id);
  if(el) el.innerHTML = html;
}

function escapeHTML(str){
  return String(str)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function debounce(fn, wait = 100){
  let t = null;
  return function(...args){
    clearTimeout(t);
    t = setTimeout(() => fn.apply(this, args), wait);
  };
}

function saveViewState(){
  try{
    localStorage.setItem(STORAGE_KEYS.month, String(currentMonth));
    localStorage.setItem(STORAGE_KEYS.week, String(currentWeekIndex));
  } catch(err){
    // Silencioso. La vida ya tiene suficiente drama.
  }
}

function restoreViewState(){
  try{
    const savedMonth = Number(localStorage.getItem(STORAGE_KEYS.month));
    const savedWeek = Number(localStorage.getItem(STORAGE_KEYS.week));

    if(Number.isInteger(savedMonth) && savedMonth >= 0 && savedMonth <= 11){
      currentMonth = savedMonth;
    }

    if(Number.isInteger(savedWeek) && savedWeek >= 0){
      currentWeekIndex = savedWeek;
    }
  } catch(err){
    // Ignorar si storage falla
  }
}

function showFatalError(message){
  const safeMessage = escapeHTML(message);

  ["calendarGrid", "totalsGrid", "lunchDaysList"].forEach(id => {
    const el = byId(id);
    if(el){
      el.innerHTML = `<div class="sub" style="padding:8px 0;">${safeMessage}</div>`;
    }
  });

  safeSetText("weekLabel", "No fue posible cargar los datos");
  safeSetText("monthLabel", `Planner ${currentYear}`);

  const kpiIds = [
    "kpiTopDay", "kpiTopDayHint", "kpiTopWeek", "kpiTopWeekHint",
    "kpiMonthTotal", "kpiMonthTotalHint", "kpiWeekAvg", "kpiWeekAvgHint",
    "kpiTopWeekday", "kpiTopWeekdayHint", "kpiDaysWithJornada", "kpiDaysWithJornadaHint",
    "kpiLunchDays", "kpiLunchDaysHint", "kpiLunchHours", "kpiLunchHoursHint",
    "kpiRawTotal", "kpiRawTotalHint",
    "kpiYearTotal", "kpiYearTotalHint", "kpiYearMonthAvg", "kpiYearMonthAvgHint",
    "kpiTopMonth", "kpiTopMonthHint", "kpiTopWeekYear", "kpiTopWeekYearHint",
    "kpiTopWeekdayYear", "kpiTopWeekdayYearHint", "kpiLunchHoursYear", "kpiLunchHoursYearHint",
    "kpiDaysWithJornadaYear", "kpiDaysWithJornadaYearHint",
    "kpiDaysWithoutJornadaYear", "kpiDaysWithoutJornadaYearHint",
    "kpiYearRawTotal", "kpiYearRawTotalHint"
  ];

  for(const id of kpiIds){
    safeSetText(id, id.toLowerCase().includes("hint") ? "No disponible" : "--");
  }
}

function setLoadingState(){
  safeSetText("monthLabel", "Cargando...");
  safeSetText("weekLabel", "Cargando semana...");

  const loadingText = `<div class="sub" style="padding:8px 0;">Cargando información...</div>`;

  safeSetHTML("calendarGrid", loadingText);
  safeSetHTML("totalsGrid", loadingText);
  safeSetHTML("lunchDaysList", loadingText);
}

/**********************************************************
 * REGLA DE ALMUERZO
 * Si rawHours > 6h, se descuenta 1h
 **********************************************************/
function lunchDeduction(rawHours){
  return rawHours > 6 ? 1 : 0;
}

function effectiveHours(rawHours){
  const lunch = lunchDeduction(rawHours);
  return Math.max(0, rawHours - lunch);
}

/**********************************************************
 * FETCH ROBUSTO
 **********************************************************/
async function fetchTextWithTimeout(url, timeoutMs = FETCH_TIMEOUT_MS){
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);

  try{
    const res = await fetch(url, {
      method: "GET",
      cache: "no-store",
      signal: controller.signal
    });

    if(!res.ok){
      throw new Error(`HTTP ${res.status} al cargar el TSV`);
    }

    return await res.text();
  } finally {
    clearTimeout(timer);
  }
}

/**********************************************************
 * NORMALIZACIÓN / ÍNDICES
 **********************************************************/
function rebuildIndexes(){
  dayMap = new Map();
  monthBuckets = Array.from({ length: 12 }, () => []);

  for(const day of allDays){
    dayMap.set(isoKeyFromDate(day.date), day);

    if(day.m >= 0 && day.m < 12){
      monthBuckets[day.m].push(day);
    }
  }

  for(let i = 0; i < 12; i++){
    monthBuckets[i].sort((a, b) => a.date - b.date);
  }
}

function getDayRecord(date){
  return dayMap.get(isoKeyFromDate(date)) || null;
}

function getMonthData(month){
  return monthBuckets[month] || [];
}

function getAvailableYears(){
  const years = [...new Set(allDays.map(d => d.y))].sort((a, b) => a - b);
  return years;
}

function resolveInitialYear(){
  const availableYears = getAvailableYears();
  if(!availableYears.length){
    return new Date().getFullYear();
  }

  const todayYear = new Date().getFullYear();
  return availableYears.includes(todayYear) ? todayYear : availableYears[0];
}

function syncToCurrentDate(forceCurrentDate = true){
  const today = todayAtLocal();

  if(forceCurrentDate){
    if(today.getFullYear() === currentYear){
      currentMonth = today.getMonth();
      const weeks = getWeeksForMonth(currentYear, currentMonth);
      const index = weeks.findIndex(w => today >= w.start && today <= w.end);
      currentWeekIndex = index >= 0 ? index : 0;
      return;
    }

    const monthsWithData = monthBuckets
      .map((bucket, idx) => ({ idx, count: bucket.length }))
      .filter(item => item.count > 0);

    currentMonth = monthsWithData.length ? monthsWithData[0].idx : 0;
    currentWeekIndex = 0;
    return;
  }

  restoreViewState();
  currentMonth = clamp(currentMonth, 0, 11);

  const weeks = getWeeksForMonth(currentYear, currentMonth);
  currentWeekIndex = clamp(currentWeekIndex, 0, Math.max(0, weeks.length - 1));
}

/**********************************************************
 * LOAD DATA
 **********************************************************/
async function load(){
  setLoadingState();

  const tsv = await fetchTextWithTimeout(`${TSV_URL}&t=${Date.now()}`);
  const rows = tsvToRows(tsv);

  if(!rows.length){
    throw new Error("La fuente de datos llegó vacía");
  }

  const headerCell = (rows[0]?.[0] || "").toLowerCase();
  const startRowIndex = headerCell.includes("día") || headerCell.includes("dia") ? 1 : 0;

  const parsed = [];

  for(let i = startRowIndex; i < rows.length; i++){
    const r = rows[i];

    // Asumido:
    // r[1] = fecha
    // r[2] = hora inicio
    // r[3] = hora fin
    // r[4]/r[5] = nota
    const date = parseDMY(r[1]);
    if(!date) continue;

    const startMin = parseTime(r[2]);
    const endMin = parseTime(r[3]);
    const nota = r[5] || r[4] || "";

    const hasJornada = startMin !== null && endMin !== null && endMin > startMin;
    const rawHours = hasJornada ? Math.max(0, (endMin - startMin) / 60) : 0;
    const lunchHours = hasJornada ? lunchDeduction(rawHours) : 0;
    const hours = hasJornada ? effectiveHours(rawHours) : 0;

    const timeLabel = hasJornada
      ? `${minToHHMM(startMin)} – ${minToHHMM(endMin)}`
      : "";

    parsed.push({
      date,
      y: date.getFullYear(),
      m: date.getMonth(),
      d: date.getDate(),
      weekday: (date.getDay() + 6) % 7, // lunes = 0

      hasJornada,
      startMin,
      endMin,

      rawHours,
      lunchHours,
      hours,

      note: nota,
      label: hasJornada ? timeLabel : (nota || "Sin jornada")
    });
  }

  allDays = parsed.sort((a, b) => a.date - b.date);
  rebuildIndexes();

  if(!allDays.length){
    throw new Error("No hay registros válidos en la fuente de datos.");
  }

  currentYear = resolveInitialYear();

  // Filtramos solo el año activo para los buckets mensuales
  const filtered = allDays.filter(d => d.y === currentYear);
  allDays = filtered.sort((a, b) => a.date - b.date);
  rebuildIndexes();

  if(!allDays.length){
    throw new Error(`No hay registros válidos para ${currentYear}`);
  }

  syncToCurrentDate(true);
  render();
}

/**********************************************************
 * DATA HELPERS POR MES / AÑO
 **********************************************************/
function getMonthJornadaDays(month = currentMonth){
  return getMonthData(month).filter(d => d.hasJornada);
}

function getMonthLunchDays(month = currentMonth){
  return getMonthJornadaDays(month).filter(d => d.lunchHours > 0);
}

function getMonthWeekTotals(month = currentMonth){
  const weeks = getWeeksForMonth(currentYear, month);
  const monthData = getMonthJornadaDays(month);

  return weeks.map(w => {
    let sum = 0;
    for(const d of monthData){
      if(inRange(d.date, w.start, w.end)) sum += d.hours;
    }
    return sum;
  });
}

function getMonthWeekdayTotals(month = currentMonth){
  const totals = [0, 0, 0, 0, 0, 0, 0];
  const monthData = getMonthJornadaDays(month);

  for(const d of monthData){
    totals[d.weekday] += d.hours;
  }

  return totals;
}

function getYearJornadaDays(){
  return allDays.filter(d => d.y === currentYear && d.hasJornada);
}

function getYearMonthTotals(){
  const totals = Array(12).fill(0);
  for(const d of getYearJornadaDays()){
    totals[d.m] += d.hours;
  }
  return totals;
}

function getYearMonthRawTotals(){
  const totals = Array(12).fill(0);
  for(const d of getYearJornadaDays()){
    totals[d.m] += d.rawHours;
  }
  return totals;
}

function getYearWeekdayTotals(){
  const totals = [0, 0, 0, 0, 0, 0, 0];
  for(const d of getYearJornadaDays()){
    totals[d.weekday] += d.hours;
  }
  return totals;
}

/**********************************************************
 * UI HELPERS
 **********************************************************/
function getCompactCellText(dayRecord){
  if(!dayRecord) return "Sin jornada";

  if(!dayRecord.hasJornada){
    return dayRecord.label || "Sin jornada";
  }

  const lunchMark = dayRecord.lunchHours ? " 🍽️" : "";

  if(isMobileLayout()){
    return `${fmtHours(dayRecord.hours)} · ${minToHHMM(dayRecord.startMin)}-${minToHHMM(dayRecord.endMin)}${lunchMark}`;
  }

  return `${fmtHours(dayRecord.hours)} · ${dayRecord.label}${lunchMark}`;
}

function getCellTitle(dayRecord){
  if(!dayRecord) return "Sin jornada";

  if(!dayRecord.hasJornada){
    return dayRecord.label || "Sin jornada";
  }

  const lunchText = dayRecord.lunchHours
    ? ` · incluye ajuste de ${fmtHours(dayRecord.lunchHours)} por almuerzo`
    : "";

  return `${fmtDMY(dayRecord.date)} · ${fmtHours(dayRecord.hours)} · ${dayRecord.label}${lunchText}`;
}

function setWeekNavState(){
  const prevWeek = byId("prevWeek");
  const nextWeek = byId("nextWeek");

  if(prevWeek){
    prevWeek.disabled = currentWeekIndex <= 0;
    prevWeek.setAttribute("aria-disabled", prevWeek.disabled ? "true" : "false");
  }

  if(nextWeek){
    nextWeek.disabled = currentWeekIndex >= (monthWeeks.length - 1);
    nextWeek.setAttribute("aria-disabled", nextWeek.disabled ? "true" : "false");
  }
}

function createEmptyMessage(text){
  const el = document.createElement("div");
  el.className = "sub";
  el.style.padding = "8px 0";
  el.textContent = text;
  return el;
}

/**********************************************************
 * RENDER MASTER
 **********************************************************/
function render(){
  monthWeeks = getWeeksForMonth(currentYear, currentMonth);
  currentWeekIndex = clamp(currentWeekIndex, 0, Math.max(0, monthWeeks.length - 1));

  safeSetText("monthLabel", `${MONTHS[currentMonth]} ${currentYear}`);
  setWeekNavState();

  renderCalendar();
  renderWeekBars();
  renderTotals();
  renderKPIs();
  renderYearKPIs();

  saveViewState();
}

/**********************************************************
 * CALENDAR GRID
 **********************************************************/
function renderCalendar(){
  const grid = byId("calendarGrid");
  if(!grid) return;

  grid.innerHTML = "";

  const first = new Date(currentYear, currentMonth, 1);
  const offset = (first.getDay() + 6) % 7;
  const daysInMonth = new Date(currentYear, currentMonth + 1, 0).getDate();
  const totalCells = 42;
  const today = todayAtLocal();

  const frag = document.createDocumentFragment();

  for(let i = 0; i < totalCells; i++){
    const cell = document.createElement("div");
    cell.className = "calCell";

    const dayNumber = i - offset + 1;

    if(dayNumber < 1 || dayNumber > daysInMonth){
      cell.classList.add("off");

      const top = document.createElement("div");
      top.className = "calDate";
      top.innerHTML = "&nbsp;";

      const bottom = document.createElement("div");
      bottom.className = "calHours";
      bottom.innerHTML = "&nbsp;";

      cell.appendChild(top);
      cell.appendChild(bottom);
      frag.appendChild(cell);
      continue;
    }

    const date = new Date(currentYear, currentMonth, dayNumber);
    const data = getDayRecord(date);

    const top = document.createElement("div");
    top.className = "calDate";
    top.textContent = dayNumber;

    const bottom = document.createElement("div");
    bottom.className = "calHours";
    bottom.textContent = getCompactCellText(data);

    if(data && data.hasJornada){
      cell.classList.add("on");
    }

    if(sameDay(date, today)){
      cell.classList.add("isToday");
    }

    const ariaText = data && data.hasJornada
      ? `${fmtDMY(date)} · ${fmtHours(data.hours)} · ${data.label}`
      : `${fmtDMY(date)} · ${data ? data.label : "Sin jornada"}`;

    cell.title = getCellTitle(data);
    cell.setAttribute("aria-label", ariaText);

    cell.appendChild(top);
    cell.appendChild(bottom);
    frag.appendChild(cell);
  }

  grid.appendChild(frag);
}

/**********************************************************
 * WEEK BARS
 **********************************************************/
function renderWeekBars(){
  const week = monthWeeks[currentWeekIndex] || null;
  const map = [
    ["lun", 0], ["mar", 1], ["mie", 2], ["jue", 3],
    ["vie", 4], ["sab", 5], ["dom", 6]
  ];

  if(!week){
    safeSetText("weekLabel", "Semana · Sin datos");

    for(const [id] of map){
      const bar = byId(`bar-${id}`);
      const h = byId(`hours-${id}`);
      if(bar) bar.innerHTML = "";
      if(h) h.textContent = "0.0h";
    }
    return;
  }

  safeSetText(
    "weekLabel",
    `Semana ${currentWeekIndex + 1} · Del ${fmtDM(week.start)} al ${fmtDM(week.end)}`
  );

  const totals = [0, 0, 0, 0, 0, 0, 0];
  const monthData = getMonthJornadaDays(currentMonth);

  for(const d of monthData){
    if(inRange(d.date, week.start, week.end)){
      totals[d.weekday] += d.hours;
    }
  }

  const max = Math.max(...totals, 1);

  for(const [id, i] of map){
    const bar = byId(`bar-${id}`);
    const h = byId(`hours-${id}`);
    if(!bar || !h) continue;

    bar.innerHTML = "";

    const fill = document.createElement("div");
    fill.className = "barFill";
    fill.style.height = `${(totals[i] / max) * 100}%`;
    fill.title = `${DAYS[i]} · ${fmtHours(totals[i])}`;
    fill.setAttribute("aria-label", `${DAYS[i]} ${fmtHours(totals[i])}`);

    bar.appendChild(fill);
    h.textContent = fmtHours(totals[i]);
  }
}

/**********************************************************
 * TOTALS
 **********************************************************/
function renderTotals(){
  const totalsGrid = byId("totalsGrid");
  if(!totalsGrid) return;

  const weekdayTotals = getMonthWeekdayTotals(currentMonth);
  const weekTotals = getMonthWeekTotals(currentMonth);
  const monthTotal = weekTotals.reduce((a, b) => a + b, 0);

  totalsGrid.innerHTML = "";

  if(!weekTotals.length){
    totalsGrid.appendChild(createEmptyMessage("No hay datos disponibles para este mes."));
    return;
  }

  const box1 = document.createElement("div");
  box1.className = "totalsBox";
  box1.innerHTML = `
    <div class="totalsTitle">Totales por día del mes</div>
    <div class="totalsList">
      ${DAYS.map((d, i) => `
        <div class="totalsRow">
          <span>${d}</span>
          <span>${fmtHours(weekdayTotals[i])}</span>
        </div>
      `).join("")}
    </div>
  `;

  const box2 = document.createElement("div");
  box2.className = "totalsBox";
  box2.innerHTML = `
    <div class="totalsTitle">Totales por semana del mes</div>
    <div class="totalsList">
      ${weekTotals.map((hours, i) => `
        <div class="totalsRow">
          <span>Semana ${i + 1}</span>
          <span>${fmtHours(hours)}</span>
        </div>
      `).join("")}
      <div class="totalsRow" style="margin-top:8px; font-weight:950; color:rgba(15,23,42,.88);">
        <span>Total del mes</span>
        <span>${fmtHours(monthTotal)}</span>
      </div>
    </div>
  `;

  totalsGrid.appendChild(box1);
  totalsGrid.appendChild(box2);
}

/**********************************************************
 * KPIs MENSUALES
 **********************************************************/
function renderKPIs(){
  const monthData = getMonthData(currentMonth);
  const jornadaDays = monthData.filter(d => d.hasJornada);
  const lunchDays = jornadaDays.filter(d => d.lunchHours > 0);

  const rawTotal = jornadaDays.reduce((a, d) => a + d.rawHours, 0);
  const effectiveTotal = jornadaDays.reduce((a, d) => a + d.hours, 0);
  const lunchHoursTotal = lunchDays.reduce((a, d) => a + d.lunchHours, 0);

  let topDay = null;
  for(const d of jornadaDays){
    if(!topDay || d.hours > topDay.hours){
      topDay = d;
    }
  }

  const weekTotals = getMonthWeekTotals(currentMonth);
  let topWeekIndex = 0;
  for(let i = 1; i < weekTotals.length; i++){
    if(weekTotals[i] > weekTotals[topWeekIndex]) topWeekIndex = i;
  }

  const weekAvg = weekTotals.length
    ? weekTotals.reduce((a, b) => a + b, 0) / weekTotals.length
    : 0;

  const weekdayTotals = getMonthWeekdayTotals(currentMonth);
  let topWeekday = 0;
  for(let i = 1; i < 7; i++){
    if(weekdayTotals[i] > weekdayTotals[topWeekday]) topWeekday = i;
  }

  if(topDay){
    safeSetText("kpiTopDay", fmtDM(topDay.date));
    safeSetText(
      "kpiTopDayHint",
      `${fmtHours(topDay.hours)} · ${topDay.label}${topDay.lunchHours ? " 🍽️" : ""}`
    );
  } else {
    safeSetText("kpiTopDay", "--");
    safeSetText("kpiTopDayHint", "No hay jornadas registradas en este mes");
  }

  if(weekTotals.length){
    safeSetText("kpiTopWeek", `Semana ${topWeekIndex + 1}`);

    if(monthWeeks[topWeekIndex]){
      safeSetText(
        "kpiTopWeekHint",
        `${fmtHours(weekTotals[topWeekIndex] || 0)} · Del ${fmtDM(monthWeeks[topWeekIndex].start)} al ${fmtDM(monthWeeks[topWeekIndex].end)}`
      );
    } else {
      safeSetText("kpiTopWeekHint", fmtHours(weekTotals[topWeekIndex] || 0));
    }
  } else {
    safeSetText("kpiTopWeek", "--");
    safeSetText("kpiTopWeekHint", "No hay semanas con horas registradas");
  }

  safeSetText("kpiMonthTotal", fmtHours(effectiveTotal));
  safeSetText("kpiMonthTotalHint", "Tiempo acumulado del mes");

  safeSetText("kpiWeekAvg", fmtHours(weekAvg));
  safeSetText("kpiWeekAvgHint", "Promedio semanal del mes");

  safeSetText("kpiTopWeekday", DAYS[topWeekday]);
  safeSetText("kpiTopWeekdayHint", `${fmtHours(weekdayTotals[topWeekday])} acumuladas`);

  safeSetText("kpiDaysWithJornada", String(jornadaDays.length));
  safeSetText("kpiDaysWithJornadaHint", `Días con jornada en ${MONTHS[currentMonth]}`);

  safeSetText("kpiLunchDays", String(lunchDays.length));
  safeSetText("kpiLunchDaysHint", "Días con ajuste de almuerzo");

  safeSetText("kpiLunchHours", fmtHours(lunchHoursTotal));
  safeSetText("kpiLunchHoursHint", "Tiempo total ajustado por almuerzo");

  safeSetText("kpiRawTotal", fmtHours(rawTotal));
  safeSetText("kpiRawTotalHint", "Referencia antes de ajustes");

  renderLunchDaysList(lunchDays);
}

/**********************************************************
 * KPIs ANUALES
 **********************************************************/
function renderYearKPIs(){
  if(!byId("kpiYearTotal")) return;

  const jornadaYear = getYearJornadaDays();

  const effectiveYearTotal = jornadaYear.reduce((a, d) => a + d.hours, 0);
  const rawYearTotal = jornadaYear.reduce((a, d) => a + d.rawHours, 0);
  const lunchHoursYear = jornadaYear.reduce((a, d) => a + d.lunchHours, 0);

  const monthTotals = getYearMonthTotals();
  const monthRawTotals = getYearMonthRawTotals();
  const weekdayTotalsYear = getYearWeekdayTotals();

  const monthAvg = monthTotals.reduce((a, b) => a + b, 0) / 12;

  let topMonth = 0;
  for(let i = 1; i < 12; i++){
    if(monthTotals[i] > monthTotals[topMonth]) topMonth = i;
  }

  let topWeekdayYear = 0;
  for(let i = 1; i < 7; i++){
    if(weekdayTotalsYear[i] > weekdayTotalsYear[topWeekdayYear]) topWeekdayYear = i;
  }

  const yearWeeks = getWeeksForYear(currentYear);
  const yearWeekTotals = yearWeeks.map(w => {
    let sum = 0;
    for(const d of jornadaYear){
      if(inRange(d.date, w.start, w.end)) sum += d.hours;
    }
    return sum;
  });

  let topWeekYearIndex = 0;
  for(let i = 1; i < yearWeekTotals.length; i++){
    if(yearWeekTotals[i] > yearWeekTotals[topWeekYearIndex]) topWeekYearIndex = i;
  }

  const topWeekObj = yearWeeks[topWeekYearIndex];

  let daysWithJornadaYear = 0;
  let daysWithoutJornadaYear = 0;

  const yearStart = new Date(currentYear, 0, 1);
  const yearEnd = new Date(currentYear, 11, 31);

  for(let dt = new Date(yearStart); dt <= yearEnd; dt.setDate(dt.getDate() + 1)){
    const rec = getDayRecord(new Date(dt));
    if(rec && rec.hasJornada) daysWithJornadaYear++;
    else daysWithoutJornadaYear++;
  }

  safeSetText("kpiYearTotal", fmtHours(effectiveYearTotal));
  safeSetText("kpiYearTotalHint", "Tiempo acumulado del año");

  safeSetText("kpiYearMonthAvg", fmtHours(monthAvg));
  safeSetText("kpiYearMonthAvgHint", "Promedio mensual del año");

  safeSetText("kpiTopMonth", MONTHS[topMonth]);
  safeSetText(
    "kpiTopMonthHint",
    `${fmtHours(monthTotals[topMonth])} acumuladas · ${fmtHours(monthRawTotals[topMonth])} antes de ajustes`
  );

  safeSetText("kpiTopWeekYear", `Semana ${topWeekYearIndex + 1}`);
  safeSetText(
    "kpiTopWeekYearHint",
    topWeekObj
      ? `${fmtHours(yearWeekTotals[topWeekYearIndex])} · Del ${fmtDM(topWeekObj.start)} al ${fmtDM(topWeekObj.end)}`
      : fmtHours(yearWeekTotals[topWeekYearIndex] || 0)
  );

  safeSetText("kpiTopWeekdayYear", DAYS[topWeekdayYear]);
  safeSetText("kpiTopWeekdayYearHint", `${fmtHours(weekdayTotalsYear[topWeekdayYear])} acumuladas`);

  safeSetText("kpiLunchHoursYear", fmtHours(lunchHoursYear));
  safeSetText("kpiLunchHoursYearHint", "Tiempo total ajustado por almuerzo");

  safeSetText("kpiDaysWithJornadaYear", String(daysWithJornadaYear));
  safeSetText("kpiDaysWithJornadaYearHint", "Días con jornada durante el año");

  safeSetText("kpiDaysWithoutJornadaYear", String(daysWithoutJornadaYear));
  safeSetText("kpiDaysWithoutJornadaYearHint", "Días sin jornada o sin registro");

  safeSetText("kpiYearRawTotal", fmtHours(rawYearTotal));
  safeSetText("kpiYearRawTotalHint", "Referencia anual antes de ajustes");
}

/**********************************************************
 * LUNCH DAYS LIST
 **********************************************************/
function renderLunchDaysList(lunchDays){
  const list = byId("lunchDaysList");
  if(!list) return;

  list.innerHTML = "";

  if(!lunchDays.length){
    list.appendChild(
      createEmptyMessage("Este mes no tiene días con ajuste de almuerzo.")
    );
    return;
  }

  const sorted = [...lunchDays].sort((a, b) => a.date - b.date);
  const frag = document.createDocumentFragment();

  for(const d of sorted){
    const chip = document.createElement("div");
    chip.className = "chip chip--on";
    chip.textContent = `${DAYS[d.weekday]} ${fmtDM(d.date)} · ${fmtHours(d.hours)} (−${fmtHours(d.lunchHours)} 🍽️)`;
    chip.title = `${fmtDMY(d.date)} · ${d.label} · ${fmtHours(d.hours)}`;
    frag.appendChild(chip);
  }

  list.appendChild(frag);
}

/**********************************************************
 * NAVEGACIÓN
 **********************************************************/
function goToMonth(month){
  currentMonth = clamp(month, 0, 11);
  currentWeekIndex = 0;
  render();
}

function goToWeek(index){
  currentWeekIndex = clamp(index, 0, Math.max(0, monthWeeks.length - 1));
  render();
}

function wireNav(){
  const prevMonth = byId("prevMonth");
  const nextMonth = byId("nextMonth");
  const prevWeek = byId("prevWeek");
  const nextWeek = byId("nextWeek");

  if(prevMonth){
    prevMonth.addEventListener("click", () => {
      goToMonth((currentMonth + 11) % 12);
    });
  }

  if(nextMonth){
    nextMonth.addEventListener("click", () => {
      goToMonth((currentMonth + 1) % 12);
    });
  }

  if(prevWeek){
    prevWeek.addEventListener("click", () => {
      goToWeek(currentWeekIndex - 1);
    });
  }

  if(nextWeek){
    nextWeek.addEventListener("click", () => {
      goToWeek(currentWeekIndex + 1);
    });
  }

  window.addEventListener("resize", debounce(() => {
    renderCalendar();
  }, 120));
}

/**********************************************************
 * START
 **********************************************************/
async function start(){
  wireNav();

  try{
    await load();
  } catch(err){
    console.error(err);
    showFatalError("No se pudieron cargar los datos. Revisa la URL publicada del Sheet o sus permisos.");
  }
}

start();