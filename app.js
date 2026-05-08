/* ──────────────────────────────────────────────────────────
   FLOORPLAN · Production Planning Console
   App logic: Excel I/O, validation, KPIs, scheduling, export
   ────────────────────────────────────────────────────────── */
(() => {
'use strict';

/* ============================================================
   1. SCHEMA — required columns per sheet
   ============================================================ */
const SCHEMA = {
  WorkOrders: {
    required: ['OrderNumber','Material','OrderQty','ConfirmedQty','WorkCenter','StartDate','FinishDate','Status'],
    optional: ['Description','Priority']
  },
  Materials: {
    required: ['Material','OnHand','Reserved','SafetyStock','ReorderPoint'],
    optional: ['Description','LeadTimeDays','UoM']
  },
  BOM: {
    required: ['ParentMaterial','Component','QtyPerParent'],
    optional: ['UoM']
  },
  WorkCenters: {
    required: ['WorkCenter','HoursPerDay','DaysPerWeek'],
    optional: ['Description','Efficiency','HoursPerUnit']
  }
};

const SETTINGS = {
  riskWindowDays: 5,
  pastDueGraceDays: 0,
  lowStockMultiplier: 1.20,
  capacityWarn: 0.85,
  capacityCrit: 1.00,
  defaultHoursPerUnit: 1.0,
  defaultEfficiency: 0.90
};

/* ============================================================
   2. UTILITIES
   ============================================================ */
const $ = id => document.getElementById(id);

const normKey = k => String(k || '').replace(/[\s_\-]+/g,'').toLowerCase();

function toast(msg, isErr=false){
  const t = $('toast');
  t.textContent = msg;
  t.classList.toggle('toast--err', isErr);
  t.classList.add('is-show');
  clearTimeout(t._timer);
  t._timer = setTimeout(()=>t.classList.remove('is-show'), 2800);
}

function setStatus(state, text){
  const pill = $('statusPill');
  pill.dataset.state = state;
  $('statusText').textContent = text;
}

function fmtNum(n, decimals=0){
  if (n === null || n === undefined || isNaN(n)) return '—';
  return Number(n).toLocaleString(undefined,{minimumFractionDigits:decimals,maximumFractionDigits:decimals});
}
function fmtPct(n, decimals=0){
  if (n === null || n === undefined || isNaN(n)) return '—';
  return (n*100).toFixed(decimals) + '%';
}

function parseDate(v){
  if (v === null || v === undefined || v === '') return null;
  if (v instanceof Date && !isNaN(v)) return v;
  // Excel serial date number
  if (typeof v === 'number'){
    // SheetJS already converts dates if cellDates:true, but defensively:
    const epoch = new Date(Date.UTC(1899,11,30));
    return new Date(epoch.getTime() + v*86400000);
  }
  const d = new Date(v);
  return isNaN(d) ? null : d;
}
function fmtDate(d){
  if (!d) return '—';
  return d.toISOString().slice(0,10);
}
function todayStartOfDay(){
  const d = new Date();
  d.setHours(0,0,0,0);
  return d;
}
function daysBetween(a,b){
  if (!a || !b) return null;
  const ms = b.getTime() - a.getTime();
  return Math.round(ms / 86400000);
}

/* Map raw row keys (case/space-insensitive) to canonical schema keys */
function canonicalizeRow(row, required, optional){
  const all = [...required, ...optional];
  const lookup = {};
  for (const canon of all) lookup[normKey(canon)] = canon;
  const out = {};
  for (const k of Object.keys(row)){
    const c = lookup[normKey(k)];
    if (c) out[c] = row[k];
    else out['_extra_' + k] = row[k]; // preserve extras for export
  }
  return out;
}

/* ============================================================
   3. WORKBOOK INTAKE & VALIDATION
   ============================================================ */
let STATE = {
  fileName: null,
  raw: null,        // canonicalized rows per sheet
  warnings: [],
  computed: null,   // { kpis, schedule, shortages, capacity, inventory }
  edits: new Map(), // key: "Sheet:idx:field" → original value (for revert)
  confirmations: [] // shop floor confirmation log (from operator screen)
};

/* Persist to localStorage so floor.html can read and write back */
function persistShared(){
  if (!STATE.raw) return;
  const Shared = window.FloorplanShared;
  if (!Shared) return;
  Shared.save({
    raw: STATE.raw,
    confirmations: STATE.confirmations,
    fileName: STATE.fileName,
    savedAt: new Date().toISOString()
  });
}

/* Load any existing shared state on page load (operator may have already
   uploaded data via the planner in a previous session). */
function tryLoadShared(){
  const Shared = window.FloorplanShared;
  if (!Shared) return false;
  const s = Shared.load();
  if (!s || !s.raw) return false;
  STATE.raw = s.raw;
  STATE.confirmations = s.confirmations || [];
  STATE.fileName = s.fileName || 'restored';
  STATE.warnings = ['Restored from previous session — re-upload to refresh.'];
  return true;
}

/* Edit tracking helpers ------------------------------------- */
function editKey(sheet, idx, field){ return `${sheet}:${idx}:${field}`; }
function isEdited(sheet, idx, field){ return STATE.edits.has(editKey(sheet, idx, field)); }
function recordEdit(sheet, idx, field, originalValue){
  const key = editKey(sheet, idx, field);
  // Only record the FIRST time this cell is edited so reset returns to the original
  if (!STATE.edits.has(key)) STATE.edits.set(key, originalValue);
}
function clearEdits(){ STATE.edits.clear(); }
function editCount(){ return STATE.edits.size; }

function readWorkbook(file){
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
        resolve(wb);
      } catch (err) { reject(err); }
    };
    r.onerror = () => reject(new Error('File read failed'));
    r.readAsArrayBuffer(file);
  });
}

function findSheet(wb, canonName){
  const target = normKey(canonName);
  for (const name of wb.SheetNames){
    if (normKey(name) === target) return name;
  }
  return null;
}

function ingestWorkbook(wb){
  const warnings = [];
  const raw = {};

  for (const sheetName of Object.keys(SCHEMA)){
    const actual = findSheet(wb, sheetName);
    if (!actual){
      warnings.push(`Sheet "${sheetName}" not found — that section will be empty.`);
      raw[sheetName] = [];
      continue;
    }
    const ws = wb.Sheets[actual];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: false });
    const { required, optional } = SCHEMA[sheetName];

    // Canonicalize keys
    const canon = rows.map(r => canonicalizeRow(r, required, optional));

    // Check required columns
    if (canon.length){
      const missing = required.filter(col => !(col in canon[0]));
      if (missing.length){
        warnings.push(`Sheet "${sheetName}" is missing required columns: ${missing.join(', ')}. Rows will still load but related KPIs may be incomplete.`);
      }
    } else {
      warnings.push(`Sheet "${sheetName}" is empty.`);
    }
    raw[sheetName] = canon;
  }

  // Optionally pick up a Confirmations sheet (from a re-uploaded export)
  const confSheet = findSheet(wb, 'Confirmations');
  let confirmations = [];
  if (confSheet){
    const confRows = XLSX.utils.sheet_to_json(wb.Sheets[confSheet], { defval: null, raw: false });
    confirmations = confRows.map(r => {
      const ts = r.Timestamp ? new Date(r.Timestamp) : new Date();
      return {
        id: 'imp_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6),
        timestamp: isNaN(ts) ? new Date() : ts,
        OrderNumber: r.OrderNumber || '',
        WorkCenter: r.WorkCenter || '',
        Operator: r.Operator || '',
        Action: String(r.Action || '').toUpperCase(),
        Quantity: r.Quantity != null && r.Quantity !== '' ? Number(r.Quantity) : undefined,
        Minutes: r.Minutes != null && r.Minutes !== '' ? Number(r.Minutes) : undefined
      };
    }).filter(c => c.OrderNumber && c.Action);
    if (confirmations.length){
      warnings.push(`Restored ${confirmations.length} floor confirmation${confirmations.length===1?'':'s'} from the Confirmations sheet.`);
    }
  }

  return { raw, warnings, confirmations };
}

/* ============================================================
   4. KPI / FLAG COMPUTATION
   ============================================================ */
function computeAll(raw){
  const today = todayStartOfDay();
  const warnings = [];

  // Derive live state from any confirmations the floor has logged
  const confirmations = STATE.confirmations || [];
  const orderLive = window.FloorplanShared
    ? window.FloorplanShared.deriveOrderLiveState(confirmations)
    : new Map();

  // Clean & coerce — augment ConfirmedQty with floor confirmations
  const wos = (raw.WorkOrders || []).map(r => {
    const live = orderLive.get(r.OrderNumber);
    const baseConfirmed = Number(r.ConfirmedQty) || 0;
    const floorAdded = live ? live.ActualConfirmed : 0;
    const liveStatus = live ? live.LiveStatus : null;
    // If the operator marked it complete on the floor, treat as TECO
    const status = liveStatus === 'COMPLETE'
      ? 'TECO'
      : String(r.Status || 'CRTD').toUpperCase().trim();
    return {
      OrderNumber: r.OrderNumber,
      Material: r.Material,
      Description: r.Description || '',
      OrderQty: Number(r.OrderQty) || 0,
      ConfirmedQty: baseConfirmed + floorAdded,
      WorkCenter: r.WorkCenter,
      StartDate: parseDate(r.StartDate),
      FinishDate: parseDate(r.FinishDate),
      Priority: Number(r.Priority) || 3,
      Status: status,
      LiveStatus: liveStatus,
      ScrapQty: live ? live.ActualScrap : 0,
      DowntimeMinutes: live ? live.DowntimeMinutes : 0
    };
  }).filter(w => w.OrderNumber);

  const mats = (raw.Materials || []).map(r => ({
    Material: r.Material,
    Description: r.Description || '',
    OnHand: Number(r.OnHand) || 0,
    Reserved: Number(r.Reserved) || 0,
    SafetyStock: Number(r.SafetyStock) || 0,
    ReorderPoint: Number(r.ReorderPoint) || 0,
    LeadTimeDays: Number(r.LeadTimeDays) || 0,
    UoM: r.UoM || ''
  })).filter(m => m.Material);

  const bom = (raw.BOM || []).map(r => ({
    ParentMaterial: r.ParentMaterial,
    Component: r.Component,
    QtyPerParent: Number(r.QtyPerParent) || 0,
    UoM: r.UoM || ''
  })).filter(b => b.ParentMaterial && b.Component);

  const wcs = (raw.WorkCenters || []).map(r => ({
    WorkCenter: r.WorkCenter,
    Description: r.Description || '',
    HoursPerDay: Number(r.HoursPerDay) || 0,
    DaysPerWeek: Number(r.DaysPerWeek) || 0,
    Efficiency: Number(r.Efficiency) || SETTINGS.defaultEfficiency,
    HoursPerUnit: Number(r.HoursPerUnit) || SETTINGS.defaultHoursPerUnit
  })).filter(w => w.WorkCenter);

  // Index maps
  const matByCode = new Map(mats.map(m => [m.Material, m]));
  const wcByCode  = new Map(wcs.map(w => [w.WorkCenter, w]));
  const bomByParent = new Map();
  for (const b of bom){
    if (!bomByParent.has(b.ParentMaterial)) bomByParent.set(b.ParentMaterial, []);
    bomByParent.get(b.ParentMaterial).push(b);
  }

  /* ── INVENTORY ─────────────────── */
  const inventory = mats.map(m => {
    const available = m.OnHand - m.Reserved;
    let status = 'OK';
    if (available <= 0) status = 'STOCKOUT';
    else if (available <= m.ReorderPoint) status = 'BELOW REORDER';
    else if (available < m.SafetyStock * SETTINGS.lowStockMultiplier) status = 'LOW STOCK';
    return { ...m, Available: available, Status: status };
  });

  /* ── WORK ORDER FLAGS ─────────── */
  const enrichedWOs = wos.map(w => {
    const wc = wcByCode.get(w.WorkCenter);
    const hpu = wc ? wc.HoursPerUnit : SETTINGS.defaultHoursPerUnit;
    const openQty = Math.max(0, w.OrderQty - w.ConfirmedQty);
    const estHours = openQty * hpu;
    const daysToFinish = w.FinishDate ? daysBetween(today, w.FinishDate) : null;

    let flag = 'ON TRACK';
    let tone = 'ok';
    if (w.Status === 'TECO' || w.Status === 'CLSD'){
      flag = 'COMPLETE'; tone = 'ok';
    } else if (daysToFinish !== null && daysToFinish < -SETTINGS.pastDueGraceDays && openQty > 0){
      flag = 'PAST DUE'; tone = 'crit';
    } else if (daysToFinish !== null && daysToFinish <= SETTINGS.riskWindowDays){
      flag = 'AT RISK'; tone = 'warn';
    }

    return { ...w, OpenQty: openQty, EstHours: estHours, DaysToFinish: daysToFinish, Flag: flag, Tone: tone };
  });

  /* ── SHORTAGES (per WO × component) ── */
  const shortages = [];
  // Track demand per material across all open orders (so the same component
  // referenced by multiple WOs is flagged when total demand exceeds available)
  const demandByMaterial = new Map();
  for (const w of enrichedWOs){
    if (w.Flag === 'COMPLETE' || w.OpenQty <= 0) continue;
    const components = bomByParent.get(w.Material) || [];
    for (const c of components){
      const required = c.QtyPerParent * w.OpenQty;
      demandByMaterial.set(c.Component, (demandByMaterial.get(c.Component) || 0) + required);
    }
  }
  for (const w of enrichedWOs){
    if (w.Flag === 'COMPLETE' || w.OpenQty <= 0) continue;
    const components = bomByParent.get(w.Material) || [];
    for (const c of components){
      const required = c.QtyPerParent * w.OpenQty;
      const matRec = matByCode.get(c.Component);
      const available = matRec ? (matRec.OnHand - matRec.Reserved) : 0;
      const totalDemand = demandByMaterial.get(c.Component) || 0;
      const shortByLine = Math.max(0, required - available);
      let flag = 'OK', tone = 'ok';
      if (available < totalDemand){
        // material is collectively short across all WOs
        flag = 'SHORTAGE'; tone = 'crit';
      } else if (available < totalDemand * 1.05){
        flag = 'TIGHT'; tone = 'warn';
      }
      shortages.push({
        OrderNumber: w.OrderNumber,
        ParentMaterial: w.Material,
        Component: c.Component,
        ComponentDesc: matRec ? matRec.Description : '',
        Required: required,
        Available: available,
        Short: shortByLine,
        Flag: flag,
        Tone: tone
      });
    }
  }

  /* ── CAPACITY LOAD ───────────── */
  // Sum est hours per WC for open orders only
  const demandByWC = new Map();
  for (const w of enrichedWOs){
    if (w.Flag === 'COMPLETE') continue;
    demandByWC.set(w.WorkCenter, (demandByWC.get(w.WorkCenter) || 0) + w.EstHours);
  }
  const capacity = wcs.map(wc => {
    const effective = wc.HoursPerDay * wc.DaysPerWeek * wc.Efficiency;
    const demand = demandByWC.get(wc.WorkCenter) || 0;
    const load = effective > 0 ? demand / effective : 0;
    let flag = 'OK', tone = 'ok';
    if (load > SETTINGS.capacityCrit){ flag = 'OVERLOAD'; tone = 'crit'; }
    else if (load > SETTINGS.capacityWarn){ flag = 'HIGH LOAD'; tone = 'warn'; }
    return { ...wc, Effective: effective, Demand: demand, Load: load, Flag: flag, Tone: tone };
  });

  /* ── SCHEDULE (today's prioritized list) ── */
  // Sort: open first, by priority asc, then earliest finish, then past-due first
  const schedule = enrichedWOs
    .filter(w => w.Flag !== 'COMPLETE' && w.OpenQty > 0)
    .sort((a,b) => {
      // Past due first
      const aPD = a.Flag === 'PAST DUE' ? 0 : 1;
      const bPD = b.Flag === 'PAST DUE' ? 0 : 1;
      if (aPD !== bPD) return aPD - bPD;
      // Priority asc (1 = highest)
      if (a.Priority !== b.Priority) return a.Priority - b.Priority;
      // Earliest finish
      const af = a.FinishDate ? a.FinishDate.getTime() : Infinity;
      const bf = b.FinishDate ? b.FinishDate.getTime() : Infinity;
      return af - bf;
    });

  /* ── KPIs ────────────────────── */
  const openCount = enrichedWOs.filter(w => w.Flag !== 'COMPLETE' && w.OpenQty > 0).length;
  const pastDueCount = enrichedWOs.filter(w => w.Flag === 'PAST DUE').length;
  const atRiskCount = enrichedWOs.filter(w => w.Flag === 'AT RISK').length;

  // Unique component shortages
  const uniqueShortageComps = new Set(shortages.filter(s => s.Flag === 'SHORTAGE').map(s => s.Component));
  const shortageCount = uniqueShortageComps.size;
  const shortageLineCount = shortages.filter(s => s.Flag === 'SHORTAGE').length;

  const lowStockCount = inventory.filter(i =>
    i.Status === 'LOW STOCK' || i.Status === 'BELOW REORDER' || i.Status === 'STOCKOUT'
  ).length;

  const overloadCount = capacity.filter(c => c.Flag === 'OVERLOAD').length;

  // Schedule adherence: % of closed orders (TECO/CLSD) whose FinishDate is on
  // or before today. Without an "ActualFinishDate" column we can't measure
  // closure-vs-due directly, so this assumes the order closed approximately
  // when it appeared in TECO status — a common compromise when an actuals
  // column is unavailable. The Data Notices panel surfaces this caveat.
  const closed = enrichedWOs.filter(w => w.Status === 'TECO' || w.Status === 'CLSD');
  let adherence = null;
  if (closed.length){
    // An order is treated as "on time" if its planned FinishDate is <= today
    // (i.e. it had reached or passed its due date by the time we observed it
    // as closed). Orders closed *before* their due date count on-time; orders
    // still open past due reduce adherence by being counted in the denominator
    // alongside closed-on-time orders.
    const onTime = closed.filter(w => w.FinishDate && daysBetween(w.FinishDate, today) >= 0).length;
    const denom = closed.length + pastDueCount;
    adherence = denom > 0 ? onTime / denom : null;
    warnings.push('Schedule adherence is approximate — add an "ActualFinishDate" column to your WorkOrders sheet for an exact metric.');
  } else {
    warnings.push('No closed (TECO/CLSD) orders found — schedule adherence cannot be calculated.');
  }

  const kpis = {
    open: openCount,
    pastDue: pastDueCount,
    atRisk: atRiskCount,
    shortageMaterials: shortageCount,
    shortageLines: shortageLineCount,
    lowStock: lowStockCount,
    overload: overloadCount,
    adherence
  };

  return { kpis, schedule, shortages, capacity, inventory, enrichedWOs, warnings };
}

/* ============================================================
   5. RENDERING
   ============================================================ */
function renderEmpty(tbody, colspan, msg){
  tbody.innerHTML = `<tr class="empty-row"><td colspan="${colspan}">${msg}</td></tr>`;
}

function flagPill(flag, tone){
  const cls = tone === 'crit' ? 'pill--crit' : tone === 'warn' ? 'pill--warn' : tone === 'ok' ? 'pill--ok' : 'pill--mute';
  return `<span class="pill ${cls}">${flag}</span>`;
}

function escapeHtml(s){
  return String(s ?? '').replace(/[&<>"']/g, ch => ({
    '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'
  }[ch]));
}

function renderAll(c){
  // KPIs
  $('kpiOpen').textContent = fmtNum(c.kpis.open);
  $('kpiOpenSub').textContent = `${fmtNum(c.enrichedWOs.length)} total in file`;
  $('kpiPastDue').textContent = fmtNum(c.kpis.pastDue);
  $('kpiAtRisk').textContent = fmtNum(c.kpis.atRisk);
  $('kpiShortages').textContent = fmtNum(c.kpis.shortageMaterials);
  $('kpiShortagesSub').textContent = `${fmtNum(c.kpis.shortageLines)} order-component lines`;
  $('kpiLowStock').textContent = fmtNum(c.kpis.lowStock);
  $('kpiOverload').textContent = fmtNum(c.kpis.overload);
  $('kpiAdherence').textContent = c.kpis.adherence !== null ? fmtPct(c.kpis.adherence) : '—';

  // Header status
  if (c.kpis.pastDue > 0 || c.kpis.shortageMaterials > 0 || c.kpis.overload > 0){
    setStatus('crit', 'CRITICAL ISSUES');
  } else if (c.kpis.atRisk > 0 || c.kpis.lowStock > 0){
    setStatus('warn', 'WARNINGS PRESENT');
  } else {
    setStatus('ok', 'ALL CLEAR');
  }

  // SCHEDULE (editable: OrderQty, ConfirmedQty, Priority)
  const sb = $('scheduleBody');
  if (c.schedule.length === 0){
    renderEmpty(sb, 11, 'NO OPEN ORDERS');
  } else {
    sb.innerHTML = c.schedule.map((w,i) => {
      const idx = (STATE.raw.WorkOrders || []).findIndex(r => r.OrderNumber === w.OrderNumber);
      const editFlag = (field) => idx >= 0 && isEdited('WorkOrders', idx, field) ? ' is-edited' : '';
      return `
        <tr>
          <td>${i+1}</td>
          <td><b>${escapeHtml(w.OrderNumber)}</b></td>
          <td>${escapeHtml(w.Material)}</td>
          <td>${escapeHtml(w.Description)}</td>
          <td class="num">${fmtNum(w.OpenQty)}</td>
          <td class="num">${fmtNum(w.EstHours,1)}</td>
          <td>${escapeHtml(w.WorkCenter)}</td>
          <td>${fmtDate(w.FinishDate)}</td>
          <td class="num">${w.DaysToFinish === null ? '—' : w.DaysToFinish}</td>
          <td class="editable${editFlag('Priority')}" data-sheet="WorkOrders" data-idx="${idx}" data-field="Priority" tabindex="0"><span class="prio" data-p="${w.Priority}"><span class="prio__dot"></span>P${w.Priority}</span></td>
          <td>${flagPill(w.Flag, w.Tone)}</td>
        </tr>
      `;
    }).join('');
  }

  // SHORTAGES
  const sh = $('shortageBody');
  const shortageRows = c.shortages.filter(s => s.Flag !== 'OK');
  if (shortageRows.length === 0){
    renderEmpty(sh, 6, 'NO MATERIAL SHORTAGES');
  } else {
    sh.innerHTML = shortageRows.map(s => `
      <tr>
        <td><b>${escapeHtml(s.OrderNumber)}</b></td>
        <td>${escapeHtml(s.Component)}<br><span style="color:var(--ink-mute);font-size:10px">${escapeHtml(s.ComponentDesc)}</span></td>
        <td class="num">${fmtNum(s.Required,2)}</td>
        <td class="num">${fmtNum(s.Available,2)}</td>
        <td class="num">${fmtNum(s.Short,2)}</td>
        <td>${flagPill(s.Flag, s.Tone)}</td>
      </tr>
    `).join('');
  }

  // CAPACITY (editable: HoursPerDay, DaysPerWeek, Efficiency, HoursPerUnit)
  const cb = $('capacityBody');
  if (c.capacity.length === 0){
    renderEmpty(cb, 6, 'NO WORK CENTERS DEFINED');
  } else {
    cb.innerHTML = c.capacity.map(w => {
      const pct = Math.min(w.Load * 100, 200);
      const barCls = w.Tone === 'crit' ? 'bar--crit' : w.Tone === 'warn' ? 'bar--warn' : '';
      const idx = (STATE.raw.WorkCenters || []).findIndex(r => r.WorkCenter === w.WorkCenter);
      // We don't render individual editable cells in this summary table because
      // it would crowd the layout — capacity inputs are edited through their
      // own dedicated mini-form (see settings drawer below). To keep this
      // section purely informational for now.
      return `
        <tr>
          <td><b>${escapeHtml(w.WorkCenter)}</b></td>
          <td>${escapeHtml(w.Description)}</td>
          <td class="num">${fmtNum(w.Effective,1)}</td>
          <td class="num">${fmtNum(w.Demand,1)}</td>
          <td class="num">
            ${fmtPct(w.Load,0)}
            <div class="bar ${barCls}"><div class="bar__fill" style="width:${pct}%"></div></div>
          </td>
          <td>${flagPill(w.Flag, w.Tone)}</td>
        </tr>
      `;
    }).join('');
  }

  // INVENTORY (editable: OnHand, Reserved, SafetyStock, ReorderPoint, LeadTimeDays)
  const ib = $('inventoryBody');
  if (c.inventory.length === 0){
    renderEmpty(ib, 9, 'NO MATERIALS LOADED');
  } else {
    ib.innerHTML = c.inventory.map((m, idx) => {
      const tone = m.Status === 'STOCKOUT' || m.Status === 'BELOW REORDER' ? 'crit'
                 : m.Status === 'LOW STOCK' ? 'warn' : 'ok';
      const editFlag = (field) => isEdited('Materials', idx, field) ? ' is-edited' : '';
      return `
        <tr>
          <td><b>${escapeHtml(m.Material)}</b></td>
          <td>${escapeHtml(m.Description)}</td>
          <td class="num editable${editFlag('OnHand')}" data-sheet="Materials" data-idx="${idx}" data-field="OnHand" tabindex="0">${fmtNum(m.OnHand,2)}</td>
          <td class="num editable${editFlag('Reserved')}" data-sheet="Materials" data-idx="${idx}" data-field="Reserved" tabindex="0">${fmtNum(m.Reserved,2)}</td>
          <td class="num">${fmtNum(m.Available,2)}</td>
          <td class="num editable${editFlag('SafetyStock')}" data-sheet="Materials" data-idx="${idx}" data-field="SafetyStock" tabindex="0">${fmtNum(m.SafetyStock,2)}</td>
          <td class="num editable${editFlag('ReorderPoint')}" data-sheet="Materials" data-idx="${idx}" data-field="ReorderPoint" tabindex="0">${fmtNum(m.ReorderPoint,2)}</td>
          <td class="num editable${editFlag('LeadTimeDays')}" data-sheet="Materials" data-idx="${idx}" data-field="LeadTimeDays" tabindex="0">${fmtNum(m.LeadTimeDays)}</td>
          <td>${flagPill(m.Status, tone)}</td>
        </tr>
      `;
    }).join('');
  }

  // WARNINGS
  const allWarnings = [...STATE.warnings, ...c.warnings];
  if (allWarnings.length){
    $('warningsBlock').hidden = false;
    $('warningsList').innerHTML = allWarnings.map(w => `<li>${escapeHtml(w)}</li>`).join('');
  } else {
    $('warningsBlock').hidden = true;
  }

  // Update edit counter in toolbar
  updateEditCounter();

  // New sections
  renderLiveFloor(c);
  renderBacklog(c);
}

/* ── LIVE FLOOR ─────────────────────────────────────── */
function renderLiveFloor(c){
  const grid = $('liveGrid');
  if (!grid) return;
  const wcs = (STATE.raw.WorkCenters || []);
  const confirmations = STATE.confirmations || [];
  if (confirmations.length === 0){
    grid.innerHTML = '<div class="live-empty">No floor activity yet — open the operator screen and start reporting work.</div>';
    return;
  }
  const Shared = window.FloorplanShared;
  const wcLive = Shared ? Shared.deriveWCLiveState(confirmations, c.enrichedWOs) : new Map();
  grid.innerHTML = wcs.map(w => {
    const s = wcLive.get(w.WorkCenter) || { Running:0, Paused:0, Idle:0, RunningOrders:[] };
    const runningOrders = s.RunningOrders.slice(0, 3).join(', ') + (s.RunningOrders.length > 3 ? `, +${s.RunningOrders.length - 3} more` : '');
    return `
      <div class="live-cell">
        <div class="live-cell__wc">WORK CENTER</div>
        <div class="live-cell__title">${escapeHtml(w.WorkCenter)}</div>
        <div class="live-cell__counts">
          <span><span class="live-cell__dot live-cell__dot--run"></span><b>${s.Running}</b> RUN</span>
          <span><span class="live-cell__dot live-cell__dot--pause"></span><b>${s.Paused}</b> PAUSE</span>
          <span><span class="live-cell__dot live-cell__dot--idle"></span><b>${s.Idle}</b> IDLE</span>
        </div>
        <div class="live-cell__orders">${escapeHtml(runningOrders || '—')}</div>
      </div>
    `;
  }).join('');
}

/* ── BACKLOG AGING ──────────────────────────────────── */
function renderBacklog(c){
  if (!$('backlog13')) return;
  const today = todayStartOfDay();
  // Past-due open orders only
  const pastDue = c.enrichedWOs.filter(w => w.Flag === 'PAST DUE');
  const total = c.enrichedWOs.filter(w => w.Flag !== 'COMPLETE' && w.OpenQty > 0).length;

  let b13 = 0, b47 = 0, b8 = 0;
  const rows = [];
  for (const w of pastDue){
    const daysLate = w.FinishDate ? -daysBetween(today, w.FinishDate) : 0;
    let bucket;
    if (daysLate <= 3){ b13++; bucket = 'warn'; }
    else if (daysLate <= 7){ b47++; bucket = 'crit'; }
    else { b8++; bucket = 'crit'; }
    rows.push({ ...w, daysLate, bucket });
  }
  $('backlog13').textContent = b13;
  $('backlog47').textContent = b47;
  $('backlog8').textContent = b8;
  $('backlogTotalOpen').textContent = total;

  const list = $('backlogList');
  if (rows.length === 0){
    list.innerHTML = '<div style="padding:18px;text-align:center;color:var(--ink-mute);font-family:var(--mono);font-size:11px;letter-spacing:0.12em">NO PAST-DUE ORDERS</div>';
    return;
  }
  // Sort by days late descending
  rows.sort((a,b) => b.daysLate - a.daysLate);
  list.innerHTML = rows.slice(0, 30).map(w => `
    <div class="backlog__row" data-age="${w.bucket}">
      <span><b>${escapeHtml(w.OrderNumber)}</b></span>
      <span>${escapeHtml(w.Material)} · ${escapeHtml(w.Description || '')}</span>
      <span>${escapeHtml(w.WorkCenter || '')}</span>
      <span style="text-align:right">${fmtNum(w.OpenQty)} open</span>
      <span class="backlog__age">${w.daysLate}D LATE</span>
    </div>
  `).join('');
}

/* ============================================================
   5b. CLICK-TO-EDIT
   ============================================================ */
function updateEditCounter(){
  const el = $('editCounter');
  if (!el) return;
  const n = editCount();
  if (n === 0){
    el.textContent = '';
    el.hidden = true;
    $('resetEditsBtn').hidden = true;
  } else {
    el.textContent = `${n} edit${n === 1 ? '' : 's'}`;
    el.hidden = false;
    $('resetEditsBtn').hidden = false;
  }
}

function recomputeAndRender(){
  const computed = computeAll(STATE.raw);
  STATE.computed = computed;
  renderAll(computed);
  persistShared();
}

function applyEdit(sheet, idx, field, newValue){
  const arr = STATE.raw[sheet];
  if (!arr || !arr[idx]) return false;
  const original = arr[idx][field];
  // Coerce value
  let coerced = newValue;
  if (typeof original === 'number' || ['OnHand','Reserved','SafetyStock','ReorderPoint','LeadTimeDays','OrderQty','ConfirmedQty','Priority','HoursPerDay','DaysPerWeek','Efficiency','HoursPerUnit','QtyPerParent'].includes(field)){
    const n = Number(newValue);
    if (isNaN(n)){
      toast('Must be a number', true);
      return false;
    }
    coerced = n;
  }
  // Don't record an edit if the value didn't actually change
  if (coerced === original) return false;
  recordEdit(sheet, idx, field, original);
  arr[idx][field] = coerced;
  return true;
}

function startCellEdit(td){
  if (td.classList.contains('is-editing')) return;
  const sheet = td.dataset.sheet;
  const idx = parseInt(td.dataset.idx, 10);
  const field = td.dataset.field;
  if (!sheet || isNaN(idx) || !field) return;
  if (idx < 0) { toast('Cannot edit this row', true); return; }

  const currentValue = STATE.raw[sheet][idx][field];
  const startVal = currentValue ?? '';

  td.classList.add('is-editing');
  const input = document.createElement('input');
  input.type = 'text';
  input.inputMode = 'decimal';
  input.value = startVal;
  input.className = 'cell-input';

  // Replace cell content with input, but remember to restore on cancel
  td._beforeEditHTML = td.innerHTML;
  td.innerHTML = '';
  td.appendChild(input);
  input.focus();
  input.select();

  const finish = (commit) => {
    if (td._finished) return;
    td._finished = true;
    td.classList.remove('is-editing');
    const raw = input.value.trim();
    if (commit && raw !== String(startVal)){
      const ok = applyEdit(sheet, idx, field, raw);
      if (ok){
        recomputeAndRender();
        return; // recomputeAndRender re-renders the whole table
      }
    }
    // No commit (or no change / failed): restore previous content
    td.innerHTML = td._beforeEditHTML;
  };

  input.addEventListener('keydown', e => {
    if (e.key === 'Enter' || e.key === 'Tab'){ e.preventDefault(); finish(true); }
    else if (e.key === 'Escape'){ e.preventDefault(); finish(false); }
  });
  input.addEventListener('blur', () => finish(true));
}

function resetAllEdits(){
  if (editCount() === 0) return;
  if (!confirm(`Discard ${editCount()} edit${editCount()===1?'':'s'} and revert to the uploaded file?`)) return;
  // Walk the edits map and write originals back
  for (const [key, originalValue] of STATE.edits.entries()){
    const [sheet, idxStr, field] = key.split(':');
    const idx = parseInt(idxStr, 10);
    if (STATE.raw[sheet] && STATE.raw[sheet][idx]){
      STATE.raw[sheet][idx][field] = originalValue;
    }
  }
  clearEdits();
  recomputeAndRender();
  toast('Edits reverted');
}

/* ============================================================
   6. EXPORT (with computed sheets appended)
   ============================================================ */
function exportEnrichedWorkbook(){
  if (!STATE.computed){ toast('No data to export', true); return; }
  const c = STATE.computed;
  const wb = XLSX.utils.book_new();

  const fmtRows = arr => arr.map(o => {
    const out = {};
    for (const k of Object.keys(o)){
      let v = o[k];
      if (v instanceof Date) v = fmtDate(v);
      out[k] = v;
    }
    return out;
  });

  // Original sheets (canonicalized)
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(fmtRows(STATE.raw.WorkOrders)), 'WorkOrders');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(fmtRows(STATE.raw.Materials)), 'Materials');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(fmtRows(STATE.raw.BOM)), 'BOM');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(fmtRows(STATE.raw.WorkCenters)), 'WorkCenters');

  // Computed sheets
  const scheduleOut = c.schedule.map((w,i) => ({
    Rank: i+1,
    OrderNumber: w.OrderNumber,
    Material: w.Material,
    Description: w.Description,
    OpenQty: w.OpenQty,
    EstHours: Number(w.EstHours.toFixed(2)),
    WorkCenter: w.WorkCenter,
    FinishDate: fmtDate(w.FinishDate),
    DaysToFinish: w.DaysToFinish,
    Priority: w.Priority,
    Status: w.Status,
    Flag: w.Flag
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(scheduleOut), 'Schedule');

  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(c.shortages.map(s => ({
    OrderNumber: s.OrderNumber, ParentMaterial: s.ParentMaterial,
    Component: s.Component, ComponentDesc: s.ComponentDesc,
    Required: Number(s.Required.toFixed(3)),
    Available: Number(s.Available.toFixed(3)),
    Short: Number(s.Short.toFixed(3)),
    Flag: s.Flag
  }))), 'Shortages');

  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(c.capacity.map(w => ({
    WorkCenter: w.WorkCenter, Description: w.Description,
    HoursPerDay: w.HoursPerDay, DaysPerWeek: w.DaysPerWeek,
    Efficiency: w.Efficiency, HoursPerUnit: w.HoursPerUnit,
    EffectiveHrsPerWeek: Number(w.Effective.toFixed(2)),
    DemandHrs: Number(w.Demand.toFixed(2)),
    Load: Number(w.Load.toFixed(4)),
    Flag: w.Flag
  }))), 'CapacityLoad');

  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(c.inventory.map(m => ({
    Material: m.Material, Description: m.Description,
    OnHand: m.OnHand, Reserved: m.Reserved, Available: m.Available,
    SafetyStock: m.SafetyStock, ReorderPoint: m.ReorderPoint,
    LeadTimeDays: m.LeadTimeDays, UoM: m.UoM, Status: m.Status
  }))), 'InventoryHealth');

  // Floor confirmations log (operator data capture)
  const confirmations = STATE.confirmations || [];
  if (confirmations.length){
    const confRows = confirmations.map(c => ({
      Timestamp: c.timestamp instanceof Date ? c.timestamp.toISOString() : c.timestamp,
      OrderNumber: c.OrderNumber || '',
      WorkCenter: c.WorkCenter || '',
      Operator: c.Operator || '',
      Action: c.Action || '',
      Quantity: c.Quantity ?? '',
      Minutes: c.Minutes ?? ''
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(confRows), 'Confirmations');
  }

  // KPI snapshot
  const k = c.kpis;
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([
    { KPI:'Snapshot Date',          Value: fmtDate(todayStartOfDay()) },
    { KPI:'Open Orders',            Value: k.open },
    { KPI:'Past Due',               Value: k.pastDue },
    { KPI:'At Risk (5d)',           Value: k.atRisk },
    { KPI:'Material Shortages',     Value: k.shortageMaterials },
    { KPI:'Shortage Lines',         Value: k.shortageLines },
    { KPI:'Low Stock Items',        Value: k.lowStock },
    { KPI:'Overloaded Work Centers',Value: k.overload },
    { KPI:'Schedule Adherence',     Value: k.adherence !== null ? fmtPct(k.adherence,1) : 'N/A' }
  ]), 'KPI_Snapshot');

  const stamp = new Date().toISOString().slice(0,10);
  const filename = `floorplan_${stamp}.xlsx`;
  XLSX.writeFile(wb, filename);
  toast(`Exported ${filename}`);
}

/* ============================================================
   7. STARTER TEMPLATE & DEMO DATA
   ============================================================ */
function buildStarterTemplate(){
  const wb = XLSX.utils.book_new();
  const woHeader = ['OrderNumber','Material','Description','OrderQty','ConfirmedQty','WorkCenter','StartDate','FinishDate','Priority','Status'];
  const matHeader = ['Material','Description','OnHand','Reserved','SafetyStock','ReorderPoint','LeadTimeDays','UoM'];
  const bomHeader = ['ParentMaterial','Component','QtyPerParent','UoM'];
  const wcHeader  = ['WorkCenter','Description','HoursPerDay','DaysPerWeek','Efficiency','HoursPerUnit'];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([woHeader]),  'WorkOrders');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([matHeader]), 'Materials');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([bomHeader]), 'BOM');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([wcHeader]),  'WorkCenters');
  // README sheet
  const readme = [
    ['FLOORPLAN — Starter Template'],
    [''],
    ['Fill in the four sheets, then upload this file at the website.'],
    ['Column names are case- and space-insensitive.'],
    ['Dates can be Excel dates or text in YYYY-MM-DD format.'],
    [''],
    ['Status values: CRTD (created), REL (released), TECO (technically complete), CLSD (closed)'],
    ['Priority: 1 = highest, 5 = lowest']
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(readme), 'README');
  XLSX.writeFile(wb, 'floorplan_template.xlsx');
  toast('Template downloaded');
}

function buildDemoWorkbook(){
  const today = todayStartOfDay();
  const d = (offset) => {
    const x = new Date(today); x.setDate(x.getDate() + offset);
    return fmtDate(x);
  };
  const wos = [
    ['1000234','FG-A100','Pump Assembly Type A',500,350,'WC-ASSY-01',d(-7),d(-1),1,'REL'],
    ['1000235','FG-B200','Valve Body B-Series',200,200,'WC-MACH-02',d(-12),d(-5),3,'TECO'],
    ['1000236','FG-C300','Control Module v3',150,0,'WC-ELEC-01',d(0),d(4),2,'CRTD'],
    ['1000237','FG-A100','Pump Assembly Type A',800,100,'WC-ASSY-01',d(-1),d(3),1,'REL'],
    ['1000238','FG-D400','Housing D-Type',1000,600,'WC-MACH-01',d(-5),d(1),3,'REL'],
    ['1000239','FG-E500','Sensor Assembly',300,0,'WC-ELEC-01',d(8),d(12),4,'CRTD'],
    ['1000240','FG-B200','Valve Body B-Series',250,0,'WC-MACH-02',d(1),d(6),2,'REL'],
    ['1000241','FG-A100','Pump Assembly Type A',600,0,'WC-ASSY-01',d(4),d(9),3,'REL'],
    ['1000242','FG-D400','Housing D-Type',400,400,'WC-MACH-01',d(-20),d(-15),4,'TECO'],
    ['1000243','FG-C300','Control Module v3',100,0,'WC-ELEC-01',d(2),d(7),5,'CRTD']
  ];
  const mats = [
    ['RM-STEEL-01','Steel Sheet 2mm',4500,1200,1000,1500,14,'KG'],
    ['RM-STEEL-02','Steel Rod 10mm',800,600,500,750,21,'KG'],
    ['RM-PLASTIC-01','ABS Pellets',200,150,300,400,10,'KG'],
    ['RM-ELEC-01','PCB Type A',120,200,150,250,28,'EA'],
    ['RM-ELEC-02','Sensor Chip X',50,30,80,120,35,'EA'],
    ['RM-SEAL-01','O-Ring 25mm',12000,3000,2000,3000,7,'EA'],
    ['FG-A100','Pump Assembly Type A',240,100,200,300,0,'EA'],
    ['FG-B200','Valve Body B-Series',80,50,60,100,0,'EA'],
    ['FG-C300','Control Module v3',30,80,50,80,0,'EA'],
    ['FG-D400','Housing D-Type',400,200,250,400,0,'EA'],
    ['FG-E500','Sensor Assembly',60,0,40,70,0,'EA']
  ];
  const bom = [
    ['FG-A100','RM-STEEL-01',2.5,'KG'],['FG-A100','RM-SEAL-01',4,'EA'],['FG-A100','RM-ELEC-01',1,'EA'],
    ['FG-B200','RM-STEEL-02',0.8,'KG'],['FG-B200','RM-SEAL-01',2,'EA'],
    ['FG-C300','RM-ELEC-01',1,'EA'],['FG-C300','RM-ELEC-02',2,'EA'],['FG-C300','RM-PLASTIC-01',0.3,'KG'],
    ['FG-D400','RM-STEEL-01',1.8,'KG'],['FG-D400','RM-PLASTIC-01',0.5,'KG'],
    ['FG-E500','RM-ELEC-02',1,'EA'],['FG-E500','RM-PLASTIC-01',0.2,'KG']
  ];
  const wcs = [
    ['WC-ASSY-01','Assembly Line 1',16,5,0.92,0.4],
    ['WC-MACH-01','CNC Machining 1',24,7,0.88,0.25],
    ['WC-MACH-02','CNC Machining 2',16,5,0.90,0.3],
    ['WC-ELEC-01','Electronics Assembly',8,5,0.95,0.8]
  ];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ['OrderNumber','Material','Description','OrderQty','ConfirmedQty','WorkCenter','StartDate','FinishDate','Priority','Status'],
    ...wos
  ]), 'WorkOrders');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ['Material','Description','OnHand','Reserved','SafetyStock','ReorderPoint','LeadTimeDays','UoM'],
    ...mats
  ]), 'Materials');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ['ParentMaterial','Component','QtyPerParent','UoM'],
    ...bom
  ]), 'BOM');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ['WorkCenter','Description','HoursPerDay','DaysPerWeek','Efficiency','HoursPerUnit'],
    ...wcs
  ]), 'WorkCenters');
  return wb;
}

/* ============================================================
   8. WIRING — events, lifecycle
   ============================================================ */
async function handleFile(file){
  if (!file) return;
  STATE.fileName = file.name;
  STATE.edits = new Map();
  STATE.confirmations = []; // fresh upload wipes confirmations from prior session
  setStatus('warn', 'PARSING…');
  try{
    const wb = await readWorkbook(file);
    const { raw, warnings, confirmations } = ingestWorkbook(wb);
    STATE.raw = raw;
    STATE.warnings = warnings;
    if (confirmations && confirmations.length) STATE.confirmations = confirmations;
    const computed = computeAll(raw);
    STATE.computed = computed;
    showDashboard();
    renderAll(computed);
    persistShared();
    toast(`Loaded ${file.name}`);
  } catch(err){
    console.error(err);
    setStatus('crit', 'PARSE FAILED');
    toast('Could not parse file: ' + err.message, true);
  }
}

function showDashboard(){
  $('intake').hidden = true;
  $('dash').hidden = false;
  $('loadedFile').textContent = STATE.fileName ? `LOADED · ${STATE.fileName}` : '—';
}
function showIntake(){
  $('intake').hidden = false;
  $('dash').hidden = true;
  setStatus('', 'NO DATA LOADED');
}

function tickClock(){
  const d = new Date();
  const pad = n => String(n).padStart(2,'0');
  $('clock').textContent = `${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}

function init(){
  // Clock
  tickClock();
  setInterval(tickClock, 1000);

  // File input
  const fileInput = $('fileInput');
  const dz = $('dropzone');

  dz.addEventListener('click', () => fileInput.click());
  dz.addEventListener('keydown', e => {
    if (e.key === 'Enter' || e.key === ' '){ e.preventDefault(); fileInput.click(); }
  });
  fileInput.addEventListener('change', e => {
    const f = e.target.files[0];
    if (f) handleFile(f);
  });

  ['dragenter','dragover'].forEach(ev =>
    dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.add('is-drag'); }));
  ['dragleave','drop'].forEach(ev =>
    dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.remove('is-drag'); }));
  dz.addEventListener('drop', e => {
    const f = e.dataTransfer.files && e.dataTransfer.files[0];
    if (f) handleFile(f);
  });

  // Buttons
  $('downloadTemplateBtn').addEventListener('click', buildStarterTemplate);
  $('loadDemoBtn').addEventListener('click', () => {
    const wb = buildDemoWorkbook();
    const { raw, warnings } = ingestWorkbook(wb);
    STATE.raw = raw;
    STATE.warnings = warnings;
    STATE.fileName = 'demo_data.xlsx';
    STATE.edits = new Map();
    STATE.confirmations = [];
    const computed = computeAll(raw);
    STATE.computed = computed;
    showDashboard();
    renderAll(computed);
    persistShared();
    toast('Demo data loaded');
  });
  $('exportBtn').addEventListener('click', exportEnrichedWorkbook);
  $('reuploadBtn').addEventListener('click', () => {
    if (editCount() > 0 && !confirm(`You have ${editCount()} unsaved edit${editCount()===1?'':'s'}. Replace file anyway?`)) return;
    showIntake();
    fileInput.value = '';
    STATE = { fileName:null, raw:null, warnings:[], computed:null, edits: new Map(), confirmations: [] };
    if (window.FloorplanShared) window.FloorplanShared.clear();
  });

  // Listen for updates pushed from the floor page (operator confirmations)
  if (window.FloorplanShared){
    window.FloorplanShared.onSync(msg => {
      if (msg.type === 'state-cleared') return;
      // Re-load shared state and re-render
      const s = window.FloorplanShared.load();
      if (s && s.confirmations){
        STATE.confirmations = s.confirmations;
        if (STATE.raw){
          const computed = computeAll(STATE.raw);
          STATE.computed = computed;
          renderAll(computed);
        }
      }
    });
    // On load, attempt to restore from shared (if user just came from the floor screen)
    if (!STATE.raw && tryLoadShared()){
      const computed = computeAll(STATE.raw);
      STATE.computed = computed;
      showDashboard();
      renderAll(computed);
    }
  }

  // Reset edits button
  $('resetEditsBtn').addEventListener('click', resetAllEdits);

  // Click-to-edit: event delegation on the dashboard
  $('dash').addEventListener('click', e => {
    const td = e.target.closest('td.editable');
    if (td) startCellEdit(td);
  });
  // Keyboard: Enter on focused editable cell starts edit
  $('dash').addEventListener('keydown', e => {
    if (e.key !== 'Enter') return;
    const td = e.target.closest('td.editable');
    if (td && !td.classList.contains('is-editing')){
      e.preventDefault();
      startCellEdit(td);
    }
  });
}

document.addEventListener('DOMContentLoaded', init);

})();
