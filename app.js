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
  computed: null    // { kpis, schedule, shortages, capacity, inventory }
};

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
  return { raw, warnings };
}

/* ============================================================
   4. KPI / FLAG COMPUTATION
   ============================================================ */
function computeAll(raw){
  const today = todayStartOfDay();
  const warnings = [];

  // Clean & coerce
  const wos = (raw.WorkOrders || []).map(r => ({
    OrderNumber: r.OrderNumber,
    Material: r.Material,
    Description: r.Description || '',
    OrderQty: Number(r.OrderQty) || 0,
    ConfirmedQty: Number(r.ConfirmedQty) || 0,
    WorkCenter: r.WorkCenter,
    StartDate: parseDate(r.StartDate),
    FinishDate: parseDate(r.FinishDate),
    Priority: Number(r.Priority) || 3,
    Status: String(r.Status || 'CRTD').toUpperCase().trim()
  })).filter(w => w.OrderNumber);

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

  // SCHEDULE
  const sb = $('scheduleBody');
  if (c.schedule.length === 0){
    renderEmpty(sb, 11, 'NO OPEN ORDERS');
  } else {
    sb.innerHTML = c.schedule.map((w,i) => `
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
        <td><span class="prio" data-p="${w.Priority}"><span class="prio__dot"></span>P${w.Priority}</span></td>
        <td>${flagPill(w.Flag, w.Tone)}</td>
      </tr>
    `).join('');
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

  // CAPACITY
  const cb = $('capacityBody');
  if (c.capacity.length === 0){
    renderEmpty(cb, 6, 'NO WORK CENTERS DEFINED');
  } else {
    cb.innerHTML = c.capacity.map(w => {
      const pct = Math.min(w.Load * 100, 200);
      const barCls = w.Tone === 'crit' ? 'bar--crit' : w.Tone === 'warn' ? 'bar--warn' : '';
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

  // INVENTORY
  const ib = $('inventoryBody');
  if (c.inventory.length === 0){
    renderEmpty(ib, 9, 'NO MATERIALS LOADED');
  } else {
    ib.innerHTML = c.inventory.map(m => {
      const tone = m.Status === 'STOCKOUT' || m.Status === 'BELOW REORDER' ? 'crit'
                 : m.Status === 'LOW STOCK' ? 'warn' : 'ok';
      return `
        <tr>
          <td><b>${escapeHtml(m.Material)}</b></td>
          <td>${escapeHtml(m.Description)}</td>
          <td class="num">${fmtNum(m.OnHand,2)}</td>
          <td class="num">${fmtNum(m.Reserved,2)}</td>
          <td class="num">${fmtNum(m.Available,2)}</td>
          <td class="num">${fmtNum(m.SafetyStock,2)}</td>
          <td class="num">${fmtNum(m.ReorderPoint,2)}</td>
          <td class="num">${fmtNum(m.LeadTimeDays)}</td>
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
  setStatus('warn', 'PARSING…');
  try{
    const wb = await readWorkbook(file);
    const { raw, warnings } = ingestWorkbook(wb);
    STATE.raw = raw;
    STATE.warnings = warnings;
    const computed = computeAll(raw);
    STATE.computed = computed;
    showDashboard();
    renderAll(computed);
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
    const computed = computeAll(raw);
    STATE.computed = computed;
    showDashboard();
    renderAll(computed);
    toast('Demo data loaded');
  });
  $('exportBtn').addEventListener('click', exportEnrichedWorkbook);
  $('reuploadBtn').addEventListener('click', () => {
    showIntake();
    fileInput.value = '';
    STATE = { fileName:null, raw:null, warnings:[], computed:null };
  });
}

document.addEventListener('DOMContentLoaded', init);

})();
