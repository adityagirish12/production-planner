/* ──────────────────────────────────────────────────────────
   FLOORPLAN · Floor (operator) screen logic
   ────────────────────────────────────────────────────────── */
(() => {
'use strict';

const $ = id => document.getElementById(id);

let SHARED = null;       // loaded from localStorage
let CURRENT_WC = null;   // selected work center code
let CURRENT_ORDER = null;// order being acted on (drawer open)
let NUMPAD_CALLBACK = null;
let NUMPAD_VALUE = '0';

/* ── UTILITIES ────────────────────────────────────── */
function toast(msg, isErr=false){
  const t = $('toast');
  t.textContent = msg;
  t.classList.toggle('toast--err', isErr);
  t.classList.add('is-show');
  clearTimeout(t._timer);
  t._timer = setTimeout(()=>t.classList.remove('is-show'), 2200);
}
function fmtNum(n){
  if (n === null || n === undefined || isNaN(n)) return '—';
  return Number(n).toLocaleString();
}
function fmtTime(d){
  if (!d) return '—';
  const dt = d instanceof Date ? d : new Date(d);
  return dt.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false });
}
function fmtDate(s){
  if (!s) return '—';
  if (typeof s === 'string') return s.slice(0,10);
  return s.toISOString().slice(0,10);
}
function escapeHtml(s){
  return String(s ?? '').replace(/[&<>"']/g, ch => ({
    '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'
  }[ch]));
}

/* ── LOAD SHARED STATE ────────────────────────────── */
function refreshFromShared(){
  SHARED = window.FloorplanShared.load();
  if (!SHARED || !SHARED.raw || !SHARED.raw.WorkOrders){
    $('emptyState').hidden = false;
    $('floorMain').hidden = true;
    return;
  }
  $('emptyState').hidden = true;
  $('floorMain').hidden = false;
  populateWCDropdown();
  renderAll();
}

function populateWCDropdown(){
  const sel = $('wcSelect');
  const wcs = (SHARED.raw.WorkCenters || []);
  const prevValue = CURRENT_WC || sel.value;
  sel.innerHTML = '<option value="">— SELECT WORK CENTER —</option>' +
    wcs.map(w => `<option value="${escapeHtml(w.WorkCenter)}">${escapeHtml(w.WorkCenter)} — ${escapeHtml(w.Description || '')}</option>`).join('');
  if (prevValue && wcs.find(w => w.WorkCenter === prevValue)){
    sel.value = prevValue;
    CURRENT_WC = prevValue;
  }
}

/* ── RENDER ───────────────────────────────────────── */
function renderAll(){
  if (!SHARED || !SHARED.raw) return;
  renderBanner();
  renderQueue();
  renderRecent();
}

function renderBanner(){
  if (!CURRENT_WC){
    $('wcName').textContent = '—';
    $('wcDesc').textContent = 'select a work center to begin';
    $('statRunning').textContent = '0';
    $('statPaused').textContent = '0';
    $('statQueue').textContent = '0';
    return;
  }
  const wc = (SHARED.raw.WorkCenters || []).find(w => w.WorkCenter === CURRENT_WC);
  $('wcName').textContent = CURRENT_WC;
  $('wcDesc').textContent = wc ? (wc.Description || '') : '';

  const wos = getOrdersForWC();
  const orderState = window.FloorplanShared.deriveOrderLiveState(SHARED.confirmations || []);
  let running = 0, paused = 0, queue = 0;
  for (const w of wos){
    const s = orderState.get(w.OrderNumber);
    const status = s ? s.LiveStatus : 'IDLE';
    if (status === 'COMPLETE') continue;
    if (status === 'RUNNING') running++;
    else if (status === 'PAUSED') paused++;
    else queue++;
  }
  $('statRunning').textContent = running;
  $('statPaused').textContent = paused;
  $('statQueue').textContent = queue;
}

function getOrdersForWC(){
  const wos = (SHARED.raw.WorkOrders || []).filter(w => w.WorkCenter === CURRENT_WC);
  // Filter to non-TECO and with open quantity
  return wos.filter(w => {
    const status = String(w.Status || '').toUpperCase();
    if (status === 'TECO' || status === 'CLSD') return false;
    const open = (Number(w.OrderQty) || 0) - (Number(w.ConfirmedQty) || 0);
    return open > 0;
  });
}

function renderQueue(){
  const queue = $('queue');
  if (!CURRENT_WC){
    queue.innerHTML = '<div class="queue-empty">SELECT A WORK CENTER FROM THE DROPDOWN ABOVE</div>';
    return;
  }
  const wos = getOrdersForWC();
  const orderState = window.FloorplanShared.deriveOrderLiveState(SHARED.confirmations || []);
  const today = new Date(); today.setHours(0,0,0,0);

  // Sort: RUNNING first, then PAUSED, then by Priority asc + earliest finish
  const enriched = wos.map(w => {
    const s = orderState.get(w.OrderNumber);
    const status = s ? s.LiveStatus : 'IDLE';
    const actualConfirmed = (Number(w.ConfirmedQty) || 0) + (s ? s.ActualConfirmed : 0);
    const pct = w.OrderQty ? Math.min(actualConfirmed / w.OrderQty, 1) : 0;
    const finish = w.FinishDate ? (typeof w.FinishDate === 'string' ? new Date(w.FinishDate) : w.FinishDate) : null;
    const daysLeft = finish ? Math.round((finish - today) / 86400000) : null;
    let flag = 'ON TRACK';
    if (daysLeft !== null && daysLeft < 0) flag = 'PAST DUE';
    else if (daysLeft !== null && daysLeft <= 5) flag = 'AT RISK';
    return { ...w, _status: status, _actualConfirmed: actualConfirmed, _pct: pct, _daysLeft: daysLeft, _flag: flag };
  });

  enriched.sort((a, b) => {
    const order = { 'RUNNING': 0, 'PAUSED': 1, 'IDLE': 2, 'COMPLETE': 3 };
    if (order[a._status] !== order[b._status]) return order[a._status] - order[b._status];
    const aP = Number(a.Priority) || 3;
    const bP = Number(b.Priority) || 3;
    if (aP !== bP) return aP - bP;
    return (a._daysLeft ?? 999) - (b._daysLeft ?? 999);
  });

  if (enriched.length === 0){
    queue.innerHTML = '<div class="queue-empty">NO OPEN WORK ORDERS AT THIS WORK CENTER</div>';
    return;
  }

  queue.innerHTML = enriched.map(w => {
    const prio = Number(w.Priority) || 3;
    return `
      <div class="order-card" data-order="${escapeHtml(w.OrderNumber)}" data-status="${w._status}" data-flag="${w._flag}" tabindex="0">
        <div class="order-card__top">
          <div class="order-card__num">${escapeHtml(w.OrderNumber)}</div>
          <div class="order-card__prio" data-p="${prio}">P${prio}</div>
        </div>
        <div class="order-card__mat">${escapeHtml(w.Material)}</div>
        <div class="order-card__desc">${escapeHtml(w.Description || '')}</div>
        <div class="order-card__progress">
          <div class="order-card__progress-row">
            <span><b>${fmtNum(w._actualConfirmed)}</b> / ${fmtNum(w.OrderQty)}</span>
            <span>${(w._pct*100).toFixed(0)}%</span>
          </div>
          <div class="order-card__bar"><div class="order-card__bar-fill" style="width:${w._pct*100}%"></div></div>
        </div>
        <div class="order-card__bottom">
          <span>DUE ${fmtDate(w.FinishDate)}</span>
          <span>${w._daysLeft !== null && w._daysLeft < 0 ? Math.abs(w._daysLeft)+'D LATE' : (w._daysLeft !== null ? w._daysLeft+'D LEFT' : '')}</span>
        </div>
      </div>
    `;
  }).join('');
}

function renderRecent(){
  const list = $('recentList');
  const recent = (SHARED.confirmations || []).slice().reverse().slice(0, 10);
  if (recent.length === 0){
    list.innerHTML = '<li class="recent-empty">NO ACTIVITY YET — TAP AN ORDER TO BEGIN</li>';
    return;
  }
  list.innerHTML = recent.map(c => {
    const detail = c.Action === 'QUANTITY' ? `+${fmtNum(c.Quantity)} units`
                 : c.Action === 'SCRAP'    ? `${fmtNum(c.Quantity)} scrap`
                 : c.Action === 'DOWNTIME' ? `${fmtNum(c.Minutes)} min downtime`
                 : (c.Operator || '');
    return `
      <li>
        <span class="recent-time">${fmtTime(c.timestamp)}</span>
        <span class="recent-action">${escapeHtml(c.Action)}</span>
        <span class="recent-order">${escapeHtml(c.OrderNumber)}</span>
        <span class="recent-detail">${escapeHtml(detail)}${c.Operator && c.Action !== 'QUANTITY' && c.Action !== 'SCRAP' && c.Action !== 'DOWNTIME' ? '' : c.Operator ? ` · ${escapeHtml(c.Operator)}` : ''}</span>
      </li>
    `;
  }).join('');
}

/* ── DRAWER (action picker) ───────────────────────── */
function openDrawer(orderNumber){
  CURRENT_ORDER = orderNumber;
  const w = (SHARED.raw.WorkOrders || []).find(x => x.OrderNumber === orderNumber);
  if (!w) return;
  const orderState = window.FloorplanShared.deriveOrderLiveState(SHARED.confirmations || []);
  const s = orderState.get(orderNumber);
  const status = s ? s.LiveStatus : 'IDLE';
  const actualConfirmed = (Number(w.ConfirmedQty) || 0) + (s ? s.ActualConfirmed : 0);
  const scrap = s ? s.ActualScrap : 0;
  const pct = w.OrderQty ? Math.min(actualConfirmed / w.OrderQty, 1) : 0;

  $('drawerOrder').textContent = w.OrderNumber;
  $('drawerMaterial').textContent = `${w.Material} · ${w.Description || ''}`;
  $('drawerConfirmed').textContent = `${fmtNum(actualConfirmed)} / ${fmtNum(w.OrderQty)}`;
  $('drawerBar').style.width = (pct * 100) + '%';
  const sp = $('drawerStatus');
  sp.textContent = status;
  sp.className = 'pill ' + (
    status === 'RUNNING'  ? 'pill--ok' :
    status === 'PAUSED'   ? 'pill--warn' :
    status === 'COMPLETE' ? 'pill--mute' : 'pill--mute');
  $('drawerScrap').textContent = scrap > 0 ? `· ${fmtNum(scrap)} scrap` : '';

  // Toggle which action buttons are enabled based on current status
  const toggle = (action, enabled) => {
    const btn = document.querySelector(`.action-btn[data-action="${action}"]`);
    if (btn) btn.disabled = !enabled;
  };
  toggle('START',    status === 'IDLE');
  toggle('PAUSE',    status === 'RUNNING');
  toggle('RESUME',   status === 'PAUSED');
  toggle('QUANTITY', status === 'RUNNING' || status === 'PAUSED');
  toggle('SCRAP',    status === 'RUNNING' || status === 'PAUSED');
  toggle('DOWNTIME', status === 'RUNNING' || status === 'PAUSED');
  toggle('COMPLETE', status === 'RUNNING' || status === 'PAUSED' || status === 'IDLE');

  $('drawer').hidden = false;
  $('drawerBackdrop').hidden = false;
}

function closeDrawer(){
  $('drawer').hidden = true;
  $('drawerBackdrop').hidden = true;
  CURRENT_ORDER = null;
}

/* ── NUMPAD ───────────────────────────────────────── */
function openNumpad(title, hint, callback){
  $('numpadTitle').textContent = title;
  $('numpadHint').textContent = hint;
  NUMPAD_VALUE = '0';
  $('numpadDisplay').textContent = '0';
  NUMPAD_CALLBACK = callback;
  $('numpad').hidden = false;
  $('numpadBackdrop').hidden = false;
}
function closeNumpad(){
  $('numpad').hidden = true;
  $('numpadBackdrop').hidden = true;
  NUMPAD_CALLBACK = null;
}
function numpadKey(k){
  if (k === 'C'){ NUMPAD_VALUE = '0'; }
  else if (k === '←'){
    NUMPAD_VALUE = NUMPAD_VALUE.length > 1 ? NUMPAD_VALUE.slice(0,-1) : '0';
  }
  else if (NUMPAD_VALUE === '0'){
    NUMPAD_VALUE = k;
  }
  else if (NUMPAD_VALUE.length < 8){
    NUMPAD_VALUE += k;
  }
  $('numpadDisplay').textContent = NUMPAD_VALUE;
}

/* ── ACTIONS ──────────────────────────────────────── */
function getOperatorName(){
  return ($('operatorName').value || '').trim().toUpperCase() || 'UNASSIGNED';
}

function logConfirmation(action, extra = {}){
  if (!CURRENT_ORDER){ return; }
  if (!SHARED.confirmations) SHARED.confirmations = [];
  window.FloorplanShared.addConfirmation(SHARED, {
    OrderNumber: CURRENT_ORDER,
    WorkCenter: CURRENT_WC,
    Operator: getOperatorName(),
    Action: action,
    ...extra
  });
  window.FloorplanShared.save(SHARED);
  renderAll();
  toast(`${action} · ${CURRENT_ORDER}`);
}

function handleAction(action){
  if (!CURRENT_ORDER){ return; }
  switch(action){
    case 'START':
    case 'PAUSE':
    case 'RESUME':
      logConfirmation(action);
      // Re-open drawer to reflect new state (for stacking actions)
      openDrawer(CURRENT_ORDER);
      break;
    case 'COMPLETE':
      if (!confirm(`Mark order ${CURRENT_ORDER} as COMPLETE? This sets the order to TECO status.`)) return;
      // Mark the actual order in raw data as TECO
      const w = (SHARED.raw.WorkOrders || []).find(x => x.OrderNumber === CURRENT_ORDER);
      if (w) w.Status = 'TECO';
      logConfirmation('COMPLETE');
      closeDrawer();
      break;
    case 'QUANTITY':
      openNumpad('REPORT QUANTITY', 'units completed since last report', (n) => {
        if (n <= 0){ toast('Enter a quantity > 0', true); return; }
        logConfirmation('QUANTITY', { Quantity: n });
        openDrawer(CURRENT_ORDER);
      });
      break;
    case 'SCRAP':
      openNumpad('REPORT SCRAP', 'units scrapped', (n) => {
        if (n <= 0){ toast('Enter a quantity > 0', true); return; }
        logConfirmation('SCRAP', { Quantity: n });
        openDrawer(CURRENT_ORDER);
      });
      break;
    case 'DOWNTIME':
      openNumpad('REPORT DOWNTIME', 'minutes of downtime', (n) => {
        if (n <= 0){ toast('Enter minutes > 0', true); return; }
        logConfirmation('DOWNTIME', { Minutes: n });
        openDrawer(CURRENT_ORDER);
      });
      break;
  }
}

/* ── EVENT WIRING ─────────────────────────────────── */
function tickClock(){
  const d = new Date();
  const pad = n => String(n).padStart(2,'0');
  $('floorClock').textContent = `${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}

function init(){
  tickClock();
  setInterval(tickClock, 1000);

  $('wcSelect').addEventListener('change', e => {
    CURRENT_WC = e.target.value || null;
    try { sessionStorage.setItem('floorplan.lastWC', CURRENT_WC || ''); } catch(e){}
    renderAll();
  });

  // Restore last-selected WC for this tab
  try {
    const last = sessionStorage.getItem('floorplan.lastWC');
    if (last){ CURRENT_WC = last; }
  } catch(e){}

  // Persist operator name
  $('operatorName').addEventListener('input', e => {
    try { sessionStorage.setItem('floorplan.operator', e.target.value); } catch(err){}
  });
  try {
    const op = sessionStorage.getItem('floorplan.operator');
    if (op) $('operatorName').value = op;
  } catch(e){}

  // Click an order card → open drawer
  $('queue').addEventListener('click', e => {
    const card = e.target.closest('.order-card');
    if (card) openDrawer(card.dataset.order);
  });
  $('queue').addEventListener('keydown', e => {
    if (e.key === 'Enter' || e.key === ' '){
      const card = e.target.closest('.order-card');
      if (card){ e.preventDefault(); openDrawer(card.dataset.order); }
    }
  });

  // Drawer actions
  $('drawerClose').addEventListener('click', closeDrawer);
  $('drawerBackdrop').addEventListener('click', closeDrawer);
  document.querySelectorAll('.action-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      if (btn.disabled) return;
      handleAction(btn.dataset.action);
    });
  });

  // Numpad
  $('numpadClose').addEventListener('click', closeNumpad);
  $('numpadBackdrop').addEventListener('click', closeNumpad);
  document.querySelectorAll('.numpad__key').forEach(b => {
    b.addEventListener('click', () => numpadKey(b.dataset.k));
  });
  $('numpadConfirm').addEventListener('click', () => {
    const n = Number(NUMPAD_VALUE) || 0;
    const cb = NUMPAD_CALLBACK;
    closeNumpad();
    if (cb) cb(n);
  });

  // ESC closes modals
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape'){
      if (!$('numpad').hidden) closeNumpad();
      else if (!$('drawer').hidden) closeDrawer();
    }
  });

  // Sync from planner page edits / cross-tab
  window.FloorplanShared.onSync(msg => {
    refreshFromShared();
  });

  refreshFromShared();
}

document.addEventListener('DOMContentLoaded', init);

})();
