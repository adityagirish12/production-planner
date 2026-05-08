/* ──────────────────────────────────────────────────────────
   FLOORPLAN · Shared data layer
   Used by both index.html (planner) and floor.html (operator)
   ────────────────────────────────────────────────────────── */
(function(){
'use strict';

const STORAGE_KEY = 'floorplan.shared.v1';
const CHANNEL_NAME = 'floorplan-sync';

/* ── Persistence ────────────────────────────────────────── */
function loadShared(){
  try{
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    // revive dates in confirmations
    if (parsed.confirmations){
      parsed.confirmations.forEach(c => { if (c.timestamp) c.timestamp = new Date(c.timestamp); });
    }
    return parsed;
  } catch(e){
    console.warn('Failed to load shared state:', e);
    return null;
  }
}

function saveShared(state){
  try{
    // Serialize dates as ISO strings
    const serializable = JSON.parse(JSON.stringify(state, (key, val) => {
      if (val instanceof Date) return val.toISOString();
      return val;
    }));
    localStorage.setItem(STORAGE_KEY, JSON.stringify(serializable));
    broadcast({ type: 'state-updated' });
  } catch(e){
    console.error('Failed to save shared state:', e);
  }
}

function clearShared(){
  localStorage.removeItem(STORAGE_KEY);
  broadcast({ type: 'state-cleared' });
}

/* ── Cross-tab sync via BroadcastChannel ──────────────── */
let channel = null;
const listeners = [];
function getChannel(){
  if (channel) return channel;
  if (typeof BroadcastChannel === 'undefined') return null;
  channel = new BroadcastChannel(CHANNEL_NAME);
  channel.onmessage = (msg) => {
    listeners.forEach(fn => { try { fn(msg.data); } catch(e){ console.error(e); } });
  };
  return channel;
}
function broadcast(msg){
  const ch = getChannel();
  if (ch) ch.postMessage(msg);
}
function onSync(fn){
  listeners.push(fn);
  getChannel(); // ensure channel is initialized
  // Also listen for cross-tab via storage events (fallback for browsers without BC)
  window.addEventListener('storage', e => {
    if (e.key === STORAGE_KEY){ fn({ type: 'state-updated' }); }
  });
}

/* ── Confirmation log helpers ──────────────────────────── */
function addConfirmation(state, conf){
  if (!state.confirmations) state.confirmations = [];
  state.confirmations.push({
    id: 'conf_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6),
    timestamp: new Date(),
    ...conf
  });
}

/**
 * Walk confirmations and produce derived per-order state:
 *   - LiveStatus: IDLE | RUNNING | PAUSED | COMPLETE
 *   - LiveOperator: last operator who touched it
 *   - ActualConfirmed: sum of QUANTITY confirmations
 *   - ActualScrap: sum of SCRAP confirmations
 *   - LastEvent: timestamp of most recent confirmation
 */
function deriveOrderLiveState(confirmations){
  const byOrder = new Map();
  // Sort by timestamp ascending so we replay events in order
  const sorted = (confirmations || []).slice().sort((a,b) => {
    const at = a.timestamp instanceof Date ? a.timestamp.getTime() : new Date(a.timestamp).getTime();
    const bt = b.timestamp instanceof Date ? b.timestamp.getTime() : new Date(b.timestamp).getTime();
    return at - bt;
  });
  for (const c of sorted){
    if (!c.OrderNumber) continue;
    let s = byOrder.get(c.OrderNumber);
    if (!s){
      s = { LiveStatus: 'IDLE', LiveOperator: null, ActualConfirmed: 0, ActualScrap: 0,
            LastEvent: null, DowntimeMinutes: 0, EventCount: 0 };
      byOrder.set(c.OrderNumber, s);
    }
    s.EventCount++;
    s.LastEvent = c.timestamp;
    if (c.Operator) s.LiveOperator = c.Operator;
    switch ((c.Action || '').toUpperCase()){
      case 'START':    s.LiveStatus = 'RUNNING'; break;
      case 'PAUSE':    s.LiveStatus = 'PAUSED'; break;
      case 'RESUME':   s.LiveStatus = 'RUNNING'; break;
      case 'COMPLETE': s.LiveStatus = 'COMPLETE'; break;
      case 'QUANTITY': s.ActualConfirmed += Number(c.Quantity) || 0; break;
      case 'SCRAP':    s.ActualScrap     += Number(c.Quantity) || 0; break;
      case 'DOWNTIME': s.DowntimeMinutes += Number(c.Minutes)  || 0; break;
    }
  }
  return byOrder;
}

/* Per-WC live aggregates */
function deriveWCLiveState(confirmations, wos){
  const orderState = deriveOrderLiveState(confirmations);
  const byWC = new Map();
  for (const w of (wos || [])){
    const wc = w.WorkCenter;
    if (!wc) continue;
    let s = byWC.get(wc);
    if (!s){ s = { Running: 0, Paused: 0, Idle: 0, DowntimeMinutes: 0, RunningOrders: [] }; byWC.set(wc, s); }
    const os = orderState.get(w.OrderNumber);
    const status = os ? os.LiveStatus : 'IDLE';
    if (status === 'RUNNING'){ s.Running++; s.RunningOrders.push(w.OrderNumber); }
    else if (status === 'PAUSED') s.Paused++;
    else s.Idle++;
    if (os) s.DowntimeMinutes += os.DowntimeMinutes;
  }
  return byWC;
}

/* ── Public API ─────────────────────────────────────────── */
window.FloorplanShared = {
  STORAGE_KEY,
  load: loadShared,
  save: saveShared,
  clear: clearShared,
  onSync,
  addConfirmation,
  deriveOrderLiveState,
  deriveWCLiveState
};

})();
