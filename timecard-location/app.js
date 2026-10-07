/* =============================================================================
 * 勤怠・位置記録アプリ（同意ベース）
 *
 * 設計上の約束:
 *   - 記録するのは「勤務開始」〜「勤務終了」の間だけ。OFF の間は何もしない。
 *   - 記録中は常にバナーを表示し、送信内容を本人に見せる（隠れて動かない）。
 *   - 位置情報はブラウザの許可ダイアログを経てのみ取得する。
 *   - 設定・ログ・未送信キューはこの端末の localStorage にのみ保持する。
 * ========================================================================== */

'use strict';

const LS = {
  config: 'tcl.config',
  session: 'tcl.session',   // 勤務中セッション（復帰用）
  queue: 'tcl.queue',       // 未送信キュー
  log: 'tcl.log',           // 画面表示用ログ
};

const $ = (id) => document.getElementById(id);

const state = {
  config: null,
  onDuty: false,
  sessionId: null,
  startedAt: null,
  tickTimer: null,          // 定期記録タイマー
  clockTimer: null,         // 経過時間表示タイマー
};

/* ------------------------------ 永続化ヘルパ ------------------------------ */

function load(key, fallback) {
  try {
    const v = localStorage.getItem(key);
    return v == null ? fallback : JSON.parse(v);
  } catch {
    return fallback;
  }
}
function save(key, value) {
  try { localStorage.setItem(key, JSON.stringify(value)); } catch {}
}
function remove(key) {
  try { localStorage.removeItem(key); } catch {}
}

/* ------------------------------ 画面の切替 -------------------------------- */

function show(view) {
  for (const id of ['setupView', 'consentView', 'onDutyView']) {
    $(id).classList.toggle('hidden', id !== view);
  }
}

function render() {
  if (!state.config) { show('setupView'); return; }
  if (state.onDuty) { show('onDutyView'); return; }
  show('consentView');
}

/* ------------------------------ 初期設定 ---------------------------------- */

function fillSetupForm() {
  const c = state.config || {};
  $('cfgEndpoint').value = c.endpoint || '';
  $('cfgWorkerId').value = c.workerId || '';
  $('cfgWorkerName').value = c.workerName || '';
  $('cfgInterval').value = c.intervalSec || 300;
  $('cfgGeocode').checked = !!c.geocode;
}

function saveSetup() {
  const endpoint = $('cfgEndpoint').value.trim();
  const workerId = $('cfgWorkerId').value.trim();
  const workerName = $('cfgWorkerName').value.trim();
  const intervalSec = Math.max(60, parseInt($('cfgInterval').value, 10) || 300);
  const geocode = $('cfgGeocode').checked;

  if (!/^https:\/\/script\.google\.com\/macros\/s\/.+\/exec/.test(endpoint)) {
    alert('記録先 URL は Apps Script のウェブアプリ URL（.../exec）を入力してください。');
    return;
  }
  if (!workerId || !workerName) {
    alert('従業員ID と氏名を入力してください。');
    return;
  }

  state.config = { endpoint, workerId, workerName, intervalSec, geocode };
  save(LS.config, state.config);
  updateConsentCopy();
  render();
}

function updateConsentCopy() {
  const c = state.config || {};
  $('consentInterval').textContent = Math.round((c.intervalSec || 300) / 60);
  $('consentAddr').style.display = c.geocode ? '' : 'none';
}

/* ------------------------------ 勤務ON/OFF -------------------------------- */

function goOnDuty() {
  if (!state.config) return;

  // ブラウザの位置情報許可を先に取得（拒否なら勤務開始しない）
  navigator.geolocation.getCurrentPosition(
    (pos) => {
      state.onDuty = true;
      state.sessionId = 's-' + Date.now();
      state.startedAt = Date.now();
      save(LS.session, { sessionId: state.sessionId, startedAt: state.startedAt });

      $('onWorker').textContent = `${state.config.workerName}（${state.config.workerId}）`;
      $('recordingBanner').classList.remove('hidden');
      render();

      recordPoint('clock_in', pos);     // 開始地点を記録
      startTimers();
      updateClock();
    },
    (err) => {
      alert('位置情報を取得できませんでした。ブラウザの位置情報を許可してください。\n（' + err.message + '）');
    },
    { enableHighAccuracy: true, timeout: 20000, maximumAge: 0 }
  );
}

function goOffDuty() {
  // 終了地点を記録してから停止
  navigator.geolocation.getCurrentPosition(
    (pos) => finishDuty(pos),
    () => finishDuty(null),
    { enableHighAccuracy: true, timeout: 15000, maximumAge: 0 }
  );
}

function finishDuty(pos) {
  if (pos) recordPoint('clock_out', pos);
  else recordPoint('clock_out', null);

  stopTimers();
  state.onDuty = false;
  state.sessionId = null;
  state.startedAt = null;
  remove(LS.session);
  $('recordingBanner').classList.add('hidden');
  render();
}

function startTimers() {
  stopTimers();
  state.tickTimer = setInterval(() => {
    navigator.geolocation.getCurrentPosition(
      (pos) => recordPoint('ping', pos),
      (err) => addLog({ type: 'ping', error: err.message, time: Date.now() }, 'pending'),
      { enableHighAccuracy: true, timeout: 20000, maximumAge: 0 }
    );
  }, state.config.intervalSec * 1000);

  state.clockTimer = setInterval(updateClock, 1000);
}

function stopTimers() {
  if (state.tickTimer) clearInterval(state.tickTimer);
  if (state.clockTimer) clearInterval(state.clockTimer);
  state.tickTimer = null;
  state.clockTimer = null;
}

function updateClock() {
  if (!state.startedAt) return;
  const s = Math.floor((Date.now() - state.startedAt) / 1000);
  const hh = String(Math.floor(s / 3600)).padStart(2, '0');
  const mm = String(Math.floor((s % 3600) / 60)).padStart(2, '0');
  const ss = String(s % 60).padStart(2, '0');
  $('elapsed').textContent = `${hh}:${mm}:${ss}`;
}

/* ------------------------------ 位置の記録 -------------------------------- */

async function recordPoint(type, pos) {
  const now = Date.now();
  const point = {
    workerId: state.config.workerId,
    workerName: state.config.workerName,
    sessionId: state.sessionId,
    type,                                  // clock_in | ping | clock_out
    timestamp: new Date(now).toISOString(),
    lat: pos ? round(pos.coords.latitude, 6) : null,
    lng: pos ? round(pos.coords.longitude, 6) : null,
    accuracy: pos ? Math.round(pos.coords.accuracy) : null,
    address: '',
  };

  // 画面表示を更新
  if (pos) {
    $('lastTime').textContent = fmtTime(now);
    $('lastPos').textContent = `${point.lat}, ${point.lng}`;
    $('lastAcc').textContent = `±${point.accuracy} m`;
  }

  // 任意: 逆ジオコーディング（住所）。設定ONのときだけ。
  if (state.config.geocode && pos) {
    try { point.address = await reverseGeocode(point.lat, point.lng); } catch {}
  }

  enqueueAndFlush(point);
}

function round(n, d) {
  const f = Math.pow(10, d);
  return Math.round(n * f) / f;
}

/* ------------------------- 送信（キュー＋リトライ） ------------------------ */

function enqueueAndFlush(point) {
  const q = load(LS.queue, []);
  q.push(point);
  save(LS.queue, q);
  $('pendingCount').textContent = q.length;
  addLog(point, 'pending');
  flushQueue();
}

let flushing = false;
async function flushQueue() {
  if (flushing) return;
  flushing = true;
  try {
    let q = load(LS.queue, []);
    while (q.length) {
      const point = q[0];
      const ok = await sendPoint(point);
      if (!ok) break;                 // 失敗したら次回に持ち越し
      q.shift();
      save(LS.queue, q);
      $('pendingCount').textContent = q.length;
      markLogSent(point);
    }
  } finally {
    flushing = false;
  }
}

async function sendPoint(point) {
  try {
    // Apps Script のウェブアプリへ送信。
    // text/plain にして CORS プリフライトを避ける（Apps Script の定石）。
    const res = await fetch(state.config.endpoint, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(point),
      redirect: 'follow',
    });
    return res.ok;
  } catch {
    return false;   // オフライン等。キューに残して後で再送。
  }
}

/* ------------------------------ 逆ジオコーディング ------------------------ *
 * OpenStreetMap / Nominatim を利用（無料）。
 * 【注意】Nominatim は本番の大量利用に不向き（1 req/s 上限・商用制限あり）。
 *         恒常運用では Google Maps / Mapbox 等の正規サービスに差し替えてください。
 * -------------------------------------------------------------------------- */
async function reverseGeocode(lat, lng) {
  const url = `https://nominatim.openstreetmap.org/reverse?format=jsonv2&lat=${lat}&lon=${lng}&accept-language=ja`;
  const res = await fetch(url, { headers: { 'Accept': 'application/json' } });
  if (!res.ok) return '';
  const data = await res.json();
  return data.display_name || '';
}

/* ------------------------------ 画面ログ ---------------------------------- */

function addLog(point, status) {
  const log = load(LS.log, []);
  log.unshift({ ...point, status, _id: point.timestamp + ':' + point.type });
  save(LS.log, log.slice(0, 100));
  renderLog();
}
function markLogSent(point) {
  const log = load(LS.log, []);
  const id = point.timestamp + ':' + point.type;
  for (const e of log) if (e._id === id) e.status = 'ok';
  save(LS.log, log);
  renderLog();
}
function renderLog() {
  const log = load(LS.log, []);
  $('log').innerHTML = log.map((e) => {
    const label = { clock_in: '勤務開始', ping: '記録', clock_out: '勤務終了' }[e.type] || e.type;
    const pos = (e.lat != null)
      ? `<a href="https://www.google.com/maps?q=${e.lat},${e.lng}" target="_blank" rel="noopener">地図</a>`
      : (e.error ? '取得失敗' : '—');
    const st = e.status === 'ok'
      ? '<span class="k ok">送信済</span>'
      : '<span class="k pending">未送信</span>';
    const addr = e.address ? `<div>${escapeHtml(e.address)}</div>` : '';
    return `<li>
      <span class="t">${fmtTime(Date.parse(e.timestamp))}</span>
      <span class="k">${label}</span> ${st}
      <div>${pos}${e.accuracy != null ? ` · ±${e.accuracy}m` : ''}</div>
      ${addr}
    </li>`;
  }).join('');
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}
function fmtTime(ms) {
  const d = new Date(ms);
  const p = (n) => String(n).padStart(2, '0');
  return `${p(d.getMonth() + 1)}/${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
}

/* ------------------------------ リセット ---------------------------------- */

function resetAll() {
  if (!confirm('設定・ログ・未送信キューをこの端末から消去します。よろしいですか？')) return;
  stopTimers();
  for (const k of Object.values(LS)) remove(k);
  state.config = null;
  state.onDuty = false;
  state.sessionId = null;
  state.startedAt = null;
  $('recordingBanner').classList.add('hidden');
  fillSetupForm();
  renderLog();
  render();
}

/* ------------------------------ 起動処理 ---------------------------------- */

function init() {
  state.config = load(LS.config, null);

  // イベント配線
  $('saveSetup').addEventListener('click', saveSetup);
  $('editSetup').addEventListener('click', () => { fillSetupForm(); show('setupView'); });
  $('consentCheck').addEventListener('change', (e) => {
    $('goOnDuty').disabled = !e.target.checked;
  });
  $('goOnDuty').addEventListener('click', goOnDuty);
  $('goOffDuty').addEventListener('click', goOffDuty);
  $('sendNow').addEventListener('click', () => {
    navigator.geolocation.getCurrentPosition(
      (pos) => recordPoint('ping', pos),
      (err) => alert('位置を取得できませんでした: ' + err.message),
      { enableHighAccuracy: true, timeout: 20000, maximumAge: 0 }
    );
  });
  $('resetAll').addEventListener('click', resetAll);

  // ネットワーク復帰時に未送信を再送
  window.addEventListener('online', flushQueue);

  fillSetupForm();
  updateConsentCopy();
  renderLog();

  // 勤務中にリロードされた場合でも、記録は「本人の再操作」で再開する。
  // （バックグラウンドでこっそり続けない＝同意と透明性を優先）
  remove(LS.session);

  render();

  // ブラウザ非対応チェック
  if (!('geolocation' in navigator)) {
    alert('このブラウザは位置情報に対応していません。');
  }
}

document.addEventListener('DOMContentLoaded', init);
