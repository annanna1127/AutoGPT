/* =============================================================================
 * 現在地リンク共有 — 共有者ページ
 *
 * 約束:
 *   - 自分が「共有を開始」した時だけ送信。許可しなければ何も起きない。
 *   - 共有中はバナーを常時表示。タブを閉じると送信は止まる（裏で動き続けない）。
 *   - 設定・共有IDはこの端末の localStorage にのみ保持。
 * ========================================================================== */
'use strict';

var LS = { cfg: 'lls.cfg', share: 'lls.share' };
var $ = function (id) { return document.getElementById(id); };

var S = {
  cfg: null,         // { endpoint, name }
  shareId: null,
  startedAt: null,
  expiresAt: null,
  watchId: null,
  lastSent: 0,
  timer: null,
};

var SEND_INTERVAL_MS = 8000; // 送信の最短間隔

/* ------------------------------ 保存ヘルパ ------------------------------ */
function load(k, d) { try { var v = localStorage.getItem(k); return v == null ? d : JSON.parse(v); } catch (e) { return d; } }
function save(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) {} }
function del(k) { try { localStorage.removeItem(k); } catch (e) {} }

/* ------------------------------ 画面切替 -------------------------------- */
function show(view) {
  ['setup', 'idle', 'sharing'].forEach(function (id) {
    $(id).classList.toggle('hidden', id !== view);
  });
}
function render() {
  if (!S.cfg) { show('setup'); return; }
  if (S.watchId != null) { show('sharing'); return; }
  show('idle');
}

/* ------------------------------ 設定 ------------------------------------ */
function saveSetup() {
  var endpoint = $('endpoint').value.trim();
  var name = $('name').value.trim();
  if (!/^https:\/\/script\.google\.com\/macros\/s\/.+\/exec/.test(endpoint)) {
    alert('共有先 URL は Apps Script のウェブアプリ URL（.../exec）を入力してください。');
    return;
  }
  if (!name) { alert('表示名を入力してください。'); return; }
  S.cfg = { endpoint: endpoint, name: name };
  save(LS.cfg, S.cfg);
  render();
}

/* ------------------------------ 共有開始/停止 --------------------------- */
function startShare() {
  if (!S.cfg) return;
  // 先に位置情報の許可を取る
  navigator.geolocation.getCurrentPosition(function (pos) {
    S.shareId = (crypto.randomUUID ? crypto.randomUUID().slice(0, 8) : String(Date.now()).slice(-8));
    S.startedAt = Date.now();
    var hours = parseInt($('ttl').value, 10) || 4;
    S.expiresAt = new Date(S.startedAt + hours * 3600000);
    save(LS.share, { shareId: S.shareId, startedAt: S.startedAt, expiresAt: S.expiresAt.toISOString() });

    $('who').textContent = S.cfg.name;
    var url = viewerLink();
    $('link').value = url;
    $('preview').href = url;
    $('expires').textContent = fmt(S.expiresAt);
    $('banner').classList.remove('hidden');

    sendPosition(pos, true);               // 開始位置を即送信

    // リアルタイム追従（Zenly 風）。送信は最短間隔でスロットル。
    S.watchId = navigator.geolocation.watchPosition(
      function (p) { sendPosition(p, false); },
      function (err) { console.warn(err); },
      { enableHighAccuracy: true, maximumAge: 0, timeout: 20000 }
    );
    S.timer = setInterval(tickClock, 1000);
    render();
    tickClock();
  }, function (err) {
    alert('位置情報を取得できませんでした。ブラウザの位置情報を許可してください。\n（' + err.message + '）');
  }, { enableHighAccuracy: true, timeout: 20000, maximumAge: 0 });
}

function stopShare(viaBeacon) {
  if (S.watchId != null) navigator.geolocation.clearWatch(S.watchId);
  if (S.timer) clearInterval(S.timer);
  S.watchId = null; S.timer = null;

  // 相手の画面を「共有終了」にするため停止を通知
  if (S.shareId) {
    var body = JSON.stringify({ shareId: S.shareId, stop: true });
    if (viaBeacon && navigator.sendBeacon) {
      navigator.sendBeacon(S.cfg.endpoint, new Blob([body], { type: 'text/plain;charset=utf-8' }));
    } else {
      fetch(S.cfg.endpoint, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: body }).catch(function () {});
    }
  }
  S.shareId = null; S.startedAt = null; S.expiresAt = null;
  del(LS.share);
  $('banner').classList.add('hidden');
  render();
}

/* ------------------------------ 送信 ------------------------------------ */
function sendPosition(pos, force) {
  var now = Date.now();
  if (!force && now - S.lastSent < SEND_INTERVAL_MS) return;  // スロットル
  S.lastSent = now;

  if (S.expiresAt && now > S.expiresAt.getTime()) { stopShare(false); return; }

  var payload = {
    shareId: S.shareId,
    name: S.cfg.name,
    lat: round(pos.coords.latitude, 6),
    lng: round(pos.coords.longitude, 6),
    accuracy: Math.round(pos.coords.accuracy),
    expiresAt: S.expiresAt.toISOString(),
  };
  fetch(S.cfg.endpoint, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(payload),
  }).then(function () {
    $('updated').textContent = fmt(new Date());
    $('acc').textContent = '±' + payload.accuracy + ' m';
  }).catch(function () {
    $('updated').textContent = '送信失敗（通信待ち）';
  });
}

/* ------------------------------ 補助 ------------------------------------ */
function viewerLink() {
  var base = location.href.replace(/index\.html(\?.*)?$/, '').replace(/\/?$/, '/');
  return base + 'view.html?e=' + encodeURIComponent(S.cfg.endpoint) + '&id=' + encodeURIComponent(S.shareId);
}
function tickClock() {
  if (!S.startedAt) return;
  var s = Math.floor((Date.now() - S.startedAt) / 1000);
  var m = String(Math.floor(s / 60)).padStart(2, '0');
  var ss = String(s % 60).padStart(2, '0');
  $('elapsed').textContent = m + ':' + ss;
  if (S.expiresAt && Date.now() > S.expiresAt.getTime()) stopShare(false);
}
function fmt(d) {
  var p = function (n) { return String(n).padStart(2, '0'); };
  return p(d.getHours()) + ':' + p(d.getMinutes()) + ':' + p(d.getSeconds());
}
function round(n, d) { var f = Math.pow(10, d); return Math.round(n * f) / f; }

function resetAll() {
  if (!confirm('設定を消去します。よろしいですか？')) return;
  if (S.watchId != null) stopShare(false);
  del(LS.cfg); del(LS.share);
  S.cfg = null;
  $('endpoint').value = ''; $('name').value = '';
  render();
}

/* ------------------------------ 起動 ------------------------------------ */
function init() {
  S.cfg = load(LS.cfg, null);
  if (S.cfg) { $('endpoint').value = S.cfg.endpoint; $('name').value = S.cfg.name; }

  $('saveSetup').addEventListener('click', saveSetup);
  $('editSetup').addEventListener('click', function () { show('setup'); });
  $('start').addEventListener('click', startShare);
  $('stop').addEventListener('click', function () { stopShare(false); });
  $('reset').addEventListener('click', resetAll);
  $('copy').addEventListener('click', function () {
    $('link').select();
    navigator.clipboard && navigator.clipboard.writeText($('link').value);
    $('copy').textContent = 'コピー済';
    setTimeout(function () { $('copy').textContent = 'コピー'; }, 1500);
  });

  // タブを閉じる/離れる時は停止を通知（裏で共有し続けない）
  window.addEventListener('pagehide', function () { if (S.watchId != null) stopShare(true); });

  // 共有中にリロードされても自動再開しない（本人の再操作を必須にする）
  del(LS.share);

  if (!('geolocation' in navigator)) alert('このブラウザは位置情報に対応していません。');
  render();
}
document.addEventListener('DOMContentLoaded', init);
