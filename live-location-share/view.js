/* =============================================================================
 * 現在地リンク共有 — 閲覧ページ
 *
 * リンクの ?e=（共有先URL）と ?id=（共有ID）だけで動く（閲覧者は設定不要）。
 * CORS 回避のため JSONP で現在地を取得し、Leaflet で地図表示する。
 * ========================================================================== */
'use strict';

var POLL_MS = 8000;

var params = new URLSearchParams(location.search);
var endpoint = params.get('e');
var shareId = params.get('id');

var map, marker, circle, trailLine;

function el(id) { return document.getElementById(id); }

function initMap() {
  map = L.map('map', { zoomControl: true }).setView([35.681236, 139.767125], 15); // 既定: 東京駅
  L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
    maxZoom: 19,
    attribution: '&copy; OpenStreetMap contributors'
  }).addTo(map);
}

// JSONP: <script> を差し込んで callback で受け取る（CORS 不要）
function fetchJsonp(url) {
  return new Promise(function (resolve, reject) {
    var cb = '_lls' + Math.random().toString(36).slice(2);
    var s = document.createElement('script');
    var timer = setTimeout(function () { cleanup(); reject(new Error('timeout')); }, 15000);
    window[cb] = function (data) { cleanup(); resolve(data); };
    function cleanup() { clearTimeout(timer); delete window[cb]; s.remove(); }
    s.onerror = function () { cleanup(); reject(new Error('network')); };
    s.src = url + (url.indexOf('?') >= 0 ? '&' : '?') + 'callback=' + cb;
    document.body.appendChild(s);
  });
}

function agoText(iso) {
  var sec = Math.max(0, Math.floor((Date.now() - Date.parse(iso)) / 1000));
  if (sec < 60) return sec + '秒前';
  if (sec < 3600) return Math.floor(sec / 60) + '分前';
  return Math.floor(sec / 3600) + '時間前';
}

function renderLive(d) {
  el('who').textContent = (d.name || '相手') + ' の現在地';
  el('sub').textContent = '最終更新: ' + agoText(d.updatedAt) +
    (d.accuracy ? '（精度 ±' + d.accuracy + 'm）' : '');

  if (typeof d.lat !== 'number' || typeof d.lng !== 'number') return;
  var latlng = [d.lat, d.lng];

  if (!marker) {
    marker = L.marker(latlng).addTo(map);
    circle = L.circle(latlng, { radius: d.accuracy || 30, color: '#2f81f7', fillOpacity: 0.12 }).addTo(map);
    map.setView(latlng, 16);
  } else {
    marker.setLatLng(latlng);
    circle.setLatLng(latlng).setRadius(d.accuracy || 30);
    map.panTo(latlng);
  }

  // 移動経路（直近の点列）
  if (Array.isArray(d.trail) && d.trail.length > 1) {
    var pts = d.trail.map(function (t) { return [t[0], t[1]]; });
    if (!trailLine) trailLine = L.polyline(pts, { color: '#2f81f7', weight: 3, opacity: 0.6 }).addTo(map);
    else trailLine.setLatLngs(pts);
  }
}

function renderEnded(name) {
  el('who').textContent = (name || '相手') + ' の共有は終了しました';
  el('sub').innerHTML = '<span class="ended">現在は位置を共有していません。</span>';
}

function renderNotFound() {
  el('who').textContent = 'リンクが見つかりません';
  el('sub').textContent = 'リンクが正しいか、共有が開始されているか確認してください。';
}

async function poll() {
  try {
    var url = endpoint + '?id=' + encodeURIComponent(shareId);
    var d = await fetchJsonp(url);
    if (!d || !d.ok) { el('sub').textContent = '取得に失敗しました。再試行します…'; return; }
    if (d.status === 'live') renderLive(d);
    else if (d.status === 'ended') renderEnded(d.name);
    else renderNotFound();
  } catch (e) {
    el('sub').textContent = '通信エラー。再試行します…';
  }
}

function start() {
  if (!endpoint || !shareId) {
    el('who').textContent = 'リンクが不正です';
    el('sub').textContent = 'e と id のパラメータが必要です。';
    return;
  }
  initMap();
  poll();
  setInterval(poll, POLL_MS);
}

document.addEventListener('DOMContentLoaded', start);
