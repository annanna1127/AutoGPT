/* =============================================================================
 * 勤怠・位置記録 — Google Apps Script バックエンド
 *
 * 役割:
 *   - Web アプリ本体（index.html）からの POST を受け、「記録」シートに1行追記する。
 *   - メニューから「日次サマリー」を生成し、セッションごとの勤務時間・滞在地点・
 *     滞在時間・移動経路を計算する。
 *
 * セットアップ:
 *   1. Google スプレッドシートを新規作成。
 *   2. 拡張機能 → Apps Script を開き、このコードを貼り付けて保存。
 *   3. デプロイ → 新しいデプロイ → 種類「ウェブアプリ」
 *        - 実行ユーザー: 自分
 *        - アクセスできるユーザー: 全員（本人の端末から送信できるように）
 *   4. 発行された .../exec の URL を Web アプリ本体の「記録先 URL」に設定。
 *
 * プライバシー運用の注意:
 *   - アクセスできるユーザーを「全員」にすると URL を知る人が POST できます。
 *     必要に応じて SHARED_TOKEN を設定し、本体からも同じ値を送るよう改修してください。
 *   - 記録データは個人情報です。閲覧権限・保管期間・利用目的を規程に従って管理してください。
 * ========================================================================== */

var SHEET_LOG = '記録';
var SHEET_SUMMARY = 'サマリー';

// 滞在とみなす距離のしきい値（メートル）。この範囲内の連続点を同一滞在とみなす。
var STAY_RADIUS_M = 150;
// 滞在として集計する最小滞在時間（分）。これ未満は「通過」として無視。
var STAY_MIN_MINUTES = 5;

var LOG_HEADERS = [
  '受信時刻', 'timestamp', 'workerId', 'workerName', 'sessionId',
  'type', 'lat', 'lng', 'accuracy(m)', 'address'
];

/* ----------------------------- POST 受信 -------------------------------- */

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = getOrCreateSheet(SHEET_LOG, LOG_HEADERS);
    sheet.appendRow([
      new Date(),
      data.timestamp || '',
      data.workerId || '',
      data.workerName || '',
      data.sessionId || '',
      data.type || '',
      data.lat != null ? data.lat : '',
      data.lng != null ? data.lng : '',
      data.accuracy != null ? data.accuracy : '',
      data.address || ''
    ]);
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

// 動作確認用（ブラウザで URL を開くと表示される）
function doGet() {
  return json({ ok: true, service: 'timecard-location', time: new Date().toISOString() });
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* --------------------------- メニュー登録 ------------------------------- */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('勤怠集計')
    .addItem('日次サマリーを作成', 'buildSummary')
    .addToUi();
}

/* --------------------------- サマリー生成 ------------------------------- *
 * セッション（clock_in 〜 clock_out）ごとに:
 *   - 勤務時間（開始〜終了）
 *   - 移動経路の点数と総移動距離
 *   - 滞在地点（半径 STAY_RADIUS_M 内の連続点をまとめ、滞在時間を算出）
 * ------------------------------------------------------------------------ */

function buildSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = ss.getSheetByName(SHEET_LOG);
  if (!log || log.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('「記録」シートにデータがありません。');
    return;
  }

  var rows = log.getRange(2, 1, log.getLastRow() - 1, LOG_HEADERS.length).getValues();

  // sessionId ごとにまとめる
  var sessions = {};
  rows.forEach(function (r) {
    var p = {
      ts: parseTs(r[1]),
      workerId: r[2],
      workerName: r[3],
      sessionId: r[4],
      type: r[5],
      lat: typeof r[6] === 'number' ? r[6] : null,
      lng: typeof r[7] === 'number' ? r[7] : null,
      address: r[9]
    };
    if (!p.sessionId || !p.ts) return;
    (sessions[p.sessionId] = sessions[p.sessionId] || []).push(p);
  });

  var out = [[
    '日付', 'workerId', '氏名', 'sessionId',
    '勤務開始', '勤務終了', '勤務時間(h)',
    '記録点数', '総移動距離(km)',
    '滞在地点数', '滞在内訳（住所/座標 : 滞在分）'
  ]];

  Object.keys(sessions).forEach(function (sid) {
    var pts = sessions[sid].filter(function (p) { return p.lat != null && p.lng != null; });
    pts.sort(function (a, b) { return a.ts - b.ts; });
    if (!pts.length) return;

    var start = pts[0].ts, end = pts[pts.length - 1].ts;
    var workHours = (end - start) / 3600000;

    // 総移動距離
    var dist = 0;
    for (var i = 1; i < pts.length; i++) {
      dist += haversine(pts[i - 1].lat, pts[i - 1].lng, pts[i].lat, pts[i].lng);
    }

    // 滞在地点の抽出（連続点が半径内にとどまる区間）
    var stays = extractStays(pts);
    var stayText = stays.map(function (s) {
      var label = s.address || (round(s.lat, 5) + ',' + round(s.lng, 5));
      return label + ' : ' + Math.round(s.minutes) + '分';
    }).join('\n');

    out.push([
      Utilities.formatDate(start, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      pts[0].workerId,
      pts[0].workerName,
      sid,
      Utilities.formatDate(start, Session.getScriptTimeZone(), 'HH:mm'),
      Utilities.formatDate(end, Session.getScriptTimeZone(), 'HH:mm'),
      round(workHours, 2),
      pts.length,
      round(dist / 1000, 2),
      stays.length,
      stayText
    ]);
  });

  var summary = getOrCreateSheet(SHEET_SUMMARY, out[0]);
  summary.clearContents();
  summary.getRange(1, 1, out.length, out[0].length).setValues(out);
  summary.setFrozenRows(1);
  summary.getRange(1, 1, 1, out[0].length).setFontWeight('bold');

  SpreadsheetApp.getUi().alert('サマリーを作成しました（' + (out.length - 1) + ' セッション）。');
}

/* 連続する点をまとめて滞在区間を作る */
function extractStays(pts) {
  var stays = [];
  var cluster = [pts[0]];

  function flush() {
    if (cluster.length < 2) return;
    var minutes = (cluster[cluster.length - 1].ts - cluster[0].ts) / 60000;
    if (minutes < STAY_MIN_MINUTES) return;
    // 重心
    var lat = 0, lng = 0;
    cluster.forEach(function (p) { lat += p.lat; lng += p.lng; });
    var addr = '';
    for (var i = 0; i < cluster.length; i++) { if (cluster[i].address) { addr = cluster[i].address; break; } }
    stays.push({
      lat: lat / cluster.length,
      lng: lng / cluster.length,
      minutes: minutes,
      address: addr
    });
  }

  for (var i = 1; i < pts.length; i++) {
    var d = haversine(cluster[0].lat, cluster[0].lng, pts[i].lat, pts[i].lng);
    if (d <= STAY_RADIUS_M) {
      cluster.push(pts[i]);
    } else {
      flush();
      cluster = [pts[i]];
    }
  }
  flush();
  return stays;
}

/* ------------------------------ ユーティリティ -------------------------- */

function getOrCreateSheet(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sh;
}

function parseTs(v) {
  if (v instanceof Date) return v;
  var t = Date.parse(v);
  return isNaN(t) ? null : new Date(t);
}

// 2点間の距離（メートル）
function haversine(lat1, lng1, lat2, lng2) {
  var R = 6371000;
  var toRad = function (d) { return d * Math.PI / 180; };
  var dLat = toRad(lat2 - lat1);
  var dLng = toRad(lng2 - lng1);
  var a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) *
    Math.sin(dLng / 2) * Math.sin(dLng / 2);
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

function round(n, d) {
  var f = Math.pow(10, d);
  return Math.round(n * f) / f;
}
