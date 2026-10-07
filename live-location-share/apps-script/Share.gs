/* =============================================================================
 * 現在地リンク共有 — Google Apps Script バックエンド（小さなKV）
 *
 * 役割:
 *   - 共有者の端末からの POST を受け、shareId ごとに「最新の現在地」を保存する。
 *   - 閲覧リンク（view.html）からの GET に、その現在地を JSONP で返す。
 *
 * なぜ JSONP か:
 *   閲覧ページ（GitHub Pages など別ドメイン）から読めるよう、callback 付き GET は
 *   JSONP で返します。これにより CORS 設定なしでどこからでも読み取れます。
 *
 * セットアップ:
 *   1. Google スプレッドシートを新規作成。
 *   2. 拡張機能 → Apps Script にこのコードを貼り付けて保存。
 *   3. デプロイ → 新しいデプロイ → ウェブアプリ
 *        実行ユーザー: 自分 / アクセス: 全員
 *   4. 発行された .../exec の URL を共有ページ（index.html）に設定。
 *
 * 注意:
 *   - ここに保存されるのは「現在地」です。個人情報として扱ってください。
 *   - 期限切れ・停止済みのデータは下の cleanup() を時間トリガーで回すと自動削除できます。
 * ========================================================================== */

var SHEET = 'LIVE';
var HEADERS = ['shareId', 'name', 'lat', 'lng', 'accuracy', 'updatedAt', 'expiresAt', 'active', 'trail'];
var TRAIL_MAX = 30; // 閲覧側で経路を描くために保持する直近の点数

/* ------------------------------ 書き込み（共有者→） ---------------------- */

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    var d = JSON.parse(e.postData.contents);
    if (!d.shareId) return json({ ok: false, error: 'no shareId' });

    var sh = getSheet();
    var row = findRow(sh, d.shareId);

    // 停止指示
    if (d.stop) {
      if (row > 0) sh.getRange(row, col('active')).setValue(false);
      return json({ ok: true, status: 'stopped' });
    }

    var now = new Date();
    var expiresAt = d.expiresAt ? new Date(d.expiresAt) : new Date(now.getTime() + 4 * 3600000);

    // 直近経路（trail）を更新
    var trail = [];
    if (row > 0) {
      try { trail = JSON.parse(sh.getRange(row, col('trail')).getValue() || '[]'); } catch (_) { trail = []; }
    }
    if (typeof d.lat === 'number' && typeof d.lng === 'number') {
      trail.push([round(d.lat, 6), round(d.lng, 6), now.getTime()]);
      if (trail.length > TRAIL_MAX) trail = trail.slice(trail.length - TRAIL_MAX);
    }

    var values = [
      d.shareId,
      (d.name || '').toString().slice(0, 40),
      typeof d.lat === 'number' ? d.lat : '',
      typeof d.lng === 'number' ? d.lng : '',
      typeof d.accuracy === 'number' ? Math.round(d.accuracy) : '',
      now.toISOString(),
      expiresAt.toISOString(),
      true,
      JSON.stringify(trail)
    ];

    if (row > 0) {
      sh.getRange(row, 1, 1, HEADERS.length).setValues([values]);
    } else {
      sh.appendRow(values);
    }
    return json({ ok: true, status: 'live' });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* ------------------------------ 読み取り（→閲覧者） ---------------------- */

function doGet(e) {
  var cb = e && e.parameter && e.parameter.callback;
  var id = e && e.parameter && e.parameter.id;

  if (!id) return reply(cb, { ok: true, service: 'live-location-share' });

  var sh = getSheet();
  var row = findRow(sh, id);
  if (row <= 0) return reply(cb, { ok: true, status: 'notfound' });

  var r = sh.getRange(row, 1, 1, HEADERS.length).getValues()[0];
  var active = r[col('active') - 1] === true || r[col('active') - 1] === 'TRUE';
  var expiresAt = new Date(r[col('expiresAt') - 1]);
  if (!active || new Date() > expiresAt) {
    return reply(cb, { ok: true, status: 'ended', name: r[1] });
  }

  var trail = [];
  try { trail = JSON.parse(r[col('trail') - 1] || '[]'); } catch (_) {}

  return reply(cb, {
    ok: true,
    status: 'live',
    name: r[1],
    lat: r[2],
    lng: r[3],
    accuracy: r[4],
    updatedAt: r[5],
    trail: trail
  });
}

/* ------------------------------ 後始末（任意） -------------------------- *
 * 時間主導トリガーで1日1回など回すと、停止済み・期限切れの行を削除します。
 * ------------------------------------------------------------------------ */
function cleanup() {
  var sh = getSheet();
  var last = sh.getLastRow();
  if (last < 2) return;
  var data = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
  var now = new Date();
  for (var i = data.length - 1; i >= 0; i--) {
    var active = data[i][col('active') - 1] === true;
    var exp = new Date(data[i][col('expiresAt') - 1]);
    if (!active || now > exp) sh.deleteRow(i + 2);
  }
}

/* ------------------------------ ユーティリティ -------------------------- */

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function findRow(sh, shareId) {
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0] === shareId) return i + 2;
  }
  return -1;
}

function col(name) { return HEADERS.indexOf(name) + 1; }

function round(n, d) { var f = Math.pow(10, d); return Math.round(n * f) / f; }

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// callback があれば JSONP、なければ JSON を返す
function reply(callback, obj) {
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(obj) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return json(obj);
}
