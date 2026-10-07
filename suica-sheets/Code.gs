/**
 * Suica 利用履歴 → Google スプレッドシート記録用 Web アプリ
 *
 * iPhone の「ショートカット」アプリ（ウォレットのオートメーション）から
 * POST された取引データを、スプレッドシートの「Suica履歴」シートに1行ずつ追記します。
 *
 * 初回だけ setup() を実行してください（シート作成と合言葉トークンの発行）。
 */

const SHEET_NAME = 'Suica履歴';
const HEADERS = ['記録日時', '利用日時', '利用先', '金額(円)', 'カード名', '元の金額表記', '受信データ'];

/** 初回セットアップ: シートを作成し、合言葉トークンを発行してログに表示する */
function setup() {
  getSheet_();
  const props = PropertiesService.getScriptProperties();
  let token = props.getProperty('TOKEN');
  if (!token) {
    token = Utilities.getUuid();
    props.setProperty('TOKEN', token);
  }
  Logger.log('ショートカットに設定するトークン: ' + token);
}

/** ショートカットからの POST を受け取って1行追記する */
function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'JSON の形式が正しくありません' });
  }
  return json_(record_(body));
}

/**
 * GET でも記録できるようにする（POST が通らない環境向け）。
 * 例: .../exec?token=xxx&amount=¥210&merchant=新宿
 * token がなければ動作確認用の応答だけ返す。
 */
function doGet(e) {
  const params = (e && e.parameter) || {};
  if (!params.token) {
    return json_({ ok: true, message: 'Suica 記録 Web アプリは動作しています' });
  }
  return json_(record_(params));
}

/** トークンを確認して「Suica履歴」シートに1行追記する */
function record_(data) {
  const token = PropertiesService.getScriptProperties().getProperty('TOKEN');
  if (!token || data.token !== token) {
    return { ok: false, error: 'トークンが一致しません' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const now = new Date();
    getSheet_().appendRow([
      now,
      parseDate_(data.date) || now,
      data.merchant || '',
      parseAmount_(data.amount),
      data.card || '',
      String(data.amount || ''),
      rawData_(data),
    ]);
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange('A:B').setNumberFormat('yyyy/mm/dd hh:mm:ss');
    sheet.getRange('D:D').setNumberFormat('#,##0');
  } else if (sheet.getRange(1, HEADERS.length).getValue() === '') {
    // 以前のバージョンで作ったシートに、増えた列の見出しを足す
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
  return sheet;
}

/** 届いた内容をそのまま残す（トークンは除く）。項目が空になるときの原因調査用 */
function rawData_(data) {
  const copy = {};
  Object.keys(data).forEach(function (key) {
    if (key !== 'token') copy[key] = data[key];
  });
  return JSON.stringify(copy);
}

/** 「¥210」「-210円」「JP¥1,000」などから数値だけを取り出す */
function parseAmount_(value) {
  if (typeof value === 'number') return value;
  const digits = String(value || '').replace(/[^0-9.\-]/g, '');
  const n = Number(digits);
  return digits && !isNaN(n) ? n : '';
}

function parseDate_(value) {
  if (!value) return null;
  const d = new Date(value);
  return isNaN(d.getTime()) ? null : d;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
