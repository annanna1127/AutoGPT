/**
 * クレジットカード明細CSV → Googleスプレッドシート 自動取り込み
 *
 * 使い方は README.md を参照。
 * 1. CONFIG.INBOX_FOLDER_ID に、CSVを置くGoogleドライブのフォルダIDを設定
 * 2. setup() を一度だけ実行（明細・月別集計シートの作成と、1時間ごとの自動実行を登録）
 */

const CONFIG = {
  // CSVを置くフォルダのID（フォルダURLの .../folders/ の後ろの文字列）
  INBOX_FOLDER_ID: 'ここにフォルダIDを貼り付け',
  // 取り込み済みCSVの移動先（INBOXフォルダ内に自動作成）
  DONE_FOLDER_NAME: '取込済み',
  // 書き込み先のシート名
  SHEET_NAME: '明細',
  SUMMARY_SHEET_NAME: '月別集計',
  // 明細の「カード」列に入れる名前
  CARD_NAME: 'アメックス',
  // 店名にこれらを含む行は取り込まない（カード代金の引き落とし＝支払い行など）
  EXCLUDE_PATTERNS: [/お支払.*ありがとう/, /ご入金/, /口座振替/],
  // 自動実行の間隔（時間）
  TRIGGER_HOURS: 1,
};

const HEADERS = ['利用日', '利用店名', '金額', 'カード', '取込日時', '重複判定キー'];

// 見出しの候補（上ほど優先）
const DATE_KEYS = ['ご利用年月日', '利用年月日', 'ご利用日', '利用日', '日付', '年月日'];
const STORE_KEYS = ['ご利用内容', 'ご利用店名', '利用店名', 'ご利用先', '利用先', '店名', '摘要', '内容', '加盟店'];
const AMOUNT_KEYS = ['金額', 'ご利用金額', '利用金額', 'お支払金額', '支払金額', '請求金額'];
// 外貨建ての金額列は使わない
const FOREIGN_RE = /外貨|外国|現地|通貨|レート/;

/** 初回に一度だけ実行 */
function setup() {
  getSheet_();
  setupSummary_();
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === 'importCsvFiles')
    .forEach((t) => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('importCsvFiles').timeBased().everyHours(CONFIG.TRIGGER_HOURS).create();
  importCsvFiles();
}

/** スプレッドシートを開いたときにメニューを追加 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('カード明細')
    .addItem('今すぐ取り込む', 'importCsvFiles')
    .addItem('月別集計シートを作り直す', 'rebuildSummary')
    .addToUi();
}

/** INBOXフォルダ内のCSVをすべて取り込む */
function importCsvFiles() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) return;
  try {
    const inbox = DriveApp.getFolderById(CONFIG.INBOX_FOLDER_ID);
    const done = getOrCreateSubfolder_(inbox, CONFIG.DONE_FOLDER_NAME);
    const sheet = getSheet_();
    const existingKeys = loadExistingKeys_(sheet);
    const now = new Date();

    let added = 0;
    let skipped = 0;
    const files = inbox.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      if (!/\.csv$/i.test(file.getName()) && file.getMimeType() !== 'text/csv') continue;

      const text = decodeCsv_(file.getBlob());
      const records = parseCardCsv(Utilities.parseCsv(text));

      const rows = [];
      for (const r of records) {
        if (existingKeys.has(r.key)) {
          skipped++;
          continue;
        }
        existingKeys.add(r.key);
        rows.push([r.date, r.store, r.amount, CONFIG.CARD_NAME, now, r.key]);
      }
      if (rows.length) {
        sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, HEADERS.length).setValues(rows);
        added += rows.length;
      }
      file.moveTo(done);
    }

    if (sheet.getLastRow() > 1) {
      sheet
        .getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length)
        .sort([{ column: 1, ascending: true }]);
    }
    console.log(`取り込み完了: 追加 ${added} 件 / 重複スキップ ${skipped} 件`);
  } finally {
    lock.releaseLock();
  }
}

/**
 * CSVの2次元配列から明細を取り出す。
 * 見出し行があれば列名で、なければ「日付,店名,金額,...」の並びとみなして読む。
 * 合計行・注意書きなど、日付が読めない行は無視する。
 */
function parseCardCsv(table) {
  const header = findHeader_(table);
  const cols = header ? header.cols : { date: 0, store: 1, amount: 2 };
  const start = header ? header.row + 1 : 0;

  const records = [];
  const seen = {};
  for (let i = start; i < table.length; i++) {
    const row = table[i];
    const date = parseDate_(row[cols.date]);
    const amount = parseAmount_(row[cols.amount]);
    if (!date || amount === null) continue;
    const store = String(row[cols.store] || '').trim();
    if (CONFIG.EXCLUDE_PATTERNS.some((re) => re.test(store))) continue;

    // 同じCSV内の「同日・同店・同額」は別の利用として残すため連番を付ける
    const base = [formatDate_(date), store, amount].join('|');
    seen[base] = (seen[base] || 0) + 1;
    records.push({ date, store, amount, key: `${base}#${seen[base]}` });
  }
  return records;
}

function findHeader_(table) {
  for (let i = 0; i < Math.min(table.length, 20); i++) {
    const cells = table[i].map((c) => String(c).replace(/\s/g, ''));
    const date = findCol_(cells, DATE_KEYS);
    const amount = findCol_(cells, AMOUNT_KEYS);
    if (date < 0 || amount < 0) continue;
    let store = findCol_(cells, STORE_KEYS);
    if (store < 0) store = date + 1;
    return { row: i, cols: { date, store, amount } };
  }
  return null;
}

function findCol_(cells, keys) {
  // 完全一致を優先し、次に部分一致（外貨の列は除く）
  for (const key of keys) {
    const idx = cells.indexOf(key);
    if (idx >= 0) return idx;
  }
  for (const key of keys) {
    const idx = cells.findIndex((c) => c.includes(key) && !FOREIGN_RE.test(c));
    if (idx >= 0) return idx;
  }
  return -1;
}

function parseDate_(value) {
  if (value === undefined || value === null) return null;
  const s = String(value).trim();
  let m = s.match(/^(\d{2,4})[\/\-.年](\d{1,2})[\/\-.月](\d{1,2})日?$/);
  if (!m) m = s.match(/^(\d{4})(\d{2})(\d{2})$/);
  if (!m) {
    // 月/日/年（英語表記のCSV）
    const us = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (us) m = [s, us[3], us[1], us[2]];
  }
  if (!m) return null;
  let y = Number(m[1]);
  if (y < 100) y += 2000;
  const d = new Date(y, Number(m[2]) - 1, Number(m[3]));
  return d.getMonth() === Number(m[2]) - 1 ? d : null;
}

function parseAmount_(value) {
  if (value === undefined || value === null) return null;
  const s = String(value).replace(/[,，¥￥円\s]/g, '').replace(/^▲/, '-');
  if (!/^-?\d+(\.\d+)?$/.test(s)) return null;
  return Number(s);
}

function formatDate_(d) {
  const p = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}

/** UTF-8で読めなければShift_JISとして読む（日本のカード会社のCSVはShift_JISが多い） */
function decodeCsv_(blob) {
  let text = blob.getDataAsString('UTF-8');
  if (text.includes('�')) text = blob.getDataAsString('Shift_JIS');
  return text.replace(/^﻿/, '');
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange('A:A').setNumberFormat('yyyy/mm/dd');
    sheet.getRange('C:C').setNumberFormat('#,##0');
    sheet.getRange('E:E').setNumberFormat('yyyy/mm/dd hh:mm');
    sheet.hideColumns(HEADERS.length);
  }
  return sheet;
}

function loadExistingKeys_(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return new Set();
  const values = sheet.getRange(2, HEADERS.length, last - 1, 1).getValues();
  return new Set(values.map((v) => String(v[0])));
}

/** 月別集計シート（数式で自動更新されるので、作るのは一度だけ） */
function setupSummary_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(CONFIG.SUMMARY_SHEET_NAME)) return;
  const sheet = ss.insertSheet(CONFIG.SUMMARY_SHEET_NAME);
  const src = `'${CONFIG.SHEET_NAME}'`;

  // 左：月別の合計・件数（新しい月が上）
  sheet.getRange('A1').setFormula(
    `=QUERY({ARRAYFORMULA(IF(${src}!A2:A="",,TEXT(${src}!A2:A,"yyyy-mm"))),${src}!C2:C},` +
      `"select Col1, sum(Col2), count(Col2) where Col1 is not null group by Col1 order by Col1 desc ` +
      `label Col1 '利用月', sum(Col2) '合計金額', count(Col2) '件数'",0)`
  );
  // 右：利用店別の合計（金額の大きい順）
  sheet.getRange('E1').setFormula(
    `=QUERY(${src}!B2:C,"select B, sum(C), count(C) where B is not null group by B order by sum(C) desc ` +
      `label B '利用店名', sum(C) '合計金額', count(C) '件数'",0)`
  );
  sheet.getRange('B:B').setNumberFormat('#,##0');
  sheet.getRange('F:F').setNumberFormat('#,##0');
  sheet.setFrozenRows(1);

  // 月別合計の棒グラフ（直近24か月分まで）
  const chart = sheet
    .newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sheet.getRange('A1:B25'))
    .setOption('title', '月別のご利用金額')
    .setOption('legend', { position: 'none' })
    .setOption('hAxis', { direction: -1 })
    .setPosition(1, 9, 0, 0)
    .build();
  sheet.insertChart(chart);
}

/** 月別集計シートを削除して作り直す */
function rebuildSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const old = ss.getSheetByName(CONFIG.SUMMARY_SHEET_NAME);
  if (old) ss.deleteSheet(old);
  setupSummary_();
}

function getOrCreateSubfolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
