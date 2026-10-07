/**
 * 最初に1回だけ実行してください（エディタ上部で「setup」を選んで ▶ 実行）。
 * スプレッドシート・ドライブのフォルダ・Gmailラベル・自動取込トリガーを作成します。
 */
function setup() {
  const ss = getSpreadsheet_();
  ss.setSpreadsheetTimeZone(CONFIG.TIMEZONE);
  ss.setSpreadsheetLocale('ja_JP');

  setupTxSheet_(ss);
  setupCategorySheet_(ss);
  setupSourceSheet_(ss);
  setupSummarySheets_(ss);
  const defaultSheet = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) ss.deleteSheet(defaultSheet);

  getRootFolder_();
  GmailApp.getUserLabelByName(CONFIG.GMAIL_DONE_LABEL) || GmailApp.createLabel(CONFIG.GMAIL_DONE_LABEL);

  // Gmail の利用通知を1時間ごとに自動取込
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'importFromGmail')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('importFromGmail').timeBased().everyHours(1).create();

  if (!props_().getProperty('ANTHROPIC_API_KEY')) {
    Logger.log('⚠ スクリプトプロパティに ANTHROPIC_API_KEY を設定してください。');
  }
  Logger.log('セットアップ完了: ' + ss.getUrl());
}

function getSpreadsheet_() {
  const id = props_().getProperty('SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const ss = SpreadsheetApp.create(CONFIG.SPREADSHEET_NAME);
  props_().setProperty('SPREADSHEET_ID', ss.getId());
  return ss;
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function setupTxSheet_(ss) {
  const sh = getOrCreateSheet_(ss, CONFIG.SHEET_TX);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, TX_HEADERS.length).setValues([TX_HEADERS]).setFontWeight('bold').setBackground('#e8f0fe');
    sh.setFrozenRows(1);
    sh.getRange('B:B').setNumberFormat('yyyy-mm-dd');
    sh.getRange('C:C').setNumberFormat('@'); // 年月は文字列のまま保持（集計の軸）
    sh.getRange('E:E').setNumberFormat('#,##0');
    sh.getRange('L:L').setNumberFormat('yyyy-mm-dd hh:mm');
    sh.hideColumns(TX_COL.id);
  }
}

function setupCategorySheet_(ss) {
  const sh = getOrCreateSheet_(ss, CONFIG.SHEET_CAT);
  if (sh.getLastRow() > 0) return;
  const rows = [['勘定科目', '種別']]
    .concat(DEFAULT_EXPENSE_CATEGORIES.map(c => [c, '支出']))
    .concat(DEFAULT_INCOME_CATEGORIES.map(c => [c, '収入']));
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#e8f0fe');
  sh.setFrozenRows(1);
}

function setupSourceSheet_(ss) {
  const sh = getOrCreateSheet_(ss, CONFIG.SHEET_SRC);
  if (sh.getLastRow() > 0) return;
  // 検索条件は「例」です。届いている利用通知メールの差出人・件名に合わせて調整し、有効を TRUE にしてください。
  const rows = [
    ['有効', '名前', '支払方法', 'Gmail検索条件'],
    [false, '楽天カード（例）', 'クレジットカード', 'from:rakuten-card.co.jp subject:利用'],
    [false, '三井住友カード（例）', 'クレジットカード', 'from:vpass.ne.jp subject:ご利用'],
    [false, 'PayPayカード（例）', 'クレジットカード', 'from:paypay-card.co.jp subject:利用'],
    [false, '銀行の入出金通知（例）', '銀行口座', 'subject:(入出金 OR 振込 OR 引落)'],
  ];
  sh.getRange(1, 1, rows.length, 4).setValues(rows);
  sh.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#e8f0fe');
  sh.getRange(2, 1, rows.length - 1, 1).insertCheckboxes();
  sh.setFrozenRows(1);
  sh.setColumnWidth(2, 200);
  sh.setColumnWidth(4, 420);
}

function setupSummarySheets_(ss) {
  const tx = CONFIG.SHEET_TX;
  const monthly = getOrCreateSheet_(ss, CONFIG.SHEET_MONTHLY);
  if (monthly.getLastRow() === 0) {
    // 年月 × 収入/支出 の集計（取引シートが更新されると自動で再計算）
    monthly.getRange('A1').setFormula(
      `=QUERY('${tx}'!A:L, "select C, sum(E) where C is not null group by C pivot D order by C desc label C '年月'", 1)`);
    monthly.getRange('B:D').setNumberFormat('#,##0');
  }
  const byCat = getOrCreateSheet_(ss, CONFIG.SHEET_BY_CAT);
  if (byCat.getLastRow() === 0) {
    byCat.getRange('A1').setFormula(
      `=QUERY('${tx}'!A:L, "select G, sum(E) where D = '支出' and C is not null group by G pivot C label G '勘定科目'", 1)`);
    byCat.getRange('B:Z').setNumberFormat('#,##0');
  }
}
