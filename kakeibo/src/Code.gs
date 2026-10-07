/**
 * Webアプリの入口と、画面（index.html）から google.script.run で呼ばれる関数。
 */

function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('かんたん家計簿')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function getInitData() {
  const ss = getSpreadsheet_();
  return {
    today: today_(),
    categories: getCategories_(),
    payments: PAYMENT_METHODS,
    spreadsheetUrl: ss.getUrl(),
    driveUrl: getRootFolder_().getUrl(),
    hasApiKey: !!props_().getProperty('ANTHROPIC_API_KEY'),
    lastGmailImport: props_().getProperty('LAST_GMAIL_IMPORT') || '',
  };
}

function getDashboard(ym) {
  return dashboard_(ym);
}

function listTransactions(ym) {
  return listTransactions_(ym);
}

function processImage(payload) {
  return processImage_(payload);
}

/** @param {string} source '手動' | 'レシート' | 'スクショ' */
function saveTransactions(list, source) {
  const allowed = ['手動', 'レシート', 'スクショ'];
  return addTransactions_(list || [], allowed.indexOf(source) >= 0 ? source : '手動');
}

function updateTransaction(tx) {
  updateTransaction_(tx);
  return true;
}

function deleteTransaction(id) {
  deleteTransaction_(id);
  return true;
}

function runGmailImport() {
  return importFromGmail();
}

function importCsv(payload) {
  return importCsv_(payload);
}
