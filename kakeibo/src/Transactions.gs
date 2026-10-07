/**
 * 取引データ（スプレッドシートの「取引」シート）の読み書き。
 * Webアプリで登録した内容はこのシートに直接書き込まれるため、スプレッドシートにも即時反映されます。
 */

function txSheet_() {
  return getSpreadsheet_().getSheetByName(CONFIG.SHEET_TX);
}

function getCategories_() {
  const sh = getSpreadsheet_().getSheetByName(CONFIG.SHEET_CAT);
  const values = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues() : [];
  const expense = values.filter(r => r[0] && r[1] !== '収入').map(r => String(r[0]));
  const income = values.filter(r => r[0] && r[1] === '収入').map(r => String(r[0]));
  return {
    expense: expense.length ? expense : DEFAULT_EXPENSE_CATEGORIES,
    income: income.length ? income : DEFAULT_INCOME_CATEGORIES,
  };
}

function readAllTx_() {
  const sh = txSheet_();
  const n = sh.getLastRow() - 1;
  if (n <= 0) return [];
  return sh.getRange(2, 1, n, TX_HEADERS.length).getValues().map((r, i) => ({
    row: i + 2,
    id: String(r[TX_COL.id - 1]),
    date: formatDate_(r[TX_COL.date - 1]),
    type: String(r[TX_COL.type - 1] || '支出'),
    amount: Number(r[TX_COL.amount - 1]) || 0,
    store: String(r[TX_COL.store - 1] || ''),
    category: String(r[TX_COL.category - 1] || ''),
    payment: String(r[TX_COL.payment - 1] || ''),
    source: String(r[TX_COL.source - 1] || ''),
    memo: String(r[TX_COL.memo - 1] || ''),
    receipt: String(r[TX_COL.receipt - 1] || ''),
  }));
}

function normalizeTx_(tx) {
  const date = formatDate_(tx.date);
  const amount = Math.abs(Math.round(Number(String(tx.amount).replace(/[^\d.-]/g, '')) || 0));
  return {
    date: date,
    type: tx.type === '収入' ? '収入' : '支出',
    amount: amount,
    store: String(tx.store || '').trim(),
    category: String(tx.category || 'その他').trim(),
    payment: String(tx.payment || tx.payment_method || '現金').trim(),
    memo: String(tx.memo || '').trim(),
    receipt: String(tx.receipt || '').trim(),
  };
}

function toRow_(tx, source) {
  const row = new Array(TX_HEADERS.length).fill('');
  row[TX_COL.id - 1] = Utilities.getUuid();
  row[TX_COL.date - 1] = tx.date;
  row[TX_COL.ym - 1] = tx.date.slice(0, 7);
  row[TX_COL.type - 1] = tx.type;
  row[TX_COL.amount - 1] = tx.amount;
  row[TX_COL.store - 1] = tx.store;
  row[TX_COL.category - 1] = tx.category;
  row[TX_COL.payment - 1] = tx.payment;
  row[TX_COL.source - 1] = source;
  row[TX_COL.memo - 1] = tx.memo;
  row[TX_COL.receipt - 1] = tx.receipt;
  row[TX_COL.createdAt - 1] = new Date();
  return row;
}

function storeKey_(s) {
  return String(s || '').normalize('NFKC').toLowerCase().replace(/[\s・\-ー_.,()（）]/g, '').slice(0, 8);
}

/**
 * 取引をまとめて登録します。二重計上を避けるため次のように扱います。
 *  - 同じ日付・金額・店舗の取引が既にある → スキップ
 *  - レシートを登録したとき、同じ日付・金額のカード/Gmail/CSV取込の行がある
 *    → 新しい行は作らず、その行にレシートの店舗名・勘定科目・画像リンクを補完
 *  - Gmail/CSVで取り込むとき、同じ日付・金額のレシート行（現金以外）がある → スキップ
 */
function addTransactions_(list, source) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const existing = readAllTx_();
    const isReceipt = source === 'レシート';
    const rows = [];
    let skipped = 0, merged = 0;

    list.map(normalizeTx_).filter(tx => tx.amount > 0).forEach(tx => {
      if (source === '手動') { // 手入力は本人の意思なので重複チェックしない
        rows.push(toRow_(tx, source));
        return;
      }
      const sameDayAmount = existing.filter(e => e.date === tx.date && e.amount === tx.amount && e.type === tx.type);
      if (sameDayAmount.some(e => storeKey_(e.store) === storeKey_(tx.store))) {
        skipped++;
        return;
      }
      if (isReceipt && tx.payment !== '現金') {
        const target = sameDayAmount.find(e => e.source !== 'レシート' && e.source !== '手動' && !e.receipt);
        if (target) {
          const sh = txSheet_();
          sh.getRange(target.row, TX_COL.store).setValue(tx.store || target.store);
          sh.getRange(target.row, TX_COL.category).setValue(tx.category);
          sh.getRange(target.row, TX_COL.receipt).setValue(tx.receipt);
          if (tx.memo) sh.getRange(target.row, TX_COL.memo).setValue(tx.memo);
          target.receipt = tx.receipt;
          merged++;
          return;
        }
      }
      if (!isReceipt && sameDayAmount.some(e => e.source === 'レシート' && e.payment !== '現金')) {
        skipped++;
        return;
      }
      rows.push(toRow_(tx, source));
      existing.push(Object.assign({ source: source, row: -1 }, tx));
    });

    if (rows.length) {
      const sh = txSheet_();
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, TX_HEADERS.length).setValues(rows);
    }
    return { added: rows.length, merged: merged, skipped: skipped };
  } finally {
    lock.releaseLock();
  }
}

function findRowById_(id) {
  const sh = txSheet_();
  const n = sh.getLastRow() - 1;
  if (n <= 0) return -1;
  const ids = sh.getRange(2, TX_COL.id, n, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function updateTransaction_(tx) {
  const row = findRowById_(tx.id);
  if (row < 0) throw new Error('取引が見つかりません');
  const t = normalizeTx_(tx);
  const sh = txSheet_();
  sh.getRange(row, TX_COL.date, 1, 6).setValues([[t.date, t.date.slice(0, 7), t.type, t.amount, t.store, t.category]]);
  sh.getRange(row, TX_COL.payment).setValue(t.payment);
  sh.getRange(row, TX_COL.memo).setValue(t.memo);
}

function deleteTransaction_(id) {
  const row = findRowById_(id);
  if (row < 0) throw new Error('取引が見つかりません');
  txSheet_().deleteRow(row);
}

function listTransactions_(ym) {
  return readAllTx_()
    .filter(t => !ym || t.date.slice(0, 7) === ym)
    .sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : b.row - a.row))
    .map(t => { delete t.row; return t; });
}

function dashboard_(ym) {
  const list = listTransactions_(ym);
  const sum = (arr) => arr.reduce((s, t) => s + t.amount, 0);
  const expenses = list.filter(t => t.type === '支出');
  const income = list.filter(t => t.type === '収入');
  const group = (arr, key) => {
    const m = {};
    arr.forEach(t => { m[t[key] || '未分類'] = (m[t[key] || '未分類'] || 0) + t.amount; });
    return Object.keys(m).map(k => ({ name: k, amount: m[k] })).sort((a, b) => b.amount - a.amount);
  };
  return {
    ym: ym,
    expense: sum(expenses),
    income: sum(income),
    byCategory: group(expenses, 'category'),
    byPayment: group(expenses, 'payment'),
    recent: list.slice(0, 8),
  };
}
