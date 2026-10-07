/**
 * 設定値。APIキーなどの秘密情報はコードに書かず、
 * 「プロジェクトの設定 > スクリプト プロパティ」に保存します。
 */
const CONFIG = {
  // レシート・明細の読み取りに使う Claude のモデル
  MODEL: 'claude-opus-5-5',
  TIMEZONE: 'Asia/Tokyo',

  SPREADSHEET_NAME: '家計簿',
  SHEET_TX: '取引',
  SHEET_CAT: '勘定科目',
  SHEET_SRC: '連携設定',
  SHEET_MONTHLY: '月次集計',
  SHEET_BY_CAT: '科目別集計',

  DRIVE_ROOT_NAME: '家計簿レシート',
  GMAIL_DONE_LABEL: '家計簿/取込済',
  GMAIL_MAX_THREADS: 30,
};

// 取引シートの列（順番を変えるときは TX_COL も合わせて変更）
const TX_HEADERS = ['ID', '日付', '年月', '種別', '金額', '店舗・内容', '勘定科目', '支払方法', '入力元', 'メモ', 'レシート', '登録日時'];
const TX_COL = {
  id: 1, date: 2, ym: 3, type: 4, amount: 5, store: 6, category: 7,
  payment: 8, source: 9, memo: 10, receipt: 11, createdAt: 12,
};

const DEFAULT_EXPENSE_CATEGORIES = [
  '食費', '外食', '日用品', '交通費', '住居費', '水道・光熱費', '通信費',
  '医療・健康', '衣服・美容', '趣味・娯楽', '交際費', '教育・教養',
  '保険', '税金・社会保険', '特別な支出', 'その他',
];
const DEFAULT_INCOME_CATEGORIES = ['給与', '賞与', '副業', 'その他収入'];
const PAYMENT_METHODS = ['現金', 'クレジットカード', 'PayPay', 'Suica', '銀行口座', '電子マネー・その他'];

function props_() {
  return PropertiesService.getScriptProperties();
}

function today_() {
  return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');
}

function formatDate_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  const s = String(value || '').trim().replace(/\//g, '-');
  const m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (!m) return today_();
  return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
}
