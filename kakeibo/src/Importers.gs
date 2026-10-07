/**
 * 自動取込: レシート・スクリーンショット画像 / Gmail の利用通知 / 明細CSV
 */

/**
 * レシート（またはPayPay・Suicaなどの履歴画面のスクショ）を読み取り、画像をドライブに保存します。
 * 取引はまだ登録せず、画面で確認してから saveTransactions で登録します。
 */
function processImage_(payload) {
  const mimeType = payload.mimeType || 'image/jpeg';
  const kind = payload.kind === 'screenshot' ? 'スクショ' : 'レシート';
  const hint = kind === 'スクショ'
    ? 'これは決済アプリや交通系ICの利用履歴画面のスクリーンショットです。写っている取引をすべて抽出してください。'
    : 'これはお店のレシート（または領収書）の写真です。';

  let txs = [];
  let error = '';
  try {
    txs = extractTransactions_([
      { type: 'image', source: { type: 'base64', media_type: mimeType, data: payload.base64 } },
      { type: 'text', text: hint + '\n支払方法が判断できない場合の既定値: ' + (payload.defaultPayment || '現金') },
    ]);
  } catch (e) {
    error = e.message;
  }
  // 読み取りに失敗しても画像は保存しておく（あとで手入力できるように）
  const file = saveReceiptImage_(payload.base64, mimeType, txs[0]);
  return { kind: kind, transactions: txs, fileUrl: file.url, error: error };
}

/** Gmail のクレジットカード・銀行の利用通知メールから取引を取り込みます（1時間ごとに自動実行）。 */
function importFromGmail() {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(CONFIG.SHEET_SRC);
  if (!sh || sh.getLastRow() < 2) return { added: 0, merged: 0, skipped: 0, mails: 0 };
  const sources = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues()
    .filter(r => r[0] === true && r[3]);

  const label = GmailApp.getUserLabelByName(CONFIG.GMAIL_DONE_LABEL) || GmailApp.createLabel(CONFIG.GMAIL_DONE_LABEL);
  const total = { added: 0, merged: 0, skipped: 0, mails: 0 };

  sources.forEach(([, name, payment, query]) => {
    const q = query + ' -label:' + CONFIG.GMAIL_DONE_LABEL.replace('/', '-') + ' newer_than:45d';
    const threads = GmailApp.search(q, 0, CONFIG.GMAIL_MAX_THREADS);
    if (!threads.length) return;

    const mails = [];
    threads.forEach(th => th.getMessages().forEach(m => {
      mails.push('受信日: ' + formatDate_(m.getDate()) + '\n件名: ' + m.getSubject() + '\n本文:\n' +
        m.getPlainBody().replace(/\n{3,}/g, '\n\n').slice(0, 3000));
    }));

    // 8通ずつまとめて読み取る（API呼び出し回数を減らすため）
    for (let i = 0; i < mails.length; i += 8) {
      const chunk = mails.slice(i, i + 8);
      const txs = extractTransactions_([{
        type: 'text',
        text: '以下は「' + name + '」から届いた利用通知メールです。利用日（なければ受信日）・金額・利用先を取引として抽出してください。' +
          'payment_method は「' + payment + '」にしてください。\n\n' +
          chunk.map((m, j) => '===== メール' + (j + 1) + ' =====\n' + m).join('\n\n'),
      }]);
      const r = addTransactions_(txs.map(t => Object.assign(t, { payment: payment })), 'Gmail:' + name);
      total.added += r.added; total.merged += r.merged; total.skipped += r.skipped;
    }
    total.mails += mails.length;
    label.addToThreads(threads);
  });

  props_().setProperty('LAST_GMAIL_IMPORT', new Date().toISOString());
  return total;
}

/**
 * カード会社・銀行・PayPay などからダウンロードした明細CSVを取り込みます。
 * 列の並びは会社ごとに違うので、Claude に読み取らせて共通の形式に変換します。
 */
function importCsv_(payload) {
  const lines = String(payload.text || '').split(/\r?\n/).filter(l => l.trim());
  if (lines.length < 1) throw new Error('CSVが空です');
  const header = lines[0];
  const body = lines.slice(1);
  const name = payload.sourceName || 'CSV';
  const total = { added: 0, merged: 0, skipped: 0 };

  for (let i = 0; i < Math.max(body.length, 1); i += 40) {
    const chunk = [header].concat(body.slice(i, i + 40)).join('\n');
    const txs = extractTransactions_([{
      type: 'text',
      text: '以下は「' + name + '」の利用明細CSVです（先頭行は見出し）。各行を取引に変換してください。' +
        '支払方法が判断できない場合は「' + (payload.payment || 'クレジットカード') + '」にしてください。\n\n' + chunk,
    }]);
    const r = addTransactions_(txs, 'CSV:' + name);
    total.added += r.added; total.merged += r.merged; total.skipped += r.skipped;
  }
  return total;
}
