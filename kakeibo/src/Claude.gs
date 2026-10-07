/**
 * Claude API でレシート画像・利用通知メール・CSV明細から取引を読み取ります。
 * Apps Script には公式SDKがないため、Messages API を UrlFetchApp で直接呼び出します。
 */

const EXTRACT_SYSTEM_PROMPT = [
  'あなたは日本の家計簿アプリの入力アシスタントです。与えられたレシート・決済画面のスクリーンショット・利用通知メール・明細CSVから、家計簿に記録すべき取引を抽出します。',
  '',
  'ルール:',
  '- 金額は税込の支払総額を円単位の整数で。レシート1枚は原則1件の取引（合計金額）にまとめ、主な品目をmemoに短く書く。',
  '- 決済履歴のスクリーンショットや明細のように複数の取引が写っている場合は、1行ずつ別の取引にする。',
  '- 日付は YYYY-MM-DD。年が書かれていなければ、今日の日付を基準に直近の過去日として推定する。',
  '- store には店舗名・支払先を、読みやすい名前で入れる（例: 「ｾﾌﾞﾝｲﾚﾌﾞﾝ」→「セブン-イレブン」）。',
  '- category は与えられた選択肢から最も適切なものを選ぶ。スーパー・コンビニの食品は「食費」、飲食店は「外食」、ドラッグストアの日用品は「日用品」、電車・バスは「交通費」。',
  '- payment_method はレシートの「お預り/お釣り」なら現金、「クレジット」「VISA」等ならクレジットカード、「PayPay」ならPayPay、「交通系IC」「Suica」ならSuica。判断できなければ指定された既定値を使う。',
  '- 二重計上になるものは含めない: 電子マネーや残高への「チャージ」、口座間の振替、クレジットカード利用分の口座引き落とし、ポイント付与のみの行、キャンセル済みの取引。',
  '- 給与・振込入金などの入金は type を「収入」にする。',
  '- 取引が読み取れない（家計簿に関係ない内容）場合は transactions を空配列にする。',
].join('\n');

function buildTxSchema_() {
  const cats = getCategories_();
  return {
    type: 'object',
    additionalProperties: false,
    required: ['transactions'],
    properties: {
      transactions: {
        type: 'array',
        items: {
          type: 'object',
          additionalProperties: false,
          required: ['date', 'type', 'amount', 'store', 'category', 'payment_method', 'memo'],
          properties: {
            date: { type: 'string', description: 'YYYY-MM-DD' },
            type: { type: 'string', enum: ['支出', '収入'] },
            amount: { type: 'integer', description: '税込の円金額（正の整数）' },
            store: { type: 'string' },
            category: { type: 'string', enum: cats.expense.concat(cats.income) },
            payment_method: { type: 'string', enum: PAYMENT_METHODS },
            memo: { type: 'string' },
          },
        },
      },
    },
  };
}

/**
 * @param {Object[]} content Messages API の user content ブロック
 * @return {Object[]} 取引の配列
 */
function extractTransactions_(content) {
  const apiKey = props_().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) throw new Error('Claude APIキーが未設定です。スクリプトプロパティに ANTHROPIC_API_KEY を設定してください。');

  const body = {
    model: CONFIG.MODEL,
    max_tokens: 16000,
    system: EXTRACT_SYSTEM_PROMPT + '\n\n今日の日付: ' + today_(),
    // 読み取りは単純な作業なので effort を下げて速度と料金を抑える
    output_config: { effort: 'low', format: { type: 'json_schema', schema: buildTxSchema_() } },
    // 安全性分類で断られた場合に、推奨の代替モデルで自動再実行する
    fallbacks: 'default',
    messages: [{ role: 'user', content: content }],
  };

  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
      'anthropic-beta': 'server-side-fallback-2026-07-01',
    },
    payload: JSON.stringify(body),
  });

  const status = res.getResponseCode();
  const json = JSON.parse(res.getContentText() || '{}');
  if (status === 401) throw new Error('Claude APIキーが正しくありません。');
  if (status === 429 || status >= 500) throw new Error('Claude API が混み合っています。少し待ってから再度お試しください。(' + status + ')');
  if (status !== 200) throw new Error('Claude API エラー (' + status + '): ' + ((json.error && json.error.message) || ''));

  if (json.stop_reason === 'refusal') throw new Error('この画像・データは読み取れませんでした。手入力してください。');
  if (json.stop_reason === 'max_tokens') throw new Error('データが多すぎて読み切れませんでした。分割してお試しください。');

  const textBlock = (json.content || []).find(b => b.type === 'text');
  if (!textBlock) return [];
  const parsed = JSON.parse(textBlock.text);
  return (parsed.transactions || []).map(t => ({
    date: t.date,
    type: t.type,
    amount: t.amount,
    store: t.store,
    category: t.category,
    payment: t.payment_method,
    memo: t.memo,
  }));
}
