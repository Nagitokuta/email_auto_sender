// 設定定数
var CONFIG = {
  // シート名
  SHEET_NAMES: {
    MAIL_SETTINGS: "メール設定",
    RECIPIENTS: "宛先",
    SEND_HISTORY: "送信履歴",
  },

  // デフォルトテンプレート
  DEFAULT_TEMPLATE: {
    templateId: "default",
    subject: "お知らせ",
    body: "こんにちは、{{名前}}さん。",
  },

  // 送信間隔（ミリ秒）
  SEND_INTERVAL: 1000,

  // 送信可能ステータス
  SENDABLE_STATUS: "送信可",
};

// 設定値を取得する関数
function getConfig(key) {
  return CONFIG[key];
}
