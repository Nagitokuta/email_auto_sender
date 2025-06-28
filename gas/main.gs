// メイン実行関数
function main() {
  console.log("=== メール自動送信開始 ===");

  try {
    var results = sendMailsWithCondition();

    console.log("=== 送信結果 ===");
    console.log("処理件数: " + results.length);

    var successCount = 0;
    var failureCount = 0;

    for (var i = 0; i < results.length; i++) {
      if (results[i].result === "成功") {
        successCount++;
      } else {
        failureCount++;
      }
    }

    console.log("成功: " + successCount + "件");
    console.log("失敗: " + failureCount + "件");
  } catch (error) {
    console.log("エラーが発生しました: " + error.toString());
  }

  console.log("=== メール自動送信終了 ===");
}

// テスト用関数（1件だけ送信）
function testSendMail() {
  var recipients = getRecipientList();

  if (recipients.length > 0) {
    var recipient = recipients[0]; // 最初の1件のみ

    if (canSendMail(recipient)) {
      var template = getMailTemplate(recipient.templateId);
      var result = sendMailToRecipient(recipient, template);
      console.log("テスト送信結果: " + result);
    } else {
      console.log("テスト対象が送信条件を満たしていません");
    }
  } else {
    console.log("宛先データがありません");
  }
}

// 送信可能かどうかをチェック
function canSendMail(recipient) {
  // メールアドレスが空の場合
  if (!recipient.email || recipient.email.trim() === "") {
    return false;
  }

  // ステータスが「送信可」でない場合
  if (recipient.status !== "送信可") {
    return false;
  }

  // メールアドレスの形式チェック（簡易版）
  var emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailPattern.test(recipient.email)) {
    return false;
  }

  return true;
}

// 条件付き送信関数
function sendMailsWithCondition() {
  var recipients = getRecipientList();
  var results = [];

  for (var i = 0; i < recipients.length; i++) {
    var recipient = recipients[i];

    if (canSendMail(recipient)) {
      var template = getMailTemplate(recipient.templateId);
      var result = sendMailToRecipient(recipient, template);

      results.push({
        recipient: recipient,
        template: template,
        result: result,
      });

      // 送信間隔を空ける
      Utilities.sleep(1000);
    } else {
      console.log("送信条件未満: " + recipient.name);
    }
  }

  return results;
}

function main() {
  console.log("=== メール自動送信開始 ===");

  try {
    var results = sendMailsWithCondition();

    console.log("=== 送信結果 ===");
    console.log("処理件数: " + results.length);

    var successCount = 0;
    var failureCount = 0;

    for (var i = 0; i < results.length; i++) {
      if (results[i].result === "成功") {
        successCount++;
      } else {
        failureCount++;
      }
    }

    console.log("成功: " + successCount + "件");
    console.log("失敗: " + failureCount + "件");

    // 全体統計も表示
    console.log("");
    showStatistics();
  } catch (error) {
    console.log("エラーが発生しました: " + error.toString());
  }

  console.log("=== メール自動送信終了 ===");
}

// 履歴確認用関数
function checkHistory() {
  showRecentHistory();
  console.log("");
  showStatistics();
}
