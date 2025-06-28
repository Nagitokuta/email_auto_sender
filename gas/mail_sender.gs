// 差し込み変数を置換する関数
function replaceTemplate(template, name) {
  return template.replace(/{{名前}}/g, name);
}

// 単一の宛先にメールを送信
function sendMailToRecipient(recipient, template) {
  try {
    // テンプレート本文の「{{名前}}」を宛先名で置換
    var personalizedBody = replaceTemplate(template.body, recipient.name);

    // メール送信
    GmailApp.sendEmail(recipient.email, template.subject, personalizedBody);

    console.log("送信成功: " + recipient.name + " (" + recipient.email + ")");
    return "成功";
  } catch (error) {
    console.log("送信失敗: " + recipient.name + " - " + error.toString());
    return "失敗: " + error.toString();
  }
}

// 全宛先にメールを送信
function sendMailsToAll() {
  var recipients = getRecipientList();
  var results = [];

  for (var i = 0; i < recipients.length; i++) {
    var recipient = recipients[i];

    // ステータスが「送信可」の場合のみ送信
    if (recipient.status === "送信可") {
      var template = getMailTemplate(recipient.templateId);
      var result = sendMailToRecipient(recipient, template);

      results.push({
        recipient: recipient,
        template: template,
        result: result,
      });

      // 送信間隔を空ける（1秒待機）
      Utilities.sleep(1000);
    } else {
      console.log(
        "スキップ: " +
          recipient.name +
          " (ステータス: " +
          recipient.status +
          ")"
      );
    }
  }

  return results;
}

// スプレッドシートを取得する共通関数
function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

// シート名を指定してシートを取得
function getSheetByName(sheetName) {
  var ss = getSpreadsheet();
  return ss.getSheetByName(sheetName);
}

// メールテンプレート（件名・本文）を取得
function getMailTemplate(templateId) {
  var sheet = getSheetByName("メール設定");
  var data = sheet.getDataRange().getValues();

  // 1行目はヘッダーなので2行目から検索
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === templateId) {
      return {
        templateId: data[i][0],
        subject: data[i][1],
        body: data[i][2],
      };
    }
  }

  // テンプレートが見つからない場合はデフォルト
  return {
    templateId: "default",
    subject: "お知らせ",
    body: "こんにちは、{{名前}}さん。",
  };
}

// 宛先リストを取得
function getRecipientList() {
  var sheet = getSheetByName("宛先");
  var data = sheet.getDataRange().getValues();
  var recipients = [];

  // 1行目はヘッダーなので2行目から
  for (var i = 1; i < data.length; i++) {
    recipients.push({
      no: data[i][0],
      name: data[i][1],
      email: data[i][2],
      status: data[i][3],
      templateId: data[i][4],
    });
  }

  return recipients;
}

// 単一の宛先にメールを送信（履歴記録付き）
function sendMailToRecipient(recipient, template) {
  var personalizedBody = "";
  var status = "";

  try {
    // テンプレート本文の「{{名前}}」を宛先名で置換
    personalizedBody = replaceTemplate(template.body, recipient.name);

    // メール送信
    GmailApp.sendEmail(recipient.email, template.subject, personalizedBody);

    status = "成功";
    console.log("送信成功: " + recipient.name + " (" + recipient.email + ")");
  } catch (error) {
    status = "失敗: " + error.toString();
    console.log("送信失敗: " + recipient.name + " - " + error.toString());
  }

  // 送信履歴を記録（成功・失敗問わず記録）
  addSendHistory(recipient, template, status, personalizedBody);

  return status;
}

// 全宛先にメールを送信（履歴記録付き）
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

      // 送信しなかった場合も履歴に記録
      var template = getMailTemplate(recipient.templateId || "temp001");
      addSendHistory(recipient, template, "スキップ: " + recipient.status, "");
    }
  }

  return results;
}
