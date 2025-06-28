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

function addSendHistory(recipient, template, status, personalizedBody) {
  try {
    var sheet = getSheetByName("送信履歴");
    var dateString = getJSTDateString();

    // 本文は長いので省略版を作成（最初の50文字まで）
    var bodyPreview = personalizedBody
      ? personalizedBody.substring(0, 50) + "..."
      : "";

    sheet.appendRow([
      dateString,
      recipient.name,
      recipient.email,
      template.subject,
      bodyPreview,
      template.templateId,
      status,
    ]);

    console.log("履歴記録完了: " + recipient.name);
  } catch (error) {
    console.log("履歴記録エラー: " + error.toString());
  }
}

// 日時フォーマット関数
function getJSTDateString() {
  var date = new Date();
  return Utilities.formatDate(date, "Asia/Tokyo", "yyyy/MM/dd HH:mm");
}

// 安全な履歴記録関数（エラー時も動作継続）
function addSendHistorySafe(recipient, template, status, personalizedBody) {
  try {
    addSendHistory(recipient, template, status, personalizedBody);
  } catch (error) {
    // 履歴記録に失敗してもメール送信処理は継続
    console.log(
      "履歴記録に失敗しましたが、処理を継続します: " + error.toString()
    );

    // 代替手段として、ログに記録
    console.log(
      "代替履歴: " +
        getJSTDateString() +
        " | " +
        recipient.name +
        " | " +
        status
    );
  }
}

// 古い履歴データを削除する関数（指定日数より古いデータを削除）
function cleanupOldHistory(daysToKeep) {
  var sheet = getSheetByName("送信履歴");
  var data = sheet.getDataRange().getValues();

  if (data.length <= 1) return; // ヘッダーのみの場合は何もしない

  var cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - daysToKeep);

  var rowsToDelete = [];

  // 1行目はヘッダーなので2行目から確認
  for (var i = 1; i < data.length; i++) {
    var dateString = data[i][0];
    var recordDate = new Date(dateString);

    if (recordDate < cutoffDate) {
      rowsToDelete.push(i + 1); // シートの行番号は1から始まる
    }
  }

  // 後ろから削除（行番号がずれないように）
  for (var j = rowsToDelete.length - 1; j >= 0; j--) {
    sheet.deleteRow(rowsToDelete[j]);
  }

  console.log("古い履歴データを削除しました: " + rowsToDelete.length + "件");
}

// 送信統計を取得する関数
function getSendStatistics() {
  var sheet = getSheetByName("送信履歴");
  var data = sheet.getDataRange().getValues();

  var totalCount = data.length - 1; // ヘッダー行を除く
  var successCount = 0;
  var failureCount = 0;
  var skipCount = 0;

  // 1行目はヘッダーなので2行目から集計
  for (var i = 1; i < data.length; i++) {
    var status = data[i][6]; // ステータス列

    if (status === "成功") {
      successCount++;
    } else if (status.indexOf("失敗") === 0) {
      failureCount++;
    } else if (status.indexOf("スキップ") === 0) {
      skipCount++;
    }
  }

  return {
    total: totalCount,
    success: successCount,
    failure: failureCount,
    skip: skipCount,
  };
}

// 統計情報を表示する関数
function showStatistics() {
  var stats = getSendStatistics();

  console.log("=== 送信統計 ===");
  console.log("総件数: " + stats.total);
  console.log("成功: " + stats.success + "件");
  console.log("失敗: " + stats.failure + "件");
  console.log("スキップ: " + stats.skip + "件");

  if (stats.total > 0) {
    var successRate = Math.round((stats.success / stats.total) * 100);
    console.log("成功率: " + successRate + "%");
  }
}

// 送信履歴を取得する関数
function getSendHistory(limit) {
  var sheet = getSheetByName("送信履歴");
  var data = sheet.getDataRange().getValues();

  // limitが指定されている場合は最新のN件のみ取得
  if (limit && limit > 0) {
    var startRow = Math.max(1, data.length - limit);
    return data.slice(startRow);
  }

  return data;
}

// 最新の送信履歴を表示する関数
function showRecentHistory() {
  var history = getSendHistory(10); // 最新10件

  console.log("=== 最新の送信履歴 ===");
  for (var i = 1; i < history.length; i++) {
    // 1行目はヘッダーなのでスキップ
    var record = history[i];
    console.log(record[0] + " | " + record[1] + " | " + record[6]); // 日時 | 宛先名 | ステータス
  }
}
