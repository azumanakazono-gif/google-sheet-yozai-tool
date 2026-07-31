/**
 * 週次トリガーをスクリプトから設定したい場合に一度だけ実行する。
 * (GASエディタの「トリガー」画面から手動設定する場合はこの関数は不要)
 * 既存の syncWeeklyReports 用トリガーがあれば一旦削除してから作り直す。
 * 毎週月曜日の午前6時〜7時の間に実行される（GASの時間主導型トリガーは実行時刻の厳密な
 * 指定ができないため、atHour(6)で「6時〜7時の間のいずれかのタイミング」を指定する）。
 */
function installWeeklyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === 'syncWeeklyReports') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('syncWeeklyReports')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(6)
    .create();
}

/**
 * onEdit(e)のシンプルトリガーではなく、handleEdit_をインストーラブルトリガーとして
 * 登録したい場合に一度だけ実行する(任意)。インストーラブルトリガーは追加の権限承認が
 * 必要になる代わりに、MailApp等の外部サービス呼び出しも可能になる。
 * 既存の handleEdit_ 用トリガーがあれば一旦削除してから作り直す。
 */
function installEditTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === 'handleEdit_') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('handleEdit_')
    .forSpreadsheet(SpreadsheetApp.openById(CONFIG.TARGET_SPREADSHEET_ID))
    .onEdit()
    .create();
}
