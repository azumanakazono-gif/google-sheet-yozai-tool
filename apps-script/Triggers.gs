/**
 * 週次トリガーをスクリプトから設定したい場合に一度だけ実行する。
 * (GASエディタの「トリガー」画面から手動設定する場合はこの関数は不要)
 * 既存の syncWeeklyReports 用トリガーがあれば一旦削除してから作り直す。
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
 * 「今月」等の予材シートで見込確度が手動編集された際に、変更理由の入力ダイアログ
 * (onConfidenceCellEdited)を実行するインストーラブルトリガーを設定する。一度だけ実行する。
 * (GASエディタの「トリガー」画面から手動設定する場合はこの関数は不要。ただしその場合も
 * 必ず「インストーラブル トリガー」として、イベントの種類を「編集時」で設定すること。
 * 単純トリガーのonEdit(e)ではダイアログが表示できないため、この機能には使えない。)
 * 既存の onConfidenceCellEdited 用トリガーがあれば一旦削除してから作り直す。
 */
function installConfidenceChangeTrigger() {
  const ss = SpreadsheetApp.openById(CONFIG.TARGET_SPREADSHEET_ID);
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === 'onConfidenceCellEdited') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('onConfidenceCellEdited')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
}
