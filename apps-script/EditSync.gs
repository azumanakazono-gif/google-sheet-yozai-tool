/**
 * 編集時: シート名に「今月」を含むシートの、EDIT_MIN_ROW行目以降・EDIT_TARGET_COLUMN列目
 * (既定: O列=見込確度)の変更を検知し、「週次報告記録」と「週次ステータス変更履歴」の
 * 両方へ自動転記する。
 *
 * この機能が発火するのは、このスクリプトが「31期予材リスト」スプレッドシートの
 * コンテナバインド型スクリプトとして設置されている場合のみ(スタンドアロン型では発火しない)。
 * 「拡張機能 > Apps Script」から「31期予材リスト」スプレッドシートに紐づく形でスクリプトを
 * 開いていれば、このシンプルトリガー(onEdit)は追加設定なしで自動的に動作する。
 */

/**
 * シンプルトリガー。関数名onEditはGASにおいて特別な意味を持ち、インストール操作なしで
 * このスプレッドシートが編集されるたびに自動実行される(手動実行や権限承認は不要)。
 * より強い権限(メール送信等)が必要な場合は installEditTrigger() でインストーラブル
 * トリガーとして handleEdit_ を登録することもできる。
 */
function onEdit(e) {
  try {
    handleEdit_(e);
  } catch (err) {
    Logger.log('onEdit処理でエラーが発生しました: ' + err);
  }
}

/**
 * シート名に「今月」を含むシートの、EDIT_MIN_ROW行目以降・EDIT_TARGET_COLUMN列目
 * (既定: O列=見込確度)の変更を検知し、「週次報告記録」と「週次ステータス変更履歴」の
 * 両方へ自動転記する。範囲コピー&ペーストなど複数セル一括編集にも対応する。
 */
function handleEdit_(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  if (sheetName.indexOf(CONFIG.EDIT_SHEET_NAME_KEYWORD) === -1) return;

  const range = e.range;
  const firstRow = range.getRow();
  const lastRow = firstRow + range.getNumRows() - 1;
  const firstCol = range.getColumn();
  const lastCol = firstCol + range.getNumColumns() - 1;

  if (lastRow < CONFIG.EDIT_MIN_ROW) return;
  if (CONFIG.EDIT_TARGET_COLUMN < firstCol || CONFIG.EDIT_TARGET_COLUMN > lastCol) return;

  const isSingleCell = range.getNumRows() === 1 && range.getNumColumns() === 1;
  const startRow = Math.max(firstRow, CONFIG.EDIT_MIN_ROW);
  let endRow = lastRow;
  if (endRow - startRow + 1 > CONFIG.EDIT_MAX_ROWS_PER_EVENT) {
    endRow = startRow + CONFIG.EDIT_MAX_ROWS_PER_EVENT - 1;
    Logger.log('編集行数が多いため、先頭' + CONFIG.EDIT_MAX_ROWS_PER_EVENT + '行のみ処理します。');
  }

  const editorEmail = getEditorEmail_(e);
  const weeklySheet = getTargetSheet_();
  const historySheet = getStatusHistorySheet_();

  for (let row = startRow; row <= endRow; row++) {
    const newValue = sheet.getRange(row, CONFIG.EDIT_TARGET_COLUMN).getValue();
    const oldValue = isSingleCell ? e.oldValue : '';
    if (isSingleCell && String(oldValue === undefined ? '' : oldValue) === String(newValue)) continue;

    appendEditToWeeklyReport_(sheet, row, weeklySheet);
    appendStatusHistory_(sheet, row, oldValue, newValue, editorEmail, historySheet);
  }
}

/**
 * 編集したユーザーのメールアドレスを可能な範囲で取得する。
 * 権限設定によっては取得できないことがあるため、その場合は空文字を返す。
 */
function getEditorEmail_(e) {
  try {
    if (e.user && e.user.getEmail && e.user.getEmail()) return e.user.getEmail();
  } catch (err) {
    // メールアドレス取得不可(権限設定による)。空文字で続行する。
  }
  try {
    return Session.getActiveUser().getEmail() || '';
  } catch (err) {
    return '';
  }
}

/**
 * 「今月」シートの1行分を、列名マッチングで「週次報告記録」シートへ追記する。
 * ヘッダー行位置(CONFIG.EDIT_HEADER_ROW)は「今月」シートの実際のレイアウトに合わせて
 * 調整すること(既定値: 4行目)。
 */
function appendEditToWeeklyReport_(sourceSheet, row, targetSheet) {
  const lastCol = sourceSheet.getLastColumn();
  const headerRow = sourceSheet.getRange(CONFIG.EDIT_HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const sourceHeader = headerRow.map(function (h) { return String(h).trim(); });
  const rowValues = sourceSheet.getRange(row, 1, 1, lastCol).getValues()[0];

  const headerMap = ensureColumnsExist_(targetSheet, [CONFIG.PERIOD_COLUMN_NAME].concat(sourceHeader));
  const totalCols = targetSheet.getLastColumn();
  const timestamp = new Date();
  const today = Utilities.formatDate(timestamp, CONFIG.TIME_ZONE, 'yyyy/MM/dd');

  const outRow = new Array(totalCols).fill('');
  outRow[headerMap['取込日時'] - 1] = timestamp;
  outRow[headerMap['元ファイル名'] - 1] = '(シート編集) ' + sourceSheet.getName();
  if (headerMap[CONFIG.PERIOD_COLUMN_NAME]) {
    outRow[headerMap[CONFIG.PERIOD_COLUMN_NAME] - 1] = today + ' 〜 ' + today;
  }
  sourceHeader.forEach(function (colName, idx) {
    if (!colName) return;
    const col = headerMap[colName];
    if (col) outRow[col - 1] = rowValues[idx];
  });

  const lastRow = targetSheet.getLastRow();
  targetSheet.getRange(lastRow + 1, 1, 1, totalCols).setValues([outRow]);
}

/**
 * 「週次ステータス変更履歴」シートを取得する。存在しない場合は作成し、ヘッダーを設定する。
 */
function getStatusHistorySheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.TARGET_SPREADSHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.STATUS_HISTORY_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.STATUS_HISTORY_SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 8).setValues([[
      '変更日時', 'シート名', '行番号', '列名', '変更前', '変更後', '編集者', '識別情報'
    ]]);
  }
  return sheet;
}

/**
 * 見込確度などの変更を「週次ステータス変更履歴」に1行追記する。
 * 識別情報には、対象行の1列目からEDIT_IDENTIFIER_COLUMN_COUNT列目までの値のうち
 * 空でないものだけを" / "区切りで連結して記録する(案件名・得意先名などを想定)。
 */
function appendStatusHistory_(sourceSheet, row, oldValue, newValue, editorEmail, historySheet) {
  const idCount = Math.min(CONFIG.EDIT_IDENTIFIER_COLUMN_COUNT, CONFIG.EDIT_TARGET_COLUMN - 1);
  const idValues = idCount > 0 ? sourceSheet.getRange(row, 1, 1, idCount).getValues()[0] : [];
  const identifier = idValues
    .map(function (v) { return String(v).trim(); })
    .filter(function (v) { return v !== ''; })
    .join(' / ');

  const columnName = getColumnHeaderName_(sourceSheet, CONFIG.EDIT_TARGET_COLUMN);

  historySheet.appendRow([
    new Date(),
    sourceSheet.getName(),
    row,
    columnName,
    oldValue,
    newValue,
    editorEmail,
    identifier
  ]);
}

/**
 * 指定列のヘッダー名(EDIT_HEADER_ROW行目)を取得する。取得できない場合は「列N」を返す。
 */
function getColumnHeaderName_(sheet, col) {
  try {
    const value = sheet.getRange(CONFIG.EDIT_HEADER_ROW, col).getValue();
    const name = String(value).trim();
    return name !== '' ? name : ('列' + col);
  } catch (err) {
    return '列' + col;
  }
}
