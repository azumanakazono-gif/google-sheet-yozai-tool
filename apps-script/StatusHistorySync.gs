/**
 * 「今月」を含む名前のシート（例: 「【 今月 】」など、全角/半角スペースや装飾記号付き）で
 * ステータス列（O列 = 15列目）が編集された時に、変更内容を「週次ステータス変更履歴」シートへ
 * 自動転記するスクリプト。
 *
 * - シート名判定は「今月」を"含むか"の部分一致で行い、シート名中のスペース（全角/半角問わず）は
 *   判定前に除去するため、「今月」「 今月 」「【 今月 】」等の表記ゆれをすべて検知できる。
 * - 対象列(EDIT_TARGET_COLUMN)・対象行(EDIT_MIN_ROW以上)の判定を、単一セル編集だけでなく
 *   複数セル貼り付け(範囲編集)でも取りこぼさないよう、編集範囲全体を走査して判定する。
 * - myOnEdit(e) はインストール型トリガー（installStatusHistoryEditTrigger で設定）として
 *   動かす想定。編集者に「週次ステータス変更履歴」への書き込み権限が無くても、
 *   スクリプト所有者権限で確実に転記できるようにするため。
 */

// ===== 編集時トリガー(onEdit)設定 =====
const CONFIG_STATUS_HISTORY = {
  EDIT_SHEET_NAME_KEYWORD: '今月',
  EDIT_MIN_ROW: 5,
  EDIT_TARGET_COLUMN: 15,
  EDIT_HEADER_ROW: 4,
  EDIT_IDENTIFIER_COLUMN_COUNT: 14,
  STATUS_HISTORY_SHEET_NAME: '週次ステータス変更履歴'
};

/**
 * インストール型onEditトリガーから呼び出されるハンドラー。
 * 対象シート(「今月」を含む名前)・対象列(EDIT_TARGET_COLUMN)・対象行(EDIT_MIN_ROW以上)を
 * 厳密に判定し、条件を満たす変更のみ「週次ステータス変更履歴」へ転記する。
 */
function myOnEdit(e) {
  try {
    if (!e || !e.range) {
      Logger.log('myOnEdit: イベントオブジェクトが不正のため終了します。');
      return;
    }

    const sheet = e.range.getSheet();
    if (!isTargetSheetName_(sheet.getName())) return;

    const editedRange = e.range;
    const firstRow = editedRange.getRow();
    const lastRow = firstRow + editedRange.getNumRows() - 1;
    const firstCol = editedRange.getColumn();
    const lastCol = firstCol + editedRange.getNumColumns() - 1;

    // 編集範囲に対象列(O列)が含まれていない場合は何もしない(単一セル編集・範囲貼り付けの両方に対応)
    const targetCol = CONFIG_STATUS_HISTORY.EDIT_TARGET_COLUMN;
    if (targetCol < firstCol || targetCol > lastCol) return;

    const isSingleCellEdit = editedRange.getNumRows() === 1 && editedRange.getNumColumns() === 1;
    const editor = getEditorEmail_(e);
    const timestamp = new Date();
    const historyRows = [];

    for (let row = firstRow; row <= lastRow; row++) {
      // ヘッダー行や見出し領域(EDIT_MIN_ROW未満)は対象外
      if (row < CONFIG_STATUS_HISTORY.EDIT_MIN_ROW) continue;

      const newValue = sheet.getRange(row, targetCol).getValue();
      const oldValue = isSingleCellEdit && row === firstRow ? (e.oldValue !== undefined ? e.oldValue : '') : '';

      // 変更前後が同一(実質変化なし)の場合は記録しない
      if (isSingleCellEdit && String(oldValue) === String(newValue)) continue;

      historyRows.push([
        timestamp,
        sheet.getName(),
        row,
        buildRowIdentifier_(sheet, row),
        oldValue,
        newValue,
        editor
      ]);
    }

    if (historyRows.length === 0) return;

    const historySheet = getStatusHistorySheet_(e.source);
    historySheet
      .getRange(historySheet.getLastRow() + 1, 1, historyRows.length, historyRows[0].length)
      .setValues(historyRows);
  } catch (err) {
    Logger.log('myOnEdit failed: ' + err);
  }
}

/**
 * シート名に EDIT_SHEET_NAME_KEYWORD("今月")が含まれるかを判定する。
 * 全角/半角スペースなど空白文字は判定前に除去するため、
 * 「今月」「 今月 」「【 今月 】」のような表記ゆれをすべて同一視できる。
 */
function isTargetSheetName_(sheetName) {
  const normalized = normalizeSheetNameForMatch_(sheetName);
  const keyword = normalizeSheetNameForMatch_(CONFIG_STATUS_HISTORY.EDIT_SHEET_NAME_KEYWORD);
  return keyword !== '' && normalized.indexOf(keyword) !== -1;
}

/**
 * 全角スペース(　)・半角スペース・タブ等の空白文字をすべて除去する。
 */
function normalizeSheetNameForMatch_(text) {
  return String(text).replace(/[\s　]/g, '');
}

/**
 * EDIT_HEADER_ROW のヘッダー名と、対象行のA列〜EDIT_IDENTIFIER_COLUMN_COUNT列目までの値を
 * 「ヘッダー名: 値」の形式で連結し、変更行を一意に識別できる文字列を作る。
 */
function buildRowIdentifier_(sheet, row) {
  const count = CONFIG_STATUS_HISTORY.EDIT_IDENTIFIER_COLUMN_COUNT;
  const headers = sheet.getRange(CONFIG_STATUS_HISTORY.EDIT_HEADER_ROW, 1, 1, count).getValues()[0];
  const values = sheet.getRange(row, 1, 1, count).getValues()[0];

  return headers
    .map(function (header, idx) {
      const label = String(header).trim();
      const value = values[idx];
      return label ? label + ': ' + value : String(value);
    })
    .filter(function (part) { return part !== ''; })
    .join(' / ');
}

/**
 * インストール型トリガーの場合は編集者のメールアドレスを取得できるため優先し、
 * 取得できない場合(権限不足・シンプルトリガー実行時等)はアクティブユーザーにフォールバックする。
 */
function getEditorEmail_(e) {
  try {
    if (e.user && e.user.getEmail) {
      const email = e.user.getEmail();
      if (email) return email;
    }
  } catch (err) {
    // 権限不足等で取得できない場合は下のフォールバックに委ねる
  }
  try {
    return Session.getActiveUser().getEmail() || '';
  } catch (err) {
    return '';
  }
}

/**
 * 「週次ステータス変更履歴」シートを取得する。存在しない場合はヘッダー付きで新規作成する。
 */
function getStatusHistorySheet_(spreadsheet) {
  const name = CONFIG_STATUS_HISTORY.STATUS_HISTORY_SHEET_NAME;
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(name);
    sheet.getRange(1, 1, 1, 7).setValues([['日時', 'シート名', '行番号', '識別子', '変更前', '変更後', '編集者']]);
  }
  return sheet;
}

// ===== トリガー設定 =====

/**
 * myOnEdit をインストール型トリガーとして設定する(初回に一度だけ実行する)。
 * シンプルトリガー(onEdit(e))ではなくインストール型にしているのは、
 * 編集者の権限に関わらず「週次ステータス変更履歴」への書き込みを確実に行うため。
 * 既存の myOnEdit 用トリガーがあれば一旦削除してから作り直す。
 */
function installStatusHistoryEditTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === 'myOnEdit') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('myOnEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
}
