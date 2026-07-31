/**
 * 予材リスト週次同期スクリプト（単一ファイル版）
 *
 * - SOURCE_FOLDER_ID 配下のANDPAD出力Excelファイル（Googleスプレッドシートも可）を読み取り、
 *   「31期予材リスト」の「週次報告記録」シート末尾に追記する。
 * - Excelファイルは DriveApp + Utilities.unzip + XmlService のみで直接パースする（Advanced Drive
 *   Service、UrlFetchApp、外部API呼び出しは一切使わない）。ANDPAD等の「拡張子は.xlsxだが中身は
 *   標準ZIPではない(HTMLテーブル等)」出力にも対応するため、先頭バイトで実体を判別し、
 *   ZIP/.xlsx・HTMLテーブルの2通りの読み取り方法へ自動でフォールバックする。
 * - 列はヘッダー名でマッチングするため、ANDPAD側の列順変更や列追加に強い。
 * - 転記済みファイルは「処理済み」フォルダへ移動し、同名ファイルがあればリネームして衝突を回避する。
 * - 万一の移動失敗に備え、処理済みファイルIDをスクリプトプロパティに記録し、再走時の二重追記を防ぐ。
 * - syncWeeklyReportsの最後に、「週次報告記録」を「報告日」列(ヘッダー名で検出、見つからない
 *   場合はI列=9列目)で昇順ソートする。そのうえで、シート全体のデータから対象期間(最小の
 *   報告日〜最大の報告日。報告日列が判定できない場合は「元ファイル名」列に含まれる日付から
 *   フォールバック)を算出し、その開始日・終了日をそれぞれが属する週の月曜日・日曜日まで
 *   拡張したうえで、「2行目」に「【対象期間：2026/07/27 〜 2026/08/02】」のような区切り行を
 *   1行だけ挿入する(A〜S列を結合し、背景色と太字で装飾)。区切り行は常に1行のみで、新しい
 *   ファイルを取り込むたびに削除・再算出・再挿入されるため、複数の区切り行が並んだり
 *   ソート後に末尾へ移動したりすることはない。
 * - 「今月」を含むシート名の5行目以降・O列(見込確度)の編集を検知し、「週次報告記録」と
 *   「週次ステータス変更履歴」へ自動転記する(onEdit)。この機能は、このスクリプトが
 *   「31期予材リスト」のコンテナバインド型スクリプトとして設置されている場合のみ動作する。
 *
 * 事前準備:
 *   1. appsscript.json の内容をマニフェストに反映する（エディタで「"appsscript.json"
 *      マニフェスト ファイルをエディタで表示する」を有効にすると編集できる）。
 *      GASの「サービス」から何かを追加する必要はない。
 *   2. このファイルの内容を「コード.gs」に丸ごと貼り付ける（既存の中身は全削除してから貼り付けること）。
 *      貼り付け後は Ctrl+End (Mac: Cmd+End) でファイル末尾に移動し、最後の "}" の後に
 *      余分な文字が残っていないことを目視確認する。
 */

// ===== 設定値 =====
const CONFIG = {
  // データ取得元: Excel/スプレッドシートが置かれているフォルダ
  SOURCE_FOLDER_ID: '1WrjMUtIpe2JEwChiJygRs1TeqWGgPuyc',

  // 二重処理防止用: 処理済み/失敗ファイルの退避先フォルダ名
  // (SOURCE_FOLDER_ID の直下にサブフォルダとして自動作成される)
  PROCESSED_FOLDER_NAME: '処理済み',
  ERROR_FOLDER_NAME: 'エラー',

  // 追記先: 「31期予材リスト」スプレッドシート
  TARGET_SPREADSHEET_ID: '1zTz2lLUD6M4SPBCcEO3OBxNMmv1U7RsUYRJhkKkgOOQ',
  TARGET_SHEET_NAME: '週次報告記録',

  // 取り込み元ファイルの1行目がヘッダー行かどうか
  SOURCE_HAS_HEADER: true,

  // エラー発生時に通知したいメールアドレス。空文字なら通知しない。
  NOTIFY_EMAIL: '',

  TIME_ZONE: 'Asia/Tokyo',

  // 処理済みファイルIDを記録しておく数（moveTo失敗時などの二重処理を防ぐ保険）
  PROCESSED_LOG_MAX: 300,

  // ===== 対象期間の区切り行 =====
  // 対象期間はデータ列(J列等)には記録せず、シート全体のデータの先頭(2行目)に
  // 1行だけの区切り行として表示する(複数ファイルにまたがっても、最小の報告日〜最大の
  // 報告日を1行に集約する)。ラベルの接頭辞・接尾辞（例:「【対象期間：2026/07/20 〜 2026/07/26】」）
  PERIOD_SEPARATOR_PREFIX: '【対象期間：',
  PERIOD_SEPARATOR_SUFFIX: '】',
  // 対象期間が判定できなかった場合に区切り行へ表示する文言
  PERIOD_UNKNOWN_LABEL: '不明',
  // 区切り行をA列からこの列数まで結合する（既定: 19列目 = S列）
  PERIOD_SEPARATOR_MERGE_COLUMNS: 19,
  // 区切り行の背景色
  PERIOD_SEPARATOR_BACKGROUND_COLOR: '#f3f3f3',

  // ===== 報告日順の自動ソート・対象期間の集計に使う列 =====
  // ヘッダーからこの順で列名を探し、最初に見つかった列を「報告日」列として扱う。
  // どれも見つからない場合はREPORT_DATE_FALLBACK_COLUMN(既定: 9列目 = I列)を使う。
  REPORT_DATE_COLUMN_CANDIDATES: ['報告日', '日付', '対象日', '実施日', '報告日時', '登録日'],
  REPORT_DATE_FALLBACK_COLUMN: 9,

  // ===== 編集時トリガー(onEdit)設定 =====
  // シート名にこの文字列を含む場合のみ編集を監視する
  EDIT_SHEET_NAME_KEYWORD: '今月',
  // 監視対象の開始行（この行以降の編集のみを対象にする。見出し・集計行等を除外するため）
  EDIT_MIN_ROW: 5,
  // 監視対象の列（既定: O列 = 15列目, 見込確度）
  EDIT_TARGET_COLUMN: 15,
  // 「今月」シート側のヘッダー行番号（列名マッチングに使用。実際のシート構成に合わせて調整すること）
  EDIT_HEADER_ROW: 4,
  // 識別情報として転記する列数（1列目からこの列数まで。案件名・得意先名などの識別列を想定）
  EDIT_IDENTIFIER_COLUMN_COUNT: 14,
  // 1回の編集イベントで処理する最大行数（大量範囲貼り付け時のタイムアウト防止）
  EDIT_MAX_ROWS_PER_EVENT: 200,
  // ステータス変更履歴の記録先シート名
  STATUS_HISTORY_SHEET_NAME: '週次ステータス変更履歴'
};

// ===== メイン処理 (週次自動転記) =====

function syncWeeklyReports() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    Logger.log('別の実行が進行中のため終了します。');
    return;
  }

  try {
    const sourceFolder = DriveApp.getFolderById(CONFIG.SOURCE_FOLDER_ID);
    const processedFolder = getOrCreateSubFolder_(sourceFolder, CONFIG.PROCESSED_FOLDER_NAME);
    const errorFolder = getOrCreateSubFolder_(sourceFolder, CONFIG.ERROR_FOLDER_NAME);
    const targetSheet = getTargetSheet_();

    const files = sourceFolder.getFiles();
    const errors = [];
    let processedCount = 0;
    let skippedCount = 0;

    while (files.hasNext()) {
      const file = files.next();

      if (!isSupportedSpreadsheet_(file)) continue;
      if (file.getId() === CONFIG.TARGET_SPREADSHEET_ID) continue;

      if (isAlreadyProcessed_(file.getId())) {
        Logger.log('既に処理済み(記録あり)のためスキップ、処理済みフォルダへ移動します: ' + file.getName());
        moveToFolderSafely_(file, processedFolder);
        skippedCount++;
        continue;
      }

      try {
        appendFileToTargetSheet_(file, targetSheet);
        // 追記が終わった時点で記録する。この後の移動が失敗しても次回は二重追記されない。
        markAsProcessed_(file.getId());
        moveToFolderSafely_(file, processedFolder);
        processedCount++;
      } catch (err) {
        Logger.log('処理失敗: ' + file.getName() + ' / ' + err);
        errors.push(file.getName() + ': ' + err);
        try {
          moveToFolderSafely_(file, errorFolder);
        } catch (moveErr) {
          Logger.log('エラーフォルダへの移動にも失敗: ' + file.getName() + ' / ' + moveErr);
        }
      }
    }

    if (processedCount > 0) {
      // 既存の対象期間区切り行を取り除いてから報告日順にソートし、シート全体のデータから
      // 算出した対象期間(最小の報告日〜最大の報告日)を表す区切り行を先頭(2行目)に挿入し直す。
      removeExistingPeriodSeparatorRows_(targetSheet);
      sortWeeklyReportByReportDate_(targetSheet);
      insertPeriodSummaryRow_(targetSheet);
    }

    Logger.log(
      '処理完了: ' + processedCount + '件 / スキップ: ' + skippedCount + '件 / 失敗: ' + errors.length + '件'
    );

    if (errors.length > 0 && CONFIG.NOTIFY_EMAIL) {
      MailApp.sendEmail(
        CONFIG.NOTIFY_EMAIL,
        '[予材リスト週次同期] 処理エラーのお知らせ',
        '以下のファイルの処理に失敗しました。「' + CONFIG.ERROR_FOLDER_NAME + '」フォルダをご確認ください。\n\n' +
          errors.join('\n')
      );
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * 1ファイル分のデータを読み取り、対象シートの末尾に追記する。
 * 列はヘッダー名でマッチングするため、ANDPAD側の列順・列追加の揺れを吸収する。
 * 対象期間の区切り行はここでは挿入しない(syncWeeklyReportsが全ファイル処理後に
 * removeExistingPeriodSeparatorRows_ / sortWeeklyReportByReportDate_ / insertPeriodSummaryRow_
 * を通じて、シート全体のデータから算出した区切り行を先頭にまとめて1行だけ挿入する)。
 */
function appendFileToTargetSheet_(file, targetSheet) {
  const mimeType = file.getMimeType();
  let values;

  if (mimeType === MimeType.GOOGLE_SHEETS) {
    try {
      values = SpreadsheetApp.openById(file.getId()).getSheets()[0].getDataRange().getValues();
    } catch (e) {
      throw new Error('Googleスプレッドシートとして開けませんでした(id: ' + file.getId() + ')。詳細: ' + e);
    }
  } else {
    values = readXlsxAsValues_(file);
  }

  if (values.length === 0) {
    Logger.log('データなし: ' + file.getName());
    return;
  }

  const hasHeader = CONFIG.SOURCE_HAS_HEADER;
  const rawHeader = hasHeader ? values[0] : values[0].map(function (_, idx) { return '列' + (idx + 1); });
  const sourceHeader = rawHeader.map(function (h) { return String(h).trim(); });
  const dataRows = (hasHeader ? values.slice(1) : values).filter(function (row) {
    return row.some(function (cell) { return cell !== '' && cell !== null; });
  });

  if (dataRows.length === 0) {
    Logger.log('データ行なし: ' + file.getName());
    return;
  }

  const headerMap = ensureColumnsExist_(targetSheet, sourceHeader);
  const totalCols = targetSheet.getLastColumn();
  const timestamp = new Date();

  const rowsToAppend = dataRows.map(function (row) {
    const outRow = new Array(totalCols).fill('');
    outRow[headerMap['取込日時'] - 1] = timestamp;
    outRow[headerMap['元ファイル名'] - 1] = file.getName();
    sourceHeader.forEach(function (colName, idx) {
      if (!colName) return;
      const col = headerMap[colName];
      if (col) outRow[col - 1] = row[idx];
    });
    return outRow;
  });

  const lastRow = targetSheet.getLastRow();
  targetSheet.getRange(lastRow + 1, 1, rowsToAppend.length, totalCols).setValues(rowsToAppend);
}

/**
 * 対象シートのヘッダー行(1行目)に、渡された列名のうち未登録のものを追加する。
 * ヘッダーが空の場合は「取込日時」「元ファイル名」から作成する。
 * 戻り値: { 列名: 列番号(1始まり) } のマップ
 */
function ensureColumnsExist_(targetSheet, headerNames) {
  const lastCol = targetSheet.getLastColumn();
  let headerRow = lastCol > 0 ? targetSheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  let existing = headerRow.map(function (h) { return String(h).trim(); }).filter(function (h) { return h !== ''; });

  if (existing.length === 0) {
    existing = ['取込日時', '元ファイル名'];
  }

  let changed = existing.length !== headerRow.length;
  headerNames.forEach(function (name) {
    const trimmed = String(name).trim();
    if (trimmed === '') return;
    if (existing.indexOf(trimmed) === -1) {
      existing.push(trimmed);
      changed = true;
    }
  });

  if (changed) {
    targetSheet.getRange(1, 1, 1, existing.length).setValues([existing]);
  }

  const map = {};
  existing.forEach(function (name, idx) { map[name] = idx + 1; });
  return map;
}

// ===== 対象期間の区切り行 =====

/**
 * セルの値(Date/Excelシリアル値/文字列)をDateに変換する。変換できない場合はnullを返す。
 */
function toDateValue_(value) {
  if (value instanceof Date) {
    return isNaN(value.getTime()) ? null : value;
  }
  if (typeof value === 'number') {
    // Excelのシリアル値(1899-12-30起点)をDateに変換する。極端な値は日付列とみなさない。
    if (value <= 0 || value > 100000) return null;
    const epoch = Date.UTC(1899, 11, 30);
    const date = new Date(epoch + value * 86400000);
    return isNaN(date.getTime()) ? null : date;
  }
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (trimmed === '') return null;
    const match = trimmed.match(/^(\d{4})[\/\-年](\d{1,2})[\/\-月](\d{1,2})/);
    if (match) {
      const date = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
      return isNaN(date.getTime()) ? null : date;
    }
    const parsed = new Date(trimmed);
    return isNaN(parsed.getTime()) ? null : parsed;
  }
  return null;
}

/**
 * 開始日・終了日を「yyyy/MM/dd 〜 yyyy/MM/dd」形式の文字列にする。
 */
function formatPeriod_(start, end) {
  const from = Utilities.formatDate(start, CONFIG.TIME_ZONE, 'yyyy/MM/dd');
  const to = Utilities.formatDate(end, CONFIG.TIME_ZONE, 'yyyy/MM/dd');
  return from + ' 〜 ' + to;
}

/**
 * 対象期間の区切り行を作る。A列に「【対象期間：yyyy/MM/dd 〜 yyyy/MM/dd】」を入れ、
 * それ以外の列はすべて空欄にする。対象期間が判定できなかった場合はPERIOD_UNKNOWN_LABELを表示する。
 */
function buildPeriodSeparatorRow_(periodValue, totalCols) {
  const row = new Array(totalCols).fill('');
  row[0] = CONFIG.PERIOD_SEPARATOR_PREFIX + (periodValue || CONFIG.PERIOD_UNKNOWN_LABEL) + CONFIG.PERIOD_SEPARATOR_SUFFIX;
  return row;
}

/**
 * 区切り行をA列からPERIOD_SEPARATOR_MERGE_COLUMNS列目(既定: S列)まで結合し、
 * 背景色(PERIOD_SEPARATOR_BACKGROUND_COLOR)と太字で装飾する。
 * データ列数(totalCols)がPERIOD_SEPARATOR_MERGE_COLUMNS未満でも結合できるよう、
 * 事前にシートの列数を必要分だけ拡張する。この処理は指定した1行のみに閉じており、
 * データの書き込み(setValues)には影響しない。
 */
function formatPeriodSeparatorRow_(targetSheet, rowNumber) {
  const mergeCols = CONFIG.PERIOD_SEPARATOR_MERGE_COLUMNS;
  ensureMinColumns_(targetSheet, mergeCols);

  const range = targetSheet.getRange(rowNumber, 1, 1, mergeCols);
  range.setBackground(CONFIG.PERIOD_SEPARATOR_BACKGROUND_COLOR);
  range.setFontWeight('bold');
  range.merge();
}

/**
 * シートの列数がminColsに満たない場合、不足分の列を末尾に追加する。
 */
function ensureMinColumns_(sheet, minCols) {
  const currentMax = sheet.getMaxColumns();
  if (currentMax < minCols) {
    sheet.insertColumnsAfter(currentMax, minCols - currentMax);
  }
}

// ===== 報告日順の自動ソートと対象期間区切り行の再構築 =====

/**
 * シート上に残っている対象期間の区切り行(A列がPERIOD_SEPARATOR_PREFIXで始まる行)を
 * すべて削除する。syncWeeklyReportsが新規ファイルを処理するたびに、この関数で古い
 * 区切り行を取り除いてからソート・再集計・再挿入することで、区切り行が常に1行だけ、
 * かつ先頭(2行目)に保たれる。
 */
function removeExistingPeriodSeparatorRows_(targetSheet) {
  const lastRow = targetSheet.getLastRow();
  if (lastRow < 2) return;

  const colAValues = targetSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const rowsToDelete = [];
  colAValues.forEach(function (row, idx) {
    const value = String(row[0]);
    if (value.indexOf(CONFIG.PERIOD_SEPARATOR_PREFIX) === 0) {
      rowsToDelete.push(idx + 2);
    }
  });

  // 行番号のズレを避けるため、後ろの行から削除する。
  rowsToDelete.sort(function (a, b) { return b - a; });
  rowsToDelete.forEach(function (rowNum) {
    targetSheet.deleteRow(rowNum);
  });
}

/**
 * 「週次報告記録」の2行目以降を「報告日」列(見つからなければI列=9列目)で昇順ソートする。
 * このソートは removeExistingPeriodSeparatorRows_ で区切り行を取り除いた後に呼ぶ前提だが、
 * 万一結合セルが残っていてもソートできるよう、念のためソート前に対象範囲の結合を解除する
 * (GASの仕様: 'You are trying to sort a range that contains merged cells')。
 */
function sortWeeklyReportByReportDate_(targetSheet) {
  const firstDataRow = 2;
  const lastRow = targetSheet.getLastRow();
  if (lastRow < firstDataRow) return;

  const contentLastCol = targetSheet.getLastColumn();
  const sortColumn = resolveReportDateColumn_(targetSheet, contentLastCol);
  if (!sortColumn) {
    Logger.log('報告日列が見つからないため、並べ替えをスキップします。');
    return;
  }

  const dataRange = targetSheet.getRange(firstDataRow, 1, lastRow - firstDataRow + 1, contentLastCol);
  dataRange.breakApart();
  dataRange.sort({ column: sortColumn, ascending: true });
}

/**
 * ヘッダー行からREPORT_DATE_COLUMN_CANDIDATES(既定:「報告日」「日付」等)の順で列を探す。
 * どれも見つからない場合はREPORT_DATE_FALLBACK_COLUMN(既定: 9列目 = I列)を使う。
 */
function resolveReportDateColumn_(targetSheet, lastCol) {
  if (lastCol > 0) {
    const headerRow = targetSheet.getRange(1, 1, 1, lastCol).getValues()[0]
      .map(function (h) { return String(h).trim(); });
    for (let c = 0; c < CONFIG.REPORT_DATE_COLUMN_CANDIDATES.length; c++) {
      const idx = headerRow.indexOf(CONFIG.REPORT_DATE_COLUMN_CANDIDATES[c]);
      if (idx !== -1) return idx + 1;
    }
  }
  return CONFIG.REPORT_DATE_FALLBACK_COLUMN;
}

/**
 * シート全体のデータ(2行目以降)から対象期間(最小の報告日〜最大の報告日)を算出し、
 * 先頭(2行目)に区切り行として挿入する。ソート済みであることを前提とせず、シートの
 * 現在の内容だけから算出するため、常に「今シートにある全データを覆う最も広い期間」になる。
 */
function insertPeriodSummaryRow_(targetSheet) {
  const lastRow = targetSheet.getLastRow();
  if (lastRow < 2) return;

  const contentLastCol = targetSheet.getLastColumn();
  const periodValue = computeOverallPeriodLabel_(targetSheet, contentLastCol, lastRow);

  targetSheet.insertRowBefore(2);
  const separatorRow = buildPeriodSeparatorRow_(periodValue, contentLastCol);
  targetSheet.getRange(2, 1, 1, contentLastCol).setValues([separatorRow]);
  formatPeriodSeparatorRow_(targetSheet, 2);
}

/**
 * シート全体のデータから対象期間を算出し、「2026/07/20 〜 2026/07/26」形式の文字列で返す。
 * まず「報告日」列(resolveReportDateColumn_)の値から最小値・最大値を求め、判定できない
 * 場合は「元ファイル名」列に含まれる日付(ファイル名パターン)からフォールバックする。
 * どちらからも判定できない場合は空文字を返す。
 * 開始日・終了日はそれぞれの日付が属する週の月曜日・日曜日まで拡張する
 * (例: データが2026/07/28〜2026/07/30なら「2026/07/27 〜 2026/08/02」)。
 */
function computeOverallPeriodLabel_(targetSheet, contentLastCol, lastRow) {
  const dateColumn = resolveReportDateColumn_(targetSheet, contentLastCol);
  const range = extractDateRangeFromColumn_(targetSheet, dateColumn, lastRow)
    || extractDateRangeFromFileNameColumn_(targetSheet, contentLastCol, lastRow);
  if (!range) return '';
  return formatPeriod_(getMondayOfWeek_(range.min), getSundayOfWeek_(range.max));
}

/**
 * 指定日を含む週(月曜始まり)の月曜日を返す。
 */
function getMondayOfWeek_(date) {
  const day = date.getDay(); // 0=日,1=月,2=火,...,6=土
  const diffToMonday = (day + 6) % 7;
  return new Date(date.getFullYear(), date.getMonth(), date.getDate() - diffToMonday);
}

/**
 * 指定日を含む週(月曜始まり)の日曜日を返す。
 */
function getSundayOfWeek_(date) {
  const day = date.getDay(); // 0=日,1=月,2=火,...,6=土
  const diffToSunday = (7 - day) % 7;
  return new Date(date.getFullYear(), date.getMonth(), date.getDate() + diffToSunday);
}

/**
 * 指定列(2行目以降)の値をDateに変換し、その最小値・最大値を返す。判定できる値が
 * 1件もない場合はnullを返す。
 */
function extractDateRangeFromColumn_(targetSheet, column, lastRow) {
  if (!column) return null;
  const values = targetSheet.getRange(2, column, lastRow - 1, 1).getValues();

  let min = null;
  let max = null;
  values.forEach(function (row) {
    const date = toDateValue_(row[0]);
    if (!date) return;
    if (!min || date.getTime() < min.getTime()) min = date;
    if (!max || date.getTime() > max.getTime()) max = date;
  });

  return (min && max) ? { min: min, max: max } : null;
}

/**
 * 「元ファイル名」列(2行目以降)からファイル名パターンの日付を抽出し、その最小値・最大値を
 * 返す。「元ファイル名」列自体が見つからない場合、または判定できる値が1件もない場合はnullを返す。
 */
function extractDateRangeFromFileNameColumn_(targetSheet, contentLastCol, lastRow) {
  const headerRow = targetSheet.getRange(1, 1, 1, contentLastCol).getValues()[0]
    .map(function (h) { return String(h).trim(); });
  const fileNameCol = headerRow.indexOf('元ファイル名') + 1;
  if (fileNameCol === 0) return null;

  const values = targetSheet.getRange(2, fileNameCol, lastRow - 1, 1).getValues();

  let min = null;
  let max = null;
  values.forEach(function (row) {
    const bounds = extractDateBoundsFromFileName_(row[0]);
    if (!bounds) return;
    if (!min || bounds.start.getTime() < min.getTime()) min = bounds.start;
    if (!max || bounds.end.getTime() > max.getTime()) max = bounds.end;
  });

  return (min && max) ? { min: min, max: max } : null;
}

/**
 * ファイル名に含まれる2つの8桁日付(yyyyMMdd)を開始日・終了日として取り出す。
 * 例: 報告一覧_20260720_20260726.xlsx → { start: 2026/07/20, end: 2026/07/26 }
 * 判定できない場合はnullを返す。
 */
function extractDateBoundsFromFileName_(fileName) {
  const base = String(fileName).replace(/\.[^.]+$/, '');
  const match = base.match(/(\d{4})(\d{2})(\d{2}).*?(\d{4})(\d{2})(\d{2})/);
  if (!match) return null;

  const start = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
  const end = new Date(Number(match[4]), Number(match[5]) - 1, Number(match[6]));
  if (isNaN(start.getTime()) || isNaN(end.getTime())) return null;

  return { start: start, end: end };
}

/**
 * DriveAppとGAS標準サービスだけでExcelファイルを読み取り、
 * SpreadsheetApp.getDataRange().getValues()相当の2次元配列を返す。
 * Advanced Drive Service・UrlFetchApp・外部API呼び出しは一切使わない。
 *
 * ANDPAD等の出力は「拡張子は.xlsxだが中身は標準ZIP構造ではない」ケースがあるため、
 * ファイル先頭のマジックバイトで実体を判別し、3通りの読み取り方法にフォールバックする。
 *   1. 本物の.xlsx(ZIP構造): Utilities.unzip + XmlServiceで直接パース。
 *   2. 旧形式バイナリ.xls(OLE構造): パース不可のため、その旨を明確なエラーで通知する。
 *   3. HTMLテーブルを.xlsx/.xlsとして出力している場合: <table>をHTMLとして解析する。
 */
function readXlsxAsValues_(file) {
  const blob = file.getBlob();
  const bytes = blob.getBytes();

  if (looksLikeZip_(bytes)) {
    return readValuesFromXlsxZip_(blob);
  }

  if (looksLikeLegacyOle_(bytes)) {
    throw new Error(
      '旧形式のバイナリExcelファイル(.xls, Excel 97-2003形式)は読み取れません。.xlsx形式で保存し直してください。'
    );
  }

  const text = decodeTextBlob_(blob);
  if (/<html[\s>]|<table[\s>]/i.test(text.slice(0, 4000))) {
    return readValuesFromHtmlTable_(text);
  }

  throw new Error(
    '対応していないファイル形式です(ZIP/.xlsxでもHTMLでもありません)。ANDPAD側の出力設定やファイルの破損を確認してください。'
  );
}

/**
 * ZIPファイルの先頭マジックバイト("PK")かどうかを判定する。
 */
function looksLikeZip_(bytes) {
  return bytes.length >= 2 && bytes[0] === 0x50 && bytes[1] === 0x4b;
}

/**
 * 旧形式バイナリOffice文書(OLE Compound File)の先頭マジックバイトかどうかを判定する。
 */
function looksLikeLegacyOle_(bytes) {
  if (bytes.length < 4) return false;
  const b0 = bytes[0] < 0 ? bytes[0] + 256 : bytes[0];
  const b1 = bytes[1] < 0 ? bytes[1] + 256 : bytes[1];
  const b2 = bytes[2] < 0 ? bytes[2] + 256 : bytes[2];
  const b3 = bytes[3] < 0 ? bytes[3] + 256 : bytes[3];
  return b0 === 0xd0 && b1 === 0xcf && b2 === 0x11 && b3 === 0xe0;
}

/**
 * blobのテキストをUTF-8で読み、内部にShift_JIS系のcharset宣言が見つかればShift_JISで読み直す。
 * (ANDPAD等、日本語業務システムのHTML書き出しはShift_JISであることが多いため)
 */
function decodeTextBlob_(blob) {
  const utf8Text = blob.getDataAsString('UTF-8');
  const metaMatch = utf8Text.slice(0, 2000).match(/charset\s*=\s*["']?([\w-]+)/i);
  const charset = metaMatch ? metaMatch[1].toLowerCase() : '';
  if (/shift[-_]?jis|sjis|ms932|windows-31j/.test(charset)) {
    return blob.getDataAsString('Shift_JIS');
  }
  return utf8Text;
}

/**
 * 本物の.xlsx(ZIP構造)をUtilities.unzip + XmlServiceで直接パースする。
 */
function readValuesFromXlsxZip_(blob) {
  const zipBlob = blob.setContentType('application/zip');
  let entries;
  try {
    entries = Utilities.unzip(zipBlob);
  } catch (e) {
    throw new Error(
      '.xlsx(ZIP形式)として展開できませんでした。詳細: ' + e
    );
  }

  const entryMap = {};
  entries.forEach(function (entry) {
    entryMap[entry.getName()] = entry;
  });

  const sheetPath = resolveFirstSheetPath_(entryMap);
  const sheetEntry = entryMap[sheetPath];
  if (!sheetEntry) {
    throw new Error('ワークシートのXMLが見つかりませんでした(' + sheetPath + ')。');
  }

  const sharedStrings = parseSharedStrings_(entryMap['xl/sharedStrings.xml']);
  return parseWorksheetXml_(sheetEntry.getDataAsString('UTF-8'), sharedStrings);
}

/**
 * ANDPAD等が「Excelファイル」と称して実際にはHTMLテーブルを出力しているケース向けの
 * フォールバック。複数の<table>がある場合は行数が最も多いものをデータ本体とみなす。
 * XmlServiceでの厳密なXMLパースはHTMLの崩れ(閉じタグ漏れ等)で失敗しやすいため、
 * 正規表現ベースの緩いパースで抽出する。
 */
function readValuesFromHtmlTable_(html) {
  const tableBlocks = html.match(/<table[\s\S]*?<\/table>/gi) || [html];
  let bestBlock = tableBlocks[0];
  let bestRowCount = -1;
  tableBlocks.forEach(function (block) {
    const rowCount = (block.match(/<tr[\s>]/gi) || []).length;
    if (rowCount > bestRowCount) {
      bestRowCount = rowCount;
      bestBlock = block;
    }
  });

  const rowMatches = bestBlock.match(/<tr[\s\S]*?<\/tr>/gi) || [];
  const rows = rowMatches.map(function (rowHtml) {
    const cellMatches = rowHtml.match(/<t[dh][\s\S]*?<\/t[dh]>/gi) || [];
    return cellMatches.map(function (cellHtml) {
      const inner = cellHtml.replace(/^<t[dh][^>]*>/i, '').replace(/<\/t[dh]>$/i, '');
      const withBreaks = inner.replace(/<br\s*\/?>/gi, '\n');
      const stripped = withBreaks.replace(/<[^>]+>/g, '');
      return decodeHtmlEntities_(stripped).trim();
    });
  });

  const maxCol = rows.reduce(function (max, r) { return Math.max(max, r.length); }, 0);
  return rows.map(function (row) {
    const padded = row.slice();
    while (padded.length < maxCol) padded.push('');
    return padded;
  });
}

/**
 * HTMLの主要なエンティティ(&nbsp;等)をデコードする。
 */
function decodeHtmlEntities_(text) {
  return text
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/g, "'")
    .replace(/&#x([0-9a-fA-F]+);/g, function (_, hex) { return String.fromCharCode(parseInt(hex, 16)); })
    .replace(/&#(\d+);/g, function (_, dec) { return String.fromCharCode(parseInt(dec, 10)); });
}

/**
 * xl/workbook.xml と xl/_rels/workbook.xml.rels から、1番目のシート(タブ順で先頭)の
 * 実体XMLパスを求める。解決できない場合は慣例的な既定パスにフォールバックする。
 */
function resolveFirstSheetPath_(entryMap) {
  const fallback = 'xl/worksheets/sheet1.xml';
  const workbookEntry = entryMap['xl/workbook.xml'];
  const relsEntry = entryMap['xl/_rels/workbook.xml.rels'];
  if (!workbookEntry || !relsEntry) return fallback;

  try {
    const workbookDoc = XmlService.parse(workbookEntry.getDataAsString('UTF-8'));
    const ns = workbookDoc.getRootElement().getNamespace();
    const rIdNs = XmlService.getNamespace(
      'r',
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    );
    const sheetsEl = workbookDoc.getRootElement().getChild('sheets', ns);
    const firstSheet = sheetsEl.getChildren('sheet', ns)[0];
    const rId = firstSheet.getAttribute('id', rIdNs).getValue();

    const relsDoc = XmlService.parse(relsEntry.getDataAsString('UTF-8'));
    const relsNs = relsDoc.getRootElement().getNamespace();
    const relationships = relsDoc.getRootElement().getChildren('Relationship', relsNs);
    let target = null;
    relationships.forEach(function (rel) {
      if (rel.getAttribute('Id').getValue() === rId) {
        target = rel.getAttribute('Target').getValue();
      }
    });
    if (!target) return fallback;
    return target.indexOf('xl/') === 0 ? target : 'xl/' + target.replace(/^\/+/, '');
  } catch (e) {
    return fallback;
  }
}

/**
 * xl/sharedStrings.xml を読み、インデックス順の文字列配列を返す（存在しなければ空配列）。
 */
function parseSharedStrings_(entry) {
  if (!entry) return [];
  const doc = XmlService.parse(entry.getDataAsString('UTF-8'));
  const ns = doc.getRootElement().getNamespace();
  const siList = doc.getRootElement().getChildren('si', ns);
  return siList.map(function (si) {
    return extractText_(si, ns);
  });
}

/**
 * <si>要素からテキストを取り出す。直接<t>を持つ場合と、<r><t>...</t></r>の
 * リッチテキストの場合の両方に対応する。
 */
function extractText_(si, ns) {
  const directT = si.getChild('t', ns);
  if (directT) return directT.getText();
  const runs = si.getChildren('r', ns);
  return runs
    .map(function (r) {
      const t = r.getChild('t', ns);
      return t ? t.getText() : '';
    })
    .join('');
}

/**
 * xl/worksheets/sheetN.xml をパースし、getDataRange().getValues()相当の2次元配列を作る。
 * Excelは空セル・空行を省略して出力するため、行番号(r属性)・列文字(セル参照)から
 * 実際の位置に値を配置し、抜けている箇所は空文字で埋める。
 */
function parseWorksheetXml_(xmlText, sharedStrings) {
  const doc = XmlService.parse(xmlText);
  const ns = doc.getRootElement().getNamespace();
  const sheetData = doc.getRootElement().getChild('sheetData', ns);
  if (!sheetData) return [];

  const rows = sheetData.getChildren('row', ns);
  let maxCol = 0;

  const parsedRows = rows.map(function (row) {
    const rowIndex = parseInt(row.getAttribute('r').getValue(), 10);
    const cells = row.getChildren('c', ns);
    const rowValues = {};
    cells.forEach(function (cell) {
      const ref = cell.getAttribute('r').getValue();
      const letters = ref.match(/^[A-Za-z]+/)[0].toUpperCase();
      const colIndex = columnLetterToIndex_(letters);
      if (colIndex > maxCol) maxCol = colIndex;
      rowValues[colIndex] = extractCellValue_(cell, ns, sharedStrings);
    });
    return { rowIndex: rowIndex, values: rowValues };
  });

  const maxRow = parsedRows.reduce(function (max, r) { return Math.max(max, r.rowIndex); }, 0);
  const values = [];
  for (let r = 0; r < maxRow; r++) {
    values.push(new Array(maxCol).fill(''));
  }
  parsedRows.forEach(function (row) {
    Object.keys(row.values).forEach(function (colKey) {
      const colIndex = Number(colKey);
      values[row.rowIndex - 1][colIndex - 1] = row.values[colKey];
    });
  });

  return values;
}

/**
 * <c>(セル)要素から値を取り出す。t属性でshared string/真偽値/インライン文字列/数値を判別する。
 */
function extractCellValue_(cell, ns, sharedStrings) {
  const typeAttr = cell.getAttribute('t');
  const type = typeAttr ? typeAttr.getValue() : null;

  if (type === 'inlineStr') {
    const isEl = cell.getChild('is', ns);
    return isEl ? extractText_(isEl, ns) : '';
  }

  const vEl = cell.getChild('v', ns);
  if (!vEl) return '';
  const raw = vEl.getText();

  if (type === 's') {
    const idx = parseInt(raw, 10);
    return sharedStrings[idx] !== undefined ? sharedStrings[idx] : '';
  }
  if (type === 'b') {
    return raw === '1';
  }
  if (type === 'str' || type === 'e') {
    return raw;
  }
  const num = Number(raw);
  return isNaN(num) ? raw : num;
}

/**
 * "A"→1, "Z"→26, "AA"→27 のように列文字を1始まりの列番号に変換する。
 */
function columnLetterToIndex_(letters) {
  let index = 0;
  for (let i = 0; i < letters.length; i++) {
    index = index * 26 + (letters.charCodeAt(i) - 64);
  }
  return index;
}

/**
 * 対応する形式かどうかを判定する。
 * ANDPAD等の出力ではmimeTypeが正しく設定されないことがあるため、拡張子でもフォールバック判定する。
 */
function isSupportedSpreadsheet_(file) {
  const mimeType = file.getMimeType();
  if (
    mimeType === MimeType.GOOGLE_SHEETS ||
    mimeType === MimeType.MICROSOFT_EXCEL ||
    mimeType === 'application/vnd.ms-excel'
  ) {
    return true;
  }
  const name = file.getName().toLowerCase();
  return name.endsWith('.xlsx') || name.endsWith('.xls');
}

/**
 * フォルダ内に同名ファイルが既にある場合はタイムスタンプ付きにリネームしてから移動する。
 * ANDPADの出力はファイル名が固定/重複しがちなため、上書き混同を避ける。
 */
function moveToFolderSafely_(file, folder) {
  const existing = folder.getFilesByName(file.getName());
  if (existing.hasNext()) {
    const stamp = Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyyyMMdd_HHmmss');
    const name = file.getName();
    const dotIndex = name.lastIndexOf('.');
    const newName = dotIndex > -1
      ? name.slice(0, dotIndex) + '_' + stamp + name.slice(dotIndex)
      : name + '_' + stamp;
    file.setName(newName);
  }
  file.moveTo(folder);
}

function isAlreadyProcessed_(fileId) {
  return getProcessedIds_().indexOf(fileId) !== -1;
}

function markAsProcessed_(fileId) {
  const props = PropertiesService.getScriptProperties();
  const ids = getProcessedIds_();
  ids.push(fileId);
  const trimmed = ids.slice(-CONFIG.PROCESSED_LOG_MAX);
  props.setProperty('PROCESSED_FILE_IDS', JSON.stringify(trimmed));
}

function getProcessedIds_() {
  const raw = PropertiesService.getScriptProperties().getProperty('PROCESSED_FILE_IDS');
  return raw ? JSON.parse(raw) : [];
}

function getOrCreateSubFolder_(parentFolder, name) {
  const folders = parentFolder.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return parentFolder.createFolder(name);
}

function getTargetSheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.TARGET_SPREADSHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.TARGET_SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 2).setValues([['取込日時', '元ファイル名']]);
  }
  return sheet;
}

// ===== 編集時トリガー (onEdit) =====

/**
 * シンプルトリガー。関数名onEditはGASにおいて特別な意味を持ち、インストール操作なしで
 * このスプレッドシートが編集されるたびに自動実行される(手動実行や権限承認は不要)。
 * 動作するのは、このスクリプトが「31期予材リスト」スプレッドシートの
 * コンテナバインド型スクリプトとして設置されている場合のみ(スタンドアロン型では発火しない)。
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

  const headerMap = ensureColumnsExist_(targetSheet, sourceHeader);
  const totalCols = targetSheet.getLastColumn();
  const timestamp = new Date();

  const outRow = new Array(totalCols).fill('');
  outRow[headerMap['取込日時'] - 1] = timestamp;
  outRow[headerMap['元ファイル名'] - 1] = '(シート編集) ' + sourceSheet.getName();
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

// ===== トリガー設定 =====

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

/**
 * 処理済みファイルIDの記録をリセットする。
 * テストのため同じファイルをもう一度処理させたい場合などに、この関数だけを手動実行する。
 */
function resetProcessedLog() {
  PropertiesService.getScriptProperties().deleteProperty('PROCESSED_FILE_IDS');
  Logger.log('処理済みログをリセットしました。');
}
