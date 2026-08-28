/**
 * 予材リスト週次同期スクリプト（単一ファイル版）
 *
 * - SOURCE_FOLDER_ID 配下のANDPAD出力Excelファイル（Googleスプレッドシートも可）を読み取り、
 *   「31期予材リスト」の「週次報告記録」シート末尾に追記する。
 * - Excelファイルは DriveApp + Utilities.unzip + XmlService のみで直接パースする（Advanced Drive
 *   Service、UrlFetchApp、外部API呼び出しは一切使わない）。ANDPAD等の「拡張子は.xlsxだが中身は
 *   標準ZIPではない(HTMLテーブルやCSV/TSV等)」出力にも対応するため、先頭バイトや中身のテキストで
 *   実体を判別し、ZIP/.xlsx・HTMLテーブル・CSV/TSVの読み取り方法へ自動でフォールバックする。
 * - 列はヘッダー名でマッチングするため、ANDPAD側の列順変更や列追加に強い。
 * - 転記済みファイルは「処理済み」フォルダへ移動し、同名ファイルがあればリネームして衝突を回避する。
 * - 万一の移動失敗に備え、処理済みファイルIDをスクリプトプロパティに記録し、再走時の二重追記を防ぐ。
 *
 * 事前準備:
 *   1. appsscript.json の内容をマニフェストに反映する（エディタで「"appsscript.json"
 *      マニフェスト ファイルをエディタで表示する」を有効にすると編集できる）。
 *      GASの「サービス」から何かを追加する必要はない。
 *   2. このファイルの内容を「コード.gs」に丸ごと貼り付ける（既存の中身は全削除してから貼り付けること）。
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
  TARGET_SHEET_NAME: '報告記録',

  // 週次ステータス変更履歴: 「週次報告記録」への追記直後に、直近の取り込み分から
  // 「見込確度」が変更前後で変化した案件だけを抽出して記録するシート名。
  STATUS_HISTORY_SHEET_NAME: '週次ステータス変更履歴',
  // 履歴シートの実際のレイアウト: 1行目=タイトル行、2行目=列見出し行、3行目以降=データ行。
  // ヘッダー(列見出し)が入っている行番号。ensureColumnsExist_・追記位置の判定の両方で使う。
  STATUS_HISTORY_HEADER_ROW: 2,
  // 「記録日時」列(A列)に書き込む日時文字列のフォーマット(Utilities.formatDate準拠)。
  // 日付だけでなく時刻まで含めて記録するため、Dateオブジェクトのままではなく
  // この書式の文字列に変換してから書き込む(buildStatusHistoryChange_を参照)。
  STATUS_HISTORY_TIMESTAMP_FORMAT: 'yyyy/MM/dd HH:mm:ss',
  // 履歴シートの列構成(既存シートの列順・列名に合わせてある)。
  // シートが空の場合はこの並びでヘッダーを新規作成し、既にヘッダーがある場合は
  // 名前が一致する列にだけ書き込む(他の既存列には影響しない)。
  // 「昇降フラグ」列はH列(変更前見込確度)・I列(変更後見込確度)を比較する数式が
  // シート側に組み込まれているため、GASからは値を書き込まない(既存行の数式を
  // コピーして新規行に適用するのみ。applyRankFlagFormula_を参照)。
  STATUS_HISTORY_COLUMNS: [
    '記録日時', '担当', '案件名', 'カテゴリ',
    '次回アクション予定日時', 'ネクストアクション', '昇降フラグ',
    '変更前見込確度', '変更後見込確度', '変更経緯',
    '最新見積格納日', '契約見込月', '税込想定売上', '貢献利益率'
  ],
  // 昇降フラグ列の名前。数式で↑/↓が算出される列のため、GASは値を書き込まず、
  // 既存行の数式をコピーして新規行に適用するためだけに使う。
  STATUS_HISTORY_RANK_FLAG_COLUMN: '昇降フラグ',
  // 案件名列(既存シートの見た目に合わせたリンク色)。既定はGoogle Sheetsのリンク標準色。
  STATUS_HISTORY_LINK_COLOR: '#1155cc',
  // 「週次報告記録」内で案件を一意に識別するための列名候補(先頭から順に探し、
  // 最初に見つかった列をキーとして使う)。実際のANDPAD出力列名に合わせて調整すること。
  STATUS_HISTORY_KEY_COLUMN_CANDIDATES: ['案件名'],
  // 転記条件となる「見込確度」列の名前候補(「週次報告記録」側の列名)。
  // 実際のANDPAD出力列名に合わせて調整すること。
  STATUS_HISTORY_CONFIDENCE_COLUMN_CANDIDATES: ['見込確度', '確度'],
  // 履歴シートの各列に、「週次報告記録」側のどの列の値を転記するかの対応表。
  // 値は「週次報告記録」側の列名候補の配列(先頭から探して最初に見つかったものを使う)。
  // 記録日時・案件名・昇降フラグ・変更前後の見込確度・変更経緯は別途組み立てるため、
  // ここには含めない。
  STATUS_HISTORY_FIELD_SOURCE_COLUMNS: {
    '担当': ['担当'],
    'カテゴリ': ['カテゴリ'],
    '次回アクション予定日時': ['次回アクション予定日時'],
    'ネクストアクション': ['ネクストアクション'],
    '最新見積格納日': ['最新見積格納日'],
    '契約見込月': ['契約見込月'],
    '税込想定売上': ['税込想定売上'],
    '貢献利益率': ['貢献利益率']
  },

  // 「今月」等の予材シートでO列(見込確度)が手動編集された際に、変更理由の入力ダイアログを
  // 表示して「週次ステータス変更履歴」に記録する機能(onConfidenceCellEdited)で使う設定。
  // 対象とする予材シート名(複数可)。実際の運用シート名に合わせて追加・変更すること。
  CONFIDENCE_EDIT_SHEET_NAMES: ['今月'],
  // 「今月」シートは列見出しが「週次報告記録」側と一致しない(または見出し行の構成が異なる)ため、
  // ヘッダー名でのマッチングは行わず、ここに設定した固定の列位置から直接値を抽出する。
  // 実際の「今月」シートの列位置に合わせて調整すること。
  CONFIDENCE_EDIT_COLUMN_LETTERS: {
    '案件名': 'H',
    '担当': 'G',
    'カテゴリ': 'P',
    '次回アクション予定日時': 'C',
    'ネクストアクション': 'D',
    '見込確度': 'O',
    '最新見積格納日': 'J',
    '契約見込月': 'R',
    '税込想定売上': 'K',
    '貢献利益率': 'L'
  },

  // 取り込み元ファイルの1行目がヘッダー行かどうか
  SOURCE_HAS_HEADER: true,

  // エラー発生時に通知したいメールアドレス。空文字なら通知しない。
  NOTIFY_EMAIL: '',

  TIME_ZONE: 'Asia/Tokyo',

  // 処理済みファイルIDを記録しておく数（moveTo失敗時などの二重処理を防ぐ保険）
  PROCESSED_LOG_MAX: 300
};

// ===== メイン処理 =====

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
    values = parseWeeklyReportFile_(file);
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
  // 変更検知は「今回追記する内容」と「追記前に既にシートにあった内容」を比較するため、
  // 追記より前にこの時点の既存データを読み取っておく。
  const previousRows = lastRow > 1
    ? targetSheet.getRange(2, 1, lastRow - 1, totalCols).getValues()
    : [];

  const appendRange = targetSheet.getRange(lastRow + 1, 1, rowsToAppend.length, totalCols);
  writeValuesIgnoringValidation_(appendRange, rowsToAppend);

  try {
    updateWeeklyStatusHistory_(targetSheet, lastRow, headerMap, previousRows, rowsToAppend, timestamp);
  } catch (historyErr) {
    // 履歴記録の失敗は「報告記録」への追記そのものを失敗させたくないため、ここで握りつぶす。
    Logger.log(
      '週次ステータス変更履歴の更新に失敗しました(報告記録への追記自体は成功しています): ' +
        file.getName() + ' / ' + historyErr
    );
  }
}

/**
 * 「週次報告記録」への直近の追記分(newRows)を、追記前の既存データ(previousRows)と
 * 案件単位で突き合わせ、「見込確度」(CONFIG.STATUS_HISTORY_CONFIDENCE_COLUMN_CANDIDATES)が
 * 変更前後で変化した案件だけを抽出して「週次ステータス変更履歴」シートの末尾に記録する。
 * (ステータスや進捗など見込確度以外の変更は転記対象にしない)
 * 案件を識別するキー列や見込確度列が「週次報告記録」に存在しない場合は、何もせずスキップする。
 * reportSheet・reportLastRow(追記前の最終行)は、履歴側の案件名セルに「週次報告記録」の
 * 該当行へ飛べるリンクを張るために使う(リンク自体の見た目はbuildStatusHistoryChange_を参照)。
 */
function updateWeeklyStatusHistory_(reportSheet, reportLastRow, headerMap, previousRows, newRows, timestamp) {
  const keyCol = findHeaderColumnIndex_(headerMap, CONFIG.STATUS_HISTORY_KEY_COLUMN_CANDIDATES);
  const confidenceCol = findHeaderColumnIndex_(headerMap, CONFIG.STATUS_HISTORY_CONFIDENCE_COLUMN_CANDIDATES);

  if (!keyCol || !confidenceCol) {
    Logger.log(
      '週次ステータス変更履歴: 案件識別列(' + CONFIG.STATUS_HISTORY_KEY_COLUMN_CANDIDATES.join('/') +
        ')または見込確度列(' + CONFIG.STATUS_HISTORY_CONFIDENCE_COLUMN_CANDIDATES.join('/') +
        ')が「' + CONFIG.TARGET_SHEET_NAME + '」に見つからないため、変更履歴の記録をスキップします。'
    );
    return;
  }

  // 案件(キー列の値)ごとに「直前の状態」を引けるようにする。同じ案件が複数行あっても、
  // 後の行で上書きされるため常に最後に登場した行が残る。
  const previousByKey = {};
  previousRows.forEach(function (row) {
    const key = String(row[keyCol - 1]).trim();
    if (key === '') return;
    previousByKey[key] = row;
  });

  const reportSheetId = reportSheet.getSheetId();
  const changes = [];
  newRows.forEach(function (row, index) {
    const key = String(row[keyCol - 1]).trim();
    if (key === '') return;

    const previousRow = previousByKey[key];
    if (previousRow) {
      const oldConfidence = previousRow[confidenceCol - 1];
      const newConfidence = row[confidenceCol - 1];
      // 転記条件: 変更前見込確度が存在し(空欄の新規登録は除外)、かつ変更前後で
      // 値が異なる(実際に確度の昇降があった)案件のみを対象にする。
      const hasOldConfidence = normalizeHistoryValue_(oldConfidence) !== '';
      if (hasOldConfidence && !historyValuesEqual_(oldConfidence, newConfidence)) {
        // newRowsはreportLastRow+1行目から順に書き込まれているため、行番号を逆算できる。
        const reportRow = reportLastRow + 1 + index;
        const reportRowLink = '#gid=' + reportSheetId + '&range=A' + reportRow;
        const fields = {};
        Object.keys(CONFIG.STATUS_HISTORY_FIELD_SOURCE_COLUMNS).forEach(function (name) {
          fields[name] = readTrackedField_(headerMap, row, CONFIG.STATUS_HISTORY_FIELD_SOURCE_COLUMNS[name]);
        });
        changes.push(
          buildStatusHistoryChange_(fields, key, oldConfidence, newConfidence, timestamp, reportRowLink)
        );
      }
    }

    // 初登場の案件は比較対象がないため対象外。以降の比較のために直前状態を更新する。
    previousByKey[key] = row;
  });

  if (changes.length === 0) return;

  appendStatusHistoryRows_(changes);
}

/**
 * 1件の見込確度変更を、「週次ステータス変更履歴」シートの列名をキーとしたオブジェクトに組み立てる。
 * fieldsは呼び出し元(週次同期側はヘッダー名マッチング、手動編集側は固定列位置)で既に
 * 解決済みの値を{ '担当': ..., 'カテゴリ': ..., ... }の形で渡す(値の取得元を問わない)。
 * 昇降フラグ(↑/↓)はH列・I列(変更前後の見込確度)を比較するスプレッドシート側の数式で
 * 算出される列のため、ここでは値をセットしない(applyRankFlagFormula_で数式をコピーする)。
 * reportRowLinkは案件名セルに張るリンク(該当行への内部リンク)のURLで、
 * applyProjectNameLinks_でリッチテキストとして反映される(この関数自体はセル値を書かない)。
 * reasonText(省略可)は手動編集時の変更理由(onConfidenceCellEdited参照)で、そのまま
 * 「変更経緯」に書き込む(自動生成の説明文は付与しない。未入力/キャンセル時・バッチ取り込み時は空文字)。
 */
function buildStatusHistoryChange_(fields, projectName, oldConfidence, newConfidence, timestamp, reportRowLink, reasonText) {
  const change = {};
  change['記録日時'] = Utilities.formatDate(timestamp, CONFIG.TIME_ZONE, CONFIG.STATUS_HISTORY_TIMESTAMP_FORMAT);
  change['担当'] = fields['担当'] || '';
  change['案件名'] = projectName;
  change['__案件名リンクURL'] = reportRowLink;
  change['カテゴリ'] = fields['カテゴリ'] || '';
  change['次回アクション予定日時'] = fields['次回アクション予定日時'] || '';
  change['ネクストアクション'] = fields['ネクストアクション'] || '';
  change['変更前見込確度'] = oldConfidence;
  change['変更後見込確度'] = newConfidence;
  change['変更経緯'] = reasonText || '';
  change['最新見積格納日'] = fields['最新見積格納日'] || '';
  change['契約見込月'] = fields['契約見込月'] || '';
  change['税込想定売上'] = fields['税込想定売上'] || '';
  change['貢献利益率'] = fields['貢献利益率'] || '';
  return change;
}

/**
 * 「週次報告記録」側の候補列名リストから最初に見つかった列の値を読み取る。
 * どれも見つからない場合は空文字を返す。
 */
function readTrackedField_(headerMap, row, candidateNames) {
  const col = findHeaderColumnIndex_(headerMap, candidateNames || []);
  return col ? row[col - 1] : '';
}

/**
 * headerMapの中から、候補列名リストの先頭から見て最初に存在する列番号を返す。
 * 1つも見つからない場合はnullを返す。
 */
function findHeaderColumnIndex_(headerMap, candidateNames) {
  for (let i = 0; i < candidateNames.length; i++) {
    if (headerMap[candidateNames[i]]) return headerMap[candidateNames[i]];
  }
  return null;
}

/**
 * 履歴の変更判定用に2つのセル値を比較する。null/undefined/空文字は同一視し、
 * それ以外は文字列化してトリムした上で比較する(数値と数値文字列などの揺れを吸収)。
 */
function historyValuesEqual_(a, b) {
  return normalizeHistoryValue_(a) === normalizeHistoryValue_(b);
}

function normalizeHistoryValue_(value) {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return String(value.getTime());
  return String(value).trim();
}

/**
 * 履歴の表示用に値を整形する。空欄は「(空欄)」と表示する。
 */
function formatHistoryValue_(value) {
  const normalized = value === null || value === undefined ? '' : String(value).trim();
  return normalized === '' ? '(空欄)' : normalized;
}

/**
 * 見込確度が変化した案件を「週次ステータス変更履歴」シートの末尾に追記する。
 * 列はCONFIG.STATUS_HISTORY_COLUMNSの名前でマッチングするため、シート側の列順が
 * 変わっていても正しい列に入り、シートが空の場合はこの並びでヘッダーを新規作成する。
 */
function appendStatusHistoryRows_(changes) {
  const historySheet = getOrCreateStatusHistorySheet_();
  const historyColumns = CONFIG.STATUS_HISTORY_COLUMNS;
  const headerRowNumber = CONFIG.STATUS_HISTORY_HEADER_ROW || 1;
  const historyHeaderMap = ensureColumnsExist_(historySheet, historyColumns, historyColumns, headerRowNumber);
  const historyTotalCols = historySheet.getLastColumn();

  const historyRows = changes.map(function (change) {
    const outRow = new Array(historyTotalCols).fill('');
    historyColumns.forEach(function (colName) {
      const col = historyHeaderMap[colName];
      if (col && change[colName] !== undefined) {
        outRow[col - 1] = change[colName];
      }
    });
    return outRow;
  });

  // getLastRow()はシート全体(どの列でもよい)の最終行を見るため、他の列に離れた場所まで
  // 書式やゴミデータが残っていると、実際のデータより大幅に手前/先の行を誤って返すことがある。
  // また、このシートは1行目=タイトル・headerRowNumber行目=列見出しの構成のため、
  // 記録日時列(通常A列)のうちデータ領域(見出し行の次の行以降)だけを対象に、
  // 値が入っている最終行を明示的に探して、その次の行に追記する。
  const dateCol = historyHeaderMap[historyColumns[0]];
  const dataStartRow = headerRowNumber + 1;
  const historyLastRow = findLastRowWithValueInColumn_(historySheet, dateCol, dataStartRow);
  const historyAppendRange = historySheet.getRange(historyLastRow + 1, 1, historyRows.length, historyTotalCols);
  writeValuesIgnoringValidation_(historyAppendRange, historyRows);

  applyRankFlagFormula_(historySheet, historyHeaderMap, historyLastRow, historyRows.length);
  applyProjectNameLinks_(historySheet, historyHeaderMap, historyLastRow, changes);
}

/**
 * 指定した列(col)のうち、startRow行目以降(タイトル行・ヘッダー行より下のデータ領域)を
 * 上から走査し、値が入っている最終行を返す。1行も無ければ(startRow - 1)を返す
 * (=呼び出し側はその次の行、つまりstartRowから安全に追記できる)。
 * (Sheet.getLastRow()はシート全体のどこかに値/書式があれば影響を受けるため、
 * 特定の列・特定の開始行に絞って「実際のデータの最終行」を求めたい場合に使う)
 */
function findLastRowWithValueInColumn_(sheet, col, startRow) {
  const searchStartRow = startRow || 1;
  const maxRow = sheet.getLastRow();
  if (maxRow < searchStartRow) return searchStartRow - 1;

  const numRows = maxRow - searchStartRow + 1;
  const values = sheet.getRange(searchStartRow, col, numRows, 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    const value = values[i][0];
    if (value !== '' && value !== null) return searchStartRow + i;
  }
  return searchStartRow - 1;
}

/**
 * 案件名列に、既存シートの見た目(リンク青色 + 下線)に合わせたリッチテキストを設定する。
 * リンク先はchange['__案件名リンクURL']に組み立て済みの「週次報告記録」該当行への
 * 内部リンクで、リンクが無い場合(reportSheetの行が特定できなかった場合)でも
 * 色・下線のスタイルだけは既存の見た目に合わせて適用する。
 */
function applyProjectNameLinks_(historySheet, headerMap, templateRow, changes) {
  const nameCol = headerMap['案件名'];
  if (!nameCol) return;

  const richTextValues = changes.map(function (change) {
    return [buildProjectNameRichText_(change['案件名'], change['__案件名リンクURL'])];
  });

  const destinationRange = historySheet.getRange(templateRow + 1, nameCol, changes.length, 1);
  destinationRange.setRichTextValues(richTextValues);
}

/**
 * 案件名セル用のリッチテキストを組み立てる。既存シートの案件名リンクと同じ見た目
 * (フォントカラー: CONFIG.STATUS_HISTORY_LINK_COLOR、下線あり)にし、linkUrlが
 * 渡された場合はそのURLへのリンクとしてセットする。
 */
function buildProjectNameRichText_(projectName, linkUrl) {
  const style = SpreadsheetApp.newTextStyle()
    .setForegroundColor(CONFIG.STATUS_HISTORY_LINK_COLOR)
    .setUnderline(true)
    .build();
  const builder = SpreadsheetApp.newRichTextValue().setText(String(projectName)).setTextStyle(style);
  if (linkUrl) {
    builder.setLinkUrl(linkUrl);
  }
  return builder.build();
}

/**
 * 昇降フラグ列(H列の変更前見込確度・I列の変更後見込確度を比較する数式が入る列)に、
 * 既存データ行の数式をコピーして新規追記行へ適用する。
 * GAS側では↑/↓の値を直接書き込まず、既存の数式をそのまま踏襲する方針とする
 * (数式の中身はスプレッドシート側の既存実装に委ね、GASは複製するだけに留める。
 * appendStatusHistoryRows_側でもこの列には値をセットしないため、GASからの上書きは発生しない)。
 * コピー元は直前行(templateRow)固定ではなく、templateRowから見出し行の次の行まで
 * 遡って実際に数式が入っている最も近い行を探す(直前行だけがたまたま数式なし・空欄等の
 * 場合でも、確実に既存の数式を引き継げるようにするため)。数式が入った行が1件も
 * 見つからない場合のみ、数式の設定をスキップする。
 */
function applyRankFlagFormula_(historySheet, headerMap, templateRow, newRowCount) {
  const flagCol = headerMap[CONFIG.STATUS_HISTORY_RANK_FLAG_COLUMN];
  if (!flagCol) return;

  const headerRowNumber = CONFIG.STATUS_HISTORY_HEADER_ROW || 1;
  const dataStartRow = headerRowNumber + 1;
  if (templateRow < dataStartRow) {
    Logger.log(
      '週次ステータス変更履歴: 「' + CONFIG.STATUS_HISTORY_RANK_FLAG_COLUMN +
        '」列の数式コピー元となる既存データ行が無いため、数式の設定をスキップしました。' +
        '(いずれか1行に手動で数式を設定しておけば、以降の追記時に自動でコピーされます)'
    );
    return;
  }

  const sourceRow = findNearestFormulaRow_(historySheet, flagCol, templateRow, dataStartRow);
  if (!sourceRow) {
    Logger.log(
      '週次ステータス変更履歴: 「' + CONFIG.STATUS_HISTORY_RANK_FLAG_COLUMN +
        '」列に数式が設定されている既存データ行が見つからなかったため、数式の設定をスキップしました。' +
        '(いずれか1行に手動で数式を設定しておけば、以降の追記時に自動でコピーされます)'
    );
    return;
  }

  const sourceCell = historySheet.getRange(sourceRow, flagCol);
  const destinationRange = historySheet.getRange(templateRow + 1, flagCol, newRowCount, 1);
  sourceCell.copyTo(destinationRange, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
}

/**
 * 指定した列(col)を、fromRow行目からminRow行目まで下から上に遡って走査し、
 * 数式が入っている最初の行番号を返す。見つからなければnullを返す。
 */
function findNearestFormulaRow_(sheet, col, fromRow, minRow) {
  for (let row = fromRow; row >= minRow; row--) {
    if (sheet.getRange(row, col).getFormula()) return row;
  }
  return null;
}

/**
 * 「週次ステータス変更履歴」シートを取得する。存在しなければ新規作成する。
 */
function getOrCreateStatusHistorySheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.TARGET_SPREADSHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.STATUS_HISTORY_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.STATUS_HISTORY_SHEET_NAME);
  }
  return sheet;
}

/**
 * 「今月」等の予材シート(CONFIG.CONFIDENCE_EDIT_SHEET_NAMES)で見込確度列が手動編集された際に、
 * ポップアップ(Ui.prompt)で変更理由を入力させ、「週次ステータス変更履歴」へ記録する。
 *
 * 【重要】単純トリガーのonEdit(e)はダイアログ(Ui.prompt/Browser.inputBox)を表示できないため、
 * この関数は installConfidenceChangeTrigger() で作成するインストーラブルトリガーとして
 * 登録すること(GASエディタで直接「onEdit」という名前の関数を作っても、この制限により
 * ダイアログの部分は動作しない)。
 *
 * - 「今月」シートは列見出しが「週次報告記録」側と一致しない(または見出し行の構成が異なる)ため、
 *   ヘッダー名でのマッチングは行わず、CONFIG.CONFIDENCE_EDIT_COLUMN_LETTERSに設定した
 *   固定の列位置から直接値を抽出する(readConfidenceEditField_参照)。
 * - 対象は「見込確度」列(CONFIG.CONFIDENCE_EDIT_COLUMN_LETTERS['見込確度']、既定でO列)の
 *   単一セル編集のみ。複数セルへの一括貼り付けは、編集前の値(oldValue)を個別に取得できない
 *   ため対象外とする。
 * - 変更理由の入力がキャンセル/空欄でも、変更自体は理由なしで記録する(セルの値は
 *   ユーザーの編集内容のまま変更しない)。
 * - 記録に失敗しても例外は投げず、ログとアラートに留める(セルの編集自体は既に確定しているため)。
 */
function onConfidenceCellEdited(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (CONFIG.CONFIDENCE_EDIT_SHEET_NAMES.indexOf(sheet.getName()) === -1) return;

  // ペースト等の複数セル編集はoldValueを個別に取得できないため対象外。
  if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;

  const editedRow = e.range.getRow();
  if (editedRow === 1) return; // ヘッダー行自体の編集は対象外

  const confidenceCol = columnLetterToIndex_(CONFIG.CONFIDENCE_EDIT_COLUMN_LETTERS['見込確度']);
  if (e.range.getColumn() !== confidenceCol) return;

  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return;
  // 編集された行の現在値(編集確定後の状態)をまとめて取得し、以降はここから固定列位置で
  // 直接読み取る(ヘッダー名マッチングには依存しない)。
  const rowValues = sheet.getRange(editedRow, 1, 1, lastCol).getValues()[0];

  const oldConfidence = e.oldValue !== undefined ? e.oldValue : '';
  // 変更後の値はe.valueではなく、確定後の状態を反映するrowValues(上で取得済み)から読む方が
  // 確実(貼り付け以外の一部の編集種別ではe.valueが入らないことがあるため)。
  const newConfidence = readConfidenceEditField_(rowValues, '見込確度');
  // 変更前見込確度が空欄(=新規登録による初回設定)の場合は「確度の変更」ではないため対象外とする。
  if (normalizeHistoryValue_(oldConfidence) === '') return;
  if (historyValuesEqual_(oldConfidence, newConfidence)) return;

  try {
    const projectName = String(readConfidenceEditField_(rowValues, '案件名')).trim();
    const displayName = projectName || ('(' + sheet.getName() + ' ' + editedRow + '行目)');

    const reason = promptForChangeReason_(displayName, oldConfidence, newConfidence);
    // 案件名セル(「今月」シート側)に既にリッチテキストのリンクが設定されていれば、
    // そのリンク先URLをそのまま引き継ぐ。リンクが無いセルの場合のみ、従来通り
    // 「今月」シートの該当行への内部リンクにフォールバックする。
    const projectNameCol = columnLetterToIndex_(CONFIG.CONFIDENCE_EDIT_COLUMN_LETTERS['案件名']);
    const projectNameCell = sheet.getRange(editedRow, projectNameCol);
    const sourceLinkUrl = getCellLinkUrl_(projectNameCell);
    const reportRowLink = sourceLinkUrl || ('#gid=' + sheet.getSheetId() + '&range=A' + editedRow);

    const fields = {
      '担当': readConfidenceEditField_(rowValues, '担当'),
      'カテゴリ': readConfidenceEditField_(rowValues, 'カテゴリ'),
      '次回アクション予定日時': readConfidenceEditField_(rowValues, '次回アクション予定日時'),
      'ネクストアクション': readConfidenceEditField_(rowValues, 'ネクストアクション'),
      '最新見積格納日': readConfidenceEditField_(rowValues, '最新見積格納日'),
      '契約見込月': readConfidenceEditField_(rowValues, '契約見込月'),
      '税込想定売上': readConfidenceEditField_(rowValues, '税込想定売上'),
      '貢献利益率': readConfidenceEditField_(rowValues, '貢献利益率')
    };

    const change = buildStatusHistoryChange_(
      fields, displayName, oldConfidence, newConfidence, new Date(), reportRowLink, reason
    );
    appendStatusHistoryRows_([change]);
  } catch (err) {
    Logger.log(
      '週次ステータス変更履歴(手動編集分)の記録に失敗しました: ' +
        sheet.getName() + '!' + e.range.getA1Notation() + ' / ' + err
    );
    try {
      SpreadsheetApp.getUi().alert('見込確度の変更履歴の記録に失敗しました。管理者に連絡してください。\n' + err);
    } catch (uiErr) {
      // アラート表示自体に失敗しても、記録失敗の実害はログに残っているため無視する。
    }
  }
}

/**
 * 見込確度の変更理由をUi.promptで入力させる。キャンセル/未入力の場合、またはダイアログを
 * 表示できない場合は空文字を返す(変更自体の記録は妨げない)。
 */
function promptForChangeReason_(projectName, oldConfidence, newConfidence) {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
      '見込確度の変更理由を入力してください',
      '案件: ' + projectName + '\n見込確度: ' +
        formatHistoryValue_(oldConfidence) + ' → ' + formatHistoryValue_(newConfidence),
      ui.ButtonSet.OK_CANCEL
    );
    if (response.getSelectedButton() !== ui.Button.OK) return '';
    return response.getResponseText().trim();
  } catch (err) {
    Logger.log('変更理由の入力ダイアログを表示できませんでした(理由は空欄で記録します): ' + err);
    return '';
  }
}

/**
 * 「今月」シートの1行分の値(rowValues、A列起点の配列)から、CONFIG.CONFIDENCE_EDIT_COLUMN_LETTERS
 * に設定された固定列位置を使ってfieldNameに対応する値を読み取る。列位置の設定が無い場合や、
 * rowValuesの範囲外(シートの右端より右を指している)の場合は空文字を返す。
 */
function readConfidenceEditField_(rowValues, fieldName) {
  const letter = CONFIG.CONFIDENCE_EDIT_COLUMN_LETTERS[fieldName];
  if (!letter) return '';
  const col = columnLetterToIndex_(letter);
  return col <= rowValues.length ? rowValues[col - 1] : '';
}

/**
 * セルに設定されているリッチテキストのリンク先URLを取得する。セル全体が単一のリンクとして
 * 設定されている場合はそのURLを返し、リンクが無い場合(または取得できない場合)は空文字を返す。
 */
function getCellLinkUrl_(cell) {
  try {
    const richText = cell.getRichTextValue();
    if (!richText) return '';
    return richText.getLinkUrl() || '';
  } catch (err) {
    return '';
  }
}

/**
 * データ入力規則(プルダウン等のリスト検証)が設定された範囲でも、リスト外の値を含む
 * 元データの追記が「入力規則違反」で失敗しないようにする。
 * 対象範囲の入力規則を一時的にクリアしてから値を書き込み、直後に元の規則を復元する。
 * (Sheetsの入力規則は入力時にのみ検証され、既存セルへ規則を再設定しても遡って
 * 再検証はされないため、復元後もエラーにはならず、プルダウン自体は維持される)
 */
function writeValuesIgnoringValidation_(range, values) {
  const existingValidations = range.getDataValidations();
  range.clearDataValidations();
  try {
    range.setValues(values);
  } finally {
    range.setDataValidations(existingValidations);
  }
}

/**
 * 対象シートのヘッダー行(headerRowNumber省略時は1行目)に、渡された列名のうち
 * 未登録のものを追加する。ヘッダーが空の場合はdefaultHeader(省略時は「取込日時」「元ファイル名」)
 * から作成する。タイトル行がヘッダーより上にあるシート(例: 週次ステータス変更履歴)では
 * headerRowNumberで実際のヘッダー行番号を指定する。
 * 戻り値: { 列名: 列番号(1始まり) } のマップ
 */
function ensureColumnsExist_(targetSheet, headerNames, defaultHeader, headerRowNumber) {
  const row = headerRowNumber || 1;
  const lastCol = targetSheet.getLastColumn();
  const headerRow = lastCol > 0 ? targetSheet.getRange(row, 1, 1, lastCol).getValues()[0] : [];
  const trimmedHeaderRow = headerRow.map(function (h) { return String(h).trim(); });
  const hasAnyHeader = trimmedHeaderRow.some(function (h) { return h !== ''; });

  // 列名は実際の列位置を保ったまま保持する(空欄の列があっても詰めない)。
  // 詰めてしまうと、それより右側の列名と実際の列位置が全てズレてしまう。
  let columns = hasAnyHeader ? trimmedHeaderRow.slice() : (defaultHeader || ['取込日時', '元ファイル名']).slice();
  let changed = !hasAnyHeader;

  headerNames.forEach(function (name) {
    const trimmed = String(name).trim();
    if (trimmed === '') return;
    if (columns.indexOf(trimmed) === -1) {
      columns.push(trimmed);
      changed = true;
    }
  });

  if (changed) {
    targetSheet.getRange(row, 1, 1, columns.length).setValues([columns]);
  }

  const map = {};
  columns.forEach(function (name, idx) {
    if (name !== '') map[name] = idx + 1;
  });
  return map;
}

/**
 * DriveAppとGAS標準サービスだけで週次報告ファイルを読み取り、
 * SpreadsheetApp.getDataRange().getValues()相当の2次元配列を返す。
 * Advanced Drive Service・UrlFetchApp・外部API呼び出しは一切使わない。
 *
 * ANDPAD等の出力は「拡張子は.xlsxだが中身は標準ZIP構造ではない」「実体はCSV/TSVや
 * HTMLテーブル」等のケースがあるため、ファイル先頭のマジックバイトや中身のテキストで
 * 実体を判別し、以下の順にフォールバックする。途中の形式でのパースに失敗しても即エラー
 * にはせず、ログを残して次の形式を試す。
 *   1. 本物の.xlsx(ZIP構造): Utilities.unzip + XmlServiceで直接パース。
 *   2. 旧形式バイナリ.xls(OLE構造): パース不可のため、その旨を明確なエラーで通知する。
 *   3. HTMLテーブルを.xlsx/.xlsとして出力している場合: <table>をHTMLとして解析する。
 *   4. CSV/TSVをテキストとして出力している場合: 区切り文字(カンマ/タブ)を自動判定し、
 *      ダブルクォート囲み・エスケープにも対応した簡易CSVパーサで解析する。
 * どの形式にも当てはまらない場合は、原因調査用に検出したContentTypeと先頭バイトを
 * 含むエラーを投げる。
 */
function parseWeeklyReportFile_(file) {
  const blob = file.getBlob();
  const bytes = blob.getBytes();

  if (looksLikeZip_(bytes)) {
    try {
      return parseXlsxZipDirectly_(blob);
    } catch (zipErr) {
      Logger.log(
        'parseXlsxZipDirectly_ 失敗、HTMLフォールバックを試みます: ' + file.getName() + ' (' + zipErr + ')'
      );
    }
  }

  if (looksLikeLegacyOle_(bytes)) {
    throw new Error(
      '旧形式のバイナリExcelファイル(.xls, Excel 97-2003形式)は読み取れません。' +
        '.xlsxまたはCSV形式で保存し直してください。(ファイル: ' + file.getName() + ')'
    );
  }

  const text = decodeTextBlob_(blob);

  if (/<html[\s>]|<table[\s>]/i.test(text.slice(0, 4000))) {
    try {
      return parseHtmlTableBlob_(text);
    } catch (htmlErr) {
      Logger.log(
        'parseHtmlTableBlob_ 失敗、CSVフォールバックを試みます: ' + file.getName() + ' (' + htmlErr + ')'
      );
    }
  }

  if (looksLikeCsvText_(text)) {
    return parseCsvBlob_(text);
  }

  throw new Error(
    'サポートされていないファイル形式です(ZIP/.xlsx・HTML・CSVのいずれでもありません)。' +
      'ファイル: ' + file.getName() +
      ' / ContentType: ' + blob.getContentType() +
      ' / 先頭バイト: ' + bytesToHexPreview_(bytes)
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
 * テキストの先頭数行にカンマまたはタブが含まれるかどうかで、CSV/TSVらしさを判定する。
 * (HTML判定・ZIP判定・OLE判定のいずれにも当てはまらなかった場合の最終フォールバックとして使う)
 */
function looksLikeCsvText_(text) {
  const lines = text.split(/\r\n|\r|\n/).slice(0, 5).filter(function (line) {
    return line.trim() !== '';
  });
  if (lines.length === 0) return false;
  return lines.some(function (line) {
    return line.indexOf(',') !== -1 || line.indexOf('\t') !== -1;
  });
}

/**
 * バイト列先頭16バイトを16進数プレビュー文字列にする(未対応形式のエラー診断用)。
 */
function bytesToHexPreview_(bytes) {
  return bytes.slice(0, 16).map(function (b) {
    const unsigned = b < 0 ? b + 256 : b;
    const hex = unsigned.toString(16).toUpperCase();
    return hex.length === 1 ? '0' + hex : hex;
  }).join(' ');
}

/**
 * blobのテキストをUTF-8で読み、内部にShift_JIS系のcharset宣言が見つかればShift_JISで読み直す。
 * (ANDPAD等、日本語業務システムのHTML/CSV書き出しはShift_JISであることが多いため)
 * 先頭のUTF-8 BOMが付いている場合は取り除く。
 */
function decodeTextBlob_(blob) {
  const utf8Text = blob.getDataAsString('UTF-8');
  const metaMatch = utf8Text.slice(0, 2000).match(/charset\s*=\s*["']?([\w-]+)/i);
  const charset = metaMatch ? metaMatch[1].toLowerCase() : '';
  const text = /shift[-_]?jis|sjis|ms932|windows-31j/.test(charset)
    ? blob.getDataAsString('Shift_JIS')
    : utf8Text;
  return text.replace(/^\uFEFF/, '');
}

/**
 * 本物の.xlsx(ZIP構造)をUtilities.unzip + XmlServiceで直接パースする。
 */
function parseXlsxZipDirectly_(blob) {
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
 * <tr>を1件も抽出できなかった場合はエラーを投げ、呼び出し元でCSVフォールバックに
 * つなげられるようにする。
 */
function parseHtmlTableBlob_(html) {
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
  if (rowMatches.length === 0) {
    throw new Error('HTMLとして解析できる<table>/<tr>が見つかりませんでした。');
  }

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
 * CSV/TSVテキストを2次元配列にパースする。区切り文字は先頭行のカンマ/タブの出現数で
 * 自動判定する。ダブルクォート囲み・囲み内のカンマ/改行・""によるエスケープに対応した
 * 簡易パーサ(RFC4180相当)。
 */
function parseCsvBlob_(text) {
  const delimiter = detectCsvDelimiter_(text);
  return parseDelimitedText_(text, delimiter);
}

/**
 * テキスト先頭行のカンマ数とタブ数を比較し、区切り文字を推定する。
 */
function detectCsvDelimiter_(text) {
  const firstLine = text.split(/\r\n|\r|\n/)[0] || '';
  const commaCount = (firstLine.match(/,/g) || []).length;
  const tabCount = (firstLine.match(/\t/g) || []).length;
  return tabCount > commaCount ? '\t' : ',';
}

/**
 * 指定した区切り文字でテキストを2次元配列にパースする(ダブルクォート対応)。
 * 各行の列数は最大列数に合わせて空文字でパディングする。
 */
function parseDelimitedText_(text, delimiter) {
  const rows = [];
  let row = [];
  let field = '';
  let inQuotes = false;
  let i = 0;
  const len = text.length;

  while (i < len) {
    const ch = text.charAt(i);

    if (inQuotes) {
      if (ch === '"') {
        if (text.charAt(i + 1) === '"') {
          field += '"';
          i += 2;
        } else {
          inQuotes = false;
          i++;
        }
      } else {
        field += ch;
        i++;
      }
      continue;
    }

    if (ch === '"') {
      inQuotes = true;
      i++;
    } else if (ch === delimiter) {
      row.push(field);
      field = '';
      i++;
    } else if (ch === '\r' || ch === '\n') {
      if (ch === '\r' && text.charAt(i + 1) === '\n') i++;
      row.push(field);
      field = '';
      rows.push(row);
      row = [];
      i++;
    } else {
      field += ch;
      i++;
    }
  }

  if (field !== '' || row.length > 0) {
    row.push(field);
    rows.push(row);
  }

  while (rows.length > 0 && rows[rows.length - 1].every(function (cell) { return cell === ''; })) {
    rows.pop();
  }

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
    mimeType === 'application/vnd.ms-excel' ||
    mimeType === MimeType.CSV ||
    mimeType === 'text/csv'
  ) {
    return true;
  }
  const name = file.getName().toLowerCase();
  return name.endsWith('.xlsx') || name.endsWith('.xls') || name.endsWith('.csv') || name.endsWith('.tsv');
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

// ===== トリガー設定 =====

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

/**
 * 処理済みファイルIDの記録をリセットする。
 * テストのため同じファイルをもう一度処理させたい場合などに、この関数だけを手動実行する。
 */
function resetProcessedLog() {
  PropertiesService.getScriptProperties().deleteProperty('PROCESSED_FILE_IDS');
  Logger.log('処理済みログをリセットしました。');
}
