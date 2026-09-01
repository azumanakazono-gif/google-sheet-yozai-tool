/**
 * 「週次ステータス変更履歴」シートの自動更新処理。
 * - バッチ取り込み時: updateWeeklyStatusHistory_ (appendFileToTargetSheet_ から呼ばれる)
 * - 「今月」等の予材シートでの手動編集時: onConfidenceCellEdited (編集トリガーから呼ばれる)
 * 依存: Config.gs (CONFIG) / Utils.gs (ensureColumnsExist_, writeValuesIgnoringValidation_)
 */

/**
 * 「報告記録」への直近の追記分(newRows)を、追記前の既存データ(previousRows)と
 * 案件単位で突き合わせ、「見込確度」(CONFIG.STATUS_HISTORY_CONFIDENCE_COLUMN_CANDIDATES)が
 * 変更前後で変化した案件だけを抽出して「週次ステータス変更履歴」シートの末尾に記録する。
 * (ステータスや進捗など見込確度以外の変更は転記対象にしない)
 * 案件を識別するキー列や見込確度列が「報告記録」に存在しない場合は、何もせずスキップする。
 * reportSheet・reportLastRow(追記前の最終行)は、履歴側の案件名セルに「報告記録」の
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
  change['備考（状況）'] = fields['備考（状況）'] || '';
  return change;
}

/**
 * 「報告記録」側の候補列名リストから最初に見つかった列の値を読み取る。
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
 * 戻り値: 実際に書き込んだ最初の行番号(1件のみ追記した場合は、その行番号そのもの)。
 * onConfidenceCellEdited側で、後から変更経緯セルだけを追記更新するために使う。
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

  // 昇降フラグの数式コピー・案件名リンクの設定は見た目を整えるための付随処理。
  // ここで例外が発生しても、直前に書き込んだ履歴データ自体(記録日時・見込確度の変化等)を
  // 無かったことにしてはならないため、失敗してもロールバックや呼び出し元への例外伝播は
  // せず、ログにのみ残す(そうしないと「本体は記録できたのに全体が失敗扱いになり、
  // 履歴が更新されていないように見える」事態になる)。
  try {
    applyRankFlagFormula_(historySheet, historyHeaderMap, historyLastRow, historyRows.length);
  } catch (err) {
    Logger.log('週次ステータス変更履歴: 昇降フラグの数式コピーに失敗しました(記録データ自体は書き込み済み): ' + err);
  }
  try {
    applyProjectNameLinks_(historySheet, historyHeaderMap, historyLastRow, changes);
  } catch (err) {
    Logger.log('週次ステータス変更履歴: 案件名リンクの設定に失敗しました(記録データ自体は書き込み済み): ' + err);
  }

  return historyLastRow + 1;
}

/**
 * 履歴シートの指定行の「変更経緯」セルだけを追記更新する。
 * onConfidenceCellEdited側で、まず理由なしで記録を確定させた後、変更理由の入力
 * ダイアログ(Ui.prompt、失敗しうる)が成功した場合にのみベストエフォートで呼び出す。
 */
function writeChangeReason_(row, reason) {
  const historySheet = getOrCreateStatusHistorySheet_();
  const headerRowNumber = CONFIG.STATUS_HISTORY_HEADER_ROW || 1;
  const historyHeaderMap = ensureColumnsExist_(
    historySheet, CONFIG.STATUS_HISTORY_COLUMNS, CONFIG.STATUS_HISTORY_COLUMNS, headerRowNumber
  );
  const col = historyHeaderMap['変更経緯'];
  if (!col) return;
  historySheet.getRange(row, col).setValue(reason);
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
 * リンク先はchange['__案件名リンクURL']に組み立て済みの「報告記録」該当行への
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
 * - 「今月」シートは列見出しが「報告記録」側と一致しない(または見出し行の構成が異なる)ため、
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
  // 【重要】以前はガード判定(シート名・列位置チェック等)をtryの外に書いていたため、
  // その部分で例外が起きると(例: CONFIG.CONFIDENCE_EDIT_COLUMN_LETTERSの設定ミスや
  // シート構成の変更によるgetRange失敗など)ログにすら残らず、ポップアップも履歴記録も
  // 一切発生しないまま処理が終わっていた。関数全体をtryで包み、原因を問わず失敗時は
  // 必ずログ・トーストのいずれかで気づけるようにする。
  try {
    if (!e || !e.range) {
      Logger.log('onConfidenceCellEdited: イベントオブジェクトにrangeが無いため終了します。');
      return;
    }

    const sheet = e.range.getSheet();
    if (CONFIG.CONFIDENCE_EDIT_SHEET_NAMES.indexOf(sheet.getName()) === -1) return;

    // ペースト等の複数セル編集はoldValueを個別に取得できないため対象外。
    if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) {
      Logger.log(
        'onConfidenceCellEdited: 複数セル編集(' + sheet.getName() + '!' + e.range.getA1Notation() +
          ')のため対象外とします。'
      );
      return;
    }

    const editedRow = e.range.getRow();
    if (editedRow === 1) return; // ヘッダー行自体の編集は対象外

    const confidenceCol = columnLetterToIndex_(CONFIG.CONFIDENCE_EDIT_COLUMN_LETTERS['見込確度']);
    if (e.range.getColumn() !== confidenceCol) return; // 見込確度以外の列編集は対象外(頻出のためログ省略)

    const lastCol = sheet.getLastColumn();
    if (lastCol === 0) {
      Logger.log('onConfidenceCellEdited: 「' + sheet.getName() + '」にデータが無いため終了します。');
      return;
    }
    // 編集された行の現在値(編集確定後の状態)をまとめて取得し、以降はここから固定列位置で
    // 直接読み取る(ヘッダー名マッチングには依存しない)。
    const rowValues = sheet.getRange(editedRow, 1, 1, lastCol).getValues()[0];

    const oldConfidence = e.oldValue !== undefined ? e.oldValue : '';
    // 変更後の値はe.valueではなく、確定後の状態を反映するrowValues(上で取得済み)から読む方が
    // 確実(貼り付け以外の一部の編集種別ではe.valueが入らないことがあるため)。
    const newConfidence = readConfidenceEditField_(rowValues, '見込確度');
    // 変更前見込確度が空欄(=新規登録による初回設定)の場合は「確度の変更」ではないため対象外とする。
    if (normalizeHistoryValue_(oldConfidence) === '') {
      Logger.log(
        'onConfidenceCellEdited: ' + sheet.getName() + '!' + e.range.getA1Notation() +
          ' は変更前見込確度が空欄(新規登録扱い)のため対象外とします。(新しい値: ' + newConfidence + ')' +
          (e.oldValue === undefined
            ? ' ※e.oldValueが未定義でした(GASの仕様上、編集前セルが空欄だった場合等に発生します)。'
            : '')
      );
      return;
    }
    if (historyValuesEqual_(oldConfidence, newConfidence)) {
      Logger.log(
        'onConfidenceCellEdited: ' + sheet.getName() + '!' + e.range.getA1Notation() +
          ' は値に実質的な変化が無いため対象外とします。(旧: ' + oldConfidence + ' / 新: ' + newConfidence + ')'
      );
      return;
    }

    Logger.log(
      'onConfidenceCellEdited: 見込確度の変更を検知しました。' + sheet.getName() + '!' + e.range.getA1Notation() +
        ' (旧: ' + oldConfidence + ' → 新: ' + newConfidence + ')'
    );

    const projectName = String(readConfidenceEditField_(rowValues, '案件名')).trim();
    const displayName = projectName || ('(' + sheet.getName() + ' ' + editedRow + '行目)');

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
      '貢献利益率': readConfidenceEditField_(rowValues, '貢献利益率'),
      '備考（状況）': readConfidenceEditField_(rowValues, '備考（状況）')
    };

    // 【重要】変更理由の入力ダイアログ(Ui.prompt)は、このスクリプトプロジェクトが
    // 対象スプレッドシートのコンテナバインド型スクリプトとして設置されていない場合
    // (スタンドアロンのスクリプトプロジェクトの場合)、SpreadsheetApp.getUi()自体が
    // 例外を投げ、常に失敗する。ダイアログの成否に関わらず「見込確度が変わった」という
    // 事実の記録は必ず残すべきなので、先に理由なしで記録を確定させ、ダイアログの表示・
    // 入力に成功した場合だけ、後から変更経緯セルを追記更新する(記録自体を
    // ダイアログの可否に依存させない)。
    const change = buildStatusHistoryChange_(
      fields, displayName, oldConfidence, newConfidence, new Date(), reportRowLink, ''
    );
    const appendedRow = appendStatusHistoryRows_([change]);
    Logger.log(
      'onConfidenceCellEdited: 「' + CONFIG.STATUS_HISTORY_SHEET_NAME + '」の' + appendedRow + '行目に記録しました。'
    );

    const reason = promptForChangeReason_(displayName, oldConfidence, newConfidence);
    if (reason) {
      writeChangeReason_(appendedRow, reason);
    }
  } catch (err) {
    Logger.log(
      '週次ステータス変更履歴(手動編集分)の記録に失敗しました: ' +
        (e && e.range ? e.range.getSheet().getName() + '!' + e.range.getA1Notation() : '(不明なセル)') +
        ' / ' + err
    );
    notifyConfidenceEditError_(err, e);
  }
}

/**
 * onConfidenceCellEdited内の例外を、Logger.logだけでなく画面上にも必ず気づける形で通知する。
 * SpreadsheetApp.getActiveSpreadsheet()はスタンドアロンのスクリプトプロジェクトでは
 * 使えない場合があるため、まずイベントオブジェクトから直接取得したSpreadsheet
 * (e.range.getSheet().getParent())でtoastを試み、それも使えない場合のみ
 * getActiveSpreadsheet()にフォールバックする。どちらも失敗する場合はログのみに留める。
 */
function notifyConfidenceEditError_(err, e) {
  const message = String(err);
  const title = '見込確度の変更履歴の記録でエラーが発生しました';
  try {
    const ss = (e && e.range) ? e.range.getSheet().getParent() : SpreadsheetApp.getActiveSpreadsheet();
    ss.toast(message, title, 10);
    return;
  } catch (toastErr) {
    Logger.log('エラー通知(toast)に失敗しました(e由来のSpreadsheet): ' + toastErr);
  }
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title, 10);
  } catch (toastErr2) {
    Logger.log(
      'エラー通知(toast)にも失敗しました(このスクリプトがスプレッドシートに' +
        'バインドされていない可能性があります): ' + toastErr2
    );
  }
}

/**
 * 見込確度の変更理由をUi.promptで入力させる。キャンセル/未入力の場合、またはダイアログを
 * 表示できない場合は空文字を返す(変更自体の記録は妨げない)。
 */
function promptForChangeReason_(projectName, oldConfidence, newConfidence) {
  let ui;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (err) {
    // 【重要】SpreadsheetApp.getUi()は、このスクリプトプロジェクトが対象スプレッドシートの
    // コンテナバインド型スクリプト(そのスプレッドシートの「拡張機能」>「Apps Script」から
    // 開けるプロジェクト)として設置されていない場合、常にこの時点で例外を投げる。
    // ポップアップが一度も表示されない場合は、まずこのログが出ていないか確認すること。
    Logger.log(
      '変更理由の入力ダイアログを表示できませんでした(このスクリプトが対象スプレッドシートの' +
        'コンテナバインド型スクリプトとして設置されていない可能性があります。記録自体は理由なしで' +
        '完了しています): ' + err
    );
    return '';
  }
  try {
    const response = ui.prompt(
      '見込確度の変更理由を入力してください',
      '案件: ' + projectName + '\n見込確度: ' +
        formatHistoryValue_(oldConfidence) + ' → ' + formatHistoryValue_(newConfidence),
      ui.ButtonSet.OK_CANCEL
    );
    if (response.getSelectedButton() !== ui.Button.OK) return '';
    return response.getResponseText().trim();
  } catch (err) {
    Logger.log('変更理由の入力ダイアログの表示中にエラーが発生しました(理由は空欄で記録します): ' + err);
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
