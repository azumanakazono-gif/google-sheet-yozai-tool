/**
 * 予材リスト週次同期スクリプト（単一ファイル版）
 *
 * - SOURCE_FOLDER_ID 配下のANDPAD出力Excelファイル（Googleスプレッドシートも可）を読み取り、
 *   「31期予材リスト」の「報告記録」シート末尾に追記する。
 * - Excelファイルは DriveApp + Utilities.unzip + XmlService のみで直接パースする（Advanced Drive
 *   Service、UrlFetchApp、外部API呼び出しは一切使わない）。ANDPAD等の「拡張子は.xlsxだが中身は
 *   標準ZIPではない(HTMLテーブル等)」出力にも対応するため、先頭バイトで実体を判別し、
 *   ZIP/.xlsx・HTMLテーブルの2通りの読み取り方法へ自動でフォールバックする。
 * - 列はヘッダー名でマッチングするため、ANDPAD側の列順変更や列追加に強い。
 * - 転記済みファイルは「処理済み」フォルダへ移動し、同名ファイルがあればリネームして衝突を回避する。
 * - 万一の移動失敗に備え、処理済みファイルIDをスクリプトプロパティに記録し、再走時の二重追記を防ぐ。
 * - 「対象期間」の区切り行は挿入せず、読み取ったデータをそのまま「報告記録」シート末尾に
 *   追記する(ソートや区切り行の再構築は行わない)。
 * - 「今月」を含むシート名の5行目以降・O列(見込確度)の編集を検知し、「報告記録」と
 *   「週次ステータス変更履歴」へ自動転記する(onEdit)。この機能は、このスクリプトが
 *   「31期予材リスト」のコンテナバインド型スクリプトとして設置されている場合のみ動作する。
 * - 「今月」タブのH列(案件名テキスト+ハイパーリンク)を、「週次ステータス変更履歴」のC列
 *   (案件名)へ自動でコピーする。既存行に一括適用したい場合は applyProjectLinksToStatusHistory()
 *   を手動実行する。
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
  TARGET_SHEET_NAME: '報告記録',

  // 取り込み元ファイルの1行目がヘッダー行かどうか
  SOURCE_HAS_HEADER: true,

  // エラー発生時に通知したいメールアドレス。空文字なら通知しない。
  NOTIFY_EMAIL: '',

  TIME_ZONE: 'Asia/Tokyo',

  // 処理済みファイルIDを記録しておく数（moveTo失敗時などの二重処理を防ぐ保険）
  PROCESSED_LOG_MAX: 300,

  // CVR集計用の正規列名。ANDPAD側の表記ゆれ(別名)を吸収し、必ずこの列名に正規化して書き込む。
  // key: 報告記録シート上の正規列名 / value: ANDPAD側で使われうる別名(表記ゆれ)の配列。
  // ここに無い列名は元の名前のまま追記される(自動追加)ので、ANDPAD側の実際の項目名を
  // 確認のうえ、必要な別名を随時この配列に追加すること。
  // CVR集計はANDPAD報告画面の4ステータス「アプローチ」「面談」「提案（見積提示）」「契約」を
  // 基準とする。各ステータスの到達日を以下の正規列名で保持し、
  //   ①アプローチ→面談CVR = 面談数 ÷ アプローチ数
  //   ②面談→提案（見積提示）CVR = 提案数 ÷ 面談数
  //   ③提案（見積提示）→契約CVR = 契約数 ÷ 提案数
  // を算出する(詳細はCVR_KPI_DESIGN.md参照)。「提案（見積提示）」はANDPAD側の表記が
  // 全角/半角カッコや「見積日」「提案日」など揺れうるため、別名を広めに登録している。
  STAGE_COLUMN_ALIASES: {
    '案件種別': ['案件種別', '案件区分', '種別'],
    '属性': ['属性', '顧客属性', '反響属性'],
    'アプローチ日': ['アプローチ日', '初回アプローチ日', '初回接触日', '反響日', 'アプローチ'],
    '面談日': ['面談日', '面談実施日', '打合せ日', '初回面談日', '面談'],
    '提案日': [
      '提案日', '提案（見積提示）', '提案(見積提示)',
      '提案（見積提示）日', '提案(見積提示)日',
      '見積日', '見積提出日', '御見積日', '見積書提出日', '提案書提出日'
    ],
    '契約日': ['契約日', '成約日', '受注日', '契約']
  },

  // ===== 案件種別の自動判定(カテゴリ分類) =====
  // 「案件（引合）名」のテキストに含まれるキーワードから「案件種別」列の値を自動判定する
  // (classifyCategory_)。データ入力規則で定義されている15カテゴリ(「風力」は対象外)を
  // 優先順位の高い順に並べている。上から順にキーワードを走査し、最初に一致したカテゴリを
  // 採用する。どれにも一致しない場合はCATEGORY_FALLBACKを使う。
  // 実際のANDPAD案件名の表記に合わせて、随時keywordsを追加・調整すること。
  CATEGORY_RULES: [
    { category: '事業者用PV', keywords: ['事業者用PV', '事業用PV', '産業用PV'] },
    { category: 'リパワリング', keywords: ['リパワリング', 'リパワー'] },
    { category: '事業者PCS交換', keywords: ['事業者PCS交換', '事業用PCS交換', '産業用PCS交換'] },
    { category: 'リプレイス', keywords: ['リプレイス'] },
    { category: '事業者用創蓄', keywords: ['事業者用創蓄', '事業用創蓄', '産業用創蓄', '事業者用蓄電池'] },
    { category: '住宅用PV', keywords: ['住宅用PV'] },
    { category: '住宅用創蓄', keywords: ['住宅用創蓄', '住宅用蓄電池'] },
    { category: 'EV充放電システム', keywords: ['EV充放電システム', 'EV充放電'] },
    { category: '戸建てPCS交換', keywords: ['戸建てPCS交換', '戸建PCS交換'] },
    { category: 'EQ/オール電化', keywords: ['EQ/オール電化', 'EQ／オール電化', 'オール電化', 'EQ'] },
    { category: 'O&M', keywords: ['O&M', 'O＆M', 'Ｏ&Ｍ', 'Ｏ＆Ｍ'] },
    { category: 'メンテナンス', keywords: ['メンテナンス', 'メンテ'] },
    { category: "LED'S", keywords: ["LED'S", 'LED’S', 'LEDS'] },
    { category: '行政案件', keywords: ['行政案件', '行政'] }
    // 「その他」は上記いずれにも一致しなかった場合のフォールバック(CATEGORY_FALLBACK)として扱うため、
    // ルール一覧には含めない。
  ],
  // 上記いずれのキーワードにも一致しなかった場合に「案件種別」へ設定する値。
  CATEGORY_FALLBACK: 'その他',
  // 「案件種別」を自動判定する元になる、案件名が入っている列のヘッダー名候補。
  // ANDPAD側の表記ゆれ(「案件名」「引合名」など)を吸収するため、上から順に一致する列を探す。
  PROJECT_NAME_COLUMN_CANDIDATES: ['案件（引合）名', '案件(引合)名', '案件名', '引合名', '案件名称'],

  // ===== 転記対象の最大列数 =====
  // 「報告記録」シートへ転記する列はこの列数まで（既定: 30列目 = AD列）。
  // それ以降の列は不要なデータのため、ヘッダー追加・値の書き込みともに対象外とする。
  // ANDPAD側の列(元24列)に加え、STAGE_COLUMN_ALIASESのCVR集計用6列(案件種別/属性/
  // アプローチ日/面談日/提案日/契約日)が収まるよう6列分の余裕を持たせている。
  MAX_DATA_COLUMNS: 30,

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
  STATUS_HISTORY_SHEET_NAME: '週次ステータス変更履歴',

  // ===== 「週次ステータス変更履歴」の案件名ハイパーリンク =====
  // 「今月」タブのこの列(既定: H列)には、案件名テキストにURLへのハイパーリンクが
  // 設定されている。同じテキスト+リンクを「週次ステータス変更履歴」の
  // STATUS_HISTORY_PROJECT_COLUMNへ自動でコピーする。
  EDIT_URL_COLUMN: 8,
  // 「週次ステータス変更履歴」側で案件名+ハイパーリンクを表示する列（既定: 3列目 = C列）
  STATUS_HISTORY_PROJECT_COLUMN: 3
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
 * 「対象期間」の区切り行は挿入せず、読み取った行をそのままシート末尾へ追記するだけである。
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
    // 空のまま「処理済み」として静かに移動してしまうと、パースの取りこぼしに気づけない
    // (ファイル自体は失われる)ため、エラーとして扱いエラーフォルダへ移動・通知の対象にする。
    throw new Error(
      'ファイルからデータを1行も読み取れませんでした。ANDPAD側の出力形式が変わっていないか確認してください。'
    );
  }

  const hasHeader = CONFIG.SOURCE_HAS_HEADER;
  const rawHeader = hasHeader ? values[0] : values[0].map(function (_, idx) { return '列' + (idx + 1); });
  const sourceHeader = rawHeader.map(function (h) { return String(h).trim(); });
  const dataRows = (hasHeader ? values.slice(1) : values).filter(function (row) {
    return row.some(function (cell) { return cell !== '' && cell !== null; });
  });

  if (dataRows.length === 0) {
    // 上と同様、ヘッダー行しか読み取れなかった場合も静かにスキップせずエラーとして扱う。
    throw new Error(
      'ヘッダー行を除くとデータ行が1件もありませんでした。ANDPAD側の出力形式が変わっていないか確認してください。'
    );
  }

  // CVR集計列(案件種別/属性/アプローチ日/面談日/提案日/契約日)は表記ゆれを吸収して正規化し、
  // かつ今回のファイルに列が無くてもシート上には必ず存在させる(CVR集計式の列位置を安定させるため)。
  const canonicalHeader = sourceHeader.map(resolveCanonicalHeader_);
  const stageColumnNames = Object.keys(CONFIG.STAGE_COLUMN_ALIASES || {});
  const headerMap = ensureColumnsExist_(targetSheet, canonicalHeader.concat(stageColumnNames));
  const totalCols = Math.min(targetSheet.getLastColumn(), CONFIG.MAX_DATA_COLUMNS);
  const timestamp = new Date();

  // 「案件種別」はANDPAD側の値をそのまま転記せず、案件(引合)名のテキストから
  // classifyCategory_で自動判定して上書きする(カテゴリ別CVR集計を固定15分類で行うため)。
  const projectNameColIdx = sourceHeader.findIndex(function (name) {
    return (CONFIG.PROJECT_NAME_COLUMN_CANDIDATES || []).indexOf(name) !== -1;
  });

  const rowsToAppend = dataRows.map(function (row) {
    const outRow = new Array(totalCols).fill('');
    outRow[headerMap['取込日時'] - 1] = timestamp;
    outRow[headerMap['元ファイル名'] - 1] = file.getName();
    canonicalHeader.forEach(function (colName, idx) {
      if (!colName) return;
      const col = headerMap[colName];
      if (col && col <= totalCols) outRow[col - 1] = row[idx];
    });
    if (projectNameColIdx !== -1 && headerMap['案件種別'] && headerMap['案件種別'] <= totalCols) {
      outRow[headerMap['案件種別'] - 1] = classifyCategory_(row[projectNameColIdx]);
    }
    return outRow;
  });

  const lastRow = targetSheet.getLastRow();
  targetSheet.getRange(lastRow + 1, 1, rowsToAppend.length, totalCols).setValues(rowsToAppend);
}

/**
 * 対象シートのヘッダー行(1行目)に、渡された列名のうち未登録のものを追加する。
 * ヘッダーが空の場合は「取込日時」「元ファイル名」から作成する。
 * MAX_DATA_COLUMNS(既定: 24列目 = X列)に達している場合、それ以降の新規列は追加しない
 * (Y列以降は転記対象外のため)。
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
    if (existing.length >= CONFIG.MAX_DATA_COLUMNS) return;
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

/**
 * ANDPAD側の列名がCONFIG.STAGE_COLUMN_ALIASESに登録された別名(表記ゆれ)と一致する場合、
 * CVR集計用の正規列名(案件種別/属性/アプローチ日/面談日/提案日/契約日)に変換する。
 * 一致しなければ元の列名をそのまま返す。
 */
function resolveCanonicalHeader_(name) {
  const aliases = CONFIG.STAGE_COLUMN_ALIASES || {};
  const trimmed = String(name).trim();
  const canonicalNames = Object.keys(aliases);
  for (let i = 0; i < canonicalNames.length; i++) {
    const canonical = canonicalNames[i];
    if (aliases[canonical].indexOf(trimmed) !== -1) return canonical;
  }
  return trimmed;
}

/**
 * 案件(引合)名のテキストからCONFIG.CATEGORY_RULESを先頭から順に走査し、
 * 最初にキーワードが一致したカテゴリ名を返す(案件種別の自動判定)。
 * どれにも一致しない場合はCONFIG.CATEGORY_FALLBACK(既定:「その他」)を返す。
 */
function classifyCategory_(projectName) {
  const text = String(projectName || '').toUpperCase();
  const rules = CONFIG.CATEGORY_RULES || [];
  for (let i = 0; i < rules.length; i++) {
    const keywords = rules[i].keywords || [];
    for (let j = 0; j < keywords.length; j++) {
      if (text.indexOf(String(keywords[j]).toUpperCase()) !== -1) return rules[i].category;
    }
  }
  return CONFIG.CATEGORY_FALLBACK || 'その他';
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
 * フォールバック。以下の手順で、実際にデータが入っているテーブルを取りこぼしなく抽出する。
 *   1. splitTopLevelTables_で最上位の<table>ブロックを入れ子を壊さずに分割する。
 *      単純な非貪欲正規表現(<table[\s\S]*?<\/table>)は、テーブルが入れ子になっている場合
 *      (レイアウト用の外側<table>の中に実データの<table>がある等)、内側の</table>で
 *      打ち切られてしまい、実データを含む外側テーブルを正しく取得できないことがある。
 *   2. 各ブロックについて実際に行・セルを抽出し、<tr>タグの出現数ではなく「値が入っている
 *      行数」で比較して最もデータらしいブロックを選ぶ。タグの出現数だけで比較すると、
 *      見た目のレイアウト目的の(値の少ない)テーブルの方が<tr>数が多く、誤って選ばれることがある。
 * XmlServiceでの厳密なXMLパースはHTMLの崩れ(閉じタグ漏れ等)で失敗しやすいため、
 * 正規表現ベースの緩いパースで抽出する。
 */
function readValuesFromHtmlTable_(html) {
  const tableBlocks = splitTopLevelTables_(html);
  if (tableBlocks.length === 0) return [];

  let bestRows = null;
  let bestScore = -1;
  tableBlocks.forEach(function (block) {
    const rows = extractRowsFromTableHtml_(block);
    const populatedRowCount = rows.filter(function (row) {
      return row.some(function (cell) { return cell !== ''; });
    }).length;
    if (populatedRowCount > bestScore) {
      bestScore = populatedRowCount;
      bestRows = rows;
    }
  });

  const rows = bestRows || [];
  const maxCol = rows.reduce(function (max, r) { return Math.max(max, r.length); }, 0);
  return rows.map(function (row) {
    const padded = row.slice();
    while (padded.length < maxCol) padded.push('');
    return padded;
  });
}

/**
 * htmlのうち最上位(トップレベル)の<table>...</table>ブロックを、開閉タグの深さを数えながら
 * 分割する。<table>が入れ子になっている場合でも、外側の開始タグから対応する外側の終了タグ
 * までを1ブロックとして正しく切り出す(非貪欲な正規表現では内側の</table>で打ち切られてしまう)。
 */
function splitTopLevelTables_(html) {
  const tagRe = /<table\b[^>]*>|<\/table\s*>/gi;
  const blocks = [];
  let depth = 0;
  let start = -1;
  let match;
  while ((match = tagRe.exec(html)) !== null) {
    const isCloseTag = match[0].charAt(1) === '/';
    if (!isCloseTag) {
      if (depth === 0) start = match.index;
      depth++;
    } else if (depth > 0) {
      depth--;
      if (depth === 0 && start !== -1) {
        blocks.push(html.slice(start, match.index + match[0].length));
        start = -1;
      }
    }
  }
  return blocks;
}

/**
 * 1つの<table>ブロック(内部に入れ子テーブルを含んでもよい)から<tr>単位で行・セルを抽出する。
 * colspan指定があるセルは、その回数だけ同じ値を繰り返して列位置のズレを抑える
 * (rowspanは複雑になるため非対応。ANDPAD側の一覧データ出力では通常使われない)。
 */
function extractRowsFromTableHtml_(tableHtml) {
  const rowMatches = tableHtml.match(/<tr[\s\S]*?<\/tr>/gi) || [];
  return rowMatches.map(function (rowHtml) {
    const cellMatches = rowHtml.match(/<t[dh][\s\S]*?<\/t[dh]>/gi) || [];
    const cells = [];
    cellMatches.forEach(function (cellHtml) {
      const openTagMatch = cellHtml.match(/^<t[dh][^>]*>/i);
      const openTag = openTagMatch ? openTagMatch[0] : '';
      const colspanMatch = openTag.match(/colspan\s*=\s*["']?(\d+)/i);
      const colspan = colspanMatch ? Math.max(1, parseInt(colspanMatch[1], 10)) : 1;

      const inner = cellHtml.replace(/^<t[dh][^>]*>/i, '').replace(/<\/t[dh]>$/i, '');
      const withBreaks = inner.replace(/<br\s*\/?>/gi, '\n');
      const stripped = withBreaks.replace(/<[^>]+>/g, '');
      const value = decodeHtmlEntities_(stripped).trim();

      for (let i = 0; i < colspan; i++) cells.push(value);
    });
    return cells;
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

/**
 * このスクリプトがバインドされているスプレッドシートを返す。
 *
 * 【重要】SpreadsheetApp.openById(CONFIG.TARGET_SPREADSHEET_ID) をonEdit(e)経由の呼び出し
 * (getTargetSheet_ / getStatusHistorySheet_ / findMonthlySheet_)で使わないこと。
 * onEdit(e)はシンプルトリガーとして実行され認可(authorization)が一切ない状態で動くため、
 * 対象が同じファイルであってもID指定でスプレッドシートを開く操作は例外になる。その例外は
 * onEdit側でLogger.logされるだけで画面上には何も表示されず、「今月」シートのO列(見込確度)を
 * 編集しても報告記録・週次ステータス変更履歴への追記や案件名リンクのコピーが一切発生しない、
 * という不具合の原因になっていた。getActiveSpreadsheet()はコンテナバインド型スクリプトで
 * あれば認可不要で使え、時間主導型トリガー・シンプルトリガーいずれの実行コンテキストでも
 * バインド先のファイルを返すため、こちらを使う。
 */
function getBoundSpreadsheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getId() !== CONFIG.TARGET_SPREADSHEET_ID) {
    throw new Error(
      'このスクリプトは想定外のスプレッドシート(id: ' + ss.getId() + ')にバインドされています。' +
        '「31期予材リスト」(id: ' + CONFIG.TARGET_SPREADSHEET_ID + ')に紐づいたコンテナバインド型' +
        'プロジェクトとして設置してください。'
    );
  }
  return ss;
}

function getTargetSheet_() {
  const ss = getBoundSpreadsheet_();
  let sheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.TARGET_SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    // 新規シート作成時はCVR集計列まで含めて初期ヘッダーを作る
    // (取込日時, 元ファイル名, 案件種別, 属性, アプローチ日, 面談日, 提案日, 契約日)。
    // 既存シートの場合はここを通らず、ensureColumnsExist_ が不足列を末尾に自動追加する。
    const defaultHeader = ['取込日時', '元ファイル名'].concat(Object.keys(CONFIG.STAGE_COLUMN_ALIASES || {}));
    sheet.getRange(1, 1, 1, defaultHeader.length).setValues([defaultHeader]);
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
    notifyEditError_(err);
  }
}

/**
 * onEdit中の例外はGASのUIダイアログ(getUi().alert等)を出せないため、代わりにtoastで
 * スプレッドシート右下にエラーを表示する。Logger.logだけに埋もれて気づけなくなる
 * (「編集しても何も起きないように見える」)事態を防ぐ。
 */
function notifyEditError_(err) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(String(err), '編集時処理でエラーが発生しました', 10);
  } catch (toastErr) {
    Logger.log('エラー通知(toast)にも失敗しました: ' + toastErr);
  }
}

/**
 * シート名に「今月」を含むシートの、EDIT_MIN_ROW行目以降・EDIT_TARGET_COLUMN列目
 * (既定: O列=見込確度)の変更を検知し、「報告記録」と「週次ステータス変更履歴」の
 * 両方へ自動転記する。範囲コピー&ペーストなど複数セル一括編集にも対応する。
 * また、「週次ステータス変更履歴」のSTATUS_HISTORY_PROJECT_COLUMN(既定: C列)が手動で
 * 編集された場合は、対応するハイパーリンクの自動セットのみを行う(handleStatusHistoryProjectEdit_)。
 */
function handleEdit_(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  if (sheetName === CONFIG.STATUS_HISTORY_SHEET_NAME) {
    handleStatusHistoryProjectEdit_(e, sheet);
    return;
  }

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
 * 「今月」シートの1行分を、列名マッチングで「報告記録」シートへ追記する。
 * ヘッダー行位置(CONFIG.EDIT_HEADER_ROW)は「今月」シートの実際のレイアウトに合わせて
 * 調整すること(既定値: 4行目)。
 */
function appendEditToWeeklyReport_(sourceSheet, row, targetSheet) {
  const lastCol = sourceSheet.getLastColumn();
  const headerRow = sourceSheet.getRange(CONFIG.EDIT_HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const sourceHeader = headerRow.map(function (h) { return String(h).trim(); });
  const rowValues = sourceSheet.getRange(row, 1, 1, lastCol).getValues()[0];

  const headerMap = ensureColumnsExist_(targetSheet, sourceHeader);
  const totalCols = Math.min(targetSheet.getLastColumn(), CONFIG.MAX_DATA_COLUMNS);
  const timestamp = new Date();

  const outRow = new Array(totalCols).fill('');
  outRow[headerMap['取込日時'] - 1] = timestamp;
  outRow[headerMap['元ファイル名'] - 1] = '(シート編集) ' + sourceSheet.getName();
  sourceHeader.forEach(function (colName, idx) {
    if (!colName) return;
    const col = headerMap[colName];
    if (col && col <= totalCols) outRow[col - 1] = rowValues[idx];
  });

  const lastRow = targetSheet.getLastRow();
  targetSheet.getRange(lastRow + 1, 1, 1, totalCols).setValues([outRow]);
}

/**
 * 「週次ステータス変更履歴」シートを取得する。存在しない場合は作成し、ヘッダーを設定する。
 */
function getStatusHistorySheet_() {
  const ss = getBoundSpreadsheet_();
  let sheet = ss.getSheetByName(CONFIG.STATUS_HISTORY_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.STATUS_HISTORY_SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 8).setValues([[
      '変更日時', 'シート名', '案件名', '列名', '変更前', '変更後', '編集者', '識別情報'
    ]]);
  }
  return sheet;
}

/**
 * 見込確度などの変更を「週次ステータス変更履歴」に1行追記する。
 * 識別情報には、対象行の1列目からEDIT_IDENTIFIER_COLUMN_COUNT列目までの値のうち
 * 空でないものだけを" / "区切りで連結して記録する(案件名・得意先名などを想定)。
 * 追記後、STATUS_HISTORY_PROJECT_COLUMN(既定: C列)には「今月」タブのEDIT_URL_COLUMN
 * (既定: H列、案件名テキスト+ハイパーリンク)をそのままコピーする。
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
    '',
    columnName,
    oldValue,
    newValue,
    editorEmail,
    identifier
  ]);

  copyProjectHyperlink_(sourceSheet, row, historySheet, historySheet.getLastRow());
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

/**
 * sourceSheetのEDIT_URL_COLUMN(既定: H列、案件名テキスト+ハイパーリンク)のリッチテキストを、
 * historySheetのSTATUS_HISTORY_PROJECT_COLUMN(既定: C列)へそのままコピーする。
 * 表示文字列とリンクURLをまとめて複製するため、案件名の表示とリンク先が常に一致する。
 */
function copyProjectHyperlink_(sourceSheet, sourceRow, historySheet, historyRow) {
  const richText = sourceSheet.getRange(sourceRow, CONFIG.EDIT_URL_COLUMN).getRichTextValue();
  if (!richText) return; // セルが空、または文字列以外の場合はnullになるためコピーしない
  historySheet.getRange(historyRow, CONFIG.STATUS_HISTORY_PROJECT_COLUMN).setRichTextValue(richText);
}

/**
 * 「週次ステータス変更履歴」のSTATUS_HISTORY_PROJECT_COLUMN(既定: C列)が手動で入力・更新
 * された際、EDIT_SHEET_NAME_KEYWORD(既定:「今月」)を含むシートのEDIT_URL_COLUMN(既定: H列)
 * と案件名テキストが完全一致する行を探し、そのハイパーリンクを自動でセットする。
 */
function handleStatusHistoryProjectEdit_(e, sheet) {
  const range = e.range;
  const firstRow = range.getRow();
  const lastRow = firstRow + range.getNumRows() - 1;
  const firstCol = range.getColumn();
  const lastCol = firstCol + range.getNumColumns() - 1;
  const col = CONFIG.STATUS_HISTORY_PROJECT_COLUMN;

  if (col < firstCol || col > lastCol) return;
  if (lastRow < 2) return;

  const monthlySheet = findMonthlySheet_();
  if (!monthlySheet) return;
  const urlMap = buildProjectUrlMap_(monthlySheet);

  const startRow = Math.max(firstRow, 2);
  const endRow = Math.min(lastRow, startRow + CONFIG.EDIT_MAX_ROWS_PER_EVENT - 1);
  for (let row = startRow; row <= endRow; row++) {
    const cell = sheet.getRange(row, col);
    const text = String(cell.getValue()).trim();
    const richText = urlMap[text];
    if (richText) cell.setRichTextValue(richText);
  }
}

/**
 * 「週次ステータス変更履歴」のSTATUS_HISTORY_PROJECT_COLUMN(既定: C列)のうち、まだ
 * ハイパーリンクが設定されていないセルへ一括でリンクを設定する。EDIT_SHEET_NAME_KEYWORD
 * (既定:「今月」)を含むシートのEDIT_URL_COLUMN(既定: H列)の案件名テキストと完全一致する
 * 行を探し、そのリッチテキスト(表示文字列+リンクURL)をそのままコピーする。GASエディタから
 * 手動で実行する(既にリンク済みのセルはスキップするため、何度でも再実行できる)。
 */
function applyProjectLinksToStatusHistory() {
  const historySheet = getStatusHistorySheet_();
  const monthlySheet = findMonthlySheet_();
  if (!monthlySheet) {
    Logger.log('「' + CONFIG.EDIT_SHEET_NAME_KEYWORD + '」を含むシートが見つかりませんでした。');
    return;
  }

  const urlMap = buildProjectUrlMap_(monthlySheet);
  const lastRow = historySheet.getLastRow();
  if (lastRow < 2) return;

  const col = CONFIG.STATUS_HISTORY_PROJECT_COLUMN;
  let updatedCount = 0;
  for (let row = 2; row <= lastRow; row++) {
    const cell = historySheet.getRange(row, col);
    const richText = cell.getRichTextValue();
    if (richText && richText.getLinkUrl()) continue; // 既にリンク済み

    const text = (richText ? richText.getText() : String(cell.getValue())).trim();
    const match = urlMap[text];
    if (!match) continue;

    cell.setRichTextValue(match);
    updatedCount++;
  }
  Logger.log(updatedCount + '件のセルへハイパーリンクを設定しました。');
}

/**
 * EDIT_SHEET_NAME_KEYWORD(既定:「今月」)を含む最初のシートを返す。見つからない場合はnull。
 */
function findMonthlySheet_() {
  const ss = getBoundSpreadsheet_();
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().indexOf(CONFIG.EDIT_SHEET_NAME_KEYWORD) !== -1) return sheets[i];
  }
  return null;
}

/**
 * monthlySheetのEDIT_URL_COLUMN(既定: H列)から、{ 案件名テキスト: RichTextValue } の
 * マップを作る。案件名が重複する場合は最後に見つかった行の値を採用する。
 */
function buildProjectUrlMap_(monthlySheet) {
  const map = {};
  const lastRow = monthlySheet.getLastRow();
  if (lastRow < CONFIG.EDIT_MIN_ROW) return map;

  const range = monthlySheet.getRange(
    CONFIG.EDIT_MIN_ROW, CONFIG.EDIT_URL_COLUMN, lastRow - CONFIG.EDIT_MIN_ROW + 1, 1
  );
  range.getRichTextValues().forEach(function (cellArr) {
    const richText = cellArr[0];
    if (!richText) return;
    const text = richText.getText().trim();
    if (text !== '') map[text] = richText;
  });
  return map;
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
