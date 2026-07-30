/**
 * 予材リスト週次同期スクリプト（単一ファイル版）
 *
 * - SOURCE_FOLDER_ID 配下のANDPAD出力.xlsx（Googleスプレッドシートも可）を読み取り、
 *   「31期予材リスト」の「週次報告記録」シート末尾に追記する。
 * - .xlsxは手動変換不要。Drive API (Advanced Service) で一時的にGoogleスプレッドシートへ変換して読み取る。
 * - 列はヘッダー名でマッチングするため、ANDPAD側の列順変更や列追加に強い。
 * - 転記済みファイルは「処理済み」フォルダへ移動し、同名ファイルがあればリネームして衝突を回避する。
 * - 万一の移動失敗に備え、処理済みファイルIDをスクリプトプロパティに記録し、再走時の二重追記を防ぐ。
 *
 * 事前準備:
 *   1. GASエディタの「サービス」から Drive API (Advanced Service, v3) を追加する。
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
  TARGET_SHEET_NAME: '週次報告記録',

  // 取り込み元ファイルの1行目がヘッダー行かどうか
  SOURCE_HAS_HEADER: true,

  // エラー発生時に通知したいメールアドレス。空文字なら通知しない。
  NOTIFY_EMAIL: '',

  TIME_ZONE: 'Asia/Tokyo',

  // 処理済みファイルIDを記録しておく数（moveTo失敗時などの二重処理を防ぐ保険）
  PROCESSED_LOG_MAX: 300
};

// ===== メイン処理 =====

/**
 * 週次実行のエントリーポイント。トリガーにはこの関数を設定する。
 */
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
  let workingSpreadsheetId;
  let tempFileId = null;

  if (mimeType === MimeType.GOOGLE_SHEETS) {
    workingSpreadsheetId = file.getId();
  } else {
    const converted = convertToTemporaryGoogleSheet_(file);
    workingSpreadsheetId = converted.id;
    tempFileId = converted.id;
  }

  try {
    const sourceSpreadsheet = SpreadsheetApp.openById(workingSpreadsheetId);
    const sourceSheet = sourceSpreadsheet.getSheets()[0];
    const values = sourceSheet.getDataRange().getValues();

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
  } finally {
    // 変換のために作った一時ファイルはゴミ箱へ送って残さない
    if (tempFileId) {
      DriveApp.getFileById(tempFileId).setTrashed(true);
    }
  }
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

/**
 * Excel(.xlsx/.xls)ファイルを一時的にGoogleスプレッドシート形式へ変換する。
 * Drive API (Advanced Service) で、アップロード時にmimeTypeを変えることで自動変換される。
 * ファイル自体のmimeTypeがxlsxとして正しく設定されていなくても、中身のバイト列から変換される。
 */
function convertToTemporaryGoogleSheet_(file) {
  const resource = {
    name: '__tmp_convert_' + file.getId(),
    mimeType: MimeType.GOOGLE_SHEETS
  };
  return Drive.Files.create(resource, file.getBlob());
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

/**
 * 処理済みファイルIDの記録をリセットする。
 * テストのため同じファイルをもう一度処理させたい場合などに、この関数だけを手動実行する。
 */
function resetProcessedLog() {
  PropertiesService.getScriptProperties().deleteProperty('PROCESSED_FILE_IDS');
  Logger.log('処理済みログをリセットしました。');
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
