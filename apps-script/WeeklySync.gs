/**
 * 週次: SOURCE_FOLDER_ID 配下のExcel/スプレッドシートを読み取り、
 * 「31期予材リスト」の「週次報告記録」シート末尾に追記する。
 * 処理が終わったファイルは「処理済み」フォルダへ移動し、二重処理を防ぐ。
 * 失敗したファイルは「エラー」フォルダへ移動し、次回以降のスキャン対象から外す。
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

    while (files.hasNext()) {
      const file = files.next();

      if (!isSupportedSpreadsheet_(file)) continue;
      if (file.getId() === CONFIG.TARGET_SPREADSHEET_ID) continue;

      try {
        appendFileToTargetSheet_(file, targetSheet);
        file.moveTo(processedFolder);
        processedCount++;
      } catch (err) {
        Logger.log('処理失敗: ' + file.getName() + ' / ' + err);
        errors.push(file.getName() + ': ' + err);
        try {
          file.moveTo(errorFolder);
        } catch (moveErr) {
          Logger.log('エラーフォルダへの移動にも失敗: ' + file.getName() + ' / ' + moveErr);
        }
      }
    }

    Logger.log('処理完了: ' + processedCount + '件 / 失敗: ' + errors.length + '件');

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
 * Excel形式のファイルは一時的にGoogleスプレッドシートへ変換して読み取る。
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

    const startRow = CONFIG.SOURCE_HAS_HEADER ? 1 : 0;
    const dataRows = values.slice(startRow).filter(function (row) {
      return row.some(function (cell) { return cell !== '' && cell !== null; });
    });

    if (dataRows.length === 0) {
      Logger.log('データなし: ' + file.getName());
      return;
    }

    const timestamp = new Date();
    const rowsToAppend = dataRows.map(function (row) {
      return [timestamp, file.getName()].concat(row);
    });

    const lastRow = targetSheet.getLastRow();
    targetSheet
      .getRange(lastRow + 1, 1, rowsToAppend.length, rowsToAppend[0].length)
      .setValues(rowsToAppend);
  } finally {
    // 変換のために作った一時ファイルはゴミ箱へ送って残さない
    if (tempFileId) {
      DriveApp.getFileById(tempFileId).setTrashed(true);
    }
  }
}

/**
 * Excel(.xlsx/.xls)ファイルを一時的にGoogleスプレッドシート形式へ変換する。
 * Drive API (Advanced Service) で、アップロード時にmimeTypeを変えることで自動変換される。
 */
function convertToTemporaryGoogleSheet_(file) {
  const resource = {
    name: '__tmp_convert_' + file.getId(),
    mimeType: MimeType.GOOGLE_SHEETS
  };
  return Drive.Files.create(resource, file.getBlob());
}

function isSupportedSpreadsheet_(file) {
  const mimeType = file.getMimeType();
  return (
    mimeType === MimeType.GOOGLE_SHEETS ||
    mimeType === MimeType.MICROSOFT_EXCEL ||
    mimeType === 'application/vnd.ms-excel'
  );
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
