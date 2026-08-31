/**
 * 「報告記録」シートへの追記処理。
 * 依存: Config.gs (CONFIG) / Utils.gs (parseWeeklyReportFile_, ensureColumnsExist_,
 * writeValuesIgnoringValidation_) / StatusHistory.gs (updateWeeklyStatusHistory_)
 */

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
