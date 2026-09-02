/**
 * メイン処理: syncWeeklyReports のエントリポイントと、処理済みログのリセット関数をまとめたファイル。
 * トリガー設定用ヘルパーは Triggers.gs を参照。
 * 依存: Config.gs (CONFIG) / Utils.gs (getOrCreateSubFolder_, getTargetSheet_,
 * isSupportedSpreadsheet_, isAlreadyProcessed_, markAsProcessed_, moveToFolderSafely_) /
 * ReportSync.gs (appendFileToTargetSheet_)
 */

/**
 * 週次: SOURCE_FOLDER_ID 配下に置かれたANDPAD出力の.xlsx（Googleスプレッドシートも可）を読み取り、
 * 「31期予材リスト」の「報告記録」シート末尾に追記する。
 *
 * - Excelファイルは DriveApp + Utilities.unzip + XmlService のみで直接パースする（Advanced Drive
 *   Service、UrlFetchApp、外部API呼び出しは一切使わない）。ANDPAD等の「拡張子は.xlsxだが中身は
 *   標準ZIPではない(HTMLテーブルやCSV/TSV等)」出力にも対応するため、先頭バイトや中身のテキストで
 *   実体を判別し、ZIP/.xlsx・HTMLテーブル・CSV/TSVの読み取り方法へ自動でフォールバックする。
 * - 列はヘッダー名でマッチングして書き込むため、ANDPAD側の列順変更や列追加（フォーマットの揺れ）に強い。
 * - 転記済みファイルは「処理済み」フォルダへ移動し、同名ファイルがあればリネームして衝突を回避する。
 * - 万一の移動失敗に備え、処理済みファイルIDをプロパティに記録し、再走時の二重追記を防ぐ。
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
    // CONFIG.TARGET_SPREADSHEET_IDが空文字(=getActiveSpreadsheet()で自動解決)の場合でも
    // 対象スプレッドシート自身をソースファイルとして誤って処理しないよう、解決済みの
    // targetSheetから実際のIDを取得して比較する(CONFIG.TARGET_SPREADSHEET_IDを直接見ない)。
    const targetSpreadsheetId = targetSheet.getParent().getId();

    const files = sourceFolder.getFiles();
    const errors = [];
    let processedCount = 0;
    let skippedCount = 0;

    while (files.hasNext()) {
      const file = files.next();

      if (!isSupportedSpreadsheet_(file)) continue;
      if (file.getId() === targetSpreadsheetId) continue;

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

// ===== ログリセット =====

/**
 * 処理済みファイルIDの記録をリセットする。
 * テストのため同じファイルをもう一度処理させたい場合などに、この関数だけを手動実行する。
 */
function resetProcessedLog() {
  PropertiesService.getScriptProperties().deleteProperty('PROCESSED_FILE_IDS');
  Logger.log('処理済みログをリセットしました。');
}
