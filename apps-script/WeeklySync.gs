/**
 * 週次: SOURCE_FOLDER_ID 配下に置かれたANDPAD出力の.xlsx（Googleスプレッドシートも可）を読み取り、
 * 「31期予材リスト」の「報告記録」シート末尾に追記する。
 *
 * - Excelファイルは DriveApp + Utilities.unzip + XmlService のみで直接パースする（Advanced Drive
 *   Service、UrlFetchApp、外部API呼び出しは一切使わない）。ANDPAD等の「拡張子は.xlsxだが中身は
 *   標準ZIPではない(HTMLテーブル等)」出力にも対応するため、先頭バイトで実体を判別し、
 *   ZIP/.xlsx・HTMLテーブルの2通りの読み取り方法へ自動でフォールバックする。
 * - 列はヘッダー名でマッチングして書き込むため、ANDPAD側の列順変更や列追加（フォーマットの揺れ）に強い。
 * - 転記済みファイルは「処理済み」フォルダへ移動し、同名ファイルがあればリネームして衝突を回避する。
 * - 万一の移動失敗に備え、処理済みファイルIDをプロパティに記録し、再走時の二重追記を防ぐ。
 * - 「対象期間」の区切り行は挿入せず、読み取ったデータをそのまま「報告記録」シート末尾に
 *   追記する(ソートや区切り行の再構築は行わない)。
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
 * MAX_DATA_COLUMNS(既定: 30列目 = AD列)に達している場合、それ以降の新規列は追加しない
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

function getTargetSheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.TARGET_SPREADSHEET_ID);
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
