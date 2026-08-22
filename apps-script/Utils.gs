/**
 * 取り込み元ファイルのパース(xlsx/HTMLテーブル/CSV・TSVの自動フォールバック判定を含む)、
 * および Drive・シート操作まわりの共通ユーティリティ関数をまとめたファイル。
 * 依存: Config.gs (CONFIG)
 */

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
