/**
 * 週次同期スクリプトの設定値。
 * 運用環境に合わせてここだけを書き換えれば動くようにしてある。
 */
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

  // ===== 対象期間の自動判定 =====
  // 対象期間はデータ列(J列等)には記録せず、データ群先頭の区切り行として挿入する。
  // 区切り行のA列に入るラベルの接頭辞・接尾辞（例:「【対象期間：2026/07/20 〜 2026/07/26】」）
  PERIOD_SEPARATOR_PREFIX: '【対象期間：',
  PERIOD_SEPARATOR_SUFFIX: '】',
  // 対象期間が判定できなかった場合に区切り行へ表示する文言
  PERIOD_UNKNOWN_LABEL: '不明',
  // ファイル名から日付が拾えなかった場合に、取り込みデータの中から探す列名の候補
  PERIOD_DATE_COLUMN_CANDIDATES: ['報告日', '日付', '対象日', '実施日', '報告日時', '登録日'],
  // 区切り行をA列からこの列数まで結合する（既定: 19列目 = S列）
  PERIOD_SEPARATOR_MERGE_COLUMNS: 19,
  // 区切り行の背景色
  PERIOD_SEPARATOR_BACKGROUND_COLOR: '#f3f3f3',

  // ===== 報告日順の自動ソート =====
  // syncWeeklyReportsの最後に、この列名をヘッダーから探して昇順ソートする。
  REPORT_DATE_COLUMN_NAME: '報告日',
  // ヘッダーに上記の列名が見つからない場合に使う列番号（既定: 9列目 = I列）
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
