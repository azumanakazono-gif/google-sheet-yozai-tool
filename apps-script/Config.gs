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

  // CVR集計用の正規列名。ANDPAD側の表記ゆれ(別名)を吸収し、必ずこの列名に正規化して書き込む。
  // key: 週次報告記録シート上の正規列名 / value: ANDPAD側で使われうる別名(表記ゆれ)の配列。
  // ここに無い列名は元の名前のまま追記される(自動追加)ので、ANDPAD側の実際の項目名を
  // 確認のうえ、必要な別名を随時この配列に追加すること。
  STAGE_COLUMN_ALIASES: {
    '案件種別': ['案件種別', '案件区分', '種別'],
    '属性': ['属性', '顧客属性', '反響属性'],
    'アプローチ日': ['アプローチ日', '初回アプローチ日', '初回接触日', '反響日'],
    '面談日': ['面談日', '面談実施日', '打合せ日', '初回面談日'],
    '見積日': ['見積日', '見積提出日', '御見積日', '見積書提出日'],
    '契約日': ['契約日', '成約日', '受注日']
  }
};
