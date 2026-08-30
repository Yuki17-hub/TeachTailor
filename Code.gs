/**
 * ============================================================
 *  TeachTailor — 個別最適化プリント用プロンプトジェネレーター
 *  （完全無料版 / API連携なし）
 *
 *  ■ セットアップ手順（推奨: コンテナバインド方式）
 *  1. スプレッドシートを新規作成し、「拡張機能 > Apps Script」を開く
 *  2. このコードと index.html を貼り付ける
 *  3. デプロイ > 新しいデプロイ > ウェブアプリ として公開
 *  ※ シートとヘッダー行は初回アクセス時に自動生成されます
 *  ※ 他の人に配布するときは、スプレッドシートごと「コピーを作成」
 *     してもらえば、各自のアカウント内で独立して動作します
 *  （従来どおりスクリプトプロパティ SPREADSHEET_ID で外部の
 *    スプレッドシートを指定する運用も引き続き可能です）
 * ============================================================
 */

// アプリのバージョン（配布後のサポート用。更新時はindex.html側の表記も合わせること）
var APP_VERSION = '1.5.1';

var SHEET_STUDENTS = '生徒カルテ';
var SHEET_LOGS = '学習ログ';
var SHEET_HISTORY = '出力履歴';
var SHEET_PENDING = '未処理チェック';

// teaching_policy は既存データの列ずれを避けるため末尾に追加している
var HEADERS_STUDENTS = ['id', 'display_name', 'grade', 'learning_level', 'short_term_goal', 'goal', 'environment', 'personality', 'memo', 'teaching_policy'];
var HEADERS_LOGS = ['student_id', 'date', 'item', 'status', 'note'];
var HEADERS_HISTORY = ['id', 'student_id', 'date', 'title', 'source', 'html'];
// 授業後にまとめて学習ログへ記録するための「未処理のチェック項目」。
// 処理が終わった行は削除するため、蓄積はしない（記録は学習ログ・出力履歴に残る）
var HEADERS_PENDING = ['id', 'student_id', 'datetime', 'title', 'items'];

// シートの1行目に表示するヘッダー（内部キーとは別の、人が読むための表示名）。
// 「name」は本名の登録を前提としない設計のため「管理名」とする（曜日・時間・イニシャル等での運用を想定）
var DISPLAY_HEADERS = {};
DISPLAY_HEADERS[SHEET_STUDENTS] = ['ID', '管理名', '学年', '学習度合', '短期目標', '目標', '環境', '性格', '前回のメモ', '指導方針'];
DISPLAY_HEADERS[SHEET_LOGS] = ['生徒ID', '日付', '項目', '状態', 'メモ'];
DISPLAY_HEADERS[SHEET_HISTORY] = ['ID', '生徒ID', '日付', 'タイトル', '出典', 'HTML'];
DISPLAY_HEADERS[SHEET_PENDING] = ['ID', '生徒ID', '日時', 'タイトル', 'チェック項目'];

// スプレッドシートのセル文字数制限(5万字)対策のチャンクサイズ
var HISTORY_CHUNK_SIZE = 45000;

/* ---------- 共通ヘルパー ---------- */

function getSpreadsheet_() {
  // 1. スクリプトプロパティ SPREADSHEET_ID があればそれを優先（従来の運用と互換）
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (id) {
    return SpreadsheetApp.openById(id);
  }
  // 2. なければ、スプレッドシートに紐づくスクリプト（コンテナバインド）として動作。
  //    配布時は「スプレッドシートごとコピー」してもらうだけで、設定なしで動く
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    return active;
  }
  throw new Error('スプレッドシートが見つかりません。スクリプトプロパティ「SPREADSHEET_ID」を設定するか、スプレッドシートの「拡張機能 > Apps Script」からこのコードを実行してください。');
}

/**
 * シートを取得。存在しなければ作成し、ヘッダー行を整備する。
 */
function getSheet_(name, headers) {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  var firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  var hasHeader = firstRow.some(function (v) { return String(v).trim() !== ''; });
  if (!hasHeader) {
    // 1行目は人が読むための日本語表示名を書く。
    // データの読み書きは列の順番（HEADERS_*の並び）で行っており、
    // ヘッダーの文字列は参照していないため、表示名を変えても動作に影響しない。
    var labels = DISPLAY_HEADERS[name] || headers;
    if (labels.length !== headers.length) labels = headers; // 定義漏れ時は内部キーで安全に倒す
    sheet.getRange(1, 1, 1, headers.length).setValues([labels]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sheet;
}

/**
 * Date型を含む値を安全な文字列に変換（google.script.run のシリアライズ対策）
 */
function toSafeString_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(value);
}

/* ---------- doGet ---------- */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('TeachTailor v' + APP_VERSION)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

/* ---------- 生徒カルテ ---------- */

/**
 * 全生徒データをJSON文字列で返す
 */
function getStudents() {
  var sheet = getSheet_(SHEET_STUDENTS, HEADERS_STUDENTS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return JSON.stringify([]);

  var values = sheet.getRange(2, 1, lastRow - 1, HEADERS_STUDENTS.length).getValues();
  var students = values
    .filter(function (row) { return String(row[0]).trim() !== ''; })
    .map(function (row) {
      var obj = {};
      HEADERS_STUDENTS.forEach(function (key, i) {
        obj[key] = toSafeString_(row[i]);
      });
      return obj;
    });
  return JSON.stringify(students);
}

/**
 * 生徒の新規登録・更新。
 * studentData.id が空なら新規採番、既存IDなら該当行を更新する。
 */
function saveStudent(studentData) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet_(SHEET_STUDENTS, HEADERS_STUDENTS);
    var lastRow = sheet.getLastRow();
    var id = String(studentData.id || '').trim();

    if (id !== '' && lastRow >= 2) {
      // 既存生徒の更新
      var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === id) {
          var rowValues = HEADERS_STUDENTS.map(function (key) {
            return key === 'id' ? id : String(studentData[key] || '');
          });
          sheet.getRange(i + 2, 1, 1, HEADERS_STUDENTS.length).setValues([rowValues]);
          return JSON.stringify({ ok: true, id: id, mode: 'update' });
        }
      }
    }

    // 新規登録（既存IDの最大値 + 1 を採番）
    var newId = 1;
    if (lastRow >= 2) {
      var existing = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      existing.forEach(function (row) {
        var n = parseInt(row[0], 10);
        if (!isNaN(n) && n >= newId) newId = n + 1;
      });
    }
    var newRow = HEADERS_STUDENTS.map(function (key) {
      return key === 'id' ? newId : String(studentData[key] || '');
    });
    sheet.appendRow(newRow);
    return JSON.stringify({ ok: true, id: String(newId), mode: 'create' });
  } finally {
    lock.releaseLock();
  }
}

/**
 * 授業メモ（memo列）のみを更新する。
 * カルテ全体の保存(saveStudent)とは独立して、ログ画面から素早く保存するために使う。
 */
function updateStudentMemo(studentId, memo) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet_(SHEET_STUDENTS, HEADERS_STUDENTS);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error('生徒が見つかりません');

    var memoCol = HEADERS_STUDENTS.indexOf('memo') + 1;
    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(studentId)) {
        sheet.getRange(i + 2, memoCol).setValue(String(memo || ''));
        return JSON.stringify({ ok: true });
      }
    }
    throw new Error('生徒が見つかりません: ID ' + studentId);
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 学習ログ ---------- */

/**
 * 指定生徒の全ログを返す（rowIndex = 実際のシート行番号を含む）。
 * プロンプト生成時の「未定着」フィルタリングはフロント側で行う。
 */
function getLogs(studentId) {
  var sheet = getSheet_(SHEET_LOGS, HEADERS_LOGS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return JSON.stringify([]);

  var values = sheet.getRange(2, 1, lastRow - 1, HEADERS_LOGS.length).getValues();
  var logs = [];
  values.forEach(function (row, i) {
    if (String(row[0]) === String(studentId)) {
      logs.push({
        rowIndex: i + 2, // シート上の実行番号（1行目はヘッダー）
        student_id: toSafeString_(row[0]),
        date: toSafeString_(row[1]),
        item: toSafeString_(row[2]),
        status: toSafeString_(row[3]),
        note: toSafeString_(row[4])
      });
    }
  });
  return JSON.stringify(logs);
}

/**
 * 新しい未定着項目を追記する（デフォルトステータス: 未定着）
 */
function saveLog(studentId, item) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet_(SHEET_LOGS, HEADERS_LOGS);
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    sheet.appendRow([String(studentId), today, String(item), '未定着', '']);
    return JSON.stringify({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

/**
 * 複数の未定着項目を1回の通信でまとめて追記する（デフォルトステータス: 未定着）
 * items: 文字列配列。setValues() による一括書き込みで実装し、1件ずつの
 * appendRow ループは避ける。
 */
function saveLogsBulk(studentId, items) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var validItems = (items || []).filter(function (item) {
      return String(item || '').trim() !== '';
    });
    if (validItems.length === 0) {
      return JSON.stringify({ ok: true, count: 0 });
    }

    var sheet = getSheet_(SHEET_LOGS, HEADERS_LOGS);
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var startRow = sheet.getLastRow() + 1;

    var rows = validItems.map(function (item) {
      return [String(studentId), today, String(item), '未定着', ''];
    });

    sheet.getRange(startRow, 1, rows.length, HEADERS_LOGS.length).setValues(rows);
    return JSON.stringify({ ok: true, count: rows.length });
  } finally {
    lock.releaseLock();
  }
}

/**
 * 指定項目のステータスを一括で切り替える（「定着」⇔「未定着」双方向）。
 * 同じ項目が複数回登録されていても全行に適用する。
 * 生徒ID＋項目名の完全一致で対象を特定するため、行ズレの影響を受けない。
 */
function setLogStatusByItem(studentId, item, newStatus) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet_(SHEET_LOGS, HEADERS_LOGS);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error('データが更新されています。ログ一覧を再読み込みします。');

    var values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    var status = (newStatus === '定着') ? '定着' : '未定着';
    var count = 0;
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][0]) === String(studentId) && String(values[i][2]) === String(item)) {
        sheet.getRange(i + 2, 4).setValue(status); // 4列目 = status
        count++;
      }
    }
    if (count === 0) throw new Error('データが更新されています。ログ一覧を再読み込みします。');
    return JSON.stringify({ ok: true, status: status, count: count });
  } finally {
    lock.releaseLock();
  }
}

/**
 * 指定項目のログをすべて削除する（重複登録もまとめて整理）。
 * 行番号のズレを避けるため、下の行から順に削除する。
 */
function deleteLogsByItem(studentId, item) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet_(SHEET_LOGS, HEADERS_LOGS);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error('データが更新されています。ログ一覧を再読み込みします。');

    var values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    var targetRows = [];
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][0]) === String(studentId) && String(values[i][2]) === String(item)) {
        targetRows.push(i + 2);
      }
    }
    if (targetRows.length === 0) throw new Error('データが更新されています。ログ一覧を再読み込みします。');
    // 下の行から削除して行番号のズレを防ぐ
    for (var j = targetRows.length - 1; j >= 0; j--) {
      sheet.deleteRow(targetRows[j]);
    }
    return JSON.stringify({ ok: true, count: targetRows.length });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 出力履歴（再印刷・授業の振り返り用） ---------- */

/**
 * 生成済みプリントのHTMLを履歴として保存する。
 * HTMLはセル文字数制限を避けるため、6列目以降にチャンク分割して保存する。
 */
function saveHistory(studentId, title, source, html, checkItems) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet_(SHEET_HISTORY, HEADERS_HISTORY);
    var lastRow = sheet.getLastRow();

    // 新規ID採番
    var newId = 1;
    if (lastRow >= 2) {
      var existing = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      existing.forEach(function (row) {
        var n = parseInt(row[0], 10);
        if (!isNaN(n) && n >= newId) newId = n + 1;
      });
    }

    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var htmlStr = String(html || '');

    // チャンク分割
    var chunks = [];
    for (var i = 0; i < htmlStr.length; i += HISTORY_CHUNK_SIZE) {
      chunks.push(htmlStr.substring(i, i + HISTORY_CHUNK_SIZE));
    }
    if (chunks.length === 0) chunks.push('');

    var row = [newId, String(studentId), today, String(title || '無題プリント'), String(source || '')].concat(chunks);
    sheet.appendRow(row);

    // このプリントのチェック項目を「未処理」として預かる。
    // 授業後にまとめて学習ログへ記録できるようにするためで、
    // 「授業で使う＝履歴に保存する」という講師の操作をそのまま起点にしている。
    var items = (checkItems || []).filter(function (s) { return String(s || '').trim() !== ''; });
    if (items.length > 0) {
      var ps = getSheet_(SHEET_PENDING, HEADERS_PENDING);
      var pLast = ps.getLastRow();
      var pId = 1;
      if (pLast >= 2) {
        ps.getRange(2, 1, pLast - 1, 1).getValues().forEach(function (r) {
          var n = parseInt(r[0], 10);
          if (!isNaN(n) && n >= pId) pId = n + 1;
        });
      }
      var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
      ps.appendRow([pId, String(studentId), stamp, String(title || '無題プリント'),
                    JSON.stringify(items)]);
    }
    return JSON.stringify({ ok: true, id: String(newId) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 未処理チェック（授業後のログ記録） ---------- */

/**
 * 未処理のチェック項目をすべて返す（新しい順）。
 * 生徒名はフロント側で解決するため、ここでは生徒IDのみ返す。
 */
function getPendingChecks() {
  var sheet = getSheet_(SHEET_PENDING, HEADERS_PENDING);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return JSON.stringify([]);

  var values = sheet.getRange(2, 1, lastRow - 1, HEADERS_PENDING.length).getValues();
  var list = [];
  values.forEach(function (row) {
    var items = [];
    try { items = JSON.parse(String(row[4])); } catch (e) { items = []; }
    list.push({
      id: toSafeString_(row[0]),
      student_id: toSafeString_(row[1]),
      datetime: toSafeString_(row[2]),
      title: toSafeString_(row[3]),
      items: items
    });
  });
  list.sort(function (a, b) { return parseInt(b.id, 10) - parseInt(a.id, 10); });
  return JSON.stringify(list);
}

/**
 * 未処理チェックを完了させる。
 * items が指定されていれば、それを未定着として学習ログに登録する。
 * このシートは「処理待ちのキュー」であり記録ではないため、
 * 完了した行は削除する（プリント本体は出力履歴に、記録した項目は学習ログに残る）。
 */
function resolvePendingCheck(pendingId, studentId, items) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var registered = 0;
    var valid = (items || []).filter(function (s) { return String(s || '').trim() !== ''; });
    if (valid.length > 0) {
      var logSheet = getSheet_(SHEET_LOGS, HEADERS_LOGS);
      var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      var rows = valid.map(function (item) {
        return [String(studentId), today, String(item), '未定着', ''];
      });
      logSheet.getRange(logSheet.getLastRow() + 1, 1, rows.length, HEADERS_LOGS.length).setValues(rows);
      registered = rows.length;
    }

    var sheet = getSheet_(SHEET_PENDING, HEADERS_PENDING);
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === String(pendingId)) {
          sheet.deleteRow(i + 2);
          return JSON.stringify({ ok: true, count: registered });
        }
      }
    }
    throw new Error('データが更新されています。画面を再読み込みします。');
  } finally {
    lock.releaseLock();
  }
}

/**
 * 指定生徒の履歴一覧を返す（HTML本体は含めず、メタ情報のみで軽量に）
 */
function getHistoryList(studentId) {
  var sheet = getSheet_(SHEET_HISTORY, HEADERS_HISTORY);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return JSON.stringify([]);

  // メタ情報の5列だけ読む（HTMLチャンク列は読まない）
  var values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  var list = [];
  values.forEach(function (row) {
    if (String(row[1]) === String(studentId)) {
      list.push({
        id: toSafeString_(row[0]),
        student_id: toSafeString_(row[1]),
        date: toSafeString_(row[2]),
        title: toSafeString_(row[3]),
        source: toSafeString_(row[4])
      });
    }
  });
  // 新しい順（ID降順）
  list.sort(function (a, b) { return parseInt(b.id, 10) - parseInt(a.id, 10); });
  return JSON.stringify(list);
}

/**
 * 履歴IDを指定してHTML本体を返す（チャンクを結合して復元）
 */
function getHistoryHtml(historyId) {
  var sheet = getSheet_(SHEET_HISTORY, HEADERS_HISTORY);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('履歴が見つかりません');

  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(historyId)) {
      var rowNum = i + 2;
      var lastCol = sheet.getLastColumn();
      if (lastCol < 6) return JSON.stringify({ ok: true, html: '' });
      var cells = sheet.getRange(rowNum, 6, 1, lastCol - 5).getValues()[0];
      var html = cells
        .map(function (c) { return String(c); })
        .join('');
      return JSON.stringify({ ok: true, html: html });
    }
  }
  throw new Error('履歴が見つかりません: ID ' + historyId);
}

/**
 * 履歴を削除する（IDで行を特定して削除）
 */
function deleteHistory(historyId) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet_(SHEET_HISTORY, HEADERS_HISTORY);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error('履歴が見つかりません');

    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(historyId)) {
        sheet.deleteRow(i + 2);
        return JSON.stringify({ ok: true });
      }
    }
    throw new Error('履歴が見つかりません: ID ' + historyId);
  } finally {
    lock.releaseLock();
  }
}
