// ============================================================
// Google Apps Script — 社労士業務管理ツール
// スプレッドシート: TODO / 記録 / 給与計算記録 / クライアント / 行政問い合わせ記録
// ============================================================

// --------------- 定数 ---------------
var SHEET_TODO    = 'TODO';
var SHEET_RECORD  = '記録';
var SHEET_PAYROLL = '給与計算記録';
var SHEET_CLIENT  = 'クライアント';
var SHEET_INQUIRY = '行政問い合わせ記録';
var SHEET_CONSULT = '相談記録';
var SHEET_BUSY    = '算定年更管理';
var SHEET_STRESS  = 'ストレスチェック管理';
var SHEET_ACTIVITY = 'アクティビティログ';
var SHEET_HOLIDAY  = '休日マスタ';
var SHEET_MESSAGE  = '一言メッセージ';
var SHEET_FEE_MASTER   = '顧問料設定';
var SHEET_FEE_MONTHLY  = '月別請求管理';
var SHEET_WORK_AVAIL   = '勤務予定DB';

// 給与計算系の業務種別（ここに追加すれば分岐が増やせる）
var PAYROLL_CATEGORIES = ['給与計算', '賞与計算', '給与修正・再計算', '会計入力'];

// --------------- doGet: シートデータを返す ---------------
// ?sheet=シート名 で取得先を切り替え（最優先）
// ?action=xxx    で専用ハンドラを呼び出し
// パラメータなしのデフォルト: TODO
function doGet(e) {
  var action    = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';
  var sheetName = (e && e.parameter && e.parameter.sheet)  ? e.parameter.sheet  : '';
  Logger.log('[doGet] action=' + action + ' sheet=' + sheetName);

  // ── sheet パラメータで明示指定（最優先）──
  if (sheetName === 'ストレスチェック管理') {
    Logger.log('[doGet] ストレスチェック管理シート取得');
    return getSheetData_(SHEET_STRESS);
  }
  if (sheetName === 'クライアント') {
    Logger.log('[doGet] クライアントシート取得');
    return getSheetData_(SHEET_CLIENT);
  }
  if (sheetName === '算定年更管理') {
    Logger.log('[doGet] 算定年更管理シート取得');
    return getSheetData_(SHEET_BUSY);
  }
  if (sheetName === '記録') {
    Logger.log('[doGet] 記録シート取得');
    return getSheetData_(SHEET_RECORD);
  }
  if (sheetName === '給与計算記録') {
    Logger.log('[doGet] 給与計算記録シート取得');
    return getSheetData_(SHEET_PAYROLL);
  }
  if (sheetName) {
    Logger.log('[doGet] 汎用シート取得: ' + sheetName);
    return getSheetData_(sheetName);
  }

  // ── action ハンドラ ──
  if (action === 'getConsultTodos')       return handleGetConsultTodos_();
  if (action === 'getBusySeasonRecords')  return handleGetBusySeasonRecords_(e);
  if (action === 'getStressCheckRecords') return handleGetStressCheckRecords_();
  if (action === 'getActivityLog')        return handleGetActivityLog_(e);
  if (action === 'getHolidayMaster')      return handleGetHolidayMaster_();
  if (action === 'getMemberStatuses')     return handleGetMemberStatuses_();
  if (action === 'getMessages')           return handleGetMessages_();
  if (action === 'getFeeMonthly')         return handleGetFeeMonthly_(e);
  if (action === 'getWorkAvailability')   return handleGetWorkAvailability_(e);

  // ── デフォルト: TODO シート ──
  Logger.log('[doGet] 取得シート: ' + SHEET_TODO + ' (デフォルト)');
  return getSheetData_(SHEET_TODO);
}

// 指定シートの全行をオブジェクト配列で返す汎用ヘルパー
function getSheetData_(sheetName) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return jsonResponse_({ error: 'シート「' + sheetName + '」が見つかりません' });
  }
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var result  = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j];
    }
    result.push(obj);
  }
  Logger.log('[getSheetData_] シート=' + sheetName + ' 件数=' + result.length);
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// --------------- 相談記録TODO取得 ---------------
function handleGetConsultTodos_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_CONSULT);
  if (!sheet) {
    return jsonResponse_({ error: 'シート「' + SHEET_CONSULT + '」が見つかりません' });
  }

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  Logger.log('[getConsultTodos] ヘッダー: ' + JSON.stringify(headers));

  // ヘッダーから列インデックスを取得
  var flagCol   = headers.indexOf('要対応フラグ');
  var statusCol = headers.indexOf('ステータス');
  Logger.log('[getConsultTodos] 要対応フラグ列=' + flagCol + ', ステータス列=' + statusCol);

  var result = [];
  for (var i = 1; i < data.length; i++) {
    // 要対応フラグ: TRUE / "TRUE" / true を全て吸収
    var flagVal = (flagCol >= 0) ? String(data[i][flagCol]).toUpperCase().trim() : '';
    var statusVal = (statusCol >= 0) ? String(data[i][statusCol]).trim() : '';

    if (flagVal !== 'TRUE') continue;
    if (statusVal === '完了') continue;

    var obj = { _rowIndex: i + 1 }; // スプレッドシートの行番号
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j];
    }
    result.push(obj);
  }

  Logger.log('[getConsultTodos] 対象件数: ' + result.length);
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// --------------- doPost: アクション振り分け ---------------
function doPost(e) {
  try {
    // URLSearchParams(data=JSON) または application/json の両方に対応
    var rawContents = (e.parameter && e.parameter.data) ? e.parameter.data : e.postData.contents;
    Logger.log('[doPost] ★受信rawデータ: ' + rawContents);
    var data = JSON.parse(rawContents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    Logger.log('[doPost] action=' + (data.action || '(なし)'));
    Logger.log('[doPost] ★data全体: ' + JSON.stringify(data));

    // ── アクティビティログ追加 ──
    if (data.action === 'logActivity') {
      return handleLogActivity_(ss, data);
    }

    // ── TODO完了アクション（完了処理を一括実行） ──
    if (data.action === 'completeTodo') {
      Logger.log('[doPost] ★completeTodo呼出前: 開始時刻=' + data['開始時刻'] + ', 終了時刻=' + data['終了時刻'] + ', 作業時間=' + data['作業時間']);
      return handleCompleteTodo_(ss, data);
    }

    // ── 相談記録TODO完了（O列をFALSEに） ──
    if (data.action === 'completeConsultTodo') {
      return handleCompleteConsultTodo_(ss, data);
    }

    // ── 担当者状態更新 ──
    if (data.action === 'setMemberStatus') {
      return handleSetMemberStatus_(data);
    }

    // ── 一言メッセージ投稿 ──
    if (data.action === 'postMessage') {
      return handlePostMessage_(data);
    }

    // ── 一言メッセージ削除 ──
    if (data.action === 'deleteMessage') {
      return handleDeleteMessage_(data);
    }

    // ── 算定年更管理 一括保存 ──
    if (data.action === 'saveBusySeasonRecords') {
      return handleSaveBusySeasonRecords_(ss, data);
    }

    // ── ストレスチェック管理 一括保存（事業所名キー upsert） ──
    if (data.action === 'saveStressCheckRecords') {
      return handleSaveStressCheckRecords_(ss, data);
    }

    // ── ストレスチェック管理 新規追加（旧・互換のため残す） ──
    if (data.action === 'saveStressCheckRecord') {
      return handleSaveStressCheckRecord_(ss, data);
    }

    // ── ストレスチェック管理 更新 ──
    if (data.action === 'updateStressCheckRecord') {
      return handleUpdateStressCheckRecord_(ss, data);
    }

    // ── ストレスチェック 重要フラグ専用更新 ──
    if (data.action === 'updateStressCheckFlag') {
      return handleUpdateStressCheckFlag_(ss, data);
    }

    // ── 行更新アクション ──
    if (data.action === 'updateRow') {
      return handleUpdateRow_(ss, data);
    }

    // ── 顧問料設定 一括保存 ──
    if (data.action === 'saveFeeMaster') {
      return handleSaveFeeMaster_(ss, data);
    }

    // ── 月別請求管理 一括保存 ──
    if (data.action === 'saveFeeMonthly') {
      return handleSaveFeeMonthly_(ss, data);
    }

    // ── 勤務予定保存 ──
    if (data.action === 'saveWorkAvailability') {
      return handleSaveWorkAvailability_(ss, data);
    }

    // ── 従来の新規追加処理 ──
    var sheetName = data.sheet || SHEET_TODO;
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return jsonResponse_({ error: 'シート「' + sheetName + '」が見つかりません' });
    }

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var nowJST = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

    // IDの自動採番（TODO新規追加時）
    if (sheetName === SHEET_TODO && !data['ID']) {
      data['ID'] = getNextId_(sheet, headers);
    }

    // クライアントシート: client_id・作成日時・更新日時・有効フラグを自動設定
    if (sheetName === SHEET_CLIENT) {
      if (!data['client_id']) {
        data['client_id'] = getNextIdByCol_(sheet, headers, 'client_id');
      }
      if (!data['作成日時']) data['作成日時'] = nowJST;
      if (!data['更新日時']) data['更新日時'] = nowJST;
      if (data['有効フラグ'] === undefined || data['有効フラグ'] === '') data['有効フラグ'] = 'TRUE';
      Logger.log('[doPost] クライアント新規追加: client_id=' + data['client_id']);
    }

    // 行政問い合わせ記録シート: inquiry_id・作成日時・更新日時・有効フラグを自動設定
    if (sheetName === SHEET_INQUIRY) {
      if (!data['inquiry_id']) {
        data['inquiry_id'] = getNextIdByCol_(sheet, headers, 'inquiry_id');
      }
      if (!data['作成日時']) data['作成日時'] = nowJST;
      if (!data['更新日時']) data['更新日時'] = nowJST;
      if (data['有効フラグ'] === undefined || data['有効フラグ'] === '') data['有効フラグ'] = 'TRUE';
      Logger.log('[doPost] 行政問い合わせ新規追加: inquiry_id=' + data['inquiry_id']);
    }

    // 相談記録シート: ヘッダー名の揺れ吸収
    if (sheetName === SHEET_CONSULT) {
      Logger.log('[doPost] ★相談記録 ヘッダー: ' + JSON.stringify(headers));
      // 「作業時間（分）」→「作業時間」のマッピング
      if (data['作業時間（分）'] !== undefined && data['作業時間'] === undefined) {
        data['作業時間'] = data['作業時間（分）'];
      }
      // 「要対応フラグ（TRUE / FALSE）」→「要対応フラグ」のマッピング
      if (data['要対応フラグ（TRUE / FALSE）'] !== undefined && data['要対応フラグ'] === undefined) {
        data['要対応フラグ'] = data['要対応フラグ（TRUE / FALSE）'];
      }
    }

    var row = headers.map(function(h) {
      return data[h] !== undefined ? data[h] : '';
    });

    // appendRow 実行前ログ
    var lastRowBefore = sheet.getLastRow();
    Logger.log('[doPost] ★appendRow実行前: シート=' + sheetName + ', 最終行=' + lastRowBefore);
    Logger.log('[doPost] ★書き込み配列: ' + JSON.stringify(row));

    sheet.appendRow(row);

    // appendRow 実行後ログ
    var lastRowAfter = sheet.getLastRow();
    Logger.log('[doPost] ★appendRow実行後: 最終行=' + lastRowAfter);

    if (lastRowAfter <= lastRowBefore) {
      Logger.log('[doPost] ❌ appendRow後に行が増えていない');
      return jsonResponse_({ success: false, error: 'appendRowが反映されませんでした', sheetName: sheetName });
    }

    Logger.log('[doPost] ✅ 新規追加成功: シート=' + sheetName + ', 行=' + lastRowAfter);
    return jsonResponse_({
      success: true,
      sheetName: sheetName,
      writtenRow: lastRowAfter,
      savedData: row,
      headers: headers
    });

  } catch (err) {
    Logger.log('[doPost] ❌ エラー: ' + err.message);
    return jsonResponse_({ success: false, error: err.message });
  }
}

// ===============================================================
// TODO完了処理（一括）
// ブラウザ → completeTodo → TODO更新＋記録保存＋記録済セット
// ===============================================================
function handleCompleteTodo_(ss, data) {
  var todoSheet = ss.getSheetByName(SHEET_TODO);
  if (!todoSheet) return jsonResponse_({ error: 'TODOシートが見つかりません' });

  var allData = todoSheet.getDataRange().getValues();
  var headers = allData[0];
  var idCol = headers.indexOf('ID');
  if (idCol === -1) return jsonResponse_({ error: 'ID列が見つかりません' });

  var targetId = String(data.id);
  Logger.log('[completeTodo] 対象ID: ' + targetId);

  // --- 対象行を探す ---
  var rowNum = -1;
  for (var i = 1; i < allData.length; i++) {
    if (String(allData[i][idCol]) === targetId) {
      rowNum = i + 1;
      break;
    }
  }
  if (rowNum === -1) {
    Logger.log('[completeTodo] ID=' + targetId + ' が見つかりません');
    return jsonResponse_({ error: '該当ID(' + targetId + ')がTODOシートに見つかりません' });
  }
  Logger.log('[completeTodo] ID=' + targetId + ' → 行' + rowNum);

  // --- 二重保存チェック ---
  var recordedCol = headers.indexOf('記録済');
  if (recordedCol !== -1) {
    var recordedVal = todoSheet.getRange(rowNum, recordedCol + 1).getValue();
    if (recordedVal !== '' && recordedVal !== null && recordedVal !== undefined) {
      Logger.log('[completeTodo] 既に記録済 → スキップ');
      return jsonResponse_({ success: true, skipped: true, message: '既に記録済です' });
    }
  }

  // --- TODOシートの状態を「完了」に更新 ---
  var statusCol = headers.indexOf('状態');
  if (statusCol !== -1) {
    todoSheet.getRange(rowNum, statusCol + 1).setValue('完了');
  }

  // 更新日もセット（JST）
  var updatedCol = headers.indexOf('更新日');
  var todoNowJST = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  if (updatedCol !== -1) {
    todoSheet.getRange(rowNum, updatedCol + 1).setValue(todoNowJST);
  }

  // --- 最新の行データを取得 ---
  var colCount = headers.length;
  var todoValues = todoSheet.getRange(rowNum, 1, 1, colCount).getValues()[0];

  // ヘルパー：列名から値を取得
  function tv(colName) {
    var idx = headers.indexOf(colName);
    return idx !== -1 ? todoValues[idx] : '';
  }

  // --- 業務種別を判定（C列を参照、trimして比較） ---
  var categoryRaw = tv('業務種別');
  var category = String(categoryRaw || '').trim();
  var isPayroll = PAYROLL_CATEGORIES.indexOf(category) !== -1;
  var targetSheetName = isPayroll ? SHEET_PAYROLL : SHEET_RECORD;
  Logger.log('[completeTodo] 業務種別(元値): "' + categoryRaw + '"');
  Logger.log('[completeTodo] 業務種別(trim後): "' + category + '"');
  Logger.log('[completeTodo] 分岐結果: ' + (isPayroll ? '給与計算記録' : '記録') + ' → 保存先: ' + targetSheetName);

  // --- 記録シートへ保存 ---
  var targetSheet = ss.getSheetByName(targetSheetName);
  if (!targetSheet) {
    Logger.log('[completeTodo] 保存先シート「' + targetSheetName + '」が見つかりません！');
    return jsonResponse_({ error: '保存先シート「' + targetSheetName + '」が見つかりません' });
  }

  // --- POSTデータから作業時間を取得（フロントから直接渡される） ---
  var postWorkStart   = data['開始時刻'] || '';
  var postWorkEnd     = data['終了時刻'] || '';
  var postWorkMinutes = data['作業時間'] || '';
  Logger.log('[completeTodo] POST受信: 開始時刻=' + postWorkStart + ', 終了時刻=' + postWorkEnd + ', 作業時間=' + postWorkMinutes);

  if (isPayroll) {
    // --- 給与計算記録シートへ保存 ---
    var payrollMonth = '';
    var dueVal = tv('期限');
    if (dueVal) {
      var dueDate = new Date(dueVal);
      if (!isNaN(dueDate.getTime())) {
        payrollMonth = dueDate.getFullYear() + '-' + ('0' + (dueDate.getMonth() + 1)).slice(-2);
      } else {
        payrollMonth = String(dueVal);
      }
    }
    var workMinutes = convertToMinutes_(postWorkMinutes);
    var newId = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss');

    var nowJST = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    var payrollRow = [
      newId,                 // A: ID
      tv('ID'),              // B: 元タスクID
      tv('クライアント'),    // C
      payrollMonth,          // D: 給与対象月
      tv('タスク内容'),      // E
      tv('担当者'),          // F
      nowJST,                // G: 完了日（JST）
      workMinutes,           // H: 作業時間（分）
      '',                    // I: 人数（空欄）
      '',                    // J: メモ（空欄）
      '',                    // K: （空欄）
      postWorkStart,         // L: 開始時刻（POSTから）
      postWorkEnd,           // M: 終了時刻（POSTから）
      postWorkMinutes        // N: 作業時間（POSTから）
    ];
    Logger.log('[completeTodo] ★給与計算記録 書込直前: payrollRow[11]=' + payrollRow[11] + ', [12]=' + payrollRow[12] + ', [13]=' + payrollRow[13]);
    Logger.log('[completeTodo] ★payrollRow全体: ' + JSON.stringify(payrollRow));
    targetSheet.appendRow(payrollRow);
    Logger.log('[completeTodo] 給与計算記録に追加完了: 元ID=' + tv('ID') + ', クライアント=' + tv('クライアント'));

  } else {
    // --- 記録シートへ保存 ---
    var nowJST2 = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    var recordRow = [
      tv('ID'),             // A: ID
      tv('クライアント'),    // B
      tv('業務種別'),        // C
      tv('タスク内容'),      // D
      tv('期限'),            // E
      tv('担当者'),          // F
      tv('状態'),            // G（完了）
      tv('優先度'),          // H
      tv('作成日'),          // I
      nowJST2,               // J: 更新日（JST）
      nowJST2,               // K: 完了日時（JST）
      postWorkStart,         // L: 開始時刻（POSTから）
      postWorkEnd,           // M: 終了時刻（POSTから）
      postWorkMinutes,       // N: 作業時間（POSTから）
      'TODO',                // O: 元シート
      rowNum                 // P: 元行番号
    ];
    Logger.log('[completeTodo] ★記録シート 書込直前: recordRow[11]=' + recordRow[11] + ', [12]=' + recordRow[12] + ', [13]=' + recordRow[13]);
    Logger.log('[completeTodo] ★recordRow全体: ' + JSON.stringify(recordRow));
    targetSheet.appendRow(recordRow);
    Logger.log('[completeTodo] 記録シートに追加完了: ID=' + tv('ID') + ', 業務種別=' + category + ', タスク=' + tv('タスク内容'));
  }

  // --- 記録済に日時を入れる ---
  if (recordedCol !== -1) {
    var recordedNowJST = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    todoSheet.getRange(rowNum, recordedCol + 1).setValue(recordedNowJST);
    Logger.log('[completeTodo] TODO行' + rowNum + ' 記録済に日時セット: ' + recordedNowJST);
  }

  Logger.log('[completeTodo] 完了処理 すべて成功');
  return jsonResponse_({
    success: true,
    savedTo: targetSheetName,
    todoId: targetId,
    category: category
  });
}

// --------------- 行更新処理（既存行のステータス等を更新） ---------------
// appendRow は使わず、必ず該当IDの行を探して上書きする
function handleUpdateRow_(ss, data) {
  // シート名の指定があればそちらを使う（なければTODOシート）
  var sheetName = data.sheet || SHEET_TODO;
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return jsonResponse_({ error: 'シート「' + sheetName + '」が見つかりません' });

  var allData = sheet.getDataRange().getValues();
  var headers = allData[0];

  // 検索キー列の指定（デフォルトは 'ID'）
  var keyCol = data.updateKey || 'ID';
  var idCol = headers.indexOf(keyCol);
  if (idCol === -1) return jsonResponse_({ error: keyCol + '列が見つかりません（シート: ' + sheetName + '）' });

  var targetId = String(data[keyCol] || data.id || '');
  Logger.log('[updateRow] シート=' + sheetName + ', キー=' + keyCol + ', 対象ID=' + targetId);

  for (var i = 1; i < allData.length; i++) {
    if (String(allData[i][idCol]) === targetId) {
      var rowNum = i + 1;
      Logger.log('[updateRow] ID=' + targetId + ' → 行' + rowNum);

      // 送信されたフィールドをヘッダーに基づいて更新
      for (var j = 0; j < headers.length; j++) {
        var colName = headers[j];
        // ID・作成日・記録済は変更しない
        if (colName === keyCol || colName === '作成日' || colName === '記録済') continue;
        // 更新日・更新日時はGAS側の現在日時（JST）を強制使用
        if (colName === '更新日' || colName === '更新日時') {
          var updateNowJST = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
          sheet.getRange(rowNum, j + 1).setValue(updateNowJST);
          continue;
        }
        if (data[colName] !== undefined) {
          sheet.getRange(rowNum, j + 1).setValue(data[colName]);
        }
      }

      Logger.log('[updateRow] 行' + rowNum + ' の更新完了');
      return jsonResponse_({ success: true, updatedRow: rowNum
      });
    }
  }

  return jsonResponse_({ error: '該当ID(' + targetId + ')がシート「' + sheetName + '」に見つかりません' });
}

// --------------- 相談記録TODO完了 ---------------
function handleCompleteConsultTodo_(ss, data) {
  var sheet = ss.getSheetByName(SHEET_CONSULT);
  if (!sheet) return jsonResponse_({ success: false, error: 'シート「' + SHEET_CONSULT + '」が見つかりません' });

  var allData = sheet.getDataRange().getValues();
  var headers = allData[0];
  var idCol     = headers.indexOf('ID');
  var flagCol   = headers.indexOf('要対応フラグ');
  var statusCol = headers.indexOf('ステータス');
  var updateCol = headers.indexOf('更新日時');

  Logger.log('[completeConsultTodo] ID列=' + idCol + ', 要対応フラグ列=' + flagCol + ', ステータス列=' + statusCol);

  if (idCol === -1) return jsonResponse_({ success: false, error: 'ID列が見つかりません' });
  if (flagCol === -1) return jsonResponse_({ success: false, error: '要対応フラグ列が見つかりません' });

  var targetId = String(data.consultId || data.id || '');
  Logger.log('[completeConsultTodo] 対象ID=' + targetId);

  for (var i = 1; i < allData.length; i++) {
    if (String(allData[i][idCol]) === targetId) {
      var rowNum = i + 1;
      // O列を FALSE に
      sheet.getRange(rowNum, flagCol + 1).setValue('FALSE');
      Logger.log('[completeConsultTodo] 行' + rowNum + ' 要対応フラグ → FALSE');

      // H列(ステータス)を「完了」に
      if (statusCol >= 0) {
        sheet.getRange(rowNum, statusCol + 1).setValue('完了');
        Logger.log('[completeConsultTodo] 行' + rowNum + ' ステータス → 完了');
      }

      // K列(更新日時)を現在時刻に
      if (updateCol >= 0) {
        var nowJST = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
        sheet.getRange(rowNum, updateCol + 1).setValue(nowJST);
      }

      return jsonResponse_({ success: true, updatedRow: rowNum, sheetName: SHEET_CONSULT });
    }
  }

  return jsonResponse_({ success: false, error: '該当ID(' + targetId + ')が見つかりません' });
}

// --------------- 休日マスタ ---------------
function handleGetHolidayMaster_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_HOLIDAY);
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var val = data[i][j];
      // 日付列はyyyy-MM-dd文字列に変換
      if (headers[j] === '日付' && val instanceof Date) {
        val = Utilities.formatDate(val, 'Asia/Tokyo', 'yyyy-MM-dd');
      }
      obj[headers[j]] = val;
    }
    result.push(obj);
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// --------------- アクティビティログ ---------------

// ログ追加
function handleLogActivity_(ss, data) {
  var sheet = ss.getSheetByName(SHEET_ACTIVITY);
  if (!sheet) {
    // シートがなければ作成
    var headers = ['日時','操作種別','対象種別','クライアント名','件名','詳細','実行者','元データID','元シート','備考'];
    sheet = ss.insertSheet(SHEET_ACTIVITY);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  var nowJST = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  var row = [
    nowJST,
    data['操作種別'] || '',
    data['対象種別'] || '',
    data['クライアント名'] || '',
    data['件名'] || '',
    data['詳細'] || '',
    data['実行者'] || '',
    data['元データID'] || '',
    data['元シート'] || '',
    data['備考'] || ''
  ];
  sheet.appendRow(row);
  return jsonResponse_({ success: true });
}

// ログ取得（直近20件、新しい順）
function handleGetActivityLog_(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_ACTIVITY);
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // limit（既定20・上限200）。全件は読まず、末尾 limit 行だけ取得して負荷を抑える。
  var limit = 20;
  if (e && e.parameter && e.parameter.limit) {
    var n = parseInt(e.parameter.limit, 10);
    if (!isNaN(n) && n > 0) limit = Math.min(n, 200);
  }
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    return ContentService.createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  // データ行は 2 〜 lastRow。末尾 limit 行だけ読む。
  var dataRows = lastRow - 1;
  var take = Math.min(limit, dataRows);
  var startRow = lastRow - take + 1;
  var values = sheet.getRange(startRow, 1, take, lastCol).getValues();
  var result = [];
  // 新しい順（末尾から先頭へ）
  for (var i = values.length - 1; i >= 0; i--) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = values[i][j];
    }
    result.push(obj);
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// --------------- 算定年更管理 ---------------

// データ取得（年度フィルタ対応）
function handleGetBusySeasonRecords_(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_BUSY);
  if (!sheet) {
    // シートが無ければ空配列を返す（フロント側で初回読込時に対応）
    Logger.log('[getBusySeason] シート未作成');
    return ContentService.createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var year = (e && e.parameter && e.parameter.year) ? String(e.parameter.year) : '';
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var result = [];

  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var val = data[i][j];
      // Date型はJST文字列に変換（JSON.stringifyのUTC変換を防止）
      if (val instanceof Date) {
        val = Utilities.formatDate(val, 'Asia/Tokyo', 'yyyy/MM/dd');
      }
      obj[headers[j]] = val;
    }
    // 年度フィルタ
    if (year && String(obj['年度']) !== year) continue;
    result.push(obj);
  }

  Logger.log('[getBusySeason] 年度=' + (year || '全件') + ', 件数=' + result.length);
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// 一括保存（全行を差し替え）
function handleSaveBusySeasonRecords_(ss, data) {
  var rows = data.rows;
  if (!Array.isArray(rows)) {
    return jsonResponse_({ success: false, error: 'rows が配列ではありません' });
  }

  var sheet = ss.getSheetByName(SHEET_BUSY);
  var headers = ['年度','顧問先区分','顧問先名',
    '年度更新_資料回収','年度更新_データ作成','年度更新_申告書作成','年度更新_申告',
    '年度更新_納付書作成','年度更新_納付額通知','年度更新_公文書通知',
    '算定基礎_データ作成','算定基礎_申告書作成','算定基礎_申請',
    '算定基礎_結果取込','算定基礎_保険料通知','算定基礎_公文書通知',
    'コメント','更新日'];

  // シートが無ければ作成
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_BUSY);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    Logger.log('[saveBusySeason] シート新規作成');
  }

  var year = data.year ? String(data.year) : '';
  Logger.log('[saveBusySeason] 年度=' + year + ', 行数=' + rows.length);

  // 既存データ読み込み
  var existingData = sheet.getDataRange().getValues();
  var existingHeaders = existingData[0];
  var yearCol = existingHeaders.indexOf('年度');

  // 指定年度以外の行を保持
  var keepRows = [];
  for (var i = 1; i < existingData.length; i++) {
    if (year && String(existingData[i][yearCol]) !== year) {
      keepRows.push(existingData[i]);
    }
  }

  // 新しい行を構築
  var nowJST = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  var newRows = rows.map(function(row) {
    return headers.map(function(h) {
      if (h === '更新日') return nowJST;
      return row[h] !== undefined ? row[h] : '';
    });
  });

  // 全データ = ヘッダー + 他年度 + 今回年度
  var allRows = [headers].concat(keepRows).concat(newRows);

  // シートをクリアして書き直し
  sheet.clearContents();
  if (allRows.length > 0) {
    var range = sheet.getRange(1, 1, allRows.length, headers.length);
    // 日付列（D〜P: 4列目〜16列目）をテキスト形式にしてDate自動変換を防止
    for (var ci = 3; ci <= 15; ci++) {
      sheet.getRange(2, ci + 1, Math.max(allRows.length - 1, 1), 1).setNumberFormat('@');
    }
    range.setValues(allRows);
  }

  Logger.log('[saveBusySeason] 保存完了: 全' + (allRows.length - 1) + '行（今回' + newRows.length + '行）');
  return jsonResponse_({ success: true, savedCount: newRows.length, totalCount: allRows.length - 1 });
}

// --------------- ストレスチェック管理 ---------------

// 一括保存（事業所名をキーに upsert、算定年更管理の handleSaveBusySeasonRecords_ と同じ考え方）
function handleSaveStressCheckRecords_(ss, data) {
  var rows = data.rows;
  if (!Array.isArray(rows)) {
    return jsonResponse_({ success: false, error: 'rows が配列ではありません' });
  }

  var headers = ['実施年度','事業所名','担当者','実施月','説明日','同意取得日','開始日','終了日',
    '集団分析同意','集団分析同封','出力日','郵送日','ステータス','備考','重要フラグ','作成日','更新日'];

  var sheet = ss.getSheetByName(SHEET_STRESS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_STRESS);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    Logger.log('[saveStressCheckRecords] シート新規作成');
  }

  Logger.log('[saveStressCheckRecords] 保存行数=' + rows.length);

  // 既存データから作成日を事業所名キーで保持
  var existingData    = sheet.getDataRange().getValues();
  var existingHeaders = existingData[0];
  var jigyoshoCol     = existingHeaders.indexOf('事業所名');
  var sosakuCol       = existingHeaders.indexOf('作成日');
  var createdMap      = {};
  for (var i = 1; i < existingData.length; i++) {
    var name = String(existingData[i][jigyoshoCol] || '').trim();
    if (name) createdMap[name] = existingData[i][sosakuCol] || '';
  }

  var nowJST = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

  // 新しい行を構築（事業所名なし行はスキップ）
  var newRows = [];
  rows.forEach(function(row) {
    var name = String(row['事業所名'] || '').trim();
    if (!name) return;
    newRows.push(headers.map(function(h) {
      if (h === '更新日') return nowJST;
      if (h === '作成日') return createdMap[name] || nowJST;
      return row[h] !== undefined ? row[h] : '';
    }));
  });

  // シートをクリアして書き直し（ヘッダー + データ）
  sheet.clearContents();
  var allRows = [headers].concat(newRows);
  if (allRows.length > 0) {
    sheet.getRange(1, 1, allRows.length, headers.length).setValues(allRows);
  }

  Logger.log('[saveStressCheckRecords] 保存完了: ' + newRows.length + '行');
  return jsonResponse_({ success: true, savedCount: newRows.length });
}

// データ取得（全件）
function handleGetStressCheckRecords_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_STRESS);
  if (!sheet) {
    Logger.log('[getStressCheck] シート未作成');
    return ContentService.createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return ContentService.createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var headers = data[0];
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var val = data[i][j];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, 'Asia/Tokyo', 'yyyy-MM-dd');
      }
      obj[headers[j]] = val;
    }
    result.push(obj);
  }
  Logger.log('[getStressCheck] 件数: ' + result.length);
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// 新規追加（ID・作成日・更新日を自動セット）
function handleSaveStressCheckRecord_(ss, data) {
  var sheet = ss.getSheetByName(SHEET_STRESS);
  if (!sheet) {
    return jsonResponse_({ success: false, error: 'シート「' + SHEET_STRESS + '」が見つかりません' });
  }
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var nowJST = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

  data['ID'] = getNextId_(sheet, headers);
  data['作成日'] = nowJST;
  data['更新日'] = nowJST;

  var row = headers.map(function(h) {
    return data[h] !== undefined ? data[h] : '';
  });
  sheet.appendRow(row);
  Logger.log('[saveStressCheck] 新規追加: ID=' + data['ID'] + ', 事業所=' + data['事業所名']);
  return jsonResponse_({ success: true, id: data['ID'] });
}

// 更新（ID指定で1行更新、更新日はGAS側で自動セット）
// 重要フラグ専用更新（ID行を特定して「重要フラグ」列だけを書き込む）
function handleUpdateStressCheckFlag_(ss, data) {
  var action    = data.action || '';
  var id        = String(data['ID'] || '');
  var flagValue = String(data['重要フラグ'] || '');

  Logger.log('[updateStressCheckFlag] action=' + action + ', ID=' + id + ', 重要フラグ=' + flagValue);

  if (!id) {
    Logger.log('[updateStressCheckFlag] ❌ IDが指定されていません');
    return jsonResponse_({ success: false, message: 'IDが指定されていません' });
  }
  if (flagValue !== 'TRUE' && flagValue !== 'FALSE') {
    Logger.log('[updateStressCheckFlag] ❌ 重要フラグの値が不正: ' + flagValue);
    return jsonResponse_({ success: false, message: '重要フラグの値は TRUE または FALSE である必要があります（受信値: "' + flagValue + '"）' });
  }

  var sheet = ss.getSheetByName(SHEET_STRESS);
  if (!sheet) {
    Logger.log('[updateStressCheckFlag] ❌ シートが見つかりません: ' + SHEET_STRESS);
    return jsonResponse_({ success: false, message: 'シート「' + SHEET_STRESS + '」が見つかりません' });
  }

  var allData = sheet.getDataRange().getValues();
  var headers  = allData[0];

  var idCol   = headers.indexOf('ID');
  var flagCol = headers.indexOf('重要フラグ');

  Logger.log('[updateStressCheckFlag] ID列index=' + idCol + ', 重要フラグ列index=' + flagCol);
  Logger.log('[updateStressCheckFlag] ヘッダー一覧: ' + JSON.stringify(headers));

  if (idCol === -1) {
    return jsonResponse_({ success: false, message: 'シートに「ID」列が見つかりません' });
  }
  if (flagCol === -1) {
    return jsonResponse_({ success: false, message: 'シート「' + SHEET_STRESS + '」に「重要フラグ」列が見つかりません。ヘッダー行に「重要フラグ」という列を追加してください。現在のヘッダー: ' + headers.join(', ') });
  }

  for (var i = 1; i < allData.length; i++) {
    if (String(allData[i][idCol]) === id) {
      var rowNum = i + 1;
      Logger.log('[updateStressCheckFlag] 書き込み対象行=' + rowNum + ', 列=' + (flagCol + 1) + ', 値=' + flagValue);
      sheet.getRange(rowNum, flagCol + 1).setValue(flagValue);
      Logger.log('[updateStressCheckFlag] ✅ 書き込み完了: 行' + rowNum + ' の重要フラグ → ' + flagValue);
      return jsonResponse_({ success: true, updatedRow: rowNum, flagValue: flagValue });
    }
  }

  Logger.log('[updateStressCheckFlag] ❌ 対象IDが見つかりません: ' + id);
  return jsonResponse_({ success: false, message: '対象ID(' + id + ')が「' + SHEET_STRESS + '」シートに見つかりません' });
}

function handleUpdateStressCheckRecord_(ss, data) {
  var sheet = ss.getSheetByName(SHEET_STRESS);
  if (!sheet) {
    return jsonResponse_({ success: false, error: 'シート「' + SHEET_STRESS + '」が見つかりません' });
  }
  var allData = sheet.getDataRange().getValues();
  var headers = allData[0];
  var idCol = headers.indexOf('ID');
  if (idCol === -1) {
    return jsonResponse_({ success: false, error: 'ID列が見つかりません' });
  }
  // 「重要フラグ」を含む更新の場合、列が存在するか先に確認
  if (data['重要フラグ'] !== undefined && headers.indexOf('重要フラグ') === -1) {
    Logger.log('[updateStressCheck] ❌ 「重要フラグ」列が見つかりません。ヘッダー: ' + JSON.stringify(headers));
    return jsonResponse_({ success: false, error: 'スプレッドシートの「ストレスチェック管理」シートに「重要フラグ」列がありません。シートのヘッダー行に「重要フラグ」列を追加してください。' });
  }

  var targetId = String(data['ID'] || '');
  for (var i = 1; i < allData.length; i++) {
    if (String(allData[i][idCol]) === targetId) {
      var rowNum = i + 1;
      var nowJST = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
      for (var j = 0; j < headers.length; j++) {
        var colName = headers[j];
        if (colName === 'ID' || colName === '作成日') continue;
        if (colName === '更新日') {
          sheet.getRange(rowNum, j + 1).setValue(nowJST);
          continue;
        }
        if (data[colName] !== undefined) {
          sheet.getRange(rowNum, j + 1).setValue(data[colName]);
        }
      }
      Logger.log('[updateStressCheck] 行' + rowNum + ' 更新完了: ID=' + targetId);
      return jsonResponse_({ success: true, updatedRow: rowNum });
    }
  }
  return jsonResponse_({ success: false, error: '該当ID(' + targetId + ')が見つかりません' });
}

// ===============================================================
// 顧問料管理
// ===============================================================

// --------------- 月別請求管理 取得 ---------------
// ?action=getFeeMonthly&month=YYYY/MM で対象月のみ / month省略で全件
function handleGetFeeMonthly_(e) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_FEE_MONTHLY);
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var month = (e && e.parameter && e.parameter.month) ? String(e.parameter.month).trim() : '';
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return ContentService.createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var headers = data[0];
  var result  = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j];
    }
    if (!month || String(obj['対象月'] || '').trim() === month) {
      result.push(obj);
    }
  }
  Logger.log('[getFeeMonthly] month=' + (month || '全件') + ' 件数=' + result.length);
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// --------------- 顧問料設定 一括保存（client_id キー upsert） ---------------
function handleSaveFeeMaster_(ss, data) {
  var rows = data.rows;
  if (!Array.isArray(rows)) {
    return jsonResponse_({ success: false, error: 'rows が配列ではありません' });
  }
  var headers = ['client_id','会社名','担当者','月額顧問料','給与計算費','社会保険費',
    '支払方法','源泉対象','消費税対象','備考','有効フラグ','作成日時','更新日時'];

  var sheet = ss.getSheetByName(SHEET_FEE_MASTER);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_FEE_MASTER);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    Logger.log('[saveFeeMaster] シート新規作成');
  }

  var existingData = sheet.getDataRange().getValues();
  var existHeaders = existingData[0];
  var idCol      = existHeaders.indexOf('client_id');
  var createdCol = existHeaders.indexOf('作成日時');
  var createdMap = {};
  for (var i = 1; i < existingData.length; i++) {
    var cid = String(existingData[i][idCol] || '').trim();
    if (cid) createdMap[cid] = existingData[i][createdCol] || '';
  }

  var nowJST = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  var newRows = [];
  rows.forEach(function(row) {
    var cid = String(row['client_id'] || '').trim();
    if (!cid) return;
    newRows.push(headers.map(function(h) {
      if (h === '更新日時') return nowJST;
      if (h === '作成日時') return createdMap[cid] || nowJST;
      return row[h] !== undefined ? row[h] : '';
    }));
  });

  sheet.clearContents();
  var allRows = [headers].concat(newRows);
  sheet.getRange(1, 1, allRows.length, headers.length).setValues(allRows);
  Logger.log('[saveFeeMaster] 保存完了: ' + newRows.length + '行');
  return jsonResponse_({ success: true, savedCount: newRows.length });
}

// --------------- 月別請求管理 一括保存（対象月内を全置換、他月は保持） ---------------
// GAS側でも税額を再計算して保存（JS側との不一致防止）
// 源泉対象額 = 顧問料 + 給与計算費 + スポット請求 + 調整額（社会保険費は除外）
function handleSaveFeeMonthly_(ss, data) {
  var rows = data.rows;
  if (!Array.isArray(rows)) {
    return jsonResponse_({ success: false, error: 'rows が配列ではありません' });
  }
  var headers = ['対象月','client_id','会社名','担当者','顧問料','給与計算費','社会保険費',
    'スポット請求','スポット内容','調整額','調整理由','小計','消費税','源泉所得税',
    '入金予定額','源泉対象','消費税対象','支払方法','入金状況','入金日','メモ','作成日時','更新日時'];

  var sheet = ss.getSheetByName(SHEET_FEE_MONTHLY);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_FEE_MONTHLY);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    Logger.log('[saveFeeMonthly] シート新規作成');
  }

  var targetMonths = {};
  rows.forEach(function(r) {
    var m = String(r['対象月'] || '').trim();
    if (m) targetMonths[m] = true;
  });

  var existingData = sheet.getDataRange().getValues();
  var existHeaders = existingData[0];
  var monthColE   = existHeaders.indexOf('対象月');
  var cidColE     = existHeaders.indexOf('client_id');
  var createdColE = existHeaders.indexOf('作成日時');

  var createdMap = {};
  var remainRows = [];
  for (var i = 1; i < existingData.length; i++) {
    var rowM   = String(existingData[i][monthColE]  || '').trim();
    var rowCid = String(existingData[i][cidColE] || '').trim();
    var key    = rowM + '_' + rowCid;
    createdMap[key] = existingData[i][createdColE] || '';
    if (!targetMonths[rowM]) {
      var obj = {};
      for (var j = 0; j < existHeaders.length; j++) {
        obj[existHeaders[j]] = existingData[i][j];
      }
      remainRows.push(headers.map(function(h) {
        return obj[h] !== undefined ? obj[h] : '';
      }));
    }
  }

  var nowJST  = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  var newRows = [];
  rows.forEach(function(row) {
    var targetMonth = String(row['対象月'] || '').trim();
    var cid         = String(row['client_id'] || '').trim();
    if (!targetMonth) return;
    var key = targetMonth + '_' + cid;

    var fee     = parseFloat(row['顧問料'])      || 0;
    var payroll = parseFloat(row['給与計算費'])  || 0;
    var social  = parseFloat(row['社会保険費'])  || 0;
    var spot    = parseFloat(row['スポット請求']) || 0;
    var adjust  = parseFloat(row['調整額'])      || 0;

    var subtotal = fee + payroll + social + spot + adjust;
    var isTax    = String(row['消費税対象'] || '').toUpperCase() === 'TRUE';
    var isWh     = String(row['源泉対象']   || '').toUpperCase() === 'TRUE';
    var tax      = isTax ? Math.floor(subtotal * 0.1) : 0;
    var whTarget = fee + payroll + spot + adjust;  // 社会保険費は源泉対象外
    var wh       = isWh  ? Math.floor(whTarget * 0.1021) : 0;
    var expected = subtotal + tax - wh;

    newRows.push(headers.map(function(h) {
      if (h === '更新日時')   return nowJST;
      if (h === '作成日時')   return createdMap[key] || nowJST;
      if (h === '小計')       return subtotal;
      if (h === '消費税')     return tax;
      if (h === '源泉所得税') return wh;
      if (h === '入金予定額') return expected;
      return row[h] !== undefined ? row[h] : '';
    }));
  });

  var allRows = [headers].concat(remainRows).concat(newRows);
  sheet.clearContents();
  sheet.getRange(1, 1, allRows.length, headers.length).setValues(allRows);
  Logger.log('[saveFeeMonthly] 保存完了: 新規/更新=' + newRows.length + '行, 保持=' + remainRows.length + '行');
  return jsonResponse_({ success: true, savedCount: newRows.length });
}

// --------------- 担当者状態（一時保存） ---------------

var VALID_MEMBER_STATUSES = ['normal', 'phone', 'toilet', 'smoke', 'clean', 'meal'];

function handleGetMemberStatuses_() {
  var cache = CacheService.getScriptCache();
  var raw = cache.get('memberStatuses');
  var statuses = raw ? JSON.parse(raw) : {};
  return jsonResponse_(statuses);
}

function handleSetMemberStatus_(data) {
  var name = data.name || '';
  var status = data.status || 'normal';
  if (VALID_MEMBER_STATUSES.indexOf(status) === -1) status = 'normal';
  if (!name) return jsonResponse_({ success: false, error: 'name is required' });

  var cache = CacheService.getScriptCache();
  var raw = cache.get('memberStatuses');
  var statuses = raw ? JSON.parse(raw) : {};
  var nowJST = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  statuses[name] = { status: status, updatedAt: nowJST };
  // CacheService最大6時間（21600秒）
  cache.put('memberStatuses', JSON.stringify(statuses), 21600);
  return jsonResponse_({ success: true, name: name, status: status, updatedAt: nowJST });
}

// --------------- 一言メッセージ（スプレッドシート保存） ---------------

function handleGetMessages_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_MESSAGE);
  if (!sheet) return jsonResponse_([]);

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return jsonResponse_([]); // ヘッダーのみ

  var messages = [];
  for (var i = 1; i < data.length; i++) {
    var ts = data[i][1];
    if (ts instanceof Date) {
      ts = Utilities.formatDate(ts, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
    }
    messages.push({
      id: String(data[i][0]),
      ts: String(ts || ''),
      from: String(data[i][2] || ''),
      to: String(data[i][3] || ''),
      text: String(data[i][4] || '')
    });
  }
  // 新しい順にソートして最新3件
  messages.sort(function(a, b) { return a.ts > b.ts ? -1 : a.ts < b.ts ? 1 : 0; });
  return jsonResponse_(messages.slice(0, 3));
}

function handlePostMessage_(data) {
  var text = (data.text || '').trim();
  if (!text) return jsonResponse_({ success: false, error: 'text is required' });
  var from = data.from || '';
  var to = data.to || '';
  if (!from || !to) return jsonResponse_({ success: false, error: 'from and to are required' });

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_MESSAGE);
  if (!sheet) return jsonResponse_({ success: false, error: 'シートが見つかりません' });

  var id = String(new Date().getTime());
  var nowJST = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
  sheet.appendRow([id, nowJST, from, to, text]);

  return jsonResponse_({ success: true, message: { id: id, ts: nowJST, from: from, to: to, text: text } });
}

function handleDeleteMessage_(data) {
  var id = String(data.id || '');
  if (!id) return jsonResponse_({ success: false, error: 'id is required' });

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_MESSAGE);
  if (!sheet) return jsonResponse_({ success: false, error: 'シートが見つかりません' });

  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === id) {
      sheet.deleteRow(i + 1);
      return jsonResponse_({ success: true });
    }
  }
  return jsonResponse_({ success: false, error: 'not found' });
}

// --------------- ユーティリティ ---------------

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getNextId_(sheet, headers) {
  var idCol = headers.indexOf('ID');
  if (idCol === -1) return 1;
  var data = sheet.getDataRange().getValues();
  var maxId = 0;
  for (var i = 1; i < data.length; i++) {
    var v = parseInt(data[i][idCol], 10);
    if (v > maxId) maxId = v;
  }
  return maxId + 1;
}

// 指定カラム名でID採番（クライアント・行政問い合わせ用）
function getNextIdByCol_(sheet, headers, colName) {
  var idCol = headers.indexOf(colName);
  if (idCol === -1) return 1;
  var data = sheet.getDataRange().getValues();
  var maxId = 0;
  for (var i = 1; i < data.length; i++) {
    var v = parseInt(data[i][idCol], 10);
    if (v > maxId) maxId = v;
  }
  return maxId + 1;
}

function convertToMinutes_(val) {
  if (val === '' || val === null || val === undefined) return '';

  if (typeof val === 'number') {
    if (val < 1 && val > 0) {
      return Math.round(val * 24 * 60);
    }
    return val;
  }

  if (val instanceof Date) {
    return val.getHours() * 60 + val.getMinutes();
  }

  var str = String(val);
  var match = str.match(/^(\d+):(\d+)$/);
  if (match) {
    return parseInt(match[1], 10) * 60 + parseInt(match[2], 10);
  }

  var num = parseInt(str, 10);
  if (!isNaN(num)) return num;

  return '';
}

// ===============================================================
// 勤務予定共有
// ===============================================================

function ensureWorkAvailabilitySheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_WORK_AVAIL);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_WORK_AVAIL);
    sheet.appendRow(['id', 'staff_id', 'staff_name', 'date', 'status', 'start_time', 'end_time', 'memo', 'updated_at']);
    Logger.log('[ensureWorkAvailabilitySheet_] シート「' + SHEET_WORK_AVAIL + '」を新規作成しました');
  }
  return sheet;
}

function handleGetWorkAvailability_(e) {
  Logger.log('===== handleGetWorkAvailability_ NEW CODE v76 =====');
  var sheet   = ensureWorkAvailabilitySheet_();
  var staffId = (e && e.parameter && e.parameter.staff_id) ? e.parameter.staff_id : '';
  var month   = (e && e.parameter && e.parameter.month)    ? e.parameter.month    : '';
  Logger.log('[getWorkAvailability] staff_id=' + staffId + ' month=' + month);

  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  Logger.log('[getWorkAvailability] 全行数=' + (data.length - 1));
  var timeColumns = ['start_time', 'end_time'];
  // 日付をキーに1件ずつ集約（時間入り行を優先）
  var resultMap = {};
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var v = data[i][j];
      // 先頭3行のみ型ログ（詳細確認用）
      if (i <= 3 && (headers[j] === 'start_time' || headers[j] === 'end_time')) {
        Logger.log('[WA raw] row' + i + ' col=' + headers[j] + ' value=' + v + ' type=' + Object.prototype.toString.call(v) + ' isDate=' + (v instanceof Date));
      }
      if (v instanceof Date) {
        var fmt = (timeColumns.indexOf(headers[j]) !== -1) ? 'HH:mm' : 'yyyy-MM-dd';
        v = Utilities.formatDate(v, 'Asia/Tokyo', fmt);
        if (i <= 3 && (headers[j] === 'start_time' || headers[j] === 'end_time')) {
          Logger.log('[WA fmt]  row' + i + ' col=' + headers[j] + ' formatted=' + v);
        }
      }
      obj[headers[j]] = (v === null || v === undefined) ? '' : v;
    }
    var rowStaffId = String(obj['staff_id'] || '');
    var rowDate    = String(obj['date']     || '');
    if (staffId && rowStaffId !== staffId) continue;
    if (month   && rowDate.substring(0, 7) !== month) continue;
    if (!rowDate) continue;
    var existing = resultMap[rowDate];
    if (!existing) {
      resultMap[rowDate] = obj;
    } else {
      // 時間入りの行を優先し、両方同条件なら updated_at が新しい方を採用
      var newHasTime  = hasValidTime_(obj['start_time'])      || hasValidTime_(obj['end_time']);
      var exstHasTime = hasValidTime_(existing['start_time']) || hasValidTime_(existing['end_time']);
      if (newHasTime && !exstHasTime) {
        resultMap[rowDate] = obj;
      } else if (newHasTime === exstHasTime) {
        if (String(obj['updated_at']) > String(existing['updated_at'])) resultMap[rowDate] = obj;
      }
      // !newHasTime && exstHasTime → 既存を維持
    }
  }
  var result = Object.keys(resultMap).map(function(k) { return resultMap[k]; });
  Logger.log('[getWorkAvailability] 重複除外後返却件数=' + result.length);
  if (result.length > 0) Logger.log('[getWorkAvailability] 先頭row=' + JSON.stringify(result[0]));
  return jsonResponse_(result);
}

// 「時間データが何かある」判定：空欄かどうかだけ見る
// "1899-12-30"（時刻セルが日付で返ってきた値）も "09:30" と同様に「あり」とみなす
function hasValidTime_(val) {
  return !!(val && String(val).trim() !== '');
}

function handleSaveWorkAvailability_(ss, data) {
  var sheet     = ensureWorkAvailabilitySheet_();
  var staffId   = String(data.staff_id   || '');
  var staffName = String(data.staff_name || '');
  var records   = data.records || [];
  var nowJST    = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  var baseTs    = nowJST.replace(/[^0-9]/g, '');
  Logger.log('[saveWorkAvailability] staff_id=' + staffId + ' 件数=' + records.length);

  var allData    = sheet.getDataRange().getValues();
  var headers    = allData[0];
  var idCol         = headers.indexOf('id');
  var dateCol       = headers.indexOf('date');
  var staffIdCol    = headers.indexOf('staff_id');
  var startTimeCol  = headers.indexOf('start_time') + 1; // 1-indexed (0→未見つかり)
  var endTimeCol    = headers.indexOf('end_time')   + 1;

  // 既存行を (staff_id + '_' + date) → シート行番号 でマップ化
  // date列がDate型の場合は yyyy-MM-dd に正規化してキーを作る
  var existMap = {};
  for (var i = 1; i < allData.length; i++) {
    var dv = allData[i][dateCol];
    if (dv instanceof Date) dv = Utilities.formatDate(dv, 'Asia/Tokyo', 'yyyy-MM-dd');
    var key = String(allData[i][staffIdCol]) + '_' + String(dv);
    existMap[key] = i + 1;
  }

  var count = 0;
  for (var k = 0; k < records.length; k++) {
    var rec     = records[k];
    var dateStr = String(rec.date || '');
    if (!dateStr) continue;

    var rowKey  = staffId + '_' + dateStr;
    var startVal = String(rec.start_time || '');
    var endVal   = String(rec.end_time   || '');
    Logger.log('[saveWorkAvailability] key=' + rowKey + ' start=' + startVal + ' end=' + endVal);
    var rowData = [
      '',                             // A: id
      staffId,                        // B: staff_id
      staffName,                      // C: staff_name
      dateStr,                        // D: date
      String(rec.status     || ''),   // E: status
      startVal,                       // F: start_time
      endVal,                         // G: end_time
      String(rec.memo       || ''),   // H: memo
      nowJST                          // I: updated_at
    ];

    if (existMap[rowKey]) {
      var rowNum  = existMap[rowKey];
      rowData[idCol] = allData[rowNum - 1][idCol]; // id は変えない
      // start_time / end_time セルをテキスト形式にしてから保存（Sheetsの自動変換を防ぐ）
      if (startTimeCol > 0) sheet.getRange(rowNum, startTimeCol).setNumberFormat('@');
      if (endTimeCol   > 0) sheet.getRange(rowNum, endTimeCol).setNumberFormat('@');
      sheet.getRange(rowNum, 1, 1, headers.length).setValues([rowData]);
    } else {
      rowData[idCol] = baseTs + ('0' + k).slice(-2);
      // 新規行：テキスト形式指定後に書き込む（appendRow はフォーマット指定不可のため setValues を使用）
      var newRow = sheet.getLastRow() + 1;
      if (startTimeCol > 0) sheet.getRange(newRow, startTimeCol).setNumberFormat('@');
      if (endTimeCol   > 0) sheet.getRange(newRow, endTimeCol).setNumberFormat('@');
      sheet.getRange(newRow, 1, 1, headers.length).setValues([rowData]);
      existMap[rowKey] = newRow;
    }
    count++;
  }

  Logger.log('[saveWorkAvailability] 完了 count=' + count);
  return jsonResponse_({ success: true, count: count });
}

// ============================================================
// メンテナンス: 実タブ名の確認用ヘルパー
// ------------------------------------------------------------
// このツールのコード上は完了データ専用タブ（TODO（完了）/記録（完了））は
// 存在せず、実体は「TODO」「記録」シート内に完了行が混在している。
// 本削除の前に、まずこの関数を実行して実際のタブ名を確認すること。
// ============================================================
function maintenanceListSheetNames_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const names = ss.getSheets().map(s => s.getName());
  console.log('========== 現在のタブ名一覧 ==========');
  names.forEach((n, i) => console.log(`${i + 1}. ${n}`));
  return names;
}

// ============================================================
// メンテナンス: 古い履歴・完了データ行の削除
// ------------------------------------------------------------
// スプレッドシートが重くなるのを防ぐため、指定日(2026/06/30)以前の
// 「古い完了済み・履歴系データ」だけを行ごと削除する。
// 未完了TODOや現在使用中の記録は削除しない。
//
// 【重要 / 完了判定について】
//   実タブ名がコード定義（SHEET_TODO='TODO' / SHEET_RECORD='記録'）の場合、
//   完了データと未完了データが同じシートに混在している。そのため：
//     - TODO系シート: 状態=「完了」または記録済フラグがある行のみ削除対象
//     - 記録系シート: 完了区分(状態)列がある場合は「完了」行のみ削除対象
//     - 給与計算記録 / アクティビティログ: 全行が履歴のため日付条件のみで削除
//   完了判定列が必須のシート(TODO系)で完了列が見つからない場合は、
//   未完了誤削除を避けるためシートごとスキップする。
//
// 【使い方】
//   0. 先に maintenanceListSheetNames_() を実行して実タブ名を確認
//   1. DRY_RUN = true のまま実行し、ログで削除対象を確認
//   2. 問題なければ DRY_RUN = false に変更して再実行（本削除）
//   3. 削除後、アプリを再読み込みして空欄TODOが発生しないか確認
//
// 【安全対策】
//   - 本削除の前に対象シートをバックアップコピー
//   - 行番号ズレ防止のため必ず「下の行から」削除（deleteRows で一括）
//   - 値クリアではなく deleteRows で行自体を削除（空行・書式残りを防ぐ）
//   - 日付列・完了列はヘッダー名から自動判定
// ============================================================
function maintenanceDeleteOldRowsUntil20260630() {
  const DRY_RUN = true; // 確認後 false に変更
  const CUTOFF = new Date('2026-06-30T23:59:59+09:00');

  // 対象シートの設定。
  //   names          : 実タブ名の候補（先に見つかったものを使用。完了専用タブがあれば優先）
  //   completionMode : 'required'  … 完了判定必須。完了列が無ければシートごとスキップ（TODO系）
  //                    'ifPresent' … 完了列があれば完了行のみ、無ければ日付のみ（記録系）
  //                    'none'      … 完了判定なし・日付条件のみ（純粋な履歴ログ）
  const TARGET_CONFIGS = [
    // --- 第1段階：判定列が明確な履歴ログのみ削除（全行が過去データ・日付条件のみ） ---
    { names: ['給与計算記録'],               completionMode: 'none'      },
    { names: ['アクティビティログ'],         completionMode: 'none'      }

    // --- 第2段階：TODO / 記録 は実タブ名・完了列・削除方針を確認できるまで対象外 ---
    //   ※ TODO は状態=「完了」または記録済フラグのある行のみ削除（未完了は絶対に削除しない）
    //   ※ 記録 は「完了行のみ」か「古い記録全体か」を確認するまで本削除しない
    //   確認後、下記コメントを解除して再度 DRY_RUN で件数を確認すること。
    // , { names: ['TODO（完了）', 'TODO'],    completionMode: 'required'  }
    // , { names: ['記録（完了）', '記録'],     completionMode: 'ifPresent' }
  ];

  const DATE_HEADER_CANDIDATES = [
    '日付', '期限', '完了日', '対応日', '作成日', '日時',
    'createdAt', 'updatedAt', 'date', 'timestamp'
  ];

  // 完了状態を示す「状態/区分」列の候補
  const STATUS_HEADER_CANDIDATES = ['状態', 'ステータス', '完了区分', '完了フラグ', 'status'];
  // 値が入っていれば完了とみなす「フラグ」列の候補（記録済など）
  const FLAG_HEADER_CANDIDATES = ['記録済', '完了日時'];
  // 状態列がこれらの値なら「完了」とみなす
  const COMPLETED_VALUES = ['完了', '済', '完了済', '対応済', 'done', 'true'];

  const isCompletedStatus = (v) => {
    if (v === true) return true;
    const s = String(v).trim().toLowerCase();
    if (!s) return false;
    return COMPLETED_VALUES.indexOf(s) !== -1;
  };

  const findHeaderIndex = (headers, candidates) => {
    for (const c of candidates) {
      const idx = headers.indexOf(c);
      if (idx !== -1) return idx;
    }
    return -1;
  };

  // 削除対象行の内容プレビュー（先頭最大6列）を作る
  const buildPreview = (headers, row) => {
    const obj = {};
    const limit = Math.min(headers.length, 6);
    for (let c = 0; c < limit; c++) {
      const key = headers[c] || ('列' + (c + 1));
      let v = row[c];
      if (Object.prototype.toString.call(v) === '[object Date]') {
        v = Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
      }
      obj[key] = v;
    }
    return obj;
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const backupSuffix = Utilities.formatDate(new Date(), tz, 'yyyyMMdd_HHmmss');

  TARGET_CONFIGS.forEach(config => {
    // --- 実タブ名を解決（候補の先頭から存在するものを採用） ---
    let sheet = null;
    let sheetName = null;
    for (const name of config.names) {
      const s = ss.getSheetByName(name);
      if (s) { sheet = s; sheetName = name; break; }
    }
    if (!sheet) {
      console.warn(`[SKIP] シートが見つかりません: 候補=${JSON.stringify(config.names)}`);
      return;
    }

    console.log(`========== ${sheetName} ==========`);

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow <= 1) {
      console.log(`[SKIP] データ行なし: ${sheetName}`);
      return;
    }

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());

    // --- 日付列の判定 ---
    const dateColIndex = findHeaderIndex(headers, DATE_HEADER_CANDIDATES);
    if (dateColIndex === -1) {
      console.warn(`[SKIP] 日付列が見つかりません: ${sheetName} / headers=${JSON.stringify(headers)}`);
      return;
    }

    // --- 完了判定列の判定 ---
    const statusColIndex = findHeaderIndex(headers, STATUS_HEADER_CANDIDATES);
    const flagColIndex = findHeaderIndex(headers, FLAG_HEADER_CANDIDATES);
    const hasCompletionCol = statusColIndex !== -1 || flagColIndex !== -1;

    // 完了判定が必須のシートで完了列が無ければ、未完了誤削除を避けてスキップ
    if (config.completionMode === 'required' && !hasCompletionCol) {
      console.warn(`[SKIP] 完了判定列が見つからないため削除しません（未完了誤削除防止）: ${sheetName} / headers=${JSON.stringify(headers)}`);
      return;
    }

    // このシートで完了フィルタを適用するか
    const applyCompletionFilter =
      config.completionMode === 'required' ||
      (config.completionMode === 'ifPresent' && hasCompletionCol);

    // ログ用：完了判定に使う列名
    const completionColNames = [];
    if (applyCompletionFilter) {
      if (statusColIndex !== -1) completionColNames.push(headers[statusColIndex] + '(状態)');
      if (flagColIndex !== -1)   completionColNames.push(headers[flagColIndex] + '(フラグ)');
    }

    const isRowCompleted = (row) => {
      if (statusColIndex !== -1 && isCompletedStatus(row[statusColIndex])) return true;
      if (flagColIndex !== -1) {
        const fv = row[flagColIndex];
        if (fv !== '' && fv !== null && fv !== undefined) return true;
      }
      return false;
    };

    const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const rowsToDelete = [];
    const previews = [];
    let skippedIncomplete = 0; // 日付は該当するが未完了のためスキップした件数

    values.forEach((row, i) => {
      const actualRowNumber = i + 2;
      const parsedDate = parseMaintenanceDate_(row[dateColIndex]);
      if (!parsedDate) return;
      if (parsedDate.getTime() > CUTOFF.getTime()) return;

      // 日付条件は満たす。完了フィルタが有効なら未完了行を除外
      if (applyCompletionFilter && !isRowCompleted(row)) {
        skippedIncomplete++;
        return;
      }

      rowsToDelete.push(actualRowNumber);
      if (previews.length < 5) {
        previews.push({ 行: actualRowNumber, 内容: buildPreview(headers, row) });
      }
    });

    // --- DRY_RUN ログ出力 ---
    console.log(`対象シート名: ${sheetName}`);
    console.log(`判定に使った日付列: ${headers[dateColIndex]} (${dateColIndex + 1}列目)`);
    console.log(`完了判定に使った列: ${applyCompletionFilter ? (completionColNames.join(', ') || '(なし)') : '(なし・日付条件のみ)'}`);
    if (applyCompletionFilter) {
      console.log(`未完了のため除外した件数: ${skippedIncomplete}`);
    }
    console.log(`削除対象件数: ${rowsToDelete.length}`);
    console.log(`削除対象行: ${rowsToDelete.join(', ')}`);
    console.log(`削除対象プレビュー(先頭最大5件): ${JSON.stringify(previews, null, 2)}`);

    if (rowsToDelete.length === 0) {
      console.log(`[SKIP] 削除対象なし: ${sheetName}`);
      return;
    }

    if (DRY_RUN) {
      console.log(`[DRY_RUN] ${sheetName}: 実削除は行いません`);
      return;
    }

    // --- バックアップ作成 ---
    const backupName = `${sheetName}_backup_${backupSuffix}`;
    sheet.copyTo(ss).setName(backupName);
    console.log(`[BACKUP] ${backupName} を作成しました`);

    // --- 連続行はまとめて、必ず「下の行から」削除する ---
    deleteRowsFromBottom_(sheet, rowsToDelete);

    console.log(`[DONE] ${sheetName}: ${rowsToDelete.length} 行を削除しました`);
  });
}

// 日付セルの値を Date に正規化する（Dateオブジェクト・各種文字列表記に対応）
function parseMaintenanceDate_(value) {
  if (!value) return null;

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return value;
  }

  const text = String(value).trim();
  if (!text) return null;

  // 2026/06/30, 2026-06-30, 2026.06.30, 2026年6月30日 などに対応
  const normalized = text
    .replace(/[年月]/g, '/')
    .replace(/[日]/g, '')
    .replace(/[.]/g, '/')
    .replace(/-/g, '/');

  const match = normalized.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})/);
  if (!match) return null;

  const y = Number(match[1]);
  const m = Number(match[2]);
  const d = Number(match[3]);

  if (!y || !m || !d) return null;

  return new Date(y, m - 1, d, 23, 59, 59);
}

// 行番号配列を「下から」連続範囲ごとにまとめて削除する
function deleteRowsFromBottom_(sheet, rows) {
  const sorted = [...rows].sort((a, b) => b - a);

  let start = sorted[0];
  let count = 1;

  for (let i = 1; i < sorted.length; i++) {
    const current = sorted[i];
    const previous = sorted[i - 1];

    if (current === previous - 1) {
      count++;
    } else {
      sheet.deleteRows(start - count + 1, count);
      start = current;
      count = 1;
    }
  }

  sheet.deleteRows(start - count + 1, count);
}


// ================================================================
// ================================================================
// 【引き渡し前・一回限りの初期化セクション ここから】
// ----------------------------------------------------------------
//  ★ 用途：複製先（社労士法人 えみふる）スプレッドに残っている
//          「濱田事務所の運用データ」を全消去し、先方が最初から
//          入力できる状態にするための、当方専用・一回限りのツール。
//  ★ 先方が日常的に使う関数ではない。
//  ★ 引き渡し時には、この「ここから 〜 ここまで」セクションごと削除する。
//
//  【三重の安全ロック】本番削除は次の3つを「すべて」満たす時だけ実行:
//     1) DRY_RUN === false
//     2) CONFIRM_INITIALIZE === 'INITIALIZE_NEW_OFFICE'
//     3) 現在のスプレッドID === TARGET_SPREADSHEET_ID（複製先IDのみ）
//     → 元の濱田事務所用スプレッドでは（IDが一致しないため）絶対に実行不可。
//
//  【実行手順】
//     手順0. listAllSheetNames_() を実行 → 実際のタブ名をログで確認し、
//            下の CLEAR_SHEETS / KEEP_SHEETS と相違がないか突き合わせる。
//     手順1. initializeSpreadsheetForNewOffice を DRY_RUN=true のまま実行
//            → 対象シートと削除予定件数がログに出るだけ（データは消えない）。
//     手順2. 問題なければ関数内の3つの定数を本番設定にして再実行:
//              const DRY_RUN = false;
//              const CONFIRM_INITIALIZE = 'INITIALIZE_NEW_OFFICE';
//            （TARGET_SPREADSHEET_ID は変更不要）
//            → 実行時にスプレッド全体のバックアップコピーを自動作成 → クリア。
//            ※ 初回はDrive承認ダイアログが出る（バックアップ作成のため）。
//     手順3. 実行後、削除件数・対象シートのログを確認し、アプリを再読込して
//            各画面が空の状態で正常表示されるか確認する。
// ================================================================

// 複製先（社労士法人 えみふる）スプレッドシートID。ここ以外では実行させない。
var INIT_TARGET_SPREADSHEET_ID = '1dgr3_BUMAvYxyWZutfkyjaysFAS1FqHlWwp-v-Azl6k';

// クリア対象（＝濱田事務所の運用データが入っている）13シートのホワイトリスト。
// ここに載っているシートの「2行目以降」のみクリアする。
var INIT_CLEAR_SHEETS = [
  'TODO',               // タスク（未完了・完了が混在）
  '記録',               // 完了タスクの記録
  '給与計算記録',       // 給与計算の履歴
  'クライアント',       // 顧客マスタ（濱田事務所の顧客。要削除）
  '行政問い合わせ記録', // 行政への問い合わせ記録
  '相談記録',           // 相談記録
  '算定年更管理',       // 算定基礎・年度更新の進捗（顧問先別）
  'ストレスチェック管理', // ストレスチェック進捗（事業所別）
  'アクティビティログ', // 操作履歴（実行者に旧氏名）
  '一言メッセージ',     // スタッフ間メッセージ（from/to に旧氏名）
  '顧問料設定',         // 顧問料マスタ（顧客別金額。要削除）
  '月別請求管理',       // 月次請求データ
  '勤務予定DB'          // 勤務予定（staff_name/staff_id に旧氏名）
];

// 保持（マスタ／設定）— 絶対に触らない。
var INIT_KEEP_SHEETS = ['休日マスタ']; // 祝日など汎用マスタ。事務所固有ではないので残す

// --- 手順0：実際のタブ名を一覧表示するだけの確認用ヘルパー（読み取り専用） ---
function listAllSheetNames_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var names = ss.getSheets().map(function(s) { return s.getName(); });
  console.log('========== 現在のタブ名一覧（' + names.length + '枚） ==========');
  console.log('現在のスプレッドID: ' + ss.getId());
  console.log('複製先(想定)ID    : ' + INIT_TARGET_SPREADSHEET_ID
    + (ss.getId() === INIT_TARGET_SPREADSHEET_ID ? '  → ✅一致' : '  → ⚠不一致（初期化対象外のスプレッドです）'));
  names.forEach(function(n, i) {
    var role = INIT_CLEAR_SHEETS.indexOf(n) !== -1 ? '［クリア対象］'
             : INIT_KEEP_SHEETS.indexOf(n)  !== -1 ? '［保持・マスタ］'
             : '［ホワイトリスト外 → 変更しない］';
    console.log((i + 1) + '. ' + n + '  ' + role);
  });
  // コード上の対象名で、実在しないものを警告（タブ名の相違チェック）
  INIT_CLEAR_SHEETS.forEach(function(name) {
    if (!ss.getSheetByName(name)) console.warn('[要確認] コード上のクリア対象「' + name + '」が実タブに見つかりません');
  });
  return names;
}

// --- 本体：運用データのクリア（三重ロック／既定は DRY_RUN） ---
function initializeSpreadsheetForNewOffice() {
  // ★★★ 本番実行するときだけ、次の2つを変更する（TARGET_SPREADSHEET_IDは変更不要） ★★★
  const DRY_RUN = true;                     // 本番時のみ false
  const CONFIRM_INITIALIZE = '';            // 本番時のみ 'INITIALIZE_NEW_OFFICE'
  const TARGET_SPREADSHEET_ID = INIT_TARGET_SPREADSHEET_ID; // 複製先IDのみ許可

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const currentId = ss.getId();

  // ---- 三重ロックの判定 ----
  const idMatches  = (currentId === TARGET_SPREADSHEET_ID);
  const confirmed  = (CONFIRM_INITIALIZE === 'INITIALIZE_NEW_OFFICE');
  const isProdRun  = (DRY_RUN === false) && confirmed && idMatches;

  console.log('========== 初期化 事前サマリー ' + (isProdRun ? '[本番実行]' : '[確認のみ・削除しません]') + ' ==========');
  console.log('現在のスプレッドID: ' + currentId);
  console.log('複製先(許可)ID    : ' + TARGET_SPREADSHEET_ID + (idMatches ? '  → ✅一致' : '  → ⚠不一致'));
  console.log('安全ロック: DRY_RUN=' + DRY_RUN + ' / CONFIRM一致=' + confirmed + ' / ID一致=' + idMatches);

  // ---- 対象シートと削除予定件数を集計してログ出力（読み取りのみ） ----
  const plan = [];
  let grandTotal = 0;
  INIT_CLEAR_SHEETS.forEach(function(name) {
    const sheet = ss.getSheetByName(name);
    if (!sheet) { console.warn('[SKIP] シートが見つかりません: ' + name); return; }
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const dataRows = Math.max(lastRow - 1, 0); // 1行目（見出し）は除外
    let formulaCount = 0;
    if (dataRows > 0 && lastCol > 0) {
      const formulas = sheet.getRange(2, 1, dataRows, lastCol).getFormulas();
      formulas.forEach(function(r) { r.forEach(function(f) { if (f) formulaCount++; }); });
    }
    plan.push({ name: name, dataRows: dataRows, lastCol: lastCol, formulaCount: formulaCount });
    grandTotal += dataRows;
    console.log('- ' + name + '：削除予定 ' + dataRows + ' 行'
      + (formulaCount ? '（うち数式セル ' + formulaCount + ' 個は保持）' : ''));
  });
  console.log('保持（マスタ）：' + INIT_KEEP_SHEETS.join(', ') + ' ／ ホワイトリスト外のタブも一切触りません');
  console.log('削除予定 合計：' + grandTotal + ' 行');

  // ---- 本番でなければここで終了（削除しない） ----
  if (!isProdRun) {
    if (DRY_RUN !== false)  console.log('→ DRY_RUN が true のため削除しません。');
    if (!confirmed)         console.log('→ CONFIRM_INITIALIZE が \'INITIALIZE_NEW_OFFICE\' でないため削除しません。');
    if (!idMatches)         console.warn('→ 現在のスプレッドIDが複製先IDと一致しないため削除しません（元・濱田事務所での誤実行を防止）。');
    console.log('[中止] 実削除は行っていません。3つの条件がすべて揃った時のみ実行されます。');
    return;
  }

  // ---- 本番：スプレッドシート全体のバックアップを作成 ----
  const stamp = Utilities.formatDate(new Date(), tz, 'yyyyMMdd_HHmmss');
  const backup = DriveApp.getFileById(ss.getId()).makeCopy(ss.getName() + '_backup_' + stamp);
  console.log('[BACKUP] バックアップ作成完了：' + backup.getName() + '（ID: ' + backup.getId() + '）');

  // ---- クリア実行（見出し・数式・書式・入力規則は維持） ----
  const cleared = [];
  plan.forEach(function(p) {
    if (p.dataRows <= 0) { console.log('[SKIP] データ行なし：' + p.name); return; }
    const sheet = ss.getSheetByName(p.name);
    const range = sheet.getRange(2, 1, p.dataRows, p.lastCol);
    const formulas = p.formulaCount > 0 ? range.getFormulas() : null; // 数式を退避
    range.clearContent();                       // 値のみ消去（書式・入力規則・見出しは保持）
    if (formulas) range.setFormulas(formulas);  // 数式セルのみ復元
    cleared.push(p.name + '(' + p.dataRows + '行)');
    console.log('[DONE] ' + p.name + '：' + p.dataRows + ' 行をクリア'
      + (p.formulaCount ? '（数式 ' + p.formulaCount + ' 個は復元）' : ''));
  });

  // ---- 実行後サマリー（削除件数と対象シート） ----
  console.log('========== 初期化 完了サマリー ==========');
  console.log('クリアした合計：' + grandTotal + ' 行');
  console.log('対象シート：' + (cleared.length ? cleared.join(' , ') : '（なし）'));
  console.log('バックアップ：' + backup.getName());
  console.log('※ アプリを再読み込みし、各画面が空の状態で正常表示されるか確認してください。');
}

// ================================================================
// 【引き渡し前・一回限りの初期化セクション ここまで】
//   ★ 引き渡し時は、上の「ここから」からこの行までを丸ごと削除する ★
// ================================================================
