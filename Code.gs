/**
 * ダブル受験 進捗ダッシュボード - Google Apps Script
 * 
 * 機能:
 * - 進捗データの読み込み（GET）
 * - 進捗データの保存（POST）
 * - スプレッドシートとの連携
 * 
 * セットアップ:
 * 1. 新しいGoogleスプレッドシートを作成
 * 2. 拡張機能 > Apps Script を開く
 * 3. このコードを貼り付け
 * 4. デプロイ > 新しいデプロイ でウェブアプリとして公開
 */

// スプレッドシートの設定
const SHEET_NAME = 'ダッシュボードデータ';

/**
 * スプレッドシートを初期化（初回のみ実行）
 */
function initializeSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    
    // ヘッダー行を設定
    sheet.getRange('A1:D1').setValues([['キー', 'データ', '更新日時', '更新者']]);
    sheet.getRange('A1:D1').setFontWeight('bold');
    sheet.getRange('A1:D1').setBackground('#f3f4f6');
    
    // 初期データ行を作成
    sheet.getRange('A2').setValue('progress');
    sheet.getRange('B2').setValue('{}');
    sheet.getRange('A3').setValue('reviewLog');
    sheet.getRange('B3').setValue('{}');
    
    // 列幅を調整
    sheet.setColumnWidth(1, 120);
    sheet.setColumnWidth(2, 500);
    sheet.setColumnWidth(3, 180);
    sheet.setColumnWidth(4, 150);
    
    // フィルタを有効化
    sheet.getRange('A1:D3').createFilter();
    
    Logger.log('スプレッドシートを初期化しました');
  }
  
  return sheet;
}

/**
 * GETリクエストを処理（データ読み込み）
 */
function doGet(e) {
  try {
    const sheet = getOrCreateSheet();
    const data = readDataFromSheet(sheet);
    
    return createJsonResponse({
      success: true,
      data: data,
      timestamp: new Date().toISOString()
    });
    
  } catch (error) {
    Logger.log('GET Error: ' + error.toString());
    return createJsonResponse({
      success: false,
      error: error.toString()
    });
  }
}

/**
 * POSTリクエストを処理（データ保存）
 */
function doPost(e) {
  try {
    const sheet = getOrCreateSheet();
    
    // リクエストボディをパース
    let requestData;
    try {
      requestData = JSON.parse(e.postData.contents);
    } catch (parseError) {
      throw new Error('JSONの解析に失敗しました: ' + parseError.toString());
    }
    
    // データを保存
    saveDataToSheet(sheet, requestData);
    
    return createJsonResponse({
      success: true,
      message: 'データを保存しました',
      timestamp: new Date().toISOString()
    });
    
  } catch (error) {
    Logger.log('POST Error: ' + error.toString());
    return createJsonResponse({
      success: false,
      error: error.toString()
    });
  }
}

/**
 * シートを取得または作成
 */
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = initializeSpreadsheet();
  }
  
  return sheet;
}

/**
 * シートからデータを読み込む
 */
function readDataFromSheet(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), 1);
  
  if (lastRow < 2) {
    return { progress: {}, reviewLog: {} };
  }
  
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 2);
  const values = dataRange.getValues();
  
  const result = {
    progress: {},
    reviewLog: {}
  };
  
  values.forEach(row => {
    const key = row[0];
    const value = row[1];
    
    if (key && value) {
      try {
        const parsedValue = JSON.parse(value);
        if (key === 'progress') {
          result.progress = parsedValue;
        } else if (key === 'reviewLog') {
          result.reviewLog = parsedValue;
        }
      } catch (e) {
        Logger.log('JSON parse error for key: ' + key);
      }
    }
  });
  
  return result;
}

/**
 * シートにデータを保存
 */
function saveDataToSheet(sheet, data) {
  const now = new Date();
  const timestamp = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  const user = Session.getActiveUser().getEmail() || '匿名';
  
  // progressデータを保存
  if (data.progress !== undefined) {
    updateOrInsertRow(sheet, 'progress', JSON.stringify(data.progress), timestamp, user);
  }
  
  // reviewLogデータを保存
  if (data.reviewLog !== undefined) {
    updateOrInsertRow(sheet, 'reviewLog', JSON.stringify(data.reviewLog), timestamp, user);
  }
}

/**
 * 行を更新または挿入
 */
function updateOrInsertRow(sheet, key, value, timestamp, user) {
  const lastRow = Math.max(sheet.getLastRow(), 1);
  
  // 既存の行を検索
  for (let i = 2; i <= lastRow; i++) {
    if (sheet.getRange(i, 1).getValue() === key) {
      // 既存行を更新
      sheet.getRange(i, 2, 1, 3).setValues([[value, timestamp, user]]);
      return;
    }
  }
  
  // 新規行を追加
  const newRow = lastRow + 1;
  sheet.getRange(newRow, 1, 1, 4).setValues([[key, value, timestamp, user]]);
}

/**
 * JSONレスポンスを作成
 */
function createJsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * テスト用: データを手動で読み込む
 */
function testRead() {
  const sheet = getOrCreateSheet();
  const data = readDataFromSheet(sheet);
  Logger.log(JSON.stringify(data, null, 2));
}

/**
 * テスト用: サンプルデータを保存
 */
function testSave() {
  const sheet = getOrCreateSheet();
  const testData = {
    progress: {
      'sr1': 3,
      'sr2': 2,
      'gs1': 1
    },
    reviewLog: {
      '2025-12-L1': true,
      '2026-01-L1': true
    }
  };
  saveDataToSheet(sheet, testData);
  Logger.log('テストデータを保存しました');
}

/**
 * すべてのデータをリセット
 */
function resetAllData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '確認',
    'すべての進捗データをリセットしますか？\nこの操作は取り消せません。',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    const sheet = getOrCreateSheet();
    const emptyData = {
      progress: {},
      reviewLog: {}
    };
    saveDataToSheet(sheet, emptyData);
    ui.alert('データをリセットしました');
  }
}

/**
 * カスタムメニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📚 ダッシュボード管理')
    .addItem('🔧 初期化（初回のみ）', 'initializeSpreadsheet')
    .addItem('📖 データ読み込みテスト', 'testRead')
    .addItem('💾 テストデータ保存', 'testSave')
    .addSeparator()
    .addItem('🗑️ すべてのデータをリセット', 'resetAllData')
    .addToUi();
}
