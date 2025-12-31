// ========================================
// スタディダッシュボード Google Apps Script
// ========================================

// ========== 設定 ==========
// ※ スプレッドシートIDを自分のIDに置き換えてください
const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID';
const SHEET_PROGRESS = '進捗データ';
const SHEET_LOG = '更新ログ';

// ========== GETリクエスト（データ読込）==========
function doGet(e) {
  try {
    const data = loadProgressData();
    logOperation('読込', 'データを読み込みました');
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      data: data,
      timestamp: new Date().toISOString()
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    console.error('GET エラー:', error);
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ========== POSTリクエスト（データ保存）==========
function doPost(e) {
  try {
    const postData = JSON.parse(e.postData.contents);
    saveProgressData(postData);
    logOperation('保存', 'データを保存しました');
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: 'データを保存しました',
      timestamp: new Date().toISOString()
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    console.error('POST エラー:', error);
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ========== データ読込 ==========
function loadProgressData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_PROGRESS);
  
  if (!sheet) {
    throw new Error('進捗データシートが見つかりません');
  }
  
  const data = sheet.getDataRange().getValues();
  
  if (data.length < 2) {
    // データがない場合は空のオブジェクトを返す
    return {
      progress: {},
      reviewLog: {}
    };
  }
  
  const headers = data[0];
  const values = data[1]; // ID="main"の行
  
  const result = {
    progress: {},
    reviewLog: {}
  };
  
  // ヘッダーとデータをマッピング
  for (let i = 2; i < headers.length; i++) { // ID, 最終更新日時をスキップ
    const key = headers[i];
    const value = values[i];
    
    if (key.startsWith('sr') || key.startsWith('gs')) {
      // 科目の周回数
      result.progress[key] = value || 0;
    } else if (key.match(/^\d{4}-\d{2}-L\d$/)) {
      // レビューログ
      result.reviewLog[key] = value === true || value === 'TRUE' || value === 'true';
    }
  }
  
  return result;
}

// ========== データ保存 ==========
function saveProgressData(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_PROGRESS);
  
  if (!sheet) {
    throw new Error('進捗データシートが見つかりません');
  }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = ['main', new Date().toISOString()]; // ID, 最終更新日時
  
  // ヘッダーに合わせてデータを配置
  for (let i = 2; i < headers.length; i++) {
    const key = headers[i];
    
    if (data.progress && data.progress[key] !== undefined) {
      values.push(data.progress[key]);
    } else if (data.reviewLog && data.reviewLog[key] !== undefined) {
      values.push(data.reviewLog[key]);
    } else {
      values.push('');
    }
  }
  
  // 2行目にデータを書き込み
  sheet.getRange(2, 1, 1, values.length).setValues([values]);
}

// ========== ログ記録 ==========
function logOperation(operation, detail) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let logSheet = ss.getSheetByName(SHEET_LOG);
    
    if (!logSheet) {
      return; // ログシートがなければスキップ
    }
    
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
    const user = Session.getActiveUser().getEmail() || '不明';
    
    logSheet.appendRow([timestamp, operation, user, detail]);
    
  } catch (error) {
    console.error('ログ記録エラー:', error);
  }
}

// ========== 初期化用テスト関数 ==========
/**
 * スプレッドシートの初期化を行います
 * 注意: この関数は実行ログで結果を確認してください（alert()は使用していません）
 */
function testInit() {
  try {
    console.log('=== 初期化開始 ===');
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // 進捗データシート作成
    console.log('進捗データシート作成中...');
    let progressSheet = ss.getSheetByName(SHEET_PROGRESS);
    
    if (progressSheet) {
      ss.deleteSheet(progressSheet);
      console.log('既存シートを削除しました');
    }
    
    progressSheet = ss.insertSheet(SHEET_PROGRESS);
    
    // ヘッダー作成
    const headers = ['ID', '最終更新日時'];
    
    // 社労士科目（sr1〜sr10）
    for (let i = 1; i <= 10; i++) {
      headers.push(`sr${i}`);
    }
    
    // 行政書士科目（gs1〜gs10）
    for (let i = 1; i <= 10; i++) {
      headers.push(`gs${i}`);
    }
    
    // レビューログ
    const reviewPeriods = [
      { month: '2025-12', levels: ['L1', 'L2'] },
      { month: '2026-01', levels: ['L1', 'L2'] },
      { month: '2026-02', levels: ['L1', 'L2'] },
      { month: '2026-03', levels: ['L1', 'L2', 'L3'] },
      { month: '2026-04', levels: ['L1', 'L2'] },
      { month: '2026-05', levels: ['L1', 'L2'] },
      { month: '2026-06', levels: ['L1', 'L2', 'L3'] }
    ];
    
    reviewPeriods.forEach(period => {
      period.levels.forEach(level => {
        headers.push(`${period.month}-${level}`);
      });
    });
    
    // ヘッダー書き込み
    progressSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // 初期データ行作成
    const initialData = ['main', new Date().toISOString()];
    for (let i = 2; i < headers.length; i++) {
      initialData.push(0);
    }
    progressSheet.getRange(2, 1, 1, initialData.length).setValues([initialData]);
    
    console.log(`進捗データシート作成完了（${headers.length}列）`);
    
    // 更新ログシート作成
    console.log('更新ログシート作成中...');
    let logSheet = ss.getSheetByName(SHEET_LOG);
    
    if (logSheet) {
      ss.deleteSheet(logSheet);
      console.log('既存ログシートを削除しました');
    }
    
    logSheet = ss.insertSheet(SHEET_LOG);
    
    const logHeaders = ['タイムスタンプ', '操作', 'ユーザー', '詳細'];
    logSheet.getRange(1, 1, 1, logHeaders.length).setValues([logHeaders]);
    
    console.log('更新ログシート作成完了');
    
    // 初期ログ追加
    logOperation('初期化', 'スプレッドシートを初期化しました');
    
    console.log('=== ✅ 初期化が完了しました ===');
    console.log(`シート名: ${SHEET_PROGRESS}, ${SHEET_LOG}`);
    console.log(`カラム数: ${headers.length}`);
    
  } catch (error) {
    console.error('❌ 初期化エラー:', error);
    console.error('エラー詳細:', error.stack);
  }
}

// ========== 書式設定（オプション）==========
/**
 * シートの書式を整えます（testInitとは別に実行）
 */
function formatSheets() {
  try {
    console.log('=== 書式設定開始 ===');
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // 進捗データシートの書式
    const progressSheet = ss.getSheetByName(SHEET_PROGRESS);
    if (progressSheet) {
      // ヘッダー行の書式
      const headerRange = progressSheet.getRange(1, 1, 1, progressSheet.getLastColumn());
      headerRange.setBackground('#4285f4');
      headerRange.setFontColor('#ffffff');
      headerRange.setFontWeight('bold');
      headerRange.setHorizontalAlignment('center');
      
      // 列幅の自動調整
      progressSheet.autoResizeColumns(1, progressSheet.getLastColumn());
      
      // 固定行
      progressSheet.setFrozenRows(1);
      
      console.log('進捗データシートの書式設定完了');
    }
    
    // 更新ログシートの書式
    const logSheet = ss.getSheetByName(SHEET_LOG);
    if (logSheet) {
      // ヘッダー行の書式
      const headerRange = logSheet.getRange(1, 1, 1, logSheet.getLastColumn());
      headerRange.setBackground('#34a853');
      headerRange.setFontColor('#ffffff');
      headerRange.setFontWeight('bold');
      headerRange.setHorizontalAlignment('center');
      
      // 列幅の調整
      logSheet.setColumnWidth(1, 150); // タイムスタンプ
      logSheet.setColumnWidth(2, 80);  // 操作
      logSheet.setColumnWidth(3, 200); // ユーザー
      logSheet.setColumnWidth(4, 300); // 詳細
      
      // 固定行
      logSheet.setFrozenRows(1);
      
      console.log('更新ログシートの書式設定完了');
    }
    
    console.log('=== ✅ 書式設定が完了しました ===');
    
  } catch (error) {
    console.error('❌ 書式設定エラー:', error);
  }
}

// ========== テスト用: データ読込確認 ==========
function testLoad() {
  try {
    console.log('=== データ読込テスト ===');
    const data = loadProgressData();
    console.log('読み込んだデータ:', JSON.stringify(data, null, 2));
    console.log('✅ 読込テスト成功');
  } catch (error) {
    console.error('❌ 読込テストエラー:', error);
  }
}

// ========== テスト用: データ保存確認 ==========
function testSave() {
  try {
    console.log('=== データ保存テスト ===');
    
    const testData = {
      progress: {
        sr1: 5,
        sr2: 3,
        gs1: 2
      },
      reviewLog: {
        '2025-12-L1': true,
        '2025-12-L2': false
      }
    };
    
    saveProgressData(testData);
    console.log('保存したデータ:', JSON.stringify(testData, null, 2));
    console.log('✅ 保存テスト成功');
    
  } catch (error) {
    console.error('❌ 保存テストエラー:', error);
  }
}
