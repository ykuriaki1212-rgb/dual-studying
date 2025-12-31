// ========================================
// スタディダッシュボード Google Apps Script（完全版）
// ========================================

// ========== 設定 ==========
const SHEET_PROGRESS = '進捗データ';
const SHEET_LOG = '更新ログ';

// ========== スプレッドシート取得 ==========
function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

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
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_PROGRESS);
  
  if (!sheet) {
    throw new Error('進捗データシートが見つかりません');
  }
  
  const data = sheet.getDataRange().getValues();
  
  if (data.length < 2) {
    return {
      progress: {},
      reviewLog: {}
    };
  }
  
  const headers = data[0];
  const values = data[1];
  
  const result = {
    progress: {},
    reviewLog: {}
  };
  
  for (let i = 2; i < headers.length; i++) {
    const key = headers[i];
    const value = values[i];
    
    if (key.startsWith('sr') || key.startsWith('gs')) {
      result.progress[key] = value || 0;
    } else if (key.match(/^\d{4}-\d{2}-L\d$/)) {
      result.reviewLog[key] = value === true || value === 'TRUE' || value === 'true';
    }
  }
  
  return result;
}

// ========== データ保存 ==========
function saveProgressData(data) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_PROGRESS);
  
  if (!sheet) {
    throw new Error('進捗データシートが見つかりません');
  }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = ['main', new Date().toISOString()];
  
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
  
  sheet.getRange(2, 1, 1, values.length).setValues([values]);
}

// ========== ログ記録 ==========
function logOperation(operation, detail) {
  try {
    const ss = getSpreadsheet();
    let logSheet = ss.getSheetByName(SHEET_LOG);
    
    if (!logSheet) {
      return;
    }
    
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
    const user = Session.getActiveUser().getEmail() || '不明';
    
    logSheet.appendRow([timestamp, operation, user, detail]);
    
  } catch (error) {
    console.error('ログ記録エラー:', error);
  }
}

// ========== 初期化用テスト関数 ==========
function testInit() {
  try {
    console.log('=== 初期化開始 ===');
    console.log('');
    
    const ss = getSpreadsheet();
    console.log('✓ スプレッドシート取得成功');
    console.log('  名前:', ss.getName());
    console.log('  ID:', ss.getId());
    console.log('');
    
    // 進捗データシート作成
    console.log('[1/4] 進捗データシート作成中...');
    let progressSheet = ss.getSheetByName(SHEET_PROGRESS);
    
    if (progressSheet) {
      ss.deleteSheet(progressSheet);
      console.log('  既存シートを削除しました');
    }
    
    progressSheet = ss.insertSheet(SHEET_PROGRESS);
    console.log('  新しいシートを作成しました');
    
    // ヘッダー作成
    const headers = ['ID', '最終更新日時'];
    
    for (let i = 1; i <= 10; i++) {
      headers.push(`sr${i}`);
    }
    
    for (let i = 1; i <= 10; i++) {
      headers.push(`gs${i}`);
    }
    
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
    
    progressSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    const initialData = ['main', new Date().toISOString()];
    for (let i = 2; i < headers.length; i++) {
      initialData.push(0);
    }
    progressSheet.getRange(2, 1, 1, initialData.length).setValues([initialData]);
    
    console.log('  ✓ 進捗データシート作成完了（' + headers.length + '列）');
    console.log('');
    
    // 更新ログシート作成
    console.log('[2/4] 更新ログシート作成中...');
    let logSheet = ss.getSheetByName(SHEET_LOG);
    
    if (logSheet) {
      ss.deleteSheet(logSheet);
      console.log('  既存ログシートを削除しました');
    }
    
    logSheet = ss.insertSheet(SHEET_LOG);
    
    const logHeaders = ['タイムスタンプ', '操作', 'ユーザー', '詳細'];
    logSheet.getRange(1, 1, 1, logHeaders.length).setValues([logHeaders]);
    
    console.log('  ✓ 更新ログシート作成完了');
    console.log('');
    
    // 初期ログ追加
    console.log('[3/4] 初期ログ記録中...');
    logOperation('初期化', 'スプレッドシートを初期化しました');
    console.log('  ✓ 初期ログ記録完了');
    console.log('');
    
    // 書式設定
    console.log('[4/4] 書式設定中...');
    formatSheets();
    console.log('  ✓ 書式設定完了');
    console.log('');
    
    console.log('========================================');
    console.log('✅ 初期化が完了しました！');
    console.log('========================================');
    console.log('');
    console.log('📋 作成されたシート:');
    console.log('  1. ' + SHEET_PROGRESS + ' (' + headers.length + '列)');
    console.log('  2. ' + SHEET_LOG);
    console.log('');
    console.log('📝 次のステップ:');
    console.log('  1. スプレッドシートを確認してください');
    console.log('  2. 「デプロイ」→「新しいデプロイ」を実行');
    console.log('  3. アクセス権限を「全員」に設定');
    console.log('  4. デプロイURLをコピー');
    console.log('  5. dashboard.htmlのAPI_URLに貼り付け');
    console.log('');
    
  } catch (error) {
    console.error('========================================');
    console.error('❌ 初期化エラーが発生しました');
    console.error('========================================');
    console.error('エラー内容:', error.message);
    console.error('');
    console.error('🔍 トラブルシューティング:');
    console.error('');
    console.error('【方法1】認証を再実行');
    console.error('  1. Apps Scriptエディタを閉じる');
    console.error('  2. スプレッドシートに戻る');
    console.error('  3. もう一度「拡張機能」→「Apps Script」を開く');
    console.error('  4. コードを貼り付けて保存');
    console.error('  5. testInitを実行して認証を許可');
    console.error('');
    console.error('【方法2】新しいスプレッドシートで試す');
    console.error('  1. 新しいスプレッドシートを作成');
    console.error('  2. 「拡張機能」→「Apps Script」を開く');
    console.error('  3. このコードを貼り付け');
    console.error('  4. testInitを実行');
    console.error('');
    console.error('【方法3】マニフェストファイルを確認');
    console.error('  1. 左メニューの「プロジェクトの設定」（歯車アイコン）');
    console.error('  2. 「appsscript.json」マニフェストをエディタで表示にチェック');
    console.error('  3. 左メニューに「appsscript.json」が表示される');
    console.error('  4. 中身を確認（次の応答で提供します）');
    console.error('');
  }
}

// ========== 書式設定 ==========
function formatSheets() {
  const ss = getSpreadsheet();
  
  // 進捗データシート
  const progressSheet = ss.getSheetByName(SHEET_PROGRESS);
  if (progressSheet) {
    const headerRange = progressSheet.getRange(1, 1, 1, progressSheet.getLastColumn());
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    progressSheet.autoResizeColumns(1, progressSheet.getLastColumn());
    progressSheet.setFrozenRows(1);
  }
  
  // 更新ログシート
  const logSheet = ss.getSheetByName(SHEET_LOG);
  if (logSheet) {
    const headerRange = logSheet.getRange(1, 1, 1, logSheet.getLastColumn());
    headerRange.setBackground('#34a853');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    logSheet.setColumnWidth(1, 150);
    logSheet.setColumnWidth(2, 80);
    logSheet.setColumnWidth(3, 200);
    logSheet.setColumnWidth(4, 300);
    logSheet.setFrozenRows(1);
  }
}

// ========== テスト関数 ==========
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

function testSave() {
  try {
    console.log('=== データ保存テスト ===');
    const testData = {
      progress: { sr1: 5, sr2: 3, gs1: 2 },
      reviewLog: { '2025-12-L1': true, '2025-12-L2': false }
    };
    saveProgressData(testData);
    console.log('保存したデータ:', JSON.stringify(testData, null, 2));
    console.log('✅ 保存テスト成功');
  } catch (error) {
    console.error('❌ 保存テストエラー:', error);
  }
}

function checkSpreadsheetInfo() {
  try {
    const ss = getSpreadsheet();
    console.log('=== スプレッドシート情報 ===');
    console.log('名前:', ss.getName());
    console.log('ID:', ss.getId());
    console.log('URL:', ss.getUrl());
    console.log('シート数:', ss.getSheets().length);
    console.log('シート名一覧:');
    ss.getSheets().forEach(sheet => {
      console.log('  -', sheet.getName());
    });
  } catch (error) {
    console.error('❌ エラー:', error);
  }
}
