/**
 * ダブル受験 進捗ダッシュボード用 Google Apps Script
 * 
 * 【セットアップ手順】
 * 1. Google スプレッドシートを新規作成
 * 2. 拡張機能 → Apps Script を開く
 * 3. このコードを貼り付けて保存
 * 4. デプロイ → 新しいデプロイ
 * 5. 種類: ウェブアプリ
 * 6. 実行するユーザー: 自分
 * 7. アクセスできるユーザー: 全員
 * 8. デプロイしてURLをコピー
 * 9. ダッシュボードの設定でURLを入力
 */

// スプレッドシートのシート名
const SHEET_NAME = 'Data';

// データを保存するセル
const DATA_CELL = 'A1';

/**
 * GETリクエスト: データを読み込む
 */
function doGet(e) {
  try {
    const sheet = getOrCreateSheet();
    const data = sheet.getRange(DATA_CELL).getValue();
    
    let parsedData = {
      progress: {},
      reviewLog: {},
      customTargets: {}
    };
    
    if (data) {
      try {
        parsedData = JSON.parse(data);
        // 後方互換性: customTargetsがない場合は空オブジェクトを設定
        if (!parsedData.customTargets) {
          parsedData.customTargets = {};
        }
      } catch (parseError) {
        console.error('データのパースエラー:', parseError);
      }
    }
    
    return createJsonResponse({
      success: true,
      data: parsedData
    });
    
  } catch (error) {
    console.error('読み込みエラー:', error);
    return createJsonResponse({
      success: false,
      error: error.message
    });
  }
}

/**
 * POSTリクエスト: データを保存する
 */
function doPost(e) {
  try {
    const sheet = getOrCreateSheet();
    
    // リクエストボディをパース
    let postData;
    try {
      postData = JSON.parse(e.postData.contents);
    } catch (parseError) {
      return createJsonResponse({
        success: false,
        error: 'リクエストデータのパースに失敗しました'
      });
    }
    
    // データを保存
    const dataToSave = {
      progress: postData.progress || {},
      reviewLog: postData.reviewLog || {},
      customTargets: postData.customTargets || {},
      lastUpdated: new Date().toISOString()
    };
    
    sheet.getRange(DATA_CELL).setValue(JSON.stringify(dataToSave));
    
    // 更新日時を別セルにも記録（確認用）
    sheet.getRange('B1').setValue('最終更新: ' + new Date().toLocaleString('ja-JP'));
    
    return createJsonResponse({
      success: true,
      message: 'データを保存しました'
    });
    
  } catch (error) {
    console.error('保存エラー:', error);
    return createJsonResponse({
      success: false,
      error: error.message
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
    sheet = ss.insertSheet(SHEET_NAME);
    // 初期データを設定
    sheet.getRange(DATA_CELL).setValue(JSON.stringify({
      progress: {},
      reviewLog: {},
      customTargets: {}
    }));
    sheet.getRange('B1').setValue('初期化: ' + new Date().toLocaleString('ja-JP'));
  }
  
  return sheet;
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
 * テスト用: データをリセット
 */
function resetData() {
  const sheet = getOrCreateSheet();
  sheet.getRange(DATA_CELL).setValue(JSON.stringify({
    progress: {},
    reviewLog: {},
    customTargets: {}
  }));
  sheet.getRange('B1').setValue('リセット: ' + new Date().toLocaleString('ja-JP'));
}

/**
 * テスト用: 現在のデータを確認
 */
function checkData() {
  const sheet = getOrCreateSheet();
  const data = sheet.getRange(DATA_CELL).getValue();
  console.log('現在のデータ:', data);
  return data;
}
