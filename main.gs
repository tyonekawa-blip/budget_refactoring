// =========================================
// SPenD Dashboard - Backend Script (Main)
// =========================================

const MAIN_SS_ID = '1MC933tAVIxMIPX4VR68GNKcmC-VtTp0fbUGOofAVeb8';
const TEAM_SHEET_NAME = 'マスタ_チーム'; 

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('SPenD Dashboard')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function logErrorToSheet(e, funcName) {
  // ★ エラーログの書き込みにもロックをかける（10秒待機）
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openById(MAIN_SS_ID);
    let logSheet = ss.getSheetByName('ログシート') || ss.insertSheet('ログシート');
    if (logSheet.getLastRow() === 0) {
      logSheet.appendRow(['日時', 'ユーザー', '処理名', '状態 / 内容', 'スタックトレース（エラー時のみ）']);
      logSheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#f1f3f5');
    }
    const user = Session.getActiveUser().getEmail() || "Unknown User";
    logSheet.appendRow([new Date(), user, funcName, "ERROR: " + e.message, e.stack]);
  } catch (err) { 
    console.error("Logging failed: " + err.message); 
  } finally {
    lock.releaseLock(); // 必ずロックを解除する
  }
}

function logActionToSheet(funcName, actionDetail) {
  // ★ 操作ログの書き込みにもロックをかける
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openById(MAIN_SS_ID);
    let logSheet = ss.getSheetByName('ログシート') || ss.insertSheet('ログシート');
    
    if (logSheet.getLastRow() === 0) {
      logSheet.appendRow(['日時', 'ユーザー', '処理名', '状態 / 内容', 'スタックトレース（エラー時のみ）']);
      logSheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#f1f3f5');
    }
    
    const user = Session.getActiveUser().getEmail() || "Unknown User";
    logSheet.appendRow([new Date(), user, funcName, actionDetail, '']);
  } catch (err) { 
    console.error("Action logging failed: " + err.message); 
  } finally {
    lock.releaseLock();
  }
}

function getFiscalYear(yyyymm) {
  const y = parseInt(yyyymm.substring(0, 4));
  const m = parseInt(yyyymm.substring(4, 6));
  return m >= 10 ? y : y - 1;
}

function getSystemSettings() {
  try {
    // スプレッドシートを読まず、爆速のスクリプトプロパティから取得！
    const props = PropertiesService.getScriptProperties().getProperties();
    return props;
  } catch (e) {
    logErrorToSheet(e, 'getSystemSettings');
    return {};
  }
}