// =========================================
// マスタ管理：勘定科目マスタ用 API
// =========================================

const ACCOUNT_MASTER_SHEET_NAME = 'マスタ_勘定科目';

/**
 * 勘定科目マスタの全データを取得し、HTMLへ返す
 */
function getAccountMasterData() {
  try {
    const ss = SpreadsheetApp.openById(MAIN_SS_ID);
    const sheet = ss.getSheetByName(ACCOUNT_MASTER_SHEET_NAME);
    
    if (!sheet) return { error: `シート「${ACCOUNT_MASTER_SHEET_NAME}」が見つかりません。` };
    
    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) return { success: true, headers: [], data: [] }; // データなし
    
    const headers = values[0];
    const data = [];
    
    // 2行目からデータをオブジェクトに変換
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const obj = {};
      headers.forEach((h, index) => {
        obj[h] = row[index];
      });
      data.push(obj);
    }
    
    return { success: true, headers: headers, data: data };
  } catch (e) {
    logErrorToSheet(e, 'getAccountMasterData');
    return { error: `マスタ取得エラー: ${e.message}` };
  }
}

/**
 * 画面で編集された勘定科目マスタのデータをシートに保存する（排他制御あり）
 * @param {Array} updatedDataArray - 更新後のオブジェクト配列
 */
function saveAccountMasterData(updatedDataArray) {
  const lock = LockService.getScriptLock();
  
  try {
    // 最大10秒間、他の人の処理が終わるのを待つ
    if (!lock.waitLock(10000)) {
      return { error: "現在他のユーザーが更新中です。少し待ってから再度お試しください。" };
    }
    
    const ss = SpreadsheetApp.openById(MAIN_SS_ID);
    const sheet = ss.getSheetByName(ACCOUNT_MASTER_SHEET_NAME);
    
    if (!sheet) return { error: `シート「${ACCOUNT_MASTER_SHEET_NAME}」が見つかりません。` };
    
    // 現在のヘッダーを取得
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // HTMLから来たオブジェクト配列を、シート書き込み用の2次元配列に変換
    const outputValues = [];
    updatedDataArray.forEach(obj => {
      const row = [];
      headers.forEach(h => {
         // プロパティが存在しない場合は空文字にする
         row.push(obj[h] !== undefined ? obj[h] : ""); 
      });
      outputValues.push(row);
    });
    
    // データ行（2行目以降）を一旦クリアして、新しいデータを一括書き込み
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
    }
    
    if (outputValues.length > 0) {
      sheet.getRange(2, 1, outputValues.length, headers.length).setValues(outputValues);
    }
    
    // ログを残す
    logActionToSheet('saveAccountMasterData', `勘定科目マスタを更新しました（${outputValues.length}件）`);
    
    return { success: true, message: "勘定科目マスタの保存が完了しました。" };
    
  } catch (e) {
    logErrorToSheet(e, 'saveAccountMasterData');
    return { error: `保存エラー: ${e.message}` };
  } finally {
    // 【重要】処理が終わったら必ずロックを解除
    lock.releaseLock();
  }// =========================================
// マスタ管理：CSVエクスポート・インポート機能
// =========================================

/**
 * 勘定科目マスタの全データをCSV文字列としてエクスポートする
 */
function exportAccountMasterCsv() {
  try {
    const ss = SpreadsheetApp.openById(MAIN_SS_ID);
    const sheet = ss.getSheetByName(ACCOUNT_MASTER_SHEET_NAME);
    const values = sheet.getDataRange().getValues();
    
    // 2次元配列をCSV形式の文字列に変換
    let csvString = "";
    values.forEach(row => {
      // カンマや改行が含まれるデータ対策としてダブルクォーテーションで囲む
      const escapedRow = row.map(cell => `"${String(cell).replace(/"/g, '""')}"`);
      csvString += escapedRow.join(",") + "\n";
    });
    
    // 文字化け防止のためのBOM付きUTF-8として返す準備
    return { success: true, csvData: csvString };
  } catch (e) {
    logErrorToSheet(e, 'exportAccountMasterCsv');
    return { error: `エクスポートエラー: ${e.message}` };
  }
}

/**
 * CSVデータを受け取り、勘定科目マスタを上書き更新する
 * @param {string} csvString - アップロードされたCSVの文字列データ
 */
function importAccountMasterCsv(csvString) {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.waitLock(10000)) {
      return { error: "現在他のユーザーが更新中です。少し待ってから再度お試しください。" };
    }
    
    // CSV文字列を2次元配列にパース
    const csvData = Utilities.parseCsv(csvString);
    if (!csvData || csvData.length < 2) {
      return { error: "CSVデータが空、またはヘッダーしかありません。" };
    }
    
    const ss = SpreadsheetApp.openById(MAIN_SS_ID);
    const sheet = ss.getSheetByName(ACCOUNT_MASTER_SHEET_NAME);
    
    // 一旦既存のデータをすべてクリア
    sheet.clearContents();
    
    // CSVのデータを一括書き込み
    sheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);
    
    logActionToSheet('importAccountMasterCsv', `CSVインポートで勘定科目マスタを更新しました（${csvData.length - 1}件）`);
    return { success: true, message: "CSVインポートが完了し、マスタが更新されました。" };
    
  } catch (e) {
    logErrorToSheet(e, 'importAccountMasterCsv');
    return { error: `インポートエラー: ${e.message}` };
  } finally {
    lock.releaseLock();
  }
}
}