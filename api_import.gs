// =========================================
// SPenD Dashboard - Backend Script (Settings API)
// =========================================

function getSettingsData(year) {
  try {
    const ss = SpreadsheetApp.openById(MAIN_SS_ID);

    const subSheet = ss.getSheetByName('マスタ_勘定科目');
    let subjects = [];
    if (subSheet) {
      const subData = subSheet.getDataRange().getValues();
      const headers = subData[0];
      
      // 新しいヘッダー列名「free_cd」「free_name」などの位置を動的に検索してズレを解消
      const cdIdx = headers.indexOf('free_cd');
      const nameIdx = headers.indexOf('free_name');
      const descIdx = headers.indexOf('free_description');
      const typeMIdx = headers.indexOf('typeM');

      subjects = subData.slice(1).map(r => ({
        cd: cdIdx !== -1 ? String(r[cdIdx]).trim() : "",
        name: nameIdx !== -1 ? String(r[nameIdx]).trim() : "",
        desc: descIdx !== -1 ? String(r[descIdx]).trim() : "",
        typeM: typeMIdx !== -1 ? (parseInt(r[typeMIdx], 10) || 0) : 0
      })).filter(s => s.name); // 勘定科目名が存在するものに絞り込み
    }
    
    const subjectSettings = {};
    const subSetSheet = ss.getSheetByName('html科目設定保管');
    if (subSetSheet && subSetSheet.getLastRow() > 0) {
      const subSetData = subSetSheet.getDataRange().getValues();
      for (let i = 1; i < subSetData.length; i++) { 
        if (String(subSetData[i][0])) {
          subjectSettings[String(subSetData[i][0]).trim()] = (subSetData[i][1] === true || String(subSetData[i][1]).toUpperCase() === 'TRUE'); 
        }
      }
    }
    
    const actualSettings = {};
    const actSetSheet = ss.getSheetByName('html実績表示設定保管');
    if (actSetSheet && actSetSheet.getLastRow() > 0) {
      const actSetData = actSetSheet.getDataRange().getValues();
      for (let i = 1; i < actSetData.length; i++) { 
        const ymKey = String(actSetData[i][0]).trim(); 
        if (ymKey) {
          actualSettings[ymKey] = String(actSetData[i][1]).trim(); 
        }
      }
    }
    
    let config = {};
    const sysSheet = ss.getSheetByName('システム設定');
    if (sysSheet) {
      const data = sysSheet.getDataRange().getValues();
      for(let i = 0; i < data.length; i++) { 
        if(data[i][0]) {
          config[String(data[i][0]).trim()] = data[i][1]; 
        }
      }
    }

    let zaimuSubjects = [];
    let diffFrom = 0;
    let diffTo = 9999999999;
    
    const zaimuSheet = ss.getSheetByName('html重要科目・表示設定');
    if (zaimuSheet && zaimuSheet.getLastRow() > 0) {
      const zData = zaimuSheet.getDataRange().getValues();
      for (let i = 0; i < zData.length; i++) {
        const key = String(zData[i][0]).trim();
        const val = zData[i][1];
        if (key === 'ZAIMU_DIFF_FROM') { diffFrom = Number(val) || 0; }
        else if (key === 'ZAIMU_DIFF_TO') { diffTo = Number(val) || 9999999999; }
        else if (key && key !== '項目名' && key !== '勘定科目名' && (val === true || String(val).toUpperCase() === 'TRUE')) {
          zaimuSubjects.push(key);
        }
      }
    }
    const zaimuSettings = { subjects: zaimuSubjects, diffFrom: diffFrom, diffTo: diffTo };

    const targetYear = parseInt(year) || parseInt(config['CURRENT_YEAR']) || 2025;
    const monthList = [];
    for (let i = 0; i < 12; i++) { 
      let y = targetYear; 
      let m = 10 + i; 
      if (m > 12) { y += 1; m -= 12; } 
      monthList.push(`${y}${m.toString().padStart(2, '0')}`); 
    }
    
    return { subjects: subjects, subjectSettings: subjectSettings, actualSettings: actualSettings, zaimuSettings: zaimuSettings, monthList: monthList };
  } catch(e) { return { error: e.message }; }
}

function saveZaimuSettingsConfig(arg1, arg2, arg3) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); 
    const ss = SpreadsheetApp.openById(MAIN_SS_ID);
    let subjectsObj = {};
    let diffFrom = 0;
    let diffTo = 9999999999;
    
    if (arg1) {
      if (Array.isArray(arg1)) {
        arg1.forEach(k => { if(k) subjectsObj[k] = true; });
        diffFrom = arg2 !== undefined ? Number(arg2) : 0;
        diffTo = arg3 !== undefined ? Number(arg3) : 9999999999;
      } else if (typeof arg1 === 'object') {
        if (Array.isArray(arg1.subjects)) {
          arg1.subjects.forEach(k => { if(k) subjectsObj[k] = true; });
        } else if (typeof arg1.subjects === 'object' && arg1.subjects !== null) {
          subjectsObj = arg1.subjects;
        } else {
          subjectsObj = arg1;
        }
        diffFrom = arg1.diffFrom !== undefined ? Number(arg1.diffFrom) : (arg2 !== undefined ? Number(arg2) : 0);
        diffTo = arg1.diffTo !== undefined ? Number(arg1.diffTo) : (arg3 !== undefined ? Number(arg3) : 9999999999);
      }
    }

    let zSheet = ss.getSheetByName('html重要科目・表示設定') || ss.insertSheet('html重要科目・表示設定');
    zSheet.clearContents();
    
    const rows = [['項目名', '設定値']];
    rows.push(['ZAIMU_DIFF_FROM', diffFrom]);
    rows.push(['ZAIMU_DIFF_TO', diffTo]);
    
    Object.keys(subjectsObj).forEach(name => {
      if(subjectsObj[name] && name !== 'ZAIMU_DIFF_FROM' && name !== 'ZAIMU_DIFF_TO') {
        rows.push([name, true]);
      }
    });
    zSheet.getRange(1, 1, rows.length, 2).setValues(rows);
    logActionToSheet('saveZaimuSettingsConfig', `財務戦略レポートの設定を更新しました。(対象科目数: ${Object.keys(subjectsObj).length}件)`);
    return { success: true };
  } catch (e) { 
    return { error: "サーバーが混雑しています。しばらく待ってから再度保存してください。(" + e.message + ")" }; 
  } finally {
    lock.releaseLock();
  }
}

function exportZaimuSettingsCsv() {
  try {
    const ss = SpreadsheetApp.openById(MAIN_SS_ID);
    const sheet = ss.getSheetByName('html重要科目・表示設定');
    if (!sheet) return "";
    const data = sheet.getDataRange().getValues();
    logActionToSheet('exportZaimuSettingsCsv', '財務レポート設定のCSVエクスポートが実行されました。');
    return data.map(row => row.join(",")).join("\n");
  } catch(e) { return ""; }
}

function importZaimuSettingsCsv(csvString) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const ss = SpreadsheetApp.openById(MAIN_SS_ID);
    let sheet = ss.getSheetByName('html重要科目・表示設定') || ss.insertSheet('html重要科目・表示設定');
    const data = Utilities.parseCsv(csvString);
    sheet.clearContents();
    if (data.length > 0) sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    logActionToSheet('importZaimuSettingsCsv', `CSVから財務レポート設定をインポートしました。`);
    return { success: true };
  } catch (e) { 
    return { error: "同時操作による書き込みエラーです。時間をおいて再試行してください。" }; 
  } finally {
    lock.releaseLock();
  }
}

function saveSubjectSettings(arg) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const ss = SpreadsheetApp.openById(MAIN_SS_ID);
    let sheet = ss.getSheetByName('html科目設定保管') || ss.insertSheet('html科目設定保管');
    sheet.getDataRange().clearContent();
    let rows = [];
    if (Array.isArray(arg)) { rows = arg; } 
    else if (typeof arg === 'object' && arg !== null) { rows = Object.keys(arg).map(k => [k, arg[k]]); }
    if (rows.length > 0) sheet.getRange(1, 1, rows.length, 2).setValues(rows);
    logActionToSheet('saveSubjectSettings', '勘定科目 PL表示設定を保存しました。');
    return true;
  } catch (e) { 
    throw new Error("サーバー混雑エラー: " + e.message); 
  } finally {
    lock.releaseLock();
  }
}

function saveActualSettings(arg) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const ss = SpreadsheetApp.openById(MAIN_SS_ID);
    let sheet = ss.getSheetByName('html実績表示設定保管') || ss.insertSheet('html実績表示設定保管');
    sheet.getDataRange().clearContent();
    let rows = [];
    if (Array.isArray(arg)) { rows = arg; } 
    else if (typeof arg === 'object' && arg !== null) { rows = Object.keys(arg).map(k => [k, arg[k]]); }
    if (rows.length > 0) sheet.getRange(1, 1, rows.length, 2).setValues(rows);
    logActionToSheet('saveActualSettings', '実績データの表示切り替え設定を保存しました。');
    return true;
  } catch (e) { 
    throw new Error("サーバー混雑エラー: " + e.message); 
  } finally {
    lock.releaseLock();
  }
}

function saveWeatherSettings(sunny, cloudy, rainy, sunnyAmt, cloudyAmt, rainyAmt, maxLength) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const ss = SpreadsheetApp.openById(MAIN_SS_ID);
    let sheet = ss.getSheetByName('システム設定');
    if (!sheet) return { error: "システム設定シートがありません" };
    const data = sheet.getDataRange().getValues();
    const val = JSON.stringify({ 
        sunny: sunny, cloudy: cloudy, rainy: rainy,
        sunnyAmt: sunnyAmt || 0, cloudyAmt: cloudyAmt || 0, rainyAmt: rainyAmt || 0,
        maxLength: maxLength || 40
    });
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === 'WEATHER_SETTINGS') { 
        sheet.getRange(i + 1, 2).setValue(val); 
        found = true; 
        break; 
      }
    }
    if (!found) sheet.appendRow(['WEATHER_SETTINGS', val, '創会レポート天気設定']);
    logActionToSheet('saveWeatherSettings', `創会レポートの天気・判定設定を更新しました。`);
    return { success: true };
  } catch (e) { 
    return { error: "同時書き込みによるエラーです。再度保存を押してください。" }; 
  } finally {
    lock.releaseLock();
  }
}

function saveTeamSettings(arr) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const ss = SpreadsheetApp.openById(MAIN_SS_ID);
    let sheet = ss.getSheetByName(typeof TEAM_SHEET_NAME !== 'undefined' ? TEAM_SHEET_NAME : 'マスタ_チーム');
    if (!sheet) { sheet = ss.insertSheet(typeof TEAM_SHEET_NAME !== 'undefined' ? TEAM_SHEET_NAME : 'マスタ_チーム'); sheet.appendRow(['チームCD', 'チーム名', '表示順', 'セグメントCDリスト', '非表示']); }
    if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).clearContent();
    if (arr && arr.length > 0) sheet.getRange(2, 1, arr.length, 5).setValues(arr); 
    logActionToSheet('saveTeamSettings', `報告部署（チーム）の編成を更新しました。`);
    return true;
  } catch (e) { 
    throw new Error("サーバー混雑エラー: " + e.message); 
  } finally {
    lock.releaseLock();
  }
}

function saveUserSettings(arr) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const ss = SpreadsheetApp.openById(MAIN_SS_ID);
    let sheet = ss.getSheetByName('ユーザー設定');
    if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).clearContent();
    if (arr.length > 0) sheet.getRange(2, 1, arr.length, 7).setValues(arr);
    logActionToSheet('saveUserSettings', `ユーザー権限設定を更新しました。`);
    return true;
  } catch (e) { 
    throw new Error("サーバー混雑エラー: " + e.message); 
  } finally {
    lock.releaseLock();
  }
}

function saveComment(period, segCd, key, text) { 
  const lock = LockService.getScriptLock();
  try { 
    lock.waitLock(15000); 
    const ss = SpreadsheetApp.openById(MAIN_SS_ID);
    let config = {};
    const sysSheet = ss.getSheetByName('システム設定');
    if(sysSheet){ const data = sysSheet.getDataRange().getValues(); for(let i=0; i<data.length; i++) { if(data[i][0]) config[String(data[i][0]).trim()] = data[i][1]; } }
    const sheetName = String(config['COMMENT_SHEET_NAME'] || 'htmlコメント保管').trim();
    let sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
    const data = sheet.getDataRange().getValues(); 
    
    let isUpdate = false;
    for (let i = 1; i < data.length; i++) { 
      if (String(data[i][0]) === period && String(data[i][1]) === segCd && String(data[i][2]) === key) { 
        sheet.getRange(i + 1, 4).setValue(text); 
        isUpdate = true;
        break;
      } 
    } 
    if(!isUpdate) sheet.appendRow([period, segCd, key, text]);
    logActionToSheet('saveComment', `【予実分析コメント】対象期間: ${period} のコメントを保存しました。`);
    return true;
  } catch (e) { 
    throw new Error("他ユーザーの保存処理と重なりました。数秒待ってから再度保存してください。"); 
  } finally {
    lock.releaseLock();
  }
}

function saveReportComment(period, segCd, text) { 
  const lock = LockService.getScriptLock();
  try { 
    lock.waitLock(15000);
    const ss = SpreadsheetApp.openById(MAIN_SS_ID);
    let config = {};
    const sysSheet = ss.getSheetByName('システム設定');
    if(sysSheet){ const data = sysSheet.getDataRange().getValues(); for(let i=0; i<data.length; i++) { if(data[i][0]) config[String(data[i][0]).trim()] = data[i][1]; } }
    const sheetName = String(config['REPORT_COMMENT_SHEET_NAME'] || 'html報告資料コメント保管').trim();
    let sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
    const data = sheet.getDataRange().getValues(); 
    
    let isUpdate = false;
    for (let i = 1; i < data.length; i++) { 
      if (String(data[i][0]) === period && String(data[i][1]) === segCd) { 
        sheet.getRange(i + 1, 3).setValue(text); 
        isUpdate = true;
        break;
      } 
    } 
    if(!isUpdate) sheet.appendRow([period, segCd, text]);
    logActionToSheet('saveReportComment', `【創会レポート コメント保存】対象: ${segCd} のコメントを保存しました。`);
    return true;
  } catch (e) { 
    throw new Error("他ユーザーの保存処理と重なりました。数秒待ってから再度保存してください。"); 
  } finally {
    lock.releaseLock();
  }
}

function updateCurrentYear(y) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openById(MAIN_SS_ID);
    let sheet = ss.getSheetByName('システム設定');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === 'CURRENT_YEAR') { sheet.getRange(i + 1, 2).setValue(y); return true; }
    }
    sheet.appendRow(['CURRENT_YEAR', y, '現在の処理年度']);
    logActionToSheet('updateCurrentYear', `【システム年度更新】処理年度を『FY${y}』に変更しました。`);
    return true;
  } catch (e) { 
    throw new Error("サーバー混雑エラー: " + e.message); 
  } finally {
    lock.releaseLock();
  }
}

function getCloseStatuses(year) {
  try {
    const ss = SpreadsheetApp.openById(MAIN_SS_ID);
    const sheet = ss.getSheetByName('締め');
    const closedMonths = [];
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][2] === true && String(data[i][0]) === 'ALL') {
          closedMonths.push(String(data[i][1]));
        }
      }
    }
    
    const targetYear = parseInt(year) || 2024;
    const result = [];
    for (let i = 0; i < 12; i++) {
      let y = targetYear;
      let m = 10 + i;
      if (m > 12) { y += 1; m -= 12; }
      const ym = `${y}${m.toString().padStart(2, '0')}`;
      result.push({ ym: ym, isClosed: closedMonths.includes(ym) });
    }
    return result;
  } catch (e) { return { error: e.message }; }
}

function updateCloseStatus(ym, lockParam) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openById(MAIN_SS_ID);
    let sheet = ss.getSheetByName('締め') || ss.insertSheet('締め');
    const data = sheet.getDataRange().getValues();
    
    let updated = false;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]) === ym && String(data[i][0]) === 'ALL') {
        sheet.getRange(i + 1, 3).setValue(lockParam);
        sheet.getRange(i + 1, 4).setValue(new Date());
        updated = true;
        break;
      }
    }
    
    if (!updated && lockParam) {
      sheet.appendRow(['ALL', ym, true, new Date()]);
      sheet.getRange(sheet.getLastRow(), 3).insertCheckboxes();
    }
    
    const actionStr = lockParam ? '締め（ロック）' : 'ロック解除';
    logActionToSheet('updateCloseStatus', `【月次締め処理】${ym.substring(0,4)}年${ym.substring(4,6)}月度を ${actionStr} しました。`);
    return { success: true };
  } catch (e) { 
    return { error: "他ユーザーと操作が重なりました。画面を更新して再度お試しください。" }; 
  } finally {
    lock.releaseLock();
  }
}