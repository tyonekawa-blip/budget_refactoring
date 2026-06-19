// =========================================
// SPenD Dashboard - Backend Script (Data API)
// =========================================

function getSS() {
  const TARGET_SSID = '1MC933tAVIxMIPX4VR68GNKcmC-VtTp0fbUGOofAVeb8';
  try { return SpreadsheetApp.openById(TARGET_SSID); } 
  catch(e) { try { return SpreadsheetApp.getActiveSpreadsheet(); } catch(err) { return null; } }
}

function getSystemSettings() {
  try { return PropertiesService.getScriptProperties().getProperties(); } catch (e) { return {}; }
}

function getAppInitialData() {
  try {
    const ss = getSS();
    if (!ss) return { error: "SS取得失敗" };
    const segSheet = ss.getSheetByName('マスタ_セグメント');
    let segments = [];
    if (segSheet) {
      const segData = segSheet.getDataRange().getValues();
      segments = segData.slice(2).map(r => ({ cd: String(r[7]).trim(), name: String(r[8]).trim() })).filter(r => r.cd && r.name && r.cd !== 'seg1CD');
    }

    const config = getSystemSettings();
    const currentYear = config['CURRENT_YEAR'] || '2025'; 

    let userSheet = ss.getSheetByName('ユーザー設定');
    const users = [];
    if (userSheet) {
      const userData = userSheet.getDataRange().getValues();
      for (let i = 1; i < userData.length; i++) {
        const email = String(userData[i][0]).trim().toLowerCase();
        if (email) users.push({ email, role: String(userData[i][1]).toLowerCase(), menus: String(userData[i][2]).split(','), segments: String(userData[i][3]).split(','), name: String(userData[i][4] || ""), status: String(userData[i][5] || "有効"), memo: String(userData[i][6] || "") });
      }
    }

    const teamSettings = [];
    const teamSheetName = typeof TEAM_SHEET_NAME !== 'undefined' ? TEAM_SHEET_NAME : 'マスタ_チーム';
    let teamSheet = ss.getSheetByName(teamSheetName);
    if (teamSheet && teamSheet.getLastRow() > 1) {
      const teamData = teamSheet.getDataRange().getValues();
      for (let i = 1; i < teamData.length; i++) {
        if (String(teamData[i][0]).trim()) {
          teamSettings.push({ teamCd: String(teamData[i][0]).trim(), teamName: String(teamData[i][1]).trim(), order: Number(teamData[i][2]) || 999, segments: String(teamData[i][3]).split(',').map(s=>s.trim()).filter(s=>s), isHidden: (String(teamData[i][4]).trim() === 'true' || teamData[i][4] === true) });
        }
      }
      teamSettings.sort((a, b) => a.order - b.order);
    }

    const periods = [];
    const masterSheet = ss.getSheetByName('マスタ_年度');
    if (masterSheet) {
      const mData = masterSheet.getDataRange().getDisplayValues();
      for (let i = 1; i < mData.length; i++) {
        const label = String(mData[i][0]).trim();
        const fy = String(mData[i][1]).trim();
        const val = String(mData[i][3] || label).trim(); 
        if (label) periods.push({ label: label, value: val, fy: fy });
      }
    }

    let activeEmail = "";
    try { activeEmail = String(Session.getActiveUser().getEmail()).toLowerCase(); } catch(e) {}

    return { segments, teamSettings, currentYear, users, activeEmail, actualSettings: {}, subjectSettings: {}, masters: { subject: [] }, periods: periods };
  } catch (e) { return { error: e.message }; }
}

function getFiscalYear(ym) {
  if (!ym || ym.length < 6) return parseInt(new Date().getFullYear());
  const y = parseInt(ym.substring(0, 4));
  const m = parseInt(ym.substring(4, 6));
  if (m >= 10 && m <= 12) return y;
  return y - 1; 
}

function getAppRawDataPackage(targetPeriod, targetTeamCd) {
  let debugLogs = [];
  debugLogs.push(`🚀 [GAS開始] 選択された期間: '${targetPeriod}', チーム: '${targetTeamCd}'`);
  try {
    const config = getSystemSettings();
    const ss = getSS();
    if (!ss) return { error: "SS取得失敗" };
    
    const zaimuKeywordsStr = String(config['ZAIMU_KEYWORDS'] || '');
    let zaimuSubjects = [];
    let diffFrom = 0; let diffTo = 9999999999;
    const zaimuSheet = ss.getSheetByName('html重要科目・表示設定');
    if (zaimuSheet && zaimuSheet.getLastRow() > 0) {
      const zData = zaimuSheet.getDataRange().getValues();
      for (let i = 0; i < zData.length; i++) {
        const key = String(zData[i][0]).trim(); const val = zData[i][1];
        if (key === 'ZAIMU_DIFF_FROM') diffFrom = Number(val) || 0;
        else if (key === 'ZAIMU_DIFF_TO') diffTo = Number(val) || 9999999999;
        else if (key && key !== '項目名' && key !== '勘定科目名' && (val === true || String(val).toUpperCase() === 'TRUE')) zaimuSubjects.push(key);
      }
    }
    const zaimuSettings = { subjects: zaimuSubjects, diffFrom, diffTo };
    let weatherSettings = { sunny: 5, cloudy: 10, rainy: 20 };
    if (config['WEATHER_SETTINGS']) try { weatherSettings = JSON.parse(config['WEATHER_SETTINGS']); } catch(e){}

    let email = ""; try { email = String(Session.getActiveUser().getEmail()).toLowerCase(); } catch(e) {}
    let currentUser = { role: 'user', segments: [] };
    const userSheet = ss.getSheetByName('ユーザー設定');
    if (userSheet) {
      const userData = userSheet.getDataRange().getValues();
      for(let i = 1; i < userData.length; i++) {
        if (String(userData[i][0]).trim().toLowerCase() === email) {
          currentUser.role = String(userData[i][1]).toLowerCase();
          currentUser.segments = String(userData[i][3]).split(',').map(s=>s.trim());
          break;
        }
      }
    }
    const isAdmin = currentUser.role === 'admin';
    const allowedSegments = currentUser.segments;

    const masters = { subject: ss.getSheetByName('マスタ_勘定科目') ? ss.getSheetByName('マスタ_勘定科目').getDataRange().getValues() : [] };
    const teamSheetName = typeof TEAM_SHEET_NAME !== 'undefined' ? TEAM_SHEET_NAME : 'マスタ_チーム';
    const teamSheet = ss.getSheetByName(teamSheetName);
    let teamSettings = teamSheet ? teamSheet.getDataRange().getValues().slice(1).map(r => ({ teamCd: String(r[0]), teamName: String(r[1]), segments: String(r[3]).split(',').map(s=>s.trim()), isHidden: (String(r[4]).trim() === 'true' || r[4] === true) })) : [];
    const targetTeamSettings = targetTeamCd === 'all' ? teamSettings : teamSettings.filter(t => t.teamCd === targetTeamCd);
    targetTeamSettings.forEach(t => { if (!isAdmin) t.segments = t.segments.filter(seg => allowedSegments.includes(seg)); });

    const segToTeam = {};
    teamSettings.forEach(t => { t.segments.forEach(segCd => { if (isAdmin || allowedSegments.includes(segCd)) segToTeam[segCd] = { cd: t.teamCd, name: t.teamName, isHidden: t.isHidden }; }); });
    const actualSettings = {};
    const actSetSheet = ss.getSheetByName('html実績表示設定保管');
    if (actSetSheet && actSetSheet.getLastRow() > 0) {
      const actSetData = actSetSheet.getDataRange().getValues();
      for (let i = 1; i < actSetData.length; i++) { const ymKey = String(actSetData[i][0]).trim(); if (ymKey) actualSettings[ymKey] = String(actSetData[i][1]).trim(); }
    }

    const subjectSettings = {}; 
    const subSetSheet = ss.getSheetByName('html科目設定保管');
    if (subSetSheet && subSetSheet.getLastRow() > 0) {
      const subSetData = subSetSheet.getDataRange().getValues();
      for (let i = 1; i < subSetData.length; i++) { if (String(subSetData[i][0])) subjectSettings[String(subSetData[i][0]).trim()] = (subSetData[i][1] === true || String(subSetData[i][1]).toUpperCase() === 'TRUE'); }
    }

    let targetFY = parseInt(config['CURRENT_YEAR']) || new Date().getFullYear(); 
    let reportTargetMonths = [];
    const targetVal = String(targetPeriod).trim(); 
    const masterSheetForQuery = ss.getSheetByName('マスタ_年度');
    
    if (masterSheetForQuery) {
      const mData = masterSheetForQuery.getDataRange().getDisplayValues();
      for (let i = 1; i < mData.length; i++) {
        const colA = String(mData[i][0]).trim();
        const colC = String(mData[i][2]).trim(); 
        const colD = String(mData[i][3] || "").trim(); 
        
        if (colD === targetVal || colC === targetVal || colA === targetVal) {
          debugLogs.push(`✅ [マスタ合致] 行${i+1}で一致! A='${colA}', C='${colC}', D='${colD}'`);
          const fyRaw = String(mData[i][1]).replace(/\D/g, '');
          if (fyRaw.length >= 4) targetFY = parseInt(fyRaw.substring(0,4), 10);
          
          let rawMonths = colC;
          if (/[\|｜,，]/.test(colD) || colD.length >= 6) rawMonths = colD;
          if (/[\|｜,，]/.test(colC) || colC.length >= 6) rawMonths = colC; 
          
          reportTargetMonths = rawMonths.split(/[\|｜,，]+/).map(m => m.replace(/\D/g, '')).filter(m => m.length === 6);
          debugLogs.push(`📅 [月リスト抽出] 対象月: ${reportTargetMonths.join(', ')} (判定年度: FY${targetFY})`);
          break;
        }
      }
    }

    if (reportTargetMonths.length === 0) {
      debugLogs.push(`⚠️ [警告] マスタに合致する期間がありませんでした。`);
      const cleanYm = targetVal.replace(/\D/g, ''); 
      if (cleanYm.length === 6) {
        reportTargetMonths = [cleanYm];
        let y = parseInt(cleanYm.substring(0, 4), 10); let m = parseInt(cleanYm.substring(4, 6), 10);
        targetFY = (m >= 10 && m <= 12) ? y : y - 1;
      }
    }

    const targetMonths = []; const prevTargetMonths = [];
    for (let i = 0; i < 12; i++) { 
      let m = 10 + i; let y = targetFY;
      if (m > 12) { m -= 12; y += 1; }
      targetMonths.push(`${y}${m.toString().padStart(2, '0')}`);
      prevTargetMonths.push(`${y - 1}${m.toString().padStart(2, '0')}`);
    }

    let ytdMonths = [];
    if (reportTargetMonths.length > 0) {
      const lastTargetMonth = reportTargetMonths[reportTargetMonths.length - 1];
      for (let ym of targetMonths) { ytdMonths.push(ym); if (ym === lastTargetMonth) break; }
    }
    const allValidMonths = [...targetMonths, ...prevTargetMonths];
    const budgetList = [], actualList = [], prevActualList = [], currentMonthAllSegments = [], ytdAllSegments = [], annualBudgetList = [];
    
    function extractData(sheetName, isCurrentYear) {
      debugLogs.push(`📂 [シート検索] '${sheetName}' を探しています...`);
      const cacheSheet = ss.getSheetByName(sheetName); 
      if (!cacheSheet) {
        debugLogs.push(`❌ [エラー] シート '${sheetName}' が見つかりません！`);
        return;
      }
      const data = cacheSheet.getDataRange().getValues();
      let matchCount = 0;
      
      for (let i = 1; i < data.length; i++) {
        let ym = "";
        if (data[i][1] instanceof Date) {
          const d = data[i][1]; ym = d.getFullYear() + String(d.getMonth() + 1).padStart(2, '0');
        } else {
          ym = String(data[i][1]).replace(/\D/g, '').substring(0, 6); 
        }

        if (ym.length !== 6 || !allValidMonths.includes(ym)) continue;
        const seg = String(data[i][2]).trim(); 
        const tInfo = segToTeam[seg]; 
        if (!tInfo) continue; 
        
        matchCount++;
        const rowType = String(data[i][0]).trim();
        let useRow = false; let outputType = '実績';
        
        if (rowType === '予算' || rowType.includes('予算')) { 
          useRow = true; outputType = '予算';
        } else {
          useRow = true; outputType = '実績'; 
        }

        if (!useRow) continue;
        const rowObj = { month: ym, teamCd: tInfo.cd, teamName: tInfo.name, segCd: seg, segName: seg, acctCd: String(data[i][3]), acctName: String(data[i][4]), partnerName: String(data[i][6]), amount: Number(data[i][7]), type: outputType };
        
        if (isCurrentYear) {
          if (outputType === '予算') annualBudgetList.push(rowObj);
          if (reportTargetMonths.includes(ym)) currentMonthAllSegments.push(rowObj);
          if (ytdMonths.includes(ym)) ytdAllSegments.push(rowObj);
          if ((targetTeamCd === 'all' || (tInfo.cd === targetTeamCd)) && targetMonths.includes(ym)) { 
            if (outputType === '予算') budgetList.push(rowObj); else actualList.push(rowObj); 
          }
        } else {
          const currCorrespondingYm = `${parseInt(ym.substring(0,4))+1}${ym.substring(4,6)}`;
          const prevRowObj = { ...rowObj, month: currCorrespondingYm, isPrev: true };
          if (reportTargetMonths.includes(currCorrespondingYm)) currentMonthAllSegments.push(prevRowObj);
          if (ytdMonths.includes(currCorrespondingYm)) ytdAllSegments.push(prevRowObj);
          if ((targetTeamCd === 'all' || (tInfo.cd === targetTeamCd)) && targetMonths.includes(currCorrespondingYm)) {
            if (outputType !== '予算') prevActualList.push(prevRowObj);
          }
        }
      }
      debugLogs.push(`📊 [抽出完了] '${sheetName}' から ${matchCount} 件の有効データを取得しました。`);
    }
    
    extractData('DATA_CACHE_FY' + targetFY, true);
    extractData('DATA_CACHE_FY' + (targetFY - 1), false);
    
    const closedStatus = {};
    const sokaiPeriod = { ranking: [], perfects: [] }; 
    const sokaiYTD = { ranking: [], perfects: [] }; 

    return { 
      debugLogs: debugLogs, 
      segments: [], teamSettings: targetTeamSettings, masters, 
      comments: {}, reportComments: {}, 
      subjectSettings, actualSettings, closedStatus, 
      cacheData: { budget: budgetList, actuals: actualList, prevActuals: prevActualList, currentMonthAllSegments, ytdAllSegments, annualBudget: annualBudgetList }, 
      params: { period: targetPeriod, reportTargetMonths, ytdMonths, monthList: targetMonths, segmentCd: targetTeamCd }, 
      zaimuKeywords: zaimuKeywordsStr, zaimuSettings, weatherSettings,
      rankingPeriod: sokaiPeriod.ranking, perfectPeriod: sokaiPeriod.perfects,
      rankingYTD: sokaiYTD.ranking, perfectYTD: sokaiYTD.perfects
    };
  } catch (e) { return { error: e.message, debugLogs: debugLogs }; }
}

// =========================================
// 🚀 魔法のスライド連携関数 (Route B: 画像化)
// =========================================
function exportDashboardToSlides(imageDataArray, slideId) {
  try {
    const presentation = SlidesApp.openById(slideId);
    
    // オプション: 既存のスライドをリセットしたい場合は以下のコメントアウトを外す
    // const slides = presentation.getSlides();
    // if(slides.length > 1) { for(let i=1; i<slides.length; i++){ slides[i].remove(); } }
    
    imageDataArray.forEach(data => {
      // 画像の種類(JPEG/PNG)を動的に判別し、Blobを生成
      const mimeType = data.image.match(/data:(.*?);/)[1] || 'image/jpeg';
      const base64Data = data.image.split(',')[1];
      const ext = mimeType === 'image/png' ? '.png' : '.jpg';
      
      const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, data.teamName + ext);
      
      const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
      const image = slide.insertImage(blob);
      
      // 画像をスライドの中央にピッタリ合わせる計算
      const pageWidth = presentation.getPageWidth();
      const pageHeight = presentation.getPageHeight();
      const imgWidth = image.getWidth();
      const imgHeight = image.getHeight();
      
      const ratio = Math.min(pageWidth / imgWidth, pageHeight / imgHeight);
      const newWidth = imgWidth * ratio;
      const newHeight = imgHeight * ratio;
      
      image.setWidth(newWidth);
      image.setHeight(newHeight);
      image.setLeft((pageWidth - newWidth) / 2);
      image.setTop((pageHeight - newHeight) / 2);
    });

    return { success: true, url: presentation.getUrl() };
  } catch (e) {
    return { error: e.message };
  }
}

function authorizeSlides() {
  SlidesApp.openById('1IrJGgwbPIy6uc8lhylHXzDTflQxnWSFKdfP0ADCpJl8');
  console.log("スライドへのアクセス権限が許可されました！");
}