// =========================================
// SPenD Dashboard - Backend Script (Data API)
// =========================================

function getSS() {
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch(e) {}
  if (!ss) {
    try { if (typeof MAIN_SS_ID !== 'undefined') ss = SpreadsheetApp.openById(MAIN_SS_ID); } catch(e) {}
  }
  return ss;
}

function getSystemSettings() {
  try {
    const props = PropertiesService.getScriptProperties().getProperties();
    return props;
  } catch (e) {
    return {};
  }
}

function getInitialData() {
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

    let activeEmail = "";
    try { activeEmail = String(Session.getActiveUser().getEmail()).toLowerCase(); } catch(e) {}

    return { segments, teamSettings, currentYear, users, activeEmail, actualSettings: {}, subjectSettings: {}, masters: { subject: [] } };
  } catch (e) { return { error: e.message }; }
}

function getFiscalYear(ym) {
  if (!ym || ym.length < 6) return parseInt(new Date().getFullYear());
  const y = parseInt(ym.substring(0, 4));
  const m = parseInt(ym.substring(4, 6));
  if (m >= 10 && m <= 12) return y;
  return y - 1; 
}

function getRawDataPackage(targetPeriod, targetTeamCd) {
  try {
    const config = getSystemSettings();
    const ss = getSS();
    if (!ss) return { error: "SS取得失敗" };
    
    const zaimuKeywordsStr = String(config['ZAIMU_KEYWORDS'] || '');
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
    } else {
      diffFrom = Number(config['ZAIMU_DIFF_FROM']) || 0;
      diffTo = Number(config['ZAIMU_DIFF_TO']) || 9999999999;
      if (config['ZAIMU_SETTINGS']) { try { const zs = JSON.parse(config['ZAIMU_SETTINGS']); if(zs.subjects) zaimuSubjects = Array.isArray(zs.subjects) ? zs.subjects : Object.keys(zs.subjects).filter(k => zs.subjects[k]); } catch(e){} }
    }
    const zaimuSettings = { subjects: zaimuSubjects, diffFrom: diffFrom, diffTo: diffTo };
    let weatherSettings = { sunny: 5, cloudy: 10, rainy: 20 };
    if (config['WEATHER_SETTINGS']) { try { weatherSettings = JSON.parse(config['WEATHER_SETTINGS']); } catch(e){} }

    let email = "";
    try { email = String(Session.getActiveUser().getEmail()).toLowerCase(); } catch(e) {}
    
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
    targetTeamSettings.forEach(t => { if (!isAdmin) { t.segments = t.segments.filter(seg => allowedSegments.includes(seg)); } });

    const segToTeam = {};
    teamSettings.forEach(t => { t.segments.forEach(segCd => { if (isAdmin || allowedSegments.includes(segCd)) { segToTeam[segCd] = { cd: t.teamCd, name: t.teamName, isHidden: t.isHidden }; } }); });
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
      for (let i = 1; i < subSetData.length; i++) { if (String(subSetData[i][0])) { subjectSettings[String(subSetData[i][0]).trim()] = (subSetData[i][1] === true || String(subSetData[i][1]).toUpperCase() === 'TRUE'); } }
    }

    const cmtSheetName = String(config['COMMENT_SHEET_NAME'] || 'htmlコメント保管').trim();
    const repCmtSheetName = String(config['REPORT_COMMENT_SHEET_NAME'] || 'html報告資料コメント保管').trim();

    const savedComments = {};
    const cmtSheet = ss.getSheetByName(cmtSheetName);
    if(cmtSheet && cmtSheet.getLastRow() > 0) {
      const commentData = cmtSheet.getDataRange().getValues();
      for (let i = 1; i < commentData.length; i++) { 
         const cmtPeriod = String(commentData[i][0]);
         savedComments[`${cmtPeriod}_${commentData[i][1]}_${commentData[i][2]}`] = String(commentData[i][3]); 
      }
    }
    
    const reportComments = {}; 
    const repCmtSheet = ss.getSheetByName(repCmtSheetName);
    if(repCmtSheet && repCmtSheet.getLastRow() > 0) {
      const repCommentData = repCmtSheet.getDataRange().getValues();
      for (let i = 1; i < repCommentData.length; i++) { 
         reportComments[`${String(repCommentData[i][0])}_${String(repCommentData[i][1])}`] = String(repCommentData[i][2]);
      }
    }

    const labelBudget = String(config['LABEL_BUDGET'] || "予算").trim();
    const labelQuick = String(config['LABEL_ACTUAL_QUICK'] || "実績(速報)").trim();
    const labelFixed = String(config['LABEL_ACTUAL_FIXED'] || "実績(確定)").trim();
    
    // ==========================================
    // 💡 ユーザー様発案！最強の「D列完全一致」ロジック
    // ==========================================
    let targetFY = parseInt(config['CURRENT_YEAR']) || new Date().getFullYear(); 
    let reportTargetMonths = [];
    let foundInMaster = false;
    
    const targetVal = String(targetPeriod).trim(); // 画面から飛んできた値（例: "2025_3Q"）
    
    const masterSheet = ss.getSheetByName('マスタ_年度');
    if (masterSheet) {
      const mData = masterSheet.getDataRange().getValues();
      
      // i=1 からループ (1行目はヘッダー想定)
      for (let i = 1; i < mData.length; i++) {
        const colA = String(mData[i][0]).trim(); // A列: 表示名
        const colD = String(mData[i][3] || "").trim(); // D列: システム値 (2025_3Qなど)
        
        // D列の値と完全に一致するかチェック！ (D列が空の時のためにA列でもチェック)
        if (colD === targetVal || colA === targetVal) {
          targetFY = parseInt(String(mData[i][1]).replace(/\D/g, ''), 10);
          const monthsStr = String(mData[i][2]);
          reportTargetMonths = monthsStr.split(',').map(m => m.trim().replace(/\D/g, '')).filter(m => m.length === 6);
          foundInMaster = true;
          break;
        }
      }
    }

    // マスタに設定がない場合（単月など）の自動フェイルセーフ
    if (!foundInMaster) {
      const cleanYm = targetVal.replace(/\D/g, ''); 
      if (cleanYm.length >= 6) {
        reportTargetMonths = [cleanYm.substring(0, 6)];
        targetFY = getFiscalYear(reportTargetMonths[0]); 
      }
    }
    // ==========================================

    const targetMonths = []; const prevTargetMonths = [];
    for (let i = 0; i < 12; i++) { 
      let m = 10 + i;
      let y = targetFY;
      if (m > 12) {
         m -= 12;
         y += 1; 
      }
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
      const cacheSheet = ss.getSheetByName(sheetName); if (!cacheSheet) return;
      const data = cacheSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const ym = String(data[i][1]).trim();
        if (!allValidMonths.includes(ym)) continue;
        
        const seg = String(data[i][2]).trim(); 
        const tInfo = segToTeam[seg]; 
        if (!tInfo) continue; 
        
        const rowType = String(data[i][0]).trim();
        let useRow = false; 
        let outputType = '実績';
        
        if (rowType === '予算' || rowType === labelBudget || rowType.includes('予算')) { 
          useRow = true;
          outputType = '予算';
        } else {
          const actSet = String(actualSettings[ym] || '確定');
          if (actSet.includes('速報') && (rowType === labelQuick || rowType.includes('速報'))) {
            useRow = true;
            outputType = '速報実績';
          } else if (actSet.includes('確定') && (rowType === labelFixed || rowType.includes('確定'))) {
            useRow = true;
            outputType = '確定実績';
          } else if (rowType === '実績' || rowType === labelFixed) {
            useRow = true;
            outputType = '実績';
          }
        }

        if (!useRow) continue;
        const rowObj = { month: ym, teamCd: tInfo.cd, teamName: tInfo.name, segCd: seg, segName: seg, acctCd: String(data[i][3]), acctName: String(data[i][4]), partnerName: String(data[i][6]), amount: Number(data[i][7]), type: outputType };
        if (isCurrentYear) {
          if (outputType === '予算') annualBudgetList.push(rowObj);
          if (reportTargetMonths.includes(ym)) currentMonthAllSegments.push(rowObj);
          if (ytdMonths.includes(ym)) ytdAllSegments.push(rowObj);
          if ((targetTeamCd === 'all' || (tInfo.cd === targetTeamCd)) && targetMonths.includes(ym)) { 
            if (outputType === '予算') budgetList.push(rowObj);
            else actualList.push(rowObj); 
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
    }
    
    extractData('DATA_CACHE_FY' + targetFY, true);
    extractData('DATA_CACHE_FY' + (targetFY - 1), false);
    
    const closedStatus = {};
    const closeSheet = ss.getSheetByName('締め');
    if (closeSheet) {
      const closeData = closeSheet.getDataRange().getValues();
      for (let i = 1; i < closeData.length; i++) {
        if (closeData[i][2] === true) closedStatus[`${closeData[i][0]}_${closeData[i][1]}`] = true;
      }
    }

    function calcSokai(dataList) {
      const tStats = {};
      const perfs = [];
      dataList.forEach(row => {
        if(row.isPrev) return;
        const tCd = row.teamCd;
        if(!tStats[tCd]) tStats[tCd] = { name: row.teamName, budget: 0, actual: 0, items: {} };
        
        const isBudget = row.type === '予算';
        const isActual = row.type !== '予算';
        
        tStats[tCd].budget += isBudget ? row.amount : 0;
        tStats[tCd].actual += isActual ? row.amount : 0;
        if(!tStats[tCd].items[row.acctName]) tStats[tCd].items[row.acctName] = { acctName: row.acctName, budget: 0, actual: 0 };
        tStats[tCd].items[row.acctName].budget += isBudget ? row.amount : 0;
        tStats[tCd].items[row.acctName].actual += isActual ? row.amount : 0;
      });
      const ranking = [];
      Object.values(tStats).forEach(t => {
        if(t.budget === 0 && t.actual === 0) return;
        t.diffRatio = t.budget > 0 ? Math.abs(t.actual - t.budget) / t.budget : (t.actual > 0 ? 1 : 0);
        const itemArr = [];
        Object.values(t.items).forEach(item => {
           if(item.budget > 0 && item.budget === item.actual) perfs.push({ teamName: t.name, acctName: item.acctName, amount: item.actual });
           item.diff = item.actual - item.budget; item.diffAbs = Math.abs(item.diff);
           item.ratio = t.budget > 0 ? (item.budget / t.budget) : 0;
           if(item.budget > 0 || item.actual > 0) itemArr.push(item);
        });
        itemArr.sort((a,b) => b.diffAbs - a.diffAbs);
        t.worst3 = itemArr.slice(0, 3);
        const rest = itemArr.slice(3);
        
        if(rest.length > 0) { t.others = rest.reduce((acc, curr) => { acc.budget += curr.budget; acc.actual += curr.actual; acc.diff += curr.diff; acc.ratio += curr.ratio; return acc; }, { acctName: 'その他', budget: 0, actual: 0, diff: 0, ratio: 0 });
        } else { t.others = null; }
        delete t.items; ranking.push(t);
      });
      ranking.sort((a,b) => a.diffRatio - b.diffRatio);
      return { ranking, perfects: perfs };
    }
    
    const sokaiPeriod = calcSokai(currentMonthAllSegments);
    const sokaiYTD = calcSokai(ytdAllSegments);
    
    const filteredComments = {};
    Object.keys(savedComments).forEach(k => {
       const parts = k.split('_'); 
       if (reportTargetMonths.includes(parts[0])) {
           const targetKey = `${parts[1]}_${parts.slice(2).join('_')}`;
           if(filteredComments[targetKey]) {
               filteredComments[targetKey] += `|||${savedComments[k]}`;
           } else {
               filteredComments[targetKey] = savedComments[k];
           }
       }
    });

    return { 
      segments: [], teamSettings: targetTeamSettings, masters, 
      comments: filteredComments, reportComments: reportComments, 
      subjectSettings, actualSettings, closedStatus, 
      cacheData: { budget: budgetList, actuals: actualList, prevActuals: prevActualList, currentMonthAllSegments, ytdAllSegments, annualBudget: annualBudgetList }, 
      params: { period: targetPeriod, reportTargetMonths, ytdMonths, monthList: targetMonths, segmentCd: targetTeamCd }, 
      zaimuKeywords: zaimuKeywordsStr, zaimuSettings, weatherSettings,
      rankingPeriod: sokaiPeriod.ranking, perfectPeriod: sokaiPeriod.perfects,
      rankingYTD: sokaiYTD.ranking, perfectYTD: sokaiYTD.perfects
    };
  } catch (e) { return { error: e.message }; }
}