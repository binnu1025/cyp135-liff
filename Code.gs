// ==========================================
// 1. 核心設定區
// ==========================================
const APP_VERSION = '2.03';
const CHANNEL_ACCESS_TOKEN = 'J0kKhBXygzwYubPguLX7yObD2nX2/V5CgTCmsHYlr9fLWhHouO4PAJHUTRoMgR40w9qPcAdWvnt3l4vo5cWd9n22uj48ZX9lFWeNCNw/bSqkx3ruVqOdDo6xTCH0Ivxmd68tyGo1ReJhz8RLwbo5CQdB04t89/1O/w1cDnyilFU=';
const SPREADSHEET_ID = '1U7F0L6WvZuF-71UfWQltoCrKXfgAtyocZmOOWCe68jc';
const DRIVE_FOLDER_ID = '1zNJvKi7uknsSG1eW2NYu8XX9R-0DZHBh';

// ==========================================
// 2. 請求入口 (doGet & doPost)
// ==========================================

function doGet(e) {
  try {
    const action = e.parameter.action;
    const uid = e.parameter.uid;
    if (action === 'getInitData')        return responseJson({ success: true, data: apiGetInitData(uid) });
    if (action === 'getAdminDashboard')  return responseJson({ success: true, data: apiGetAdminDashboard(uid) });
    if (action === 'checkTarId')         return responseJson(apiCheckTarId(e.parameter.tarId));
    if (action === 'getGiftData')        return responseJson({ success: true, data: apiGetGiftData(uid) });
    if (action === 'getContestStatus')   return responseJson({ success: true, data: apiGetContestStatus() });
    if (action === 'checkRegisterId')    return responseJson(apiCheckRegisterId(e.parameter.amId));
    if (action === 'getGroupList')        return responseJson({ success: true, data: apiGetGroupList() });
    if (action === 'checkMyRegistration') return responseJson(apiCheckMyRegistration(uid));
    return responseJson({ success: false, message: "未知的 GET 請求" });
  } catch (err) {
    return responseJson({ success: false, message: "後端錯誤：" + err.toString() });
  }
}

function doPost(e) {
  try {
    const postData = JSON.parse(e.postData.contents);
    if (postData.events && postData.events.length > 0) {
      postData.events.forEach(handleLineEvent);
      return ContentService.createTextOutput("ok");
    }
    let result;
    const { action, payload } = postData;
    if      (action === 'submitReferral')    result = apiSubmitReferral(payload);
    else if (action === 'approve')           result = apiProcessReview(payload, "通過");
    else if (action === 'reject')            result = apiProcessReview(payload, "退回");
    else if (action === 'linkIdentity')      result = apiLinkIdentity(payload);
    else if (action === 'registerUser')      result = apiRegisterUser(payload);
    else if (action === 'assignGift')        result = apiAssignGift(payload);
    else if (action === 'confirmGift')       result = apiConfirmGift(payload);
    else if (action === 'addGift')           result = apiAddGift(payload);
    else if (action === 'deleteGift')        result = apiDeleteGift(payload);
    else if (action === 'setContestStatus')  result = apiSetContestStatus(payload);
    else if (action === 'submitRegister')    result = apiSubmitRegister(payload);
    else if (action === 'approveRegister')   result = apiApproveRegister(payload);
    else if (action === 'rejectRegister')    result = apiRejectRegister(payload);
    else result = { success: false, message: "無效的 POST Action" };
    return responseJson(result);
  } catch(err) {
    return responseJson({ success: false, message: "系統異常：" + err.toString() });
  }
}

// ==========================================
// 3. LINE Webhook 處理邏輯
// ==========================================

function handleLineEvent(event) {
  if (event.type !== 'message' || event.message.type !== 'text') return;
  const userMsg = event.message.text.trim();
  const userId = event.source.userId;
  const replyToken = event.replyToken;

  if (userMsg === '報名' || userMsg === '參加競賽') {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const rosterData = ss.getSheetByName('雁群總名冊').getDataRange().getValues();
    if (isUserRegistered(userId, rosterData)) {
      replyLine(replyToken, "⚠️ 您已在名冊中，無法重複報名。");
      return;
    }
    const copyText = `📝註冊申請\n安麗編號：\n👬請輸入完整姓名：`;
    replyLine(replyToken, `歡迎參賽！請複製並填寫以下內容傳回：\n\n${copyText}`);
    return;
  }

  if (userMsg.includes('📝註冊申請')) {
    try {
      const lines = userMsg.split('\n');
      let amId = "", name = "";
      lines.forEach(line => {
        const cleanLine = line.replace(/\s/g, '');
        if (cleanLine.includes('編號')) {
          const p = cleanLine.includes('：') ? cleanLine.split('：') : cleanLine.split(':');
          if (p[1]) amId = p[1].trim();
        }
        if (cleanLine.includes('姓名')) {
          const p = cleanLine.includes('：') ? cleanLine.split('：') : cleanLine.split(':');
          if (p[1]) name = p[1].trim();
        }
      });
      if (!amId || !name) throw new Error("解析失敗");
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      const rosterData = ss.getSheetByName('雁群總名冊').getDataRange().getValues();
      if (isAmIdRegistered(amId, rosterData)) {
        replyLine(replyToken, `⚠️ 報名失敗！編號「${amId}」已在名冊中。`);
        return;
      }
      sendGroupFlex(replyToken, amId, name);
    } catch(e) {
      replyLine(replyToken, "❌ 格式讀取失敗，請確保冒號後有填寫內容。");
    }
    return;
  }

  if (userMsg.includes('【確認報名】')) {
    try {
      const data = userMsg.split('|');
      const amId = data[0].split(':')[1];
      const name = data[1].split(':')[1];
      const group = data[2].split(':')[1];
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      const rosterSheet = ss.getSheetByName('雁群總名冊');
      const rosterData = rosterSheet.getDataRange().getValues();
      if (isAmIdRegistered(amId, rosterData) || isUserRegistered(userId, rosterData)) {
        replyLine(replyToken, "⚠️ 您已完成報名，無法再變更雁群。如需修改請聯絡管理員。");
        return;
      }
      rosterSheet.appendRow([userId, amId, name, group, "待確認"]);
      replyLine(replyToken, `📋 ${name}，您的報名申請已送出！請等待會長確認後即可使用系統。`);
    } catch(e) {
      replyLine(replyToken, "❌ 同步至系統失敗。");
    }
  }
}

// ==========================================
// 4. Web API 核心邏輯
// ==========================================

function apiGetInitData(uid) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const uidStr = String(uid).trim();
  // 系統設定只讀一次，門檻與競賽狀態共用（原本會各讀一次、且各自再開一次試算表）
  const settings = getSettings(ss);
  const effAmt = settings.effectiveAmount;

  const permSheet   = ss.getSheetByName('LINE權限表');
  const rosterSheet = ss.getSheetByName('雁群總名冊');
  const groupSheet  = ss.getSheetByName('雁群清單');
  const totalSheet  = ss.getSheetByName('推薦總表');

  const permData   = permSheet.getDataRange().getValues();
  const rosterData = rosterSheet.getDataRange().getValues();

  let authorizedGroups = [];
  let isPresident = false;
  let myClassroom = "";

  for (let i = 1; i < permData.length; i++) {
    if (String(permData[i][0]).trim() === uidStr) {
      const groupStr = String(permData[i][2] || "").replace(/，/g, ',').trim();
      if (groupStr) authorizedGroups = groupStr.split(',').map(g => g.trim()).filter(g => g !== "");
      if (String(permData[i][3]).trim() === "會長") isPresident = true;
      myClassroom = String(permData[i][4] || "").trim();
      break;
    }
  }

  const maps = getMapsFromData(rosterData);
  const myAmId = maps.uidToAmId[uidStr];
  const myGroup = myAmId ? maps.amIdToGroup[myAmId] : "";

  let groupList = [];
  const groupToClassroom = {};
  if (groupSheet.getLastRow() >= 2) {
    const groupData = groupSheet.getRange(2, 1, groupSheet.getLastRow() - 1, 2).getValues();
    groupData.forEach(r => {
      const gName = String(r[0]).trim();
      const cName = String(r[1] || "").trim();
      if (gName) {
        groupList.push(gName);
        if (cName) groupToClassroom[gName] = cName;
      }
    });
  }

  if (!myClassroom && myGroup) myClassroom = groupToClassroom[myGroup] || "";
  if (authorizedGroups.length === 0 && myGroup) authorizedGroups = [myGroup];
  if (isPresident && authorizedGroups.length === 0) authorizedGroups = groupList;
  if (!isPresident && authorizedGroups.length === 0 && !myAmId)
    return { error: "未註冊", groupList: groupList };

  const participants = [];
  rosterData.slice(1).forEach(r => {
    const rGroup = String(r[3]).trim();
    const amId = String(r[1]).trim();
    const status = String(r[4] || "").trim();
    if (authorizedGroups.includes(rGroup) && amId && status === "是") {
      participants.push({ id: amId, name: String(r[2]).trim(), group: rGroup });
    }
  });

  const groupPerfMap = {};
  const newbiesForCons = [];

  if (totalSheet && totalSheet.getLastRow() > 1) {
    const allRecords = totalSheet.getDataRange().getValues().slice(1);
    allRecords.sort((a, b) => new Date(a[0]) - new Date(b[0]));

    allRecords.forEach(row => {
      const rGroup = String(row[3]).trim();
      if (!authorizedGroups.includes(rGroup)) return;
      const invId = String(row[1]).trim(), invName = String(row[2]).trim();
      const tarId = String(row[4]).trim(), tarName = String(row[5]).trim();
      const amt = Number(row[6]) || 0;
      const dateStr = Utilities.formatDate(new Date(row[8] || row[0]), "GMT+8", "MM/dd");

      if (!groupPerfMap[rGroup]) groupPerfMap[rGroup] = {};
      if (!groupPerfMap[rGroup][invId])
        groupPerfMap[rGroup][invId] = { name: invName, recruits: {}, count: 0, totalPoints: 0 };

      const invRef = groupPerfMap[rGroup][invId];
      if (!invRef.recruits[tarId]) {
        invRef.recruits[tarId] = { name: tarName, tarName, spent: 0, recDate: dateStr, effDate: "", points: 1, tarId, invId, invName, group: rGroup };
        invRef.count++;
      }
      const rec = invRef.recruits[tarId];
      const oldSpent = rec.spent;
      rec.spent += amt;
      if (oldSpent < effAmt && rec.spent >= effAmt) { rec.effDate = dateStr; rec.points = 2; }
    });
  }

  // 計算每組總分
  const groupPerf = Object.entries(groupPerfMap).map(([gName, invs]) => {
    const influencers = Object.entries(invs).map(([id, val]) => {
      val.totalPoints = Object.values(val.recruits).reduce((sum, r) => sum + r.points, 0);
      return { id, ...val };
    }).sort((a, b) => b.totalPoints - a.totalPoints);

    // 整組加總
    const groupTotalPoints = influencers.reduce((sum, inv) => sum + inv.totalPoints, 0);
    const groupTotalReferrals = influencers.reduce((sum, inv) => sum + inv.count, 0);
    const groupEffectiveCount = influencers.reduce((sum, inv) =>
      sum + Object.values(inv.recruits).filter(r => r.points === 2).length, 0);

    return { groupName: gName, influencers, groupTotalPoints, groupTotalReferrals, groupEffectiveCount };
  });

  Object.values(groupPerfMap).forEach(g => {
    Object.values(g).forEach(inv => {
      Object.values(inv.recruits).forEach(r => newbiesForCons.push({ ...r, isReached: r.spent >= effAmt }));
    });
  });
  newbiesForCons.sort((a, b) => (a.isReached === b.isReached) ? 0 : a.isReached ? 1 : -1);

  // 各雁群的報名成員名單（從名冊讀取，與推薦總表無關）
  const groupMembersMap = {};
  rosterData.slice(1).forEach(r => {
    const rGroup = String(r[3]).trim();
    const amId   = String(r[1]).trim();
    const name   = String(r[2]).trim();
    const status = String(r[4] || "").trim();
    if (!authorizedGroups.includes(rGroup) || !amId) return;
    if (status !== "是" && status !== "待確認") return;
    if (!groupMembersMap[rGroup]) groupMembersMap[rGroup] = [];
    groupMembersMap[rGroup].push({ name, amId, status });
  });

  // 把成員名單整理成陣列獨立回傳
  const groupMembers = authorizedGroups.map(gName => ({
    groupName: gName,
    members: groupMembersMap[gName] || []
  })).filter(g => g.members.length > 0);

  const displayProfileName = myAmId ? (maps.amIdToName[myAmId] || "使用者") : "管理員";

  // 競賽狀態沿用上面那次讀取的結果，不再重開試算表
  const contestStatus = { status: settings.status, label: settings.label };

  return {
    myProfile: { id: myAmId || "", group: myGroup || "", name: displayProfileName, classroom: myClassroom },
    participants, newbies: newbiesForCons, groupPerf, groupMembers, isPresident, groupList,
    contestStatus, version: APP_VERSION, effectiveAmount: effAmt
  };
}

// ==========================================
// 5. 輔助 API
// ==========================================

function apiLinkIdentity(p) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('雁群總名冊');
    const data = sheet.getDataRange().getValues();
    const alreadyBound = data.slice(1).some(r => String(r[0]).trim() === String(p.uid).trim());
    if (alreadyBound) return { success: false, message: "此 LINE 帳號已綁定其他編號。" };
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === String(p.amId).trim()) {
        const existingUid = String(data[i][0]).trim();
        if (existingUid && existingUid !== "") return { success: false, message: "此編號已被綁定，請聯絡管理員。" };
        sheet.getRange(i + 1, 1).setValue(p.uid);
        return { success: true, message: "身分認領成功！" };
      }
    }
    return { success: false, message: "找不到編號，請確認是否已報名。" };
  } catch(e) { return { success: false, message: "錯誤：" + e.toString() }; }
}

function apiSubmitReferral(p) {
  try {
    // 整支只開一次試算表（原本身分檢查、競賽狀態、寫入各開一次）
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // 身分把關：未完成認領者一律不得提交
    if (!isRegisteredUid(p && p.uid, ss))
      return { success: false, message: "請先完成身分認領後再提交。" };

    // 檢查競賽狀態
    const contestStatus = apiGetContestStatus(ss);
    if (contestStatus.status !== 'active') {
      const msg = contestStatus.status === 'pending' ? '競賽尚未開始，無法提交。' : '競賽已截止，無法提交。';
      return { success: false, message: msg };
    }

    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const imageList = Array.isArray(p.imageData) ? p.imageData : (p.imageData ? [p.imageData] : []);
    const imgUrls = [];
    imageList.forEach((imgData, i) => {
      if (!imgData || imgData.length <= 50) return;
      const parts = imgData.split(',');
      const mimeType = parts[0].split(':')[1].split(';')[0];
      const blob = Utilities.newBlob(Utilities.base64Decode(parts[1]), mimeType, `[待審]_${p.invName}推薦_${p.tarName}_${i + 1}`);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      imgUrls.push("https://lh3.googleusercontent.com/d/" + file.getId());
    });
    const imgUrlStr = imgUrls.join(',');
    const firstFileId = imgUrls.length > 0 ? imgUrls[0].replace("https://lh3.googleusercontent.com/d/", "") : "";
    const sheet = ss.getSheetByName('待審核清單');
    sheet.appendRow([new Date(), p.uid, p.invId, p.invName, p.tarId, p.tarName, p.date, Number(p.amt) || 0, "待審核", imgUrlStr, p.type, firstFileId]);
    return { success: true, message: "✨ 資料已提交。" };
  } catch(e) { return { success: false, message: "失敗：" + e.toString() }; }
}

function apiCheckTarId(tarId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tarIdStr = String(tarId).trim();
  const rosterSheet = ss.getSheetByName('雁群總名冊');
  if (rosterSheet && rosterSheet.getLastRow() > 1) {
    const roster = rosterSheet.getDataRange().getValues().slice(1);
    if (roster.some(r => String(r[1]).trim() === tarIdStr))
      return { exists: true, message: `⚠️ 編號 ${tarId} 已是報名者，請再次確認編號。` };
  }
  const totalSheet = ss.getSheetByName('推薦總表');
  if (totalSheet && totalSheet.getLastRow() > 1) {
    const records = totalSheet.getDataRange().getValues().slice(1);
    if (records.some(r => String(r[4]).trim() === tarIdStr))
      return { exists: true, message: `⚠️ 編號 ${tarId} 已是被推薦者，請再次確認編號。` };
  }
  return { exists: false };
}

// ==========================================
// 5a. 雁群清單（輕量版，供報名頁使用）
// ==========================================

function apiGetGroupList() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const groupSheet = ss.getSheetByName('雁群清單');
    if (!groupSheet || groupSheet.getLastRow() < 2) return [];
    return groupSheet.getRange(2, 1, groupSheet.getLastRow() - 1, 1)
      .getValues().flat()
      .map(g => String(g).trim())
      .filter(g => g !== "");
  } catch(e) { return []; }
}

// ==========================================
// 5b. 報名相關查詢
// ==========================================

function apiCheckRegisterId(amId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const amIdStr = String(amId || "").trim();
    if (!amIdStr) return { exists: false };
    const rosterSheet = ss.getSheetByName('雁群總名冊');
    if (rosterSheet && rosterSheet.getLastRow() > 1) {
      const roster = rosterSheet.getDataRange().getValues().slice(1);
      if (roster.some(r => String(r[1]).trim() === amIdStr))
        return { exists: true, message: '⚠️ 此編號已在名冊中，請確認是否重複報名。' };
    }
    return { exists: false };
  } catch(e) { return { exists: false }; }
}

function apiCheckMyRegistration(uid) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const uidStr = String(uid || "").trim();
    if (!uidStr) return { found: false };
    const rosterSheet = ss.getSheetByName('雁群總名冊');
    if (!rosterSheet || rosterSheet.getLastRow() < 2) return { found: false };
    const roster = rosterSheet.getDataRange().getValues().slice(1);
    const row = roster.find(r => String(r[0]).trim() === uidStr);
    if (!row) return { found: false };
    const status = String(row[4] || '').trim();
    return {
      found: true,
      amId: String(row[1]).trim(),
      name: String(row[2]).trim(),
      group: String(row[3]).trim(),
      status: status
    };
  } catch(e) { return { found: false }; }
}

// ==========================================
// 6. 競賽狀態管理
// ==========================================

// 一次把「系統設定」分頁讀完。純讀取、不寫入 —— 寫入放在讀取路徑上會拖慢每一次請求。
// 呼叫端已開啟試算表時請把 ss 傳進來，省下一次 openById。
function getSettings(ss) {
  const out = { status: 'active', label: '競賽進行中', effectiveAmount: 3000 };
  try {
    const book = ss || SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = book.getSheetByName('系統設定');
    if (!sheet) return out;
    sheet.getDataRange().getValues().forEach(row => {
      const key = String(row[0]).trim();
      if (key === 'contest_status') { const v = String(row[1]).trim(); if (v) out.status = v; }
      if (key === 'contest_label')  { const v = String(row[1]).trim(); if (v) out.label  = v; }
      if (key === 'effective_amount') {
        // 容許使用者填成 "5,000" 或 "$5000"
        const v = Number(String(row[1]).replace(/[^0-9.]/g, ''));
        if (v > 0) out.effectiveAmount = v;
      }
    });
  } catch(e) {}
  return out;
}

function apiGetContestStatus(ss) {
  const s = getSettings(ss);
  return { status: s.status, label: s.label };
}

function apiSetContestStatus(p) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    if (!getPresidentInfo(p && p.uid, ss).isPresident)
      return { success: false, message: "❌ 只有會長能變更競賽狀態。" };
    let settingSheet = ss.getSheetByName('系統設定');
    if (!settingSheet) {
      settingSheet = ss.insertSheet('系統設定');
      settingSheet.appendRow(['contest_status', p.status]);
      settingSheet.appendRow(['contest_label', p.label || '']);
      return { success: true, message: '設定已儲存。' };
    }
    const data = settingSheet.getDataRange().getValues();
    let statusRow = -1, labelRow = -1;
    data.forEach((row, i) => {
      if (String(row[0]).trim() === 'contest_status') statusRow = i + 1;
      if (String(row[0]).trim() === 'contest_label')  labelRow  = i + 1;
    });
    if (statusRow > 0) settingSheet.getRange(statusRow, 2).setValue(p.status);
    else settingSheet.appendRow(['contest_status', p.status]);
    if (labelRow > 0) settingSheet.getRange(labelRow, 2).setValue(p.label || '');
    else settingSheet.appendRow(['contest_label', p.label || '']);
    return { success: true, message: '競賽狀態已更新。' };
  } catch(e) {
    return { success: false, message: '儲存失敗：' + e.toString() };
  }
}

// ==========================================
// 7. 會長儀表板
// ==========================================

function apiGetAdminDashboard(uid) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const uidStr = String(uid || "").trim();
  const effAmt = getEffectiveAmount(ss);

  // 權限檢查：前端只是隱藏分頁，這裡才是真的把關
  const auth = getPresidentInfo(uidStr, ss);
  if (!auth.isPresident) return { error: "無權限" };

  const groupSheet   = ss.getSheetByName('雁群清單');
  const rosterSheet  = ss.getSheetByName('雁群總名冊');
  const pendingSheet = ss.getSheetByName('待審核清單');
  const totalSheet   = ss.getSheetByName('推薦總表');

  const rosterData    = rosterSheet ? rosterSheet.getDataRange().getValues() : [];
  const pendingValues = pendingSheet && pendingSheet.getLastRow() > 1 ? pendingSheet.getDataRange().getValues() : [];
  const totalData     = totalSheet && totalSheet.getLastRow() > 1 ? totalSheet.getDataRange().getValues() : [];

  // 取得審核者姓名
  const myClassroom = auth.classroom;
  let myName = "會長";
  if (auth.amId && rosterData.length > 1) {
    const found = rosterData.slice(1).find(r => String(r[1]).trim() === auth.amId);
    if (found) myName = String(found[2]).trim();
  }

  const groupToClassroom = {};
  if (groupSheet && groupSheet.getLastRow() >= 2) {
    groupSheet.getRange(2, 1, groupSheet.getLastRow() - 1, 2).getValues().forEach(r => {
      const gName = String(r[0]).trim();
      const cName = String(r[1] || "").trim();
      if (gName && cName) groupToClassroom[gName] = cName;
    });
  }

  const amIdToGroup = {};
  const allUsers = [];
  const groupedRoster = {};

  const pendingRegistrations = [];
  rosterData.slice(1).forEach(r => {
    const amId   = String(r[1]).trim();
    const name   = String(r[2]).trim();
    const group  = String(r[3]).trim();
    const status = String(r[4] || "").trim();
    const lineId = String(r[0] || "").trim();
    if (!amId) return;
    amIdToGroup[amId] = group;
    if (status === "是") {
      allUsers.push({ amId, name, group });
      if (!groupedRoster[group]) groupedRoster[group] = [];
      groupedRoster[group].push({ amId, name });
    } else if (status === "待確認") {
      // 過濾教室：只顯示屬於自己教室的待確認報名
      const gClassroom = groupToClassroom[group] || "";
      const canSee = !myClassroom || myClassroom === "全部" || gClassroom === myClassroom;
      if (canSee) pendingRegistrations.push({ amId, name, group, lineId, status });
    }
  });

  // 待審核 & 已審核紀錄
  const groupedPending = {};
  const reviewedItems = [];

  for (let i = 1; i < pendingValues.length; i++) {
    const row = pendingValues[i];
    const statusCell = String(row[8] || "").trim();
    const invId = String(row[2]).trim();
    const gName = amIdToGroup[invId] || "未知雁群";
    const gClassroom = groupToClassroom[gName] || "";
    const canSee = !myClassroom || myClassroom === "全部" || gClassroom === myClassroom;
    if (!canSee) continue;

    if (statusCell === "待審核") {
      if (!groupedPending[gName]) groupedPending[gName] = [];
      groupedPending[gName].push({ row: i + 1, invName: row[3], tarName: row[5], amt: row[7], img: row[9], type: row[10] });
    } else if (statusCell === "通過" || statusCell === "退回") {
      // 已審核紀錄，col[12] 存審核人姓名（新欄位）
      reviewedItems.push({
        invName:    row[3],
        tarName:    row[5],
        amt:        row[7],
        type:       row[10],
        status:     statusCell,
        reviewedBy: String(row[12] || "").trim() || "未記錄",
        reviewedAt: row[13] ? Utilities.formatDate(new Date(row[13]), "GMT+8", "MM/dd HH:mm") : (row[0] ? Utilities.formatDate(new Date(row[0]), "GMT+8", "MM/dd HH:mm") : ""),
        img:        row[9],
        groupName:  gName
      });
    }
  }

  const pending = Object.keys(groupedPending).map(k => ({ groupName: k, items: groupedPending[k] }));

  let totalReferrals = 0, totalEffective = 0;
  const yanqunMap = {};
  const recruitMap = {};

  totalData.slice(1).sort((a, b) => new Date(a[0]) - new Date(b[0])).forEach(row => {
    const invId  = String(row[1]).trim(), invName = String(row[2]).trim();
    const gName  = String(row[3]).trim();
    const tarId  = String(row[4]).trim(), tarName = String(row[5]).trim();
    const amt    = Number(row[6]) || 0;
    const key    = invId + '_' + tarId;
    if (!recruitMap[key]) {
      const dateStr = Utilities.formatDate(new Date(row[8] || row[0]), "GMT+8", "MM/dd");
      recruitMap[key] = { invId, invName, gName, tarId, tarName, spent: 0, points: 1, recDate: dateStr, effDate: "" };
      totalReferrals++;
    }
    const rec = recruitMap[key];
    const oldSpent = rec.spent;
    rec.spent += amt;
    if (oldSpent < effAmt && rec.spent >= effAmt) {
      rec.points = 2;
      rec.effDate = Utilities.formatDate(new Date(row[8] || row[0]), "GMT+8", "MM/dd");
      totalEffective++;
    }
  });

  Object.values(recruitMap).forEach(rec => {
    const gName = rec.gName;
    if (!yanqunMap[gName]) yanqunMap[gName] = {};
    if (!yanqunMap[gName][rec.invId]) yanqunMap[gName][rec.invId] = { name: rec.invName, recruits: [] };
    yanqunMap[gName][rec.invId].recruits.push({ name: rec.tarName, points: rec.points, recDate: rec.recDate, effDate: rec.effDate });
  });

  const yanqunStats = Object.entries(yanqunMap).map(([gName, invs]) => {
    let groupTotalPoints = 0, groupEffective = 0, groupReferrals = 0;
    const influencers = Object.entries(invs).map(([invId, inv]) => {
      const invTotal = inv.recruits.reduce((sum, r) => sum + r.points, 0);
      const invEff   = inv.recruits.filter(r => r.points === 2).length;
      groupTotalPoints += invTotal;
      groupEffective   += invEff;
      groupReferrals   += inv.recruits.length;
      return { name: inv.name, totalPoints: invTotal, recruits: inv.recruits };
    }).sort((a, b) => b.totalPoints - a.totalPoints);
    return { groupName: gName, totalPoints: groupTotalPoints, totalReferrals: groupReferrals, effectiveCount: groupEffective, influencers };
  }).sort((a, b) => b.totalPoints - a.totalPoints);

  return { pending, reviewedItems, allUsers, groupedRoster, pendingRegistrations, stats: { totalMembers: allUsers.length, totalReferrals, totalEffective }, yanqunStats, myName };
}

// ==========================================
// 8. 審核處理（記錄審核人）
// ==========================================

function apiProcessReview(p, status) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // 權限檢查
  const auth = getPresidentInfo(p && p.uid, ss);
  if (!auth.isPresident) return { success: false, message: "❌ 您沒有審核權限。" };

  const row = Number(p.row);
  if (!row || row < 2) return { success: false, message: "資料列號無效，請重新整理。" };

  // 上鎖：避免兩人同時審同一筆，或同一人連點兩下
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch(e) {
    return { success: false, message: "系統忙碌中，請稍後再試。" };
  }

  try {
    const sheet = ss.getSheetByName('待審核清單');
    if (row > sheet.getLastRow()) return { success: false, message: "找不到這筆資料，請重新整理。" };

    const d = sheet.getRange(row, 1, 1, 13).getValues()[0];

    // 冪等檢查：這筆若已被處理過就直接擋下，避免重複寫入推薦總表造成分數多算
    const currentStatus = String(d[8] || "").trim();
    if (currentStatus !== "待審核") {
      return { success: false, message: `⚠️ 這筆已經處理過了（目前狀態：${currentStatus || "未知"}），請重新整理。` };
    }

    const roster = ss.getSheetByName('雁群總名冊').getDataRange().getValues();

    // 審核人姓名以後端權限表對應的編號為準，不採信前端傳來的值
    let reviewerName = "";
    if (auth.amId) {
      for (let i = 1; i < roster.length; i++) {
        if (String(roster[i][1]).trim() === auth.amId) { reviewerName = String(roster[i][2]).trim(); break; }
      }
    }
    if (!reviewerName) reviewerName = String(p.reviewerName || "").trim() || "未記錄";

    if (status === "通過") {
      let g = "未知";
      for(let i = 1; i < roster.length; i++) {
        if(String(roster[i][1]).trim() === String(d[2]).trim()) { g = roster[i][3]; break; }
      }
      const refDate = d[6] ? new Date(d[6]) : new Date();
      ss.getSheetByName('推薦總表').appendRow([new Date(), d[2], d[3], g, d[4], d[5], d[7], d[10], refDate]);
    }

    // 更新狀態（第9欄）
    sheet.getRange(row, 9).setValue(status);
    // 記錄審核人姓名（第13欄）
    sheet.getRange(row, 13).setValue(reviewerName);
    // 記錄審核時間（第14欄）
    sheet.getRange(row, 14).setValue(new Date());
    // 確保狀態確實寫回試算表後才放鎖，否則下一個請求可能讀到舊值
    SpreadsheetApp.flush();

    // 更新 Drive 檔名
    const fileId = d[11];
    if (fileId) { try { DriveApp.getFileById(fileId).setName(`[${status}]_${d[3]}推薦_${d[5]}`); } catch(e) {} }

    return { success: true, message: "審核完成: " + status };
  } catch(e) {
    return { success: false, message: "審核失敗：" + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 9. 頒獎管理
// ==========================================

function apiGetGiftData(uid) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const uidStr = String(uid || "").trim();
  const effAmt = getEffectiveAmount(ss);

  const permSheet = ss.getSheetByName('LINE權限表');
  const permData  = permSheet.getDataRange().getValues();
  let isPresident = false;
  let myClassroom = "";
  for (let i = 1; i < permData.length; i++) {
    if (String(permData[i][0]).trim() === uidStr) {
      if (String(permData[i][3]).trim() === "會長") isPresident = true;
      myClassroom = String(permData[i][4] || "").trim();
      break;
    }
  }
  if (!isPresident) return { error: "無權限" };

  const giftSheet   = ss.getSheetByName('禮物清單');
  const recordSheet = ss.getSheetByName('領獎紀錄');
  const totalSheet  = ss.getSheetByName('推薦總表');

  const giftData = giftSheet && giftSheet.getLastRow() > 1
    ? giftSheet.getRange(2, 1, giftSheet.getLastRow() - 1, 4).getValues() : [];

  const recordData = recordSheet && recordSheet.getLastRow() > 1
    ? recordSheet.getDataRange().getValues().slice(1) : [];

  const giftUsedCount = {};
  recordData.forEach(r => {
    const giftName  = String(r[4]).trim();
    const classroom = String(r[9] || "").trim();
    if (classroom === myClassroom && giftName) {
      giftUsedCount[giftName] = (giftUsedCount[giftName] || 0) + 1;
    }
  });

  const gifts = giftData
    .filter(r => String(r[2] || "").trim() === myClassroom)
    .map(r => {
      const name      = String(r[0]).trim();
      const total     = Number(r[1]) || 0;
      const used      = giftUsedCount[name] || 0;
      const remaining = Math.max(0, total - used);
      return { name, total, used, remaining, note: String(r[3] || "").trim() };
    }).filter(g => g.name !== "");

  const giftRecordMap = {};
  recordData.forEach((r, i) => {
    const classroom = String(r[9] || "").trim();
    if (classroom !== myClassroom) return;
    const invId = String(r[1]).trim();
    const tarId = String(r[2]).trim();
    const key   = invId + '_' + tarId;
    giftRecordMap[key] = {
      row:          i + 2,
      invId,
      tarId,
      giftName:     String(r[4]).trim(),
      status:       String(r[5]).trim(),
      claimDate:    r[6] ? Utilities.formatDate(new Date(r[6]), "GMT+8", "yyyy-MM-dd") : "",
      registerDate: r[7] ? Utilities.formatDate(new Date(r[7]), "GMT+8", "yyyy-MM-dd") : "",
      note:         String(r[8] || "").trim(),
    };
  });

  const totalData = totalSheet && totalSheet.getLastRow() > 1
    ? totalSheet.getDataRange().getValues().slice(1) : [];

  const groupMap = {};
  totalData.sort((a, b) => new Date(a[0]) - new Date(b[0])).forEach(row => {
    const gName   = String(row[3]).trim();
    const invId   = String(row[1]).trim();
    const invName = String(row[2]).trim();
    const tarId   = String(row[4]).trim();
    const tarName = String(row[5]).trim();
    const amt     = Number(row[6]) || 0;

    if (!groupMap[gName]) groupMap[gName] = {};
    if (!groupMap[gName][invId]) groupMap[gName][invId] = { invName, recruits: {} };

    if (!groupMap[gName][invId].recruits[tarId]) {
      groupMap[gName][invId].recruits[tarId] = { tarId, tarName, spent: 0, points: 1 };
    }
    const rec = groupMap[gName][invId].recruits[tarId];
    const oldSpent = rec.spent;
    rec.spent += amt;
    if (oldSpent < effAmt && rec.spent >= effAmt) rec.points = 2;
  });

  const groupedReferrals = Object.entries(groupMap).map(([gName, invs]) => {
    const influencers = Object.entries(invs).map(([invId, inv]) => {
      const recruits = Object.values(inv.recruits).map(rec => {
        const rKey = invId + '_' + rec.tarId;
        const awardRecord = giftRecordMap[rKey] || null;
        return { ...rec, awardRecord };
      });
      return { invId, invName: inv.invName, recruits };
    });
    return { groupName: gName, influencers };
  });

  return { gifts, groupedReferrals, myClassroom };
}

function apiAddGift(p) {
  try {
    if (!p.name) return { success: false, message: "請填寫禮物名稱" };
    if (!p.classroom) return { success: false, message: "無法取得教室資訊" };
    const ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
    if (!getPresidentInfo(p && p.uid, ss).isPresident)
      return { success: false, message: "❌ 只有會長能管理禮物。" };
    const giftSheet = ss.getSheetByName('禮物清單');
    const existing  = giftSheet.getLastRow() > 1 ? giftSheet.getDataRange().getValues().slice(1) : [];
    if (existing.some(r => String(r[0]).trim() === p.name.trim() && String(r[2]).trim() === p.classroom.trim()))
      return { success: false, message: "此教室已有相同禮物名稱。" };
    giftSheet.appendRow([p.name.trim(), Number(p.total) || 0, p.classroom.trim(), p.note || ""]);
    return { success: true, message: `✅ 已新增「${p.name}」。` };
  } catch(e) { return { success: false, message: "新增失敗：" + e.toString() }; }
}

function apiDeleteGift(p) {
  try {
    const ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
    if (!getPresidentInfo(p && p.uid, ss).isPresident)
      return { success: false, message: "❌ 只有會長能管理禮物。" };
    const giftSheet = ss.getSheetByName('禮物清單');
    const data      = giftSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(p.name).trim() &&
          String(data[i][2]).trim() === String(p.classroom).trim()) {
        giftSheet.deleteRow(i + 1);
        return { success: true, message: `已刪除「${p.name}」。` };
      }
    }
    return { success: false, message: "找不到此禮物。" };
  } catch(e) { return { success: false, message: "刪除失敗：" + e.toString() }; }
}

function apiAssignGift(p) {
  try {
    const ss          = SpreadsheetApp.openById(SPREADSHEET_ID);
    if (!getPresidentInfo(p && p.uid, ss).isPresident)
      return { success: false, message: "❌ 只有會長能登記領獎。" };
    const recordSheet = ss.getSheetByName('領獎紀錄');
    const giftName    = p.giftName || "";

    if (giftName && p.status === "已領取") {
      const giftSheet = ss.getSheetByName('禮物清單');
      const giftData  = giftSheet.getDataRange().getValues().slice(1);
      let giftTotal   = 0;
      for (const r of giftData) {
        if (String(r[0]).trim() === giftName && String(r[2]).trim() === (p.classroom || "")) {
          giftTotal = Number(r[1]) || 0;
          break;
        }
      }
      const recordData = recordSheet.getLastRow() > 1 ? recordSheet.getDataRange().getValues().slice(1) : [];
      const usedCount  = recordData.filter(r =>
        String(r[4]).trim() === giftName && String(r[9] || "").trim() === (p.classroom || "")
      ).length;
      if (giftTotal > 0 && usedCount >= giftTotal)
        return { success: false, message: `「${giftName}」已額滿，無法再登記。` };
    }

    const recordData = recordSheet.getLastRow() > 1 ? recordSheet.getDataRange().getValues().slice(1) : [];
    const alreadyAssigned = recordData.some(
      r => String(r[1]).trim() === String(p.invId).trim() &&
           String(r[2]).trim() === String(p.tarId).trim() &&
           String(r[9] || "").trim() === (p.classroom || "")
    );
    if (alreadyAssigned) return { success: false, message: `此推薦紀錄已登記過禮物。` };

    const today     = new Date();
    const claimDate = (p.status === "已領取" && p.claimDate) ? new Date(p.claimDate) : "";
    recordSheet.appendRow([today, p.invId, p.tarId, p.tarName, giftName, p.status, claimDate, today, p.note || "", p.classroom || ""]);
    const msg = p.status === "待領取"
      ? `⏳ 已將「${p.tarName}」標記為待領取。`
      : `✅ 已為「${p.tarName}」登記「${giftName}」。`;
    return { success: true, message: msg };
  } catch(e) { return { success: false, message: "寫入失敗：" + e.toString() }; }
}

function apiConfirmGift(p) {
  try {
    const ss          = SpreadsheetApp.openById(SPREADSHEET_ID);
    if (!getPresidentInfo(p && p.uid, ss).isPresident)
      return { success: false, message: "❌ 只有會長能確認領取。" };
    const recordSheet = ss.getSheetByName('領獎紀錄');

    if (p.giftName) {
      const giftSheet = ss.getSheetByName('禮物清單');
      const giftData  = giftSheet.getDataRange().getValues().slice(1);
      let giftTotal   = 0;
      for (const r of giftData) {
        if (String(r[0]).trim() === p.giftName && String(r[2]).trim() === (p.classroom || "")) {
          giftTotal = Number(r[1]) || 0;
          break;
        }
      }
      if (giftTotal > 0) {
        const recordData = recordSheet.getLastRow() > 1 ? recordSheet.getDataRange().getValues().slice(1) : [];
        const usedCount  = recordData.filter(r =>
          String(r[4]).trim() === p.giftName &&
          String(r[5]).trim() === "已領取" &&
          String(r[9] || "").trim() === (p.classroom || "")
        ).length;
        if (usedCount >= giftTotal)
          return { success: false, message: `「${p.giftName}」已額滿，無法再登記。` };
      }
    }

    const claimDate = p.claimDate ? new Date(p.claimDate) : new Date();
    recordSheet.getRange(p.row, 5).setValue(p.giftName || "");
    recordSheet.getRange(p.row, 6).setValue("已領取");
    recordSheet.getRange(p.row, 7).setValue(claimDate);
    return { success: true, message: "✅ 已確認領取。" };
  } catch(e) { return { success: false, message: "更新失敗：" + e.toString() }; }
}


// ==========================================
// 11. 報名審核
// ==========================================

function apiSubmitRegister(p) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('雁群總名冊');
    const rosterData = sheet.getDataRange().getValues();
    if (isAmIdRegistered(p.amId, rosterData)) return { success: false, message: "此安麗編號已存在名冊中。" };
    // 代報名：uid 為空；自己報名：uid 為自己的 LINE UID
    const uid = p.uid || "";
    sheet.appendRow([uid, p.amId, p.name, p.group, "待確認"]);
    return { success: true, message: "✅ 報名申請已送出，等待會長確認。" };
  } catch(e) { return { success: false, message: "寫入失敗：" + e.toString() }; }
}

function apiApproveRegister(p) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    if (!getPresidentInfo(p && p.uid, ss).isPresident)
      return { success: false, message: "❌ 您沒有審核報名的權限。" };
    const sheet = ss.getSheetByName('雁群總名冊');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === String(p.amId).trim()) {
        sheet.getRange(i + 1, 5).setValue("是");
        return { success: true, message: `✅ 已確認「${p.name}」的報名。` };
      }
    }
    return { success: false, message: "找不到此報名資料。" };
  } catch(e) { return { success: false, message: "操作失敗：" + e.toString() }; }
}

function apiRejectRegister(p) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    if (!getPresidentInfo(p && p.uid, ss).isPresident)
      return { success: false, message: "❌ 您沒有審核報名的權限。" };
    const sheet = ss.getSheetByName('雁群總名冊');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === String(p.amId).trim()) {
        sheet.getRange(i + 1, 5).setValue("退回");
        return { success: true, message: `已退回「${p.name}」的報名。` };
      }
    }
    return { success: false, message: "找不到此報名資料。" };
  } catch(e) { return { success: false, message: "操作失敗：" + e.toString() }; }
}

// ==========================================
// 10. 共用工具函式
// ==========================================

// 依 LINE UID 查權限表，回傳身分資訊。所有「寫入型」API 都要先過這關。
// ss 可選：呼叫端已開啟試算表時傳進來，避免重複開檔。
function getPresidentInfo(uid, ss) {
  const result = { found: false, isPresident: false, classroom: "", amId: "" };
  const uidStr = String(uid || "").trim();
  if (!uidStr) return result;
  try {
    const book = ss || SpreadsheetApp.openById(SPREADSHEET_ID);
    const permSheet = book.getSheetByName('LINE權限表');
    if (!permSheet) return result;
    const permData = permSheet.getDataRange().getValues();
    for (let i = 1; i < permData.length; i++) {
      if (String(permData[i][0]).trim() === uidStr) {
        result.found       = true;
        result.isPresident = String(permData[i][3]).trim() === "會長";
        result.classroom   = String(permData[i][4] || "").trim();
        result.amId        = String(permData[i][1] || "").trim();
        break;
      }
    }
  } catch(e) {}
  return result;
}

// 有效推薦的消費門檻。在「系統設定」分頁用 effective_amount 這個鍵調整，
// 例如 A 欄填 effective_amount、B 欄填 5000。讀不到或填了無效值時退回 3000。
function getEffectiveAmount(ss) {
  return getSettings(ss).effectiveAmount;
}

// 這個 LINE UID 是否已完成身分認領（綁定名冊），或列在權限表中。
// 未通過的人不得提交任何資料 —— 前端關卡是體驗，這裡才是真正的門。
function isRegisteredUid(uid, ss) {
  const uidStr = String(uid || "").trim();
  if (!uidStr) return false;
  try {
    const book = ss || SpreadsheetApp.openById(SPREADSHEET_ID);
    const rosterSheet = book.getSheetByName('雁群總名冊');
    if (rosterSheet && rosterSheet.getLastRow() > 1) {
      const data = rosterSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === uidStr) return true;
      }
    }
    // 管理者可能只在權限表、不在名冊
    return getPresidentInfo(uidStr, book).found;
  } catch(e) { return false; }
}

function getMapsFromData(rosterData) {
  const uidToAmId = {}, amIdToGroup = {}, amIdToName = {};
  rosterData.slice(1).forEach(r => {
    const uid = String(r[0]).trim(), amId = String(r[1]).trim();
    if (uid) uidToAmId[uid] = amId;
    if (amId) { amIdToGroup[amId] = String(r[3]).trim(); amIdToName[amId] = String(r[2]).trim(); }
  });
  return { uidToAmId, amIdToGroup, amIdToName };
}

function isAmIdRegistered(amId, rosterData) {
  return rosterData.some(r => String(r[1]).trim() === String(amId).trim());
}

function isUserRegistered(uid, rosterData) {
  return rosterData.some(r => String(r[0]).trim() === String(uid).trim());
}

function responseJson(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function replyLine(token, text) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN, 'Content-Type': 'application/json' },
    payload: JSON.stringify({ replyToken: token, messages: [{ type: 'text', text: text }] })
  });
}

function pushLineFlex(token, flex) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN, 'Content-Type': 'application/json' },
    payload: JSON.stringify({ replyToken: token, messages: [{ type: "flex", altText: "選擇雁群", contents: flex }] })
  });
}

function sendGroupFlex(token, amId, name) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const groups = ss.getSheetByName('雁群清單').getRange("A2:A" + ss.getSheetByName('雁群清單').getLastRow()).getValues().flat().filter(g => g !== "");
  const buttons = groups.map(g => ({
    type: "button", height: "sm",
    action: { type: "message", label: g, text: `【確認報名】編號:${amId}|姓名:${name}|雁群:${g}` },
    style: "secondary", margin: "sm"
  }));
  const flex = {
    type: "bubble",
    header: { type: "box", layout: "vertical", contents: [{ type: "text", text: "最後一步：選擇您的雁群", weight: "bold", size: "lg" }] },
    body: { type: "box", layout: "vertical", spacing: "sm", contents: [{ type: "text", text: `您好 ${name}，請選所屬雁群：`, size: "sm", color: "#666666", wrap: true }, ...buttons] }
  };
  pushLineFlex(token, flex);
}
