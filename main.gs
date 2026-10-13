const MASTER_SHEET_NAME = '總表';
const USER_MAP_SHEET = '使用者對照表';
const HANDLER_HEADERS = ['負責人', '購案名稱', '承攬廠商',	'聯絡人', '開始日期', '結束日期', '計畫金額','年度分配','類型','經費分擔計畫',	'分攤經費',	'採購案號','備註', '狀態', '最後更新時間'];
const MASTER_HEADERS = ['負責人', '購案名稱', '承攬廠商',	'聯絡人', '開案日期', '結案日期', '計畫金額','年度分配','類型','經費分擔計畫',	'分攤經費',	'採購案號','備註', '狀態'];

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('計畫管理')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------- 使用者身分 ---------- */

function getCurrentEmail_() {
  const email = Session.getActiveUser().getEmail();
  if (!email) throw new Error('無法取得使用者 Email,請確認部署設定的「執行身分」與「存取權限」');
  return email;
}

function getUserMapSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(USER_MAP_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(USER_MAP_SHEET);
    sheet.getRange(1, 1, 1, 2).setValues([['Email', '承辦人']]);
  }
  return sheet;
}

function getCurrentHandlerName_() {
  const email = getCurrentEmail_();
  const sheet = getUserMapSheet_();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) return data[i][1];
  }
  return null;
}

function getCurrentUser() {
  return {
    email: getCurrentEmail_(),
    handlerName: getCurrentHandlerName_()
  };
}

function setHandlerName(name) {
  name = String(name).trim();
  const email = getCurrentEmail_();
  const sheet = getUserMapSheet_();
  const data = sheet.getDataRange().getValues();
  let updated = false;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      sheet.getRange(i + 1, 2).setValue(name);
      updated = true;
      break;
    }
  }
  if (!updated) sheet.appendRow([email, name]);
  getOrCreateHandlerSheet_(name); // 確保個人分頁存在
  return { email: email, handlerName: name };
}

/* ---------- 分頁管理 ---------- */

function sanitizeSheetName_(name) {
  let s = String(name).replace(/[\[\]\*\?\/\\:]/g, '_').trim();
  if (s.length > 90) s = s.substring(0, 90);
  if (!s) s = '未命名';
  return s;
}

function getOrCreateHandlerSheet_(handlerName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = sanitizeSheetName_(handlerName);
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, HANDLER_HEADERS.length).setValues([HANDLER_HEADERS]);
  }
  return sheet;
}

function getMasterSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(MASTER_SHEET_NAME);
    sheet.getRange(1, 1, 1, MASTER_HEADERS.length).setValues([MASTER_HEADERS]);
  }
  return sheet;
}

/* ---------- 日期格式 ---------- */

function formatDate_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return v;
}

function formatDateTime_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  }
  return v;
}

function searchMasteryr(yr){
  const sheet = getMasterSheet_();
  const data = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[1]) continue; // 略過空列
      const yrAmount = row[7];
      let matchAmount = '';
      let totalAmount = 0;
      let monthDetail = [];
      for (let m = 0 ; m < 12 ; m++){
          monthDetail.push('');
      }


      paygroup = yrAmount.split('\n');
      for (let k = 0 ; k < paygroup.length ; k++){
          payAmount = parseFloat(paygroup[k].split('__')[0]);
          payDate = paygroup[k].split('__')[2];
          if (payDate == ''){
            continue;
          }
          if (payDate.split('-')[0] == yr){
              matchAmount += paygroup[k]+'\n';
              totalAmount += payAmount;
              month = parseInt(payDate.split('-')[1],10)
              monthDetail[ month-1 ] += String( payAmount )+'\n';
          }

      }
      if (totalAmount != 0){
        result.push({
          handler: row[0],
          planName: row[1],
          contractor: row[2],
          contectName: row[3],
          startDate: formatDate_(row[4]),
          endDate: formatDate_(row[5]),
          amount: totalAmount,
          yrAmount: matchAmount,
          planType: row[8],
          otherPlan: row[9],
          otherAmount: row[10],
          planNo:row[11],
          remark:row[12],
          status:row[13],
          monthDetail : monthDetail
        });
      }

  }
  return result;
}


/* ---------- 讀取自己名下的資料(完整歷史) ---------- */

function getMyRecords() {
  const handlerName = getCurrentHandlerName_();
  if (!handlerName) return [];
  const sheet = getOrCreateHandlerSheet_(handlerName);
  const data = sheet.getDataRange().getValues();
  const result = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[1]) continue; // 略過空列
    result.push({
      handler: row[0],
      planName: row[1],
      contractor: row[2],
      contectName: row[3],
      startDate: formatDate_(row[4]),
      endDate: formatDate_(row[5]),
      amount: row[6],
      yrAmount: row[7],
      planType: row[8],
      otherPlan: row[9],
      otherAmount: row[10],
      planNo:row[11],
      remark:row[12],
      status:row[13],
      timestamp:formatDateTime_(row[14]),
    });
  }
  
  result.reverse(); // 新到舊
  return result;
}
/*------------刪除舊的-----------------------------*/
function removeRecord(rowIndex) {
  const handlerName = getCurrentHandlerName_();   
  const sheet = getOrCreateHandlerSheet_(handlerName);
  const rowData = sheet.getRange(rowIndex, 1, 1, HANDLER_HEADERS.length).getValues()[0];
  if (rowData[2] !== handlerName) {
    throw new Error('無權刪除此資料');
  }
  sheet.deleteRow(rowIndex);
  return { success: true };

}
/* ---------- 新增一筆(遞增, 不修改不刪除) ---------- */

function addRecord(record) {
  const handlerName = getCurrentHandlerName_();
  if (!handlerName) throw new Error('尚未設定承辦人身分');
  const planName = String(record.planName || '').trim();
  if (!planName) throw new Error('請填寫計劃名');

  const timestamp = new Date();
  handlerrow =  [
      handlerName,
      planName,
      record.contractor,
      record.contectName,
      record.startDate,
      record.endDate,
      record.amount,
      record.yrAmount,
      record.planType,
      record.otherPlan,
      record.otherAmount,
      record.planNo,
      record.remark,
      record.status,
      timestamp,
     
  ];
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);


  try {
    const sheet = getOrCreateHandlerSheet_(handlerName);
    let update = 0;
    
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][1] === planName) {
        sheet.getRange(i + 1, 1, 1, 15).setValues([handlerrow])
        update = 1;
      }
    }
    if (update == 0){
      sheet.appendRow(handlerrow);
    }


    //recalcLatestStatus_(sheet, planName);

    const latest = getplanFromHandlerSheet_(sheet, planName);
    upsertMaster_(latest);
  } finally {
    lock.releaseLock();
  }
  return { success: true };
}

/** 重新計算某個計劃名在此分頁中,哪一列是時間戳記最新的一筆,標記狀態欄 */
function recalcLatestStatus_(sheet, planName) {
  const data = sheet.getDataRange().getValues();
  let latestRow = -1;
  let latestTime = null;
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === planName) {
      const t = data[i][15];
      if (latestTime === null || (t instanceof Date && t > latestTime)) {
        latestTime = t;
        latestRow = i;
      }
    }
  }
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === planName) {
      sheet.getRange(i + 1, 15).setValue(i === latestRow ? '現況' : '');
    }
  }
}

/** 取出某分頁中,該計劃名相同那一筆 */
function getplanFromHandlerSheet_(sheet, planName) {
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][1] === planName ) {
      return {
        handler: data[i][0],
        planName: data[i][1],
        contractor: data[i][2],
        contectName: data[i][3],
        startDate: formatDate_(data[i][4]),
        endDate: formatDate_(data[i][5]),
        amount: data[i][6],
        yrAmount: data[i][7],
        planType: data[i][8],
        otherPlan: data[i][9],
        otherAmount: data[i][10],
        planNo:data[i][11],
        remark:data[i][12],
        status:data[i][13]
      };
    }
  }
  return null;
}

/** 把最新資料同步(覆蓋或新增)到總表,以「計劃名」作為比對鍵 */
function upsertMaster_(record) {
  if (!record) return;
  const sheet = getMasterSheet_();
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === record.planName) {
      sheet.getRange(i + 1, 1, 1, 14).setValues([[
              record.handler,
              record.planName,
              record.contractor,
              record.contectName,
              record.startDate,
              record.endDate,
              record.amount,
              record.yrAmount,
              record.planType,
              record.otherPlan,
              record.otherAmount,
              record.planNo,
              record.remark,
              record.status,
      ]]);
      return;
    }
  }
  sheet.appendRow([
              record.handler,
              record.planName,
              record.contractor,
              record.contectName,
              record.startDate,
              record.endDate,
              record.amount,
              record.yrAmount,
              record.planType,
              record.otherPlan,
              record.otherAmount,
              record.planNo,
              record.remark,
              record.status,]);
}