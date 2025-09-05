/**
 * Apps Script 後端 (Code.gs)
 * 提供前端存取 Google Sheets 的簡單 Key-Value 存儲 (模擬 localStorage)
 * 使用方式：
 *  - 把前端 index.html 放到同一個 Apps Script 專案
 *  - 部署為「網頁應用程式」，預設 anyone with link 可以存取（或依需求調整）
 */

/** 頁面入口 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index').setTitle('班級管理系統');
}

/** 確保 storage 工作表存在，並回傳該工作表 */
function _ensureStorageSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('storage');
  if (!sheet) {
    sheet = ss.insertSheet('storage');
    sheet.getRange(1,1,1,2).setValues([['key','value']]);
  }
  return sheet;
}

/** 取得 key 的內容（如果找不到，回傳 null） */
function getStorage(key) {
  const sheet = _ensureStorageSheet();
  const values = sheet.getDataRange().getValues();
  // skip header row
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(key)) {
      const raw = values[i][1];
      if (raw === '' || raw === null || raw === undefined) return null;
      try {
        return JSON.parse(raw);
      } catch (e) {
        // 如果無法 parse，直接回傳原始字串
        return raw;
      }
    }
  }
  return null;
}

/** 設定 key 的內容（會覆蓋或新增） */
function setStorage(key, value) {
  const sheet = _ensureStorageSheet();
  const values = sheet.getDataRange().getValues();
  const json = JSON.stringify(value === undefined ? null : value);
  // 找到 row
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(key)) {
      sheet.getRange(i+1, 2).setValue(json);
      return { status: 'updated', key: key };
    }
  }
  // not found -> append
  sheet.appendRow([key, json]);
  return { status: 'created', key: key };
}

/** 刪除 key（如需） */
function removeStorage(key) {
  const sheet = _ensureStorageSheet();
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(key)) {
      sheet.deleteRow(i+1);
      return { status: 'deleted', key: key };
    }
  }
  return { status: 'not_found', key: key };
}

/** 方便的 debug：回傳所有 key */
function listStorageKeys() {
  const sheet = _ensureStorageSheet();
  const values = sheet.getDataRange().getValues();
  const keys = [];
  for (let i = 1; i < values.length; i++) {
    keys.push(String(values[i][0]));
  }
  return keys;
}
