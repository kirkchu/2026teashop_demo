// 春水堂班級點餐 - Google Apps Script
// 部署步驟：
//   1. 開啟 Google Sheet（名稱：2026春水堂）
//   2. 擴充功能 → Apps Script → 貼上此程式碼
//   3. 部署 → 新增部署 → 類型選「網頁應用程式」
//   4. 執行身分：我、存取對象：所有人
//   5. 複製部署 URL → 貼到 index.html 與 orders.html 的 SCRIPT_URL

const SHEET_NAME = '2026春水堂';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getOrCreateSheet();

    sheet.appendRow([
      new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' }),
      data.name,
      data.items,
      data.total
    ]);

    return jsonResponse({ success: true });
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

function doGet(e) {
  try {
    const sheet = getOrCreateSheet();
    const lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      return jsonResponse({ success: true, data: [] });
    }

    const values = sheet.getDataRange().getValues();
    return jsonResponse({ success: true, data: values });
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    const headers = ['時間戳記', '姓名', '餐點明細', '總金額(元)'];
    sheet.appendRow(headers);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#8B2635');
    headerRange.setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
