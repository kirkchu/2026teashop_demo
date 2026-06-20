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
  const ss = SpreadsheetApp.openById('1zb7X_fYhPoYP_-NckPOKzN6J221Mkmrr82iBzahzuaM');
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    try {
      sheet = ss.insertSheet(SHEET_NAME);
    } catch(e) {
      // sheet might exist with slight name difference - find it by iterating
      sheet = ss.getSheets().find(s => s.getName().trim() === SHEET_NAME.trim()) || ss.getSheets()[0];
    }
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
