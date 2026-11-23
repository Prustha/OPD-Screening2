const APP_USERS = Object.freeze({
  admin: 'admin'
});

/** Serves the application as a Google Apps Script web app. */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Smart_OPD_Triage_Demo')
    .setTitle('Smart OPD Triage');
}

/** Validates the login form without exposing credentials to the browser. */
function validateLogin(username, password) {
  const account = String(username || '').trim();
  return Object.prototype.hasOwnProperty.call(APP_USERS, account) &&
    APP_USERS[account] === String(password || '');
}

/** Handles login requests sent from the standalone HTML page. */
function doPost(e) {
  let request;
  try {
    request = JSON.parse((e.postData && e.postData.contents) || '{}');
  } catch (error) {
    return jsonResponse({ ok: false, error: 'Invalid request' });
  }

  if (request.action === 'login') {
    return jsonResponse({
      ok: validateLogin(request.username, request.password)
    });
  }

  return jsonResponse({ ok: false, error: 'Unsupported action' });
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Creates the OPD Triage worksheet and its column headers.
 * Run once from the Apps Script editor before receiving records.
 */
function SetupSheet() {
  const sheetName = 'OPD Triage';
  const headers = [
    'วันที่และเวลาบันทึก',
    'รหัสผู้ป่วยจำลอง',
    'อายุ (ปี)',
    'เพศ',
    'เริ่มมีอาการ',
    'อาการสำคัญ',
    'SBP (mmHg)',
    'DBP (mmHg)',
    'HR (ครั้ง/นาที)',
    'RR (ครั้ง/นาที)',
    'BT (°C)',
    'SpO₂ (%)',
    'Pain score (0–10)',
    'ระดับ Triage',
    'สี Triage',
    'สถานะ',
    'บันทึกเพิ่มเติม'
  ];

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#172b49')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    sheet.autoResizeColumns(1, headers.length);
  }

  return sheet;
}
