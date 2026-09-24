/* SCU Small Group — Admin Console web app.
   This project serves only the admin console; members use the separate
   SCU Small Group app, which reads the same spreadsheet. */

const SHEET_ID = '1XhVmwTimVD1VaPkNVCpriNn_g7wcJXZjPGC7wJ-1vaI';
var SS_CACHE_ = null;
function ss_() {
  return SS_CACHE_ || (SS_CACHE_ = SpreadsheetApp.openById(SHEET_ID));
}

function doGet(e) {
  try {
    return adminDoGet_();
  } catch (error) {
    return HtmlService.createHtmlOutput(`<div style="padding:40px;text-align:center;"><h2>System Notice</h2><p>${error.message}</p></div>`);
  }
}

/* Inlines another HTML file of this project (Admin.html pulls in AdminApp). */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* Each tab is read at most once per request: every reader goes through here.
   Writers call clearSheetMemo_() so later reads in the same request see them. */
var SHEET_MEMO_ = {};
function sheetValues_(sheetName) {
  if (!(sheetName in SHEET_MEMO_)) {
    const sheet = ss_().getSheetByName(sheetName);
    SHEET_MEMO_[sheetName] = sheet ? sheet.getDataRange().getValues() : null;
  }
  return SHEET_MEMO_[sheetName];
}
function clearSheetMemo_() { SHEET_MEMO_ = {}; }

function readSheetAsMap(sheetName) {
  const data = sheetValues_(sheetName);
  if (!data || data.length < 2) return [];
  const headers = data[0].map(h => String(h).trim());
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { if (h) obj[h] = row[i]; });
    return obj;
  });
}

function getInitials(name) {
  if (!name) return "US";
  const parts = name.trim().split(/\s+/);
  if (parts.length === 1) return (parts[0][0] + parts[0][0]).toUpperCase();
  return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
}
