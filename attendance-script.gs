// ============================================================
//  Attendance Scanner — Google Apps Script
//  Columns: Timestamp | Name | Email | Location | Handphone Number
//  Sheet: https://docs.google.com/spreadsheets/d/113KIEqWC3NyGhiVz99Dg4lREqvcrVBTYuRktw5y6KGo
// ============================================================

var SPREADSHEET_ID = "113KIEqWC3NyGhiVz99Dg4lREqvcrVBTYuRktw5y6KGo";
var SHEET_GID      = 246908960;
var SHEET_NAME     = "Attendance";

function doGet(e) {
  var output;
  try {
    var p = e.parameter || {};

    if (!p.name) {
      output = { status: "ok", message: "Script is running!" };
    } else {
      var sheet = getTargetSheet();

      // Header row if empty — match exact Google Sheet columns
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(["Timestamp", "Name", "Email", "Location", "Handphone Number"]);
        sheet.getRange(1, 1, 1, 5)
          .setFontWeight("bold")
          .setBackground("#4a1a8c")
          .setFontColor("#ffffff");
        sheet.setFrozenRows(1);
      }

      // Write row matching exact column order
      sheet.appendRow([
        new Date().toLocaleString("en-SG"),  // Timestamp
        p.name     || "",                     // Name
        p.email    || "",                     // Email
        p.location || "",                     // Location
        p.phone    || ""                      // Handphone Number
      ]);

      sheet.autoResizeColumns(1, 5);
      output = { status: "ok", message: "Logged: " + p.name };
    }
  } catch(err) {
    output = { status: "error", message: err.toString() };
  }

  // JSONP support
  var cb   = (e.parameter || {}).callback;
  var json = JSON.stringify(output);
  return ContentService
    .createTextOutput(cb ? cb + "(" + json + ")" : json)
    .setMimeType(cb ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}

function getTargetSheet() {
  var ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === SHEET_GID) return sheets[i];
  }
  var byName = ss.getSheetByName(SHEET_NAME);
  if (byName) return byName;
  return ss.insertSheet(SHEET_NAME);
}
