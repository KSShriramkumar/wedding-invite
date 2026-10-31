// ============================================================
// Wedding RSVP -> Google Sheet (dedup by Name + Phone)
// ------------------------------------------------------------
// SETUP / UPDATE:
//  1. Paste this into Extensions > Apps Script.
//  2. Deploy > Manage deployments > (pencil) > Version: New version > Deploy
//     (Required! Editing code without a new version keeps the old code live.)
//
// FLOW:
//  - Page asks GET ?action=check&name=..&phone=..&callback=.. (JSONP)
//    -> { exists: true|false }
//  - If guest already responded: page says "We already have your response,
//    submit again to change it". Second submit = update.
//  - Page POSTs JSON -> appended, or the existing row is updated.
// ============================================================

var SHEET_NAME = "RSVPs";
var HEADERS = ["Timestamp", "Name", "Phone", "Attending", "Guests", "Message"];

function norm_(s) {
  return String(s == null ? "" : s).trim().toLowerCase();
}
function digits_(s) {
  return norm_(s).replace(/[^0-9]/g, "");
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
    return sheet;
  }

  // Migrate old 5-column layout (no Phone) if present.
  var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (header.length < 6 || norm_(header[2]) !== "phone") {
    sheet.insertColumnAfter(2); // new empty column C
    sheet.getRange(1, 3).setValue("Phone");
  }
  return sheet;
}

// Returns the row number (>=2) matching name+phone, or -1.
function findRow_(sheet, name, phone) {
  var last = sheet.getLastRow();
  if (last < 2) return -1;
  var n = norm_(name), p = digits_(phone);
  var values = sheet.getRange(2, 2, last - 1, 2).getValues(); // cols B (Name), C (Phone)
  for (var i = 0; i < values.length; i++) {
    if (norm_(values[i][0]) === n && digits_(values[i][1]) === p) return i + 2;
  }
  return -1;
}

// GET — JSONP check (+ plain health check when no callback given)
function doGet(e) {
  var p = (e && e.parameter) || {};
  var payload = { ok: true };

  if (p.action === "check") {
    payload.exists = findRow_(getSheet_(), p.name, p.phone) > 0;
  }

  var json = JSON.stringify(payload);
  var cb = String(p.callback || "");
  if (/^[A-Za-z_$][A-Za-z0-9_$]*$/.test(cb)) {
    return ContentService.createTextOutput(cb + "(" + json + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// POST — create or update (dedup by name+phone)
function doPost(e) {
  var out = { ok: false };
  try {
    var data = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    var sheet = getSheet_();
    var row = findRow_(sheet, data.name, data.phone);
    var line = [
      new Date(),
      String(data.name || ""),
      String(data.phone || ""),
      String(data.attending || ""),
      Number(data.guests) || 0,
      String(data.message || "")
    ];

    if (row > 0) {
      sheet.getRange(row, 1, 1, HEADERS.length).setValues([line]);
      out.updated = true;
    } else {
      sheet.appendRow(line);
      out.updated = false;
    }
    out.ok = true;
  } catch (err) {
    out = { ok: false, error: String(err) };
    Logger.log("RSVP ERROR: " + err);
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}
