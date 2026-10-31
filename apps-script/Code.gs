// ============================================================
// Wedding RSVP -> Google Sheet
// ------------------------------------------------------------
// SETUP (one time, ~2 minutes):
//  1. Create a new Google Sheet (this is where RSVPs are stored).
//  2. Extensions > Apps Script, delete the default code, paste this file.
//  3. (Optional) Change SHEET_NAME below if you renamed the tab.
//  4. Deploy > New deployment > Type: Web app
//       - Description: "RSVP"
//       - Execute as: Me
//       - Who has access: Anyone
//     Authorize when prompted (it's your own script).
//  5. Copy the Web app URL (ends in /exec) and paste it into
//     index.html -> CONFIG.rsvpEndpoint
// ============================================================

var SHEET_NAME = "RSVPs";

function doPost(e) {
  var out = { ok: false };

  try {
    // The page sends JSON as a plain-text body (avoids CORS preflight).
    var data = JSON.parse((e && e.postData && e.postData.contents) || "{}");

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Timestamp", "Name", "Attending", "Guests", "Message"]);
      sheet.getRange(1, 1, 1, 5).setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([
      new Date(),
      String(data.name || ""),
      String(data.attending || ""),
      Number(data.guests) || 0,
      String(data.message || "")
    ]);

    out = { ok: true };
  } catch (err) {
    out = { ok: false, error: String(err) };
  }

  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

// Needed so the browser can "see" a response if you later drop mode:"no-cors".
function doOptions() {
  return ContentService.createTextOutput("")
    .setMimeType(ContentService.MimeType.TEXT);
}
