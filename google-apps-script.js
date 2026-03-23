/**
 * Google Apps Script — Kerala Survey 2026
 *
 * HOW TO SET UP:
 * 1. Open your Google Sheet:
 *    https://docs.google.com/spreadsheets/d/1kSuwyGwwgW_eMojhRBP5iFmVvdHzvtcanTndeIld1E4/edit
 *
 * 2. Go to  Extensions → Apps Script
 *
 * 3. Delete everything in Code.gs and paste THIS entire file.
 *
 * 4. Click  Deploy → New deployment
 *      • Type  = Web app
 *      • Execute as = Me
 *      • Who has access = Anyone
 *
 * 5. Copy the Web App URL and paste it into:
 *    src/data/surveyData.js → GOOGLE_SCRIPT_URL
 *
 * 6. Click "Authorize" when prompted.
 *
 * That's it! The form will now write data to your sheet.
 */

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);

    // Add header row if sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "Timestamp",
        "Assembly Constituency",
        "FA Name",
        "Caste",
        "Gender",
        "Age Group",
        "Vote in 2021 AE",
        "Vote in 2024 GE",
        "Vote in 2026 AE",
        "Who Will Win"
      ]);
    }

    sheet.appendRow([
      data.timestamp,
      data.ac,
      data.faName,
      data.caste,
      data.gender,
      data.age,
      data.vote2021,
      data.vote2024,
      data.vote2026,
      data.whoWillWin
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput("Kerala Survey 2026 API is running.")
    .setMimeType(ContentService.MimeType.TEXT);
}
