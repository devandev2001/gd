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
 * 7. SET UP DAILY REPORT TRIGGER:
 *    In Apps Script, go to Triggers (clock icon on left) →
 *    + Add Trigger →
 *      Function: generateDailyReport
 *      Event source: Time-driven
 *      Type: Day timer
 *      Time: 8pm to 9pm
 *    → Save
 *
 * That's it! The form will now write data to your sheet,
 * and a daily report tab is auto-created at 8 PM.
 */

// ═══════════════════════════════════════════════════════════
// 1. RECEIVE FORM SUBMISSIONS
// ═══════════════════════════════════════════════════════════

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
        "Caste Weight",
        "Gender Weight",
        "Age Weight",
        "Vote in 2021 AE",
        "Vote in 2024 GE",
        "Vote in 2026 AE",
        "Who Will Win",
        "Who Will Win Normalized"
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
      data.whoWillWin,
      data.normalizedScore
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

// ═══════════════════════════════════════════════════════════
// 2. DAILY REPORT — runs at 8 PM via time-trigger
// ═══════════════════════════════════════════════════════════

function generateDailyReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheets()[0]; // first sheet = raw data

  // Today's date string for the tab name (e.g. "24-Mar-2026")
  var today = new Date();
  var tabName = Utilities.formatDate(today, "Asia/Kolkata", "dd-MMM-yyyy");

  // Delete existing tab with same name (if re-running)
  var existing = ss.getSheetByName(tabName);
  if (existing) ss.deleteSheet(existing);

  // Get all data rows (skip header)
  var lastRow = dataSheet.getLastRow();
  if (lastRow < 2) return; // no data

  var data = dataSheet.getRange(2, 1, lastRow - 1, 11).getValues();
  // Columns: 0=Timestamp, 1=AC, 2=FA, 3=Caste, 4=Gender, 5=Age,
  //          6=Vote2021, 7=Vote2024, 8=Vote2026, 9=WhoWillWin, 10=NormalizedScore

  // Today's date string for filtering (match dd/mm/yyyy or d/m/yyyy format)
  var todayStr = Utilities.formatDate(today, "Asia/Kolkata", "d/M/yyyy");
  var todayStr2 = Utilities.formatDate(today, "Asia/Kolkata", "dd/MM/yyyy");
  var todayStr3 = Utilities.formatDate(today, "Asia/Kolkata", "d/M/yy");

  // Filter only today's entries
  var todayData = data.filter(function(row) {
    var ts = String(row[0]);
    return ts.indexOf(todayStr) !== -1 || ts.indexOf(todayStr2) !== -1 || ts.indexOf(todayStr3) !== -1;
  });

  if (todayData.length === 0) return; // no entries today

  // Parties we track
  var parties = ["LDF", "UDF", "BJP/NDA", "Others"];

  // Build AC → party → { sum, count }
  var acMap = {};
  for (var i = 0; i < todayData.length; i++) {
    var row = todayData[i];
    var ac = String(row[1]).trim();
    var party = String(row[9]).trim();
    var score = parseFloat(row[10]) || 0;

    if (!acMap[ac]) {
      acMap[ac] = {};
      for (var p = 0; p < parties.length; p++) {
        acMap[ac][parties[p]] = { sum: 0, count: 0 };
      }
    }

    if (acMap[ac][party]) {
      acMap[ac][party].sum += score;
      acMap[ac][party].count += 1;
    }
  }

  // Create report sheet
  var reportSheet = ss.insertSheet(tabName);

  // Header row
  var header = [
    "Assembly Constituency",
    "Total Entries",
    "LDF Avg",
    "UDF Avg",
    "BJP/NDA Avg",
    "Others Avg",
    "Predicted Winner"
  ];
  reportSheet.getRange(1, 1, 1, header.length).setValues([header]);

  // Style header
  var headerRange = reportSheet.getRange(1, 1, 1, header.length);
  headerRange.setBackground("#ff9933");
  headerRange.setFontColor("#ffffff");
  headerRange.setFontWeight("bold");
  headerRange.setHorizontalAlignment("center");

  // Fill data rows
  var acNames = Object.keys(acMap).sort();
  var reportRows = [];

  for (var a = 0; a < acNames.length; a++) {
    var acName = acNames[a];
    var partyData = acMap[acName];
    var totalEntries = 0;
    var avgs = {};
    var maxAvg = -1;
    var winner = "-";

    for (var p = 0; p < parties.length; p++) {
      var pt = parties[p];
      var d = partyData[pt];
      totalEntries += d.count;
      avgs[pt] = d.count > 0 ? d.sum / d.count : 0;
      if (avgs[pt] > maxAvg) {
        maxAvg = avgs[pt];
        winner = pt;
      }
    }

    reportRows.push([
      acName,
      totalEntries,
      avgs["LDF"] ? avgs["LDF"].toFixed(8) : "0",
      avgs["UDF"] ? avgs["UDF"].toFixed(8) : "0",
      avgs["BJP/NDA"] ? avgs["BJP/NDA"].toFixed(8) : "0",
      avgs["Others"] ? avgs["Others"].toFixed(8) : "0",
      maxAvg > 0 ? winner : "No data"
    ]);
  }

  if (reportRows.length > 0) {
    reportSheet.getRange(2, 1, reportRows.length, header.length).setValues(reportRows);
  }

  // Auto-resize columns
  for (var c = 1; c <= header.length; c++) {
    reportSheet.autoResizeColumn(c);
  }

  // Highlight winner column with green
  if (reportRows.length > 0) {
    var winnerRange = reportSheet.getRange(2, 7, reportRows.length, 1);
    winnerRange.setFontWeight("bold");
    winnerRange.setFontColor("#138808");
  }

  // Add summary row at the bottom
  var summaryRow = reportRows.length + 3;
  reportSheet.getRange(summaryRow, 1).setValue("Report generated at:");
  reportSheet.getRange(summaryRow, 2).setValue(
    Utilities.formatDate(new Date(), "Asia/Kolkata", "dd-MMM-yyyy hh:mm a")
  );
  reportSheet.getRange(summaryRow, 1, 1, 2).setFontStyle("italic").setFontColor("#718096");
}
