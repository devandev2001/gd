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
 * 4. Click  Deploy → Manage deployments → Edit (pencil icon)
 *      • Version = New version
 *      → Deploy
 *
 * 5. The Web App URL stays the same (no need to update surveyData.js).
 *
 * 6. SET UP DAILY REPORT TRIGGER:
 *    In Apps Script, go to Triggers (clock icon on left) →
 *    + Add Trigger →
 *      Function: generateDailyReport
 *      Event source: Time-driven
 *      Type: Day timer
 *      Time: 8pm to 9pm
 *    → Save
 *
 * 7. RUN fixTextRows ONCE to clean up old text rows:
 *    Select fixTextRows in dropdown → ▶ Run
 */

// ═══════════════════════════════════════════════════════════
// DEMOGRAPHIC DATA (shared by doPost, fixTextRows, report)
// ═══════════════════════════════════════════════════════════

var AC_DEMOGRAPHICS = {
  Kattakkada:         { male:48.01, female:51.99, Nair:34.99, Ezhava:14.83, Muslim:6.04,  Christian:11.03, "SC/ST":11.03, Others:9.50  },
  Kovalam:            { male:48.88, female:51.11, Nair:14.71, Ezhava:14.35, Muslim:9.56,  Christian:11.49, "SC/ST":11.49, Others:12.59 },
  Vattiyoorkavu:      { male:47.56, female:52.44, Nair:35.23, Ezhava:10.24, Muslim:6.00,  Christian:9.75,  "SC/ST":9.75,  Others:12.74 },
  Thiruvananthapuram: { male:48.06, female:51.93, Nair:23.21, Ezhava:8.43,  Muslim:18.00, Christian:8.65,  "SC/ST":8.65,  Others:10.41 },
  Attingal:           { male:46.28, female:53.72, Nair:19.10, Ezhava:28.12, Muslim:17.30, Christian:17.21, "SC/ST":17.21, Others:13.96 },
  Chathannoor:        { male:46.81, female:53.19, Nair:26.85, Ezhava:29.60, Muslim:12.10, Christian:14.36, "SC/ST":14.36, Others:4.29  },
  Aranmula:           { male:47.95, female:52.05, Nair:20.89, Ezhava:20.89, Muslim:4.20,  Christian:15.30, "SC/ST":15.30, Others:0.86  },
  Thiruvalla:         { male:47.98, female:52.02, Nair:16.78, Ezhava:10.36, Muslim:0.00,  Christian:11.46, "SC/ST":11.46, Others:10.03 },
  Chengannur:         { male:47.64, female:52.35, Nair:29.92, Ezhava:15.54, Muslim:3.89,  Christian:15.75, "SC/ST":15.75, Others:5.71  },
  Adoor:              { male:47.21, female:52.79, Nair:25.15, Ezhava:19.69, Muslim:6.80,  Christian:18.60, "SC/ST":18.60, Others:1.84  },
  Poonjar:            { male:49.51, female:50.49, Nair:7.30,  Ezhava:15.11, Muslim:20.39, Christian:11.37, "SC/ST":11.37, Others:2.45  },
  Pala:               { male:48.71, female:51.29, Nair:16.41, Ezhava:13.73, Muslim:1.58,  Christian:8.39,  "SC/ST":8.39,  Others:2.20  },
  Thrissur:           { male:47.55, female:52.45, Nair:18.03, Ezhava:17.11, Muslim:5.20,  Christian:7.66,  "SC/ST":7.66,  Others:8.61  },
  Kunnathunad:        { male:48.85, female:51.14, Nair:11.78, Ezhava:14.57, Muslim:19.70, Christian:12.69, "SC/ST":12.69, Others:4.36  },
  Palakkad:           { male:48.63, female:51.37, Nair:9.66,  Ezhava:22.08, Muslim:27.84, Christian:11.75, "SC/ST":11.75, Others:9.90  },
  "Kozhikode North":  { male:47.41, female:52.59, Nair:14.07, Ezhava:33.16, Muslim:25.10, Christian:0.00,  "SC/ST":0.00,  Others:11.55 },
  Kasaragod:          { male:50.00, female:50.00, Nair:3.30,  Ezhava:15.00, Muslim:50.42, Christian:5.43,  "SC/ST":5.43,  Others:15.20 },
  Manjeshwaram:       { male:50.38, female:49.62, Nair:0.00,  Ezhava:12.00, Muslim:52.89, Christian:4.77,  "SC/ST":4.77,  Others:20.01 },
};

var AGE_WEIGHTS = {
  "18-19": 0.01574992977,
  "20-29": 0.16726013,
  "30-39": 0.1839168018,
  "40-49": 0.2081028821,
  "50-59": 0.19009771,
  "60-69": 0.1395625022,
  "70-79": 0.07469513213,
  "80+":   0.02061491203,
};

var CASTE_LIST = ["Nair","Ezhava","Muslim","Christian","SC/ST","Others"];
var GENDER_LIST = ["Male","Female"];

function isTextLabel(val) {
  var s = String(val).trim();
  if (s === "") return false;
  return isNaN(parseFloat(s));
}

function resolveWeights(ac, casteVal, genderVal, ageVal) {
  var acData = AC_DEMOGRAPHICS[ac];
  var casteW, genderW, ageW;

  if (isTextLabel(casteVal)) {
    casteW = acData ? (acData[String(casteVal).trim()] || 0) / 100 : 0;
  } else {
    casteW = parseFloat(casteVal) || 0;
  }

  if (isTextLabel(genderVal)) {
    var g = String(genderVal).trim();
    genderW = acData ? (g === "Male" ? acData.male : acData.female) / 100 : 0;
  } else {
    genderW = parseFloat(genderVal) || 0;
  }

  if (isTextLabel(ageVal)) {
    ageW = AGE_WEIGHTS[String(ageVal).trim()] || 0;
  } else {
    ageW = parseFloat(ageVal) || 0;
  }

  return { casteW: casteW, genderW: genderW, ageW: ageW, norm: casteW * genderW * ageW };
}

// ═══════════════════════════════════════════════════════════
// 1. RECEIVE FORM SUBMISSIONS
//    Now auto-converts text labels to numbers server-side
// ═══════════════════════════════════════════════════════════

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);

    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "Timestamp", "Assembly Constituency", "FA Name",
        "Caste Weight", "Gender Weight", "Age Weight",
        "Vote in 2021 AE", "Vote in 2024 GE", "Vote in 2026 AE",
        "Who Will Win", "Who Will Win Normalized"
      ]);
    }

    var w = resolveWeights(data.ac, data.caste, data.gender, data.age);

    sheet.appendRow([
      data.timestamp,
      data.ac,
      data.faName,
      w.casteW,
      w.genderW,
      w.ageW,
      data.vote2021,
      data.vote2024,
      data.vote2026,
      data.whoWillWin,
      w.norm
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
// 2. ONE-TIME CLEANUP — run manually once to fix old text rows
//    Select fixTextRows → ▶ Run
// ═══════════════════════════════════════════════════════════

function fixTextRows() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log("No data rows found."); return; }

  var data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
  var fixed = 0;

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var ac         = String(row[1]).trim();
    var casteCell  = row[3];
    var genderCell = row[4];
    var ageCell    = row[5];

    var casteIsText  = isTextLabel(casteCell);
    var genderIsText = isTextLabel(genderCell);
    var ageIsText    = isTextLabel(ageCell);

    if (!casteIsText && !genderIsText && !ageIsText) continue;

    var w = resolveWeights(ac, casteCell, genderCell, ageCell);

    var sheetRow = i + 2;
    sheet.getRange(sheetRow, 4).setValue(w.casteW);
    sheet.getRange(sheetRow, 5).setValue(w.genderW);
    sheet.getRange(sheetRow, 6).setValue(w.ageW);
    sheet.getRange(sheetRow, 11).setValue(w.norm);

    Logger.log("Fixed row " + sheetRow + ": " + ac +
      " | " + casteCell + " → " + w.casteW +
      " | " + genderCell + " → " + w.genderW +
      " | " + ageCell + " → " + w.ageW);
    fixed++;
  }

  Logger.log("Done. Fixed " + fixed + " row(s).");
}

// ═══════════════════════════════════════════════════════════
// 3. DAILY REPORT — runs at 8 PM via time-trigger
//    Party % = party score sum / grand total × 100
// ═══════════════════════════════════════════════════════════

function generateDailyReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheets()[0];

  var today = new Date();
  var tabName = Utilities.formatDate(today, "Asia/Kolkata", "dd-MMM-yyyy");

  var existing = ss.getSheetByName(tabName);
  if (existing) ss.deleteSheet(existing);

  var lastRow = dataSheet.getLastRow();
  if (lastRow < 2) return;

  var data = dataSheet.getRange(2, 1, lastRow - 1, 11).getValues();

  var todayStr  = Utilities.formatDate(today, "Asia/Kolkata", "d/M/yyyy");
  var todayStr2 = Utilities.formatDate(today, "Asia/Kolkata", "dd/MM/yyyy");
  var todayStr3 = Utilities.formatDate(today, "Asia/Kolkata", "d/M/yy");

  var todayData = data.filter(function(row) {
    var ts = String(row[0]);
    return ts.indexOf(todayStr) !== -1 || ts.indexOf(todayStr2) !== -1 || ts.indexOf(todayStr3) !== -1;
  });

  if (todayData.length === 0) return;

  var parties = ["LDF", "UDF", "BJP/NDA", "Others"];

  var acMap = {};
  for (var i = 0; i < todayData.length; i++) {
    var row = todayData[i];
    var ac    = String(row[1]).trim();
    var party = String(row[9]).trim();
    var score = parseFloat(row[10]) || 0;

    if (!acMap[ac]) {
      acMap[ac] = {};
      for (var p = 0; p < parties.length; p++) acMap[ac][parties[p]] = { sum: 0, count: 0 };
    }
    if (acMap[ac][party]) {
      acMap[ac][party].sum += score;
      acMap[ac][party].count += 1;
    }
  }

  var reportSheet = ss.insertSheet(tabName);

  var header = [
    "Assembly Constituency", "Total Entries",
    "LDF %", "UDF %", "BJP/NDA %", "Others %",
    "Predicted Winner"
  ];
  reportSheet.getRange(1, 1, 1, header.length).setValues([header]);

  var headerRange = reportSheet.getRange(1, 1, 1, header.length);
  headerRange.setBackground("#1d4ed8");
  headerRange.setFontColor("#ffffff");
  headerRange.setFontWeight("bold");
  headerRange.setHorizontalAlignment("center");

  var acNames = Object.keys(acMap).sort();
  var reportRows = [];

  for (var a = 0; a < acNames.length; a++) {
    var acName = acNames[a];
    var partyData = acMap[acName];
    var totalEntries = 0;
    var partySums = {};
    var grandTotal = 0;

    for (var p = 0; p < parties.length; p++) {
      var pt = parties[p];
      var d = partyData[pt];
      totalEntries += d.count;
      partySums[pt] = d.sum;
      grandTotal += d.sum;
    }

    var partyPct = {};
    var maxPct = -1;
    var winner = "-";

    for (var p = 0; p < parties.length; p++) {
      var pt = parties[p];
      partyPct[pt] = grandTotal > 0 ? (partySums[pt] / grandTotal) * 100 : 0;
      if (partyPct[pt] > maxPct) {
        maxPct = partyPct[pt];
        winner = pt;
      }
    }

    reportRows.push([
      acName,
      totalEntries,
      partyPct["LDF"].toFixed(2) + "%",
      partyPct["UDF"].toFixed(2) + "%",
      partyPct["BJP/NDA"].toFixed(2) + "%",
      partyPct["Others"].toFixed(2) + "%",
      grandTotal > 0 ? winner : "No data"
    ]);
  }

  if (reportRows.length > 0) {
    reportSheet.getRange(2, 1, reportRows.length, header.length).setValues(reportRows);
    reportSheet.getRange(2, 7, reportRows.length, 1).setFontWeight("bold").setFontColor("#1d4ed8");
  }

  for (var c = 1; c <= header.length; c++) reportSheet.autoResizeColumn(c);

  var summaryRow = reportRows.length + 3;
  reportSheet.getRange(summaryRow, 1).setValue("Report generated at:");
  reportSheet.getRange(summaryRow, 2).setValue(
    Utilities.formatDate(new Date(), "Asia/Kolkata", "dd-MMM-yyyy hh:mm a")
  );
  reportSheet.getRange(summaryRow, 1, 1, 2).setFontStyle("italic").setFontColor("#718096");
}
