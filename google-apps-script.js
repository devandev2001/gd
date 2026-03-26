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
  var action = e && e.parameter && e.parameter.action;

  try {
    if (action === "entries") {
      var fromP = e.parameter.from;
      var toP = e.parameter.to;
      if (fromP && toP) {
        return jsonResponse(getEntriesForDateRange(String(fromP), String(toP)));
      }
      return jsonResponse(getEntriesForDate(e.parameter.date || ""));
    }
    if (action === "summary") {
      return jsonResponse(getSummaryForDate(e.parameter.date || ""));
    }
    if (action === "dates") {
      return jsonResponse(getAvailableDates());
    }
    // Health check
    return ContentService
      .createTextOutput("Kerala Survey 2026 API is running.")
      .setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return jsonResponse({ error: err.toString() });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function pad2(n) {
  var x = parseInt(n, 10);
  return x < 10 ? "0" + x : String(x);
}

// Normalize timestamp cell (Date or string) to yyyy-MM-dd in Asia/Kolkata.
function cellToYmdKolkata(cell) {
  if (cell instanceof Date) {
    if (isNaN(cell.getTime())) return null;
    return Utilities.formatDate(cell, "Asia/Kolkata", "yyyy-MM-dd");
  }
  var s = String(cell);
  var m = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (m) {
    var d = parseInt(m[1], 10);
    var mo = parseInt(m[2], 10);
    var y = parseInt(m[3], 10);
    if (y < 100) y += 2000;
    return y + "-" + pad2(mo) + "-" + pad2(d);
  }
  m = s.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[1] + "-" + m[2] + "-" + m[3];
  var monthMap = {
    Jan: 1, Feb: 2, Mar: 3, Apr: 4, May: 5, Jun: 6,
    Jul: 7, Aug: 8, Sep: 9, Oct: 10, Nov: 11, Dec: 12
  };
  var mon = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})/);
  if (mon) {
    var key = mon[2].charAt(0).toUpperCase() + mon[2].slice(1).toLowerCase();
    var moNum = monthMap[key];
    if (moNum) return mon[3] + "-" + pad2(moNum) + "-" + pad2(mon[1]);
  }
  return null;
}

// API date param → yyyy-MM-dd (single-day filter).
function parseDateParamToYmd(dateStr) {
  if (!dateStr) return null;
  var s = String(dateStr).trim();
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return m[1] + "-" + m[2] + "-" + m[3];
  var monthMap = {
    Jan: 1, Feb: 2, Mar: 3, Apr: 4, May: 5, Jun: 6,
    Jul: 7, Aug: 8, Sep: 9, Oct: 10, Nov: 11, Dec: 12
  };
  var mon = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/);
  if (mon) {
    var key = mon[2].charAt(0).toUpperCase() + mon[2].slice(1).toLowerCase();
    var moNum = monthMap[key];
    if (moNum) return mon[3] + "-" + pad2(moNum) + "-" + pad2(mon[1]);
  }
  var sl = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (sl) return sl[3] + "-" + pad2(sl[2]) + "-" + pad2(sl[1]);
  return null;
}

// Builds substring patterns that may appear in Sheet1 timestamps for the same calendar day.
// Summary tabs use "dd-MMM-yyyy" (e.g. 25-Mar-2026) but form timestamps usually use d/M/yyyy.
function timestampPatternsForDateKey(dateStr) {
  var tz = "Asia/Kolkata";
  var patterns = [String(dateStr)];
  var monthMap = {
    Jan: 1, Feb: 2, Mar: 3, Apr: 4, May: 5, Jun: 6,
    Jul: 7, Aug: 8, Sep: 9, Oct: 10, Nov: 11, Dec: 12
  };
  var mon = dateStr.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/);
  if (mon) {
    var day = parseInt(mon[1], 10);
    var year = parseInt(mon[3], 10);
    var key = mon[2].charAt(0).toUpperCase() + mon[2].slice(1).toLowerCase();
    var mo = monthMap[key];
    if (mo) {
      var cal = new Date(year, mo - 1, day);
      patterns.push(Utilities.formatDate(cal, tz, "d/M/yyyy"));
      patterns.push(Utilities.formatDate(cal, tz, "dd/MM/yyyy"));
      patterns.push(Utilities.formatDate(cal, tz, "d/M/yy"));
      patterns.push(Utilities.formatDate(cal, tz, "dd/MM/yy"));
      patterns.push(Utilities.formatDate(cal, tz, "yyyy-MM-dd"));
    }
  }
  return patterns;
}

function legacySubstringRowMatch(row, dateStr) {
  if (!dateStr) return true;
  var ts = String(row[0]);
  var patterns = timestampPatternsForDateKey(dateStr);
  for (var i = 0; i < patterns.length; i++) {
    if (ts.indexOf(patterns[i]) !== -1) return true;
  }
  return false;
}

function rowMatchesDateFilter(row, dateStr) {
  if (!dateStr) return true;
  var targetYmd = parseDateParamToYmd(dateStr);
  if (targetYmd) {
    var rowYmd = cellToYmdKolkata(row[0]);
    if (rowYmd) return rowYmd === targetYmd;
  }
  return legacySubstringRowMatch(row, dateStr);
}

function mapRowsToEntries(filtered) {
  var headers = ["timestamp","ac","faName","casteWeight","genderWeight","ageWeight",
                 "vote2021","vote2024","vote2026","whoWillWin","normalizedScore"];
  return filtered.map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i]; });
    return obj;
  });
}

// Returns all raw entries for a given date string (e.g. "2026-03-26", "25/3/2026", "25-Mar-2026")
function getEntriesForDate(dateStr) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { entries: [], total: 0 };

  // Rows 2 … lastRow (row 1 = header). Do not use lastRow-1 — that drops the final data row.
  var data = sheet.getRange(2, 1, lastRow, 11).getValues();

  var filtered = dateStr
    ? data.filter(function(row) { return rowMatchesDateFilter(row, dateStr); })
    : data;

  var entries = mapRowsToEntries(filtered);
  return { entries: entries, total: entries.length };
}

// Inclusive range using calendar yyyy-MM-dd (matches Sheet1 timestamps, not summary tab names).
function getEntriesForDateRange(fromYmd, toYmd) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { entries: [], total: 0 };

  var data = sheet.getRange(2, 1, lastRow, 11).getValues();
  var filtered = data.filter(function(row) {
    var rowYmd = cellToYmdKolkata(row[0]);
    if (rowYmd) return rowYmd >= fromYmd && rowYmd <= toYmd;
    if (fromYmd === toYmd) return rowMatchesDateFilter(row, fromYmd);
    return false;
  });
  var entries = mapRowsToEntries(filtered);
  return { entries: entries, total: entries.length };
}

// Returns summary data from a named tab (e.g. "25-Mar-2026")
function getSummaryForDate(tabName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(tabName);
  if (!sheet) return { rows: [], tabName: tabName, found: false };

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { rows: [], tabName: tabName, found: true };

  var data = sheet.getRange(2, 1, lastRow, 7).getValues();
  var rows = data
    .filter(function(row) { return String(row[0]).trim() !== "" && String(row[0]).indexOf("Report") === -1; })
    .map(function(row) {
      return {
        ac: row[0], totalEntries: row[1],
        ldf: row[2], udf: row[3], bjp: row[4], others: row[5],
        winner: row[6]
      };
    });

  return { rows: rows, tabName: tabName, found: true };
}

// Returns list of available summary tab names (excludes Sheet1/raw data sheet)
function getAvailableDates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var dates = [];
  for (var i = 1; i < sheets.length; i++) {
    dates.push(sheets[i].getName());
  }
  return { dates: dates };
}

// ═══════════════════════════════════════════════════════════
// 2. ONE-TIME CLEANUP — run manually once to fix old text rows
//    Select fixTextRows → ▶ Run
// ═══════════════════════════════════════════════════════════

function fixTextRows() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log("No data rows found."); return; }

  var data = sheet.getRange(2, 1, lastRow, 11).getValues();
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

  var data = dataSheet.getRange(2, 1, lastRow, 11).getValues();

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
