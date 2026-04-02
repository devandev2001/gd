    var d = parseInt(m[1], 10), mo = parseInt(m[2], 10), y = parseInt(m[3], 10);
    if (y < 100) y += 2000;
    return y + "-" + pad2(mo) + "-" + pad2(d);
  }

  m = s.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[1] + "-" + m[2] + "-" + m[3];

  var monthMap = { Jan:1, Feb:2, Mar:3, Apr:4, May:5, Jun:6, Jul:7, Aug:8, Sep:9, Oct:10, Nov:11, Dec:12 };
  var mon = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})/);
  if (mon) {
    var key = mon[2].charAt(0).toUpperCase() + mon[2].slice(1).toLowerCase();
    var moNum = monthMap[key];
    if (moNum) return mon[3] + "-" + pad2(moNum) + "-" + pad2(mon[1]);
  }
  return null;
}

function parseDateParamToYmd(dateStr) {
  if (!dateStr) return null;
  var s = String(dateStr).trim();

  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return m[1] + "-" + m[2] + "-" + m[3];

  var monthMap = { Jan:1, Feb:2, Mar:3, Apr:4, May:5, Jun:6, Jul:7, Aug:8, Sep:9, Oct:10, Nov:11, Dec:12 };
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

function timestampPatternsForDateKey(dateStr) {
  var tz = "Asia/Kolkata";
  var patterns = [String(dateStr)];
  var monthMap = { Jan:1, Feb:2, Mar:3, Apr:4, May:5, Jun:6, Jul:7, Aug:8, Sep:9, Oct:10, Nov:11, Dec:12 };

  var mon = dateStr.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/);
  if (mon) {
    var day = parseInt(mon[1], 10), year = parseInt(mon[3], 10);
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

function legacySubstringRowMatch(tsCell, dateStr) {
  if (!dateStr) return true;
  var ts = String(tsCell);
  var patterns = timestampPatternsForDateKey(dateStr);
  for (var i = 0; i < patterns.length; i++) if (ts.indexOf(patterns[i]) !== -1) return true;
  return false;
}

/** Prefer formatted cell text for date logic — avoids US-locale Date serials becoming the wrong calendar day. */
function timestampForDateFilter(displayVal, rawVal) {
  if (displayVal !== null && displayVal !== undefined && String(displayVal).trim() !== "") {
    return displayVal;
  }
  return rawVal;
}

function rowMatchesDateFilter(tsCell, dateStr) {
  if (!dateStr) return true;
  var targetYmd = parseDateParamToYmd(dateStr);
  if (targetYmd) {
    var rowYmd = cellToYmdKolkata(tsCell);
    if (rowYmd) return rowYmd === targetYmd;
  }
  return legacySubstringRowMatch(tsCell, dateStr);
}

function mapRowsToEntries(filtered, displayTimestamps, rowNumbers) {
  var headers = [
    "timestamp","ac","faName",
    "casteWeight","casteLabel",
    "genderWeight","genderLabel",
    "ageWeight","ageLabel",
    "vote2021","vote2024","vote2026",
    "whoWillWin",
    "rawNormalizedScore",
    "gevsve",
    "finalValue"
  ];

  return filtered.map(function(row, idx) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i]; });
    obj.normalizedScore = row[15];
    if (displayTimestamps && displayTimestamps[idx] !== undefined && displayTimestamps[idx] !== null && String(displayTimestamps[idx]).trim() !== "") {
      obj.timestamp = displayTimestamps[idx];
    }
    // Sheet row number (row 2 = first data row, so sheetRow = originalIndex + 2)
    if (rowNumbers && rowNumbers[idx] !== undefined) {
      obj.sheetRow = rowNumbers[idx];
    }
    return obj;
  });
}

function getEntriesForDate(dateStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME) || ss.getSheets()[0];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { entries: [], total: 0 };

  var numRows = lastRow - 1;
  var data = sheet.getRange(2, 1, numRows, 16).getValues();
  var displayA = sheet.getRange(2, 1, numRows, 1).getDisplayValues();
  var filtered = [];
  var dispTs = [];
  var rowNums = [];
  var i;
  for (i = 0; i < numRows; i++) {
    var tsF = timestampForDateFilter(displayA[i][0], data[i][0]);
    if (!dateStr || rowMatchesDateFilter(tsF, dateStr)) {
      filtered.push(data[i]);
      dispTs.push(displayA[i][0]);
      rowNums.push(i + 2); // +2: row 1 is header, data starts at row 2
    }
  }
  var entries = mapRowsToEntries(filtered, dispTs, rowNums);
  return { entries: entries, total: entries.length };
}

function getEntriesForDateRange(fromYmd, toYmd) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME) || ss.getSheets()[0];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { entries: [], total: 0 };

  var numRows = lastRow - 1;
  var data = sheet.getRange(2, 1, numRows, 16).getValues();
  var displayA = sheet.getRange(2, 1, numRows, 1).getDisplayValues();
  var filtered = [];
  var dispTs = [];
  var rowNums = [];
  for (var j = 0; j < numRows; j++) {
    var tsF = timestampForDateFilter(displayA[j][0], data[j][0]);
    var rowYmd = cellToYmdKolkata(tsF);
    var include = false;
    if (rowYmd) {
      include = rowYmd >= fromYmd && rowYmd <= toYmd;
    } else if (fromYmd === toYmd) {
      include = rowMatchesDateFilter(tsF, fromYmd);
    }
    if (include) {
      filtered.push(data[j]);
      dispTs.push(displayA[j][0]);
      rowNums.push(j + 2); // +2: row 1 is header
    }
  }

  var entries = mapRowsToEntries(filtered, dispTs, rowNums);
  return { entries: entries, total: entries.length };
}

function getSummaryForDate(tabName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(tabName);
  if (!sheet) return { rows: [], tabName: tabName, found: false };

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { rows: [], tabName: tabName, found: true };

  var numRows = lastRow - 1;
  var data = sheet.getRange(2, 1, numRows, 7).getValues();
  var rows = data
    .filter(function(row) { return String(row[0]).trim() !== "" && String(row[0]).indexOf("Report") === -1; })
    .map(function(row) {
      return {
        ac: row[0], totalEntries: row[1],
        ldf: row[2], udf: row[3], bjp: row[4], others: row[5], winner: row[6]
      };
    });

  return { rows: rows, tabName: tabName, found: true };
}

function getAvailableDates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var dates = [];
  for (var i = 0; i < sheets.length; i++) {
    var n = sheets[i].getName();
    if (n !== MAIN_SHEET_NAME && n !== DEMOGRAPHICS_SHEET_NAME) dates.push(n);
  }
  return { dates: dates };
}

/**
 * Prefer label columns E,G,I when present so weights match what the user originally chose.
 */
function pickCasteInputForResolve(row) {
  if (isValidCasteLabel(row[4])) return canonicalCasteKey(row[4]);
  return row[3];
}

function pickGenderInputForResolve(row) {
  if (isValidGenderLabel(row[6])) return row[6];
  return row[5];
}

function pickAgeInputForResolve(row) {
  if (isValidAgeLabel(row[8])) return normalizeAgeLabel(row[8]);
  return row[7];
}

function fixTextRows() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME) || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log("No rows."); return; }

  ensureHeaders(sheet);
  var numRows = lastRow - 1;
  var data = sheet.getRange(2, 1, numRows, 16).getValues();

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var ac = normalizeAcName(row[1]);
    var v26 = normalizeParty(row[11]);
    var ww = normalizeParty(row[12]);

    var cIn = pickCasteInputForResolve(row);
    var gIn = pickGenderInputForResolve(row);
    var aIn = pickAgeInputForResolve(row);

    var w = resolveWeights(ac, cIn, gIn, aIn);

    var casteLabel = casteLabelForSheet(ac, cIn, w.casteW);
    var genderLabel = genderLabelForSheet(ac, gIn, w.genderW);
    var ageLabel = ageLabelForSheet(aIn, w.ageW);

    var r = i + 2;
    sheet.getRange(r, 2).setValue(ac);
    sheet.getRange(r, 4).setValue(w.casteW);
    sheet.getRange(r, 5).setValue(casteLabel);
    sheet.getRange(r, 6).setValue(w.genderW);
    sheet.getRange(r, 7).setValue(genderLabel);
    sheet.getRange(r, 8).setValue(w.ageW);
    sheet.getRange(r, 9).setValue(ageLabel);
    sheet.getRange(r, 12).setValue(v26);
    sheet.getRange(r, 13).setValue(ww);
    sheet.getRange(r, 14).setValue(w.norm);
  }

  sheet.getRange(2, 4, numRows, 1).setNumberFormat("0.00000000");
  sheet.getRange(2, 6, numRows, 1).setNumberFormat("0.00000000");
  sheet.getRange(2, 8, numRows, 1).setNumberFormat("0.00000000");
  sheet.getRange(2, 14, numRows, 1).setNumberFormat("0.0000000000");

  recalcGevsveAndFinalValues();
  Logger.log("fixTextRows updated " + numRows + " rows.");
}

function recalcAllNormalizedScores() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME) || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lr = sheet.getLastRow();
  if (lr < 2) { Logger.log("No rows."); return; }

  var numRows = lr - 1;
  var data = sheet.getRange(2, 1, numRows, 16).getValues();
  var updated = 0;

  for (var i = 0; i < data.length; i++) {
    var cIn = pickCasteInputForResolve(data[i]);
    var gIn = pickGenderInputForResolve(data[i]);
    var aIn = pickAgeInputForResolve(data[i]);
    var ac = normalizeAcName(data[i][1]);
    var w = resolveWeights(ac, cIn, gIn, aIn);
    sheet.getRange(i + 2, 14).setValue(w.norm);
    updated++;
  }

  sheet.getRange(2, 14, numRows, 1).setNumberFormat("0.0000000000");
  recalcGevsveAndFinalValues();

  Logger.log("recalcAllNormalizedScores updated " + updated + " rows.");
}

function buildWeightedReportRows(dataRows, partyColumnIndex) {
  var parties = ["LDF", "UDF", "BJP/NDA"];
  var acMap = {};

  for (var i = 0; i < dataRows.length; i++) {
    var row = dataRows[i];
    var ac = normalizeAcName(row[1]);
    var party = normalizeParty(row[partyColumnIndex]);

    var score = asFiniteNumber(row[15], 0);

    if (!acMap[ac]) {
      acMap[ac] = { totalRows: 0, othersSum: 0, parties: {} };
      for (var p = 0; p < parties.length; p++) acMap[ac].parties[parties[p]] = { sum: 0, count: 0 };
    }

    acMap[ac].totalRows++;
    if (party === "Others") {
      acMap[ac].othersSum += score;
    } else {
      acMap[ac].parties[party].sum += score;
      acMap[ac].parties[party].count += 1;
    }
  }

  var acNames = Object.keys(acMap).sort();
  var reportRows = [];

  for (var a = 0; a < acNames.length; a++) {
    var acName = acNames[a];
    var rec = acMap[acName];

    var partySums = {};
    var recognizedTotal = 0;
    var maxSum = 0;
    var winner = "No data";

    for (var p2 = 0; p2 < parties.length; p2++) {
      var pt = parties[p2];
      partySums[pt] = rec.parties[pt].sum;
      recognizedTotal += rec.parties[pt].sum;
      if (rec.parties[pt].sum > maxSum) { maxSum = rec.parties[pt].sum; winner = pt; }
    }

    var denom = recognizedTotal + rec.othersSum;
    var ldfPct = denom > 0 ? (partySums["LDF"] / denom) * 100 : 0;
    var udfPct = denom > 0 ? (partySums["UDF"] / denom) * 100 : 0;
    var bjpPct = denom > 0 ? (partySums["BJP/NDA"] / denom) * 100 : 0;

    reportRows.push([
      acName,
      rec.totalRows,
      ldfPct.toFixed(2) + "%",
      udfPct.toFixed(2) + "%",
      bjpPct.toFixed(2) + "%",
      winner
    ]);
  }

  return reportRows;
}

function writeDailyReportSheet(ss, tabName, title, reportRows) {
  var existing = ss.getSheetByName(tabName);
  if (existing) ss.deleteSheet(existing);
  var sheet = ss.insertSheet(tabName);

  var header = ["Assembly Constituency","Total Entries","LDF %","UDF %","BJP/NDA %","Predicted Winner"];
  sheet.getRange(1, 1).setValue(title).setFontWeight("bold").setFontColor("#1d4ed8");
  sheet.getRange(2, 1, 1, header.length).setValues([header]);

  var h = sheet.getRange(2, 1, 1, header.length);
  h.setBackground("#1d4ed8");
  h.setFontColor("#ffffff");
  h.setFontWeight("bold");
  h.setHorizontalAlignment("center");

  if (reportRows.length > 0) {
    sheet.getRange(3, 1, reportRows.length, header.length).setValues(reportRows);
    sheet.getRange(3, 6, reportRows.length, 1).setFontWeight("bold").setFontColor("#1d4ed8");
  }

  for (var c = 1; c <= header.length; c++) sheet.autoResizeColumn(c);

  var summaryRow = reportRows.length + 5;
  sheet.getRange(summaryRow, 1).setValue("Report generated at:");
  sheet.getRange(summaryRow, 2).setValue(Utilities.formatDate(new Date(), "Asia/Kolkata", "dd-MMM-yyyy hh:mm a"));
  sheet.getRange(summaryRow, 1, 1, 2).setFontStyle("italic").setFontColor("#718096");
}

function generateDailyReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheetByName(MAIN_SHEET_NAME) || ss.getSheets()[0];
  var today = new Date();
  var baseName = Utilities.formatDate(today, "Asia/Kolkata", "dd-MMM-yyyy");

  var lastRow = dataSheet.getLastRow();
  if (lastRow < 2) return;

  var numRows = lastRow - 1;
  var data = dataSheet.getRange(2, 1, numRows, 16).getValues();

  var todayStr1 = Utilities.formatDate(today, "Asia/Kolkata", "d/M/yyyy");
  var todayStr2 = Utilities.formatDate(today, "Asia/Kolkata", "dd/MM/yyyy");
  var todayStr3 = Utilities.formatDate(today, "Asia/Kolkata", "d/M/yy");

  var todayData = data.filter(function(row) {
    var ts = String(row[0]);
    return ts.indexOf(todayStr1) !== -1 || ts.indexOf(todayStr2) !== -1 || ts.indexOf(todayStr3) !== -1;
  });

  if (todayData.length === 0) return;

  var wwRows = buildWeightedReportRows(todayData, 12);
  var v26Rows = buildWeightedReportRows(todayData, 11);

  writeDailyReportSheet(ss, baseName + "-WW", "Who Will Win (weighted by Final Values)", wwRows);
  writeDailyReportSheet(ss, baseName + "-V26", "Vote 2026 (weighted by Final Values)", v26Rows);
}
