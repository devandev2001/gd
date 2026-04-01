
function ageLabelForSheet(formAge, ageW) {
  if (formAge != null && isValidAgeLabel(formAge)) return normalizeAgeLabel(formAge);
  return getAgeLabelFromWeight(ageW);
}

function getCasteLabelFromWeight(ac, casteWeight) {
  if (isTextLabel(casteWeight)) {
    var t = canonicalCasteKey(casteWeight);
    return isValidCasteLabel(t) ? t : "Others";
  }

  var acName = normalizeAcName(ac);
  var acData = getDemographicsMap()[acName];
  if (!acData) return "Unknown";

  var w = asFiniteNumber(casteWeight, NaN);
  if (!isFinite(w)) return "Unknown";
  if (w > 1.5) w = w / 100;
  if (w === 0) return "Unknown";

  var best = "Others";
  var bestDiff = Infinity;
  var i;
  for (i = 0; i < CASTE_LIST.length; i++) {
    var c = CASTE_LIST[i];
    var diff = Math.abs((acData[c] || 0) / 100 - w);
    if (diff < bestDiff) { bestDiff = diff; best = c; }
  }
  return best;
}

function getGenderLabelFromWeight(ac, genderWeight) {
  if (isTextLabel(genderWeight)) {
    var t = String(genderWeight).trim();
    if (isValidGenderLabel(t)) return displayGenderLabel(t);
    return "Unknown";
  }

  var acName = normalizeAcName(ac);
  var acData = getDemographicsMap()[acName];
  if (!acData) return "Unknown";

  var w = asFiniteNumber(genderWeight, NaN);
  if (!isFinite(w)) return "Unknown";
  if (w > 1.5) w = w / 100;
  if (w === 0) return "Unknown";

  var maleW = acData.male / 100;
  var femaleW = acData.female / 100;
  var md = Math.abs(maleW - w);
  var fd = Math.abs(femaleW - w);
  if (md < fd) return "Male";
  if (fd < md) return "Female";
  if (Math.abs(acData.male - acData.female) < 1e-9) return "Unknown";
  return femaleW >= maleW ? "Female" : "Male";
}

function getAgeLabelFromWeight(ageWeight) {
  return getAgeLabelFromAny(ageWeight);
}

/**
 * Numerator from GEVSVE_BASE, or null if (party, AC) is not in the Excel formula table.
 * Do not use numeric 1 as “missing” — a real numerator could be 1 and would break === 1 checks.
 */
function getBaseGevsveValue(party, ac) {
  var t = GEVSVE_BASE[party];
  if (!t || !t.hasOwnProperty(ac)) return null;
  return t[ac];
}

/** COUNTIFS($K:$K, party, $B:$B, AC) — column K index 10, B index 1 */
function buildCountifsMapForKandB(dataRows) {
  var map = {};
  for (var i = 0; i < dataRows.length; i++) {
    var ac = normalizeAcName(dataRows[i][1]);
    var party = normalizeParty(dataRows[i][10]);
    if (party !== "UDF" && party !== "LDF" && party !== "BJP/NDA") continue;
    var key = party + "||" + ac;
    map[key] = (map[key] || 0) + 1;
  }
  return map;
}

function calcGevsveForRow(ac, vote2024Party, countMap) {
  var acNorm = normalizeAcName(ac);
  var party = normalizeParty(vote2024Party);

  if (party !== "UDF" && party !== "LDF" && party !== "BJP/NDA") return 1;

  var base = getBaseGevsveValue(party, acNorm);
  if (base == null) return 1;

  var denom = countMap[party + "||" + acNorm] || 0;
  if (denom <= 0) return 1;

  return base / denom;
}

/**
 * After appendRow, only rows sharing the same (AC, 2024 party) as the new row need a new GevsVE.
 * Reads B, K, N only — avoids full-sheet 16-column read + rewrite on every form submit.
 * Use action=recalcOP (or fixTextRows) if sheet data was edited and O/P must be fully reconciled.
 */
function incrementalUpdateGevsveAfterAppend(sheet, lastRow) {
  var numRows = lastRow - 1;
  if (numRows < 1) return;

  var colB = sheet.getRange(2, 2, lastRow, 2).getValues();
  var colK = sheet.getRange(2, 11, lastRow, 11).getValues();
  var colN = sheet.getRange(2, 14, lastRow, 14).getValues();

  var countMap = {};
  var i;
  for (i = 0; i < numRows; i++) {
    var ac = normalizeAcName(colB[i][0]);
    var party = normalizeParty(colK[i][0]);
    if (party !== "UDF" && party !== "LDF" && party !== "BJP/NDA") continue;
    var key = party + "||" + ac;
    countMap[key] = (countMap[key] || 0) + 1;
  }

  var acNew = normalizeAcName(colB[numRows - 1][0]);
  var partyNew = normalizeParty(colK[numRows - 1][0]);
  var nLast = asFiniteNumber(colN[numRows - 1][0], 0);

  if (partyNew !== "UDF" && partyNew !== "LDF" && partyNew !== "BJP/NDA") {
    sheet.getRange(lastRow, 15, lastRow, 16).setValues([[1, nLast]]);
    sheet.getRange(lastRow, 15, lastRow, 16).setNumberFormat("0.0000000000");
    return;
  }

  var targetKey = partyNew + "||" + acNew;
  var base = getBaseGevsveValue(partyNew, acNew);
  var denom = countMap[targetKey] || 0;
  var oVal = (base == null || denom <= 0) ? 1 : base / denom;

  var matchIdx = [];
  for (i = 0; i < numRows; i++) {
    var acI = normalizeAcName(colB[i][0]);
    var partyI = normalizeParty(colK[i][0]);
    if ((partyI + "||" + acI) !== targetKey) continue;
    matchIdx.push(i);
  }

  if (matchIdx.length === 0) {
    sheet.getRange(lastRow, 15, lastRow, 16).setValues([[oVal, oVal * nLast]]);
    sheet.getRange(lastRow, 15, lastRow, 16).setNumberFormat("0.0000000000");
    return;
  }

  var m = 0;
  var firstSheetRow = matchIdx[0] + 2;
  var lastSheetRow = matchIdx[matchIdx.length - 1] + 2;
  while (m < matchIdx.length) {
    var blockStart = matchIdx[m];
    var blockEnd = blockStart;
    while (m + 1 < matchIdx.length && matchIdx[m + 1] === matchIdx[m] + 1) {
      m++;
      blockEnd = matchIdx[m];
    }
    var vals = [];
    var j;
    for (j = blockStart; j <= blockEnd; j++) {
      var nV = asFiniteNumber(colN[j][0], 0);
      vals.push([oVal, oVal * nV]);
    }
    sheet.getRange(blockStart + 2, 15, blockEnd + 2, 16).setValues(vals);
    m++;
  }

  sheet.getRange(firstSheetRow, 15, lastSheetRow, 16).setNumberFormat("0.0000000000");
}

/** Script property: "1" while a deferred O/P recalc trigger is scheduled (avoids trigger spam). */
var PROP_GEVSVE_RECALC_PENDING = "GEVSVE_RECALC_PENDING";

/**
 * Schedule full O:P recalc shortly after this request ends — form HTTP response returns fast.
 * Large Sheet1: skip incremental (avoids scanning B+K+N for 100k+ rows). If trigger creation fails on a large sheet,
 * runs one synchronous full recalc as last resort.
 */
function scheduleDeferredGevsveRecalc(sheet, lastRow) {
  var numDataRows = lastRow - 1;
  if (numDataRows >= 1 && numDataRows <= GEVSVE_INCREMENTAL_MAX_DATA_ROWS) {
    try {
      incrementalUpdateGevsveAfterAppend(sheet, lastRow);
    } catch (err) {
      Logger.log("incrementalUpdateGevsveAfterAppend: " + err);
    }
  }

  var props = PropertiesService.getScriptProperties();
  if (props.getProperty(PROP_GEVSVE_RECALC_PENDING) === "1") return;

  try {
    props.setProperty(PROP_GEVSVE_RECALC_PENDING, "1");
    ScriptApp.newTrigger("runDeferredGevsveRecalc")
      .timeBased()
      .after(2000)
      .create();
  } catch (err) {
    props.deleteProperty(PROP_GEVSVE_RECALC_PENDING);
    Logger.log("scheduleDeferredGevsveRecalc: " + err);
    if (numDataRows > GEVSVE_INCREMENTAL_MAX_DATA_ROWS) {
      try {
        recalcGevsveAndFinalValues();
      } catch (e2) {
        Logger.log("recalcGevsveAndFinalValues fallback: " + e2);
      }
    }
  }
}

/** Called by the one-shot clock trigger; then clears the pending flag so new submits can schedule again. */
function runDeferredGevsveRecalc() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(PROP_GEVSVE_RECALC_PENDING);
  try {
    recalcGevsveAndFinalValues();
  } catch (err) {
    Logger.log("runDeferredGevsveRecalc: " + err);
  }
}

/**
 * Optional: Triggers → Add trigger → pingWarm → Time-driven → Every 5–10 minutes.
 * Keeps the script/sheet binding warm so the first form submit after idle is not 15–30s cold start only.
 */
function pingWarm() {
  try {
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME);
  } catch (err) {
    Logger.log("pingWarm: " + err);
  }
}

function recalcGevsveAndFinalValues() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME) || ss.getSheets()[0];
  var lr = sheet.getLastRow();
  if (lr < 2) return;

  ensureHeaders(sheet);

  var numRows = lr - 1;
  var data = sheet.getRange(2, 1, numRows, 16).getValues();
  var countMap = buildCountifsMapForKandB(data);

  var out = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var ac = row[1];
    var vote2024 = row[10];
    var nVal = asFiniteNumber(row[13], 0);

    var oVal = calcGevsveForRow(ac, vote2024, countMap);
    var pVal = oVal * nVal;

    out.push([oVal, pVal]);
  }

  sheet.getRange(2, 15, out.length, 2).setValues(out);
  sheet.getRange(2, 15, out.length, 1).setNumberFormat("0.0000000000");
  sheet.getRange(2, 16, out.length, 1).setNumberFormat("0.0000000000");
}

function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(MAIN_SHEET_NAME) || ss.getSheets()[0];
    ensureHeaders(sheet);

    var data = JSON.parse(e.postData.contents);

    var ac = normalizeAcName(data.ac);
    var vote2026 = normalizeParty(data.vote2026);
    var whoWillWin = normalizeParty(data.whoWillWin);

    var w = resolveWeights(ac, data.caste, data.gender, data.age);
    var casteLabel = casteLabelForSheet(ac, data.caste, w.casteW);
    var genderLabel = genderLabelForSheet(ac, data.gender, w.genderW);
    var ageLabel = ageLabelForSheet(data.age, w.ageW);

    sheet.appendRow([
      data.timestamp,
      ac,
      data.faName,
      w.casteW, casteLabel,
      w.genderW, genderLabel,
      w.ageW, ageLabel,
      data.vote2021,
      vote2024ForSheet(data.vote2024),
      vote2026,
      whoWillWin,
      w.norm,
      1,
      w.norm
    ]);

    var lr = sheet.getLastRow();
    sheet.getRange(lr, 4).setNumberFormat("0.00000000");
    sheet.getRange(lr, 6).setNumberFormat("0.00000000");
    sheet.getRange(lr, 8).setNumberFormat("0.00000000");
    sheet.getRange(lr, 14).setNumberFormat("0.0000000000");

    scheduleDeferredGevsveRecalc(sheet, lr);

    return jsonResponse({ status: "success" });
  } catch (err) {
    return jsonResponse({ status: "error", message: String(err) });
  }
}

function doGet(e) {
  var action = e && e.parameter && e.parameter.action;
  try {
    if (action === "entries") {
      var fromP = e.parameter.from;
      var toP = e.parameter.to;
      if (fromP && toP) return jsonResponse(getEntriesForDateRange(String(fromP), String(toP)));
      return jsonResponse(getEntriesForDate(e.parameter.date || ""));
    }
    if (action === "summary") return jsonResponse(getSummaryForDate(e.parameter.date || ""));
    if (action === "dates") return jsonResponse(getAvailableDates());
    if (action === "refreshDemographics") {
      clearDemographicsCache();
      return jsonResponse({ status: "ok", count: Object.keys(getDemographicsMap()).length });
    }
    if (action === "recalcOP") {
      recalcGevsveAndFinalValues();
      return jsonResponse({ status: "ok", message: "O/P recalculated" });
    }
    if (action === "ping") {
      pingWarm();
      return jsonResponse({ status: "ok", message: "pong" });
    }

    return ContentService.createTextOutput("Kerala Survey 2026 API is running.")
      .setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return jsonResponse({ error: String(err) });
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function cellToYmdKolkata(cell) {
  if (cell instanceof Date) {
    if (isNaN(cell.getTime())) return null;
    return Utilities.formatDate(cell, "Asia/Kolkata", "yyyy-MM-dd");
  }

  var s = String(cell);

  var m = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (m) {
