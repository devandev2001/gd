/**
 * Kerala Survey 2026 — Main sheet A:P (deploy this version when columns include labels + GevsVE + Final Values)
 *
 * A Timestamp
 * B Assembly Constituency
 * C FA Name
 * D Caste Weight          E Caste Label
 * F Gender Weight         G Gender Label
 * H Age Weight            I Age Label
 * J Vote 2021 AE          K Vote 2024 GE          L Vote 2026 AE
 * M Who Will Win          N Who Will Win Normalized (D×F×H)
 * O GevsVE                P Final Values (O×N) — dashboard uses P as normalizedScore
 *
 * Form values: whatever the user selects (caste / gender / age) is written to E, G, I when valid;
 * weights in D, F, H always come from demographics + that selection.
 *
 * doPost returns immediately after the new row is written; O/P are recalculated ~2s later via trigger so the
 * form does not wait on full-sheet or column-scans (avoids 30–40s spinners on large Sheet1). Column P may lag
 * a few seconds vs N. After bulk edits, open ?action=recalcOP. Optional: time-driven trigger on pingWarm() every
 * 5–10 min reduces Apps Script cold starts.
 */

var MAIN_SHEET_NAME = "Sheet1";
var DEMOGRAPHICS_SHEET_NAME = "FINAL GENDER CASTE";
var USE_DEMOGRAPHICS_SHEET = true;

var AGE_WEIGHTS = {
  "18-19": 0.01574992977,
  "20-29": 0.16726013,
  "30-39": 0.1839168018,
  "40-49": 0.2081028821,
  "50-59": 0.19009771,
  "60-69": 0.1395625022,
  "70-79": 0.07469513213,
  "80+":   0.02061491203
};

var CASTE_LIST = ["Nair", "Ezhava", "Muslim", "Christian", "SC/ST", "Others"];

var AC_DEMOGRAPHICS_FALLBACK = {
  Kattakkada:         { male:48.01, female:51.99, Muslim:6.04,  Christian:22.27, Nair:34.99, Ezhava:14.83, Others:10.07, "SC/ST":11.37 },
  Kovalam:            { male:48.88, female:51.11, Muslim:9.56,  Christian:35.91, Nair:14.71, Ezhava:14.35, Others:13.70, "SC/ST":11.66 },
  Kazhakkoottam:      { male:47.82, female:52.18, Muslim:7.30,  Christian:14.50, Nair:27.98, Ezhava:28.30, Others:11.72, "SC/ST":10.20 },
  Vattiyoorkavu:      { male:47.56, female:52.44, Muslim:6.00,  Christian:18.00, Nair:35.23, Ezhava:10.24, Others:20.33, "SC/ST":10.20 },
  Thiruvananthapuram: { male:48.06, female:51.93, Muslim:18.00, Christian:24.00, Nair:23.21, Ezhava:8.43,  Others:17.27, "SC/ST":9.09  },
  Nemom:              { male:48.30, female:51.69, Muslim:15.30, Christian:8.50,  Nair:30.88, Ezhava:13.25, Others:22.30, "SC/ST":9.82  },
  Attingal:           { male:46.28, female:53.72, Muslim:17.30, Christian:1.70,  Nair:19.10, Ezhava:28.12, Others:16.21, "SC/ST":17.47 },
  Chathannoor:        { male:46.81, female:53.19, Muslim:12.10, Christian:10.30, Nair:26.85, Ezhava:29.60, Others:6.25,  "SC/ST":14.70 },
  Aranmula:           { male:47.95, female:52.05, Muslim:4.20,  Christian:38.70, Nair:20.89, Ezhava:16.37, Others:4.14,  "SC/ST":15.60 },
  Thiruvalla:         { male:47.98, female:52.02, Muslim:2.10,  Christian:48.30, Nair:16.78, Ezhava:10.36, Others:10.41, "SC/ST":11.85 },
  Chengannur:         { male:47.64, female:52.35, Muslim:3.89,  Christian:26.81, Nair:29.92, Ezhava:15.54, Others:7.72,  "SC/ST":15.97 },
  Adoor:              { male:47.21, female:52.79, Muslim:6.80,  Christian:26.40, Nair:25.15, Ezhava:19.67, Others:3.15,  "SC/ST":18.84 },
  Poonjar:            { male:49.51, female:50.49, Muslim:20.39, Christian:39.26, Nair:7.30,  Ezhava:15.11, Others:6.58,  "SC/ST":11.37 },
  Kanjirappally:      { male:48.55, female:51.45, Muslim:10.20, Christian:40.00, Nair:23.92, Ezhava:12.00, Others:4.21,  "SC/ST":9.66  },
  Pala:               { male:48.71, female:51.29, Muslim:1.58,  Christian:56.26, Nair:16.41, Ezhava:13.73, Others:3.61,  "SC/ST":8.39  },
  Thrissur:           { male:47.55, female:52.45, Muslim:5.20,  Christian:38.70, Nair:16.30, Ezhava:14.00, Others:17.96, "SC/ST":7.85  },
  Kunnathunad:        { male:48.85, female:51.14, Muslim:19.70, Christian:35.40, Nair:11.78, Ezhava:14.57, Others:5.42,  "SC/ST":13.13 },
  Palakkad:           { male:48.63, female:51.37, Muslim:27.90, Christian:2.94,  Nair:9.66,  Ezhava:22.08, Others:25.37, "SC/ST":11.89 },
  "Kozhikode North":  { male:47.41, female:52.59, Muslim:25.10, Christian:7.90,  Nair:14.07, Ezhava:32.16, Others:16.34, "SC/ST":4.43  },
  Kasaragod:          { male:50.00, female:50.00, Muslim:50.42, Christian:2.40,  Nair:3.30,  Ezhava:15.00, Others:22.17, "SC/ST":6.71  },
  Manjeshwaram:       { male:50.38, female:49.62, Muslim:52.89, Christian:2.70,  Nair:0.44,  Ezhava:12.00, Others:25.60, "SC/ST":6.37  }
};

var GEVSVE_BASE = {
  "UDF": {
    "Chathannoor": 24.93, "Attingal": 25.02, "Kattakkada": 29.55,
    "Manjeshwaram": 38.14, "Kasaragod": 49.57, "Poonjar": 24.76, "Nemom": 28.85
  },
  "LDF": {
    "Chathannoor": 43.12, "Attingal": 47.35, "Kattakkada": 45.49,
    "Manjeshwaram": 23.57, "Kasaragod": 17.67, "Poonjar": 41.94, "Nemom": 24.59
  },
  "BJP/NDA": {
    "Chathannoor": 30.61, "Attingal": 25.92, "Kattakkada": 23.77,
    "Manjeshwaram": 37.7, "Kasaragod": 31.76, "Poonjar": 29.92, "Nemom": 45.18
  }
};

var _DEMOGRAPHICS_CACHE = null;

function pad2(n) {
  n = parseInt(n, 10);
  return n < 10 ? "0" + n : String(n);
}

function asFiniteNumber(val, fallback) {
  if (val === null || val === undefined || val === "") return fallback;
  var n = typeof val === "number" ? val : parseFloat(String(val).replace(/,/g, "").replace(/%/g, ""));
  return isFinite(n) ? n : fallback;
}

function isTextLabel(val) {
  var s = String(val === null || val === undefined ? "" : val).trim();
  if (s === "") return false;
  return isNaN(parseFloat(s));
}

function normalizeAgeLabel(label) {
  return String(label || "").trim().replace(/[–—]/g, "-");
}

/** Map caste answer to CASTE_LIST key (case / SC-ST variants). */
function canonicalCasteKey(casteStr) {
  var s = String(casteStr || "").trim();
  if (!s) return s;
  var low = s.toLowerCase().replace(/\s+/g, " ");
  if (low === "sc/st" || low === "scst" || low === "sc-st" || low === "sc / st") return "SC/ST";
  var i;
  for (i = 0; i < CASTE_LIST.length; i++) {
    if (CASTE_LIST[i].toLowerCase() === low) return CASTE_LIST[i];
  }
  return s;
}

function isValidCasteLabel(val) {
  var k = canonicalCasteKey(val);
  return CASTE_LIST.indexOf(k) >= 0;
}

function isValidGenderLabel(val) {
  var g = String(val || "").trim().toLowerCase();
  return g === "male" || g === "female";
}

function displayGenderLabel(val) {
  var g = String(val || "").trim().toLowerCase();
  return g === "male" ? "Male" : (g === "female" ? "Female" : "");
}

function isValidAgeLabel(val) {
  var k = normalizeAgeLabel(val);
  return AGE_WEIGHTS.hasOwnProperty(k);
}

function normalizeAcName(acRaw) {
  var s = String(acRaw || "").trim();
  var k = s.toLowerCase().replace(/\s+/g, " ");
  if (k === "kattakada") return "Kattakkada";
  if (k === "kowalam") return "Kovalam";
  if (k === "trivandrum") return "Thiruvananthapuram";
  if (k === "naimam" || k === "nemam" || k === "nemeom" || k === "naiyamam" || k === "neyyattinkara") return "Nemom";

  var dem = getDemographicsMap();
  if (dem[s]) return s;

  var keys = Object.keys(dem);
  for (var i = 0; i < keys.length; i++) {
    if (keys[i].toLowerCase() === k) return keys[i];
  }
  return s;
}

function normalizeParty(p) {
  var s = String(p || "").trim().toUpperCase().replace(/\s+/g, "");
  if (!s) return "Others";
  if (s === "LDF") return "LDF";
  if (s === "UDF") return "UDF";
  if (s === "BJP/NDA" || s === "BJP-NDA" || s === "BJPNDA" || s === "BJP" || s === "NDA") return "BJP/NDA";
  return "Others";
}

function ensureHeaders(sheet) {
  var a1 = String(sheet.getRange(1, 1).getValue() || "").trim();
  if (a1 === "Timestamp") return;

  var headers = [[
    "Timestamp","Assembly Constituency","FA Name",
    "Caste Weight","Caste Label",
    "Gender Weight","Gender Label",
    "Age Weight","Age Label",
    "Vote in 2021 AE","Vote in 2024 GE","Vote in 2026 AE",
    "Who Will Win","Who Will Win Normalized",
    "GevsVE","Final Values"
  ]];
  sheet.getRange(1, 1, 1, 16).setValues(headers);
}

function parsePct(val) {
  return asFiniteNumber(val, 0);
}

function readDemographicsFromSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(DEMOGRAPHICS_SHEET_NAME);
  if (!sh) return null;

  var lr = sh.getLastRow();
  var lc = sh.getLastColumn();
  if (lr < 2 || lc < 9) return null;

  var data = sh.getRange(1, 1, lr, lc).getValues();
  var head = data[0].map(function(h) { return String(h).trim(); });

  function idx(name) { return head.indexOf(name); }

  var iAc = idx("AC Name");
  var iMale = idx("Male Percentage");
  var iFemale = idx("Female Percentage");
  var iMuslim = idx("%Muslims");
  var iChristian = idx("%Christians");
  var iNair = idx("%Nairs");
  var iEzhava = idx("%Ezhavas");
  var iOthers = idx("%Others");
  var iScst = idx("%SC/ST");

  if ([iAc, iMale, iFemale, iMuslim, iChristian, iNair, iEzhava, iOthers, iScst].some(function(x){ return x < 0; })) {
    throw new Error("Demographics sheet headers missing/changed. Keep exact column names.");
  }

  var map = {};
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var ac = normalizeAcName(row[iAc]);
    if (!ac) continue;
    map[ac] = {
      male: parsePct(row[iMale]),
      female: parsePct(row[iFemale]),
      Muslim: parsePct(row[iMuslim]),
      Christian: parsePct(row[iChristian]),
      Nair: parsePct(row[iNair]),
      Ezhava: parsePct(row[iEzhava]),
      Others: parsePct(row[iOthers]),
      "SC/ST": parsePct(row[iScst])
    };
  }
  return map;
}

function getDemographicsMap() {
  if (_DEMOGRAPHICS_CACHE) return _DEMOGRAPHICS_CACHE;

  if (USE_DEMOGRAPHICS_SHEET) {
    try {
      var fromSheet = readDemographicsFromSheet();
      if (fromSheet && Object.keys(fromSheet).length > 0) {
        _DEMOGRAPHICS_CACHE = fromSheet;
        return _DEMOGRAPHICS_CACHE;
      }
    } catch (err) {
      Logger.log("readDemographicsFromSheet failed: " + err);
    }
  }

  _DEMOGRAPHICS_CACHE = AC_DEMOGRAPHICS_FALLBACK;
  return _DEMOGRAPHICS_CACHE;
}

function clearDemographicsCache() {
  _DEMOGRAPHICS_CACHE = null;
}

function getClosestAgeLabelFromNormalizedWeight(w) {
  var best = "Unknown";
  var bestDiff = Infinity;
  for (var label in AGE_WEIGHTS) {
    if (!AGE_WEIGHTS.hasOwnProperty(label)) continue;
    var diff = Math.abs(AGE_WEIGHTS[label] - w);
    if (diff < bestDiff) { bestDiff = diff; best = label; }
  }
  return best;
}

function getAgeLabelFromAny(ageVal) {
  if (isTextLabel(ageVal)) {
    var txt = normalizeAgeLabel(ageVal);
    if (AGE_WEIGHTS.hasOwnProperty(txt)) return txt;
    return "Unknown";
  }

  var n = asFiniteNumber(ageVal, NaN);
  if (!isFinite(n)) return "Unknown";

  if (n > 0 && n < 1.2) return getClosestAgeLabelFromNormalizedWeight(n);
  if (n >= 80) return "80+";
  if (n >= 70) return "70-79";
  if (n >= 60) return "60-69";
  if (n >= 50) return "50-59";
  if (n >= 40) return "40-49";
  if (n >= 30) return "30-39";
  if (n >= 20) return "20-29";
  if (n >= 18) return "18-19";
  return "Unknown";
}

function getAgeWeightFromAny(ageVal) {
  var label = getAgeLabelFromAny(ageVal);
  return AGE_WEIGHTS[label] || 0;
}

function resolveWeights(ac, casteVal, genderVal, ageVal) {
  var acName = normalizeAcName(ac);
  var dem = getDemographicsMap();
  var acData = dem[acName];

  var casteW, genderW, ageW;

  if (isTextLabel(casteVal)) {
    var ck = canonicalCasteKey(casteVal);
    casteW = acData ? (acData[ck] || 0) / 100 : 0;
  } else {
    casteW = asFiniteNumber(casteVal, 0);
    if (casteW > 1.5) casteW = casteW / 100;
  }

  if (isTextLabel(genderVal)) {
    var g = String(genderVal).trim().toLowerCase();
    genderW = acData ? (g === "male" ? acData.male : acData.female) / 100 : 0;
  } else {
    genderW = asFiniteNumber(genderVal, 0);
    if (genderW > 1.5) genderW = genderW / 100;
  }

  ageW = getAgeWeightFromAny(ageVal);
  var norm = casteW * genderW * ageW;

  return { casteW: casteW, genderW: genderW, ageW: ageW, norm: norm };
}

/**
 * Prefer the form’s caste string when it matches a known label; otherwise infer from weight.
 */
function casteLabelForSheet(ac, formCaste, casteW) {
  if (formCaste != null && isValidCasteLabel(formCaste)) return canonicalCasteKey(formCaste);
  return getCasteLabelFromWeight(ac, casteW);
}

/**
 * Prefer the form’s gender when Male/Female; avoids 50/50 AC tie bugs (e.g. Kasaragod).
 */
function genderLabelForSheet(ac, formGender, genderW) {
  if (formGender != null && isValidGenderLabel(formGender)) {
    var d = displayGenderLabel(formGender);
    if (d) return d;
  }
  return getGenderLabelFromWeight(ac, genderW);
}

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
  return femaleW >= maleW ? "Female" : "Male";
}

function getAgeLabelFromWeight(ageWeight) {
  return getAgeLabelFromAny(ageWeight);
}

function getBaseGevsveValue(party, ac) {
  if (GEVSVE_BASE[party] && GEVSVE_BASE[party].hasOwnProperty(ac)) return GEVSVE_BASE[party][ac];
  return 1;
}

/** Count rows by AC + 2024 vote (column K = index 10) for GevsVE denominator. */
function buildCountifsMapForK(dataRows) {
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
  if (base === 1) return 1;

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
  var oVal = (base === 1 || denom <= 0) ? 1 : base / denom;

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
 * If trigger creation fails (quota / permissions), falls back to synchronous incremental update.
 */
function scheduleDeferredGevsveRecalc(sheet, lastRow) {
  var props = PropertiesService.getScriptProperties();
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) {
    incrementalUpdateGevsveAfterAppend(sheet, lastRow);
    return;
  }
  try {
    if (props.getProperty(PROP_GEVSVE_RECALC_PENDING) === "1") return;

    ScriptApp.newTrigger("runDeferredGevsveRecalc")
      .timeBased()
      .after(2000)
      .create();

    props.setProperty(PROP_GEVSVE_RECALC_PENDING, "1");
  } catch (err) {
    Logger.log("scheduleDeferredGevsveRecalc failed, using sync incremental: " + err);
    incrementalUpdateGevsveAfterAppend(sheet, lastRow);
  } finally {
    lock.releaseLock();
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
  var countMap = buildCountifsMapForK(data);

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
      data.vote2024,
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

function legacySubstringRowMatch(row, dateStr) {
  if (!dateStr) return true;
  var ts = String(row[0]);
  var patterns = timestampPatternsForDateKey(dateStr);
  for (var i = 0; i < patterns.length; i++) if (ts.indexOf(patterns[i]) !== -1) return true;
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

  return filtered.map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i]; });
    obj.normalizedScore = row[15];
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
  var filtered = dateStr ? data.filter(function(row) { return rowMatchesDateFilter(row, dateStr); }) : data;
  var entries = mapRowsToEntries(filtered);
  return { entries: entries, total: entries.length };
}

function getEntriesForDateRange(fromYmd, toYmd) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME) || ss.getSheets()[0];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { entries: [], total: 0 };

  var numRows = lastRow - 1;
  var data = sheet.getRange(2, 1, numRows, 16).getValues();
  var filtered = data.filter(function(row) {
    var rowYmd = cellToYmdKolkata(row[0]);
    if (rowYmd) return rowYmd >= fromYmd && rowYmd <= toYmd;
    if (fromYmd === toYmd) return rowMatchesDateFilter(row, fromYmd);
    return false;
  });

  var entries = mapRowsToEntries(filtered);
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
