/**
 * Kerala Survey 2026 — Main sheet A:P (deploy when columns include labels + GevsVE + Final Values)
 *
 * A Timestamp
 * B Assembly Constituency
 * C FA Name
 * D Caste Weight          E Caste Label
 * F Gender Weight         G Gender Label
 * H Age Weight            I Age Label
 * J Vote 2021 AE          K Vote 2024 GE          L Vote 2026 AE
 * M Who Will Win          N Who Will Win Normalized (D×F×H)
 * O GevsVE                P Final Values (O×N)
 *
 * IMPORTANT: Dashboard/report weighting uses Final Values (P), not N.
 *
 * GevsVE (O) = same as your sheet: GEVSVE_BASE[party][B] / COUNTIFS($K:$K,party,$B:$B,B) for UDF/LDF/BJP/NDA; else 1.
 * Kasaragod spelling variants (e.g. Kasargod) normalize to “Kasaragod” so denominators are not split.
 * doPost does not run that full recalc inline — it schedules it (~2s) so the form returns fast on large Sheet1.
 * Sheet filters hide rows but do not remove data; clear filters on K if rows seem “missing”.
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

// GevsVE base values (sheet formula constants) — keep in sync with Excel
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

/**
 * incrementalUpdateGevsveAfterAppend reads B+K+N for every row — OK for small sheets; on ~100k+ rows it can take 30s+.
 * Above this row count, skip incremental and rely on deferred full recalc only (fast doPost; O/P in ~2s).
 */
var GEVSVE_INCREMENTAL_MAX_DATA_ROWS = 12000;

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
  var s = String(acRaw || "").replace(/[\u200B-\u200D\uFEFF\u00A0]/g, " ").trim().replace(/\s+/g, " ");
  var k = s.toLowerCase();
  if (k === "kattakada") return "Kattakkada";
  if (k === "kowalam") return "Kovalam";
  if (k === "trivandrum") return "Thiruvananthapuram";
  if (k === "naimam" || k === "nemam" || k === "nemeom" || k === "naiyamam") return "Nemom";
  // GEVSVE_BASE / demographics / Excel use spelling "Kasaragod" — merge common variants so COUNTIFS denominator is not split
  if (k === "kasargod" || k === "kasaragodu" || k === "kasargodu" || k === "kasrgod" || k === "kassaragod") return "Kasaragod";
  if (k === "manjeswaram" || k === "manjeshwar") return "Manjeshwaram";

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

/** Column K: store UDF / LDF / BJP/NDA exactly like Excel COUNTIFS; keep “Not Voted” / “Others” as submitted. */
function vote2024ForSheet(v) {
  var p = normalizeParty(v);
  if (p === "UDF" || p === "LDF" || p === "BJP/NDA") return p;
  return String(v === null || v === undefined ? "" : v).trim();
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
