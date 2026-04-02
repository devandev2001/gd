/**
 * Demographic weight lookup tables
 * Source: Cast_Mapping_Split.xlsx - FINAL GENDER CASTE.csv (sync with google-apps-script.js AC_DEMOGRAPHICS)
 *
 * Caste: percentage for that caste in that AC / 100
 * Gender: male/female percentage for that AC / 100
 * Age: "Age Normal" value from Sheet2 (already a decimal fraction)
 */

// ── AC-level caste & gender data ──────────────────────────────
// Key = AC name as it appears in the form dropdown
export const acDemographics = {
  Kattakkada: {
    male: 48.01, female: 51.99,
    Muslim: 6.04, Christian: 22.27, Nair: 34.99, Ezhava: 14.83, Others: 10.07, "SC/ST": 11.37,
  },
  Kovalam: {
    male: 48.88, female: 51.11,
    Muslim: 9.56, Christian: 35.91, Nair: 14.71, Ezhava: 14.35, Others: 13.7, "SC/ST": 11.66,
  },
  Kazhakkoottam: {
    male: 47.82, female: 52.18,
    Muslim: 7.3, Christian: 14.5, Nair: 27.98, Ezhava: 28.3, Others: 11.72, "SC/ST": 10.2,
  },
  Vattiyoorkavu: {
    male: 47.56, female: 52.44,
    Muslim: 6.0, Christian: 18.0, Nair: 35.23, Ezhava: 10.24, Others: 20.33, "SC/ST": 10.2,
  },
  Thiruvananthapuram: {
    male: 48.06, female: 51.93,
    Muslim: 18.0, Christian: 24.0, Nair: 23.21, Ezhava: 8.43, Others: 17.27, "SC/ST": 9.09,
  },
  Nemom: {
    male: 48.3, female: 51.69,
    Muslim: 15.3, Christian: 8.5, Nair: 30.88, Ezhava: 13.25, Others: 22.3, "SC/ST": 9.82,
  },
  Attingal: {
    male: 46.28, female: 53.72,
    Muslim: 17.3, Christian: 1.7, Nair: 19.1, Ezhava: 28.12, Others: 16.21, "SC/ST": 17.47,
  },
  Chathannoor: {
    male: 46.81, female: 53.19,
    Muslim: 12.1, Christian: 10.3, Nair: 26.85, Ezhava: 29.6, Others: 6.25, "SC/ST": 14.7,
  },
  Aranmula: {
    male: 47.95, female: 52.05,
    Muslim: 4.2, Christian: 38.7, Nair: 20.89, Ezhava: 16.37, Others: 4.14, "SC/ST": 15.6,
  },
  Thiruvalla: {
    male: 47.98, female: 52.02,
    Muslim: 2.1, Christian: 48.3, Nair: 16.78, Ezhava: 10.36, Others: 10.41, "SC/ST": 11.85,
  },
  Chengannur: {
    male: 47.64, female: 52.35,
    Muslim: 3.89, Christian: 26.81, Nair: 29.92, Ezhava: 15.54, Others: 7.72, "SC/ST": 15.97,
  },
  Adoor: {
    male: 47.21, female: 52.79,
    Muslim: 6.8, Christian: 26.4, Nair: 25.15, Ezhava: 19.67, Others: 3.15, "SC/ST": 18.84,
  },
  Poonjar: {
    male: 49.51, female: 50.49,
    Muslim: 20.39, Christian: 39.26, Nair: 7.3, Ezhava: 15.11, Others: 6.58, "SC/ST": 11.37,
  },
  Kanjirappally: {
    male: 48.55, female: 51.45,
    Muslim: 10.2, Christian: 40.0, Nair: 23.92, Ezhava: 12.0, Others: 4.21, "SC/ST": 9.66,
  },
  Pala: {
    male: 48.71, female: 51.29,
    Muslim: 1.58, Christian: 56.26, Nair: 16.41, Ezhava: 13.73, Others: 3.61, "SC/ST": 8.39,
  },
  Thrissur: {
    male: 47.55, female: 52.45,
    Muslim: 5.2, Christian: 38.7, Nair: 16.3, Ezhava: 14.0, Others: 17.96, "SC/ST": 7.85,
  },
  Kunnathunad: {
    male: 48.85, female: 51.14,
    Muslim: 19.7, Christian: 35.4, Nair: 11.78, Ezhava: 14.57, Others: 5.42, "SC/ST": 13.13,
  },
  Palakkad: {
    male: 48.63, female: 51.37,
    Muslim: 27.9, Christian: 2.94, Nair: 9.66, Ezhava: 22.08, Others: 25.37, "SC/ST": 11.89,
  },
  "Kozhikode North": {
    male: 47.41, female: 52.59,
    Muslim: 25.1, Christian: 7.9, Nair: 14.07, Ezhava: 32.16, Others: 16.34, "SC/ST": 4.43,
  },
  Kasaragod: {
    male: 50.0, female: 50.0,
    Muslim: 50.42, Christian: 2.4, Nair: 3.3, Ezhava: 15.0, Others: 22.17, "SC/ST": 6.71,
  },
  Manjeshwaram: {
    male: 50.38, female: 49.62,
    Muslim: 52.89, Christian: 2.7, Nair: 0.44, Ezhava: 12.0, Others: 25.6, "SC/ST": 6.37,
  },
  "Nattika (SC)": {
    male: 48.09, female: 51.91,
    Muslim: 16.31, Christian: 14.12, Nair: 9.0, Ezhava: 34.7, Others: 15.26, "SC/ST": 10.45,
  },
  Malampuzha: {
    male: 48.76, female: 51.24,
    Muslim: 11.31, Christian: 6.01, Nair: 10.75, Ezhava: 33.89, Others: 22.76, "SC/ST": 15.26,
  },
  Manalur: {
    male: 48.89, female: 51.11,
    Muslim: 21.05, Christian: 21.52, Nair: 8.6, Ezhava: 29.9, Others: 10.67, "SC/ST": 8.28,
  },
  Perumbavoor: {
    male: 49.14, female: 50.85,
    Muslim: 18.6, Christian: 35.5, Nair: 12.39, Ezhava: 12.39, Others: 11.04, "SC/ST": 10.07,
  },
  Thripunithura: {
    male: 48.32, female: 51.68,
    Muslim: 11.82, Christian: 26.31, Nair: 11.08, Ezhava: 21.55, Others: 20.95, "SC/ST": 7.99,
  },
  /** Caste shares TBD in FINAL GENDER CASTE; gender from Male female CSV — caste % proxied from Thripunithura until split exists. */
  Thrikkakara: {
    male: 48.06, female: 51.94,
    Muslim: 11.82, Christian: 26.31, Nair: 11.08, Ezhava: 21.55, Others: 20.95, "SC/ST": 7.99,
  },
};

const AC_ALIASES = {
  kattakada: "Kattakkada",
  kowalam: "Kovalam",
  neyyattinkara: "Nemom",
  naimam: "Nemom",
  naiyamam: "Nemom",
  nemeom: "Nemom",
  nemam: "Nemom",
  nattika: "Nattika (SC)",
  nattikasc: "Nattika (SC)",
  thrissurac: "Thrissur",
  manalurac: "Manalur",
  perumbaavoor: "Perumbavoor",
  perumbavoor: "Perumbavoor",
  thripunitura: "Thripunithura",
  thrippunithura: "Thripunithura",
  thrikakkara: "Thrikkakara",
};

const CANONICAL_AC_BY_KEY = Object.keys(acDemographics).reduce((acc, ac) => {
  acc[String(ac).toLowerCase().replace(/[^a-z0-9]/g, "")] = ac;
  return acc;
}, {});

function normalizeAcName(ac) {
  const raw = String(ac || "").trim();
  if (!raw) return "";
  const key = raw.toLowerCase().replace(/[^a-z0-9]/g, "");
  return AC_ALIASES[key] || CANONICAL_AC_BY_KEY[key] || raw;
}

// ── Age Normal weights (statewide, from Sheet2) ───────────────
// Key = age group label as shown in the form
// Value = "Age Normal" directly from the CSV (already a decimal)
export const ageNormalWeights = {
  "18-19": 0.01574992977,
  "20-29": 0.16726013,
  "30-39": 0.1839168018,
  "40-49": 0.2081028821,
  "50-59": 0.19009771,
  "60-69": 0.1395625022,
  "70-79": 0.07469513213,
  "80+":   0.02061491203,
};

/**
 * Look up weights for a survey entry.
 *
 * @param {string} ac       - Assembly constituency name (from form)
 * @param {string} caste    - Caste selection (Nair / Ezhava / Muslim / Christian / SC/ST / Others)
 * @param {string} gender   - "Male" or "Female"
 * @param {string} ageGroup - e.g. "30-39"
 *
 * @returns {{ casteWeight: number, genderWeight: number, ageWeight: number }}
 *          Caste & Gender = percentage / 100. Age = Age Normal value as-is.
 *          Returns 0 if lookup fails.
 */
export function getWeights(ac, caste, gender, ageGroup) {
  const acKey = normalizeAcName(ac);
  const acData = acDemographics[acKey];

  // Caste weight: that caste's % in this AC / 100
  const casteWeight = acData ? (acData[caste] ?? 0) / 100 : 0;

  // Gender weight: male/female % in this AC / 100 (case-insensitive)
  const g = String(gender || "").trim().toLowerCase();
  const genderKey = g === "male" ? "male" : "female";
  const genderWeight = acData ? (acData[genderKey] ?? 0) / 100 : 0;

  // Age weight: Age Normal value from Sheet2 (already a decimal)
  const ageWeight = ageNormalWeights[ageGroup] ?? 0;

  return { casteWeight, genderWeight, ageWeight };
}
