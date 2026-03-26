/**
 * Demographic weight lookup tables
 * Source: Cast_Mapping_Split.xlsx - Data Caste (1).csv (sync with google-apps-script AC_DEMOGRAPHICS)
 *
 * Caste: percentage for that caste in that AC / 100
 * Gender: male/female percentage for that AC / 100
 * Age: "Age Normal" value from Sheet2 (already a decimal fraction)
 *
 * Aranmula row: CSV had SC/ST and Others negative; stored values use max(0, %) for weights.
 */

// ── AC-level caste & gender data ──────────────────────────────
// Key = AC name as it appears in the form dropdown
export const acDemographics = {
  Kattakkada: {
    male: 48.01, female: 51.99,
    Nair: 34.99, Ezhava: 14.83, Muslim: 6.04, Christian: 22.27, "SC/ST": 11.37, Others: 10.5,
  },
  Kovalam: {
    male: 48.88, female: 51.11,
    Nair: 14.71, Ezhava: 14.35, Muslim: 9.56, Christian: 35.91, "SC/ST": 11.66, Others: 13.81,
  },
  Vattiyoorkavu: {
    male: 47.56, female: 52.44,
    Nair: 35.23, Ezhava: 10.24, Muslim: 6.0, Christian: 18.0, "SC/ST": 10.2, Others: 20.33,
  },
  Thiruvananthapuram: {
    male: 48.06, female: 51.93,
    Nair: 23.21, Ezhava: 8.43, Muslim: 18.0, Christian: 24.0, "SC/ST": 9.09, Others: 17.27,
  },
  Nemom: {
    male: 48.3, female: 51.69,
    Nair: 30.88, Ezhava: 13.25, Muslim: 15.3, Christian: 8.5, "SC/ST": 9.41, Others: 22.66,
  },
  Attingal: {
    male: 46.28, female: 53.72,
    Nair: 19.1, Ezhava: 28.12, Muslim: 17.3, Christian: 1.7, "SC/ST": 17.47, Others: 16.31,
  },
  Chathannoor: {
    male: 46.81, female: 53.19,
    Nair: 26.85, Ezhava: 29.6, Muslim: 12.1, Christian: 10.3, "SC/ST": 14.7, Others: 6.45,
  },
  Aranmula: {
    male: 47.95, female: 52.05,
    Nair: 20.89, Ezhava: 20.89, Muslim: 4.2, Christian: 38.7, "SC/ST": 15.6, Others: 0,
  },
  Thiruvalla: {
    male: 47.98, female: 52.02,
    Nair: 16.78, Ezhava: 10.36, Muslim: 2.1, Christian: 48.3, "SC/ST": 11.93, Others: 10.53,
  },
  Chengannur: {
    male: 47.64, female: 52.35,
    Nair: 29.92, Ezhava: 15.54, Muslim: 3.89, Christian: 26.81, "SC/ST": 15.97, Others: 7.87,
  },
  Adoor: {
    male: 47.21, female: 52.79,
    Nair: 25.15, Ezhava: 19.69, Muslim: 6.8, Christian: 26.4, "SC/ST": 18.84, Others: 3.12,
  },
  Poonjar: {
    male: 49.51, female: 50.49,
    Nair: 7.3, Ezhava: 15.11, Muslim: 20.39, Christian: 39.26, "SC/ST": 11.37, Others: 6.57,
  },
  Kanjirappally: {
    male: 48.55, female: 51.45,
    Nair: 23.92, Ezhava: 12.0, Muslim: 10.2, Christian: 40.0, "SC/ST": 9.66, Others: 4.22,
  },
  Pala: {
    male: 48.71, female: 51.29,
    Nair: 16.41, Ezhava: 13.73, Muslim: 1.58, Christian: 56.26, "SC/ST": 8.39, Others: 3.63,
  },
  Thrissur: {
    male: 47.55, female: 52.45,
    Nair: 18.03, Ezhava: 17.11, Muslim: 5.2, Christian: 38.7, "SC/ST": 7.85, Others: 13.11,
  },
  Kunnathunad: {
    male: 48.85, female: 51.14,
    Nair: 11.78, Ezhava: 14.57, Muslim: 19.7, Christian: 35.4, "SC/ST": 13.13, Others: 5.42,
  },
  Palakkad: {
    male: 48.63, female: 51.37,
    Nair: 9.66, Ezhava: 22.08, Muslim: 27.84, Christian: 2.94, "SC/ST": 11.88, Others: 25.6,
  },
  "Kozhikode North": {
    male: 47.41, female: 52.59,
    Nair: 14.07, Ezhava: 33.16, Muslim: 25.1, Christian: 7.9, "SC/ST": 4.42, Others: 15.35,
  },
  Kasaragod: {
    male: 50.0, female: 50.0,
    Nair: 3.3, Ezhava: 15.0, Muslim: 50.42, Christian: 2.4, "SC/ST": 6.7, Others: 22.18,
  },
  Manjeshwaram: {
    male: 50.38, female: 49.62,
    Nair: 0.44, Ezhava: 12.0, Muslim: 52.89, Christian: 2.7, "SC/ST": 6.36, Others: 25.61,
  },
};

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
  const acData = acDemographics[ac];

  // Caste weight: that caste's % in this AC / 100
  const casteWeight = acData ? (acData[caste] ?? 0) / 100 : 0;

  // Gender weight: male/female % in this AC / 100
  const genderKey = gender === "Male" ? "male" : "female";
  const genderWeight = acData ? (acData[genderKey] ?? 0) / 100 : 0;

  // Age weight: Age Normal value from Sheet2 (already a decimal)
  const ageWeight = ageNormalWeights[ageGroup] ?? 0;

  return { casteWeight, genderWeight, ageWeight };
}
