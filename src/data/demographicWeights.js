/**
 * Demographic weight lookup tables
 * Source: Cast_Mapping_Split.xlsx
 *
 * Values are percentages from the CSV.
 * The form divides by 100 before sending to Google Sheets.
 *
 * Caste columns in CSV: Nair, Ezhava, Muslim, Christain (sic), SC/ST, Others
 * Gender columns: Male Percentage, Female Percentage
 * Age-Gender columns: from Sheet2 — Male %, Female % per age group
 */

// ── AC-level caste & gender data ──────────────────────────────
// Key = AC name as it appears in the form dropdown
export const acDemographics = {
  Kattakkada: {
    male: 48.01, female: 51.99,
    Nair: 34.99, Ezhava: 14.83, Muslim: 6.04, Christian: 11.03, "SC/ST": 11.03, Others: 9.50,
  },
  Kovalam: {
    male: 48.88, female: 51.11,
    Nair: 14.71, Ezhava: 14.35, Muslim: 9.56, Christian: 11.49, "SC/ST": 11.49, Others: 12.59,
  },
  Vattiyoorkavu: {
    male: 47.56, female: 52.44,
    Nair: 35.23, Ezhava: 10.24, Muslim: 6.00, Christian: 9.75, "SC/ST": 9.75, Others: 12.74,
  },
  Thiruvananthapuram: {
    male: 48.06, female: 51.93,
    Nair: 23.21, Ezhava: 8.43, Muslim: 18.00, Christian: 8.65, "SC/ST": 8.65, Others: 10.41,
  },
  Attingal: {
    male: 46.28, female: 53.72,
    Nair: 19.10, Ezhava: 28.12, Muslim: 17.30, Christian: 17.21, "SC/ST": 17.21, Others: 13.96,
  },
  Chathannoor: {
    male: 46.81, female: 53.19,
    Nair: 26.85, Ezhava: 29.60, Muslim: 12.10, Christian: 14.36, "SC/ST": 14.36, Others: 4.29,
  },
  Aranmula: {
    male: 47.95, female: 52.05,
    Nair: 20.89, Ezhava: 20.89, Muslim: 4.20, Christian: 15.30, "SC/ST": 15.30, Others: 0.86,
  },
  Thiruvalla: {
    male: 47.98, female: 52.02,
    Nair: 16.78, Ezhava: 10.36, Muslim: 0.00, Christian: 11.46, "SC/ST": 11.46, Others: 10.03,
  },
  Chengannur: {
    male: 47.64, female: 52.35,
    Nair: 29.92, Ezhava: 15.54, Muslim: 3.89, Christian: 15.75, "SC/ST": 15.75, Others: 5.71,
  },
  Adoor: {
    male: 47.21, female: 52.79,
    Nair: 25.15, Ezhava: 19.69, Muslim: 6.80, Christian: 18.60, "SC/ST": 18.60, Others: 1.84,
  },
  Poonjar: {
    male: 49.51, female: 50.49,
    Nair: 7.30, Ezhava: 15.11, Muslim: 20.39, Christian: 11.37, "SC/ST": 11.37, Others: 2.45,
  },
  Pala: {
    male: 48.71, female: 51.29,
    Nair: 16.41, Ezhava: 13.73, Muslim: 1.58, Christian: 8.39, "SC/ST": 8.39, Others: 2.20,
  },
  Thrissur: {
    male: 47.55, female: 52.45,
    Nair: 18.03, Ezhava: 17.11, Muslim: 5.20, Christian: 7.66, "SC/ST": 7.66, Others: 8.61,
  },
  Kunnathunad: {
    male: 48.85, female: 51.14,
    Nair: 11.78, Ezhava: 14.57, Muslim: 19.70, Christian: 12.69, "SC/ST": 12.69, Others: 4.36,
  },
  Palakkad: {
    male: 48.63, female: 51.37,
    Nair: 9.66, Ezhava: 22.08, Muslim: 27.84, Christian: 11.75, "SC/ST": 11.75, Others: 9.90,
  },
  "Kozhikode North": {
    male: 47.41, female: 52.59,
    Nair: 14.07, Ezhava: 33.16, Muslim: 25.10, Christian: 0.00, "SC/ST": 0.00, Others: 11.55,
  },
  Kasaragod: {
    male: 50.00, female: 50.00,
    Nair: 3.30, Ezhava: 15.00, Muslim: 50.42, Christian: 5.43, "SC/ST": 5.43, Others: 15.20,
  },
  Manjeshwaram: {
    male: 50.38, female: 49.62,
    Nair: 0.00, Ezhava: 12.00, Muslim: 52.89, Christian: 4.77, "SC/ST": 4.77, Others: 20.01,
  },
};

// ── Age-group gender weights (statewide) ──────────────────────
// Key = age group label as shown in the form
export const ageGenderWeights = {
  "18-19": { male: 50.31, female: 49.69 },
  "20-29": { male: 50.06, female: 49.94 },
  "30-39": { male: 51.57, female: 48.43 },
  "40-49": { male: 49.49, female: 50.50 },
  "50-59": { male: 46.26, female: 53.74 },
  "60-69": { male: 47.57, female: 52.43 },
  "70-79": { male: 46.71, female: 53.29 },
  "80+":   { male: 40.12, female: 59.88 },
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
 *          All values are decimal fractions (percentage / 100).
 *          Returns 0 if lookup fails.
 */
export function getWeights(ac, caste, gender, ageGroup) {
  const acData = acDemographics[ac];
  const ageData = ageGenderWeights[ageGroup];

  // Caste weight: that caste's % in this AC / 100
  const casteWeight = acData ? (acData[caste] ?? 0) / 100 : 0;

  // Gender weight: male/female % in this AC / 100
  const genderKey = gender === "Male" ? "male" : "female";
  const genderWeight = acData ? (acData[genderKey] ?? 0) / 100 : 0;

  // Age weight: male/female % for this age group / 100
  const ageWeight = ageData ? (ageData[genderKey] ?? 0) / 100 : 0;

  return { casteWeight, genderWeight, ageWeight };
}
