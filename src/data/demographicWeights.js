/**
 * Demographic weight lookup tables
 * Source: Cast_Mapping_Split.xlsx + FINAL GENDER CASTE sheet
 *
 * Caste: percentage for that caste in that AC / 100
 * Gender: male/female percentage for that AC / 100
 * Age: "Age Normal" value from Sheet2 (already a decimal fraction)
 *
 * NOTE: The Apps Script server-side (AC_DEMOGRAPHICS_FALLBACK + the FINAL GENDER CASTE sheet)
 * is the authoritative source. This client-side copy is used only for the initial
 * normalizedScore sent with each submission (D×F×H). The server then recomputes
 * raking weights for the entire AC on every append.
 */

// ── AC-level caste & gender data ──────────────────────────────
// Key = AC name as it appears in the form dropdown
export const acDemographics = {
  Adoor:              { male: 47.21, female: 52.79, Nair: 25.15, Ezhava: 19.67, Muslim: 6.80,  Christian: 26.40, "SC/ST": 18.84, Others: 3.15  },
  Ambalapuzha:        { male: 48.78, female: 51.21, Nair: 13.69, Ezhava: 27.72, Muslim: 24.90, Christian: 15.60, "SC/ST": 3.76,  Others: 14.33 },
  Aranmula:           { male: 47.95, female: 52.05, Nair: 20.89, Ezhava: 16.37, Muslim: 4.20,  Christian: 38.70, "SC/ST": 15.60, Others: 4.14  },
  Aroor:              { male: 49.05, female: 50.95, Nair: 13.24, Ezhava: 37.20, Muslim: 10.80, Christian: 19.40, "SC/ST": 12.61, Others: 6.65  },
  Attingal:           { male: 46.28, female: 53.72, Nair: 19.10, Ezhava: 28.12, Muslim: 17.30, Christian: 1.70,  "SC/ST": 17.47, Others: 16.21 },
  Beypur:             { male: 48.97, female: 51.03, Nair: 10.00, Ezhava: 15.00, Muslim: 40.00, Christian: 5.00,  "SC/ST": 10.00, Others: 20.00 },
  Chalakkudy:         { male: 48.87, female: 51.13, Nair: 7.70,  Ezhava: 17.80, Muslim: 2.91,  Christian: 49.09, "SC/ST": 11.58, Others: 10.97 },
  Changanassery:      { male: 48.20, female: 51.80, Nair: 12.57, Ezhava: 17.34, Muslim: 7.90,  Christian: 45.40, "SC/ST": 10.35, Others: 6.44  },
  Chathannoor:        { male: 46.81, female: 53.19, Nair: 26.85, Ezhava: 29.60, Muslim: 12.10, Christian: 10.30, "SC/ST": 14.70, Others: 6.25  },
  Chelakkara:         { male: 48.42, female: 51.58, Nair: 20.00, Ezhava: 12.00, Muslim: 26.93, Christian: 6.51,  "SC/ST": 15.90, Others: 18.71 },
  Chengannur:         { male: 47.64, female: 52.35, Nair: 29.92, Ezhava: 15.54, Muslim: 3.89,  Christian: 26.81, "SC/ST": 15.97, Others: 7.72  },
  Cherthala:          { male: 48.48, female: 51.52, Nair: 11.47, Ezhava: 41.95, Muslim: 2.30,  Christian: 21.70, "SC/ST": 5.59,  Others: 16.88 },
  Chirayankeezhu:     { male: 46.37, female: 53.63, Nair: 14.63, Ezhava: 23.56, Muslim: 17.90, Christian: 15.20, "SC/ST": 16.34, Others: 11.97 },
  Devikulam:          { male: 49.26, female: 50.74, Nair: 5.00,  Ezhava: 10.00, Muslim: 5.00,  Christian: 30.00, "SC/ST": 30.00, Others: 20.00 },
  Ettumanoor:         { male: 48.89, female: 51.11, Nair: 13.86, Ezhava: 24.32, Muslim: 5.37,  Christian: 40.46, "SC/ST": 6.17,  Others: 9.82  },
  Guruvayoor:         { male: 48.73, female: 51.27, Nair: 12.00, Ezhava: 20.00, Muslim: 20.00, Christian: 20.00, "SC/ST": 10.00, Others: 18.00 },
  Harippad:           { male: 47.61, female: 52.39, Nair: 21.16, Ezhava: 40.40, Muslim: 10.66, Christian: 8.90,  "SC/ST": 9.11,  Others: 9.63  },
  Idukki:             { male: 49.53, female: 50.47, Nair: 6.30,  Ezhava: 17.23, Muslim: 3.22,  Christian: 54.76, "SC/ST": 10.16, Others: 8.33  },
  Irinjalakuda:       { male: 48.29, female: 51.71, Nair: 9.30,  Ezhava: 27.90, Muslim: 6.49,  Christian: 31.46, "SC/ST": 14.21, Others: 10.60 },
  Irinjalakkuda:      { male: 48.29, female: 51.71, Nair: 9.30,  Ezhava: 27.90, Muslim: 6.49,  Christian: 31.46, "SC/ST": 14.21, Others: 10.60 },
  Irikkur:            { male: 49.58, female: 50.42, Nair: 8.84,  Ezhava: 8.03,  Muslim: 16.26, Christian: 43.57, "SC/ST": 7.83,  Others: 15.46 },
  Kalpetta:           { male: 48.76, female: 51.24, Nair: 6.00,  Ezhava: 15.00, Muslim: 15.00, Christian: 20.00, "SC/ST": 20.00, Others: 24.00 },
  Kanhangad:          { male: 48.62, female: 51.37, Nair: 11.78, Ezhava: 17.01, Muslim: 20.11, Christian: 14.47, "SC/ST": 10.40, Others: 26.24 },
  Kanjirappally:      { male: 48.55, female: 51.45, Nair: 23.92, Ezhava: 12.00, Muslim: 10.20, Christian: 40.00, "SC/ST": 9.66,  Others: 4.21  },
  Karunagappally:     { male: 48.64, female: 51.35, Nair: 16.62, Ezhava: 33.78, Muslim: 23.90, Christian: 4.00,  "SC/ST": 8.15,  Others: 13.35 },
  Kasaragod:          { male: 50.00, female: 50.00, Nair: 3.30,  Ezhava: 15.00, Muslim: 50.42, Christian: 2.40,  "SC/ST": 6.71,  Others: 22.17 },
  Kattakkada:         { male: 48.01, female: 51.99, Nair: 34.99, Ezhava: 14.83, Muslim: 6.04,  Christian: 22.27, "SC/ST": 11.37, Others: 10.07 },
  Kayamkulam:         { male: 47.75, female: 52.25, Nair: 24.39, Ezhava: 33.06, Muslim: 17.50, Christian: 8.90,  "SC/ST": 11.22, Others: 4.73  },
  Kazhakkoottam:      { male: 47.82, female: 52.18, Nair: 27.98, Ezhava: 28.30, Muslim: 7.30,  Christian: 14.50, "SC/ST": 10.20, Others: 11.72 },
  Kodungallur:        { male: 48.68, female: 51.32, Nair: 10.30, Ezhava: 26.00, Muslim: 16.96, Christian: 22.51, "SC/ST": 10.48, Others: 13.73 },
  Konni:              { male: 47.57, female: 52.43, Nair: 23.45, Ezhava: 17.33, Muslim: 5.00,  Christian: 30.70, "SC/ST": 14.70, Others: 8.82  },
  Kottarakkara:       { male: 47.39, female: 52.61, Nair: 32.94, Ezhava: 15.14, Muslim: 5.08,  Christian: 21.17, "SC/ST": 16.19, Others: 9.49  },
  Kottayam:           { male: 48.13, female: 51.87, Nair: 16.15, Ezhava: 20.06, Muslim: 4.97,  Christian: 43.23, "SC/ST": 7.25,  Others: 8.23  },
  Kovalam:            { male: 48.88, female: 51.11, Nair: 14.71, Ezhava: 14.35, Muslim: 9.56,  Christian: 35.91, "SC/ST": 11.66, Others: 13.70 },
  "Kozhikode North":  { male: 47.41, female: 52.59, Nair: 14.07, Ezhava: 32.16, Muslim: 25.10, Christian: 7.90,  "SC/ST": 4.43,  Others: 16.34 },
  Kunnamkulam:        { male: 48.56, female: 51.43, Nair: 12.90, Ezhava: 22.80, Muslim: 20.14, Christian: 21.37, "SC/ST": 13.84, Others: 8.95  },
  Kunnathunad:        { male: 48.85, female: 51.14, Nair: 11.78, Ezhava: 14.57, Muslim: 19.70, Christian: 35.40, "SC/ST": 13.13, Others: 5.42  },
  Kunnathur:          { male: 47.69, female: 52.31, Nair: 30.78, Ezhava: 13.64, Muslim: 13.70, Christian: 15.10, "SC/ST": 18.53, Others: 8.26  },
  Kuttanad:           { male: 49.24, female: 50.76, Nair: 14.32, Ezhava: 28.56, Muslim: 1.23,  Christian: 38.90, "SC/ST": 9.64,  Others: 7.34  },
  Malampuzha:         { male: 48.76, female: 51.24, Nair: 10.75, Ezhava: 33.89, Muslim: 11.31, Christian: 6.01,  "SC/ST": 15.26, Others: 22.76 },
  Manalur:            { male: 48.89, female: 51.11, Nair: 8.60,  Ezhava: 29.90, Muslim: 21.05, Christian: 21.52, "SC/ST": 8.28,  Others: 10.67 },
  Manjeshwaram:       { male: 50.38, female: 49.62, Nair: 0.44,  Ezhava: 12.00, Muslim: 52.89, Christian: 2.70,  "SC/ST": 6.37,  Others: 25.60 },
  Mankada:            { male: 49.91, female: 50.09, Nair: 8.00,  Ezhava: 10.00, Muslim: 45.00, Christian: 3.00,  "SC/ST": 14.00, Others: 20.00 },
  Mavelikkara:        { male: 47.00, female: 52.99, Nair: 31.08, Ezhava: 22.00, Muslim: 9.60,  Christian: 14.50, "SC/ST": 16.42, Others: 6.10  },
  Nattika:            { male: 48.09, female: 51.91, Nair: 9.00,  Ezhava: 34.70, Muslim: 16.31, Christian: 14.12, "SC/ST": 10.45, Others: 15.26 },
  Nedumangad:         { male: 47.58, female: 52.42, Nair: 34.61, Ezhava: 13.78, Muslim: 20.74, Christian: 9.80,  "SC/ST": 10.91, Others: 9.90  },
  Nemom:              { male: 48.30, female: 51.69, Nair: 30.88, Ezhava: 13.25, Muslim: 15.30, Christian: 8.50,  "SC/ST": 9.82,  Others: 22.30 },
  Nenmara:            { male: 49.60, female: 50.40, Nair: 7.97,  Ezhava: 26.28, Muslim: 17.06, Christian: 3.28,  "SC/ST": 21.66, Others: 23.74 },
  Ollur:              { male: 48.55, female: 51.45, Nair: 8.40,  Ezhava: 24.00, Muslim: 3.97,  Christian: 40.10, "SC/ST": 8.97,  Others: 14.51 },
  Ottapalam:          { male: 48.54, female: 51.46, Nair: 18.52, Ezhava: 17.83, Muslim: 29.06, Christian: 2.31,  "SC/ST": 13.36, Others: 18.87 },
  Pala:               { male: 48.71, female: 51.29, Nair: 16.41, Ezhava: 13.73, Muslim: 1.58,  Christian: 56.26, "SC/ST": 8.39,  Others: 3.61  },
  Palakkad:           { male: 48.63, female: 51.37, Nair: 9.66,  Ezhava: 22.08, Muslim: 27.90, Christian: 2.94,  "SC/ST": 11.89, Others: 25.37 },
  Peerumade:          { male: 49.32, female: 50.68, Nair: 4.08,  Ezhava: 16.34, Muslim: 5.97,  Christian: 42.98, "SC/ST": 25.15, Others: 5.48  },
  Peravoor:           { male: 49.21, female: 50.79, Nair: 8.00,  Ezhava: 12.00, Muslim: 15.00, Christian: 30.00, "SC/ST": 15.00, Others: 20.00 },
  Perumbavoor:        { male: 49.14, female: 50.85, Nair: 12.39, Ezhava: 12.39, Muslim: 18.60, Christian: 35.50, "SC/ST": 10.07, Others: 11.04 },
  Ponnani:            { male: 49.56, female: 50.44, Nair: 5.00,  Ezhava: 5.00,  Muslim: 60.00, Christian: 3.00,  "SC/ST": 7.00,  Others: 20.00 },
  Poonjar:            { male: 49.51, female: 50.49, Nair: 7.30,  Ezhava: 15.11, Muslim: 20.39, Christian: 39.26, "SC/ST": 11.37, Others: 6.58  },
  Puthukkad:          { male: 49.08, female: 50.92, Nair: 11.40, Ezhava: 25.90, Muslim: 5.91,  Christian: 30.80, "SC/ST": 11.53, Others: 14.41 },
  Ranni:              { male: 48.67, female: 51.33, Nair: 17.27, Ezhava: 14.44, Muslim: 4.70,  Christian: 47.00, "SC/ST": 10.40, Others: 6.19  },
  Shornur:            { male: 48.81, female: 51.19, Nair: 18.07, Ezhava: 15.39, Muslim: 31.77, Christian: 1.27,  "SC/ST": 14.78, Others: 18.68 },
  "Sultan Bathery":   { male: 48.63, female: 51.37, Nair: 6.40,  Ezhava: 17.46, Muslim: 16.74, Christian: 24.65, "SC/ST": 22.68, Others: 11.65 },
  "Sulthan Bathery":  { male: 48.63, female: 51.37, Nair: 6.40,  Ezhava: 17.46, Muslim: 16.74, Christian: 24.65, "SC/ST": 22.68, Others: 11.65 },
  Thiruvalla:         { male: 47.98, female: 52.02, Nair: 16.78, Ezhava: 10.36, Muslim: 2.10,  Christian: 48.30, "SC/ST": 11.85, Others: 10.41 },
  Thiruvambady:       { male: 49.45, female: 50.55, Nair: 4.75,  Ezhava: 8.86,  Muslim: 44.68, Christian: 23.66, "SC/ST": 10.01, Others: 8.04  },
  Thiruvananthapuram: { male: 48.06, female: 51.93, Nair: 23.21, Ezhava: 8.43,  Muslim: 18.00, Christian: 24.00, "SC/ST": 9.09,  Others: 17.27 },
  Thodupuzha:         { male: 49.37, female: 50.63, Nair: 7.84,  Ezhava: 15.29, Muslim: 16.63, Christian: 44.10, "SC/ST": 10.18, Others: 5.90  },
  Thrikkakara:        { male: 48.06, female: 51.94, Nair: 12.00, Ezhava: 18.00, Muslim: 15.00, Christian: 30.00, "SC/ST": 8.00,  Others: 17.00 },
  Thripunithura:      { male: 48.32, female: 51.68, Nair: 11.08, Ezhava: 21.55, Muslim: 11.82, Christian: 26.31, "SC/ST": 7.99,  Others: 20.95 },
  Thrissur:           { male: 47.55, female: 52.45, Nair: 16.30, Ezhava: 14.00, Muslim: 5.20,  Christian: 38.70, "SC/ST": 7.85,  Others: 17.96 },
  Udumbanchola:       { male: 49.42, female: 50.58, Nair: 7.38,  Ezhava: 21.09, Muslim: 4.59,  Christian: 42.68, "SC/ST": 10.96, Others: 13.29 },
  Vaikom:             { male: 48.72, female: 51.28, Nair: 15.44, Ezhava: 34.18, Muslim: 4.75,  Christian: 17.98, "SC/ST": 12.57, Others: 15.01 },
  Vamanapuram:        { male: 46.98, female: 53.02, Nair: 25.86, Ezhava: 12.63, Muslim: 23.06, Christian: 9.58,  "SC/ST": 14.71, Others: 14.16 },
  Varkala:            { male: 46.99, female: 53.00, Nair: 21.51, Ezhava: 24.32, Muslim: 28.70, Christian: 1.10,  "SC/ST": 15.70, Others: 8.57  },
  Vattiyoorkavu:      { male: 47.56, female: 52.44, Nair: 35.23, Ezhava: 10.24, Muslim: 6.00,  Christian: 18.00, "SC/ST": 10.20, Others: 20.33 },
  Wadakkanchery:      { male: 48.29, female: 51.70, Nair: 13.40, Ezhava: 21.60, Muslim: 6.72,  Christian: 29.45, "SC/ST": 11.25, Others: 17.40 },
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
