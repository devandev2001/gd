/**
 * Kerala Survey 2026 — survey weights aligned with scripts/weight_ground_input.py
 *
 * Caste / Gender / Age: raking weight = population target ÷ sample share within each AC
 *   (targets from FINAL GENDER CASTE sheet + AGE_WEIGHTS).
 * GevsVE (column O): seven ACs only — historical vote share as fraction ÷ (party count / n_ac);
 *   Kasaragod & Nemom use Vote 2024 GE (K); others use Vote 2021 AE (J). Others = 1 − LDF − UDF − BJP.
 * On append: all rows in that AC get new D–H and N; then O and P (P = O × N).
 *
 * HOW TO SET UP:
 * 1. Open your Google Sheet and go to Extensions → Apps Script
 * 2. Delete everything in Code.gs and paste THIS entire file.
 * 3. Deploy → New deployment (Web app, Execute as Me, Anyone)
 * 4. Copy the Web App URL → src/data/surveyData.js → GOOGLE_SCRIPT_URL
 * 5. Click "Authorize" when prompted.
 * 6. SET UP DAILY REPORT TRIGGER:
 *    Triggers → + Add Trigger → generateDailyReport → Time-driven → Day timer → 8pm–9pm
 */

var MAIN_SHEET_NAME = "Sheet1";
var ENTRIES_CACHE_TTL_SEC = 60;
var ENTRIES_CACHE_MAX_CHARS = 95000;
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

var CASTE_LIST = ["Nair","Ezhava","Muslim","Christian","SC/ST","Others"];

var AC_DEMOGRAPHICS_FALLBACK = {
  "Adoor": {"male": 47.21, "female": 52.79, "Muslim": 6.8, "Christian": 26.4, "Nair": 25.15, "Ezhava": 19.67, "Others": 3.15, "SC/ST": 18.84},
  "Alappuzha": {"male": 48.77, "female": 51.23},
  "Alathur": {"male": 49.34, "female": 50.66},
  "Aluva": {"male": 48.64, "female": 51.36},
  "Ambalapuzha": {"male": 48.78, "female": 51.21, "Muslim": 24.9, "Christian": 15.6, "Nair": 13.69, "Ezhava": 27.72, "Others": 14.33, "SC/ST": 3.76},
  "Angamaly": {"male": 49.32, "female": 50.68},
  "Aranmula": {"male": 47.95, "female": 52.05, "Muslim": 4.2, "Christian": 38.7, "Nair": 20.89, "Ezhava": 16.37, "Others": 4.14, "SC/ST": 15.6},
  "Aroor": {"male": 49.05, "female": 50.95, "Muslim": 10.8, "Christian": 19.4, "Nair": 13.24, "Ezhava": 37.2, "Others": 6.65, "SC/ST": 12.61},
  "Aruvikkara": {"male": 47.49, "female": 52.5},
  "Attingal": {"male": 46.28, "female": 53.72, "Muslim": 17.3, "Christian": 1.7, "Nair": 19.1, "Ezhava": 28.12, "Others": 16.21, "SC/ST": 17.47},
  "Azhikode": {"male": 47.58, "female": 52.42},
  "Balussery": {"male": 48.58, "female": 51.41},
  "Beypore": {"male": 48.97, "female": 51.03},
  "Chadayamangalam": {"male": 47.41, "female": 52.58},
  "Chalakkudy": {"male": 48.87, "female": 51.13, "Muslim": 2.91, "Christian": 49.09, "Nair": 7.7, "Ezhava": 17.8, "Others": 10.97, "SC/ST": 11.58},
  "Changanassery": {"male": 48.2, "female": 51.8, "Muslim": 7.9, "Christian": 45.4, "Nair": 12.57, "Ezhava": 17.34, "Others": 6.44, "SC/ST": 10.35},
  "Chathannoor": {"male": 46.81, "female": 53.19, "Muslim": 12.1, "Christian": 10.3, "Nair": 26.85, "Ezhava": 29.6, "Others": 6.25, "SC/ST": 14.7},
  "Chavara": {"male": 48.69, "female": 51.31},
  "Chelakkara": {"male": 48.42, "female": 51.58, "Muslim": 26.93, "Christian": 6.51, "Nair": 20.0, "Ezhava": 12.0, "Others": 18.71, "SC/ST": 15.9},
  "Chengannur": {"male": 47.64, "female": 52.35, "Muslim": 3.89, "Christian": 26.81, "Nair": 29.92, "Ezhava": 15.54, "Others": 7.72, "SC/ST": 15.97},
  "Cherthala": {"male": 48.48, "female": 51.52, "Muslim": 2.3, "Christian": 21.7, "Nair": 11.47, "Ezhava": 41.95, "Others": 16.88, "SC/ST": 5.59},
  "Chirayankeezhu": {"male": 46.37, "female": 53.63, "Muslim": 17.9, "Christian": 15.2, "Nair": 14.63, "Ezhava": 23.56, "Others": 11.97, "SC/ST": 16.34},
  "Chittur": {"male": 48.84, "female": 51.16},
  "Devikulam": {"male": 49.26, "female": 50.74},
  "Dharmadam": {"male": 47.49, "female": 52.5},
  "Elathur": {"male": 48.26, "female": 51.74},
  "Eranad": {"male": 50.34, "female": 49.66},
  "Eravipuram": {"male": 48.12, "female": 51.88},
  "Ernakulam": {"male": 48.15, "female": 51.84},
  "Ettumanoor": {"male": 48.89, "female": 51.11, "Muslim": 5.37, "Christian": 40.46, "Nair": 13.86, "Ezhava": 24.32, "Others": 9.82, "SC/ST": 6.17},
  "Guruvayoor": {"male": 48.73, "female": 51.27},
  "Harippad": {"male": 47.61, "female": 52.39, "Muslim": 10.66, "Christian": 8.9, "Nair": 21.16, "Ezhava": 40.4, "Others": 9.63, "SC/ST": 9.11},
  "Idukki": {"male": 49.53, "female": 50.47, "Muslim": 3.22, "Christian": 54.76, "Nair": 6.3, "Ezhava": 17.23, "Others": 8.33, "SC/ST": 10.16},
  "Irikkur": {"male": 49.58, "female": 50.42, "Muslim": 16.26, "Christian": 43.57, "Nair": 8.84, "Ezhava": 8.03, "Others": 15.46, "SC/ST": 7.83},
  "Irinjalakuda": {"male": 48.29, "female": 51.71, "Muslim": 6.49, "Christian": 31.46, "Nair": 9.3, "Ezhava": 27.9, "Others": 10.6, "SC/ST": 14.21},
  "Kaduthuruthy": {"male": 48.77, "female": 51.23},
  "Kaipamangalam": {"male": 48.13, "female": 51.87},
  "Kalamassery": {"male": 48.45, "female": 51.55},
  "Kalliasseri": {"male": 47.73, "female": 52.27},
  "Kalpetta": {"male": 48.76, "female": 51.24},
  "Kanhangad": {"male": 48.62, "female": 51.37, "Muslim": 20.11, "Christian": 14.47, "Nair": 11.78, "Ezhava": 17.01, "Others": 26.24, "SC/ST": 10.4},
  "Kanjirappally": {"male": 48.55, "female": 51.45, "Muslim": 10.2, "Christian": 40.0, "Nair": 23.92, "Ezhava": 12.0, "Others": 4.21, "SC/ST": 9.66},
  "Kannur": {"male": 47.97, "female": 52.03},
  "Karunagappally": {"male": 48.64, "female": 51.35, "Muslim": 23.9, "Christian": 4.0, "Nair": 16.62, "Ezhava": 33.78, "Others": 13.35, "SC/ST": 8.15},
  "Kasaragod": {"male": 50.0, "female": 50.0, "Muslim": 50.42, "Christian": 2.4, "Nair": 3.3, "Ezhava": 15.0, "Others": 22.17, "SC/ST": 6.71},
  "Kattakkada": {"male": 48.01, "female": 51.99, "Muslim": 6.04, "Christian": 22.27, "Nair": 34.99, "Ezhava": 14.83, "Others": 10.07, "SC/ST": 11.37},
  "Kayamkulam": {"male": 47.75, "female": 52.25, "Muslim": 17.5, "Christian": 8.9, "Nair": 24.39, "Ezhava": 33.06, "Others": 4.73, "SC/ST": 11.22},
  "Kazhakkoottam": {"male": 47.82, "female": 52.18, "Muslim": 7.3, "Christian": 14.5, "Nair": 27.98, "Ezhava": 28.3, "Others": 11.72, "SC/ST": 10.2},
  "Kochi": {"male": 48.64, "female": 51.36},
  "Kodungallur": {"male": 48.68, "female": 51.32, "Muslim": 16.96, "Christian": 22.51, "Nair": 10.3, "Ezhava": 26.0, "Others": 13.73, "SC/ST": 10.48},
  "Koduvally": {"male": 49.66, "female": 50.34},
  "Kollam": {"male": 48.17, "female": 51.83},
  "Kondotty": {"male": 50.56, "female": 49.44},
  "Kongad": {"male": 49.21, "female": 50.79},
  "Konni": {"male": 47.57, "female": 52.43, "Muslim": 5.0, "Christian": 30.7, "Nair": 23.45, "Ezhava": 17.33, "Others": 8.82, "SC/ST": 14.7},
  "Kothamangalam": {"male": 49.23, "female": 50.77},
  "Kottakkal": {"male": 50.54, "female": 49.46},
  "Kottarakkara": {"male": 47.39, "female": 52.61, "Muslim": 5.08, "Christian": 21.17, "Nair": 32.94, "Ezhava": 15.14, "Others": 9.49, "SC/ST": 16.19},
  "Kottayam": {"male": 48.13, "female": 51.87, "Muslim": 4.97, "Christian": 43.23, "Nair": 16.15, "Ezhava": 20.06, "Others": 8.23, "SC/ST": 7.25},
  "Kovalam": {"male": 48.88, "female": 51.11, "Muslim": 9.56, "Christian": 35.91, "Nair": 14.71, "Ezhava": 14.35, "Others": 13.7, "SC/ST": 11.66},
  "Kozhikode": {"male": 48.82, "female": 51.18},
  "Kozhikode North": {"male": 47.41, "female": 52.59, "Muslim": 25.1, "Christian": 7.9, "Nair": 14.07, "Ezhava": 32.16, "Others": 16.34, "SC/ST": 4.43},
  "Kozhikode South": {"Muslim": 48.0, "Christian": 3.0, "Nair": 10.29, "Ezhava": 23.52, "Others": 11.94, "SC/ST": 3.25},
  "Kundara": {"male": 47.82, "female": 52.17},
  "Kunnamangalam": {"male": 48.79, "female": 51.21},
  "Kunnamkulam": {"male": 48.56, "female": 51.43, "Muslim": 20.14, "Christian": 21.37, "Nair": 12.9, "Ezhava": 22.8, "Others": 8.95, "SC/ST": 13.84},
  "Kunnathunad": {"male": 48.85, "female": 51.14},
  "Kunnathur": {"male": 47.69, "female": 52.31, "Muslim": 13.7, "Christian": 15.1, "Nair": 30.78, "Ezhava": 13.64, "Others": 8.26, "SC/ST": 18.53},
  "Kuthuparamba": {"male": 48.67, "female": 51.33},
  "Kuttanad": {"male": 49.24, "female": 50.76, "Muslim": 1.23, "Christian": 38.9, "Nair": 14.32, "Ezhava": 28.56, "Others": 7.34, "SC/ST": 9.64},
  "Kuttiadi": {"male": 49.19, "female": 50.8},
  "Malampuzha": {"male": 48.76, "female": 51.24, "Muslim": 11.31, "Christian": 6.01, "Nair": 10.75, "Ezhava": 33.89, "Others": 22.76, "SC/ST": 15.26},
  "Malappuram": {"male": 50.38, "female": 49.62},
  "Manalur": {"male": 48.89, "female": 51.11, "Muslim": 21.05, "Christian": 21.52, "Nair": 8.6, "Ezhava": 29.9, "Others": 10.67, "SC/ST": 8.28},
  "Mananthavady": {"male": 49.26, "female": 50.74},
  "Manjeri": {"male": 49.86, "female": 50.14},
  "Manjeshwaram": {"male": 50.38, "female": 49.62, "Muslim": 52.89, "Christian": 2.7, "Nair": 0.44, "Ezhava": 12.0, "Others": 25.6, "SC/ST": 6.37},
  "Mankada": {"male": 49.91, "female": 50.09},
  "Mannarkad": {"male": 49.12, "female": 50.88},
  "Mattannur": {"male": 48.27, "female": 51.73},
  "Mavelikkara": {"male": 47.0, "female": 52.99, "Muslim": 9.6, "Christian": 14.5, "Nair": 31.08, "Ezhava": 22.0, "Others": 6.1, "SC/ST": 16.42},
  "Muvattupuzha": {"male": 49.14, "female": 50.86},
  "Nadapuram": {"male": 49.69, "female": 50.31},
  "Nattika": {"male": 48.09, "female": 51.91, "Muslim": 16.31, "Christian": 14.12, "Nair": 9.0, "Ezhava": 34.7, "Others": 15.26, "SC/ST": 10.45},
  "Nedumangad": {"male": 47.58, "female": 52.42, "Muslim": 20.74, "Christian": 9.8, "Nair": 34.61, "Ezhava": 13.78, "Others": 9.9, "SC/ST": 10.91},
  "Nemom": {"male": 48.3, "female": 51.69, "Muslim": 15.3, "Christian": 8.5, "Nair": 30.88, "Ezhava": 13.25, "Others": 22.3, "SC/ST": 9.82},
  "Nenmara": {"male": 49.6, "female": 50.4, "Muslim": 17.06, "Christian": 3.28, "Nair": 7.97, "Ezhava": 26.28, "Others": 23.74, "SC/ST": 21.66},
  "Neyyattinkara": {"male": 49.12, "female": 50.88},
  "Nilambur": {"male": 48.97, "female": 51.02},
  "Ollur": {"male": 48.55, "female": 51.45, "Muslim": 3.97, "Christian": 40.1, "Nair": 8.4, "Ezhava": 24.0, "Others": 14.51, "SC/ST": 8.97},
  "Ottapalam": {"male": 48.54, "female": 51.46, "Muslim": 29.06, "Christian": 2.31, "Nair": 18.52, "Ezhava": 17.83, "Others": 18.87, "SC/ST": 13.36},
  "Pala": {"male": 48.71, "female": 51.29, "Muslim": 1.58, "Christian": 56.26, "Nair": 16.41, "Ezhava": 13.73, "Others": 3.61, "SC/ST": 8.39},
  "Palakkad": {"male": 48.63, "female": 51.37, "Muslim": 27.9, "Christian": 2.94, "Nair": 9.66, "Ezhava": 22.08, "Others": 25.37, "SC/ST": 11.89},
  "Parassala": {"male": 48.58, "female": 51.42},
  "Paravur": {"male": 48.75, "female": 51.25},
  "Pathanapuram": {"male": 47.38, "female": 52.62},
  "Pattambi": {"male": 49.55, "female": 50.45},
  "Payyannur": {"male": 48.11, "female": 51.89},
  "Peerumade": {"male": 49.32, "female": 50.68, "Muslim": 5.97, "Christian": 42.98, "Nair": 4.08, "Ezhava": 16.34, "Others": 5.48, "SC/ST": 25.15},
  "Perambra": {"male": 48.95, "female": 51.05},
  "Peravoor": {"male": 49.21, "female": 50.79},
  "Perinthalmanna": {"male": 49.42, "female": 50.58},
  "Perumbavoor": {"male": 49.14, "female": 50.85, "Muslim": 18.6, "Christian": 35.5, "Nair": 12.39, "Ezhava": 12.39, "Others": 11.04, "SC/ST": 10.07},
  "Piravom": {"male": 48.4, "female": 51.6},
  "Ponnani": {"male": 49.56, "female": 50.44},
  "Poonjar": {"male": 49.51, "female": 50.49, "Muslim": 20.39, "Christian": 39.26, "Nair": 7.3, "Ezhava": 15.11, "Others": 6.58, "SC/ST": 11.37},
  "Punalur": {"male": 47.56, "female": 52.44},
  "Puthukkad": {"male": 49.08, "female": 50.92, "Muslim": 5.91, "Christian": 30.8, "Nair": 11.4, "Ezhava": 25.9, "Others": 14.41, "SC/ST": 11.53},
  "Puthuppally": {"male": 48.71, "female": 51.28},
  "Quilandi": {"male": 48.46, "female": 51.54},
  "Ranni": {"male": 48.67, "female": 51.33, "Muslim": 4.7, "Christian": 47.0, "Nair": 17.27, "Ezhava": 14.44, "Others": 6.19, "SC/ST": 10.4},
  "Shornur": {"male": 48.81, "female": 51.19, "Muslim": 31.77, "Christian": 1.27, "Nair": 18.07, "Ezhava": 15.39, "Others": 18.68, "SC/ST": 14.78},
  "Sultan Bathery": {"male": 48.63, "female": 51.37, "Muslim": 16.74, "Christian": 24.65, "Nair": 6.4, "Ezhava": 17.46, "Others": 11.65, "SC/ST": 22.68},
  "Taliparamba": {"male": 48.09, "female": 51.91},
  "Tanur": {"male": 50.48, "female": 49.52},
  "Tarur": {"male": 49.21, "female": 50.79},
  "Thalassery": {"male": 47.8, "female": 52.2},
  "Thavanur": {"male": 49.75, "female": 50.25},
  "Thiruvalla": {"male": 47.98, "female": 52.02, "Muslim": 2.1, "Christian": 48.3, "Nair": 16.78, "Ezhava": 10.36, "Others": 10.41, "SC/ST": 11.85},
  "Thiruvambady": {"male": 49.45, "female": 50.55, "Muslim": 44.68, "Christian": 23.66, "Nair": 4.75, "Ezhava": 8.86, "Others": 8.04, "SC/ST": 10.01},
  "Thiruvananthapuram": {"male": 48.06, "female": 51.93, "Muslim": 18.0, "Christian": 24.0, "Nair": 23.21, "Ezhava": 8.43, "Others": 17.27, "SC/ST": 9.09},
  "Thodupuzha": {"male": 49.37, "female": 50.63, "Muslim": 16.63, "Christian": 44.1, "Nair": 7.84, "Ezhava": 15.29, "Others": 5.9, "SC/ST": 10.18},
  "Thrikkakara": {"male": 48.06, "female": 51.94},
  "Thripunithura": {"male": 48.32, "female": 51.68, "Muslim": 11.82, "Christian": 26.31, "Nair": 11.08, "Ezhava": 21.55, "Others": 20.95, "SC/ST": 7.99},
  "Thrissur": {"male": 47.55, "female": 52.45, "Muslim": 5.2, "Christian": 38.7, "Nair": 16.3, "Ezhava": 14.0, "Others": 17.96, "SC/ST": 7.85},
  "Thrithala": {"male": 49.18, "female": 50.82},
  "Tirur": {"male": 50.12, "female": 49.87},
  "Tirurangadi": {"male": 50.73, "female": 49.27},
  "Trikaripur": {"male": 48.52, "female": 51.48},
  "Udma": {"male": 49.06, "female": 50.93},
  "Udumbanchola": {"male": 49.42, "female": 50.58, "Muslim": 4.59, "Christian": 42.68, "Nair": 7.38, "Ezhava": 21.09, "Others": 13.29, "SC/ST": 10.96},
  "Vadakara": {"male": 48.63, "female": 51.37},
  "Vaikom": {"male": 48.72, "female": 51.28, "Muslim": 4.75, "Christian": 17.98, "Nair": 15.44, "Ezhava": 34.18, "Others": 15.01, "SC/ST": 12.57},
  "Vallikkunnu": {"male": 50.38, "female": 49.62},
  "Vamanapuram": {"male": 46.98, "female": 53.02, "Muslim": 23.06, "Christian": 9.58, "Nair": 25.86, "Ezhava": 12.63, "Others": 14.16, "SC/ST": 14.71},
  "Varkala": {"male": 46.99, "female": 53.0, "Muslim": 28.7, "Christian": 1.1, "Nair": 21.51, "Ezhava": 24.32, "Others": 8.57, "SC/ST": 15.7},
  "Vattiyoorkavu": {"male": 47.56, "female": 52.44, "Muslim": 6.0, "Christian": 18.0, "Nair": 35.23, "Ezhava": 10.24, "Others": 20.33, "SC/ST": 10.2},
  "Vengara": {"male": 51.21, "female": 48.79},
  "Vypin": {"male": 48.69, "female": 51.31},
  "Wadakkanchery": {"male": 48.29, "female": 51.7, "Muslim": 6.72, "Christian": 29.45, "Nair": 13.4, "Ezhava": 21.6, "Others": 17.4, "SC/ST": 11.25},
  "Wandoor": {"male": 49.13, "female": 50.87}
};

var AC_FA_ROSTER = {
  "Kattakkada": ["Dinu S", "Ajin Anand"],
  "Kovalam": ["Anandhu RS", "Arun R"],
  "Vattiyoorkavu": ["Gokul G", "Jijomon SP"],
  "Thiruvananthapuram": ["Sudheesh", "Akhil R"],
  "Attingal": ["Anand KS", "SREEHARI .S"],
  "Chathannoor": ["Aromal A", "Sajikumar S"],
  "Aranmula": ["Saneesh", "Syam Chandran"],
  "Thiruvalla": ["Nikhil MR", "Nidhin P Nair"],
  "Chengannur": ["Aswin S", "Ananthu R Pillai"],
  "Adoor": ["KG Gokulkrishnan Nair", "Vyshakh  V"],
  "Poonjar": ["Sajan P Nair", "Sreehari Babu"],
  "Pala": ["Anirudh Vinod", "Vishnu B Ambalammatom"],
  "Thrissur": ["Jitesh", "Akshay", "Sree Hari"],
  "Kunnathunad": ["Sarath Sadan", "Adarshjith"],
  "Palakkad": ["Goshal krishna", "Vignesh R"],
  "Kozhikode North": ["Abhishek", "Abhinav KP"],
  "Kasaragod": ["Sivadas MM"],
  "Manjeshwaram": ["Viswajith BV", "Vishnu B"],
  "Nemom": ["Anandhu RS", "Arun R"],
  "Kazhakkoottam": ["Abhinanad BS"],
  "Nattika": ["Roshith", "Amal ks"],
  "Malampuzha": ["Jayakrishnan", "Adeep das"],
  "Manalur": ["Sooraj Krishna", "SANJAY"],
  "Perumbavoor": ["Sarath Sadan", "JithuKrishanan Babu"],
  "Kanjirappally": ["Gokul PG", "Harikrishnan AB"],
  "Thripunithura": ["FA1", "FA2"],
  "Thrikkakara": ["FA1", "FA2"]
};

// ═══════════════════════════════════════════════════════════
// UTILITY HELPERS
// ═══════════════════════════════════════════════════════════

var _DEMOGRAPHICS_CACHE = null;

function pad2(n) {
  n = parseInt(n, 10);
  if (n < 10) return "0" + n;
  return String(n);
}

function asFiniteNumber(val, fallback) {
  if (val === null || val === undefined || val === "") return fallback;
  var n;
  if (typeof val === "number") { n = val; }
  else { n = parseFloat(String(val).replace(/,/g, "").replace(/%/g, "")); }
  if (isFinite(n)) return n;
  return fallback;
}

function isTextLabel(val) {
  var s = String(val === null || val === undefined ? "" : val).trim();
  if (s === "") return false;
  return isNaN(parseFloat(s));
}

function normalizeAgeLabel(label) {
  var s = String(label || "").trim().replace(/[\u2013\u2014]/g, "-");
  if (/^18\s*-\s*19$/i.test(s)) return "18-19";
  return s;
}

function normalizeAgeLabelForWeight(label) {
  return normalizeAgeLabel(label);
}

function canonicalCasteKey(casteStr) {
  var s = String(casteStr || "").trim();
  if (!s) return s;
  var low = s.toLowerCase().replace(/\s+/g, " ");
  if (low === "sc/st" || low === "scst" || low === "sc-st" || low === "sc / st") return "SC/ST";
  for (var i = 0; i < CASTE_LIST.length; i++) {
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
  return (g === "male" || g === "female");
}

function displayGenderLabel(val) {
  var g = String(val || "").trim().toLowerCase();
  if (g === "male") return "Male";
  if (g === "female") return "Female";
  return "";
}

function isValidAgeLabel(val) {
  var k = normalizeAgeLabel(val);
  return AGE_WEIGHTS.hasOwnProperty(k);
}

function normalizeAcName(acRaw) {
  var s = String(acRaw || "").trim().replace(/[\u200B-\u200D\uFEFF]/g, "");
  var k = s.toLowerCase().replace(/\s+/g, " ");
  if (k === "kattakada") return "Kattakkada";
  if (k === "thripunitura" || k === "thrippunithura") return "Thripunithura";
  if (k === "thrikakkara") return "Thrikkakara";
  if (k === "kowalam") return "Kovalam";
  if (k === "trivandrum") return "Thiruvananthapuram";
  if (k === "naimam" || k === "nemam" || k === "nemeom" || k === "naiyamam" || k === "neyyattinkara") return "Nemom";
  if (k === "kasargod" || k === "kasragod" || k === "kasaragode" || k === "kasargode" || k === "kasaragod") return "Kasaragod";
  if (k === "manjeswaram" || k === "manjeshwar" || k === "manjeshwaram" || k === "manjeswar") return "Manjeshwaram";
  if (k === "nattika" || k === "nattika (sc)" || k === "nattika ac") return "Nattika";
  if (k === "thrissur ac" || k === "thrissur") return "Thrissur";
  if (k === "malampuzha") return "Malampuzha";
  if (k === "chengannur") return "Chengannur";
  if (k === "manalur ac" || k === "manalur") return "Manalur";
  if (k === "perumbaavoor" || k === "perumbavoor ac" || k === "perumbavoor") return "Perumbavoor";
  if (k === "kanjirapalli" || k === "kanjirappally" || k === "kanjirappalli") return "Kanjirappally";
  if (k === "kunnathunad" || k === "kunnathunad ac" || k === "kunnathumar" ||
      k === "kunnathunadu" || k === "kunnathunadu ac" || k === "kunnathunadac" ||
      k === "gunnar thunadu" || k === "kunnathumad") return "Kunnathunad";
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

function normalizeVoteForGevs(raw) {
  var s = String(raw === null || raw === undefined ? "" : raw).trim();
  if (!s || s.toLowerCase() === "nan") return "";
  var u = s.toUpperCase().replace(/\s+/g, "");
  if (u.indexOf("NOT") >= 0 && u.indexOf("VOT") >= 0) return "Not Voted";
  if (u === "LDF") return "LDF";
  if (u === "UDF") return "UDF";
  if (u === "BJP/NDA" || u === "BJP-NDA" || u === "BJPNDA" || u === "BJP" || u === "NDA") return "BJP/NDA";
  if (u === "OTHERS" || u === "OTHER") return "Others";
  return s;
}

function voteForSheet(v) {
  var p = normalizeParty(v);
  if (p === "UDF" || p === "LDF" || p === "BJP/NDA") return p;
  return String(v === null || v === undefined ? "" : v).trim();
}

// ═══════════════════════════════════════════════════════════
// DEMOGRAPHICS LOOKUP
// ═══════════════════════════════════════════════════════════

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
  var missing = [iAc, iMale, iFemale, iMuslim, iChristian, iNair, iEzhava, iOthers, iScst].some(function(x) { return x < 0; });
  if (missing) throw new Error("Demographics sheet headers missing/changed.");
  var map = {};
  for (var r = 1; r < data.length; r++) {
    var ac = normalizeAcName(data[r][iAc]);
    if (!ac) continue;
    map[ac] = {
      male: asFiniteNumber(data[r][iMale], 0),
      female: asFiniteNumber(data[r][iFemale], 0),
      Muslim: asFiniteNumber(data[r][iMuslim], 0),
      Christian: asFiniteNumber(data[r][iChristian], 0),
      Nair: asFiniteNumber(data[r][iNair], 0),
      Ezhava: asFiniteNumber(data[r][iEzhava], 0),
      Others: asFiniteNumber(data[r][iOthers], 0),
      "SC/ST": asFiniteNumber(data[r][iScst], 0)
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
    } catch (err) { Logger.log("Failed to read demographics: " + err); }
  }
  _DEMOGRAPHICS_CACHE = AC_DEMOGRAPHICS_FALLBACK;
  return _DEMOGRAPHICS_CACHE;
}

function clearDemographicsCache() { _DEMOGRAPHICS_CACHE = null; }

// ═══════════════════════════════════════════════════════════
// LABEL / WEIGHT REVERSE LOOKUP (for legacy/numeric rows)
// ═══════════════════════════════════════════════════════════

function getClosestAgeLabelFromNormalizedWeight(w) {
  var best = "Unknown", bestDiff = Infinity;
  for (var label in AGE_WEIGHTS) {
    if (AGE_WEIGHTS.hasOwnProperty(label)) {
      var diff = Math.abs(AGE_WEIGHTS[label] - w);
      if (diff < bestDiff) { bestDiff = diff; best = label; }
    }
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

function getCasteLabelFromWeight(ac, casteWeight) {
  if (isTextLabel(casteWeight)) {
    var t = canonicalCasteKey(casteWeight);
    if (isValidCasteLabel(t)) return t;
    return "Others";
  }
  var acData = getDemographicsMap()[normalizeAcName(ac)];
  if (!acData) return "Unknown";
  var w = asFiniteNumber(casteWeight, NaN);
  if (!isFinite(w) || w === 0) return "Unknown";
  if (w > 1.5) w = w / 100;
  var best = "Others", bestDiff = Infinity;
  for (var i = 0; i < CASTE_LIST.length; i++) {
    var diff = Math.abs((acData[CASTE_LIST[i]] || 0) / 100 - w);
    if (diff < bestDiff) { bestDiff = diff; best = CASTE_LIST[i]; }
  }
  return best;
}

function getGenderLabelFromWeight(ac, genderWeight) {
  if (isTextLabel(genderWeight)) {
    var t = String(genderWeight).trim();
    if (isValidGenderLabel(t)) return displayGenderLabel(t);
    return "Unknown";
  }
  var acData = getDemographicsMap()[normalizeAcName(ac)];
  if (!acData) return "Unknown";
  var w = asFiniteNumber(genderWeight, NaN);
  if (!isFinite(w) || w === 0) return "Unknown";
  if (w > 1.5) w = w / 100;
  var mD = Math.abs((acData.male / 100) - w);
  var fD = Math.abs((acData.female / 100) - w);
  if (mD < fD) return "Male";
  if (fD < mD) return "Female";
  if (Math.abs(acData.male - acData.female) < 1e-9) return "Unknown";
  if (acData.female >= acData.male) return "Female";
  return "Male";
}

function getAgeLabelFromWeight(ageWeight) {
  return getAgeLabelFromAny(ageWeight);
}

function casteLabelForSheet(ac, formCaste, casteW) {
  if (formCaste != null && isValidCasteLabel(formCaste)) return canonicalCasteKey(formCaste);
  return getCasteLabelFromWeight(ac, casteW);
}

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

// ═══════════════════════════════════════════════════════════
// RAKING WEIGHT CORE
// ═══════════════════════════════════════════════════════════

function computeRakingWeight(targetPct, acRows, valueFn) {
  var total = acRows.length;
  if (total === 0) return 1;
  var count = 0;
  for (var i = 0; i < total; i++) {
    if (valueFn(acRows[i])) count++;
  }
  if (count === 0) return 1;
  var sampleShare = count / total;
  var popShare = targetPct / 100;
  if (popShare <= 0) return 1;
  return popShare / sampleShare;
}

function resolveWeightsFromLabelsAndCounts(ac, casteLabel, genderLabel, ageLabel, acRows) {
  var demo = getDemographicsMap()[normalizeAcName(ac)];
  if (!demo) return { caste: 1, gender: 1, age: 1 };

  var castePct = demo[casteLabel] || 0;
  var cw = computeRakingWeight(castePct, acRows, function(row) {
    return canonicalCasteKey(row.caste) === casteLabel;
  });

  var gNorm = String(genderLabel).trim().toLowerCase();
  var genderPct = (gNorm === "male") ? demo.male : (gNorm === "female") ? demo.female : 0;
  var gw = computeRakingWeight(genderPct, acRows, function(row) {
    return String(row.gender).trim().toLowerCase() === gNorm;
  });

  var agePct = (AGE_WEIGHTS[ageLabel] || 0) * 100;
  var aw = computeRakingWeight(agePct, acRows, function(row) {
    return normalizeAgeLabel(row.age) === ageLabel;
  });

  return { caste: cw, gender: gw, age: aw };
}

function resolveWeights(ac, rawCaste, rawGender, rawAge, acRows) {
  var cl = casteLabelForSheet(ac, rawCaste, rawCaste);
  var gl = genderLabelForSheet(ac, rawGender, rawGender);
  var al = ageLabelForSheet(rawAge, rawAge);
  var wObj = resolveWeightsFromLabelsAndCounts(ac, cl, gl, al, acRows);
  return {
    casteWeight: wObj.caste,
    casteLabel: cl,
    genderWeight: wObj.gender,
    genderLabel: gl,
    ageWeight: wObj.age,
    ageLabel: al
  };
}

// ═══════════════════════════════════════════════════════════
// GEVS-VE (Historical vote alignment for 7 ACs)
// ═══════════════════════════════════════════════════════════
var GEVS_EMPTY_PARTY_KEY = "__empty__";

// Kasaragod and Nemom use col K (Vote 2024 GE); others use col J (Vote 2021 AE).
var GEVS_USE_COL_K = { "kasaragod": true, "nemom": true };

function gevsVoteColumnIndex(acNorm) {
  // Returns 0-based column index: 9 = col J (2021 AE), 10 = col K (2024 GE)
  return GEVS_USE_COL_K[String(acNorm || "").toLowerCase()] ? 10 : 9;
}

function getGevsHistoricalPercents(acNorm) {
  // Must match Python GEVSVE_TARGETS_J_2021 / GEVSVE_TARGETS_K_2024 exactly.
  // Chathannoor/Attingal/Kattakkada/Manjeshwaram/Poonjar → 2021 AE results (col J).
  // Kasaragod/Nemom → 2024 GE results (col K).
  // "Others" = 100 − LDF − UDF − BJP/NDA.
  var percents = {
    "chathannoor": { "LDF": 43.12, "UDF": 24.93, "BJP/NDA": 30.61 },
    "attingal":    { "LDF": 47.35, "UDF": 25.02, "BJP/NDA": 25.92 },
    "kattakkada":  { "LDF": 45.49, "UDF": 29.55, "BJP/NDA": 23.77 },
    "manjeshwaram":{ "LDF": 23.57, "UDF": 38.14, "BJP/NDA": 37.70 },
    "poonjar":     { "LDF": 41.94, "UDF": 24.76, "BJP/NDA": 29.92 },
    "kasaragod":   { "LDF": 17.67, "UDF": 49.57, "BJP/NDA": 31.76 },
    "nemom":       { "LDF": 24.59, "UDF": 28.85, "BJP/NDA": 45.18 }
  };
  var entry = percents[String(acNorm || "").toLowerCase()];
  if (!entry) return null;
  // Compute "Others" = 100 − LDF − UDF − BJP/NDA
  var oth = 100 - (entry["LDF"] + entry["UDF"] + entry["BJP/NDA"]);
  if (oth < 0) oth = 0;
  return { "LDF": entry["LDF"], "UDF": entry["UDF"], "BJP/NDA": entry["BJP/NDA"], "Others": oth };
}

function percentsToGevsFractions(pObj) {
  var total = 0;
  for (var k in pObj) if (pObj.hasOwnProperty(k)) total += pObj[k];
  if (total <= 0) return {};
  var res = {};
  for (var k2 in pObj) if (pObj.hasOwnProperty(k2)) res[k2] = pObj[k2] / total;
  return res;
}

function isGevsSevenAc(acNorm) {
  return !!getGevsHistoricalPercents(acNorm);
}

function getGevsTargetFractionForParty(acNorm, partyKey) {
  var pct = getGevsHistoricalPercents(acNorm);
  if (!pct) return null;
  var frac = percentsToGevsFractions(pct);
  if (frac.hasOwnProperty(partyKey)) return frac[partyKey];
  if (partyKey === GEVS_EMPTY_PARTY_KEY) return 0;
  return frac["Others"] || 0;
}

function computeGevsVE(ac, voteRaw, allAcVotes) {
  var acNorm = normalizeAcName(ac);
  if (!isGevsSevenAc(acNorm)) return 1;
  var partyKey = normalizeVoteForGevs(voteRaw);
  if (partyKey === GEVS_EMPTY_PARTY_KEY || partyKey === "") return 1;
  var total = allAcVotes.length;
  if (total === 0) return 1;
  var count = 0;
  for (var i = 0; i < total; i++) {
    if (normalizeVoteForGevs(allAcVotes[i]) === partyKey) count++;
  }
  if (count === 0) return 1;
  var sampleFrac = count / total;
  var targetFrac = getGevsTargetFractionForParty(acNorm, partyKey);
  if (targetFrac == null || targetFrac <= 0) return 1;
  return targetFrac / sampleFrac;
}

// ═══════════════════════════════════════════════════════════
// SHEET HEADER
// ═══════════════════════════════════════════════════════════
function ensureHeaders(sh) {
  var first = sh.getRange(1, 1).getValue();
  if (first === "Timestamp") return;
  sh.getRange(1, 1, 1, 16).setValues([[
    "Timestamp",
    "AC Name",
    "FA Name",
    "Caste Weight",
    "Caste Label",
    "Gender Weight",
    "Gender Label",
    "Age Weight",
    "Age Label",
    "Vote 2021",
    "Vote 2024 LS",
    "Vote 2026",
    "Who Will Win 2026",
    "Normalized (D×F×H)",
    "GevsVE",
    "Final Values"
  ]]);
}

// ═══════════════════════════════════════════════════════════
// FAST RECALC AFTER APPEND
// ═══════════════════════════════════════════════════════════

function fastRecalcDemographicWeightsAfterAppend(sh) {
  var lr = sh.getLastRow();
  if (lr < 2) return;
  var data = sh.getRange(2, 1, lr - 1, 16).getValues();
  // group by AC
  var acMap = {};
  for (var i = 0; i < data.length; i++) {
    var acNorm = normalizeAcName(data[i][1]); // col B
    if (!acNorm) continue;
    if (!acMap[acNorm]) acMap[acNorm] = [];
    acMap[acNorm].push({
      rowIdx: i,
      caste: data[i][4],  // col E label
      gender: data[i][6], // col G label
      age: data[i][8]     // col I label
    });
  }
  var updates = []; // [ [row1idx, [D, E, F, G, H, I, ..., N]] ]
  for (var ac in acMap) {
    if (!acMap.hasOwnProperty(ac)) continue;
    var rows = acMap[ac];
    var demo = getDemographicsMap()[ac];
    if (!demo) continue;
    for (var j = 0; j < rows.length; j++) {
      var r = rows[j];
      var cl = casteLabelForSheet(ac, r.caste, data[r.rowIdx][3]);
      var gl = genderLabelForSheet(ac, r.gender, data[r.rowIdx][5]);
      var al = ageLabelForSheet(r.age, data[r.rowIdx][7]);
      var wObj = resolveWeightsFromLabelsAndCounts(ac, cl, gl, al, rows);
      var norm = wObj.caste * wObj.gender * wObj.age;
      data[r.rowIdx][3] = wObj.caste;   // D
      data[r.rowIdx][4] = cl;           // E
      data[r.rowIdx][5] = wObj.gender;  // F
      data[r.rowIdx][6] = gl;           // G
      data[r.rowIdx][7] = wObj.age;     // H
      data[r.rowIdx][8] = al;           // I
      data[r.rowIdx][13] = norm;        // N
    }
  }
  // write back D:I and N
  var colDI = [];
  var colN = [];
  for (var k = 0; k < data.length; k++) {
    colDI.push([data[k][3], data[k][4], data[k][5], data[k][6], data[k][7], data[k][8]]);
    colN.push([data[k][13]]);
  }
  sh.getRange(2, 4, colDI.length, 6).setValues(colDI);
  sh.getRange(2, 14, colN.length, 1).setValues(colN);
}

function fastRecalcGevsveAfterAppend(sh) {
  var lr = sh.getLastRow();
  if (lr < 2) return;
  var data = sh.getRange(2, 1, lr - 1, 16).getValues();
  // group votes by AC — use correct column (J or K) per AC
  var acVotes = {};
  for (var i = 0; i < data.length; i++) {
    var acN = normalizeAcName(data[i][1]);
    if (!acN) continue;
    if (!isGevsSevenAc(acN)) continue;
    if (!acVotes[acN]) acVotes[acN] = [];
    var voteCol = gevsVoteColumnIndex(acN);
    acVotes[acN].push({ idx: i, vote: data[i][voteCol] });
  }
  var colO = [];
  var colP = [];
  for (var r = 0; r < data.length; r++) {
    var acNorm = normalizeAcName(data[r][1]);
    if (acNorm && isGevsSevenAc(acNorm) && acVotes[acNorm]) {
      var votes = acVotes[acNorm].map(function(v) { return v.vote; });
      var voteColR = gevsVoteColumnIndex(acNorm);
      var gev = computeGevsVE(acNorm, data[r][voteColR], votes);
      colO.push([gev]);
      colP.push([asFiniteNumber(data[r][13], 1) * gev]);
    } else {
      colO.push([1]);
      colP.push([asFiniteNumber(data[r][13], 1)]);
    }
  }
  sh.getRange(2, 15, colO.length, 1).setValues(colO);
  sh.getRange(2, 16, colP.length, 1).setValues(colP);
}

// ═══════════════════════════════════════════════════════════
// SINGLE-AC RECALC (used by doPost — only recalcs the one AC that changed)
// ═══════════════════════════════════════════════════════════

function fastRecalcSingleAcAfterAppend(sh, targetAc) {
  var targetAcNorm = normalizeAcName(targetAc);
  if (!targetAcNorm) return;

  var lr = sh.getLastRow();
  if (lr < 2) return;
  var data = sh.getRange(2, 1, lr - 1, 16).getValues();

  // Collect only rows belonging to this AC
  var acRows = [];   // { rowIdx, caste, gender, age }
  var acVotes = [];   // raw vote values for GevsVE
  var isGevs = isGevsSevenAc(targetAcNorm);
  var voteCol = isGevs ? gevsVoteColumnIndex(targetAcNorm) : 9;

  for (var i = 0; i < data.length; i++) {
    var rowAc = normalizeAcName(data[i][1]);
    if (rowAc !== targetAcNorm) continue;
    acRows.push({
      rowIdx: i,
      caste: data[i][4],   // col E label
      gender: data[i][6],  // col G label
      age: data[i][8]      // col I label
    });
    if (isGevs) acVotes.push(data[i][voteCol]);
  }

  if (acRows.length === 0) return;

  var demo = getDemographicsMap()[targetAcNorm];

  // Recalc D, E, F, G, H, I, N for each row in this AC
  for (var j = 0; j < acRows.length; j++) {
    var r = acRows[j];
    var idx = r.rowIdx;
    if (demo) {
      var cl = casteLabelForSheet(targetAcNorm, r.caste, data[idx][3]);
      var gl = genderLabelForSheet(targetAcNorm, r.gender, data[idx][5]);
      var al = ageLabelForSheet(r.age, data[idx][7]);
      var wObj = resolveWeightsFromLabelsAndCounts(targetAcNorm, cl, gl, al, acRows);
      data[idx][3] = wObj.caste;   // D
      data[idx][4] = cl;           // E
      data[idx][5] = wObj.gender;  // F
      data[idx][6] = gl;           // G
      data[idx][7] = wObj.age;     // H
      data[idx][8] = al;           // I
      data[idx][13] = wObj.caste * wObj.gender * wObj.age; // N
    }

    // GevsVE (O) and Final (P)
    if (isGevs && acVotes.length > 0) {
      var gev = computeGevsVE(targetAcNorm, data[idx][voteCol], acVotes);
      data[idx][14] = gev;                                        // O
      data[idx][15] = asFiniteNumber(data[idx][13], 1) * gev;     // P
    } else {
      data[idx][14] = 1;                                          // O
      data[idx][15] = asFiniteNumber(data[idx][13], 1);           // P
    }
  }

  // Write back only the changed rows using batch setValues per contiguous block
  // Group consecutive row indices to minimize API calls
  var ranges = []; // { startIdx, count }
  var sorted = acRows.map(function(r) { return r.rowIdx; }).sort(function(a, b) { return a - b; });

  var blockStart = sorted[0], blockEnd = sorted[0];
  for (var k = 1; k < sorted.length; k++) {
    if (sorted[k] === blockEnd + 1) {
      blockEnd = sorted[k];
    } else {
      ranges.push({ start: blockStart, end: blockEnd });
      blockStart = sorted[k];
      blockEnd = sorted[k];
    }
  }
  ranges.push({ start: blockStart, end: blockEnd });

  for (var b = 0; b < ranges.length; b++) {
    var s = ranges[b].start, en = ranges[b].end;
    var numRows = en - s + 1;
    var diBlock = [], nopBlock = [];
    for (var ri = s; ri <= en; ri++) {
      diBlock.push([data[ri][3], data[ri][4], data[ri][5], data[ri][6], data[ri][7], data[ri][8]]);
      nopBlock.push([data[ri][13], data[ri][14], data[ri][15]]);
    }
    sh.getRange(s + 2, 4, numRows, 6).setValues(diBlock);   // D:I
    sh.getRange(s + 2, 14, numRows, 3).setValues(nopBlock);  // N:P
  }
}

// ═══════════════════════════════════════════════════════════
// FULL SHEET RECALC
// ═══════════════════════════════════════════════════════════

function recalcAllDemographicWeightsAndNormalized() {
  clearDemographicsCache();
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME);
  if (!sh) return;
  fastRecalcDemographicWeightsAfterAppend(sh);
}

function recalcGevsveAndFinalValues() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME);
  if (!sh) return;
  fastRecalcGevsveAfterAppend(sh);
}

function recalcAllNormalizedScores() {
  clearDemographicsCache();
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME);
  if (!sh) return;
  fastRecalcDemographicWeightsAfterAppend(sh);
  fastRecalcGevsveAfterAppend(sh);
}

// ═══════════════════════════════════════════════════════════
// doPost — RECEIVES FORM SUBMISSIONS (16 columns)
// ═══════════════════════════════════════════════════════════

function doPost(e) {
  try {
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);

    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(MAIN_SHEET_NAME);
    if (!sh) { sh = ss.insertSheet(MAIN_SHEET_NAME); }
    ensureHeaders(sh);

    var ac = normalizeAcName(data.acName || data.ac_name || data.ac || "");
    var fa = data.faName || data.fa_name || "";
    var rawCaste = data.caste || "";
    var rawGender = data.gender || "";
    var rawAge = data.age || "";
    // Prefer explicit label fields sent by the form (text: "Muslim", "Male", "30-39")
    var rawCasteLabel = data.casteLabel || data.caste_label || "";
    var rawGenderLabel = data.genderLabel || data.gender_label || "";
    var rawAgeLabel = data.ageLabel || data.age_label || "";
    var vote2021 = data.vote2021 || data.vote_2021 || "";
    var vote2024 = data.vote2024 || data.vote_2024 || "";
    var vote2026 = data.vote2026 || data.vote_2026 || "";
    var whoWillWin = data.whoWillWin || data.who_will_win || "";

    // Resolve labels — prefer the explicit text label; fall back to reverse-lookup from weight
    var cl = casteLabelForSheet(ac, rawCasteLabel || rawCaste, rawCaste);
    var gl = genderLabelForSheet(ac, rawGenderLabel || rawGender, rawGender);
    var al = ageLabelForSheet(rawAgeLabel || rawAge, rawAge);

    // Initial weights = 1 (will be recalculated after append)
    var timestamp = new Date();
    var newRow = [
      timestamp,    // A - Timestamp
      ac,           // B - AC Name
      fa,           // C - FA Name
      1,            // D - Caste Weight (placeholder)
      cl,           // E - Caste Label
      1,            // F - Gender Weight (placeholder)
      gl,           // G - Gender Label
      1,            // H - Age Weight (placeholder)
      al,           // I - Age Label
      vote2021,     // J - Vote 2021
      vote2024,     // K - Vote 2024 LS
      vote2026,     // L - Vote 2026
      whoWillWin,   // M - Who Will Win 2026
      1,            // N - Normalized (placeholder)
      1,            // O - GevsVE (placeholder)
      1             // P - Final Values (placeholder)
    ];
    sh.appendRow(newRow);

    // Recalc only the affected AC (raking weights change for ALL rows in this AC when a new entry is added)
    clearDemographicsCache();
    fastRecalcSingleAcAfterAppend(sh, ac);

    clearEntriesCache();
    lock.releaseLock();

    return jsonResponse({ result: "success", message: "Entry appended and weights recalculated." });
  } catch (err) {
    return jsonResponse({ result: "error", message: err.toString() });
  }
}

// ═══════════════════════════════════════════════════════════
// doGet — API ENDPOINTS
// ═══════════════════════════════════════════════════════════

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "";

  if (action === "entries") {
    var dateStr = e.parameter.date || "";
    var rangeStr = e.parameter.range || "";
    if (rangeStr) {
      var parts = rangeStr.split(",");
      return jsonResponse(getEntriesForDateRange(parts[0], parts[1] || parts[0]));
    }
    if (dateStr) return jsonResponse(getEntriesForDate(dateStr));
    return jsonResponse(getEntriesForDate(""));
  }

  if (action === "summary") {
    var sDate = e.parameter.date || "";
    return jsonResponse(getSummaryForDate(sDate));
  }

  if (action === "dates") {
    return jsonResponse(getAvailableDates());
  }

  if (action === "recalcOP") {
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      recalcAllNormalizedScores();
    } finally {
      lock.releaseLock();
    }
    return jsonResponse({ result: "success", message: "All weights and scores recalculated." });
  }

  if (action === "ping") {
    return jsonResponse({ result: "pong", ts: new Date().toISOString() });
  }

  return jsonResponse({ result: "error", message: "Unknown action: " + action });
}

// ═══════════════════════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════════════════════

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

var _ENTRIES_CACHE = {};
function clearEntriesCache() { _ENTRIES_CACHE = {}; }

function pingWarm() {
  Logger.log("Warm ping at " + new Date().toISOString());
}

// ═══════════════════════════════════════════════════════════
// DATE PARSING
// ═══════════════════════════════════════════════════════════

function parseTimestamp(ts) {
  if (ts instanceof Date) return ts;
  if (typeof ts === "number") return new Date(ts);
  var s = String(ts).trim();
  // dd/mm/yyyy hh:mm:ss
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2}):(\d{2})$/);
  if (m) return new Date(+m[3], +m[2]-1, +m[1], +m[4], +m[5], +m[6]);
  // ISO or other
  var d = new Date(s);
  if (!isNaN(d.getTime())) return d;
  return null;
}

function dateToYMD(d) {
  return d.getFullYear() + "-" + pad2(d.getMonth()+1) + "-" + pad2(d.getDate());
}

function parseDateParam(str) {
  if (!str) return null;
  var parts = String(str).split(/[-\/]/);
  if (parts.length === 3) {
    var y = +parts[0], m = +parts[1] - 1, d = +parts[2];
    if (parts[0].length <= 2) { d = +parts[0]; m = +parts[1] - 1; y = +parts[2]; }
    return new Date(y, m, d);
  }
  return null;
}

// ═══════════════════════════════════════════════════════════
// ENTRIES API
// ═══════════════════════════════════════════════════════════

function mapRowsToEntries(data) {
  var entries = [];
  for (var i = 0; i < data.length; i++) {
    var r = data[i];
    entries.push({
      timestamp: r[0],
      acName: r[1],
      faName: r[2],
      casteWeight: r[3],
      casteLabel: r[4],
      genderWeight: r[5],
      genderLabel: r[6],
      ageWeight: r[7],
      ageLabel: r[8],
      vote2021: r[9],
      vote2024: r[10],
      vote2026: r[11],
      whoWillWin: r[12],
      normalized: r[13],
      gevsve: r[14],
      finalValue: r[15]
    });
  }
  return entries;
}

function getEntriesForDate(dateStr) {
  var cacheKey = "entries_" + dateStr;
  if (_ENTRIES_CACHE[cacheKey]) {
    var cached = _ENTRIES_CACHE[cacheKey];
    if ((new Date().getTime() - cached.ts) < ENTRIES_CACHE_TTL_SEC * 1000) return cached.data;
  }
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME);
  if (!sh) return [];
  var lr = sh.getLastRow();
  if (lr < 2) return [];
  var data = sh.getRange(2, 1, lr - 1, 16).getValues();
  var filtered = data;
  if (dateStr) {
    var target = parseDateParam(dateStr);
    if (target) {
      var ty = target.getFullYear(), tm = target.getMonth(), td = target.getDate();
      filtered = data.filter(function(r) {
        var ts = parseTimestamp(r[0]);
        if (!ts) return false;
        return ts.getFullYear() === ty && ts.getMonth() === tm && ts.getDate() === td;
      });
    }
  }
  var result = mapRowsToEntries(filtered);
  var json = JSON.stringify(result);
  if (json.length < ENTRIES_CACHE_MAX_CHARS) {
    _ENTRIES_CACHE[cacheKey] = { ts: new Date().getTime(), data: result };
  }
  return result;
}

function getEntriesForDateRange(startStr, endStr) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME);
  if (!sh) return [];
  var lr = sh.getLastRow();
  if (lr < 2) return [];
  var data = sh.getRange(2, 1, lr - 1, 16).getValues();
  var startD = parseDateParam(startStr);
  var endD = parseDateParam(endStr);
  if (!startD || !endD) return mapRowsToEntries(data);
  var s = startD.getTime(), en = endD.getTime() + 86400000;
  var filtered = data.filter(function(r) {
    var ts = parseTimestamp(r[0]);
    if (!ts) return false;
    var t = ts.getTime();
    return t >= s && t < en;
  });
  return mapRowsToEntries(filtered);
}

function getSummaryForDate(dateStr) {
  var entries = getEntriesForDate(dateStr);
  var byAc = {};
  for (var i = 0; i < entries.length; i++) {
    var e = entries[i];
    var ac = normalizeAcName(e.acName);
    if (!ac) continue;
    if (!byAc[ac]) byAc[ac] = { count: 0, parties: {} };
    byAc[ac].count++;
    var party2026 = normalizeParty(e.vote2026);
    if (party2026) {
      if (!byAc[ac].parties[party2026]) byAc[ac].parties[party2026] = { raw: 0, weighted: 0 };
      byAc[ac].parties[party2026].raw++;
      byAc[ac].parties[party2026].weighted += asFiniteNumber(e.finalValue, 1);
    }
  }
  return { date: dateStr || "all", totalEntries: entries.length, byAc: byAc };
}

function getAvailableDates() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME);
  if (!sh) return [];
  var lr = sh.getLastRow();
  if (lr < 2) return [];
  var col = sh.getRange(2, 1, lr - 1, 1).getValues();
  var seen = {};
  var dates = [];
  for (var i = 0; i < col.length; i++) {
    var ts = parseTimestamp(col[i][0]);
    if (!ts) continue;
    var ymd = dateToYMD(ts);
    if (!seen[ymd]) { seen[ymd] = true; dates.push(ymd); }
  }
  dates.sort();
  return dates;
}

// ═══════════════════════════════════════════════════════════
// FIX TEXT ROWS (legacy numeric → label backfill)
// ═══════════════════════════════════════════════════════════

function fixTextRows() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME);
  if (!sh) return;
  var lr = sh.getLastRow();
  if (lr < 2) return;
  var data = sh.getRange(2, 1, lr - 1, 16).getValues();
  var changed = false;
  for (var i = 0; i < data.length; i++) {
    var ac = normalizeAcName(data[i][1]);
    // col E – caste label
    if (!isTextLabel(data[i][4]) || !isValidCasteLabel(data[i][4])) {
      var cl = getCasteLabelFromWeight(ac, data[i][3]);
      if (cl !== data[i][4]) { data[i][4] = cl; changed = true; }
    }
    // col G – gender label
    if (!isTextLabel(data[i][6]) || !isValidGenderLabel(data[i][6])) {
      var gl = getGenderLabelFromWeight(ac, data[i][5]);
      if (gl !== data[i][6]) { data[i][6] = gl; changed = true; }
    }
    // col I – age label
    if (!isTextLabel(data[i][8]) || !isValidAgeLabel(data[i][8])) {
      var al = getAgeLabelFromWeight(data[i][7]);
      if (al !== data[i][8]) { data[i][8] = al; changed = true; }
    }
  }
  if (changed) {
    var labels = data.map(function(r) { return [r[4], r[6], r[8]]; });
    // Write E, G, I
    sh.getRange(2, 5, labels.length, 1).setValues(labels.map(function(l) { return [l[0]]; }));
    sh.getRange(2, 7, labels.length, 1).setValues(labels.map(function(l) { return [l[1]]; }));
    sh.getRange(2, 9, labels.length, 1).setValues(labels.map(function(l) { return [l[2]]; }));
  }
  Logger.log("fixTextRows done; changed=" + changed);
}

// ═══════════════════════════════════════════════════════════
// DAILY REPORT
// ═══════════════════════════════════════════════════════════

function buildWeightedReportRows(data) {
  var acMap = {};
  for (var i = 0; i < data.length; i++) {
    var acNorm = normalizeAcName(data[i][1]);
    if (!acNorm) continue;
    if (!acMap[acNorm]) acMap[acNorm] = { raw: {}, total: 0 };
    acMap[acNorm].total++;
    var party = normalizeParty(data[i][11]); // col L (Vote 2026)
    if (!party) party = "No Response";
    if (!acMap[acNorm].raw[party]) acMap[acNorm].raw[party] = { count: 0, weightedSum: 0 };
    acMap[acNorm].raw[party].count++;
    acMap[acNorm].raw[party].weightedSum += asFiniteNumber(data[i][15], 1); // col P
  }

  var rows = [];
  var acs = Object.keys(acMap).sort();
  for (var a = 0; a < acs.length; a++) {
    var ac = acs[a];
    var info = acMap[ac];
    var totalW = 0;
    for (var p in info.raw) if (info.raw.hasOwnProperty(p)) totalW += info.raw[p].weightedSum;
    if (totalW <= 0) totalW = 1;
    var parties = Object.keys(info.raw).sort();
    for (var pi = 0; pi < parties.length; pi++) {
      var pName = parties[pi];
      var pData = info.raw[pName];
      rows.push([
        ac,
        pName,
        pData.count,
        info.total,
        Math.round(pData.weightedSum / totalW * 10000) / 100
      ]);
    }
  }
  return rows;
}

function writeDailyReportSheet(reportDate, rows) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = "Report_" + reportDate;
  var sh = ss.getSheetByName(sheetName);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(sheetName);
  sh.getRange(1, 1, 1, 5).setValues([["AC", "Party", "Raw Count", "Total Responses", "Weighted %"]]);
  if (rows.length > 0) {
    sh.getRange(2, 1, rows.length, 5).setValues(rows);
  }
  return sheetName;
}

function generateDailyReport(dateStr) {
  if (!dateStr) dateStr = dateToYMD(new Date());
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME);
  if (!sh) return "No main sheet";
  var lr = sh.getLastRow();
  if (lr < 2) return "No data";
  var data = sh.getRange(2, 1, lr - 1, 16).getValues();
  var target = parseDateParam(dateStr);
  var filtered = data;
  if (target) {
    var ty = target.getFullYear(), tm = target.getMonth(), td = target.getDate();
    filtered = data.filter(function(r) {
      var ts = parseTimestamp(r[0]);
      if (!ts) return false;
      return ts.getFullYear() === ty && ts.getMonth() === tm && ts.getDate() === td;
    });
  }
  var rows = buildWeightedReportRows(filtered);
  var name = writeDailyReportSheet(dateStr, rows);
  return "Report written to sheet: " + name + " (" + filtered.length + " entries, " + rows.length + " rows)";
}
