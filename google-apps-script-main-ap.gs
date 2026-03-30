/**
 * Kerala Survey 2026 — DYNAMIC MATCHING EXCEL FORMULA SCRIPT
 *
 * Implements the "Shared Power" model: when a new entry is added, 
 * the GevsVE (Column O) updates for all existing rows in that group.
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

var CASTE_LIST = [
  "Nair", 
  "Ezhava", 
  "Muslim", 
  "Christian", 
  "SC/ST", 
  "Others"
];

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
  "Wandoor": {"male": 49.13, "female": 50.87},
};

/**
 * Allowed FA display names per AC (matches survey form). Optional reference; not enforced server-side.
 */
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
  "Kanjirappally": ["Gokul PG", "Harikrishnan AB"]
};

var _DEMOGRAPHICS_CACHE = null;

function pad2(n) {
  n = parseInt(n, 10);
  if (n < 10) {
    return "0" + n;
  }
  return String(n);
}

function asFiniteNumber(val, fallback) {
  if (val === null || val === undefined || val === "") {
    return fallback;
  }
  var n;
  if (typeof val === "number") {
    n = val;
  } else {
    n = parseFloat(String(val).replace(/,/g, "").replace(/%/g, ""));
  }
  if (isFinite(n)) {
    return n;
  }
  return fallback;
}

function isTextLabel(val) {
  var s = String(val === null || val === undefined ? "" : val).trim();
  if (s === "") {
    return false;
  }
  return isNaN(parseFloat(s));
}

function normalizeAgeLabel(label) {
  return String(label || "").trim().replace(/[–—]/g, "-");
}

function canonicalCasteKey(casteStr) {
  var s = String(casteStr || "").trim();
  if (!s) {
    return s;
  }
  var low = s.toLowerCase().replace(/\s+/g, " ");
  if (low === "sc/st" || low === "scst" || low === "sc-st" || low === "sc / st") {
    return "SC/ST";
  }
  for (var i = 0; i < CASTE_LIST.length; i++) {
    if (CASTE_LIST[i].toLowerCase() === low) {
      return CASTE_LIST[i];
    }
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
  if (g === "male") { return "Male"; }
  if (g === "female") { return "Female"; }
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

  var dem = getDemographicsMap();
  if (dem[s]) {
    return s;
  }

  var keys = Object.keys(dem);
  for (var i = 0; i < keys.length; i++) {
    if (keys[i].toLowerCase() === k) {
      return keys[i];
    }
  }
  return s;
}

function normalizeParty(p) {
  var s = String(p || "").trim().toUpperCase().replace(/\s+/g, "");
  if (!s) {
    return "Others";
  }
  if (s === "LDF") {
    return "LDF";
  }
  if (s === "UDF") {
    return "UDF";
  }
  if (s === "BJP/NDA" || s === "BJP-NDA" || s === "BJPNDA" || s === "BJP" || s === "NDA") {
    return "BJP/NDA";
  }
  return "Others";
}

function voteForSheet(v) {
  var p = normalizeParty(v);
  if (p === "UDF" || p === "LDF" || p === "BJP/NDA") {
    return p;
  }
  return String(v === null || v === undefined ? "" : v).trim();
}

function ensureHeaders(sheet) {
  var a1 = String(sheet.getRange(1, 1).getValue() || "").trim();
  if (a1 === "Timestamp") {
    return;
  }

  var headers = [[
    "Timestamp", "Assembly Constituency", "FA Name",
    "Caste Weight", "Caste Label",
    "Gender Weight", "Gender Label",
    "Age Weight", "Age Label",
    "Vote in 2021 AE", "Vote in 2024 GE", "Vote in 2026 AE",
    "Who Will Win", "Who Will Win Normalized",
    "GevsVE", "Final Values"
  ]];
  sheet.getRange(1, 1, 1, 16).setValues(headers);
}

function readDemographicsFromSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(DEMOGRAPHICS_SHEET_NAME);
  if (!sh) {
    return null;
  }

  var lr = sh.getLastRow();
  var lc = sh.getLastColumn();
  if (lr < 2 || lc < 9) {
    return null;
  }

  var data = sh.getRange(1, 1, lr, lc).getValues();
  var head = data[0].map(function(h) { 
    return String(h).trim(); 
  });

  function idx(name) { 
    return head.indexOf(name); 
  }

  var iAc = idx("AC Name");
  var iMale = idx("Male Percentage");
  var iFemale = idx("Female Percentage");
  var iMuslim = idx("%Muslims");
  var iChristian = idx("%Christians");
  var iNair = idx("%Nairs");
  var iEzhava = idx("%Ezhavas");
  var iOthers = idx("%Others");
  var iScst = idx("%SC/ST");

  var missing = [iAc, iMale, iFemale, iMuslim, iChristian, iNair, iEzhava, iOthers, iScst].some(function(x) { 
    return x < 0; 
  });

  if (missing) {
    throw new Error("Demographics sheet headers missing/changed.");
  }

  var map = {};
  for (var r = 1; r < data.length; r++) {
    var ac = normalizeAcName(data[r][iAc]);
    if (!ac) {
      continue;
    }
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
  if (_DEMOGRAPHICS_CACHE) {
    return _DEMOGRAPHICS_CACHE;
  }

  if (USE_DEMOGRAPHICS_SHEET) {
    try {
      var fromSheet = readDemographicsFromSheet();
      if (fromSheet && Object.keys(fromSheet).length > 0) {
        _DEMOGRAPHICS_CACHE = fromSheet;
        return _DEMOGRAPHICS_CACHE;
      }
    } catch (err) {
      Logger.log("Failed to read demographics: " + err);
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
    if (AGE_WEIGHTS.hasOwnProperty(label)) {
      var diff = Math.abs(AGE_WEIGHTS[label] - w);
      if (diff < bestDiff) { 
        bestDiff = diff; 
        best = label; 
      }
    }
  }
  return best;
}

function getAgeLabelFromAny(ageVal) {
  if (isTextLabel(ageVal)) {
    var txt = normalizeAgeLabel(ageVal);
    if (AGE_WEIGHTS.hasOwnProperty(txt)) {
      return txt;
    }
    return "Unknown";
  }

  var n = asFiniteNumber(ageVal, NaN);
  if (!isFinite(n)) {
    return "Unknown";
  }

  if (n > 0 && n < 1.2) {
    return getClosestAgeLabelFromNormalizedWeight(n);
  }
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
  if (AGE_WEIGHTS[label]) {
    return AGE_WEIGHTS[label];
  }
  return 0;
}

function resolveWeights(ac, casteVal, genderVal, ageVal) {
  var acName = normalizeAcName(ac);
  var dem = getDemographicsMap();
  var acData = dem[acName];

  var casteW = 0;
  var genderW = 0;
  var ageW = 0;

  if (isTextLabel(casteVal)) {
    var ck = canonicalCasteKey(casteVal);
    if (acData && acData[ck]) {
      casteW = acData[ck] / 100;
    }
  } else {
    casteW = asFiniteNumber(casteVal, 0);
    if (casteW > 1.5) {
      casteW = casteW / 100;
    }
  }

  if (isTextLabel(genderVal)) {
    var g = String(genderVal).trim().toLowerCase();
    if (acData) {
      if (g === "male") {
        genderW = acData.male / 100;
      } else {
        genderW = acData.female / 100;
      }
    }
  } else {
    genderW = asFiniteNumber(genderVal, 0);
    if (genderW > 1.5) {
      genderW = genderW / 100;
    }
  }

  ageW = getAgeWeightFromAny(ageVal);
  var norm = casteW * genderW * ageW;

  return { 
    casteW: casteW, 
    genderW: genderW, 
    ageW: ageW, 
    norm: norm 
  };
}

function casteLabelForSheet(ac, formCaste, casteW) {
  if (formCaste != null && isValidCasteLabel(formCaste)) {
    return canonicalCasteKey(formCaste);
  }
  return getCasteLabelFromWeight(ac, casteW);
}

function genderLabelForSheet(ac, formGender, genderW) {
  if (formGender != null && isValidGenderLabel(formGender)) {
    var d = displayGenderLabel(formGender);
    if (d) {
      return d;
    }
  }
  return getGenderLabelFromWeight(ac, genderW);
}

function ageLabelForSheet(formAge, ageW) {
  if (formAge != null && isValidAgeLabel(formAge)) {
    return normalizeAgeLabel(formAge);
  }
  return getAgeLabelFromWeight(ageW);
}

function getCasteLabelFromWeight(ac, casteWeight) {
  if (isTextLabel(casteWeight)) {
    var t = canonicalCasteKey(casteWeight);
    if (isValidCasteLabel(t)) {
      return t;
    }
    return "Others";
  }

  var acData = getDemographicsMap()[normalizeAcName(ac)];
  if (!acData) {
    return "Unknown";
  }

  var w = asFiniteNumber(casteWeight, NaN);
  if (!isFinite(w) || w === 0) {
    return "Unknown";
  }
  if (w > 1.5) {
    w = w / 100;
  }

  var best = "Others";
  var bestDiff = Infinity;
  for (var i = 0; i < CASTE_LIST.length; i++) {
    var diff = Math.abs((acData[CASTE_LIST[i]] || 0) / 100 - w);
    if (diff < bestDiff) { 
      bestDiff = diff; 
      best = CASTE_LIST[i]; 
    }
  }
  return best;
}

function getGenderLabelFromWeight(ac, genderWeight) {
  if (isTextLabel(genderWeight)) {
    var t = String(genderWeight).trim();
    if (isValidGenderLabel(t)) {
      return displayGenderLabel(t);
    }
    return "Unknown";
  }

  var acData = getDemographicsMap()[normalizeAcName(ac)];
  if (!acData) {
    return "Unknown";
  }

  var w = asFiniteNumber(genderWeight, NaN);
  if (!isFinite(w) || w === 0) {
    return "Unknown";
  }
  if (w > 1.5) {
    w = w / 100;
  }

  var mD = Math.abs((acData.male / 100) - w);
  var fD = Math.abs((acData.female / 100) - w);
  if (mD < fD) {
    return "Male";
  }
  if (fD < mD) {
    return "Female";
  }
  if (acData.female >= acData.male) {
    return "Female";
  }
  return "Male";
}

function getAgeLabelFromWeight(ageWeight) {
  return getAgeLabelFromAny(ageWeight);
}


/* =========================================================================
   PERFECT EXCEL REPLICA 
   This perfectly recreates your exact IF/OR/SWITCH formula logic line by line.
   ========================================================================= */

function executeExcelFormulaForO(acName, vote2021, vote2024, countColK_UDF, countColK_LDF, countColK_BJP, countColJ_UDF, countColJ_LDF, countColJ_BJP) {
  // Translate standard variables from Row directly into formula equivalents
  var B2 = normalizeAcName(acName);       // Column B
  var J2 = normalizeParty(vote2021);      // Column J
  var K2 = normalizeParty(vote2024);      // Column K
  
  // IF(OR(B2="Kasaragod", B2="Nemom")
  if (B2 === "Kasaragod" || B2 === "Nemom") {
    
    // SWITCH(K2...
    if (K2 === "UDF") {
      // SWITCH(B2, "Kasaragod", 49.57, "Nemom", 28.85) / COUNTIFS($B:$B, B2, $K:$K, "UDF")
      var base = (B2 === "Kasaragod") ? 49.57 : 28.85;
      return base / Math.max(1, countColK_UDF);
    } 
    else if (K2 === "LDF") {
      // SWITCH(B2, "Kasaragod", 17.67, "Nemom", 24.59) / COUNTIFS($B:$B, B2, $K:$K, "LDF")
      var base = (B2 === "Kasaragod") ? 17.67 : 24.59;
      return base / Math.max(1, countColK_LDF);
    } 
    else if (K2 === "BJP/NDA") {
      // SWITCH(B2, "Kasaragod", 31.76, "Nemom", 45.18) / COUNTIFS($B:$B, B2, $K:$K, "BJP/NDA")
      var base = (B2 === "Kasaragod") ? 31.76 : 45.18;
      return base / Math.max(1, countColK_BJP);
    } 
    else {
      // 1)
      return 1;
    }
  } 
  
  // IF(OR(B2="Chathannoor", B2="Attingal", B2="Kattakkada", B2="Manjeshwaram", B2="Poonjar")
  else if (B2 === "Chathannoor" || B2 === "Attingal" || B2 === "Kattakkada" || B2 === "Manjeshwaram" || B2 === "Poonjar") {
    
    // SWITCH(J2...
    if (J2 === "UDF") {
      // SWITCH(B2, "Chathannoor", 24.93, "Attingal", 25.02, "Kattakkada", 29.55, "Manjeshwaram", 38.14, "Poonjar", 24.76) / COUNTIFS($B:$B, B2, $J:$J, "UDF")
      var baseUDF = 1;
      if (B2 === "Chathannoor") baseUDF = 24.93;
      if (B2 === "Attingal") baseUDF = 25.02;
      if (B2 === "Kattakkada") baseUDF = 29.55;
      if (B2 === "Manjeshwaram") baseUDF = 38.14;
      if (B2 === "Poonjar") baseUDF = 24.76;
      return baseUDF / Math.max(1, countColJ_UDF);
    }
    else if (J2 === "LDF") {
      // SWITCH(...) / COUNTIFS($B:$B, B2, $J:$J, "LDF")
      var baseLDF = 1;
      if (B2 === "Chathannoor") baseLDF = 43.12;
      if (B2 === "Attingal") baseLDF = 47.35;
      if (B2 === "Kattakkada") baseLDF = 45.49;
      if (B2 === "Manjeshwaram") baseLDF = 23.57;
      if (B2 === "Poonjar") baseLDF = 41.94;
      return baseLDF / Math.max(1, countColJ_LDF);
    }
    else if (J2 === "BJP/NDA") {
      // SWITCH(...) / COUNTIFS($B:$B, B2, $J:$J, "BJP/NDA")
      var baseBJP = 1;
      if (B2 === "Chathannoor") baseBJP = 30.61;
      if (B2 === "Attingal") baseBJP = 25.92;
      if (B2 === "Kattakkada") baseBJP = 23.77;
      if (B2 === "Manjeshwaram") baseBJP = 37.7;
      if (B2 === "Poonjar") baseBJP = 29.92;
      return baseBJP / Math.max(1, countColJ_BJP);
    }
    else {
      // 1)
      return 1;
    }
  } 
  
  // 1)), 1) -> IFERROR Fallback / Logic Fallback
  else {
    return 1;
  }
}

/** 
 * STATIC ENTRY ENGINE: Calculates only for the new row 
 */
function fastRecalcGevsveAfterAppend(sheet, lastRow) {
  if (lastRow < 2) return;
  
  var numRows = lastRow - 1;
  // Performance Optimization: Read only once, and grab only the core columns (1=A to 16=P)
  var allData = sheet.getRange(2, 1, numRows, 16).getValues();

  // Identify the new row's group (the last item in our array)
  var lastIdx = numRows - 1;
  var acNew = normalizeAcName(allData[lastIdx][1]);
  var jNew = normalizeParty(allData[lastIdx][9]);
  var kNew = normalizeParty(allData[lastIdx][10]);
  var partyNew = getGroupPartyForAc(acNew, jNew, kNew);

  // 1. Calculate the TOTAL count for this group across history
  var matchingIndices = [];
  for (var i = 0; i < numRows; i++) {
    var acIter = normalizeAcName(allData[i][1]);
    if (acIter === acNew) {
      var pIter = getGroupPartyForAc(acIter, normalizeParty(allData[i][9]), normalizeParty(allData[i][10]));
      if (pIter === partyNew) {
        matchingIndices.push(i);
      }
    }
  }

  var totalCount = matchingIndices.length;
  if (totalCount === 0) return;

  // 2. Map the count to the Excel Formula Helper
  var counts = { kUDF: 0, kLDF: 0, kBJP: 0, jUDF: 0, jLDF: 0, jBJP: 0 };
  if (acNew === "Kasaragod" || acNew === "Nemom") {
    if (partyNew === "UDF") counts.kUDF = totalCount;
    else if (partyNew === "LDF") counts.kLDF = totalCount;
    else if (partyNew === "BJP/NDA") counts.kBJP = totalCount;
  } else {
    if (partyNew === "UDF") counts.jUDF = totalCount;
    else if (partyNew === "LDF") counts.jLDF = totalCount;
    else if (partyNew === "BJP/NDA") counts.jBJP = totalCount;
  }

  var updatedO = executeExcelFormulaForO(acNew, jNew, kNew, counts.kUDF, counts.kLDF, counts.kBJP, counts.jUDF, counts.jLDF, counts.jBJP);

  // 3. Efficient Batch Write: Only write to the Columns O and P (indices 14 and 15)
  // Instead of rewriting the whole sheet, we just prepare the update array for columns 15-16
  var updateRange = sheet.getRange(2, 15, numRows, 2);
  var colOPValues = updateRange.getValues(); // Initial values for O and P

  for (var m = 0; m < matchingIndices.length; m++) {
    var idx = matchingIndices[m];
    var nVal = asFiniteNumber(allData[idx][13], 0);
    colOPValues[idx][0] = updatedO;
    colOPValues[idx][1] = updatedO * nVal;
  }

  updateRange.setValues(colOPValues);
  updateRange.setNumberFormat("0.0000000000");
}

/** Helper to determine the group party based on AC */
function getGroupPartyForAc(ac, jParty, kParty) {
  if (ac === "Kasaragod" || ac === "Nemom") return kParty;
  if (ac === "Chathannoor" || ac === "Attingal" || ac === "Kattakkada" || ac === "Manjeshwaram" || ac === "Poonjar") return jParty;
  return "Others";
}

/** 
 * FULL SHEET STATIC RECALC
 * Also respects EXACT Excel logic if triggered manually over the entire sheet.
 */
function recalcGevsveAndFinalValues() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME) || ss.getSheets()[0];
  var lr = sheet.getLastRow();
  if (lr < 2) return;

  ensureHeaders(sheet);

  var numRows = lr - 1;
  var data = sheet.getRange(2, 1, numRows, 16).getValues();
  
  // 1. First Pass: Calculate Global Totals for all groups
  var groupTotals = {};
  for (var i = 0; i < numRows; i++) {
    var ac = normalizeAcName(data[i][1]);
    var p = getGroupPartyForAc(ac, normalizeParty(data[i][9]), normalizeParty(data[i][10]));
    var key = ac + "||" + p;
    groupTotals[key] = (groupTotals[key] || 0) + 1;
  }

  // 2. Second Pass: Apply totals to every row
  var out = [];
  for (var j = 0; j < numRows; j++) {
    var acJ = normalizeAcName(data[j][1]);
    var jVal = normalizeParty(data[j][9]);
    var kVal = normalizeParty(data[j][10]);
    var nVal = asFiniteNumber(data[j][13], 0);
    var pJ = getGroupPartyForAc(acJ, jVal, kVal);
    
    var keyJ = acJ + "||" + pJ;
    var total = groupTotals[keyJ] || 1;

    // Direct counts for the formula helper
    var c = { kUDF: 0, kLDF: 0, kBJP: 0, jUDF: 0, jLDF: 0, jBJP: 0 };
    if (acJ === "Kasaragod" || acJ === "Nemom") {
       if (pJ === "UDF") c.kUDF = total;
       else if (pJ === "LDF") c.kLDF = total;
       else if (pJ === "BJP/NDA") c.kBJP = total;
    } else {
       if (pJ === "UDF") c.jUDF = total;
       else if (pJ === "LDF") c.jLDF = total;
       else if (pJ === "BJP/NDA") c.jBJP = total;
    }

    var oVal = executeExcelFormulaForO(acJ, jVal, kVal, c.kUDF, c.kLDF, c.kBJP, c.jUDF, c.jLDF, c.jBJP);
    out.push([oVal, oVal * nVal]);
  }

  sheet.getRange(2, 15, numRows, 2).setValues(out);
  sheet.getRange(2, 15, numRows, 2).setNumberFormat("0.0000000000");
  SpreadsheetApp.flush();
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
      w.casteW, 
      casteLabel,
      w.genderW, 
      genderLabel,
      w.ageW, 
      ageLabel,
      voteForSheet(data.vote2021),
      voteForSheet(data.vote2024),
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

    try {
      fastRecalcGevsveAfterAppend(sheet, lr);
    } catch (fastErr) {
      Logger.log("fastRecalcGevsveAfterAppend failed: " + fastErr);
    }

    return jsonResponse({ status: "success", opRecalculated: true });
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
    if (action === "refreshDemographics") {
      clearDemographicsCache();
      return jsonResponse({ status: "ok" });
    }
    if (action === "recalcOP") {
      recalcGevsveAndFinalValues();
      return jsonResponse({ status: "ok", message: "O/P recalculated statically" });
    }
    if (action === "ping") {
      pingWarm();
      return jsonResponse({ status: "ok" });
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

function pingWarm() {
  try {
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME);
  } catch (err) {
    // do nothing
  }
}

function cellToYmdKolkata(cell) {
  if (cell instanceof Date) {
    if (isNaN(cell.getTime())) {
      return null;
    }
    return Utilities.formatDate(cell, "Asia/Kolkata", "yyyy-MM-dd");
  }

  var s = String(cell);

  var m = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (m) {
    var d = parseInt(m[1], 10);
    var mo = parseInt(m[2], 10);
    var y = parseInt(m[3], 10);
    if (y < 100) {
      y += 2000;
    }
    return y + "-" + pad2(mo) + "-" + pad2(d);
  }

  m = s.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (m) {
    return m[1] + "-" + m[2] + "-" + m[3];
  }

  var monthMap = { Jan:1, Feb:2, Mar:3, Apr:4, May:5, Jun:6, Jul:7, Aug:8, Sep:9, Oct:10, Nov:11, Dec:12 };
  var mon = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})/);
  if (mon) {
    var key = mon[2].charAt(0).toUpperCase() + mon[2].slice(1).toLowerCase();
    var moNum = monthMap[key];
    if (moNum) {
      return mon[3] + "-" + pad2(moNum) + "-" + pad2(mon[1]);
    }
  }
  return null;
}

function parseDateParamToYmd(dateStr) {
  if (!dateStr) return null;
  var s = String(dateStr).trim();

  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) {
    return m[1] + "-" + m[2] + "-" + m[3];
  }

  var monthMap = { Jan:1, Feb:2, Mar:3, Apr:4, May:5, Jun:6, Jul:7, Aug:8, Sep:9, Oct:10, Nov:11, Dec:12 };
  var mon = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/);
  if (mon) {
    var key = mon[2].charAt(0).toUpperCase() + mon[2].slice(1).toLowerCase();
    var moNum = monthMap[key];
    if (moNum) {
      return mon[3] + "-" + pad2(moNum) + "-" + pad2(mon[1]);
    }
  }

  var sl = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (sl) {
    return sl[3] + "-" + pad2(sl[2]) + "-" + pad2(sl[1]);
  }

  return null;
}

function timestampPatternsForDateKey(dateStr) {
  var tz = "Asia/Kolkata";
  var patterns = [String(dateStr)];
  var monthMap = { Jan:1, Feb:2, Mar:3, Apr:4, May:5, Jun:6, Jul:7, Aug:8, Sep:9, Oct:10, Nov:11, Dec:12 };

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
    if (ts.indexOf(patterns[i]) !== -1) {
      return true;
    }
  }
  return false;
}

function rowMatchesDateFilter(row, dateStr) {
  if (!dateStr) return true;
  var targetYmd = parseDateParamToYmd(dateStr);
  if (targetYmd) {
    var rowYmd = cellToYmdKolkata(row[0]);
    if (rowYmd) {
      return rowYmd === targetYmd;
    }
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
    headers.forEach(function(h, i) { 
      obj[h] = row[i]; 
    });
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
  var filtered = data;
  if (dateStr) {
    filtered = data.filter(function(row) { 
      return rowMatchesDateFilter(row, dateStr); 
    });
  }
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
    if (rowYmd) {
      return rowYmd >= fromYmd && rowYmd <= toYmd;
    }
    if (fromYmd === toYmd) {
      return rowMatchesDateFilter(row, fromYmd);
    }
    return false;
  });

  var entries = mapRowsToEntries(filtered);
  return { entries: entries, total: entries.length };
}

function getSummaryForDate(tabName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    return { rows: [], tabName: tabName, found: false };
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { rows: [], tabName: tabName, found: true };
  }

  var numRows = lastRow - 1;
  var data = sheet.getRange(2, 1, numRows, 7).getValues();
  var rows = data
    .filter(function(row) { 
      return String(row[0]).trim() !== "" && String(row[0]).indexOf("Report") === -1; 
    })
    .map(function(row) {
      return {
        ac: row[0], 
        totalEntries: row[1],
        ldf: row[2], 
        udf: row[3], 
        bjp: row[4], 
        others: row[5], 
        winner: row[6]
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
    if (n !== MAIN_SHEET_NAME && n !== DEMOGRAPHICS_SHEET_NAME) {
      dates.push(n);
    }
  }
  return { dates: dates };
}

function pickCasteInputForResolve(row) {
  if (isValidCasteLabel(row[4])) {
    return canonicalCasteKey(row[4]);
  }
  return row[3];
}

function pickGenderInputForResolve(row) {
  if (isValidGenderLabel(row[6])) {
    return row[6];
  }
  return row[5];
}

function pickAgeInputForResolve(row) {
  if (isValidAgeLabel(row[8])) {
    return normalizeAgeLabel(row[8]);
  }
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
      for (var p = 0; p < parties.length; p++) {
        acMap[ac].parties[parties[p]] = { sum: 0, count: 0 };
      }
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
      if (rec.parties[pt].sum > maxSum) { 
        maxSum = rec.parties[pt].sum; 
        winner = pt; 
      }
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
  if (existing) {
    ss.deleteSheet(existing);
  }
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

  for (var c = 1; c <= header.length; c++) {
    sheet.autoResizeColumn(c);
  }

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
