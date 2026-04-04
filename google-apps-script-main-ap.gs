/**
 * Kerala Survey 2026 — survey weights aligned with scripts/weight_ground_input.py
 *
 * Caste / Gender / Age: raking weight = population target ÷ sample share within each AC
 *   (targets from FINAL GENDER CASTE sheet + AGE_WEIGHTS).
 * GevsVE (column O): seven ACs only — historical vote share as fraction ÷ (party count / n_ac);
 *   Kasaragod & Nemom use Vote 2024 GE (K); others use Vote 2021 AE (J). Others = 1 − LDF − UDF − BJP.
 * On append: all rows in that AC get new D–H and N; then O and P (P = O × N).
 */

var MAIN_SHEET_NAME = "Sheet1";

/** Short-lived cache for entries API (reduces repeat dashboard hits; TTL keeps new submissions visible quickly). Max ~100KB per CacheService entry. */
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
  "Beypur": {"male": 48.97, "female": 51.03},
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
  "Adoor": ["KG Gokulkrishnan Nair", "Vyshakh V"],
  "Ambalapuzha": ["FA 1", "FA 2"],
  "Aranmula": ["Saneesh", "Syam Chandran", "Sarath Mohan"],
  "Aroor": ["Akshay Das", "Abhijith Anilkumar", "Ananthu Gokul"],
  "Attingal": ["Anand KS", "Vignesh lal", "SREEHARI .S"],
  "Beypur": ["Amal Manoj", "Abhinav KP"],
  "Chalakkudy": ["FA 1", "FA 2"],
  "Changanassery": ["FA 1", "FA 2"],
  "Chathannoor": ["Aromal A", "Sajikumar S", "Aravind AS"],
  "Chelakkara": ["Jithesh Nair"],
  "Chengannur": ["Aswin S", "Ananthu R Pillai"],
  "Cherthala": ["FA 1", "FA 2"],
  "Chirayankeezhu": ["FA 1", "FA 2"],
  "Devikulam": ["Ananthan"],
  "Ettumanoor": ["Sreehari Babu"],
  "Guruvayoor": ["Sreehari", "Ranjith"],
  "Harippad": ["Vaishakh V", "Akul R Kamath"],
  "Idukki": ["FA 1", "FA 2"],
  "Irinjalakkuda": ["Akshay Sasidharan", "Athul", "Sanjay Babu S"],
  "Irinjalakuda": ["Akshay Sasidharan", "Athul", "Sanjay Babu S"],
  "Irikkur": ["FA 1", "FA 2"],
  "Kalpetta": ["Ragesh Ambadi"],
  "Kanhangad": ["FA 1", "FA 2"],
  "Kanjirappally": ["Gokul PG", "Harikrishnan AB", "Shiju Abraham"],
  "Karunagappally": ["Akshay M"],
  "Kasaragod": ["Sivadas MM"],
  "Kattakkada": ["Ajin Anand", "Anandhu RS", "Sadheerthan"],
  "Kayamkulam": ["FA 1", "FA 2"],
  "Kazhakkoottam": ["ASWIN DAS T.S.", "Anandu Ashok A", "Abhinanad BS"],
  "Kodungallur": ["Akhil PU", "Samal Krishnan"],
  "Konni": ["FA 1", "FA 2"],
  "Kottarakkara": ["Anoop RK", "Vishnu K J", "Midhun"],
  "Kottayam": ["FA 1", "FA 2"],
  "Kovalam": ["Anandhu RS", "Arun R"],
  "Kozhikode North": ["Abhishek", "Arjun PR"],
  "Kunnamkulam": ["FA 1", "FA 2"],
  "Kunnathunad": ["Sarath Sadan"],
  "Kunnathur": ["FA 1", "FA 2"],
  "Kuttanad": ["FA 1", "FA 2"],
  "Malampuzha": ["Adeep Das"],
  "Manalur": ["Sooraj Krishna", "Anukrishna"],
  "Manjeshwaram": ["Viswajith BV", "Vishnu B"],
  "Mankada": ["Pradeep P K", "Arun Krishnan P C"],
  "Mavelikkara": ["FA 1", "FA 2"],
  "Nattika": ["Krishnadathan KS", "Roshith", "Sarath PV"],
  "Nedumangad": ["FA 1", "FA 2"],
  "Nemom": ["Ashwin Kumar VA", "Arun R", "Dinu S"],
  "Nenmara": ["FA 1", "FA 2"],
  "Ollur": ["FA 1", "FA 2"],
  "Ottapalam": ["FA 1", "FA 2"],
  "Pala": ["Anirudh Vinod", "Vishnu B Ambalamattam", "Jishnu Suresh"],
  "Palakkad": ["Goshal krishna", "Vignesh R", "Prasad A", "Jayakrishnan", "Shibu P", "Ajith K"],
  "Peerumade": ["FA 1", "FA 2"],
  "Peravoor": ["Kiran Dev G", "Vishnu Kunnummal"],
  "Perumbavoor": ["Sarath Sadan", "JithuKrishanan Babu"],
  "Ponnani": ["YADHU KRISHNAN V"],
  "Poonjar": ["Sajan P Nair", "Rahul", "Anandhu Anil Olickka"],
  "Puthukkad": ["FA 1", "FA 2"],
  "Ranni": ["FA 1", "FA 2"],
  "Shornur": ["FA 1", "FA 2"],
  "Sultan Bathery": ["FA 1", "FA 2"],
  "Sulthan Bathery": ["Nithin KV"],
  "Thiruvalla": ["Nikhil MR", "Nidhin P Nair", "Umeshkumar"],
  "Thiruvambady": ["FA 1", "FA 2"],
  "Thiruvananthapuram": ["Sudheesh", "Akhil R", "Arun A"],
  "Thodupuzha": ["FA 1", "FA 2"],
  "Thrikkakara": ["Adarshjith"],
  "Thripunithura": ["ASWIN Sabu", "JithuKrishanan Babu"],
  "Thrissur": ["Yadhukrishnan MP", "Amal KS"],
  "Udumbanchola": ["FA 1", "FA 2"],
  "Vaikom": ["FA 1", "FA 2"],
  "Vamanapuram": ["FA 1", "FA 2"],
  "Varkala": ["FA 1", "FA 2"],
  "Vattiyoorkavu": ["Gokul G", "Jijomon SP"],
  "Wadakkanchery": ["Vysakh", "ADARSH V M"]
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
  var s = String(label || "").trim().replace(/[–—]/g, "-");
  if (/^18\s*-\s*19$/i.test(s)) {
    return "18-19";
  }
  return s;
}

/** Age key for weight tables (matches weight_ground_input.normalize_age_label). */
function normalizeAgeLabelForWeight(label) {
  return normalizeAgeLabel(label);
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
  if (k === "sulthan bathery" || k === "sultan bathery") return "Sultan Bathery";
  if (k === "irinjalakkuda" || k === "irinjalakuda") return "Irinjalakuda";
  if (k === "beypur" || k === "beypore") return "Beypur";
  if (k === "kazhakootam" || k === "kazhakkootam" || k === "kazhakkoottam") return "Kazhakkoottam";
  // Kunnathunad — common sheet/typo variants (if AC unknown, caste×gender×age norm stays 0)
  if (
    k === "kunnathunad" ||
    k === "kunnathunad ac" ||
    k === "kunnathumar" ||
    k === "kunnathunadu" ||
    k === "kunnathunadu ac" ||
    k === "kunnathunadac" ||
    k === "gunnar thunadu" ||
    k === "kunnathumad"
  ) {
    return "Kunnathunad";
  }

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

/**
 * Vote normalization for GevsVE — matches scripts/weight_ground_input.normalize_party.
 * Empty → ""; Not Voted; LDF/UDF/BJP/NDA; explicit Others; else raw trimmed (unknown → weight 1).
 */
function normalizeVoteForGevs(raw) {
  var s = String(raw === null || raw === undefined ? "" : raw).trim();
  if (!s || s.toLowerCase() === "nan") {
    return "";
  }
  var u = s.toUpperCase().replace(/\s+/g, "");
  if (u.indexOf("NOT") >= 0 && u.indexOf("VOT") >= 0) {
    return "Not Voted";
  }
  if (u === "LDF") {
    return "LDF";
  }
  if (u === "UDF") {
    return "UDF";
  }
  if (u === "BJP/NDA" || u === "BJP-NDA" || u === "BJPNDA" || u === "BJP" || u === "NDA") {
    return "BJP/NDA";
  }
  if (u === "OTHERS" || u === "OTHER") {
    return "Others";
  }
  return s;
}

/** Column J vs K for historical GevsVE (Kasaragod & Nemom use 2024 GE). */
function getVoteRawForGevs(ac, jRaw, kRaw) {
  var B = normalizeAcName(ac);
  if (B === "Kasaragod" || B === "Nemom") {
    return kRaw;
  }
  return jRaw;
}

/** Historical vote % for seven ACs (same numbers as weight_ground_input.py). */
function getGevsHistoricalPercents(ac) {
  var B = normalizeAcName(ac);
  if (B === "Chathannoor") {
    return { ldf: 43.12, udf: 24.93, bjp: 30.61 };
  }
  if (B === "Attingal") {
    return { ldf: 47.35, udf: 25.02, bjp: 25.92 };
  }
  if (B === "Kattakkada") {
    return { ldf: 45.49, udf: 29.55, bjp: 23.77 };
  }
  if (B === "Manjeshwaram") {
    return { ldf: 23.57, udf: 38.14, bjp: 37.7 };
  }
  if (B === "Poonjar") {
    return { ldf: 41.94, udf: 24.76, bjp: 29.92 };
  }
  if (B === "Kasaragod") {
    return { ldf: 17.67, udf: 49.57, bjp: 31.76 };
  }
  if (B === "Nemom") {
    return { ldf: 24.59, udf: 28.85, bjp: 45.18 };
  }
  return null;
}

function percentsToGevsFractions(p) {
  var ldf = p.ldf / 100;
  var udf = p.udf / 100;
  var bjp = p.bjp / 100;
  var oth = Math.max(0, 1 - ldf - udf - bjp);
  return { ldf: ldf, udf: udf, bjp: bjp, oth: oth };
}

function isGevsSevenAc(ac) {
  return getGevsHistoricalPercents(ac) !== null;
}

/**
 * Target fraction for party (LDF/UDF/BJP/NDA/Others). Others = 1 - LDF - UDF - BJP.
 */
function getGevsTargetFractionForParty(ac, party) {
  var p = getGevsHistoricalPercents(ac);
  if (!p) {
    return null;
  }
  var f = percentsToGevsFractions(p);
  if (party === "LDF") {
    return f.ldf;
  }
  if (party === "UDF") {
    return f.udf;
  }
  if (party === "BJP/NDA") {
    return f.bjp;
  }
  if (party === "Others") {
    return f.oth;
  }
  return null;
}

/**
 * GevsVE = target_fraction / (count_party / n_ac); matches weight_ground_input.compute_weights.
 */
function computeGevsVE(ac, jRaw, kRaw, partyCountMap, nAc, emptyKey) {
  var acN = normalizeAcName(ac);
  if (!isGevsSevenAc(acN) || nAc <= 0) {
    return 1;
  }
  var party = normalizeVoteForGevs(getVoteRawForGevs(acN, jRaw, kRaw));
  if (party === "" || party === "Not Voted") {
    return 1;
  }
  var t = getGevsTargetFractionForParty(acN, party);
  if (t === null || !isFinite(t) || t <= 0) {
    return 1;
  }
  var key = party === "" ? emptyKey : party;
  var cnt = partyCountMap[key];
  if (cnt === undefined || cnt === null) {
    cnt = 0;
  }
  var sample = cnt / nAc;
  if (!isFinite(sample) || sample <= 0) {
    return 1;
  }
  var r = t / sample;
  if (!isFinite(r) || r <= 0) {
    return 1;
  }
  return r;
}

/** weight = target / (bin_count / n_ac) — same as Python. */
function computeRakingWeight(targetFrac, binCount, nAc) {
  if (!isFinite(targetFrac) || targetFrac <= 0 || !nAc || nAc <= 0) {
    return 1;
  }
  var sample = binCount / nAc;
  if (!isFinite(sample) || sample <= 0) {
    return 1;
  }
  var w = targetFrac / sample;
  if (!isFinite(w) || w <= 0) {
    return 1;
  }
  return w;
}

/**
 * Labels for J/K (2021 AE / 2024 GE). Never leave blank: missing payload → "Not Voted"
 * so the sheet stays consistent with L/M (2026 / who will win).
 */
function voteForSheet(v) {
  var raw = String(v === null || v === undefined ? "" : v).trim();
  if (!raw) {
    return "Not Voted";
  }
  var p = normalizeParty(v);
  if (p === "UDF" || p === "LDF" || p === "BJP/NDA") {
    return p;
  }
  return raw;
}

/**
 * Labels for L/M (2026 vote / who will win). Empty → "Others" (three-front contest default).
 */
function votePredictionForSheet(v) {
  return normalizeParty(v);
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

/**
 * Raking weights: target_pop_share / sample_share within AC (matches weight_ground_input.compute_weights).
 * Pass nAc and stratum counts for (caste & gender & age) for this row's AC.
 */
function resolveWeightsFromLabelsAndCounts(ac, casteLabel, genderLabel, ageLabel, nAc, casteCount, genderCount, ageCount) {
  var acName = normalizeAcName(ac);
  var dem = getDemographicsMap();
  var acData = dem[acName];

  var ck = canonicalCasteKey(casteLabel);
  var targetC = (acData && acData[ck] !== undefined && acData[ck] !== null) ? acData[ck] / 100 : NaN;
  var casteW = computeRakingWeight(targetC, casteCount, nAc);

  var g = String(genderLabel || "").trim().toLowerCase();
  var targetG = NaN;
  if (acData) {
    if (g === "male") {
      targetG = acData.male / 100;
    } else if (g === "female") {
      targetG = acData.female / 100;
    }
  }
  var genderW = computeRakingWeight(targetG, genderCount, nAc);

  var ageKey = normalizeAgeLabelForWeight(ageLabel);
  var targetA = AGE_WEIGHTS.hasOwnProperty(ageKey) ? AGE_WEIGHTS[ageKey] : NaN;
  var ageW = computeRakingWeight(targetA, ageCount, nAc);

  var norm = casteW * genderW * ageW;
  return {
    casteW: casteW,
    genderW: genderW,
    ageW: ageW,
    norm: norm
  };
}

/**
 * Legacy entry point: when counts unknown, uses nAc=1 so weight = target; fastRecalcDemographicWeightsAfterAppend fixes.
 */
function resolveWeights(ac, casteVal, genderVal, ageVal) {
  var acName = normalizeAcName(ac);

  var cLab = "";
  if (isValidCasteLabel(casteVal)) {
    cLab = canonicalCasteKey(casteVal);
  } else if (isTextLabel(casteVal)) {
    cLab = canonicalCasteKey(casteVal);
  } else {
    cLab = getCasteLabelFromWeight(acName, casteVal);
  }

  var gLab = "";
  if (isValidGenderLabel(genderVal)) {
    gLab = displayGenderLabel(genderVal);
  } else if (isTextLabel(genderVal)) {
    gLab = displayGenderLabel(genderVal);
  } else {
    gLab = getGenderLabelFromWeight(acName, genderVal);
  }

  var aLab = "20-29";
  if (ageVal != null && isValidAgeLabel(ageVal)) {
    aLab = normalizeAgeLabel(ageVal);
  } else if (!isTextLabel(ageVal)) {
    aLab = getAgeLabelFromAny(ageVal);
  } else {
    aLab = normalizeAgeLabel(ageVal);
  }

  if (!cLab || cLab === "Unknown") {
    cLab = "Others";
  }
  if (!gLab || gLab === "Unknown") {
    gLab = "Male";
  }
  if (!aLab || aLab === "Unknown" || !AGE_WEIGHTS.hasOwnProperty(normalizeAgeLabelForWeight(aLab))) {
    aLab = "20-29";
  }

  return resolveWeightsFromLabelsAndCounts(acName, cLab, gLab, aLab, 1, 1, 1, 1);
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
  // Equal male/female % (e.g. Kasaragod 50/50): same weight maps to both — cannot infer from weight alone.
  if (Math.abs(acData.male - acData.female) < 1e-9) {
    return "Unknown";
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
   GevsVE + raking weights — aligned with scripts/weight_ground_input.py
   GevsVE = target_fraction / (party_count / n_ac); seven ACs only.
   ========================================================================= */

var GEVS_EMPTY_PARTY_KEY = "__empty__";

/**
 * Recompute D,F,H,N (raking) for every row in the same AC as the last appended row.
 */
function fastRecalcDemographicWeightsAfterAppend(sheet, lastRow) {
  if (lastRow < 2) {
    return;
  }
  var numRows = lastRow - 1;
  var block = sheet.getRange(2, 2, lastRow, 9).getValues();
  var lastIdx = numRows - 1;
  var acNew = normalizeAcName(block[lastIdx][0]);

  var casteCounts = {};
  var genderCounts = {};
  var ageCounts = {};
  var nAc = 0;
  var i;
  for (i = 0; i < numRows; i++) {
    if (normalizeAcName(block[i][0]) !== acNew) {
      continue;
    }
    nAc++;
    var cLab = canonicalCasteKey(block[i][3]);
    var gLab = String(block[i][5] || "").trim();
    var gKey = gLab.toLowerCase() === "male" ? "Male" : (gLab.toLowerCase() === "female" ? "Female" : gLab);
    var ageKey = normalizeAgeLabelForWeight(block[i][7]);
    if (!AGE_WEIGHTS.hasOwnProperty(ageKey)) {
      ageKey = "20-29";
    }
    casteCounts[cLab] = (casteCounts[cLab] || 0) + 1;
    genderCounts[gKey] = (genderCounts[gKey] || 0) + 1;
    ageCounts[ageKey] = (ageCounts[ageKey] || 0) + 1;
  }
  if (nAc === 0) {
    return;
  }

  var dVals = sheet.getRange(2, 4, numRows, 1).getValues();
  var fVals = sheet.getRange(2, 6, numRows, 1).getValues();
  var hVals = sheet.getRange(2, 8, numRows, 1).getValues();
  var nVals = sheet.getRange(2, 14, numRows, 1).getValues();

  for (i = 0; i < numRows; i++) {
    if (normalizeAcName(block[i][0]) !== acNew) {
      continue;
    }
    var cLab2 = canonicalCasteKey(block[i][3]);
    var gLab2 = String(block[i][5] || "").trim();
    var gKey2 = gLab2.toLowerCase() === "male" ? "Male" : (gLab2.toLowerCase() === "female" ? "Female" : gLab2);
    var ageKey2 = normalizeAgeLabelForWeight(block[i][7]);
    if (!AGE_WEIGHTS.hasOwnProperty(ageKey2)) {
      ageKey2 = "20-29";
    }
    var w = resolveWeightsFromLabelsAndCounts(
      acNew,
      cLab2,
      gKey2,
      ageKey2,
      nAc,
      casteCounts[cLab2] || 1,
      genderCounts[gKey2] || 1,
      ageCounts[ageKey2] || 1
    );
    dVals[i][0] = w.casteW;
    fVals[i][0] = w.genderW;
    hVals[i][0] = w.ageW;
    nVals[i][0] = w.norm;
  }

  sheet.getRange(2, 4, numRows, 1).setValues(dVals);
  sheet.getRange(2, 6, numRows, 1).setValues(fVals);
  sheet.getRange(2, 8, numRows, 1).setValues(hVals);
  sheet.getRange(2, 14, numRows, 1).setValues(nVals);
  sheet.getRange(2, 4, numRows, 1).setNumberFormat("0.00000000");
  sheet.getRange(2, 6, numRows, 1).setNumberFormat("0.00000000");
  sheet.getRange(2, 8, numRows, 1).setNumberFormat("0.00000000");
  sheet.getRange(2, 14, numRows, 1).setNumberFormat("0.0000000000");
}

/**
 * After a new row: all rows in that AC get new O,P (n_ac and party shares changed).
 */
function fastRecalcGevsveAfterAppend(sheet, lastRow) {
  if (lastRow < 2) {
    return;
  }

  var numRows = lastRow - 1;
  var allData = sheet.getRange(2, 1, numRows, 16).getValues();
  var lastIdx = numRows - 1;
  var acNew = normalizeAcName(allData[lastIdx][1]);
  var EMPTY = GEVS_EMPTY_PARTY_KEY;

  var nAc = 0;
  var partyMap = {};
  var i;
  for (i = 0; i < numRows; i++) {
    if (normalizeAcName(allData[i][1]) !== acNew) {
      continue;
    }
    nAc++;
    var rawV = getVoteRawForGevs(acNew, allData[i][9], allData[i][10]);
    var p = normalizeVoteForGevs(rawV);
    var key = p === "" ? EMPTY : p;
    partyMap[key] = (partyMap[key] || 0) + 1;
  }
  if (nAc === 0) {
    return;
  }

  var colOPValues = sheet.getRange(2, 15, numRows, 2).getValues();
  for (i = 0; i < numRows; i++) {
    if (normalizeAcName(allData[i][1]) !== acNew) {
      continue;
    }
    var nVal = asFiniteNumber(allData[i][13], 0);
    var oVal = computeGevsVE(acNew, allData[i][9], allData[i][10], partyMap, nAc, EMPTY);
    colOPValues[i][0] = oVal;
    colOPValues[i][1] = oVal * nVal;
  }

  sheet.getRange(2, 15, numRows, 2).setValues(colOPValues);
  sheet.getRange(2, 15, numRows, 2).setNumberFormat("0.0000000000");
}

/** Full sheet: O = GevsVE, P = O × N (matches Python batch). */
function recalcGevsveAndFinalValues() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME) || ss.getSheets()[0];
  var lr = sheet.getLastRow();
  if (lr < 2) {
    return;
  }

  ensureHeaders(sheet);

  var numRows = lr - 1;
  var data = sheet.getRange(2, 1, numRows, 16).getValues();
  var EMPTY = GEVS_EMPTY_PARTY_KEY;

  var nByAc = {};
  var partyByAc = {};
  var i;
  for (i = 0; i < numRows; i++) {
    var ac = normalizeAcName(data[i][1]);
    nByAc[ac] = (nByAc[ac] || 0) + 1;
    var rawV = getVoteRawForGevs(ac, data[i][9], data[i][10]);
    var party = normalizeVoteForGevs(rawV);
    var key = party === "" ? EMPTY : party;
    if (!partyByAc[ac]) {
      partyByAc[ac] = {};
    }
    partyByAc[ac][key] = (partyByAc[ac][key] || 0) + 1;
  }

  var out = [];
  for (i = 0; i < numRows; i++) {
    var acJ = normalizeAcName(data[i][1]);
    var nAc = nByAc[acJ] || 1;
    var pmap = partyByAc[acJ] || {};
    var nVal = asFiniteNumber(data[i][13], 0);
    var oVal = computeGevsVE(acJ, data[i][9], data[i][10], pmap, nAc, EMPTY);
    out.push([oVal, oVal * nVal]);
  }

  sheet.getRange(2, 15, numRows, 2).setValues(out);
  sheet.getRange(2, 15, numRows, 2).setNumberFormat("0.0000000000");
  SpreadsheetApp.flush();
}

/**
 * Full sheet raking: D,F,H,N from demographics sheet + AGE_WEIGHTS (same as Python).
 * May be slow on very large sheets; use after bulk import or column fixes.
 */
function recalcAllDemographicWeightsAndNormalized() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME) || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lr = sheet.getLastRow();
  if (lr < 2) {
    return;
  }
  ensureHeaders(sheet);
  var numRows = lr - 1;
  var block = sheet.getRange(2, 2, lr, 9).getValues();

  var nByAc = {};
  var casteByAc = {};
  var genderByAc = {};
  var ageByAc = {};
  var i;
  for (i = 0; i < numRows; i++) {
    var ac = normalizeAcName(block[i][0]);
    nByAc[ac] = (nByAc[ac] || 0) + 1;
    var cLab = canonicalCasteKey(block[i][3]);
    var gLab = String(block[i][5] || "").trim();
    var gKey = gLab.toLowerCase() === "male" ? "Male" : (gLab.toLowerCase() === "female" ? "Female" : gLab);
    var ageKey = normalizeAgeLabelForWeight(block[i][7]);
    if (!AGE_WEIGHTS.hasOwnProperty(ageKey)) {
      ageKey = "20-29";
    }
    if (!casteByAc[ac]) {
      casteByAc[ac] = {};
    }
    if (!genderByAc[ac]) {
      genderByAc[ac] = {};
    }
    if (!ageByAc[ac]) {
      ageByAc[ac] = {};
    }
    casteByAc[ac][cLab] = (casteByAc[ac][cLab] || 0) + 1;
    genderByAc[ac][gKey] = (genderByAc[ac][gKey] || 0) + 1;
    ageByAc[ac][ageKey] = (ageByAc[ac][ageKey] || 0) + 1;
  }

  var dVals = sheet.getRange(2, 4, numRows, 1).getValues();
  var fVals = sheet.getRange(2, 6, numRows, 1).getValues();
  var hVals = sheet.getRange(2, 8, numRows, 1).getValues();
  var nVals = sheet.getRange(2, 14, numRows, 1).getValues();

  for (i = 0; i < numRows; i++) {
    var ac2 = normalizeAcName(block[i][0]);
    var nA = nByAc[ac2] || 1;
    var cLab2 = canonicalCasteKey(block[i][3]);
    var gLab2 = String(block[i][5] || "").trim();
    var gKey2 = gLab2.toLowerCase() === "male" ? "Male" : (gLab2.toLowerCase() === "female" ? "Female" : gLab2);
    var ageKey2 = normalizeAgeLabelForWeight(block[i][7]);
    if (!AGE_WEIGHTS.hasOwnProperty(ageKey2)) {
      ageKey2 = "20-29";
    }
    var cb = casteByAc[ac2] || {};
    var gb = genderByAc[ac2] || {};
    var ab = ageByAc[ac2] || {};
    var w = resolveWeightsFromLabelsAndCounts(
      ac2,
      cLab2,
      gKey2,
      ageKey2,
      nA,
      cb[cLab2] || 1,
      gb[gKey2] || 1,
      ab[ageKey2] || 1
    );
    dVals[i][0] = w.casteW;
    fVals[i][0] = w.genderW;
    hVals[i][0] = w.ageW;
    nVals[i][0] = w.norm;
  }

  sheet.getRange(2, 4, numRows, 1).setValues(dVals);
  sheet.getRange(2, 6, numRows, 1).setValues(fVals);
  sheet.getRange(2, 8, numRows, 1).setValues(hVals);
  sheet.getRange(2, 14, numRows, 1).setValues(nVals);
  sheet.getRange(2, 4, numRows, 1).setNumberFormat("0.00000000");
  sheet.getRange(2, 6, numRows, 1).setNumberFormat("0.00000000");
  sheet.getRange(2, 8, numRows, 1).setNumberFormat("0.00000000");
  sheet.getRange(2, 14, numRows, 1).setNumberFormat("0.0000000000");
}

/**
 * Dedup TTL (seconds). A submissionId is remembered for this long —
 * any retry with the same ID within this window is silently accepted
 * without appending a duplicate row.
 */
var DEDUP_TTL_SEC = 300;   // 5 minutes

function doPost(e) {
  // --- Acquire a script-wide lock so concurrent POSTs are serialized ---
  var lock = LockService.getScriptLock();
  try {
    // Wait up to 20 s for the lock; if another POST is in-flight, this queues.
    lock.waitLock(20000);
  } catch (lockErr) {
    return jsonResponse({ status: "error", message: "Server busy, please retry." });
  }

  try {
    var data = JSON.parse(e.postData.contents);

    // --- Dedup: reject if this submissionId was already processed ---
    var subId = data.submissionId || "";
    if (subId) {
      var dedupCache = CacheService.getScriptCache();
      var dedupKey = "dedup_" + subId;
      if (dedupCache.get(dedupKey)) {
        // Already appended — return success without writing again
        return jsonResponse({ status: "success", duplicate: true });
      }
      // Mark as seen *before* the append — worst case we skip a true retry
      // rather than accidentally writing twice.
      dedupCache.put(dedupKey, "1", DEDUP_TTL_SEC);
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(MAIN_SHEET_NAME) || ss.getSheets()[0];
    ensureHeaders(sheet);

    var ac = normalizeAcName(data.ac);
    var vote2026 = votePredictionForSheet(data.vote2026);
    var whoWillWin = votePredictionForSheet(data.whoWillWin);

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
      fastRecalcDemographicWeightsAfterAppend(sheet, lr);
    } catch (demErr) {
      Logger.log("fastRecalcDemographicWeightsAfterAppend failed: " + demErr);
    }

    try {
      fastRecalcGevsveAfterAppend(sheet, lr);
    } catch (fastErr) {
      Logger.log("fastRecalcGevsveAfterAppend failed: " + fastErr);
    }

    try {
      clearEntriesCache();
    } catch (cacheErr) {
      Logger.log("clearEntriesCache failed: " + cacheErr);
    }

    return jsonResponse({ status: "success", opRecalculated: true });
  } catch (err) {
    return jsonResponse({ status: "error", message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  var action = e && e.parameter && e.parameter.action;
  try {
    if (action === "entries") {
      var fromP = e.parameter.from;
      var toP = e.parameter.to;
      var dateP = e.parameter.date || "";
      var cache = CacheService.getScriptCache();
      var cacheKey;
      var payload;
      if (fromP && toP) {
        cacheKey = "ent_v1_r_" + String(fromP) + "_" + String(toP);
        var cachedR = cache.get(cacheKey);
        if (cachedR) {
          return ContentService.createTextOutput(cachedR).setMimeType(ContentService.MimeType.JSON);
        }
        payload = getEntriesForDateRange(String(fromP), String(toP));
      } else {
        cacheKey = "ent_v1_d_" + (dateP === "" ? "all" : String(dateP));
        var cachedD = cache.get(cacheKey);
        if (cachedD) {
          return ContentService.createTextOutput(cachedD).setMimeType(ContentService.MimeType.JSON);
        }
        payload = getEntriesForDate(dateP);
      }
      var jsonStr = JSON.stringify(payload);
      if (jsonStr.length <= ENTRIES_CACHE_MAX_CHARS) {
        cache.put(cacheKey, jsonStr, ENTRIES_CACHE_TTL_SEC);
      }
      return ContentService.createTextOutput(jsonStr).setMimeType(ContentService.MimeType.JSON);
    }
    if (action === "summary") {
      return jsonResponse(getSummaryForDate(e.parameter.date || ""));
    }
    if (action === "dates") {
      return jsonResponse(getAvailableDates());
    }
    if (action === "timestamps") {
      return jsonResponse(getUniqueTimestampDates());
    }
    if (action === "refreshDemographics") {
      clearDemographicsCache();
      return jsonResponse({ status: "ok" });
    }
    if (action === "recalcOP") {
      recalcAllDemographicWeightsAndNormalized();
      recalcGevsveAndFinalValues();
      return jsonResponse({ status: "ok", message: "D–N and O/P recalculated (raking + GevsVE)" });
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

function entriesCacheKey(suffix) {
  var gen = PropertiesService.getScriptProperties().getProperty("entriesCacheGen") || "0";
  return "ent_v1_g" + gen + "_" + suffix;
}

/** Bump generation so all prior entry-cache keys are ignored (call after sheet changes). */
function clearEntriesCache() {
  var p = PropertiesService.getScriptProperties();
  var g = parseInt(p.getProperty("entriesCacheGen") || "0", 10) + 1;
  p.setProperty("entriesCacheGen", String(g));
}

/** Merge sorted 0-based row indices into contiguous ranges for fewer getRange calls. */
function groupContiguousIndices(sorted) {
  if (sorted.length === 0) return [];
  var groups = [];
  var s = sorted[0];
  var e = sorted[0];
  var k;
  for (k = 1; k < sorted.length; k++) {
    if (sorted[k] === e + 1) {
      e = sorted[k];
    } else {
      groups.push([s, e]);
      s = e = sorted[k];
    }
  }
  groups.push([s, e]);
  return groups;
}

/**
 * After filtering by timestamp (column A only), load full 16-column rows only for matching indices.
 * When most rows are excluded (e.g. one day in a large sheet), this is much faster than reading all columns for every row.
 */
function readEntryRowsByIndices(sheet, indices, numRows, displayA) {
  if (indices.length === 0) {
    return { rows: [], dispTs: [] };
  }
  var sorted = indices.slice().sort(function(a, b) { return a - b; });
  var allRows = [];
  var allDisp = [];
  var i;
  var j;
  if (sorted.length === numRows) {
    var dataAll = sheet.getRange(2, 1, numRows, 16).getValues();
    for (i = 0; i < numRows; i++) {
      allRows.push(dataAll[i]);
      allDisp.push(displayA[i][0]);
    }
    return { rows: allRows, dispTs: allDisp };
  }
  var groups = groupContiguousIndices(sorted);
  for (i = 0; i < groups.length; i++) {
    var s = groups[i][0];
    var e = groups[i][1];
    var height = e - s + 1;
    var sheetRow = 2 + s;
    var chunk = sheet.getRange(sheetRow, 1, height, 16).getValues();
    for (j = 0; j < chunk.length; j++) {
      allRows.push(chunk[j]);
      allDisp.push(displayA[s + j][0]);
    }
  }
  return { rows: allRows, dispTs: allDisp };
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

  // Sheets sometimes returns a serial number (days since 1899-12-30) instead of Date
  if (typeof cell === "number" && isFinite(cell) && cell > 0) {
    var fromSerial = new Date((cell - 25569) * 86400000);
    if (isNaN(fromSerial.getTime())) {
      return null;
    }
    return Utilities.formatDate(fromSerial, "Asia/Kolkata", "yyyy-MM-dd");
  }

  var s = String(cell).trim();

  // 1) ISO yyyy-mm-dd (matches sv-SE Kolkata strings like "2026-04-01 15:30:45" and Sheets ISO)
  var m = s.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (m) {
    return m[1] + "-" + m[2] + "-" + m[3];
  }

  // 2) DD-MM-YYYY with dashes (some browsers / CSV; en-IN sometimes uses dashes)
  m = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})(?:\s|,|$)/);
  if (m) {
    var dDm = parseInt(m[1], 10);
    var moDm = parseInt(m[2], 10);
    var yDm = parseInt(m[3], 10);
    if (moDm >= 1 && moDm <= 12 && dDm >= 1 && dDm <= 31 && yDm >= 1900 && yDm <= 2100) {
      return yDm + "-" + pad2(moDm) + "-" + pad2(dDm);
    }
  }

  // 3) d/m/y — treated as DAY/MONTH/YEAR (India). Note: ambiguous with US m/d/y for days ≤12.
  m = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (m) {
    var d = parseInt(m[1], 10);
    var mo = parseInt(m[2], 10);
    var y = parseInt(m[3], 10);
    if (y < 100) {
      y += 2000;
    }
    return y + "-" + pad2(mo) + "-" + pad2(d);
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

  // yyyy-mm-dd (admin date picker) → add slash formats actually stored in Sheet1 (e.g. "1/4/2026, 7:50 AM")
  var isoOnly = String(dateStr).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (isoOnly) {
    var yIso = parseInt(isoOnly[1], 10);
    var moIso = parseInt(isoOnly[2], 10);
    var dIso = parseInt(isoOnly[3], 10);
    var calIso = new Date(yIso, moIso - 1, dIso);
    patterns.push(Utilities.formatDate(calIso, tz, "d/M/yyyy"));
    patterns.push(Utilities.formatDate(calIso, tz, "dd/MM/yyyy"));
    patterns.push(Utilities.formatDate(calIso, tz, "d/M/yy"));
    patterns.push(Utilities.formatDate(calIso, tz, "dd/MM/yy"));
    patterns.push(Utilities.formatDate(calIso, tz, "M/d/yyyy"));
    patterns.push(Utilities.formatDate(calIso, tz, "MM/dd/yyyy"));
    patterns.push(dIso + "/" + moIso + "/" + yIso);
    patterns.push(pad2(dIso) + "/" + pad2(moIso) + "/" + yIso);
  }

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

function legacySubstringRowMatch(tsCell, dateStr) {
  if (!dateStr) return true;
  var ts = String(tsCell);
  var patterns = timestampPatternsForDateKey(dateStr);
  for (var i = 0; i < patterns.length; i++) {
    if (ts.indexOf(patterns[i]) !== -1) {
      return true;
    }
  }
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
    if (rowYmd) {
      return rowYmd === targetYmd;
    }
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
    headers.forEach(function(h, i) { 
      obj[h] = row[i]; 
    });
    obj.normalizedScore = row[15];
    if (displayTimestamps && displayTimestamps[idx] !== undefined && displayTimestamps[idx] !== null && String(displayTimestamps[idx]).trim() !== "") {
      obj.timestamp = displayTimestamps[idx];
    }
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
      rowNums.push(i + 2);
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
  var i;
  for (i = 0; i < numRows; i++) {
    var tsF = timestampForDateFilter(displayA[i][0], data[i][0]);
    var rowYmd = cellToYmdKolkata(tsF);
    var include = false;
    if (rowYmd) {
      include = rowYmd >= fromYmd && rowYmd <= toYmd;
    } else if (fromYmd === toYmd) {
      include = rowMatchesDateFilter(tsF, fromYmd);
    }
    if (include) {
      filtered.push(data[i]);
      dispTs.push(displayA[i][0]);
      rowNums.push(i + 2);
    }
  }

  var entries = mapRowsToEntries(filtered, dispTs, rowNums);
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

/**
 * Returns unique yyyy-MM-dd dates found in column A of Sheet1.
 * The dashboard can call ?action=timestamps to know which calendar days have data,
 * regardless of whether a summary tab exists for that day.
 */
function getUniqueTimestampDates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME) || ss.getSheets()[0];
  var lr = sheet.getLastRow();
  if (lr < 2) return { dates: [] };

  var numRows = lr - 1;
  var tsCol = sheet.getRange(2, 1, numRows, 1).getDisplayValues();
  var raw   = sheet.getRange(2, 1, numRows, 1).getValues();
  var seen = {};
  var result = [];
  for (var i = 0; i < numRows; i++) {
    var display = tsCol[i][0];
    var rawVal  = raw[i][0];
    var tsF = timestampForDateFilter(display, rawVal);
    var ymd = cellToYmdKolkata(tsF);
    if (ymd && !seen[ymd]) {
      seen[ymd] = true;
      result.push(ymd);
    }
  }
  result.sort();
  return { dates: result };
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

    var casteLabel = casteLabelForSheet(ac, cIn, row[3]);
    var genderLabel = genderLabelForSheet(ac, gIn, row[5]);
    var ageLabel = ageLabelForSheet(aIn, row[7]);

    var r = i + 2;
    sheet.getRange(r, 2).setValue(ac);
    sheet.getRange(r, 5).setValue(casteLabel);
    sheet.getRange(r, 7).setValue(genderLabel);
    sheet.getRange(r, 9).setValue(ageLabel);
    sheet.getRange(r, 12).setValue(v26);
    sheet.getRange(r, 13).setValue(ww);
  }

  recalcAllDemographicWeightsAndNormalized();
  recalcGevsveAndFinalValues();
  Logger.log("fixTextRows updated " + numRows + " rows.");
}

/**
 * One-time / occasional repair: rows where J–M are blank get explicit labels
 * (same defaults as doPost). Run from Apps Script ▶ after imports or legacy rows.
 */
function repairBlankVoteColumnsJthroughM() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME) || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lr = sheet.getLastRow();
  if (lr < 2) {
    return;
  }
  ensureHeaders(sheet);
  var numRows = lr - 1;
  var data = sheet.getRange(2, 1, numRows, 13).getValues();
  var fixed = 0;
  for (var i = 0; i < data.length; i++) {
    var r = i + 2;
    var jv = String(data[i][9] === null || data[i][9] === undefined ? "" : data[i][9]).trim();
    var kv = String(data[i][10] === null || data[i][10] === undefined ? "" : data[i][10]).trim();
    var lv = String(data[i][11] === null || data[i][11] === undefined ? "" : data[i][11]).trim();
    var mv = String(data[i][12] === null || data[i][12] === undefined ? "" : data[i][12]).trim();
    if (!jv) {
      sheet.getRange(r, 10).setValue("Not Voted");
      fixed++;
    }
    if (!kv) {
      sheet.getRange(r, 11).setValue("Not Voted");
      fixed++;
    }
    if (!lv) {
      sheet.getRange(r, 12).setValue("Others");
      fixed++;
    }
    if (!mv) {
      sheet.getRange(r, 13).setValue("Others");
      fixed++;
    }
  }
  if (fixed > 0) {
    try {
      recalcGevsveAndFinalValues();
    } catch (e) {
      Logger.log("repairBlankVoteColumnsJthroughM: recalc failed " + e);
    }
  }
  Logger.log("repairBlankVoteColumnsJthroughM: filled " + fixed + " empty cells (J–M).");
}

function recalcAllNormalizedScores() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME) || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lr = sheet.getLastRow();
  if (lr < 2) {
    Logger.log("No rows.");
    return;
  }

  recalcAllDemographicWeightsAndNormalized();
  recalcGevsveAndFinalValues();

  Logger.log("recalcAllNormalizedScores: full raking + GevsVE done.");
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
