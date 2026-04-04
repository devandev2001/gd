/** Official AC numbers for display only — submitted `ac` remains the constituency name. */
export const AC_NUMBER_BY_NAME = {
  Manjeshwaram: "1",
  Kasaragod: "2",
  Kanhangad: "4",
  Irikkur: "9",
  Kalpetta: "14",
  "Sultan Bathery": "15",
  "Sulthan Bathery": "15",
  Mananthavady: "16",
  Peravoor: "18",
  "Kozhikode North": "27",
  Beypur: "35",
  Thiruvambady: "31",
  Ponnani: "39",
  Mankada: "43",
  Shornur: "51",
  Ottapalam: "50",
  Nenmara: "53",
  Malampuzha: "55",
  Palakkad: "56",
  Puthukkad: "59",
  Chelakkara: "60",
  Kunnamkulam: "62",
  Wadakkanchery: "61",
  Guruvayoor: "63",
  Manalur: "64",
  Ollur: "65",
  Irinjalakuda: "66",
  Thrissur: "67",
  Nattika: "68",
  Chalakkudy: "69",
  Kodungallur: "70",
  Perumbavoor: "74",
  Kalamassery: "77",
  Thripunithura: "81",
  Thrikkakara: "83",
  Kunnathunad: "84",
  Thodupuzha: "86",
  Udumbanchola: "87",
  Devikulam: "88",
  Idukki: "89",
  Peerumade: "90",
  Ettumanoor: "91",
  Pala: "93",
  Vaikom: "94",
  Kanjirappally: "100",
  Kottayam: "104",
  Poonjar: "101",
  Changanassery: "103",
  Ranni: "117",
  Kuttanad: "106",
  Ambalapuzha: "108",
  Chengannur: "110",
  Thiruvalla: "111",
  Aranmula: "113",
  Adoor: "115",
  Konni: "116",
  Harippad: "119",
  Kayamkulam: "120",
  Mavelikkara: "121",
  Cherthala: "122",
  Aroor: "123",
  Kunnathur: "123",
  Karunagappally: "124",
  Kottarakkara: "125",
  Chathannoor: "126",
  Attingal: "128",
  Vamanapuram: "129",
  Varkala: "130",
  Chirayankeezhu: "129",
  Kazhakkoottam: "132",
  Vattiyoorkavu: "133",
  Thiruvananthapuram: "134",
  Nemom: "135",
  Nedumangad: "136",
  Kattakkada: "138",
  Kovalam: "139",
};

function normalizeAcKey(v) {
  return String(v || "")
    .toLowerCase()
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/\(sc\)/g, "")
    .replace(/ac/g, "")
    .replace(/[^a-z]/g, "");
}

export function getAcNo(acName) {
  const key = String(acName || "").trim();
  if (AC_NUMBER_BY_NAME[key]) return AC_NUMBER_BY_NAME[key];

  // Aliases/typos seen in tracker sheets — keep display stable while submitted `ac` stays canonical.
  const k = normalizeAcKey(key);
  if (k === "kazhakootam") return AC_NUMBER_BY_NAME.Kazhakkoottam || "";
  if (k === "perumbaavoor") return AC_NUMBER_BY_NAME.Perumbavoor || "";
  if (k === "kanjirapalli" || k === "kanjirappalli") return AC_NUMBER_BY_NAME.Kanjirappally || "";
  if (k === "thripunitura" || k === "thrippunithura") return AC_NUMBER_BY_NAME.Thripunithura || "";
  if (k === "thrikakkara") return AC_NUMBER_BY_NAME.Thrikkakara || "";
  if (k === "sulthanbathery") return AC_NUMBER_BY_NAME["Sulthan Bathery"] || "";
  if (k === "irinjalakkuda") return AC_NUMBER_BY_NAME.Irinjalakuda || "";
  if (k === "beypur" || k === "beypore") return AC_NUMBER_BY_NAME.Beypur || "";
  if (k === "nattikasc" || k === "nattika") return AC_NUMBER_BY_NAME.Nattika || "";
  if (k === "nemam" || k === "naimam") return AC_NUMBER_BY_NAME.Nemom || "";
  if (k === "kasargod" || k === "kasaragode") return AC_NUMBER_BY_NAME.Kasaragod || "";

  return "";
}

export function formatAcSelectLabel(acName) {
  const no = getAcNo(acName);
  const name = String(acName || "").trim();
  return no ? `${no} — ${name}` : name || "—";
}

/** Sort `constituencyData`-style rows by official AC number. */
export function sortConstituenciesByAcNo(rows) {
  return [...rows].sort((a, b) => {
    const na = getAcNo(a.ac);
    const nb = getAcNo(b.ac);
    if (na && nb) return parseInt(na, 10) - parseInt(nb, 10);
    if (na && !nb) return -1;
    if (!na && nb) return 1;
    return String(a.ac).localeCompare(String(b.ac));
  });
}

/**
 * Sort by constituency name A–Z (case-insensitive).
 * Use this for the survey dropdown so names like Thripunithura / Thrikkakara appear with other “Th…”
 * ACs instead of between unrelated numbers (81/83 vs 134).
 */
export function sortConstituenciesByName(rows) {
  return [...rows].sort((a, b) =>
    String(a.ac).localeCompare(String(b.ac), "en", { sensitivity: "base" })
  );
}
