/** Official AC numbers for display only — submitted `ac` remains the constituency name. */
export const AC_NUMBER_BY_NAME = {
  Kattakkada: "138",
  Kovalam: "139",
  Kazhakkoottam: "132",
  Vattiyoorkavu: "133",
  Thiruvananthapuram: "134",
  Nemom: "135",
  Attingal: "128",
  Chathannoor: "126",
  Aranmula: "113",
  Thiruvalla: "111",
  Chengannur: "110",
  Adoor: "115",
  Poonjar: "101",
  Kanjirappally: "100",
  Pala: "93",
  Thrissur: "67",
  Thripunithura: "81",
  Thrikkakara: "83",
  Kunnathunad: "84",
  Palakkad: "56",
  "Kozhikode North": "27",
  Kasaragod: "2",
  Manjeshwaram: "1",
  Perumbavoor: "74",
  Nattika: "68",
  Manalur: "64",
  Malampuzha: "55",
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
  if (k === "kanjirapalli") return AC_NUMBER_BY_NAME.Kanjirappally || "";
  if (k === "thripunitura" || k === "thrippunithura") return AC_NUMBER_BY_NAME.Thripunithura || "";
  if (k === "thrikakkara") return AC_NUMBER_BY_NAME.Thrikkakara || "";

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
