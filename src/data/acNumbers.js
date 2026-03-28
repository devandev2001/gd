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
  Malampuzha: "58",
  Thrissur: "67",
  Nattika: "68",
  Manalur: "69",
  Kunnathunad: "84",
  Perumbavoor: "74",
  Palakkad: "56",
  "Kozhikode North": "27",
  Kasaragod: "2",
  Manjeshwaram: "1",
};

export function getAcNo(acName) {
  const key = String(acName || "").trim();
  return AC_NUMBER_BY_NAME[key] || "";
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
