// =============================================================================
// PharmaTrack v2 — Exclusion / Filter Rules
// =============================================================================
// This file must be loaded BEFORE script.js.
// All data-exclusion logic lives here so it can be maintained independently
// without touching the main application logic.
// =============================================================================

// ── ALLOWED MEDICAL MATERIAL GROUP NAMES (ALLOWLIST) ─────────────────────────
// ONLY groups whose name starts with one of these strings (case-insensitive,
// trimmed) are kept. Everything else is excluded from data load AND dropdowns.
const ALLOWED_MEDICAL_GROUPS = [
  "anesthesia",
  "antidotes",
  "antihistamines",
  "anti-infective medic",
  "antineoplastics",
  "blood products and m",
  "cardiovascular medic",
  "central nervous syst",
  "chemicals and stains",
  "clinical chemistry d",
  "culture, biochemical",
  "dermatological agent",
  "dressing supplies",
  "ear, nose and throat",
  "endocrine disorder a",
  "external quality ass",
  "fluid, electrolyte a",
  "gastroinstinal medic",
  "hematology diagnosti",
  "imaging chemicals an",
  "immunomodulators",
  "injection supplies",
  "laboratory materials",
  "laboratory supplies",
  "miscellaneous chemic",
  "miscellaneous medici",
  "miscellaneous suppl",
  "molecular diagnostic",
  "musculoskeletelal an",
  "obstetrics and gynec",
  "ophthalmic agents",
  "protective supplies",
  "respiratory medicine",
  "sera and immunoglobu",
  "serology and imm dig",
  "surgical sutures",
  "tubes and drains",
  "vitamins",
];

// ── NON-MEDICAL GROUP NAME CHECK ─────────────────────────────────────────────
// Returns TRUE (exclude) if the Material Group Name is NOT in the allowlist.
// Uses startsWith (case-insensitive, trimmed) so SAP-truncated names still match.
const isNonMedicalGroup = name => {
  if (!name) return true;
  const lower = String(name).trim().toLowerCase();
  return !ALLOWED_MEDICAL_GROUPS.some(g => lower.startsWith(g));
};

// ── NON-MEDICAL MATERIAL CODE PATTERNS ───────────────────────────────────────
// Returns TRUE if the material code should be EXCLUDED.
//
// Rules (applied after normalising Excel scientific notation):
//  1. Excel scientific notation (e.g. 7E+09) is normalised to integer string first.
//  2. ALL codes whose first character is a digit 1–9 are excluded.
//     This covers codes starting with 1, 2, 3, 4, 5, 6, 7, 8, or 9 —
//     including project/non-trade (4xxxxx), SAP header codes (10000, 20000…),
//     and all other numeric ranges (100001, 200050, etc.).
//     Only codes with a non-numeric or zero-leading prefix are kept.
const isNonMedicalCode = code => {
  let s = String(code).trim();
  if (/e/i.test(s)) s = Math.round(Number(s)).toString();
  return /^[1-9]/.test(s);
};
