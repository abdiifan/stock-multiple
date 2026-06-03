// =============================================================================
// PharmaTrack v2 — Exclusion / Filter Rules
// =============================================================================
// This file must be loaded BEFORE script.js.
// All data-exclusion logic lives here so it can be maintained independently
// without touching the main application logic.
// =============================================================================

// ── NON-MEDICAL MATERIAL GROUP KEYWORDS ──────────────────────────────────────
// Case-insensitive partial match against Material Group Name.
// Any group whose name contains one of these strings will be excluded from
// data load AND from all filter dropdowns.
const NON_MEDICAL_GROUPS = [
  "accessory", "accessories",
  "building", "construction",
  "furniture", "office supply", "office supplies",
  "stationery", "vehicle", "it equipment", "computer",
  "clothing", "uniform", "textile",
  "food", "beverage",
];

// ── NON-MEDICAL MATERIAL CODE PATTERNS ───────────────────────────────────────
// Returns true if the material code should be excluded.
//
// Rules:
//  1. Excel scientific notation (e.g. 7E+09) is normalised to integer first.
//  2. Codes that are exactly one digit followed by four or more zeros
//     (e.g. 10000, 20000, 90000) are excluded — these are non-product
//     material group header codes in SAP.
//  3. Codes starting with "4" are excluded (non-trade / project stock range).
//
// NOTE: The old regex /^\d0000/ was too broad — it matched valid 6-digit codes
// like 100001. The corrected /^\d0{4,}$/ anchors both ends.
const isNonMedicalCode = code => {
  let s = String(code).trim();
  if (/e/i.test(s)) s = Math.round(Number(s)).toString(); // fix Excel scientific notation
  return /^\d0{4,}$/.test(s) || /^4/.test(s);
};

// ── NON-MEDICAL GROUP NAME CHECK ─────────────────────────────────────────────
// Returns true if the Material Group Name matches any NON_MEDICAL_GROUPS entry.
const isNonMedicalGroup = name => {
  if (!name) return false;
  const lower = String(name).toLowerCase();
  return NON_MEDICAL_GROUPS.some(g => lower.includes(g));
};
