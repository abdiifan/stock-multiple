// =============================================================================
// PharmaTrack v2 — filters.js
// Exclusion rules for non-medical / non-pharmaceutical materials.
// This file MUST be loaded before script.js.
//
// isNonMedicalCode(materialCode) → true  = exclude this row
// isNonMedicalGroup(groupName)   → true  = exclude this row
// =============================================================================

/**
 * Returns true if the material code looks like a non-medical / non-trade item
 * that should be excluded from pharmaceutical inventory analysis.
 *
 * Common EPSS patterns to exclude:
 *   - Codes starting with "NT" (Non-Trade)
 *   - Codes starting with "Q"  (Project Stock / Q-type)
 *   - Empty / blank codes
 *
 * Extend this list to match your actual data as needed.
 */
function isNonMedicalCode(code) {
  if (!code) return false;
  const c = String(code).trim().toUpperCase();
  if (!c) return false;

  // Non-Trade prefix
  if (c.startsWith("NT")) return true;

  // You can add more prefixes here, e.g.:
  // if (c.startsWith("SRV")) return true;  // services
  // if (c.startsWith("ASSET")) return true;

  return false;
}

/**
 * Returns true if the material group name is a non-medical category
 * that should be excluded from pharmaceutical inventory analysis.
 *
 * Common EPSS group names to exclude.
 * Extend this list to match your actual material group naming.
 */
function isNonMedicalGroup(groupName) {
  if (!groupName) return false;
  const g = String(groupName).trim().toUpperCase();
  if (!g) return false;

  const EXCLUDED_GROUPS = [
    "NON TRADE",
    "NON-TRADE",
    "NONTRADE",
    "PROJECT STOCK",
    "SERVICES",
    "ASSETS",
    "OFFICE SUPPLIES",
    "STATIONERY",
    "SPARE PARTS",
    "EQUIPMENT",
    "FURNITURE",
  ];

  return EXCLUDED_GROUPS.some(ex => g.includes(ex));
}
