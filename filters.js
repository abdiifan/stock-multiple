// =============================================================================
// PharmaTrack v2 — filters.js
// Exclusion rules for non-medical / non-pharmaceutical materials.
// This file MUST be loaded before script.js.
//
// isNonMedicalCode(materialCode)           → true = exclude this row
// isNonMedicalGroup(groupName)             → true = exclude this row
// isProjectStockDescription(description)   → true = exclude this row
// isExcludedStorageLocation(storageLoc)    → true = exclude this row
// =============================================================================

/**
 * Returns true if the material code looks like a non-medical / non-trade item
 * that should be excluded from pharmaceutical inventory analysis.
 *
 * Common EPSS patterns to exclude:
 *   - Codes starting with "NT" (Non-Trade)
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

/**
 * Returns true if the Special Stock Type Description indicates Project Stock.
 *
 * This catches rows where the Special Stock Type code is not "Q" but the
 * description still resolves to "Project Stock" — both must be excluded.
 */
function isProjectStockDescription(description) {
  if (!description) return false;
  const d = String(description).trim().toUpperCase();
  if (!d) return false;
  return d === "PROJECT STOCK";
}

/**
 * Returns true if the Storage Location code is in the excluded list.
 *
 * These locations hold non-pharmaceutical / project / administrative stock
 * and must be excluded from all inventory analysis.
 */
function isExcludedStorageLocation(storageLoc) {
  if (!storageLoc) return false;
  const s = String(storageLoc).trim().toUpperCase();
  if (!s) return false;

  const EXCLUDED_LOCATIONS = [
    "AA1G", "AA2G", "ADG1", "ARG1", "ASG1",
    "BDG1", "DDG1", "DEG1", "GAG1", "GOG1",
    "HAG1", "HOG1", "HOG2", "JIG1", "JJG1",
    "KDG1", "MKG1", "NBG1", "NKG1", "SEG1",
    "SHG1",
  ];

  return EXCLUDED_LOCATIONS.includes(s);
}
