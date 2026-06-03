// =============================================================================
// PharmaTrack — Pharmaceutical Inventory Management System
// Pure static JS converted from material_fixed.py (Streamlit)
// =============================================================================

// ── CONSTANTS ──────────────────────────────────────────────────────────────
const REQUIRED_COLUMNS = [
  "Material","Material Description","Plant","Plant Name",
  "Storage Location","Description of Storage Location",
  "Special Stock Type","Special Stock Type Description",
  "Unrestricted Stock","Stock in Quality Inspection","Blocked Stock",
  "Batch","Inventory Valuation Type","Material Group Name",
  "Shelf Life Expiration Date","Stock in Transit",
  "Value of Stock in Quality Inspection","Value of Stock in Transit",
  "Value of Unrestricted Stock",
];

const CATEGORY_KEYWORDS = {
  "Anti-infective":   ["antibiotic","anti-infect","antimicrobial","antifungal","antiviral","amoxicillin","ciprofloxacin","azithromycin","metronidazole","ceftriaxone","penicillin","ampicillin","doxycycline","cotrimoxazole","trimethoprim","anti infect"],
  "Analgesic":        ["analgesic","pain","paracetamol","ibuprofen","aspirin","diclofenac","tramadol","morphine","codeine","naproxen","nsaid","anti-inflam","anti inflam"],
  "Cardiovascular":   ["cardiovasc","cardiac","heart","antihypertens","amlodipine","atenolol","captopril","enalapril","lisinopril","losartan","valsartan","digoxin","warfarin","heparin","statin","simvastatin","atorvastatin","hypertens"],
  "Antidiabetic":     ["antidiabet","diabet","insulin","metformin","glibenclamide","glimepiride","pioglitazone","sitagliptin","hypoglyc"],
  "Vitamins & Supplements": ["vitamin","supplement","mineral","zinc","iron","calcium","magnesium","folic","ferrous","multivitamin","nutritional","micronutrient","electrolyte"],
  "Antiallergic":     ["antiallerg","anti-allerg","antihistamine","cetirizine","loratadine","fexofenadine","chlorphenamine","promethazine","allerg","histamine"],
};

const COLORWAY = ["#58a6ff","#3fb950","#d29922","#f85149","#a371f7","#79c0ff","#56d364","#e3b341","#ff7b72","#d2a8ff"];

const PLOTLY_LAYOUT = {
  paper_bgcolor: "rgba(0,0,0,0)",
  plot_bgcolor:  "rgba(0,0,0,0)",
  font: { family: "IBM Plex Sans", color: "#8b949e", size: 12 },
  xaxis: { gridcolor: "#21262d", zerolinecolor: "#21262d", tickfont: { color: "#8b949e" } },
  yaxis: { gridcolor: "#21262d", zerolinecolor: "#21262d", tickfont: { color: "#8b949e" } },
  legend: { bgcolor: "rgba(0,0,0,0)", font: { color: "#8b949e" } },
  margin: { l: 20, r: 20, t: 40, b: 40 },
  colorway: COLORWAY,
};

const PLOTLY_CONFIG = { displayModeBar: false, responsive: true };

// ── STATE ──────────────────────────────────────────────────────────────────
let rawDf   = [];   // full filtered data
let filtDf  = [];   // preview-page filtered subset
let currentPage = "dashboard";

// ── FORMAT HELPERS ─────────────────────────────────────────────────────────
const fmtETB = v => `ETB ${Number(v || 0).toLocaleString("en-US", { maximumFractionDigits: 0 })}`;
const fmtQty = v => Number(v || 0).toLocaleString("en-US", { maximumFractionDigits: 0 });

// ── CLASSIFY CATEGORY ──────────────────────────────────────────────────────
function classifyCategory(row) {
  const text = `${row["Material Description"] || ""} ${row["Material Group Name"] || ""}`.toLowerCase();
  for (const [cat, kws] of Object.entries(CATEGORY_KEYWORDS)) {
    if (kws.some(k => text.includes(k))) return cat;
  }
  return "Other";
}

// ── LOAD & PROCESS EXCEL ───────────────────────────────────────────────────
function loadFile(file) {
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb   = XLSX.read(new Uint8Array(e.target.result), { type: "array", cellDates: true });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(ws, { defval: "" });

      if (!data.length) { showError("The uploaded file contains no data."); return; }

      // Trim column names
      const trimmed = data.map(row => {
        const r = {};
        for (const [k, v] of Object.entries(row)) r[k.trim()] = v;
        return r;
      });

      // Validate columns
      const cols = Object.keys(trimmed[0]);
      const missing = REQUIRED_COLUMNS.filter(c => !cols.includes(c));
      if (missing.length) { showError(`Missing required columns: ${missing.join(", ")}`); return; }

      // Filter: exclude Project Stock (Q) and Non-Trade (Material starts with '4')
      let df = trimmed
        .filter(r => String(r["Special Stock Type"]).trim() !== "Q")
        .filter(r => !String(r["Material"]).startsWith("4"));

      // Numeric coercion
      const numCols = ["Unrestricted Stock","Stock in Quality Inspection","Blocked Stock","Stock in Transit",
                       "Value of Stock in Quality Inspection","Value of Stock in Transit","Value of Unrestricted Stock"];
      df.forEach(row => {
        numCols.forEach(c => { row[c] = parseFloat(row[c]) || 0; });
        // Parse dates
        const d = row["Shelf Life Expiration Date"];
        if (d instanceof Date) {
          row._expiry = d;
        } else if (d) {
          const parsed = new Date(d);
          row._expiry = isNaN(parsed) ? null : parsed;
        } else {
          row._expiry = null;
        }
        // Derived
        row["Category"]    = classifyCategory(row);
        row["Total Value"] = row["Value of Unrestricted Stock"] + row["Value of Stock in Transit"] + row["Value of Stock in Quality Inspection"];
      });

      rawDf  = df;
      filtDf = df;
      showSuccess(file.name, df.length);
      clearError();
      hideLanding();
      renderPage(currentPage);
    } catch (err) {
      showError(`Could not read Excel file: ${err.message}`);
    }
  };
  reader.readAsArrayBuffer(file);
}

// ── UI HELPERS ─────────────────────────────────────────────────────────────
function showError(msg) {
  const el = document.getElementById("errorBanner");
  el.textContent = `⚠️ ${msg}`;
  el.style.display = "block";
}
function clearError() { document.getElementById("errorBanner").style.display = "none"; }

function showSuccess(name, n) {
  const el = document.getElementById("fileStatus");
  el.style.display = "block";
  el.innerHTML = `<div class="status-ok">✓ FILE LOADED</div><div class="status-name">${name} (${n.toLocaleString()} records)</div>`;
  document.getElementById("uploadBtnText").textContent = "📂 Change File";
}

function hideLanding() { document.getElementById("landingView").style.display = "none"; }

function kpiCard(label, value, sub, color) {
  return `<div class="kpi-card ${color}">
    <div class="kpi-label">${label}</div>
    <div class="kpi-value">${value}</div>
    <div class="kpi-sub">${sub}</div>
  </div>`;
}

function setKpis(id, cards) {
  document.getElementById(id).innerHTML = cards.map(([l, v, s, c]) => kpiCard(l, v, s, c)).join("");
}

// ── GROUPBY HELPERS ────────────────────────────────────────────────────────
function groupBy(data, key, aggCols) {
  const map = {};
  data.forEach(row => {
    const k = row[key] || "";
    if (!map[k]) { map[k] = { [key]: k }; aggCols.forEach(([c]) => { map[k][c] = 0; }); }
    aggCols.forEach(([c, src]) => { map[k][c] += row[src] || 0; });
  });
  return Object.values(map);
}

function groupBy2(data, k1, k2, valCol) {
  const map = {};
  data.forEach(row => {
    const a = row[k1] || "", b = row[k2] || "";
    if (!map[a]) map[a] = {};
    map[a][b] = (map[a][b] || 0) + (row[valCol] || 0);
  });
  return map;
}

function sortBy(arr, key, asc = false) {
  return [...arr].sort((a, b) => asc ? a[key] - b[key] : b[key] - a[key]);
}

// ── TABLE BUILDER ──────────────────────────────────────────────────────────
function buildTable(rows, cols, rowClass) {
  if (!rows.length) return `<div class="alert-info">No data to display.</div>`;
  const thead = `<thead><tr>${cols.map(c => `<th>${c.label}</th>`).join("")}</tr></thead>`;
  const tbody = `<tbody>${rows.map(row => {
    const cls = rowClass ? rowClass(row) : "";
    return `<tr class="${cls}">${cols.map(c => `<td>${c.fmt ? c.fmt(row[c.key]) : (row[c.key] ?? "")}</td>`).join("")}</tr>`;
  }).join("")}</tbody>`;
  return `<div class="tbl-wrap"><table>${thead}${tbody}</table></div>`;
}

// ── CSV DOWNLOAD ───────────────────────────────────────────────────────────
function downloadCSV(data, cols, filename) {
  const header = cols.map(c => c.label).join(",");
  const rows   = data.map(row => cols.map(c => {
    let v = row[c.key] ?? "";
    if (typeof v === "string" && v.includes(",")) v = `"${v}"`;
    return v;
  }).join(","));
  const blob = new Blob([header + "\n" + rows.join("\n")], { type: "text/csv" });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement("a"); a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}

// ── PLOTLY LAYOUT MERGE ────────────────────────────────────────────────────
function pl(extra = {}) {
  return Object.assign({}, PLOTLY_LAYOUT, extra);
}

// ═══════════════════════════════════════════════════════════════════════════
// PAGE RENDERERS
// ═══════════════════════════════════════════════════════════════════════════

// ── DASHBOARD ──────────────────────────────────────────────────────────────
function renderDashboard() {
  const df = rawDf;
  const totalVal   = df.reduce((s, r) => s + r["Total Value"], 0);
  const transitVal = df.reduce((s, r) => s + r["Value of Stock in Transit"], 0);
  const qcVal      = df.reduce((s, r) => s + r["Value of Stock in Quality Inspection"], 0);
  const availVal   = df.reduce((s, r) => s + r["Value of Unrestricted Stock"], 0);

  setKpis("dash-kpis", [
    ["Total Inventory Value",   fmtETB(totalVal),   `${df.length.toLocaleString()} records`,     "blue"],
    ["Stock in Transit Value",  fmtETB(transitVal), `${fmtQty(df.reduce((s,r)=>s+r["Stock in Transit"],0))} units`, "amber"],
    ["Value in QC Inspection",  fmtETB(qcVal),      `${fmtQty(df.reduce((s,r)=>s+r["Stock in Quality Inspection"],0))} units`, "red"],
    ["Total Available Value",   fmtETB(availVal),   `${fmtQty(df.reduce((s,r)=>s+r["Unrestricted Stock"],0))} units`, "green"],
  ]);

  // Plant bar
  const plantVal = sortBy(groupBy(df, "Plant Name", [["val","Total Value"]]), "val");
  Plotly.newPlot("chart-plant-val", [{
    type: "bar", x: plantVal.map(r=>r["Plant Name"]), y: plantVal.map(r=>r.val),
    marker: { color: plantVal.map(r=>r.val), colorscale: [[0,"#1c2128"],[0.5,"#1f6feb"],[1,"#58a6ff"]], showscale: false },
    hovertemplate: "<b>%{x}</b><br>ETB %{y:,.0f}<extra></extra>",
  }], pl({ height: 280, margin: { l:20,r:20,t:20,b:80 } }), PLOTLY_CONFIG);

  // Category pie
  const catVal = groupBy(df, "Category", [["val","Total Value"]]);
  Plotly.newPlot("chart-cat-pie", [{
    type: "pie", labels: catVal.map(r=>r["Category"]), values: catVal.map(r=>r.val),
    hole: 0.55, textposition: "outside", textinfo: "percent+label",
    marker: { colors: COLORWAY },
    hovertemplate: "<b>%{label}</b><br>ETB %{value:,.0f}<br>%{percent}<extra></extra>",
  }], pl({ showlegend: false, height: 280, margin: { l:10,r:10,t:30,b:10 } }), PLOTLY_CONFIG);

  // Material group bar
  const mgVal = sortBy(groupBy(df, "Material Group Name", [["val","Total Value"]]), "val", true).slice(-15);
  Plotly.newPlot("chart-mg-bar", [{
    type: "bar", orientation: "h", y: mgVal.map(r=>r["Material Group Name"]), x: mgVal.map(r=>r.val),
    marker: { color: mgVal.map(r=>r.val), colorscale: [[0,"#1c2128"],[0.5,"#1f6feb"],[1,"#58a6ff"]], showscale: false },
    hovertemplate: "<b>%{y}</b><br>ETB %{x:,.0f}<extra></extra>",
  }], pl({ height: 420, margin: { l:10,r:20,t:10,b:20 } }), PLOTLY_CONFIG);

  // Category stacked
  const cats   = [...new Set(df.map(r=>r["Category"]))];
  const byType = { "Unrestricted": "Value of Unrestricted Stock", "Transit": "Value of Stock in Transit", "QC": "Value of Stock in Quality Inspection" };
  const colors = { "Unrestricted": "#58a6ff", "Transit": "#d29922", "QC": "#f85149" };
  const traces = Object.entries(byType).map(([name, col]) => {
    const vals = cats.map(cat => df.filter(r=>r["Category"]===cat).reduce((s,r)=>s+r[col],0));
    return { type: "bar", name, x: cats, y: vals, marker: { color: colors[name] }, hovertemplate: `<b>%{x}</b> · ${name}<br>ETB %{y:,.0f}<extra></extra>` };
  });
  Plotly.newPlot("chart-cat-breakdown", traces, pl({ barmode: "group", height: 300 }), PLOTLY_CONFIG);
}

// ── TRANSIT ────────────────────────────────────────────────────────────────
function renderTransit() {
  const df = rawDf.filter(r => r["Stock in Transit"] > 0);

  const totalTV  = df.reduce((s,r)=>s+r["Value of Stock in Transit"],0);
  const totalTQ  = df.reduce((s,r)=>s+r["Stock in Transit"],0);
  const uniqMat  = new Set(df.map(r=>r["Material"])).size;

  setKpis("transit-kpis", [
    ["Total Transit Value",          fmtETB(totalTV), "Across all plants",  "amber"],
    ["Total Transit Quantity",        fmtQty(totalTQ), "Units in movement",  "blue"],
    ["Unique Materials in Transit",   String(uniqMat), "Distinct SKUs",      "green"],
  ]);

  if (!df.length) { document.getElementById("transit-table-wrap").innerHTML = `<div class="alert-info">ℹ️ No items are currently in transit.</div>`; return; }

  // Plant bar
  const plantTV = sortBy(groupBy(df, "Plant Name", [["val","Value of Stock in Transit"]]), "val");
  Plotly.newPlot("chart-transit-plant", [{
    type: "bar", x: plantTV.map(r=>r["Plant Name"]), y: plantTV.map(r=>r.val),
    marker: { color: plantTV.map(r=>r.val), colorscale: [[0,"#1c2128"],[0.5,"#d29922"],[1,"#f0a500"]], showscale: false },
    hovertemplate: "<b>%{x}</b><br>ETB %{y:,.0f}<extra></extra>",
  }], pl({ height: 280, margin: { l:20,r:20,t:20,b:80 } }), PLOTLY_CONFIG);

  // Pie
  Plotly.newPlot("chart-transit-pie", [{
    type: "pie", labels: plantTV.map(r=>r["Plant Name"]), values: plantTV.map(r=>r.val),
    hole: 0.55, textposition: "outside", textinfo: "percent+label",
    marker: { colors: ["#d29922","#58a6ff","#3fb950","#f85149","#a371f7","#79c0ff","#e3b341"] },
    hovertemplate: "<b>%{label}</b><br>ETB %{value:,.0f}<br>%{percent}<extra></extra>",
  }], pl({ showlegend: false, height: 280, margin: { l:10,r:10,t:30,b:10 } }), PLOTLY_CONFIG);

  // Heatmap
  const heatMap  = groupBy2(df, "Material Group Name", "Plant Name", "Value of Stock in Transit");
  const mgGroups = Object.entries(
    Object.fromEntries(Object.entries(heatMap).map(([mg, plants]) => [mg, Object.values(plants).reduce((a,b)=>a+b,0)]))
  ).sort((a,b)=>b[1]-a[1]).slice(0,12).map(e=>e[0]);
  const plants   = [...new Set(df.map(r=>r["Plant Name"]))];
  const zData    = mgGroups.map(mg => plants.map(p => (heatMap[mg] && heatMap[mg][p]) || 0));
  Plotly.newPlot("chart-transit-heat", [{ type: "heatmap", z: zData, x: plants, y: mgGroups,
    colorscale: [[0,"#0d1117"],[0.3,"#1c2128"],[0.7,"#d29922"],[1,"#f0a500"]],
    hovertemplate: "Plant: <b>%{x}</b><br>Group: <b>%{y}</b><br>ETB %{z:,.0f}<extra></extra>",
  }], pl({ height: 400, margin: { l:20,r:20,t:20,b:80 } }), PLOTLY_CONFIG);

  // Table
  const transitCols = [
    { key: "Material",                      label: "Material" },
    { key: "Material Description",          label: "Description" },
    { key: "Category",                      label: "Category" },
    { key: "Plant Name",                    label: "Plant" },
    { key: "Stock in Transit",              label: "Transit Qty", fmt: fmtQty },
    { key: "Value of Stock in Transit",     label: "Transit Value (ETB)", fmt: fmtETB },
    { key: "_status",                       label: "Status" },
  ];
  const transitRows = sortBy([...df], "Value of Stock in Transit").map(r => ({
    ...r,
    _status: r["Value of Stock in Transit"] > 100000 ? "🔴 Critical" : r["Value of Stock in Transit"] > 50000 ? "🟠 High" : r["Value of Stock in Transit"] > 10000 ? "🟡 Medium" : "🟢 Low",
  }));
  document.getElementById("transit-table-wrap").innerHTML = buildTable(transitRows, transitCols);
  document.getElementById("btn-dl-transit").onclick = () => downloadCSV(transitRows, transitCols.slice(0,-1), "transit_analysis.csv");

  // HO01 section
  const ho01 = df.filter(r => String(r["Plant"]).toUpperCase() === "HO01" || String(r["Plant Name"]).toUpperCase().includes("HO01") || String(r["Plant Name"]).toUpperCase().includes("HEAD OFFICE") || String(r["Plant Name"]).toUpperCase().includes("CENTRAL"));
  if (!ho01.length) {
    document.getElementById("ho01-kpis").innerHTML = `<div class="alert-info">ℹ️ No transit records found originating from HO01 (Central).</div>`;
    return;
  }
  setKpis("ho01-kpis", [
    ["HO01 Transit Value", fmtETB(ho01.reduce((s,r)=>s+r["Value of Stock in Transit"],0)), "From Central", "blue"],
    ["HO01 Transit Qty",   fmtQty(ho01.reduce((s,r)=>s+r["Stock in Transit"],0)),           "Units dispatched", "amber"],
    ["Destinations",       String(new Set(ho01.map(r=>r["Plant Name"])).size),               "Receiving plants", "green"],
  ]);
  const ho01Cols = [
    { key: "Material", label: "Material" },
    { key: "Material Description", label: "Description" },
    { key: "Category", label: "Category" },
    { key: "Plant Name", label: "Plant" },
    { key: "Stock in Transit", label: "Transit Qty", fmt: fmtQty },
    { key: "Value of Stock in Transit", label: "Transit Value (ETB)", fmt: fmtETB },
  ];
  const ho01Sorted = sortBy([...ho01], "Value of Stock in Transit");
  document.getElementById("ho01-table-wrap").innerHTML = buildTable(ho01Sorted, ho01Cols);
  document.getElementById("btn-dl-ho01").onclick = () => downloadCSV(ho01Sorted, ho01Cols, "ho01_transit.csv");
}

// ── EXPIRY ─────────────────────────────────────────────────────────────────
function renderExpiry() {
  const today   = new Date(); today.setHours(0,0,0,0);
  const maxDate = new Date("2030-12-31");

  const valid = rawDf.filter(r => {
    const d = r._expiry;
    return d && d <= maxDate && d.getFullYear() !== 9999;
  });

  const months  = parseInt(document.querySelector('input[name="expWin"]:checked').value);
  const cutoff  = new Date(today); cutoff.setMonth(cutoff.getMonth() + months);

  const expiring = valid.filter(r => r._expiry >= today && r._expiry <= cutoff);
  const expired  = valid.filter(r => r._expiry < today);

  setKpis("expiry-kpis", [
    ["Expiring in Window", String(expiring.length), `Items within next ${months} months`, "amber"],
    ["Already Expired",    String(expired.length),  "Requires immediate action",          "red"],
    ["At-Risk Value",      fmtETB(expiring.reduce((s,r)=>s+r["Value of Unrestricted Stock"],0)), "Unrestricted stock value", "purple"],
  ]);

  // Timeline chart
  if (expiring.length) {
    const monthMap = {};
    expiring.forEach(r => {
      const key = `${r._expiry.getFullYear()}-${String(r._expiry.getMonth()+1).padStart(2,"0")}`;
      monthMap[key] = (monthMap[key] || 0) + 1;
    });
    const months2 = Object.keys(monthMap).sort();
    Plotly.newPlot("chart-expiry-timeline", [{
      type: "bar", x: months2, y: months2.map(m => monthMap[m]),
      marker: { color: months2.map(m => monthMap[m]), colorscale: [[0,"#1c2128"],[0.5,"#d29922"],[1,"#f85149"]], showscale: false },
      hovertemplate: "<b>%{x}</b><br>%{y} items<extra></extra>",
    }], pl({ height: 260, margin: { l:20,r:20,t:20,b:60 } }), PLOTLY_CONFIG);
  } else {
    document.getElementById("chart-expiry-timeline").innerHTML = "";
  }

  // Expiry table
  const expCols = [
    { key: "Material", label: "Material" },
    { key: "Material Description", label: "Description" },
    { key: "Plant Name", label: "Plant" },
    { key: "_expiryStr", label: "Expiry Date" },
    { key: "Unrestricted Stock", label: "Qty", fmt: fmtQty },
    { key: "Value of Unrestricted Stock", label: "Value (ETB)", fmt: fmtETB },
    { key: "Category", label: "Category" },
    { key: "_daysLeft", label: "Days Until Expiry" },
  ];
  const expRows = sortBy([...expiring], "_daysLeft", true).map(r => ({
    ...r,
    _expiryStr: r._expiry ? r._expiry.toISOString().slice(0,10) : "",
    _daysLeft:  r._expiry ? Math.floor((r._expiry - today) / 86400000) : 9999,
  }));
  const expRowClass = row => row._daysLeft <= 30 ? "row-red" : row._daysLeft <= 90 ? "row-amber" : "";

  document.getElementById("expiry-table-wrap").innerHTML = expRows.length
    ? buildTable(expRows, expCols, expRowClass)
    : `<div class="alert-info">✓ No items expiring within the selected window.</div>`;

  document.getElementById("btn-dl-expiry").onclick = () => downloadCSV(expRows, expCols, `expiry_watchlist_${months}months.csv`);

  // Expired section
  if (expired.length) {
    document.getElementById("expired-section").style.display = "block";
    document.getElementById("expired-header").textContent = `🔴 Already Expired Items (${expired.length})`;
    const expiredCols = [
      { key: "Material", label: "Material" },
      { key: "Material Description", label: "Description" },
      { key: "Plant Name", label: "Plant" },
      { key: "_expiryStr", label: "Expiry Date" },
      { key: "Unrestricted Stock", label: "Qty", fmt: fmtQty },
    ];
    const expiredRows = sortBy([...expired], "_expiry2", true).map(r => ({ ...r, _expiryStr: r._expiry ? r._expiry.toISOString().slice(0,10) : "" }));
    document.getElementById("expired-table-wrap").innerHTML = buildTable(expiredRows, expiredCols);
    document.getElementById("btn-dl-expired").onclick = () => downloadCSV(expiredRows, expiredCols, "expired_items.csv");
  } else {
    document.getElementById("expired-section").style.display = "none";
  }
}

// ── QC ─────────────────────────────────────────────────────────────────────
function renderQC() {
  const df = rawDf.filter(r => r["Stock in Quality Inspection"] > 0);

  const totalQCVal = df.reduce((s,r)=>s+r["Value of Stock in Quality Inspection"],0);
  const totalQCQty = df.reduce((s,r)=>s+r["Stock in Quality Inspection"],0);
  const uniqQC     = new Set(df.map(r=>r["Material"])).size;

  setKpis("qc-kpis", [
    ["Total Value Stuck in QC",     fmtETB(totalQCVal), "Across all plants",       "red"],
    ["Total QC Quantity",            fmtQty(totalQCQty), "Units under inspection",  "amber"],
    ["Unique Materials in QC",       String(uniqQC),     "Distinct SKUs",           "blue"],
  ]);

  if (!df.length) { document.getElementById("qc-table-wrap").innerHTML = `<div class="alert-info">✓ No items currently in quality inspection.</div>`; return; }

  // Plant bar
  const plantQC = sortBy(groupBy(df, "Plant Name", [["val","Value of Stock in Quality Inspection"]]), "val");
  Plotly.newPlot("chart-qc-plant", [{
    type: "bar", x: plantQC.map(r=>r["Plant Name"]), y: plantQC.map(r=>r.val),
    marker: { color: plantQC.map(r=>r.val), colorscale: [[0,"#1c2128"],[0.5,"#f85149"],[1,"#ff7b72"]], showscale: false },
    hovertemplate: "<b>%{x}</b><br>ETB %{y:,.0f}<extra></extra>",
  }], pl({ height: 280, margin: { l:20,r:20,t:20,b:80 } }), PLOTLY_CONFIG);

  // MG horizontal bar
  const mgQC = sortBy(groupBy(df, "Material Group Name", [["val","Value of Stock in Quality Inspection"]]), "val", true).slice(-12);
  Plotly.newPlot("chart-qc-mg", [{
    type: "bar", orientation: "h", y: mgQC.map(r=>r["Material Group Name"]), x: mgQC.map(r=>r.val),
    marker: { color: mgQC.map(r=>r.val), colorscale: [[0,"#1c2128"],[0.5,"#f85149"],[1,"#ff7b72"]], showscale: false },
    hovertemplate: "<b>%{y}</b><br>ETB %{x:,.0f}<extra></extra>",
  }], pl({ height: 280, margin: { l:10,r:20,t:10,b:20 } }), PLOTLY_CONFIG);

  // Table
  const qcCols = [
    { key: "Material", label: "Material" },
    { key: "Material Description", label: "Description" },
    { key: "Category", label: "Category" },
    { key: "Plant Name", label: "Plant" },
    { key: "Stock in Quality Inspection", label: "QC Qty", fmt: fmtQty },
    { key: "Value of Stock in Quality Inspection", label: "QC Value (ETB)", fmt: fmtETB },
  ];
  const qcRows = sortBy([...df], "Value of Stock in Quality Inspection");
  const qcRowClass = row => row["Value of Stock in Quality Inspection"] > 10000 ? "row-red" : "";
  document.getElementById("qc-table-wrap").innerHTML = buildTable(qcRows, qcCols, qcRowClass);
  document.getElementById("btn-dl-qc").onclick = () => downloadCSV(qcRows, qcCols, "qc_inspection.csv");
}

// ── BRANCH COMPARISON ───────────────────────────────────────────────────────
function renderBranch() {
  const df = rawDf;
  const plants = [...new Set(df.map(r=>r["Plant"]).map(v=>String(v).toUpperCase()))];
  let centralCode, centralName;

  if (plants.includes("HO01")) {
    centralCode = "HO01";
    centralName = df.find(r=>String(r["Plant"]).toUpperCase()==="HO01")?.["Plant Name"] || "HO01";
    document.getElementById("branch-central-info").style.display = "none";
  } else {
    // Fallback: highest total value plant
    const totals = {};
    df.forEach(r => { const p = r["Plant Name"]; totals[p] = (totals[p]||0) + r["Total Value"]; });
    centralName = Object.entries(totals).sort((a,b)=>b[1]-a[1])[0]?.[0] || "";
    document.getElementById("branch-central-info").style.display = "block";
    document.getElementById("branch-central-info").innerHTML = `ℹ️ HO01 not found — using <b>${centralName}</b> as central branch (highest inventory value).`;
  }

  // Aggregate
  const aggMap = {};
  df.forEach(r => {
    const k = r["Plant Name"];
    if (!aggMap[k]) aggMap[k] = { "Plant Name": k, Plant: r["Plant"], Total_Value: 0, Unrestricted: 0, Transit: 0, QC: 0, Items: 0 };
    aggMap[k].Total_Value  += r["Total Value"];
    aggMap[k].Unrestricted += r["Value of Unrestricted Stock"];
    aggMap[k].Transit      += r["Value of Stock in Transit"];
    aggMap[k].QC           += r["Value of Stock in Quality Inspection"];
    aggMap[k].Items++;
  });
  const branchAgg = Object.values(aggMap);
  const allBranches = branchAgg.map(r=>r["Plant Name"]);
  const others = allBranches.filter(b => b !== centralName);

  // Populate select
  const sel = document.getElementById("branch-select");
  sel.innerHTML = "";
  others.forEach(b => {
    const opt = document.createElement("option");
    opt.value = b; opt.textContent = b; opt.selected = true;
    sel.appendChild(opt);
  });

  function updateBranchCharts() {
    const selected = [...sel.selectedOptions].map(o=>o.value);
    if (!selected.length) {
      document.getElementById("branch-table-wrap").innerHTML = `<div class="alert-warning">⚠️ Please select at least one branch to compare.</div>`;
      return;
    }
    const compareNames = [centralName, ...selected];
    const compareDf    = branchAgg.filter(r => compareNames.includes(r["Plant Name"]));

    // Table
    const bCols = [
      { key: "Plant Name", label: "Plant Name" },
      { key: "Total_Value", label: "Total Value (ETB)", fmt: fmtETB },
      { key: "Unrestricted", label: "Unrestricted (ETB)", fmt: fmtETB },
      { key: "Transit", label: "Transit (ETB)", fmt: fmtETB },
      { key: "QC", label: "QC (ETB)", fmt: fmtETB },
      { key: "Items", label: "# Items" },
    ];
    document.getElementById("branch-table-wrap").innerHTML = buildTable(compareDf, bCols, r => r["Plant Name"] === centralName ? "row-blue" : "");

    // Grouped bar
    const types = ["Unrestricted","Transit","QC"];
    const colors2 = { Unrestricted: "#58a6ff", Transit: "#d29922", QC: "#f85149" };
    const traces = types.map(t => ({
      type: "bar", name: t,
      x: compareDf.map(r=>r["Plant Name"]),
      y: compareDf.map(r=>r[t]),
      marker: { color: colors2[t] },
      hovertemplate: `<b>%{x}</b> · ${t}<br>ETB %{y:,.0f}<extra></extra>`,
    }));
    Plotly.newPlot("chart-branch-grouped", traces, pl({ barmode: "group", height: 300 }), PLOTLY_CONFIG);

    // Scatter
    Plotly.newPlot("chart-branch-scatter", [{
      type: "scatter", mode: "markers+text",
      x: compareDf.map(r=>r.Total_Value), y: compareDf.map(r=>r.Transit),
      text: compareDf.map(r=>r["Plant Name"]),
      textposition: "top center",
      marker: { size: compareDf.map(r=>Math.max(8, Math.sqrt(r.Items)*2)), color: COLORWAY },
      hovertemplate: "<b>%{text}</b><br>Total: ETB %{x:,.0f}<br>Transit: ETB %{y:,.0f}<extra></extra>",
    }], pl({ height: 320, xaxis: { ...PLOTLY_LAYOUT.xaxis, title: "Total Inventory Value (ETB)" }, yaxis: { ...PLOTLY_LAYOUT.yaxis, title: "Transit Value (ETB)" } }), PLOTLY_CONFIG);
  }

  sel.addEventListener("change", updateBranchCharts);
  updateBranchCharts();
}

// ── SUPPLY CHAIN FLOW ───────────────────────────────────────────────────────
function renderFlow() {
  const df = rawDf;
  const totalVal    = df.reduce((s,r)=>s+r["Total Value"],0);
  const transitVal  = df.reduce((s,r)=>s+r["Value of Stock in Transit"],0);
  const qcVal       = df.reduce((s,r)=>s+r["Value of Stock in Quality Inspection"],0);
  const availVal    = df.reduce((s,r)=>s+r["Value of Unrestricted Stock"],0);

  // Sankey
  Plotly.newPlot("chart-sankey", [{
    type: "sankey",
    node: {
      pad: 20, thickness: 24,
      line: { color: "#21262d", width: 0.5 },
      label: ["Total Inventory","In Transit","In QC Inspection","Available (Unrestricted)","Released to Market"],
      color: ["#58a6ff","#d29922","#f85149","#3fb950","#a371f7"],
      hovertemplate: "%{label}<extra></extra>",
    },
    link: {
      source: [0, 0, 0, 1, 2],
      target: [1, 2, 3, 4, 4],
      value:  [
        transitVal  || 1,
        qcVal       || 1,
        availVal    || 1,
        (transitVal || 1) * 0.85,
        (qcVal      || 1) * 0.80,
      ],
      color: ["rgba(210,153,34,0.35)","rgba(248,81,73,0.35)","rgba(63,185,80,0.35)","rgba(210,153,34,0.25)","rgba(248,81,73,0.25)"],
      hovertemplate: "ETB %{value:,.0f}<extra></extra>",
    },
  }], pl({ height: 480, margin: { l:20,r:20,t:40,b:20 } }), PLOTLY_CONFIG);

  const pct = v => totalVal ? `${(v/totalVal*100).toFixed(1)}% of total` : "N/A";
  setKpis("flow-kpis", [
    ["📦 Total Inventory",  fmtETB(totalVal),   "All stock (excl. blocked)",   "blue"],
    ["🚚 In Transit",       fmtETB(transitVal), pct(transitVal),               "amber"],
    ["🔬 In QC Inspection", fmtETB(qcVal),      pct(qcVal),                    "red"],
    ["✅ Available Stock",  fmtETB(availVal),   pct(availVal),                 "green"],
  ]);

  // Per-plant flow
  const plantFlow = {};
  df.forEach(r => {
    const p = r["Plant Name"];
    if (!plantFlow[p]) plantFlow[p] = { Unrestricted: 0, Transit: 0, QC: 0 };
    plantFlow[p].Unrestricted += r["Value of Unrestricted Stock"];
    plantFlow[p].Transit      += r["Value of Stock in Transit"];
    plantFlow[p].QC           += r["Value of Stock in Quality Inspection"];
  });
  const pfPlants = Object.keys(plantFlow).sort((a,b) => plantFlow[b].Unrestricted - plantFlow[a].Unrestricted);
  const flowColors = { Unrestricted: "#3fb950", Transit: "#d29922", QC: "#f85149" };
  const flowTraces = ["Unrestricted","Transit","QC"].map(t => ({
    type: "bar", name: t,
    x: pfPlants,
    y: pfPlants.map(p=>plantFlow[p][t]),
    marker: { color: flowColors[t] },
    hovertemplate: `<b>%{x}</b> · ${t}<br>ETB %{y:,.0f}<extra></extra>`,
  }));
  Plotly.newPlot("chart-flow-plant", flowTraces, pl({ barmode: "stack", height: 320 }), PLOTLY_CONFIG);
}

// ── DATA PREVIEW ────────────────────────────────────────────────────────────
function populatePreviewFilters() {
  const df = rawDf;
  function populate(id, key) {
    const sel = document.getElementById(id);
    const vals = [...new Set(df.map(r=>r[key]))].filter(Boolean).sort();
    sel.innerHTML = vals.map(v => `<option value="${v}">${v}</option>`).join("");
  }
  populate("filter-plant", "Plant Name");
  populate("filter-cat",   "Category");
  populate("filter-mg",    "Material Group Name");
}

function applyPreviewFilters() {
  const getSelected = id => [...document.querySelectorAll(`#${id} option:checked`)].map(o=>o.value);
  const plants = getSelected("filter-plant");
  const cats   = getSelected("filter-cat");
  const mgs    = getSelected("filter-mg");
  filtDf = rawDf.filter(r =>
    (!plants.length || plants.includes(r["Plant Name"])) &&
    (!cats.length   || cats.includes(r["Category"])) &&
    (!mgs.length    || mgs.includes(r["Material Group Name"]))
  );
  renderPreviewTable();
}

function renderPreviewTable() {
  const df = filtDf;
  setKpis("preview-kpis", [
    ["Total Records",     df.length.toLocaleString(),                               "After filtering",         "blue"],
    ["Unique Materials",  new Set(df.map(r=>r["Material"])).size.toLocaleString(),  "Distinct SKUs",           "green"],
    ["Total Plants",      new Set(df.map(r=>r["Plant"])).size.toLocaleString(),     "Stocking locations",      "amber"],
    ["Material Groups",   new Set(df.map(r=>r["Material Group Name"])).size.toLocaleString(), "Therapeutic categories", "purple"],
  ]);

  document.getElementById("preview-count").innerHTML = `Showing <b>${df.length.toLocaleString()}</b> of <b>${rawDf.length.toLocaleString()}</b> records`;

  const cols = [
    { key: "Material",                          label: "Material" },
    { key: "Material Description",              label: "Description" },
    { key: "Plant Name",                        label: "Plant" },
    { key: "Category",                          label: "Category" },
    { key: "Material Group Name",               label: "Material Group" },
    { key: "Unrestricted Stock",                label: "Available Qty",    fmt: fmtQty },
    { key: "Stock in Transit",                  label: "Transit Qty",      fmt: fmtQty },
    { key: "Stock in Quality Inspection",       label: "QC Qty",           fmt: fmtQty },
    { key: "Value of Unrestricted Stock",       label: "Unrestricted (ETB)", fmt: fmtETB },
    { key: "Value of Stock in Transit",         label: "Transit (ETB)",    fmt: fmtETB },
    { key: "Value of Stock in Quality Inspection", label: "QC (ETB)",      fmt: fmtETB },
    { key: "Total Value",                       label: "Total Value (ETB)", fmt: fmtETB },
    { key: "_expiryStr",                        label: "Expiry Date" },
  ];
  const rows = df.slice(0, 500).map(r => ({ ...r, _expiryStr: r._expiry ? r._expiry.toISOString().slice(0,10) : "" }));
  document.getElementById("preview-table-wrap").innerHTML = buildTable(rows, cols);
  if (df.length > 500) {
    document.getElementById("preview-table-wrap").insertAdjacentHTML("afterend",
      `<div class="alert-info">Showing first 500 of ${df.length.toLocaleString()} records. Download CSV for full data.</div>`);
  }
  document.getElementById("btn-dl-preview").onclick = () => downloadCSV(df.map(r=>({...r, _expiryStr: r._expiry?r._expiry.toISOString().slice(0,10):""})), cols, "pharma_inventory_filtered.csv");
}

function renderPreview() {
  populatePreviewFilters();
  filtDf = rawDf;
  renderPreviewTable();
}

// ═══════════════════════════════════════════════════════════════════════════
// PAGE SWITCHING
// ═══════════════════════════════════════════════════════════════════════════
const PAGE_RENDERERS = {
  dashboard: renderDashboard,
  transit:   renderTransit,
  expiry:    renderExpiry,
  qc:        renderQC,
  branch:    renderBranch,
  flow:      renderFlow,
  preview:   renderPreview,
};

function renderPage(id) {
  if (!rawDf.length) return;
  currentPage = id;

  document.querySelectorAll(".page").forEach(el => { el.style.display = "none"; });
  const pg = document.getElementById(`page-${id}`);
  if (pg) pg.style.display = "block";

  document.querySelectorAll(".nav-btn").forEach(btn => {
    btn.classList.toggle("active", btn.dataset.page === id);
  });

  try { PAGE_RENDERERS[id]?.(); }
  catch(e) { console.error(`Error rendering ${id}:`, e); }
}

// ═══════════════════════════════════════════════════════════════════════════
// EVENT LISTENERS
// ═══════════════════════════════════════════════════════════════════════════
document.addEventListener("DOMContentLoaded", () => {
  // Nav
  document.querySelectorAll(".nav-btn").forEach(btn => {
    btn.addEventListener("click", () => renderPage(btn.dataset.page));
  });

  // File upload
  document.getElementById("fileInput").addEventListener("change", e => {
    const f = e.target.files[0];
    if (f) loadFile(f);
  });

  // Expiry window radio
  document.getElementById("expiry-window-group").addEventListener("change", () => {
    if (rawDf.length && currentPage === "expiry") renderExpiry();
  });

  // Preview filters
  document.getElementById("btn-apply-filter").addEventListener("click", applyPreviewFilters);
  document.getElementById("btn-clear-filter").addEventListener("click", () => {
    document.querySelectorAll("#filter-plant option, #filter-cat option, #filter-mg option").forEach(o => { o.selected = false; });
    filtDf = rawDf;
    renderPreviewTable();
  });
});
