// =============================================================================
// PharmaTrack v2 — Pharmaceutical Inventory Management System
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

const COLORWAY = ["#58a6ff","#3fb950","#d29922","#f85149","#a371f7","#79c0ff","#56d364","#e3b341","#ff7b72","#d2a8ff","#ffa657","#70d9a0"];

// NOTE: Exclusion rules (isNonMedicalCode, isNonMedicalGroup) are loaded from
// filters.js which MUST be included before this script in the HTML.
const PLOTLY_LAYOUT = {
  paper_bgcolor: "rgba(0,0,0,0)", plot_bgcolor: "rgba(0,0,0,0)",
  font: { family: "IBM Plex Sans", color: "#8b949e", size: 12 },
  xaxis: { gridcolor: "#21262d", zerolinecolor: "#21262d", tickfont: { color: "#8b949e" } },
  yaxis: { gridcolor: "#21262d", zerolinecolor: "#21262d", tickfont: { color: "#8b949e" } },
  legend: { bgcolor: "rgba(0,0,0,0)", font: { color: "#8b949e" } },
  margin: { l: 20, r: 20, t: 40, b: 40 },
  colorway: COLORWAY,
};
const PLOTLY_CONFIG = { displayModeBar: false, responsive: true };

// ── STATE ──────────────────────────────────────────────────────────────────
let rawDf  = [];
let filtDf = [];
let currentPage = "dashboard";

// Page-level filter state
const pageFilters = {
  dashboard: { plant: "", mg: "" },
  transit:   { plant: "", mg: "" },
  expiry:    { plant: "", mg: "" },
  qc:        { plant: "", mg: "" },
  branch:    { mg: "" },
  flow:      { plant: "", mg: "" },
};

// ── RECONCILIATION STATE ───────────────────────────────────────────────────
let reconcileGroups = []; // [{name, codes:[]}]
let reconcilePending = []; // [{code, desc}]

// ── FORMAT HELPERS ─────────────────────────────────────────────────────────
const fmtETB = v => `ETB ${Number(v || 0).toLocaleString("en-US", { maximumFractionDigits: 0 })}`;
const fmtQty = v => Number(v || 0).toLocaleString("en-US", { maximumFractionDigits: 0 });

// ── HTML ESCAPE (used by buildTable and reconciliation UI) ──────────────────
function escHtml(str) {
  return String(str ?? "").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}

// ── LOAD & PROCESS EXCEL ───────────────────────────────────────────────────
function loadFile(file) {
  // Show loading state immediately before the synchronous XLSX parse blocks the UI
  const statusEl = document.getElementById("fileStatus");
  statusEl.style.display = "block";
  statusEl.innerHTML = `<div class="status-ok">⏳ LOADING…</div><div class="status-name">Parsing ${file.name}</div>`;

  const reader = new FileReader();
  reader.onload = e => {
    // Defer heavy work one tick so the loading message renders first
    setTimeout(() => {
      try {
      const wb   = XLSX.read(new Uint8Array(e.target.result), { type: "array", cellDates: true });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(ws, { defval: "" });
      if (!data.length) { showError("The uploaded file contains no data."); return; }

      const trimmed = data.map(row => {
        const r = {};
        for (const [k, v] of Object.entries(row)) r[k.trim()] = v;
        return r;
      });

      const cols = Object.keys(trimmed[0]);
      const missing = REQUIRED_COLUMNS.filter(c => !cols.includes(c));
      if (missing.length) { showError(`Missing columns: ${missing.join(", ")}`); return; }

      let df = trimmed
        .filter(r => String(r["Special Stock Type"]).trim() !== "Q")
        .filter(r => !isNonMedicalCode(r["Material"]))
        .filter(r => !isNonMedicalGroup(r["Material Group Name"]));

      const numCols = ["Unrestricted Stock","Stock in Quality Inspection","Blocked Stock","Stock in Transit",
                       "Value of Stock in Quality Inspection","Value of Stock in Transit","Value of Unrestricted Stock"];
      df.forEach(row => {
        numCols.forEach(c => { row[c] = parseFloat(row[c]) || 0; });
        const d = row["Shelf Life Expiration Date"];
        if (d instanceof Date) row._expiry = d;
        else if (d) { const p = new Date(d); row._expiry = isNaN(p) ? null : p; }
        else row._expiry = null;
        row["Total Value"] = row["Value of Unrestricted Stock"] + row["Value of Stock in Transit"] + row["Value of Stock in Quality Inspection"];
        row["Total Qty"]   = row["Unrestricted Stock"] + row["Stock in Transit"] + row["Stock in Quality Inspection"];
      });

      // Exclude any row where ALL stock quantities are zero —
      // this covers: expired items with zero qty, unrestricted-only items
      // with zero qty, and any other stock-type combination that nets to zero.
      df = df.filter(r =>
        r["Unrestricted Stock"] > 0 ||
        r["Stock in Transit"] > 0 ||
        r["Stock in Quality Inspection"] > 0 ||
        r["Blocked Stock"] > 0
      );

      rawDf  = df;
      filtDf = df;
      showSuccess(file.name, df.length);
      clearError();
      hideLanding();
      populateAllFilters();
      renderPage(currentPage);
      } catch (err) { showError(`Could not read Excel file: ${err.message}`); }
    }, 30); // 30 ms — enough for the browser to paint the loading state
  };
  reader.readAsArrayBuffer(file);
}

// ── POPULATE FILTER DROPDOWNS ──────────────────────────────────────────────
function populateAllFilters() {
  const plants = [...new Set(rawDf.map(r=>r["Plant Name"]))].filter(Boolean).sort();
  // Explicitly exclude non-medical groups from dropdowns — mirrors the exclusion
  // applied during data load, so the filter never surfaces excluded categories.
  const mgs    = [...new Set(rawDf.map(r=>r["Material Group Name"]))]
    .filter(Boolean)
    .filter(name => !isNonMedicalGroup(name))
    .sort();

  const plantSelectors = ["dash-filter-plant","transit-filter-plant","expiry-filter-plant","qc-filter-plant","flow-filter-plant","filter-plant"];
  const mgSelectors    = ["dash-filter-mg","transit-filter-mg","expiry-filter-mg","qc-filter-mg","branch-filter-mg","flow-filter-mg","filter-mg"];
  const mgNameSelectors = ["filter-mgname"];

  plantSelectors.forEach(id => {
    const el = document.getElementById(id);
    if (!el) return;
    el.innerHTML = `<option value="">All Plants</option>` + plants.map(p=>`<option value="${p}">${p}</option>`).join("");
  });
  mgSelectors.forEach(id => {
    const el = document.getElementById(id);
    if (!el) return;
    el.innerHTML = `<option value="">All Material Groups</option>` + mgs.map(m=>`<option value="${m}">${m}</option>`).join("");
  });
  mgNameSelectors.forEach(id => {
    const el = document.getElementById(id);
    if (!el) return;
    const vals = [...new Set(rawDf.map(r=>r["Material Group Name"]))]
      .filter(Boolean)
      .filter(name => !isNonMedicalGroup(name))
      .sort();
    el.innerHTML = vals.map(v=>`<option value="${v}">${v}</option>`).join("");
  });
}

// ── APPLY PAGE FILTER ──────────────────────────────────────────────────────
function applyPageFilter(page) {
  const f = pageFilters[page] || {};
  const filtered = rawDf.filter(r =>
    (!f.plant || r["Plant Name"] === f.plant) &&
    (!f.mg    || r["Material Group Name"] === f.mg)
  );
  return applyReconciliationToData(filtered);
}

// ── UI HELPERS ─────────────────────────────────────────────────────────────
function showError(msg) { const el = document.getElementById("errorBanner"); el.textContent = `⚠️ ${msg}`; el.style.display = "block"; }
function clearError() { document.getElementById("errorBanner").style.display = "none"; }
function showSuccess(name, n) {
  const el = document.getElementById("fileStatus"); el.style.display = "block";
  el.innerHTML = `<div class="status-ok">✓ FILE LOADED</div><div class="status-name">${name} (${n.toLocaleString()} records)</div>`;
  document.getElementById("uploadBtnText").textContent = "📂 Change File";
}
function hideLanding() { document.getElementById("landingView").style.display = "none"; }

function kpiCard(label, value, sub, color) {
  return `<div class="kpi-card ${color}"><div class="kpi-label">${label}</div><div class="kpi-value">${value}</div><div class="kpi-sub">${sub}</div></div>`;
}
function setKpis(id, cards) { document.getElementById(id).innerHTML = cards.map(([l,v,s,c])=>kpiCard(l,v,s,c)).join(""); }

// ── GROUPBY HELPERS ────────────────────────────────────────────────────────
function groupBy(data, key, aggCols) {
  const map = {};
  data.forEach(row => {
    const k = row[key] || "";
    if (!map[k]) { map[k] = { [key]: k }; aggCols.forEach(([c])=>{ map[k][c] = 0; }); }
    aggCols.forEach(([c,src])=>{ map[k][c] += row[src]||0; });
  });
  return Object.values(map);
}
function sortBy(arr, key, asc=false) { return [...arr].sort((a,b)=>asc ? a[key]-b[key] : b[key]-a[key]); }

// ── TABLE BUILDER ──────────────────────────────────────────────────────────
// Columns with raw:true may contain trusted HTML (badges etc.) — all others
// are escaped to prevent XSS from Excel data landing in the DOM.
function buildTable(rows, cols, rowClass, extraClass="") {
  if (!rows.length) return `<div class="alert-info">No data to display.</div>`;
  const thead = `<thead><tr>${cols.map(c=>`<th>${escHtml(c.label)}</th>`).join("")}</tr></thead>`;
  const tbody = `<tbody>${rows.map(row=>{
    const cls = rowClass ? rowClass(row) : "";
    return `<tr class="${cls}">${cols.map(c=>{
      const raw = c.fmt ? c.fmt(row[c.key]) : (row[c.key]??"");
      const val = c.raw ? raw : escHtml(String(raw));
      const cellCls = c.cellClass || "";
      return `<td class="${cellCls}">${val}</td>`;
    }).join("")}</tr>`;
  }).join("")}</tbody>`;
  return `<div class="tbl-wrap"><table class="${extraClass}">${thead}${tbody}</table></div>`;
}

// ── EXCEL DOWNLOAD ─────────────────────────────────────────────────────────
function downloadExcel(data, cols, filename) {
  const header = cols.map(c=>c.label);
  const rows   = data.map(row => cols.map(c => {
    const v = row[c.key];
    // Use rawKey value when available for numeric fidelity
    const raw = c.rawKey ? (row[c.rawKey] ?? v) : v;
    // For formatted columns return the raw number; for plain text return the value
    // Use "" fallback (not 0) so text fields don't become 0 in Excel
    if (c.fmt) return (typeof raw === "number") ? raw : (raw ?? "");
    return raw ?? "";
  }));
  const wsData = [header, ...rows];
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Data");
  XLSX.writeFile(wb, filename);
}

// ── CSV DOWNLOAD ───────────────────────────────────────────────────────────
function downloadCSV(data, cols, filename) {
  const header = cols.map(c=>c.label).join(",");
  const rows   = data.map(row=>cols.map(c=>{
    let v = c.rawKey ? (row[c.rawKey]??row[c.key]??"") : (row[c.key]??"");
    v = String(v ?? "");
    // FIX 2: CSV injection — prefix dangerous leading chars with single quote
    if (/^[=+\-@\t\r]/.test(v)) v = `'${v}`;
    // Wrap in quotes if contains comma, quote or newline
    if (v.includes(",") || v.includes('"') || v.includes("\n")) {
      v = `"${v.replace(/"/g, '""')}"`;
    }
    return v;
  }).join(","));
  const blob = new Blob(["\uFEFF"+header+"\n"+rows.join("\n")],{type:"text/csv;charset=utf-8"});
  const url  = URL.createObjectURL(blob);
  const a = document.createElement("a"); a.href=url; a.download=filename; a.click();
  URL.revokeObjectURL(url);
}

// ── PLOTLY LAYOUT MERGE ────────────────────────────────────────────────────
// Deep-merge nested axis objects so callers can override xaxis/yaxis properties
// without clobbering the shared PLOTLY_LAYOUT reference (Object.assign is shallow).
function pl(extra={}) {
  return Object.assign({}, PLOTLY_LAYOUT, extra, {
    xaxis:  Object.assign({}, PLOTLY_LAYOUT.xaxis,  extra.xaxis  || {}),
    yaxis:  Object.assign({}, PLOTLY_LAYOUT.yaxis,  extra.yaxis  || {}),
    legend: Object.assign({}, PLOTLY_LAYOUT.legend, extra.legend || {}),
    margin: Object.assign({}, PLOTLY_LAYOUT.margin, extra.margin || {}),
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// DASHBOARD
// ═══════════════════════════════════════════════════════════════════════════
function renderDashboard() {
  const pf = pageFilters.dashboard;
  document.getElementById("dash-filter-plant").value = pf.plant||"";
  document.getElementById("dash-filter-mg").value    = pf.mg||"";
  const df = applyPageFilter("dashboard");

  const totalVal   = df.reduce((s,r)=>s+r["Total Value"],0);
  const transitVal = df.reduce((s,r)=>s+r["Value of Stock in Transit"],0);
  const qcVal      = df.reduce((s,r)=>s+r["Value of Stock in Quality Inspection"],0);
  const availVal   = df.reduce((s,r)=>s+r["Value of Unrestricted Stock"],0);
  const totalQty   = df.reduce((s,r)=>s+r["Total Qty"],0);

  setKpis("dash-kpis",[
    ["Total Inventory Value",   fmtETB(totalVal),   `${fmtQty(totalQty)} total units`,     "blue"],
    ["Stock in Transit Value",  fmtETB(transitVal), `${fmtQty(df.reduce((s,r)=>s+r["Stock in Transit"],0))} units`, "amber"],
    ["Value in QC",             fmtETB(qcVal),      `${fmtQty(df.reduce((s,r)=>s+r["Stock in Quality Inspection"],0))} units`, "red"],
    ["Available (Unrestricted)",fmtETB(availVal),   `${fmtQty(df.reduce((s,r)=>s+r["Unrestricted Stock"],0))} units`, "green"],
    ["Unique Materials",        new Set(df.map(r=>r["Material"])).size.toLocaleString(), `${new Set(df.map(r=>r["Plant"])).size} plants`, "purple"],
  ]);

  // Plant bar — dual axis qty+value
  const plantAgg = sortBy(groupBy(df,"Plant Name",[["val","Total Value"],["qty","Total Qty"]]),"val");
  Plotly.newPlot("chart-plant-val",[
    { type:"bar", name:"Value (ETB)", x:plantAgg.map(r=>r["Plant Name"]), y:plantAgg.map(r=>r.val), yaxis:"y", marker:{color:"#58a6ff"}, hovertemplate:"<b>%{x}</b><br>ETB %{y:,.0f}<extra></extra>" },
    { type:"scatter", mode:"lines+markers", name:"Quantity", x:plantAgg.map(r=>r["Plant Name"]), y:plantAgg.map(r=>r.qty), yaxis:"y2", marker:{color:"#3fb950",size:8}, line:{color:"#3fb950"}, hovertemplate:"<b>%{x}</b><br>Qty: %{y:,.0f}<extra></extra>" },
  ], pl({ height:280, margin:{l:20,r:60,t:20,b:80}, yaxis2:{overlaying:"y",side:"right",gridcolor:"transparent",tickfont:{color:"#3fb950"},title:{text:"Qty",font:{color:"#3fb950"}}}, barmode:"group" }), PLOTLY_CONFIG);

  // Material Group pie (by value)
  const mgAgg = sortBy(groupBy(df,"Material Group Name",[["val","Total Value"]]),"val").slice(0,12);
  Plotly.newPlot("chart-cat-pie",[{
    type:"pie", labels:mgAgg.map(r=>r["Material Group Name"]), values:mgAgg.map(r=>r.val),
    hole:0.55, textposition:"outside", textinfo:"percent+label",
    marker:{colors:COLORWAY}, hovertemplate:"<b>%{label}</b><br>ETB %{value:,.0f}<br>%{percent}<extra></extra>",
  }], pl({ showlegend:false, height:280, margin:{l:10,r:10,t:30,b:10} }), PLOTLY_CONFIG);

  // MG bar top 15
  const mgVal = sortBy(groupBy(df,"Material Group Name",[["val","Total Value"],["qty","Total Qty"]]),"val",true).slice(-15);
  Plotly.newPlot("chart-mg-bar",[
    { type:"bar", orientation:"h", name:"Value (ETB)", y:mgVal.map(r=>r["Material Group Name"]), x:mgVal.map(r=>r.val), marker:{color:"#58a6ff"}, hovertemplate:"<b>%{y}</b><br>ETB %{x:,.0f}<extra></extra>" },
    { type:"bar", orientation:"h", name:"Quantity", y:mgVal.map(r=>r["Material Group Name"]), x:mgVal.map(r=>r.qty), marker:{color:"#3fb950"}, hovertemplate:"<b>%{y}</b><br>Qty: %{x:,.0f}<extra></extra>" },
  ], pl({ barmode:"group", height:420, margin:{l:10,r:20,t:10,b:20} }), PLOTLY_CONFIG);

  // Download handlers
  const dlCols=[
    {key:"Plant Name",label:"Plant"},
    {key:"Material Group Name",label:"Material Group"},
    {key:"Total Value",label:"Total Value (ETB)",fmt:fmtETB,rawKey:"Total Value"},
    {key:"Total Qty",label:"Total Qty",fmt:fmtQty,rawKey:"Total Qty"},
  ];
  const aggForDl=groupBy(df,"Plant Name",[["Total Value","Total Value"],["Total Qty","Total Qty"]]);
  document.getElementById("btn-dl-dash-xlsx").onclick=()=>downloadExcel(aggForDl,dlCols,"dashboard_summary.xlsx");
  document.getElementById("btn-dl-dash-csv").onclick=()=>downloadCSV(aggForDl,dlCols,"dashboard_summary.csv");
}

// ═══════════════════════════════════════════════════════════════════════════
// TRANSIT
// ═══════════════════════════════════════════════════════════════════════════
function renderTransit() {
  const pf = pageFilters.transit;
  document.getElementById("transit-filter-plant").value = pf.plant||"";
  document.getElementById("transit-filter-mg").value    = pf.mg||"";
  const df = applyPageFilter("transit").filter(r=>r["Stock in Transit"]>0);

  const totalTV  = df.reduce((s,r)=>s+r["Value of Stock in Transit"],0);
  const totalTQ  = df.reduce((s,r)=>s+r["Stock in Transit"],0);
  const uniqMat  = new Set(df.map(r=>r["Material"])).size;
  setKpis("transit-kpis",[
    ["Total Transit Value",         fmtETB(totalTV), "Across all plants","amber"],
    ["Total Transit Quantity",       fmtQty(totalTQ), "Units in movement","blue"],
    ["Unique Materials in Transit",  String(uniqMat), "Distinct SKUs","green"],
  ]);

  if (!df.length) { document.getElementById("transit-table-wrap").innerHTML=`<div class="alert-info">ℹ️ No items are currently in transit.</div>`; return; }

  // Dual chart: qty + value by plant
  const plantAgg=sortBy(groupBy(df,"Plant Name",[["val","Value of Stock in Transit"],["qty","Stock in Transit"]]),"val");
  Plotly.newPlot("chart-transit-plant",[
    {type:"bar",name:"Value (ETB)",x:plantAgg.map(r=>r["Plant Name"]),y:plantAgg.map(r=>r.val),yaxis:"y",marker:{color:"#d29922"},hovertemplate:"<b>%{x}</b><br>ETB %{y:,.0f}<extra></extra>"},
    {type:"scatter",mode:"lines+markers",name:"Qty",x:plantAgg.map(r=>r["Plant Name"]),y:plantAgg.map(r=>r.qty),yaxis:"y2",marker:{color:"#3fb950",size:8},line:{color:"#3fb950"},hovertemplate:"<b>%{x}</b><br>Qty: %{y:,.0f}<extra></extra>"},
  ],pl({height:280,margin:{l:20,r:60,t:20,b:80},yaxis2:{overlaying:"y",side:"right",gridcolor:"transparent",tickfont:{color:"#3fb950"}}}),PLOTLY_CONFIG);

  // Table
  const transitCols=[
    {key:"Material",label:"Material"},
    {key:"Material Description",label:"Description"},
    {key:"Material Group Name",label:"Material Group"},
    {key:"Plant Name",label:"Plant"},
    {key:"Stock in Transit",label:"Transit Qty",fmt:fmtQty,rawKey:"Stock in Transit",cellClass:"col-qty"},
    {key:"Value of Stock in Transit",label:"Transit Value (ETB)",fmt:fmtETB,rawKey:"Value of Stock in Transit",cellClass:"col-val"},
    {key:"_status",label:"Status", raw:true},
  ];
  const transitRows=sortBy([...df],"Value of Stock in Transit").map(r=>({
    ...r,
    _status: r["Value of Stock in Transit"]>100000?"<span class='badge badge-red'>Critical</span>":r["Value of Stock in Transit"]>50000?"<span class='badge badge-amber'>High</span>":r["Value of Stock in Transit"]>10000?"<span class='badge badge-amber'>Medium</span>":"<span class='badge badge-green'>Low</span>",
  }));
  document.getElementById("transit-table-wrap").innerHTML=buildTable(transitRows,transitCols);
  document.getElementById("btn-dl-transit").onclick=()=>downloadCSV(transitRows,transitCols.slice(0,-1),"transit_analysis.csv");
  document.getElementById("btn-dl-transit-xlsx").onclick=()=>downloadExcel(transitRows,transitCols.slice(0,-1),"transit_analysis.xlsx");

  // FIX 6: use filtered df (not rawDf) so HO01 respects active plant/MG filters
  const allTransitDf = applyPageFilter("transit");
  const ho01=allTransitDf.filter(r=>r["Stock in Transit"]>0).filter(r=>String(r["Plant"]).toUpperCase()==="HO01"||String(r["Plant Name"]).toUpperCase().includes("HO01")||String(r["Plant Name"]).toUpperCase().includes("HEAD OFFICE")||String(r["Plant Name"]).toUpperCase().includes("CENTRAL"));
  if (ho01.length) {
    setKpis("ho01-kpis",[
      ["HO01 Transit Value",   fmtETB(ho01.reduce((s,r)=>s+r["Value of Stock in Transit"],0)),"From central hub","amber"],
      ["HO01 Transit Qty",     fmtQty(ho01.reduce((s,r)=>s+r["Stock in Transit"],0)),"Units in movement","blue"],
      ["Unique SKUs",          String(new Set(ho01.map(r=>r["Material"])).size),"Distinct materials","green"],
    ]);
    document.getElementById("ho01-table-wrap").innerHTML=buildTable(sortBy([...ho01],"Value of Stock in Transit"),transitCols.slice(0,-1));
    document.getElementById("btn-dl-ho01").onclick=()=>downloadCSV(ho01,transitCols.slice(0,-1),"ho01_transit.csv");
  } else {
    document.getElementById("ho01-kpis").innerHTML=`<div class="alert-info">No HO01 transit records found.</div>`;
    document.getElementById("ho01-table-wrap").innerHTML="";
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// EXPIRY
// ═══════════════════════════════════════════════════════════════════════════
function renderExpiry() {
  const pf = pageFilters.expiry;
  document.getElementById("expiry-filter-plant").value = pf.plant||"";
  document.getElementById("expiry-filter-mg").value    = pf.mg||"";
  const baseDf = applyPageFilter("expiry");
  const months = parseInt(document.querySelector('input[name="expWin"]:checked')?.value||6);
  const today  = new Date();
  const cutoff = new Date(today); cutoff.setMonth(cutoff.getMonth()+months);
  const valid   = baseDf.filter(r=>r._expiry instanceof Date && !isNaN(r._expiry));
  const expiring= valid.filter(r=>r._expiry>=today&&r._expiry<=cutoff);
  const expired = valid.filter(r=>r._expiry<today);

  setKpis("expiry-kpis",[
    ["Expiring in Window", String(expiring.length), `Items within next ${months} months`,"amber"],
    ["Already Expired",    String(expired.length),  "Requires immediate action","red"],
    ["At-Risk Value",      fmtETB(expiring.reduce((s,r)=>s+r["Value of Unrestricted Stock"],0)),"Unrestricted stock value","purple"],
    ["At-Risk Quantity",   fmtQty(expiring.reduce((s,r)=>s+r["Unrestricted Stock"],0)),"Units expiring soon","amber"],
  ]);

  if (expiring.length) {
    const monthMap={}, valMap={};
    expiring.forEach(r=>{
      const key=`${r._expiry.getFullYear()}-${String(r._expiry.getMonth()+1).padStart(2,"0")}`;
      monthMap[key]=(monthMap[key]||0)+1;
      valMap[key]=(valMap[key]||0)+r["Value of Unrestricted Stock"];
    });
    const ms=Object.keys(monthMap).sort();
    Plotly.newPlot("chart-expiry-timeline",[
      {type:"bar",name:"Items Count",x:ms,y:ms.map(m=>monthMap[m]),marker:{color:"#d29922"},hovertemplate:"<b>%{x}</b><br>%{y} items<extra></extra>"},
      {type:"scatter",mode:"lines+markers",name:"Value at Risk",x:ms,y:ms.map(m=>valMap[m]),yaxis:"y2",marker:{color:"#f85149",size:8},line:{color:"#f85149"},hovertemplate:"<b>%{x}</b><br>ETB %{y:,.0f}<extra></extra>"},
    ],pl({height:260,margin:{l:20,r:60,t:20,b:60},yaxis2:{overlaying:"y",side:"right",gridcolor:"transparent",tickfont:{color:"#f85149"}}}),PLOTLY_CONFIG);
  } else { document.getElementById("chart-expiry-timeline").innerHTML=""; }

  const expCols=[
    {key:"Material",label:"Material"},
    {key:"Material Description",label:"Description"},
    {key:"Material Group Name",label:"Material Group"},
    {key:"Plant Name",label:"Plant"},
    {key:"Description of Storage Location",label:"Storage Location"},
    {key:"_expiryStr",label:"Expiry Date"},
    {key:"Unrestricted Stock",label:"Qty",fmt:fmtQty,rawKey:"Unrestricted Stock",cellClass:"col-qty"},
    {key:"Value of Unrestricted Stock",label:"Value (ETB)",fmt:fmtETB,rawKey:"Value of Unrestricted Stock",cellClass:"col-val"},
    {key:"_daysLeft",label:"Days Until Expiry"},
  ];
  const expRows=sortBy(expiring.map(r=>({...r,_expiryStr:r._expiry?r._expiry.toISOString().slice(0,10):"",_daysLeft:r._expiry?Math.floor((r._expiry-today)/86400000):9999})),"_daysLeft",true);
  document.getElementById("expiry-table-wrap").innerHTML=expRows.length?buildTable(expRows,expCols,r=>r._daysLeft<=30?"row-red":r._daysLeft<=90?"row-amber":""):`<div class="alert-info">✓ No items expiring within the selected window.</div>`;
  document.getElementById("btn-dl-expiry").onclick=()=>downloadCSV(expRows,expCols,`expiry_${months}months.csv`);
  document.getElementById("btn-dl-expiry-xlsx").onclick=()=>downloadExcel(expRows,expCols,`expiry_${months}months.xlsx`);

  if (expired.length) {
    document.getElementById("expired-section").style.display="block";
    document.getElementById("expired-header").textContent=`🔴 Already Expired Items (${expired.length})`;
    const expiredRows=expired.map(r=>({...r,_expiryStr:r._expiry?r._expiry.toISOString().slice(0,10):""}));
    document.getElementById("expired-table-wrap").innerHTML=buildTable(expiredRows,[
      {key:"Material",label:"Material"},{key:"Material Description",label:"Description"},
      {key:"Material Group Name",label:"Material Group"},{key:"Plant Name",label:"Plant"},
      {key:"Description of Storage Location",label:"Storage Location"},
      {key:"_expiryStr",label:"Expiry Date"},
      {key:"Unrestricted Stock",label:"Qty",fmt:fmtQty,rawKey:"Unrestricted Stock",cellClass:"col-qty"},
    ]);
    document.getElementById("btn-dl-expired").onclick=()=>downloadCSV(expiredRows,[{key:"Material",label:"Material"},{key:"Material Description",label:"Description"},{key:"Plant Name",label:"Plant"},{key:"Description of Storage Location",label:"Storage Location"},{key:"_expiryStr",label:"Expiry Date"},{key:"Unrestricted Stock",label:"Qty",rawKey:"Unrestricted Stock"}],"expired_items.csv");
  } else { document.getElementById("expired-section").style.display="none"; }
}

// ═══════════════════════════════════════════════════════════════════════════
// QC
// ═══════════════════════════════════════════════════════════════════════════
function renderQC() {
  const pf = pageFilters.qc;
  document.getElementById("qc-filter-plant").value = pf.plant||"";
  document.getElementById("qc-filter-mg").value    = pf.mg||"";
  const df = applyPageFilter("qc").filter(r=>r["Stock in Quality Inspection"]>0);

  const totalQCVal=df.reduce((s,r)=>s+r["Value of Stock in Quality Inspection"],0);
  const totalQCQty=df.reduce((s,r)=>s+r["Stock in Quality Inspection"],0);
  setKpis("qc-kpis",[
    ["Total Value in QC", fmtETB(totalQCVal),"Across all plants","red"],
    ["Total QC Quantity",  fmtQty(totalQCQty),"Units under inspection","amber"],
    ["Unique Materials",   String(new Set(df.map(r=>r["Material"])).size),"Distinct SKUs","blue"],
  ]);

  if (!df.length) { document.getElementById("qc-table-wrap").innerHTML=`<div class="alert-info">✓ No items in quality inspection.</div>`; return; }

  const plantQC=sortBy(groupBy(df,"Plant Name",[["val","Value of Stock in Quality Inspection"],["qty","Stock in Quality Inspection"]]),"val");
  Plotly.newPlot("chart-qc-plant",[
    {type:"bar",name:"Value (ETB)",x:plantQC.map(r=>r["Plant Name"]),y:plantQC.map(r=>r.val),yaxis:"y",marker:{color:"#f85149"},hovertemplate:"<b>%{x}</b><br>ETB %{y:,.0f}<extra></extra>"},
    {type:"scatter",mode:"lines+markers",name:"Qty",x:plantQC.map(r=>r["Plant Name"]),y:plantQC.map(r=>r.qty),yaxis:"y2",marker:{color:"#3fb950",size:8},line:{color:"#3fb950"},hovertemplate:"<b>%{x}</b><br>Qty: %{y:,.0f}<extra></extra>"},
  ],pl({height:280,margin:{l:20,r:60,t:20,b:80},yaxis2:{overlaying:"y",side:"right",gridcolor:"transparent",tickfont:{color:"#3fb950"}}}),PLOTLY_CONFIG);

  const qcCols=[
    {key:"Material",label:"Material"},{key:"Material Description",label:"Description"},
    {key:"Material Group Name",label:"Material Group"},{key:"Plant Name",label:"Plant"},
    {key:"Description of Storage Location",label:"Storage Location"},
    {key:"_expiryStr",label:"Shelf Life Expiry"},
    {key:"Stock in Quality Inspection",label:"QC Qty",fmt:fmtQty,rawKey:"Stock in Quality Inspection",cellClass:"col-qty"},
    {key:"Value of Stock in Quality Inspection",label:"QC Value (ETB)",fmt:fmtETB,rawKey:"Value of Stock in Quality Inspection",cellClass:"col-val"},
  ];
  const qcRows=sortBy([...df].map(r=>({...r,_expiryStr:r._expiry?r._expiry.toISOString().slice(0,10):""})),"Value of Stock in Quality Inspection");
  document.getElementById("qc-table-wrap").innerHTML=buildTable(qcRows,qcCols,r=>r["Value of Stock in Quality Inspection"]>10000?"row-red":"");
  document.getElementById("btn-dl-qc").onclick=()=>downloadCSV(qcRows,qcCols,"qc_inspection.csv");
  document.getElementById("btn-dl-qc-xlsx").onclick=()=>downloadExcel(qcRows,qcCols,"qc_inspection.xlsx");
}

// ═══════════════════════════════════════════════════════════════════════════
// BRANCH COMPARISON
// ═══════════════════════════════════════════════════════════════════════════
function renderBranch() {
  const pf = pageFilters.branch;
  document.getElementById("branch-filter-mg").value = pf.mg||"";
  // FIX 5: apply reconciliation consistently — was reading rawDf directly
  const baseDf = applyReconciliationToData(rawDf);
  const df = pf.mg ? baseDf.filter(r=>r["Material Group Name"]===pf.mg) : baseDf;

  const plants=[...new Set(df.map(r=>String(r["Plant"]).toUpperCase()))];
  let centralCode,centralName;
  if (plants.includes("HO01")) {
    centralCode="HO01";
    centralName=df.find(r=>String(r["Plant"]).toUpperCase()==="HO01")?.["Plant Name"]||"HO01";
    document.getElementById("branch-central-info").style.display="none";
  } else {
    const totals={};
    df.forEach(r=>{const p=r["Plant Name"];totals[p]=(totals[p]||0)+r["Total Value"];});
    centralName=Object.entries(totals).sort((a,b)=>b[1]-a[1])[0]?.[0]||"";
    document.getElementById("branch-central-info").style.display="block";
    document.getElementById("branch-central-info").innerHTML=`ℹ️ HO01 not found — using <b>${centralName}</b> as central branch (highest inventory value).`;
  }

  const aggMap={};
  df.forEach(r=>{
    const k=r["Plant Name"];
    if (!aggMap[k]) aggMap[k]={PlantName:k,Plant:r["Plant"],TotalValue:0,Unrestricted:0,Transit:0,QC:0,UnrestrictedQty:0,TransitQty:0,QCQty:0,Items:0};
    aggMap[k].TotalValue  +=r["Total Value"];
    aggMap[k].Unrestricted+=r["Value of Unrestricted Stock"];
    aggMap[k].Transit     +=r["Value of Stock in Transit"];
    aggMap[k].QC          +=r["Value of Stock in Quality Inspection"];
    aggMap[k].UnrestrictedQty+=r["Unrestricted Stock"];
    aggMap[k].TransitQty  +=r["Stock in Transit"];
    aggMap[k].QCQty       +=r["Stock in Quality Inspection"];
    aggMap[k].Items++;
  });
  const branchAgg=Object.values(aggMap);
  const others=branchAgg.map(r=>r.PlantName).filter(b=>b!==centralName);

  const matPlantMap={};
  df.forEach(r=>{
    const mat=r["Material"],pln=r["Plant Name"];
    if (!matPlantMap[mat]) matPlantMap[mat]={desc:r["Material Description"],group:r["Material Group Name"]};
    if (!matPlantMap[mat][pln]) matPlantMap[mat][pln]={Unrestricted:0,Transit:0,QC:0,TotalValue:0,TotalQty:0,UnrestrictedQty:0};
    matPlantMap[mat][pln].Unrestricted+=r["Value of Unrestricted Stock"];
    matPlantMap[mat][pln].Transit     +=r["Value of Stock in Transit"];
    matPlantMap[mat][pln].QC          +=r["Value of Stock in Quality Inspection"];
    matPlantMap[mat][pln].TotalValue  +=r["Total Value"];
    matPlantMap[mat][pln].TotalQty    +=r["Total Qty"];
    matPlantMap[mat][pln].UnrestrictedQty+=r["Unrestricted Stock"];
  });

  const tabsHtml=`
    <div class="branch-tabs" id="branch-tabs">
      <button class="branch-tab active" data-tab="value">📊 Total Value Comparison</button>
      <button class="branch-tab" data-tab="material">🔬 Line-Item (Material Across Branches)</button>
    </div>
    <div id="branch-tab-value"></div>
    <div id="branch-tab-material" style="display:none"></div>`;
  document.getElementById("branch-tabs-wrap").innerHTML=tabsHtml;

  document.querySelectorAll(".branch-tab").forEach(btn=>{
    btn.addEventListener("click",()=>{
      document.querySelectorAll(".branch-tab").forEach(b=>b.classList.remove("active"));
      btn.classList.add("active");
      const tab=btn.dataset.tab;
      document.getElementById("branch-tab-value").style.display   =tab==="value"   ?"block":"none";
      document.getElementById("branch-tab-material").style.display=tab==="material"?"block":"none";
      if (tab==="material") renderMaterialTab();
    });
  });

  const sel=document.getElementById("branch-select");
  sel.innerHTML="";
  others.forEach(b=>{
    const opt=document.createElement("option");
    opt.value=b; opt.textContent=b; opt.selected=true;
    sel.appendChild(opt);
  });

  // ── TAB 1: Total Value ──
  function updateBranchCharts() {
    const selected=[...sel.selectedOptions].map(o=>o.value);
    const wrap=document.getElementById("branch-tab-value");
    if (!selected.length) { wrap.innerHTML=`<div class="alert-warning">⚠️ Select at least one branch.</div>`; return; }
    const compareNames=[centralName,...selected];
    const compareDf=branchAgg.filter(r=>compareNames.includes(r.PlantName));

    const bCols=[
      {key:"PlantName",label:"Plant Name"},
      {key:"TotalValue",label:"Total Value (ETB)",fmt:fmtETB,rawKey:"TotalValue"},
      {key:"Unrestricted",label:"Unrestricted (ETB)",fmt:fmtETB,rawKey:"Unrestricted"},
      {key:"UnrestrictedQty",label:"Avail Qty",fmt:fmtQty,rawKey:"UnrestrictedQty",cellClass:"col-qty"},
      {key:"Transit",label:"Transit (ETB)",fmt:fmtETB,rawKey:"Transit"},
      {key:"TransitQty",label:"Transit Qty",fmt:fmtQty,rawKey:"TransitQty",cellClass:"col-qty"},
      {key:"QC",label:"QC (ETB)",fmt:fmtETB,rawKey:"QC"},
      {key:"QCQty",label:"QC Qty",fmt:fmtQty,rawKey:"QCQty",cellClass:"col-qty"},
      {key:"Items",label:"# Line Items"},
    ];
    wrap.innerHTML=`
      <div id="branch-table-wrap-inner" style="margin-bottom:1rem">${buildTable(compareDf,bCols,r=>r.PlantName===centralName?"row-blue":"")}</div>
      <div class="chart-box full" style="margin-top:1rem"><div class="section-header">Value Comparison by Stock Type</div><div id="chart-branch-grouped"></div></div>
      <div class="chart-box full"><div class="section-header">Available Quantity vs Transit Quantity</div><div id="chart-branch-qty"></div></div>`;

    const types=["Unrestricted","Transit","QC"];
    const colors2={Unrestricted:"#58a6ff",Transit:"#d29922",QC:"#f85149"};
    Plotly.newPlot("chart-branch-grouped",types.map(t=>({
      type:"bar",name:t,x:compareDf.map(r=>r.PlantName),y:compareDf.map(r=>r[t]),
      marker:{color:colors2[t]},hovertemplate:`<b>%{x}</b> · ${t}<br>ETB %{y:,.0f}<extra></extra>`,
    })),pl({barmode:"group",height:300}),PLOTLY_CONFIG);

    const qtyTypes={UnrestrictedQty:"#3fb950",TransitQty:"#d29922",QCQty:"#f85149"};
    Plotly.newPlot("chart-branch-qty",Object.entries(qtyTypes).map(([k,col])=>({
      type:"bar",name:k.replace("Qty",""),x:compareDf.map(r=>r.PlantName),y:compareDf.map(r=>r[k]),
      marker:{color:col},hovertemplate:`<b>%{x}</b><br>Qty: %{y:,.0f}<extra></extra>`,
    })),pl({barmode:"group",height:280}),PLOTLY_CONFIG);

    document.getElementById("btn-dl-branch-csv").onclick=()=>downloadCSV(compareDf,bCols,"branch_comparison.csv");
    document.getElementById("btn-dl-branch-xlsx").onclick=()=>downloadExcel(compareDf,bCols,"branch_comparison.xlsx");
  }

  // ── TAB 2: Material Across Branches ──
  // FIX 9: initialize as false each time renderBranch runs so filter changes
  // cause the Material tab UI (including its own MG dropdown) to rebuild correctly.
  let matTabInitialized=false;
  function renderMaterialTab() {
    const wrap=document.getElementById("branch-tab-material");
    const allPlantNames=[...new Set(df.map(r=>r["Plant Name"]))].sort((a,b)=>{
      if (a===centralName) return -1; if (b===centralName) return 1; return a.localeCompare(b);
    });

    if (!matTabInitialized) {
      matTabInitialized=true;
      const mgNamesForFilter=[...new Set(df.map(r=>r["Material Group Name"]))].filter(Boolean).filter(name=>!isNonMedicalGroup(name)).sort();
      wrap.innerHTML=`
        <div style="display:flex;gap:0.8rem;flex-wrap:wrap;align-items:flex-end;margin-bottom:1rem;background:var(--surface);border:1px solid var(--border);border-radius:8px;padding:0.8rem">
          <div>
            <div class="nav-label" style="font-size:0.65rem;margin-bottom:3px">Search Material</div>
            <input id="mat-search" type="text" placeholder="code or description…" style="background:var(--surface2);border:1px solid var(--border2);color:var(--text);padding:6px 10px;border-radius:6px;width:220px;font-size:13px">
          </div>
          <div>
            <div class="nav-label" style="font-size:0.65rem;margin-bottom:3px">Metric</div>
            <select id="mat-metric" style="background:var(--surface2);border:1px solid var(--border2);color:var(--text);padding:6px 10px;border-radius:6px;font-size:13px">
              <option value="TotalValue">Total Value (ETB)</option>
              <option value="Unrestricted">Unrestricted Value (ETB)</option>
              <option value="Transit">Transit Value (ETB)</option>
              <option value="QC">QC Value (ETB)</option>
              <option value="TotalQty">Total Quantity</option>
              <option value="UnrestrictedQty">Available Quantity</option>
            </select>
          </div>
          <div>
            <div class="nav-label" style="font-size:0.65rem;margin-bottom:3px">Material Group</div>
            <select id="mat-mgfilter" style="background:var(--surface2);border:1px solid var(--border2);color:var(--text);padding:6px 10px;border-radius:6px;font-size:13px">
              <option value="">All Material Groups</option>
              ${mgNamesForFilter.map(m=>`<option value="${m}">${m}</option>`).join("")}
            </select>
          </div>
          <div>
            <div class="nav-label" style="font-size:0.65rem;margin-bottom:3px">Sort By</div>
            <select id="mat-sort" style="background:var(--surface2);border:1px solid var(--border2);color:var(--text);padding:6px 10px;border-radius:6px;font-size:13px">
              <option value="total_desc">Highest Total ↓</option>
              <option value="total_asc">Lowest Total ↑</option>
              <option value="desc_asc">Description A–Z</option>
              <option value="spread_desc">Most Branches ↓</option>
            </select>
          </div>
          <button id="mat-apply" class="apply-btn">Apply</button>
          <button id="mat-dl-csv" class="dl-btn">⬇ CSV</button>
          <button id="mat-dl-xlsx" class="dl-btn">⬇ Excel</button>
        </div>
        <div id="mat-chart-wrap" style="margin-bottom:1rem"></div>
        <div id="mat-table-wrap"></div>`;
      document.getElementById("mat-apply").addEventListener("click",refreshMaterialView);
      document.getElementById("mat-search").addEventListener("keydown",e=>{if(e.key==="Enter")refreshMaterialView();});
    }
    refreshMaterialView();

    function refreshMaterialView() {
      const searchVal=(document.getElementById("mat-search").value||"").toLowerCase().trim();
      const metric=document.getElementById("mat-metric").value;
      const sortMode=document.getElementById("mat-sort").value;
      const mgFilter=document.getElementById("mat-mgfilter").value;
      const isQty=metric.includes("Qty");
      const fmtFn=isQty?fmtQty:fmtETB;

      let materials=Object.entries(matPlantMap)
        .filter(([mat,info])=>{
          if (mgFilter&&info.group!==mgFilter) return false;
          if (searchVal) return mat.toLowerCase().includes(searchVal)||info.desc.toLowerCase().includes(searchVal);
          return true;
        })
        .map(([mat,info])=>{
          const plantData={};
          let grandTotal=0,branchCount=0;
          allPlantNames.forEach(pn=>{
            const v=info[pn]?info[pn][metric]:0;
            plantData[pn]=v||0;
            grandTotal+=plantData[pn];
            if ((info[pn]?.TotalValue||0)>0) branchCount++;
          });
          return {mat,desc:info.desc,group:info.group,plantData,grandTotal,branchCount};
        });

      if (sortMode==="total_desc") materials.sort((a,b)=>b.grandTotal-a.grandTotal);
      if (sortMode==="total_asc")  materials.sort((a,b)=>a.grandTotal-b.grandTotal);
      if (sortMode==="desc_asc")   materials.sort((a,b)=>a.desc.localeCompare(b.desc));
      if (sortMode==="spread_desc")materials.sort((a,b)=>b.branchCount-a.branchCount);

      const top=materials.slice(0,30);
      const chartWrap=document.getElementById("mat-chart-wrap");
      if (!top.length) { chartWrap.innerHTML=`<div class="alert-info">No materials found.</div>`; document.getElementById("mat-table-wrap").innerHTML=""; return; }

      if (top.length===1) {
        const info=top[0];
        chartWrap.innerHTML=`<div class="chart-box full"><div class="section-header">${info.desc} (${info.mat}) — ${metric} across branches</div><div id="chart-mat-detail"></div></div>`;
        Plotly.newPlot("chart-mat-detail",[{type:"bar",x:allPlantNames,y:allPlantNames.map(pn=>info.plantData[pn]||0),marker:{color:allPlantNames.map((_,i)=>COLORWAY[i%COLORWAY.length])},hovertemplate:`<b>%{x}</b><br>${metric}: %{y:,.0f}<extra></extra>`}],pl({height:320}),PLOTLY_CONFIG);
      } else {
        const plantTraces=allPlantNames.map((pn,i)=>({
          type:"bar",name:pn,x:top.map(m=>m.mat),y:top.map(m=>m.plantData[pn]||0),
          customdata:top.map(m=>m.desc),marker:{color:COLORWAY[i%COLORWAY.length]},
          hovertemplate:`<b>%{customdata}</b><br>${pn}<br>${metric}: %{y:,.0f}<extra></extra>`,
        }));
        chartWrap.innerHTML=`<div class="chart-box full"><div class="section-header">Top ${top.length} Materials — ${metric} by Branch</div><div id="chart-mat-multi"></div></div>`;
        Plotly.newPlot("chart-mat-multi",plantTraces,pl({barmode:"group",height:Math.max(360,20*top.length),xaxis:{...PLOTLY_LAYOUT.xaxis,tickangle:-40}}),PLOTLY_CONFIG);
      }

      const colDefs=[
        {key:"mat",label:"Material"},
        {key:"desc",label:"Description"},
        {key:"group",label:"Material Group"},
        ...allPlantNames.map(pn=>({key:`__p__${pn}`,label:pn,fmt:fmtFn,rawKey:`__r__${pn}`,cellClass:isQty?"col-qty":"col-val"})),
        {key:"grandTotal",label:"Grand Total",fmt:fmtFn,rawKey:"grandTotal",cellClass:isQty?"col-qty":"col-val"},
        {key:"branchCount",label:"# Branches"},
      ];
      const tableRows=materials.slice(0,200).map(m=>{
        const row={mat:m.mat,desc:m.desc,group:m.group,grandTotal:m.grandTotal,branchCount:m.branchCount};
        allPlantNames.forEach(pn=>{row[`__p__${pn}`]=m.plantData[pn]||0; row[`__r__${pn}`]=m.plantData[pn]||0;});
        row["__r__grandTotal"]=m.grandTotal;
        return row;
      });

      const centralKey=`__p__${centralName}`;
      const thead=`<thead><tr>${colDefs.map(c=>`<th${c.key===centralKey?' style="color:#58a6ff;background:#0d2035"':""} >${c.label}</th>`).join("")}</tr></thead>`;
      const tbody=tableRows.map(r=>{
        const cells=colDefs.map(c=>{
          const v=r[c.key];
          const display=c.fmt?c.fmt(v):(v==null?"":v);
          const isZero=typeof v==="number"&&v===0;
          const style=c.key===centralKey?'style="color:#58a6ff;background:#0d2035"':isZero?'style="color:#484f58"':"";
          const cls=c.cellClass||"";
          return `<td class="${cls}" ${style}>${display}</td>`;
        }).join("");
        return `<tr>${cells}</tr>`;
      }).join("");
      document.getElementById("mat-table-wrap").innerHTML=`
        <div style="color:var(--muted);font-size:12px;margin-bottom:6px">Showing ${tableRows.length} of ${materials.length} materials · Blue = Central (${centralName})</div>
        <div class="tbl-wrap"><table>${thead}<tbody>${tbody}</tbody></table></div>
        ${materials.length>200?`<div class="alert-info">Showing first 200 of ${materials.length}. Refine search.</div>`:""}`;

      // DL handlers
      const flatCols=[{key:"mat",label:"Material"},{key:"desc",label:"Description"},{key:"group",label:"Material Group"},...allPlantNames.map(pn=>({key:`__p__${pn}`,label:pn,rawKey:`__r__${pn}`})),{key:"grandTotal",label:"Grand Total"}];
      const btn=document.getElementById("mat-dl-csv");
      if (btn) btn.onclick=()=>downloadCSV(tableRows,flatCols,"materials_by_branch.csv");
      const btnX=document.getElementById("mat-dl-xlsx");
      if (btnX) btnX.onclick=()=>downloadExcel(tableRows,flatCols,"materials_by_branch.xlsx");
    }
  }

  sel.addEventListener("change",updateBranchCharts);
  updateBranchCharts();
}

// ═══════════════════════════════════════════════════════════════════════════
// INVENTORY FLOW (rebuilt)
// ═══════════════════════════════════════════════════════════════════════════
function renderFlow() {
  const pf = pageFilters.flow;
  document.getElementById("flow-filter-plant").value = pf.plant||"";
  document.getElementById("flow-filter-mg").value    = pf.mg||"";
  const df = applyPageFilter("flow");

  const totalVal   = df.reduce((s,r)=>s+r["Total Value"],0);
  const transitVal = df.reduce((s,r)=>s+r["Value of Stock in Transit"],0);
  const qcVal      = df.reduce((s,r)=>s+r["Value of Stock in Quality Inspection"],0);
  const availVal   = df.reduce((s,r)=>s+r["Value of Unrestricted Stock"],0);
  const totalQty   = df.reduce((s,r)=>s+r["Total Qty"],0);
  const availQty   = df.reduce((s,r)=>s+r["Unrestricted Stock"],0);

  // Reorder alerts: unrestricted=0 but transit or QC > 0
  const reorderItems = df.filter(r=>r["Unrestricted Stock"]===0&&(r["Stock in Transit"]>0||r["Stock in Quality Inspection"]>0));

  setKpis("flow-kpis",[
    ["Total Inventory",     fmtETB(totalVal),  `${fmtQty(totalQty)} units`,"blue"],
    ["Available Stock",     fmtETB(availVal),  `${fmtQty(availQty)} units unrestricted`,"green"],
    ["In Transit (Inbound)",fmtETB(transitVal),`${fmtQty(df.reduce((s,r)=>s+r["Stock in Transit"],0))} units`,"amber"],
    ["In QC",               fmtETB(qcVal),     `${fmtQty(df.reduce((s,r)=>s+r["Stock in Quality Inspection"],0))} units`,"red"],
    ["Reorder Alerts",      String(reorderItems.length),"Zero unrestricted stock","red"],
  ]);

  // Reorder table
  const reorderCols=[
    {key:"Material",label:"Material"},{key:"Material Description",label:"Description"},
    {key:"Material Group Name",label:"Material Group"},{key:"Plant Name",label:"Plant"},
    {key:"Unrestricted Stock",label:"Avail Qty",fmt:fmtQty,rawKey:"Unrestricted Stock",cellClass:"col-qty"},
    {key:"Stock in Transit",label:"In Transit",fmt:fmtQty,rawKey:"Stock in Transit",cellClass:"col-qty"},
    {key:"Stock in Quality Inspection",label:"In QC",fmt:fmtQty,rawKey:"Stock in Quality Inspection",cellClass:"col-qty"},
    {key:"Value of Stock in Transit",label:"Transit Value (ETB)",fmt:fmtETB,rawKey:"Value of Stock in Transit",cellClass:"col-val"},
    {key:"_alert",label:"Alert", raw:true},
  ];
  const reorderRows = reorderItems.map(r=>({...r,
    _alert: r["Stock in Transit"]>0&&r["Stock in Quality Inspection"]>0
      ? "<span class='badge badge-red'>Transit+QC</span>"
      : r["Stock in Transit"]>0
      ? "<span class='badge badge-amber'>Awaiting Transit</span>"
      : "<span class='badge badge-amber'>Awaiting QC Release</span>",
  }));
  document.getElementById("reorder-table-wrap").innerHTML = reorderRows.length
    ? buildTable(reorderRows,reorderCols,()=>"row-amber")
    : `<div class="alert-info">✓ No reorder alerts — all materials have available unrestricted stock.</div>`;

  // Stock levels by plant — stacked bar qty
  const plantAgg=sortBy(groupBy(df,"Plant Name",[["avail","Unrestricted Stock"],["transit","Stock in Transit"],["qc","Stock in Quality Inspection"],["availVal","Value of Unrestricted Stock"],["transitVal","Value of Stock in Transit"]]),"avail");
  Plotly.newPlot("chart-stock-levels",[
    {type:"bar",name:"Available (Qty)",x:plantAgg.map(r=>r["Plant Name"]),y:plantAgg.map(r=>r.avail),marker:{color:"#3fb950"},hovertemplate:"<b>%{x}</b><br>Available: %{y:,.0f}<extra></extra>"},
    {type:"bar",name:"In Transit (Qty)",x:plantAgg.map(r=>r["Plant Name"]),y:plantAgg.map(r=>r.transit),marker:{color:"#d29922"},hovertemplate:"<b>%{x}</b><br>Transit: %{y:,.0f}<extra></extra>"},
    {type:"bar",name:"In QC (Qty)",x:plantAgg.map(r=>r["Plant Name"]),y:plantAgg.map(r=>r.qc),marker:{color:"#f85149"},hovertemplate:"<b>%{x}</b><br>QC: %{y:,.0f}<extra></extra>"},
  ],pl({barmode:"stack",height:300,margin:{l:20,r:20,t:20,b:80}}),PLOTLY_CONFIG);

  // Transfers table (transit items with destination)
  const transferData = df.filter(r=>r["Stock in Transit"]>0);
  const transferCols=[
    {key:"Material",label:"Material"},{key:"Material Description",label:"Description"},
    {key:"Material Group Name",label:"Material Group"},{key:"Plant Name",label:"Destination Plant"},
    {key:"Stock in Transit",label:"Transit Qty",fmt:fmtQty,rawKey:"Stock in Transit",cellClass:"col-qty"},
    {key:"Value of Stock in Transit",label:"Transit Value (ETB)",fmt:fmtETB,rawKey:"Value of Stock in Transit",cellClass:"col-val"},
  ];
  document.getElementById("transfer-table-wrap").innerHTML = transferData.length
    ? buildTable(sortBy([...transferData],"Value of Stock in Transit"),transferCols)
    : `<div class="alert-info">No active transfers found.</div>`;

  // Inbound vs available
  Plotly.newPlot("chart-inbound-outbound",[
    {type:"bar",name:"Available Value (ETB)",x:plantAgg.map(r=>r["Plant Name"]),y:plantAgg.map(r=>r.availVal),marker:{color:"#3fb950"},hovertemplate:"<b>%{x}</b><br>Available: ETB %{y:,.0f}<extra></extra>"},
    {type:"bar",name:"Inbound Transit (ETB)",x:plantAgg.map(r=>r["Plant Name"]),y:plantAgg.map(r=>r.transitVal),marker:{color:"#d29922"},hovertemplate:"<b>%{x}</b><br>Inbound: ETB %{y:,.0f}<extra></extra>"},
  ],pl({barmode:"group",height:300,margin:{l:20,r:20,t:20,b:80}}),PLOTLY_CONFIG);

  const flowForDl = df.map(r=>({...r}));
  const flowDlCols=[
    {key:"Material",label:"Material"},{key:"Material Description",label:"Description"},
    {key:"Plant Name",label:"Plant"},{key:"Material Group Name",label:"Material Group"},
    {key:"Unrestricted Stock",label:"Available Qty",rawKey:"Unrestricted Stock"},
    {key:"Stock in Transit",label:"Transit Qty",rawKey:"Stock in Transit"},
    {key:"Stock in Quality Inspection",label:"QC Qty",rawKey:"Stock in Quality Inspection"},
    {key:"Value of Unrestricted Stock",label:"Available Value (ETB)",rawKey:"Value of Unrestricted Stock"},
    {key:"Value of Stock in Transit",label:"Transit Value (ETB)",rawKey:"Value of Stock in Transit"},
  ];
  document.getElementById("btn-dl-flow-csv").onclick=()=>downloadCSV(flowForDl,flowDlCols,"inventory_flow.csv");
  document.getElementById("btn-dl-flow-xlsx").onclick=()=>downloadExcel(flowForDl,flowDlCols,"inventory_flow.xlsx");
}

// ═══════════════════════════════════════════════════════════════════════════
// DATA PREVIEW
// ═══════════════════════════════════════════════════════════════════════════
function renderPreview() {
  filtDf = rawDf;
  populatePreviewFilters();
  renderPreviewTable();
}

function populatePreviewFilters() {
  function fill(id, key, excludeFn) {
    const sel = document.getElementById(id); if (!sel) return;
    const vals = [...new Set(rawDf.map(r=>r[key]))]
      .filter(Boolean)
      .filter(v => !excludeFn || !excludeFn(v))
      .sort();
    sel.innerHTML = vals.map(v=>`<option value="${v}">${v}</option>`).join("");
  }
  fill("filter-plant",  "Plant Name",         null);
  fill("filter-mg",     "Material Group Name", isNonMedicalGroup);
  fill("filter-mgname", "Material Group Name", isNonMedicalGroup);
}

function applyPreviewFilters() {
  const getSelected=id=>[...document.querySelectorAll(`#${id} option:checked`)].map(o=>o.value);
  const plants=getSelected("filter-plant");
  const mgs=getSelected("filter-mg");
  const mgnames=getSelected("filter-mgname");
  filtDf=rawDf.filter(r=>
    (!plants.length||plants.includes(r["Plant Name"]))&&
    (!mgs.length||mgs.includes(r["Material Group Name"]))&&
    (!mgnames.length||mgnames.includes(r["Material Group Name"]))
  );
  renderPreviewTable();
}

function renderPreviewTable() {
  const df=filtDf;
  setKpis("preview-kpis",[
    ["Total Records",     df.length.toLocaleString(),"After filtering","blue"],
    ["Unique Materials",  new Set(df.map(r=>r["Material"])).size.toLocaleString(),"Distinct SKUs","green"],
    ["Total Plants",      new Set(df.map(r=>r["Plant"])).size.toLocaleString(),"Stocking locations","amber"],
    ["Material Groups",   new Set(df.map(r=>r["Material Group Name"])).size.toLocaleString(),"Therapeutic categories","purple"],
  ]);
  document.getElementById("preview-count").innerHTML=`Showing <b>${df.length.toLocaleString()}</b> of <b>${rawDf.length.toLocaleString()}</b> records`;

  const cols=[
    {key:"Material",label:"Material"},
    {key:"Material Description",label:"Description"},
    {key:"Plant Name",label:"Plant"},
    {key:"Material Group Name",label:"Material Group"},
    {key:"Unrestricted Stock",label:"Avail Qty",fmt:fmtQty,rawKey:"Unrestricted Stock",cellClass:"col-qty"},
    {key:"Stock in Transit",label:"Transit Qty",fmt:fmtQty,rawKey:"Stock in Transit",cellClass:"col-qty"},
    {key:"Stock in Quality Inspection",label:"QC Qty",fmt:fmtQty,rawKey:"Stock in Quality Inspection",cellClass:"col-qty"},
    {key:"Value of Unrestricted Stock",label:"Avail Value (ETB)",fmt:fmtETB,rawKey:"Value of Unrestricted Stock",cellClass:"col-val"},
    {key:"Value of Stock in Transit",label:"Transit Value (ETB)",fmt:fmtETB,rawKey:"Value of Stock in Transit",cellClass:"col-val"},
    {key:"Value of Stock in Quality Inspection",label:"QC Value (ETB)",fmt:fmtETB,rawKey:"Value of Stock in Quality Inspection",cellClass:"col-val"},
    {key:"Total Value",label:"Total Value (ETB)",fmt:fmtETB,rawKey:"Total Value",cellClass:"col-val"},
    {key:"_expiryStr",label:"Expiry Date"},
  ];
  const rows=df.slice(0,500).map(r=>({...r,_expiryStr:r._expiry?r._expiry.toISOString().slice(0,10):""}));
  document.getElementById("preview-table-wrap").innerHTML=
    buildTable(rows,cols) +
    (df.length>500 ? `<div class="alert-warning">⚠️ Showing first 500 of ${df.length.toLocaleString()} records. Apply filters to narrow results or use the Excel/CSV download for the full dataset.</div>` : "");
  document.getElementById("btn-dl-preview").onclick=()=>downloadCSV(rows,cols,"pharma_inventory_filtered.csv");
  document.getElementById("btn-dl-preview-xlsx").onclick=()=>downloadExcel(rows,cols,"pharma_inventory_filtered.xlsx");
}

// ═══════════════════════════════════════════════════════════════════════════
// MATERIAL CODE RECONCILIATION
// ═══════════════════════════════════════════════════════════════════════════

// Returns the canonical (primary) code for a given material code,
// based on defined reconcile groups. Used by all renderers.
function getCanonicalCode(code) {
  for (const g of reconcileGroups) {
    if (g.codes.includes(code)) return g.codes[0]; // first code = primary
  }
  return code;
}

// Merges rawDf rows so that reconciled codes are aggregated together.
// Returns a new array where aliased materials are summed into their primary code.
function applyReconciliationToData(df) {
  if (!reconcileGroups.length) return df;
  const merged = [];
  const primaryMap = {}; // primaryCode -> merged row index

  // Numeric columns that get summed across reconciled rows
  const NUM_COLS = ["Unrestricted Stock","Stock in Quality Inspection","Blocked Stock","Stock in Transit",
    "Value of Stock in Quality Inspection","Value of Stock in Transit",
    "Value of Unrestricted Stock","Total Value","Total Qty"];

  df.forEach(row => {
    const primary = getCanonicalCode(row["Material"]);
    if (primary === row["Material"]) {
      const idx = primaryMap[primary];
      if (idx !== undefined) {
        // Primary seen before — aggregate numeric cols only
        const target = merged[idx];
        NUM_COLS.forEach(c => { target[c] = (target[c]||0) + (row[c]||0); });
        // FIX 3: carry forward non-numeric fields if the primary row is missing them
        if (!target["_expiry"] && row["_expiry"]) target["_expiry"] = row["_expiry"];
        if (!target["Description of Storage Location"] && row["Description of Storage Location"])
          target["Description of Storage Location"] = row["Description of Storage Location"];
        if (!target["Batch"] && row["Batch"]) target["Batch"] = row["Batch"];
      } else {
        primaryMap[primary] = merged.length;
        merged.push({...row});
      }
    } else {
      // This is an alias — merge into primary
      const idx = primaryMap[primary];
      if (idx !== undefined) {
        const target = merged[idx];
        NUM_COLS.forEach(c => { target[c] = (target[c]||0) + (row[c]||0); });
        // FIX 3: carry forward non-numeric fields from alias if primary is missing them
        if (!target["_expiry"] && row["_expiry"]) target["_expiry"] = row["_expiry"];
        if (!target["Description of Storage Location"] && row["Description of Storage Location"])
          target["Description of Storage Location"] = row["Description of Storage Location"];
        if (!target["Batch"] && row["Batch"]) target["Batch"] = row["Batch"];
        // Append alias code to description for visibility
        if (!target["Material Description"].includes(`[+${row["Material"]}]`)) {
          target["Material Description"] += ` [+${row["Material"]}]`;
        }
      } else {
        // Primary not yet seen — add this row as if it were the primary
        primaryMap[primary] = merged.length;
        merged.push({...row, Material: primary});
      }
    }
  });
  return merged;
}

function openReconcilePanel() {
  document.getElementById("reconcile-panel").style.display="flex";
  document.getElementById("reconcile-overlay").style.display="block";
  refreshReconcileGroupsList();
  // Restore search state
  const q = document.getElementById("rp-search").value;
  if (q.trim()) searchReconcileMaterials(q);
  else refreshReconcilePending();
}

function closeReconcilePanel() {
  document.getElementById("reconcile-panel").style.display="none";
  document.getElementById("reconcile-overlay").style.display="none";
}

function searchReconcileMaterials(query) {
  const resultsEl = document.getElementById("rp-search-results");
  if (!rawDf.length || !query.trim()) {
    resultsEl.innerHTML=`<div style="color:var(--dim);font-size:0.78rem;padding:0.5rem">Upload data and type to search materials.</div>`;
    return;
  }
  const q = query.toLowerCase().trim();
  const allMaterials = [...new Map(rawDf.map(r=>[r["Material"],{code:r["Material"],desc:r["Material Description"]}])).values()];
  const matches = allMaterials.filter(m=>m.code.toLowerCase().includes(q)||m.desc.toLowerCase().includes(q)).slice(0,20);

  if (!matches.length) { resultsEl.innerHTML=`<div class="alert-info">No materials found.</div>`; return; }

  // FIX Bug 1: Use index-based closure instead of data-attributes for click handler
  // This avoids HTML attribute encoding issues with special chars in descriptions
  resultsEl.innerHTML = matches.map((m, i) => {
    const isPending = reconcilePending.some(p=>p.code===m.code);
    const inGroup   = reconcileGroups.some(g=>g.codes.includes(m.code));
    const tag = inGroup ? `<span style="font-size:0.65rem;color:var(--green);margin-left:4px">✓ In group</span>` : "";
    return `<div class="rp-result-item${isPending?" selected":""}" data-match-idx="${i}">
      <span class="rp-result-code">${escHtml(m.code)}</span>
      <span class="rp-result-desc">${escHtml(m.desc)}${tag}</span>
      <button class="rp-add-btn${isPending?" added":""}" data-match-idx="${i}">${isPending?"✓ Added":"+ Add"}</button>
    </div>`;
  }).join("");

  // Attach listeners using closure over `matches` array — no data-attribute encoding needed
  resultsEl.querySelectorAll(".rp-add-btn").forEach(btn => {
    btn.addEventListener("click", e => {
      e.stopPropagation();
      const m = matches[parseInt(btn.dataset.matchIdx)];
      if (!m) return;
      if (!reconcilePending.some(p=>p.code===m.code)) {
        reconcilePending.push({code: m.code, desc: m.desc});
        refreshReconcilePending();
        searchReconcileMaterials(query); // re-render results to update "Added" state
      }
    });
  });
}


function refreshReconcilePending() {
  const el    = document.getElementById("rp-pending");
  const count = document.getElementById("rp-pending-count");
  count.textContent = reconcilePending.length ? `(${reconcilePending.length})` : "";
  if (!reconcilePending.length) {
    el.innerHTML=`<div style="color:var(--dim);font-size:0.78rem;padding:0.5rem">No codes selected yet. Search and add codes above.</div>`;
    return;
  }
  el.innerHTML = reconcilePending.map((p, i) => `
    <div class="rp-pending-item">
      <span class="rp-result-code">${escHtml(p.code)}</span>
      <span class="rp-result-desc" style="flex:1;margin:0 0.5rem;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:0.74rem;color:var(--muted)">${escHtml(p.desc)}</span>
      <button class="rp-pending-remove" data-idx="${i}">✕</button>
    </div>`).join("");

  // FIX Bug 2 (partial): Use stable index from current reconcilePending snapshot
  el.querySelectorAll(".rp-pending-remove").forEach(btn => {
    const idx = parseInt(btn.dataset.idx);
    btn.addEventListener("click", () => {
      reconcilePending.splice(idx, 1);
      refreshReconcilePending();
      const q = document.getElementById("rp-search").value;
      if (q.trim()) searchReconcileMaterials(q);
    });
  });
}

function confirmReconcileGroup() {
  if (reconcilePending.length < 2) { alert("Select at least 2 material codes to link."); return; }

  // Check if any pending code already belongs to an existing group
  const conflicts = reconcilePending.filter(p => reconcileGroups.some(g=>g.codes.includes(p.code)));
  if (conflicts.length) {
    alert(`These codes are already in a group: ${conflicts.map(c=>c.code).join(", ")}\nRemove them from their existing group first.`);
    return;
  }

  // Prompt for a meaningful group name
  const defaultName = `Group ${reconcileGroups.length+1} (${reconcilePending[0].code} +${reconcilePending.length-1} more)`;
  const name = (prompt("Enter a name for this reconciliation group:", defaultName) || "").trim() || defaultName;

  reconcileGroups.push({ name, codes: reconcilePending.map(p=>p.code) });
  reconcilePending.length = 0;
  refreshReconcilePending();
  refreshReconcileGroupsList();
  document.getElementById("rp-search").value = "";
  document.getElementById("rp-search-results").innerHTML = "";

  // Save to localStorage so groups persist across sessions
  saveReconcileGroups();

  // Re-render current page so reconciliation takes effect immediately
  if (rawDf.length) renderPage(currentPage);
}

function refreshReconcileGroupsList() {
  const el = document.getElementById("rp-groups-list");
  if (!reconcileGroups.length) {
    el.innerHTML=`<div style="color:var(--dim);font-size:0.78rem;padding:0.5rem">No link groups created yet.</div>`;
    return;
  }

  // FIX Bug 2: Use event delegation on the container instead of per-button listeners
  // This avoids stale index issues after deletes
  el.innerHTML = reconcileGroups.map((g, i) => `
    <div class="rp-group-card">
      <div class="rp-group-header">
        <span class="rp-group-name">🔗 ${escHtml(g.name)}</span>
        <div style="display:flex;gap:0.4rem;align-items:center">
          <span style="font-size:0.65rem;color:var(--dim)">${g.codes.length} codes · primary: <span style="color:var(--purple)">${escHtml(g.codes[0])}</span></span>
          <button class="rp-group-del" data-group-idx="${i}">Delete</button>
        </div>
      </div>
      <div class="rp-group-codes">${g.codes.map((c,ci)=>`<span class="rp-code-tag" title="${ci===0?'Primary code':'Alias'}">${escHtml(c)}${ci===0?' ★':''}</span>`).join("")}</div>
    </div>`).join("");

  // Single delegated listener on the container — no stale index issues
  el.onclick = e => {
    const btn = e.target.closest(".rp-group-del");
    if (!btn) return;
    const idx = parseInt(btn.dataset.groupIdx);
    if (!isNaN(idx)) {
      reconcileGroups.splice(idx, 1);
      saveReconcileGroups();
      refreshReconcileGroupsList();
      if (rawDf.length) renderPage(currentPage);
    }
  };
}

// ── Persistence ──────────────────────────────────────────────────────────
const RECONCILE_STORE_KEY = "pharmatrack_reconcile_v2";
function saveReconcileGroups() {
  try { localStorage.setItem(RECONCILE_STORE_KEY, JSON.stringify(reconcileGroups)); } catch(e) {}
}
function loadReconcileGroups() {
  try {
    const saved = localStorage.getItem(RECONCILE_STORE_KEY);
    if (saved) reconcileGroups = JSON.parse(saved);
  } catch(e) { reconcileGroups = []; }
}

// ═══════════════════════════════════════════════════════════════════════════
// PAGE SWITCHING
// ═══════════════════════════════════════════════════════════════════════════
const PAGE_RENDERERS={ dashboard:renderDashboard, transit:renderTransit, expiry:renderExpiry, qc:renderQC, branch:renderBranch, flow:renderFlow, preview:renderPreview };

function renderPage(id) {
  if (!rawDf.length) return;
  currentPage=id;
  document.querySelectorAll(".page").forEach(el=>{el.style.display="none";});
  const pg=document.getElementById(`page-${id}`);
  if (pg) pg.style.display="block";
  document.querySelectorAll(".nav-btn").forEach(btn=>btn.classList.toggle("active",btn.dataset.page===id));
  try { PAGE_RENDERERS[id]?.(); } catch(e) { console.error(`Error rendering ${id}:`,e); }
}

// ═══════════════════════════════════════════════════════════════════════════
// EVENT LISTENERS
// ═══════════════════════════════════════════════════════════════════════════
document.addEventListener("DOMContentLoaded",()=>{
  // Load persisted reconcile groups before anything else
  loadReconcileGroups();

  // Nav
  document.querySelectorAll(".nav-btn[data-page]").forEach(btn=>{
    btn.addEventListener("click",()=>renderPage(btn.dataset.page));
  });

  // File upload
  document.getElementById("fileInput").addEventListener("change",e=>{
    const f=e.target.files[0]; if (f) loadFile(f);
  });

  // Expiry window
  document.getElementById("expiry-window-group").addEventListener("change",()=>{
    if (rawDf.length&&currentPage==="expiry") renderExpiry();
  });

  // Preview filters
  document.getElementById("btn-apply-filter").addEventListener("click",applyPreviewFilters);
  document.getElementById("btn-clear-filter").addEventListener("click",()=>{
    document.querySelectorAll("#filter-plant option,#filter-mg option,#filter-mgname option").forEach(o=>{o.selected=false;});
    filtDf=rawDf; renderPreviewTable();
  });

  // ── Page filter wiring ──
  function wirePageFilters(page, plantId, mgId, applyId, clearId) {
    const applyBtn=document.getElementById(applyId);
    const clearBtn=document.getElementById(clearId);
    if (applyBtn) applyBtn.addEventListener("click",()=>{
      if (plantId) pageFilters[page].plant=document.getElementById(plantId)?.value||"";
      if (mgId)    pageFilters[page].mg   =document.getElementById(mgId)?.value||"";
      renderPage(page);
    });
    if (clearBtn) clearBtn.addEventListener("click",()=>{
      if (plantId) { pageFilters[page].plant=""; const el=document.getElementById(plantId); if(el) el.value=""; }
      if (mgId)    { pageFilters[page].mg="";    const el=document.getElementById(mgId);    if(el) el.value=""; }
      renderPage(page);
    });
  }

  wirePageFilters("dashboard","dash-filter-plant","dash-filter-mg","dash-filter-apply","dash-filter-clear");
  wirePageFilters("transit","transit-filter-plant","transit-filter-mg","transit-filter-apply","transit-filter-clear");
  wirePageFilters("expiry","expiry-filter-plant","expiry-filter-mg","expiry-filter-apply","expiry-filter-clear");
  wirePageFilters("qc","qc-filter-plant","qc-filter-mg","qc-filter-apply","qc-filter-clear");
  wirePageFilters("branch",null,"branch-filter-mg","branch-filter-apply","branch-filter-clear");
  wirePageFilters("flow","flow-filter-plant","flow-filter-mg","flow-filter-apply","flow-filter-clear");

  // ── Reconciliation panel ──
  document.getElementById("open-reconcile-btn").addEventListener("click",openReconcilePanel);
  document.getElementById("reconcile-close").addEventListener("click",closeReconcilePanel);
  document.getElementById("reconcile-overlay").addEventListener("click",closeReconcilePanel);
  document.getElementById("rp-search").addEventListener("input",e=>searchReconcileMaterials(e.target.value));
  document.getElementById("rp-confirm").addEventListener("click",confirmReconcileGroup);
  document.getElementById("rp-clear-pending").addEventListener("click",()=>{ reconcilePending.length=0; refreshReconcilePending(); });
});
