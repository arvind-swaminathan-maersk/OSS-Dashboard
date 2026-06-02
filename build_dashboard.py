#!/usr/bin/env python3
"""
build_dashboard.py (Multi-File Version)
=========================================
Scans the data/ folder for ALL .xlsx files, consolidates them into a
single dataset with deduplication, and generates the interactive
OSS Incident Coaching Heatmap & Feedback Guide dashboard.

Usage (local):
    python build_dashboard.py

Usage (GitHub Actions):
    Triggered automatically when any file in data/ is updated.

Output:
    output/index.html
"""

import pandas as pd
import json
import os
import sys
import glob
from datetime import datetime


# --- Configuration --------------------------------------------------------
DATA_DIR    = "data"
OUTPUT_DIR  = "output"
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "index.html")
DEDUP_KEY   = "Case No."          # column used to deduplicate

FLAG_COLUMNS = [
    "Flag_HowToOrConsulting",
    "Flag_PriorityInflation",
    "Flag_CustomerDelay",
    "Flag_ComponentMismatch",
    "Flag_Reopened",
    "Flag_AutoConfirmed",
    "Flag_TotalCoaching",
]

FLAG_LABELS = {
    "Flag_HowToOrConsulting": "How-To / Consulting",
    "Flag_PriorityInflation": "Priority Inflation",
    "Flag_CustomerDelay": "Customer Delay",
    "Flag_ComponentMismatch": "Component Mismatch",
    "Flag_Reopened": "Reopened",
    "Flag_AutoConfirmed": "Auto-Confirmed",
    "Flag_TotalCoaching": "Total Coaching",
}

FLAG_GUIDANCE = {
    "Flag_HowToOrConsulting": {
        "icon": "\u2753",
        "means": "Case was a how-to or consulting question, not a software defect.",
        "do": "Search SAP KBAs & documentation before raising; use SAP Community forums.",
        "avoid": "Raising OSS cases for configuration guidance or how-to questions.",
    },
    "Flag_PriorityInflation": {
        "icon": "\u26a0\ufe0f",
        "means": "Priority set higher than SAP guidelines warrant for the issue type.",
        "do": "Follow SAP priority matrix; set P1 only for production-down scenarios.",
        "avoid": "Inflating priority to get faster response \u2014 it triggers audits.",
    },
    "Flag_CustomerDelay": {
        "icon": "\u23f3",
        "means": "Significant delays caused by late customer responses.",
        "do": "Respond to SAP within 24h; assign a backup if primary contact is away.",
        "avoid": "Leaving cases in \u2018Customer Action\u2019 for days without responding.",
    },
    "Flag_ComponentMismatch": {
        "icon": "\ud83d\udee0\ufe0f",
        "means": "Case was logged under the wrong SAP component.",
        "do": "Verify the correct component using SAP\u2019s component selector before logging.",
        "avoid": "Guessing the component \u2014 misrouting causes 3\u20135 day delays.",
    },
    "Flag_Reopened": {
        "icon": "\ud83d\udd04",
        "means": "Case was reopened after being closed, indicating incomplete resolution.",
        "do": "Test the fix thoroughly before confirming; document remaining issues clearly.",
        "avoid": "Confirming closure without validating the solution in your environment.",
    },
    "Flag_AutoConfirmed": {
        "icon": "\u2705",
        "means": "Case was auto-confirmed by SAP due to no customer response within SLA.",
        "do": "Set reminders to respond before auto-confirm deadline (typically 8 days).",
        "avoid": "Ignoring SAP\u2019s \u2018Proposed Solution\u2019 notifications.",
    },
    "Flag_TotalCoaching": {
        "icon": "\ud83c\udfaf",
        "means": "One or more coaching opportunities identified in this case.",
        "do": "Review flagged cases as learning opportunities; share insights with your team.",
        "avoid": "Dismissing coaching flags \u2014 they highlight patterns that waste effort.",
    },
}


# --- Data Loading (Multi-File) -------------------------------------------

def load_all_files(data_dir):
    """Scan data/ for all .xlsx files, load and concatenate them."""
    pattern = os.path.join(data_dir, "*.xlsx")
    files = sorted(glob.glob(pattern))

    if not files:
        print(f"\n[ERROR] No .xlsx files found in '{data_dir}/' folder.")
        print("  Place at least one Excel file in the data/ directory.")
        sys.exit(1)

    print(f"\n[1/5] Scanning {data_dir}/ ... found {len(files)} file(s):")
    frames = []
    for fpath in files:
        fname = os.path.basename(fpath)
        try:
            tmp = pd.read_excel(fpath, engine="openpyxl")
            rows = len(tmp)
            tmp["_source_file"] = fname  # track origin
            frames.append(tmp)
            print(f"       {fname:50s}  {rows:>6,} rows  \u2705")
        except Exception as e:
            print(f"       {fname:50s}  SKIPPED \u26a0\ufe0f  ({e})")

    if not frames:
        print("\n[ERROR] No files could be loaded successfully.")
        sys.exit(1)

    df = pd.concat(frames, ignore_index=True)
    print(f"\n       Total rows loaded: {len(df):,}")
    return df


def deduplicate(df):
    """Remove duplicate cases, keeping the most recent entry."""
    if DEDUP_KEY not in df.columns:
        print(f"  [WARN] Column '{DEDUP_KEY}' not found - skipping deduplication.")
        return df

    before = len(df)

    # Sort by Closing Date desc so the latest record is kept
    if "Closing Date" in df.columns:
        df["_sort_date"] = pd.to_datetime(df["Closing Date"], errors="coerce")
        df = df.sort_values("_sort_date", ascending=False, na_position="last")
        df = df.drop(columns=["_sort_date"])

    df = df.drop_duplicates(subset=[DEDUP_KEY], keep="first")
    after = len(df)
    removed = before - after

    print(f"[2/5] Deduplication on '{DEDUP_KEY}':")
    print(f"       Before: {before:,} rows")
    print(f"       After:  {after:,} rows")
    if removed > 0:
        print(f"       Removed: {removed:,} duplicate(s)")
    else:
        print(f"       No duplicates found \u2705")

    return df.reset_index(drop=True)


# --- Data Cleaning --------------------------------------------------------

def clean(df):
    """Clean flag, numeric, string, and date columns."""
    for col in FLAG_COLUMNS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    for col in ["Days_AtCustomer", "Days_AtSAP"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    for col in ["Workstream", "Priority", "Error Category"]:
        if col in df.columns:
            df[col] = df[col].fillna("Unknown").astype(str).str.strip()

    for col in ["Creation Date", "Closing Date"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    return df


def compute_date_range(df):
    d_min = df["Creation Date"].min() if "Creation Date" in df.columns else None
    d_max = df["Closing Date"].max() if "Closing Date" in df.columns else None
    if pd.notna(d_min) and pd.notna(d_max):
        return f"{d_min.strftime('%d %b %Y')} \u2013 {d_max.strftime('%d %b %Y')}"
    return "Date range unavailable"


# --- Aggregation ----------------------------------------------------------

def aggregate(df):
    flag_cols_no_total = [c for c in FLAG_COLUMNS if c != "Flag_TotalCoaching"]
    rows = []

    groups = list(df.groupby("Workstream")) + [("All Workstreams", df)]

    for ws, grp in groups:
        n = len(grp)
        if n == 0:
            continue

        rec = {"workstream": ws, "total_cases": n}

        for fc in FLAG_COLUMNS:
            cnt = int(grp[fc].sum()) if fc in grp.columns else 0
            rec[fc + "_count"] = cnt
            rec[fc + "_pct"] = round(cnt / n * 100, 1)

        if "Flag_TotalCoaching" in grp.columns:
            flagged = int((grp["Flag_TotalCoaching"] > 0).sum())
        else:
            flagged = 0
        rec["cases_with_flags"] = flagged
        rec["cases_with_flags_pct"] = round(flagged / n * 100, 1)

        rec["avg_days_customer"] = round(grp["Days_AtCustomer"].mean(), 1) if "Days_AtCustomer" in grp.columns else 0
        rec["avg_days_sap"] = round(grp["Days_AtSAP"].mean(), 1) if "Days_AtSAP" in grp.columns else 0

        if "Priority" in grp.columns:
            pri = grp["Priority"].value_counts().to_dict()
        else:
            pri = {}
        rec["priority_dist"] = {str(k): int(v) for k, v in pri.items()}

        if "Error Category" in grp.columns:
            ec = grp["Error Category"].value_counts().head(10).to_dict()
        else:
            ec = {}
        rec["error_categories"] = {str(k): int(v) for k, v in ec.items()}

        flag_counts = {fc: rec[fc + "_count"] for fc in flag_cols_no_total}
        sorted_flags = sorted(flag_counts.items(), key=lambda x: x[1], reverse=True)
        if sorted_flags and sorted_flags[0][1] > 0:
            rec["top_flag"] = FLAG_LABELS.get(sorted_flags[0][0], sorted_flags[0][0])
            rec["top_flag_key"] = sorted_flags[0][0]
        else:
            rec["top_flag"] = "None"
            rec["top_flag_key"] = ""
        if len(sorted_flags) > 1 and sorted_flags[1][1] > 0:
            rec["second_flag"] = FLAG_LABELS.get(sorted_flags[1][0], sorted_flags[1][0])
            rec["second_flag_key"] = sorted_flags[1][0]
        else:
            rec["second_flag"] = ""
            rec["second_flag_key"] = ""

        rows.append(rec)

    all_ws = [r for r in rows if r["workstream"] == "All Workstreams"]
    others = sorted([r for r in rows if r["workstream"] != "All Workstreams"],
                    key=lambda x: x["total_cases"], reverse=True)
    return all_ws + others


# --- HTML Generation ------------------------------------------------------

def build_html(data, date_range, total_cases, file_count, dedup_info):
    data_json = json.dumps(data, indent=2)
    flag_guidance_json = json.dumps(FLAG_GUIDANCE, indent=2)
    flag_labels_json = json.dumps(FLAG_LABELS, indent=2)
    flag_columns_json = json.dumps([c for c in FLAG_COLUMNS if c != "Flag_TotalCoaching"])
    now_str = datetime.now().strftime("%d %b %Y %H:%M")

    html = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>OSS Incident Coaching Heatmap &amp; Feedback Guide</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;background:#f0f2f5;color:#333}
.disclaimer{background:#fff3cd;border-bottom:2px solid #ffc107;padding:10px 24px;font-size:13px;color:#856404;text-align:center}
.disclaimer a{color:#856404;font-weight:600}
header{background:linear-gradient(135deg,#1a1a2e 0%,#16213e 50%,#0f3460 100%);color:#fff;padding:28px 24px 20px;text-align:center}
header h1{font-size:26px;margin-bottom:6px;letter-spacing:0.5px}
header p{font-size:14px;opacity:.85}
.data-badge{display:inline-block;background:rgba(255,255,255,0.15);border-radius:20px;padding:4px 14px;font-size:12px;margin-top:8px}
.how-to-link{color:#ffd700;cursor:pointer;text-decoration:underline;font-size:13px;margin-top:6px;display:inline-block}
.controls{display:flex;justify-content:center;gap:16px;padding:16px 24px;background:#fff;border-bottom:1px solid #ddd;flex-wrap:wrap;align-items:center}
.controls label{font-weight:600;font-size:14px}
.controls select{padding:8px 14px;border-radius:6px;border:1px solid #ccc;font-size:14px;min-width:220px}
.container{max-width:1300px;margin:0 auto;padding:20px 24px}
.kpi-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:14px;margin-bottom:22px}
.kpi-card{background:#fff;border-radius:10px;padding:18px 16px;text-align:center;box-shadow:0 2px 8px rgba(0,0,0,.07);border-top:4px solid #0f3460}
.kpi-card .kpi-value{font-size:28px;font-weight:700;color:#0f3460;margin:6px 0}
.kpi-card .kpi-label{font-size:12px;color:#666;text-transform:uppercase;letter-spacing:.5px}
.section{background:#fff;border-radius:10px;padding:20px;margin-bottom:20px;box-shadow:0 2px 8px rgba(0,0,0,.07)}
.section h2{font-size:18px;color:#1a1a2e;margin-bottom:14px;padding-bottom:8px;border-bottom:2px solid #e9ecef}
table{width:100%;border-collapse:collapse;font-size:13px}
th,td{padding:9px 10px;text-align:center;border:1px solid #e0e0e0}
th{background:#1a1a2e;color:#fff;font-weight:600;font-size:12px;position:sticky;top:0}
td.ws-name{text-align:left;font-weight:600;background:#f8f9fa;white-space:nowrap}
.heat-cell{cursor:pointer;font-weight:600;font-size:12px;transition:transform .15s}
.heat-cell:hover{transform:scale(1.08);outline:2px solid #0f3460}
.tooltip-box{display:none;position:fixed;background:#1a1a2e;color:#fff;padding:14px 18px;border-radius:8px;font-size:13px;max-width:340px;z-index:1000;box-shadow:0 8px 24px rgba(0,0,0,.25);line-height:1.5}
.tooltip-box .tip-title{font-weight:700;font-size:14px;margin-bottom:6px;color:#ffd700}
.tooltip-box .tip-do{color:#82ffb5}
.tooltip-box .tip-avoid{color:#ff8a8a}
.coaching-table .flag-icon{font-size:18px}
.coaching-table td{font-size:12px;text-align:left;padding:8px 10px}
.coaching-table td:nth-child(6),.coaching-table td:nth-child(7){text-align:center}
.donut-container{display:flex;justify-content:center;align-items:center;gap:30px;flex-wrap:wrap;padding:10px 0}
.legend-item{display:flex;align-items:center;gap:8px;font-size:13px;margin-bottom:6px}
.legend-color{width:14px;height:14px;border-radius:3px}
.summary-block{background:#f8f9fa;border-left:4px solid #0f3460;padding:14px 18px;border-radius:6px;margin-top:10px;font-size:13px;line-height:1.7}
footer{text-align:center;padding:18px;font-size:12px;color:#999;border-top:1px solid #e0e0e0;margin-top:10px}
.modal-overlay{display:none;position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,.6);z-index:2000;justify-content:center;align-items:center}
.modal-box{background:#fff;border-radius:12px;padding:28px 32px;max-width:600px;width:90%;max-height:80vh;overflow-y:auto;box-shadow:0 12px 40px rgba(0,0,0,.3)}
.modal-box h2{font-size:20px;color:#1a1a2e;margin-bottom:14px}
.modal-box ol{padding-left:20px;font-size:14px;line-height:2}
.modal-box .close-btn{float:right;font-size:22px;cursor:pointer;color:#999;background:none;border:none}
.modal-box .close-btn:hover{color:#333}
</style>
</head>
<body>

<div class="disclaimer">
\u26a0\ufe0f <strong>AI-Generated Coaching Dashboard</strong> \u2014 This dashboard was generated using AI analysis of SAP OSS closed-case data. Flags and recommendations are indicative, not definitive. Always validate with case details before taking action. Questions? Contact <a href="mailto:arvind.swaminathan@maersk.com">arvind.swaminathan@maersk.com</a>
</div>

<header>
<h1>\ud83d\udcca OSS Incident Coaching Heatmap &amp; Feedback Guide</h1>
<p id="subtitle">""" + str(total_cases) + """ Closed Cases | """ + date_range + """</p>
<div class="data-badge">\ud83d\udcc1 """ + str(file_count) + """ data file(s) consolidated | """ + dedup_info + """</div>
<br>
<span class="how-to-link" onclick="document.getElementById('howToModal').style.display='flex'">\u2753 How to use this dashboard</span>
</header>

<div class="modal-overlay" id="howToModal">
<div class="modal-box">
<button class="close-btn" onclick="this.parentElement.parentElement.style.display='none'">&times;</button>
<h2>\ud83d\udcd6 How to Use This Dashboard</h2>
<ol>
<li><strong>Select a Workstream</strong> from the dropdown to filter all panels.</li>
<li><strong>Review KPI Cards</strong> for a quick summary of case volume, coaching flags, and response times.</li>
<li><strong>Examine the Heatmap</strong> \u2014 darker red cells indicate higher issue concentration. <strong>Click any cell</strong> for specific coaching advice.</li>
<li><strong>Read the Coaching Flags table</strong> to understand what each flag means and recommended actions.</li>
<li><strong>Check Error Categories</strong> to see the most common error types.</li>
<li><strong>View Priority Distribution</strong> in the donut chart to assess priority mix.</li>
<li><strong>Use the Coaching Summary</strong> for quick, actionable takeaways per workstream.</li>
</ol>
</div>
</div>

<div class="controls">
<label for="wsFilter">Workstream:</label>
<select id="wsFilter" onchange="applyFilter()">
</select>
</div>

<div class="container">
<div class="kpi-row" id="kpiRow"></div>
<div class="section" id="heatmapSection"><h2>\ud83d\udd25 Coaching Heatmap</h2><div style="overflow-x:auto"><table id="heatmapTable"></table></div></div>
<div class="section" id="coachingSection"><h2>\ud83c\udfaf Coaching Flags \u2014 What They Mean</h2><div style="overflow-x:auto"><table class="coaching-table" id="coachingTable"></table></div></div>
<div class="section" id="errorSection"><h2>\ud83d\udcc2 Error Category Breakdown</h2><div style="overflow-x:auto"><table id="errorTable"></table></div></div>
<div class="section" id="prioritySection"><h2>\ud83c\udfaf Priority Distribution</h2><div class="donut-container" id="donutContainer"></div></div>
<div class="section" id="summarySection"><h2>\ud83d\udcdd Actionable Coaching Summary</h2><div class="summary-block" id="summaryBlock"></div></div>
</div>

<footer>Generated on """ + now_str + """ | OSS Incident Coaching Heatmap &amp; Feedback Guide | Data-driven coaching for SAP support excellence</footer>

<div class="tooltip-box" id="tooltip"></div>

<script>
const DATA = """ + data_json + """;
const FLAG_GUIDANCE = """ + flag_guidance_json + """;
const FLAG_LABELS = """ + flag_labels_json + """;
const FLAG_COLS = """ + flag_columns_json + """;
const PRIORITY_COLORS = {"1 - Very High":"#dc3545","2 - High":"#fd7e14","3 - Medium":"#ffc107","4 - Low":"#28a745","Unknown":"#6c757d"};

function initDropdown(){
  const sel=document.getElementById('wsFilter');
  DATA.forEach(d=>{const o=document.createElement('option');o.value=d.workstream;o.textContent=d.workstream;sel.appendChild(o);});
}

function getRow(ws){return DATA.find(d=>d.workstream===ws)||DATA[0];}

function applyFilter(){
  const ws=document.getElementById('wsFilter').value;
  const r=getRow(ws);
  renderKPIs(r);
  renderHeatmap(ws);
  renderCoachingTable(r);
  renderErrorTable(r);
  renderDonut(r);
  renderSummary(r);
}

function renderKPIs(r){
  const h=document.getElementById('kpiRow');
  h.innerHTML=`
    <div class="kpi-card"><div class="kpi-label">Total Cases</div><div class="kpi-value">${r.total_cases}</div></div>
    <div class="kpi-card"><div class="kpi-label">Coaching Flags</div><div class="kpi-value">${r.Flag_TotalCoaching_count}</div></div>
    <div class="kpi-card"><div class="kpi-label">Cases Flagged</div><div class="kpi-value">${r.cases_with_flags} (${r.cases_with_flags_pct}%)</div></div>
    <div class="kpi-card"><div class="kpi-label">Avg Days @ Customer</div><div class="kpi-value">${r.avg_days_customer}</div></div>
    <div class="kpi-card"><div class="kpi-label">Avg Days @ SAP</div><div class="kpi-value">${r.avg_days_sap}</div></div>
    <div class="kpi-card"><div class="kpi-label">Top Coaching Issue</div><div class="kpi-value" style="font-size:16px">${r.top_flag}</div></div>`;
}

function heatColor(pct){
  if(pct===0) return '#f8f9fa';
  const i=Math.min(pct/50,1);
  const r_=255, g=Math.round(255*(1-i*.85)), b=Math.round(255*(1-i*.85));
  return `rgb(${r_},${g},${b})`;
}

function renderHeatmap(ws){
  const t=document.getElementById('heatmapTable');
  let rows = ws==='All Workstreams' ? DATA.filter(d=>d.workstream!=='All Workstreams') : [getRow(ws)];
  let hdr='<thead><tr><th>Workstream</th><th>Cases</th>';
  FLAG_COLS.forEach(f=>{hdr+=`<th>${FLAG_LABELS[f]}</th>`;});
  hdr+='<th>Total Coaching</th></tr></thead>';
  let body='<tbody>';
  rows.forEach(r=>{
    body+=`<tr><td class="ws-name">${r.workstream}</td><td>${r.total_cases}</td>`;
    FLAG_COLS.forEach(f=>{
      const cnt=r[f+'_count'], pct=r[f+'_pct'];
      body+=`<td class="heat-cell" style="background:${heatColor(pct)}" onclick="showTip(event,'${f}',${cnt},${pct},'${r.workstream.replace(/'/g,"\\'")}')" title="Click for coaching advice">${cnt}<br><small>${pct}%</small></td>`;
    });
    body+=`<td style="background:${heatColor(r.Flag_TotalCoaching_pct)};font-weight:700">${r.Flag_TotalCoaching_count}<br><small>${r.Flag_TotalCoaching_pct}%</small></td></tr>`;
  });
  body+='</tbody>';
  t.innerHTML=hdr+body;
}

function showTip(e,flag,cnt,pct,ws){
  const g=FLAG_GUIDANCE[flag]; if(!g) return;
  const tip=document.getElementById('tooltip');
  tip.innerHTML=`<div class="tip-title">${g.icon} ${FLAG_LABELS[flag]}</div><div><strong>Workstream:</strong> ${ws}</div><div><strong>Count:</strong> ${cnt} (${pct}%)</div><br><div class="tip-do">\u2705 <strong>DO:</strong> ${g.do}</div><br><div class="tip-avoid">\u274c <strong>AVOID:</strong> ${g.avoid}</div>`;
  tip.style.display='block';
  let x=e.clientX+15,y=e.clientY+15;
  if(x+350>window.innerWidth) x=e.clientX-360;
  if(y+200>window.innerHeight) y=e.clientY-210;
  tip.style.left=x+'px'; tip.style.top=y+'px';
}
document.addEventListener('click',function(e){if(!e.target.classList.contains('heat-cell'))document.getElementById('tooltip').style.display='none';});

function renderCoachingTable(r){
  const t=document.getElementById('coachingTable');
  let html='<thead><tr><th></th><th>Flag</th><th>What It Means</th><th>What To Do</th><th>What To Avoid</th><th>Count</th><th>%</th></tr></thead><tbody>';
  FLAG_COLS.forEach(f=>{
    const g=FLAG_GUIDANCE[f]||{};
    html+=`<tr><td class="flag-icon">${g.icon||''}</td><td><strong>${FLAG_LABELS[f]}</strong></td><td>${g.means||''}</td><td>${g.do||''}</td><td>${g.avoid||''}</td><td>${r[f+'_count']}</td><td>${r[f+'_pct']}%</td></tr>`;
  });
  html+='</tbody>'; t.innerHTML=html;
}

function renderErrorTable(r){
  const t=document.getElementById('errorTable');
  const ec=r.error_categories||{};
  const sorted=Object.entries(ec).sort((a,b)=>b[1]-a[1]);
  let html='<thead><tr><th>Error Category</th><th>Count</th><th>% of Cases</th></tr></thead><tbody>';
  sorted.forEach(([cat,cnt])=>{html+=`<tr><td style="text-align:left">${cat}</td><td>${cnt}</td><td>${(cnt/r.total_cases*100).toFixed(1)}%</td></tr>`;});
  if(sorted.length===0) html+='<tr><td colspan="3">No data</td></tr>';
  html+='</tbody>'; t.innerHTML=html;
}

function renderDonut(r){
  const c=document.getElementById('donutContainer');
  const pri=r.priority_dist||{};
  const total=Object.values(pri).reduce((a,b)=>a+b,0);
  if(total===0){c.innerHTML='<p>No priority data</p>';return;}
  let svg='<svg width="200" height="200" viewBox="0 0 42 42"><circle cx="21" cy="21" r="15.91549430918954" fill="transparent" stroke-width="0.5" stroke="#e0e0e0"/>';
  let offset=25;
  const entries=Object.entries(pri).sort((a,b)=>b[1]-a[1]);
  entries.forEach(([p,cnt])=>{
    const pct=cnt/total*100;
    const color=PRIORITY_COLORS[p]||'#6c757d';
    svg+=`<circle cx="21" cy="21" r="15.91549430918954" fill="transparent" stroke="${color}" stroke-width="8" stroke-dasharray="${pct} ${100-pct}" stroke-dashoffset="${-offset}" stroke-linecap="round"/>`;
    offset-=pct;
  });
  svg+='<text x="21" y="20" text-anchor="middle" font-size="5" font-weight="700" fill="#1a1a2e">'+total+'</text><text x="21" y="24.5" text-anchor="middle" font-size="2.5" fill="#666">cases</text></svg>';
  let legend='<div>';
  entries.forEach(([p,cnt])=>{
    const color=PRIORITY_COLORS[p]||'#6c757d';
    legend+=`<div class="legend-item"><div class="legend-color" style="background:${color}"></div>${p}: ${cnt} (${(cnt/total*100).toFixed(1)}%)</div>`;
  });
  legend+='</div>';
  c.innerHTML=svg+legend;
}

function renderSummary(r){
  const s=document.getElementById('summaryBlock');
  let html=`<strong>${r.workstream}</strong> \u2014 ${r.total_cases} cases, ${r.cases_with_flags} flagged (${r.cases_with_flags_pct}%).<br>`;
  if(r.top_flag && r.top_flag!=='None'){
    const g1=FLAG_GUIDANCE[r.top_flag_key]||{};
    html+=`<br>\ud83d\udd34 <strong>Top Issue: ${r.top_flag}</strong> (${r[r.top_flag_key+'_count']} cases) \u2014 ${g1.means||''}<br>&nbsp;&nbsp;&nbsp;\u2705 ${g1.do||''}<br>&nbsp;&nbsp;&nbsp;\u274c ${g1.avoid||''}`;
  }
  if(r.second_flag && r.second_flag!==''){
    const g2=FLAG_GUIDANCE[r.second_flag_key]||{};
    html+=`<br><br>\ud83d\udfe0 <strong>Second Issue: ${r.second_flag}</strong> (${r[r.second_flag_key+'_count']} cases) \u2014 ${g2.means||''}<br>&nbsp;&nbsp;&nbsp;\u2705 ${g2.do||''}<br>&nbsp;&nbsp;&nbsp;\u274c ${g2.avoid||''}`;
  }
  html+=`<br><br>\u23f1 Avg turnaround: <strong>${r.avg_days_customer} days</strong> at customer, <strong>${r.avg_days_sap} days</strong> at SAP.`;
  s.innerHTML=html;
}

initDropdown();
applyFilter();
</script>
</body>
</html>"""

    return html


# --- Main -----------------------------------------------------------------

def main():
    print("=" * 60)
    print("OSS Coaching Dashboard Builder (Multi-File)")
    print("=" * 60)

    # Step 1: Load all files
    df = load_all_files(DATA_DIR)
    file_count = df["_source_file"].nunique()

    # Step 2: Deduplicate
    before_dedup = len(df)
    df = deduplicate(df)
    after_dedup = len(df)
    dedup_info = f"{before_dedup - after_dedup} duplicates removed" if before_dedup > after_dedup else "no duplicates"

    # Step 3: Clean
    print("[3/5] Cleaning data ...")
    df = clean(df)
    date_range = compute_date_range(df)
    total_cases = len(df)
    print(f"       {total_cases:,} cases | {date_range}")

    # Step 4: Aggregate
    print("[4/5] Aggregating by workstream ...")
    data = aggregate(df)
    print(f"       {len(data) - 1} workstreams + All Workstreams")

    # Step 5: Generate HTML
    print("[5/5] Generating HTML dashboard ...")
    html = build_html(data, date_range, total_cases, file_count, dedup_info)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(html)

    size_kb = round(os.path.getsize(OUTPUT_FILE) / 1024, 1)
    print(f"\n{'=' * 60}")
    print(f"DONE!")
    print(f"  Files processed:  {file_count}")
    print(f"  Total cases:      {total_cases:,}")
    print(f"  Date range:       {date_range}")
    print(f"  Deduplication:    {dedup_info}")
    print(f"  Output:           {OUTPUT_FILE} ({size_kb} KB)")
    print(f"{'=' * 60}")


if __name__ == "__main__":
    main()
