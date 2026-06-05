## 🚀 Quick Start Guide

### Prerequisites
- Python 3.8+
- pandas
- openpyxl

### Installation

```bash
# Install dependencies
pip install pandas openpyxl
```

---

## ⚡ Quick Setup (5 minutes)

### Step 1: Initialize Excel Workbook
```bash
python init_excel_workbook.py
```
Creates: `Data/consolidated.xlsx` with 4 sheets (Raw Data, Historical Data, Insights, Metadata)

### Step 2: Add Your Data
1. Open `Data/consolidated.xlsx`
2. Go to **"Raw Data"** sheet
3. Paste your case data (starting from row 2)
4. Save the file

### Step 3: Process Data
```bash
python process_data.py
```
✓ Reads Raw Data → ✓ Cleans & validates → ✓ Appends to Historical Data → ✓ Generates Insights

### Step 4: Generate Dashboard
```bash
python build_dashboard.py
```
✓ Creates interactive HTML dashboard

### Step 5: View Results
1. Open `output/index.html` in your browser
2. Filter by workstream
3. Click heatmap cells for coaching tips
4. Review KPIs, trends, and error categories

---

## 📅 Regular Workflow

**Weekly/Monthly Data Update:**

```bash
# 1. Add new data to Raw Data sheet in Data/consolidated.xlsx
# 2. Run processing
python process_data.py

# 3. Rebuild dashboard
python build_dashboard.py

# 4. Commit changes (if using Git)
git add Data/consolidated.xlsx output/
git commit -m "Update: new case data and insights"
```

---

## 🤖 Automatic Updates (GitHub Actions)

The workflow runs **automatically**:
- **Daily** at 2 AM UTC
- **On demand** via Actions tab
- **On file change** when you push new data to Data folder

No manual scheduling needed!

---

## 📊 What You Get

### Excel File (`Data/consolidated.xlsx`)
- **Raw Data**: Your input sheet
- **Historical Data**: Complete deduplicated history
- **Insights**: Auto-generated KPIs
- **Metadata**: Processing info

### Dashboard (`output/index.html`)
- Interactive workstream filter
- Heatmap with coaching guidance
- Priority distribution chart
- Error category breakdown
- Time metrics summary
- Actionable coaching summary

---

## 🎯 Current Metrics

| Metric | Updated | Location |
|--------|---------|----------|
| Total Cases | process_data.py | Insights sheet + Dashboard KPI |
| Flagged Cases | process_data.py | Insights sheet + Dashboard KPI |
| Coaching Flags (7 types) | process_data.py | Insights table + Heatmap |
| Avg Days @ Customer | process_data.py | Insights sheet + Dashboard KPI |
| Avg Days @ SAP | process_data.py | Insights sheet + Dashboard KPI |
| Priority Distribution | process_data.py | Insights table + Donut chart |
| Error Categories (Top 10) | process_data.py | Error Category Breakdown table |
| Top Coaching Issues | process_data.py | Insights + Summary |

---

## 🔄 Data Flow

```
Your Data
   ↓
Raw Data Sheet (paste here)
   ↓
process_data.py
   ├→ Clean & validate
   ├→ Historical Data Sheet (auto-append)
   ├→ Deduplicate
   └→ Insights Sheet (generate KPIs)
   ↓
build_dashboard.py
   ├→ Read Historical Data
   ├→ Generate HTML
   └→ Output/index.html
   ↓
View in Browser / Share
```

---

## 📝 Data Format Requirements

### Required Columns (Raw Data sheet)
- Case No. (unique identifier)
- Creation Date (YYYY-MM-DD or system date format)
- Closing Date (YYYY-MM-DD or system date format)
- Workstream (text, e.g., "Finance", "HR")
- Priority (text, e.g., "1 - Very High", "2 - High")
- Error Category (text)
- Days_AtCustomer (number)
- Days_AtSAP (number)
- Flag_HowToOrConsulting (0 or 1)
- Flag_PriorityInflation (0 or 1)
- Flag_CustomerDelay (0 or 1)
- Flag_ComponentMismatch (0 or 1)
- Flag_Reopened (0 or 1)
- Flag_AutoConfirmed (0 or 1)
- Flag_TotalCoaching (0 or 1)

---

## 🛠️ Troubleshooting

| Problem | Solution |
|---------|----------|
| No files found | Run `python init_excel_workbook.py` |
| Column not found | Check column names match requirements |
| Dashboard empty | Verify data in Historical Data sheet |
| Automation not running | Check GitHub Actions > daily-process.yml |
| Permission denied | Add execute: `chmod +x *.py` |

---

## 📖 Full Documentation

See `DATA_STRUCTURE.md` for:
- Detailed sheet explanations
- Advanced configuration
- Backup strategy
- Dashboard features
- Future enhancements

---

**Version**: 1.0 | **Last Updated**: 2026-06-05
