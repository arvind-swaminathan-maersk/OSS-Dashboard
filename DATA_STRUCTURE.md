## 📊 OSS Dashboard Data Structure & Automation Guide

This document explains the new multi-sheet Excel structure and automated data processing workflow.

---

## 🏗️ File Structure Overview

```
Data/
├── consolidated.xlsx          ← Main workbook (4 sheets)
├── backups/                   ← Automatic backups
│   └── consolidated_*.xlsx
└── process.log                ← Processing log
```

### Excel Workbook Sheets

#### 1. **Raw Data** 📥
- **Purpose**: Input sheet where you paste new/updated case data
- **Columns**: Case No., Creation Date, Closing Date, Workstream, Priority, Error Category, Days_AtCustomer, Days_AtSAP, Flag_* columns
- **When to use**: Each time you have new data to process
- **Action**: 
  1. Paste your latest data here
  2. Run `python process_data.py`
  3. Data automatically moves to Historical Data

#### 2. **Historical Data** 📚
- **Purpose**: Complete deduplicated history of all cases over time
- **Columns**: Same as Raw Data + `_processed_date` (auto-added)
- **Features**:
  - Automatically appended by `process_data.py`
  - Duplicates removed (keeps most recent entry per Case No.)
  - Includes processing timestamp
  - Forms the basis for trend analysis
- **Note**: Do not manually edit; managed by automation

#### 3. **Insights** 📊
- **Purpose**: Auto-generated KPIs and metrics
- **Content**:
  - Overall metrics (total cases, date range, flagged cases)
  - Coaching flags breakdown (count & percentage)
  - Time metrics (avg days at customer/SAP)
  - Top coaching issues
- **Updated**: Every time `process_data.py` runs
- **Use**: Quick reference for current metrics

#### 4. **Metadata** 🔍
- **Purpose**: Tracking and processing information
- **Content**:
  - Last Updated timestamp
  - Total rows processed
  - Data quality status
  - Processing instructions
- **Updated**: Automatically by `process_data.py`

---

## 🚀 Workflow: From Data to Dashboard

### Initial Setup (One-time)

```bash
# 1. Initialize the Excel workbook
python init_excel_workbook.py
```

This creates `Data/consolidated.xlsx` with all 4 sheets properly formatted.

### Regular Data Processing Flow

```
┌─────────────────────────────────────────────────────────┐
│ 1. Add Data to "Raw Data" Sheet                          │
│    (Paste your latest case export)                       │
└────────────────┬────────────────────────────────────────┘
                 │
┌────────────────▼────────────────────────────────────────┐
│ 2. Run Data Processing: python process_data.py           │
│    ✓ Reads Raw Data                                      │
│    ✓ Cleans & validates                                  │
│    ✓ Appends to Historical Data with timestamp           │
│    ✓ Deduplicates on Case No.                            │
│    ✓ Generates Insights sheet                            │
│    ✓ Creates backup                                      │
│    ✓ Logs to process.log                                 │
└────────────────┬────────────────────────────────────────┘
                 │
┌────────────────▼────────────────────────────────────────┐
│ 3. Build Dashboard: python build_dashboard.py            │
│    ✓ Reads all .xlsx files in Data/ folder              │
│    ✓ Consolidates into master dataset                   │
│    ✓ Generates interactive HTML dashboard               │
│    ✓ Output: Output/index.html                          │
└────────────────┬────────────────────────────────────────┘
                 │
┌────────────────▼────────────────────────────────────────┐
│ 4. Review Dashboard                                      │
│    Open output/index.html in your browser               │
│    ✓ Filter by workstream                               │
│    ✓ Click heatmap cells for coaching advice            │
│    ✓ Review KPIs, error categories, priorities          │
└─────────────────────────────────────────────────────────┘
```

---

## 🔄 Automated Workflow: GitHub Actions

The repository includes a daily automated workflow that:

- **Runs daily** at 2 AM UTC (configurable)
- **Triggers on** Data folder file changes
- **Executes**:
  1. `process_data.py` - Update historical records & insights
  2. `build_dashboard.py` - Regenerate dashboard
  3. Auto-commits changes with `[skip ci]` flag
  4. Saves logs and output as artifacts

### Manual Trigger

To run the workflow manually:
1. Go to **Actions** tab in GitHub
2. Select **Daily Data Processing & Dashboard Update**
3. Click **Run workflow**

### Adjust Schedule

Edit `.github/workflows/daily-process.yml`:

```yaml
schedule:
  - cron: '0 2 * * *'  # Change this to adjust time
```

Common cron expressions:
- `0 8 * * *` - 8 AM UTC
- `0 0 * * 0` - Sundays at midnight UTC
- `0 */6 * * *` - Every 6 hours

---

## 📋 Current Metrics Tracked

### KPIs Generated
- **Total Cases**: Count of all closed cases
- **Flagged Cases**: Cases with ≥1 coaching flag
- **Coaching Flags Breakdown**:
  - How-To / Consulting
  - Priority Inflation
  - Customer Delay
  - Component Mismatch
  - Reopened
  - Auto-Confirmed
  - Total Coaching
- **Time Metrics**:
  - Avg Days @ Customer
  - Avg Days @ SAP
- **Priority Distribution**: P1, P2, P3, P4 mix
- **Error Categories**: Top 10 error types
- **Top Coaching Issues**: #1 and #2 issues by frequency

### By Workstream
All metrics are available both:
- **Aggregate** (All Workstreams combined)
- **Per-Workstream** (filtered by Workstream column)

---

## 📝 Usage Examples

### Example 1: Weekly Data Update

```bash
# Monday morning - new data arrives
1. Export latest closed cases to CSV/Excel
2. Open Data/consolidated.xlsx
3. Paste data into "Raw Data" sheet
4. Save Excel file
5. Run: python process_data.py
6. Run: python build_dashboard.py
7. Share output/index.html with team
```

### Example 2: Monthly Trend Analysis

The **Historical Data** sheet maintains a complete record. Over months/quarters:
- Copy Historical Data to external analytics tool
- Plot coaching flags over time
- Identify trending issues
- Share improvement metrics with leadership

### Example 3: Spot-Check Current Metrics

```bash
# Quick check without new data
1. Open Data/consolidated.xlsx
2. Check "Insights" sheet for current KPIs
3. No need to run scripts - already up to date
```

---

## 🔧 Technical Details

### Data Cleaning Logic

The `process_data.py` script cleans:

- **Flag columns**: Converted to 0/1 integers
- **Numeric columns** (Days_AtCustomer, Days_AtSAP): Filled with 0 if missing
- **String columns** (Workstream, Priority, Error Category): "Unknown" for nulls
- **Date columns**: Parsed to datetime, errors become NaT

### Deduplication Strategy

- **Key**: Case No. column
- **Method**: Keep most recent entry (sorted by Closing Date DESC)
- **Result**: True unique cases over time with latest state

### Excel File Structure

- Uses `openpyxl` library (Python-native)
- Styled headers (dark background, white text)
- No formulas (all values are computed Python-side)
- Human-readable format

### Backup Strategy

- Automatic backup created before each `process_data.py` run
- Location: `Data/backups/consolidated_YYYYMMDD_HHMMSS.xlsx`
- Retention: 7 days (configurable)

---

## 📊 Dashboard Features

The generated HTML dashboard (`output/index.html`):

### Interactive Elements
- **Workstream Filter**: Dropdown to focus on specific workstream
- **Heatmap**: Click any cell for coaching guidance
- **Priority Donut Chart**: Visual priority distribution
- **Error Category Table**: Top 10 errors

### Sections
1. **KPI Cards**: 6 key metrics at a glance
2. **Coaching Heatmap**: Flags × Workstreams with color intensity
3. **Coaching Flags Guide**: What each flag means + do/avoid
4. **Error Categories**: Most common errors
5. **Priority Distribution**: Visual breakdown
6. **Actionable Summary**: Per-workstream takeaways

---

## 🎯 Future Enhancements

### Coming Soon (Optional)
- Additional KPI definitions
- Custom date range filtering
- Export insights to PDF
- Comparison views (month-over-month)
- Trend charts (coaching flags over time)

### Configuration
When ready, modify `process_data.py` or `build_dashboard.py` to:
- Add new FLAG_COLUMNS
- Update FLAG_GUIDANCE
- Add computed metrics
- Change styling

---

## 🐛 Troubleshooting

### Issue: `No .xlsx files found in 'data/' folder`
**Solution**: Run `python init_excel_workbook.py` first, then add data to Raw Data sheet.

### Issue: `Column 'Case No.' not found`
**Solution**: Ensure your data has a "Case No." column. Update column names to match expected schema.

### Issue: Deduplication seems off
**Solution**: Check that "Case No." and "Closing Date" columns exist and are properly formatted.

### Issue: Dashboard not updating
**Solution**: 
1. Check `Data/process.log` for errors
2. Run `python process_data.py` manually
3. Run `python build_dashboard.py` manually
4. Verify `output/index.html` exists

### Issue: GitHub Action failing
**Solution**:
1. Check **Actions** tab for error logs
2. Ensure `Data/consolidated.xlsx` is committed
3. Verify workflow file syntax: `.github/workflows/daily-process.yml`

---

## 📞 Support

For detailed script documentation, see:
- `process_data.py` - Data processing logic
- `build_dashboard.py` - Dashboard generation
- `init_excel_workbook.py` - Workbook initialization
- `.github/workflows/daily-process.yml` - Automation schedule

---

**Last Updated**: 2026-06-05
**Version**: 1.0 - Multi-Sheet Architecture with Automated Insights
