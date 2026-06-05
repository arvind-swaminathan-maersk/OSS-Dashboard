## 🚀 Automated Workflow - Excel File Drop & Process

Your OSS Dashboard now operates **fully automatically**. Just drop an Excel file in the Data folder and everything else happens automatically!

---

## 📋 How It Works

### Local Development (Continuous Monitoring)

```bash
# Terminal 1: Start the file watcher
python watch_and_process.py

# Now whenever you drop an Excel file in Data/ folder:
# 1. File is detected
# 2. Data is read & cleaned
# 3. Appended to consolidated.xlsx
# 4. Deduplicated
# 5. Insights generated
# 6. Dashboard rebuilt
# 7. Original file moved to processed/ folder
```

**That's it!** No manual scripts to run. Just place the file and watch the magic happen.

### Cloud/CI-CD (GitHub Actions)

When you push an Excel file to the Data folder:

```bash
git add Data/Closed_Cases_May2024.xlsx
git commit -m "Add May 2024 closed cases data"
git push
```

GitHub Actions **automatically**:
1. Detects the new file
2. Processes it with `watch_and_process.py`
3. Updates consolidated.xlsx
4. Rebuilds the dashboard
5. Commits updated output/index.html back to repo

---

## 🎯 Complete Workflow

### Scenario 1: Local Development

```
1. Export closed cases to Excel (e.g., from SAP)
   └─→ Closed_Cases_2024_05_15.xlsx

2. Place file in Data/ folder
   └─→ Data/Closed_Cases_2024_05_15.xlsx

3. watch_and_process.py detects it
   ├─→ Reads the file
   ├─→ Cleans & validates data
   ├─→ Appends to consolidated.xlsx Historical Data sheet
   ├─→ Deduplicates on Case No.
   ├─→ Generates Insights sheet
   ├─→ Builds dashboard
   └─→ Moves file to Data/processed/

4. Check console for status
   └─→ ✅ Processing complete!

5. View results
   ├─→ Data/consolidated.xlsx (updated)
   ├─→ Output/index.html (refreshed)
   └─→ Data/process.log (detailed log)
```

### Scenario 2: Cloud/GitHub

```
1. Push new Excel extract to Data folder
   git push origin main

2. GitHub Actions detects change
   ├─→ Automatically triggers workflow
   ├─→ Runs watch_and_process.py --single-run
   ├─→ Builds dashboard
   ├─→ Commits changes back

3. View workflow run
   └─→ Actions tab > Daily Data Processing

4. Dashboard automatically updated
   └─→ output/index.html with latest data
```

---

## 🔧 Two Modes of Operation

### Mode 1: Local Continuous Watcher (Recommended)

```bash
pip install pandas openpyxl watchdog
python watch_and_process.py
```

**Best for:**
- Development
- Interactive testing
- Local file processing
- Real-time feedback

**Features:**
- Runs continuously
- Instant file detection
- Live console logging
- Ctrl+C to stop

**Output:**
```
[2026-06-05 14:32:00] 🔔 New file detected: Closed_Cases_May.xlsx
[2026-06-05 14:32:02] [1/7] Reading data
[2026-06-05 14:32:03] Read 1,245 rows from file
[2026-06-05 14:32:04] [2/7] Cleaning data
[2026-06-05 14:32:05] [3/7] Reading historical data
[2026-06-05 14:32:06] [4/7] Combining and deduplicating
[2026-06-05 14:32:07] Deduplicated: removed 23 duplicate case(s)
[2026-06-05 14:32:08] Historical data now contains 5,847 unique cases
[2026-06-05 14:32:09] [5/7] Creating backup
[2026-06-05 14:32:10] Created backup: consolidated_20260605_143210.xlsx
[2026-06-05 14:32:11] [6/7] Updating consolidated.xlsx
[2026-06-05 14:32:12] Wrote 5,847 rows to Historical Data sheet
[2026-06-05 14:32:13] Updated Insights sheet with 5,847 cases
[2026-06-05 14:32:14] [7/7] Building dashboard
[2026-06-05 14:32:45] ✅ Dashboard built successfully
[2026-06-05 14:32:46] Moved processed file to: Closed_Cases_May_20260605_143246.xlsx
======================================================================
✅ Processing complete!
======================================================================
```

### Mode 2: GitHub Actions (Hands-Off)

```yaml
# Triggered automatically when:
# 1. Excel file pushed to Data folder
# 2. Daily schedule (2 AM UTC)
# 3. Manual trigger from Actions tab
```

**Best for:**
- Production
- Team collaboration
- Audit trail (Git history)
- Scheduled processing

**No action needed:**
- Just push files to repo
- Actions handles everything
- Results auto-committed

---

## 📊 Data Flow

```
Your Excel Extract (from SAP, etc.)
        ↓
    Drop in Data/ folder
        ↓
File Watcher Detects File
(watch_and_process.py)
        ↓
    Read & Clean Data
        ├─→ Parse dates, numbers, flags
        ├─→ Validate columns exist
        └─→ Add metadata (timestamp, source file)
        ↓
Read Historical Data from consolidated.xlsx
        ↓
Combine & Deduplicate
    ├─→ Merge new + old data
    ├─→ Remove Case No. duplicates
    └─→ Keep most recent entry per case
        ↓
Create Backup
    └─→ Data/backups/consolidated_YYYYMMDD_HHMMSS.xlsx
        ↓
Write Updated Data
    └─→ consolidated.xlsx Historical Data sheet
        ↓
Generate Insights
    ├─→ Total cases, date range
    ├─→ Coaching flags (counts & percentages)
    ├─→ Time metrics (avg days)
    ├─→ Top coaching issues
    └─→ Error categories
        ↓
Build Dashboard
    ├─→ Read data from consolidated.xlsx
    ├─→ Generate HTML with interactive features
    └─→ Output/index.html
        ↓
Move Processed File
    └─→ Data/processed/Closed_Cases_May_*.xlsx
        ↓
    ✅ DONE!
```

---

## 📁 Folder Structure After Processing

```
Data/
├── consolidated.xlsx              ← Main workbook (auto-updated)
├── process.log                    ← Processing log (auto-appended)
├── backups/
│   ├── consolidated_20260605_143210.xlsx
│   ├── consolidated_20260605_150445.xlsx
│   └── consolidated_20260605_152018.xlsx
└── processed/
    ├── Closed_Cases_May_20260605_143246.xlsx
    └── Closed_Cases_June_20260605_144512.xlsx

Output/
└── index.html                     ← Dashboard (auto-refreshed)
```

---

## ⚙️ Configuration

### Adjust Watch Folder (in watch_and_process.py)

```python
DATA_FOLDER = "Data"        # Watch this folder
WATCH_EXTENSIONS = {'.xlsx', '.xls'}  # Watch these file types
```

### Adjust GitHub Actions Schedule (in workflows/daily-process.yml)

```yaml
schedule:
  - cron: '0 2 * * *'  # Daily at 2 AM UTC
  # Examples:
  # '0 8 * * *'     → Daily at 8 AM UTC
  # '0 0 * * 0'     → Sundays at midnight UTC
  # '0 */6 * * *'   → Every 6 hours
```

---

## 🎯 Expected Excel Format

When you drop an Excel file, it should have columns:

| Column | Type | Example |
|--------|------|---------|
| Case No. | Text | CAS-2024-001234 |
| Creation Date | Date | 2024-05-01 |
| Closing Date | Date | 2024-05-15 |
| Workstream | Text | Finance |
| Priority | Text | 1 - Very High |
| Error Category | Text | Authorization Error |
| Days_AtCustomer | Number | 5 |
| Days_AtSAP | Number | 8 |
| Flag_HowToOrConsulting | 0/1 | 0 |
| Flag_PriorityInflation | 0/1 | 1 |
| Flag_CustomerDelay | 0/1 | 0 |
| Flag_ComponentMismatch | 0/1 | 0 |
| Flag_Reopened | 0/1 | 0 |
| Flag_AutoConfirmed | 0/1 | 0 |
| Flag_TotalCoaching | 0/1 | 1 |

---

## 🚨 Troubleshooting

### Issue: File not being processed

**Check:**
1. File is in `Data/` folder (not subfolder)
2. File extension is `.xlsx` or `.xls`
3. watch_and_process.py is running (see console output)
4. Check `Data/process.log` for errors

### Issue: GitHub Actions not triggering

**Check:**
1. File was added/modified in `Data/` folder
2. Committed and pushed to main branch
3. Check **Actions** tab for workflow run status
4. View job logs for detailed error messages

### Issue: Columns not found error

**Fix:**
1. Ensure Excel file has required columns (see table above)
2. Column names must match exactly
3. If column names differ, update them before placing in Data folder

### Issue: Deduplication removing too many rows

**Check:**
1. Verify Case No. column has unique values
2. Check Closing Date is correct (used for keeping latest)
3. Review `Data/process.log` for deduplicate details

---

## 📋 Processing Steps (What Happens Automatically)

When you drop an Excel file:

1. **Detection** (< 1 sec)
   - File system watcher detects new file

2. **Reading** (1-5 sec)
   - Excel file is read with pandas
   - Metadata added (timestamp, source file)

3. **Cleaning** (1-2 sec)
   - Data types validated
   - Dates parsed
   - Flags converted to 0/1

4. **Historical Append** (2-5 sec)
   - Combined with existing historical data
   - Deduplication on Case No.

5. **Insights Generation** (1-2 sec)
   - KPIs calculated
   - Insights sheet updated

6. **Dashboard Build** (10-30 sec)
   - HTML generated
   - Interactive features created

7. **Archival** (1 sec)
   - Original file moved to processed/
   - Backup created

**Total Time: ~20-50 seconds**

---

## ✅ Verification

After placing a file in Data/:

**Check consolidated.xlsx:**
```
1. Open Data/consolidated.xlsx
2. Historical Data sheet → should have more rows
3. Insights sheet → metrics should be updated
4. Metadata sheet → "Last Updated" should be recent
```

**Check Dashboard:**
```
1. Open Output/index.html in browser
2. KPI cards should show updated numbers
3. Heatmap should show current data
4. Workstream filter should work
```

**Check Logs:**
```
1. Open Data/process.log
2. Should show processing steps
3. Look for ✅ indicators
4. No ERROR messages = success
```

---

## 🔐 Data Safety

**Automatic Protections:**
- ✅ Backups created before each update
- ✅ Deduplication prevents data loss
- ✅ Processing log tracks all changes
- ✅ Processed files moved to archive folder
- ✅ Original Excel not modified

**You can always:**
1. Roll back to a backup
2. Review process.log for audit trail
3. Check processed/ folder for originals

---

## 🎓 Usage Patterns

### Weekly Coaching Review
```
Monday 9 AM:
1. Export closed cases from SAP → Closed_Cases_W21.xlsx
2. Drop in Data/ folder
3. Watch the console for completion
4. Open Output/index.html
5. Review insights with team
6. File automatically moved to Data/processed/
```

### Monthly Reporting
```
End of month:
1. Export all month's cases
2. Drop in Data/ folder (watch processes it)
3. consolidated.xlsx now contains month's history
4. Generate monthly report from Insights sheet
5. Share HTML dashboard with stakeholders
```

### Automated Daily Updates (GitHub)
```
Every night at 2 AM UTC:
1. GitHub Actions runs automatically
2. Any pending files are processed
3. Dashboard updated
4. Changes auto-committed
5. Team sees latest data each morning
```

---

## 📞 Quick Reference

| Task | Command/Action |
|------|-----------------|
| Start local watcher | `python watch_and_process.py` |
| Process single file manually | `python process_data.py Data/file.xlsx` |
| Build dashboard | `python build_dashboard.py` |
| View logs | `cat Data/process.log` |
| Check backups | `ls Data/backups/` |
| View processed files | `ls Data/processed/` |

---

**Status**: ✅ Fully Automated
**Setup Time**: < 2 minutes
**Learning Curve**: Drop file → that's it!
**Support**: Check Data/process.log for any issues

---

*Updated: 2026-06-05*
*Mode: Automatic File-Based Processing*
