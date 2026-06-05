## ⚡ Automatic Setup - 2 Minutes

Your OSS Dashboard is now **fully automated**. Here's how to activate it:

---

## 🎯 Quick Start - Choose Your Mode

### Option A: Local Development (Recommended)

Perfect for testing, development, and interactive use.

**Step 1: Install dependency**
```bash
pip install watchdog
```

**Step 2: Start the watcher**
```bash
python watch_and_process.py
```

You should see:
```
======================================================================
OSS Dashboard - File Watcher Service Started
======================================================================
Monitoring: Data folder for new Excel files
Action: Auto-process & append to consolidated.xlsx
Press Ctrl+C to stop
======================================================================

✅ Watcher ready. Watching Data/ for new Excel files...
```

**Step 3: Drop Excel files**
- Export closed cases from SAP → Excel file
- Save to `Data/` folder
- Watch console for automatic processing
- Output updates to `consolidated.xlsx` and `output/index.html`

**Done!** The watcher runs continuously. Ctrl+C to stop.

---

### Option B: Cloud/GitHub (Hands-Off)

Perfect for team collaboration and production.

**Step 1: Commit the watcher script**
```bash
git add watch_and_process.py process_data.py workflows/daily-process.yml
git commit -m "Enable automatic data processing"
git push
```

**Step 2: Push Excel files**
```bash
# Export your closed cases to Excel
# Commit to repo
git add Data/Closed_Cases_May2024.xlsx
git commit -m "Add May 2024 closed cases"
git push
```

**Step 3: Done!**
- GitHub Actions detects the file
- Automatically processes it
- Updates dashboard
- Commits results back

No manual intervention needed!

---

## 📊 What Happens Automatically

```
Drop Excel File in Data/
        ↓
File Watcher Detects (< 1 sec)
        ↓
Read & Clean Data (2-5 sec)
        ↓
Append to Historical Data (2-5 sec)
        ↓
Deduplicate (1-2 sec)
        ↓
Generate Insights (1-2 sec)
        ↓
Build Dashboard (10-30 sec)
        ↓
Move File to Processed (1 sec)
        ↓
✅ DONE! (Total: ~20-50 seconds)
```

---

## 📁 What Gets Created

After processing, you'll have:

```
Data/
├── consolidated.xlsx          ← Updated with your new data
├── process.log                ← Detailed processing log
├── backups/
│   └── consolidated_*.xlsx    ← Automatic backups
└── processed/
    └── Your_File_*.xlsx       ← Original file archived

Output/
└── index.html                 ← Fresh dashboard
```

---

## ✅ Verification

**After dropping a file, check:**

1. **Console** (if using local watcher)
   - Should show "Processing complete!" ✅

2. **Data/consolidated.xlsx**
   - Historical Data sheet → more rows
   - Insights sheet → updated KPIs
   - Metadata sheet → recent timestamp

3. **Output/index.html**
   - Open in browser
   - KPI numbers updated
   - Dashboard refreshed

4. **Data/process.log**
   - Shows all processing steps
   - Look for ✅ indicators
   - No ERROR = success

---

## 🔑 Key Features

✅ **Automatic Detection**
- Just drop the file in Data/
- No commands to run

✅ **Historical Tracking**
- All data preserved in consolidated.xlsx
- Complete audit trail in process.log

✅ **Deduplication**
- Duplicates automatically removed
- Keeps most recent entry per case

✅ **Instant Insights**
- KPIs generated immediately
- Dashboard refreshed

✅ **Data Safety**
- Automatic backups
- Original files archived
- No data loss

✅ **Flexible Deployment**
- Works locally (watch_and_process.py)
- Works in cloud (GitHub Actions)
- Works on schedule (cron jobs)

---

## 🚨 If Something Goes Wrong

1. **Check the log file**
   ```bash
   cat Data/process.log
   ```
   - Last entries show what happened
   - ERROR messages explain problems

2. **Verify file format**
   - Must have required columns (see AUTOMATION_GUIDE.md)
   - Excel file must be readable

3. **Check file location**
   - File must be in `Data/` folder
   - Not in subfolders

4. **Restart the watcher**
   ```bash
   # Stop with Ctrl+C
   # Then:
   python watch_and_process.py
   ```

---

## 📋 File Requirements

Your Excel file needs these columns:

- Case No.
- Creation Date
- Closing Date
- Workstream
- Priority
- Error Category
- Days_AtCustomer
- Days_AtSAP
- Flag_HowToOrConsulting (0 or 1)
- Flag_PriorityInflation (0 or 1)
- Flag_CustomerDelay (0 or 1)
- Flag_ComponentMismatch (0 or 1)
- Flag_Reopened (0 or 1)
- Flag_AutoConfirmed (0 or 1)
- Flag_TotalCoaching (0 or 1)

---

## 🎓 Common Usage

### Local Testing
```bash
# Start watcher
python watch_and_process.py

# In another terminal:
# Drop files in Data/ and watch them process
```

### Daily Updates (Team)
```bash
# Push new data to GitHub
git push origin main

# Actions automatically processes it
# Check Actions tab for status
# Dashboard updates on GitHub Pages
```

### Monthly Review
```bash
# End of month:
# 1. Export all month's cases
# 2. Drop in Data/
# 3. consolidated.xlsx has full month history
# 4. Share dashboard with team
```

---

## 📞 Documentation

- **How does it work?** → See AUTOMATION_GUIDE.md
- **Data structure?** → See DATA_STRUCTURE.md
- **What metrics?** → See build_dashboard.py comments
- **Troubleshooting?** → See AUTOMATION_GUIDE.md troubleshooting section

---

## ✨ You're All Set!

1. ✅ Installed watchdog: `pip install watchdog`
2. ✅ Started watcher: `python watch_and_process.py`
3. ✅ Understood the workflow: Drop file → auto-process
4. ✅ Know how to verify: Check consolidated.xlsx + dashboard

**Now just:**
- Drop Excel files in Data/
- Watch the magic happen
- View updated insights

That's it! No more manual scripts to run. 🎉

---

**Need Help?**
- Check `Data/process.log` for errors
- Review `AUTOMATION_GUIDE.md` for detailed info
- Ensure Excel has required columns

**Questions?**
- See inline comments in `watch_and_process.py`
- Check GitHub Actions logs for cloud issues
- Review `DATA_STRUCTURE.md` for architecture

---

*Setup Complete: 2026-06-05*
*Mode: Fully Automated*
*Ready to Process: Yes ✅*
