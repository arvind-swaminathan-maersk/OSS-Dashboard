## 🚀 GitHub Actions Automation - Complete Setup Guide

You'll have **fully automated processing** without installing anything locally. Just push files and GitHub handles everything!

---

## ✅ Prerequisites Check

Before starting, verify you have:

- [ ] GitHub account
- [ ] Repository: `arvind-swaminathan-maersk/OSS-Dashboard`
- [ ] Git installed on your computer (or use GitHub Desktop)
- [ ] Access to push to the repository

---

## 📋 Step-by-Step Setup

### **STEP 1: Enable GitHub Actions** (5 minutes)

1. **Go to your repository**
   - Open browser: https://github.com/arvind-swaminathan-maersk/OSS-Dashboard
   - You should see your repo

2. **Click the "Settings" tab**
   - At the top of the page, next to "Code", "Issues", "Pull requests"
   - Click **Settings**

3. **In the left sidebar, click "Actions"**
   - You'll see a menu on the left
   - Look for **Actions** (may be under "Code and automation")
   - Click **Actions**

4. **Click "General"** (under Actions)
   - Left sidebar under Actions
   - Click **General**

5. **Find "Actions permissions"**
   - Scroll down to find **"Actions permissions"** section
   - You should see three radio button options:
     - "Disable all"
     - "Allow all actions and reusable workflows" ← **Select this one**
     - "Allow enterprise, and select non-enterprise, actions and reusable workflows"

6. **Select "Allow all actions and reusable workflows"**
   - Click the radio button next to it

7. **Scroll down and click "Save"**
   - Blue button at the bottom
   - This saves your settings

✅ **GitHub Actions is now enabled!**

---

### **STEP 2: Verify Workflow File Exists** (2 minutes)

The workflow file should already be in your repository. Let's verify:

1. **Go to your repository main page**
   - https://github.com/arvind-swaminathan-maersk/OSS-Dashboard

2. **Navigate to the workflow file**
   - Click on the **Code** tab
   - Click on **workflows** folder
   - You should see **daily-process.yml** file

   Or directly: https://github.com/arvind-swaminathan-maersk/OSS-Dashboard/blob/main/workflows/daily-process.yml

3. **If you see the file, continue to Step 3**

If the file is missing, don't worry - it was already created earlier.

✅ **Workflow file verified!**

---

### **STEP 3: Prepare Your First Excel File** (5 minutes)

1. **Export your closed cases data**
   - From SAP or your system
   - Save as Excel file (`.xlsx`)
   - Name it something like: `Closed_Cases_May2024.xlsx`

2. **Verify it has the right columns:**
   - Case No.
   - Creation Date
   - Closing Date
   - Workstream
   - Priority
   - Error Category
   - Days_AtCustomer
   - Days_AtSAP
   - Flag_HowToOrConsulting
   - Flag_PriorityInflation
   - Flag_CustomerDelay
   - Flag_ComponentMismatch
   - Flag_Reopened
   - Flag_AutoConfirmed
   - Flag_TotalCoaching

   ✅ If all columns are there, you're ready!

---

### **STEP 4: Clone Repository (If You Haven't Already)** (5 minutes)

You need the repo on your computer to add files.

#### **Option A: Using Git Command Line**

1. **Open Command Prompt or Terminal**
   - Windows: Search for "Command Prompt" or "PowerShell"
   - Mac/Linux: Open Terminal

2. **Navigate to where you want the repo**
   ```bash
   cd Desktop
   ```
   Or any folder you prefer.

3. **Clone the repository**
   ```bash
   git clone https://github.com/arvind-swaminathan-maersk/OSS-Dashboard.git
   ```

4. **Enter the folder**
   ```bash
   cd OSS-Dashboard
   ```

5. **You're ready!** You now have the repo locally.

#### **Option B: Using GitHub Desktop (Easier)**

1. **Install GitHub Desktop**
   - https://desktop.github.com/
   - Download and install

2. **Open GitHub Desktop**

3. **Click "File" → "Clone Repository"**

4. **Enter:**
   - URL: `https://github.com/arvind-swaminathan-maersk/OSS-Dashboard.git`
   - Local Path: Choose a folder on your computer

5. **Click "Clone"**

6. **Done!** You now have the repo locally in the chosen folder.

✅ **Repository cloned!**

---

### **STEP 5: Add Your Excel File to the Data Folder** (3 minutes)

1. **Open the folder where you cloned the repo**
   - If you used command line: Open File Explorer and navigate to the folder
   - If you used GitHub Desktop: Click "Show in Explorer" (right-click on repo name)

2. **Navigate to the "Data" folder**
   - You should see folders: `Data`, `Output`, `workflows`
   - Open the **Data** folder

3. **Copy your Excel file here**
   - Take your `Closed_Cases_May2024.xlsx` file
   - Copy it into the `Data` folder

   ✅ Your file is now in the right place!

---

### **STEP 6: Commit and Push (The Magic Happens Here!)** (5 minutes)

#### **Option A: Using Command Line**

1. **Open Command Prompt/Terminal**

2. **Navigate to your repo**
   ```bash
   cd path/to/OSS-Dashboard
   ```

3. **Check status**
   ```bash
   git status
   ```
   You should see your Excel file listed as "Untracked files"

4. **Add the file**
   ```bash
   git add Data/Closed_Cases_May2024.xlsx
   ```

5. **Commit the file**
   ```bash
   git commit -m "Add May 2024 closed cases data"
   ```

6. **Push to GitHub**
   ```bash
   git push
   ```

   ✅ Your file is now on GitHub!

#### **Option B: Using GitHub Desktop**

1. **Open GitHub Desktop**

2. **You should see your file in the "Changes" tab**
   - If you don't see it, click "Fetch" first

3. **You'll see a summary:**
   - "1 changed file" or similar
   - Shows your Excel file

4. **In the "Summary" box at bottom left, type:**
   ```
   Add May 2024 closed cases data
   ```

5. **Click "Commit to main"**
   - Blue button

6. **Click "Push origin"**
   - Blue button at top
   - This uploads to GitHub

   ✅ Your file is now on GitHub!

---

### **STEP 7: Watch GitHub Actions Process Your File** (1-2 minutes)

1. **Go to your repository on GitHub**
   - https://github.com/arvind-swaminathan-maersk/OSS-Dashboard

2. **Click the "Actions" tab**
   - At the top, next to "Code", "Issues", etc.

3. **You should see a workflow running!**
   - Title: Something like "Add May 2024 closed cases data"
   - Yellow dot = Currently running
   - Green checkmark = Completed successfully
   - Red X = Failed (check logs)

4. **Click on the workflow to see details**
   - Shows each step:
     - ✅ Checkout repository
     - ✅ Set up Python
     - ✅ Install dependencies
     - ✅ Detect and process new Excel files
     - ✅ Build dashboard
     - ✅ Commit and push changes
     - ✅ Upload logs and output

5. **Wait for it to complete** (usually 2-5 minutes)
   - Green checkmark = Success!
   - Check the box below the workflow name

✅ **Your data was processed automatically!**

---

### **STEP 8: View Your Results** (3 minutes)

1. **Pull the latest changes**
   - If using command line:
     ```bash
     git pull
     ```
   - If using GitHub Desktop: Click "Pull origin"

2. **Check the consolidated.xlsx file**
   - Open `Data/consolidated.xlsx`
   - You should see:
     - **Historical Data** sheet: Your rows + any previous data
     - **Insights** sheet: Updated KPIs
     - **Metadata** sheet: Recent timestamp

3. **View the dashboard**
   - Open `Output/index.html` in your browser
   - Should show:
     - Updated KPI cards
     - Heatmap with your data
     - Coaching flags breakdown
     - Error categories
     - Priority distribution

4. **Check the processing log**
   - Open `Data/process.log`
   - Shows everything that happened:
     ```
     [2026-06-05 14:32:00] Processing: Closed_Cases_May2024.xlsx
     [2026-06-05 14:32:02] Read 1,245 rows from file
     [2026-06-05 14:32:05] Historical data now contains 5,847 unique cases
     [2026-06-05 14:32:30] ✅ Dashboard built successfully
     ```

✅ **Everything is working!**

---

## 🔄 From Now On - Regular Usage

Every time you want to add new data:

1. **Export closed cases from SAP** → Excel file
2. **Add to your repo:**
   ```bash
   git add Data/New_Closed_Cases.xlsx
   git commit -m "Add new cases"
   git push
   ```
3. **GitHub Actions automatically:**
   - Processes the file
   - Updates consolidated.xlsx
   - Regenerates dashboard
   - Commits results

4. **Pull the updates:**
   ```bash
   git pull
   ```

5. **View your updated dashboard**
   - Open `Output/index.html`

**That's it! No manual scripts, no watching. Just push → GitHub handles everything.**

---

## 🎯 Checklist - Did You Complete Everything?

- [ ] Step 1: Enabled GitHub Actions
- [ ] Step 2: Verified workflow file exists
- [ ] Step 3: Prepared Excel file with correct columns
- [ ] Step 4: Cloned repository locally
- [ ] Step 5: Added Excel file to Data folder
- [ ] Step 6: Committed and pushed file
- [ ] Step 7: Watched workflow run in Actions tab
- [ ] Step 8: Pulled results and viewed consolidated.xlsx + dashboard

✅ **All done?** You now have fully automated processing!

---

## 🚨 Troubleshooting

### **Workflow not running?**
1. Check "Actions" tab - do you see a workflow listed?
2. If not, go back to Step 1 and verify "Allow all actions" is selected
3. Try pushing a file again

### **Workflow failed (red X)?**
1. Click on the failed workflow
2. Click on "Detect and process new Excel files" step
3. Read the error message
4. Common issues:
   - File doesn't have all required columns → Check the column names
   - File is corrupted → Try exporting again

### **consolidated.xlsx not updating?**
1. Check `Data/process.log` for errors
2. Make sure you **git pulled** the latest changes
3. Verify file was in Data folder (not subdirectory)

### **Dashboard not updating?**
1. Open `Output/index.html` in browser (not cached - use Ctrl+Shift+Del to clear cache)
2. Check the file was processed (look in Actions logs)
3. Make sure you git pulled the latest

---

## 📞 Quick Reference

| Task | Command |
|------|---------|
| Clone repo | `git clone https://github.com/arvind-swaminathan-maersk/OSS-Dashboard.git` |
| Add file | `git add Data/filename.xlsx` |
| Commit | `git commit -m "Add description"` |
| Push (trigger automation) | `git push` |
| Pull results | `git pull` |
| View workflow status | Go to Actions tab on GitHub |

---

## ✨ You're All Set!

You now have:
- ✅ GitHub Actions enabled
- ✅ Workflow ready to process files
- ✅ Knowledge to push files and trigger automation
- ✅ Understanding of where to find results

**Next step:** Push your first Excel file and watch it process automatically!

---

*Setup Time: ~30 minutes total*
*Ongoing Time: 1 minute per data update (just git push)*
*Automation: 100% - No manual intervention needed*

---

**Questions?** Check the troubleshooting section or review the AUTOMATION_GUIDE.md file.
