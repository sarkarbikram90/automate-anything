# TASK SCHEDULER SETUP GUIDE
## Windows Automation for Excel to MySQL Script

---

## PREREQUISITES

Before setting up Task Scheduler, ensure you have:

1. **Python 3.8+** installed and added to system PATH
   - Test: Open Command Prompt and run `python --version`
   - If not found, download from https://www.python.org
   - During installation, CHECK "Add Python to PATH"

2. **Required Python packages installed**
   ```
   pip install python-dotenv openpyxl mysql-connector-python exchangelib
   ```

3. **Configuration file (.env) created**
   - Copy `.env.example` to `.env` in the same directory as the script
   - Fill in your email, password, and database credentials
   - Verify the .env file exists in the script directory

4. **MySQL database setup**
   - Create your database if it doesn't exist
   - Create the target table with appropriate columns
   - Test the connection with MySQL Workbench or similar tool

---

## STEP 1: CREATE A BATCH FILE (OPTIONAL BUT RECOMMENDED)

Creating a batch file makes it easier to run the script and capture output.

**File: run_excel_automation.bat**

```batch
@echo off
REM Excel to MySQL Automation
REM Change to script directory
cd /d "C:\Path\To\Your\Script\Directory"

REM Run the Python script
python excel_to_mysql_automation.py

REM Pause to see any errors (remove this in production)
REM pause
```

**Save this file in the same directory as your Python script**

---

## STEP 2: OPEN TASK SCHEDULER

### Method 1: Search Method
1. Click the **Windows Start button**
2. Type: `Task Scheduler`
3. Click "Task Scheduler" in the results

### Method 2: Run Command
1. Press **Win + R**
2. Type: `taskschd.msc`
3. Press Enter

---

## STEP 3: CREATE A NEW TASK

1. In Task Scheduler, on the right panel, click **"Create Task..."**

2. **General Tab:**
   - **Name:** `Excel to MySQL Sync`
   - **Description:** `Automated Excel file import from email to MySQL database`
   - ☑ **Check: "Run whether user is logged in or not"**
   - ☑ **Check: "Run with highest privileges"**
   - **Configure for:** Windows 10 (or your OS)

   ![General Tab Settings]
   ```
   ┌─────────────────────────────────┐
   │ Name: Excel to MySQL Sync       │
   │ Description: Automated Excel... │
   │ ☑ Run whether user logged in    │
   │ ☑ Run with highest privileges   │
   │ Configure for: Windows 10       │
   └─────────────────────────────────┘
   ```

3. Click **"Next"** (or go to **Triggers** tab)

---

## STEP 4: SET UP TRIGGERS (WHEN TO RUN)

### For Daily Execution at Specific Time

1. Go to **Triggers** tab
2. Click **New...**
3. **Begin the task:** Select `On a schedule`
4. **Settings:**
   - **Daily**
   - **Start:** Today's date
   - **Recur every:** 1 day
   - **At:** `09:00:00` (9 AM - adjust as needed)
   - ☑ **Enabled**

5. Click **OK**

### For Hourly Execution

1. Go to **Triggers** tab
2. Click **New...**
3. **Begin the task:** Select `On a schedule`
4. **Settings:**
   - **Repeat task every:** `1 hour` (or your preferred interval)
   - **For a duration of:** `23 hours` (runs all day)
   - **Stop all running instances at:** Check if needed

5. Click **OK**

### For Specific Days Only

1. Go to **Triggers** tab
2. Click **New...**
3. **Begin the task:** Select `On a schedule`
4. **Settings:**
   - **Weekly**
   - **Start:** Today's date
   - **Recur every:** 1 week
   - **At:** `09:00:00`
   - ☑ Select days: Mon, Tue, Wed, Thu, Fri (uncheck weekends if not needed)

5. Click **OK**

---

## STEP 5: SET UP ACTIONS (WHAT TO RUN)

1. Go to **Actions** tab
2. Click **New...**

### Option A: Using Batch File (RECOMMENDED)

1. **Action:** `Start a program`
2. **Program/script:** Browse and select `run_excel_automation.bat`
3. **Add arguments:** (leave blank)
4. **Start in:** `C:\Path\To\Your\Script\Directory`
5. Click **OK**

### Option B: Direct Python Execution

1. **Action:** `Start a program`
2. **Program/script:** `C:\Users\YourUsername\AppData\Local\Programs\Python\Python311\python.exe`
   (Adjust Python version number)
3. **Add arguments:** `C:\Path\To\excel_to_mysql_automation.py`
4. **Start in:** `C:\Path\To\Your\Script\Directory`
5. Click **OK**

---

## STEP 6: CONFIGURE CONDITIONS (OPTIONAL)

1. Go to **Conditions** tab
2. **Power:**
   - ☑ Start the task only if the computer is on AC power
   (Optional: uncheck if you need it to run on battery)

3. **Network:**
   - ☑ Start only if the following network connection is available: Any
   (Ensures internet is available for Outlook/email)

4. Click **OK** when done with all tabs

---

## STEP 7: CONFIGURE SETTINGS (OPTIONAL)

1. Go to **Settings** tab
2. **Recommended Settings:**
   - ☑ Allow task to be run on demand
   - ☑ Stop the task if it runs longer than: 01:00:00 (1 hour)
   - ☑ If the task fails, restart every: 10 minutes
   - Repeat up to: 3 times
   - ☑ If the task is already running, then the following rule applies:
     → "Do not start a new instance"

3. Click **OK** to save the task

---

## STEP 8: TEST THE TASK

1. In Task Scheduler, find your task: `Excel to MySQL Sync`
2. Right-click → **Run**
3. Watch the Status column - it should show "Running"
4. Check the logs in your script directory (look for `logs/excel_to_mysql.log`)
5. Verify:
   - Temporary files are downloaded
   - Data is inserted into MySQL
   - Log file shows success/errors

---

## TROUBLESHOOTING

### Task Runs But Nothing Happens

**Solution:**
1. Check the task history: Right-click task → View → Show History
2. Look for "Last Run Result" - if it shows error code, note it
3. Check the log file: `logs/excel_to_mysql.log`
4. Verify .env file exists in the script directory

### "File Not Found" Error

**Solution:**
- Use absolute paths (full paths) in Task Scheduler, not relative paths
- Example: `C:\Users\YourName\Documents\excel_automation\excel_to_mysql_automation.py`
- Not: `./excel_to_mysql_automation.py`

### Office 365 Authentication Failed

**Solution:**
1. Verify you're using an App Password (not regular password)
2. Check if 2FA is enabled on your account
3. Generate a NEW app password and update .env file
4. Ensure password has no special characters or quotes

### Database Connection Failed

**Solution:**
1. Verify MySQL is running: `net start MySQL80` (Windows)
2. Test connection with MySQL Workbench
3. Verify database and table exist
4. Check firewall isn't blocking port 3306
5. Verify credentials in .env file

### Task Runs Too Frequently or Not at All

**Solution:**
- Double-check trigger settings
- Edit task → Triggers tab → Verify schedule
- Make sure "Enabled" checkbox is checked
- Consider time zones in your scheduled times

### Log File Not Created

**Solution:**
1. Verify `logs` folder exists in script directory
2. Create it manually: `mkdir logs`
3. Check folder permissions (script needs write access)
4. Run script manually to verify it creates logs

---

## ADVANCED: RUN TASK IMMEDIATELY (TESTING)

```batch
REM Run task immediately from Command Prompt
taskkill /tn "Excel to MySQL Sync" /f
schtasks /run /tn "Excel to MySQL Sync"
```

---

## ADVANCED: DELETE TASK

```batch
REM Delete the scheduled task
schtasks /delete /tn "Excel to MySQL Sync" /f
```

---

## MONITORING TASK EXECUTION

### View Task History
1. In Task Scheduler, select your task
2. Go to **History** tab
3. Look for recent execution records
4. Double-click entry to see details

### Enable Detailed Logging
In the .env file:
```
NOTIFY_ON_SUCCESS=True
NOTIFY_EMAIL=your-email@domain.com
```

This will email you after each run.

---

## SECURITY BEST PRACTICES

1. **Protect .env file:**
   - Right-click .env → Properties → Security
   - Remove "Authenticated Users" access
   - Only your user account should have access

2. **Use App Passwords:**
   - Never put your actual Outlook password in .env
   - Always use generated App Passwords for Office 365

3. **Run with minimal privileges:**
   - The script only needs access to email and database
   - Don't run as SYSTEM user unnecessarily

4. **Monitor logs regularly:**
   - Check `logs/` folder for security issues
   - Look for failed authentication attempts

---

## EXAMPLE SCHEDULE SCENARIOS

### Scenario 1: Daily at 9 AM
- **Trigger:** On a schedule, Daily, 09:00:00
- **Use case:** Process emails received overnight

### Scenario 2: Every 4 Hours
- **Trigger:** On a schedule, Weekly, every 4 hours, 00:00:00
- **Use case:** Real-time processing throughout the day

### Scenario 3: Weekdays Only
- **Trigger:** On a schedule, Weekly, 08:00:00, Mon-Fri only
- **Use case:** Business hours automation

### Scenario 4: Immediately When Email Arrives
- Not possible with Task Scheduler alone
- Use Power Automate instead for real-time triggers

---

## QUICK REFERENCE CHECKLIST

- [ ] Python 3.8+ installed and in PATH
- [ ] Required packages installed (`pip install ...`)
- [ ] `.env` file created and configured
- [ ] MySQL database and table created
- [ ] Batch file created (optional)
- [ ] Task Scheduler opened
- [ ] Task created with correct name
- [ ] General tab configured (highest privileges, run whether logged in)
- [ ] Trigger set up (schedule and frequency)
- [ ] Action configured (points to correct script/batch)
- [ ] Conditions set (optional)
- [ ] Settings configured (optional)
- [ ] Task tested manually
- [ ] Log file verified
- [ ] Schedule confirmed working

---

## SUPPORT

If you encounter issues:

1. **Check logs:** `logs/excel_to_mysql.log`
2. **Run manually:** `python excel_to_mysql_automation.py`
3. **View Task History:** Task Scheduler → Select task → History tab
4. **Test components separately:**
   - Test Outlook connection
   - Test MySQL connection
   - Test file reading

---

Last Updated: 2024
Compatible with: Windows 10, Windows 11, Windows Server 2019+
