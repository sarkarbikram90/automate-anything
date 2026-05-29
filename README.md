# QUICK START GUIDE
## Excel to MySQL Automation with Outlook Integration

---

## 📋 REQUIREMENTS

- **Windows 10/11** (or Windows Server)
- **Python 3.8+** (with pip)
- **MySQL 5.7+** or **MariaDB 10.3+**
- **Outlook/Office 365 account** with app passwords enabled

---

## 🚀 INSTALLATION (5 MINUTES)

### Step 1: Install Python Dependencies

```bash
# Navigate to your script directory
cd C:\path\to\your\script

# Install required packages
pip install -r requirements.txt
```

### Step 2: Configure Environment Variables

```bash
# Copy the example file
copy .env.example .env

# Edit .env with your credentials (use Notepad or any text editor)
notepad .env
```

**Required fields in .env:**
- `OUTLOOK_EMAIL` = your outlook email
- `OUTLOOK_PASSWORD` = app password (not regular password!)
- `MYSQL_USER` = database username
- `MYSQL_PASSWORD` = database password
- `MYSQL_DATABASE` = database name

### Step 3: Prepare MySQL Database

```sql
-- Connect to your MySQL database

-- Create database (if needed)
CREATE DATABASE your_database_name;
USE your_database_name;

-- Create table to store Excel data
CREATE TABLE excel_data (
    id INT AUTO_INCREMENT PRIMARY KEY,
    -- Add columns matching your Excel file
    -- Examples:
    name VARCHAR(255),
    email VARCHAR(255),
    phone VARCHAR(20),
    amount DECIMAL(10, 2),
    date_created TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    -- Adjust based on your actual Excel columns
    UNIQUE KEY unique_email (email)
);
```

### Step 4: Test the Script Manually

```bash
# From Command Prompt/PowerShell in script directory
python excel_to_mysql_automation.py
```

**Expected output:**
```
2024-01-15 09:00:00 - ExcelToMySQL - INFO - EXCEL TO MYSQL AUTOMATION STARTED
2024-01-15 09:00:01 - ExcelToMySQL - INFO - Connecting to Outlook...
2024-01-15 09:00:03 - ExcelToMySQL - INFO - ✓ Successfully connected to Outlook
2024-01-15 09:00:05 - ExcelToMySQL - INFO - Searching for emails from past 1 days...
...
```

### Step 5: Set Up Windows Task Scheduler

See **TASK_SCHEDULER_SETUP_GUIDE.md** for detailed instructions.

Quick summary:
1. Open Task Scheduler
2. Create New Task → "Excel to MySQL Sync"
3. Set Trigger → Daily at 9 AM (adjust as needed)
4. Set Action → Run `excel_to_mysql_automation.py`
5. Test by right-clicking → Run

---

## 📁 FILE STRUCTURE

```
your-project/
├── excel_to_mysql_automation.py    ← Main script
├── requirements.txt                ← Python dependencies
├── .env                            ← Configuration (create from .env.example)
├── .env.example                    ← Configuration template
├── run_excel_automation.bat        ← Batch file for Task Scheduler
├── logs/                           ← Created automatically
│   └── excel_to_mysql.log         ← Execution logs
└── temp_files/                     ← Created automatically
    └── [downloaded Excel files]
```

---

## ⚙️ CONFIGURATION OPTIONS

### Basic Configuration (.env)

```env
# Required
OUTLOOK_EMAIL=your-email@outlook.com
OUTLOOK_PASSWORD=xxxxxxxxxxxxxxxx
MYSQL_HOST=localhost
MYSQL_USER=root
MYSQL_PASSWORD=password
MYSQL_DATABASE=your_db
TABLE_NAME=excel_data

# Optional
DAYS_BACK=1                    # How many days to search for emails
NOTIFY_EMAIL=admin@email.com   # Email for notifications
NOTIFY_ON_SUCCESS=True
NOTIFY_ON_ERROR=True
MAX_RETRIES=3
RETRY_DELAY=5
```

### Advanced Configuration

**DAYS_BACK**
- `1` = Only emails from today
- `7` = Emails from past week
- `30` = Emails from past month

**Scheduling Options**
- **Daily at 9 AM:** Task Scheduler → Daily → 09:00
- **Every 4 hours:** Task Scheduler → Weekly → 4-hour interval
- **Only weekdays:** Task Scheduler → Weekly → Mon-Fri

---

## 🔐 SECURITY SETUP

### Get Office 365 App Password

1. Go to https://account.microsoft.com/account/manage-my-microsoft-account
2. Click **Security** → **Advanced security options**
3. **App passwords** → Select "Mail" and "Windows"
4. Copy generated password → Paste in .env as `OUTLOOK_PASSWORD`

### Secure .env File

```bash
# Right-click .env file → Properties → Security
# Remove all users except your account
# Click "Edit" → Select "Authenticated Users" → Remove → Apply
```

---

## 📊 MONITORING & LOGS

### Check Execution Logs

```bash
# View latest log file
type logs\excel_to_mysql.log

# For PowerShell:
Get-Content logs\excel_to_mysql.log -Tail 50
```

### Expected Log Output

```
2024-01-15 09:00:00 - ExcelToMySQL - INFO - ============================================================================
2024-01-15 09:00:00 - ExcelToMySQL - INFO - EXCEL TO MYSQL AUTOMATION STARTED
2024-01-15 09:00:00 - ExcelToMySQL - INFO - ============================================================================
2024-01-15 09:00:01 - ExcelToMySQL - INFO - Connecting to Outlook... (Attempt 1/3)
2024-01-15 09:00:03 - ExcelToMySQL - INFO - ✓ Successfully connected to Outlook
2024-01-15 09:00:04 - ExcelToMySQL - INFO - Searching for emails from past 1 days...
2024-01-15 09:00:06 - ExcelToMySQL - INFO - ✓ Downloaded: data.xlsx
2024-01-15 09:00:07 - ExcelToMySQL - INFO - Reading Excel file: ./temp_files/20240115_090006_data.xlsx
2024-01-15 09:00:08 - ExcelToMySQL - INFO - ✓ Successfully read 250 rows from Excel file
2024-01-15 09:00:09 - ExcelToMySQL - INFO - Inserting data into table: excel_data
2024-01-15 09:00:12 - ExcelToMySQL - INFO - ✓ Successfully inserted 250 rows (Failed: 0)
2024-01-15 09:00:13 - ExcelToMySQL - INFO - ✓ Cleaned up: ./temp_files/20240115_090006_data.xlsx

EXCEL TO MYSQL AUTOMATION REPORT
======================================================================
Status: ✓ SUCCESS
Duration: 0:00:13.456789
STATISTICS:
- Files Downloaded: 1
- Files Processed: 1
- Rows Inserted: 250
- Errors: 0
======================================================================
```

---

## ❌ TROUBLESHOOTING

### Issue: "Python not found"
```bash
# Verify Python is installed
python --version

# If not found, install Python 3.8+
# During installation: CHECK "Add Python to PATH"
```

### Issue: "Module not found"
```bash
# Reinstall dependencies
pip install -r requirements.txt --force-reinstall
```

### Issue: "Office 365 authentication failed"
```
Solution:
1. Verify you're using App Password (not regular password)
2. Enable 2-factor authentication on Office 365 account
3. Generate new app password
4. Update OUTLOOK_PASSWORD in .env
```

### Issue: "MySQL connection refused"
```bash
# Verify MySQL is running
mysql -h localhost -u root -p -e "SELECT 1"

# Or check MySQL service status (Windows)
wmic service where name="MySQL80" get status
```

### Issue: "Table doesn't exist"
```sql
-- Verify table exists
SHOW TABLES IN your_database_name;

-- If missing, create it:
CREATE TABLE excel_data (
    id INT AUTO_INCREMENT PRIMARY KEY,
    name VARCHAR(255),
    email VARCHAR(255),
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
```

### Issue: "Task Scheduler shows error code"
```
Error Code: 0x1 = Run with highest privileges needed
Error Code: 0xFFFF = Path to script is wrong
Error Code: 0x534 = Task not found

Solution:
1. Right-click task → Properties
2. Check "Run with highest privileges"
3. Verify action path is absolute (C:\path\...) not relative
4. Verify .env file is in same directory as script
```

---

## 🔄 WORKFLOW SUMMARY

```
1. Email arrives with Excel attachment
   ↓
2. Task Scheduler triggers at specified time
   ↓
3. Script connects to Outlook
   ↓
4. Downloads Excel files from past N days
   ↓
5. Reads Excel data
   ↓
6. Connects to MySQL database
   ↓
7. Inserts rows into table
   ↓
8. Cleans up temporary files
   ↓
9. Logs results and sends notification email (optional)
```

---

## 📧 EMAIL NOTIFICATIONS (OPTIONAL)

Enable notifications in .env:
```env
NOTIFY_EMAIL=your-email@gmail.com
NOTIFY_ON_SUCCESS=True
NOTIFY_ON_ERROR=True
```

You'll receive emails like:
- ✓ "Excel to MySQL Sync Successful - 250 rows inserted"
- ✗ "Excel to MySQL Sync Failed - Error details..."

---

## 🎯 NEXT STEPS

1. **Test manually:** Run `python excel_to_mysql_automation.py`
2. **Check logs:** Open `logs/excel_to_mysql.log`
3. **Set up Task Scheduler:** Follow TASK_SCHEDULER_SETUP_GUIDE.md
4. **Monitor:** Check logs regularly, set up email notifications
5. **Optimize:** Adjust DAYS_BACK, TABLE_NAME, schedule as needed

---

## 📞 SUPPORT

**Common Issues:**
- Check `logs/excel_to_mysql.log` for detailed error messages
- Run script manually to test
- Verify all credentials in .env
- Test MySQL connection separately
- Test Outlook credentials separately

**Files to Review:**
- `excel_to_mysql_automation.py` - Main script with full error handling
- `TASK_SCHEDULER_SETUP_GUIDE.md` - Detailed Task Scheduler instructions
- `.env.example` - Configuration reference

---

## 📝 CHANGELOG

**v1.0 (2024-01-15)**
- Initial release
- Outlook/Office 365 integration
- MySQL database support
- Comprehensive error handling
- Email notifications
- Task Scheduler compatible
