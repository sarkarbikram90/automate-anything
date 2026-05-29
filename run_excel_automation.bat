@echo off
REM ========================================================================
REM Excel to MySQL Automation Batch File
REM This file is used by Windows Task Scheduler to run the Python script
REM ========================================================================

REM Get current date and time for logging
for /f "tokens=2-4 delims=/ " %%a in ('date /t') do (set mydate=%%c%%a%%b)
for /f "tokens=1-2 delims=/:" %%a in ('time /t') do (set mytime=%%a%%b)

REM Set script directory (CHANGE THIS TO YOUR ACTUAL DIRECTORY)
set SCRIPT_DIR=C:\Path\To\Your\Script\Directory
set PYTHON_SCRIPT=%SCRIPT_DIR%\excel_to_mysql_automation.py

REM Set log file
set LOG_FILE=%SCRIPT_DIR%\task_runs.log

REM ========================================================================
REM MAIN EXECUTION
REM ========================================================================

REM Echo to log file (append mode with >>)
echo. >> "%LOG_FILE%"
echo ========================================================================>> "%LOG_FILE%"
echo Task Run: %mydate% at %mytime%>> "%LOG_FILE%"
echo ========================================================================>> "%LOG_FILE%"

REM Change to script directory
cd /d "%SCRIPT_DIR%"
if %errorlevel% neq 0 (
    echo ERROR: Failed to change directory to %SCRIPT_DIR% >> "%LOG_FILE%"
    exit /b 1
)

REM Run Python script and log output
python "%PYTHON_SCRIPT%" >> "%LOG_FILE%" 2>&1

if %errorlevel% neq 0 (
    echo ERROR: Python script failed with exit code %errorlevel% >> "%LOG_FILE%"
    exit /b %errorlevel%
) else (
    echo SUCCESS: Script completed successfully >> "%LOG_FILE%"
    exit /b 0
)

REM ========================================================================
REM NOTES:
REM - Uncomment the following line to see console output (for testing)
REM - Remove for production to run silently
REM - @pause
REM ========================================================================
