@echo off
cd /d "%~dp0"
echo [1/2] Fetching tickets from Jira (PowerShell)...
powershell -ExecutionPolicy Bypass -File "%~dp0analytics_tickets.ps1"
if errorlevel 1 (
    echo [ERROR] PowerShell step failed. Aborting.
    exit /b 1
)
echo.
echo [2/2] Generating Excel file (Python)...
python analytics_tickets.py
if errorlevel 1 (
    echo [ERROR] Excel generation failed.
    exit /b 1
)
echo.
echo Done!
