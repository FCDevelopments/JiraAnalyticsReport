# setup_schedule.ps1
# Run this once to (re)register the weekly Jira report task in Windows Task Scheduler.
# Does NOT require admin — schtasks.exe registers for the current user by default.
# After setup, the task fires every Monday at 10:00 AM automatically.

$TaskName   = "JiraAnalytics_WeeklyReport"
$ScriptPath = "C:\JiraAnalyticsReport\send_weekly_report.ps1"
$LogPath    = "C:\JiraAnalyticsReport\send_weekly_report.log"

$trArgs = "-NonInteractive -ExecutionPolicy Bypass -File `"$ScriptPath`""

# /f overwrites if the task already exists
schtasks /create /tn $TaskName /tr "powershell.exe $trArgs" /sc WEEKLY /d MON /st 10:00 /f

if ($LASTEXITCODE -eq 0) {
    Write-Host ""
    Write-Host "Task '$TaskName' registered successfully."
    Write-Host "  Schedule : Every Monday at 10:00 AM"
    Write-Host "  Script   : $ScriptPath"
    Write-Host "  Log file : $LogPath"
    Write-Host ""
    Write-Host "To verify : open Task Scheduler and look for '$TaskName' under Task Scheduler Library."
    Write-Host "To test   : schtasks /run /tn `"$TaskName`""
} else {
    Write-Error "Registration failed (exit code $LASTEXITCODE)."
}
