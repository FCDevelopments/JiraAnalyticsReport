# send_weekly_report.ps1
# Full pipeline: fetch Jira data -> generate Excel -> email via Outlook COM
# Designed to run via Windows Task Scheduler every Monday at 10:00 AM.
# Outlook must be installed and the user's account signed in (same approach as AmazonReportBot).

[CmdletBinding()]
param(
    # TODO: set this to your own Analytics team distribution list address
    [string]$To      = "analytics-team@example.com",
    [string]$Cc      = "",
    [string]$Subject = ""
)

$ErrorActionPreference = "Stop"
$SCRIPT_DIR = $PSScriptRoot
if (-not $SCRIPT_DIR) { $SCRIPT_DIR = Split-Path -Parent $MyInvocation.MyCommand.Path }

$LogFile = Join-Path $SCRIPT_DIR "send_weekly_report.log"
function Log($msg) {
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $msg"
    Write-Host $line
    Add-Content -Path $LogFile -Value $line -Encoding utf8
}

# ── Step 1: Fetch Jira data ────────────────────────────────────────────────────
Log "Step 1/3 - Fetching Jira tickets..."
& powershell.exe -NonInteractive -ExecutionPolicy Bypass -File (Join-Path $SCRIPT_DIR "analytics_tickets.ps1")
if ($LASTEXITCODE -ne 0) {
    Log "ERROR: analytics_tickets.ps1 exited with code $LASTEXITCODE"
    exit 1
}
Log "Jira fetch complete."

# ── Step 2: Generate Excel ─────────────────────────────────────────────────────
Log "Step 2/3 - Generating Excel report..."
$pyOut = & python (Join-Path $SCRIPT_DIR "analytics_tickets.py") 2>&1
$pyOut | ForEach-Object { Log "  [py] $_" }
if ($LASTEXITCODE -ne 0) {
    Log "ERROR: analytics_tickets.py exited with code $LASTEXITCODE"
    exit 1
}
Log "Excel generation complete."

# ── Step 3: Find the Excel generated in this run (within last 2 min) ──────────
$cutoff     = (Get-Date).AddMinutes(-2)
$latestXlsx = Get-ChildItem -Path $SCRIPT_DIR -Filter "analytics_tickets_*.xlsx" |
              Where-Object { $_.LastWriteTime -ge $cutoff } |
              Sort-Object LastWriteTime -Descending |
              Select-Object -First 1

$weekOf  = (Get-Date).ToString("MMMM d, yyyy")
if (-not $Subject) { $Subject = "Analytics Team - Open Jira Tickets ($weekOf)" }

# ── Step 4: Send via Outlook COM ──────────────────────────────────────────────
Log "Step 3/3 - Sending email to $To via Outlook..."

if (-not $latestXlsx) {
    Log "No active tickets found - sending notification with no attachment."
    $body = @"
Hi Team,

The Analytics team currently has no open or in-progress Jira tickets as of $(Get-Date -Format 'dddd, MMMM d, yyyy').

This report is sent automatically every Monday at 10:00 AM.

-- IT Team
"@
} else {
    Log "Attaching: $($latestXlsx.Name)"
    $body = @"
Hi Team,

Please find attached this week's open Jira tickets for the Analytics team.

This report includes all tickets currently open, in progress, or pending - excluding anything already completed, fulfilled, or cancelled.

Report generated: $(Get-Date -Format 'dddd, MMMM d, yyyy') at $(Get-Date -Format 'h:mm tt')
File: $($latestXlsx.Name)

This report is sent automatically every Monday at 10:00 AM.

-- IT Team
"@
}

try {
    $outlook     = New-Object -ComObject Outlook.Application
    $session     = $outlook.Session
    $mail        = $outlook.CreateItem(0)   # 0 = olMailItem
    $store       = $session.DefaultStore
    $sentFolder  = $null
    if ($store) { $sentFolder = $store.GetDefaultFolder(5) }
    if ($sentFolder) { $mail.SaveSentMessageFolder = $sentFolder }

    # Pick the account tied to the default store
    $account = $null
    foreach ($acc in $session.Accounts) {
        if ($null -ne $acc.DeliveryStore -and $null -ne $store -and
            $acc.DeliveryStore.StoreID -eq $store.StoreID) {
            $account = $acc; break
        }
    }
    if (-not $account -and $session.Accounts.Count -gt 0) {
        $account = $session.Accounts.Item(1)
    }
    if ($account) { $mail.SendUsingAccount = $account }

    $mail.To      = $To
    if ($Cc) { $mail.CC = $Cc }
    $mail.Subject = $Subject
    $mail.Body    = $body
    if ($latestXlsx) {
        [void]$mail.Attachments.Add((Resolve-Path $latestXlsx.FullName).Path)
    }
    $mail.Send()

    $usedAccount = if ($account) { "$($account.DisplayName) <$($account.SmtpAddress)>" } else { "<none>" }
    Log "Email sent: To='$To' | Subject='$Subject' | Account='$usedAccount'"
} catch {
    Log "ERROR sending via Outlook: $($_.Exception.Message)"
    exit 1
}

Log "Pipeline complete."
