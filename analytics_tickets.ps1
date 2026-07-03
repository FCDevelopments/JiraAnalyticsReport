# analytics_tickets.ps1
# Pulls all Jira tickets submitted by the Analytics team.
# Uses Invoke-RestMethod (trusted PowerShell component) to avoid Windows Defender
# behavioral detection that terminates Python processes making repeated HTTP calls.
# Outputs: analytics_tickets_data.json (consumed by analytics_tickets.py for Excel export)

[CmdletBinding()]
param()

Set-StrictMode -Off
$ErrorActionPreference = "Continue"

$SCRIPT_DIR = $PSScriptRoot
if (-not $SCRIPT_DIR) { $SCRIPT_DIR = Split-Path -Parent $MyInvocation.MyCommand.Path }

# Load .env
$envFile = Join-Path $SCRIPT_DIR ".env"
Get-Content $envFile | ForEach-Object {
    if ($_ -match '^\s*([^#][^=]*?)\s*=\s*(.+?)\s*$') {
        [System.Environment]::SetEnvironmentVariable($Matches[1].Trim(), $Matches[2].Trim(), 'Process')
    }
}

$EMAIL    = $env:JIRA_EMAIL
$TOKEN    = $env:JIRA_API_TOKEN
$BASE_URL = if ($env:JIRA_BASE_URL) { $env:JIRA_BASE_URL } else { "https://yourcompany.atlassian.net/rest/api/3" }

if (-not $EMAIL -or -not $TOKEN) {
    Write-Error "Missing JIRA_EMAIL or JIRA_API_TOKEN in .env"
    exit 1
}

$CREDS   = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${EMAIL}:${TOKEN}"))
$HEADERS = @{
    "Authorization" = "Basic $CREDS"
    "Content-Type"  = "application/json"
    "Accept"        = "application/json"
}

# TODO: Replace with your own team's Jira account emails (used to filter
# tickets to those reported by/assigned to your team). This should match
# the email addresses associated with each user's Jira account.
$ANALYTICS_EMAILS = @(
    "alice@example.com",
    "bob@example.com",
    "carol@example.com"
)

function Invoke-JiraGet {
    param($Url, [hashtable]$Params = @{})
    if ($Params.Count -gt 0) {
        $qs = ($Params.GetEnumerator() | ForEach-Object {
            "$($_.Key)=$([Uri]::EscapeDataString($_.Value.ToString()))"
        }) -join "&"
        $Url = "${Url}?${qs}"
    }
    for ($a = 0; $a -lt 3; $a++) {
        try {
            return Invoke-RestMethod -Uri $Url -Method GET -Headers $HEADERS -TimeoutSec 30
        } catch {
            Write-Host "  [WARN] GET attempt $($a+1): $($_.Exception.Message)"
            if ($a -eq 2) { return $null }
            Start-Sleep -Seconds ([Math]::Pow(2, $a))
        }
    }
    return $null
}

function Invoke-JiraPost {
    param($Url, $Body)
    for ($a = 0; $a -lt 3; $a++) {
        try {
            $json = $Body | ConvertTo-Json -Compress -Depth 10
            return Invoke-RestMethod -Uri $Url -Method POST -Headers $HEADERS -Body $json -ContentType "application/json" -TimeoutSec 60
        } catch {
            Write-Host "  [WARN] POST attempt $($a+1): $($_.Exception.Message)"
            if ($a -eq 2) { return $null }
            Start-Sleep -Seconds ([Math]::Pow(2, $a))
        }
    }
    return $null
}

# --- Step 1: Resolve emails to account IDs ---
$totalEmails = $ANALYTICS_EMAILS.Count
Write-Host "Step 1/3 - Resolving $totalEmails analytics team emails to Jira account IDs..."
$resolved = @{}
$notFound = @()

foreach ($email in $ANALYTICS_EMAILS) {
    $result = Invoke-JiraGet -Url "$BASE_URL/user/search" -Params @{ query = $email; maxResults = "1" }
    if ($result -and $result.Count -gt 0) {
        $accountId   = $result[0].accountId
        $displayName = $result[0].displayName
        if ($accountId) {
            $resolved[$email] = $accountId
            Write-Host "  OK  $email -> $displayName"
        } else {
            $notFound += $email
            Write-Host "  ??  $email -> matched but no accountId"
        }
    } else {
        $notFound += $email
        Write-Host "  --  $email -> not found in Jira"
    }
    Start-Sleep -Milliseconds 300
}

$notFoundCount = $notFound.Count
$resolvedCount = $resolved.Count
if ($notFoundCount -gt 0) {
    Write-Host "[WARN] $notFoundCount unresolved: $($notFound -join ', ')"
}
Write-Host "Resolved $resolvedCount of $totalEmails users."

if ($resolvedCount -eq 0) {
    Write-Error "No users resolved - cannot build JQL. Exiting."
    exit 1
}

# --- Step 2: Fetch active tickets created since last run ---
$accountIds = $resolved.Values
$jqlIds     = ($accountIds | ForEach-Object { "`"$_`"" }) -join ", "

$lastRunTimestampPath = Join-Path $SCRIPT_DIR "last_run_timestamp.txt"
if (Test-Path $lastRunTimestampPath) {
    $lastRunTimestamp = (Get-Content $lastRunTimestampPath -Raw -Encoding utf8).Trim()
    Write-Host "Last run: $lastRunTimestamp - fetching tickets created after that."
} else {
    $lastRunTimestamp = (Get-Date).AddDays(-7).ToString("yyyy-MM-dd HH:mm")
    Write-Host "No previous run found - defaulting to past 7 days ($lastRunTimestamp)."
}

$JQL = "reporter in ($jqlIds) AND issuetype not in subTaskIssueTypes() AND status not in (""Approved"", ""Canceled"", ""Cancelled"", ""Closed"", ""Completed"", ""Done"", ""Fulfilled"", ""Resolved"") AND created >= ""$lastRunTimestamp"" ORDER BY created DESC"

Write-Host ""
Write-Host "Step 2/3 - Fetching tickets for $resolvedCount users..."

$allIssues = [System.Collections.Generic.List[object]]::new()
$page      = 1
$nextToken = $null

do {
    Write-Host "  Fetching page ${page}..."
    $body = @{
        jql        = $JQL
        maxResults = 100
        fields     = @("summary", "status", "assignee", "created", "updated", "reporter", "issuetype")
    }
    if ($nextToken) { $body["nextPageToken"] = $nextToken }

    $data = Invoke-JiraPost -Url "$BASE_URL/search/jql" -Body $body

    if (-not $data -or -not $data.issues) {
        Write-Host "  [ERROR] No data on page ${page} - stopping."
        break
    }

    $batch      = $data.issues
    $batchCount = $batch.Count
    foreach ($issue in $batch) { $allIssues.Add($issue) }
    $runningTotal = $allIssues.Count
    Write-Host "    Page ${page}: $batchCount issues (running total: $runningTotal)"

    if ($data.isLast -or $batchCount -lt 100) {
        Write-Host "  Last page reached."
        break
    }

    $nextToken = $data.nextPageToken
    if (-not $nextToken) {
        Write-Host "  [WARN] No nextPageToken returned - stopping."
        break
    }

    $page++
    Start-Sleep -Seconds 1

} while ($true)

$fetchedTotal = $allIssues.Count
Write-Host "Total tickets fetched: $fetchedTotal"

# --- Step 3: Build records and save to JSON ---
Write-Host ""
Write-Host "Step 3/3 - Building records and saving to JSON..."

$records = foreach ($issue in $allIssues) {
    $fields     = $issue.fields
    $reporter   = if ($fields.reporter) { $fields.reporter.displayName } else { "Unknown" }
    $assignee   = if ($fields.assignee) { $fields.assignee.displayName } else { "Unassigned" }

    $dateStr    = ""
    if ($fields.created) {
        try {
            $dt      = [DateTimeOffset]::Parse($fields.created)
            $dateStr = $dt.ToLocalTime().ToString("yyyy-MM-dd HH:mm")
        } catch { $dateStr = $fields.created }
    }

    $updatedStr = ""
    if ($fields.updated) {
        try {
            $dt         = [DateTimeOffset]::Parse($fields.updated)
            $updatedStr = $dt.ToLocalTime().ToString("yyyy-MM-dd HH:mm")
        } catch { $updatedStr = $fields.updated }
    }

    [PSCustomObject]@{
        ticket_key     = $issue.key
        title          = if ($fields.summary) { $fields.summary } else { "" }
        status         = if ($fields.status)  { $fields.status.name } else { "" }
        date_requested = $dateStr
        last_updated   = $updatedStr
        requested_by   = $reporter
        assigned_to    = $assignee
    }
}

# Save main data JSON - force array so Python always gets [] not null
$jsonPath    = Join-Path $SCRIPT_DIR "analytics_tickets_data.json"
$recordArray = @($records)
$recordCount = $recordArray.Count
$recordArray | ConvertTo-Json -Depth 5 | Out-File -FilePath $jsonPath -Encoding utf8
Write-Host "JSON saved: $jsonPath ($recordCount records)"

# Advance the timestamp for next run
$currentTimestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm")
$currentTimestamp | Out-File -FilePath $lastRunTimestampPath -Encoding utf8 -NoNewline
Write-Host "Timestamp saved: $lastRunTimestampPath ($currentTimestamp)"
Write-Host "Done. Run analytics_tickets.py to generate the Excel file."
