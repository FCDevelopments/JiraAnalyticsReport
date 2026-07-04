# JiraAnalyticsReport

Lightweight weekly status report: fetch a team's open Jira tickets, render a styled Excel workbook, and email it to a distribution list every Monday morning — three cleanly separated steps.

```
analytics_tickets.ps1  →  analytics_tickets_data.json   (fetch — PowerShell, Jira REST API)
analytics_tickets.py   →  analytics_tickets_YYYYMMDD.xlsx (format — Python, openpyxl)
send_weekly_report.ps1 →  Outlook email                  (send — Outlook COM)
```

The Excel step deliberately makes **no network calls** — it reads pre-fetched JSON, which keeps fetch/format/send independently testable and replaceable.

## Design notes

- **PowerShell for HTTP** — `Invoke-RestMethod` is a trusted Windows component, avoiding endpoint-protection rules that flag script runtimes making repeated outbound HTTP calls
- **Incremental pulls** — `last_run_timestamp.txt` scopes each run to tickets created since the previous one
- **Account-ID resolution** — team member emails are resolved to Jira account IDs at runtime, so the JQL survives display-name changes
- **No-ticket weeks still notify** — the email goes out with a "nothing open" note instead of silently skipping

## Setup

```powershell
pip install -r requirements.txt
copy .env.example .env        # Jira URL, email, API token
```

1. Edit `analytics_tickets.ps1` → replace the `$ANALYTICS_EMAILS` placeholder list with your team's Jira account emails.
2. Edit `send_weekly_report.ps1` → set the `-To` distribution list address.
3. Register the Monday 10:00 AM task (no admin needed):

```powershell
.\setup_schedule.ps1
```

Or run everything now: `run.bat` (fetch + Excel) or `send_weekly_report.ps1` (full pipeline including email).

## Requirements

Windows with classic Outlook installed and signed in (COM automation sends from your default account — no SMTP credentials stored).

## Stack

PowerShell · Python · openpyxl · Jira REST API v3 · Outlook COM · Task Scheduler
