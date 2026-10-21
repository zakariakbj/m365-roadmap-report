# M365 Roadmap Report

A PowerShell script that emails a weekly digest of Microsoft 365 changes for the workloads you choose: Exchange, Outlook and Security (Defender for Office 365, Purview) by default, and SharePoint, Teams, OneDrive, Copilot or any other workload by editing one file.

It collects what changed in the last 7 days from two sources:

- **Message Center** (Microsoft Graph): admin announcements, planned changes and required actions.
- **Microsoft 365 Roadmap** (public API): new and updated features, with their status.

Each item shows its workload and links to the Admin Center or the roadmap. If a source is unavailable, the report is still sent with a warning.

## Requirements

- PowerShell 7+ (Windows, Linux, macOS) or Windows PowerShell 5.1. No module to install.
- An Entra ID app registration with the Microsoft Graph application permissions `ServiceMessage.Read.All` and `Mail.Send`.
- An Exchange Online mailbox to send the report from.

[docs/azure-app-setup.md](docs/azure-app-setup.md) walks through the app registration and shows how to restrict `Mail.Send` to the sender mailbox.

## Configuration

Copy `env.ini.example` to `env.ini`, next to the script, and fill it in:

```ini
AZURE_TENANT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
AZURE_CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
AZURE_CLIENT_SECRET=your_client_secret_here
SENDER_EMAIL=it-reports@contoso.com
RECIPIENT_EMAIL=team-it@contoso.com
```

Environment variables with the same names take precedence over the file, so the secret can also come from your scheduler or secret store. `-ConfigPath` reads another file.

`env.ini` is git-ignored. Restrict its permissions, since it holds the client secret (`chmod 600 env.ini` on Linux).

## Choosing the workloads

The workloads are listed in [workloads.psd1](workloads.psd1). Each workload has a name, shown as a badge on its items, and the names to look for in each source:

```powershell
@{
    Name          = "SharePoint & OneDrive"
    MessageCenter = @("SharePoint Online", "Microsoft OneDrive")   # Message Center services
    Roadmap       = @("SharePoint", "OneDrive")                    # Roadmap products
}
```

The two sources do not use the same names. List them, with their number of items:

```powershell
pwsh -File ./Send-M365RoadmapReport.ps1 -ListWorkloads
```

Roadmap products come from the public roadmap. Message Center services come from your tenant and are listed only when the Entra ID settings are configured.

Matching ignores case, and `*` is a wildcard (`"*Purview*"`). A workload can list names for one source only. An item that matches several workloads shows all their badges. `Title` sets the report header and the email subject.

To keep your changes out of the way of updates, copy the file elsewhere and pass it with `-WorkloadsPath`.

## Usage

```powershell
# Send the report
pwsh -File ./Send-M365RoadmapReport.ps1

# Preview: write the HTML report to a file instead of sending it
pwsh -File ./Send-M365RoadmapReport.ps1 -OutFile report.html

# Look back 14 days instead of 7
pwsh -File ./Send-M365RoadmapReport.ps1 -Days 14

# Use another workloads file
pwsh -File ./Send-M365RoadmapReport.ps1 -WorkloadsPath ./my-workloads.psd1

# Windows PowerShell 5.1
powershell -ExecutionPolicy Bypass -File .\Send-M365RoadmapReport.ps1
```

The script exits with code 1 when it fails, so schedulers and monitoring can detect it.

## Scheduling

### Linux (cron), every Monday at 08:00

```cron
TZ=Europe/Paris
0 8 * * 1 /usr/bin/pwsh -NoProfile -File /opt/m365-roadmap-report/Send-M365RoadmapReport.ps1
```

The report dates use the time zone of the process (`TZ` here).

### Windows (Task Scheduler), every Monday at 08:00

```powershell
$action  = New-ScheduledTaskAction -Execute "pwsh.exe" -Argument "-NoProfile -File C:\m365-roadmap-report\Send-M365RoadmapReport.ps1"
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At "08:00"
Register-ScheduledTask -TaskName "M365 Roadmap Report" -Action $action -Trigger $trigger
```

With Windows PowerShell 5.1, use `-Execute "powershell.exe"` and add `-ExecutionPolicy Bypass` to the arguments.

## Logs

Each run appends to `logs/m365_roadmap_report.log` next to the script. The folder is created on the first run and git-ignored.

A weekly run writes a few lines, so rotation is optional. On Linux, logrotate can handle it:

```
/opt/m365-roadmap-report/logs/m365_roadmap_report.log {
    monthly
    rotate 12
    compress
    missingok
    notifempty
}
```

## Updating

Releases are listed on the [releases page](https://github.com/zakariakbj/m365-roadmap-report/releases), with their changes in [CHANGELOG.md](CHANGELOG.md). The version that sent a report is shown in its footer and in the log. A major version (`2.0.0`) means your configuration needs a change: read its changelog entry before updating.

## Development

Tests use [Pester](https://pester.dev) 5+, and the code is checked with [PSScriptAnalyzer](https://github.com/PowerShell/PSScriptAnalyzer):

```powershell
Install-Module Pester, PSScriptAnalyzer -Scope CurrentUser
Invoke-Pester ./tests
Invoke-ScriptAnalyzer -Path . -Recurse -Settings ./PSScriptAnalyzerSettings.psd1
```

The CI runs both on Linux and Windows with PowerShell 7, and on Windows PowerShell 5.1.

## License

[MIT](LICENSE)
