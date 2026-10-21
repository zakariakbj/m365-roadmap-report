<#
.SYNOPSIS
    Weekly email report of Microsoft 365 changes for the workloads you choose.

.DESCRIPTION
    Fetches Microsoft 365 changes for the workloads listed in workloads.psd1
    (Exchange, Outlook and Security by default) from two sources:
      - Message Center (Microsoft Graph API): admin announcements created or
        modified in the last N days, matched on their services
      - Public Microsoft 365 Roadmap: feature items modified in the last N days,
        matched on their product tags
    Renders a consolidated HTML report and sends it by email through the Graph API.

    Settings come from environment variables, or from an INI file
    (KEY=value lines, env.ini next to the script by default). Environment
    variables take precedence over the file:
      AZURE_TENANT_ID      Entra ID tenant ID
      AZURE_CLIENT_ID      App registration client ID
      AZURE_CLIENT_SECRET  App registration client secret
      SENDER_EMAIL         Exchange Online mailbox used to send the report
      RECIPIENT_EMAIL      Report recipient address

.PARAMETER ConfigPath
    INI file to read the settings from. Defaults to env.ini next to the script.

.PARAMETER WorkloadsPath
    PowerShell data file listing the workloads to report on. Defaults to
    workloads.psd1 next to the script.

.PARAMETER Days
    Number of days to look back. Defaults to 7.

.PARAMETER OutFile
    Writes the HTML report to this file instead of sending it by email.
    SENDER_EMAIL and RECIPIENT_EMAIL are not needed in this mode.

.PARAMETER ListWorkloads
    Lists the Roadmap products and Message Center services available for
    workloads.psd1, with their number of items, instead of sending the report.
    Message Center services are listed only when the Entra ID settings are set.

.EXAMPLE
    pwsh -File Send-M365RoadmapReport.ps1

.EXAMPLE
    pwsh -File Send-M365RoadmapReport.ps1 -OutFile report.html

.EXAMPLE
    pwsh -File Send-M365RoadmapReport.ps1 -ListWorkloads

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File Send-M365RoadmapReport.ps1

.NOTES
    Version : 1.0.0
    Author  : Zakaria KABBAJ
    Project : https://github.com/zakariakbj/m365-roadmap-report
    Requires: PowerShell 7+ or Windows PowerShell 5.1, no external module
    License : MIT
#>
[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "env.ini"),
    [string]$WorkloadsPath = (Join-Path $PSScriptRoot "workloads.psd1"),
    [ValidateRange(1, 90)]
    [int]$Days = 7,
    [string]$OutFile,
    [switch]$ListWorkloads
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
$GRAPH_MESSAGES_URL = "https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/messages"
$ROADMAP_URL        = "https://www.microsoft.com/releasecommunications/api/v1/m365"
$SEVERITIES         = @("normal", "high", "critical")
$ROADMAP_STATUSES   = @{ "in development" = "dev"; "rolling out" = "rolling"; "launched" = "launched"; "cancelled" = "cancelled" }
$EM_DASH            = [char]0x2014
$SCRIPT_VERSION     = "1.0.0"  # also in .NOTES and CHANGELOG.md
$PROJECT_URL        = "https://github.com/zakariakbj/m365-roadmap-report"

$LogFile = Join-Path (Join-Path $PSScriptRoot "logs") "m365_roadmap_report.log"

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
function Write-ReportLog {
    param([string]$Level, [string]$Message)
    $line = "{0} [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level.ToUpper(), $Message
    $dir = Split-Path -Parent $LogFile
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
    Add-Content -Path $LogFile -Value $line
    Write-Information $line -InformationAction Continue
}

# ---------------------------------------------------------------------------
# Load the INI file; environment variables already set win
# ---------------------------------------------------------------------------
function Import-Config {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return }
    foreach ($line in Get-Content $Path) {
        if ($line -match '^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*?)\s*$') {
            $name  = $Matches[1]
            $value = $Matches[2] -replace '^"(.*)"$', '$1' -replace "^'(.*)'$", '$1'
            if (-not [System.Environment]::GetEnvironmentVariable($name)) {
                [System.Environment]::SetEnvironmentVariable($name, $value)
            }
        }
    }
}

# ---------------------------------------------------------------------------
# Workloads: which Message Center services and Roadmap products to report on
# ---------------------------------------------------------------------------
function Import-Workload {
    param([string]$Path)
    if (-not (Test-Path $Path)) { throw "Workloads file not found: $Path" }
    $data = Import-PowerShellDataFile -Path $Path

    $workloads = @()
    foreach ($w in @(Get-Prop $data "Workloads" @())) {
        $name          = [string](Get-Prop $w "Name")
        $messageCenter = @(Get-Prop $w "MessageCenter" @() | Where-Object { $_ })
        $roadmap       = @(Get-Prop $w "Roadmap" @() | Where-Object { $_ })
        if (-not $name) { throw "${Path}: every workload needs a Name" }
        if ($messageCenter.Count -eq 0 -and $roadmap.Count -eq 0) { throw "${Path}: workload '$name' needs MessageCenter or Roadmap names" }
        $workloads += [PSCustomObject]@{ Name = $name; MessageCenter = $messageCenter; Roadmap = $roadmap }
    }
    if ($workloads.Count -eq 0) { throw "${Path}: no workload defined" }

    $title = [string](Get-Prop $data "Title")
    if (-not $title) { $title = ($workloads.Name) -join ", " }
    return [PSCustomObject]@{ Title = $title; Workloads = $workloads }
}

# Case-insensitive; patterns with * or ? are wildcards, others must match exactly
function Test-NameMatch {
    param([string[]]$Names, [string[]]$Patterns)
    foreach ($n in $Names) {
        $n = "$n".Trim()
        foreach ($p in $Patterns) {
            if ($p -match '[*?]') { if ($n -like $p) { return $true } }
            elseif ($n -eq $p.Trim()) { return $true }
        }
    }
    return $false
}

function Get-MatchingWorkload {
    param([string[]]$Names, $Workloads, [ValidateSet("MessageCenter", "Roadmap")][string]$Source)
    return @(foreach ($w in @($Workloads)) { if (Test-NameMatch $Names $w.$Source) { $w.Name } })
}

# Roadmap product tags; items without tagsContainer fall back to their flat tags
function Get-RoadmapProduct {
    param($Item)
    $container = Get-Prop $Item "tagsContainer" $null
    $tags = if ($container) { @(Get-Prop $container "products" @()) } else { @(Get-Prop $Item "tags" @()) }
    return @(foreach ($t in $tags) {
        if ($t -is [string]) { $t }
        elseif ((Get-Prop $t "tagName" $null) -is [string]) { $t.tagName }
    })
}

# ---------------------------------------------------------------------------
# OAuth2 token (client credentials)
# ---------------------------------------------------------------------------
function Get-AccessToken {
    param([string]$TenantId, [string]$ClientId, [string]$ClientSecret)

    $url  = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $body = @{
        grant_type    = "client_credentials"
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = "https://graph.microsoft.com/.default"
    }
    $resp = Invoke-RestMethod -Method Post -Uri $url -Body $body -ContentType "application/x-www-form-urlencoded"
    if (-not $resp.PSObject.Properties["access_token"]) { throw "Authentication failed: no access token returned" }
    return $resp.access_token
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
# Safe property or hashtable key accessor (compatible with Set-StrictMode -Version Latest)
function Get-Prop {
    param($Object, [string]$Name, $Default = "")
    if ($null -eq $Object) { return $Default }
    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name) -and $null -ne $Object[$Name]) { return $Object[$Name] }
        return $Default
    }
    if ($Object.PSObject.Properties[$Name] -and $null -ne $Object.$Name) { return $Object.$Name }
    return $Default
}

function ConvertTo-PlainText {
    param([string]$Text)
    $Text = [regex]::Replace($Text, '<[^>]+>', ' ')
    $Text = [System.Net.WebUtility]::HtmlDecode($Text)
    $Text = [regex]::Replace($Text, '\s+', ' ')
    return $Text.Trim()
}

function ConvertTo-HtmlText {
    param([string]$Text)
    return [System.Net.WebUtility]::HtmlEncode($Text)
}

# Cuts at a word boundary, with a 5-character leeway (same as Jinja's truncate)
function Get-Excerpt {
    param([string]$Text, [int]$Length = 300)
    if ($Text.Length -le $Length + 5) { return $Text }
    $cut   = $Text.Substring(0, $Length - 3)
    $space = $cut.LastIndexOf(" ")
    if ($space -gt 0) { $cut = $cut.Substring(0, $space) }
    return $cut + "..."
}

function Get-DatePart {
    param([string]$Text)
    return $Text.Substring(0, [Math]::Min(10, $Text.Length))
}

function Test-InWindow {
    param([string]$Date, [datetime]$Since)
    if (-not $Date) { return $false }
    try {
        $dt = [datetime]::Parse($Date, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind)
    } catch {
        return $false
    }
    if ($dt.Kind -eq [System.DateTimeKind]::Unspecified) { $dt = [datetime]::SpecifyKind($dt, [System.DateTimeKind]::Utc) }
    return $dt.ToUniversalTime() -ge $Since
}

# ---------------------------------------------------------------------------
# Message Center (Microsoft Graph)
# ---------------------------------------------------------------------------
function Get-MessageCenterData {
    param([string]$Token)
    $headers = @{ Authorization = "Bearer $Token" }
    $all     = @()
    $url     = $GRAPH_MESSAGES_URL
    while ($url) {
        $resp = Invoke-RestMethod -Method Get -Uri $url -Headers $headers
        $all += @(Get-Prop $resp "value" @())
        $url = Get-Prop $resp "@odata.nextLink" $null
    }
    return $all
}

function Get-MessageCenterMessage {
    param([string]$Token, $Workloads, [int]$Days = 7)

    $since    = (Get-Date).ToUniversalTime().AddDays(-$Days)
    $filtered = @()
    foreach ($msg in @(Get-MessageCenterData -Token $Token)) {
        $matched = @(Get-MatchingWorkload -Names @(Get-Prop $msg "services" @()) -Workloads $Workloads -Source MessageCenter)
        if ($matched.Count -eq 0) { continue }

        # Messages without a modification date are kept, as in the Message Center
        $lastModified = [string](Get-Prop $msg "lastModifiedDateTime")
        if ($lastModified -and -not (Test-InWindow $lastModified $since)) { continue }

        $msgId = [string](Get-Prop $msg "id")
        $filtered += [PSCustomObject]@{
            id                 = $msgId
            title              = [string](Get-Prop $msg "title")
            message_type       = [string](Get-Prop $msg "messageType")
            severity           = [string](Get-Prop $msg "severity" "normal")
            workloads          = $matched
            start_date         = [string](Get-Prop $msg "startDateTime")
            action_required_by = [string](Get-Prop $msg "actionRequiredByDateTime")
            description        = ConvertTo-PlainText ([string](Get-Prop (Get-Prop $msg "body" $null) "content"))
            web_link           = if ($msgId) { "https://admin.microsoft.com/Adminportal/Home#/MessageCenter/:/messages/$msgId" } else { "" }
        }
    }
    return $filtered
}

# ---------------------------------------------------------------------------
# Microsoft 365 Roadmap (public API)
# ---------------------------------------------------------------------------
function Get-RoadmapData {
    $resp = Invoke-RestMethod -Method Get -Uri $ROADMAP_URL
    if ($resp -is [array]) { return $resp }
    if (Get-Prop $resp "items" $null) { return $resp.items }
    if (Get-Prop $resp "value" $null) { return $resp.value }
    return @()
}

function Get-RoadmapItem {
    param($Workloads, [int]$Days = 7)

    $since    = (Get-Date).ToUniversalTime().AddDays(-$Days)
    $filtered = @()
    foreach ($item in @(Get-RoadmapData)) {
        $matched = @(Get-MatchingWorkload -Names (Get-RoadmapProduct $item) -Workloads $Workloads -Source Roadmap)
        if ($matched.Count -eq 0) { continue }

        $modified = [string](Get-Prop $item "modified")
        if (-not (Test-InWindow $modified $since)) { continue }

        $itemId   = [string](Get-Prop $item "id")
        $moreInfo = [string](Get-Prop $item "moreInfoLink")
        if (-not $moreInfo -and $itemId) { $moreInfo = "https://www.microsoft.com/en-us/microsoft-365/roadmap?id=$itemId" }

        $filtered += [PSCustomObject]@{
            id             = $itemId
            title          = [string](Get-Prop $item "title")
            description    = ConvertTo-PlainText ([string](Get-Prop $item "description"))
            status         = [string](Get-Prop $item "status")
            workloads      = $matched
            available_date = [string](Get-Prop $item "publicDisclosureAvailabilityDate")
            modified       = $modified
            more_info_link = $moreInfo
        }
    }
    return $filtered
}

# ---------------------------------------------------------------------------
# Names available for workloads.psd1, with their number of items
# ---------------------------------------------------------------------------
function Get-AvailableWorkload {
    param([string]$Token)
    $roadmap = @(foreach ($item in @(Get-RoadmapData)) { Get-RoadmapProduct $item })
    $roadmap | Group-Object | Sort-Object Name | ForEach-Object {
        [PSCustomObject]@{ Source = "Roadmap"; Name = $_.Name; Items = $_.Count }
    }
    if ($Token) {
        $services = @(foreach ($msg in @(Get-MessageCenterData -Token $Token)) { @(Get-Prop $msg "services" @()) | ForEach-Object { "$_".Trim() } })
        $services | Group-Object | Sort-Object Name | ForEach-Object {
            [PSCustomObject]@{ Source = "MessageCenter"; Name = $_.Name; Items = $_.Count }
        }
    }
}

# ---------------------------------------------------------------------------
# HTML report
# ---------------------------------------------------------------------------
function Get-WorkloadBadge {
    param($Item)
    return (@(Get-Prop $Item "workloads" @()) | ForEach-Object { "<span class=`"badge workload`">$(ConvertTo-HtmlText $_.ToUpper())</span>" }) -join " "
}

function Get-StatusBadge {
    param([string]$Status)
    if (-not $Status) { return "" }
    $class = $ROADMAP_STATUSES[$Status.ToLower()]
    if (-not $class) { $class = "dev" }
    return "<span class=`"badge $class`">$(ConvertTo-HtmlText $Status.ToUpper())</span>"
}

function ConvertTo-ReportHtml {
    param($Messages, $RoadmapItems, $Warnings, [string]$Title = "Microsoft 365", [datetime]$Now = (Get-Date))

    $weekStart   = $Now.Date.AddDays(-(([int]$Now.DayOfWeek + 6) % 7))  # Monday
    $weekEnd     = $weekStart.AddDays(6)
    $generatedAt = "{0} {1}" -f $Now.ToString("MM/dd/yyyy HH:mm"), [System.TimeZoneInfo]::Local.Id

    $warningsHtml = ""
    foreach ($w in @($Warnings)) {
        $warningsHtml += "  <div class=`"warning`">&#9888; $(ConvertTo-HtmlText $w)</div>`n"
    }

    $messagesHtml = ""
    foreach ($msg in @($Messages)) {
        $sev           = $msg.severity.ToLower()
        if ($SEVERITIES -notcontains $sev) { $sev = "normal" }
        $workloadBadge = Get-WorkloadBadge $msg
        $startSpan     = if ($msg.start_date) { "<span>Start: $(ConvertTo-HtmlText (Get-DatePart $msg.start_date))</span>" } else { "" }
        $actionSpan    = if ($msg.action_required_by) { "<span>&#9888; Action required by: $(ConvertTo-HtmlText (Get-DatePart $msg.action_required_by))</span>" } else { "" }
        $linkHtml      = if (Get-Prop $msg "web_link") { "<div class=`"card-link`"><a href=`"$(ConvertTo-HtmlText $msg.web_link)`">View in Admin Center &#8594;</a></div>" } else { "" }

        $messagesHtml += @"
      <div class="card $sev">
        <div class="card-title">$(ConvertTo-HtmlText $msg.id) &mdash; $(ConvertTo-HtmlText $msg.title)</div>
        <div class="card-meta">
          <span><span class="badge $sev">$($sev.ToUpper())</span></span>
          $workloadBadge
          <span>Type: $(ConvertTo-HtmlText $msg.message_type)</span>
          $startSpan
          $actionSpan
        </div>
        <div class="card-desc">$(ConvertTo-HtmlText (Get-Excerpt $msg.description))</div>
        $linkHtml
      </div>

"@
    }
    if (-not $messagesHtml) { $messagesHtml = "      <p class=`"empty`">No Message Center events this week.</p>`n" }

    $roadmapHtml = ""
    foreach ($item in @($RoadmapItems)) {
        $statusBadge   = Get-StatusBadge $item.status
        $workloadBadge = Get-WorkloadBadge $item
        $availSpan     = if ($item.available_date) { "<span>Available: $(ConvertTo-HtmlText $item.available_date)</span>" } else { "" }
        $modSpan       = if (Get-Prop $item "modified") { "<span>Updated: $(ConvertTo-HtmlText (Get-DatePart $item.modified))</span>" } else { "" }
        $linkHtml      = if ($item.more_info_link) { "<div class=`"card-link`"><a href=`"$(ConvertTo-HtmlText $item.more_info_link)`">Learn more &#8594;</a></div>" } else { "" }

        $roadmapHtml += @"
      <div class="card roadmap">
        <div class="card-title">$(ConvertTo-HtmlText $item.title)</div>
        <div class="card-meta">
          $statusBadge
          $workloadBadge
          $availSpan
          $modSpan
        </div>
        <div class="card-desc">$(ConvertTo-HtmlText (Get-Excerpt $item.description))</div>
        $linkHtml
      </div>

"@
    }
    if (-not $roadmapHtml) { $roadmapHtml = "      <p class=`"empty`">No roadmap items this week.</p>`n" }

    return @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<style>
  body { font-family: Segoe UI, Arial, sans-serif; background: #f4f4f4; margin: 0; padding: 20px; }
  .container { max-width: 800px; margin: auto; background: white; border-radius: 8px; overflow: hidden; }
  .header { background: #0078d4; color: white; padding: 24px 32px; }
  .header h1 { margin: 0; font-size: 20px; }
  .header p { margin: 4px 0 0; font-size: 13px; opacity: 0.85; }
  .section { padding: 24px 32px; border-bottom: 1px solid #e0e0e0; }
  .section h2 { font-size: 15px; text-transform: uppercase; letter-spacing: 1px; color: #555; margin: 0 0 16px; }
  .card { border-left: 4px solid #ccc; padding: 12px 16px; margin-bottom: 12px; background: #fafafa; border-radius: 4px; }
  .card.critical { border-left-color: #d13438; background: #fff4f4; }
  .card.high { border-left-color: #ca5010; background: #fff9f5; }
  .card.normal { border-left-color: #0078d4; background: #f5f9ff; }
  .card.roadmap { border-left-color: #107c10; background: #f4fff4; }
  .card-title { font-weight: 600; font-size: 14px; margin-bottom: 6px; }
  .card-meta { font-size: 12px; color: #666; margin-bottom: 6px; }
  .card-meta span { margin-right: 16px; }
  .badge { display: inline-block; padding: 2px 8px; border-radius: 12px; font-size: 11px; font-weight: 600; margin-right: 6px; }
  .badge.critical { background: #d13438; color: white; }
  .badge.high { background: #ca5010; color: white; }
  .badge.normal { background: #0078d4; color: white; }
  .badge.rolling { background: #107c10; color: white; }
  .badge.dev { background: #8764b8; color: white; }
  .badge.launched { background: #038387; color: white; }
  .badge.cancelled { background: #8a8886; color: white; }
  .badge.workload { background: #004b6b; color: white; }
  .card-desc { font-size: 13px; color: #333; line-height: 1.5; }
  .card-link { font-size: 12px; margin-top: 6px; }
  .card-link a { color: #0078d4; }
  .empty { color: #999; font-size: 13px; font-style: italic; }
  .warning { background: #fff4ce; border: 1px solid #f7c948; padding: 12px 16px; border-radius: 4px; margin: 16px 32px 0; font-size: 13px; color: #7a5200; }
  .footer { padding: 16px 32px; font-size: 11px; color: #999; text-align: center; }
  .footer a { color: #999; }
</style>
</head>
<body>
<div class="container">
  <div class="header">
    <h1>M365 Weekly Report &mdash; $(ConvertTo-HtmlText $Title)</h1>
    <p>Week of $($weekStart.ToString("MM/dd/yyyy")) to $($weekEnd.ToString("MM/dd/yyyy"))</p>
  </div>
$warningsHtml
  <div class="section">
    <h2>Message Center</h2>
$messagesHtml  </div>

  <div class="section">
    <h2>Roadmap</h2>
$roadmapHtml  </div>

  <div class="footer">Generated by <a href="$PROJECT_URL">m365-roadmap-report</a> $SCRIPT_VERSION &mdash; $(ConvertTo-HtmlText $generatedAt)</div>
</div>
</body>
</html>
"@
}

# ---------------------------------------------------------------------------
# Send the report through the Graph API
# ---------------------------------------------------------------------------
function Send-Report {
    param([string]$Token, [string]$From, [string]$To, [string]$Subject, [string]$HtmlBody)

    $url     = "https://graph.microsoft.com/v1.0/users/$From/sendMail"
    $headers = @{ Authorization = "Bearer $Token" }
    $payload = @{
        message = @{
            subject      = $Subject
            body         = @{ contentType = "HTML"; content = $HtmlBody }
            toRecipients = @(@{ emailAddress = @{ address = $To } })
        }
        saveToSentItems = $false
    } | ConvertTo-Json -Depth 10

    # UTF-8 bytes: Windows PowerShell 5.1 would otherwise send the body as ISO-8859-1
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($payload)
    $resp  = Invoke-WebRequest -Method Post -Uri $url -Headers $headers -Body $bytes -ContentType "application/json; charset=utf-8" -UseBasicParsing
    if ($resp.StatusCode -notin 200, 202) {
        throw "Mail send error $($resp.StatusCode): $($resp.Content)"
    }
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
function Invoke-Report {
    param([string]$ConfigPath, [string]$WorkloadsPath, [int]$Days, [string]$OutFile)

    Write-ReportLog "INFO" "Starting M365 roadmap report $SCRIPT_VERSION"
    Import-Config $ConfigPath
    $config = Import-Workload $WorkloadsPath

    $required = @("AZURE_TENANT_ID", "AZURE_CLIENT_ID", "AZURE_CLIENT_SECRET")
    if (-not $OutFile) { $required += @("SENDER_EMAIL", "RECIPIENT_EMAIL") }
    $missing = @($required | Where-Object { -not [System.Environment]::GetEnvironmentVariable($_) })
    if ($missing.Count -gt 0) { throw "Missing configuration: $($missing -join ', ')" }

    $token = Get-AccessToken -TenantId $env:AZURE_TENANT_ID -ClientId $env:AZURE_CLIENT_ID -ClientSecret $env:AZURE_CLIENT_SECRET
    Write-ReportLog "INFO" "Entra ID authentication successful"

    $warnings = @()

    $messages = @()
    try {
        $messages = @(Get-MessageCenterMessage -Token $token -Workloads $config.Workloads -Days $Days)
        Write-ReportLog "INFO" "$($messages.Count) Message Center message(s) retrieved"
    } catch {
        Write-ReportLog "ERROR" "Message Center unavailable: $_"
        $warnings += "Message Center unavailable this week: $_"
    }

    $roadmapItems = @()
    try {
        $roadmapItems = @(Get-RoadmapItem -Workloads $config.Workloads -Days $Days)
        Write-ReportLog "INFO" "$($roadmapItems.Count) roadmap item(s) retrieved"
    } catch {
        Write-ReportLog "ERROR" "Roadmap unavailable: $_"
        $warnings += "M365 Roadmap unavailable this week: $_"
    }

    $now      = Get-Date
    $htmlBody = ConvertTo-ReportHtml -Messages $messages -RoadmapItems $roadmapItems -Warnings $warnings -Title $config.Title -Now $now

    if ($OutFile) {
        [System.IO.File]::WriteAllText($ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutFile), $htmlBody, (New-Object System.Text.UTF8Encoding $false))
        Write-ReportLog "INFO" "Report written to $OutFile"
        return
    }

    $subject = "[M365] $($config.Title) Report $EM_DASH week of $($now.ToString("MM/dd/yyyy"))"
    Send-Report -Token $token -From $env:SENDER_EMAIL -To $env:RECIPIENT_EMAIL -Subject $subject -HtmlBody $htmlBody
    Write-ReportLog "INFO" "Report sent to $env:RECIPIENT_EMAIL"
}

function Show-AvailableWorkload {
    param([string]$ConfigPath)
    Import-Config $ConfigPath
    $token = $null
    if ($env:AZURE_TENANT_ID -and $env:AZURE_CLIENT_ID -and $env:AZURE_CLIENT_SECRET) {
        $token = Get-AccessToken -TenantId $env:AZURE_TENANT_ID -ClientId $env:AZURE_CLIENT_ID -ClientSecret $env:AZURE_CLIENT_SECRET
    } else {
        Write-Warning "Entra ID settings not found: listing Roadmap products only"
    }
    Get-AvailableWorkload -Token $token
}

# Dot-sourcing (tests) only loads the functions
if ($MyInvocation.InvocationName -ne ".") {
    if ($PSVersionTable.PSEdition -ne "Core") {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    }
    try {
        if ($ListWorkloads) { Show-AvailableWorkload -ConfigPath $ConfigPath }
        else { Invoke-Report -ConfigPath $ConfigPath -WorkloadsPath $WorkloadsPath -Days $Days -OutFile $OutFile }
    } catch {
        Write-ReportLog "ERROR" "$_"
        exit 1
    }
}
