BeforeAll {
    $RepoRoot = Split-Path -Parent $PSScriptRoot
    . (Join-Path $RepoRoot "Send-M365RoadmapReport.ps1")
    $LogFile = Join-Path $TestDrive "test.log"
    $DefaultWorkloads = (Import-Workload (Join-Path $RepoRoot "workloads.psd1")).Workloads

    function Get-IsoDate([int]$DaysAgo) {
        return (Get-Date).ToUniversalTime().AddDays(-$DaysAgo).ToString("yyyy-MM-ddTHH:mm:ssZ")
    }

    function Get-FakeMessage($Id, $Title, $Services, [int]$DaysAgo = 1) {
        return [PSCustomObject]@{
            id                       = $Id
            title                    = $Title
            messageType              = "planForChange"
            severity                 = "high"
            startDateTime            = Get-IsoDate $DaysAgo
            lastModifiedDateTime     = Get-IsoDate $DaysAgo
            endDateTime              = $null
            actionRequiredByDateTime = $null
            body                     = [PSCustomObject]@{ content = "<p>Description of $Id.</p>" }
            services                 = $Services
        }
    }

    function Get-FakeRoadmapEntry($Id, $Title, $Products, $Status = "In development", [int]$DaysAgo = 1) {
        $productTags = @($Products | ForEach-Object { [PSCustomObject]@{ tagName = $_ } })
        return [PSCustomObject]@{
            id                               = $Id
            title                            = $Title
            description                      = "Description of $Id."
            status                           = $Status
            publicDisclosureAvailabilityDate = "October CY2026"
            modified                         = Get-IsoDate $DaysAgo
            moreInfoLink                     = "https://aka.ms/$Id"
            tags                             = @($productTags) + @([PSCustomObject]@{ tagName = "Web" })
            tagsContainer                    = [PSCustomObject]@{ products = $productTags; platforms = @([PSCustomObject]@{ tagName = "Web" }) }
        }
    }

    function Get-FakeReportMessage($Id = "MC123", $Title = "Test Exchange", $Workloads = @("Exchange & Outlook")) {
        return [PSCustomObject]@{
            id = $Id; title = $Title; severity = "high"; message_type = "planForChange"; workloads = $Workloads
            start_date = "2026-03-08T00:00:00Z"; action_required_by = ""; description = "Desc test"; web_link = ""
        }
    }

    function Get-FakeReportItem($Id = "r001", $Title = "Outlook calendar", $Status = "Rolling out", $Workloads = @("Exchange & Outlook")) {
        return [PSCustomObject]@{
            id = $Id; title = $Title; description = "New view"; status = $Status; workloads = $Workloads
            available_date = "2026-03-05"; modified = ""; more_info_link = "https://aka.ms/$Id"
        }
    }

    function Export-WorkloadFile([string]$Content) {
        $path = Join-Path $TestDrive "workloads.psd1"
        Set-Content -Path $path -Value $Content
        return $path
    }
}

Describe "Version" {
    It "matches the .NOTES header and the latest CHANGELOG.md release" {
        $SCRIPT_VERSION | Should -Match '^\d+\.\d+\.\d+$'
        $script = Get-Content (Join-Path $RepoRoot "Send-M365RoadmapReport.ps1") -Raw
        $script | Should -Match "Version : $([regex]::Escape($SCRIPT_VERSION))\s"
        $changelog = Get-Content (Join-Path $RepoRoot "CHANGELOG.md") -Raw
        [regex]::Match($changelog, '(?m)^## \[(\d+\.\d+\.\d+)\]').Groups[1].Value | Should -Be $SCRIPT_VERSION
        $changelog | Should -Match "(?m)^\[$([regex]::Escape($SCRIPT_VERSION))\]: .*/releases/tag/v$([regex]::Escape($SCRIPT_VERSION))$"
    }

    It "is shown in the report footer" {
        ConvertTo-ReportHtml -Messages @() -RoadmapItems @() -Warnings @() | Should -Match "m365-roadmap-report</a> $([regex]::Escape($SCRIPT_VERSION))"
    }
}

Describe "Import-Config" {
    BeforeEach {
        $names = "M365RR_TEST_A", "M365RR_TEST_B", "M365RR_TEST_C"
        foreach ($n in $names) { [System.Environment]::SetEnvironmentVariable($n, $null) }
    }
    AfterEach {
        foreach ($n in $names) { [System.Environment]::SetEnvironmentVariable($n, $null) }
    }

    It "reads KEY=value lines, strips quotes and skips comments" {
        $path = Join-Path $TestDrive "env.ini"
        Set-Content $path "# comment`nM365RR_TEST_A = value a `nM365RR_TEST_B=`"quoted value`"`n"
        Import-Config $path
        $env:M365RR_TEST_A | Should -Be "value a"
        $env:M365RR_TEST_B | Should -Be "quoted value"
    }

    It "keeps environment variables already set" {
        $env:M365RR_TEST_C = "from environment"
        $path = Join-Path $TestDrive "env.ini"
        Set-Content $path "M365RR_TEST_C=from file"
        Import-Config $path
        $env:M365RR_TEST_C | Should -Be "from environment"
    }

    It "ignores a missing file" {
        { Import-Config (Join-Path $TestDrive "missing.ini") } | Should -Not -Throw
    }
}

Describe "Import-Workload" {
    It "loads the shipped workloads.psd1" {
        $config = Import-Workload (Join-Path $RepoRoot "workloads.psd1")
        $config.Title | Should -Be "Exchange, Outlook & Security"
        $config.Workloads.Name | Should -Be @("Exchange & Outlook", "Security")
        $config.Workloads[0].Roadmap | Should -Be @("Exchange", "Outlook")
    }

    It "builds the title from the workload names when it is not set" {
        $path = Export-WorkloadFile '@{ Workloads = @(@{ Name = "SharePoint"; Roadmap = @("SharePoint") }, @{ Name = "Teams"; MessageCenter = @("Microsoft Teams") }) }'
        $config = Import-Workload $path
        $config.Title | Should -Be "SharePoint, Teams"
        $config.Workloads[0].MessageCenter.Count | Should -Be 0
        $config.Workloads[1].Roadmap.Count | Should -Be 0
    }

    It "rejects a workload without a name" {
        $path = Export-WorkloadFile '@{ Workloads = @(@{ Roadmap = @("SharePoint") }) }'
        { Import-Workload $path } | Should -Throw "*needs a Name*"
    }

    It "rejects a workload without names to match" {
        $path = Export-WorkloadFile '@{ Workloads = @(@{ Name = "Empty" }) }'
        { Import-Workload $path } | Should -Throw "*'Empty' needs MessageCenter or Roadmap names*"
    }

    It "rejects a file without workloads" {
        $path = Export-WorkloadFile '@{ Title = "Nothing" }'
        { Import-Workload $path } | Should -Throw "*no workload defined*"
    }

    It "rejects a missing file" {
        { Import-Workload (Join-Path $TestDrive "missing.psd1") } | Should -Throw "*Workloads file not found*"
    }
}

Describe "Test-NameMatch" {
    It "matches exact names, ignoring case and surrounding spaces" {
        Test-NameMatch @("Microsoft Clipchamp ") @("microsoft clipchamp") | Should -BeTrue
        Test-NameMatch @("SharePoint Syntex") @("SharePoint") | Should -BeFalse
    }

    It "matches wildcards" {
        Test-NameMatch @("Microsoft Purview Data Loss Prevention") @("*Purview*") | Should -BeTrue
        Test-NameMatch @("SharePoint Syntex") @("SharePoint*") | Should -BeTrue
        Test-NameMatch @("Microsoft Teams") @("*Purview*") | Should -BeFalse
    }
}

Describe "Get-AccessToken" {
    It "returns the token of the client credentials flow" {
        Mock Invoke-RestMethod { [PSCustomObject]@{ access_token = "fake_token_123" } }
        Get-AccessToken "tenant" "client" "secret" | Should -Be "fake_token_123"
        Should -Invoke Invoke-RestMethod -Times 1 -ParameterFilter {
            $Uri -eq "https://login.microsoftonline.com/tenant/oauth2/v2.0/token" -and
            $Body.scope -eq "https://graph.microsoft.com/.default" -and $Body.grant_type -eq "client_credentials"
        }
    }

    It "fails when no token is returned" {
        Mock Invoke-RestMethod { [PSCustomObject]@{ error = "invalid_client" } }
        { Get-AccessToken "t" "c" "s" } | Should -Throw "*Authentication failed*"
    }
}

Describe "Get-MessageCenterMessage" {
    It "keeps Exchange Online messages and drops other services" {
        Mock Invoke-RestMethod {
            [PSCustomObject]@{ value = @(
                Get-FakeMessage "MC123456" "Exchange: new retention policy" @("Exchange Online")
                Get-FakeMessage "MC999999" "Teams: interface update" @("Microsoft Teams")
            ) }
        }
        $messages = @(Get-MessageCenterMessage -Token "fake" -Workloads $DefaultWorkloads)
        $messages.Count | Should -Be 1
        $messages[0].id | Should -Be "MC123456"
        $messages[0].workloads | Should -Be @("Exchange & Outlook")
        $messages[0].description | Should -Be "Description of MC123456."
        $messages[0].web_link | Should -BeLike "*/messages/MC123456"
    }

    It "keeps Defender for Office 365 and Purview messages as security" {
        Mock Invoke-RestMethod {
            [PSCustomObject]@{ value = @(
                Get-FakeMessage "MC777777" "Defender rule" @("Microsoft Defender for Office 365")
                Get-FakeMessage "MC888888" "Purview policy" @("Microsoft Purview Data Loss Prevention")
            ) }
        }
        $messages = @(Get-MessageCenterMessage -Token "fake" -Workloads $DefaultWorkloads)
        $messages.id | Should -Be @("MC777777", "MC888888")
        $messages.workloads | Should -Be @("Security", "Security")
    }

    It "reports on any workload added to the configuration" {
        Mock Invoke-RestMethod {
            [PSCustomObject]@{ value = @(
                Get-FakeMessage "MC1" "SharePoint change" @("SharePoint Online")
                Get-FakeMessage "MC2" "Shared change" @("SharePoint Online", "Exchange Online")
                Get-FakeMessage "MC3" "Teams change" @("Microsoft Teams")
            ) }
        }
        $workloads = @($DefaultWorkloads) + [PSCustomObject]@{ Name = "SharePoint"; MessageCenter = @("SharePoint Online"); Roadmap = @() }
        $messages = @(Get-MessageCenterMessage -Token "fake" -Workloads $workloads)
        $messages.id | Should -Be @("MC1", "MC2")
        $messages[1].workloads | Should -Be @("Exchange & Outlook", "SharePoint")
    }

    It "drops messages modified before the window" {
        Mock Invoke-RestMethod { [PSCustomObject]@{ value = @(Get-FakeMessage "MC1" "Old" @("Exchange Online") -DaysAgo 30) } }
        @(Get-MessageCenterMessage -Token "fake" -Workloads $DefaultWorkloads).Count | Should -Be 0
    }

    It "follows @odata.nextLink" {
        Mock Invoke-RestMethod -ParameterFilter { $Uri -like "*page2*" } {
            [PSCustomObject]@{ value = @(Get-FakeMessage "MC2" "Page 2" @("Exchange Online")) }
        }
        Mock Invoke-RestMethod -ParameterFilter { $Uri -notlike "*page2*" } {
            [PSCustomObject]@{ value = @(Get-FakeMessage "MC1" "Page 1" @("Exchange Online")); "@odata.nextLink" = "https://graph.microsoft.com/page2" }
        }
        @(Get-MessageCenterMessage -Token "fake" -Workloads $DefaultWorkloads).id | Should -Be @("MC1", "MC2")
    }

    It "fails on an HTTP error" {
        Mock Invoke-RestMethod { throw "Response status code does not indicate success: 401 (Unauthorized)." }
        { Get-MessageCenterMessage -Token "bad" -Workloads $DefaultWorkloads } | Should -Throw "*401*"
    }
}

Describe "Get-RoadmapItem" {
    It "keeps recent Exchange and Outlook items, matched on product tags" {
        Mock Invoke-RestMethod {
            , @(
                Get-FakeRoadmapEntry "roadmap-001" "Outlook: new calendar view" @("Outlook") "Rolling out" 3
                Get-FakeRoadmapEntry "roadmap-002" "Teams: live reactions" @("Microsoft Teams") "In development" 2
                Get-FakeRoadmapEntry "roadmap-003" "Exchange: transport rules" @("Exchange") "In development" 30
                Get-FakeRoadmapEntry "roadmap-004" "Project Operations: manage exchange rate variances" @("Dynamics 365 Project Operations")
            )
        }
        $items = @(Get-RoadmapItem -Workloads $DefaultWorkloads)
        $items.id | Should -Be @("roadmap-001")
        $items[0].workloads | Should -Be @("Exchange & Outlook")
    }

    It "keeps Defender for Office 365 and Purview items as security" {
        Mock Invoke-RestMethod {
            , @(
                Get-FakeRoadmapEntry "roadmap-101" "New impersonation detection" @("Microsoft Defender for Office 365")
                Get-FakeRoadmapEntry "roadmap-102" "New eDiscovery export option" @("Microsoft Purview")
            )
        }
        $items = @(Get-RoadmapItem -Workloads $DefaultWorkloads)
        $items.id | Should -Be @("roadmap-101", "roadmap-102")
        $items.workloads | Should -Be @("Security", "Security")
    }

    It "reports on any workload added to the configuration" {
        Mock Invoke-RestMethod {
            , @(
                Get-FakeRoadmapEntry "1" "SharePoint feature" @("SharePoint")
                Get-FakeRoadmapEntry "2" "OneDrive feature" @("OneDrive", "SharePoint")
                Get-FakeRoadmapEntry "3" "Teams feature" @("Microsoft Teams")
            )
        }
        $workloads = @([PSCustomObject]@{ Name = "SharePoint & OneDrive"; MessageCenter = @(); Roadmap = @("SharePoint", "OneDrive") })
        $items = @(Get-RoadmapItem -Workloads $workloads)
        $items.id | Should -Be @("1", "2")
        $items[1].workloads | Should -Be @("SharePoint & OneDrive")
    }

    It "falls back to the flat tags when tagsContainer is missing" {
        Mock Invoke-RestMethod {
            $entry = Get-FakeRoadmapEntry "1" "Outlook feature" @("Outlook")
            $entry.PSObject.Properties.Remove("tagsContainer")
            , @($entry)
        }
        @(Get-RoadmapItem -Workloads $DefaultWorkloads).id | Should -Be @("1")
    }

    It "builds the roadmap link when moreInfoLink is missing" {
        Mock Invoke-RestMethod {
            $entry = Get-FakeRoadmapEntry "424242" "Outlook feature" @("Outlook")
            $entry.moreInfoLink = $null
            , @($entry)
        }
        @(Get-RoadmapItem -Workloads $DefaultWorkloads)[0].more_info_link | Should -Be "https://www.microsoft.com/en-us/microsoft-365/roadmap?id=424242"
    }

    It "turns HTML descriptions into plain text" {
        Mock Invoke-RestMethod {
            $entry = Get-FakeRoadmapEntry "1" "Outlook feature" @("Outlook")
            $entry.description = "<p>First&nbsp;line.</p><br><p>Second &amp; last.</p>"
            , @($entry)
        }
        @(Get-RoadmapItem -Workloads $DefaultWorkloads)[0].description | Should -Be "First line. Second & last."
    }

    It "fails on an HTTP error" {
        Mock Invoke-RestMethod { throw "Response status code does not indicate success: 500 (Internal Server Error)." }
        { Get-RoadmapItem -Workloads $DefaultWorkloads } | Should -Throw "*500*"
    }
}

Describe "Get-AvailableWorkload" {
    It "lists Roadmap products, and Message Center services with a token" {
        Mock Get-RoadmapData {
            Get-FakeRoadmapEntry "1" "A" @("SharePoint")
            Get-FakeRoadmapEntry "2" "B" @("SharePoint", "OneDrive")
        }
        Mock Get-MessageCenterData { Get-FakeMessage "MC1" "A" @("SharePoint Online ") }
        $names = @(Get-AvailableWorkload -Token "fake")
        $names | Where-Object Source -eq "Roadmap" | ForEach-Object { "$($_.Name)=$($_.Items)" } | Should -Be @("OneDrive=1", "SharePoint=2")
        ($names | Where-Object Source -eq "MessageCenter").Name | Should -Be "SharePoint Online"
    }

    It "lists only Roadmap products without a token" {
        Mock Get-RoadmapData { Get-FakeRoadmapEntry "1" "A" @("Planner") }
        Mock Get-MessageCenterData { throw "must not be called" }
        @(Get-AvailableWorkload -Token $null).Source | Should -Be @("Roadmap")
    }
}

Describe "Get-Excerpt" {
    It "keeps texts up to the length plus the leeway" {
        $text = "a" * 305
        Get-Excerpt $text | Should -Be $text
    }

    It "cuts long texts at a word boundary" {
        $text = ("word " * 80).Trim()
        $excerpt = Get-Excerpt $text
        $excerpt | Should -Match "^(word )+word\.\.\.$"
        $excerpt.Length | Should -BeLessOrEqual 300
    }
}

Describe "ConvertTo-ReportHtml" {
    It "includes the Message Center data" {
        $html = ConvertTo-ReportHtml -Messages @(Get-FakeReportMessage) -RoadmapItems @() -Warnings @()
        $html | Should -Match "MC123"
        $html | Should -Match "Test Exchange"
        $html | Should -Match 'class="badge high">HIGH'
    }

    It "includes the roadmap data" {
        $html = ConvertTo-ReportHtml -Messages @() -RoadmapItems @(Get-FakeReportItem) -Warnings @()
        $html | Should -Match "Outlook calendar"
        $html | Should -Match "ROLLING OUT"
    }

    It "shows a badge for each workload of an item" {
        $html = ConvertTo-ReportHtml -Messages @(Get-FakeReportMessage -Workloads @("Security")) -RoadmapItems @(Get-FakeReportItem -Workloads @("SharePoint", "Teams")) -Warnings @()
        $html | Should -Match 'class="badge workload">SECURITY<'
        $html | Should -Match 'class="badge workload">SHAREPOINT</span> <span class="badge workload">TEAMS<'
    }

    It "shows the roadmap status" {
        $items = @("In development", "Rolling out", "Launched", "Cancelled", "Something new") | ForEach-Object { Get-FakeReportItem -Status $_ }
        $html = ConvertTo-ReportHtml -Messages @() -RoadmapItems $items -Warnings @()
        $html | Should -Match 'class="badge dev">IN DEVELOPMENT<'
        $html | Should -Match 'class="badge rolling">ROLLING OUT<'
        $html | Should -Match 'class="badge launched">LAUNCHED<'
        $html | Should -Match 'class="badge cancelled">CANCELLED<'
        $html | Should -Match 'class="badge dev">SOMETHING NEW<'
    }

    It "shows the title" {
        ConvertTo-ReportHtml -Messages @() -RoadmapItems @() -Warnings @() -Title "SharePoint & Teams" | Should -Match "M365 Weekly Report &mdash; SharePoint &amp; Teams"
    }

    It "shows the warning banner" {
        ConvertTo-ReportHtml -Messages @() -RoadmapItems @() -Warnings @("Roadmap API unavailable") | Should -Match "Roadmap API unavailable"
    }

    It "shows the empty texts when there is no data" {
        $html = ConvertTo-ReportHtml -Messages @() -RoadmapItems @() -Warnings @()
        $html | Should -Match "No Message Center events this week."
        $html | Should -Match "No roadmap items this week."
    }

    It "escapes HTML from the sources" {
        $html = ConvertTo-ReportHtml -Messages @(Get-FakeReportMessage -Title "<script>alert(1)</script>") -RoadmapItems @(Get-FakeReportItem -Title "A & B <b>") -Warnings @("<i>warning</i>")
        $html | Should -Not -Match "<script>"
        $html | Should -Match "&lt;script&gt;alert\(1\)&lt;/script&gt;"
        $html | Should -Match "A &amp; B &lt;b&gt;"
        $html | Should -Match "&lt;i&gt;warning&lt;/i&gt;"
    }

    It "starts the week on Monday, also on Sunday" {
        $sunday = New-Object DateTime 2026, 10, 25, 9, 0, 0
        ConvertTo-ReportHtml -Messages @() -RoadmapItems @() -Warnings @() -Now $sunday | Should -Match "Week of 10/19/2026 to 10/25/2026"
        $monday = New-Object DateTime 2026, 10, 19, 8, 0, 0
        ConvertTo-ReportHtml -Messages @() -RoadmapItems @() -Warnings @() -Now $monday | Should -Match "Week of 10/19/2026 to 10/25/2026"
    }
}

Describe "Send-Report" {
    It "posts the message to the sender's sendMail endpoint" {
        Mock Invoke-WebRequest { [PSCustomObject]@{ StatusCode = 202; Content = "" } }
        Send-Report -Token "fake" -From "sender@contoso.com" -To "dest@contoso.com" -Subject "Test report" -HtmlBody "<h1>Report</h1>"
        Should -Invoke Invoke-WebRequest -Times 1 -ParameterFilter {
            $payload = [System.Text.Encoding]::UTF8.GetString($Body) | ConvertFrom-Json
            $Uri -eq "https://graph.microsoft.com/v1.0/users/sender@contoso.com/sendMail" -and
            $payload.message.subject -eq "Test report" -and
            $payload.message.toRecipients[0].emailAddress.address -eq "dest@contoso.com" -and
            $payload.saveToSentItems -eq $false
        }
    }

    It "fails on an unexpected status" {
        Mock Invoke-WebRequest { [PSCustomObject]@{ StatusCode = 204; Content = "nothing" } }
        { Send-Report "fake" "s@x.com" "d@x.com" "Subject" "<p>body</p>" } | Should -Throw "*Mail send error 204*"
    }
}

Describe "Invoke-Report" {
    BeforeAll {
        $configNames = "AZURE_TENANT_ID", "AZURE_CLIENT_ID", "AZURE_CLIENT_SECRET", "SENDER_EMAIL", "RECIPIENT_EMAIL"
        $saved = @{}
        foreach ($n in $configNames) { $saved[$n] = [System.Environment]::GetEnvironmentVariable($n) }
    }
    BeforeEach {
        foreach ($n in $configNames) { [System.Environment]::SetEnvironmentVariable($n, "value-$n") }
        Mock Write-Information { }
        Mock Get-AccessToken { "token" }
        Mock Get-RoadmapItem { @() }
        Mock Send-Report { }
    }
    AfterAll {
        foreach ($n in $configNames) { [System.Environment]::SetEnvironmentVariable($n, $saved[$n]) }
    }

    It "sends the report" {
        Mock Get-MessageCenterMessage { @(Get-FakeReportMessage -Id "MC001") }
        Invoke-Report -ConfigPath (Join-Path $TestDrive "none.ini") -WorkloadsPath (Join-Path $RepoRoot "workloads.psd1") -Days 7
        Should -Invoke Send-Report -Times 1 -ParameterFilter {
            $HtmlBody -match "MC001" -and $To -eq "value-RECIPIENT_EMAIL" -and $Subject -like '`[M365`] Exchange, Outlook & Security Report * week of *'
        }
    }

    It "uses the workloads and title of the workloads file" {
        $path = Export-WorkloadFile '@{ Title = "SharePoint"; Workloads = @(@{ Name = "SharePoint"; Roadmap = @("SharePoint") }) }'
        Mock Get-MessageCenterMessage { @() }
        Invoke-Report -ConfigPath (Join-Path $TestDrive "none.ini") -WorkloadsPath $path -Days 14
        Should -Invoke Get-RoadmapItem -Times 1 -ParameterFilter { $Workloads[0].Name -eq "SharePoint" -and $Days -eq 14 }
        Should -Invoke Send-Report -Times 1 -ParameterFilter { $Subject -like '`[M365`] SharePoint Report *' }
    }

    It "sends the report with a warning when the Message Center fails" {
        Mock Get-MessageCenterMessage { throw "API down" }
        Invoke-Report -ConfigPath (Join-Path $TestDrive "none.ini") -WorkloadsPath (Join-Path $RepoRoot "workloads.psd1") -Days 7
        Should -Invoke Send-Report -Times 1 -ParameterFilter { $HtmlBody -match "Message Center unavailable this week: API down" }
    }

    It "writes the report to a file without sending it, sender and recipient not needed" {
        Mock Get-MessageCenterMessage { @() }
        [System.Environment]::SetEnvironmentVariable("SENDER_EMAIL", $null)
        [System.Environment]::SetEnvironmentVariable("RECIPIENT_EMAIL", $null)
        $out = Join-Path $TestDrive "report.html"
        Invoke-Report -ConfigPath (Join-Path $TestDrive "none.ini") -WorkloadsPath (Join-Path $RepoRoot "workloads.psd1") -Days 7 -OutFile $out
        Get-Content $out -Raw | Should -Match "M365 Weekly Report"
        Should -Invoke Send-Report -Times 0
    }

    It "fails on missing configuration" {
        [System.Environment]::SetEnvironmentVariable("AZURE_CLIENT_SECRET", $null)
        { Invoke-Report -ConfigPath (Join-Path $TestDrive "none.ini") -WorkloadsPath (Join-Path $RepoRoot "workloads.psd1") -Days 7 } | Should -Throw "*Missing configuration: AZURE_CLIENT_SECRET*"
    }
}
