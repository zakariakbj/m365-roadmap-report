# Workloads covered by the report.
#
# Each workload has a Name, shown as a badge on its items, and the names to
# match in each source (at least one of the two lists):
#   MessageCenter  Message Center services, for example "SharePoint Online"
#   Roadmap        Microsoft 365 Roadmap products, for example "SharePoint"
# Matching ignores case. Use * as a wildcard, for example "*Purview*".
#
# The two sources do not use the same names. To list them, with the number of
# items for each:
#   pwsh -File ./Send-M365RoadmapReport.ps1 -ListWorkloads
@{
    # Shown in the report header and in the email subject
    Title     = "Exchange, Outlook & Security"

    Workloads = @(
        @{
            Name          = "Exchange & Outlook"
            MessageCenter = @("Exchange Online")
            Roadmap       = @("Exchange", "Outlook")
        }
        @{
            Name          = "Security"
            MessageCenter = @("*Defender for Office 365*", "*Purview*")
            Roadmap       = @("Microsoft Defender for Office 365", "Microsoft Purview")
        }

        # More examples: remove the # signs to add a workload.
        # @{
        #     Name          = "SharePoint & OneDrive"
        #     MessageCenter = @("SharePoint Online", "Microsoft OneDrive")
        #     Roadmap       = @("SharePoint", "OneDrive")
        # }
        # @{
        #     Name          = "Teams"
        #     MessageCenter = @("Microsoft Teams")
        #     Roadmap       = @("Microsoft Teams")
        # }
        # @{
        #     Name          = "Copilot"
        #     MessageCenter = @("Microsoft 365 Copilot Chat")
        #     Roadmap       = @("Microsoft Copilot (Microsoft 365)")
        # }
    )
}
