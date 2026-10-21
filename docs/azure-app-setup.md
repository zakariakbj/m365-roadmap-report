# Entra ID app registration

The script signs in as an application (client credentials flow) and needs two Microsoft Graph application permissions:

| Permission | Used for |
|---|---|
| `ServiceMessage.Read.All` | Reading the Message Center |
| `Mail.Send` | Sending the report from the sender mailbox |

The Microsoft 365 Roadmap is a public API and needs no permission.

## 1. Register the application

You need the **Application Administrator** role, or Global Administrator to grant the consent.

1. In the [Microsoft Entra admin center](https://entra.microsoft.com), go to **Identity > Applications > App registrations > New registration**.
2. Name: `M365 Roadmap Report`. Supported account types: **Accounts in this organizational directory only**. No redirect URI.
3. Select **Register**, then note:
   - **Application (client) ID**: `AZURE_CLIENT_ID`
   - **Directory (tenant) ID**: `AZURE_TENANT_ID`

## 2. Create a client secret

1. Go to **Certificates & secrets > Client secrets > New client secret**.
2. Pick an expiry (24 months at most) and select **Add**.
3. Copy the **Value** right away into `AZURE_CLIENT_SECRET`. It is shown only once.

Set a reminder before the expiry date: the report stops when the secret expires, and the log shows an authentication error.

## 3. Grant the Message Center permission

1. Go to **API permissions > Add a permission > Microsoft Graph > Application permissions**.
2. Add `ServiceMessage.Read.All`.
3. Select **Grant admin consent for <tenant>**.

## 4. Allow sending from the sender mailbox only

Granted in Entra ID, the `Mail.Send` application permission lets the app send as **any** mailbox in the tenant. Grant it through Exchange Online instead, limited to the sender mailbox, with [RBAC for Applications](https://learn.microsoft.com/exchange/permissions-online/application-rbac).

In Exchange Online PowerShell (`Connect-ExchangeOnline`), with the Exchange Administrator role:

```powershell
$appId    = "<AZURE_CLIENT_ID>"
$objectId = "<object ID of the app's Enterprise application, not of the app registration>"
$sender   = "it-reports@contoso.com"

New-ServicePrincipal -AppId $appId -ObjectId $objectId -DisplayName "M365 Roadmap Report"
New-ManagementScope -Name "M365 Roadmap Report sender" -RecipientRestrictionFilter "PrimarySmtpAddress -eq '$sender'"
New-ManagementRoleAssignment -App $appId -Role "Application Mail.Send" -CustomResourceScope "M365 Roadmap Report sender"

# Check: InScope must be True for the sender and False for any other mailbox
Test-ServicePrincipalAuthorization -Identity $appId -Resource $sender
```

Do not add `Mail.Send` in Entra ID as well. Exchange combines both grants, so the Entra ID permission would allow every mailbox again.

If you prefer the Entra ID permission, add `Mail.Send` in step 3 next to `ServiceMessage.Read.All`.

## 5. Sender mailbox

`SENDER_EMAIL` must be an Exchange Online mailbox in the tenant, either licensed or a shared mailbox. Reports are not saved to its Sent Items.

## 6. Test

Fill in `env.ini` (see the README), then render the report without sending it:

```powershell
pwsh -File ./Send-M365RoadmapReport.ps1 -OutFile report.html
```

This checks the sign-in and the Message Center permission. Run it without `-OutFile` to check the sending permission.
