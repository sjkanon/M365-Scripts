**English** · [Nederlands](readme.nl.md) · [Français](readme.fr.md)

[M365-Scripts](../../../readme.md) › [scripts](../../readme.md) › [PatronToolkit](../readme.md) › **Exchange**

# Patron Toolkit — Exchange

Mail flow diagnostics via Exchange Online.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Get-MessageTraceReport.ps1`](Get-MessageTraceReport.ps1) ([docs](#get-messagetracereportps1)) | Export a message trace (mail flow) report, with optional per-message delivery detail |

---

### Get-MessageTraceReport.ps1

Runs a message trace for a given time window (default: last 48 hours), optionally
filtered by sender, recipient, and delivery status, with `Get-MessageTraceV2`. Exports a
summary CSV and, with `-IncludeDetail`, a per-message delivery detail CSV
(`Get-MessageTraceDetailV2`).

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-Hours` | No | Hours back to trace from now (default: `48`). Ignored if `-StartDate` is given |
| `-StartDate` | No | Explicit start of the trace window |
| `-EndDate` | No | Explicit end of the trace window (default: now) |
| `-SenderAddress` | No | Filter to a specific sender |
| `-RecipientAddress` | No | Filter to a specific recipient |
| `-Status` | No | Filter by delivery status (e.g. `Delivered`, `Failed`, `Pending`, `Quarantined`) |
| `-IncludeDetail` | No | Also export per-message delivery detail (slower on large result sets) |
| `-MaxResults` | No | Stop paging after this many trace rows (default: `50000`) |
| `-OutputPath` | No | Folder for the CSV report(s) (default: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | No | Tenant ID or domain. Defaults to the GDAP customer (`load.config.ps1`) or your own tenant; required for app-only |
| `-ClientId` | No | App registration for app-only sign-in (with `-CertificateThumbprint`). Without it the script signs in delegated, as you |
| `-CertificateThumbprint` | No | Certificate thumbprint for app-only sign-in with `-ClientId` |
| `-AppOnly` | No | App-only sign-in with the ClientId and thumbprint for the tenant from `graph.appid.json` |

**Examples**

```powershell
.\Get-MessageTraceReport.ps1

.\Get-MessageTraceReport.ps1 -SenderAddress "user@contoso.com" -Hours 24

.\Get-MessageTraceReport.ps1 -RecipientAddress "user@contoso.com" -Status Failed -IncludeDetail
```

**Notes**
- Message trace V2 reaches back 90 days; an older `-StartDate` is cut to 90 days with a
  warning. One `Get-MessageTraceV2` call covers at most 10 days and 5000 rows, so the
  script splits the window into 10-day slices and pages through each one
  (`EndDate` + `StartingRecipientAddress` from the last row) — earlier versions stopped at
  5000 rows without saying so
- The retired `Get-MessageTrace` / `Get-MessageTraceDetail` are no longer used; the script
  stops with a clear message when the session has no V2 cmdlets (update
  ExchangeOnlineManagement)
- Sign-in via [`Connect-M365.ps1`](../../Startup/readme.md#connect-m365ps1): Exchange
  Online, delegated by default (device code and the GDAP customer via
  `-DelegatedOrganization`, per `load.config.ps1`); app-only with `-ClientId` +
  `-CertificateThumbprint` or `-AppOnly` (needs `-TenantId` as a domain). Stays on
  Exchange Online because Microsoft Graph has no message trace API

**Required module**
```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```
