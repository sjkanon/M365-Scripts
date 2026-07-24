# Patron Toolkit — Exchange

Mail flow diagnostics via Exchange Online.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Get-MessageTraceReport.ps1`](#get-messagetracereportps1) | Export a message trace (mail flow) report, with optional per-message delivery detail |

---

### Get-MessageTraceReport.ps1

Runs a message trace for a given time window (default: last 48 hours), optionally
filtered by sender, recipient, and delivery status. Exports a summary CSV and, with
`-IncludeDetail`, a per-message delivery detail CSV.

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
| `-OutputPath` | No | Folder for the CSV report(s) (default: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | No | Entra ID tenant ID or domain |

**Examples**

```powershell
.\Get-MessageTraceReport.ps1

.\Get-MessageTraceReport.ps1 -SenderAddress "user@contoso.com" -Hours 24

.\Get-MessageTraceReport.ps1 -RecipientAddress "user@contoso.com" -Status Failed -IncludeDetail
```

**Notes**
- `Get-MessageTrace` only covers the last 10 days — for older mail, use the Exchange
  admin center's historical search instead

**Required module**
```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```
