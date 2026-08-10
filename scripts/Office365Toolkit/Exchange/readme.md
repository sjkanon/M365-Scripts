# Office365Toolkit / Exchange

Mailbox hygiene baseline, inbox-rule forwarding risk, mailbox add-ins, Unified
Audit Log search, and message trace reporting. Connect to Exchange Online
automatically if no session is active; reuse an existing session if already
connected.

> These complement, and don't duplicate, the existing Exchange audit scripts in
> [`scripts/Exchange/`](../../Exchange/readme.md) — see each script's Notes
> section below for the exact boundary.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Test-MailboxSecurityBaseline.ps1`](#test-mailboxsecuritybaselineps1) | Audit mailboxes against a hygiene baseline (audit logging, retention, litigation hold, archive, legacy protocols) |
| [`Test-MailboxForwardingRisk.ps1`](#test-mailboxforwardingriskps1) | Audit inbox rules and Sweep rules for forwarding/exfiltration patterns (BEC indicator) |
| [`Get-MailboxAddIns.ps1`](#get-mailboxaddinsps1) | Report Outlook add-ins installed per mailbox |
| [`Search-MailboxAuditLog.ps1`](#search-mailboxauditlogps1) | Search the Unified Audit Log for sign-in and mailbox login events |

> Message trace reporting lives in [`scripts/PatronToolkit/Exchange/Get-MessageTraceReport.ps1`](../../PatronToolkit/Exchange/readme.md#get-messagetracereportps1) — an equivalent script was built independently for both toolkits, so only one was kept (with `-IncludeDetail` support merged in from this one).

---

### Test-MailboxSecurityBaseline.ps1

Checks every mailbox (or a single one) against a hygiene baseline: audit
logging enabled + log age limit, deleted item retention, litigation hold,
archive status, no mailbox-level forwarding, POP3/IMAP disabled. Each mailbox
gets a Pass/Fail per check and an overall `Status`. Read-only.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-Mailbox` | No | UPN of a single mailbox. If omitted, all user and shared mailboxes are checked |
| `-MinAuditLogAgeLimitDays` | No | Minimum acceptable audit log age limit in days (default `90`) |
| `-MinRetainDeletedItemsDays` | No | Minimum acceptable deleted item retention in days (default `30`) |
| `-OutputPath` | No | CSV report path (default: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | No | Entra ID tenant ID or domain |

**Examples**

```powershell
.\Test-MailboxSecurityBaseline.ps1

.\Test-MailboxSecurityBaseline.ps1 -Mailbox "user@contoso.com"

.\Test-MailboxSecurityBaseline.ps1 -MinAuditLogAgeLimitDays 180 -MinRetainDeletedItemsDays 30
```

**Notes**
- Mailbox-level external forwarding is a simple present/absent check here; for
  domain-aware external-vs-internal breakdown use
  [`Get-ExternalForwards.ps1`](../../Exchange/readme.md#get-externalforwardsps1).
- For inbox-rule/Sweep-rule based forwarding, use `Test-MailboxForwardingRisk.ps1`
  below instead.

**Required module:** `ExchangeOnlineManagement`

---

### Test-MailboxForwardingRisk.ps1

Audits every mailbox's inbox rules and Sweep rules for forwarding/redirect/
exfiltration patterns — a classic business email compromise (BEC) indicator
that doesn't show up on the mailbox object itself. Flags rule recipients as
External/Internal/Unknown based on the tenant's accepted domains. Read-only.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-Mailbox` | No | UPN of a single mailbox. If omitted, all mailboxes are checked |
| `-IncludeDisabledRules` | No | Also report disabled rules matching the risky patterns |
| `-OutputPath` | No | CSV report path |
| `-TenantId` | No | Entra ID tenant ID or domain |

**Examples**

```powershell
.\Test-MailboxForwardingRisk.ps1

.\Test-MailboxForwardingRisk.ps1 -Mailbox "user@contoso.com" -IncludeDisabledRules
```

**Notes**
- Complements [`Get-ExternalForwards.ps1`](../../Exchange/readme.md#get-externalforwardsps1)
  (mailbox-level `ForwardingSmtpAddress`) and `Test-MailboxSecurityBaseline.ps1`
  above (same check) — this script covers the inbox-rule/Sweep-rule layer only.

**Required module:** `ExchangeOnlineManagement`

---

### Get-MailboxAddIns.ps1

Lists Outlook add-ins (`Get-App`) present on each mailbox — centrally deployed
and user/sideloaded. Useful for spotting unapproved add-ins, a known
phishing/consent-grant vector. Read-only.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-Mailbox` | No | UPN of a single mailbox. If omitted, all user and shared mailboxes are checked |
| `-OutputPath` | No | CSV report path |
| `-TenantId` | No | Entra ID tenant ID or domain |

**Examples**

```powershell
.\Get-MailboxAddIns.ps1

.\Get-MailboxAddIns.ps1 -Mailbox "user@contoso.com"
```

**Required module:** `ExchangeOnlineManagement`

---

### Search-MailboxAuditLog.ps1

Generic Unified Audit Log search wrapper (`Search-UnifiedAuditLog`), paginating
through all matching results. Defaults to the last 2 days covering interactive
sign-ins (success/failure) and mailbox logins. Fully parameterized for other
record types/operations/date ranges/users. Read-only.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-Days` | No | Days back from now to search (default `2`); ignored if `-StartDate` given |
| `-StartDate` | No | Explicit window start (overrides `-Days`) |
| `-EndDate` | No | Explicit window end (default: now) |
| `-RecordType` | No | Audit log record type(s) (default `AzureActiveDirectoryStsLogon`, `ExchangeItem`) |
| `-Operations` | No | Operation name(s) (default `UserLoggedIn`, `UserLoginFailed`, `MailboxLogin`) |
| `-UserIds` | No | Restrict to specific user(s) |
| `-OutputPath` | No | CSV report path |
| `-TenantId` | No | Entra ID tenant ID or domain |

**Examples**

```powershell
# Last 2 days, sign-ins + mailbox logins
.\Search-MailboxAuditLog.ps1

# Last 30 days, failed sign-ins only, one user
.\Search-MailboxAuditLog.ps1 -Days 30 -RecordType AzureActiveDirectoryStsLogon -Operations UserLoginFailed -UserIds "user@contoso.com"

# Explicit window
.\Search-MailboxAuditLog.ps1 -StartDate (Get-Date "2026-07-01") -EndDate (Get-Date "2026-07-15")
```

**Notes**
- The Unified Audit Log is not immediate — allow up to 30-60 minutes for recent
  activity to appear.

**Required module:** `ExchangeOnlineManagement`

---

### Get-MessageTraceReport.ps1

Reports mail flow for a recent window (message trace only covers ~10 days),
with optional sender/recipient/status filters. Prefers the newer
`Get-MessageTraceV2` cmdlet, falling back to classic `Get-MessageTrace` on
older module versions. Read-only.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-Hours` | No | Hours back from now to search (default `48`); ignored if `-StartDate` given |
| `-StartDate` | No | Explicit window start |
| `-EndDate` | No | Explicit window end (default: now) |
| `-SenderAddress` | No | Filter to a specific sender |
| `-RecipientAddress` | No | Filter to a specific recipient |
| `-Status` | No | Filter to a delivery status (e.g. `Delivered`, `Failed`, `Quarantined`) |
| `-OutputPath` | No | CSV report path |
| `-TenantId` | No | Entra ID tenant ID or domain |

**Examples**

```powershell
.\Get-MessageTraceReport.ps1

.\Get-MessageTraceReport.ps1 -Hours 24 -RecipientAddress "user@contoso.com"

.\Get-MessageTraceReport.ps1 -SenderAddress "billing@vendor.com" -Status Failed
```

**Required module:** `ExchangeOnlineManagement`
