**English** · [Nederlands](readme.nl.md) · [Français](readme.fr.md)

[M365-Scripts](../../../readme.md) › [scripts](../../readme.md) › [Office365Toolkit](../readme.md) › **Exchange**

# Office365Toolkit / Exchange

Mailbox hygiene baseline, inbox-rule forwarding risk, mailbox add-ins and
Unified Audit Log search.

**Sign-in.** Every script signs in through
[`Connect-M365.ps1`](../../Startup/readme.md#connect-m365ps1): delegated (you sign
in as the admin) by default, with device code and the GDAP customer taken from
`load.config.ps1` — under GDAP Exchange Online is reached with
`-DelegatedOrganization`. App-only with `-ClientId` + `-CertificateThumbprint`,
or `-AppOnly` (ClientId and thumbprint from `graph.appid.json`). A session that
already fits is reused and left connected; the scripts only disconnect what they
opened themselves. `Search-MailboxAuditLog.ps1` uses Microsoft Graph; the other
three stay on Exchange Online because Graph has no API for what they read (see
each script's Notes).

> These complement, and don't duplicate, the existing Exchange audit scripts in
> [`scripts/Exchange/`](../../Exchange/readme.md) — see each script's Notes
> section below for the exact boundary.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Test-MailboxSecurityBaseline.ps1`](Test-MailboxSecurityBaseline.ps1) ([docs](#test-mailboxsecuritybaselineps1)) | Audit mailboxes against a hygiene baseline (audit logging, retention, litigation hold, archive, legacy protocols) |
| [`Test-MailboxForwardingRisk.ps1`](Test-MailboxForwardingRisk.ps1) ([docs](#test-mailboxforwardingriskps1)) | Audit inbox rules and Sweep rules for forwarding/exfiltration patterns (BEC indicator) |
| [`Get-MailboxAddIns.ps1`](Get-MailboxAddIns.ps1) ([docs](#get-mailboxaddinsps1)) | Report Outlook add-ins installed per mailbox |
| [`Search-MailboxAuditLog.ps1`](Search-MailboxAuditLog.ps1) ([docs](#search-mailboxauditlogps1)) | Search the Unified Audit Log (Graph Audit Log Query API) for sign-in and mailbox login events |

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
| `-TenantId` | No | Tenant ID or domain; defaults to the GDAP customer from `load.config.ps1` (app-only needs the domain form) |
| `-ClientId` | No | App registration for app-only sign-in (with `-CertificateThumbprint`) |
| `-CertificateThumbprint` | No | Certificate thumbprint for app-only sign-in |
| `-AppOnly` | No | App-only with ClientId and thumbprint from `graph.appid.json` |

**Examples**

```powershell
.\Test-MailboxSecurityBaseline.ps1

.\Test-MailboxSecurityBaseline.ps1 -Mailbox "user@contoso.com"

.\Test-MailboxSecurityBaseline.ps1 -MinAuditLogAgeLimitDays 180 -MinRetainDeletedItemsDays 30
```

**Notes**
- Stays on Exchange Online: audit, retention, hold, archive, forwarding and
  POP/IMAP settings (`Get-Mailbox` / `Get-CASMailbox`) have no Graph API.
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
| `-TenantId` | No | Tenant ID or domain; defaults to the GDAP customer from `load.config.ps1` (app-only needs the domain form) |
| `-ClientId` | No | App registration for app-only sign-in (with `-CertificateThumbprint`) |
| `-CertificateThumbprint` | No | Certificate thumbprint for app-only sign-in |
| `-AppOnly` | No | App-only with ClientId and thumbprint from `graph.appid.json` |

**Examples**

```powershell
.\Test-MailboxForwardingRisk.ps1

.\Test-MailboxForwardingRisk.ps1 -Mailbox "user@contoso.com" -IncludeDisabledRules
```

**Notes**
- Stays on Exchange Online: Graph reads another user's inbox rules
  (`messageRules`) only with an app-only Mail permission, and has no API for
  Sweep rules.
- A rule recipient inside the organization (`[EX:/o=...]`) now counts as
  Internal; it used to be flagged External because it has no domain to compare.
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
| `-TenantId` | No | Tenant ID or domain; defaults to the GDAP customer from `load.config.ps1` (app-only needs the domain form) |
| `-ClientId` | No | App registration for app-only sign-in (with `-CertificateThumbprint`) |
| `-CertificateThumbprint` | No | Certificate thumbprint for app-only sign-in |
| `-AppOnly` | No | App-only with ClientId and thumbprint from `graph.appid.json` |

**Examples**

```powershell
.\Get-MailboxAddIns.ps1

.\Get-MailboxAddIns.ps1 -Mailbox "user@contoso.com"
```

**Notes**
- Stays on Exchange Online: Outlook add-ins per mailbox (`Get-App`) have no
  Graph API.

**Required module:** `ExchangeOnlineManagement`

---

### Search-MailboxAuditLog.ps1

Generic Unified Audit Log search. By default it uses the Microsoft Graph Audit
Log Query API: it creates a query (`POST /security/auditLog/queries`), polls it
every 30 seconds until it has succeeded, then pages through the records.
Defaults to the last 2 days covering interactive sign-ins (success/failure) and
mailbox logins. Fully parameterized for other record types/operations/date
ranges/users. `-UseExchange` runs the same search with `Search-UnifiedAuditLog`
in Exchange Online instead. Read-only.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-Days` | No | Days back from now to search (default `2`); ignored if `-StartDate` given |
| `-StartDate` | No | Explicit window start (overrides `-Days`) |
| `-EndDate` | No | Explicit window end (default: now) |
| `-RecordType` | No | Audit log record type(s) (default `AzureActiveDirectoryStsLogon`, `ExchangeItem`) |
| `-Operations` | No | Operation name(s) (default `UserLoggedIn`, `UserLoginFailed`, `MailboxLogin`) |
| `-UserIds` | No | Restrict to specific user(s) (UPN) |
| `-UseExchange` | No | Search with `Search-UnifiedAuditLog` in Exchange Online instead of Graph |
| `-TimeoutMinutes` | No | How long to wait for the Graph query (default `60`) |
| `-OutputPath` | No | CSV report path |
| `-TenantId` | No | Tenant ID or domain; defaults to the GDAP customer from `load.config.ps1` |
| `-ClientId` | No | App registration for app-only sign-in (with `-CertificateThumbprint`) |
| `-CertificateThumbprint` | No | Certificate thumbprint for app-only sign-in |
| `-AppOnly` | No | App-only with ClientId and thumbprint from `graph.appid.json` |

**Examples**

```powershell
# Last 2 days, sign-ins + mailbox logins
.\Search-MailboxAuditLog.ps1

# Last 30 days, failed sign-ins only, one user
.\Search-MailboxAuditLog.ps1 -Days 30 -RecordType AzureActiveDirectoryStsLogon -Operations UserLoginFailed -UserIds "user@contoso.com"

# Explicit window
.\Search-MailboxAuditLog.ps1 -StartDate (Get-Date "2026-07-01") -EndDate (Get-Date "2026-07-15")

# Same search through Exchange Online
.\Search-MailboxAuditLog.ps1 -UseExchange
```

**Notes**
- The Graph query runs asynchronously in the service and can take several
  minutes. After `-TimeoutMinutes` the script stops waiting and prints the query
  id; the query keeps running and can be read later.
- Record types are given in their audit-log form (`ExchangeItem`); the script
  converts them to the API's camelCase (`exchangeItem`).
- `Search-UnifiedAuditLog` takes one record type per call. Earlier versions
  passed both default record types in one call; `-UseExchange` now runs one
  paged search per record type.
- The Unified Audit Log is not immediate — allow up to 30-60 minutes for recent
  activity to appear.

**Required scope:** `AuditLogsQuery.Read.All` (plus a Purview role that may search the audit log, e.g. Audit Logs)
**Required modules:** `Microsoft.Graph.Authentication`; `ExchangeOnlineManagement` for `-UseExchange`
