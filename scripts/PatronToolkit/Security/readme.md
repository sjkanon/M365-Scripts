# Patron Toolkit — Security

Tenant security posture reporting: app consent risk, mailbox-rule BEC detection, security
alerts, email security configuration, audit logging, and SPF/DMARC validation.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Get-EntraAppConsents.ps1`](#get-entraappconsentsps1) | Audit OAuth delegated + application permission grants tenant-wide, flag high-risk scopes |
| [`Get-SuspiciousInboxRules.ps1`](#get-suspiciousinboxrulesps1) | Detect BEC-style inbox rules (external forward, silent delete, hidden-folder + keyword) |
| [`Get-SecurityAlerts.ps1`](#get-securityalertsps1) | Report unified Defender/Entra security alerts |
| [`Test-EmailSecurityPosture.ps1`](#test-emailsecuritypostureps1) | Consolidated Defender for O365 / anti-spam / DLP / mail flow security report |
| [`Test-MailboxAuditingConfig.ps1`](#test-mailboxauditingconfigps1) | Report and optionally fix unified audit log + per-mailbox auditing gaps |
| [`Test-EmailAuthenticationRecords.ps1`](#test-emailauthenticationrecordsps1) | Validate SPF and DMARC DNS records |

---

### Get-EntraAppConsents.ps1

Audits every enterprise application (service principal) with delegated permission
grants (`OAuth2PermissionGrants`) or application permission grants
(`AppRoleAssignments`), flagging grants that match a list of commonly-abused
high-privilege scopes. This is the classic "illicit consent grant" / third-party app
risk check.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-RiskyOnly` | No | Only include grants flagged High risk |
| `-OutputPath` | No | CSV report path (default: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | No | Entra ID tenant ID or domain |

**Examples**

```powershell
.\Get-EntraAppConsents.ps1

.\Get-EntraAppConsents.ps1 -RiskyOnly
```

**Notes**
- Risk flagging is a heuristic (string match against a known-risky scope list) — review
  flagged entries manually, don't treat "Normal" as a guarantee of safety
- Required scopes: `Application.Read.All`, `Directory.Read.All`

---

### Get-SuspiciousInboxRules.ps1

Scans every mailbox's inbox rules for patterns commonly left behind by a compromised
account: external forward/redirect, silent delete, or moving keyword-matching messages
(invoice, payment, wire, password, ...) into a rarely-checked folder. Default is a
read-only report; `-Apply` disables (not deletes) high-confidence matches.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-Mailbox` | No | UPN of a single mailbox. If omitted, all mailboxes are checked |
| `-Apply` | No | Disable high-confidence flagged rules (default: report only) |
| `-OutputPath` | No | CSV report path (default: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | No | Entra ID tenant ID or domain |

**Examples**

```powershell
.\Get-SuspiciousInboxRules.ps1

.\Get-SuspiciousInboxRules.ps1 -Mailbox "user@contoso.com"

# Preview what would be disabled
.\Get-SuspiciousInboxRules.ps1 -Apply -WhatIf

.\Get-SuspiciousInboxRules.ps1 -Apply
```

Supports `-WhatIf` (`SupportsShouldProcess`).

**Notes**
- "External" is determined against the tenant's accepted domains (`Get-AcceptedDomain`)
- Disabling (not deleting) is intentional — reversible, and preserves the rule for
  incident-response review

---

### Get-SecurityAlerts.ps1

Reports the unified Microsoft Graph security alerts feed (`security/alerts_v2`), which
spans Defender for Office 365, Defender for Endpoint, Defender for Identity, Defender for
Cloud Apps, and Entra ID Protection.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-Days` | No | Lookback window in days (default: `30`) |
| `-Severity` | No | Filter: `informational`, `low`, `medium`, `high` |
| `-Status` | No | Filter: `new`, `inProgress`, `resolved` |
| `-OutputPath` | No | CSV report path (default: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | No | Entra ID tenant ID or domain |

**Examples**

```powershell
.\Get-SecurityAlerts.ps1

.\Get-SecurityAlerts.ps1 -Days 7 -Severity high,medium -Status new,inProgress
```

**Notes**
- Required scope: `SecurityAlert.Read.All`

---

### Test-EmailSecurityPosture.ps1

Consolidated read-only report on Exchange Online / Defender for Office 365 email
security configuration: Safe Links, Safe Attachments, anti-malware, anti-spam
(inbound/outbound), connection filter, remote domains (external auto-forward), DLP
policies, and alert policy summary. Optionally (`-IncludeMailboxDetail`) also checks
POP/IMAP enablement and litigation hold per mailbox.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-IncludeMailboxDetail` | No | Also check per-mailbox legacy protocols + litigation hold (slower) |
| `-OutputPath` | No | CSV report path (default: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | No | Entra ID tenant ID or domain |

**Examples**

```powershell
.\Test-EmailSecurityPosture.ps1

.\Test-EmailSecurityPosture.ps1 -IncludeMailboxDetail
```

**Notes**
- This is the flagship consolidation of ~18 single-purpose upstream scripts — see the
  script's `.NOTES` for the full list
- The mutating `*-set.ps1` / `*-del.ps1` counterparts from the source project were
  intentionally not ported — each hardcoded one MSP's specific "recommended" values with
  no per-tenant override

**Required modules**
```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

---

### Test-MailboxAuditingConfig.ps1

Reports (and, with `-Apply`, fixes) unified audit log and per-mailbox audit logging
gaps: whether the org-wide Unified Audit Log is enabled, and whether each mailbox has
`AuditEnabled` on with a sufficient `AuditLogAgeLimit`.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-MinimumAuditLogAgeDays` | No | Minimum acceptable retention in days (default: `180`) |
| `-Apply` | No | Enable the unified audit log and fix flagged mailboxes (default: report only) |
| `-OutputPath` | No | CSV report path (default: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | No | Entra ID tenant ID or domain |

**Examples**

```powershell
.\Test-MailboxAuditingConfig.ps1

.\Test-MailboxAuditingConfig.ps1 -Apply -WhatIf

.\Test-MailboxAuditingConfig.ps1 -MinimumAuditLogAgeDays 365 -Apply
```

Supports `-WhatIf` (`SupportsShouldProcess`).

---

### Test-EmailAuthenticationRecords.ps1

Validates SPF and DMARC DNS records for one or more domains — checks SPF exists, includes
`spf.protection.outlook.com` where expected, and isn't permissively `+all`; checks DMARC
exists, reports its policy (`none`/`quarantine`/`reject`), and whether an aggregate report
address is configured. DKIM is intentionally out of scope — use
[`Test-DkimConfig.ps1`](../../Exchange/readme.md) for that.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-Domain` | No | One or more domains. If omitted, auto-discovered via Microsoft Graph |
| `-OutputPath` | No | CSV report path (default: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | No | Used only for auto-discovering domains via Graph |

**Examples**

```powershell
.\Test-EmailAuthenticationRecords.ps1 -Domain "contoso.com"

.\Test-EmailAuthenticationRecords.ps1
```

**Notes**
- Windows-only (uses `Resolve-DnsName`)
