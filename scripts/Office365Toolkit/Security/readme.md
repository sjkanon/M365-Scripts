# Office365Toolkit / Security

Security posture reporting and hardening scripts: Secure Score trend, enterprise
app consent review, shared mailbox sign-in lockdown, and an EOP anti-spam/
anti-malware baseline.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Get-SecureScoreReport.ps1`](#get-securescorereportps1) | Report Secure Score trend and weakest controls |
| [`Remove-EnterpriseAppConsent.ps1`](#remove-enterpriseappconsentps1) | Audit and optionally revoke an enterprise app's OAuth consent grants |
| [`Test-SharedMailboxSignIn.ps1`](#test-sharedmailboxsigninps1) | Report/disable direct sign-in on shared mailboxes |
| [`New-EOPProtectionBaseline.ps1`](#new-eopprotectionbaselineps1) | Create/update baseline anti-spam + anti-malware EOP policies |

---

### Get-SecureScoreReport.ps1

Reports the tenant's Microsoft Secure Score trend over time and breaks down the
latest snapshot's control scores, sorted weakest-first. Read-only.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-HistoryCount` | No | Number of historical snapshots to include (default `30`) |
| `-OutputPath` | No | CSV report path (default: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | No | Entra ID tenant ID or domain |

**Examples**

```powershell
.\Get-SecureScoreReport.ps1

.\Get-SecureScoreReport.ps1 -HistoryCount 90 -OutputPath C:\Reports
```

**Required scope:** `SecurityEvents.Read.All`
**Required module:** `Microsoft.Graph.Security`

---

### Remove-EnterpriseAppConsent.ps1

Audits (and, with `-Apply`, revokes) the delegated and application permission
grants held by a single enterprise application — useful for reviewing illicit
consent grants or cleaning up before removing an app. Exactly one of `-AppId` /
`-AppDisplayName` is required, so a single run can't accidentally sweep every
app in the tenant.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-AppId` | * | Application (client) ID or service principal Object ID |
| `-AppDisplayName` | * | Exact display name of the enterprise app |
| `-IncludeUserConsent` | No | Also report/revoke per-user delegated consent (not just tenant-wide admin consent) |
| `-Apply` | No | Actually revoke the reported grants (default: preview only) |
| `-OutputPath` | No | CSV report path |
| `-TenantId` | No | Entra ID tenant ID or domain |

*Exactly one of `-AppId` / `-AppDisplayName` is required.

**Examples**

```powershell
# Preview only
.\Remove-EnterpriseAppConsent.ps1 -AppDisplayName "Suspicious Reporting Tool"

# Revoke tenant-wide admin consent + application permissions
.\Remove-EnterpriseAppConsent.ps1 -AppDisplayName "Suspicious Reporting Tool" -Apply

# Also revoke individual users' delegated consent
.\Remove-EnterpriseAppConsent.ps1 -AppId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -IncludeUserConsent -Apply
```

Supports `-WhatIf` (`SupportsShouldProcess`).

**Required scopes:** `Application.Read.All`, `DelegatedPermissionGrant.ReadWrite.All`, `AppRoleAssignment.ReadWrite.All`
**Required module:** `Microsoft.Graph.Applications`

---

### Test-SharedMailboxSignIn.ps1

Cross-references Exchange Online shared mailboxes with their Entra ID account
state and reports which ones still allow direct interactive sign-in — a common
soft target since shared mailboxes are rarely licensed/MFA-protected. With
`-Apply`, disables sign-in (`AccountEnabled = $false`) for any found enabled.
Delegated access (Full Access / Send As) is unaffected.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-Mailbox` | No | UPN of a single shared mailbox. If omitted, all shared mailboxes are checked |
| `-Apply` | No | Actually disable sign-in for enabled accounts found (default: preview only) |
| `-OutputPath` | No | CSV report path |
| `-TenantId` | No | Entra ID tenant ID or domain |

**Examples**

```powershell
.\Test-SharedMailboxSignIn.ps1

.\Test-SharedMailboxSignIn.ps1 -Apply

.\Test-SharedMailboxSignIn.ps1 -Mailbox "helpdesk@contoso.com" -Apply
```

Supports `-WhatIf` (`SupportsShouldProcess`).

**Required scope:** `User.ReadWrite.All`
**Required modules:** `ExchangeOnlineManagement`, `Microsoft.Graph.Users`

---

### New-EOPProtectionBaseline.ps1

Reports the tenant's current anti-spam/anti-malware policies against a
recommended baseline, and with `-Apply` creates (or, with `-UpdateExisting`,
updates) a baseline hosted content filter (anti-spam) policy and/or malware
filter policy + rule, scoped to given recipient domains (defaults to all
accepted domains).

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-Domains` | No | Recipient domain(s) to scope the rule(s) to (default: all accepted domains) |
| `-Protection` | No | `Spam`, `Malware`, or `Both` (default) |
| `-SpamPolicyName` | No | Name for the anti-spam policy/rule (default `MSP Baseline Anti-Spam`) |
| `-MalwarePolicyName` | No | Name for the anti-malware policy/rule (default `MSP Baseline Anti-Malware`) |
| `-UpdateExisting` | No | Update the policy in place if one with the same name already exists |
| `-Apply` | No | Actually create/update the policies (default: preview only) |
| `-TenantId` | No | Entra ID tenant ID or domain |

**Examples**

```powershell
# Preview only
.\New-EOPProtectionBaseline.ps1

# Create both baseline policies for all accepted domains
.\New-EOPProtectionBaseline.ps1 -Apply

# Anti-spam only, specific domains, update in place if it exists
.\New-EOPProtectionBaseline.ps1 -Protection Spam -Domains "contoso.com" -UpdateExisting -Apply
```

**Notes**
- These are baseline recommendations, not a full hardening pass — review
  against your tenant's Standard/Strict preset security policies before
  applying broadly.

Supports `-WhatIf` (`SupportsShouldProcess`).

**Required module:** `ExchangeOnlineManagement`
