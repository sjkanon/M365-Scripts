# Exchange Scripts

Scripts for Exchange Online calendar, mailbox, and distribution group management.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Migrate-Calendar.ps1`](#migrate-calendarps1) | Migrate a shared M365 Group calendar to a Room Mailbox |
| [`Set-Calendar-rights.ps1`](#set-calendar-rightsps1) | Grant calendar folder permissions to a user |
| [`Set-Distributionlist-dynamic-static.ps1`](#set-distributionlist-dynamic-staticps1) | Resolve a dynamic distribution group's members into a regular (static) group |
| [`Move-InboxToArchive.ps1`](#move-inboxtoarchiveps1) | Move all (or date-filtered) Inbox messages of a mailbox to its Archive folder |
| [`Test-CalendarPermissions.ps1`](#test-calendarpermissionsps1) | Audit calendar folder permissions |
| [`Test-MailboxPermissions.ps1`](#test-mailboxpermissionsps1) | Audit Full Access, Send As, Send on Behalf delegation |
| [`Test-DistributionGroupPermissions.ps1`](#test-distributiongrouppermissionsps1) | Audit DG managers, Send As, Send on Behalf, member counts |
| [`Test-DkimConfig.ps1`](#test-dkimconfigps1) | Validate DKIM signing config and DNS records |
| [`Get-ExternalForwards.ps1`](#get-externalforwardsps1) | Audit mailboxes with external forwarding |
| [`Get-MailboxSizes.ps1`](#get-mailboxsizesps1) | Report mailbox sizes and item counts |
| [`Get-MessageTraceReport.ps1`](#get-messagetracereportps1) | Trace who received what, at what exact time, and where it was forwarded to |
| [`Remove-PhishingMessage.ps1`](#remove-phishingmessageps1) | Delete a phishing message from one, several, or all mailboxes — dry-run by default |

---

### Migrate-Calendar.ps1

Migrates a shared M365 Group calendar to a Room Mailbox. Solves the problem of group members receiving email notifications for every calendar event — a Room Mailbox uses the same booking mechanism as a meeting room: no notifications, auto-accept, visible to everyone.

**How it works**

1. Creates an Entra ID App Registration (or reuses an existing one)
2. Creates a Room Mailbox as the destination calendar
3. Configures AutoAccept and sets Default permissions to Reviewer
4. Reads events from the M365 Group calendar via delegated access
5. Copies events to the Room Mailbox via app auth
6. Optionally deletes the source M365 Group

> The script uses a dual-auth flow because Microsoft requires delegated access to read group calendars but application permissions to write to other mailboxes.

**Parameters**

| Parameter | Required | Default | Description |
|-----------|----------|---------|-------------|
| `-TenantId` | Yes | — | Entra ID Tenant ID |
| `-AdminUPN` | Yes | — | UPN of the executing admin (must be a member of the source group) |
| `-ClientId` | No | — | App Registration Client ID. If omitted, a new registration is created automatically |
| `-ClientSecret` | No | — | Client Secret. If omitted, created automatically |
| `-AppName` | No | `HolidaysCalendarMigration` | Name for the App Registration |
| `-SourceGroupMail` | No | — | Email address of the source M365 Group |
| `-SourceGroupDisplayName` | No | — | Display name of the source group (used as fallback lookup) |
| `-DestinationType` | No | `Room` | Destination mailbox type: `Room` or `Shared` |
| `-DestinationDisplayName` | No | `Holidays Calendar` | Display name of the destination mailbox |
| `-DestinationAlias` | No | `holidays-calendar` | Alias of the destination mailbox |
| `-DestinationEmail` | No | — | SMTP address of the destination mailbox |
| `-DaysBack` | No | `365` | Days back for event retrieval |
| `-DaysForward` | No | `730` | Days forward for event retrieval |
| `-DeleteSourceGroup` | No | `$false` | Delete the M365 Group after migration |

**Examples**

```powershell
# First run — create App Registration automatically
.\Migrate-Calendar.ps1 `
    -TenantId     "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -AdminUPN     "admin@contoso.com" `
    -SourceGroupMail "holidays@contoso.com"

# Subsequent runs — reuse existing App Registration
.\Migrate-Calendar.ps1 `
    -TenantId     "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -AdminUPN     "admin@contoso.com" `
    -ClientId     "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -ClientSecret "your-client-secret" `
    -SourceGroupMail "holidays@contoso.com"

# Dry run — no changes
.\Migrate-Calendar.ps1 `
    -TenantId     "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -AdminUPN     "admin@contoso.com" `
    -SourceGroupMail "holidays@contoso.com" `
    -WhatIf
```

**Required permissions**

| Permission | Purpose |
|-----------|---------|
| Exchange Admin or Global Admin | Create Room Mailbox |
| Global Admin | Create App Registration + grant admin consent |
| Member of the source M365 Group | Read group calendar via delegated access |

**Required modules**

```powershell
Install-Module ExchangeOnlineManagement     -Scope CurrentUser
Install-Module Microsoft.Graph.Applications  -Scope CurrentUser
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module Microsoft.Graph.Calendar      -Scope CurrentUser
Install-Module Microsoft.Graph.Groups        -Scope CurrentUser
Install-Module Microsoft.Graph.Users         -Scope CurrentUser
```

---

### Set-Calendar-rights.ps1

Grants a user access rights on another user's calendar folder in Exchange Online. Supports NL / FR / EN mailbox locales.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-User` | Yes | Username (without domain) receiving the permissions |
| `-TargetMailbox` | Yes | Username (without domain) of the target mailbox |
| `-AccessRights` | Yes | Access level (see table below) |

**Access levels**

| Value | Description |
|-------|-------------|
| `Owner` | Full control including deletion and folder management |
| `PublishingEditor` | Read, create, modify, delete and create subfolders |
| `Editor` | Read, create, modify and delete |
| `Author` | Read and create, modify/delete own items |
| `Reviewer` | Read-only |
| `AvailabilityOnly` | Free/busy only |
| `LimitedDetails` | Free/busy with limited details |

**Examples**

```powershell
# Grant Reviewer rights
.\Set-Calendar-rights.ps1 -User j.doe -TargetMailbox a.smith -AccessRights Reviewer

# Dry run
.\Set-Calendar-rights.ps1 -User j.doe -TargetMailbox a.smith -AccessRights Editor -WhatIf
```

**Required module**

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

---

### Set-Distributionlist-dynamic-static.ps1

Resolves the members currently matching a Dynamic Distribution Group's filter and copies them into a regular (static) distribution group — creating the target group if it doesn't exist. Also exports the resolved member list to CSV. Requires an active Exchange Online session (`Connect-ExchangeOnline`).

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-DynamicGroupIdentity` | Yes | Source dynamic distribution group (name, alias, DN, or SMTP address) |
| `-TargetGroupIdentity` | Yes | Target regular distribution group — created if it doesn't exist |
| `-TargetDisplayName` | No | Display name for a new target group (default: `<DynamicDisplayName> Static`) |
| `-TargetAlias` | No | Alias for a new target group (default: `<DynamicAlias>-static`) |
| `-TargetPrimarySmtpAddress` | No | SMTP address for a new target group (default: dynamic group's current primary SMTP) |
| `-CopyManagersFromDynamic` | No | Copy `ManagedBy` owners from the dynamic group to the target group (default: on) |
| `-DisableCopyManagersFromDynamic` | No | Disable copying `ManagedBy` owners |
| `-MakeDynamicAddressTemporary` | No | Give the dynamic group a temporary primary SMTP first, freeing its address for the target group (default: on) |
| `-DisableMakeDynamicAddressTemporary` | No | Disable the automatic temporary SMTP change |
| `-ClearTargetMembers` | No | Remove existing target members before adding the resolved dynamic members |
| `-ExportCsvPath` | No | CSV export path for resolved members (default: `C:\Temp\DynamicGroupMembers_<timestamp>.csv` on Windows, `~/Downloads` on Linux/macOS) |
| `-SkipMemberAdd` | No | Only resolve and export members, do not modify the target group |
| `-RenameDynamicGroupTo` | No | Rename the source dynamic distribution group after processing |

**Examples**

```powershell
# Resolve dynamic group and populate regular group
.\Set-Distributionlist-dynamic-static.ps1 -DynamicGroupIdentity "All Sales" -TargetGroupIdentity "All Sales Static"

# Full refresh of target group members
.\Set-Distributionlist-dynamic-static.ps1 -DynamicGroupIdentity sales@contoso.com -TargetGroupIdentity sales-static@contoso.com -ClearTargetMembers

# Dry run
.\Set-Distributionlist-dynamic-static.ps1 -DynamicGroupIdentity "All Staff" -TargetGroupIdentity "All Staff Static" -WhatIf

# Convert and rename the original dynamic group
.\Set-Distributionlist-dynamic-static.ps1 -DynamicGroupIdentity "All Sales" -TargetGroupIdentity "All Sales Static" -RenameDynamicGroupTo "All Sales (Legacy Dynamic)"

# Free up the dynamic group's SMTP address for the new static group
.\Set-Distributionlist-dynamic-static.ps1 -DynamicGroupIdentity "All Sales" -TargetGroupIdentity "All Sales Static" -MakeDynamicAddressTemporary
```

**Notes**
- Requires Exchange Online PowerShell module and active EXO session (`Connect-ExchangeOnline`)
- Dynamic Distribution Groups are Exchange objects; this script uses Exchange cmdlets, not Graph

---

### Move-InboxToArchive.ps1

Moves every message in a mailbox's Inbox to its Archive folder — the same folder Outlook's "Archive" button targets. Optionally restrict the scope to a date range (`-After` / `-Before`). Uses Microsoft Graph's `$batch` endpoint to move messages in batches of 20, with retry/backoff on throttling (429/503). Default behavior is safe preview mode — pass `-Apply` to actually move messages.

**Authentication (default: automatic, no Full Access needed)**

By default the script archives any mailbox in the tenant without requiring Full Access on it. It connects interactively (delegated, `Application.ReadWrite.All` + `AppRoleAssignment.ReadWrite.All`), creates a short-lived temporary App Registration, self-grants it `Mail.ReadWrite` application permission (no separate admin-consent screen — the delegated role does that), uses it for the mailbox operations, and removes it again when the script finishes. This mirrors the temporary-app pattern in `Get-SharePointStorageReport.ps1` / `Remove-SharePointFileVersionsByDate.ps1`. Requires Global Administrator or Privileged Role Administrator for that one-time setup, and the `Microsoft.Graph.Applications` module.

- `-Delegated` skips all of that and uses a plain delegated `Mail.ReadWrite` session instead — needs Exchange Admin, not Entra app-creation rights. For a mailbox other than the signed-in user's own, the script connects to Exchange Online, grants that account temporary Full Access, polls `Get-MailboxPermission` until it's actually visible (up to ~3 minutes — Exchange Online permission changes don't propagate instantly), archives, then removes the grant again (with a few retries, since the removal can likewise hit a domain controller that hasn't caught up yet).
  > **Known limitation:** `Get-MailboxPermission` reflects Exchange's own state almost immediately, but Microsoft Graph's authorization cache for delegate mailbox access can lag up to **~60 minutes** behind that — this is a Microsoft-side limitation. If the actual Inbox read still 403s after the Full Access poll, the script keeps retrying it (60s apart) against a **`-MaxWaitMinutes`** deadline (default 65, covering Microsoft's documented worst case) — the Full Access grant stays in place for the whole wait, since revoking and re-granting between attempts would reset the propagation clock. Raise `-MaxWaitMinutes` if 65 isn't enough, or drop `-Delegated` to use the default app-only mode, which has no such delay.
- `-ClientId` + `-ClientSecret`/`-CertificateThumbprint` reuses your own existing App Registration instead of creating a temporary one — that app must already have `Mail.ReadWrite` application permission (admin consent granted).

**Parameters**

| Parameter | Required | Description |
|-----------|----------|--------------|
| `-Mailbox` | Yes | UPN or object ID of the mailbox whose Inbox to archive |
| `-After` | No | Only archive messages received on or after this date |
| `-Before` | No | Only archive messages received before this date |
| `-TenantId` | No | Entra ID tenant ID (GUID) **or** a verified domain of the tenant (e.g. `contoso.com`) — either works. Optional if already connected or resolvable from a GDAP customer tenant context; required for app-only auth if not resolvable |
| `-ClientId` | No | Existing App Registration client ID for app-only auth — skips the automatic temporary app. Use with `-TenantId` and `-ClientSecret` or `-CertificateThumbprint` |
| `-ClientSecret` | No | Client secret for the app registration in `-ClientId` |
| `-CertificateThumbprint` | No | Certificate thumbprint for the app registration in `-ClientId` |
| `-Delegated` | No | Skip the automatic temporary app-only setup; connect delegated instead. For other mailboxes, auto-grants + polls + revokes temporary Full Access via Exchange Online (needs Exchange Admin) |
| `-MaxWaitMinutes` | No | `-Delegated` only. How long to keep retrying while waiting for Graph to honor the Full Access grant, before giving up and revoking it. Default `65` |
| `-Apply` | No | Actually move the messages. Without it, the script only reports how many messages would be archived |

**Examples**

```powershell
# Preview — auto app-only setup, reports the count, makes no changes
.\Move-InboxToArchive.ps1 -Mailbox "user@contoso.com"

# Archive everything in the Inbox
.\Move-InboxToArchive.ps1 -Mailbox "user@contoso.com" -Apply

# Only messages received before 2025
.\Move-InboxToArchive.ps1 -Mailbox "user@contoso.com" -Before (Get-Date "2025-01-01") -Apply

# Only messages received in 2024
.\Move-InboxToArchive.ps1 -Mailbox "user@contoso.com" -After (Get-Date "2024-01-01") -Before (Get-Date "2025-01-01") -Apply

# Delegated — auto-grants + polls + revokes temporary Full Access via Exchange Online
# instead of the Entra app-only setup (needs Exchange Admin, not Global Admin)
.\Move-InboxToArchive.ps1 -Mailbox "user@contoso.com" -Delegated -Apply

# Delegated, willing to wait out Microsoft's full ~90-minute worst case for Graph
# to honor the Full Access grant, instead of the 65-minute default
.\Move-InboxToArchive.ps1 -Mailbox "user@contoso.com" -Delegated -MaxWaitMinutes 90 -Apply

# Reuse an existing App Registration instead of creating a temporary one.
# -TenantId accepts the tenant's domain instead of its GUID.
.\Move-InboxToArchive.ps1 -Mailbox "user@contoso.com" -TenantId "contoso.com" `
    -ClientId "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" -ClientSecret "your-client-secret" -Apply
```

**Notes**
- Mailbox reads/moves always go through Microsoft Graph, not Exchange Online cmdlets — requires `Microsoft.Graph.Authentication` (and `Microsoft.Graph.Applications` for the default automatic temporary-app mode, or `ExchangeOnlineManagement` for `-Delegated`'s temporary Full Access grant)
- Prints timestamped progress while paginating Inbox messages, while moving batches (`[HH:mm:ss] N / total moved (...%)`), and while polling for Full Access propagation in `-Delegated` mode
- GDAP-aware: under a GDAP session (`$global:authMode -eq 'GDAP'`, set via `Connect-Tenant` / `load.ps1`), `-TenantId` is resolved automatically from the selected customer tenant (`$global:cid`) if omitted — same fallback as `Get-SharePointStorageReport.ps1` / `Remove-SharePointFileVersionsByDate.ps1`. `$env:M365_CUSTOMER_TENANTID` / `$env:M365_AUTH_MODE` are honored too

**Required modules**

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module Microsoft.Graph.Applications    -Scope CurrentUser
Install-Module ExchangeOnlineManagement        -Scope CurrentUser
```

---

## Audit scripts

Connect to Exchange Online automatically if no session is active; reuse an existing session if already connected.

---

### Test-CalendarPermissions.ps1

Retrieves calendar folder permissions for one or all mailboxes. Uses `FolderType` to locate the calendar folder independently of the mailbox locale (NL/FR/EN). Exports results to CSV.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-Mailbox` | No | UPN of a single mailbox. If omitted, all user and shared mailboxes are checked |
| `-OutputPath` | No | CSV report path (default: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | No | Entra ID tenant ID or domain |

**Examples**

```powershell
# Audit all mailboxes
.\Test-CalendarPermissions.ps1

# Single mailbox
.\Test-CalendarPermissions.ps1 -Mailbox "user@contoso.com"

# Custom output path
.\Test-CalendarPermissions.ps1 -OutputPath "C:\Reports\calendar.csv"
```

---

### Test-MailboxPermissions.ps1

Audits all three delegation types for one or all mailboxes:

- **Full Access** — users who can open the mailbox
- **Send As** — users who can send as the mailbox identity
- **Send on Behalf** — users listed in `GrantSendOnBehalfTo`

Inherited and SELF entries are filtered out automatically.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-Mailbox` | No | UPN of a single mailbox. If omitted, all user and shared mailboxes are checked |
| `-OutputPath` | No | CSV report path (default: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | No | Entra ID tenant ID or domain |

**Examples**

```powershell
# Audit all mailboxes
.\Test-MailboxPermissions.ps1

# Shared mailbox only
.\Test-MailboxPermissions.ps1 -Mailbox "shared@contoso.com"
```

---

### Test-DistributionGroupPermissions.ps1

Audits distribution groups and mail-enabled security groups:

- **Settings** — member count, join/leave restrictions, external sender policy
- **ManagedBy** — group owners/managers
- **Send As** — who can send as the group
- **Send on Behalf** — `GrantSendOnBehalfTo` delegates
- **Members** (optional, use `-IncludeMembers`)

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-Group` | No | Name, alias, or email of a single group. If omitted, all DGs are audited |
| `-IncludeMembers` | No | Also list individual group members in the report |
| `-OutputPath` | No | CSV report path (default: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | No | Entra ID tenant ID or domain |

**Examples**

```powershell
# Audit all distribution groups
.\Test-DistributionGroupPermissions.ps1

# Single group with member list
.\Test-DistributionGroupPermissions.ps1 -Group "helpdesk@contoso.com" -IncludeMembers
```

**Required module**
```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

---

### Test-DkimConfig.ps1

Validates the DKIM signing configuration for one or all accepted domains:

- Checks whether DKIM signing is enabled
- Resolves `selector1/2._domainkey.<domain>` CNAME records and compares against Exchange config
- Resolves Microsoft's TXT public key records and verifies the key matches
- Lists required actions for any issues found

> DNS lookups use `Resolve-DnsName` (Windows only). On macOS/Linux the Exchange config is shown without DNS validation.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-Domain` | No | Domain to validate. If omitted, all domains with a signing config are checked |
| `-ShowAll` | No | Show full signing config object instead of the summarised view |
| `-TenantId` | No | Entra ID tenant ID or domain |

**Examples**

```powershell
# Check all domains
.\Test-DkimConfig.ps1

# Single domain
.\Test-DkimConfig.ps1 -Domain "contoso.com"
```

---

### Get-ExternalForwards.ps1

Audits all mailboxes for forwarding rules that point to external (non-tenant) domains. External forwarding is a common security/compliance risk and should be reviewed regularly. Exports to CSV.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-Mailbox` | No | UPN of a single mailbox. If omitted, all mailboxes are checked |
| `-OutputPath` | No | CSV report path (default: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | No | Entra ID tenant ID or domain |

**Examples**

```powershell
# Check all mailboxes
.\Get-ExternalForwards.ps1

# Single mailbox
.\Get-ExternalForwards.ps1 -Mailbox "user@contoso.com"
```

---

### Get-MailboxSizes.ps1

Reports mailbox sizes (MB/GB), item counts, and quota status. Sorted by size descending. Exports to CSV.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-Mailbox` | No | UPN of a single mailbox. If omitted, all user and shared mailboxes are reported |
| `-OutputPath` | No | CSV report path (default: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | No | Entra ID tenant ID or domain |

**Examples**

```powershell
# Report all mailboxes
.\Get-MailboxSizes.ps1

# Single mailbox
.\Get-MailboxSizes.ps1 -Mailbox "user@contoso.com"
```

---

### Get-MessageTraceReport.ps1

Answers "who received this, when exactly, and where did it go next?". Runs a message trace and reports per message the exact timestamp in **both local time and UTC** (Exchange stores message trace timestamps in UTC), sender, recipient, subject, status, size, originating/delivering IP, `MessageId` and `MessageTraceId`.

**Forward detection** — the `ForwardedTo` column is filled from three independent signals, and `ForwardDetection` names which one fired:

| Method | Detects |
|--------|---------|
| `SameMessageId` | Other recipients that received the same `MessageId` — SMTP forwarding, redirect rules, distribution group expansion |
| `RedirectHop` | Redirect / transport-rule hops pulled from `Get-MessageTraceDetailV2` (requires `-IncludeDetails`) |
| `ClientForward(subject match)` | A later message sent **by** the recipient carrying the same normalized subject — an Outlook "Forward", which gets a brand new `MessageId`. Heuristic; replies back to the original sender are excluded |

> **Why the forward target is traced separately:** a mailbox forward or a redirect rule keeps the **original sender** on the forwarded copy. The trace row for the delivery to the forward target therefore mentions the traced mailbox *neither as sender nor as recipient* — filtering on the mailbox alone would never return it. The script resolves the configured forward targets **before** tracing and adds them as extra recipient filters, so the actual hand-off shows up with its own exact timestamp. `-ResolveSiblings` goes further and re-traces every matched `MessageId` without any filter, which also catches forward targets that are no longer configured (a rule deleted after it did its work still leaves its deliveries in the trace).

On top of that, the script reports the **configured** forwarding of every internal mailbox that appears in the trace — `ForwardingSMTPAddress` / `ForwardingAddress` plus any inbox rule with `ForwardTo` / `RedirectTo` / `ForwardAsAttachmentTo` — so a forward that has not fired inside the traced window is still visible.

Uses `Get-MessageTraceV2` when available and falls back to the retired `Get-MessageTrace`. Ranges longer than the V2 limit are split into 10-day chunks automatically, and every chunk is paginated until exhausted.

**Parameters**

| Parameter | Required | Default | Description |
|-----------|----------|---------|-------------|
| `-Mailbox` | No | — | Trace both directions for this address (sent **and** received) and pull its forwarding config |
| `-SenderAddress` | No | — | Filter on sender address. Aliased as `-Sender` (`$Sender` is a PowerShell automatic variable, so it cannot be the parameter's real name) |
| `-Recipient` | No | — | Filter on recipient address |
| `-ForwardAddress` | No | — | Known forward/exfiltration address(es) to trace as recipients on top of whatever forwarding config is discovered. Use when the forward has **already been removed** — there is then no config left to find, but its past deliveries are still in the trace |
| `-Subject` | No | — | Client-side subject filter, wildcards allowed (message trace cannot filter on subject server-side) |
| `-MessageId` | No | — | Internet MessageId to trace, with or without angle brackets |
| `-Days` | No | `2` | Days back from `-EndDate`. Ignored if `-StartDate` is given |
| `-StartDate` | No | — | Explicit window start (local time) |
| `-EndDate` | No | now | Explicit window end (local time) |
| `-Status` | No | — | `Delivered`, `Failed`, `Pending`, `Expanded`, `Quarantined`, `FilteredAsSpam`, `GettingStatus`, `None` |
| `-IncludeDetails` | No | off | Retrieve per-hop delivery detail — this is what exposes redirect / transport-rule targets. Slow and throttled |
| `-MaxDetailLookups` | No | `50` | Cap on hop-detail lookups; truncation is reported explicitly |
| `-ResolveSiblings` | No | off | Re-trace every matched `MessageId` without sender/recipient filter to reveal **all** recipients — catches forward targets that are no longer configured. One extra call per MessageId |
| `-MaxSiblingLookups` | No | `100` | Cap on sibling lookups |
| `-SkipForwardingConfig` | No | off | Skip the mailbox forwarding / inbox rule inspection |
| `-OutputPath` | No | `C:\Temp\` / `~/Downloads` | Main CSV path. Detail and forwarding reports are written alongside it with `_Details` / `_ForwardingConfig` suffixes |
| `-TenantId` | No | — | Entra ID tenant ID or domain |

**Examples**

```powershell
# Everything one mailbox sent and received in the last 2 days, incl. forwards
.\Get-MessageTraceReport.ps1 -Mailbox "user@contoso.com"

# One specific flow over the last 30 days, with per-hop detail
.\Get-MessageTraceReport.ps1 -Sender "boss@contoso.com" -Recipient "user@contoso.com" -Days 30 -IncludeDetails

# Where did this specific message end up?
.\Get-MessageTraceReport.ps1 -MessageId "<abc123@contoso.com>" -Days 10 -IncludeDetails

# Suspected external forward — full picture, incl. targets no longer configured
.\Get-MessageTraceReport.ps1 -Mailbox "facturen@contoso.com" -Days 10 -ResolveSiblings -IncludeDetails

# The forward was already removed, but the address is known — trace it anyway
.\Get-MessageTraceReport.ps1 -Mailbox "facturen@contoso.com" -Days 10 -ResolveSiblings `
    -ForwardAddress "exfil@lookalike-domain.nl"

# All failed mail from one sender in an explicit window
.\Get-MessageTraceReport.ps1 -Sender "noreply@contoso.com" -Status Failed `
    -StartDate (Get-Date "2026-08-01") -EndDate (Get-Date "2026-08-08")
```

**Notes**
- Message trace retains **90 days**; the script warns when the requested window reaches past that
- Reading inbox rules requires permissions on the mailbox — mailboxes that cannot be read are skipped silently (use `-Verbose` to see which)
- `-IncludeDetails` issues one API call per message and is subject to Exchange Online throttling; raise `-MaxDetailLookups` deliberately

**Required module**

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

---

### Remove-PhishingMessage.ps1

Incident-response companion to `Get-MessageTraceReport.ps1`: the trace tells you **who received** the phish, this script **removes it again**. Dry-run by default — nothing is deleted without `-Apply`.

**Two engines**

| Engine | How it finds messages | Use it when |
|--------|----------------------|-------------|
| `Purview` | One KQL Content Search across the tenant, then `New-ComplianceSearchAction -Purge` | You do **not** know the recipients, or you need a **HardDelete**. Reports counts **per mailbox**, not individual messages |
| `Graph` | Enumerates each target mailbox over the Graph mail API and deletes message by message | You **do** know the recipients (from the trace) and want it gone **now**, with a per-message report |

Engine defaults to `Graph` when `-Mailbox` is given and `Purview` otherwise. Override with `-Engine`.

> **Why two engines.** Purview reads the **search index**, which lags delivery by roughly 15–30 minutes — a purge fired straight after the phish lands can honestly report *0 hits* and still leave the message sitting in every inbox. Graph queries the mailbox directly and has no such lag, but it needs the recipient list and cannot write to `Recoverable Items\Purges`, so it cannot hard-delete. During a live campaign the usual sequence is: trace → **Graph** the known recipients immediately → **Purview** sweep tenant-wide half an hour later to catch the rest.

**Delete types**

| Value | Lands in | User can recover? | Engines |
|-------|----------|-------------------|---------|
| `Recycle` | Deleted Items | Yes, trivially | Graph |
| `SoftDelete` *(default)* | `Recoverable Items\Deletions` | Yes, via "Recover deleted items" | Both |
| `HardDelete` | `Recoverable Items\Purges` | No — retained only if the mailbox is on hold | Purview |

**Parameters**

| Parameter | Required | Default | Description |
|-----------|----------|---------|-------------|
| `-Mailbox` | No | — | Target mailbox address(es). Required for Graph unless `-AllMailboxes`. Omitted on Purview = every mailbox in the tenant |
| `-AllMailboxes` | No | off | Sweep every mailbox. Implicit for Purview; for Graph this is one query per mailbox and is slow |
| `-MessageId` | No | — | Internet MessageId, with or without angle brackets. **The precise selector** — matches that one message and nothing else |
| `-SenderAddress` | No | — | Sender address. Aliased as `-Sender` (`$Sender` is a PowerShell automatic variable) |
| `-Subject` | No | — | Purview matches this as an indexed phrase; Graph matches client-side and accepts wildcards |
| `-AttachmentName` | No | — | Attachment filename, wildcards allowed (e.g. `*.html`) |
| `-BodyContains` | No | — | Word or phrase in the body. **Purview only** — Graph would have to download every body |
| `-ReceivedAfter` | No | — | Only messages received at or after this moment (local time) |
| `-ReceivedBefore` | No | — | Only messages received at or before this moment (local time) |
| `-Engine` | No | see above | `Purview` or `Graph` |
| `-DeleteType` | No | `SoftDelete` | `Recycle`, `SoftDelete` or `HardDelete` (see table above) |
| `-Apply` | No | off | **Actually delete.** Without it the script only reports what it found |
| `-SearchName` | No | `Phish_<timestamp>` | Name of the Content Search to create. Purview requires unique names |
| `-KeepSearch` | No | off | Keep the Content Search afterwards so you can inspect it in the Purview portal |
| `-IncludeCalendar` | No | off | Also remove matching **calendar items**, not just mail. Works on **both engines**; needs `-Subject` or `-SenderAddress` |
| `-CalendarDaysBack` | No | `30` | How far back to scan the calendar |
| `-CalendarDaysForward` | No | `365` | How far forward to scan the calendar |
| `-VerifyWithGraph` | No | off | After a Purview purge, check the affected mailboxes over Graph to confirm the messages are really gone. Needs the same app-only Graph session as `-Engine Graph` |
| `-MaxPurgeRounds` | No | `10` | Purview purges max 10 items per mailbox per action, so the script loops rounds. 10 rounds = up to 100 items per mailbox |
| `-MaxMessagesPerMailbox` | No | `500` | Graph safety cap per mailbox; hitting it is reported explicitly |
| `-TimeoutMinutes` | No | `30` | How long to wait for a search or purge action to complete |
| `-OutputPath` | No | `C:\Temp\` / `~/Downloads` | CSV report path |
| `-TenantId` | No | — | Tenant ID or domain, used when the script has to connect itself |
| `-ClientId` | No | — | Your own App Registration for app-only Graph auth — skips the automatic temporary app |
| `-ClientSecret` | No | — | Client secret for `-ClientId` |
| `-CertificateThumbprint` | No | — | Certificate thumbprint for `-ClientId` |

**Examples**

```powershell
# 1. What would be removed, tenant-wide? (no -Apply = nothing is deleted)
.\Remove-PhishingMessage.ps1 -MessageId "<abc123@evil.example>"

# 1b. Purge, then confirm over Graph that it is really gone
.\Remove-PhishingMessage.ps1 -MessageId "<abc123@evil.example>" `
    -DeleteType HardDelete -Apply -VerifyWithGraph

# 2. Same, now actually purge it beyond user recovery
.\Remove-PhishingMessage.ps1 -MessageId "<abc123@evil.example>" `
    -DeleteType HardDelete -Apply

# 3. Known recipients from the message trace — immediate, no index lag
.\Remove-PhishingMessage.ps1 `
    -Mailbox "a@contoso.com","b@contoso.com" `
    -Sender  "no-reply@evil.example" `
    -Subject "*password expires*" `
    -Apply

# 4. Campaign sweep: everything from one sender in a window, tenant-wide
.\Remove-PhishingMessage.ps1 -Sender "no-reply@evil.example" `
    -ReceivedAfter (Get-Date "2026-08-30") -DeleteType HardDelete -Apply

# 5. Phishing MEETING INVITE, tenant-wide: hard-delete the mail AND the events
.\Remove-PhishingMessage.ps1 -Sender "no-reply@evil.example" -Subject "kick-off" `
    -DeleteType HardDelete -IncludeCalendar -Apply -VerifyWithGraph

# 6. HTML attachment campaign
.\Remove-PhishingMessage.ps1 -AttachmentName "*.html" `
    -Sender "billing@evil.example" -Apply
```

**Typical incident flow**

```powershell
# Who got it, and where did it go?
.\Get-MessageTraceReport.ps1 -Sender "no-reply@evil.example" -Days 2 -ResolveSiblings

# Pull it from the known recipients right now (no index lag)
.\Remove-PhishingMessage.ps1 -Mailbox $recipients -MessageId "<abc123@evil.example>" -Apply

# ~30 minutes later, sweep the tenant for anything the trace missed
.\Remove-PhishingMessage.ps1 -MessageId "<abc123@evil.example>" -DeleteType HardDelete -Apply
```

**Required permissions**

| Engine | Permission |
|--------|-----------|
| `Purview` | Membership of the **Search And Purge** role — in practice the *Organization Management* or *eDiscovery Manager* role group in the Purview compliance portal. Connects via `Connect-IPPSSession -EnableSearchOnlySession` |
| `Graph` | App-only `Mail.ReadWrite`. **You do not have to arrange this yourself** — see the three routes below |

**How the Graph engine (and `-VerifyWithGraph`) gets its access**

The same three-way pattern as [`Move-InboxToArchive.ps1`](#move-inboxtoarchiveps1) and the SharePoint reporting scripts, tried in order:

| # | Route | What it needs |
|---|-------|---------------|
| 1 | An app-only Graph session you already established | Nothing — it is used as-is |
| 2 | `-ClientId` + `-TenantId` + (`-ClientSecret` or `-CertificateThumbprint`) | Your own app with `Mail.ReadWrite` application permission, admin consent granted. **With `-ClientSecret` this is the most robust route** — it takes its token over plain REST and never loads the Graph SDK |
| 3 | **Automatic** — device code sign-in, then a short-lived App Registration that self-grants `Mail.ReadWrite`, hands over an app-only token, and is **removed again when the run finishes** | Global Administrator or Privileged Role Administrator for that one-time sign-in. No extra modules |

Route 3 is what happens when you pass nothing, so `-VerifyWithGraph` works out of the box. The delegated role grants the consent, so there is no separate admin-consent screen. If setup fails halfway, the partly-created app is removed before the error is reported — no orphans left in Entra ID.

Routes 2 (with `-ClientSecret`) and 3 are both built on plain REST — device code flow for the sign-in, the Graph REST API for creating and deleting the app registration. **Neither loads the Graph SDK**, which is what lets them work in the same session that already connected to Exchange. Route 3 shows a code to enter at `microsoft.com/devicelogin`:

```
  ------------------------------------------------------------
   To sign in, use a web browser to open https://microsoft.com/devicelogin
   and enter the code ABCD-EFGH to authenticate.
  ------------------------------------------------------------
```

> **Exchange and Graph fight over MSAL.** `ExchangeOnlineManagement` and `Microsoft.Graph.Authentication` each bundle their own `Microsoft.Identity.Client`, and .NET loads only the first one a process touches. So a Purview purge (which connects Exchange) followed by `-VerifyWithGraph` in the same window makes the Graph SDK call into an MSAL whose API does not match, and it fails with `Method not found: ... WithLogging(...)` — which looks nothing like the version clash it is.
>
> Routes 2 (`-ClientSecret`) and 3 sidestep this entirely by never loading the SDK. Only the `-CertificateThumbprint` variant and reusing an existing `Connect-MgGraph` session still go through it, and both report the clash for what it is rather than leaving you to read the stack trace.
>
> When `-VerifyWithGraph` is used with the Purview engine, Graph access is established **before** the purge, so a verification that cannot run is reported up front instead of after the messages are gone. The purge still runs either way — a failed verification never means a failed purge.

> Delegated `Mail.ReadWrite` only ever reaches *your own* mailbox, so a delegated session is deliberately **not** accepted for the Graph engine; the script falls through to route 2 or 3 instead. Note that `Mail.ReadWrite` (application) grants access to **every** mailbox in the tenant; scope the app with `New-ApplicationAccessPolicy` if that is wider than you want.

> GDAP-aware: under a GDAP session (`$global:authMode -eq 'GDAP'`, set by `Connect-Tenant` / `load.ps1`) `-TenantId` is resolved from the selected customer tenant, same as the SharePoint scripts.

**Notes**
- The calendar needs `Calendars.ReadWrite`, which `Mail.ReadWrite` does not cover. The automatic temporary app grants it **only when `-IncludeCalendar` is used**, and the script checks the roles claim in the issued token before doing any work — an app-only token is handed out whether or not the grant has landed, and a token minted too early is cached for an hour, which otherwise turns into a 403 on every single mailbox for the whole run
- **A phishing meeting invite is only half gone when the mail is deleted.** The invitation leaves an event in the calendar, and `/messages` and `/events` are separate collections — neither engine touches the calendar by default. `-IncludeCalendar` sweeps those too, matching on `-Subject` or `-SenderAddress` (as organiser). Without a selector it refuses rather than walk the whole calendar
- **`-IncludeCalendar` works with the Purview engine too**, and that pairing is the full clean-up of a meeting-invite phish: Purview hard-deletes the invitation across the tenant, then a Graph pass removes the events it left behind in exactly the mailboxes the search hit. The calendar pass needs the same Graph access as `-VerifyWithGraph`, and a mailbox it cannot reach is reported rather than counted as clean
- Verification follows suit: with `-IncludeCalendar` it checks the calendar as well, and without it prints *"mail only — calendar items are not checked"* rather than reporting a mailbox clean on incomplete evidence
- **Nothing in Purview can confirm a purge.** The purge action reports what the service believes it did, and the search index keeps listing purged items for up to ~30 minutes — so re-running the script is not a check. `-VerifyWithGraph` is the only lag-free verification: it re-asks the *same* query the Graph engine deletes on, directly against the mailboxes the search hit. Soft- and hard-deleted items sit in Recoverable Items, which Graph does not list, so a purged message correctly reads as gone
- Verification distinguishes **"could not check"** from **"clean"**. A mailbox that returns 403 is reported as unverified, never as confirmed. It also never fails the run — a purge that already happened is not reported as failed because the check could not run
- **Purview cannot show you the individual messages.** Content Search reports item counts per mailbox; the preview action that used to return sender and subject per message is [documented as on-premises only](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/new-compliancesearchaction?view=exchange-ps) since the May 2025 eDiscovery changes. For per-message detail, take the mailbox list from the Purview run and re-run those addresses through `-Engine Graph`
- Purview runs searches and purges server-side and they routinely take minutes. The script reports the job status and elapsed time roughly every 15 seconds while it waits, so a slow step is visibly slow rather than looking hung, and `-TimeoutMinutes` (default 30) bounds it
- **Rounds are planned, not polled.** A purge removes at most 10 items per mailbox per action, so the script computes `ceil(max items per mailbox / 10)` from the first search. It deliberately does *not* loop until the index goes quiet: the index lags a purge by up to ~30 minutes, so that would re-purge the same items and then report a false truncation. What each round actually removed is read back from the purge action itself
- A single content search purges at most **50,000 mailboxes**; beyond that the script warns and you should batch with `-Mailbox`. Microsoft points at the Graph `ediscoverySearch: purgeData` API (100 items per location) for bulk work
- **Requires ExchangeOnlineManagement 3.9.0+** for the Purview engine. Content Search runs on a backend that a plain IPPS connection no longer reaches: without `-EnableSearchOnlySession` the cmdlets are present but `Start-ComplianceSearch` fails at initialisation. The script passes the switch when it connects itself. **If you were already connected without it, the session cannot be repaired from inside the process** — open a new PowerShell window and let the script connect
- **You do not need the whole subject.** `-Subject` matches a fragment: on Purview it becomes a KQL phrase, which matches anywhere in the subject, so `-Subject "kick-off meeting"` finds every mail containing those words in that order. A long, punctuation-heavy subject is in fact the *fragile* choice — commas, apostrophes and times like `14:09` word-break badly. Short and distinctive wins
- **Matching inside a word** is the one thing KQL cannot do. A leading `*` is dropped (the script warns rather than silently ignoring it) and a trailing `*` only survives on a single word, since a wildcard is inert inside a quoted phrase. For true substring matching use `-Engine Graph` with `-Mailbox`, which matches client-side and takes wildcards as written
- **Selectors combine with AND.** The usual phishing pair is the sender plus a subject fragment: `-Sender "no-reply@evil.example" -Subject "kick-off meeting"`. When the sender rotates addresses, `-BodyContains "a distinctive sentence"` (Purview only) is often the most durable selector
- A run that fails partway no longer leaves its Content Search behind — cleanup runs in a `finally`. Searches orphaned by older runs can be listed with `Get-ComplianceSearch | Where-Object Name -like 'Phish_*'` and removed with `Remove-ComplianceSearch`
- The script **refuses to run** without at least one of `-MessageId`, `-SenderAddress`, `-Subject`, `-AttachmentName` or `-BodyContains` — a date range on its own would match every message in every mailbox
- **Index lag** (Purview only): a message delivered in the last ~30 minutes may not be searchable yet. A `0 hits` result straight after delivery is not proof the phish is gone — wait and re-run, or use the Graph engine
- Purge covers the **primary mailbox only** — neither engine reaches the archive mailbox
- KQL has no wildcard-in-phrase support, so `-Subject "*invoice*"` has its wildcards stripped on the Purview engine and is matched as a phrase; on Graph the wildcards work as written
- Every run writes a CSV report of what was matched and what was deleted

**Required modules**

```powershell
Install-Module ExchangeOnlineManagement       -Scope CurrentUser
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser   # optional, see below
```

Microsoft.Graph.Authentication is only needed to reuse an existing `Connect-MgGraph` session or to use `-CertificateThumbprint`. The `-ClientSecret` and automatic temporary-app routes run on plain REST and need nothing beyond ExchangeOnlineManagement.
