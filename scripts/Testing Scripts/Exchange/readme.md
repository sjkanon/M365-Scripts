# Testing — Exchange

Audit and diagnostic scripts for Exchange Online.
All scripts connect automatically if no session is active and reuse an existing session if already connected.

---

## Scripts

### Test-CalendarPermissions.ps1

Retrieves calendar folder permissions for one or all mailboxes. Uses `FolderType` to locate the calendar folder independently of the mailbox locale (NL/FR/EN). Exports results to CSV.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-Mailbox` | No | UPN of a single mailbox. If omitted, all user and shared mailboxes are checked |
| `-OutputPath` | No | CSV report path (default: `.\CalendarPermissions_<timestamp>.csv`) |
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
| `-OutputPath` | No | CSV report path (default: `.\MailboxPermissions_<timestamp>.csv`) |
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
| `-OutputPath` | No | CSV report path (default: `.\GroupPermissions_<timestamp>.csv`) |
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
| `-OutputPath` | No | CSV report path (default: `.\ExternalForwards_<timestamp>.csv`) |
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
| `-OutputPath` | No | CSV report path (default: `.\MailboxSizes_<timestamp>.csv`) |
| `-TenantId` | No | Entra ID tenant ID or domain |

**Examples**

```powershell
# Report all mailboxes
.\Get-MailboxSizes.ps1

# Single mailbox
.\Get-MailboxSizes.ps1 -Mailbox "user@contoso.com"
```
