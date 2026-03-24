# Exchange Scripts

Scripts for Exchange Online calendar and mailbox management.

---

## Scripts

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
