# Legacy Utilities — Exchange

Mailbox and contact management scripts modernized from a batch of old ad hoc
scripts. Connect to Exchange Online / Microsoft Graph automatically if no
session is active; reuse an existing session if already connected.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Set-MailboxFolderPermission.ps1`](#set-mailboxfolderpermissionps1) | Grant a folder permission across every folder in a mailbox |
| [`Add-MailboxDelegateAccess.ps1`](#add-mailboxdelegateaccessps1) | Grant Full Access / Send As on one mailbox, a CSV list, or every mailbox |
| [`New-BulkSharedMailboxes.ps1`](#new-bulksharedmailboxesps1) | Bulk-create shared mailboxes from CSV |
| [`New-BulkMailContacts.ps1`](#new-bulkmailcontactsps1) | Bulk-create Mail Contacts from CSV, optionally add to a distribution group |
| [`Sync-UserContacts.ps1`](#sync-usercontactsps1) | Push a shared contact list into users' personal Outlook Contacts |
| [`Start-MailboxMessageTraceReport.ps1`](#start-mailboxmessagetracereportps1) | Submit historical message trace report requests |
| [`Remove-DuplicateMailItems.ps1`](#remove-duplicatemailitemsps1) | Find/remove duplicate messages in a mailbox folder via Graph |

---

### Set-MailboxFolderPermission.ps1

Applies one folder permission across every folder in a mailbox (skipping Sync
Issues, Recoverable Items, Purges, Versions, Deletions). For full-mailbox
delegate access where `Add-MailboxPermission -AccessRights FullAccess` on its
own isn't the right shape. Dry-run by default.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-Mailbox` | Yes | Target mailbox |
| `-User` | Yes | User to grant access to |
| `-AccessRights` | Yes | Folder permission level (Owner, Editor, Reviewer, ...) |
| `-Apply` | No | Actually grant the permission (default: preview) |
| `-TenantId` | No | Entra ID tenant ID or domain |

```powershell
.\Set-MailboxFolderPermission.ps1 -Mailbox "shared@contoso.com" -User "j.doe@contoso.com" -AccessRights Editor -Apply
```

---

### Add-MailboxDelegateAccess.ps1

Consolidates three old near-duplicate scripts (grant Full Access + Send As to
one mailbox; the same across a CSV list of shared mailboxes; a blanket grant
across every mailbox in the org) into one script. Dry-run by default.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-User` | Yes | User to grant delegate access to |
| `-Mailbox` | * | Single target mailbox |
| `-CsvPath` | * | CSV/TXT list of target mailboxes |
| `-AllMailboxes` | * | Every user/shared mailbox in the tenant |
| `-AccessRights` | No | `FullAccess`, `SendAs`, or `Both` (default) |
| `-AutoMapping` | No | Enable Outlook auto-mapping (default: off) |
| `-Apply` | No | Actually grant access (default: preview) |
| `-OutputPath` | No | CSV report path |
| `-TenantId` | No | Entra ID tenant ID or domain |

*Exactly one of `-Mailbox` / `-CsvPath` / `-AllMailboxes` selects the scope.

```powershell
.\Add-MailboxDelegateAccess.ps1 -Mailbox "sales@contoso.com" -User "j.doe@contoso.com" -Apply
.\Add-MailboxDelegateAccess.ps1 -AllMailboxes -User "helpdesk@contoso.com"   # preview scope first
```

---

### New-BulkSharedMailboxes.ps1

Creates shared mailboxes from a CSV (`Name`, `PrimarySmtpAddress`, optional
`Alias`). Dry-run by default.

```powershell
.\New-BulkSharedMailboxes.ps1 -CsvPath .\sharedmailboxes.csv -Apply
```

---

### New-BulkMailContacts.ps1

Creates Mail Contacts from a CSV (`Name`, `ExternalEmailAddress`), optionally
adding each new contact to a distribution group. Dry-run by default.

```powershell
.\New-BulkMailContacts.ps1 -CsvPath .\contacts.csv -DistributionGroup "everyone@contoso.com" -Apply
```

---

### Sync-UserContacts.ps1

Pushes a CSV contact list into the personal Contacts folder of an explicit user
list or every member of a group, via Microsoft Graph. Every contact it creates
is tagged in `PersonalNotes` so a later run with `-RemoveExisting` can clean up
and refresh without touching the user's own contacts. Dry-run by default.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-CsvPath` | Yes | Contact list (DisplayName, EmailAddress required; GivenName/Surname/CompanyName/BusinessPhone/MobilePhone optional) |
| `-UserList` / `-GroupId` | * | Target scope |
| `-Tag` | No | Marker written to PersonalNotes (default: `Synced-by-Sync-UserContacts`) |
| `-RemoveExisting` | No | Remove previously-synced contacts before re-importing |
| `-Apply` | No | Actually write contacts (default: preview) |

```powershell
.\Sync-UserContacts.ps1 -CsvPath .\companycontacts.csv -GroupId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -RemoveExisting -Apply
```

---

### Start-MailboxMessageTraceReport.ps1

Submits one `Start-HistoricalSearch` request per mailbox (message trace,
sender filter), delivered by email as a compressed export — the way to pull
message trace data older than the 10-day window `Get-MessageTrace` covers.
Modernized replacement for an old script that used the retired MSOnline module
to build the mailbox list.

```powershell
.\Start-MailboxMessageTraceReport.ps1 -NotifyAddress "admin@contoso.com" -AllMailboxes -Apply
```

---

### Remove-DuplicateMailItems.ps1

Graph-based replacement for a third-party EWS duplicate-item-removal tool that
was not carried forward — Microsoft is retiring the EWS API for Exchange
Online. Groups messages in a folder by `internetMessageId`, keeps the oldest of
each group, and (with `-Apply`) deletes the rest. Dry-run by default.

```powershell
.\Remove-DuplicateMailItems.ps1 -Mailbox "user@contoso.com" -IncludeSubfolders -Apply
```

---

**Required modules**

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Install-Module Microsoft.Graph -Scope CurrentUser
```
