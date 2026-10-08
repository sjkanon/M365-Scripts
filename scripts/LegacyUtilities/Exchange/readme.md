**English** · [Nederlands](readme.nl.md) · [Français](readme.fr.md)

[M365-Scripts](../../../readme.md) › [scripts](../../readme.md) › [LegacyUtilities](../readme.md) › **Exchange**

# Legacy Utilities — Exchange

Mailbox and contact management scripts modernized from a batch of old ad hoc scripts. They sign in through [`Connect-M365.ps1`](../../Startup/readme.md): delegated as the admin by default (browser, or device code / GDAP customer per `load.config.ps1`), app-only with `-ClientId` + `-CertificateThumbprint` or `-AppOnly` (app from `graph.appid.json`). A session for the right tenant that already fits is reused and left connected; only a session the script opened is disconnected. Every script accepts `-TenantId`, `-ClientId`, `-CertificateThumbprint` and `-AppOnly`.

Graph is the default. Five scripts stay on Exchange Online PowerShell because Graph has no API for what they do (see each section); under GDAP they now reach the customer with `-DelegatedOrganization` — before, `-TenantId` was passed as `-Organization`, which only applies to app-only sign-in. `Sync-UserContacts.ps1` is the one script that signs in **app-only by default**, because a delegated token cannot write other users' contacts.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Set-MailboxFolderPermission.ps1`](Set-MailboxFolderPermission.ps1) ([docs](#set-mailboxfolderpermissionps1)) | Grant a folder permission across every folder in a mailbox |
| [`Add-MailboxDelegateAccess.ps1`](Add-MailboxDelegateAccess.ps1) ([docs](#add-mailboxdelegateaccessps1)) | Grant Full Access / Send As on one mailbox, a CSV list, or every mailbox |
| [`New-BulkSharedMailboxes.ps1`](New-BulkSharedMailboxes.ps1) ([docs](#new-bulksharedmailboxesps1)) | Bulk-create shared mailboxes from CSV |
| [`New-BulkMailContacts.ps1`](New-BulkMailContacts.ps1) ([docs](#new-bulkmailcontactsps1)) | Bulk-create Mail Contacts from CSV, optionally add to a distribution group |
| [`Sync-UserContacts.ps1`](Sync-UserContacts.ps1) ([docs](#sync-usercontactsps1)) | Push a shared contact list into users' personal Outlook Contacts |
| [`Start-MailboxMessageTraceReport.ps1`](Start-MailboxMessageTraceReport.ps1) ([docs](#start-mailboxmessagetracereportps1)) | Submit historical message trace report requests |
| [`Remove-DuplicateMailItems.ps1`](Remove-DuplicateMailItems.ps1) ([docs](#remove-duplicatemailitemsps1)) | Find/remove duplicate messages in a mailbox folder via Graph |

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
| `-TenantId` | No | Tenant domain or ID (default: the GDAP customer when `authMode` is GDAP); app-only needs the domain form |
| `-ClientId` / `-CertificateThumbprint` | No | App-only sign-in with this app registration and certificate |
| `-AppOnly` | No | App-only sign-in with the app from `graph.appid.json` |

```powershell
.\Set-MailboxFolderPermission.ps1 -Mailbox "shared@contoso.com" -User "j.doe@contoso.com" -AccessRights Editor -Apply
```

**Notes**
- Stays on Exchange Online: Graph has no API for mailbox folder permissions (only the calendar has `calendarPermission`)

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
| `-TenantId` | No | Tenant domain or ID (default: the GDAP customer when `authMode` is GDAP); app-only needs the domain form |
| `-ClientId` / `-CertificateThumbprint` | No | App-only sign-in with this app registration and certificate |
| `-AppOnly` | No | App-only sign-in with the app from `graph.appid.json` |

*Exactly one of `-Mailbox` / `-CsvPath` / `-AllMailboxes` selects the scope.

```powershell
.\Add-MailboxDelegateAccess.ps1 -Mailbox "sales@contoso.com" -User "j.doe@contoso.com" -Apply
.\Add-MailboxDelegateAccess.ps1 -AllMailboxes -User "helpdesk@contoso.com"   # preview scope first
```

**Notes**
- Stays on Exchange Online: Full Access and Send As are Exchange permissions without a Graph API

---

### New-BulkSharedMailboxes.ps1

Creates shared mailboxes from a CSV (`Name`, `PrimarySmtpAddress`, optional
`Alias`). Dry-run by default.

```powershell
.\New-BulkSharedMailboxes.ps1 -CsvPath .\sharedmailboxes.csv -Apply
```

**Notes**
- Stays on Exchange Online: Graph cannot create shared mailboxes

---

### New-BulkMailContacts.ps1

Creates Mail Contacts from a CSV (`Name`, `ExternalEmailAddress`), optionally
adding each new contact to a distribution group. Dry-run by default.

```powershell
.\New-BulkMailContacts.ps1 -CsvPath .\contacts.csv -DistributionGroup "everyone@contoso.com" -Apply
```

**Notes**
- Stays on Exchange Online: mail contacts and distribution-group membership are Exchange objects; Graph's `orgContact` is read-only and Graph cannot change distribution-group members

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
| `-ClientId` / `-CertificateThumbprint` | No | App registration with the `Contacts.ReadWrite` application permission (plus `GroupMember.Read.All` for `-GroupId`) |
| `-AppOnly` | No | The app from `graph.appid.json` — already the default |
| `-Delegated` | No | Sign in as yourself; only your own mailbox can be a target |
| `-TenantId` | No | Tenant (default: the GDAP customer); also picks the entry in `graph.appid.json` |

```powershell
.\Sync-UserContacts.ps1 -CsvPath .\companycontacts.csv -GroupId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -RemoveExisting -Apply
```

**Notes**
- **App-only by default**: a delegated token, even a Global Administrator's, can only write the signed-in user's own contacts. Without `-ClientId`/`-CertificateThumbprint` the app comes from `graph.appid.json`; when there is none, the script stops and says why
- With `-Delegated` every target must be you; the script stops when another user is in the list or the group
- Only user members of `-GroupId` are targeted (nested groups and devices are skipped)
- A CSV without `DisplayName` or `EmailAddress` is rejected (that check never fired before)

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

**Notes**
- Stays on Exchange Online: `Start-HistoricalSearch` (historical message trace) exists only in Exchange Online PowerShell

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

**Notes**
- Delegated, Graph only reaches another user's mailbox when it is shared with you: the script asks for `Mail.ReadWrite` + `Mail.ReadWrite.Shared`, and the signed-in admin needs **Full Access** on the target mailbox (e.g. `Add-MailboxDelegateAccess.ps1 -AccessRights FullAccess`)
- Without Full Access, run app-only (`-ClientId`/`-CertificateThumbprint` or `-AppOnly`) with the `Mail.ReadWrite` application permission, preferably scoped with RBAC for Applications
- `-IncludeSubfolders` works again: the folder walk failed on the first folder without subfolders
