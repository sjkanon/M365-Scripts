# Testing — Exchange

Audit scripts for diagnosing and reporting on Exchange Online permissions.
Connect to Exchange Online automatically if no session is active; reuse an existing session if already connected.

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

**Required module**
```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```
