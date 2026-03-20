# Migrate-HolidaysCalendar.ps1

> **BraveHub Internal Script**
> Ticket: #0298048 | Customer: Onco3R Therapeutics
> Author: Sjoerd Kanon | Date: 19/03/2026

---

## Background and problem statement

Onco3R wanted a shared calendar where employees could book leave, giving everyone an overview of who is absent and when. The initial solution used a **Microsoft 365 Group** as the shared calendar. This worked technically, but had a major unwanted side effect: **all group members received an email notification for every new event** in the calendar. For a company-wide calendar this means everyone receives a mail every time someone books leave.

The solution is a **Room/Resource Mailbox** — the same mechanism used for booking a meeting room in Outlook. Employees add the resource as an attendee to their leave appointment, the booking is automatically approved, and the event appears on the shared calendar. No email notifications, no group membership required.

### M365 Group vs Room Mailbox comparison

| | M365 Group | Room Mailbox |
|---|---|---|
| Shared calendar | ✅ | ✅ |
| Visible to everyone | ❌ (members only) | ✅ |
| Email notifications on events | ❌ (always, cannot be disabled) | ✅ (none) |
| Works like a meeting room | ❌ | ✅ |
| AutoAccept leave | ❌ | ✅ |
| Overlapping bookings allowed | ❌ | ✅ (configurable) |

---

## What the script does

The script performs the full migration in one run:

1. **Detect platform** — selects the correct authentication method (Windows vs macOS/Linux)
2. **Connect Exchange Online** — to create and configure the Room Mailbox
3. **Create App Registration** — automatically creates an Entra ID app with the correct application permissions (or reuses an existing one)
4. **Grant admin consent** — automatically grants consent for all required Graph permissions
5. **Connect Graph (app auth)** — connects using client credentials for write access to other mailboxes
6. **Create Room Mailbox** — creates `holidays-calendar@onco3r.com` as Room type
7. **Set permissions** — sets Default to Reviewer so everyone can read the calendar
8. **Configure AutoAccept** — leave bookings are automatically approved, overlaps allowed
9. **Reconnect Graph (delegated)** — temporarily as a delegated user to read the group calendar (Microsoft limitation: group calendars cannot be read via app auth)
10. **Find M365 Group** — searches for the Holidays group in four ways (mail lowercase, mail original, displayName, Search)
11. **Retrieve events** — fetches all events from the group calendar within the specified date range
12. **Switch back to app auth** — for writing to the Room Mailbox
13. **Copy events** — copies each event to the Room Mailbox calendar with Out of Office status
14. **Delete M365 Group** — optional, removes the M365 Group after migration
15. **Summary** — displays results and user instructions

### Technical note: dual-auth flow

The script deliberately uses **two Graph connections** during execution. This is required because Microsoft has two conflicting limitations:

- **Reading a group calendar** requires *delegated* access (as a signed-in user) — app auth is blocked with 403
- **Writing to a Room Mailbox** requires *application* permissions — delegated access returns 403 on other users' mailboxes

The script therefore automatically switches between both connections at the right moment.

---

## Requirements

### PowerShell version

PowerShell 7+ is required for macOS and Linux. On Windows, PowerShell 5.1 is also supported.

```powershell
$PSVersionTable.PSVersion  # check version
```

Install PowerShell 7: https://aka.ms/powershell

### Install modules

```powershell
Install-Module ExchangeOnlineManagement       -Scope CurrentUser
Install-Module Microsoft.Graph.Applications   -Scope CurrentUser
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module Microsoft.Graph.Calendar       -Scope CurrentUser
Install-Module Microsoft.Graph.Groups         -Scope CurrentUser
Install-Module Microsoft.Graph.Users          -Scope CurrentUser
```

Update modules if already installed:

```powershell
Update-Module ExchangeOnlineManagement
Update-Module Microsoft.Graph
```

### Required permissions

The admin running the script needs:

| Permission | Purpose |
|---|---|
| Exchange Admin or Global Admin | Create Room Mailbox, set permissions |
| Global Admin | Create App Registration + grant admin consent |
| Member of the Holidays M365 Group | Read group calendar via delegated access |

> **Important:** the executing admin must be a member of the Holidays group. Add the admin via M365 Admin Center → Groups → Holidays → Members if not already done.

---

## Platform support

The script automatically detects the operating system:

| Platform | Auth method | Notes |
|---|---|---|
| **Windows** | Interactive browser | Browser opens automatically |
| **macOS** | Device code flow | Code + URL appears in terminal |
| **Linux** | Device code flow | Code + URL appears in terminal |

With device code flow you will see this in the terminal:

```
To sign in, use a web browser to open the page https://login.microsoft.com/device
and enter the code XXXXXXXXX to authenticate.
```

Open the URL in your browser, enter the code, and sign in with your admin account. The script automatically detects when you are done and continues.

---

## Usage

### First run — fully automatic (Mode A)

Do not provide `ClientId` or `ClientSecret`. The script creates an App Registration automatically.

```powershell
.\Migrate-HolidaysCalendar.ps1 `
    -TenantId        "6d5ec429-5783-4739-852c-c872af7302ca" `
    -AdminUPN        "admin@onco3r.onmicrosoft.com" `
    -SourceGroupMail "holidays@onco3r.com"
```

The script displays the `ClientId` and `ClientSecret` at the end. **Store these in Vaultwarden** — the secret is only shown once.

### Subsequent runs — existing App Registration (Mode B)

```powershell
.\Migrate-HolidaysCalendar.ps1 `
    -TenantId        "6d5ec429-5783-4739-852c-c872af7302ca" `
    -AdminUPN        "admin@onco3r.onmicrosoft.com" `
    -ClientId        "717164c5-7927-4066-bdbc-f4ab9f10be56" `
    -ClientSecret    "u2K8Q~JDCMJ~95MilAYfB4o0YMpRU5L2w3H77b0m" `
    -SourceGroupMail "holidays@onco3r.com"
```

### Dry run (no changes)

```powershell
.\Migrate-HolidaysCalendar.ps1 `
    -TenantId        "6d5ec429-5783-4739-852c-c872af7302ca" `
    -AdminUPN        "admin@onco3r.onmicrosoft.com" `
    -SourceGroupMail "holidays@onco3r.com" `
    -WhatIf
```

### With M365 Group deletion after migration

```powershell
.\Migrate-HolidaysCalendar.ps1 `
    -TenantId          "6d5ec429-5783-4739-852c-c872af7302ca" `
    -AdminUPN          "admin@onco3r.onmicrosoft.com" `
    -ClientId          "717164c5-7927-4066-bdbc-f4ab9f10be56" `
    -ClientSecret      "u2K8Q~JDCMJ~95MilAYfB4o0YMpRU5L2w3H77b0m" `
    -SourceGroupMail   "holidays@onco3r.com" `
    -DeleteSourceGroup $true
```

---

## Parameters

| Parameter | Required | Default | Description |
|---|---|---|---|
| `TenantId` | Yes | — | Azure AD Tenant ID (Entra ID → Overview) |
| `AdminUPN` | Yes | — | UPN of the executing admin |
| `ClientId` | No | `""` | AppId of existing App Registration. Empty = create automatically |
| `ClientSecret` | No | `""` | Client Secret. Empty = create automatically |
| `AppName` | No | `BraveHub-HolidaysCalendarMigration` | Name of the App Registration |
| `SourceGroupMail` | No | `holidays@onco3r.com` | Email of the source M365 Group calendar |
| `SourceGroupDisplayName` | No | `Holidays` | DisplayName of the source M365 Group (fallback) |
| `DestinationType` | No | `Room` | Destination mailbox type: `Room` or `Shared` (see below) |
| `DestinationDisplayName` | No | `Holidays Calendar` | Display name of the destination mailbox |
| `DestinationAlias` | No | `holidays-calendar` | Alias of the destination mailbox (must be unique) |
| `DestinationEmail` | No | `holidays-calendar@onco3r.com` | SMTP address of the destination mailbox |
| `DaysBack` | No | `365` | Days back for event retrieval |
| `DaysForward` | No | `730` | Days forward for event retrieval |
| `DeleteSourceGroup` | No | `$false` | Delete M365 Group after migration |

### CalendarProcessing settings (Room Mailbox)

The script configures the following settings on the Room Mailbox:

| Setting | Value | Notes |
|---|---|---|
| `AutomateProcessing` | `AutoAccept` | Automatically approve bookings |
| `AllowConflicts` | `$true` | Multiple people can book the same day |
| `MaximumDurationInMinutes` | `0` | No duration limit (default 1440 = 1 day, too short for multi-day leave) |
| `BookingWindowInDays` | `0` | Book unlimited days in advance |
| `AddOrganizerToSubject` | `$false` | Do not add organiser name to subject |
| `DeleteComments` | `$false` | Preserve comments |
| `DeleteSubject` | `$false` | Preserve subject |

> **Note:** the default `MaximumDurationInMinutes` of 1440 (= 24 hours) causes multi-day leave bookings to be rejected with *"This resource doesn't accept meetings longer than 1440 minutes."* This is why it is explicitly set to `0`.

### DestinationType: Room vs Shared

| | Room Mailbox | Shared Mailbox |
|---|---|---|
| **How to book** | Add as attendee to appointment (like a meeting room) | Create appointment directly from the shared calendar |
| **AutoAccept** | ✅ Automatically approved | ❌ Not applicable |
| **Notifications** | ❌ None | ❌ None |
| **Visible to everyone** | ✅ Via directory | ✅ Via directory |
| **Recommended for leave** | ✅ | ⚠️ Less intuitive |

```powershell
# Room Mailbox (default, recommended)
.\Migrate-HolidaysCalendar.ps1 -DestinationType "Room" ...

# Shared Mailbox
.\Migrate-HolidaysCalendar.ps1 -DestinationType "Shared" ...
```

### Find Tenant ID

```powershell
# Via Graph (if already connected)
(Get-MgOrganization).Id
```

Or via Azure Portal: **Entra ID → Overview → Tenant ID**

---

## Authentication during execution

Depending on the mode you will see two or three login prompts:

| Login | When | Purpose |
|---|---|---|
| Login 1 (delegated) | Mode A only | Create App Registration + grant admin consent |
| Login 2 (delegated) | Always | Read group calendar (Microsoft limitation) |
| Login 3 (automatic) | Always | App auth for writing to Room Mailbox — no interaction needed |

With Mode B (existing app) Login 1 is skipped and you go directly to Login 2.

---

## After migration

### Booking leave (end users)

1. Create an appointment in Outlook
2. Set the duration to **All day** and the status to **Out of office**
3. Add `holidays-calendar@onco3r.com` as an **attendee** (just like a meeting room)
4. Save — the booking is automatically approved
5. The event appears on the shared Holidays Calendar for everyone

### Add Holidays Calendar in Outlook (once per user)

1. Outlook → Calendar → **Add calendar**
2. Choose **Add from directory**
3. Search for `Holidays Calendar` or `holidays-calendar@onco3r.com`
4. Click **Add** — the calendar appears under **People's calendars**

---

## Troubleshooting

### Admin is not a member of the Holidays group

```
[FAIL] Group not found after 4 attempts.
```

If the group exists but is not found via delegated access, the admin is probably not a member. Add the admin:

**M365 Admin Center → Groups → Active groups → Holidays → Members → Add members**

### Destination mailbox alias conflict

```
New-Mailbox: The alias 'holidays-calendar' is already in use.
```

```powershell
.\Migrate-HolidaysCalendar.ps1 `
    -DestinationAlias "leave-calendar" `
    -DestinationEmail "leave-calendar@onco3r.com"
```

### Graph 403 on group calendar

This is a known Microsoft limitation — `Get-MgGroupCalendarEvent` does not work with application permissions. The script resolves this via the dual-auth flow (automatically). If you still see this, verify the admin is a member of the group (see above).

Reference: https://learn.microsoft.com/en-us/graph/known-issues#group-calendar

### Graph 403 on Room Mailbox write

Verify that admin consent has been granted correctly in Entra ID:

**Entra ID → App Registrations → BraveHub-HolidaysCalendarMigration → API Permissions**

All permissions must show status **Granted for Onco3R**. If not, click **Grant admin consent for Onco3R**.

### Client credentials auth failed

```
ClientSecretCredential authentication failed
```

The script automatically attempts a fallback via environment variables. If both methods fail, verify the secret has not expired (expiry date is shown in the summary). Create a new secret if needed:

**Entra ID → App Registrations → BraveHub-HolidaysCalendarMigration → Certificates & secrets → New client secret**

### Events partially failed

Events that could not be copied are logged as `[WARN]` with an error message. The script does not stop on individual event failures but continues. Check the `[WARN]` lines in the output after completion.

---

## Changelog

| Date | Version | Change |
|---|---|---|
| 19/03/2026 | 1.0 | Initial version |
| 19/03/2026 | 1.1 | Platform detection (macOS/Linux device code flow) |
| 19/03/2026 | 1.2 | Fallback group lookup on displayName and Search |
| 19/03/2026 | 1.3 | App Registration setup integrated into main script |
| 19/03/2026 | 1.4 | Dual-auth flow: delegated read + app auth write |
| 19/03/2026 | 1.5 | Fix read-only `$IsWindows`/`$IsMacOS`/`$IsLinux` variables |
| 19/03/2026 | 1.6 | Fix `Get-MgGroupCalendarEvent` 403 via module reload between connections |
| 19/03/2026 | 1.7 | Renamed to Source/Destination, support for Room and Shared Mailbox as destination |
| 19/03/2026 | 1.8 | MaximumDurationInMinutes and BookingWindowInDays set to 0 (unlimited) |

---

---

# Set-Calendar-rights.ps1

Grants a user access rights on the calendar folder of another user in Exchange Online.

Supports **NL / FR / EN** mailbox locales — suitable for Belgian environments with mixed language settings.

---

## Requirements

| Requirement | Value |
|---|---|
| PowerShell | 7.0 or later |
| Module | `ExchangeOnlineManagement` ≥ 3.0 |
| Permissions | Exchange Administrator or delegated mailbox permissions |
| Connection | Active session via `Connect-ExchangeOnline` |

---

## Parameters

| Parameter | Required | Description |
|---|---|---|
| `-User` | Yes | Username (without domain) receiving the permissions |
| `-TargetMailbox` | Yes | Username (without domain) of the target mailbox |
| `-AccessRights` | Yes | Access level (see table below) |

### Access levels

| Value | Description |
|---|---|
| `Owner` | Full control, including deletion and folder management |
| `PublishingEditor` | Read, create, modify, delete and create subfolders |
| `Editor` | Read, create, modify and delete |
| `PublishingAuthor` | Read, create, modify/delete own items and create subfolders |
| `Author` | Read and create, modify/delete own items |
| `NonEditingAuthor` | Read and create, no modifications |
| `Reviewer` | Read-only |
| `Contributor` | Create only (no read access) |
| `AvailabilityOnly` | Free/busy information only |
| `LimitedDetails` | Free/busy with limited details |

---

## Usage

```powershell
# Connect
Connect-ExchangeOnline

# Grant Reviewer rights
.\Set-Calendar-rights.ps1 -User Sjoerd.Kanon -TargetMailbox Jan.Jansen -AccessRights Reviewer

# Grant Editor rights with WhatIf (dry run)
.\Set-Calendar-rights.ps1 -User Sjoerd.Kanon -TargetMailbox Jan.Jansen -AccessRights Editor -WhatIf

# Verbose output for diagnostics
.\Set-Calendar-rights.ps1 -User Sjoerd.Kanon -TargetMailbox Jan.Jansen -AccessRights Author -Verbose
```

---

## Locale support

The script automatically tries the following folder names:

| Language | Folder name |
|---|---|
| Dutch | `\Agenda` |
| French | `\Calendrier` |
| English | `\Calendar` |

Only the path that actually exists in the mailbox succeeds. The others are silently skipped.

---

## Changelog

| Date | Version | Change |
|---|---|---|
| 20/03/2026 | 1.0 | Initial version — replaces MSOnline with ExchangeOnlineManagement |
| 20/03/2026 | 1.1 | French locale (`\Calendrier`) added for Belgium |
