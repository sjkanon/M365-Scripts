**English** · [Nederlands](readme.nl.md) · [Français](readme.fr.md)

# M365-Scripts

> A collection of PowerShell scripts and M365 management tools for MSP engineers, maintained by Sjoerd Kanon.

---

## Table of Contents

- [Folders](#folders)
- [Getting Started](#getting-started)
- [Finding a script](#finding-a-script)
- [Quick Launcher](#quick-launcher)
- [Requirements](#requirements)
- [Menu](#menu)
- [Script Categories](#script-categories)
  - [M365 Management](#️-m365-management)
  - [Exchange](#-exchange)
  - [Entra ID / Graph](#-entra-id--graph)
  - [Intune & Autopilot](#-intune--autopilot)
  - [SharePoint & OneDrive](#-sharepoint--onedrive)
  - [Testing & Diagnostics](#-testing--diagnostics)
  - [Reporting](#-reporting)
  - [Infrastructure & Devices](#️-infrastructure--devices)
  - [Custom Tools](#-custom-tools)
  - [Azure Infrastructure](#️-azure-infrastructure)
  - [Legacy Toolkit Rewrites](#️-legacy-toolkit-rewrites)
- [Repository Structure](#repository-structure)
- [Contributing](#contributing)
- [Version History](#version-history)

---

## Folders

Every workload has its own folder under [`scripts/`](scripts/readme.md), and every folder has a readme: what each script is for, its parameters, examples and notes. Start here and click through; every readme has a breadcrumb at the top to get back up.

| Folder | Description |
|--------|-------------|
| [`ActiveDirectory/`](scripts/ActiveDirectory/readme.md) | On-prem AD DS monitoring (account lockout watcher) — targets a DC/file server directly, not Entra ID |
| [`Azure/`](scripts/Azure/readme.md) | Azure IaaS VM management (disk controller conversion) — targets Azure directly via `Az`, not the M365 tenant |
| [`Entra/`](scripts/Entra/readme.md) | User lifecycle, manager assignment, license reporting, Conditional Access baseline, temporary CA windows, TAP codes, M365 Group audit (Microsoft Graph) |
| [`Exchange/`](scripts/Exchange/readme.md) | Calendar migration/permissions, distribution groups, mailbox/calendar/DKIM/forwarding audits |
| [`Graph/`](scripts/Graph/readme.md) | Microsoft Graph application permission management |
| [`Intune/`](scripts/Intune/readme.md) | Autopilot enrollment, iOS compliance policy updater, corporate wallpaper/lockscreen deployment |
| [`SharePoint/`](scripts/SharePoint/readme.md) | SharePoint Online / OneDrive content operations — recycle bin restore per site or tenant-wide (PnP PowerShell, auto app registration), and where a file went: renamed, moved or deleted (audit log) |
| [`Reporting/`](scripts/Reporting/readme.md) | Computer last-logon report, SharePoint storage report, monthly licensing report |
| [`Device/`](scripts/Device/readme.md) | Windows endpoint maintenance — activation, cleanup, temp files, time sync, audio, OpenVPN diagnostics, Azure/AVD temp disk + pagefile, printer drivers + printers from a JSON file |
| [`Linux/`](scripts/Linux/readme.md) | Linux servers (Debian/Ubuntu, 3CX Phone System) — bash disk cleanup: packages, journal, logs, temp, user caches, Docker, 3CX logs and backups |
| [`Network/`](scripts/Network/readme.md) | TCP port checks, auth/network diagnostics, file I/O stress testing |
| [`RDS/`](scripts/RDS/readme.md) | RDP / RD Web Access login diagnostics, live session monitoring, FSLogix profile diagnostics and disk shrinking, session host image preparation (Teams, Outlook, Copilot) |
| [`SMTP/`](scripts/SMTP/readme.md) | SMTP relay connectivity tests (one-time and recurring) |
| [`Deployment/`](scripts/Deployment/readme.md) | USB toolkit for Windows setup and Autopilot enrollment during OOBE |
| [`DNS/`](scripts/DNS/readme.md) | Resolve and import DNS records into AD-integrated DNS zones |
| [`SAS/`](scripts/SAS/readme.md) | SAS batch job error monitoring with Zabbix integration |
| [`Teams/`](scripts/Teams/readme.md) | Microsoft Teams / SharePoint export and archiving |
| [`Startup/`](scripts/Startup/readme.md) | [`functies.ps1`](scripts/Startup/functies.ps1) M365 function library + module bootstrap + syntax checker, dot-sourced by the menu |
| [`Custom Scripts/`](scripts/Custom%20Scripts/readme.md) | Path-pinned scripts — Office theme deployment (hardcodes its download URL to this repo path) |
| [`TenantOnboarding/`](scripts/TenantOnboarding/readme.md) | New-tenant provisioning, multi-tenant/GDAP reporting, app deployment, device config, OneDrive management, user management — modernized from a retired internal tenant-setup toolkit |
| [`Office365Toolkit/`](scripts/Office365Toolkit/readme.md) | Security/Exchange/Intune rewrites of still-useful capabilities from the retired `directorcia/Office365` (CIAOPS) toolkit |
| [`PatronToolkit/`](scripts/PatronToolkit/readme.md) | Entra/Exchange/Intune/Security/SharePoint/Teams rewrites of still-useful capabilities from the retired `directorcia/patron` toolkit |
| [`LegacyUtilities/`](scripts/LegacyUtilities/readme.md) | Misc modernized scripts (Exchange, Entra, Teams, Network, Device, Workspace 365) from assorted small tools in the retired internal toolkit |

Know the script name but not the folder? [`scripts/INDEX.md`](scripts/INDEX.md) lists every script A–Z.

---

## Getting Started

```powershell
.\load.ps1
```

On first run, [`load.ps1`](load.ps1) will:

1. Ask for your admin UPN and display name — saved to a gitignored `load.config.ps1`
2. Ask whether you want delegated GDAP mode as default and optionally store a default customer domain
3. Ask whether Graph should use device code sign-in by default
4. Check the required modules — missing, older than their minimum, or with an update on the PowerShell Gallery — and install or update them
5. Import the core modules
6. Open the interactive menu

From then on it skips the questions. The module check (step 4) runs at every start without asking; it asks the gallery at most once every 24 hours, and with everything current it prints one line. `.\load.ps1 -SkipModuleCheck` skips it once.

To run the launcher automatically at Windows sign-in:

```powershell
.\load.ps1 -SetupStartup
```

To remove the startup shortcut later:

```powershell
.\load.ps1 -RemoveStartup
```

> You can also run [`.\menu.ps1`](menu.ps1) directly — it will ask for your UPN as a fallback.
> To reinstall or update modules manually: [`.\scripts\Startup\Install-Modules.ps1`](scripts/Startup/Install-Modules.ps1) or [`.\scripts\Startup\Update-Modules.ps1`](scripts/Startup/Update-Modules.ps1). The list of modules is [`RequiredModules.psd1`](scripts/Startup/RequiredModules.psd1) — add a module there and every machine installs it at its next start. The docs hook reports a module a script loads that is not on the list.

---

## Finding a script

| Where | What it gives you |
|-------|-------------------|
| [`scripts/INDEX.md`](scripts/INDEX.md) | Every script A–Z on one page — name, folder and what it does. Ctrl-F this when you know roughly what you want but not where it lives |
| [`scripts/readme.md`](scripts/readme.md) | The other direction: what each workload folder is for |
| [`.\menu.ps1`](menu.ps1) | The curated interactive launcher for the everyday tasks |
| `f <term>` | Fuzzy search from your shell, described under [Quick Launcher](#quick-launcher) below |
| Any folder readme | Every script name in a `Scripts` table links straight to the file, with a `docs` link to its section on the same page |

[`INDEX.md`](scripts/INDEX.md) is generated from the scripts' own `.SYNOPSIS` headers by [`scripts/Startup/Update-ScriptIndex.ps1`](scripts/Startup/Update-ScriptIndex.ps1) — rerun it (or run it with `-Check`) whenever a script is added, renamed, moved or removed.

Those links are checked, not assumed: [`scripts/Startup/Test-MarkdownLinks.ps1`](scripts/Startup/Test-MarkdownLinks.ps1) walks every readme and fails on a file that is not there, or an anchor with no heading behind it.

---

## Quick Launcher

[`f.ps1`](f.ps1) creates one short command per script in [scripts/](scripts/), so you no longer have to navigate to a folder first. Dot-source it from your profile:

```powershell
notepad $PROFILE
. "C:\Users\<you>\Git\M365-Scripts\f.ps1"
```

The leading dot is not optional — without it the file runs in its own scope and the commands are gone again immediately. `.\f.ps1 -Install` writes the line for you (and does nothing if it is already there).

After restarting your shell:

```powershell
f-test-dkimconfig -Domain contoso.com
f-dkimconfig -Domain contoso.com          # short form, when the noun is unique
f-get-mailboxsizes
f-m365                                    # the whole catalogue
f-m365 mailbox                            # filtered
```

The wrappers copy the parameter block of the target script from its AST, so `-Dom<Tab>` completes and a `ValidateSet` is enforced before anything runs. Default values are dropped from the wrapper on purpose: only bound parameters are forwarded, so the script's own defaults still apply.

### Fuzzy search

For when you know roughly what a script is called but not exactly, `f` matches on name, folder and `.SYNOPSIS`:

```powershell
f dkim                    # one match -> runs it
f entra group             # several matches -> numbered picker
f dkim -Domain contoso.com
f -List mailbox           # show matches, run nothing
f -Show trace             # path, synopsis and parameters
f -Edit bloatware         # open in $env:EDITOR, VS Code, or notepad
```

### Notes

| | |
|---|---|
| Naming | `f-<full-script-name>` always exists; `f-<noun>` is added only where it stays unambiguous. Stripping the verb collides 11 times here (`Detect-`, `Install-` and `Uninstall-ClaudeDesktop-Intune` become the same noun), so the full name is the one you can always count on. |
| Cache | Generated wrappers live in `.f-index.json` (gitignored). Parsing every script costs ~500 ms, reading the cache ~30 ms, which is what keeps shell start quick. It refreshes itself when a script is added, removed or changed; `f-refresh` forces it. |
| Collisions | Existing commands are never overwritten. A profile can load more than one of these — `ScriptRunner.Profile.ps1` in *itce-testing* owns `f-scripts`, which is why this catalogue is called `f-m365`. Anything skipped is reported on load. |
| Uninstall | `.\f.ps1 -Uninstall` removes the marked block from `$PROFILE`. A hand-written dot-source line is reported, not deleted. |

---

## Requirements

| Requirement | Details |
|-------------|---------|
| PowerShell | 7.0+ (cross-platform); individual scripts support PS 5.1 on Windows |
| Permissions | Microsoft 365 admin permissions for the target workload |
| Sign-in | Every M365 script signs in through [`Connect-M365.ps1`](scripts/Startup/readme.md#connect-m365ps1): Microsoft Graph first, **delegated by default** (device code and GDAP customer from `load.config.ps1`), app-only with `-ClientId`/`-CertificateThumbprint` or `-AppOnly` (`graph.appid.json`). Exchange Online, Teams and PnP only where Graph has no API |
| Execution Policy | Windows only: `Set-ExecutionPolicy RemoteSigned -Scope CurrentUser` |

---

## Menu

The launcher ([`menu.ps1`](menu.ps1)) covers all tools in this repo. Press a key to launch:

| Key | Category | Tool |
|-----|----------|------|
| `1` / `F1` | Testing | [Test-Ports](scripts/Network/Test-Ports.ps1) — TCP port checker |
| `2` / `F2` | Exchange | [Migrate-Calendar](scripts/Exchange/Migrate-Calendar.ps1) |
| `3` / `F3` | Exchange | [Set-Calendar-rights](scripts/Exchange/Set-Calendar-rights.ps1) |
| `4` / `F4` | Testing | [Test-SMTP (one-time)](scripts/SMTP/testsmtp.ps1) |
| `5` / `F5` | Testing | [Test-SMTP (every 5 min)](scripts/SMTP/testsmtp_5min.ps1) |
| `6` / `F6` | Device | [Restart-Time-Sync](scripts/Device/Time%20sync/Restart-Time-Sync.ps1) |
| `7` / `F7` | Device | [Detect-AudioDevices](scripts/Device/audio/detect-audiodevices.ps1) |
| `8` / `F8` | Device | [Disable-InternalMic](scripts/Device/audio/Disable-internalmic.ps1) |
| `I` | Device | [Remove-OemBloatware](scripts/Device/Remove-OemBloatware.ps1) — remove OEM + generic Store bloatware |
| `T` | Device | [Update-TeamsClient](scripts/Device/Update-TeamsClient.ps1) — update new Teams + the Outlook meeting add-in when outdated |
| `R` | Device | [Repair-AppxPackageStore](scripts/Device/Repair-AppxPackageStore.ps1) — repair AppX packages failing with 0x80070490 (Teams, new Outlook, FSLogix) |
| `K` | Device | [FSLogix-Shrink](scripts/RDS/Invoke-FSLogixShrink.ps1) — shrink FSLogix profile disks on a share, or check compaction at sign-out |
| `J` | Device | [Update-SessionHostImage](scripts/RDS/Update-SessionHostImage.ps1) — prepare a multi-session image / AVD hosts for Teams, Outlook and Copilot with FSLogix |
| `N` | Device | [Install-Printer](scripts/Device/Printer/Install-Printer.ps1) — install printer drivers (from GitHub) and printers from a JSON file |
| `9` / `F9` | Startup | [Install-Modules](scripts/Startup/Install-Modules.ps1) |
| `U` | Startup | [Update-Modules](scripts/Startup/Update-Modules.ps1) — check/update the required modules, optionally every other installed module too |
| `Z` | Startup | [Test-RequiredModules](scripts/Startup/Test-RequiredModules.ps1) — modules scripts load that `RequiredModules.psd1` does not list |
| `X` | Startup | [Update-ScriptIndex](scripts/Startup/Update-ScriptIndex.ps1) — rebuild [`scripts/INDEX.md`](scripts/INDEX.md), the A–Z list of every script |
| `L` | Startup | [Test-MarkdownLinks](scripts/Startup/Test-MarkdownLinks.ps1) — check every readme link: files and in-page anchors |
| `M` | Startup | [Convert-MarkdownToHtml](scripts/Startup/Convert-MarkdownToHtml.ps1) — build a styled HTML page from a markdown document, for IT Glue |
| `A` / `F10` | Reporting | [Licensing-Report](scripts/Reporting/Licensing/genereer_rapport.ps1) |
| `P` | Reporting | [SharePoint-Perms](scripts/Reporting/Get-SharePointPermissionsReport.ps1) — report who has access to what, at every level |
| `S` | SharePoint | [SharePoint-Structure](scripts/SharePoint/Provisioning/readme.md) — provision/check metadata, libraries and rights |
| `F` | Startup | [Enable-LauncherStartup](menu.ps1) — add launcher to Windows Startup |
| `G` | Startup | [Disable-LauncherStartup](menu.ps1) — remove launcher from Windows Startup |
| `B` | M365 | [Connect-Tenant](scripts/Startup/readme.md#functiesps1) |
| `H` | M365 | [Test-GdapConnection](scripts/Startup/readme.md#functiesps1) — validate delegated GDAP access |
| `C` | M365 | [Exchange Online submenu](scripts/Startup/readme.md#functiesps1) |
| `D` | M365 | [Entra ID / Graph submenu](scripts/Startup/readme.md#functiesps1) |
| `E` | M365 | [MSP Admin submenu](scripts/Startup/readme.md#functiesps1) |

M365 options (`B`, `C`, `D`, `E`, `H`) lazy-load [`functies.ps1`](scripts/Startup/functies.ps1) on first use — Graph authentication is only triggered when needed.

**Exchange submenu (`C`)**

| Key | Tool |
|-----|------|
| `8` | [Test-CalendarPermissions](scripts/Exchange/Test-CalendarPermissions.ps1) — audit calendar folder permissions (all or single mailbox) |
| `9` | [Test-MailboxPermissions](scripts/Exchange/Test-MailboxPermissions.ps1) — audit Full Access, Send As, Send on Behalf |
| `A` | [Test-GroupPermissions](scripts/Exchange/Test-DistributionGroupPermissions.ps1) — audit DG managers, Send As, Send on Behalf, member counts |
| `B` | [Test-DkimConfig](scripts/Exchange/Test-DkimConfig.ps1) — validate DKIM signing config and DNS CNAME/TXT records |
| `C` | [Get-ExternalForwards](scripts/Exchange/Get-ExternalForwards.ps1) — audit mailboxes with external forwarding configured |
| `D` | [Get-MailboxSizes](scripts/Exchange/Get-MailboxSizes.ps1) — mailbox size report sorted by storage used |
| `E` | [Move-InboxToArchive](scripts/Exchange/Move-InboxToArchive.ps1) — archive Inbox messages to Archive folder |
| `F` | [Set-DL-Dynamic-Static](scripts/Exchange/Set-Distributionlist-dynamic-static.ps1) — resolve a dynamic distribution group into a static group |
| `H` | [Get-CalendarMappings](scripts/Exchange/Get-CalendarMappings.ps1) — where a calendar is mapped in Outlook, next to the rights (search by keyword, e.g. `balie`, or all/selected mailboxes) |
| `I` | [Convert-SharedCalendar](scripts/Exchange/Convert-SharedCalendarToResource.ps1) — move a shared calendar out of a user's mailbox into a room/equipment mailbox (always previews first) |
| `J` | [Move-SharedCalendar](scripts/Exchange/Move-SharedCalendar.ps1) — all in one: find a calendar by keyword, move it into a resource mailbox, list who has to switch |
| `K` | [Get-DLMembers](scripts/Exchange/Get-DistributionGroupMembers.ps1) — export every distribution list with its members to Excel, or only the lists holding one address, one domain, or a domain tree (`-Recurse` to expand nested lists) |
| `L` | [Restore-MailboxMessages](scripts/Exchange/Restore-MailboxMessages.ps1) — put back mail that was moved or deleted on a given day, or from a date until now, and show who did it (always previews first) |

**Entra ID submenu (`D`)**

| Key | Tool |
|-----|------|
| `A` | [Test-M365GroupMembership](scripts/Entra/Test-M365GroupMembership.ps1) — audit M365 Group / Teams owners and members |
| `B` | [New-M365User](scripts/Entra/New-M365User.ps1) — create a single new user (auto-generated password, optional license) |
| `C` | [Import-M365Users](scripts/Entra/Import-M365Users.ps1) — bulk create users from CSV, dry-run by default |
| `D` | [New-TemporaryCA](scripts/Entra/New-TemporaryConditionalAccessPolicy.ps1) — create temporary Conditional Access policy for user/group (duration or start/end datetime) |
| `E` | [Remove-TemporaryCA](scripts/Entra/Remove-TemporaryConditionalAccessPolicies.ps1) — remove expired or all temporary CA policies |
| `F` | [New-UserTAP](scripts/Entra/New-UserTemporaryAccessPass.ps1) — create Temporary Access Pass for a user |
| `G` | [Get-M365UserLicenses](scripts/Entra/Get-M365UserLicenses.ps1) — report assigned licenses for a set of users |
| `H` | [Import-CA-Baseline](scripts/Entra/Import-ConditionalAccessBaseline.ps1) — import the community Conditional Access baseline |
| `I` | [Set-UserManager](scripts/Entra/Set-UserManager.ps1) — report/bulk-set manager for a set of users |

---

## Script Categories

### ☁️ M365 Management

📂 Folder: [`Startup/`](scripts/Startup/readme.md)

Interactive M365 management functions via Microsoft Graph and Exchange Online. Loaded as a library through the menu. CSV and log exports go to `C:\Temp\` on Windows or `~/Downloads/` on macOS.

| Area | Features |
|------|----------|
| Exchange Online | Shared mailbox access, locale, aliases, distribution groups, auto-reply, sent-items copy |
| Entra ID / Graph | Tenant admins, domains, licenses, users, password reset, sign-in logs, bulk create/remove, temporary CA windows, TAP codes |
| MSP Admin | Create/manage MSP admin account across customer tenants |

---

### 📧 Exchange

📂 Folder: [`Exchange/`](scripts/Exchange/readme.md)

Scripts for calendar and mailbox management.

- Calendar migration between users
- Set calendar folder permissions (NL/FR/EN locale support)
- **[Get-DistributionGroupMembers.ps1](scripts/Exchange/Get-DistributionGroupMembers.ps1)** — who is on which distribution list, as one Excel workbook meant to go straight to the customer
  - `Overzicht` sheet (one row per list) and `Leden` sheet (one row per member), both filterable tables with a frozen header row, headers in Dutch
  - `-Member jan@contoso.com` answers "which lists is this person on?"; `-Member @be.verizon.com` answers it for one domain and `-Member *.verizon.com` for a domain and every subdomain, matching aliases and `ExternalEmailAddress` so external contacts are actually found
  - `-Recurse` expands nested lists — without it someone who only receives mail through a nested group is invisible, and a filter reports "no hits" on a list that does deliver to them
- **Get-MessageTraceReport.ps1** — trace who received what, at what exact time, and where it was forwarded to
- **[Remove-PhishingMessage.ps1](scripts/Exchange/Remove-PhishingMessage.ps1)** — delete a phishing message from one, several, or all mailboxes; dry-run by default
  - Two engines: **Purview** Content Search + purge (tenant-wide, the only one that can HardDelete) and **Graph** (per-mailbox, no search-index lag, per-message report)
  - `Recycle` / `SoftDelete` / `HardDelete`; refuses to run without a content selector so a date range alone can never match every message
  - Loops purge rounds automatically around Purview's 10-items-per-mailbox limit, and writes a CSV of everything matched and deleted
- **[Restore-MailboxMessages.ps1](scripts/Exchange/Restore-MailboxMessages.ps1)** — put back messages that were moved or deleted on a given day, and report who did it; preview by default
  - Deleted messages go back via `Restore-RecoverableItems` (Deleted Items, Recoverable Items, Purges); moved messages are traced to their original folder through the audit log and moved back over Graph
  - Names the actor from the Unified Audit Log — account, owner/delegate/admin, client, IP — and says which actions are not audited on the mailbox

---

### 👤 Entra ID / Graph

📂 Folder: [`Entra/`](scripts/Entra/readme.md), [`Graph/`](scripts/Graph/readme.md)

Scripts for user lifecycle management via Microsoft Graph.

- Create a single M365 user (auto-generated password, optional license)
- Bulk create users from CSV — dry-run by default, passwords in CSV output
- Bulk remove users from CSV — dry-run by default, CSV report

---

### 📱 Intune & Autopilot

📂 Folder: [`Intune/`](scripts/Intune/readme.md)

Scripts for device enrollment, Autopilot registration, and compliance policy management.

- Retrieve Windows Autopilot hardware info
- CMD-based Autopilot enrollment helper
- **[Compare-IntuneConfig.ps1](scripts/Intune/Compare-IntuneConfig.ps1)** — compare a customer tenant's Intune configuration against an MSP baseline backup (drift detection), via the `IntuneBackupAndRestore` module — read-only
- **iOS Compliance Updater** — automatically keeps the minimum iOS version requirement in Intune up to date
  - Fetches latest iOS version from Apple's RSS feed (with fallback to Apple Support page)
  - Compares against current policy minimum and patches via Microsoft Graph API
  - One-time setup via [`Setup.ps1`](scripts/Intune/iOS-Compliance-Updater/Setup.ps1) (creates App Registration, assigns permissions, writes `config.json`)
  - Runs weekly as a Windows scheduled task (SYSTEM, every Monday 07:00)
  - Dry-run mode (`-WhatIf`) — shows what would change without applying
- **Desktop** — deploy lockscreen to start and desktop; set corporate wallpaper via Intune:
  - [`Set-CorporateWallpaper.ps1`](scripts/Intune/Desktop/Background/Desktop/Set-CorporateWallpaper.ps1) — generic, reusable per customer; only the CONFIGURATION block needs updating
  - [`Make-lockscreen.ps1`](scripts/Intune/Desktop/Background/Lockscreen/Make-lockscreen.ps1) — applies the same corporate image as Windows lockscreen via PersonalizationCSP
  - Downloads wallpaper from a public URL; compares SHA256 hash against existing file — skips if already up to date, applies if new or changed
  - Lockscreen flow downloads from internet via `Invoke-WebRequest`, validates image headers (`jpg/png/bmp`), blocks HTML responses, and normalizes common GitHub blob/raw URLs
  - Applies via PersonalizationCSP (MDM enforcement), WinAPI (immediate), HKCU registry (style), and Default User profile (new accounts)
  - Log: `C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\CorporateWallpaper-<CLIENTNAME>.log`
  - Lockscreen log: `C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\CorporateLockscreen-<CLIENTNAME>.log`
  - Deploy via Intune: **Run as SYSTEM**, 64-bit PowerShell

  | Variable | Description |
  |---|---|
  | `$ImageUrl` | Public URL to the wallpaper image (PNG or JPG) |
  | `$WallpaperStyle` | `10` = Fill · `6` = Fit · `2` = Stretch · `0` = Tile · `22` = Span |
  | `$ClientName` | Customer name — used in log filename and local image path |

---

### 📁 SharePoint & OneDrive

📂 Folder: [`SharePoint/`](scripts/SharePoint/readme.md)

Content operations on SharePoint Online sites and OneDrive via PnP PowerShell.

#### Trace a File

**[Trace-SharePointFile.ps1](scripts/SharePoint/Trace-SharePointFile.ps1)** — "where did my file go?" for OneDrive and SharePoint, from the Unified Audit Log (Exchange Online, not PnP). Read-only.

- Follows renames, moves, copies, deletes and restores of one file by name (wildcards), old URL or item ID — a chain `A → B → C` ends at C
- Replays folder renames, moves and deletes onto the file, because those move every file inside without a record per file
- Every time in Brussels time with the UTC offset; `-StartDate` / `-EndDate` in Belgian notation (`15-09-2026 08:30`), a bare end date includes the whole day
- Reads per day and splits a slice with more than 50,000 records, retries failing searches; last known location and status per item, CSV plus the raw audit records as JSON

#### Revoke User Access

**[Revoke-SharePointUserAccess.ps1](scripts/SharePoint/Revoke-SharePointUserAccess.ps1)** — the counterpart to the permissions report: that one says who can reach what, this one takes it away. Reports by default, removes with `-Apply`, and writes a CSV of every grant found and what happened to it.

- Site collection administrator first, because it overrides every role assignment below it
- Direct role assignments on the site, a sub-site, a list or library, a folder or a single file
- SharePoint group membership, and **sharing links** — the `SharingLinks.*` groups that "Anyone with the link" and "Specific people" actually put a person in
- It deliberately never changes Entra ID group membership: a user who gets in through a security or Microsoft 365 group keeps that access, and removing them from SharePoint does not take it away. Those routes are reported with the group named, so offboarding is two steps and the second one is visible
- Grants to `Everyone` are left alone for the same reason in reverse — removing one revokes access for the whole tenant, not for this person
- `ConfirmImpact = 'High'`, so it asks per removal unless `-Confirm:$false`

**[Test-SharePointAccessScripts.ps1](scripts/SharePoint/Test-SharePointAccessScripts.ps1)** verifies this script and the permissions report without touching a tenant: the app-only auth layer both share must stay byte-identical, and the revocation funnel must record a dry run without executing it, execute and record under `-Apply`, and keep refusing the grants it must not remove.

#### Recycle Bin Restore

Restore deleted files and folders from a site or OneDrive recycle bin — dry-run by default, `-Apply` to actually restore.

- One site (`-SiteUrl`, works for OneDrive too) or every SharePoint site in the tenant (`-AllSites`) — the tenant sweep skips OneDrive, system and locked sites, and a failing site does not abort the run
- Narrow the sweep with `-SiteFilter` and try it on a handful of sites first with `-MaxSites`
- Filter by name, original folder, who deleted it, and a deletion time window
- First-stage (user) and second-stage (site collection) recycle bin, or both
- Restores folders first, shallow paths first — a file cannot be restored into a folder that is itself still deleted
- Creates the required Entra app registration automatically on the first run against a tenant, then caches the client ID in `pnp.appid.json` (gitignored) — later runs go straight to the interactive login
- `-GrantSiteAdmin` temporarily makes you site collection admin per site and removes the rights afterwards — needed for another user's OneDrive and effectively required for `-AllSites`
- Restores in batches of up to 200 via a single server call (`-BatchSize`) — a failed batch falls back to item-by-item so one bad file does not sink the rest
- Timing throughout: how long reading the bin took, an up-front estimate, a progress bar with live ETA, and the real duration in the summary
- CSV report of every item, restored or failed, including the site, the SharePoint error, its batch number and how long it took

#### Structure Provisioning

Provision and maintain a whole SharePoint structure — metadata model, content types, libraries and group permissions — from one JSON config. See [`scripts/SharePoint/Provisioning/`](scripts/SharePoint/Provisioning/readme.md).

- [`Install-SharePointStructure.ps1`](scripts/SharePoint/Provisioning/Install-SharePointStructure.ps1) — **the one-command build**: registers the Entra app itself, runs the three provisioning steps in the only order that works, verifies the result, and with `-TemporaryApp` removes the app registration again so nothing is left behind in a tenant you do not manage day to day
- The model lives in the config, not in the code: a second MSP client is a second config file, not a second fork of four scripts
- [`New-SharePointMetadata.ps1`](scripts/SharePoint/Provisioning/New-SharePointMetadata.ps1) — managed metadata term set, site columns and content types, on **every** site in the config (a Teams private channel is its own site collection, and a site column does not reach across one)
- [`Set-SharePointLibraries.ps1`](scripts/SharePoint/Provisioning/Set-SharePointLibraries.ps1) — libraries, Teams channel folders, content type binding, per-folder content type order, default column values, grouped views, and one Entra ID security group per pillar per access level; `-EnsureGroups` creates the groups on the way
- [`Update-SharePointShareStatus.ps1`](scripts/SharePoint/Provisioning/Update-SharePointShareStatus.ps1) — derives a Deelstatus column from the permissions actually on each file (Anyone link, guest, organisation link, or nothing) and flags anything tagged Intern/Vertrouwelijk sitting behind an external link; exit code 2 for a scheduled RMM job
- [`Test-SharePointStructure.ps1`](scripts/SharePoint/Provisioning/Test-SharePointStructure.ps1) — read-only drift check classifying every difference as Missing / Different / Extra; exit code 2 means somebody changed something
- All four are idempotent and support `-WhatIf`; interactive or app-only with a certificate
- Cross-cutting brand views (`Scope = RecursiveAll`) make "brand as a tag" real: *Alles - Northwind* is one flat list across every pillar folder, including everything tagged **Beide** — one file, two brands, no copies. Plus *Nog te taggen*, *Extern gedeeld* and *Te archiveren*
- [`SharePoint-Handleiding.md`](scripts/SharePoint/Provisioning/SharePoint-Handleiding.md) — end-user documentation in Dutch to hand to the customer: the three ways of adding a file and why they behave differently, what each label means, and what happens the moment you tag something
- Documented rather than hidden: unique permissions on a **standard**-channel folder are what this model asks for and what Microsoft does not support — members keep seeing the channel and get an error on the Files tab. `-SkipChannelFolderPermissions` is the conservative alternative

---

### 🧪 Testing & Diagnostics

Audit and diagnostic scripts, organised by workload. Self-connecting where applicable — reuse an existing session or connect automatically. CSV exports go to `C:\Temp\` on Windows or `~/Downloads/` on macOS.

#### Exchange Online

📂 Folder: [`Exchange/`](scripts/Exchange/readme.md)

- Audit calendar folder permissions (locale-independent, exports CSV)
- Audit Full Access, Send As, Send on Behalf delegation (exports CSV)
- Audit distribution group managers, Send As, Send on Behalf, member counts (exports CSV)
- Validate DKIM signing config and DNS CNAME/TXT records; lists required actions
- Audit mailboxes with external forwarding to non-tenant domains (security audit, exports CSV)
- Report mailbox sizes and item counts sorted by storage used (exports CSV)

#### Entra ID / Graph

📂 Folder: [`Entra/`](scripts/Entra/readme.md)

- Audit M365 Group (incl. Teams) owners and members — one row per entry, exports CSV
- Create temporary Conditional Access policies for installation windows (duration or exact local start/end)
- Auto-clean temporary CA policy at end time (same session) and cleanup script for missed sessions
- Create Temporary Access Pass (TAP) codes for user onboarding/support

#### SharePoint Online

📂 Folder: [`SharePoint/`](scripts/SharePoint/readme.md)

- Report storage usage across all sites in a tenant — current file sizes + version history per library and per file
- Two-phase: first enumerates all sites and document libraries, then retrieves storage data
- Quick mode (quota data only) or full recursive scan with `-Apply`

#### Network & Connectivity

📂 Folder: [`Network/`](scripts/Network/readme.md)

- Test TCP port connectivity on any host — single ports, ranges (`1294:1494`), combinations (`80,443,1294:1494`)
- One-time SMTP test with interactive credential prompt
- Recurring SMTP test (every 5 minutes) with saved encrypted password
- Auth & network diagnostics — Event Viewer (logon failures, Kerberos, NTLM, DC availability), time sync, DNS, TCP, UNC shares, optional log scan; exports txt report to `C:\Temp\`
- File I/O diagnostics — write/append/read/delete loop on any path; classifies failures as AUTH/NETWORK/TIMEOUT/DISK/PATH; on each failure captures FileSystemWatcher events, NTFS permission diff vs baseline, open process handles (Handle.exe auto-downloaded from Sysinternals), new process snapshot, Kerberos tickets, and Security event log; stops after 3 failures
- **UniFi** — network documentation HTML report (devices, firmware, uptime, per site) and firmware upgrade tooling for a UniFi Controller/UniFi OS console; credentials via `Get-Credential`, never hardcoded

#### Device

📂 Folder: [`Device/`](scripts/Device/readme.md)

- OpenVPN Connect diagnostics — PnP adapters, services, routes, DNS, Event Log, conflicting VPN software; exports txt report to `C:\Temp\`

#### RDS

📂 Folder: [`RDS/`](scripts/RDS/readme.md)

- RDP + RD Web Access diagnostics ([`Test-RDSDiagnostics.ps1`](scripts/RDS/Test-RDSDiagnostics.ps1)) — diagnose why users cannot log in to an RDP or RDWeb server:
  - Services (TermService, SessionEnv, UmRdpService), RDP enabled/disabled, NLA, session limits, RD Licensing, firewall rules, active sessions
  - HTTPS certificate validity and expiry on RDWeb; IIS app pool and RD Gateway status (local only)
  - User account checks: enabled, locked, password expired, Remote Desktop Users membership
  - Event log analysis: failed logons (4625), lockouts (4740), Kerberos failures (4771), session disconnect reasons (20/40)
  - Timestamped log file saved to `C:\Temp\`; `-IncludeEventLogs` for event analysis

- Real-time RDS monitor ([`Watch-RDSLive.ps1`](scripts/RDS/Watch-RDSLive.ps1)) — polls event logs every N seconds and streams new events to console + log file:
  - Session events: logon (21), reconnect (22/25), logoff (23), disconnect (24), logon failed (20), disconnect reason (40) with human-readable reason codes
  - Security: failed RDP logons (4625 type 10), account lockouts (4740)
  - Licensing: `TerminalServices-Licensing/Admin` events + System log `TermServLicensing` provider
  - Heartbeat line per poll showing active session count and new event count
  - Run directly on each RDS/RDWeb server; `-IntervalSeconds` (default 20), `-NoLogFile` to skip file output

- FSLogix profile diagnostics ([`Get-FSlogix-errors.ps1`](scripts/RDS/Get-FSlogix-errors.ps1)) — collects version, configuration, attached containers, SMB/Azure Files state and FSLogix/disk events on an AVD session host into one transcript

- FSLogix disk shrink ([`Invoke-FSLogixShrink.ps1`](scripts/RDS/Invoke-FSLogixShrink.ps1)) — gives back the space dynamic profile/ODFC VHDX files keep:
  - Downloads Invoke-FslShrinkDisk (FSLogix team) at a pinned commit and verifies its SHA-256
  - `-ReportOnly` lists every container on the share, largest first; otherwise shrinks them and summarises GB recovered and disks not processed (in use)
  - `-CheckHost` checks whether FSLogix's built-in compaction at sign-out can run (version, `VHDCompactDisk`, `defragsvc`, dynamic disks)

- Session host image ([`Update-SessionHostImage.ps1`](scripts/RDS/Update-SessionHostImage.ps1)) — makes a Windows 11 multi-session image or AVD host fit for new Teams, new Outlook and Copilot with FSLogix, without changing FSLogix:
  - Checks WebView2, the AppX frameworks the apps depend on, provisioned builds and per-user drift, Teams on AVD (SlimCore, the WebRTC redirector retired on 1 October 2026), Shared Computer Activation and the sign-in broker
  - Sets the Store and Teams update hold-back and fixes the apps through `Repair-AppxPackageStore.ps1` and `Update-TeamsClient.ps1`; `-ComputerName` compares the whole pool, `-ForCapture` checks Sysprep readiness

---

### 📊 Reporting

📂 Folder: [`Reporting/`](scripts/Reporting/readme.md)

#### Computer Last Logon Report

Report last logon date for all computer objects in one or more OUs and export to CSV.

- Queries Active Directory for computers in specified OUs (e.g. `OU=Laptops`, `OU=Computers`)
- Two accuracy modes: `LastLogonTimestamp` (fast, max 14-day delay) or `-AllDCs` (queries every DC for exact `LastLogon`)
- Four statuses: **Active** · **Active (pwd recent)** · **Stale** · **Never** · **Disabled**
- `Active (pwd recent)`: device falsely marked stale due to 14-day replication delay — `PasswordLastSet` < 35 days confirms the machine is online (computer accounts auto-rotate password every ~30 days)
- CSV columns: Name, Status, Enabled, LastLogon, DaysSinceLogon, PasswordLastSet, DaysSincePasswordSet, OS, IPv4, OU path, Created, Description
- Supports multiple OUs in one run; `-IncludeDisabled` to include disabled objects

#### SharePoint Permissions Report

**[Get-SharePointPermissionsReport.ps1](scripts/Reporting/Get-SharePointPermissionsReport.ps1)** — who can reach which SharePoint, through which group, at what level. Read-only: every call it makes is a GET.

- Starts from one consolidated view — one row per person per site, naming the group their access runs through and the level it grants. Grants and membership otherwise live in separate reports, and "Site Owners has Full Control" plus "Site Owners contains five people" is not yet an answer
- Underneath it: site collection admins, web/list/item role assignments, inheritance breaks, SharePoint groups with their membership, Entra groups resolved to transitive membership, sharing links with their kind, external and guest principals, `Everyone` grants
- An item is only reported as its own scope when it has unique permissions, so the report maps the permission structure instead of repeating a row per file
- Role assignments are not readable through Graph and are not covered by SharePoint's Read/Write/Manage roles, so it creates a short-lived certificate-backed app with `Sites.FullControl.All` and deletes it again. A client secret cannot work — SharePoint Online refuses secret-based app-only tokens
- `-Excel` writes one workbook with a sheet per report plus ready-made pivots; the CSVs are always written and the workbook is built from them
- Resumes after an interruption from the last completed list, and says at the end whether every scope could actually be read

#### SharePoint Storage Report

**[Get-SharePointStorageReport.ps1](scripts/Reporting/Get-SharePointStorageReport.ps1)** — tenant-wide storage per site, library, version history and recycle bin, the longest paths measured against SharePoint's and Windows' limits, with site collection totals comparable to the admin centre.

#### SharePoint Version Cleanup

**[Remove-SharePointFileVersionsByDate.ps1](scripts/Reporting/Remove-SharePointFileVersionsByDate.ps1)** — report (and with `-Apply`, delete) file versions older than a cutoff date. The current version is always preserved.

#### Licensing Report

📂 Folder: [`Reporting/Licensing/`](scripts/Reporting/Licensing/readme.md)

Monthly licensing and Azure cost report generator.

- Combines Pax8 (CSV) and Ingram (Excel) billing data into one formatted Excel
- Per-customer tab with Azure consumption, licenses, and Acronis breakdown
- Summary tab with totals and margin per customer
- PowerShell launcher with pre-flight validation
- Optional Windows scheduled task (runs on the 6th of each month)

---

### 🖥️ Infrastructure & Devices

#### USB Setup Toolkit

📂 Folder: [`Deployment/`](scripts/Deployment/readme.md)

USB toolkit for Windows setup and Autopilot enrollment during OOBE.

- Interactive menu (Device Manager, Autopilot, AD join, device rename, product key, Windows Update, restart)
- Customer install browser from USB toolkit menu:
  - Local `Install` folder by customer (`D`)
  - Network share from `INSTALL_SHARE` in `start.local.cmd` by customer (`E`)
- For local option `D`, copy both [`Browse-InstallScripts.ps1`](scripts/Deployment/Browse-InstallScripts.ps1) and the complete `Install` folder next to [`start.bat`](scripts/Deployment/start.bat)
- Before options `D` and `E`, the toolkit creates/updates local admin `LocalAdmin` with the password from `start.local.cmd` (asked, hidden, when missing), adds it to `Administrators`, and sets OOBE skip flags
- Self-elevating, OOBE-compatible via Shift+F10
- Split "Do it all": `A` = Intune (Rename + Autopilot + Update), `C` = AD (Rename + Domain join + Update)

#### Audio Management

📂 Folder: [`Device/audio/`](scripts/Device/audio/readme.md)

Three scripts that work together to detect, disable, and roll back internal microphones on endpoints — deployed via NinjaOne.

| Script | Doel |
|---|---|
| [`detect-audiodevices.ps1`](scripts%5CDevice%5Caudio%5Cdetect-audiodevices.ps1) | Inventory van alle audio endpoints op het toestel |
| [`Disable-internalmic.ps1`](scripts%5CDevice%5Caudio%5CDisable-internalmic.ps1) | Disable interne microfoon(s), headsets worden overgeslagen |
| [`Rollback-InternalMic.ps1`](scripts%5CDevice%5Caudio%5CRollback-InternalMic.ps1) | Heractiveer eerder uitgeschakelde interne microfoons |

**NinjaOne uitrol (alle drie de scripts):**

| Instelling | Waarde |
|---|---|
| Run as | **SYSTEM** |
| Script parameters | _(geen)_ |
| Custom field vereist | `AudioDeviceInventory` (device, tekstveld/textarea) |
| Exit code | `0` = succes · `1` = fout (script gemarkeerd als mislukt) |

> Het custom field `AudioDeviceInventory` moet aangemaakt zijn als device-level custom field in NinjaOne voordat je de scripts uitrolt. De output van elk script wordt daarin weggeschreven via `Ninja-Property-Set AudioDeviceInventory`.

#### Time Sync

📂 Folder: [`Device/Time sync/`](scripts/Device/Time%20sync/readme.md)

- Restart and force Windows Time service sync

#### Windows Cleanup

📂 Folder: [`Device/`](scripts/Device/readme.md)

Comprehensive disk space cleanup for Windows endpoints.

- Cleans user/system temp folders, Windows Update download cache, Delivery Optimization cache, Prefetch, memory dumps, WER queues, thumbnail cache, DirectX shader cache, Recycle Bin, browser caches (Edge, Chrome, Firefox), and event logs
- DISM component store cleanup (`/StartComponentCleanup /ResetBase`) after Windows Updates
- DNS cache flush
- Dry-run by default — shows reclaimable space per category without deleting anything
- Run with `-Apply` to perform the actual cleanup; individual categories can be skipped with `-SkipBrowserCache`, `-SkipEventLogs`, `-SkipDism`, `-SkipRecycleBin`
- Exports a CSV report with bytes freed per category to `C:\Temp\`

#### OEM Bloatware Removal

📂 Folder: [`Device/`](scripts/Device/readme.md)

- Detects device manufacturer (HP/Lenovo/Dell) and removes known OEM bloatware via `winget`, plus a generic list of consumer Microsoft Store apps (Xbox, Solitaire, Bing News/Weather, Cortana, Clipchamp)
- Dry-run by default; `-Apply` to actually remove. CSV report of found/removed apps to `C:\Temp\`

#### Cloud Drive Mapping

📂 Folder: [`Device/DriveMapping/`](scripts/Device/DriveMapping/readme.md)

- Maps SharePoint/OneDrive document libraries to persistent drive letters via WebDAV (`net use`), for use as a per-user logon script (Intune Win32 app or scheduled task)
- Config-driven via a mappings CSV (`DriveLetter`, `Url`, optional `Label`); dry-run by default, `-Apply` to actually map
- No stored credentials — relies on the signed-in user's existing tenant session (same as browser WebDAV access)

#### Temp Disk & Pagefile (Azure / AVD)

📂 Folder: [`Device/TempDisk/`](scripts/Device/TempDisk/readme.md)

Two scripts that keep the ephemeral temp disk (`D:`) of an Azure VM or AVD session host in place, and keep the pagefile on it.

| Script | Doel |
|---|---|
| [`Init-TempDisk.ps1`](scripts/Device/TempDisk/Init-TempDisk.ps1) | Restore the temp disk as `D:` and configure the pagefile on it |
| [`Register-InitTempDiskTask.ps1`](scripts/Device/TempDisk/Register-InitTempDiskTask.ps1) | Install that script on the device and run it at every boot as SYSTEM |

- The temp disk is wiped on every deallocate, resize or host move — and Windows reads the pagefile configuration at boot, so a pagefile on a drive letter that is not there at boot is never created and the machine pages on `C:` again
- Restores the volume (RAW disks only — a disk that still carries partitions is reported, never formatted), moves an optical drive off `D:` when it is in the way, then points the pagefile at `D:\pagefile.sys` and removes the entry for every other drive
- Windows only reads that configuration at boot, so `-RestartIfNeeded` (what the boot task uses) restarts the machine once when that is the only thing left - never after a failed run, never while someone is signed in, and at most once an hour. The countdown only applies when someone is signed in to see it; at boot it restarts within seconds
- `-CheckOnly` reports without changing anything (exit code `2` = work is due); `-WhatIf` walks the whole flow; `-Quiet` keeps a healthy boot silent

#### Printers from JSON

📂 Folder: [`Device/Printer/`](scripts/Device/Printer/readme.md)

| Script | Doel |
|---|---|
| [`Install-Printer.ps1`](scripts/Device/Printer/Install-Printer.ps1) | Install printer drivers (downloaded from GitHub) and printers as described in a JSON file |

- One JSON says which drivers exist, where each is downloaded from — a GitHub release asset, a folder in a GitHub repository, any https URL or a share — and which TCP/IP printers use them. `-Printer` picks a subset
- Downloads only what is missing or older than the JSON's `version`; optional `sha256`, catalog signature check, then `pnputil /add-driver /install` + `Add-PrinterDriver`, port, printer and print defaults (duplex, colour, paper size). `"ensure": "absent"` removes a printer
- Hardened so a run cannot half-fail: the whole JSON and every INF are validated before anything changes, one run at a time per machine, a stalled spooler is restarted and the step retried, the JSON from a URL is cached as a fallback, and a clean run writes a config hash under `HKLM:\SOFTWARE\M365-Scripts\InstallPrinter` for detection
- Made for the first boot of a server or session host provisioned from a golden image: waits for the Print Spooler and retries downloads while the network comes up (`-WaitSeconds`), relaunches itself 64-bit, and is idempotent so it can also run at every startup
- `-CheckOnly` reports without changing anything (exit code `2` = work is due); `-WhatIf` walks the whole flow; `-Quiet` keeps a host that is in order silent

---

### 🔧 Custom Tools

#### SAS Batch Monitoring

📂 Folder: [`SAS/`](scripts/SAS/readme.md)

Monitor SAS batch job logs and Windows Event Viewer for errors, with optional Zabbix integration and email alerts.

- Detects spawn errors, WORK library auth failures, aborts, disk errors, and general `ERROR:` lines
- Text, JSON, and Zabbix output formats; configurable look-back period
- One-time setup script — installs to `C:\Scripts\`, creates daily scheduled task
- Optional Zabbix UserParameter config for automated alerting

#### Windows Device Management

📂 Folder: [`Device/`](scripts/Device/readme.md)

Scripts for managing and maintaining Windows devices.

**[Invoke-WindowsActivation.ps1](scripts/Device/Invoke-WindowsActivation.ps1)** — Activate Windows or manage licensing settings:
- Install a retail or KMS generic product key (`-ProductKey`)
- Configure a corporate KMS activation server (`-KmsServer`, `-KmsPort`)
- Trigger online or KMS-based activation (`-Activate`)
- Show activation status via WMI and `slmgr /dli` (`-Status`)
- Remove the product key before reimage or license transfer (`-RemoveKey`)
- Reset the grace-period counter (`-ReArm`, max ~3-5x per install)
- Confirmation prompts by default; use `-Force` to skip

**NinjaOne uitrol:**

| Instelling | Waarde |
|---|---|
| Run as | **Administrator** |
| Custom fields | _(geen — output via console/script log)_ |
| Exit code | `0` = succes · `1` = fout |

> **Let op:** `-RemoveKey` en `-ReArm` vragen interactieve bevestiging. Voeg altijd `-Force` toe wanneer je deze via NinjaOne uitvoert, anders blijft het script hangen.

Veelgebruikte NinjaOne script parameters:

| Scenario | Parameters |
|---|---|
| Status controleren | `-Status` |
| KMS activatie | `-KmsServer kms.bedrijf.local -Activate -Status` |
| KMS met afwijkende poort | `-KmsServer kms.bedrijf.local -KmsPort 2500 -Activate` |
| Retail key installeren + activeren | `-ProductKey XXXXX-XXXXX-XXXXX-XXXXX-XXXXX -Activate -Status` |
| Key verwijderen (voor reimage) | `-RemoveKey -Force` |
| Grace period resetten | `-ReArm -Force` |

**[Invoke-WindowsCleanup.ps1](scripts/Device/Invoke-WindowsCleanup.ps1)** — Scan and optionally remove reclaimable disk space:
- User + system temp, Windows Update cache, Delivery Optimization, Prefetch
- Memory dumps, WER queues, thumbnail/DirectX shader cache, font cache
- Recycle Bin, browser caches (Edge/Chrome multi-profile + Firefox)
- Event logs, DISM component store (`/StartComponentCleanup /ResetBase`)
- Application & system logs: dynamic scan of entire C:\ for `logs`/`log`/`logging` folders
- Dry-run by default; use `-Apply` to delete. Per-category summary with space freed

**[Invoke-LinuxCleanup.sh](scripts/Linux/Invoke-LinuxCleanup.sh)** — the same for a Debian/Ubuntu server, 3CX Phone System included (bash, run as root on the server; 📂 [`Linux/`](scripts/Linux/readme.md)):
- APT cache, `autoremove` (old kernels), leftover package configuration, disabled snap revisions
- systemd journal, rotated logs in `/var/log`, crash dumps, `/tmp`, user caches and trash, optionally Docker (`--docker`)
- With 3CX installed: 3CX logs, and backups beyond the newest N (`--keep-backups`); recordings are only reported, never deleted
- Dry run by default, `--apply` to delete, `--check-only` for monitoring (exit code `2` when there is work)

**[Repair-AppxPackageStore.ps1](scripts/Device/Repair-AppxPackageStore.ps1)** — Repair AppX packages (Teams, new Outlook, any other) that fail with `0x80070490` / "Deployment Register operation ... from:  (AppxManifest.xml)":
- Diagnoses registrations whose files are gone, provisioned copies without files, and orphaned `AppxAllUserStore` entries (no profile, no files, no manifest)
- On FSLogix hosts reads the `Microsoft-FSLogix-Apps` errors: which exact version the profiles ask for against what this host provisions, the FSLogix build, `InstallAppxPackages`, ODFC `IncludeTeams`, and AppX install policies
- Lists **every** app that failed to install, update or register in the last `-Days` days (AppX deployment log + FSLogix log), with the meaning of each error code
- Repairs in a fixed order — deprovision, re-register, remove, then back up every registry key to `.reg` before removing it — and reads everything back; `-Provision` puts Teams / new Outlook back for all users with Microsoft's own installer, or from winget with `-UseWinget`; `-WingetId` does the same for any other app
- `-CheckOnly` changes nothing; with `-Name '*'` system/framework packages and Deprovisioned markers are never touched

#### DNS Management

📂 Folder: [`DNS/`](scripts/DNS/readme.md)

Scripts for managing DNS records in Active Directory-integrated DNS zones.

- Resolve public DNS records via Google DNS (dig) and import them as A or CNAME records into AD DNS
- Dry-run by default — shows what would be created before applying
- Idempotent — skips records that already exist

---

### ☁️ Azure Infrastructure

📂 Folder: [`Azure/`](scripts/Azure/readme.md)

Scripts that target Azure IaaS directly via the `Az` module — not the M365 tenant, and not wired into [`menu.ps1`](menu.ps1).

- **[Azure-NVMe-Conversion.ps1](scripts/Azure/VM/Azure-NVMe-Conversion.ps1)** — vendored third-party script (Microsoft, MIT licensed, from `Azure/SAP-on-Azure-Scripts-and-Utilities`) that converts a VM's disk controller type between SCSI and NVMe, including in-guest driver readiness checks and fixes for both Windows and Linux guests
- **[Search-AADDSUserActivity.ps1](scripts/Azure/Search-AADDSUserActivity.ps1)** — searches all Azure AD Domain Services audit tables in Log Analytics for a single user in one `union` query, instead of guessing which table an event landed in

---

### 🗄️ Legacy Toolkit Rewrites

A now-retired internal PowerShell repo (and two forked third-party GitHub toolkits it vendored) was reviewed script-by-script and modernized into this repo's house style — Graph/Exchange Online instead of the retired `MSOnline`/`AzureAD` modules, dry-run-by-default with `-Apply` for anything mutating, no hardcoded customer data or secrets. None of these are wired into [`menu.ps1`](menu.ps1) — they're audit/reporting/setup scripts meant to be run directly, following the same pattern as [`scripts/RDS/`](scripts/RDS/readme.md), [`scripts/Azure/`](scripts/Azure/readme.md), and [`scripts/Network/UniFi/`](scripts/Network/UniFi/readme.md). Each folder has its own readme with full parameter/usage docs.

| Folder | Source | Covers |
|--------|--------|--------|
| [`TenantOnboarding/`](scripts/TenantOnboarding/readme.md) | Internal tenant-setup toolkit | New-tenant provisioning (break-glass admin, baseline groups/Intune assignment), multi-tenant/GDAP license + break-glass password reporting, Win32/Chocolatey app deployment, device config (kiosk power, Office uninstall, Start menu layout), OneDrive management, dynamic-DG/feature-group user management |
| [`Office365Toolkit/`](scripts/Office365Toolkit/readme.md) | Fork of [`directorcia/Office365`](https://github.com/directorcia/Office365) (CIAOPS) | Secure Score reporting, enterprise app consent cleanup, shared mailbox sign-in lockdown, EOP baseline, mailbox hygiene/forwarding-risk audits, Unified Audit Log search, Intune policy inventory |
| [`PatronToolkit/`](scripts/PatronToolkit/readme.md) | Fork of [`directorcia/patron`](https://github.com/directorcia/patron) | MFA registration + CA policy export, enterprise app consent + suspicious inbox rule + unified security alert audits, consolidated email security posture + mailbox auditing checks, SPF/DMARC validation, Intune policy assignment + Autopilot device reports, message trace, SharePoint sharing config, Teams config report |
| [`LegacyUtilities/`](scripts/LegacyUtilities/readme.md) | Internal toolkit (misc small scripts) | Mailbox folder permissions/delegate access, bulk shared mailbox/contact creation, contact sync, duplicate mail item cleanup, M365 group membership, CA policy backup, Teams/Planner cloning, Azure Files drive mapping, NumLock/lock-workstation device tweaks, Workspace 365 environment provisioning |

Both GitHub forks were reviewed capability-by-capability rather than ported 1:1 — near-duplicate single-purpose report scripts were consolidated into fewer well-parameterized ones, and capabilities already covered elsewhere in this repo were skipped rather than duplicated (see each folder's readme for the full skip list and reasoning). All code is a fresh implementation in this repo's style, not copied from the source projects.

---

## Repository Structure

Every folder has its own [`readme.md`](readme.md) — this tree is a map; follow the links for full parameter/usage docs.

<pre>
<a href="readme.md">M365-Scripts/</a>
├── <a href=".github/workflows/ci.yml">.github/workflows/ci.yml</a>         ← CI: syntax, analyzer, links, generated docs, ShellCheck
├── <a href=".gitignore">.gitignore</a>
├── <a href=".vscode">.vscode/</a>
│   └── <a href=".vscode/settings.json">settings.json</a>
├── <a href="load.ps1">load.ps1</a>                         ← Entry point: first-run setup + launches menu
├── <a href="menu.ps1">menu.ps1</a>                         ← Interactive launcher (all scripts + M365 functions)
├── <a href="readme.md">readme.md</a>
└── <a href="scripts/readme.md">scripts/</a>
    ├── <a href="scripts/readme.md">readme.md</a>                    ← Index of all categories below
    ├── <a href="scripts/INDEX.md">INDEX.md</a>                     ← Every script A-Z with its folder (generated)
    ├── <a href="scripts/Azure/readme.md">Azure/</a>                        ← targets Azure IaaS directly via Az, not the M365 tenant
    │   ├── <a href="scripts/Azure/readme.md">readme.md</a>
    │   └── <a href="scripts/Azure/VM/readme.md">VM/</a>
    │       ├── <a href="scripts/Azure/VM/readme.md">readme.md</a>
    │       └── <a href="scripts/Azure/VM/Azure-NVMe-Conversion.ps1">Azure-NVMe-Conversion.ps1</a>   ← vendored (Microsoft, MIT) — SCSI/NVMe disk controller conversion
    ├── <a href="scripts/Entra/readme.md">Entra/</a>
    │   ├── <a href="scripts/Entra/readme.md">readme.md</a>
    │   ├── <a href="scripts/Entra/Set-UserManager.ps1">Set-UserManager.ps1</a>
    │   ├── <a href="scripts/Entra/Remove-M365Users.ps1">Remove-M365Users.ps1</a>
    │   ├── <a href="scripts/Entra/New-M365User.ps1">New-M365User.ps1</a>
    │   ├── <a href="scripts/Entra/Import-M365Users.ps1">Import-M365Users.ps1</a>
    │   ├── <a href="scripts/Entra/Get-M365UserLicenses.ps1">Get-M365UserLicenses.ps1</a>
    │   ├── <a href="scripts/Entra/Import-ConditionalAccessBaseline.ps1">Import-ConditionalAccessBaseline.ps1</a>
    │   └── <a href="scripts/Entra/Test-M365GroupMembership.ps1">Test-M365GroupMembership.ps1</a>   ← audit M365 Group / Teams owners and members
    ├── <a href="scripts/Exchange/readme.md">Exchange/</a>
    │   ├── <a href="scripts/Exchange/readme.md">readme.md</a>
    │   ├── <a href="scripts/Exchange/Migrate-Calendar.ps1">Migrate-Calendar.ps1</a>
    │   ├── <a href="scripts/Exchange/Convert-SharedCalendarToResource.ps1">Convert-SharedCalendarToResource.ps1</a>  ← shared calendar in a user's mailbox → its own room/equipment mailbox
    │   ├── <a href="scripts/Exchange/Move-SharedCalendar.ps1">Move-SharedCalendar.ps1</a>  ← all in one: search by keyword + convert + who has to switch
    │   ├── <a href="scripts/Exchange/Set-Calendar-rights.ps1">Set-Calendar-rights.ps1</a>
    │   ├── <a href="scripts/Exchange/Set-Distributionlist-dynamic-static.ps1">Set-Distributionlist-dynamic-static.ps1</a>
    │   ├── <a href="scripts/Exchange/Move-InboxToArchive.ps1">Move-InboxToArchive.ps1</a>
    │   ├── <a href="scripts/Exchange/Restore-MailboxMessages.ps1">Restore-MailboxMessages.ps1</a>  ← put back mail moved/deleted on a date, and who did it
    │   ├── <a href="scripts/Exchange/Test-CalendarPermissions.ps1">Test-CalendarPermissions.ps1</a>
    │   ├── <a href="scripts/Exchange/Get-CalendarMappings.ps1">Get-CalendarMappings.ps1</a>  ← where each calendar is mapped in Outlook, next to the rights behind it
    │   ├── <a href="scripts/Exchange/Test-MailboxPermissions.ps1">Test-MailboxPermissions.ps1</a>
    │   ├── <a href="scripts/Exchange/Test-DistributionGroupPermissions.ps1">Test-DistributionGroupPermissions.ps1</a>
    │   ├── <a href="scripts/Exchange/Test-DkimConfig.ps1">Test-DkimConfig.ps1</a>
    │   ├── <a href="scripts/Exchange/Get-ExternalForwards.ps1">Get-ExternalForwards.ps1</a>
    │   ├── <a href="scripts/Exchange/Get-MailboxSizes.ps1">Get-MailboxSizes.ps1</a>
    │   └── <a href="scripts/Exchange/Get-DistributionGroupMembers.ps1">Get-DistributionGroupMembers.ps1</a>  ← who is on which distribution list, as an Excel workbook for the customer
    ├── <a href="scripts/Graph/readme.md">Graph/</a>
    │   ├── <a href="scripts/Graph/readme.md">readme.md</a>
    │   └── <a href="scripts/Graph/logic-permissies.ps1">logic-permissies.ps1</a>     ← grant a Graph app role to a Logic App managed identity
    ├── <a href="scripts/Intune/readme.md">Intune/</a>
    │   ├── <a href="scripts/Intune/readme.md">readme.md</a>
    │   ├── <a href="scripts/Intune/Compare-IntuneConfig.ps1">Compare-IntuneConfig.ps1</a>  ← Intune config drift vs an MSP baseline backup (IntuneBackupAndRestore)
    │   ├── <a href="scripts/Intune/Get-Autopilot/readme.md">Get-Autopilot/</a>
    │   │   ├── <a href="scripts/Intune/Get-Autopilot/readme.md">readme.md</a>
    │   │   ├── <a href="scripts/Intune/Get-Autopilot/Get-WindowsAutoPilotInfo.ps1">Get-WindowsAutoPilotInfo.ps1</a>
    │   │   └── <a href="scripts/Intune/Get-Autopilot/GetAutoPilot.CMD">GetAutoPilot.CMD</a>
    │   ├── <a href="scripts/Intune/iOS-Compliance-Updater/readme.md">iOS-Compliance-Updater/</a>
    │   │   ├── <a href="scripts/Intune/iOS-Compliance-Updater/readme.md">readme.md</a>
    │   │   ├── <a href="scripts/Intune/iOS-Compliance-Updater/Update-iOSCompliancePolicy.ps1">Update-iOSCompliancePolicy.ps1</a>   ← main script (run or scheduled task)
    │   │   ├── <a href="scripts/Intune/iOS-Compliance-Updater/Setup.ps1">Setup.ps1</a>                        ← one-time: App Registration + config.json
    │   │   ├── <a href="scripts/Intune/iOS-Compliance-Updater/Install-ScheduledTask.ps1">Install-ScheduledTask.ps1</a>        ← register weekly scheduled task
    │   │   └── <a href="scripts/Intune/iOS-Compliance-Updater/config.example.json">config.example.json</a>
    │   └── <a href="scripts/Intune/Desktop/readme.md">Desktop/</a>                  ← corporate wallpaper/lockscreen (Office theme lives in Custom Scripts/, see below)
    │       ├── <a href="scripts/Intune/Desktop/readme.md">readme.md</a>
    │       ├── <a href="scripts/Intune/Desktop/Add%20Lockscreen%20to%20start%20and%20desktop/readme.md">Add Lockscreen to start and desktop/</a>
    │       │   ├── <a href="scripts/Intune/Desktop/Add%20Lockscreen%20to%20start%20and%20desktop/readme.md">readme.md</a>
    │       │   ├── <a href="scripts/Intune/Desktop/Add%20Lockscreen%20to%20start%20and%20desktop/add-lock.ps1">add-lock.ps1</a>               ← taskbar "Lock Workstation" shortcut
    │       │   └── <a href="scripts/Intune/Desktop/Add%20Lockscreen%20to%20start%20and%20desktop/add-shortcut-lock.ps1">add-shortcut-lock.ps1</a>
    │       └── <a href="scripts/Intune/Desktop/Background/readme.md">Background/</a>
    │           ├── <a href="scripts/Intune/Desktop/Background/readme.md">readme.md</a>
    │           ├── <a href="scripts/Intune/Desktop/Background/Desktop/readme.md">Desktop/</a>
    │           │   ├── <a href="scripts/Intune/Desktop/Background/Desktop/readme.md">readme.md</a>
    │           │   ├── <a href="scripts/Intune/Desktop/Background/Desktop/Set-CorporateWallpaper.ps1">Set-CorporateWallpaper.ps1</a>  ← corporate wallpaper via Intune (hash check, PersonalizationCSP)
    │           │   └── <a href="scripts/Intune/Desktop/Background/Desktop/Remove-CorporateWallpaper.ps1">Remove-CorporateWallpaper.ps1</a>
    │           └── <a href="scripts/Intune/Desktop/Background/Lockscreen/readme.md">Lockscreen/</a>
    │               ├── <a href="scripts/Intune/Desktop/Background/Lockscreen/readme.md">readme.md</a>
    │               └── <a href="scripts/Intune/Desktop/Background/Lockscreen/Make-lockscreen.ps1">Make-lockscreen.ps1</a>         ← corporate lockscreen via Intune (validated download, PersonalizationCSP)
    ├── <a href="scripts/Device/readme.md">Device/</a>
    │   ├── <a href="scripts/Device/readme.md">readme.md</a>
    │   ├── <a href="scripts/Device/Invoke-WindowsActivation.ps1">Invoke-WindowsActivation.ps1</a> ← activate Windows, set product key / KMS server
    │   ├── <a href="scripts/Device/Invoke-WindowsCleanup.ps1">Invoke-WindowsCleanup.ps1</a>    ← temp, cache, WU, DISM, browser, event logs
    │   ├── <a href="scripts/Device/Clear-TempFiles.ps1">Clear-TempFiles.ps1</a>
    │   ├── <a href="scripts/Device/Remove-OemBloatware.ps1">Remove-OemBloatware.ps1</a>      ← HP/Lenovo/Dell + generic Store bloatware removal
    │   ├── <a href="scripts/Device/Repair-AppxPackageStore.ps1">Repair-AppxPackageStore.ps1</a>  ← repair AppX 0x80070490 (orphaned store entries, FSLogix replay)
    │   ├── <a href="scripts/Device/Test-OpenVpnDiagnostics.ps1">Test-OpenVpnDiagnostics.ps1</a>  ← OpenVPN Connect diagnostics
    │   ├── <a href="scripts/Device/Update-TeamsClient.ps1">Update-TeamsClient.ps1</a>       ← update new Teams + meeting add-in when a newer build exists
    │   ├── <a href="scripts/Device/Update-TeamsClient.md">Update-TeamsClient.md</a>        ← how that script decides, step by step
    │   ├── <a href="scripts/Device/Update-TeamsClient-ITGlue.md">Update-TeamsClient-ITGlue.md</a> ← servicedeskversie (NL) om in IT Glue te plakken
    │   ├── <a href="scripts/Device/audio/readme.md">audio/</a>
    │   │   ├── <a href="scripts/Device/audio/readme.md">readme.md</a>
    │   │   ├── <a href="scripts/Device/audio/detect-audiodevices.ps1">detect-audiodevices.ps1</a>
    │   │   ├── <a href="scripts/Device/audio/Disable-internalmic.ps1">Disable-internalmic.ps1</a>
    │   │   └── <a href="scripts/Device/audio/Rollback-InternalMic.ps1">Rollback-InternalMic.ps1</a>
    │   ├── <a href="scripts/Device/DriveMapping/readme.md">DriveMapping/</a>
    │   │   ├── <a href="scripts/Device/DriveMapping/readme.md">readme.md</a>
    │   │   └── <a href="scripts/Device/DriveMapping/New-CloudDriveMapping.ps1">New-CloudDriveMapping.ps1</a>   ← map SharePoint/OneDrive libraries to drive letters (WebDAV)
    │   ├── <a href="scripts/Device/Printer/readme.md">Printer/</a>
    │   │   ├── <a href="scripts/Device/Printer/readme.md">readme.md</a>
    │   │   ├── <a href="scripts/Device/Printer/Install-Printer.ps1">Install-Printer.ps1</a>            ← printer drivers (from GitHub) + printers from a JSON file
    │   │   └── <a href="scripts/Device/Printer/printers.example.json">printers.example.json</a>          ← every field and every driver source
    │   ├── <a href="scripts/Device/TempDisk/readme.md">TempDisk/</a>
    │   │   ├── <a href="scripts/Device/TempDisk/readme.md">readme.md</a>
    │   │   ├── <a href="scripts/Device/TempDisk/Init-TempDisk.ps1">Init-TempDisk.ps1</a>              ← restore the ephemeral temp disk as D: and put the pagefile on it
    │   │   └── <a href="scripts/Device/TempDisk/Register-InitTempDiskTask.ps1">Register-InitTempDiskTask.ps1</a>  ← install that script and run it at every boot as SYSTEM
    │   └── <a href="scripts/Device/Time%20sync/readme.md">Time sync/</a>
    │       ├── <a href="scripts/Device/Time%20sync/readme.md">readme.md</a>
    │       └── <a href="scripts/Device/Time%20sync/Restart-Time-Sync.ps1">Restart-Time-Sync.ps1</a>
    ├── <a href="scripts/Linux/readme.md">Linux/</a>
    │   ├── <a href="scripts/Linux/readme.md">readme.md</a>
    │   └── <a href="scripts/Linux/Invoke-LinuxCleanup.sh">Invoke-LinuxCleanup.sh</a>        ← bash: disk cleanup for Debian/Ubuntu, 3CX included
    ├── <a href="scripts/Network/readme.md">Network/</a>
    │   ├── <a href="scripts/Network/readme.md">readme.md</a>
    │   ├── <a href="scripts/Network/Test-Ports.ps1">Test-Ports.ps1</a>
    │   ├── <a href="scripts/Network/Test-AuthNetworkDiagnostics.ps1">Test-AuthNetworkDiagnostics.ps1</a>   ← auth/network issue diagnostics
    │   ├── <a href="scripts/Network/Test-FileIODiagnostics.ps1">Test-FileIODiagnostics.ps1</a>        ← file I/O test + real-time directory monitor
    │   └── <a href="scripts/Network/UniFi/readme.md">UniFi/</a>
    │       ├── <a href="scripts/Network/UniFi/readme.md">readme.md</a>
    │       ├── <a href="scripts/Network/UniFi/UnifiApi.ps1">UnifiApi.ps1</a>                  ← shared login/session helper (dot-sourced)
    │       ├── <a href="scripts/Network/UniFi/Get-UnifiNetworkReport.ps1">Get-UnifiNetworkReport.ps1</a>     ← HTML network documentation report
    │       └── <a href="scripts/Network/UniFi/Update-UnifiFirmware.ps1">Update-UnifiFirmware.ps1</a>       ← list/trigger firmware upgrades across sites
    ├── <a href="scripts/RDS/readme.md">RDS/</a>
    │   ├── <a href="scripts/RDS/readme.md">readme.md</a>
    │   ├── <a href="scripts/RDS/Get-FSlogix-errors.ps1">Get-FSlogix-errors.ps1</a>            ← FSLogix / Azure Files profile diagnostics
    │   ├── <a href="scripts/RDS/Invoke-FSLogixShrink.ps1">Invoke-FSLogixShrink.ps1</a>          ← shrink FSLogix profile disks, check compaction
    │   ├── <a href="scripts/RDS/Test-RDSDiagnostics.ps1">Test-RDSDiagnostics.ps1</a>           ← RDP/RDWeb login failure diagnostics
    │   ├── <a href="scripts/RDS/Update-SessionHostImage.ps1">Update-SessionHostImage.ps1</a>       ← Teams / Outlook / Copilot fit for FSLogix on the image
    │   └── <a href="scripts/RDS/Watch-RDSLive.ps1">Watch-RDSLive.ps1</a>                 ← real-time session + licensing monitor
    ├── <a href="scripts/SMTP/readme.md">SMTP/</a>
    │   ├── <a href="scripts/SMTP/readme.md">readme.md</a>
    │   ├── <a href="scripts/SMTP/testsmtp.ps1">testsmtp.ps1</a>
    │   └── <a href="scripts/SMTP/testsmtp_5min.ps1">testsmtp_5min.ps1</a>
    ├── <a href="scripts/Deployment/readme.md">Deployment/</a>                   ← USB setup toolkit (OOBE / Autopilot)
    │   ├── <a href="scripts/Deployment/readme.md">readme.md</a>
    │   ├── <a href="scripts/Deployment/start.bat">start.bat</a>
    │   ├── <a href="scripts/Deployment/autorun.inf">autorun.inf</a>
    │   └── <a href="scripts/Deployment/Browse-InstallScripts.ps1">Browse-InstallScripts.ps1</a>
    ├── <a href="scripts/DNS/readme.md">DNS/</a>
    │   ├── <a href="scripts/DNS/readme.md">readme.md</a>
    │   ├── <a href="scripts/DNS/Import-DnsRecords.ps1">Import-DnsRecords.ps1</a>   ← resolve via Google DNS + import into AD DNS
    │   └── <a href="scripts/DNS/example-records.csv">example-records.csv</a>
    ├── <a href="scripts/SAS/readme.md">SAS/</a>
    │   ├── <a href="scripts/SAS/readme.md">readme.md</a>
    │   ├── <a href="scripts/SAS/rca.md">rca.md</a>
    │   ├── <a href="scripts/SAS/Monitor-SASBatchErrors.ps1">Monitor-SASBatchErrors.ps1</a>   ← scan logs + Event Viewer for SAS errors
    │   ├── <a href="scripts/SAS/Setup-SASMonitoring.ps1">Setup-SASMonitoring.ps1</a>      ← install script, scheduled task, Zabbix config
    │   ├── <a href="scripts/SAS/Test-SASWorkDirectory.ps1">Test-SASWorkDirectory.ps1</a>    ← validate WORK directory health
    │   └── <a href="scripts/SAS/zabbix_sas_monitor.conf">zabbix_sas_monitor.conf</a>
    ├── <a href="scripts/SharePoint/readme.md">SharePoint/</a>
    │   ├── <a href="scripts/SharePoint/readme.md">readme.md</a>
    │   ├── <a href="scripts/SharePoint/Find-SiteContent.ps1">Find-SiteContent.ps1</a>         ← search a whole site (name/path/type/date or full text) + report the permissions on every hit (PnP)
    │   ├── <a href="scripts/SharePoint/Search-SharePointContent.ps1">Search-SharePointContent.ps1</a> ← same, tenant-wide via Graph app-only: delta + /permissions, sharing links and guests (files/folders)
    │   ├── <a href="scripts/SharePoint/Restore-RecycleBinItems.ps1">Restore-RecycleBinItems.ps1</a>  ← restore deleted files from a recycle bin: one site/OneDrive or tenant-wide (PnP, auto app registration)
    │   ├── <a href="scripts/SharePoint/Trace-SharePointFile.ps1">Trace-SharePointFile.ps1</a>     ← where did a file go: renames, moves, copies, deletes (incl. via a folder) from the audit log, in Brussels time
    │   ├── <a href="scripts/SharePoint/Revoke-SharePointUserAccess.ps1">Revoke-SharePointUserAccess.ps1</a> ← take one user's access away at every level, sharing links included (reports unless -Apply)
    │   ├── <a href="scripts/SharePoint/Test-SharePointAccessScripts.ps1">Test-SharePointAccessScripts.ps1</a> ← verify the two access scripts without a tenant (shared auth block + revocation funnel)
    │   └── <a href="scripts/SharePoint/Provisioning/readme.md">Provisioning/</a>                ← provision a whole structure from one JSON config (PnP + Graph)
    │       ├── <a href="scripts/SharePoint/Provisioning/readme.md">readme.md</a>
    │       ├── <a href="scripts/SharePoint/Provisioning/SharePoint-Handleiding.md">SharePoint-Handleiding.md</a>      ← end-user guide (NL) to hand to the customer (template)
    │       ├── <a href="scripts/SharePoint/Provisioning/example.config.json">example.config.json</a>          ← example model (CHANGEME) — client configs next to it are git-ignored
    │       ├── <a href="scripts/SharePoint/Provisioning/Install-SharePointStructure.ps1">Install-SharePointStructure.ps1</a> ← build it all in one run, incl. (temporary) app registration
    │       ├── <a href="scripts/SharePoint/Provisioning/SharePointStructure.Common.ps1">SharePointStructure.Common.ps1</a> ← shared helpers (dot-sourced by every script here)
    │       ├── <a href="scripts/SharePoint/Provisioning/New-SharePointMetadata.ps1">New-SharePointMetadata.ps1</a>   ← term set, site columns, content types (every site in the config)
    │       ├── <a href="scripts/SharePoint/Provisioning/Set-SharePointLibraries.ps1">Set-SharePointLibraries.ps1</a>  ← libraries/channel folders, content types, defaults, views, group rights
    │       ├── <a href="scripts/SharePoint/Provisioning/Update-SharePointShareStatus.ps1">Update-SharePointShareStatus.ps1</a> ← derive the Deelstatus column, flag over-sharing (exit 2)
    │       └── <a href="scripts/SharePoint/Provisioning/Test-SharePointStructure.ps1">Test-SharePointStructure.ps1</a> ← read-only drift check vs the config (exit 2)
    ├── <a href="scripts/Teams/readme.md">Teams/</a>
    │   ├── <a href="scripts/Teams/readme.md">readme.md</a>
    │   └── <a href="scripts/Teams/Invoke-TeamsArchive.ps1">Invoke-TeamsArchive.ps1</a>  ← Teams/SharePoint export + archiving (Graph, PS7+, Global Admin)
    ├── <a href="scripts/Reporting/readme.md">Reporting/</a>
    │   ├── <a href="scripts/Reporting/readme.md">readme.md</a>
    │   ├── <a href="scripts/Reporting/Get-ComputerLastLogon.ps1">Get-ComputerLastLogon.ps1</a>        ← last logon per computer in OU(s), export to CSV
    │   ├── <a href="scripts/Reporting/Get-SharePointStorageReport.ps1">Get-SharePointStorageReport.ps1</a>  ← tenant-wide SharePoint storage report
    │   ├── <a href="scripts/Reporting/Get-SharePointPermissionsReport.ps1">Get-SharePointPermissionsReport.ps1</a> ← who has access to what and via which group, to CSV + Excel
    │   ├── <a href="scripts/Reporting/Remove-SharePointFileVersionsByDate.ps1">Remove-SharePointFileVersionsByDate.ps1</a> ← delete file versions older than a date
    │   └── <a href="scripts/Reporting/Licensing/readme.md">Licensing/</a>
    │       ├── <a href="scripts/Reporting/Licensing/readme.md">readme.md</a>
    │       ├── <a href="scripts/Reporting/Licensing/genereer_licentie_overzicht.py">genereer_licentie_overzicht.py</a>
    │       ├── <a href="scripts/Reporting/Licensing/genereer_rapport.ps1">genereer_rapport.ps1</a>
    │       ├── <a href="scripts/Reporting/Licensing/genereer_rapport.bat">genereer_rapport.bat</a>
    │       └── <a href="scripts/Reporting/Licensing/create_scheduled_task.ps1">create_scheduled_task.ps1</a>
    ├── <a href="scripts/Startup/readme.md">Startup/</a>
    │   ├── <a href="scripts/Startup/readme.md">readme.md</a>
    │   ├── <a href="scripts/Startup/functies.ps1">functies.ps1</a>             ← M365 function library (dot-sourced by menu)
    │   ├── <a href="scripts/Startup/RequiredModules.psd1">RequiredModules.psd1</a>     ← The one list of required modules
    │   ├── <a href="scripts/Startup/Test-RequiredModules.ps1">Test-RequiredModules.ps1</a> ← Report modules scripts load that the list misses
    │   ├── <a href="scripts/Startup/Install-Modules.ps1">Install-Modules.ps1</a>      ← Bootstrap: install &amp; import all modules
    │   ├── <a href="scripts/Startup/Update-Modules.ps1">Update-Modules.ps1</a>       ← Check/update the required modules (load.ps1 runs it at startup), then the rest
    │   ├── <a href="scripts/Startup/Test-PowerShellSyntax.ps1">Test-PowerShellSyntax.ps1</a>
    │   ├── <a href="scripts/Startup/Update-ScriptIndex.ps1">Update-ScriptIndex.ps1</a>   ← Regenerates scripts/INDEX.md from the .SYNOPSIS headers
    │   ├── <a href="scripts/Startup/Test-MarkdownLinks.ps1">Test-MarkdownLinks.ps1</a>   ← Checks every readme link: files and in-page anchors
    │   └── <a href="scripts/Startup/Convert-MarkdownToHtml.ps1">Convert-MarkdownToHtml.ps1</a> ← Markdown doc → one self-contained styled HTML page
    ├── <a href="scripts/Custom%20Scripts/readme.md">Custom Scripts/</a>                 ← Office theme/color deployment, theme URL as a parameter
    │   ├── <a href="scripts/Custom%20Scripts/readme.md">readme.md</a>
    │   └── <a href="scripts/Custom%20Scripts/Intune/readme.md">Intune/</a>
    │       ├── <a href="scripts/Custom%20Scripts/Intune/readme.md">readme.md</a>
    │       └── <a href="scripts/Custom%20Scripts/Intune/Desktop/readme.md">Desktop/</a>
    │           ├── <a href="scripts/Custom%20Scripts/Intune/Desktop/readme.md">readme.md</a>
    │           ├── <a href="scripts/Custom%20Scripts/Intune/Desktop/Deploy-OfficeTheme.ps1">Deploy-OfficeTheme.ps1</a>        ← installs an Office .thmx theme from a URL
    │           └── <a href="scripts/Custom%20Scripts/Intune/Desktop/Office%20Themes/readme.md">Office Themes/</a>
    │               ├── <a href="scripts/Custom%20Scripts/Intune/Desktop/Office%20Themes/readme.md">readme.md</a>
    │               └── <a href="scripts/Custom%20Scripts/Intune/Desktop/Office%20Themes/Deploy-Officecolors.ps1">Deploy-Officecolors.ps1</a>   ← installs just a color scheme from a URL
    ├── <a href="scripts/TenantOnboarding/readme.md">TenantOnboarding/</a>                ← modernized from a retired internal tenant-setup toolkit, not menu-wired
    │   ├── <a href="scripts/TenantOnboarding/readme.md">readme.md</a>
    │   ├── <a href="scripts/TenantOnboarding/Provisioning/readme.md">Provisioning/</a>         (3 scripts)  ← break-glass admin, baseline groups, Intune policy assignment
    │   ├── <a href="scripts/TenantOnboarding/MultiTenant/readme.md">MultiTenant/</a>          (3 scripts)  ← GDAP license report, break-glass password rotation, customer portal index
    │   ├── <a href="scripts/TenantOnboarding/AppDeployment/readme.md">AppDeployment/</a>        (7 scripts)  ← Win32/Chocolatey install, shortcuts, file associations, printer connections
    │   ├── <a href="scripts/TenantOnboarding/DeviceConfig/readme.md">DeviceConfig/</a>         (6 scripts)  ← kiosk power, Office uninstall, Start menu layout, Teams firewall rule
    │   ├── <a href="scripts/TenantOnboarding/OneDriveManagement/readme.md">OneDriveManagement/</a>   (3 scripts)  ← sync watchdog, library sync stop, known-folder redirect
    │   └── <a href="scripts/TenantOnboarding/UserManagement/readme.md">UserManagement/</a>       (2 scripts)  ← dynamic DG by filter, feature-group membership
    ├── <a href="scripts/Office365Toolkit/readme.md">Office365Toolkit/</a>                ← rewrite of retired directorcia/Office365 (CIAOPS) fork, not menu-wired
    │   ├── <a href="scripts/Office365Toolkit/readme.md">readme.md</a>
    │   ├── <a href="scripts/Office365Toolkit/Security/readme.md">Security/</a>             (4 scripts)  ← Secure Score, app consent cleanup, shared mailbox lockdown, EOP baseline
    │   ├── <a href="scripts/Office365Toolkit/Exchange/readme.md">Exchange/</a>             (4 scripts)  ← mailbox hygiene baseline, forwarding risk, add-ins, audit log search
    │   └── <a href="scripts/Office365Toolkit/Intune/readme.md">Intune/</a>               (1 script)   ← tenant-wide policy inventory
    ├── <a href="scripts/PatronToolkit/readme.md">PatronToolkit/</a>                    ← rewrite of retired directorcia/patron fork, not menu-wired
    │   ├── <a href="scripts/PatronToolkit/readme.md">readme.md</a>
    │   ├── <a href="scripts/PatronToolkit/Entra/readme.md">Entra/</a>                (2 scripts)  ← MFA registration report, CA policy export
    │   ├── <a href="scripts/PatronToolkit/Security/readme.md">Security/</a>             (5 scripts)  ← app consents, suspicious inbox rules, security alerts, email security posture, mailbox auditing
    │   ├── <a href="scripts/PatronToolkit/Exchange/readme.md">Exchange/</a>             (1 script)   ← message trace report
    │   ├── <a href="scripts/PatronToolkit/Intune/readme.md">Intune/</a>               (2 scripts)  ← policy assignments, Autopilot devices
    │   ├── <a href="scripts/PatronToolkit/SharePoint/readme.md">SharePoint/</a>           (1 script)   ← sharing config audit
    │   └── <a href="scripts/PatronToolkit/Teams/readme.md">Teams/</a>                (1 script)   ← Teams config report
    └── <a href="scripts/LegacyUtilities/readme.md">LegacyUtilities/</a>                  ← misc modernized scripts from the retired internal toolkit, not menu-wired
        ├── <a href="scripts/LegacyUtilities/readme.md">readme.md</a>
        ├── <a href="scripts/LegacyUtilities/Exchange/readme.md">Exchange/</a>             (7 scripts)  ← folder permissions, delegate access, bulk mailboxes/contacts, contact sync, dedup, message trace
        ├── <a href="scripts/LegacyUtilities/Entra/readme.md">Entra/</a>                (2 scripts)  ← group membership, CA policy backup
        ├── <a href="scripts/LegacyUtilities/Teams/readme.md">Teams/</a>                (3 scripts)  ← team/plan cloning, project team provisioning
        ├── <a href="scripts/LegacyUtilities/Network/readme.md">Network/</a>              (1 script)   ← Azure Files drive mapping
        ├── <a href="scripts/LegacyUtilities/Device/readme.md">Device/</a>               (2 scripts)  ← NumLock default, lock-workstation shortcut
        └── <a href="scripts/LegacyUtilities/Workspace365/readme.md">Workspace365/</a>         (2 scripts)  ← environment provisioning/removal
</pre>

[`Deploy-OfficeTheme.ps1`](scripts/Custom%20Scripts/Intune/Desktop/Deploy-OfficeTheme.ps1) and [`Deploy-Officecolors.ps1`](scripts/Custom%20Scripts/Intune/Desktop/Office%20Themes/Deploy-Officecolors.ps1) take the theme's download URL as a parameter; the theme files are not kept in the repo.

---

## Contributing

When adding new scripts:

1. Follow the existing naming convention (`Verb-Noun.ps1`)
2. Include a header comment block with Synopsis, Description, Parameters, and Example
3. Test against a non-production tenant before committing
4. Place the script in the appropriate workload folder
5. Add it to [`menu.ps1`](menu.ps1) and update this readme
6. Update `Version History` in this file for every functional or structural change (required), including changes requested or applied via Copilot/AI assistant
7. Write the change in all three readme languages ([`readme.md`](readme.md), [`readme.nl.md`](readme.nl.md), [`readme.fr.md`](readme.fr.md))

**What updates itself.** [`scripts/INDEX.md`](scripts/INDEX.md), the language switcher and breadcrumb at the top of every readme, and the link check are kept current automatically — you do not run them by hand:

| When | What runs |
|------|-----------|
| You commit | The git hook [`.githooks/pre-commit`](.githooks/pre-commit) regenerates the index and the readme headers, adds them to the commit, and stops the commit on a broken link. It warns when an English readme changed without its Dutch/French version — ask Claude to translate |
| Claude Code edits a file | A hook in [`.claude/settings.json`](.claude/settings.json) does the same in the background after every edit, and before Claude finishes it checks that changed readmes were translated and changed scripts were documented |
| You push to `devel`/`main` or open a PR | [GitHub Actions](.github/workflows/ci.yml) runs the same checks once more, for commits made without the hook or in the web editor: PowerShell syntax (7 and 5.1), PSScriptAnalyzer errors, the link check, whether `INDEX.md` and the readme headers are current and every folder has all three languages, and ShellCheck on the `.sh` scripts |

Turn the git hook on once per clone:

```powershell
git config core.hooksPath .githooks
```

---

## Disclaimer

These scripts are provided as-is. Always test in a non-production environment before running against live tenants. The maintainer accepts no liability for unintended changes resulting from misuse or misconfiguration.

---

## Version History

> Note: Older entries can reference historical folder names such as [`Custom Scripts/`](scripts/Custom%20Scripts/readme.md) and `Testing Scripts/`. These path names reflect the repository structure at the time of that change.

### 2026-10-08 (11)
| Change |
|--------|
| **[`PatronToolkit/`](scripts/PatronToolkit/readme.md) signs in through [`Connect-M365.ps1`](scripts/Startup/readme.md#connect-m365ps1), and drops the last SharePoint Online Management Shell code.** All 13 scripts: delegated by default (device code and GDAP customer from `load.config.ps1`), new `-ClientId`, `-CertificateThumbprint` and `-AppOnly`, and they close only what they opened. [`Test-SharePointSharingConfig.ps1`](scripts/PatronToolkit/SharePoint/readme.md#test-sharepointsharingconfigps1) used `Connect-SPOService` and ignored its documented `-TenantId`; tenant sharing now comes from Graph `/admin/sharepoint/settings`, and only what Graph has no API for (link defaults with the new `-IncludeLinkSettings`, site overrides, external users) uses PnP, skipped with a warning when PnP cannot sign in. [`Test-EmailSecurityPosture.ps1`](scripts/PatronToolkit/Security/readme.md#test-emailsecuritypostureps1) called `Connect-IPPSSession` without parameters and read your own tenant under GDAP |
| **Results that were wrong or empty.** [`Get-IntunePolicyAssignments.ps1`](scripts/PatronToolkit/Intune/readme.md#get-intunepolicyassignmentsps1) and [`Get-AutopilotDevices.ps1`](scripts/PatronToolkit/Intune/readme.md#get-autopilotdevicesps1) called cmdlets that do not exist in Microsoft.Graph v2, hidden by `SilentlyContinue`: Settings Catalog assignments and deployment profiles were always missing. Both now use Graph with paging, and Autopilot no longer counts every device as unassigned (Graph says `assignedInSync`, not `assigned`). [`Get-MessageTraceReport.ps1`](scripts/PatronToolkit/Exchange/readme.md#get-messagetracereportps1) stopped at 5000 rows and used the retired `Get-MessageTraceDetail`; it now pages V2 in 10-day slices up to 90 days (new `-MaxResults`). DMARC parsing in `Test-EmailAuthenticationRecords` could read `sp=` as the policy; `Get-SuspiciousInboxRules` called internal `EX:` recipients external; alert and consent reports sorted severity alphabetically; `Get-TeamsConfigReport` read an anonymous-join property its cmdlet does not have, and now takes the team inventory from Graph. `Get-SuspiciousInboxRules` deliberately stays on Exchange: Graph `messageRules` never returns hidden rules, which are what the script hunts |
| Verified: syntax check; link check; every Graph 2.41.1 and PnP 3.4.1 cmdlet and parameter used exists, and the three removed Graph cmdlets do not; unit tests of the DMARC regex, the internal/external address check and the severity sort; help blocks matched to the parameters. **Not** verified: any run against a tenant - `$expand=assignments` on all Intune collections, the `/admin/sharepoint/settings` values, V2 paging beyond 5000 rows, PnP sign-in and the Teams configuration property names |

### 2026-10-08 (10)
| Change |
|--------|
| **[`Connect-M365.ps1`](scripts/Startup/readme.md#connect-m365ps1): Exchange sessions closed by connection id, and what the first scripts on it asked for.** `Disconnect-M365Exchange` ran `Disconnect-ExchangeOnline` without `-ConnectionId`, so a script that only added a Security & Compliance session - or a child script called by another - closed its caller's Exchange session too. `Connect-M365Exchange` now returns the ids of the sessions it opened and only those are closed. A delegated Exchange sign-in only passes `-DelegatedOrganization` under GDAP; outside GDAP it signed a home-tenant admin in as a "partner" of their own tenant. New: `-EnableSearchOnlySession` (Content Search purges), `Disconnect-M365Teams`, `Connect-M365PnP -AppOnly` (the certificate app from `graph.appid.json`) and `Invoke-M365GraphPaged` (follows `@odata.nextLink`), which two scripts had each written for themselves |
| Verified: syntax check; with mocked Exchange cmdlets: Direct with `-TenantId` signs in without `-DelegatedOrganization`, a second call reuses the session, a child that adds only IPPS closes only that session and the parent's stays open, GDAP uses the customer domain and a different customer opens a new session; `Invoke-M365GraphPaged` with a mocked two-page and an empty collection. **Not** verified: against a tenant |

### 2026-10-08 (9)
| Change |
|--------|
| **[`Office365Toolkit/`](scripts/Office365Toolkit/readme.md) signs in through [`Connect-M365.ps1`](scripts/Startup/readme.md#connect-m365ps1).** All nine scripts: delegated by default (device code and GDAP customer from `load.config.ps1`), new `-ClientId`, `-CertificateThumbprint` and `-AppOnly` for app-only, and they only disconnect what they opened. The four Exchange-only scripts and `Test-SharedMailboxSignIn` passed `-Organization` with an interactive sign-in, so under GDAP they landed in the partner's own tenant; they now reach the customer with `-DelegatedOrganization` |
| **Scripts that silently returned too little.** [`Get-IntunePolicyInventory.ps1`](scripts/Office365Toolkit/Intune/readme.md#get-intunepolicyinventoryps1) called `Get-MgDeviceManagementConfigurationPolicy` and `Get-MgDeviceManagementIntent`, which do not exist in Microsoft.Graph v2: Settings Catalog and Endpoint Security were always missing. It now reads all five policy types through Graph with paging. [`Get-SecureScoreReport.ps1`](scripts/Office365Toolkit/Security/readme.md#get-securescorereportps1) always printed an empty control table (it read typed properties from `AdditionalProperties`); it now joins each control on its control profile for the maximum points and title |
| **More Graph, fewer bugs.** [`Search-MailboxAuditLog.ps1`](scripts/Office365Toolkit/Exchange/readme.md#search-mailboxauditlogps1) now uses the Graph Audit Log Query API (create, poll, page) with new `-TimeoutMinutes`; `-UseExchange` keeps `Search-UnifiedAuditLog`, now one paged search per record type. [`Remove-EnterpriseAppConsent.ps1`](scripts/Office365Toolkit/Security/readme.md#remove-enterpriseappconsentps1) filters grants server-side instead of reading every grant in the tenant, stops on a failed read instead of reporting "no grants", escapes apostrophes in filters, shows role names instead of GUIDs and asks for write scopes only with `-Apply`. [`Test-MailboxForwardingRisk.ps1`](scripts/Office365Toolkit/Exchange/readme.md#test-mailboxforwardingriskps1) no longer calls an `[EX:/o=...]` recipient external. What stays on Exchange (add-ins, forwarding and sweep rules, CAS settings, EOP policies, the shared-mailbox list) has no Graph API; each readme says so |
| Verified: syntax check; link check; every Mg cmdlet used exists in Microsoft.Graph 2.41.1 and the two removed ones do not; offline runs with mocked Graph calls of the Secure Score, Intune inventory (paging, assignment counts), audit query (request body, two pages, failed status) and app-consent logic (filters, escaping, role lookup); the forwarding recipient classifier unit-tested. **Not** verified: any run against a tenant - the real audit-query status values and record fields, `$expand=assignments` on beta `configurationPolicies`, whether the preview scopes suffice, and the GDAP path in practice |

### 2026-10-08 (9)
| Change |
|--------|
| **New [`Update-SessionHostImage.ps1`](scripts/RDS/Update-SessionHostImage.ps1): a Windows 11 multi-session image (or AVD host) that keeps new Teams, new Outlook and Copilot working with FSLogix, without updating FSLogix.** Three pooled hosts built from a 2024 image (24H2, 26100) kept breaking the same way: an app updates itself per user on one host, FSLogix replays that exact version on another host that does not have it, and registration fails with `0x80070490`. Repair-AppxPackageStore and Update-TeamsClient fix the apps, but nothing checked what the image carries around them, or stopped the drift. The script checks the edition and pending reboot, the FSLogix build (read only), the Store / Teams update hold-back, WebView2 against Edge Stable, the provisioned builds and users holding a newer one, every framework the apps' manifests depend on, Teams on AVD (IsWVDEnvironment, SlimCore minimum build, the meeting add-in, the WebRTC redirector out of support since 1 October 2026 and gone 1 April 2027), Shared Computer Activation, the sign-in broker and `redirections.xml`; `-ForCapture` adds Sysprep readiness. It fixes the policies, SCA and WebView2 itself, the apps through the two existing scripts, reads everything back, and with `-ComputerName` compares the pool. Menu entry `J` |
| Verified: syntax check; the Edge Stable lookup, the WebView2 / Edge / Office channel registry readers and the package-family regex were run locally on Windows 11 (not multi-session, not elevated). **Not** verified: a full run, the fixes, `-ComputerName` or `-ForCapture` on a real session host or image VM |

### 2026-10-08 (8)
| Change |
|--------|
| **New [`Connect-M365.ps1`](scripts/Startup/Connect-M365.ps1): one sign-in for every script, Graph first and delegated by default.** An audit of every M365 script showed that each one signed in its own way: only the startup of `functies.ps1` honoured `useDeviceCodeAuth` from `load.config.ps1`, no Exchange script reached a GDAP customer (they passed `-Organization`, which Exchange only applies to app-only sign-in; a partner needs `-DelegatedOrganization`), about a dozen scripts could only run app-only and some 35 only in the browser. The new file gives `Connect-M365Graph`, `Connect-M365Exchange` (with `-IncludeCompliance`), `Connect-M365Teams` and `Connect-M365PnP` the same rules: delegated by default, device code when `load.config.ps1` says so, the GDAP customer from `$global:cid` / `Connect-Tenant`, app-only with `-ClientId` + `-CertificateThumbprint` or `-AppOnly` from `graph.appid.json`; an existing session is reused when it fits, and `Disconnect-M365Graph` / `Disconnect-M365Exchange` only close what the script opened itself. Exchange, Teams and PnP are only for work Graph has no API for. The scripts move to it in the commits that follow |
| Verified: syntax check; the tenant resolution (no setting, GDAP with `cid`, explicit `-TenantId`, Direct), the Exchange customer domain under GDAP, the device-code choice (`useDeviceCodeAuth`, `-Interactive` overriding it) and the `graph.appid.json` lookup (by tenant, the only entry, an unknown tenant with a clear error) were run locally; parameter names were checked against ExchangeOnlineManagement 3.10.1 and Microsoft.Graph.Authentication 2.41.1. **Not** verified: an actual sign-in against a tenant |

### 2026-10-08 (7)
| Change |
|--------|
| **GitHub Actions: [`.github/workflows/ci.yml`](.github/workflows/ci.yml).** Until now every check ran only locally, through the pre-commit hook and the Claude Code hooks; a commit from a clone without `core.hooksPath`, or from the GitHub web editor, reached `main` unchecked. The workflow runs on every push and PR to `devel`/`main`, and executes nothing - it only parses and lints. Four jobs: **PowerShell 7** ([`Test-PowerShellSyntax.ps1`](scripts/Startup/Test-PowerShellSyntax.ps1) on every `.ps1`/`.psm1`, and PSScriptAnalyzer at error level, without `PSAvoidUsingConvertToSecureStringWithPlainText`, which the break-glass scripts trigger on purpose); **Windows PowerShell 5.1** (parse every script that has no `#Requires -Version 7`); **readmes** ([`Test-MarkdownLinks.ps1`](scripts/Startup/Test-MarkdownLinks.ps1), [`Update-ReadmeHeader.ps1`](scripts/Startup/Update-ReadmeHeader.ps1) - which fails on a missing language - and [`Update-ScriptIndex.ps1`](scripts/Startup/Update-ScriptIndex.ps1), followed by `git diff --exit-code` so stale generated files fail); **shell** (`bash -n` and ShellCheck at warning level on every `.sh`). Findings appear as annotations on the file and line |
| **The 5.1 job found 66 scripts that do not parse in Windows PowerShell 5.1.** Nearly all are UTF-8 without BOM containing a character such as `—` or `é`: 5.1 reads them as ANSI, the bytes become a stray quote, and the script breaks with "string is missing the terminator". PowerShell 7 has no such problem, so it went unnoticed - but Intune, GPO and scheduled tasks run 5.1. A few others use `??` or `?.` without `#Requires -Version 7` (`Import-M365Users.ps1`, `create_scheduled_task.ps1`, `Get-SharePointStorageReport.ps1`). The job is therefore `continue-on-error` for now: it reports, it does not block. Fixing the scripts is a separate change |
| Verified: on a clean checkout of `HEAD`, locally, PowerShell 7 syntax (194 files), the analyzer at error level (only the excluded rule), the link check, header and index generation (no difference) and ShellCheck 0.11 on [`Invoke-LinuxCleanup.sh`](scripts/Linux/Invoke-LinuxCleanup.sh) are all clean; the 5.1 parse was run with `powershell.exe` and gives the list above. If the first run on GitHub needed a fix, that is in the commit after this one |

### 2026-10-08 (6)
| Change |
|--------|
| **Modules are installed and updated at startup without asking.** [`load.ps1`](load.ps1) now runs [`Update-Modules.ps1`](scripts/Startup/Update-Modules.ps1) with the new `-Auto` instead of `-Prompt`: what is missing is installed, what is outdated is updated, and only that is shown; with everything current it is one line, `Modules OK`. `-Prompt` stays for whoever wants to be asked. The same `-RequiredOnly -Auto -MaxAgeHours 24` line is what goes into a PowerShell profile |
| **New [`Test-RequiredModules.ps1`](scripts/Startup/Test-RequiredModules.ps1), run by the docs hook after every edit.** The startup check only installs what [`RequiredModules.psd1`](scripts/Startup/RequiredModules.psd1) lists, and a script that starts using a new module worked on its author's machine and failed everywhere else until someone remembered to add it. The script searches every `.ps1`/`.psm1` for `#Requires -Modules`, `Import-Module` and `Install-Module` with a literal name and reports each one the list lacks. [`sync-docs.ps1`](.claude/hooks/sync-docs.ps1) runs it after every edit of a `.ps1`, `.psd1` or `.md` (waking Claude) and warns with it at pre-commit. Modules deliberately not installed from the gallery go into the new `NotManaged` section with a reason: `ActiveDirectory`, `WebAdministration`, `AzureAD`, `Microsoft.Graph` |
| The first run of that check found five modules scripts used that no list had: `Microsoft.Graph.Reports` and `Az.OperationalInsights` (in `#Requires` lines), `DCToolbox`, `IntuneBackupAndRestore` and `Az.Accounts`. All five are now in the list. Menu keys `U` (Update-Modules, which had no key yet) and `Z` (Test-RequiredModules) |
| Verified: syntax check; `Test-RequiredModules.ps1` on the repository is clean under PowerShell 7.6 and 5.1, and a probe file with a `#Requires` list including a module specification and an `Import-Module -Name '...'` gave exactly those three names, while `Import-Module $dynamic` was skipped; the PostEdit hook reported the missing module for a probe file and stayed silent for the clean repository. `-Auto` ran for real on this machine: it installed the missing `DCToolbox` 2.1.6 without asking (about 30 s, once), and the next start printed only `Modules OK (20 vereist, actueel)` in 2 s including the pwsh start. **Not** verified: updating a module installed for all users from a non-elevated session (expected to report an error) |

### 2026-10-08 (5)
| Change |
|--------|
| **[`Invoke-LinuxCleanup.sh`](scripts/Linux/Invoke-LinuxCleanup.sh) after its first run on a live 3CX server.** The dry run there showed three problems. It said `apt/dpkg is running` and skipped every package step, while `apt-get` worked fine: the check looked for a process named `unattended-upgr`, and `unattended-upgrades` keeps one running all the time. It now reads dpkg's own lock from `/proc/locks`. The journal (3.2 GB there) was left out of the estimate; the dry run now works out what `journalctl --vacuum-time/--vacuum-size` frees, the way journald decides: only archived files, oldest first, by the time in the file name. And a 1.2 GB backup from 2020 in `/var/lib/3cxpbx/Data/Backups`, the folder from before `Instance1`, did not show up as a backup. That folder and its `Logs` are now included, and every backup is listed with date and size, so `--keep-backups` is an informed choice |
| Verified: `bash -n`; in a Debian 12 container with a running `systemd-journald` and rotated journals, the dry-run estimate (74.2 MB) matched what `--apply` freed (74.2 MB), and a fake archived file whose name dates it 60 days back was counted and removed; with `--keep-backups 3` the backups from 2020 (old folder) and the v18 one were removed and the 3 newest kept; a process named `unattended-upgr` no longer blocks the package steps, while a held dpkg lock (`fcntl`) still does. **Not** verified: `--apply` on the live 3CX server |

### 2026-10-08 (4)
| Change |
|--------|
| **[`Install-Printer.ps1`](scripts/Device/Printer/Install-Printer.ps1) hardened for unattended runs after a golden image.** One real bug: the preflight asked the spooler whether a driver was installed *before* waiting for it, so on a first boot every driver would have been called missing. Beyond that, a run now cannot get half-way on bad input: the whole JSON is validated first and every problem listed at once (unknown fields such as `adress`, duplicates, names Windows refuses, addresses, ports, enum values, `"true"` as text, two printers on one port with different addresses); the INF is read before pnputil sees it (printer class, a section for this architecture, the exact driver name among its models — naming the closest when it does not match — and a signed catalog for this architecture); only the native-architecture driver counts as installed. Documented in the [Printer readme](scripts/Device/Printer/readme.md#install-printerps1) |
| Also: a machine-wide lock so a startup task and an RMM job do not both install; the spooler waited for until it answers, and restarted with one retry when it stalls mid-install (deliberately not after every driver, as some published scripts do — that interrupts printing on a session host in use); print defaults read and written in a job with a timeout; proxy/login pages and non-zip files refused as downloads, 1 GB free required; the JSON from a URL cached as fallback; `-Proxy`; public GitHub files fetched through the release download link and `raw.githubusercontent.com`, outside the 60-requests-an-hour API limit a pool behind one NAT would exhaust; LFS pointers and truncated trees refused; a changed address moves or rebuilds the port; one rotated `Install-Printer.log`; and a clean run writes `ConfigSha256`/`LastSuccess` under `HKLM:\SOFTWARE\M365-Scripts\InstallPrinter` for detection. Checked against published scripts by MSEndpointMgr, Konrad Brunner (AlyaKoni) and Printune |
| Verified on Windows PowerShell 5.1 and PowerShell 7: the INF reader on six of Windows' own printer INFs, exact and near-miss names on `prnms009.inf`, synthetic x86-only, `%token%`, missing-catalog, installer-only and non-printer packages; HTML and junk saved as `.zip`; the timeout; the lock with three parallel processes (one waits and gets it, one gives up after its budget); the cache fallback on a 404; a raw download while the API was rate limited (that limit was hit for real during testing); a broken JSON giving 19 problems; full `-CheckOnly`/`-WhatIf` runs, a typo, an empty file and a bad proxy. Those ran with the elevation check bypassed, as this session was not elevated. **Not** verified: an elevated run that installs a driver and printer, the port rebuild, and a first boot from a golden image |

### 2026-10-08 (3)
| Change |
|--------|
| **[`load.ps1`](load.ps1) checks the modules at every start: missing, too old, or with an update.** Before, it only looked whether seven hardcoded modules existed — an outdated module, or a module a newer script needed, went unnoticed until a command failed. Now it runs [`Update-Modules.ps1`](scripts/Startup/Update-Modules.ps1) `-RequiredOnly -Prompt`, lists what is missing, below its minimum version or behind the PowerShell Gallery, and asks `Deze n module(s) nu installeren/updaten? [J/n]`. The gallery versions are cached for 24 hours (`%LOCALAPPDATA%\M365-Scripts\module-gallery-cache.json`), so a normal start costs under a second instead of six. `-SkipModuleCheck` skips it once. The same call, `-RequiredOnly -Prompt -MaxAgeHours 24`, can go in a PowerShell profile for anyone who starts through `$PROFILE` instead of `load.ps1` |
| **One module list: new [`RequiredModules.psd1`](scripts/Startup/RequiredModules.psd1).** `load.ps1`, [`Install-Modules.ps1`](scripts/Startup/Install-Modules.ps1) and `Update-Modules.ps1` each had their own list (7, 14 and 7 modules) and they did not agree. All three now read this file, so adding a module there is enough for every machine to be offered it at its next start. Added: `PnP.PowerShell` (used by 10 SharePoint/Teams scripts, skipped below PowerShell 7.4) and `MicrosoftTeams`, which scripts used but no list installed. Removed: `AzureAD` — Microsoft took it off the PowerShell Gallery, so `Install-Modules.ps1` failed on it on every run |
| `Update-Modules.ps1` rewritten: per-module status (`Missing`, `BelowMinimum`, `UpdateAvailable`, `OK`, `Unknown`, `Skipped`), new `-CheckOnly`, `-RequiredOnly`, `-MaxAgeHours`, `-Scope`, `-Quiet`, `-PassThru`, `-Prompt`. It reads the installed version with `Get-Module -ListAvailable` (also sees modules not installed through PowerShellGet), asks the gallery with `Find-PSResource` when available (3 s instead of 9 s for `Find-Module`), and installs side by side where `Update-Module` would refuse. Without parameters it still updates every other installed module afterwards, as before |
| Verified: syntax check; on PowerShell 7.6 `-RequiredOnly -CheckOnly` against the live gallery reports all 15 modules `OK` on this machine (6 s), and from the cache in 1.7 s including the pwsh start (0.8 s inside a running session); `-Prompt` prints a single `Modules OK` line, and on a non-interactive host falls back to changing nothing; with a test list and a doctored cache it reports `Missing`, `BelowMinimum`, `UpdateAvailable`, `OK` and `Skipped` correctly, in list order, and `-PassThru` returns the objects `load.ps1` reads; the same test under Windows PowerShell 5.1 works and sees that runtime's own module folders. **Not** verified: an actual install or update run (nothing was installed on this machine), and `load.ps1` interactively from start to menu |

### 2026-10-08 (2)
| Change |
|--------|
| **New [`Invoke-LinuxCleanup.sh`](scripts/Linux/Invoke-LinuxCleanup.sh) in the new folder [`Linux/`](scripts/Linux/readme.md)** — the Linux counterpart of `Invoke-WindowsCleanup.ps1`, for a Debian/Ubuntu server and in particular one that runs 3CX Phone System. A bash script, because such a server has no PowerShell. Cleans the APT cache, `autoremove` (old kernels; refused when the list holds a 3CX package, skipped while apt/dpkg is running), leftover package configuration, disabled snap revisions, the systemd journal, rotated logs, crash dumps, `/tmp`, user caches and trash, optionally Docker (`--docker`, volumes never), and, when 3CX is detected, 3CX's logs and backups beyond the newest N (`--keep-backups`). Call recordings, the database and configuration are only reported. Dry run by default, `--apply` to clean up, `--check-only` for monitoring with exit code `2` when there is work |
| New [`.gitattributes`](.gitattributes): `*.sh` is checked out with LF line endings, also on Windows — with CRLF, bash on the server fails on the first line |
| Verified: `bash -n` and ShellCheck (no warnings); in a Debian 12 container with a recreated 3CX folder layout, aged files, 7 backups, a recording, an autoremove candidate and a package with leftover config: dry run and `--check-only` change nothing (exit `0` / `2`), `--apply` removes exactly the old files (active logs, recent files, the recording, PostgreSQL's lock file and `systemd-private-*` stay), keeps the 3 newest backups, a second `--apply` finds nothing, without 3CX the 3CX sections are left out, and totals above 2 GB add up (Debian's `mawk` overflowed `printf "%d"` at 2 GiB, now `%.0f`). **Not** verified: a live 3CX server (its log and backup paths come from 3CX's Linux layout), the journal (the container has no systemd), snap and `--docker` |

### 2026-10-08
| Change |
|--------|
| **New [`Install-Printer.ps1`](scripts/Device/Printer/Install-Printer.ps1) in the new folder [`Device/Printer/`](scripts/Device/Printer/readme.md).** Installs printer drivers and TCP/IP printers as described in one JSON file ([`printers.example.json`](scripts/Device/Printer/printers.example.json)). Drivers come from a GitHub release asset, a folder in a GitHub repository (through the API, so a private repository works with `-GitHubToken` / `GITHUB_TOKEN`; the token is only ever sent to GitHub's own hosts), any https URL or a share. Only a missing driver, or one older than the JSON's `version`, is downloaded; an optional `sha256` and the catalog signature are checked before `pnputil /add-driver /install` and `Add-PrinterDriver`. Then port, printer, location/comment/sharing and print defaults; `"ensure": "absent"` removes a printer. The existing [`Add-NetworkPrinterConnection.ps1`](scripts/TenantOnboarding/AppDeployment/Add-NetworkPrinterConnection.ps1) only adds a printer for a driver that is already there |
| Built to run unattended on a server or session host fresh from a golden image: a Print Spooler that has not started yet is started and waited for, downloads that fail on DNS or a timeout are retried within `-WaitSeconds` (a 401/403/404 fails at once), it relaunches itself 64-bit because `pnputil` does not exist under SysWOW64, and every run is idempotent so it can also run at every startup. Practices taken from published Intune/RMM printer guides: pnputil exit codes `259`/`3010`/`1641` as success with a pointer to `setupapi.dev.log`, SNMP off on new ports unless `"snmp": true`, `Set-PrintConfiguration` in a job with a 2-minute timeout because some universal drivers hang in it. Menu key `N` |
| Verified: syntax check; on Windows PowerShell 5.1 and PowerShell 7, the functions on their own — a real release asset from `cli/cli` (and the refusal when a pattern matches 5 assets), a 7-file folder from `actions/checkout` and a single file through the API, a 404 with the token hint, `sha256` match and mismatch, INF lookup in a UTF-16 package with x86/x64 folders, the catalog signature of Windows' own `prnms009.inf` (Microsoft Print To PDF) and its version decoded as `10.0.26100.8951`, retry on a transient error, no retry on a 404, giving up within the budget; the whole script with `-CheckOnly` (exit `2`), `-WhatIf`, `-Printer` with one match, `-Quiet` on a device in order (no output, exit `0`), RMM variables, an incomplete JSON, and the 32-bit relaunch. Those runs had the elevation check bypassed, because this session was not elevated. **Not** verified: an elevated run that actually installs a driver and a printer, and a first boot from a golden image |

### 2026-10-07
| Change |
|--------|
| **[`Update-TeamsClient.ps1`](scripts/Device/Update-TeamsClient.ps1) installs the build it checked against.** The version check asks the Teams config service which build is current, but the install let `teamsbootstrapper.exe -p` download whatever its own staged rollout handed out — and the two disagreed for weeks. A run ended with `Installed build 26246.1604.5133.838 is still older than the published 26260.1701.5139.3736`, three weeks after that build went up, and the next scheduled run would call the host outdated and reinstall it again. The config service also returns a `buildLink` to the exact MSIX, which the script read and never used. Step 5 now downloads that package and checks its size and Microsoft signature before anything is uninstalled, and step 7 provisions it with `-p -o`. If the download or the check fails, the run warns and falls back to the bootstrapper's own choice, since nothing has been removed yet |
| New `-UseBootstrapperBuild` (NinjaOne: `useBootstrapperBuild`) restores the old behaviour, leaving Microsoft's staged rollout in charge. It cannot be combined with `-UseWinget`. Documented in the [Device readme](scripts/Device/readme.md#update-teamsclientps1), [Update-TeamsClient.md](scripts/Device/Update-TeamsClient.md) and the IT Glue version (md + regenerated html) |
| Verified: syntax check; the live config service publishes `26260.1701.5139.3736` with a `buildLink` on `teamsinstaller.public.onecdn.static.microsoft`; the script's own `Get-LatestTeamsBuild` and `Save-VerifiedDownload` with the step 5 logic, under Windows PowerShell 5.1, downloaded the 274 MB package and accepted its signature (`Valid`, `O=Microsoft Corporation`, also under PowerShell 7); a link answering 404 gave the fallback warning and left no file behind; `-UseWinget -UseBootstrapperBuild` is refused with exit `1` on both runtimes. **Not** verified: a full elevated run that provisions from the downloaded package — no host was updated with this version yet |
### 2026-10-06
| Change |
|--------|
| **[`Get-SharePointStorageReport.ps1`](scripts/Reporting/Get-SharePointStorageReport.ps1) no longer contradicts itself about version history.** Skipping version history is what `-SkipVersions` has always done, and `-FastMode` implies it — but that implication was applied *after* the mode banner was printed, so a `-FastMode` run first announced `Full scan including version history` and then, two lines further down, `Fast scan (no version history, no detail rows)`. On a run of several hours that is the difference between trusting the output and not. `-FastMode` now sets `-SkipVersions` before the banner and has its own line in it |
| Verified: syntax check; the banner logic run through for all four combinations — `-Apply` gives `Full scan including version history`, `-Apply -SkipVersions` gives `Full scan (version history skipped)`, `-Apply -FastMode` and `-Apply -FastMode -SkipVersions` both give exactly one line, `Fast scan (no version history, no detail rows)`. Nothing changed about what is scanned |

### 2026-10-05 (15)
| Change |
|--------|
| **[`Get-SharePointStorageReport.ps1`](scripts/Reporting/Get-SharePointStorageReport.ps1) also checks PDFs against Adobe's limit.** Acrobat and Reader cannot open a PDF whose path is longer than 255 characters — notably from a synced or network folder, where Adobe's 2021 fix does not always help — so a PDF between 256 and 259 local characters passed the Windows check yet would not open. Such a file is now marked `Adobe (255)` in the long paths CSV, the console and the Markdown report. Documented in the [Reporting readme](scripts/Reporting/readme.md#long-paths-windows-limits) |
| Verified: syntax check; the measuring function run on PDFs of 255, 256, 259 and 260 local characters — nothing, `Adobe (255)`, `Adobe (255)` and `Windows (260)`. Not run against a live tenant |

### 2026-10-05 (14)
| Change |
|--------|
| **[`Get-SharePointStorageReport.ps1`](scripts/Reporting/Get-SharePointStorageReport.ps1) now reports the longest paths and checks them against Windows' limits.** A library that is fine in SharePoint can still break once it is synced with OneDrive: the local path `C:\Users\<user>\<Organisation>\<Site> - <Library>\...` is longer than the SharePoint path, and nothing in the report showed that. With `-Apply` (also `-FastMode`) every file and folder is now measured against SharePoint's 400 characters, Windows `MAX_PATH` (260) and, for workbooks, Excel's 218. Result in `SharePoint_LongPaths_<timestamp>.csv`, longest first, plus a top 10 in the console and the Markdown report; it survives a resume from a checkpoint |
| The local path is measured for the enabled member account with the **longest UPN** in the tenant, because the profile folder is named after the UPN prefix — a path that fits for that user fits for everyone. This adds `User.Read.All` to the interactive sign-in. New parameters `-SyncProfilePath`, `-OrganizationName` and `-LongPathThreshold` (default 200) override the estimate. Documented in the [Reporting readme](scripts/Reporting/readme.md#long-paths-windows-limits) |
| Verified: syntax check; the measuring function run against sample paths — an Excel file at 205 local characters stays under 218, a deep path at 282 is flagged `Windows (260)`, a path of 441 SharePoint characters `SharePoint (400)`, and a OneDrive personal site is measured as `OneDrive - <Organisation>`. Not yet run against a live tenant, so the UPN and organisation lookups through Graph are untested |

### 2026-10-05 (13)
| Change |
|--------|
| **`Remove-CorporateWallpaper.ps1` also removed the corporate lockscreen.** It deleted the whole `PersonalizationCSP` key and the whole `C:\ProgramData\Wallpapers` folder, but `Make-lockscreen.ps1` keeps its `LockScreen*` values in that key and its image in that folder. It now removes only the `Desktop*` values and the `corporate-background-*` files, and the key or folder only when nothing else is left in it |
| It also missed half of what `Set-CorporateWallpaper.ps1` writes: the `Policies\System` fallback and every other loaded user hive stayed on the corporate wallpaper, and run as SYSTEM its `HKCU` reset hit SYSTEM's own profile instead of a user's. It now resets the policy values, every loaded user hive and the Default User profile — only where they point at a corporate file — back to the Windows default image, clears those users' transcoded wallpaper cache, and touches `HKCU` only when not running as SYSTEM. Logs to the Intune log folder; `-WhatIf` shows what it would do |
| Verified: syntax check; the functions tested against a scratch registry key — a corporate value is reset to `img0.jpg` with style Fill, a lockscreen file or a wallpaper elsewhere is not recognised as ours, a foreign wallpaper is kept, and `-WhatIf` changes nothing; and a full `-WhatIf` run on a workstation, which skips the Default User hive with a message when not elevated. Not run as SYSTEM or through Intune |

### 2026-10-05 (12)
| Change |
|--------|
| **The USB toolkit's `LocalAdmin` password was hardcoded in `scripts/Deployment/start.bat`**, printed on screen after every run, and written out in the English readmes — together with the internal IP address of the install share. `start.bat` now reads both from `start.local.cmd` next to it (git-ignored; template `start.local.example.cmd`). When that file or a value is missing it asks: the password with hidden input, the share path when option `E` is chosen. With no password, options `D`/`E` stop instead of creating an account, and the password is no longer echoed |
| **The password is still in the git history and on every device the toolkit has prepared — change it.** Copy `start.local.example.cmd` to `start.local.cmd` on the USB stick with the new password and the share |
| Verified with a copy of `start.bat` in which `net`, `wmic` and `reg` only print what they would do: with `start.local.cmd` the account is created with the password from the file (including `$`, `&` and `!`) and the menu shows the share; without it and with an empty answer, nothing is created and option `E` asks for the path. The hidden prompt itself needs a console and was checked separately (SecureString back to text in Windows PowerShell). Not run in OOBE |

### 2026-10-05 (11)
| Change |
|--------|
| The root readme is clickable throughout. The **Repository Structure** tree was a code block, so none of its 215 entries could be clicked; it is now a `<pre>` block in which every folder links to its readme (in the reader's language) and every file to the file |
| The **Menu** tables link each tool to what the key actually runs, taken from `menu.ps1`: 48 rows, including `Test-GroupPermissions` → `Test-DistributionGroupPermissions.ps1` and `Get-DLMembers` → `Get-DistributionGroupMembers.ps1`, whose names differ from the label; the M365 functions link to the `functies.ps1` section. Every category in **Script Categories** gained a folder line, and file names in the running text link to the file — 384 per language, only where the name points to exactly one tracked file |
| Verified: every link in the tree points to a tracked file or folder, the same number of links was added in each language, and the link check passes. Two stale tree notes corrected along the way (Custom Scripts "path-pinned", the shared Provisioning module "used by all four") |

### 2026-10-05 (10)
| Change |
|--------|
| Every folder readme now follows the `scripts/Intune/` pattern: each script in the `## Scripts` table links to the file **and** to its section (`([docs](#…))`). 19 folders did not, in all three languages: Azure/VM, DNS, Deployment, Device/DriveMapping, Graph and Intune/Desktop/Background/Desktop and /Lockscreen had no scripts table at all, and Reporting/Licensing only a tree; Device/audio, ClaudeDesktop, CoworkPrerequisites, DiskCleanup, iOS-Compliance-Updater, Get-Autopilot (no file link for `GetAutoPilot.CMD`), UniFi, SAS, SMTP, Startup and one row in SharePoint/Provisioning had rows without a docs link |
| Where a script was documented under a descriptive heading, the heading now carries the file name so the anchor is the same in every language (checked first that nothing linked to the old anchor). Scripts that had no section got a short one from their own header comment: `Remove-CorporateWallpaper.ps1`, `Invoke-`/`Detect-DiskCleanupIntune.ps1`, `UnifiApi.ps1`, `Test-SASWorkDirectory.ps1`, `SharePointStructure.Common.ps1`, `Browse-InstallScripts.ps1`, and a shared section for the ClaudeDesktop install/uninstall/detect scripts. The `Remove-CorporateWallpaper.ps1` notes record that it does not undo everything `Set-CorporateWallpaper.ps1` 2.3+ writes |
| Also corrected: the Intune readme still said the Office theme scripts hardcode their URL, and the Provisioning notes said the shared module is dot-sourced by "all four" scripts — all ten in the folder use it |
| Verified: every script in every folder readme has a file link and a docs link whose anchor exists (scripted check over all 57 folder readmes), and the link check passes over 4562 internal links. Nothing was run |

### 2026-10-05 (9)
| Change |
|--------|
| `Test-MarkdownLinks.ps1` dropped every underscore when it turned a heading into an anchor, so `## testsmtp_5min.ps1` became `#testsmtp5minps1` while GitHub makes it `#testsmtp_5minps1`. A correct link was reported broken (and a broken one would have passed). Underscores are now removed only at a word edge, where they mean emphasis; inside a word they are kept, as GitHub does |
| Verified: syntax check, and the link check over all 204 markdown files with the new `[docs]` link to `#testsmtp_5minps1` resolving. `testsmtp_5min.ps1` is the only heading in the repo with an underscore inside a word, so no other result changes |

### 2026-10-05 (8)
| Change |
|--------|
| [`Install-SharePointStructure.ps1`](scripts/SharePoint/Provisioning/Install-SharePointStructure.ps1) ended a successful build by telling you to fill one client's security groups by name prefix, whatever config had just been built. It now names the number of groups and the config they come from, read through `Get-ConfigValue` so a config without groups does not trip strict mode. Sample output in [`Update-TeamsClient.md`](scripts/Device/Update-TeamsClient.md) showed a real domain account; it shows `CONTOSO\admin` now |
| Verified: syntax check, and the message rendered against the example config (13 groups) and a config without groups (0) under strict mode. A repo-wide search outside the version history finds no remaining customer names |

### 2026-10-05 (7)
| Change |
|--------|
| **`Invoke-TeamsArchive.ps1 -DryRun` did not dry-run on a first run.** The script removes conflicting Graph modules and restarts itself in a clean `pwsh` session, but the restart passed only the script path — every parameter was dropped. The restarted session ran with defaults: no `-DryRun`, no `-Step10Only`, no `-ChannelAction`, so a run meant as a simulation went through the real export and the interactive Step 10. Only a run in a session where the restart flag was already set kept its parameters |
| The restart now passes every bound parameter on (switches only when set, values as they were given) and exits with the restarted run's exit code instead of always 0 |
| Verified with a stand-in script built from the real parameter block and forwarding code: `-DryRun -Step10Only -Step10Action undo` and values containing spaces arrive intact in the restarted session, defaults stay defaults, and the child's exit code comes back. The archiver itself was not run |

### 2026-10-05 (6)
| Change |
|--------|
| Renamed `scripts/Teams/vias_archiver.ps1` to [`Invoke-TeamsArchive.ps1`](scripts/Teams/Invoke-TeamsArchive.ps1) and took one customer out of it: the wizard titles, the prompts for tenant and admin account, the example SharePoint URL, the temporary app name, the temp files, the eDiscovery case name and the report file name all named that customer, and the default Excel file and archive folder pointed at its own file and network drive. The defaults are now `C:\Temp\Teams_Channels.xlsx` and `C:\Temp\Teams_Archive` |
| The Excel worksheet was read by a hardcoded, customer-named sheet. New `-WorksheetName`; without it the first worksheet is read. The restart marker environment variable is renamed with the script |
| The Teams readme gained the `## Scripts` table it lacked, the new parameter, and the columns the Excel file needs. Not in [`menu.ps1`](menu.ps1), so nothing to rename there |
| Verified: syntax check and a search for remaining customer names. Not run against a tenant |

### 2026-10-05 (5)
| Change |
|--------|
| [`SharePoint/Provisioning/`](scripts/SharePoint/Provisioning/readme.md) said it was client-neutral but shipped one client's complete configuration (tenant, owner, site URLs, groups), a user guide written for another, and five scripts whose default `-ConfigPath` was a client-named config file that did not exist — so running them without `-ConfigPath` failed on a missing file |
| The client config is replaced by [`example.config.json`](scripts/SharePoint/Provisioning/example.config.json): the same model with Contoso names and `CHANGEME` in the tenant, owner and site URLs. Client configs (`<client>.config.json`) stay next to it but are git-ignored, so an existing one keeps working locally and is no longer published |
| New `Resolve-StructureConfigPath` in [`SharePointStructure.Common.ps1`](scripts/SharePoint/Provisioning/SharePointStructure.Common.ps1): with no `-ConfigPath` the five scripts now take the one `*.config.json` there that no longer contains `CHANGEME` — the rule [`Install-SharePointStructure.ps1`](scripts/SharePoint/Provisioning/Install-SharePointStructure.ps1), [`Remove-SharePointStructure.ps1`](scripts/SharePoint/Provisioning/Remove-SharePointStructure.ps1) and [`Sync-SharePointChannelMember.ps1`](scripts/SharePoint/Provisioning/Sync-SharePointChannelMember.ps1) already used. [`menu.ps1`](menu.ps1) says so in its prompt instead of naming a file |
| The user guide renamed to [`SharePoint-Handleiding.md`](scripts/SharePoint/Provisioning/SharePoint-Handleiding.md) and turned into a template (Contoso NV, brands Northwind and Fabrikam, with a note to replace them). [`New-StructureConfig.ps1`](scripts/SharePoint/Provisioning/New-StructureConfig.ps1) no longer suggests one client's name and brands as defaults; examples and readmes use Contoso. Internal column names (`PsMerk`, …) are unchanged — they live in each config and in sites already built |
| Verified: syntax check on the folder and [`menu.ps1`](menu.ps1); the example imports cleanly once `CHANGEME` is filled in and is refused as shipped; the resolver picks the one filled-in config and skips the example. No run against a tenant |

### 2026-10-05 (4)
| Change |
|--------|
| [`Deploy-OfficeTheme.ps1`](scripts/Custom%20Scripts/Intune/Desktop/Deploy-OfficeTheme.ps1) and [`Deploy-Officecolors.ps1`](scripts/Custom%20Scripts/Intune/Desktop/Office%20Themes/Deploy-Officecolors.ps1) were built for one customer: the theme's download URL and file name were hardcoded, pointing at that customer's `.thmx` and colour XML inside a GitHub repo. They now take `-ThemeUrl`/`-ThemeName` and `-ColorsUrl`/`-ColorsName`; the name defaults to the last segment of the URL and must end in `.thmx` or `.xml`. Without a URL they stop with exit code 1 instead of prompting, since nobody answers a prompt under Intune |
| Removed the customer's theme files (`.thmx` and colour-scheme `.xml`) from the repo. The readmes no longer say the scripts are pinned to this path, and explain how to pass parameters under Intune (Win32 app command line, or a copy with defaults filled in) |
| **Existing Intune deployments carry their own copy of the old script and keep downloading from the URL inside it** — that URL points at a different GitHub repository, and if that repository is synced from this one, those deployments lose the file once this reaches `main` |
| Verified: syntax check, the missing-URL and wrong-extension paths exit 1 with their message, and a percent-encoded URL yields the expected file name. No download or Intune deployment was run |

### 2026-10-05 (3)
| Change |
|--------|
| [`Init-TempDisk.ps1`](scripts/Device/TempDisk/Init-TempDisk.ps1) remembered its last self-triggered restart under a registry key named after one company. The key is now the `-RestartMarkerPath` parameter, default `HKLM:\SOFTWARE\M365-Scripts\InitTempDisk`, validated to be an `HKLM:` or `HKCU:` path |
| **On machines that already ran it, the old key is no longer read**, so the restart cooldown starts afresh once: at most one extra restart per machine, and only when every other restart condition holds. Pass the old key as `-RestartMarkerPath` to keep it |
| Verified: syntax check, and the parameter's pattern accepts the default and rejects a file path. Not run on a device |

### 2026-10-05 (2)
| Change |
|--------|
| Replaced customer data in examples with Contoso placeholders: [`Set-UserManager.ps1`](scripts/Entra/Set-UserManager.ps1) (a customer mail domain), [`Import-DnsRecords.ps1`](scripts/DNS/Import-DnsRecords.ps1) (a customer DNS zone and DC), [`Get-ComputerLastLogon.ps1`](scripts/Reporting/Get-ComputerLastLogon.ps1) (a customer OU path) and [`Get-SharePointStorageReport.ps1`](scripts/Reporting/Get-SharePointStorageReport.ps1), whose help text named a real person's mailbox |
| The licensing report had a company OneDrive path (`C:\OneDrive\<Company>\...`) hardcoded in both [`genereer_rapport.ps1`](scripts/Reporting/Licensing/genereer_rapport.ps1) and [`genereer_licentie_overzicht.py`](scripts/Reporting/Licensing/genereer_licentie_overzicht.py), and the readme told you to edit the scripts. The folder now comes from `-ExportDir` / `--export-dir`, else the `LICENSING_EXPORT_DIR` environment variable; with neither, both stop with exit code 2 instead of guessing. The launcher checks this before `Join-Path` is reached, which would otherwise throw on an empty path |
| [`create_scheduled_task.ps1`](scripts/Reporting/Licensing/create_scheduled_task.ps1) took its settings from variables to edit, including a fixed service account. `-ExportDir` and `-RunAsUser` are now required parameters, `-RunDay`/`-RunTime` optional, and the task passes `--export-dir` to Python. `-RunTime` was also ignored when computing the first run (always 08:00); it is used now. **A task registered earlier runs without `--export-dir` and now stops at once — re-register it** |
| Verified: syntax check on the touched PowerShell, `py_compile` on the Python engine, and the launcher without an export directory exits 2 with the message. The Python engine itself was not run (no `pandas` on this machine) and the scheduled task was not re-registered |

### 2026-10-05
| Change |
|--------|
| Removed `scripts/djm` and `scripts/djm.pub` — an OpenSSH private key (`djm-portaal`) and its public half that were committed to the repo. Nothing in the repo used them. [`.gitignore`](.gitignore) now excludes `id_*`, `*.pem`, `*.key` and `*.pub`. **The key is still in the git history and must be considered compromised: rotate it on every host that trusts it** |

### 2026-10-02 (5)
| Change |
|--------|
| A live run of [`Revoke-SharePointUserAccess.ps1`](scripts/SharePoint/Revoke-SharePointUserAccess.ps1) reported `Sites with access: 15` and `Grants found: 0` — with no errors. The permissions report named 30 grants for the same user on those same sites, so the two disagreed and the revoke side was wrong |
| **The role assignments were read from the wrong URL.** `Get-ScopeRoleAssignments` was handed the scope base (`/_api/web`) and queried that directly instead of `/_api/web/roleassignments`. SharePoint answered with the web object, which carries no `value` array, so the paging helper found nothing to iterate and returned an empty collection. Nothing failed; it simply found nothing. The function now appends `/roleassignments` itself, which also matches what the removal URL is built from |
| **`\24384` is a read-only automatic variable.** `\24384 = [int]\.PrincipalId` throws `Cannot overwrite variable PID`. It sat behind the URL bug so it never surfaced, and would have turned every scope evaluation into a caught exception. Renamed, and a check now walks both scripts for assignments to any read-only automatic |
| Added the general guard for the silent half: a collection endpoint always answers with `value` (or `d.results`) even when empty, so a response carrying neither is the wrong URL rather than an empty result. The paging helper now throws instead of returning nothing, in both scripts |
| Verified with 126 checks (6 new) plus a reproduction built from the shapes that tenant actually returned: a direct Dutch `Beperkte toegang` grant is now matched, as are sharing links, joined Entra groups and `Everyone`, while another person and an unjoined group are not. **The corrected read is not yet verified against a live tenant** |

### 2026-10-02 (4)
| Change |
|--------|
| Paired the two SharePoint access scripts through the reports own output. [`Revoke-SharePointUserAccess.ps1`](scripts/SharePoint/Revoke-SharePointUserAccess.ps1) gained `-FromReport`: it takes the sites to visit from a [`Get-SharePointPermissionsReport.ps1`](scripts/Reporting/Get-SharePointPermissionsReport.ps1) run instead of walking the tenant a second time. The report answers who can reach what, you read it and decide, and the revoke acts on exactly what you were looking at — on a tenant where a full sweep takes a quarter of an hour, a user with access to a handful of sites is now revoked in seconds |
| It reads the reports site-access file in preference to the raw grant list. That file already resolved every group to its people, so a site is only visited when the user is genuinely in the group granting access. The first version used the grant list, which names the group but not its members — on a tenant where most sites grant through `Site Members` that meant visiting nearly every site, losing the entire point. A test now proves the preferred path visits fewer sites than the fallback |
| The report decides where to look, never what to remove: every site it names is still read live, so a grant that disappeared in between comes back as `AlreadyGone` rather than a failure, and one removed by hand is not resurrected. The reverse is called out rather than assumed — anything granted after the report, and anything the report itself could not read, is named in the summary, and a report older than a day says so |
| `-FromReport` accepts the detail CSV, any sibling from the same run, or the folder; without `-TenantUrl` the tenant is taken from the report as well, since repeating a URL the file already contains is a way to get it wrong |
| Verified with 120 checks (22 new), 20 of them driving the real functions against real report files on disk: resolving the detail CSV from a folder, a sibling or the workbook; the preferred and fallback paths; a guest matched on either their mail or their tenant UPN; another persons rows ignored; an error row adding no site; and a substring of a real UPN matching nothing. **Not yet verified against a live tenant** |

### 2026-10-02 (3)
| Change |
|--------|
| `-RemoveFromEntraGroups` could only ever act on the groups the run''s own scan found, and nothing said so. A single-site run, a narrowed `-Scope`, excluded OneDrive or hidden lists, or scopes that failed to read all shrink that list — and "removed every group that grants access" then reads as complete when it is not, which is how an offboarding gets signed off half-finished |
| The run now works out what it did **not** cover and says so twice: before removing anything, and again in the summary, naming each limit. A group granting access somewhere that was never searched is explicitly called out as absent from the list |
| Worth stating plainly, because it was a fair question: the revoke script runs its own scan. [`Get-SharePointPermissionsReport.ps1`](scripts/Reporting/Get-SharePointPermissionsReport.ps1) is not a prerequisite — discovery, revocation and the Entra phase happen in one run, in that order |
| `$scanLimits` is declared alongside `$stats` rather than inside the scan, so a run that dies early leaves the summary an empty list instead of an undefined variable |
| Verified with 83 checks (7 new): each limit is collected, the warning appears before the removals and again at the end, and the list survives an early exit |

### 2026-10-02 (2)
| Change |
|--------|
| Added `-RemoveFromEntraGroups` to [`scripts/SharePoint/Revoke-SharePointUserAccess.ps1`](scripts/SharePoint/Revoke-SharePointUserAccess.ps1), completing the second half of an offboarding instead of only reporting it. Until now the script removed every SharePoint-level grant and then told you to go and handle the Entra groups yourself |
| **Only the groups this run actually caught holding a role assignment on a scope in range are touched** — never every group the user belongs to. A leaver can be in fifty groups, and widening this to all of them would be the difference between revoking an access and detaching someone from the organisation |
| An Entra group is not a SharePoint object: the same membership commonly carries a Teams team, a mailbox, licences and app assignments that this report cannot see. The switch is off by default, the banner and summary say what it reaches, and the menu entry defaults to no |
| Four cases are reported rather than forced, because forcing them would fail or do the wrong thing: a dynamic group (membership follows a rule, so there is nothing stored to remove), a group synced from on-premises AD (read-only in the cloud), a membership inherited through a nested group (the user is not a direct member, so the cut has to be made at the group that actually holds them), and a user who could not be resolved in Entra |
| The write permission follows the switch: `GroupMember.ReadWrite.All` is only requested when `-RemoveFromEntraGroups` is given, so a report-only run holds nothing that can change group membership tenant-wide. Removals go through the same funnel as every other change, so `-Apply`, `-WhatIf`, the confirmation prompt and the audit CSV behave identically |
| Verified with 76 checks (11 new): the switch exists and gates the write role, the Entra phase iterates the groups seen granting access and never `$userGroupIds`, all four refusals are present, the removal goes through the funnel, and the per-scope pass still only records an Entra grant rather than acting on it |

### 2026-10-02
| Change |
|--------|
| Added [`scripts/RDS/Invoke-FSLogixShrink.ps1`](scripts/RDS/Invoke-FSLogixShrink.ps1): dynamic FSLogix profile/ODFC VHDX files grow but never give space back, and the shrink was being done by hand from pasted commands that downloaded whatever was on Invoke-FslShrinkDisk's master branch, wrote the log to a `C:\Temp` that might not exist (its `Export-Csv` then fails), and named one customer's share. The script downloads Invoke-FslShrinkDisk at a pinned commit (`bfe0504`, 2025-06-19) and refuses it unless the SHA-256 matches, lists every container on the share largest first (`-ReportOnly`), shrinks with the same defaults (≥ 5 GB, ≥ 10% free, 4 at a time), creates the log folder, and summarises GB recovered and the disks it could not process — usually attached because the user is signed in |
| Looked on GitHub for something better: Invoke-FslShrinkDisk is still maintained by the FSLogix team and remains the tool; the forks and ShrinkVHD do the same with less behind them. The real improvement is FSLogix's own VHD Disk Compaction at every sign-out (2210 and later, on by default), so `-CheckHost` tells whether that can run on a host: version, `VHDCompactDisk`, `defragsvc` not Disabled, dynamic disks. When it passes, a manual shrink only catches up |
| Added menu entry `K` (FSLogix-Shrink), and [`Get-FSlogix-errors.ps1`](scripts/RDS/Get-FSlogix-errors.ps1) to the root readme's RDS category and structure tree, where it was missing |
| Verified in Windows PowerShell 5.1 on a workstation: the download of the pinned commit and the hash check, the second run reusing it, a tampered copy refused, `-ReportOnly` on a test folder with two VHDX files (6 GB and 1 GB, plus a non-VHD file skipped), `-CheckHost` reporting FSLogix not installed (exit 1), and the CSV summary on a sample log (4.75 GB recovered, an in-use disk named, exit 1). The actual shrink has not been run against a share — that needs an elevated session on a host with access to the profile share |

### 2026-10-01 (4)
| Change |
|--------|
| For a package that keeps failing with the right build provisioned, [`Repair-AppxPackageStore.ps1`](scripts/Device/Repair-AppxPackageStore.ps1) now answers the question that decides whether it matters: does every signed-in user have the app? lem-avd-4 had Outlook 915 provisioned and FSLogix still logging `Deployment Register ... from:  (AppxManifest.xml) failed with error 0x80070490` that afternoon — FSLogix registering with an empty path. Whether users were without Outlook or only the log was noisy could not be read from the error |
| The run compares the loaded user hives with the users the package is registered (Installed) for, and names the newest build each has. All covered: the failures are FSLogix's own replay, the run says so, points at `InstallAppxPackages = 0` as Microsoft's documented way to silence it without changing it, and ends with 0. Anyone missing: named, and the run fails |
| Verified end-to-end in Windows PowerShell 5.1 with mocked AppX state: lem-avd-4 with Outlook registered for the signed-in user → "all 1 signed-in user(s) have it (1.2026.915.300)", exit 0; lem-avd-5 with it registered for someone else → the user named, exit 2; an empty host → exit 0. Not yet run on the hosts |

### 2026-10-01 (3)
| Change |
|--------|
| [`Repair-AppxPackageStore.ps1`](scripts/Device/Repair-AppxPackageStore.ps1) aborted on the first live pool run with `The property 'Name' cannot be found on this object` (lem-avd-4). `@($exactTargets.Name)` throws under `Set-StrictMode` in Windows PowerShell 5.1 when the list is empty — which it is on a host that already has the newest build. Built from the items instead |
| Found by the end-to-end run that should have existed before: with `-Name teams,outlook -Provision`, Microsoft's installers also ran for packages already provisioned, and they deliver an older last-known-good build — Outlook 818 over a provisioned 915, a downgrade. The installers now only run for a package that is not provisioned at all; newer builds come from the exact-build / `-Latest` route |
| Verified by running the **whole** script in Windows PowerShell 5.1 with the AppX cmdlets, event logs, downloads and signatures mocked, in the lem-avd-4 scenario (915 provisioned, FSLogix failing on 902/915), the lem-avd-5 scenario (profiles asking 922) and an empty host, with `-Name teams,outlook -Latest -Provision -RemoveOld` and with `-CheckOnly`: no abort, no installer over a provisioned package, 922 provisioned exactly on the lem-avd-5 scenario, exit codes 0/1/2 as expected. Not yet re-run on the hosts |

### 2026-10-01 (2)
| Change |
|--------|
| `Repair-AppxPackageStore.ps1 -RemoveOld` removes every reference a host keeps to an older build of the named packages, after the newest is provisioned: older provisioned copies, older builds registered for any user (for all users, per user where that refuses), and what `AppxAllUserStore` still remembers of them under user, end-of-life, deferred-removal and machine entries — each key backed up to `.reg` first. On the host that kept failing, Outlook 818 was still provisioned next to 915 and 902 still registered, which keeps older builds within reach of a sign-in |
| Deliberately bounded: only packages named one by one (`-Name teams,outlook`; ignored with a wildcard), nothing at all when the build to keep is not provisioned, so no user is left without the app. The `WindowsApps` folders are left to Windows, which owns them and deletes them once nothing references them, and the list in each profile container to FSLogix, which rewrites it at the next sign-out — both are reported. Menu `R` asks for it after provisioning, together with `-Latest` |
| Verified in PowerShell 5.1 with the cmdlets mocked and a scratch `AppxAllUserStore`: keeping 915 removed the provisioned 818, the registered 902 and the three store entries for 902/818 (three `.reg` backups), left 915 everywhere and named the two old `WindowsApps` folders; with nothing provisioned it removed nothing. Not run on a session host |

### 2026-10-01
| Change |
|--------|
| `Repair-AppxPackageStore.ps1 -Latest` provisions the newest Teams / Outlook build there is, not only the build FSLogix failed on. Teams comes from Microsoft's config service, the feed the client itself uses (26246 today, with its MSIX link). Outlook has no such feed — the Store catalog answered 1.2026.818.0 while 915.300 was already on the CDN and in users' profiles — so the newest build that can be proven is used (asked for by FSLogix, registered for a user on the host, or present in `WindowsApps`), and the run says which source it used |
| The run no longer says "Nothing to repair" when a package keeps failing with the right build provisioned. A live run after the exact-build fix showed Outlook 1.2026.915.300 provisioned, FSLogix 26.01, the profiles asking for 902 and 915 — and FSLogix still failing that afternoon, plus 55× `0x80073CF9`. Step 1b now says that is not a version gap (where it used to guess "an old saved version that clears at the next sign-out"), the run exits 1, and step 1c shows the evidence: the newest AppX deployment error with Windows' specific error text and its `Get-AppPackageLog -ActivityID`, and the package's lines in FSLogix's profile log |
| Verified in PowerShell 5.1: `-Latest` against the **real** Teams config service (26246 found newer than a provisioned 26225, with the right MSIX URL) and with Outlook taken from what was seen (915 over 818); the evidence against this machine's **real** AppX log, which printed the specific error text and an ActivityId. Not yet run on the session hosts |

### 2026-09-30 (12)
| Change |
|--------|
| Fixed `Request_UnsupportedQuery: Unsupported or invalid query filter clause specified for property appId` at startup, introduced by the previous commit. Moving the role list out of the shared block into `\` put it **above** the well-known application ids it is built from, so every `ResourceAppId` was empty and the lookup filter became `appId eq ` |
| The constants now sit above the role list in each script rather than inside the shared block, which is where they have to be if the role list is going to use them |
| The parser cannot catch a variable used before it is assigned, so the test now checks the order directly: the application ids must precede the role list, and the role list must precede the shared block. Also proved by executing each scripts prologue and confirming every role resolves to a real GUID |

### 2026-09-30 (11)
| Change |
|--------|
| [`scripts/SharePoint/Revoke-SharePointUserAccess.ps1`](scripts/SharePoint/Revoke-SharePointUserAccess.ps1) reported a real, enabled account as `Not found in Entra ID`. The temporary app was never granted Graph `User.Read.All` — `GroupMember.Read.All` does not allow reading an arbitrary user object — so `GET /users/{upn}` came back `403`, and the catch treated that as the user not existing |
| The same mistake as before in a new place: being refused a lookup is a different fact from the thing not being there, and only one of them is safe to shrug at. It now stops with the missing permission named. That matters beyond the wrong message — without the user resolved, the Entra groups that also grant access are never listed, and that list is the half of the report saying what this script **cannot** revoke |
| Added `User.Read.All` to the roles the revoke script asks for. To keep the report at least privilege, the shared block no longer hardcodes the role list: each script sets `$RequiredAppRoles` before it, and the token-role check validates whatever that script asked for rather than a fixed pair. The report still does not ask for `User.Read.All`, because it expands groups and never reads a user object |
| Verified with 63 checks (10 new): the revoke script asks for `User.Read.All` and the report does not, both still ask for the three they share, the token check follows the per-script list, a `403` on the lookup is fatal and names the permission, and a genuine absence still only warns |

### 2026-09-30 (10)
| Change |
|--------|
| A live tenant-wide run of [`scripts/Reporting/Get-SharePointPermissionsReport.ps1`](scripts/Reporting/Get-SharePointPermissionsReport.ps1) showed the select-ladder rediscovering the same answer on every site: three system lists (`Galerie van thema''s`, `Galerie met basispagina''s`, `Bibliotheek met onderhoudslogboeken`) rejected the same fields on all 131 sites, each costing a wasted round trip and a log line |
| Which fields a list accepts is a property of its **template**, not of the site, so the outcome is now learned once per template and reused. A simulation of the observed pattern over 131 sites puts it at half the round trips (1048 to 528) with every template still landing on exactly the rung it accepts, so no field is lost to the shortcut |
| The log says it once per template instead of once per site — the `[SKIP]` for the User Information List too, which appeared on all 131. Roughly 350 repeated lines removed, which is what was hiding everything else |
| Verified by simulating the ladder across 131 sites with and without the memory: fewer calls, identical resolution per template, and a template that accepts nothing still terminates instead of looping |

### 2026-09-30 (9)
| Change |
|--------|
| Fixed [`scripts/Reporting/Get-SharePointPermissionsReport.ps1`](scripts/Reporting/Get-SharePointPermissionsReport.ps1) failing on **every** site with `Cannot validate argument on parameter 'Kind'. The argument "A" does not belong to the set "U,G"`. Adding the consolidated site-access view taught the checkpoint to *read* a third key kind (`A`) but never widened the `ValidateSet` on the function that *writes* one, so the first web threw and each of the 131 sites reported a failure |
| Introduced alongside the `Toegang` sheet and not caught because the report's test suite lived in a session scratchpad that was cleared between sessions — the cost of that loss, exactly as flagged at the time |
| Added a check for the whole class rather than this one case: every kind written must be in the `ValidateSet` **and** be read back by the resume switch, and every allowed kind must actually be used. Proved it fires by running it against both broken variants — the kind missing from the set, and a kind written but never read, which would silently lose resume state instead of throwing |
| No data was lost. The failed webs wrote error rows carrying their `UnitKey`, and the supersede logic drops those once the web succeeds, so a plain re-run cleans up after itself |

### 2026-09-30 (8)
| Change |
|--------|
| Hardening pass on [`scripts/SharePoint/Revoke-SharePointUserAccess.ps1`](scripts/SharePoint/Revoke-SharePointUserAccess.ps1), driven by reading the code for the failure modes a destructive script has rather than the ones a report has. Three were real |
| **Matching a guest could have revoked the wrong person.** The fallback lookup compared with `-like "*needle*"`, and `an@contoso.com` is a substring of `jan@contoso.com`. Replaced with exact comparison against the UPN, the mail address, the claim suffix and a properly decoded guest login (`jan_partner.com#ext#@tenant` back to `jan@partner.com`, splitting on the last underscore so a local part may contain one). When two different accounts answer to the same address the site is left untouched and the run stops naming both — choosing is the operator's call, not the script's |
| **The audit CSV was written once, at the end.** A run that revoked two hundred things and then died would have left no record of what it removed, which is the one thing a script like this must never do. Rows are now appended as they happen, through the shared helper that retries a locked file and stops the run rather than dropping a row |
| **A `404` on a removal counted as a failure.** It means the grant is already gone, which on a second pass is the normal outcome — a clean re-run would have reported failures. Recorded as `AlreadyGone` instead |
| A failed site collection administrator removal is now loud and counts as a failure: that role reaches every scope in the site, so every other removal there is cosmetic while it stands. The summary says so explicitly rather than reading like a success |
| `-WhatIf` now takes the same branch as a dry run, so it records `WouldRevoke` instead of `Skipped`, which had implied someone declined a prompt |
| [`Test-SharePointAccessScripts.ps1`](scripts/SharePoint/Test-SharePointAccessScripts.ps1) grew from 27 to 50 checks: guest-login decoding including an underscored local part, exact matching against the near-misses a substring test would have accepted (shorter, longer, suffixed domain, another tenant's guest, empty), the audit CSV existing and holding every row mid-run, and a 404 reading as already gone. **Still not verified against a live tenant** |

### 2026-09-30 (7)
| Change |
|--------|
| Removed a dead `-Restart` parameter from [`scripts/SharePoint/Revoke-SharePointUserAccess.ps1`](scripts/SharePoint/Revoke-SharePointUserAccess.ps1). It was declared and its help promised it would `discard any existing checkpoint and start over instead of resuming` - but the script has no checkpoint and no resume, so the switch did nothing and the help described behaviour that does not exist. Found by comparing the parameter block against the comment-based help and the folder readme rather than assuming they agreed |
| Replaced it with a `.NOTES` entry saying why there is deliberately no resume: revoking is idempotent, so a second run finds only what the first did not remove. Re-running after an interruption is both the recovery and the verification, and safer than resuming a partly applied destructive operation from a saved position |
| Verified all 17 remaining parameters appear in the comment-based help and the folder readme, that `Get-Help` no longer mentions `-Restart`, and that the 27 checks in [`Test-SharePointAccessScripts.ps1`](scripts/SharePoint/Test-SharePointAccessScripts.ps1) still pass |

### 2026-09-30 (6)
| Change |
|--------|
| New [`scripts/SharePoint/Trace-SharePointFile.ps1`](scripts/SharePoint/Trace-SharePointFile.ps1): finds where a OneDrive or SharePoint file went — renamed, moved, copied, deleted, restored — by whom and when, from the Unified Audit Log. Before, this meant clicking through the Purview audit search by hand, where a rename chain or a renamed parent folder is easy to miss |
| The trail is followed by item ID and by the path a file was renamed or moved to, so `A → B → C` ends at C. Folder renames, moves and deletes are replayed onto the file's path, because SharePoint writes no record per file for those |
| Times are shown in Brussels time (`Europe/Brussels`, with the UTC offset, summer/winter time handled) and the period is given as Brussels wall-clock time in day-first notation; a bare end date includes the whole day |
| Robustness: the window is read per day, a slice with more than 50,000 records is split (down to 15 minutes), a failing or inconsistent search (`ResultIndex -1`) is retried with backoff, and duplicate records are dropped. [`menu.ps1`](menu.ps1) has it under key `O` |
| Verified: syntax check; runs in PowerShell 7 and Windows PowerShell 5.1 against a **mocked** `Search-UnifiedAuditLog` — rename chain, copy reported but not followed, folder rename and folder recycle replayed onto the file, old URL and `/:w:/r/` sharing link, `-SiteUrl` not matching a neighbouring site with the same prefix, the DST switch on 29 March 2026 (+01:00 → +02:00), and the 50,000-record split. **Not yet run against a live tenant**; the audit field layout (`SourceRelativeUrl`, `DestinationFileName`, `ListItemUniqueId`) follows Microsoft's documented schema |

### 2026-09-30 (5)
| Change |
|--------|
| Documentation now keeps itself current. [`.claude/hooks/sync-docs.ps1`](.claude/hooks/sync-docs.ps1) regenerates [`scripts/INDEX.md`](scripts/INDEX.md) and the readme headers and runs the link check; it is called by a git pre-commit hook ([`.githooks/pre-commit`](.githooks/pre-commit), for changes made by hand) and by Claude Code hooks in [`.claude/settings.json`](.claude/settings.json) (after every edit, in the background). Before, all three had to be remembered — and [`INDEX.md`](scripts/INDEX.md) had already fallen two scripts behind |
| Before Claude finishes, a Stop hook checks that an English readme change was also made in Dutch and French, and that a changed script has its folder readme touched; the git hook warns about the first, because a hook cannot translate. A broken link stops the commit |
| Verified: every mode run against this repository — a clean edit is silent, an injected broken link exits 2 with the file and target, an untranslated readme and an undocumented script each block Stop once (and not a second time), and the Claude hook was seen firing after an edit. The pre-commit hook ran on this commit |

### 2026-09-30 (4)
| Change |
|--------|
| Every readme now exists in three languages: [`readme.md`](readme.md) (English, still the main version), [`readme.nl.md`](readme.nl.md) (Dutch) and [`readme.fr.md`](readme.fr.md) (French) — 66 folders, the root readme including its full Version History. The switcher at the top of each page goes to the same page in the other language, and the breadcrumbs stay within the language you are reading |
| Script-name headings (`### Set-UserManager.ps1`) are not translated, so every `#…ps1` anchor is the same in all three languages; other headings are, with their in-page links adjusted. Parameter names, commands, paths and the literal strings a script prints or writes (Dutch Excel tab names, error messages) stay as they are in every language |
| `Reporting/readme.md` and part of `SharePoint/readme.md` were Dutch in an English set; they were made English first, and the Dutch versions keep the original wording |
| The local-admin password that appears in plain text in [`scripts/Deployment/readme.md`](scripts/Deployment/readme.md) and in this Version History is **not** copied into the Dutch and French versions; there it reads as omitted |
| [`.claude/CLAUDE.md`](.claude/CLAUDE.md) now requires a change to a readme to be made in all three languages, followed by [`Update-ReadmeHeader.ps1`](scripts/Startup/Update-ReadmeHeader.ps1) |
| Verified with [`Test-MarkdownLinks.ps1`](scripts/Startup/Test-MarkdownLinks.ps1): 204 markdown files, every internal link resolves; `Update-ReadmeHeader.ps1 -Check` reports every header current. The translations were checked for structure (sections, tables, line counts against the English), not proofread line by line by a native speaker |

### 2026-09-30 (3)
| Change |
|--------|
| New [`scripts/Startup/Update-ReadmeHeader.ps1`](scripts/Startup/Update-ReadmeHeader.ps1) writes the two lines at the top of every readme: a language switcher (`English · Nederlands · Français`) and the breadcrumb back up the tree, each level linking to its readme in the current language. Hand-kept, those relative paths are what breaks when a folder moves; generated from the folder a readme sits in, they cannot. `-Check` exits 1 on a stale header or a missing language version |
| [`scripts/INDEX.md`](scripts/INDEX.md) regenerated: besides the new script it now also lists [`Revoke-SharePointUserAccess.ps1`](scripts/SharePoint/Revoke-SharePointUserAccess.ps1) and [`Test-SharePointAccessScripts.ps1`](scripts/SharePoint/Test-SharePointAccessScripts.ps1), which had been added without rerunning [`Update-ScriptIndex.ps1`](scripts/Startup/Update-ScriptIndex.ps1) |
| Verified: syntax check clean; run in PowerShell 7 and a `-Check` run in Windows PowerShell 5.1 against this repository — 66 folders, 198 readmes, every header current afterwards |

### 2026-09-30 (2)
| Change |
|--------|
| Every readme now starts with a breadcrumb (`M365-Scripts › scripts › Intune › Desktop`) linking each level back up. Before, 40 of the 66 folder readmes had no way back to their parent except the browser's back button |
| The root readme opens with a `## Folders` table linking every workload folder, so the repository can be browsed from the front page down instead of via [`scripts/readme.md`](scripts/readme.md) only |
| Subfolders are listed under a `## Folders` heading everywhere. [`Device/`](scripts/Device/readme.md), [`Network/`](scripts/Network/readme.md), [`Reporting/`](scripts/Reporting/readme.md) and [`SharePoint/`](scripts/SharePoint/readme.md) mixed them into the Scripts table, [`Intune/Desktop/`](scripts/Intune/Desktop/readme.md) and [`Custom Scripts/Intune/Desktop/`](scripts/Custom%20Scripts/Intune/Desktop/readme.md) used a `Contents` table, [`TenantOnboarding/`](scripts/TenantOnboarding/readme.md) said `Subfolders` and [`LegacyUtilities/`](scripts/LegacyUtilities/readme.md) had no heading at all |
| Verified with [`Test-MarkdownLinks.ps1`](scripts/Startup/Test-MarkdownLinks.ps1): 1,072 internal links across 72 markdown files resolve. Documentation only; no script changed |

### 2026-09-29 (11)
| Change |
|--------|
| [`Repair-AppxPackageStore.ps1`](scripts/Device/Repair-AppxPackageStore.ps1) provisions the **exact** Teams / Outlook build FSLogix fails on. A production host showed Outlook failing 196× in a week — 186× `0x80070490` — for 1.2026.902 and 915 while the host provisioned 818 and the files of both requested builds were on disk. The earlier verdict ("with a current FSLogix the gap is harmless, no action needed") was wrong, and so was chasing it with the installers, which only deliver an older last-known-good build |
| The MSIX for one exact build is on Microsoft's CDN at the versioned URL winget's manifests use (`res.cdn.office.net/.../v2/<version>/Microsoft.OutlookForWindows_x64.msix`, `teamsinstaller.public.onecdn.static.microsoft/production-windows-x64/<version>/MSTeams-x64.msix`) — checked to answer for Outlook 812/818/902/915 and Teams 26198/26225/26246. `-Provision` now takes the newest build FSLogix failed on, downloads it, checks the Microsoft signature, provisions it and reads the provisioned version back; the installer is skipped for that package |
| Error codes are decoded with Windows' own message for every Win32 code instead of a short hand-written list, one code per line: `0x80073D19` turned out to be "An error occurred because a user was logged off" — harmless — and is now labelled as such |
| Verified in PowerShell 5.1: the scenario from that host (FSLogix asking 902 and 915, host at 818) yields 915 with the right URL, a **real** download of that 32 MB MSIX with a valid Microsoft signature, the provisioned version read back (with `Add-AppxProvisionedPackage` mocked), and no target once the host has 915; earlier scenarios unchanged. **Not run on the host itself** |

### 2026-09-29 (10)
| Change |
|--------|
| `Repair-AppxPackageStore.ps1 -Copilot -Provision` now installs the **new**, unified Microsoft Copilot app instead of the old Microsoft 365 Copilot app. The installer added in (9) delivers the old AppX package, which the unification then has to move over; the new app is installed machine-wide by Edge Update. The documented way is used: `Install{C50565E9-...}` = 5 (Force Installs), `UpdaterExperimentationAndConfigurationServiceControl` = 1 (which Force Installs requires) and `CopilotUnificationAllowed{...}` = 1 under `HKLM\SOFTWARE\Policies\Microsoft\EdgeUpdate`, written after a `.reg` backup of that key; then Edge Update's machine task is started and the run waits up to 10 minutes for the app under `EdgeUpdate\Clients`. If it does not appear the old installer is the fallback. `Install` = 0 is never overridden. Diagnosis and verification now count Copilot as present only when the new app is |
| `-Name` understands `teams`, `outlook` and `copilot`, so new Outlook and the new Copilot app on the pool is `-ComputerName lem-avd-4,lem-avd-5,lem-avd-6 -Name outlook,copilot -Provision`. The menu (`R`) takes the same words |
| Found while testing, fixed: under Windows PowerShell 5.1 with `ErrorActionPreference = Stop`, `reg.exe` writing to stderr is a terminating error, so a single failed `.reg` backup would have ended the whole run instead of leaving that key alone. The backup also takes `HKLM:`/`HKCU:` paths now |
| Verified in PowerShell 5.1 and 7: the shorthands (five combinations), the Edge Update install against a mocked policy key and task — values written, backup made, the app "appearing" is picked up — and `Install` = 0 left untouched; the earlier store scenario unchanged. **Not run on a session host**: whether Edge Update installs the app within the 10 minutes there is untested |

### 2026-09-29 (9)
| Change |
|--------|
| `Repair-AppxPackageStore.ps1 -Copilot` diagnoses and restores Copilot. Step 1d reports the Microsoft 365 Copilot app (`Microsoft.MicrosoftOfficeHub`) and the Windows Copilot app (`Microsoft.Copilot`) as registered and provisioned, the unified Microsoft Copilot app that Edge Update installs since the September 2026 unification (read from `EdgeUpdate\Clients\{C50565E9-...}`), the Edge Update version against the 1.3.253.25 it needs, and every policy that keeps Copilot away — `Install` / `Uninstall` / `Update{C50565E9-...}` under `Policies\Microsoft\EdgeUpdate` (with Force Installs overriding Uninstall, as documented), the unification pause, and Windows' `WindowsCopilot` / `WindowsAI` policies machine-wide and per signed-in user. Policies are reported with path and value and fail the run, but are never changed: they come from GPO or Intune |
| With `-Provision` Copilot is installed for all users with Microsoft's documented `M365CopilotDesktopInstaller.exe --quiet --start -p` from `go.microsoft.com/fwlink/?linkid=2325486` — checked today to deliver a Microsoft-signed `xpdBootstrapper` 16.0.19305 — and accepted when either the AppX package is provisioned or the unified app appears under Edge Update. `-Copilot` also puts both packages' Deprovisioned markers in scope, which is what a debloat tool leaves behind. winget has only an `.exe` for it, so `-UseWinget` falls back to the installer. The pool table gained a Copilot column; the menu (`R`) takes `copilot` as the package answer |
| Renumbered the previous entry to (8): it and the SharePoint entry below were both committed as (7) the same afternoon |
| Verified in PowerShell 5.1 and 7: step 1d against this machine's real registry and packages, and against mocked Edge Update policies (Uninstall alone fails the run, Uninstall with Force Installs does not); the orchestrator's pool table with the new column; the installer download and its signature. **Not run on a session host**: the installer's `-p` provisioning and the unified app appearing afterwards are untested |

### 2026-09-29 (8)
| Change |
|--------|
| [`Repair-AppxPackageStore.ps1`](scripts/Device/Repair-AppxPackageStore.ps1) runs across a pool with `-ComputerName lem-avd-4,lem-avd-5,lem-avd-6` (optionally `-Credential`): it copies itself to `C:\IT\AppxRepair` on each host over PowerShell remoting, runs there with the same parameters — the host's own output streams back — and ends with one table across the pool (exit code, FSLogix build, provisioned Teams / Outlook) that names any difference between hosts. A repair is confirmed once for the whole pool, because a remote session cannot answer a confirmation prompt reliably. Wired into the menu (key `R` asks for the hosts) |
| The verify step's advice was wrong. After a live `-Provision` run it said Teams (26225) and Outlook (1.2026.818) were "still older than the 26246 / 902 profiles ask for - bring the other hosts to the same build". But no host is ahead: both apps update themselves per user, and Microsoft's installers provision a last-known-good build that is behind that, so the profile will always be ahead of every host and provisioning newer only lasts until the next update. What decides whether it hurts is FSLogix: from 2210 HF4 (Teams) / 25.06 (Outlook) it registers by family name and the gap is harmless (now reported as OK, exit code 0); on an older build the advice is to update FSLogix. Such a gap no longer counts as something to provision |
| A transcript that will not start — as in some remote and RMM sessions — no longer aborts the repair; it is a warning |
| Verified in PowerShell 5.1 and 7: the orchestrator against two unreachable hosts (each named with its WinRM error, the pool table, exit code 1), and the version gap with a mocked FSLogix above and below the minimum. **Not run against real session hosts**: no WinRM to lem-avd-4/5/6 from here, so copying, the remote run and the pool table with real values are untested |

### 2026-09-29 (7)
| Change |
|--------|
| Added [`scripts/SharePoint/Revoke-SharePointUserAccess.ps1`](scripts/SharePoint/Revoke-SharePointUserAccess.ps1) — the counterpart to the permissions report. It finds every place one named user holds access and removes it: site collection administrator first (it overrides everything below it, so leaving it would make the rest cosmetic), direct role assignments on sites, sub-sites, lists, folders and single files, SharePoint group membership, and the `SharingLinks.*` groups that carry "Anyone with the link" and "Specific people". Reporting is the default; nothing changes without `-Apply`, and every run writes a CSV of what was found and what happened to it |
| It deliberately refuses two things and says so loudly. An Entra ID group grant is not revoked — the group *is* the grant, and removing the user from SharePoint would leave access in place while looking like it was closed; the group is named in the CSV under `Action = CannotRevoke` so offboarding is visibly two steps. A grant to `Everyone` is left alone for the reverse reason: removing it revokes access for the whole tenant rather than for this person |
| The app-only authentication and SharePoint REST layer is shared with the permissions report, byte for byte, delimited by `SHARED BLOCK START/END`. It took four live runs against a tenant to get right, and a second copy that quietly drifts is a correctness risk in the script that deletes permissions. To make that shareable the report's banner moved above the block and the temp app name now comes from `$TempAppNamePrefix`; the report's behaviour is unchanged |
| Added [`scripts/SharePoint/Test-SharePointAccessScripts.ps1`](scripts/SharePoint/Test-SharePointAccessScripts.ps1), which asserts the two copies are identical (printing the first differing line when they are not) and drives the revocation funnel for real: a dry run records its intent and executes nothing, `-Apply` executes and records, a failure lands in the audit trail instead of vanishing, and the refusals above stay refused. 27 checks, all passing, runnable from any directory |
| Added the menu entry (key `W`) and documented both scripts in the SharePoint folder readme, the repository tree and this category list. Verified every `docs` anchor in that readme resolves |

### 2026-09-29 (6)
| Change |
|--------|
| [`Repair-AppxPackageStore.ps1`](scripts/Device/Repair-AppxPackageStore.ps1) no longer tries to re-register a package that has been superseded. The first live run tried `aimgr_0.20.61.0` in step 3 and got `0x80073D06` ("a higher version 0.20.62.0 of this package is already installed"): an old version whose status is not Ok while a newer one of the same package sits next to it is not damage but Windows waiting to remove it, and re-registering it can never succeed. The diagnosis now compares versions per package name, architecture and resource id, reports these as *Superseded* in one grey line, and leaves them out of the repair count; step 3 also treats a `0x80073D06` that turns up anyway as "left for Windows" rather than a warning |
| Verified offline in PowerShell 7 and 5.1 with that exact pair (0.20.61.0 `Modified` next to 0.20.62.0 `Ok`): reported as superseded and not counted, while a genuinely damaged package in the same run is still picked up for re-registration. Not yet re-run on the host where it happened |

### 2026-09-29 (5)
| Change |
|--------|
| [`Repair-AppxPackageStore.ps1`](scripts/Device/Repair-AppxPackageStore.ps1) can take Teams and new Outlook from winget (`-UseWinget`): `winget download` of `Microsoft.Teams` / `Microsoft.Outlook`, whose manifests point at the MSIX on Microsoft's CDN, then `Add-AppxProvisionedPackage` with any dependencies winget brought, so the package lands for all users rather than only for whoever ran `winget install`. Every file must carry a valid Microsoft signature. winget's manifests lag behind Microsoft's installers (checked today: Teams 26198 against the 26246 profiles ask for, Outlook 1.2026.812 against 902), so the installers stay the default and the run warns when winget's build is older than what the profiles ask for. `-WingetId` provisions any other app the same way |
| It now lists every app that fails, not only the ones in the package store: step 1c reads the AppX deployment log and the FSLogix Apps log over `-Days`, grouped per package, with the error codes named (`0x80073D02` in use, `0x80073CF6` registration failed, ...), the versions asked for and whether this host has their files; `0x80070490` first, top 15 |
| Found while testing, fixed: `Get-WinEvent` throws a terminating error for a provider that is not registered — any machine without FSLogix — which `-ErrorAction SilentlyContinue` does not catch, so the FSLogix check would have aborted the run there; all event reads go through one wrapper now. And the winget progress filter held two non-ASCII characters, which Windows PowerShell 5.1 reads as ANSI in a file without BOM and then fails to parse — the whole script would not have run from NinjaOne. The file is pure ASCII again, checked |
| Verified in PowerShell 7 and 5.1: the failing-apps overview against this machine's **real** AppX log (20 packages, codes translated, capped at 15); a **real** `winget download` of `Microsoft.Outlook` (32 MB MSIX, hash verified by winget, signature by the script) through to a mocked `Add-AppxProvisionedPackage`; and the earlier mocked store scenario, unchanged. Provisioning itself and the Teams download (271 MB) were not run here |

### 2026-09-29 (4)
| Change |
|--------|
| New [`scripts/Device/Repair-AppxPackageStore.ps1`](scripts/Device/Repair-AppxPackageStore.ps1): repairs AppX packages that fail with `0x80070490` and an empty path ("Deployment Register operation ... from:  (AppxManifest.xml)"), for any package — the same error came back for `Microsoft.OutlookForWindows` on a host where only Teams had a repair, and `Update-TeamsClient.ps1 -RepairAppxStore` is scoped to `MSTeams` by design |
| The errors on that host were logged by `Apps (Microsoft-FSLogix-Apps)`, which is a different cause than a damaged store: FSLogix saves each user's packages by exact version in `AppxPackages.xml` and replays them at sign-in (`InstallAppxPackages`, default on), so a host that provisions another build — or none — answers `0x80070490`. The script reads those events and compares the version the profiles ask for with what the host provisions, checks the FSLogix build against the first releases that register Teams (2210 HF4) and Outlook (25.06) by family name, and with `-Provision` puts Teams / new Outlook back for all users with Microsoft's own installer (`teamsbootstrapper.exe -p`, Outlook `Setup.exe --provision true --quiet --start-`; both links checked to resolve to Microsoft's CDN, both signature-checked before running) |
| The store repair generalises the Teams one and adds what makes it safe to run on a whole store: every registry key is exported to a `.reg` backup before it is removed, and not removed when the backup fails; with a wildcard `-Name`, system and framework packages are reported but never touched and Deprovisioned markers (how bloatware removals are remembered) are left alone; a user registration also counts as orphaned when its package has no files anywhere, not only when its SID has no profile. Editing `StateRepository-Machine.srd` or `AppxPackages.xml` was researched and deliberately left out — both unsupported |
| Wired into [`menu.ps1`](menu.ps1) as Device key `R` (diagnose unless you confirm the repair; asks separately about provisioning), documented in the Device readme |
| Verified offline in PowerShell 7 and 5.1 with mocked AppX cmdlets, FSLogix events and a scratch `AppxAllUserStore` in HKCU: the empty-path Teams package is found as a ghost, its user and machine entries as orphans, an Outlook entry for a SID without profile as orphan, the older provisioned Teams and the missing Outlook as needing provisioning, a framework ghost and a Deprovisioned marker are left alone with `*` and included when named, and the `.reg` backup is written. That run also caught `-Name A,B` arriving as one string through `powershell.exe -File` (and through the script's own relaunches), now split. **Not run on a live host**: no elevation or AVD host was available here, so the removals, the registry edits and both installers have not been exercised for real |

### 2026-09-29 (3)
| Change |
|--------|
| [`Restore-MailboxMessages.ps1`](scripts/Exchange/Restore-MailboxMessages.ps1) now handles "everything from this date until now" as a first-class case: `-After` gained the aliases `-From` and `-Since`, and the menu entry (`L`) asks whether to restore only that day or everything since. The window itself already allowed it, but the audit search ran as one query over the whole window, and a single search session stops at 50,000 records tenant-wide — over a few weeks that silently dropped actions on the mailbox being restored |
| The audit log is now searched one day at a time, each day in its own session with paging, and only records that mention this mailbox are kept in memory. A day that alone exceeds 50,000 records is named in a warning |
| A long window can reach past what the mailbox still keeps, which would read as "nothing was deleted". The run now warns when the window starts before the mailbox's `RetainDeletedItemsFor` (14 days by default) and the mailbox is not on hold, and when it starts more than 180 days ago, beyond the usual audit retention |
| Verified offline in PowerShell 7 and 5.1: `-Since` binds to `-After`; a mocked `Search-UnifiedAuditLog` over a 2.6-day window was called per day with UTC boundaries, the last slice ending at the window's end, day one paged across two calls in one session, and only this mailbox's records kept. The retention warning is **not** exercised - it needs a live `Get-Mailbox` |

### 2026-09-29 (2)
| Change |
|--------|
| [`Restore-MailboxMessages.ps1`](scripts/Exchange/Restore-MailboxMessages.ps1) no longer needs the Mailbox Import Export role to bring deleted mail back. The first real run stopped at "Get-RecoverableItems is not available" and skipped every deleted message, while the role is in no role group by default — so on most tenants the deleted part simply did nothing |
| Without the role the run now switches to Graph and restores **everything** deleted in the window, not only what the audit log saw: every message in Deleted Items and Recoverable Items\Deletions whose modification time falls in the window goes back. Audited deletions (`MoveToDeletedItems`, `SoftDelete`, which Exchange audits for the owner by default) go to the folder the record says they left, with the actor matched exactly by MessageId; the rest go to the Inbox. A message deleted *out of* Deleted Items goes to the Inbox too, because putting it back in Deleted Items is not recovering it. Hard-deleted items (Purges) are out of Graph's reach and reported as `Unreachable` instead of silently missing |
| Message lookups now also search `recoverableitemsdeletions`, which `/messages` does not cover, so a message that was moved and then deleted is found in either part. The moved and deleted parts share one lookup / move / report path instead of two copies |
| Verified offline in PowerShell 7 and 5.1 against a mocked Graph: moved-then-soft-deleted goes back to the original subfolder, deleted-from-Deleted-Items goes to the Inbox, unaudited items from both folders are restored, an item deleted five days earlier is left alone, a hard delete is reported `Unreachable`, and preview and `-Apply` issue exactly the expected moves. **Not run against a live tenant**; in particular whether Graph allows a move out of `recoverableitemsdeletions`, and whether a move changes `lastModifiedDateTime`, are untested |

### 2026-09-29
| Change |
|--------|
| New [`scripts/Exchange/Restore-MailboxMessages.ps1`](scripts/Exchange/Restore-MailboxMessages.ps1): put back the messages that were moved or deleted in one mailbox on a given day, and say who did it. There was no way back from a bad archive run or a mass delete short of restoring by hand in Outlook, and no answer to "who did this" without writing an audit log query from scratch |
| Deleted messages go back through `Get-/Restore-RecoverableItems` (Deleted Items, Recoverable Items, Purges), one `EntryID` at a time, filtered on the moment of deletion — Exchange knows the original folder itself. After an `-Apply` the folders are read again, and anything still there is reported as `NotRestored` instead of trusting the cmdlet's silence |
| Moved messages have no such memory: neither Graph nor Exchange records where a moved message came from. The Unified Audit Log does, so every audited `Move` is traced to the **first** folder the message left that day, located over Graph by Internet MessageId and moved back through `$batch`. Folders are matched on their path as the audit log writes it, which is in the mailbox's own language (`\Postvak IN\Projecten`). Moves out of Deleted Items or Recoverable Items are skipped, because those were restores and reversing them would delete the message again |
| The same audit records name the actor — account, owner/delegate/admin, client (Outlook, OWA, Graph app with app ID), IP — per message in the CSV, as a grouped "who moved / deleted what" table on screen, and as a raw `_Audit.csv`. Deletions are attributed by subject and nearest time, since recoverable items carry no MessageId. The run also lists which of the four actions are not audited on the mailbox, because by default the owner's own `Move` is not, and a missing record would otherwise read as "nobody did it" |
| Archive items without an audit record (e.g. after [`Move-InboxToArchive.ps1`](scripts/Exchange/Move-InboxToArchive.ps1)) are listed by modification time and only moved to the Inbox with `-UnauditedArchiveToInbox`, since reading or flagging also changes that time. Graph access reuses the REST-only three-route pattern of [`Remove-PhishingMessage.ps1`](scripts/Exchange/Remove-PhishingMessage.ps1), so it runs next to the Exchange session without the MSAL clash. Added to the Exchange submenu as `L` (preview first, then `-Apply`) |
| Verified offline only, in PowerShell 7 and Windows PowerShell 5.1: syntax check, audit-record parsing against fabricated records (mailbox filter, UTC to local time, logon types, client labels, subject/time attribution), and the whole moved-message path against a mocked Graph — a chain of moves going back to the first folder, Dutch folder names, a user's restore left alone, an already-returned message skipped, an audited Archive item kept out of the unaudited list, and the resulting move requests. **Not yet run against a live tenant**: the exact output properties of `Get-RecoverableItems`, how it interprets the filter times, and whether the audit records carry `InternetMessageId` for every client are all untested |

### 2026-09-28 (6)
| Change |
|--------|
| The restart `-RestartIfNeeded` triggers waited 60 seconds on a host where nobody could be watching. The countdown exists to warn people, so it now only applies when there are people: with someone signed in it is `-RestartDelaySeconds` and `shutdown /a` stops it; with nobody signed in - the normal case at boot, and the guaranteed one on a session host whose pool is set to drain - it restarts within seconds. A few seconds are kept so the run's own log line is written before the shutdown starts |
| Considered draining logons from inside the guest (`change logon /drainuntilrestart`) to close the window between the task starting and the restart, and dropped it: a host being restarted this way already has its pool on drain, so the guest-side switch would only duplicate what the pool guarantees - and it would leave logons blocked on any run that crashed before re-enabling them |
| Verified that `change.exe` and `chglogon.exe` exist on this Windows 11 build and that `change logon /query` reports "Session logins are currently ENABLED" while exiting 1, which is why the idea was measured before being dropped rather than after |

### 2026-09-28 (5)
| Change |
|--------|
| [`Update-TeamsClient.ps1`](scripts/Device/Update-TeamsClient.ps1) can now repair the host instead of only diagnosing it. A session host where every route answered `0x80070490` had an `AppxAllUserStore` full of entries Windows can no longer resolve, and the honest advice at that point was "redeploy" — which is not what anyone wants to hear about a machine that is otherwise fine |
| `-RepairAppxStore` does it in two steps. A package whose files are still on disk is re-registered from its own manifest (`Add-AppxPackage -Register`), which rebuilds the store's knowledge of it and usually makes the ordinary removal work again. What survives that is removed key by key: registrations under a SID with no profile on this host (also under `EndOfLife` and `DeferredRemoval`), a machine-wide `Applications` entry whose manifest is gone, and the `Deprovisioned` marker that refuses the provision outright. Each key is named with its full registry path before it goes, and nothing outside MSTeams is ever touched |
| Preflight reports those orphans whether or not the switch is given, so `-CheckOnly` is the diagnosis and the repair is a separate decision — the same shape as `-ClearOrphanedAddInRegistration` for Windows Installer |
| `-UseWinget` attacks it from the other side: winget downloads the MSIX, checking it against the SHA256 in its own manifest, and the bootstrapper provisions that file with `-p -o`. The deployment then has an explicit source rather than a store entry it has to resolve, and the run knows which build it installed. winget's manifest lags the config service — measured at `26198.304.4946.9672` against a `26246` build — and the run says so when it does |
| Run as System, winget is not on `PATH` at all: its alias is a per-user MSIX shim. It is resolved from `Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*` instead, verified by emptying `PATH` and watching the fallback find it |
| The store reader was run against this workstation's live `AppxAllUserStore`, where it found two genuine orphans (`S-1-0-0` and a deleted profile's SID under `EndOfLife`), reported none for healthy packages and kept every path inside the store. The **removal** was exercised for real against a store rebuilt under `HKCU`: 9 MSTeams entries, 6 orphaned, all 6 removed, while the healthy registrations, the `Staged` entry, another product's orphan and the dead SID's own key all survived |
| `winget download` was measured end to end: 271 MB in 23 seconds, no Store account, its own hash check, a valid `O=Microsoft Corporation` signature, the staging folder emptied first and removed after. Against a host whose package store is actually damaged, both switches are **untested** — no machine here has one |
| The reboot line stopped naming a reason. It said "MSI returned 3010" while three different things set it, and pointed readers at an installer that had never run |

### 2026-09-28 (4)
| Change |
|--------|
| [`Update-TeamsClient.ps1`](scripts/Device/Update-TeamsClient.ps1) was giving advice that could not help. A production session host answered `0x80070490` ("Element not found") for **every** holder of the package, `NT AUTHORITY\SYSTEM` among them, and the script still said to drain the host and sign users out — on a host that was already drained and had nobody on it. A registration the package store cannot find is not a user holding the package |
| Preflight now says whether a pending removal has anybody left to wait for. `Installed(pending removal)` only completes at a sign-out, so a SID with no profile under `ProfileList` waits for an event that can never happen; those are reported apart from the ones that really are waiting, and only the latter flag a reboot |
| It also checks the one state nothing recovers from by itself: a package the store lists whose `InstallLocation` is gone, or that has no install location at all. That single line explains the whole failure — every removal answers `0x80070490` because there is nothing to remove, and provisioning the same version answers it too |
| When the per-user removal answers that code for every holder, the run says so plainly, and the final failure changes its advice with it: not "drain the host", but that `Remove-AppxPackage`, the bootstrapper and DISM all read the same inconsistent store, so none of them can repair it — a pooled session host is redeployed from its image, a personal one is repaired in place |
| Exercised against the strings that host really printed, plus a live SID from this machine as the contrast case: four orphaned profiles read as orphaned, a real profile still reads as "sign them out", a package with no install location is flagged, and a healthy package stays quiet. The **remediation** is untested — this machine has no damaged package store to try it on |

### 2026-09-28 (3)
| Change |
|--------|
| [`Init-TempDisk.ps1`](scripts/Device/TempDisk/Init-TempDisk.ps1) repaired the temp disk and configured the pagefile on it, and then let the machine run the rest of that session without one - Windows reads the pagefile configuration at boot and never re-reads it, so the boot that had to rebuild `D:` is exactly the boot on which the pagefile does not exist. Added `-RestartIfNeeded`, which closes that gap instead of waiting for the next boot |
| A script that runs at every boot and may restart the machine is a reboot loop waiting to happen, so it only fires when all of this holds: the run finished clean (a failed run never restarts - that would hide the failure behind a reboot), the disk is there, the pagefile is configured on it, and the only thing missing is that this session is not using it |
| Nobody may be signed in, connected or disconnected. Sessions are counted as one `explorer.exe` per interactive desktop rather than by parsing `query.exe`, whose column headers follow the display language and would read an empty list out of a Dutch session host. `-RestartEvenIfUsersSignedIn` overrides it where the countdown is warning enough |
| At most one restart per `-RestartCooldownMinutes` (default 60), remembered as a round-trip timestamp under `HKLM:\SOFTWARE\ICTKanon\InitTempDisk` - a locale-formatted timestamp written by one run and read by another is how a cooldown quietly stops working. A second restart for the same thing means the first one did not help, and the run says so instead of repeating it |
| The restart goes through `shutdown.exe` with a 60-second countdown and the planned "Operating System: Reconfiguration" reason, so anyone on the machine sees it coming, `shutdown /a` stops it, and it is not reported as an unexpected restart |
| [`Register-InitTempDiskTask.ps1`](scripts/Device/TempDisk/Register-InitTempDiskTask.ps1) now deploys the task with `-Quiet -RestartIfNeeded`, and says at registration time whether the task may restart the machine. `-ScriptArguments '-Quiet'` leaves the restart out |
| Numbered the two earlier entries of today: two identical `### 2026-09-28` headings had ended up in the history, which reads as one change split in half |
| Verified on this machine under PowerShell 5.1 and 7: session detection names the signed-in account (so this machine would refuse to restart), a missing marker reads as `$null`, a marker written and read back parses to a `DateTime` and blocks a second restart inside the cooldown, a 90-minute-old marker allows one, and a corrupt marker degrades to "no marker" rather than throwing |
| Also measured how a failing `shutdown.exe` reports itself, because the code branches on it: `shutdown /a` with nothing pending answers 1116 and sets `$LASTEXITCODE` in both 5.1 and 7.6 without throwing, so the exit-code branch is the one that runs. The `try` around it stays for `$PSNativeCommandUseErrorActionPreference`, which can turn that into a terminating error on 7.4 and later |
| **No restart was triggered from this session** and the guards remain unverified against a live Azure VM |

### 2026-09-28 (2)
| Change |
|--------|
| [`Update-TeamsClient.ps1`](scripts/Device/Update-TeamsClient.ps1) crashed in preflight on any machine where the meeting add-in is registered nowhere: `The property 'Count' cannot be found on this object`. `$x = if (...) { @() }` assigns `$null`, because an empty array written to the pipeline is zero objects — the `@()` has to go around the whole `if`, not inside its branches. Reproduced against the committed version and fixed; all five paths through the reporting function now pass, and the three empty ones provably threw before |
| A package the AppX stack refuses to remove no longer ends the run. `Remove-AppxPackage -AllUsers` answered `Catastrophic failure` on a session host carrying two `MSTeams` versions, and `-ErrorAction` does not cover a terminating error, so it needed a `try`/`catch`. The run continues and the provision upgrades in place whatever survived — aborting there had left the host with the add-in uninstalled and no Teams put back |
| The add-in **uninstall** moved from step 6 to step 8, next to the install that replaces it. The sweep had already moved there; leaving the uninstall behind meant any later failure produced the same outcome by a different route. Everything destructive about the add-in now sits with the thing that undoes it |
| Exit codes are readable. `teamsbootstrapper.exe` answers with an HRESULT, which PowerShell prints as a large negative integer: "exit code -2147023728" says nothing, `0x80070490 - Element not found` says where to look. MSI codes stay plain numbers, and an HRESULT outside the Win32 facility falls back to bare hex rather than inventing a meaning |
| A failed provision now tries Microsoft's documented machine-wide uninstall (`teamsbootstrapper.exe -x -m`) once and provisions again before giving up, and the failure it raises names the usual cause on a session host: a package held by a signed-in user. **Untested** — that recovery has not yet run on a host that needed it |
| That failure message then ate its own exit code: `-f` binds tighter than `+`, so formatting a concatenated string applied the format to the last piece only and printed a literal `{0}`. Reported from a live host, where it hid the very code the reader needed |
| The bootstrapper prints its own verdict, and running it hidden threw that away. `Invoke-Installer` can now capture stdout and stderr, and the bootstrapper's lines are echoed as `bootstrapper:` output. Verified with a process that prints a JSON verdict and exits non-zero; without the switch nothing is captured and no temp files are left behind |
| Capturing that output then broke every verdict, and only on the runtime that matters: under Windows PowerShell 5.1 a redirected `Start-Process -PassThru` reports no exit code at all unless the process handle is touched first. A `-x -m` that printed `{"success": true}` was reported as a failure. Measured on both runtimes; `$null = $proc.Handle` fixes it, and PowerShell 7 never had the problem |
| Success is now decided by the bootstrapper's own JSON verdict where there is one, not by an exit code that can go missing. Non-JSON, foreign JSON and malformed output all fall back to the exit code rather than guessing |
| Preflight counted `PackageUserInformation` entries and called a package "installed for 1 user profile" when its only entry was `S-1-5-18` **staging** it — not a user, and not installed. It reads the states now: staged-only says so, and a package whose entries read `Installed(pending removal)` is reported as already removed and waiting for those users to sign out, by name, with the reboot flagged. Measured against the two packages a production session host really reported |
| `Remove-AppxPackage -AllUsers` is all or nothing, so a single profile it cannot touch fails the whole call. When it does, the package is now removed per user instead, and whoever still holds it is named — "a signed-in user is holding it" is not actionable until you know which user. The SID comes out of `PackageUserInformation`'s string form, which differs across builds, exercised against seven shapes including the Entra `S-1-12-1` one. **Untested** against a package that really refuses removal |
| The AppX log then answered the question the bootstrapper's `0x80070490` hid: it is provisioning `MSTeams_26246...`, the build registered for one profile, and Windows cannot find that package's files. Two MSTeams versions side by side with one of them broken is the state to look for |
| A failed provision now also prints the AppX deployment errors Windows logged, which carry the reason its `0x80070490` hides — "Unable to install because the following apps need to be closed &lt;package&gt;". Read-only, 10 ms, and quiet when the log holds nothing recent |
| Preflight reports an AppX package whose `Status` is not `Ok`. Windows considering a package Modified or Tampered is exactly what makes `Remove-AppxPackage` answer `Catastrophic failure` and the provision fail after it, and it was invisible until now |

### 2026-09-28
| Change |
|--------|
| Added [`scripts/Device/TempDisk/Init-TempDisk.ps1`](scripts/Device/TempDisk/Init-TempDisk.ps1): an Azure VM's ephemeral temp disk is wiped on every deallocate, resize or host move and comes back RAW, offline or without its drive letter. Windows reads the pagefile configuration at boot and never re-reads it, so a pagefile configured on `D:` that is not there at boot is simply never created and the machine pages on `C:` again - or runs with no pagefile at all. The script restores the volume as `D:` and points the pagefile back at it |
| A temp disk that only lost its drive letter is given the letter back rather than reformatted, recognised by its label (`Temporary Storage`) or by the `DataLoss_Warning_Readme.txt` Azure writes on the resource disk. Only a RAW, non-boot, non-system disk is ever initialised: an empty temp disk and an unformatted data disk look identical from the outside, so a disk with partitions is reported and left alone, and more than one RAW candidate makes the script refuse to guess and ask for `-DiskNumber`. `-Force` plus `-DiskNumber` is the only route to formatting a disk that still carries data |
| An optical drive holding `D:` is moved out of the way first - Windows hands `D:` to the DVD on an image with no temp disk and never gives it back, which is the second way the pagefile ends up on `C:` |
| The run distinguishes the pagefile as *configured* (registry) from the pagefile *in use* (this session) and says which is which, instead of reporting success for a change that only lands at the next restart. Configuring a pagefile on a drive that could not be restored is a hard failure rather than a setting Windows quietly ignores |
| Added [`scripts/Device/TempDisk/Register-InitTempDiskTask.ps1`](scripts/Device/TempDisk/Register-InitTempDiskTask.ps1), built on a draft that had two faults: its `-ScriptSourcePath` default referenced `$ScriptTargetDir`, a parameter declared *after* it, so the default expanded to `\Init-TempDisk.ps1` and never resolved; and the copy ran with `-ErrorAction SilentlyContinue`, so a missing source registered a boot task against a file that is not there - it then fails at every boot with nobody watching. The source now defaults to the copy next to the script, a missing source is a hard error, and the task is verified to exist after registering |
| Documented both in a new [`scripts/Device/TempDisk/readme.md`](scripts/Device/TempDisk/readme.md), added the folder to the [`Device/`](scripts/Device/readme.md) readme and the repository tree, gave the root readme a **Temp Disk & Pagefile (Azure / AVD)** entry, and wired [`Init-TempDisk.ps1`](scripts/Device/TempDisk/Init-TempDisk.ps1) into [`menu.ps1`](menu.ps1) as Device key `V` (defaults to `-CheckOnly` unless you confirm the repair) |
| Verified: both files parse clean through [`Test-PowerShellSyntax.ps1`](scripts/Startup/Test-PowerShellSyntax.ps1). The read-only helpers were run for real on this Windows 11 machine under both PowerShell 5.1 and 7 with `Set-StrictMode -Version Latest` - free-letter search, optical-drive lookup, temp-disk and RAW-candidate detection, and the pagefile state read (which correctly reported automatic management on and `C:\pagefile.sys` in use) - and the scheduled-task trigger, principal and settings objects were constructed and their values checked. **Not yet verified on a live Azure VM or session host**: no disk was initialised, no pagefile was changed and no task was registered from this session |

### 2026-09-25 (20)
| Change |
|--------|
| Turned the `Pivot toegang` sheet around to nest **site → group → person** instead of site → person → group. That is how SharePoint actually grants access — a site has groups, groups have people — so collapsed it lists the groups on a site and expanded it names everyone they let in |
| Added `Pivot per persoon` (person → site → group) so the other direction is still answerable from the same sheet: what does this one person reach, and through what. That is the offboarding question, and a site-first pivot cannot answer it |
| A directly granted person has no group, and in a three-level hierarchy that empty middle level reads as missing data rather than as "granted without a group". `ViaName` now says `(direct toegekend)` for those rows instead of being blank |
| Verified locally with 240 checks across eleven suites: both pivots read back out of the workbook with their row fields in the intended order, and a direct grant is named rather than empty |

### 2026-09-25 (19)
| Change |
|--------|
| Audited the documentation against the repository rules rather than assuming it was complete, and found three gaps. Parameters checked out: all 19 of [`Get-SharePointPermissionsReport.ps1`](scripts/Reporting/Get-SharePointPermissionsReport.ps1)'s parameters are present in the comment-based help and in the folder readme's parameter table, with nothing stale in either |
| [`scripts/Reporting/readme.md`](scripts/Reporting/readme.md) had no `## Scripts` table at all, while [`Exchange/`](scripts/Exchange/readme.md), [`Entra/`](scripts/Entra/readme.md), [`Device/`](scripts/Device/readme.md) and [`SharePoint/`](scripts/SharePoint/readme.md) all have one. Added it, covering all five entries in the folder — not just the new script — so the table describes the folder rather than the last change to it. Every link and anchor in it was verified to resolve |
| The root readme's `### 📊 Reporting` category listed only the Computer Last Logon and Licensing reports. All three SharePoint reporting scripts were missing from it, including two that predate this work. Added an entry for [`Get-SharePointPermissionsReport.ps1`](scripts/Reporting/Get-SharePointPermissionsReport.ps1) and short ones for [`Get-SharePointStorageReport.ps1`](scripts/Reporting/Get-SharePointStorageReport.ps1) and [`Remove-SharePointFileVersionsByDate.ps1`](scripts/Reporting/Remove-SharePointFileVersionsByDate.ps1) |
| The repository tree still described the permissions report as going "to CSV", which stopped being true when `-Excel` was added; it now says CSV + Excel. The [`menu.ps1`](menu.ps1) label said "who has access to what, at every level", which describes the old shape of the report rather than the consolidated per-site view it now leads with |
| Confirmed no action needed for [`f.ps1`](f.ps1): its index rebuilds itself when a script's write time changes, so `f-sharepointpermissionsreport -Excel` picks up new parameters without `f-refresh` |

### 2026-09-25 (18)
| Change |
|--------|
| [`scripts/Reporting/Get-SharePointPermissionsReport.ps1`](scripts/Reporting/Get-SharePointPermissionsReport.ps1) answered "which grants exist" but not the question people actually open it with: **who can reach this SharePoint, and how did they get there.** `Rechten` said a group had rights, `Groepen` said who was in it, and nothing joined the two — "Site Owners has Full Control" plus "Site Owners contains five people" is not an answer. Added `SharePoint_Permissions_SiteAccess_<ts>.csv` (worksheet `Toegang`): one row per person per site, with the group their access runs through, that group's id, and the permission level |
| Consolidated per site collection on purpose: someone reaching thirty folders in one site through the same group is one row, not thirty. A different level or a different group is a separate row, because that is different access. Per-scope detail stays behind `-IncludeEffectiveAccess` |
| Three things deliberately do not fall out of that view: a directly granted person appears as themselves with `ViaType = Direct`; `Everyone` and `Everyone except external users` resolve to nobody but get a row naming the claim, since they are exactly what a reviewer is looking for; and with `-SkipGroupExpansion` the direct grants still show, only the group members are missing |
| Added `SiteTitle` — a consolidated view of 130 sites is not readable as 130 URLs, and the root web title is only known while that web is being scanned, so it is captured there and looked up per row. Added `ViaId` alongside `ViaName` after comparing with [NovaPoint](https://github.com/Barbarur/NovaPoint/wiki/Solution-Report-PermissionsReport), which carries `GroupId` next to `AccessType` for the same reason: a title like `Site Owners` repeats on every site in the tenant |
| Added a `Pivot toegang` sheet nesting site → person → group against permission level, filtered by external and access type. NovaPoint puts its users in one `Users` column as a list; each user gets their own row here instead, which reads less compactly but is the difference between being able to filter or pivot on a person and not |
| Verified locally with 238 checks across eleven suites (25 new): a group grant lists its people with the group name, id and level; the same access through the same group on a deeper scope is not repeated while a different level is; direct grants, `Everyone` claims and `-SkipGroupExpansion` all behave as described; checkpoint keys are one per row and unique; and the sheet and its pivot are read back out of the workbook. **Not yet verified against a live tenant** |

### 2026-09-25 (17)
| Change |
|--------|
| Made the `-Excel` workbook from [`scripts/Reporting/Get-SharePointPermissionsReport.ps1`](scripts/Reporting/Get-SharePointPermissionsReport.ps1) genuinely pivotable. Checked first rather than assumed: numeric columns already arrive in Excel as numbers, not text, so aggregation was never the problem — the obstacle was `PermissionLevels`, which SharePoint fills with several levels at once (`Read; Limited Access`). A pivot treats each combination as its own value, so `Full Control` and `Full Control; Limited Access` land on separate rows |
| Sheets carrying `PermissionLevels` now get a `PrimaryPermission` column immediately beside it, holding the single strongest level of that grant. `Limited Access` always loses to a real level — SharePoint adds it automatically for traversal — and a custom level ranks above `Read` but below `Full Control`, because it was created deliberately and should not vanish behind a built-in. Dutch and English level names are both recognised, which matters on a Dutch-language tenant |
| Added three ready-made pivot sheets: `Pivot rechten` (site × permission level, count of grants, filtered by principal and scope type), `Pivot principals` (principal × scope type, count of scopes, filtered by site and external), and `Pivot groepen` (group × external member, count of members, filtered by site and group type). Each only references columns its source sheet actually has, and a missing or narrowed source is skipped rather than producing a broken pivot |
| Pivot creation is best-effort and isolated: a failure warns and leaves the data sheets untouched, on the same principle as the workbook itself not being allowed to cost the CSVs |
| Verified locally with 213 checks across ten suites (33 new): the level ranking across single, joined, reversed, Dutch, custom, case-varying and empty inputs; the derived column landing next to the original on the right sheets and not on the others; and the pivots read back out of the package with the right row, column, data and filter fields, skipping absent sources and column-less sheets. **Not yet verified against a live tenant** |

### 2026-09-25 (16)
| Change |
|--------|
| Added `-Excel` to [`scripts/Reporting/Get-SharePointPermissionsReport.ps1`](scripts/Reporting/Get-SharePointPermissionsReport.ps1): one `.xlsx` alongside the CSVs with a worksheet per report — `Samenvatting`, `Rechten`, `Groepen` and, with `-IncludeEffectiveAccess`, `Effectief` — each a real Excel table with filter dropdowns and a frozen header, using the same `ImportExcel` pattern as [`Get-DistributionGroupMembers.ps1`](scripts/Exchange/Get-DistributionGroupMembers.ps1) |
| The CSVs are still always written and the workbook is built from them, not instead of them. They are what the scan streams into and what a resumed run appends to, so they exist regardless — and a workbook that fails to write (module missing, file open, out of memory) then costs a convenience copy rather than the report |
| A worksheet stops at 1,048,576 rows and drops the rest without complaint, so sheets are capped at 1,000,000 with a warning naming the sheet and the CSV that still holds everything. On a large tenant only `Effectief` realistically approaches that |
| Fixed a real gap in the group membership while wiring this up: an Entra ID group granted **directly** on a site, list or item never passes through `/sitegroups`, so it was the one kind of group whose membership the report never listed — only the first ten names in `MemberPreview`. Those groups now get their own rows in the Groups output, resolved to people, recorded once per group rather than once per grant, and sharing the schema the SharePoint-group rows already use |
| Added the Excel prompt to the [`menu.ps1`](menu.ps1) entry |
| Verified locally with 180 checks across nine suites (20 new): a workbook is written and read back with all four sheets in order and their rows intact, a re-run replaces rather than appends, absent/empty/missing sources are skipped, nothing to write leaves no file behind, an oversized sheet is capped rather than truncated by Excel, a directly granted Entra group is listed once with its real members, SharePoint groups are left to `/sitegroups`, and both group sources share one schema. **Not yet verified against a live tenant** |

### 2026-09-25 (15)
| Change |
|--------|
| First complete live run of [`scripts/Reporting/Get-SharePointPermissionsReport.ps1`](scripts/Reporting/Get-SharePointPermissionsReport.ps1): 130 webs, 2066 lists, 453 unique scopes, 1554 grants, 83 sharing links, 13 external grants, 111 `Everyone` grants. Three faults in the output itself, all found by reading the produced CSVs rather than the logs |
| `-IncludeEffectiveAccess` produced an empty file on a tenant with 1554 grants and over a thousand resolved members. The guard was `if ($IncludeEffectiveAccess -and $EffectiveRows)`, and **an empty `List[object]` is falsy in PowerShell** — so the test failed on the very first row and the list could never fill, which kept it empty, which kept the test failing. Now an explicit `$null -ne` check |
| 121 of the 158 "could not be read" rows were a single hidden system list, `Lijst met gebruikersgegevens` (template 112, the User Information List), on every site. SharePoint rejects `/items` on it with `400` at every `$select` width, including the narrowest rung of the ladder. Its items are directory records rather than content, so item-level scopes there mean nothing for an access review — the item sweep now skips template 112 and says so, while still reporting the list's own scope |
| The remaining 37 were stale: error rows written by the interrupted earlier attempt, carried into the final CSV by the resume even though those lists succeeded on the retry. A failed unit is deliberately left unmarked so it is retried, but nothing removed its old rows. Detail rows now carry the `UnitKey` that produced them, and a row whose unit is marked complete is dropped when the final CSV is written — so the incompleteness count describes the file the reader opens |
| The summary is now built from the published detail CSV instead of the partial, so its counts and the file agree |
| Verified locally with 158 checks across eight suites (26 new): effective rows are emitted one per resolved user with the group they came through, a null list is tolerated, superseded error rows are dropped while still-failing and unkeyed ones survive, nothing else is lost, and the summary counts only what was published. Two earlier assertions were found to be mis-parenthesised — one of them a false pass — and corrected. **Fixes to this run's findings are not themselves verified against a live tenant yet** |

### 2026-09-25 (14)
| Change |
|--------|
| Second live run of [`scripts/Reporting/Get-SharePointPermissionsReport.ps1`](scripts/Reporting/Get-SharePointPermissionsReport.ps1) authenticated cleanly — the token role check caught the replication delay on its first attempt and waited it out, SharePoint and Graph both accepted their tokens, and 130 webs were discovered and started scanning. Two scan-level faults then repeated on every site |
| `ConvertTo-PermissionRows` rejected an empty role assignment collection: a `Mandatory [object[]]` parameter refuses `@()`, so every system list that has unique permissions but no remaining role assignments (`User Information List`, `Converted Forms`, `Bibliotheek met onderhoudslogboeken`) failed with `Cannot bind argument to parameter 'RoleAssignments'`. Fixed with `[AllowEmptyCollection()]` — a scope with no assignments legitimately produces no rows |
| More consequentially, an unreadable role assignment list was indistinguishable from an empty one. `Invoke-SPGet` swallows `403`/`404` and returns `$null`, which `Get-SPCollection` turns into an empty collection — and an empty collection reads as "nobody has rights on this scope". Role assignment reads now use `-ThrowOnDenied`, so a refusal becomes an error row saying the permissions are unknown rather than a silent claim that there are none. This is the only place a 403 is not skipped, because it is the only place where "not allowed to look" would be misread as a finding |
| The gallery lists (`Galerie van thema's`, `Galerie met basispagina's`) answered `400 Bad Request` to the item `$select`, because their schema does not carry every field it names, and a 400 is not something retrying fixes. The item sweep now steps down a four-rung ladder of progressively narrower `$select` clauses until SharePoint accepts one; every rung keeps `Id` and `HasUniqueRoleAssignments`, so the worst case loses a file name rather than the list's unique scopes. Only a 400 triggers narrowing — a denial, a throttle or a view threshold answers the same way however few fields are asked for |
| Verified locally with 132 checks across seven suites (21 new): an empty assignment set no longer crashes and produces no rows, a denied read throws with "unknown rather than empty", a genuine empty `200` stays empty end to end, an ordinary sweep still skips a 403, every ladder rung keeps the fields the scan depends on, and only a 400 narrows. **The scan phase past web 11 is still unverified against a live tenant** |

### 2026-09-25 (13)
| Change |
|--------|
| First live run of [`scripts/Reporting/Get-SharePointPermissionsReport.ps1`](scripts/Reporting/Get-SharePointPermissionsReport.ps1) got past SharePoint — the certificate credential worked and the preflight reported `SharePoint accepted the token (root web: ...)` — and then failed on Graph with `401` while retrieving sites. Cause: the Graph token was minted immediately after the app roles were granted, before the grant had replicated, so it carried no `roles` claim at all. Graph answers such a token with `401`, not `403`, and because the token was cached for its full hour every one of the six retries was handed the same dead token back |
| Tokens now have to prove themselves: `Get-ResourceToken` takes `-RequiredRoles`, decodes the issued JWT, and refuses to cache a token whose `roles` claim is missing what the run needs. It keeps re-minting (up to 15 attempts, backoff capped at 20s) until the grant appears, then fails with the missing role named. Both tokens are validated up front — Graph for `Sites.Read.All` + `GroupMember.Read.All`, SharePoint for `Sites.FullControl.All` — so a replication delay is waited out before the scan starts rather than discovered 130 sites in |
| `Invoke-GraphGet` now drops the cached token and re-mints once on a `401`, the same recovery `Invoke-SPGet` already had. Retrying the request alone could never have worked against a poisoned cache entry |
| A user-supplied `-ClientId` app is deliberately **not** role-validated: a working app may hold broader roles (`Directory.Read.All` instead of `GroupMember.Read.All`), and rejecting it would be a false failure. The SharePoint preflight still catches a genuinely under-permissioned app |
| `Get-JwtClaim` returns `$null` for an empty token instead of throwing a parameter binding error, and `Disconnect-MgGraph` no longer leaks its context object as a stray `ClientId`/`TenantId`/`Scopes` table after the summary |
| Verified locally with 111 checks across six suites (23 of them new, driving the real `Get-ResourceToken` against a fake token endpoint): a role that arrives late is waited out and only the token carrying it is cached, a role that never arrives fails loudly with nothing cached, the backoff grows and stays capped, and caching stays isolated per resource. **The corrected Graph path is not yet verified against a live tenant** |

### 2026-09-25 (12)
| Change |
|--------|
| Hardened [`scripts/Reporting/Get-SharePointPermissionsReport.ps1`](scripts/Reporting/Get-SharePointPermissionsReport.ps1) for long tenant-wide runs. A `401` was still recoverable into a per-site error row: the per-list handler rethrew it but the per-web handler caught it again, so a credential that stopped working mid-run would have written one error row per remaining site — the same failure shape the certificate fix had just removed. Both handlers now let a 401 through and the run stops |
| The parallel item lookups read the bearer token once per library instead of once per wave. A library with enough unique scopes outlives a token, so the tail of it would have failed with no indication why. The token is now re-read before every wave |
| Item enumeration no longer materialises an entire library before filtering. `Invoke-SPCollectionPaged` hands each page to a callback and only items that actually have their own scope are kept — a million-item library now costs one page of memory instead of a million live objects |
| Added a paging guard: SharePoint echoing back an identical `nextLink` used to be an infinite loop against a live tenant, and is now detected and stopped |
| Checkpoint keys moved from the JSON state file to an append-only `.keys.partial.log`. Rewriting a sorted list of every completed key after every list is quadratic; on a tenant with thousands of lists the checkpoint cost more than the scanning. A torn final line from a killed process is tolerated — that unit is simply re-scanned |
| A failed CSV write (the partial open in Excel) is retried five times and then stops the run. It previously threw while the unit was already marked complete, so those rows were gone from the report for good |
| A list that fails now costs that list, not the rest of the site: per-list error handling writes an error row, keeps whatever the list already produced, and deliberately leaves the unit unmarked so a resumed run retries it. A failed item sweep gets its own row, because without it the list looks like it simply had nothing with unique permissions |
| Per-item workers no longer report `401`/`403` as an empty permission set — only `404` (item genuinely deleted mid-scan) means "no permissions". Claiming an unreadable item has no rights on it is worse than saying so |
| Added an output-folder write probe before authenticating, an abort when discovery finds no sites at all, and a `trap` that removes the temporary Full Control app registration on any unhandled error |
| Suppressed the `Set-MgRequestContext` context table that leaked to stdout, and the run now closes by stating whether every targeted scope was read or how many were missed |
| Verified locally with 53 checks across four suites: principal/claims parsing, CSV row-schema consistency (six row shapes, 26 columns each), certificate and signed-assertion generation, and HTTP behaviour driven through a fake transport — 401 aborts after one re-auth, 403/404 stay per-object, 429 retries to success, paging loops are broken, locked files are retried, and the checkpoint log survives a torn line. **Still not verified against a live tenant** |

### 2026-09-25 (11)
| Change |
|--------|
| Fixed [`scripts/Reporting/Get-SharePointPermissionsReport.ps1`](scripts/Reporting/Get-SharePointPermissionsReport.ps1) producing an empty report against a live tenant: all 130 webs came back `[SKIP] Web not accessible with the current permissions`. The temporary App Registration authenticated with a client secret, and **SharePoint Online refuses every app-only token obtained with a secret** — `401` with `x-ms-diagnostics: ... Unsupported app only token`. Graph accepted the same credential, so site enumeration worked and only the `_api` calls failed, which is why it looked like a per-site permissions problem |
| The temporary app is now given a certificate instead of a secret. It is generated in memory with `CertificateRequest`, registered as a `keyCredential`, and used to sign an RFC 7523 client assertion — it never touches the certificate store or disk, so an interrupted run leaves nothing behind |
| `Invoke-SPGet` no longer swallows `401` alongside `403`/`404`. A `401` is never per-site — it is the same answer for the whole tenant — and treating it as "this one site is not accessible" is what turned a single credential fault into 130 lines that read like findings. It now throws, with the `x-ms-diagnostics` reason and, for the secret case, what to do about it |
| Added a SharePoint preflight: one call against the tenant root after connecting, before enumerating anything. Whether SharePoint accepts the credential is one yes/no for the whole run, so it now costs one request to find out instead of a full sweep |
| `-ClientSecret` now warns at startup that the SharePoint half will fail, and the docs say the same. Graph-only alternatives were considered and rejected: Graph has no endpoint for web role assignments, SharePoint groups, site collection administrators or named permission levels, and would need one `/permissions` call per item instead of one `HasUniqueRoleAssignments` sweep per list |
| Suppressed a stray `ClientTimeout RetryDelay MaxRetry` table that `Set-MgRequestContext` printed to stdout at the end of every run |
| Verified locally: 21 checks on certificate generation and the signed assertion, including that the signature verifies against the certificate's public key, that `x5t` matches its SHA-1 hash, and that nothing is written to `Cert:\CurrentUser\My`. The corrected auth path itself is **not yet verified against a live tenant** |

### 2026-09-25 (10)
| Change |
|--------|
| [`Convert-MarkdownToHtml.ps1`](scripts/Startup/Convert-MarkdownToHtml.ps1) rendered every numbered list as empty bullets. `$Matches` is one variable per scope: the list branch captured the item text, then ran a second `-match` to decide whether the list was ordered, and that second match threw the capture away. Dashed lists survived only because their second match failed and left `$Matches` alone. Both captures now come from one match and the marker decides the type without matching again |
| A fenced code block indented to line up inside a numbered step kept that indentation, so anyone copying the command out of the page copied leading spaces with it. The fence's own indentation is now stripped from its content — and only that: a block fenced at column 0 keeps every space, which is what the sample output of the script needs |
| Verified on the regenerated page: 36 list items and **none** empty, still parses as XML, 29 tables and 19 code blocks intact, the indented command comes out clean, and the seven code blocks that legitimately start with whitespace still do |

### 2026-09-25 (9)
| Change |
|--------|
| You can now click from a readme straight to the script it describes. Every script name in a folder readme's `Scripts` table linked to a section further down the same page, never to the file — so the readme told you what a script did but gave you no way to open it. 172 links across 47 readmes now point at the file, with a `([docs](#…))` link beside them for the section that was there before |
| 34 of those were not links at all: a script name in a table cell, set in backticks, with nothing behind it. Those are file links now too |
| Added [`scripts/Startup/Test-MarkdownLinks.ps1`](scripts/Startup/Test-MarkdownLinks.ps1), because links that are never checked are links that quietly rot. It walks every `.md` and fails on two things: a relative link to a file that is not there (percent-encoded spaces decoded first, the way GitHub serves them), and an anchor with no heading behind it. Anchors are resolved the way GitHub builds them, including the `-1`/`-2` suffix for repeated headings |
| It found three anchors that had never worked: `#watch-rdslivesps1` had an `s` too many for `### Watch-RDSLive.ps1`, and two links in the Intune readme used `#detect--remediate-…` where the heading `### Detect- / Remediate-StuckWin32AppEnforcement.ps1` produces `#detect---remediate-…` — three hyphens, because the slash becomes nothing and the spaces around it each become one. Nobody would find that by eye |
| Invisible characters are stripped from both the heading and the link before they are compared. Without that, the four emoji entries in the root table of contents read as broken: the heading and the link both carry a variation selector, which is not a letter and not a digit. Reproducing GitHub's slugger byte for byte on characters nobody can see is not the point — establishing that a link and a heading correspond is |
| Verified: 854 internal links across 71 markdown files all resolve, all 181 files parse, and the script index is current at 178 scripts. `L` added to the menu for the link check, alongside `X` for the index |

### 2026-09-25 (8)
| Change |
|--------|
| The Teams IT Glue procedure now also exists as a styled HTML page, [`scripts/Device/Update-TeamsClient-ITGlue.html`](scripts/Device/Update-TeamsClient-ITGlue.html), for pasting into IT Glue or printing. It is **generated** by the new [`scripts/Startup/Convert-MarkdownToHtml.ps1`](scripts/Startup/Convert-MarkdownToHtml.ps1) rather than written by hand: that document changed six times in two days, and a hand-made copy would have been wrong by the next morning |
| The converter covers what these documents actually use — headings, tables, fenced code blocks including the indented ones inside numbered steps, blockquotes, both list kinds, rules, and inline code, bold, italic and links — and passes anything else through as text instead of guessing. `-Check` writes nothing and exits `1` when the committed page has fallen behind its markdown, which is what a hook or pipeline would call |
| Void elements are emitted self-closing, so the page parses as XML as well as HTML. That is how it was verified rather than by looking at it: the output parses, and holds 29 tables, 190 rows, 19 code blocks and 46 headings, with **no** table containing a row that disagrees with its header width. Also checked: no `**`, backtick or `](` left anywhere in the rendered text, and the ✅/❌ and accented characters survive |
| Both `-Check` failure paths exercised: a page that does not exist yet, and a markdown file that has moved on — each exits `1` with the reason. The generated-on line is excluded from the comparison, so an unchanged document does not report a difference |
| Added as menu key **M**, documented in [`scripts/Startup/readme.md`](scripts/Startup/readme.md), and [`scripts/INDEX.md`](scripts/INDEX.md) regenerated — adding a script had made it stale, which `Update-ScriptIndex.ps1 -Check` reported |

### 2026-09-25 (7)
| Change |
|--------|
| Merged the readable parts of an older IT Glue version of this same procedure into [`Update-TeamsClient-ITGlue.md`](scripts/Device/Update-TeamsClient-ITGlue.md): the one-sentence statement of what the procedure covers, and the ✅/❌ shape for "does this script fit this ticket", including the two user complaints that version listed and ours did not — Teams hanging on startup and Teams closing unexpectedly |
| The two versions disagreed on who may run the update — the older one puts it at level 1, ours at level 2 — which serves a service desk worse than either answer on its own. The blanket level is replaced by a "wie mag wat" table that assigns a level per action: checking stays level 1, the update and the repairs sit at level 2, and the three switches that skip the signature check or edit the Windows Installer database sit at level 3. Changing the escalation policy is now one table, not a re-read of the document |
| Deliberately not merged, measured against the script rather than judged by eye: that version's parameter table covers 8 of the 16 parameters, states that classic Teams is never touched (`-RemoveClassicTeams` does exactly that), and shows two sample output lines the script does not produce — `Found MSTeams ...` and `Teams installation completed` |
| Numbering fix: two entries were both labelled `(4)`. The IT Glue sync is newer than the script-index entry above it, so it is now `(6)` and sits in the order the work happened |

### 2026-09-25 (6)
| Change |
|--------|
| The IT Glue document was brought back in line with the script, checked by comparing its text against the parameter block rather than by reading it: it was missing `-BootstrapperUrl` and `-SkipSignatureCheck` entirely, and its NinjaOne variable table was missing `removeWebRtcRedirector`, `clearOrphanedAddInRegistration`, `skipSignatureCheck` and `webRtcUrl`. All three documents now cover all 16 parameters, and the variable table all 16 environment variables |
| That same comparison found a real gap in the script: every other text field could be set from a NinjaOne variable except `bootstrapperUrl`, which was simply never read. An admin who set it would have had it silently ignored. It is read now, alongside `webRtcUrl` |
| Two statements in the IT Glue document were no longer true. "It does not touch classic Teams" is only true without `-RemoveClassicTeams`, and the plain-language summary still promised that every copy of the add-in is removed during an update - that sweep is now conditional on a replacement being installable |
| Added a level 3 walkthrough for the `1612` + `1638` deadlock the production host hit: what each code means, the read-only command that says whether Windows Installer still has its cached MSI, which of the two outcomes needs `-ClearOrphanedAddInRegistration`, and the note that starting Teams and restarting Outlook gives a user the meeting button back in the meantime |

### 2026-09-25 (5)
| Change |
|--------|
| Finding a script on GitHub meant guessing which of the 56 workload folders it was under and opening readmes until it turned up. There is now one page that answers it: [`scripts/INDEX.md`](scripts/INDEX.md) lists all 176 scripts A-Z with a link to the file, a link to its folder readme, and what it does — Ctrl-F instead of a hunt |
| The page is **generated**, by the new [`scripts/Startup/Update-ScriptIndex.ps1`](scripts/Startup/Update-ScriptIndex.ps1), so it cannot drift from the files the way a hand-kept table does. `-Check` reports a stale index without writing (exit `1`), which is what a hook or a pipeline would call; a run that finds the page current writes nothing at all |
| Descriptions come from the scripts themselves: the `.SYNOPSIS` block, joined across the lines it wraps over rather than taking the first line, which was leaving half-sentences like "Grant Full Access and/or Send As delegate rights on one mailbox, a CSV list of" in the table. Where a synopsis opens with a sentence and then lists its cases, the lead-in is kept and the list is not dragged in behind it |
| For the older scripts that have no `.SYNOPSIS`, a leading `#` comment block is used instead — but only a real header. A single comment line sitting straight on top of code describes that line, not the script: `# URL van de theme` above a `$ThemeUrl` assignment was being read as a description, which is worse in a table than a blank |
| The eight scripts that still ended up with nothing got a real `.SYNOPSIS` instead of a blank cell: [`add-lock.ps1`](scripts/Intune/Desktop/Add%20Lockscreen%20to%20start%20and%20desktop/add-lock.ps1), [`add-shortcut-lock.ps1`](scripts/Intune/Desktop/Add%20Lockscreen%20to%20start%20and%20desktop/add-shortcut-lock.ps1), [`logic-permissies.ps1`](scripts/Graph/logic-permissies.ps1), [`Test-OpenVpnDiagnostics.ps1`](scripts/Device/Test-OpenVpnDiagnostics.ps1), [`Deploy-OfficeTheme.ps1`](scripts/Custom%20Scripts/Intune/Desktop/Deploy-OfficeTheme.ps1), [`Restart-Time-Sync.ps1`](scripts/Device/Time%20sync/Restart-Time-Sync.ps1), [`Test-PowerShellSyntax.ps1`](scripts/Startup/Test-PowerShellSyntax.ps1) and [`functies.ps1`](scripts/Startup/functies.ps1). All 176 scripts now describe themselves, so the index has no "without a description" section left |
| Documented the two scripts no readme mentioned at all: [`Phising-rollout.ps1`](scripts/Entra/Phising-rollout.ps1) in [`scripts/Entra/readme.md`](scripts/Entra/readme.md) (the two-way sync between the phishing-resistant MFA rollout and registered groups, what counts as registered and why the default is an AAGUID filter) and [`Get-FSlogix-errors.ps1`](scripts/RDS/Get-FSlogix-errors.ps1) in [`scripts/RDS/readme.md`](scripts/RDS/readme.md) (what the FSLogix diagnostic collects and that it must run on the session host). Its header still pointed at a filename that no longer exists, and named a real customer in the example; both corrected |
| The root `Menu` table had drifted from [`menu.ps1`](menu.ps1) — `I`, `T`, `S` and `P` were missing. Synced, and `X` added for the index generator, which is also in the Startup readme and the repository tree |
| Verified: all 176 files parse; the generator is idempotent (a second run reports "already up to date" and writes nothing); `-Check` exits `0` when current; every markdown link in the repository resolves, percent-encoded folder names included; and [`f.ps1`](f.ps1) still finds both the new script and the newly described ones |
| Numbering fix: two entries below were both labelled `(3)`. Renumbered to the order the work actually happened in |

### 2026-09-25 (4)
| Change |
|--------|
| Correction to the previous entry: the `1638` on the meeting add-in was **not** caused by `-Force` downgrading the client. Measured on the host itself, the registered add-in was `1.25.28902` and the MSI being installed `1.26.21803` - newer, and still refused. This MSI declines to install while any other copy of the add-in is registered, whichever version that is. The version comparison added in the last change would therefore not have prevented the failure; the guard now asks whether a registration survived the uninstall, which is the thing that actually decides it |
| `-ClearOrphanedAddInRegistration` is the way out of the state that host is in. An uninstall answering `1612` means Windows Installer has lost the cached MSI it needs and can no longer remove the product by any supported means, while its registration keeps refusing every reinstall. The switch makes the installer forget that one product: its keys under `Installer\Products`, `Installer\Features` and `Installer\UserData\S-1-5-18\Products`, its entry under the upgrade code, and the Programs and Features entry. What MsiZap used to do, scoped to one product, only after msiexec has proved it cannot, off by default, and every key through `ShouldProcess` |
| Finding those keys needs the ProductCode as Windows Installer's 32-character "packed" GUID. That transform was validated before anything used it to point at keys for deletion: of 57 GUID-named uninstall entries on a workstation, the 32 with machine-wide product data all mapped onto an existing packed key with an identical `DisplayName`, and the 25 that did not are per-user installs living under the user's own SID |
| Verified read-only against three real products: each yields three product keys plus exactly one upgrade-code entry, and the constructed path is readable as written. A bogus product code returns nothing and an unknown product returns no keys, so the cleanup cannot fire on thin air. **Untested:** the removal itself, and the reinstall that should follow it |

### 2026-09-25 (3)
| Change |
|--------|
| Added [`scripts/Reporting/Get-SharePointPermissionsReport.ps1`](scripts/Reporting/Get-SharePointPermissionsReport.ps1) — an exhaustive read-only SharePoint Online permissions report: site collection admins, web role assignments including inheritance breaks, SharePoint groups with their full membership, list and library role assignments, every folder and item with a unique scope, sharing links with their kind, external/guest principals, `Everyone` grants, and Entra group grants resolved to transitive membership. Four CSVs: detail, per-site summary, group membership, and — behind `-IncludeEffectiveAccess` — one row per resolved user per scope with the group the access runs through |
| Inheritance is followed the way SharePoint models it: an item is only reported as its own scope when `HasUniqueRoleAssignments` is true, so the CSV is a map of the permission structure rather than a row per file. Site discovery is deliberately redundant — Graph `getAllSites`, then sub-sites through both Graph and SharePoint REST (`/_api/web/webs`), de-duplicated on URL — because Graph omits classic sub-webs |
| Authentication had to go app-only: role assignments are not readable through Graph at all, and are not covered by SharePoint's Read/Write/Manage application roles either — only `Sites.FullControl.All` can enumerate them. The script signs in interactively once, creates a short-lived App Registration with that role plus Graph `Sites.Read.All` and `GroupMember.Read.All`, and deletes it again on exit. Despite the Full Control role it only ever issues `GET`: it never writes and never changes a permission. `-ClientId`/`-TenantId` with a secret or certificate skips the temporary app |
| Resumable like the other long SharePoint scans: a checkpoint per completed list, keyed on a hash of the scan parameters, so an interrupted tenant run continues instead of starting over; `-Restart` discards it. Checkpoint files are only cleaned up once the final CSVs are written, so their presence is itself the signal that a run was interrupted |
| Documented in [`scripts/Reporting/readme.md`](scripts/Reporting/readme.md) (coverage, authentication, the four output files, checkpoints, full parameter table, examples), added to the root repository tree and to [`menu.ps1`](menu.ps1) under Reporting as `P` — which also prompts for tenant or single site, scope, and whether to write the effective-access CSV |
| The repository tree also listed neither this script nor [`Remove-SharePointFileVersionsByDate.ps1`](scripts/Reporting/Remove-SharePointFileVersionsByDate.ps1); both are in it now |
| **Untested by me**: this script has not been run against a live tenant in this session — only its syntax was checked. The temporary App Registration path, the throttling retries and the checkpoint resume are unverified here and should be exercised on a pilot tenant, starting with `-SiteUrl` and `-Scope Site`, before a tenant-wide run |

### 2026-09-25 (2)
| Change |
|--------|
| An add-in that is registered but whose files are gone no longer counts as installed. That is exactly the state the failed run above left behind, and the script would have answered "Teams is up to date - nothing to do" on it: the work decision only looked at the Programs and Features entry, while the DLL check that spots this was report-only |
| [`Update-TeamsClient.ps1`](scripts/Device/Update-TeamsClient.ps1) no longer deletes a working meeting add-in before knowing it can install a replacement. A production `-Force` run removed every copy in step 6 and then failed in step 8 with `1638`, leaving the session host with no add-in at all. The sweep moved to step 8, behind the version comparison: an older MSI than the registered add-in now means the sweep and the install are skipped and the working add-in is left exactly as it is |
| Three things had to line up for that, and all three are now handled. `-Force` on a host whose build is newer than the published one is a **downgrade**, which the version check now warns about by name. An add-in uninstall answering `1612` means Windows Installer lost its source, so it is retried against its own cached MSI under `C:\Windows\Installer` (via `Installer\UserData\S-1-5-18\Products\*\InstallProperties`, `LocalPackage`); when that is gone too, the run says the registration cannot be removed and what it will cause. A `1638` on the add-in is now a warning rather than an abort, so verification still runs and reports what Outlook is actually left with |
| Preflight prints the registered add-in version instead of just "is installed" - that single number was the whole diagnosis of the failure and it was the one thing not on screen |
| The AppLocker line no longer prints an empty summary when `SrpV2` exists with no rule collections under it (as on the production host): it says "no rule collections configured, so it blocks nothing" |
| Verified: the cached-package lookup against real installed products, and that asking for a product this machine lacks returns nothing without throwing; the version guard in all four combinations (older, newer, equal, unparsable); the empty-collection AppLocker line. **Untested:** the `msiexec /x <cached msi>` retry and the `1638` warning path on a live host |

### 2026-09-25
| Change |
|--------|
| [`Update-TeamsClient.ps1`](scripts/Device/Update-TeamsClient.ps1) stops crying wolf about AppLocker. The SlimCore blocker check warned whenever `HKLM:\SOFTWARE\Policies\Microsoft\Windows\SrpV2` existed — which it does on any fleet that has ever written a single Exe rule — so the warning fired on machines where nothing was blocked at all. A check that always fires is a check nobody reads |
| The policy is now read instead of detected, on the three points that decide whether it can stop the MSIX: only the packaged-app (`Appx`) collection applies, because an MSIX never meets the `Exe`/`Msi`/`Script`/`Dll` rules; a collection holding rules whose enforcement is *not configured* is enforced all the same, per Microsoft, and only an explicit `EnforcementMode = 0` lets everything through; and nothing is enforced at all while the Application Identity service (`AppIDSvc`) is stopped, which is now said out loud rather than assumed either way |
| The report is something a technician can act on: the registry path, the mode per collection, the service state and the first five `Appx` rule names with their action. A rule that already allows the packages by name is reported as `[ OK ]`; a rule allowing anything signed by `O=MICROSOFT CORPORATION` is reported as probably sufficient, with a note to check it has not been narrowed to one product name |
| A policy found on a session host is now named as a `[SKIP]` reference line instead of being hidden: it blocks nothing there, because the staging happens on the endpoint, but it is usually the same GPO — so the thing worth checking is whether it also reaches the endpoints |
| Every branch exercised against a stubbed policy tree: an enforced `Exe` collection with no `Appx` collection produces no blocker (the old false positive, gone); an enforced `Appx` collection with no matching allow rule warns with path and rule count; explicit-SlimCore and Microsoft-publisher allow rules produce their two different notes; `EnforcementMode = 0` reads as audit only; enforcement-not-configured-with-rules reads as enforced; seven rules print five and `... and 2 more`; a missing `SrpV2` key produces nothing. Confirmed silent on this machine, which has no AppLocker policy. **Untested** against a live enforced AppLocker policy on a real endpoint |
| Known limitation, documented rather than hidden: the allow-rule match is a text match on the rule XML, so a broad rule that names neither Microsoft nor the packages (`PublisherName="*"`) would really let SlimCore through but is still reported as a blocker. The rule names printed beside it are what settles that |

### 2026-09-24
| Change |
|--------|
| [`Update-TeamsClient.ps1`](scripts/Device/Update-TeamsClient.ps1) answers the question the inventory could not: preflight now reads the `Microsoft Teams VDI` events from the Application log on any session host — not just with `-AvdOptimizations` — and translates the codes from Microsoft's connection error table, so a plain `-CheckOnly` reports whether users are actually optimized instead of only whether the parts are installed |
| `24002`/`24010` say the user is on SlimCore, `16002` that an endpoint still has no plugin, `16389`/`10083`/`1951` that policy on the endpoint blocks the MSIX. A zero `errc` is deliberately not in the table: it means that phase raised no error, and printing "OK" next to a real failure in the other phase would be a lie |
| The query uses `-FilterXPath`, because `Get-WinEvent -FilterHashtable @{ ProviderName = ... }` throws when the provider has never written an event — which is the normal case on a healthy non-VDI machine. Measured: 357 ms and a soft error when absent, 104 ms when present |
| New `-RemoveWebRtcRedirector` removes the old optimization, retired 1 October 2026. Mutually exclusive with `-AvdOptimizations` and refused before the UAC prompt, reuses the `msiexec /x` + stale-`1605`-entry path proven for classic Teams, and leaves `IsWVDEnvironment` set because SlimCore needs that flag too. Off by default: an endpoint that cannot do SlimCore and no longer finds the redirector silently falls back to rendering media on the session host |
| Both docs corrected where they still told a technician to look for SlimCore on the session host |

### 2026-09-20 (8)
| Change |
|--------|
| "The add-in still does not load" now gets an answer instead of a status. A registration that is present but not loading is checked against the three causes that leave no trace in `LoadBehavior` itself, each reported as a `why:` line: a bitness mismatch between Outlook and the registered loader, Outlook having parked the add-in in its `DisabledItems`/`CrashedAddins` resiliency lists, and a group policy overriding the user's load behaviour |
| The resiliency check decodes the binary values in that user's hive and matches on the add-in path, so it reports the one cause a technician cannot see from `LoadBehavior` at all — Outlook disables a crashed add-in and keeps it disabled, which is why ticking the box back on does not stick |
| When nothing on the machine blocks it, it says that too, which is also an answer: what remains is a full Outlook restart and a user who has signed in to Teams at least once |
| Both detections exercised: an x86 loader path against this x64 Office produces the bitness reason, and a planted binary `CrashedAddins` value is decoded and reported. A healthy registration produces no `why:` line |

### 2026-09-20 (7)
| Change |
|--------|
| A full reinstall now removes **every** copy of the meeting add-in before installing the new one, not just the one the MSI knows about: the machine-wide folder, the per-profile folders under `%LOCALAPPDATA%\Microsoft\TeamsMeetingAdd-in`, and the per-user COM registrations in each loaded hive |
| That closes the loop on the `LoadBehavior 2` this script has been chasing for two days. A copy left behind during a reinstall is precisely what becomes a per-user registration shadowing the fresh machine-wide one while pointing at files that no longer exist — which is how `admin` ended up registered against add-in `1.24.19202` |
| `Get-TeamsAddInFolder` verified against this device: it finds the real per-profile copy. The `-WhatIf` plan shows the folder and both CLSID views being removed before the reinstall. The removal itself reuses mechanics already proven live in the classic-Teams and repair tests, but the sweep as a whole runs for the first time on a production host |

### 2026-09-20 (6)
| Change |
|--------|
| Corrected a check that was looking in the wrong place: the script warned `SlimCore packages not found` on session hosts, but Microsoft stages SlimCore **on the endpoint**, not on the VM — *"Step 3: SlimCore MSIX staging and registration on the endpoint ... the plugin silently executes this step, without user or admin intervention"*. The warning was noise on every session host, and a check that looks in the wrong place does not fail, it lies |
| The report is now context-aware. On a session host it confirms the Teams build against the documented minimum `24193.1805.3040.8975` and states that SlimCore belongs on the endpoint. On an endpoint it reports whether the packages are staged, and checks the three policies Microsoft documents as blocking that staging, each with the Teams error code it surfaces: `BlockNonAdminUserInstall` (16389), `AllowAllTrustedApps` (15615) and AppLocker (10083) |
| Also settled the version question from last week: Windows App for Windows `2.0.352.0` is the documented minimum on the endpoint, and the classic Remote Desktop client is no longer supported for this at all |
| Resilience: MSI exit code `1641` (success, reboot already initiated) counted as a failure and aborted the run. It is now a success with a reboot flagged, alongside `3010` |
| Verified on both sides: this endpoint reports `Microsoft.Teams.SlimCoreVdiHost.win-x64 2026.31.1.16`; with `RDInfraAgent` faked the session-host wording appears instead; the three blockers were exercised against stubbed registry reads |

### 2026-09-20 (5)
| Change |
|--------|
| A clean production run confirmed three earlier fixes on a real session host: the redirector repaired in place (`The download is the installed version (1.56.2603.20001)`, so the previous run really did upgrade 1.54 → 1.56), the add-in resolved from the staged package after provisioning, and the whole flow finished at exit `0` |
| It also pinned down the one remaining warning: `BAKKERPARTNERS\admin` has a registration pointing at add-in `1.24.19202`, a per-user copy long gone, which shadows a perfectly healthy machine-wide `1.26.21803`. `-RepairOutlookAddIn` (Ninja variable `repairOutlookAddIn`) now clears that stale `Classes\CLSID\{19A6E644-...}` key and puts `LoadBehavior` back to 3, so COM resolves to the machine-wide registration again |
| It only acts when that machine-wide registration is healthy — clearing the shadow with nothing behind it would leave the user worse off — and it is off by default, because it writes into another user's hive. It counts as work, so `-CheckOnly` reports it and `-Quiet` surfaces it |
| Tested against planted keys in both registry views: `-WhatIf` plans both actions, an applied run clears the CLSID keys, sets `LoadBehavior` to 3 and exits `0`. Untested: whether Outlook then actually loads the add-in for that user — that is the next production run |

### 2026-09-20 (4)
| Change |
|--------|
| "I do not see it loaded on all profiles yet" was a visibility gap, not only a Teams one: a profile whose hive is not mounted cannot be read at all, and the script simply left it out — so an unreadable profile and a healthy one looked identical in the output. It now lists those profiles by name, with what it means for them: with a healthy machine-wide registration they pick the add-in up at the first Outlook start, without one there is nothing to fall back on |
| Worth stating plainly, because it decides whether there is anything to fix: a profile that is not signed in is not broken. The machine-wide registration covers users who have no per-user state; only a user who already has their own (disabled, or pointing at a removed DLL) keeps shadowing it |
| Both messages verified against stubbed profile lists. Not verified here: mounting an unmounted hive to inspect or repair a signed-out profile — `reg load` needs privileges this workstation does not have, so that machinery is deliberately not built on an untested assumption |

### 2026-09-20 (3)
| Change |
|--------|
| Answered a question the script could not: **why** an account shows `LoadBehavior 2`. Outlook resolves the add-in through `Classes\CLSID\{19A6E644-...}\InprocServer32`, and a per-user registration in `HKCU\SOFTWARE\Classes` outranks the machine-wide one — so a user keeps loading the copy from their own profile even after an `ALLUSERS=1` install lands in `Program Files (x86)`. Measured on a device: the class resolves to `%LOCALAPPDATA%\Microsoft\TeamsMeetingAdd-in\<version>\x64\Microsoft.Teams.AddinLoader.dll` |
| The script now resolves that path per signed-in user and reports the two cases apart, because they need different fixes: the add-in switched off but its DLL present (tick the box back on) versus a registration pointing at a DLL that is gone (ticking will not stick — it has to be installed again for that user). The targeted lookup across loaded hives costs ~100 ms |
| Tested with a planted registration pointing at a missing DLL, without touching the real one |

### 2026-09-20 (2)
| Change |
|--------|
| Third production failure on the same session host, third fix: `Uninstall of Teams Machine-Wide Installer failed (exit code 1605)`. 1605 is "this action is only valid for products that are currently installed" — the entry in Programs and Features outlived the product, which is common once the new Teams bootstrapper has been over a machine |
| There is nothing to uninstall in that case, but the stale entry would keep the script reporting classic Teams on every run, so it now removes the registry entry instead and carries on. Tested live: a real `msiexec /x` against an unknown product code returns 1605, the run warns, removes a planted stale entry, verifies clean and exits `0` |
| The uninstall-entry objects now carry their `RegistryPath` and `UninstallString`, which is what makes that cleanup possible |

### 2026-09-20
| Change |
|--------|
| Fixed the second production failure on a session host: `WebRTC Redirector install failed (exit code 1638)`. That MSI keeps one ProductCode across versions, so `msiexec /i` over an existing install refuses with "another version of this product is already installed" rather than upgrading — and `-Force` walks straight into it on any host that already has the redirector |
| The script now reads the downloaded ProductVersion and decides: same version → repair in place (`REINSTALL=ALL REINSTALLMODE=vomus`), different version → uninstall the old one first, then install. A `1638` that still slips through is reported as "leaving the existing one in place" instead of failing the whole run |
| Measured while fixing it: `aka.ms/msrdcwebrtcsvc/msi` now serves `1.56.2603.20001`, while that host had `1.54.2408.19001` installed — so this was an upgrade being refused, not a duplicate install. Both paths are planned correctly under `-WhatIf`; neither msiexec call has been run for real yet, which the docs say out loud |
| Also written down explicitly: an installed redirector is **not** silently upgraded by a normal run. Only `-Force` replaces it. With WebRTC losing support on 1 October 2026, keeping that deliberate beats auto-upgrading a component on its way out |

### 2026-09-18 (4)
| Change |
|--------|
| [`Get-DistributionGroupMembers.ps1`](scripts/Exchange/Get-DistributionGroupMembers.ps1) — `-Member "*.verizon.com"` now matches a domain **and every subdomain of it** (`.verizon.com` and `*@*.verizon.com` are the same thing). Without the leading `*.` the filter stays on that one domain, so `@be.verizon.com` still deliberately does not reach `@us.verizon.com` |
| The run says which of the two it is doing — *"scanning N list(s) for members on verizon.com and its subdomains"* — because a filter whose scope you have to infer is a filter you cannot trust in a customer report |
| Matching is on the full domain label, verified against `@notverizon.com` and the suffix trick `@verizon.com.evil.test`; neither matches a `*.verizon.com` run. A wildcard anywhere but the front is escaped rather than quietly widening the filter |

### 2026-09-18 (3)
| Change |
|--------|
| [`Get-DistributionGroupMembers.ps1`](scripts/Exchange/Get-DistributionGroupMembers.ps1) — **`-Recurse`**, after checking whether the report really covered everyone: it did not. Exchange only ever returns *direct* members, so a list containing another list reported that list as one member and never the people inside it. Someone who receives mail only through a nested group was invisible, and `-Member` reported "no hits" on a list that does deliver to them — a wrong answer that looks like a confident one |
| `Via groep` names the group a person came in through (empty for a direct member), and someone reachable by several routes gets one row with the routes joined rather than a row per route |
| `Aantal leden` keeps counting direct members, because that is the number Exchange and the EAC show; the new `Aantal personen` counts the real recipients reached |
| A group already expanded is not expanded again, which is also what stops a membership cycle (A contains B, B contains A) from recursing forever. Verified against a deliberately cyclic pair of test lists; nesting past 20 levels is reported and left alone |
| Documented what the report still does *not* cover: it reads group membership, so a user on no list at all appears nowhere |

### 2026-09-18 (2)
| Change |
|--------|
| [`Get-DistributionGroupMembers.ps1`](scripts/Exchange/Get-DistributionGroupMembers.ps1) — `-Member` now also takes a **domain**: `-Member "@be.verizon.com"` reports every list that still holds an address on that domain (`be.verizon.com` and `*@be.verizon.com` mean the same). An address is matched by Exchange itself; a domain cannot be, so every list is read and then filtered — slower, and documented as such |
| Matching covers the primary address, every alias, **and `ExternalEmailAddress`**. That is the whole point for a partner domain: such a member is usually a mail contact whose primary SMTP is `...@contoso.onmicrosoft.com`, with the real `@be.verizon.com` only in its external address. Matching on the primary address would have found nothing and reported "none" with a straight face |
| New `Extern adres` column in the `Leden` sheet, so the address that actually receives the mail is visible for contacts instead of only the internal placeholder |
| With a filter active: `Treffers` per list in `Overzicht`, and `Treffer op` per member in `Leden`. `Treffer op` holds the matching **address**, not Ja/Nee — a hit on an alias is otherwise unexplainable in a report that does not show aliases |
| A domain filter that matches nothing says so and writes no file, rather than handing over an empty workbook that reads as a failed export |

### 2026-09-18
| Change |
|--------|
| Added [`Get-DistributionGroupMembers.ps1`](scripts/Exchange/Get-DistributionGroupMembers.ps1) — every distribution list with its members in one Excel workbook: an `Overzicht` sheet (one row per list) and a `Leden` sheet (one row per member), both filterable tables with a frozen header row. Sheet headers and recipient types are in Dutch, because the workbook is what the customer reads |
| `-Member user@domain` answers "which lists is this person on?" server-side via `Get-Recipient -Filter "Members -eq '<DN>'"` instead of walking every group, and still exports the matched lists in full so the customer sees who else is on them |
| `-IncludeDynamic` and `-IncludeM365Groups` widen the report beyond plain distribution groups; dynamic groups are evaluated live, since they store no membership to query |
| Falls back to two CSV files when `ImportExcel` is missing (and offers to install it first), so a missing module never costs you the report. `ImportExcel` added to [`Install-Modules.ps1`](scripts/Startup/Install-Modules.ps1) — `vias_archiver.ps1` already needed it |
| Exchange submenu: `K` Get-DLMembers |

### 2026-09-17 (3)
| Change |
|--------|
| Fixed the bug a production run on an AVD session host surfaced: `teamsbootstrapper.exe -p` **provisions** the package for future sign-ins, it does not install it for whoever ran the script. The bootstrapper reported success and the add-in step then died on `New Teams package not found after install`, because `Get-AppxPackage -Name MSTeams` asks about the current user and the admin running the script had no Teams |
| The add-in MSI is now found by globbing `%ProgramFiles%\WindowsApps\MSTeams_*_x64__8wekyb3d8bbwe\MicrosoftTeamsMeetingAddinInstaller.msi` and taking the newest version, so it works whether or not any user has the package installed. Re-tested for a per-user install, a provisioned-only host and a `-Force` run |
| Same per-user blind spot in the SlimCore check, which reported "not found" on a host that has new Teams for one profile: it now asks `-AllUsers` first. And the two machine-wide Outlook registry views are labelled 64-bit/32-bit apart, because that run printed two identical `all users (machine-wide)` lines, which reads like a bug |
| That run also earned the new checks their keep: it removed a real per-profile classic Teams `1.4.00.11161`, and found `LoadBehavior 2` for one account - Outlook had switched the add-in off, which no amount of reinstalling fixes |

### 2026-09-17 (2)
| Change |
|--------|
| [`Update-TeamsClient.ps1`](scripts/Device/Update-TeamsClient.ps1) can now remove classic Teams as well, behind `-RemoveClassicTeams` (Ninja variable `removeClassicTeams`). Off by default: taking an application away from users is not a decision an update job should make on its own. It uninstalls the *Teams Machine-Wide Installer* through msiexec — the one that matters, because while it is present Windows keeps staging classic Teams into every new profile — and per profile clears the install root, the `Run\com.squirrel.Teams.Teams` autostart entry and the stale `Uninstall\Teams` key |
| The documented per-user uninstall (`Update.exe --uninstall -s`) has to run as the profile owner, which System cannot do, so the files are removed instead. Roaming data in `%APPDATA%\Microsoft\Teams` is left alone |
| Failure handling splits the two cases on purpose: a machine-wide installer that survives the uninstall is a real failure (exit 1), while a per-profile folder that survives is almost always a file lock from a running classic Teams — a warning, cleared by the next run after the user signs out |
| Tested with detection faked, since the test device has neither variant: `-WhatIf` plans the msiexec uninstall and the folder removal and skips steps 5-8, and an applied run removed a faked profile folder with the verification reporting it clean. The msiexec path itself has **not** been run against a real Machine-Wide Installer — written down in the docs rather than implied |

### 2026-09-17
| Change |
|--------|
| Fixed a bug a real session-host run exposed: `Get-AppxPackage -AllUsers` finds nothing when Teams is only *provisioned* and no user has it yet, so the version check had nothing to compare, declared the host outdated and reinstalled ~275 MB on every scheduled run. The installed version now falls back to the provisioned package version. Verified against that exact scenario: reports `Provisioned MSTeams <version>`, compares, does nothing |
| [`Update-TeamsClient.ps1`](scripts/Device/Update-TeamsClient.ps1) now checks whether **Outlook itself** sees the meeting add-in, not only that the MSI installed. It reads `HKEY_USERS\<sid>\...\Outlook\Addins\TeamsAddin.FastConnect` per signed-in user plus the machine-wide key: `LoadBehavior 3` = loaded, `2`/`0` = Outlook switched it off (the real "the button is gone" case). Reported in preflight and verification, never as a failure — a profile nobody is signed into cannot be read |
| Preflight became a full inventory of everywhere Teams can live: AppX per user, the provisioned package, the classic *Teams Machine-Wide Installer*, classic per-profile installs, the add-in in both hives, the Outlook registration. Classic Teams is reported, not removed — it shares the October 2026 end-of-support date and a leftover machine-wide installer keeps restaging it into new profiles |
| Two bugs found by running it rather than reading it. Enumerating profiles with an `S-1-5-21-*` whitelist skips **every** user on an Entra-joined device, where SIDs are `S-1-12-1-*` — the check claimed nobody had the add-in registered while `LoadBehavior=3` sat right there. And `New-PSDrive` honours `ShouldProcess`, so under `-WhatIf` the `HKEY_USERS` drive was never created and the same read-only check lied; the hives are addressed through `Registry::HKEY_USERS` now, with no drive to create |

### 2026-09-11 (5)
| Change |
|--------|
| Added [`scripts/Exchange/Move-SharedCalendar.ps1`](scripts/Exchange/Move-SharedCalendar.ps1) — all in one: `-Search balie` finds the calendar, shows who uses it, moves it into a resource mailbox with [`Convert-SharedCalendarToResource.ps1`](scripts/Exchange/Convert-SharedCalendarToResource.ps1) and lists who has to switch. One temporary App Registration with the permissions of both scripts, created once and removed at the end, so there is one sign-in instead of two. Several matches are picked from a list or narrowed with `-Owner`; a non-interactive run lists them and stops rather than guessing. The two scripts are called, not copied, so there is one implementation of each step |
| Fixed [`Convert-SharedCalendarToResource.ps1`](scripts/Exchange/Convert-SharedCalendarToResource.ps1): in a non-interactive session the typed confirmation was skipped and the original calendar removed without `-Force`. It is now left in place with a warning unless `-Force` is given |
| Both calendar scripts: an explicit `-ClientId` now takes precedence over an existing app-only Graph session, so a calling script's app is really used. [`Convert-SharedCalendarToResource.ps1`](scripts/Exchange/Convert-SharedCalendarToResource.ps1) gains `-PassThru` (result object for callers) |
| Exchange submenu (`C`) option `J` added |

### 2026-09-11 (4)
| Change |
|--------|
| [`Get-CalendarMappings.ps1`](scripts/Exchange/Get-CalendarMappings.ps1) — first real run (a tenant where "Balie planning" turned out to be a secondary calendar in an archived mailbox) confirmed the Graph assumptions: app-only reads return the calendars users added, and a shared secondary calendar appears in their list under its own name. It also exposed a wrong hint: a `NotMapped` row pointed at "Reservering vergaderzaal LBM" as "probably this one", while that is a *second* calendar the same owner shares. The hint now skips entries named after another calendar of the owner, and - for a secondary calendar - entries named after the owner, which are their main calendar |

### 2026-09-11 (3)
| Change |
|--------|
| Added [`scripts/Exchange/Convert-SharedCalendarToResource.ps1`](scripts/Exchange/Convert-SharedCalendarToResource.ps1) — moves a shared calendar (the "Balie" calendar in one person's mailbox) into a Room or Equipment mailbox of its own, with every item and every permission, then removes the original on request. Preview by default; `-Apply` creates and copies, `-RemoveSourceCalendar` removes the original only after every item has a verified copy and the calendar's name has been typed as confirmation |
| Items are copied faithfully rather than approximately: recurring series stay series with their moved and cancelled occurrences applied (matched occurrence by occurrence, and left alone with a warning if the two series do not line up), times are written back in the time zone they were created in so weekly items survive a daylight saving switch, categories keep their colour, attachments up to 3 MB are copied and larger ones saved to the backup folder. Attendees are listed in the body instead of copied, so nobody receives a fresh invitation |
| Permissions carry their exact Exchange access rights, custom rights included; `-SendSharingInvitation` sends users the standard invitation. External people, deleted accounts and delegate flags are reported, not silently dropped |
| Every copy carries its source item's id in a hidden property, so a run that stops halfway continues where it left off; a half-finished series is redone. A JSON backup of everything read is written before anything is created |
| Exchange submenu (`C`) option `I` added: always a preview first, then an explicit second step |

### 2026-09-11 (2)
| Change |
|--------|
| [`Get-CalendarMappings.ps1`](scripts/Exchange/Get-CalendarMappings.ps1) gains `-Search` (alias `-Keyword`): "where is the Balie calendar?" in one run. The keyword is matched against the owner's name and every address (a shared mailbox `balie@`, a room, a group) and against calendar names (a secondary calendar *Balie* in somebody's mailbox). The report shows where the calendar lives (new status `Source`), who has it in their calendar list, and who has rights on it |
| A matching **secondary** calendar now gets its own permissions read, instead of being compared against the owner's main calendar and landing on `MappedWithoutRight`. A new `Calendar` column says which of the owner's calendars a row is about |
| A calendar list entry carries no link back to the folder it came from, so a shared secondary calendar is matched by name. When a user has it under another name, the `NotMapped` row names the entry that is probably it rather than leaving a silent false negative |
| Menu option `H` asks for a keyword first; blank falls back to the full or per-mailbox report |

### 2026-09-11
| Change |
|--------|
| Added [`scripts/Exchange/Get-CalendarMappings.ps1`](scripts/Exchange/Get-CalendarMappings.ps1) — shows where each calendar is actually mapped: for every mailbox it reads the calendar list in Outlook and the rights on its own main calendar, and folds both into one row per owner + user with a status (`Mapped`, `MappedWithoutRight`, `NotMapped`, `MappedOwnerMissing`, `SharedExternally`, …). [`Test-CalendarPermissions.ps1`](scripts/Exchange/Test-CalendarPermissions.ps1) says who *may* open a calendar; this says where it *is*, and where the two disagree |
| Graph rather than Exchange Online PowerShell, because the entries a user added to their own calendar list are not visible to any Exchange cmdlet. App-only access follows the same three routes as [`Remove-PhishingMessage.ps1`](scripts/Exchange/Remove-PhishingMessage.ps1) (existing session, own app, or a temporary app that is removed in a `finally`), with read-only permissions `Calendars.Read`, `User.Read.All` and `Group.Read.All`. No Exchange connection, so no MSAL clash |
| Tenant-wide runs go through `$batch` (20 mailboxes per call) with throttled items retried. A mailbox that cannot be read is reported as such rather than as "nothing mapped" |
| Written down what the report cannot see: Full Access with AutoMapping (a mailbox permission — [`Test-MailboxPermissions.ps1`](scripts/Exchange/Test-MailboxPermissions.ps1)), calendars opened in classic Outlook without shared calendar improvements, and secondary calendars, which show up as `MappedWithoutRight` |
| Exchange submenu (`C`) option `H` added |

### 2026-09-16 (5)
| Change |
|--------|
| The explanation page is now step 4 of the one-command build rather than a separate script somebody remembers a week later. A structure nobody was told about is a structure nobody uses, and because the page is generated from the same configuration it describes exactly what the run just made |
| `-SkipHelpPage` leaves it out, `-HelpContact` says who people should ask. The installer passes `-Force`, because it owns that page: rerunning the build brings the explanation back in line with what the build made |
| Verification and the deelstatus audit shifted to steps 5 and 6, and the stale "step 2 changes permissions on a live team" warnings now name step 3 |
### 2026-09-16 (4)
| Change |
|--------|
| Added [`scripts/SharePoint/Provisioning/Add-SharePointHelpPage.ps1`](scripts/SharePoint/Provisioning/Add-SharePointHelpPage.ps1) — puts the end-user explanation on the team site as a SharePoint page, linked from the left-hand navigation. A handleiding in a repo is read by nobody; this writes it where the people who upload files already are |
| The page is generated from the configuration rather than typed out, so it cannot drift from what the libraries actually do: the channels it lists are the ones that exist, the labels carry the same help text that appears under each field in the upload form, and the required fields per document type are read off the content types |
| Written for the person uploading a catalogue. Two pieces of the configuration are deliberately kept off it: the `note` on a container, which names security groups, and the `description` on a view, which talks about pillars and synced folders. Permissions are left out entirely — who may see what is not something a user can act on |
| Fixed along the way, found by rendering the page rather than reading the code: it announced three ways of adding a file and listed two, and the wizard was writing the team name where the company name belonged ("Intern blijft binnen Laseto-NewTeams") |
### 2026-09-16 (3)
| Change |
|--------|
| Added [`scripts/SharePoint/Provisioning/Remove-SharePointStructure.ps1`](scripts/SharePoint/Provisioning/Remove-SharePointStructure.ps1) — takes the same configuration apart, deepest first: tabs, channels, libraries, content types (unbound from their lists first), site columns, term set, security groups, and the team itself |
| **Deliberately the reverse default of everything else in the folder: without `-Apply` it changes nothing.** Forgetting `-WhatIf` on a destructive script is the dangerous direction, so the safe state is the one you get for free |
| `-Scope All` never includes the team. Deleting a client's whole team is not something anyone should get by asking for "all" — it has to be named, and then the team's name typed to confirm |
| Refuses by default rather than asking forgiveness: a library or channel folder that still holds files is skipped unless `-IncludeContent` (the item count is reported either way), the General channel and the team's own Documents library are never removed, and a content type still in use is reported rather than forced |
| Reported with their cost before they run, because no recycle bin brings them back: removing the term set orphans the Leverancier value on every document that carried one, and removing a column takes its data with it. What *is* recoverable is said too — a deleted group or team is soft-deleted for 30 days, a channel has its own 30-day recycle, and files from a removed library land in the site recycle bin |
| Menu step `6` runs it; the menu asks about applying, about files, and about the team as three separate questions rather than one |
### 2026-09-16 (2)
| Change |
|--------|
| A restricted pillar can now be shaped as **its own library behind an ordinary channel**, which is the only arrangement that gives a real read-only role and still puts a channel in Teams. The wizard asks which pillars are restricted and then in which form - `bibliotheek` (the default) or `privekanaal` |
| The library form breaks inheritance **without copying it**, which is the whole point: copying carries every team member across as an editor, which is exactly the door the shape is meant to close. What survives is the site's own owners plus the pillar's two groups - Contribute and Read |
| A private channel offers no read-only role at all: owners and members, and members may post, edit and delete. So a pillar that needs "may look" cannot be a private channel, and the wizard now says so at the point where the choice is made |
| The channel is created as usual but gets no folder in the shared library, and the restricted library is surfaced as a tab in that channel - a standard channel's own Files tab always points at the team library and cannot be repointed, so it sits beside it |
| Flagged in the readme because it will otherwise be reported as a bug: that built-in Files tab stays, pointing at a folder nobody uses. Either point people at the named tab, or remove the Files tab from the channel once by hand |
### 2026-09-16
| Change |
|--------|
| Added [`scripts/SharePoint/Provisioning/Sync-SharePointChannelMember.ps1`](scripts/SharePoint/Provisioning/Sync-SharePointChannelMember.ps1) — makes an Entra ID security group the source of truth for who is in a private Teams channel. A private channel cannot be given rights through a group at all: Teams tracks its roster one person at a time and Graph accepts only individual users there, so the group feeds the roster instead |
| The obvious workaround is a trap and is documented as one: adding the group to the private channel site's SharePoint permissions works until Teams syncs the roster back over it, and in the meantime those people reach the files while the channel stays invisible to them in Teams. Unsupported by Microsoft |
| Nested groups are followed, non-users are dropped, and everyone is added to the parent team first — Teams refuses a private-channel member who is not on the team, and the error it returns does not mention that. `-Prune` also removes people the groups no longer list; channel owners are never removed |
| Written down because it changes the design, not just the script: **a private channel has no read-only role.** Owners and members, and members may post, edit and delete. A group named `-RO` cannot mean "may look" there, so the run reports per group how many people it brought in rather than letting that pass unnoticed. Where read-only genuinely matters, a document library with its own permissions is the right shape |
| `-EnsureGroups` now creates every group in the model rather than only the ones a library grants to. A private channel grants nothing, so the MGMT pair sat in the configuration and was never created — which is exactly the pillar whose groups you go looking for first |
| The wizard writes a `channelMembers` section for private pillars naming the groups that feed the roster; a configuration written before that key existed falls back to the configured groups named after the container, and reports the fallback |
| Menu step `5` runs the sync; it signs in to Graph on its own, so it asks for no PnP app registration |
### 2026-09-15 (3)
| Change |
|--------|
| `New-StructureConfig.ps1 -All` asks for the names that were still being derived behind the operator's back: per pillar the channel name, the folder, the content type and both group names and the view title; plus the library behind the channels, the column and content type groups, the term set, the team site URL and the label every column carries for the user. Each keeps its derivation as the suggestion, so `-All` is still mostly Enters |
| Only the column *internal* names stay fixed. They are never shown to anyone, and changing one after documents carry it loses the metadata on those documents |
| Fixed: an optional question could never be turned down, because Enter means "take the suggestion". Optional questions now say `(of "geen")` and accept geen/none/nee/- as a real "none" — before this, answering nothing to "customer library" still created FUTECH |
| Fixed a one-item list coming back as a bare string: PowerShell unrolls a single-element array on return, so a client with one brand crashed the wizard on `.Count`. Returned with a leading comma now |
| Fixed two `$x = if (...) { @() }` assignments that yield `$null` rather than an empty array — a configuration with no sales pillar or no suppliers died at the summary |
| All three paths verified end to end against the config validator: the full six-pillar default, a `-All` run with deliberately different names throughout, and a minimal two-pillar tenant with no private channel, no suppliers, no regions and no customer library |
### 2026-09-15 (2)
| Change |
|--------|
| [`New-StructureConfig.ps1`](scripts/SharePoint/Provisioning/New-StructureConfig.ps1) asks what everything should be called and writes the configuration itself — nobody should have to open a JSON file to name a channel. Enter accepts the suggestion in brackets, so a standard build is mostly Enters plus the tenant and the team owner |
| Everything else is derived from those answers: per pillar a channel, a content type, two security groups and a grouped view; per brand a cross-cutting view spanning every pillar folder. Which pillar handles suppliers and which handles sales is what decides where Leverancier and Regio become required fields |
| Two of the answers are the ones that cost something later, so they are asked last and default to no: maintaining the share-status column (the only nightly script) and enforcing per-pillar rights on standard-channel folders (the part Microsoft does not support) |
| [`New-SharePointTeam.ps1`](scripts/SharePoint/Provisioning/New-SharePointTeam.ps1) creates the Microsoft 365 team and its channels, the private MGMT one included, so the structure can be built from an empty tenant. A private channel's site collection is provisioned asynchronously and its URL cannot be known in advance — the script polls for it and writes it back into the configuration, which is what lets the following steps connect to something |
| Never renames or deletes a channel: a channel whose name does not match the config is reported, not corrected, because renaming one moves its folder and breaks every link anyone has shared |
| [`Install-SharePointStructure.ps1`](scripts/SharePoint/Provisioning/Install-SharePointStructure.ps1) runs the wizard by itself when it finds no configuration for the tenant, and the team step is now step 1 of six. `-SkipTeam` for a team that already exists |
| Column internal names and content type IDs are generated once and then fixed — SharePoint keys document metadata to both — which is why the wizard refuses to overwrite an existing configuration without `-Force`. Display names, channel names and group names stay changeable |
| Removed the last hardcoded column names: the share-status audit reads which column is which from a new `fieldRoles` section instead of assuming `PsDeelstatus` and `PsVertrouwelijkheid` |
| The app registration now also consents `Channel.Create`, `ChannelSettings.ReadWrite.All` and `Team.Create`, which the team step needs |
### 2026-09-15
| Change |
|--------|
| Added [`scripts/SharePoint/Provisioning/Install-SharePointStructure.ps1`](scripts/SharePoint/Provisioning/Install-SharePointStructure.ps1) — builds the whole structure in one run: registers the Entra app itself and admin-consents its delegated scopes, then metadata → libraries/groups/permissions → verification, optionally the first deelstatus audit. Each step stays its own script, so a failure is rerun on its own instead of starting over |
| `-TemporaryApp` deletes the app registration again at the end, for a one-off build on a tenant you do not manage day to day. It only ever deletes an app **this run created** — one that was already cached predates the run and is somebody else's to remove, so the script says so rather than quietly deleting it. Without the switch the app stays and the client ID is cached in `pnp.appid.json`, shared with the other PnP scripts in this repo |
| Flagged in the docs because it will otherwise be reported as a bug: a `-WhatIf` run needs an app to sign in with. With no cached app for the tenant there is nothing to connect as, so the dry run validates the config and stops there — run once for real, or pass `-ClientId`, to dry-run step by step |
| Closed the gap that made "merk als tag" only half true: cross-cutting views on the shared library with `Scope = RecursiveAll`, so *Alles - Butterstone* is one flat list across every pillar folder, **including everything tagged Beide** — one file, two brands, no copies to drift apart. Plus *Nog te taggen* (what drag-and-drop and OneDrive sync leave behind), *Extern gedeeld* and *Te archiveren*. The filter is raw CAML in the config rather than a mini query language of the script's own invention |
| Never group a view on `PsTaal`: SharePoint refuses to group on a multi-value column. Filtering on it works fine, and no shipped view groups on it |
| Added `Petsolutions-SharePoint-Handleiding.md` — end-user documentation in Dutch to hand to the customer. Covers the three ways of adding a file and why they behave differently, what each label means, and what happens the moment you tag something (the file does not move, links keep working, `Beide` shows up in both brand views, search lags a few minutes behind the views) |
| Two things the guide says out loud because users assume the opposite: **drag-and-drop and OneDrive sync ask nothing** — required columns are enforced by the upload form, not by the library, so bulk-dropped files land with empty labels and a "Required info" prompt rather than being blocked; and **a label is not a lock** — Vertrouwelijkheid shuts nobody out, it is an agreement plus the signal the nightly audit uses to flag over-sharing |
| Menu item `S` gained step `0` for the all-in-one build; the per-step options are unchanged |

### 2026-09-10 (4)
| Change |
|--------|
| Added [`scripts/SharePoint/Provisioning/`](scripts/SharePoint/Provisioning/readme.md) — provision and maintain a whole SharePoint structure (metadata model, content types, libraries, Entra ID group permissions) for an MSP client from one JSON config, with a sharing audit and a read-only drift check. Built for Petsolutions NV (brands Butterstone/Laseto), but nothing in the scripts is client-specific |
| The model lives in `petsolutions.config.json`, cross-checked at load time: a content type referring to an undefined column, or a container granting a group that is not in the model, fails before anything connects rather than halfway through provisioning. The shipped `CHANGEME` tenant/site URLs are refused outright |
| Column internal names carry a `Ps` prefix. "Contenttype" and "Status" are display names SharePoint already uses for something else, and the prefix keeps them unambiguous in CAML, in views and in the drift check while users still see plain Dutch labels. Content type IDs are fixed rather than generated, so the same structure is reproducible across tenants |
| [`New-SharePointMetadata.ps1`](scripts/SharePoint/Provisioning/New-SharePointMetadata.ps1) runs against **every** site in the config, not just the team site: a Teams private channel (MGMT here) is its own site collection and a site column does not reach across one. Making a column required after the fact works — the `Required` flag on an existing field link is updated in place and pushed down to the lists already using the content type |
| [`Set-SharePointLibraries.ps1`](scripts/SharePoint/Provisioning/Set-SharePointLibraries.ps1) sets a per-folder content type order, so the *New* menu inside the Leveranciers channel offers Leveranciersdocument and not the five types belonging to the other pillars — the shared library has to carry them all, the folder does not have to show them. No view is ever made the default: the default view of a Teams library is what every member of the channel sees the second they open Files |
| Written down rather than hidden: unique permissions on a **standard**-channel folder are what this model asks for and what Microsoft does not support. Members who lose access keep seeing the channel in Teams and get an error on the Files tab instead of a closed door. The script does it, warns per folder, and `-SkipChannelFolderPermissions` leaves those folders inheriting. A private channel, a shared channel or an own library (what FUTECH uses) are the supported ways to close a pillar off |
| [`Update-SharePointShareStatus.ps1`](scripts/SharePoint/Provisioning/Update-SharePointShareStatus.ps1) derives the Deelstatus column from the permissions actually on each file. It asks the cheap question first — a file that inherits is not shared — so one round trip per hundred items settles nearly the whole library; only files that broke inheritance get their role assignments read, and of those only specific-people links need expanding (an Anyone or Organization link already says in its name whether a guest can be behind it). Writes with `SystemUpdate` so Modified/Modified By stay put and no version is created |
| The audit never revokes a link. It reports files tagged Intern or Vertrouwelijk sitting behind an external one and exits `2`, so a scheduled RMM job surfaces exactly when there is a decision for a person to make. [`Test-SharePointStructure.ps1`](scripts/SharePoint/Provisioning/Test-SharePointStructure.ps1) does the same for structural drift, classified as Missing / Different / Extra — "Extra" is never fixed automatically, because an extra column holds data and an extra role assignment is usually somebody's deliberate exception |
| [`SharePointStructure.Common.ps1`](scripts/SharePoint/Provisioning/SharePointStructure.Common.ps1) is dot-sourced by all four — a deliberate exception to the "every script stands alone" rule elsewhere in this repo, because they share one config schema and three copies of the permission code would drift apart within a month |
| Menu item `S` added for the set (pick a step, `-WhatIf` unless you confirm; the drift check skips the question because it never writes) |
### 2026-09-10 (4)
| Change |
|--------|
| Attached the expiry date to the `-AvdOptimizations` feature: Microsoft retires the WebRTC-based AVD media optimization on **1 October 2026** (end of support) and **1 April 2027** (end of availability), and Teams already shows users a banner about it. The switch keeps installing the redirector because Microsoft still advises it as a fallback — with a note to revisit before April 2027 |
| Its replacement, SlimCore, needs nothing on the session host: it ships inside new Teams. Confirmed on a device with Teams `26225.1806.5074.1452`, which carries `Microsoft.Teams.SlimCoreVdiHost.win-x64` `2026.31.1.16` plus several framework packages. Preflight now reports that package under `-AvdOptimizations` |
| That report is deliberately informational and creates no work item: which media path is used depends on the Windows App version on the endpoint the user connects from, which a script running on the session host cannot see. Auditing endpoint client versions is the actual migration work |
| Added a service desk section to the IT Glue doc for the banner users are reporting: what it means (an announcement, not an outage), the two dates, that the fix is on the local device rather than the session host, how to read the `AVD SlimCore Media Optimized` / `AVD Media Optimized` line under Teams > About, and ready-made text for the user |

### 2026-09-10 (3)
| Change |
|--------|
| Corrected a wrong claim in the Teams docs and in the script comment: the meeting add-in's uninstall entry does **not** always live in `WOW6432Node`. Measured on a Windows 11 endpoint, add-in `1.26.21803` registers in the **64-bit** hive, with `InstallSource` pointing at a per-user MSI cache. Scanning both hives (which the script already did) is right — the stated reason was not |
| Documented how the add-in actually reaches a device, measured rather than assumed: the script installs it machine-wide (`ALLUSERS=1`, `Program Files (x86)`) for shared machines and session hosts, while on an ordinary endpoint the Teams client installs and updates it **per user** from `%LOCALAPPDATA%\Microsoft\TeamsMeetingAddinMsis` into `%LOCALAPPDATA%\Microsoft\TeamsMeetingAdd-in`, registering only in `HKCU\...\Office\Outlook\Addins` |
| Written down with it: what step 8 actually proves. It reads the HKLM uninstall keys, so it confirms the machine-wide install succeeded — not that a given user's Outlook shows the button. Run as System the script cannot see a user's `HKCU` at all |

### 2026-09-10 (2)
| Change |
|--------|
| [`scripts/Device/Update-TeamsClient.ps1`](scripts/Device/Update-TeamsClient.ps1) absorbs the AVD/VDI parts of the older gap-fill installer behind `-AvdOptimizations`: the `IsWVDEnvironment` media flag (set in step 3, before the client is provisioned, because Teams reads it at startup to pick its media path) and the Remote Desktop WebRTC Redirector Service from `aka.ms/msrdcwebrtcsvc/msi` |
| Deliberately a switch and not autodetection: setting that flag on a normal endpoint tells Teams to hand media to a redirector that is not there. Without the switch the script only *reports* that a device looks like a session host (`HKLM:\SOFTWARE\Microsoft\RDInfraAgent`) |
| Both components are installed only when missing (`-Force` reinstalls the redirector), so a scheduled run on a configured session host still downloads nothing and prints nothing under `-Quiet`. Verified end to end, including that the redirector MSI (1.7 MB, `1.54.2408.19001`) passes the Microsoft signature check |
| The add-in step now also skips itself when the add-in is present and the client was not replaced — before this it would reinstall the add-in on a run that was only there to fix the AVD components |
| Download and signature verification moved into one `Save-VerifiedDownload` helper shared by the bootstrapper and the redirector: https-only, minimum size, Authenticode `Valid` and signed by `O=Microsoft Corporation`, or it throws |
| Corrected in the docs: Ninja script-variable names are **not** case-sensitive. Windows environment lookups are case-insensitive, so variables named `Quiet` or `Force` work exactly like `quiet` and `force` |

### 2026-09-10
| Change |
|--------|
| [`scripts/Device/Update-TeamsClient-ITGlue.md`](scripts/Device/Update-TeamsClient-ITGlue.md) gained a NinjaOne setup appendix: which values to pick per field when adding the script (PowerShell 5.1 rather than 7, 64-bit, Run As System), the script-variable names with a way to verify they actually arrive, the test run on one device, the scheduled automation with `-Quiet -Confirm:$false`, and the optional detection job |
| Written down explicitly because it will otherwise be reported as a bug: `-CheckOnly` exits `2` when an update is available, and NinjaOne shows every non-zero exit code as a failed job. That is the intent — those are the devices needing attention — and it is what a script result condition can key on |
| Also flagged: the Ninja script timeout must exceed `-TimeoutSeconds` (900 s) plus the ~275 MB download, otherwise Ninja kills the job mid-install |

### 2026-09-08 (5)
| Change |
|--------|
| Added [`scripts/Device/Update-TeamsClient-ITGlue.md`](scripts/Device/Update-TeamsClient-ITGlue.md) — the service desk version of that documentation, in Dutch, to paste into IT Glue. Layered per support level: L1 checks with `-CheckOnly -Quiet` and reads the labelled output, L2 runs the update from NinjaOne or by hand and verifies afterwards, L3 gets parameters, exit codes, paths and the built-in safeties. Includes an error table with the escalation level per message, an FAQ, and ready-made text for the end user |
| It leads with the point that trips people up: no output means the device is already current, which is a successful run and not a failure. Download volume per device (~275 MB: a 1.9 MB bootstrapper that pulls a ~273 MB package) was measured, not estimated |

### 2026-09-08 (4)
| Change |
|--------|
| Added [`scripts/Device/Update-TeamsClient.md`](scripts/Device/Update-TeamsClient.md) — a reference for that script: the decision tree (behind → full reinstall, current but add-in missing → add-in only, current → nothing at all), the seven steps, the config-service version check with a sample response, output modes and exit codes, the NinjaOne script-variable table, the design decisions behind the order of operations, a troubleshooting table and what has actually been tested. Linked from the Device readme |

### 2026-09-08 (3)
| Change |
|--------|
| [`scripts/Device/Update-TeamsClient.ps1`](scripts/Device/Update-TeamsClient.ps1) no longer reinstalls unconditionally: it asks the Teams client config service (`config.teams.microsoft.com/config/v1/MicrosoftTeams/...`, `BuildSettings.WebView2PreAuth.<arch>.latestVersion` — the same feed the client uses to decide it is out of date) which build is published for this architecture, and leaves an up-to-date device completely alone |
| Is the client current but the meeting add-in missing? Then only the add-in is installed — no download, no uninstall, no reprovision |
| `-Quiet` holds back all output until there is news, so a scheduled NinjaOne run prints nothing on an up-to-date device and only surfaces in the activity feed when it found a newer build or hit a problem. Verified: an up-to-date `-Quiet` run produces zero bytes of output and exit code 0 |
| `-CheckOnly` reports without changing anything and exits `2` when a newer build is available, for use as a Ninja detection/condition job. `-Ring` selects a non-default update ring |
| When the config service cannot be reached the run stops instead of reinstalling blindly; `-Force` now means "reinstall even though it is current" as well as "continue without Teams or version info" |
| A transcript is only written when the run actually changes something, so an hourly check leaves no log litter in `C:\Temp` |

### 2026-09-08 (2)
| Change |
|--------|
| [`scripts/Device/Update-TeamsClient.ps1`](scripts/Device/Update-TeamsClient.ps1) is now safe to run unattended from an RMM (NinjaOne) *and* by hand. It relaunches itself 64-bit via `SysNative` when the agent starts PowerShell 32-bit — otherwise the HKLM reads are redirected to `WOW6432Node` and `$env:ProgramFiles` points at the x86 folder, so neither the AppX package nor the add-in MSI is ever found |
| NinjaOne script variables (`whatIf`, `force`, `skipMeetingAddIn`, `skipSignatureCheck`, `workingDir`, `logPath`) are read from the environment when the matching parameter is not passed, so a preview run can be a checkbox instead of a parameter string |
| Started by hand without elevation it now asks for UAC and continues in an elevated window, instead of failing on a `#Requires -RunAsAdministrator` line, and an interactive apply run asks for confirmation once. `-Confirm:$false` makes it unattended; the menu passes that because it already asked |
| Reordered so the bootstrapper is downloaded **and** its Microsoft Authenticode signature verified before the first uninstall — a failed download or a blocked URL can no longer leave a device without a Teams client. TLS 1.2 is forced for the download, and a non-https `-BootstrapperUrl` is refused |
| `msiexec` and the bootstrapper now run through one helper with a timeout (`-TimeoutSeconds`, default 900, process killed on expiry), a retry on 1618 (another install in progress) and 3010 handled as success with a pending-reboot note, so an RMM job can never hang the agent |
| The AppX package is also deprovisioned (`Remove-AppxProvisionedPackage`), otherwise new user profiles keep getting the old version staged from the image |
| Add-in MSI version now comes from the MSI property table via the `WindowsInstaller.Installer` COM object. `Get-AppLockerFileInformation` — what Microsoft's own sample uses — is missing on some editions and under PowerShell 7 it drags in the Windows PowerShell compatibility layer, which fails and floods a `-WhatIf` run with unrelated file-copy output |
| Apply runs write a transcript to `C:\Temp\Update-TeamsClient_<timestamp>.log`; unexpected errors abort instead of continuing half-way; exit code is 0 on success (`-WhatIf` included) and 1 on failure |

### 2026-09-08
| Change |
|--------|
| Added [`scripts/Device/Update-TeamsClient.ps1`](scripts/Device/Update-TeamsClient.ps1) — clean reinstall of new Teams on an endpoint or AVD session host: uninstall the Teams Meeting Add-in, remove the `MSTeams` AppX package for all users, download `teamsbootstrapper.exe`, provision Teams (`-p`) and install the meeting add-in MSI that ships inside the new Teams package |
| Every state-changing step runs through `ShouldProcess`, so `-WhatIf` walks the whole flow and prints each uninstall/download/install without touching the machine; the steps that only exist after a real install (new Teams version, add-in MSI path, final verification) are reported as such instead of failing the run |
| The add-in lookup reads both the 64-bit and the `WOW6432Node` uninstall hive — the add-in installs 32-bit, so the 64-bit hive alone never finds it (uninstall and verification both missed it before) |
| Exit codes and msiexec/bootstrapper exit codes are checked instead of assumed; `-SkipMeetingAddIn` replaces only the client, `-Force` installs on a device without any Teams. Wired into [`menu.ps1`](menu.ps1) (key T), which defaults to a `-WhatIf` preview |

### 2026-09-07 (2)
| Change |
|--------|
| Added [`scripts/SharePoint/Search-SharePointContent.ps1`](scripts/SharePoint/Search-SharePointContent.ps1) — the Microsoft Graph counterpart of [`Find-SiteContent.ps1`](scripts/SharePoint/Find-SiteContent.ps1): app-only, no interactive sign-in, and it searches one site or **every site and OneDrive in the tenant** |
| Finding content uses `/drives/{id}/root/delta` (a whole library tree in pages of a thousand items, so `*contains*` wildcards work) or `/search/query` with `-Content` for text inside documents; the filters are the same as in the PnP script |
| Permissions come from `/drives/{id}/items/{id}/permissions`, 20 per `/$batch` call. One call yields the roles, the granted-to identities, the sharing link with its scope (anyone/organization/specific people), edit-or-view, expiry and URL, and `inheritedFrom` — which is what decides `PermissionSource = Item` (unique) versus `Inherited`. "Anyone with the link" gets its own counter because those need no sign-in at all |
| Written down explicitly, in the script and the readme: Graph has no API for SharePoint role assignments, so site- and list-level rights and items in ordinary (non-library) lists stay the domain of [`Find-SiteContent.ps1`](scripts/SharePoint/Find-SiteContent.ps1). The readme has a comparison table for picking between the two |
| Sign-in is app-only: the first run registers an app, consents the application role `Sites.Read.All`, creates a self-signed certificate in `CurrentUser\My` and uploads its public key — no secret on disk — and caches client ID plus thumbprint per tenant in `graph.appid.json` (added to [`.gitignore`](.gitignore)). Later runs connect without a prompt, so it also works from a scheduled task. Throttling (429) is retried, honouring `Retry-After`, for single calls and batch sub-requests alike |

### 2026-09-07
| Change |
|--------|
| Added [`scripts/SharePoint/Find-SiteContent.ps1`](scripts/SharePoint/Find-SiteContent.ps1) — search an entire SharePoint site or OneDrive for content and report which permissions apply to every hit. Read-only |
| Two engines: a crawl over every list and library (sees everything, `-IncludeSubsites` for the subsites) and a KQL query against the search index (`-Content`) that also matches text *inside* documents. Both share the filters `-Name`, `-Path`, `-Extension`, `-ItemType`, `-ListName`, `-ModifiedBy`, `-ModifiedAfter`/`-ModifiedBefore` and `-MinSizeMB` |
| Per hit the script resolves where the permissions come from — the item itself (broken inheritance), its list, or the site — and flattens the role assignments to one CSV row per principal with type, login, e-mail and role names. `Limited Access` is filtered out unless `-IncludeLimitedAccess` |
| Sharing links (the `SharingLinks.*` groups behind "Copy link") are always expanded to the people in them and labelled Anyone/Organization/Specific people; external guests (`#ext#`) and "Everyone (except external users)" are flagged separately in the summary and the CSV |
| Site and list permissions are read once and cached and item permissions only for items that broke inheritance, so cost scales with the number of hits, not the size of the site; `-Permissions Unique` reports only what is shared differently, `-Permissions None` skips permissions, and `-MaxPermissionLookups` caps a too-broad search |
| Reuses the app-registration flow and the per-tenant `pnp.appid.json` cache of [`Restore-RecycleBinItems.ps1`](scripts/SharePoint/Restore-RecycleBinItems.ps1), and can temporarily grant itself site collection admin (`-GrantSiteAdmin`) to search a site or OneDrive it has no rights on |

### 2026-08-28
| Change |
|--------|
| Added [`scripts/SharePoint/`](scripts/SharePoint/readme.md) with [`Restore-RecycleBinItems.ps1`](scripts/SharePoint/Restore-RecycleBinItems.ps1) — restore deleted files/folders from a SharePoint site or OneDrive recycle bin, dry-run by default, with filters on name, original folder, who deleted it, and a deletion time window |
| Two scopes: `-SiteUrl` for one site collection (OneDrive included), or `-AllSites -TenantUrl` to walk every SharePoint site in the tenant. The tenant sweep excludes OneDrive personal sites, the My Site host, redirect sites and locked sites, supports `-SiteFilter`/`-MaxSites`, and keeps going when a single site errors out — per-site results land in a summary table and in a `Site` column in the CSV |
| Restores run in batches of up to 200 items through `Restore-PnPRecycleBinItem -IdList` (one server call per batch) instead of one call per item; folders and files never share a batch, and a batch that fails as a whole is retried item by item so per-item errors are still reported. Parallel runspaces were deliberately not used — PnP PowerShell is not thread-safe and concurrent calls against one site collection hit SharePoint throttling |
| The script reports timing at every step — how long reading the recycle bin took, an up-front estimate of the restore, a progress bar with a live ETA from the measured rate, the real duration in the summary, and a `DurationSeconds` column per item in the CSV |
| The script registers its own Entra app on the first run against a tenant (public client, delegated `AllSites.FullControl`, admin-consented) because PnP PowerShell no longer ships a shared multi-tenant app; the client ID is cached per tenant in `pnp.appid.json` (added to [`.gitignore`](.gitignore)) |

### 2026-07-24 (3)
| Change |
|--------|
| Retired a now-unmaintained internal PowerShell repo (`Windows-Powershell`, last commit March 2023) by reviewing every script in it and modernizing whatever still had value into this repo — nothing was copied verbatim; everything was rewritten against Microsoft Graph / Exchange Online (the source repo's `MSOnline`/`AzureAD`-based scripts are fully non-functional since Microsoft retired those endpoints) |
| Added [`scripts/TenantOnboarding/`](scripts/TenantOnboarding/readme.md) (24 scripts across Provisioning/MultiTenant/AppDeployment/DeviceConfig/OneDriveManagement/UserManagement) — modernized from the source repo's tenant-setup/onboarding scripts |
| Added [`scripts/Office365Toolkit/`](scripts/Office365Toolkit/readme.md) (9 scripts across Security/Exchange/Intune) — modernized from a forked copy of the retired `directorcia/Office365` (CIAOPS) GitHub project found in the source repo; reviewed capability-by-capability and consolidated, not ported 1:1 |
| Added [`scripts/PatronToolkit/`](scripts/PatronToolkit/readme.md) (13 scripts across Entra/Security/Exchange/Intune/SharePoint/Teams) — modernized from a forked copy of the retired `directorcia/patron` GitHub project found in the source repo, same capability-consolidation approach |
| Added [`scripts/LegacyUtilities/`](scripts/LegacyUtilities/readme.md) (17 scripts across Exchange/Entra/Teams/Network/Device/Workspace365) — modernized from assorted small tools in the source repo not covered by the above |
| Data-handling boundary applied throughout: the source repo's `Klanten`, `created-users`, `csv files`, and `Archief` folders (real customer names/tenant domains/generated passwords) were never read or ported; any other script found to hardcode real customer/tenant identifiers or secrets was generalized into parameters instead, or skipped outright — see each new folder's readme for its specific skip list |
| None of the ~63 new scripts are wired into [`menu.ps1`](menu.ps1) — they're audit/reporting/setup scripts meant to be run directly, matching the existing pattern for [`scripts/RDS/`](scripts/RDS/readme.md), [`scripts/Azure/`](scripts/Azure/readme.md), and [`scripts/Network/UniFi/`](scripts/Network/UniFi/readme.md) |

### 2026-07-24 (2)
| Change |
|--------|
| Fixed [`scripts/Reporting/Get-SharePointStorageReport.ps1`](scripts/Reporting/Get-SharePointStorageReport.ps1) silently abandoning version-history lookups on very large libraries: `Invoke-GraphBatchGet`'s retry-pass ceiling was hardcoded at 8, but SharePoint Online's per-app activity throttle allows only ~1500-2500 resolved version lookups per pass before a ~60-90s cool-down repeats — on a 200k-file tenant this meant ~90% of files got marked "gave up" before the scan actually finished |
| Applied the identical fix to [`scripts/Reporting/Remove-SharePointFileVersionsByDate.ps1`](scripts/Reporting/Remove-SharePointFileVersionsByDate.ps1)'s own copy of the same batch-retry function (`Get-FileVersionsBatch`), which had the same hardcoded 8-pass ceiling |
| Added `-MaxVersionRetryPasses` parameter to both scripts (default `0` = auto-scales the pass ceiling to the request volume, capped at 500 passes); a clear `Write-Warning` is now emitted listing exactly how many files were abandoned and suggesting the parameter if the ceiling is still hit |
| Updated [`scripts/Reporting/readme.md`](scripts/Reporting/readme.md) to document `-VersionBatchConcurrency` and `-MaxVersionRetryPasses` for both scripts |

### 2026-07-24 (1)
| Change |
|--------|
| Documented [`scripts/Azure/VM/Azure-NVMe-Conversion.ps1`](scripts/Azure/VM/Azure-NVMe-Conversion.ps1) (previously untracked in any readme/structure tree): added [`scripts/Azure/readme.md`](scripts/Azure/readme.md) and [`scripts/Azure/VM/readme.md`](scripts/Azure/VM/readme.md), and a new "Azure Infrastructure" category in this readme |
| Wired 5 previously menu-less but fully-documented scripts into [`menu.ps1`](menu.ps1): [`Move-InboxToArchive.ps1`](scripts/Exchange/Move-InboxToArchive.ps1) and [`Set-Distributionlist-dynamic-static.ps1`](scripts/Exchange/Set-Distributionlist-dynamic-static.ps1) (Exchange submenu, keys E/F), [`Get-M365UserLicenses.ps1`](scripts/Entra/Get-M365UserLicenses.ps1), [`Import-ConditionalAccessBaseline.ps1`](scripts/Entra/Import-ConditionalAccessBaseline.ps1), [`Set-UserManager.ps1`](scripts/Entra/Set-UserManager.ps1) (Entra submenu, keys G/H/I) — and added them to the Menu tables in this readme |
| Added [`scripts/Device/Remove-OemBloatware.ps1`](scripts/Device/Remove-OemBloatware.ps1) — detects HP/Lenovo/Dell and removes OEM bloatware via winget plus generic Microsoft Store junk via AppX; dry-run by default; wired into [`menu.ps1`](menu.ps1) (key I) |
| Added [`scripts/Device/DriveMapping/New-CloudDriveMapping.ps1`](scripts/Device/DriveMapping/New-CloudDriveMapping.ps1) — maps SharePoint/OneDrive document libraries to drive letters via WebDAV for use as a logon script; dry-run by default |
| Added [`scripts/Network/UniFi/`](scripts/Network/UniFi/readme.md) — [`UnifiApi.ps1`](scripts/Network/UniFi/UnifiApi.ps1) shared login helper (classic controller + UniFi OS auto-detect), [`Get-UnifiNetworkReport.ps1`](scripts/Network/UniFi/Get-UnifiNetworkReport.ps1) (HTML documentation report), [`Update-UnifiFirmware.ps1`](scripts/Network/UniFi/Update-UnifiFirmware.ps1) (dry-run firmware upgrade tooling); credentials always via `Get-Credential`, never hardcoded |
| Added [`scripts/Intune/Compare-IntuneConfig.ps1`](scripts/Intune/Compare-IntuneConfig.ps1) — Intune configuration drift detection between a customer tenant and an MSP baseline backup, via the `IntuneBackupAndRestore` module; read-only |

### 2026-07-22 (6)
| Change |
|--------|
| Updated [`scripts/Reporting/Get-SharePointStorageReport.ps1`](scripts/Reporting/Get-SharePointStorageReport.ps1) to make full-site scans GDAP-proof by resolving and pinning one effective tenant context (`-TenantId`, or GDAP customer context from `$global:cid`) for Graph sign-in, temporary app creation, and app-only token issuance |
| Added guard rails for GDAP/app-only flows: clearer errors when customer tenant context is missing (run `Connect-Tenant` first or pass `-TenantId`) and when `-ClientId` is provided without a resolvable tenant ID |
| Updated checkpoint signature inputs in [`Get-SharePointStorageReport.ps1`](scripts/Reporting/Get-SharePointStorageReport.ps1) to include `-ForceAppOnlySingleSite` and resolved tenant context, preventing cross-context resume collisions |
| Updated [`scripts/Reporting/readme.md`](scripts/Reporting/readme.md) to document full-site GDAP tenant binding behavior |

### 2026-07-22 (5)
| Change |
|--------|
| Updated [`scripts/Reporting/Get-SharePointStorageReport.ps1`](scripts/Reporting/Get-SharePointStorageReport.ps1) to make single-site scans GDAP-proof: when `authMode=GDAP` is detected, the script now automatically uses the temporary app/app-only bootstrap path for `-SiteUrl` scans to avoid delegated permission gaps |
| Added `-ForceAppOnlySingleSite` parameter to explicitly force app-only bootstrap for single-site scans, independent of detected auth mode |
| Updated [`scripts/Reporting/readme.md`](scripts/Reporting/readme.md) to document the new GDAP behavior and `-ForceAppOnlySingleSite` parameter |

### 2026-07-22 (4)
| Change |
|--------|
| Hardened [`scripts/Reporting/Get-SharePointStorageReport.ps1`](scripts/Reporting/Get-SharePointStorageReport.ps1) delegated path to remove dependency on missing Graph cmdlets: replaced `Get-MgDriveItemChild` usage with Graph REST pagination via `Invoke-MgGraphRequest` for drive children traversal |
| Updated site-drive enumeration fallbacks in [`Get-SharePointStorageReport.ps1`](scripts/Reporting/Get-SharePointStorageReport.ps1) to use Graph REST (`/sites/{id}/drives`) instead of `Get-MgSiteDrive` in delegated/non-app-only branches |
| Updated recycle-bin retrieval in [`Get-SharePointStorageReport.ps1`](scripts/Reporting/Get-SharePointStorageReport.ps1) to use Graph REST pagination (`/sites/{id}/recycleBin/items`) as fallback/primary delegated path instead of `Get-MgSiteRecycleBinItem` cmdlet dependency |
| Suppressed non-fatal MSAL authority warning noise on disconnect by wrapping `Disconnect-MgGraph` with temporary warning suppression in cleanup |

### 2026-07-22 (3)
| Change |
|--------|
| Updated [`scripts/Reporting/Get-SharePointStorageReport.ps1`](scripts/Reporting/Get-SharePointStorageReport.ps1) scan-mode handling so single-site runs (`-SiteUrl` with `/sites/...` or `/teams/...`) no longer trigger temporary app registration + app-only bootstrap; app-only setup is now only used for tenant-wide enumeration |
| Improved single-site lookup performance and reliability by resolving the exact site directly via Graph URL path (`/sites/{hostname}:{path}`) instead of search/filter flow |
| Updated [`scripts/Reporting/readme.md`](scripts/Reporting/readme.md) performance notes to document the single-site optimized path and expected startup speed behavior |

### 2026-07-22 (2)
| Change |
|--------|
| Fixed SharePoint report dependency issue causing `Get-MgSite` command-not-found errors: updated [`load.ps1`](load.ps1), [`scripts/Startup/Install-Modules.ps1`](scripts/Startup/Install-Modules.ps1), and [`scripts/Startup/Update-Modules.ps1`](scripts/Startup/Update-Modules.ps1) to include `Microsoft.Graph.Sites` |
| Updated [`scripts/Reporting/Get-SharePointStorageReport.ps1`](scripts/Reporting/Get-SharePointStorageReport.ps1) with an explicit module preflight check for `Microsoft.Graph.Authentication` and `Microsoft.Graph.Sites`, including a clear install hint when modules are missing |
| Updated [`scripts/Startup/readme.md`](scripts/Startup/readme.md) module dependency documentation to include `Microsoft.Graph.Sites` in install/update requirements |

### 2026-07-22 (1)
| Change |
|--------|
| Updated [`load.ps1`](load.ps1) — added startup launcher switches `-SetupStartup` / `-RemoveStartup`; first-run config now stores delegated auth defaults (`authMode`, optional `defaultCustomerDomain`, `useDeviceCodeAuth`) in `load.config.ps1` |
| Updated [`menu.ps1`](menu.ps1) — added Startup actions `F` (Enable-LauncherStartup) and `G` (Disable-LauncherStartup); added M365 action `H` (Test-GdapConnection); Entra submenu now includes temporary CA and TAP actions (`D`/`E`/`F`) |
| Updated [`scripts/Startup/functies.ps1`](scripts/Startup/functies.ps1) — Graph startup connection now supports delegated device-auth preference + required scopes for GDAP flow; added `Test-GdapConnection` helper for delegated contract/connectivity checks |
| Updated startup module maintenance: [`scripts/Startup/Install-Modules.ps1`](scripts/Startup/Install-Modules.ps1) and [`scripts/Startup/Update-Modules.ps1`](scripts/Startup/Update-Modules.ps1) now include `Microsoft.Graph.Identity.DirectoryManagement` |
| Added [`scripts/Entra/New-TemporaryConditionalAccessPolicy.ps1`](scripts/Entra/New-TemporaryConditionalAccessPolicy.ps1) — create temporary CA policy for user/group with either duration-based window or exact local start/end datetime; optional same-session auto-cleanup at end time |
| Added [`scripts/Entra/Remove-TemporaryConditionalAccessPolicies.ps1`](scripts/Entra/Remove-TemporaryConditionalAccessPolicies.ps1) — remove one or multiple temporary CA policies (`TEMP-CA -`), including expired-only or remove-all modes |
| Added [`scripts/Entra/New-UserTemporaryAccessPass.ps1`](scripts/Entra/New-UserTemporaryAccessPass.ps1) — create Temporary Access Pass (TAP) for a user with configurable lifetime and one-time option |
| Updated docs for the above across [`readme.md`](readme.md), [`scripts/readme.md`](scripts/readme.md), [`scripts/Startup/readme.md`](scripts/Startup/readme.md), and [`scripts/Entra/readme.md`](scripts/Entra/readme.md); clarified that temporary CA auto-cleanup runs in the current session (no Scheduled Task created) |

### 2026-07-09 (8)
| Change |
|--------|
| [`scripts/Reporting/Get-SharePointStorageReport.ps1`](scripts/Reporting/Get-SharePointStorageReport.ps1) — added a "site collection totals" phase (`-Apply` only, Phase 2c): sub-sites/Teams-kanalen and the recycle bin are now automatically rolled up per root site collection into `SharePoint_SiteCollectionTotals_<timestamp>.csv`, so the grand total is directly comparable to the single "storage used" figure shown per site in the SharePoint admin center |
| Documented the new output and the most likely causes of a remaining mismatch with the admin portal figure (timing lag, silently skipped folders on permission errors, failed version lookups) in [`scripts/Reporting/readme.md`](scripts/Reporting/readme.md) |

### 2026-07-09 (7)
| Change |
|--------|
| [`scripts/Reporting/Get-SharePointStorageReport.ps1`](scripts/Reporting/Get-SharePointStorageReport.ps1) — briefly extended recycle bin lookups (`-Apply`'s Phase 2b and `-RecycleBinOnly`) to also cover OneDrive personal sites, then reverted the same day on request — recycle bin scope stays SharePoint site collections only, OneDrive stays fully excluded (both storage scan and recycle bin) |
| Added a "Prullenbak (recycle bin)" section to [`scripts/Reporting/readme.md`](scripts/Reporting/readme.md) documenting the (SharePoint-only) recycle bin scope |

### 2026-07-09 (6)
| Change |
|--------|
| Removed the `Testing Scripts/` wrapper folder entirely — its subfolders duplicated existing top-level category names by verb (`Test-`/`Get-` prefix) rather than by domain. Merged its contents into the matching domain folder: `Testing Scripts/Entra/Test-M365GroupMembership.ps1` → [`Entra/`](scripts/Entra/readme.md), `Testing Scripts/Exchange/*` (6 scripts) → [`Exchange/`](scripts/Exchange/readme.md), `Testing Scripts/Device/Test-OpenVpnDiagnostics.ps1` → [`Device/`](scripts/Device/readme.md). [`Network/`](scripts/Network/readme.md), [`RDS/`](scripts/RDS/readme.md), and [`SMTP/`](scripts/SMTP/readme.md) (no existing top-level counterpart) were promoted to their own top-level category folders instead |
| Merged the corresponding readmes into each destination folder's existing readme.md rather than keeping separate "Testing —" docs |
| Updated 10 [`menu.ps1`](menu.ps1) script paths (Exchange audit submenu, Entra audit submenu, Test-Ports, SMTP tests) for the new locations |

### 2026-07-09 (5)
| Change |
|--------|
| Tidied `Testing Scripts/` and the repo root: moved `vias_archiver.ps1` out of `Testing Scripts/Device/` into a new [`scripts/Teams/`](scripts/Teams/readme.md) category — it's a Teams/SharePoint export & archiving tool, not a diagnostic script, so it didn't belong under "Testing" |
| Moved root-level [`Update-modules.ps1`](scripts/Startup/Update-Modules.ps1) into [`scripts/Startup/`](scripts/Startup/readme.md) (renamed [`Update-Modules.ps1`](scripts/Startup/Update-Modules.ps1) for naming consistency) — it's a module-maintenance script like [`Install-Modules.ps1`](scripts/Startup/Install-Modules.ps1), not a repo entry point like [`load.ps1`](load.ps1)/[`menu.ps1`](menu.ps1) |
| Removed the `Testing Scripts/SharePoint/` folder (it held only a pointer readme, no script) — that pointer now lives directly in `Testing Scripts/readme.md` |
| Documented [`Test-PowerShellSyntax.ps1`](scripts/Startup/Test-PowerShellSyntax.ps1) in [`scripts/Startup/readme.md`](scripts/Startup/readme.md), which had no docs before |

### 2026-07-09 (4)
| Change |
|--------|
| Optimized [`scripts/Reporting/Get-SharePointStorageReport.ps1`](scripts/Reporting/Get-SharePointStorageReport.ps1) version-history lookups, which were the main cause of the script appearing to hang on large libraries (one sequential Graph call per file, each eligible for up to 6 retries with backoff up to ~2 minutes on throttling): (1) skip the lookup entirely when a library is positively known to have versioning disabled, (2) batch up to 20 file version lookups per HTTP call via Graph's `$batch` endpoint instead of one call per file, (3) use a short, cheap 3-attempt retry for these specific calls instead of the main retry/backoff policy, since a failed lookup safely falls back to "0 versions" |
| Removed the now-unused `Get-VersionSize` function, replaced by `Invoke-GraphBatchGet` + batched resolution in `Get-AllDriveItems` |
| Updated [`scripts/Reporting/readme.md`](scripts/Reporting/readme.md) with a "Performance" section documenting the above |

### 2026-07-09 (3)
| Change |
|--------|
| Moved [`Deploy-OfficeTheme.ps1`](scripts/Custom%20Scripts/Intune/Desktop/Deploy-OfficeTheme.ps1), [`Deploy-Officecolors.ps1`](scripts/Custom%20Scripts/Intune/Desktop/Office%20Themes/Deploy-Officecolors.ps1), and their theme assets (`2026 Vias institute colours (2).thmx`, `Office Themes/`) back to [`Custom Scripts/Intune/Desktop/`](scripts/Custom%20Scripts/Intune/Desktop/readme.md) — both scripts hardcode their download URL to that exact repo path, so this keeps the URL valid instead of requiring a script update + Intune redeploy. Recreated [`Custom Scripts/`](scripts/Custom%20Scripts/readme.md) and [`Custom Scripts/Intune/`](scripts/Custom%20Scripts/Intune/readme.md) as minimal path-pinned folders (just this one item) rather than the full former category |
| [`scripts/Intune/Desktop/`](scripts/Intune/Desktop/readme.md) now holds only wallpaper/lockscreen/taskbar-shortcut deployment; both `Intune/readme.md` and `Intune/Desktop/readme.md` cross-reference [`Custom Scripts/Intune/Desktop/`](scripts/Custom%20Scripts/Intune/Desktop/readme.md) for the Office theme scripts |

### 2026-07-09 (2)
| Change |
|--------|
| Removed the [`Custom Scripts/`](scripts/Custom%20Scripts/readme.md) wrapper folder — it mixed generic tooling with customer-specific scripts under one confusing label, and duplicated the [`Intune/`](scripts/Intune/readme.md) category. Contents redistributed to proper top-level categories: `Custom Scripts/device/` → [`Device/`](scripts/Device/readme.md), `Custom Scripts/DNS/` → [`DNS/`](scripts/DNS/readme.md), `Custom Scripts/SAS/` → [`SAS/`](scripts/SAS/readme.md), `Custom Scripts/Save install time/` → [`Deployment/`](scripts/Deployment/readme.md) (renamed), [`Custom Scripts/Intune/Desktop/`](scripts/Custom%20Scripts/Intune/Desktop/readme.md) → merged into [`Intune/Desktop/`](scripts/Intune/Desktop/readme.md) |
| Updated [`menu.ps1`](menu.ps1) script paths for [`Restart-Time-Sync.ps1`](scripts/Device/Time%20sync/Restart-Time-Sync.ps1), [`detect-audiodevices.ps1`](scripts/Device/audio/detect-audiodevices.ps1), [`Disable-internalmic.ps1`](scripts/Device/audio/Disable-internalmic.ps1) to their new [`scripts/Device/`](scripts/Device/readme.md) location |
| Updated cross-references in [`scripts/Intune/readme.md`](scripts/Intune/readme.md), [`scripts/Intune/Get-Autopilot/readme.md`](scripts/Intune/Get-Autopilot/readme.md), and [`scripts/readme.md`](scripts/readme.md) for the new folder locations |
| ~~**Known issue (intentional):** [`Deploy-OfficeTheme.ps1`](scripts/Custom%20Scripts/Intune/Desktop/Deploy-OfficeTheme.ps1) and [`Deploy-Officecolors.ps1`](scripts/Custom%20Scripts/Intune/Desktop/Office%20Themes/Deploy-Officecolors.ps1) still hardcode their download URL to the old path — left unchanged on request.~~ **Resolved above** — the scripts moved back to match their hardcoded URL instead. |

### 2026-07-09
| Change |
|--------|
| Added [`readme.md`](readme.md) to every folder that lacked one: [`scripts/`](scripts/readme.md), [`scripts/Custom Scripts/`](scripts/Custom%20Scripts/readme.md), [`scripts/Custom Scripts/Intune/`](scripts/Custom%20Scripts/Intune/readme.md) (+ `Desktop/`, `Office Themes/`, `Add Lockscreen to start and desktop/`, `Background/`), `scripts/Custom Scripts/device/Time sync/`, [`scripts/Graph/`](scripts/Graph/readme.md), [`scripts/Intune/`](scripts/Intune/readme.md) (+ `Get-Autopilot/`), `scripts/Testing Scripts/`, `scripts/Testing Scripts/Network/`, `scripts/Testing Scripts/RDS/` — each with a file list and parameter/usage docs |
| Fixed `scripts/Custom Scripts/device/audio/Rollback-InternalMic` — file was missing its `.ps1` extension |
| Renamed `scripts/Entra/remove-m365users.ps1` → [`Remove-M365Users.ps1`](scripts/Entra/Remove-M365Users.ps1) for naming consistency (menu.ps1 and readmes already referenced the PascalCase form) |
| Fixed [`scripts/Entra/readme.md`](scripts/Entra/readme.md) — removed a stale `Distributionlist.ps1` entry that actually documented [`scripts/Exchange/Set-Distributionlist-dynamic-static.ps1`](scripts/Exchange/Set-Distributionlist-dynamic-static.ps1); moved accurate docs to [`scripts/Exchange/readme.md`](scripts/Exchange/readme.md); added missing [`Set-UserManager.ps1`](scripts/Entra/Set-UserManager.ps1) docs |
| Fixed `scripts/Testing Scripts/SharePoint/readme.md` — was a stale duplicate of [`Get-SharePointStorageReport.ps1`](scripts/Reporting/Get-SharePointStorageReport.ps1) docs (script doesn't live in this folder); replaced with a pointer to [`scripts/Reporting/readme.md`](scripts/Reporting/readme.md), which now documents the script's full current parameter set (`-ClientId`, `-ClientSecret`, `-CertificateThumbprint`, `-RecycleBinOnly`, `-GraphTimeoutSec`, `-MaxGraphRetry` were previously undocumented) |
| Fixed `scripts/Custom Scripts/device/audio/readme.md` — corrected script name casing to match the actual files on disk |
| Removed tracked `.DS_Store` files from git and added `.DS_Store` to [`.gitignore`](.gitignore) |

### 2026-04-17
| Change |
|--------|
| Updated `scripts/Custom Scripts/Save install time/start.bat` — added option `D` (customer install scripts from local `Install` folder) and option `E` (customer install scripts from a network share); before deployment starts, creates/updates local admin `LocalAdmin` (`<password omitted>`), adds it to `Administrators`, and sets OOBE skip flags |
| Added `scripts/Custom Scripts/Save install time/Browse-InstallScripts.ps1` — customer-first browser that lists customer folders as menu items and launches `.ps1`, `.bat`, and `.cmd` scripts |
| Updated `scripts/Custom Scripts/Save install time/readme.md` — documented new `D`/`E` menu options, including that option `D` requires copying both [`Browse-InstallScripts.ps1`](scripts/Deployment/Browse-InstallScripts.ps1) and the full `Install` folder, and that customer deploy options prepare `LocalAdmin` plus OOBE skip flags |

### 2026-04-16
| Change |
|--------|
| Updated `scripts/Custom Scripts/Intune/Desktop/Background/Lockscreen/Make-lockscreen.ps1` to v2.0 — aligned lockscreen source with corporate wallpaper config (`$ImageUrl`, `$ClientName`), replaced direct `WebClient` with validated internet download flow (`Invoke-WebRequest`), added GitHub blob/raw URL normalization, image signature checks (`jpg/png/bmp`), HTML-response guard, structured Intune logging, and safer temporary download handling |
| Added `scripts/Custom Scripts/Intune/Desktop/Background/Lockscreen/readme.md` — documentation for configuration, deployment, logging, workflow, and lockscreen-specific version history |
| Updated root [`readme.md`](readme.md) — expanded Intune Desktop/Background documentation and repository structure to include the lockscreen script and docs |

### 2026-04-15
| Change |
|--------|
| Updated `scripts/Custom Scripts/SAS/Test-SASWorkDirectory.ps1` — added drive profile diagnostics with explicit ephemeral drive awareness (`G:`/`U:`), improved I/O failure detail logging (exception type, inner exception, HResult), and adaptive low-space warning threshold for ephemeral scratch disks |
| Updated `scripts/Custom Scripts/SAS/Test-SASWorkDirectory.ps1` — refined SAS Application event analysis: explicit classification of `hc_disk_delete*` access-denied (`Return code 5`) as actionable failures, de-duplication of repeated SAS events, and informational handling of `ARM Application data not available` telemetry noise |
| Updated `scripts/Custom Scripts/SAS/Test-SASWorkDirectory.ps1` — added AV/EDR diagnostics: Defender status + exclusions, Defender Operational event scan, FilterManager/WdFilter System event scan, and active minifilter snapshot (`fltmc`) for correlation with intermittent WORK delete access-denied failures |
| Updated `scripts/Custom Scripts/SAS/readme.md` — documented ephemeral disk behavior for SAS `WORK`/`USERWORK` and clarified why switching between `G:` and `U:` is not a long-term failover strategy when both are ephemeral |
| Added `scripts/Custom Scripts/SAS/rca.md` — formal root cause analysis for intermittent SAS WORK delete failures, including evidence timeline, AV/ASR findings, root cause assessment, and remediation plan |

### 2026-04-13
| Change |
|--------|
| Updated `scripts/Testing Scripts/Device/vias_archiver.ps1` to v8.11 — Step 10 now supports non-interactive mode via `-Step10Only -Step10Action undo|archive|skip`; added explicit warning that archive/unarchive in Teams is a team-level action (not per channel) |
| Updated `scripts/Testing Scripts/Device/vias_archiver.ps1` to v8.12 — added per-channel soft-archive mode in Step 10 (rename marker with undo), including quick mode parameters `-Step10Only -ChannelAction archive|undo` and optional `-ChannelArchiveTag` |
| Updated `scripts/Testing Scripts/Device/vias_archiver.ps1` to v8.13 — switched per-channel flow to real Graph channel archive/unarchive API, added channel scope `ChannelSettings.ReadWrite.All`, and optional rename fallback via `-ChannelFallbackToRename` |
| Updated `scripts/Testing Scripts/Device/vias_archiver.ps1` to v8.14 — added `-DryRun` mode for Step 10 so team/channel archive/unarchive (and optional rename fallback) can be simulated without making changes |
| Updated `scripts/Testing Scripts/Device/vias_archiver.ps1` to v8.15 — expanded `-DryRun` to whole-script behavior: skips mutating setup/export/report/cleanup actions while keeping verification and simulated Step 10 output |
| Updated `scripts/Testing Scripts/Device/vias_archiver.ps1` to v8.16 — `-DryRun` now keeps temporary app creation, permission bootstrap, full login and export/report flow active; only Step 10 archive/unarchive mutations remain simulated |
| Updated `scripts/Testing Scripts/Device/vias_archiver.ps1` to v8.17 — refined `-DryRun` for Steps 6-9 to validate existence/counts (Teams/SharePoint/Graph) without writing exports; Step 11 report now uses these probe counts |
| Updated `scripts/Testing Scripts/Device/vias_archiver.ps1` to v8.18 — fixed channel lookup reliability (trimmed Excel Team/Channel values and reused cached Graph-fallback channel resolver in Step 9) to reduce false "Kanaal niet gevonden" in dry-run |
| Updated `scripts/Testing Scripts/Device/vias_archiver.ps1` to v8.19 — added normalized channel-name matching (trim/whitespace/case) in channel cache + Graph fallback lookup to better handle subtle name differences while dry-running |

### 2026-04-09
| Change |
|--------|
| Updated `scripts/Custom Scripts/Intune/Desktop/Background/Desktop/Set-CorporateWallpaper.ps1` v2.1 — added `Set-ExecutionPolicy Bypass -Scope Process` at the top to prevent exit code 3 when Intune's execution policy blocks the script |

### 2026-04-08
| Change |
|--------|
| Updated [`scripts/Reporting/Licensing/genereer_licentie_overzicht.py`](scripts/Reporting/Licensing/genereer_licentie_overzicht.py) — when the same product appears with multiple billing periods on one invoice (Pax8 and Ingram), each period is now shown as a separate row with the period range in the Category/Detail column instead of being summed incorrectly |
| Updated [`scripts/Reporting/Licensing/genereer_rapport.ps1`](scripts/Reporting/Licensing/genereer_rapport.ps1) — CMD window now closes automatically when run as a scheduled task; `Read-Host` pauses are skipped when `[Environment]::UserInteractive` is false |
| Updated [`scripts/Reporting/Licensing/genereer_rapport.bat`](scripts/Reporting/Licensing/genereer_rapport.bat) — pipes stdin from `NUL` so Python's interactive pause is never triggered when run as a scheduled task |

### 2026-04-01
| Change |
|--------|
| Updated [`scripts/Exchange/Migrate-Calendar.ps1`](scripts/Exchange/Migrate-Calendar.ps1) — fixed Room Mailbox booking issues: increased provisioning wait from 15s to 60s; added retry loop (5×30s) for `Set-CalendarProcessing` with error handling and fallback instructions; changed `BookingWindowInDays 0` to `1825` and added `EnforceSchedulingHorizon $false` to prevent silent booking rejections |

### 2026-03-30
| Change |
|--------|
| Updated `scripts/Custom Scripts/Intune/Desktop/Background/Desktop/Set-CorporateWallpaper.ps1` — added idempotency check: downloads image to temp, compares SHA256 hash against existing file; skips if hash matches and PersonalizationCSP is correct; applies (without second download) if image is new or changed |
| Updated readme — Intune & Autopilot section: expanded [`Set-CorporateWallpaper.ps1`](scripts/Intune/Desktop/Background/Desktop/Set-CorporateWallpaper.ps1) documentation with configuration table, deployment steps, log path, and NinjaOne/Intune deploy instructions |

### 2026-03-27
| Change |
|--------|
| Added [`scripts/Reporting/Get-ComputerLastLogon.ps1`](scripts/Reporting/Get-ComputerLastLogon.ps1) — last logon report for computers in one or more OUs; `LastLogonTimestamp` (fast) or `-AllDCs` (accurate) mode; marks Active/Stale/Never/Disabled; exports timestamped CSV to `C:\Temp\`; `-InactiveDays`, `-IncludeDisabled`, `-ExportPath` parameters |
| Added [`scripts/Reporting/readme.md`](scripts/Reporting/readme.md) — documents [`Get-ComputerLastLogon.ps1`](scripts/Reporting/Get-ComputerLastLogon.ps1) with parameter table, CSV column reference, and usage examples |
| Updated [`scripts/Reporting/Get-ComputerLastLogon.ps1`](scripts/Reporting/Get-ComputerLastLogon.ps1) — added `PasswordLastSet` / `DaysSincePasswordSet` columns; new `Active (pwd recent)` status for devices falsely marked stale due to 14-day `LastLogonTimestamp` replication delay |

### 2026-03-26
| Change |
|--------|
| Updated `scripts/Custom Scripts/device/audio/Disable-internalmic.ps1` — expanded internal mic patterns: added Conexant language variants (EN/FR/NL), Synaptics EN, Intel SST, IDT, Cirrus Logic drivers, and multilingual microphone array names (FR/DE/ES/PT/IT) |
| Updated readme — Audio Management section: added NinjaOne deployment table (Run as SYSTEM, no parameters, `AudioDeviceInventory` custom field, exit codes) for all three audio scripts |
| Updated readme — [`Invoke-WindowsActivation.ps1`](scripts/Device/Invoke-WindowsActivation.ps1): added NinjaOne deployment table with Run as Administrator, exit codes, and parameter examples per scenario; warning added for `-RemoveKey`/`-ReArm` requiring `-Force` |
| Added `scripts/Testing Scripts/RDS/Watch-RDSLive.ps1` — real-time RDS monitor: polls session events (20/21/22/23/24/25/40), failed RDP logons (4625), lockouts (4740), and licensing events every 20s; heartbeat per poll with session count; run directly on each RDS server |

### 2026-03-25
| Change |
|--------|
| Updated `scripts/Testing Scripts/Network/Test-FileIODiagnostics.ps1` — merged real-time monitor: FileSystemWatcher, NTFS permission diff vs baseline, auto-download Sysinternals Handle.exe, process snapshot diff, Kerberos tickets at failure time, Security audit events (4625/4740/4656/4663/4670); stops after 3 failures |
| Added `scripts/Testing Scripts/Network/Test-FileIODiagnostics.ps1` — file I/O stress test on any path (local or UNC/mapped drive); categorises failures as AUTH / NETWORK / TIMEOUT / DISK / PATH; auto-collects Kerberos tickets, net use, SMB port check and Security event log on first failure; `-Iterations`, `-StopOnFirstError`, `-DelayMs` parameters |
| Updated `scripts/Testing Scripts/RDS/Test-RDSDiagnostics.ps1` — add RDWeb Event Viewer analysis: TerminalServices-WebAccess/Admin+Operational, TerminalServices-Gateway/Admin+Operational, IIS/ASP.NET errors from Application log; triggered by -IncludeEventLogs |
| Added `scripts/Testing Scripts/RDS/Test-RDSDiagnostics.ps1` — diagnose RDP/RDWeb login failures: services, registry, NLA, session limits, licensing, firewall, HTTPS cert, IIS app pool, user account (enabled/locked/expired/group), event logs (4625/4740/4771/20/40); timestamped log to C:\Temp\ |
| Added `scripts/Custom Scripts/device/Invoke-WindowsActivation.ps1` — activate Windows, install product key, configure KMS server/port, remove key, ReArm grace period; dry-run safe with confirmation prompts; -Force to skip |
| Updated `scripts/Custom Scripts/device/Invoke-WindowsCleanup.ps1` — add dynamic log folder scan (section 13): recursively scans entire C:\ drive (max depth 7) for folders named logs/log/logging/diagnostics; skips Windows system dirs and dev artifacts (node_modules, .git, venv); fixed Windows system log paths (CBS archived .cab, DISM, WU, Panther, IIS); new `-SkipAppLogs` parameter |
| Updated `scripts/Custom Scripts/device/Invoke-WindowsCleanup.ps1` — scan all user profiles in C:\Users\ for temp, WER, thumbnail/shader cache and browser caches (Edge multi-profile, Chrome multi-profile, Firefox); summary shows reclaimable space per category |
| Fixed `scripts/Custom Scripts/device/Invoke-WindowsCleanup.ps1` — fix PS analyzer warnings: rename `$profile` to `$ffProfile`, drop unused `$dismResult` assignment |
| Fixed `scripts/Custom Scripts/device/Invoke-WindowsCleanup.ps1` — replace `??` null-coalescing operator with `-as [int64]` for PowerShell 5.1 compatibility |
| Added `scripts/Custom Scripts/device/Invoke-WindowsCleanup.ps1` — comprehensive Windows disk cleanup: temp files, WU cache, Delivery Optimization, Prefetch, memory dumps, WER, thumbnail/shader cache, Recycle Bin, browser caches, event logs, DISM component store; dry-run by default, `-Apply` to execute |
| Updated [`Get-SharePointStorageReport.ps1`](scripts/Reporting/Get-SharePointStorageReport.ps1) — two-phase approach (enumerate all sites/libraries first, then retrieve storage data); fixed inline `if` syntax error |
| Added [`Test-AuthNetworkDiagnostics.ps1`](scripts/Network/Test-AuthNetworkDiagnostics.ps1) — auth & network diagnostics: Event Viewer (4625/4771/4776/4740/5719), time sync, Kerberos cache, DNS, TCP, UNC shares, optional log scan |
| Added `scripts/Custom Scripts/SAS/` — SAS batch error monitoring with Zabbix integration, email alerts, and Event Viewer analysis |

### 2026-03-24
| Change |
|--------|
| Added [`scripts/Intune/iOS-Compliance-Updater/`](scripts/Intune/iOS-Compliance-Updater/readme.md) — auto-update minimum iOS version in Intune compliance policy via Graph API; weekly scheduled task, dry-run support, one-time Setup.ps1 |
| Added [`Get-SharePointStorageReport.ps1`](scripts/Reporting/Get-SharePointStorageReport.ps1) — tenant-wide SharePoint storage report with version history per file; quick mode (quota) and full recursive scan |
| Added [`Test-OpenVpnDiagnostics.ps1`](scripts/Device/Test-OpenVpnDiagnostics.ps1) — OpenVPN Connect diagnostics: PnP adapters, services, routes, DNS, Event Log, conflicting VPN software; txt export to `C:\Temp\` |

### 2026-03-23
| Change |
|--------|
| All CSV exports now go to `C:\Temp\` (Windows) or `~/Downloads/` (macOS/Linux) |
| Added [`Import-DnsRecords.ps1`](scripts/DNS/Import-DnsRecords.ps1) — resolve public DNS via dig (Google 8.8.8.8) and import A/CNAME records into AD DNS, dry-run by default |
| Added [`New-M365User.ps1`](scripts/Entra/New-M365User.ps1) — create single M365 user via Graph, auto-generated password, optional license |
| Added [`Import-M365Users.ps1`](scripts/Entra/Import-M365Users.ps1) — bulk user creation from CSV via Graph, dry-run by default, passwords in CSV output |
| Added [`Remove-M365Users.ps1`](scripts/Entra/Remove-M365Users.ps1) — bulk Entra ID user removal, dry-run by default, CSV report |
| Added [`Get-ExternalForwards.ps1`](scripts/Exchange/Get-ExternalForwards.ps1) — audit external forwarding rules across all mailboxes, CSV export |
| Added [`Get-MailboxSizes.ps1`](scripts/Exchange/Get-MailboxSizes.ps1) — mailbox size + item count report, sorted by storage, CSV export |
| Added [`Test-DkimConfig.ps1`](scripts/Exchange/Test-DkimConfig.ps1) — DKIM signing config + DNS CNAME/TXT validation with required-actions output |
| Added [`Test-M365GroupMembership.ps1`](scripts/Entra/Test-M365GroupMembership.ps1) — M365 Group / Teams owner and member audit via Graph, CSV export |
| Added [`Test-DistributionGroupPermissions.ps1`](scripts/Exchange/Test-DistributionGroupPermissions.ps1) — DG managers, Send As, Send on Behalf, member counts, CSV export |
| Added [`Test-CalendarPermissions.ps1`](scripts/Exchange/Test-CalendarPermissions.ps1) — locale-independent calendar permission audit, CSV export |
| Added [`Test-MailboxPermissions.ps1`](scripts/Exchange/Test-MailboxPermissions.ps1) — Full Access / Send As / Send on Behalf audit, CSV export |
| Exchange submenu (`C`) and Entra submenu (`D`): permission audit tools added |
| Moved [`Test-Ports.ps1`](scripts/Network/Test-Ports.ps1) to `scripts/Testing Scripts/Network/` |
| [`load.ps1`](load.ps1) — auto-imports modules at startup; detects missing modules and offers install |
| Added [`load.ps1`](load.ps1) — first-run setup (UPN + name), saves to gitignored `load.config.ps1`, launches menu |
| Extended [`menu.ps1`](menu.ps1) with M365 section (B–E): Exchange, Entra ID, MSP Admin submenus |
| Added [`menu.ps1`](menu.ps1) — interactive launcher, single-keypress, number + F-keys, cross-platform |
| Added [`scripts/Network/Test-Ports.ps1`](scripts/Network/Test-Ports.ps1) — TCP port checker, range/list syntax, multi-target |
| Licensing scripts: translated to English, genericised, configurable export dir |
| Rewrote [`create_scheduled_task.ps1`](scripts/Reporting/Licensing/create_scheduled_task.ps1) — admin check, auto-detect Python, dynamic trigger date |
| Added [`scripts/Reporting/Licensing/`](scripts/Reporting/Licensing/readme.md) — Pax8 + Ingram → Excel report toolkit |

### 2026-03-20
| Change |
|--------|
| [`start.bat`](scripts/Deployment/start.bat) v2.8 — split Do it all: A = Intune, C = AD; added device rename (B) and AD join (8) |
| Rewrote [`Migrate-Calendar.ps1`](scripts/Exchange/Migrate-Calendar.ps1) v2.0 — English, generic, mandatory params |
| Added SMTP test scripts; added `Test-SmtpRelay` to [`functies.ps1`](scripts/Startup/functies.ps1) |
| Rewrote [`functies.ps1`](scripts/Startup/functies.ps1) — replaced MSOnline/AzureAD with Microsoft Graph, cross-platform |
| Added [`Set-CorporateWallpaper.ps1`](scripts/Intune/Desktop/Background/Desktop/Set-CorporateWallpaper.ps1), [`Set-Calendar-rights.ps1`](scripts/Exchange/Set-Calendar-rights.ps1), [`Restart-Time-Sync.ps1`](scripts/Device/Time%20sync/Restart-Time-Sync.ps1) |
| Removed all company-specific references; translated all readmes to English |

### 2026-03-19
| Change |
|--------|
| Initial repository upload |

---

## Maintainer

**Sjoerd Kanon** — Security-minded | Team- & Projectgericht | Microsoft 365 & Infrastructure
