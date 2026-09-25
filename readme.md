# M365-Scripts

> A collection of PowerShell scripts and M365 management tools for MSP engineers, maintained by Sjoerd Kanon.

---

## Table of Contents

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

## Getting Started

```powershell
.\load.ps1
```

On first run, `load.ps1` will:

1. Ask for your admin UPN and display name — saved to a gitignored `load.config.ps1`
2. Ask whether you want delegated GDAP mode as default and optionally store a default customer domain
3. Ask whether Graph should use device code sign-in by default
4. Detect missing modules and offer to install them automatically
5. Import all required modules
6. Open the interactive menu

From then on it starts directly without any prompts.

To run the launcher automatically at Windows sign-in:

```powershell
.\load.ps1 -SetupStartup
```

To remove the startup shortcut later:

```powershell
.\load.ps1 -RemoveStartup
```

> You can also run `.\menu.ps1` directly — it will ask for your UPN as a fallback.
> To reinstall or update modules manually: `.\scripts\Startup\Install-Modules.ps1`

---

## Finding a script

| Where | What it gives you |
|-------|-------------------|
| [`scripts/INDEX.md`](scripts/INDEX.md) | Every script A–Z on one page — name, folder and what it does. Ctrl-F this when you know roughly what you want but not where it lives |
| [`scripts/readme.md`](scripts/readme.md) | The other direction: what each workload folder is for |
| [`.\menu.ps1`](menu.ps1) | The curated interactive launcher for the everyday tasks |
| `f <term>` | Fuzzy search from your shell, described under [Quick Launcher](#quick-launcher) below |
| Any folder readme | Every script name in a `Scripts` table links straight to the file, with a `docs` link to its section on the same page |

`INDEX.md` is generated from the scripts' own `.SYNOPSIS` headers by [`scripts/Startup/Update-ScriptIndex.ps1`](scripts/Startup/Update-ScriptIndex.ps1) — rerun it (or run it with `-Check`) whenever a script is added, renamed, moved or removed.

Those links are checked, not assumed: [`scripts/Startup/Test-MarkdownLinks.ps1`](scripts/Startup/Test-MarkdownLinks.ps1) walks every readme and fails on a file that is not there, or an anchor with no heading behind it.

---

## Quick Launcher

`f.ps1` creates one short command per script in [scripts/](scripts/), so you no longer have to navigate to a folder first. Dot-source it from your profile:

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
| Execution Policy | Windows only: `Set-ExecutionPolicy RemoteSigned -Scope CurrentUser` |

---

## Menu

The launcher (`menu.ps1`) covers all tools in this repo. Press a key to launch:

| Key | Category | Tool |
|-----|----------|------|
| `1` / `F1` | Testing | Test-Ports — TCP port checker |
| `2` / `F2` | Exchange | Migrate-Calendar |
| `3` / `F3` | Exchange | Set-Calendar-rights |
| `4` / `F4` | Testing | Test-SMTP (one-time) |
| `5` / `F5` | Testing | Test-SMTP (every 5 min) |
| `6` / `F6` | Device | Restart-Time-Sync |
| `7` / `F7` | Device | Detect-AudioDevices |
| `8` / `F8` | Device | Disable-InternalMic |
| `I` | Device | Remove-OemBloatware — remove OEM + generic Store bloatware |
| `T` | Device | Update-TeamsClient — update new Teams + the Outlook meeting add-in when outdated |
| `9` / `F9` | Startup | Install-Modules |
| `X` | Startup | Update-ScriptIndex — rebuild [`scripts/INDEX.md`](scripts/INDEX.md), the A–Z list of every script |
| `L` | Startup | Test-MarkdownLinks — check every readme link: files and in-page anchors |
| `M` | Startup | Convert-MarkdownToHtml — build a styled HTML page from a markdown document, for IT Glue |
| `A` / `F10` | Reporting | Licensing-Report |
| `P` | Reporting | SharePoint-Perms — report who has access to what, at every level |
| `S` | SharePoint | SharePoint-Structure — provision/check metadata, libraries and rights |
| `F` | Startup | Enable-LauncherStartup — add launcher to Windows Startup |
| `G` | Startup | Disable-LauncherStartup — remove launcher from Windows Startup |
| `B` | M365 | Connect-Tenant |
| `H` | M365 | Test-GdapConnection — validate delegated GDAP access |
| `C` | M365 | Exchange Online submenu |
| `D` | M365 | Entra ID / Graph submenu |
| `E` | M365 | MSP Admin submenu |

M365 options (`B`, `C`, `D`, `E`, `H`) lazy-load `functies.ps1` on first use — Graph authentication is only triggered when needed.

**Exchange submenu (`C`)**

| Key | Tool |
|-----|------|
| `8` | Test-CalendarPermissions — audit calendar folder permissions (all or single mailbox) |
| `9` | Test-MailboxPermissions — audit Full Access, Send As, Send on Behalf |
| `A` | Test-GroupPermissions — audit DG managers, Send As, Send on Behalf, member counts |
| `B` | Test-DkimConfig — validate DKIM signing config and DNS CNAME/TXT records |
| `C` | Get-ExternalForwards — audit mailboxes with external forwarding configured |
| `D` | Get-MailboxSizes — mailbox size report sorted by storage used |
| `E` | Move-InboxToArchive — archive Inbox messages to Archive folder |
| `F` | Set-DL-Dynamic-Static — resolve a dynamic distribution group into a static group |
| `H` | Get-CalendarMappings — where a calendar is mapped in Outlook, next to the rights (search by keyword, e.g. `balie`, or all/selected mailboxes) |
| `I` | Convert-SharedCalendar — move a shared calendar out of a user's mailbox into a room/equipment mailbox (always previews first) |
| `J` | Move-SharedCalendar — all in one: find a calendar by keyword, move it into a resource mailbox, list who has to switch |
| `K` | Get-DLMembers — export every distribution list with its members to Excel, or only the lists holding one address, one domain, or a domain tree (`-Recurse` to expand nested lists) |

**Entra ID submenu (`D`)**

| Key | Tool |
|-----|------|
| `A` | Test-M365GroupMembership — audit M365 Group / Teams owners and members |
| `B` | New-M365User — create a single new user (auto-generated password, optional license) |
| `C` | Import-M365Users — bulk create users from CSV, dry-run by default |
| `D` | New-TemporaryCA — create temporary Conditional Access policy for user/group (duration or start/end datetime) |
| `E` | Remove-TemporaryCA — remove expired or all temporary CA policies |
| `F` | New-UserTAP — create Temporary Access Pass for a user |
| `G` | Get-M365UserLicenses — report assigned licenses for a set of users |
| `H` | Import-CA-Baseline — import the community Conditional Access baseline |
| `I` | Set-UserManager — report/bulk-set manager for a set of users |

---

## Script Categories

### ☁️ M365 Management

Interactive M365 management functions via Microsoft Graph and Exchange Online. Loaded as a library through the menu. CSV and log exports go to `C:\Temp\` on Windows or `~/Downloads/` on macOS.

| Area | Features |
|------|----------|
| Exchange Online | Shared mailbox access, locale, aliases, distribution groups, auto-reply, sent-items copy |
| Entra ID / Graph | Tenant admins, domains, licenses, users, password reset, sign-in logs, bulk create/remove, temporary CA windows, TAP codes |
| MSP Admin | Create/manage MSP admin account across customer tenants |

---

### 📧 Exchange

Scripts for calendar and mailbox management.

- Calendar migration between users
- Set calendar folder permissions (NL/FR/EN locale support)
- **Get-DistributionGroupMembers.ps1** — who is on which distribution list, as one Excel workbook meant to go straight to the customer
  - `Overzicht` sheet (one row per list) and `Leden` sheet (one row per member), both filterable tables with a frozen header row, headers in Dutch
  - `-Member jan@contoso.com` answers "which lists is this person on?"; `-Member @be.verizon.com` answers it for one domain and `-Member *.verizon.com` for a domain and every subdomain, matching aliases and `ExternalEmailAddress` so external contacts are actually found
  - `-Recurse` expands nested lists — without it someone who only receives mail through a nested group is invisible, and a filter reports "no hits" on a list that does deliver to them
- **Get-MessageTraceReport.ps1** — trace who received what, at what exact time, and where it was forwarded to
- **Remove-PhishingMessage.ps1** — delete a phishing message from one, several, or all mailboxes; dry-run by default
  - Two engines: **Purview** Content Search + purge (tenant-wide, the only one that can HardDelete) and **Graph** (per-mailbox, no search-index lag, per-message report)
  - `Recycle` / `SoftDelete` / `HardDelete`; refuses to run without a content selector so a date range alone can never match every message
  - Loops purge rounds automatically around Purview's 10-items-per-mailbox limit, and writes a CSV of everything matched and deleted

---

### 👤 Entra ID / Graph

Scripts for user lifecycle management via Microsoft Graph.

- Create a single M365 user (auto-generated password, optional license)
- Bulk create users from CSV — dry-run by default, passwords in CSV output
- Bulk remove users from CSV — dry-run by default, CSV report

---

### 📱 Intune & Autopilot

Scripts for device enrollment, Autopilot registration, and compliance policy management.

- Retrieve Windows Autopilot hardware info
- CMD-based Autopilot enrollment helper
- **Compare-IntuneConfig.ps1** — compare a customer tenant's Intune configuration against an MSP baseline backup (drift detection), via the `IntuneBackupAndRestore` module — read-only
- **iOS Compliance Updater** — automatically keeps the minimum iOS version requirement in Intune up to date
  - Fetches latest iOS version from Apple's RSS feed (with fallback to Apple Support page)
  - Compares against current policy minimum and patches via Microsoft Graph API
  - One-time setup via `Setup.ps1` (creates App Registration, assigns permissions, writes `config.json`)
  - Runs weekly as a Windows scheduled task (SYSTEM, every Monday 07:00)
  - Dry-run mode (`-WhatIf`) — shows what would change without applying
- **Desktop** — deploy lockscreen to start and desktop; set corporate wallpaper via Intune:
  - `Set-CorporateWallpaper.ps1` — generic, reusable per customer; only the CONFIGURATION block needs updating
  - `Make-lockscreen.ps1` — applies the same corporate image as Windows lockscreen via PersonalizationCSP
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

Content operations on SharePoint Online sites and OneDrive via PnP PowerShell.

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

- `Install-SharePointStructure.ps1` — **the one-command build**: registers the Entra app itself, runs the three provisioning steps in the only order that works, verifies the result, and with `-TemporaryApp` removes the app registration again so nothing is left behind in a tenant you do not manage day to day
- The model lives in the config, not in the code: a second MSP client is a second config file, not a second fork of four scripts
- `New-SharePointMetadata.ps1` — managed metadata term set, site columns and content types, on **every** site in the config (a Teams private channel is its own site collection, and a site column does not reach across one)
- `Set-SharePointLibraries.ps1` — libraries, Teams channel folders, content type binding, per-folder content type order, default column values, grouped views, and one Entra ID security group per pillar per access level; `-EnsureGroups` creates the groups on the way
- `Update-SharePointShareStatus.ps1` — derives a Deelstatus column from the permissions actually on each file (Anyone link, guest, organisation link, or nothing) and flags anything tagged Intern/Vertrouwelijk sitting behind an external link; exit code 2 for a scheduled RMM job
- `Test-SharePointStructure.ps1` — read-only drift check classifying every difference as Missing / Different / Extra; exit code 2 means somebody changed something
- All four are idempotent and support `-WhatIf`; interactive or app-only with a certificate
- Cross-cutting brand views (`Scope = RecursiveAll`) make "brand as a tag" real: *Alles - Butterstone* is one flat list across every pillar folder, including everything tagged **Beide** — one file, two brands, no copies. Plus *Nog te taggen*, *Extern gedeeld* and *Te archiveren*
- [`Petsolutions-SharePoint-Handleiding.md`](scripts/SharePoint/Provisioning/Petsolutions-SharePoint-Handleiding.md) — end-user documentation in Dutch to hand to the customer: the three ways of adding a file and why they behave differently, what each label means, and what happens the moment you tag something
- Documented rather than hidden: unique permissions on a **standard**-channel folder are what this model asks for and what Microsoft does not support — members keep seeing the channel and get an error on the Files tab. `-SkipChannelFolderPermissions` is the conservative alternative

---

### 🧪 Testing & Diagnostics

Audit and diagnostic scripts, organised by workload. Self-connecting where applicable — reuse an existing session or connect automatically. CSV exports go to `C:\Temp\` on Windows or `~/Downloads/` on macOS.

#### Exchange Online

- Audit calendar folder permissions (locale-independent, exports CSV)
- Audit Full Access, Send As, Send on Behalf delegation (exports CSV)
- Audit distribution group managers, Send As, Send on Behalf, member counts (exports CSV)
- Validate DKIM signing config and DNS CNAME/TXT records; lists required actions
- Audit mailboxes with external forwarding to non-tenant domains (security audit, exports CSV)
- Report mailbox sizes and item counts sorted by storage used (exports CSV)

#### Entra ID / Graph

- Audit M365 Group (incl. Teams) owners and members — one row per entry, exports CSV
- Create temporary Conditional Access policies for installation windows (duration or exact local start/end)
- Auto-clean temporary CA policy at end time (same session) and cleanup script for missed sessions
- Create Temporary Access Pass (TAP) codes for user onboarding/support

#### SharePoint Online

- Report storage usage across all sites in a tenant — current file sizes + version history per library and per file
- Two-phase: first enumerates all sites and document libraries, then retrieves storage data
- Quick mode (quota data only) or full recursive scan with `-Apply`

#### Network & Connectivity

- Test TCP port connectivity on any host — single ports, ranges (`1294:1494`), combinations (`80,443,1294:1494`)
- One-time SMTP test with interactive credential prompt
- Recurring SMTP test (every 5 minutes) with saved encrypted password
- Auth & network diagnostics — Event Viewer (logon failures, Kerberos, NTLM, DC availability), time sync, DNS, TCP, UNC shares, optional log scan; exports txt report to `C:\Temp\`
- File I/O diagnostics — write/append/read/delete loop on any path; classifies failures as AUTH/NETWORK/TIMEOUT/DISK/PATH; on each failure captures FileSystemWatcher events, NTFS permission diff vs baseline, open process handles (Handle.exe auto-downloaded from Sysinternals), new process snapshot, Kerberos tickets, and Security event log; stops after 3 failures
- **UniFi** — network documentation HTML report (devices, firmware, uptime, per site) and firmware upgrade tooling for a UniFi Controller/UniFi OS console; credentials via `Get-Credential`, never hardcoded

#### Device

- OpenVPN Connect diagnostics — PnP adapters, services, routes, DNS, Event Log, conflicting VPN software; exports txt report to `C:\Temp\`

#### RDS

- RDP + RD Web Access diagnostics (`Test-RDSDiagnostics.ps1`) — diagnose why users cannot log in to an RDP or RDWeb server:
  - Services (TermService, SessionEnv, UmRdpService), RDP enabled/disabled, NLA, session limits, RD Licensing, firewall rules, active sessions
  - HTTPS certificate validity and expiry on RDWeb; IIS app pool and RD Gateway status (local only)
  - User account checks: enabled, locked, password expired, Remote Desktop Users membership
  - Event log analysis: failed logons (4625), lockouts (4740), Kerberos failures (4771), session disconnect reasons (20/40)
  - Timestamped log file saved to `C:\Temp\`; `-IncludeEventLogs` for event analysis

- Real-time RDS monitor (`Watch-RDSLive.ps1`) — polls event logs every N seconds and streams new events to console + log file:
  - Session events: logon (21), reconnect (22/25), logoff (23), disconnect (24), logon failed (20), disconnect reason (40) with human-readable reason codes
  - Security: failed RDP logons (4625 type 10), account lockouts (4740)
  - Licensing: `TerminalServices-Licensing/Admin` events + System log `TermServLicensing` provider
  - Heartbeat line per poll showing active session count and new event count
  - Run directly on each RDS/RDWeb server; `-IntervalSeconds` (default 20), `-NoLogFile` to skip file output

---

### 📊 Reporting

#### Computer Last Logon Report

Report last logon date for all computer objects in one or more OUs and export to CSV.

- Queries Active Directory for computers in specified OUs (e.g. `OU=Laptops`, `OU=Computers`)
- Two accuracy modes: `LastLogonTimestamp` (fast, max 14-day delay) or `-AllDCs` (queries every DC for exact `LastLogon`)
- Four statuses: **Active** · **Active (pwd recent)** · **Stale** · **Never** · **Disabled**
- `Active (pwd recent)`: device falsely marked stale due to 14-day replication delay — `PasswordLastSet` < 35 days confirms the machine is online (computer accounts auto-rotate password every ~30 days)
- CSV columns: Name, Status, Enabled, LastLogon, DaysSinceLogon, PasswordLastSet, DaysSincePasswordSet, OS, IPv4, OU path, Created, Description
- Supports multiple OUs in one run; `-IncludeDisabled` to include disabled objects

#### Licensing Report

Monthly licensing and Azure cost report generator.

- Combines Pax8 (CSV) and Ingram (Excel) billing data into one formatted Excel
- Per-customer tab with Azure consumption, licenses, and Acronis breakdown
- Summary tab with totals and margin per customer
- PowerShell launcher with pre-flight validation
- Optional Windows scheduled task (runs on the 6th of each month)

---

### 🖥️ Infrastructure & Devices

#### USB Setup Toolkit

USB toolkit for Windows setup and Autopilot enrollment during OOBE.

- Interactive menu (Device Manager, Autopilot, AD join, device rename, product key, Windows Update, restart)
- Customer install browser from USB toolkit menu:
  - Local `Install` folder by customer (`D`)
  - Network share `\\10.222.3.94\Software` by customer (`E`)
- For local option `D`, copy both `Browse-InstallScripts.ps1` and the complete `Install` folder next to `start.bat`
- Before options `D` and `E`, the toolkit creates/updates local admin `LocalAdmin` with password `Er@smus_Roter0`, adds it to `Administrators`, and sets OOBE skip flags
- Self-elevating, OOBE-compatible via Shift+F10
- Split "Do it all": `A` = Intune (Rename + Autopilot + Update), `C` = AD (Rename + Domain join + Update)

#### Audio Management

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

- Restart and force Windows Time service sync

#### Windows Cleanup

Comprehensive disk space cleanup for Windows endpoints.

- Cleans user/system temp folders, Windows Update download cache, Delivery Optimization cache, Prefetch, memory dumps, WER queues, thumbnail cache, DirectX shader cache, Recycle Bin, browser caches (Edge, Chrome, Firefox), and event logs
- DISM component store cleanup (`/StartComponentCleanup /ResetBase`) after Windows Updates
- DNS cache flush
- Dry-run by default — shows reclaimable space per category without deleting anything
- Run with `-Apply` to perform the actual cleanup; individual categories can be skipped with `-SkipBrowserCache`, `-SkipEventLogs`, `-SkipDism`, `-SkipRecycleBin`
- Exports a CSV report with bytes freed per category to `C:\Temp\`

#### OEM Bloatware Removal

- Detects device manufacturer (HP/Lenovo/Dell) and removes known OEM bloatware via `winget`, plus a generic list of consumer Microsoft Store apps (Xbox, Solitaire, Bing News/Weather, Cortana, Clipchamp)
- Dry-run by default; `-Apply` to actually remove. CSV report of found/removed apps to `C:\Temp\`

#### Cloud Drive Mapping

- Maps SharePoint/OneDrive document libraries to persistent drive letters via WebDAV (`net use`), for use as a per-user logon script (Intune Win32 app or scheduled task)
- Config-driven via a mappings CSV (`DriveLetter`, `Url`, optional `Label`); dry-run by default, `-Apply` to actually map
- No stored credentials — relies on the signed-in user's existing tenant session (same as browser WebDAV access)

---

### 🔧 Custom Tools

#### SAS Batch Monitoring

Monitor SAS batch job logs and Windows Event Viewer for errors, with optional Zabbix integration and email alerts.

- Detects spawn errors, WORK library auth failures, aborts, disk errors, and general `ERROR:` lines
- Text, JSON, and Zabbix output formats; configurable look-back period
- One-time setup script — installs to `C:\Scripts\`, creates daily scheduled task
- Optional Zabbix UserParameter config for automated alerting

#### Windows Device Management

Scripts for managing and maintaining Windows devices.

**Invoke-WindowsActivation.ps1** — Activate Windows or manage licensing settings:
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

**Invoke-WindowsCleanup.ps1** — Scan and optionally remove reclaimable disk space:
- User + system temp, Windows Update cache, Delivery Optimization, Prefetch
- Memory dumps, WER queues, thumbnail/DirectX shader cache, font cache
- Recycle Bin, browser caches (Edge/Chrome multi-profile + Firefox)
- Event logs, DISM component store (`/StartComponentCleanup /ResetBase`)
- Application & system logs: dynamic scan of entire C:\ for `logs`/`log`/`logging` folders
- Dry-run by default; use `-Apply` to delete. Per-category summary with space freed

#### DNS Management

Scripts for managing DNS records in Active Directory-integrated DNS zones.

- Resolve public DNS records via Google DNS (dig) and import them as A or CNAME records into AD DNS
- Dry-run by default — shows what would be created before applying
- Idempotent — skips records that already exist

---

### ☁️ Azure Infrastructure

Scripts that target Azure IaaS directly via the `Az` module — not the M365 tenant, and not wired into `menu.ps1`.

- **Azure-NVMe-Conversion.ps1** — vendored third-party script (Microsoft, MIT licensed, from `Azure/SAP-on-Azure-Scripts-and-Utilities`) that converts a VM's disk controller type between SCSI and NVMe, including in-guest driver readiness checks and fixes for both Windows and Linux guests
- **Search-AADDSUserActivity.ps1** — searches all Azure AD Domain Services audit tables in Log Analytics for a single user in one `union` query, instead of guessing which table an event landed in

---

### 🗄️ Legacy Toolkit Rewrites

A now-retired internal PowerShell repo (and two forked third-party GitHub toolkits it vendored) was reviewed script-by-script and modernized into this repo's house style — Graph/Exchange Online instead of the retired `MSOnline`/`AzureAD` modules, dry-run-by-default with `-Apply` for anything mutating, no hardcoded customer data or secrets. None of these are wired into `menu.ps1` — they're audit/reporting/setup scripts meant to be run directly, following the same pattern as `scripts/RDS/`, `scripts/Azure/`, and `scripts/Network/UniFi/`. Each folder has its own readme with full parameter/usage docs.

| Folder | Source | Covers |
|--------|--------|--------|
| [`TenantOnboarding/`](scripts/TenantOnboarding/readme.md) | Internal tenant-setup toolkit | New-tenant provisioning (break-glass admin, baseline groups/Intune assignment), multi-tenant/GDAP license + break-glass password reporting, Win32/Chocolatey app deployment, device config (kiosk power, Office uninstall, Start menu layout), OneDrive management, dynamic-DG/feature-group user management |
| [`Office365Toolkit/`](scripts/Office365Toolkit/readme.md) | Fork of [`directorcia/Office365`](https://github.com/directorcia/Office365) (CIAOPS) | Secure Score reporting, enterprise app consent cleanup, shared mailbox sign-in lockdown, EOP baseline, mailbox hygiene/forwarding-risk audits, Unified Audit Log search, Intune policy inventory |
| [`PatronToolkit/`](scripts/PatronToolkit/readme.md) | Fork of [`directorcia/patron`](https://github.com/directorcia/patron) | MFA registration + CA policy export, enterprise app consent + suspicious inbox rule + unified security alert audits, consolidated email security posture + mailbox auditing checks, SPF/DMARC validation, Intune policy assignment + Autopilot device reports, message trace, SharePoint sharing config, Teams config report |
| [`LegacyUtilities/`](scripts/LegacyUtilities/readme.md) | Internal toolkit (misc small scripts) | Mailbox folder permissions/delegate access, bulk shared mailbox/contact creation, contact sync, duplicate mail item cleanup, M365 group membership, CA policy backup, Teams/Planner cloning, Azure Files drive mapping, NumLock/lock-workstation device tweaks, Workspace 365 environment provisioning |

Both GitHub forks were reviewed capability-by-capability rather than ported 1:1 — near-duplicate single-purpose report scripts were consolidated into fewer well-parameterized ones, and capabilities already covered elsewhere in this repo were skipped rather than duplicated (see each folder's readme for the full skip list and reasoning). All code is a fresh implementation in this repo's style, not copied from the source projects.

---

## Repository Structure

Every folder has its own `readme.md` — this tree is a map; follow the links for full parameter/usage docs.

```
M365-Scripts/
├── .gitignore
├── .vscode/
│   └── settings.json
├── load.ps1                         ← Entry point: first-run setup + launches menu
├── menu.ps1                         ← Interactive launcher (all scripts + M365 functions)
├── readme.md
└── scripts/
    ├── readme.md                    ← Index of all categories below
    ├── INDEX.md                     ← Every script A-Z with its folder (generated)
    ├── Azure/                        ← targets Azure IaaS directly via Az, not the M365 tenant
    │   ├── readme.md
    │   └── VM/
    │       ├── readme.md
    │       └── Azure-NVMe-Conversion.ps1   ← vendored (Microsoft, MIT) — SCSI/NVMe disk controller conversion
    ├── Entra/
    │   ├── readme.md
    │   ├── Set-UserManager.ps1
    │   ├── Remove-M365Users.ps1
    │   ├── New-M365User.ps1
    │   ├── Import-M365Users.ps1
    │   ├── Get-M365UserLicenses.ps1
    │   ├── Import-ConditionalAccessBaseline.ps1
    │   └── Test-M365GroupMembership.ps1   ← audit M365 Group / Teams owners and members
    ├── Exchange/
    │   ├── readme.md
    │   ├── Migrate-Calendar.ps1
    │   ├── Convert-SharedCalendarToResource.ps1  ← shared calendar in a user's mailbox → its own room/equipment mailbox
    │   ├── Move-SharedCalendar.ps1  ← all in one: search by keyword + convert + who has to switch
    │   ├── Set-Calendar-rights.ps1
    │   ├── Set-Distributionlist-dynamic-static.ps1
    │   ├── Move-InboxToArchive.ps1
    │   ├── Test-CalendarPermissions.ps1
    │   ├── Get-CalendarMappings.ps1  ← where each calendar is mapped in Outlook, next to the rights behind it
    │   ├── Test-MailboxPermissions.ps1
    │   ├── Test-DistributionGroupPermissions.ps1
    │   ├── Test-DkimConfig.ps1
    │   ├── Get-ExternalForwards.ps1
    │   ├── Get-MailboxSizes.ps1
    │   └── Get-DistributionGroupMembers.ps1  ← who is on which distribution list, as an Excel workbook for the customer
    ├── Graph/
    │   ├── readme.md
    │   └── logic-permissies.ps1     ← grant a Graph app role to a Logic App managed identity
    ├── Intune/
    │   ├── readme.md
    │   ├── Compare-IntuneConfig.ps1  ← Intune config drift vs an MSP baseline backup (IntuneBackupAndRestore)
    │   ├── Get-Autopilot/
    │   │   ├── readme.md
    │   │   ├── Get-WindowsAutoPilotInfo.ps1
    │   │   └── GetAutoPilot.CMD
    │   ├── iOS-Compliance-Updater/
    │   │   ├── readme.md
    │   │   ├── Update-iOSCompliancePolicy.ps1   ← main script (run or scheduled task)
    │   │   ├── Setup.ps1                        ← one-time: App Registration + config.json
    │   │   ├── Install-ScheduledTask.ps1        ← register weekly scheduled task
    │   │   └── config.example.json
    │   └── Desktop/                  ← corporate wallpaper/lockscreen (Office theme lives in Custom Scripts/, see below)
    │       ├── readme.md
    │       ├── Add Lockscreen to start and desktop/
    │       │   ├── readme.md
    │       │   ├── add-lock.ps1               ← taskbar "Lock Workstation" shortcut
    │       │   └── add-shortcut-lock.ps1
    │       └── Background/
    │           ├── readme.md
    │           ├── Desktop/
    │           │   ├── readme.md
    │           │   ├── Set-CorporateWallpaper.ps1  ← corporate wallpaper via Intune (hash check, PersonalizationCSP)
    │           │   └── Remove-CorporateWallpaper.ps1
    │           └── Lockscreen/
    │               ├── readme.md
    │               └── Make-lockscreen.ps1         ← corporate lockscreen via Intune (validated download, PersonalizationCSP)
    ├── Device/
    │   ├── readme.md
    │   ├── Invoke-WindowsActivation.ps1 ← activate Windows, set product key / KMS server
    │   ├── Invoke-WindowsCleanup.ps1    ← temp, cache, WU, DISM, browser, event logs
    │   ├── Clear-TempFiles.ps1
    │   ├── Remove-OemBloatware.ps1      ← HP/Lenovo/Dell + generic Store bloatware removal
    │   ├── Test-OpenVpnDiagnostics.ps1  ← OpenVPN Connect diagnostics
    │   ├── Update-TeamsClient.ps1       ← update new Teams + meeting add-in when a newer build exists
    │   ├── Update-TeamsClient.md        ← how that script decides, step by step
    │   ├── Update-TeamsClient-ITGlue.md ← servicedeskversie (NL) om in IT Glue te plakken
    │   ├── audio/
    │   │   ├── readme.md
    │   │   ├── detect-audiodevices.ps1
    │   │   ├── Disable-internalmic.ps1
    │   │   └── Rollback-InternalMic.ps1
    │   ├── DriveMapping/
    │   │   ├── readme.md
    │   │   └── New-CloudDriveMapping.ps1   ← map SharePoint/OneDrive libraries to drive letters (WebDAV)
    │   └── Time sync/
    │       ├── readme.md
    │       └── Restart-Time-Sync.ps1
    ├── Network/
    │   ├── readme.md
    │   ├── Test-Ports.ps1
    │   ├── Test-AuthNetworkDiagnostics.ps1   ← auth/network issue diagnostics
    │   ├── Test-FileIODiagnostics.ps1        ← file I/O test + real-time directory monitor
    │   └── UniFi/
    │       ├── readme.md
    │       ├── UnifiApi.ps1                  ← shared login/session helper (dot-sourced)
    │       ├── Get-UnifiNetworkReport.ps1     ← HTML network documentation report
    │       └── Update-UnifiFirmware.ps1       ← list/trigger firmware upgrades across sites
    ├── RDS/
    │   ├── readme.md
    │   ├── Test-RDSDiagnostics.ps1           ← RDP/RDWeb login failure diagnostics
    │   └── Watch-RDSLive.ps1                 ← real-time session + licensing monitor
    ├── SMTP/
    │   ├── readme.md
    │   ├── testsmtp.ps1
    │   └── testsmtp_5min.ps1
    ├── Deployment/                   ← USB setup toolkit (OOBE / Autopilot)
    │   ├── readme.md
    │   ├── start.bat
    │   ├── autorun.inf
    │   └── Browse-InstallScripts.ps1
    ├── DNS/
    │   ├── readme.md
    │   ├── Import-DnsRecords.ps1   ← resolve via Google DNS + import into AD DNS
    │   └── example-records.csv
    ├── SAS/
    │   ├── readme.md
    │   ├── rca.md
    │   ├── Monitor-SASBatchErrors.ps1   ← scan logs + Event Viewer for SAS errors
    │   ├── Setup-SASMonitoring.ps1      ← install script, scheduled task, Zabbix config
    │   ├── Test-SASWorkDirectory.ps1    ← validate WORK directory health
    │   └── zabbix_sas_monitor.conf
    ├── SharePoint/
    │   ├── readme.md
    │   ├── Find-SiteContent.ps1         ← search a whole site (name/path/type/date or full text) + report the permissions on every hit (PnP)
    │   ├── Search-SharePointContent.ps1 ← same, tenant-wide via Graph app-only: delta + /permissions, sharing links and guests (files/folders)
    │   ├── Restore-RecycleBinItems.ps1  ← restore deleted files from a recycle bin: one site/OneDrive or tenant-wide (PnP, auto app registration)
    │   └── Provisioning/                ← provision a whole structure from one JSON config (PnP + Graph)
    │       ├── readme.md
    │       ├── Petsolutions-SharePoint-Handleiding.md ← end-user guide (NL) to hand to the customer
    │       ├── petsolutions.config.json     ← the model: columns, content types, groups, libraries, views, permissions
    │       ├── Install-SharePointStructure.ps1 ← build it all in one run, incl. (temporary) app registration
    │       ├── SharePointStructure.Common.ps1 ← shared helpers (dot-sourced by all four)
    │       ├── New-SharePointMetadata.ps1   ← term set, site columns, content types (every site in the config)
    │       ├── Set-SharePointLibraries.ps1  ← libraries/channel folders, content types, defaults, views, group rights
    │       ├── Update-SharePointShareStatus.ps1 ← derive the Deelstatus column, flag over-sharing (exit 2)
    │       └── Test-SharePointStructure.ps1 ← read-only drift check vs the config (exit 2)
    ├── Teams/
    │   ├── readme.md
    │   └── vias_archiver.ps1        ← Teams/SharePoint export + archiving (Graph, PS7+, Global Admin)
    ├── Reporting/
    │   ├── readme.md
    │   ├── Get-ComputerLastLogon.ps1        ← last logon per computer in OU(s), export to CSV
    │   ├── Get-SharePointStorageReport.ps1  ← tenant-wide SharePoint storage report
    │   ├── Get-SharePointPermissionsReport.ps1 ← who has access to what, at every level, to CSV
    │   ├── Remove-SharePointFileVersionsByDate.ps1 ← delete file versions older than a date
    │   └── Licensing/
    │       ├── readme.md
    │       ├── genereer_licentie_overzicht.py
    │       ├── genereer_rapport.ps1
    │       ├── genereer_rapport.bat
    │       └── create_scheduled_task.ps1
    ├── Startup/
    │   ├── readme.md
    │   ├── functies.ps1             ← M365 function library (dot-sourced by menu)
    │   ├── Install-Modules.ps1      ← Bootstrap: install & import all modules
    │   ├── Update-Modules.ps1       ← Update every installed PowerShell module
    │   ├── Test-PowerShellSyntax.ps1
    │   ├── Update-ScriptIndex.ps1   ← Regenerates scripts/INDEX.md from the .SYNOPSIS headers
    │   ├── Test-MarkdownLinks.ps1   ← Checks every readme link: files and in-page anchors
    │   └── Convert-MarkdownToHtml.ps1 ← Markdown doc → one self-contained styled HTML page
    ├── Custom Scripts/                 ← path-pinned scripts (see note above)
    │   ├── readme.md
    │   └── Intune/
    │       ├── readme.md
    │       └── Desktop/
    │           ├── readme.md
    │           ├── Deploy-OfficeTheme.ps1        ← installs the full VIAS .thmx Office theme
    │           ├── 2026 Vias institute colours (2).thmx
    │           └── Office Themes/
    │               ├── readme.md
    │               ├── Deploy-Officecolors.ps1   ← installs just the color scheme
    │               └── Test VIAS.xml
    ├── TenantOnboarding/                ← modernized from a retired internal tenant-setup toolkit, not menu-wired
    │   ├── readme.md
    │   ├── Provisioning/         (3 scripts)  ← break-glass admin, baseline groups, Intune policy assignment
    │   ├── MultiTenant/          (3 scripts)  ← GDAP license report, break-glass password rotation, customer portal index
    │   ├── AppDeployment/        (7 scripts)  ← Win32/Chocolatey install, shortcuts, file associations, printer connections
    │   ├── DeviceConfig/         (6 scripts)  ← kiosk power, Office uninstall, Start menu layout, Teams firewall rule
    │   ├── OneDriveManagement/   (3 scripts)  ← sync watchdog, library sync stop, known-folder redirect
    │   └── UserManagement/       (2 scripts)  ← dynamic DG by filter, feature-group membership
    ├── Office365Toolkit/                ← rewrite of retired directorcia/Office365 (CIAOPS) fork, not menu-wired
    │   ├── readme.md
    │   ├── Security/             (4 scripts)  ← Secure Score, app consent cleanup, shared mailbox lockdown, EOP baseline
    │   ├── Exchange/             (4 scripts)  ← mailbox hygiene baseline, forwarding risk, add-ins, audit log search
    │   └── Intune/               (1 script)   ← tenant-wide policy inventory
    ├── PatronToolkit/                    ← rewrite of retired directorcia/patron fork, not menu-wired
    │   ├── readme.md
    │   ├── Entra/                (2 scripts)  ← MFA registration report, CA policy export
    │   ├── Security/             (5 scripts)  ← app consents, suspicious inbox rules, security alerts, email security posture, mailbox auditing
    │   ├── Exchange/             (1 script)   ← message trace report
    │   ├── Intune/               (2 scripts)  ← policy assignments, Autopilot devices
    │   ├── SharePoint/           (1 script)   ← sharing config audit
    │   └── Teams/                (1 script)   ← Teams config report
    └── LegacyUtilities/                  ← misc modernized scripts from the retired internal toolkit, not menu-wired
        ├── readme.md
        ├── Exchange/             (7 scripts)  ← folder permissions, delegate access, bulk mailboxes/contacts, contact sync, dedup, message trace
        ├── Entra/                (2 scripts)  ← group membership, CA policy backup
        ├── Teams/                (3 scripts)  ← team/plan cloning, project team provisioning
        ├── Network/              (1 script)   ← Azure Files drive mapping
        ├── Device/               (2 scripts)  ← NumLock default, lock-workstation shortcut
        └── Workspace365/         (2 scripts)  ← environment provisioning/removal
```

`Deploy-OfficeTheme.ps1` and `Deploy-Officecolors.ps1` hardcode their download URL to this exact repo path (`main` branch) — they stay here rather than under `Intune/Desktop/` so the URL keeps resolving.

---

## Contributing

When adding new scripts:

1. Follow the existing naming convention (`Verb-Noun.ps1`)
2. Include a header comment block with Synopsis, Description, Parameters, and Example
3. Test against a non-production tenant before committing
4. Place the script in the appropriate workload folder
5. Add it to `menu.ps1` and update this readme
6. Update `Version History` in this file for every functional or structural change (required), including changes requested or applied via Copilot/AI assistant

---

## Disclaimer

These scripts are provided as-is. Always test in a non-production environment before running against live tenants. The maintainer accepts no liability for unintended changes resulting from misuse or misconfiguration.

---

## Version History

> Note: Older entries can reference historical folder names such as `Custom Scripts/` and `Testing Scripts/`. These path names reflect the repository structure at the time of that change.

### 2026-09-25 (10)
| Change |
|--------|
| `Convert-MarkdownToHtml.ps1` rendered every numbered list as empty bullets. `$Matches` is one variable per scope: the list branch captured the item text, then ran a second `-match` to decide whether the list was ordered, and that second match threw the capture away. Dashed lists survived only because their second match failed and left `$Matches` alone. Both captures now come from one match and the marker decides the type without matching again |
| A fenced code block indented to line up inside a numbered step kept that indentation, so anyone copying the command out of the page copied leading spaces with it. The fence's own indentation is now stripped from its content — and only that: a block fenced at column 0 keeps every space, which is what the sample output of the script needs |
| Verified on the regenerated page: 36 list items and **none** empty, still parses as XML, 29 tables and 19 code blocks intact, the indented command comes out clean, and the seven code blocks that legitimately start with whitespace still do |

### 2026-09-25 (9)
| Change |
|--------|
| You can now click from a readme straight to the script it describes. Every script name in a folder readme's `Scripts` table linked to a section further down the same page, never to the file — so the readme told you what a script did but gave you no way to open it. 172 links across 47 readmes now point at the file, with a `([docs](#…))` link beside them for the section that was there before |
| 34 of those were not links at all: a script name in a table cell, set in backticks, with nothing behind it. Those are file links now too |
| Added `scripts/Startup/Test-MarkdownLinks.ps1`, because links that are never checked are links that quietly rot. It walks every `.md` and fails on two things: a relative link to a file that is not there (percent-encoded spaces decoded first, the way GitHub serves them), and an anchor with no heading behind it. Anchors are resolved the way GitHub builds them, including the `-1`/`-2` suffix for repeated headings |
| It found three anchors that had never worked: `#watch-rdslivesps1` had an `s` too many for `### Watch-RDSLive.ps1`, and two links in the Intune readme used `#detect--remediate-…` where the heading `### Detect- / Remediate-StuckWin32AppEnforcement.ps1` produces `#detect---remediate-…` — three hyphens, because the slash becomes nothing and the spaces around it each become one. Nobody would find that by eye |
| Invisible characters are stripped from both the heading and the link before they are compared. Without that, the four emoji entries in the root table of contents read as broken: the heading and the link both carry a variation selector, which is not a letter and not a digit. Reproducing GitHub's slugger byte for byte on characters nobody can see is not the point — establishing that a link and a heading correspond is |
| Verified: 854 internal links across 71 markdown files all resolve, all 181 files parse, and the script index is current at 178 scripts. `L` added to the menu for the link check, alongside `X` for the index |

### 2026-09-25 (8)
| Change |
|--------|
| The Teams IT Glue procedure now also exists as a styled HTML page, `scripts/Device/Update-TeamsClient-ITGlue.html`, for pasting into IT Glue or printing. It is **generated** by the new [`scripts/Startup/Convert-MarkdownToHtml.ps1`](scripts/Startup/Convert-MarkdownToHtml.ps1) rather than written by hand: that document changed six times in two days, and a hand-made copy would have been wrong by the next morning |
| The converter covers what these documents actually use — headings, tables, fenced code blocks including the indented ones inside numbered steps, blockquotes, both list kinds, rules, and inline code, bold, italic and links — and passes anything else through as text instead of guessing. `-Check` writes nothing and exits `1` when the committed page has fallen behind its markdown, which is what a hook or pipeline would call |
| Void elements are emitted self-closing, so the page parses as XML as well as HTML. That is how it was verified rather than by looking at it: the output parses, and holds 29 tables, 190 rows, 19 code blocks and 46 headings, with **no** table containing a row that disagrees with its header width. Also checked: no `**`, backtick or `](` left anywhere in the rendered text, and the ✅/❌ and accented characters survive |
| Both `-Check` failure paths exercised: a page that does not exist yet, and a markdown file that has moved on — each exits `1` with the reason. The generated-on line is excluded from the comparison, so an unchanged document does not report a difference |
| Added as menu key **M**, documented in [`scripts/Startup/readme.md`](scripts/Startup/readme.md), and `scripts/INDEX.md` regenerated — adding a script had made it stale, which `Update-ScriptIndex.ps1 -Check` reported |

### 2026-09-25 (7)
| Change |
|--------|
| Merged the readable parts of an older IT Glue version of this same procedure into `Update-TeamsClient-ITGlue.md`: the one-sentence statement of what the procedure covers, and the ✅/❌ shape for "does this script fit this ticket", including the two user complaints that version listed and ours did not — Teams hanging on startup and Teams closing unexpectedly |
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
| The page is **generated**, by the new `scripts/Startup/Update-ScriptIndex.ps1`, so it cannot drift from the files the way a hand-kept table does. `-Check` reports a stale index without writing (exit `1`), which is what a hook or a pipeline would call; a run that finds the page current writes nothing at all |
| Descriptions come from the scripts themselves: the `.SYNOPSIS` block, joined across the lines it wraps over rather than taking the first line, which was leaving half-sentences like "Grant Full Access and/or Send As delegate rights on one mailbox, a CSV list of" in the table. Where a synopsis opens with a sentence and then lists its cases, the lead-in is kept and the list is not dragged in behind it |
| For the older scripts that have no `.SYNOPSIS`, a leading `#` comment block is used instead — but only a real header. A single comment line sitting straight on top of code describes that line, not the script: `# URL van de theme` above a `$ThemeUrl` assignment was being read as a description, which is worse in a table than a blank |
| The eight scripts that still ended up with nothing got a real `.SYNOPSIS` instead of a blank cell: `add-lock.ps1`, `add-shortcut-lock.ps1`, `logic-permissies.ps1`, `Test-OpenVpnDiagnostics.ps1`, `Deploy-OfficeTheme.ps1`, `Restart-Time-Sync.ps1`, `Test-PowerShellSyntax.ps1` and `functies.ps1`. All 176 scripts now describe themselves, so the index has no "without a description" section left |
| Documented the two scripts no readme mentioned at all: `Phising-rollout.ps1` in [`scripts/Entra/readme.md`](scripts/Entra/readme.md) (the two-way sync between the phishing-resistant MFA rollout and registered groups, what counts as registered and why the default is an AAGUID filter) and `Get-FSlogix-errors.ps1` in [`scripts/RDS/readme.md`](scripts/RDS/readme.md) (what the FSLogix diagnostic collects and that it must run on the session host). Its header still pointed at a filename that no longer exists, and named a real customer in the example; both corrected |
| The root `Menu` table had drifted from `menu.ps1` — `I`, `T`, `S` and `P` were missing. Synced, and `X` added for the index generator, which is also in the Startup readme and the repository tree |
| Verified: all 176 files parse; the generator is idempotent (a second run reports "already up to date" and writes nothing); `-Check` exits `0` when current; every markdown link in the repository resolves, percent-encoded folder names included; and `f.ps1` still finds both the new script and the newly described ones |
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
| Added `scripts/Reporting/Get-SharePointPermissionsReport.ps1` — an exhaustive read-only SharePoint Online permissions report: site collection admins, web role assignments including inheritance breaks, SharePoint groups with their full membership, list and library role assignments, every folder and item with a unique scope, sharing links with their kind, external/guest principals, `Everyone` grants, and Entra group grants resolved to transitive membership. Four CSVs: detail, per-site summary, group membership, and — behind `-IncludeEffectiveAccess` — one row per resolved user per scope with the group the access runs through |
| Inheritance is followed the way SharePoint models it: an item is only reported as its own scope when `HasUniqueRoleAssignments` is true, so the CSV is a map of the permission structure rather than a row per file. Site discovery is deliberately redundant — Graph `getAllSites`, then sub-sites through both Graph and SharePoint REST (`/_api/web/webs`), de-duplicated on URL — because Graph omits classic sub-webs |
| Authentication had to go app-only: role assignments are not readable through Graph at all, and are not covered by SharePoint's Read/Write/Manage application roles either — only `Sites.FullControl.All` can enumerate them. The script signs in interactively once, creates a short-lived App Registration with that role plus Graph `Sites.Read.All` and `GroupMember.Read.All`, and deletes it again on exit. Despite the Full Control role it only ever issues `GET`: it never writes and never changes a permission. `-ClientId`/`-TenantId` with a secret or certificate skips the temporary app |
| Resumable like the other long SharePoint scans: a checkpoint per completed list, keyed on a hash of the scan parameters, so an interrupted tenant run continues instead of starting over; `-Restart` discards it. Checkpoint files are only cleaned up once the final CSVs are written, so their presence is itself the signal that a run was interrupted |
| Documented in `scripts/Reporting/readme.md` (coverage, authentication, the four output files, checkpoints, full parameter table, examples), added to the root repository tree and to `menu.ps1` under Reporting as `P` — which also prompts for tenant or single site, scope, and whether to write the effective-access CSV |
| The repository tree also listed neither this script nor `Remove-SharePointFileVersionsByDate.ps1`; both are in it now |
| **Untested by me**: this script has not been run against a live tenant in this session — only its syntax was checked. The temporary App Registration path, the throttling retries and the checkpoint resume are unverified here and should be exercised on a pilot tenant, starting with `-SiteUrl` and `-Scope Site`, before a tenant-wide run |

### 2026-09-25 (2)
| Change |
|--------|
| An add-in that is registered but whose files are gone no longer counts as installed. That is exactly the state the failed run above left behind, and the script would have answered "Teams is up to date - nothing to do" on it: the work decision only looked at the Programs and Features entry, while the DLL check that spots this was report-only |
| `Update-TeamsClient.ps1` no longer deletes a working meeting add-in before knowing it can install a replacement. A production `-Force` run removed every copy in step 6 and then failed in step 8 with `1638`, leaving the session host with no add-in at all. The sweep moved to step 8, behind the version comparison: an older MSI than the registered add-in now means the sweep and the install are skipped and the working add-in is left exactly as it is |
| Three things had to line up for that, and all three are now handled. `-Force` on a host whose build is newer than the published one is a **downgrade**, which the version check now warns about by name. An add-in uninstall answering `1612` means Windows Installer lost its source, so it is retried against its own cached MSI under `C:\Windows\Installer` (via `Installer\UserData\S-1-5-18\Products\*\InstallProperties`, `LocalPackage`); when that is gone too, the run says the registration cannot be removed and what it will cause. A `1638` on the add-in is now a warning rather than an abort, so verification still runs and reports what Outlook is actually left with |
| Preflight prints the registered add-in version instead of just "is installed" - that single number was the whole diagnosis of the failure and it was the one thing not on screen |
| The AppLocker line no longer prints an empty summary when `SrpV2` exists with no rule collections under it (as on the production host): it says "no rule collections configured, so it blocks nothing" |
| Verified: the cached-package lookup against real installed products, and that asking for a product this machine lacks returns nothing without throwing; the version guard in all four combinations (older, newer, equal, unparsable); the empty-collection AppLocker line. **Untested:** the `msiexec /x <cached msi>` retry and the `1638` warning path on a live host |

### 2026-09-25
| Change |
|--------|
| `Update-TeamsClient.ps1` stops crying wolf about AppLocker. The SlimCore blocker check warned whenever `HKLM:\SOFTWARE\Policies\Microsoft\Windows\SrpV2` existed — which it does on any fleet that has ever written a single Exe rule — so the warning fired on machines where nothing was blocked at all. A check that always fires is a check nobody reads |
| The policy is now read instead of detected, on the three points that decide whether it can stop the MSIX: only the packaged-app (`Appx`) collection applies, because an MSIX never meets the `Exe`/`Msi`/`Script`/`Dll` rules; a collection holding rules whose enforcement is *not configured* is enforced all the same, per Microsoft, and only an explicit `EnforcementMode = 0` lets everything through; and nothing is enforced at all while the Application Identity service (`AppIDSvc`) is stopped, which is now said out loud rather than assumed either way |
| The report is something a technician can act on: the registry path, the mode per collection, the service state and the first five `Appx` rule names with their action. A rule that already allows the packages by name is reported as `[ OK ]`; a rule allowing anything signed by `O=MICROSOFT CORPORATION` is reported as probably sufficient, with a note to check it has not been narrowed to one product name |
| A policy found on a session host is now named as a `[SKIP]` reference line instead of being hidden: it blocks nothing there, because the staging happens on the endpoint, but it is usually the same GPO — so the thing worth checking is whether it also reaches the endpoints |
| Every branch exercised against a stubbed policy tree: an enforced `Exe` collection with no `Appx` collection produces no blocker (the old false positive, gone); an enforced `Appx` collection with no matching allow rule warns with path and rule count; explicit-SlimCore and Microsoft-publisher allow rules produce their two different notes; `EnforcementMode = 0` reads as audit only; enforcement-not-configured-with-rules reads as enforced; seven rules print five and `... and 2 more`; a missing `SrpV2` key produces nothing. Confirmed silent on this machine, which has no AppLocker policy. **Untested** against a live enforced AppLocker policy on a real endpoint |
| Known limitation, documented rather than hidden: the allow-rule match is a text match on the rule XML, so a broad rule that names neither Microsoft nor the packages (`PublisherName="*"`) would really let SlimCore through but is still reported as a blocker. The rule names printed beside it are what settles that |

### 2026-09-24
| Change |
|--------|
| `Update-TeamsClient.ps1` answers the question the inventory could not: preflight now reads the `Microsoft Teams VDI` events from the Application log on any session host — not just with `-AvdOptimizations` — and translates the codes from Microsoft's connection error table, so a plain `-CheckOnly` reports whether users are actually optimized instead of only whether the parts are installed |
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
| `Get-DistributionGroupMembers.ps1` — `-Member "*.verizon.com"` now matches a domain **and every subdomain of it** (`.verizon.com` and `*@*.verizon.com` are the same thing). Without the leading `*.` the filter stays on that one domain, so `@be.verizon.com` still deliberately does not reach `@us.verizon.com` |
| The run says which of the two it is doing — *"scanning N list(s) for members on verizon.com and its subdomains"* — because a filter whose scope you have to infer is a filter you cannot trust in a customer report |
| Matching is on the full domain label, verified against `@notverizon.com` and the suffix trick `@verizon.com.evil.test`; neither matches a `*.verizon.com` run. A wildcard anywhere but the front is escaped rather than quietly widening the filter |

### 2026-09-18 (3)
| Change |
|--------|
| `Get-DistributionGroupMembers.ps1` — **`-Recurse`**, after checking whether the report really covered everyone: it did not. Exchange only ever returns *direct* members, so a list containing another list reported that list as one member and never the people inside it. Someone who receives mail only through a nested group was invisible, and `-Member` reported "no hits" on a list that does deliver to them — a wrong answer that looks like a confident one |
| `Via groep` names the group a person came in through (empty for a direct member), and someone reachable by several routes gets one row with the routes joined rather than a row per route |
| `Aantal leden` keeps counting direct members, because that is the number Exchange and the EAC show; the new `Aantal personen` counts the real recipients reached |
| A group already expanded is not expanded again, which is also what stops a membership cycle (A contains B, B contains A) from recursing forever. Verified against a deliberately cyclic pair of test lists; nesting past 20 levels is reported and left alone |
| Documented what the report still does *not* cover: it reads group membership, so a user on no list at all appears nowhere |

### 2026-09-18 (2)
| Change |
|--------|
| `Get-DistributionGroupMembers.ps1` — `-Member` now also takes a **domain**: `-Member "@be.verizon.com"` reports every list that still holds an address on that domain (`be.verizon.com` and `*@be.verizon.com` mean the same). An address is matched by Exchange itself; a domain cannot be, so every list is read and then filtered — slower, and documented as such |
| Matching covers the primary address, every alias, **and `ExternalEmailAddress`**. That is the whole point for a partner domain: such a member is usually a mail contact whose primary SMTP is `...@contoso.onmicrosoft.com`, with the real `@be.verizon.com` only in its external address. Matching on the primary address would have found nothing and reported "none" with a straight face |
| New `Extern adres` column in the `Leden` sheet, so the address that actually receives the mail is visible for contacts instead of only the internal placeholder |
| With a filter active: `Treffers` per list in `Overzicht`, and `Treffer op` per member in `Leden`. `Treffer op` holds the matching **address**, not Ja/Nee — a hit on an alias is otherwise unexplainable in a report that does not show aliases |
| A domain filter that matches nothing says so and writes no file, rather than handing over an empty workbook that reads as a failed export |

### 2026-09-18
| Change |
|--------|
| Added `Get-DistributionGroupMembers.ps1` — every distribution list with its members in one Excel workbook: an `Overzicht` sheet (one row per list) and a `Leden` sheet (one row per member), both filterable tables with a frozen header row. Sheet headers and recipient types are in Dutch, because the workbook is what the customer reads |
| `-Member user@domain` answers "which lists is this person on?" server-side via `Get-Recipient -Filter "Members -eq '<DN>'"` instead of walking every group, and still exports the matched lists in full so the customer sees who else is on them |
| `-IncludeDynamic` and `-IncludeM365Groups` widen the report beyond plain distribution groups; dynamic groups are evaluated live, since they store no membership to query |
| Falls back to two CSV files when `ImportExcel` is missing (and offers to install it first), so a missing module never costs you the report. `ImportExcel` added to `Install-Modules.ps1` — `vias_archiver.ps1` already needed it |
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
| `Update-TeamsClient.ps1` can now remove classic Teams as well, behind `-RemoveClassicTeams` (Ninja variable `removeClassicTeams`). Off by default: taking an application away from users is not a decision an update job should make on its own. It uninstalls the *Teams Machine-Wide Installer* through msiexec — the one that matters, because while it is present Windows keeps staging classic Teams into every new profile — and per profile clears the install root, the `Run\com.squirrel.Teams.Teams` autostart entry and the stale `Uninstall\Teams` key |
| The documented per-user uninstall (`Update.exe --uninstall -s`) has to run as the profile owner, which System cannot do, so the files are removed instead. Roaming data in `%APPDATA%\Microsoft\Teams` is left alone |
| Failure handling splits the two cases on purpose: a machine-wide installer that survives the uninstall is a real failure (exit 1), while a per-profile folder that survives is almost always a file lock from a running classic Teams — a warning, cleared by the next run after the user signs out |
| Tested with detection faked, since the test device has neither variant: `-WhatIf` plans the msiexec uninstall and the folder removal and skips steps 5-8, and an applied run removed a faked profile folder with the verification reporting it clean. The msiexec path itself has **not** been run against a real Machine-Wide Installer — written down in the docs rather than implied |

### 2026-09-17
| Change |
|--------|
| Fixed a bug a real session-host run exposed: `Get-AppxPackage -AllUsers` finds nothing when Teams is only *provisioned* and no user has it yet, so the version check had nothing to compare, declared the host outdated and reinstalled ~275 MB on every scheduled run. The installed version now falls back to the provisioned package version. Verified against that exact scenario: reports `Provisioned MSTeams <version>`, compares, does nothing |
| `Update-TeamsClient.ps1` now checks whether **Outlook itself** sees the meeting add-in, not only that the MSI installed. It reads `HKEY_USERS\<sid>\...\Outlook\Addins\TeamsAddin.FastConnect` per signed-in user plus the machine-wide key: `LoadBehavior 3` = loaded, `2`/`0` = Outlook switched it off (the real "the button is gone" case). Reported in preflight and verification, never as a failure — a profile nobody is signed into cannot be read |
| Preflight became a full inventory of everywhere Teams can live: AppX per user, the provisioned package, the classic *Teams Machine-Wide Installer*, classic per-profile installs, the add-in in both hives, the Outlook registration. Classic Teams is reported, not removed — it shares the October 2026 end-of-support date and a leftover machine-wide installer keeps restaging it into new profiles |
| Two bugs found by running it rather than reading it. Enumerating profiles with an `S-1-5-21-*` whitelist skips **every** user on an Entra-joined device, where SIDs are `S-1-12-1-*` — the check claimed nobody had the add-in registered while `LoadBehavior=3` sat right there. And `New-PSDrive` honours `ShouldProcess`, so under `-WhatIf` the `HKEY_USERS` drive was never created and the same read-only check lied; the hives are addressed through `Registry::HKEY_USERS` now, with no drive to create |

### 2026-09-11 (5)
| Change |
|--------|
| Added `scripts/Exchange/Move-SharedCalendar.ps1` — all in one: `-Search balie` finds the calendar, shows who uses it, moves it into a resource mailbox with `Convert-SharedCalendarToResource.ps1` and lists who has to switch. One temporary App Registration with the permissions of both scripts, created once and removed at the end, so there is one sign-in instead of two. Several matches are picked from a list or narrowed with `-Owner`; a non-interactive run lists them and stops rather than guessing. The two scripts are called, not copied, so there is one implementation of each step |
| Fixed `Convert-SharedCalendarToResource.ps1`: in a non-interactive session the typed confirmation was skipped and the original calendar removed without `-Force`. It is now left in place with a warning unless `-Force` is given |
| Both calendar scripts: an explicit `-ClientId` now takes precedence over an existing app-only Graph session, so a calling script's app is really used. `Convert-SharedCalendarToResource.ps1` gains `-PassThru` (result object for callers) |
| Exchange submenu (`C`) option `J` added |

### 2026-09-11 (4)
| Change |
|--------|
| `Get-CalendarMappings.ps1` — first real run (a tenant where "Balie planning" turned out to be a secondary calendar in an archived mailbox) confirmed the Graph assumptions: app-only reads return the calendars users added, and a shared secondary calendar appears in their list under its own name. It also exposed a wrong hint: a `NotMapped` row pointed at "Reservering vergaderzaal LBM" as "probably this one", while that is a *second* calendar the same owner shares. The hint now skips entries named after another calendar of the owner, and - for a secondary calendar - entries named after the owner, which are their main calendar |

### 2026-09-11 (3)
| Change |
|--------|
| Added `scripts/Exchange/Convert-SharedCalendarToResource.ps1` — moves a shared calendar (the "Balie" calendar in one person's mailbox) into a Room or Equipment mailbox of its own, with every item and every permission, then removes the original on request. Preview by default; `-Apply` creates and copies, `-RemoveSourceCalendar` removes the original only after every item has a verified copy and the calendar's name has been typed as confirmation |
| Items are copied faithfully rather than approximately: recurring series stay series with their moved and cancelled occurrences applied (matched occurrence by occurrence, and left alone with a warning if the two series do not line up), times are written back in the time zone they were created in so weekly items survive a daylight saving switch, categories keep their colour, attachments up to 3 MB are copied and larger ones saved to the backup folder. Attendees are listed in the body instead of copied, so nobody receives a fresh invitation |
| Permissions carry their exact Exchange access rights, custom rights included; `-SendSharingInvitation` sends users the standard invitation. External people, deleted accounts and delegate flags are reported, not silently dropped |
| Every copy carries its source item's id in a hidden property, so a run that stops halfway continues where it left off; a half-finished series is redone. A JSON backup of everything read is written before anything is created |
| Exchange submenu (`C`) option `I` added: always a preview first, then an explicit second step |

### 2026-09-11 (2)
| Change |
|--------|
| `Get-CalendarMappings.ps1` gains `-Search` (alias `-Keyword`): "where is the Balie calendar?" in one run. The keyword is matched against the owner's name and every address (a shared mailbox `balie@`, a room, a group) and against calendar names (a secondary calendar *Balie* in somebody's mailbox). The report shows where the calendar lives (new status `Source`), who has it in their calendar list, and who has rights on it |
| A matching **secondary** calendar now gets its own permissions read, instead of being compared against the owner's main calendar and landing on `MappedWithoutRight`. A new `Calendar` column says which of the owner's calendars a row is about |
| A calendar list entry carries no link back to the folder it came from, so a shared secondary calendar is matched by name. When a user has it under another name, the `NotMapped` row names the entry that is probably it rather than leaving a silent false negative |
| Menu option `H` asks for a keyword first; blank falls back to the full or per-mailbox report |

### 2026-09-11
| Change |
|--------|
| Added `scripts/Exchange/Get-CalendarMappings.ps1` — shows where each calendar is actually mapped: for every mailbox it reads the calendar list in Outlook and the rights on its own main calendar, and folds both into one row per owner + user with a status (`Mapped`, `MappedWithoutRight`, `NotMapped`, `MappedOwnerMissing`, `SharedExternally`, …). `Test-CalendarPermissions.ps1` says who *may* open a calendar; this says where it *is*, and where the two disagree |
| Graph rather than Exchange Online PowerShell, because the entries a user added to their own calendar list are not visible to any Exchange cmdlet. App-only access follows the same three routes as `Remove-PhishingMessage.ps1` (existing session, own app, or a temporary app that is removed in a `finally`), with read-only permissions `Calendars.Read`, `User.Read.All` and `Group.Read.All`. No Exchange connection, so no MSAL clash |
| Tenant-wide runs go through `$batch` (20 mailboxes per call) with throttled items retried. A mailbox that cannot be read is reported as such rather than as "nothing mapped" |
| Written down what the report cannot see: Full Access with AutoMapping (a mailbox permission — `Test-MailboxPermissions.ps1`), calendars opened in classic Outlook without shared calendar improvements, and secondary calendars, which show up as `MappedWithoutRight` |
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
| Added `scripts/SharePoint/Provisioning/Add-SharePointHelpPage.ps1` — puts the end-user explanation on the team site as a SharePoint page, linked from the left-hand navigation. A handleiding in a repo is read by nobody; this writes it where the people who upload files already are |
| The page is generated from the configuration rather than typed out, so it cannot drift from what the libraries actually do: the channels it lists are the ones that exist, the labels carry the same help text that appears under each field in the upload form, and the required fields per document type are read off the content types |
| Written for the person uploading a catalogue. Two pieces of the configuration are deliberately kept off it: the `note` on a container, which names security groups, and the `description` on a view, which talks about pillars and synced folders. Permissions are left out entirely — who may see what is not something a user can act on |
| Fixed along the way, found by rendering the page rather than reading the code: it announced three ways of adding a file and listed two, and the wizard was writing the team name where the company name belonged ("Intern blijft binnen Laseto-NewTeams") |
### 2026-09-16 (3)
| Change |
|--------|
| Added `scripts/SharePoint/Provisioning/Remove-SharePointStructure.ps1` — takes the same configuration apart, deepest first: tabs, channels, libraries, content types (unbound from their lists first), site columns, term set, security groups, and the team itself |
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
| Added `scripts/SharePoint/Provisioning/Sync-SharePointChannelMember.ps1` — makes an Entra ID security group the source of truth for who is in a private Teams channel. A private channel cannot be given rights through a group at all: Teams tracks its roster one person at a time and Graph accepts only individual users there, so the group feeds the roster instead |
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
| `New-StructureConfig.ps1` asks what everything should be called and writes the configuration itself — nobody should have to open a JSON file to name a channel. Enter accepts the suggestion in brackets, so a standard build is mostly Enters plus the tenant and the team owner |
| Everything else is derived from those answers: per pillar a channel, a content type, two security groups and a grouped view; per brand a cross-cutting view spanning every pillar folder. Which pillar handles suppliers and which handles sales is what decides where Leverancier and Regio become required fields |
| Two of the answers are the ones that cost something later, so they are asked last and default to no: maintaining the share-status column (the only nightly script) and enforcing per-pillar rights on standard-channel folders (the part Microsoft does not support) |
| `New-SharePointTeam.ps1` creates the Microsoft 365 team and its channels, the private MGMT one included, so the structure can be built from an empty tenant. A private channel's site collection is provisioned asynchronously and its URL cannot be known in advance — the script polls for it and writes it back into the configuration, which is what lets the following steps connect to something |
| Never renames or deletes a channel: a channel whose name does not match the config is reported, not corrected, because renaming one moves its folder and breaks every link anyone has shared |
| `Install-SharePointStructure.ps1` runs the wizard by itself when it finds no configuration for the tenant, and the team step is now step 1 of six. `-SkipTeam` for a team that already exists |
| Column internal names and content type IDs are generated once and then fixed — SharePoint keys document metadata to both — which is why the wizard refuses to overwrite an existing configuration without `-Force`. Display names, channel names and group names stay changeable |
| Removed the last hardcoded column names: the share-status audit reads which column is which from a new `fieldRoles` section instead of assuming `PsDeelstatus` and `PsVertrouwelijkheid` |
| The app registration now also consents `Channel.Create`, `ChannelSettings.ReadWrite.All` and `Team.Create`, which the team step needs |
### 2026-09-15
| Change |
|--------|
| Added `scripts/SharePoint/Provisioning/Install-SharePointStructure.ps1` — builds the whole structure in one run: registers the Entra app itself and admin-consents its delegated scopes, then metadata → libraries/groups/permissions → verification, optionally the first deelstatus audit. Each step stays its own script, so a failure is rerun on its own instead of starting over |
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
| Added `scripts/SharePoint/Provisioning/` — provision and maintain a whole SharePoint structure (metadata model, content types, libraries, Entra ID group permissions) for an MSP client from one JSON config, with a sharing audit and a read-only drift check. Built for Petsolutions NV (brands Butterstone/Laseto), but nothing in the scripts is client-specific |
| The model lives in `petsolutions.config.json`, cross-checked at load time: a content type referring to an undefined column, or a container granting a group that is not in the model, fails before anything connects rather than halfway through provisioning. The shipped `CHANGEME` tenant/site URLs are refused outright |
| Column internal names carry a `Ps` prefix. "Contenttype" and "Status" are display names SharePoint already uses for something else, and the prefix keeps them unambiguous in CAML, in views and in the drift check while users still see plain Dutch labels. Content type IDs are fixed rather than generated, so the same structure is reproducible across tenants |
| `New-SharePointMetadata.ps1` runs against **every** site in the config, not just the team site: a Teams private channel (MGMT here) is its own site collection and a site column does not reach across one. Making a column required after the fact works — the `Required` flag on an existing field link is updated in place and pushed down to the lists already using the content type |
| `Set-SharePointLibraries.ps1` sets a per-folder content type order, so the *New* menu inside the Leveranciers channel offers Leveranciersdocument and not the five types belonging to the other pillars — the shared library has to carry them all, the folder does not have to show them. No view is ever made the default: the default view of a Teams library is what every member of the channel sees the second they open Files |
| Written down rather than hidden: unique permissions on a **standard**-channel folder are what this model asks for and what Microsoft does not support. Members who lose access keep seeing the channel in Teams and get an error on the Files tab instead of a closed door. The script does it, warns per folder, and `-SkipChannelFolderPermissions` leaves those folders inheriting. A private channel, a shared channel or an own library (what FUTECH uses) are the supported ways to close a pillar off |
| `Update-SharePointShareStatus.ps1` derives the Deelstatus column from the permissions actually on each file. It asks the cheap question first — a file that inherits is not shared — so one round trip per hundred items settles nearly the whole library; only files that broke inheritance get their role assignments read, and of those only specific-people links need expanding (an Anyone or Organization link already says in its name whether a guest can be behind it). Writes with `SystemUpdate` so Modified/Modified By stay put and no version is created |
| The audit never revokes a link. It reports files tagged Intern or Vertrouwelijk sitting behind an external one and exits `2`, so a scheduled RMM job surfaces exactly when there is a decision for a person to make. `Test-SharePointStructure.ps1` does the same for structural drift, classified as Missing / Different / Extra — "Extra" is never fixed automatically, because an extra column holds data and an extra role assignment is usually somebody's deliberate exception |
| `SharePointStructure.Common.ps1` is dot-sourced by all four — a deliberate exception to the "every script stands alone" rule elsewhere in this repo, because they share one config schema and three copies of the permission code would drift apart within a month |
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
| `scripts/Device/Update-TeamsClient.ps1` absorbs the AVD/VDI parts of the older gap-fill installer behind `-AvdOptimizations`: the `IsWVDEnvironment` media flag (set in step 3, before the client is provisioned, because Teams reads it at startup to pick its media path) and the Remote Desktop WebRTC Redirector Service from `aka.ms/msrdcwebrtcsvc/msi` |
| Deliberately a switch and not autodetection: setting that flag on a normal endpoint tells Teams to hand media to a redirector that is not there. Without the switch the script only *reports* that a device looks like a session host (`HKLM:\SOFTWARE\Microsoft\RDInfraAgent`) |
| Both components are installed only when missing (`-Force` reinstalls the redirector), so a scheduled run on a configured session host still downloads nothing and prints nothing under `-Quiet`. Verified end to end, including that the redirector MSI (1.7 MB, `1.54.2408.19001`) passes the Microsoft signature check |
| The add-in step now also skips itself when the add-in is present and the client was not replaced — before this it would reinstall the add-in on a run that was only there to fix the AVD components |
| Download and signature verification moved into one `Save-VerifiedDownload` helper shared by the bootstrapper and the redirector: https-only, minimum size, Authenticode `Valid` and signed by `O=Microsoft Corporation`, or it throws |
| Corrected in the docs: Ninja script-variable names are **not** case-sensitive. Windows environment lookups are case-insensitive, so variables named `Quiet` or `Force` work exactly like `quiet` and `force` |

### 2026-09-10
| Change |
|--------|
| `scripts/Device/Update-TeamsClient-ITGlue.md` gained a NinjaOne setup appendix: which values to pick per field when adding the script (PowerShell 5.1 rather than 7, 64-bit, Run As System), the script-variable names with a way to verify they actually arrive, the test run on one device, the scheduled automation with `-Quiet -Confirm:$false`, and the optional detection job |
| Written down explicitly because it will otherwise be reported as a bug: `-CheckOnly` exits `2` when an update is available, and NinjaOne shows every non-zero exit code as a failed job. That is the intent — those are the devices needing attention — and it is what a script result condition can key on |
| Also flagged: the Ninja script timeout must exceed `-TimeoutSeconds` (900 s) plus the ~275 MB download, otherwise Ninja kills the job mid-install |

### 2026-09-08 (5)
| Change |
|--------|
| Added `scripts/Device/Update-TeamsClient-ITGlue.md` — the service desk version of that documentation, in Dutch, to paste into IT Glue. Layered per support level: L1 checks with `-CheckOnly -Quiet` and reads the labelled output, L2 runs the update from NinjaOne or by hand and verifies afterwards, L3 gets parameters, exit codes, paths and the built-in safeties. Includes an error table with the escalation level per message, an FAQ, and ready-made text for the end user |
| It leads with the point that trips people up: no output means the device is already current, which is a successful run and not a failure. Download volume per device (~275 MB: a 1.9 MB bootstrapper that pulls a ~273 MB package) was measured, not estimated |

### 2026-09-08 (4)
| Change |
|--------|
| Added `scripts/Device/Update-TeamsClient.md` — a reference for that script: the decision tree (behind → full reinstall, current but add-in missing → add-in only, current → nothing at all), the seven steps, the config-service version check with a sample response, output modes and exit codes, the NinjaOne script-variable table, the design decisions behind the order of operations, a troubleshooting table and what has actually been tested. Linked from the Device readme |

### 2026-09-08 (3)
| Change |
|--------|
| `scripts/Device/Update-TeamsClient.ps1` no longer reinstalls unconditionally: it asks the Teams client config service (`config.teams.microsoft.com/config/v1/MicrosoftTeams/...`, `BuildSettings.WebView2PreAuth.<arch>.latestVersion` — the same feed the client uses to decide it is out of date) which build is published for this architecture, and leaves an up-to-date device completely alone |
| Is the client current but the meeting add-in missing? Then only the add-in is installed — no download, no uninstall, no reprovision |
| `-Quiet` holds back all output until there is news, so a scheduled NinjaOne run prints nothing on an up-to-date device and only surfaces in the activity feed when it found a newer build or hit a problem. Verified: an up-to-date `-Quiet` run produces zero bytes of output and exit code 0 |
| `-CheckOnly` reports without changing anything and exits `2` when a newer build is available, for use as a Ninja detection/condition job. `-Ring` selects a non-default update ring |
| When the config service cannot be reached the run stops instead of reinstalling blindly; `-Force` now means "reinstall even though it is current" as well as "continue without Teams or version info" |
| A transcript is only written when the run actually changes something, so an hourly check leaves no log litter in `C:\Temp` |

### 2026-09-08 (2)
| Change |
|--------|
| `scripts/Device/Update-TeamsClient.ps1` is now safe to run unattended from an RMM (NinjaOne) *and* by hand. It relaunches itself 64-bit via `SysNative` when the agent starts PowerShell 32-bit — otherwise the HKLM reads are redirected to `WOW6432Node` and `$env:ProgramFiles` points at the x86 folder, so neither the AppX package nor the add-in MSI is ever found |
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
| Added `scripts/Device/Update-TeamsClient.ps1` — clean reinstall of new Teams on an endpoint or AVD session host: uninstall the Teams Meeting Add-in, remove the `MSTeams` AppX package for all users, download `teamsbootstrapper.exe`, provision Teams (`-p`) and install the meeting add-in MSI that ships inside the new Teams package |
| Every state-changing step runs through `ShouldProcess`, so `-WhatIf` walks the whole flow and prints each uninstall/download/install without touching the machine; the steps that only exist after a real install (new Teams version, add-in MSI path, final verification) are reported as such instead of failing the run |
| The add-in lookup reads both the 64-bit and the `WOW6432Node` uninstall hive — the add-in installs 32-bit, so the 64-bit hive alone never finds it (uninstall and verification both missed it before) |
| Exit codes and msiexec/bootstrapper exit codes are checked instead of assumed; `-SkipMeetingAddIn` replaces only the client, `-Force` installs on a device without any Teams. Wired into `menu.ps1` (key T), which defaults to a `-WhatIf` preview |

### 2026-09-07 (2)
| Change |
|--------|
| Added `scripts/SharePoint/Search-SharePointContent.ps1` — the Microsoft Graph counterpart of `Find-SiteContent.ps1`: app-only, no interactive sign-in, and it searches one site or **every site and OneDrive in the tenant** |
| Finding content uses `/drives/{id}/root/delta` (a whole library tree in pages of a thousand items, so `*contains*` wildcards work) or `/search/query` with `-Content` for text inside documents; the filters are the same as in the PnP script |
| Permissions come from `/drives/{id}/items/{id}/permissions`, 20 per `/$batch` call. One call yields the roles, the granted-to identities, the sharing link with its scope (anyone/organization/specific people), edit-or-view, expiry and URL, and `inheritedFrom` — which is what decides `PermissionSource = Item` (unique) versus `Inherited`. "Anyone with the link" gets its own counter because those need no sign-in at all |
| Written down explicitly, in the script and the readme: Graph has no API for SharePoint role assignments, so site- and list-level rights and items in ordinary (non-library) lists stay the domain of `Find-SiteContent.ps1`. The readme has a comparison table for picking between the two |
| Sign-in is app-only: the first run registers an app, consents the application role `Sites.Read.All`, creates a self-signed certificate in `CurrentUser\My` and uploads its public key — no secret on disk — and caches client ID plus thumbprint per tenant in `graph.appid.json` (added to `.gitignore`). Later runs connect without a prompt, so it also works from a scheduled task. Throttling (429) is retried, honouring `Retry-After`, for single calls and batch sub-requests alike |

### 2026-09-07
| Change |
|--------|
| Added `scripts/SharePoint/Find-SiteContent.ps1` — search an entire SharePoint site or OneDrive for content and report which permissions apply to every hit. Read-only |
| Two engines: a crawl over every list and library (sees everything, `-IncludeSubsites` for the subsites) and a KQL query against the search index (`-Content`) that also matches text *inside* documents. Both share the filters `-Name`, `-Path`, `-Extension`, `-ItemType`, `-ListName`, `-ModifiedBy`, `-ModifiedAfter`/`-ModifiedBefore` and `-MinSizeMB` |
| Per hit the script resolves where the permissions come from — the item itself (broken inheritance), its list, or the site — and flattens the role assignments to one CSV row per principal with type, login, e-mail and role names. `Limited Access` is filtered out unless `-IncludeLimitedAccess` |
| Sharing links (the `SharingLinks.*` groups behind "Copy link") are always expanded to the people in them and labelled Anyone/Organization/Specific people; external guests (`#ext#`) and "Everyone (except external users)" are flagged separately in the summary and the CSV |
| Site and list permissions are read once and cached and item permissions only for items that broke inheritance, so cost scales with the number of hits, not the size of the site; `-Permissions Unique` reports only what is shared differently, `-Permissions None` skips permissions, and `-MaxPermissionLookups` caps a too-broad search |
| Reuses the app-registration flow and the per-tenant `pnp.appid.json` cache of `Restore-RecycleBinItems.ps1`, and can temporarily grant itself site collection admin (`-GrantSiteAdmin`) to search a site or OneDrive it has no rights on |

### 2026-08-28
| Change |
|--------|
| Added `scripts/SharePoint/` with `Restore-RecycleBinItems.ps1` — restore deleted files/folders from a SharePoint site or OneDrive recycle bin, dry-run by default, with filters on name, original folder, who deleted it, and a deletion time window |
| Two scopes: `-SiteUrl` for one site collection (OneDrive included), or `-AllSites -TenantUrl` to walk every SharePoint site in the tenant. The tenant sweep excludes OneDrive personal sites, the My Site host, redirect sites and locked sites, supports `-SiteFilter`/`-MaxSites`, and keeps going when a single site errors out — per-site results land in a summary table and in a `Site` column in the CSV |
| Restores run in batches of up to 200 items through `Restore-PnPRecycleBinItem -IdList` (one server call per batch) instead of one call per item; folders and files never share a batch, and a batch that fails as a whole is retried item by item so per-item errors are still reported. Parallel runspaces were deliberately not used — PnP PowerShell is not thread-safe and concurrent calls against one site collection hit SharePoint throttling |
| The script reports timing at every step — how long reading the recycle bin took, an up-front estimate of the restore, a progress bar with a live ETA from the measured rate, the real duration in the summary, and a `DurationSeconds` column per item in the CSV |
| The script registers its own Entra app on the first run against a tenant (public client, delegated `AllSites.FullControl`, admin-consented) because PnP PowerShell no longer ships a shared multi-tenant app; the client ID is cached per tenant in `pnp.appid.json` (added to `.gitignore`) |

### 2026-07-24 (3)
| Change |
|--------|
| Retired a now-unmaintained internal PowerShell repo (`Windows-Powershell`, last commit March 2023) by reviewing every script in it and modernizing whatever still had value into this repo — nothing was copied verbatim; everything was rewritten against Microsoft Graph / Exchange Online (the source repo's `MSOnline`/`AzureAD`-based scripts are fully non-functional since Microsoft retired those endpoints) |
| Added `scripts/TenantOnboarding/` (24 scripts across Provisioning/MultiTenant/AppDeployment/DeviceConfig/OneDriveManagement/UserManagement) — modernized from the source repo's tenant-setup/onboarding scripts |
| Added `scripts/Office365Toolkit/` (9 scripts across Security/Exchange/Intune) — modernized from a forked copy of the retired `directorcia/Office365` (CIAOPS) GitHub project found in the source repo; reviewed capability-by-capability and consolidated, not ported 1:1 |
| Added `scripts/PatronToolkit/` (13 scripts across Entra/Security/Exchange/Intune/SharePoint/Teams) — modernized from a forked copy of the retired `directorcia/patron` GitHub project found in the source repo, same capability-consolidation approach |
| Added `scripts/LegacyUtilities/` (17 scripts across Exchange/Entra/Teams/Network/Device/Workspace365) — modernized from assorted small tools in the source repo not covered by the above |
| Data-handling boundary applied throughout: the source repo's `Klanten`, `created-users`, `csv files`, and `Archief` folders (real customer names/tenant domains/generated passwords) were never read or ported; any other script found to hardcode real customer/tenant identifiers or secrets was generalized into parameters instead, or skipped outright — see each new folder's readme for its specific skip list |
| None of the ~63 new scripts are wired into `menu.ps1` — they're audit/reporting/setup scripts meant to be run directly, matching the existing pattern for `scripts/RDS/`, `scripts/Azure/`, and `scripts/Network/UniFi/` |

### 2026-07-24 (2)
| Change |
|--------|
| Fixed `scripts/Reporting/Get-SharePointStorageReport.ps1` silently abandoning version-history lookups on very large libraries: `Invoke-GraphBatchGet`'s retry-pass ceiling was hardcoded at 8, but SharePoint Online's per-app activity throttle allows only ~1500-2500 resolved version lookups per pass before a ~60-90s cool-down repeats — on a 200k-file tenant this meant ~90% of files got marked "gave up" before the scan actually finished |
| Applied the identical fix to `scripts/Reporting/Remove-SharePointFileVersionsByDate.ps1`'s own copy of the same batch-retry function (`Get-FileVersionsBatch`), which had the same hardcoded 8-pass ceiling |
| Added `-MaxVersionRetryPasses` parameter to both scripts (default `0` = auto-scales the pass ceiling to the request volume, capped at 500 passes); a clear `Write-Warning` is now emitted listing exactly how many files were abandoned and suggesting the parameter if the ceiling is still hit |
| Updated `scripts/Reporting/readme.md` to document `-VersionBatchConcurrency` and `-MaxVersionRetryPasses` for both scripts |

### 2026-07-24 (1)
| Change |
|--------|
| Documented `scripts/Azure/VM/Azure-NVMe-Conversion.ps1` (previously untracked in any readme/structure tree): added `scripts/Azure/readme.md` and `scripts/Azure/VM/readme.md`, and a new "Azure Infrastructure" category in this readme |
| Wired 5 previously menu-less but fully-documented scripts into `menu.ps1`: `Move-InboxToArchive.ps1` and `Set-Distributionlist-dynamic-static.ps1` (Exchange submenu, keys E/F), `Get-M365UserLicenses.ps1`, `Import-ConditionalAccessBaseline.ps1`, `Set-UserManager.ps1` (Entra submenu, keys G/H/I) — and added them to the Menu tables in this readme |
| Added `scripts/Device/Remove-OemBloatware.ps1` — detects HP/Lenovo/Dell and removes OEM bloatware via winget plus generic Microsoft Store junk via AppX; dry-run by default; wired into `menu.ps1` (key I) |
| Added `scripts/Device/DriveMapping/New-CloudDriveMapping.ps1` — maps SharePoint/OneDrive document libraries to drive letters via WebDAV for use as a logon script; dry-run by default |
| Added `scripts/Network/UniFi/` — `UnifiApi.ps1` shared login helper (classic controller + UniFi OS auto-detect), `Get-UnifiNetworkReport.ps1` (HTML documentation report), `Update-UnifiFirmware.ps1` (dry-run firmware upgrade tooling); credentials always via `Get-Credential`, never hardcoded |
| Added `scripts/Intune/Compare-IntuneConfig.ps1` — Intune configuration drift detection between a customer tenant and an MSP baseline backup, via the `IntuneBackupAndRestore` module; read-only |

### 2026-07-22 (6)
| Change |
|--------|
| Updated `scripts/Reporting/Get-SharePointStorageReport.ps1` to make full-site scans GDAP-proof by resolving and pinning one effective tenant context (`-TenantId`, or GDAP customer context from `$global:cid`) for Graph sign-in, temporary app creation, and app-only token issuance |
| Added guard rails for GDAP/app-only flows: clearer errors when customer tenant context is missing (run `Connect-Tenant` first or pass `-TenantId`) and when `-ClientId` is provided without a resolvable tenant ID |
| Updated checkpoint signature inputs in `Get-SharePointStorageReport.ps1` to include `-ForceAppOnlySingleSite` and resolved tenant context, preventing cross-context resume collisions |
| Updated `scripts/Reporting/readme.md` to document full-site GDAP tenant binding behavior |

### 2026-07-22 (5)
| Change |
|--------|
| Updated `scripts/Reporting/Get-SharePointStorageReport.ps1` to make single-site scans GDAP-proof: when `authMode=GDAP` is detected, the script now automatically uses the temporary app/app-only bootstrap path for `-SiteUrl` scans to avoid delegated permission gaps |
| Added `-ForceAppOnlySingleSite` parameter to explicitly force app-only bootstrap for single-site scans, independent of detected auth mode |
| Updated `scripts/Reporting/readme.md` to document the new GDAP behavior and `-ForceAppOnlySingleSite` parameter |

### 2026-07-22 (4)
| Change |
|--------|
| Hardened `scripts/Reporting/Get-SharePointStorageReport.ps1` delegated path to remove dependency on missing Graph cmdlets: replaced `Get-MgDriveItemChild` usage with Graph REST pagination via `Invoke-MgGraphRequest` for drive children traversal |
| Updated site-drive enumeration fallbacks in `Get-SharePointStorageReport.ps1` to use Graph REST (`/sites/{id}/drives`) instead of `Get-MgSiteDrive` in delegated/non-app-only branches |
| Updated recycle-bin retrieval in `Get-SharePointStorageReport.ps1` to use Graph REST pagination (`/sites/{id}/recycleBin/items`) as fallback/primary delegated path instead of `Get-MgSiteRecycleBinItem` cmdlet dependency |
| Suppressed non-fatal MSAL authority warning noise on disconnect by wrapping `Disconnect-MgGraph` with temporary warning suppression in cleanup |

### 2026-07-22 (3)
| Change |
|--------|
| Updated `scripts/Reporting/Get-SharePointStorageReport.ps1` scan-mode handling so single-site runs (`-SiteUrl` with `/sites/...` or `/teams/...`) no longer trigger temporary app registration + app-only bootstrap; app-only setup is now only used for tenant-wide enumeration |
| Improved single-site lookup performance and reliability by resolving the exact site directly via Graph URL path (`/sites/{hostname}:{path}`) instead of search/filter flow |
| Updated `scripts/Reporting/readme.md` performance notes to document the single-site optimized path and expected startup speed behavior |

### 2026-07-22 (2)
| Change |
|--------|
| Fixed SharePoint report dependency issue causing `Get-MgSite` command-not-found errors: updated `load.ps1`, `scripts/Startup/Install-Modules.ps1`, and `scripts/Startup/Update-Modules.ps1` to include `Microsoft.Graph.Sites` |
| Updated `scripts/Reporting/Get-SharePointStorageReport.ps1` with an explicit module preflight check for `Microsoft.Graph.Authentication` and `Microsoft.Graph.Sites`, including a clear install hint when modules are missing |
| Updated `scripts/Startup/readme.md` module dependency documentation to include `Microsoft.Graph.Sites` in install/update requirements |

### 2026-07-22 (1)
| Change |
|--------|
| Updated `load.ps1` — added startup launcher switches `-SetupStartup` / `-RemoveStartup`; first-run config now stores delegated auth defaults (`authMode`, optional `defaultCustomerDomain`, `useDeviceCodeAuth`) in `load.config.ps1` |
| Updated `menu.ps1` — added Startup actions `F` (Enable-LauncherStartup) and `G` (Disable-LauncherStartup); added M365 action `H` (Test-GdapConnection); Entra submenu now includes temporary CA and TAP actions (`D`/`E`/`F`) |
| Updated `scripts/Startup/functies.ps1` — Graph startup connection now supports delegated device-auth preference + required scopes for GDAP flow; added `Test-GdapConnection` helper for delegated contract/connectivity checks |
| Updated startup module maintenance: `scripts/Startup/Install-Modules.ps1` and `scripts/Startup/Update-Modules.ps1` now include `Microsoft.Graph.Identity.DirectoryManagement` |
| Added `scripts/Entra/New-TemporaryConditionalAccessPolicy.ps1` — create temporary CA policy for user/group with either duration-based window or exact local start/end datetime; optional same-session auto-cleanup at end time |
| Added `scripts/Entra/Remove-TemporaryConditionalAccessPolicies.ps1` — remove one or multiple temporary CA policies (`TEMP-CA -`), including expired-only or remove-all modes |
| Added `scripts/Entra/New-UserTemporaryAccessPass.ps1` — create Temporary Access Pass (TAP) for a user with configurable lifetime and one-time option |
| Updated docs for the above across `readme.md`, `scripts/readme.md`, `scripts/Startup/readme.md`, and `scripts/Entra/readme.md`; clarified that temporary CA auto-cleanup runs in the current session (no Scheduled Task created) |

### 2026-07-09 (8)
| Change |
|--------|
| `scripts/Reporting/Get-SharePointStorageReport.ps1` — added a "site collection totals" phase (`-Apply` only, Phase 2c): sub-sites/Teams-kanalen and the recycle bin are now automatically rolled up per root site collection into `SharePoint_SiteCollectionTotals_<timestamp>.csv`, so the grand total is directly comparable to the single "storage used" figure shown per site in the SharePoint admin center |
| Documented the new output and the most likely causes of a remaining mismatch with the admin portal figure (timing lag, silently skipped folders on permission errors, failed version lookups) in `scripts/Reporting/readme.md` |

### 2026-07-09 (7)
| Change |
|--------|
| `scripts/Reporting/Get-SharePointStorageReport.ps1` — briefly extended recycle bin lookups (`-Apply`'s Phase 2b and `-RecycleBinOnly`) to also cover OneDrive personal sites, then reverted the same day on request — recycle bin scope stays SharePoint site collections only, OneDrive stays fully excluded (both storage scan and recycle bin) |
| Added a "Prullenbak (recycle bin)" section to `scripts/Reporting/readme.md` documenting the (SharePoint-only) recycle bin scope |

### 2026-07-09 (6)
| Change |
|--------|
| Removed the `Testing Scripts/` wrapper folder entirely — its subfolders duplicated existing top-level category names by verb (`Test-`/`Get-` prefix) rather than by domain. Merged its contents into the matching domain folder: `Testing Scripts/Entra/Test-M365GroupMembership.ps1` → `Entra/`, `Testing Scripts/Exchange/*` (6 scripts) → `Exchange/`, `Testing Scripts/Device/Test-OpenVpnDiagnostics.ps1` → `Device/`. `Network/`, `RDS/`, and `SMTP/` (no existing top-level counterpart) were promoted to their own top-level category folders instead |
| Merged the corresponding readmes into each destination folder's existing readme.md rather than keeping separate "Testing —" docs |
| Updated 10 `menu.ps1` script paths (Exchange audit submenu, Entra audit submenu, Test-Ports, SMTP tests) for the new locations |

### 2026-07-09 (5)
| Change |
|--------|
| Tidied `Testing Scripts/` and the repo root: moved `vias_archiver.ps1` out of `Testing Scripts/Device/` into a new `scripts/Teams/` category — it's a Teams/SharePoint export & archiving tool, not a diagnostic script, so it didn't belong under "Testing" |
| Moved root-level `Update-modules.ps1` into `scripts/Startup/` (renamed `Update-Modules.ps1` for naming consistency) — it's a module-maintenance script like `Install-Modules.ps1`, not a repo entry point like `load.ps1`/`menu.ps1` |
| Removed the `Testing Scripts/SharePoint/` folder (it held only a pointer readme, no script) — that pointer now lives directly in `Testing Scripts/readme.md` |
| Documented `Test-PowerShellSyntax.ps1` in `scripts/Startup/readme.md`, which had no docs before |

### 2026-07-09 (4)
| Change |
|--------|
| Optimized `scripts/Reporting/Get-SharePointStorageReport.ps1` version-history lookups, which were the main cause of the script appearing to hang on large libraries (one sequential Graph call per file, each eligible for up to 6 retries with backoff up to ~2 minutes on throttling): (1) skip the lookup entirely when a library is positively known to have versioning disabled, (2) batch up to 20 file version lookups per HTTP call via Graph's `$batch` endpoint instead of one call per file, (3) use a short, cheap 3-attempt retry for these specific calls instead of the main retry/backoff policy, since a failed lookup safely falls back to "0 versions" |
| Removed the now-unused `Get-VersionSize` function, replaced by `Invoke-GraphBatchGet` + batched resolution in `Get-AllDriveItems` |
| Updated `scripts/Reporting/readme.md` with a "Performance" section documenting the above |

### 2026-07-09 (3)
| Change |
|--------|
| Moved `Deploy-OfficeTheme.ps1`, `Deploy-Officecolors.ps1`, and their theme assets (`2026 Vias institute colours (2).thmx`, `Office Themes/`) back to `Custom Scripts/Intune/Desktop/` — both scripts hardcode their download URL to that exact repo path, so this keeps the URL valid instead of requiring a script update + Intune redeploy. Recreated `Custom Scripts/` and `Custom Scripts/Intune/` as minimal path-pinned folders (just this one item) rather than the full former category |
| `scripts/Intune/Desktop/` now holds only wallpaper/lockscreen/taskbar-shortcut deployment; both `Intune/readme.md` and `Intune/Desktop/readme.md` cross-reference `Custom Scripts/Intune/Desktop/` for the Office theme scripts |

### 2026-07-09 (2)
| Change |
|--------|
| Removed the `Custom Scripts/` wrapper folder — it mixed generic tooling with customer-specific scripts under one confusing label, and duplicated the `Intune/` category. Contents redistributed to proper top-level categories: `Custom Scripts/device/` → `Device/`, `Custom Scripts/DNS/` → `DNS/`, `Custom Scripts/SAS/` → `SAS/`, `Custom Scripts/Save install time/` → `Deployment/` (renamed), `Custom Scripts/Intune/Desktop/` → merged into `Intune/Desktop/` |
| Updated `menu.ps1` script paths for `Restart-Time-Sync.ps1`, `detect-audiodevices.ps1`, `Disable-internalmic.ps1` to their new `scripts/Device/` location |
| Updated cross-references in `scripts/Intune/readme.md`, `scripts/Intune/Get-Autopilot/readme.md`, and `scripts/readme.md` for the new folder locations |
| ~~**Known issue (intentional):** `Deploy-OfficeTheme.ps1` and `Deploy-Officecolors.ps1` still hardcode their download URL to the old path — left unchanged on request.~~ **Resolved above** — the scripts moved back to match their hardcoded URL instead. |

### 2026-07-09
| Change |
|--------|
| Added `readme.md` to every folder that lacked one: `scripts/`, `scripts/Custom Scripts/`, `scripts/Custom Scripts/Intune/` (+ `Desktop/`, `Office Themes/`, `Add Lockscreen to start and desktop/`, `Background/`), `scripts/Custom Scripts/device/Time sync/`, `scripts/Graph/`, `scripts/Intune/` (+ `Get-Autopilot/`), `scripts/Testing Scripts/`, `scripts/Testing Scripts/Network/`, `scripts/Testing Scripts/RDS/` — each with a file list and parameter/usage docs |
| Fixed `scripts/Custom Scripts/device/audio/Rollback-InternalMic` — file was missing its `.ps1` extension |
| Renamed `scripts/Entra/remove-m365users.ps1` → `Remove-M365Users.ps1` for naming consistency (menu.ps1 and readmes already referenced the PascalCase form) |
| Fixed `scripts/Entra/readme.md` — removed a stale `Distributionlist.ps1` entry that actually documented `scripts/Exchange/Set-Distributionlist-dynamic-static.ps1`; moved accurate docs to `scripts/Exchange/readme.md`; added missing `Set-UserManager.ps1` docs |
| Fixed `scripts/Testing Scripts/SharePoint/readme.md` — was a stale duplicate of `Get-SharePointStorageReport.ps1` docs (script doesn't live in this folder); replaced with a pointer to `scripts/Reporting/readme.md`, which now documents the script's full current parameter set (`-ClientId`, `-ClientSecret`, `-CertificateThumbprint`, `-RecycleBinOnly`, `-GraphTimeoutSec`, `-MaxGraphRetry` were previously undocumented) |
| Fixed `scripts/Custom Scripts/device/audio/readme.md` — corrected script name casing to match the actual files on disk |
| Removed tracked `.DS_Store` files from git and added `.DS_Store` to `.gitignore` |

### 2026-04-17
| Change |
|--------|
| Updated `scripts/Custom Scripts/Save install time/start.bat` — added option `D` (customer install scripts from local `Install` folder) and option `E` (customer install scripts from `\\10.222.3.94\Software`); before deployment starts, creates/updates local admin `LocalAdmin` (`Er@smus_Roter0`), adds it to `Administrators`, and sets OOBE skip flags |
| Added `scripts/Custom Scripts/Save install time/Browse-InstallScripts.ps1` — customer-first browser that lists customer folders as menu items and launches `.ps1`, `.bat`, and `.cmd` scripts |
| Updated `scripts/Custom Scripts/Save install time/readme.md` — documented new `D`/`E` menu options, including that option `D` requires copying both `Browse-InstallScripts.ps1` and the full `Install` folder, and that customer deploy options prepare `LocalAdmin` plus OOBE skip flags |

### 2026-04-16
| Change |
|--------|
| Updated `scripts/Custom Scripts/Intune/Desktop/Background/Lockscreen/Make-lockscreen.ps1` to v2.0 — aligned lockscreen source with corporate wallpaper config (`$ImageUrl`, `$ClientName`), replaced direct `WebClient` with validated internet download flow (`Invoke-WebRequest`), added GitHub blob/raw URL normalization, image signature checks (`jpg/png/bmp`), HTML-response guard, structured Intune logging, and safer temporary download handling |
| Added `scripts/Custom Scripts/Intune/Desktop/Background/Lockscreen/readme.md` — documentation for configuration, deployment, logging, workflow, and lockscreen-specific version history |
| Updated root `readme.md` — expanded Intune Desktop/Background documentation and repository structure to include the lockscreen script and docs |

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
| Updated `scripts/Reporting/Licensing/genereer_licentie_overzicht.py` — when the same product appears with multiple billing periods on one invoice (Pax8 and Ingram), each period is now shown as a separate row with the period range in the Category/Detail column instead of being summed incorrectly |
| Updated `scripts/Reporting/Licensing/genereer_rapport.ps1` — CMD window now closes automatically when run as a scheduled task; `Read-Host` pauses are skipped when `[Environment]::UserInteractive` is false |
| Updated `scripts/Reporting/Licensing/genereer_rapport.bat` — pipes stdin from `NUL` so Python's interactive pause is never triggered when run as a scheduled task |

### 2026-04-01
| Change |
|--------|
| Updated `scripts/Exchange/Migrate-Calendar.ps1` — fixed Room Mailbox booking issues: increased provisioning wait from 15s to 60s; added retry loop (5×30s) for `Set-CalendarProcessing` with error handling and fallback instructions; changed `BookingWindowInDays 0` to `1825` and added `EnforceSchedulingHorizon $false` to prevent silent booking rejections |

### 2026-03-30
| Change |
|--------|
| Updated `scripts/Custom Scripts/Intune/Desktop/Background/Desktop/Set-CorporateWallpaper.ps1` — added idempotency check: downloads image to temp, compares SHA256 hash against existing file; skips if hash matches and PersonalizationCSP is correct; applies (without second download) if image is new or changed |
| Updated readme — Intune & Autopilot section: expanded `Set-CorporateWallpaper.ps1` documentation with configuration table, deployment steps, log path, and NinjaOne/Intune deploy instructions |

### 2026-03-27
| Change |
|--------|
| Added `scripts/Reporting/Get-ComputerLastLogon.ps1` — last logon report for computers in one or more OUs; `LastLogonTimestamp` (fast) or `-AllDCs` (accurate) mode; marks Active/Stale/Never/Disabled; exports timestamped CSV to `C:\Temp\`; `-InactiveDays`, `-IncludeDisabled`, `-ExportPath` parameters |
| Added `scripts/Reporting/readme.md` — documents `Get-ComputerLastLogon.ps1` with parameter table, CSV column reference, and usage examples |
| Updated `scripts/Reporting/Get-ComputerLastLogon.ps1` — added `PasswordLastSet` / `DaysSincePasswordSet` columns; new `Active (pwd recent)` status for devices falsely marked stale due to 14-day `LastLogonTimestamp` replication delay |

### 2026-03-26
| Change |
|--------|
| Updated `scripts/Custom Scripts/device/audio/Disable-internalmic.ps1` — expanded internal mic patterns: added Conexant language variants (EN/FR/NL), Synaptics EN, Intel SST, IDT, Cirrus Logic drivers, and multilingual microphone array names (FR/DE/ES/PT/IT) |
| Updated readme — Audio Management section: added NinjaOne deployment table (Run as SYSTEM, no parameters, `AudioDeviceInventory` custom field, exit codes) for all three audio scripts |
| Updated readme — `Invoke-WindowsActivation.ps1`: added NinjaOne deployment table with Run as Administrator, exit codes, and parameter examples per scenario; warning added for `-RemoveKey`/`-ReArm` requiring `-Force` |
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
| Updated `Get-SharePointStorageReport.ps1` — two-phase approach (enumerate all sites/libraries first, then retrieve storage data); fixed inline `if` syntax error |
| Added `Test-AuthNetworkDiagnostics.ps1` — auth & network diagnostics: Event Viewer (4625/4771/4776/4740/5719), time sync, Kerberos cache, DNS, TCP, UNC shares, optional log scan |
| Added `scripts/Custom Scripts/SAS/` — SAS batch error monitoring with Zabbix integration, email alerts, and Event Viewer analysis |

### 2026-03-24
| Change |
|--------|
| Added `scripts/Intune/iOS-Compliance-Updater/` — auto-update minimum iOS version in Intune compliance policy via Graph API; weekly scheduled task, dry-run support, one-time Setup.ps1 |
| Added `Get-SharePointStorageReport.ps1` — tenant-wide SharePoint storage report with version history per file; quick mode (quota) and full recursive scan |
| Added `Test-OpenVpnDiagnostics.ps1` — OpenVPN Connect diagnostics: PnP adapters, services, routes, DNS, Event Log, conflicting VPN software; txt export to `C:\Temp\` |

### 2026-03-23
| Change |
|--------|
| All CSV exports now go to `C:\Temp\` (Windows) or `~/Downloads/` (macOS/Linux) |
| Added `Import-DnsRecords.ps1` — resolve public DNS via dig (Google 8.8.8.8) and import A/CNAME records into AD DNS, dry-run by default |
| Added `New-M365User.ps1` — create single M365 user via Graph, auto-generated password, optional license |
| Added `Import-M365Users.ps1` — bulk user creation from CSV via Graph, dry-run by default, passwords in CSV output |
| Added `Remove-M365Users.ps1` — bulk Entra ID user removal, dry-run by default, CSV report |
| Added `Get-ExternalForwards.ps1` — audit external forwarding rules across all mailboxes, CSV export |
| Added `Get-MailboxSizes.ps1` — mailbox size + item count report, sorted by storage, CSV export |
| Added `Test-DkimConfig.ps1` — DKIM signing config + DNS CNAME/TXT validation with required-actions output |
| Added `Test-M365GroupMembership.ps1` — M365 Group / Teams owner and member audit via Graph, CSV export |
| Added `Test-DistributionGroupPermissions.ps1` — DG managers, Send As, Send on Behalf, member counts, CSV export |
| Added `Test-CalendarPermissions.ps1` — locale-independent calendar permission audit, CSV export |
| Added `Test-MailboxPermissions.ps1` — Full Access / Send As / Send on Behalf audit, CSV export |
| Exchange submenu (`C`) and Entra submenu (`D`): permission audit tools added |
| Moved `Test-Ports.ps1` to `scripts/Testing Scripts/Network/` |
| `load.ps1` — auto-imports modules at startup; detects missing modules and offers install |
| Added `load.ps1` — first-run setup (UPN + name), saves to gitignored `load.config.ps1`, launches menu |
| Extended `menu.ps1` with M365 section (B–E): Exchange, Entra ID, MSP Admin submenus |
| Added `menu.ps1` — interactive launcher, single-keypress, number + F-keys, cross-platform |
| Added `scripts/Network/Test-Ports.ps1` — TCP port checker, range/list syntax, multi-target |
| Licensing scripts: translated to English, genericised, configurable export dir |
| Rewrote `create_scheduled_task.ps1` — admin check, auto-detect Python, dynamic trigger date |
| Added `scripts/Reporting/Licensing/` — Pax8 + Ingram → Excel report toolkit |

### 2026-03-20
| Change |
|--------|
| `start.bat` v2.8 — split Do it all: A = Intune, C = AD; added device rename (B) and AD join (8) |
| Rewrote `Migrate-Calendar.ps1` v2.0 — English, generic, mandatory params |
| Added SMTP test scripts; added `Test-SmtpRelay` to `functies.ps1` |
| Rewrote `functies.ps1` — replaced MSOnline/AzureAD with Microsoft Graph, cross-platform |
| Added `Set-CorporateWallpaper.ps1`, `Set-Calendar-rights.ps1`, `Restart-Time-Sync.ps1` |
| Removed all company-specific references; translated all readmes to English |

### 2026-03-19
| Change |
|--------|
| Initial repository upload |

---

## Maintainer

**Sjoerd Kanon** — Security-minded | Team- & Projectgericht | Microsoft 365 & Infrastructure
