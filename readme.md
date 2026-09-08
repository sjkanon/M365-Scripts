# M365-Scripts

> A collection of PowerShell scripts and M365 management tools for MSP engineers, maintained by Sjoerd Kanon.

---

## Table of Contents

- [Getting Started](#getting-started)
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
| `9` / `F9` | Startup | Install-Modules |
| `A` / `F10` | Reporting | Licensing-Report |
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
| `detect-audiodevices.ps1` | Inventory van alle audio endpoints op het toestel |
| `Disable-internalmic.ps1` | Disable interne microfoon(s), headsets worden overgeslagen |
| `Rollback-InternalMic.ps1` | Heractiveer eerder uitgeschakelde interne microfoons |

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
| [`TenantOnboarding/`](../scripts/TenantOnboarding/readme.md) | Internal tenant-setup toolkit | New-tenant provisioning (break-glass admin, baseline groups/Intune assignment), multi-tenant/GDAP license + break-glass password reporting, Win32/Chocolatey app deployment, device config (kiosk power, Office uninstall, Start menu layout), OneDrive management, dynamic-DG/feature-group user management |
| [`Office365Toolkit/`](../scripts/Office365Toolkit/readme.md) | Fork of [`directorcia/Office365`](https://github.com/directorcia/Office365) (CIAOPS) | Secure Score reporting, enterprise app consent cleanup, shared mailbox sign-in lockdown, EOP baseline, mailbox hygiene/forwarding-risk audits, Unified Audit Log search, Intune policy inventory |
| [`PatronToolkit/`](../scripts/PatronToolkit/readme.md) | Fork of [`directorcia/patron`](https://github.com/directorcia/patron) | MFA registration + CA policy export, enterprise app consent + suspicious inbox rule + unified security alert audits, consolidated email security posture + mailbox auditing checks, SPF/DMARC validation, Intune policy assignment + Autopilot device reports, message trace, SharePoint sharing config, Teams config report |
| [`LegacyUtilities/`](../scripts/LegacyUtilities/readme.md) | Internal toolkit (misc small scripts) | Mailbox folder permissions/delegate access, bulk shared mailbox/contact creation, contact sync, duplicate mail item cleanup, M365 group membership, CA policy backup, Teams/Planner cloning, Azure Files drive mapping, NumLock/lock-workstation device tweaks, Workspace 365 environment provisioning |

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
    │   ├── Set-Calendar-rights.ps1
    │   ├── Set-Distributionlist-dynamic-static.ps1
    │   ├── Move-InboxToArchive.ps1
    │   ├── Test-CalendarPermissions.ps1
    │   ├── Test-MailboxPermissions.ps1
    │   ├── Test-DistributionGroupPermissions.ps1
    │   ├── Test-DkimConfig.ps1
    │   ├── Get-ExternalForwards.ps1
    │   └── Get-MailboxSizes.ps1
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
    │   └── Restore-RecycleBinItems.ps1  ← restore deleted files from a recycle bin: one site/OneDrive or tenant-wide (PnP, auto app registration)
    ├── Teams/
    │   ├── readme.md
    │   └── vias_archiver.ps1        ← Teams/SharePoint export + archiving (Graph, PS7+, Global Admin)
    ├── Reporting/
    │   ├── readme.md
    │   ├── Get-ComputerLastLogon.ps1        ← last logon per computer in OU(s), export to CSV
    │   ├── Get-SharePointStorageReport.ps1  ← tenant-wide SharePoint storage report
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
    │   └── Test-PowerShellSyntax.ps1
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
