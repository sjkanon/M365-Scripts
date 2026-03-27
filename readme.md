# M365-Scripts

> A collection of PowerShell scripts and M365 management tools for MSP engineers, maintained by Sjoerd Kanon.

---

## Table of Contents

- [Getting Started](#getting-started)
- [Requirements](#requirements)
- [Menu](#menu)
- [Script Categories](#script-categories)
  - [M365 Management](#️-m365-management)
  - [Exchange](#-exchange)
  - [Entra ID / Graph](#-entra-id--graph)
  - [Intune & Autopilot](#-intune--autopilot)
  - [Testing & Diagnostics](#-testing--diagnostics)
  - [Reporting](#-reporting)
  - [Infrastructure & Devices](#️-infrastructure--devices)
  - [Custom Tools](#-custom-tools)
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
2. Detect missing modules and offer to install them automatically
3. Import all required modules
4. Open the interactive menu

From then on it starts directly without any prompts.

> You can also run `.\menu.ps1` directly — it will ask for your UPN as a fallback.
> To reinstall or update modules manually: `.\scripts\Startup\Install-Modules.ps1`

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
| `B` | M365 | Connect-Tenant |
| `C` | M365 | Exchange Online submenu |
| `D` | M365 | Entra ID / Graph submenu |
| `E` | M365 | MSP Admin submenu |

M365 options (`B`–`E`) lazy-load `functies.ps1` on first use — Graph authentication is only triggered when needed.

**Exchange submenu (`C`)**

| Key | Tool |
|-----|------|
| `8` | Test-CalendarPermissions — audit calendar folder permissions (all or single mailbox) |
| `9` | Test-MailboxPermissions — audit Full Access, Send As, Send on Behalf |
| `A` | Test-GroupPermissions — audit DG managers, Send As, Send on Behalf, member counts |
| `B` | Test-DkimConfig — validate DKIM signing config and DNS CNAME/TXT records |
| `C` | Get-ExternalForwards — audit mailboxes with external forwarding configured |
| `D` | Get-MailboxSizes — mailbox size report sorted by storage used |

**Entra ID submenu (`D`)**

| Key | Tool |
|-----|------|
| `A` | Test-M365GroupMembership — audit M365 Group / Teams owners and members |
| `B` | New-M365User — create a single new user (auto-generated password, optional license) |
| `C` | Import-M365Users — bulk create users from CSV, dry-run by default |

---

## Script Categories

### ☁️ M365 Management

Interactive M365 management functions via Microsoft Graph and Exchange Online. Loaded as a library through the menu. CSV and log exports go to `C:\Temp\` on Windows or `~/Downloads/` on macOS.

| Area | Features |
|------|----------|
| Exchange Online | Shared mailbox access, locale, aliases, distribution groups, auto-reply, sent-items copy |
| Entra ID / Graph | Tenant admins, domains, licenses, users, password reset, sign-in logs, bulk create/remove |
| MSP Admin | Create/manage MSP admin account across customer tenants |

---

### 📧 Exchange

Scripts for calendar and mailbox management.

- Calendar migration between users
- Set calendar folder permissions (NL/FR/EN locale support)

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
- **iOS Compliance Updater** — automatically keeps the minimum iOS version requirement in Intune up to date
  - Fetches latest iOS version from Apple's RSS feed (with fallback to Apple Support page)
  - Compares against current policy minimum and patches via Microsoft Graph API
  - One-time setup via `Setup.ps1` (creates App Registration, assigns permissions, writes `config.json`)
  - Runs weekly as a Windows scheduled task (SYSTEM, every Monday 07:00)
  - Dry-run mode (`-WhatIf`) — shows what would change without applying
- **Desktop** — deploy lockscreen to start and desktop; set corporate wallpaper via Intune (PersonalizationCSP + WinAPI + Default User)

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
- Marks each computer as **Active**, **Stale**, **Never**, or **Disabled** based on `-InactiveDays` threshold (default 90)
- CSV columns: Name, Status, Enabled, LastLogon, DaysSinceLogon, OS, IPv4, OU path, Created, Description
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

## Repository Structure

```
M365-Scripts/
├── .gitignore
├── .vscode/
│   └── settings.json
├── load.ps1                         ← Entry point: first-run setup + launches menu
├── menu.ps1                         ← Interactive launcher (all scripts + M365 functions)
├── readme.md
└── scripts/
    ├── Custom Scripts/
    │   ├── device/
    │   │   ├── Invoke-WindowsActivation.ps1 ← activate Windows, set product key / KMS server
    │   │   ├── Invoke-WindowsCleanup.ps1    ← temp, cache, WU, DISM, browser, event logs
    │   │   ├── audio/
    │   │   │   ├── detect-audiodevices.ps1
    │   │   │   ├── Disable-internalmic.ps1
    │   │   │   └── readme.md
    │   │   └── Time sync/
    │   │       └── Restart-Time-Sync.ps1
    │   ├── Intune/Desktop/
    │   │   ├── Add Lockscreen to start and desktop/
    │   │   └── Background/Desktop/
    │   │       ├── Set-CorporateWallpaper.ps1
    │   │       └── readme.md
    │   ├── DNS/
    │   │   ├── Import-DnsRecords.ps1   ← resolve via Google DNS + import into AD DNS
    │   │   ├── example-records.csv
    │   │   └── readme.md
    │   ├── SAS/
    │   │   ├── Monitor-SASBatchErrors.ps1   ← scan logs + Event Viewer for SAS errors
    │   │   ├── Setup-SASMonitoring.ps1      ← install script, scheduled task, Zabbix config
    │   │   ├── Test-SASWorkDirectory.ps1    ← validate WORK directory health
    │   │   ├── zabbix_sas_monitor.conf
    │   │   └── readme.md
    │   └── Save install time/       ← USB setup toolkit
    │       ├── start.bat
    │       ├── autorun.inf
    │       └── readme.md
    ├── Entra/
    │   ├── Import-M365Users.ps1
    │   ├── New-M365User.ps1
    │   ├── Remove-M365Users.ps1
    │   └── readme.md
    ├── Exchange/
    │   ├── Migrate-Calendar.ps1
    │   ├── Set-Calendar-rights.ps1
    │   └── readme.md
    ├── Intune/
    │   ├── Get-Autopilot/
    │   │   ├── Get-WindowsAutoPilotInfo.ps1
    │   │   └── GetAutoPilot.CMD
    │   └── iOS-Compliance-Updater/
    │       ├── Update-iOSCompliancePolicy.ps1   ← main script (run or scheduled task)
    │       ├── Setup.ps1                        ← one-time: App Registration + config.json
    │       ├── Install-ScheduledTask.ps1        ← register weekly scheduled task
    │       ├── config.example.json
    │       └── readme.md
    ├── Reporting/
    │   ├── Get-ComputerLastLogon.ps1        ← last logon per computer in OU(s), export to CSV
    │   ├── readme.md
    │   └── Licensing/
    │       ├── genereer_licentie_overzicht.py
    │       ├── genereer_rapport.ps1
    │       ├── genereer_rapport.bat
    │       ├── create_scheduled_task.ps1
    │       └── readme.md
    ├── Startup/
    │   ├── functies.ps1             ← M365 function library (dot-sourced by menu)
    │   ├── Install-Modules.ps1      ← Bootstrap: install & import all modules
    │   └── readme.md
    └── Testing Scripts/
        ├── Entra/
        │   ├── Test-M365GroupMembership.ps1
        │   └── readme.md
        ├── Exchange/
        │   ├── Get-ExternalForwards.ps1
        │   ├── Get-MailboxSizes.ps1
        │   ├── Test-CalendarPermissions.ps1
        │   ├── Test-DkimConfig.ps1
        │   ├── Test-DistributionGroupPermissions.ps1
        │   ├── Test-MailboxPermissions.ps1
        │   └── readme.md
        ├── Device/
        │   ├── Test-OpenVpnDiagnostics.ps1
        │   └── readme.md
        ├── Network/
        │   ├── Test-Ports.ps1
        │   ├── Test-AuthNetworkDiagnostics.ps1   ← auth/network issue diagnostics
        │   └── Test-FileIODiagnostics.ps1        ← file I/O test + real-time directory monitor
        ├── RDS/
        │   ├── Test-RDSDiagnostics.ps1           ← RDP/RDWeb login failure diagnostics
        │   └── Watch-RDSLive.ps1                 ← real-time session + licensing monitor
        ├── SharePoint/
        │   ├── Get-SharePointStorageReport.ps1
        │   └── readme.md
        └── SMTP/
            ├── testsmtp.ps1
            ├── testsmtp_5min.ps1
            └── readme.md
```

---

## Contributing

When adding new scripts:

1. Follow the existing naming convention (`Verb-Noun.ps1`)
2. Include a header comment block with Synopsis, Description, Parameters, and Example
3. Test against a non-production tenant before committing
4. Place the script in the appropriate workload folder
5. Add it to `menu.ps1` and update this readme

---

## Disclaimer

These scripts are provided as-is. Always test in a non-production environment before running against live tenants. The maintainer accepts no liability for unintended changes resulting from misuse or misconfiguration.

---

## Version History

### 2026-03-27
| Change |
|--------|
| Added `scripts/Reporting/Get-ComputerLastLogon.ps1` — last logon report for computers in one or more OUs; `LastLogonTimestamp` (fast) or `-AllDCs` (accurate) mode; marks Active/Stale/Never/Disabled; exports timestamped CSV to `C:\Temp\`; `-InactiveDays`, `-IncludeDisabled`, `-ExportPath` parameters |
| Added `scripts/Reporting/readme.md` — documents `Get-ComputerLastLogon.ps1` with parameter table, CSV column reference, and usage examples |

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
