# M365-Scripts

> A collection of PowerShell scripts and M365 management tools for MSP engineers, maintained by Sjoerd Kanon.

---

## Overview

This repository contains production-ready tooling used by engineers to automate, manage, and report on Microsoft 365 tenants. Everything is accessible through an interactive menu — scripts are organised by workload and actively maintained.

---

## Getting Started

```powershell
.\load.ps1
```

That's it. On first run, `load.ps1` will:

1. Ask for your admin UPN and display name — saved to a gitignored `load.config.ps1`
2. Detect missing modules and offer to install them automatically
3. Import all required modules
4. Open the interactive menu

From then on it starts directly without any prompts.

> You can also run `.\menu.ps1` directly — it will ask for your UPN as a fallback.
> To reinstall or update modules manually: `.\scripts\Startup\Install-Modules.ps1`

---

## Requirements

- **PowerShell** 7.0+ (cross-platform) — individual scripts support PS 5.1 on Windows
- **Microsoft 365 admin permissions** for the target workload
- Windows only — enable script execution if needed:

```powershell
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

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
| `C` | M365 | Exchange Online submenu (incl. calendar, mailbox, and DG permission audits) |
| `D` | M365 | Entra ID / Graph submenu (incl. user create/import/remove + M365 group audit) |
| `E` | M365 | MSP Admin submenu |

M365 options (`B`–`E`) lazy-load `functies.ps1` on first use — Graph authentication is only triggered when needed.

**Exchange submenu (`C`)** audit tools:

| Key | Tool |
|-----|------|
| `8` | Test-CalendarPermissions — audit calendar folder permissions (all or single mailbox) |
| `9` | Test-MailboxPermissions — audit Full Access, Send As, Send on Behalf |
| `A` | Test-GroupPermissions — audit DG managers, Send As, Send on Behalf, member counts |
| `B` | Test-DkimConfig — validate DKIM signing config and DNS CNAME/TXT records |
| `C` | Get-ExternalForwards — audit mailboxes with external forwarding configured |
| `D` | Get-MailboxSizes — mailbox size report sorted by storage used |

**Entra ID submenu (`D`)** audit tool:

| Key | Tool |
|-----|------|
| `A` | Test-M365GroupMembership — audit M365 Group / Teams owners and members |
| `B` | New-M365User — create a single new user (auto-generated password, optional license) |
| `C` | Import-M365Users — bulk create users from CSV, dry-run by default |

---

## Categories

### 📧 Exchange
Scripts for calendar and mailbox management.

- Calendar migration between users
- Set calendar folder permissions (NL/FR/EN locale support)

### ☁️ M365 Management (`functies.ps1`)
Interactive M365 management functions via Microsoft Graph and Exchange Online. Loaded as a library through the menu. CSV and log exports go to `C:\Temp\` on Windows or `~/Downloads/` on macOS.

- **Exchange Online** — shared mailbox access, locale, aliases, distribution groups, auto-reply, sent-items copy
- **Entra ID / Graph** — tenant admins, domains, licenses, users, password reset, sign-in logs, bulk user creation, bulk user removal
- **MSP Admin** — create/manage MSP admin account across customer tenants

### 📱 Intune / Autopilot
Scripts for device enrollment and Autopilot registration.

- Retrieve Windows Autopilot hardware info
- CMD-based Autopilot enrollment helper

### 💾 USB Setup Toolkit
USB toolkit for Windows setup and Autopilot enrollment during OOBE.

- Interactive menu (Device Manager, Autopilot, AD join, device rename, product key, Windows Update, restart)
- Self-elevating, OOBE-compatible via Shift+F10
- Split "Do it all": `A` = Intune (Rename + Autopilot + Update), `C` = AD (Rename + Domain join + Update)

### 🖥️ Device — Audio
Scripts for managing audio device configuration on endpoints.

- Detect connected audio devices
- Disable internal microphone via policy
- Rollback internal mic changes

### ⏱️ Device — Time Sync
Scripts for managing Windows time synchronisation.

- Restart and force Windows Time service sync

### 📊 Reporting — Licensing
Monthly licensing and Azure cost report generator.

- Combines Pax8 (CSV) and Ingram (Excel) billing data into one formatted Excel
- Per-customer tab with Azure consumption, licenses, and Acronis breakdown
- Summary tab with totals and margin per customer
- PowerShell launcher with pre-flight validation
- Optional Windows scheduled task (runs on the 6th of each month)

### 🧪 Testing — Device
Diagnostic scripts for Windows endpoints.

- OpenVPN Connect diagnostics — PnP adapters, services, routes, DNS, Event Log, conflicting VPN software; exports txt report to `C:\Temp\`

### 🧪 Testing — Connectivity
Scripts for diagnosing network and mail connectivity.

- Test TCP port connectivity on any host — single ports, ranges (`1294:1494`), combinations (`80,443,1294:1494`)
- One-time SMTP test with interactive credential prompt
- Recurring SMTP test (every 5 minutes) with saved encrypted password
- Auth & network diagnostics — Event Viewer (logon failures, Kerberos, NTLM, DC availability), time sync, Kerberos ticket cache, DNS, TCP connectivity, UNC share access, optional log file scan; exports txt report to `C:\Temp\`

### 🧪 Testing — Exchange
Audit and diagnostic scripts for Exchange Online. Self-connecting — reuse an existing session or connect automatically. CSV exports go to `C:\Temp\` on Windows or `~/Downloads/` on macOS.

- Audit calendar folder permissions (locale-independent, exports CSV)
- Audit Full Access, Send As, Send on Behalf delegation (exports CSV)
- Audit distribution group managers, Send As, Send on Behalf, member counts (exports CSV)
- Validate DKIM signing config and DNS CNAME/TXT records; lists required actions
- Audit mailboxes with external forwarding to non-tenant domains (security audit, exports CSV)
- Report mailbox sizes and item counts sorted by storage used (exports CSV)

### 🧪 Testing — Entra ID / Graph
Audit scripts for Microsoft 365 groups via Microsoft Graph. CSV exports go to `C:\Temp\` on Windows or `~/Downloads/` on macOS.

- Audit M365 Group (incl. Teams) owners and members — one row per entry, exports CSV

### 🧪 Testing — SharePoint
Storage audit scripts for SharePoint Online via Microsoft Graph. CSV exports go to `C:\Temp\` on Windows or `~/Downloads/` on macOS.

- Report storage usage across all sites in a tenant — current file sizes + version history per library and per file
- Quick mode (quota data only) or full recursive scan with `-Apply`

### 📊 SAS Batch Monitoring
Monitor SAS batch job logs and Windows Event Viewer for errors, with optional Zabbix integration and email alerts.

- Detects spawn errors, WORK library auth failures, aborts, disk errors, and general `ERROR:` lines
- Text, JSON, and Zabbix output formats; configurable look-back period
- One-time setup script — installs to `C:\Scripts\`, creates daily scheduled task
- Optional Zabbix UserParameter config for automated alerting

### 🌐 DNS Management
Scripts for managing DNS records in Active Directory-integrated DNS zones.

- Resolve public DNS records via Google DNS (dig) and import them as A or CNAME records into AD DNS
- Dry-run by default — shows what would be created before applying
- Idempotent — skips records that already exist

---

### 📱 Intune — iOS Compliance Updater
Automatically keeps the minimum iOS version requirement in an Intune compliance policy up to date.

- Fetches the latest iOS version from Apple's RSS feed (with fallback to Apple Support page)
- Compares against the current policy minimum and patches via Microsoft Graph API
- One-time setup via `Setup.ps1` (creates App Registration, assigns permissions, writes `config.json`)
- Runs weekly as a Windows scheduled task (SYSTEM, every Monday 07:00)
- Dry-run mode (`-WhatIf`) — shows what would change without applying

### 🖼️ Intune — Desktop
Scripts and assets for managing desktop and lockscreen configuration.

- Deploy lockscreen to start and desktop
- Set corporate wallpaper via Intune (PersonalizationCSP + WinAPI + Default User)

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
        │   └── Test-AuthNetworkDiagnostics.ps1   ← auth/network issue diagnostics
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

| Date | Change |
|------|--------|
| 2026-03-25 | Added `scripts/Testing Scripts/Network/Test-AuthNetworkDiagnostics.ps1` — auth & network diagnostics: Event Viewer (4625/4771/4776/4740/5719), time sync, Kerberos cache, DNS, TCP, UNC shares, optional log scan |
| 2026-03-25 | Added `scripts/Custom Scripts/SAS/` — SAS batch error monitoring with Zabbix integration, email alerts, and Event Viewer analysis |
| 2026-03-24 | Added `scripts/Intune/iOS-Compliance-Updater/` — auto-update minimum iOS version in Intune compliance policy via Graph API; weekly scheduled task, dry-run support, one-time Setup.ps1 |
| 2026-03-24 | Added `scripts/Testing Scripts/SharePoint/Get-SharePointStorageReport.ps1` — tenant-wide SharePoint storage report with version history per file; quick mode (quota) and full recursive scan |
| 2026-03-24 | Added `scripts/Testing Scripts/Device/Test-OpenVpnDiagnostics.ps1` — OpenVPN Connect diagnostics: PnP adapters, services, routes, DNS, Event Log, conflicting VPN software; txt export to `C:\Temp\` |
| 2026-03-23 | All CSV exports now go to `C:\Temp\` (Windows) or `~/Downloads/` (macOS/Linux) — no longer written to current directory |
| 2026-03-23 | Added `scripts/Custom Scripts/DNS/Import-DnsRecords.ps1` — resolve public DNS via dig (Google 8.8.8.8) and import A/CNAME records into AD DNS, dry-run by default |
| 2026-03-23 | Added `scripts/Entra/New-M365User.ps1` — create single M365 user via Graph, auto-generated password, optional license |
| 2026-03-23 | Added `scripts/Entra/Import-M365Users.ps1` — bulk user creation from CSV via Graph, dry-run by default, passwords in CSV output |
| 2026-03-23 | Added `scripts/Testing Scripts/Exchange/Get-ExternalForwards.ps1` — audit external forwarding rules across all mailboxes, CSV export |
| 2026-03-23 | Added `scripts/Testing Scripts/Exchange/Get-MailboxSizes.ps1` — mailbox size + item count report, sorted by storage, CSV export |
| 2026-03-23 | Added `scripts/Testing Scripts/Exchange/Test-DkimConfig.ps1` — DKIM signing config + DNS CNAME/TXT validation with required-actions output |
| 2026-03-23 | Added `scripts/Testing Scripts/Entra/Test-M365GroupMembership.ps1` — M365 Group / Teams owner and member audit via Graph, CSV export |
| 2026-03-23 | Added `scripts/Testing Scripts/Exchange/Test-DistributionGroupPermissions.ps1` — DG managers, Send As, Send on Behalf, member counts, CSV export |
| 2026-03-23 | Added `scripts/Testing Scripts/Exchange/Test-CalendarPermissions.ps1` — locale-independent calendar permission audit, CSV export |
| 2026-03-23 | Added `scripts/Testing Scripts/Exchange/Test-MailboxPermissions.ps1` — Full Access / Send As / Send on Behalf audit, CSV export |
| 2026-03-23 | Exchange submenu (`C`) items 8/9/A and Entra submenu (`D`) item A: permission audit tools |
| 2026-03-23 | Moved `Test-Ports.ps1` from `scripts/Network/` to `scripts/Testing Scripts/Network/` |
| 2026-03-23 | Added `scripts/Entra/Remove-M365Users.ps1` — bulk Entra ID user removal, dry-run by default, CSV report |
| 2026-03-23 | `load.ps1` — auto-imports modules at startup; detects missing modules and offers to run `Install-Modules.ps1` |
| 2026-03-23 | Added `load.ps1` — first-run setup (UPN + name), saves to gitignored `load.config.ps1`, launches menu |
| 2026-03-23 | Extended `menu.ps1` with M365 section (B–E): Exchange, Entra ID, MSP Admin submenus via `functies.ps1` |
| 2026-03-23 | Added `menu.ps1` — interactive launcher, single-keypress, number + F-keys, cross-platform |
| 2026-03-23 | Added `scripts/Network/Test-Ports.ps1` — TCP port checker, range/list syntax, multi-target |
| 2026-03-23 | Licensing scripts: translated to English, genericised (no hardcoded customers/paths), configurable export dir |
| 2026-03-23 | Rewrote `create_scheduled_task.ps1` — admin check, auto-detect Python, dynamic trigger date |
| 2026-03-23 | Added `scripts/Reporting/Licensing/` — Pax8 + Ingram → Excel report toolkit |
| 2026-03-20 | `start.bat` v2.8 — split Do it all: A = Intune, C = AD; added device rename (B) and AD join (8) |
| 2026-03-20 | Rewrote `Migrate-Calendar.ps1` v2.0 — English, generic, mandatory params |
| 2026-03-20 | Added SMTP test scripts; added `Test-SmtpRelay` to `functies.ps1` |
| 2026-03-20 | Rewrote `functies.ps1` — replaced MSOnline/AzureAD with Microsoft Graph, cross-platform |
| 2026-03-20 | Added `Set-CorporateWallpaper.ps1`, `Set-Calendar-rights.ps1`, `Restart-Time-Sync.ps1` |
| 2026-03-20 | Removed all company-specific references; translated all readmes to English |
| 2026-03-19 | Initial repository upload |

---

## Maintainer

**Sjoerd Kanon** — Security-minded | Team- & Projectgericht | Microsoft 365 & Infrastructure
