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
| `1` / `F1` | Network | Test-Ports — TCP port checker |
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
| `D` | M365 | Entra ID / Graph submenu (incl. bulk user removal) |
| `E` | M365 | MSP Admin submenu |

M365 options (`B`–`E`) lazy-load `functies.ps1` on first use — Graph authentication is only triggered when needed.

---

## Categories

### 🔌 Network
Scripts for network diagnostics.

- Test TCP port connectivity on any host — single ports, ranges (`1294:1494`), combinations (`80,443,1294:1494`)

### 📧 Exchange
Scripts for calendar and mailbox management.

- Calendar migration between users
- Set calendar folder permissions (NL/FR/EN locale support)

### ☁️ M365 Management (`functies.ps1`)
Interactive M365 management functions via Microsoft Graph and Exchange Online. Loaded as a library through the menu.

- **Exchange Online** — shared mailbox access, locale, aliases, distribution groups, auto-reply, sent-items copy
- **Entra ID / Graph** — tenant admins, domains, licenses, users, password reset, sign-in logs, bulk user removal
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

### 🧪 Testing — SMTP
Scripts for diagnosing SMTP connectivity and authentication.

- One-time SMTP test with interactive credential prompt
- Recurring SMTP test (every 5 minutes) with saved encrypted password

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
    │   └── Save install time/       ← USB setup toolkit
    │       ├── start.bat
    │       ├── autorun.inf
    │       └── readme.md
    ├── Entra/
    │   ├── Remove-M365Users.ps1
    │   └── readme.md
    ├── Exchange/
    │   ├── Migrate-Calendar.ps1
    │   ├── Set-Calendar-rights.ps1
    │   └── readme.md
    ├── Intune/Get-Autopilot/
    │   ├── Get-WindowsAutoPilotInfo.ps1
    │   └── GetAutoPilot.CMD
    ├── Network/
    │   └── Test-Ports.ps1
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
    └── Testing Scripts/SMTP/
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
