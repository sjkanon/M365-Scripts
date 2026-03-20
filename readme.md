# M365-Scripts

> A collection of PowerShell scripts for managing Microsoft 365 environments, maintained by Sjoerd Kanon.

---

## Overview

This repository contains production-ready PowerShell scripts used by engineers to automate, manage, and report on Microsoft 365 tenants. Scripts are organized by workload and are actively maintained and expanded over time.

---

## Categories

### 📧 Exchange
Scripts for calendar and mailbox management.

- Calendar migration between users
- Set calendar folder permissions (NL/FR/EN locale support)

### 📱 Intune / Autopilot
Scripts for device enrollment and Autopilot registration.

- Retrieve Windows Autopilot hardware info
- CMD-based Autopilot enrollment helper

### 💾 USB Setup Toolkit
USB toolkit for Windows setup and Autopilot enrollment during OOBE.

- Interactive menu (Device Manager, Autopilot, product key, restart)
- Self-elevating, OOBE-compatible via Shift+F10

### 🖥️ Custom Scripts — Device / Audio
Scripts for managing audio device configuration on endpoints.

- Detect connected audio devices
- Disable internal microphone via policy
- Rollback internal mic changes

### ⏱️ Custom Scripts — Device / Time Sync
Scripts for managing Windows time synchronization.

- Restart and force Windows Time service sync

### 🧪 Testing Scripts — SMTP
Scripts for diagnosing SMTP connectivity and authentication.

- One-time SMTP test with interactive credential prompt
- Recurring SMTP test (every 5 minutes) with saved encrypted password

### 🖼️ Custom Scripts — Intune / Desktop
Scripts and assets for managing desktop and lockscreen configuration.

- Deploy lockscreen to start and desktop
- Set corporate wallpaper via Intune (PersonalizationCSP + WinAPI + Default User)

---

## Getting Started

Run the bootstrap script once to install and import all required modules across any platform:

```powershell
# Windows / macOS / Linux
.\scripts\Startup\Install-Modules.ps1
```

Optional flags:

```powershell
.\scripts\Startup\Install-Modules.ps1 -Force              # Reinstall even if already present
.\scripts\Startup\Install-Modules.ps1 -Scope AllUsers     # Install for all users (requires elevation)
```

The script detects the current platform and skips Windows-only modules (`AzureAD`, `WindowsAutopilotIntune`) automatically on macOS and Linux.

## Requirements

- **PowerShell** 7.0 or later (cross-platform) — 5.1 is supported per-script on Windows but the bootstrap requires PS 7+
- Appropriate **Microsoft 365 admin permissions** for the target workload
- Script execution must be enabled (Windows only):

```powershell
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

---

## Usage

Each script is self-contained and includes a comment block at the top describing:

- **Purpose** – what the script does
- **Parameters** – required and optional inputs
- **Prerequisites** – required modules and permissions
- **Example** – sample usage

Run any script with `-Help` or check the header comment for usage instructions.

---

## Repository Structure

```
M365-Scripts/
├── .vscode/
│   └── settings.json
├── scripts/
│   ├── Custom Scripts/
│   │   ├── device/
│   │   │   ├── audio/
│   │   │   │   ├── detect-audiodevices.ps1
│   │   │   │   ├── Disable-internalmic.ps1
│   │   │   │   └── readme.md
│   │   │   └── Time sync/
│   │   │       └── Restart-Time-Sync.ps1
│   │   ├── Intune/Desktop/
│   │   │   ├── Add Lockscreen to start and desktop/
│   │   │   └── Background/
│   │   │       └── Desktop/
│   │   │           ├── Set-CorporateWallpaper.ps1
│   │   │           └── readme.md
│   │   └── Save install time/             ← USB setup toolkit
│   │       ├── start.bat
│   │       ├── autorun.inf
│   │       └── readme.md
│   ├── Exchange/
│   │   ├── Migrate-Calendar.ps1
│   │   ├── Set-Calendar-rights.ps1
│   │   └── readme.md
│   ├── Intune/Get-Autopilot/
│   │   ├── Get-WindowsAutoPilotInfo.ps1
│   │   └── GetAutoPilot.CMD
│   ├── Testing Scripts/SMTP/
│   │   ├── testsmtp.ps1
│   │   ├── testsmtp_5min.ps1
│   │   └── readme.md
│   └── Startup/
│       ├── functies.ps1                 ← Function library: dot-source at startup
│       ├── Install-Modules.ps1          ← Bootstrap: install & import all modules
│       └── readme.md
└── readme.md
```

---

## Contributing

This repository is actively maintained and expanded. When adding new scripts:

1. Follow the existing naming convention (`Verb-Noun.ps1`)
2. Include a header comment block with Purpose, Parameters, Prerequisites, and Example
3. Test against a non-production tenant before committing
4. Place the script in the appropriate workload folder

---

## Disclaimer

These scripts are provided as-is. Always test in a non-production environment before running against live tenants. The maintainer accepts no liability for unintended changes resulting from misuse or misconfiguration.

---

## Version History

| Date | Change |
|------|--------|
| 2026-03-20 | `start.bat` v2.8 — Split Do it all: A = Intune (Rename + Autopilot + Update), C = AD (Rename + Domain join + Update) |
| 2026-03-20 | `start.bat` v2.7 — Do it all updated: AD domain join added as step 3 |
| 2026-03-20 | `start.bat` v2.6 — Do it all updated: device rename now first step before Autopilot |
| 2026-03-20 | `start.bat` v2.5 — Added device rename option (B): prompts for prefix, auto-appends serial number |
| 2026-03-20 | `start.bat` v2.4 — Added Active Directory domain join option (8); option 9 = restart |
| 2026-03-20 | `start.bat` v2.3 — Autopilot online option (`-Online` uploads directly to Intune); Do it all uses online enrollment |
| 2026-03-20 | `start.bat` v2.2 — Windows Update via `PSWindowsUpdate` module (`Install-WindowsUpdate -AcceptAll`) |
| 2026-03-20 | `start.bat` v2.1 — added Windows Update, PowerShell 7 install (`winget`), Do it all option |
| 2026-03-20 | `start.bat` v2.0 — English, self-elevation, `cd /d %~dp0`, OOBE-compatible; added `autorun.inf` and readme |
| 2026-03-20 | Rewrote `Migrate-Calendar.ps1` to v2.0 — English, generic, mandatory params, linter fixes |
| 2026-03-20 | Removed all BraveHub references from scripts and readmes |
| 2026-03-20 | Added `Test-SmtpRelay` to `functies.ps1` |
| 2026-03-20 | SMTP scripts v2.1 — cross-platform (Windows: saved DPAPI file, macOS/Linux: prompt once) |
| 2026-03-20 | Added SMTP test scripts — `testsmtp.ps1` and `testsmtp_5min.ps1` with `param()`, `System.Net.Mail.SmtpClient` |
| 2026-03-20 | Rewrote `Set-CorporateWallpaper.ps1` to v2.0 — English, `Invoke-WebRequest`, `#Requires -Version 5.1`, generic CDN URL |
| 2026-03-20 | Added `Set-CorporateWallpaper.ps1` with Intune deployment via PersonalizationCSP |
| 2026-03-20 | Moved `Install-Modules.ps1` to `scripts/Startup/`; added `Restart-Time-Sync.ps1` |
| 2026-03-20 | Translated all readme files to English |
| 2026-03-20 | Rewrote `functies.ps1` — replaced MSOnline/AzureAD with Microsoft Graph, made generic and cross-platform |
| 2026-03-20 | Added `scripts/Startup/readme.md` |
| 2026-03-20 | Added `Set-Calendar-rights.ps1` with NL/FR/EN locale support |
| 2026-03-20 | Added `Install-Modules.ps1` cross-platform bootstrap script |
| 2026-03-20 | Updated readme structure; added device/audio and Intune/Desktop categories |
| 2026-03-19 | Added Exchange calendar migration script |
| 2026-03-19 | Added Intune/Get-Autopilot scripts (`Get-WindowsAutoPilotInfo.ps1`, `GetAutoPilot.CMD`) |
| 2026-03-19 | Initial repository upload |

---

## Maintainer

**Sjoerd Kanon** — Security-minded | Team- & Projectgericht | Microsoft 365 & Infrastructure