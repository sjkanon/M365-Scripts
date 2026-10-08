**English** · [Nederlands](readme.nl.md) · [Français](readme.fr.md)

[M365-Scripts](../../readme.md) › [scripts](../readme.md) › **Deployment**

# Setup Toolkit — USB / OOBE

> Author: Sjoerd Kanon

A USB toolkit for Windows setup and Autopilot enrollment. Designed to be used during OOBE (Out-of-Box Experience) via `Shift+F10`.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`start.bat`](start.bat) ([docs](#startbat)) | Main menu of the toolkit — self-elevates and offers Autopilot enrollment, Windows Update, rename, domain join and the customer install browser |
| [`start.local.example.cmd`](start.local.example.cmd) ([docs](#startlocalcmd)) | Template for `start.local.cmd` — the site's LocalAdmin password and install share, kept out of the repo |
| [`Browse-InstallScripts.ps1`](Browse-InstallScripts.ps1) ([docs](#browse-installscriptsps1)) | Interactive customer/script browser behind menu options `D` and `E` — browse customer folders and launch `.ps1` / `.bat` / `.cmd` files |

Also in this folder: [`autorun.inf`](autorun.inf) — USB drive label only ([details](#autoruninf)).

---

## USB folder structure

All files must be in the **same folder** on the USB drive:

```
USB:\
├── start.bat                      ← Main menu — run this
├── start.local.cmd                ← Site settings: LocalAdmin password, install share (not in the repo)
├── GetAutoPilot.CMD               ← Autopilot enrollment script
├── Get-WindowsAutoPilotInfo.ps1   ← PowerShell module for hardware hash
├── Browse-InstallScripts.ps1       ← Customer install browser for option D/E
├── Install\                        ← Local customer install content for option D
└── autorun.inf                    ← USB drive label (cosmetic only)
```

> `GetAutoPilot.CMD` and `Get-WindowsAutoPilotInfo.ps1` are in `scripts/Intune/Get-Autopilot/` — copy both to the USB.
> For menu option `D`, copy both `Browse-InstallScripts.ps1` and the full `Install` folder to the same USB location as `start.bat`.

---

## OOBE usage

Windows does **not** auto-run USB scripts (blocked since Vista). Manual steps:

1. Plug in the USB drive
2. During OOBE, press **Shift+F10** to open a command prompt
3. Find the USB drive letter — usually `D:` or `E:`:
   ```
   D:
   dir
   ```
4. Run the toolkit:
   ```
   start.bat
   ```
5. The script requests administrator privileges automatically and shows the menu

---

## start.bat

The main menu of the toolkit. It switches to its own folder (`cd /d %~dp0`), requests administrator privileges, and shows these options:

| Option | Action | Works in OOBE |
|---|---|---|
| `1` | Open Device Manager | ✅ |
| `2` | Autopilot enrollment — save hash to `compHash.csv` | ✅ |
| `3` | Remove hash file + re-run Autopilot (save to CSV) | ✅ |
| `4` | **Autopilot enrollment online** — upload hash directly to Intune | ✅ (needs internet + admin account) |
| `5` | Windows Update via `PSWindowsUpdate` (`Install-WindowsUpdate -AcceptAll -AutoReboot`) | ✅ (needs internet) |
| `6` | Install PowerShell 7 via `winget` | ✅ (needs internet) |
| `7` | Enter product key (`slui.exe`) | ✅ |
| `8` | Join Active Directory domain | ✅ (needs domain connectivity) |
| `9` | Restart (5 second delay) | ✅ |
| `A` | **Do it all — Intune** — Rename + Autopilot online + Windows Update + restart | ✅ (needs internet + admin account) |
| `B` | **Rename device** — prompts for prefix, appends serial number (`PREFIX-SERIALNUMBER`) | ✅ |
| `C` | **Do it all — AD** — Rename + Domain join + Windows Update + restart | ✅ (needs domain connectivity) |
| `D` | **Customer install scripts (local)** — open customer menu from local `Install` folder | ✅ |
| `E` | **Customer install scripts (network share)** — open customer menu from the share in `INSTALL_SHARE` (asked when not set) | ✅ (needs network access) |
| `0` | Exit | ✅ |

### Browse-InstallScripts.ps1

Customer install browser behind menu options `D` and `E`. Shows the first-level customer folders under `-RootPath` as a menu, then lets you browse into them and launch `.ps1`, `.bat` and `.cmd` files (`AppDeployToolkit` folders are hidden).

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-RootPath` | Yes | Folder whose subfolders are the customers (local `Install` folder or a network share) |
| `-SourceLabel` | No | Title shown above the customer menu (default `Install Scripts`) |

```powershell
# What option D runs
powershell -NoProfile -ExecutionPolicy Bypass -File .\Browse-InstallScripts.ps1 -RootPath .\Install -SourceLabel "Local Install"
```

- Option `D` needs local files: `Browse-InstallScripts.ps1` and the complete `Install` folder next to `start.bat`.
- Option `E` reads customer folders from the share in `INSTALL_SHARE` (from [`start.local.cmd`](#startlocalcmd), otherwise asked) and needs network access.
- Before option `D` or `E` opens the deploy browser, `start.bat` prepares the device for deployment:
   - Creates or updates local admin user `LocalAdmin`
   - Password: `LOCALADMIN_PASSWORD` from [`start.local.cmd`](#startlocalcmd), otherwise asked with hidden input. No password, no account: the option stops and the menu returns
   - Adds `LocalAdmin` to the local `Administrators` group
   - Sets OOBE skip registry flags so the remaining OOBE flow can be skipped more easily

### start.local.cmd

Site-specific settings that do not belong in the repo. `start.bat` loads it from its own folder when it exists; copy [`start.local.example.cmd`](start.local.example.cmd) to `start.local.cmd` on the USB stick and fill it in. `start.local.cmd` is git-ignored.

| Variable | Description |
|----------|-------------|
| `LOCALADMIN_PASSWORD` | Password for the `LocalAdmin` account options `D` and `E` create. Empty or missing: asked with hidden input. Avoid `%` — batch expands it |
| `INSTALL_SHARE` | UNC path of the customer install share for option `E`, e.g. `\\server\Software`. Empty or missing: asked when option `E` is chosen |

### Autopilot online (option 4)

Runs `Get-WindowsAutoPilotInfo.ps1 -Online -DeviceCode` — uploads the hardware hash directly to Intune through Microsoft Graph without generating a CSV file. Sign in with a Microsoft 365 admin account by device code: open the address shown on a phone or another PC and enter the code, so no browser is needed during OOBE. Device appears in **Intune → Devices → Enroll devices → Windows enrollment → Autopilot devices** within a few minutes.

> Windows Update Settings panel is not available in OOBE, but `UsoClient` triggers updates directly from the command line and works fine.

### PowerShell 7 (option 5)

Uses `winget install Microsoft.PowerShell`. Requires internet. If `winget` is not available (older Windows 10), the script shows the manual download URL. After install, launch with `pwsh.exe`.

### Do it all — Intune (option A)

For Intune/cloud-managed environments. Runs in sequence:
1. Renames the device — prompts for prefix, appends serial number (`PREFIX-SERIALNUMBER`)
2. Removes existing `compHash.csv`
3. Runs Autopilot enrollment online (`Get-WindowsAutoPilotInfo.ps1 -Online -DeviceCode`)
4. Installs Windows updates via `PSWindowsUpdate`
5. Restarts after 30 seconds (Ctrl+C to cancel)

### Do it all — Active Directory (option C)

For on-premises AD environments (no Intune). Runs in sequence:
1. Renames the device — prompts for prefix, appends serial number (`PREFIX-SERIALNUMBER`)
2. Joins Active Directory domain — prompts for domain name and admin credentials
3. Installs Windows updates via `PSWindowsUpdate`
4. Restarts after 30 seconds (Ctrl+C to cancel)

---

## autorun.inf

Sets the USB drive label to `Setup Toolkit` when plugged in. Does **not** auto-execute anything — Windows blocks USB autorun on all modern versions (Vista+).

---

## Changelog

| Date | Version | Change |
|---|---|---|
| 2026-10-05 | 3.0 | The LocalAdmin password and the install share's address are no longer in `start.bat`: they come from `start.local.cmd` (git-ignored), and are asked when it is missing — the password with hidden input. With no password, options `D`/`E` stop instead of creating an account. Added `start.local.example.cmd` |
| 2026-04-17 | 2.9 | Added customer-based install browser to `start.bat`: option `D` opens local `Install` customer folders and option `E` opens a network share; added `Browse-InstallScripts.ps1` to browse customer folders and run `.ps1` / `.bat` / `.cmd` scripts; documented that option `D` requires copying both `Browse-InstallScripts.ps1` and the full `Install` folder; options `D` and `E` now create/update local admin `LocalAdmin` and set OOBE skip flags before deployment starts |

| Date | Version | Change |
|---|---|---|
| 2026-10-08 | 2.9 | Autopilot online signs in by device code (`-DeviceCode`): the script now talks to Microsoft Graph instead of the retired AzureAD/WindowsAutopilotIntune modules, and a browser sign-in may not open during OOBE |
| 2026-03-20 | 2.8 | Split Do it all into A (Intune) and C (Active Directory); AD variant skips Autopilot |
| 2026-03-20 | 2.7 | Do it all updated: AD domain join added as step 3 |
| 2026-03-20 | 2.6 | Do it all updated: device rename added as first step |
| 2026-03-20 | 2.5 | Added device rename option (B) — prompts for prefix, appends serial number (`Get-WmiObject Win32_BIOS`), max 15 chars |
| 2026-03-20 | 2.4 | Added Active Directory domain join option (8) — `Add-Computer` via PowerShell, prompts for domain + credentials |
| 2026-03-20 | 2.3 | Added Autopilot online option (4) — `Get-WindowsAutoPilotInfo.ps1 -Online`; Do it all now uses online enrollment |
| 2026-03-20 | 2.2 | Windows Update now uses `PSWindowsUpdate` module (`Install-WindowsUpdate -AcceptAll -AutoReboot`) instead of `UsoClient` |
| 2026-03-20 | 2.1 | Added Windows Update, PowerShell 7 install (`winget`), and Do it all (`A`) option |
| 2026-03-20 | 2.0 | Rewritten to English; `cd /d %~dp0` for USB path; self-elevation; OOBE-compatible options only; quoted paths; added `autorun.inf` and readme |
