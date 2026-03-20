# Setup Toolkit — USB / OOBE

> Author: Sjoerd Kanon

A USB toolkit for Windows setup and Autopilot enrollment. Designed to be used during OOBE (Out-of-Box Experience) via `Shift+F10`.

---

## USB folder structure

All files must be in the **same folder** on the USB drive:

```
USB:\
├── start.bat                      ← Main menu — run this
├── GetAutoPilot.CMD               ← Autopilot enrollment script
├── Get-WindowsAutoPilotInfo.ps1   ← PowerShell module for hardware hash
└── autorun.inf                    ← USB drive label (cosmetic only)
```

> `GetAutoPilot.CMD` and `Get-WindowsAutoPilotInfo.ps1` are in `scripts/Intune/Get-Autopilot/` — copy both to the USB.

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

## Menu options

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
| `A` | **Do it all** — Autopilot online + Windows Update + restart (30s) | ✅ |
| `B` | **Rename device** — prompts for prefix, appends serial number (`PREFIX-SERIALNUMBER`) | ✅ |
| `0` | Exit | ✅ |

### Autopilot online (option 4)

Runs `Get-WindowsAutoPilotInfo.ps1 -Online` — uploads the hardware hash directly to Intune without generating a CSV file. Prompts for Microsoft 365 admin credentials. Device appears in **Intune → Devices → Enroll devices → Windows enrollment → Autopilot devices** within a few minutes.

> Windows Update Settings panel is not available in OOBE, but `UsoClient` triggers updates directly from the command line and works fine.

### PowerShell 7 (option 5)

Uses `winget install Microsoft.PowerShell`. Requires internet. If `winget` is not available (older Windows 10), the script shows the manual download URL. After install, launch with `pwsh.exe`.

### Do it all (option A)

Runs in sequence:
1. Removes existing `compHash.csv`
2. Runs Autopilot enrollment (`GetAutoPilot.CMD`)
3. Triggers Windows Update (`UsoClient StartScan` → `StartDownload` → `StartInstall`)
4. Restarts after 30 seconds (Ctrl+C to cancel)

---

## autorun.inf

Sets the USB drive label to `Setup Toolkit` when plugged in. Does **not** auto-execute anything — Windows blocks USB autorun on all modern versions (Vista+).

---

## Changelog

| Date | Version | Change |
|---|---|---|
| 2026-03-20 | 2.5 | Added device rename option (B) — prompts for prefix, appends serial number (`Get-WmiObject Win32_BIOS`), max 15 chars |
| 2026-03-20 | 2.4 | Added Active Directory domain join option (8) — `Add-Computer` via PowerShell, prompts for domain + credentials |
| 2026-03-20 | 2.3 | Added Autopilot online option (4) — `Get-WindowsAutoPilotInfo.ps1 -Online`; Do it all now uses online enrollment |
| 2026-03-20 | 2.2 | Windows Update now uses `PSWindowsUpdate` module (`Install-WindowsUpdate -AcceptAll -AutoReboot`) instead of `UsoClient` |
| 2026-03-20 | 2.1 | Added Windows Update, PowerShell 7 install (`winget`), and Do it all (`A`) option |
| 2026-03-20 | 2.0 | Rewritten to English; `cd /d %~dp0` for USB path; self-elevation; OOBE-compatible options only; quoted paths; added `autorun.inf` and readme |
