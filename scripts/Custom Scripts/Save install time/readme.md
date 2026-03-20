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
| `2` | Autopilot enrollment | ✅ |
| `3` | Remove hash file + re-run Autopilot | ✅ |
| `4` | Enter product key (`slui.exe`) | ✅ |
| `5` | Restart (5 second delay) | ✅ |
| `0` | Exit | ✅ |

> Windows Update settings are **not** available during OOBE — open them after the first login.

---

## autorun.inf

Sets the USB drive label to `Setup Toolkit` when plugged in. Does **not** auto-execute anything — Windows blocks USB autorun on all modern versions (Vista+).

---

## Changelog

| Date | Version | Change |
|---|---|---|
| 2026-03-20 | 2.0 | Rewritten to English; `cd /d %~dp0` for USB path; self-elevation; OOBE-compatible options only; quoted paths; added `autorun.inf` and readme |
