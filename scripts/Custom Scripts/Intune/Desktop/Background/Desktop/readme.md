# Set-CorporateWallpaper.ps1

> **BraveHub Internal Script**
> Author: Sjoerd Kanon

---

## What it does

Downloads a corporate wallpaper image from a URL and applies it to Windows devices via Intune. Sets the wallpaper for:

- The **current user** (via WinAPI and HKCU registry)
- **All users** via MDM (PersonalizationCSP — HKLM)
- **New users** via the Default User profile (NTUSER.DAT)

This ensures the wallpaper is applied regardless of who logs in, both existing and new accounts.

---

## Configuration

Only the three variables in the `CONFIGURATIE` block at the top of the script need to be changed per customer:

```powershell
$ImageUrl       = "https://your-cdn.com/CUSTOMERNAME/wallpaper.png"
$WallpaperStyle = "10"
$ClientName     = "CUSTOMERNAME"
```

Everything else is fixed and does not need to be modified.

### Wallpaper styles

| Value | Style | Notes |
|---|---|---|
| `10` | Fill | Recommended — fills the screen without distortion |
| `6` | Fit | Fits within screen, black borders possible |
| `2` | Stretch | Stretches to fill, may distort |
| `0` | Tile | Repeats the image |
| `22` | Span | Spreads across multiple monitors |

---

## Logging

Logs are written to:

```
C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\CorporateWallpaper-<CUSTOMERNAME>.log
```

This is the standard Intune log location, readable directly from Intune or locally.

---

## Intune deployment

### As PowerShell script

1. Intune → **Devices** → **Scripts and remediations** → **Platform scripts**
2. Click **Add** → **Windows 10 and later**
3. Settings:
   - **Script**: upload `Set-CorporateWallpaper.ps1`
   - **Run this script using the logged on credentials**: No (run as SYSTEM)
   - **Enforce script signature check**: No
   - **Run script in 64-bit PowerShell**: Yes
4. Assign to the desired device group
5. **Save**

### As Win32 app (recommended for re-run control)

Package as `.intunewin` and use the following:

| Field | Value |
|---|---|
| **Install command** | `powershell.exe -ExecutionPolicy Bypass -File Set-CorporateWallpaper.ps1` |
| **Uninstall command** | `cmd.exe /c echo uninstall` |
| **Detection rule** | File exists: `C:\ProgramData\Wallpapers\corporate-background-<customername>.jpg` |
| **Run as** | System |
| **Architecture** | 64-bit |

---

## How it works (step by step)

| Step | Action |
|---|---|
| 1 | Create `C:\ProgramData\Wallpapers\` if it does not exist |
| 2 | Download the image via TLS 1.2 to the local folder |
| 3 | Load the Windows `user32.dll` API for wallpaper control |
| 4 | Write `PersonalizationCSP` registry keys (HKLM) — applies via MDM to all users |
| 5 | Call `SystemParametersInfo` (WinAPI) — applies immediately for the current user |
| 6 | Write `HKCU\Control Panel\Desktop` — sets style and path for the current user |
| 7 | Load `Default\NTUSER.DAT` and write the same keys — applies to new user accounts |

---

## Changelog

| Date | Version | Change |
|---|---|---|
| — | 1.0 | Initial version |
| — | 1.2 | Made generic for reuse per customer |
