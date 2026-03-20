# Set-CorporateWallpaper.ps1

> **BraveHub Internal Script**
> Author: Sjoerd Kanon

---

## What it does

Downloads a corporate wallpaper image from a URL and applies it to Windows devices via Intune. Sets the wallpaper for:

- The **current user** (WinAPI — applies immediately)
- The **current user** (HKCU registry — persists the style setting)
- **All users** via MDM (PersonalizationCSP — enforced via HKLM)
- **New user accounts** via the Default User profile (NTUSER.DAT)

This ensures the wallpaper is applied regardless of who logs in, both for existing and newly created accounts.

---

## Configuration

Only the three parameters at the top of the script need to be changed per customer. Everything else is fixed.

```powershell
.\Set-CorporateWallpaper.ps1 `
    -ImageUrl      "https://cdn.example.com/acme/wallpaper.png" `
    -ClientName    "Acme" `
    -WallpaperStyle "10"
```

**Intune does not pass parameters.** When deploying as a Platform Script, edit the defaults directly in the `param()` block before uploading:

```powershell
[string]$ImageUrl       = "https://your-cdn.com/CUSTOMERNAME/wallpaper.png"
[string]$ClientName     = "CUSTOMERNAME"
[string]$WallpaperStyle = "10"
```

### Wallpaper styles

| Value | Style | Notes |
|---|---|---|
| `10` | Fill | Recommended — fills the screen without distortion |
| `6` | Fit | Fits within the screen, black borders possible |
| `2` | Stretch | Stretches to fill, may distort |
| `0` | Tile | Repeats the image |
| `22` | Span | Spreads across multiple monitors |

---

## Logging

Logs are written to the standard Intune log directory:

```
C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\CorporateWallpaper-<CUSTOMERNAME>.log
```

Readable from Intune Device Diagnostics or locally on the device.

---

## Intune deployment

### As PowerShell script

1. Intune → **Devices** → **Scripts and remediations** → **Platform scripts**
2. Click **Add** → **Windows 10 and later**
3. Settings:
   - **Script**: upload `Set-CorporateWallpaper.ps1`
   - **Run this script using the logged on credentials**: No
   - **Enforce script signature check**: No
   - **Run script in 64-bit PowerShell**: Yes
4. Assign to the desired device group
5. **Save**

> Note: with this method the script runs as SYSTEM. Step 5 (WinAPI) and Step 6 (HKCU) will apply to the SYSTEM context, not the logged-in user. PersonalizationCSP (Step 4) and Default User (Step 7) will still apply correctly for all users.

### As Win32 app (recommended)

Packaging as a Win32 app allows re-run control and detection rules.

**Wrap the script:**
```powershell
.\IntuneWinAppUtil.exe -c . -s Set-CorporateWallpaper.ps1 -o .
```

**Intune app settings:**

| Field | Value |
|---|---|
| **Install command** | `powershell.exe -ExecutionPolicy Bypass -File Set-CorporateWallpaper.ps1` |
| **Uninstall command** | `cmd.exe /c echo uninstall` |
| **Detection rule** | File exists: `C:\ProgramData\Wallpapers\corporate-background-<customername>.jpg` |
| **Run as** | System |
| **Architecture** | 64-bit |

---

## How it works (step by step)

| Step | Action | Scope |
|---|---|---|
| 1 | Create `C:\ProgramData\Wallpapers\` if it does not exist | — |
| 2 | Download the image via `Invoke-WebRequest` | — |
| 3 | Load `user32.dll` Windows API | — |
| 4 | Write `PersonalizationCSP` registry keys (HKLM) | All users via MDM |
| 5 | Call `SystemParametersInfo` (WinAPI) | Current user — immediate |
| 6 | Write `HKCU\Control Panel\Desktop` + `RUNDLL32` refresh | Current user — persistent |
| 7 | Load `Default\NTUSER.DAT` and write the same keys | New user accounts |

---

## Dry run

Use `-WhatIf` to preview all actions without making any changes:

```powershell
.\Set-CorporateWallpaper.ps1 -ImageUrl "https://..." -ClientName "Acme" -WhatIf
```

---

## Changelog

| Date | Version | Change |
|---|---|---|
| — | 1.0 | Initial version |
| — | 1.2 | Made generic for reuse per customer |
| 2026-03-20 | 2.0 | Rewritten to English; `[CmdletBinding]`, `param()`, `-WhatIf`, `#Requires -Version 5.1`; `Invoke-WebRequest` replaces `WebClient`; comment-based help |
