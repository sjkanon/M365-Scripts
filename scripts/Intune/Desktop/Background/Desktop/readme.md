**English** · [Nederlands](readme.nl.md) · [Français](readme.fr.md)

[M365-Scripts](../../../../../readme.md) › [scripts](../../../../readme.md) › [Intune](../../../readme.md) › [Desktop](../../readme.md) › [Background](../readme.md) › **Desktop**

# Desktop

Corporate desktop wallpaper via Intune — set it, and take it off again.

## Scripts

| Script | Description |
|--------|-------------|
| [`Set-CorporateWallpaper.ps1`](Set-CorporateWallpaper.ps1) ([docs](#set-corporatewallpaperps1)) | Download the corporate wallpaper and enforce it for all users (PersonalizationCSP, current user, Default User) |
| [`Remove-CorporateWallpaper.ps1`](Remove-CorporateWallpaper.ps1) ([docs](#remove-corporatewallpaperps1)) | Undo the corporate wallpaper for every user — only what `Set-CorporateWallpaper.ps1` wrote, the corporate lockscreen stays |

---

## Set-CorporateWallpaper.ps1

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

Edit the three variables in the `CONFIGURATION` block at the top of the script before uploading to Intune:

```powershell
$ImageUrl       = "https://your-cdn.com/CUSTOMERNAME/wallpaper.png"
$ClientName     = "CUSTOMERNAME"
$WallpaperStyle = "10"
```

Replace `CUSTOMERNAME` with the customer's name (e.g. `acme`). Everything else is fixed and does not need to be modified.

If you host the image on GitHub, use a raw file URL (`raw.githubusercontent.com`) instead of a `github.com/.../blob/...` page URL.
The script can normalize common GitHub blob/raw page URLs automatically, but using a raw URL directly is best.

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

> Note: the script runs as SYSTEM. Steps 5 (WinAPI) and 6 (HKCU) apply to the SYSTEM context, not the logged-in user. PersonalizationCSP (Step 4) and Default User (Step 7) apply correctly for all users regardless.

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
| **Detection rule** | File exists: `C:\ProgramData\Wallpapers\corporate-background-<customername>.jpg` (or `.png` if source image is PNG) |
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

## Changelog

| Date | Version | Change |
|---|---|---|
| — | 1.0 | Initial version |
| — | 1.2 | Made generic for reuse per customer |
| 2026-03-20 | 2.0 | Translated to English; `Invoke-WebRequest` replaces `WebClient`; `#Requires -Version 5.1`; generic CDN URL placeholder |
| 2026-04-14 | 2.1 | Added image signature validation and dynamic local extension (`.jpg/.png/.bmp`) to prevent invalid wallpaper files |
| 2026-04-14 | 2.2 | Added GitHub URL normalization (`github.com/.../blob/...` to `raw.githubusercontent.com`) and HTML-response guard to avoid black/empty backgrounds |
| 2026-04-14 | 2.3 | Added fallback enforcement for black backgrounds: machine wallpaper policy keys + update of all loaded user hives; `DesktopImageUrl` now uses source URL |
| 2026-04-14 | 2.4 | Fixed cleanup order bug where temporary `.download` file could be deleted before `Move-Item`, causing path-not-found error |
| 2026-04-14 | 2.5 | Added `explorer.exe` restart step so wallpaper/theme changes become visible immediately for logged-on users |
| 2026-04-14 | 2.6 | Restored generic default configuration values (`$ImageUrl`, `$ClientName`) for reusable customer deployments |
| 2026-04-14 | 2.7 | Added fail-safe backup of current wallpaper and changed replacement order so previous wallpaper stays available if update fails |

---

## Remove-CorporateWallpaper.ps1

Reverts what `Set-CorporateWallpaper.ps1` put in place, and only that. Run as SYSTEM (Intune platform script, or the uninstall command of the Win32 app).

1. **PersonalizationCSP** — removes the `DesktopImagePath`, `DesktopImageUrl` and `DesktopImageStatus` values. The key itself is removed only when it is then empty: [`Make-lockscreen.ps1`](../Lockscreen/Make-lockscreen.ps1) keeps its `LockScreen*` values in the same key
2. **`HKLM\...\Policies\System`** — removes `Wallpaper` and `WallpaperStyle`, but only when they point at a corporate wallpaper file; a policy set by something else is left alone
3. **Every loaded user hive** (`S-1-5-21-*`) — a wallpaper pointing at a corporate file is put back to the Windows default (`img0.jpg`, style Fill), and that user's transcoded wallpaper cache is cleared so the old image does not linger
4. **The current user's `HKCU`** — the same reset, but only when not running as SYSTEM (as SYSTEM, `HKCU` is SYSTEM's own profile and step 3 already covered the users)
5. **Default User profile** (`C:\Users\Default\NTUSER.DAT`) — the same reset, so new accounts no longer get the wallpaper. Skipped with a message when the hive cannot be loaded (not elevated)
6. **Files** — deletes the `corporate-background-*` files in `C:\ProgramData\Wallpapers`. The folder is removed only when it is empty — the lockscreen image lives there too
7. Restarts `explorer.exe` so the change shows immediately

"A corporate file" means a `corporate-background-*` file in `C:\ProgramData\Wallpapers` — the name `Set-CorporateWallpaper.ps1` gives it.

**Parameters**

| Parameter | Description |
|-----------|-------------|
| `-WhatIf` | Show every value and file that would be removed or reset; change nothing |

**Examples**

```powershell
# As SYSTEM (Intune platform script / Win32 uninstall command)
powershell.exe -ExecutionPolicy Bypass -File Remove-CorporateWallpaper.ps1

# See what it would do
.\Remove-CorporateWallpaper.ps1 -WhatIf
```

**Notes**

- Log: `C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\CorporateWallpaper-Remove.log`
- Users whose profile is not loaded (not signed in) keep the wallpaper value in their own hive; Windows shows the default once the file is gone, and the next run of this script while they are signed in resets the value
- Earlier versions deleted the whole PersonalizationCSP key and the whole `C:\ProgramData\Wallpapers` folder, which also removed the corporate lockscreen
