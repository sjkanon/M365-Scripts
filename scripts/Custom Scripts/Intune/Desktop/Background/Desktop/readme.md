# Set-CorporateWallpaper.ps1

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
