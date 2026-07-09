# Make-lockscreen.ps1

> Author: Sjoerd Kanon

---

## What it does

Downloads the corporate wallpaper image from the internet and applies the same file as the Windows lockscreen via Intune.

The script writes the `PersonalizationCSP` lockscreen registry values in `HKLM`, so the lockscreen is enforced device-wide.

---

## Configuration

Edit the variables in the `CONFIGURATION` block at the top of the script before uploading to Intune:

```powershell
$ImageUrl   = "https://your-cdn.com/CUSTOMERNAME/wallpaper.png"
$ClientName = "CUSTOMERNAME"
```

Replace `CUSTOMERNAME` with the customer name you want to use in the local filename and log filename.

This script is intended to use the same image as `Set-CorporateWallpaper.ps1`, so desktop and lockscreen branding stay aligned.

If you host the image on GitHub, a raw URL is preferred. Common `github.com/.../blob/...` URLs are normalized automatically.

---

## Logging

Logs are written to:

```text
C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\CorporateLockscreen-<CUSTOMERNAME>.log
```

---

## Intune deployment

### As PowerShell script

1. Intune -> Devices -> Scripts and remediations -> Platform scripts
2. Add a new Windows 10 and later PowerShell script
3. Upload `Make-lockscreen.ps1`
4. Use these settings:
   - Run this script using the logged on credentials: No
   - Enforce script signature check: No
   - Run script in 64-bit PowerShell: Yes
5. Assign to the required device group

### As Win32 app

Use a file detection rule for the downloaded image in:

```text
C:\ProgramData\Wallpapers\corporate-lockscreen-<customername>.<jpg|png|bmp>
```

---

## How it works

| Step | Action |
|---|---|
| 1 | Create `C:\ProgramData\Wallpapers\` if it does not exist |
| 2 | Normalize common GitHub download URLs when needed |
| 3 | Download the image with `Invoke-WebRequest` |
| 4 | Validate file size and detect unsupported or HTML content |
| 5 | Detect the actual image type from file headers |
| 6 | Save the file locally with the correct extension |
| 7 | Write `LockScreenImagePath`, `LockScreenImageUrl`, and `LockScreenImageStatus` in `PersonalizationCSP` |
| 8 | Trigger a parameter refresh with `RUNDLL32.EXE USER32.DLL, UpdatePerUserSystemParameters 1, True` |

---

## Version history

| Date | Version | Change |
|---|---|---|
| — | 1.0 | Initial lockscreen version using direct `WebClient` download and static `.jpg` output |
| 2026-04-16 | 2.0 | Updated to use the same corporate wallpaper source, added GitHub/raw URL normalization, image validation, HTML detection, structured logging, and safer temporary download handling |