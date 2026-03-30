#Requires -Version 5.1
# ==============================================================================
# Set-CorporateWallpaper.ps1
# Version 2.0 - Generic version for reuse per customer
#
# Usage:
#   Only change the variables in the CONFIGURATION block below.
#   The rest of the script does not need to be modified.
#
# Deploy via Intune:
#   - Type: PowerShell script
#   - Run as: SYSTEM
#   - Run in 64-bit PowerShell: Yes
# ==============================================================================

# ==============================================================================
# CONFIGURATION — change these per customer
# ==============================================================================

# URL to the wallpaper image (PNG or JPG)
# Use a publicly accessible URL hosted by or on behalf of the customer
$ImageUrl = "https://your-cdn.com/CUSTOMERNAME/wallpaper.png"

# Display style:
#   10 = Fill    (recommended — fills screen without distortion)
#    6 = Fit     (fits within screen, black borders possible)
#    2 = Stretch (stretches to fill, may distort)
#    0 = Tile
#   22 = Span    (spreads across multiple monitors)
$WallpaperStyle = "10"

# Customer name — used in log filename and local image filename
$ClientName = "CUSTOMERNAME"

# ==============================================================================
# INTERNAL VARIABLES — do not modify
# ==============================================================================

$WallpaperFolder   = "$env:ProgramData\Wallpapers"
$WallpaperFileName = "corporate-background-$($ClientName.ToLower()).jpg"
$WallpaperPath     = "$WallpaperFolder\$WallpaperFileName"
$LogFilePath       = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs\CorporateWallpaper-$($ClientName.ToUpper()).log"

# ==============================================================================
# FUNCTIONS
# ==============================================================================

function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message
    )
    $logDir = Split-Path -Parent $LogFilePath
    if (-not (Test-Path -Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Out-File -FilePath $LogFilePath -Append
}

# ==============================================================================
# SCRIPT START
# ==============================================================================

Write-Log "====== Start Set-CorporateWallpaper for client: $ClientName ======"
Write-Log "Source URL : $ImageUrl"
Write-Log "Target path: $WallpaperPath"
Write-Log "Style      : $WallpaperStyle"

# Step 1: Create target folder
if (-not (Test-Path -Path $WallpaperFolder)) {
    try {
        New-Item -ItemType Directory -Path $WallpaperFolder -Force | Out-Null
        Write-Log "Wallpaper folder created: $WallpaperFolder"
    } catch {
        Write-Log "ERROR creating folder: $_"
        exit 1
    }
}

# Step 2: Download image
try {
    Invoke-WebRequest -Uri $ImageUrl -OutFile $WallpaperPath -UseBasicParsing
    Write-Log "Image downloaded to: $WallpaperPath"
} catch {
    Write-Log "ERROR downloading image: $_"
    exit 1
}

if (-not (Test-Path -Path $WallpaperPath)) {
    Write-Log "ERROR: File not present after download."
    exit 1
}

# Step 3: Load Windows API
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class Wallpaper {
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
}
"@ -ErrorAction SilentlyContinue

$SPI_SETDESKWALLPAPER = 0x0014
$SPIF_UPDATEINIFILE   = 0x01
$SPIF_SENDCHANGE      = 0x02

# Step 4: PersonalizationCSP (MDM/Intune — enforces wallpaper for all users)
try {
    $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP"
    if (-not (Test-Path -Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }
    New-ItemProperty -Path $regPath -Name "DesktopImagePath"   -Value $WallpaperPath -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $regPath -Name "DesktopImageUrl"    -Value $WallpaperPath -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $regPath -Name "DesktopImageStatus" -Value 1              -PropertyType DWord  -Force | Out-Null
    Write-Log "PersonalizationCSP registry keys updated"
} catch {
    Write-Log "ERROR setting PersonalizationCSP: $_"
}

# Step 5: Current user — WinAPI (applies immediately)
try {
    $result = [Wallpaper]::SystemParametersInfo($SPI_SETDESKWALLPAPER, 0, $WallpaperPath, $SPIF_UPDATEINIFILE -bor $SPIF_SENDCHANGE)
    if ($result) {
        Write-Log "Wallpaper applied via WinAPI for current user"
    } else {
        Write-Log "WinAPI returned false — non-critical, PersonalizationCSP will apply the wallpaper"
    }
} catch {
    Write-Log "ERROR applying wallpaper via WinAPI: $_"
}

# Step 6: Current user — HKCU registry (persists style setting)
try {
    $regPath = "HKCU:\Control Panel\Desktop"
    Set-ItemProperty -Path $regPath -Name "Wallpaper"      -Value $WallpaperPath -Force
    Set-ItemProperty -Path $regPath -Name "WallpaperStyle" -Value $WallpaperStyle -Force
    Set-ItemProperty -Path $regPath -Name "TileWallpaper"  -Value "0"            -Force
    RUNDLL32.EXE USER32.DLL,UpdatePerUserSystemParameters 1, $true
    Write-Log "HKCU registry keys updated (style: $WallpaperStyle)"
} catch {
    Write-Log "ERROR updating HKCU registry: $_"
}

# Step 7: Default User profile (applies to new user accounts created after deployment)
try {
    $defaultUserRegPath = "$env:SystemDrive\Users\Default\NTUSER.DAT"
    $tempRegPath        = "HKLM\DefaultUserTemp"

    if (Test-Path -Path $defaultUserRegPath) {
        reg load $tempRegPath $defaultUserRegPath | Out-Null
        reg add "$tempRegPath\Control Panel\Desktop" /v Wallpaper      /t REG_SZ /d $WallpaperPath  /f | Out-Null
        reg add "$tempRegPath\Control Panel\Desktop" /v WallpaperStyle /t REG_SZ /d $WallpaperStyle /f | Out-Null
        reg add "$tempRegPath\Control Panel\Desktop" /v TileWallpaper  /t REG_SZ /d 0              /f | Out-Null
        [gc]::Collect()
        Start-Sleep -Seconds 1
        reg unload $tempRegPath | Out-Null
        Write-Log "Default User profile updated for new user accounts"
    } else {
        Write-Log "Default User NTUSER.DAT not found — step skipped"
    }
} catch {
    Write-Log "ERROR updating Default User profile: $_"
    try { reg unload $tempRegPath | Out-Null } catch { }
}

Write-Log "====== Script completed for client: $ClientName ======"
exit 0
