# ==============================================================================
# Set-CorporateWallpaper.ps1
# Versie 1.2 - Generieke versie voor hergebruik per klant
#
# Gebruik:
#   Pas enkel de variabelen in het blok "CONFIGURATIE" hieronder aan.
#   De rest van het script hoeft niet gewijzigd te worden.
#
# Uitrollen via Intune:
#   - Type: PowerShell script  /  Win32 app
#   - Uitvoeren als: SYSTEM
#   - 64-bit PowerShell: Ja
# ==============================================================================

# ==============================================================================
# CONFIGURATIE - pas dit aan per klant
# ==============================================================================

# URL naar de achtergrondafbeelding (PNG of JPG)
# Tip: gebruik https://config.support.bravehub.io/<KLANTNAAM>/wallpaper.png
$ImageUrl = "https://config.support.bravehub.io/KLANTNAAM/wallpaper.png"

# Weergavestijl:
#   10 = Fill (aanbevolen - vult scherm zonder vervorming)
#    6 = Fit  (past binnen scherm, zwarte randen mogelijk)
#    2 = Stretch (uitrekken, kan vervormen)
#    0 = Tile
#   22 = Span (multi-monitor)
$WallpaperStyle = "10"

# Naam van de klant - wordt gebruikt in logberichten en bestandsnaam
$ClientName = "KLANTNAAM"

# ==============================================================================
# INTERNE VARIABELEN - niet aanpassen
# ==============================================================================

$WallpaperFolder = "$env:ProgramData\Wallpapers"
$WallpaperFileName = "corporate-background-$($ClientName.ToLower()).jpg"
$WallpaperPath    = "$WallpaperFolder\$WallpaperFileName"
$LogFilePath      = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs\CorporateWallpaper-$($ClientName.ToUpper()).log"

# ==============================================================================
# FUNCTIES
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
    $timeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timeStamp - $Message" | Out-File -FilePath $LogFilePath -Append
}

# ==============================================================================
# SCRIPT START
# ==============================================================================

Write-Log "====== Start Set-CorporateWallpaper voor klant: $ClientName ======"
Write-Log "Bron-URL  : $ImageUrl"
Write-Log "Doel-pad  : $WallpaperPath"
Write-Log "Stijl     : $WallpaperStyle"

# Stap 1: Doelmap aanmaken
if (-not (Test-Path -Path $WallpaperFolder)) {
    try {
        New-Item -ItemType Directory -Path $WallpaperFolder -Force | Out-Null
        Write-Log "Wallpapermap aangemaakt: $WallpaperFolder"
    } catch {
        Write-Log "FOUT bij aanmaken map: $_"
        exit 1
    }
}

# Stap 2: Afbeelding downloaden
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $webClient = New-Object System.Net.WebClient
    $webClient.DownloadFile($ImageUrl, $WallpaperPath)
    Write-Log "Afbeelding gedownload naar: $WallpaperPath"
} catch {
    Write-Log "FOUT bij downloaden: $_"
    exit 1
}

if (-not (Test-Path -Path $WallpaperPath)) {
    Write-Log "FOUT: Bestand niet aanwezig na download."
    exit 1
}

# Stap 3: Windows API laden
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

# Stap 4: PersonalizationCSP (MDM/Intune - voor alle gebruikers)
try {
    $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP"
    if (-not (Test-Path -Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }
    New-ItemProperty -Path $regPath -Name "DesktopImagePath"   -Value $WallpaperPath -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $regPath -Name "DesktopImageUrl"    -Value $WallpaperPath -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $regPath -Name "DesktopImageStatus" -Value 1              -PropertyType DWord  -Force | Out-Null
    Write-Log "PersonalizationCSP registerwaarden bijgewerkt"
} catch {
    Write-Log "FOUT bij PersonalizationCSP: $_"
}

# Stap 5: Huidige gebruiker - WinAPI
try {
    $result = [Wallpaper]::SystemParametersInfo($SPI_SETDESKWALLPAPER, 0, $WallpaperPath, $SPIF_UPDATEINIFILE -bor $SPIF_SENDCHANGE)
    if ($result) {
        Write-Log "Achtergrond ingesteld via WinAPI voor huidige gebruiker"
    } else {
        Write-Log "WinAPI retourneerde false - geen kritieke fout"
    }
} catch {
    Write-Log "FOUT bij WinAPI: $_"
}

# Stap 6: Huidige gebruiker - register (HKCU)
try {
    $regPath = "HKCU:\Control Panel\Desktop"
    Set-ItemProperty -Path $regPath -Name "Wallpaper"      -Value $WallpaperPath -Force
    Set-ItemProperty -Path $regPath -Name "WallpaperStyle" -Value $WallpaperStyle -Force
    Set-ItemProperty -Path $regPath -Name "TileWallpaper"  -Value "0"            -Force
    RUNDLL32.EXE USER32.DLL,UpdatePerUserSystemParameters 1, $true
    Write-Log "HKCU registerwaarden bijgewerkt (stijl: $WallpaperStyle)"
} catch {
    Write-Log "FOUT bij HKCU register: $_"
}

# Stap 7: Default User profiel (voor nieuwe gebruikers)
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
        Write-Log "Default User profiel bijgewerkt voor nieuwe gebruikers"
    } else {
        Write-Log "Default User NTUSER.DAT niet gevonden - stap overgeslagen"
    }
} catch {
    Write-Log "FOUT bij Default User profiel: $_"
    try { reg unload $tempRegPath | Out-Null } catch { }
}

Write-Log "====== Script voltooid voor klant: $ClientName ======"
exit 0