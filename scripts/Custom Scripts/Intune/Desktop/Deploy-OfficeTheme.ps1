<#
.SYNOPSIS
    Download the corporate Office theme (.thmx) and apply it for the signed-in user.

.DESCRIPTION
    Fetches the theme from the URL pinned at the top of this script into
    %ProgramData%\OfficeThemes and points Office at it, so Word, Excel and PowerPoint
    open in the corporate colours. Intended for Intune deployment.

    The download URL is hardcoded to a raw GitHub path inside this repository, which
    is why this script lives under "Custom Scripts" - moving or renaming the .thmx
    file breaks it.
#>

# URL van de theme
$ThemeUrl = "https://github.com/FirstITHub/M365-Scripts/raw/refs/heads/main/scripts/Custom%20Scripts/Intune/Desktop/2026%20Vias%20institute%20colours%20(2).thmx"

# Naam
$ThemeName = "2026 Vias institute colours (2).thmx"

# Lokale opslag
$LocalFolder = "$env:ProgramData\OfficeThemes"
$LocalFile = Join-Path $LocalFolder $ThemeName

# Office Theme map
$OfficeFolder = Join-Path $env:APPDATA "Microsoft\Templates\Document Themes"

# Mappen aanmaken
New-Item -ItemType Directory -Path $LocalFolder -Force | Out-Null
New-Item -ItemType Directory -Path $OfficeFolder -Force | Out-Null

# Download theme
Invoke-WebRequest -Uri $ThemeUrl -OutFile $LocalFile -UseBasicParsing

# Kopiëren naar Office
Copy-Item -Path $LocalFile -Destination $OfficeFolder -Force

Write-Output "Office Theme deployed successfully."
