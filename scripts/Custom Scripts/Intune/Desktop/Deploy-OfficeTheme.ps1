# ==============================================================================
# Deploy Office Theme
# FirstIT
# ==============================================================================

# CONFIGURATION
$ThemeUrl  = "https://raw.githubusercontent.com/FirstITHub/M365-Scripts/refs/heads/main/scripts/Custom%20Scripts/Intune/Desktop/2026%20Vias%20institute%20colours%20(2).thmx"
$ThemeName = "2026 Vias institute colours (2).thmx"

# Local storage
$LocalFolder = "$env:ProgramData\FirstIT\OfficeThemes"
$LocalFile   = Join-Path $LocalFolder $ThemeName

# Office Theme folder
$OfficeThemeFolder = Join-Path $env:APPDATA "Microsoft\Templates\Document Themes"
$OfficeThemeFile   = Join-Path $OfficeThemeFolder $ThemeName

# Create folders
New-Item -ItemType Directory -Path $LocalFolder -Force | Out-Null
New-Item -ItemType Directory -Path $OfficeThemeFolder -Force | Out-Null

# Download theme
Invoke-WebRequest -Uri $ThemeUrl -OutFile $LocalFile

# Copy to Office Themes
Copy-Item -Path $LocalFile -Destination $OfficeThemeFile -Force

Write-Host "Office Theme deployed successfully."
