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
