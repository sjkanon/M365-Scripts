# ============================================================================
# Deploy Office Theme
# ============================================================================

$Source = "C:\temp\VIAS\hosting\2026 Vias institute colours.thmx"

$DestinationFolder = Join-Path $env:APPDATA "Microsoft\Templates\Document Themes"
$DestinationFile = Join-Path $DestinationFolder "2026 Vias institute colours.thmx"

# Maak de doelmap aan indien nodig
if (!(Test-Path $DestinationFolder)) {
    New-Item -ItemType Directory -Path $DestinationFolder -Force | Out-Null
}

# Controleer of het bronbestand bestaat
if (Test-Path $Source) {
    Copy-Item -Path $Source -Destination $DestinationFile -Force
    Write-Output "Office theme succesvol gedeployed."
} else {
    Write-Error "Bronbestand niet gevonden: $Source"
    exit 1
}
