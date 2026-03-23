# ================================================
# Licentie Overzicht Generator
# ================================================

$ErrorActionPreference = "Stop"

# ── Vaste mappen ──
$ExportDir   = "C:\OneDrive\BraveHub\BraveHub - Finance - Licenses_facturatie_upload"
$ImportDir   = Join-Path $ExportDir "Import"
$IngramDir   = Join-Path $ImportDir "Ingram"
$Pax8Dir     = Join-Path $ImportDir "Pax8"
$ArchiveDir  = Join-Path $ExportDir "Archive"
$ScriptDir   = Split-Path -Parent $MyInvocation.MyCommand.Path
$LogFile     = Join-Path $ScriptDir "licentie_rapport.log"
$Maand       = Get-Date -Format "yyyy-MM"
$ArchiveMaand = Join-Path $ArchiveDir $Maand
$OutFile     = Join-Path $ExportDir "Licentie_Overzicht_$Maand.xlsx"

# ── Logging ──
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "$timestamp [$Level] $Message"
    Add-Content -Path $LogFile -Value $line
    Write-Host $Message
}

Write-Log "========================================"
Write-Log "Start licentie rapport generatie - $Maand"

# ── Controleer OneDrive ──
if (-not (Test-Path $ExportDir)) {
    Write-Log "OneDrive map niet bereikbaar: $ExportDir" "FOUT"
    Write-Host ""
    Write-Host "[FOUT] OneDrive map niet bereikbaar." -ForegroundColor Red
    Write-Host "Controleer of je ingelogd bent en OneDrive gesynchroniseerd is."
    Write-Host "Zie logfile: $LogFile"
    Read-Host "Druk op Enter om af te sluiten"
    exit 2
}
Write-Log "OneDrive map bereikbaar: OK"

# ── Controleer Python ──
try {
    $null = & python --version 2>&1
} catch {
    Write-Log "Python niet gevonden." "FOUT"
    Write-Host "[FOUT] Python niet gevonden." -ForegroundColor Red
    Read-Host "Druk op Enter om af te sluiten"
    exit 3
}
Write-Log "Python beschikbaar: OK"

# ── Maak mappen aan ──
foreach ($dir in @($IngramDir, $Pax8Dir, $ArchiveDir, $ArchiveMaand)) {
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir | Out-Null
        Write-Log "Map aangemaakt: $dir"
    }
}

# ── Ingram bestand detecteren ──
Write-Log "Zoeken naar Ingram bestand in Import\Ingram\ ..."
$ingramFiles = Get-ChildItem -Path $IngramDir -Filter "*.xlsx" -ErrorAction SilentlyContinue

if ($ingramFiles.Count -eq 0) {
    Write-Log "Geen Excel bestand gevonden in Import\Ingram\" "FOUT"
    Write-Host ""
    Write-Host "[FOUT] Geen Ingram bestand (.xlsx) gevonden in:" -ForegroundColor Red
    Write-Host "  $IngramDir"
    Write-Host ""
    Write-Host "Plaats het Ingram billing bestand in de Ingram map en probeer opnieuw."
    Read-Host "Druk op Enter om af te sluiten"
    exit 2
}

if ($ingramFiles.Count -gt 1) {
    Write-Log "Meerdere Excel bestanden gevonden in Import\Ingram\ ($($ingramFiles.Count) bestanden)." "FOUT"
    Write-Host ""
    Write-Host "[FOUT] Meerdere Excel bestanden gevonden in Import\Ingram\" -ForegroundColor Red
    Write-Host "Zorg dat er precies 1 Ingram bestand aanwezig is:"
    $ingramFiles | ForEach-Object { Write-Host "  $($_.Name)" }
    Write-Host ""
    Write-Host "Verwijder of verplaats de overbodige bestanden en probeer opnieuw."
    Read-Host "Druk op Enter om af te sluiten"
    exit 2
}

$IngramFile = $ingramFiles[0].FullName
Write-Log "Ingram bestand gevonden: $($ingramFiles[0].Name)"

# ── Pax8 bestand detecteren ──
Write-Log "Zoeken naar Pax8 bestand in Import\Pax8\ ..."
$pax8Files = Get-ChildItem -Path $Pax8Dir -Filter "*.csv" -ErrorAction SilentlyContinue

if ($pax8Files.Count -eq 0) {
    Write-Log "Geen CSV bestand gevonden in Import\Pax8\" "FOUT"
    Write-Host ""
    Write-Host "[FOUT] Geen Pax8 bestand (.csv) gevonden in:" -ForegroundColor Red
    Write-Host "  $Pax8Dir"
    Write-Host ""
    Write-Host "Plaats het Pax8 factuur CSV bestand in de Pax8 map en probeer opnieuw."
    Read-Host "Druk op Enter om af te sluiten"
    exit 2
}

if ($pax8Files.Count -gt 1) {
    Write-Log "Meerdere CSV bestanden gevonden in Import\Pax8\ ($($pax8Files.Count) bestanden)." "FOUT"
    Write-Host ""
    Write-Host "[FOUT] Meerdere CSV bestanden gevonden in Import\Pax8\" -ForegroundColor Red
    Write-Host "Zorg dat er precies 1 Pax8 bestand aanwezig is:"
    $pax8Files | ForEach-Object { Write-Host "  $($_.Name)" }
    Write-Host ""
    Write-Host "Verwijder of verplaats de overbodige bestanden en probeer opnieuw."
    Read-Host "Druk op Enter om af te sluiten"
    exit 2
}

$Pax8File = $pax8Files[0].FullName
Write-Log "Pax8 bestand gevonden: $($pax8Files[0].Name)"

# ── Bevestiging ──
Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  Licentie Overzicht Generator - $Maand" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Ingram : $IngramFile"
Write-Host "Pax8   : $Pax8File"
Write-Host "Output : $OutFile"
Write-Host ""

# ── Python script uitvoeren ──
Write-Log "Python script starten..."

try {
    $output = & python "$ScriptDir\genereer_licentie_overzicht.py" --ingram "$IngramFile" --pax8 "$Pax8File" --output "$OutFile" 2>&1
    $output | ForEach-Object { Add-Content -Path $LogFile -Value $_ }
    $output | Where-Object { $_ -notmatch "PerformanceWarning|highly fragmented|frame.insert|pd.concat|newframe" } | Write-Host
    if ($LASTEXITCODE -ne 0) { throw "Exit code $LASTEXITCODE" }
} catch {
    Write-Log "Python script gefaald: $_" "FOUT"
    Write-Host ""
    Write-Host "[FOUT] Rapport aanmaken mislukt. Zie logfile:" -ForegroundColor Red
    Write-Host "  $LogFile"
    Read-Host "Druk op Enter om af te sluiten"
    exit 4
}

# ── Controleer output ──
if (-not (Test-Path $OutFile)) {
    Write-Log "Output bestand niet aangemaakt: $OutFile" "FOUT"
    Read-Host "Druk op Enter om af te sluiten"
    exit 5
}
Write-Log "Output bestand aangemaakt: $OutFile"

# ── Archiveren ──
Move-Item -Path $IngramFile -Destination $ArchiveMaand -Force
Move-Item -Path $Pax8File   -Destination $ArchiveMaand -Force
Write-Log "Importbestanden gearchiveerd naar: $ArchiveMaand"
Write-Log "Rapport succesvol aangemaakt."
Write-Log "========================================"

Write-Host ""
Write-Host "================================================" -ForegroundColor Green
Write-Host "Klaar! Rapport opgeslagen als:" -ForegroundColor Green
Write-Host "  $OutFile" -ForegroundColor Green
Write-Host "================================================" -ForegroundColor Green
Write-Host ""
Read-Host "Druk op Enter om af te sluiten"
exit 0