#Requires -Version 5.1
# ================================================
# Generate-LicensingReport.ps1
# PowerShell launcher for the licensing report generator.
# Validates the environment, then calls the Python engine.
# ================================================

$ErrorActionPreference = "Stop"

# ── Configuration ────────────────────────────────────────────────────────────
# Set ExportDir to the folder where input files are placed and output is written.
# Create the folder structure below manually, or let this script create it on first run.
#
# Expected structure:
#   <ExportDir>\
#   ├── Import\
#   │   ├── Ingram\    ← place exactly 1 .xlsx here before running
#   │   └── Pax8\      ← place exactly 1 .csv here before running
#   ├── Archive\       ← input files are moved here automatically
#   └── Licensing_Report_YYYY-MM.xlsx
$ExportDir = "C:\Reports\Licensing"
# ─────────────────────────────────────────────────────────────────────────────

$ImportDir   = Join-Path $ExportDir "Import"
$IngramDir   = Join-Path $ImportDir "Ingram"
$Pax8Dir     = Join-Path $ImportDir "Pax8"
$ArchiveDir  = Join-Path $ExportDir "Archive"
$ScriptDir   = Split-Path -Parent $MyInvocation.MyCommand.Path
$LogFile     = Join-Path $ScriptDir "licensing_report.log"
$Period      = Get-Date -Format "yyyy-MM"
$ArchivePeriod = Join-Path $ArchiveDir $Period
$OutFile     = Join-Path $ExportDir "Licensing_Report_$Period.xlsx"

# ── Logging ──────────────────────────────────────────────────────────────────
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $LogFile -Value "$timestamp [$Level] $Message"
    Write-Host $Message
}

Write-Log "========================================"
Write-Log "Starting licensing report — $Period"

# ── Validate export directory ─────────────────────────────────────────────────
if (-not (Test-Path $ExportDir)) {
    Write-Log "Export directory not found: $ExportDir" "ERROR"
    Write-Host ""
    Write-Host "[ERROR] Export directory not found:" -ForegroundColor Red
    Write-Host "  $ExportDir"
    Write-Host ""
    Write-Host "Update the ExportDir variable at the top of this script."
    Write-Host "See log: $LogFile"
    Read-Host "Press Enter to exit"
    exit 2
}
Write-Log "Export directory: OK"

# ── Validate Python ───────────────────────────────────────────────────────────
try {
    $null = & python --version 2>&1
} catch {
    Write-Log "Python not found in PATH." "ERROR"
    Write-Host "[ERROR] Python not found in PATH." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 3
}
Write-Log "Python: OK"

# ── Create subfolders if missing ──────────────────────────────────────────────
foreach ($dir in @($IngramDir, $Pax8Dir, $ArchiveDir, $ArchivePeriod)) {
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir | Out-Null
        Write-Log "Created folder: $dir"
    }
}

# ── Locate Ingram file ────────────────────────────────────────────────────────
Write-Log "Looking for Ingram file in Import\Ingram\ ..."
$ingramFiles = Get-ChildItem -Path $IngramDir -Filter "*.xlsx" -ErrorAction SilentlyContinue

if ($ingramFiles.Count -eq 0) {
    Write-Log "No .xlsx file found in Import\Ingram\" "ERROR"
    Write-Host ""
    Write-Host "[ERROR] No Ingram file (.xlsx) found in:" -ForegroundColor Red
    Write-Host "  $IngramDir"
    Write-Host ""
    Write-Host "Place the Ingram billing export in the Ingram folder and try again."
    Read-Host "Press Enter to exit"
    exit 2
}

if ($ingramFiles.Count -gt 1) {
    Write-Log "Multiple .xlsx files found in Import\Ingram\ ($($ingramFiles.Count) files)." "ERROR"
    Write-Host ""
    Write-Host "[ERROR] Multiple .xlsx files found in Import\Ingram\" -ForegroundColor Red
    Write-Host "Ensure exactly 1 Ingram file is present:"
    $ingramFiles | ForEach-Object { Write-Host "  $($_.Name)" }
    Write-Host ""
    Write-Host "Remove or archive the extra files and try again."
    Read-Host "Press Enter to exit"
    exit 2
}

$IngramFile = $ingramFiles[0].FullName
Write-Log "Ingram file: $($ingramFiles[0].Name)"

# ── Locate Pax8 file ──────────────────────────────────────────────────────────
Write-Log "Looking for Pax8 file in Import\Pax8\ ..."
$pax8Files = Get-ChildItem -Path $Pax8Dir -Filter "*.csv" -ErrorAction SilentlyContinue

if ($pax8Files.Count -eq 0) {
    Write-Log "No .csv file found in Import\Pax8\" "ERROR"
    Write-Host ""
    Write-Host "[ERROR] No Pax8 file (.csv) found in:" -ForegroundColor Red
    Write-Host "  $Pax8Dir"
    Write-Host ""
    Write-Host "Place the Pax8 invoice export in the Pax8 folder and try again."
    Read-Host "Press Enter to exit"
    exit 2
}

if ($pax8Files.Count -gt 1) {
    Write-Log "Multiple .csv files found in Import\Pax8\ ($($pax8Files.Count) files)." "ERROR"
    Write-Host ""
    Write-Host "[ERROR] Multiple .csv files found in Import\Pax8\" -ForegroundColor Red
    Write-Host "Ensure exactly 1 Pax8 file is present:"
    $pax8Files | ForEach-Object { Write-Host "  $($_.Name)" }
    Write-Host ""
    Write-Host "Remove or archive the extra files and try again."
    Read-Host "Press Enter to exit"
    exit 2
}

$Pax8File = $pax8Files[0].FullName
Write-Log "Pax8 file: $($pax8Files[0].Name)"

# ── Confirmation ──────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  Licensing Report Generator — $Period" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Ingram : $IngramFile"
Write-Host "Pax8   : $Pax8File"
Write-Host "Output : $OutFile"
Write-Host ""

# ── Run Python engine ─────────────────────────────────────────────────────────
Write-Log "Starting Python engine..."

try {
    $output = & python "$ScriptDir\genereer_licentie_overzicht.py" --ingram "$IngramFile" --pax8 "$Pax8File" --output "$OutFile" 2>&1
    $output | ForEach-Object { Add-Content -Path $LogFile -Value $_ }
    $output | Where-Object { $_ -notmatch "PerformanceWarning|highly fragmented|frame.insert|pd.concat|newframe" } | Write-Host
    if ($LASTEXITCODE -ne 0) { throw "Exit code $LASTEXITCODE" }
} catch {
    Write-Log "Python engine failed: $_" "ERROR"
    Write-Host ""
    Write-Host "[ERROR] Report generation failed. See log:" -ForegroundColor Red
    Write-Host "  $LogFile"
    Read-Host "Press Enter to exit"
    exit 4
}

# ── Validate output ───────────────────────────────────────────────────────────
if (-not (Test-Path $OutFile)) {
    Write-Log "Output file was not created: $OutFile" "ERROR"
    Read-Host "Press Enter to exit"
    exit 5
}
Write-Log "Output file created: $OutFile"

# ── Archive input files ───────────────────────────────────────────────────────
Move-Item -Path $IngramFile -Destination $ArchivePeriod -Force
Move-Item -Path $Pax8File   -Destination $ArchivePeriod -Force
Write-Log "Input files archived to: $ArchivePeriod"
Write-Log "Report generated successfully."
Write-Log "========================================"

Write-Host ""
Write-Host "================================================" -ForegroundColor Green
Write-Host "Done! Report saved as:" -ForegroundColor Green
Write-Host "  $OutFile" -ForegroundColor Green
Write-Host "================================================" -ForegroundColor Green
Write-Host ""
Read-Host "Press Enter to exit"
exit 0
