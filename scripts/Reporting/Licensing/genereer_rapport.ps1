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
$ExportDir = "C:\OneDrive\BraveHub\BraveHub - Finance - Licenses_facturatie_upload"
# ─────────────────────────────────────────────────────────────────────────────

$ImportDir   = Join-Path $ExportDir "Import"
$IngramDir   = Join-Path $ImportDir "Ingram"
$Pax8Dir     = Join-Path $ImportDir "Pax8"
$ArchiveDir  = Join-Path $ExportDir "Archive"
$ScriptDir   = Split-Path -Parent $MyInvocation.MyCommand.Path
$LogFile     = Join-Path $ScriptDir "licensing_report.log"
$PythonDetailLog = Join-Path $ScriptDir "python_engine_last_run.log"
$PythonAppLog    = Join-Path (Join-Path $ScriptDir "Log") "licensing_report.log"
$Period      = Get-Date -Format "yyyy-MM"
$ArchivePeriod = Join-Path $ArchiveDir $Period
$OutFile     = $null

# ── Logging ──────────────────────────────────────────────────────────────────
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $LogFile -Value "$timestamp [$Level] $Message"
    Write-Host $Message
}

function Pause-IfInteractive {
    if ([Environment]::UserInteractive) {
        Read-Host "Press Enter to exit"
    }
}

Write-Log "========================================"
Write-Log "Starting licensing report — $Period"
if (Test-Path $PythonDetailLog) {
    Remove-Item -Path $PythonDetailLog -Force -ErrorAction SilentlyContinue
}

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

# ── Validate Python modules ──────────────────────────────────────────────────
try {
    $null = & python -c "import pandas, openpyxl" 2>&1
    if ($LASTEXITCODE -ne 0) { throw "Missing required Python packages" }
} catch {
    Write-Log "Required Python packages missing (pandas/openpyxl)." "ERROR"
    Write-Host ""
    Write-Host "[ERROR] Python packages ontbreken: pandas en/of openpyxl" -ForegroundColor Red
    Write-Host "Installeer met: pip install pandas openpyxl"
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 3
}
Write-Log "Python packages: OK"

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
    Write-Log "No .xlsx file found in Import\Ingram\ - continuing without Ingram data." "WARN"
    $IngramFile = $null
} elseif ($ingramFiles.Count -gt 1) {
    Write-Log "Multiple .xlsx files found in Import\Ingram\ ($($ingramFiles.Count) files)." "ERROR"
    Write-Host ""
    Write-Host "[ERROR] Multiple .xlsx files found in Import\Ingram\" -ForegroundColor Red
    Write-Host "Ensure exactly 1 Ingram file is present:"
    $ingramFiles | ForEach-Object { Write-Host "  $($_.Name)" }
    Write-Host ""
    Write-Host "Remove or archive the extra files and try again."
    Read-Host "Press Enter to exit"
    exit 2
} else {
    $IngramFile = $ingramFiles[0].FullName
    Write-Log "Ingram file: $($ingramFiles[0].Name)"
}

# ── Locate Pax8 file ──────────────────────────────────────────────────────────
Write-Log "Looking for Pax8 file in Import\Pax8\ ..."
$pax8Files = Get-ChildItem -Path $Pax8Dir -Filter "*.csv" -ErrorAction SilentlyContinue

if ($pax8Files.Count -eq 0) {
    Write-Log "No .csv file found in Import\Pax8\ - continuing without Pax8 data." "WARN"
    $Pax8File = $null
} elseif ($pax8Files.Count -gt 1) {
    Write-Log "Multiple .csv files found in Import\Pax8\ ($($pax8Files.Count) files)." "ERROR"
    Write-Host ""
    Write-Host "[ERROR] Multiple .csv files found in Import\Pax8\" -ForegroundColor Red
    Write-Host "Ensure exactly 1 Pax8 file is present:"
    $pax8Files | ForEach-Object { Write-Host "  $($_.Name)" }
    Write-Host ""
    Write-Host "Remove or archive the extra files and try again."
    Read-Host "Press Enter to exit"
    exit 2
} else {
    $Pax8File = $pax8Files[0].FullName
    Write-Log "Pax8 file: $($pax8Files[0].Name)"
}

if (-not $IngramFile -and -not $Pax8File) {
    Write-Log "No input files found in Import\Ingram\ or Import\Pax8\" "ERROR"
    Write-Host ""
    Write-Host "[ERROR] Geen invoerbestand gevonden." -ForegroundColor Red
    Write-Host "Plaats minimaal 1 bestand in één van deze mappen:"
    Write-Host "  $IngramDir"
    Write-Host "  $Pax8Dir"
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 2
}

$sourceLabel = if ($IngramFile -and $Pax8File) { "Ingram-Pax8" }
               elseif ($IngramFile) { "IngramOnly" }
               else { "Pax8Only" }

$isPreliminary = -not ($IngramFile -and $Pax8File)
$nameSuffix = if ($isPreliminary) { "_Voorlopig" } else { "" }

$OutFile = Join-Path $ExportDir "Licensing_Report_${Period}_${sourceLabel}${nameSuffix}.xlsx"

# ── Confirmation ──────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  Licensing Report Generator — $Period" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Ingram : $(if ($IngramFile) { $IngramFile } else { 'niet gevonden - overgeslagen' })"
Write-Host "Pax8   : $(if ($Pax8File) { $Pax8File } else { 'niet gevonden - overgeslagen' })"
Write-Host "Output : $OutFile"
Write-Host ""

# ── Run Python engine ─────────────────────────────────────────────────────────
Write-Log "Starting Python engine..."

try {
    $PythonScript = Join-Path $ScriptDir "genereer_licentie_overzicht.py"
    if (-not (Test-Path $PythonScript)) {
        throw "Python script not found: $PythonScript"
    }

    $pythonArgs = @("$PythonScript", "--output", "$OutFile")
    if ($IngramFile) { $pythonArgs += @("--ingram", "$IngramFile") }
    if ($Pax8File)   { $pythonArgs += @("--pax8", "$Pax8File") }

    $output = & python @pythonArgs 2>&1
    $exitCode = $LASTEXITCODE
    $outputLines = @($output | ForEach-Object { $_.ToString() })

    $outputLines | Set-Content -Path $PythonDetailLog -Encoding UTF8
    $outputLines | ForEach-Object { Add-Content -Path $LogFile -Value $_ }
    $outputLines | Where-Object { $_ -notmatch "PerformanceWarning|highly fragmented|frame.insert|pd.concat|newframe" } | Write-Host

    if ($exitCode -ne 0) {
        $tail = ($outputLines | Select-Object -Last 8) -join [Environment]::NewLine
        if ($tail) {
            Write-Log "Python error output (last lines):`n$tail" "ERROR"
        }
        throw "Exit code $exitCode"
    }
} catch {
    Write-Log "Python engine failed: $_" "ERROR"
    Write-Host ""
    Write-Host "[ERROR] Report generation failed. See log:" -ForegroundColor Red
    Write-Host "  Launcher log : $LogFile"
    Write-Host "  Python run   : $PythonDetailLog"
    Write-Host "  Python app   : $PythonAppLog"

    if (Test-Path $PythonDetailLog) {
        Write-Host ""
        Write-Host "Last Python output lines:" -ForegroundColor Yellow
        Get-Content -Path $PythonDetailLog -Tail 25 | ForEach-Object { Write-Host "  $_" }
    }

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
if ($IngramFile -and $Pax8File) {
    Move-Item -Path $IngramFile -Destination $ArchivePeriod -Force
    Move-Item -Path $Pax8File   -Destination $ArchivePeriod -Force
    Write-Log "Input files archived to: $ArchivePeriod"
} else {
    Write-Log "Archive skipped: waiting for complete Ingram + Pax8 set for full overview."
}
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
