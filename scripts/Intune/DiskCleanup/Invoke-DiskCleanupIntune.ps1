#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Intune Win32-app install command: schoont C:\ op en herstart het apparaat.

.DESCRIPTION
    Dunne Intune-wrapper rond het bestaande cleanup-script Invoke-WindowsCleanup.ps1
    (scripts/Device/) — voert zelf geen cleanup-logica uit, roept alleen dat script aan met
    -Apply en herstart het apparaat direct daarna. Verwacht Invoke-WindowsCleanup.ps1 in
    dezelfde map (standaard bij Win32-app content packaging: alle bronbestanden staan naast
    elkaar) — zie readme.md in deze map voor de packaging-stappen.

    Schrijft na een geslaagde run een tijdstempel naar
    HKLM:\SOFTWARE\DiskCleanupDeploy\LastRunUtc. Detect-DiskCleanupIntune.ps1 gebruikt die
    waarde om de Win32-app periodiek (standaard elke 30 dagen) opnieuw te laten uitrollen,
    zodat de cleanup automatisch terugkerend blijft draaien zonder handmatige heruitrol.

    Gebruik als Intune "Install command":
        %SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -File Invoke-DiskCleanupIntune.ps1

    Gebruik het meegeleverde Detect-DiskCleanupIntune.ps1 als "Custom detection script".

.PARAMETER SkipDism
    Sla de DISM component store cleanup (/StartComponentCleanup /ResetBase) over. Standaard
    NIET overgeslagen — dit levert vaak de meeste ruimtewinst op, maar kan tientallen minuten
    duren. Verhoog bij gebruik zonder -SkipDism de Intune install-timeout (standaard 60 min)
    als dit device een trage schijf heeft of lang niet is opgeschoond.

.PARAMETER NoRestart
    Voer de cleanup uit maar sla de herstart over. Alleen voor handmatig testen buiten Intune —
    laat dit weg in de echte Intune install command.

.NOTES
    Logt naar %ProgramData%\DiskCleanupDeploy\cleanup.log (incl. volledige output van
    Invoke-WindowsCleanup.ps1 en het CSV-rapport van die run).
#>

[CmdletBinding()]
param(
    [switch]$SkipDism,
    [switch]$NoRestart
)

$ErrorActionPreference = "Stop"

$logDir = Join-Path $env:ProgramData "DiskCleanupDeploy"
New-Item -ItemType Directory -Path $logDir -Force | Out-Null
$logFile = Join-Path $logDir "cleanup.log"

function Write-Log {
    param([string]$Message)
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $Message"
    Add-Content -Path $logFile -Value $line
    Write-Host $line
}

try {
    Write-Log "=== Start Intune disk cleanup (C:) ==="

    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    $cleanupScript = Join-Path $scriptDir "Invoke-WindowsCleanup.ps1"
    if (-not (Test-Path $cleanupScript)) {
        throw "Invoke-WindowsCleanup.ps1 niet gevonden op verwacht pad: $cleanupScript"
    }

    $cleanupParams = @{ Apply = $true; OutputPath = $logDir }
    if ($SkipDism) { $cleanupParams['SkipDism'] = $true }

    Write-Log ("Cleanup starten via {0} (SkipDism={1})..." -f $cleanupScript, [bool]$SkipDism)
    & $cleanupScript @cleanupParams *>> $logFile
    Write-Log "Cleanup script afgerond."

    $regPath = "HKLM:\SOFTWARE\DiskCleanupDeploy"
    if (-not (Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
    New-ItemProperty -Path $regPath -Name "LastRunUtc" -PropertyType String `
        -Value ((Get-Date).ToUniversalTime().ToString("o")) -Force | Out-Null
    Write-Log "Detectie-tijdstempel geschreven onder $regPath."

    if ($NoRestart) {
        Write-Log "=== Cleanup afgerond, herstart overgeslagen (-NoRestart) ==="
        exit 0
    }

    Write-Log "Cleanup afgerond, apparaat wordt nu direct herstart..."
    Restart-Computer -Force
    exit 0
}
catch {
    Write-Log "FOUT: $($_.Exception.Message)"
    exit 1
}
