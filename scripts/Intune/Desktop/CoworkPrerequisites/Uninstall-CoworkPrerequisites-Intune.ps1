<#
.SYNOPSIS
    Intune Win32-app uninstall script voor de Cowork Windows-vereisten.

.DESCRIPTION
    Schakelt VirtualMachinePlatform en Fast Startup NIET standaard terug, omdat andere
    applicaties (WSL, Hyper-V-based tools, andere Cowork-achtige apps) hier ook van afhankelijk
    kunnen zijn, en Fast Startup uitgeschakeld laten sowieso geen nadeel heeft voor apparaten
    zonder Cowork. Gebruik -DisableVirtualMachinePlatform als je zeker weet dat dit device het
    feature nergens anders voor nodig heeft.

    Gebruik als Intune "Uninstall command":
        %SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -File Uninstall-CoworkPrerequisites-Intune.ps1

.PARAMETER DisableVirtualMachinePlatform
    Schakelt ook het Windows-feature VirtualMachinePlatform uit. Standaard uitgeschakeld
    gelaten (aanbevolen), tenzij je zeker weet dat niets anders op dit apparaat dit nodig heeft.

.NOTES
    Logt naar %ProgramData%\CoworkPrereqDeploy\uninstall.log
#>

[CmdletBinding()]
param(
    [switch]$DisableVirtualMachinePlatform
)

$ErrorActionPreference = "Stop"

$logDir = Join-Path $env:ProgramData "CoworkPrereqDeploy"
New-Item -ItemType Directory -Path $logDir -Force | Out-Null
$logFile = Join-Path $logDir "uninstall.log"

function Write-Log {
    param([string]$Message)
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $Message"
    Add-Content -Path $logFile -Value $line
    Write-Host $line
}

try {
    Write-Log "=== Start Cowork Prerequisites uninstall ==="

    if ($DisableVirtualMachinePlatform) {
        Write-Log "Uitschakelen VirtualMachinePlatform (op verzoek via -DisableVirtualMachinePlatform)..."
        Disable-WindowsOptionalFeature -Online -FeatureName VirtualMachinePlatform -NoRestart | Out-Null
        Write-Log "VirtualMachinePlatform uitgeschakeld (herstart kan nodig zijn)."
    } else {
        Write-Log "VirtualMachinePlatform blijft ingeschakeld (standaardgedrag, gebruik -DisableVirtualMachinePlatform om ook uit te schakelen)."
    }

    Write-Log "Fast Startup (HiberbootEnabled) blijft uitgeschakeld staan — dit is een onschadelijke, apparaatbrede instelling die niet specifiek is voor Cowork."

    Write-Log "=== Uninstall script succesvol afgerond ==="
    exit 0
}
catch {
    Write-Log "FOUT: $($_.Exception.Message)"
    exit 1
}
