<#
.SYNOPSIS
    Intune Win32-app install script voor Claude Desktop (machine-breed, incl. Cowork-vereiste).

.DESCRIPTION
    Bedoeld als "install command" van een Intune Win32-app (.intunewin). Wordt niet los
    uitgevoerd, maar als content meegepakt door Deploy-ClaudeDesktopIntune.ps1 in dezelfde
    map als Claude.msix (standaard bij Win32-app packaging: alle bronbestanden staan naast
    elkaar in de content-map).

    Voert uit:
      1. Enable-WindowsOptionalFeature -FeatureName VirtualMachinePlatform (indien nodig,
         vereist voor Cowork)
      2. Add-AppxProvisionedPackage voor machine-brede installatie (bereikt ook standaard-
         gebruikers zonder adminrechten; Add-AppxPackage alleen zou dat niet doen)
      3. HKLM:\SOFTWARE\Policies\Claude\disableAutoUpdates = 1, zodat de ingebouwde
         auto-updater de machine-brede provisioning niet per-user kan overschrijven —
         versiebeheer loopt voortaan via de maandelijkse Deploy-ClaudeDesktopIntune.ps1 run

    Gebruik als Intune "Install command":
        %SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -File Install-ClaudeDesktop-Intune.ps1

    Gebruik het meegeleverde Detect-ClaudeDesktop-Intune.ps1 als "Custom detection script".

.NOTES
    Logt naar %ProgramData%\ClaudeDeploy\install.log voor troubleshooting via Intune
    diagnostics / IME-logs.
#>

[CmdletBinding()]
param(
    [string]$MsixFileName = "Claude.msix"
)

$ErrorActionPreference = "Stop"

$logDir = Join-Path $env:ProgramData "ClaudeDeploy"
New-Item -ItemType Directory -Path $logDir -Force | Out-Null
$logFile = Join-Path $logDir "install.log"

function Write-Log {
    param([string]$Message)
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $Message"
    Add-Content -Path $logFile -Value $line
    Write-Host $line
}

try {
    Write-Log "=== Start Claude Desktop install ==="

    # 1. Virtual Machine Platform (vereist voor Cowork)
    $vmp = Get-WindowsOptionalFeature -Online -FeatureName VirtualMachinePlatform
    if ($vmp.State -ne "Enabled") {
        Write-Log "VirtualMachinePlatform niet ingeschakeld, wordt nu ingeschakeld..."
        Enable-WindowsOptionalFeature -Online -FeatureName VirtualMachinePlatform -All -NoRestart | Out-Null
        Write-Log "VirtualMachinePlatform ingeschakeld (herstart kan later nodig zijn)."
    } else {
        Write-Log "VirtualMachinePlatform was al ingeschakeld."
    }

    # 2. MSIX pad bepalen (naast dit script, zoals Intune Win32-content dat plaatst)
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    $msixPath = Join-Path $scriptDir $MsixFileName

    if (-not (Test-Path $msixPath)) {
        throw "MSIX niet gevonden op verwacht pad: $msixPath"
    }

    # 3. Machine-brede provisioning
    Write-Log "Provisioneren van $msixPath machine-breed..."
    Add-AppxProvisionedPackage -Online -PackagePath $msixPath -SkipLicense -Regions "all" | Out-Null
    Write-Log "Add-AppxProvisionedPackage voltooid."

    # 4. Auto-updater uitschakelen: versiebeheer verloopt via de maandelijkse Intune-content-update,
    #    niet via Claude's eigen updater (die zou de machine-brede/Cowork-registratie per-user
    #    kunnen overschrijven — zie support.claude.com "Enterprise configuration for Claude Desktop")
    $policyPath = "HKLM:\SOFTWARE\Policies\Claude"
    if (-not (Test-Path $policyPath)) {
        New-Item -Path $policyPath -Force | Out-Null
    }
    New-ItemProperty -Path $policyPath -Name "disableAutoUpdates" -PropertyType DWord -Value 1 -Force | Out-Null
    Write-Log "disableAutoUpdates=1 gezet onder $policyPath."

    Write-Log "=== Install script succesvol afgerond ==="
    exit 0
}
catch {
    Write-Log "FOUT: $($_.Exception.Message)"
    exit 1
}
