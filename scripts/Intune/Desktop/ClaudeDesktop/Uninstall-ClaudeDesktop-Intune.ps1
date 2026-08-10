<#
.SYNOPSIS
    Intune Win32-app uninstall script voor Claude Desktop.

.DESCRIPTION
    Verwijdert Claude Desktop dat via Add-AppxProvisionedPackage machine-breed is
    geïnstalleerd, en ruimt als fallback ook eventuele per-user Add-AppxPackage
    installaties op (voor alle profielen die al zijn ingelogd).

    Gebruik als Intune "Uninstall command":
        %SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -File Uninstall-ClaudeDesktop-Intune.ps1

    Raakt de Windows-kant van Cowork (VirtualMachinePlatform, Fast Startup) niet aan — dat is een
    eigen, onafhankelijke Win32-app met zijn eigen uninstall
    (../CoworkPrerequisites/Uninstall-CoworkPrerequisites-Intune.ps1), losstaand van of Claude
    Desktop hier wordt verwijderd of niet.

.NOTES
    Logt naar %ProgramData%\ClaudeDeploy\uninstall.log
#>

[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

$logDir = Join-Path $env:ProgramData "ClaudeDeploy"
New-Item -ItemType Directory -Path $logDir -Force | Out-Null
$logFile = Join-Path $logDir "uninstall.log"

function Write-Log {
    param([string]$Message)
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $Message"
    Add-Content -Path $logFile -Value $line
    Write-Host $line
}

try {
    Write-Log "=== Start Claude Desktop uninstall ==="

    # 1. Machine-brede provisioning verwijderen
    $provisioned = Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -like "*Claude*" }

    if ($provisioned) {
        foreach ($pkg in $provisioned) {
            Write-Log "Verwijderen provisioned package: $($pkg.DisplayName) $($pkg.Version)"
            Remove-AppxProvisionedPackage -Online -PackageName $pkg.PackageName -ErrorAction Stop | Out-Null
        }
    } else {
        Write-Log "Geen provisioned Claude-package gevonden."
    }

    # 2. Fallback: per-user installaties verwijderen voor reeds geladen profielen
    #    (Get-AppxPackage -AllUsers vereist ook admin en pakt alle geregistreerde profielen mee)
    $userPackages = Get-AppxPackage -AllUsers -Name "*Claude*" -ErrorAction SilentlyContinue
    if ($userPackages) {
        foreach ($pkg in $userPackages) {
            Write-Log "Verwijderen per-user package: $($pkg.PackageFullName) voor $($pkg.PackageUserInformation.UserSecurityId.Sid -join ', ')"
            try {
                Remove-AppxPackage -Package $pkg.PackageFullName -AllUsers -ErrorAction Stop
            }
            catch {
                Write-Log "Waarschuwing: kon $($pkg.PackageFullName) niet verwijderen voor alle users: $($_.Exception.Message)"
            }
        }
    } else {
        Write-Log "Geen per-user Claude-installaties gevonden."
    }

    # 3. Policy-sleutel opruimen die het install-script zette
    $policyPath = "HKLM:\SOFTWARE\Policies\Claude"
    if (Test-Path $policyPath) {
        Remove-ItemProperty -Path $policyPath -Name "disableAutoUpdates" -ErrorAction SilentlyContinue
        Write-Log "disableAutoUpdates policy-waarde verwijderd."
    }

    Write-Log "=== Uninstall script succesvol afgerond ==="
    exit 0
}
catch {
    Write-Log "FOUT: $($_.Exception.Message)"
    exit 1
}
