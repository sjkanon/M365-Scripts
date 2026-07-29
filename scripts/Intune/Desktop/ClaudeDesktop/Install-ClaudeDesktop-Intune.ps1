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
         vereist voor Cowork). Stond de feature al aan, dan gebeurt er verder niets bijzonders.
         Moest de feature net worden ingeschakeld, dan stuurt dit script aan het einde een
         Engelstalige melding (msg.exe) naar de actief ingelogde gebruiker (de sessie met status
         "Active" volgens quser, niet een broadcast naar alle sessies) dat die zelf moet
         herstarten — dit script herstart het apparaat NIET zelf. Een net ingeschakelde optional
         feature is pas na een herstart écht actief, en Cowork zou tot dan niet werken.
      2. Verwijdert eerder geprovisioneerde Claude-versies (Remove-AppxProvisionedPackage).
         Add-AppxProvisionedPackage vervangt een bestaande, andere versie niet automatisch —
         zonder deze stap stapelen oude versies zich op in de image. Dit raakt alleen de
         provisioning-laag (toekomstige profielen); al ingelogde gebruikers met de oude versie
         geregistreerd blijven gewoon werken totdat zij zelf opnieuw inloggen of de app herstarten.
      3. Add-AppxProvisionedPackage voor machine-brede installatie van de nieuwe versie (bereikt
         ook standaardgebruikers zonder adminrechten; Add-AppxPackage alleen zou dat niet doen)
      4. HKLM:\SOFTWARE\Policies\Claude\disableAutoUpdates = 1, zodat de ingebouwde
         auto-updater de machine-brede provisioning niet per-user kan overschrijven —
         versiebeheer loopt voortaan via de maandelijkse Deploy-ClaudeDesktopIntune.ps1 run

    Gebruik als Intune "Install command":
        %SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -File Install-ClaudeDesktop-Intune.ps1

    Gebruik het meegeleverde Detect-ClaudeDesktop-Intune.ps1 als "Custom detection script".

.PARAMETER MsixFileName
    Bestandsnaam van de MSIX naast dit script (standaard: Claude.msix).

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

function Get-ActiveConsoleSessionId {
    # quser geeft alle sessies terug; alleen de sessie met status "Active" is de daadwerkelijk
    # ingelogde (interactieve) gebruiker — een msg.exe naar "*" zou ook losstaande/disconnected
    # sessies raken, wat hier niet de bedoeling is.
    try {
        $output = quser 2>$null
        if (-not $output) { return $null }
        $activeLine = $output | Select-Object -Skip 1 | Where-Object { $_ -match '\bActive\b' } | Select-Object -First 1
        if ($activeLine -match '\s(\d+)\s+Active\b') {
            return $Matches[1]
        }
        return $null
    } catch {
        return $null
    }
}

try {
    Write-Log "=== Start Claude Desktop install ==="

    # 1. Virtual Machine Platform (vereist voor Cowork)
    $vmp = Get-WindowsOptionalFeature -Online -FeatureName VirtualMachinePlatform
    $vmpJustEnabled = $false
    if ($vmp.State -ne "Enabled") {
        Write-Log "VirtualMachinePlatform niet ingeschakeld, wordt nu ingeschakeld..."
        Enable-WindowsOptionalFeature -Online -FeatureName VirtualMachinePlatform -All -NoRestart | Out-Null
        $vmpJustEnabled = $true
        Write-Log "VirtualMachinePlatform ingeschakeld (gebruiker wordt aan het einde van dit script gevraagd zelf te herstarten)."
    } else {
        Write-Log "VirtualMachinePlatform was al ingeschakeld, geen herstart nodig."
    }

    # 2. Eerder geprovisioneerde Claude-versies opruimen vóór het provisioneren van de nieuwe
    #    (zie .DESCRIPTION hierboven waarom dit nodig is)
    $oldProvisioned = Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -like "*Claude*" }
    if ($oldProvisioned) {
        foreach ($old in $oldProvisioned) {
            Write-Log "Oudere geprovisioneerde versie verwijderen: $($old.DisplayName) $($old.Version)"
            try {
                Remove-AppxProvisionedPackage -Online -PackageName $old.PackageName -ErrorAction Stop | Out-Null
            } catch {
                Write-Log "Waarschuwing: kon oude provisioned package niet verwijderen: $($_.Exception.Message)"
            }
        }
    } else {
        Write-Log "Geen eerder geprovisioneerde Claude-versie gevonden (eerste installatie)."
    }

    # 3. MSIX pad bepalen (naast dit script, zoals Intune Win32-content dat plaatst)
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    $msixPath = Join-Path $scriptDir $MsixFileName

    if (-not (Test-Path $msixPath)) {
        throw "MSIX niet gevonden op verwacht pad: $msixPath"
    }

    # 4. Machine-brede provisioning
    Write-Log "Provisioneren van $msixPath machine-breed..."
    Add-AppxProvisionedPackage -Online -PackagePath $msixPath -SkipLicense -Regions "all" | Out-Null
    Write-Log "Add-AppxProvisionedPackage voltooid."

    # 5. Auto-updater uitschakelen: versiebeheer verloopt via de maandelijkse Intune-content-update,
    #    niet via Claude's eigen updater (die zou de machine-brede/Cowork-registratie per-user
    #    kunnen overschrijven — zie support.claude.com "Enterprise configuration for Claude Desktop")
    $policyPath = "HKLM:\SOFTWARE\Policies\Claude"
    if (-not (Test-Path $policyPath)) {
        New-Item -Path $policyPath -Force | Out-Null
    }
    New-ItemProperty -Path $policyPath -Name "disableAutoUpdates" -PropertyType DWord -Value 1 -Force | Out-Null
    Write-Log "disableAutoUpdates=1 gezet onder $policyPath."

    Write-Log "=== Install script succesvol afgerond ==="

    # 6. Alleen een melding sturen als VirtualMachinePlatform in déze run net is ingeschakeld —
    #    dit script herstart het apparaat zelf NIET, de gebruiker moet dat zelf doen.
    if ($vmpJustEnabled) {
        $restartMessage = "A required Windows feature (Virtual Machine Platform) was just enabled to support Claude Cowork. Please restart this computer as soon as possible to finish enabling it."
        $sessionId = Get-ActiveConsoleSessionId
        if ($sessionId) {
            Write-Log "Melding sturen naar ingelogde gebruiker (sessie $sessionId) om zelf te herstarten (geen automatische herstart)."
            try {
                & msg.exe $sessionId /TIME:0 $restartMessage
            } catch {
                Write-Log "Waarschuwing: kon geen melding naar ingelogde gebruiker sturen: $($_.Exception.Message)"
            }
        } else {
            Write-Log "Geen actief ingelogde gebruiker gevonden — melding overgeslagen (VirtualMachinePlatform vereist alsnog een herstart bij volgend gebruik)."
        }
    }

    exit 0
}
catch {
    Write-Log "FOUT: $($_.Exception.Message)"
    exit 1
}
