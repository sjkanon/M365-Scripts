<#
.SYNOPSIS
    Intune Proactive Remediation — remediatiehelft — Cowork Windows-vereisten.

.DESCRIPTION
    Schakelt VirtualMachinePlatform in en Fast Startup (Hiberboot) uit — de twee Windows-kant-
    vereisten voor Claude Cowork. Wordt door Intune alleen aangeroepen als
    Detect-CoworkPrerequisites.ps1 "niet compliant" rapporteerde.

    Voert uit:
      1. Enable-WindowsOptionalFeature -FeatureName VirtualMachinePlatform (indien nodig), met
         retry tegen voorbijgaande DISM-storingen. Stond de feature al aan, dan gebeurt er verder
         niets bijzonders.
      2. Fast Startup (HiberbootEnabled) uitschakelen — expliciet genoemd in Anthropic's eigen
         Cowork-documentatie: "Restart the machine using Restart, not shut down and power on.
         With Windows Fast Startup enabled, a shutdown cycle can leave the virtualization
         services uninitialized." Bij elke run gezet, niet alleen als VMP in déze run net is
         ingeschakeld.
      3. Als VMP in déze run net is ingeschakeld: een Engelstalige melding (msg.exe) naar de
         actief ingelogde gebruiker (herkend via de niet-vertaalde SESSIONNAME "console", niet de
         per OS-taal wisselende STATE-tekst "Active"). De gebruiker moet zelf herstarten — dit
         script herstart het apparaat NIET zelf (Proactive Remediations hebben, anders dan
         Win32-apps, geen return-code-gebaseerd herstart-mechanisme om op terug te vallen, dus
         forceren zou hier alleen via een expliciete Restart-Computer-aanroep kunnen, en dat is
         bewust niet de bedoeling — zie readme.md).

    Omdat Proactive Remediations standaard dagelijks opnieuw draaien: zolang de gebruiker niet
    herstart, blijft Detect-CoworkPrerequisites.ps1 "niet compliant" melden, en stuurt deze
    remediation de volgende cyclus gewoon opnieuw een herinnering — geen eenmalige melding die
    kan worden gemist, maar een terugkerende totdat het apparaat daadwerkelijk herstart is.

.NOTES
    Logt naar %ProgramData%\CoworkPrereqDeploy\remediate.log voor troubleshooting via Intune
    diagnostics / IME-logs.
#>

$ErrorActionPreference = "Stop"

$logDir = Join-Path $env:ProgramData "CoworkPrereqDeploy"
New-Item -ItemType Directory -Path $logDir -Force | Out-Null
$logFile = Join-Path $logDir "remediate.log"

function Write-Log {
    param([string]$Message)
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $Message"
    Add-Content -Path $logFile -Value $line
    Write-Host $line
}

function Invoke-WithRetry {
    # Vangt voorbijgaande DISM-storingen op (bv. "busy" doordat er tegelijk andere Win32-apps/
    # remediations verwerkt worden).
    param([scriptblock]$Action, [int]$MaxAttempts = 3, [int]$DelaySeconds = 10)
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            return & $Action
        } catch {
            if ($attempt -eq $MaxAttempts) { throw }
            Start-Sleep -Seconds $DelaySeconds
        }
    }
}

function Get-ActiveConsoleSessionId {
    # SESSIONNAME "console" is een vaste, niet-vertaalde WinStation-naam, in tegenstelling tot de
    # STATE-kolom ("Active"), die per OS-weergavetaal verschilt (bv. "Actief" op nl-NL).
    try {
        $output = quser 2>$null
        if (-not $output) { return $null }
        $consoleLine = $output | Select-Object -Skip 1 |
            Where-Object { $_ -match '^\s*\S+\s+console\s+(\d+)\s' } | Select-Object -First 1
        if ($consoleLine -match '^\s*\S+\s+console\s+(\d+)\s') {
            return $Matches[1]
        }
        return $null
    } catch {
        return $null
    }
}

try {
    Write-Log "=== Start Cowork Prerequisites remediation ==="

    # 1. Virtual Machine Platform
    $vmpJustEnabled = $false
    $vmp = Invoke-WithRetry -Action { Get-WindowsOptionalFeature -Online -FeatureName VirtualMachinePlatform -ErrorAction Stop }
    if ($vmp.State -ne "Enabled") {
        Write-Log "VirtualMachinePlatform niet ingeschakeld, wordt nu ingeschakeld..."
        Invoke-WithRetry -Action { Enable-WindowsOptionalFeature -Online -FeatureName VirtualMachinePlatform -All -NoRestart -ErrorAction Stop | Out-Null }
        $vmpJustEnabled = $true
        Write-Log "VirtualMachinePlatform ingeschakeld — herstart vereist voordat de HCS-services actief worden."
    } else {
        Write-Log "VirtualMachinePlatform was al ingeschakeld."
    }

    # 2. Fast Startup (Hiberboot) uitschakelen — zie .DESCRIPTION.
    $powerPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power"
    New-ItemProperty -Path $powerPath -Name "HiberbootEnabled" -PropertyType DWord -Value 0 -Force | Out-Null
    Write-Log "Fast Startup (HiberbootEnabled) uitgeschakeld, zodat een shutdown/power-on-cyclus de Cowork-virtualisatieservices niet ongeïnitialiseerd achterlaat."

    # 3. Melding bij een net ingeschakelde VMP — de gebruiker moet zelf herstarten, dit script
    #    doet dat NIET automatisch (zie .DESCRIPTION).
    if ($vmpJustEnabled) {
        $restartMessage = "A required Windows feature (Virtual Machine Platform) was just enabled to support Claude Cowork. Please restart this computer as soon as possible to finish enabling it."
        $sessionId = Get-ActiveConsoleSessionId
        if ($sessionId) {
            Write-Log "Melding sturen naar ingelogde gebruiker (sessie $sessionId) om zelf te herstarten."
            try {
                & msg.exe $sessionId /TIME:0 $restartMessage
            } catch {
                Write-Log "Waarschuwing: kon geen melding naar ingelogde gebruiker sturen: $($_.Exception.Message)"
            }
        } else {
            Write-Log "Geen actief ingelogde gebruiker gevonden — melding overgeslagen. Detectie blijft 'niet compliant' tot een herstart, dus de melding volgt bij de eerstvolgende dagelijkse remediation-cyclus alsnog."
        }
    }

    Write-Log "=== Remediation succesvol afgerond ==="
    exit 0
}
catch {
    Write-Log "FOUT: $($_.Exception.Message)"
    exit 1
}
