<#
.SYNOPSIS
    Cowork Windows-vereisten (VirtualMachinePlatform + Fast Startup) als Intune "Platform script"
    (Devices > Scripts and remediations > Platform scripts) - los alternatief naast de Win32-app
    in deze map (Deploy-/Install-/Uninstall-/Detect-CoworkPrerequisites-Intune.ps1).

.DESCRIPTION
    Zelfde inhoud/logica als Install-CoworkPrerequisites-Intune.ps1, aangepast voor het Platform
    Scripts-model in plaats van een Win32-app "install command":

      - Platform scripts kennen GEEN return-code-naar-restart-mapping zoals Win32-apps
        (-RestartBehavior 'basedOnReturnCode' + exit 3010). Intune interpreteert bij een platform
        script alleen exit 0 = geslaagd, alles anders = mislukt - exit 3010 zou hier dus ten
        onrechte als FOUT gerapporteerd worden, terwijl VirtualMachinePlatform wel degelijk correct
        is ingeschakeld. Dit script sluit daarom altijd af met exit 0 na een geslaagde configuratie,
        ongeacht of een herstart nog moet gebeuren.
      - Geen "Sysnative"-pad nodig in een handmatige command line - Intune's eigen "Run script in
        64-bit PowerShell Host"-instelling in de portal (zet op Ja) regelt dat al.
      - Geen detection rule, requirement rule of GRS-gerelateerde complexiteit - een platform
        script is gewoon: script draait, exit code bepaalt geslaagd/mislukt. Wel een gevolg: geen
        automatische periodieke herevaluatie zoals een Proactive Remediation (die was hier sowieso
        niet bruikbaar - zie readme.md, licentie-eis Windows Enterprise/Education/VDA, niet gedekt
        door Business Premium). Bij een net ingeschakelde VirtualMachinePlatform komt er dus maar
        Eenmalig een melding - geen herhaalde herinnering totdat de gebruiker herstart, in
        tegenstelling tot de Win32-app-route (die via return-code 3010 een echte, door Intune
        afgedwongen herstart-policy triggert) of een remediation (die dagelijks opnieuw zou nagen).

    Voert uit:
      1. Enable-WindowsOptionalFeature -FeatureName VirtualMachinePlatform (indien nodig), met
         retry tegen voorbijgaande DISM-storingen. Stond de feature al aan, dan gebeurt er verder
         niets bijzonders.
      2. Fast Startup (HiberbootEnabled) uitschakelen - expliciet genoemd in Anthropic's eigen
         Cowork-documentatie: "Restart the machine using Restart, not shut down and power on. With
         Windows Fast Startup enabled, a shutdown cycle can leave the virtualization services
         uninitialized." Bij elke run gezet, niet alleen als VMP in deze run net is ingeschakeld.
      3. Als VMP in deze run net is ingeschakeld: een Engelstalige melding (msg.exe) naar de actief
         ingelogde gebruiker (herkend via de niet-vertaalde SESSIONNAME "console") dat die zelf
         moet herstarten. Dit script herstart het apparaat NIET zelf en kan dat ook niet via Intune
         laten afdwingen (dat vereist het Win32-app-return-code-mechanisme).

    Uploaden in Intune:
      1. Devices > Scripts and remediations > Platform scripts > Add > Windows 10 and later
      2. Upload dit .ps1-bestand
      3. Script settings: "Run this script using the logged on credentials" = No (SYSTEM),
         "Enforce script signature check" = No, "Run script in 64-bit PowerShell Host" = Yes
      4. Wijs toe aan dezelfde groep als Claude Desktop

.NOTES
    Logt naar %ProgramData%\CoworkPrereqDeploy\platformscript.log (los van het install.log van de
    Win32-app-variant, zodat de twee elkaars logs niet overschrijven als je ooit beide test op
    hetzelfde apparaat).
#>

[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

$logDir = Join-Path $env:ProgramData "CoworkPrereqDeploy"
New-Item -ItemType Directory -Path $logDir -Force | Out-Null
$logFile = Join-Path $logDir "platformscript.log"

function Write-Log {
    param([string]$Message)
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $Message"
    Add-Content -Path $logFile -Value $line
    Write-Host $line
}

function Invoke-WithRetry {
    # Vangt voorbijgaande DISM-storingen op (bv. "busy" doordat er tegelijk andere scripts/
    # Win32-apps/features verwerkt worden).
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
    Write-Log "=== Start Cowork Prerequisites (platform script) ==="

    # 1. Virtual Machine Platform
    $vmpJustEnabled = $false
    $vmp = Invoke-WithRetry -Action { Get-WindowsOptionalFeature -Online -FeatureName VirtualMachinePlatform -ErrorAction Stop }
    if ($vmp.State -ne "Enabled") {
        Write-Log "VirtualMachinePlatform niet ingeschakeld, wordt nu ingeschakeld..."
        Invoke-WithRetry -Action { Enable-WindowsOptionalFeature -Online -FeatureName VirtualMachinePlatform -All -NoRestart -ErrorAction Stop | Out-Null }
        $vmpJustEnabled = $true
        Write-Log "VirtualMachinePlatform ingeschakeld - herstart vereist voordat de HCS-services actief worden."
    } else {
        Write-Log "VirtualMachinePlatform was al ingeschakeld, geen herstart nodig."
    }

    # 2. Fast Startup (Hiberboot) uitschakelen - zie .DESCRIPTION.
    $powerPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power"
    New-ItemProperty -Path $powerPath -Name "HiberbootEnabled" -PropertyType DWord -Value 0 -Force | Out-Null
    Write-Log "Fast Startup (HiberbootEnabled) uitgeschakeld, zodat een shutdown/power-on-cyclus de Cowork-virtualisatieservices niet ongeinitialiseerd achterlaat."

    # 3. Melding bij een net ingeschakelde VMP - eenmalig, geen door Intune afgedwongen herstart
    #    mogelijk vanuit een platform script (zie .DESCRIPTION).
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
            Write-Log "Geen actief ingelogde gebruiker gevonden - melding overgeslagen (bv. tijdens Autopilot ESP)."
        }
    }

    # Platform scripts kennen geen 3010/soft-reboot-semantiek - altijd exit 0 bij geslaagde
    # configuratie, ongeacht of een herstart nog moet gebeuren (zie .DESCRIPTION).
    Write-Log "=== Platform script succesvol afgerond ==="
    exit 0
}
catch {
    Write-Log "FOUT: $($_.Exception.Message)"
    exit 1
}
