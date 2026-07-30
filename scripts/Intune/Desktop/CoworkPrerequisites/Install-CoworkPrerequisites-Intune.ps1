<#
.SYNOPSIS
    Intune Win32-app install script voor de Windows-vereisten van Claude Cowork
    (VirtualMachinePlatform + Fast Startup), los van de Claude Desktop-app zelf.

.DESCRIPTION
    Losgetrokken van Install-ClaudeDesktop-Intune.ps1: dit stuk heeft niets met Claude Desktop
    zelf te maken (dat werkt prima zonder dit), maar is puur de Windows-kant die Cowork nodig
    heeft. Door dit als eigen Win32-app te draaien, zonder Intune-dependency naar de Claude
    Desktop-app, kan Claude Desktop altijd geïnstalleerd worden ongeacht of Cowork-prereqs
    slagen — en zie je in Intune apart of het de Windows-kant of de Claude-kant is die faalt,
    in plaats van één opaak install command dat alles combineert.

    Voert uit:
      1. Enable-WindowsOptionalFeature -FeatureName VirtualMachinePlatform (indien nodig), met
         retry tegen voorbijgaande DISM-storingen. Stond de feature al aan, dan gebeurt er
         verder niets bijzonders.
      2. Fast Startup (HiberbootEnabled) uitschakelen — expliciet genoemd in Anthropic's eigen
         Cowork-documentatie: "Restart the machine using Restart, not shut down and power on.
         With Windows Fast Startup enabled, a shutdown cycle can leave the virtualization
         services uninitialized." Bij elke run gezet, niet alleen als VMP in déze run net is
         ingeschakeld.
      3. Als VMP in déze run net is ingeschakeld: een Engelstalige melding (msg.exe) naar de
         actief ingelogde gebruiker (herkend via de niet-vertaalde SESSIONNAME "console", niet
         de per OS-taal wisselende STATE-tekst "Active"), én exitcode 3010 ("soft reboot
         required") zodat Intune's eigen herstart-UX (RestartBehavior 'basedOnReturnCode' in
         Deploy-CoworkPrerequisitesIntune.ps1) de herstart afdwingt/plant — dit script herstart
         het apparaat NIET zelf. De melding wordt niet rechtstreeks vanuit deze SYSTEM-context
         verstuurd (dat levert een dialoog op waarvan de OK-knop niet reageert op klikken — een
         bekend msg.exe-euvel bij cross-session-berichten), maar via een kortstondige geplande
         taak die msg.exe binnen de eigen sessie van de gebruiker draait.

    Gebruik als Intune "Install command":
        %SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -File Install-CoworkPrerequisites-Intune.ps1

    Gebruik het meegeleverde Detect-CoworkPrerequisites-Intune.ps1 als "Custom detection script".

.NOTES
    Logt naar %ProgramData%\CoworkPrereqDeploy\install.log voor troubleshooting via Intune
    diagnostics / IME-logs.
#>

[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

$logDir = Join-Path $env:ProgramData "CoworkPrereqDeploy"
New-Item -ItemType Directory -Path $logDir -Force | Out-Null
$logFile = Join-Path $logDir "install.log"

function Write-Log {
    param([string]$Message)
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $Message"
    Add-Content -Path $logFile -Value $line
    Write-Host $line
}

function Invoke-WithRetry {
    # Vangt voorbijgaande DISM-storingen op (bv. "busy" doordat Autopilot ESP meerdere
    # Win32-apps/features tegelijk aan het verwerken is).
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

function Get-ActiveConsoleSession {
    # SESSIONNAME "console" is een vaste, niet-vertaalde WinStation-naam, in tegenstelling tot de
    # STATE-kolom ("Active"), die per OS-weergavetaal verschilt (bv. "Actief" op nl-NL).
    try {
        $output = quser 2>$null
        if (-not $output) { return $null }
        $consoleLine = $output | Select-Object -Skip 1 |
            Where-Object { $_ -match '^\s*>?(\S+)\s+console\s+(\d+)\s' } | Select-Object -First 1
        if ($consoleLine -match '^\s*>?(\S+)\s+console\s+(\d+)\s') {
            return [PSCustomObject]@{
                UserName  = $Matches[1]
                SessionId = $Matches[2]
            }
        }
        return $null
    } catch {
        return $null
    }
}

function Show-UserRestartNotification {
    # Rechtstreeks "msg.exe <sessieId> ..." aanroepen vanuit déze SYSTEM-context levert een
    # dialoog op waarvan de OK-knop niet op klikken reageert — een bekend msg.exe-euvel bij
    # cross-session-berichten die van een niet-interactieve afzender komen. Door msg.exe via een
    # kortstondige geplande taak in de eigen sessie van de ingelogde gebruiker te draaien, hoort
    # de dialoog bij een echt interactief bureaublad en werkt de OK-knop wél gewoon.
    param([Parameter(Mandatory = $true)][string]$Message)

    $session = Get-ActiveConsoleSession
    if (-not $session) {
        Write-Log "Geen actief ingelogde gebruiker gevonden — melding overgeslagen, Intune plant de herstart af via exitcode 3010 (bv. tijdens Autopilot ESP, waar nog niemand is ingelogd)."
        return
    }

    $taskName = "CoworkPrereqNotify_$([guid]::NewGuid().ToString('N').Substring(0, 8))"
    try {
        $action = New-ScheduledTaskAction -Execute "$env:SystemRoot\System32\msg.exe" -Argument "$($session.SessionId) /TIME:0 `"$Message`""
        $principal = New-ScheduledTaskPrincipal -UserId $session.UserName -LogonType Interactive -RunLevel Limited
        $task = New-ScheduledTask -Action $action -Principal $principal
        Register-ScheduledTask -TaskName $taskName -InputObject $task -Force | Out-Null
        Start-ScheduledTask -TaskName $taskName
        Start-Sleep -Seconds 2
        Write-Log "Melding gestuurd naar ingelogde gebruiker ($($session.UserName), sessie $($session.SessionId)) via geplande taak."
    } catch {
        Write-Log "Waarschuwing: kon geen melding naar ingelogde gebruiker sturen: $($_.Exception.Message)"
    } finally {
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
    }
}

try {
    Write-Log "=== Start Cowork Prerequisites install ==="

    # 1. Virtual Machine Platform
    $vmpJustEnabled = $false
    $vmp = Invoke-WithRetry -Action { Get-WindowsOptionalFeature -Online -FeatureName VirtualMachinePlatform -ErrorAction Stop }
    if ($vmp.State -ne "Enabled") {
        Write-Log "VirtualMachinePlatform niet ingeschakeld, wordt nu ingeschakeld..."
        Invoke-WithRetry -Action { Enable-WindowsOptionalFeature -Online -FeatureName VirtualMachinePlatform -All -NoRestart -ErrorAction Stop | Out-Null }
        $vmpJustEnabled = $true
        Write-Log "VirtualMachinePlatform ingeschakeld (herstart vereist — zie exitcode 3010 aan het einde van dit script)."
    } else {
        Write-Log "VirtualMachinePlatform was al ingeschakeld, geen herstart nodig."
    }

    # 2. Fast Startup (Hiberboot) uitschakelen — zie .DESCRIPTION.
    $powerPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power"
    New-ItemProperty -Path $powerPath -Name "HiberbootEnabled" -PropertyType DWord -Value 0 -Force | Out-Null
    Write-Log "Fast Startup (HiberbootEnabled) uitgeschakeld, zodat een shutdown/power-on-cyclus de Cowork-virtualisatieservices niet ongeïnitialiseerd achterlaat."

    Write-Log "=== Install script succesvol afgerond ==="

    # 3. Alleen relevant als VirtualMachinePlatform in déze run net is ingeschakeld.
    if ($vmpJustEnabled) {
        $restartMessage = "A required Windows feature (Virtual Machine Platform) was just enabled to support Claude Cowork. Please restart this computer as soon as possible to finish enabling it."
        Show-UserRestartNotification -Message $restartMessage
        Write-Log "Exit 3010 (herstart vereist); Intune plant daarnaast zelf de herstart af via exitcode 3010."
        exit 3010
    }

    exit 0
}
catch {
    Write-Log "FOUT: $($_.Exception.Message)"
    exit 1
}
