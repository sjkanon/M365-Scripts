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
      2. Verwijdert Claude Desktop VOLLEDIG van dit apparaat vóór de nieuwe installatie: sluit
         eventueel actieve Claude-processen, verwijdert alle per-user Appx-installaties
         (Get-AppxPackage -AllUsers / Remove-AppxPackage -AllUsers), eerder geprovisioneerde
         machine-brede versies (Remove-AppxProvisionedPackage), én eventuele "klassieke"
         (niet-Appx) per-user installaties zoals de consumer-installer van claude.ai/download
         die neerzet — die registreert zichzelf via een gewone per-user Uninstall-registry-key,
         niet als Appx-package, dus Get-AppxPackage ziet die nooit. Voor elk lokaal profiel
         (ingelogd of niet — niet-geladen profielhives worden tijdelijk geladen) wordt de
         Uninstall-registry doorzocht op "*Claude*" en de bijbehorende
         QuietUninstallString/UninstallString uitgevoerd. Dit is bewust grondiger dan alleen de
         provisioning-laag opschonen: elke blijvende, niet via deze route beheerde installatie
         kan een eigen, niet-Cowork-geregistreerde Claude-sessie in stand houden, ook nadat de
         machine-brede versie is bijgewerkt. Een ingelogde gebruiker die Claude open heeft staan
         verliest hierdoor die sessie.
      3. Add-AppxProvisionedPackage voor machine-brede installatie van de nieuwe versie (bereikt
         ook standaardgebruikers zonder adminrechten; Add-AppxPackage alleen zou dat niet doen) —
         zo staat Cowork voor elke gebruiker die hierna inlogt vanaf de eerste keer klaar.
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

function Get-ClassicClaudeUninstallEntries {
    # Vindt "klassieke" (niet-Appx) Claude Desktop-installaties: de consumer-installer van
    # claude.ai/download registreert zichzelf per-user via een gewone Uninstall-registry-key
    # (zoals de meeste Electron-apps), niet als Appx-package — Get-AppxPackage ziet die dus
    # nooit. Doorzoekt HKLM (zeldzaam voor deze installer, maar goedkoop om mee te nemen) en de
    # per-user hive van elk lokaal profiel, inclusief profielen die nu niet zijn ingelogd (hive
    # wordt daarvoor tijdelijk geladen vanuit NTUSER.DAT en aan het einde weer ontladen).
    $uninstallSubPaths = @(
        'Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    $entries = [System.Collections.Generic.List[object]]::new()

    foreach ($subPath in $uninstallSubPaths) {
        Get-ItemProperty -Path "HKLM:\$subPath" -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -like "*Claude*" } |
            ForEach-Object { $entries.Add($_) }
    }

    # Al geladen user-hives (ingelogde/onlangs ingelogde profielen) — SIDs van echte
    # gebruikersaccounts, geen ..._Classes-subhives.
    Get-ChildItem 'Registry::HKEY_USERS' -ErrorAction SilentlyContinue |
        Where-Object { $_.PSChildName -match '^S-1-5-21-\d+-\d+-\d+-\d+$' } |
        ForEach-Object {
            $sid = $_.PSChildName
            foreach ($subPath in $uninstallSubPaths) {
                Get-ItemProperty -Path "Registry::HKEY_USERS\$sid\$subPath" -ErrorAction SilentlyContinue |
                    Where-Object { $_.DisplayName -like "*Claude*" } |
                    ForEach-Object { $entries.Add($_) }
            }
        }

    # Niet-geladen profielen: NTUSER.DAT tijdelijk laden, doorzoeken, weer ontladen.
    $systemProfileNames = @('Public', 'Default', 'Default User', 'All Users')
    Get-ChildItem 'C:\Users' -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notin $systemProfileNames } |
        ForEach-Object {
            $ntUserPath = Join-Path $_.FullName 'NTUSER.DAT'
            $tempHiveName = "ClaudeCleanup_$($_.Name)"
            if ((Test-Path $ntUserPath) -and -not (Test-Path "Registry::HKEY_USERS\$tempHiveName")) {
                $loaded = $false
                try {
                    & reg.exe load "HKU\$tempHiveName" $ntUserPath *> $null
                    $loaded = ($LASTEXITCODE -eq 0)
                } catch {
                    $loaded = $false
                }

                if ($loaded) {
                    try {
                        foreach ($subPath in $uninstallSubPaths) {
                            Get-ItemProperty -Path "Registry::HKEY_USERS\$tempHiveName\$subPath" -ErrorAction SilentlyContinue |
                                Where-Object { $_.DisplayName -like "*Claude*" } |
                                ForEach-Object { $entries.Add($_) }
                        }
                    } finally {
                        # .NET houdt registry-handles soms vast tot een expliciete GC-cyclus —
                        # zonder dit faalt "reg unload" regelmatig met "Access is denied".
                        [System.GC]::Collect()
                        [System.GC]::WaitForPendingFinalizers()
                        & reg.exe unload "HKU\$tempHiveName" *> $null
                    }
                }
            }
        }

    return $entries
}

function Invoke-ClassicUninstall {
    param([Parameter(Mandatory = $true)][string]$UninstallCommand, [int]$TimeoutSeconds = 120)
    $proc = Start-Process -FilePath 'cmd.exe' -ArgumentList "/c `"$UninstallCommand`"" -WindowStyle Hidden -PassThru
    if (-not $proc.WaitForExit($TimeoutSeconds * 1000)) {
        try { $proc.Kill() } catch {}
        throw "Uninstall reageerde niet binnen $TimeoutSeconds seconden en is afgebroken."
    }
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

    # 2. Claude Desktop volledig verwijderen vóór de nieuwe installatie (zie .DESCRIPTION)
    Write-Log "Actieve Claude-processen sluiten (indien aanwezig)..."
    Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -like "*Claude*" } |
        Stop-Process -Force -ErrorAction SilentlyContinue

    $oldUserPackages = Get-AppxPackage -AllUsers -Name "*Claude*" -ErrorAction SilentlyContinue
    if ($oldUserPackages) {
        foreach ($pkg in $oldUserPackages) {
            Write-Log "Per-user installatie verwijderen: $($pkg.PackageFullName) voor $($pkg.PackageUserInformation.UserSecurityId.Sid -join ', ')"
            try {
                Remove-AppxPackage -Package $pkg.PackageFullName -AllUsers -ErrorAction Stop
            } catch {
                Write-Log "Waarschuwing: kon $($pkg.PackageFullName) niet verwijderen voor alle users: $($_.Exception.Message)"
            }
        }
    } else {
        Write-Log "Geen per-user Claude-installaties gevonden."
    }

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

    # Klassieke (niet-Appx) per-user installaties opruimen — bv. de consumer-installer van
    # claude.ai/download, die geen Appx-package is (zie .DESCRIPTION).
    $classicEntries = Get-ClassicClaudeUninstallEntries
    if ($classicEntries.Count -gt 0) {
        foreach ($entry in $classicEntries) {
            $uninstallCmd = if ($entry.QuietUninstallString) { $entry.QuietUninstallString } else { $entry.UninstallString }
            if (-not $uninstallCmd) { continue }
            Write-Log "Klassieke installatie verwijderen: $($entry.DisplayName) $($entry.DisplayVersion) via: $uninstallCmd"
            try {
                Invoke-ClassicUninstall -UninstallCommand $uninstallCmd
            } catch {
                Write-Log "Waarschuwing: klassieke uninstall mislukt voor $($entry.DisplayName): $($_.Exception.Message)"
            }
        }
    } else {
        Write-Log "Geen klassieke (niet-Appx) Claude-installaties gevonden."
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
