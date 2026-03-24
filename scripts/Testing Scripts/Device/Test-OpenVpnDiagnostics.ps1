# Maak C:\Temp aan als het niet bestaat
if (-not (Test-Path "C:\Temp")) { New-Item -ItemType Directory -Path "C:\Temp" | Out-Null }

$logPath = "C:\Temp\OpenVPN_Diagnostics_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
$output = @()
$issues = @()
$output += "=== OPENVPN DIAGNOSTICS ==="
$output += "Datum: $(Get-Date)"
$output += ""

# --- Wintun/TAP Adapters inclusief verborgen/oude ---
$output += "--- Wintun/TAP Adapters (PnP) inclusief verborgen/oude ---"
$wintun = Get-PnpDevice -Class Net -PresentOnly:$false | Where-Object {
    $_.FriendlyName -like "*Wintun*" -or 
    $_.FriendlyName -like "*TAP*" -or
    $_.FriendlyName -like "*DCO*"
}
if ($wintun) {
    $wintun | ForEach-Object {
        $output += "Status: $($_.Status) | Naam: $($_.FriendlyName) | ID: $($_.InstanceId)"
        if ($_.Status -ne "OK") {
            $issues += "PROBLEEM: Adapter '$($_.FriendlyName)' heeft status '$($_.Status)'"
        }
    }
    $tapCount = ($wintun | Where-Object {$_.FriendlyName -like "*TAP*"}).Count
    if ($tapCount -gt 1) {
        $issues += "PROBLEEM: $tapCount TAP adapter instanties gevonden - conflicteren mogelijk"
    }
} else {
    $output += "GEEN Wintun/TAP/DCO adapters gevonden"
    $issues += "PROBLEEM: Geen VPN adapter aanwezig - driver mogelijk niet geinstalleerd"
}
$output += ""

# --- Alle Virtual Network Adapters ---
$output += "--- Virtual Network Adapters ---"
$adapters = Get-NetAdapter | Where-Object {
    $_.InterfaceDescription -like "*Wintun*" -or 
    $_.InterfaceDescription -like "*TAP*" -or 
    $_.InterfaceDescription -like "*Virtual*" -or 
    $_.InterfaceDescription -like "*VPN*" -or 
    $_.InterfaceDescription -like "*DCO*"
}
if ($adapters) {
    $adapters | ForEach-Object {
        $output += "Naam: $($_.Name) | Omschrijving: $($_.InterfaceDescription) | Status: $($_.Status)"
    }
} else {
    $output += "GEEN virtuele adapters zichtbaar"
    $issues += "WAARSCHUWING: Geen virtuele netwerkadapter zichtbaar"
}
$output += ""

# --- Network Profiles ---
$output += "--- Network Profiles ---"
Get-NetConnectionProfile | ForEach-Object {
    $output += "Netwerk: $($_.Name) | Adapter: $($_.InterfaceAlias) | Categorie: $($_.NetworkCategory)"
    if ($_.NetworkCategory -eq "Public" -and (
        $_.InterfaceAlias -like "*Wintun*" -or 
        $_.InterfaceAlias -like "*TAP*" -or 
        $_.InterfaceAlias -like "*OpenVPN*"
    )) {
        $issues += "PROBLEEM: VPN adapter '$($_.InterfaceAlias)' staat op Public - moet Private zijn"
    }
}
$output += ""

# --- Actieve Windows Gebruikerssessies ---
$output += "--- Actieve Windows Gebruikerssessies ---"
try {
    $sessions = query user 2>&1
    $output += $sessions
    $activeSessions = ($sessions | Where-Object {$_ -match "Active"}).Count
    if ($activeSessions -gt 1) {
        $issues += "PROBLEEM: $activeSessions actieve gebruikerssessies - meerdere sessies conflicteren met OpenVPN adapter"
    }
} catch {
    $output += "Kon gebruikerssessies niet ophalen: $_"
}
$output += ""

# --- VPN Software + versie check ---
$output += "--- VPN Software geinstalleerd ---"
$vpnApps = Get-WmiObject -Class Win32_Product | Where-Object {
    $_.Name -like "*VPN*" -or 
    $_.Name -like "*Cisco*" -or 
    $_.Name -like "*WireGuard*" -or 
    $_.Name -like "*OpenVPN*"
}
if ($vpnApps) {
    $vpnApps | ForEach-Object {
        $output += "App: $($_.Name) | Versie: $($_.Version)"
        if ($_.Name -notlike "*OpenVPN Connect*") {
            $issues += "WAARSCHUWING: Mogelijk conflicterende VPN software: '$($_.Name)'"
        }
        if ($_.Name -like "*OpenVPN Connect*") {
            $version = [version]$_.Version
            if ($version -lt [version]"3.5.0") {
                $issues += "WAARSCHUWING: OpenVPN Connect $($_.Version) is ouder dan 3.5.0 - bekende agent compatibiliteitsproblemen"
            }
        }
    }
} else {
    $output += "Geen VPN software gevonden"
    $issues += "PROBLEEM: OpenVPN Connect niet gevonden"
}
$output += ""

# --- MDM/Intune check ---
$output += "--- MDM / Intune Enrollment ---"
try {
    $mdmInfo = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Enrollments\*" -ErrorAction SilentlyContinue | 
        Select-Object PSChildName, UPN, EnrollmentType, ProviderID, EnrollmentState
    if ($mdmInfo) {
        $mdmInfo | ForEach-Object {
            $output += "MDM Provider: $($_.ProviderID) | UPN: $($_.UPN) | Type: $($_.EnrollmentType) | Status: $($_.EnrollmentState)"
        }
        $issues += "INFO: Toestel is MDM/Intune enrolled - controleer of Intune een oudere OpenVPN versie terug pusht"
    } else {
        $output += "Geen MDM enrollment gevonden"
    }
} catch {
    $output += "Kon MDM info niet ophalen: $_"
}
$output += ""

# --- Intune app deployment check ---
$output += "--- Intune Gedeployede Apps (OpenVPN gerelateerd) ---"
try {
    $intuneApps = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps\*\*" -ErrorAction SilentlyContinue |
        Where-Object {$_.PSChildName -match "OpenVPN" -or $_.DisplayName -match "OpenVPN"}
    if ($intuneApps) {
        $intuneApps | ForEach-Object {
            $output += "Intune App: $($_.PSChildName) | Status: $($_.ComplianceState)"
            $issues += "INFO: Intune beheert OpenVPN - mogelijke versie override bij herinstallatie"
        }
    } else {
        $output += "Geen OpenVPN Intune app deployment gevonden via registry"
    }
} catch {
    $output += "Kon Intune app info niet ophalen: $_"
}
$output += ""

# --- Hyper-V / WSL / Virtualisatie ---
$output += "--- Hyper-V / WSL / Virtualisatie ---"
$virtFeatures = Get-WindowsOptionalFeature -Online | Where-Object {
    $_.FeatureName -like "*Hyper-V*" -or 
    $_.FeatureName -like "*VirtualMachinePlatform*" -or 
    $_.FeatureName -like "*WSL*"
}
$virtFeatures | ForEach-Object {
    $output += "Feature: $($_.FeatureName) | Status: $($_.State)"
    if ($_.State -eq "Enabled" -and $_.FeatureName -like "*Hyper-V*") {
        $issues += "WAARSCHUWING: Hyper-V ingeschakeld - kan conflicteren met Wintun driver"
    }
}
$output += ""

# --- PnPUtil ---
$output += "--- Wintun/TAP in PnPUtil ---"
$pnpOutput = pnputil /enum-drivers 2>&1
$lines = $pnpOutput -split "`n"
$foundInPnp = $false
for ($i = 0; $i -lt $lines.Count; $i++) {
    if ($lines[$i] -match "wintun|tap|ovpn|dco") {
        $foundInPnp = $true
        $start = [Math]::Max(0, $i - 1)
        $end = [Math]::Min($lines.Count - 1, $i + 5)
        $output += $lines[$start..$end]
        $output += ""
    }
}
if (-not $foundInPnp) {
    $output += "Geen Wintun/TAP/DCO driver gevonden in PnPUtil"
    $issues += "PROBLEEM: VPN driver niet geregistreerd in PnPUtil"
}
$output += ""

# --- OpenVPN Services + Recovery ---
$output += "--- OpenVPN Services ---"
$services = Get-Service | Where-Object {
    $_.DisplayName -like "*OpenVPN*" -or 
    $_.DisplayName -like "*Wintun*"
}
if ($services) {
    $services | ForEach-Object {
        $output += "Service: $($_.DisplayName) | Status: $($_.Status) | Starttype: $($_.StartType)"
        if ($_.Status -ne "Running") {
            $issues += "PROBLEEM: Service '$($_.DisplayName)' is niet actief"
        }
        $scOutput = sc.exe qfailure $($_.Name) 2>&1
        $output += "  Recovery: $($scOutput -join ' ')"
    }
} else {
    $output += "Geen OpenVPN services gevonden"
    $issues += "PROBLEEM: Geen OpenVPN service aanwezig"
}
$output += ""

# --- DCO registry instelling ---
$output += "--- OpenVPN Connect DCO Instelling ---"
try {
    $dcoPaths = @("HKLM:\SOFTWARE\OpenVPN Connect", "HKCU:\SOFTWARE\OpenVPN Connect")
    $dcoFound = $false
    foreach ($path in $dcoPaths) {
        if (Test-Path $path) {
            $dcoSetting = Get-ItemProperty $path -ErrorAction SilentlyContinue
            $output += "Registry pad: $path"
            $output += "  Waarden: $($dcoSetting | Out-String)"
            $dcoFound = $true
        }
    }
    if (-not $dcoFound) {
        $output += "Geen OpenVPN Connect registry instellingen gevonden"
        $issues += "INFO: Probeer DCO in te schakelen via OpenVPN Connect advanced settings"
    }
} catch {
    $output += "Kon DCO registry niet lezen: $_"
}
$output += ""

# --- OpenVPN Connect Logs (alle gebruikersprofielen) ---
$output += "--- OpenVPN Connect Logs (alle gebruikersprofielen) ---"
$logFiles = Get-ChildItem "C:\Users\*\AppData\Local\OpenVPN Connect\logs\*.log" -ErrorAction SilentlyContinue | 
    Sort-Object LastWriteTime -Descending | 
    Select-Object -First 5

if ($logFiles) {
    foreach ($logFile in $logFiles) {
        $output += ""
        $output += ">> Logbestand: $($logFile.FullName)"
        $output += ">> Laatst gewijzigd: $($logFile.LastWriteTime)"

        # Kopieer ook het volledige logbestand naar C:\Temp
        $destName = "OpenVPN_Log_" + ($logFile.Name -replace "[^a-zA-Z0-9.]", "_")
        Copy-Item $logFile.FullName "C:\Temp\$destName" -Force -ErrorAction SilentlyContinue
        $output += ">> Kopie opgeslagen: C:\Temp\$destName"

        $logContent = Get-Content $logFile.FullName -Tail 100 -ErrorAction SilentlyContinue
        $relevantLines = $logContent | Where-Object {
            $_ -match "error|warn|fail|timeout|disconnect|connect|tun|tap|wintun|dco|protect|socket|fatal|denied|permission"
        }
        if ($relevantLines) {
            $output += "Relevante log regels:"
            $relevantLines | ForEach-Object {
                $output += "  $_"
                if ($_ -match "protect\(\) method") { $issues += "PROBLEEM: 'protect() method' fout in log" }
                if ($_ -match "general tun error") { $issues += "PROBLEEM: 'General tun error' in log" }
                if ($_ -match "access.denied|access denied") { $issues += "PROBLEEM: 'Access Denied' in log" }
                if ($_ -match "fatal") { $issues += "PROBLEEM: Fatale fout in log: $_" }
            }
        } else {
            $output += "Geen relevante foutmeldingen gevonden in laatste 100 regels"
        }
    }
} else {
    $output += "Geen OpenVPN Connect logbestanden gevonden in enig gebruikersprofiel"
    $issues += "WAARSCHUWING: Geen logbestanden gevonden - script draait als SYSTEM, logs staan in gebruikersprofiel"
    
    # Probeer alle OpenVPN log locaties te vinden
    $output += ""
    $output += "Zoeken naar alle OpenVPN bestanden op C:\Users..."
    $allOpenVPNFiles = Get-ChildItem "C:\Users\*\AppData\Local\OpenVPN*" -Recurse -ErrorAction SilentlyContinue | 
        Select-Object -First 20
    if ($allOpenVPNFiles) {
        $allOpenVPNFiles | ForEach-Object {
            $output += "  Gevonden: $($_.FullName) | $($_.LastWriteTime)"
        }
    } else {
        $output += "  Geen OpenVPN bestanden gevonden in gebruikersprofielen"
        $issues += "PROBLEEM: Geen OpenVPN Connect gebruikersdata gevonden - app mogelijk nooit succesvol gestart of profiel ontbreekt"
    }
}
$output += ""

# --- OpenVPN Connect config bestanden ---
$output += "--- OpenVPN Connect Configuratie ---"
$configFiles = Get-ChildItem "C:\Users\*\AppData\Local\OpenVPN Connect\profiles\*" -ErrorAction SilentlyContinue
if ($configFiles) {
    $configFiles | ForEach-Object {
        $output += "Config: $($_.FullName) | Gewijzigd: $($_.LastWriteTime)"
    }
} else {
    $output += "Geen configuratiebestanden gevonden"
    $issues += "WAARSCHUWING: Geen OpenVPN Connect profiel/config gevonden"
}
$output += ""

# --- Windows Event Log laatste 24 uur ---
$output += "--- Windows Event Log VPN/Driver fouten (laatste 24 uur) ---"
try {
    $events24u = Get-WinEvent -FilterHashtable @{
        LogName   = 'System'
        Level     = 1, 2, 3
        StartTime = (Get-Date).AddHours(-24)
    } -ErrorAction SilentlyContinue | Where-Object {
        $_.Message -match "VPN|TAP|Wintun|OpenVPN|network adapter|DCO"
    } | Select-Object -First 10
    if ($events24u) {
        $output += "Aantal: $($events24u.Count)"
        $events24u | ForEach-Object {
            $output += "Tijd: $($_.TimeCreated) | Level: $($_.LevelDisplayName) | Bron: $($_.ProviderName)"
            $output += "  Bericht: $($_.Message.Substring(0, [Math]::Min(300, $_.Message.Length)))"
            $output += ""
            if ($_.Level -le 2 -and $_.ProviderName -notlike "*DCOM*") {
                $issues += "PROBLEEM: Event Log fout van '$($_.ProviderName)' om $($_.TimeCreated)"
            }
        }
    } else {
        $output += "Geen relevante fouten gevonden"
    }
} catch {
    $output += "Kon Windows Event Log niet lezen: $_"
}
$output += ""

# --- Windows Event Log laatste maand ---
$output += "--- Windows Event Log VPN/Driver fouten (laatste maand) ---"
try {
    $events1m = Get-WinEvent -FilterHashtable @{
        LogName   = 'System'
        Level     = 1, 2, 3
        StartTime = (Get-Date).AddDays(-30)
        EndTime   = (Get-Date).AddHours(-24)
    } -ErrorAction SilentlyContinue | Where-Object {
        $_.Message -match "VPN|TAP|Wintun|OpenVPN|network adapter|DCO"
    } | Select-Object -First 20
    if ($events1m) {
        $output += "Aantal: $($events1m.Count)"
        $events1m | ForEach-Object {
            $output += "Tijd: $($_.TimeCreated) | Level: $($_.LevelDisplayName) | Bron: $($_.ProviderName)"
            $output += "  Bericht: $($_.Message.Substring(0, [Math]::Min(300, $_.Message.Length)))"
            $output += ""
        }
    } else {
        $output += "Geen relevante fouten gevonden"
    }
} catch {
    $output += "Kon Windows Event Log niet lezen: $_"
}
$output += ""

# --- Applicatie Event Log OpenVPN ---
$output += "--- Applicatie Event Log OpenVPN (laatste maand) ---"
try {
    $appEvents = Get-WinEvent -FilterHashtable @{
        LogName   = 'Application'
        StartTime = (Get-Date).AddDays(-30)
    } -ErrorAction SilentlyContinue | Where-Object {
        $_.ProviderName -like "*OpenVPN*" -or 
        $_.Message -match "OpenVPN|tun error|protect.*socket"
    } | Select-Object -First 20
    if ($appEvents) {
        $output += "Aantal: $($appEvents.Count)"
        $appEvents | ForEach-Object {
            $output += "Tijd: $($_.TimeCreated) | Level: $($_.LevelDisplayName) | Bron: $($_.ProviderName)"
            $output += "  Bericht: $($_.Message.Substring(0, [Math]::Min(300, $_.Message.Length)))"
            $output += ""
            if ($_.Level -le 2) {
                $issues += "PROBLEEM: OpenVPN App Event fout om $($_.TimeCreated)"
            }
        }
    } else {
        $output += "Geen OpenVPN events gevonden in Applicatie log"
    }
} catch {
    $output += "Kon Applicatie Event Log niet lezen: $_"
}
$output += ""

# === SAMENVATTING ===
$output += "=== SAMENVATTING ==="
if ($issues.Count -eq 0) {
    $output += "RESULTAAT: Geen problemen gevonden"
    $exitCode = 0
} else {
    $output += "RESULTAAT: $($issues.Count) probleem/waarschuwing(en) gevonden:"
    $issues | ForEach-Object { $output += "  >> $_" }
    $hardIssues = $issues | Where-Object {$_ -like "PROBLEEM:*"}
    if ($hardIssues) {
        $exitCode = 2
    } else {
        $exitCode = 1
    }
}
$output += ""
$output += "Log weggeschreven naar: $logPath"
$output += "=== DONE ==="

# Wegschrijven naar C:\Temp én output naar NinjaOne
$output | Out-File -FilePath $logPath -Encoding UTF8
$output | ForEach-Object { Write-Output $_ }

exit $exitCode