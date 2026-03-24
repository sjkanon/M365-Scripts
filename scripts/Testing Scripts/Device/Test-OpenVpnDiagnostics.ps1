$output = @()
$issues = @()
$output += "=== OPENVPN DIAGNOSTICS ==="
$output += "Datum: $(Get-Date)"
$output += ""

# --- Wintun/TAP Adapters (PnP) ---
$output += "--- Wintun/TAP Adapters (PnP) ---"
$wintun = Get-PnpDevice | Where-Object {$_.FriendlyName -like "*Wintun*" -or $_.FriendlyName -like "*TAP*"}
if ($wintun) {
    $wintun | ForEach-Object {
        $output += "Status: $($_.Status) | Naam: $($_.FriendlyName) | ID: $($_.InstanceId)"
        if ($_.Status -ne "OK") {
            $issues += "PROBLEEM: Wintun/TAP adapter '$($_.FriendlyName)' heeft status '$($_.Status)'"
        }
    }
} else {
    $output += "GEEN Wintun/TAP adapters gevonden"
    $issues += "PROBLEEM: Geen Wintun/TAP adapter aanwezig - driver mogelijk niet geinstalleerd"
}
$output += ""

# --- Virtual Network Adapters ---
$output += "--- Virtual Network Adapters ---"
$adapters = Get-NetAdapter | Where-Object {$_.InterfaceDescription -like "*Wintun*" -or $_.InterfaceDescription -like "*TAP*" -or $_.InterfaceDescription -like "*Virtual*" -or $_.InterfaceDescription -like "*VPN*" -or $_.InterfaceDescription -like "*DCO*"}
if ($adapters) {
    $adapters | ForEach-Object {
        $output += "Naam: $($_.Name) | Omschrijving: $($_.InterfaceDescription) | Status: $($_.Status)"
    }
} else {
    $output += "GEEN virtuele adapters gevonden"
    $issues += "WAARSCHUWING: Geen virtuele netwerkadapter zichtbaar"
}
$output += ""

# --- Network Profiles ---
$output += "--- Network Profiles ---"
Get-NetConnectionProfile | ForEach-Object {
    $output += "Netwerk: $($_.Name) | Adapter: $($_.InterfaceAlias) | Categorie: $($_.NetworkCategory)"
    if ($_.NetworkCategory -eq "Public" -and ($_.InterfaceAlias -like "*Wintun*" -or $_.InterfaceAlias -like "*TAP*" -or $_.InterfaceAlias -like "*OpenVPN*")) {
        $issues += "PROBLEEM: VPN adapter '$($_.InterfaceAlias)' staat op Public netwerk - moet Private zijn"
    }
}
$output += ""

# --- VPN Software ---
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
            $issues += "WAARSCHUWING: Mogelijk conflicterende VPN software gevonden: '$($_.Name)'"
        }
    }
} else {
    $output += "Geen VPN software gevonden"
    $issues += "PROBLEEM: OpenVPN Connect niet gevonden in geinstalleerde software"
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
        $issues += "WAARSCHUWING: Hyper-V is ingeschakeld - kan conflicteren met Wintun driver"
    }
}
$output += ""

# --- PnPUtil ---
$output += "--- Wintun/TAP in PnPUtil ---"
$pnpOutput = pnputil /enum-drivers 2>&1
$lines = $pnpOutput -split "`n"
$foundInPnp = $false
for ($i = 0; $i -lt $lines.Count; $i++) {
    if ($lines[$i] -match "wintun|tap") {
        $foundInPnp = $true
        $start = [Math]::Max(0, $i - 1)
        $end = [Math]::Min($lines.Count - 1, $i + 3)
        $output += $lines[$start..$end]
    }
}
if (-not $foundInPnp) {
    $output += "Geen Wintun/TAP driver gevonden in PnPUtil"
    $issues += "PROBLEEM: Wintun/TAP driver niet geregistreerd in PnPUtil"
}
$output += ""

# --- OpenVPN Services ---
$output += "--- OpenVPN Services ---"
$services = Get-Service | Where-Object {$_.DisplayName -like "*OpenVPN*" -or $_.DisplayName -like "*Wintun*"}
if ($services) {
    $services | ForEach-Object {
        $output += "Service: $($_.DisplayName) | Status: $($_.Status) | Starttype: $($_.StartType)"
        if ($_.Status -ne "Running") {
            $issues += "WAARSCHUWING: Service '$($_.DisplayName)' is niet actief (Status: $($_.Status))"
        }
    }
} else {
    $output += "Geen OpenVPN services gevonden"
    $issues += "PROBLEEM: Geen OpenVPN service aanwezig"
}
$output += ""

# --- OpenVPN Connect Logs ---
$output += "--- OpenVPN Connect Logs ---"
$logPaths = @(
    "$env:LOCALAPPDATA\OpenVPN Connect\logs",
    "C:\Users\*\AppData\Local\OpenVPN Connect\logs"
)

$logFound = $false
foreach ($path in $logPaths) {
    $resolvedPaths = Resolve-Path $path -ErrorAction SilentlyContinue
    foreach ($resolvedPath in $resolvedPaths) {
        if (Test-Path $resolvedPath) {
            $logFiles = Get-ChildItem -Path $resolvedPath -Filter "*.log" | Sort-Object LastWriteTime -Descending | Select-Object -First 3
            foreach ($logFile in $logFiles) {
                $logFound = $true
                $output += ""
                $output += ">> Logbestand: $($logFile.FullName) (Laatst gewijzigd: $($logFile.LastWriteTime))"
                
                $logContent = Get-Content $logFile.FullName -Tail 50 -ErrorAction SilentlyContinue
                $relevantLines = $logContent | Where-Object {
                    $_ -match "error|warn|fail|timeout|disconnect|connect|tun|tap|wintun|dco|protect|socket|fatal" 
                }
                
                if ($relevantLines) {
                    $output += "Relevante log regels:"
                    $relevantLines | ForEach-Object { 
                        $output += "  $_"
                        if ($_ -match "protect\(\) method") {
                            $issues += "PROBLEEM: 'protect() method' fout gevonden in log - TAP/Wintun driver conflict"
                        }
                        if ($_ -match "general tun error") {
                            $issues += "PROBLEEM: 'General tun error' gevonden in log - tunnel interface kan niet aangemaakt worden"
                        }
                        if ($_ -match "access.denied|access denied") {
                            $issues += "PROBLEEM: 'Access Denied' in log - rechtenprobleem op driver of adapter"
                        }
                        if ($_ -match "fatal") {
                            $issues += "PROBLEEM: Fatale fout gevonden in log: $_"
                        }
                    }
                } else {
                    $output += "Geen relevante foutmeldingen gevonden in laatste 50 regels"
                }
            }
        }
    }
}

if (-not $logFound) {
    $output += "Geen OpenVPN Connect logbestanden gevonden"
    $issues += "WAARSCHUWING: Geen logbestanden gevonden - OpenVPN Connect mogelijk nooit gestart"
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
        $_.Message -match "VPN|TAP|Wintun|OpenVPN|tun|network adapter|DCO"
    } | Select-Object -First 10

    if ($events24u) {
        $events24u | ForEach-Object {
            $output += "Tijd: $($_.TimeCreated) | Level: $($_.LevelDisplayName) | Bron: $($_.ProviderName)"
            $output += "  Bericht: $($_.Message.Substring(0, [Math]::Min(200, $_.Message.Length)))"
            $output += ""
            if ($_.Level -le 2) {
                $issues += "PROBLEEM: Windows Event Log fout van '$($_.ProviderName)' om $($_.TimeCreated)"
            }
        }
    } else {
        $output += "Geen relevante fouten gevonden in laatste 24 uur"
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
        $_.Message -match "VPN|TAP|Wintun|OpenVPN|tun|network adapter|DCO"
    } | Select-Object -First 20

    if ($events1m) {
        $output += "Aantal gevonden events: $($events1m.Count)"
        $output += ""
        $events1m | ForEach-Object {
            $output += "Tijd: $($_.TimeCreated) | Level: $($_.LevelDisplayName) | Bron: $($_.ProviderName)"
            $output += "  Bericht: $($_.Message.Substring(0, [Math]::Min(200, $_.Message.Length)))"
            $output += ""
        }
    } else {
        $output += "Geen relevante fouten gevonden in laatste maand"
    }
} catch {
    $output += "Kon Windows Event Log niet lezen: $_"
}
$output += ""

# --- SAMENVATTING ---
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
$output += "=== DONE ==="

$output | ForEach-Object { Write-Output $_ }

exit $exitCode