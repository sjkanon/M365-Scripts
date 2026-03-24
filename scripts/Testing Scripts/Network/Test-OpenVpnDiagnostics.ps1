#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Diagnose OpenVPN Connect issues on a Windows machine.

.DESCRIPTION
    Collects and evaluates diagnostic information relevant to OpenVPN Connect:
      - Wintun / TAP adapter status (PnP)
      - Virtual network adapter visibility
      - Network profile categories (Public vs Private)
      - Installed VPN software (conflicting apps)
      - Hyper-V / WSL / virtualisation features
      - OpenVPN service status
      - Active network routes
      - DNS configuration
      - Recent OpenVPN entries in the Windows Event Log

    Results are printed to screen. Use -ExportTxt to save the full report to a file.

.PARAMETER ExportTxt
    Save the full diagnostic report to a text file.

.PARAMETER OutputPath
    Custom path for the report file. Implies -ExportTxt.

.EXAMPLE
    .\Test-OpenVpnDiagnostics.ps1

.EXAMPLE
    .\Test-OpenVpnDiagnostics.ps1 -ExportTxt

.EXAMPLE
    .\Test-OpenVpnDiagnostics.ps1 -OutputPath "C:\Temp\vpn-report.txt"
#>
[CmdletBinding()]
param (
    [switch] $ExportTxt,
    [string] $OutputPath
)

# ── Output folder ─────────────────────────────────────────────────────────────
$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

if ($OutputPath) {
    $ExportTxt = $true
} else {
    $OutputPath = Join-Path $outputDir "OpenVpnDiagnostics_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
}

# ── Helpers ───────────────────────────────────────────────────────────────────
$output = [System.Collections.Generic.List[string]]::new()
$issues = [System.Collections.Generic.List[string]]::new()

function Add-Line {
    param([string]$Text = '')
    $output.Add($Text)
    Write-Host $Text
}

function Add-Section {
    param([string]$Title)
    Add-Line
    Add-Line "--- $Title ---"
}

function Add-Issue {
    param([string]$Message, [string]$Level = 'PROBLEEM')
    $issues.Add("[$Level] $Message")
}

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Test-OpenVpnDiagnostics" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

Add-Line "=== OPENVPN DIAGNOSTICS ==="
Add-Line "Datum    : $(Get-Date)"
Add-Line "Computer : $env:COMPUTERNAME"

# ── Wintun / TAP adapters (PnP) ───────────────────────────────────────────────
Add-Section "Wintun/TAP Adapters (PnP)"
$wintun = Get-PnpDevice | Where-Object { $_.FriendlyName -like '*Wintun*' -or $_.FriendlyName -like '*TAP*' }
if ($wintun) {
    $wintun | ForEach-Object {
        Add-Line ("  Status: {0,-8} | Naam: {1} | ID: {2}" -f $_.Status, $_.FriendlyName, $_.InstanceId)
        if ($_.Status -ne 'OK') {
            Add-Issue "Wintun/TAP adapter '$($_.FriendlyName)' heeft status '$($_.Status)'"
        }
    }
} else {
    Add-Line "  GEEN Wintun/TAP adapters gevonden"
    Add-Issue "Geen Wintun/TAP adapter aanwezig — driver mogelijk niet geïnstalleerd"
}

# ── Virtual network adapters ──────────────────────────────────────────────────
Add-Section "Virtual Network Adapters"
$adapters = Get-NetAdapter | Where-Object {
    $_.InterfaceDescription -like '*Wintun*' -or
    $_.InterfaceDescription -like '*TAP*' -or
    $_.InterfaceDescription -like '*Virtual*' -or
    $_.InterfaceDescription -like '*VPN*'
}
if ($adapters) {
    $adapters | ForEach-Object {
        Add-Line ("  Naam: {0} | Omschrijving: {1} | Status: {2}" -f $_.Name, $_.InterfaceDescription, $_.Status)
    }
} else {
    Add-Line "  GEEN virtuele adapters gevonden"
    Add-Issue "Geen virtuele netwerkadapter zichtbaar — OpenVPN Connect mogelijk niet actief tijdens scan" "WAARSCHUWING"
}

# ── Network profiles ──────────────────────────────────────────────────────────
Add-Section "Network Profiles"
Get-NetConnectionProfile | ForEach-Object {
    Add-Line ("  Netwerk: {0} | Adapter: {1} | Categorie: {2}" -f $_.Name, $_.InterfaceAlias, $_.NetworkCategory)
    if ($_.NetworkCategory -eq 'Public' -and (
        $_.InterfaceAlias -like '*Wintun*' -or
        $_.InterfaceAlias -like '*TAP*' -or
        $_.InterfaceAlias -like '*OpenVPN*')) {
        Add-Issue "VPN adapter '$($_.InterfaceAlias)' staat op Public netwerk — moet Private zijn"
    }
}

# ── Installed VPN software ────────────────────────────────────────────────────
Add-Section "VPN Software geïnstalleerd"
try {
    $vpnApps = Get-WmiObject -Class Win32_Product -ErrorAction Stop | Where-Object {
        $_.Name -like '*VPN*' -or
        $_.Name -like '*Cisco*' -or
        $_.Name -like '*WireGuard*' -or
        $_.Name -like '*OpenVPN*'
    }
    if ($vpnApps) {
        $vpnApps | ForEach-Object {
            Add-Line ("  App: {0} | Versie: {1}" -f $_.Name, $_.Version)
            if ($_.Name -notlike '*OpenVPN Connect*') {
                Add-Issue "Mogelijk conflicterende VPN software gevonden: '$($_.Name)'" "WAARSCHUWING"
            }
        }
    } else {
        Add-Line "  Geen VPN software gevonden"
        Add-Issue "OpenVPN Connect niet gevonden in geïnstalleerde software"
    }
} catch {
    Add-Line "  WMI query mislukt: $_"
}

# ── Hyper-V / WSL / virtualisation ───────────────────────────────────────────
Add-Section "Hyper-V / WSL / Virtualisatie"
try {
    $virtFeatures = Get-WindowsOptionalFeature -Online -ErrorAction Stop | Where-Object {
        $_.FeatureName -like '*Hyper-V*' -or
        $_.FeatureName -like '*VirtualMachinePlatform*' -or
        $_.FeatureName -like '*WSL*'
    }
    if ($virtFeatures) {
        $virtFeatures | ForEach-Object {
            Add-Line ("  {0,-40} : {1}" -f $_.FeatureName, $_.State)
            if ($_.State -eq 'Enabled' -and $_.FeatureName -like '*Hyper-V*') {
                Add-Issue "Hyper-V is ingeschakeld — kan conflicteren met Wintun driver" "WAARSCHUWING"
            }
        }
    } else {
        Add-Line "  Geen virtualisatie-features gevonden"
    }
} catch {
    Add-Line "  Ophalen Windows features mislukt: $_"
}

# ── OpenVPN service ───────────────────────────────────────────────────────────
Add-Section "OpenVPN Service"
$vpnServices = Get-Service | Where-Object { $_.DisplayName -like '*OpenVPN*' -or $_.Name -like '*OpenVPN*' }
if ($vpnServices) {
    $vpnServices | ForEach-Object {
        Add-Line ("  Service: {0} | Status: {1} | StartType: {2}" -f $_.DisplayName, $_.Status, $_.StartType)
        if ($_.Status -ne 'Running') {
            Add-Issue "OpenVPN service '$($_.DisplayName)' is niet actief (status: $($_.Status))"
        }
    }
} else {
    Add-Line "  Geen OpenVPN service gevonden"
    Add-Issue "Geen OpenVPN service aanwezig — OpenVPN Connect mogelijk niet correct geïnstalleerd"
}

# ── Active routes ─────────────────────────────────────────────────────────────
Add-Section "Actieve Routes (VPN-gerelateerd)"
$vpnAdapterIndices = (Get-NetAdapter | Where-Object {
    $_.InterfaceDescription -like '*Wintun*' -or
    $_.InterfaceDescription -like '*TAP*' -or
    $_.InterfaceDescription -like '*VPN*'
}).ifIndex

if ($vpnAdapterIndices) {
    $routes = Get-NetRoute | Where-Object { $vpnAdapterIndices -contains $_.InterfaceIndex }
    if ($routes) {
        $routes | ForEach-Object {
            Add-Line ("  {0,-20} via {1,-16} (metric: {2})" -f $_.DestinationPrefix, $_.NextHop, $_.RouteMetric)
        }
    } else {
        Add-Line "  Geen routes via VPN adapter"
        Add-Issue "VPN adapter aanwezig maar geen routes — tunnel mogelijk niet actief" "WAARSCHUWING"
    }
} else {
    Add-Line "  Geen VPN adapter gevonden — routes overgeslagen"
}

# ── DNS configuration ─────────────────────────────────────────────────────────
Add-Section "DNS Configuratie"
Get-NetAdapter | Where-Object { $_.Status -eq 'Up' } | ForEach-Object {
    $dns = (Get-DnsClientServerAddress -InterfaceIndex $_.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue).ServerAddresses
    if ($dns) {
        Add-Line ("  Adapter: {0,-25} | DNS: {1}" -f $_.Name, ($dns -join ', '))
    }
}

# ── Event log ─────────────────────────────────────────────────────────────────
Add-Section "Event Log (laatste 20 OpenVPN entries)"
try {
    $events = Get-WinEvent -FilterHashtable @{ LogName = 'Application'; ProviderName = '*OpenVPN*' } `
        -MaxEvents 20 -ErrorAction Stop
    $events | ForEach-Object {
        Add-Line ("  [{0}] {1} — {2}" -f $_.LevelDisplayName, $_.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss'), $_.Message.Split("`n")[0])
    }
} catch {
    Add-Line "  Geen OpenVPN entries gevonden in Event Log (of toegang geweigerd)"
}

# ── Issues summary ────────────────────────────────────────────────────────────
Add-Line
Add-Line "=== SAMENVATTING ==="
if ($issues.Count -eq 0) {
    Add-Line "  Geen problemen gevonden."
    Write-Host "  Geen problemen gevonden." -ForegroundColor Green
} else {
    $issues | ForEach-Object {
        $color = if ($_ -like '*[PROBLEEM]*') { 'Red' } else { 'Yellow' }
        Add-Line "  $_"
        Write-Host "  $_" -ForegroundColor $color
    }
}
Add-Line

# ── Export ────────────────────────────────────────────────────────────────────
if ($ExportTxt) {
    $output | Set-Content -Path $OutputPath -Encoding UTF8
    Write-Host ""
    Write-Host "  Rapport opgeslagen: $OutputPath" -ForegroundColor Cyan
}

Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ("  Gevonden issues : {0}" -f $issues.Count) -ForegroundColor $(if ($issues.Count -gt 0) { 'Yellow' } else { 'Green' })
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""
