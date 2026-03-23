#Requires -Version 5.1
<#
.SYNOPSIS
    Resolve public DNS records and optionally import them into Active Directory DNS.

.DESCRIPTION
    Reads a list of FQDNs from a CSV file and resolves each one via Google DNS (8.8.8.8)
    using dig. Results are shown on screen.

    Use -ExportCsv to save the resolved records to a CSV — useful for manual review
    or importing into AD DNS by hand.

    Use -Apply to automatically write the resolved records into Active Directory DNS.

    Required CSV column:
      FQDN    Fully qualified domain name to resolve (e.g. briefings.vias.be)

    Resolution logic per FQDN:
      1. Check for CNAME  → if found, records type CNAME + target
      2. Check for A      → if found, records type A + IP address
      3. No answer        → reported as unresolvable, skipped

.PARAMETER CsvPath
    Path to the CSV file with a FQDN column.

.PARAMETER ZoneName
    The AD DNS zone (e.g. vias.be). Only required when using -Apply.

.PARAMETER ExportCsv
    Export resolved records to a CSV file for manual review or import.
    Default path: .\ResolvedRecords_<timestamp>.csv

.PARAMETER ExportPath
    Custom path for the exported CSV. Implies -ExportCsv.

.PARAMETER Apply
    Write resolved records into Active Directory DNS (requires DnsServer module).

.PARAMETER DnsServer
    DNS server to write records to. Default: localhost. Only used with -Apply.

.PARAMETER Ttl
    TTL for created records in seconds. Default: 3600. Only used with -Apply.

.EXAMPLE
    # Resolve and show on screen only
    .\Import-DnsRecords.ps1 -CsvPath .\records.csv -ZoneName vias.be

.EXAMPLE
    # Resolve and export to CSV for manual review
    .\Import-DnsRecords.ps1 -CsvPath .\records.csv -ZoneName vias.be -ExportCsv

.EXAMPLE
    # Resolve and import directly into AD DNS
    .\Import-DnsRecords.ps1 -CsvPath .\records.csv -ZoneName vias.be -Apply

.EXAMPLE
    # Remote DNS server
    .\Import-DnsRecords.ps1 -CsvPath .\records.csv -ZoneName vias.be -DnsServer dc01.vias.be -Apply
#>
[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string] $CsvPath,

    [string] $ZoneName,

    [switch] $ExportCsv,

    [string] $ExportPath = ".\ResolvedRecords_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",

    [switch] $Apply,

    [string] $DnsServer = 'localhost',

    [int] $Ttl = 3600
)

# ── Preflight ─────────────────────────────────────────────────────────────────
if ($ExportPath -ne ".\ResolvedRecords_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv") {
    $ExportCsv = $true
}

$digCmd = Get-Command dig -ErrorAction SilentlyContinue
if (-not $digCmd) {
    Write-Error "'dig' not found. Install BIND tools: choco install bind-toolsonly"
    exit 1
}

if ($Apply -and -not $ZoneName) {
    Write-Error "-ZoneName is required when using -Apply."
    exit 1
}

if ($Apply -and -not (Get-Module -Name DnsServer -ListAvailable)) {
    Write-Error "DnsServer module not found. Run on a Windows DNS Server or install RSAT: Add-WindowsFeature RSAT-DNS-Server"
    exit 1
}

# ── Load CSV ──────────────────────────────────────────────────────────────────
$rows = Import-Csv -Path $CsvPath

if ($rows.Count -eq 0) { Write-Error "CSV is empty."; exit 1 }

$fqdnCol = $rows[0].PSObject.Properties.Name | Where-Object { $_ -match '^fqdn$' } | Select-Object -First 1
if (-not $fqdnCol) {
    Write-Error "CSV must have a 'FQDN' column. Found: $($rows[0].PSObject.Properties.Name -join ', ')"
    exit 1
}

# ── Helper: resolve via dig ───────────────────────────────────────────────────
$googleDns = '8.8.8.8'

function Resolve-Public {
    param([string]$Fqdn, [string]$Type)
    $result = & dig "@$googleDns" $Fqdn $Type +short 2>&1
    return ($result | Where-Object { $_ -and $_ -notmatch '^;' })
}

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Import-DnsRecords" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  CSV        : $CsvPath"
if ($ZoneName) { Write-Host "  Zone       : $ZoneName" }
Write-Host "  Records    : $($rows.Count) FQDN(s)"
if ($Apply)     { Write-Host "  DNS server : $DnsServer" }
if ($ExportCsv) { Write-Host "  Export to  : $ExportPath" }
Write-Host ""
Write-Host "  --- Resolving via Google DNS ($googleDns) ---" -ForegroundColor Cyan
Write-Host ""

# ── Process records ───────────────────────────────────────────────────────────
$ttlSpan      = [System.TimeSpan]::FromSeconds($Ttl)
$commonParams  = @{ ZoneName = $ZoneName; ComputerName = $DnsServer }
$resolved      = [System.Collections.Generic.List[PSCustomObject]]::new()
$created       = 0
$skipped       = 0
$errors        = 0

foreach ($row in $rows) {
    $fqdn = ($row.$fqdnCol).Trim()
    if (-not $fqdn) { continue }

    # Derive host name within zone (only needed for -Apply)
    $hostName = $null
    if ($ZoneName) {
        if ($fqdn -notmatch "\.$([regex]::Escape($ZoneName))$") {
            Write-Host ("  [SKIP]       {0,-42} not in zone '{1}'" -f $fqdn, $ZoneName) -ForegroundColor DarkYellow
            $skipped++
            continue
        }
        $hostName = $fqdn -replace "\.$([regex]::Escape($ZoneName))$", ''
    }

    # Resolve: CNAME first, then A
    $cnameResult = Resolve-Public -Fqdn $fqdn -Type 'CNAME'
    $aResult     = Resolve-Public -Fqdn $fqdn -Type 'A'

    if ($cnameResult) {
        $target      = $cnameResult | Select-Object -First 1
        $targetClean = $target.TrimEnd('.')

        Write-Host ("  [CNAME]      {0,-42} -> {1}" -f $fqdn, $targetClean) -NoNewline

        $resolved.Add([PSCustomObject]@{ FQDN = $fqdn; HostName = $hostName; Type = 'CNAME'; Value = $targetClean })

        if ($Apply) {
            $existing = Get-DnsServerResourceRecord @commonParams -Name $hostName -RRType CName -ErrorAction SilentlyContinue
            if ($existing) {
                Write-Host '  [EXISTS]' -ForegroundColor DarkGray; $skipped++
            } else {
                try {
                    Add-DnsServerResourceRecordCName @commonParams -Name $hostName -HostNameAlias $target -TimeToLive $ttlSpan -ErrorAction Stop
                    Write-Host '  [CREATED]' -ForegroundColor Green; $created++
                } catch {
                    Write-Host ""; Write-Host ("  [ERROR] {0}" -f $_.Exception.Message) -ForegroundColor Red; $errors++
                }
            }
        } else {
            Write-Host ""
        }

    } elseif ($aResult) {
        $ip = $aResult | Select-Object -First 1

        Write-Host ("  [A]          {0,-42} -> {1}" -f $fqdn, $ip) -NoNewline

        $resolved.Add([PSCustomObject]@{ FQDN = $fqdn; HostName = $hostName; Type = 'A'; Value = $ip })

        if ($Apply) {
            $existing = Get-DnsServerResourceRecord @commonParams -Name $hostName -RRType A -ErrorAction SilentlyContinue
            if ($existing) {
                Write-Host '  [EXISTS]' -ForegroundColor DarkGray; $skipped++
            } else {
                try {
                    Add-DnsServerResourceRecordA @commonParams -Name $hostName -IPv4Address $ip -TimeToLive $ttlSpan -ErrorAction Stop
                    Write-Host '  [CREATED]' -ForegroundColor Green; $created++
                } catch {
                    Write-Host ""; Write-Host ("  [ERROR] {0}" -f $_.Exception.Message) -ForegroundColor Red; $errors++
                }
            }
        } else {
            Write-Host ""
        }

    } else {
        Write-Host ("  [UNRESOLVED] {0,-42} no public DNS answer" -f $fqdn) -ForegroundColor Red
        $errors++
    }
}

# ── Export CSV ────────────────────────────────────────────────────────────────
if ($ExportCsv -and $resolved.Count -gt 0) {
    $resolved | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host ""
    Write-Host "  Resolved records saved to: $ExportPath" -ForegroundColor Cyan
    Write-Host "  Columns: FQDN, HostName, Type, Value" -ForegroundColor DarkGray
}

# ── Summary ───────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Summary" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "  Resolved    : $($resolved.Count)" -ForegroundColor Green
Write-Host "  Unresolved  : $errors"             -ForegroundColor $(if ($errors -gt 0) { 'Red' } else { 'DarkGray' })
if ($Apply) {
    Write-Host "  Created     : $created"         -ForegroundColor Green
    Write-Host "  Skipped     : $skipped (already existed)" -ForegroundColor DarkGray
} elseif (-not $ExportCsv) {
    Write-Host ""
    Write-Host "  Use -ExportCsv to save results for manual import." -ForegroundColor DarkYellow
    Write-Host "  Use -Apply to import directly into AD DNS." -ForegroundColor DarkYellow
}
Write-Host ""
