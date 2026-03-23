#Requires -Version 5.1
<#
.SYNOPSIS
    Resolve public DNS records and import them into Active Directory DNS.

.DESCRIPTION
    Reads a list of FQDNs from a CSV file, resolves each one via Google DNS (8.8.8.8)
    using dig, and adds the result as an A or CNAME record in Active Directory DNS.

    The script shows what it found and what it would create before doing anything.
    Pass -Apply to actually write the records.

    Required CSV column:
      FQDN    Fully qualified domain name to resolve (e.g. briefings.vias.be)

    Resolution logic per FQDN:
      1. Check for CNAME  → if found, creates a CNAME record in AD DNS
      2. Check for A      → if found, creates an A record in AD DNS
      3. No answer        → reported as unresolvable, skipped

.PARAMETER CsvPath
    Path to the CSV file with a FQDN column.

.PARAMETER ZoneName
    The AD DNS zone to add records to (e.g. vias.be).

.PARAMETER DnsServer
    DNS server to create the records on. Default: localhost.

.PARAMETER Ttl
    Time-to-live for the records in seconds. Default: 3600 (1 hour).

.PARAMETER Apply
    Actually create the DNS records in AD. Without this switch runs in dry-run mode.

.EXAMPLE
    # Check what would be created (dry run)
    .\Import-DnsRecords.ps1 -CsvPath .\vias-dns.csv -ZoneName vias.be

.EXAMPLE
    # Resolve and import into AD DNS
    .\Import-DnsRecords.ps1 -CsvPath .\vias-dns.csv -ZoneName vias.be -Apply

.EXAMPLE
    # Remote DNS server
    .\Import-DnsRecords.ps1 -CsvPath .\vias-dns.csv -ZoneName vias.be -DnsServer dc01.vias.be -Apply
#>
[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string] $CsvPath,

    [Parameter(Mandatory)]
    [string] $ZoneName,

    [string] $DnsServer = 'localhost',

    [int] $Ttl = 3600,

    [switch] $Apply
)

# ── Preflight ─────────────────────────────────────────────────────────────────
$digCmd = Get-Command dig -ErrorAction SilentlyContinue
if (-not $digCmd) {
    Write-Error "'dig' not found. Install BIND tools: choco install bind-toolsonly"
    exit 1
}

if ($Apply -and -not (Get-Module -Name DnsServer -ListAvailable)) {
    Write-Error "DnsServer module not found. Run on a Windows DNS Server or install RSAT: Add-WindowsFeature RSAT-DNS-Server"
    exit 1
}

# ── Load CSV ──────────────────────────────────────────────────────────────────
$rows = Import-Csv -Path $CsvPath

if ($rows.Count -eq 0) {
    Write-Error "CSV is empty."
    exit 1
}

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
Write-Host "  Zone       : $ZoneName"
Write-Host "  DNS server : $DnsServer"
Write-Host "  TTL        : $Ttl seconds"
Write-Host "  Records    : $($rows.Count) FQDN(s)"
Write-Host ""

if (-not $Apply) {
    Write-Host "  ================================================" -ForegroundColor Yellow
    Write-Host "   DRY RUN MODE — no records will be created" -ForegroundColor Yellow
    Write-Host "   Add -Apply to write the records to AD DNS." -ForegroundColor Yellow
    Write-Host "  ================================================" -ForegroundColor Yellow
    Write-Host ""
}

Write-Host "  --- Resolving via Google DNS (8.8.8.8) ---" -ForegroundColor Cyan
Write-Host ""

$ttlSpan     = [System.TimeSpan]::FromSeconds($Ttl)
$commonParams = @{ ZoneName = $ZoneName; ComputerName = $DnsServer }

$created = 0
$skipped = 0
$errors  = 0

foreach ($row in $rows) {
    $fqdn = ($row.$fqdnCol).Trim()

    if (-not $fqdn) { continue }

    # Derive host name within zone (strip zone suffix)
    if ($fqdn -notmatch "\.$([regex]::Escape($ZoneName))$") {
        Write-Host ("  [SKIP] {0,-40} : not in zone '{1}'" -f $fqdn, $ZoneName) -ForegroundColor DarkYellow
        $skipped++
        continue
    }
    $hostName = $fqdn -replace "\.$([regex]::Escape($ZoneName))$", ''

    # ── Resolve: CNAME first, then A ──────────────────────────────────────────
    $cnameResult = Resolve-Public -Fqdn $fqdn -Type 'CNAME'
    $aResult     = Resolve-Public -Fqdn $fqdn -Type 'A'

    if ($cnameResult) {
        $target = $cnameResult | Select-Object -First 1
        $targetClean = $target.TrimEnd('.')

        Write-Host ("  [CNAME] {0,-40} -> {1}" -f $fqdn, $targetClean) -NoNewline

        if (-not $Apply) {
            Write-Host "  [DRY RUN]" -ForegroundColor DarkYellow
            $skipped++
            continue
        }

        $existing = Get-DnsServerResourceRecord @commonParams -Name $hostName -RRType CName -ErrorAction SilentlyContinue
        if ($existing) {
            Write-Host "  [EXISTS]" -ForegroundColor DarkGray
            $skipped++
            continue
        }

        try {
            Add-DnsServerResourceRecordCName @commonParams `
                -Name $hostName `
                -HostNameAlias $target `
                -TimeToLive $ttlSpan `
                -ErrorAction Stop
            Write-Host "  [CREATED]" -ForegroundColor Green
            $created++
        } catch {
            Write-Host ""
            Write-Host ("  [ERROR]   {0}" -f $_.Exception.Message) -ForegroundColor Red
            $errors++
        }

    } elseif ($aResult) {
        $ip = $aResult | Select-Object -First 1

        Write-Host ("  [A]     {0,-40} -> {1}" -f $fqdn, $ip) -NoNewline

        if (-not $Apply) {
            Write-Host "  [DRY RUN]" -ForegroundColor DarkYellow
            $skipped++
            continue
        }

        $existing = Get-DnsServerResourceRecord @commonParams -Name $hostName -RRType A -ErrorAction SilentlyContinue
        if ($existing) {
            Write-Host "  [EXISTS]" -ForegroundColor DarkGray
            $skipped++
            continue
        }

        try {
            Add-DnsServerResourceRecordA @commonParams `
                -Name $hostName `
                -IPv4Address $ip `
                -TimeToLive $ttlSpan `
                -ErrorAction Stop
            Write-Host "  [CREATED]" -ForegroundColor Green
            $created++
        } catch {
            Write-Host ""
            Write-Host ("  [ERROR]   {0}" -f $_.Exception.Message) -ForegroundColor Red
            $errors++
        }

    } else {
        Write-Host ("  [UNRESOLVED] {0,-40} : no public DNS answer" -f $fqdn) -ForegroundColor Red
        $errors++
    }
}

# ── Summary ───────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Summary" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
if ($Apply) {
    Write-Host "  Created     : $created" -ForegroundColor Green
    Write-Host "  Skipped     : $skipped" -ForegroundColor DarkGray
    Write-Host "  Unresolved  : $errors"  -ForegroundColor $(if ($errors -gt 0) { 'Red' } else { 'DarkGray' })
} else {
    Write-Host "  Would process : $($rows.Count) FQDN(s)" -ForegroundColor Yellow
    Write-Host "  Run with -Apply to write the records to AD DNS." -ForegroundColor DarkYellow
}
Write-Host ""
