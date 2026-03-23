#Requires -Version 5.1
<#
.SYNOPSIS
    Create DNS records in Active Directory DNS from a CSV file.

.DESCRIPTION
    Reads a CSV file and adds A and CNAME records to an Active Directory-integrated DNS zone.
    Defaults to dry-run — pass -Apply to create records.

    After creation, use -Verify to validate all records against Google DNS (8.8.8.8) using dig.
    -Verify can also be used standalone to check existing records without creating anything.

    Skips records that already exist (idempotent — safe to re-run).

    Required CSV columns:
      Name    Host name within the zone (e.g. 'briefings' or 'www.briefings')
      Type    Record type: A or CNAME
      Value   IP address for A records, target FQDN for CNAME records.
              Leave empty for A records to use -DefaultIP.

    Example CSV:
      Name,Type,Value
      briefings,A,
      helpdesk,A,1.2.3.5
      www.briefings,CNAME,briefings.vias.be

.PARAMETER CsvPath
    Path to the CSV file containing the DNS records to create.

.PARAMETER ZoneName
    DNS zone name (e.g. vias.be).

.PARAMETER DefaultIP
    Default IP address for A records that have no Value in the CSV.

.PARAMETER DnsServer
    DNS server to create the records on. Default: localhost.

.PARAMETER Ttl
    Time-to-live for all records in seconds. Default: 3600 (1 hour).

.PARAMETER Apply
    Actually create the DNS records. Without this switch the script runs in dry-run mode.

.PARAMETER Verify
    Validate all records against Google DNS (8.8.8.8) using dig.
    Runs automatically after -Apply. Can also be used standalone.

.EXAMPLE
    # Dry run
    .\Add-DnsRecords.ps1 -CsvPath .\vias-dns.csv -ZoneName vias.be -DefaultIP 1.2.3.4

.EXAMPLE
    # Create and verify
    .\Add-DnsRecords.ps1 -CsvPath .\vias-dns.csv -ZoneName vias.be -DefaultIP 1.2.3.4 -Apply -Verify

.EXAMPLE
    # Verify only (no creation)
    .\Add-DnsRecords.ps1 -CsvPath .\vias-dns.csv -ZoneName vias.be -DefaultIP 1.2.3.4 -Verify

.EXAMPLE
    # Remote DNS server
    .\Add-DnsRecords.ps1 -CsvPath .\records.csv -ZoneName contoso.com -DefaultIP 10.0.0.10 -DnsServer dc01.contoso.com -Apply -Verify
#>
[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string] $CsvPath,

    [Parameter(Mandatory)]
    [string] $ZoneName,

    [string] $DefaultIP,

    [string] $DnsServer = 'localhost',

    [int] $Ttl = 3600,

    [switch] $Apply,

    [switch] $Verify
)

# ── Validation ────────────────────────────────────────────────────────────────
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

$cols = $rows[0].PSObject.Properties.Name
$nameCol  = $cols | Where-Object { $_ -match '^name$' }  | Select-Object -First 1
$typeCol  = $cols | Where-Object { $_ -match '^type$' }  | Select-Object -First 1
$valueCol = $cols | Where-Object { $_ -match '^value$' } | Select-Object -First 1

if (-not $nameCol -or -not $typeCol -or -not $valueCol) {
    Write-Error "CSV must have columns: Name, Type, Value. Found: $($cols -join ', ')"
    exit 1
}

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Add-DnsRecords" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Zone       : $ZoneName"
Write-Host "  DNS server : $DnsServer"
if ($DefaultIP) { Write-Host "  Default IP : $DefaultIP" }
Write-Host "  TTL        : $Ttl seconds"
Write-Host "  Records    : $($rows.Count) row(s) in CSV"
Write-Host ""

if (-not $Apply -and -not $Verify) {
    Write-Host "  ================================================" -ForegroundColor Yellow
    Write-Host "   DRY RUN MODE — no records will be created" -ForegroundColor Yellow
    Write-Host "   Add -Apply to create the records." -ForegroundColor Yellow
    Write-Host "  ================================================" -ForegroundColor Yellow
    Write-Host ""
}

# ── Create records ────────────────────────────────────────────────────────────
if ($Apply -or -not $Verify) {
    $created = 0
    $skipped = 0
    $errors  = 0

    $ttlSpan     = [System.TimeSpan]::FromSeconds($Ttl)
    $commonParams = @{ ZoneName = $ZoneName; ComputerName = $DnsServer }

    foreach ($row in $rows) {
        $name  = ($row.$nameCol).Trim()
        $type  = ($row.$typeCol).Trim().ToUpper()
        $value = ($row.$valueCol).Trim()
        $fqdn  = "$name.$ZoneName"

        # Resolve value for A records
        if ($type -eq 'A' -and -not $value) {
            if (-not $DefaultIP) {
                Write-Host ("  [SKIP] A  {0,-35} : no Value and no -DefaultIP supplied" -f $fqdn) -ForegroundColor DarkYellow
                $skipped++
                continue
            }
            $value = $DefaultIP
        }

        if (-not $name -or $type -notin 'A','CNAME') {
            Write-Host ("  [SKIP] {0,-40} : invalid name or unsupported type '{1}'" -f $fqdn, $type) -ForegroundColor DarkYellow
            $skipped++
            continue
        }

        if (-not $Apply) {
            Write-Host ("  [DRY RUN] {0}  {1,-35} -> {2}" -f $type.PadRight(5), $fqdn, $value) -ForegroundColor DarkYellow
            $skipped++
            continue
        }

        switch ($type) {
            'A' {
                $existing = Get-DnsServerResourceRecord @commonParams -Name $name -RRType A -ErrorAction SilentlyContinue
                if ($existing) {
                    Write-Host ("  [EXISTS]  A     {0,-35} -> {1}" -f $fqdn, ($existing.RecordData.IPv4Address -join ', ')) -ForegroundColor DarkGray
                    $skipped++
                    break
                }
                try {
                    Add-DnsServerResourceRecordA @commonParams -Name $name -IPv4Address $value -TimeToLive $ttlSpan -ErrorAction Stop
                    Write-Host ("  [CREATED] A     {0,-35} -> {1}" -f $fqdn, $value) -ForegroundColor Green
                    $created++
                } catch {
                    Write-Host ("  [ERROR]   A     {0,-35} : {1}" -f $fqdn, $_.Exception.Message) -ForegroundColor Red
                    $errors++
                }
            }
            'CNAME' {
                $existing = Get-DnsServerResourceRecord @commonParams -Name $name -RRType CName -ErrorAction SilentlyContinue
                if ($existing) {
                    Write-Host ("  [EXISTS]  CNAME {0,-35} -> {1}" -f $fqdn, $existing.RecordData.HostNameAlias) -ForegroundColor DarkGray
                    $skipped++
                    break
                }
                # Ensure trailing dot for FQDN
                $target = if ($value.EndsWith('.')) { $value } else { "$value." }
                try {
                    Add-DnsServerResourceRecordCName @commonParams -Name $name -HostNameAlias $target -TimeToLive $ttlSpan -ErrorAction Stop
                    Write-Host ("  [CREATED] CNAME {0,-35} -> {1}" -f $fqdn, $value) -ForegroundColor Green
                    $created++
                } catch {
                    Write-Host ("  [ERROR]   CNAME {0,-35} : {1}" -f $fqdn, $_.Exception.Message) -ForegroundColor Red
                    $errors++
                }
            }
        }
    }

    Write-Host ""
    Write-Host "  ================================================" -ForegroundColor Cyan
    Write-Host "   Summary" -ForegroundColor Cyan
    Write-Host "  ================================================" -ForegroundColor Cyan
    if ($Apply) {
        Write-Host "  Created : $created" -ForegroundColor Green
        Write-Host "  Skipped : $skipped" -ForegroundColor DarkGray
        Write-Host "  Errors  : $errors"  -ForegroundColor $(if ($errors -gt 0) { 'Red' } else { 'DarkGray' })
    } else {
        Write-Host "  Would process : $($rows.Count) record(s)" -ForegroundColor Yellow
        Write-Host "  Run with -Apply to create them." -ForegroundColor DarkYellow
    }
    Write-Host ""
}

# ── Verify via Google DNS (dig) ───────────────────────────────────────────────
if ($Verify) {
    $digCmd = Get-Command dig -ErrorAction SilentlyContinue
    if (-not $digCmd) {
        Write-Host "  [WARN] 'dig' not found. Install BIND tools to use -Verify." -ForegroundColor Yellow
        Write-Host "  Windows: choco install bind-toolsonly  or  https://www.isc.org/download/" -ForegroundColor DarkGray
        Write-Host ""
        exit 0
    }

    $googleDns  = '8.8.8.8'
    $verifyOk   = 0
    $verifyFail = 0

    Write-Host "  ================================================" -ForegroundColor Cyan
    Write-Host "   Verifying against Google DNS ($googleDns)" -ForegroundColor Cyan
    Write-Host "  ================================================" -ForegroundColor Cyan
    Write-Host ""

    if ($Apply) {
        Write-Host "  Note: records were just added to AD DNS. Google DNS reflects the" -ForegroundColor DarkYellow
        Write-Host "  public zone — allow propagation time if results appear unresolved." -ForegroundColor DarkYellow
        Write-Host ""
    }

    function Invoke-Dig {
        param([string]$Name, [string]$Type)
        $output = & dig @$googleDns $Name $Type +short 2>&1
        return ($output | Where-Object { $_ -and $_ -notmatch '^;' })
    }

    foreach ($row in $rows) {
        $name  = ($row.$nameCol).Trim()
        $type  = ($row.$typeCol).Trim().ToUpper()
        $value = ($row.$valueCol).Trim()
        $fqdn  = "$name.$ZoneName"

        if ($type -eq 'A' -and -not $value) { $value = $DefaultIP }
        if (-not $name -or $type -notin 'A','CNAME' -or -not $value) { continue }

        $result = Invoke-Dig -Name $fqdn -Type $type

        switch ($type) {
            'A' {
                if ($result -contains $value) {
                    Write-Host ("  [OK]       A     {0,-35} -> {1}" -f $fqdn, ($result -join ', ')) -ForegroundColor Green
                    $verifyOk++
                } elseif ($result) {
                    Write-Host ("  [MISMATCH] A     {0,-35} -> {1} (expected: {2})" -f $fqdn, ($result -join ', '), $value) -ForegroundColor Yellow
                    $verifyFail++
                } else {
                    Write-Host ("  [FAIL]     A     {0,-35} -> (no answer)" -f $fqdn) -ForegroundColor Red
                    $verifyFail++
                }
            }
            'CNAME' {
                $resultClean = ($result | ForEach-Object { $_.TrimEnd('.') })
                $expected    = $value.TrimEnd('.')
                if ($resultClean -contains $expected) {
                    Write-Host ("  [OK]       CNAME {0,-35} -> {1}" -f $fqdn, ($result -join ', ')) -ForegroundColor Green
                    $verifyOk++
                } elseif ($result) {
                    Write-Host ("  [MISMATCH] CNAME {0,-35} -> {1} (expected: {2})" -f $fqdn, ($result -join ', '), $expected) -ForegroundColor Yellow
                    $verifyFail++
                } else {
                    Write-Host ("  [FAIL]     CNAME {0,-35} -> (no answer)" -f $fqdn) -ForegroundColor Red
                    $verifyFail++
                }
            }
        }
    }

    Write-Host ""
    Write-Host "  ================================================" -ForegroundColor Cyan
    Write-Host "   Verification Summary" -ForegroundColor Cyan
    Write-Host "  ================================================" -ForegroundColor Cyan
    Write-Host "  OK   : $verifyOk"   -ForegroundColor Green
    Write-Host "  Fail : $verifyFail" -ForegroundColor $(if ($verifyFail -gt 0) { 'Red' } else { 'DarkGray' })
    Write-Host ""
}
