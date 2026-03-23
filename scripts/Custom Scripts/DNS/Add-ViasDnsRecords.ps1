#Requires -Version 5.1
<#
.SYNOPSIS
    Create DNS records for vias.be subdomains in Active Directory DNS.

.DESCRIPTION
    Adds A records and CNAME records for vias.be web properties to an
    Active Directory-integrated DNS zone. Defaults to dry-run — pass -Apply to create records.

    After creation, use -Verify to validate all records against Google DNS (8.8.8.8) using dig.
    -Verify can also be used standalone (without -Apply) to check existing records.

    A records (all pointing to -TargetIP):
      briefings.vias.be
      helpdesk.vias.be
      meetweek.vias.be
      mobisafetyscan.vias.be
      semaindecomptage.vias.be

    CNAME records (pointing to their base subdomain):
      www.briefings.vias.be        → briefings.vias.be
      www.meetweek.vias.be         → meetweek.vias.be
      www.mobisafetyscan.vias.be   → mobisafetyscan.vias.be
      www.semaindecomptage.vias.be → semaindecomptage.vias.be

.PARAMETER TargetIP
    IP address for the A records (e.g. 1.2.3.4).
    Not required when using -Verify only.

.PARAMETER DnsServer
    DNS server to create the records on. Default: localhost.

.PARAMETER ZoneName
    DNS zone name. Default: vias.be.

.PARAMETER Ttl
    Time-to-live for the records in seconds. Default: 3600 (1 hour).

.PARAMETER Apply
    Actually create the DNS records. Without this switch the script runs in dry-run mode.

.PARAMETER Verify
    Validate all records against Google DNS (8.8.8.8) using dig.
    Automatically runs after -Apply. Can also be used standalone to check existing records.

.EXAMPLE
    # Dry run — shows what would be created
    .\Add-ViasDnsRecords.ps1 -TargetIP 1.2.3.4

.EXAMPLE
    # Create and verify
    .\Add-ViasDnsRecords.ps1 -TargetIP 1.2.3.4 -Apply -Verify

.EXAMPLE
    # Only verify existing records (no creation)
    .\Add-ViasDnsRecords.ps1 -TargetIP 1.2.3.4 -Verify

.EXAMPLE
    # Create the records on a remote DNS server
    .\Add-ViasDnsRecords.ps1 -TargetIP 1.2.3.4 -DnsServer dc01.vias.be -Apply -Verify
#>
[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter()]
    [ValidateScript({
        if ($_ -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$') { $true }
        else { throw "TargetIP must be a valid IPv4 address (e.g. 1.2.3.4)" }
    })]
    [string] $TargetIP,

    [string] $DnsServer = 'localhost',

    [string] $ZoneName = 'vias.be',

    [int] $Ttl = 3600,

    [switch] $Apply,

    [switch] $Verify
)

# ── Validation ────────────────────────────────────────────────────────────────
if (-not $Verify -and -not $TargetIP) {
    Write-Error "TargetIP is required unless using -Verify only."
    exit 1
}

if (($Apply -or -not $Verify) -and -not (Get-Module -Name DnsServer -ListAvailable)) {
    Write-Error "DnsServer module not found. Run this script on a Windows DNS Server or install RSAT: Add-WindowsFeature RSAT-DNS-Server"
    exit 1
}

# ── Record definitions ────────────────────────────────────────────────────────
$aRecords = @(
    'briefings'
    'helpdesk'
    'meetweek'
    'mobisafetyscan'
    'semaindecomptage'
)

# CNAME: key = host name in zone, value = FQDN target (without trailing dot for display)
$cnameRecords = [ordered]@{
    'www.briefings'        = 'briefings.vias.be'
    'www.meetweek'         = 'meetweek.vias.be'
    'www.mobisafetyscan'   = 'mobisafetyscan.vias.be'
    'www.semaindecomptage' = 'semaindecomptage.vias.be'
}

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Add-ViasDnsRecords" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""
if ($TargetIP) { Write-Host "  Target IP  : $TargetIP" }
Write-Host "  Zone       : $ZoneName"
Write-Host "  DNS server : $DnsServer"
Write-Host "  TTL        : $Ttl seconds"
Write-Host ""

# ── Skip creation block if only verifying ────────────────────────────────────
if ($Apply -or -not $Verify) {
    if (-not $Apply) {
        Write-Host "  ================================================" -ForegroundColor Yellow
        Write-Host "   DRY RUN MODE — no records will be created" -ForegroundColor Yellow
        Write-Host "   Add -Apply to create the records." -ForegroundColor Yellow
        Write-Host "  ================================================" -ForegroundColor Yellow
        Write-Host ""
    }

    $created = 0
    $skipped = 0
    $errors  = 0

    $ttlSpan    = [System.TimeSpan]::FromSeconds($Ttl)
    $commonParams = @{ ZoneName = $ZoneName; ComputerName = $DnsServer }

    # ── A records ─────────────────────────────────────────────────────────────
    Write-Host "  --- A Records ---" -ForegroundColor Cyan
    Write-Host ""

    foreach ($name in $aRecords) {
        $fqdn = "$name.$ZoneName"

        if (-not $Apply) {
            Write-Host ("  [DRY RUN] A  {0,-35} -> {1}" -f $fqdn, $TargetIP) -ForegroundColor DarkYellow
            $skipped++
            continue
        }

        $existing = Get-DnsServerResourceRecord @commonParams -Name $name -RRType A -ErrorAction SilentlyContinue
        if ($existing) {
            Write-Host ("  [EXISTS]  A  {0,-35} -> {1}" -f $fqdn, ($existing.RecordData.IPv4Address -join ', ')) -ForegroundColor DarkGray
            $skipped++
            continue
        }

        try {
            Add-DnsServerResourceRecordA @commonParams `
                -Name $name `
                -IPv4Address $TargetIP `
                -TimeToLive $ttlSpan `
                -ErrorAction Stop
            Write-Host ("  [CREATED] A  {0,-35} -> {1}" -f $fqdn, $TargetIP) -ForegroundColor Green
            $created++
        } catch {
            Write-Host ("  [ERROR]   A  {0,-35} : {1}" -f $fqdn, $_.Exception.Message) -ForegroundColor Red
            $errors++
        }
    }

    Write-Host ""

    # ── CNAME records ─────────────────────────────────────────────────────────
    Write-Host "  --- CNAME Records ---" -ForegroundColor Cyan
    Write-Host ""

    foreach ($entry in $cnameRecords.GetEnumerator()) {
        $name   = $entry.Key
        $target = $entry.Value
        $fqdn   = "$name.$ZoneName"

        if (-not $Apply) {
            Write-Host ("  [DRY RUN] CNAME {0,-35} -> {1}" -f $fqdn, $target) -ForegroundColor DarkYellow
            $skipped++
            continue
        }

        $existing = Get-DnsServerResourceRecord @commonParams -Name $name -RRType CName -ErrorAction SilentlyContinue
        if ($existing) {
            Write-Host ("  [EXISTS]  CNAME {0,-35} -> {1}" -f $fqdn, $existing.RecordData.HostNameAlias) -ForegroundColor DarkGray
            $skipped++
            continue
        }

        try {
            Add-DnsServerResourceRecordCName @commonParams `
                -Name $name `
                -HostNameAlias "$target." `
                -TimeToLive $ttlSpan `
                -ErrorAction Stop
            Write-Host ("  [CREATED] CNAME {0,-35} -> {1}" -f $fqdn, $target) -ForegroundColor Green
            $created++
        } catch {
            Write-Host ("  [ERROR]   CNAME {0,-35} : {1}" -f $fqdn, $_.Exception.Message) -ForegroundColor Red
            $errors++
        }
    }

    # ── Summary ───────────────────────────────────────────────────────────────
    Write-Host ""
    Write-Host "  ================================================" -ForegroundColor Cyan
    Write-Host "   Summary" -ForegroundColor Cyan
    Write-Host "  ================================================" -ForegroundColor Cyan
    if ($Apply) {
        Write-Host "  Created : $created" -ForegroundColor Green
        Write-Host "  Skipped : $skipped (already existed)" -ForegroundColor DarkGray
        Write-Host "  Errors  : $errors" -ForegroundColor $(if ($errors -gt 0) { 'Red' } else { 'DarkGray' })
    } else {
        Write-Host "  Would create : $($aRecords.Count) A record(s) + $($cnameRecords.Count) CNAME record(s)" -ForegroundColor Yellow
        Write-Host "  Run with -Apply to create them." -ForegroundColor DarkYellow
    }
    Write-Host ""
}

# ── Verify via Google DNS (dig) ───────────────────────────────────────────────
if ($Verify) {
    # Check dig availability
    $digCmd = Get-Command dig -ErrorAction SilentlyContinue
    if (-not $digCmd) {
        Write-Host "  [WARN] 'dig' not found. Install BIND tools or use nslookup manually." -ForegroundColor Yellow
        Write-Host "  Download: https://www.isc.org/download/" -ForegroundColor DarkGray
        Write-Host ""
        exit 0
    }

    $googleDns   = '8.8.8.8'
    $verifyOk    = 0
    $verifyFail  = 0

    Write-Host "  ================================================" -ForegroundColor Cyan
    Write-Host "   Verifying against Google DNS ($googleDns)" -ForegroundColor Cyan
    Write-Host "  ================================================" -ForegroundColor Cyan
    Write-Host ""

    # Warn if just applied — Google DNS may not have propagated yet
    if ($Apply) {
        Write-Host "  Note: Records were just created in AD DNS. Google DNS resolves" -ForegroundColor DarkYellow
        Write-Host "  the public zone — allow time for DNS propagation if needed." -ForegroundColor DarkYellow
        Write-Host ""
    }

    # Helper: run dig and extract the answer
    function Invoke-Dig {
        param([string]$Name, [string]$Type)
        # +short returns only the answer value(s)
        $output = & dig @$googleDns $Name $Type +short 2>&1
        return ($output | Where-Object { $_ -and $_ -notmatch '^;' })
    }

    Write-Host "  --- A Records ---" -ForegroundColor Cyan
    Write-Host ""

    foreach ($name in $aRecords) {
        $fqdn   = "$name.$ZoneName"
        $result = Invoke-Dig -Name $fqdn -Type 'A'

        if ($result -contains $TargetIP) {
            Write-Host ("  [OK]   A  {0,-35} -> {1}" -f $fqdn, ($result -join ', ')) -ForegroundColor Green
            $verifyOk++
        } elseif ($result) {
            Write-Host ("  [MISMATCH] A  {0,-35} -> {1} (expected: {2})" -f $fqdn, ($result -join ', '), $TargetIP) -ForegroundColor Yellow
            $verifyFail++
        } else {
            Write-Host ("  [FAIL] A  {0,-35} -> (no answer)" -f $fqdn) -ForegroundColor Red
            $verifyFail++
        }
    }

    Write-Host ""
    Write-Host "  --- CNAME Records ---" -ForegroundColor Cyan
    Write-Host ""

    foreach ($entry in $cnameRecords.GetEnumerator()) {
        $name     = $entry.Key
        $expected = $entry.Value
        $fqdn     = "$name.$ZoneName"

        # dig CNAME +short returns the canonical name with trailing dot
        $result = Invoke-Dig -Name $fqdn -Type 'CNAME'

        # Normalize: strip trailing dot for comparison
        $resultClean = ($result | ForEach-Object { $_.TrimEnd('.') })

        if ($resultClean -contains $expected) {
            Write-Host ("  [OK]   CNAME {0,-35} -> {1}" -f $fqdn, ($result -join ', ')) -ForegroundColor Green
            $verifyOk++
        } elseif ($result) {
            Write-Host ("  [MISMATCH] CNAME {0,-35} -> {1} (expected: {2})" -f $fqdn, ($result -join ', '), $expected) -ForegroundColor Yellow
            $verifyFail++
        } else {
            Write-Host ("  [FAIL] CNAME {0,-35} -> (no answer)" -f $fqdn) -ForegroundColor Red
            $verifyFail++
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
