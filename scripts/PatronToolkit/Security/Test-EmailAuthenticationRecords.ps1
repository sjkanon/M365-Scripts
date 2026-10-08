#Requires -Version 7.0
<#
.SYNOPSIS
    Validate SPF and DMARC DNS records for one or more domains.

.DESCRIPTION
    Resolves and validates the email authentication DNS records that anti-spoofing
    depends on:
      - SPF (TXT record starting "v=spf1") — checks it exists, includes
        "include:spf.protection.outlook.com" when the domain routes mail through
        Exchange Online, and does not end in the permissive "+all" mechanism
      - DMARC (TXT record at _dmarc.<domain> starting "v=DMARC1") — checks it exists,
        reports its policy (none/quarantine/reject) and whether an aggregate report
        address (rua=) is configured

    DKIM is intentionally out of scope here — use Test-DkimConfig.ps1 (in
    scripts/Exchange/) for DKIM signing configuration and selector validation.

    Domains can be supplied explicitly, or auto-discovered from Entra ID verified
    domains via Microsoft Graph when -TenantId is given (or an existing Graph session is
    active). Results are exported to CSV.

.PARAMETER Domain
    One or more domains to check. If omitted, verified domains are auto-discovered via
    Microsoft Graph.

.PARAMETER OutputPath
    CSV report path. Defaults to .\EmailAuthenticationRecords_<timestamp>.csv.

.PARAMETER TenantId
    Entra ID tenant ID or domain, used only for auto-discovering domains via Graph when
    -Domain is omitted. Defaults to the GDAP customer (load.config.ps1) or your own
    tenant. Required for app-only sign-in.

.PARAMETER ClientId
    App registration for app-only sign-in (with -CertificateThumbprint). Without it the
    script signs in delegated, as you. Only used when -Domain is omitted.

.PARAMETER CertificateThumbprint
    Certificate thumbprint for app-only sign-in with -ClientId.

.PARAMETER AppOnly
    App-only sign-in with ClientId and CertificateThumbprint from graph.appid.json.

.EXAMPLE
    .\Test-EmailAuthenticationRecords.ps1 -Domain "contoso.com"

.EXAMPLE
    .\Test-EmailAuthenticationRecords.ps1

.NOTES
    Windows-only (uses Resolve-DnsName). Inspired by the capability list of the retired
    directorcia/patron toolkit (o365-dns-get.ps1) — the original dumped every DNS record
    type for a domain with no interpretation. This script is deliberately narrower and
    opinionated: it only checks the two records that matter for anti-spoofing and reports
    pass/fail against known-good patterns.

    Sign-in: only when -Domain is omitted, to Microsoft Graph through
    scripts\Startup\Connect-M365.ps1 - delegated by default (scope Domain.Read.All;
    device code / GDAP customer per load.config.ps1), app-only with
    -ClientId/-CertificateThumbprint or -AppOnly (application permission Domain.Read.All).
    An existing fitting Graph session is reused and left connected.
#>
[CmdletBinding()]
param(
    [string[]] $Domain,
    [string] $OutputPath,
    [string] $TenantId,
    [string] $ClientId,
    [string] $CertificateThumbprint,
    [switch] $AppOnly
)

. (Join-Path $PSScriptRoot '..\..\Startup\Connect-M365.ps1')

if (-not (Get-Command Resolve-DnsName -ErrorAction SilentlyContinue)) {
    Write-Host "  [ERROR] Resolve-DnsName is not available on this platform (Windows-only). Aborting." -ForegroundColor Red
    exit 1
}

# ── Output folder ─────────────────────────────────────────────────────────────
$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   SPF / DMARC Record Validation" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

# ── Domain discovery ──────────────────────────────────────────────────────────
$graph = $null
if (-not $Domain) {
    $graph = Connect-M365Graph -Scopes 'Domain.Read.All' -TenantId $TenantId `
        -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -AppOnly:$AppOnly
    Write-Host "  Discovering verified domains via Microsoft Graph..." -ForegroundColor DarkGray
    try {
        $Domain = @(Get-MgDomain -All -ErrorAction Stop | Where-Object { $_.IsVerified } | Select-Object -ExpandProperty Id)
    } catch {
        Write-Host "  [ERROR] Could not list the tenant's domains: $($_.Exception.Message)" -ForegroundColor Red
        Disconnect-M365Graph $graph
        exit 1
    }
    Write-Host "  Found $($Domain.Count) verified domain(s)." -ForegroundColor DarkGray
    Write-Host ""
}

$results = [System.Collections.Generic.List[PSObject]]::new()

foreach ($d in $Domain) {
    Write-Host "  Checking $d" -ForegroundColor Cyan

    # ── SPF ───────────────────────────────────────────────────────────────────
    $spfRecord = $null
    try {
        $txtRecords = Resolve-DnsName -Name $d -Type TXT -ErrorAction Stop
        $spfRecord = ($txtRecords | Where-Object { $_.Strings -match '^v=spf1' } | Select-Object -First 1).Strings -join ''
    } catch {}

    if (-not $spfRecord) {
        Write-Host "    [FAIL] SPF record not found" -ForegroundColor Red
        $results.Add([PSCustomObject]@{ Domain = $d; RecordType = 'SPF'; Found = $false; Value = ''; Flag = 'Fail'; Detail = 'No SPF TXT record found' })
    } else {
        $includesEXO = $spfRecord -match 'include:spf\.protection\.outlook\.com'
        $permissive = $spfRecord -match '\+all\s*$'
        $flag = if ($permissive) { 'Fail' } elseif (-not $includesEXO) { 'Warn' } else { 'Pass' }
        $detail = if ($permissive) { 'Ends in "+all" — allows any server to send as this domain' }
                  elseif (-not $includesEXO) { 'Does not include spf.protection.outlook.com — verify this is intentional' }
                  else { 'OK' }
        Write-Host "    [$flag] SPF: $spfRecord" -ForegroundColor $(if ($flag -eq 'Pass') { 'Green' } elseif ($flag -eq 'Warn') { 'Yellow' } else { 'Red' })
        $results.Add([PSCustomObject]@{ Domain = $d; RecordType = 'SPF'; Found = $true; Value = $spfRecord; Flag = $flag; Detail = $detail })
    }

    # ── DMARC ─────────────────────────────────────────────────────────────────
    $dmarcRecord = $null
    try {
        $dmarcTxt = Resolve-DnsName -Name "_dmarc.$d" -Type TXT -ErrorAction Stop
        $dmarcRecord = ($dmarcTxt | Where-Object { $_.Strings -match '^v=DMARC1' } | Select-Object -First 1).Strings -join ''
    } catch {}

    if (-not $dmarcRecord) {
        Write-Host "    [FAIL] DMARC record not found" -ForegroundColor Red
        $results.Add([PSCustomObject]@{ Domain = $d; RecordType = 'DMARC'; Found = $false; Value = ''; Flag = 'Fail'; Detail = 'No DMARC TXT record found at _dmarc.' + $d })
    } else {
        # Anchor on the tag boundary, so 'sp=' (subdomain policy) is not read as 'p='.
        $policy = if ($dmarcRecord -match '(?:^|;)\s*p\s*=\s*(\w+)') { $Matches[1] } else { 'unknown' }
        $hasRua = $dmarcRecord -match 'rua='
        $flag = if ($policy -eq 'none') { 'Warn' } elseif (-not $hasRua) { 'Warn' } else { 'Pass' }
        $detail = if ($policy -eq 'none') { 'Policy is "none" — monitoring only, not enforced' }
                  elseif (-not $hasRua) { 'No aggregate report address (rua=) configured' }
                  else { "Policy: $policy" }
        Write-Host "    [$flag] DMARC: $dmarcRecord" -ForegroundColor $(if ($flag -eq 'Pass') { 'Green' } else { 'Yellow' })
        $results.Add([PSCustomObject]@{ Domain = $d; RecordType = 'DMARC'; Found = $true; Value = $dmarcRecord; Flag = $flag; Detail = $detail })
    }
    Write-Host ""
}

# ── Output ────────────────────────────────────────────────────────────────────
if (-not $OutputPath) {
    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $OutputPath = Join-Path $outputDir "EmailAuthenticationRecords_$ts.csv"
}
$results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host "  Report saved: $OutputPath" -ForegroundColor Green

$failCount = @($results | Where-Object { $_.Flag -eq 'Fail' }).Count
Write-Host ""
Write-Host ("  {0} domain(s) checked — {1} record(s) failed" -f $Domain.Count, $failCount) -ForegroundColor $(if ($failCount -gt 0) { 'Red' } else { 'Green' })
Write-Host ""

Disconnect-M365Graph $graph
