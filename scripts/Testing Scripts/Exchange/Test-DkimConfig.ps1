#Requires -Version 5.1
<#
.SYNOPSIS
    Validate DKIM signing configuration and DNS records for Exchange Online domains.

.DESCRIPTION
    Connects to Exchange Online and checks the DKIM signing configuration for one
    or all accepted domains. For each domain it verifies:
      - Whether DKIM signing is enabled
      - Whether selector CNAME records exist in DNS and point to the correct targets
      - Whether Microsoft's TXT public key records are resolvable and match the config
    Outputs a list of required actions where issues are found.

    Note: Resolve-DnsName is Windows-only. On macOS/Linux DNS lookups are skipped
    and only the Exchange config is shown.

.PARAMETER Domain
    Domain name to check. If omitted, all domains with a DKIM signing config are checked.

.PARAMETER ShowAll
    Show full DKIM signing config object instead of the summarised view.

.PARAMETER TenantId
    Entra ID tenant ID or domain. Optional if already connected.

.EXAMPLE
    .\Test-DkimConfig.ps1

.EXAMPLE
    .\Test-DkimConfig.ps1 -Domain "contoso.com"

.EXAMPLE
    .\Test-DkimConfig.ps1 -Domain "contoso.com" -ShowAll
#>
[CmdletBinding()]
param(
    [string] $Domain,
    [switch] $ShowAll,
    [string] $TenantId
)

$isWindows = $PSVersionTable.PSEdition -eq 'Desktop' -or $IsWindows

# ── Connection ────────────────────────────────────────────────────────────────
$script:ConnectedHere = $false
try {
    $null = Get-DkimSigningConfig -ErrorAction Stop | Select-Object -First 1
} catch {
    $connectParams = @{ ShowBanner = $false }
    if ($TenantId) { $connectParams['Organization'] = $TenantId }
    Connect-ExchangeOnline @connectParams
    $script:ConnectedHere = $true
}

# ── Helper: compare public key in DNS TXT to Exchange config ──────────────────
function Compare-DkimKeys {
    param([string] $DnsKey, [string] $ConfigKey)
    if ([string]::IsNullOrWhiteSpace($DnsKey) -or [string]::IsNullOrWhiteSpace($ConfigKey)) { return $false }
    $regex = 'p=(.*?);'
    $dnsVal    = if ($DnsKey    -match $regex) { $Matches[1].Trim() } else { $null }
    $configVal = if ($ConfigKey -match $regex) { $Matches[1].Trim() } else { $null }
    return ($dnsVal -and $configVal -and $dnsVal -eq $configVal)
}

# ── Validate one domain config ────────────────────────────────────────────────
function Test-DomainDkim {
    param($Config)

    $domain = $Config.Domain
    $isOnmicrosoft = $domain -match '(onmicrosoft|microsoftonline)\.com$'
    $actions = @()

    Write-Host ""
    Write-Host "  ================================================" -ForegroundColor Cyan
    Write-Host "   $domain" -ForegroundColor Cyan
    Write-Host "  ================================================" -ForegroundColor Cyan
    Write-Host ""

    if ($ShowAll) {
        $Config | Format-List
    } else {
        $Config | Select-Object Identity, Enabled, Status, Selector1CNAME, Selector2CNAME,
            KeyCreationTime, LastChecked, RotateOnDate | Format-List
    }

    if (-not $Config.Enabled) {
        Write-Host "  [WARN] DKIM signing is NOT enabled for $domain" -ForegroundColor Yellow
        $actions += "Enable DKIM signing: Set-DkimSigningConfig -Identity '$domain' -Enabled `$true"
    }

    if (-not $isOnmicrosoft) {
        if (-not $isWindows) {
            Write-Host "  [INFO] DNS lookup skipped — Resolve-DnsName is Windows-only." -ForegroundColor DarkGray
        } else {
            Write-Host "  Checking DNS..." -ForegroundColor DarkGray
            Write-Host ""

            foreach ($n in 1, 2) {
                $selectorProp = "Selector${n}CNAME"
                $keyProp      = "Selector${n}PublicKey"
                $cnameHost    = "selector${n}._domainkey.$domain"
                $expected     = $Config.$selectorProp
                $expectedKey  = $Config.$keyProp

                Write-Host "  Selector $n" -ForegroundColor DarkGray
                Write-Host ("  {0,-16} : {1}" -f "Expected CNAME", $expected)

                $cnameDns = Resolve-DnsName -Name $cnameHost -Type CNAME -ErrorAction SilentlyContinue
                if ($cnameDns -and $cnameDns.NameHost) {
                    Write-Host ("  {0,-16} : {1}" -f "DNS CNAME", $cnameDns.NameHost)
                    $match = $expected.Trim() -eq $cnameDns.NameHost.Trim()
                    if ($match) {
                        Write-Host "  CNAME match      : OK" -ForegroundColor Green
                    } else {
                        Write-Host "  CNAME match      : MISMATCH" -ForegroundColor Red
                        $actions += "Selector $n — update CNAME $cnameHost → $expected"
                    }
                } else {
                    Write-Host "  CNAME             : NOT FOUND ($cnameHost)" -ForegroundColor Red
                    $actions += "Selector $n — create CNAME $cnameHost → $expected"
                }

                $txtDns = Resolve-DnsName -Name $expected -Type TXT -ErrorAction SilentlyContinue
                if ($txtDns -and $txtDns.Strings) {
                    $dnsKey = $txtDns.Strings[0].Trim()
                    $keyMatch = Compare-DkimKeys $dnsKey $expectedKey
                    if ($keyMatch) {
                        Write-Host "  Public key match : OK" -ForegroundColor Green
                    } else {
                        Write-Host "  Public key match : MISMATCH" -ForegroundColor Red
                        $actions += "Selector $n — TXT key at $expected does not match Exchange config"
                    }
                } else {
                    Write-Host "  Public key TXT   : NOT FOUND ($expected)" -ForegroundColor Red
                    $actions += "Selector $n — Microsoft TXT record $expected missing; DKIM config may need to be recreated"
                }

                Write-Host ""
            }
        }
    } else {
        Write-Host "  [INFO] Skipping DNS check for onmicrosoft.com domain." -ForegroundColor DarkGray
        Write-Host ""
    }

    if ($actions.Count -gt 0) {
        Write-Host "  Required actions:" -ForegroundColor Yellow
        foreach ($a in $actions) { Write-Host "    - $a" }
        Write-Host ""
    } else {
        Write-Host "  No issues found." -ForegroundColor Green
        Write-Host ""
    }
}

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   DKIM Configuration Validator" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan

# ── Run ───────────────────────────────────────────────────────────────────────
if ($Domain) {
    $config = Get-DkimSigningConfig -Identity $Domain -ErrorAction SilentlyContinue
    if ($config) {
        Test-DomainDkim $config
    } else {
        Write-Host ""
        Write-Host "  [WARN] No DKIM signing config found for '$Domain'." -ForegroundColor Yellow

        if ($Domain -notmatch '(onmicrosoft|microsoftonline)\.com$') {
            # Try DNS-only check (no EXO config yet)
            if ($isWindows) {
                Write-Host "  Checking DNS CNAME records only..." -ForegroundColor DarkGray
                Write-Host ""
                foreach ($n in 1, 2) {
                    $cnameHost = "selector${n}._domainkey.$Domain"
                    $cnameDns  = Resolve-DnsName -Name $cnameHost -Type CNAME -ErrorAction SilentlyContinue
                    if ($cnameDns) {
                        Write-Host ("  selector{0}._domainkey : {1}" -f $n, $cnameDns.NameHost) -ForegroundColor Green
                    } else {
                        Write-Host "  selector${n}._domainkey : NOT FOUND ($cnameHost)" -ForegroundColor Red
                    }
                }
                Write-Host ""
            }
            Write-Host "  To create a signing config: New-DkimSigningConfig -DomainName '$Domain' -Enabled `$true" -ForegroundColor DarkGray
        }
        Write-Host ""
    }
} else {
    $configs = Get-DkimSigningConfig -ErrorAction SilentlyContinue
    if ($configs) {
        foreach ($cfg in $configs) { Test-DomainDkim $cfg }
    } else {
        Write-Host ""
        Write-Host "  No DKIM signing configurations found in this tenant." -ForegroundColor Yellow
        Write-Host ""
    }
}

# ── Disconnect if we connected ────────────────────────────────────────────────
if ($script:ConnectedHere) { Disconnect-ExchangeOnline -Confirm:$false | Out-Null }
