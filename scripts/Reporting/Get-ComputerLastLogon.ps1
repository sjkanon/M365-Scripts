#Requires -Version 5.1
<#
.SYNOPSIS
    Report last logon date for all computer objects in one or more OUs.

.DESCRIPTION
    Queries Active Directory for computer objects within the specified OUs and
    reports the last logon date, OS, enabled status, and staleness per device.

    Two accuracy modes:
      - Default  : Uses LastLogonTimestamp (replicated, up to 14 days behind).
                   Fast and sufficient for stale-device reporting.
      - -AllDCs  : Queries every domain controller for the raw LastLogon
                   attribute and picks the most recent value per computer.
                   Accurate but slower in large environments.

    Output:
      - Timestamped CSV exported to -ExportPath (default C:\Temp\)
      - Console summary with counts per status

.PARAMETER SearchBase
    One or more OU distinguished names to search.
    Accepts a string array. Searches all child OUs recursively.
    If omitted the script asks interactively or falls back to the whole domain.

.PARAMETER AllDCs
    Query all domain controllers for the most accurate LastLogon value.
    Slower, but eliminates the 9-14 day replication delay.

.PARAMETER InactiveDays
    Number of days without logon before a computer is marked Stale.
    Default: 90.

.PARAMETER IncludeDisabled
    Include disabled computer objects in the report.
    By default only enabled computers are included.

.PARAMETER ExportPath
    Folder where the CSV file is saved. Default: C:\Temp\.

.EXAMPLE
    # Vias Institute — Laptops OU
    .\Get-ComputerLastLogon.ps1 `
        -SearchBase "OU=Laptops,OU=Computers,OU=Vias Institute,DC=ad,DC=vias,DC=be"

.EXAMPLE
    # Vias Institute — alle computers (Laptops + rest van de Computers OU)
    .\Get-ComputerLastLogon.ps1 `
        -SearchBase "OU=Computers,OU=Vias Institute,DC=ad,DC=vias,DC=be"

.EXAMPLE
    # Nauwkeurigste modus — bevraagt alle DC's
    .\Get-ComputerLastLogon.ps1 `
        -SearchBase "OU=Laptops,OU=Computers,OU=Vias Institute,DC=ad,DC=vias,DC=be" `
        -AllDCs

.EXAMPLE
    # Inclusief uitgeschakelde computers, drempel op 60 dagen
    .\Get-ComputerLastLogon.ps1 `
        -SearchBase "OU=Computers,OU=Vias Institute,DC=ad,DC=vias,DC=be" `
        -IncludeDisabled -InactiveDays 60
#>
[CmdletBinding()]
param (
    [string[]] $SearchBase,
    [switch]   $AllDCs,
    [int]      $InactiveDays   = 90,
    [switch]   $IncludeDisabled,
    [string]   $ExportPath     = 'C:\Temp'
)

#Requires -Modules ActiveDirectory
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Helpers ────────────────────────────────────────────────────────────────────
function Write-Status {
    param([string]$Msg, [string]$Level = 'INFO')
    $color = switch ($Level) {
        'OK'   { 'Green'  } 'WARN' { 'Yellow' }
        'FAIL' { 'Red'    } 'HEAD' { 'Cyan'   }
        default{ 'Gray'   }
    }
    Write-Host "  [$Level] $Msg" -ForegroundColor $color
}

function Get-OUFromDN {
    param([string]$DN)
    ($DN -split ',' | Where-Object { $_ -like 'OU=*' } | ForEach-Object { $_ -replace '^OU=' }) -join ' > '
}

# ── Output folder ──────────────────────────────────────────────────────────────
if (-not (Test-Path $ExportPath)) { New-Item -ItemType Directory -Path $ExportPath | Out-Null }
$ts      = Get-Date -Format 'yyyyMMdd_HHmmss'
$csvFile = Join-Path $ExportPath "ComputerLastLogon_$ts.csv"

# ── Banner ─────────────────────────────────────────────────────────────────────
Write-Host ''
Write-Host '  ════════════════════════════════════════════' -ForegroundColor Cyan
Write-Host '   Get-ComputerLastLogon' -ForegroundColor Cyan
Write-Host '  ════════════════════════════════════════════' -ForegroundColor Cyan
Write-Host ''

# ── Resolve SearchBase ─────────────────────────────────────────────────────────
if (-not $SearchBase) {
    try {
        $domain = (Get-ADDomain).DistinguishedName
        Write-Status "No -SearchBase specified — searching entire domain: $domain" 'WARN'
        $SearchBase = @($domain)
    } catch {
        throw "Cannot resolve domain. Provide -SearchBase explicitly."
    }
}

# ── Collect computers from all specified OUs ───────────────────────────────────
$adFilter = if ($IncludeDisabled) { '*' } else { "Enabled -eq 'True'" }

$adProps = @(
    'Name', 'Enabled', 'OperatingSystem', 'OperatingSystemVersion',
    'lastLogonTimestamp', 'PasswordLastSet', 'whenCreated', 'DistinguishedName', 'Description', 'IPv4Address'
)

$allComputers = [System.Collections.Generic.List[object]]::new()

foreach ($ou in $SearchBase) {
    Write-Status "Querying OU: $ou" 'INFO'
    try {
        $computers = Get-ADComputer -Filter $adFilter -SearchBase $ou `
            -Properties $adProps -ResultPageSize 500 -ErrorAction Stop
        $allComputers.AddRange([object[]]$computers)
        Write-Status "  Found $($computers.Count) computer(s)" 'OK'
    } catch {
        Write-Status "Failed to query '$ou': $_" 'FAIL'
    }
}

if ($allComputers.Count -eq 0) {
    Write-Status 'No computers found in the specified OUs.' 'WARN'
    exit 0
}

Write-Status "Total computers to process: $($allComputers.Count)" 'HEAD'
Write-Host ''

# ── All-DC mode: collect LastLogon from every DC ──────────────────────────────
$lastLogonPerComputer = @{}

if ($AllDCs) {
    Write-Status 'Querying all domain controllers for accurate LastLogon...' 'HEAD'

    $dcs = Get-ADDomainController -Filter * | Select-Object -ExpandProperty HostName
    Write-Status "Domain controllers found: $($dcs.Count)" 'INFO'

    $processed = 0
    foreach ($computer in $allComputers) {
        $processed++
        if ($processed % 25 -eq 0) {
            Write-Status "  Processed $processed / $($allComputers.Count)..." 'INFO'
        }
        $best = [DateTime]::MinValue
        foreach ($dc in $dcs) {
            try {
                $obj = Get-ADComputer $computer.SamAccountName -Server $dc `
                    -Properties LastLogon -ErrorAction Stop
                if ($obj.LastLogon -gt 0) {
                    $dt = [DateTime]::FromFileTime($obj.LastLogon)
                    if ($dt -gt $best) { $best = $dt }
                }
            } catch {}
        }
        $lastLogonPerComputer[$computer.SamAccountName] = if ($best -gt [DateTime]::MinValue) { $best } else { $null }
    }
    Write-Status 'All-DC query complete.' 'OK'
    Write-Host ''
}

# ── Build report rows ──────────────────────────────────────────────────────────
$now    = Get-Date
$report = foreach ($c in $allComputers) {

    $lastLogon = if ($AllDCs) {
        $lastLogonPerComputer[$c.SamAccountName]
    } elseif ($c.lastLogonTimestamp -gt 0) {
        [DateTime]::FromFileTime($c.lastLogonTimestamp)
    } else {
        $null
    }

    $daysSince    = if ($lastLogon) { [math]::Round(($now - $lastLogon).TotalDays) } else { $null }
    $pwdLastSet   = $c.PasswordLastSet
    $daysSincePwd = if ($pwdLastSet) { [math]::Round(($now - $pwdLastSet).TotalDays) } else { $null }

    # PasswordLastSet: computer accounts auto-rotate every ~30 days when online.
    # If LastLogonTimestamp looks stale but password was set recently, the device
    # is still active — the replication delay is misleading.
    $pwdRecent = $pwdLastSet -and $daysSincePwd -le 35

    $status = if (-not $c.Enabled) {
        'Disabled'
    } elseif (-not $lastLogon -and -not $pwdRecent) {
        'Never'
    } elseif ($daysSince -le $InactiveDays -or ($null -eq $daysSince -and $pwdRecent)) {
        'Active'
    } elseif ($pwdRecent) {
        'Active (pwd recent)'   # LastLogon stale, maar wachtwoord recent vernieuwd → device is online
    } else {
        'Stale'
    }

    [PSCustomObject]@{
        Name                   = $c.Name
        Status                 = $status
        Enabled                = $c.Enabled
        LastLogon              = if ($lastLogon) { $lastLogon.ToString('dd/MM/yyyy HH:mm') } else { 'Never' }
        DaysSinceLogon         = if ($null -ne $daysSince) { $daysSince } else { '' }
        PasswordLastSet        = if ($pwdLastSet) { $pwdLastSet.ToString('dd/MM/yyyy') } else { '' }
        DaysSincePasswordSet   = if ($null -ne $daysSincePwd) { $daysSincePwd } else { '' }
        OperatingSystem        = $c.OperatingSystem
        OperatingSystemVersion = $c.OperatingSystemVersion
        IPv4Address            = $c.IPv4Address
        OU                     = Get-OUFromDN $c.DistinguishedName
        Created                = $c.whenCreated.ToString('dd/MM/yyyy')
        Description            = $c.Description
        DistinguishedName      = $c.DistinguishedName
    }
}

# ── Export CSV ─────────────────────────────────────────────────────────────────
$report | Sort-Object DaysSinceLogon -Descending |
    Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8

# ── Console summary ────────────────────────────────────────────────────────────
Write-Host ''
Write-Host '  ── Summary ────────────────────────────────────' -ForegroundColor Cyan

$active   = @($report | Where-Object { $_.Status -eq 'Active'   }).Count
$stale    = @($report | Where-Object { $_.Status -eq 'Stale'    }).Count
$never    = @($report | Where-Object { $_.Status -eq 'Never'    }).Count
$disabled = @($report | Where-Object { $_.Status -eq 'Disabled' }).Count

Write-Status "Total computers  : $($report.Count)" 'INFO'
Write-Status "Active           : $active  (logged on within $InactiveDays days)" 'OK'
Write-Status "Stale            : $stale  (no logon for > $InactiveDays days)" 'WARN'
Write-Status "Never logged on  : $never" 'WARN'
if ($IncludeDisabled) {
    Write-Status "Disabled         : $disabled" 'INFO'
}

$accuracyNote = if ($AllDCs) { 'All DCs (accurate)' } else { 'LastLogonTimestamp (up to 14 days behind)' }
Write-Status "Accuracy mode    : $accuracyNote" 'INFO'
Write-Status "CSV exported to  : $csvFile" 'OK'
Write-Host ''
