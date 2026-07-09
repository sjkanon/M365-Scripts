#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Diagnose authentication and network issues on a Windows server or endpoint.

.DESCRIPTION
    Scans Windows Event Viewer and optional log files for authentication failures,
    network errors, and Kerberos/NTLM issues. Also runs active connectivity checks.

    Checks performed:
      - Event Viewer: logon failures (4625), Kerberos errors (4771/4768), NTLM failures (4776)
      - Event Viewer: SMB/network errors, Netlogon errors, DNS client errors
      - Windows Time sync (Kerberos requires < 5 min offset)
      - Kerberos ticket cache (klist)
      - DNS resolution for specified hosts
      - TCP port connectivity for specified endpoints
      - UNC share access for specified paths
      - Network adapter status
      - Log file scan for auth/network error patterns (optional)

    Output: txt report to C:\Temp\ or a custom path.

.PARAMETER TestHosts
    Hostnames or IPs to test DNS resolution and TCP connectivity against.

.PARAMETER TestPorts
    TCP ports to test on each host in -TestHosts (default: 445, 88, 389, 636).

.PARAMETER TestShares
    UNC paths to test read access (e.g. \\server\share).

.PARAMETER LogDirectory
    Optional directory to scan for auth/network error patterns in log files.

.PARAMETER LogDaysBack
    How many days back to include log files (default: 1).

.PARAMETER EventLogHours
    How many hours back to check Windows Event Viewer (default: 24).

.PARAMETER OutputPath
    Override the default output folder (C:\Temp\).

.EXAMPLE
    # Basic run — Event Viewer only
    .\Test-AuthNetworkDiagnostics.ps1

.EXAMPLE
    # Test connectivity + UNC shares
    .\Test-AuthNetworkDiagnostics.ps1 -TestHosts "dc01","fileserver" -TestShares "\\fileserver\data"

.EXAMPLE
    # Full scan including SAS log files
    .\Test-AuthNetworkDiagnostics.ps1 -TestHosts "sasserver" -LogDirectory "E:\SAS\Logs"

.EXAMPLE
    # Check last 48 hours of events
    .\Test-AuthNetworkDiagnostics.ps1 -EventLogHours 48
#>
[CmdletBinding()]
param (
    [string[]] $TestHosts     = @(),
    [int[]]    $TestPorts     = @(445, 88, 389, 636),
    [string[]] $TestShares    = @(),
    [string]   $LogDirectory  = '',
    [int]      $LogDaysBack   = 1,
    [int]      $EventLogHours = 24,
    [string]   $OutputPath    = ''
)

# ── Output ────────────────────────────────────────────────────────────────────
$outDir = if ($OutputPath) { $OutputPath } else { 'C:\Temp' }
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }
$reportFile = Join-Path $outDir "AuthNetworkDiag-$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"

$report  = [System.Collections.Generic.List[string]]::new()
$issues  = [System.Collections.Generic.List[string]]::new()

function Add-Line  { param([string]$l) $report.Add($l) }
function Add-Issue { param([string]$severity, [string]$msg)
    $line = "[$severity] $msg"
    $report.Add("  $line")
    $issues.Add($line)
}
function Add-Ok    { param([string]$msg) $report.Add("  [OK]  $msg") }
function Add-Info  { param([string]$msg) $report.Add("  [INFO] $msg") }
function Section   { param([string]$title)
    $report.Add('')
    $report.Add(('─' * 60))
    $report.Add("  $title")
    $report.Add(('─' * 60))
}

$cutoff = (Get-Date).AddHours(-$EventLogHours)

# ── Header ────────────────────────────────────────────────────────────────────
Add-Line ('=' * 60)
Add-Line "  Auth & Network Diagnostics"
Add-Line "  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  |  $env:COMPUTERNAME"
Add-Line ('=' * 60)
Add-Info "Event log lookback : last $EventLogHours hours (since $($cutoff.ToString('yyyy-MM-dd HH:mm')))"
if ($LogDirectory) {
    Add-Info "Log file scan      : $LogDirectory (last $LogDaysBack day(s))"
}

# ── 1. Windows Time Sync ──────────────────────────────────────────────────────
Section "1. Windows Time Sync (Kerberos requires < 5 min offset)"

try {
    $svc = Get-Service -Name W32Time -ErrorAction Stop
    if ($svc.Status -ne 'Running') {
        Add-Issue 'WARN' "Windows Time service is not running (status: $($svc.Status))"
    } else {
        Add-Ok "Windows Time service is running"
    }
} catch {
    Add-Issue 'WARN' "Could not check Windows Time service: $_"
}

try {
    $w32tmOutput = & w32tm /query /status 2>&1 | Out-String
    if ($w32tmOutput -match 'Last Successful Sync Time:\s*(.+)') {
        Add-Info "Last sync : $($Matches[1].Trim())"
    }
    if ($w32tmOutput -match 'Source:\s*(.+)') {
        Add-Info "Source    : $($Matches[1].Trim())"
    }
    if ($w32tmOutput -match 'Stratum:\s*(\d+)') {
        $stratum = [int]$Matches[1]
        if ($stratum -ge 16) {
            Add-Issue 'ERROR' "Time stratum is $stratum — server is not synchronised"
        }
    }
} catch {
    Add-Issue 'WARN' "Could not run w32tm: $_"
}

# ── 2. Kerberos Ticket Cache ───────────────────────────────────────────────────
Section "2. Kerberos Ticket Cache (klist)"

try {
    $klist = & klist 2>&1 | Out-String
    if ($klist -match 'Cached Tickets: \(0\)') {
        Add-Issue 'WARN' "No Kerberos tickets cached — possible auth issue"
    } elseif ($klist -match 'Credentials cache: .+ has no tickets') {
        Add-Issue 'WARN' "Kerberos credential cache is empty"
    } else {
        $ticketCount = ([regex]::Matches($klist, 'Client:')).Count
        Add-Ok "$ticketCount Kerberos ticket(s) cached"

        # Check for expired tickets
        $expiredMatches = [regex]::Matches($klist, 'End Time:\s+(.+)')
        $expiredCount = 0
        foreach ($m in $expiredMatches) {
            $endTimeStr = $m.Groups[1].Value.Trim() -replace '\s+\(.*\)', ''
            try {
                if ([datetime]::Parse($endTimeStr) -lt (Get-Date)) { $expiredCount++ }
            } catch {}
        }
        if ($expiredCount -gt 0) {
            Add-Issue 'WARN' "$expiredCount expired Kerberos ticket(s) found — run 'klist purge'"
        }
    }
} catch {
    Add-Issue 'INFO' "klist not available or no tickets: $_"
}

# ── 3. Network Adapters ────────────────────────────────────────────────────────
Section "3. Network Adapters"

try {
    $adapters = Get-NetAdapter -ErrorAction Stop | Where-Object { $_.Status -ne 'Not Present' }
    foreach ($adapter in $adapters) {
        if ($adapter.Status -eq 'Up') {
            Add-Ok "$($adapter.Name) — $($adapter.Status) ($($adapter.LinkSpeed))"
        } elseif ($adapter.Status -eq 'Disconnected') {
            Add-Issue 'WARN' "Adapter '$($adapter.Name)' is disconnected"
        } else {
            Add-Issue 'ERROR' "Adapter '$($adapter.Name)' status: $($adapter.Status)"
        }
    }
} catch {
    Add-Issue 'WARN' "Could not enumerate network adapters: $_"
}

# ── 4. Event Viewer — Authentication ──────────────────────────────────────────
Section "4. Event Viewer — Authentication Failures (last $EventLogHours h)"

$authEvents = @(
    @{ Id = 4625; Log = 'Security';  Sev = 'ERROR'; Desc = 'Account logon failure' }
    @{ Id = 4771; Log = 'Security';  Sev = 'ERROR'; Desc = 'Kerberos pre-auth failure' }
    @{ Id = 4768; Log = 'Security';  Sev = 'WARN';  Desc = 'Kerberos TGT request failure' }
    @{ Id = 4776; Log = 'Security';  Sev = 'ERROR'; Desc = 'NTLM credential validation failure' }
    @{ Id = 4740; Log = 'Security';  Sev = 'ERROR'; Desc = 'Account locked out' }
)

foreach ($ev in $authEvents) {
    try {
        $events = Get-WinEvent -FilterHashtable @{
            LogName   = $ev.Log
            Id        = $ev.Id
            StartTime = $cutoff
        } -ErrorAction SilentlyContinue

        if ($events -and $events.Count -gt 0) {
            Add-Issue $ev.Sev "$($events.Count)x Event $($ev.Id) — $($ev.Desc)"
            $events | Select-Object -First 3 | ForEach-Object {
                $report.Add("         $($_.TimeCreated.ToString('HH:mm:ss')) — $($_.Message -replace '\s+', ' ' | Select-Object -First 1)")
            }
        } else {
            Add-Ok "No Event $($ev.Id) ($($ev.Desc))"
        }
    } catch {
        Add-Info "Could not query Event $($ev.Id): $_"
    }
}

# ── 5. Event Viewer — Network & SMB ────────────────────────────────────────────
Section "5. Event Viewer — Network & SMB Errors (last $EventLogHours h)"

$netEvents = @(
    @{ Id = 10016; Log = 'System'; Sev = 'WARN';  Desc = 'DCOM permission error' }
    @{ Id = 5719;  Log = 'System'; Sev = 'ERROR'; Desc = 'No DC available (NETLOGON)' }
    @{ Id = 1129;  Log = 'System'; Sev = 'ERROR'; Desc = 'Group Policy processing failed (no network)' }
    @{ Id = 1014;  Log = 'System'; Sev = 'WARN';  Desc = 'DNS name resolution timeout' }
    @{ Id = 1085;  Log = 'Microsoft-Windows-DNS-Client/Operational'; Sev = 'WARN'; Desc = 'DNS client resolver failure' }
)

foreach ($ev in $netEvents) {
    try {
        $filter = @{ LogName = $ev.Log; Id = $ev.Id; StartTime = $cutoff }
        $events = Get-WinEvent -FilterHashtable $filter -ErrorAction SilentlyContinue
        if ($events -and $events.Count -gt 0) {
            Add-Issue $ev.Sev "$($events.Count)x Event $($ev.Id) — $($ev.Desc)"
        } else {
            Add-Ok "No Event $($ev.Id) ($($ev.Desc))"
        }
    } catch {
        Add-Info "Could not query Event $($ev.Id)"
    }
}

# Check Netlogon log for errors
try {
    $netlogonLog = "$env:SystemRoot\debug\netlogon.log"
    if (Test-Path $netlogonLog) {
        $netlogonErrors = Get-Content $netlogonLog -ErrorAction Stop |
            Where-Object { $_ -match '\[ERROR\]|\bNO_LOGON_SERVERS\b|\bNO_SUCH_DOMAIN\b|\bTIMEOUT\b' } |
            Select-Object -Last 10
        if ($netlogonErrors.Count -gt 0) {
            Add-Issue 'ERROR' "$($netlogonErrors.Count) recent error(s) in Netlogon.log"
            $netlogonErrors | Select-Object -First 3 | ForEach-Object { $report.Add("         $_") }
        } else {
            Add-Ok "Netlogon.log: no recent errors"
        }
    }
} catch {}

# ── 6. DNS Resolution ─────────────────────────────────────────────────────────
if ($TestHosts.Count -gt 0) {
    Section "6. DNS Resolution"

    foreach ($hostName in $TestHosts) {
        try {
            $resolved = [System.Net.Dns]::GetHostAddresses($hostName)
            Add-Ok "$hostName → $($resolved.IPAddressToString -join ', ')"
        } catch {
            Add-Issue 'ERROR' "DNS resolution failed for '$hostName': $_"
        }
    }
}

# ── 7. TCP Connectivity ───────────────────────────────────────────────────────
if ($TestHosts.Count -gt 0) {
    Section "7. TCP Connectivity"

    $portNames = @{ 88 = 'Kerberos'; 389 = 'LDAP'; 445 = 'SMB'; 636 = 'LDAPS'; 3268 = 'GC'; 80 = 'HTTP'; 443 = 'HTTPS' }

    foreach ($h in $TestHosts) {
        foreach ($port in $TestPorts) {
            $label = if ($portNames.ContainsKey($port)) { "$port/$($portNames[$port])" } else { "$port" }
            try {
                $tcp = [System.Net.Sockets.TcpClient]::new()
                $ar  = $tcp.BeginConnect($h, $port, $null, $null)
                $ok  = $ar.AsyncWaitHandle.WaitOne(3000)
                if ($ok -and $tcp.Connected) {
                    Add-Ok "${h}:${label}"
                } else {
                    Add-Issue 'ERROR' "${h}:${label} — connection refused or timed out"
                }
                $tcp.Close()
            } catch {
                Add-Issue 'ERROR' "${h}:${label} — $($_.Exception.Message)"
            }
        }
    }
}

# ── 8. UNC Share Access ────────────────────────────────────────────────────────
if ($TestShares.Count -gt 0) {
    Section "8. UNC Share Access"

    foreach ($share in $TestShares) {
        try {
            $null = Get-Item -Path $share -ErrorAction Stop
            Add-Ok "Accessible: $share"
        } catch {
            Add-Issue 'ERROR' "Cannot access '$share': $($_.Exception.Message)"
        }
    }
}

# ── 9. Log File Scan ──────────────────────────────────────────────────────────
if ($LogDirectory -and (Test-Path $LogDirectory)) {
    Section "9. Log File Scan — Auth & Network Errors"

    $logPatterns = @(
        @{ Pattern = 'access.?denied|unauthorized|permission.?denied';                 Severity = 'ERROR'; Desc = 'Access denied / unauthorized' }
        @{ Pattern = 'authentication.?fail|login.?fail|logon.?fail';                   Severity = 'ERROR'; Desc = 'Authentication / login failure' }
        @{ Pattern = 'cannot.?connect|connection.?refused|connection.?timed.?out';     Severity = 'ERROR'; Desc = 'Connection failure' }
        @{ Pattern = 'network.?path.?not.?found|no.?such.?host|name.?resolution';     Severity = 'ERROR'; Desc = 'Network / DNS error' }
        @{ Pattern = 'certificate.?expired|ssl.?error|tls.?error|cert.?invalid';       Severity = 'ERROR'; Desc = 'Certificate / TLS error' }
        @{ Pattern = 'kerberos|ntlm|ldap.?error|domain.?controller';                  Severity = 'WARN';  Desc = 'Kerberos / LDAP / DC reference' }
        @{ Pattern = 'disk.?full|no.?space.?left|storage.?error|i\/o.?error';         Severity = 'WARN';  Desc = 'Disk / I/O error' }
    )

    $cutoffDate  = (Get-Date).AddDays(-$LogDaysBack)
    $logFiles    = Get-ChildItem -Path $LogDirectory -Recurse -File -ErrorAction SilentlyContinue |
                   Where-Object { $_.LastWriteTime -gt $cutoffDate -and $_.Extension -in @('.log','.txt') }

    Add-Info "Scanning $($logFiles.Count) log file(s) modified in the last $LogDaysBack day(s)"

    $matchSummary = @{}
    foreach ($logFile in $logFiles) {
        try {
            $lines = Get-Content $logFile.FullName -ErrorAction Stop
            foreach ($pat in $logPatterns) {
                $hits = $lines | Where-Object { $_ -imatch $pat.Pattern }
                if ($hits.Count -gt 0) {
                    $key = $pat.Desc
                    if (-not $matchSummary.ContainsKey($key)) {
                        $matchSummary[$key] = @{ Severity = $pat.Severity; Count = 0; Examples = @() }
                    }
                    $matchSummary[$key].Count += $hits.Count
                    if ($matchSummary[$key].Examples.Count -lt 3) {
                        $matchSummary[$key].Examples += "[$($logFile.Name)] $($hits[0].Trim())"
                    }
                }
            }
        } catch {}
    }

    if ($matchSummary.Count -eq 0) {
        Add-Ok "No auth/network patterns found in log files"
    } else {
        foreach ($key in $matchSummary.Keys) {
            $entry = $matchSummary[$key]
            Add-Issue $entry.Severity "$($entry.Count)x — $key"
            $entry.Examples | ForEach-Object { $report.Add("         $_") }
        }
    }
} elseif ($LogDirectory) {
    Add-Issue 'WARN' "Log directory not found: $LogDirectory"
}

# ── Summary ───────────────────────────────────────────────────────────────────
$report.Add('')
$report.Add('=' * 60)
$report.Add('  SUMMARY')
$report.Add('=' * 60)

if ($issues.Count -eq 0) {
    $report.Add('  No issues found.')
} else {
    $errors   = @($issues | Where-Object { $_ -match '^\[ERROR\]' })
    $warnings = @($issues | Where-Object { $_ -match '^\[WARN\]' })
    $report.Add("  Errors   : $($errors.Count)")
    $report.Add("  Warnings : $($warnings.Count)")
    $report.Add('')
    $issues | ForEach-Object { $report.Add("  $_") }
}

$report.Add('')

# ── Write report ──────────────────────────────────────────────────────────────
$report | Out-File -FilePath $reportFile -Encoding UTF8

# Echo to console
$report | ForEach-Object {
    if ($_ -match '^\s*\[ERROR\]') { Write-Host $_ -ForegroundColor Red }
    elseif ($_ -match '^\s*\[WARN\]') { Write-Host $_ -ForegroundColor Yellow }
    elseif ($_ -match '^\s*\[OK\]')   { Write-Host $_ -ForegroundColor Green }
    else { Write-Host $_ }
}

Write-Host ""
Write-Host "  Report saved: $reportFile" -ForegroundColor Cyan
Write-Host ""
