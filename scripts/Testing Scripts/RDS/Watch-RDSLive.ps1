#Requires -Version 5.1
<#
.SYNOPSIS
    Real-time RDS session and licensing monitor.

.DESCRIPTION
    Polls Windows event logs every N seconds and shows new events as they occur.
    Run this script directly on each RDS/RDWeb server for full visibility.

    Events monitored:
      - TerminalServices-LocalSessionManager: logon (21), reconnect (22/25),
        logoff (23), disconnect (24), logon failed (20), disconnect reason (40)
      - Security: failed RDP logon (4625 type 10), account lockout (4740)
      - TerminalServices-Licensing: license granted/denied/warning events
      - System: TermServLicensing provider (grace period, license server errors)

    Designed for use after licensing changes — shows whether licenses are being
    issued correctly and whether sessions succeed or fail.

.PARAMETER IntervalSeconds
    Polling interval in seconds. Default: 20.

.PARAMETER LogPath
    Folder for the output log file. Default: C:\Temp\.

.PARAMETER NoLogFile
    Output to console only, do not write a log file.

.EXAMPLE
    # Run on TRTS002 or TR-AZ-TSE-01
    .\Watch-RDSLive.ps1

.EXAMPLE
    # Faster polling, no log file
    .\Watch-RDSLive.ps1 -IntervalSeconds 10 -NoLogFile

.NOTES
    Press Ctrl+C to stop.
    Run as Administrator for Security log access.
#>
[CmdletBinding()]
param (
    [int]    $IntervalSeconds = 20,
    [string] $LogPath         = 'C:\Temp',
    [switch] $NoLogFile
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'SilentlyContinue'

# ── Log setup ──────────────────────────────────────────────────────────────────
$logFile = $null
if (-not $NoLogFile) {
    if (-not (Test-Path $LogPath)) { New-Item -ItemType Directory -Path $LogPath | Out-Null }
    $ts      = Get-Date -Format 'yyyyMMdd_HHmmss'
    $logFile = Join-Path $LogPath "RDS_Live_$ts.log"
}

function Write-Live {
    param([string]$Message, [string]$Level = 'INFO')
    $ts    = Get-Date -Format 'HH:mm:ss'
    $line  = "[$ts] [$Level] $Message"
    $color = switch ($Level) {
        'OK'      { 'Green'   }
        'WARN'    { 'Yellow'  }
        'FAIL'    { 'Red'     }
        'LICENSE' { 'Cyan'    }
        'SESSION' { 'White'   }
        'HEAD'    { 'Magenta' }
        default   { 'Gray'    }
    }
    Write-Host $line -ForegroundColor $color
    if ($logFile) { Add-Content -Path $logFile -Value $line -Encoding UTF8 }
}

# ── Session event descriptions ─────────────────────────────────────────────────
$sessionEvents = @{
    20 = @{ Label = 'Logon FAILED';      Level = 'FAIL'    }
    21 = @{ Label = 'Logon';             Level = 'SESSION' }
    22 = @{ Label = 'Shell start';       Level = 'SESSION' }
    23 = @{ Label = 'Logoff';            Level = 'SESSION' }
    24 = @{ Label = 'Disconnect';        Level = 'WARN'    }
    25 = @{ Label = 'Reconnect';         Level = 'SESSION' }
    39 = @{ Label = 'Session disconnect (limit)'; Level = 'WARN' }
    40 = @{ Label = 'Disconnect reason'; Level = 'WARN'    }
}

# Disconnect reason codes (Event 40 – Reason field)
$disconnectReasons = @{
    0  = 'No additional information'
    1  = 'User initiated'
    2  = 'Logoff by user'
    3  = 'Disconnect (idle timeout)'
    4  = 'Disconnect (logon timeout)'
    5  = 'Connection replaced'
    6  = 'Out of memory'
    7  = 'Server denied connection'
    8  = 'Server denied connection (no license)'
    9  = 'New connection limit reached'
    11 = 'Total login limit reached'
    12 = 'Disk full'
    16 = 'RDP keep-alive timeout'
    17 = 'Logon/reconnect failed'
    23 = 'License error'
    1025 = 'Internal error'
}

# ── Banner ─────────────────────────────────────────────────────────────────────
Clear-Host
Write-Host ''
Write-Host '  ════════════════════════════════════════════════' -ForegroundColor Magenta
Write-Host "   Watch-RDSLive  ·  $env:COMPUTERNAME" -ForegroundColor Magenta
Write-Host "   Started: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')  ·  interval: ${IntervalSeconds}s" -ForegroundColor Magenta
Write-Host '  ════════════════════════════════════════════════' -ForegroundColor Magenta
Write-Host '  Press Ctrl+C to stop.' -ForegroundColor DarkGray
if ($logFile) { Write-Host "  Log: $logFile" -ForegroundColor DarkGray }
Write-Host ''

# ── Initial status snapshot ────────────────────────────────────────────────────
Write-Live '--- INITIAL STATUS ---' 'HEAD'

# Active sessions
$sessions = @(quser 2>$null | Select-Object -Skip 1 | Where-Object { $_ -match '\S' })
Write-Live "Active sessions: $($sessions.Count)" 'INFO'
foreach ($s in $sessions) { Write-Live "  $($s.Trim())" 'SESSION' }

# Services
foreach ($svcName in @('TermService', 'SessionEnv', 'UmRdpService', 'TermServLicensing')) {
    $svc = Get-Service -Name $svcName -ErrorAction SilentlyContinue
    if ($svc) {
        $lvl = if ($svc.Status -eq 'Running') { 'OK' } else { 'FAIL' }
        Write-Live "Service $svcName : $($svc.Status)" $lvl
    }
}

# Check if licensing mode is configured
$licReg  = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services'
$licMode = (Get-ItemProperty $licReg -ErrorAction SilentlyContinue).LicensingMode
$licSrv  = (Get-ItemProperty $licReg -ErrorAction SilentlyContinue).LicenseServers
$licStr  = switch ($licMode) { 2 { 'Per Device' } 4 { 'Per User' } $null { 'Not set via GPO' } default { "Mode $licMode" } }
Write-Live "Licensing mode  : $licStr$(if ($licSrv) { "  |  server: $licSrv" })" 'LICENSE'

Write-Live '--- WATCHING FOR EVENTS ---' 'HEAD'
Write-Host ''

# ── Polling loop ───────────────────────────────────────────────────────────────
$lastCheck = (Get-Date).AddSeconds(-2)   # slight overlap to avoid missing events

while ($true) {
    $pollStart = Get-Date
    $newEvents = 0

    # ── TerminalServices-LocalSessionManager ──────────────────────────────────
    try {
        $smEvents = Get-WinEvent -FilterHashtable @{
            LogName   = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational'
            Id        = @(20, 21, 22, 23, 24, 25, 39, 40)
            StartTime = $lastCheck
        } -ErrorAction Stop | Sort-Object TimeCreated

        foreach ($ev in $smEvents) {
            $info   = $sessionEvents[$ev.Id]
            $label  = $info.Label
            $level  = $info.Level
            $user   = try { $ev.Properties[0].Value } catch { '?' }
            $session= try { $ev.Properties[1].Value } catch { '' }

            $msg = "$label | user: $user$(if ($session) { "  session: $session" })"

            # Enrich disconnect reason (Event 40)
            if ($ev.Id -eq 40) {
                $reasonCode = try { [int]$ev.Properties[2].Value } catch { -1 }
                $reasonText = if ($disconnectReasons.ContainsKey($reasonCode)) {
                    $disconnectReasons[$reasonCode]
                } else { "code $reasonCode" }
                $msg += "  reason: $reasonText"
                if ($reasonCode -in @(8, 23)) { $level = 'FAIL' }  # license-related
            }

            Write-Live $msg $level
            $newEvents++
        }
    } catch {}

    # ── Security: failed RDP logons (4625 type 10) + lockouts (4740) ──────────
    try {
        $secEvents = Get-WinEvent -FilterHashtable @{
            LogName   = 'Security'
            Id        = @(4625, 4740)
            StartTime = $lastCheck
        } -MaxEvents 50 -ErrorAction Stop | Sort-Object TimeCreated

        foreach ($ev in $secEvents) {
            if ($ev.Id -eq 4625) {
                $logonType = try { [int]$ev.Properties[10].Value } catch { 0 }
                if ($logonType -ne 10) { continue }   # only RDP (type 10)
                $user   = try { $ev.Properties[5].Value } catch { '?' }
                $domain = try { $ev.Properties[6].Value } catch { '' }
                $reason = try { $ev.Properties[9].Value } catch { '' }
                Write-Live "FAILED LOGON | user: $domain\$user  reason: $reason" 'FAIL'
                $newEvents++
            } elseif ($ev.Id -eq 4740) {
                $user = try { $ev.Properties[0].Value } catch { '?' }
                $src  = try { $ev.Properties[1].Value } catch { '?' }
                Write-Live "ACCOUNT LOCKED OUT | user: $user  source: $src" 'FAIL'
                $newEvents++
            }
        }
    } catch {}

    # ── Licensing events ───────────────────────────────────────────────────────
    # TerminalServices-Licensing/Admin
    try {
        $licEvents = Get-WinEvent -FilterHashtable @{
            LogName   = 'Microsoft-Windows-TerminalServices-Licensing/Admin'
            StartTime = $lastCheck
        } -MaxEvents 20 -ErrorAction Stop | Sort-Object TimeCreated

        foreach ($ev in $licEvents) {
            $level = switch ($ev.Level) { 1 { 'FAIL' } 2 { 'FAIL' } 3 { 'WARN' } default { 'LICENSE' } }
            $msg   = ($ev.Message -replace '\s+', ' ')
            $msg   = $msg.Substring(0, [Math]::Min(160, $msg.Length))
            Write-Live "LICENSE | ID $($ev.Id): $msg" $level
            $newEvents++
        }
    } catch {}

    # System log — TermServLicensing provider
    try {
        $sysLic = Get-WinEvent -FilterHashtable @{
            LogName      = 'System'
            ProviderName = 'TermServLicensing'
            StartTime    = $lastCheck
        } -MaxEvents 20 -ErrorAction Stop | Sort-Object TimeCreated

        foreach ($ev in $sysLic) {
            $level = if ($ev.Level -le 2) { 'FAIL' } elseif ($ev.Level -eq 3) { 'WARN' } else { 'LICENSE' }
            $msg   = ($ev.Message -replace '\s+', ' ')
            $msg   = $msg.Substring(0, [Math]::Min(160, $msg.Length))
            Write-Live "LICENSE (System) | ID $($ev.Id): $msg" $level
            $newEvents++
        }
    } catch {}

    # ── Heartbeat line (every poll) ────────────────────────────────────────────
    $activeSessions = @(quser 2>$null | Select-Object -Skip 1 | Where-Object { $_ -match '\S' })
    $elapsed = [math]::Round(((Get-Date) - $pollStart).TotalMilliseconds)
    $heartbeat = "[ poll $($pollStart.ToString('HH:mm:ss')) | sessions: $($activeSessions.Count) | new events: $newEvents | ${elapsed}ms ]"
    Write-Host $heartbeat -ForegroundColor DarkGray
    if ($logFile) { Add-Content -Path $logFile -Value $heartbeat -Encoding UTF8 }

    $lastCheck = $pollStart
    Start-Sleep -Seconds $IntervalSeconds
}
