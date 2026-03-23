#Requires -Version 5.1
# Cross-platform: Windows (PS 5.1+), macOS and Linux (PS 7+)
# Install PowerShell 7: https://aka.ms/powershell
<#
.SYNOPSIS
    Tests TCP connectivity on one or more ports against a target host.

.DESCRIPTION
    Checks which ports are open or closed on a given IP address or hostname.
    Supports individual ports, ranges (1294:1494), and comma-separated combinations.
    Uses TCP connect attempts with a configurable timeout.
    Works on Windows (PowerShell 5.1+), macOS, and Linux (PowerShell 7+).

.PARAMETER Target
    IP address or hostname to test. Accepts multiple targets.

.PARAMETER Ports
    Port specification. Supports:
      - Single port:    80
      - Range:          1294:1494
      - List:           80,443,3389
      - Mix:            80,443,1294:1494,8080:8090

.PARAMETER TimeoutMs
    TCP connection timeout in milliseconds. Default: 500.

.PARAMETER ShowClosed
    Include closed ports in the output. Default: only open ports are shown.

.EXAMPLE
    .\Test-Ports.ps1 -Target 192.168.1.1 -Ports 1294:1494

.EXAMPLE
    .\Test-Ports.ps1 -Target 10.0.0.1 -Ports 80,443,3389,8080:8090

.EXAMPLE
    .\Test-Ports.ps1 -Target server01.contoso.local -Ports 22,3389 -TimeoutMs 1000 -ShowClosed

.EXAMPLE
    .\Test-Ports.ps1 -Target 10.0.0.1,10.0.0.2 -Ports 80,443
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory, Position = 0)]
    [string[]] $Target,

    [Parameter(Mandatory, Position = 1)]
    [string[]] $Ports,

    [int]    $TimeoutMs  = 500,
    [switch] $ShowClosed
)

# ── Parse port specification into a list of integers ──────────────────────────
function Resolve-Ports {
    param([string[]] $PortSpec)

    $result = [System.Collections.Generic.List[int]]::new()

    foreach ($entry in $PortSpec) {
        # Handle comma-separated values passed as a single string
        foreach ($part in $entry -split ',') {
            $part = $part.Trim()
            if ($part -match '^(\d+)[:\-](\d+)$') {
                $from = [int]$Matches[1]
                $to   = [int]$Matches[2]
                if ($from -gt $to) { $from, $to = $to, $from }
                for ($p = $from; $p -le $to; $p++) { $result.Add($p) }
            } elseif ($part -match '^\d+$') {
                $result.Add([int]$part)
            } else {
                Write-Warning "Skipping unrecognised port specification: '$part'"
            }
        }
    }

    return $result | Sort-Object -Unique
}

# ── Test a single TCP port ────────────────────────────────────────────────────
function Test-TcpPort {
    param([string] $Hostname, [int] $Port, [int] $Timeout)

    $client = [System.Net.Sockets.TcpClient]::new()
    try {
        $ar = $client.BeginConnect($Hostname, $Port, $null, $null)
        $ok = $ar.AsyncWaitHandle.WaitOne($Timeout, $false)
        if ($ok -and $client.Connected) {
            return $true
        }
        return $false
    } catch {
        return $false
    } finally {
        $client.Close()
    }
}

# ── Main ──────────────────────────────────────────────────────────────────────
$portList = Resolve-Ports -PortSpec $Ports

if ($portList.Count -eq 0) {
    Write-Error "No valid ports specified."
    exit 1
}

$totalPorts = $portList.Count * $Target.Count
Write-Host ""
Write-Host "  Target(s) : $($Target -join ', ')"
if ($portList.Count -gt 10) {
    Write-Host "  Ports     : $($portList[0])..$($portList[-1]) ($($portList.Count) ports)"
} else {
    Write-Host "  Ports     : $($portList -join ', ')"
}
Write-Host "  Timeout   : ${TimeoutMs}ms per port"
Write-Host "  Total     : $totalPorts checks"
Write-Host ""

$results = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($t in $Target) {
    Write-Host "  Scanning $t ..." -ForegroundColor Cyan

    # Resolve hostname to IP once
    try {
        $resolved = ([System.Net.Dns]::GetHostAddresses($t) | Where-Object { $_.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork } | Select-Object -First 1).ToString()
    } catch {
        $resolved = $t
    }

    $open   = 0
    $closed = 0

    foreach ($port in $portList) {
        $isOpen = Test-TcpPort -Host $t -Port $port -Timeout $TimeoutMs

        if ($isOpen) {
            $open++
            $results.Add([PSCustomObject]@{
                Target   = $t
                Resolved = $resolved
                Port     = $port
                Status   = "OPEN"
            })
        } else {
            $closed++
            if ($ShowClosed) {
                $results.Add([PSCustomObject]@{
                    Target   = $t
                    Resolved = $resolved
                    Port     = $port
                    Status   = "closed"
                })
            }
        }
    }

    Write-Host "    Open: $open  /  Closed: $closed" -ForegroundColor DarkGray
}

# ── Output ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor White
Write-Host "   Results" -ForegroundColor White
Write-Host "  ================================================" -ForegroundColor White
Write-Host ""

if ($results.Count -eq 0) {
    Write-Host "  No open ports found." -ForegroundColor Yellow
} else {
    foreach ($r in $results | Sort-Object Target, Port) {
        $color = if ($r.Status -eq "OPEN") { "Green" } else { "DarkGray" }
        $label = "{0,-22} Port {1,-6} {2}" -f $r.Target, $r.Port, $r.Status
        Write-Host "  $label" -ForegroundColor $color
    }
}

Write-Host ""

# Return results as objects so the caller can pipe/filter
return $results
