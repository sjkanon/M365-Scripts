#Requires -Version 5.1
#Requires -Modules ActiveDirectory
<#
.SYNOPSIS
    Monitor on-prem Active Directory for locked-out user accounts and log new lockouts.

.DESCRIPTION
    Queries the domain's PDC Emulator for every currently locked-out user account. The PDC
    Emulator is used deliberately — bad-password-attempt counters and lockoutTime are tracked
    per-DC and only guaranteed authoritative on the PDC Emulator until the rest of the lockout
    replicates, so querying a random DC (or the local one) can under-report or miss lockouts.

    State (which accounts were already known-locked, and since when) is kept in a small JSON
    file next to the log, so only a NEW lockout — first time seen, or an unlock-then-relock of
    the same account — is written to the log. An account that stays locked across many runs is
    not re-logged every run.

    Run without -RegisterTask for a single check (useful to test manually). Pass -RegisterTask
    once to register this script as a Windows Scheduled Task that repeats every
    -TaskIntervalMinutes (default 20), running as SYSTEM. Intended to run on a Domain
    Controller, or a file server with the ActiveDirectory RSAT module installed and network
    access to a DC.

.PARAMETER SearchBase
    Optional. Limit the search to a specific OU (distinguished name). Default: whole domain.

.PARAMETER LogPath
    Override the default log file (default: C:\ProgramData\ADLockoutMonitor\lockouts.log).
    The state JSON file is written next to it.

.PARAMETER RegisterTask
    Register this script as a recurring Scheduled Task instead of running a single check.

.PARAMETER TaskIntervalMinutes
    Repetition interval in minutes, only used with -RegisterTask (default: 20).

.EXAMPLE
    # One-off check — also what the scheduled task runs every 20 minutes
    .\Watch-ADAccountLockouts.ps1

.EXAMPLE
    # Register the recurring scheduled task (run once, on the DC/file server)
    .\Watch-ADAccountLockouts.ps1 -RegisterTask

.EXAMPLE
    # Limit to one OU, every 5 minutes
    .\Watch-ADAccountLockouts.ps1 -RegisterTask -TaskIntervalMinutes 5 -SearchBase "OU=Sales,DC=contoso,DC=com"
#>

[CmdletBinding()]
param(
    [string]$SearchBase,
    [string]$LogPath = (Join-Path $env:ProgramData 'ADLockoutMonitor\lockouts.log'),
    [switch]$RegisterTask,
    [int]$TaskIntervalMinutes = 20
)

$ErrorActionPreference = 'Stop'

$logDir = Split-Path -Parent $LogPath
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }

if ($RegisterTask) {
    $scriptPath = $MyInvocation.MyCommand.Path
    $taskArgs = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -LogPath `"$LogPath`""
    if ($SearchBase) { $taskArgs += " -SearchBase `"$SearchBase`"" }

    $action    = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $taskArgs
    $trigger   = New-ScheduledTaskTrigger -Once -At (Get-Date) `
        -RepetitionInterval (New-TimeSpan -Minutes $TaskIntervalMinutes) `
        -RepetitionDuration ([TimeSpan]::MaxValue)
    $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest

    Register-ScheduledTask -TaskName 'AD Account Lockout Monitor' -Action $action -Trigger $trigger `
        -Principal $principal -Force | Out-Null
    Write-Host "Scheduled Task 'AD Account Lockout Monitor' geregistreerd — draait elke $TaskIntervalMinutes minuten als SYSTEM." -ForegroundColor Green
    exit 0
}

function Write-LockLog {
    param([string]$Message)
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $Message"
    Add-Content -Path $LogPath -Value $line
    Write-Host $line
}

$statePath = Join-Path $logDir 'state.json'

# Authoritative lockout source — see .DESCRIPTION.
$pdc = (Get-ADDomain).PDCEmulator

$searchParams = @{ LockedOut = $true; UsersOnly = $true; Server = $pdc }
if ($SearchBase) { $searchParams['SearchBase'] = $SearchBase }

$lockedNow = @(Search-ADAccount @searchParams | ForEach-Object {
    Get-ADUser -Identity $_.SamAccountName -Server $pdc -Properties LockoutTime, DisplayName, DistinguishedName
})

$previousMap = @{}
if (Test-Path $statePath) {
    $previousState = Get-Content $statePath -Raw | ConvertFrom-Json -ErrorAction SilentlyContinue
    foreach ($p in @($previousState)) {
        if ($p -and $p.SamAccountName) { $previousMap[$p.SamAccountName] = [int64]$p.LockoutTime }
    }
}

$newState = [System.Collections.Generic.List[object]]::new()
$newLockoutCount = 0

foreach ($user in $lockedNow) {
    $lockoutTimeUtc = if ($user.LockoutTime) { [datetime]::FromFileTimeUtc($user.LockoutTime) } else { [datetime]::MinValue }
    $key   = $user.SamAccountName
    $stamp = [int64]$lockoutTimeUtc.Ticks

    $newState.Add([PSCustomObject]@{ SamAccountName = $key; LockoutTime = $stamp })

    if (-not $previousMap.ContainsKey($key) -or $previousMap[$key] -ne $stamp) {
        $newLockoutCount++
        Write-LockLog ("NIEUWE LOCKOUT: {0} ({1}) - locked op {2:yyyy-MM-dd HH:mm:ss} UTC - {3}" -f `
            $key, $user.DisplayName, $lockoutTimeUtc, $user.DistinguishedName)
    }
}

if ($newLockoutCount -eq 0) {
    Write-Verbose ("Geen nieuwe lockouts ({0} account(s) momenteel locked, al eerder gemeld)." -f $lockedNow.Count)
}

(ConvertTo-Json -InputObject $newState) | Set-Content -Path $statePath -Encoding UTF8
