#Requires -Version 3
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Create the inbound firewall rule Microsoft Teams needs for LAN peer-to-peer screen sharing.

.DESCRIPTION
    Modernized wrapper around the well-known community "Update-TeamsFWRule" pattern
    (originally by Michael Mardahl, based on Microsoft's own sample at
    https://docs.microsoft.com/microsoftteams/get-clients#sample-powershell-script):
    creates an inbound firewall rule scoped to the currently logged-on user's
    Teams.exe, allowing same-domain-profile traffic (needed for optimal in-office LAN
    screen sharing) while blocking it on Public/Private profiles.

    Designed to run as SYSTEM (Intune Win32 app or logon-triggered scheduled task) so
    it can resolve "the user who is currently logged on".

.PARAMETER Force
    Remove any pre-existing rule(s) for the resolved Teams.exe path before creating a
    fresh pair. Default: on (matches the original script's recommended setting).

.PARAMETER Apply
    Actually create the firewall rule(s). Without this switch, the script only
    reports what it would do.

.EXAMPLE
    .\Set-TeamsFirewallRule.ps1

.EXAMPLE
    .\Set-TeamsFirewallRule.ps1 -Apply

.NOTES
    Original concept (c) Microsoft Corporation 2018 and Michael Mardahl
    (www.iphase.dk / www.msendpointmgr.com), provided as-is. This is a house-style
    rewrite, not a verbatim copy.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [switch] $Force = $true,
    [switch] $Apply
)

function Get-LoggedOnUserProfile {
    $loggedOnUser = (Get-CimInstance -ClassName Win32_ComputerSystem).UserName
    if (-not $loggedOnUser) { throw "No user is currently logged on to the primary session." }
    $username = ($loggedOnUser -split '\\')[-1]
    Get-ChildItem (Join-Path $env:SystemDrive 'Users') -ErrorAction Stop |
        Where-Object Name -like "$username*" | Select-Object -First 1
}

Write-Host ""
Write-Host "  Set-TeamsFirewallRule" -ForegroundColor Cyan
Write-Host "  Mode : $(if ($Apply) { 'Apply' } else { 'Preview only' })" -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

try {
    $profileObj = Get-LoggedOnUserProfile
} catch {
    Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

$progPath = Join-Path $profileObj.FullName 'AppData\Local\Microsoft\Teams\Current\Teams.exe'
Write-Host "  User        : $($profileObj.Name)"
Write-Host "  Teams.exe   : $progPath $(if (-not (Test-Path $progPath)) { '(not found yet)' })"
Write-Host ""

if (-not $Apply) {
    Write-Host "  Would create inbound Allow rule (Domain profile) and inbound Block rule (Public/Private) for $progPath" -ForegroundColor Yellow
    Write-Host "  Re-run with -Apply to create these rules." -ForegroundColor Yellow
    Write-Host ""
    exit 0
}

if (-not $PSCmdlet.ShouldProcess($progPath, "Create Teams firewall rules")) { exit 0 }

if ($Force) {
    Get-NetFirewallApplicationFilter -Program $progPath -ErrorAction SilentlyContinue |
        Remove-NetFirewallRule -ErrorAction SilentlyContinue
}

if (-not (Get-NetFirewallApplicationFilter -Program $progPath -ErrorAction SilentlyContinue)) {
    $ruleName = "Teams.exe for user $($profileObj.Name)"
    New-NetFirewallRule -DisplayName $ruleName -Direction Inbound -Profile Domain -Program $progPath -Action Allow -Protocol Any | Out-Null
    New-NetFirewallRule -DisplayName $ruleName -Direction Inbound -Profile Public, Private -Program $progPath -Action Block -Protocol Any | Out-Null
    Write-Host "  [OK]   Firewall rules created: '$ruleName'" -ForegroundColor Green
} else {
    Write-Host "  [SKIP] A rule for this Teams.exe path already exists." -ForegroundColor DarkGray
}

Write-Host ""
