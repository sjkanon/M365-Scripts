#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Set the default Num Lock state (on/off) for new user profiles and the sign-in
    screen.

.DESCRIPTION
    Writes InitialKeyboardIndicators under HKEY_USERS\.DEFAULT — the template
    profile new user accounts are based on — and to the currently logged-on
    user's own hive, so the setting takes effect immediately without waiting for
    a new profile to be created. Also applies it to the sign-in screen (HKU\.DEFAULT
    is what's loaded there). Defaults to a safe preview — pass -Apply to actually
    write the value.

.PARAMETER State
    "On" or "Off". Default: On.

.PARAMETER Apply
    Actually write the registry value. Without this switch, the script only
    reports what it would do.

.EXAMPLE
    .\Set-NumLockDefault.ps1

.EXAMPLE
    .\Set-NumLockDefault.ps1 -State On -Apply

.EXAMPLE
    .\Set-NumLockDefault.ps1 -State Off -Apply

.NOTES
    Registry value: InitialKeyboardIndicators — 2147483650 (0x80000002) / "2" = on
    at every logon, 2147483648 (0x80000000) / "0" = off. This script uses the
    simple "2"/"0" values, matching Windows' own default profile convention.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [ValidateSet('On', 'Off')]
    [string] $State = 'On',

    [switch] $Apply
)

$value = if ($State -eq 'On') { '2' } else { '0' }
$targets = @(
    @{ Path = 'Registry::HKEY_USERS\.DEFAULT\Control Panel\Keyboard'; Label = 'Default profile / sign-in screen' }
)

# Include the current user's own hive too, so an interactive admin sees the
# change immediately without needing a new profile.
if (Test-Path 'Registry::HKEY_CURRENT_USER\Control Panel\Keyboard') {
    $targets += @{ Path = 'Registry::HKEY_CURRENT_USER\Control Panel\Keyboard'; Label = 'Current user' }
}

Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Set-NumLockDefault" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  State : $State (InitialKeyboardIndicators = $value)"
Write-Host ("  Mode  : {0}" -f $(if ($Apply) { 'Apply' } else { 'Preview only' })) -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

foreach ($target in $targets) {
    if (-not $Apply) {
        Write-Host "  [PREVIEW] $($target.Label): $($target.Path)\InitialKeyboardIndicators -> $value" -ForegroundColor DarkGray
        continue
    }

    if (-not $PSCmdlet.ShouldProcess($target.Path, "Set InitialKeyboardIndicators = $value")) { continue }

    try {
        Set-ItemProperty -Path $target.Path -Name 'InitialKeyboardIndicators' -Value $value -ErrorAction Stop
        Write-Host "  [OK]   $($target.Label)" -ForegroundColor Green
    } catch {
        Write-Host "  [WARN] $($target.Label): $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

Write-Host ""
if (-not $Apply) { Write-Host "  Re-run with -Apply to set this default." -ForegroundColor Yellow }
Write-Host ""
