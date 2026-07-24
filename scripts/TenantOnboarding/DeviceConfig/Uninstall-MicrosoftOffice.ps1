#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Silently uninstall Microsoft Office (Click-to-Run) from a device.

.DESCRIPTION
    Finds installed Microsoft Office 365 / Microsoft 365 Apps entries in the registry
    uninstall keys (both 32-bit and 64-bit views) and runs each uninstall string
    silently with DisplayLevel=False appended — the modernized, dry-run-by-default
    equivalent of an old always-silent removal script. Useful before a clean
    reinstall/migration.

.PARAMETER NameFilter
    Wildcard filter (PowerShell -like syntax) applied to each uninstall entry's
    DisplayName. Default: "*Microsoft 365*" and "*Microsoft Office*" are both matched.

.PARAMETER Apply
    Actually run the uninstall(s). Without this switch, the script only lists what
    would be removed.

.EXAMPLE
    .\Uninstall-MicrosoftOffice.ps1

.EXAMPLE
    .\Uninstall-MicrosoftOffice.ps1 -Apply
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [switch] $Apply
)

$uninstallKeys = @(
    'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
)

Write-Host ""
Write-Host "  Uninstall-MicrosoftOffice" -ForegroundColor Cyan
Write-Host "  Mode : $(if ($Apply) { 'Apply' } else { 'Preview only' })" -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

$entries = Get-ItemProperty -Path $uninstallKeys -ErrorAction SilentlyContinue |
    Where-Object { $_.DisplayName -like '*Microsoft 365*' -or $_.DisplayName -like '*Microsoft Office*' } |
    Where-Object { $_.UninstallString }

if (-not $entries) {
    Write-Host "  No Microsoft Office / Microsoft 365 Apps installation found." -ForegroundColor DarkGray
    exit 0
}

foreach ($entry in $entries) {
    Write-Host "  Found: $($entry.DisplayName)" -ForegroundColor Cyan

    if (-not $Apply) {
        Write-Host "    [PREVIEW] Would run: $($entry.UninstallString) DisplayLevel=False" -ForegroundColor Yellow
        continue
    }
    if (-not $PSCmdlet.ShouldProcess($entry.DisplayName, "Silently uninstall")) { continue }

    $parts = $entry.UninstallString -split '"'
    $exe  = $parts[1]
    $args = "$($parts[2]) DisplayLevel=False".Trim()

    try {
        Start-Process -FilePath $exe -ArgumentList $args -Wait -ErrorAction Stop
        Write-Host "    [OK]   Uninstalled '$($entry.DisplayName)'." -ForegroundColor Green
    } catch {
        Write-Host "    [WARN] Failed: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

Write-Host ""
if (-not $Apply) { Write-Host "  Re-run with -Apply to uninstall the item(s) shown above." -ForegroundColor Yellow }
Write-Host ""
