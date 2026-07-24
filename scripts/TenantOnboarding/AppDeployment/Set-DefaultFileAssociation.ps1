#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Set the default application for a file extension, bypassing Windows' UserChoice hash protection.

.DESCRIPTION
    Since Windows 10 1803, direct registry writes to the per-extension "UserChoice"
    key are protected by a hash Windows validates, so a plain Set-ItemProperty no
    longer sticks. This script wraps the community PS-SFTA tool
    (https://github.com/DanysysTeam/PS-SFTA), which computes the correct hash, to set
    a default file association the supported way.

    Generic replacement for a set of one-off scripts that each hardcoded one
    ProgId/extension pair (7-Zip for .zip/.rar, Adobe Acrobat for .pdf, ...).

.PARAMETER ProgId
    The registered ProgId (or Application\<exe> path) to set as default, e.g.
    "Applications\7zFM.exe" or "Acrobat.Document.DC".

.PARAMETER Extension
    File extension(s) to associate, e.g. ".zip" or ".pdf". Accepts multiple values.

.PARAMETER Apply
    Actually change the association. Without this switch, the script only reports
    what it would do.

.EXAMPLE
    # Dry run
    .\Set-DefaultFileAssociation.ps1 -ProgId "Applications\7zFM.exe" -Extension ".zip",".rar"

.EXAMPLE
    .\Set-DefaultFileAssociation.ps1 -ProgId "Applications\7zFM.exe" -Extension ".zip",".rar" -Apply

.EXAMPLE
    .\Set-DefaultFileAssociation.ps1 -ProgId "Acrobat.Document.DC" -Extension ".pdf" -Apply

.NOTES
    Downloads and imports the PS-SFTA module from GitHub at runtime
    (DanysysTeam/PS-SFTA). Review/vendor that module locally first if your policy
    requires offline or pre-approved script sources.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string] $ProgId,

    [Parameter(Mandatory)]
    [string[]] $Extension,

    [switch] $Apply
)

Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Set-DefaultFileAssociation" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "  ProgId     : $ProgId"
Write-Host "  Extensions : $($Extension -join ', ')"
Write-Host "  Mode       : $(if ($Apply) { 'Apply' } else { 'Preview only' })" -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

if (-not $Apply) {
    foreach ($ext in $Extension) { Write-Host "  Would set '$ProgId' as default handler for '$ext'." -ForegroundColor Yellow }
    Write-Host "  Re-run with -Apply to perform this change." -ForegroundColor Yellow
    Write-Host ""
    return
}

if (-not $PSCmdlet.ShouldProcess("$($Extension -join ', ')", "Set default app to $ProgId")) { return }

$tempDir = Join-Path $env:TEMP 'PS-SFTA'
$zipPath = Join-Path $env:TEMP 'PS-SFTA.zip'

try {
    Write-Host "  Downloading PS-SFTA..." -ForegroundColor DarkGray
    Invoke-WebRequest -Uri 'https://github.com/DanysysTeam/PS-SFTA/archive/refs/heads/master.zip' -OutFile $zipPath -ErrorAction Stop
    if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force }
    Expand-Archive -Path $zipPath -DestinationPath $tempDir -Force

    $module = Get-ChildItem -Path $tempDir -Filter 'sfta.ps1' -Recurse | Select-Object -First 1
    Import-Module $module.FullName -Force

    foreach ($ext in $Extension) {
        Set-FTA -ProgId $ProgId -Extension $ext
        Write-Host "  [OK]   '$ProgId' set as default for '$ext'." -ForegroundColor Green
    }
} catch {
    Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
    exit 1
} finally {
    Remove-Item -Path $zipPath -Force -ErrorAction SilentlyContinue
    Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host ""
