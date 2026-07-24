#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Download a zipped PSAppDeployToolkit (or similar) package and run its silent install.

.DESCRIPTION
    Generic replacement for a whole family of near-identical old scripts that each
    hardcoded one vendor's download URL (Adobe Reader, AnyDesk, Citrix Workspace,
    Google Drive, Jabra Direct, TeamViewer, ...) and repeated the same four steps:
    download a zip from an internal software repository, expand it, run its
    `Deploy-<App>.ps1 -DeploymentType Install -DeployMode NonInteractive` (the
    standard PSAppDeployToolkit entry point), then clean up the temp files.

    This script performs the same four steps generically for any app — point it at
    your own internal package repository with -SourceUri. No vendor URL is hardcoded.

    Optionally registers a "run at every logon" scheduled task that re-invokes this
    same script, for self-updating desktop clients that need to check for a newer
    version on each sign-in (the old per-vendor scripts did this by hand for a
    specific in-house desktop client installer).

.PARAMETER AppName
    Friendly name of the application — used for the working folder name and the
    deployment script name (Deploy-<AppName>.ps1) inside the package if
    -DeployScriptName is not given.

.PARAMETER SourceUri
    URL to the .zip package to download. Required — no default (do not hardcode an
    internal endpoint here; pass your own repository URL per app).

.PARAMETER DeployScriptName
    Name of the PSAppDeployToolkit entry-point script inside the expanded zip.
    Default: "Deploy-<AppName>.ps1".

.PARAMETER DeploymentType
    PSAppDeployToolkit -DeploymentType value. Default: "Install".

.PARAMETER DeployMode
    PSAppDeployToolkit -DeployMode value. Default: "NonInteractive".

.PARAMETER WorkingRoot
    Local folder used for the download + extraction + execution. Default: C:\Install.

.PARAMETER RegisterLogonUpdateTask
    Also register a scheduled task (trigger: at logon) that re-runs this exact command
    line, so the app is re-checked/updated on every sign-in. Requires -Apply.

.PARAMETER Apply
    Actually download and run the installer. Without this switch, the script only
    reports what it would do.

.EXAMPLE
    # Dry run
    .\Install-Win32AppPackage.ps1 -AppName "AdobeReaderDC" -SourceUri "https://packages.contoso.com/Installers/AdobeReaderDC.zip"

.EXAMPLE
    .\Install-Win32AppPackage.ps1 -AppName "AdobeReaderDC" -SourceUri "https://packages.contoso.com/Installers/AdobeReaderDC.zip" -Apply

.EXAMPLE
    # Self-updating desktop client, re-checked at every logon
    .\Install-Win32AppPackage.ps1 -AppName "LineOfBusinessDesktop" `
        -SourceUri "https://packages.contoso.com/Installers/LineOfBusinessDesktop.zip" `
        -RegisterLogonUpdateTask -Apply

.NOTES
    Designed for zip packages built with PSAppDeployToolkit (https://psappdeploytoolkit.com/),
    the de facto standard for silent Win32 app deployment via Intune/SCCM.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string] $AppName,

    [Parameter(Mandatory)]
    [string] $SourceUri,

    [string] $DeployScriptName,
    [string] $DeploymentType = 'Install',
    [string] $DeployMode = 'NonInteractive',
    [string] $WorkingRoot = 'C:\Install',
    [switch] $RegisterLogonUpdateTask,
    [switch] $Apply
)

if (-not $DeployScriptName) { $DeployScriptName = "Deploy-$AppName.ps1" }

$tempZip     = Join-Path $env:TEMP "$AppName.zip"
$extractPath = Join-Path $WorkingRoot $AppName

Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Install-Win32AppPackage : $AppName" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "  Source : $SourceUri"
Write-Host "  Mode   : $(if ($Apply) { 'Apply' } else { 'Preview only' })" -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

if (-not $Apply) {
    Write-Host "  Would download $SourceUri -> $tempZip" -ForegroundColor Yellow
    Write-Host "  Would expand to $extractPath and run $DeployScriptName -DeploymentType $DeploymentType -DeployMode $DeployMode" -ForegroundColor Yellow
    if ($RegisterLogonUpdateTask) { Write-Host "  Would register a logon scheduled task to re-run this install." -ForegroundColor Yellow }
    Write-Host "  Re-run with -Apply to perform the installation." -ForegroundColor Yellow
    Write-Host ""
    return
}

if (-not $PSCmdlet.ShouldProcess($AppName, "Download and silently install")) { return }

try {
    if (-not (Test-Path $WorkingRoot)) { New-Item -ItemType Directory -Path $WorkingRoot -Force | Out-Null }

    Write-Host "  Downloading package..." -ForegroundColor DarkGray
    Invoke-WebRequest -Uri $SourceUri -OutFile $tempZip -ErrorAction Stop

    Write-Host "  Expanding package..." -ForegroundColor DarkGray
    if (Test-Path $extractPath) { Remove-Item $extractPath -Recurse -Force }
    Expand-Archive -Path $tempZip -DestinationPath $extractPath -Force

    $deployScript = Get-ChildItem -Path $extractPath -Filter $DeployScriptName -Recurse | Select-Object -First 1
    if (-not $deployScript) { throw "Could not find $DeployScriptName inside the extracted package." }

    Write-Host "  Running $($deployScript.Name) -DeploymentType $DeploymentType -DeployMode $DeployMode ..." -ForegroundColor DarkGray
    Push-Location $deployScript.DirectoryName
    try {
        & powershell.exe -ExecutionPolicy Bypass -File $deployScript.FullName -DeploymentType $DeploymentType -DeployMode $DeployMode
    } finally {
        Pop-Location
    }

    Write-Host "  [OK]   Deployment script completed." -ForegroundColor Green
} catch {
    Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
    exit 1
} finally {
    Remove-Item -Path $tempZip -Force -ErrorAction SilentlyContinue
    Remove-Item -Path $extractPath -Recurse -Force -ErrorAction SilentlyContinue
}

if ($RegisterLogonUpdateTask) {
    Write-Host "  Registering logon update task..." -ForegroundColor DarkGray
    $scriptPath = $MyInvocation.MyCommand.Path
    $argument = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`" -AppName `"$AppName`" -SourceUri `"$SourceUri`" -Apply"
    $trigger = New-ScheduledTaskTrigger -AtLogOn
    $action  = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $argument
    Register-ScheduledTask -TaskName "$AppName Update" -Trigger $trigger -Action $action -User $env:USERNAME -Force | Out-Null
    Write-Host "  [OK]   Scheduled task '$AppName Update' registered (runs at logon)." -ForegroundColor Green
}

Write-Host ""
