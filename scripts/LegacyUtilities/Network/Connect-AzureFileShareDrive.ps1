#Requires -Version 5.1
<#
.SYNOPSIS
    Mount an Azure Files SMB share as a persistent drive letter.

.DESCRIPTION
    Generalized replacement for a script that hardcoded a specific storage
    account name, share path, and storage account key. Tests SMB (port 445)
    connectivity to the storage account first — Azure Files over SMB requires
    outbound TCP 445, which many ISPs/networks block — then saves the storage
    account key via cmdkey and maps the share with New-PSDrive.

    Storage account key is always supplied by the caller (-StorageAccountKey or
    a SecureString via -Credential), never hardcoded.

.PARAMETER StorageAccountName
    Azure Storage account name (the part before ".file.core.windows.net").

.PARAMETER ShareName
    Name of the file share within the storage account.

.PARAMETER StorageAccountKey
    Storage account access key (plain string — pass securely, e.g. from a secret
    store; not saved to disk by this script beyond what cmdkey itself persists).

.PARAMETER DriveLetter
    Drive letter to mount the share as (e.g. "S"). Default: "S".

.PARAMETER Persist
    Make the New-PSDrive mapping persist across reboots.

.PARAMETER Apply
    Actually save the credential and mount the drive. Without this switch, the
    script only tests connectivity and reports what it would do.

.EXAMPLE
    # Preview — test connectivity only
    .\Connect-AzureFileShareDrive.ps1 -StorageAccountName "contosofiles" -ShareName "documents" -StorageAccountKey $key

.EXAMPLE
    .\Connect-AzureFileShareDrive.ps1 -StorageAccountName "contosofiles" -ShareName "documents" -StorageAccountKey $key -DriveLetter S -Persist -Apply

.NOTES
    Requires outbound TCP 445 to *.file.core.windows.net — blocked by some
    ISPs/firewalls. Use an Azure P2S/S2S VPN or ExpressRoute to tunnel SMB
    traffic over a different port if 445 is unavailable.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string] $StorageAccountName,

    [Parameter(Mandatory)]
    [string] $ShareName,

    [Parameter(Mandatory)]
    [string] $StorageAccountKey,

    [string] $DriveLetter = 'S',
    [switch] $Persist,
    [switch] $Apply
)

$hostName = "$StorageAccountName.file.core.windows.net"
$sharePath = "\\$hostName\$ShareName"

Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Connect-AzureFileShareDrive" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Share  : $sharePath"
Write-Host "  Drive  : ${DriveLetter}:"
Write-Host ("  Mode   : {0}" -f $(if ($Apply) { 'Apply' } else { 'Preview only' })) -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

# ── Connectivity check ────────────────────────────────────────────────────────
Write-Host "  Testing SMB (port 445) connectivity to $hostName..." -ForegroundColor DarkGray
$connectTest = Test-NetConnection -ComputerName $hostName -Port 445 -WarningAction SilentlyContinue

if (-not $connectTest.TcpTestSucceeded) {
    Write-Host "  [ERROR] Cannot reach $hostName on port 445." -ForegroundColor Red
    Write-Host "  This network/ISP may be blocking outbound SMB (445). Use an Azure P2S/S2S VPN" -ForegroundColor Yellow
    Write-Host "  or ExpressRoute to tunnel SMB traffic over a different port instead." -ForegroundColor Yellow
    exit 1
}
Write-Host "  [OK]   Port 445 reachable." -ForegroundColor Green
Write-Host ""

if (-not $Apply) {
    Write-Host "  Would save credentials for '$hostName' and mount $sharePath as ${DriveLetter}:." -ForegroundColor Yellow
    Write-Host "  Re-run with -Apply to perform the mount." -ForegroundColor Yellow
    exit 0
}

if (-not $PSCmdlet.ShouldProcess($sharePath, "Mount as ${DriveLetter}:")) { exit 0 }

# ── Save credential and mount ──────────────────────────────────────────────────
try {
    cmd.exe /C "cmdkey /add:`"$hostName`" /user:`"localhost\$StorageAccountName`" /pass:`"$StorageAccountKey`"" | Out-Null
    Write-Host "  [OK]   Credential saved for $hostName." -ForegroundColor Green

    $psDriveParams = @{ Name = $DriveLetter; PSProvider = 'FileSystem'; Root = $sharePath; ErrorAction = 'Stop' }
    if ($Persist) { $psDriveParams['Persist'] = $true }
    New-PSDrive @psDriveParams | Out-Null
    Write-Host "  [OK]   Mounted $sharePath as ${DriveLetter}:" -ForegroundColor Green
} catch {
    Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
