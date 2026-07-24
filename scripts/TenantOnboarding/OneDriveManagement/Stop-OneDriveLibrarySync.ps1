#Requires -Version 5.1
<#
.SYNOPSIS
    Stop syncing one specific SharePoint/OneDrive library, without touching any other synced libraries.

.DESCRIPTION
    OneDrive doesn't offer a supported "stop syncing just this one library" command —
    doing it cleanly means shutting OneDrive down, removing that library's client
    policy file and its registry mount-point cache entry, editing the internal sync
    database ini file to drop just that library's entry, and removing its local
    folder, before relaunching OneDrive. This script automates that sequence for a
    single library identified by a distinctive text pattern in its policy file
    (originally used to stop syncing one customer/project's SharePoint library on a
    shared machine without disrupting the technician's other synced libraries).

.PARAMETER MatchPattern
    A distinctive string that appears in the target library's ClientPolicy_*.ini file
    — e.g. part of the SharePoint site name or library GUID. Used to identify which
    of the (possibly many) synced libraries to stop.

.PARAMETER LocalFolderPath
    Full path to the local synced folder to remove once sync is stopped, e.g.
    "$env:USERPROFILE\Contoso\ProjectX - Documents".

.PARAMETER Apply
    Actually stop OneDrive, edit its sync state, and remove the local folder. Without
    this switch, the script only reports what it would do.

.EXAMPLE
    .\Stop-OneDriveLibrarySync.ps1 -MatchPattern "ProjectX" -LocalFolderPath "$env:USERPROFILE\Contoso\ProjectX - Documents"

.EXAMPLE
    .\Stop-OneDriveLibrarySync.ps1 -MatchPattern "ProjectX" -LocalFolderPath "$env:USERPROFILE\Contoso\ProjectX - Documents" -Apply

.NOTES
    Operates on OneDrive's internal state files under
    %LOCALAPPDATA%\Microsoft\OneDrive\settings\Business1 — these are undocumented and
    may change format between OneDrive client versions. Test on one machine before
    wider rollout.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string] $MatchPattern,

    [Parameter(Mandatory)]
    [string] $LocalFolderPath,

    [switch] $Apply
)

$settingsRoot = "$env:USERPROFILE\AppData\Local\Microsoft\OneDrive\settings\Business1"

Write-Host ""
Write-Host "  Stop-OneDriveLibrarySync : '$MatchPattern'" -ForegroundColor Cyan
Write-Host "  Mode : $(if ($Apply) { 'Apply' } else { 'Preview only' })" -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

$policyFile = Get-ChildItem -Path $settingsRoot -Filter 'ClientPolicy_*.ini' -ErrorAction SilentlyContinue |
    Where-Object { (Get-Content -Path $_.FullName -Encoding Unicode -ErrorAction SilentlyContinue) -match [regex]::Escape($MatchPattern) } |
    Select-Object -First 1

if (-not $policyFile) {
    Write-Host "  [WARN] No ClientPolicy_*.ini file matching '$MatchPattern' found under $settingsRoot." -ForegroundColor Yellow
    exit 0
}

Write-Host "  Matched policy file : $($policyFile.Name)"
Write-Host "  Local folder        : $LocalFolderPath"
Write-Host ""

if (-not $Apply) {
    Write-Host "  Would shut down OneDrive, remove '$($policyFile.Name)', clear its mount-point cache entry," -ForegroundColor Yellow
    Write-Host "  remove its line from the sync database ini, delete '$LocalFolderPath', and relaunch OneDrive." -ForegroundColor Yellow
    Write-Host "  Re-run with -Apply to perform these actions." -ForegroundColor Yellow
    Write-Host ""
    exit 0
}

if (-not $PSCmdlet.ShouldProcess($LocalFolderPath, "Stop OneDrive sync for this library")) { exit 0 }

try {
    Start-Process "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe" -ArgumentList '/shutdown'
    Start-Sleep -Milliseconds 500

    Remove-Item $policyFile.FullName -Force
    Write-Host "  [OK]   Removed policy file." -ForegroundColor Green

    $key = Get-Item 'HKCU:\Software\Microsoft\OneDrive\Accounts\Business1\ScopeIdToMountPointPathCache' -ErrorAction SilentlyContinue
    if ($key) {
        foreach ($property in $key.Property) {
            if ($key.GetValue($property) -eq $LocalFolderPath) {
                Remove-ItemProperty -Path $key.PSPath -Name $property
                Write-Host "  [OK]   Removed mount-point cache entry." -ForegroundColor Green
                break
            }
        }
    }

    $syncDbFile = Get-ChildItem -Path $settingsRoot -Filter '????????-????-????-????-????????????*.ini' -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($syncDbFile) {
        $syncId = $null
        foreach ($line in (Get-Content $syncDbFile.FullName -Encoding Unicode)) {
            if ($line -match " = \d+ ([a-f0-9+]+) .*$([regex]::Escape($MatchPattern))") {
                $syncId = $Matches[1].Replace('+', '\+')
                break
            }
        }
        if ($syncId) {
            $tempFile = "$($syncDbFile.FullName)2"
            Get-Content $syncDbFile.FullName -Encoding Unicode | Where-Object { $_ -notmatch $syncId } | Set-Content $tempFile -Encoding Unicode
            Remove-Item $syncDbFile.FullName
            Rename-Item $tempFile $syncDbFile.Name
            Write-Host "  [OK]   Removed sync database entry." -ForegroundColor Green
        }
    }

    if (Test-Path $LocalFolderPath) {
        Remove-Item $LocalFolderPath -Recurse -Force
        Write-Host "  [OK]   Removed local folder." -ForegroundColor Green
    }

    Start-Process "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe" -ArgumentList '/background'
    Write-Host "  [OK]   OneDrive relaunched." -ForegroundColor Green
} catch {
    Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
