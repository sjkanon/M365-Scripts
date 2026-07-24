#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Redirect Desktop/Documents/Pictures/Downloads to the OneDrive for Business sync folder.

.DESCRIPTION
    OneDrive's built-in Known Folder Move (KFM) policy only covers Desktop/Documents/
    Pictures — this script extends the same technique (SHSetKnownFolderPath +
    Robocopy migration) to also redirect Downloads and LocalAppData if desired, and
    works standalone (e.g. for a manual/logon-script rollout instead of the Intune
    KFM policy). Existing file/folder contents are moved with Robocopy (auditable,
    resumable) rather than deleted; the original folder is hidden, not removed.

    Based on the widely-used community "OneDrive KFM/redirect" script pattern
    (SHSetKnownFolderPath via P/Invoke), generalized here to avoid any hardcoded
    environment specifics.

    Also optionally sets OneDrive's "Timerautomount" flag, so previously-synced
    libraries the signed-in account has access to auto-mount on a new device instead
    of requiring the user to manually reconnect each one.

.PARAMETER Folders
    Which known folders to redirect into OneDrive. Default: Desktop, Documents,
    Pictures, Downloads.

.PARAMETER EnableAutoMountSharedLibraries
    Also set the OneDrive "Timerautomount" registry flag for the current user, so
    previously-synced SharePoint/OneDrive libraries auto-mount on this device.

.PARAMETER Apply
    Actually redirect folders / move data. Without this switch, the script only
    reports what it would do.

.EXAMPLE
    .\Set-OneDriveKnownFolderRedirect.ps1

.EXAMPLE
    .\Set-OneDriveKnownFolderRedirect.ps1 -Apply

.EXAMPLE
    .\Set-OneDriveKnownFolderRedirect.ps1 -Folders Desktop,Documents -EnableAutoMountSharedLibraries -Apply

.NOTES
    Run in the signed-in user's context (not SYSTEM) — it reads that user's OneDrive
    Business1 account registry key to find the sync root.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [ValidateSet('Desktop', 'Documents', 'Pictures', 'Downloads', 'LocalAppData')]
    [string[]] $Folders = @('Desktop', 'Documents', 'Pictures', 'Downloads'),

    [switch] $EnableAutoMountSharedLibraries,
    [switch] $Apply
)

function Set-KnownFolderPath {
    param(
        [Parameter(Mandatory)] [string] $KnownFolder,
        [Parameter(Mandatory)] [string] $Path
    )
    $knownFolderGuids = @{
        'Desktop'      = 'B4BFCC3A-DB2C-424C-B029-7FE99A87C641'
        'Documents'    = 'FDD39AD0-238F-46AF-ADB4-6C85480369C7'
        'Pictures'     = '33E28130-4E1E-4676-835A-98395C3BC3BB'
        'Downloads'    = '374DE290-123F-4565-9164-39C4925E467B'
        'LocalAppData' = 'F1B32785-6FBA-4FCF-9D55-7B8E7F157091'
    }
    $type = ([System.Management.Automation.PSTypeName]'TenantOnboarding.KnownFolders').Type
    if (-not $type) {
        $signature = '[DllImport("shell32.dll")] public static extern int SHSetKnownFolderPath(ref Guid folderId, uint flags, IntPtr token, [MarshalAs(UnmanagedType.LPWStr)] string path);'
        $type = Add-Type -MemberDefinition $signature -Name 'KnownFolders' -Namespace 'TenantOnboarding' -PassThru
    }
    if (-not (Test-Path $Path)) { New-Item -Path $Path -ItemType Directory -Force | Out-Null }
    $guid = [guid]$knownFolderGuids[$KnownFolder]
    $result = $type::SHSetKnownFolderPath([ref]$guid, 0, 0, $Path)
    if ($result -ne 0) { throw "Error redirecting $KnownFolder. Return code $result." }
    attrib +r $Path
}

function Move-RedirectedFolderContents {
    param([string] $Source, [string] $Destination, [string] $Log)
    if (-not (Test-Path (Split-Path $Log))) { New-Item -Path (Split-Path $Log) -ItemType Directory -Force | Out-Null }
    Robocopy.exe $Source $Destination /E /MOV /XJ /R:1 /W:1 /NP "/LOG+:$Log" | Out-Null
}

$folderMap = @{
    'Desktop'      = @{ Shell = 'Desktop';       Target = 'Desktop' }
    'Documents'    = @{ Shell = 'MyDocuments';    Target = 'Documents' }
    'Pictures'     = @{ Shell = 'MyPictures';     Target = 'Pictures' }
    'Downloads'    = @{ Shell = $null;            Target = 'Downloads' }
    'LocalAppData' = @{ Shell = 'LocalApplicationData'; Target = 'LocalAppData' }
}

Write-Host ""
Write-Host "  Set-OneDriveKnownFolderRedirect" -ForegroundColor Cyan
Write-Host "  Mode : $(if ($Apply) { 'Apply' } else { 'Preview only' })" -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ""

$oneDriveFolder = (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\OneDrive\Accounts\Business1' -Name UserFolder -ErrorAction SilentlyContinue).UserFolder
if (-not $oneDriveFolder -or -not (Test-Path $oneDriveFolder)) {
    Write-Host "  [ERROR] Could not resolve the signed-in user's OneDrive for Business sync folder. Is OneDrive signed in?" -ForegroundColor Red
    exit 1
}
Write-Host "  OneDrive folder : $oneDriveFolder"
Write-Host ""

foreach ($folder in $Folders) {
    $target = Join-Path $oneDriveFolder $folderMap[$folder].Target

    if (-not $Apply) {
        Write-Host "  [PREVIEW] Would redirect '$folder' -> $target" -ForegroundColor Yellow
        continue
    }
    if (-not $PSCmdlet.ShouldProcess($folder, "Redirect into OneDrive")) { continue }

    try {
        if ($folder -eq 'Downloads') {
            # Downloads has no legacy [Environment] mapping usable with SHSetKnownFolderPath the same way; use its own GUID directly.
            Set-KnownFolderPath -KnownFolder 'Downloads' -Path $target
        } else {
            $currentPath = [Environment]::GetFolderPath($folderMap[$folder].Shell)
            if ($currentPath -ne $target) {
                Set-KnownFolderPath -KnownFolder $folder -Path $target
                Move-RedirectedFolderContents -Source $currentPath -Destination $target -Log "$env:LocalAppData\RedirectLogs\Robocopy$folder.log"
                attrib +h $currentPath
            }
        }
        Write-Host "  [OK]   '$folder' redirected to $target" -ForegroundColor Green
    } catch {
        Write-Host "  [WARN] Failed to redirect '$folder': $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

if ($EnableAutoMountSharedLibraries) {
    if (-not $Apply) {
        Write-Host "  [PREVIEW] Would set OneDrive Timerautomount = 1 for the current user." -ForegroundColor Yellow
    } elseif ($PSCmdlet.ShouldProcess('OneDrive Timerautomount', 'Enable')) {
        New-ItemProperty -Path 'HKCU:\Software\Microsoft\OneDrive\Accounts\Business1' -Name 'Timerautomount' -Value 1 -PropertyType QWord -Force | Out-Null
        Write-Host "  [OK]   Timerautomount enabled." -ForegroundColor Green
    }
}

Write-Host ""
