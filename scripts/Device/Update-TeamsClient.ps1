#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Remove the installed Microsoft Teams client and reinstall the latest new Teams,
    including the Teams Meeting Add-in for Outlook. Supports -WhatIf.

.DESCRIPTION
    Endpoint/AVD script that performs a clean reinstall of new Teams:

      1. Detects an installed MSTeams AppX package (aborts if none is present).
      2. Uninstalls the Microsoft Teams Meeting Add-in (MSI) if registered.
      3. Removes the MSTeams AppX package for all users.
      4. Downloads teamsbootstrapper.exe to the working folder.
      5. Provisions new Teams for all users (teamsbootstrapper.exe -p).
      6. Installs the Teams Meeting Add-in MSI shipped inside the new Teams package.
      7. Verifies the add-in registration and the provisioned AppX package.

    Every state-changing step is wrapped in ShouldProcess, so -WhatIf walks the whole
    flow and reports exactly what would be uninstalled, downloaded, installed and
    provisioned without touching the machine. Steps that can only be evaluated after
    a real install (new Teams version, add-in MSI path, final verification) are
    reported as such under -WhatIf instead of failing the run.

.PARAMETER WorkingDir
    Folder used for the bootstrapper download (default: C:\IT\AVD\Teams).

.PARAMETER BootstrapperUrl
    Download URL for teamsbootstrapper.exe (default: the Microsoft fwlink for new Teams).

.PARAMETER SkipMeetingAddIn
    Leave the Teams Meeting Add-in alone - do not uninstall it up front and do not
    install it afterwards. Only the Teams client itself is replaced.

.PARAMETER Force
    Continue even when no existing MSTeams package is detected (clean install instead
    of a reinstall).

.EXAMPLE
    # Dry run - show every uninstall/download/install this would perform
    .\Update-TeamsClient.ps1 -WhatIf

.EXAMPLE
    # Actually reinstall Teams plus the Outlook meeting add-in
    .\Update-TeamsClient.ps1

.EXAMPLE
    # Install new Teams on a device that has no Teams at all yet
    .\Update-TeamsClient.ps1 -Force

.EXAMPLE
    # Replace only the client, keep the currently installed meeting add-in
    .\Update-TeamsClient.ps1 -SkipMeetingAddIn

.NOTES
    Author  : Sjoerd Kanon
    Platform: Windows only (AppX + MSI), run as administrator
#>
[CmdletBinding(SupportsShouldProcess)]
param (
    [string] $WorkingDir      = 'C:\IT\AVD\Teams',
    [string] $BootstrapperUrl = 'https://go.microsoft.com/fwlink/?linkid=2243204&clcid=0x409',
    [switch] $SkipMeetingAddIn,
    [switch] $Force
)

Set-StrictMode -Version Latest

$simulate = [bool] $WhatIfPreference
$teamsExe = 'teamsbootstrapper.exe'
$exePath  = Join-Path $WorkingDir $teamsExe

# The add-in MSI installs 32-bit, so on x64 its uninstall entry lands in the
# WOW6432Node hive - checking only the 64-bit hive misses it.
$UninstallRoots = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
)

function Write-Step { param([string] $Message) Write-Host "  $Message" -ForegroundColor Cyan }
function Write-Ok   { param([string] $Message) Write-Host "  [ OK ] $Message" -ForegroundColor Green }
function Write-Skip { param([string] $Message) Write-Host "  [SKIP] $Message" -ForegroundColor DarkGray }
function Write-Bad  { param([string] $Message) Write-Host "  [FAIL] $Message" -ForegroundColor Red }

function Get-TeamsMeetingAddInEntry {
    <# Uninstall entries for the Teams Meeting Add-in, from both registry views. #>
    foreach ($root in $UninstallRoots) {
        if (-not (Test-Path $root)) { continue }
        foreach ($key in Get-ChildItem $root) {
            $name = $key.GetValue('DisplayName')
            if ($name -like '*Microsoft Teams Meeting Add-in*') {
                [PSCustomObject]@{
                    ProductCode = $key.PSChildName
                    DisplayName = $name
                    Version     = $key.GetValue('DisplayVersion')
                }
            }
        }
    }
}

Write-Host ''
Write-Host "  Mode: $(if ($simulate) { '-WhatIf - nothing will be changed' } else { 'APPLY - Teams will be reinstalled' })" -ForegroundColor Cyan
Write-Host ''

# -- 1. Detect the current Teams installation ---------------------------------
Write-Step '1. Detecting installed Teams client'
$existingTeams = @(Get-AppxPackage -AllUsers -Name '*MSTEAMS*' -ErrorAction SilentlyContinue)

if ($existingTeams.Count -gt 0) {
    foreach ($pkg in $existingTeams) { Write-Ok "Found $($pkg.Name) $($pkg.Version)" }
} elseif ($Force) {
    Write-Skip 'No Teams installation detected - continuing because -Force was given'
} else {
    Write-Bad 'No Teams installation detected. Use -Force to install new Teams anyway.'
    exit 1
}

# -- 2. Uninstall the Teams Meeting Add-in ------------------------------------
Write-Host ''
Write-Step '2. Teams Meeting Add-in (uninstall)'
if ($SkipMeetingAddIn) {
    Write-Skip 'Skipped (-SkipMeetingAddIn)'
} else {
    $addInEntries = @(Get-TeamsMeetingAddInEntry)
    if ($addInEntries.Count -eq 0) {
        Write-Skip 'Microsoft Teams Meeting Add-in is not installed'
    }
    foreach ($entry in $addInEntries) {
        $target = "$($entry.DisplayName) $($entry.Version) [$($entry.ProductCode)]"
        if ($PSCmdlet.ShouldProcess($target, 'msiexec /x /qn (uninstall)')) {
            $proc = Start-Process msiexec.exe -ArgumentList "/x $($entry.ProductCode) /qn" -Wait -PassThru
            if ($proc.ExitCode -eq 0) { Write-Ok "Uninstalled $($entry.DisplayName)" }
            else { Write-Bad "Uninstall of $($entry.DisplayName) returned exit code $($proc.ExitCode)" }
        }
    }
}

# -- 3. Remove the Teams AppX package -----------------------------------------
Write-Host ''
Write-Step '3. Teams AppX package (remove for all users)'
if ($existingTeams.Count -eq 0) { Write-Skip 'Nothing to remove' }
foreach ($pkg in $existingTeams) {
    if ($PSCmdlet.ShouldProcess("$($pkg.Name) $($pkg.Version)", 'Remove-AppxPackage -AllUsers')) {
        $pkg | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
        if (Get-AppxPackage -AllUsers -Name $pkg.Name -ErrorAction SilentlyContinue) {
            Write-Bad "$($pkg.Name) is still present after removal"
        } else {
            Write-Ok "Removed $($pkg.Name)"
        }
    }
}

# -- 4. Download the bootstrapper ---------------------------------------------
Write-Host ''
Write-Step '4. Download new Teams bootstrapper'
if ($PSCmdlet.ShouldProcess($WorkingDir, 'Create working directory')) {
    New-Item -ItemType Directory -Path $WorkingDir -Force | Out-Null
    Write-Ok "Working directory ready: $WorkingDir"
}
if ($PSCmdlet.ShouldProcess($exePath, "Download $BootstrapperUrl")) {
    try {
        $progressBackup = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'
        Invoke-WebRequest -Uri $BootstrapperUrl -OutFile $exePath -UseBasicParsing -ErrorAction Stop
        $ProgressPreference = $progressBackup
        Write-Ok "Downloaded $teamsExe ($([math]::Round((Get-Item $exePath).Length / 1MB, 1)) MB)"
    } catch {
        Write-Bad "Download failed: $($_.Exception.Message)"
        exit 1
    }
}

# -- 5. Install / provision new Teams -----------------------------------------
Write-Host ''
Write-Step '5. Install new Teams'
if ($PSCmdlet.ShouldProcess($exePath, 'Provision new Teams for all users (-p)')) {
    if (-not (Test-Path $exePath)) {
        Write-Bad "Bootstrapper not found at $exePath"
        exit 1
    }
    $proc = Start-Process $exePath -ArgumentList '-p' -Wait -PassThru
    if ($proc.ExitCode -eq 0) { Write-Ok 'Bootstrapper completed' }
    else { Write-Bad "Bootstrapper returned exit code $($proc.ExitCode)" }
}

# -- 6. Install the Teams Meeting Add-in for all users ------------------------
Write-Host ''
Write-Step '6. Teams Meeting Add-in (install)'
if ($SkipMeetingAddIn) {
    Write-Skip 'Skipped (-SkipMeetingAddIn)'
} else {
    $newTeams        = Get-AppxPackage -Name 'MSTeams' -ErrorAction SilentlyContinue
    $newTeamsVersion = if ($newTeams) { $newTeams.Version } else { $null }

    if (-not $newTeamsVersion) {
        if ($simulate) {
            Write-Skip 'New Teams is not installed yet - on a real run the add-in MSI comes from the freshly installed package folder'
            Write-Skip 'Would run: msiexec.exe /i "<ProgramFiles>\WindowsApps\MSTeams_<version>_x64__8wekyb3d8bbwe\MicrosoftTeamsMeetingAddinInstaller.msi" TARGETDIR="<ProgramFiles(x86)>\Microsoft\TeamsMeetingAddin\<version>\" /qn ALLUSERS=1'
        } else {
            Write-Bad 'New Teams package not found. Install new Teams from https://aka.ms/GetTeams first.'
            exit 1
        }
    } else {
        Write-Ok "Found new Teams version: $newTeamsVersion"

        $tmaPath    = '{0}\WindowsApps\MSTeams_{1}_x64__8wekyb3d8bbwe\MicrosoftTeamsMeetingAddinInstaller.msi' -f $env:ProgramFiles, $newTeamsVersion
        $publisher  = Get-AppLockerFileInformation -Path $tmaPath -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Publisher
        $tmaVersion = if ($publisher) { $publisher.BinaryVersion } else { $null }

        if (-not $tmaVersion) {
            Write-Bad "Teams Meeting Add-in installer not found at $tmaPath"
            if (-not $simulate) { exit 1 }
        } else {
            Write-Ok "Found Teams Meeting Add-in version: $tmaVersion"
            $targetDir = '{0}\Microsoft\TeamsMeetingAddin\{1}\' -f ${env:ProgramFiles(x86)}, $tmaVersion
            $params    = '/i "{0}" TARGETDIR="{1}" /qn ALLUSERS=1' -f $tmaPath, $targetDir

            if ($PSCmdlet.ShouldProcess("Teams Meeting Add-in $tmaVersion", "msiexec.exe $params")) {
                $proc = Start-Process msiexec.exe -ArgumentList $params -Wait -PassThru
                if ($proc.ExitCode -eq 0) { Write-Ok "Installed Teams Meeting Add-in to $targetDir" }
                else { Write-Bad "msiexec returned exit code $($proc.ExitCode)" }
            }
        }
    }
}

# -- 7. Verify ----------------------------------------------------------------
Write-Host ''
Write-Step '7. Verification'
if ($simulate) {
    Write-Skip 'Skipped - nothing was changed, so there is nothing to verify (-WhatIf)'
    Write-Host ''
    Write-Host '  Dry run only - rerun without -WhatIf to apply these changes.' -ForegroundColor Yellow
    Write-Host ''
    exit 0
}

$failed = $false

if (-not $SkipMeetingAddIn) {
    if (Get-TeamsMeetingAddInEntry) { Write-Ok 'Teams Meeting Add-in installed' }
    else { Write-Bad 'Teams Meeting Add-in installation failed'; $failed = $true }
}

if (Get-ProvisionedAppPackage -Online | Where-Object { $_.DisplayName -eq 'MSTeams' }) {
    Write-Ok 'Teams provisioned for all users'
} else {
    Write-Bad 'Teams installation failed'
    $failed = $true
}

Write-Host ''
if ($failed) { exit 1 }
exit 0
