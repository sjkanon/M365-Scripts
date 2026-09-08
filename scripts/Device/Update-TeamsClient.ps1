#Requires -Version 5.1
<#
.SYNOPSIS
    Remove the installed Microsoft Teams client and reinstall the latest new Teams,
    including the Teams Meeting Add-in for Outlook. Supports -WhatIf.

.DESCRIPTION
    Endpoint/AVD script that performs a clean reinstall of new Teams:

      1. Preflight  - detect the installed MSTeams AppX package, the meeting add-in
                      and any running Teams/Outlook process.
      2. Download   - fetch teamsbootstrapper.exe and verify its Microsoft signature
                      BEFORE anything is uninstalled, so a failed download can never
                      leave the device without a Teams client.
      3. Uninstall  - Teams Meeting Add-in (MSI), the MSTeams AppX package for all
                      users, and the provisioned package.
      4. Install    - provision new Teams for all users (teamsbootstrapper.exe -p).
      5. Add-in     - install the Teams Meeting Add-in MSI shipped inside the new
                      Teams package (ALLUSERS=1).
      6. Verify     - re-check the add-in registration and the provisioned package.

    Every state-changing step is wrapped in ShouldProcess, so -WhatIf walks the whole
    flow and reports exactly what would be uninstalled, downloaded, installed and
    provisioned without touching the machine. Steps that can only be evaluated after
    a real install (new Teams version, add-in MSI path, final verification) are
    reported as such under -WhatIf instead of failing the run.

    Safety
    ------
      - The installer is downloaded and signature-checked before the first uninstall.
      - Every msiexec/bootstrapper call runs with a timeout and is killed if it hangs,
        so an RMM job can never block the agent indefinitely.
      - MSI exit code 1618 (another install in progress) is retried; 3010 is treated
        as success with a reboot flagged in the summary.
      - Unexpected errors abort the run instead of continuing half-way, and the exit
        code is 0 on success (a -WhatIf run included) or 1 on failure.
      - An apply run writes a transcript to the log folder for after-the-fact review.

    Running it by hand
    ------------------
    From an ordinary PowerShell window the script elevates itself (UAC) and continues
    in a new elevated window that stays open, so no "run as administrator" dance is
    needed first. An interactive apply run asks for confirmation once before the first
    uninstall; -Confirm:$false skips that question. A -WhatIf run never asks.

        powershell -ExecutionPolicy Bypass -File .\Update-TeamsClient.ps1 -WhatIf

    RMM / NinjaOne
    --------------
    The script is safe to deploy from NinjaOne (run as System):

      - Script variables arrive as environment variables, so a checkbox named
        whatIf, force or skipMeetingAddIn, or a text field named workingDir, is
        picked up when the matching parameter is not passed on the command line.
        A Parameters field of -WhatIf works just as well.
      - If the agent starts PowerShell 32-bit, the script relaunches itself 64-bit
        via SysNative first. Without that, registry reads are redirected to
        WOW6432Node and $env:ProgramFiles points at the x86 folder, so the AppX
        package and the add-in MSI are never found.
      - Add -Confirm:$false to the Parameters field so the confirmation question can
        never appear, whatever the agent reports about the session.

.PARAMETER WorkingDir
    Folder used for the bootstrapper download (default: C:\IT\AVD\Teams).

.PARAMETER LogPath
    Folder for the transcript of an apply run (default: C:\Temp). A -WhatIf run
    writes no transcript.

.PARAMETER BootstrapperUrl
    Download URL for teamsbootstrapper.exe (default: the Microsoft fwlink for new
    Teams). Must be https.

.PARAMETER SkipMeetingAddIn
    Leave the Teams Meeting Add-in alone - do not uninstall it up front and do not
    install it afterwards. Only the Teams client itself is replaced.

.PARAMETER SkipSignatureCheck
    Accept the downloaded bootstrapper without verifying its Authenticode signature.
    Only for an air-gapped or internally hosted -BootstrapperUrl that is not signed
    by Microsoft.

.PARAMETER TimeoutSeconds
    Per-process timeout for msiexec and the bootstrapper (default: 900). A process
    that outlives it is killed and the step is reported as failed.

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

.EXAMPLE
    # NinjaOne: preview run driven by a script variable instead of a parameter
    $env:whatIf = 'true'; .\Update-TeamsClient.ps1

.NOTES
    Author  : Sjoerd Kanon
    Platform: Windows only (AppX + MSI), run as administrator or as System
#>
[CmdletBinding(SupportsShouldProcess)]
param (
    [string] $WorkingDir      = 'C:\IT\AVD\Teams',
    [string] $LogPath         = 'C:\Temp',
    [string] $BootstrapperUrl = 'https://go.microsoft.com/fwlink/?linkid=2243204&clcid=0x409',
    [switch] $SkipMeetingAddIn,
    [switch] $SkipSignatureCheck,
    [ValidateRange(60, 7200)]
    [int]    $TimeoutSeconds  = 900,
    [switch] $Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-ForwardedArgument {
    <# Rebuild the caller's own parameters as a command line for a relaunch. #>
    param([Parameter(Mandatory)] $Bound)

    $list = @()
    foreach ($entry in $Bound.GetEnumerator()) {
        if ($entry.Value -is [switch] -or $entry.Value -is [bool]) {
            # -Confirm:$false and friends must survive as an explicit :$false.
            if ($entry.Value) { $list += "-$($entry.Key)" } else { $list += "-$($entry.Key):`$false" }
        } else {
            $list += "-$($entry.Key)"
            $list += [string] $entry.Value
        }
    }
    return $list
}

function Test-Elevated {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    return ([Security.Principal.WindowsPrincipal] $identity).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# -- RMM: run 64-bit -----------------------------------------------------------
# An RMM agent (NinjaOne among them) can start PowerShell 32-bit. Under WOW64 the
# HKLM reads land in WOW6432Node and $env:ProgramFiles points at the x86 folder,
# which hides both the AppX package and the add-in MSI. Relaunch ourselves 64-bit
# and hand the parameters over unchanged.
if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
    $nativeShell = Join-Path $env:WINDIR 'SysNative\WindowsPowerShell\v1.0\powershell.exe'
    if (Test-Path $nativeShell) {
        $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath) +
                   (Get-ForwardedArgument -Bound $PSBoundParameters)
        & $nativeShell @argList
        exit $LASTEXITCODE
    }
    Write-Warning 'Running 32-bit and SysNative is unavailable - AppX and registry lookups may fail.'
}

# -- Elevation -----------------------------------------------------------------
# AppX enumeration, the MSI calls and the bootstrapper all need administrator rights.
# Run as System from an RMM this is already true; started by hand it is usually not,
# so ask for elevation instead of failing on a #Requires line.
if (-not (Test-Elevated)) {
    if (-not [Environment]::UserInteractive) {
        Write-Error 'Administrator rights are required. Run the script as System or from an elevated session.'
        exit 1
    }

    Write-Host ''
    Write-Host '  Not running elevated - asking for administrator rights...' -ForegroundColor Yellow
    $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-NoExit', '-File', $PSCommandPath) +
               (Get-ForwardedArgument -Bound $PSBoundParameters)
    try {
        Start-Process -FilePath (Get-Process -Id $PID).Path -ArgumentList $argList -Verb RunAs | Out-Null
        Write-Host '  Continued in an elevated window.' -ForegroundColor Cyan
        exit 0
    } catch {
        Write-Error "Elevation was declined or failed: $($_.Exception.Message)"
        exit 1
    }
}

# -- RMM: script variables -----------------------------------------------------
# NinjaOne exposes script variables as environment variables. Honour them only
# when the matching parameter was not passed on the command line, so a plain
# .\Update-TeamsClient.ps1 -WhatIf keeps working exactly as before.
$rmmTrue = @('true', '1', 'yes')
if (-not $PSBoundParameters.ContainsKey('WhatIf')             -and $env:whatIf             -in $rmmTrue) { $WhatIfPreference   = $true }
if (-not $PSBoundParameters.ContainsKey('Force')              -and $env:force              -in $rmmTrue) { $Force              = $true }
if (-not $PSBoundParameters.ContainsKey('SkipMeetingAddIn')   -and $env:skipMeetingAddIn   -in $rmmTrue) { $SkipMeetingAddIn   = $true }
if (-not $PSBoundParameters.ContainsKey('SkipSignatureCheck') -and $env:skipSignatureCheck -in $rmmTrue) { $SkipSignatureCheck = $true }
if (-not $PSBoundParameters.ContainsKey('WorkingDir')         -and $env:workingDir)                      { $WorkingDir         = $env:workingDir }
if (-not $PSBoundParameters.ContainsKey('LogPath')            -and $env:logPath)                         { $LogPath            = $env:logPath }

$simulate       = [bool] $WhatIfPreference
$teamsExe       = 'teamsbootstrapper.exe'
$exePath        = Join-Path $WorkingDir $teamsExe
$exitCode       = 0
$rebootRequired = $false
$transcribing   = $false
$cancelled      = $false

# -Confirm:$false means "never ask", for an unattended run from a scheduler or RMM.
$confirmSuppressed = $PSBoundParameters.ContainsKey('Confirm') -and -not $PSBoundParameters['Confirm']

# The add-in MSI installs 32-bit, so on x64 its uninstall entry lands in the
# WOW6432Node hive - checking only the 64-bit hive misses it.
$UninstallRoots = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
)

function Write-Step { param([string] $Message) Write-Host "  $Message" -ForegroundColor Cyan }
function Write-Ok   { param([string] $Message) Write-Host "  [ OK ] $Message" -ForegroundColor Green }
function Write-Skip { param([string] $Message) Write-Host "  [SKIP] $Message" -ForegroundColor DarkGray }
function Write-Warn { param([string] $Message) Write-Host "  [WARN] $Message" -ForegroundColor Yellow }
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

function Get-MsiProductVersion {
    <#
        ProductVersion straight from the MSI property table. Microsoft's own sample
        reads it with Get-AppLockerFileInformation, but that module is absent on some
        editions and under PowerShell 7 it drags in the Windows PowerShell
        compatibility layer, which fails and floods a -WhatIf run with its own output.
    #>
    param([Parameter(Mandatory)] [string] $Path)

    $installer = $null; $database = $null; $view = $null
    try {
        $installer = New-Object -ComObject WindowsInstaller.Installer
        $database  = $installer.GetType().InvokeMember('OpenDatabase', 'InvokeMethod', $null, $installer, @($Path, 0))
        $view      = $database.GetType().InvokeMember('OpenView', 'InvokeMethod', $null, $database, @("SELECT Value FROM Property WHERE Property = 'ProductVersion'"))
        $null      = $view.GetType().InvokeMember('Execute', 'InvokeMethod', $null, $view, @($null))
        $record    = $view.GetType().InvokeMember('Fetch', 'InvokeMethod', $null, $view, $null)
        if (-not $record) { return $null }
        return $record.GetType().InvokeMember('StringData', 'GetProperty', $null, $record, @(1))
    } catch {
        return $null
    } finally {
        foreach ($com in @($view, $database, $installer)) {
            if ($com) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($com) }
        }
    }
}

function Invoke-Installer {
    <#
        Run an installer and never hang the caller: the process is killed when it
        outlives the timeout, 1618 (another install in progress) is retried, and
        3010 counts as success with a reboot pending.
    #>
    param(
        [Parameter(Mandatory)] [string] $FilePath,
        [string] $Arguments = '',
        [int]    $Attempts  = 3
    )

    for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
        $startArgs = @{ FilePath = $FilePath; PassThru = $true; WindowStyle = 'Hidden' }
        if ($Arguments) { $startArgs['ArgumentList'] = $Arguments }

        $proc = Start-Process @startArgs
        if (-not $proc.WaitForExit($TimeoutSeconds * 1000)) {
            try { $proc.Kill() } catch { Write-Warn "Could not kill PID $($proc.Id): $($_.Exception.Message)" }
            return [PSCustomObject]@{ Success = $false; ExitCode = $null; RebootRequired = $false
                                      Message = "timed out after $TimeoutSeconds seconds and was killed" }
        }

        $code = $proc.ExitCode
        if ($code -eq 1618 -and $attempt -lt $Attempts) {
            Write-Warn "Another installation is in progress (1618) - retrying in 30 seconds ($attempt/$($Attempts - 1))"
            Start-Sleep -Seconds 30
            continue
        }

        return [PSCustomObject]@{
            Success        = ($code -in @(0, 3010))
            ExitCode       = $code
            RebootRequired = ($code -eq 3010)
            Message        = "exit code $code"
        }
    }
}

# -- Run -----------------------------------------------------------------------
try {
    if (-not $simulate) {
        if (-not (Test-Path $LogPath)) { New-Item -ItemType Directory -Path $LogPath -Force | Out-Null }
        $logFile = Join-Path $LogPath ("Update-TeamsClient_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
        Start-Transcript -Path $logFile | Out-Null
        $transcribing = $true
    }

    Write-Host ''
    Write-Host "  Mode: $(if ($simulate) { '-WhatIf - nothing will be changed' } else { 'APPLY - Teams will be reinstalled' })" -ForegroundColor Cyan
    Write-Host ''

    # -- 1. Preflight ----------------------------------------------------------
    Write-Step '1. Preflight'
    try {
        $existingTeams = @(Get-AppxPackage -AllUsers -Name '*MSTEAMS*' -ErrorAction SilentlyContinue)
    } catch {
        throw "Could not enumerate AppX packages ($($_.Exception.Message)). This needs administrator or System rights."
    }

    if ($existingTeams.Count -gt 0) {
        foreach ($pkg in $existingTeams) { Write-Ok "Found $($pkg.Name) $($pkg.Version)" }
    } elseif ($Force) {
        Write-Skip 'No Teams installation detected - continuing because -Force was given'
    } else {
        throw 'No Teams installation detected. Use -Force to install new Teams anyway.'
    }

    if (Get-Process -Name 'ms-teams', 'Teams' -ErrorAction SilentlyContinue) {
        Write-Warn 'Teams is running - it will be closed by the removal; warn the user before an apply run'
    }
    if (-not $SkipMeetingAddIn -and (Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue)) {
        Write-Warn 'Outlook is running - the meeting add-in registers reliably only after Outlook restarts'
    }

    # Last exit before anything is touched. Only for a hands-on run: an RMM or
    # scheduled session is not interactive, and -Confirm:$false silences it outright.
    if (-not $simulate -and -not $confirmSuppressed -and [Environment]::UserInteractive) {
        Write-Host ''
        Write-Warn 'Teams and the meeting add-in will be uninstalled and reinstalled on this device.'
        $answer = Read-Host '  Continue? [y/N]'
        if ($answer -notmatch '^[Yy]') {
            $cancelled = $true
            throw 'Cancelled - nothing was changed.'
        }
    }

    # -- 2. Download and verify the bootstrapper -------------------------------
    # Deliberately before any uninstall: a failed download must never leave the
    # device without a Teams client.
    Write-Host ''
    Write-Step '2. Download new Teams bootstrapper'
    if ($BootstrapperUrl -notmatch '^https://') {
        throw "BootstrapperUrl must be https: $BootstrapperUrl"
    }

    if ($PSCmdlet.ShouldProcess($WorkingDir, 'Create working directory')) {
        New-Item -ItemType Directory -Path $WorkingDir -Force | Out-Null
        Write-Ok "Working directory ready: $WorkingDir"
    }

    if ($PSCmdlet.ShouldProcess($exePath, "Download $BootstrapperUrl")) {
        # PowerShell 5.1 on older builds still defaults to TLS 1.0, which the CDN refuses.
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $progressBackup     = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'
        try {
            Invoke-WebRequest -Uri $BootstrapperUrl -OutFile $exePath -UseBasicParsing
        } catch {
            throw "Download failed: $($_.Exception.Message)"
        } finally {
            $ProgressPreference = $progressBackup
        }

        $downloaded = Get-Item $exePath
        if ($downloaded.Length -lt 100KB) {
            throw "Downloaded file is only $($downloaded.Length) bytes - not a valid bootstrapper"
        }
        Write-Ok "Downloaded $teamsExe ($([math]::Round($downloaded.Length / 1MB, 1)) MB)"

        if ($SkipSignatureCheck) {
            Write-Warn 'Signature check skipped (-SkipSignatureCheck)'
        } else {
            $signature = Get-AuthenticodeSignature -FilePath $exePath
            if ($signature.Status -ne 'Valid') {
                throw "Bootstrapper signature is $($signature.Status) - refusing to run it"
            }
            if ($signature.SignerCertificate.Subject -notmatch 'O=Microsoft Corporation') {
                throw "Bootstrapper is not signed by Microsoft: $($signature.SignerCertificate.Subject)"
            }
            Write-Ok 'Signature verified: Microsoft Corporation'
        }
    }

    # -- 3. Uninstall the Teams Meeting Add-in ---------------------------------
    Write-Host ''
    Write-Step '3. Teams Meeting Add-in (uninstall)'
    if ($SkipMeetingAddIn) {
        Write-Skip 'Skipped (-SkipMeetingAddIn)'
    } else {
        $addInEntries = @(Get-TeamsMeetingAddInEntry)
        if ($addInEntries.Count -eq 0) { Write-Skip 'Microsoft Teams Meeting Add-in is not installed' }

        foreach ($entry in $addInEntries) {
            $target = "$($entry.DisplayName) $($entry.Version) [$($entry.ProductCode)]"
            if ($PSCmdlet.ShouldProcess($target, 'msiexec /x /qn (uninstall)')) {
                $result = Invoke-Installer -FilePath 'msiexec.exe' -Arguments "/x $($entry.ProductCode) /qn /norestart"
                if ($result.Success) {
                    Write-Ok "Uninstalled $($entry.DisplayName)"
                    if ($result.RebootRequired) { $rebootRequired = $true }
                } else {
                    # Not fatal: the reinstall below replaces the add-in anyway.
                    Write-Warn "Uninstall of $($entry.DisplayName) failed ($($result.Message))"
                }
            }
        }
    }

    # -- 4. Remove the Teams AppX package --------------------------------------
    Write-Host ''
    Write-Step '4. Teams AppX package (remove for all users)'
    if ($existingTeams.Count -eq 0) { Write-Skip 'Nothing to remove' }

    foreach ($pkg in $existingTeams) {
        if ($PSCmdlet.ShouldProcess("$($pkg.Name) $($pkg.Version)", 'Remove-AppxPackage -AllUsers')) {
            $pkg | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
            if (Get-AppxPackage -AllUsers -Name $pkg.Name -ErrorAction SilentlyContinue) {
                Write-Warn "$($pkg.Name) is still present after removal - the reinstall will upgrade it in place"
            } else {
                Write-Ok "Removed $($pkg.Name)"
            }
        }
    }

    # Without dropping the provisioned copy, new user profiles keep getting the old
    # version staged from the image.
    $provisioned = @(Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -like 'MSTeams*' })
    if ($provisioned.Count -eq 0) {
        Write-Skip 'No provisioned MSTeams package to remove'
    }
    foreach ($prov in $provisioned) {
        if ($PSCmdlet.ShouldProcess($prov.PackageName, 'Remove-AppxProvisionedPackage -Online')) {
            try {
                Remove-AppxProvisionedPackage -Online -PackageName $prov.PackageName | Out-Null
                Write-Ok "Deprovisioned $($prov.PackageName)"
            } catch {
                Write-Warn "Could not deprovision $($prov.PackageName): $($_.Exception.Message)"
            }
        }
    }

    # -- 5. Install / provision new Teams --------------------------------------
    Write-Host ''
    Write-Step '5. Install new Teams'
    if ($PSCmdlet.ShouldProcess($exePath, 'Provision new Teams for all users (-p)')) {
        if (-not (Test-Path $exePath)) { throw "Bootstrapper not found at $exePath" }

        $result = Invoke-Installer -FilePath $exePath -Arguments '-p'
        if (-not $result.Success) { throw "Bootstrapper failed ($($result.Message))" }
        if ($result.RebootRequired) { $rebootRequired = $true }
        Write-Ok 'Bootstrapper completed'
    }

    # -- 6. Install the Teams Meeting Add-in for all users ---------------------
    Write-Host ''
    Write-Step '6. Teams Meeting Add-in (install)'
    if ($SkipMeetingAddIn) {
        Write-Skip 'Skipped (-SkipMeetingAddIn)'
    } else {
        $newTeams        = Get-AppxPackage -Name 'MSTeams' -ErrorAction SilentlyContinue
        $newTeamsVersion = if ($newTeams) { $newTeams.Version } else { $null }

        if (-not $newTeamsVersion -and $simulate) {
            Write-Skip 'New Teams is not installed yet - on a real run the add-in MSI comes from the freshly installed package folder'
            Write-Skip 'Would run: msiexec.exe /i "<ProgramFiles>\WindowsApps\MSTeams_<version>_x64__8wekyb3d8bbwe\MicrosoftTeamsMeetingAddinInstaller.msi" TARGETDIR="<ProgramFiles(x86)>\Microsoft\TeamsMeetingAddin\<version>\" /qn ALLUSERS=1'
        } elseif (-not $newTeamsVersion) {
            throw 'New Teams package not found after install. Check the bootstrapper output.'
        } else {
            Write-Ok "Found new Teams version: $newTeamsVersion"

            $tmaPath    = '{0}\WindowsApps\MSTeams_{1}_x64__8wekyb3d8bbwe\MicrosoftTeamsMeetingAddinInstaller.msi' -f $env:ProgramFiles, $newTeamsVersion
            $tmaVersion = if (Test-Path $tmaPath) { Get-MsiProductVersion -Path $tmaPath } else { $null }

            if (-not $tmaVersion -and $simulate) {
                Write-Skip "Add-in MSI not present in the current package - a real run takes it from $tmaPath"
            } elseif (-not $tmaVersion) {
                throw "Teams Meeting Add-in installer not found at $tmaPath"
            } else {
                Write-Ok "Found Teams Meeting Add-in version: $tmaVersion"
                $targetDir = '{0}\Microsoft\TeamsMeetingAddin\{1}\' -f ${env:ProgramFiles(x86)}, $tmaVersion
                $params    = '/i "{0}" TARGETDIR="{1}" /qn /norestart ALLUSERS=1' -f $tmaPath, $targetDir

                if ($PSCmdlet.ShouldProcess("Teams Meeting Add-in $tmaVersion", "msiexec.exe $params")) {
                    $result = Invoke-Installer -FilePath 'msiexec.exe' -Arguments $params
                    if (-not $result.Success) { throw "Teams Meeting Add-in install failed ($($result.Message))" }
                    if ($result.RebootRequired) { $rebootRequired = $true }
                    Write-Ok "Installed Teams Meeting Add-in to $targetDir"
                }
            }
        }
    }

    # -- 7. Verify -------------------------------------------------------------
    Write-Host ''
    Write-Step '7. Verification'
    if ($simulate) {
        Write-Skip 'Skipped - nothing was changed, so there is nothing to verify (-WhatIf)'
        Write-Host ''
        Write-Host '  Dry run only - rerun without -WhatIf to apply these changes.' -ForegroundColor Yellow
    } else {
        if (-not $SkipMeetingAddIn) {
            if (Get-TeamsMeetingAddInEntry) { Write-Ok 'Teams Meeting Add-in installed' }
            else { Write-Bad 'Teams Meeting Add-in installation failed'; $exitCode = 1 }
        }

        if (Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -eq 'MSTeams' }) {
            Write-Ok 'Teams provisioned for all users'
        } else {
            Write-Bad 'Teams installation failed'
            $exitCode = 1
        }

        Write-Host ''
        if ($rebootRequired) { Write-Warn 'A reboot is required to complete the installation (MSI returned 3010)' }
        if ($exitCode -eq 0) { Write-Ok 'Done' }
    }
} catch {
    Write-Host ''
    if ($cancelled) {
        Write-Skip $_.Exception.Message
    } else {
        Write-Bad "Aborted: $($_.Exception.Message)"
        $exitCode = 1
    }
} finally {
    if ($transcribing) {
        Write-Host ''
        try { Stop-Transcript | Out-Null } catch { Write-Warning "Transcript not closed cleanly: $($_.Exception.Message)" }
    }
}

Write-Host ''
exit $exitCode
