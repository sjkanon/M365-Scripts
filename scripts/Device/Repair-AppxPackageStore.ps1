#Requires -Version 5.1
<#
.SYNOPSIS
    Find and repair AppX packages that Windows lists but can no longer find, the state
    behind "Deployment Register operation ... from:  (AppxManifest.xml) failed with
    error 0x80070490".

.DESCRIPTION
    0x80070490 is ERROR_NOT_FOUND. Windows' package store says a package is there,
    the deployment engine goes looking and finds neither its files nor the user it
    belongs to. The empty path before (AppxManifest.xml) in the error is the giveaway:
    the registration being replayed has no install location left. Every supported
    route - Remove-AppxPackage, DISM, an installer such as teamsbootstrapper.exe -
    reads that same store, so they all fail the same way, and draining the host,
    signing users out or rebooting changes nothing.

    On an AVD host with FSLogix profile containers there is a second, more common
    route to the same error, logged under "Apps (Microsoft-FSLogix-Apps)" at sign-in.
    At sign-out FSLogix saves the user's packages - by full name, so exact version -
    to AppxPackages.xml in the profile, and at the next sign-in it re-registers them
    (Profiles\InstallAppxPackages, on by default). A host that does not have that
    exact version answers 0x80070490. The store on that host can be perfectly
    healthy; what fixes it is provisioning the package for all users, at the same
    build on every host in the pool, and an FSLogix build that registers Teams and
    Outlook by family name (2210 HF4 for Teams, 25.06 for Outlook). The saved list in
    the profile is not edited: FSLogix rewrites it at the next sign-out.

    This script handles both, for any package (MSTeams, Microsoft.OutlookForWindows,
    ...), in a fixed order: diagnose first, repair only what is demonstrably broken,
    and read everything back afterwards.

      1. Diagnose     Packages registered for a user whose files are gone, packages
                      provisioned from files that are gone, and the entries under
                      AppxAllUserStore that no longer correspond to anything. Always
                      runs; -CheckOnly stops here.
      1b. FSLogix     The FSLogix build, whether it replays packages at sign-in, which
                      packages it failed to register in the last -Days days against
                      what this host provisions, and AppX policies that block installs.
      1c. Failing     Every package that failed to install, update or register in the
                      last -Days days, from the AppX and FSLogix logs together: count,
                      error codes, versions asked for, and whether this host has them.
      1d. Copilot     With -Copilot: the Microsoft 365 Copilot app (MicrosoftOfficeHub)
                      and the Windows Copilot app (AppX, registered and provisioned),
                      the unified Microsoft Copilot app Edge Update installs, and every
                      policy that removes or blocks it - Edge Update's Install /
                      Uninstall / Update for the Copilot app id, the unification pause,
                      and Windows' WindowsCopilot / WindowsAI policies. Policies are
                      reported with path and value, never changed.
      2. Provisioned  Drop provisioned copies whose files are gone, so new profiles
                      stop being handed a package that cannot be registered.
      3. Re-register  A package whose files are still on disk but whose status is
                      not Ok is registered again from its own manifest.
      4. Remove       Registrations with nothing behind them are removed for all users,
                      and per user where -AllUsers refuses.
      5. Store        What is left is registry: each orphaned AppxAllUserStore key is
                      exported to a .reg backup and only then removed.
      6. Provision    With -Provision, Teams and new Outlook are provisioned for all
                      users with Microsoft's own installer (teamsbootstrapper.exe -p,
                      Outlook Setup.exe --provision true), or with -UseWinget from the
                      MSIX winget downloads. -WingetId does the same for any other
                      package, -Source for an MSIX you supply. All only after the
                      Microsoft signature is checked.
      7. Verify       The diagnosis runs again. Exit code 0 means nothing is broken
                      any more, 1 that something survived.

    With -Name '*' (the default) system packages and framework packages are reported
    but never touched, and Deprovisioned markers - which is how bloatware removals are
    remembered - are left alone. Name the package to have those considered as well.

.PARAMETER Name
    Package names to look at, wildcards allowed. Default '*': the whole store.
    Examples: 'MSTeams', 'Microsoft.OutlookForWindows', 'MSTeams','Microsoft.OutlookForWindows'.
    Shorthands: teams (MSTeams), outlook (Microsoft.OutlookForWindows) and copilot
    (the same as -Copilot), so -Name outlook,copilot is enough.

.PARAMETER CheckOnly
    Diagnose and report, change nothing. Exit code 2 when there is something to repair.

.PARAMETER Source
    Path to an .msix / .msixbundle / .appx(bundle) to provision once the store is
    clean. Its Authenticode signature must be valid and from Microsoft unless
    -SkipSignatureCheck is given.

.PARAMETER Provision
    Provision for all users, with Microsoft's own installer, the packages FSLogix
    failed to register that this host does not provision, and any of MSTeams /
    Microsoft.OutlookForWindows named explicitly in -Name. Other packages need
    -Source.

    Where FSLogix fails on a newer Teams or Outlook build than this host
    provisions, that exact build is provisioned instead: the MSIX is downloaded
    from Microsoft's CDN at the versioned URL winget's manifests use (checked to
    serve Outlook 1.2026.812-915 and Teams 26198-26246), its signature is checked,
    and the provisioned version is read back. The installers only ever deliver an
    older last-known-good build, and older is what fails.

.PARAMETER UseWinget
    With -Provision, take Teams / new Outlook from winget (Microsoft.Teams,
    Microsoft.Outlook) instead of Microsoft's installers: winget downloads the MSIX
    and checks its SHA256, the script checks the signature and provisions it with
    Add-AppxProvisionedPackage. winget's manifests lag behind the installers; the run
    says so when that build is older than what the profiles ask for.

.PARAMETER WingetId
    winget ids of any other packages to provision for all users the same way, e.g.
    to put back an app from the failing-apps overview. Only works when winget's
    manifest for it is an MSIX.

.PARAMETER IncludeDeprovisioned
    Also clear Deprovisioned markers for the named packages. Without this they are
    cleared only when -Name names a package without wildcards, because a marker is an
    instruction ("never provision this again") rather than damage, and with a wildcard
    it would undo every bloatware removal ever made on the host.

.PARAMETER SkipSignatureCheck
    Run the downloaded installer, or provision -Source, without checking that
    Microsoft signed it.

.PARAMETER Days
    How far back to read the FSLogix Apps event log. Default 7.

.PARAMETER WorkingDir
    Where the installers for -Provision are downloaded. Default C:\IT\AppxRepair.

.PARAMETER LogPath
    Folder for the transcript and the .reg backups. Default C:\Temp.

.PARAMETER ComputerName
    Session hosts to run on instead of this machine, e.g. lem-avd-4,lem-avd-5,lem-avd-6.
    The script copies itself to C:\IT\AppxRepair on each host over PowerShell
    remoting (WinRM), runs there with the same parameters, and ends with one table
    across the pool: exit code, FSLogix build and provisioned Teams / Outlook per
    host, with any difference between hosts named. A repair is confirmed once for
    the whole pool, not per change on each host.

.PARAMETER Credential
    Credential for the remoting sessions, when the current account is not an
    administrator on the hosts.

.PARAMETER Copilot
    Look at Copilot: adds Microsoft.MicrosoftOfficeHub and Microsoft.Copilot to -Name
    (so their Deprovisioned markers are in scope) and runs step 1d. What counts as
    installed is the new, unified Microsoft Copilot app that Edge Update puts on the
    machine.

    With -Provision that app is installed machine-wide the way Microsoft documents
    it: Edge Update's Install{C50565E9-CCCF-44B4-BA15-5AC5C6569197} = 5 (Force
    Installs), UpdaterExperimentationAndConfigurationServiceControl = 1 and
    CopilotUnificationAllowed{...} = 1 are written under
    HKLM\SOFTWARE\Policies\Microsoft\EdgeUpdate after a .reg backup of that key,
    Edge Update is asked to check now, and the run waits up to 10 minutes for the
    app. When it does not appear, the Microsoft 365 Copilot installer
    (M365CopilotDesktopInstaller.exe --quiet --start -p, signature-checked) is the
    fallback; the unification moves that old app over later. A policy that forbids
    or removes Copilot (Install = 0, Uninstall without Force Installs, Windows'
    TurnOff/Remove policies) is reported and fails the run, but is never overridden.

.PARAMETER Latest
    Provision the newest build there is of Teams / Outlook in scope, not only the
    one FSLogix failed on. Teams: Microsoft's config service, the feed the client
    itself uses (version and MSIX link). Outlook has no such feed - the Store
    catalog answered 1.2026.818.0 while 915.300 was already out - so the newest
    build that can be proven is used: the newest FSLogix asked for, registered for
    any user here, or present in WindowsApps. Use with -Provision.

.PARAMETER RemoveOld
    After provisioning, remove every reference this host keeps to an older build of
    the named packages: older provisioned copies, older builds registered for any
    user, and what AppxAllUserStore still remembers of them (backed up to .reg
    first). Only for packages named explicitly, and only once the build to keep is
    provisioned. The WindowsApps folders are left to Windows, and the list in each
    profile container to FSLogix, which rewrites it at the next sign-out.

.EXAMPLE
    # The very newest Teams and Outlook on the whole pool, and nothing older left
    .\Repair-AppxPackageStore.ps1 -ComputerName lem-avd-4,lem-avd-5,lem-avd-6 -Name teams,outlook -Latest -Provision -RemoveOld -Confirm:$false

.EXAMPLE
    # What is broken on this host? Changes nothing.
    .\Repair-AppxPackageStore.ps1 -CheckOnly

.EXAMPLE
    # Repair Teams and new Outlook and provision both at the current build, unattended
    .\Repair-AppxPackageStore.ps1 -Name 'MSTeams','Microsoft.OutlookForWindows' -Provision -Confirm:$false

.EXAMPLE
    # Same, but take both packages from winget
    .\Repair-AppxPackageStore.ps1 -Name 'MSTeams','Microsoft.OutlookForWindows' -Provision -UseWinget -Confirm:$false

.EXAMPLE
    # The whole pool: diagnose first, then repair and provision on every host
    .\Repair-AppxPackageStore.ps1 -ComputerName lem-avd-4,lem-avd-5,lem-avd-6 -Name MSTeams,Microsoft.OutlookForWindows -CheckOnly
    .\Repair-AppxPackageStore.ps1 -ComputerName lem-avd-4,lem-avd-5,lem-avd-6 -Name MSTeams,Microsoft.OutlookForWindows -Provision -Confirm:$false

.EXAMPLE
    # New Outlook and the new Copilot app on the whole pool: diagnose, then install
    .\Repair-AppxPackageStore.ps1 -ComputerName lem-avd-4,lem-avd-5,lem-avd-6 -Name outlook,copilot -CheckOnly
    .\Repair-AppxPackageStore.ps1 -ComputerName lem-avd-4,lem-avd-5,lem-avd-6 -Name outlook,copilot -Provision -Confirm:$false

.EXAMPLE
    # Every app that failed in the last 14 days, and the store for all of them
    .\Repair-AppxPackageStore.ps1 -CheckOnly -Days 14

.EXAMPLE
    # Repair and provision the current Teams from an MSIX in hand
    .\Repair-AppxPackageStore.ps1 -Name MSTeams -Source C:\IT\MSTeams-x64.msix

.NOTES
    Author  : Sjoerd Kanon
    Platform: Windows only (AppX), run as administrator or as System
    Exit    : 0 clean, 1 failed or something survived, 2 check-only found work
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param (
    [string[]] $Name = @('*'),
    [switch]   $CheckOnly,
    [string]   $Source,
    [switch]   $Provision,
    [switch]   $UseWinget,
    [string[]] $WingetId,
    [switch]   $IncludeDeprovisioned,
    [switch]   $SkipSignatureCheck,
    [ValidateRange(1, 90)]
    [int]      $Days       = 7,
    [string]   $WorkingDir = 'C:\IT\AppxRepair',
    [string]   $LogPath    = 'C:\Temp',
    [string[]] $ComputerName,
    [pscredential] $Credential,
    [switch]   $Copilot,
    [switch]   $Latest,
    [switch]   $RemoveOld
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# -- Several session hosts -----------------------------------------------------
# With -ComputerName this run only orchestrates: it copies itself to each host over
# PowerShell remoting, runs there with the same parameters, and ends with one table
# across the pool - which is the question on a pooled host ("is every host on the
# same build?") that no single host can answer.
$ComputerName = @($ComputerName | ForEach-Object { $_ -split '[,;\s]' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
if ($ComputerName.Count -gt 0) {
    $forward = @{}
    foreach ($entry in $PSBoundParameters.GetEnumerator()) {
        if ($entry.Key -in @('ComputerName', 'Credential', 'Confirm')) { continue }
        $forward[$entry.Key] = if ($entry.Value -is [switch]) { [bool] $entry.Value } else { $entry.Value }
    }

    # A remote session cannot answer a confirmation prompt reliably, so the
    # question is asked once here, for the whole pool, and not again per change.
    $changes = -not ($CheckOnly -or $WhatIfPreference)
    $confirmOff = $PSBoundParameters.ContainsKey('Confirm') -and -not $PSBoundParameters['Confirm']
    if ($changes -and -not $confirmOff) {
        $answer = Read-Host ("  Repair the package store on {0}? [y/N]" -f ($ComputerName -join ', '))
        if ($answer -notmatch '^[Yy]') { Write-Host '  Cancelled - nothing was changed.' -ForegroundColor DarkGray; exit 0 }
    }
    if ($changes) { $forward['Confirm'] = $false }
    if ($WhatIfPreference) { $forward['WhatIf'] = $true }

    $remotePath = 'C:\IT\AppxRepair\Repair-AppxPackageStore.ps1'
    $watch      = @('MSTeams', 'Microsoft.OutlookForWindows', 'Microsoft.MicrosoftOfficeHub')
    $summary    = [System.Collections.Generic.List[object]]::new()

    foreach ($computer in $ComputerName) {
        Write-Host ''
        Write-Host ("  ==== {0} " -f $computer).PadRight(80, '=') -ForegroundColor Cyan
        $row = [ordered]@{ Host = $computer; Exit = $null; FSLogix = ''; MSTeams = ''; Outlook = ''; Copilot = ''; Note = '' }
        $session = $null
        try {
            $sessionArgs = @{ ComputerName = $computer; ErrorAction = 'Stop' }
            if ($Credential) { $sessionArgs['Credential'] = $Credential }
            $session = New-PSSession @sessionArgs

            Invoke-Command -Session $session -ScriptBlock {
                param($Path) New-Item -ItemType Directory -Path (Split-Path $Path) -Force | Out-Null
            } -ArgumentList $remotePath
            Copy-Item -Path $PSCommandPath -Destination $remotePath -ToSession $session -Force

            # The host's own output streams back as it runs; the last object is the
            # state afterwards, read on the host itself.
            $result = Invoke-Command -Session $session -ScriptBlock {
                param($Path, $Params, $Watch)
                & $Path @Params
                $code = $LASTEXITCODE
                $frx  = Join-Path $env:ProgramFiles 'FSLogix\Apps\frxsvc.exe'
                $info = if (Test-Path $frx) { (Get-Item $frx).VersionInfo } else { $null }
                $prov = @{}
                foreach ($p in @(Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -in $Watch })) {
                    if (-not $prov.ContainsKey($p.DisplayName) -or [version] $prov[$p.DisplayName] -lt [version] $p.Version) { $prov[$p.DisplayName] = $p.Version }
                }
                $unified = Get-ItemProperty 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{C50565E9-CCCF-44B4-BA15-5AC5C6569197}' -ErrorAction SilentlyContinue
                [PSCustomObject]@{
                    RepairResult = $true
                    ExitCode     = $code
                    FSLogix      = if ($info) { '{0}.{1}.{2}.{3}' -f $info.FileMajorPart, $info.FileMinorPart, $info.FileBuildPart, $info.FilePrivatePart } else { 'not installed' }
                    Provisioned  = $prov
                    Copilot      = if ($unified -and $unified.pv) { "$($unified.pv) (unified)" } else { $null }
                }
            } -ArgumentList $remotePath, $forward, $watch

            $state = @($result | Where-Object { $_ -and $_.PSObject.Properties.Name -contains 'RepairResult' }) | Select-Object -Last 1
            if ($state) {
                $row.Exit    = $state.ExitCode
                $row.FSLogix = $state.FSLogix
                $row.MSTeams = if ($state.Provisioned.ContainsKey('MSTeams')) { $state.Provisioned['MSTeams'] } else { '-' }
                $row.Outlook = if ($state.Provisioned.ContainsKey('Microsoft.OutlookForWindows')) { $state.Provisioned['Microsoft.OutlookForWindows'] } else { '-' }
                $row.Copilot = if ($state.Copilot) { $state.Copilot }
                               elseif ($state.Provisioned.ContainsKey('Microsoft.MicrosoftOfficeHub')) { $state.Provisioned['Microsoft.MicrosoftOfficeHub'] }
                               else { '-' }
            } else {
                $row.Exit = 1; $row.Note = 'no result came back'
            }
        } catch {
            $row.Exit = 1
            # The full reason is printed above; the table keeps its gist.
            $gist     = ($_.Exception.Message -replace '^Connecting to remote server \S+ failed with the following error message : ', '') -replace '\s+', ' '
            $row.Note = 'not reached: ' + $(if ($gist.Length -gt 70) { $gist.Substring(0, 70) + '...' } else { $gist })
            Write-Host "  [FAIL] $computer - $($_.Exception.Message)" -ForegroundColor Red
        } finally {
            if ($session) { Remove-PSSession $session -ErrorAction SilentlyContinue }
        }
        $summary.Add([PSCustomObject] $row)
    }

    # Differences between hosts are the finding, so they are named, not left for
    # the reader to spot in the table.
    Write-Host ''
    Write-Host '  ==== Pool '.PadRight(80, '=') -ForegroundColor Cyan
    $summary | Format-Table -AutoSize | Out-String -Width 200 | Write-Host
    foreach ($column in 'FSLogix', 'MSTeams', 'Outlook', 'Copilot') {
        $values = @($summary | Where-Object { $null -ne $_.Exit -and -not $_.Note } | ForEach-Object { $_.$column } | Select-Object -Unique)
        if ($values.Count -gt 1) {
            Write-Host "  [WARN] $column differs between hosts ($($values -join ' / ')) - users moving between them get a different build" -ForegroundColor Yellow
        }
    }
    $worst = @($summary | ForEach-Object { [int] $_.Exit } | Sort-Object -Descending | Select-Object -First 1)[0]
    # 1 (failed) outranks 2 (check-only found work).
    if (@($summary | Where-Object { $_.Exit -eq 1 }).Count -gt 0) { $worst = 1 }
    Write-Host ''
    exit $worst
}

function Get-ForwardedArgument {
    <# Rebuild the caller's own parameters as a command line for a relaunch. #>
    param([Parameter(Mandatory)] $Bound)

    $list = @()
    foreach ($entry in $Bound.GetEnumerator()) {
        if ($entry.Value -is [switch] -or $entry.Value -is [bool]) {
            if ($entry.Value) { $list += "-$($entry.Key)" } else { $list += "-$($entry.Key):`$false" }
        } elseif ($entry.Value -is [array]) {
            $list += "-$($entry.Key)"
            $list += ($entry.Value -join ',')
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
# Under WOW64 HKLM reads land in WOW6432Node and the AppX cmdlets see a different
# world, so a 32-bit RMM agent gets a 64-bit relaunch with the same parameters.
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
# Get-AppxPackage -AllUsers answers "Access is denied" to anything less.
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
# NinjaOne hands script variables over as environment variables. They count only
# when the parameter itself was not passed.
$rmmTrue = @('true', '1', 'yes')
if (-not $PSBoundParameters.ContainsKey('WhatIf')               -and $env:whatIf               -in $rmmTrue) { $WhatIfPreference     = $true }
if (-not $PSBoundParameters.ContainsKey('CheckOnly')            -and $env:checkOnly            -in $rmmTrue) { $CheckOnly            = $true }
if (-not $PSBoundParameters.ContainsKey('IncludeDeprovisioned') -and $env:includeDeprovisioned -in $rmmTrue) { $IncludeDeprovisioned = $true }
if (-not $PSBoundParameters.ContainsKey('SkipSignatureCheck')   -and $env:skipSignatureCheck   -in $rmmTrue) { $SkipSignatureCheck   = $true }
if (-not $PSBoundParameters.ContainsKey('Provision')            -and $env:provision            -in $rmmTrue) { $Provision            = $true }
if (-not $PSBoundParameters.ContainsKey('UseWinget')            -and $env:useWinget            -in $rmmTrue) { $UseWinget            = $true }
if (-not $PSBoundParameters.ContainsKey('WingetId')             -and $env:wingetId)     { $WingetId = @($env:wingetId) }
if (-not $PSBoundParameters.ContainsKey('Days')                 -and $env:days -match '^\d+$')  { $Days    = [int] $env:days }
if (-not $PSBoundParameters.ContainsKey('WorkingDir')           -and $env:workingDir)   { $WorkingDir = $env:workingDir }
if (-not $PSBoundParameters.ContainsKey('Name')                 -and $env:packageName)  { $Name    = @($env:packageName -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }
if (-not $PSBoundParameters.ContainsKey('Source')               -and $env:source)       { $Source  = $env:source }
if (-not $PSBoundParameters.ContainsKey('LogPath')              -and $env:logPath)      { $LogPath = $env:logPath }

# powershell.exe -File passes 'MSTeams,Microsoft.OutlookForWindows' as one string -
# which is also how the relaunches above hand it over - so the list is split here.
$Name = @($Name | ForEach-Object { $_ -split '[,;]' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
if (-not $PSBoundParameters.ContainsKey('Copilot') -and $env:copilot -in $rmmTrue) { $Copilot = $true }
if (-not $PSBoundParameters.ContainsKey('Latest')  -and $env:latest  -in $rmmTrue) { $Latest  = $true }
if (-not $PSBoundParameters.ContainsKey('RemoveOld') -and $env:removeOld -in $rmmTrue) { $RemoveOld = $true }
# Shorthands, so -Name outlook,copilot is enough. 'copilot' is the -Copilot switch.
$Name = @(foreach ($entry in $Name) {
    switch ($entry) {
        'teams'   { 'MSTeams' }
        'outlook' { 'Microsoft.OutlookForWindows' }
        'copilot' { $Copilot = $true }
        default   { $entry }
    }
})
if ($Name.Count -eq 0 -and -not $Copilot) { $Name = @('*') }
# -Copilot names both Copilot packages, so they are explicit: their Deprovisioned
# markers are in scope and -Provision installs the app. Added to -Name when that was
# given, instead of it when it was left at '*'.
if ($Copilot) {
    $keep = @($Name | Where-Object { $_ -ne '*' -or $PSBoundParameters.ContainsKey('Name') -or $env:packageName })
    $Name = @($keep + @('Microsoft.MicrosoftOfficeHub', 'Microsoft.Copilot') | Select-Object -Unique)
}
$WingetId = @($WingetId | ForEach-Object { $_ -split '[,;]' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })

$confirmSuppressed = $PSBoundParameters.ContainsKey('Confirm') -and -not $PSBoundParameters['Confirm']
if ($confirmSuppressed) { $ConfirmPreference = 'None' }

$simulate     = [bool] $WhatIfPreference
$exitCode     = 0
$plannedExit  = $null
$transcribing = $false
$storeEdited  = $false
$script:unexplained = @()
$script:coverage    = @{}

# A name without wildcards is a deliberate choice of package; only then are system
# packages, frameworks and Deprovisioned markers in scope.
$explicitName = -not ($Name | Where-Object { [WildcardPattern]::ContainsWildcardCharacters($_) })

$AppxAllUserStorePath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Appx\AppxAllUserStore'
$ProfileListPath      = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'

# Where package files can live. A registration whose package folder is in none of
# them has nothing behind it.
$PackageRoots = @(
    (Join-Path $env:ProgramFiles 'WindowsApps')
    (Join-Path $env:WINDIR 'SystemApps')
    (Join-Path $env:WINDIR 'ImmersiveControlPanel')
)

# Packages Microsoft ships its own all-users provisioning installer for. -Provision
# uses these; anything else needs -Source. Both links were checked to resolve to
# Microsoft's CDN (teamsbootstrapper.exe, and the new Outlook Setup.exe).
# WingetId is the same package in the winget source, whose manifest points at the
# MSIX on Microsoft's CDN (checked: Microsoft.Teams -> MSTeams-x64.msix,
# Microsoft.Outlook -> Microsoft.OutlookForWindows_x64.msix). winget's manifests lag
# behind the installers above, which is why the installers are the default.
# VersionUrl is where Microsoft's CDN serves one exact build as an MSIX - the URLs
# winget's manifests point at, with the version in the path. Checked for Outlook
# 1.2026.812/818/902/915 and Teams 26198/26225/26246: every one answers 200. That is
# what lets -Provision put down exactly the build the profiles ask for, instead of
# the older last-known-good build the installers provision.
$KnownInstallers = @{
    'MSTeams' = @{
        Url        = 'https://go.microsoft.com/fwlink/?linkid=2243204&clcid=0x409'
        File       = 'teamsbootstrapper.exe'
        Args       = '-p'
        WingetId   = 'Microsoft.Teams'
        VersionUrl = 'https://teamsinstaller.public.onecdn.static.microsoft/production-windows-{1}/{0}/MSTeams-{1}.msix'
    }
    'Microsoft.OutlookForWindows' = @{
        Url        = 'https://go.microsoft.com/fwlink/?linkid=2207851'
        File       = 'OutlookSetup.exe'
        Args       = '--provision true --quiet --start-'
        WingetId   = 'Microsoft.Outlook'
        VersionUrl = 'https://res.cdn.office.net/nativehost/5mttl/installer/v2/{0}/Microsoft.OutlookForWindows_{1}.msix'
    }
    # The Microsoft 365 Copilot app, renamed Microsoft Copilot in 2026. Installer and
    # switches exactly as Microsoft Learn documents them ("Deploy the Microsoft 365
    # Copilot app"); the link was checked to deliver M365CopilotDesktopInstaller.exe,
    # an xpdBootstrapper signed by Microsoft. winget only has an .exe for it, so there
    # is no winget route. After unification the app can arrive through Edge Update
    # instead of as this AppX package, which is why EdgeUpdateId is checked as well.
    'Microsoft.MicrosoftOfficeHub' = @{
        Url          = 'https://go.microsoft.com/fwlink/?linkid=2325486'
        File         = 'M365CopilotDesktopInstaller.exe'
        Args         = '--quiet --start -p'
        WingetId     = $null
        EdgeUpdateId = '{C50565E9-CCCF-44B4-BA15-5AC5C6569197}'
    }
}

# Copilot on Windows, as of September 2026. The Microsoft 365 Copilot app
# (Microsoft.MicrosoftOfficeHub) and the Windows Copilot app (Microsoft.Copilot) are
# being unified into one Microsoft Copilot app that Edge Update installs and updates
# (app id below). Policies that decide whether it may be there live under
# EdgeUpdate; Microsoft Learn "Microsoft Copilot update policies for Windows" and
# "Pause the unified Microsoft Copilot application deployment".
$CopilotPackages       = @('Microsoft.MicrosoftOfficeHub', 'Microsoft.Copilot')
$CopilotEdgeUpdateId   = '{C50565E9-CCCF-44B4-BA15-5AC5C6569197}'
$EdgeUpdatePolicyPath  = 'HKLM:\SOFTWARE\Policies\Microsoft\EdgeUpdate'
$EdgeUpdateStatePath   = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate'
$EdgeUpdateMinimum     = [version] '1.3.253.25'

# FSLogix builds that register these packages by family name instead of by the
# exact version saved in the profile (release notes: 2210 HF4 for new Teams, 25.06
# for new Outlook). Older builds replay the saved full name, which is the 0x80070490.
$FslogixMinimum = @{
    'MSTeams'                     = [version] '2.9.8884.27471'
    'Microsoft.OutlookForWindows' = [version] '3.25.626.21064'
}
$FslogixServicePath  = Join-Path $env:ProgramFiles 'FSLogix\Apps\frxsvc.exe'
$FslogixProfilesPath = 'HKLM:\SOFTWARE\FSLogix\Profiles'
$FslogixOdfcPath     = 'HKLM:\SOFTWARE\Policies\FSLogix\ODFC'
$AppxPolicyPath      = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Appx'

# -- Output --------------------------------------------------------------------
function Write-Out  { param([string] $Message = '', [string] $Color = 'Gray') Write-Host $Message -ForegroundColor $Color }
function Write-Step { param([string] $Message) Write-Out "  $Message" 'Cyan' }
function Write-Ok   { param([string] $Message) Write-Out "  [ OK ] $Message" 'Green' }
function Write-Skip { param([string] $Message) Write-Out "  [SKIP] $Message" 'DarkGray' }
function Write-Warn { param([string] $Message) Write-Out "  [WARN] $Message" 'Yellow' }
function Write-Bad  { param([string] $Message) Write-Out "  [FAIL] $Message" 'Red' }

function Get-PropertyValue {
    <# Property access that returns $null instead of tripping Set-StrictMode. #>
    param($Object, [string] $Name)

    if ($null -eq $Object) { return $null }
    if ($Object.PSObject.Properties.Name -notcontains $Name) { return $null }
    return $Object.$Name
}

function Resolve-SidName {
    <# SID to DOMAIN\user, or the SID itself when it cannot be translated. #>
    param([Parameter(Mandatory)] [string] $Sid)

    try { return (New-Object Security.Principal.SecurityIdentifier $Sid).Translate([Security.Principal.NTAccount]).Value }
    catch { return $Sid }
}

function Test-ProfileOnHost {
    <# Whether a SID still has a profile on this machine. #>
    param([Parameter(Mandatory)] [string] $Sid)
    return (Test-Path "$ProfileListPath\$Sid")
}

function Test-NameInScope {
    <# Whether a package name or full name matches one of -Name. #>
    param([string] $Value)

    if (-not $Value) { return $false }
    # A full name is Name_Version_Arch_ResourceId_PublisherId; a family Name_PublisherId.
    $short = ($Value -split '_')[0]
    foreach ($pattern in $Name) {
        if ($short -like $pattern -or $Value -like $pattern) { return $true }
    }
    return $false
}

function Test-PackageFiles {
    <# Whether a package full name has a folder in any of the package roots. #>
    param([Parameter(Mandatory)] [string] $FullName)

    foreach ($root in $PackageRoots) {
        if (Test-Path -LiteralPath (Join-Path $root $FullName)) { return $true }
    }
    return $false
}

function Get-AppxPackageHolder {
    <#
        The users a package is registered for. PackageUserInformation renders as
        "S-1-5-21-... [DOMAIN\user]: Installed", so the SID and state are taken from
        that string rather than a property path that differs between builds.
    #>
    param($Package)

    foreach ($holder in @(Get-PropertyValue $Package 'PackageUserInformation')) {
        $text = [string] $holder
        if ($text -notmatch '(S-1-[0-9\-]+)') { continue }
        $sid     = $Matches[1]
        $state   = if ($text -match ':\s*([^:]+)$') { $Matches[1].Trim() } else { 'unknown' }
        $account = if ($text -match '\[([^\]]+)\]' -and $Matches[1] -ne $sid) { $Matches[1] } else { Resolve-SidName $sid }
        [PSCustomObject]@{ Sid = $sid; Account = $account; State = $state }
    }
}

# -- Diagnosis -----------------------------------------------------------------
function Get-BrokenInstalledPackage {
    <#
        Packages registered for at least one user whose files are gone (Ghost) or
        whose files are there but whose status is not Ok (Damaged). Frameworks and
        system packages are only considered when named explicitly: they are shared by
        other apps, and "reported but not touched" is the safe default for them.
    #>
    param($Installed)

    # Newest installed version per package (name + architecture + resource). An older
    # version next to a newer one is superseded, not damaged: its status is not Ok
    # because Windows is waiting to remove it, and re-registering it can only answer
    # 0x80073D06 ("a higher version is already installed") - measured on aimgr
    # 0.20.61.0 next to 0.20.62.0.
    $newest = @{}
    foreach ($pkg in $Installed) {
        $key     = '{0}|{1}|{2}' -f $pkg.Name, (Get-PropertyValue $pkg 'Architecture'), (Get-PropertyValue $pkg 'ResourceId')
        $version = try { [version] (Get-PropertyValue $pkg 'Version') } catch { $null }
        if ($version -and (-not $newest.ContainsKey($key) -or $newest[$key] -lt $version)) { $newest[$key] = $version }
    }

    foreach ($pkg in $Installed) {
        if (-not (Test-NameInScope $pkg.Name)) { continue }

        $location  = Get-PropertyValue $pkg 'InstallLocation'
        $filesGone = (-not $location) -or (-not (Test-Path -LiteralPath $location))
        $status    = [string] (Get-PropertyValue $pkg 'Status')
        $damaged   = (-not $filesGone) -and $status -and $status -ne 'Ok'
        if (-not ($filesGone -or $damaged)) { continue }

        $key        = '{0}|{1}|{2}' -f $pkg.Name, (Get-PropertyValue $pkg 'Architecture'), (Get-PropertyValue $pkg 'ResourceId')
        $version    = try { [version] (Get-PropertyValue $pkg 'Version') } catch { $null }
        $superseded = $damaged -and $version -and $newest.ContainsKey($key) -and $version -lt $newest[$key]

        $protected = ([string] (Get-PropertyValue $pkg 'SignatureKind') -eq 'System') -or
                     [bool] (Get-PropertyValue $pkg 'IsFramework') -or
                     [bool] (Get-PropertyValue $pkg 'NonRemovable')

        [PSCustomObject]@{
            Package   = $pkg
            FullName  = $pkg.PackageFullName
            Name      = $pkg.Name
            Location  = $location
            Problem   = if ($filesGone) { 'Ghost' } elseif ($superseded) { 'Superseded' } else { 'Damaged' }
            Reason    = if ($filesGone -and $location) { "its files are gone ($location)" }
                        elseif ($filesGone)          { 'the store holds no install location for it at all' }
                        elseif ($superseded)         { "its status is $status and $($newest[$key]) is installed next to it - Windows removes this one once no user holds it any more; nothing to repair" }
                        else                         { "its status is $status" }
            Protected = ($protected -and -not $explicitName) -or $superseded
            Holders   = @(Get-AppxPackageHolder -Package $pkg)
        }
    }
}

function Get-BrokenProvisionedPackage {
    <# Provisioned packages whose files are gone - every new profile trips over them. #>
    foreach ($prov in @(Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue)) {
        if (-not (Test-NameInScope $prov.DisplayName)) { continue }
        $manifest = [Environment]::ExpandEnvironmentVariables([string] (Get-PropertyValue $prov 'InstallLocation'))
        if ($manifest -and (Test-Path -LiteralPath $manifest)) { continue }
        if (-not $manifest -and (Test-PackageFiles $prov.PackageName)) { continue }
        [PSCustomObject]@{
            PackageName = $prov.PackageName
            DisplayName = $prov.DisplayName
            Reason      = if ($manifest) { "the manifest it provisions from is gone ($manifest)" } else { 'it records no install location' }
        }
    }
}

function Get-AppxStoreRegistration {
    <#
        Every place AppxAllUserStore remembers a package, and whether that memory
        still corresponds to anything:

          <SID>\<package>                 registered for that user
          EndOfLife\<SID>\<package>       superseded, waiting for that user to go
          DeferredRemoval\<SID>\<package> removal parked until that user signs out
          Applications\<package>          machine-wide entry, with its manifest path
          Deprovisioned\<family>          "never provision this again"

        A user entry is orphaned when its SID has no profile left or its package has
        no files anywhere; an Applications entry when its manifest is gone. Those are
        what the deployment engine walks and answers 0x80070490 to. Staged entries are
        never touched.
    #>
    param([hashtable] $LiveLocations, [hashtable] $Protected = @{})

    if (-not (Test-Path $AppxAllUserStorePath)) { return }

    function Test-Gone([string] $FullName) {
        # A full name has five parts; anything else is not a package entry.
        if (($FullName -split '_').Count -lt 5) { return $false }
        if ($LiveLocations.ContainsKey($FullName)) { return -not $LiveLocations[$FullName] }
        return -not (Test-PackageFiles $FullName)
    }

    foreach ($scope in @(
        @{ Root = $AppxAllUserStorePath;                   Kind = 'user registration' }
        @{ Root = "$AppxAllUserStorePath\EndOfLife";       Kind = 'end-of-life entry' }
        @{ Root = "$AppxAllUserStorePath\DeferredRemoval"; Kind = 'deferred removal' }
    )) {
        foreach ($sidKey in @(Get-ChildItem $scope.Root -ErrorAction SilentlyContinue |
                              Where-Object { $_.PSChildName -like 'S-1-*' })) {
            $sid       = $sidKey.PSChildName
            $noProfile = -not (Test-ProfileOnHost $sid)
            foreach ($pkgKey in @(Get-ChildItem $sidKey.PSPath -ErrorAction SilentlyContinue |
                                  Where-Object { Test-NameInScope $_.PSChildName })) {
                $noFiles = Test-Gone $pkgKey.PSChildName
                # A system or framework package that still has its files is left to
                # Windows unless it was named explicitly.
                if (-not $noFiles -and -not $explicitName -and $Protected.ContainsKey($pkgKey.PSChildName)) { continue }
                [PSCustomObject]@{
                    Kind     = $scope.Kind
                    Account  = Resolve-SidName $sid
                    Name     = $pkgKey.PSChildName
                    Path     = $pkgKey.Name
                    Orphaned = $noProfile -or $noFiles
                    Reason   = if ($noFiles)       { 'the package has no files on this host' }
                               elseif ($noProfile) { "$sid has no profile on this host, so nothing will ever complete it" }
                               else                { '' }
                }
            }
        }
    }

    foreach ($appKey in @(Get-ChildItem "$AppxAllUserStorePath\Applications" -ErrorAction SilentlyContinue |
                          Where-Object { Test-NameInScope $_.PSChildName })) {
        $manifest = Get-PropertyValue (Get-ItemProperty $appKey.PSPath -ErrorAction SilentlyContinue) 'Path'
        $gone     = (-not $manifest) -or (-not (Test-Path -LiteralPath $manifest))
        [PSCustomObject]@{
            Kind     = 'machine registration'
            Account  = $null
            Name     = $appKey.PSChildName
            Path     = $appKey.Name
            Orphaned = $gone
            Reason   = if (-not $gone)   { '' }
                       elseif ($manifest) { "the manifest it was staged from is gone ($manifest)" }
                       else               { 'it records no manifest path at all' }
        }
    }

    foreach ($key in @(Get-ChildItem "$AppxAllUserStorePath\Deprovisioned" -ErrorAction SilentlyContinue |
                       Where-Object { Test-NameInScope $_.PSChildName })) {
        $inScope = $IncludeDeprovisioned -or $explicitName
        [PSCustomObject]@{
            Kind     = 'deprovisioned marker'
            Account  = $null
            Name     = $key.PSChildName
            Path     = $key.Name
            Orphaned = $inScope
            Reason   = if ($inScope) { 'it tells Windows never to provision this package again, which refuses the reinstall' } else { '' }
        }
    }
}

function Get-OrphanedStoreEntry {
    <#
        The orphaned AppxAllUserStore entries, judged against what Get-AppxPackage
        lists right now: full name -> whether its install location exists, and which
        full names are system or framework packages.
    #>
    param($Installed = @(Get-AppxPackage -AllUsers -ErrorAction Stop))

    $live = @{}; $protected = @{}
    foreach ($pkg in $Installed) {
        $location = Get-PropertyValue $pkg 'InstallLocation'
        $live[$pkg.PackageFullName] = [bool] ($location -and (Test-Path -LiteralPath $location))
        if (([string] (Get-PropertyValue $pkg 'SignatureKind') -eq 'System') -or
            [bool] (Get-PropertyValue $pkg 'IsFramework') -or
            [bool] (Get-PropertyValue $pkg 'NonRemovable')) { $protected[$pkg.PackageFullName] = $true }
    }
    return @(Get-AppxStoreRegistration -LiveLocations $live -Protected $protected | Where-Object { $_.Orphaned })
}

function Get-StoreHealth {
    <# One full reading of the store: installed, provisioned and registry. #>
    $installed = @(Get-AppxPackage -AllUsers -ErrorAction Stop)

    [PSCustomObject]@{
        Installed   = @(Get-BrokenInstalledPackage -Installed $installed)
        Provisioned = @(Get-BrokenProvisionedPackage)
        Registry    = @(Get-OrphanedStoreEntry -Installed $installed)
    }
}

function Write-StoreHealth {
    param($Health)

    foreach ($item in $Health.Installed) {
        # Superseded is not a fault, so it is one grey line instead of a warning.
        if ($item.Problem -eq 'Superseded') {
            Write-Skip "Superseded: $($item.FullName) - $($item.Reason)"
            continue
        }
        $tag = if ($item.Protected) { ' (system or framework package - reported, not touched; name it to include it)' } else { '' }
        Write-Warn "$($item.Problem): $($item.FullName)$tag"
        Write-Warn "  $($item.Reason)"
        foreach ($holder in $item.Holders) { Write-Skip "  registered for $($holder.Account): $($holder.State)" }
    }
    foreach ($item in $Health.Provisioned) {
        Write-Warn "Provisioned without files: $($item.PackageName)"
        Write-Warn "  $($item.Reason)"
    }
    foreach ($entry in $Health.Registry) {
        $who = if ($entry.Account) { " for $($entry.Account)" } else { '' }
        Write-Warn "Store orphan - $($entry.Kind)$who`: $($entry.Name)"
        Write-Warn "  $($entry.Reason)"
        Write-Skip "  $($entry.Path)"
    }
}

function Get-RepairableCount {
    param($Health)
    return @($Health.Installed | Where-Object { -not $_.Protected }).Count +
           $Health.Provisioned.Count + $Health.Registry.Count
}

function Write-AppxDeploymentError {
    <#
        The deployment log says which package and why, where the error code alone
        does not. Read-only, and quiet when the log holds nothing recent.
    #>
    param([int] $Minutes = 60, [int] $MaxEvents = 5)

    $since  = (Get-Date).AddMinutes(-$Minutes)
    $events = @(Get-EventSafe -Filter @{ LogName = 'Microsoft-Windows-AppXDeploymentServer/Operational'; Level = 2; StartTime = $since } -MaxEvents 50 |
                Where-Object { $_.Message -match '([\w\.\-]+_[\d\.]+_\w*_[\w\.\-]*_\w{13})' -and
                               (Test-NameInScope $Matches[1]) } |
                Select-Object -First $MaxEvents)

    if ($events.Count -eq 0) {
        Write-Skip "No AppX deployment errors for these packages in the last $Minutes minutes"
        return
    }
    foreach ($entry in $events) {
        $text = (($entry.Message -replace '\s+', ' ')).Trim()
        if ($text.Length -gt 240) { $text = $text.Substring(0, 240) + '...' }
        Write-Warn ('AppX deployment {0:HH:mm}: {1}' -f $entry.TimeCreated, $text)
    }
}

function Backup-RegistryKey {
    <# Export a key to a .reg file before it is removed; returns whether that worked. #>
    param([Parameter(Mandatory)] [string] $Key, [Parameter(Mandatory)] [string] $Folder)

    if (-not (Test-Path $Folder)) { New-Item -ItemType Directory -Path $Folder -Force | Out-Null }
    # reg.exe wants the long hive name.
    $Key  = $Key -replace '^HKLM:\\?', 'HKEY_LOCAL_MACHINE\' -replace '^HKCU:\\?', 'HKEY_CURRENT_USER\'
    $file = Join-Path $Folder (($Key -replace '^HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Appx\\', '' -replace '[\\/:*?"<>|]', '_') + '.reg')
    # Under Windows PowerShell 5.1 with ErrorActionPreference Stop, a native command
    # writing to stderr is a terminating error - a failed export would have ended the
    # whole run instead of answering "no backup". Measured with reg.exe.
    $ErrorActionPreference = 'Continue'
    & reg.exe export $Key $file /y 2>$null | Out-Null
    return ($LASTEXITCODE -eq 0 -and (Test-Path $file))
}

function Test-MicrosoftSignature {
    <# Valid Authenticode signature from Microsoft on a package file. #>
    param([Parameter(Mandatory)] [string] $Path)

    $sig = Get-AuthenticodeSignature -FilePath $Path
    return ($sig.Status -eq 'Valid' -and $sig.SignerCertificate.Subject -match 'O=Microsoft Corporation')
}

function Get-EventSafe {
    <#
        Get-WinEvent throws a terminating "The parameter is incorrect" for a provider
        that is not registered (no FSLogix on the machine) and "No events were found"
        for an empty result; -ErrorAction covers neither. Both mean "nothing" here.
    #>
    param([Parameter(Mandatory)] [hashtable] $Filter, [int] $MaxEvents = 5000)
    try { return @(Get-WinEvent -FilterHashtable $Filter -MaxEvents $MaxEvents -ErrorAction Stop) }
    catch { return @() }
}

# -- FSLogix -------------------------------------------------------------------
function Get-PackageVersionFromFullName {
    <# The version part of Name_Version_Arch_ResourceId_PublisherId. #>
    param([string] $FullName)
    $parts = $FullName -split '_'
    if ($parts.Count -lt 5) { return $null }
    try { return [version] $parts[1] } catch { return $null }
}

function Get-ProvisionedVersion {
    <# Highest provisioned version per package name, for the names in scope. #>
    $map = @{}
    foreach ($prov in @(Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue)) {
        $version = try { [version] $prov.Version } catch { $null }
        if (-not $version) { continue }
        if (-not $map.ContainsKey($prov.DisplayName) -or $map[$prov.DisplayName] -lt $version) {
            $map[$prov.DisplayName] = $version
        }
    }
    return $map
}

function Get-FslogixVersion {
    <#
        The installed FSLogix build, or $null. The numeric file version, not
        ProductVersion: that one can carry a suffix that does not parse.
    #>
    if (-not (Test-Path $FslogixServicePath)) { return $null }
    $info = (Get-Item $FslogixServicePath).VersionInfo
    return [version] ('{0}.{1}.{2}.{3}' -f $info.FileMajorPart, $info.FileMinorPart, $info.FileBuildPart, $info.FilePrivatePart)
}

function Get-FslogixRequest {
    <#
        What FSLogix tried to register at sign-in and could not. At sign-out FSLogix
        writes the user's packages - full names, so exact versions - to
        AppxPackages.xml in the profile container, and at the next sign-in it
        replays that list (Profiles\InstallAppxPackages, on by default). A host that
        does not have that exact version answers 0x80070490 with an empty path,
        which is the error in the Apps event log. One object per requested package,
        with how often and when it last failed.

        The same request also shows in the AppX deployment log, under the user's own
        SID, and on some hosts only there: LEM-AVD-5 failed new Outlook 10-16 times per
        sign-in with 0x80073CF9 / 0x80070490 while the FSLogix log named nothing, so
        a host repair found no build to provision. A 0x80070490 there for a version
        whose files are not on this host counts as a request as well.
    #>
    $since  = (Get-Date).AddDays(-$Days)
    $events = @(Get-EventSafe -Filter @{ ProviderName = 'Microsoft-FSLogix-Apps'; Level = 2; StartTime = $since })
    $appx   = @(Get-EventSafe -Filter @{ LogName = 'Microsoft-Windows-AppXDeploymentServer/Operational'; Level = 2; StartTime = $since } |
                Where-Object { $_.Message -match '0x80070490' })

    $found = @{}
    foreach ($entry in @($events) + @($appx)) {
        $fullName = $null
        if ($entry.ProviderName -eq 'Microsoft-FSLogix-Apps') {
            if ($entry.Message -match 'on Package\s+(\S+?)\s+from') { $fullName = $Matches[1] }
        } elseif ($entry.Message -match '([\w\.\-]+_\d+\.\d+\.\d+\.\d+_\w*_[\w\.\-]*_\w{13})') {
            # AppX: only a build this host does not have is a request it can answer.
            $fullName = $Matches[1]
            if (Test-PackageFiles $fullName) { continue }
        }
        if (-not $fullName) { continue }
        if (-not (Test-NameInScope $fullName)) { continue }
        $code = if ($entry.Message -match '(0x[0-9A-Fa-f]{8})') { $Matches[1] } else { 'unknown' }
        if (-not $found.ContainsKey($fullName)) {
            $found[$fullName] = [PSCustomObject]@{
                FullName = $fullName
                Name     = ($fullName -split '_')[0]
                Version  = Get-PackageVersionFromFullName $fullName
                Code     = $code
                Count    = 0
                Last     = $entry.TimeCreated
            }
        }
        $found[$fullName].Count++
        if ($entry.TimeCreated -gt $found[$fullName].Last) { $found[$fullName].Last = $entry.TimeCreated }
    }
    return @($found.Values | Sort-Object Name, Version)
}

function Write-VersionGap {
    <#
        Profiles asking for a newer build than this host provisions, and FSLogix
        failing on it - which is only ever called with failures in hand. Teams and new
        Outlook update themselves per user, and Microsoft's installers provision a
        last-known-good build behind that (measured: Teams 26225 against 26246,
        Outlook 1.2026.818 against 902 and 915).

        An earlier version called this harmless once FSLogix was new enough to
        register by family name. A production host proved otherwise: Outlook failed
        186 times with 0x80070490 while the files of both requested builds were on
        disk. So the gap is work whenever the failures are there, and the fix is to
        provision exactly the build asked for. Returns $true.
    #>
    param([string] $Package, $Have, $Asked, $Fslogix)

    $exact   = $KnownInstallers.ContainsKey($Package) -and $KnownInstallers[$Package].ContainsKey('VersionUrl')
    $minimum = if ($FslogixMinimum.ContainsKey($Package)) { $FslogixMinimum[$Package] } else { $null }
    if ($exact) {
        Write-Warn "  $Package provisioned at $Have while the profiles ask for $Asked and FSLogix fails on it - -Provision installs exactly $Asked from Microsoft's CDN"
    } else {
        Write-Warn "  $Package provisioned at $Have, older than the $Asked profiles ask for - provision the newer build on every host in the pool"
    }
    if ($minimum -and $Fslogix -and $Fslogix -lt $minimum) {
        Write-Warn "  FSLogix $Fslogix predates $minimum and replays the exact saved version, so the next self-update opens the same gap - update FSLogix on the image as well"
    } else {
        Write-Skip '  The app updates itself per user, so a later update can open a new gap - run this on a schedule to keep the hosts level'
    }
    return $true
}

function Get-ExactTarget {
    <#
        The builds to provision exactly: per package with a VersionUrl, the newest
        version FSLogix failed on, when that is newer than what this host provisions.
        One object per package, with the URL for this host's architecture.
    #>
    param([hashtable] $Provisioned, $Requests)

    foreach ($group in @($Requests | Group-Object Name)) {
        if (-not $KnownInstallers.ContainsKey($group.Name)) { continue }
        $spec = $KnownInstallers[$group.Name]
        if (-not $spec.ContainsKey('VersionUrl')) { continue }
        $newest = $group.Group | Sort-Object Version -Descending | Select-Object -First 1
        if (-not $newest.Version) { continue }
        if ($Provisioned.ContainsKey($group.Name) -and $Provisioned[$group.Name] -ge $newest.Version) { continue }
        $arch = (($newest.FullName -split '_')[2]).ToLowerInvariant()
        if ($arch -notin @('x64', 'arm64', 'x86')) { $arch = 'x64' }
        [PSCustomObject]@{
            Name     = $group.Name
            Version  = $newest.Version
            FullName = $newest.FullName
            Url      = $spec.VersionUrl -f $newest.Version, $arch
        }
    }
}

function Get-LatestTarget {
    <#
        -Latest: the newest build there is, per Teams / Outlook in scope, when it is
        newer than what this host provisions.

          Teams    Microsoft's config service, the feed the client itself uses to
                   decide it is out of date - newest version and its MSIX link.
          Outlook  No such feed exists: the Store catalog answered 1.2026.818.0 while
                   915.300 was already on the CDN and in users' profiles. So the
                   newest build that can be proven is used - the newest of what
                   FSLogix asked for, what is registered for any user here, and what
                   has a folder in WindowsApps - and the run says that is what it is.
    #>
    param([hashtable] $Provisioned, $Requests)

    foreach ($pkgName in @('MSTeams', 'Microsoft.OutlookForWindows')) {
        if (-not (Test-NameInScope $pkgName)) { continue }
        $spec = $KnownInstallers[$pkgName]
        $arch = if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64') { 'arm64' } else { 'x64' }
        $best = $null; $source = $null; $url = $null

        if ($pkgName -eq 'MSTeams') {
            try {
                [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
                $config = Invoke-RestMethod -UseBasicParsing -TimeoutSec 30 -Uri ('https://config.teams.microsoft.com/config/v1/MicrosoftTeams/0.0.0.0' +
                          '?environment=prod&audienceGroup=general&teamsRing=general&agent=TeamsBuilds')
                $node = Get-PropertyValue (Get-PropertyValue (Get-PropertyValue $config 'BuildSettings') 'WebView2PreAuth') $arch
                $v    = Get-PropertyValue $node 'latestVersion'
                if ($v) { $best = [version] $v; $source = 'the Teams config service'; $url = Get-PropertyValue $node 'buildLink' }
            } catch {
                Write-Warn "  Could not reach the Teams config service: $($_.Exception.Message)"
            }
        }

        # Every build of this package this host can prove exists.
        $seen = @()
        $seen += @($Requests | Where-Object { $_.Name -eq $pkgName } | ForEach-Object { $_.Version })
        $seen += @(Get-AppxPackage -AllUsers -Name $pkgName -ErrorAction SilentlyContinue | ForEach-Object { try { [version] $_.Version } catch { } })
        $seen += @(Get-ChildItem (Join-Path $env:ProgramFiles 'WindowsApps') -Directory -Filter "${pkgName}_*_${arch}__*" -ErrorAction SilentlyContinue |
                   ForEach-Object { Get-PackageVersionFromFullName $_.Name })
        $seenBest = @($seen | Where-Object { $_ } | Sort-Object -Descending) | Select-Object -First 1
        if ($seenBest -and (-not $best -or $seenBest -gt $best)) {
            $best = $seenBest; $source = 'the newest build seen on this host and in the profiles'; $url = $null
        }
        if (-not $best) { continue }

        $have = if ($Provisioned.ContainsKey($pkgName)) { $Provisioned[$pkgName] } else { $null }
        if ($have -and $have -ge $best) {
            Write-Ok "  $pkgName $have provisioned - the newest build there is ($source)"
            continue
        }
        Write-Warn "  $pkgName newest build is $best ($source), this host provisions $(if ($have) { $have } else { 'nothing' })"
        [PSCustomObject]@{
            Name     = $pkgName
            Version  = $best
            FullName = $null
            Url      = if ($url) { $url } else { $spec.VersionUrl -f $best, $arch }
        }
    }
}

function Invoke-ExactProvision {
    <#
        Provision one exact build for all users: the MSIX straight from Microsoft's
        CDN, its signature checked, then Add-AppxProvisionedPackage, then the
        provisioned version read back. Returns $true when that version is there.
    #>
    param([Parameter(Mandatory)] $Target)

    if (-not (Test-Path $WorkingDir)) { New-Item -ItemType Directory -Path $WorkingDir -Force | Out-Null }
    $file = Join-Path $WorkingDir ('{0}_{1}.msix' -f $Target.Name, $Target.Version)

    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    $ProgressPreference = 'SilentlyContinue'
    Invoke-WebRequest -Uri $Target.Url -OutFile $file -UseBasicParsing
    Write-Skip ("  Downloaded {0} ({1} MB) from {2}" -f (Split-Path $file -Leaf), [math]::Round((Get-Item $file).Length / 1MB, 0), $Target.Url)
    if (-not $SkipSignatureCheck -and -not (Test-MicrosoftSignature -Path $file)) {
        throw "$(Split-Path $file -Leaf) is not validly signed by Microsoft - refusing to provision it"
    }

    Add-AppxProvisionedPackage -Online -PackagePath $file -SkipLicense -ErrorAction Stop | Out-Null
    $now = (Get-ProvisionedVersion)[$Target.Name]
    if ($now -and $now -ge $Target.Version) {
        Write-Ok "$($Target.Name) $now provisioned for all users - the build the profiles ask for"
        Remove-Item -LiteralPath $file -Force -ErrorAction SilentlyContinue
        return $true
    }
    Write-Bad "$($Target.Name): Add-AppxProvisionedPackage completed but the provisioned version is $now, not $($Target.Version)"
    return $false
}

function Write-FslogixStatus {
    <#
        The FSLogix side: which build, whether it replays AppX packages at sign-in,
        and what it failed to register lately against what this host provisions.
        Returns the package names that need provisioning on this host.
    #>
    param([hashtable] $Provisioned, $Requests)

    $needs = @()
    $fslogixVersion = Get-FslogixVersion
    if (-not $fslogixVersion) {
        Write-Skip 'FSLogix is not installed'
        return $needs
    }
    Write-Ok "FSLogix $fslogixVersion"

    $replay = Get-PropertyValue (Get-ItemProperty $FslogixProfilesPath -ErrorAction SilentlyContinue) 'InstallAppxPackages'
    if ($null -eq $replay -or $replay -eq 1) {
        Write-Skip '  InstallAppxPackages is on (default): packages saved in the profile are re-registered at sign-in'
    } else {
        Write-Skip '  InstallAppxPackages is 0: FSLogix does not re-register packages at sign-in'
    }

    $odfc = Get-ItemProperty $FslogixOdfcPath -ErrorAction SilentlyContinue
    if ((Get-PropertyValue $odfc 'Enabled') -eq 1 -and (Get-PropertyValue $odfc 'IncludeTeams') -ne 1) {
        Write-Warn '  ODFC is enabled without IncludeTeams=1 - new Teams is then not registered by family name at sign-in'
    }

    foreach ($pkgName in $FslogixMinimum.Keys) {
        if (-not (Test-NameInScope $pkgName)) { continue }
        if ($fslogixVersion -and $fslogixVersion -lt $FslogixMinimum[$pkgName]) {
            Write-Warn "  FSLogix $fslogixVersion predates $($FslogixMinimum[$pkgName]), the first build that registers $pkgName by family instead of by the exact saved version - update FSLogix on the image"
        }
    }

    if ($Requests.Count -eq 0) {
        Write-Ok "  No failed package registrations from FSLogix in the last $Days day(s)"
        return $needs
    }

    foreach ($group in @($Requests | Group-Object Name)) {
        $have   = if ($Provisioned.ContainsKey($group.Name)) { $Provisioned[$group.Name] } else { $null }
        $newest = ($group.Group | Sort-Object Version -Descending | Select-Object -First 1).Version
        $total  = ($group.Group | Measure-Object -Property Count -Sum).Sum
        $last   = ($group.Group | Sort-Object Last -Descending | Select-Object -First 1).Last

        Write-Warn ("FSLogix could not register {0} at sign-in: {1}x, last {2:yyyy-MM-dd HH:mm}" -f $group.Name, $total, $last)
        foreach ($request in $group.Group) {
            $onDisk = if (Test-PackageFiles $request.FullName) { 'on this host' } else { 'not on this host' }
            Write-Skip "  asked for $($request.FullName) ($onDisk, $($request.Code))"
        }

        if (-not $have) {
            Write-Bad "  This host does not provision $($group.Name) at all - users whose profile asks for it get nothing. Provision it (-Provision)"
            $needs += $group.Name
        } elseif ($newest -and $have -lt $newest) {
            # [void]: its $true/$false would otherwise land in this function's output
            # and be returned as a package name to provision.
            [void] (Write-VersionGap -Package $group.Name -Have $have -Asked $newest -Fslogix $fslogixVersion)
        } elseif ($last -gt (Get-Date).AddHours(-24)) {
            # Measured on a production host: Outlook 1.2026.915.300 provisioned, the
            # profiles asking for 902 and 915, FSLogix 26.01 - and still failing that
            # same afternoon. Calling that "an old saved version that clears at the
            # next sign-out" was a guess. Not a version gap, so it is said as such
            # and the evidence is shown in step 1c.
            Write-Warn "  This host provisions $have, which is what the profiles ask for, yet FSLogix still failed on it in the last 24 hours - not a version gap; step 1c shows what Windows and FSLogix logged"
            $script:unexplained += $group.Name
        } else {
            Write-Ok "  This host provisions $have and nothing failed in the last 24 hours - the older failures were FSLogix replaying a saved version"
        }
    }
    return $needs
}

function Write-UserCoverage {
    <#
        The question that decides whether a failure matters: do the users who are
        signed in right now have the app? FSLogix's replay can fail while Windows
        registers the provisioned build at sign-in anyway, and then the error is
        noise. Signed-in means a loaded user hive; "has it" means registered and
        Installed for that SID, whatever the version.
    #>
    param([Parameter(Mandatory)] [string] $PackageName)

    $signedIn = @(Get-ChildItem 'Registry::HKEY_USERS' -ErrorAction SilentlyContinue |
                  Where-Object { $_.PSChildName -match '^S-1-(5-21|12-1)-[\d-]+$' } |
                  ForEach-Object { $_.PSChildName })
    if ($signedIn.Count -eq 0) {
        Write-Skip "    nobody is signed in, so whether users get $PackageName cannot be seen right now"
        return
    }

    $holders = @{}
    foreach ($pkg in @(Get-AppxPackage -AllUsers -Name $PackageName -ErrorAction SilentlyContinue)) {
        foreach ($holder in @(Get-AppxPackageHolder -Package $pkg)) {
            if ($holder.State -notmatch '^Installed' -or $holder.State -match 'pending removal') { continue }
            # The newest build a user has, not whichever was listed last.
            $v = try { [version] $pkg.Version } catch { $null }
            if (-not $holders.ContainsKey($holder.Sid) -or ($v -and $v -gt [version] $holders[$holder.Sid])) { $holders[$holder.Sid] = $pkg.Version }
        }
    }
    $have    = @($signedIn | Where-Object { $holders.ContainsKey($_) })
    $missing = @($signedIn | Where-Object { -not $holders.ContainsKey($_) })
    $builds  = (@($have | ForEach-Object { $holders[$_] } | Select-Object -Unique) -join ', ')

    $script:coverage[$PackageName] = ($missing.Count -eq 0)
    if ($missing.Count -eq 0) {
        Write-Ok ("    all {0} signed-in user(s) have {1} ({2}) - the failures are FSLogix's own replay; Windows registers the provisioned build at sign-in regardless, so users are not affected" -f
                  $signedIn.Count, $PackageName, $builds)
        Write-Skip '    to silence it: FSLogix Profiles\InstallAppxPackages = 0 (Microsoft''s documented workaround; it stops FSLogix replaying any AppX package) - not changed by this script'
    } else {
        Write-Bad ("    {0} of {1} signed-in user(s) do NOT have {2}: {3}" -f
                   $missing.Count, $signedIn.Count, $PackageName, (($missing | ForEach-Object { Resolve-SidName $_ }) -join ', '))
        if ($have.Count -gt 0) { Write-Skip "    the other $($have.Count) have it ($builds)" }
    }
}

function Write-FailureEvidence {
    <#
        Why a package keeps failing, in the words of the two components involved:
        the newest AppX deployment error for it, whose text carries the reason the
        bare code does not, with its ActivityId for Get-AppPackageLog; and the lines
        about it in FSLogix's own profile log.
    #>
    param([Parameter(Mandatory)] [string] $PackageName)

    $since  = (Get-Date).AddDays(-$Days)
    $latest = @(Get-EventSafe -Filter @{ LogName = 'Microsoft-Windows-AppXDeploymentServer/Operational'; Level = 2; StartTime = $since } |
                Where-Object { $_.Message -match [regex]::Escape($PackageName) }) | Select-Object -First 1
    if ($latest) {
        $text = (($latest.Message -replace '\s+', ' ')).Trim()
        if ($text.Length -gt 500) { $text = $text.Substring(0, 500) + '...' }
        Write-Skip ("    newest AppX error, {0:yyyy-MM-dd HH:mm}: {1}" -f $latest.TimeCreated, $text)
        $activity = Get-PropertyValue $latest 'ActivityId'
        if ($activity) { Write-Skip "    full trace: Get-AppPackageLog -ActivityID $activity" }
    }

    Write-UserCoverage -PackageName $PackageName

    $logDir = 'C:\ProgramData\FSLogix\Logs\Profile'
    $log    = @(Get-ChildItem -Path $logDir -Filter '*.log' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending) | Select-Object -First 1
    if ($log) {
        $lines = @(Select-String -LiteralPath $log.FullName -Pattern ([regex]::Escape($PackageName)) -ErrorAction SilentlyContinue |
                   Select-Object -Last 4)
        foreach ($line in $lines) {
            $text = $line.Line.Trim()
            if ($text.Length -gt 300) { $text = $text.Substring(0, 300) + '...' }
            Write-Skip "    FSLogix log ($($log.Name)): $text"
        }
        if ($lines.Count -eq 0) { Write-Skip "    FSLogix log $($log.FullName) says nothing about $PackageName" }
    }
}

function Write-AppxPolicyStatus {
    <# Policies that refuse package installs outright, whatever the store says. #>
    $policy = Get-ItemProperty $AppxPolicyPath -ErrorAction SilentlyContinue
    if (-not $policy) { Write-Skip 'No AppX installation policies set'; return }
    if ((Get-PropertyValue $policy 'BlockNonAdminUserInstall') -eq 1) {
        Write-Warn 'Policy BlockNonAdminUserInstall=1 - users cannot register packages themselves; provisioning for all users still works'
    }
    if ((Get-PropertyValue $policy 'AllowAllTrustedApps') -eq 0) {
        Write-Warn 'Policy AllowAllTrustedApps=0 - sideloading is off, which can refuse an MSIX from -Source'
    }
}

function Invoke-KnownInstaller {
    <#
        Provision a package for all users with Microsoft's own installer, then read
        the provisioned packages back instead of trusting the exit code alone.
    #>
    param([Parameter(Mandatory)] [string] $PackageName)

    $spec = $KnownInstallers[$PackageName]
    if (-not (Test-Path $WorkingDir)) { New-Item -ItemType Directory -Path $WorkingDir -Force | Out-Null }
    $exe = Join-Path $WorkingDir $spec.File

    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    $ProgressPreference = 'SilentlyContinue'
    Invoke-WebRequest -Uri $spec.Url -OutFile $exe -UseBasicParsing
    if (-not $SkipSignatureCheck -and -not (Test-MicrosoftSignature -Path $exe)) {
        throw "$($spec.File) is not validly signed by Microsoft - refusing to run it"
    }

    $before = Get-ProvisionedVersion
    $proc   = Start-Process -FilePath $exe -ArgumentList $spec.Args -PassThru -WindowStyle Hidden
    if (-not $proc.WaitForExit(900 * 1000)) {
        try { $proc.Kill() } catch { }
        throw "$($spec.File) $($spec.Args) did not finish within 15 minutes"
    }

    $after = Get-ProvisionedVersion
    if ($after.ContainsKey($PackageName)) {
        $was = if ($before.ContainsKey($PackageName)) { $before[$PackageName] } else { 'nothing' }
        Write-Ok "$PackageName provisioned for all users: $($after[$PackageName]) (was $was, exit code $($proc.ExitCode))"
        return $true
    }
    # Copilot can land as the unified app through Edge Update instead of as AppX.
    if ($spec.ContainsKey('EdgeUpdateId') -and $spec.EdgeUpdateId) {
        $unified = Get-EdgeUpdateClientVersion -AppId $spec.EdgeUpdateId
        if ($unified) {
            Write-Ok "$PackageName is installed machine-wide as the unified Microsoft Copilot app $unified (Edge Update, exit code $($proc.ExitCode))"
            return $true
        }
    }
    Write-Bad "$($spec.File) $($spec.Args) exited with $($proc.ExitCode) and $PackageName is still not provisioned"
    return $false
}

# -- Copilot -------------------------------------------------------------------
function Get-EdgeUpdateClientVersion {
    <# The version Edge Update has installed for an app id, or $null. #>
    param([Parameter(Mandatory)] [string] $AppId)
    $key = Get-ItemProperty "$EdgeUpdateStatePath\Clients\$AppId" -ErrorAction SilentlyContinue
    return Get-PropertyValue $key 'pv'
}

function Install-UnifiedCopilot {
    <#
        The new, unified Microsoft Copilot app, the way Microsoft documents putting it
        on a machine: Edge Update's Install policy for the Copilot app id set to
        Force Installs (5, machine-wide - which is what a session host needs), with
        UpdaterExperimentationAndConfigurationServiceControl = 1, which Force Installs
        requires, and CopilotUnificationAllowed = 1 for machines that already carry
        one of the old apps. Then Edge Update is asked to check now, and this waits
        for the app to appear under its Clients key.

        These are written as local policy, after a .reg backup of the key. A policy
        that forbids the install (Install = 0) is not overridden: somebody decided
        that, in a GPO or Intune, and a local value would be undone at the next
        refresh anyway. Returns the installed version, or $null.
    #>
    param([string] $BackupFolder, [int] $WaitMinutes = 10)

    $policy = Get-ItemProperty $EdgeUpdatePolicyPath -ErrorAction SilentlyContinue
    if ((Get-PropertyValue $policy "Install$CopilotEdgeUpdateId") -eq 0) {
        Write-Bad "Policy Install$CopilotEdgeUpdateId = 0 forbids the Copilot install - not overridden; change it in the GPO or Intune profile that sets it"
        return $null
    }

    $updater = Join-Path ${env:ProgramFiles(x86)} 'Microsoft\EdgeUpdate\MicrosoftEdgeUpdate.exe'
    if (-not (Test-Path $updater)) {
        Write-Warn 'Edge Update is not installed, so the unified Copilot app cannot come through it'
        return $null
    }

    if ((Test-Path $EdgeUpdatePolicyPath) -and $BackupFolder) {
        if (Backup-RegistryKey -Key $EdgeUpdatePolicyPath -Folder $BackupFolder) { Write-Skip "  Backed up $EdgeUpdatePolicyPath to $BackupFolder" }
        else { Write-Warn "  Could not back up $EdgeUpdatePolicyPath - continuing, the values written are listed below" }
    }
    if (-not (Test-Path $EdgeUpdatePolicyPath)) { New-Item -Path $EdgeUpdatePolicyPath -Force | Out-Null }

    foreach ($value in @(
        @{ Name = "Install$CopilotEdgeUpdateId";                   Data = 5; Why = 'Force Installs, machine-wide' }
        @{ Name = 'UpdaterExperimentationAndConfigurationServiceControl'; Data = 1; Why = 'required for Force Installs' }
        @{ Name = "CopilotUnificationAllowed$CopilotEdgeUpdateId";  Data = 1; Why = 'moves an old Copilot app to the unified one' }
    )) {
        if ((Get-PropertyValue $policy $value.Name) -eq $value.Data) { continue }
        New-ItemProperty -Path $EdgeUpdatePolicyPath -Name $value.Name -Value $value.Data -PropertyType DWord -Force | Out-Null
        Write-Ok "  Set $($value.Name) = $($value.Data) ($($value.Why))"
    }

    if ((Get-PropertyValue (Get-ItemProperty $EdgeUpdateStatePath -ErrorAction SilentlyContinue) 'PauseCopilotAppUnificationRollout') -eq 1) {
        Write-Warn "  PauseCopilotAppUnificationRollout = 1 is set under $EdgeUpdateStatePath - left alone; it can hold the unified app back"
    }

    # Edge Update checks by itself on a schedule; asking now saves waiting for it.
    # Its own machine task first, the updater's documented /ua check otherwise.
    $task = @(Get-ScheduledTask -TaskName 'MicrosoftEdgeUpdateTaskMachineUA*' -ErrorAction SilentlyContinue) | Select-Object -First 1
    if ($task) {
        Start-ScheduledTask -InputObject $task
        Write-Skip "  Started the Edge Update task $($task.TaskName)"
    } else {
        Start-Process -FilePath $updater -ArgumentList '/ua /installsource scheduler' -WindowStyle Hidden | Out-Null
        Write-Skip '  Started an Edge Update check (MicrosoftEdgeUpdate.exe /ua)'
    }

    $deadline = (Get-Date).AddMinutes($WaitMinutes)
    while ((Get-Date) -lt $deadline) {
        $version = Get-EdgeUpdateClientVersion -AppId $CopilotEdgeUpdateId
        if ($version) {
            Write-Ok "Unified Microsoft Copilot app $version installed machine-wide by Edge Update"
            return $version
        }
        Start-Sleep -Seconds 15
    }
    Write-Warn "  The unified Copilot app did not appear within $WaitMinutes minutes - the policy is in place, so Edge Update installs it at its next check"
    return $null
}

function Write-CopilotStatus {
    <#
        Where Copilot stands on this machine and what keeps it away. Three places:
        the two AppX packages it used to be, the unified app Edge Update installs,
        and the policies - Edge Update's own and Windows' - that remove it or refuse
        the install. Policies are reported with their path and value, never changed:
        they come from a GPO or Intune, and a local edit would be undone at the next
        refresh. Returns $true when something blocks Copilot.
    #>
    $blocked = $false

    $installed = @(Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue | Where-Object { $_.Name -in $CopilotPackages })
    $provisioned = Get-ProvisionedVersion
    foreach ($pkgName in $CopilotPackages) {
        $mine  = @($installed | Where-Object { $_.Name -eq $pkgName })
        $users = @($mine | ForEach-Object { @(Get-AppxPackageHolder -Package $_) } | Select-Object -ExpandProperty Sid -Unique).Count
        $prov  = if ($provisioned.ContainsKey($pkgName)) { "provisioned $($provisioned[$pkgName])" } else { 'not provisioned' }
        if ($mine.Count -gt 0) {
            $newest = ($mine | Sort-Object { [version] $_.Version } -Descending | Select-Object -First 1).Version
            Write-Ok "$pkgName $newest - registered for $users user(s), $prov"
        } else {
            Write-Skip "$pkgName - not registered for anyone, $prov"
        }
    }

    $unified = Get-EdgeUpdateClientVersion -AppId $CopilotEdgeUpdateId
    if ($unified) { Write-Ok "Unified Microsoft Copilot app $unified (installed by Edge Update)" }
    else          { Write-Skip 'Unified Microsoft Copilot app (Edge Update) is not installed' }

    $updater = Join-Path ${env:ProgramFiles(x86)} 'Microsoft\EdgeUpdate\MicrosoftEdgeUpdate.exe'
    if (Test-Path $updater) {
        $v = (Get-Item $updater).VersionInfo
        $updaterVersion = [version] ('{0}.{1}.{2}.{3}' -f $v.FileMajorPart, $v.FileMinorPart, $v.FileBuildPart, $v.FilePrivatePart)
        if ($updaterVersion -lt $EdgeUpdateMinimum) {
            Write-Warn "Edge Update $updaterVersion is older than $EdgeUpdateMinimum, the first that can install the unified Copilot app"
        }
    } else {
        Write-Warn 'Edge Update is not installed - the unified Copilot app is delivered through it'
    }

    # Edge Update policies for the Copilot app id.
    $policy  = Get-ItemProperty $EdgeUpdatePolicyPath -ErrorAction SilentlyContinue
    $install = Get-PropertyValue $policy "Install$CopilotEdgeUpdateId"
    $remove  = Get-PropertyValue $policy "Uninstall$CopilotEdgeUpdateId"
    $update  = Get-PropertyValue $policy "Update$CopilotEdgeUpdateId"
    $where   = "$EdgeUpdatePolicyPath (a GPO or Intune setting - change it there, a local edit is undone at the next refresh)"
    if ($install -eq 0) {
        $blocked = $true
        Write-Bad "Policy Install$CopilotEdgeUpdateId = 0 - Copilot may not be installed through Edge Update. $where"
    } elseif ($install -eq 5) {
        Write-Ok 'Policy Install = 5 (Force Installs, machine-wide) - Edge Update puts Copilot on this machine itself'
    }
    if ($remove -in @(1, 2)) {
        # Microsoft: Force Installs (5) overrides the Uninstall policy.
        if ($install -eq 5) {
            Write-Skip "Policy Uninstall$CopilotEdgeUpdateId = $remove is set as well, but Install = 5 overrides it"
        } else {
            $blocked = $true
            Write-Bad "Policy Uninstall$CopilotEdgeUpdateId = $remove - Edge Update removes Copilot at every check$(if ($remove -eq 2) { ', user data included' }). $where"
        }
    }
    if ($update -eq 0) {
        Write-Warn "Policy Update$CopilotEdgeUpdateId = 0 - Copilot is never updated. $where"
    }
    if (($install -in @(1, 5)) -and (Get-PropertyValue $policy 'UpdaterExperimentationAndConfigurationServiceControl') -ne 1) {
        Write-Warn "Policy Install = $install only works with UpdaterExperimentationAndConfigurationServiceControl = 1, which is not set. $where"
    }
    if ((Get-PropertyValue (Get-ItemProperty $EdgeUpdateStatePath -ErrorAction SilentlyContinue) 'PauseCopilotAppUnificationRollout') -eq 1) {
        Write-Warn "PauseCopilotAppUnificationRollout = 1 under $EdgeUpdateStatePath - the move to the unified Copilot app is paused on this machine"
    }

    # Windows' own Copilot policies, machine-wide and for every signed-in user. Every
    # value is shown as it is rather than interpreted: the names have changed more
    # than once and a wrong reading would send someone after the wrong setting.
    $roots = @('HKLM:\SOFTWARE\Policies\Microsoft\Windows') +
             @(Get-ChildItem 'Registry::HKEY_USERS' -ErrorAction SilentlyContinue |
               Where-Object { $_.PSChildName -match '^S-1-(5-21|12-1)-[\d-]+$' } |
               ForEach-Object { "Registry::HKEY_USERS\$($_.PSChildName)\SOFTWARE\Policies\Microsoft\Windows" })
    foreach ($root in $roots) {
        foreach ($leaf in 'WindowsCopilot', 'WindowsAI') {
            $values = Get-ItemProperty "$root\$leaf" -ErrorAction SilentlyContinue
            if (-not $values) { continue }
            foreach ($property in $values.PSObject.Properties) {
                if ($property.Name -like 'PS*') { continue }
                if ($leaf -eq 'WindowsAI' -and $property.Name -notmatch 'Copilot') { continue }
                $display = ($root -replace '^Registry::HKEY_USERS', 'HKU') + "\$leaf\$($property.Name) = $($property.Value)"
                if ($property.Name -match '^(TurnOff|Remove|Disable)' -and $property.Value -eq 1) {
                    $blocked = $true
                    Write-Bad "Policy $display - this turns Copilot off; change it in the GPO or Intune profile that sets it"
                } else {
                    Write-Skip "Policy $display"
                }
            }
        }
    }
    return $blocked
}

# -- Older builds --------------------------------------------------------------
function Remove-OlderBuild {
    <#
        -RemoveOld: every reference this host keeps to an older build of one package,
        once the build to keep is provisioned - so a sign-in can only ever land on
        the current one. In this order, each read back:

          provisioned   older provisioned copies (Remove-AppxProvisionedPackage)
          registered    older builds registered for any user (Remove-AppxPackage
                        -AllUsers, per user where that refuses)
          store         what AppxAllUserStore still remembers of older builds -
                        user, end-of-life, deferred-removal and machine entries -
                        backed up to .reg first, like step 5

        Not touched: the files under WindowsApps - TrustedInstaller owns them and
        Windows deletes them itself once nothing references them - and the list in
        each user's profile container (AppxPackages.xml), which FSLogix rewrites at
        the user's next sign-out. Both are reported.
    #>
    param([Parameter(Mandatory)] [string] $PackageName, [string] $BackupFolder)

    $keep = (Get-ProvisionedVersion)[$PackageName]
    if (-not $keep) {
        Write-Warn "  $PackageName is not provisioned on this host - nothing is removed, so users are never left without it"
        return
    }
    Write-Skip "  $PackageName - keeping $keep, removing every older build"
    $removed = 0

    foreach ($prov in @(Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -eq $PackageName })) {
        $v = try { [version] $prov.Version } catch { $null }
        if (-not $v -or $v -ge $keep) { continue }
        if (-not $PSCmdlet.ShouldProcess($prov.PackageName, 'Remove-AppxProvisionedPackage -Online (older build)')) { continue }
        try {
            Remove-AppxProvisionedPackage -Online -PackageName $prov.PackageName -ErrorAction Stop | Out-Null
            Write-Ok "  Deprovisioned the older $($prov.PackageName)"; $removed++
        } catch { Write-Warn "  Could not deprovision $($prov.PackageName): $($_.Exception.Message)" }
    }

    foreach ($pkg in @(Get-AppxPackage -AllUsers -Name $PackageName -ErrorAction SilentlyContinue)) {
        $v = try { [version] $pkg.Version } catch { $null }
        if (-not $v -or $v -ge $keep) { continue }
        if (-not $PSCmdlet.ShouldProcess($pkg.PackageFullName, 'Remove-AppxPackage -AllUsers (older build)')) { continue }
        try {
            Remove-AppxPackage -Package $pkg.PackageFullName -AllUsers -ErrorAction Stop
            Write-Ok "  Removed the older $($pkg.PackageFullName) for all users"; $removed++
        } catch {
            Write-Warn "  Removing $($pkg.PackageFullName) for all users failed: $($_.Exception.Message)"
            foreach ($holder in @(Get-AppxPackageHolder -Package $pkg)) {
                try {
                    Remove-AppxPackage -Package $pkg.PackageFullName -User $holder.Sid -ErrorAction Stop
                    Write-Ok "    Removed it for $($holder.Account)"; $removed++
                } catch {
                    Write-Warn "    Still there for $($holder.Account): $($_.Exception.Message)"
                }
            }
        }
    }

    # What the registry still remembers of older builds, after the cmdlets above.
    $roots = @($AppxAllUserStorePath, "$AppxAllUserStorePath\EndOfLife", "$AppxAllUserStorePath\DeferredRemoval")
    $keys  = @(foreach ($root in $roots) {
        Get-ChildItem $root -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -like 'S-1-*' } |
            ForEach-Object { Get-ChildItem $_.PSPath -ErrorAction SilentlyContinue }
    }) + @(Get-ChildItem "$AppxAllUserStorePath\Applications" -ErrorAction SilentlyContinue)
    foreach ($key in @($keys | Where-Object { $_.PSChildName -like "${PackageName}_*" })) {
        $v = Get-PackageVersionFromFullName $key.PSChildName
        if (-not $v -or $v -ge $keep) { continue }
        if (-not $PSCmdlet.ShouldProcess($key.Name, 'Back up and remove the store entry of an older build')) { continue }
        if ($BackupFolder -and -not (Backup-RegistryKey -Key $key.Name -Folder $BackupFolder)) {
            Write-Warn "  Could not back up $($key.Name) - left in place"
            continue
        }
        Remove-Item -Path $key.PSPath -Recurse -Force -ErrorAction SilentlyContinue
        if (Test-Path $key.PSPath) { Write-Warn "  Could not remove $($key.Name)" }
        else { Write-Ok "  Removed the store entry $($key.Name)"; $removed++; $script:storeEdited = $true }
    }

    if ($removed -eq 0) { Write-Ok "  No older build of $PackageName referenced on this host" }

    $folders = @(Get-ChildItem (Join-Path $env:ProgramFiles 'WindowsApps') -Directory -Filter "${PackageName}_*" -ErrorAction SilentlyContinue |
                 Where-Object { ($v = Get-PackageVersionFromFullName $_.Name) -and $v -lt $keep })
    if ($folders.Count -gt 0) {
        Write-Skip ("  {0} older folder(s) left in WindowsApps ({1}) - Windows deletes them once nothing references them; they are not touched here" -f
                    $folders.Count, (($folders | ForEach-Object { Get-PackageVersionFromFullName $_.Name }) -join ', '))
    }
    Write-Skip '  Users whose profile still lists an older build get it rewritten by FSLogix at their next sign-out'
}

# -- winget --------------------------------------------------------------------
function Get-WingetPath {
    <#
        winget as a path rather than a PATH lookup: its alias only exists per user,
        so a run as System finds nothing on PATH. The App Installer's own folder
        under WindowsApps is reachable either way; newest version wins.
    #>
    $command = Get-Command 'winget.exe' -ErrorAction SilentlyContinue
    if ($command -and $command.Source -and (Test-Path $command.Source)) { return $command.Source }

    $folders = @(Get-ChildItem (Join-Path $env:ProgramFiles 'WindowsApps') -Directory -ErrorAction SilentlyContinue `
                               -Filter 'Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe' |
                 Sort-Object -Property @{ Expression = {
                     $raw = $_.Name -replace '^Microsoft\.DesktopAppInstaller_', '' -replace '_x64__8wekyb3d8bbwe$', ''
                     try { [version] $raw } catch { [version] '0.0.0.0' }
                 } })
    for ($i = $folders.Count - 1; $i -ge 0; $i--) {
        $exe = Join-Path $folders[$i].FullName 'winget.exe'
        if (Test-Path $exe) { return $exe }
    }
    return $null
}

function Invoke-WingetProvision {
    <#
        Provision a package for all users from the MSIX winget downloads for it.
        winget checks the SHA256 its manifest publishes, this checks the Microsoft
        signature on the package and on every dependency winget brought along, and
        Add-AppxProvisionedPackage installs the lot for every profile - not only for
        the account running this, which is what "winget install" would do.
        Returns the provisioned version, or $null when it did not take.
    #>
    param([Parameter(Mandatory)] [string] $PackageId)

    $winget = Get-WingetPath
    if (-not $winget) { throw 'winget was not found - the App Installer package provides it' }

    # Its own folder, emptied first, so a previous attempt's file cannot be picked up.
    $staging = Join-Path $WorkingDir ('winget_' + ($PackageId -replace '[^\w\.]', '_'))
    if (Test-Path $staging) { Remove-Item -LiteralPath $staging -Recurse -Force -ErrorAction SilentlyContinue }
    New-Item -ItemType Directory -Path $staging -Force | Out-Null

    $output = & $winget 'download' '--id' $PackageId '--exact' '--source' 'winget' '--download-directory' $staging `
                        '--accept-package-agreements' '--accept-source-agreements' '--disable-interactivity' 2>&1
    foreach ($line in @($output)) {
        $text = "$line".Trim()
        # Progress bars are dropped: spinner characters and block elements (U+2580-259F),
        # built from char codes because Windows PowerShell reads this file as ANSI.
        $progress = '^[\s\-\\|/' + [char] 0x2580 + '-' + [char] 0x259F + ']+$'
        if ($text -and $text -notmatch $progress) { Write-Skip "  winget: $text" }
    }
    if ($LASTEXITCODE -ne 0) { throw "winget download --id $PackageId failed (exit code $LASTEXITCODE)" }

    $packages = @(Get-ChildItem -LiteralPath $staging -File -Recurse -ErrorAction SilentlyContinue |
                  Where-Object { $_.Extension -in @('.msix', '.msixbundle', '.appx', '.appxbundle') })
    # The package itself is the largest; everything else winget fetched is a dependency.
    $main = $packages | Sort-Object Length -Descending | Select-Object -First 1
    if (-not $main) { throw "winget downloaded no MSIX for $PackageId - its installer is not a package that can be provisioned" }
    $deps = @($packages | Where-Object { $_.FullName -ne $main.FullName })

    if (-not $SkipSignatureCheck) {
        foreach ($file in @($main) + $deps) {
            if (-not (Test-MicrosoftSignature -Path $file.FullName)) {
                throw "$($file.Name) is not validly signed by Microsoft - refusing to provision it"
            }
        }
    }

    $before = Get-ProvisionedVersion
    $params = @{ Online = $true; PackagePath = $main.FullName; SkipLicense = $true; ErrorAction = 'Stop' }
    if ($deps.Count -gt 0) { $params['DependencyPackagePath'] = @($deps.FullName) }
    Add-AppxProvisionedPackage @params | Out-Null

    # Which package that was is read back from the store rather than assumed from
    # the winget id: the two names are not the same (Microsoft.Outlook provisions
    # Microsoft.OutlookForWindows).
    $after   = Get-ProvisionedVersion
    $changed = @($after.Keys | Where-Object { -not $before.ContainsKey($_) -or $before[$_] -ne $after[$_] })
    $mainName = @($changed | Where-Object { $_ -notmatch 'VCLibs|UI\.Xaml|WindowsAppRuntime|NET\.Native' }) + $changed |
                Select-Object -First 1
    if (-not $mainName) {
        Write-Warn "  winget $PackageId was provisioned, but no provisioned package changed version - it was already there at that build"
        return $null
    }
    Write-Ok "$mainName provisioned for all users from winget ($PackageId): $($after[$mainName])"
    return $after[$mainName]
}

# -- Everything that fails -----------------------------------------------------
function Get-FailingApp {
    <#
        Every package that failed to deploy in the last -Days days, from both places
        that log it: the AppX deployment log (any install, update or registration)
        and the FSLogix Apps log (the replay at sign-in). One object per package name,
        with the versions asked for, the error codes, and whether this host has them.
    #>
    $since = (Get-Date).AddDays(-$Days)
    $rows  = [System.Collections.Generic.List[object]]::new()

    foreach ($source in @(
        @{ Label = 'AppX'; Filter = @{ LogName = 'Microsoft-Windows-AppXDeploymentServer/Operational'; Level = 2; StartTime = $since } }
        @{ Label = 'FSLogix'; Filter = @{ ProviderName = 'Microsoft-FSLogix-Apps'; Level = 2; StartTime = $since } }
    )) {
        foreach ($entry in @(Get-EventSafe -Filter $source.Filter)) {
            if ($entry.Message -notmatch '([\w\.\-]+_\d+\.\d+\.\d+\.\d+_\w*_[\w\.\-]*_\w{13})') { continue }
            $fullName = $Matches[1]
            if (-not (Test-NameInScope $fullName)) { continue }
            $rows.Add([PSCustomObject]@{
                Source   = $source.Label
                FullName = $fullName
                Name     = ($fullName -split '_')[0]
                Code     = if ($entry.Message -match '(0x8[0-9A-Fa-f]{7})') { ($Matches[1].ToUpper() -replace '^0X', '0x') } else { '?' }
                Time     = $entry.TimeCreated
            })
        }
    }

    foreach ($group in @($rows | Group-Object Name | Sort-Object Count -Descending)) {
        [PSCustomObject]@{
            Name     = $group.Name
            Count    = $group.Count
            Last     = ($group.Group | Sort-Object Time -Descending | Select-Object -First 1).Time
            Sources  = (@($group.Group.Source | Select-Object -Unique) -join '+')
            Codes    = (@($group.Group | Group-Object Code | Sort-Object Count -Descending |
                          ForEach-Object { "$($_.Name) x$($_.Count)" }) -join ', ')
            Versions = @($group.Group.FullName | Select-Object -Unique)
        }
    }
}

# What a code means for this script, on top of what Windows calls it.
$AppxErrorText = @{
    '0x80070490' = 'the store or FSLogix asks for something it cannot find - this script repairs that'
    '0x80073D02' = 'retries when the app is closed'
    '0x80073D06' = 'harmless, the newer version stays'
    '0x80073D19' = 'harmless, the user signed out during registration'
}

function Get-HResultText {
    <#
        Windows' own text for an AppX HRESULT - every 0x8007xxxx is a Win32 code,
        and Windows has the message - plus this script's note where it has one.
        Measured: 0x80073D19 reads "An error occurred because a user was logged off".
    #>
    param([string] $Code)

    $text = $null
    if ($Code -match '^0x8007([0-9A-Fa-f]{4})$') {
        $text = ([ComponentModel.Win32Exception]::new([Convert]::ToInt32($Matches[1], 16)).Message -replace '\s+', ' ').Trim().TrimEnd('.')
        if ($text -match '^Unknown error') { $text = $null }
    }
    $note = if ($AppxErrorText.ContainsKey($Code)) { $AppxErrorText[$Code] } else { $null }
    return (@($text, $note) | Where-Object { $_ }) -join ' - '
}

function Write-FailingApp {
    <# The overview, one block per package, worst first; -Top keeps it readable. #>
    param([hashtable] $Provisioned, [int] $Top = 15)

    $failing = @(Get-FailingApp)
    if ($failing.Count -eq 0) {
        Write-Ok "No package failed to deploy in the last $Days day(s)"
        return
    }

    # 0x80070490 first: that is what this script can repair.
    $failing = @($failing | Sort-Object @{ Expression = { $_.Codes -notmatch '0x80070490' } }, @{ Expression = 'Count'; Descending = $true })
    Write-Warn "$($failing.Count) package(s) failed to deploy in the last $Days day(s):"
    foreach ($app in @($failing | Select-Object -First $Top)) {
        Write-Warn ("  {0}: {1}x ({2}), last {3:yyyy-MM-dd HH:mm}" -f $app.Name, $app.Count, $app.Sources, $app.Last)
        # One code per line: three codes with their meaning on one line wrapped
        # into something nobody reads.
        foreach ($part in @($app.Codes -split ', ')) {
            $meaning = Get-HResultText (($part -split ' ')[0])
            Write-Skip ("    {0}{1}" -f $part, $(if ($meaning) { " - $meaning" } else { '' }))
        }
        foreach ($version in $app.Versions) {
            $onDisk = if (Test-PackageFiles $version) { 'files on this host' } else { 'no files on this host' }
            Write-Skip "    $version ($onDisk)"
        }
        if ($Provisioned.ContainsKey($app.Name)) { Write-Skip "    provisioned here: $($Provisioned[$app.Name])" }
        # The reason, for what is still failing today - the code alone is not one.
        if ($app.Last -gt (Get-Date).AddHours(-24)) { Write-FailureEvidence -PackageName $app.Name }
    }
    if ($failing.Count -gt $Top) {
        Write-Skip "  ... and $($failing.Count - $Top) more - narrow it down with -Name"
    }
}

# -- Main ----------------------------------------------------------------------
try {
    Write-Out ''
    Write-Step ("AppX package store - {0}" -f ($Name -join ', '))

    if ($Source -and -not (Test-Path -LiteralPath $Source)) { throw "Source not found: $Source" }

    # -- 1. Diagnose -----------------------------------------------------------
    Write-Out ''
    Write-Step '1. Diagnose'
    $health = Get-StoreHealth
    Write-StoreHealth $health
    $work = Get-RepairableCount $health
    if ($work -eq 0) { Write-Ok 'Nothing broken in the package store for these packages' }

    # -- 1b. FSLogix and policy ------------------------------------------------
    # On a pooled host the store is often fine and the error comes from FSLogix
    # replaying a version this host never had. That is fixed by provisioning, not by
    # editing the store, so it is diagnosed separately.
    Write-Out ''
    Write-Step '1b. FSLogix and installation policy'
    $provisionedNow = Get-ProvisionedVersion
    $requests       = @(Get-FslogixRequest)
    $needs          = @(Write-FslogixStatus -Provisioned $provisionedNow -Requests $requests)
    # The exact builds FSLogix fails on, where Microsoft's CDN serves that build.
    # They win over the installer for the same package: the installer's build is
    # older, and older is what fails.
    $exactTargets   = @(Get-ExactTarget -Provisioned $provisionedNow -Requests $requests)
    if ($Latest) {
        # The newest build replaces the one FSLogix asked for when it is newer.
        foreach ($newest in @(Get-LatestTarget -Provisioned $provisionedNow -Requests $requests)) {
            $same = @($exactTargets | Where-Object { $_.Name -eq $newest.Name })
            if ($same.Count -eq 0 -or $same[0].Version -lt $newest.Version) {
                $exactTargets = @($exactTargets | Where-Object { $_.Name -ne $newest.Name }) + $newest
            }
        }
    }
    Write-AppxPolicyStatus

    # -- 1c. Everything that fails --------------------------------------------
    # Not only what the store gets wrong: every package that failed to install,
    # update or register lately, from the AppX and FSLogix logs together.
    Write-Out ''
    Write-Step '1c. Failing apps'
    Write-FailingApp -Provisioned $provisionedNow
    # A failure every signed-in user is unaffected by is noise, not a fault.
    $script:unexplained = @($script:unexplained | Where-Object { -not ($script:coverage.ContainsKey($_) -and $script:coverage[$_]) })

    # -- 1d. Copilot -----------------------------------------------------------
    # Only when asked for: with -Copilot, or a Copilot package named explicitly.
    $copilotBlocked = $false
    $copilotMissing = $false
    $copilotAsked   =$Copilot -or ($explicitName -and @($CopilotPackages | Where-Object { Test-NameInScope $_ }).Count -gt 0)
    if ($copilotAsked) {
        Write-Out ''
        Write-Step '1d. Copilot'
        $copilotBlocked = [bool] (Write-CopilotStatus)
        if ($copilotBlocked) { $work++ }
        # What counts is the new, unified app installed machine-wide: the old
        # Microsoft 365 Copilot app is what the unification replaces, and a package
        # registered for some users is not what a first sign-in on a session host gets.
        $copilotMissing = -not (Get-EdgeUpdateClientVersion -AppId $CopilotEdgeUpdateId)
        if ($copilotMissing) {
            Write-Warn 'The new Microsoft Copilot app is not installed on this machine - -Copilot -Provision installs it machine-wide through Edge Update'
            $work++
        }
    }

    # What -Provision installs: what FSLogix showed is missing or behind, plus any
    # known package named explicitly (a deliberate "bring this to the current build").
    $provisionTargets = @($needs)
    if ($Provision -and $explicitName) {
        # Only what is not provisioned at all. The installers deliver an older
        # last-known-good build (Outlook 818 while 915 was provisioned), so running
        # them over a provisioned package is a downgrade - which a full end-to-end
        # run of -Name teams,outlook -Provision showed it would have done.
        # Copilot is the exception: its state lives in Edge Update, not here.
        $provisionTargets += @($KnownInstallers.Keys | Where-Object {
            (Test-NameInScope $_) -and (-not $provisionedNow.ContainsKey($_) -or $_ -eq 'Microsoft.MicrosoftOfficeHub')
        })
    }
    # Not @($exactTargets.Name): under Set-StrictMode in Windows PowerShell 5.1 that
    # throws "The property 'Name' cannot be found" when the list is empty - which
    # aborted a live run on a host that already had the newest build.
    $exactNames       = @($exactTargets | ForEach-Object { $_.Name })
    $provisionTargets = @($provisionTargets | Where-Object { $KnownInstallers.ContainsKey($_) -and $_ -notin $exactNames } | Select-Object -Unique)
    $work += @($exactTargets | Where-Object { $_.Name -notin $needs }).Count
    foreach ($missing in @($needs | Where-Object { -not $KnownInstallers.ContainsKey($_) })) {
        Write-Warn "No known installer for $missing - provision it with -Source <msix>"
    }

    $work += $needs.Count
    # -RemoveOld only ever acts on packages named one by one: "every older build of
    # everything" is not a request anyone should be able to make by accident.
    $removeOldTargets = @()
    if ($RemoveOld) {
        if (-not $explicitName) {
            Write-Warn '-RemoveOld needs the packages named, e.g. -Name teams,outlook - ignored for a wildcard'
        } else {
            $removeOldTargets = @($Name | Where-Object { $_ -notin $CopilotPackages })
        }
    }

    if ($work -eq 0 -and $removeOldTargets.Count -eq 0 -and -not $Source -and $WingetId.Count -eq 0 -and -not ($Provision -and ($provisionTargets.Count + $exactTargets.Count) -gt 0)) {
        # "Nothing to repair" while a package failed today is not the same as healthy.
        if ($script:unexplained.Count -gt 0) {
            $plannedExit = 1
            throw ("Nothing this script can repair, but {0} still failed in the last 24 hours with the right build provisioned - the evidence in step 1c is the next lead" -f ($script:unexplained -join ', '))
        }
        $plannedExit = 0
        throw 'Nothing to repair.'
    }
    if ($CheckOnly) {
        $plannedExit = 2
        throw "Check only - $work thing(s) to repair, nothing was changed."
    }

    if (-not $simulate) {
        if (-not (Test-Path $LogPath)) { New-Item -ItemType Directory -Path $LogPath -Force | Out-Null }
        $stamp     = Get-Date -Format 'yyyyMMdd_HHmmss'
        $backupDir = Join-Path $LogPath "Repair-AppxPackageStore_$stamp"
        # A transcript is worth having, not worth aborting a repair over - some
        # remote and RMM hosts refuse it.
        try {
            Start-Transcript -Path (Join-Path $LogPath "Repair-AppxPackageStore_$stamp.log") | Out-Null
            $transcribing = $true
        } catch {
            Write-Warn "No transcript: $($_.Exception.Message)"
        }
    } else {
        $backupDir = Join-Path $LogPath 'Repair-AppxPackageStore_WhatIf'
    }

    # -- 2. Provisioned copies without files -----------------------------------
    Write-Out ''
    Write-Step '2. Provisioned packages'
    if ($health.Provisioned.Count -eq 0) { Write-Skip 'None without files' }
    foreach ($prov in $health.Provisioned) {
        if (-not $PSCmdlet.ShouldProcess($prov.PackageName, 'Remove-AppxProvisionedPackage -Online')) { continue }
        try {
            Remove-AppxProvisionedPackage -Online -PackageName $prov.PackageName -ErrorAction Stop | Out-Null
            Write-Ok "Deprovisioned $($prov.PackageName)"
        } catch {
            Write-Warn "Could not deprovision $($prov.PackageName): $($_.Exception.Message)"
        }
    }

    # -- 3. Re-register what still has files ------------------------------------
    # Rebuilds the store's view of the package from its own manifest. Registers it
    # for the account running this, which on a System run is System.
    Write-Out ''
    Write-Step '3. Re-register damaged packages'
    $damaged = @($health.Installed | Where-Object { $_.Problem -eq 'Damaged' -and -not $_.Protected })
    if ($damaged.Count -eq 0) { Write-Skip 'None with files on disk and a status other than Ok' }
    foreach ($item in $damaged) {
        $manifest = Join-Path $item.Location 'AppxManifest.xml'
        if (-not (Test-Path -LiteralPath $manifest)) { Write-Skip "  $($item.FullName) has no manifest to register from"; continue }
        if (-not $PSCmdlet.ShouldProcess($item.FullName, 'Add-AppxPackage -Register (from its own manifest)')) { continue }
        try {
            Add-AppxPackage -Register $manifest -DisableDevelopmentMode -ForceApplicationShutdown -ErrorAction Stop
            Write-Ok "Re-registered $($item.FullName)"
        } catch {
            # A newer version turned up after the diagnosis: nothing to re-register.
            if ($_.Exception.Message -match '0x80073D06') {
                Write-Skip "  $($item.FullName): a newer version is installed, so this one is left for Windows to remove"
            } else {
                Write-Warn "Could not re-register $($item.FullName): $($_.Exception.Message)"
            }
        }
    }

    # -- 4. Remove registrations with nothing behind them -----------------------
    Write-Out ''
    Write-Step '4. Remove ghost registrations'
    $ghosts = @($health.Installed | Where-Object { $_.Problem -eq 'Ghost' -and -not $_.Protected })
    if ($ghosts.Count -eq 0) { Write-Skip 'None' }
    foreach ($item in $ghosts) {
        if (-not $PSCmdlet.ShouldProcess($item.FullName, 'Remove-AppxPackage -AllUsers')) { continue }
        try {
            Remove-AppxPackage -Package $item.FullName -AllUsers -ErrorAction Stop
            Write-Ok "Removed $($item.FullName) for all users"
            continue
        } catch {
            Write-Warn "Remove for all users failed: $($_.Exception.Message)"
        }
        # -AllUsers is all or nothing; per user often still goes.
        foreach ($holder in $item.Holders) {
            try {
                Remove-AppxPackage -Package $item.FullName -User $holder.Sid -ErrorAction Stop
                Write-Ok "  Removed it for $($holder.Account)"
            } catch {
                if ($_.Exception.Message -match '0x80070490|Element not found') {
                    Write-Skip "  Nothing the cmdlet can find for $($holder.Account) (0x80070490) - step 5 clears the entry itself"
                } else {
                    Write-Warn "  Still held by $($holder.Account): $($_.Exception.Message)"
                }
            }
        }
    }

    # -- 5. Orphaned registry entries -------------------------------------------
    # Read again: the removals above may have taken some entries with them.
    Write-Out ''
    Write-Step '5. Orphaned package store entries'
    $orphans = @(Get-OrphanedStoreEntry)
    if ($orphans.Count -eq 0) { Write-Skip 'None left' }
    foreach ($entry in $orphans) {
        if (-not $PSCmdlet.ShouldProcess($entry.Path, "Back up and remove the orphaned $($entry.Kind)")) { continue }
        # No backup, no removal: an edit to the store that cannot be undone is not a repair.
        if (-not (Backup-RegistryKey -Key $entry.Path -Folder $backupDir)) {
            Write-Warn "Could not back up $($entry.Path) - left in place"
            continue
        }
        $psPath = 'Registry::' + $entry.Path
        Remove-Item -Path $psPath -Recurse -Force -ErrorAction SilentlyContinue
        if (Test-Path $psPath) {
            Write-Warn "Could not remove $($entry.Path)"
        } else {
            $storeEdited = $true
            Write-Ok "Removed the orphaned $($entry.Kind) for $($entry.Name)"
        }
    }
    if ($storeEdited) { Write-Skip "  Backups: $backupDir (double-click a .reg file to put an entry back)" }

    # -- 6. Provision -----------------------------------------------------------
    # After the store is clean, not before: provisioning over an orphaned entry of the
    # same package is the operation that answered 0x80070490 in the first place.
    Write-Out ''
    Write-Step '6. Provision'
    if ($Provision) {
        foreach ($target in $exactTargets) {
            if (-not $PSCmdlet.ShouldProcess("$($target.Name) $($target.Version)", "Provision for all users (exact build from $($target.Url))")) { continue }
            try {
                if (-not (Invoke-ExactProvision -Target $target)) { $exitCode = 1 }
            } catch {
                Write-Bad "Provisioning $($target.Name) $($target.Version) failed: $($_.Exception.Message)"
                Write-AppxDeploymentError -Minutes 20
                $exitCode = 1
            }
        }
        if (($provisionTargets.Count + $exactTargets.Count) -eq 0) { Write-Skip 'Nothing to provision with a known installer' }
        foreach ($target in $provisionTargets) {
            $spec = $KnownInstallers[$target]
            # Copilot has no MSIX in winget, so -UseWinget falls back to its installer.
            $viaWinget = $UseWinget -and $spec.WingetId
            if ($UseWinget -and -not $spec.WingetId) {
                Write-Skip "  winget has no MSIX for $target - using Microsoft's installer instead"
            }
            $how  = if ($viaWinget) { "winget download $($spec.WingetId) + Add-AppxProvisionedPackage" } else { "$($spec.File) $($spec.Args)" }
            if (-not $PSCmdlet.ShouldProcess($target, "Provision for all users ($how)")) { continue }
            try {
                if ($target -eq 'Microsoft.MicrosoftOfficeHub' -and $copilotAsked) {
                    # The new Copilot app comes through Edge Update. Only when that
                    # does not deliver is the Microsoft 365 Copilot installer used -
                    # the old app, which the unification then moves over.
                    $forbidden = (Get-PropertyValue (Get-ItemProperty $EdgeUpdatePolicyPath -ErrorAction SilentlyContinue) "Install$CopilotEdgeUpdateId") -eq 0
                    if (-not (Install-UnifiedCopilot -BackupFolder $backupDir)) {
                        if ($forbidden) {
                            # A forbidding policy would undo the old app just the same.
                            $exitCode = 1
                        } else {
                            Write-Skip '  Falling back to the Microsoft 365 Copilot installer'
                            if (-not (Invoke-KnownInstaller -PackageName $target)) { $exitCode = 1 }
                        }
                    }
                } elseif ($viaWinget) {
                    $got = Invoke-WingetProvision -PackageId $spec.WingetId
                    # winget's manifest lags behind Microsoft's installers; say so
                    # when it is behind what the profiles on this pool ask for.
                    $asked = @(Get-FslogixRequest | Where-Object { $_.Name -eq $target } |
                               Sort-Object Version -Descending | Select-Object -First 1 | ForEach-Object { $_.Version })
                    if ($got -and $asked.Count -gt 0 -and $got -lt $asked[0]) {
                        Write-Warn "  winget publishes $got while profiles ask for $($asked[0]) - drop -UseWinget to provision Microsoft's current build"
                    }
                } elseif (-not (Invoke-KnownInstaller -PackageName $target)) {
                    Write-AppxDeploymentError -Minutes 20
                    $exitCode = 1
                }
            } catch {
                Write-Bad "Provisioning $target failed: $($_.Exception.Message)"
                $exitCode = 1
            }
        }
    } elseif (($provisionTargets.Count + $exactTargets.Count) -gt 0) {
        $todo = @($provisionTargets) + @($exactTargets | ForEach-Object { "$($_.Name) $($_.Version)" })
        Write-Bad ("Still to do: provision {0} - run again with -Provision" -f ($todo -join ', '))
        $exitCode = 1
    }
    if ($copilotMissing -and -not $Provision) {
        Write-Bad 'Still to do: install Copilot for all users - run again with -Copilot -Provision'
        $exitCode = 1
    }
    # Any other package, by its winget id - the way to put back an app that is not
    # Teams or Outlook, as long as winget's manifest for it is an MSIX.
    foreach ($id in $WingetId) {
        if (-not $PSCmdlet.ShouldProcess($id, 'Provision for all users (winget download + Add-AppxProvisionedPackage)')) { continue }
        try {
            [void] (Invoke-WingetProvision -PackageId $id)
        } catch {
            Write-Bad "Provisioning $id failed: $($_.Exception.Message)"
            Write-AppxDeploymentError -Minutes 20
            $exitCode = 1
        }
    }

    if (-not $Source) {
        if (-not $Provision -and $WingetId.Count -eq 0) { Write-Skip 'No -Provision, -WingetId or -Source given' }
    } else {
        if (-not $SkipSignatureCheck -and -not (Test-MicrosoftSignature -Path $Source)) {
            throw "$Source is not validly signed by Microsoft - use -SkipSignatureCheck to provision it anyway"
        }
        if ($PSCmdlet.ShouldProcess($Source, 'Add-AppxProvisionedPackage -Online -SkipLicense')) {
            try {
                Add-AppxProvisionedPackage -Online -PackagePath $Source -SkipLicense -ErrorAction Stop | Out-Null
                Write-Ok "Provisioned $Source for all users"
            } catch {
                Write-Bad "Provisioning failed: $($_.Exception.Message)"
                Write-AppxDeploymentError -Minutes 10
                $exitCode = 1
            }
        }
    }

    # -- 6b. Older builds -------------------------------------------------------
    # After provisioning, never before: what is kept has to be in place first.
    if ($removeOldTargets.Count -gt 0) {
        Write-Out ''
        Write-Step '6b. Remove older builds'
        foreach ($proc in @(Get-Process -Name 'olk', 'ms-teams' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name -Unique)) {
            Write-Warn "  $proc is running - an older build in use is closed or finishes at that user's sign-out"
        }
        foreach ($pkgName in $removeOldTargets) {
            Remove-OlderBuild -PackageName $pkgName -BackupFolder $backupDir
        }
    }

    # -- 7. Verify -------------------------------------------------------------
    Write-Out ''
    Write-Step '7. Verify'
    if ($simulate) {
        Write-Skip 'Nothing was changed (-WhatIf)'
    } else {
        $after = Get-StoreHealth
        $left  = Get-RepairableCount $after
        if ($left -eq 0) {
            Write-Ok 'The package store holds nothing broken for these packages any more'
        } else {
            Write-StoreHealth $after
            Write-Bad "$left thing(s) survived the repair - the lines above say which"
            $exitCode = 1
        }

        # What FSLogix asked for, against what this host provisions now.
        $provisionedAfter = Get-ProvisionedVersion
        $requestsAfter    = @(Get-FslogixRequest)
        foreach ($target in @($exactTargets | Where-Object { $_.Name -notin $needs })) {
            $now = $provisionedAfter[$target.Name]
            if ($now -and $now -ge $target.Version) {
                Write-Ok "$($target.Name) is provisioned at $now, the build the profiles ask for"
            } elseif ($Provision) {
                Write-Bad "$($target.Name) is provisioned at $now, still older than the $($target.Version) the profiles ask for"
                $exitCode = 1
            }
        }
        foreach ($target in $needs) {
            $asked = ($requestsAfter | Where-Object { $_.Name -eq $target } | Sort-Object Version -Descending | Select-Object -First 1 | ForEach-Object { $_.Version })
            if ($provisionedAfter.ContainsKey($target) -and $asked -and $provisionedAfter[$target] -lt $asked) {
                if (Write-VersionGap -Package $target -Have $provisionedAfter[$target] -Asked $asked -Fslogix (Get-FslogixVersion)) {
                    $exitCode = 1
                }
            } elseif ($provisionedAfter.ContainsKey($target)) {
                Write-Ok "$target is provisioned: $($provisionedAfter[$target]) - users get it at their next sign-in, and FSLogix saves that version at their next sign-out"
            } else {
                Write-Bad "$target is still not provisioned on this host"
                $exitCode = 1
            }
        }
        if ($copilotAsked) {
            $officeHub = (Get-ProvisionedVersion)['Microsoft.MicrosoftOfficeHub']
            $unified   = Get-EdgeUpdateClientVersion -AppId $CopilotEdgeUpdateId
            if ($unified)       { Write-Ok "Copilot: the new Microsoft Copilot app $unified is installed machine-wide" }
            elseif ($officeHub) { Write-Warn "Copilot: only the old Microsoft 365 Copilot app ($officeHub) is provisioned - the new app follows at Edge Update's next check if the policy is in place" }
            else                { Write-Warn 'Copilot is not installed machine-wide - run with -Copilot -Provision' }
            # The one thing this script will not change: a policy that removes Copilot
            # wins over any install, so it is the answer, not a footnote.
            if ($copilotBlocked) {
                Write-Bad 'A policy removes or blocks Copilot (see step 1d) - fix it in the GPO or Intune profile that sets it; until then any install is undone'
                $exitCode = 1
            }
        }
        foreach ($pkgName in $script:unexplained) {
            Write-Warn "$pkgName was failing before this run with the right build provisioned - step 1c says whether signed-in users have it; check again after the next sign-ins"
        }
        if ($storeEdited) {
            Write-Warn 'The package store was edited - restart this host when convenient so the deployment engine rereads it from scratch'
        }
    }
} catch {
    if ($null -ne $plannedExit) {
        Write-Out ''
        Write-Skip $_.Exception.Message
        $exitCode = $plannedExit
    } else {
        Write-Out ''
        Write-Bad "Aborted: $($_.Exception.Message)"
        $exitCode = 1
    }
} finally {
    if ($transcribing) {
        try { Stop-Transcript | Out-Null } catch { Write-Warning "Transcript not closed cleanly: $($_.Exception.Message)" }
    }
}

Write-Host ''
exit $exitCode
