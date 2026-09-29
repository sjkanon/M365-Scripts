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
                      Outlook Setup.exe --provision true). With -Source, any MSIX you
                      supply. Both only after the Microsoft signature is checked.
      7. Verify       The diagnosis runs again. Exit code 0 means nothing is broken
                      any more, 1 that something survived.

    With -Name '*' (the default) system packages and framework packages are reported
    but never touched, and Deprovisioned markers - which is how bloatware removals are
    remembered - are left alone. Name the package to have those considered as well.

.PARAMETER Name
    Package names to look at, wildcards allowed. Default '*': the whole store.
    Examples: 'MSTeams', 'Microsoft.OutlookForWindows', 'MSTeams','Microsoft.OutlookForWindows'.

.PARAMETER CheckOnly
    Diagnose and report, change nothing. Exit code 2 when there is something to repair.

.PARAMETER Source
    Path to an .msix / .msixbundle / .appx(bundle) to provision once the store is
    clean. Its Authenticode signature must be valid and from Microsoft unless
    -SkipSignatureCheck is given.

.PARAMETER Provision
    Provision for all users, with Microsoft's own installer, the packages FSLogix
    failed to register that this host does not provision (or provisions older), and
    any of MSTeams / Microsoft.OutlookForWindows named explicitly in -Name. Other
    packages need -Source.

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

.EXAMPLE
    # What is broken on this host? Changes nothing.
    .\Repair-AppxPackageStore.ps1 -CheckOnly

.EXAMPLE
    # Repair Teams and new Outlook and provision both at the current build, unattended
    .\Repair-AppxPackageStore.ps1 -Name 'MSTeams','Microsoft.OutlookForWindows' -Provision -Confirm:$false

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
    [switch]   $IncludeDeprovisioned,
    [switch]   $SkipSignatureCheck,
    [ValidateRange(1, 90)]
    [int]      $Days       = 7,
    [string]   $WorkingDir = 'C:\IT\AppxRepair',
    [string]   $LogPath    = 'C:\Temp'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

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
if (-not $PSBoundParameters.ContainsKey('Days')                 -and $env:days -match '^\d+$')  { $Days    = [int] $env:days }
if (-not $PSBoundParameters.ContainsKey('WorkingDir')           -and $env:workingDir)   { $WorkingDir = $env:workingDir }
if (-not $PSBoundParameters.ContainsKey('Name')                 -and $env:packageName)  { $Name    = @($env:packageName -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }
if (-not $PSBoundParameters.ContainsKey('Source')               -and $env:source)       { $Source  = $env:source }
if (-not $PSBoundParameters.ContainsKey('LogPath')              -and $env:logPath)      { $LogPath = $env:logPath }

# powershell.exe -File passes 'MSTeams,Microsoft.OutlookForWindows' as one string -
# which is also how the relaunches above hand it over - so the list is split here.
$Name = @($Name | ForEach-Object { $_ -split '[,;]' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
if ($Name.Count -eq 0) { $Name = @('*') }

$confirmSuppressed = $PSBoundParameters.ContainsKey('Confirm') -and -not $PSBoundParameters['Confirm']
if ($confirmSuppressed) { $ConfirmPreference = 'None' }

$simulate     = [bool] $WhatIfPreference
$exitCode     = 0
$plannedExit  = $null
$transcribing = $false
$storeEdited  = $false

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
$KnownInstallers = @{
    'MSTeams' = @{
        Url  = 'https://go.microsoft.com/fwlink/?linkid=2243204&clcid=0x409'
        File = 'teamsbootstrapper.exe'
        Args = '-p'
    }
    'Microsoft.OutlookForWindows' = @{
        Url  = 'https://go.microsoft.com/fwlink/?linkid=2207851'
        File = 'OutlookSetup.exe'
        Args = '--provision true --quiet --start-'
    }
}

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

    foreach ($pkg in $Installed) {
        if (-not (Test-NameInScope $pkg.Name)) { continue }

        $location  = Get-PropertyValue $pkg 'InstallLocation'
        $filesGone = (-not $location) -or (-not (Test-Path -LiteralPath $location))
        $status    = [string] (Get-PropertyValue $pkg 'Status')
        $damaged   = (-not $filesGone) -and $status -and $status -ne 'Ok'
        if (-not ($filesGone -or $damaged)) { continue }

        $protected = ([string] (Get-PropertyValue $pkg 'SignatureKind') -eq 'System') -or
                     [bool] (Get-PropertyValue $pkg 'IsFramework') -or
                     [bool] (Get-PropertyValue $pkg 'NonRemovable')

        [PSCustomObject]@{
            Package   = $pkg
            FullName  = $pkg.PackageFullName
            Name      = $pkg.Name
            Location  = $location
            Problem   = if ($filesGone) { 'Ghost' } else { 'Damaged' }
            Reason    = if ($filesGone -and $location) { "its files are gone ($location)" }
                        elseif ($filesGone)          { 'the store holds no install location for it at all' }
                        else                         { "its status is $status" }
            Protected = $protected -and -not $explicitName
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
    $events = @(Get-WinEvent -LogName 'Microsoft-Windows-AppXDeploymentServer/Operational' `
                             -FilterXPath '*[System[(Level=2)]]' -MaxEvents 50 -ErrorAction SilentlyContinue |
                Where-Object { $_.TimeCreated -ge $since -and
                               $_.Message -match '([\w\.\-]+_[\d\.]+_\w*_[\w\.\-]*_\w{13})' -and
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
    $file = Join-Path $Folder (($Key -replace '^HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Appx\\', '' -replace '[\\/:*?"<>|]', '_') + '.reg')
    & reg.exe export $Key $file /y 2>&1 | Out-Null
    return ($LASTEXITCODE -eq 0 -and (Test-Path $file))
}

function Test-MicrosoftSignature {
    <# Valid Authenticode signature from Microsoft on a package file. #>
    param([Parameter(Mandatory)] [string] $Path)

    $sig = Get-AuthenticodeSignature -FilePath $Path
    return ($sig.Status -eq 'Valid' -and $sig.SignerCertificate.Subject -match 'O=Microsoft Corporation')
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

function Get-FslogixRequest {
    <#
        What FSLogix tried to register at sign-in and could not. At sign-out FSLogix
        writes the user's packages - full names, so exact versions - to
        AppxPackages.xml in the profile container, and at the next sign-in it
        replays that list (Profiles\InstallAppxPackages, on by default). A host that
        does not have that exact version answers 0x80070490 with an empty path,
        which is the error in the Apps event log. One object per requested package,
        with how often and when it last failed.
    #>
    $since  = (Get-Date).AddDays(-$Days)
    $events = @(Get-WinEvent -FilterHashtable @{ ProviderName = 'Microsoft-FSLogix-Apps'; Level = 2; StartTime = $since } `
                             -ErrorAction SilentlyContinue)

    $found = @{}
    foreach ($entry in $events) {
        if ($entry.Message -notmatch 'on Package\s+(\S+?)\s+from') { continue }
        $fullName = $Matches[1]
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

function Write-FslogixStatus {
    <#
        The FSLogix side: which build, whether it replays AppX packages at sign-in,
        and what it failed to register lately against what this host provisions.
        Returns the package names that need provisioning on this host.
    #>
    param([hashtable] $Provisioned, $Requests)

    $needs = @()
    if (-not (Test-Path $FslogixServicePath)) {
        Write-Skip 'FSLogix is not installed'
        return $needs
    }

    # The numeric file version, not ProductVersion: that one can carry a suffix
    # that does not parse as a version.
    $info           = (Get-Item $FslogixServicePath).VersionInfo
    $fslogixVersion = [version] ('{0}.{1}.{2}.{3}' -f $info.FileMajorPart, $info.FileMinorPart, $info.FileBuildPart, $info.FilePrivatePart)
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
            Write-Warn "  This host provisions $have, older than the $newest profiles ask for - another host in the pool is ahead. Bring every host to the same build (-Provision)"
            $needs += $group.Name
        } else {
            Write-Ok "  This host provisions $have - Windows registers that at sign-in; the error is FSLogix replaying an old saved version and stops once each user has signed out once on a current host"
        }
    }
    return $needs
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
    Write-Bad "$($spec.File) $($spec.Args) exited with $($proc.ExitCode) and $PackageName is still not provisioned"
    return $false
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
    Write-AppxDeploymentError

    # -- 1b. FSLogix and policy ------------------------------------------------
    # On a pooled host the store is often fine and the error comes from FSLogix
    # replaying a version this host never had. That is fixed by provisioning, not by
    # editing the store, so it is diagnosed separately.
    Write-Out ''
    Write-Step '1b. FSLogix and installation policy'
    $provisionedNow = Get-ProvisionedVersion
    $needs = @(Write-FslogixStatus -Provisioned $provisionedNow -Requests @(Get-FslogixRequest))
    Write-AppxPolicyStatus

    # What -Provision installs: what FSLogix showed is missing or behind, plus any
    # known package named explicitly (a deliberate "bring this to the current build").
    $provisionTargets = @($needs)
    if ($Provision -and $explicitName) {
        $provisionTargets += @($KnownInstallers.Keys | Where-Object { Test-NameInScope $_ })
    }
    $provisionTargets = @($provisionTargets | Where-Object { $KnownInstallers.ContainsKey($_) } | Select-Object -Unique)
    foreach ($missing in @($needs | Where-Object { -not $KnownInstallers.ContainsKey($_) })) {
        Write-Warn "No known installer for $missing - provision it with -Source <msix>"
    }

    $work += $needs.Count
    if ($work -eq 0 -and -not $Source -and -not ($Provision -and $provisionTargets.Count -gt 0)) {
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
        Start-Transcript -Path (Join-Path $LogPath "Repair-AppxPackageStore_$stamp.log") | Out-Null
        $transcribing = $true
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
            Write-Warn "Could not re-register $($item.FullName): $($_.Exception.Message)"
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
        if ($provisionTargets.Count -eq 0) { Write-Skip 'Nothing to provision with a known installer' }
        foreach ($target in $provisionTargets) {
            if (-not $PSCmdlet.ShouldProcess($target, "Provision for all users ($($KnownInstallers[$target].File) $($KnownInstallers[$target].Args))")) { continue }
            try {
                if (-not (Invoke-KnownInstaller -PackageName $target)) {
                    Write-AppxDeploymentError -Minutes 20
                    $exitCode = 1
                }
            } catch {
                Write-Bad "Provisioning $target failed: $($_.Exception.Message)"
                $exitCode = 1
            }
        }
    } elseif ($provisionTargets.Count -gt 0) {
        Write-Bad ("Still to do: provision {0} - run again with -Provision" -f ($provisionTargets -join ', '))
        $exitCode = 1
    }
    if (-not $Source) {
        if (-not $Provision) { Write-Skip 'No -Source or -Provision given' }
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
        foreach ($target in $needs) {
            $asked = ($requestsAfter | Where-Object { $_.Name -eq $target } | Sort-Object Version -Descending | Select-Object -First 1 | ForEach-Object { $_.Version })
            if ($provisionedAfter.ContainsKey($target) -and $asked -and $provisionedAfter[$target] -lt $asked) {
                Write-Warn "$target is provisioned at $($provisionedAfter[$target]), still older than the $asked profiles ask for - bring the other hosts in the pool to the same build"
                $exitCode = 1
            } elseif ($provisionedAfter.ContainsKey($target)) {
                Write-Ok "$target is provisioned: $($provisionedAfter[$target]) - users get it at their next sign-in, and FSLogix saves that version at their next sign-out"
            } else {
                Write-Bad "$target is still not provisioned on this host"
                $exitCode = 1
            }
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
