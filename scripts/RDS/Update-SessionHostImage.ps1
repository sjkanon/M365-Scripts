#Requires -Version 5.1
<#
.SYNOPSIS
    Check and prepare a Windows 11 multi-session image (or a running AVD session host)
    so new Teams, new Outlook and Copilot keep working with FSLogix profile containers,
    without touching FSLogix itself.

.DESCRIPTION
    Teams, new Outlook and Copilot are MSIX apps. On a pooled AVD host with FSLogix
    they break in one recurring way: a user's app updates itself on host A, FSLogix
    saves that exact version in the profile at sign-out, and at the next sign-in on
    host B - which does not have that version - registering it fails with 0x80070490.
    FSLogix 2210 HF4 (Teams) and 25.06 (Outlook) register by package family instead,
    but this script deliberately leaves the FSLogix build alone. What it does instead
    is keep every host at the newest build of the apps and of everything they need,
    the same on every host:

      1. Windows      Edition (multi-session), build, pending reboot.
      2. FSLogix      Build and InstallAppxPackages, read only.
      3. Updates      New Outlook updates itself weekly from the Office CDN, not
                      through the Store, and has no switch to stop it - so with
                      FSLogix below 25.06 the cure is running this script weekly on
                      every host. Teams: below 2210 HF4 its self-update is turned off
                      (disableAutoUpdate = 1) and every run updates it centrally;
                      on a newer FSLogix it may update itself. The Store setting is
                      shown, not changed. Edge Update policies that block WebView2 or
                      Edge are reported.
      4. WebView2     The Evergreen runtime all three apps render with, against the
                      current Edge Stable build. Behind or missing: the Evergreen
                      Standalone installer (signature-checked) is run.
      5. Apps         Teams, new Outlook, the Microsoft 365 Copilot app and the
                      unified Copilot app: provisioned build, and users holding a
                      newer build than the image provisions (the drift FSLogix
                      replays).
      6. Frameworks   Every PackageDependency in those apps' manifests (VCLibs,
                      UI.Xaml, WindowsAppRuntime, ...) must be on the machine at the
                      MinVersion the manifest asks for. An image from 2024 is where
                      these are typically too old.
      7. Teams VDI    IsWVDEnvironment, a Teams build new enough for SlimCore
                      (24193.1805.3040.8975), the Teams Meeting add-in, and the WebRTC
                      redirector - out of support since 1 October 2026, stops working
                      1 April 2027, kept only as a fallback for endpoints that cannot
                      do SlimCore yet.
      8. Office       Shared Computer Activation (required on multi-session) and the
                      update channel.
      9. Sign-in      Microsoft.AAD.BrokerPlugin present, and no FSLogix
                      redirections.xml exclusion of AppData\Local\Packages or of the
                      app folders - without them users sign in over and over.
     10. Capture      With -ForCapture: packages a user installed that are not
                      provisioned (Sysprep fails on them) and a pending reboot.

    Without -CheckOnly every run updates, whether a check found something or not:
    Teams' self-update setting where FSLogix needs it, Shared Computer Activation,
    WebView2, then the apps to their newest build through the existing repo scripts -
    Repair-AppxPackageStore.ps1 (-Name teams,outlook -Latest -Provision -RemoveOld,
    then -Name copilot -Provision), whose Microsoft installers bring the frameworks
    their package depends on, and Update-TeamsClient.ps1 -AvdOptimizations for the
    newest Teams, IsWVDEnvironment, the WebRTC redirector and the meeting add-in.
    They only change what is behind. Everything is read back afterwards; a framework
    still too old then stays a finding.

    Newest Outlook: Microsoft publishes no version feed for it, so -Latest takes the
    newest build users on the host already updated to (and Microsoft's installer
    otherwise). On a pool that means every host provisions what the most recent user
    got - which is exactly the build FSLogix will ask the others for.

    The two scripts are taken from the repo when this one runs from it (or from the
    folder -ComputerName copies them to). Run on its own - only this file on the
    image VM - it fetches them from GitHub at a pinned commit of main and runs them
    only when their SHA-256 matches; see $HelperCommit / $HelperHashes.

    It never restarts the machine, and neither do the scripts and installers it calls
    (every msiexec runs with /norestart). A pending reboot is reported and left to you.

    Run it on the image VM before capture, or on every session host - -ComputerName
    does the latter and ends with one table across the pool, because "is every host
    on the same build?" is the question that matters on a pooled host.

.PARAMETER CheckOnly
    Report only, change nothing. Exit code 2 means there is work to do.

.PARAMETER ForCapture
    The machine is the image VM and is about to be sysprepped and captured: also
    fail on a pending reboot and on per-user packages that are not provisioned.

.PARAMETER SkipApps
    Do not call Repair-AppxPackageStore.ps1 and Update-TeamsClient.ps1; only the
    policies, Shared Computer Activation and WebView2 are fixed.

.PARAMETER ComputerName
    Session hosts to run on over PowerShell remoting. This script and the two it
    calls are copied to C:\IT\SessionHostImage on each host.

.PARAMETER Credential
    Credential for -ComputerName.

.PARAMETER WorkingDir
    Folder for downloads - WebView2 and, when they are not next to this script,
    the two helper scripts (default: C:\IT\SessionHostImage).

.PARAMETER LogPath
    Folder for the transcript of a run that changes something (default: C:\Temp).

.EXAMPLE
    # What does this host or image need? Changes nothing.
    .\Update-SessionHostImage.ps1 -CheckOnly

.EXAMPLE
    # All three session hosts side by side, read only
    .\Update-SessionHostImage.ps1 -ComputerName avd-0,avd-1,avd-2 -CheckOnly

.EXAMPLE
    # Prepare the image VM, then check it is fit for capture
    .\Update-SessionHostImage.ps1 -Confirm:$false
    .\Update-SessionHostImage.ps1 -CheckOnly -ForCapture

.EXAMPLE
    # Bring all three hosts level (drain them first)
    .\Update-SessionHostImage.ps1 -ComputerName avd-0,avd-1,avd-2 -Confirm:$false

.NOTES
    Author  : Sjoerd Kanon
    Platform: Windows 11 Enterprise multi-session, run as administrator or as System
    Exit    : 0 clean, 1 failed, 2 check-only found work
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param (
    [switch]   $CheckOnly,
    [switch]   $ForCapture,
    [switch]   $SkipApps,
    [string[]] $ComputerName,
    [pscredential] $Credential,
    [string]   $WorkingDir = 'C:\IT\SessionHostImage',
    [string]   $LogPath    = 'C:\Temp',
    # Set by the -ComputerName orchestrator: emit the host's state as an object.
    [Parameter(DontShow)]
    [switch]   $PassThru
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# -- Constants -------------------------------------------------------------------
$Apps = [ordered]@{
    'MSTeams'                      = 'Teams'
    'Microsoft.OutlookForWindows'  = 'Outlook'
    'Microsoft.MicrosoftOfficeHub' = 'M365 Copilot'
    'Microsoft.Copilot'            = 'Copilot'
}
# FSLogix builds that register by package family (release notes: 2210 HF4, 25.06).
$FslogixFamilyBuild = @{
    'MSTeams'                     = [version] '2.9.8884.27471'
    'Microsoft.OutlookForWindows' = [version] '3.25.626.21064'
}
$SlimCoreMinimumTeams = [version] '24193.1805.3040.8975'
$WebView2Guid   = '{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}'
$EdgeGuid       = '{56EB18F8-B008-4CBD-B6D2-8C97FE7E9062}'
$CopilotGuid    = '{C50565E9-CCCF-44B4-BA15-5AC5C6569197}'
$WebView2Url    = 'https://go.microsoft.com/fwlink/?linkid=2124701'   # Evergreen Standalone x64
$EdgeReleaseApi = 'https://edgeupdates.microsoft.com/api/products'
# The two scripts this one calls, when they are not next to it: fetched from this
# repo at a pinned commit of main and refused unless the SHA-256 matches. A newer
# version is never run unseen - to move up, put the new commit and the hashes of
# both files at that commit here, after reading the diff.
$HelperCommit = '746541e4b3bb47fd54867dc3b79c40a629c7fea6'   # main, 2026-10-05
$HelperHashes = @{
    'Update-TeamsClient.ps1'      = 'CC6D97EA9ADBEF12E3614FA482A219CE04E5AB421ECFAF5D11F86AEEC0397E7D'
    'Repair-AppxPackageStore.ps1' = '88A9CEAB9F801CF60BFD25AEB69320A97AECD9F19323BC6E354A60118A04C45C'
}
$WebRtcEndOfSupport      = [datetime] '2026-10-01'
$WebRtcEndOfAvailability = [datetime] '2027-04-01'
$OfficeChannels = @{
    '492350f6-3a01-4f97-b9c0-c7c6ddf67d60' = 'Current'
    '55336b82-a18d-4dd6-b5f6-9e5095c314a6' = 'Monthly Enterprise'
    '7ffbc6bf-bc32-4f92-8982-f9dd17fd3114' = 'Semi-Annual Enterprise'
    'b8f9b850-328d-4355-9145-c59439a0c4cf' = 'Semi-Annual Enterprise (Preview)'
    '64256afe-f5d9-4f86-8936-8840a6a4f5be' = 'Current (Preview)'
    '5440fd1f-7ecb-4221-8110-145efaa6372f' = 'Beta'
}
$StorePolicyPath  = 'HKLM:\SOFTWARE\Policies\Microsoft\WindowsStore'
$TeamsPath        = 'HKLM:\SOFTWARE\Microsoft\Teams'
$EdgeUpdatePolicy = 'HKLM:\SOFTWARE\Policies\Microsoft\EdgeUpdate'
$C2RConfigPath    = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
$FslogixProfiles  = 'HKLM:\SOFTWARE\FSLogix\Profiles'

# -- Output ------------------------------------------------------------------------
function Write-Step { param([string] $Message) Write-Host ''; Write-Host "  $Message" -ForegroundColor Cyan }
function Write-Ok   { param([string] $Message) Write-Host "  [ OK ] $Message" -ForegroundColor Green }
function Write-Skip { param([string] $Message) Write-Host "  [SKIP] $Message" -ForegroundColor DarkGray }
function Write-Warn { param([string] $Message) Write-Host "  [WARN] $Message" -ForegroundColor Yellow }
function Write-Bad  { param([string] $Message) Write-Host "  [FAIL] $Message" -ForegroundColor Red }

$Findings = [System.Collections.Generic.List[object]]::new()
function Add-Finding {
    <# One line of work: what, how bad, and whether this run can fix it. #>
    param([string] $Area, [string] $Message, [ValidateSet('Warn', 'Fail')] [string] $Level = 'Warn', [string] $Fix)
    if ($Level -eq 'Fail') { Write-Bad $Message } else { Write-Warn $Message }
    $Findings.Add([PSCustomObject]@{ Area = $Area; Level = $Level; Message = $Message; Fix = $Fix })
}

function Get-PropertyValue {
    <# Property access that returns $null instead of tripping Set-StrictMode. #>
    param($Object, [string] $Name)
    if ($null -eq $Object) { return $null }
    $prop = $Object.PSObject.Properties[$Name]
    if ($prop) { return $prop.Value }
    return $null
}

function ConvertTo-Version {
    param($Text)
    $v = $null
    if ($Text -and [version]::TryParse(([string] $Text).Trim(), [ref] $v)) { return $v }
    return $null
}

function Get-ForwardedArgument {
    <# Rebuild the caller's own parameters as a command line for a relaunch. #>
    param([Parameter(Mandatory)] $Bound)
    $list = @()
    foreach ($entry in $Bound.GetEnumerator()) {
        if ($entry.Value -is [switch] -or $entry.Value -is [bool]) {
            if ($entry.Value) { $list += "-$($entry.Key)" } else { $list += "-$($entry.Key):`$false" }
        } elseif ($entry.Value -is [array]) {
            $list += "-$($entry.Key)"; $list += ($entry.Value -join ',')
        } else {
            $list += "-$($entry.Key)"; $list += [string] $entry.Value
        }
    }
    return $list
}

# -- Several session hosts -----------------------------------------------------------
$ComputerName = @($ComputerName | ForEach-Object { $_ -split '[,;\s]' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
if ($ComputerName.Count -gt 0) {
    $forward = @{ PassThru = $true }
    foreach ($entry in $PSBoundParameters.GetEnumerator()) {
        if ($entry.Key -in @('ComputerName', 'Credential', 'Confirm', 'PassThru')) { continue }
        $forward[$entry.Key] = if ($entry.Value -is [switch]) { [bool] $entry.Value } else { $entry.Value }
    }
    $changes    = -not ($CheckOnly -or $WhatIfPreference)
    $confirmOff = $PSBoundParameters.ContainsKey('Confirm') -and -not $PSBoundParameters['Confirm']
    if ($changes -and -not $confirmOff) {
        $answer = Read-Host ("  Prepare {0} (apps are re-provisioned - drain the hosts first)? [y/N]" -f ($ComputerName -join ', '))
        if ($answer -notmatch '^[Yy]') { Write-Host '  Cancelled - nothing was changed.' -ForegroundColor DarkGray; exit 0 }
    }
    if ($changes) { $forward['Confirm'] = $false }
    if ($WhatIfPreference) { $forward['WhatIf'] = $true }

    # The two scripts this one calls travel along, flat in one folder.
    $remoteDir = 'C:\IT\SessionHostImage'
    $toCopy = @($PSCommandPath)
    foreach ($helper in 'Repair-AppxPackageStore.ps1', 'Update-TeamsClient.ps1') {
        $local = Join-Path (Join-Path (Split-Path $PSScriptRoot) 'Device') $helper
        if (Test-Path $local) { $toCopy += $local }
    }

    $summary = [System.Collections.Generic.List[object]]::new()
    foreach ($computer in $ComputerName) {
        Write-Host ''
        Write-Host ("  ==== {0} " -f $computer).PadRight(80, '=') -ForegroundColor Cyan
        $row = [ordered]@{ Host = $computer; Exit = 1; Build = ''; FSLogix = ''; WebView2 = ''; Teams = ''; Outlook = ''; Copilot = ''; Findings = '' }
        $session = $null
        try {
            $sessionArgs = @{ ComputerName = $computer; ErrorAction = 'Stop' }
            if ($Credential) { $sessionArgs['Credential'] = $Credential }
            $session = New-PSSession @sessionArgs
            Invoke-Command -Session $session -ScriptBlock { param($Dir) New-Item -ItemType Directory -Path $Dir -Force | Out-Null } -ArgumentList $remoteDir
            foreach ($file in $toCopy) { Copy-Item -Path $file -Destination $remoteDir -ToSession $session -Force }

            $result = Invoke-Command -Session $session -ScriptBlock {
                param($Path, $Params)
                & $Path @Params
            } -ArgumentList (Join-Path $remoteDir (Split-Path $PSCommandPath -Leaf)), $forward

            $state = @($result | Where-Object { $_ -and $_.PSObject.Properties.Name -contains 'ImageResult' }) | Select-Object -Last 1
            if ($state) {
                foreach ($key in @($row.Keys | Where-Object { $_ -ne 'Host' })) { $row[$key] = $state.$key }
            } else {
                $row.Findings = 'no result came back'
            }
        } catch {
            $gist = ($_.Exception.Message -replace '\s+', ' ')
            $row.Findings = 'not reached: ' + $(if ($gist.Length -gt 60) { $gist.Substring(0, 60) + '...' } else { $gist })
            Write-Host "  [FAIL] $computer - $($_.Exception.Message)" -ForegroundColor Red
        } finally {
            if ($session) { Remove-PSSession $session -ErrorAction SilentlyContinue }
        }
        $summary.Add([PSCustomObject] $row)
    }

    Write-Host ''
    Write-Host '  ==== Pool '.PadRight(80, '=') -ForegroundColor Cyan
    $summary | Format-Table -AutoSize | Out-String -Width 220 | Write-Host
    $reached = @($summary | Where-Object { $_.Build })
    foreach ($column in 'Build', 'FSLogix', 'WebView2', 'Teams', 'Outlook', 'Copilot') {
        $values = @($reached | ForEach-Object { $_.$column } | Sort-Object -Unique)
        if ($values.Count -gt 1) {
            Write-Host ("  [WARN] {0} differs between hosts: {1}" -f $column, ($values -join ' / ')) -ForegroundColor Yellow
        }
    }
    $worst = ($summary | Measure-Object -Property Exit -Maximum).Maximum
    if (@($summary | Where-Object { $_.Exit -eq 1 }).Count -gt 0) { $worst = 1 }
    exit [int] $worst
}

# -- Windows PowerShell, 64-bit, elevated --------------------------------------------
# The AppX cmdlets need Windows PowerShell; under WOW64 HKLM lands in WOW6432Node.
$nativeShell = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
    $nativeShell = Join-Path $env:WINDIR 'SysNative\WindowsPowerShell\v1.0\powershell.exe'
}
if ($PSVersionTable.PSEdition -eq 'Core' -or ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess)) {
    $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath) + (Get-ForwardedArgument -Bound $PSBoundParameters)
    & $nativeShell @argList
    exit $LASTEXITCODE
}
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
if (-not ([Security.Principal.WindowsPrincipal] $identity).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error 'Administrator rights are required. Run the script as System or from an elevated session.'
    exit 1
}

$changing = -not ($CheckOnly -or $WhatIfPreference)
if ($changing) {
    New-Item -ItemType Directory -Path $LogPath, $WorkingDir -Force | Out-Null
    Start-Transcript -Path (Join-Path $LogPath ('Update-SessionHostImage_{0}.log' -f (Get-Date -Format 'yyyyMMdd_HHmmss'))) | Out-Null
}

# -- Readers -------------------------------------------------------------------------
function Get-EdgeUpdateClientVersion {
    param([string] $Guid)
    foreach ($root in 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients', 'HKLM:\SOFTWARE\Microsoft\EdgeUpdate\Clients') {
        $v = ConvertTo-Version (Get-PropertyValue (Get-ItemProperty (Join-Path $root $Guid) -ErrorAction SilentlyContinue) 'pv')
        if ($v -and $v -ne [version] '0.0.0.0') { return $v }
    }
    return $null
}

function Get-EdgeStableVersion {
    <# The current Edge Stable build for Windows x64 - WebView2 Evergreen follows it. #>
    try {
        $products = Invoke-RestMethod -Uri $EdgeReleaseApi -UseBasicParsing -TimeoutSec 20
        $stable   = @($products | Where-Object { $_.Product -eq 'Stable' }) | Select-Object -First 1
        $release  = @($stable.Releases | Where-Object { $_.Platform -eq 'Windows' -and $_.Architecture -eq 'x64' }) |
                    Sort-Object { [version] $_.ProductVersion } -Descending | Select-Object -First 1
        return ConvertTo-Version $release.ProductVersion
    } catch {
        return $null
    }
}

function Get-FslogixVersion {
    $frx = Join-Path $env:ProgramFiles 'FSLogix\Apps\frxsvc.exe'
    if (-not (Test-Path $frx)) { return $null }
    $i = (Get-Item $frx).VersionInfo
    return [version] ('{0}.{1}.{2}.{3}' -f $i.FileMajorPart, $i.FileMinorPart, $i.FileBuildPart, $i.FilePrivatePart)
}

function Get-ProvisionedVersion {
    <# Newest provisioned build per DisplayName, for the apps in $Apps. #>
    $map = @{}
    foreach ($p in @(Get-AppxProvisionedPackage -Online | Where-Object { $Apps.Contains($_.DisplayName) })) {
        $v = ConvertTo-Version $p.Version
        if (-not $map.ContainsKey($p.DisplayName) -or $map[$p.DisplayName] -lt $v) { $map[$p.DisplayName] = $v }
    }
    return $map
}

function Get-WebRtcRedirector {
    foreach ($root in 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall') {
        if (-not (Test-Path $root)) { continue }
        foreach ($key in Get-ChildItem $root) {
            if ($key.GetValue('DisplayName') -like '*Remote Desktop WebRTC Redirector Service*') { return $key.GetValue('DisplayVersion') }
        }
    }
    return $null
}

function Get-PendingReboot {
    $reasons = @()
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') { $reasons += 'servicing (CBS)' }
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') { $reasons += 'Windows Update' }
    return $reasons
}

function Set-RegistryDword {
    param([string] $Path, [string] $Name, [int] $Value, [string] $Why)
    if ($PSCmdlet.ShouldProcess("$Path\$Name", "set to $Value ($Why)")) {
        if (-not (Test-Path $Path)) { New-Item -Path $Path -Force | Out-Null }
        New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType DWord -Force | Out-Null
        Write-Ok "$Path\$Name = $Value"
    }
}

# =====================================================================================
function Invoke-Check {
    <# Every check, read only. Returns the host state for the pool table. #>
    $Findings.Clear()
    $os = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
    $build = '{0}.{1}' -f $os.CurrentBuildNumber, $os.UBR

    Write-Step '1. Windows'
    if ($os.EditionID -eq 'ServerRdsh') { Write-Ok "$($os.ProductName) (multi-session), build $build" }
    else { Add-Finding 'Windows' "Edition is $($os.EditionID), not Enterprise multi-session (ServerRdsh) - build $build" }
    $pending = @(Get-PendingReboot)
    if ($pending.Count -gt 0) {
        $level = if ($ForCapture) { 'Fail' } else { 'Warn' }
        Add-Finding 'Windows' ("A reboot is pending ({0}) - AppX changes and capture both need it done first" -f ($pending -join ', ')) $level 'reboot'
    } else { Write-Ok 'No reboot pending' }

    Write-Step '2. FSLogix (read only - not changed by this script)'
    $fslogix = Get-FslogixVersion
    if (-not $fslogix) { Write-Skip 'FSLogix is not installed' }
    else {
        Write-Ok "FSLogix $fslogix"
        $replay = Get-PropertyValue (Get-ItemProperty $FslogixProfiles -ErrorAction SilentlyContinue) 'InstallAppxPackages'
        if ($null -eq $replay -or $replay -eq 1) { Write-Skip '  InstallAppxPackages on (default): packages saved in the profile are re-registered at sign-in' }
        if ($fslogix -lt $FslogixFamilyBuild['MSTeams']) {
            Write-Warn "  Below $($FslogixFamilyBuild['MSTeams']): Teams is re-registered at its exact saved version - its self-update must stay off and every run updates it centrally (step 3)"
        }
        if ($fslogix -lt $FslogixFamilyBuild['Microsoft.OutlookForWindows']) {
            Write-Warn "  Below $($FslogixFamilyBuild['Microsoft.OutlookForWindows']): Outlook is re-registered at its exact saved version, and it updates itself weekly - run this script weekly on every host so they provision the newest build users have"
        }
    }

    Write-Step '3. Updates'
    # New Outlook updates itself weekly from the Office CDN, not through the Store,
    # and Microsoft documents no switch to stop it. The Store setting is shown only.
    $store = Get-PropertyValue (Get-ItemProperty $StorePolicyPath -ErrorAction SilentlyContinue) 'AutoDownload'
    Write-Skip ("Microsoft Store automatic app updates: {0} (Outlook does not update through the Store)" -f $(if ($store -eq 2) { 'off (AutoDownload = 2)' } elseif ($null -eq $store) { 'on (not set)' } else { "AutoDownload = $store" }))
    $teamsKey  = Get-ItemProperty $TeamsPath -ErrorAction SilentlyContinue
    $teamsSelf = (Get-PropertyValue $teamsKey 'disableAutoUpdate') -ne 1
    if ($fslogix -and $fslogix -lt $FslogixFamilyBuild['MSTeams']) {
        if ($teamsSelf) { Add-Finding 'Updates' "Teams updates itself per user, and this FSLogix build replays exact versions - hosts drift apart" 'Warn' 'teamsupdate' }
        else { Write-Ok 'Teams self-update is off (disableAutoUpdate = 1); every run of this script updates it to the newest build' }
    } else {
        Write-Ok ("Teams self-update is {0}; FSLogix registers Teams by family, so that is safe" -f $(if ($teamsSelf) { 'on' } else { 'off (disableAutoUpdate = 1)' }))
    }
    $edgePolicy = Get-ItemProperty $EdgeUpdatePolicy -ErrorAction SilentlyContinue
    foreach ($item in @(@{ Name = 'WebView2'; Guid = $WebView2Guid }, @{ Name = 'Edge'; Guid = $EdgeGuid })) {
        $value = Get-PropertyValue $edgePolicy "Update$($item.Guid)"
        if ($null -eq $value) { $value = Get-PropertyValue $edgePolicy 'UpdateDefault' }
        if ($value -eq 0) { Add-Finding 'Updates' "Edge Update policy blocks $($item.Name) updates ($EdgeUpdatePolicy) - reported, not changed" }
    }

    Write-Step '4. WebView2 runtime'
    $webView2 = Get-EdgeUpdateClientVersion $WebView2Guid
    $edge     = Get-EdgeUpdateClientVersion $EdgeGuid
    $stable   = Get-EdgeStableVersion
    if (-not $webView2) { Add-Finding 'WebView2' 'WebView2 Runtime is not installed - Teams, Outlook and Copilot cannot render' 'Fail' 'webview2' }
    elseif ($stable -and $webView2 -lt $stable) { Add-Finding 'WebView2' "WebView2 $webView2 is behind Edge Stable $stable" 'Warn' 'webview2' }
    else { Write-Ok ("WebView2 {0}{1}" -f $webView2, $(if ($stable) { " (Edge Stable: $stable)" } else { ' (Edge Stable could not be looked up)' })) }
    if ($edge) {
        if ($stable -and $edge.Major -lt $stable.Major) { Add-Finding 'WebView2' "Edge $edge is behind Edge Stable $stable - Edge Update keeps it current unless a policy blocks it" }
        else { Write-Ok "Edge $edge" }
    } else { Add-Finding 'WebView2' 'Microsoft Edge is not installed - the unified Copilot app is installed through Edge Update' }

    Write-Step '5. Apps'
    $provisioned = Get-ProvisionedVersion
    $registered  = @{}
    foreach ($name in $Apps.Keys) {
        $all = @(Get-AppxPackage -AllUsers -Name $name -ErrorAction SilentlyContinue)
        $registered[$name] = $all
        $have = if ($provisioned.ContainsKey($name)) { $provisioned[$name] } else { $null }
        if ($have) {
            $newer = @($all | Where-Object { (ConvertTo-Version $_.Version) -gt $have } | ForEach-Object { $_.Version } | Sort-Object -Unique)
            if ($newer.Count -gt 0) {
                Add-Finding 'Apps' ("{0} provisioned at {1}, but users hold {2} - FSLogix replays those on hosts that do not have them" -f $Apps[$name], $have, ($newer -join ', ')) 'Warn' 'apps'
            } else { Write-Ok "$($Apps[$name]) provisioned at $have" }
        } elseif ($name -in 'MSTeams', 'Microsoft.OutlookForWindows') {
            Add-Finding 'Apps' "$($Apps[$name]) is not provisioned for all users" 'Fail' 'apps'
        } else {
            Write-Skip "$($Apps[$name]) ($name) is not provisioned"
        }
    }
    $unified = Get-EdgeUpdateClientVersion $CopilotGuid
    if ($unified) { Write-Ok "Copilot app (unified, Edge Update) $unified" }
    elseif (-not ($provisioned.ContainsKey('Microsoft.MicrosoftOfficeHub') -or $provisioned.ContainsKey('Microsoft.Copilot'))) {
        Add-Finding 'Apps' 'No Copilot app on this machine - neither the unified app nor a provisioned package' 'Warn' 'copilot'
    }

    Write-Step '6. Frameworks the apps depend on'
    $frameworks = @(Get-AppxPackage -AllUsers -PackageTypeFilter Framework -ErrorAction SilentlyContinue)
    $checked = @{}
    foreach ($name in $Apps.Keys) {
        $pkg = @($registered[$name] | Sort-Object { ConvertTo-Version $_.Version } -Descending) | Select-Object -First 1
        if (-not $pkg -or -not $pkg.InstallLocation) { continue }
        $manifestPath = Join-Path $pkg.InstallLocation 'AppxManifest.xml'
        if (-not (Test-Path $manifestPath)) {
            Add-Finding 'Frameworks' "$($Apps[$name]) $($pkg.Version): the manifest is gone from $($pkg.InstallLocation) - the registration has nothing behind it" 'Fail' 'apps'
            continue
        }
        [xml] $manifest = Get-Content -LiteralPath $manifestPath -Raw
        foreach ($dep in @($manifest.GetElementsByTagName('PackageDependency'))) {
            $min = ConvertTo-Version $dep.MinVersion
            $key = "$($dep.Name)|$min"
            if ($checked.ContainsKey($key)) { continue }
            $checked[$key] = $true
            $best = @($frameworks | Where-Object { $_.Name -eq $dep.Name -and $_.Architecture -in 'X64', 'Neutral' } |
                      Sort-Object { ConvertTo-Version $_.Version } -Descending) | Select-Object -First 1
            if (-not $best) { Add-Finding 'Frameworks' "$($dep.Name) is missing ($($Apps[$name]) needs $min or newer)" 'Fail' 'apps' }
            elseif ((ConvertTo-Version $best.Version) -lt $min) { Add-Finding 'Frameworks' "$($dep.Name) $($best.Version) is older than the $min $($Apps[$name]) needs" 'Fail' 'apps' }
            else { Write-Ok "$($dep.Name) $($best.Version) (needs $min)" }
        }
    }
    if ($checked.Count -eq 0) { Write-Skip 'No installed app to read dependencies from' }

    Write-Step '7. Teams on AVD'
    if ((Get-PropertyValue $teamsKey 'IsWVDEnvironment') -eq 1) { Write-Ok 'IsWVDEnvironment = 1' }
    else { Add-Finding 'Teams VDI' 'IsWVDEnvironment is not 1 - Teams does not optimize media at all, neither SlimCore nor WebRTC' 'Fail' 'teams' }
    $teams = if ($provisioned.ContainsKey('MSTeams')) { $provisioned['MSTeams'] } else { $null }
    if ($teams -and $teams -lt $SlimCoreMinimumTeams) { Add-Finding 'Teams VDI' "Teams $teams is older than $SlimCoreMinimumTeams, the first build that can use SlimCore" 'Fail' 'apps' }
    elseif ($teams) { Write-Ok "Teams $teams can use SlimCore (endpoints need Windows App 2.0.352.0 or newer)" }
    $webRtc = Get-WebRtcRedirector
    $today  = Get-Date
    if ($today -ge $WebRtcEndOfAvailability) {
        if ($webRtc) { Write-Warn "WebRTC redirector $webRtc is still installed; it stopped working on $($WebRtcEndOfAvailability.ToString('d MMMM yyyy')) - remove it with Update-TeamsClient.ps1 -RemoveWebRtcRedirector" }
        else { Write-Ok 'WebRTC redirector not installed (no longer available - SlimCore only)' }
    } elseif ($webRtc) {
        $note = if ($today -ge $WebRtcEndOfSupport) { "out of support since $($WebRtcEndOfSupport.ToString('d MMMM yyyy')), stops working $($WebRtcEndOfAvailability.ToString('d MMMM yyyy'))" } else { "end of support $($WebRtcEndOfSupport.ToString('d MMMM yyyy'))" }
        Write-Ok "WebRTC redirector $webRtc - fallback for endpoints without SlimCore, $note"
    } else {
        Add-Finding 'Teams VDI' "WebRTC redirector is not installed - endpoints that cannot do SlimCore yet render media on this host (fallback until $($WebRtcEndOfAvailability.ToString('d MMMM yyyy')))" 'Warn' 'teams'
    }
    $addIn = $null
    foreach ($root in 'HKLM:\SOFTWARE\Microsoft\Office\Outlook\Addins\TeamsAddin.FastConnect', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\Outlook\Addins\TeamsAddin.FastConnect') {
        $value = Get-PropertyValue (Get-ItemProperty $root -ErrorAction SilentlyContinue) 'LoadBehavior'
        if ($null -ne $value) { $addIn = $value }
    }
    if ($addIn -eq 3) { Write-Ok 'Teams Meeting add-in registered for classic Outlook (LoadBehavior 3)' }
    else { Add-Finding 'Teams VDI' 'Teams Meeting add-in is not registered machine-wide for classic Outlook' 'Warn' 'teams' }
    Write-Skip 'Tenant side, not checkable here: Teams VDI policy "VDI2Optimization" must be Enabled for SlimCore'

    Write-Step '8. Microsoft 365 Apps'
    $c2r = Get-ItemProperty $C2RConfigPath -ErrorAction SilentlyContinue
    if (-not $c2r) { Write-Skip 'Microsoft 365 Apps (Click-to-Run) is not installed' }
    else {
        $channelUrl = [string] $(if (Get-PropertyValue $c2r 'UpdateChannel') { Get-PropertyValue $c2r 'UpdateChannel' } else { Get-PropertyValue $c2r 'CDNBaseUrl' })
        $channel    = 'unknown'
        foreach ($id in $OfficeChannels.Keys) { if ($channelUrl -match $id) { $channel = $OfficeChannels[$id] } }
        Write-Ok ("Microsoft 365 Apps {0}, {1} channel" -f (Get-PropertyValue $c2r 'VersionToReport'), $channel)
        if ((Get-PropertyValue $c2r 'SharedComputerLicensing') -ne '1') {
            Add-Finding 'Office' 'Shared Computer Activation is off - required on multi-session, without it Office (and Copilot in Office) activates per device' 'Fail' 'sca'
        } else { Write-Ok 'Shared Computer Activation is on' }
    }

    Write-Step '9. Sign-in (WAM) and FSLogix redirections'
    $broker = @(Get-AppxPackage -AllUsers -Name 'Microsoft.AAD.BrokerPlugin' -ErrorAction SilentlyContinue) | Select-Object -First 1
    if ($broker) { Write-Ok "Microsoft.AAD.BrokerPlugin $($broker.Version)" }
    else { Add-Finding 'Sign-in' 'Microsoft.AAD.BrokerPlugin is not on this machine - single sign-on for Teams, Outlook and Copilot fails' 'Fail' }
    $redirFolder = Get-PropertyValue (Get-ItemProperty $FslogixProfiles -ErrorAction SilentlyContinue) 'RedirXMLSourceFolder'
    if ($redirFolder) {
        $redirFile = Join-Path $redirFolder 'redirections.xml'
        if (Test-Path -LiteralPath $redirFile) {
            [xml] $redir = Get-Content -LiteralPath $redirFile -Raw
            $bad = @($redir.GetElementsByTagName('Exclude') | ForEach-Object { $_.InnerText } |
                     Where-Object { $_ -match '(?i)AppData\\Local\\Packages\\?$|MSTeams|OutlookForWindows|AAD\.BrokerPlugin|MicrosoftOfficeHub' })
            if ($bad.Count -gt 0) { Add-Finding 'Sign-in' ("redirections.xml excludes {0} - these must roam in the profile container" -f ($bad -join ', ')) 'Fail' }
            else { Write-Ok "redirections.xml excludes none of the app or broker folders" }
        } else { Write-Warn "RedirXMLSourceFolder is set to $redirFolder, but redirections.xml is not reachable there" }
    } else { Write-Skip 'No FSLogix redirections.xml configured' }

    if ($ForCapture) {
        Write-Step '10. Ready for capture'
        $provisionedFamilies = @(Get-AppxProvisionedPackage -Online | ForEach-Object { $_.PackageName -replace '_[^_]+_[^_]+_[^_]*_', '_' })
        $blockers = @(Get-AppxPackage -AllUsers -PackageTypeFilter Main -ErrorAction SilentlyContinue |
                      Where-Object { -not $_.NonRemovable -and $_.SignatureKind -ne 'System' -and
                                     $_.PackageFamilyName -notin $provisionedFamilies -and @($_.PackageUserInformation).Count -gt 0 })
        if ($blockers.Count -gt 0) {
            Add-Finding 'Capture' ("Installed for a user but not provisioned - Sysprep stops on these: {0}" -f (($blockers | ForEach-Object { $_.Name } | Sort-Object -Unique) -join ', ')) 'Fail'
        } else { Write-Ok 'No per-user packages that would stop Sysprep' }
    }

    return [PSCustomObject]@{
        Build    = $build
        FSLogix  = if ($fslogix) { [string] $fslogix } else { '-' }
        WebView2 = if ($webView2) { [string] $webView2 } else { '-' }
        Teams    = if ($provisioned.ContainsKey('MSTeams')) { [string] $provisioned['MSTeams'] } else { '-' }
        Outlook  = if ($provisioned.ContainsKey('Microsoft.OutlookForWindows')) { [string] $provisioned['Microsoft.OutlookForWindows'] } else { '-' }
        Copilot  = if ($unified) { "$unified (unified)" } elseif ($provisioned.ContainsKey('Microsoft.MicrosoftOfficeHub')) { [string] $provisioned['Microsoft.MicrosoftOfficeHub'] } else { '-' }
    }
}

function Install-WebView2 {
    $installer = Join-Path $WorkingDir 'MicrosoftEdgeWebView2RuntimeInstallerX64.exe'
    if (-not $PSCmdlet.ShouldProcess('WebView2 Runtime', 'download and run the Evergreen Standalone installer')) { return }
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -Uri $WebView2Url -OutFile $installer -UseBasicParsing
    $signature = Get-AuthenticodeSignature -FilePath $installer
    if ($signature.Status -ne 'Valid' -or $signature.SignerCertificate.Subject -notmatch 'O=Microsoft Corporation') {
        Remove-Item $installer -Force -ErrorAction SilentlyContinue
        throw "The WebView2 installer is not signed by Microsoft ($($signature.Status)) - not run"
    }
    $process = Start-Process -FilePath $installer -ArgumentList '/silent', '/install' -Wait -PassThru
    if ($process.ExitCode -ne 0) { Write-Warn "WebView2 installer exited with $($process.ExitCode)" }
    else { Write-Ok "WebView2 installer finished - now $(Get-EdgeUpdateClientVersion $WebView2Guid)" }
}

function Find-Helper {
    <#
        The repo layout (..\Device) or the flat folder -ComputerName copies to.
        Neither there: this script runs on its own, so the helper is fetched from
        GitHub at $HelperCommit and only used when its SHA-256 matches.
    #>
    param([string] $Name)
    foreach ($path in (Join-Path (Join-Path (Split-Path $PSScriptRoot) 'Device') $Name), (Join-Path $PSScriptRoot $Name)) {
        if (Test-Path $path) { return $path }
    }
    if (-not $HelperHashes.ContainsKey($Name)) { return $null }

    $url    = 'https://raw.githubusercontent.com/sjkanon/M365-Scripts/{0}/scripts/Device/{1}' -f $HelperCommit, $Name
    $target = Join-Path $WorkingDir $Name
    try {
        New-Item -ItemType Directory -Path $WorkingDir -Force | Out-Null
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $url -OutFile $target -UseBasicParsing
        $hash = (Get-FileHash -Path $target -Algorithm SHA256).Hash
        if ($hash -ne $HelperHashes[$Name]) {
            throw "SHA-256 is $hash, expected $($HelperHashes[$Name]) - not used"
        }
        Write-Ok "$Name fetched from GitHub at $($HelperCommit.Substring(0, 7)), SHA-256 verified"
        return $target
    } catch {
        Remove-Item $target -Force -ErrorAction SilentlyContinue
        Write-Bad "$Name from $url - $($_.Exception.Message)"
        return $null
    }
}

function Invoke-Helper {
    <# Run a repo script in its own process, so its exit does not end this one. #>
    param([string] $Name, [string[]] $Arguments)
    $path = Find-Helper $Name
    if (-not $path) { Write-Bad "$Name not found locally and not verified from GitHub - skipped"; return 1 }
    Write-Step "$Name $($Arguments -join ' ')"
    # To the host, not the pipeline: the return value is the exit code alone.
    & $nativeShell -NoProfile -ExecutionPolicy Bypass -File $path @Arguments | Out-Host
    $code = $LASTEXITCODE
    if ($code -eq 0) { Write-Ok "$Name finished" } else { Write-Warn "$Name exited with $code" }
    return $code
}

# =====================================================================================
Write-Host ''
Write-Host ("  Session host image - {0} - {1}" -f $env:COMPUTERNAME, $(if ($changing) { 'prepare' } else { 'check only' })) -ForegroundColor Cyan

$exitCode = 0
try {
    $state = Invoke-Check
    $work  = @($Findings)

    # A prepare run always brings the apps to the newest build, finding or not: the
    # helpers only change what is behind, so an up-to-date host costs a check.
    if ($changing) {
        Write-Host ''
        Write-Host '  ==== Updating '.PadRight(80, '=') -ForegroundColor Cyan
        $fixes = @(@($work | ForEach-Object { $_.Fix }) + @('apps') | Where-Object { $_ } | Sort-Object -Unique)

        if ('teamsupdate' -in $fixes) { Set-RegistryDword $TeamsPath 'disableAutoUpdate' 1 'Teams updates centrally, at one build on every host' }
        if ('sca' -in $fixes) {
            if ($PSCmdlet.ShouldProcess("$C2RConfigPath\SharedComputerLicensing", 'set to 1')) {
                New-ItemProperty -Path $C2RConfigPath -Name 'SharedComputerLicensing' -Value '1' -PropertyType String -Force | Out-Null
                Write-Ok 'Shared Computer Activation on - users sign in to Office once more to activate'
            }
        }
        if ('webview2' -in $fixes) { Install-WebView2 }

        if (-not $SkipApps) {
            if ('apps' -in $fixes) {
                if ((Invoke-Helper 'Repair-AppxPackageStore.ps1' @('-Name', 'teams,outlook', '-Latest', '-Provision', '-RemoveOld', '-Confirm:$false')) -eq 1) { $exitCode = 1 }
            }
            if ('copilot' -in $fixes -or 'apps' -in $fixes) {
                if ((Invoke-Helper 'Repair-AppxPackageStore.ps1' @('-Name', 'copilot', '-Provision', '-Confirm:$false')) -eq 1) { $exitCode = 1 }
            }
            if ('teams' -in $fixes -or 'apps' -in $fixes) {
                if ((Invoke-Helper 'Update-TeamsClient.ps1' @('-AvdOptimizations', '-Confirm:$false')) -eq 1) { $exitCode = 1 }
            }
        } elseif (@($fixes | Where-Object { $_ -in 'apps', 'copilot', 'teams' }).Count -gt 0) {
            Write-Skip '-SkipApps: Teams, Outlook and Copilot left as they are'
        }
        if ('reboot' -in $fixes) { Write-Warn 'A reboot is still pending - reboot before the next run or the capture' }

        Write-Host ''
        Write-Host '  ==== Read back '.PadRight(80, '=') -ForegroundColor Cyan
        $state = Invoke-Check
        $work  = @($Findings)
    }

    Write-Host ''
    if ($work.Count -eq 0) {
        Write-Ok 'Nothing left to do'
    } else {
        Write-Host ("  {0} finding(s) left:" -f $work.Count) -ForegroundColor Yellow
        $work | Format-Table Area, Level, Message -AutoSize -Wrap | Out-String -Width 200 | Write-Host
        if ($changing) { $exitCode = 1 } elseif ($exitCode -eq 0) { $exitCode = 2 }
    }
} catch {
    Write-Bad $_.Exception.Message
    $exitCode = 1
    $state = $null
} finally {
    if ($changing) { Stop-Transcript | Out-Null }
}

if ($PassThru) {
    [PSCustomObject]@{
        ImageResult = $true
        Exit        = $exitCode
        Build       = Get-PropertyValue $state 'Build'
        FSLogix     = Get-PropertyValue $state 'FSLogix'
        WebView2    = Get-PropertyValue $state 'WebView2'
        Teams       = Get-PropertyValue $state 'Teams'
        Outlook     = Get-PropertyValue $state 'Outlook'
        Copilot     = Get-PropertyValue $state 'Copilot'
        Findings    = @($Findings).Count
    }
}
exit $exitCode
