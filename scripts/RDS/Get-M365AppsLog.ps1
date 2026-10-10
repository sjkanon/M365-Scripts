#Requires -Version 5.1
<#
.SYNOPSIS
    Collects everything about new Teams, new Outlook and Copilot on a session host into
    one folder (and zip): per user what is registered and running, the events, and what
    the watchdog (Watch-M365Apps.ps1) did and would conclude - to find out why an app did
    not work for someone and why the watchdog did or did not repair it.

.DESCRIPTION
    Read only: nothing is started, registered, repaired or closed. Run it on the host,
    elevated, as soon as possible after the complaint. It writes:

      summary.txt    Everything below as it appeared on screen.
      sessions.csv   Every signed-in user: session, since when, watched account or
                     customer, and whether the watchdog checks them yet (10 minutes).
      packages.csv   Per user and app: the package(s) registered for them, version,
                     status, files present, running in their session, and the verdict
                     the watchdog would reach - with BLIND SPOT where it would call it
                     fine while it is not tested at all.
      allusers.csv   Per app package, from Get-AppxPackage -AllUsers: every user the
                     package is known for and its install state - users who are signed
                     off included.
      events.csv     From the last -Hours: refused opens (TWinUI 5961), registrations
                     (AppXDeploymentServer, errors and 401/404), app model runtime
                     errors, crashes and hangs (Application Error 1000 / Hang 1002) of
                     the apps - with the user, the code, and whether the watchdog reads
                     that event at all.
      watchdog\      The watchdog's task (state, last run, result, history), config.json
                     with the webhook and token masked, state.json, and its logs and
                     repair logs from the window.

.PARAMETER User
    Only these users, by user name, UPN or DOMAIN\user - also when they are signed off
    (their events and package state are still there). Default: everyone.

.PARAMETER App
    Teams, Outlook, Copilot. Default: all three.

.PARAMETER Hours
    How far back to read events and logs. Default 24.

.PARAMETER OutputPath
    Folder the collection is written in. Default C:\Temp.

.PARAMETER WorkingDir
    The watchdog's folder. Default C:\IT\AppWatchdog.

.PARAMETER IncludeAppLogs
    Also copy the apps' own logs from each signed-in user in scope: Teams (LocalCache -
    gone at sign-out with FSLogix), Teams diagnostics in Downloads, new Outlook (only
    when the user turned on troubleshooting logging). They hold names and mail
    addresses, so they are left out unless asked for.

.PARAMETER NoZip
    Leave the folder, do not zip it.

.EXAMPLE
    # A customer says Copilot did not open this morning
    .\Get-M365AppsLog.ps1 -User jansen -App Copilot -Hours 12

.EXAMPLE
    # Everything on this host from the last two days
    .\Get-M365AppsLog.ps1 -Hours 48

.NOTES
    Author  : Sjoerd Kanon
    Platform: Windows 11 Enterprise multi-session / Windows Server with FSLogix, run as
              administrator (the watchdog's folder is System and Administrators only)
    Exit    : 0 collected, 1 failed
#>
[CmdletBinding()]
param (
    [string[]] $User,
    [ValidateSet('Teams', 'Outlook', 'Copilot')]
    [string[]] $App = @('Teams', 'Outlook', 'Copilot'),
    [ValidateRange(1, 336)]
    [int]      $Hours = 24,
    [string]   $OutputPath = 'C:\Temp',
    [string]   $WorkingDir = 'C:\IT\AppWatchdog',
    [switch]   $IncludeAppLogs,
    [switch]   $NoZip
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# -- Constants, the same as Watch-M365Apps.ps1 -----------------------------------------
$Apps = @(
    [PSCustomObject]@{ App = 'Teams';   Packages = @('MSTeams');                                         Exe = @('ms-teams.exe') }
    [PSCustomObject]@{ App = 'Outlook'; Packages = @('Microsoft.OutlookForWindows');                     Exe = @('olk.exe') }
    # copilotapp.exe is the unified Copilot app (Edge Update, 152.x and up), M365Copilot.exe the packaged one.
    [PSCustomObject]@{ App = 'Copilot'; Packages = @('Microsoft.MicrosoftOfficeHub', 'Microsoft.Copilot'); Exe = @('M365Copilot.exe', 'copilotapp.exe') }
)
$CopilotGuid      = '{C50565E9-CCCF-44B4-BA15-5AC5C6569197}'
$TaskName         = 'M365 App Watchdog'
$UserGraceMinutes = 10
$BenignAppxCodes  = @('0x80073D02', '0x80073CFB', '0x80073D06')
$MaxEvents        = 5000
$WebView2Guid     = '{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}'
$AppxTimeoutSeconds = 120
$MaxFileMB        = 50     # a copied log larger than this keeps only its last 50 MB

# -- Output ------------------------------------------------------------------------
function Write-Step { param([string] $Message) Write-Host ''; Write-Host "  $Message" -ForegroundColor Cyan }
function Write-Ok   { param([string] $Message) Write-Host "  [ OK ] $Message" -ForegroundColor Green }
function Write-Info { param([string] $Message) Write-Host "  [INFO] $Message" }
function Write-Warn { param([string] $Message) Write-Host "  [WARN] $Message" -ForegroundColor Yellow }
function Write-Bad  { param([string] $Message) Write-Host "  [FAIL] $Message" -ForegroundColor Red }

function Get-PropertyValue {
    <# Property access that returns $null instead of tripping Set-StrictMode. #>
    param($Object, [string] $Name)
    if ($null -eq $Object) { return $null }
    $prop = $Object.PSObject.Properties[$Name]
    if ($prop) { return $prop.Value }
    return $null
}

$script:failedSections = [System.Collections.Generic.List[string]]::new()
function Invoke-Section {
    <#
        One part of the collection: when it fails, say so and go on with the next -
        half a collection is worth more than none. Dot-sourced, so what it sets stays
        in the script scope for the parts after it.
    #>
    param([string] $SectionTitle, [scriptblock] $Body)
    # Held in the script scope: the body is dot-sourced here and may use any variable name.
    $script:currentSection = $SectionTitle
    Write-Step $SectionTitle
    try { . $Body }
    catch {
        Write-Bad "$($script:currentSection) - stopped: $($_.Exception.Message)"
        $script:failedSections.Add($script:currentSection)
    }
}

function Invoke-WithTimeout {
    <#
        A script block in its own Windows PowerShell process, given up after $Seconds:
        Get-AppxPackage -AllUsers and -User hang for minutes on a host whose AppX
        state store is in trouble - exactly the host this is run on. Returns flat
        objects (they cross a process boundary), or $null on a timeout.
    #>
    param([scriptblock] $ScriptBlock, [object[]] $ArgumentList = @(), [int] $Seconds = 120, [string] $What = 'command')
    $job = Start-Job -ScriptBlock $ScriptBlock -ArgumentList $ArgumentList
    try {
        if (Wait-Job -Job $job -Timeout $Seconds) {
            $result = @(Receive-Job -Job $job -ErrorAction SilentlyContinue)
            $errors = @($job.ChildJobs | ForEach-Object { $_.Error } | Where-Object { $_ } | ForEach-Object { "$_".Trim() } | Sort-Object -Unique)
            if ($errors.Count -gt 0) { Write-Warn ('{0} - {1}' -f $What, ($errors -join '; ')) }
            return ,$result
        }
        Write-Warn "$What did not answer within $Seconds s - skipped (the AppX state store may be stuck)"
        return $null
    } finally {
        Stop-Job -Job $job -ErrorAction SilentlyContinue
        Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
    }
}

function Get-EventSafe {
    <#
        Get-WinEvent that returns nothing instead of throwing: a channel that does not
        exist on this host (FSLogix not installed, an older build) ends Get-WinEvent
        with "The parameter is incorrect", whatever -ErrorAction says.
    #>
    param([hashtable] $Filter)
    try { return @(Get-WinEvent -FilterHashtable $Filter -MaxEvents $MaxEvents -ErrorAction Stop) }
    catch { return @() }
}

function Read-SharedText {
    <# A file another process still writes to (an open transcript, a live app log): Get-Content and Copy-Item refuse it. #>
    param([string] $Path, [long] $Tail = 0)
    $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete)
    try {
        if ($Tail -gt 0 -and $stream.Length -gt $Tail) { $stream.Seek(-$Tail, [IO.SeekOrigin]::End) | Out-Null }
        $reader = New-Object IO.StreamReader($stream, $true)
        return $reader.ReadToEnd()
    } finally {
        $stream.Dispose()
    }
}

function Copy-LimitedFile {
    <# Copy a log even while it is open, and only its end when it is larger than $MaxFileMB. #>
    param([string] $Path, [string] $Destination, [string] $Name)
    if (-not $Name) { $Name = Split-Path $Path -Leaf }
    $target = Join-Path $Destination $Name
    try {
        $length = (Get-Item -LiteralPath $Path).Length
        if ($length -le $MaxFileMB * 1MB) {
            try { Copy-Item -LiteralPath $Path -Destination $target -Force; return }
            catch { }   # locked: read it shared below
        }
        [IO.File]::WriteAllText($target, (Read-SharedText -Path $Path -Tail ($MaxFileMB * 1MB)))
    } catch {
        Write-Warn "Not copied: $Path - $($_.Exception.Message)"
    }
}

# -- Windows PowerShell, 64-bit, elevated --------------------------------------------
# The AppX cmdlets need Windows PowerShell; under WOW64 HKLM lands in WOW6432Node.
$nativeShell = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
    $nativeShell = Join-Path $env:WINDIR 'SysNative\WindowsPowerShell\v1.0\powershell.exe'
}
if ($PSVersionTable.PSEdition -eq 'Core' -or ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess)) {
    $quote = { param($s) "'" + ([string] $s -replace "'", "''") + "'" }
    $line  = '& ' + (& $quote $PSCommandPath)
    foreach ($entry in $PSBoundParameters.GetEnumerator()) {
        if ($entry.Value -is [switch]) { $line += ' -{0}:${1}' -f $entry.Key, ([bool] $entry.Value).ToString().ToLower() }
        elseif ($entry.Value -is [array]) { $line += ' -{0} {1}' -f $entry.Key, (($entry.Value | ForEach-Object { & $quote $_ }) -join ',') }
        else { $line += ' -{0} {1}' -f $entry.Key, (& $quote $entry.Value) }
    }
    & $nativeShell -NoProfile -ExecutionPolicy Bypass -Command ($line + '; exit $LASTEXITCODE')
    exit $LASTEXITCODE
}
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error 'Administrator rights are required: the watchdog folder and other users'' packages are not readable otherwise.'
    exit 1
}

$Apps = @($Apps | Where-Object { $_.App -in $App })

# -- Readers -------------------------------------------------------------------------
function Get-EdgeUpdateClientVersion {
    param([string] $Guid)
    foreach ($root in 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients', 'HKLM:\SOFTWARE\Microsoft\EdgeUpdate\Clients') {
        $pv = Get-PropertyValue (Get-ItemProperty (Join-Path $root $Guid) -ErrorAction SilentlyContinue) 'pv'
        if ($pv -and $pv -ne '0.0.0.0') { return $pv }
    }
    return $null
}

function Get-UserSession {
    <# Everyone signed in: the owner of each explorer.exe, as in the watchdog. #>
    param([string[]] $Watched)
    $seen = @{}
    foreach ($proc in @(Get-CimInstance -ClassName Win32_Process -Filter "Name = 'explorer.exe'" -ErrorAction SilentlyContinue)) {
        $owner = Invoke-CimMethod -InputObject $proc -MethodName GetOwner -ErrorAction SilentlyContinue
        if (-not $owner -or $owner.ReturnValue -ne 0) { continue }
        $sid = (Invoke-CimMethod -InputObject $proc -MethodName GetOwnerSid -ErrorAction SilentlyContinue).Sid
        if (-not $sid -or $seen.ContainsKey($sid)) { continue }
        $seen[$sid] = $true
        [PSCustomObject]@{
            User      = '{0}\{1}' -f $owner.Domain, $owner.User
            Name      = $owner.User
            Sid       = $sid
            SessionId = $proc.SessionId
            Since     = $proc.CreationDate
            Watched   = @($Watched | Where-Object { (($_ -split '\\')[-1] -split '@')[0] -ieq $owner.User }).Count -gt 0
        }
    }
}

function Get-ShortName {
    <# itceadmin from DOMAIN\itceadmin or itceadmin@contoso.com. #>
    param([string] $Name)
    return (($Name -split '\\')[-1] -split '@')[0]
}

function Resolve-UserName {
    <# A SID as a name: a signed-in user, else the profile folder, else the SID. #>
    param([string] $Sid)
    if (-not $Sid) { return '' }
    if ($Sid -eq 'S-1-5-18') { return 'SYSTEM' }
    if ($script:nameCache.ContainsKey($Sid)) { return $script:nameCache[$Sid] }
    $name = $null
    $session = @($script:sessions | Where-Object { $_.Sid -eq $Sid }) | Select-Object -First 1
    if ($session) { $name = $session.User }
    if (-not $name) {
        $profilePath = Get-PropertyValue (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$Sid" -ErrorAction SilentlyContinue) 'ProfileImagePath'
        if ($profilePath) { $name = Split-Path $profilePath -Leaf }
    }
    if (-not $name) {
        try { $name = ([Security.Principal.SecurityIdentifier] $Sid).Translate([Security.Principal.NTAccount]).Value } catch { $name = $Sid }
    }
    $script:nameCache[$Sid] = $name
    return $name
}

function Test-UserMatch {
    <# Whether a SID belongs to one of -User; true for everything without -User. #>
    param([string] $Sid)
    if (-not $User) { return $true }
    if (-not $Sid) { return $false }
    $short = Get-ShortName (Resolve-UserName $Sid)
    return @($User | Where-Object { (Get-ShortName $_) -ieq $short -or $short -ilike "$(Get-ShortName $_).*" }).Count -gt 0
}

function Find-App {
    <# The watched app a package name, AUMID, exe or message text belongs to. #>
    param([string] $Text)
    foreach ($entry in $Apps) {
        foreach ($pkg in $entry.Packages) { if ($Text -like "*$($pkg)_*" -or $Text -like "*$($pkg)!*") { return $entry } }
        foreach ($exe in $entry.Exe) { if ($Text -like "*$exe*") { return $entry } }
    }
    return $null
}

function Format-HResult {
    param($Value)
    try { return '0x{0:X8}' -f ([long] $Value -band [long] 4294967295) } catch { return [string] $Value }
}

function Get-AppEvent {
    <#
        Every event about the apps since $Since, one row each, with whether the
        watchdog reads it: it reads TWinUI 5961, AppXDeploymentServer 401/404 minus
        the benign codes, and Application Error 1000 / Hang 1002 - nothing else.
    #>
    param([datetime] $Since)
    $sources = @(
        # 5960: activation blocked because the package state is Modified (0x80073CFC); 5961: refused.
        @{ Filter = @{ LogName = 'Microsoft-Windows-TWinUI/Operational'; Id = 5960, 5961; StartTime = $Since } }
        @{ Filter = @{ LogName = 'Microsoft-Windows-AppXDeploymentServer/Operational'; Id = 401, 404; StartTime = $Since } }
        @{ Filter = @{ LogName = 'Microsoft-Windows-AppXDeploymentServer/Operational'; Level = 1, 2, 3; StartTime = $Since } }
        @{ Filter = @{ LogName = 'Microsoft-Windows-AppXDeployment/Operational'; Level = 1, 2, 3; StartTime = $Since } }
        @{ Filter = @{ LogName = 'Microsoft-Windows-AppxPackaging/Operational'; Level = 1, 2, 3; StartTime = $Since } }
        @{ Filter = @{ LogName = 'Microsoft-Windows-AppReadiness/Admin'; Level = 1, 2, 3; StartTime = $Since } }
        @{ Filter = @{ LogName = 'Microsoft-Windows-AppReadiness/Operational'; Level = 1, 2, 3; StartTime = $Since } }
        @{ Filter = @{ LogName = 'Microsoft-Windows-AppModel-Runtime/Admin'; Level = 1, 2, 3; StartTime = $Since } }
        @{ Filter = @{ LogName = 'Application'; ProviderName = 'Application Error', 'Application Hang'; Id = 1000, 1002; StartTime = $Since } }
    )
    $seen = @{}
    foreach ($source in $sources) {
        $events = @(Get-EventSafe $source.Filter)
        foreach ($record in $events) {
            if ($seen.ContainsKey("$($record.LogName)|$($record.RecordId)")) { continue }
            $seen["$($record.LogName)|$($record.RecordId)"] = $true
            $values  = @($record.Properties | ForEach-Object { $_.Value })
            $message = ''
            try { $message = [string] $record.FormatDescription() } catch { }
            if (-not $message) { $message = ($values | ForEach-Object { [string] $_ }) -join ' | ' }
            $entry = Find-App (($values | ForEach-Object { [string] $_ }) -join ' ')
            if (-not $entry) { $entry = Find-App $message }
            if (-not $entry) { continue }

            $code = ''
            $read = 'no - the watchdog does not read this event'
            switch ($record.LogName) {
                'Microsoft-Windows-TWinUI/Operational' {
                    if ($record.Id -eq 5961) {
                        if ($values.Count -ge 2) { $code = Format-HResult $values[1] }
                        $read = 'yes (WontOpen)'
                    }
                }
                'Microsoft-Windows-AppXDeploymentServer/Operational' {
                    if ($record.Id -in 401, 404) {
                        $at = if ($record.Id -eq 401) { 3 } else { 2 }
                        if ($values.Count -gt $at) { $code = Format-HResult $values[$at] }
                        $read = if ($code -in $BenignAppxCodes) { 'no - benign code, skipped on purpose' } else { 'yes (RegisterFailed)' }
                    } else {
                        if ($message -match '0x8[0-9A-Fa-f]{7}') { $code = $Matches[0].ToUpper() -replace '^0X', '0x' }
                        if ($code -in $BenignAppxCodes) { $read = 'no - benign code' }
                    }
                }
                'Application' {
                    if ($record.Id -eq 1000 -and $values.Count -gt 6) { $code = '0x' + [string] $values[6] }
                    # The watchdog knows the unified Copilot app's process by neither package nor name.
                    $read = if ([string] $values[0] -ieq 'copilotapp.exe' -and -not $script:knowsCopilotApp) { 'no - this watchdog does not know copilotapp.exe' } else { 'yes (crash report, no repair)' }
                }
            }
            if (-not $code -and $message -match '0x8[0-9A-Fa-f]{7}') { $code = $Matches[0].ToUpper() -replace '^0X', '0x' }
            $sid = [string] $record.UserId
            if (-not (Test-UserMatch $sid) -and $record.LogName -ne 'Application') { continue }
            [PSCustomObject]@{
                Time     = $record.TimeCreated
                App      = $entry.App
                Log      = $record.LogName -replace '^Microsoft-Windows-', ''
                Id       = $record.Id
                Level    = $record.LevelDisplayName
                User     = if ($record.LogName -eq 'Application') { '(not in event)' } else { Resolve-UserName $sid }
                Code     = $code
                Watchdog = $read
                Message  = (($message -split "`r?`n") | Where-Object { $_.Trim() } | Select-Object -First 2) -join ' / '
            }
        }
    }
}

function Get-UserAppRow {
    <#
        One user and app as the watchdog sees it (Test-UserApp), plus what it does not
        look at: whether the app is running, and which package actually answered.
    #>
    param($Session, $Entry, [string] $Unified, [bool] $Settled, $Registered)
    if ($null -eq $Registered) {
        return [PSCustomObject]@{ User = $Session.User; Kind = $(if ($Session.Watched) { 'watched' } else { 'customer' }); App = $Entry.App
                                  Packages = '?'; Running = '?'; Watchdog = '?'; Note = 'Get-AppxPackage did not answer for this user - the watchdog hangs on this too' }
    }
    $registered = @($Registered | Where-Object { $_.Name -in $Entry.Packages })
    $exeNames = @($Entry.Exe | ForEach-Object { [IO.Path]::GetFileNameWithoutExtension($_) })
    $running  = @(Get-Process -Name $exeNames -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq $Session.SessionId })

    # The package the watchdog picks: the first name in the list that is registered.
    $pick = $null
    foreach ($name in $Entry.Packages) {
        $pick = @($registered | Where-Object { $_.Name -eq $name }) | Sort-Object { [version] $_.Version } -Descending | Select-Object -First 1
        if ($pick) { break }
    }
    $customer = -not $Session.Watched
    $verdict  = ''
    $note     = ''
    if (-not $pick) {
        if ($Entry.App -eq 'Copilot' -and $Unified) {
            $verdict = 'OK'
            $note    = "BLIND SPOT: no Copilot package for this user, but the unified app ($Unified) is on the host, so the watchdog calls it fine without testing it for them"
        } else {
            $verdict = 'NotRegistered'
        }
    } else {
        $location = Get-PropertyValue $pick 'InstallLocation'
        $status   = [string] (Get-PropertyValue $pick 'Status')
        if (-not $location -or -not (Test-Path -LiteralPath $location)) { $verdict = 'Broken (files gone)' }
        elseif ($status -and $status -ne 'Ok') { $verdict = "Broken (status $status)" }
        else { $verdict = 'OK' }
        if ($Entry.App -eq 'Copilot' -and $pick.Name -eq 'Microsoft.Copilot') {
            $note = 'BLIND SPOT: only consumer Copilot (Microsoft.Copilot) is registered, not Microsoft 365 Copilot (MicrosoftOfficeHub) - the watchdog counts either as Copilot'
        }
        if ($verdict -eq 'OK' -and $customer -and $running.Count -eq 0 -and -not $note) {
            $note = 'Registered and Ok, not running: for a customer the watchdog only sees an open that Windows refused (TWinUI 5961), not an app that starts and then fails'
        }
    }
    if ($customer -and -not $Settled) { $note = (@("Signed in less than $UserGraceMinutes minutes ago - the watchdog leaves them alone until then", $note) | Where-Object { $_ }) -join '; ' }

    [PSCustomObject]@{
        User       = $Session.User
        Kind       = if ($customer) { 'customer' } else { 'watched' }
        App        = $Entry.App
        Packages   = (@($registered | ForEach-Object { '{0} {1} [{2}]' -f $_.Name, $_.Version, (Get-PropertyValue $_ 'Status') }) -join '; ')
        Running    = if ($running.Count -gt 0) { (@($running | ForEach-Object { '{0} ({1})' -f $_.ProcessName, $_.Id }) -join ', ') } else { 'no' }
        Watchdog   = $verdict
        Note       = $note
    }
}

# =====================================================================================
$runStart = Get-Date
$since    = $runStart.AddHours(-$Hours)
$folderName = 'M365AppsLog_{0}_{1:yyyyMMdd-HHmm}' -f $env:COMPUTERNAME, $runStart
$folder   = $null
# -OutputPath first; when that cannot be written (no C:\Temp, a full disk), the temp folder.
foreach ($base in @($OutputPath, $env:TEMP) | Select-Object -Unique) {
    try {
        $candidate = Join-Path $base $folderName
        New-Item -ItemType Directory -Path $candidate -Force | Out-Null
        [IO.File]::WriteAllText((Join-Path $candidate '.write-test'), '')
        Remove-Item (Join-Path $candidate '.write-test') -Force
        $folder = $candidate
        break
    } catch {
        Write-Warn "Cannot write to $base - $($_.Exception.Message)"
    }
}
if (-not $folder) { Write-Error 'No folder to write the collection to.'; exit 1 }

$transcribing = $false
try { Start-Transcript -Path (Join-Path $folder 'summary.txt') | Out-Null; $transcribing = $true }
catch { Write-Warn "No summary.txt: $($_.Exception.Message)" }

$script:nameCache = @{}
$script:sessions  = @()
$script:knowsCopilotApp = $false
$config   = $null
$watched  = @('itceadmin')
$unified  = $null
try {
    Write-Host ''
    Write-Host ('  M365 apps log - {0} - {1:yyyy-MM-dd HH:mm} - last {2} h{3}' -f $env:COMPUTERNAME, $runStart, $Hours, $(if ($User) { ' - ' + ($User -join ', ') } else { '' })) -ForegroundColor Cyan
    $os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    if ($os) { Write-Info ('{0} {1} (build {2}), up since {3:yyyy-MM-dd HH:mm}' -f $os.Caption, $os.Version, $os.BuildNumber, $os.LastBootUpTime) }

    $watchDir = Join-Path $folder 'watchdog'
    New-Item -ItemType Directory -Path $watchDir -Force | Out-Null

    Invoke-Section 'Watchdog settings' {
        $configPath = Join-Path $WorkingDir 'config.json'
        if (-not (Test-Path $configPath)) { Write-Bad "No $configPath - the watchdog is not installed on this host (or in another -WorkingDir)"; return }
        $script:config = Get-Content -Path $configPath -Raw | ConvertFrom-Json
        # The webhook and its token stay on the host: only whether they are set.
        $masked = $config | Select-Object *
        $url = Get-PropertyValue $masked 'WebhookUrl'
        if ($url) { $masked.WebhookUrl = $(try { ([uri] $url).Host + ' (path masked)' } catch { '(set, masked)' }) }
        if (Get-PropertyValue $masked 'WebhookToken') { $masked.WebhookToken = '(set, masked)' }
        $masked | ConvertTo-Json | Set-Content -Path (Join-Path $watchDir 'config.masked.json') -Encoding UTF8
        if (Get-PropertyValue $config 'Account') { $script:watched = @(Get-PropertyValue $config 'Account') }
        Write-Info ('Accounts {0}; apps {1}; NoRepair {2}; NoUserRepair {3}; SkipLaunchTest {4}; cooldown {5} h; webhook {6}' -f `
            ($watched -join ','), ((@(Get-PropertyValue $config 'App')) -join ','),
            [bool] (Get-PropertyValue $config 'NoRepair'), [bool] (Get-PropertyValue $config 'NoUserRepair'),
            [bool] (Get-PropertyValue $config 'SkipLaunchTest'), (Get-PropertyValue $config 'RepairCooldownHours'),
            $(if ($url) { 'set' } else { 'NOT set - nothing is reported' }))
        if (Get-PropertyValue $config 'NoRepair')     { Write-Warn 'NoRepair is on: the watchdog repairs nothing on this host' }
        if (Get-PropertyValue $config 'NoUserRepair') { Write-Warn 'NoUserRepair is on: nothing is repaired in a customer''s session, only reported' }
        foreach ($entry in $Apps) {
            if (-not ((@(Get-PropertyValue $config 'App')) -contains $entry.App)) { Write-Warn "$($entry.App) is not among the watched apps" }
        }
        $self = Join-Path $WorkingDir 'Watch-M365Apps.ps1'
        if (Test-Path $self) {
            $item = Get-Item $self
            Write-Info ('Installed watchdog: {0:yyyy-MM-dd HH:mm}, SHA-256 {1}' -f $item.LastWriteTime, (Get-FileHash $self -Algorithm SHA256).Hash)
            # Older versions only test our own account: then customers were never looked at.
            # Watchdogs from before 2026-10-10 do not know the unified Copilot app's process.
            $script:knowsCopilotApp = [bool] (Select-String -Path $self -Pattern 'copilotapp.exe' -SimpleMatch -Quiet)
            if (-not (Select-String -Path $self -Pattern 'Get-OpenFailure' -SimpleMatch -Quiet)) {
                Write-Bad 'This watchdog is an older version without the per-user checks (step 2b): customers are not checked or repaired - reinstall from the readme'
            }
        } else {
            Write-Bad "No $self"
        }
    }

    Invoke-Section 'Watchdog task' {
        $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
        if (-not $task) { Write-Bad "Scheduled task '$TaskName' not found"; return }
        $info = Get-ScheduledTaskInfo -TaskName $TaskName
        $interval = @($task.Triggers | ForEach-Object { Get-PropertyValue (Get-PropertyValue $_ 'Repetition') 'Interval' } | Where-Object { $_ }) | Select-Object -First 1
        $line = 'Task {0}: last run {1:yyyy-MM-dd HH:mm}, result {2} (0x{2:X}), next {3:yyyy-MM-dd HH:mm}, every {4}' -f $task.State, $info.LastRunTime, $info.LastTaskResult, $info.NextRunTime, $(if ($interval) { $interval } else { '?' })
        if ($task.State -eq 'Disabled') { Write-Bad $line }
        elseif ($info.LastRunTime -lt $runStart.AddHours(-2)) { Write-Bad "$line - has not run for over 2 hours" }
        elseif ($info.LastTaskResult -notin 0, 1, 267009) { Write-Warn "$line - 0 is healthy, 1 is problems left, 267009 is running" }
        else { Write-Ok $line }
        $info | Select-Object TaskName, LastRunTime, LastTaskResult, NextRunTime, NumberOfMissedRuns |
            Export-Csv -Path (Join-Path $watchDir 'task.csv') -NoTypeInformation -Encoding UTF8
        $history = @(Get-EventSafe @{ LogName = 'Microsoft-Windows-TaskScheduler/Operational'; StartTime = $since } |
                     Where-Object { $_.Message -like "*\$TaskName*" })
        if ($history.Count -gt 0) {
            $history | Select-Object TimeCreated, Id, LevelDisplayName, Message | Export-Csv -Path (Join-Path $watchDir 'task-history.csv') -NoTypeInformation -Encoding UTF8
            Write-Info "$($history.Count) task history events (task-history.csv)"
        } else {
            Write-Info 'No task history (the TaskScheduler/Operational log is off by default)'
        }
        # One-off tasks the watchdog leaves behind only when it was killed halfway.
        $probes = @(Get-ScheduledTask -TaskPath '\M365AppWatchdog\' -ErrorAction SilentlyContinue)
        if ($probes.Count -gt 0) { Write-Warn "$($probes.Count) probe task(s) left in \M365AppWatchdog\ - a run was cut off" }
    }

    Invoke-Section 'Watchdog state and logs' {
        $statePath = Join-Path $WorkingDir 'state.json'
        if (Test-Path $statePath) {
            Copy-Item -Path $statePath -Destination $watchDir
            $state = Get-Content -Path $statePath -Raw | ConvertFrom-Json
            Write-Info ('Last run {0}, last host repair {1}, last reported {2}' -f (Get-PropertyValue $state 'LastRun'), (Get-PropertyValue $state 'LastRepair'), (Get-PropertyValue $state 'LastNotified'))
            $signature = Get-PropertyValue $state 'Signature'
            if ($signature) { Write-Warn "Problems still open at the last run: $signature" }
            $repairs = Get-PropertyValue $state 'UserRepairs'
            if ($repairs) {
                foreach ($p in $repairs.PSObject.Properties) {
                    Write-Info ('Re-registered in a user''s session: {0} {1} at {2}' -f (Resolve-UserName (($p.Name -split '\|')[0])), (($p.Name -split '\|')[-1]), $p.Value)
                }
            }
        }

        # The logs of the window, and the lines about these apps and users pulled out.
        $logs = @()
        foreach ($sub in 'Logs', 'Repair') {
            $dir = Join-Path $WorkingDir $sub
            if (-not (Test-Path $dir)) { continue }
            $target = Join-Path $watchDir $sub
            New-Item -ItemType Directory -Path $target -Force | Out-Null
            foreach ($file in @(Get-ChildItem -Path $dir -File -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -ge $since })) {
                Copy-LimitedFile -Path $file.FullName -Destination $target
                if ($sub -eq 'Logs') { $logs += $file }
            }
        }
        if ($logs.Count -eq 0) {
            if (Test-Path $WorkingDir) { Write-Warn "No watchdog log in $WorkingDir\Logs from the last $Hours h" }
            return
        }
        $runs  = 0
        $lines = [System.Collections.Generic.List[string]]::new()
        $words = @(@($Apps | ForEach-Object { $_.App }) + @($User | Where-Object { $_ } | ForEach-Object { Get-ShortName $_ }))
        foreach ($file in $logs | Sort-Object LastWriteTime) {
            # A transcript that is still open is locked for Get-Content: read it shared.
            $text = Read-SharedText $file.FullName
            foreach ($line in @($text -split "`r?`n")) {
                if ($line -match 'M365 app watchdog - ') { $runs++; $lines.Add($line.Trim()); continue }
                if ($line -notmatch '\[(FAIL|WARN)\]|re-registered|reset|Repair-AppxPackageStore|signed off|running for them|not repaired|Reported|Aborted') { continue }
                if (@($words | Where-Object { $line -like "*$_*" }).Count -eq 0 -and $line -notmatch 'Repair-AppxPackageStore|Reported|Aborted') { continue }
                $lines.Add('    ' + $line.Trim())
            }
        }
        Write-Info "$runs watchdog runs in $($logs.Count) log file(s); what they said about $($words -join ', ') (last 80 lines):"
        $lines | Select-Object -Last 80 | ForEach-Object { Write-Host "    $_" }
    }

    Invoke-Section 'Host' {
        $provisioned = @(Get-AppxProvisionedPackage -Online | Where-Object { $p = $_.DisplayName; @($Apps | Where-Object { $p -in $_.Packages }).Count -gt 0 })
        foreach ($entry in $Apps) {
            $have = @($provisioned | Where-Object { $_.DisplayName -in $entry.Packages })
            if ($have.Count -gt 0) { Write-Ok ('{0} provisioned: {1}' -f $entry.App, ((@($have | ForEach-Object { "$($_.DisplayName) $($_.Version)" })) -join ', ')) }
            else { Write-Warn "$($entry.App) not provisioned for new profiles" }
        }
        if (@($Apps | Where-Object { $_.App -eq 'Copilot' }).Count -gt 0) {
            $script:unified = Get-EdgeUpdateClientVersion $CopilotGuid
            if ($unified) { Write-Info "Copilot unified app (Edge Update) $unified - machine-wide, not a package per user; the watchdog accepts it as Copilot for everyone" }
            else { Write-Info 'No Copilot unified app (Edge Update)' }
        }
        $webView = Get-EdgeUpdateClientVersion $WebView2Guid
        if ($webView) { Write-Info "WebView2 runtime $webView" } else { Write-Warn 'No WebView2 runtime found - Teams, new Outlook and Copilot all need it' }
    }

    Invoke-Section 'Packages for all users' {
        $names = @($Apps | ForEach-Object { $_.Packages })
        $raw = Invoke-WithTimeout -What 'Get-AppxPackage -AllUsers' -Seconds $AppxTimeoutSeconds -ArgumentList (, $names) -ScriptBlock {
            param($Names)
            foreach ($name in $Names) {
                foreach ($pkg in @(Get-AppxPackage -AllUsers -Name $name -ErrorAction SilentlyContinue)) {
                    foreach ($info in @($pkg.PackageUserInformation)) {
                        [PSCustomObject]@{ Name = $pkg.Name; Package = $pkg.PackageFullName; Sid = [string] $info.UserSecurityId.Sid; InstallState = [string] $info.InstallState }
                    }
                }
            }
        }
        if ($null -eq $raw) { return }
        $allUsers = @(foreach ($row in $raw) {
            if (-not (Test-UserMatch $row.Sid)) { continue }
            $entry = @($Apps | Where-Object { $row.Name -in $_.Packages }) | Select-Object -First 1
            [PSCustomObject]@{ App = $entry.App; Package = $row.Package; User = Resolve-UserName $row.Sid; Sid = $row.Sid; InstallState = $row.InstallState }
        })
        $allUsers | Export-Csv -Path (Join-Path $folder 'allusers.csv') -NoTypeInformation -Encoding UTF8
        Write-Info "$($allUsers.Count) user/package registrations (allusers.csv)"
        foreach ($row in @($allUsers | Where-Object { $_.InstallState -and $_.InstallState -ne 'Installed' })) {
            Write-Warn ('{0} - {1}: {2} ({3})' -f $row.User, $row.App, $row.InstallState, $row.Package)
        }
        foreach ($name in @($User | Where-Object { $_ })) {
            foreach ($entry in $Apps) {
                if (@($allUsers | Where-Object { $_.App -eq $entry.App -and (Get-ShortName $_.User) -ieq (Get-ShortName $name) }).Count -eq 0) {
                    Write-Warn "$name - $($entry.App): no package known for this user on this host"
                }
            }
        }
    }

    Invoke-Section 'Signed-in users' {
        $script:sessions = @(Get-UserSession -Watched $watched)
        $inScope = @($script:sessions | Where-Object { Test-UserMatch $_.Sid })
        Write-Info ('{0} signed in, {1} in scope' -f $script:sessions.Count, $inScope.Count)
        $script:sessions | Select-Object User, Sid, SessionId, Since, Watched |
            Export-Csv -Path (Join-Path $folder 'sessions.csv') -NoTypeInformation -Encoding UTF8
        foreach ($name in @($User | Where-Object { $_ })) {
            if (@($inScope | Where-Object { (Get-ShortName $_.User) -ieq (Get-ShortName $name) }).Count -eq 0) {
                Write-Warn "$name is not signed in: the watchdog cannot check or repair them now, only the events below say what happened"
            }
        }
        foreach ($name in $watched) {
            if (@($script:sessions | Where-Object { $_.Watched -and (Get-ShortName $_.User) -ieq (Get-ShortName $name) }).Count -eq 0) {
                Write-Warn "Watched account $name is not signed in here - the watchdog's launch test does not run on this host"
            }
        }

        $names = @($Apps | ForEach-Object { $_.Packages })
        $rows = @(foreach ($session in $inScope) {
            $settled = $session.Watched -or -not $session.Since -or $session.Since -lt $runStart.AddMinutes(-$UserGraceMinutes)
            # By DOMAIN\user, not SID - as the watchdog does, for Entra SIDs.
            $registered = Invoke-WithTimeout -What "Get-AppxPackage -User $($session.User)" -Seconds $AppxTimeoutSeconds -ArgumentList $session.User, $names -ScriptBlock {
                param($UserName, $Names)
                foreach ($name in $Names) {
                    Get-AppxPackage -User $UserName -Name $name -ErrorAction SilentlyContinue |
                        Select-Object Name, Version, PackageFullName, InstallLocation, @{ n = 'Status'; e = { [string] $_.Status } }
                }
            }
            foreach ($entry in $Apps) { Get-UserAppRow -Session $session -Entry $entry -Unified $unified -Settled $settled -Registered $registered }
        })
        $rows | Export-Csv -Path (Join-Path $folder 'packages.csv') -NoTypeInformation -Encoding UTF8
        foreach ($row in $rows) {
            $text = '{0} ({1}) - {2}: watchdog {3}, running {4}; {5}' -f $row.User, $row.Kind, $row.App, $row.Watchdog, $row.Running, $(if ($row.Packages) { $row.Packages } else { 'no package' })
            if ($row.Watchdog -ne 'OK') { Write-Bad $text } elseif ($row.Note) { Write-Warn $text } else { Write-Ok $text }
            if ($row.Note) { Write-Host "         $($row.Note)" -ForegroundColor Yellow }
        }
    }

    Invoke-Section ('Events since {0:yyyy-MM-dd HH:mm}' -f $since) {
        $events = @(Get-AppEvent -Since $since | Sort-Object Time)
        $events | Export-Csv -Path (Join-Path $folder 'events.csv') -NoTypeInformation -Encoding UTF8
        if ($events.Count -eq 0) { Write-Ok 'None about these apps' }
        foreach ($group in @($events | Group-Object App, Log, Id, User, Code, Watchdog)) {
            $first = $group.Group[0]
            $last  = $group.Group[-1]
            $text  = '{0} - {1} {2} {3}{4}: {5}x, last {6:dd-MM HH:mm} - watchdog reads it: {7}' -f $first.App, $first.Log, $first.Id, $first.User, $(if ($first.Code) { " $($first.Code)" } else { '' }), $group.Count, $last.Time, $first.Watchdog
            if ($first.Watchdog -like 'no -*' -and $first.Level -in 'Error', 'Fout', 'Erreur', 'Critical') { Write-Warn $text } else { Write-Info $text }
            Write-Host "         $($last.Message)" -ForegroundColor DarkGray
        }
    }

    Invoke-Section 'Registry' {
        $regDir = Join-Path $folder 'registry'
        New-Item -ItemType Directory -Path $regDir -Force | Out-Null
        $patterns = @($Apps | ForEach-Object { $_.Packages } | ForEach-Object { "$($_)_*" })

        # A non-zero PackageStatus is a package Windows marked as modified or bad: TWinUI 5960.
        $stateRoot = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModel\StateChange\PackageList'
        foreach ($key in @(Get-ChildItem -Path $stateRoot -ErrorAction SilentlyContinue)) {
            if (@($patterns | Where-Object { $key.PSChildName -like $_ }).Count -eq 0) { continue }
            $status = Get-PropertyValue (Get-ItemProperty -Path $key.PSPath -ErrorAction SilentlyContinue) 'PackageStatus'
            if ($status) { Write-Bad "$($key.PSChildName): PackageStatus $status - Windows marked this package bad (activation blocked, 0x80073CFC)" }
        }
        # A deprovisioned package is never given to a new profile again.
        $deprovisioned = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Appx\AppxAllUserStore\Deprovisioned'
        foreach ($key in @(Get-ChildItem -Path $deprovisioned -ErrorAction SilentlyContinue)) {
            if (@($patterns | Where-Object { $key.PSChildName -like $_ }).Count -gt 0) { Write-Bad "Deprovisioned: $($key.PSChildName) - new profiles will not get it" }
        }

        $keys = [ordered]@{
            'fslogix'           = 'HKLM\SOFTWARE\FSLogix'
            'fslogix-policy'    = 'HKLM\SOFTWARE\Policies\FSLogix'
            'edgeupdate-policy' = 'HKLM\SOFTWARE\Policies\Microsoft\EdgeUpdate'
            'teams'             = 'HKLM\SOFTWARE\Microsoft\Teams'
            'appx-deprovisioned' = 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Appx\AppxAllUserStore\Deprovisioned'
        }
        foreach ($name in $keys.Keys) {
            if (-not (Test-Path "Registry::$($keys[$name])")) { continue }
            # Through cmd, so a reg.exe complaint does not become a terminating error under Stop.
            & cmd.exe /c ('reg.exe export "{0}" "{1}" /y >nul 2>&1' -f $keys[$name], (Join-Path $regDir "$name.reg"))
            if ($LASTEXITCODE -ne 0) { Write-Warn "reg export $($keys[$name]) - exit code $LASTEXITCODE" }
        }
        $appx = Get-PropertyValue (Get-ItemProperty 'HKLM:\SOFTWARE\FSLogix\Profiles' -ErrorAction SilentlyContinue) 'InstallAppxPackages'
        if ($null -ne $appx -and $appx -eq 0) { Write-Warn 'FSLogix InstallAppxPackages is 0: FSLogix does not register the apps again at sign-in' }
        foreach ($policy in 'CopilotUnificationAllowed', 'PauseCopilotAppUnificationRollout') {
            $value = Get-PropertyValue (Get-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\EdgeUpdate' -ErrorAction SilentlyContinue) $policy
            if ($null -ne $value) { Write-Info "Edge Update policy $policy = $value" }
        }
        Write-Info "Exported to registry\ (missing keys are left out)"
    }

    Invoke-Section 'FSLogix and Edge Update logs' {
        $fsDir = Get-PropertyValue (Get-ItemProperty 'HKLM:\SOFTWARE\FSLogix\Logging' -ErrorAction SilentlyContinue) 'LogDir'
        if (-not $fsDir) { $fsDir = Join-Path $env:ProgramData 'FSLogix\Logs' }
        $sources = [ordered]@{ 'fslogix' = $fsDir; 'edgeupdate' = (Join-Path $env:ProgramData 'Microsoft\EdgeUpdate\Log') }
        foreach ($name in $sources.Keys) {
            $dir = $sources[$name]
            if (-not (Test-Path $dir)) { Write-Info "No $dir"; continue }
            $files = @(Get-ChildItem -Path $dir -File -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -ge $since })
            if ($files.Count -eq 0) { Write-Info "Nothing in $dir from the last $Hours h"; continue }
            $target = Join-Path $folder $name
            foreach ($file in $files) {
                $relative = $file.FullName.Substring($dir.TrimEnd('\').Length).TrimStart('\')
                $parent   = Split-Path $relative -Parent
                $sub      = if ($parent) { Join-Path $target $parent } else { $target }
                New-Item -ItemType Directory -Path $sub -Force | Out-Null
                Copy-LimitedFile -Path $file.FullName -Destination $sub
            }
            Write-Info "$($files.Count) file(s) from $dir"
        }
        # FSLogix writes per user which packages it registered at sign-in and how long it took.
        $words = @($Apps | ForEach-Object { $_.Packages })
        $profileLogs = @(Get-ChildItem -Path (Join-Path $fsDir 'Profile') -Filter '*.log' -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -ge $since })
        $hits = @(foreach ($file in $profileLogs) {
            foreach ($line in @((Read-SharedText $file.FullName) -split "`r?`n")) {
                if (@($words | Where-Object { $line -like "*$_*" }).Count -gt 0 -and $line -match 'ERROR|fail|WARN|0x8') { $line.Trim() }
            }
        })
        if ($hits.Count -gt 0) {
            Write-Warn "FSLogix Profile log lines about the apps with an error ($($hits.Count), last 20):"
            $hits | Select-Object -Last 20 | ForEach-Object { Write-Host "    $_" }
        }
        $fsEvents = @(foreach ($log in 'Microsoft-FSLogix-Apps/Admin', 'Microsoft-FSLogix-Apps/Operational') {
            Get-EventSafe @{ LogName = $log; Level = 1, 2, 3; StartTime = $since }
        })
        if ($fsEvents.Count -gt 0) {
            $fsEvents | Select-Object TimeCreated, LogName, Id, LevelDisplayName, Message | Export-Csv -Path (Join-Path $folder 'fslogix-events.csv') -NoTypeInformation -Encoding UTF8
            Write-Warn "$($fsEvents.Count) FSLogix error/warning event(s) (fslogix-events.csv)"
        }
    }

    Invoke-Section 'Copilot app' {
        if (@($Apps | Where-Object { $_.App -eq 'Copilot' }).Count -eq 0) { Write-Info 'Copilot not asked for'; return }
        foreach ($root in 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall') {
            foreach ($key in @(Get-ChildItem -Path $root -ErrorAction SilentlyContinue)) {
                $p = Get-ItemProperty -Path $key.PSPath -ErrorAction SilentlyContinue
                if ([string] (Get-PropertyValue $p 'DisplayName') -notlike '*Copilot*') { continue }
                Write-Info ('Installed: {0} {1} in {2}' -f (Get-PropertyValue $p 'DisplayName'), (Get-PropertyValue $p 'DisplayVersion'), (Get-PropertyValue $p 'InstallLocation'))
            }
        }
        foreach ($proc in @(Get-Process -Name 'copilotapp', 'M365Copilot' -ErrorAction SilentlyContinue)) {
            $owner = @($script:sessions | Where-Object { $_.SessionId -eq $proc.SessionId }) | Select-Object -First 1
            $path  = $null
            try { $path = $proc.Path } catch { }
            Write-Info ('Running: {0} for {1} (session {2}) - {3}' -f $proc.ProcessName, $(if ($owner) { $owner.User } else { '?' }), $proc.SessionId, $path)
        }
        if (@(Get-Process -Name 'copilotapp' -ErrorAction SilentlyContinue).Count -gt 0) {
            if ($script:knowsCopilotApp) { Write-Warn 'The unified Copilot app (copilotapp.exe) is in use here: the watchdog sees its crashes, but does not test per user whether it opens' }
            else { Write-Warn 'The unified Copilot app (copilotapp.exe) is in use here: this watchdog checks neither its start nor its crashes' }
        }
    }

    if ($IncludeAppLogs) {
        Invoke-Section 'App logs per user' {
            foreach ($session in @($script:sessions | Where-Object { Test-UserMatch $_.Sid })) {
                $profilePath = Get-PropertyValue (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$($session.Sid)" -ErrorAction SilentlyContinue) 'ProfileImagePath'
                if (-not $profilePath -or -not (Test-Path $profilePath)) { Write-Warn "$($session.User): profile folder not found"; continue }
                $local  = Join-Path $profilePath 'AppData\Local'
                $target = Join-Path $folder ('applogs\' + ($session.User -replace '[\\/:*?"<>|]', '_'))
                $dirs = [ordered]@{
                    'teams'         = Join-Path $local 'Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\Logs'
                    'outlook'       = Join-Path $local 'Microsoft\Olk\logs'
                    'fslogix-appx'  = Join-Path $local 'FSLogix'
                }
                $count = 0
                foreach ($name in $dirs.Keys) {
                    if (-not (Test-Path $dirs[$name])) { continue }
                    # Logs only: Teams keeps model caches and settings under Logs as well.
                    $files = @(Get-ChildItem -Path $dirs[$name] -File -Recurse -ErrorAction SilentlyContinue |
                               Where-Object { $_.LastWriteTime -ge $since -and $_.Extension -in '.log', '.txt', '.xml', '.json' } |
                               Sort-Object LastWriteTime -Descending | Select-Object -First 200)
                    if ($files.Count -eq 0) { continue }
                    New-Item -ItemType Directory -Path (Join-Path $target $name) -Force | Out-Null
                    foreach ($file in $files) { Copy-LimitedFile -Path $file.FullName -Destination (Join-Path $target $name) -Name ($file.FullName.Substring($dirs[$name].Length).TrimStart('\') -replace '\\', '_') }
                    $count += $files.Count
                }
                # Ctrl+Alt+Shift+1 in Teams writes a diagnostics bundle to Downloads.
                foreach ($item in @(Get-ChildItem -Path (Join-Path $profilePath 'Downloads') -Filter 'MSTeams Diagnostics Log*' -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -ge $since })) {
                    New-Item -ItemType Directory -Path (Join-Path $target 'teams-diagnostics') -Force | Out-Null
                    Copy-Item -LiteralPath $item.FullName -Destination (Join-Path $target 'teams-diagnostics') -Recurse -Force -ErrorAction SilentlyContinue
                    $count++
                }
                Write-Info "$($session.User): $count app log file(s)"
            }
        }
    }

    Write-Host ''
    if ($script:failedSections.Count -gt 0) { Write-Warn "Incomplete - stopped in: $($script:failedSections -join ', ')" }
    Write-Ok "Collected in $folder"
} catch {
    Write-Bad "Aborted: $($_.Exception.Message)"
    $script:failedSections.Add('(run)')
} finally {
    if ($transcribing) { try { Stop-Transcript | Out-Null } catch { } }
}

if (-not $NoZip) {
    try {
        $zip = "$folder.zip"
        Compress-Archive -Path (Join-Path $folder '*') -DestinationPath $zip -Force
        Write-Ok "Zipped: $zip"
    } catch {
        Write-Warn "Not zipped, the folder is there: $($_.Exception.Message)"
    }
}
if ($script:failedSections -contains '(run)') { exit 1 }
exit 0
