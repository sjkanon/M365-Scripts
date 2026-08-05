<#
.SYNOPSIS
    FSLogix / Azure Files Premium diagnosescript voor AVD session hosts.

.DESCRIPTION
    Verzamelt in één run alle relevante informatie om FSLogix profielproblemen
    (mount errors, locked VHDX, SMB/Azure Files issues, disk errors) te analyseren.

.NOTES
    Uitvoeren in een verhoogde (Administrator) PowerShell sessie op de session host.
    Voorbeeld:  .\FSLogix-AVD-Diagnose.ps1 -User dedonder -Days 7 -OutputPath C:\Temp
#>

[CmdletBinding()]
param(
    [string]$User,
    [int]$Days = 7,
    [string]$OutputPath = "$env:SystemDrive\Temp\FSLogixDiag"
)

$ErrorActionPreference = 'Continue'
$since = (Get-Date).AddDays(-$Days)

if (-not (Test-Path $OutputPath)) { New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null }
$logFile = Join-Path $OutputPath ("FSLogixDiag_{0}_{1}.log" -f $env:COMPUTERNAME,(Get-Date -Format 'yyyyMMdd_HHmmss'))
Start-Transcript -Path $logFile -Force | Out-Null

function Write-Section {
    param([string]$Title)
    Write-Host ""
    Write-Host ("========== {0} ==========" -f $Title.ToUpper()) -ForegroundColor Cyan
}

function Write-Warn {
    param([string]$Message)
    Write-Host $Message -ForegroundColor Yellow
}

# ---------------------------------------------------------------------------
Write-Section "SYSTEEM"
# ---------------------------------------------------------------------------
[PSCustomObject]@{
    ComputerName = $env:COMPUTERNAME
    Tijdstip     = Get-Date
    OS           = (Get-CimInstance Win32_OperatingSystem).Caption
    Build        = (Get-CimInstance Win32_OperatingSystem).BuildNumber
    LastBoot     = (Get-CimInstance Win32_OperatingSystem).LastBootUpTime
    Uptime       = (New-TimeSpan -Start (Get-CimInstance Win32_OperatingSystem).LastBootUpTime -End (Get-Date)).ToString()
} | Format-List

# ---------------------------------------------------------------------------
Write-Section "FSLogix versie"
# ---------------------------------------------------------------------------
$frx = Join-Path $env:ProgramFiles 'FSLogix\Apps\frx.exe'
if (Test-Path $frx) {
    (Get-Item $frx).VersionInfo | Select-Object FileVersion, ProductVersion, FileName | Format-List
} else {
    Write-Warn "frx.exe niet gevonden op $frx - is FSLogix wel geinstalleerd?"
}

# ---------------------------------------------------------------------------
Write-Section "FSLogix config"
# ---------------------------------------------------------------------------
$profilesKey = 'HKLM:\SOFTWARE\FSLogix\Profiles'
$vhdLocations = @()

if (Test-Path $profilesKey) {
    $cfg = Get-ItemProperty $profilesKey
    $cfg | Select-Object Enabled, VHDLocations, VolumeType, SizeInMBs, IsDynamic,
                          DeleteLocalProfileWhenVHDShouldApply, FlipFlopProfileDirectoryName,
                          LockedRetryCount, LockedRetryInterval, ReAttachIntervalSeconds,
                          ReAttachRetryCount, ProfileType, CCDLocations |
        Format-List
    $vhdLocations = @($cfg.VHDLocations) | Where-Object { $_ }
} else {
    Write-Warn "Registry key $profilesKey bestaat niet."
}

Write-Host "--- Office Container (ODFC) ---"
$odfcKey = 'HKLM:\SOFTWARE\Policies\FSLogix\ODFC'
if (Test-Path $odfcKey) {
    Get-ItemProperty $odfcKey | Format-List
} else {
    Write-Host "Geen ODFC configuratie aanwezig."
}

# ---------------------------------------------------------------------------
Write-Section "FSLogix services"
# ---------------------------------------------------------------------------
Get-Service frxsvc, frxccds, frxdrv, frxdrvvt -ErrorAction SilentlyContinue |
    Select-Object Name, DisplayName, Status, StartType |
    Format-Table -AutoSize

# ---------------------------------------------------------------------------
Write-Section "Actieve VHD(X) containers"
# ---------------------------------------------------------------------------
# LET OP: Get-DiskImage vereist -ImagePath. Nooit los aanroepen (vraagt dan om input).
$mounted = Get-CimInstance -Namespace root\Microsoft\Windows\Storage -ClassName MSFT_DiskImage -ErrorAction SilentlyContinue |
           Where-Object { $_.ImagePath }

if ($mounted) {
    $mounted | Select-Object ImagePath, Attached, DevicePath, Size, StorageType |
        Format-Table -AutoSize -Wrap
} else {
    Write-Host "Geen gemounte disk images gevonden via CIM. Fallback via Get-Disk:"
    Get-Disk -ErrorAction SilentlyContinue |
        Where-Object BusType -eq 'File Backed Virtual' |
        ForEach-Object { Get-DiskImage -DevicePath $_.Path.TrimEnd('\') -ErrorAction SilentlyContinue } |
        Select-Object ImagePath, Attached, Number |
        Format-Table -AutoSize -Wrap
}

# ---------------------------------------------------------------------------
Write-Section "FSLogix sessies (registry)"
# ---------------------------------------------------------------------------
$sessionsKey = 'HKLM:\SOFTWARE\FSLogix\Profiles\Sessions'
if (Test-Path $sessionsKey) {
    Get-ChildItem $sessionsKey | ForEach-Object {
        $sid = Split-Path $_.Name -Leaf
        $acct = try {
            (New-Object System.Security.Principal.SecurityIdentifier($sid)).Translate([System.Security.Principal.NTAccount]).Value
        } catch { 'onbekend' }
        $p = Get-ItemProperty $_.PSPath
        [PSCustomObject]@{
            SID       = $sid
            Account   = $acct
            VHDPath   = $p.LocalProfilePath
            Status    = $p.Status
            Session   = $p.SessionId
        }
    } | Format-List
} else {
    Write-Host "Geen actieve FSLogix sessies in de registry."
}

# ---------------------------------------------------------------------------
Write-Section "Profielpaden (ProfileList)"
# ---------------------------------------------------------------------------
Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList' |
ForEach-Object {
    $sid = Split-Path $_.Name -Leaf
    $p = Get-ItemProperty $_.PSPath
    $acct = try {
        (New-Object System.Security.Principal.SecurityIdentifier($sid)).Translate([System.Security.Principal.NTAccount]).Value
    } catch { 'onbekend' }
    [PSCustomObject]@{ SID = $sid; Account = $acct; ProfilePath = $p.ProfileImagePath; State = $p.State }
} | Where-Object { $_.Account -notmatch 'NT AUTHORITY' } |
    Format-Table -AutoSize -Wrap

# ---------------------------------------------------------------------------
Write-Section "SMB connections (Azure Files)"
# ---------------------------------------------------------------------------
Get-SmbConnection -ErrorAction SilentlyContinue |
    Select-Object ServerName, ShareName, UserName, Dialect, Encrypted, NumOpens |
    Format-Table -AutoSize

Write-Host "--- SMB client config ---"
Get-SmbClientConfiguration | Select-Object EnableMultiChannel, SessionTimeout, ConnectionCountPerRssNetworkInterface, RequireSecuritySignature | Format-List

Write-Host "--- SMB open files (mogelijke locks) ---"
Get-SmbOpenFile -ErrorAction SilentlyContinue |
    Where-Object { $_.Path -like '*.vhd*' } |
    Select-Object ClientUserName, ClientComputerName, Path |
    Format-Table -AutoSize -Wrap

# ---------------------------------------------------------------------------
Write-Section "VHD share bereikbaarheid"
# ---------------------------------------------------------------------------
foreach ($loc in $vhdLocations) {
    Write-Host "Pad: $loc"
    if (Test-Path $loc) {
        Write-Host "  -> Bereikbaar." -ForegroundColor Green

        $server = ($loc -replace '^\\\\','').Split('\')[0]
        Test-NetConnection -ComputerName $server -Port 445 -WarningAction SilentlyContinue |
            Select-Object ComputerName, RemoteAddress, TcpTestSucceeded | Format-List

        Get-ChildItem $loc -Recurse -Filter *.VHDX -ErrorAction SilentlyContinue |
            Select-Object FullName, @{N='SizeGB';E={[math]::Round($_.Length/1GB,2)}}, LastWriteTime |
            Sort-Object LastWriteTime -Descending |
            Format-Table -AutoSize -Wrap
    } else {
        Write-Warn "  -> NIET bereikbaar vanaf deze host."
    }
}

# ---------------------------------------------------------------------------
Write-Section "FSLogix events (laatste $Days dagen)"
# ---------------------------------------------------------------------------
$fslogixLogs = @(
    'Microsoft-FSLogix-Apps/Operational',
    'Microsoft-FSLogix-Apps/Admin',
    'Microsoft-FSLogix-CloudCache/Operational'
)

foreach ($log in $fslogixLogs) {
    Write-Host "--- $log ---"
    $events = Get-WinEvent -FilterHashtable @{ LogName = $log; StartTime = $since } -ErrorAction SilentlyContinue
    if ($events) {
        $events | Group-Object Id, LevelDisplayName |
            Select-Object @{N='Id/Level';E={$_.Name}}, Count |
            Sort-Object Count -Descending | Format-Table -AutoSize

        $events | Where-Object { $_.LevelDisplayName -in 'Error','Warning' } |
            Sort-Object TimeCreated -Descending |
            Select-Object -First 50 TimeCreated, Id, LevelDisplayName, Message |
            Format-Table -AutoSize -Wrap
    } else {
        Write-Host "Geen events gevonden."
    }
}

# ---------------------------------------------------------------------------
Write-Section "Kritieke foutpatronen in FSLogix events"
# ---------------------------------------------------------------------------
$patterns = 'Failed to attach','Access is denied','cannot access the file','locked',
            'STATUS_SHARING_VIOLATION','Error 32','Error 5','corrupt','network path'

$hits = Get-WinEvent -FilterHashtable @{ LogName = 'Microsoft-FSLogix-Apps/Operational'; StartTime = $since } -ErrorAction SilentlyContinue |
        Where-Object { $m = $_.Message; $patterns | Where-Object { $m -match [regex]::Escape($_) } }

if ($hits) {
    $hits | Sort-Object TimeCreated -Descending |
        Select-Object -First 40 TimeCreated, Id, Message | Format-List
} else {
    Write-Host "Geen matches op de bekende foutpatronen." -ForegroundColor Green
}

# ---------------------------------------------------------------------------
Write-Section "Disk / NTFS errors (laatste $Days dagen)"
# ---------------------------------------------------------------------------
$sys = Get-WinEvent -FilterHashtable @{ LogName = 'System'; StartTime = $since } -ErrorAction SilentlyContinue |
       Where-Object { $_.Id -in 7,11,51,55,98,140,153 -and $_.ProviderName -match 'disk|Ntfs|volmgr|storahci|BitLocker' }

if ($sys) {
    $sys | Sort-Object TimeCreated -Descending |
        Select-Object TimeCreated, Id, ProviderName, Message |
        Format-Table -AutoSize -Wrap
} else {
    Write-Host "Geen disk/NTFS errors gevonden." -ForegroundColor Green
}

# ---------------------------------------------------------------------------
Write-Section "User Profile Service events"
# ---------------------------------------------------------------------------
Get-WinEvent -FilterHashtable @{ LogName = 'Application'; StartTime = $since } -ErrorAction SilentlyContinue |
    Where-Object { $_.ProviderName -match 'User Profile' } |
    Sort-Object TimeCreated -Descending |
    Select-Object -First 30 TimeCreated, Id, LevelDisplayName, Message |
    Format-Table -AutoSize -Wrap

# ---------------------------------------------------------------------------
Write-Section "Lokale profielen (C:\Users)"
# ---------------------------------------------------------------------------
Get-ChildItem 'C:\Users' -Directory |
    Select-Object Name, LastWriteTime, CreationTime |
    Sort-Object LastWriteTime -Descending |
    Format-Table -AutoSize

if (Get-ChildItem 'C:\Users' -Directory | Where-Object Name -match '^local_') {
    Write-Warn "Let op: 'local_*' profielen aanwezig - duidt op mislukte FSLogix mounts."
}

# ---------------------------------------------------------------------------
Write-Section "FSLogix logbestanden"
# ---------------------------------------------------------------------------
$logRoot = 'C:\ProgramData\FSLogix\Logs'
if (Test-Path $logRoot) {
    Get-ChildItem $logRoot -Recurse -File |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 20 FullName, LastWriteTime, @{N='KB';E={[math]::Round($_.Length/1KB,1)}} |
        Format-Table -AutoSize -Wrap

    Write-Host "--- Laatste ERROR/WARN regels uit Profile logs ---"
    Get-ChildItem (Join-Path $logRoot 'Profile') -Filter *.log -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 3 |
        ForEach-Object {
            Write-Host "[$($_.Name)]" -ForegroundColor Yellow
            Select-String -Path $_.FullName -Pattern 'ERROR','WARN','failed','denied','locked' -ErrorAction SilentlyContinue |
                Select-Object -Last 25 | ForEach-Object { $_.Line }
        }
} else {
    Write-Warn "$logRoot niet gevonden."
}

# ---------------------------------------------------------------------------
if ($User) {
    Write-Section "Gefilterd op gebruiker: $User"

    Get-WinEvent -FilterHashtable @{ LogName = 'Microsoft-FSLogix-Apps/Operational'; StartTime = $since } -ErrorAction SilentlyContinue |
        Where-Object { $_.Message -match $User } |
        Sort-Object TimeCreated -Descending |
        Select-Object -First 40 TimeCreated, Id, LevelDisplayName, Message |
        Format-List

    foreach ($loc in $vhdLocations) {
        Get-ChildItem $loc -Recurse -Filter *.VHDX -ErrorAction SilentlyContinue |
            Where-Object { $_.FullName -match $User } |
            Select-Object FullName, @{N='SizeGB';E={[math]::Round($_.Length/1GB,2)}}, LastWriteTime, IsReadOnly |
            Format-List
    }
}

# ---------------------------------------------------------------------------
Write-Section "Voltooid"
Write-Host "Rapport opgeslagen in: $logFile" -ForegroundColor Green
Stop-Transcript | Out-Null
