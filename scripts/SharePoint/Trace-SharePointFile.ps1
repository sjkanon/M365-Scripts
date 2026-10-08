#Requires -Version 5.1
<#
.SYNOPSIS
    Find out where a file in OneDrive or SharePoint went: renamed, moved, copied,
    deleted or restored, by whom and when - from the Unified Audit Log, in Brussels time.

.DESCRIPTION
    "The file is gone" usually means it was renamed, moved, or deleted along with
    the folder it was in. SharePoint itself does not remember any of that; the
    Unified Audit Log does, for OneDrive and SharePoint alike. This script reads
    it and reconstructs the trail of one file:

      1. Read    Every rename, move, copy, delete, recycle, restore and upload of
                 files AND folders in the window (Search-UnifiedAuditLog). The
                 window is read in slices; a slice that holds more than the
                 50,000 records one search can return is split until it fits,
                 and a failing or inconsistent search is retried.

      2. Seed    The records that mention the file by -Name, -Url or -ItemId.

      3. Follow  Every record of the same item (ListItemUniqueId, which survives
                 renames and moves within a site) and every record that starts at
                 a path the file was renamed or moved to. A chain A -> B -> C is
                 followed to C, even though C no longer resembles the name you
                 searched for. A move to another site gets a new item ID there;
                 it is followed by its destination path.

      4. Folders Renaming, moving or deleting a folder moves every file in it
                 WITHOUT an audit record per file. Folder records are therefore
                 replayed against the file's path at that moment, so "moved along
                 with folder X" and "deleted along with folder X" show up too.

    Every time shown and exported is in the -TimeZone (default Europe/Brussels,
    summer and winter time handled), with the UTC offset alongside. -StartDate
    and -EndDate are read as wall-clock time in that zone as well.

    Read-only: nothing in the tenant is changed. The result is printed per item
    (timeline, last known location, status) and exported to CSV, with the raw
    audit records of the trail in a JSON file next to it as evidence.

.PARAMETER Name
    File name to look for, as it was at some point. Case-insensitive. Without an
    extension it also matches the name with any extension ("Offerte" finds
    "Offerte.docx"). * and ? are wildcards ("Offerte*2026*.xlsx").

.PARAMETER Url
    Full URL of the file as it was (copy it from an old e-mail, a Teams message,
    the browser history). A "?web=1" query is ignored, and a sharing link of the
    form https://tenant.sharepoint.com/:w:/r/sites/... is turned back into the
    path. Opaque sharing links (/:w:/s/..., /:w:/g/...) cannot be traced.

.PARAMETER ItemId
    The ListItemUniqueId of the file, if you already know it (it is in the CSV of
    a previous run).

.PARAMETER SiteUrl
    Only consider records of this site or OneDrive (and below). Much faster in a
    large tenant. Moves out of this site are still seen at the source; what
    happens to the file at the destination site afterwards is not.

.PARAMETER StartDate
    Start of the window, as wall-clock time in -TimeZone. A [datetime], or a string
    in Belgian day-first notation or ISO: '15-09-2026', '15/09/2026 08:30',
    '2026-09-15', '2026-09-15 08:30'. A [datetime] from Get-Date is taken as the
    moment it denotes on this machine. Default: -Days before -EndDate.

.PARAMETER EndDate
    End of the window, same notation. A date without a time includes that whole
    day ('30-09-2026' runs up to 1 October 00:00). Default: now.

.PARAMETER Days
    Length of the window when -StartDate is not given. Default 30. Audit Standard
    keeps 180 days; Audit Premium (E5) longer.

.PARAMETER TimeZone
    IANA or Windows time zone ID for input and output. Default 'Europe/Brussels'
    (Windows: 'Romance Standard Time'); both are accepted in Windows PowerShell
    5.1 and PowerShell 7.

.PARAMETER FollowCopies
    Also follow copies of the file (FileCopied). Off by default: a copy leaves the
    original in place, so it is reported, but its own trail is not followed.

.PARAMETER IncludeActivity
    Also read opens, edits, downloads and sync activity (FileAccessed,
    FileModified, FileDownloaded, FileSyncDownloadedFull, FileSyncUploadedFull,
    FileCheckedIn, FileCheckedOut) - shows who last worked in the file. This
    multiplies the number of audit records and so the run time.

.PARAMETER OutputPath
    CSV path. Default C:\Temp\FileTrail_<name>_<timestamp>.csv (~/Downloads on
    macOS/Linux). The raw audit records go to the same name with .json.

.PARAMETER TenantId
    Tenant domain to read (e.g. contoso.onmicrosoft.com). Delegated, a partner
    reaches a customer with it (-DelegatedOrganization); left out, the GDAP
    customer from Connect-Tenant, else your own tenant. Optional when already
    connected to the right organisation.

.PARAMETER ClientId
    App-only instead of delegated: an app registration with Exchange.ManageAsApp
    and an Exchange role that includes View-Only Audit Logs. Needs
    -CertificateThumbprint and -TenantId (domain).

.PARAMETER CertificateThumbprint
    Certificate of that app, in CurrentUser\My.

.PARAMETER AppOnly
    App-only with the app registration for the tenant in graph.appid.json.

.PARAMETER PassThru
    Also return the timeline rows as objects.

.EXAMPLE
    # Where did "Offerte Janssens.docx" go in the last 30 days?
    .\Trace-SharePointFile.ps1 -Name "Offerte Janssens.docx" -TenantId contoso.onmicrosoft.com

.EXAMPLE
    # A period in Brussels time, one OneDrive only
    .\Trace-SharePointFile.ps1 -Name "Budget*" -StartDate '01-09-2026' -EndDate '15-09-2026' `
        -SiteUrl https://contoso-my.sharepoint.com/personal/jan_contoso_com

.EXAMPLE
    # From the link someone once sent, a specific afternoon
    .\Trace-SharePointFile.ps1 -Url "https://contoso.sharepoint.com/sites/Sales/Shared Documents/2026/Prijslijst.xlsx" `
        -StartDate '2026-09-12 13:00' -EndDate '2026-09-12 18:00'

.EXAMPLE
    # Include who opened / edited it, follow copies too
    .\Trace-SharePointFile.ps1 -Name Prijslijst.xlsx -Days 90 -IncludeActivity -FollowCopies

.NOTES
    Permissions: Search-UnifiedAuditLog needs the View-Only Audit Logs or Audit
    Logs role in Exchange Online (Organization Management, Compliance Management,
    or Global Reader / Compliance Administrator via Entra ID).

    Required module: ExchangeOnlineManagement (v3).

    Sign-in: delegated by default, through Connect-M365Exchange in
    scripts\Startup\Connect-M365.ps1 (device code per load.config.ps1, the GDAP
    customer via -DelegatedOrganization); app-only with -ClientId +
    -CertificateThumbprint or -AppOnly. Under Windows PowerShell 5.1, which cannot
    load that helper, the script connects delegated on its own with the same rules.

    Why Exchange Online and not Graph: Graph's audit log query API
    (/security/auditLog/queries) is asynchronous - a query is created, runs for
    minutes, and its records are fetched afterwards - and returns a different
    record shape. The slicing, splitting and retry logic below is built and tested
    on Search-UnifiedAuditLog, so this script keeps it.

    The audit log runs 30-90 minutes (occasionally 24 hours) behind; a run on
    the same day can miss the most recent actions.

    Only what happened INSIDE the window can be followed. A folder that was
    renamed before -StartDate is invisible, so the path at the start of the
    window is taken from the first record found. If the trail seems to start
    in the middle, widen the window.

    Actions by the OneDrive sync client (a rename in Explorer) are audited just
    like actions in the browser; the UserAgent column tells them apart.

    The last known location is what the audit log says; the script does not
    check that the file is still there. Status "In recycle bin" means
    Restore-RecycleBinItems.ps1 can bring it back.
#>
[CmdletBinding(DefaultParameterSetName = 'Name')]
param(
    [Parameter(Mandatory, ParameterSetName = 'Name', Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]   $Name,

    [Parameter(Mandatory, ParameterSetName = 'Url')]
    [ValidateNotNullOrEmpty()]
    [string]   $Url,

    [Parameter(Mandatory, ParameterSetName = 'ItemId')]
    [guid]     $ItemId,

    [string]   $SiteUrl,
    [object]   $StartDate,
    [object]   $EndDate,
    [ValidateRange(1, 3650)]
    [int]      $Days = 30,
    [string]   $TimeZone = 'Europe/Brussels',
    [switch]   $FollowCopies,
    [switch]   $IncludeActivity,
    [string]   $OutputPath,
    [string]   $TenantId,
    [string]   $ClientId,
    [string]   $CertificateThumbprint,
    [switch]   $AppOnly,
    [switch]   $PassThru
)

$ErrorActionPreference = 'Stop'

# -- Time zone -----------------------------------------------------------------
function Resolve-TimeZone([string]$Id) {
    $candidates = [System.Collections.Generic.List[string]]::new()
    $candidates.Add($Id)
    # .NET 6+ can translate IANA <-> Windows itself; Windows PowerShell 5.1 cannot.
    $out = $null
    $m = [TimeZoneInfo].GetMethod('TryConvertIanaIdToWindowsId', [type[]]@([string], [string].MakeByRefType()))
    if ($m) { $args2 = @($Id, $null); if ($m.Invoke($null, $args2)) { $candidates.Add($args2[1]) } }
    $known = @{ 'europe/brussels' = 'Romance Standard Time'; 'europe/amsterdam' = 'W. Europe Standard Time'
                'europe/paris' = 'Romance Standard Time'; 'europe/luxembourg' = 'W. Europe Standard Time'
                'utc' = 'UTC'; 'europe/london' = 'GMT Standard Time' }
    if ($known.ContainsKey($Id.ToLowerInvariant())) { $candidates.Add($known[$Id.ToLowerInvariant()]) }
    foreach ($c in $candidates) {
        try { return [TimeZoneInfo]::FindSystemTimeZoneById($c) } catch { $out = $_ }
    }
    throw "Time zone '$Id' is not known on this machine. Use e.g. 'Europe/Brussels' or 'Romance Standard Time'."
}
$tz = Resolve-TimeZone $TimeZone

# Wall-clock time in $tz -> UTC. A time in the spring-forward gap does not exist; move it an hour on.
function ConvertTo-UtcFromZone([datetime]$Wall) {
    $w = [datetime]::SpecifyKind($Wall, [DateTimeKind]::Unspecified)
    if ($tz.IsInvalidTime($w)) { $w = $w.AddHours(1) }
    [TimeZoneInfo]::ConvertTimeToUtc($w, $tz)
}
function Format-ZoneTime([datetime]$Utc) {
    $local  = [TimeZoneInfo]::ConvertTimeFromUtc($Utc, $tz)
    $offset = $tz.GetUtcOffset($Utc)
    $sign   = if ($offset -lt [timespan]::Zero) { '-' } else { '+' }
    '{0} {1}{2:hh\:mm}' -f $local.ToString('yyyy-MM-dd HH:mm:ss'), $sign, $offset.Duration()
}

# Accepts a [datetime] or a string (day-first or ISO). Returns the moment in UTC.
function Resolve-InputDate([object]$Value, [string]$ParamName, [switch]$IsEnd) {
    if ($Value -is [datetime]) {
        if ($Value.Kind -eq [DateTimeKind]::Unspecified) { return (ConvertTo-UtcFromZone $Value) }
        return $Value.ToUniversalTime()
    }
    $s = "$Value".Trim()
    $formats = @(
        'yyyy-MM-dd', 'yyyy-MM-dd HH:mm', 'yyyy-MM-dd HH:mm:ss', "yyyy-MM-dd'T'HH:mm", "yyyy-MM-dd'T'HH:mm:ss",
        'dd-MM-yyyy', 'dd-MM-yyyy HH:mm', 'dd-MM-yyyy HH:mm:ss', 'd-M-yyyy', 'd-M-yyyy H:mm',
        'dd/MM/yyyy', 'dd/MM/yyyy HH:mm', 'dd/MM/yyyy HH:mm:ss', 'd/M/yyyy', 'd/M/yyyy H:mm',
        'dd.MM.yyyy', 'dd.MM.yyyy HH:mm'
    )
    $parsed = [datetime]::MinValue
    if (-not [datetime]::TryParseExact($s, [string[]]$formats, [Globalization.CultureInfo]::InvariantCulture,
                                       [Globalization.DateTimeStyles]::None, [ref]$parsed)) {
        if (-not [datetime]::TryParse($s, [Globalization.CultureInfo]::GetCultureInfo('nl-BE'),
                                      [Globalization.DateTimeStyles]::None, [ref]$parsed)) {
            throw "-$ParamName '$s' is not a date. Use e.g. '15-09-2026', '15/09/2026 08:30' or '2026-09-15 08:30'."
        }
    }
    # A bare date as end date means "up to and including that day".
    if ($IsEnd -and $parsed.TimeOfDay -eq [timespan]::Zero -and $s -notmatch '\d:\d') { $parsed = $parsed.AddDays(1) }
    ConvertTo-UtcFromZone $parsed
}

$nowUtc   = [datetime]::UtcNow
$endUtc   = if ($PSBoundParameters.ContainsKey('EndDate'))   { Resolve-InputDate $EndDate 'EndDate' -IsEnd } else { $nowUtc }
$startUtc = if ($PSBoundParameters.ContainsKey('StartDate')) { Resolve-InputDate $StartDate 'StartDate' } else { $endUtc.AddDays(-$Days) }
if ($endUtc -gt $nowUtc) { $endUtc = $nowUtc }
if ($startUtc -ge $endUtc) { throw "-StartDate must be earlier than -EndDate (and not in the future)." }

# -- Paths ---------------------------------------------------------------------
function ConvertTo-PathKey([string]$Path) {
    if (-not $Path) { return '' }
    $p = $Path.Trim()
    $p = ($p -split '\?')[0] -replace '#.*$', ''
    try { $p = [uri]::UnescapeDataString($p) } catch { }
    $p = $p -replace '/:[a-z]:/r/', '/'          # https://t.sharepoint.com/:w:/r/sites/x/... -> /sites/x/...
    $p.TrimEnd('/').ToLowerInvariant()
}
function Join-SpPath([string]$Site, [string]$Folder, [string]$Leaf) {
    $base = if ($Folder -match '^https?://') { $Folder } elseif ($Site) { $Site.TrimEnd('/') + '/' + $Folder.Trim('/') } else { $Folder }
    $base = $base.TrimEnd('/')
    if ($Leaf) { $base = "$base/$Leaf" }
    try { [uri]::UnescapeDataString($base) } catch { $base }
}
function Split-Leaf([string]$Path) {
    $p = $Path.TrimEnd('/')
    $i = $p.LastIndexOf('/')
    if ($i -lt 0) { return @('', $p) }
    @($p.Substring(0, $i), $p.Substring($i + 1))
}

$siteKey = if ($SiteUrl) { ConvertTo-PathKey $SiteUrl } else { '' }
$sitePre = if ($siteKey) { ($siteKey -replace '^https?://', '') } else { '' }
function Test-UnderSite([string]$Key) { $siteKey -and ($Key -eq $siteKey -or $Key.StartsWith($siteKey + '/')) }
$mode = $PSCmdlet.ParameterSetName
if ($Url -and $Url -match '/:[a-z]:/[sg]/') {
    throw "This is an opaque sharing link (/:x:/s/ or /:x:/g/) - it does not contain the file's path. Open it in a browser and copy the address it lands on, or search by -Name."
}

# -- Operations ----------------------------------------------------------------
$fileOps = @('FileRenamed', 'FileMoved', 'FileCopied', 'FileDeleted', 'FileRecycled',
             'FileDeletedFirstStageRecycleBin', 'FileDeletedSecondStageRecycleBin',
             'FileRestored', 'FileUploaded', 'FileVersionsAllDeleted')
$folderOps = @('FolderRenamed', 'FolderMoved', 'FolderDeleted', 'FolderRecycled',
               'FolderDeletedFirstStageRecycleBin', 'FolderDeletedSecondStageRecycleBin', 'FolderRestored')
$activityOps = @('FileAccessed', 'FileModified', 'FileDownloaded', 'FileSyncDownloadedFull',
                 'FileSyncUploadedFull', 'FileCheckedIn', 'FileCheckedOut')
$ops = $fileOps + $folderOps
if ($IncludeActivity) { $ops += $activityOps }

# -- Output --------------------------------------------------------------------
$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not $OutputPath) {
    $label = switch ($mode) {
        'Name'   { $Name }
        'Url'    { (Split-Leaf (ConvertTo-PathKey $Url))[1] }
        'ItemId' { "$ItemId" }
    }
    $label = ($label -replace '[\\/:*?"<>|\s]+', '_').Trim('_')
    if (-not $label) { $label = 'file' }
    if ($label.Length -gt 60) { $label = $label.Substring(0, 60) }
    $OutputPath = Join-Path $outputDir ("FileTrail_{0}_{1}.csv" -f $label, (Get-Date -Format 'yyyyMMdd_HHmmss'))
}
$outDir = Split-Path -Parent $OutputPath
if ($outDir -and -not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
$rawPath = [IO.Path]::ChangeExtension($OutputPath, '.json')

# -- Connection ----------------------------------------------------------------
if (-not (Get-Command Search-UnifiedAuditLog -ErrorAction SilentlyContinue)) {
    if (-not (Get-Module -ListAvailable ExchangeOnlineManagement)) {
        throw "Module ExchangeOnlineManagement is not installed: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
    }
}
$connectedHere = $false
$exoConnection = $null
if ($PSVersionTable.PSVersion.Major -ge 7) {
    # Delegated by default (device code / GDAP customer per load.config.ps1), app-only on
    # request; reuses a session to the right organisation and only closes what it opened.
    . (Join-Path $PSScriptRoot '..\Startup\Connect-M365.ps1')
    $exoArgs = @{}
    if ($TenantId)              { $exoArgs['TenantId'] = $TenantId }
    if ($ClientId)              { $exoArgs['ClientId'] = $ClientId }
    if ($CertificateThumbprint) { $exoArgs['CertificateThumbprint'] = $CertificateThumbprint }
    if ($AppOnly)               { $exoArgs['AppOnly'] = $true }
    $exoConnection = Connect-M365Exchange @exoArgs
} else {
    # Windows PowerShell 5.1 cannot load Connect-M365.ps1 (PowerShell 7): the same
    # rules, inline.
    $connected = $false
    try {
        if (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue) {
            $conn = @(Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object { $_.State -eq 'Connected' })
            if ($TenantId) { $conn = @($conn | Where-Object { $_.TenantID -eq $TenantId -or $_.DelegatedOrganization -eq $TenantId -or $_.Organization -eq $TenantId -or $_.UserPrincipalName -like "*@$TenantId" }) }
            $connected = $conn.Count -gt 0
        }
    } catch { $connected = $false }
    if (-not $connected) {
        $cp = @{ ShowBanner = $false; ErrorAction = 'Stop' }
        if ($ClientId) {
            if (-not $TenantId -or -not $CertificateThumbprint) { throw 'App-only needs -TenantId (domain) and -CertificateThumbprint.' }
            $cp['AppId'] = $ClientId; $cp['CertificateThumbprint'] = $CertificateThumbprint; $cp['Organization'] = $TenantId
        } else {
            $org = if ($TenantId) { $TenantId } elseif ("$global:authMode" -eq 'GDAP' -and $global:connectmsoldomain) { [string]$global:connectmsoldomain } else { $null }
            # -Organization only applies to app-only sign-in; a partner reaches a customer with -DelegatedOrganization.
            if ($org) { $cp['DelegatedOrganization'] = $org }
            if ($global:useDeviceCodeAuth) { $cp['Device'] = $true } elseif ($global:upn) { $cp['UserPrincipalName'] = [string]$global:upn }
        }
        Connect-ExchangeOnline @cp
        $connectedHere = $true
    }
}

try {
    # -- Header ----------------------------------------------------------------
    $target = switch ($mode) { 'Name' { "name   $Name" } 'Url' { "url    $Url" } 'ItemId' { "itemid $ItemId" } }
    Write-Host ""
    Write-Host "  ================================================" -ForegroundColor Cyan
    Write-Host "   Trace a OneDrive / SharePoint file" -ForegroundColor Cyan
    Write-Host "  ================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Looking for : $target" -ForegroundColor DarkGray
    Write-Host "  Window      : $(Format-ZoneTime $startUtc)  ->  $(Format-ZoneTime $endUtc)  ($TimeZone)" -ForegroundColor DarkGray
    if ($SiteUrl) { Write-Host "  Site        : $SiteUrl" -ForegroundColor DarkGray }
    Write-Host "  Operations  : $($ops.Count) ($(if ($IncludeActivity) { 'incl. activity' } else { 'renames, moves, copies, deletes, restores, uploads' }))" -ForegroundColor DarkGray
    Write-Host ""

    if (($nowUtc - $startUtc).TotalDays -gt 180) {
        Write-Warning "The window starts more than 180 days ago. Audit Standard keeps 180 days; older records only exist with Audit Premium or a longer retention policy."
    }
    try {
        $cfg = Get-AdminAuditLogConfig -ErrorAction Stop
        if ($cfg.PSObject.Properties['UnifiedAuditLogIngestionEnabled'] -and -not $cfg.UnifiedAuditLogIngestionEnabled) {
            Write-Warning "UnifiedAuditLogIngestionEnabled is False in this tenant: there may be no records at all. Turn it on for the future with Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled `$true."
        }
    } catch { Write-Verbose "Could not read the audit log configuration: $($_.Exception.Message)" }

    # -- 1. Read the audit log -------------------------------------------------
    $script:truncated = $false
    $script:incomplete = $false
    $minSlice = [timespan]::FromMinutes(15)

    # Reads one slice. Returns @{ Split = $true } when it holds more than one search can return.
    function Read-AuditSlice([datetime]$From, [datetime]$To) {
        $fromS = $From.ToString('yyyy-MM-dd HH:mm:ss') + 'Z'
        $toS   = $To.ToString('yyyy-MM-dd HH:mm:ss') + 'Z'
        $best = [System.Collections.Generic.List[object]]::new()
        for ($attempt = 1; $attempt -le 4; $attempt++) {
            $sid   = "FileTrail_$([guid]::NewGuid())"
            $buf   = [System.Collections.Generic.List[object]]::new()
            $total = 0; $bad = $false
            try {
                for ($page = 1; $page -le 12; $page++) {
                    $batch = @(Search-UnifiedAuditLog -StartDate $fromS -EndDate $toS -Operations $ops `
                                 -SessionId $sid -SessionCommand ReturnLargeSet -ResultSize 5000 -ErrorAction Stop)
                    if ($batch.Count -eq 0) { break }
                    # ResultIndex -1 / ResultCount 0 on a non-empty page is the service's way of saying "this went wrong".
                    if (@($batch | Where-Object { $_.ResultIndex -eq -1 }).Count -gt 0 -or [int]$batch[0].ResultCount -eq 0) { $bad = $true; break }
                    if ($total -eq 0) { $total = [int]$batch[0].ResultCount }
                    if ($total -gt 50000 -and ($To - $From) -gt $minSlice) { return @{ Split = $true } }
                    $buf.AddRange([object[]]$batch)
                    if ($buf.Count -ge $total) { break }
                }
            } catch {
                $msg = $_.Exception.Message
                if ($msg -match 'not (be )?enabled|UnifiedAuditLogIngestionEnabled') { throw "The Unified Audit Log is not enabled in this tenant: $msg" }
                if ($msg -match 'not recognized|is not recognized|Access.*denied|not authorized|RoleAssignment') { throw "Search-UnifiedAuditLog is not available to this account (needs View-Only Audit Logs / Audit Logs): $msg" }
                Write-Verbose "Attempt $attempt for $fromS - $toS failed: $msg"
                $bad = $true
            }
            if (-not $bad -and ($total -eq 0 -or $buf.Count -ge [math]::Min($total, 50000))) {
                if ($total -gt 50000) { $script:truncated = $true }
                return @{ Split = $false; Records = $buf }
            }
            if ($buf.Count -gt $best.Count) { $best = $buf }
            if ($attempt -lt 4) { Start-Sleep -Seconds (5 * $attempt * $attempt) }
        }
        $script:incomplete = $true
        Write-Warning "Audit search $fromS - $toS stayed incomplete after 4 attempts; using the $($best.Count) record(s) it did return."
        @{ Split = $false; Records = $best }
    }

    $seen    = [System.Collections.Generic.HashSet[string]]::new()
    $records = [System.Collections.Generic.List[object]]::new()
    $rawById = @{}
    $queue   = [System.Collections.Generic.Queue[object]]::new()
    $cursor  = $startUtc
    while ($cursor -lt $endUtc) {
        $next = $cursor.AddDays(1); if ($next -gt $endUtc) { $next = $endUtc }
        $queue.Enqueue(@($cursor, $next)); $cursor = $next
    }
    $totalSpan = ($endUtc - $startUtc).TotalSeconds
    $doneSpan  = 0.0
    $rawCount  = 0
    $sw = [Diagnostics.Stopwatch]::StartNew()

    while ($queue.Count -gt 0) {
        $slice = $queue.Dequeue()
        $pct = [int][math]::Min(100, 100 * $doneSpan / $totalSpan)
        Write-Progress -Activity 'Reading the audit log' -Status ("{0}  ->  {1}   ({2} records kept, {3:mm\:ss} elapsed)" -f (Format-ZoneTime $slice[0]), (Format-ZoneTime $slice[1]), $records.Count, $sw.Elapsed) -PercentComplete $pct
        $r = Read-AuditSlice $slice[0] $slice[1]
        if ($r.Split) {
            $mid = $slice[0].AddTicks([long](($slice[1] - $slice[0]).Ticks / 2))
            Write-Verbose "Splitting $(Format-ZoneTime $slice[0]) - $(Format-ZoneTime $slice[1]): more than 50,000 records."
            # Keep chronological order: second half behind the first.
            $rest = @($queue.ToArray()); $queue.Clear()
            $queue.Enqueue(@($slice[0], $mid)); $queue.Enqueue(@($mid, $slice[1]))
            foreach ($q in $rest) { $queue.Enqueue($q) }
            continue
        }
        $doneSpan += ($slice[1] - $slice[0]).TotalSeconds

        foreach ($rec in $r.Records) {
            $rawCount++
            $json = "$($rec.AuditData)"
            # Cheap pre-filter before parsing: the site's host/path must appear somewhere in the record.
            if ($sitePre -and $json.Replace('\/', '/').ToLowerInvariant().IndexOf($sitePre) -lt 0) { continue }
            try { $d = $json | ConvertFrom-Json } catch { continue }
            $id = if ($d.Id) { "$($d.Id)" } else { "$($rec.Identity)" }
            if (-not $seen.Add($id)) { continue }

            $t = $d.CreationTime
            if ($t -is [datetime]) {
                $t = if ($t.Kind -eq [DateTimeKind]::Local) { $t.ToUniversalTime() } else { [datetime]::SpecifyKind($t, [DateTimeKind]::Utc) }
            } else {
                $t = [datetime]::Parse("$t", [Globalization.CultureInfo]::InvariantCulture,
                        [Globalization.DateTimeStyles]::AssumeUniversal -bor [Globalization.DateTimeStyles]::AdjustToUniversal)
            }
            if ($t -lt $startUtc -or $t -ge $endUtc) { continue }

            $site = "$($d.SiteUrl)"
            $op   = "$($d.Operation)"
            $isFolder = ("$($d.ItemType)" -eq 'Folder') -or $op -like 'Folder*'

            $srcFolder = "$($d.SourceRelativeUrl)"; $srcName = "$($d.SourceFileName)"
            $objPath = if ($d.ObjectId) { try { [uri]::UnescapeDataString("$($d.ObjectId)") } catch { "$($d.ObjectId)" } } else { '' }
            if (-not $srcName -and $objPath) { $srcName = (Split-Leaf $objPath)[1] }
            $srcPath = if ($srcFolder) { Join-SpPath $site $srcFolder $srcName } else { $objPath }
            # Some records put the item itself in SourceRelativeUrl; then ObjectId is the right path.
            if ($srcFolder -and $objPath -and (ConvertTo-PathKey (Join-SpPath $site $srcFolder '')) -eq (ConvertTo-PathKey $objPath)) { $srcPath = $objPath }
            if (-not $srcPath) { $srcPath = Join-SpPath $site '' $srcName }

            $dstPath = ''
            if ($op -in 'FileRenamed', 'FileMoved', 'FileCopied', 'FolderRenamed', 'FolderMoved') {
                $dstFolder = "$($d.DestinationRelativeUrl)"; $dstName = "$($d.DestinationFileName)"
                if (-not $dstName)   { $dstName = $srcName }
                if (-not $dstFolder) { $dstPath = Join-SpPath '' (Split-Leaf $srcPath)[0] $dstName }
                else                 { $dstPath = Join-SpPath $site $dstFolder $dstName }
            }
            if ($siteKey -and -not ((Test-UnderSite (ConvertTo-PathKey $srcPath)) -or (Test-UnderSite (ConvertTo-PathKey $site)))) { continue }

            $records.Add([pscustomobject]@{
                Id        = $id
                TimeUtc   = $t
                Operation = $op
                IsFolder  = $isFolder
                User      = "$($d.UserId)"
                Site      = $site
                SrcPath   = $srcPath
                DstPath   = $dstPath
                SrcKey    = ConvertTo-PathKey $srcPath
                DstKey    = ConvertTo-PathKey $dstPath
                ItemId    = "$($d.ListItemUniqueId)".ToLowerInvariant()
                ClientIP  = "$($d.ClientIP)"
                UserAgent = "$($d.UserAgent)"
            })
            $rawById[$id] = $json
        }
    }
    Write-Progress -Activity 'Reading the audit log' -Completed
    Write-Host ("  Read {0:N0} audit record(s) in {1:mm\:ss}; {2:N0} file/folder record(s) kept." -f $rawCount, $sw.Elapsed, $records.Count) -ForegroundColor DarkGray

    # -- 2. Seed ---------------------------------------------------------------
    $fileRecs   = @($records | Where-Object { -not $_.IsFolder } | Sort-Object TimeUtc)
    $folderRecs = @($records | Where-Object { $_.IsFolder -and $_.Operation -in $folderOps } | Sort-Object TimeUtc)

    $wild = $false; $pattern = $null; $hasExt = $false
    if ($Name) {
        $wild    = $Name -match '[\*\?]'
        $pattern = $Name -replace '\[', '`[' -replace '\]', '`]'
        $hasExt  = [IO.Path]::HasExtension($Name) -and -not $wild
    }
    $urlKey = if ($Url) { ConvertTo-PathKey $Url } else { '' }
    $idKey  = if ($PSBoundParameters.ContainsKey('ItemId')) { "$ItemId".ToLowerInvariant() } else { '' }

    function Test-Seed($r) {
        switch ($mode) {
            'ItemId' { return $r.ItemId -eq $idKey }
            'Url'    { return ($r.SrcKey -eq $urlKey -or $r.DstKey -eq $urlKey) }
            'Name'   {
                foreach ($p in @($r.SrcPath, $r.DstPath)) {
                    if (-not $p) { continue }
                    $leaf = (Split-Leaf $p)[1]
                    if ($wild) { if ($leaf -like $pattern) { return $true } }
                    elseif ($leaf -eq $Name) { return $true }
                    elseif (-not $hasExt -and [IO.Path]::GetFileNameWithoutExtension($leaf) -eq $Name) { return $true }
                }
                return $false
            }
        }
    }

    # -- 3. Follow -------------------------------------------------------------
    $inTrail   = [System.Collections.Generic.HashSet[string]]::new()
    $knownIds  = [System.Collections.Generic.HashSet[string]]::new()
    $knownPath = @{}   # path key -> earliest UTC time the file was at that path

    function Add-ToTrail($r) {
        if (-not $inTrail.Add($r.Id)) { return $false }
        $isCopy = $r.Operation -eq 'FileCopied'
        if ($r.ItemId) { [void]$knownIds.Add($r.ItemId) }
        # A copy's source is the original, which stays; only follow the copy itself with -FollowCopies.
        if ($r.DstKey -and (-not $isCopy -or $FollowCopies)) {
            if (-not $knownPath.ContainsKey($r.DstKey) -or $knownPath[$r.DstKey] -gt $r.TimeUtc) { $knownPath[$r.DstKey] = $r.TimeUtc }
        }
        $true
    }
    foreach ($r in $fileRecs) { if (Test-Seed $r) { [void](Add-ToTrail $r) } }
    if ($mode -eq 'ItemId') { [void]$knownIds.Add($idKey) }
    if ($mode -eq 'Url')    { $knownPath[$urlKey] = [datetime]::MinValue }

    do {
        $added = 0
        foreach ($r in $fileRecs) {
            if ($inTrail.Contains($r.Id)) { continue }
            $hit = ($r.ItemId -and $knownIds.Contains($r.ItemId)) -or
                   ($r.SrcKey -and $knownPath.ContainsKey($r.SrcKey) -and $r.TimeUtc -ge $knownPath[$r.SrcKey])
            if ($hit -and (Add-ToTrail $r)) { $added++ }
        }
    } while ($added -gt 0)

    $trail = @($fileRecs | Where-Object { $inTrail.Contains($_.Id) })

    if ($trail.Count -eq 0) {
        Write-Host ""
        Write-Host "  Nothing found for $target in this window." -ForegroundColor Yellow
        Write-Host "   - Widen the window (-Days / -StartDate); Audit Standard goes back 180 days." -ForegroundColor DarkGray
        Write-Host "   - Try part of the name with wildcards: -Name '*offerte*'." -ForegroundColor DarkGray
        Write-Host "   - The audit log runs 30-90 minutes behind; very recent actions may not be in it yet." -ForegroundColor DarkGray
        Write-Host "   - If only the folder was touched, the file itself has no record: search for the folder name." -ForegroundColor DarkGray
        if ($script:incomplete -or $script:truncated) { Write-Host "   - The read was incomplete (see warnings above); a rerun may find more." -ForegroundColor DarkGray }
        Write-Host ""
        return
    }

    # Group into items: records sharing an item ID or a path the item moved through belong together.
    $parent = @{}
    function Find-Root([string]$k) { while ($parent[$k] -ne $k) { $parent[$k] = $parent[$parent[$k]]; $k = $parent[$k] }; $k }
    function Join-Set([string]$a, [string]$b) {
        foreach ($k in $a, $b) { if (-not $parent.ContainsKey($k)) { $parent[$k] = $k } }
        $ra = Find-Root $a; $rb = Find-Root $b; if ($ra -ne $rb) { $parent[$rb] = $ra }
    }
    foreach ($r in $trail) {
        $self = "r:$($r.Id)"
        Join-Set $self $self
        if ($r.ItemId) { Join-Set $self "i:$($r.ItemId)" }
        if ($r.SrcKey) { Join-Set $self "p:$($r.SrcKey)" }
        if ($r.DstKey -and ($r.Operation -ne 'FileCopied')) { Join-Set $self "p:$($r.DstKey)" }
    }
    $groups = $trail | Group-Object { Find-Root "r:$($_.Id)" } | Sort-Object { ($_.Group | Sort-Object TimeUtc | Select-Object -First 1).TimeUtc }

    # -- 4. Replay folder actions and build the timeline ----------------------
    $describe = @{
        FileRenamed = 'Renamed'; FileMoved = 'Moved'; FileCopied = 'Copied'; FileDeleted = 'Deleted (to recycle bin)'
        FileRecycled = 'Deleted (to recycle bin)'; FileDeletedFirstStageRecycleBin = 'Removed from recycle bin (now in 2nd stage)'
        FileDeletedSecondStageRecycleBin = 'Permanently deleted'; FileRestored = 'Restored'; FileUploaded = 'Uploaded'
        FileVersionsAllDeleted = 'All versions deleted'; FileAccessed = 'Opened'; FileModified = 'Modified'
        FileDownloaded = 'Downloaded'; FileSyncDownloadedFull = 'Synced down'; FileSyncUploadedFull = 'Synced up'
        FileCheckedIn = 'Checked in'; FileCheckedOut = 'Checked out'
        FolderRenamed = 'Folder renamed'; FolderMoved = 'Folder moved'; FolderDeleted = 'Folder deleted (to recycle bin)'
        FolderRecycled = 'Folder deleted (to recycle bin)'; FolderDeletedFirstStageRecycleBin = 'Folder removed from recycle bin (2nd stage)'
        FolderDeletedSecondStageRecycleBin = 'Folder permanently deleted'; FolderRestored = 'Folder restored'
    }

    $rows = [System.Collections.Generic.List[object]]::new()
    $summaries = [System.Collections.Generic.List[object]]::new()
    $itemNo = 0

    foreach ($g in $groups) {
        $itemNo++
        $events = @($g.Group | Sort-Object TimeUtc)
        $first  = $events[0]
        $current = $first.SrcPath
        $status  = 'Present'
        $statusNote = ''
        $timeline = [System.Collections.Generic.List[object]]::new()

        # Merge file and folder events chronologically.
        $merged = @($events | ForEach-Object { [pscustomobject]@{ Kind = 'File'; R = $_ } }) +
                  @($folderRecs | Where-Object { $_.TimeUtc -ge $first.TimeUtc } | ForEach-Object { [pscustomobject]@{ Kind = 'Folder'; R = $_ } })
        $merged = @($merged | Sort-Object { $_.R.TimeUtc }, { if ($_.Kind -eq 'Folder') { 0 } else { 1 } })

        foreach ($m in $merged) {
            $r = $m.R
            $from = ''; $to = ''; $via = ''
            if ($m.Kind -eq 'Folder') {
                $curKey = ConvertTo-PathKey $current
                if (-not $r.SrcKey -or -not $curKey.StartsWith($r.SrcKey + '/')) { continue }
                $withFolder = "together with folder $($r.SrcPath)"
                # A file already in the recycle bin on its own does not move with its old folder;
                # one deleted with the folder only follows that folder's own recycle bin steps.
                $applies = if ($r.Operation -in 'FolderRenamed', 'FolderMoved', 'FolderDeleted', 'FolderRecycled') { $status -eq 'Present' }
                           else { $statusNote -eq $withFolder }
                if (-not $applies) { continue }
                $via  = $r.SrcPath
                $from = $current
                switch ($r.Operation) {
                    { $_ -in 'FolderRenamed', 'FolderMoved' } {
                        if ($r.DstPath) { $current = $r.DstPath + $current.Substring($r.SrcPath.TrimEnd('/').Length) }
                        $to = $current
                    }
                    'FolderRestored'                             { $status = 'Present'; $statusNote = '' }
                    { $_ -in 'FolderDeleted', 'FolderRecycled' } { $status = 'In recycle bin'; $statusNote = $withFolder }
                    'FolderDeletedFirstStageRecycleBin'          { $status = 'In second-stage recycle bin' }
                    'FolderDeletedSecondStageRecycleBin'         { $status = 'Permanently deleted' }
                }
            } else {
                $from = $r.SrcPath
                switch ($r.Operation) {
                    { $_ -in 'FileRenamed', 'FileMoved' } { $to = $r.DstPath; if ($r.DstPath) { $current = $r.DstPath }; $status = 'Present'; $statusNote = '' }
                    'FileCopied' { $to = $r.DstPath }
                    { $_ -in 'FileDeleted', 'FileRecycled' } { $current = $r.SrcPath; $status = 'In recycle bin'; $statusNote = '' }
                    'FileDeletedFirstStageRecycleBin'  { $status = 'In second-stage recycle bin'; $statusNote = '' }
                    'FileDeletedSecondStageRecycleBin' { $status = 'Permanently deleted'; $statusNote = '' }
                    { $_ -in 'FileRestored', 'FileUploaded' } { $current = $r.SrcPath; $status = 'Present'; $statusNote = '' }
                    default { if ($status -eq 'Present' -and $r.SrcPath) { $current = $r.SrcPath } }
                }
            }
            $row = [pscustomobject]@{
                Item       = $itemNo
                Time       = Format-ZoneTime $r.TimeUtc
                TimeUtc    = $r.TimeUtc.ToString('yyyy-MM-dd HH:mm:ss') + 'Z'
                Action     = if ($describe.ContainsKey($r.Operation)) { $describe[$r.Operation] } else { $r.Operation }
                Operation  = $r.Operation
                User       = $r.User
                From       = $from
                To         = $to
                ViaFolder  = $via
                ItemId     = $r.ItemId
                ClientIP   = $r.ClientIP
                UserAgent  = $r.UserAgent
                RecordId   = $r.Id
            }
            $timeline.Add($row); $rows.Add($row)
        }

        $names = @($events | ForEach-Object { (Split-Leaf $_.SrcPath)[1]; if ($_.DstPath) { (Split-Leaf $_.DstPath)[1] } } | Where-Object { $_ } | Select-Object -Unique)
        $summaries.Add([pscustomobject]@{
            Item = $itemNo; Names = $names; LastKnown = $current; Status = $status; Note = $statusNote
            Timeline = $timeline; ItemIds = @($events.ItemId | Where-Object { $_ } | Select-Object -Unique)
        })
    }

    # -- Display ---------------------------------------------------------------
    foreach ($s in $summaries) {
        Write-Host ""
        Write-Host ("  -- Item {0}: {1}" -f $s.Item, ($s.Names -join '  |  ')) -ForegroundColor Cyan
        if ($s.ItemIds) { Write-Host "     ItemId: $($s.ItemIds -join ', ')" -ForegroundColor DarkGray }
        foreach ($row in $s.Timeline) {
            $color = if     ($row.Operation -match 'Deleted|Recycled') { 'Red' }
                     elseif ($row.Operation -match 'Renamed|Moved')    { 'Yellow' }
                     elseif ($row.Operation -match 'Restored')         { 'Green' }
                     elseif ($row.Operation -eq 'FileCopied')          { 'Magenta' }
                     else                                              { 'Gray' }
            Write-Host ("     {0}  {1,-32} {2}" -f $row.Time, $row.Action, $row.User) -ForegroundColor $color
            if ($row.To -and $row.To -ne $row.From) {
                Write-Host "         from $($row.From)" -ForegroundColor DarkGray
                Write-Host "         to   $($row.To)" -ForegroundColor DarkGray
            } elseif ($row.From) {
                Write-Host "         at   $($row.From)" -ForegroundColor DarkGray
            }
            if ($row.ViaFolder) { Write-Host "         (the file itself has no record: its folder $($row.ViaFolder) was changed)" -ForegroundColor DarkGray }
        }
        $sc = switch ($s.Status) { 'Present' { 'Green' } 'Permanently deleted' { 'Red' } default { 'Yellow' } }
        Write-Host ""
        Write-Host "     Last known location : $($s.LastKnown)" -ForegroundColor White
        Write-Host ("     Status              : {0}{1}" -f $s.Status, $(if ($s.Note) { " ($($s.Note))" } else { '' })) -ForegroundColor $sc
        if ($s.Status -like '*recycle bin*') {
            Write-Host "     Restore it with Restore-RecycleBinItems.ps1 (same folder) against the site above." -ForegroundColor DarkGray
        }
    }

    # -- Export ----------------------------------------------------------------
    $rows | Select-Object Item, Time, TimeUtc, Action, Operation, User, From, To, ViaFolder, ItemId, ClientIP, UserAgent, RecordId |
        Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    $rawJson = '[' + ((@($rows | Select-Object -ExpandProperty RecordId -Unique) | ForEach-Object { $rawById[$_] }) -join ",`n") + ']'
    [IO.File]::WriteAllText($rawPath, $rawJson, [Text.UTF8Encoding]::new($false))

    Write-Host ""
    if ($script:truncated)  { Write-Warning "At least one 15-minute slice held more than 50,000 records; records beyond that could not be read. Narrow with -SiteUrl." }
    if ($script:incomplete) { Write-Warning "At least one audit search stayed incomplete after retries; the trail may miss steps. Rerun later." }
    Write-Host "  $($summaries.Count) item(s), $($rows.Count) event(s). Times in $TimeZone." -ForegroundColor Cyan
    Write-Host "  Report : $OutputPath" -ForegroundColor Green
    Write-Host "  Raw    : $rawPath" -ForegroundColor Green
    Write-Host ""

    if ($PassThru) { $rows }
}
finally {
    if ($exoConnection) { Disconnect-M365Exchange $exoConnection }
    elseif ($connectedHere) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null }
}
