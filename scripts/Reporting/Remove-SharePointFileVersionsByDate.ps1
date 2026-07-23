#Requires -Version 5.1
<#
.SYNOPSIS
    Remove SharePoint Online file versions older than a cutoff date while preserving the current version.

.DESCRIPTION
    Connects to SharePoint Online using PnP.PowerShell and scans either a single site or all
    SharePoint sites in the tenant. For each document library, all files are enumerated and their
    previous versions are inspected. Versions older than -BeforeDate are reported by default.

    Only when -Apply is specified are matching versions actually removed. The current version of a
    file is always preserved because Get-PnPFileVersion only returns previous versions.

    Default behavior is safe preview mode.

.PARAMETER BeforeDate
    Delete versions older than this date/time.

.PARAMETER SiteUrl
    Optional. Scan a single SharePoint site.

.PARAMETER TenantUrl
    Optional. Tenant root URL, for example https://contoso.sharepoint.com.
    Required when scanning all sites.

.PARAMETER ClientId
    Entra ID app/client ID for interactive PnP login. If omitted, the script tries
    ENTRAID_APP_ID first and then AZURE_CLIENT_ID.

.PARAMETER OutputPath
    Override the default output folder.

.PARAMETER Apply
    Actually remove matching versions. Without this switch, the script only reports.

.PARAMETER IncludeOneDriveSites
    Include personal OneDrive sites in an all-sites scan.

.PARAMETER IncludeHiddenLibraries
    Also inspect hidden document libraries. Disabled by default for safety.

.PARAMETER LibraryTitle
    Optional filter. Limit the scan to one or more library titles.

.EXAMPLE
    .\Remove-SharePointFileVersionsByDate.ps1 -TenantUrl "https://contoso.sharepoint.com" -BeforeDate "2025-01-01"

.EXAMPLE
    .\Remove-SharePointFileVersionsByDate.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/Finance" -BeforeDate "2025-01-01" -Apply
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [datetime] $BeforeDate,
    [string] $SiteUrl,
    [string] $TenantUrl,
    [string] $ClientId,
    [string] $OutputPath,
    [switch] $Apply,
    [switch] $IncludeOneDriveSites,
    [switch] $IncludeHiddenLibraries,
    [string[]] $LibraryTitle = @()
)

$outputDir = if ($OutputPath) { $OutputPath }
             elseif ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' }
             else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

$ts = Get-Date -Format 'yyyyMMdd_HHmmss'
$detailCsv = Join-Path $outputDir "SharePoint_VersionCleanup_Detail_$ts.csv"
$summaryCsv = Join-Path $outputDir "SharePoint_VersionCleanup_Summary_$ts.csv"

function Write-ProgressHost {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [ConsoleColor]$ForegroundColor = [ConsoleColor]::DarkGray
    )
    Write-Host ("[{0}] {1}" -f (Get-Date -Format 'HH:mm:ss'), $Message) -ForegroundColor $ForegroundColor
}

function Resolve-PnPClientId {
    param([string]$ConfiguredClientId)

    if (-not [string]::IsNullOrWhiteSpace($ConfiguredClientId)) { return $ConfiguredClientId }
    if (-not [string]::IsNullOrWhiteSpace($env:ENTRAID_APP_ID)) { return $env:ENTRAID_APP_ID }
    if (-not [string]::IsNullOrWhiteSpace($env:AZURE_CLIENT_ID)) { return $env:AZURE_CLIENT_ID }
    throw 'No ClientId provided. Pass -ClientId or set ENTRAID_APP_ID / AZURE_CLIENT_ID.'
}

function Get-TenantAdminUrl {
    param(
        [string]$TenantRootUrl,
        [string]$AnySiteUrl
    )

    $baseUrl = if ($TenantRootUrl) { $TenantRootUrl } elseif ($AnySiteUrl) {
        $siteUri = [Uri]$AnySiteUrl
        "$($siteUri.Scheme)://$($siteUri.Host)"
    }

    if ([string]::IsNullOrWhiteSpace($baseUrl)) { return $null }
    return ($baseUrl -replace '(https://[^.]+)(\.sharepoint\.com)', '$1-admin$2').TrimEnd('/')
}

function Connect-PnPSafely {
    param(
        [Parameter(Mandatory = $true)][string]$Url,
        [Parameter(Mandatory = $true)][string]$ClientId
    )

    Connect-PnPOnline -Url $Url -Interactive -ClientId $ClientId -ErrorAction Stop
}

function Get-VersionCreatedDate {
    param([object]$Version)

    foreach ($name in @('Created', 'LastModifiedDateTime', 'CreatedDate', 'SnapshotDate')) {
        if ($Version.PSObject.Properties.Name -contains $name) {
            $value = $Version.$name
            if ($value) {
                try { return [datetime]$value } catch {}
            }
        }
    }
    return $null
}

function Get-VersionIdentityValue {
    param([object]$Version)

    foreach ($name in @('Id', 'ID', 'VersionLabel', 'Label')) {
        if ($Version.PSObject.Properties.Name -contains $name) {
            $value = $Version.$name
            if (-not [string]::IsNullOrWhiteSpace([string]$value)) { return [string]$value }
        }
    }
    return $null
}

function Get-VersionLabelValue {
    param([object]$Version)

    foreach ($name in @('VersionLabel', 'Label', 'Id', 'ID')) {
        if ($Version.PSObject.Properties.Name -contains $name) {
            $value = $Version.$name
            if (-not [string]::IsNullOrWhiteSpace([string]$value)) { return [string]$value }
        }
    }
    return $null
}

function Get-VersionSizeBytes {
    param([object]$Version)

    foreach ($name in @('Size', 'Length')) {
        if ($Version.PSObject.Properties.Name -contains $name) {
            try { return [int64]($Version.$name ?? 0) } catch {}
        }
    }
    return [int64]0
}

Write-Host ''
Write-Host '  ================================================' -ForegroundColor Cyan
Write-Host '   Remove-SharePointFileVersionsByDate' -ForegroundColor Cyan
Write-Host '  ================================================' -ForegroundColor Cyan
Write-Host ''
Write-Host ("  Cutoff    : older than {0}" -f $BeforeDate.ToString('yyyy-MM-dd HH:mm:ss')) -ForegroundColor Cyan
Write-Host ("  Mode      : {0}" -f $(if ($Apply) { 'Apply (delete matching versions)' } else { 'Preview only (no deletion)' })) -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'DarkGray' })
Write-Host ''

$missingModules = @('PnP.PowerShell') | Where-Object { -not (Get-Module -ListAvailable -Name $_) }
if ($missingModules.Count -gt 0) {
    Write-Host "  [ERROR] Missing required module(s): $($missingModules -join ', ')" -ForegroundColor Red
    exit 1
}

Import-Module PnP.PowerShell -ErrorAction Stop
$resolvedClientId = Resolve-PnPClientId -ConfiguredClientId $ClientId

$allSitesMode = [string]::IsNullOrWhiteSpace($SiteUrl)
if ($allSitesMode -and [string]::IsNullOrWhiteSpace($TenantUrl)) {
    Write-Host '  [ERROR] -TenantUrl is required when scanning all sites.' -ForegroundColor Red
    exit 1
}

$detailRows = [System.Collections.Generic.List[object]]::new()
$summaryRows = [System.Collections.Generic.List[object]]::new()

try {
    $targetSites = [System.Collections.Generic.List[object]]::new()

    if ($allSitesMode) {
        $adminUrl = Get-TenantAdminUrl -TenantRootUrl $TenantUrl -AnySiteUrl $null
        Write-ProgressHost -Message ("Connecting to tenant admin: {0}" -f $adminUrl) -ForegroundColor Cyan
        Connect-PnPSafely -Url $adminUrl -ClientId $resolvedClientId

        $tenantSites = @(Get-PnPTenantSite -Detailed -ErrorAction Stop)
        foreach ($tenantSite in $tenantSites) {
            if (-not $tenantSite.Url) { continue }
            if (-not $IncludeOneDriveSites -and $tenantSite.Url -match '-my\.sharepoint\.com/personal/') { continue }
            $targetSites.Add([PSCustomObject]@{
                Url   = $tenantSite.Url
                Title = $(if ($tenantSite.Title) { $tenantSite.Title } else { $tenantSite.Url })
            }) | Out-Null
        }
    } else {
        $targetSites.Add([PSCustomObject]@{ Url = $SiteUrl.TrimEnd('/'); Title = $SiteUrl.TrimEnd('/') }) | Out-Null
    }

    Write-ProgressHost -Message ("Target sites: {0}" -f $targetSites.Count) -ForegroundColor Green

    $siteIndex = 0
    foreach ($targetSite in $targetSites) {
        $siteIndex++
        Write-ProgressHost -Message ("[{0}/{1}] {2}" -f $siteIndex, $targetSites.Count, $targetSite.Url) -ForegroundColor White

        try {
            Connect-PnPSafely -Url $targetSite.Url -ClientId $resolvedClientId
            $lists = @(Get-PnPList -Includes RootFolder, Hidden, BaseTemplate, Title, ItemCount -ErrorAction Stop | Where-Object {
                $_.BaseTemplate -eq 101 -and ($IncludeHiddenLibraries -or -not $_.Hidden)
            })

            if ($LibraryTitle.Count -gt 0) {
                $lists = @($lists | Where-Object { $_.Title -in $LibraryTitle })
            }

            foreach ($list in $lists) {
                Write-ProgressHost -Message ("Library: {0}" -f $list.Title) -ForegroundColor DarkGray
                try {
                    $files = @(Get-PnPFolderItem -List $list -ItemType File -ErrorAction Stop)
                } catch {
                    Write-ProgressHost -Message ("[WARN] Cannot enumerate files in {0}: {1}" -f $list.Title, $_.Exception.Message) -ForegroundColor Yellow
                    continue
                }

                $fileCounter = 0
                $candidateCount = 0
                $deletedCount = 0
                $candidateBytes = [int64]0
                $deletedBytes = [int64]0

                foreach ($file in $files) {
                    $fileCounter++
                    $fileUrl = [string]$file.ServerRelativeUrl
                    if ([string]::IsNullOrWhiteSpace($fileUrl)) { continue }

                    try {
                        $versions = @(Get-PnPFileVersion -Url $fileUrl -ErrorAction Stop)
                    } catch {
                        $detailRows.Add([PSCustomObject]@{
                            SiteUrl         = $targetSite.Url
                            Library         = $list.Title
                            FileUrl         = $fileUrl
                            FileName        = $file.Name
                            VersionLabel    = $null
                            VersionIdentity = $null
                            VersionCreated  = $null
                            VersionSizeMB   = $null
                            Action          = 'Error'
                            Message         = $_.Exception.Message
                        }) | Out-Null
                        continue
                    }

                    foreach ($version in $versions) {
                        $versionCreated = Get-VersionCreatedDate -Version $version
                        if (-not $versionCreated -or $versionCreated -ge $BeforeDate) { continue }

                        $versionIdentity = Get-VersionIdentityValue -Version $version
                        $versionLabel = Get-VersionLabelValue -Version $version
                        $versionSizeBytes = Get-VersionSizeBytes -Version $version
                        $candidateCount++
                        $candidateBytes += $versionSizeBytes

                        $action = 'WouldDelete'
                        $message = $null
                        if ($Apply) {
                            try {
                                if ($versionIdentity) {
                                    Remove-PnPFileVersion -Url $fileUrl -Identity $versionIdentity -Force -ErrorAction Stop
                                } else {
                                    throw 'Version identity could not be determined.'
                                }
                                $action = 'Deleted'
                                $deletedCount++
                                $deletedBytes += $versionSizeBytes
                            } catch {
                                $action = 'Error'
                                $message = $_.Exception.Message
                            }
                        }

                        $detailRows.Add([PSCustomObject]@{
                            SiteUrl         = $targetSite.Url
                            Library         = $list.Title
                            FileUrl         = $fileUrl
                            FileName        = $file.Name
                            VersionLabel    = $versionLabel
                            VersionIdentity = $versionIdentity
                            VersionCreated  = if ($versionCreated) { $versionCreated.ToString('s') } else { $null }
                            VersionSizeMB   = [math]::Round($versionSizeBytes / 1MB, 3)
                            Action          = $action
                            Message         = $message
                        }) | Out-Null
                    }
                }

                $summaryRows.Add([PSCustomObject]@{
                    SiteUrl           = $targetSite.Url
                    Library           = $list.Title
                    FilesScanned      = $fileCounter
                    CandidateVersions = $candidateCount
                    CandidateSizeMB   = [math]::Round($candidateBytes / 1MB, 2)
                    DeletedVersions   = $deletedCount
                    DeletedSizeMB     = [math]::Round($deletedBytes / 1MB, 2)
                    Mode              = if ($Apply) { 'Apply' } else { 'Preview' }
                }) | Out-Null
            }
        } catch {
            $summaryRows.Add([PSCustomObject]@{
                SiteUrl           = $targetSite.Url
                Library           = $null
                FilesScanned      = 0
                CandidateVersions = 0
                CandidateSizeMB   = 0
                DeletedVersions   = 0
                DeletedSizeMB     = 0
                Mode              = if ($Apply) { 'Apply' } else { 'Preview' }
                Error             = $_.Exception.Message
            }) | Out-Null
            Write-ProgressHost -Message ("[ERROR] Site failed: {0}" -f $_.Exception.Message) -ForegroundColor Red
        }
    }
} finally {
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
}

$detailRows | Export-Csv -Path $detailCsv -NoTypeInformation -Encoding UTF8
$summaryRows | Export-Csv -Path $summaryCsv -NoTypeInformation -Encoding UTF8

$totalCandidates = ($summaryRows | Measure-Object -Property CandidateVersions -Sum).Sum ?? 0
$totalDeleted = ($summaryRows | Measure-Object -Property DeletedVersions -Sum).Sum ?? 0
$totalCandidateMB = ($summaryRows | Measure-Object -Property CandidateSizeMB -Sum).Sum ?? 0
$totalDeletedMB = ($summaryRows | Measure-Object -Property DeletedSizeMB -Sum).Sum ?? 0

Write-Host ''
Write-Host '  ================================================' -ForegroundColor Cyan
Write-Host '   Summary' -ForegroundColor Cyan
Write-Host '  ================================================' -ForegroundColor Cyan
Write-Host ("  Detail   : {0}" -f $detailCsv) -ForegroundColor Green
Write-Host ("  Summary  : {0}" -f $summaryCsv) -ForegroundColor Green
Write-Host ("  Candidates: {0} version(s) | {1} MB" -f $totalCandidates, [math]::Round($totalCandidateMB, 2)) -ForegroundColor Yellow
Write-Host ("  Deleted   : {0} version(s) | {1} MB" -f $totalDeleted, [math]::Round($totalDeletedMB, 2)) -ForegroundColor $(if ($Apply) { 'Magenta' } else { 'DarkGray' })
Write-Host ''