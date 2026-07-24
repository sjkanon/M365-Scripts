#Requires -Version 5.1
<#
.SYNOPSIS
    Map SharePoint Online / OneDrive document libraries to drive letters via WebDAV.

.DESCRIPTION
    Maps one or more SharePoint/OneDrive document library URLs to persistent drive
    letters using the WebDAV redirector (net use), for use as a per-user logon script
    (Intune Win32 app or a scheduled task triggered at logon).

    Converts each https:// library URL into its WebDAV UNC form
    (\\<host>@SSL\DavWWWRoot\<path>) and maps it with `net use`. This relies on the
    Windows "WebClient" service (WebDAV redirector) — present and running by default
    on Windows 10/11, but usually not installed on Windows Server without the
    "Desktop Experience" feature.

    Run without -Apply for a dry run — shows the resolved WebDAV path and the exact
    `net use` command for each mapping without changing anything. Run with -Apply to
    perform the mapping.

    Auth: relies on the signed-in user already having a valid session/SSO to the
    tenant (same as browser-based WebDAV access to SharePoint). This is not app-only
    Graph auth — if a mapping prompts for credentials or fails with access denied,
    the user's session doesn't yet have access to that library.

.PARAMETER MappingsCsv
    Path to a CSV with columns: DriveLetter, Url, and optional Label.
    Example row: C,https://contoso.sharepoint.com/sites/Finance/Shared Documents,Finance Docs

.PARAMETER DriveLetter
    Single ad hoc mapping: drive letter to use (e.g. "Z"). Use with -Url instead of -MappingsCsv.

.PARAMETER Url
    Single ad hoc mapping: SharePoint/OneDrive document library URL. Use with -DriveLetter.

.PARAMETER Label
    Optional friendly name for the single ad hoc mapping (sets the drive label where supported).

.PARAMETER RemoveExisting
    Remove any existing mapping on the target drive letter(s) first (`net use <letter>: /delete /y`)
    before mapping, so reruns don't fail with "already in use".

.PARAMETER Persist
    Make the mapping persistent across reboots (`net use ... /persistent:yes`). Off by default,
    since a logon script typically remaps on every sign-in anyway.

.PARAMETER Apply
    Actually perform the mapping. Without this switch, only a dry run is shown.

.PARAMETER OutputPath
    Log file folder (default: $env:TEMP — logon scripts usually run in user context,
    not as admin, so C:\Temp may not be writable).

.EXAMPLE
    # Dry run from a CSV of mappings
    .\New-CloudDriveMapping.ps1 -MappingsCsv .\mappings.csv

.EXAMPLE
    # Apply mappings from CSV, clearing any existing mapping on those letters first
    .\New-CloudDriveMapping.ps1 -MappingsCsv .\mappings.csv -RemoveExisting -Apply

.EXAMPLE
    # Single ad hoc mapping
    .\New-CloudDriveMapping.ps1 -DriveLetter Z -Url "https://contoso.sharepoint.com/sites/Finance/Shared Documents" -Apply

.NOTES
    Author  : Sjoerd Kanon
    Platform: Windows only (WebDAV redirector / WebClient service)
#>
[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Csv')]
param (
    [Parameter(ParameterSetName = 'Csv', Mandatory)]
    [string] $MappingsCsv,

    [Parameter(ParameterSetName = 'Single', Mandatory)]
    [string] $DriveLetter,

    [Parameter(ParameterSetName = 'Single', Mandatory)]
    [string] $Url,

    [Parameter(ParameterSetName = 'Single')]
    [string] $Label,

    [switch] $RemoveExisting,
    [switch] $Persist,
    [switch] $Apply,
    [string] $OutputPath = $env:TEMP
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function ConvertTo-WebDavPath {
    param([Parameter(Mandatory)][string] $HttpsUrl)

    $uri = [Uri]$HttpsUrl
    $decodedPath = [Uri]::UnescapeDataString($uri.AbsolutePath) -replace '/', '\'
    return "\\$($uri.Host)@SSL\DavWWWRoot$decodedPath"
}

function Test-WebClientService {
    $svc = Get-Service -Name WebClient -ErrorAction SilentlyContinue
    if (-not $svc) {
        Write-Warning "The 'WebClient' (WebDAV redirector) service is not installed. On Windows Server, enable the 'Desktop Experience' feature first. Drive mapping will fail without it."
        return $false
    }
    if ($svc.Status -ne 'Running') {
        Write-Warning "The 'WebClient' service is installed but not running (Status: $($svc.Status)). Attempting to start it..."
        try { Start-Service -Name WebClient } catch { Write-Warning "Could not start WebClient service: $($_.Exception.Message)" }
    }
    return $true
}

function Get-Mappings {
    if ($PSCmdlet.ParameterSetName -eq 'Single') {
        return @([PSCustomObject]@{ DriveLetter = $DriveLetter.TrimEnd(':'); Url = $Url; Label = $Label })
    }

    if (-not (Test-Path -LiteralPath $MappingsCsv)) {
        throw "Mappings CSV not found: $MappingsCsv"
    }

    $rows = Import-Csv -LiteralPath $MappingsCsv
    foreach ($row in $rows) {
        if (-not $row.DriveLetter -or -not $row.Url) {
            throw "Mappings CSV row missing DriveLetter or Url: $($row | Out-String)"
        }
        [PSCustomObject]@{ DriveLetter = $row.DriveLetter.TrimEnd(':'); Url = $row.Url; Label = $row.Label }
    }
}

# ── Output ─────────────────────────────────────────────────────────────────────
if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath | Out-Null }
$ts        = Get-Date -Format 'yyyyMMdd_HHmmss'
$logFile   = Join-Path $OutputPath "CloudDriveMapping_$ts.log"
$results   = [System.Collections.Generic.List[object]]::new()

Write-Host ''
Write-Host "  Mode: $(if ($Apply) { 'APPLY — drives will be mapped' } else { 'DRY RUN — no changes will be made' })" -ForegroundColor Cyan
Write-Host ''

$webClientOk = Test-WebClientService
$mappings = @(Get-Mappings)

foreach ($m in $mappings) {
    $letter  = "$($m.DriveLetter):"
    $webdav  = ConvertTo-WebDavPath -HttpsUrl $m.Url
    $labelTxt = if ($m.Label) { " ($($m.Label))" } else { '' }

    Write-Host "  $letter -> $($m.Url)$labelTxt" -ForegroundColor Yellow
    Write-Host "    WebDAV path: $webdav" -ForegroundColor DarkGray

    $status = 'Skipped (dry run)'
    if ($Apply -and $webClientOk -and $PSCmdlet.ShouldProcess($letter, "Map to $webdav")) {
        if ($RemoveExisting) {
            & net use $letter /delete /y *> $null
        }

        $persistFlag = if ($Persist) { 'yes' } else { 'no' }
        $netUseOutput = & net use $letter $webdav /persistent:$persistFlag 2>&1
        if ($LASTEXITCODE -eq 0) {
            $status = 'Mapped'
            Write-Host "    [OK] Mapped $letter" -ForegroundColor Green
        } else {
            $status = "Failed: $($netUseOutput -join ' ')"
            Write-Host "    [FAIL] $status" -ForegroundColor Red
        }
    }

    $results.Add([PSCustomObject]@{
        DriveLetter = $letter
        Url         = $m.Url
        WebDavPath  = $webdav
        Status      = $status
    })
}

$results | ForEach-Object { "$($_.DriveLetter) | $($_.Url) | $($_.WebDavPath) | $($_.Status)" } | Out-File -FilePath $logFile -Encoding UTF8

Write-Host ''
Write-Host "  Log: $logFile" -ForegroundColor DarkGray
if (-not $Apply) {
    Write-Host "  Dry run only — rerun with -Apply to map these drives." -ForegroundColor Yellow
}
