#Requires -Version 5.1
<#
.SYNOPSIS
    Bulk-create M365 users from a CSV file via Microsoft Graph.

.DESCRIPTION
    Reads a CSV file and creates M365 user accounts in Entra ID.
    Defaults to dry-run mode — pass -Apply to actually create accounts.

    If no Password column is present in the CSV, a unique 16-character random
    password is generated for each user. Passwords are written to the results CSV.

    A LicenseSkuId column in the CSV assigns a license per user. Overridden for
    all users by the -LicenseSkuId parameter if provided.

    Required CSV columns:
      UserPrincipalName   UPN of the new user
      DisplayName         Display name

    Optional CSV columns:
      GivenName           First name
      Surname             Last name (also accepts: LastName)
      Department          Department
      JobTitle            Job title
      MobilePhone         Mobile phone number
      UsageLocation       Two-letter ISO country code (overrides -UsageLocation)
      Password            Initial password (auto-generated if absent)
      LicenseSkuId        SKU part number to assign (e.g. ENTERPRISEPACK)

.PARAMETER CsvPath
    Path to the input CSV file.

.PARAMETER Apply
    Actually create the accounts. Without this switch the script runs in dry-run mode.

.PARAMETER UsageLocation
    Default two-letter ISO country code for all users. Default: NL.
    Overridden per row if a UsageLocation column is present in the CSV.

.PARAMETER LicenseSkuId
    Assign this license to all created users. Overrides any LicenseSkuId column
    in the CSV.

.PARAMETER NoPasswordReset
    Do not force a password change on first sign-in.

.PARAMETER OutputPath
    CSV report path. Default: .\CreatedAccounts_<timestamp>.csv

.PARAMETER TenantId
    Entra ID tenant ID or domain. Optional if already connected.

.EXAMPLE
    # Dry run (default — shows what would be created)
    .\Import-M365Users.ps1 -CsvPath .\users.csv

.EXAMPLE
    # Create accounts
    .\Import-M365Users.ps1 -CsvPath .\users.csv -Apply

.EXAMPLE
    # Create with a default license for all users
    .\Import-M365Users.ps1 -CsvPath .\users.csv -LicenseSkuId "ENTERPRISEPACK" -Apply
#>
[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string] $CsvPath,

    [switch] $Apply,

    [string] $UsageLocation = 'NL',

    [string] $LicenseSkuId,

    [switch] $NoPasswordReset,

    [string] $OutputPath,

    [string] $TenantId
)

# ── Output folder ─────────────────────────────────────────────────────────────
$outputDir = if ($IsWindows -or $env:OS -eq 'Windows_NT') { 'C:\Temp' } else { "$HOME/Downloads" }
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }
if (-not $OutputPath) { $OutputPath = Join-Path $outputDir "CreatedAccounts_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv" }

# ── Password generator ────────────────────────────────────────────────────────
function New-RandomPassword {
    $upper   = 'ABCDEFGHJKLMNPQRSTUVWXYZ'.ToCharArray()
    $lower   = 'abcdefghjkmnpqrstuvwxyz'.ToCharArray()
    $digits  = '23456789'.ToCharArray()
    $special = '!@#$%&*'.ToCharArray()
    $all     = $upper + $lower + $digits + $special

    $chars  = @(
        $upper   | Get-Random
        $lower   | Get-Random
        $digits  | Get-Random
        $special | Get-Random
    )
    $chars += 1..12 | ForEach-Object { $all | Get-Random }
    ($chars | Sort-Object { Get-Random }) -join ''
}

# ── Connection ────────────────────────────────────────────────────────────────
$script:ConnectedHere = $false
$ctx = Get-MgContext -ErrorAction SilentlyContinue
if (-not $ctx) {
    $connectParams = @{
        Scopes    = @('User.ReadWrite.All', 'Directory.ReadWrite.All')
        NoWelcome = $true
    }
    if ($TenantId) { $connectParams['TenantId'] = $TenantId }
    Connect-MgGraph @connectParams
    $script:ConnectedHere = $true
}

# ── Load and validate CSV ─────────────────────────────────────────────────────
$rows = Import-Csv -Path $CsvPath

if ($rows.Count -eq 0) {
    Write-Error "CSV is empty."
    if ($script:ConnectedHere) { Disconnect-MgGraph | Out-Null }
    exit 1
}

$cols = $rows[0].PSObject.Properties.Name

$upnCol = $cols | Where-Object { $_ -match '^userprincipalname$|^upn$' } | Select-Object -First 1
$dnCol  = $cols | Where-Object { $_ -match '^displayname$' }             | Select-Object -First 1

if (-not $upnCol) {
    Write-Error "CSV must have a 'UserPrincipalName' or 'UPN' column. Found: $($cols -join ', ')"
    if ($script:ConnectedHere) { Disconnect-MgGraph | Out-Null }
    exit 1
}
if (-not $dnCol) {
    Write-Error "CSV must have a 'DisplayName' column. Found: $($cols -join ', ')"
    if ($script:ConnectedHere) { Disconnect-MgGraph | Out-Null }
    exit 1
}

# Helper: get optional column value or null
function Get-Col { param($Row, $Name) ($Row.PSObject.Properties | Where-Object Name -eq $Name | Select-Object -First 1)?.Value }

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Import M365 Users" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host ""

if (-not $Apply) {
    Write-Host "  ================================================" -ForegroundColor Yellow
    Write-Host "   DRY RUN MODE — no accounts will be created" -ForegroundColor Yellow
    Write-Host "   Add -Apply to create the accounts." -ForegroundColor Yellow
    Write-Host "  ================================================" -ForegroundColor Yellow
    Write-Host ""
}

Write-Host "  Input file : $CsvPath"
Write-Host "  Rows found : $($rows.Count)"
Write-Host ""

# ── Cache SKUs if needed ──────────────────────────────────────────────────────
$skuCache = @{}
if ($LicenseSkuId -or ($cols | Where-Object { $_ -match 'licenseskuid' })) {
    Get-MgSubscribedSku -ErrorAction SilentlyContinue | ForEach-Object {
        $skuCache[$_.SkuPartNumber] = $_.SkuId
        $skuCache[$_.SkuId]         = $_.SkuId
    }
}

# ── Process rows ──────────────────────────────────────────────────────────────
$results = [System.Collections.Generic.List[PSCustomObject]]::new()
$created = 0
$skipped = 0
$errors  = 0

foreach ($row in $rows) {
    $upn = (Get-Col $row $upnCol)?.Trim()
    $dn  = (Get-Col $row $dnCol)?.Trim()

    if (-not $upn -or $upn -notmatch '@') {
        Write-Host "  [SKIP] Invalid UPN: '$upn'" -ForegroundColor DarkYellow
        $skipped++
        continue
    }

    $given   = Get-Col $row 'GivenName'
    $surname = (Get-Col $row 'Surname') ?? (Get-Col $row 'LastName')
    $dept    = Get-Col $row 'Department'
    $title   = Get-Col $row 'JobTitle'
    $mobile  = Get-Col $row 'MobilePhone'
    $loc     = (Get-Col $row 'UsageLocation') ?? $UsageLocation
    $pw      = Get-Col $row 'Password'
    $rowSku  = Get-Col $row 'LicenseSkuId'

    $generated = $false
    if (-not $pw) { $pw = New-RandomPassword; $generated = $true }

    $effectiveSku = if ($LicenseSkuId) { $LicenseSkuId } elseif ($rowSku) { $rowSku } else { $null }

    Write-Host "  -> $upn" -NoNewline

    if (-not $Apply) {
        $licNote = if ($effectiveSku) { " | license: $effectiveSku" } else { '' }
        Write-Host "  [DRY RUN — $dn$licNote]" -ForegroundColor DarkYellow
        $results.Add([PSCustomObject]@{
            UserPrincipalName = $upn
            DisplayName       = $dn
            Password          = if ($generated) { $pw } else { '(provided)' }
            License           = $effectiveSku
            Status            = 'DryRun'
            Timestamp         = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        })
        $skipped++
        continue
    }

    # Build params
    $userParams = @{
        UserPrincipalName = $upn
        DisplayName       = $dn
        AccountEnabled    = $true
        UsageLocation     = $loc
        PasswordProfile   = @{
            Password                      = $pw
            ForceChangePasswordNextSignIn = -not $NoPasswordReset
        }
    }
    if ($given)   { $userParams['GivenName']  = $given }
    if ($surname) { $userParams['Surname']     = $surname }
    if ($dept)    { $userParams['Department']  = $dept }
    if ($title)   { $userParams['JobTitle']    = $title }
    if ($mobile)  { $userParams['MobilePhone'] = $mobile }

    try {
        $newUser = New-MgUser @userParams -ErrorAction Stop
        Write-Host "  Created ($dn)" -ForegroundColor Green
        $status = 'Created'
        $created++

        # Assign license
        if ($effectiveSku) {
            $skuId = $skuCache[$effectiveSku]
            if ($skuId) {
                try {
                    Set-MgUserLicense -UserId $newUser.Id -BodyParameter @{
                        AddLicenses    = @(@{ SkuId = $skuId })
                        RemoveLicenses = @()
                    } | Out-Null
                    $status = "Created+Licensed($effectiveSku)"
                } catch {
                    Write-Host "    [WARN] License '$effectiveSku' failed: $($_.Exception.Message)" -ForegroundColor Yellow
                    $status = "Created+LicenseFailed"
                }
            } else {
                Write-Host "    [WARN] SKU '$effectiveSku' not found in tenant." -ForegroundColor Yellow
                $status = "Created+SkuNotFound"
            }
        }
    } catch {
        Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
        $status = "Error: $($_.Exception.Message)"
        $errors++
    }

    $results.Add([PSCustomObject]@{
        UserPrincipalName = $upn
        DisplayName       = $dn
        Password          = if ($generated) { $pw } else { '(provided)' }
        License           = $effectiveSku
        Status            = $status
        Timestamp         = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    })
}

# ── Summary ───────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Summary" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
if ($Apply) {
    Write-Host "  Created : $created" -ForegroundColor Green
    Write-Host "  Skipped : $skipped" -ForegroundColor DarkYellow
    Write-Host "  Errors  : $errors"  -ForegroundColor Red
} else {
    Write-Host "  Would create : $($rows.Count - $skipped) user(s)" -ForegroundColor Yellow
    Write-Host "  Skipped      : $skipped (invalid UPN)" -ForegroundColor DarkYellow
}

# ── Export ────────────────────────────────────────────────────────────────────
$results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host ""
Write-Host "  Report saved : $OutputPath" -ForegroundColor Cyan
if ($Apply -and $created -gt 0) {
    Write-Host "  Passwords are in the report — share securely and advise users to change on first login." -ForegroundColor DarkYellow
}
Write-Host ""

if ($script:ConnectedHere) { Disconnect-MgGraph | Out-Null }
