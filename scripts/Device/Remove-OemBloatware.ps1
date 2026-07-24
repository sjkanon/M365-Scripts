#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Remove OEM bloatware and generic Microsoft Store junk apps from a Windows endpoint.

.DESCRIPTION
    Detects the device manufacturer (Get-CimInstance Win32_ComputerSystem) and removes
    known bloatware for that manufacturer via winget, plus a generic list of consumer
    Microsoft Store apps (Xbox, Solitaire, Bing News/Weather, etc.) via Remove-AppxPackage.

    Run without -Apply for a dry run — lists every matching app actually installed on
    this device without removing anything. Run with -Apply to remove them.

    The bloatware lists below are a starting point, not exhaustive — package/app IDs
    change over time and vary by device model and preload image. Run `winget list` and
    `Get-AppxPackage | Select Name` on a representative device first and adjust the
    lists if something is missing or something wanted gets flagged.

    Output:
      - Summary CSV saved to C:\Temp\ with what was found/removed per app.

.PARAMETER Apply
    Actually remove the matched apps. Without this switch, only a scan is performed.

.PARAMETER SkipOem
    Skip manufacturer-specific winget removal (HP/Lenovo/Dell), only process the generic
    Microsoft Store app list.

.PARAMETER SkipAppx
    Skip the generic Microsoft Store AppX list, only process manufacturer-specific apps.

.PARAMETER Manufacturer
    Override the auto-detected manufacturer (HP, Lenovo, or Dell). Useful for testing
    the removal list on a device of a different make, or when Win32_ComputerSystem
    reports an unexpected string.

.PARAMETER OutputPath
    Override the default output folder (default: C:\Temp\).

.EXAMPLE
    # Dry run — see what would be removed on this device
    .\Remove-OemBloatware.ps1

.EXAMPLE
    # Actually remove OEM + generic bloatware
    .\Remove-OemBloatware.ps1 -Apply

.EXAMPLE
    # Only remove the generic Microsoft Store junk, leave OEM apps alone
    .\Remove-OemBloatware.ps1 -Apply -SkipOem

.NOTES
    Author  : Sjoerd Kanon
    Platform: Windows only (winget + AppX)
#>
[CmdletBinding(SupportsShouldProcess)]
param (
    [switch] $Apply,
    [switch] $SkipOem,
    [switch] $SkipAppx,
    [ValidateSet('HP', 'Lenovo', 'Dell')]
    [string] $Manufacturer,
    [string] $OutputPath = 'C:\Temp'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'SilentlyContinue'

# ── OEM winget bloatware lists ────────────────────────────────────────────────
# Package IDs as seen via `winget list`. Verify against your own fleet before relying
# on this list wholesale — OEMs change preload images and package IDs over time.
$OemPackages = @{
    'HP' = @(
        'HP Inc.HPSupportAssistant'
        'HP Inc.MyHP'
        'HP Inc.HPSystemInformation'
        'HP Inc.HPPrivacySettings'
        'HP Inc.HPPCHardwareDiagnosticsWindows'
        'HP Inc.HPDesktopSupportUtilities'
        'HP Inc.HPQuickDrop'
        'HP Inc.HPEasyClean'
    )
    'Lenovo' = @(
        'E0469063.LenovoCompanion'
        '9WZDNCRFJ4MV'   # Lenovo Vantage (Microsoft Store)
        'E046963F.LenovoSystemInterfaceFoundation'
        'E0469028.LenovoUtility'
        '4505Fortemedia.LenovoAudioDriverCompatibleWithFalcon'
    )
    'Dell' = @(
        'DellInc.PartnerPromo'
        'DellInc.DellDigitalDelivery'
        'DellInc.DellSupportAssistforPCs'
        'DellInc.DellCommandUpdate'
        'DellInc.DellPeripheralManager'
        'DellInc.DellOptimizer'
        'DellInc.DellCustomerConnect'
    )
}

# ── Generic consumer Microsoft Store junk (any manufacturer) ───────────────────
$GenericAppxPackages = @(
    'Microsoft.BingNews'
    'Microsoft.BingWeather'
    'Microsoft.GamingApp'
    'Microsoft.MicrosoftSolitaireCollection'
    'Microsoft.Todos'
    'Microsoft.XboxApp'
    'Microsoft.XboxGamingOverlay'
    'Microsoft.XboxIdentityProvider'
    'Microsoft.XboxSpeechToTextOverlay'
    'Microsoft.Xbox.TCUI'
    'Microsoft.YourPhone'
    'Microsoft.ZuneMusic'
    'Microsoft.ZuneVideo'
    'Clipchamp.Clipchamp'
    'MicrosoftCorporationII.MicrosoftFamily'
    'Microsoft.549981C3F5F10'   # Cortana
)

# ── Output ─────────────────────────────────────────────────────────────────────
if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath | Out-Null }
$ts        = Get-Date -Format 'yyyyMMdd_HHmmss'
$reportCsv = Join-Path $OutputPath "OemBloatwareRemoval_$ts.csv"
$results   = [System.Collections.Generic.List[object]]::new()

function Resolve-Manufacturer {
    if ($Manufacturer) { return $Manufacturer }

    $raw = (Get-CimInstance -ClassName Win32_ComputerSystem).Manufacturer
    switch -Regex ($raw) {
        'HP|Hewlett' { return 'HP' }
        'Lenovo'     { return 'Lenovo' }
        'Dell'       { return 'Dell' }
        default      { return $null }
    }
}

function Test-WingetAvailable {
    return [bool](Get-Command winget.exe -ErrorAction SilentlyContinue)
}

function Remove-OemWingetPackages {
    param([string[]] $PackageIds)

    if (-not (Test-WingetAvailable)) {
        Write-Warning 'winget.exe not found — skipping OEM package removal. Install App Installer from the Microsoft Store.'
        return
    }

    $installed = winget list --accept-source-agreements 2>$null

    foreach ($id in $PackageIds) {
        $found = $installed | Select-String -SimpleMatch $id -Quiet
        if (-not $found) {
            $results.Add([PSCustomObject]@{ Type = 'OEM (winget)'; Id = $id; Found = $false; Removed = $false })
            continue
        }

        Write-Host "  [Found]  $id" -ForegroundColor Yellow
        $removed = $false
        if ($Apply -and $PSCmdlet.ShouldProcess($id, 'winget uninstall')) {
            winget uninstall --id $id --silent --accept-source-agreements --disable-interactivity | Out-Null
            $removed = ($LASTEXITCODE -eq 0)
            Write-Host "  [$(if ($removed) { 'Removed' } else { 'Failed ' })] $id" -ForegroundColor $(if ($removed) { 'Green' } else { 'Red' })
        }
        $results.Add([PSCustomObject]@{ Type = 'OEM (winget)'; Id = $id; Found = $true; Removed = $removed })
    }
}

function Remove-GenericAppxPackages {
    param([string[]] $PackageNames)

    foreach ($name in $PackageNames) {
        $pkgs = Get-AppxPackage -AllUsers -Name $name
        if (-not $pkgs) {
            $results.Add([PSCustomObject]@{ Type = 'Generic (AppX)'; Id = $name; Found = $false; Removed = $false })
            continue
        }

        Write-Host "  [Found]  $name" -ForegroundColor Yellow
        $removed = $false
        if ($Apply -and $PSCmdlet.ShouldProcess($name, 'Remove-AppxPackage -AllUsers')) {
            $pkgs | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
            $removed = -not (Get-AppxPackage -AllUsers -Name $name)
            Write-Host "  [$(if ($removed) { 'Removed' } else { 'Failed ' })] $name" -ForegroundColor $(if ($removed) { 'Green' } else { 'Red' })
        }
        $results.Add([PSCustomObject]@{ Type = 'Generic (AppX)'; Id = $name; Found = $true; Removed = $removed })
    }
}

# ── Run ────────────────────────────────────────────────────────────────────────
Write-Host ''
Write-Host "  Mode: $(if ($Apply) { 'APPLY — apps will be removed' } else { 'DRY RUN — no changes will be made' })" -ForegroundColor Cyan
Write-Host ''

if (-not $SkipOem) {
    $detected = Resolve-Manufacturer
    if (-not $detected) {
        Write-Warning "Could not map manufacturer to a known bloatware list (raw: '$((Get-CimInstance -ClassName Win32_ComputerSystem).Manufacturer)'). Skipping OEM removal — use -Manufacturer to force one."
    } else {
        Write-Host "  Manufacturer: $detected" -ForegroundColor Cyan
        Remove-OemWingetPackages -PackageIds $OemPackages[$detected]
    }
}

if (-not $SkipAppx) {
    Write-Host ''
    Write-Host '  Generic Microsoft Store apps:' -ForegroundColor Cyan
    Remove-GenericAppxPackages -PackageNames $GenericAppxPackages
}

# ── Report ─────────────────────────────────────────────────────────────────────
$results | Export-Csv -Path $reportCsv -NoTypeInformation -Encoding UTF8

Write-Host ''
$foundCount   = ($results | Where-Object Found).Count
$removedCount = ($results | Where-Object Removed).Count
Write-Host "  Found: $foundCount   Removed: $removedCount" -ForegroundColor Cyan
Write-Host "  Report: $reportCsv" -ForegroundColor DarkGray
if (-not $Apply -and $foundCount -gt 0) {
    Write-Host "  Dry run only — rerun with -Apply to remove these apps." -ForegroundColor Yellow
}
