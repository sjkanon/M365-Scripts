#Requires -Version 5.1
<#
.SYNOPSIS
    Automatically updates the minimum iOS version in an Intune compliance policy.

.DESCRIPTION
    Fetches the latest iOS version from Apple's RSS feed and updates the specified
    Intune compliance policy via Microsoft Graph API.
    Designed to run as a scheduled task on a Windows server.

.PARAMETER ConfigPath
    Path to the configuration file (default: config.json in script directory).

.PARAMETER LogPath
    Path to the log directory (default: logs\ in script directory).

.PARAMETER WhatIf
    Dry-run mode — shows what would happen without making changes.

.EXAMPLE
    .\Update-iOSCompliancePolicy.ps1

.EXAMPLE
    .\Update-iOSCompliancePolicy.ps1 -WhatIf

.EXAMPLE
    .\Update-iOSCompliancePolicy.ps1 -ConfigPath "C:\Scripts\config.json"

.NOTES
    Requires Graph API permission: DeviceManagementConfiguration.ReadWrite.All
#>
[CmdletBinding(SupportsShouldProcess)]
param (
    [string] $ConfigPath = "$PSScriptRoot\config.json",
    [string] $LogPath    = "$PSScriptRoot\logs",
    [switch] $WhatIf
)

# ── Logging ───────────────────────────────────────────────────────────────────
function Write-Log {
    param (
        [string] $Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')]
        [string] $Level = 'INFO'
    )

    $timestamp  = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logMessage = "[$timestamp] [$($Level.PadRight(7))] $Message"

    switch ($Level) {
        'INFO'    { Write-Host $logMessage -ForegroundColor Cyan }
        'WARN'    { Write-Host $logMessage -ForegroundColor Yellow }
        'ERROR'   { Write-Host $logMessage -ForegroundColor Red }
        'SUCCESS' { Write-Host $logMessage -ForegroundColor Green }
    }

    if (-not (Test-Path $LogPath)) {
        New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
    }

    $logFile = Join-Path $LogPath "compliance-updater-$(Get-Date -Format 'yyyy-MM').log"
    Add-Content -Path $logFile -Value $logMessage
}

# ── Configuration ─────────────────────────────────────────────────────────────
function Get-Config {
    param ([string] $Path)

    if (-not (Test-Path $Path)) {
        Write-Log "Config file not found: $Path" -Level ERROR
        Write-Log "Copy config.example.json to config.json and fill in your values." -Level ERROR
        exit 1
    }

    try {
        $config = Get-Content $Path -Raw | ConvertFrom-Json
        Write-Log "Configuration loaded from: $Path"
        return $config
    } catch {
        Write-Log "Failed to parse config file: $_" -Level ERROR
        exit 1
    }
}

function Test-Config {
    param ($Config)

    $required = @('TenantId', 'ClientId', 'ClientSecret', 'CompliancePolicyId')
    $missing  = $required | Where-Object { -not $Config.$_ -or $Config.$_ -eq 'FILL_IN' }

    if ($missing) {
        Write-Log "Missing required config fields: $($missing -join ', ')" -Level ERROR
        exit 1
    }

    Write-Log "Configuration validated."
}

# ── Authentication ────────────────────────────────────────────────────────────
function Get-GraphToken {
    param ($Config)

    Write-Log "Requesting access token..."

    $body = @{
        grant_type    = 'client_credentials'
        scope         = 'https://graph.microsoft.com/.default'
        client_id     = $Config.ClientId
        client_secret = $Config.ClientSecret
    }

    try {
        $response = Invoke-RestMethod `
            -Uri    "https://login.microsoftonline.com/$($Config.TenantId)/oauth2/v2.0/token" `
            -Method POST `
            -Body   $body `
            -ErrorAction Stop

        Write-Log "Access token obtained." -Level SUCCESS
        return $response.access_token
    } catch {
        Write-Log "Failed to obtain access token: $($_.Exception.Message)" -Level ERROR
        exit 1
    }
}

# ── iOS version detection ─────────────────────────────────────────────────────
function Get-LatestIOSVersion {
    Write-Log "Fetching latest iOS version from Apple..."

    try {
        $rss     = Invoke-RestMethod -Uri 'https://developer.apple.com/news/releases/rss/releases.rss' -ErrorAction Stop
        $entries = $rss.channel.item.title | Where-Object { $_ -match '^iOS \d+(\.\d+)*' }

        if (-not $entries) { throw 'No iOS entries found in Apple RSS feed.' }

        $version = ($entries | Select-Object -First 1) -replace '^iOS (\S+).*', '$1'
        $version = $version.Trim()

        if ($version -notmatch '^\d+(\.\d+)+$') { throw "Unexpected version format: '$version'" }

        Write-Log "Latest iOS version: $version" -Level SUCCESS
        return $version
    } catch {
        Write-Log "Apple RSS failed: $($_.Exception.Message) — trying fallback..." -Level WARN

        try {
            $page  = Invoke-WebRequest -Uri 'https://support.apple.com/en-us/111900' -UseBasicParsing -ErrorAction Stop
            $match = [regex]::Match($page.Content, 'iOS (\d+\.\d+(?:\.\d+)?)')

            if (-not $match.Success) { throw 'Could not parse iOS version from Apple support page.' }

            $version = $match.Groups[1].Value
            Write-Log "Latest iOS version (fallback): $version" -Level SUCCESS
            return $version
        } catch {
            Write-Log "Fallback also failed: $($_.Exception.Message)" -Level ERROR
            exit 1
        }
    }
}

# ── Intune compliance policy ───────────────────────────────────────────────────
function Get-CompliancePolicy {
    param ([string] $PolicyId, [hashtable] $Headers)

    Write-Log "Fetching compliance policy (ID: $PolicyId)..."

    try {
        $policy = Invoke-RestMethod `
            -Uri     "https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies/$PolicyId" `
            -Headers $Headers `
            -Method  GET `
            -ErrorAction Stop

        Write-Log "Policy: '$($policy.displayName)' — current minimum iOS: $($policy.osMinimumVersion)"
        return $policy
    } catch {
        Write-Log "Failed to fetch compliance policy: $($_.Exception.Message)" -Level ERROR
        exit 1
    }
}

function Set-CompliancePolicyVersion {
    param ([string] $PolicyId, [string] $NewVersion, [hashtable] $Headers, [switch] $WhatIf)

    if ($WhatIf) {
        Write-Log "[WHATIF] Would update compliance policy to iOS $NewVersion" -Level WARN
        return
    }

    Write-Log "Updating compliance policy to iOS $NewVersion..."

    $body = @{ osMinimumVersion = $NewVersion } | ConvertTo-Json

    try {
        Invoke-RestMethod `
            -Uri         "https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies/$PolicyId" `
            -Method      PATCH `
            -Headers     $Headers `
            -Body        $body `
            -ContentType 'application/json' `
            -ErrorAction Stop

        Write-Log "Compliance policy updated to iOS $NewVersion." -Level SUCCESS
    } catch {
        Write-Log "Failed to update compliance policy: $($_.Exception.Message)" -Level ERROR
        exit 1
    }
}

# ── Main ──────────────────────────────────────────────────────────────────────
Write-Log "========================================"
Write-Log " Intune iOS Compliance Updater started"
Write-Log "========================================"

if ($WhatIf) { Write-Log "WHATIF mode — no changes will be made." -Level WARN }

$config = Get-Config -Path $ConfigPath
Test-Config -Config $config

$token   = Get-GraphToken -Config $config
$headers = @{ Authorization = "Bearer $token"; 'Content-Type' = 'application/json' }

$latestVersion  = Get-LatestIOSVersion
$currentPolicy  = Get-CompliancePolicy -PolicyId $config.CompliancePolicyId -Headers $headers
$currentVersion = $currentPolicy.osMinimumVersion

if ($currentVersion -eq $latestVersion) {
    Write-Log "Policy is up to date (iOS $currentVersion). No changes needed." -Level SUCCESS
} else {
    Write-Log "Update needed: $currentVersion -> $latestVersion"
    Set-CompliancePolicyVersion -PolicyId $config.CompliancePolicyId -NewVersion $latestVersion -Headers $headers -WhatIf:$WhatIf
}

Write-Log "========================================"
Write-Log " Script completed"
Write-Log "========================================"
