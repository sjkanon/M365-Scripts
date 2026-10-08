#Requires -Version 7.0
<#
.SYNOPSIS
    Automatically updates the minimum iOS version in an Intune compliance policy.

.DESCRIPTION
    Fetches the latest released iOS version from Apple's RSS feed and raises the
    minimum OS version of the specified Intune iOS compliance policy via Microsoft
    Graph. Designed to run as a scheduled task on a Windows server.

    Only final releases count (beta and RC entries are skipped), the highest version
    in the feed wins (Apple also publishes updates for older major versions), and the
    policy is only ever raised, never lowered.

    Sign-in goes through scripts\Startup\Connect-M365.ps1, so keep this folder inside
    the repository layout (..\..\Startup\Connect-M365.ps1 must exist).

    App-only (default for this unattended updater)
        Uses TenantId, ClientId and CertificateThumbprint from config.json (written by
        Setup.ps1). A ClientSecret in config.json is still accepted for configs made by
        older versions of Setup.ps1, but a certificate is preferred: a secret sits in
        plaintext in config.json.

    Delegated (manual run)
        Pass -CompliancePolicyId and no ClientId (none in config.json either): you sign
        in as an Intune admin, with device code when $global:useDeviceCodeAuth is set.

.PARAMETER ConfigPath
    Path to the configuration file (default: config.json in script directory).

.PARAMETER LogPath
    Path to the log directory (default: logs\ in script directory).

.PARAMETER CompliancePolicyId
    ID of the iOS compliance policy. Overrides CompliancePolicyId in config.json.

.PARAMETER TenantId
    Tenant ID or domain. Overrides TenantId in config.json.

.PARAMETER ClientId
    App registration for app-only sign-in. Overrides ClientId in config.json.

.PARAMETER CertificateThumbprint
    Certificate thumbprint for -ClientId. Overrides CertificateThumbprint in config.json.

.PARAMETER AppOnly
    App-only sign-in with ClientId and CertificateThumbprint from graph.appid.json.

.PARAMETER WhatIf
    Dry-run mode — shows what would happen without making changes (common parameter).

.EXAMPLE
    .\Update-iOSCompliancePolicy.ps1

.EXAMPLE
    .\Update-iOSCompliancePolicy.ps1 -WhatIf

.EXAMPLE
    .\Update-iOSCompliancePolicy.ps1 -ConfigPath "C:\Scripts\config.json"

.EXAMPLE
    # Manual, delegated run against one policy
    .\Update-iOSCompliancePolicy.ps1 -CompliancePolicyId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -WhatIf

.NOTES
    Graph permission: DeviceManagementConfiguration.ReadWrite.All (application for
    app-only, delegated scope for a manual run)
    Required module : Microsoft.Graph.Authentication
#>
[CmdletBinding(SupportsShouldProcess)]
param (
    [string] $ConfigPath = "$PSScriptRoot\config.json",
    [string] $LogPath    = "$PSScriptRoot\logs",
    [string] $CompliancePolicyId,
    [string] $TenantId,
    [string] $ClientId,
    [string] $CertificateThumbprint,
    [switch] $AppOnly
)

. (Join-Path $PSScriptRoot '..\..\Startup\Connect-M365.ps1')

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
        New-Item -ItemType Directory -Path $LogPath -Force -WhatIf:$false | Out-Null
    }

    $logFile = Join-Path $LogPath "compliance-updater-$(Get-Date -Format 'yyyy-MM').log"
    Add-Content -Path $logFile -Value $logMessage -WhatIf:$false
}

# ── Configuration ─────────────────────────────────────────────────────────────
function Get-Config {
    param ([string] $Path)

    if (-not (Test-Path $Path)) { return $null }
    try {
        $config = Get-Content $Path -Raw | ConvertFrom-Json
        Write-Log "Configuration loaded from: $Path"
        return $config
    } catch {
        Write-Log "Failed to parse config file: $_" -Level ERROR
        exit 1
    }
}

function Get-ConfigValue {
    # A parameter wins over config.json; 'FILL_IN' placeholders count as empty.
    param ([string] $Value, $Config, [string] $Name)
    if ($Value) { return $Value }
    if ($Config -and $Config.$Name -and $Config.$Name -ne 'FILL_IN') { return [string]$Config.$Name }
    return $null
}

# ── iOS version detection ─────────────────────────────────────────────────────
function Get-LatestIOSVersion {
    Write-Log "Fetching latest iOS version from Apple..."

    try {
        $rss = Invoke-RestMethod -Uri 'https://developer.apple.com/news/releases/rss/releases.rss' -ErrorAction Stop
        # Invoke-RestMethod returns the RSS <item> elements themselves, not the document; the
        # old $rss.channel.item.title was always empty, so every run fell through to the
        # fallback page (which no longer matched either).
        $items = if ($rss -is [xml]) { $rss.rss.channel.item } else { $rss }
        # Final releases only ("iOS 18.3.2 (22D82)", "iOS 18 (22A3354)"), not "iOS 18.4 beta 2"
        # or "iOS 18.4 RC". A bare major version becomes 18.0 for the comparison.
        $versions = @($items.title |
            Where-Object { $_ -match '^iOS \d+(\.\d+)*\s*(\(|$)' -and $_ -notmatch '\b(beta|RC)\b' } |
            ForEach-Object {
                $v = [regex]::Match($_, '^iOS (\d+(?:\.\d+)*)').Groups[1].Value
                if ($v -notmatch '\.') { $v = "$v.0" }
                [version]$v
            })

        if (-not $versions) { throw 'No iOS release entries found in Apple RSS feed.' }

        # The feed is ordered by date and also lists updates for older majors
        # (e.g. 17.7.x after 18.1), so take the highest version, not the first entry.
        $version = ($versions | Sort-Object -Descending | Select-Object -First 1).ToString()
        Write-Log "Latest iOS version: $version" -Level SUCCESS
        return $version
    } catch {
        Write-Log "Apple RSS failed: $($_.Exception.Message) — trying fallback..." -Level WARN

        try {
            # "About iOS updates": "The latest version of iOS and iPadOS is 18.3.2."
            $page  = Invoke-WebRequest -Uri 'https://support.apple.com/en-us/100100' -UseBasicParsing -ErrorAction Stop
            $match = [regex]::Match($page.Content, 'latest version of iOS[^.<]{0,40}?(\d+\.\d+(?:\.\d+)?)')
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
    param ([string] $PolicyId)

    Write-Log "Fetching compliance policy (ID: $PolicyId)..."

    try {
        $policy = Invoke-MgGraphRequest -Method GET `
            -Uri "https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies/$PolicyId" -ErrorAction Stop
        if ($policy.'@odata.type' -ne '#microsoft.graph.iosCompliancePolicy') {
            Write-Log "Policy '$($policy.displayName)' is a $($policy.'@odata.type'), not an iOS compliance policy." -Level ERROR
            exit 1
        }
        Write-Log "Policy: '$($policy.displayName)' — current minimum iOS: $($policy.osMinimumVersion)"
        return $policy
    } catch {
        Write-Log "Failed to fetch compliance policy: $($_.Exception.Message)" -Level ERROR
        exit 1
    }
}

function Set-CompliancePolicyVersion {
    param ([string] $PolicyId, [string] $NewVersion)

    Write-Log "Updating compliance policy to iOS $NewVersion..."

    # Intune wants the derived type on a PATCH of a compliance policy.
    $body = @{
        '@odata.type'    = '#microsoft.graph.iosCompliancePolicy'
        osMinimumVersion = $NewVersion
    } | ConvertTo-Json

    try {
        Invoke-MgGraphRequest -Method PATCH `
            -Uri "https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies/$PolicyId" `
            -Body $body -ContentType 'application/json' -ErrorAction Stop | Out-Null

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

if ($WhatIfPreference) { Write-Log "WHATIF mode — no changes will be made." -Level WARN }

$config = Get-Config -Path $ConfigPath
if (-not $config -and -not $CompliancePolicyId) {
    Write-Log "Config file not found: $ConfigPath" -Level ERROR
    Write-Log "Run Setup.ps1, or copy config.example.json to config.json and fill in your values." -Level ERROR
    exit 1
}

$policyId   = Get-ConfigValue $CompliancePolicyId    $config 'CompliancePolicyId'
$tenant     = Get-ConfigValue $TenantId              $config 'TenantId'
$appId      = Get-ConfigValue $ClientId              $config 'ClientId'
$thumbprint = Get-ConfigValue $CertificateThumbprint $config 'CertificateThumbprint'
$secretText = Get-ConfigValue ''                     $config 'ClientSecret'

if (-not $policyId) { Write-Log "No CompliancePolicyId (config.json or -CompliancePolicyId)." -Level ERROR; exit 1 }
if ($appId -and -not $tenant) { Write-Log "App-only sign-in needs TenantId in config.json or -TenantId." -Level ERROR; exit 1 }
if ($appId -and -not $thumbprint -and -not $secretText) {
    Write-Log "ClientId needs CertificateThumbprint (preferred) or ClientSecret in config.json." -Level ERROR; exit 1
}
if ($appId -and -not $thumbprint) {
    Write-Log "Signing in with the client secret from config.json. Re-run Setup.ps1 to switch to a certificate." -Level WARN
}

try {
    $connect = @{ Scopes = 'DeviceManagementConfiguration.ReadWrite.All'; TenantId = $tenant; AppOnly = $AppOnly }
    if ($appId) {
        $connect['ClientId'] = $appId
        if ($thumbprint) { $connect['CertificateThumbprint'] = $thumbprint }
        else { $connect['ClientSecret'] = ConvertTo-SecureString $secretText -AsPlainText -Force }
    }
    $graph = Connect-M365Graph @connect
    Write-Log "Connected to Microsoft Graph ($($graph.AuthType), tenant $($graph.TenantId))." -Level SUCCESS
} catch {
    Write-Log "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -Level ERROR
    exit 1
}

$latestVersion  = Get-LatestIOSVersion
$currentPolicy  = Get-CompliancePolicy -PolicyId $policyId
$currentVersion = [string]$currentPolicy.osMinimumVersion

$current = $null
if ($currentVersion) { [void][version]::TryParse($currentVersion, [ref]$current) }

if ($current -and $current -ge [version]$latestVersion) {
    Write-Log "Policy is up to date (minimum iOS $currentVersion, latest $latestVersion). No changes needed." -Level SUCCESS
} else {
    Write-Log "Update needed: $(if ($currentVersion) { $currentVersion } else { '(none)' }) -> $latestVersion"
    if ($PSCmdlet.ShouldProcess("compliance policy $policyId", "Set minimum iOS version to $latestVersion")) {
        Set-CompliancePolicyVersion -PolicyId $policyId -NewVersion $latestVersion
    } else {
        Write-Log "[WHATIF] Would update compliance policy to iOS $latestVersion" -Level WARN
    }
}

Disconnect-M365Graph $graph

Write-Log "========================================"
Write-Log " Script completed"
Write-Log "========================================"
