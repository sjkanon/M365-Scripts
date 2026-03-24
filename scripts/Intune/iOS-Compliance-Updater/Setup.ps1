#Requires -Version 5.1
<#
.SYNOPSIS
    One-time setup for the Intune iOS Compliance Updater.

.DESCRIPTION
    Automates the full setup:
    1. Installs required PowerShell modules (Microsoft.Graph)
    2. Creates an App Registration in Entra ID
    3. Assigns the required Graph API permissions
    4. Grants admin consent
    5. Creates a Client Secret
    6. Looks up the Intune iOS compliance policy
    7. Writes config.json

.PARAMETER CompliancePolicyName
    Name of the Intune compliance policy to update.
    If omitted, the script shows a list of available iOS policies to choose from.

.PARAMETER AppName
    Name of the App Registration (default: "Intune iOS Compliance Updater").

.PARAMETER SecretExpiryYears
    Validity of the client secret in years (default: 2).

.PARAMETER ConfigPath
    Where to save config.json (default: script directory).

.EXAMPLE
    .\Setup.ps1

.EXAMPLE
    .\Setup.ps1 -CompliancePolicyName "iOS - Minimum version compliance"

.NOTES
    Required role: Global Administrator or Application Administrator + Intune Administrator
#>
[CmdletBinding()]
param (
    [string] $CompliancePolicyName = '',
    [string] $AppName              = 'Intune iOS Compliance Updater',
    [int]    $SecretExpiryYears    = 2,
    [string] $ConfigPath           = "$PSScriptRoot\config.json"
)

# ── Helpers ───────────────────────────────────────────────────────────────────
function Write-Step { param([string]$m) Write-Host "`n  === $m ===" -ForegroundColor Magenta }
function Write-Ok   { param([string]$m) Write-Host "  [OK]   $m" -ForegroundColor Green }
function Write-Warn { param([string]$m) Write-Host "  [WARN] $m" -ForegroundColor Yellow }
function Write-Err  { param([string]$m) Write-Host "  [ERR]  $m" -ForegroundColor Red }
function Write-Info { param([string]$m) Write-Host "  [INFO] $m" -ForegroundColor Cyan }

# ── Prerequisites ─────────────────────────────────────────────────────────────
function Test-Prerequisites {
    Write-Step "Checking prerequisites"

    $allOk = $true

    # PowerShell version
    $psv = $PSVersionTable.PSVersion
    if ($psv.Major -gt 5 -or ($psv.Major -eq 5 -and $psv.Minor -ge 1)) {
        Write-Ok "PowerShell $($psv)"
    } else {
        Write-Err "PowerShell 5.1 or later required (current: $psv)"
        $allOk = $false
    }

    # Administrator check
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)
    if ($isAdmin) {
        Write-Ok "Running as Administrator"
    } else {
        Write-Warn "Not running as Administrator — modules will be installed for CurrentUser only"
    }

    # Execution policy
    $execPolicy = Get-ExecutionPolicy -Scope CurrentUser
    if ($execPolicy -in @('Unrestricted', 'RemoteSigned', 'Bypass')) {
        Write-Ok "Execution policy: $execPolicy"
    } else {
        Write-Warn "Execution policy is '$execPolicy' — adjusting to RemoteSigned..."
        try {
            Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force -ErrorAction Stop
            Write-Ok "Execution policy set to RemoteSigned (CurrentUser)"
        } catch {
            Write-Err "Could not set execution policy: $($_.Exception.Message)"
            Write-Info "Run manually: Set-ExecutionPolicy RemoteSigned -Scope CurrentUser"
            $allOk = $false
        }
    }

    # Connectivity
    Write-Info "Testing connectivity..."
    @(
        @{ Name = 'Microsoft Graph';    Url = 'https://graph.microsoft.com' }
        @{ Name = 'Microsoft Login';    Url = 'https://login.microsoftonline.com' }
        @{ Name = 'PowerShell Gallery'; Url = 'https://www.powershellgallery.com' }
        @{ Name = 'Apple RSS';          Url = 'https://developer.apple.com' }
    ) | ForEach-Object {
        try {
            $null = Invoke-WebRequest -Uri $_.Url -UseBasicParsing -TimeoutSec 5 -ErrorAction Stop
            Write-Ok "Reachable: $($_.Name)"
        } catch {
            Write-Err "Not reachable: $($_.Name) ($($_.Url))"
            $allOk = $false
        }
    }

    if (-not $allOk) {
        Write-Host "`n  One or more prerequisites failed. Fix the issues above and retry." -ForegroundColor Red
        exit 1
    }
}

# ── Module installation ───────────────────────────────────────────────────────
function Install-RequiredModules {
    Write-Step "Installing PowerShell modules"

    # NuGet provider
    $nuget = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
    if (-not $nuget -or $nuget.Version -lt [Version]'2.8.5.201') {
        Write-Info "Installing NuGet package provider..."
        try {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -ErrorAction Stop | Out-Null
            Write-Ok "NuGet provider installed"
        } catch {
            Write-Err "Could not install NuGet provider: $($_.Exception.Message)"; exit 1
        }
    } else {
        Write-Ok "NuGet provider present (v$($nuget.Version))"
    }

    # Trust PSGallery
    if ((Get-PSRepository -Name PSGallery).InstallationPolicy -ne 'Trusted') {
        try {
            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction Stop
            Write-Ok "PSGallery set as trusted"
        } catch {
            Write-Warn "Could not trust PSGallery: $($_.Exception.Message)"
        }
    } else {
        Write-Ok "PSGallery already trusted"
    }

    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)
    $scope   = if ($isAdmin) { 'AllUsers' } else { 'CurrentUser' }
    Write-Info "Install scope: $scope"

    @(
        @{ Name = 'Microsoft.Graph.Authentication';    MinVersion = '2.0.0' }
        @{ Name = 'Microsoft.Graph.Applications';      MinVersion = '2.0.0' }
        @{ Name = 'Microsoft.Graph.DeviceManagement';  MinVersion = '2.0.0' }
    ) | ForEach-Object {
        $mod       = $_
        $installed = Get-Module -ListAvailable -Name $mod.Name | Sort-Object Version -Descending | Select-Object -First 1

        if ($installed -and $installed.Version -ge [Version]$mod.MinVersion) {
            Write-Ok "$($mod.Name) v$($installed.Version)"
        } else {
            Write-Info "Installing $($mod.Name)..."
            try {
                Install-Module -Name $mod.Name -Scope $scope -Force -AllowClobber -Repository PSGallery -ErrorAction Stop
                $ver = (Get-Module -ListAvailable -Name $mod.Name | Sort-Object Version -Descending | Select-Object -First 1).Version
                Write-Ok "$($mod.Name) v$ver installed"
            } catch {
                Write-Err "Could not install $($mod.Name): $($_.Exception.Message)"; exit 1
            }
        }

        Import-Module $mod.Name -ErrorAction SilentlyContinue
    }

    Write-Ok "All modules ready"
}

# ── Authentication ────────────────────────────────────────────────────────────
function Connect-ToGraph {
    Write-Step "Connecting to Microsoft Graph"
    Write-Info "A browser window will open for authentication."
    Write-Info "Required role: Global Administrator or Application Administrator + Intune Administrator"

    try {
        Connect-MgGraph -Scopes @(
            'Application.ReadWrite.All'
            'AppRoleAssignment.ReadWrite.All'
            'Directory.ReadWrite.All'
            'DeviceManagementConfiguration.ReadWrite.All'
        ) -NoWelcome -ErrorAction Stop

        $ctx = Get-MgContext
        Write-Ok "Signed in as: $($ctx.Account)"
        Write-Ok "Tenant: $($ctx.TenantId)"
        return $ctx.TenantId
    } catch {
        Write-Err "Authentication failed: $($_.Exception.Message)"; exit 1
    }
}

# ── App Registration ──────────────────────────────────────────────────────────
function New-AppRegistration {
    param ([string] $Name)
    Write-Step "Creating App Registration"

    $existing = Get-MgApplication -Filter "displayName eq '$Name'" -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Warn "App Registration '$Name' already exists (ID: $($existing.AppId))"
        $choice = Read-Host "  Reuse existing app? (y/n)"
        if ($choice -eq 'y') {
            Write-Ok "Reusing existing App Registration"
            return $existing
        }
        $Name = "$Name $(Get-Date -Format 'yyyyMMdd-HHmm')"
        Write-Info "Creating new app as: $Name"
    }

    try {
        $app = New-MgApplication -DisplayName $Name -ErrorAction Stop
        Write-Ok "App Registration created: '$Name'"
        Write-Ok "Client ID: $($app.AppId)"
        return $app
    } catch {
        Write-Err "Could not create App Registration: $($_.Exception.Message)"; exit 1
    }
}

# ── Service Principal ─────────────────────────────────────────────────────────
function New-AppServicePrincipal {
    param ([string] $AppId)
    Write-Step "Creating Service Principal"

    $existing = Get-MgServicePrincipal -Filter "appId eq '$AppId'" -ErrorAction SilentlyContinue
    if ($existing) { Write-Ok "Service Principal already exists"; return $existing }

    try {
        $sp = New-MgServicePrincipal -AppId $AppId -ErrorAction Stop
        Write-Ok "Service Principal created (ID: $($sp.Id))"
        return $sp
    } catch {
        Write-Err "Could not create Service Principal: $($_.Exception.Message)"; exit 1
    }
}

# ── API permissions ───────────────────────────────────────────────────────────
function Set-GraphPermissions {
    param ([string] $AppObjectId, [string] $ServicePrincipalId)
    Write-Step "Assigning API permissions"

    $graphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'" -ErrorAction Stop
    $appRole = $graphSp.AppRoles | Where-Object { $_.Value -eq 'DeviceManagementConfiguration.ReadWrite.All' }

    if (-not $appRole) {
        Write-Err "Permission 'DeviceManagementConfiguration.ReadWrite.All' not found in Microsoft Graph"; exit 1
    }

    $existing = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ServicePrincipalId -ErrorAction SilentlyContinue |
                Where-Object { $_.AppRoleId -eq $appRole.Id }

    if ($existing) {
        Write-Ok "Permission already assigned"
    } else {
        try {
            New-MgServicePrincipalAppRoleAssignment `
                -ServicePrincipalId $ServicePrincipalId `
                -PrincipalId        $ServicePrincipalId `
                -ResourceId         $graphSp.Id `
                -AppRoleId          $appRole.Id `
                -ErrorAction Stop | Out-Null
            Write-Ok "Permission assigned + admin consent granted"
        } catch {
            Write-Err "Could not assign permission: $($_.Exception.Message)"
            Write-Warn "Assign manually in Entra ID > App registrations > API permissions"
        }
    }
}

# ── Client Secret ─────────────────────────────────────────────────────────────
function New-AppClientSecret {
    param ([string] $AppObjectId, [int] $ExpiryYears)
    Write-Step "Creating Client Secret"

    $endDate = (Get-Date).AddYears($ExpiryYears)

    try {
        $secret = Add-MgApplicationPassword `
            -ApplicationId       $AppObjectId `
            -PasswordCredential  @{
                displayName = "iOS Compliance Updater - $(Get-Date -Format 'yyyy-MM-dd')"
                endDateTime = $endDate
            } `
            -ErrorAction Stop

        Write-Ok "Client Secret created (valid until: $($endDate.ToString('yyyy-MM-dd')))"
        Write-Warn "Store this value securely — it is only shown once!"
        Write-Host "  Secret: $($secret.SecretText)" -ForegroundColor Yellow
        return $secret.SecretText
    } catch {
        Write-Err "Could not create Client Secret: $($_.Exception.Message)"; exit 1
    }
}

# ── Compliance policy lookup ──────────────────────────────────────────────────
function Get-CompliancePolicyId {
    param ([string] $PolicyName)
    Write-Step "Looking up Intune compliance policy"

    try {
        $policies    = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies' -ErrorAction Stop
        $iosPolicies = $policies.value | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.iosCompliancePolicy' }

        if (-not $iosPolicies) {
            Write-Err "No iOS compliance policies found in Intune"; exit 1
        }

        if ($PolicyName) {
            $match = $iosPolicies | Where-Object { $_.displayName -eq $PolicyName }
            if ($match) {
                Write-Ok "Policy found: '$($match.displayName)' (ID: $($match.id))"
                return $match.id
            }
            Write-Warn "Policy '$PolicyName' not found. Available iOS policies:"
        } else {
            Write-Info "Available iOS compliance policies:"
        }

        $i = 1
        $iosPolicies | ForEach-Object { Write-Host "  [$i] $($_.displayName)" -ForegroundColor White; $i++ }

        $choice = [int](Read-Host "`n  Choose a policy (number)") - 1
        if ($choice -lt 0 -or $choice -ge $iosPolicies.Count) {
            Write-Err "Invalid choice"; exit 1
        }

        $selected = $iosPolicies[$choice]
        Write-Ok "Selected: '$($selected.displayName)' (ID: $($selected.id))"
        return $selected.id
    } catch {
        Write-Err "Could not retrieve compliance policies: $($_.Exception.Message)"; exit 1
    }
}

# ── Write config ──────────────────────────────────────────────────────────────
function Save-Config {
    param ([string] $TenantId, [string] $ClientId, [string] $ClientSecret, [string] $PolicyId, [string] $Path)
    Write-Step "Writing config.json"

    $config = @{
        TenantId           = $TenantId
        ClientId           = $ClientId
        ClientSecret       = $ClientSecret
        CompliancePolicyId = $PolicyId
    } | ConvertTo-Json -Depth 3

    try {
        $config | Set-Content -Path $Path -Encoding UTF8 -ErrorAction Stop
        Write-Ok "config.json saved to: $Path"
        Write-Warn "Never commit config.json to Git — it contains secrets!"
    } catch {
        Write-Err "Could not write config.json: $($_.Exception.Message)"
        Write-Info "Create config.json manually with these values:"
        Write-Host $config -ForegroundColor Yellow
    }
}

# ── Main ──────────────────────────────────────────────────────────────────────
Clear-Host
Write-Host ""
Write-Host "  ================================================" -ForegroundColor Magenta
Write-Host "   Intune iOS Compliance Updater - Setup" -ForegroundColor Magenta
Write-Host "  ================================================" -ForegroundColor Magenta
Write-Host ""

Test-Prerequisites
Install-RequiredModules

$tenantId     = Connect-ToGraph
$app          = New-AppRegistration -Name $AppName
$sp           = New-AppServicePrincipal -AppId $app.AppId
Set-GraphPermissions -AppObjectId $app.Id -ServicePrincipalId $sp.Id
$clientSecret = New-AppClientSecret -AppObjectId $app.Id -ExpiryYears $SecretExpiryYears
$policyId     = Get-CompliancePolicyId -PolicyName $CompliancePolicyName
Save-Config -TenantId $tenantId -ClientId $app.AppId -ClientSecret $clientSecret -PolicyId $policyId -Path $ConfigPath

Write-Host ""
Write-Host "  ================================================" -ForegroundColor Green
Write-Host "   Setup complete!" -ForegroundColor Green
Write-Host "  ================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Next steps:" -ForegroundColor White
Write-Host "  1. Test: .\Update-iOSCompliancePolicy.ps1 -WhatIf" -ForegroundColor Cyan
Write-Host "  2. Register scheduled task: .\Install-ScheduledTask.ps1" -ForegroundColor Cyan
Write-Host "  3. Check logs in the logs\ folder" -ForegroundColor Cyan
Write-Host ""
