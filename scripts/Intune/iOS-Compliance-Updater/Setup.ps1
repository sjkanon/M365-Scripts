#Requires -Version 7.0
<#
.SYNOPSIS
    One-time setup for the Intune iOS Compliance Updater.

.DESCRIPTION
    Automates the full setup:
    1. Installs required PowerShell modules (Microsoft.Graph.Authentication, .Applications)
    2. Signs you in (delegated) as an administrator
    3. Creates (or reuses) an App Registration in Entra ID
    4. Assigns the Graph application permission DeviceManagementConfiguration.ReadWrite.All
       and grants admin consent
    5. Creates the app's credential: a self-signed certificate (default) whose public
       key is uploaded to the app, or - with -CredentialType Secret - a client secret
    6. Looks up the Intune iOS compliance policy (all pages)
    7. Writes config.json

    Sign-in goes through scripts\Startup\Connect-M365.ps1: delegated (you, as an
    admin; device code when $global:useDeviceCodeAuth is set; under GDAP the customer
    tenant from $global:cid unless -TenantId names one). The updater itself then runs
    app-only with the credential created here.

    Certificate (default): the private key stays in the Windows certificate store and
    config.json holds only its thumbprint. When this setup runs elevated the
    certificate goes to LocalMachine\My, so the scheduled task (SYSTEM) can use it;
    otherwise to CurrentUser\My - then run the task as that user.
    Secret: config.json holds the secret in plaintext. Anyone who can read the file
    can change Intune configuration in the tenant. Use it only where a certificate
    is not possible, and protect the file.

.PARAMETER TenantId
    Tenant ID or domain to set up. Defaults to the GDAP customer tenant, else the tenant
    you sign in to.

.PARAMETER CompliancePolicyName
    Name of the Intune compliance policy to update.
    If omitted, the script shows a list of available iOS policies to choose from.

.PARAMETER AppName
    Name of the App Registration (default: "Intune iOS Compliance Updater").

.PARAMETER CredentialType
    Certificate (default) or Secret.

.PARAMETER CertificateStoreLocation
    LocalMachine or CurrentUser. Default: LocalMachine when running elevated, else CurrentUser.

.PARAMETER CertificateExpiryYears
    Validity of the self-signed certificate in years (default: 2).

.PARAMETER SecretExpiryYears
    Validity of the client secret in years (default: 2). Only with -CredentialType Secret.

.PARAMETER ConfigPath
    Where to save config.json (default: script directory).

.EXAMPLE
    # Run elevated so the certificate lands in LocalMachine\My for the SYSTEM task
    .\Setup.ps1

.EXAMPLE
    .\Setup.ps1 -TenantId "contoso.onmicrosoft.com" -CompliancePolicyName "iOS - Minimum version compliance"

.EXAMPLE
    # Old behaviour: a client secret in config.json
    .\Setup.ps1 -CredentialType Secret

.NOTES
    Required role: Global Administrator, or Application Administrator + Privileged Role
    Administrator (for the admin consent) + Intune Administrator
    Scopes: Application.ReadWrite.All, AppRoleAssignment.ReadWrite.All,
            DeviceManagementConfiguration.Read.All
    Windows only (New-SelfSignedCertificate, certificate store).
#>
[CmdletBinding()]
param (
    [string] $TenantId,
    [string] $CompliancePolicyName = '',
    [string] $AppName              = 'Intune iOS Compliance Updater',
    [ValidateSet('Certificate', 'Secret')]
    [string] $CredentialType       = 'Certificate',
    [ValidateSet('LocalMachine', 'CurrentUser')]
    [string] $CertificateStoreLocation,
    [int]    $CertificateExpiryYears = 2,
    [int]    $SecretExpiryYears    = 2,
    [string] $ConfigPath           = "$PSScriptRoot\config.json"
)

. (Join-Path $PSScriptRoot '..\..\Startup\Connect-M365.ps1')

# ── Helpers ───────────────────────────────────────────────────────────────────
function Write-Step { param([string]$m) Write-Host "`n  === $m ===" -ForegroundColor Magenta }
function Write-Ok   { param([string]$m) Write-Host "  [OK]   $m" -ForegroundColor Green }
function Write-Warn { param([string]$m) Write-Host "  [WARN] $m" -ForegroundColor Yellow }
function Write-Err  { param([string]$m) Write-Host "  [ERR]  $m" -ForegroundColor Red }
function Write-Info { param([string]$m) Write-Host "  [INFO] $m" -ForegroundColor Cyan }

$isAdmin = $IsWindows -and ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $CertificateStoreLocation) { $CertificateStoreLocation = if ($isAdmin) { 'LocalMachine' } else { 'CurrentUser' } }

# ── Prerequisites ─────────────────────────────────────────────────────────────
function Test-Prerequisites {
    Write-Step "Checking prerequisites"

    $allOk = $true

    Write-Ok "PowerShell $($PSVersionTable.PSVersion)"

    if (-not $IsWindows) {
        Write-Err "Windows is required (certificate store, scheduled task)."
        $allOk = $false
    }

    if ($isAdmin) {
        Write-Ok "Running as Administrator"
    } else {
        Write-Warn "Not running as Administrator — modules install for CurrentUser and the certificate goes to CurrentUser\My"
    }

    if ($CredentialType -eq 'Certificate' -and $CertificateStoreLocation -eq 'LocalMachine' -and -not $isAdmin) {
        Write-Err "-CertificateStoreLocation LocalMachine needs an elevated session."
        $allOk = $false
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
            # Graph answers its root with 4xx; any HTTP response means it is reachable.
            if ($_.Exception.Response) { Write-Ok "Reachable: $($_.Name)" }
            else { Write-Err "Not reachable: $($_.Name) ($($_.Url))"; $allOk = $false }
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

    $scope = if ($isAdmin) { 'AllUsers' } else { 'CurrentUser' }
    Write-Info "Install scope: $scope"

    @(
        @{ Name = 'Microsoft.Graph.Authentication'; MinVersion = '2.0.0' }
        @{ Name = 'Microsoft.Graph.Applications';   MinVersion = '2.0.0' }
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
    Write-Info "Required role: Global Administrator, or Application Administrator + Privileged Role Administrator + Intune Administrator"

    try {
        $script:graph = Connect-M365Graph -TenantId $TenantId -Scopes @(
            'Application.ReadWrite.All'
            'AppRoleAssignment.ReadWrite.All'
            'DeviceManagementConfiguration.Read.All'
        )
        Write-Ok "Signed in as: $($script:graph.Account)"
        Write-Ok "Tenant: $($script:graph.TenantId)"
        return $script:graph.TenantId
    } catch {
        Write-Err "Authentication failed: $($_.Exception.Message)"; exit 1
    }
}

# ── App Registration ──────────────────────────────────────────────────────────
function New-AppRegistration {
    param ([string] $Name)
    Write-Step "Creating App Registration"

    $existing = Get-MgApplication -Filter "displayName eq '$($Name -replace "'", "''")'" -ErrorAction SilentlyContinue | Select-Object -First 1
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

# ── Credential: certificate (default) ─────────────────────────────────────────
function New-AppCertificate {
    param ([string] $AppObjectId, [string] $Name, [int] $ExpiryYears, [string] $StoreLocation)
    Write-Step "Creating certificate ($StoreLocation\My)"

    try {
        $cert = New-SelfSignedCertificate `
            -Subject           "CN=$Name" `
            -CertStoreLocation "Cert:\$StoreLocation\My" `
            -KeyExportPolicy   NonExportable `
            -KeySpec           Signature `
            -KeyAlgorithm      RSA `
            -KeyLength         2048 `
            -HashAlgorithm     SHA256 `
            -NotAfter          (Get-Date).AddYears($ExpiryYears) `
            -ErrorAction Stop
        Write-Ok "Certificate created: $($cert.Thumbprint) (valid until $($cert.NotAfter.ToString('yyyy-MM-dd')))"
    } catch {
        Write-Err "Could not create certificate: $($_.Exception.Message)"; exit 1
    }

    try {
        # keyCredentials is replaced as a whole on PATCH. The existing keys (with their
        # public key bytes, which Graph only returns on $select) are sent back so a
        # reused app keeps working on other machines.
        $current = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/applications/$AppObjectId`?`$select=keyCredentials" -ErrorAction Stop
        $keys = @($current.keyCredentials | Where-Object { $_.key } | ForEach-Object {
            @{ keyId = $_.keyId; type = $_.type; usage = $_.usage; key = $_.key; displayName = $_.displayName }
        })
        $keys += @{
            type        = 'AsymmetricX509Cert'
            usage       = 'Verify'
            key         = [Convert]::ToBase64String($cert.RawData)
            displayName = "iOS Compliance Updater - $env:COMPUTERNAME - $(Get-Date -Format 'yyyy-MM-dd')"
        }
        Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/applications/$AppObjectId" `
            -Body (@{ keyCredentials = $keys } | ConvertTo-Json -Depth 5) -ContentType 'application/json' -ErrorAction Stop | Out-Null
        Write-Ok "Certificate uploaded to the App Registration"
        return $cert.Thumbprint
    } catch {
        Write-Err "Could not upload the certificate: $($_.Exception.Message)"
        Write-Info "Upload the .cer manually in Entra ID > App registrations > Certificates & secrets."
        exit 1
    }
}

# ── Credential: client secret (-CredentialType Secret) ────────────────────────
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
        Write-Warn "The secret is written to config.json in plaintext — protect that file."
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
        $all = [System.Collections.Generic.List[object]]::new()
        $uri = 'https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies'
        while ($uri) {
            $page = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
            foreach ($p in @($page.value)) { $all.Add($p) }
            $uri = $page.'@odata.nextLink'
        }
        $iosPolicies = @($all | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.iosCompliancePolicy' })

        if (-not $iosPolicies) {
            Write-Err "No iOS compliance policies found in Intune"; exit 1
        }

        if ($PolicyName) {
            $match = $iosPolicies | Where-Object { $_.displayName -eq $PolicyName } | Select-Object -First 1
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
    param ([string] $TenantId, [string] $ClientId, [string] $CertificateThumbprint, [string] $ClientSecret, [string] $PolicyId, [string] $Path)
    Write-Step "Writing config.json"

    $values = [ordered]@{ TenantId = $TenantId; ClientId = $ClientId }
    if ($CertificateThumbprint) { $values['CertificateThumbprint'] = $CertificateThumbprint }
    if ($ClientSecret)          { $values['ClientSecret'] = $ClientSecret }
    $values['CompliancePolicyId'] = $PolicyId
    $config = $values | ConvertTo-Json -Depth 3

    try {
        $config | Set-Content -Path $Path -Encoding UTF8 -ErrorAction Stop
        Write-Ok "config.json saved to: $Path"
        if ($ClientSecret) { Write-Warn "Never commit config.json to Git — it contains a secret!" }
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

$tenantId = Connect-ToGraph
$app      = New-AppRegistration -Name $AppName
$sp       = New-AppServicePrincipal -AppId $app.AppId
Set-GraphPermissions -AppObjectId $app.Id -ServicePrincipalId $sp.Id

$thumbprint   = $null
$clientSecret = $null
if ($CredentialType -eq 'Certificate') {
    $thumbprint = New-AppCertificate -AppObjectId $app.Id -Name $AppName -ExpiryYears $CertificateExpiryYears -StoreLocation $CertificateStoreLocation
} else {
    $clientSecret = New-AppClientSecret -AppObjectId $app.Id -ExpiryYears $SecretExpiryYears
}

$policyId = Get-CompliancePolicyId -PolicyName $CompliancePolicyName
Save-Config -TenantId $tenantId -ClientId $app.AppId -CertificateThumbprint $thumbprint -ClientSecret $clientSecret -PolicyId $policyId -Path $ConfigPath

Disconnect-M365Graph $script:graph

Write-Host ""
Write-Host "  ================================================" -ForegroundColor Green
Write-Host "   Setup complete!" -ForegroundColor Green
Write-Host "  ================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Next steps:" -ForegroundColor White
Write-Host "  1. Test: .\Update-iOSCompliancePolicy.ps1 -WhatIf" -ForegroundColor Cyan
Write-Host "     (new app permissions can take a few minutes to apply)" -ForegroundColor DarkGray
Write-Host "  2. Register scheduled task: .\Install-ScheduledTask.ps1 (elevated)" -ForegroundColor Cyan
if ($thumbprint -and $CertificateStoreLocation -eq 'CurrentUser') {
    Write-Warn "The certificate is in CurrentUser\My: the SYSTEM task cannot use it. Re-run Setup.ps1 elevated, or run the task as this user with a stored password."
}
Write-Host "  3. Check logs in the logs\ folder" -ForegroundColor Cyan
Write-Host ""
