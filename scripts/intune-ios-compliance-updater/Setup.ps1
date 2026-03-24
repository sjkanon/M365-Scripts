#Requires -Version 5.1
<#
.SYNOPSIS
    One-time setup script voor de Intune iOS Compliance Updater.

.DESCRIPTION
    Dit script doet het volgende automatisch:
    1. Installeert benodigde PowerShell modules (Microsoft.Graph)
    2. Maakt een App Registration aan in Entra ID
    3. Wijst de juiste Graph API permissies toe
    4. Geeft admin consent
    5. Maakt een Client Secret aan
    6. Zoekt de Intune compliance policy op
    7. Schrijft automatisch config.json

.PARAMETER CompliancePolicyName
    Naam van de Intune compliance policy om te updaten.
    Als leeg, toont het script een lijst van beschikbare policies.

.PARAMETER AppName
    Naam van de App Registration (default: "Intune iOS Compliance Updater")

.PARAMETER SecretExpiryYears
    Hoe lang het client secret geldig is in jaren (default: 2)

.PARAMETER ConfigPath
    Waar config.json wordt opgeslagen (default: script directory)

.EXAMPLE
    .\Setup.ps1
    .\Setup.ps1 -CompliancePolicyName "iOS - Minimum versie compliance"
    .\Setup.ps1 -AppName "BraveHub iOS Updater" -SecretExpiryYears 1

.NOTES
    Author:  BraveHub
    Version: 1.0.0
    Vereist: Global Administrator of Application Administrator + Intune Administrator rol
#>

[CmdletBinding()]
param (
    [string]$CompliancePolicyName = "",
    [string]$AppName              = "Intune iOS Compliance Updater",
    [int]$SecretExpiryYears       = 2,
    [string]$ConfigPath           = "$PSScriptRoot\config.json"
)

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────

function Write-Step {
    param ([string]$Message)
    Write-Host "`n━━━ $Message ━━━" -ForegroundColor Magenta
}

function Write-Ok   { param([string]$m) Write-Host "  ✓ $m" -ForegroundColor Green }
function Write-Warn { param([string]$m) Write-Host "  ⚠ $m" -ForegroundColor Yellow }
function Write-Err  { param([string]$m) Write-Host "  ✗ $m" -ForegroundColor Red }
function Write-Info { param([string]$m) Write-Host "  → $m" -ForegroundColor Cyan }

# ─────────────────────────────────────────────
# SYSTEEM CHECKS & MODULE INSTALLATIE
# ─────────────────────────────────────────────

function Test-Prerequisites {
    Write-Step "Systeem vereisten controleren"

    $allOk = $true

    # 1. PowerShell versie
    $psVersion = $PSVersionTable.PSVersion
    if ($psVersion.Major -ge 5 -and $psVersion.Minor -ge 1) {
        Write-Ok "PowerShell versie: $($psVersion.ToString())"
    }
    else {
        Write-Err "PowerShell 5.1 of hoger vereist. Huidige versie: $($psVersion.ToString())"
        Write-Info "Download: https://github.com/PowerShell/PowerShell/releases"
        $allOk = $false
    }

    # 2. Wordt uitgevoerd als Administrator?
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if ($isAdmin) {
        Write-Ok "Script draait als Administrator"
    }
    else {
        Write-Warn "Script draait NIET als Administrator"
        Write-Info "Module installatie voor AllUsers vereist Administrator."
        Write-Info "Modules worden geïnstalleerd voor huidige gebruiker (CurrentUser scope)."
    }

    # 3. Execution Policy
    $execPolicy = Get-ExecutionPolicy -Scope CurrentUser
    if ($execPolicy -in @("Unrestricted", "RemoteSigned", "Bypass")) {
        Write-Ok "Execution Policy (CurrentUser): $execPolicy"
    }
    else {
        Write-Warn "Execution Policy staat op '$execPolicy' — wordt tijdelijk aangepast..."
        try {
            Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force -ErrorAction Stop
            Write-Ok "Execution Policy aangepast naar RemoteSigned (CurrentUser)"
        }
        catch {
            Write-Err "Kon Execution Policy niet aanpassen: $($_.Exception.Message)"
            Write-Info "Voer handmatig uit: Set-ExecutionPolicy RemoteSigned -Scope CurrentUser"
            $allOk = $false
        }
    }

    # 4. Internetverbinding
    Write-Info "Internetverbinding testen..."
    $endpoints = @(
        @{ Name = "Microsoft Graph";    Url = "https://graph.microsoft.com" },
        @{ Name = "Microsoft Login";    Url = "https://login.microsoftonline.com" },
        @{ Name = "PowerShell Gallery"; Url = "https://www.powershellgallery.com" },
        @{ Name = "Apple RSS";          Url = "https://developer.apple.com" }
    )

    foreach ($ep in $endpoints) {
        try {
            $null = Invoke-WebRequest -Uri $ep.Url -UseBasicParsing -TimeoutSec 5 -ErrorAction Stop
            Write-Ok "Bereikbaar: $($ep.Name)"
        }
        catch {
            Write-Err "Niet bereikbaar: $($ep.Name) ($($ep.Url))"
            Write-Info "Controleer firewall/proxy instellingen"
            $allOk = $false
        }
    }

    if (-not $allOk) {
        Write-Host "`n  Een of meer vereisten zijn niet voldaan. Los bovenstaande problemen op en probeer opnieuw." -ForegroundColor Red
        exit 1
    }

    Write-Ok "Alle systeem vereisten OK"
}

function Install-RequiredModules {
    Write-Step "PowerShell modules installeren"

    # Zorg dat NuGet provider beschikbaar is (vereist voor Install-Module)
    $nuget = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
    if (-not $nuget -or $nuget.Version -lt [Version]"2.8.5.201") {
        Write-Info "NuGet package provider installeren..."
        try {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -ErrorAction Stop | Out-Null
            Write-Ok "NuGet provider geïnstalleerd"
        }
        catch {
            Write-Err "Kon NuGet provider niet installeren: $($_.Exception.Message)"
            exit 1
        }
    }
    else {
        Write-Ok "NuGet provider aanwezig (v$($nuget.Version))"
    }

    # Vertrouw PSGallery als dat nog niet het geval is
    $gallery = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
    if ($gallery.InstallationPolicy -ne "Trusted") {
        Write-Info "PSGallery als trusted repository instellen..."
        try {
            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction Stop
            Write-Ok "PSGallery ingesteld als trusted"
        }
        catch {
            Write-Warn "Kon PSGallery niet als trusted instellen: $($_.Exception.Message)"
            Write-Info "Modules worden geïnstalleerd met -Force parameter"
        }
    }
    else {
        Write-Ok "PSGallery is al trusted"
    }

    # Benodigde modules
    $modules = @(
        @{ Name = "Microsoft.Graph.Authentication"; MinVersion = "2.0.0" },
        @{ Name = "Microsoft.Graph.Applications";   MinVersion = "2.0.0" },
        @{ Name = "Microsoft.Graph.DeviceManagement"; MinVersion = "2.0.0" }
    )

    $scope = if (([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        "AllUsers"
    } else {
        "CurrentUser"
    }

    Write-Info "Installatie scope: $scope"

    foreach ($mod in $modules) {
        $installed = Get-Module -ListAvailable -Name $mod.Name |
                     Sort-Object Version -Descending |
                     Select-Object -First 1

        if ($installed -and $installed.Version -ge [Version]$mod.MinVersion) {
            Write-Ok "$($mod.Name) v$($installed.Version) — al geïnstalleerd"
        }
        else {
            if ($installed) {
                Write-Warn "$($mod.Name) v$($installed.Version) gevonden maar verouderd (min: $($mod.MinVersion)) — updaten..."
            }
            else {
                Write-Info "$($mod.Name) installeren..."
            }

            try {
                Install-Module `
                    -Name         $mod.Name `
                    -Scope        $scope `
                    -Force        `
                    -AllowClobber `
                    -Repository   PSGallery `
                    -ErrorAction  Stop

                $newVersion = (Get-Module -ListAvailable -Name $mod.Name | Sort-Object Version -Descending | Select-Object -First 1).Version
                Write-Ok "$($mod.Name) v$newVersion geïnstalleerd"
            }
            catch {
                Write-Err "Kon $($mod.Name) niet installeren: $($_.Exception.Message)"
                Write-Info "Probeer handmatig: Install-Module $($mod.Name) -Scope CurrentUser -Force"
                exit 1
            }
        }

        # Module importeren
        try {
            Import-Module $mod.Name -ErrorAction Stop
        }
        catch {
            Write-Warn "Kon $($mod.Name) niet importeren: $($_.Exception.Message)"
        }
    }

    Write-Ok "Alle modules klaar"
}

# ─────────────────────────────────────────────
# AUTHENTICATIE
# ─────────────────────────────────────────────

function Connect-ToGraph {
    Write-Step "Aanmelden bij Microsoft Graph"
    Write-Info "Er opent een browservenster om aan te melden..."
    Write-Info "Vereiste rol: Global Administrator of Application Administrator + Intune Administrator"

    try {
        Connect-MgGraph -Scopes @(
            "Application.ReadWrite.All",
            "AppRoleAssignment.ReadWrite.All",
            "Directory.ReadWrite.All",
            "DeviceManagementConfiguration.ReadWrite.All"
        ) -ErrorAction Stop

        $context = Get-MgContext
        Write-Ok "Aangemeld als: $($context.Account)"
        Write-Ok "Tenant: $($context.TenantId)"
        return $context.TenantId
    }
    catch {
        Write-Err "Aanmelden mislukt: $($_.Exception.Message)"
        exit 1
    }
}

# ─────────────────────────────────────────────
# APP REGISTRATION
# ─────────────────────────────────────────────

function New-AppRegistration {
    param ([string]$Name)

    Write-Step "App Registration aanmaken"

    # Controleer of de app al bestaat
    $existing = Get-MgApplication -Filter "displayName eq '$Name'" -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Warn "App Registration '$Name' bestaat al (ID: $($existing.AppId))"
        $choice = Read-Host "  Bestaande app hergebruiken? (j/n)"
        if ($choice -eq "j") {
            Write-Ok "Bestaande App Registration wordt hergebruikt"
            return $existing
        }
        else {
            Write-Info "Nieuwe app aanmaken met timestamp..."
            $Name = "$Name $(Get-Date -Format 'yyyyMMdd-HHmm')"
        }
    }

    try {
        $app = New-MgApplication -DisplayName $Name -ErrorAction Stop
        Write-Ok "App Registration aangemaakt: '$Name'"
        Write-Ok "Application (Client) ID: $($app.AppId)"
        return $app
    }
    catch {
        Write-Err "Kon App Registration niet aanmaken: $($_.Exception.Message)"
        exit 1
    }
}

# ─────────────────────────────────────────────
# SERVICE PRINCIPAL
# ─────────────────────────────────────────────

function New-ServicePrincipalForApp {
    param ([string]$AppId)

    Write-Step "Service Principal aanmaken"

    $existing = Get-MgServicePrincipal -Filter "appId eq '$AppId'" -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Ok "Service Principal bestaat al"
        return $existing
    }

    try {
        $sp = New-MgServicePrincipal -AppId $AppId -ErrorAction Stop
        Write-Ok "Service Principal aangemaakt (ID: $($sp.Id))"
        return $sp
    }
    catch {
        Write-Err "Kon Service Principal niet aanmaken: $($_.Exception.Message)"
        exit 1
    }
}

# ─────────────────────────────────────────────
# API PERMISSIES
# ─────────────────────────────────────────────

function Set-GraphPermissions {
    param (
        [string]$AppObjectId,
        [string]$ServicePrincipalId
    )

    Write-Step "API permissies toewijzen"

    # Microsoft Graph App ID (altijd hetzelfde in elke tenant)
    $graphAppId = "00000003-0000-0000-c000-000000000000"
    $graphSp    = Get-MgServicePrincipal -Filter "appId eq '$graphAppId'" -ErrorAction Stop

    # Permissie die we nodig hebben
    $permissionName = "DeviceManagementConfiguration.ReadWrite.All"
    $appRole        = $graphSp.AppRoles | Where-Object { $_.Value -eq $permissionName }

    if (-not $appRole) {
        Write-Err "Kon permissie '$permissionName' niet vinden in Microsoft Graph"
        exit 1
    }

    # Controleer of permissie al toegewezen is
    $existing = Get-MgServicePrincipalAppRoleAssignment `
        -ServicePrincipalId $ServicePrincipalId `
        -ErrorAction SilentlyContinue |
        Where-Object { $_.AppRoleId -eq $appRole.Id }

    if ($existing) {
        Write-Ok "Permissie '$permissionName' is al toegewezen"
    }
    else {
        try {
            New-MgServicePrincipalAppRoleAssignment `
                -ServicePrincipalId $ServicePrincipalId `
                -PrincipalId        $ServicePrincipalId `
                -ResourceId         $graphSp.Id `
                -AppRoleId          $appRole.Id `
                -ErrorAction Stop | Out-Null

            Write-Ok "Permissie '$permissionName' toegewezen"
            Write-Ok "Admin consent automatisch verleend via Service Principal assignment"
        }
        catch {
            Write-Err "Kon permissie niet toewijzen: $($_.Exception.Message)"
            Write-Warn "Wijs de permissie handmatig toe in Entra ID > App registrations > API permissions"
        }
    }
}

# ─────────────────────────────────────────────
# CLIENT SECRET
# ─────────────────────────────────────────────

function New-ClientSecret {
    param (
        [string]$AppObjectId,
        [int]$ExpiryYears
    )

    Write-Step "Client Secret aanmaken"

    $endDate = (Get-Date).AddYears($ExpiryYears)

    try {
        $secret = Add-MgApplicationPassword `
            -ApplicationId    $AppObjectId `
            -PasswordCredential @{
                displayName = "Intune iOS Compliance Updater - $(Get-Date -Format 'yyyy-MM-dd')"
                endDateTime = $endDate
            } `
            -ErrorAction Stop

        Write-Ok "Client Secret aangemaakt (geldig tot: $($endDate.ToString('yyyy-MM-dd')))"
        Write-Warn "Bewaar deze waarde — hij wordt maar één keer getoond!"
        Write-Host "  Secret: $($secret.SecretText)" -ForegroundColor Yellow

        return $secret.SecretText
    }
    catch {
        Write-Err "Kon Client Secret niet aanmaken: $($_.Exception.Message)"
        exit 1
    }
}

# ─────────────────────────────────────────────
# COMPLIANCE POLICY OPZOEKEN
# ─────────────────────────────────────────────

function Get-CompliancePolicyId {
    param ([string]$PolicyName)

    Write-Step "Intune compliance policy opzoeken"

    try {
        $policies = Invoke-MgGraphRequest `
            -Method GET `
            -Uri    "https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies" `
            -ErrorAction Stop

        $iosPolicies = $policies.value | Where-Object {
            $_.'@odata.type' -eq "#microsoft.graph.iosCompliancePolicy"
        }

        if (-not $iosPolicies) {
            Write-Err "Geen iOS compliance policies gevonden in Intune"
            exit 1
        }

        # Als naam opgegeven is, zoek direct
        if ($PolicyName) {
            $match = $iosPolicies | Where-Object { $_.displayName -eq $PolicyName }
            if ($match) {
                Write-Ok "Policy gevonden: '$($match.displayName)' (ID: $($match.id))"
                return $match.id
            }
            else {
                Write-Warn "Policy '$PolicyName' niet gevonden. Beschikbare iOS policies:"
            }
        }
        else {
            Write-Info "Beschikbare iOS compliance policies:"
        }

        # Toon lijst en laat gebruiker kiezen
        $i = 1
        foreach ($policy in $iosPolicies) {
            Write-Host "  [$i] $($policy.displayName)" -ForegroundColor White
            $i++
        }

        $choice = Read-Host "`n  Kies een policy (nummer)"
        $index  = [int]$choice - 1

        if ($index -lt 0 -or $index -ge $iosPolicies.Count) {
            Write-Err "Ongeldige keuze"
            exit 1
        }

        $selected = $iosPolicies[$index]
        Write-Ok "Gekozen policy: '$($selected.displayName)' (ID: $($selected.id))"
        return $selected.id
    }
    catch {
        Write-Err "Kon compliance policies niet ophalen: $($_.Exception.Message)"
        exit 1
    }
}

# ─────────────────────────────────────────────
# CONFIG SCHRIJVEN
# ─────────────────────────────────────────────

function Write-Config {
    param (
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientSecret,
        [string]$PolicyId,
        [string]$Path
    )

    Write-Step "config.json aanmaken"

    $config = @{
        TenantId           = $TenantId
        ClientId           = $ClientId
        ClientSecret       = $ClientSecret
        CompliancePolicyId = $PolicyId
    } | ConvertTo-Json -Depth 3

    try {
        $config | Set-Content -Path $Path -Encoding UTF8 -ErrorAction Stop
        Write-Ok "config.json aangemaakt op: $Path"
        Write-Warn "Voeg config.json NOOIT toe aan Git! (staat al in .gitignore)"
    }
    catch {
        Write-Err "Kon config.json niet schrijven: $($_.Exception.Message)"
        Write-Info "Maak config.json handmatig aan met deze waarden:"
        Write-Host $config -ForegroundColor Yellow
    }
}

# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

Clear-Host
Write-Host "╔════════════════════════════════════════════╗" -ForegroundColor Magenta
Write-Host "║   Intune iOS Compliance Updater - Setup    ║" -ForegroundColor Magenta
Write-Host "║   BraveHub                                 ║" -ForegroundColor Magenta
Write-Host "╚════════════════════════════════════════════╝" -ForegroundColor Magenta

# 1. Systeem checks
Test-Prerequisites

# 2. Modules installeren
Install-RequiredModules

# 3. Aanmelden
$tenantId = Connect-ToGraph

# 4. App Registration aanmaken
$app = New-AppRegistration -Name $AppName

# 5. Service Principal aanmaken
$sp = New-ServicePrincipalForApp -AppId $app.AppId

# 6. API permissies toewijzen
Set-GraphPermissions -AppObjectId $app.Id -ServicePrincipalId $sp.Id

# 7. Client Secret aanmaken
$clientSecret = New-ClientSecret -AppObjectId $app.Id -ExpiryYears $SecretExpiryYears

# 8. Compliance Policy opzoeken
$policyId = Get-CompliancePolicyId -PolicyName $CompliancePolicyName

# 9. Config schrijven
Write-Config `
    -TenantId     $tenantId `
    -ClientId     $app.AppId `
    -ClientSecret $clientSecret `
    -PolicyId     $policyId `
    -Path         $ConfigPath

# Samenvatting
Write-Host "`n╔════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "║   Setup voltooid!                          ║" -ForegroundColor Green
Write-Host "╚════════════════════════════════════════════╝" -ForegroundColor Green
Write-Host ""
Write-Host "  Volgende stappen:" -ForegroundColor White
Write-Host "  1. Test het script: .\Update-iOSCompliancePolicy.ps1 -WhatIf" -ForegroundColor Cyan
Write-Host "  2. Registreer de scheduled task: .\Register-ScheduledTask.ps1" -ForegroundColor Cyan
Write-Host "  3. Controleer de logs in de logs\ map" -ForegroundColor Cyan
Write-Host ""
