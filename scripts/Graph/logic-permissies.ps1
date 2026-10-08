<#
.SYNOPSIS
    Grant a Microsoft Graph application permission to a Logic App's managed identity.
    Supports -WhatIf.

.DESCRIPTION
    A Logic App that calls Graph needs an app role on its own managed identity, which
    the portal cannot assign - only Graph itself can. This looks up the service
    principal of the managed identity by display name, resolves the requested app role
    on the Graph service principal, and assigns it when it is not there already.
    Re-running it changes nothing.

.PARAMETER TenantId
    Entra ID tenant to connect to. Optional: defaults to the GDAP customer when
    authMode is GDAP, else the tenant you sign in to.

.PARAMETER LogicAppName
    Display name of the Logic App's managed identity (its enterprise application).
    The name has to resolve to exactly one service principal.

.PARAMETER PermissionValue
    App role to assign (default: AuditLog.Read.All).

.PARAMETER ResourceAppId
    Application the role belongs to (default: Microsoft Graph).

.PARAMETER ModuleHandling
    Skip (default) leaves the modules to load.ps1 / scripts\Startup\Update-Modules.ps1.
    InstallIfMissing installs Microsoft.Graph.Authentication / .Applications when
    they are absent. InstallOrUpdate is kept for old calls and now does the same as
    InstallIfMissing: it no longer runs Update-Module -Force on every run.

.PARAMETER ClientId
    App registration for app-only sign-in (with -CertificateThumbprint and -TenantId).

.PARAMETER CertificateThumbprint
    Certificate thumbprint for app-only sign-in with -ClientId.

.PARAMETER AppOnly
    App-only sign-in with ClientId and CertificateThumbprint from graph.appid.json.

.NOTES
    Sign-in goes through scripts\Startup\Connect-M365.ps1: delegated as the admin by
    default (device code / GDAP customer per load.config.ps1), app-only on request.
    A fitting Graph session is reused and left connected; only a session this script
    opened is disconnected. Delegated scopes: Application.Read.All,
    AppRoleAssignment.ReadWrite.All (role: Privileged Role Administrator, or
    Cloud Application Administrator for non-Graph resources).

.EXAMPLE
    .\logic-permissies.ps1 -LogicAppName 'la-signin-report' -PermissionValue 'AuditLog.Read.All' -WhatIf

.EXAMPLE
    .\logic-permissies.ps1 -TenantId contoso.onmicrosoft.com -LogicAppName 'la-signin-report' -PermissionValue 'User.Read.All'
#>

#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$TenantId,

    [Parameter(Mandatory)]
    [string]$LogicAppName,

    [string]$PermissionValue = "AuditLog.Read.All",

    [string]$ResourceAppId = "00000003-0000-0000-c000-000000000000",

    [ValidateSet("InstallOrUpdate", "InstallIfMissing", "Skip")]
    [string]$ModuleHandling = "Skip",

    [string]$ClientId,

    [string]$CertificateThumbprint,

    [switch]$AppOnly
)

. (Join-Path $PSScriptRoot '..\Startup\Connect-M365.ps1')

$requiredScopes = @("Application.Read.All", "AppRoleAssignment.ReadWrite.All")
$requiredModules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.Applications")

function Ensure-ModuleReady {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][ValidateSet("InstallOrUpdate", "InstallIfMissing", "Skip")][string]$Handling
    )

    # Updating is the job of load.ps1 (scripts\Startup\Update-Modules.ps1); an
    # Update-Module -Force on every run was slow and could break a loaded session.
    if ($Handling -ne "Skip" -and -not (Get-Module -ListAvailable -Name $Name)) {
        Write-Host "Module '$Name' ontbreekt, installeren..." -ForegroundColor Yellow
        Install-Module -Name $Name -Scope CurrentUser -Force -ErrorAction Stop
    }

    if (-not (Get-Module -ListAvailable -Name $Name)) {
        throw "Benodigde module '$Name' ontbreekt. Draai scripts\Startup\Install-Modules.ps1 of gebruik -ModuleHandling InstallIfMissing."
    }

    Import-Module -Name $Name -ErrorAction Stop | Out-Null
}

foreach ($moduleName in $requiredModules) {
    Ensure-ModuleReady -Name $moduleName -Handling $ModuleHandling
}

$graph = Connect-M365Graph -Scopes $requiredScopes -TenantId $TenantId `
    -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -AppOnly:$AppOnly

try {
    $safeLogicAppName = $LogicAppName.Replace("'", "''")
    $logicAppMatches = @(Get-MgServicePrincipal -Filter "displayName eq '$safeLogicAppName'" -All)

    if ($logicAppMatches.Count -eq 0) {
        throw "Service principal niet gevonden voor '$LogicAppName'. Controleer of de Managed Identity al bestaat als enterprise app."
    }

    if ($logicAppMatches.Count -gt 1) {
        $ids = ($logicAppMatches | Select-Object -ExpandProperty Id) -join ", "
        throw "Meerdere service principals gevonden voor '$LogicAppName': $ids. Maak de naam uniek of pas filtering aan."
    }

    $logicAppSP = $logicAppMatches[0]

    $graphMatches = @(Get-MgServicePrincipal -Filter "appId eq '$ResourceAppId'" -All)
    if ($graphMatches.Count -eq 0) {
        throw "Resource service principal met appId '$ResourceAppId' niet gevonden."
    }

    $graphSP = $graphMatches[0]

    $appRole = $graphSP.AppRoles | Where-Object {
        $_.Value -eq $PermissionValue -and $_.AllowedMemberTypes -contains "Application"
    }

    if (-not $appRole) {
        throw "Application permission '$PermissionValue' niet gevonden op resource app '$ResourceAppId'."
    }

    $existingAssignments = @(Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $logicAppSP.Id -All)
    $assignmentAlreadyExists = $existingAssignments | Where-Object {
        $_.ResourceId -eq $graphSP.Id -and $_.AppRoleId -eq $appRole.Id
    }

    if ($assignmentAlreadyExists) {
        Write-Host "Permission '$PermissionValue' is al toegewezen aan '$LogicAppName'." -ForegroundColor Cyan
        return
    }

    if ($PSCmdlet.ShouldProcess($LogicAppName, "Ken permission '$PermissionValue' toe")) {
        $assignmentParams = @{
            ServicePrincipalId = $logicAppSP.Id
            PrincipalId        = $logicAppSP.Id
            ResourceId         = $graphSP.Id
            AppRoleId          = $appRole.Id
        }

        New-MgServicePrincipalAppRoleAssignment @assignmentParams | Out-Null

        Write-Host "Permission '$PermissionValue' toegekend aan '$LogicAppName' managed identity." -ForegroundColor Green
    }
}
finally {
    # Only a session this script opened; a session you already had stays.
    Disconnect-M365Graph $graph
}
