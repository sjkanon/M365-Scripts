[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string]$TenantId,

    [Parameter(Mandatory)]
    [string]$LogicAppName,

    [string]$PermissionValue = "AuditLog.Read.All",

    [string]$ResourceAppId = "00000003-0000-0000-c000-000000000000",

    [ValidateSet("InstallOrUpdate", "InstallIfMissing", "Skip")]
    [string]$ModuleHandling = "InstallOrUpdate"
)

$requiredScopes = @("Application.Read.All", "AppRoleAssignment.ReadWrite.All")
$requiredModules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.Applications")

function Ensure-ModuleReady {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][ValidateSet("InstallOrUpdate", "InstallIfMissing", "Skip")][string]$Handling
    )

    $installed = Get-Module -ListAvailable -Name $Name

    if ($Handling -eq "InstallOrUpdate") {
        if (-not $installed) {
            Write-Host "Module '$Name' ontbreekt, installeren..." -ForegroundColor Yellow
            Install-Module -Name $Name -Scope CurrentUser -Force -ErrorAction Stop
        }
        else {
            Write-Host "Module '$Name' gevonden, updaten..." -ForegroundColor Yellow
            Update-Module -Name $Name -Force -ErrorAction Stop
        }
    }
    elseif ($Handling -eq "InstallIfMissing") {
        if (-not $installed) {
            Write-Host "Module '$Name' ontbreekt, installeren..." -ForegroundColor Yellow
            Install-Module -Name $Name -Scope CurrentUser -Force -ErrorAction Stop
        }
    }

    if (-not (Get-Module -ListAvailable -Name $Name)) {
        throw "Benodigde module '$Name' ontbreekt."
    }

    Import-Module -Name $Name -ErrorAction Stop | Out-Null
}

foreach ($moduleName in $requiredModules) {
    Ensure-ModuleReady -Name $moduleName -Handling $ModuleHandling
}

Connect-MgGraph -Scopes $requiredScopes -TenantId $TenantId | Out-Null

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