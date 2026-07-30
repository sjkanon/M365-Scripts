#Requires -Version 5.1
<#
.SYNOPSIS
    Maakt/update de "Cowork Windows Prerequisites" Proactive Remediation in Intune en wijst 'm
    toe aan een Entra ID-groep met een dagelijks schema.

.DESCRIPTION
    Proactive Remediations (Detect + Remediate scriptparen) zijn geen Win32-apps — ze zitten
    achter een aparte Graph-resource (`deviceManagement/deviceHealthScripts`, alleen op de
    `beta`-API, er is geen v1.0-endpoint), dus de `IntuneWin32App`-module (gebruikt door
    Deploy-ClaudeDesktopIntune.ps1) is hier niet van toepassing. Dit script praat rechtstreeks
    met Graph via `Invoke-MgGraphRequest` (Microsoft.Graph.Authentication), delegated — geen
    tijdelijke App Registration nodig zoals bij de Win32-app-deploys, want dit vereist alleen
    DeviceManagementScripts.ReadWrite.All, geen los af te bakenen app-only recht.

    Leest Detect-/Remediate-CoworkPrerequisites.ps1 (in dezelfde map), base64-encodeert ze, en:
      1. Bestaat de remediation nog niet (op -DisplayName)? Dan wordt hij aangemaakt en direct
         toegewezen aan de opgegeven Entra-groep met een dagelijks schema.
      2. Bestaat hij al? Dan wordt de content (detectie-/remediatiescript) bijgewerkt via PATCH —
         altijd, ongeacht of er sinds de vorige run daadwerkelijk iets gewijzigd is (dat opzoeken
         zou de content moeten terugdownloaden en vergelijken; een PATCH is goedkoop en idempotent,
         dus dat is niet de moeite waard). De toewijzing zelf wordt bij een update NIET opnieuw
         gezet — de Graph "assign"-actie vervangt de VOLLEDIGE toewijzingenlijst van de remediation,
         dus dat gebeurt alleen bij eerste aanmaak of expliciet met -ReassignGroup.

    Vereiste rol: een rol die DeviceManagementScripts.ReadWrite.All-consent kan geven (bv. Intune
    Administrator, of Global Administrator).

.PARAMETER AssignmentGroupName
    Entra ID-groep. Gebruikt bij eerste aanmaak, en bij -ReassignGroup op een bestaande run.

.PARAMETER DisplayName
    Naam van de remediation in Intune. Gebruikt om 'm terug te vinden bij latere runs — niet
    wijzigen zonder de remediation ook in Intune zelf te hernoemen.

.PARAMETER IntervalDays
    Herhaalinterval in dagen voor het schema (standaard: 1 = dagelijks). Alleen toegepast bij
    eerste aanmaak of -ReassignGroup.

.PARAMETER RunTimeLocal
    Lokale tijd (HH:mm) waarop de remediation dagelijks draait. Standaard 03:00. Alleen toegepast
    bij eerste aanmaak of -ReassignGroup.

.PARAMETER ReassignGroup
    Zet de toewijzing (groep/schema) ook opnieuw als de remediation al bestaat — nodig om
    -AssignmentGroupName/-IntervalDays/-RunTimeLocal te wijzigen op een bestaande remediation.

.PARAMETER TenantId
    Entra ID tenant ID. Standaard automatisch bepaald uit de delegated sessie.

.PARAMETER Force
    Slaat de "typ JA om door te gaan"-bevestiging over.

.EXAMPLE
    .\Deploy-CoworkPrerequisitesRemediation.ps1 -AssignmentGroupName "SG-Apps-ClaudeDesktop"

.EXAMPLE
    # Schema wijzigen op een bestaande remediation
    .\Deploy-CoworkPrerequisitesRemediation.ps1 -AssignmentGroupName "SG-Apps-ClaudeDesktop" -IntervalDays 1 -RunTimeLocal "06:00" -ReassignGroup
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$AssignmentGroupName,

    [string]$DisplayName = 'Cowork Windows Prerequisites',

    [int]$IntervalDays = 1,

    [string]$RunTimeLocal = '03:00',

    [switch]$ReassignGroup,

    [string]$TenantId,

    [switch]$Force
)

function Write-Step {
    param([string]$Message)
    Write-Host ("`n==> {0}" -f $Message) -ForegroundColor Cyan
}

function Write-Info {
    param([string]$Message, [ConsoleColor]$ForegroundColor = 'DarkGray')
    Write-Host ("    {0}" -f $Message) -ForegroundColor $ForegroundColor
}

function Confirm-Action {
    param([Parameter(Mandatory = $true)][string]$Message)
    if ($Force) { return $true }
    Write-Host ""
    Write-Host "    $Message" -ForegroundColor Yellow
    $response = Read-Host "    Typ JA om door te gaan"
    return ($response -eq 'JA')
}

function Get-AllGraphPages {
    # $filter wordt niet betrouwbaar ondersteund op deviceHealthScripts (niet gedocumenteerd) —
    # daarom altijd de volledige lijst ophalen (incl. paginering) en client-side filteren.
    param([Parameter(Mandatory = $true)][string]$Uri)
    $results = [System.Collections.Generic.List[object]]::new()
    $nextUri = $Uri
    while ($nextUri) {
        $response = Invoke-MgGraphRequest -Method GET -Uri $nextUri -ErrorAction Stop
        if ($response.value) { $results.AddRange(@($response.value)) }
        $nextUri = $response.'@odata.nextLink'
    }
    return $results
}

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ''
Write-Host '  ================================================' -ForegroundColor Cyan
Write-Host '   Deploy-CoworkPrerequisitesRemediation' -ForegroundColor Cyan
Write-Host '  ================================================' -ForegroundColor Cyan

# ── Module preflight ──────────────────────────────────────────────────────────
$requiredModules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups')
$missingModules = $requiredModules | Where-Object { -not (Get-Module -ListAvailable -Name $_) }
if ($missingModules.Count -gt 0) {
    Write-Host "  [ERROR] Ontbrekende module(s): $($missingModules -join ', ')" -ForegroundColor Red
    Write-Host "  Installeer met: .\scripts\Startup\Install-Modules.ps1" -ForegroundColor Yellow
    exit 1
}

$detectPath = Join-Path $PSScriptRoot 'Detect-CoworkPrerequisites.ps1'
$remediatePath = Join-Path $PSScriptRoot 'Remediate-CoworkPrerequisites.ps1'
foreach ($p in @($detectPath, $remediatePath)) {
    if (-not (Test-Path $p)) {
        Write-Host "  [ERROR] Bestand niet gevonden: $p" -ForegroundColor Red
        exit 1
    }
}

$script:ConnectedHere = $false

try {
    Write-Step "Verbinden met Microsoft Graph (delegated)"
    Write-Info "Vereiste rol: rol die DeviceManagementScripts.ReadWrite.All-consent kan geven." -ForegroundColor Yellow
    $connectParams = @{ Scopes = @('DeviceManagementScripts.ReadWrite.All', 'Group.Read.All'); NoWelcome = $true }
    if ($TenantId) { $connectParams['TenantId'] = $TenantId }
    Connect-MgGraph @connectParams -ErrorAction Stop
    $script:ConnectedHere = $true
    Write-Info "[OK]   Verbonden."

    Write-Step "Toewijzingsgroep opzoeken"
    $escapedGroupName = $AssignmentGroupName.Replace("'", "''")
    $groups = @(Get-MgGroup -Filter "displayName eq '$escapedGroupName'" -ErrorAction Stop)
    if ($groups.Count -eq 0) {
        throw "Entra-groep '$AssignmentGroupName' niet gevonden."
    } elseif ($groups.Count -gt 1) {
        throw "Meerdere Entra-groepen met de naam '$AssignmentGroupName' gevonden — gebruik een unieke naam."
    }
    $assignmentGroup = $groups[0]
    Write-Info "[OK]   Groep gevonden: $($assignmentGroup.DisplayName) ($($assignmentGroup.Id))"

    Write-Step "Content-scripts inlezen"
    $detectionB64   = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes((Get-Content $detectPath -Raw)))
    $remediationB64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes((Get-Content $remediatePath -Raw)))
    Write-Info "[OK]   Detect-CoworkPrerequisites.ps1 + Remediate-CoworkPrerequisites.ps1 gelezen."

    $scriptBody = @{
        '@odata.type'            = '#microsoft.graph.deviceHealthScript'
        displayName              = $DisplayName
        description              = 'Enables VirtualMachinePlatform and disables Fast Startup for Claude Cowork, independent of Claude Desktop. Managed by Deploy-CoworkPrerequisitesRemediation.ps1.'
        publisher                = 'IT'
        runAsAccount             = 'system'
        enforceSignatureCheck    = $false
        runAs32Bit               = $false
        detectionScriptContent   = $detectionB64
        remediationScriptContent = $remediationB64
    }

    Write-Step "Remediation opzoeken in Intune: '$DisplayName'"
    $allScripts = Get-AllGraphPages -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts?`$top=200"
    $existingScript = $allScripts | Where-Object { $_.displayName -eq $DisplayName } | Select-Object -First 1

    $targetId = $null
    $needsAssignment = $false

    if ($existingScript) {
        Write-Step "Bestaande remediation bijwerken (ID $($existingScript.id))"
        if (-not (Confirm-Action "Content van '$DisplayName' bijwerken in PRODUCTIE-Intune?")) {
            Write-Host "  Afgebroken door gebruiker." -ForegroundColor Yellow
            exit 1
        }
        Invoke-MgGraphRequest -Method PATCH `
            -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts/$($existingScript.id)" `
            -Body ($scriptBody | ConvertTo-Json -Depth 5) -ErrorAction Stop | Out-Null
        Write-Info "[OK]   Content bijgewerkt." -ForegroundColor Green
        $targetId = $existingScript.id
        $needsAssignment = [bool]$ReassignGroup
        if (-not $needsAssignment) {
            Write-Info "Toewijzing ongewijzigd gelaten (gebruik -ReassignGroup om groep/schema te wijzigen)."
        }
    } else {
        Write-Step "Nieuwe remediation aanmaken"
        if (-not (Confirm-Action "Nieuwe remediation '$DisplayName' aanmaken en toewijzen aan groep '$($assignmentGroup.DisplayName)'?")) {
            Write-Host "  Afgebroken door gebruiker." -ForegroundColor Yellow
            exit 1
        }
        $created = Invoke-MgGraphRequest -Method POST `
            -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts" `
            -Body ($scriptBody | ConvertTo-Json -Depth 5) -ErrorAction Stop
        $targetId = $created.id
        $needsAssignment = $true
        Write-Info "[OK]   Aangemaakt (ID $targetId)." -ForegroundColor Green
    }

    if ($needsAssignment) {
        Write-Step ("Toewijzen aan '{0}' (elke {1} dag(en) om {2} lokale tijd)" -f $assignmentGroup.DisplayName, $IntervalDays, $RunTimeLocal)
        $assignBody = @{
            deviceHealthScriptAssignments = @(
                @{
                    target = @{
                        '@odata.type' = '#microsoft.graph.groupAssignmentTarget'
                        groupId       = $assignmentGroup.Id
                    }
                    runRemediationScript = $true
                    runSchedule = @{
                        '@odata.type' = '#microsoft.graph.deviceHealthScriptDailySchedule'
                        interval      = $IntervalDays
                        useUtc        = $false
                        time          = "$($RunTimeLocal):00"
                    }
                }
            )
        }
        Invoke-MgGraphRequest -Method POST `
            -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts/$targetId/assign" `
            -Body ($assignBody | ConvertTo-Json -Depth 6) -ErrorAction Stop | Out-Null
        Write-Info "[OK]   Toegewezen." -ForegroundColor Green
    }
}
catch {
    Write-Host "`n  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
finally {
    if ($script:ConnectedHere) {
        try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
    }
}

Write-Host "`n  ================================================" -ForegroundColor Cyan
Write-Host "   Klaar — Cowork Prerequisites Remediation" -ForegroundColor Cyan
Write-Host "  ================================================`n" -ForegroundColor Cyan
exit 0
