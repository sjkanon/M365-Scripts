#Requires -Version 5.1
<#
.SYNOPSIS
    Draait Deploy-CoworkPrerequisitesIntune.ps1 en Deploy-ClaudeDesktopIntune.ps1 na elkaar, in
    één aanroep.

.DESCRIPTION
    Pure orchestrator — bevat zelf geen Intune/Graph-logica, roept alleen de twee bestaande
    deploy-scripts na elkaar aan met dezelfde -AssignmentGroupName/-TenantId/-Force. Cowork
    Prerequisites draait eerst: als je later -RequireCoworkPrerequisites gebruikt bij Claude
    Desktop (zie ClaudeDesktop/readme.md), moet die app al bestaan in Intune.

    Elke deelrun beheert zijn eigen Microsoft Graph-sessie en tijdelijke App Registration
    volledig zelfstandig (aanmaken, gebruiken, opruimen) — dit script deelt niets tussen de twee
    runs, het is puur twee losse, achter-elkaar uitgevoerde aanroepen. Faalt de Cowork-run, dan
    stopt dit script vóór Claude Desktop (geen zin om door te gaan als de vereiste al niet lukte),
    tenzij -ContinueOnError is opgegeven.

.PARAMETER AssignmentGroupName
    Entra ID-groep, doorgegeven aan beide deelscripts.

.PARAMETER TenantId
    Optioneel, doorgegeven aan beide deelscripts. Standaard automatisch bepaald uit de
    delegated sessie (per deelscript apart).

.PARAMETER Force
    Slaat de "typ JA om door te gaan"-bevestiging over voor beide deelscripts.

.PARAMETER SkipCoworkPrerequisites
    Sla de Cowork Prerequisites-deploy over, draai alleen Claude Desktop.

.PARAMETER ContinueOnError
    Ga door met de Claude Desktop-deploy ook als de Cowork Prerequisites-deploy is mislukt.
    Standaard stopt dit script in dat geval.

.EXAMPLE
    .\Deploy-AllIntune.ps1 -AssignmentGroupName "SG-Apps-ClaudeDesktop"

.EXAMPLE
    # Onbeheerd, bv. via een geplande taak
    .\Deploy-AllIntune.ps1 -AssignmentGroupName "SG-Apps-ClaudeDesktop" -Force

.EXAMPLE
    # Alleen Claude Desktop, Cowork Prerequisites al eerder/apart gedaan
    .\Deploy-AllIntune.ps1 -AssignmentGroupName "SG-Apps-ClaudeDesktop" -SkipCoworkPrerequisites
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$AssignmentGroupName,

    [string]$TenantId,

    [switch]$Force,

    [switch]$SkipCoworkPrerequisites,

    [switch]$ContinueOnError
)

$sharedParams = @{ AssignmentGroupName = $AssignmentGroupName }
if ($TenantId) { $sharedParams['TenantId'] = $TenantId }
if ($Force) { $sharedParams['Force'] = $true }

$coworkScript = Join-Path $PSScriptRoot 'CoworkPrerequisites\Deploy-CoworkPrerequisitesIntune.ps1'
$claudeScript = Join-Path $PSScriptRoot 'ClaudeDesktop\Deploy-ClaudeDesktopIntune.ps1'

if (-not $SkipCoworkPrerequisites) {
    Write-Host "`n########## Cowork Windows Prerequisites ##########" -ForegroundColor Magenta
    & $coworkScript @sharedParams
    if ($LASTEXITCODE -ne 0) {
        if (-not $ContinueOnError) {
            Write-Host "`n[ERROR] Cowork Prerequisites-deploy is mislukt (exit code $LASTEXITCODE) — Claude Desktop wordt niet gestart." -ForegroundColor Red
            Write-Host "        Gebruik -ContinueOnError om dat te forceren." -ForegroundColor Red
            exit 1
        }
        Write-Host "`n[WARN] Cowork Prerequisites-deploy is mislukt (exit code $LASTEXITCODE) — doorgegaan met Claude Desktop door -ContinueOnError." -ForegroundColor Yellow
    }
} else {
    Write-Host "`nCowork Prerequisites-deploy overgeslagen (-SkipCoworkPrerequisites)." -ForegroundColor DarkGray
}

Write-Host "`n########## Claude Desktop ##########" -ForegroundColor Magenta
& $claudeScript @sharedParams
$claudeExitCode = $LASTEXITCODE

if ($claudeExitCode -ne 0) {
    Write-Host "`n[ERROR] Claude Desktop-deploy is mislukt (exit code $claudeExitCode)." -ForegroundColor Red
    exit $claudeExitCode
}

Write-Host "`nBeide deploys afgerond." -ForegroundColor Green
