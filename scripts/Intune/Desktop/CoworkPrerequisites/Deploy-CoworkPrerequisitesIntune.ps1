#Requires -Version 5.1
<#
.SYNOPSIS
    Deploy-script voor de Cowork Windows-vereisten (VirtualMachinePlatform + Fast Startup) als
    eigen, onafhankelijke Intune Win32-app — los van Claude Desktop zelf.

.DESCRIPTION
    Zusje van Deploy-ClaudeDesktopIntune.ps1, maar dan voor de Windows-kant van Cowork. Er is
    hier geen MSIX om te downloaden — de content is alleen de drie meegeleverde .ps1-scripts in
    deze map. Draai dit script opnieuw wanneer je Install-/Uninstall-/
    Detect-CoworkPrerequisites-Intune.ps1 zelf wijzigt; zonder codewijziging doet een herhaalde
    run niets.

      1. Kopieert Install-/Uninstall-CoworkPrerequisites-Intune.ps1 naast dit script naar de
         staging-map en pakt ze in tot een .intunewin-bestand.
      2. Verbindt delegated met Microsoft Graph en maakt — net als Deploy-ClaudeDesktopIntune.ps1
         — een kortlevende tijdelijke App Registration aan met alleen
         DeviceManagementApps.ReadWrite.All, gebruikt om de app te uploaden/updaten, en aan het
         einde weer verwijderd.
      3. Bestaat de Intune-app nog niet? Dan wordt hij aangemaakt met detectie-/requirement-
         regels en toegewezen (Required) aan de opgegeven Entra-groep. Bestaat hij al en zijn de
         drie content-scripts gewijzigd sinds de vorige run (SHA256-hash in het Notes-veld,
         ScriptsHash=...)? Dan wordt de package-inhoud bijgewerkt — de bestaande toewijzing
         blijft ongewijzigd.
      4. Detectie- en requirement-regel worden bij ELKE run opnieuw opgebouwd en meegestuurd,
         niet alleen bij eerste aanmaak — zelfde reden als in Deploy-ClaudeDesktopIntune.ps1 (een
         fout die ooit bij de allereerste Add-IntuneWin32App-call is vastgelegd blijft anders voor
         altijd op de app staan). Bij een UPDATE wordt de requirement rule bewust niet via
         Set-IntuneWin32App's eigen -RequirementRule-parameter gezet — die cmdlet heeft in de
         geïnstalleerde IntuneWin32App-module (1.5.0) een bevestigde bug waarbij die parameter een
         niet-bestaand '@odata.type'-veld eist en bij het ontbreken daarvan de HELE rest van de
         functie afbreekt (dus ook Notes/DetectionRule/RestartBehavior zouden dan nooit wegschrijven)
         — zie Set-Win32AppArchitectureRequirement hieronder, die architecture/OS-release rechtstreeks
         via Graph PATCHt, buiten die kapotte route om. Add-IntuneWin32App (eerste aanmaak) heeft
         deze bug niet.
      5. De app is `-CompanyPortalFeaturedApp` op elke run, zodat hij zichtbaar/uitgelicht in
         Company Portal staat i.p.v. verborgen te blijven als achtergrond-vereiste.

    Bewust een ONAFHANKELIJKE app, zonder Intune-dependency naar/van de Claude Desktop-app: een
    mislukte of nog-niet-voltooide Cowork-prereqs-install mag Claude Desktop zelf niet
    tegenhouden (dat werkt prima zonder Cowork), en omgekeerd moet een probleem in Claude
    Desktop's eigen installatie niet verward kunnen worden met een Cowork/Windows-featureprobleem
    — beide apps hebben hun eigen, apart zichtbare install-status in Intune.

    Vereiste rol: Global Administrator, of Application Administrator in combinatie met een rol
    die AppRoleAssignment.ReadWrite.All-consent mag geven.

.PARAMETER AssignmentGroupName
    Displaynaam van een bestaande Entra ID-groep. Alleen gebruikt bij de allereerste aanmaak van
    de Intune-app; latere content-updates raken de toewijzing niet aan.

.PARAMETER WorkingDirectory
    Werkmap voor build-artefacten (Source/, Output/). Standaard: C:\Temp\CoworkPrereqDeploy.

.PARAMETER AppDisplayName
    Displaynaam van de Win32-app in Intune. Wordt gebruikt om de bestaande app terug te vinden bij
    latere runs — niet wijzigen zonder de app ook in Intune zelf te hernoemen.

.PARAMETER MinimumSupportedWindowsRelease
    Requirement rule: minimale Windows-release (bv. 'W10_21H2', 'W11_23H2').

.PARAMETER TenantId
    Entra ID tenant ID. Standaard automatisch bepaald uit de delegated sessie.

.PARAMETER IntuneWinAppUtilPath
    Pad naar een al aanwezige IntuneWinAppUtil.exe. Zonder opgave downloadt de IntuneWin32App-
    module deze automatisch naar %TEMP%.

.PARAMETER Force
    Slaat de "typ JA om door te gaan"-bevestiging over.

.EXAMPLE
    .\Deploy-CoworkPrerequisitesIntune.ps1 -AssignmentGroupName "SG-Apps-ClaudeDesktop"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$AssignmentGroupName,

    [string]$WorkingDirectory = 'C:\Temp\CoworkPrereqDeploy',

    [string]$AppDisplayName = 'Cowork Windows Prerequisites (Machine-wide)',

    [string]$MinimumSupportedWindowsRelease = 'W10_21H2',

    [string]$TenantId,

    [string]$IntuneWinAppUtilPath,

    [switch]$Force
)

# ── Cleanup tracking ─────────────────────────────────────────────────────────
$script:TempAppObjectId  = $null
$script:ConnectedHere    = $false

# ── Cultuur tijdelijk op invariant zetten (zie Deploy-ClaudeDesktopIntune.ps1 voor de reden:
#    IntuneWin32App-module upstream issue #210, crasht op dd/MM/jjjj-notatie zoals nl-NL) ──
$script:OriginalCulture   = [System.Threading.Thread]::CurrentThread.CurrentCulture
$script:OriginalUICulture = [System.Threading.Thread]::CurrentThread.CurrentUICulture
[System.Threading.Thread]::CurrentThread.CurrentCulture   = [System.Globalization.CultureInfo]::InvariantCulture
[System.Threading.Thread]::CurrentThread.CurrentUICulture = [System.Globalization.CultureInfo]::InvariantCulture

function Write-Step {
    param([string]$Message, [ConsoleColor]$ForegroundColor = 'Cyan')
    Write-Host ("`n==> {0}" -f $Message) -ForegroundColor $ForegroundColor
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

function Get-ScriptsHashHex {
    param([Parameter(Mandatory = $true)][string[]]$Paths)
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        $combined = ($Paths | ForEach-Object { Get-Content -Path $_ -Raw }) -join "`n---`n"
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($combined)
        $hash  = $sha256.ComputeHash($bytes)
        return ([System.BitConverter]::ToString($hash) -replace '-', '').ToLowerInvariant().Substring(0, 12)
    } finally {
        $sha256.Dispose()
    }
}

function Set-Win32AppArchitectureRequirement {
    # Zelfde bevestigde IntuneWin32App-moduleprobleem (1.5.0) als in Deploy-ClaudeDesktopIntune.ps1:
    # Set-IntuneWin32App (de UPDATE-cmdlet) eist ten onrechte '@odata.type' op -RequirementRule
    # (architecture/minimumSupportedWindowsRelease zijn platte win32LobApp-properties, geen
    # polymorf 'rules'-lid — zie https://learn.microsoft.com/en-us/graph/api/resources/intune-apps-win32lobapp).
    # New-IntuneWin32AppRequirementRule zet '@odata.type' nooit, dus die check slaat altijd aan, en
    # de "break" erna (geen loop/switch eromheen — empirisch geverifieerd) breekt de HELE rest van
    # de functie af, ruim vóór de eigenlijke Graph-PATCH-call verderop: Notes/DetectionRule/
    # RestartBehavior zouden dus NOOIT wegschrijven zolang -RequirementRule werd meegegeven aan een
    # UPDATE-aanroep. Add-IntuneWin32App (eerste aanmaak) heeft dit euvel niet. Deze functie zet
    # architecture/minimumSupportedWindowsRelease daarom rechtstreeks via Graph, buiten de kapotte
    # cmdlet-route om, met dezelfde sessie die Connect-MSIntuneGraph al opzette.
    param(
        [Parameter(Mandatory = $true)][string]$AppId,
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$RequirementRule
    )
    if (-not $Global:AuthenticationHeader) {
        throw "Geen actieve Intune Graph-sessie (Global:AuthenticationHeader ontbreekt) — Connect-MSIntuneGraph moet al zijn uitgevoerd."
    }
    $body = @{
        '@odata.type'                    = '#microsoft.graph.win32LobApp'
        allowedArchitectures             = $RequirementRule['allowedArchitectures']
        minimumSupportedWindowsRelease   = $RequirementRule['minimumSupportedWindowsRelease']
    } | ConvertTo-Json -Depth 5

    try {
        Invoke-RestMethod -Method Patch `
            -Uri "https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps/$AppId" `
            -Headers $Global:AuthenticationHeader -Body $body -ContentType 'application/json' -ErrorAction Stop | Out-Null
    } catch {
        $graphErrorDetail = $null
        try {
            if ($_.Exception.Response) {
                $streamReader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                try {
                    $streamReader.BaseStream.Position = 0
                    $graphErrorDetail = $streamReader.ReadToEnd()
                } finally {
                    $streamReader.Dispose()
                }
            }
        } catch {}

        if ($graphErrorDetail) {
            throw "Graph PATCH van architecture requirement mislukt: $graphErrorDetail"
        } else {
            throw
        }
    }
}

function New-TempAppRegistration {
    param(
        [Parameter(Mandatory = $true)][string]$AppName,
        [Parameter(Mandatory = $true)][string]$UsedTenantId
    )

    Write-Info "Tijdelijke App Registration '$AppName' aanmaken..." -ForegroundColor Cyan
    $app = New-MgApplication -DisplayName $AppName -ErrorAction Stop
    $sp  = New-MgServicePrincipal -AppId $app.AppId -ErrorAction Stop

    $graphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'" -ErrorAction Stop
    $appRole = $graphSp.AppRoles | Where-Object { $_.Value -eq 'DeviceManagementApps.ReadWrite.All' }
    if (-not $appRole) {
        throw "Kon Graph app role 'DeviceManagementApps.ReadWrite.All' niet vinden."
    }
    New-MgServicePrincipalAppRoleAssignment `
        -ServicePrincipalId $sp.Id `
        -PrincipalId        $sp.Id `
        -ResourceId         $graphSp.Id `
        -AppRoleId          $appRole.Id `
        -ErrorAction Stop | Out-Null
    Write-Info "[OK]   DeviceManagementApps.ReadWrite.All toegekend."

    $secret = Add-MgApplicationPassword `
        -ApplicationId      $app.Id `
        -PasswordCredential @{ displayName = 'temp'; endDateTime = (Get-Date).AddDays(1) } `
        -ErrorAction Stop

    return [PSCustomObject]@{
        AppObjectId  = $app.Id
        AppId        = $app.AppId
        ClientSecret = $secret.SecretText
    }
}

function Remove-TempApp {
    param([string]$ObjectId)
    if (-not $ObjectId) { return }
    Write-Info "Tijdelijke App Registration verwijderen..."
    try {
        Remove-MgApplication -ApplicationId $ObjectId -ErrorAction Stop
        Write-Info "[OK]   Tijdelijke App Registration verwijderd."
    } catch {
        Write-Host "    [WARN] Kon tijdelijke App Registration niet verwijderen (Object ID: $ObjectId)." -ForegroundColor Yellow
        Write-Host "    [WARN] Verwijder deze handmatig in Entra ID > App registrations." -ForegroundColor Yellow
    }
}

# ── Header ────────────────────────────────────────────────────────────────────
Write-Host ''
Write-Host '  ================================================' -ForegroundColor Cyan
Write-Host '   Deploy-CoworkPrerequisitesIntune' -ForegroundColor Cyan
Write-Host '  ================================================' -ForegroundColor Cyan

# ── Module preflight ──────────────────────────────────────────────────────────
$requiredModules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Applications', 'Microsoft.Graph.Groups', 'IntuneWin32App')
$missingModules = $requiredModules | Where-Object { -not (Get-Module -ListAvailable -Name $_) }
if ($missingModules.Count -gt 0) {
    Write-Host "  [ERROR] Ontbrekende module(s): $($missingModules -join ', ')" -ForegroundColor Red
    Write-Host "  Installeer met: .\scripts\Startup\Install-Modules.ps1" -ForegroundColor Yellow
    exit 1
}

[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

# ── Werkmap voorbereiden ─────────────────────────────────────────────────────
Write-Step "Werkmap voorbereiden ($WorkingDirectory)"
$sourceDir = Join-Path $WorkingDirectory 'Source'
$outputDir = Join-Path $WorkingDirectory 'Output'
foreach ($dir in @($sourceDir, $outputDir)) {
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
}

# ── Content-scripts naast elkaar plaatsen ────────────────────────────────────
Copy-Item -Path (Join-Path $PSScriptRoot 'Install-CoworkPrerequisites-Intune.ps1')   -Destination $sourceDir -Force
Copy-Item -Path (Join-Path $PSScriptRoot 'Uninstall-CoworkPrerequisites-Intune.ps1') -Destination $sourceDir -Force
$detectScriptPath = Join-Path $PSScriptRoot 'Detect-CoworkPrerequisites-Intune.ps1'

# Hash van alle vier de scripts (incl. dit deploy-script zelf) — zie Deploy-ClaudeDesktopIntune.ps1
# voor dezelfde reden: een toekomstige fix puur in dit orchestratie-script moet zichzelf ook
# zonder MSIX-versie (die hier sowieso niet bestaat) naar een bestaande app kunnen pushen.
$newScriptsHash = Get-ScriptsHashHex -Paths @(
    $PSCommandPath,
    (Join-Path $PSScriptRoot 'Install-CoworkPrerequisites-Intune.ps1'),
    (Join-Path $PSScriptRoot 'Uninstall-CoworkPrerequisites-Intune.ps1'),
    $detectScriptPath
)

# ── .intunewin bouwen ────────────────────────────────────────────────────────
Write-Step "Package bouwen (.intunewin)"
$packageParams = @{
    SourceFolder = $sourceDir
    SetupFile    = 'Install-CoworkPrerequisites-Intune.ps1'
    OutputFolder = $outputDir
    Force        = $true
}
if ($IntuneWinAppUtilPath) { $packageParams['IntuneWinAppUtilPath'] = $IntuneWinAppUtilPath }
try {
    $win32AppPackage = New-IntuneWin32AppPackage @packageParams -ErrorAction Stop
} catch {
    Write-Host "  [ERROR] Bouwen van .intunewin mislukt: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
$intuneWinFile = $win32AppPackage.Path
if ([string]::IsNullOrWhiteSpace($intuneWinFile) -or -not (Test-Path $intuneWinFile)) {
    Write-Host "  [ERROR] New-IntuneWin32AppPackage leverde geen geldig .intunewin-pad op." -ForegroundColor Red
    exit 1
}
Write-Info "[OK]   Package gebouwd: $intuneWinFile"

# ── Microsoft Graph (delegated) ──────────────────────────────────────────────
Write-Step "Verbinden met Microsoft Graph (delegated)"
Write-Info "Vereiste rol: Global Administrator of Application Administrator." -ForegroundColor Yellow
try {
    $connectParams = @{
        Scopes    = @('Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All', 'Group.Read.All')
        NoWelcome = $true
    }
    if ($TenantId) { $connectParams['TenantId'] = $TenantId }
    Connect-MgGraph @connectParams -ErrorAction Stop
    $script:ConnectedHere = $true
    Write-Info "[OK]   Verbonden (delegated)."
} catch {
    Write-Host "  [ERROR] Verbinden met Microsoft Graph mislukt: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

try {
    $ctx = Get-MgContext
    $effectiveTenantId = if ($TenantId) { $TenantId } else { $ctx.TenantId }
    if (-not $effectiveTenantId) {
        throw "Kon tenant ID niet bepalen. Geef -TenantId op."
    }

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

    Write-Step "Tijdelijke App Registration aanmaken (DeviceManagementApps.ReadWrite.All)"
    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $tempApp = New-TempAppRegistration -AppName "CoworkPrereqDeploy-Temp-$ts" -UsedTenantId $effectiveTenantId
    $script:TempAppObjectId = $tempApp.AppObjectId

    Write-Step "Verbinden met Intune (app-only, via tijdelijke App Registration)"
    $connected = $false
    $maxAttempts = 12
    $retryDelaySeconds = 10
    for ($i = 1; $i -le $maxAttempts; $i++) {
        try {
            Connect-MSIntuneGraph -TenantID $effectiveTenantId -ClientID $tempApp.AppId -ClientSecret $tempApp.ClientSecret -ErrorAction Stop -WarningAction Stop | Out-Null
            Get-IntuneWin32App -ErrorAction Stop -WarningAction Stop | Out-Null
            $connected = $true
            break
        } catch {
            if ($i -lt $maxAttempts) {
                Write-Info "Wachten op propagatie van de tijdelijke App Registration (poging $i/$maxAttempts): $($_.Exception.Message)"
                Start-Sleep -Seconds $retryDelaySeconds
            }
        }
    }
    if (-not $connected) {
        throw "Kon niet verbinden met Intune via de tijdelijke App Registration (na $maxAttempts pogingen)."
    }
    Write-Info "[OK]   Verbonden met Intune."

    Write-Step "Intune Win32-app opzoeken: '$AppDisplayName'"
    $existingApps = @(Get-IntuneWin32App -DisplayName $AppDisplayName -ErrorAction SilentlyContinue)
    $existingApp = if ($existingApps.Count -gt 0) { $existingApps[0] } else { $null }

    $existingScriptsHash = $null
    if ($existingApp -and $existingApp.notes -match 'ScriptsHash=([0-9a-f]+)') {
        $existingScriptsHash = $Matches[1]
    }

    $notes = "ScriptsHash=$newScriptsHash; LastUpdated=$(Get-Date -Format o); Deployed by Deploy-CoworkPrerequisitesIntune.ps1"
    $scriptsUnchanged = $existingScriptsHash -eq $newScriptsHash

    # Detectie- en requirement-regel bij ELKE run herbouwd en meegestuurd — zie
    # Deploy-ClaudeDesktopIntune.ps1 voor de reden (0x80070001-les: een fout uit de allereerste
    # Add-IntuneWin32App-call blijft anders voor altijd op de app staan).
    $detectionRule = New-IntuneWin32AppDetectionRuleScript -ScriptFile $detectScriptPath -EnforceSignatureCheck $false -RunAs32Bit $false
    $requirementRule = New-IntuneWin32AppRequirementRule -Architecture 'x64' -MinimumSupportedWindowsRelease $MinimumSupportedWindowsRelease

    if ($existingApp -and $scriptsUnchanged) {
        Write-Step "Al up-to-date"
        Write-Info "Intune-app '$AppDisplayName' heeft ongewijzigde content-scripts. Geen wijzigingen nodig." -ForegroundColor Green
    }
    elseif ($existingApp) {
        Write-Step "Bestaande app bijwerken (install-/uninstall-/detect-scripts gewijzigd)"
        if (-not (Confirm-Action "Package-inhoud van '$AppDisplayName' (ID $($existingApp.id)) bijwerken in PRODUCTIE-Intune?")) {
            Write-Host "  Afgebroken door gebruiker." -ForegroundColor Yellow
            exit 1
        }
        Update-IntuneWin32AppPackageFile -ID $existingApp.id -FilePath $intuneWinFile -ErrorAction Stop | Out-Null
        Set-IntuneWin32App -ID $existingApp.id -Notes $notes -DetectionRule $detectionRule `
            -RestartBehavior 'basedOnExitCode' -CompanyPortalFeaturedApp $true -ErrorAction Stop | Out-Null
        Set-Win32AppArchitectureRequirement -AppId $existingApp.id -RequirementRule $requirementRule
        Write-Info "[OK]   App bijgewerkt (incl. ververste requirement rule)." -ForegroundColor Green
    }
    else {
        Write-Step "Nieuwe Intune Win32-app aanmaken (eerste run)"
        if (-not (Confirm-Action "Nieuwe app '$AppDisplayName' aanmaken in PRODUCTIE-Intune en toewijzen (Required) aan groep '$($assignmentGroup.DisplayName)'?")) {
            Write-Host "  Afgebroken door gebruiker." -ForegroundColor Yellow
            exit 1
        }

        $sysnativePwsh = '%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe'
        $installCommandLine   = "$sysnativePwsh -ExecutionPolicy Bypass -File Install-CoworkPrerequisites-Intune.ps1"
        $uninstallCommandLine = "$sysnativePwsh -ExecutionPolicy Bypass -File Uninstall-CoworkPrerequisites-Intune.ps1"

        $newApp = Add-IntuneWin32App `
            -FilePath              $intuneWinFile `
            -DisplayName           $AppDisplayName `
            -Description           'Windows-vereisten voor Claude Cowork (VirtualMachinePlatform + Fast Startup uit), los van Claude Desktop zelf — zie CoworkPrerequisites/readme.md.' `
            -Publisher              'Anthropic' `
            -AppVersion             '1.0' `
            -Notes                  $notes `
            -InstallCommandLine     $installCommandLine `
            -UninstallCommandLine   $uninstallCommandLine `
            -InstallExperience      'system' `
            -RestartBehavior        'basedOnExitCode' `
            -DetectionRule          $detectionRule `
            -RequirementRule        $requirementRule `
            -CompanyPortalFeaturedApp $true `
            -ErrorAction Stop

        Add-IntuneWin32AppAssignmentGroup -Include -ID $newApp.id -GroupID $assignmentGroup.Id -Intent 'required' -Notification 'showAll' -ErrorAction Stop | Out-Null
        Write-Info "[OK]   App aangemaakt (ID $($newApp.id)) en toegewezen aan '$($assignmentGroup.DisplayName)'." -ForegroundColor Green
    }
}
catch {
    Write-Host "`n  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
finally {
    Write-Step "Opruimen"
    Remove-TempApp -ObjectId $script:TempAppObjectId
    if ($script:ConnectedHere) {
        try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
    }
    [System.Threading.Thread]::CurrentThread.CurrentCulture   = $script:OriginalCulture
    [System.Threading.Thread]::CurrentThread.CurrentUICulture = $script:OriginalUICulture
}

Write-Host "`n  ================================================" -ForegroundColor Cyan
Write-Host "   Klaar — Cowork Windows Prerequisites" -ForegroundColor Cyan
Write-Host "  ================================================`n" -ForegroundColor Cyan
