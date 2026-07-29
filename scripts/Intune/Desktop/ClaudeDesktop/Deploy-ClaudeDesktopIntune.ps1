#Requires -Version 5.1
<#
.SYNOPSIS
    Eén-op-de-maand deploy-script: haalt de nieuwste Claude Desktop MSIX op, pakt hem in als
    Intune Win32-app en maakt/update de app in Intune — volledig automatisch, zonder handmatig
    voorbereide App Registration.

.DESCRIPTION
    Dit is het enige script dat je maandelijks handmatig draait om Claude Desktop actueel te
    houden via Intune. Het combineert alle stappen uit de eerdere losse scripts
    (Deploy-ClaudeDesktop.ps1 / Install- / Detect-ClaudeDesktop-Intune.ps1) tot één run:

      1. Downloadt de nieuwste Claude Desktop MSIX (x64) via de officiële "latest"-redirect-URL.
      2. Leest de versie uit AppxManifest.xml in de MSIX.
      3. Kopieert de meegeleverde Install-/Uninstall-ClaudeDesktop-Intune.ps1 (in dezelfde map als
         dit script) naast de MSIX en pakt alles in tot een .intunewin-bestand.
      4. Verbindt delegated met Microsoft Graph en maakt — net als
         Remove-SharePointFileVersionsByDate.ps1 / Get-SharePointStorageReport.ps1 — een
         kortlevende tijdelijke App Registration aan met alleen het Graph app-only recht
         DeviceManagementApps.ReadWrite.All. Die wordt gebruikt om via de IntuneWin32App-module
         de app te uploaden/updaten, en aan het einde van de run weer verwijderd. Er is dus geen
         permanente App Registration of los te beheren client secret nodig.
      5. Bestaat de Intune-app "Claude Desktop (Machine-wide)" nog niet? Dan wordt hij aangemaakt
         met detectie-/requirement-regels en toegewezen (Required) aan de opgegeven Entra-groep.
         Bestaat hij al en is er een nieuwere MSIX-versie EN/OF zijn Install-/Uninstall-/
         Detect-ClaudeDesktop-Intune.ps1 zelf gewijzigd sinds de vorige run? Dan wordt de
         package-inhoud bijgewerkt (Update-IntuneWin32AppPackageFile) én de detectieregel
         opnieuw gezet — de bestaande toewijzing blijft ongewijzigd staan en apparaten krijgen
         de nieuwe content automatisch gepusht.
      6. Zijn zowel de MSIX-versie als de drie content-scripts ongewijzigd t.o.v. de vorige run?
         Dan gebeurt er niets.

    Detectie is bewust NIET versie-specifiek (zie Detect-ClaudeDesktop-Intune.ps1) — Intune
    herinstalleert een Win32-app op reeds-toegewezen apparaten zodra de content-versie in Intune
    wijzigt, ongeacht wat de detectieregel teruggeeft. Je hoeft dus nooit meer handmatig een
    $MinimumVersion op te hogen.

    Omdat "content wijzigt" hierboven niet alleen de MSIX raakt: een SHA256-hash van de drie
    .ps1-content-scripts wordt naast ClaudeMsixVersion bewaard in het Notes-veld van de
    Intune-app (ScriptsHash=...). Zo triggert ook een pure code-wijziging in dit script-drietal
    (bv. de Install-script restart-fix) een nieuwe Update-IntuneWin32AppPackageFile-run, ook als
    de MSIX-versie zelf niet is veranderd — zonder dat je zelf hoeft te onthouden wanneer dat
    nodig is.

    Vereiste rol tijdens het draaien van dit script: Global Administrator, of Application
    Administrator in combinatie met een rol die AppRoleAssignment.ReadWrite.All-consent mag geven
    (zelfde vereiste als bij de tijdelijke App Registration in Remove-SharePointFileVersionsByDate.ps1).

.PARAMETER AssignmentGroupName
    Displaynaam van een bestaande Entra ID-groep. Alleen gebruikt bij de allereerste aanmaak van
    de Intune-app (Required-toewijzing); latere maandelijkse content-updates raken de toewijzing
    niet aan.

.PARAMETER WorkingDirectory
    Werkmap voor download/build-artefacten (Source/, Output/). Wordt bij elke run leeggemaakt en
    opnieuw gevuld. Standaard: C:\Temp\ClaudeDeploy.

.PARAMETER AppDisplayName
    Displaynaam van de Win32-app in Intune. Wordt gebruikt om de bestaande app terug te vinden bij
    latere runs — niet wijzigen zonder de app ook in Intune zelf te hernoemen.

.PARAMETER MsixDownloadUrl
    Override voor de MSIX-downloadlocatie. Standaard de officiële x64 "latest"-redirect van
    Anthropic.

.PARAMETER MinimumSupportedWindowsRelease
    Requirement rule: minimale Windows-release (bv. 'W10_21H2', 'W11_23H2'). Zie
    New-IntuneWin32AppRequirementRule voor geldige waarden van de geïnstalleerde module-versie.

.PARAMETER TenantId
    Entra ID tenant ID. Standaard automatisch bepaald uit de delegated sessie.

.PARAMETER IntuneWinAppUtilPath
    Pad naar een al aanwezige IntuneWinAppUtil.exe. Zonder opgave downloadt de IntuneWin32App-
    module deze automatisch naar %TEMP%.

.PARAMETER Force
    Slaat de "typ JA om door te gaan"-bevestiging over voordat er iets in Intune wordt
    aangemaakt/bijgewerkt. Gebruik dit als je het script onbeheerd (bv. via een geplande taak)
    wilt laten draaien.

.EXAMPLE
    .\Deploy-ClaudeDesktopIntune.ps1 -AssignmentGroupName "SG-Apps-ClaudeDesktop"

.EXAMPLE
    .\Deploy-ClaudeDesktopIntune.ps1 -AssignmentGroupName "SG-Apps-ClaudeDesktop" -Force
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$AssignmentGroupName,

    [string]$WorkingDirectory = 'C:\Temp\ClaudeDeploy',

    [string]$AppDisplayName = 'Claude Desktop (Machine-wide)',

    [string]$MsixDownloadUrl = 'https://claude.ai/api/desktop/win32/x64/msix/latest/redirect',

    [string]$MinimumSupportedWindowsRelease = 'W10_21H2',

    [string]$TenantId,

    [string]$IntuneWinAppUtilPath,

    [switch]$Force
)

# ── Cleanup tracking ─────────────────────────────────────────────────────────
$script:TempAppObjectId  = $null
$script:ConnectedHere    = $false

# ── Cultuur tijdelijk op invariant zetten ────────────────────────────────────
# De IntuneWin32App-module heeft een bekende, nog niet gemergede bug (upstream issue #210):
# de token-vervaldatum wordt met .ToString()/.Parse() weggeschreven en teruggelezen zonder
# expliciete cultuur op te geven. Op een systeem met dd/MM/jjjj-notatie (zoals nl-NL) crasht
# dat zodra de module de datum terugleest ("was not recognized as a valid DateTime"), wat
# elke retry na de eerste laat mislukken op een heel andere fout dan de eigenlijke Graph-fout.
# Invariant culture (M/d/jjjj) is intern consistent, dus dit voorkomt de crash zonder dat de
# module zelf gepatcht hoeft te worden.
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

function Get-MsixVersion {
    param([Parameter(Mandatory = $true)][string]$Path)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entry = $zip.Entries | Where-Object { $_.FullName -eq 'AppxManifest.xml' }
        if (-not $entry) { throw "AppxManifest.xml niet gevonden in $Path." }
        $stream = $entry.Open()
        try {
            $reader = New-Object System.IO.StreamReader($stream)
            try {
                [xml]$manifest = $reader.ReadToEnd()
            } finally {
                $reader.Dispose()
            }
        } finally {
            $stream.Dispose()
        }
        return [string]$manifest.Package.Identity.Version
    } finally {
        $zip.Dispose()
    }
}

function Get-ScriptsHashHex {
    # SHA256 over de gecombineerde inhoud van de content-scripts (Install/Uninstall/Detect),
    # zodat een pure code-wijziging in een van die drie ook zonder MSIX-versiebump wordt
    # gedetecteerd als "content gewijzigd" — zie .DESCRIPTION hierboven.
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

function New-TempAppRegistration {
    # Zelfde patroon als Remove-SharePointFileVersionsByDate.ps1: een kortlevende App
    # Registration + service principal, met alleen het app-only Graph-recht dat nodig is voor
    # deze run (hier: DeviceManagementApps.ReadWrite.All voor de IntuneWin32App-module).
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

    # 1-daags secret: deze run is kort, de tijdelijke app (incl. secret) wordt aan het einde
    # van de run verwijderd door Remove-TempApp. Kortere levensduur = kleinere blootstelling
    # als opruimen om wat voor reden ooit niet lukt.
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
Write-Host '   Deploy-ClaudeDesktopIntune' -ForegroundColor Cyan
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

# ── MSIX downloaden ──────────────────────────────────────────────────────────
Write-Step "Nieuwste Claude Desktop MSIX downloaden"
$msixPath = Join-Path $sourceDir 'Claude.msix'
try {
    Invoke-WebRequest -Uri $MsixDownloadUrl -OutFile $msixPath -MaximumRedirection 5 -UseBasicParsing -ErrorAction Stop
} catch {
    Write-Host "  [ERROR] Download van MSIX mislukt: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
if (-not (Test-Path $msixPath) -or (Get-Item $msixPath).Length -lt 1MB) {
    Write-Host "  [ERROR] Gedownload bestand ziet er niet uit als een geldige MSIX (te klein of ontbreekt)." -ForegroundColor Red
    exit 1
}

$newVersion = Get-MsixVersion -Path $msixPath
Write-Info "[OK]   Gedownload: Claude Desktop $newVersion"

# ── Content-scripts naast de MSIX plaatsen ───────────────────────────────────
Copy-Item -Path (Join-Path $PSScriptRoot 'Install-ClaudeDesktop-Intune.ps1')   -Destination $sourceDir -Force
Copy-Item -Path (Join-Path $PSScriptRoot 'Uninstall-ClaudeDesktop-Intune.ps1') -Destination $sourceDir -Force
$detectScriptPath = Join-Path $PSScriptRoot 'Detect-ClaudeDesktop-Intune.ps1'

# Hash van de drie content-scripts, zodat een pure code-wijziging (zonder nieuwe MSIX-versie)
# ook als "content gewijzigd" wordt herkend verderop — zie .DESCRIPTION.
$newScriptsHash = Get-ScriptsHashHex -Paths @(
    (Join-Path $PSScriptRoot 'Install-ClaudeDesktop-Intune.ps1'),
    (Join-Path $PSScriptRoot 'Uninstall-ClaudeDesktop-Intune.ps1'),
    $detectScriptPath
)

# ── .intunewin bouwen ────────────────────────────────────────────────────────
Write-Step "Package bouwen (.intunewin)"
$packageParams = @{
    SourceFolder = $sourceDir
    SetupFile    = 'Install-ClaudeDesktop-Intune.ps1'
    OutputFolder = $outputDir
    # Zonder -Force slaat New-IntuneWin32AppPackage het bouwen stilzwijgend over (en geeft
    # geen object terug) als er al een .intunewin met dezelfde naam in $outputDir staat — dat
    # moet elke maandelijkse run juist wél opnieuw gebeuren, met de zojuist gedownloade MSIX.
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

    # ── Doelgroep opzoeken (alleen nodig bij eerste aanmaak, maar we falen liever nu snel dan
    #    pas nadat de tijdelijke App Registration en de Intune-app al zijn aangemaakt) ──
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

    # ── Tijdelijke App Registration ──────────────────────────────────────────
    Write-Step "Tijdelijke App Registration aanmaken (DeviceManagementApps.ReadWrite.All)"
    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $tempApp = New-TempAppRegistration -AppName "ClaudeDeploy-Temp-$ts" -UsedTenantId $effectiveTenantId
    $script:TempAppObjectId = $tempApp.AppObjectId

    # ── Verbinden met Intune Graph via de tijdelijke app (met propagatie-retry) ─────────────
    Write-Step "Verbinden met Intune (app-only, via tijdelijke App Registration)"
    # Connect-MSIntuneGraph meldt een mislukte token-aanvraag (bv. AADSTS7000215 doordat het
    # secret van de zojuist aangemaakte app nog niet is gerepliceerd) via Write-Warning, niet
    # via een terminating error — -ErrorAction Stop op zichzelf vangt dat dus niet af. Door
    # -WarningAction Stop mee te geven wordt zo'n waarschuwing wél een exception die de retry-
    # lus kan opvangen. Een echte vervolgcall (Get-IntuneWin32App) valideert bovendien dat er
    # ook daadwerkelijk een bruikbaar token is, niet alleen dat Connect-MSIntuneGraph zonder
    # fouten doorliep.
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

    # ── Bestaande app opzoeken ───────────────────────────────────────────────
    Write-Step "Intune Win32-app opzoeken: '$AppDisplayName'"
    $existingApps = @(Get-IntuneWin32App -DisplayName $AppDisplayName -ErrorAction SilentlyContinue)
    $existingApp = if ($existingApps.Count -gt 0) { $existingApps[0] } else { $null }

    $existingVersion = $null
    if ($existingApp -and $existingApp.notes -match 'ClaudeMsixVersion=([0-9.]+)') {
        $existingVersion = $Matches[1]
    }

    $notes = "ClaudeMsixVersion=$newVersion; LastUpdated=$(Get-Date -Format o); Deployed by Deploy-ClaudeDesktopIntune.ps1"

    if ($existingApp -and $existingVersion -eq $newVersion) {
        # ── Geen actie nodig ──────────────────────────────────────────────
        Write-Step "Al up-to-date"
        Write-Info "Intune-app '$AppDisplayName' staat al op versie $newVersion. Geen wijzigingen nodig." -ForegroundColor Green
    }
    elseif ($existingApp) {
        # ── Content-update op de bestaande app; toewijzing blijft ongewijzigd ──
        $existingVersionLabel = if ($existingVersion) { $existingVersion } else { 'onbekend' }
        Write-Step "Bestaande app bijwerken naar versie $newVersion (was: $existingVersionLabel)"
        if (-not (Confirm-Action "Package-inhoud van '$AppDisplayName' (ID $($existingApp.id)) bijwerken in PRODUCTIE-Intune?")) {
            Write-Host "  Afgebroken door gebruiker." -ForegroundColor Yellow
            exit 1
        }
        Update-IntuneWin32AppPackageFile -ID $existingApp.id -FilePath $intuneWinFile -ErrorAction Stop | Out-Null
        Set-IntuneWin32App -ID $existingApp.id -AppVersion $newVersion -Notes $notes -ErrorAction Stop | Out-Null
        Write-Info "[OK]   App bijgewerkt naar versie $newVersion." -ForegroundColor Green
    }
    else {
        # ── Eerste aanmaak ────────────────────────────────────────────────
        Write-Step "Nieuwe Intune Win32-app aanmaken (eerste run)"
        if (-not (Confirm-Action "Nieuwe app '$AppDisplayName' aanmaken in PRODUCTIE-Intune en toewijzen (Required) aan groep '$($assignmentGroup.DisplayName)'?")) {
            Write-Host "  Afgebroken door gebruiker." -ForegroundColor Yellow
            exit 1
        }

        $detectionRule = New-IntuneWin32AppDetectionRuleScript -ScriptFile $detectScriptPath -EnforceSignatureCheck $false -RunAs32Bit $false
        $requirementRule = New-IntuneWin32AppRequirementRule -Architecture 'x64' -MinimumSupportedWindowsRelease $MinimumSupportedWindowsRelease

        $sysnativePwsh = '%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe'
        $installCommandLine   = "$sysnativePwsh -ExecutionPolicy Bypass -File Install-ClaudeDesktop-Intune.ps1"
        $uninstallCommandLine = "$sysnativePwsh -ExecutionPolicy Bypass -File Uninstall-ClaudeDesktop-Intune.ps1"

        $newApp = Add-IntuneWin32App `
            -FilePath              $intuneWinFile `
            -DisplayName           $AppDisplayName `
            -Description           'Claude Desktop, machine-breed geinstalleerd via Add-AppxProvisionedPackage (incl. Cowork-vereiste: VirtualMachinePlatform). Wordt maandelijks bijgewerkt door Deploy-ClaudeDesktopIntune.ps1.' `
            -Publisher              'Anthropic' `
            -AppVersion             $newVersion `
            -Notes                  $notes `
            -InstallCommandLine     $installCommandLine `
            -UninstallCommandLine   $uninstallCommandLine `
            -InstallExperience      'system' `
            -RestartBehavior        'suppress' `
            -DetectionRule          $detectionRule `
            -RequirementRule        $requirementRule `
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
Write-Host "   Klaar — Claude Desktop $newVersion" -ForegroundColor Cyan
Write-Host "  ================================================`n" -ForegroundColor Cyan
