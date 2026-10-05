[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../../readme.nl.md) › [scripts](../../../readme.nl.md) › [Intune](../../readme.nl.md) › [Desktop](../readme.nl.md) › **ClaudeDesktop**

# ClaudeDesktop

Machinebrede Intune-uitrol van [Claude Desktop](https://claude.com/download) voor Windows. Voer **één script eens per maand** uit om de Intune-app actueel te houden — geen blijvende App Registration of client secret om te beheren.

Gebaseerd op [Deploy Claude Desktop for Windows](https://support.claude.com/en/articles/12622703-deploy-claude-desktop-for-windows) en [Enterprise configuration for Claude Desktop](https://support.claude.com/en/articles/12622667-enterprise-configuration-for-claude-desktop).

> **De Windows-vereisten voor Cowork (VirtualMachinePlatform, Snel opstarten) zijn een aparte, onafhankelijke Win32-app** — zie [`../CoworkPrerequisites/readme.md`](../CoworkPrerequisites/readme.nl.md). Rol beide uit als je Cowork wilt; rol alleen deze uit als je alleen Claude Desktop zelf wilt. **Standaard is er geen Intune-afhankelijkheid** tussen de twee: een mislukte Cowork-vereiste mag Claude Desktop nooit blokkeren (dat werkt prima zonder Cowork), en een installatieprobleem van Claude Desktop mag nooit worden verward met een probleem met een Windows-onderdeel — elke app krijgt in Intune een eigen, apart zichtbare installatiestatus. Behandelt je omgeving Cowork als harde eis in plaats van optioneel, geef dan `-RequireCoworkPrerequisites` mee aan *dit* uitrolscript om in plaats daarvan een echte Intune-afhankelijkheid toe te voegen — zie "Optioneel: Cowork Prerequisites verplicht stellen" hieronder.

---

## Waarom de MSIX niet gewoon als line-of-business-app uploaden?

Intune installeert LOB-MSIX-apps per gebruiker. Dat mislukt voor standaardgebruikers zonder beheerdersrechten. In plaats daarvan wordt `Add-AppxProvisionedPackage` verpakt als Win32-app, precies zoals het officiële artikel aanbeveelt.

## Inhoud

| Script | Rol in Intune |
|---|---|
| [`Deploy-ClaudeDesktopIntune.ps1`](Deploy-ClaudeDesktopIntune.ps1) ([docs](#deploy-claudedesktopintuneps1)) | **Het script dat je uitvoert.** Regelt alles hieronder — zie de sectie "Maandelijkse run". |
| [`Install-ClaudeDesktop-Intune.ps1`](Install-ClaudeDesktop-Intune.ps1) ([docs](#install---uninstall---detect-claudedesktop-intuneps1)) | Contentscript voor de installatieopdracht |
| [`Uninstall-ClaudeDesktop-Intune.ps1`](Uninstall-ClaudeDesktop-Intune.ps1) ([docs](#install---uninstall---detect-claudedesktop-intuneps1)) | Contentscript voor de verwijderopdracht |
| [`Detect-ClaudeDesktop-Intune.ps1`](Detect-ClaudeDesktop-Intune.ps1) ([docs](#install---uninstall---detect-claudedesktop-intuneps1)) | Aangepast detectiescript |

`Install-`/`Uninstall-`/`Detect-ClaudeDesktop-Intune.ps1` voer je nooit handmatig uit — `Deploy-ClaudeDesktopIntune.ps1` verpakt ze automatisch in de `.intunewin` (of uploadt het detectiescript als onderdeel van de detectieregel).

## Deploy-ClaudeDesktopIntune.ps1

Wat het doet:

1. Downloadt de nieuwste x64-MSIX van Claude Desktop via de officiële "latest"-redirect-URL van Anthropic.
2. Leest de versie uit `AppxManifest.xml` in de MSIX.
3. Kopieert de installatie-/verwijderscripts naast de MSIX en bouwt een `.intunewin`-pakket (module `IntuneWin32App` — `IntuneWinAppUtil.exe` wordt automatisch gedownload als die er nog niet is).
4. Maakt gedelegeerd verbinding met Microsoft Graph (interactieve aanmelding) en maakt een kortlevende **tijdelijke App Registration** aan — hetzelfde patroon als [`Remove-SharePointFileVersionsByDate.ps1`](../../../Reporting/readme.nl.md) — met alleen de applicatiemachtiging `DeviceManagementApps.ReadWrite.All`. Die wordt gebruikt om de module `IntuneWin32App` te authenticeren en aan het eind van de run weer verwijderd. Tussen runs blijft niets bestaan behalve de Intune-app zelf.
5. Eerste run: maakt in Intune de Win32-app "Claude Desktop (Machine-wide)" aan met detectie- en vereistenregels, en wijst die als **Required** toe aan de Entra ID-groep die je meegeeft.
6. Latere runs: pusht een bijgewerkt pakket via `Update-IntuneWin32AppPackageFile` (bestaande toewijzing blijft ongemoeid, apparaten krijgen gewoon de nieuwe content) als **ofwel** de gedownloade MSIX-versie nieuwer is, **ofwel** de scripts Install-/Uninstall-/Detect-ClaudeDesktop-Intune.ps1 zelf sinds de vorige run zijn gewijzigd — beide bijgehouden in het veld Notes van de app (`ClaudeMsixVersion=...; ScriptsHash=...`), dus geen lokaal statusbestand nodig. Detectie- **en vereistenregels worden bij elke run opnieuw opgebouwd en ingediend**, niet alleen bij het aanmaken (zie "Bekend probleem" hieronder). Is de versie noch zijn de scripts gewijzigd, dan worden alleen de regels ververst.

Detectie is bewust versieonafhankelijk (aanwezigheid van het ingerichte pakket). Intune rolt een Win32-app opnieuw uit naar al getargete apparaten zodra de contentversie in Intune verandert, ongeacht wat de detectieregel meldt — er is dus geen `$MinimumVersion` die je elke maand met de hand moet ophogen, en een pure scriptwijziging (zonder nieuwe MSIX) leidt via de `ScriptsHash`-controle hierboven toch tot een nieuwe uitrol.

### Bekend probleem: installatiefouten 0x80070001 — met een veel diepere bug in de `IntuneWin32App`-module als oorzaak

Apparaten faalden bij de installatie met fout `0x80070001`, ongeacht het apparaat. Door het gecachte Win32-app-beleid op een getroffen apparaat (`AppWorkload.log`) te vergelijken met alle andere Win32-apps die aan dezelfde tenant waren toegewezen, kwam de afwijking boven: bij elke andere app hadden de `RequirementRules` de waarde `RequiredOSArchitecture: 3`, maar bij Claude Desktop stond er `RequiredOSArchitecture: 32` — een waarde die geen enkele andere app in de tenant gebruikte, en 16x wat `-Architecture 'x64'` zou moeten opleveren. Omdat dit op het app-object in Intune is opgeslagen (niet per apparaat), verklaarde dat waarom de fout op elk apparaat 100% reproduceerbaar was en geen apparaatspecifieke corruptie.

**Eerste poging tot oplossing (onvolledig):** `-RequirementRule` bij elke update opnieuw indienen, niet alleen bij het aanmaken. Dat bleek niet echt te werken, vanwege een tweede, veel ernstigere bug:

**De echte hoofdoorzaak**, bevestigd aan de hand van de eigen broncode van de module `IntuneWin32App` (v1.5.0) en het door Microsoft gedocumenteerde schema van de [`win32LobApp`-resource](https://learn.microsoft.com/en-us/graph/api/resources/intune-apps-win32lobapp):
- `applicableArchitectures` / `allowedArchitectures` / `minimumSupportedWindowsRelease` zijn **platte eigenschappen op het hoogste niveau** van `win32LobApp` — ze hebben helemaal geen `@odata.type` nodig (dat is alleen vereist voor de polymorfe collectie `rules`: detectie- of vereistenregels op basis van bestand/register/productcode/script).
- `New-IntuneWin32AppRequirementRule` zet nooit `@odata.type` op het object dat het teruggeeft (terecht — het heeft er geen nodig).
- Maar `Set-IntuneWin32App` (alleen de **update**-cmdlet — `Add-IntuneWin32App`, gebruikt bij het aanmaken, heeft volledig andere, correcte interne logica) eist ten onrechte toch `@odata.type` op `-RequirementRule`, en als die ontbreekt, volgt een `Write-Warning "...missing required '@odata.type'..."` met daarna een kale `break`.
- Die `break` staat niet binnen een lus of `switch` — empirisch geverifieerd: hij beëindigt **de hele rest van de functie**, inclusief de eigenlijke Graph PATCH-aanroep verderop. Dat betekent: elke `Set-IntuneWin32App`-aanroep met `-RequirementRule` deed stilletjes **helemaal niets** — niet de architectuurcorrectie, niet `-Notes`, niet `-DetectionRule`, niet `-Icon`, niet `-CompanyPortalFeaturedApp`. Alleen `Update-IntuneWin32AppPackageFile` (een aparte aanroep, vlak daarvoor) bereikte bij elke run daadwerkelijk Intune.

**De echte oplossing**: de update-tak geeft `-RequirementRule` helemaal niet meer door aan `Set-IntuneWin32App` (zodat Notes/AppVersion/DetectionRule/Icon/CompanyPortalFeaturedApp gewoon doorkomen), en roept in plaats daarvan `Set-Win32AppArchitectureRequirement` aan — een kleine helper die `applicableArchitectures`/`allowedArchitectures`/`minimumSupportedWindowsRelease` rechtstreeks via Microsoft Graph PATCHt (met dezelfde `$Global:AuthenticationHeader`-sessie die `Connect-MSIntuneGraph` al heeft opgezet), en zo voor alleen dat stuk het kapotte cmdlet-pad volledig omzeilt. `Add-IntuneWin32App` (het aanmaken) wordt hier helemaal niet door geraakt en blijft `-RequirementRule` normaal gebruiken.

**Faalt een apparaat na die oplossing nog steeds**: controleer dan of het niet eigenlijk vastzit achter de losstaande **GRS-retry-cooldown** van Intune (3 mislukte pogingen → 24 uur blokkade, ongeacht de appconfiguratie) — zie de [GRS-cooldownscripts](../../readme.nl.md#repair-stuckwin32appenforcementps1) een niveau hoger, ofwel de handmatige versie op het apparaat, ofwel het Detect-/Remediate-paar dat volledig via de Intune-portal draait. Dit raakt elke Win32-app op dat apparaat, niet alleen Claude, dus het is een nuttige eerste controle als ook meerdere losstaande apps stilletjes vastzitten.

### Zichtbaarheid in de Bedrijfsportal

Het uitrolscript haalt het logo van de app rechtstreeks uit de gedownloade MSIX (`Properties/Logo` in `AppxManifest.xml`, met als terugval de variant met de hoogste schaal die echt in het pakket zit) en stelt dat in als pictogram van de Win32-app, plus `-CompanyPortalFeaturedApp $true` bij elke run — zodat gebruikers in de Bedrijfsportal niet een generiek Win32-pictogram ergens in de volledige applijst zien, maar het echte Claude-logo, uitgelicht.

### Vereiste rol

Global Administrator, of Application Administrator in combinatie met een rol die toestemming voor `AppRoleAssignment.ReadWrite.All` kan verlenen — dezelfde eis als voor de tijdelijke App Registration in de SharePoint-rapportagescripts.

### Maandelijkse run

```powershell
.\Deploy-ClaudeDesktopIntune.ps1 -AssignmentGroupName "SG-Apps-ClaudeDesktop"
```

Je krijgt een interactieve aanmeldprompt en een bevestiging "type JA to continue" voordat er iets in Intune wordt aangemaakt of bijgewerkt. Voeg `-Force` toe om de bevestiging over te slaan (bijv. voor een geplande/onbeheerde run zodra je vertrouwd bent met de werkwijze).

| Parameter | Standaard | Omschrijving |
|---|---|---|
| `-AssignmentGroupName` | *(verplicht)* | Entra ID-groep die als Required wordt toegewezen. Alleen gebruikt bij het aanmaken. |
| `-WorkingDirectory` | `C:\Temp\ClaudeDeploy` | Staging-map voor download/build (`Source/`, `Output/`) |
| `-AppDisplayName` | `Claude Desktop (Machine-wide)` | Wordt gebruikt om de bestaande app bij latere runs te vinden — niet wijzigen zonder de app ook in Intune te hernoemen |
| `-MsixDownloadUrl` | Officiële x64-"latest"-redirect van Anthropic | Overschrijven om te testen |
| `-MinimumSupportedWindowsRelease` | `W10_21H2` | Vereistenregel |
| `-TenantId` | automatisch gedetecteerd | Tenant-ID van Entra ID |
| `-IntuneWinAppUtilPath` | automatische download | Gebruik een al gedownloade `IntuneWinAppUtil.exe` |
| `-Force` | uit | Sla de bevestigingsprompt(s) over |
| `-RequireCoworkPrerequisites` | uit | Voeg een echte Intune-afhankelijkheid van de app Cowork Prerequisites toe — zie hieronder |
| `-CoworkPrerequisitesAppDisplayName` | `Cowork Windows Prerequisites (Machine-wide)` | Moet overeenkomen met de `-AppDisplayName` die in `Deploy-CoworkPrerequisitesIntune.ps1` wordt gebruikt. Alleen gebruikt met `-RequireCoworkPrerequisites` |

### Optioneel: Cowork Prerequisites verplicht stellen

Standaard zijn Claude Desktop en Cowork Prerequisites onafhankelijk — er is geen Intune-afhankelijkheid tussen de twee (zie de opmerking bovenaan deze readme voor het waarom). Heeft je omgeving echt nodig dat Cowork altijd aanwezig is — een apparaat zonder werkende Cowork-vereisten mag Claude Desktop helemaal niet krijgen — geef dan `-RequireCoworkPrerequisites` mee:

```powershell
.\Deploy-ClaudeDesktopIntune.ps1 -AssignmentGroupName "SG-Apps-ClaudeDesktop" -RequireCoworkPrerequisites
```

Dit zoekt de app Cowork Prerequisites op (rol die **eerst** uit — `Deploy-CoworkPrerequisitesIntune.ps1`) en roept `Add-IntuneWin32AppDependency` aan met `DependencyType 'Detect'` (niet `'AutoInstall'`): de vereistenapp moet nog steeds zelfstandig als Required zijn toegewezen en al op het apparaat zijn gedetecteerd — deze afhankelijkheid installeert hem niet automatisch namens Claude Desktop, maar laat Intune er alleen op wachten. De afweging ten opzichte van de standaard: een apparaat dat vastzit op de Cowork-vereisten (bijv. midden in een herstart, of omdat het echt faalt) krijgt dan ook geen Claude Desktop tot dat is opgelost, in plaats van direct Claude Desktop en later Cowork.

## Install- / Uninstall- / Detect-ClaudeDesktop-Intune.ps1

De contentscripts van de Win32-app. `Deploy-ClaudeDesktopIntune.ps1` verpakt en uploadt ze; je voert ze nooit handmatig uit.

- **`Install-ClaudeDesktop-Intune.ps1`** (installatieopdracht): verwijdert elke bestaande Claude Desktop volledig, provisioneert de MSIX machinebreed met `Add-AppxProvisionedPackage` en schakelt de eigen auto-updater van Claude uit — zie de secties hieronder. Parameter `-MsixFileName` (standaard `Claude.msix`): bestandsnaam van de MSIX naast het script. Logt naar `%ProgramData%\ClaudeDeploy\install.log`.
- **`Uninstall-ClaudeDesktop-Intune.ps1`** (verwijderopdracht): verwijdert het machinebreed geprovisioneerde pakket en, als terugvaloptie, per-user Appx-installaties van profielen die al hebben aangemeld. Laat de Windows-vereisten van Cowork ongemoeid. Logt naar `%ProgramData%\ClaudeDeploy\uninstall.log`.
- **`Detect-ClaudeDesktop-Intune.ps1`** (aangepast detectiescript, 64-bit): meldt "installed" als het machinebreed geprovisioneerde pakket aanwezig is. Bewust versie-onafhankelijk, en probeert alleen opnieuw bij een tijdelijke DISM-exception.

### Beleid voor automatische updates

Het installatiescript zet `HKLM:\SOFTWARE\Policies\Claude\disableAutoUpdates = 1` (DWord). De eigen updater van Claude is uitgeschakeld, zodat het versiebeheer volledig bij deze maandelijkse Intune-run blijft in plaats van per apparaat uit de pas te lopen.

### Schone herinstallatie bij elke run

Voordat de nieuwe versie wordt ingericht, verwijdert het installatiescript Claude Desktop eerst volledig van het apparaat, in drie rondes:

1. Stopt elk draaiend Claude-proces.
2. Verwijdert elke **Appx**-installatie per gebruiker (`Get-AppxPackage -AllUsers` / `Remove-AppxPackage -AllUsers`, inclusief al aangemelde profielen), en daarna het oude machinebreed ingerichte pakket.
3. Verwijdert elke **klassieke (niet-Appx) installatie per gebruiker** — vooral de consumenteninstaller van [claude.ai/download](https://claude.ai/download), die zich registreert via een gewone Uninstall-registersleutel per gebruiker in plaats van als Appx-pakket, zodat `Get-AppxPackage` hem nooit ziet. Het script doorzoekt de Uninstall-registersleutel van elk lokaal profiel — ook van profielen die op dat moment niet zijn aangemeld, door tijdelijk hun `NTUSER.DAT` te laden — en voert van elke match de `QuietUninstallString` uit (of `UninstallString` als die ontbreekt), met een time-out van 120 s zodat een vastgelopen installer de Intune-installatie niet kan laten hangen.

Pas na alle drie de rondes wordt de nieuwe MSIX ingericht. Dit is bewust grondiger dan alleen de inrichtingslaag leegmaken — een installatie die via een andere route is achtergebleven, kan anders zijn eigen, onbeheerde Claude-sessie blijven draaien, ook nadat de machinebrede versie is bijgewerkt. Een gebruiker die Claude open heeft, verliest die sessie wanneer dit draait.

### Robuustheid op een apparaat waar het nooit geïnstalleerd was

Beide contentscripts zijn zo geschreven dat een apparaat waarop Claude nooit heeft gestaan — inclusief de allereerste Autopilot-ESP-run — er zonder problemen doorheen komt:

- **Installatiescript**: elke verwijderronde (proces stoppen, Appx, klassieke verwijdering per gebruiker) controleert eerst op een leeg/`$null`-resultaat, zodat een volledig schone machine bij elke stap alleen "none found" logt in plaats van een fout te geven.
- **Detectiescript**: een leeg/negatief resultaat van `Get-AppxProvisionedPackage` (de verwachte uitkomst op een apparaat waar het nooit geïnstalleerd was) meldt direct "not installed" (exit 1), zonder nieuwe poging — nieuwe pogingen starten alleen bij een echte **exception** (bijv. een tijdelijke DISM-vergrendeling, aannemelijk vlak na de zware Appx-activiteit van het installatiescript zelf op hetzelfde apparaat), zodat een echt schone machine nooit wordt opgehouden door nieuwe pogingen die de uitkomst toch niet kunnen veranderen.

## Vereisten

- De PowerShell-modules `Microsoft.Graph.Authentication`, `Microsoft.Graph.Applications`, `Microsoft.Graph.Groups` en `IntuneWin32App` — installeer ze met `.\scripts\Startup\Install-Modules.ps1`
- Uitvoeren vanaf Windows (de verpakkingstool en de MSIX-/AppX-cmdlets werken alleen op Windows)
