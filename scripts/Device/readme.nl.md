[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../readme.nl.md) › [scripts](../readme.nl.md) › **Device**

# Scripts voor apparaatbeheer

Scripts voor het beheren en onderhouden van Windows-endpoints. Alle scripts vereisen administratorrechten.

---

## Mappen

| Map | Omschrijving |
|--------|-------------|
| [`Time sync/`](Time%20sync/readme.nl.md) | Windows-tijdsynchronisatie herstellen door W32tm te herstarten en een geplande taak te registreren |
| [`audio/`](audio/readme.nl.md) | De interne microfoon op laptops detecteren en uitschakelen |
| [`DriveMapping/`](DriveMapping/readme.nl.md) | SharePoint-/OneDrive-documentbibliotheken bij aanmelden aan stationsletters koppelen |
| [`TempDisk/`](TempDisk/readme.nl.md) | De tijdelijke (ephemeral) schijf bij elke start herstellen als `D:` en de pagefile erop houden |

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`Clear-TempFiles.ps1`](Clear-TempFiles.ps1) ([docs](#clear-tempfilesps1)) | De gedeelde tijdelijke scriptmap leegmaken (`C:\Temp` op Windows, `/tmp` op Linux/macOS) |
| [`Invoke-WindowsActivation.ps1`](Invoke-WindowsActivation.ps1) ([docs](#invoke-windowsactivationps1)) | Windows activeren, productcodes en KMS-instellingen beheren |
| [`Invoke-WindowsCleanup.ps1`](Invoke-WindowsCleanup.ps1) ([docs](#invoke-windowscleanupps1)) | Terug te winnen schijfruimte scannen en vrijmaken |
| [`Remove-OemBloatware.ps1`](Remove-OemBloatware.ps1) ([docs](#remove-oembloatwareps1)) | OEM-bloatware (HP/Lenovo/Dell) en generieke Microsoft Store-bloatware verwijderen |
| [`Repair-AppxPackageStore.ps1`](Repair-AppxPackageStore.ps1) ([docs](#repair-appxpackagestoreps1)) | AppX-pakketten herstellen die falen met `0x80070490` — verweesde vermeldingen in de package store, en FSLogix die een versie terugzet die de host niet heeft (Teams, nieuwe Outlook, elk pakket) |
| [`Test-OpenVpnDiagnostics.ps1`](Test-OpenVpnDiagnostics.ps1) ([docs](#test-openvpndiagnosticsps1)) | Problemen met OpenVPN Connect diagnosticeren |
| [`Update-TeamsClient.ps1`](Update-TeamsClient.ps1) ([docs](#update-teamsclientps1)) | De nieuwe Teams + de Outlook-vergaderinvoegtoepassing bijwerken, alleen als Microsoft een nieuwere build heeft gepubliceerd ([hoe het werkt](Update-TeamsClient.md), [IT Glue](Update-TeamsClient-ITGlue.md)) |

---

### Clear-TempFiles.ps1

Maakt alleen de gedeelde tijdelijke scriptmap leeg.
Het standaarddoelpad is `C:\Temp` op Windows en `/tmp` op Linux/macOS.
Draait standaard als proefdraai en verwijdert alleen bestanden als `-Apply` is opgegeven.

**Parameters**

| Parameter | Omschrijving |
|-----------|-------------|
| `-Apply` | Echt verwijderen (standaard is proefdraai) |
| `-TempPath` | Ander tijdelijk doelpad gebruiken |
| `-OlderThanDays` | Alleen items ouder dan N dagen (standaard: `1`) |

**Voorbeelden**

```powershell
# Proefdraai
.\Clear-TempFiles.ps1

# Alleen tijdelijke bestanden ouder dan 7 dagen verwijderen
.\Clear-TempFiles.ps1 -Apply -OlderThanDays 7

# Ander tijdelijk pad gebruiken (voorbeeld Linux/macOS)
.\Clear-TempFiles.ps1 -Apply -TempPath /var/tmp
```

---

## Invoke-WindowsActivation.ps1

Activeer Windows of beheer licentie-instellingen vanaf de opdrachtregel.

**Parameters**

| Parameter | Omschrijving |
|-----------|-------------|
| `-Status` | Huidige activeringsstatus tonen (WMI + `slmgr /dli`) |
| `-ProductKey` | Een productcode installeren (retail of generieke KMS-code) |
| `-KmsServer` | Het adres van de KMS-activeringsserver instellen |
| `-KmsPort` | Poort van de KMS-server (standaard: 1688) |
| `-Activate` | Activering starten bij Microsoft of bij de ingestelde KMS-server |
| `-RemoveKey` | De geïnstalleerde productcode verwijderen (vóór herinstallatie / licentieoverdracht) |
| `-ReArm` | De teller van de respijtperiode resetten (max. ~3-5x per Windows-installatie) |
| `-Force` | Bevestigingsvragen overslaan |

**Voorbeelden**

```powershell
# Huidige activeringsstatus controleren
.\Invoke-WindowsActivation.ps1 -Status

# Een retailcode installeren en online activeren
.\Invoke-WindowsActivation.ps1 -ProductKey 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX' -Activate

# Naar een KMS-server van het bedrijf wijzen en activeren
.\Invoke-WindowsActivation.ps1 -KmsServer 'kms.company.local' -Activate

# Code verwijderen vóór herinstallatie
.\Invoke-WindowsActivation.ps1 -RemoveKey -Force
```

---

## Invoke-WindowsCleanup.ps1

Scant terug te winnen schijfruimte en maakt die desgewenst vrij. Draait standaard als proefdraai — zonder `-Apply` wordt er niets verwijderd.

**Wat er wordt opgeruimd**

| Categorie | Details |
|----------|---------|
| Tijdelijke bestanden | Temp van gebruikers (alle profielen) + `C:\Windows\Temp` |
| Windows Update | `SoftwareDistribution\Download` + DeliveryOptimization |
| Prefetch | `C:\Windows\Prefetch` |
| Geheugendumps | Minidumps van het systeem + CrashDumps per gebruiker |
| WER | Wachtrijen van Windows Error Reporting (systeem + per gebruiker) |
| Cache | Miniatuurcache, DirectX-shadercache, lettertypecache |
| Prullenbak | Alle stations |
| Browsercache | Edge, Chrome (meerdere profielen), Firefox — alle gebruikersprofielen |
| Eventlogs | Alle Windows-eventlogs |
| App- en systeemlogs | Dynamische scan van heel C:\ op mappen `logs`/`log`/`logging` |
| DISM | Opschonen van de component store (`/StartComponentCleanup /ResetBase`) |

**Parameters**

| Parameter | Omschrijving |
|-----------|-------------|
| `-Apply` | De opschoning echt uitvoeren (standaard: alleen proefdraai) |
| `-SkipBrowserCache` | Browsercache overslaan |
| `-SkipAppLogs` | Opschonen van applicatie- en systeemlogs overslaan |
| `-SkipEventLogs` | Leegmaken van de Windows-eventlogs overslaan |
| `-SkipDism` | Opschonen van de DISM-component store overslaan (traag) |
| `-SkipRecycleBin` | Legen van de prullenbak overslaan |
| `-OutputPath` | Andere map voor het rapport (standaard: `C:\Temp\`) |

**Voorbeelden**

```powershell
# Proefdraai — zien hoeveel ruimte er vrij kan komen
.\Invoke-WindowsCleanup.ps1

# Volledige opschoning
.\Invoke-WindowsCleanup.ps1 -Apply

# Opschonen, browsercache en DISM overslaan
.\Invoke-WindowsCleanup.ps1 -Apply -SkipBrowserCache -SkipDism
```

Na elke run wordt een CSV-rapport met resultaten per categorie opgeslagen in `C:\Temp\`.

---

## Remove-OemBloatware.ps1

Detecteert de fabrikant van het apparaat en verwijdert bekende OEM-bloatware via `winget`, plus een generieke lijst consumentenapps uit de Microsoft Store (Xbox, Solitaire, Bing News/Weather, Cortana, Clipchamp enz.) via `Remove-AppxPackage`. Draait standaard als proefdraai — zonder `-Apply` wordt er geen app verwijderd.

> De bloatwarelijsten zijn een startpunt, niet volledig — pakket-ID's verschillen per OEM-preload-image en veranderen in de loop van de tijd. Draai eerst `winget list` / `Get-AppxPackage | Select Name` op een representatief apparaat en pas zo nodig de lijsten in het script aan.

**Parameters**

| Parameter | Omschrijving |
|-----------|-------------|
| `-Apply` | Gevonden apps echt verwijderen (standaard: alleen proefdraai) |
| `-SkipOem` | Fabrikantspecifieke verwijdering overslaan, alleen de generieke Microsoft Store-lijst verwerken |
| `-SkipAppx` | De generieke Microsoft Store-lijst overslaan, alleen fabrikantspecifieke apps verwerken |
| `-Manufacturer` | Automatische detectie overschrijven (`HP`, `Lenovo`, `Dell`) |
| `-OutputPath` | Andere map voor het rapport (standaard: `C:\Temp\`) |

**Voorbeelden**

```powershell
# Proefdraai — zien wat er op dit apparaat zou worden verwijderd
.\Remove-OemBloatware.ps1

# OEM- + generieke bloatware echt verwijderen
.\Remove-OemBloatware.ps1 -Apply

# Alleen generieke Microsoft Store-rommel verwijderen, OEM-apps met rust laten
.\Remove-OemBloatware.ps1 -Apply -SkipOem
```

Na elke run wordt een CSV-rapport (gevonden/verwijderd per app) opgeslagen in `C:\Temp\`.

**Vereist:** `winget` (App Installer uit de Microsoft Store) voor het verwijderen van OEM-pakketten; administratorrechten.

---

## Repair-AppxPackageStore.ps1

Herstelt AppX-pakketten die falen met `0x80070490` ("Element not found"), meestal gelogd als:

```
MSTeams installation error: Deployment Register operation with target volume C: on Package
MSTeams_26246.1604.5133.838_x64__8wekyb3d8bbwe from:  (AppxManifest.xml)  failed with error 0x80070490.
```

Het lege pad vóór `(AppxManifest.xml)` verraadt het: Windows speelt een registratie opnieuw af die geen installatielocatie heeft. Er zijn twee oorzaken, en het script diagnosticeert ze allebei:

- **De package store van de host is inconsistent.** `AppxAllUserStore` vermeldt een pakket voor een SID zonder profiel, een pakket waarvan de bestanden weg zijn, of een machinebrede vermelding zonder manifest. `Remove-AppxPackage`, DISM en installers lezen allemaal die store, dus ze falen allemaal, en de host drainen of herstarten verandert niets.
- **FSLogix zet een versie terug die de host niet heeft** (eventbron `Apps (Microsoft-FSLogix-Apps)`, bij aanmelden). Bij afmelden slaat FSLogix de pakketten van de gebruiker op volledige naam — exacte versie — op in `AppxPackages.xml` in de profielcontainer, en registreert ze bij de volgende aanmelding opnieuw (`HKLM\SOFTWARE\FSLogix\Profiles\InstallAppxPackages`, standaard aan). Een host met een andere build, of helemaal geen, antwoordt met `0x80070490`. De oplossing is het pakket voor alle gebruikers te **provisionen**, met dezelfde build op elke host in de pool, en een FSLogix-build te draaien die op familienaam registreert (2210 HF4 voor Teams, 25.06 voor nieuwe Outlook). `AppxPackages.xml` wordt niet bewerkt: Microsoft zegt dat dat niet de bedoeling is, en FSLogix herschrijft het bij de volgende afmelding.

**Stappen**

| Stap | Wat het doet |
|------|--------------|
| 1. Diagnose | Geregistreerde pakketten waarvan de bestanden weg zijn (*Ghost*) of waarvan de status niet Ok is (*Damaged*), geprovisionde pakketten zonder bestanden, verweesde `AppxAllUserStore`-vermeldingen, recente AppX-deploymentfouten |
| 1b. FSLogix | FSLogix-build, `InstallAppxPackages`, ODFC `IncludeTeams`, de pakketten die FSLogix in de laatste `-Days` dagen niet kon registreren, afgezet tegen wat deze host provisiont, AppX-installatiebeleid |
| 1c. Falende apps | **Elk** pakket dat in de laatste `-Days` dagen niet kon installeren, bijwerken of registreren, uit het AppX-deploymentlog en het FSLogix-log samen: aantal, foutcodes met hun betekenis, de gevraagde versies en of deze host hun bestanden heeft. `0x80070490` eerst, top 15 |
| 2. Geprovisiond | `Remove-AppxProvisionedPackage` voor geprovisionde kopieën waarvan de bestanden weg zijn |
| 3. Opnieuw registreren | `Add-AppxPackage -Register` vanuit het eigen manifest van het pakket waar de bestanden er nog zijn. Een oudere versie naast een nieuwere van hetzelfde pakket is *Superseded*, niet beschadigd — in grijs gemeld en aan Windows overgelaten om te verwijderen, omdat opnieuw registreren alleen kan falen met `0x80073D06` |
| 4. Verwijderen | `Remove-AppxPackage -AllUsers` voor ghosts, per gebruiker waar dat wordt geweigerd |
| 5. Store | Elke overgebleven verweesde registersleutel wordt geëxporteerd naar een `.reg`-back-up en pas daarna verwijderd — geen back-up, geen verwijdering |
| 6. Provisionen | Waar FSLogix faalt op een nieuwere Teams-/Outlook-build dan deze host provisiont: **precies die build**, als MSIX van Microsofts CDN op de URL met versienummer, handtekening gecontroleerd, dan `Add-AppxProvisionedPackage` en teruggelezen. Anders `-Provision`: `teamsbootstrapper.exe -p` / Outlook `Setup.exe --provision true --quiet --start-`, gedownload van Microsoft en handtekening gecontroleerd. Met `-UseWinget` in plaats daarvan de MSIX uit winget (`Microsoft.Teams`, `Microsoft.Outlook`), geprovisiond met `Add-AppxProvisionedPackage` samen met eventuele afhankelijkheden die winget meebracht. `-WingetId`: hetzelfde voor elk ander pakket. `-Source`: elke MSIX die je zelf aanlevert |
| 7. Verificatie | De diagnose draait opnieuw; exitcode 1 als er iets heeft overleefd |

**Parameters**

| Parameter | Omschrijving |
|-----------|-------------|
| `-Name` | Pakketnamen, wildcards toegestaan (standaard `*`). Bijv. `MSTeams,Microsoft.OutlookForWindows`. Afkortingen: `teams`, `outlook`, `copilot` (= `-Copilot`) — `-Name outlook,copilot` is genoeg |
| `-CheckOnly` | Alleen diagnose, niets wijzigen (exitcode `2` als er werk is) |
| `-Provision` | Teams / nieuwe Outlook voor alle gebruikers provisionen met de installer van Microsoft: wat volgens FSLogix ontbreekt of achterloopt, plus elk van beide die expliciet in `-Name` staat |
| `-UseWinget` | Met `-Provision`: Teams / nieuwe Outlook uit winget halen in plaats van met de installer van Microsoft. winget controleert de SHA256, het script de Microsoft-handtekening. De manifests van winget lopen achter (gemeten: Teams 26198 vs 26246, Outlook 1.2026.812 vs 902); de run waarschuwt als de build ouder is dan wat de profielen vragen |
| `-WingetId` | winget-ID's van andere pakketten die op dezelfde manier voor alle gebruikers worden geprovisiond (alleen als het winget-manifest ervan een MSIX is) |
| `-Source` | Deze `.msix` / `.msixbundle` provisionen nadat de store schoon is |
| `-IncludeDeprovisioned` | Ook Deprovisioned-markeringen opruimen bij een wildcard-`-Name` (standaard worden ze alleen opgeruimd voor expliciet genoemde pakketten) |
| `-SkipSignatureCheck` | Geen geldige Microsoft-handtekening eisen op de installer of op `-Source` |
| `-Days` | Hoe ver terug de FSLogix Apps- en AppX-deploymentlogs worden gelezen (standaard `7`) |
| `-WorkingDir` | Downloadmap voor de installers (standaard `C:\IT\AppxRepair`) |
| `-LogPath` | Transcript en `.reg`-back-ups (standaard `C:\Temp`) |
| `-ComputerName` | Draai op deze sessiehosts in plaats van op deze machine (bijv. `lem-avd-4,lem-avd-5,lem-avd-6`): het script kopieert zichzelf via PowerShell remoting (WinRM) naar `C:\IT\AppxRepair` op elke host, draait daar met dezelfde parameters en eindigt met een pooltabel — exitcode, FSLogix-build, geprovisionde Teams / Outlook per host — die elk verschil tussen hosts benoemt. Een herstel wordt één keer voor de hele pool bevestigd |
| `-Credential` | Referenties voor die remotingsessies |
| `-Copilot` | Kijk naar Copilot (stap 1d): de Microsoft 365 Copilot-app (`Microsoft.MicrosoftOfficeHub`) en de Windows Copilot-app (`Microsoft.Copilot`), de **samengevoegde Microsoft Copilot-app** die Edge Update sinds september 2026 installeert, en elk beleid dat hem verwijdert of blokkeert. Hun Deprovisioned-markeringen vallen binnen de scope. Met `-Provision` wordt de **nieuwe** app machinebreed geïnstalleerd zoals Microsoft het documenteert: `Install{C50565E9-...}` = 5 (Force Installs), `UpdaterExperimentationAndConfigurationServiceControl` = 1 en `CopilotUnificationAllowed{...}` = 1 onder `HKLM\SOFTWARE\Policies\Microsoft\EdgeUpdate` (na een `.reg`-back-up), Edge Update wordt gevraagd nu te controleren, en de run wacht tot 10 minuten op de app. Verschijnt hij niet, dan is `M365CopilotDesktopInstaller.exe --quiet --start -p` (de oude app, die de samenvoeging overzet) de terugvaloptie. Een beleid dat de installatie verbiedt (`Install` = 0) wordt nooit overschreven |
| `-Latest` | De **nieuwste** build van Teams / Outlook binnen het bereik provisionen, niet alleen die waarop FSLogix faalde. Teams: de configuratieservice van Microsoft (de feed die de client zelf gebruikt). Voor Outlook bestaat zo'n feed niet — de Store-catalogus meldde 1.2026.818.0 terwijl 915.300 al uit was — dus wordt de nieuwste build gebruikt die aantoonbaar bestaat: de nieuwste waar FSLogix om vroeg, die voor een gebruiker op de host is geregistreerd, of die in `WindowsApps` staat. Gebruik met `-Provision` |

Ondersteunt `-WhatIf` en `-Confirm`; vraagt per wijziging tenzij `-Confirm:$false`. NinjaOne-scriptvariabelen: `packageName`, `checkOnly`, `provision`, `useWinget`, `wingetId`, `source`, `includeDeprovisioned`, `skipSignatureCheck`, `days`, `workingDir`, `logPath`.

**Voorbeelden**

```powershell
# Wat is er stuk op deze host? Wijzigt niets.
.\Repair-AppxPackageStore.ps1 -Name MSTeams,Microsoft.OutlookForWindows -CheckOnly

# Herstellen en Teams + nieuwe Outlook terugzetten voor alle gebruikers, onbeheerd
.\Repair-AppxPackageStore.ps1 -Name MSTeams,Microsoft.OutlookForWindows -Provision -Confirm:$false

# De hele pool vanaf één plek: diagnose, daarna herstel + provisionen op elke host
.\Repair-AppxPackageStore.ps1 -ComputerName lem-avd-4,lem-avd-5,lem-avd-6 -Name MSTeams,Microsoft.OutlookForWindows -CheckOnly
.\Repair-AppxPackageStore.ps1 -ComputerName lem-avd-4,lem-avd-5,lem-avd-6 -Name MSTeams,Microsoft.OutlookForWindows -Provision -Confirm:$false

# Nieuwe Outlook en de nieuwe Copilot-app op de pool: waarom ze ontbreken, daarna installeren voor alle gebruikers
.\Repair-AppxPackageStore.ps1 -ComputerName lem-avd-4,lem-avd-5,lem-avd-6 -Name outlook,copilot -CheckOnly
.\Repair-AppxPackageStore.ps1 -ComputerName lem-avd-4,lem-avd-5,lem-avd-6 -Name outlook,copilot -Provision -Confirm:$false

# De allernieuwste Teams en Outlook op de pool, niet alleen de build waar FSLogix om vraagt
.\Repair-AppxPackageStore.ps1 -ComputerName lem-avd-4,lem-avd-5,lem-avd-6 -Name teams,outlook -Latest -Provision -Confirm:$false

# Hetzelfde, met beide pakketten uit winget
.\Repair-AppxPackageStore.ps1 -Name MSTeams,Microsoft.OutlookForWindows -Provision -UseWinget -Confirm:$false

# Elke app die in de laatste 14 dagen faalde, plus de hele store; wijzigt niets
.\Repair-AppxPackageStore.ps1 -CheckOnly -Days 14
```

**Opmerkingen**

- Exitcodes: `0` schoon, `1` mislukt of er heeft iets overleefd, `2` check-only vond werk.
- Draai het op een FSLogix-pool op **elke** host — `-ComputerName` doet dat vanaf één plek en laat zien of FSLogix, Teams en Outlook overal gelijk zijn.
- Profielen vragen om een nieuwere Teams / Outlook dan een host provisiont, omdat beide apps zichzelf per gebruiker bijwerken, terwijl de installers van Microsoft een oudere, laatst bekende goede build provisionen (gemeten: Teams 26225 vs 26246, Outlook 1.2026.818 vs 902 en 915). Waar FSLogix op die nieuwere build **faalt**, zet `-Provision` precies die build neer vanaf Microsofts CDN. Een productiehost liet zien dat dit ook nodig is als de bestanden op schijf staan: Outlook faalde 186× met `0x80070490` terwijl beide gevraagde builds aanwezig waren. Omdat de apps blijven bijwerken, draai je het script volgens een schema (bijv. dagelijks vanuit NinjaOne met `-Name teams,outlook -Provision -Confirm:$false`) om de hosts gelijk te houden; met een FSLogix ouder dan 2210 HF4 (Teams) / 25.06 (Outlook) werk je FSLogix ook bij.
- Foutcodes worden getoond met de eigen tekst van Windows (bijv. `0x80073D19` = "An error occurred because a user was logged off", onschuldig) plus een opmerking waar het script er een heeft.
- Blijft een pakket falen terwijl deze host precies de gevraagde build provisiont, dan eindigt de run niet meer met "Nothing to repair" en exitcode 0: hij zegt dat het geen versieverschil is, eindigt met 1, en stap 1c toont het bewijs — de nieuwste AppX-deploymentfout met de *specifieke fouttekst* van Windows en de bijbehorende `Get-AppPackageLog -ActivityID`, en de regels over het pakket in het eigen profiellog van FSLogix (`C:\ProgramData\FSLogix\Logs\Profile`).
- **Copilot, september 2026:** Microsoft voegt de Microsoft 365 Copilot-app en de Windows Copilot-app samen tot één *Microsoft Copilot*-app, geïnstalleerd en bijgewerkt door Edge Update (app-id `{C50565E9-CCCF-44B4-BA15-5AC5C6569197}`). Wat hem tegenhoudt is meestal een beleid, en het script benoemt dat met zijn pad: `HKLM\SOFTWARE\Policies\Microsoft\EdgeUpdate` — `Install{id}` = 0 (geen installatie), `Uninstall{id}` = 1/2 (bij elke controle verwijderd, tenzij `Install{id}` = 5 Force Installs, dat het overschrijft); `PauseCopilotAppUnificationRollout` onder `HKLM\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate`; en Windows' eigen `WindowsCopilot`- / `WindowsAI`-beleid, machinebreed en per aangemelde gebruiker. Die komen uit GPO of Intune en worden door het script **niet** gewijzigd — een lokale aanpassing zou bij de volgende vernieuwing worden teruggedraaid. Het Microsoft 365 Apps-beheercentrum kan de automatische installatie ook uitzetten (*Modern Apps settings*); dat is op de machine niet zichtbaar.
- `-ComputerName` vereist WinRM van de plek waar je het draait naar de hosts (domain-joined hosts: Kerberos werkt zoals het is). Voor Entra-joined hosts is dat vaak niet ingericht; draai het script dan via NinjaOne op elke host.
- De FSLogix-fout voor een gebruiker stopt nadat die gebruiker zich één keer heeft afgemeld op een host die het pakket provisiont — FSLogix slaat dan de huidige versie op.
- Herstart na stap 5 de host wanneer het uitkomt, zodat de deployment-engine de store opnieuw inleest.
- Een back-up-`.reg`-bestand kun je dubbelklikken om een vermelding terug te zetten.
- Bewust niet gedaan: `StateRepository-Machine.srd` of `AppxPackages.xml` bewerken. Beide worden niet ondersteund, en het eerste breekt Start en elke app op een multi-session host als het misgaat.
- Specifiek voor Teams doet [`Update-TeamsClient.ps1`](Update-TeamsClient.ps1) `-RepairAppxStore` het storegedeelte als onderdeel van een update; dit script dekt elk pakket en de FSLogix-kant.

---

## Test-OpenVpnDiagnostics.ps1

Verzamelt en beoordeelt diagnostische informatie over problemen met OpenVPN Connect op een Windows-machine. Controleert elke relevante laag, van driver tot netwerk, en meldt eventuele gevonden problemen.

**Uitgevoerde controles**

| Onderdeel | Wat er wordt gecontroleerd |
|---------|----------------|
| Wintun-/TAP-adapters | PnP-apparaatstatus — markeert alles wat niet `OK` is |
| Virtuele netwerkadapters | Zichtbaarheid van adapters — waarschuwt als er geen zijn terwijl de VPN actief zou moeten zijn |
| Netwerkprofielen | Markeert VPN-adapters die op `Public` staan (moet `Private` zijn) |
| Geïnstalleerde VPN-software | Toont alle VPN-gerelateerde apps — waarschuwt voor mogelijke conflicten met OpenVPN Connect |
| Hyper-V / WSL / virtualisatie | Toont ingeschakelde features — waarschuwt als Hyper-V actief is (kan conflicteren met Wintun) |
| OpenVPN-service | Servicestatus en opstarttype — markeert als hij niet draait |
| Actieve routes | Routes via de VPN-adapter — waarschuwt als de adapter bestaat maar er geen routes zijn |
| DNS-configuratie | DNS-servers per actieve adapter |
| Eventlog | Laatste 20 OpenVPN-vermeldingen uit het Application-log |

De resultaten worden op het scherm getoond, met aan het eind een samenvatting van alle problemen.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-ExportTxt` | Nee | Het volledige rapport opslaan in een txt-bestand |
| `-OutputPath` | Nee | Eigen pad voor het rapport (impliceert `-ExportTxt`). Standaard: `C:\Temp\OpenVpnDiagnostics_<timestamp>.txt` |

**Voorbeelden**

```powershell
# Diagnose draaien, alleen uitvoer op het scherm
.\Test-OpenVpnDiagnostics.ps1

# Draaien en het rapport opslaan in C:\Temp\
.\Test-OpenVpnDiagnostics.ps1 -ExportTxt

# Opslaan op een eigen pad
.\Test-OpenVpnDiagnostics.ps1 -OutputPath "C:\Support\vpn-report.txt"
```

---

## Update-TeamsClient.ps1

> Volledige referentie: [Update-TeamsClient.md](Update-TeamsClient.md) — beslisstroom, versiecontrole, ontwerpkeuzes en probleemoplossing.
>
> Servicedeskversie voor IT Glue (Nederlands, per supportniveau): [Update-TeamsClient-ITGlue.md](Update-TeamsClient-ITGlue.md).

Houdt de nieuwe Teams-client en de Outlook-vergaderinvoegtoepassing actueel op een endpoint of AVD-sessiehost. Het vraagt de Teams-configservice welke build Microsoft voor deze architectuur publiceert en **doet alleen iets als die build nieuwer is dan wat er geïnstalleerd is** — een apparaat dat bij is, blijft volledig ongemoeid. Als er een update nodig is, downloadt het `teamsbootstrapper.exe` en controleert de handtekening, verwijdert de vergaderinvoegtoepassing, verwijdert en deprovisiont het AppX-pakket `MSTeams`, provisiont de nieuwe build voor alle gebruikers en installeert de MSI van de invoegtoepassing die erin meekomt opnieuw.

"Invoegtoepassing aanwezig" betekent dat de bestanden aanwezig zijn, niet dat een registersleutel ernaar verwijst: een machinebrede registratie die wijst naar een loader-DLL die weg is, telt als ontbrekend, en `-CheckOnly` zegt dat ook (`the machine-wide add-in registration points at files that are gone`). Alleen op de sleutel vertrouwen is precies hoe een apparaat waarvan de invoegtoepassing is verwijderd, te horen krijgt dat er niets te doen is.

Elke stap die iets wijzigt, loopt via `ShouldProcess`, dus `-WhatIf` doorloopt de volledige flow zonder de machine aan te raken. Draait handmatig (het verhoogt zichzelf via UAC en vraagt één keer om bevestiging) en onbeheerd vanuit een RMM zoals NinjaOne.

**Stappen**

| # | Stap | Respecteert `-WhatIf` |
|---|------|-------------------|
| 1 | Preflight — inventaris: AppX per gebruiker + geprovisiond, classic Teams (machinebreed + per profiel), invoegtoepassing, Outlook-registratie, de AVD-media-optimalisatie en wat Teams daarover logde, draaiende Teams/Outlook | alleen-lezen |
| 2 | Versiecontrole — gepubliceerde build vs geïnstalleerde build | alleen-lezen |
| 3 | Alleen AVD (`-AvdOptimizations`): vlag `IsWVDEnvironment` + WebRTC-redirector. Of (`-RemoveWebRtcRedirector`): die redirector verwijderen | ja |
| 4 | Alleen classic Teams (`-RemoveClassicTeams`): machinebrede installer + installaties per profiel verwijderen | ja |
| 5 | Werkmap aanmaken, bootstrapper downloaden, Microsoft-handtekening controleren | ja |
| 6 | `MSTeams`-AppX voor alle gebruikers verwijderen en deprovisionen; een pakket dat de AppX-stack weigert te verwijderen, wordt gemeld, niet fataal. De invoegtoepassing blijft hier ongemoeid | ja |
| 7 | Nieuwe Teams provisionen (`teamsbootstrapper.exe -p`) | ja |
| 8 | De volledige vervanging van de invoegtoepassing, zodra de MSI binnen is: de geregistreerde verwijderen (`1612` opnieuw geprobeerd vanuit de gecachte kopie), controleren dat er niets is overgebleven, elke andere kopie opruimen, installeren (`ALLUSERS=1`) | ja |
| 9 | Controleren: registratie van de invoegtoepassing (machinebreed + per aangemelde gebruiker in Outlook), verwijdering van classic, geprovisiond pakket en AVD-componenten | gemeld als overgeslagen onder `-WhatIf` |

Alleen wat ontbreekt, wordt gedaan: een actuele client met een ontbrekende invoegtoepassing installeert alleen de invoegtoepassing, en op een sessiehost met `-AvdOptimizations` installeert een ontbrekende WebRTC-redirector alleen die.

**De media-optimalisatie, beide generaties**

`-AvdOptimizations` installeert de WebRTC-redirector, die Microsoft op **1 oktober 2026** uitfaseert (einde beschikbaarheid 1 april 2027). De opvolger SlimCore wordt nooit op de sessiehost geïnstalleerd — de plugin in Windows App zet hem klaar op het *endpoint* van waaruit de gebruiker verbindt — dus de vraag aan de hostkant is niet "staat SlimCore hier" maar "krijgen mijn gebruikers hem". Het antwoord staat in het Application-eventlog: Teams schrijft bij elke verbinding een `Microsoft Teams VDI`-event, en preflight leest daarvan de laatste zeven dagen op elke sessiehost en vertaalt de codes (`24002`/`24010` = op SlimCore, `16002` = endpoint heeft geen plugin, `16389` = beleid blokkeerde de MSIX). `-RemoveWebRtcRedirector` haalt de oude generatie weg zodra niets meer `16002` meldt; het weigert samen met `-AvdOptimizations` te draaien, en laat `IsWVDEnvironment` staan omdat SlimCore die vlag ook nodig heeft.

> Dat log opvragen vereist `Get-WinEvent -FilterXPath`, niet `-FilterHashtable`: de hashtablevorm gooit meteen een fout als de provider nog nooit een event heeft geschreven, wat het normale geval is op een machine die geen sessiehost is.

Op een endpoint controleert preflight ook de drie beleidsinstellingen die het klaarzetten tegenhouden: `BlockNonAdminUserInstall` (fout `16389`), `AllowAllTrustedApps` (`15615`) en AppLocker (`10083`). AppLocker wordt gelezen in plaats van alleen gedetecteerd — alleen de verzameling voor packaged apps (`Appx`) kan een MSIX blokkeren, een verzameling met regels waarvan de handhaving *niet geconfigureerd* is, wordt evengoed gehandhaafd, en niets wordt gehandhaafd zolang de service Application Identity gestopt is. Het rapport noemt het registerpad, de modus per verzameling, de status van de service en de regelnamen, en zegt het als een regel de pakketten al toestaat. Een beleid dat op een *sessiehost* wordt gevonden, wordt ter informatie vermeld in plaats van als waarschuwing: het blokkeert daar niets, maar het is meestal dezelfde GPO die ook de endpoints bereikt.

**Waarom de volgorde ertoe doet:** er wordt niets aangeraakt tot een nieuwere build is bevestigd, en de installer wordt opgehaald en gecontroleerd *vóór* de eerste verwijdering — zo kan een mislukte download of een geblokkeerde URL het apparaat nooit zonder Teams-client achterlaten.

**Versiecontrole**

`https://config.teams.microsoft.com/config/v1/MicrosoftTeams/...` is de feed die de Teams-client zelf gebruikt om te bepalen dat hij verouderd is. Hij geeft de actuele build per architectuur terug (`BuildSettings.WebView2PreAuth.<arch>.latestVersion`). Een geïnstalleerde build die gelijk is aan of nieuwer is dan die, betekent dat er niets te doen is. Als de service niet bereikbaar is, stopt de run in plaats van blind opnieuw te installeren — `-Force` overschrijft dat. `-Ring` kiest een andere updatering (standaard `general`).

> Het verwijderen en controleren van de invoegtoepassing leest zowel de 64-bits als de `WOW6432Node`-uninstallhive, omdat niet vastligt in welke van de twee de vermelding terechtkomt (gemeten: 64-bits voor invoegtoepassing 1.26.21803 op Windows 11). De MSI-versie van de invoegtoepassing komt uit de MSI-propertytabel (`WindowsInstaller.Installer` COM), niet uit `Get-AppLockerFileInformation`, dat op sommige edities ontbreekt en onder PowerShell 7 stukgaat.

**Veiligheid**

- `msiexec` en de bootstrapper draaien met een time-out (`-TimeoutSeconds`, standaard 900) en worden afgebroken als ze vastlopen, zodat een RMM-job de agent niet kan blokkeren.
- MSI-exitcode `1618` (er loopt al een andere installatie) wordt twee keer opnieuw geprobeerd; `3010` telt als geslaagd en markeert in de samenvatting een openstaande herstart.
- Onverwachte fouten breken de run af in plaats van halverwege door te gaan.
- Een run die echt iets wijzigt, schrijft een transcript naar `C:\Temp\Update-TeamsClient_<timestamp>.log`; een controle die niets te doen vindt, laat geen logrommel achter.

**Exitcodes**

| Code | Betekenis |
|------|---------|
| `0` | Geslaagd, of al actueel |
| `1` | Fout |
| `2` | Alleen bij `-CheckOnly`: er is een nieuwere build beschikbaar |

**Parameters**

| Parameter | Omschrijving |
|-----------|-------------|
| `-WhatIf` | Tonen wat een update zou doen zonder hem uit te voeren |
| `-Quiet` | Niets tonen tenzij er nieuws is: een nieuwere build, een actie of een fout |
| `-CheckOnly` | Alleen melden of er een nieuwere build is (exitcode 2), niets wijzigen |
| `-AvdOptimizations` | AVD/VDI-sessiehosts: de vlag `IsWVDEnvironment` en de WebRTC-redirector afdwingen |
| `-RemoveClassicTeams` | Ook de classic Teams-client verwijderen: machinebrede installer plus installaties per profiel |
| `-RemoveWebRtcRedirector` | De oude WebRTC-media-optimalisatie verwijderen, uitgefaseerd op 1 oktober 2026. Niet te combineren met `-AvdOptimizations`; laat `IsWVDEnvironment` staan, omdat SlimCore die ook nodig heeft |
| `-ClearOrphanedAddInRegistration` | Laatste redmiddel: Windows Installer een vergaderinvoegtoepassing laten vergeten die het niet meer kan verwijderen (`1612` met de gecachte MSI weg), wat een herinstallatie steeds met `1638` laat weigeren |
| `-RepairAppxStore` | Laatste redmiddel voor de AppX-kant: een pakket waarvan de bestanden er nog zijn opnieuw registreren, en daarna de `AppxAllUserStore`-vermeldingen opruimen die Windows niet meer kan herleiden — registraties voor SID's zonder profiel, een machinebrede vermelding waarvan het manifest weg is, en de `Deprovisioned`-markering. Beperkt tot MSTeams; preflight benoemt ze, of de switch nu is opgegeven of niet |
| `-UseWinget` | De Teams-MSIX met winget ophalen en precies dat bestand provisionen (`teamsbootstrapper.exe -p -o`) in plaats van de bootstrapper er tijdens de run een te laten downloaden |
| `-RepairOutlookAddIn` | Een Outlook-registratie per gebruiker opruimen die wijst naar een invoegtoepassings-DLL die niet meer bestaat, zodat de machinebrede het weer overneemt |
| `-Confirm:$false` | Nooit om bevestiging vragen (gebruik dit voor onbeheerde runs) |
| `-Ring` | Updatering die bij de configservice wordt opgevraagd (standaard: `general`) |
| `-WorkingDir` | Downloadmap voor de bootstrapper (standaard: `C:\IT\AVD\Teams`) |
| `-LogPath` | Map voor het transcript (standaard: `C:\Temp`) |
| `-BootstrapperUrl` | De download-URL van `teamsbootstrapper.exe` overschrijven (alleen https) |
| `-WebRtcUrl` | De MSI-URL van de WebRTC-redirector overschrijven (alleen https) |
| `-SkipMeetingAddIn` | De vergaderinvoegtoepassing met rust laten, en een ontbrekende invoegtoepassing niet als werk zien. Verdedigbaar op gewone endpoints, waar de Teams-client de invoegtoepassing zelf per gebruiker actueel houdt |
| `-SkipSignatureCheck` | Een installer accepteren die niet door Microsoft is ondertekend (interne mirror) |
| `-TimeoutSeconds` | Time-out per proces voor msiexec/bootstrapper (standaard: `900`) |
| `-Force` | Opnieuw installeren ook als Teams actueel is, en doorgaan zonder Teams- of versie-informatie |

**Voorbeelden**

```powershell
# Proefdraai — op een nieuwere build controleren en tonen wat een update zou doen
.\Update-TeamsClient.ps1 -WhatIf

# Alleen bijwerken als Microsoft een nieuwere build heeft gepubliceerd
.\Update-TeamsClient.ps1

# Geplande RMM-run: stil tenzij er een nieuwere build of een probleem is
.\Update-TeamsClient.ps1 -Quiet -Confirm:$false

# Alleen detectie: exitcode 2 als er een update beschikbaar is
.\Update-TeamsClient.ps1 -CheckOnly -Quiet

# Herstel: volledige herinstallatie ongeacht de versiecontrole
.\Update-TeamsClient.ps1 -Force -Confirm:$false

# Gezondheidsrapport van een sessiehost: alles rond Teams, wijzigt niets
.\Update-TeamsClient.ps1 -CheckOnly

# Migratie klaar: de uitgefaseerde WebRTC-optimalisatie weghalen
.\Update-TeamsClient.ps1 -RemoveWebRtcRedirector -WhatIf
```

**Uitvoeren vanuit NinjaOne**

1. Voeg het script toe (Language: PowerShell, Operating System: Windows, Architecture: **All**, Run As: **System**).
2. Bekijk eerst één apparaat: draai het met `-WhatIf -Confirm:$false` in het veld *Parameters* — de joboutput toont de versievergelijking en elke stap die een update zou uitvoeren, en het apparaat blijft onaangeroerd.
3. Plan de echte run met `-Quiet -Confirm:$false`. Op een apparaat dat bij is, toont het niets en eindigt het met `0`, zodat de activiteitenfeed alleen de apparaten laat zien waar het echt iets heeft gedaan.
4. Gebruik voor een detectie-/conditiejob `-CheckOnly -Quiet`: stil en `0` als het actueel is, uitvoer en exitcode `2` als er een nieuwere build is gepubliceerd.
5. Optionele scriptvariabelen (checkboxen `whatIf`, `quiet`, `checkOnly`, `force`, `avdOptimizations`, `removeWebRtcRedirector`, `removeClassicTeams`, `repairOutlookAddIn`, `clearOrphanedAddInRegistration`, `skipMeetingAddIn`, `skipSignatureCheck`; tekstvelden `workingDir`, `logPath`, `ring`, `webRtcUrl`, `bootstrapperUrl`) worden uit de omgeving opgepikt als de bijbehorende parameter niet is meegegeven, zodat een technicus *whatIf* kan aanvinken in plaats van parameters te typen.

Als de agent PowerShell 32-bits start, start het script zichzelf eerst opnieuw als 64-bits via `SysNative` — zonder dat worden de registerleesacties omgeleid naar `WOW6432Node` en wijst `$env:ProgramFiles` naar de x86-map, zodat noch het AppX-pakket noch de MSI van de invoegtoepassing wordt gevonden.

---

## Time sync/

Verhelpt problemen met de Windows-tijdsynchronisatie door `W32tm` te herstarten tegen Nederlandse NTP-poolservers en een geplande taak te registreren die de synchronisatie elke 59 minuten opnieuw uitvoert. Zie [`Time sync/readme.nl.md`](Time%20sync/readme.nl.md) voor alle details.
