[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../readme.nl.md) › [scripts](../readme.nl.md) › **RDS**

# RDS

Scripts voor diagnose, monitoring en het klaarmaken van RDP- / RD Web Access-infrastructuur en AVD-sessiehosts. Voer ze rechtstreeks op de RDS-/RDWeb-server uit voor volledige resultaten — doelen op afstand krijgen alleen controles op connectiviteitsniveau.

---

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`Test-RDSDiagnostics.ps1`](Test-RDSDiagnostics.ps1) ([docs](#test-rdsdiagnosticsps1)) | Eenmalige gezondheidscontrole — services, configuratie, certificaten, gebruikersaccount, eventlogs |
| [`Watch-RDSLive.ps1`](Watch-RDSLive.ps1) ([docs](#watch-rdsliveps1)) | Realtime monitor van sessie- en licentie-events |
| [`Get-FSlogix-errors.ps1`](Get-FSlogix-errors.ps1) ([docs](#get-fslogix-errorsps1)) | Diagnose van FSLogix- / Azure Files-profielen op een AVD-sessiehost |
| [`Invoke-FSLogixShrink.ps1`](Invoke-FSLogixShrink.ps1) ([docs](#invoke-fslogixshrinkps1)) | FSLogix-profielschijven op een share verkleinen (Invoke-FslShrinkDisk), of controleren of FSLogix ze zelf comprimeert bij afmelden |
| [`Update-SessionHostImage.ps1`](Update-SessionHostImage.ps1) ([docs](#update-sessionhostimageps1)) | Een Windows 11 multi-session-image of AVD-sessiehost controleren en klaarmaken, zodat de nieuwe Teams, de nieuwe Outlook en Copilot blijven werken met FSLogix — FSLogix zelf blijft ongemoeid |
| [`Watch-M365Apps.ps1`](Watch-M365Apps.ps1) ([docs](#watch-m365appsps1)) | Watchdog (geplande taak) — test de nieuwe Teams, de nieuwe Outlook en Copilot met ons eigen account (`itceadmin`) op een sessiehost, herstelt wat stuk is voordat een klant er last van heeft, en meldt het aan n8n |
| [`Get-M365AppsLog.ps1`](Get-M365AppsLog.ps1) ([docs](#get-m365appslogps1)) | Verzamelt, alleen lezend, wat er met de nieuwe Teams, de nieuwe Outlook en Copilot op een sessiehost gebeurd is — per gebruiker wat er geregistreerd is en draait, de events, en wat de watchdog deed en zou concluderen, met zijn blinde vlekken — in één zip |

---

### Test-RDSDiagnostics.ps1

Achterhaalt waarom gebruikers zich niet kunnen aanmelden op een RDP- of RD Web Access-server.

**Uitgevoerde controles**

| Onderdeel | Details |
|------|---------|
| RDP-server | Status van `TermService`/`SessionEnv`/`UmRdpService`, RDP in-/uitgeschakeld, NLA, sessielimieten, RD Licensing-modus, groep Remote Desktop Users, firewallregels, actieve sessies (`quser`) |
| RDWeb-server | Bereikbaarheid van poort 443, geldigheid/verloopdatum van het HTTPS-certificaat, IIS + RDWeb-app-pool (alleen lokaal), RD Gateway-service (alleen lokaal) |
| Gebruikersaccount (optioneel, `-Username`) | Ingeschakeld/vergrendeld/verlopen, groepslidmaatschap, beperkingen op aanmeldwerkstations, laatste aanmelding, leeftijd van het wachtwoord |
| Eventlogs (optioneel, `-IncludeEventLogs`) | Security 4625 (mislukte RDP-aanmelding), 4740 (vergrendeling), `TerminalServices-LocalSessionManager` 20/40 (sessiefout/reden van verbreken) |

**Parameters**

| Parameter | Omschrijving |
|-----------|-------------|
| `-RdpServer` | Hostnaam/IP van de RDP-/Terminal Server. Laat weg om lokale controles uit te voeren |
| `-RdWebServer` | Hostnaam/IP van de RD Web Access-server |
| `-Username` | Een specifiek account controleren op zaken die het aanmelden blokkeren (vereist de module ActiveDirectory of de ADSI-terugval) |
| `-LogPath` | Map voor het logbestand (standaard: `C:\Temp\`) |
| `-IncludeEventLogs` | Analyse van de eventlogs over de laatste `-Hours` uur meenemen |
| `-Hours` | Aantal uren eventloggeschiedenis om te analyseren (standaard: `24`) |

**Voorbeelden**

```powershell
# Volledige controle — RDP + RDWeb + gebruikersaccount
.\Test-RDSDiagnostics.ps1 -RdpServer rdp01.company.local -RdWebServer rdweb.company.local -Username jdoe -IncludeEventLogs

# Lokaal op de RDS-host uitvoeren, eventlogs controleren
.\Test-RDSDiagnostics.ps1 -IncludeEventLogs -Hours 48
```

De resultaten worden naar de console geschreven en naar een logbestand met tijdstempel in `C:\Temp\`.

---

### Watch-RDSLive.ps1

Bevraagt de Windows-eventlogs elke N seconden en streamt nieuwe events naar de console + een logbestand. Voer het rechtstreeks uit op elke RDS-/RDWeb-server.

**Gemonitorde events**

| Bron | Events |
|--------|--------|
| `TerminalServices-LocalSessionManager` | Aanmelding (21), opnieuw verbinden (22/25), afmelding (23), verbinding verbroken (24), aanmelding mislukt (20), reden van verbreken (40) — leesbare redencodes |
| Security | Mislukte RDP-aanmelding (4625, type 10), accountvergrendeling (4740) |
| `TerminalServices-Licensing` | Events voor licentie toegekend/geweigerd/waarschuwing |
| System | Provider `TermServLicensing` (respijtperiode, fouten van de licentieserver) |

Toont per poll een heartbeatregel met het aantal actieve sessies.

**Parameters**

| Parameter | Omschrijving |
|-----------|-------------|
| `-IntervalSeconds` | Pollinterval in seconden (standaard: `20`) |
| `-LogPath` | Map voor het logbestand (standaard: `C:\Temp\`) |
| `-NoLogFile` | Alleen uitvoer naar de console, geen logbestand |

**Voorbeelden**

```powershell
# Uitvoeren op de RDS-server
.\Watch-RDSLive.ps1

# Sneller pollen, geen logbestand
.\Watch-RDSLive.ps1 -IntervalSeconds 10 -NoLogFile
```

> Druk op `Ctrl+C` om te stoppen. Voer uit als Administrator voor toegang tot het Security-log.

---

### Get-FSlogix-errors.ps1

Verzamelt in één run alles wat nodig is om te achterhalen waarom een FSLogix-profiel
niet mount op een AVD-sessiehost — mountfouten, een vergrendelde VHDX, problemen met
SMB/Azure Files of schijffouten. Alleen-lezen: het verzamelt en rapporteert, het repareert niets.

**Wat het verzamelt**

| Onderdeel | Details |
|------|---------|
| Systeem | Hostnaam, OS-build, uptime |
| FSLogix | Geïnstalleerde versie, de volledige configuratie van `Profiles`/`Containers` en de status van de service |
| Containers | Gekoppelde VHD(X)-bestanden, het sessieregister van FSLogix en de profielpaden uit `ProfileList` |
| Opslag | SMB-verbindingen met Azure Files, en of de VHD-share überhaupt bereikbaar is |
| Events | FSLogix-events over de laatste `-Days` dagen, vergeleken met de bekende kritieke foutpatronen, plus schijf-/NTFS-fouten en events van de User Profile Service |
| Restanten | Lokale profielen onder `C:\Users` en de logbestanden van FSLogix |

**Parameters**

| Parameter | Omschrijving |
|-----------|-------------|
| `-User` | Ook een sectie maken die gefilterd is op één gebruiker — het account waarvan het profiel faalt |
| `-Days` | Aantal dagen eventgeschiedenis om te analyseren (standaard: `7`) |
| `-OutputPath` | Map voor het transcript (standaard: `%SystemDrive%\Temp\FSLogixDiag`) |

**Voorbeelden**

```powershell
# Alles van de afgelopen week
.\Get-FSlogix-errors.ps1

# Eén gebruiker, twee weken terug, rapport op een andere plek
.\Get-FSlogix-errors.ps1 -User jdoe -Days 14 -OutputPath C:\Temp
```

> Voer het uit in een verhoogde sessie **op de sessiehost zelf** — de gegevens over
> containers, SMB en events bestaan alleen daar. De hele run wordt geschreven naar
> `FSLogixDiag_<host>_<timestamp>.log` in de uitvoermap; dat is het bestand dat je
> aan een ticket toevoegt.

---

### Invoke-FSLogixShrink.ps1

Geeft de ruimte terug die FSLogix-profiel- en ODFC-containers vasthouden nadat er data in
is verwijderd: een dynamische VHDX groeit, maar krimpt nooit vanzelf. Het is een schil om
[Invoke-FslShrinkDisk](https://github.com/FSLogix/Invoke-FslShrinkDisk), het eigen
verkleinscript van het FSLogix-team — nog steeds het beste gereedschap hiervoor; de forks en
alternatieven op GitHub doen hetzelfde met minder erachter.

**Wat het doet**

| Stap | Details |
|------|---------|
| Downloaden | Haalt Invoke-FslShrinkDisk op een **vastgezette commit** op naar `C:\Scripts\Invoke-FslShrinkDisk`, deblokkeert het en controleert de SHA-256 van het script. Een gewijzigde upstream-versie draait nooit ongezien; een aangepaste lokale kopie wordt geweigerd |
| Rapport | Elke `.vhd`/`.vhdx` op de share, grootste eerst, met map, grootte en laatste schrijfdatum, en het totaal |
| Verkleinen | Draait Invoke-FslShrinkDisk recursief en vat daarna het CSV-log samen: verkleinde schijven, teruggewonnen GB, en de schijven die niet verwerkt konden worden |
| `-CheckHost` | Op een sessiehost: kan de **ingebouwde compressie van FSLogix bij afmelden** draaien? FSLogix-versie (2210 / 2.9.8361 of later), `VHDCompactDisk`, de service Optimize Drives (`defragsvc` niet Disabled) en dynamische schijven |

**Parameters**

| Parameter | Omschrijving |
|-----------|-------------|
| `-Path` | De share met de containers, bijv. `\\<storageaccount>.file.core.windows.net\<share>\Profiles`. Wordt recursief doorzocht |
| `-ReportOnly` | Alleen de schijven en hun grootte tonen; niets verkleinen |
| `-IgnoreLessThanGB` | Schijven kleiner dan dit overslaan (standaard: `5`) |
| `-RatioFreeSpace` | Alleen een schijf verkleinen met minstens dit deel vrij erbinnen (standaard: `0.1` = 10%) |
| `-ThrottleLimit` | Aantal schijven tegelijk (standaard: `4`; hooguit twee keer het aantal CPU-cores) |
| `-LogFilePath` | CSV-log (standaard: `C:\Temp\FslShrink_<timestamp>.csv`); de map wordt aangemaakt als die ontbreekt |
| `-ToolPath` | Waar Invoke-FslShrinkDisk staat (standaard: `C:\Scripts\Invoke-FslShrinkDisk`) |
| `-Force` | Invoke-FslShrinkDisk opnieuw downloaden |
| `-CheckHost` | In plaats daarvan de eigen compressie van FSLogix op deze host controleren; geen `-Path` nodig |

**Voorbeelden**

```powershell
# Eerst kijken: elke container op de share, grootste eerst
.\Invoke-FSLogixShrink.ps1 -Path \\sa.file.core.windows.net\profiles\Profiles -ReportOnly

# Alles van 5 GB of meer verkleinen met minstens 10% vrij erbinnen
.\Invoke-FSLogixShrink.ps1 -Path \\sa.file.core.windows.net\profiles\Profiles

# Comprimeert FSLogix de schijven zelf op deze host?
.\Invoke-FSLogixShrink.ps1 -CheckHost
```

**Opmerkingen**

- FSLogix 2210 en later comprimeert een container zelf bij elke afmelding, als de schijf
  groter is dan 1 GB en er minstens 20% te winnen valt ([Microsoft Learn](https://learn.microsoft.com/en-us/fslogix/concepts-vhd-disk-compaction)).
  Draai eerst `-CheckHost`: slaagt die, dan haalt een handmatige verkleining alleen de
  schijven in van gebruikers die zelden afmelden, of die onder de drempel van 20% blijven.
- Een schijf die gekoppeld is — de gebruiker is aangemeld — kan niet verkleind worden; die
  staat in de samenvatting als niet verwerkt en de run eindigt met exitcode 1. Draai het
  buiten kantooruren of met de hosts leeggemaakt.
- Verhoogd uitvoeren (elke schijf wordt gekoppeld), met toegang tot de share: op Azure Files
  via Kerberos of de sleutel van het opslagaccount. Hyper-V is niet nodig.
- Naar een nieuwere Invoke-FslShrinkDisk: zet de nieuwe commit en de SHA-256 van zijn
  `Invoke-FslShrinkDisk.ps1` in `$ToolCommit` / `$ToolHash` bovenin het script, nadat je de
  diff hebt gelezen.

---

### Update-SessionHostImage.ps1

Teams, de nieuwe Outlook en Copilot zijn MSIX-apps, en op een gedeelde AVD-host met
FSLogix gaan ze steeds op dezelfde manier stuk: de app van een gebruiker werkt zichzelf bij
op host A, FSLogix bewaart die exacte versie bij afmelden in het profiel, en bij de volgende
aanmelding op host B — die die versie niet heeft — mislukt de registratie met `0x80070490`.
FSLogix 2210 HF4 (Teams) en 25.06 (Outlook) registreren op pakketfamilie, maar dit script
laat FSLogix bewust ongemoeid. Het houdt elke host op de nieuwste build van de apps en van
alles wat ze nodig hebben, op elke host gelijk.

**Wat het controleert**

| Stap | Details |
|------|---------|
| Windows | Editie (Enterprise multi-session), build, openstaande herstart |
| FSLogix | Build en `InstallAppxPackages` — alleen lezen |
| Updates | De nieuwe Outlook werkt zichzelf **wekelijks bij via het Office CDN**, niet via de Store, en heeft geen schakelaar om dat te stoppen: onder FSLogix 25.06 is de remedie dit script wekelijks op elke host draaien. Teams: onder 2210 HF4 gaat de zelfupdate uit (`disableAutoUpdate = 1`) en werkt elke run Teams centraal bij; op een nieuwere FSLogix mag het zichzelf bijwerken. De Store-instelling wordt getoond, niet gewijzigd. Edge Update-beleid dat WebView2 of Edge blokkeert wordt gemeld |
| WebView2 | De Evergreen-runtime waar alle drie de apps mee tekenen, vergeleken met de huidige Edge Stable-build (`edgeupdates.microsoft.com`) |
| Apps | Teams, de nieuwe Outlook, de Microsoft 365 Copilot-app en de unified Copilot-app: de klaargezette build, en gebruikers met een nieuwere build dan de image klaarzet |
| Frameworks | Elke `PackageDependency` in de manifests van die apps (VCLibs, UI.Xaml, WindowsAppRuntime, …) moet op de machine staan in de `MinVersion` die het manifest vraagt — op een image uit 2024 meestal te oud |
| Teams op AVD | `IsWVDEnvironment`, een Teams-build die nieuw genoeg is voor SlimCore (`24193.1805.3040.8975`), de Teams Meeting-invoegtoepassing, en de WebRTC-redirector: niet meer ondersteund sinds **1 oktober 2026**, werkt niet meer vanaf **1 april 2027**, alleen nog bewaard als terugvaloptie voor eindpunten die nog geen SlimCore kunnen |
| Office | Shared Computer Activation (verplicht op multi-session) en het updatekanaal |
| Aanmelden | `Microsoft.AAD.BrokerPlugin` aanwezig, en geen uitsluiting in de FSLogix-`redirections.xml` van `AppData\Local\Packages` of de app-mappen |
| Capture | Met `-ForCapture`: pakketten die voor een gebruiker zijn geïnstalleerd maar niet klaargezet (Sysprep stopt erop) en een openstaande herstart laten de controle falen |

**Wat het bijwerkt** (elke run zonder `-CheckOnly`, met of zonder bevinding, in deze
volgorde): de zelfupdate-instelling van Teams waar FSLogix dat nodig heeft, Shared Computer
Activation, WebView2 (Evergreen Standalone-installer, handtekening gecontroleerd), daarna
de apps naar hun nieuwste build via de bestaande scripts —
[`Repair-AppxPackageStore.ps1`](../Device/readme.nl.md#repair-appxpackagestoreps1)
`-Name teams,outlook -Latest -Provision -RemoveOld` en `-Name copilot -Provision`, en
[`Update-TeamsClient.ps1`](../Device/readme.nl.md#update-teamsclientps1) `-AvdOptimizations`
(nieuwste Teams, vergaderinvoegtoepassing, IsWVDEnvironment, WebRTC-redirector). Ze wijzigen
alleen wat achterloopt. Daarna wordt alles opnieuw uitgelezen.

De nieuwste Outlook: Microsoft publiceert er geen versiefeed voor, dus `-Latest` neemt de
nieuwste build waarnaar gebruikers op de host al zijn bijgewerkt (anders de installer van
Microsoft). In een pool zet elke host dan klaar wat de laatste gebruiker kreeg — precies de
build waar FSLogix de andere hosts om zal vragen.

**Parameters**

| Parameter | Omschrijving |
|-----------|-------------|
| `-CheckOnly` | Alleen rapporteren, niets wijzigen. Exitcode `2` betekent dat er werk is |
| `-ForCapture` | De machine is de image-VM die zo gesysprept wordt: ook falen op een openstaande herstart en op per gebruiker geïnstalleerde pakketten die niet klaargezet zijn |
| `-SkipApps` | Repair-AppxPackageStore / Update-TeamsClient niet aanroepen; alleen beleid, Shared Computer Activation en WebView2 |
| `-ComputerName` | Sessiehosts om via PowerShell remoting op te draaien; eindigt met één tabel over de pool en noemt elke kolom die tussen hosts verschilt |
| `-Credential` | Referenties voor `-ComputerName` |
| `-WorkingDir` | Map voor downloads — WebView2 en, als ze niet naast dit script staan, de twee hulpscripts (standaard: `C:\IT\SessionHostImage`) |
| `-LogPath` | Map voor het transcript van een run die iets wijzigt (standaard: `C:\Temp`) |

**Voorbeelden**

```powershell
# Wat heeft deze host of image nodig? Wijzigt niets.
.\Update-SessionHostImage.ps1 -CheckOnly

# Alle drie de sessiehosts naast elkaar, alleen lezen
.\Update-SessionHostImage.ps1 -ComputerName avd-0,avd-1,avd-2 -CheckOnly

# De image-VM klaarmaken en daarna controleren of hij klaar is voor capture
.\Update-SessionHostImage.ps1 -Confirm:$false
.\Update-SessionHostImage.ps1 -CheckOnly -ForCapture

# Alle drie de hosts gelijktrekken (eerst in drain mode zetten)
.\Update-SessionHostImage.ps1 -ComputerName avd-0,avd-1,avd-2 -Confirm:$false
```

**Opmerkingen**

- Het script herstart de machine nooit, en de scripts en installers die het aanroept ook
  niet (elke msiexec draait met `/norestart`). Een openstaande herstart wordt gemeld en aan
  jou overgelaten.
- Verhoogd of als System uitvoeren. Het script start zichzelf opnieuw in 64-bits Windows
  PowerShell, omdat de AppX-cmdlets dat nodig hebben.
- Draai het **wekelijks** op alle hosts tegelijk (een geplande taak als System werkt), en na
  elke Windows-update: Outlook verandert wekelijks, en elke host moet dat bijhouden.
- SlimCore heeft ook de tenantkant nodig: het Teams VDI-beleid `VDI2Optimization` aan, en
  Windows App 2.0.352.0 of nieuwer op de eindpunten. Geen van beide is vanaf de host te
  controleren; `Update-TeamsClient.ps1 -CheckOnly` leest de Teams VDI-events van de
  sessiehost, en die laten wel zien of gebruikers echt op SlimCore zitten.
- De twee scripts die het aanroept komen uit de repo als het daaruit draait, of uit de map
  waar `-ComputerName` ze neerzet. Draait het los — alleen dit bestand op de image-VM — dan
  haalt het ze van GitHub op een **vastgepinde commit** van `main` en voert het ze alleen
  uit als de SHA-256 klopt, zoals [`Invoke-FSLogixShrink.ps1`](#invoke-fslogixshrinkps1) dat met Invoke-FslShrinkDisk
  doet. Bijwerken: zet de nieuwe commit en de hashes van beide bestanden in `$HelperCommit` /
  `$HelperHashes` bovenin het script, nadat je de diff hebt gelezen.
- `-ComputerName` kopieert dit script en de twee die het aanroept naar
  `C:\IT\SessionHostImage` op elke host.

---

### Watch-M365Apps.ps1

Een watchdog voor de nieuwe Teams, de nieuwe Outlook en Copilot op een sessiehost. Hij
draait als geplande taak onder System en gebruikt **ons eigen account** — standaard
`itceadmin` — als kanarie: start een app daar niet, dan start hij voor een klant ook niet. Hij controleert ook elke andere aangemelde gebruiker en elke app die een gebruiker niet
kon openen, herstelt dat op de host en in de eigen sessie van die gebruiker voordat die
belt, en meldt het aan een n8n-webhook — met de gebruiker bij elke bevinding en elk
herstel.

> Servicedeskversie voor IT Glue (Nederlands, per supportniveau — wat te doen als een gebruiker belt, hoe je de Teams-kaarten leest): [Watch-M365Apps-ITGlue.md](Watch-M365Apps-ITGlue.md), of de opgemaakte [Watch-M365Apps-ITGlue.html](Watch-M365Apps-ITGlue.html) om te plakken.

**Elke run**

| Stap | Wat er gebeurt |
|------|----------------|
| 0. Crashes | Elke crash (Application Error `1000`) en elke hang die eindigde in afsluiten (Application Hang `1002`) van Teams, de nieuwe Outlook of Copilot sinds de vorige run, uit het Application-logboek — voor **elke gebruiker op de host**, klanten inbegrepen — gegroepeerd per app, module en foutcode. Gemeld, niet hersteld: de app van een klant wordt nooit voor hem herstart |
| 1. Host | Teams en de nieuwe Outlook zijn klaargezet (provisioned) voor alle gebruikers; Copilot is alleen de **nieuwe, unified Microsoft Copilot-app** (`copilotapp.exe`, machinebreed geïnstalleerd door Edge Update) — de oude Microsoft 365 Copilot-app (`MicrosoftOfficeHub`) telt niet meer. Zijn `copilotapp.exe` moet op schijf staan en een Start-menu-snelkoppeling voor alle gebruikers moet ernaar wijzen; een ontbrekende snelkoppeling wordt gemaakt, een ontbrekende app geïnstalleerd door het hostherstel (`Repair-AppxPackageStore.ps1 -Name copilot -Provision`, via Edge Update) |
| 1a. Bijwerken | Elke `-UpdateHours` (standaard 6): de **nieuwste** Teams en nieuwe Outlook worden op de host klaargezet — [`Repair-AppxPackageStore.ps1`](../Device/readme.nl.md#repair-appxpackagestoreps1) `-Latest -Provision` (Teams via Microsofts configuratieservice, Outlook de nieuwste build die op deze host of in een profiel gezien is; alleen nieuwer, handtekening gecontroleerd, geen `-RemoveOld`) — en Edge Update wordt gevraagd nu te controleren op de unified Copilot-app. Elke run: heeft **ons eigen account** een oudere build dan de host klaarzet, dan wordt de app in onze sessie gesloten en opnieuw geregistreerd vanaf de klaargezette build, zodat stap 2 de nieuwe build start — een build die niet draait vinden wij, niet een klant bij de volgende aanmelding. **Klanten worden nooit door de watchdog bijgewerkt**: Windows geeft hun de nieuwe build bij de volgende aanmelding. Een nieuwe build op de host — door deze stap of vanzelf — wordt één keer gemeld (`updated`, met oude en nieuwe versie) |
| 2. Accounts | Voor elk bewaakt account dat op deze host is aangemeld: het pakket is voor die gebruiker geregistreerd, de bestanden zijn er en de status is `Ok` — voor Copilot zijn identiteitspakket, als de host er een heeft (de bestanden zijn machinebreed). Daarna moet de app in die sessie draaien — zo niet, dan wordt hij daar gestart (`shell:AppsFolder\<AUMID>`, via een eenmalige taak in de eigen sessie van die gebruiker) en moet hij 15 seconden later nog draaien |
| 2b. Gebruikers | Elke andere aangemelde gebruiker (klanten), zodra die 10 minuten is aangemeld: dezelfde registratiecontrole, **zonder iets te starten**. Plus elke poging sinds de vorige run, door welke gebruiker ook, om een van de apps te openen die Windows weigerde (TWinUI `5961`), en elke mislukte registratie van hun pakketten (AppXDeploymentServer `401`/`404`; "sluit eerst de app" en "al geïnstalleerd" tellen niet mee), met de gebruiker bij wie het gebeurde |
| 2c. Aankondigen | Er is iets nieuws mis: [`Get-M365AppsLog.ps1`](#get-m365appslogps1) verzamelt het bewijs voor de betreffende gebruikers en apps in `Diag\` (één zip, 14 dagen bewaard) terwijl het nog stuk is, en er gaat een melding `repairing` naar n8n met elke gebruiker erin — voordat er iets veranderd wordt. Hetzelfde probleem opnieuw wordt binnen `-RenotifyHours` niet nog eens aangekondigd of verzameld |
| 3. Herstel | Een probleem van **één klant wordt alleen in diens sessie hersteld**; de host blijft ongemoeid. De host wordt voor een app alleen hersteld als het meer is dan die ene gebruiker — de host zelf, **ons eigen account** (`itceadmin` is de test voor de hele host), of dezelfde app bij twee of meer klanten — zonder bij iemand iets te starten: [`Repair-AppxPackageStore.ps1`](../Device/readme.nl.md#repair-appxpackagestoreps1) `-Provision` (installers van Microsoft, handtekening gecontroleerd), hooguit één keer per `-RepairCooldownHours`. Daarna **per gebruiker, in diens eigen sessie**: het pakket wordt opnieuw op familienaam geregistreerd (`Add-AppxPackage -RegisterByFamilyName`) via een eenmalige taak met een console zonder venster, zodat er niets verschijnt. Bij een klant alleen als de app op dat moment niet bij hem draait, hooguit één keer per `-RepairCooldownHours` per gebruiker en app, en nooit een reset. In ons eigen account wordt een app die niet start gereset (`Reset-AppxPackage`). Een gebruiker die de app in het laatste halfuur **zelf** probeerde te openen (TWinUI `5961`) krijgt hem **voor zich geopend** in zijn sessie zodra hij weer geregistreerd is — hij moet starten en blijven draaien, anders blijft het probleem open. Een poging in de eerste 2 minuten na het aanmelden is de autostart van de app en telt niet, en verder wordt er bij een klant nooit iets gestart (`-NoUserLaunch` zet dit uit) |
| 4. Teruglezen | Stap 1, 2 en de registratiecontrole van 2b opnieuw; het probleem van een gebruiker telt als hersteld als het pakket daarna voor hem geregistreerd en `Ok` is (of de app bij hem draait) |
| 5. Melden | Een JSON-POST naar de webhook als er iets mis is, iets hersteld is, iets vanzelf weer werkt, of een app minstens `-CrashThreshold` keer crashte — niet bij elke gezonde run. Een probleem dat blijft wordt na `-RenotifyHours` opnieuw gemeld |

Voor een klant wordt niets gesloten of verwijderd, en de app van een klant wordt nooit
gereset: geen `-RemoveOld`, `-Latest` alleen voor de host in stap 1a, geen processen van klanten die worden gestopt.
Elke melding noemt per bevinding de gebruiker (`Account`, met `Customer` op true voor een
klant) en onder `before` of het bij hem hersteld is (`Fixed`); de Teams-kaart zet
*hersteld bij deze gebruiker* bij elke regel.

**Parameters**

| Parameter | Beschrijving |
|-----------|--------------|
| `-Account` | Accounts om mee te testen — gebruikersnaam, UPN of `DOMEIN\gebruiker` (standaard: `itceadmin`; met meer accounts wordt elk getest, bv. `itceadmin,itce.user`) |
| `-App` | `Teams`, `Outlook`, `Copilot` (standaard: alle drie) |
| `-WebhookUrl` | n8n-webhook (productie-URL) waar het rapport naartoe wordt gePOST. Zonder deze logt de run alleen |
| `-WebhookToken` | Wordt meegestuurd als header `X-Watchdog-Token`; stel dezelfde waarde in als Header Auth op de n8n Webhook-node |
| `-IntervalMinutes` | Hoe vaak de taak draait (standaard: `30`) |
| `-RepairCooldownHours` | Minimale tijd tussen twee hostherstellingen, zodat een probleem dat hij niet kan oplossen niet elke run opnieuw wordt geprobeerd (standaard: `4`) |
| `-RenotifyHours` | Een probleem dat gelijk blijft na zoveel uur opnieuw melden (standaard: `12`) |
| `-UpdateHours` | Hoe vaak de nieuwste Teams / Outlook klaargezet worden en Edge Update op Copilot controleert (standaard: `6`; `0` zet bijwerken uit, ook voor ons eigen account; met `-NoRepair` wordt ook niets bijgewerkt) |
| `-CrashThreshold` | De crashes en hangs van een app melden zodra het er sinds de vorige run zoveel zijn (standaard: `1`, elke crash; `0` zet crashmeldingen uit) |
| `-NoRepair` | Alleen testen en melden, niets wijzigen |
| `-NoUserRepair` | De host en ons eigen account herstellen, maar nooit iets in de sessie van een klant draaien — hun problemen worden nog steeds gemeld, met hun naam |
| `-SkipLaunchTest` | Een app die niet draait niet starten; alleen de registratie controleren |
| `-NoUserLaunch` | Na het herstel van een app die een gebruiker niet kon openen, hem niet voor de gebruiker openen |
| `-NoDiagnostics` | `Get-M365AppsLog.ps1` niet draaien vóór een herstel |
| `-Install` | De watchdog naar `-WorkingDir` kopiëren en de taak **M365 App Watchdog** registreren, met de overige parameters als instellingen |
| `-Uninstall` | De taak en `-WorkingDir` verwijderen |
| `-TestNotification` | Eén testbericht naar de webhook sturen en stoppen |
| `-WorkingDir` | Watchdog, instellingen, status en logs (standaard: `C:\IT\AppWatchdog`) |

**Voorbeelden**

```powershell
# Eén run nu in deze console, alleen melden
.\Watch-M365Apps.ps1 -NoRepair

# Installeren op een sessiehost, met meldingen naar n8n
.\Watch-M365Apps.ps1 -Install -WebhookUrl 'https://n8n.example.com/webhook/m365-apps' -WebhookToken '<token>' -Confirm:$false

# De webhook van begin tot eind testen met de geïnstalleerde instellingen
.\Watch-M365Apps.ps1 -TestNotification

# Weer verwijderen
.\Watch-M365Apps.ps1 -Uninstall -Confirm:$false
```

**Wat n8n ontvangt**

```json
{
  "source": "Watch-M365Apps",
  "event": "repaired",
  "host": "AVD-0",
  "time": "2026-10-09T14:30:02.1234567+02:00",
  "summary": "AVD-0: 1 problem(s) found and repaired - itceadmin Outlook NotRegistered",
  "accounts": ["itceadmin"],
  "findings": [],
  "before": [{ "Account": "itceadmin", "App": "Outlook", "Problem": "NotRegistered", "Detail": "not registered for this user" }],
  "actions": ["itceadmin Outlook: re-registered as the user - result 0"],
  "crashes": [{ "App": "Teams", "Kind": "Crash", "Count": 2, "Last": "2026-10-09T14:12:40.0000000+02:00", "Exe": "ms-teams.exe", "Version": "26260.1704.5188.5238", "Module": "msedgewebview2.dll", "Code": "0xc0000005" }],
  "log": "C:\\IT\\AppWatchdog\\Logs\\Watch-M365Apps_20261009.log",
  "diagnostics": "C:\\IT\\AppWatchdog\\Diag\\M365AppsLog_AVD-0_20261009-1430.zip"
}
```

`event` is `repairing` (gevonden, wordt hersteld — verstuurd vóór het herstel, met `diagnostics`), `repaired`, `repair-failed`, `failing` (met `-NoRepair`), `crashed` (alleen
crashes deze run), `recovered`, `error` (de run zelf mislukte) of `test`; `crashes` gaat
met elk daarvan mee. `Problem` is `NotProvisioned` (host), `NotRegistered`,
`Broken` (bestanden weg of status niet `Ok`) of `WontStart` (ons account),
`WontOpen` (Windows weigerde hem te openen voor een gebruiker) of `RegisterFailed`. In n8n: een **Webhook**-node
(POST, Header Auth op `X-Watchdog-Token`), daarna routeren op `{{$json.body.event}}` naar
Teams, mail of een ticket.

**Installeren vanaf GitHub**

Een sessiehost heeft geen kopie van de repo nodig: download dit ene bestand en draai
`-Install`. Het hulpscript komt vanzelf mee — `-Install` haalt `Repair-AppxPackageStore.ps1`
van GitHub op de vastgepinde commit en controleert de SHA-256. De download hieronder is op
dezelfde manier vastgepind, op een commit en een hash, omdat wat hij installeert als System
draait. Draai het in een verhoogde PowerShell op de host:

```powershell
# Watch-M365Apps.ps1 op een vaste commit - beide regels samen bijwerken
$commit = 'f15d9b19e4ba643465664346a0a107dc485ac80c'
$sha256 = '1CB3145A65B8EA7B7A85810D12ED83740DC3A63DE897343F0FEBAAC59643CF12'
$file   = Join-Path $env:TEMP 'Watch-M365Apps.ps1'
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
Invoke-WebRequest "https://raw.githubusercontent.com/sjkanon/M365-Scripts/$commit/scripts/RDS/Watch-M365Apps.ps1" -OutFile $file -UseBasicParsing
if ((Get-FileHash $file -Algorithm SHA256).Hash -ne $sha256) { Remove-Item $file; throw 'SHA-256 klopt niet - niet uitgevoerd' }
# daarna, zoals in de voorbeelden: eerst -NoRepair om te kijken, dan installeren
& $file -Install -WebhookUrl '<n8n webhook URL>' -WebhookToken '<token>' -Confirm:$false
```

Op meerdere hosts tegelijk, vanaf je eigen pc via PowerShell remoting:

```powershell
# Dezelfde download op elke host, parallel
Invoke-Command -ComputerName avd-0, avd-1, avd-2 -ScriptBlock {
    $commit = 'f15d9b19e4ba643465664346a0a107dc485ac80c'
    $sha256 = '1CB3145A65B8EA7B7A85810D12ED83740DC3A63DE897343F0FEBAAC59643CF12'
    $file   = Join-Path $env:TEMP 'Watch-M365Apps.ps1'
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest "https://raw.githubusercontent.com/sjkanon/M365-Scripts/$commit/scripts/RDS/Watch-M365Apps.ps1" -OutFile $file -UseBasicParsing
    if ((Get-FileHash $file -Algorithm SHA256).Hash -ne $sha256) { Remove-Item $file; throw "SHA-256 mismatch on $env:COMPUTERNAME" }
    & $file -Install -WebhookUrl '<n8n webhook URL>' -WebhookToken '<token>' -Confirm:$false
}
```

Controleer daarna één host met `-TestNotification`. Naar een nieuwere versie: zet die
commit en de SHA-256 van het bestand op die commit
(`(Get-FileHash .\scripts\RDS\Watch-M365Apps.ps1).Hash` in een checkout ervan) in de twee
regels, na het lezen van de diff, en draai `-Install` opnieuw — tot dan blijft de taak zijn
eigen kopie in `C:\IT\AppWatchdog` draaien. De webhook-URL en het token komen nooit in
deze repo: die is openbaar.

**Opmerkingen**

- **Houd op elke host een sessie van `itceadmin` (en elk ander bewaakt account) open** (verbroken is prima). Een
  account dat niet is aangemeld wordt overgeslagen: zijn pakketten staan in zijn
  FSLogix-container en zijn zonder die container niet te testen. Zonder enige sessie draait
  de hostcontrole (stap 1) nog wel.
- Crash-events noemen geen gebruiker, dus een crash kan van een klant of van ons zijn. Een
  crash van de app in ons eigen account wordt in stap 2 opgevangen: de app draait niet meer,
  dus hij wordt opnieuw gestart. Op een drukke pool waar een losse Teams-crash ruis is:
  `-CrashThreshold` verhogen.
- Herstellen in de sessie van een klant gebruikt `conhost --headless`, zodat er geen
  console- of Windows Terminal-venster zou moeten verschijnen; gecontroleerd is dat het op de
  opdracht wacht en de exitcode niet doorgeeft (vandaar het teruglezen), maar **nog niet**
  bekeken in een echte klantsessie.
- `Get-AppxPackage -User` krijgt `DOMEIN\gebruiker`, geen SID: met een Entra ID-SID
  (`S-1-12-1-…`) antwoordt het *No valid SID could be determined*.
- De starttest start een app die niet draait in onze eigen sessie. Daar kan een venster
  verschijnen, en een gereset Teams vraagt ons account opnieuw aan te melden.
  `-SkipLaunchTest` zet dat uit.
- `-Install` beperkt `-WorkingDir` tot System en Administrators (de taak voert uit wat erin
  staat als System), kopieert dit script en `Repair-AppxPackageStore.ps1` daarheen — uit
  `..\Device` in een checkout van de repo, anders altijd van GitHub (nooit van naast een
  gedownloade kopie, een map waar elke gebruiker in kan hebben geschreven) op dezelfde vastgepinde commit en SHA-256 als
  [`Update-SessionHostImage.ps1`](#update-sessionhostimageps1) — en bewaart de webhook-URL en
  het token alleen in `config.json` daar, niet in de opdrachtregel van de taak. Een
  instelling wijzigen: `-Install` opnieuw draaien met alle parameters.
- Bewijs: `C:\IT\AppWatchdog\Diag`, één zip per nieuw probleem van `Get-M365AppsLog.ps1`, 14 dagen bewaard; het pad staat in `diagnostics` van de melding. `-Install` kopieert de verzamelaar naast de watchdog (uit de repo, of van GitHub op een vastgepinde commit en SHA-256); zonder hem herstelt en meldt de watchdog zoals voorheen.
- Logs: `C:\IT\AppWatchdog\Logs`, één bestand per dag, 14 dagen bewaard. Transcripts en
  `.reg`-back-ups van herstellingen: `C:\IT\AppWatchdog\Repair`.
- Verhoogd of als System draaien; het script start zichzelf opnieuw in 64-bit Windows
  PowerShell voor de AppX-cmdlets. Exitcode `0` gezond of hersteld, `1` er is nog iets stuk.

---

### Get-M365AppsLog.ps1

Verzamelt alles over de nieuwe Teams, de nieuwe Outlook en Copilot op een sessiehost in één
map en zip, om na een klacht twee vragen te beantwoorden: *waarom werkte de app niet bij
deze gebruiker*, en *waarom zag of herstelde de watchdog ([`Watch-M365Apps.ps1`](#watch-m365appsps1))
het niet*. Alleen lezend — er wordt niets gestart, geregistreerd, hersteld of gesloten.

**Wat het verzamelt**

| Deel | Wat |
|------|-----|
| Watchdog | `config.json` (webhook en token gemaskeerd), of `NoRepair` / `NoUserRepair` aan staat en de app bewaakt wordt, de geïnstalleerde versie en of die de controles per gebruiker überhaupt heeft; status, laatste run, resultaat en geschiedenis van de taak, achtergebleven probe-taken; `state.json` (laatste run, laatste hostherstel, open problemen, herregistraties per gebruiker); de logs en herstellogs uit de periode, met de regels over de betreffende apps en gebruikers eruit gelicht |
| Host | Klaargezette builds, de unified Copilot-app (Edge Update) en de WebView2-runtime |
| Alle gebruikers | `Get-AppxPackage -AllUsers`: elke gebruiker bij wie elk pakket bekend is en de installatiestatus — ook afgemelde gebruikers |
| Aangemelde gebruikers | Per gebruiker en app: de pakketten die voor hen geregistreerd zijn, status, bestanden aanwezig, draait in hun sessie, en het oordeel waar de watchdog op zou uitkomen — gemarkeerd als **BLIND SPOT** waar hij het goed zou vinden zonder het te testen |
| Events | Uit de laatste `-Hours`: TWinUI `5960`/`5961`, AppXDeploymentServer (fouten en `401`/`404`), AppXDeployment, AppxPackaging, AppReadiness, AppModel-Runtime, Application Error `1000` / Hang `1002` van de apps — elk met de gebruiker, de foutcode, en of de watchdog dat event überhaupt leest |
| Register | Een `PackageStatus` die niet 0 is (Windows heeft het pakket als kapot gemarkeerd), pakketten in `Deprovisioned`, FSLogix `InstallAppxPackages`, de Copilot-policy's van Edge Update; exports van de sleutels van FSLogix, de Edge Update-policy, Teams en Deprovisioned |
| FSLogix en Edge Update | Hun logbestanden uit de periode (FSLogix volgt `Logging\LogDir`), de regels in het Profile-log over de apps met een fout, de fout- en waarschuwingsevents van FSLogix |
| Copilot | Geïnstalleerde Copilot-apps met hun map, en welk Copilot-proces bij wie draait — `copilotapp.exe` is de unified app, `M365Copilot.exe` de verpakte |
| App-logs (`-IncludeAppLogs`) | Per aangemelde gebruiker binnen het bereik: Teams-logs (LocalCache — weg bij afmelden met FSLogix), het Teams-diagnosepakket in Downloads, logs van de nieuwe Outlook, FSLogix `AppxPackages.xml` |

Uitvoer in `-OutputPath` (anders `%TEMP%`): `summary.txt` (het scherm), `sessions.csv`,
`packages.csv`, `allusers.csv`, `events.csv`, `fslogix-events.csv`, `watchdog\`,
`registry\`, `fslogix\`, `edgeupdate\` en `applogs\`, gezipt.

**Blinde vlekken die het aanwijst**

- Met een watchdog van vóór 2026-10-10 (8): Copilot niet geregistreerd voor een gebruiker
  terwijl de unified app op de host staat wordt voor iedereen zonder test geaccepteerd, de
  unified app (`copilotapp.exe`) wordt nooit gestart, oudere zien ook zijn crashes niet, en
  consumenten-Copilot (`Microsoft.Copilot`) telt evenveel als Microsoft 365 Copilot. Vanaf
  die versie telt de watchdog alleen de unified app, controleert hij het identiteitspakket
  per gebruiker en start hij hem voor ons eigen account — de verzamelaar zegt welke hij vond.
- De app van een klant geregistreerd en `Ok` maar werkt niet: de watchdog start niets voor
  klanten, dus hij ziet alleen een opening die Windows weigerde (TWinUI `5961`).
- Een gebruiker die korter dan 10 minuten is aangemeld, een afgemelde gebruiker, het
  bewaakte account dat niet op de host is aangemeld, een taak die niet gedraaid heeft, een
  oudere watchdog zonder de controles per gebruiker, events die de watchdog niet leest.

**Parameters**

| Parameter | Beschrijving |
|-----------|--------------|
| `-User` | Alleen deze gebruikers — gebruikersnaam, UPN of `DOMAIN\user`, ook als ze afgemeld zijn (standaard: iedereen) |
| `-App` | `Teams`, `Outlook`, `Copilot` (standaard: alle drie) |
| `-Hours` | Hoe ver terug events en logs gelezen worden (standaard: `24`, hooguit `336`) |
| `-OutputPath` | Waar de map en zip komen (standaard: `C:\Temp`) |
| `-WorkingDir` | De map van de watchdog (standaard: `C:\IT\AppWatchdog`) |
| `-IncludeAppLogs` | Ook de eigen logs van de apps per aangemelde gebruiker kopiëren — daar staan namen en mailadressen in, dus alleen op verzoek |
| `-NoZip` | De map laten staan, niet zippen |

**Voorbeelden**

```powershell
# Een klant zegt dat Copilot vanochtend niet openging
.\Get-M365AppsLog.ps1 -User jansen -App Copilot -Hours 12

# Alles op deze host van de laatste twee dagen, met de logs van Teams en Outlook
.\Get-M365AppsLog.ps1 -Hours 48 -IncludeAppLogs
```

**Opmerkingen**

- Verhoogd draaien op de host, kort na de klacht; het start zichzelf opnieuw in 64-bit
  Windows PowerShell voor de AppX-cmdlets. Menu-item `Q`.
- Elk deel draait op zichzelf: een deel dat mislukt wordt aan het eind genoemd en de rest
  wordt toch verzameld. `Get-AppxPackage` draait in een apart proces en wordt na 120
  seconden opgegeven, omdat het juist op de host waarvoor dit bedoeld is kan blijven
  hangen. Een eventkanaal dat op de host niet bestaat (geen FSLogix, een oudere build)
  wordt overgeslagen.
- Logs die nog open staan worden gedeeld gelezen; van een log boven 50 MB blijft de laatste 50 MB.
- De webhook-URL en het token zijn gemaskeerd in de kopie van `config.json`.
- Vergelijkbaar met de AppX- en FSLogix-delen van Microsofts MSRD-Collect, maar beperkt tot
  deze drie apps en afgezet tegen wat de watchdog doet.
