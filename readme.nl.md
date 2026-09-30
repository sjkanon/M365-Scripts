[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

# M365-Scripts

> Een verzameling PowerShell-scripts en M365-beheertools voor MSP-engineers, onderhouden door Sjoerd Kanon.

---

## Inhoud

- [Mappen](#mappen)
- [Aan de slag](#aan-de-slag)
- [Een script vinden](#een-script-vinden)
- [Snelstarter](#snelstarter)
- [Vereisten](#vereisten)
- [Menu](#menu)
- [Scriptcategorieën](#scriptcategorieën)
  - [M365-beheer](#️-m365-beheer)
  - [Exchange](#-exchange)
  - [Entra ID / Graph](#-entra-id--graph)
  - [Intune en Autopilot](#-intune-en-autopilot)
  - [SharePoint en OneDrive](#-sharepoint-en-onedrive)
  - [Testen en diagnose](#-testen-en-diagnose)
  - [Rapportage](#-rapportage)
  - [Infrastructuur en apparaten](#️-infrastructuur-en-apparaten)
  - [Eigen tools](#-eigen-tools)
  - [Azure-infrastructuur](#️-azure-infrastructuur)
  - [Herschreven legacy-toolkits](#️-herschreven-legacy-toolkits)
- [Repositorystructuur](#repositorystructuur)
- [Bijdragen](#bijdragen)
- [Versiegeschiedenis](#versiegeschiedenis)

---

## Mappen

Elke workload heeft een eigen map onder [`scripts/`](scripts/readme.nl.md), en elke map heeft een readme: waar elk script voor dient, de parameters, voorbeelden en opmerkingen. Begin hier en klik door; bovenaan elke readme staat een kruimelpad om weer omhoog te gaan.

| Map | Omschrijving |
|--------|-------------|
| [`ActiveDirectory/`](scripts/ActiveDirectory/readme.nl.md) | Monitoring van on-prem AD DS (bewaking van accountvergrendelingen) — richt zich rechtstreeks op een DC/fileserver, niet op Entra ID |
| [`Azure/`](scripts/Azure/readme.nl.md) | Beheer van Azure IaaS-VM's (conversie van de schijfcontroller) — richt zich rechtstreeks op Azure via `Az`, niet op de M365-tenant |
| [`Entra/`](scripts/Entra/readme.nl.md) | Gebruikerslevenscyclus, managers toewijzen, licentierapportage, Conditional Access-baseline, tijdelijke CA-vensters, TAP-codes, audit van M365-groepen (Microsoft Graph) |
| [`Exchange/`](scripts/Exchange/readme.nl.md) | Agendamigratie/-rechten, distributiegroepen, audits van mailboxen/agenda's/DKIM/doorsturen |
| [`Graph/`](scripts/Graph/readme.nl.md) | Beheer van Microsoft Graph-toepassingsmachtigingen |
| [`Intune/`](scripts/Intune/readme.nl.md) | Autopilot-inschrijving, updater voor het iOS-compliancebeleid, uitrol van bedrijfsachtergrond/-vergrendelscherm |
| [`SharePoint/`](scripts/SharePoint/readme.nl.md) | Contentbewerkingen in SharePoint Online / OneDrive — prullenbak terugzetten per site of tenantbreed (PnP PowerShell, automatische app-registratie) |
| [`Reporting/`](scripts/Reporting/readme.nl.md) | Rapport laatste aanmelding van computers, SharePoint-opslagrapport, maandelijks licentierapport |
| [`Device/`](scripts/Device/readme.nl.md) | Onderhoud van Windows-endpoints — activatie, opschonen, tijdelijke bestanden, tijdsynchronisatie, audio, OpenVPN-diagnose, tijdelijke schijf + pagefile op Azure/AVD |
| [`Network/`](scripts/Network/readme.nl.md) | Controle van TCP-poorten, diagnose van authenticatie/netwerk, stresstest van bestands-I/O |
| [`RDS/`](scripts/RDS/readme.nl.md) | Diagnose van aanmeldingen op RDP / RD Web Access en live monitoring van sessies |
| [`SMTP/`](scripts/SMTP/readme.nl.md) | Connectiviteitstests voor een SMTP-relay (eenmalig en terugkerend) |
| [`Deployment/`](scripts/Deployment/readme.nl.md) | USB-toolkit voor Windows-installatie en Autopilot-inschrijving tijdens OOBE |
| [`DNS/`](scripts/DNS/readme.nl.md) | DNS-records opzoeken en importeren in AD-geïntegreerde DNS-zones |
| [`SAS/`](scripts/SAS/readme.nl.md) | Foutmonitoring van SAS-batchjobs met Zabbix-integratie |
| [`Teams/`](scripts/Teams/readme.nl.md) | Export en archivering van Microsoft Teams / SharePoint |
| [`Startup/`](scripts/Startup/readme.nl.md) | Functiebibliotheek `functies.ps1` voor M365 + module-bootstrap + syntaxcontrole, gedot-sourcet door het menu |
| [`Custom Scripts/`](scripts/Custom%20Scripts/readme.nl.md) | Scripts die aan hun pad vastzitten — uitrol van het Office-thema (de download-URL wijst hard naar dit pad in de repo) |
| [`TenantOnboarding/`](scripts/TenantOnboarding/readme.nl.md) | Inrichten van nieuwe tenants, multi-tenant-/GDAP-rapportage, app-uitrol, apparaatconfiguratie, OneDrive-beheer, gebruikersbeheer — gemoderniseerd vanuit een uitgefaseerde interne toolkit voor tenantinrichting |
| [`Office365Toolkit/`](scripts/Office365Toolkit/readme.nl.md) | Herschreven Security-/Exchange-/Intune-functies die nog nuttig waren uit de uitgefaseerde toolkit `directorcia/Office365` (CIAOPS) |
| [`PatronToolkit/`](scripts/PatronToolkit/readme.nl.md) | Herschreven Entra-/Exchange-/Intune-/Security-/SharePoint-/Teams-functies die nog nuttig waren uit de uitgefaseerde toolkit `directorcia/patron` |
| [`LegacyUtilities/`](scripts/LegacyUtilities/readme.nl.md) | Diverse gemoderniseerde scripts (Exchange, Entra, Teams, Network, Device, Workspace 365) uit allerlei kleine tools van de uitgefaseerde interne toolkit |

Ken je de naam van het script maar niet de map? [`scripts/INDEX.md`](scripts/INDEX.md) zet elk script van A tot Z op een rij.

---

## Aan de slag

```powershell
.\load.ps1
```

Bij de eerste keer starten doet `load.ps1` het volgende:

1. Het vraagt je admin-UPN en weergavenaam — die worden opgeslagen in een `load.config.ps1` die door git wordt genegeerd
2. Het vraagt of je standaard de gedelegeerde GDAP-modus wilt gebruiken, en slaat eventueel een standaard klantdomein op
3. Het vraagt of Graph standaard met device code moet aanmelden
4. Het detecteert ontbrekende modules en biedt aan ze automatisch te installeren
5. Het importeert alle benodigde modules
6. Het opent het interactieve menu

Daarna start het direct, zonder vragen.

Om de launcher automatisch te starten bij het aanmelden in Windows:

```powershell
.\load.ps1 -SetupStartup
```

Om de opstartsnelkoppeling later weer te verwijderen:

```powershell
.\load.ps1 -RemoveStartup
```

> Je kunt `.\menu.ps1` ook rechtstreeks starten — dan vraagt het als terugvaloptie om je UPN.
> Modules handmatig opnieuw installeren of bijwerken: `.\scripts\Startup\Install-Modules.ps1`

---

## Een script vinden

| Waar | Wat het je oplevert |
|-------|-------------------|
| [`scripts/INDEX.md`](scripts/INDEX.md) | Elk script van A tot Z op één pagina — naam, map en wat het doet. Gebruik Ctrl-F hierop als je ongeveer weet wat je zoekt, maar niet waar het staat |
| [`scripts/readme.md`](scripts/readme.nl.md) | De andere richting: waar elke workloadmap voor dient |
| [`.\menu.ps1`](menu.ps1) | De samengestelde interactieve launcher voor de dagelijkse taken |
| `f <term>` | Fuzzy zoeken vanuit je shell, beschreven onder [Snelstarter](#snelstarter) hieronder |
| Elke map-readme | Elke scriptnaam in een `Scripts`-tabel linkt rechtstreeks naar het bestand, met een `docs`-link naar de bijbehorende sectie op dezelfde pagina |

`INDEX.md` wordt gegenereerd uit de eigen `.SYNOPSIS`-headers van de scripts door [`scripts/Startup/Update-ScriptIndex.ps1`](scripts/Startup/Update-ScriptIndex.ps1) — voer het opnieuw uit (of met `-Check`) telkens wanneer een script wordt toegevoegd, hernoemd, verplaatst of verwijderd.

Die links worden gecontroleerd, niet verondersteld: [`scripts/Startup/Test-MarkdownLinks.ps1`](scripts/Startup/Test-MarkdownLinks.ps1) loopt elke readme langs en faalt op een bestand dat er niet is, of een anker zonder kop erachter.

---

## Snelstarter

`f.ps1` maakt één kort commando per script in [scripts/](scripts/), zodat je niet eerst naar een map hoeft te navigeren. Dot-source het vanuit je profiel:

```powershell
notepad $PROFILE
. "C:\Users\<you>\Git\M365-Scripts\f.ps1"
```

De punt vooraan is niet optioneel — zonder die punt draait het bestand in een eigen scope en zijn de commando's meteen weer weg. `.\f.ps1 -Install` schrijft de regel voor je (en doet niets als hij er al staat).

Na het herstarten van je shell:

```powershell
f-test-dkimconfig -Domain contoso.com
f-dkimconfig -Domain contoso.com          # korte vorm, als het zelfstandig naamwoord uniek is
f-get-mailboxsizes
f-m365                                    # de hele catalogus
f-m365 mailbox                            # gefilterd
```

De wrappers kopiëren het parameterblok van het doelscript uit de AST, zodat `-Dom<Tab>` aanvult en een `ValidateSet` wordt afgedwongen voordat er iets draait. Standaardwaarden worden bewust uit de wrapper weggelaten: alleen gebonden parameters worden doorgegeven, zodat de eigen standaardwaarden van het script blijven gelden.

### Fuzzy zoeken

Voor als je ongeveer weet hoe een script heet maar niet precies, zoekt `f` op naam, map en `.SYNOPSIS`:

```powershell
f dkim                    # één treffer -> voert het uit
f entra group             # meerdere treffers -> genummerde keuzelijst
f dkim -Domain contoso.com
f -List mailbox           # toon treffers, voer niets uit
f -Show trace             # pad, synopsis en parameters
f -Edit bloatware         # open in $env:EDITOR, VS Code of notepad
```

### Opmerkingen

| | |
|---|---|
| Naamgeving | `f-<volledige-scriptnaam>` bestaat altijd; `f-<noun>` wordt alleen toegevoegd waar het eenduidig blijft. Het werkwoord weghalen levert hier 11 botsingen op (`Detect-`, `Install-` en `Uninstall-ClaudeDesktop-Intune` worden hetzelfde zelfstandig naamwoord), dus op de volledige naam kun je altijd rekenen. |
| Cache | Gegenereerde wrappers staan in `.f-index.json` (door git genegeerd). Elk script parsen kost ~500 ms, de cache lezen ~30 ms, en daardoor blijft het starten van de shell snel. De cache ververst zichzelf wanneer een script wordt toegevoegd, verwijderd of gewijzigd; `f-refresh` dwingt het af. |
| Botsingen | Bestaande commando's worden nooit overschreven. Een profiel kan meer dan één van deze catalogi laden — `ScriptRunner.Profile.ps1` in *itce-testing* is eigenaar van `f-scripts`, en daarom heet deze catalogus `f-m365`. Alles wat wordt overgeslagen, wordt bij het laden gemeld. |
| Verwijderen | `.\f.ps1 -Uninstall` haalt het gemarkeerde blok uit `$PROFILE`. Een met de hand geschreven dot-source-regel wordt gemeld, niet verwijderd. |

---

## Vereisten

| Vereiste | Details |
|-------------|---------|
| PowerShell | 7.0+ (cross-platform); afzonderlijke scripts ondersteunen PS 5.1 op Windows |
| Rechten | Microsoft 365-beheerrechten voor de betreffende workload |
| Execution Policy | Alleen Windows: `Set-ExecutionPolicy RemoteSigned -Scope CurrentUser` |

---

## Menu

De launcher (`menu.ps1`) dekt alle tools in deze repo. Druk op een toets om te starten:

| Toets | Categorie | Tool |
|-----|----------|------|
| `1` / `F1` | Testing | Test-Ports — controle van TCP-poorten |
| `2` / `F2` | Exchange | Migrate-Calendar |
| `3` / `F3` | Exchange | Set-Calendar-rights |
| `4` / `F4` | Testing | Test-SMTP (eenmalig) |
| `5` / `F5` | Testing | Test-SMTP (elke 5 min) |
| `6` / `F6` | Device | Restart-Time-Sync |
| `7` / `F7` | Device | Detect-AudioDevices |
| `8` / `F8` | Device | Disable-InternalMic |
| `I` | Device | Remove-OemBloatware — OEM- en generieke Store-bloatware verwijderen |
| `T` | Device | Update-TeamsClient — de nieuwe Teams en de Outlook-invoegtoepassing voor vergaderingen bijwerken als ze verouderd zijn |
| `R` | Device | Repair-AppxPackageStore — AppX-pakketten repareren die falen met 0x80070490 (Teams, nieuwe Outlook, FSLogix) |
| `9` / `F9` | Startup | Install-Modules |
| `X` | Startup | Update-ScriptIndex — [`scripts/INDEX.md`](scripts/INDEX.md) opnieuw opbouwen, de A–Z-lijst van alle scripts |
| `L` | Startup | Test-MarkdownLinks — elke readme-link controleren: bestanden en ankers binnen de pagina |
| `M` | Startup | Convert-MarkdownToHtml — een opgemaakte HTML-pagina bouwen uit een markdown-document, voor IT Glue |
| `A` / `F10` | Reporting | Licensing-Report |
| `P` | Reporting | SharePoint-Perms — rapporteren wie waar toegang toe heeft, op elk niveau |
| `S` | SharePoint | SharePoint-Structure — metadata, bibliotheken en rechten inrichten/controleren |
| `F` | Startup | Enable-LauncherStartup — launcher toevoegen aan Opstarten van Windows |
| `G` | Startup | Disable-LauncherStartup — launcher verwijderen uit Opstarten van Windows |
| `B` | M365 | Connect-Tenant |
| `H` | M365 | Test-GdapConnection — gedelegeerde GDAP-toegang valideren |
| `C` | M365 | Submenu Exchange Online |
| `D` | M365 | Submenu Entra ID / Graph |
| `E` | M365 | Submenu MSP Admin |

De M365-opties (`B`, `C`, `D`, `E`, `H`) laden `functies.ps1` pas bij het eerste gebruik — Graph-authenticatie wordt alleen gestart als het nodig is.

**Exchange-submenu (`C`)**

| Toets | Tool |
|-----|------|
| `8` | Test-CalendarPermissions — rechten op agendamappen auditen (alle mailboxen of één) |
| `9` | Test-MailboxPermissions — Full Access, Send As en Send on Behalf auditen |
| `A` | Test-GroupPermissions — beheerders van distributiegroepen, Send As, Send on Behalf en aantallen leden auditen |
| `B` | Test-DkimConfig — DKIM-ondertekeningsconfiguratie en DNS-records (CNAME/TXT) valideren |
| `C` | Get-ExternalForwards — mailboxen met extern doorsturen auditen |
| `D` | Get-MailboxSizes — rapport van mailboxgroottes, gesorteerd op gebruikte opslag |
| `E` | Move-InboxToArchive — berichten uit Postvak IN archiveren naar de map Archief |
| `F` | Set-DL-Dynamic-Static — een dynamische distributiegroep omzetten naar een statische groep |
| `H` | Get-CalendarMappings — waar een agenda in Outlook is toegevoegd, naast de rechten (zoeken op trefwoord, bv. `balie`, of alle/geselecteerde mailboxen) |
| `I` | Convert-SharedCalendar — een gedeelde agenda uit de mailbox van een gebruiker verhuizen naar een ruimte-/apparatuurmailbox (toont altijd eerst een voorbeeld) |
| `J` | Move-SharedCalendar — alles in één: een agenda op trefwoord zoeken, verhuizen naar een resourcemailbox, en oplijsten wie moet overstappen |
| `K` | Get-DLMembers — elke distributielijst met haar leden exporteren naar Excel, of alleen de lijsten met één adres, één domein of een domeinboom (`-Recurse` om geneste lijsten uit te klappen) |
| `L` | Restore-MailboxMessages — mail terugzetten die op een bepaalde dag, of vanaf een datum tot nu, is verplaatst of verwijderd, en tonen wie het deed (toont altijd eerst een voorbeeld) |

**Entra ID-submenu (`D`)**

| Toets | Tool |
|-----|------|
| `A` | Test-M365GroupMembership — eigenaren en leden van M365-groepen / Teams auditen |
| `B` | New-M365User — één nieuwe gebruiker aanmaken (automatisch gegenereerd wachtwoord, optioneel licentie) |
| `C` | Import-M365Users — gebruikers in bulk aanmaken vanuit CSV, standaard een proefdraai |
| `D` | New-TemporaryCA — tijdelijk Conditional Access-beleid aanmaken voor een gebruiker/groep (duur of begin-/einddatum en -tijd) |
| `E` | Remove-TemporaryCA — verlopen of alle tijdelijke CA-beleidsregels verwijderen |
| `F` | New-UserTAP — een Temporary Access Pass aanmaken voor een gebruiker |
| `G` | Get-M365UserLicenses — toegewezen licenties rapporteren voor een reeks gebruikers |
| `H` | Import-CA-Baseline — de community-baseline voor Conditional Access importeren |
| `I` | Set-UserManager — manager rapporteren/in bulk instellen voor een reeks gebruikers |

---

## Scriptcategorieën

### ☁️ M365-beheer

Interactieve M365-beheerfuncties via Microsoft Graph en Exchange Online. Worden als bibliotheek geladen via het menu. CSV- en logexports gaan naar `C:\Temp\` op Windows of `~/Downloads/` op macOS.

| Gebied | Functies |
|------|----------|
| Exchange Online | Toegang tot gedeelde mailboxen, landinstelling, aliassen, distributiegroepen, automatisch antwoord, kopie van verzonden items |
| Entra ID / Graph | Tenantbeheerders, domeinen, licenties, gebruikers, wachtwoord resetten, aanmeldlogboeken, bulk aanmaken/verwijderen, tijdelijke CA-vensters, TAP-codes |
| MSP Admin | MSP-beheeraccount aanmaken/beheren over klanttenants heen |

---

### 📧 Exchange

Scripts voor agenda- en mailboxbeheer.

- Agendamigratie tussen gebruikers
- Rechten op agendamappen instellen (ondersteuning voor NL/FR/EN-landinstelling)
- **Get-DistributionGroupMembers.ps1** — wie in welke distributielijst zit, als één Excel-werkmap die zo naar de klant kan
  - Werkblad `Overzicht` (één rij per lijst) en werkblad `Leden` (één rij per lid), beide filterbare tabellen met een vastgezette kopregel, kopteksten in het Nederlands
  - `-Member jan@contoso.com` beantwoordt "in welke lijsten zit deze persoon?"; `-Member @be.verizon.com` beantwoordt dat voor één domein en `-Member *.verizon.com` voor een domein en al zijn subdomeinen, waarbij aliassen en `ExternalEmailAddress` meetellen zodat externe contactpersonen echt gevonden worden
  - `-Recurse` klapt geneste lijsten uit — zonder die optie is iemand die alleen via een geneste groep mail ontvangt onzichtbaar, en meldt een filter "geen treffers" op een lijst die wel bij hem aflevert
- **Get-MessageTraceReport.ps1** — nagaan wie wat heeft ontvangen, op welk exact tijdstip, en waarnaartoe het is doorgestuurd
- **Remove-PhishingMessage.ps1** — een phishingbericht verwijderen uit één, meerdere of alle mailboxen; standaard een proefdraai
  - Twee engines: **Purview** Content Search + purge (tenantbreed, de enige die HardDelete kan) en **Graph** (per mailbox, geen vertraging van de zoekindex, rapport per bericht)
  - `Recycle` / `SoftDelete` / `HardDelete`; weigert te draaien zonder inhoudsselector, zodat een datumbereik alleen nooit op elk bericht kan matchen
  - Herhaalt purge-rondes automatisch om de limiet van Purview van 10 items per mailbox heen, en schrijft een CSV van alles wat gematcht en verwijderd is
- **Restore-MailboxMessages.ps1** — berichten terugzetten die op een bepaalde dag zijn verplaatst of verwijderd, en rapporteren wie het deed; standaard alleen een voorbeeld
  - Verwijderde berichten gaan terug via `Restore-RecoverableItems` (Deleted Items, Recoverable Items, Purges); verplaatste berichten worden via het auditlogboek naar hun oorspronkelijke map herleid en via Graph teruggezet
  - Noemt de actor uit het Unified Audit Log — account, eigenaar/gemachtigde/beheerder, client, IP — en vermeldt welke acties op de mailbox niet worden geaudit

---

### 👤 Entra ID / Graph

Scripts voor het beheer van de gebruikerslevenscyclus via Microsoft Graph.

- Eén M365-gebruiker aanmaken (automatisch gegenereerd wachtwoord, optioneel licentie)
- Gebruikers in bulk aanmaken vanuit CSV — standaard een proefdraai, wachtwoorden in de CSV-uitvoer
- Gebruikers in bulk verwijderen vanuit CSV — standaard een proefdraai, CSV-rapport

---

### 📱 Intune en Autopilot

Scripts voor apparaatinschrijving, Autopilot-registratie en beheer van compliancebeleid.

- Windows Autopilot-hardware-informatie ophalen
- CMD-hulpmiddel voor Autopilot-inschrijving
- **Compare-IntuneConfig.ps1** — de Intune-configuratie van een klanttenant vergelijken met een back-up van een MSP-baseline (driftdetectie), via de module `IntuneBackupAndRestore` — alleen-lezen
- **iOS Compliance Updater** — houdt de minimale iOS-versie in Intune automatisch up-to-date
  - Haalt de nieuwste iOS-versie op uit de RSS-feed van Apple (met terugval op de Apple Support-pagina)
  - Vergelijkt die met het huidige minimum in het beleid en past het aan via de Microsoft Graph API
  - Eenmalige inrichting via `Setup.ps1` (maakt de App Registration aan, kent rechten toe, schrijft `config.json`)
  - Draait wekelijks als geplande taak in Windows (SYSTEM, elke maandag om 07:00)
  - Proefdraaimodus (`-WhatIf`) — toont wat er zou veranderen zonder het toe te passen
- **Desktop** — vergrendelscherm aan Start en bureaublad toevoegen; bedrijfsachtergrond instellen via Intune:
  - `Set-CorporateWallpaper.ps1` — generiek, herbruikbaar per klant; alleen het CONFIGURATION-blok moet worden aangepast
  - `Make-lockscreen.ps1` — past dezelfde bedrijfsafbeelding toe als Windows-vergrendelscherm via PersonalizationCSP
  - Downloadt de achtergrond van een openbare URL; vergelijkt de SHA256-hash met het bestaande bestand — slaat over als hij al actueel is, past toe als hij nieuw of gewijzigd is
  - De vergrendelschermflow downloadt van internet via `Invoke-WebRequest`, valideert afbeeldingsheaders (`jpg/png/bmp`), blokkeert HTML-antwoorden en normaliseert gangbare GitHub-blob/raw-URL's
  - Past toe via PersonalizationCSP (MDM-afdwinging), WinAPI (direct), HKCU-register (stijl) en het Default User-profiel (nieuwe accounts)
  - Log: `C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\CorporateWallpaper-<CLIENTNAME>.log`
  - Log vergrendelscherm: `C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\CorporateLockscreen-<CLIENTNAME>.log`
  - Uitrollen via Intune: **Run as SYSTEM**, 64-bit PowerShell

  | Variabele | Omschrijving |
  |---|---|
  | `$ImageUrl` | Openbare URL naar de achtergrondafbeelding (PNG of JPG) |
  | `$WallpaperStyle` | `10` = Vullen · `6` = Passend · `2` = Uitrekken · `0` = Naast elkaar · `22` = Over meerdere schermen |
  | `$ClientName` | Klantnaam — gebruikt in de naam van het logbestand en het lokale afbeeldingspad |

---

### 📁 SharePoint en OneDrive

Contentbewerkingen op SharePoint Online-sites en OneDrive via PnP PowerShell.

#### Gebruikerstoegang intrekken

**Revoke-SharePointUserAccess.ps1** — de tegenhanger van het rechtenrapport: dat vertelt wie waar bij kan, dit script neemt het weg. Rapporteert standaard, verwijdert met `-Apply`, en schrijft een CSV van elke gevonden toekenning en wat ermee gebeurd is.

- Eerst de sitecollectiebeheerder, omdat die elke roltoewijzing daaronder overstijgt
- Directe roltoewijzingen op de site, een subsite, een lijst of bibliotheek, een map of één bestand
- Lidmaatschap van SharePoint-groepen, en **deellinks** — de `SharingLinks.*`-groepen waar "Iedereen met de link" en "Specifieke personen" iemand in werkelijkheid in zetten
- Het wijzigt bewust nooit het lidmaatschap van Entra ID-groepen: een gebruiker die binnenkomt via een beveiligingsgroep of Microsoft 365-groep houdt die toegang, en verwijderen uit SharePoint neemt die niet weg. Die routes worden gerapporteerd met de naam van de groep, zodat offboarding twee stappen is en de tweede zichtbaar
- Toekenningen aan `Everyone` blijven om dezelfde reden, maar omgekeerd, ongemoeid — er één verwijderen trekt de toegang in voor de hele tenant, niet voor deze persoon
- `ConfirmImpact = 'High'`, dus het vraagt per verwijdering om bevestiging, tenzij `-Confirm:$false`

**Test-SharePointAccessScripts.ps1** verifieert dit script en het rechtenrapport zonder een tenant aan te raken: de app-only-authenticatielaag die beide delen moet byte-identiek blijven, en de intrekkingstrechter moet een proefdraai vastleggen zonder hem uit te voeren, onder `-Apply` uitvoeren en vastleggen, en de toekenningen die hij niet mag verwijderen blijven weigeren.

#### Prullenbak terugzetten

Verwijderde bestanden en mappen terugzetten uit de prullenbak van een site of OneDrive — standaard een proefdraai, `-Apply` om echt terug te zetten.

- Eén site (`-SiteUrl`, werkt ook voor OneDrive) of elke SharePoint-site in de tenant (`-AllSites`) — de tenantbrede ronde slaat OneDrive-, systeem- en vergrendelde sites over, en een falende site breekt de run niet af
- Beperk de ronde met `-SiteFilter` en probeer het eerst op een handvol sites met `-MaxSites`
- Filteren op naam, oorspronkelijke map, wie het verwijderde en een tijdvenster van verwijdering
- Prullenbak van de eerste fase (gebruiker) en de tweede fase (sitecollectie), of beide
- Zet eerst mappen terug, ondiepe paden eerst — een bestand kan niet worden teruggezet in een map die zelf nog verwijderd is
- Maakt de benodigde Entra-app-registratie automatisch aan bij de eerste run tegen een tenant, en cachet dan de client-ID in `pnp.appid.json` (door git genegeerd) — latere runs gaan meteen naar de interactieve aanmelding
- `-GrantSiteAdmin` maakt je per site tijdelijk sitecollectiebeheerder en verwijdert die rechten achteraf weer — nodig voor de OneDrive van een andere gebruiker en in de praktijk vereist voor `-AllSites`
- Zet terug in batches van maximaal 200 via één serveraanroep (`-BatchSize`) — een mislukte batch valt terug op item voor item, zodat één slecht bestand de rest niet meesleurt
- Tijdmeting overal: hoe lang het lezen van de prullenbak duurde, een schatting vooraf, een voortgangsbalk met live ETA, en de werkelijke duur in de samenvatting
- CSV-rapport van elk item, teruggezet of mislukt, inclusief de site, de SharePoint-fout, het batchnummer en hoe lang het duurde

#### Structuur inrichten

Een complete SharePoint-structuur inrichten en onderhouden — metadatamodel, inhoudstypen, bibliotheken en groepsrechten — vanuit één JSON-config. Zie [`scripts/SharePoint/Provisioning/`](scripts/SharePoint/Provisioning/readme.nl.md).

- `Install-SharePointStructure.ps1` — **de opbouw in één commando**: registreert zelf de Entra-app, voert de drie inrichtingsstappen uit in de enige volgorde die werkt, verifieert het resultaat, en verwijdert met `-TemporaryApp` de app-registratie weer, zodat er niets achterblijft in een tenant die je niet dagelijks beheert
- Het model zit in de config, niet in de code: een tweede MSP-klant is een tweede configbestand, geen tweede fork van vier scripts
- `New-SharePointMetadata.ps1` — termenset voor beheerde metadata, sitekolommen en inhoudstypen, op **elke** site in de config (een privékanaal in Teams is een eigen sitecollectie, en een sitekolom reikt daar niet overheen)
- `Set-SharePointLibraries.ps1` — bibliotheken, kanaalmappen in Teams, koppeling van inhoudstypen, volgorde van inhoudstypen per map, standaardkolomwaarden, gegroepeerde weergaven, en één Entra ID-beveiligingsgroep per pijler per toegangsniveau; `-EnsureGroups` maakt de groepen onderweg aan
- `Update-SharePointShareStatus.ps1` — leidt een kolom Deelstatus af uit de rechten die werkelijk op elk bestand staan (Anyone-link, gast, organisatielink of niets) en markeert alles met de tag Intern/Vertrouwelijk dat achter een externe link staat; exitcode 2 voor een geplande RMM-taak
- `Test-SharePointStructure.ps1` — alleen-lezen driftcontrole die elk verschil indeelt als Missing / Different / Extra; exitcode 2 betekent dat iemand iets heeft gewijzigd
- Alle vier zijn idempotent en ondersteunen `-WhatIf`; interactief of app-only met een certificaat
- Merkoverstijgende weergaven (`Scope = RecursiveAll`) maken "merk als tag" echt: *Alles - Butterstone* is één platte lijst over elke pijlermap heen, inclusief alles met de tag **Beide** — één bestand, twee merken, geen kopieën. Plus *Nog te taggen*, *Extern gedeeld* en *Te archiveren*
- [`Petsolutions-SharePoint-Handleiding.md`](scripts/SharePoint/Provisioning/Petsolutions-SharePoint-Handleiding.md) — Nederlandstalige eindgebruikersdocumentatie om aan de klant te geven: de drie manieren om een bestand toe te voegen en waarom ze zich anders gedragen, wat elk label betekent, en wat er gebeurt op het moment dat je iets tagt
- Gedocumenteerd in plaats van verstopt: unieke rechten op een map van een **standaard**kanaal zijn wat dit model vraagt en wat Microsoft niet ondersteunt — leden blijven het kanaal zien en krijgen een fout op het tabblad Bestanden. `-SkipChannelFolderPermissions` is het voorzichtige alternatief

---

### 🧪 Testen en diagnose

Audit- en diagnosescripts, ingedeeld per workload. Maken waar van toepassing zelf verbinding — hergebruiken een bestaande sessie of verbinden automatisch. CSV-exports gaan naar `C:\Temp\` op Windows of `~/Downloads/` op macOS.

#### Exchange Online

- Rechten op agendamappen auditen (onafhankelijk van de landinstelling, exporteert CSV)
- Delegatie via Full Access, Send As en Send on Behalf auditen (exporteert CSV)
- Beheerders van distributiegroepen, Send As, Send on Behalf en aantallen leden auditen (exporteert CSV)
- DKIM-ondertekeningsconfiguratie en DNS-records (CNAME/TXT) valideren; somt de vereiste acties op
- Mailboxen auditen die extern doorsturen naar domeinen buiten de tenant (beveiligingsaudit, exporteert CSV)
- Mailboxgroottes en aantallen items rapporteren, gesorteerd op gebruikte opslag (exporteert CSV)

#### Entra ID / Graph

- Eigenaren en leden van M365-groepen (incl. Teams) auditen — één rij per vermelding, exporteert CSV
- Tijdelijk Conditional Access-beleid aanmaken voor installatievensters (duur of exact lokaal begin/einde)
- Tijdelijk CA-beleid automatisch opruimen op het eindtijdstip (zelfde sessie) en een opruimscript voor gemiste sessies
- Temporary Access Pass-codes (TAP) aanmaken voor onboarding/ondersteuning van gebruikers

#### SharePoint Online

- Opslaggebruik rapporteren over alle sites in een tenant — huidige bestandsgroottes + versiegeschiedenis per bibliotheek en per bestand
- In twee fasen: eerst alle sites en documentbibliotheken opsommen, daarna de opslaggegevens ophalen
- Snelle modus (alleen quotagegevens) of volledige recursieve scan met `-Apply`

#### Netwerk en connectiviteit

- TCP-connectiviteit testen op elke host — losse poorten, bereiken (`1294:1494`), combinaties (`80,443,1294:1494`)
- Eenmalige SMTP-test met interactieve vraag om referenties
- Terugkerende SMTP-test (elke 5 minuten) met opgeslagen, versleuteld wachtwoord
- Diagnose van authenticatie en netwerk — Logboeken (mislukte aanmeldingen, Kerberos, NTLM, beschikbaarheid van DC's), tijdsynchronisatie, DNS, TCP, UNC-shares, optionele logscan; exporteert een txt-rapport naar `C:\Temp\`
- Diagnose van bestands-I/O — lus van schrijven/toevoegen/lezen/verwijderen op elk pad; deelt fouten in als AUTH/NETWORK/TIMEOUT/DISK/PATH; legt bij elke fout FileSystemWatcher-gebeurtenissen vast, een verschil van de NTFS-rechten ten opzichte van de baseline, open proceshandles (Handle.exe wordt automatisch gedownload van Sysinternals), een momentopname van nieuwe processen, Kerberos-tickets en het beveiligingslogboek; stopt na 3 fouten
- **UniFi** — HTML-rapport met netwerkdocumentatie (apparaten, firmware, uptime, per site) en tooling voor firmware-upgrades voor een UniFi Controller/UniFi OS-console; referenties via `Get-Credential`, nooit hardgecodeerd

#### Apparaat

- Diagnose van OpenVPN Connect — PnP-adapters, services, routes, DNS, Logboeken, conflicterende VPN-software; exporteert een txt-rapport naar `C:\Temp\`

#### RDS

- Diagnose van RDP + RD Web Access (`Test-RDSDiagnostics.ps1`) — achterhalen waarom gebruikers niet kunnen aanmelden op een RDP- of RDWeb-server:
  - Services (TermService, SessionEnv, UmRdpService), RDP in-/uitgeschakeld, NLA, sessielimieten, RD Licensing, firewallregels, actieve sessies
  - Geldigheid en vervaldatum van het HTTPS-certificaat op RDWeb; status van de IIS-app-pool en RD Gateway (alleen lokaal)
  - Controles van het gebruikersaccount: ingeschakeld, vergrendeld, wachtwoord verlopen, lidmaatschap van Remote Desktop Users
  - Analyse van logboeken: mislukte aanmeldingen (4625), vergrendelingen (4740), Kerberos-fouten (4771), redenen voor verbreken van sessies (20/40)
  - Logbestand met tijdstempel opgeslagen in `C:\Temp\`; `-IncludeEventLogs` voor analyse van gebeurtenissen

- Realtime RDS-monitor (`Watch-RDSLive.ps1`) — bevraagt de logboeken elke N seconden en streamt nieuwe gebeurtenissen naar de console + een logbestand:
  - Sessiegebeurtenissen: aanmelden (21), opnieuw verbinden (22/25), afmelden (23), verbreken (24), aanmelden mislukt (20), reden van verbreken (40) met leesbare redencodes
  - Beveiliging: mislukte RDP-aanmeldingen (4625 type 10), accountvergrendelingen (4740)
  - Licenties: gebeurtenissen van `TerminalServices-Licensing/Admin` + provider `TermServLicensing` in het systeemlogboek
  - Hartslagregel per bevraging met het aantal actieve sessies en het aantal nieuwe gebeurtenissen
  - Rechtstreeks uitvoeren op elke RDS-/RDWeb-server; `-IntervalSeconds` (standaard 20), `-NoLogFile` om geen bestand te schrijven

---

### 📊 Rapportage

#### Rapport laatste aanmelding van computers

De datum van de laatste aanmelding rapporteren voor alle computerobjecten in een of meer OU's en exporteren naar CSV.

- Bevraagt Active Directory naar computers in opgegeven OU's (bv. `OU=Laptops`, `OU=Computers`)
- Twee nauwkeurigheidsmodi: `LastLogonTimestamp` (snel, maximaal 14 dagen vertraging) of `-AllDCs` (bevraagt elke DC naar de exacte `LastLogon`)
- Vier statussen: **Active** · **Active (pwd recent)** · **Stale** · **Never** · **Disabled**
- `Active (pwd recent)`: apparaat dat ten onrechte als verouderd is gemarkeerd door de replicatievertraging van 14 dagen — een `PasswordLastSet` van minder dan 35 dagen bevestigt dat de machine online is (computeraccounts wisselen hun wachtwoord automatisch zo'n elke 30 dagen)
- CSV-kolommen: Name, Status, Enabled, LastLogon, DaysSinceLogon, PasswordLastSet, DaysSincePasswordSet, OS, IPv4, OU-pad, Created, Description
- Ondersteunt meerdere OU's in één run; `-IncludeDisabled` om uitgeschakelde objecten mee te nemen

#### SharePoint-rechtenrapport

**Get-SharePointPermissionsReport.ps1** — wie bij welke SharePoint kan, via welke groep, op welk niveau. Alleen-lezen: elke aanroep die het doet is een GET.

- Begint met één geconsolideerd overzicht — één rij per persoon per site, met de groep waarlangs de toegang loopt en het niveau dat die geeft. Toekenningen en lidmaatschap staan anders in aparte rapporten, en "Site Owners heeft Full Control" plus "Site Owners bevat vijf mensen" is nog geen antwoord
- Daaronder: sitecollectiebeheerders, roltoewijzingen op web/lijst/item, onderbroken overerving, SharePoint-groepen met hun leden, Entra-groepen herleid tot transitief lidmaatschap, deellinks met hun soort, externe en gastprincipals, toekenningen aan `Everyone`
- Een item wordt alleen als eigen bereik gerapporteerd als het unieke rechten heeft, zodat het rapport de rechtenstructuur in kaart brengt in plaats van per bestand een rij te herhalen
- Roltoewijzingen zijn niet leesbaar via Graph en vallen niet onder de rollen Read/Write/Manage van SharePoint, dus het maakt een kortlevende app met certificaat en `Sites.FullControl.All` aan en verwijdert die weer. Een clientgeheim werkt niet — SharePoint Online weigert app-only-tokens op basis van een geheim
- `-Excel` schrijft één werkmap met een werkblad per rapport plus kant-en-klare draaitabellen; de CSV's worden altijd geschreven en de werkmap wordt daaruit opgebouwd
- Hervat na een onderbreking vanaf de laatst voltooide lijst, en meldt aan het eind of elk bereik ook echt gelezen kon worden

#### SharePoint-opslagrapport

**Get-SharePointStorageReport.ps1** — tenantbrede opslag per site, bibliotheek, versiegeschiedenis en prullenbak, met totalen per sitecollectie die vergelijkbaar zijn met het beheercentrum.

#### SharePoint-versies opschonen

**Remove-SharePointFileVersionsByDate.ps1** — bestandsversies rapporteren (en met `-Apply` verwijderen) die ouder zijn dan een grensdatum. De huidige versie blijft altijd behouden.

#### Licentierapport

Generator voor het maandelijkse licentie- en Azure-kostenrapport.

- Combineert factuurgegevens van Pax8 (CSV) en Ingram (Excel) tot één opgemaakte Excel
- Tabblad per klant met Azure-verbruik, licenties en een uitsplitsing van Acronis
- Overzichtstabblad met totalen en marge per klant
- PowerShell-launcher met controle vooraf
- Optionele geplande taak in Windows (draait op de 6e van elke maand)

---

### 🖥️ Infrastructuur en apparaten

#### USB-installatietoolkit

USB-toolkit voor Windows-installatie en Autopilot-inschrijving tijdens OOBE.

- Interactief menu (Apparaatbeheer, Autopilot, AD-join, apparaat hernoemen, productcode, Windows Update, herstarten)
- Browser voor klantinstallaties vanuit het menu van de USB-toolkit:
  - Lokale map `Install` per klant (`D`)
  - Netwerkshare `\\10.222.3.94\Software` per klant (`E`)
- Kopieer voor de lokale optie `D` zowel `Browse-InstallScripts.ps1` als de volledige map `Install` naast `start.bat`
- Vóór de opties `D` en `E` maakt of werkt de toolkit de lokale beheerder `LocalAdmin` bij met wachtwoord `<wachtwoord weggelaten>`, voegt die toe aan `Administrators`, en zet de OOBE-overslaanvlaggen
- Verhoogt zichzelf, OOBE-compatibel via Shift+F10
- Gesplitste "Alles in één": `A` = Intune (Hernoemen + Autopilot + Update), `C` = AD (Hernoemen + Domein-join + Update)

#### Audiobeheer

Drie scripts die samenwerken om interne microfoons op endpoints te detecteren, uit te schakelen en weer terug te draaien — uitgerold via NinjaOne.

| Script | Doel |
|---|---|
| [`detect-audiodevices.ps1`](scripts%5CDevice%5Caudio%5Cdetect-audiodevices.ps1) | Inventaris van alle audio-endpoints op het toestel |
| [`Disable-internalmic.ps1`](scripts%5CDevice%5Caudio%5CDisable-internalmic.ps1) | Interne microfoon(s) uitschakelen, headsets worden overgeslagen |
| [`Rollback-InternalMic.ps1`](scripts%5CDevice%5Caudio%5CRollback-InternalMic.ps1) | Eerder uitgeschakelde interne microfoons weer activeren |

**NinjaOne-uitrol (alle drie de scripts):**

| Instelling | Waarde |
|---|---|
| Run as | **SYSTEM** |
| Scriptparameters | _(geen)_ |
| Vereist custom field | `AudioDeviceInventory` (device, tekstveld/textarea) |
| Exitcode | `0` = succes · `1` = fout (script gemarkeerd als mislukt) |

> Het custom field `AudioDeviceInventory` moet als custom field op apparaatniveau in NinjaOne zijn aangemaakt voordat je de scripts uitrolt. De uitvoer van elk script wordt daarin weggeschreven via `Ninja-Property-Set AudioDeviceInventory`.

#### Tijdsynchronisatie

- De Windows Time-service herstarten en synchronisatie afdwingen

#### Windows opschonen

Grondige opschoning van schijfruimte voor Windows-endpoints.

- Schoont tijdelijke mappen van gebruiker/systeem op, de downloadcache van Windows Update, de cache van Delivery Optimization, Prefetch, geheugendumps, WER-wachtrijen, de miniatuurcache, de DirectX-shadercache, de Prullenbak, browsercaches (Edge, Chrome, Firefox) en logboeken
- Opschoning van het DISM-componentarchief (`/StartComponentCleanup /ResetBase`) na Windows-updates
- DNS-cache leegmaken
- Standaard een proefdraai — toont de terug te winnen ruimte per categorie zonder iets te verwijderen
- Draai met `-Apply` om echt op te schonen; afzonderlijke categorieën kun je overslaan met `-SkipBrowserCache`, `-SkipEventLogs`, `-SkipDism`, `-SkipRecycleBin`
- Exporteert een CSV-rapport met vrijgemaakte bytes per categorie naar `C:\Temp\`

#### OEM-bloatware verwijderen

- Detecteert de fabrikant van het apparaat (HP/Lenovo/Dell) en verwijdert bekende OEM-bloatware via `winget`, plus een generieke lijst van consumentenapps uit de Microsoft Store (Xbox, Solitaire, Bing News/Weather, Cortana, Clipchamp)
- Standaard een proefdraai; `-Apply` om echt te verwijderen. CSV-rapport van gevonden/verwijderde apps naar `C:\Temp\`

#### Clouddrives koppelen

- Koppelt documentbibliotheken van SharePoint/OneDrive via WebDAV (`net use`) aan vaste stationsletters, voor gebruik als aanmeldscript per gebruiker (Intune Win32-app of geplande taak)
- Gestuurd door een mappings-CSV (`DriveLetter`, `Url`, optioneel `Label`); standaard een proefdraai, `-Apply` om echt te koppelen
- Geen opgeslagen referenties — steunt op de bestaande tenantsessie van de aangemelde gebruiker (net als WebDAV-toegang via de browser)

#### Tijdelijke schijf en pagefile (Azure / AVD)

Twee scripts die de vluchtige tijdelijke schijf (`D:`) van een Azure-VM of AVD-sessiehost op zijn plaats houden, en de pagefile erop.

| Script | Doel |
|---|---|
| [`Init-TempDisk.ps1`](scripts/Device/TempDisk/Init-TempDisk.ps1) | De tijdelijke schijf herstellen als `D:` en de pagefile erop configureren |
| [`Register-InitTempDiskTask.ps1`](scripts/Device/TempDisk/Register-InitTempDiskTask.ps1) | Dat script op het apparaat installeren en het bij elke opstart als SYSTEM uitvoeren |

- De tijdelijke schijf wordt gewist bij elke deallocate, resize of verhuizing naar een andere host — en Windows leest de pagefileconfiguratie bij het opstarten, dus een pagefile op een stationsletter die er bij het opstarten niet is, wordt nooit aangemaakt en de machine pagineert weer op `C:`
- Herstelt het volume (alleen RAW-schijven — een schijf die nog partities heeft, wordt gemeld en nooit geformatteerd), verplaatst een optisch station weg van `D:` als het in de weg zit, laat de pagefile dan naar `D:\pagefile.sys` wijzen en verwijdert de vermelding voor elk ander station
- Windows leest die configuratie alleen bij het opstarten, dus `-RestartIfNeeded` (wat de opstarttaak gebruikt) herstart de machine één keer als dat het enige is wat nog rest - nooit na een mislukte run, nooit terwijl er iemand is aangemeld, en hooguit één keer per uur. Het aftellen geldt alleen als er iemand is aangemeld om het te zien; bij het opstarten herstart het binnen enkele seconden
- `-CheckOnly` rapporteert zonder iets te wijzigen (exitcode `2` = er is werk te doen); `-WhatIf` doorloopt de hele flow; `-Quiet` houdt een gezonde opstart stil

---

### 🔧 Eigen tools

#### SAS-batchmonitoring

Logs van SAS-batchjobs en Windows Logboeken bewaken op fouten, met optionele Zabbix-integratie en e-mailwaarschuwingen.

- Detecteert spawn-fouten, authenticatiefouten op de WORK-bibliotheek, afbrekingen, schijffouten en algemene `ERROR:`-regels
- Uitvoerformaten tekst, JSON en Zabbix; instelbare terugkijkperiode
- Eenmalig installatiescript — installeert naar `C:\Scripts\`, maakt een dagelijkse geplande taak aan
- Optionele Zabbix-UserParameter-config voor geautomatiseerde waarschuwingen

#### Beheer van Windows-apparaten

Scripts voor het beheren en onderhouden van Windows-apparaten.

**Invoke-WindowsActivation.ps1** — Windows activeren of licentie-instellingen beheren:
- Een retail- of generieke KMS-productcode installeren (`-ProductKey`)
- Een KMS-activeringsserver van het bedrijf configureren (`-KmsServer`, `-KmsPort`)
- Online of KMS-activering starten (`-Activate`)
- Activeringsstatus tonen via WMI en `slmgr /dli` (`-Status`)
- De productcode verwijderen vóór een reimage of licentieoverdracht (`-RemoveKey`)
- De teller van de respijtperiode resetten (`-ReArm`, max. ~3-5x per installatie)
- Standaard met bevestigingsvragen; gebruik `-Force` om ze over te slaan

**NinjaOne-uitrol:**

| Instelling | Waarde |
|---|---|
| Run as | **Administrator** |
| Custom fields | _(geen — uitvoer via console/scriptlog)_ |
| Exitcode | `0` = succes · `1` = fout |

> **Let op:** `-RemoveKey` en `-ReArm` vragen interactieve bevestiging. Voeg altijd `-Force` toe wanneer je deze via NinjaOne uitvoert, anders blijft het script hangen.

Veelgebruikte NinjaOne-scriptparameters:

| Scenario | Parameters |
|---|---|
| Status controleren | `-Status` |
| KMS-activatie | `-KmsServer kms.bedrijf.local -Activate -Status` |
| KMS met afwijkende poort | `-KmsServer kms.bedrijf.local -KmsPort 2500 -Activate` |
| Retail key installeren + activeren | `-ProductKey XXXXX-XXXXX-XXXXX-XXXXX-XXXXX -Activate -Status` |
| Key verwijderen (voor reimage) | `-RemoveKey -Force` |
| Grace period resetten | `-ReArm -Force` |

**Invoke-WindowsCleanup.ps1** — terug te winnen schijfruimte scannen en eventueel vrijmaken:
- Tijdelijke bestanden van gebruiker + systeem, Windows Update-cache, Delivery Optimization, Prefetch
- Geheugendumps, WER-wachtrijen, miniatuur-/DirectX-shadercache, lettertypecache
- Prullenbak, browsercaches (Edge/Chrome met meerdere profielen + Firefox)
- Logboeken, DISM-componentarchief (`/StartComponentCleanup /ResetBase`)
- Applicatie- en systeemlogs: dynamische scan van heel C:\ op mappen `logs`/`log`/`logging`
- Standaard een proefdraai; gebruik `-Apply` om te verwijderen. Samenvatting per categorie met de vrijgemaakte ruimte

**Repair-AppxPackageStore.ps1** — AppX-pakketten repareren (Teams, nieuwe Outlook, elk ander pakket) die falen met `0x80070490` / "Deployment Register operation ... from:  (AppxManifest.xml)":
- Stelt registraties vast waarvan de bestanden weg zijn, geprovisioneerde kopieën zonder bestanden, en verweesde vermeldingen in `AppxAllUserStore` (geen profiel, geen bestanden, geen manifest)
- Leest op FSLogix-hosts de fouten van `Microsoft-FSLogix-Apps`: welke exacte versie de profielen vragen tegenover wat deze host provisioneert, de FSLogix-build, `InstallAppxPackages`, ODFC `IncludeTeams`, en het AppX-installatiebeleid
- Somt **elke** app op die de afgelopen `-Days` dagen niet kon installeren, bijwerken of registreren (AppX-implementatielog + FSLogix-log), met de betekenis van elke foutcode
- Repareert in een vaste volgorde — deprovisioneren, opnieuw registreren, verwijderen, en dan elke registersleutel naar `.reg` back-uppen voordat hij wordt verwijderd — en leest alles terug; `-Provision` zet Teams / nieuwe Outlook voor alle gebruikers terug met het eigen installatieprogramma van Microsoft, of met `-UseWinget` via winget; `-WingetId` doet hetzelfde voor elke andere app
- `-CheckOnly` wijzigt niets; met `-Name '*'` worden systeem-/frameworkpakketten en Deprovisioned-markeringen nooit aangeraakt

#### DNS-beheer

Scripts voor het beheren van DNS-records in Active Directory-geïntegreerde DNS-zones.

- Openbare DNS-records opzoeken via Google DNS (dig) en ze als A- of CNAME-records importeren in AD DNS
- Standaard een proefdraai — toont wat er zou worden aangemaakt voordat het wordt toegepast
- Idempotent — slaat records over die al bestaan

---

### ☁️ Azure-infrastructuur

Scripts die zich via de module `Az` rechtstreeks op Azure IaaS richten — niet op de M365-tenant, en niet opgenomen in `menu.ps1`.

- **Azure-NVMe-Conversion.ps1** — meegeleverd script van derden (Microsoft, MIT-licentie, uit `Azure/SAP-on-Azure-Scripts-and-Utilities`) dat het type schijfcontroller van een VM omzet tussen SCSI en NVMe, inclusief controles en correcties van de driverbereidheid in het gastbesturingssysteem voor zowel Windows- als Linux-gasten
- **Search-AADDSUserActivity.ps1** — doorzoekt alle audittabellen van Azure AD Domain Services in Log Analytics voor één gebruiker in één `union`-query, in plaats van te gokken in welke tabel een gebeurtenis terechtkwam

---

### 🗄️ Herschreven legacy-toolkits

Een inmiddels uitgefaseerde interne PowerShell-repo (en twee geforkte GitHub-toolkits van derden die erin waren opgenomen) is script voor script doorgelicht en gemoderniseerd naar de huisstijl van deze repo — Graph/Exchange Online in plaats van de uitgefaseerde modules `MSOnline`/`AzureAD`, standaard een proefdraai met `-Apply` voor alles wat iets wijzigt, geen hardgecodeerde klantgegevens of geheimen. Geen van deze scripts is opgenomen in `menu.ps1` — het zijn audit-, rapportage- en inrichtingsscripts die bedoeld zijn om rechtstreeks uit te voeren, volgens hetzelfde patroon als `scripts/RDS/`, `scripts/Azure/` en `scripts/Network/UniFi/`. Elke map heeft een eigen readme met volledige documentatie van parameters en gebruik.

| Map | Bron | Omvat |
|--------|--------|--------|
| [`TenantOnboarding/`](scripts/TenantOnboarding/readme.nl.md) | Interne toolkit voor tenantinrichting | Inrichten van nieuwe tenants (break-glass-beheerder, baselinegroepen/Intune-toewijzing), multi-tenant-/GDAP-rapportage van licenties + break-glass-wachtwoorden, uitrol van Win32-/Chocolatey-apps, apparaatconfiguratie (energiebeheer kiosk, Office verwijderen, indeling Startmenu), OneDrive-beheer, gebruikersbeheer via dynamische DG's/featuregroepen |
| [`Office365Toolkit/`](scripts/Office365Toolkit/readme.nl.md) | Fork van [`directorcia/Office365`](https://github.com/directorcia/Office365) (CIAOPS) | Secure Score-rapportage, opschonen van toestemmingen voor enterprise-apps, aanmeldblokkade voor gedeelde mailboxen, EOP-baseline, audits van mailboxhygiëne/doorstuurrisico, zoeken in het Unified Audit Log, inventaris van Intune-beleid |
| [`PatronToolkit/`](scripts/PatronToolkit/readme.nl.md) | Fork van [`directorcia/patron`](https://github.com/directorcia/patron) | MFA-registratie + export van CA-beleid, audits van toestemmingen voor enterprise-apps + verdachte inboxregels + gebundelde beveiligingswaarschuwingen, geconsolideerde e-mailbeveiligingsstatus + controles van mailboxauditing, SPF-/DMARC-validatie, rapporten van Intune-beleidstoewijzingen + Autopilot-apparaten, message trace, deelconfiguratie van SharePoint, Teams-configuratierapport |
| [`LegacyUtilities/`](scripts/LegacyUtilities/readme.nl.md) | Interne toolkit (diverse kleine scripts) | Rechten op mailboxmappen/gedelegeerde toegang, gedeelde mailboxen/contactpersonen in bulk aanmaken, contactsynchronisatie, opruimen van dubbele mailitems, lidmaatschap van M365-groepen, back-up van CA-beleid, klonen van Teams/Planner, stationskoppeling voor Azure Files, apparaataanpassingen voor NumLock/werkstation vergrendelen, inrichten van Workspace 365-omgevingen |

Beide GitHub-forks zijn functie voor functie doorgelicht in plaats van 1-op-1 overgezet — bijna identieke rapportagescripts met één doel zijn samengevoegd tot minder, goed geparametriseerde scripts, en functies die elders in deze repo al gedekt waren, zijn overgeslagen in plaats van gedupliceerd (zie de readme van elke map voor de volledige lijst van overgeslagen onderdelen en de motivering). Alle code is een nieuwe implementatie in de stijl van deze repo, niet gekopieerd uit de bronprojecten.

---

## Repositorystructuur

Elke map heeft een eigen `readme.md` — deze boom is een plattegrond; volg de links voor de volledige documentatie van parameters en gebruik.

```
M365-Scripts/
├── .gitignore
├── .vscode/
│   └── settings.json
├── load.ps1                         ← Startpunt: inrichting bij eerste start + start het menu
├── menu.ps1                         ← Interactieve launcher (alle scripts + M365-functies)
├── readme.md
└── scripts/
    ├── readme.md                    ← Index van alle categorieën hieronder
    ├── INDEX.md                     ← Elk script van A tot Z met zijn map (gegenereerd)
    ├── Azure/                        ← richt zich rechtstreeks op Azure IaaS via Az, niet op de M365-tenant
    │   ├── readme.md
    │   └── VM/
    │       ├── readme.md
    │       └── Azure-NVMe-Conversion.ps1   ← meegeleverd (Microsoft, MIT) — conversie van de schijfcontroller SCSI/NVMe
    ├── Entra/
    │   ├── readme.md
    │   ├── Set-UserManager.ps1
    │   ├── Remove-M365Users.ps1
    │   ├── New-M365User.ps1
    │   ├── Import-M365Users.ps1
    │   ├── Get-M365UserLicenses.ps1
    │   ├── Import-ConditionalAccessBaseline.ps1
    │   └── Test-M365GroupMembership.ps1   ← eigenaren en leden van M365-groepen / Teams auditen
    ├── Exchange/
    │   ├── readme.md
    │   ├── Migrate-Calendar.ps1
    │   ├── Convert-SharedCalendarToResource.ps1  ← gedeelde agenda in de mailbox van een gebruiker → eigen ruimte-/apparatuurmailbox
    │   ├── Move-SharedCalendar.ps1  ← alles in één: zoeken op trefwoord + omzetten + wie moet overstappen
    │   ├── Set-Calendar-rights.ps1
    │   ├── Set-Distributionlist-dynamic-static.ps1
    │   ├── Move-InboxToArchive.ps1
    │   ├── Restore-MailboxMessages.ps1  ← mail terugzetten die op een datum is verplaatst/verwijderd, en wie het deed
    │   ├── Test-CalendarPermissions.ps1
    │   ├── Get-CalendarMappings.ps1  ← waar elke agenda in Outlook is toegevoegd, naast de rechten erachter
    │   ├── Test-MailboxPermissions.ps1
    │   ├── Test-DistributionGroupPermissions.ps1
    │   ├── Test-DkimConfig.ps1
    │   ├── Get-ExternalForwards.ps1
    │   ├── Get-MailboxSizes.ps1
    │   └── Get-DistributionGroupMembers.ps1  ← wie in welke distributielijst zit, als Excel-werkmap voor de klant
    ├── Graph/
    │   ├── readme.md
    │   └── logic-permissies.ps1     ← een Graph-approl toekennen aan de managed identity van een Logic App
    ├── Intune/
    │   ├── readme.md
    │   ├── Compare-IntuneConfig.ps1  ← drift van de Intune-config t.o.v. een MSP-baselineback-up (IntuneBackupAndRestore)
    │   ├── Get-Autopilot/
    │   │   ├── readme.md
    │   │   ├── Get-WindowsAutoPilotInfo.ps1
    │   │   └── GetAutoPilot.CMD
    │   ├── iOS-Compliance-Updater/
    │   │   ├── readme.md
    │   │   ├── Update-iOSCompliancePolicy.ps1   ← hoofdscript (handmatig of als geplande taak)
    │   │   ├── Setup.ps1                        ← eenmalig: App Registration + config.json
    │   │   ├── Install-ScheduledTask.ps1        ← wekelijkse geplande taak registreren
    │   │   └── config.example.json
    │   └── Desktop/                  ← bedrijfsachtergrond/-vergrendelscherm (het Office-thema staat in Custom Scripts/, zie hieronder)
    │       ├── readme.md
    │       ├── Add Lockscreen to start and desktop/
    │       │   ├── readme.md
    │       │   ├── add-lock.ps1               ← snelkoppeling "Lock Workstation" op de taakbalk
    │       │   └── add-shortcut-lock.ps1
    │       └── Background/
    │           ├── readme.md
    │           ├── Desktop/
    │           │   ├── readme.md
    │           │   ├── Set-CorporateWallpaper.ps1  ← bedrijfsachtergrond via Intune (hashcontrole, PersonalizationCSP)
    │           │   └── Remove-CorporateWallpaper.ps1
    │           └── Lockscreen/
    │               ├── readme.md
    │               └── Make-lockscreen.ps1         ← bedrijfsvergrendelscherm via Intune (gevalideerde download, PersonalizationCSP)
    ├── Device/
    │   ├── readme.md
    │   ├── Invoke-WindowsActivation.ps1 ← Windows activeren, productcode / KMS-server instellen
    │   ├── Invoke-WindowsCleanup.ps1    ← temp, cache, WU, DISM, browser, logboeken
    │   ├── Clear-TempFiles.ps1
    │   ├── Remove-OemBloatware.ps1      ← HP/Lenovo/Dell- + generieke Store-bloatware verwijderen
    │   ├── Repair-AppxPackageStore.ps1  ← AppX 0x80070490 repareren (verweesde store-vermeldingen, FSLogix-replay)
    │   ├── Test-OpenVpnDiagnostics.ps1  ← diagnose van OpenVPN Connect
    │   ├── Update-TeamsClient.ps1       ← nieuwe Teams + vergaderinvoegtoepassing bijwerken als er een nieuwere build is
    │   ├── Update-TeamsClient.md        ← hoe dat script stap voor stap beslist
    │   ├── Update-TeamsClient-ITGlue.md ← servicedeskversie (NL) om in IT Glue te plakken
    │   ├── audio/
    │   │   ├── readme.md
    │   │   ├── detect-audiodevices.ps1
    │   │   ├── Disable-internalmic.ps1
    │   │   └── Rollback-InternalMic.ps1
    │   ├── DriveMapping/
    │   │   ├── readme.md
    │   │   └── New-CloudDriveMapping.ps1   ← SharePoint-/OneDrive-bibliotheken aan stationsletters koppelen (WebDAV)
    │   ├── TempDisk/
    │   │   ├── readme.md
    │   │   ├── Init-TempDisk.ps1              ← de vluchtige tijdelijke schijf herstellen als D: en de pagefile erop zetten
    │   │   └── Register-InitTempDiskTask.ps1  ← dat script installeren en bij elke opstart als SYSTEM uitvoeren
    │   └── Time sync/
    │       ├── readme.md
    │       └── Restart-Time-Sync.ps1
    ├── Network/
    │   ├── readme.md
    │   ├── Test-Ports.ps1
    │   ├── Test-AuthNetworkDiagnostics.ps1   ← diagnose van authenticatie-/netwerkproblemen
    │   ├── Test-FileIODiagnostics.ps1        ← bestands-I/O-test + realtime mapmonitor
    │   └── UniFi/
    │       ├── readme.md
    │       ├── UnifiApi.ps1                  ← gedeelde helper voor aanmelding/sessie (gedot-sourcet)
    │       ├── Get-UnifiNetworkReport.ps1     ← HTML-rapport met netwerkdocumentatie
    │       └── Update-UnifiFirmware.ps1       ← firmware-upgrades over sites heen oplijsten/starten
    ├── RDS/
    │   ├── readme.md
    │   ├── Test-RDSDiagnostics.ps1           ← diagnose van mislukte RDP-/RDWeb-aanmeldingen
    │   └── Watch-RDSLive.ps1                 ← realtime monitor van sessies + licenties
    ├── SMTP/
    │   ├── readme.md
    │   ├── testsmtp.ps1
    │   └── testsmtp_5min.ps1
    ├── Deployment/                   ← USB-installatietoolkit (OOBE / Autopilot)
    │   ├── readme.md
    │   ├── start.bat
    │   ├── autorun.inf
    │   └── Browse-InstallScripts.ps1
    ├── DNS/
    │   ├── readme.md
    │   ├── Import-DnsRecords.ps1   ← opzoeken via Google DNS + importeren in AD DNS
    │   └── example-records.csv
    ├── SAS/
    │   ├── readme.md
    │   ├── rca.md
    │   ├── Monitor-SASBatchErrors.ps1   ← logs + Logboeken scannen op SAS-fouten
    │   ├── Setup-SASMonitoring.ps1      ← installatiescript, geplande taak, Zabbix-config
    │   ├── Test-SASWorkDirectory.ps1    ← gezondheid van de WORK-map valideren
    │   └── zabbix_sas_monitor.conf
    ├── SharePoint/
    │   ├── readme.md
    │   ├── Find-SiteContent.ps1         ← een hele site doorzoeken (naam/pad/type/datum of volledige tekst) + de rechten op elke treffer rapporteren (PnP)
    │   ├── Search-SharePointContent.ps1 ← hetzelfde, tenantbreed via Graph app-only: delta + /permissions, deellinks en gasten (bestanden/mappen)
    │   ├── Restore-RecycleBinItems.ps1  ← verwijderde bestanden terugzetten uit een prullenbak: één site/OneDrive of tenantbreed (PnP, automatische app-registratie)
    │   ├── Revoke-SharePointUserAccess.ps1 ← de toegang van één gebruiker op elk niveau intrekken, deellinks inbegrepen (rapporteert tenzij -Apply)
    │   ├── Test-SharePointAccessScripts.ps1 ← de twee toegangsscripts verifiëren zonder tenant (gedeeld auth-blok + intrekkingstrechter)
    │   └── Provisioning/                ← een complete structuur inrichten vanuit één JSON-config (PnP + Graph)
    │       ├── readme.md
    │       ├── Petsolutions-SharePoint-Handleiding.md ← eindgebruikershandleiding (NL) om aan de klant te geven
    │       ├── petsolutions.config.json     ← het model: kolommen, inhoudstypen, groepen, bibliotheken, weergaven, rechten
    │       ├── Install-SharePointStructure.ps1 ← alles in één run opbouwen, incl. (tijdelijke) app-registratie
    │       ├── SharePointStructure.Common.ps1 ← gedeelde helpers (gedot-sourcet door alle vier)
    │       ├── New-SharePointMetadata.ps1   ← termenset, sitekolommen, inhoudstypen (elke site in de config)
    │       ├── Set-SharePointLibraries.ps1  ← bibliotheken/kanaalmappen, inhoudstypen, standaardwaarden, weergaven, groepsrechten
    │       ├── Update-SharePointShareStatus.ps1 ← de kolom Deelstatus afleiden, overmatig delen markeren (exit 2)
    │       └── Test-SharePointStructure.ps1 ← alleen-lezen driftcontrole t.o.v. de config (exit 2)
    ├── Teams/
    │   ├── readme.md
    │   └── vias_archiver.ps1        ← export + archivering van Teams/SharePoint (Graph, PS7+, Global Admin)
    ├── Reporting/
    │   ├── readme.md
    │   ├── Get-ComputerLastLogon.ps1        ← laatste aanmelding per computer in OU('s), export naar CSV
    │   ├── Get-SharePointStorageReport.ps1  ← tenantbreed SharePoint-opslagrapport
    │   ├── Get-SharePointPermissionsReport.ps1 ← wie waar toegang toe heeft en via welke groep, naar CSV + Excel
    │   ├── Remove-SharePointFileVersionsByDate.ps1 ← bestandsversies ouder dan een datum verwijderen
    │   └── Licensing/
    │       ├── readme.md
    │       ├── genereer_licentie_overzicht.py
    │       ├── genereer_rapport.ps1
    │       ├── genereer_rapport.bat
    │       └── create_scheduled_task.ps1
    ├── Startup/
    │   ├── readme.md
    │   ├── functies.ps1             ← M365-functiebibliotheek (gedot-sourcet door het menu)
    │   ├── Install-Modules.ps1      ← Bootstrap: alle modules installeren en importeren
    │   ├── Update-Modules.ps1       ← Elke geïnstalleerde PowerShell-module bijwerken
    │   ├── Test-PowerShellSyntax.ps1
    │   ├── Update-ScriptIndex.ps1   ← Genereert scripts/INDEX.md opnieuw uit de .SYNOPSIS-headers
    │   ├── Test-MarkdownLinks.ps1   ← Controleert elke readme-link: bestanden en ankers binnen de pagina
    │   └── Convert-MarkdownToHtml.ps1 ← Markdown-document → één opgemaakte HTML-pagina zonder externe afhankelijkheden
    ├── Custom Scripts/                 ← scripts die aan hun pad vastzitten (zie opmerking hierboven)
    │   ├── readme.md
    │   └── Intune/
    │       ├── readme.md
    │       └── Desktop/
    │           ├── readme.md
    │           ├── Deploy-OfficeTheme.ps1        ← installeert het volledige VIAS-Office-thema (.thmx)
    │           ├── 2026 Vias institute colours (2).thmx
    │           └── Office Themes/
    │               ├── readme.md
    │               ├── Deploy-Officecolors.ps1   ← installeert alleen het kleurenschema
    │               └── Test VIAS.xml
    ├── TenantOnboarding/                ← gemoderniseerd vanuit een uitgefaseerde interne toolkit voor tenantinrichting, niet in het menu
    │   ├── readme.md
    │   ├── Provisioning/         (3 scripts)  ← break-glass-beheerder, baselinegroepen, toewijzing van Intune-beleid
    │   ├── MultiTenant/          (3 scripts)  ← GDAP-licentierapport, rotatie van break-glass-wachtwoorden, index van het klantportaal
    │   ├── AppDeployment/        (7 scripts)  ← Win32-/Chocolatey-installatie, snelkoppelingen, bestandskoppelingen, printerverbindingen
    │   ├── DeviceConfig/         (6 scripts)  ← energiebeheer kiosk, Office verwijderen, indeling Startmenu, firewallregel voor Teams
    │   ├── OneDriveManagement/   (3 scripts)  ← synchronisatiewaakhond, synchronisatie van bibliotheken stoppen, omleiden van bekende mappen
    │   └── UserManagement/       (2 scripts)  ← dynamische DG via filter, lidmaatschap van featuregroepen
    ├── Office365Toolkit/                ← herschrijving van de uitgefaseerde fork van directorcia/Office365 (CIAOPS), niet in het menu
    │   ├── readme.md
    │   ├── Security/             (4 scripts)  ← Secure Score, opschonen van app-toestemmingen, blokkade gedeelde mailboxen, EOP-baseline
    │   ├── Exchange/             (4 scripts)  ← baseline mailboxhygiëne, doorstuurrisico, invoegtoepassingen, zoeken in auditlog
    │   └── Intune/               (1 script)   ← tenantbrede inventaris van beleid
    ├── PatronToolkit/                    ← herschrijving van de uitgefaseerde fork van directorcia/patron, niet in het menu
    │   ├── readme.md
    │   ├── Entra/                (2 scripts)  ← rapport MFA-registratie, export van CA-beleid
    │   ├── Security/             (5 scripts)  ← app-toestemmingen, verdachte inboxregels, beveiligingswaarschuwingen, e-mailbeveiligingsstatus, mailboxauditing
    │   ├── Exchange/             (1 script)   ← message trace-rapport
    │   ├── Intune/               (2 scripts)  ← beleidstoewijzingen, Autopilot-apparaten
    │   ├── SharePoint/           (1 script)   ← audit van de deelconfiguratie
    │   └── Teams/                (1 script)   ← Teams-configuratierapport
    └── LegacyUtilities/                  ← diverse gemoderniseerde scripts uit de uitgefaseerde interne toolkit, niet in het menu
        ├── readme.md
        ├── Exchange/             (7 scripts)  ← maprechten, gedelegeerde toegang, mailboxen/contactpersonen in bulk, contactsynchronisatie, ontdubbelen, message trace
        ├── Entra/                (2 scripts)  ← groepslidmaatschap, back-up van CA-beleid
        ├── Teams/                (3 scripts)  ← klonen van teams/plannen, inrichten van projectteams
        ├── Network/              (1 script)   ← stationskoppeling voor Azure Files
        ├── Device/               (2 scripts)  ← standaardinstelling NumLock, snelkoppeling werkstation vergrendelen
        └── Workspace365/         (2 scripts)  ← omgevingen inrichten/verwijderen
```

`Deploy-OfficeTheme.ps1` en `Deploy-Officecolors.ps1` hebben hun download-URL hard vastgezet op exact dit pad in de repo (branch `main`) — ze blijven hier staan in plaats van onder `Intune/Desktop/`, zodat de URL blijft werken.

---

## Bijdragen

Bij het toevoegen van nieuwe scripts:

1. Volg de bestaande naamgevingsconventie (`Verb-Noun.ps1`)
2. Neem een commentaarblok als header op met Synopsis, Description, Parameters en Example
3. Test tegen een niet-productietenant voordat je commit
4. Zet het script in de juiste workloadmap
5. Voeg het toe aan `menu.ps1` en werk deze readme bij
6. Werk de `Versiegeschiedenis` in dit bestand bij voor elke functionele of structurele wijziging (verplicht), ook voor wijzigingen die via Copilot/een AI-assistent zijn gevraagd of doorgevoerd
7. Schrijf de wijziging in alle drie de readmetalen (`readme.md`, `readme.nl.md`, `readme.fr.md`)

**Wat zichzelf bijwerkt.** `scripts/INDEX.md`, de taalwissel en het kruimelpad bovenaan elke readme, en de linkcontrole blijven automatisch actueel — die draai je niet met de hand:

| Wanneer | Wat er draait |
|---------|---------------|
| Je commit | De git-hook [`.githooks/pre-commit`](.githooks/pre-commit) genereert de index en de readme-headers opnieuw, neemt ze op in de commit, en stopt de commit bij een kapotte link. Hij waarschuwt als een Engelse readme is gewijzigd zonder de Nederlandse/Franse versie — vraag Claude dan om te vertalen |
| Claude Code past een bestand aan | Een hook in [`.claude/settings.json`](.claude/settings.json) doet na elke wijziging hetzelfde op de achtergrond, en voordat Claude klaar is controleert hij of gewijzigde readmes vertaald zijn en gewijzigde scripts gedocumenteerd |

Zet de git-hook eenmalig aan per clone:

```powershell
git config core.hooksPath .githooks
```

---

## Disclaimer

Deze scripts worden geleverd zoals ze zijn. Test altijd in een niet-productieomgeving voordat je ze op live tenants uitvoert. De beheerder aanvaardt geen aansprakelijkheid voor onbedoelde wijzigingen als gevolg van verkeerd gebruik of verkeerde configuratie.

---

## Versiegeschiedenis

> Opmerking: oudere vermeldingen kunnen verwijzen naar historische mapnamen zoals `Custom Scripts/` en `Testing Scripts/`. Die padnamen geven de structuur van de repository weer op het moment van die wijziging.

### 2026-09-30 (5)
| Wijziging |
|--------|
| De documentatie houdt zichzelf nu actueel. `.claude/hooks/sync-docs.ps1` genereert `scripts/INDEX.md` en de readme-headers opnieuw en draait de linkcontrole; hij wordt aangeroepen door een git pre-commit hook (`.githooks/pre-commit`, voor wijzigingen met de hand) en door Claude Code-hooks in `.claude/settings.json` (na elke wijziging, op de achtergrond). Voorheen moest je aan alle drie denken — en `INDEX.md` liep al twee scripts achter |
| Voordat Claude klaar is, controleert een Stop-hook of een wijziging aan een Engelse readme ook in het Nederlands en Frans is gedaan, en of bij een gewijzigd script de readme van de map is aangepast; de git-hook waarschuwt voor het eerste, omdat een hook niet kan vertalen. Een kapotte link stopt de commit |
| Geverifieerd: elke modus gedraaid op deze repository — een schone wijziging blijft stil, een ingevoegde kapotte link geeft exitcode 2 met bestand en doel, een onvertaalde readme en een ongedocumenteerd script blokkeren Stop elk één keer (en geen tweede keer), en de Claude-hook is na een wijziging zien afgaan. De pre-commit hook draaide op deze commit |

### 2026-09-30 (4)
| Wijziging |
|--------|
| Elke readme bestaat nu in drie talen: `readme.md` (Engels, nog steeds de hoofdversie), `readme.nl.md` (Nederlands) en `readme.fr.md` (Frans) — 66 mappen, de hoofd-readme inclusief de volledige Versiegeschiedenis. De taalwissel bovenaan elke pagina gaat naar dezelfde pagina in de andere taal, en het kruimelpad blijft binnen de taal die je leest |
| Koppen die een scriptnaam zijn (`### Set-UserManager.ps1`) worden niet vertaald, zodat elk `#…ps1`-anker in alle drie de talen gelijk is; andere koppen wel, met hun links binnen de pagina aangepast. Parameternamen, commando's, paden en de letterlijke teksten die een script toont of wegschrijft (Nederlandse Excel-tabnamen, foutmeldingen) blijven in elke taal zoals ze zijn |
| `Reporting/readme.md` en een deel van `SharePoint/readme.md` waren Nederlands in een Engelse set; die zijn eerst Engels gemaakt, en de Nederlandse versies houden de oorspronkelijke formulering aan |
| Het wachtwoord van de lokale admin dat in platte tekst in `scripts/Deployment/readme.md` en in deze Versiegeschiedenis staat, is **niet** meegekopieerd naar de Nederlandse en Franse versies; daar staat het als weggelaten |
| `.claude/CLAUDE.md` eist nu dat een wijziging aan een readme in alle drie de talen wordt gedaan, gevolgd door `Update-ReadmeHeader.ps1` |
| Geverifieerd met `Test-MarkdownLinks.ps1`: 204 markdownbestanden, elke interne link werkt; `Update-ReadmeHeader.ps1 -Check` meldt elke header actueel. De vertalingen zijn gecontroleerd op structuur (secties, tabellen, regelaantallen tegen het Engels), niet regel voor regel nagelezen door een moedertaalspreker |

### 2026-09-30 (3)
| Wijziging |
|--------|
| Nieuw `scripts/Startup/Update-ReadmeHeader.ps1` schrijft de twee regels bovenaan elke readme: een taalwissel (`English · Nederlands · Français`) en het kruimelpad terug omhoog, waarin elk niveau naar zijn readme in de huidige taal linkt. Met de hand bijgehouden zijn juist die relatieve paden wat breekt als een map verhuist; gegenereerd uit de map waarin een readme staat, kan dat niet. `-Check` stopt met exitcode 1 bij een verouderde header of een ontbrekende taalversie |
| `scripts/INDEX.md` opnieuw gegenereerd: naast het nieuwe script staan er nu ook `Revoke-SharePointUserAccess.ps1` en `Test-SharePointAccessScripts.ps1` in, die waren toegevoegd zonder `Update-ScriptIndex.ps1` opnieuw te draaien |
| Geverifieerd: syntaxcontrole schoon; gedraaid in PowerShell 7 en een `-Check`-run in Windows PowerShell 5.1 op deze repository — 66 mappen, 198 readmes, daarna elke header actueel |

### 2026-09-30 (2)
| Wijziging |
|--------|
| Elke readme begint nu met een breadcrumb (`M365-Scripts › scripts › Intune › Desktop`) die elk niveau terug naar boven linkt. Voorheen hadden 40 van de 66 mapreadmes geen weg terug naar hun bovenliggende map, behalve de terugknop van de browser |
| De root-readme opent met een `## Folders`-tabel die naar elke workloadmap linkt, zodat je de repository vanaf de voorpagina naar beneden kunt doorbladeren in plaats van alleen via `scripts/readme.md` |
| Submappen staan overal onder een `## Folders`-kop. `Device/`, `Network/`, `Reporting/` en `SharePoint/` mengden ze in de Scripts-tabel, `Intune/Desktop/` en `Custom Scripts/Intune/Desktop/` gebruikten een `Contents`-tabel, `TenantOnboarding/` zei `Subfolders` en `LegacyUtilities/` had helemaal geen kop |
| Geverifieerd met `Test-MarkdownLinks.ps1`: 1.072 interne links in 72 markdownbestanden worden correct opgelost. Alleen documentatie; geen script gewijzigd |

### 2026-09-29 (11)
| Wijziging |
|--------|
| `Repair-AppxPackageStore.ps1` provisiont de **exacte** Teams-/Outlook-build waarop FSLogix faalt. Een productiehost liet Outlook in een week 196× falen — 186× `0x80070490` — voor 1.2026.902 en 915, terwijl de host 818 provisionde en de bestanden van beide gevraagde builds op schijf stonden. Het eerdere oordeel ("met een actuele FSLogix is het verschil onschadelijk, geen actie nodig") was fout, en dat gold ook voor het najagen ervan met de installers, die alleen een oudere last-known-good build leveren |
| De MSIX voor één exacte build staat op Microsofts CDN op de geversioneerde URL die de manifests van winget gebruiken (`res.cdn.office.net/.../v2/<version>/Microsoft.OutlookForWindows_x64.msix`, `teamsinstaller.public.onecdn.static.microsoft/production-windows-x64/<version>/MSTeams-x64.msix`) — gecontroleerd dat die antwoordt voor Outlook 812/818/902/915 en Teams 26198/26225/26246. `-Provision` neemt nu de nieuwste build waarop FSLogix faalde, downloadt die, controleert de Microsoft-handtekening, provisiont hem en leest de geprovisionde versie terug; de installer wordt voor dat pakket overgeslagen |
| Foutcodes worden gedecodeerd met Windows' eigen melding voor elke Win32-code in plaats van een kort, handgeschreven lijstje, één code per regel: `0x80073D19` bleek "An error occurred because a user was logged off" te zijn — onschadelijk — en wordt nu ook zo gelabeld |
| Geverifieerd in PowerShell 5.1: het scenario van die host (FSLogix vraagt 902 en 915, host op 818) levert 915 met de juiste URL op, een **echte** download van die MSIX van 32 MB met een geldige Microsoft-handtekening, de geprovisionde versie teruggelezen (met `Add-AppxProvisionedPackage` gemockt), en geen doel meer zodra de host 915 heeft; eerdere scenario's ongewijzigd. **Niet op de host zelf uitgevoerd** |

### 2026-09-29 (10)
| Wijziging |
|--------|
| `Repair-AppxPackageStore.ps1 -Copilot -Provision` installeert nu de **nieuwe**, verenigde Microsoft Copilot-app in plaats van de oude Microsoft 365 Copilot-app. De installer die in (9) is toegevoegd levert het oude AppX-pakket, dat de unificatie daarna alsnog moet overzetten; de nieuwe app wordt machinebreed geïnstalleerd door Edge Update. De gedocumenteerde weg wordt gebruikt: `Install{C50565E9-...}` = 5 (Force Installs), `UpdaterExperimentationAndConfigurationServiceControl` = 1 (wat Force Installs vereist) en `CopilotUnificationAllowed{...}` = 1 onder `HKLM\SOFTWARE\Policies\Microsoft\EdgeUpdate`, geschreven na een `.reg`-back-up van die sleutel; daarna wordt de machinetaak van Edge Update gestart en wacht de run tot 10 minuten op de app onder `EdgeUpdate\Clients`. Verschijnt hij niet, dan is de oude installer de fallback. `Install` = 0 wordt nooit overschreven. Diagnose en verificatie tellen Copilot nu pas als aanwezig wanneer de nieuwe app er is |
| `-Name` begrijpt `teams`, `outlook` en `copilot`, dus de nieuwe Outlook en de nieuwe Copilot-app op de pool wordt `-ComputerName lem-avd-4,lem-avd-5,lem-avd-6 -Name outlook,copilot -Provision`. Het menu (`R`) accepteert dezelfde woorden |
| Gevonden tijdens het testen, opgelost: onder Windows PowerShell 5.1 met `ErrorActionPreference = Stop` is `reg.exe` dat naar stderr schrijft een afbrekende fout, dus één mislukte `.reg`-back-up zou de hele run hebben beëindigd in plaats van die sleutel met rust te laten. De back-up accepteert nu ook `HKLM:`/`HKCU:`-paden |
| Geverifieerd in PowerShell 5.1 en 7: de verkorte namen (vijf combinaties), de Edge Update-installatie tegen een gemockte policysleutel en taak — waarden geschreven, back-up gemaakt, het "verschijnen" van de app wordt opgepikt — en `Install` = 0 onaangeroerd gelaten; het eerdere store-scenario ongewijzigd. **Niet op een sessiehost uitgevoerd**: of Edge Update de app daar binnen de 10 minuten installeert, is niet getest |

### 2026-09-29 (9)
| Wijziging |
|--------|
| `Repair-AppxPackageStore.ps1 -Copilot` diagnosticeert en herstelt Copilot. Stap 1d rapporteert de Microsoft 365 Copilot-app (`Microsoft.MicrosoftOfficeHub`) en de Windows Copilot-app (`Microsoft.Copilot`) als geregistreerd en geprovisiond, de verenigde Microsoft Copilot-app die Edge Update sinds de unificatie van september 2026 installeert (gelezen uit `EdgeUpdate\Clients\{C50565E9-...}`), de Edge Update-versie tegenover de 1.3.253.25 die nodig is, en elke policy die Copilot tegenhoudt — `Install` / `Uninstall` / `Update{C50565E9-...}` onder `Policies\Microsoft\EdgeUpdate` (waarbij Force Installs, zoals gedocumenteerd, Uninstall overschrijft), de pauze op de unificatie, en Windows' `WindowsCopilot`- / `WindowsAI`-policy's machinebreed en per aangemelde gebruiker. Policy's worden met pad en waarde gerapporteerd en laten de run falen, maar worden nooit gewijzigd: ze komen uit GPO of Intune |
| Met `-Provision` wordt Copilot voor alle gebruikers geïnstalleerd met Microsofts gedocumenteerde `M365CopilotDesktopInstaller.exe --quiet --start -p` van `go.microsoft.com/fwlink/?linkid=2325486` — vandaag gecontroleerd dat die een door Microsoft ondertekende `xpdBootstrapper` 16.0.19305 levert — en geaccepteerd wanneer óf het AppX-pakket geprovisiond is óf de verenigde app onder Edge Update verschijnt. `-Copilot` neemt ook de Deprovisioned-markeringen van beide pakketten mee, en dat is precies wat een debloat-tool achterlaat. winget heeft er alleen een `.exe` voor, dus `-UseWinget` valt terug op de installer. De pooltabel kreeg een Copilot-kolom; het menu (`R`) accepteert `copilot` als antwoord voor het pakket |
| Het vorige item hernummerd naar (8): dat en het SharePoint-item hieronder waren dezelfde middag allebei als (7) gecommit |
| Geverifieerd in PowerShell 5.1 en 7: stap 1d tegen het echte register en de echte pakketten van deze machine, en tegen gemockte Edge Update-policy's (Uninstall alleen laat de run falen, Uninstall met Force Installs niet); de pooltabel van de orchestrator met de nieuwe kolom; de download van de installer en de handtekening ervan. **Niet op een sessiehost uitgevoerd**: het provisionen via `-p` van de installer en het daarna verschijnen van de verenigde app zijn niet getest |

### 2026-09-29 (8)
| Wijziging |
|--------|
| `Repair-AppxPackageStore.ps1` draait over een hele pool met `-ComputerName lem-avd-4,lem-avd-5,lem-avd-6` (optioneel `-Credential`): het kopieert zichzelf via PowerShell remoting naar `C:\IT\AppxRepair` op elke host, draait daar met dezelfde parameters — de eigen uitvoer van de host streamt terug — en eindigt met één tabel over de hele pool (exitcode, FSLogix-build, geprovisionde Teams / Outlook) die elk verschil tussen hosts benoemt. Een reparatie wordt eenmaal voor de hele pool bevestigd, omdat een remote sessie een bevestigingsprompt niet betrouwbaar kan beantwoorden. Opgenomen in het menu (toets `R` vraagt om de hosts) |
| Het advies van de verificatiestap was fout. Na een live `-Provision`-run zei het dat Teams (26225) en Outlook (1.2026.818) "still older than the 26246 / 902 profiles ask for - bring the other hosts to the same build" waren. Maar geen enkele host loopt voor: beide apps werken zichzelf per gebruiker bij, en Microsofts installers provisionen een last-known-good build die daarachter ligt, dus het profiel loopt altijd voor op elke host en een nieuwere build provisionen houdt alleen stand tot de volgende update. Wat bepaalt of het pijn doet is FSLogix: vanaf 2210 HF4 (Teams) / 25.06 (Outlook) registreert het op family name en is het verschil onschadelijk (nu als OK gerapporteerd, exitcode 0); op een oudere build is het advies om FSLogix bij te werken. Zo'n verschil telt niet langer als iets om te provisionen |
| Een transcript dat niet wil starten — zoals in sommige remote- en RMM-sessies — breekt de reparatie niet meer af; het is een waarschuwing |
| Geverifieerd in PowerShell 5.1 en 7: de orchestrator tegen twee onbereikbare hosts (elk benoemd met zijn WinRM-fout, de pooltabel, exitcode 1), en het versieverschil met een gemockte FSLogix boven en onder het minimum. **Niet tegen echte sessiehosts uitgevoerd**: van hieruit geen WinRM naar lem-avd-4/5/6, dus het kopiëren, de remote run en de pooltabel met echte waarden zijn niet getest |

### 2026-09-30 (2)
| Wijziging |
|--------|
| Hardeningronde op `scripts/SharePoint/Revoke-SharePointUserAccess.ps1`, gestuurd door de code te lezen op de faalwijzen die een destructief script heeft in plaats van die van een rapport. Drie ervan waren echt |
| **Het matchen van een gast had de verkeerde persoon kunnen intrekken.** De fallback-lookup vergeleek met `-like "*needle*"`, en `an@contoso.com` is een substring van `jan@contoso.com`. Vervangen door een exacte vergelijking met de UPN, het mailadres, het claim-achtervoegsel en een correct gedecodeerde gastlogin (`jan_partner.com#ext#@tenant` terug naar `jan@partner.com`, gesplitst op de laatste underscore zodat het lokale deel er zelf een mag bevatten). Wanneer twee verschillende accounts op hetzelfde adres antwoorden, blijft de site onaangeroerd en stopt de run met beide bij naam te noemen — kiezen is aan de operator, niet aan het script |
| **De audit-CSV werd één keer geschreven, aan het einde.** Een run die tweehonderd dingen had ingetrokken en daarna crashte, zou geen enkel spoor hebben achtergelaten van wat hij had verwijderd, en dat is het enige wat een script als dit nooit mag doen. Rijen worden nu toegevoegd op het moment dat het gebeurt, via de gedeelde helper die een vergrendeld bestand opnieuw probeert en de run stopt in plaats van een rij te laten vallen |
| **Een `404` bij een verwijdering telde als fout.** Het betekent dat de toekenning al weg is, wat bij een tweede ronde de normale uitkomst is — een schone herhaling zou fouten hebben gerapporteerd. Wordt nu vastgelegd als `AlreadyGone` |
| Een mislukte verwijdering van een sitecollectiebeheerder is nu luid en telt als fout: die rol reikt tot elk bereik in de site, dus elke andere verwijdering daar is cosmetisch zolang hij blijft staan. De samenvatting zegt dat expliciet in plaats van als een succes te lezen |
| `-WhatIf` neemt nu dezelfde tak als een proefdraai, dus legt het `WouldRevoke` vast in plaats van `Skipped`, wat suggereerde dat iemand een prompt had afgewezen |
| `Test-SharePointAccessScripts.ps1` groeide van 27 naar 50 controles: het decoderen van gastlogins inclusief een lokaal deel met underscore, exacte matching tegen de bijna-treffers die een substringtest zou hebben geaccepteerd (korter, langer, domein met achtervoegsel, de gast van een andere tenant, leeg), de audit-CSV die halverwege de run bestaat en elke rij bevat, en een 404 die als al verdwenen wordt gelezen. **Nog steeds niet geverifieerd op een live tenant** |

### 2026-09-30
| Wijziging |
|--------|
| Een dode `-Restart`-parameter verwijderd uit `scripts/SharePoint/Revoke-SharePointUserAccess.ps1`. Hij was gedeclareerd en de help beloofde dat hij `discard any existing checkpoint and start over instead of resuming` zou doen - maar het script heeft geen checkpoint en geen hervatting, dus de switch deed niets en de help beschreef gedrag dat niet bestaat. Gevonden door het parameterblok te vergelijken met de comment-based help en de mapreadme in plaats van aan te nemen dat ze overeenkwamen |
| Vervangen door een `.NOTES`-regel die uitlegt waarom er bewust geen hervatting is: intrekken is idempotent, dus een tweede run vindt alleen wat de eerste niet heeft verwijderd. Opnieuw draaien na een onderbreking is zowel het herstel als de verificatie, en veiliger dan een half toegepaste destructieve operatie hervatten vanaf een opgeslagen positie |
| Geverifieerd dat alle 17 resterende parameters in de comment-based help en de mapreadme staan, dat `Get-Help` `-Restart` niet meer noemt, en dat de 27 controles in `Test-SharePointAccessScripts.ps1` nog steeds slagen |

### 2026-09-29 (7)
| Wijziging |
|--------|
| `scripts/SharePoint/Revoke-SharePointUserAccess.ps1` toegevoegd — de tegenhanger van het rechtenrapport. Het vindt elke plek waar één genoemde gebruiker toegang heeft en verwijdert die: eerst sitecollectiebeheerder (dat overschrijft alles eronder, dus laten staan zou de rest cosmetisch maken), directe roltoewijzingen op sites, subsites, lijsten, mappen en losse bestanden, lidmaatschap van SharePoint-groepen, en de `SharingLinks.*`-groepen die "Anyone with the link" en "Specific people" dragen. Rapporteren is de standaard; er verandert niets zonder `-Apply`, en elke run schrijft een CSV van wat er gevonden is en wat ermee gebeurde |
| Het weigert bewust twee dingen en zegt dat luid. Een toekenning via een Entra ID-groep wordt niet ingetrokken — de groep *is* de toekenning, en de gebruiker uit SharePoint verwijderen zou de toegang laten staan terwijl het lijkt alsof hij dicht is; de groep wordt in de CSV genoemd onder `Action = CannotRevoke`, zodat offboarding zichtbaar twee stappen is. Een toekenning aan `Everyone` blijft om de omgekeerde reden staan: die verwijderen trekt de toegang in voor de hele tenant in plaats van voor deze persoon |
| De app-only-authenticatie en de SharePoint REST-laag worden byte voor byte gedeeld met het rechtenrapport, afgebakend door `SHARED BLOCK START/END`. Het kostte vier live runs tegen een tenant om dat goed te krijgen, en een tweede kopie die ongemerkt afwijkt is een correctheidsrisico in het script dat rechten verwijdert. Om het deelbaar te maken is de banner van het rapport boven het blok gezet en komt de naam van de tijdelijke app nu uit `$TempAppNamePrefix`; het gedrag van het rapport is ongewijzigd |
| `scripts/SharePoint/Test-SharePointAccessScripts.ps1` toegevoegd, dat controleert dat de twee kopieën identiek zijn (en de eerste afwijkende regel toont wanneer dat niet zo is) en de intrekkingstrechter echt doorloopt: een proefdraai legt zijn intentie vast en voert niets uit, `-Apply` voert uit en legt vast, een fout belandt in het audittrail in plaats van te verdwijnen, en de weigeringen hierboven blijven weigeringen. 27 controles, allemaal geslaagd, uit te voeren vanuit elke map |
| Het menu-item (toets `W`) toegevoegd en beide scripts gedocumenteerd in de SharePoint-mapreadme, de repositoryboom en deze categorielijst. Geverifieerd dat elk `docs`-anker in die readme correct wordt opgelost |

### 2026-09-29 (6)
| Wijziging |
|--------|
| `Repair-AppxPackageStore.ps1` probeert niet langer een pakket opnieuw te registreren dat vervangen is. De eerste live run probeerde in stap 3 `aimgr_0.20.61.0` en kreeg `0x80073D06` ("a higher version 0.20.62.0 of this package is already installed"): een oude versie waarvan de status niet Ok is terwijl er een nieuwere van hetzelfde pakket naast staat, is geen schade maar Windows dat wacht om hem te verwijderen, en opnieuw registreren kan nooit slagen. De diagnose vergelijkt nu versies per pakketnaam, architectuur en resource-id, rapporteert deze als *Superseded* in één grijze regel en laat ze buiten de reparatietelling; stap 3 behandelt een `0x80073D06` die toch opduikt ook als "overgelaten aan Windows" in plaats van als waarschuwing |
| Offline geverifieerd in PowerShell 7 en 5.1 met precies dat paar (0.20.61.0 `Modified` naast 0.20.62.0 `Ok`): gerapporteerd als vervangen en niet meegeteld, terwijl een echt beschadigd pakket in dezelfde run nog steeds wordt opgepikt voor herregistratie. Nog niet opnieuw uitgevoerd op de host waar het gebeurde |

### 2026-09-29 (5)
| Wijziging |
|--------|
| `Repair-AppxPackageStore.ps1` kan Teams en de nieuwe Outlook uit winget halen (`-UseWinget`): `winget download` van `Microsoft.Teams` / `Microsoft.Outlook`, waarvan de manifests naar de MSIX op Microsofts CDN wijzen, en daarna `Add-AppxProvisionedPackage` met eventuele afhankelijkheden die winget meebracht, zodat het pakket voor alle gebruikers landt in plaats van alleen voor wie `winget install` draaide. Elk bestand moet een geldige Microsoft-handtekening hebben. De manifests van winget lopen achter op Microsofts installers (vandaag gecontroleerd: Teams 26198 tegenover de 26246 waar de profielen om vragen, Outlook 1.2026.812 tegenover 902), dus de installers blijven de standaard en de run waarschuwt wanneer de build van winget ouder is dan waar de profielen om vragen. `-WingetId` provisiont elke andere app op dezelfde manier |
| Het toont nu elke app die faalt, niet alleen die in de package store: stap 1c leest het AppX-deploymentlog en het FSLogix Apps-log over `-Days`, gegroepeerd per pakket, met de foutcodes benoemd (`0x80073D02` in gebruik, `0x80073CF6` registratie mislukt, ...), de gevraagde versies en of deze host hun bestanden heeft; `0x80070490` eerst, top 15 |
| Gevonden tijdens het testen, opgelost: `Get-WinEvent` gooit een afbrekende fout voor een provider die niet geregistreerd is — elke machine zonder FSLogix — en die wordt niet door `-ErrorAction SilentlyContinue` afgevangen, dus de FSLogix-controle zou de run daar hebben afgebroken; alle eventreads gaan nu via één wrapper. En het voortgangsfilter voor winget bevatte twee niet-ASCII-tekens, die Windows PowerShell 5.1 in een bestand zonder BOM als ANSI leest en dan niet kan parsen — het hele script zou niet vanuit NinjaOne hebben gedraaid. Het bestand is weer puur ASCII, gecontroleerd |
| Geverifieerd in PowerShell 7 en 5.1: het overzicht van falende apps tegen het **echte** AppX-log van deze machine (20 pakketten, codes vertaald, afgekapt op 15); een **echte** `winget download` van `Microsoft.Outlook` (MSIX van 32 MB, hash geverifieerd door winget, handtekening door het script) tot en met een gemockte `Add-AppxProvisionedPackage`; en het eerdere gemockte store-scenario, ongewijzigd. Het provisionen zelf en de Teams-download (271 MB) zijn hier niet uitgevoerd |

### 2026-09-29 (4)
| Wijziging |
|--------|
| Nieuw `scripts/Device/Repair-AppxPackageStore.ps1`: repareert AppX-pakketten die falen met `0x80070490` en een leeg pad ("Deployment Register operation ... from:  (AppxManifest.xml)"), voor elk pakket — dezelfde fout kwam terug voor `Microsoft.OutlookForWindows` op een host waar alleen Teams een reparatie had, en `Update-TeamsClient.ps1 -RepairAppxStore` is bewust beperkt tot `MSTeams` |
| De fouten op die host werden gelogd door `Apps (Microsoft-FSLogix-Apps)`, en dat is een andere oorzaak dan een beschadigde store: FSLogix slaat de pakketten van elke gebruiker op exacte versie op in `AppxPackages.xml` en speelt ze bij het aanmelden opnieuw af (`InstallAppxPackages`, standaard aan), dus een host die een andere build provisiont — of geen — antwoordt `0x80070490`. Het script leest die events en vergelijkt de versie waar de profielen om vragen met wat de host provisiont, controleert de FSLogix-build tegen de eerste releases die Teams (2210 HF4) en Outlook (25.06) op family name registreren, en zet met `-Provision` Teams / de nieuwe Outlook voor alle gebruikers terug met Microsofts eigen installer (`teamsbootstrapper.exe -p`, Outlook `Setup.exe --provision true --quiet --start-`; van beide links gecontroleerd dat ze naar Microsofts CDN leiden, van beide de handtekening gecontroleerd vóór het uitvoeren) |
| De store-reparatie generaliseert die voor Teams en voegt toe wat het veilig maakt om op een hele store te draaien: elke registersleutel wordt naar een `.reg`-back-up geëxporteerd voordat hij wordt verwijderd, en niet verwijderd als de back-up mislukt; met een wildcard-`-Name` worden systeem- en frameworkpakketten gerapporteerd maar nooit aangeraakt en blijven Deprovisioned-markeringen (zo worden verwijderingen van bloatware onthouden) met rust; een gebruikersregistratie telt ook als verweesd wanneer het pakket nergens bestanden heeft, niet alleen wanneer de SID geen profiel heeft. Het bewerken van `StateRepository-Machine.srd` of `AppxPackages.xml` is onderzocht en bewust weggelaten — beide worden niet ondersteund |
| Opgenomen in `menu.ps1` als Device-toets `R` (diagnose, tenzij je de reparatie bevestigt; vraagt apart naar provisionen), gedocumenteerd in de Device-readme |
| Offline geverifieerd in PowerShell 7 en 5.1 met gemockte AppX-cmdlets, FSLogix-events en een kladversie van `AppxAllUserStore` in HKCU: het Teams-pakket met leeg pad wordt als ghost gevonden, de gebruikers- en machine-items ervan als verweesd, een Outlook-item voor een SID zonder profiel als verweesd, de oudere geprovisionde Teams en de ontbrekende Outlook als te provisionen, een framework-ghost en een Deprovisioned-markering blijven met `*` met rust en worden meegenomen wanneer ze bij naam genoemd worden, en de `.reg`-back-up wordt geschreven. Die run ving ook dat `-Name A,B` via `powershell.exe -File` (en via de eigen herstarts van het script) als één string binnenkwam; wordt nu gesplitst. **Niet op een live host uitgevoerd**: er was hier geen elevatie of AVD-host beschikbaar, dus de verwijderingen, de registerwijzigingen en beide installers zijn niet echt uitgeprobeerd |

### 2026-09-29 (3)
| Wijziging |
|--------|
| `Restore-MailboxMessages.ps1` behandelt "alles vanaf deze datum tot nu" nu als volwaardig geval: `-After` kreeg de aliassen `-From` en `-Since`, en het menu-item (`L`) vraagt of je alleen die dag wilt herstellen of alles sindsdien. Het venster zelf stond het al toe, maar de auditzoekopdracht liep als één query over het hele venster, en één zoeksessie stopt tenantbreed bij 50.000 records — over een paar weken liet dat ongemerkt acties op de te herstellen mailbox vallen |
| Het auditlog wordt nu per dag doorzocht, elke dag in een eigen sessie met paging, en alleen records die deze mailbox noemen worden in het geheugen gehouden. Een dag die op zichzelf 50.000 records overschrijdt, wordt in een waarschuwing genoemd |
| Een lang venster kan verder teruggaan dan wat de mailbox nog bewaart, wat zou lezen als "er is niets verwijderd". De run waarschuwt nu wanneer het venster begint vóór de `RetainDeletedItemsFor` van de mailbox (standaard 14 dagen) en de mailbox niet op hold staat, en wanneer het meer dan 180 dagen geleden begint, voorbij de gebruikelijke auditbewaartermijn |
| Offline geverifieerd in PowerShell 7 en 5.1: `-Since` bindt aan `-After`; een gemockte `Search-UnifiedAuditLog` over een venster van 2,6 dagen werd per dag aangeroepen met UTC-grenzen, waarbij de laatste schijf eindigt op het einde van het venster, dag één over twee aanroepen in één sessie werd gepagineerd, en alleen de records van deze mailbox werden bewaard. De bewaarwaarschuwing is **niet** getest - die heeft een live `Get-Mailbox` nodig |

### 2026-09-29 (2)
| Wijziging |
|--------|
| `Restore-MailboxMessages.ps1` heeft de rol Mailbox Import Export niet meer nodig om verwijderde mail terug te halen. De eerste echte run stopte bij "Get-RecoverableItems is not available" en sloeg elk verwijderd bericht over, terwijl de rol standaard in geen enkele rolgroep zit — dus op de meeste tenants deed het verwijderde deel gewoon niets |
| Zonder de rol schakelt de run nu over naar Graph en herstelt **alles** wat in het venster is verwijderd, niet alleen wat het auditlog zag: elk bericht in Deleted Items en Recoverable Items\Deletions waarvan de wijzigingstijd in het venster valt, gaat terug. Geauditeerde verwijderingen (`MoveToDeletedItems`, `SoftDelete`, die Exchange standaard voor de eigenaar auditeert) gaan naar de map die ze volgens het record verlieten, met de actor exact gematcht op MessageId; de rest gaat naar de Inbox. Een bericht dat *uit* Deleted Items is verwijderd gaat ook naar de Inbox, want het terugzetten in Deleted Items is geen herstel. Definitief verwijderde items (Purges) liggen buiten het bereik van Graph en worden als `Unreachable` gerapporteerd in plaats van ongemerkt te ontbreken |
| Het opzoeken van berichten doorzoekt nu ook `recoverableitemsdeletions`, dat `/messages` niet dekt, zodat een bericht dat eerst verplaatst en daarna verwijderd is in beide delen wordt gevonden. Het verplaatste en het verwijderde deel delen nu één pad voor opzoeken / verplaatsen / rapporteren in plaats van twee kopieën |
| Offline geverifieerd in PowerShell 7 en 5.1 tegen een gemockte Graph: verplaatst-en-daarna-soft-deleted gaat terug naar de oorspronkelijke submap, verwijderd-uit-Deleted-Items gaat naar de Inbox, niet-geauditeerde items uit beide mappen worden hersteld, een item dat vijf dagen eerder is verwijderd blijft met rust, een definitieve verwijdering wordt als `Unreachable` gerapporteerd, en de preview en `-Apply` voeren precies de verwachte verplaatsingen uit. **Niet tegen een live tenant uitgevoerd**; met name of Graph een verplaatsing uit `recoverableitemsdeletions` toestaat, en of een verplaatsing `lastModifiedDateTime` verandert, is niet getest |

### 2026-09-29
| Wijziging |
|--------|
| Nieuw `scripts/Exchange/Restore-MailboxMessages.ps1`: zet de berichten terug die op een bepaalde dag in één mailbox zijn verplaatst of verwijderd, en vertelt wie het deed. Er was geen weg terug na een mislukte archiveringsrun of een massale verwijdering, behalve met de hand herstellen in Outlook, en geen antwoord op "wie heeft dit gedaan" zonder vanaf nul een auditlogquery te schrijven |
| Verwijderde berichten gaan terug via `Get-/Restore-RecoverableItems` (Deleted Items, Recoverable Items, Purges), één `EntryID` per keer, gefilterd op het moment van verwijderen — Exchange kent de oorspronkelijke map zelf. Na een `-Apply` worden de mappen opnieuw gelezen, en alles wat er dan nog staat wordt als `NotRestored` gerapporteerd in plaats van op het stilzwijgen van de cmdlet te vertrouwen |
| Verplaatste berichten hebben zo'n geheugen niet: Graph noch Exchange legt vast waar een verplaatst bericht vandaan kwam. De Unified Audit Log wel, dus elke geauditeerde `Move` wordt herleid tot de **eerste** map die het bericht die dag verliet, via Graph gevonden op Internet MessageId en via `$batch` teruggezet. Mappen worden gematcht op hun pad zoals het auditlog dat schrijft, en dat is in de eigen taal van de mailbox (`\Postvak IN\Projecten`). Verplaatsingen uit Deleted Items of Recoverable Items worden overgeslagen, omdat dat herstelacties waren en het terugdraaien ervan het bericht opnieuw zou verwijderen |
| Dezelfde auditrecords noemen de actor — account, eigenaar/gedelegeerde/beheerder, client (Outlook, OWA, Graph-app met app-ID), IP — per bericht in de CSV, als gegroepeerde tabel "wie verplaatste / verwijderde wat" op het scherm, en als ruwe `_Audit.csv`. Verwijderingen worden toegeschreven op onderwerp en dichtstbijzijnde tijd, omdat herstelbare items geen MessageId hebben. De run toont ook welke van de vier acties niet op de mailbox worden geauditeerd, omdat de eigen `Move` van de eigenaar standaard niet wordt geauditeerd en een ontbrekend record anders zou lezen als "niemand heeft het gedaan" |
| Archiefitems zonder auditrecord (bijv. na `Move-InboxToArchive.ps1`) worden op wijzigingstijd getoond en alleen met `-UnauditedArchiveToInbox` naar de Inbox verplaatst, omdat lezen of markeren die tijd ook verandert. Graph-toegang hergebruikt het REST-only patroon met drie routes van `Remove-PhishingMessage.ps1`, zodat het naast de Exchange-sessie draait zonder het MSAL-conflict. Toegevoegd aan het Exchange-submenu als `L` (eerst preview, dan `-Apply`) |
| Alleen offline geverifieerd, in PowerShell 7 en Windows PowerShell 5.1: syntaxcontrole, het parsen van auditrecords tegen verzonnen records (mailboxfilter, UTC naar lokale tijd, logontypes, clientlabels, toeschrijving op onderwerp/tijd), en het hele pad voor verplaatste berichten tegen een gemockte Graph — een keten van verplaatsingen die teruggaat naar de eerste map, Nederlandse mapnamen, een herstelactie van een gebruiker die met rust wordt gelaten, een al teruggezet bericht dat wordt overgeslagen, een geauditeerd archiefitem dat buiten de niet-geauditeerde lijst blijft, en de resulterende verplaatsingsverzoeken. **Nog niet tegen een live tenant uitgevoerd**: de exacte uitvoereigenschappen van `Get-RecoverableItems`, hoe het de filtertijden interpreteert, en of de auditrecords voor elke client `InternetMessageId` bevatten, zijn allemaal niet getest |

### 2026-09-28 (6)
| Wijziging |
|--------|
| De herstart die `-RestartIfNeeded` activeert wachtte 60 seconden op een host waar niemand kon meekijken. Het aftellen bestaat om mensen te waarschuwen, dus het geldt nu alleen wanneer er mensen zijn: met iemand aangemeld is het `-RestartDelaySeconds` en stopt `shutdown /a` het; met niemand aangemeld - het normale geval bij het opstarten, en het gegarandeerde op een sessiehost waarvan de pool op drain staat - herstart hij binnen enkele seconden. Een paar seconden blijven behouden zodat de eigen logregel van de run wordt geschreven voordat het afsluiten begint |
| Overwogen om aanmeldingen vanuit de guest te drainen (`change logon /drainuntilrestart`) om het venster tussen het starten van de taak en de herstart te sluiten, en dat laten vallen: een host die op deze manier wordt herstart heeft zijn pool al op drain staan, dus de switch aan de guestkant zou alleen dupliceren wat de pool garandeert - en hij zou aanmeldingen geblokkeerd laten bij elke run die crashte voordat hij ze weer inschakelde |
| Geverifieerd dat `change.exe` en `chglogon.exe` op deze Windows 11-build bestaan en dat `change logon /query` "Session logins are currently ENABLED" meldt terwijl het met 1 afsluit; daarom is het idee gemeten voordat het werd laten vallen in plaats van erna |

### 2026-09-28 (5)
| Wijziging |
|--------|
| `Update-TeamsClient.ps1` kan de host nu repareren in plaats van hem alleen te diagnosticeren. Een sessiehost waar elke route `0x80070490` antwoordde, had een `AppxAllUserStore` vol items die Windows niet meer kan oplossen, en het eerlijke advies was op dat punt "opnieuw uitrollen" — niet wat iemand wil horen over een machine die verder prima is |
| `-RepairAppxStore` doet het in twee stappen. Een pakket waarvan de bestanden nog op schijf staan, wordt opnieuw geregistreerd vanuit zijn eigen manifest (`Add-AppxPackage -Register`), wat de kennis van de store erover opnieuw opbouwt en de gewone verwijdering meestal weer laat werken. Wat dat overleeft wordt sleutel voor sleutel verwijderd: registraties onder een SID zonder profiel op deze host (ook onder `EndOfLife` en `DeferredRemoval`), een machinebreed `Applications`-item waarvan het manifest weg is, en de `Deprovisioned`-markering die het provisionen botweg weigert. Elke sleutel wordt met zijn volledige registerpad genoemd voordat hij verdwijnt, en niets buiten MSTeams wordt ooit aangeraakt |
| Preflight rapporteert die weesitems ongeacht of de switch is opgegeven, dus `-CheckOnly` is de diagnose en de reparatie een aparte beslissing — dezelfde opzet als `-ClearOrphanedAddInRegistration` voor Windows Installer |
| `-UseWinget` pakt het van de andere kant aan: winget downloadt de MSIX en controleert die tegen de SHA256 in zijn eigen manifest, en de bootstrapper provisiont dat bestand met `-p -o`. De deployment heeft dan een expliciete bron in plaats van een store-item dat hij moet oplossen, en de run weet welke build hij heeft geïnstalleerd. Het manifest van winget loopt achter op de configservice — gemeten op `26198.304.4946.9672` tegenover een `26246`-build — en de run zegt dat wanneer het zo is |
| Als System staat winget helemaal niet op het `PATH`: de alias is een MSIX-shim per gebruiker. Hij wordt in plaats daarvan opgelost vanuit `Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*`, geverifieerd door `PATH` leeg te maken en te zien hoe de fallback hem vindt |
| De store-lezer is uitgevoerd tegen de live `AppxAllUserStore` van dit werkstation, waar hij twee echte weesitems vond (`S-1-0-0` en de SID van een verwijderd profiel onder `EndOfLife`), er geen rapporteerde voor gezonde pakketten en elk pad binnen de store hield. De **verwijdering** is echt uitgeprobeerd tegen een store die onder `HKCU` is nagebouwd: 9 MSTeams-items, 6 verweesd, alle 6 verwijderd, terwijl de gezonde registraties, het `Staged`-item, het weesitem van een ander product en de eigen sleutel van de dode SID allemaal bleven staan |
| `winget download` is van begin tot eind gemeten: 271 MB in 23 seconden, geen Store-account, een eigen hashcontrole, een geldige `O=Microsoft Corporation`-handtekening, de stagingmap eerst geleegd en daarna verwijderd. Tegen een host waarvan de package store echt beschadigd is, zijn beide switches **niet getest** — geen enkele machine hier heeft er een |
| De herstartregel noemt geen reden meer. Hij zei "MSI returned 3010" terwijl drie verschillende dingen hem zetten, en wees lezers naar een installer die nooit had gedraaid |

### 2026-09-28 (4)
| Wijziging |
|--------|
| `Update-TeamsClient.ps1` gaf advies dat niet kon helpen. Een productiesessiehost antwoordde `0x80070490` ("Element not found") voor **elke** houder van het pakket, `NT AUTHORITY\SYSTEM` inbegrepen, en het script zei nog steeds de host te drainen en gebruikers af te melden — op een host die al gedraind was en waar niemand op zat. Een registratie die de package store niet kan vinden is geen gebruiker die het pakket vasthoudt |
| Preflight zegt nu of een wachtende verwijdering nog op iemand te wachten heeft. `Installed(pending removal)` wordt pas bij een afmelding voltooid, dus een SID zonder profiel onder `ProfileList` wacht op een gebeurtenis die nooit kan plaatsvinden; die worden apart gerapporteerd van de SID's die echt wachten, en alleen die laatste markeren een herstart |
| Het controleert ook de ene toestand waar niets zichzelf van herstelt: een pakket dat de store vermeldt waarvan de `InstallLocation` weg is, of dat helemaal geen installatielocatie heeft. Die ene regel verklaart de hele fout — elke verwijdering antwoordt `0x80070490` omdat er niets te verwijderen is, en het provisionen van dezelfde versie antwoordt dat ook |
| Wanneer de verwijdering per gebruiker die code voor elke houder antwoordt, zegt de run dat duidelijk, en de uiteindelijke fout past zijn advies daarop aan: niet "drain de host", maar dat `Remove-AppxPackage`, de bootstrapper en DISM allemaal dezelfde inconsistente store lezen, zodat geen van hen hem kan repareren — een gepoolde sessiehost wordt opnieuw uitgerold vanaf zijn image, een persoonlijke wordt ter plekke gerepareerd |
| Uitgeprobeerd tegen de strings die die host echt printte, plus een live SID van deze machine als contrastgeval: vier verweesde profielen lezen als verweesd, een echt profiel leest nog steeds als "meld ze af", een pakket zonder installatielocatie wordt gemarkeerd, en een gezond pakket blijft stil. De **remediatie** is niet getest — deze machine heeft geen beschadigde package store om het op te proberen |

### 2026-09-28 (3)
| Wijziging |
|--------|
| `Init-TempDisk.ps1` repareerde de tijdelijke schijf en configureerde het wisselbestand erop, en liet de machine daarna de rest van die sessie zonder draaien - Windows leest de wisselbestandconfiguratie bij het opstarten en leest die nooit opnieuw, dus de boot die `D:` opnieuw moest opbouwen is precies de boot waarop het wisselbestand niet bestaat. `-RestartIfNeeded` toegevoegd, dat dat gat dicht in plaats van op de volgende boot te wachten |
| Een script dat bij elke boot draait en de machine mag herstarten is een herstartlus die op het punt staat te gebeuren, dus het gaat alleen af wanneer dit allemaal geldt: de run is schoon afgerond (een mislukte run herstart nooit - dat zou de fout achter een herstart verbergen), de schijf is er, het wisselbestand is erop geconfigureerd, en het enige wat ontbreekt is dat deze sessie het niet gebruikt |
| Er mag niemand aangemeld zijn, verbonden of niet verbonden. Sessies worden geteld als één `explorer.exe` per interactief bureaublad in plaats van door `query.exe` te parsen, waarvan de kolomkoppen de weergavetaal volgen en op een Nederlandse sessiehost een lege lijst zouden opleveren. `-RestartEvenIfUsersSignedIn` overschrijft dit waar het aftellen waarschuwing genoeg is |
| Hoogstens één herstart per `-RestartCooldownMinutes` (standaard 60), onthouden als round-trip-tijdstempel onder `HKLM:\SOFTWARE\ICTKanon\InitTempDisk` - een tijdstempel in landinstellingsopmaak dat door de ene run wordt geschreven en door een andere gelezen, is hoe een cooldown ongemerkt ophoudt te werken. Een tweede herstart voor hetzelfde betekent dat de eerste niet hielp, en de run zegt dat in plaats van hem te herhalen |
| De herstart gaat via `shutdown.exe` met een aftelling van 60 seconden en de geplande reden "Operating System: Reconfiguration", zodat iedereen op de machine hem ziet aankomen, `shutdown /a` hem stopt, en hij niet als onverwachte herstart wordt gerapporteerd |
| `Register-InitTempDiskTask.ps1` rolt de taak nu uit met `-Quiet -RestartIfNeeded`, en zegt bij het registreren of de taak de machine mag herstarten. `-ScriptArguments '-Quiet'` laat de herstart weg |
| De twee eerdere items van vandaag genummerd: er waren twee identieke `### 2026-09-28`-koppen in de geschiedenis beland, wat leest als één wijziging die in tweeën is gesplitst |
| Geverifieerd op deze machine onder PowerShell 5.1 en 7: de sessiedetectie noemt het aangemelde account (dus deze machine zou weigeren te herstarten), een ontbrekende markering leest als `$null`, een geschreven en teruggelezen markering wordt geparsed naar een `DateTime` en blokkeert een tweede herstart binnen de cooldown, een markering van 90 minuten oud staat er een toe, en een corrupte markering degradeert naar "geen markering" in plaats van een fout te gooien |
| Ook gemeten hoe een falende `shutdown.exe` zich meldt, omdat de code daarop vertakt: `shutdown /a` zonder iets in behandeling antwoordt 1116 en zet `$LASTEXITCODE` in zowel 5.1 als 7.6 zonder een fout te gooien, dus de tak voor de exitcode is degene die draait. De `try` eromheen blijft voor `$PSNativeCommandUseErrorActionPreference`, dat daar op 7.4 en later een afbrekende fout van kan maken |
| **Er is vanuit deze sessie geen herstart geactiveerd** en de beveiligingen blijven ongeverifieerd tegen een live Azure-VM |

### 2026-09-28 (2)
| Wijziging |
|--------|
| `Update-TeamsClient.ps1` crashte in de preflight op elke machine waar de meeting-add-in nergens geregistreerd is: `The property 'Count' cannot be found on this object`. `$x = if (...) { @() }` kent `$null` toe, omdat een lege array die naar de pipeline wordt geschreven nul objecten is — de `@()` moet om de hele `if` heen, niet binnen de takken ervan. Gereproduceerd tegen de gecommitte versie en opgelost; alle vijf paden door de rapportagefunctie slagen nu, en de drie lege gooiden aantoonbaar eerder een fout |
| Een package dat de AppX-stack weigert te verwijderen, beëindigt de run niet meer. `Remove-AppxPackage -AllUsers` antwoordde `Catastrophic failure` op een session host met twee `MSTeams`-versies, en `-ErrorAction` dekt geen terminating error af, dus er was een `try`/`catch` nodig. De run gaat door en de provisioning upgradet ter plekke wat er is overgebleven — daar afbreken had de host achtergelaten met de add-in gedeïnstalleerd en geen Teams teruggezet |
| De **deïnstallatie** van de add-in is verplaatst van stap 6 naar stap 8, naast de installatie die hem vervangt. De opruiming was daar al naartoe verhuisd; de deïnstallatie achterlaten betekende dat elke latere fout via een andere route hetzelfde resultaat opleverde. Alles wat destructief is aan de add-in staat nu bij datgene wat het ongedaan maakt |
| Exitcodes zijn leesbaar. `teamsbootstrapper.exe` antwoordt met een HRESULT, die PowerShell als een groot negatief geheel getal toont: "exit code -2147023728" zegt niets, `0x80070490 - Element not found` zegt waar je moet kijken. MSI-codes blijven gewone getallen, en een HRESULT buiten de Win32-facility valt terug op kale hex in plaats van een betekenis te verzinnen |
| Een mislukte provisioning probeert nu eenmalig de door Microsoft gedocumenteerde machinebrede deïnstallatie (`teamsbootstrapper.exe -x -m`) en provisiont opnieuw voordat hij opgeeft, en de fout die hij dan geeft noemt de gebruikelijke oorzaak op een session host: een package dat wordt vastgehouden door een aangemelde gebruiker. **Niet getest** — dat herstel heeft nog niet gedraaid op een host die het nodig had |
| Die foutmelding at vervolgens haar eigen exitcode op: `-f` bindt sterker dan `+`, dus het formatteren van een samengevoegde string paste de opmaak alleen toe op het laatste stuk en toonde een letterlijke `{0}`. Gemeld vanaf een live host, waar het juist de code verborg die de lezer nodig had |
| De bootstrapper print zijn eigen oordeel, en hem verborgen uitvoeren gooide dat weg. `Invoke-Installer` kan nu stdout en stderr opvangen, en de regels van de bootstrapper worden herhaald als `bootstrapper:`-uitvoer. Geverifieerd met een proces dat een JSON-oordeel print en met een niet-nul-code afsluit; zonder de switch wordt er niets opgevangen en blijven er geen tijdelijke bestanden achter |
| Het opvangen van die uitvoer brak vervolgens elk oordeel, en alleen op de runtime die ertoe doet: onder Windows PowerShell 5.1 meldt een omgeleide `Start-Process -PassThru` helemaal geen exitcode, tenzij de proceshandle eerst wordt aangeraakt. Een `-x -m` die `{"success": true}` printte, werd gemeld als mislukt. Gemeten op beide runtimes; `$null = $proc.Handle` lost het op, en PowerShell 7 heeft het probleem nooit gehad |
| Succes wordt nu bepaald door het eigen JSON-oordeel van de bootstrapper als dat er is, niet door een exitcode die kan ontbreken. Niet-JSON, vreemde JSON en misvormde uitvoer vallen allemaal terug op de exitcode in plaats van te gokken |
| De preflight telde `PackageUserInformation`-entries en noemde een package "geïnstalleerd voor 1 gebruikersprofiel" terwijl de enige entry `S-1-5-18` was die het **stagede** — geen gebruiker, en niet geïnstalleerd. Hij leest nu de states: alleen-gestaged wordt als zodanig gemeld, en een package waarvan de entries `Installed(pending removal)` zeggen, wordt gemeld als al verwijderd en wachtend tot die gebruikers zich afmelden, bij naam, met de herstart gemarkeerd. Gemeten tegen de twee packages die een productie-session host echt rapporteerde |
| `Remove-AppxPackage -AllUsers` is alles of niets, dus één profiel dat hij niet kan aanraken laat de hele aanroep mislukken. Als dat gebeurt, wordt het package nu per gebruiker verwijderd, en wordt genoemd wie het nog vasthoudt — "een aangemelde gebruiker houdt het vast" is pas bruikbaar als je weet welke gebruiker. De SID komt uit de stringvorm van `PackageUserInformation`, die per build verschilt, getest tegen zeven vormen, waaronder de Entra-vorm `S-1-12-1`. **Niet getest** tegen een package dat verwijdering echt weigert |
| Het AppX-log beantwoordde toen de vraag die de `0x80070490` van de bootstrapper verborg: hij provisiont `MSTeams_26246...`, de build die voor één profiel geregistreerd is, en Windows kan de bestanden van dat package niet vinden. Twee MSTeams-versies naast elkaar, waarvan er één kapot is, is de toestand waar je naar moet zoeken |
| Een mislukte provisioning print nu ook de AppX-deploymentfouten die Windows heeft gelogd, die de reden bevatten die de `0x80070490` verbergt — "Unable to install because the following apps need to be closed &lt;package&gt;". Alleen-lezen, 10 ms, en stil als het log niets recents bevat |
| De preflight meldt een AppX-package waarvan de `Status` niet `Ok` is. Dat Windows een package als Modified of Tampered beschouwt, is precies wat `Remove-AppxPackage` `Catastrophic failure` laat antwoorden en de provisioning daarna laat mislukken, en tot nu toe was dat onzichtbaar |

### 2026-09-28
| Wijziging |
|--------|
| `scripts/Device/TempDisk/Init-TempDisk.ps1` toegevoegd: de tijdelijke (ephemeral) temp disk van een Azure VM wordt gewist bij elke deallocate, resize of hostverhuizing en komt terug als RAW, offline of zonder stationsletter. Windows leest de pagefile-configuratie bij het opstarten en leest die daarna nooit opnieuw, dus een pagefile die op `D:` is geconfigureerd terwijl die er bij het opstarten niet is, wordt gewoon nooit aangemaakt en de machine pagineert weer op `C:` - of draait helemaal zonder pagefile. Het script herstelt het volume als `D:` en laat de pagefile er weer naar wijzen |
| Een temp disk die alleen zijn stationsletter kwijt is, krijgt de letter terug in plaats van opnieuw geformatteerd te worden, herkend aan zijn label (`Temporary Storage`) of aan de `DataLoss_Warning_Readme.txt` die Azure op de resource disk schrijft. Alleen een RAW-schijf die geen boot- of systeemschijf is, wordt ooit geïnitialiseerd: een lege temp disk en een ongeformatteerde datadisk zien er van buitenaf identiek uit, dus een schijf met partities wordt gemeld en met rust gelaten, en bij meer dan één RAW-kandidaat weigert het script te gokken en vraagt het om `-DiskNumber`. `-Force` plus `-DiskNumber` is de enige route om een schijf te formatteren die nog data bevat |
| Een optisch station dat `D:` bezet, wordt eerst uit de weg gezet - Windows geeft `D:` aan de dvd-speler op een image zonder temp disk en geeft hem nooit terug, wat de tweede manier is waarop de pagefile op `C:` belandt |
| De run maakt onderscheid tussen de pagefile zoals *geconfigureerd* (register) en de pagefile *in gebruik* (deze sessie) en zegt welke welke is, in plaats van succes te melden voor een wijziging die pas bij de volgende herstart ingaat. Een pagefile configureren op een station dat niet kon worden hersteld, is een harde fout in plaats van een instelling die Windows stilletjes negeert |
| `scripts/Device/TempDisk/Register-InitTempDiskTask.ps1` toegevoegd, gebouwd op een concept met twee fouten: de standaardwaarde van `-ScriptSourcePath` verwees naar `$ScriptTargetDir`, een parameter die *erna* werd gedeclareerd, dus de standaardwaarde werd `\Init-TempDisk.ps1` en loste nooit op; en de kopie draaide met `-ErrorAction SilentlyContinue`, dus een ontbrekende bron registreerde een boot-taak tegen een bestand dat er niet is - die mislukt vervolgens bij elke boot zonder dat iemand meekijkt. De bron is nu standaard de kopie naast het script, een ontbrekende bron is een harde fout, en na het registreren wordt gecontroleerd dat de taak bestaat |
| Beide gedocumenteerd in een nieuwe `scripts/Device/TempDisk/readme.md`, de map toegevoegd aan de `Device/`-readme en de repository-boom, de root-readme een entry **Temp Disk & Pagefile (Azure / AVD)** gegeven, en `Init-TempDisk.ps1` in `menu.ps1` gehangen als Device-toets `V` (standaard `-CheckOnly`, tenzij je de reparatie bevestigt) |
| Geverifieerd: beide bestanden komen schoon door `Test-PowerShellSyntax.ps1`. De alleen-lezen-helpers zijn echt uitgevoerd op deze Windows 11-machine onder zowel PowerShell 5.1 als 7 met `Set-StrictMode -Version Latest` - zoeken naar een vrije letter, opzoeken van het optische station, detectie van de temp disk en RAW-kandidaten, en het uitlezen van de pagefile-status (die correct automatisch beheer aan en `C:\pagefile.sys` in gebruik meldde) - en de trigger-, principal- en settings-objecten van de geplande taak zijn opgebouwd en hun waarden gecontroleerd. **Nog niet geverifieerd op een live Azure VM of session host**: er is vanuit deze sessie geen schijf geïnitialiseerd, geen pagefile gewijzigd en geen taak geregistreerd |

### 2026-09-25 (20)
| Wijziging |
|--------|
| Het `Pivot toegang`-blad omgedraaid zodat het **site → groep → persoon** nest in plaats van site → persoon → groep. Zo verleent SharePoint daadwerkelijk toegang — een site heeft groepen, groepen hebben mensen — dus ingeklapt toont het de groepen op een site en uitgeklapt noemt het iedereen die ze binnenlaten |
| `Pivot per persoon` (persoon → site → groep) toegevoegd, zodat de andere richting nog steeds vanuit hetzelfde blad te beantwoorden is: waar kan deze ene persoon bij, en via wat. Dat is de offboardingvraag, en een pivot die bij de site begint kan die niet beantwoorden |
| Een direct gemachtigde persoon heeft geen groep, en in een hiërarchie van drie niveaus leest dat lege middelste niveau als ontbrekende data in plaats van als "verleend zonder groep". `ViaName` zegt nu `(direct toegekend)` voor die rijen in plaats van leeg te zijn |
| Lokaal geverifieerd met 240 controles over elf suites: beide pivots worden uit de werkmap teruggelezen met hun rijvelden in de bedoelde volgorde, en een directe toekenning wordt benoemd in plaats van leeg gelaten |

### 2026-09-25 (19)
| Wijziging |
|--------|
| De documentatie getoetst aan de repositoryregels in plaats van aan te nemen dat ze compleet was, en drie gaten gevonden. De parameters klopten: alle 19 parameters van `Get-SharePointPermissionsReport.ps1` staan in de comment-based help en in de parametertabel van de mapreadme, zonder verouderde vermeldingen in een van beide |
| `scripts/Reporting/readme.md` had helemaal geen `## Scripts`-tabel, terwijl `Exchange/`, `Entra/`, `Device/` en `SharePoint/` er allemaal een hebben. Toegevoegd, met alle vijf entries in de map — niet alleen het nieuwe script — zodat de tabel de map beschrijft in plaats van de laatste wijziging eraan. Van elke link en elk anker erin is gecontroleerd dat ze werken |
| De categorie `### 📊 Reporting` in de root-readme noemde alleen de Computer Last Logon- en Licensing-rapporten. Alle drie de SharePoint-rapportagescripts ontbraken, waaronder twee die van vóór dit werk dateren. Een entry toegevoegd voor `Get-SharePointPermissionsReport.ps1` en korte entries voor `Get-SharePointStorageReport.ps1` en `Remove-SharePointFileVersionsByDate.ps1` |
| De repository-boom beschreef het rechtenrapport nog als "naar CSV", wat niet meer klopte sinds `-Excel` is toegevoegd; er staat nu CSV + Excel. Het `menu.ps1`-label zei "wie heeft toegang tot wat, op elk niveau", wat de oude vorm van het rapport beschrijft in plaats van de geconsolideerde weergave per site waarmee het nu opent |
| Bevestigd dat er voor `f.ps1` niets hoeft te gebeuren: de index bouwt zichzelf opnieuw op wanneer de schrijftijd van een script verandert, dus `f-sharepointpermissionsreport -Excel` pikt nieuwe parameters op zonder `f-refresh` |

### 2026-09-25 (18)
| Wijziging |
|--------|
| `scripts/Reporting/Get-SharePointPermissionsReport.ps1` beantwoordde "welke toekenningen bestaan er", maar niet de vraag waarmee mensen het eigenlijk openen: **wie kan bij deze SharePoint, en hoe zijn ze daar gekomen.** `Rechten` zei dat een groep rechten had, `Groepen` zei wie erin zat, en niets verbond die twee — "Site Owners heeft Full Control" plus "Site Owners bevat vijf mensen" is geen antwoord. `SharePoint_Permissions_SiteAccess_<ts>.csv` toegevoegd (werkblad `Toegang`): één rij per persoon per site, met de groep waarlangs hun toegang loopt, de id van die groep en het machtigingsniveau |
| Bewust geconsolideerd per sitecollectie: iemand die in één site via dezelfde groep bij dertig mappen kan, is één rij, geen dertig. Een ander niveau of een andere groep is een aparte rij, want dat is andere toegang. Details per scope blijven achter `-IncludeEffectiveAccess` |
| Drie dingen vallen bewust niet uit die weergave weg: een direct gemachtigde persoon verschijnt als zichzelf met `ViaType = Direct`; `Everyone` en `Everyone except external users` lossen naar niemand op maar krijgen een rij die de claim noemt, omdat dat precies is waar een reviewer naar zoekt; en met `-SkipGroupExpansion` blijven de directe toekenningen zichtbaar, alleen de groepsleden ontbreken |
| `SiteTitle` toegevoegd — een geconsolideerde weergave van 130 sites is niet leesbaar als 130 URL's, en de titel van de root web is alleen bekend terwijl die web wordt gescand, dus die wordt daar vastgelegd en per rij opgezocht. `ViaId` toegevoegd naast `ViaName` na vergelijking met [NovaPoint](https://github.com/Barbarur/NovaPoint/wiki/Solution-Report-PermissionsReport), dat om dezelfde reden `GroupId` naast `AccessType` meeneemt: een titel als `Site Owners` komt op elke site in de tenant terug |
| Een `Pivot toegang`-blad toegevoegd dat site → persoon → groep nest tegen machtigingsniveau, gefilterd op extern en toegangstype. NovaPoint zet zijn gebruikers als lijst in één `Users`-kolom; hier krijgt elke gebruiker in plaats daarvan een eigen rij, wat minder compact leest maar het verschil is tussen wel of niet kunnen filteren of pivoteren op een persoon |
| Lokaal geverifieerd met 238 controles over elf suites (25 nieuw): een groepstoekenning toont haar mensen met de groepsnaam, id en het niveau; dezelfde toegang via dezelfde groep op een diepere scope wordt niet herhaald, een ander niveau wel; directe toekenningen, `Everyone`-claims en `-SkipGroupExpansion` gedragen zich allemaal zoals beschreven; checkpointsleutels zijn één per rij en uniek; en het blad en zijn pivot worden uit de werkmap teruggelezen. **Nog niet geverifieerd tegen een live tenant** |

### 2026-09-25 (17)
| Wijziging |
|--------|
| De `-Excel`-werkmap van `scripts/Reporting/Get-SharePointPermissionsReport.ps1` echt pivoteerbaar gemaakt. Eerst gecontroleerd in plaats van aangenomen: numerieke kolommen komen al als getallen in Excel aan, niet als tekst, dus aggregatie was nooit het probleem — het obstakel was `PermissionLevels`, dat SharePoint met meerdere niveaus tegelijk vult (`Read; Limited Access`). Een pivot behandelt elke combinatie als een eigen waarde, dus `Full Control` en `Full Control; Limited Access` komen op aparte rijen terecht |
| Bladen met `PermissionLevels` krijgen nu direct daarnaast een kolom `PrimaryPermission`, met het ene sterkste niveau van die toekenning. `Limited Access` verliest altijd van een echt niveau — SharePoint voegt het automatisch toe voor doorgang — en een aangepast niveau staat boven `Read` maar onder `Full Control`, omdat het bewust is aangemaakt en niet achter een ingebouwd niveau mag verdwijnen. Zowel Nederlandse als Engelse niveaunamen worden herkend, wat ertoe doet op een Nederlandstalige tenant |
| Drie kant-en-klare pivotbladen toegevoegd: `Pivot rechten` (site × machtigingsniveau, aantal toekenningen, gefilterd op principal en scopetype), `Pivot principals` (principal × scopetype, aantal scopes, gefilterd op site en extern) en `Pivot groepen` (groep × extern lid, aantal leden, gefilterd op site en groepstype). Elk verwijst alleen naar kolommen die het bronblad echt heeft, en een ontbrekende of versmalde bron wordt overgeslagen in plaats van een kapotte pivot op te leveren |
| Het aanmaken van pivots is best-effort en geïsoleerd: een fout geeft een waarschuwing en laat de databladen ongemoeid, volgens hetzelfde principe als dat de werkmap zelf de CSV's niet mag kosten |
| Lokaal geverifieerd met 213 controles over tien suites (33 nieuw): de niveaurangschikking over enkelvoudige, samengevoegde, omgekeerde, Nederlandse, aangepaste, hoofdlettervariërende en lege invoer; de afgeleide kolom die naast het origineel landt op de juiste bladen en niet op de andere; en de pivots die uit het package worden teruggelezen met de juiste rij-, kolom-, data- en filtervelden, waarbij ontbrekende bronnen en bladen zonder kolommen worden overgeslagen. **Nog niet geverifieerd tegen een live tenant** |

### 2026-09-25 (16)
| Wijziging |
|--------|
| `-Excel` toegevoegd aan `scripts/Reporting/Get-SharePointPermissionsReport.ps1`: één `.xlsx` naast de CSV's met een werkblad per rapport — `Samenvatting`, `Rechten`, `Groepen` en, met `-IncludeEffectiveAccess`, `Effectief` — elk een echte Excel-tabel met filterdropdowns en een vastgezette koprij, volgens hetzelfde `ImportExcel`-patroon als `Get-DistributionGroupMembers.ps1` |
| De CSV's worden nog steeds altijd geschreven en de werkmap wordt ervan opgebouwd, niet in plaats ervan. Daar streamt de scan naartoe en daar voegt een hervatte run aan toe, dus ze bestaan hoe dan ook — en een werkmap die niet geschreven kan worden (module ontbreekt, bestand open, geheugen vol) kost dan een gemakskopie in plaats van het rapport |
| Een werkblad stopt bij 1.048.576 rijen en laat de rest zonder klagen vallen, dus bladen worden afgekapt op 1.000.000 met een waarschuwing die het blad noemt en de CSV die nog alles bevat. Op een grote tenant komt realistisch gezien alleen `Effectief` daarbij in de buurt |
| Bij het inbouwen hiervan een echt gat in het groepslidmaatschap gedicht: een Entra ID-groep die **direct** op een site, lijst of item is gemachtigd, komt nooit langs `/sitegroups`, dus dat was de enige soort groep waarvan het rapport het lidmaatschap nooit toonde — alleen de eerste tien namen in `MemberPreview`. Die groepen krijgen nu hun eigen rijen in de Groups-uitvoer, opgelost naar personen, eenmaal per groep vastgelegd in plaats van eenmaal per toekenning, en met hetzelfde schema dat de SharePoint-groepsrijen al gebruiken |
| De Excel-vraag toegevoegd aan de `menu.ps1`-entry |
| Lokaal geverifieerd met 180 controles over negen suites (20 nieuw): een werkmap wordt geschreven en teruggelezen met alle vier bladen in volgorde en hun rijen intact, een herhaalde run vervangt in plaats van toe te voegen, afwezige/lege/ontbrekende bronnen worden overgeslagen, niets te schrijven laat geen bestand achter, een te groot blad wordt afgekapt in plaats van door Excel afgekort, een direct gemachtigde Entra-groep wordt eenmaal getoond met haar echte leden, SharePoint-groepen worden aan `/sitegroups` overgelaten, en beide groepsbronnen delen één schema. **Nog niet geverifieerd tegen een live tenant** |

### 2026-09-25 (15)
| Wijziging |
|--------|
| Eerste volledige live run van `scripts/Reporting/Get-SharePointPermissionsReport.ps1`: 130 webs, 2066 lijsten, 453 unieke scopes, 1554 toekenningen, 83 deellinks, 13 externe toekenningen, 111 `Everyone`-toekenningen. Drie fouten in de uitvoer zelf, allemaal gevonden door de geproduceerde CSV's te lezen in plaats van de logs |
| `-IncludeEffectiveAccess` leverde een leeg bestand op een tenant met 1554 toekenningen en meer dan duizend opgeloste leden. De guard was `if ($IncludeEffectiveAccess -and $EffectiveRows)`, en **een lege `List[object]` is falsy in PowerShell** — dus de test faalde al bij de allereerste rij en de lijst kon nooit gevuld worden, waardoor hij leeg bleef, waardoor de test bleef falen. Nu een expliciete `$null -ne`-controle |
| 121 van de 158 rijen "kon niet worden gelezen" waren één enkele verborgen systeemlijst, `Lijst met gebruikersgegevens` (template 112, de User Information List), op elke site. SharePoint weigert `/items` daarop met `400` bij elke `$select`-breedte, inclusief de smalste trede van de ladder. De items daarin zijn directoryrecords in plaats van inhoud, dus scopes op itemniveau betekenen daar niets voor een toegangsreview — de item-sweep slaat template 112 nu over en meldt dat, terwijl de eigen scope van de lijst nog wel wordt gerapporteerd |
| De overige 37 waren verouderd: foutrijen geschreven door de eerder onderbroken poging, door de hervatting meegenomen in de uiteindelijke CSV, hoewel die lijsten bij de nieuwe poging slaagden. Een mislukte unit wordt bewust ongemarkeerd gelaten zodat hij opnieuw wordt geprobeerd, maar niets verwijderde de oude rijen ervan. Detailrijen dragen nu de `UnitKey` die ze heeft geproduceerd, en een rij waarvan de unit als voltooid is gemarkeerd, wordt weggelaten wanneer de uiteindelijke CSV wordt geschreven — zodat het aantal onvolledigheden het bestand beschrijft dat de lezer opent |
| De samenvatting wordt nu opgebouwd uit de gepubliceerde detail-CSV in plaats van uit de gedeeltelijke, zodat de aantallen en het bestand overeenkomen |
| Lokaal geverifieerd met 158 controles over acht suites (26 nieuw): effectieve rijen worden één per opgeloste gebruiker uitgegeven met de groep waarlangs ze kwamen, een null-lijst wordt verdragen, achterhaalde foutrijen worden weggelaten terwijl nog steeds falende en sleutelloze behouden blijven, verder gaat er niets verloren, en de samenvatting telt alleen wat is gepubliceerd. Twee eerdere asserties bleken verkeerd tussen haakjes te staan — een ervan een valse pass — en zijn gecorrigeerd. **De fixes voor de bevindingen van deze run zijn zelf nog niet geverifieerd tegen een live tenant** |

### 2026-09-25 (14)
| Wijziging |
|--------|
| De tweede live run van `scripts/Reporting/Get-SharePointPermissionsReport.ps1` authenticeerde schoon — de controle op tokenrollen ving de replicatievertraging bij de eerste poging op en wachtte die uit, SharePoint en Graph accepteerden allebei hun tokens, en 130 webs werden ontdekt en gingen scannen. Daarna herhaalden zich twee fouten op scanniveau op elke site |
| `ConvertTo-PermissionRows` weigerde een lege collectie role assignments: een `Mandatory [object[]]`-parameter weigert `@()`, dus elke systeemlijst met unieke rechten maar zonder resterende role assignments (`User Information List`, `Converted Forms`, `Bibliotheek met onderhoudslogboeken`) faalde met `Cannot bind argument to parameter 'RoleAssignments'`. Opgelost met `[AllowEmptyCollection()]` — een scope zonder assignments levert terecht geen rijen op |
| Belangrijker nog: een onleesbare lijst role assignments was niet te onderscheiden van een lege. `Invoke-SPGet` slikt `403`/`404` in en geeft `$null` terug, wat `Get-SPCollection` omzet in een lege collectie — en een lege collectie leest als "niemand heeft rechten op deze scope". Het lezen van role assignments gebruikt nu `-ThrowOnDenied`, zodat een weigering een foutrij wordt die zegt dat de rechten onbekend zijn, in plaats van een stille bewering dat er geen zijn. Dit is de enige plek waar een 403 niet wordt overgeslagen, omdat het de enige plek is waar "niet mogen kijken" als bevinding zou worden misgelezen |
| De gallerylijsten (`Galerie van thema's`, `Galerie met basispagina's`) antwoordden `400 Bad Request` op de item-`$select`, omdat hun schema niet elk veld bevat dat erin wordt genoemd, en een 400 is niet iets wat opnieuw proberen oplost. De item-sweep stapt nu een ladder van vier treden af met steeds smallere `$select`-clausules totdat SharePoint er een accepteert; elke trede houdt `Id` en `HasUniqueRoleAssignments`, dus in het slechtste geval gaat een bestandsnaam verloren in plaats van de unieke scopes van de lijst. Alleen een 400 leidt tot versmallen — een weigering, een throttle of een view threshold antwoordt hetzelfde, hoe weinig velden er ook worden gevraagd |
| Lokaal geverifieerd met 132 controles over zeven suites (21 nieuw): een lege set assignments crasht niet meer en levert geen rijen op, een geweigerde leesactie gooit een fout met "unknown rather than empty", een echte lege `200` blijft van begin tot eind leeg, een gewone sweep slaat een 403 nog steeds over, elke trede van de ladder houdt de velden waar de scan van afhangt, en alleen een 400 versmalt. **De scanfase voorbij web 11 is nog steeds niet geverifieerd tegen een live tenant** |

### 2026-09-25 (13)
| Wijziging |
|--------|
| De eerste live run van `scripts/Reporting/Get-SharePointPermissionsReport.ps1` kwam langs SharePoint — de certificaatcredential werkte en de preflight meldde `SharePoint accepted the token (root web: ...)` — en faalde daarna op Graph met `401` bij het ophalen van sites. Oorzaak: het Graph-token werd direct na het toekennen van de app-rollen aangemaakt, voordat de toekenning was gerepliceerd, dus het had helemaal geen `roles`-claim. Graph beantwoordt zo'n token met `401`, niet `403`, en omdat het token zijn volle uur in de cache stond, kreeg elk van de zes nieuwe pogingen hetzelfde dode token terug |
| Tokens moeten zich nu bewijzen: `Get-ResourceToken` neemt `-RequiredRoles`, decodeert de uitgegeven JWT en weigert een token te cachen waarvan de `roles`-claim mist wat de run nodig heeft. Hij blijft opnieuw aanmaken (tot 15 pogingen, backoff begrensd op 20s) totdat de toekenning verschijnt, en faalt dan met de ontbrekende rol bij naam. Beide tokens worden vooraf gevalideerd — Graph op `Sites.Read.All` + `GroupMember.Read.All`, SharePoint op `Sites.FullControl.All` — zodat een replicatievertraging wordt uitgewacht voordat de scan begint, in plaats van pas na 130 sites te worden ontdekt |
| `Invoke-GraphGet` gooit nu bij een `401` het gecachte token weg en maakt eenmalig een nieuw aan, hetzelfde herstel dat `Invoke-SPGet` al had. Alleen het verzoek opnieuw proberen had nooit kunnen werken tegen een vergiftigde cache-entry |
| Een door de gebruiker opgegeven `-ClientId`-app wordt bewust **niet** op rollen gevalideerd: een werkende app kan bredere rollen hebben (`Directory.Read.All` in plaats van `GroupMember.Read.All`), en die afwijzen zou een valse fout zijn. De SharePoint-preflight vangt een app met werkelijk te weinig rechten nog steeds af |
| `Get-JwtClaim` geeft `$null` terug voor een leeg token in plaats van een parameter binding error te gooien, en `Disconnect-MgGraph` lekt zijn contextobject niet meer als losse `ClientId`/`TenantId`/`Scopes`-tabel na de samenvatting |
| Lokaal geverifieerd met 111 controles over zes suites (23 daarvan nieuw, die de echte `Get-ResourceToken` tegen een nep-tokenendpoint aansturen): een rol die laat aankomt wordt uitgewacht en alleen het token dat hem draagt wordt gecachet, een rol die nooit aankomt faalt luid zonder dat er iets wordt gecachet, de backoff groeit en blijft begrensd, en de caching blijft per resource geïsoleerd. **Het gecorrigeerde Graph-pad is nog niet geverifieerd tegen een live tenant** |

### 2026-09-25 (12)
| Wijziging |
|--------|
| `scripts/Reporting/Get-SharePointPermissionsReport.ps1` robuuster gemaakt voor lange tenant-brede runs. Een `401` kon nog steeds eindigen als foutregel per site: de handler per lijst gooide hem opnieuw op, maar de handler per web ving hem weer af, dus een credential die halverwege de run ophield te werken zou één foutregel per resterende site hebben geschreven — precies het faalpatroon dat de certificaatfix net had weggenomen. Beide handlers laten een 401 nu door en de run stopt |
| De parallelle item-lookups lazen het bearer-token één keer per bibliotheek in plaats van één keer per golf. Een bibliotheek met genoeg unieke scopes leeft langer dan een token, dus het staartstuk ervan zou zijn mislukt zonder enige aanwijzing waarom. Het token wordt nu vóór elke golf opnieuw gelezen |
| Het opsommen van items materialiseert niet langer een complete bibliotheek vóór het filteren. `Invoke-SPCollectionPaged` geeft elke pagina aan een callback en alleen items die echt een eigen scope hebben worden bewaard — een bibliotheek met een miljoen items kost nu één pagina geheugen in plaats van een miljoen levende objecten |
| Een paging-beveiliging toegevoegd: als SharePoint een identieke `nextLink` terugstuurde, was dat een oneindige lus tegen een live tenant; dat wordt nu gedetecteerd en gestopt |
| Checkpoint-sleutels zijn verhuisd van het JSON-statusbestand naar een append-only `.keys.partial.log`. Na elke lijst een gesorteerde lijst van alle voltooide sleutels herschrijven is kwadratisch; op een tenant met duizenden lijsten kostte het checkpointen meer dan het scannen. Een afgebroken laatste regel van een gekild proces wordt getolereerd — die eenheid wordt gewoon opnieuw gescand |
| Een mislukte CSV-schrijfactie (het deelbestand dat open staat in Excel) wordt vijf keer opnieuw geprobeerd en stopt daarna de run. Voorheen werd er een fout gegooid terwijl de eenheid al als voltooid was gemarkeerd, waardoor die regels definitief uit het rapport verdwenen |
| Een lijst die mislukt kost nu alleen die lijst, niet de rest van de site: foutafhandeling per lijst schrijft een foutregel, behoudt wat de lijst al had opgeleverd en laat de eenheid bewust ongemarkeerd zodat een hervatte run hem opnieuw probeert. Een mislukte item-sweep krijgt een eigen regel, want zonder die regel lijkt het alsof de lijst gewoon niets met unieke rechten had |
| Workers per item melden `401`/`403` niet langer als een lege set rechten — alleen `404` (item echt verwijderd tijdens de scan) betekent "geen rechten". Beweren dat een onleesbaar item geen rechten heeft is erger dan dat gewoon zeggen |
| Een schrijftest op de uitvoermap toegevoegd vóór het authenticeren, een afbreking wanneer discovery helemaal geen sites vindt, en een `trap` die de tijdelijke Full Control app-registratie opruimt bij elke onafgehandelde fout |
| De contexttabel van `Set-MgRequestContext` die naar stdout lekte onderdrukt, en de run sluit nu af met de vermelding of elke beoogde scope is gelezen of hoeveel er zijn gemist |
| Lokaal geverifieerd met 53 controles verdeeld over vier suites: parsing van principals/claims, consistentie van het CSV-regelschema (zes regelvormen, elk 26 kolommen), generatie van het certificaat en de ondertekende assertion, en HTTP-gedrag aangestuurd via een nep-transport — 401 breekt af na één herauthenticatie, 403/404 blijven per object, 429 wordt herhaald tot succes, paging-lussen worden doorbroken, vergrendelde bestanden worden opnieuw geprobeerd en het checkpoint-log overleeft een afgebroken regel. **Nog steeds niet geverifieerd op een live tenant** |

### 2026-09-25 (11)
| Wijziging |
|--------|
| Opgelost dat `scripts/Reporting/Get-SharePointPermissionsReport.ps1` een leeg rapport opleverde tegen een live tenant: alle 130 webs kwamen terug met `[SKIP] Web not accessible with the current permissions`. De tijdelijke App Registration authenticeerde met een client secret, en **SharePoint Online weigert elk app-only token dat met een secret is verkregen** — `401` met `x-ms-diagnostics: ... Unsupported app only token`. Graph accepteerde dezelfde credential, dus het opsommen van sites werkte en alleen de `_api`-aanroepen faalden, en daarom leek het een rechtenprobleem per site |
| De tijdelijke app krijgt nu een certificaat in plaats van een secret. Het wordt in het geheugen gegenereerd met `CertificateRequest`, geregistreerd als `keyCredential` en gebruikt om een RFC 7523 client assertion te ondertekenen — het komt nooit in de certificaatopslag of op schijf, dus een onderbroken run laat niets achter |
| `Invoke-SPGet` slikt `401` niet langer in samen met `403`/`404`. Een `401` is nooit per site — het is hetzelfde antwoord voor de hele tenant — en hem behandelen als "deze ene site is niet toegankelijk" is precies wat één credential-fout veranderde in 130 regels die lazen als bevindingen. Er wordt nu een fout gegooid, met de reden uit `x-ms-diagnostics` en, bij het secret-geval, wat je eraan moet doen |
| Een SharePoint-preflight toegevoegd: één aanroep tegen de tenant-root na het verbinden, voordat er iets wordt opgesomd. Of SharePoint de credential accepteert is één ja/nee voor de hele run, dus dat uitzoeken kost nu één request in plaats van een volledige sweep |
| `-ClientSecret` waarschuwt nu bij het opstarten dat het SharePoint-deel zal mislukken, en de documentatie zegt hetzelfde. Alternatieven met alleen Graph zijn overwogen en verworpen: Graph heeft geen endpoint voor rol-toewijzingen op webs, SharePoint-groepen, sitecollectiebeheerders of benoemde machtigingsniveaus, en zou één `/permissions`-aanroep per item nodig hebben in plaats van één `HasUniqueRoleAssignments`-sweep per lijst |
| Een verdwaalde `ClientTimeout RetryDelay MaxRetry`-tabel onderdrukt die `Set-MgRequestContext` aan het eind van elke run naar stdout schreef |
| Lokaal geverifieerd: 21 controles op de certificaatgeneratie en de ondertekende assertion, onder meer dat de handtekening klopt tegen de publieke sleutel van het certificaat, dat `x5t` overeenkomt met de SHA-1-hash ervan, en dat er niets naar `Cert:\CurrentUser\My` wordt geschreven. Het gecorrigeerde auth-pad zelf is **nog niet geverifieerd op een live tenant** |

### 2026-09-25 (10)
| Wijziging |
|--------|
| `Convert-MarkdownToHtml.ps1` gaf elke genummerde lijst weer als lege opsommingstekens. `$Matches` is één variabele per scope: de lijst-tak ving de itemtekst af en voerde daarna een tweede `-match` uit om te bepalen of de lijst genummerd was, en die tweede match gooide de vangst weg. Lijsten met streepjes bleven alleen heel omdat hun tweede match mislukte en `$Matches` met rust liet. Beide vangsten komen nu uit één match en het teken bepaalt het type zonder opnieuw te matchen |
| Een fenced code block dat was ingesprongen om binnen een genummerde stap uit te lijnen, hield die inspringing, dus wie het commando uit de pagina kopieerde, kopieerde de voorloopspaties mee. De eigen inspringing van de fence wordt nu van de inhoud gestript — en alleen die: een blok dat op kolom 0 is gefenced houdt elke spatie, en dat is wat de voorbeelduitvoer van het script nodig heeft |
| Geverifieerd op de opnieuw gegenereerde pagina: 36 lijstitems en **geen enkele** leeg, parseert nog steeds als XML, 29 tabellen en 19 code blocks intact, het ingesprongen commando komt er schoon uit, en de zeven code blocks die terecht met witruimte beginnen doen dat nog steeds |

### 2026-09-25 (9)
| Wijziging |
|--------|
| Je kunt nu vanuit een readme direct doorklikken naar het script dat hij beschrijft. Elke scriptnaam in de `Scripts`-tabel van een folder-readme linkte naar een sectie verderop op dezelfde pagina, nooit naar het bestand — dus de readme vertelde je wat een script deed, maar gaf je geen manier om het te openen. 172 links in 47 readmes wijzen nu naar het bestand, met een `([docs](#…))`-link ernaast voor de sectie die er eerst stond |
| 34 daarvan waren helemaal geen links: een scriptnaam in een tabelcel, tussen backticks, zonder iets erachter. Dat zijn nu ook bestandslinks |
| `scripts/Startup/Test-MarkdownLinks.ps1` toegevoegd, want links die nooit worden gecontroleerd zijn links die ongemerkt verrotten. Het loopt elke `.md` langs en faalt op twee dingen: een relatieve link naar een bestand dat er niet is (percent-encoded spaties eerst gedecodeerd, zoals GitHub ze serveert), en een anchor zonder kop erachter. Anchors worden opgelost zoals GitHub ze opbouwt, inclusief het `-1`/`-2`-achtervoegsel voor herhaalde koppen |
| Het vond drie anchors die nooit hadden gewerkt: `#watch-rdslivesps1` had een `s` te veel voor `### Watch-RDSLive.ps1`, en twee links in de Intune-readme gebruikten `#detect--remediate-…` waar de kop `### Detect- / Remediate-StuckWin32AppEnforcement.ps1` `#detect---remediate-…` oplevert — drie streepjes, omdat de slash verdwijnt en de spaties eromheen elk één streepje worden. Dat vindt niemand met het blote oog |
| Onzichtbare tekens worden zowel uit de kop als uit de link gestript voordat ze worden vergeleken. Zonder dat lijken de vier emoji-items in de inhoudsopgave van de root kapot: de kop en de link bevatten allebei een variation selector, en dat is geen letter en geen cijfer. GitHub's slugger byte voor byte nabootsen op tekens die niemand kan zien is niet het doel — vaststellen dat een link en een kop bij elkaar horen wel |
| Geverifieerd: 854 interne links in 71 markdown-bestanden worden allemaal opgelost, alle 181 bestanden parseren, en de scriptindex is actueel met 178 scripts. `L` toegevoegd aan het menu voor de linkcontrole, naast `X` voor de index |

### 2026-09-25 (8)
| Wijziging |
|--------|
| De Teams-procedure voor IT Glue bestaat nu ook als opgemaakte HTML-pagina, `scripts/Device/Update-TeamsClient-ITGlue.html`, om in IT Glue te plakken of af te drukken. Hij wordt **gegenereerd** door het nieuwe [`scripts/Startup/Convert-MarkdownToHtml.ps1`](scripts/Startup/Convert-MarkdownToHtml.ps1) in plaats van met de hand geschreven: dat document is in twee dagen zes keer gewijzigd, en een handgemaakte kopie zou de volgende ochtend al fout zijn geweest |
| De converter dekt wat deze documenten daadwerkelijk gebruiken — koppen, tabellen, fenced code blocks inclusief de ingesprongen exemplaren binnen genummerde stappen, blockquotes, beide soorten lijsten, horizontale lijnen, en inline code, vet, cursief en links — en geeft al het andere als tekst door in plaats van te gokken. `-Check` schrijft niets en eindigt met `1` wanneer de gecommitte pagina achterloopt op zijn markdown, en dat is wat een hook of pipeline zou aanroepen |
| Void-elementen worden zelfsluitend uitgevoerd, zodat de pagina zowel als XML als als HTML parseert. Zo is het ook geverifieerd, in plaats van door ernaar te kijken: de uitvoer parseert, en bevat 29 tabellen, 190 rijen, 19 code blocks en 46 koppen, met **geen enkele** tabel die een rij bevat die afwijkt van de breedte van de kop. Ook gecontroleerd: nergens in de weergegeven tekst is nog een `**`, backtick of `](` over, en de ✅/❌ en tekens met accenten overleven |
| Beide faalpaden van `-Check` doorlopen: een pagina die nog niet bestaat, en een markdown-bestand dat verder is gegaan — beide eindigen met `1` en de reden. De regel met de generatiedatum wordt buiten de vergelijking gehouden, zodat een ongewijzigd document geen verschil meldt |
| Toegevoegd als menutoets **M**, gedocumenteerd in [`scripts/Startup/readme.md`](scripts/Startup/readme.nl.md), en `scripts/INDEX.md` opnieuw gegenereerd — het toevoegen van een script had hem verouderd gemaakt, wat `Update-ScriptIndex.ps1 -Check` meldde |

### 2026-09-25 (7)
| Wijziging |
|--------|
| De leesbare delen van een oudere IT Glue-versie van dezelfde procedure samengevoegd in `Update-TeamsClient-ITGlue.md`: de beschrijving in één zin van wat de procedure dekt, en de ✅/❌-vorm voor "past dit script bij dit ticket", inclusief de twee gebruikersklachten die die versie noemde en de onze niet — Teams die blijft hangen bij het opstarten en Teams die onverwacht afsluit |
| De twee versies waren het oneens over wie de update mag uitvoeren — de oudere legt het op niveau 1, de onze op niveau 2 — en daar heeft een servicedesk minder aan dan aan elk van beide antwoorden afzonderlijk. Het algemene niveau is vervangen door een "wie mag wat"-tabel die per actie een niveau toekent: controleren blijft niveau 1, de update en de reparaties zitten op niveau 2, en de drie switches die de handtekeningcontrole overslaan of de Windows Installer-database bewerken zitten op niveau 3. Het escalatiebeleid wijzigen is nu één tabel, niet het document opnieuw doorlezen |
| Bewust niet samengevoegd, gemeten tegen het script in plaats van op het oog beoordeeld: de parametertabel van die versie dekt 8 van de 16 parameters, stelt dat classic Teams nooit wordt aangeraakt (`-RemoveClassicTeams` doet precies dat), en toont twee voorbeeldregels uitvoer die het script niet produceert — `Found MSTeams ...` en `Teams installation completed` |
| Nummeringscorrectie: twee items hadden allebei het label `(4)`. De IT Glue-synchronisatie is nieuwer dan het scriptindex-item erboven, dus die is nu `(6)` en staat in de volgorde waarin het werk is gedaan |

### 2026-09-25 (6)
| Wijziging |
|--------|
| Het IT Glue-document is weer gelijkgetrokken met het script, gecontroleerd door de tekst te vergelijken met het parameterblok in plaats van door hem te lezen: `-BootstrapperUrl` en `-SkipSignatureCheck` ontbraken helemaal, en in de NinjaOne-variabelentabel ontbraken `removeWebRtcRedirector`, `clearOrphanedAddInRegistration`, `skipSignatureCheck` en `webRtcUrl`. Alle drie de documenten dekken nu alle 16 parameters, en de variabelentabel alle 16 omgevingsvariabelen |
| Diezelfde vergelijking vond een echt gat in het script: elk ander tekstveld kon via een NinjaOne-variabele worden ingesteld, behalve `bootstrapperUrl`, die simpelweg nooit werd gelezen. Een beheerder die hem instelde, zou hem stilzwijgend genegeerd zien worden. Hij wordt nu gelezen, samen met `webRtcUrl` |
| Twee uitspraken in het IT Glue-document klopten niet meer. "Het raakt classic Teams niet aan" klopt alleen zonder `-RemoveClassicTeams`, en de samenvatting in gewone taal beloofde nog dat bij een update elke kopie van de add-in wordt verwijderd - die sweep is nu afhankelijk van of er een vervanging kan worden geïnstalleerd |
| Een niveau 3-stappenplan toegevoegd voor de `1612` + `1638`-deadlock waar de productiehost tegenaan liep: wat elke code betekent, het alleen-lezen commando dat vertelt of Windows Installer zijn gecachte MSI nog heeft, welke van de twee uitkomsten `-ClearOrphanedAddInRegistration` nodig heeft, en de opmerking dat Teams starten en Outlook herstarten een gebruiker in de tussentijd de vergaderknop teruggeeft |

### 2026-09-25 (5)
| Wijziging |
|--------|
| Een script vinden op GitHub betekende raden onder welke van de 56 workload-mappen het stond en readmes openen tot het opdook. Er is nu één pagina die dat beantwoordt: [`scripts/INDEX.md`](scripts/INDEX.md) somt alle 176 scripts van A tot Z op met een link naar het bestand, een link naar de readme van de map, en wat het doet — Ctrl-F in plaats van zoeken |
| De pagina wordt **gegenereerd**, door het nieuwe `scripts/Startup/Update-ScriptIndex.ps1`, zodat hij niet kan afdrijven van de bestanden zoals een met de hand bijgehouden tabel dat doet. `-Check` meldt een verouderde index zonder te schrijven (exit `1`), en dat is wat een hook of pipeline zou aanroepen; een run die de pagina actueel aantreft schrijft helemaal niets |
| Beschrijvingen komen uit de scripts zelf: het `.SYNOPSIS`-blok, samengevoegd over de regels waarover het doorloopt in plaats van alleen de eerste regel te nemen, wat halve zinnen opleverde zoals "Grant Full Access and/or Send As delegate rights on one mailbox, a CSV list of" in de tabel. Waar een synopsis opent met een zin en daarna zijn gevallen opsomt, blijft de inleiding behouden en wordt de lijst er niet achteraan meegesleept |
| Voor de oudere scripts zonder `.SYNOPSIS` wordt in plaats daarvan een `#`-commentaarblok bovenaan gebruikt — maar alleen een echte header. Een enkele commentaarregel direct boven code beschrijft die regel, niet het script: `# URL van de theme` boven een `$ThemeUrl`-toewijzing werd als beschrijving gelezen, en dat is in een tabel erger dan een lege cel |
| De acht scripts die nog steeds niets hadden kregen een echte `.SYNOPSIS` in plaats van een lege cel: `add-lock.ps1`, `add-shortcut-lock.ps1`, `logic-permissies.ps1`, `Test-OpenVpnDiagnostics.ps1`, `Deploy-OfficeTheme.ps1`, `Restart-Time-Sync.ps1`, `Test-PowerShellSyntax.ps1` en `functies.ps1`. Alle 176 scripts beschrijven zichzelf nu, dus de index heeft geen sectie "zonder beschrijving" meer |
| De twee scripts gedocumenteerd die in geen enkele readme werden genoemd: `Phising-rollout.ps1` in [`scripts/Entra/readme.md`](scripts/Entra/readme.nl.md) (de tweerichtingssynchronisatie tussen de uitrol van phishing-bestendige MFA en de groepen met registraties, wat als geregistreerd telt en waarom de standaard een AAGUID-filter is) en `Get-FSlogix-errors.ps1` in [`scripts/RDS/readme.md`](scripts/RDS/readme.nl.md) (wat de FSLogix-diagnose verzamelt en dat die op de sessiehost moet draaien). De header wees nog naar een bestandsnaam die niet meer bestaat en noemde een echte klant in het voorbeeld; beide gecorrigeerd |
| De `Menu`-tabel in de root was uit de pas gaan lopen met `menu.ps1` — `I`, `T`, `S` en `P` ontbraken. Gesynchroniseerd, en `X` toegevoegd voor de indexgenerator, die ook in de Startup-readme en de repository-boom staat |
| Geverifieerd: alle 176 bestanden parseren; de generator is idempotent (een tweede run meldt "already up to date" en schrijft niets); `-Check` eindigt met `0` wanneer de index actueel is; elke markdown-link in de repository wordt opgelost, percent-encoded mapnamen inbegrepen; en `f.ps1` vindt zowel het nieuwe script als de nieuw beschreven scripts nog steeds |
| Nummeringscorrectie: twee items hieronder hadden allebei het label `(3)`. Hernummerd naar de volgorde waarin het werk werkelijk is gedaan |

### 2026-09-25 (4)
| Wijziging |
|--------|
| Correctie op het vorige item: de `1638` op de vergader-add-in werd **niet** veroorzaakt doordat `-Force` de client downgradede. Gemeten op de host zelf was de geregistreerde add-in `1.25.28902` en de MSI die werd geïnstalleerd `1.26.21803` - nieuwer, en toch geweigerd. Deze MSI weigert te installeren zolang er een andere kopie van de add-in geregistreerd is, welke versie dat ook is. De versievergelijking die in de vorige wijziging is toegevoegd zou de fout dus niet hebben voorkomen; de beveiliging vraagt nu of er een registratie de uninstall heeft overleefd, en dat is wat het werkelijk bepaalt |
| `-ClearOrphanedAddInRegistration` is de uitweg uit de toestand waarin die host verkeert. Een uninstall die `1612` antwoordt betekent dat Windows Installer de gecachte MSI kwijt is die het nodig heeft en het product met geen enkel ondersteund middel meer kan verwijderen, terwijl de registratie elke herinstallatie blijft weigeren. De switch laat de installer dat ene product vergeten: de sleutels onder `Installer\Products`, `Installer\Features` en `Installer\UserData\S-1-5-18\Products`, de vermelding onder de upgrade code, en de vermelding in Programma's en onderdelen. Wat MsiZap vroeger deed, beperkt tot één product, pas nadat msiexec heeft bewezen dat het niet kan, standaard uit, en elke sleutel via `ShouldProcess` |
| Om die sleutels te vinden is de ProductCode nodig als Windows Installer's "packed" GUID van 32 tekens. Die transformatie is gevalideerd voordat er iets mee naar sleutels werd gewezen om te verwijderen: van 57 uninstall-vermeldingen met een GUID als naam op een werkstation kwamen de 32 met machinebrede productgegevens allemaal uit op een bestaande packed sleutel met een identieke `DisplayName`, en de 25 die dat niet deden zijn installaties per gebruiker onder de eigen SID van de gebruiker |
| Alleen-lezen geverifieerd tegen drie echte producten: elk levert drie productsleutels op plus precies één upgrade code-vermelding, en het opgebouwde pad is leesbaar zoals het er staat. Een ongeldige productcode geeft niets terug en een onbekend product geeft geen sleutels, dus de opschoning kan niet in het wilde weg afgaan. **Niet getest:** de verwijdering zelf, en de herinstallatie die erop zou moeten volgen |

### 2026-09-25 (3)
| Wijziging |
|--------|
| `scripts/Reporting/Get-SharePointPermissionsReport.ps1` toegevoegd — een uitputtend, alleen-lezen rechtenrapport voor SharePoint Online: sitecollectiebeheerders, rol-toewijzingen op webs inclusief onderbrekingen van overerving, SharePoint-groepen met hun volledige lidmaatschap, rol-toewijzingen op lijsten en bibliotheken, elke map en elk item met een unieke scope, deellinks met hun soort, externe/gast-principals, `Everyone`-toekenningen, en toekenningen aan Entra-groepen uitgewerkt tot transitief lidmaatschap. Vier CSV's: detail, samenvatting per site, groepslidmaatschap, en — achter `-IncludeEffectiveAccess` — één regel per uitgewerkte gebruiker per scope met de groep waarlangs de toegang loopt |
| Overerving wordt gevolgd zoals SharePoint die modelleert: een item wordt alleen als eigen scope gerapporteerd wanneer `HasUniqueRoleAssignments` true is, dus de CSV is een kaart van de rechtenstructuur in plaats van een regel per bestand. Site-discovery is bewust redundant — Graph `getAllSites`, daarna subsites via zowel Graph als SharePoint REST (`/_api/web/webs`), ontdubbeld op URL — omdat Graph klassieke subwebs weglaat |
| Authenticatie moest app-only worden: rol-toewijzingen zijn via Graph helemaal niet leesbaar, en vallen ook niet onder de Read/Write/Manage-applicatierollen van SharePoint — alleen `Sites.FullControl.All` kan ze opsommen. Het script meldt zich één keer interactief aan, maakt een kortlevende App Registration met die rol plus Graph `Sites.Read.All` en `GroupMember.Read.All`, en verwijdert die bij het afsluiten weer. Ondanks de Full Control-rol doet het alleen `GET`: het schrijft nooit en wijzigt nooit een recht. `-ClientId`/`-TenantId` met een secret of certificaat slaat de tijdelijke app over |
| Hervatbaar zoals de andere lange SharePoint-scans: een checkpoint per voltooide lijst, met als sleutel een hash van de scanparameters, zodat een onderbroken tenant-run verdergaat in plaats van opnieuw te beginnen; `-Restart` gooit het weg. Checkpoint-bestanden worden pas opgeruimd zodra de definitieve CSV's zijn geschreven, dus hun aanwezigheid is zelf het signaal dat een run is onderbroken |
| Gedocumenteerd in `scripts/Reporting/readme.md` (dekking, authenticatie, de vier uitvoerbestanden, checkpoints, volledige parametertabel, voorbeelden), toegevoegd aan de repository-boom in de root en aan `menu.ps1` onder Reporting als `P` — dat ook vraagt naar tenant of één site, scope, en of de effective-access-CSV moet worden geschreven |
| De repository-boom vermeldde ook dit script niet, en `Remove-SharePointFileVersionsByDate.ps1` evenmin; beide staan er nu in |
| **Niet door mij getest**: dit script is in deze sessie niet tegen een live tenant uitgevoerd — alleen de syntax is gecontroleerd. Het pad met de tijdelijke App Registration, de throttling-retries en het hervatten via checkpoints zijn hier niet geverifieerd en moeten op een pilot-tenant worden doorlopen, te beginnen met `-SiteUrl` en `-Scope Site`, vóór een tenant-brede run |

### 2026-09-25 (2)
| Wijziging |
|--------|
| Een add-in die geregistreerd is maar waarvan de bestanden weg zijn, telt niet langer als geïnstalleerd. Dat is precies de toestand die de mislukte run hierboven achterliet, en het script zou daarop hebben geantwoord met "Teams is up to date - nothing to do": de beslissing om werk te doen keek alleen naar de vermelding in Programma's en onderdelen, terwijl de DLL-controle die dit opmerkt alleen rapporteerde |
| `Update-TeamsClient.ps1` verwijdert een werkende vergader-add-in niet langer voordat het weet dat het een vervanging kan installeren. Een `-Force`-run in productie verwijderde in stap 6 elke kopie en faalde daarna in stap 8 met `1638`, waardoor de sessiehost helemaal geen add-in meer had. De sweep is verplaatst naar stap 8, achter de versievergelijking: een oudere MSI dan de geregistreerde add-in betekent nu dat de sweep en de installatie worden overgeslagen en dat de werkende add-in precies blijft zoals hij is |
| Daarvoor moesten drie dingen samenvallen, en alle drie worden nu afgehandeld. `-Force` op een host waarvan de build nieuwer is dan de gepubliceerde is een **downgrade**, en de versiecontrole waarschuwt daar nu met zoveel woorden voor. Een uninstall van de add-in die `1612` antwoordt betekent dat Windows Installer zijn bron kwijt is, dus wordt het opnieuw geprobeerd met de eigen gecachte MSI onder `C:\Windows\Installer` (via `Installer\UserData\S-1-5-18\Products\*\InstallProperties`, `LocalPackage`); als die ook weg is, meldt de run dat de registratie niet kan worden verwijderd en wat dat tot gevolg zal hebben. Een `1638` op de add-in is nu een waarschuwing in plaats van een afbreking, zodat de verificatie nog steeds draait en meldt waar Outlook werkelijk mee achterblijft |
| De preflight toont de versie van de geregistreerde add-in in plaats van alleen "is installed" - dat ene getal was de hele diagnose van de fout, en het was het enige dat niet op het scherm stond |
| De AppLocker-regel toont niet langer een lege samenvatting wanneer `SrpV2` bestaat zonder regelverzamelingen eronder (zoals op de productiehost): hij zegt "no rule collections configured, so it blocks nothing" |
| Geverifieerd: het opzoeken van het gecachte pakket tegen echte geïnstalleerde producten, en dat vragen naar een product dat deze machine niet heeft niets teruggeeft zonder een fout te gooien; de versiebeveiliging in alle vier combinaties (ouder, nieuwer, gelijk, onleesbaar); de AppLocker-regel bij een lege verzameling. **Niet getest:** de retry met `msiexec /x <cached msi>` en het waarschuwingspad bij `1638` op een live host |

### 2026-09-25
| Wijziging |
|--------|
| `Update-TeamsClient.ps1` roept niet langer loos alarm over AppLocker. De controle op SlimCore-blokkades waarschuwde zodra `HKLM:\SOFTWARE\Policies\Microsoft\Windows\SrpV2` bestond — en dat is zo in elke vloot die ooit ook maar één Exe-regel heeft geschreven — dus de waarschuwing ging af op machines waar helemaal niets werd geblokkeerd. Een controle die altijd afgaat is een controle die niemand leest |
| Het beleid wordt nu gelezen in plaats van alleen gedetecteerd, op de drie punten die bepalen of het de MSIX kan tegenhouden: alleen de verzameling voor packaged apps (`Appx`) is van toepassing, omdat een MSIX nooit de `Exe`/`Msi`/`Script`/`Dll`-regels tegenkomt; een verzameling met regels waarvan de afdwinging *niet geconfigureerd* is, wordt volgens Microsoft toch afgedwongen, en alleen een expliciete `EnforcementMode = 0` laat alles door; en er wordt helemaal niets afgedwongen zolang de Application Identity-service (`AppIDSvc`) is gestopt, wat nu hardop wordt gezegd in plaats van het in de een of andere richting aan te nemen |
| Het rapport is iets waar een technicus mee aan de slag kan: het registerpad, de modus per verzameling, de status van de service en de eerste vijf `Appx`-regelnamen met hun actie. Een regel die de pakketten al op naam toestaat wordt gemeld als `[ OK ]`; een regel die alles toestaat wat is ondertekend door `O=MICROSOFT CORPORATION` wordt gemeld als waarschijnlijk voldoende, met de opmerking om te controleren dat hij niet is ingeperkt tot één productnaam |
| Een beleid dat op een sessiehost wordt gevonden, wordt nu benoemd als een `[SKIP]`-referentieregel in plaats van verborgen: het blokkeert daar niets, omdat het klaarzetten op het endpoint gebeurt, maar het is meestal hetzelfde GPO — dus wat de moeite van controleren waard is, is of het ook de endpoints bereikt |
| Elke tak doorlopen tegen een gestubde beleidsboom: een afgedwongen `Exe`-verzameling zonder `Appx`-verzameling levert geen blokkade op (het oude fout-positief, verdwenen); een afgedwongen `Appx`-verzameling zonder passende allow-regel waarschuwt met pad en aantal regels; allow-regels voor expliciet SlimCore en voor uitgever Microsoft leveren hun twee verschillende opmerkingen op; `EnforcementMode = 0` leest als alleen audit; afdwinging-niet-geconfigureerd-met-regels leest als afgedwongen; zeven regels tonen er vijf en `... and 2 more`; een ontbrekende `SrpV2`-sleutel levert niets op. Bevestigd dat het stil blijft op deze machine, die geen AppLocker-beleid heeft. **Niet getest** tegen een live afgedwongen AppLocker-beleid op een echt endpoint |
| Bekende beperking, gedocumenteerd in plaats van verborgen: de match op allow-regels is een tekstmatch op de XML van de regel, dus een brede regel die Microsoft noch de pakketten noemt (`PublisherName="*"`) zou SlimCore in werkelijkheid doorlaten, maar wordt toch als blokkade gemeld. De regelnamen die ernaast worden getoond zijn wat daarover de doorslag geeft |

### 2026-09-24
| Wijziging |
|--------|
| `Update-TeamsClient.ps1` beantwoordt de vraag waar de inventarisatie geen antwoord op had: de preflight leest nu de `Microsoft Teams VDI`-events uit het Application-log op elke session host — niet alleen met `-AvdOptimizations` — en vertaalt de codes aan de hand van Microsofts tabel met verbindingsfouten, zodat een gewone `-CheckOnly` meldt of gebruikers daadwerkelijk geoptimaliseerd zijn in plaats van alleen of de onderdelen geïnstalleerd zijn |
| `24002`/`24010` betekenen dat de gebruiker op SlimCore zit, `16002` dat een endpoint nog steeds geen plugin heeft, `16389`/`10083`/`1951` dat beleid op het endpoint de MSIX blokkeert. Een `errc` van nul staat bewust niet in de tabel: die betekent dat die fase geen fout gaf, en "OK" tonen naast een echte fout in de andere fase zou een leugen zijn |
| De query gebruikt `-FilterXPath`, omdat `Get-WinEvent -FilterHashtable @{ ProviderName = ... }` een fout gooit als de provider nog nooit een event heeft geschreven — en dat is de normale situatie op een gezonde niet-VDI-machine. Gemeten: 357 ms en een zachte fout als hij ontbreekt, 104 ms als hij aanwezig is |
| Nieuwe `-RemoveWebRtcRedirector` verwijdert de oude optimalisatie, die op 1 oktober 2026 met pensioen gaat. Sluit `-AvdOptimizations` uit en wordt al vóór de UAC-prompt geweigerd, hergebruikt het `msiexec /x` + verouderde-`1605`-vermelding-pad dat al bewezen is voor classic Teams, en laat `IsWVDEnvironment` staan omdat SlimCore die vlag ook nodig heeft. Standaard uit: een endpoint dat geen SlimCore kan en de redirector niet meer vindt, valt stilletjes terug op het renderen van media op de session host |
| Beide documenten gecorrigeerd op de plekken waar ze een technicus nog vertelden SlimCore op de session host te zoeken |

### 2026-09-20 (8)
| Wijziging |
|--------|
| "De add-in laadt nog steeds niet" krijgt nu een antwoord in plaats van een status. Een registratie die aanwezig is maar niet laadt, wordt getoetst aan de drie oorzaken die in `LoadBehavior` zelf geen spoor achterlaten, elk gemeld als een `why:`-regel: een bitness-verschil tussen Outlook en de geregistreerde loader, Outlook dat de add-in in zijn resiliency-lijsten `DisabledItems`/`CrashedAddins` heeft geparkeerd, en een group policy die het laadgedrag van de gebruiker overschrijft |
| De resiliency-controle decodeert de binaire waarden in de hive van die gebruiker en matcht op het pad van de add-in, zodat hij de ene oorzaak meldt die een technicus uit `LoadBehavior` helemaal niet kan aflezen — Outlook schakelt een gecrashte add-in uit en houdt hem uitgeschakeld, en daarom blijft het vinkje terugzetten niet plakken |
| Als niets op de machine het blokkeert, zegt hij dat ook, en ook dat is een antwoord: wat overblijft is een volledige herstart van Outlook en een gebruiker die minstens één keer bij Teams heeft aangemeld |
| Beide detecties uitgeprobeerd: een x86-loaderpad tegenover deze x64-Office levert de bitness-reden op, en een geplante binaire `CrashedAddins`-waarde wordt gedecodeerd en gemeld. Een gezonde registratie levert geen `why:`-regel op |

### 2026-09-20 (7)
| Wijziging |
|--------|
| Een volledige herinstallatie verwijdert nu **elke** kopie van de meeting add-in voordat de nieuwe wordt geïnstalleerd, niet alleen de kopie die de MSI kent: de machinebrede map, de mappen per profiel onder `%LOCALAPPDATA%\Microsoft\TeamsMeetingAdd-in`, en de COM-registraties per gebruiker in elke geladen hive |
| Daarmee is de cirkel rond de `LoadBehavior 2` waar dit script al twee dagen achteraan zit gesloten. Een kopie die bij een herinstallatie achterblijft, is precies wat een registratie per gebruiker wordt die de verse machinebrede overschaduwt terwijl hij naar bestanden wijst die niet meer bestaan — zo is `admin` geregistreerd geraakt tegen add-in `1.24.19202` |
| `Get-TeamsAddInFolder` geverifieerd op dit apparaat: hij vindt de echte kopie per profiel. Het `-WhatIf`-plan toont dat de map en beide CLSID-weergaven vóór de herinstallatie worden verwijderd. Het verwijderen zelf hergebruikt mechanismen die al live bewezen zijn in de tests voor classic Teams en reparatie, maar de opruimronde als geheel draait voor het eerst op een productiehost |

### 2026-09-20 (6)
| Wijziging |
|--------|
| Een controle gecorrigeerd die op de verkeerde plek zocht: het script waarschuwde `SlimCore packages not found` op session hosts, maar Microsoft zet SlimCore klaar **op het endpoint**, niet op de VM — *"Step 3: SlimCore MSIX staging and registration on the endpoint ... the plugin silently executes this step, without user or admin intervention"*. De waarschuwing was ruis op elke session host, en een controle die op de verkeerde plek zoekt faalt niet, die liegt |
| Het rapport houdt nu rekening met de context. Op een session host bevestigt het de Teams-build tegen het gedocumenteerde minimum `24193.1805.3040.8975` en meldt het dat SlimCore op het endpoint hoort. Op een endpoint meldt het of de pakketten klaargezet zijn, en controleert het de drie beleidsregels die volgens Microsoft dat klaarzetten blokkeren, elk met de Teams-foutcode die erbij hoort: `BlockNonAdminUserInstall` (16389), `AllowAllTrustedApps` (15615) en AppLocker (10083) |
| Ook de versievraag van vorige week beslecht: Windows App for Windows `2.0.352.0` is het gedocumenteerde minimum op het endpoint, en de klassieke Remote Desktop-client wordt hiervoor helemaal niet meer ondersteund |
| Robuustheid: MSI-exitcode `1641` (geslaagd, herstart al gestart) telde als fout en brak de run af. Die telt nu als geslaagd met een gemarkeerde herstart, net als `3010` |
| Aan beide kanten geverifieerd: dit endpoint meldt `Microsoft.Teams.SlimCoreVdiHost.win-x64 2026.31.1.16`; met een nagebootste `RDInfraAgent` verschijnt in plaats daarvan de tekst voor de session host; de drie blokkades zijn uitgeprobeerd tegen gestubde registry-leesacties |

### 2026-09-20 (5)
| Wijziging |
|--------|
| Een schone productierun bevestigde drie eerdere fixes op een echte session host: de redirector ter plekke gerepareerd (`The download is the installed version (1.56.2603.20001)`, dus de vorige run heeft echt 1.54 → 1.56 geüpgraded), de add-in na het provisionen gevonden in het klaargezette pakket, en de hele flow eindigde met exit `0` |
| Hij bracht ook de ene resterende waarschuwing in kaart: `BAKKERPARTNERS\admin` heeft een registratie die wijst naar add-in `1.24.19202`, een kopie per gebruiker die allang weg is en een prima gezonde machinebrede `1.26.21803` overschaduwt. `-RepairOutlookAddIn` (Ninja-variabele `repairOutlookAddIn`) ruimt die verouderde `Classes\CLSID\{19A6E644-...}`-sleutel nu op en zet `LoadBehavior` terug op 3, zodat COM weer naar de machinebrede registratie verwijst |
| Hij grijpt alleen in als die machinebrede registratie gezond is — de schaduw weghalen zonder dat er iets achter zit, zou de gebruiker slechter af laten zijn — en hij staat standaard uit, omdat hij in de hive van een andere gebruiker schrijft. Het telt als werk, dus `-CheckOnly` meldt het en `-Quiet` laat het zien |
| Getest tegen geplante sleutels in beide registry-weergaven: `-WhatIf` plant beide acties, een toegepaste run ruimt de CLSID-sleutels op, zet `LoadBehavior` op 3 en eindigt met exit `0`. Niet getest: of Outlook de add-in daarna voor die gebruiker ook echt laadt — dat is de volgende productierun |

### 2026-09-20 (4)
| Wijziging |
|--------|
| "Ik zie hem nog niet op alle profielen geladen" was een zichtbaarheidsgat, niet alleen een Teams-probleem: een profiel waarvan de hive niet gemount is, kan helemaal niet gelezen worden, en het script liet het gewoon weg — waardoor een onleesbaar profiel en een gezond profiel er in de uitvoer identiek uitzagen. Het noemt die profielen nu bij naam, met wat het voor ze betekent: met een gezonde machinebrede registratie pakken ze de add-in op bij de eerste start van Outlook, zonder die is er niets om op terug te vallen |
| Het verdient het om het ronduit te zeggen, omdat het bepaalt of er iets te repareren valt: een profiel dat niet aangemeld is, is niet kapot. De machinebrede registratie dekt gebruikers zonder eigen staat per gebruiker; alleen een gebruiker die al een eigen registratie heeft (uitgeschakeld, of wijzend naar een verwijderde DLL) blijft hem overschaduwen |
| Beide meldingen geverifieerd tegen gestubde profiellijsten. Hier niet geverifieerd: het mounten van een niet-gemounte hive om een afgemeld profiel te inspecteren of te repareren — `reg load` vereist rechten die dit werkstation niet heeft, dus dat mechanisme is bewust niet gebouwd op een ongeteste aanname |

### 2026-09-20 (3)
| Wijziging |
|--------|
| Een vraag beantwoord waar het script geen antwoord op had: **waarom** een account `LoadBehavior 2` toont. Outlook vindt de add-in via `Classes\CLSID\{19A6E644-...}\InprocServer32`, en een registratie per gebruiker in `HKCU\SOFTWARE\Classes` gaat vóór de machinebrede — dus een gebruiker blijft de kopie uit zijn eigen profiel laden, zelfs nadat een `ALLUSERS=1`-installatie in `Program Files (x86)` terechtkomt. Gemeten op een apparaat: de class verwijst naar `%LOCALAPPDATA%\Microsoft\TeamsMeetingAdd-in\<version>\x64\Microsoft.Teams.AddinLoader.dll` |
| Het script bepaalt dat pad nu per aangemelde gebruiker en meldt de twee gevallen apart, omdat ze een andere fix nodig hebben: de add-in uitgeschakeld maar de DLL aanwezig (vinkje terugzetten) tegenover een registratie die naar een DLL wijst die weg is (het vinkje blijft niet staan — hij moet voor die gebruiker opnieuw geïnstalleerd worden). De gerichte opzoeking over de geladen hives kost ~100 ms |
| Getest met een geplante registratie die naar een ontbrekende DLL wijst, zonder de echte aan te raken |

### 2026-09-20 (2)
| Wijziging |
|--------|
| Derde productiefout op dezelfde session host, derde fix: `Uninstall of Teams Machine-Wide Installer failed (exit code 1605)`. 1605 betekent "deze actie is alleen geldig voor producten die momenteel geïnstalleerd zijn" — de vermelding in Programma's en onderdelen heeft het product overleefd, wat vaak gebeurt zodra de nieuwe Teams-bootstrapper over een machine is gegaan |
| In dat geval valt er niets te verwijderen, maar door de verouderde vermelding zou het script bij elke run classic Teams blijven melden, dus verwijdert het nu in plaats daarvan de registry-vermelding en gaat het verder. Live getest: een echte `msiexec /x` tegen een onbekende productcode geeft 1605, de run waarschuwt, verwijdert een geplante verouderde vermelding, verifieert schoon en eindigt met exit `0` |
| De objecten voor de uninstall-vermeldingen dragen nu hun `RegistryPath` en `UninstallString` mee, en dat maakt die opruiming mogelijk |

### 2026-09-20
| Wijziging |
|--------|
| De tweede productiefout op een session host opgelost: `WebRTC Redirector install failed (exit code 1638)`. Die MSI houdt over versies heen één ProductCode aan, dus `msiexec /i` over een bestaande installatie weigert met "er is al een andere versie van dit product geïnstalleerd" in plaats van te upgraden — en `-Force` loopt daar recht in op elke host die de redirector al heeft |
| Het script leest nu de ProductVersion van de download en beslist: dezelfde versie → ter plekke repareren (`REINSTALL=ALL REINSTALLMODE=vomus`), een andere versie → eerst de oude verwijderen, dan installeren. Een `1638` die er toch doorheen glipt, wordt gemeld als "de bestaande blijft staan" in plaats van de hele run te laten mislukken |
| Gemeten tijdens het fixen: `aka.ms/msrdcwebrtcsvc/msi` levert nu `1.56.2603.20001`, terwijl op die host `1.54.2408.19001` geïnstalleerd was — dit was dus een geweigerde upgrade, geen dubbele installatie. Beide paden worden onder `-WhatIf` correct gepland; geen van beide msiexec-aanroepen is al echt uitgevoerd, en dat staat met zoveel woorden in de documentatie |
| Ook expliciet vastgelegd: een geïnstalleerde redirector wordt door een normale run **niet** stilletjes geüpgraded. Alleen `-Force` vervangt hem. Nu WebRTC op 1 oktober 2026 zijn ondersteuning verliest, is dat bewust houden beter dan automatisch een onderdeel upgraden dat op zijn retour is |

### 2026-09-18 (4)
| Wijziging |
|--------|
| `Get-DistributionGroupMembers.ps1` — `-Member "*.verizon.com"` matcht nu een domein **en al zijn subdomeinen** (`.verizon.com` en `*@*.verizon.com` zijn hetzelfde). Zonder de `*.` vooraan blijft het filter bij dat ene domein, dus `@be.verizon.com` reikt nog steeds bewust niet tot `@us.verizon.com` |
| De run zegt welke van de twee hij doet — *"scanning N list(s) for members on verizon.com and its subdomains"* — want een filter waarvan je de reikwijdte moet raden, is een filter dat je in een klantrapport niet kunt vertrouwen |
| Er wordt gematcht op het volledige domeinlabel, geverifieerd tegen `@notverizon.com` en de suffixtruc `@verizon.com.evil.test`; geen van beide matcht bij een `*.verizon.com`-run. Een wildcard ergens anders dan vooraan wordt ge-escaped in plaats van het filter stilletjes breder te maken |

### 2026-09-18 (3)
| Wijziging |
|--------|
| `Get-DistributionGroupMembers.ps1` — **`-Recurse`**, na te hebben gecontroleerd of het rapport echt iedereen dekte: dat deed het niet. Exchange geeft alleen ooit *directe* leden terug, dus een lijst die een andere lijst bevat, meldde die lijst als één lid en nooit de mensen erin. Iemand die alleen via een geneste groep mail ontvangt, was onzichtbaar, en `-Member` meldde "geen treffers" bij een lijst die wel bij hem aflevert — een fout antwoord dat eruitziet als een zelfverzekerd antwoord |
| `Via groep` noemt de groep waarlangs iemand binnenkwam (leeg voor een direct lid), en iemand die via meerdere routes bereikbaar is, krijgt één rij met de routes samengevoegd in plaats van een rij per route |
| `Aantal leden` blijft directe leden tellen, omdat dat het getal is dat Exchange en het EAC tonen; de nieuwe `Aantal personen` telt de werkelijk bereikte ontvangers |
| Een groep die al uitgevouwen is, wordt niet nog eens uitgevouwen, en dat voorkomt ook dat een lidmaatschapslus (A bevat B, B bevat A) eindeloos recursief doorgaat. Geverifieerd tegen een bewust cyclisch paar testlijsten; nesting dieper dan 20 niveaus wordt gemeld en met rust gelaten |
| Vastgelegd wat het rapport nog steeds *niet* dekt: het leest groepslidmaatschap, dus een gebruiker die op geen enkele lijst staat, komt nergens voor |

### 2026-09-18 (2)
| Wijziging |
|--------|
| `Get-DistributionGroupMembers.ps1` — `-Member` accepteert nu ook een **domein**: `-Member "@be.verizon.com"` meldt elke lijst die nog een adres op dat domein bevat (`be.verizon.com` en `*@be.verizon.com` betekenen hetzelfde). Een adres wordt door Exchange zelf gematcht; een domein kan dat niet, dus wordt elke lijst gelezen en daarna gefilterd — trager, en zo ook gedocumenteerd |
| Er wordt gematcht op het primaire adres, elke alias **en `ExternalEmailAddress`**. Daar draait het juist om bij een partnerdomein: zo'n lid is meestal een e-mailcontact met als primaire SMTP `...@contoso.onmicrosoft.com`, en het echte `@be.verizon.com` staat alleen in het externe adres. Matchen op het primaire adres had niets gevonden en met een stalen gezicht "geen" gemeld |
| Nieuwe kolom `Extern adres` in het werkblad `Leden`, zodat bij contacten het adres zichtbaar is dat de mail daadwerkelijk ontvangt, in plaats van alleen de interne placeholder |
| Met een actief filter: `Treffers` per lijst in `Overzicht`, en `Treffer op` per lid in `Leden`. `Treffer op` bevat het matchende **adres**, geen Ja/Nee — een treffer op een alias is anders onverklaarbaar in een rapport dat geen aliassen toont |
| Een domeinfilter dat niets matcht, zegt dat ook en schrijft geen bestand, in plaats van een lege werkmap af te leveren die leest als een mislukte export |

### 2026-09-18
| Wijziging |
|--------|
| `Get-DistributionGroupMembers.ps1` toegevoegd — elke distributielijst met zijn leden in één Excel-werkmap: een werkblad `Overzicht` (één rij per lijst) en een werkblad `Leden` (één rij per lid), allebei filterbare tabellen met een vastgezette kopregel. Kolomkoppen en ontvangertypen staan in het Nederlands, omdat de klant de werkmap leest |
| `-Member user@domain` beantwoordt "op welke lijsten staat deze persoon?" server-side via `Get-Recipient -Filter "Members -eq '<DN>'"` in plaats van elke groep af te lopen, en exporteert de gevonden lijsten toch volledig, zodat de klant ziet wie er verder op staat |
| `-IncludeDynamic` en `-IncludeM365Groups` breiden het rapport uit voorbij gewone distributiegroepen; dynamische groepen worden live geëvalueerd, omdat ze geen lidmaatschap opslaan om op te vragen |
| Valt terug op twee CSV-bestanden als `ImportExcel` ontbreekt (en biedt eerst aan het te installeren), zodat een ontbrekende module je nooit het rapport kost. `ImportExcel` toegevoegd aan `Install-Modules.ps1` — `vias_archiver.ps1` had het al nodig |
| Exchange-submenu: `K` Get-DLMembers |

### 2026-09-17 (3)
| Wijziging |
|--------|
| De bug opgelost die een productierun op een AVD-session host aan het licht bracht: `teamsbootstrapper.exe -p` **provisiont** het pakket voor toekomstige aanmeldingen, het installeert het niet voor degene die het script uitvoerde. De bootstrapper meldde succes en de add-in-stap strandde daarna op `New Teams package not found after install`, omdat `Get-AppxPackage -Name MSTeams` naar de huidige gebruiker vraagt en de admin die het script draaide geen Teams had |
| De add-in-MSI wordt nu gevonden door te globben op `%ProgramFiles%\WindowsApps\MSTeams_*_x64__8wekyb3d8bbwe\MicrosoftTeamsMeetingAddinInstaller.msi` en de nieuwste versie te nemen, zodat het werkt ongeacht of een gebruiker het pakket geïnstalleerd heeft. Opnieuw getest voor een installatie per gebruiker, een host met alleen provisioning en een `-Force`-run |
| Dezelfde blinde vlek per gebruiker in de SlimCore-controle, die "niet gevonden" meldde op een host die new Teams voor één profiel heeft: die vraagt nu eerst `-AllUsers`. En de twee machinebrede Outlook-registryweergaven krijgen apart het label 64-bit/32-bit, omdat die run twee identieke regels `all users (machine-wide)` toonde, wat als een bug leest |
| Die run bewees ook meteen het nut van de nieuwe controles: hij verwijderde een echte classic Teams `1.4.00.11161` per profiel, en vond `LoadBehavior 2` voor één account - Outlook had de add-in uitgeschakeld, en dat los je met geen enkele herinstallatie op |

### 2026-09-17 (2)
| Wijziging |
|--------|
| `Update-TeamsClient.ps1` kan nu ook classic Teams verwijderen, achter `-RemoveClassicTeams` (Ninja-variabele `removeClassicTeams`). Standaard uit: een applicatie bij gebruikers weghalen is geen beslissing die een updatejob op eigen houtje hoort te nemen. Het verwijdert de *Teams Machine-Wide Installer* via msiexec — de installer die ertoe doet, want zolang die aanwezig is, zet Windows classic Teams in elk nieuw profiel klaar — en ruimt per profiel de installatiemap, de autostartvermelding `Run\com.squirrel.Teams.Teams` en de verouderde sleutel `Uninstall\Teams` op |
| De gedocumenteerde verwijdering per gebruiker (`Update.exe --uninstall -s`) moet als de eigenaar van het profiel draaien, en dat kan System niet, dus worden in plaats daarvan de bestanden verwijderd. Roaming-gegevens in `%APPDATA%\Microsoft\Teams` blijven ongemoeid |
| De foutafhandeling houdt de twee gevallen bewust uit elkaar: een machinebrede installer die de verwijdering overleeft, is een echte fout (exit 1), terwijl een map per profiel die het overleeft bijna altijd een bestandslock van een draaiende classic Teams is — een waarschuwing, die de volgende run oplost nadat de gebruiker zich heeft afgemeld |
| Getest met nagebootste detectie, omdat het testapparaat geen van beide varianten heeft: `-WhatIf` plant de msiexec-verwijdering en het verwijderen van de map en slaat stap 5-8 over, en een toegepaste run verwijderde een nagebootste profielmap waarbij de verificatie hem als schoon meldde. Het msiexec-pad zelf is **niet** uitgevoerd tegen een echte Machine-Wide Installer — dat staat in de documentatie in plaats van dat het wordt gesuggereerd |

### 2026-09-17
| Wijziging |
|--------|
| Een bug opgelost die een echte run op een session host blootlegde: `Get-AppxPackage -AllUsers` vindt niets als Teams alleen *geprovisiond* is en nog geen enkele gebruiker het heeft, dus had de versiecontrole niets om te vergelijken, verklaarde de host verouderd en installeerde bij elke geplande run ~275 MB opnieuw. De geïnstalleerde versie valt nu terug op de versie van het geprovisionde pakket. Geverifieerd tegen precies dat scenario: meldt `Provisioned MSTeams <version>`, vergelijkt en doet niets |
| `Update-TeamsClient.ps1` controleert nu of **Outlook zelf** de meeting add-in ziet, niet alleen of de MSI geïnstalleerd is. Het leest `HKEY_USERS\<sid>\...\Outlook\Addins\TeamsAddin.FastConnect` per aangemelde gebruiker plus de machinebrede sleutel: `LoadBehavior 3` = geladen, `2`/`0` = Outlook heeft hem uitgeschakeld (het echte "de knop is weg"-geval). Gemeld in preflight en verificatie, nooit als fout — een profiel waarop niemand is aangemeld, kan niet gelezen worden |
| De preflight werd een volledige inventarisatie van overal waar Teams kan staan: AppX per gebruiker, het geprovisionde pakket, de classic *Teams Machine-Wide Installer*, classic installaties per profiel, de add-in in beide hives, de Outlook-registratie. Classic Teams wordt gemeld, niet verwijderd — het deelt de einddatum van ondersteuning in oktober 2026 en een achtergebleven machinebrede installer blijft het in nieuwe profielen opnieuw klaarzetten |
| Twee bugs gevonden door het uit te voeren in plaats van het te lezen. Profielen opsommen met een `S-1-5-21-*`-whitelist slaat **elke** gebruiker over op een Entra-joined apparaat, waar SID's `S-1-12-1-*` zijn — de controle beweerde dat niemand de add-in geregistreerd had terwijl `LoadBehavior=3` er gewoon stond. En `New-PSDrive` respecteert `ShouldProcess`, dus onder `-WhatIf` werd de `HKEY_USERS`-drive nooit aangemaakt en loog dezelfde alleen-lezen-controle; de hives worden nu aangesproken via `Registry::HKEY_USERS`, zonder drive om aan te maken |

### 2026-09-11 (5)
| Wijziging |
|--------|
| `scripts/Exchange/Move-SharedCalendar.ps1` toegevoegd — alles in één: `-Search balie` vindt de agenda, toont wie hem gebruikt, verplaatst hem met `Convert-SharedCalendarToResource.ps1` naar een resource-mailbox en somt op wie moet overstappen. Eén tijdelijke App Registration met de rechten van beide scripts, één keer aangemaakt en aan het eind verwijderd, dus één aanmelding in plaats van twee. Bij meerdere treffers kies je uit een lijst of beperk je met `-Owner`; een niet-interactieve run somt ze op en stopt in plaats van te gokken. De twee scripts worden aangeroepen, niet gekopieerd, dus er is van elke stap één implementatie |
| `Convert-SharedCalendarToResource.ps1` gerepareerd: in een niet-interactieve sessie werd de getypte bevestiging overgeslagen en de oorspronkelijke agenda zonder `-Force` verwijderd. Die blijft nu met een waarschuwing staan, tenzij `-Force` is opgegeven |
| Beide agendascripts: een expliciete `-ClientId` gaat nu vóór een bestaande app-only Graph-sessie, zodat de app van een aanroepend script echt gebruikt wordt. `Convert-SharedCalendarToResource.ps1` krijgt `-PassThru` (resultaatobject voor aanroepers) |
| Exchange-submenu (`C`) optie `J` toegevoegd |

### 2026-09-11 (4)
| Wijziging |
|--------|
| `Get-CalendarMappings.ps1` — de eerste echte run (een tenant waar "Balie planning" een secundaire agenda in een gearchiveerde mailbox bleek te zijn) bevestigde de Graph-aannames: app-only leesacties geven de agenda's terug die gebruikers hebben toegevoegd, en een gedeelde secundaire agenda verschijnt in hun lijst onder zijn eigen naam. Hij bracht ook een verkeerde hint aan het licht: een `NotMapped`-rij wees naar "Reservering vergaderzaal LBM" als "waarschijnlijk deze", terwijl dat een *tweede* agenda is die dezelfde eigenaar deelt. De hint slaat nu vermeldingen over die naar een andere agenda van de eigenaar zijn genoemd, en - bij een secundaire agenda - vermeldingen die naar de eigenaar zijn genoemd, want dat is diens hoofdagenda |

### 2026-09-11 (3)
| Wijziging |
|--------|
| `scripts/Exchange/Convert-SharedCalendarToResource.ps1` toegevoegd — verplaatst een gedeelde agenda (de "Balie"-agenda in de mailbox van één persoon) naar een eigen Room- of Equipment-mailbox, met elk item en elke machtiging, en verwijdert desgevraagd het origineel. Standaard een voorbeeldweergave; `-Apply` maakt aan en kopieert, `-RemoveSourceCalendar` verwijdert het origineel pas nadat elk item een geverifieerde kopie heeft en de naam van de agenda als bevestiging is getypt |
| Items worden getrouw gekopieerd in plaats van bij benadering: terugkerende reeksen blijven reeksen, met hun verplaatste en geannuleerde exemplaren toegepast (exemplaar voor exemplaar gematcht, en met een waarschuwing met rust gelaten als de twee reeksen niet overeenkomen), tijden worden teruggeschreven in de tijdzone waarin ze zijn aangemaakt zodat wekelijkse items een zomertijdwissel overleven, categorieën houden hun kleur, bijlagen tot 3 MB worden gekopieerd en grotere in de back-upmap opgeslagen. Deelnemers worden in de hoofdtekst vermeld in plaats van gekopieerd, zodat niemand een nieuwe uitnodiging ontvangt |
| Machtigingen behouden hun exacte Exchange-toegangsrechten, aangepaste rechten inbegrepen; `-SendSharingInvitation` stuurt gebruikers de standaarduitnodiging. Externe personen, verwijderde accounts en gemachtigde-vlaggen worden gemeld, niet stilletjes weggelaten |
| Elke kopie draagt het id van het bronitem in een verborgen eigenschap, zodat een run die halverwege stopt verdergaat waar hij gebleven was; een half afgemaakte reeks wordt opnieuw gedaan. Er wordt een JSON-back-up van alles wat gelezen is geschreven voordat er iets wordt aangemaakt |
| Exchange-submenu (`C`) optie `I` toegevoegd: altijd eerst een voorbeeldweergave, daarna een expliciete tweede stap |

### 2026-09-11 (2)
| Wijziging |
|--------|
| `Get-CalendarMappings.ps1` krijgt `-Search` (alias `-Keyword`): "waar staat de Balie-agenda?" in één run. Het trefwoord wordt vergeleken met de naam van de eigenaar en elk adres (een gedeelde mailbox `balie@`, een ruimte, een groep) en met agendanamen (een secundaire agenda *Balie* in iemands mailbox). Het rapport toont waar de agenda staat (nieuwe status `Source`), wie hem in de agendalijst heeft, en wie er rechten op heeft |
| Van een matchende **secundaire** agenda worden nu de eigen machtigingen gelezen, in plaats van dat hij met de hoofdagenda van de eigenaar wordt vergeleken en op `MappedWithoutRight` uitkomt. Een nieuwe kolom `Calendar` zegt over welke agenda van de eigenaar een rij gaat |
| Een vermelding in de agendalijst bevat geen verwijzing terug naar de map waar hij vandaan komt, dus een gedeelde secundaire agenda wordt op naam gematcht. Als een gebruiker hem onder een andere naam heeft, noemt de `NotMapped`-rij de vermelding die het waarschijnlijk is, in plaats van een stille fout-negatief achter te laten |
| Menuoptie `H` vraagt eerst om een trefwoord; leeg valt terug op het volledige rapport of het rapport per mailbox |

### 2026-09-11
| Wijziging |
|--------|
| `scripts/Exchange/Get-CalendarMappings.ps1` toegevoegd — toont waar elke agenda daadwerkelijk gekoppeld is: voor elke mailbox leest het de agendalijst in Outlook en de rechten op de eigen hoofdagenda, en voegt beide samen tot één rij per eigenaar + gebruiker met een status (`Mapped`, `MappedWithoutRight`, `NotMapped`, `MappedOwnerMissing`, `SharedExternally`, …). `Test-CalendarPermissions.ps1` zegt wie een agenda *mag* openen; dit zegt waar hij *staat*, en waar die twee van elkaar afwijken |
| Graph in plaats van Exchange Online PowerShell, omdat de vermeldingen die een gebruiker aan zijn eigen agendalijst heeft toegevoegd voor geen enkele Exchange-cmdlet zichtbaar zijn. App-only toegang volgt dezelfde drie routes als `Remove-PhishingMessage.ps1` (bestaande sessie, eigen app, of een tijdelijke app die in een `finally` wordt verwijderd), met alleen-lezen-rechten `Calendars.Read`, `User.Read.All` en `Group.Read.All`. Geen Exchange-verbinding, dus geen MSAL-conflict |
| Tenantbrede runs lopen via `$batch` (20 mailboxen per aanroep), waarbij gethrottelde items opnieuw worden geprobeerd. Een mailbox die niet gelezen kan worden, wordt als zodanig gemeld in plaats van als "niets gekoppeld" |
| Vastgelegd wat het rapport niet kan zien: Full Access met AutoMapping (een mailboxmachtiging — `Test-MailboxPermissions.ps1`), agenda's die in classic Outlook zonder verbeteringen voor gedeelde agenda's zijn geopend, en secundaire agenda's, die als `MappedWithoutRight` verschijnen |
| Exchange-submenu (`C`) optie `H` toegevoegd |

### 2026-09-16 (5)
| Wijziging |
|--------|
| De uitlegpagina is nu stap 4 van de build in één commando, in plaats van een los script waar iemand een week later aan denkt. Een structuur waar niemand over is verteld, is een structuur die niemand gebruikt, en omdat de pagina uit dezelfde configuratie wordt gegenereerd, beschrijft ze precies wat de run zojuist heeft gemaakt |
| `-SkipHelpPage` laat haar weg, `-HelpContact` zegt bij wie mensen terechtkunnen. De installer geeft `-Force` mee, omdat die pagina van hem is: de build opnieuw draaien brengt de uitleg weer in lijn met wat de build heeft gemaakt |
| Verificatie en de deelstatus-audit zijn opgeschoven naar stap 5 en 6, en de verouderde waarschuwingen "stap 2 wijzigt rechten op een live team" noemen nu stap 3 |
### 2026-09-16 (4)
| Wijziging |
|--------|
| `scripts/SharePoint/Provisioning/Add-SharePointHelpPage.ps1` toegevoegd — zet de uitleg voor eindgebruikers als SharePoint-pagina op de teamsite, gelinkt vanuit de navigatie links. Een handleiding in een repo leest niemand; dit zet haar waar de mensen die bestanden uploaden al zijn |
| De pagina wordt uit de configuratie gegenereerd in plaats van uitgetypt, zodat ze niet kan afwijken van wat de bibliotheken werkelijk doen: de kanalen die ze noemt zijn de kanalen die bestaan, de labels dragen dezelfde helptekst die onder elk veld in het uploadformulier verschijnt, en de verplichte velden per documenttype worden van de content types afgelezen |
| Geschreven voor wie een catalogus uploadt. Twee stukken van de configuratie blijven er bewust buiten: de `note` op een container, die security groups noemt, en de `description` op een view, die het over pijlers en gesynchroniseerde mappen heeft. Rechten blijven helemaal buiten beschouwing — wie wat mag zien is niet iets waar een gebruiker iets mee kan |
| Onderweg gerepareerd, gevonden door de pagina te renderen in plaats van de code te lezen: ze kondigde drie manieren aan om een bestand toe te voegen en noemde er twee, en de wizard schreef de teamnaam waar de bedrijfsnaam hoorde ("Intern blijft binnen Laseto-NewTeams") |
### 2026-09-16 (3)
| Wijziging |
|--------|
| `scripts/SharePoint/Provisioning/Remove-SharePointStructure.ps1` toegevoegd — haalt dezelfde configuratie weer uit elkaar, diepste eerst: tabs, kanalen, bibliotheken, content types (eerst losgekoppeld van hun lijsten), sitekolommen, term set, security groups en het team zelf |
| **Bewust de omgekeerde standaard van al het andere in de map: zonder `-Apply` verandert het niets.** `-WhatIf` vergeten bij een destructief script is de gevaarlijke kant, dus de veilige toestand is degene die je gratis krijgt |
| `-Scope All` omvat nooit het team. Het hele team van een klant verwijderen is niet iets wat iemand zou moeten krijgen door om "alles" te vragen — het moet benoemd worden, en daarna moet de naam van het team worden ingetypt ter bevestiging |
| Weigert standaard in plaats van achteraf vergeving te vragen: een bibliotheek of kanaalmap waar nog bestanden in staan wordt overgeslagen tenzij `-IncludeContent` (het aantal items wordt hoe dan ook gerapporteerd), het General-kanaal en de eigen Documents-bibliotheek van het team worden nooit verwijderd, en een content type dat nog in gebruik is wordt gerapporteerd in plaats van geforceerd |
| Met hun kosten gerapporteerd voordat ze draaien, omdat geen prullenbak ze terugbrengt: de term set verwijderen maakt de Leverancier-waarde wees op elk document dat er een had, en een kolom verwijderen neemt de data mee. Wat *wel* te herstellen is wordt ook gezegd — een verwijderde groep of team is 30 dagen soft-deleted, een kanaal heeft zijn eigen prullenbak van 30 dagen, en bestanden uit een verwijderde bibliotheek belanden in de prullenbak van de site |
| Menustap `6` draait het; het menu vraagt naar toepassen, naar bestanden en naar het team als drie aparte vragen in plaats van één |
### 2026-09-16 (2)
| Wijziging |
|--------|
| Een afgeschermde pijler kan nu worden vormgegeven als **een eigen bibliotheek achter een gewoon kanaal**, de enige opzet die een echte alleen-lezen-rol geeft en toch een kanaal in Teams zet. De wizard vraagt welke pijlers afgeschermd zijn en daarna in welke vorm - `bibliotheek` (de standaard) of `privekanaal` |
| De bibliotheekvorm verbreekt de overerving **zonder die te kopiëren**, en daar draait het om: kopiëren neemt elk teamlid mee als bewerker, precies de deur die deze vorm moet sluiten. Wat overblijft zijn de eigen owners van de site plus de twee groepen van de pijler - Contribute en Read |
| Een privékanaal biedt helemaal geen alleen-lezen-rol: owners en members, en members mogen posten, bewerken en verwijderen. Een pijler die "mag kijken" nodig heeft kan dus geen privékanaal zijn, en de wizard zegt dat nu op het moment dat de keuze wordt gemaakt |
| Het kanaal wordt zoals gewoonlijk aangemaakt maar krijgt geen map in de gedeelde bibliotheek, en de afgeschermde bibliotheek wordt als tab in dat kanaal getoond - het eigen Files-tabblad van een standaardkanaal wijst altijd naar de teambibliotheek en kan niet worden omgelegd, dus staat ze ernaast |
| In de readme gemarkeerd omdat het anders als bug gemeld wordt: dat ingebouwde Files-tabblad blijft staan en wijst naar een map die niemand gebruikt. Wijs mensen op de benoemde tab, of verwijder het Files-tabblad eenmalig met de hand uit het kanaal |
### 2026-09-16
| Wijziging |
|--------|
| `scripts/SharePoint/Provisioning/Sync-SharePointChannelMember.ps1` toegevoegd — maakt een Entra ID-security group de bron van waarheid voor wie er in een privé-Teams-kanaal zit. Een privékanaal kan helemaal geen rechten via een groep krijgen: Teams houdt de bezetting per persoon bij en Graph accepteert daar alleen individuele gebruikers, dus voedt de groep in plaats daarvan de bezetting |
| De voor de hand liggende workaround is een valkuil en is als zodanig gedocumenteerd: de groep toevoegen aan de SharePoint-rechten van de site van het privékanaal werkt tot Teams de bezetting er weer overheen synchroniseert, en intussen bereiken die mensen de bestanden terwijl het kanaal voor hen onzichtbaar blijft in Teams. Niet ondersteund door Microsoft |
| Geneste groepen worden gevolgd, niet-gebruikers vallen af, en iedereen wordt eerst aan het bovenliggende team toegevoegd — Teams weigert een lid van een privékanaal dat niet in het team zit, en de fout die het teruggeeft zegt dat niet. `-Prune` verwijdert ook mensen die de groepen niet meer noemen; kanaal-owners worden nooit verwijderd |
| Opgeschreven omdat het het ontwerp verandert, niet alleen het script: **een privékanaal heeft geen alleen-lezen-rol.** Owners en members, en members mogen posten, bewerken en verwijderen. Een groep met de naam `-RO` kan daar niet "mag kijken" betekenen, dus de run meldt per groep hoeveel mensen hij binnenbracht in plaats van dat ongemerkt voorbij te laten gaan. Waar alleen-lezen echt telt, is een documentbibliotheek met eigen rechten de juiste vorm |
| `-EnsureGroups` maakt nu elke groep in het model aan in plaats van alleen de groepen waaraan een bibliotheek rechten geeft. Een privékanaal geeft niets, dus het MGMT-paar stond in de configuratie en werd nooit aangemaakt — precies de pijler waarvan je de groepen als eerste gaat zoeken |
| De wizard schrijft voor privépijlers een `channelMembers`-sectie die de groepen noemt die de bezetting voeden; een configuratie die geschreven is voordat die sleutel bestond valt terug op de geconfigureerde groepen die naar de container zijn vernoemd, en meldt die terugval |
| Menustap `5` draait de sync; die meldt zich zelf aan bij Graph, dus vraagt niet om een PnP-app-registratie |
### 2026-09-15 (3)
| Wijziging |
|--------|
| `New-StructureConfig.ps1 -All` vraagt naar de namen die tot nu toe achter de rug van de operator werden afgeleid: per pijler de kanaalnaam, de map, het content type, beide groepsnamen en de titel van de view; plus de bibliotheek achter de kanalen, de kolom- en content-typegroepen, de term set, de URL van de teamsite en het label dat elke kolom voor de gebruiker draagt. Elk houdt zijn afleiding als suggestie, dus `-All` is nog steeds vooral Enters |
| Alleen de *interne* kolomnamen blijven vast. Ze worden aan niemand getoond, en er een wijzigen nadat documenten hem dragen kost de metadata op die documenten |
| Gerepareerd: een optionele vraag kon nooit worden afgeslagen, omdat Enter "neem de suggestie" betekent. Optionele vragen zeggen nu `(of "geen")` en accepteren geen/none/nee/- als een echt "geen" — hiervoor maakte niets antwoorden op "klantbibliotheek" nog steeds FUTECH aan |
| Gerepareerd dat een lijst met één item terugkwam als losse string: PowerShell rolt een array met één element uit bij het teruggeven, dus een klant met één merk liet de wizard crashen op `.Count`. Wordt nu met een voorloopkomma teruggegeven |
| Twee toewijzingen `$x = if (...) { @() }` gerepareerd die `$null` opleveren in plaats van een lege array — een configuratie zonder verkooppijler of zonder leveranciers liep vast bij de samenvatting |
| Alle drie de paden van begin tot eind geverifieerd tegen de configvalidator: de volledige standaard met zes pijlers, een `-All`-run met overal bewust andere namen, en een minimale tenant met twee pijlers zonder privékanaal, zonder leveranciers, zonder regio's en zonder klantbibliotheek |
### 2026-09-15 (2)
| Wijziging |
|--------|
| `New-StructureConfig.ps1` vraagt hoe alles moet heten en schrijft de configuratie zelf — niemand zou een JSON-bestand moeten openen om een kanaal een naam te geven. Enter accepteert de suggestie tussen haakjes, dus een standaard build is vooral Enters plus de tenant en de team-owner |
| Al het andere wordt uit die antwoorden afgeleid: per pijler een kanaal, een content type, twee security groups en een gegroepeerde view; per merk een doorsnijdende view over elke pijlermap. Welke pijler leveranciers afhandelt en welke verkoop, bepaalt waar Leverancier en Regio verplichte velden worden |
| Twee van de antwoorden zijn degene die later iets kosten, dus die worden als laatste gevraagd en staan standaard op nee: het bijhouden van de deelstatuskolom (het enige nachtelijke script) en het afdwingen van rechten per pijler op mappen van standaardkanalen (het deel dat Microsoft niet ondersteunt) |
| `New-SharePointTeam.ps1` maakt het Microsoft 365-team en zijn kanalen aan, het privé-MGMT-kanaal inbegrepen, zodat de structuur vanaf een lege tenant kan worden opgebouwd. De sitecollectie van een privékanaal wordt asynchroon ingericht en de URL ervan is vooraf niet te kennen — het script pollt ernaar en schrijft haar terug in de configuratie, en dat is wat de volgende stappen ergens mee laat verbinden |
| Hernoemt of verwijdert nooit een kanaal: een kanaal waarvan de naam niet overeenkomt met de config wordt gemeld, niet gecorrigeerd, omdat hernoemen de map verplaatst en elke link breekt die iemand heeft gedeeld |
| `Install-SharePointStructure.ps1` draait de wizard zelf wanneer het geen configuratie voor de tenant vindt, en de teamstap is nu stap 1 van zes. `-SkipTeam` voor een team dat al bestaat |
| Interne kolomnamen en content-type-ID's worden één keer gegenereerd en liggen daarna vast — SharePoint koppelt documentmetadata aan beide — en daarom weigert de wizard een bestaande configuratie te overschrijven zonder `-Force`. Weergavenamen, kanaalnamen en groepsnamen blijven wijzigbaar |
| De laatste hardgecodeerde kolomnamen verwijderd: de deelstatus-audit leest welke kolom welke is uit een nieuwe `fieldRoles`-sectie in plaats van `PsDeelstatus` en `PsVertrouwelijkheid` aan te nemen |
| De app-registratie geeft nu ook consent voor `Channel.Create`, `ChannelSettings.ReadWrite.All` en `Team.Create`, die de teamstap nodig heeft |
### 2026-09-15
| Wijziging |
|--------|
| `scripts/SharePoint/Provisioning/Install-SharePointStructure.ps1` toegevoegd — bouwt de hele structuur in één run: registreert zelf de Entra-app en geeft admin consent voor de delegated scopes, dan metadata → bibliotheken/groepen/rechten → verificatie, optioneel de eerste deelstatus-audit. Elke stap blijft een eigen script, dus bij een fout draai je alleen die stap opnieuw in plaats van opnieuw te beginnen |
| `-TemporaryApp` verwijdert de app-registratie aan het eind weer, voor een eenmalige build op een tenant die je niet dagelijks beheert. Het verwijdert alleen ooit een app die **deze run heeft aangemaakt** — een app die al in de cache stond dateert van vóór de run en is aan iemand anders om te verwijderen, dus het script zegt dat in plaats van hem stilletjes te verwijderen. Zonder de switch blijft de app staan en wordt de client-ID gecachet in `pnp.appid.json`, gedeeld met de andere PnP-scripts in deze repo |
| In de docs gemarkeerd omdat het anders als bug gemeld wordt: een `-WhatIf`-run heeft een app nodig om mee aan te melden. Zonder gecachete app voor de tenant is er niets om als te verbinden, dus de proefdraai valideert de config en stopt daar — draai één keer echt, of geef `-ClientId` mee, om stap voor stap een proefdraai te doen |
| Het gat gedicht dat "merk als tag" maar half waar maakte: doorsnijdende views op de gedeelde bibliotheek met `Scope = RecursiveAll`, zodat *Alles - Butterstone* één platte lijst is over elke pijlermap, **inclusief alles wat met Beide getagd is** — één bestand, twee merken, geen kopieën die uit elkaar lopen. Plus *Nog te taggen* (wat drag-and-drop en OneDrive-sync achterlaten), *Extern gedeeld* en *Te archiveren*. Het filter is ruwe CAML in de config in plaats van een mini-querytaal die het script zelf heeft verzonnen |
| Groepeer een view nooit op `PsTaal`: SharePoint weigert te groeperen op een kolom met meerdere waarden. Erop filteren werkt prima, en geen enkele meegeleverde view groepeert erop |
| `Petsolutions-SharePoint-Handleiding.md` toegevoegd — documentatie voor eindgebruikers in het Nederlands om aan de klant te geven. Behandelt de drie manieren om een bestand toe te voegen en waarom ze zich anders gedragen, wat elk label betekent, en wat er gebeurt op het moment dat je iets tagt (het bestand verhuist niet, links blijven werken, `Beide` verschijnt in beide merkviews, zoeken loopt een paar minuten achter op de views) |
| Twee dingen die de handleiding hardop zegt omdat gebruikers het tegenovergestelde aannemen: **drag-and-drop en OneDrive-sync vragen niets** — verplichte kolommen worden afgedwongen door het uploadformulier, niet door de bibliotheek, dus in bulk gedropte bestanden komen binnen met lege labels en een melding "Required info" in plaats van geblokkeerd te worden; en **een label is geen slot** — Vertrouwelijkheid sluit niemand buiten, het is een afspraak plus het signaal dat de nachtelijke audit gebruikt om overmatig delen te markeren |
| Menu-item `S` kreeg stap `0` voor de alles-in-één-build; de opties per stap zijn ongewijzigd |

### 2026-09-10 (4)
| Wijziging |
|--------|
| `scripts/SharePoint/Provisioning/` toegevoegd — een hele SharePoint-structuur (metadatamodel, content types, bibliotheken, rechten via Entra ID-groepen) voor een MSP-klant inrichten en onderhouden vanuit één JSON-config, met een deel-audit en een alleen-lezen driftcontrole. Gebouwd voor Petsolutions NV (merken Butterstone/Laseto), maar niets in de scripts is klantspecifiek |
| Het model staat in `petsolutions.config.json` en wordt bij het laden gecontroleerd: een content type dat naar een niet-gedefinieerde kolom verwijst, of een container die rechten geeft aan een groep die niet in het model staat, faalt voordat er iets verbindt in plaats van halverwege het inrichten. De meegeleverde `CHANGEME`-tenant/site-URL's worden botweg geweigerd |
| Interne kolomnamen krijgen een `Ps`-prefix. "Contenttype" en "Status" zijn weergavenamen die SharePoint al voor iets anders gebruikt, en de prefix houdt ze eenduidig in CAML, in views en in de driftcontrole, terwijl gebruikers nog steeds gewone Nederlandse labels zien. Content-type-ID's liggen vast in plaats van gegenereerd te worden, zodat dezelfde structuur over tenants heen reproduceerbaar is |
| `New-SharePointMetadata.ps1` draait tegen **elke** site in de config, niet alleen de teamsite: een privé-Teams-kanaal (hier MGMT) is een eigen sitecollectie en een sitekolom reikt daar niet overheen. Een kolom achteraf verplicht maken werkt — de `Required`-vlag op een bestaande field link wordt ter plekke bijgewerkt en doorgezet naar de lijsten die het content type al gebruiken |
| `Set-SharePointLibraries.ps1` stelt per map een volgorde van content types in, zodat het menu *Nieuw* in het Leveranciers-kanaal Leveranciersdocument aanbiedt en niet de vijf typen van de andere pijlers — de gedeelde bibliotheek moet ze allemaal dragen, de map hoeft ze niet te tonen. Geen enkele view wordt ooit de standaard gemaakt: de standaardview van een Teams-bibliotheek is wat elk lid van het kanaal ziet op het moment dat hij Files opent |
| Opgeschreven in plaats van verstopt: unieke rechten op een map van een **standaard**kanaal zijn wat dit model vraagt en wat Microsoft niet ondersteunt. Leden die toegang verliezen blijven het kanaal in Teams zien en krijgen een fout op het Files-tabblad in plaats van een dichte deur. Het script doet het, waarschuwt per map, en `-SkipChannelFolderPermissions` laat die mappen erven. Een privékanaal, een gedeeld kanaal of een eigen bibliotheek (wat FUTECH gebruikt) zijn de ondersteunde manieren om een pijler af te schermen |
| `Update-SharePointShareStatus.ps1` leidt de Deelstatus-kolom af uit de rechten die werkelijk op elk bestand staan. Het stelt eerst de goedkope vraag — een bestand dat erft is niet gedeeld — dus één round trip per honderd items beslist vrijwel de hele bibliotheek; alleen bij bestanden die de overerving verbraken worden de roltoewijzingen gelezen, en daarvan hoeven alleen links voor specifieke personen te worden uitgeklapt (een Anyone- of Organization-link zegt in zijn naam al of er een gast achter kan zitten). Schrijft met `SystemUpdate` zodat Modified/Modified By blijven staan en er geen versie wordt aangemaakt |
| De audit trekt nooit een link in. Hij meldt bestanden getagd Intern of Vertrouwelijk die achter een externe link zitten en eindigt met `2`, zodat een geplande RMM-job precies naar boven komt wanneer er een beslissing voor een mens ligt. `Test-SharePointStructure.ps1` doet hetzelfde voor structurele drift, ingedeeld als Missing / Different / Extra — "Extra" wordt nooit automatisch gerepareerd, omdat een extra kolom data bevat en een extra roltoewijzing meestal iemands bewuste uitzondering is |
| `SharePointStructure.Common.ps1` wordt door alle vier gedot-sourcet — een bewuste uitzondering op de regel "elk script staat op zichzelf" elders in deze repo, omdat ze één config-schema delen en drie kopieën van de rechtencode binnen een maand uit elkaar zouden lopen |
| Menu-item `S` toegevoegd voor de set (kies een stap, `-WhatIf` tenzij je bevestigt; de driftcontrole slaat de vraag over omdat die nooit schrijft) |
### 2026-09-10 (4)
| Wijziging |
|--------|
| De vervaldatum gekoppeld aan de `-AvdOptimizations`-functie: Microsoft stopt met de op WebRTC gebaseerde AVD-mediaoptimalisatie op **1 oktober 2026** (einde ondersteuning) en **1 april 2027** (einde beschikbaarheid), en Teams toont gebruikers er al een banner over. De switch blijft de redirector installeren omdat Microsoft die nog als fallback aanraadt — met een notitie om er vóór april 2027 nog eens naar te kijken |
| De opvolger, SlimCore, heeft niets nodig op de session host: hij zit in de nieuwe Teams. Bevestigd op een apparaat met Teams `26225.1806.5074.1452`, dat `Microsoft.Teams.SlimCoreVdiHost.win-x64` `2026.31.1.16` plus meerdere frameworkpakketten bevat. Preflight meldt dat pakket nu onder `-AvdOptimizations` |
| Die melding is bewust informatief en maakt geen werkitem aan: welk mediapad wordt gebruikt hangt af van de versie van de Windows App op het endpoint waarvandaan de gebruiker verbindt, en dat kan een script op de session host niet zien. De versies van de clients op de endpoints auditen is het eigenlijke migratiewerk |
| Een servicedesksectie toegevoegd aan het IT Glue-document voor de banner die gebruikers melden: wat hij betekent (een aankondiging, geen storing), de twee datums, dat de oplossing op het lokale apparaat ligt en niet op de session host, hoe je de regel `AVD SlimCore Media Optimized` / `AVD Media Optimized` onder Teams > About leest, en kant-en-klare tekst voor de gebruiker |

### 2026-09-10 (3)
| Wijziging |
|--------|
| Een onjuiste bewering in de Teams-docs en in het scriptcommentaar gecorrigeerd: de uninstall-vermelding van de meeting add-in staat **niet** altijd in `WOW6432Node`. Gemeten op een Windows 11-endpoint registreert add-in `1.26.21803` zich in de **64-bit** hive, met `InstallSource` wijzend naar een MSI-cache per gebruiker. Beide hives scannen (wat het script al deed) is juist — de opgegeven reden was dat niet |
| Gedocumenteerd hoe de add-in werkelijk op een apparaat komt, gemeten in plaats van aangenomen: het script installeert hem machinebreed (`ALLUSERS=1`, `Program Files (x86)`) voor gedeelde machines en session hosts, terwijl op een gewoon endpoint de Teams-client hem **per gebruiker** installeert en bijwerkt vanuit `%LOCALAPPDATA%\Microsoft\TeamsMeetingAddinMsis` naar `%LOCALAPPDATA%\Microsoft\TeamsMeetingAdd-in`, en zich alleen registreert in `HKCU\...\Office\Outlook\Addins` |
| Er meteen bij opgeschreven: wat stap 8 werkelijk bewijst. Die leest de HKLM-uninstallsleutels, dus bevestigt dat de machinebrede installatie is gelukt — niet dat de Outlook van een bepaalde gebruiker de knop toont. Als System gedraaid kan het script de `HKCU` van een gebruiker helemaal niet zien |

### 2026-09-10 (2)
| Wijziging |
|--------|
| `scripts/Device/Update-TeamsClient.ps1` neemt de AVD/VDI-onderdelen van de oudere aanvullende installer over achter `-AvdOptimizations`: de mediavlag `IsWVDEnvironment` (gezet in stap 3, voordat de client wordt geprovisioned, omdat Teams hem bij het opstarten leest om zijn mediapad te kiezen) en de Remote Desktop WebRTC Redirector Service van `aka.ms/msrdcwebrtcsvc/msi` |
| Bewust een switch en geen autodetectie: die vlag zetten op een normaal endpoint vertelt Teams media door te geven aan een redirector die er niet is. Zonder de switch *meldt* het script alleen dat een apparaat eruitziet als een session host (`HKLM:\SOFTWARE\Microsoft\RDInfraAgent`) |
| Beide onderdelen worden alleen geïnstalleerd als ze ontbreken (`-Force` installeert de redirector opnieuw), dus een geplande run op een geconfigureerde session host downloadt nog steeds niets en print niets onder `-Quiet`. Van begin tot eind geverifieerd, inclusief dat de redirector-MSI (1.7 MB, `1.54.2408.19001`) de Microsoft-handtekeningcontrole doorstaat |
| De add-in-stap slaat zichzelf nu ook over als de add-in aanwezig is en de client niet is vervangen — hiervoor installeerde hij de add-in opnieuw bij een run die er alleen was om de AVD-onderdelen te repareren |
| Download en handtekeningverificatie verhuisd naar één `Save-VerifiedDownload`-helper die de bootstrapper en de redirector delen: alleen https, minimale grootte, Authenticode `Valid` en ondertekend door `O=Microsoft Corporation`, anders volgt een throw |
| Gecorrigeerd in de docs: namen van Ninja-scriptvariabelen zijn **niet** hoofdlettergevoelig. Het opzoeken van omgevingsvariabelen in Windows is niet hoofdlettergevoelig, dus variabelen met de naam `Quiet` of `Force` werken precies zoals `quiet` en `force` |

### 2026-09-10
| Wijziging |
|--------|
| `scripts/Device/Update-TeamsClient-ITGlue.md` kreeg een bijlage voor de NinjaOne-inrichting: welke waarden je per veld kiest bij het toevoegen van het script (PowerShell 5.1 in plaats van 7, 64-bit, Run As System), de namen van de scriptvariabelen met een manier om te verifiëren dat ze echt aankomen, de testrun op één apparaat, de geplande automation met `-Quiet -Confirm:$false`, en de optionele detectiejob |
| Expliciet opgeschreven omdat het anders als bug gemeld wordt: `-CheckOnly` eindigt met `2` als er een update beschikbaar is, en NinjaOne toont elke exitcode die niet nul is als mislukte job. Dat is de bedoeling — dat zijn de apparaten die aandacht nodig hebben — en daar kan een script result condition op sleutelen |
| Ook gemarkeerd: de Ninja-scripttimeout moet groter zijn dan `-TimeoutSeconds` (900 s) plus de download van ~275 MB, anders breekt Ninja de job halverwege de installatie af |

### 2026-09-08 (5)
| Wijziging |
|--------|
| `scripts/Device/Update-TeamsClient-ITGlue.md` toegevoegd — de servicedeskversie van die documentatie, in het Nederlands, om in IT Glue te plakken. Gelaagd per supportniveau: L1 controleert met `-CheckOnly -Quiet` en leest de gelabelde output, L2 draait de update vanuit NinjaOne of met de hand en verifieert achteraf, L3 krijgt parameters, exitcodes, paden en de ingebouwde beveiligingen. Bevat een fouttabel met het escalatieniveau per melding, een FAQ en kant-en-klare tekst voor de eindgebruiker |
| Het opent met het punt waar mensen over struikelen: geen output betekent dat het apparaat al actueel is, en dat is een geslaagde run en geen fout. Het downloadvolume per apparaat (~275 MB: een bootstrapper van 1.9 MB die een pakket van ~273 MB binnenhaalt) is gemeten, niet geschat |

### 2026-09-08 (4)
| Wijziging |
|--------|
| `scripts/Device/Update-TeamsClient.md` toegevoegd — een naslag voor dat script: de beslisboom (achter → volledige herinstallatie, actueel maar add-in ontbreekt → alleen add-in, actueel → helemaal niets), de zeven stappen, de versiecontrole via de configservice met een voorbeeldantwoord, outputmodi en exitcodes, de tabel met NinjaOne-scriptvariabelen, de ontwerpkeuzes achter de volgorde van handelingen, een troubleshootingtabel en wat er werkelijk getest is. Gelinkt vanuit de Device-readme |

### 2026-09-08 (3)
| Wijziging |
|--------|
| `scripts/Device/Update-TeamsClient.ps1` herinstalleert niet langer onvoorwaardelijk: het vraagt de configservice van de Teams-client (`config.teams.microsoft.com/config/v1/MicrosoftTeams/...`, `BuildSettings.WebView2PreAuth.<arch>.latestVersion` — dezelfde feed die de client gebruikt om te bepalen dat hij verouderd is) welke build voor deze architectuur is gepubliceerd, en laat een actueel apparaat volledig met rust |
| Is de client actueel maar ontbreekt de meeting add-in? Dan wordt alleen de add-in geïnstalleerd — geen download, geen uninstall, geen herprovisioning |
| `-Quiet` houdt alle output achter tot er nieuws is, dus een geplande NinjaOne-run print niets op een actueel apparaat en verschijnt alleen in de activity feed als hij een nieuwere build vond of op een probleem stuitte. Geverifieerd: een actuele `-Quiet`-run levert nul bytes output op en exitcode 0 |
| `-CheckOnly` meldt zonder iets te wijzigen en eindigt met `2` als er een nieuwere build beschikbaar is, voor gebruik als Ninja-detectie/conditiejob. `-Ring` kiest een andere dan de standaard update-ring |
| Als de configservice niet bereikbaar is stopt de run in plaats van blind te herinstalleren; `-Force` betekent nu "herinstalleer ook al is hij actueel" en ook "ga door zonder Teams- of versie-informatie" |
| Er wordt alleen een transcript geschreven als de run daadwerkelijk iets verandert, dus een controle per uur laat geen logrommel achter in `C:\Temp` |

### 2026-09-08 (2)
| Wijziging |
|--------|
| `scripts/Device/Update-TeamsClient.ps1` is nu veilig om onbeheerd vanuit een RMM (NinjaOne) *en* met de hand te draaien. Het herstart zichzelf 64-bit via `SysNative` als de agent PowerShell 32-bit start — anders worden de HKLM-leesacties omgeleid naar `WOW6432Node` en wijst `$env:ProgramFiles` naar de x86-map, zodat noch het AppX-pakket noch de add-in-MSI ooit wordt gevonden |
| NinjaOne-scriptvariabelen (`whatIf`, `force`, `skipMeetingAddIn`, `skipSignatureCheck`, `workingDir`, `logPath`) worden uit de omgeving gelezen als de bijbehorende parameter niet is meegegeven, zodat een previewrun een vinkje kan zijn in plaats van een parameterstring |
| Met de hand gestart zonder verhoogde rechten vraagt het nu om UAC en gaat het verder in een verhoogd venster, in plaats van te falen op een regel `#Requires -RunAsAdministrator`, en een interactieve apply-run vraagt één keer om bevestiging. `-Confirm:$false` maakt het onbeheerd; het menu geeft dat mee omdat het al gevraagd heeft |
| Herschikt zodat de bootstrapper gedownload **en** zijn Microsoft Authenticode-handtekening geverifieerd is vóór de eerste uninstall — een mislukte download of een geblokkeerde URL kan een apparaat niet langer zonder Teams-client achterlaten. TLS 1.2 wordt afgedwongen voor de download, en een niet-https `-BootstrapperUrl` wordt geweigerd |
| `msiexec` en de bootstrapper draaien nu via één helper met een timeout (`-TimeoutSeconds`, standaard 900, proces wordt bij verlopen gekild), een retry bij 1618 (een andere installatie bezig) en 3010 behandeld als succes met een notitie over een openstaande herstart, zodat een RMM-job de agent nooit kan laten hangen |
| Het AppX-pakket wordt ook gedeprovisioned (`Remove-AppxProvisionedPackage`), anders krijgen nieuwe gebruikersprofielen de oude versie uit de image klaargezet |
| De versie van de add-in-MSI komt nu uit de MSI-property-tabel via het COM-object `WindowsInstaller.Installer`. `Get-AppLockerFileInformation` — wat Microsofts eigen voorbeeld gebruikt — ontbreekt op sommige edities en sleept onder PowerShell 7 de Windows PowerShell-compatibiliteitslaag mee, die faalt en een `-WhatIf`-run overspoelt met niet-gerelateerde file-copy-output |
| Apply-runs schrijven een transcript naar `C:\Temp\Update-TeamsClient_<timestamp>.log`; onverwachte fouten breken af in plaats van halverwege door te gaan; exitcode is 0 bij succes (`-WhatIf` inbegrepen) en 1 bij falen |

### 2026-09-08
| Wijziging |
|--------|
| `scripts/Device/Update-TeamsClient.ps1` toegevoegd — schone herinstallatie van de nieuwe Teams op een endpoint of AVD-session host: de Teams Meeting Add-in verwijderen, het `MSTeams`-AppX-pakket voor alle gebruikers verwijderen, `teamsbootstrapper.exe` downloaden, Teams provisionen (`-p`) en de meeting-add-in-MSI installeren die in het nieuwe Teams-pakket zit |
| Elke stap die de toestand wijzigt loopt via `ShouldProcess`, dus `-WhatIf` doorloopt de hele flow en print elke uninstall/download/install zonder de machine aan te raken; de stappen die pas na een echte installatie bestaan (nieuwe Teams-versie, pad van de add-in-MSI, eindverificatie) worden als zodanig gemeld in plaats van de run te laten falen |
| Het opzoeken van de add-in leest zowel de 64-bit als de `WOW6432Node`-uninstallhive — de add-in installeert 32-bit, dus de 64-bit hive alleen vindt hem nooit (uninstall en verificatie misten hem allebei voorheen) |
| Exitcodes en de exitcodes van msiexec/bootstrapper worden gecontroleerd in plaats van aangenomen; `-SkipMeetingAddIn` vervangt alleen de client, `-Force` installeert op een apparaat zonder enige Teams. Aangesloten op `menu.ps1` (toets T), dat standaard een `-WhatIf`-preview doet |

### 2026-09-07 (2)
| Wijziging |
|--------|
| `scripts/SharePoint/Search-SharePointContent.ps1` toegevoegd — de Microsoft Graph-tegenhanger van `Find-SiteContent.ps1`: app-only, geen interactieve aanmelding, en het doorzoekt één site of **elke site en OneDrive in de tenant** |
| Content vinden gebeurt via `/drives/{id}/root/delta` (een hele bibliotheekboom in pagina's van duizend items, zodat `*contains*`-wildcards werken) of `/search/query` met `-Content` voor tekst in documenten; de filters zijn dezelfde als in het PnP-script |
| Rechten komen uit `/drives/{id}/items/{id}/permissions`, 20 per `/$batch`-aanroep. Eén aanroep levert de rollen, de identiteiten waaraan rechten zijn gegeven, de deellink met zijn bereik (anyone/organization/specific people), bewerken-of-bekijken, vervaldatum en URL, en `inheritedFrom` — dat bepaalt `PermissionSource = Item` (uniek) versus `Inherited`. "Iedereen met de link" krijgt een eigen teller omdat die helemaal geen aanmelding nodig hebben |
| Expliciet opgeschreven, in het script en de readme: Graph heeft geen API voor SharePoint-roltoewijzingen, dus rechten op site- en lijstniveau en items in gewone lijsten (geen bibliotheken) blijven het domein van `Find-SiteContent.ps1`. De readme heeft een vergelijkingstabel om tussen de twee te kiezen |
| Aanmelden is app-only: de eerste run registreert een app, geeft consent voor de application role `Sites.Read.All`, maakt een self-signed certificaat aan in `CurrentUser\My` en uploadt de publieke sleutel — geen secret op schijf — en cachet client-ID plus thumbprint per tenant in `graph.appid.json` (toegevoegd aan `.gitignore`). Latere runs verbinden zonder prompt, dus het werkt ook vanuit een geplande taak. Throttling (429) wordt opnieuw geprobeerd met inachtneming van `Retry-After`, zowel voor losse aanroepen als voor batch-subrequests |

### 2026-09-07
| Wijziging |
|--------|
| `scripts/SharePoint/Find-SiteContent.ps1` toegevoegd — doorzoek een hele SharePoint-site of OneDrive op content en rapporteer welke rechten op elke treffer van toepassing zijn. Alleen-lezen |
| Twee engines: een crawl over elke lijst en bibliotheek (ziet alles, `-IncludeSubsites` voor de subsites) en een KQL-query tegen de zoekindex (`-Content`) die ook tekst *in* documenten vindt. Beide delen de filters `-Name`, `-Path`, `-Extension`, `-ItemType`, `-ListName`, `-ModifiedBy`, `-ModifiedAfter`/`-ModifiedBefore` en `-MinSizeMB` |
| Per treffer bepaalt het script waar de rechten vandaan komen — het item zelf (verbroken overerving), zijn lijst of de site — en vlakt de roltoewijzingen af tot één CSV-rij per principal met type, login, e-mail en rolnamen. `Limited Access` wordt eruit gefilterd tenzij `-IncludeLimitedAccess` |
| Deellinks (de `SharingLinks.*`-groepen achter "Link kopiëren") worden altijd uitgeklapt naar de mensen erin en gelabeld als Anyone/Organization/Specific people; externe gasten (`#ext#`) en "Everyone (except external users)" worden apart gemarkeerd in de samenvatting en de CSV |
| Site- en lijstrechten worden één keer gelezen en gecachet, en itemrechten alleen voor items die de overerving verbraken, dus de kosten schalen met het aantal treffers, niet met de grootte van de site; `-Permissions Unique` rapporteert alleen wat anders gedeeld is, `-Permissions None` slaat rechten over, en `-MaxPermissionLookups` begrenst een te brede zoekopdracht |
| Hergebruikt de app-registratieflow en de `pnp.appid.json`-cache per tenant van `Restore-RecycleBinItems.ps1`, en kan zichzelf tijdelijk sitecollectiebeheerder maken (`-GrantSiteAdmin`) om een site of OneDrive te doorzoeken waar het geen rechten op heeft |

### 2026-08-28
| Wijziging |
|--------|
| `scripts/SharePoint/` toegevoegd met `Restore-RecycleBinItems.ps1` — herstelt verwijderde bestanden/mappen uit de prullenbak van een SharePoint-site of OneDrive, standaard als proefdraai, met filters op naam, oorspronkelijke map, wie het verwijderd heeft en een tijdvenster voor de verwijdering |
| Twee scopes: `-SiteUrl` voor één sitecollectie (OneDrive inbegrepen), of `-AllSites -TenantUrl` om elke SharePoint-site in de tenant af te lopen. De tenantbrede sweep slaat persoonlijke OneDrive-sites, de My Site-host, redirect-sites en vergrendelde sites over, ondersteunt `-SiteFilter`/`-MaxSites`, en gaat door wanneer één site een fout geeft — de resultaten per site komen in een samenvattingstabel en in een `Site`-kolom in de CSV |
| Herstellen gebeurt in batches van maximaal 200 items via `Restore-PnPRecycleBinItem -IdList` (één serveraanroep per batch) in plaats van één aanroep per item; mappen en bestanden zitten nooit in dezelfde batch, en een batch die als geheel mislukt wordt item voor item opnieuw geprobeerd, zodat fouten per item toch gerapporteerd worden. Parallelle runspaces zijn bewust niet gebruikt — PnP PowerShell is niet thread-safe en gelijktijdige aanroepen op één sitecollectie lopen tegen SharePoint-throttling aan |
| Het script rapporteert bij elke stap hoe lang het duurt — hoe lang het uitlezen van de prullenbak duurde, vooraf een schatting van het herstel, een voortgangsbalk met een live ETA op basis van de gemeten snelheid, de werkelijke duur in de samenvatting, en een `DurationSeconds`-kolom per item in de CSV |
| Het script registreert bij de eerste run tegen een tenant zijn eigen Entra-app (public client, gedelegeerd `AllSites.FullControl`, met admin consent) omdat PnP PowerShell geen gedeelde multi-tenant-app meer meelevert; de client-ID wordt per tenant gecachet in `pnp.appid.json` (toegevoegd aan `.gitignore`) |

### 2026-07-24 (3)
| Wijziging |
|--------|
| Een niet langer onderhouden interne PowerShell-repo (`Windows-Powershell`, laatste commit maart 2023) uitgefaseerd door elk script erin te beoordelen en alles wat nog waarde had te moderniseren naar deze repo — niets is letterlijk gekopieerd; alles is herschreven tegen Microsoft Graph / Exchange Online (de op `MSOnline`/`AzureAD` gebaseerde scripts uit de bronrepo werken helemaal niet meer sinds Microsoft die endpoints heeft uitgefaseerd) |
| `scripts/TenantOnboarding/` toegevoegd (24 scripts verdeeld over Provisioning/MultiTenant/AppDeployment/DeviceConfig/OneDriveManagement/UserManagement) — gemoderniseerd uit de tenant-setup-/onboardingscripts van de bronrepo |
| `scripts/Office365Toolkit/` toegevoegd (9 scripts verdeeld over Security/Exchange/Intune) — gemoderniseerd uit een geforkte kopie van het uitgefaseerde GitHub-project `directorcia/Office365` (CIAOPS) dat in de bronrepo stond; per functionaliteit beoordeeld en samengevoegd, niet 1-op-1 overgezet |
| `scripts/PatronToolkit/` toegevoegd (13 scripts verdeeld over Entra/Security/Exchange/Intune/SharePoint/Teams) — gemoderniseerd uit een geforkte kopie van het uitgefaseerde GitHub-project `directorcia/patron` dat in de bronrepo stond, met dezelfde aanpak van samenvoegen per functionaliteit |
| `scripts/LegacyUtilities/` toegevoegd (17 scripts verdeeld over Exchange/Entra/Teams/Network/Device/Workspace365) — gemoderniseerd uit diverse kleine tools in de bronrepo die hierboven niet aan bod kwamen |
| Overal is een grens voor gegevensverwerking aangehouden: de mappen `Klanten`, `created-users`, `csv files` en `Archief` uit de bronrepo (echte klantnamen/tenantdomeinen/gegenereerde wachtwoorden) zijn nooit gelezen of overgezet; elk ander script dat echte klant-/tenant-identifiers of secrets hardcodeerde, is in plaats daarvan omgezet naar parameters, of helemaal overgeslagen — zie de readme van elke nieuwe map voor de specifieke lijst van overgeslagen scripts |
| Geen van de ~63 nieuwe scripts is aan `menu.ps1` gekoppeld — het zijn audit-/rapportage-/setupscripts die bedoeld zijn om direct uit te voeren, in lijn met het bestaande patroon voor `scripts/RDS/`, `scripts/Azure/` en `scripts/Network/UniFi/` |

### 2026-07-24 (2)
| Wijziging |
|--------|
| Opgelost dat `scripts/Reporting/Get-SharePointStorageReport.ps1` op zeer grote bibliotheken stilletjes stopte met het opzoeken van versiegeschiedenis: het plafond voor retry-passes van `Invoke-GraphBatchGet` stond hardcoded op 8, maar de throttle per app-activiteit van SharePoint Online staat slechts ~1500-2500 opgeloste versie-lookups per pass toe voordat een afkoelperiode van ~60-90s zich herhaalt — op een tenant met 200k bestanden betekende dit dat ~90% van de bestanden als "gave up" werd gemarkeerd voordat de scan echt klaar was |
| Dezelfde fix toegepast op de eigen kopie van dezelfde batch-retryfunctie (`Get-FileVersionsBatch`) in `scripts/Reporting/Remove-SharePointFileVersionsByDate.ps1`, die hetzelfde hardcoded plafond van 8 passes had |
| Parameter `-MaxVersionRetryPasses` aan beide scripts toegevoegd (standaard `0` = schaalt het pass-plafond automatisch mee met het aantal requests, begrensd op 500 passes); er wordt nu een duidelijke `Write-Warning` gegeven die precies aangeeft hoeveel bestanden zijn opgegeven en de parameter voorstelt als het plafond toch nog bereikt wordt |
| `scripts/Reporting/readme.md` bijgewerkt om `-VersionBatchConcurrency` en `-MaxVersionRetryPasses` voor beide scripts te documenteren |

### 2026-07-24 (1)
| Wijziging |
|--------|
| `scripts/Azure/VM/Azure-NVMe-Conversion.ps1` gedocumenteerd (stond eerder in geen enkele readme/structuurboom): `scripts/Azure/readme.md` en `scripts/Azure/VM/readme.md` toegevoegd, plus een nieuwe categorie "Azure Infrastructure" in deze readme |
| 5 scripts die volledig gedocumenteerd waren maar nog geen menu-ingang hadden, aan `menu.ps1` gekoppeld: `Move-InboxToArchive.ps1` en `Set-Distributionlist-dynamic-static.ps1` (Exchange-submenu, toetsen E/F), `Get-M365UserLicenses.ps1`, `Import-ConditionalAccessBaseline.ps1`, `Set-UserManager.ps1` (Entra-submenu, toetsen G/H/I) — en ze toegevoegd aan de Menu-tabellen in deze readme |
| `scripts/Device/Remove-OemBloatware.ps1` toegevoegd — detecteert HP/Lenovo/Dell en verwijdert OEM-bloatware via winget plus generieke Microsoft Store-rommel via AppX; standaard een proefdraai; gekoppeld aan `menu.ps1` (toets I) |
| `scripts/Device/DriveMapping/New-CloudDriveMapping.ps1` toegevoegd — koppelt SharePoint-/OneDrive-documentbibliotheken via WebDAV aan stationsletters, voor gebruik als aanmeldscript; standaard een proefdraai |
| `scripts/Network/UniFi/` toegevoegd — `UnifiApi.ps1` als gedeelde inloghelper (klassieke controller + automatische detectie van UniFi OS), `Get-UnifiNetworkReport.ps1` (HTML-documentatierapport), `Update-UnifiFirmware.ps1` (tooling voor firmware-upgrades met proefdraai); inloggegevens altijd via `Get-Credential`, nooit hardcoded |
| `scripts/Intune/Compare-IntuneConfig.ps1` toegevoegd — detectie van afwijkingen in de Intune-configuratie tussen een klanttenant en een MSP-baselineback-up, via de module `IntuneBackupAndRestore`; alleen-lezen |

### 2026-07-22 (6)
| Wijziging |
|--------|
| `scripts/Reporting/Get-SharePointStorageReport.ps1` bijgewerkt om scans van alle sites GDAP-bestendig te maken door één effectieve tenantcontext te bepalen en vast te zetten (`-TenantId`, of de GDAP-klantcontext uit `$global:cid`) voor de Graph-aanmelding, het aanmaken van de tijdelijke app en het uitgeven van app-only-tokens |
| Vangrails toegevoegd voor GDAP-/app-only-flows: duidelijkere foutmeldingen wanneer de tenantcontext van de klant ontbreekt (voer eerst `Connect-Tenant` uit of geef `-TenantId` mee) en wanneer `-ClientId` wordt opgegeven zonder een tenant-ID die te bepalen is |
| De invoer voor de checkpointsignatuur in `Get-SharePointStorageReport.ps1` bijgewerkt zodat die ook `-ForceAppOnlySingleSite` en de bepaalde tenantcontext bevat, wat botsingen voorkomt bij het hervatten vanuit een andere context |
| `scripts/Reporting/readme.md` bijgewerkt om het gedrag van de GDAP-tenantbinding bij scans van alle sites te documenteren |

### 2026-07-22 (5)
| Wijziging |
|--------|
| `scripts/Reporting/Get-SharePointStorageReport.ps1` bijgewerkt om scans van één site GDAP-bestendig te maken: wanneer `authMode=GDAP` wordt gedetecteerd, gebruikt het script voor `-SiteUrl`-scans nu automatisch het bootstrappad met tijdelijke app/app-only, om hiaten in gedelegeerde rechten te vermijden |
| Parameter `-ForceAppOnlySingleSite` toegevoegd om de app-only-bootstrap voor scans van één site expliciet af te dwingen, los van de gedetecteerde authenticatiemodus |
| `scripts/Reporting/readme.md` bijgewerkt om het nieuwe GDAP-gedrag en de parameter `-ForceAppOnlySingleSite` te documenteren |

### 2026-07-22 (4)
| Wijziging |
|--------|
| Het gedelegeerde pad van `scripts/Reporting/Get-SharePointStorageReport.ps1` robuuster gemaakt door de afhankelijkheid van ontbrekende Graph-cmdlets weg te nemen: het gebruik van `Get-MgDriveItemChild` vervangen door Graph REST-paginering via `Invoke-MgGraphRequest` voor het doorlopen van de onderliggende items van een drive |
| De fallbacks voor het opsommen van site-drives in `Get-SharePointStorageReport.ps1` bijgewerkt zodat ze in gedelegeerde/niet-app-only-takken Graph REST (`/sites/{id}/drives`) gebruiken in plaats van `Get-MgSiteDrive` |
| Het ophalen van de prullenbak in `Get-SharePointStorageReport.ps1` bijgewerkt zodat Graph REST-paginering (`/sites/{id}/recycleBin/items`) als fallback/primair gedelegeerd pad wordt gebruikt in plaats van de afhankelijkheid van de cmdlet `Get-MgSiteRecycleBinItem` |
| Niet-fatale ruis van MSAL-authority-waarschuwingen bij het verbreken van de verbinding onderdrukt door `Disconnect-MgGraph` in de opruimfase tijdelijk met waarschuwingsonderdrukking te omhullen |

### 2026-07-22 (3)
| Wijziging |
|--------|
| De afhandeling van de scanmodus in `scripts/Reporting/Get-SharePointStorageReport.ps1` bijgewerkt zodat runs voor één site (`-SiteUrl` met `/sites/...` of `/teams/...`) niet langer een tijdelijke app-registratie + app-only-bootstrap starten; de app-only-setup wordt nu alleen nog gebruikt voor tenantbrede opsomming |
| De prestaties en betrouwbaarheid van het opzoeken van één site verbeterd door de exacte site direct via het Graph-URL-pad (`/sites/{hostname}:{path}`) op te lossen in plaats van via een zoek-/filterflow |
| De prestatienotities in `scripts/Reporting/readme.md` bijgewerkt om het geoptimaliseerde pad voor één site en de verwachte opstartsnelheid te documenteren |

### 2026-07-22 (2)
| Wijziging |
|--------|
| Afhankelijkheidsprobleem in het SharePoint-rapport opgelost dat `Get-MgSite`-fouten van het type command-not-found veroorzaakte: `load.ps1`, `scripts/Startup/Install-Modules.ps1` en `scripts/Startup/Update-Modules.ps1` bijgewerkt zodat ze `Microsoft.Graph.Sites` bevatten |
| `scripts/Reporting/Get-SharePointStorageReport.ps1` bijgewerkt met een expliciete voorafgaande modulecontrole voor `Microsoft.Graph.Authentication` en `Microsoft.Graph.Sites`, inclusief een duidelijke installatiehint wanneer modules ontbreken |
| De documentatie van moduleafhankelijkheden in `scripts/Startup/readme.md` bijgewerkt zodat `Microsoft.Graph.Sites` bij de installatie-/updatevereisten staat |

### 2026-07-22 (1)
| Wijziging |
|--------|
| `load.ps1` bijgewerkt — switches `-SetupStartup` / `-RemoveStartup` voor de opstartlauncher toegevoegd; de configuratie bij de eerste run slaat nu standaardwaarden voor gedelegeerde authenticatie (`authMode`, optioneel `defaultCustomerDomain`, `useDeviceCodeAuth`) op in `load.config.ps1` |
| `menu.ps1` bijgewerkt — Startup-acties `F` (Enable-LauncherStartup) en `G` (Disable-LauncherStartup) toegevoegd; M365-actie `H` (Test-GdapConnection) toegevoegd; het Entra-submenu bevat nu acties voor tijdelijke CA en TAP (`D`/`E`/`F`) |
| `scripts/Startup/functies.ps1` bijgewerkt — de Graph-verbinding bij het opstarten ondersteunt nu een voorkeur voor gedelegeerde device-auth + de vereiste scopes voor de GDAP-flow; helper `Test-GdapConnection` toegevoegd voor controles van het gedelegeerde contract en de connectiviteit |
| Onderhoud van de opstartmodules bijgewerkt: `scripts/Startup/Install-Modules.ps1` en `scripts/Startup/Update-Modules.ps1` bevatten nu `Microsoft.Graph.Identity.DirectoryManagement` |
| `scripts/Entra/New-TemporaryConditionalAccessPolicy.ps1` toegevoegd — maakt een tijdelijk CA-beleid voor een gebruiker/groep aan, met een venster op basis van een duur of een exacte lokale start-/einddatum en -tijd; optioneel automatisch opruimen op de eindtijd binnen dezelfde sessie |
| `scripts/Entra/Remove-TemporaryConditionalAccessPolicies.ps1` toegevoegd — verwijdert één of meer tijdelijke CA-beleidsregels (`TEMP-CA -`), inclusief modi voor alleen verlopen of alles verwijderen |
| `scripts/Entra/New-UserTemporaryAccessPass.ps1` toegevoegd — maakt een Temporary Access Pass (TAP) voor een gebruiker aan met instelbare geldigheidsduur en een optie voor eenmalig gebruik |
| De documentatie voor het bovenstaande bijgewerkt in `readme.md`, `scripts/readme.md`, `scripts/Startup/readme.md` en `scripts/Entra/readme.md`; verduidelijkt dat het automatisch opruimen van tijdelijke CA in de huidige sessie gebeurt (er wordt geen Scheduled Task aangemaakt) |

### 2026-07-09 (8)
| Wijziging |
|--------|
| `scripts/Reporting/Get-SharePointStorageReport.ps1` — een fase "totalen per sitecollectie" toegevoegd (alleen `-Apply`, fase 2c): subsites/Teams-kanalen en de prullenbak worden nu automatisch per root-sitecollectie opgeteld in `SharePoint_SiteCollectionTotals_<timestamp>.csv`, zodat het eindtotaal direct te vergelijken is met het ene cijfer "storage used" dat het SharePoint-beheercentrum per site toont |
| De nieuwe uitvoer en de meest waarschijnlijke oorzaken van een resterend verschil met het cijfer in de beheerportal (vertraging in de timing, mappen die bij rechtenfouten stilletjes worden overgeslagen, mislukte versie-lookups) gedocumenteerd in `scripts/Reporting/readme.md` |

### 2026-07-09 (7)
| Wijziging |
|--------|
| `scripts/Reporting/Get-SharePointStorageReport.ps1` — de prullenbak-lookups (fase 2b van `-Apply` en `-RecycleBinOnly`) kort uitgebreid naar ook persoonlijke OneDrive-sites, en dat op verzoek dezelfde dag weer teruggedraaid — de prullenbakscope blijft beperkt tot SharePoint-sitecollecties, OneDrive blijft volledig uitgesloten (zowel de opslagscan als de prullenbak) |
| Een sectie "Prullenbak (recycle bin)" toegevoegd aan `scripts/Reporting/readme.md` die de (alleen-SharePoint) prullenbakscope documenteert |

### 2026-07-09 (6)
| Wijziging |
|--------|
| De omhullende map `Testing Scripts/` volledig verwijderd — de submappen dupliceerden bestaande categorienamen op het hoogste niveau op basis van werkwoord (prefix `Test-`/`Get-`) in plaats van op domein. De inhoud is samengevoegd in de bijbehorende domeinmap: `Testing Scripts/Entra/Test-M365GroupMembership.ps1` → `Entra/`, `Testing Scripts/Exchange/*` (6 scripts) → `Exchange/`, `Testing Scripts/Device/Test-OpenVpnDiagnostics.ps1` → `Device/`. `Network/`, `RDS/` en `SMTP/` (zonder bestaande tegenhanger op het hoogste niveau) zijn in plaats daarvan gepromoveerd tot eigen categoriemappen op het hoogste niveau |
| De bijbehorende readmes samengevoegd in de bestaande readme.md van elke doelmap in plaats van aparte "Testing —"-documenten te behouden |
| 10 scriptpaden in `menu.ps1` (Exchange-auditsubmenu, Entra-auditsubmenu, Test-Ports, SMTP-tests) bijgewerkt naar de nieuwe locaties |

### 2026-07-09 (5)
| Wijziging |
|--------|
| `Testing Scripts/` en de root van de repo opgeruimd: `vias_archiver.ps1` uit `Testing Scripts/Device/` verplaatst naar een nieuwe categorie `scripts/Teams/` — het is een export- en archiveringstool voor Teams/SharePoint, geen diagnosescript, dus het hoorde niet onder "Testing" |
| `Update-modules.ps1` uit de root verplaatst naar `scripts/Startup/` (hernoemd naar `Update-Modules.ps1` voor consistente naamgeving) — het is een script voor moduleonderhoud zoals `Install-Modules.ps1`, geen startpunt van de repo zoals `load.ps1`/`menu.ps1` |
| De map `Testing Scripts/SharePoint/` verwijderd (die bevatte alleen een verwijzende readme, geen script) — die verwijzing staat nu direct in `Testing Scripts/readme.md` |
| `Test-PowerShellSyntax.ps1` gedocumenteerd in `scripts/Startup/readme.md`; daar was nog geen documentatie voor |

### 2026-07-09 (4)
| Wijziging |
|--------|
| De versiegeschiedenis-lookups van `scripts/Reporting/Get-SharePointStorageReport.ps1` geoptimaliseerd; die waren de hoofdoorzaak dat het script op grote bibliotheken leek te hangen (één sequentiële Graph-aanroep per bestand, elk met tot 6 retries en een backoff tot ~2 minuten bij throttling): (1) de lookup helemaal overslaan wanneer met zekerheid bekend is dat versiebeheer op een bibliotheek uitstaat, (2) tot 20 versie-lookups van bestanden per HTTP-aanroep bundelen via het `$batch`-endpoint van Graph in plaats van één aanroep per bestand, (3) voor deze specifieke aanroepen een korte, goedkope retry met 3 pogingen gebruiken in plaats van het hoofdbeleid voor retry/backoff, omdat een mislukte lookup veilig terugvalt op "0 versies" |
| De niet meer gebruikte functie `Get-VersionSize` verwijderd, vervangen door `Invoke-GraphBatchGet` + gebundelde resolutie in `Get-AllDriveItems` |
| `scripts/Reporting/readme.md` bijgewerkt met een sectie "Performance" die het bovenstaande documenteert |

### 2026-07-09 (3)
| Wijziging |
|--------|
| `Deploy-OfficeTheme.ps1`, `Deploy-Officecolors.ps1` en hun thema-assets (`2026 Vias institute colours (2).thmx`, `Office Themes/`) teruggezet naar `Custom Scripts/Intune/Desktop/` — beide scripts hebben hun download-URL hardcoded naar precies dat repopad, dus zo blijft de URL geldig zonder dat een scriptupdate + Intune-herimplementatie nodig is. `Custom Scripts/` en `Custom Scripts/Intune/` opnieuw aangemaakt als minimale mappen die alleen dienen om het pad vast te houden (alleen dit ene item) in plaats van de volledige vroegere categorie |
| `scripts/Intune/Desktop/` bevat nu alleen nog de uitrol van achtergrond/vergrendelscherm/taakbalksnelkoppelingen; zowel `Intune/readme.md` als `Intune/Desktop/readme.md` verwijzen voor de Office-themascripts naar `Custom Scripts/Intune/Desktop/` |

### 2026-07-09 (2)
| Wijziging |
|--------|
| De omhullende map `Custom Scripts/` verwijderd — die mengde generieke tooling met klantspecifieke scripts onder één verwarrend label en dupliceerde de categorie `Intune/`. De inhoud is herverdeeld over de juiste categorieën op het hoogste niveau: `Custom Scripts/device/` → `Device/`, `Custom Scripts/DNS/` → `DNS/`, `Custom Scripts/SAS/` → `SAS/`, `Custom Scripts/Save install time/` → `Deployment/` (hernoemd), `Custom Scripts/Intune/Desktop/` → samengevoegd in `Intune/Desktop/` |
| De scriptpaden in `menu.ps1` voor `Restart-Time-Sync.ps1`, `detect-audiodevices.ps1` en `Disable-internalmic.ps1` bijgewerkt naar hun nieuwe locatie in `scripts/Device/` |
| Kruisverwijzingen in `scripts/Intune/readme.md`, `scripts/Intune/Get-Autopilot/readme.md` en `scripts/readme.md` bijgewerkt naar de nieuwe maplocaties |
| ~~**Bekend probleem (bewust):** `Deploy-OfficeTheme.ps1` en `Deploy-Officecolors.ps1` hebben hun download-URL nog steeds hardcoded naar het oude pad — op verzoek ongewijzigd gelaten.~~ **Hierboven opgelost** — de scripts zijn in plaats daarvan teruggezet zodat ze weer overeenkomen met hun hardcoded URL. |

### 2026-07-09
| Wijziging |
|--------|
| Een `readme.md` toegevoegd aan elke map die er nog geen had: `scripts/`, `scripts/Custom Scripts/`, `scripts/Custom Scripts/Intune/` (+ `Desktop/`, `Office Themes/`, `Add Lockscreen to start and desktop/`, `Background/`), `scripts/Custom Scripts/device/Time sync/`, `scripts/Graph/`, `scripts/Intune/` (+ `Get-Autopilot/`), `scripts/Testing Scripts/`, `scripts/Testing Scripts/Network/`, `scripts/Testing Scripts/RDS/` — elk met een bestandslijst en documentatie van parameters/gebruik |
| `scripts/Custom Scripts/device/audio/Rollback-InternalMic` gecorrigeerd — het bestand miste de extensie `.ps1` |
| `scripts/Entra/remove-m365users.ps1` hernoemd naar `Remove-M365Users.ps1` voor consistente naamgeving (menu.ps1 en de readmes verwezen al naar de PascalCase-vorm) |
| `scripts/Entra/readme.md` gecorrigeerd — een verouderde vermelding `Distributionlist.ps1` verwijderd die in werkelijkheid `scripts/Exchange/Set-Distributionlist-dynamic-static.ps1` documenteerde; de juiste documentatie verplaatst naar `scripts/Exchange/readme.md`; ontbrekende documentatie voor `Set-UserManager.ps1` toegevoegd |
| `scripts/Testing Scripts/SharePoint/readme.md` gecorrigeerd — dit was een verouderd duplicaat van de documentatie van `Get-SharePointStorageReport.ps1` (het script staat niet in deze map); vervangen door een verwijzing naar `scripts/Reporting/readme.md`, dat nu de volledige huidige parameterset van het script documenteert (`-ClientId`, `-ClientSecret`, `-CertificateThumbprint`, `-RecycleBinOnly`, `-GraphTimeoutSec`, `-MaxGraphRetry` waren eerder niet gedocumenteerd) |
| `scripts/Custom Scripts/device/audio/readme.md` gecorrigeerd — het hoofdlettergebruik van de scriptnamen afgestemd op de werkelijke bestanden op schijf |
| Getrackte `.DS_Store`-bestanden uit git verwijderd en `.DS_Store` toegevoegd aan `.gitignore` |

### 2026-04-17
| Wijziging |
|--------|
| `scripts/Custom Scripts/Save install time/start.bat` bijgewerkt — optie `D` (installatiescripts van klanten uit de lokale map `Install`) en optie `E` (installatiescripts van klanten uit `\\10.222.3.94\Software`) toegevoegd; voordat de uitrol begint, wordt de lokale admin `LocalAdmin` (`<wachtwoord weggelaten>`) aangemaakt/bijgewerkt, aan `Administrators` toegevoegd en worden de OOBE-skipvlaggen gezet |
| `scripts/Custom Scripts/Save install time/Browse-InstallScripts.ps1` toegevoegd — een browser die bij de klant begint, klantmappen als menu-items toont en `.ps1`-, `.bat`- en `.cmd`-scripts start |
| `scripts/Custom Scripts/Save install time/readme.md` bijgewerkt — de nieuwe menuopties `D`/`E` gedocumenteerd, inclusief dat je voor optie `D` zowel `Browse-InstallScripts.ps1` als de volledige map `Install` moet kopiëren, en dat de uitrolopties voor klanten `LocalAdmin` plus de OOBE-skipvlaggen voorbereiden |

### 2026-04-16
| Wijziging |
|--------|
| `scripts/Custom Scripts/Intune/Desktop/Background/Lockscreen/Make-lockscreen.ps1` bijgewerkt naar v2.0 — de bron van het vergrendelscherm afgestemd op de configuratie van de bedrijfsachtergrond (`$ImageUrl`, `$ClientName`), de directe `WebClient` vervangen door een gevalideerde internetdownloadflow (`Invoke-WebRequest`), normalisatie van GitHub-blob-/raw-URL's toegevoegd, controles op de afbeeldingssignatuur (`jpg/png/bmp`), een beveiliging tegen HTML-responses, gestructureerde Intune-logging en een veiligere afhandeling van tijdelijke downloads |
| `scripts/Custom Scripts/Intune/Desktop/Background/Lockscreen/readme.md` toegevoegd — documentatie voor configuratie, uitrol, logging, workflow en een versiegeschiedenis specifiek voor het vergrendelscherm |
| Root-`readme.md` bijgewerkt — de documentatie van Intune Desktop/Background en de repositorystructuur uitgebreid met het vergrendelschermscript en de bijbehorende documentatie |

### 2026-04-15
| Wijziging |
|--------|
| `scripts/Custom Scripts/SAS/Test-SASWorkDirectory.ps1` bijgewerkt — diagnose van het schijfprofiel toegevoegd met expliciete herkenning van ephemeral schijven (`G:`/`U:`), gedetailleerdere logging bij I/O-fouten (exceptiontype, inner exception, HResult) en een adaptieve drempel voor de waarschuwing bij weinig ruimte op ephemeral scratchschijven |
| `scripts/Custom Scripts/SAS/Test-SASWorkDirectory.ps1` bijgewerkt — de analyse van SAS-events in het Application-log verfijnd: `hc_disk_delete*`-access-denied (`Return code 5`) wordt expliciet geclassificeerd als fout waar actie op nodig is, herhaalde SAS-events worden gededupliceerd, en de telemetrieruis `ARM Application data not available` wordt als informatief behandeld |
| `scripts/Custom Scripts/SAS/Test-SASWorkDirectory.ps1` bijgewerkt — AV/EDR-diagnose toegevoegd: Defender-status + uitsluitingen, scan van Defender Operational-events, scan van FilterManager-/WdFilter-events in het System-log en een momentopname van actieve minifilters (`fltmc`), om te correleren met de sporadische access-denied-fouten bij het verwijderen van WORK |
| `scripts/Custom Scripts/SAS/readme.md` bijgewerkt — het gedrag van ephemeral schijven voor SAS `WORK`/`USERWORK` gedocumenteerd en verduidelijkt waarom wisselen tussen `G:` en `U:` geen failoverstrategie voor de lange termijn is wanneer beide ephemeral zijn |
| `scripts/Custom Scripts/SAS/rca.md` toegevoegd — een formele rootcauseanalyse van de sporadische fouten bij het verwijderen van SAS WORK, inclusief een tijdlijn van het bewijs, de bevindingen over AV/ASR, de beoordeling van de hoofdoorzaak en een herstelplan |

### 2026-04-13
| Wijziging |
|--------|
| `scripts/Testing Scripts/Device/vias_archiver.ps1` bijgewerkt naar v8.11 — stap 10 ondersteunt nu een niet-interactieve modus via `-Step10Only -Step10Action undo|archive|skip`; een expliciete waarschuwing toegevoegd dat archiveren/dearchiveren in Teams een actie op teamniveau is (niet per kanaal) |
| `scripts/Testing Scripts/Device/vias_archiver.ps1` bijgewerkt naar v8.12 — een soft-archive-modus per kanaal toegevoegd in stap 10 (hernoemen met een markering, met undo), inclusief de parameters voor de snelle modus `-Step10Only -ChannelAction archive|undo` en optioneel `-ChannelArchiveTag` |
| `scripts/Testing Scripts/Device/vias_archiver.ps1` bijgewerkt naar v8.13 — de flow per kanaal overgezet op de echte Graph-API voor het archiveren/dearchiveren van kanalen, de kanaalscope `ChannelSettings.ReadWrite.All` toegevoegd en een optionele fallback naar hernoemen via `-ChannelFallbackToRename` |
| `scripts/Testing Scripts/Device/vias_archiver.ps1` bijgewerkt naar v8.14 — een `-DryRun`-modus toegevoegd voor stap 10, zodat archiveren/dearchiveren van teams/kanalen (en de optionele fallback naar hernoemen) gesimuleerd kan worden zonder wijzigingen aan te brengen |
| `scripts/Testing Scripts/Device/vias_archiver.ps1` bijgewerkt naar v8.15 — `-DryRun` uitgebreid naar gedrag voor het hele script: slaat wijzigende setup-/export-/rapport-/opruimacties over, terwijl de verificatie en de gesimuleerde uitvoer van stap 10 behouden blijven |
| `scripts/Testing Scripts/Device/vias_archiver.ps1` bijgewerkt naar v8.16 — `-DryRun` laat het aanmaken van de tijdelijke app, de bootstrap van de rechten, de volledige aanmelding en de export-/rapportflow nu actief; alleen de archiveer-/dearchiveerwijzigingen van stap 10 blijven gesimuleerd |
| `scripts/Testing Scripts/Device/vias_archiver.ps1` bijgewerkt naar v8.17 — `-DryRun` voor stap 6-9 verfijnd zodat bestaan/aantallen (Teams/SharePoint/Graph) worden gevalideerd zonder exports te schrijven; het rapport van stap 11 gebruikt nu deze gepeilde aantallen |
| `scripts/Testing Scripts/Device/vias_archiver.ps1` bijgewerkt naar v8.18 — de betrouwbaarheid van het opzoeken van kanalen verbeterd (Team-/Channel-waarden uit Excel getrimd en de gecachete Graph-fallback voor kanaalresolutie hergebruikt in stap 9) om onterechte "Kanaal niet gevonden"-meldingen in de proefdraai te verminderen |
| `scripts/Testing Scripts/Device/vias_archiver.ps1` bijgewerkt naar v8.19 — genormaliseerde matching van kanaalnamen (trim/witruimte/hoofdletters) toegevoegd in de kanaalcache + de Graph-fallbacklookup, om subtiele naamverschillen tijdens een proefdraai beter af te handelen |

### 2026-04-09
| Wijziging |
|--------|
| `scripts/Custom Scripts/Intune/Desktop/Background/Desktop/Set-CorporateWallpaper.ps1` v2.1 bijgewerkt — `Set-ExecutionPolicy Bypass -Scope Process` bovenaan toegevoegd om exitcode 3 te voorkomen wanneer de execution policy van Intune het script blokkeert |

### 2026-04-08
| Wijziging |
|--------|
| `scripts/Reporting/Licensing/genereer_licentie_overzicht.py` bijgewerkt — wanneer hetzelfde product met meerdere factuurperiodes op één factuur staat (Pax8 en Ingram), wordt elke periode nu als aparte rij getoond met de periode in de kolom Category/Detail, in plaats van dat ze onjuist worden opgeteld |
| `scripts/Reporting/Licensing/genereer_rapport.ps1` bijgewerkt — het CMD-venster sluit nu automatisch wanneer het als geplande taak draait; `Read-Host`-pauzes worden overgeslagen wanneer `[Environment]::UserInteractive` false is |
| `scripts/Reporting/Licensing/genereer_rapport.bat` bijgewerkt — stdin wordt vanuit `NUL` doorgesluisd, zodat de interactieve pauze van Python nooit wordt geactiveerd wanneer het als geplande taak draait |

### 2026-04-01
| Wijziging |
|--------|
| `scripts/Exchange/Migrate-Calendar.ps1` bijgewerkt — boekingsproblemen met Room Mailboxes opgelost: de wachttijd voor provisioning verhoogd van 15s naar 60s; een retrylus (5×30s) toegevoegd voor `Set-CalendarProcessing` met foutafhandeling en fallback-instructies; `BookingWindowInDays 0` gewijzigd in `1825` en `EnforceSchedulingHorizon $false` toegevoegd om stille afwijzingen van boekingen te voorkomen |

### 2026-03-30
| Wijziging |
|--------|
| `scripts/Custom Scripts/Intune/Desktop/Background/Desktop/Set-CorporateWallpaper.ps1` bijgewerkt — een idempotentiecontrole toegevoegd: downloadt de afbeelding naar temp en vergelijkt de SHA256-hash met het bestaande bestand; slaat over als de hash overeenkomt en PersonalizationCSP correct is; past toe (zonder tweede download) als de afbeelding nieuw of gewijzigd is |
| Readme bijgewerkt — sectie Intune & Autopilot: de documentatie van `Set-CorporateWallpaper.ps1` uitgebreid met een configuratietabel, uitrolstappen, het logpad en uitrolinstructies voor NinjaOne/Intune |

### 2026-03-27
| Wijziging |
|--------|
| `scripts/Reporting/Get-ComputerLastLogon.ps1` toegevoegd — rapport van de laatste aanmelding voor computers in één of meer OU's; modus `LastLogonTimestamp` (snel) of `-AllDCs` (nauwkeurig); markeert Active/Stale/Never/Disabled; exporteert een CSV met tijdstempel naar `C:\Temp\`; parameters `-InactiveDays`, `-IncludeDisabled`, `-ExportPath` |
| `scripts/Reporting/readme.md` toegevoegd — documenteert `Get-ComputerLastLogon.ps1` met een parametertabel, een overzicht van de CSV-kolommen en gebruiksvoorbeelden |
| `scripts/Reporting/Get-ComputerLastLogon.ps1` bijgewerkt — kolommen `PasswordLastSet` / `DaysSincePasswordSet` toegevoegd; nieuwe status `Active (pwd recent)` voor apparaten die onterecht als stale werden gemarkeerd door de replicatievertraging van 14 dagen van `LastLogonTimestamp` |

### 2026-03-26
| Wijziging |
|--------|
| `scripts/Custom Scripts/device/audio/Disable-internalmic.ps1` bijgewerkt — de patronen voor interne microfoons uitgebreid: Conexant-taalvarianten (EN/FR/NL), Synaptics EN, Intel SST, IDT, Cirrus Logic-drivers en meertalige namen voor microfoonarrays (FR/DE/ES/PT/IT) toegevoegd |
| Readme bijgewerkt — sectie Audio Management: een uitroltabel voor NinjaOne toegevoegd (Run as SYSTEM, geen parameters, custom field `AudioDeviceInventory`, exitcodes) voor alle drie de audioscripts |
| Readme bijgewerkt — `Invoke-WindowsActivation.ps1`: een uitroltabel voor NinjaOne toegevoegd met Run as Administrator, exitcodes en parametervoorbeelden per scenario; een waarschuwing toegevoegd dat `-RemoveKey`/`-ReArm` `-Force` vereisen |
| `scripts/Testing Scripts/RDS/Watch-RDSLive.ps1` toegevoegd — realtime RDS-monitor: pollt elke 20s sessie-events (20/21/22/23/24/25/40), mislukte RDP-aanmeldingen (4625), lockouts (4740) en licentie-events; heartbeat per poll met het aantal sessies; draai het direct op elke RDS-server |

### 2026-03-25
| Wijziging |
|--------|
| `scripts/Testing Scripts/Network/Test-FileIODiagnostics.ps1` bijgewerkt — realtime monitor samengevoegd: FileSystemWatcher, verschil in NTFS-rechten ten opzichte van een baseline, automatische download van Sysinternals Handle.exe, verschil tussen procesmomentopnames, Kerberos-tickets op het moment van de fout, Security-auditevents (4625/4740/4656/4663/4670); stopt na 3 fouten |
| `scripts/Testing Scripts/Network/Test-FileIODiagnostics.ps1` toegevoegd — bestands-I/O-stresstest op elk pad (lokaal of UNC/gekoppelde schijf); deelt fouten in als AUTH / NETWORK / TIMEOUT / DISK / PATH; verzamelt bij de eerste fout automatisch Kerberos-tickets, net use, een SMB-poortcontrole en het Security-eventlog; parameters `-Iterations`, `-StopOnFirstError`, `-DelayMs` |
| `scripts/Testing Scripts/RDS/Test-RDSDiagnostics.ps1` bijgewerkt — analyse van RDWeb in Event Viewer toegevoegd: TerminalServices-WebAccess/Admin+Operational, TerminalServices-Gateway/Admin+Operational, IIS-/ASP.NET-fouten uit het Application-log; geactiveerd door -IncludeEventLogs |
| `scripts/Testing Scripts/RDS/Test-RDSDiagnostics.ps1` toegevoegd — diagnosticeert mislukte RDP-/RDWeb-aanmeldingen: services, register, NLA, sessielimieten, licenties, firewall, HTTPS-certificaat, IIS-app-pool, gebruikersaccount (ingeschakeld/vergrendeld/verlopen/groep), eventlogs (4625/4740/4771/20/40); log met tijdstempel naar C:\Temp\ |
| `scripts/Custom Scripts/device/Invoke-WindowsActivation.ps1` toegevoegd — Windows activeren, productcode installeren, KMS-server/-poort configureren, code verwijderen, respijtperiode resetten met ReArm; veilig als proefdraai met bevestigingsvragen; -Force om die over te slaan |
| `scripts/Custom Scripts/device/Invoke-WindowsCleanup.ps1` bijgewerkt — dynamische scan van logmappen toegevoegd (sectie 13): doorzoekt recursief de hele C:\-schijf (maximale diepte 7) naar mappen met de naam logs/log/logging/diagnostics; slaat Windows-systeemmappen en ontwikkelartefacten (node_modules, .git, venv) over; vaste paden voor Windows-systeemlogs (gearchiveerde CBS-.cab, DISM, WU, Panther, IIS); nieuwe parameter `-SkipAppLogs` |
| `scripts/Custom Scripts/device/Invoke-WindowsCleanup.ps1` bijgewerkt — scant alle gebruikersprofielen in C:\Users\ op temp, WER, thumbnail-/shadercache en browsercaches (Edge met meerdere profielen, Chrome met meerdere profielen, Firefox); de samenvatting toont de terug te winnen ruimte per categorie |
| `scripts/Custom Scripts/device/Invoke-WindowsCleanup.ps1` gecorrigeerd — waarschuwingen van de PS-analyzer opgelost: `$profile` hernoemd naar `$ffProfile`, de ongebruikte toewijzing `$dismResult` geschrapt |
| `scripts/Custom Scripts/device/Invoke-WindowsCleanup.ps1` gecorrigeerd — de null-coalescing-operator `??` vervangen door `-as [int64]` voor compatibiliteit met PowerShell 5.1 |
| `scripts/Custom Scripts/device/Invoke-WindowsCleanup.ps1` toegevoegd — uitgebreide opschoning van de Windows-schijf: tijdelijke bestanden, WU-cache, Delivery Optimization, Prefetch, geheugendumps, WER, thumbnail-/shadercache, Prullenbak, browsercaches, eventlogs, DISM-componentstore; standaard een proefdraai, `-Apply` om uit te voeren |
| `Get-SharePointStorageReport.ps1` bijgewerkt — aanpak in twee fasen (eerst alle sites/bibliotheken opsommen, daarna de opslaggegevens ophalen); syntaxfout in een inline `if` opgelost |
| `Test-AuthNetworkDiagnostics.ps1` toegevoegd — diagnose van authenticatie en netwerk: Event Viewer (4625/4771/4776/4740/5719), tijdsynchronisatie, Kerberos-cache, DNS, TCP, UNC-shares, optionele logscan |
| `scripts/Custom Scripts/SAS/` toegevoegd — monitoring van SAS-batchfouten met Zabbix-integratie, e-mailwaarschuwingen en analyse in Event Viewer |

### 2026-03-24
| Wijziging |
|--------|
| `scripts/Intune/iOS-Compliance-Updater/` toegevoegd — werkt de minimale iOS-versie in een Intune-compliancebeleid automatisch bij via de Graph API; wekelijkse geplande taak, ondersteuning voor proefdraai, eenmalige Setup.ps1 |
| `Get-SharePointStorageReport.ps1` toegevoegd — tenantbreed SharePoint-opslagrapport met versiegeschiedenis per bestand; snelle modus (quota) en volledige recursieve scan |
| `Test-OpenVpnDiagnostics.ps1` toegevoegd — diagnose van OpenVPN Connect: PnP-adapters, services, routes, DNS, Event Log, conflicterende VPN-software; txt-export naar `C:\Temp\` |

### 2026-03-23
| Wijziging |
|--------|
| Alle CSV-exports gaan nu naar `C:\Temp\` (Windows) of `~/Downloads/` (macOS/Linux) |
| `Import-DnsRecords.ps1` toegevoegd — lost publieke DNS op via dig (Google 8.8.8.8) en importeert A-/CNAME-records in AD DNS, standaard een proefdraai |
| `New-M365User.ps1` toegevoegd — maakt één M365-gebruiker aan via Graph, met automatisch gegenereerd wachtwoord en optionele licentie |
| `Import-M365Users.ps1` toegevoegd — bulksgewijs gebruikers aanmaken vanuit CSV via Graph, standaard een proefdraai, wachtwoorden in de CSV-uitvoer |
| `Remove-M365Users.ps1` toegevoegd — bulksgewijs Entra ID-gebruikers verwijderen, standaard een proefdraai, CSV-rapport |
| `Get-ExternalForwards.ps1` toegevoegd — audit van externe doorstuurregels over alle mailboxen, CSV-export |
| `Get-MailboxSizes.ps1` toegevoegd — rapport van mailboxgrootte + aantal items, gesorteerd op opslag, CSV-export |
| `Test-DkimConfig.ps1` toegevoegd — DKIM-ondertekeningsconfiguratie + validatie van DNS-CNAME/TXT, met uitvoer van de vereiste acties |
| `Test-M365GroupMembership.ps1` toegevoegd — audit van eigenaren en leden van M365-groepen / Teams via Graph, CSV-export |
| `Test-DistributionGroupPermissions.ps1` toegevoegd — beheerders van distributiegroepen, Send As, Send on Behalf, aantallen leden, CSV-export |
| `Test-CalendarPermissions.ps1` toegevoegd — taalonafhankelijke audit van agendarechten, CSV-export |
| `Test-MailboxPermissions.ps1` toegevoegd — audit van Full Access / Send As / Send on Behalf, CSV-export |
| Exchange-submenu (`C`) en Entra-submenu (`D`): tools voor rechtenaudits toegevoegd |
| `Test-Ports.ps1` verplaatst naar `scripts/Testing Scripts/Network/` |
| `load.ps1` — importeert modules automatisch bij het opstarten; detecteert ontbrekende modules en biedt aan ze te installeren |
| `load.ps1` toegevoegd — setup bij de eerste run (UPN + naam), slaat op in het door git genegeerde `load.config.ps1`, start het menu |
| `menu.ps1` uitgebreid met een M365-sectie (B–E): submenu's voor Exchange, Entra ID en MSP Admin |
| `menu.ps1` toegevoegd — interactieve launcher, met één toetsaanslag, cijfers + F-toetsen, cross-platform |
| `scripts/Network/Test-Ports.ps1` toegevoegd — TCP-poortcontrole, syntaxis voor bereiken/lijsten, meerdere doelen |
| Licentiescripts: vertaald naar het Engels, generiek gemaakt, exportmap instelbaar |
| `create_scheduled_task.ps1` herschreven — controle op adminrechten, Python automatisch detecteren, dynamische startdatum van de trigger |
| `scripts/Reporting/Licensing/` toegevoegd — toolkit voor Pax8 + Ingram → Excel-rapport |

### 2026-03-20
| Wijziging |
|--------|
| `start.bat` v2.8 — Do it all gesplitst: A = Intune, C = AD; apparaat hernoemen (B) en AD-join (8) toegevoegd |
| `Migrate-Calendar.ps1` v2.0 herschreven — Engels, generiek, verplichte parameters |
| SMTP-testscripts toegevoegd; `Test-SmtpRelay` toegevoegd aan `functies.ps1` |
| `functies.ps1` herschreven — MSOnline/AzureAD vervangen door Microsoft Graph, cross-platform |
| `Set-CorporateWallpaper.ps1`, `Set-Calendar-rights.ps1`, `Restart-Time-Sync.ps1` toegevoegd |
| Alle bedrijfsspecifieke verwijzingen verwijderd; alle readmes naar het Engels vertaald |

### 2026-03-19
| Wijziging |
|--------|
| Eerste upload van de repository |

---

## Beheerder

**Sjoerd Kanon** — Securitybewust | Team- & projectgericht | Microsoft 365 & infrastructuur
