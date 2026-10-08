[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../readme.nl.md) › [scripts](../readme.nl.md) › **Startup**

# Startup

Startscripts en de centrale M365-functiebibliotheek.

---

## Bestanden

| Bestand | Omschrijving |
|------|-------------|
| [`functies.ps1`](functies.ps1) ([docs](#functiesps1)) | M365-functiebibliotheek — bij eerste gebruik door `menu.ps1` gedot-sourced |
| [`RequiredModules.psd1`](RequiredModules.psd1) ([docs](#requiredmodulespsd1)) | De ene lijst met modules die deze repo nodig heeft — gelezen door `load.ps1`, `Install-Modules.ps1` en `Update-Modules.ps1` |
| [`Test-RequiredModules.ps1`](Test-RequiredModules.ps1) ([docs](#test-requiredmodulesps1)) | Meldt modules die scripts laden maar die niet in `RequiredModules.psd1` staan — de docs-hook draait het na elke wijziging |
| [`Connect-M365.ps1`](Connect-M365.ps1) ([docs](#connect-m365ps1)) | De ene manier waarop scripts aanmelden: Graph eerst, standaard delegated (device code en GDAP uit `load.config.ps1`), app-only op verzoek — door de scripts gedot-sourcet |
| [`Install-Modules.ps1`](Install-Modules.ps1) ([docs](#install-modulesps1)) | Bootstrapscript — installeert en importeert alle benodigde PowerShell-modules |
| [`Update-Modules.ps1`](Update-Modules.ps1) ([docs](#update-modulesps1)) | Controleert de vereiste modules (ontbrekend, te oud, update beschikbaar) en installeert/updatet ze; werkt desgewenst ook alle andere geïnstalleerde modules bij |
| [`Test-PowerShellSyntax.ps1`](Test-PowerShellSyntax.ps1) ([docs](#test-powershellsyntaxps1)) | Controleert `.ps1`-bestanden in de repo op syntaxfouten door ze te parsen, zonder ze uit te voeren |
| [`Update-ScriptIndex.ps1`](Update-ScriptIndex.ps1) ([docs](#update-scriptindexps1)) | Genereert [`scripts/INDEX.md`](../INDEX.md) opnieuw — de doorzoekbare A–Z-lijst van alle scripts |
| [`Test-MarkdownLinks.ps1`](Test-MarkdownLinks.ps1) ([docs](#test-markdownlinksps1)) | Controleert elke link in elke readme — bestanden die moeten bestaan, anchors die met een kop moeten overeenkomen |
| [`Convert-MarkdownToHtml.ps1`](Convert-MarkdownToHtml.ps1) ([docs](#convert-markdowntohtmlps1)) | Bouwt van een markdowndocument een op zichzelf staande, opgemaakte HTML-pagina — om in IT Glue te plakken of af te drukken |
| [`Update-ReadmeHeader.ps1`](Update-ReadmeHeader.ps1) ([docs](#update-readmeheaderps1)) | Zet de taalwissel en het kruimelpad bovenaan elke readme, in het Engels, Nederlands en Frans |

---

## Gedelegeerde GDAP-start

`load.ps1` kan nu gedelegeerde standaardwaarden opslaan in `load.config.ps1`:

- `authMode` (`GDAP` of `Direct`)
- `defaultCustomerDomain` (optioneel)
- `useDeviceCodeAuth` (`$true` / `$false`)

Als `authMode` op `GDAP` staat en er een `defaultCustomerDomain` is ingesteld, voert het menu na het laden van `functies.ps1` automatisch `Connect-Tenant` uit.

De launcher registreren bij het aanmelden in Windows:

```powershell
.\load.ps1 -SetupStartup
```

Weer verwijderen:

```powershell
.\load.ps1 -RemoveStartup
```

Bij elke start, voordat het menu opent, controleert `load.ps1` de modules uit [`RequiredModules.psd1`](#requiredmodulespsd1) met [`Update-Modules.ps1`](#update-modulesps1): het toont wat ontbreekt, ouder is dan het minimum of achterloopt op de PowerShell Gallery, en installeert of updatet dat meteen, zonder te vragen. De gallery wordt hooguit eens per 24 uur bevraagd, dus een normale start kost minder dan een seconde. De controle één keer overslaan:

```powershell
.\load.ps1 -SkipModuleCheck
```

Je kunt het automatisch starten ook in het launchermenu aan- en uitzetten:

- `F` = Enable-LauncherStartup
- `G` = Disable-LauncherStartup

---

## functies.ps1

Centrale functiebibliotheek voor multi-tenant M365-beheer via Microsoft Graph en Exchange Online. Wordt automatisch door het menu geladen bij het eerste gebruik van een optie B–E.

### Configuratie

Pas het blok `#region Configuration` bovenaan aan je eigen organisatie aan:

```powershell
$script:MspAdminAlias       = 'msp-admin'
$script:MspAdminDisplayName = 'MSP - Admin Account'
```

### Een klanttenant selecteren

```powershell
Connect-Tenant -Domain "customer.com"
# Zet $global:cid en $global:connectmsoldomain
# Onder GDAP richten alle volgende functies zich op de geselecteerde tenant
```

### Aanmelden

Elke functie meldt zich aan via [`Connect-M365.ps1`](#connect-m365ps1), dat `functies.ps1` dot-sourcet: Microsoft Graph, gedelegeerd, met een apparaatcode als `useDeviceCodeAuth` in `load.config.ps1` aan staat, anders in de browser. Een sessie die de scopes van een functie al heeft, wordt hergebruikt.

- **GDAP** (`authMode = 'GDAP'`): na `Connect-Tenant` verbindt elke Graph- en Exchange-functie met die klant (`$cid`, Exchange via `-DelegatedOrganization`).
- **Direct**: de functies blijven in je eigen tenant; `$cid` wordt niet gebruikt.
- De partnersessie die `Connect-Tenant` en `Test-GdapConnection` nodig hebben voor `Get-MgContract`, opent `Connect-PartnerGraph`. Die onthoudt bij de eerste start je eigen tenant (`$global:partnerTenantId`), zodat een andere klant kiezen ook nog werkt nadat een functie naar de huidige klant is overgeschakeld.
- Teams (`Invoke-Menu` optie 3) loopt via `Connect-M365Teams`, Exchange via `Connect-M365Exchange`.

### Functies

**Verbinding**

| Functie | Omschrijving |
|----------|-------------|
| `Connect-Tenant` | Selecteert een CSP-klant op domein, vult `$cid` en `$connectmsoldomain` |
| `Test-GdapConnection` | Valideert het gedelegeerde GDAP/CSP-contract en probeert daarna een gedelegeerde Graph-verbinding met de klant (`Get-MgOrganization`) en een gedelegeerde Exchange-verbinding |
| `Test-ExoConnection` | Verbindt met Exchange Online, of hergebruikt een sessie met de juiste organisatie |
| `Connect-PartnerGraph` | Graph in je eigen (partner)tenant, voor `Get-MgContract` |

**Exchange Online**

| Functie | Omschrijving |
|----------|-------------|
| `Enable-CopyOfSentItems` | Zet het bewaren van een kopie van verzonden items aan voor alle mailboxen |
| `Add-SharedMailboxAccess` | Kent FullAccess + SendAs toe op een gedeelde mailbox |
| `Set-MailboxLocale` | Stelt taal en tijdzone in op alle mailboxen (standaard: NL / W. Europe) |
| `Add-MailboxAlias` | Voegt een alias toe aan een mailbox |
| `Get-MailboxAliases` | Toont alle SMTP-aliassen per mailbox |
| `Export-DistributionGroups` | Exporteert alle distributiegroepen naar CSV (`C:\Temp\`) |
| `Set-AutoReply` | Stelt een afwezigheidsbericht in |

**Entra ID / Graph**

| Functie | Omschrijving |
|----------|-------------|
| `Get-TenantAdmins` | Toont alle Global Administrators |
| `Add-TenantDomain` | Voegt een domein toe en loopt de verificatie door |
| `Get-TenantLicenses` | Toont een licentieoverzicht met gebruik en beschikbaarheid |
| `Get-TenantUsers` | Toont alle gebruikers met UPN, weergavenaam en licenties |
| `Add-TenantAdmin` | Kent een gebruiker Global Administrator-rechten toe |
| `Get-EntraApplication` | Zoekt een Enterprise App op naam |
| `Reset-UserPassword` | Stelt het wachtwoord van een gebruiker opnieuw in |
| `Export-SignInLogs` | Exporteert aanmeldlogboeken naar CSV in `C:\Temp\` (standaard: laatste 30 dagen) |

**MSP-beheeraccount**

| Functie | Omschrijving |
|----------|-------------|
| `New-MspAdmin` | Maakt het MSP-beheeraccount aan als Global Admin in de klanttenant |
| `Set-MspAdminAsGroupOwner` | Maakt het MSP-beheeraccount eigenaar van een groep |
| `Reset-MspAdminPassword` | Stelt het wachtwoord van het MSP-beheeraccount opnieuw in |

---

## RequiredModules.psd1

De modules waar deze repository van afhangt, in één PowerShell-databestand. `load.ps1`,
`Install-Modules.ps1` en `Update-Modules.ps1` lezen het allemaal, zodat ze het niet meer
oneens kunnen zijn — voorheen had elk een eigen lijst, en `load.ps1` controleerde er maar zeven.

**Een module toevoegen:** voeg hier een regel toe. Bij de volgende start van `load.ps1` ziet
elke machine hem als ontbrekend en installeert hem. Een `MinimumVersion` verhogen werkt op
dezelfde manier. Vergeten is lastig: [`Test-RequiredModules.ps1`](#test-requiredmodulesps1)
draait na elke wijziging en meldt een module die een script laadt maar die niet in dit bestand staat.

| Sleutel | Betekenis |
|---------|-----------|
| `Name` | Modulenaam op de PowerShell Gallery |
| `MinimumVersion` | Ouder dan dit telt als *te oud* (niet alleen *update beschikbaar*) |
| `WindowsOnly` | Overgeslagen op macOS en Linux |
| `MinimumPSVersion` | Overgeslagen op een oudere PowerShell — `PnP.PowerShell` 3 vereist 7.4 |
| `ImportAtStartup` | Door `load.ps1` geïmporteerd voordat het menu opent |

Naast `Modules` staat `NotManaged`: modules die scripts laden maar die bewust niet uit de gallery worden geïnstalleerd, elk met de reden — `ActiveDirectory` en `WebAdministration` (Windows-onderdelen), `AzureAD` (uitgefaseerd), `Microsoft.Graph` (de hele SDK, alleen genoemd in installatietips).

Huidige lijst: `ExchangeOnlineManagement`, de Graph-submodules `Authentication`, `Sites`,
`Identity.DirectoryManagement`, `Identity.SignIns`, `Identity.Governance`, `Applications`,
`Calendar`, `Groups`, `Users`, `Reports`, plus `PnP.PowerShell`, `MicrosoftTeams`, `ImportExcel`,
`Az.Accounts`, `Az.OperationalInsights`, `DCToolbox`, `IntuneBackupAndRestore`, en op Windows
`WindowsAutopilotIntune` en `IntuneWin32App`.

---

## Test-RequiredModules.ps1

Een script dat een nieuwe module gaat gebruiken werkt op de machine waarop het geschreven is,
en faalt overal anders tot de module in [`RequiredModules.psd1`](#requiredmodulespsd1) staat.
Dit script doorzoekt elke `.ps1`/`.psm1` op `#Requires -Modules`, `Import-Module` en
`Install-Module` met een letterlijke naam, en meldt elke naam die niet in `Modules` of
`NotManaged` staat, met de bestanden die hem gebruiken. Een naam in een variabele (`$mod`) kan
het niet controleren.

De docs-hook (`.claude/hooks/sync-docs.ps1`) draait het na elke wijziging van een `.ps1`, `.psd1`
of `.md`, en de git pre-commit-hook waarschuwt ermee, zodat een nieuwe module opvalt zodra het
script wordt opgeslagen, niet pas als het bij iemand anders faalt.

**Parameters**

| Parameter | Beschrijving |
|-----------|--------------|
| `-Root` | Root van de repository (standaard: twee niveaus boven dit script) |

**Voorbeelden**

```powershell
pwsh -File scripts/Startup/Test-RequiredModules.ps1
```

Exitcodes: `0` = elke module die een script laadt staat in de lijst, `1` = er ontbreekt iets.

---

## Connect-M365.ps1

De aanmelding die elk script gebruikt. **Microsoft Graph is de standaard**; Exchange Online,
Teams en PnP worden alleen verbonden voor werk waar Graph geen API voor heeft (mailbox- en
SendAs-rechten, message trace, DKIM, EOP-beleid, Teams `Cs*`-beleid, SharePoint-rolverdeling, ...).

```powershell
. (Join-Path $PSScriptRoot '..\Startup\Connect-M365.ps1')   # diepte hangt af van de map van het script
$graph = Connect-M365Graph -Scopes 'User.Read.All' -TenantId $TenantId
# ... werk ...
Disconnect-M365Graph $graph    # verbreekt alleen wat deze aanroep verbond
```

| Functie | Wat het doet |
|---------|--------------|
| `Connect-M365Graph` | Microsoft Graph. `-Scopes`, `-TenantId`, `-ClientId` + `-CertificateThumbprint`/`-ClientSecret`, `-AppOnly`, `-DeviceCode`, `-Interactive`, `-Force`; `-Force` meldt opnieuw aan ook als de sessie zou passen. Veilig onder `Set-StrictMode`, ook zonder `load.ps1` |
| `Disconnect-M365Graph` | Verbreekt alleen als `Connect-M365Graph` de sessie opende |
| `Connect-M365Exchange` | Exchange Online, `-IncludeCompliance` voegt Security & Compliance toe (`Connect-IPPSSession`), `-EnableSearchOnlySession` voor Content Search |
| `Disconnect-M365Exchange` | Sluit alleen de sessies die `Connect-M365Exchange` opende (op connection id), nooit die van de aanroeper |
| `Connect-M365Teams` / `Disconnect-M365Teams` | Microsoft Teams PowerShell, voor het `Cs*`-beleid |
| `Connect-M365PnP` | PnP.PowerShell naar een site; geeft de verbinding terug. ClientId uit `-ClientId` of `pnp.appid.json`; `-AppOnly` gebruikt de certificaat-app uit `graph.appid.json` |
| `Invoke-M365GraphPaged` | Een Graph-collectie ophalen en `@odata.nextLink` tot het einde volgen |
| `Resolve-M365TenantId` | De tenant: `-TenantId`, anders de GDAP-klant, anders je eigen tenant |

**Hoe het aanmeldt**

- **Delegated, de standaard.** Je meldt je aan als jezelf, met een device code als
  `useDeviceCodeAuth` in `load.config.ps1` aan staat (of `-DeviceCode` is meegegeven), anders
  in de browser met je `upn` al ingevuld. Onder GDAP (`authMode = 'GDAP'`) is de klanttenant
  `$global:cid` / `$global:connectmsoldomain` uit `Connect-Tenant`, of
  `$env:M365_CUSTOMER_TENANTID`. Exchange bereikt de klant met `-DelegatedOrganization`;
  `-Organization` werkt alleen bij app-only aanmelden. Buiten GDAP komt een delegated
  Exchange-aanmelding uit in de tenant van het account waarmee je aanmeldt.
- **App-only, op verzoek.** `-ClientId` met `-CertificateThumbprint` (of `-ClientSecret`,
  alleen Graph), of `-AppOnly` om ClientId en thumbprint voor de tenant uit `graph.appid.json`
  in de root van de repo te lezen (gitignored). De app moet in die tenant toestemming hebben:
  GDAP geeft delegated rechten, geen app-only toegang.
- **Bestaande sessies worden hergebruikt** als ze van de juiste soort zijn, voor de juiste
  tenant, en (delegated) alle gevraagde scopes al hebben. Een delegated herverbinding houdt
  de scopes van de eerdere sessie, zodat een tweede script in hetzelfde venster ze niet afneemt.

**Opmerkingen**

- Vereist PowerShell 7. Elke `Connect-*` stopt met een installatietip als de module ontbreekt.
- Draai vanuit de repo: scripts dot-sourcen dit bestand via een relatief pad, dus een los
  gekopieerd script heeft dit bestand ernaast nodig.

---

## Install-Modules.ps1

Installeert en importeert elke module uit [`RequiredModules.psd1`](#requiredmodulespsd1). Draai het één keer op een nieuwe machine of na een schone PowerShell-installatie. Modules die al geïnstalleerd zijn blijven ongemoeid — bijwerken doet [`Update-Modules.ps1`](#update-modulesps1).

```powershell
.\scripts\Startup\Install-Modules.ps1
```

**Parameters**

| Parameter | Beschrijving |
|-----------|--------------|
| `-Force` | Modules opnieuw installeren, ook als ze er al zijn |
| `-Scope` | `CurrentUser` (standaard) of `AllUsers` (als administrator) |

---

## Update-Modules.ps1

Controleert elke module uit [`RequiredModules.psd1`](#requiredmodulespsd1) en geeft elke module een status:

| Status | Betekenis | Zonder `-CheckOnly` |
|--------|-----------|---------------------|
| `Missing` | Niet geïnstalleerd | Wordt geïnstalleerd |
| `BelowMinimum` | Ouder dan de `MinimumVersion` | Wordt geüpdatet |
| `UpdateAvailable` | Er staat een nieuwere versie op de PowerShell Gallery | Wordt geüpdatet |
| `OK` | Actueel | — |
| `Unknown` | Gallery niet bereikbaar, of de module staat er niet meer op | — |
| `Skipped` | Niet voor dit platform of deze PowerShell-versie | — |

Daarna werkt het, tenzij `-RequiredOnly`, zoals altijd ook elke andere via PowerShellGet geïnstalleerde module bij. `load.ps1` draait het bij het starten als `-RequiredOnly -Auto -MaxAgeHours 24`.

**Parameters**

| Parameter | Beschrijving |
|-----------|--------------|
| `-CheckOnly` | Alleen rapporteren; niets installeren of updaten |
| `-RequiredOnly` | Alleen de modules uit `RequiredModules.psd1`, niet al het andere dat geïnstalleerd is |
| `-MaxAgeHours` | Gallery-versies uit de cache hergebruiken als die jonger is dan dit aantal uur (standaard `0` = altijd de gallery bevragen) |
| `-Scope` | Scope voor nieuw geïnstalleerde modules: `CurrentUser` (standaard) of `AllUsers` |
| `-Quiet` | Geen regel per module, alleen fouten |
| `-PassThru` | Geeft per vereiste module een statusobject terug (`Name`, `Installed`, `Minimum`, `Latest`, `Status`, `Reason`) |
| `-Auto` | Voor bij het starten: installeert wat ontbreekt en updatet wat verouderd is zonder te vragen, en toont alleen wat het doet. Alles in orde geeft één regel, `Modules OK` |
| `-Prompt` | Als `-Auto`, maar toont eerst wat het zou doen en vraagt het dan |

**Voorbeelden**

```powershell
# Wat ontbreekt of is verouderd? Verandert niets
.\scripts\Startup\Update-Modules.ps1 -RequiredOnly -CheckOnly

# Ontbrekende installeren en achterlopende updaten, alleen de modules van de repo
.\scripts\Startup\Update-Modules.ps1 -RequiredOnly

# Hetzelfde, en daarna ook alle andere geïnstalleerde modules bijwerken
.\scripts\Startup\Update-Modules.ps1
```

**In je PowerShell-profiel.** Start je PowerShell met je eigen profiel (`$PROFILE`) in plaats van met `load.ps1`, zet dan de regel erin die `load.ps1` gebruikt, zodat de controle bij elke start van PowerShell draait:

```powershell
& "C:\pad\naar\M365-Scripts\scripts\Startup\Update-Modules.ps1" -RequiredOnly -Auto -MaxAgeHours 24
```

Als alles actueel is geeft dat één regel en kost het ruim minder dan een seconde; de gallery wordt hooguit eens per dag bevraagd.

**Opmerkingen**

- De geïnstalleerde versie wordt gelezen met `Get-Module -ListAvailable`, dus een module die niet via PowerShellGet is geïnstalleerd (handmatig gekopieerd, een MSI) telt ook als geïnstalleerd. Zo'n module krijgt de nieuwe versie ernaast via `Install-Module`, omdat `Update-Module` modules weigert die het niet zelf installeerde.
- Gallery-versies komen van `Find-PSResource` als PSResourceGet aanwezig is (ongeveer 3 s voor de hele lijst), anders van `Find-Module` (ongeveer 9 s). Ze worden gecachet in `%LOCALAPPDATA%\M365-Scripts\module-gallery-cache.json`; `-MaxAgeHours` bepaalt hoe oud die cache mag zijn.
- Offline mislukt het opvragen van de gallery stil: ontbrekende en te oude modules worden nog steeds gemeld, *update beschikbaar* niet.
- PowerShell 7 en Windows PowerShell 5.1 hebben aparte modulemappen, dus op dezelfde machine kunnen ze een ander resultaat geven. Dat klopt, het is geen fout.
- Oude versies worden niet verwijderd. Draai als administrator om modules bij te werken die voor alle gebruikers zijn geïnstalleerd.

---

## Test-PowerShellSyntax.ps1

Controleert `.ps1`-bestanden (en optioneel `.psm1`) op syntaxfouten door ze te parsen zonder ze uit te voeren — gebruikt `[System.Management.Automation.Language.Parser]::ParseFile()`.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-Path` | Nee | Bestand of map om te controleren (standaard: root van de repo) |
| `-Recurse` | Nee | Ook submappen doorlopen als `-Path` een map is |
| `-IncludePsm1` | Nee | Ook `.psm1`-modulebestanden controleren |

**Voorbeelden**

```powershell
# De hele repo controleren
.\Test-PowerShellSyntax.ps1 -Recurse

# Eén bestand controleren
.\Test-PowerShellSyntax.ps1 -Path .\scripts\Entra\New-M365User.ps1
```

Exitcodes: `0` = geen fouten, `1` = syntaxfouten gevonden, `2` = fout in pad/argument.

---

## Update-ScriptIndex.ps1

Bouwt [`scripts/INDEX.md`](../INDEX.md): één pagina met elk script in de repository
van A tot Z, met een link naar het bestand, een link naar de readme van de map en een omschrijving van één regel.

Hij bestaat omdat je een script op GitHub anders alleen vindt door te raden onder welke
workloadmap het staat en readmes te openen tot het opduikt. Eén gegenereerde pagina is
met Ctrl-F te doorzoeken en aanklikbaar, en kan — omdat hij gegenereerd is — niet uit de
pas gaan lopen met de bestanden, zoals een met de hand bijgehouden tabel wel doet.

**Waar de omschrijving vandaan komt**

| Volgorde | Bron |
|-------|--------|
| 1 | Het `.SYNOPSIS`-blok van het script, samengevoegd over de regels waarover het doorloopt |
| 2 | Anders de eerste echte regel van een `#`-commentaarblok bovenaan |
| 3 | Anders niets — en het script komt onder *Scripts without a description*, zodat het gat zichtbaar is in plaats van stilletjes leeg |

Een enkele `#`-commentaarregel die direct boven code staat, wordt bewust **niet** gebruikt: een regel
als `# URL van de theme` boven een toewijzing aan `$ThemeUrl` beschrijft die variabele, niet het
script, en die als omschrijving lezen zet iets erger dan niets in de tabel.
Een headerblok beslaat meerdere regels, of is door een lege regel van de code gescheiden.

**Parameters**

| Parameter | Omschrijving |
|-----------|-------------|
| `-Root` | Root van de repository (standaard: twee niveaus boven dit script) |
| `-Check` | Schrijft niets; exit `1` als de gecommitte pagina niet meer overeenkomt met de scripts op schijf |
| `-WhatIf` | Meldt wat er zou veranderen zonder te schrijven |

**Voorbeelden**

```powershell
# De index opnieuw opbouwen na het toevoegen, hernoemen of verwijderen van een script
pwsh -File scripts/Startup/Update-ScriptIndex.ps1

# Falen als de index verouderd is — voor een hook of een pipeline
pwsh -File scripts/Startup/Update-ScriptIndex.ps1 -Check
```

> Draai hem opnieuw zodra een script wordt toegevoegd, hernoemd, verplaatst of verwijderd — hetzelfde
> moment waarop de [werkregels](../../.claude/CLAUDE.md) je al vragen de readmes en
> `menu.ps1` bij te werken. Hij herschrijft niets als de pagina al actueel is, dus je kunt hem veilig
> bij elke commit draaien.

Exitcodes: `0` = geschreven of al actueel, `1` = `-Check` vond de pagina verouderd.

---

## Convert-MarkdownToHtml.ps1

De servicedeskdocumenten in deze repo zijn markdown, maar IT Glue en de meeste ticketsystemen willen rich text. Met de hand omzetten betekent dat de HTML verouderd is zodra de markdown voor het eerst verandert — en `Update-TeamsClient-ITGlue.md` veranderde zes keer in twee dagen — dus de pagina wordt in plaats daarvan gegenereerd.

Ondersteund, omdat deze documenten het gebruiken: koppen, tabellen met een kopregel, fenced codeblokken (ook de ingesprongen binnen genummerde stappen), blockquotes, genummerde en ongenummerde lijsten, horizontale lijnen, en inline code, vet, cursief en links. Al het andere gaat als tekst door in plaats van dat ernaar wordt gegokt.

De CSS is ingebed, dus de pagina staat op zichzelf — niets te hosten en niets wat breekt als het bestand ergens anders heen wordt gekopieerd. Void-elementen worden zelfsluitend geschreven (`<hr/>`, `<br/>`), zodat de uitvoer zowel als XML als als HTML parseert en structureel kan worden gecontroleerd in plaats van op het oog.

**Parameters**

| Parameter | Omschrijving |
|-----------|-------------|
| `-Path` | Het markdownbestand dat moet worden omgezet (verplicht) |
| `-Destination` | Waar de HTML naartoe wordt geschreven (standaard: dezelfde map en naam, `.html`) |
| `-Title` | Pagina- en browsertitel (standaard: de eerste `#`-kop van het document) |
| `-Check` | Schrijft niets; exit `1` als de HTML op schijf niet meer overeenkomt met de markdown |
| `-WhatIf` | Toont wat er zou worden geschreven en wijzigt niets |

**Voorbeelden**

```powershell
# De IT Glue-pagina naast zijn markdownbron bouwen
pwsh -File scripts/Startup/Convert-MarkdownToHtml.ps1 -Path scripts/Device/Update-TeamsClient-ITGlue.md

# Loopt de gecommitte pagina achter? Exitcode 1 als dat zo is
pwsh -File scripts/Startup/Convert-MarkdownToHtml.ps1 -Path scripts/Device/Update-TeamsClient-ITGlue.md -Check
```

Vanuit het menu: `menu.ps1`, toets **M**. Die biedt het Teams-document voor IT Glue aan als standaardpad en vraagt of je alleen wilt controleren.

**Opmerkingen**

- **In IT Glue krijgen:** open de `.html` in een browser, selecteer alles, kopieer en plak het in de documenteditor van IT Glue. De editor behoudt de koppen, tabellen en codeblokken en laat de CSS vallen — precies wat je daar wilt, want IT Glue past zijn eigen opmaak toe.
- De regel met de generatiedatum wordt uitgesloten van de `-Check`-vergelijking, dus opnieuw draaien op een ongewijzigd document meldt geen verschil.
- Draai het opnieuw na het bewerken van de markdown. `-Check` is wat een pre-commit-hook of een pipeline zou aanroepen.

Exitcodes: `0` = geschreven of al actueel, `1` = `-Check` vond de pagina verouderd, of de pagina bestaat nog niet.

---

## Test-MarkdownLinks.ps1

Loopt elk `.md`-bestand in de repository langs en meldt links die nergens heen gaan. Een dode
link in een readme is onzichtbaar tot iemand erop klikt, en dat is meestal precies het moment
dat die persoon hem nodig had.

Hij controleert twee soorten:

| Soort | Wat er mis kan gaan |
|------|-------------------|
| Een link naar een bestand of map | Het bestand is hernoemd, verplaatst of verwijderd en de readme wijst nog naar het oude pad. Procent-gecodeerde spaties (`Time%20sync/readme.md`) worden gedecodeerd voordat het pad wordt getest, want zo serveert GitHub het |
| Een anchor binnen de pagina (`#set-usermanagerps1`) | De kop waarnaar hij wijst is hernoemd, of de anchor is met de hand getypt en kwam nooit overeen. Deze rotten in stilte weg — niets waarschuwt je |

Anchors worden opgelost zoals GitHub ze bouwt: de kop wordt naar kleine letters omgezet, markdown-
opmaak wordt verwijderd, alles wat geen letter, cijfer, spatie, `_` of `-` is, valt weg,
en spaties worden koppeltekens — dus `### Watch-RDSLive.ps1` is `#watch-rdsliveps1`,
niet `#watch-rdslivesps1`. Herhaalde koppen krijgen GitHubs achtervoegsel `-1`, `-2`. Onzichtbare
tekens (variation selectors, zero-width joiners — de bytes die van een emoji een
emoji maken) worden zowel uit de kop als uit de link verwijderd voordat ze worden vergeleken, zodat een
emojikop in de inhoudsopgave niet als kapot wordt gezien.

Fenced codeblokken en inline code worden overgeslagen, zodat een readme die link-
syntaxis *documenteert* zichzelf niet als kapot meldt — het voorbeeld `([docs](#…))` twee alinea's hoger
is tekst over links, geen link.

Externe links (`http`, `https`, `mailto`) worden geteld maar niet opgehaald: dit is een
structurele controle, en die moet offline werken.

**Parameters**

| Parameter | Omschrijving |
|-----------|-------------|
| `-Root` | Root van de repository (standaard: twee niveaus boven dit script) |
| `-Path` | Controleer één bestand of map in plaats van de hele repository |

**Voorbeelden**

```powershell
# Elke readme in de repository controleren
pwsh -File scripts/Startup/Test-MarkdownLinks.ps1

# Alleen één workloadmap
pwsh -File scripts/Startup/Test-MarkdownLinks.ps1 -Path scripts/Exchange
```

Exitcodes: `0` = elke interne link klopt, `1` = er is iets kapot (elk geval vermeld
met het bestand waarin het staat en waarom het faalde).

---

## Update-ReadmeHeader.ps1

Elke map heeft zijn readme drie keer: `readme.md` (Engels), `readme.nl.md` (Nederlands) en
`readme.fr.md` (Frans). Elk begint met dezelfde twee regels, die alleen in hun paden
verschillen: een taalwissel naar dezelfde pagina in de andere talen, en een kruimelpad terug
omhoog waarin elk niveau naar zijn eigen readme **in de huidige taal** linkt.

Met de hand bijgehouden gaan juist die paden mis — één `../` te weinig nadat een map is
verplaatst, of een Nederlandse pagina die naar de Engelse bovenliggende map linkt. Daarom
worden ze gegenereerd uit de map waarin de readme staat, en uit niets anders. Alles boven de
eerste kop dat een taalwissel- of kruimelpadregel is, wordt vervangen; de rest van het bestand
blijft onaangeroerd.

Een map met een `readme.md` maar zonder Nederlandse of Franse versie wordt gemeld: de
taalwissel zou daar nergens naartoe linken.

**Parameters**

| Parameter | Omschrijving |
|-----------|--------------|
| `-Root` | Root van de repository (standaard: twee niveaus boven dit script) |
| `-Check` | Schrijft niets; stopt met exitcode `1` als een header verouderd is of een taalversie ontbreekt |
| `-WhatIf` | Meldt welke headers herschreven zouden worden, zonder te schrijven |

**Voorbeelden**

```powershell
# Na het toevoegen of verplaatsen van een map-readme (schrijf eerst de .nl.md- en .fr.md-versie)
pwsh -File scripts/Startup/Update-ReadmeHeader.ps1

# Faal als een header verouderd is of een vertaling ontbreekt — voor een hook of een pipeline
pwsh -File scripts/Startup/Update-ReadmeHeader.ps1 -Check
```

> Draai daarna [`Test-MarkdownLinks.ps1`](#test-markdownlinksps1): dit script schrijft de
> links, dat script bewijst dat ze kloppen.
