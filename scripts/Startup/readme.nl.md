[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../readme.nl.md) › [scripts](../readme.nl.md) › **Startup**

# Startup

Startscripts en de centrale M365-functiebibliotheek.

---

## Bestanden

| Bestand | Omschrijving |
|------|-------------|
| [`functies.ps1`](functies.ps1) | M365-functiebibliotheek — bij eerste gebruik door `menu.ps1` gedot-sourced |
| [`Install-Modules.ps1`](Install-Modules.ps1) | Bootstrapscript — installeert en importeert alle benodigde PowerShell-modules |
| [`Update-Modules.ps1`](Update-Modules.ps1) | Werkt elke geïnstalleerde PowerShell-module bij naar de nieuwste versie |
| [`Test-PowerShellSyntax.ps1`](Test-PowerShellSyntax.ps1) | Controleert `.ps1`-bestanden in de repo op syntaxfouten door ze te parsen, zonder ze uit te voeren |
| [`Update-ScriptIndex.ps1`](Update-ScriptIndex.ps1) | Genereert [`scripts/INDEX.md`](../INDEX.md) opnieuw — de doorzoekbare A–Z-lijst van alle scripts |
| [`Test-MarkdownLinks.ps1`](Test-MarkdownLinks.ps1) | Controleert elke link in elke readme — bestanden die moeten bestaan, anchors die met een kop moeten overeenkomen |
| [`Convert-MarkdownToHtml.ps1`](Convert-MarkdownToHtml.ps1) | Bouwt van een markdowndocument een op zichzelf staande, opgemaakte HTML-pagina — om in IT Glue te plakken of af te drukken |
| [`Update-ReadmeHeader.ps1`](Update-ReadmeHeader.ps1) | Zet de taalwissel en het kruimelpad bovenaan elke readme, in het Engels, Nederlands en Frans |

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
# Alle volgende functies richten zich automatisch op de geselecteerde tenant
```

### Functies

**Verbinding**

| Functie | Omschrijving |
|----------|-------------|
| `Connect-Tenant` | Selecteert een CSP-klant op domein, vult `$cid` en `$connectmsoldomain` |
| `Test-GdapConnection` | Valideert het gedelegeerde GDAP/CSP-contract + probeert een gedelegeerde Exchange-verbinding |
| `Test-ExoConnection` | Controleert / herstelt de Exchange Online-verbinding |

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

## Install-Modules.ps1

Installeert en importeert alle PowerShell-modules die deze repository nodig heeft. Draai het één keer op een nieuwe machine of na een schone PowerShell-installatie.

```powershell
.\scripts\Startup\Install-Modules.ps1
```

De geïnstalleerde kernmodules zijn onder meer `ExchangeOnlineManagement` en de benodigde Microsoft Graph-submodules (`Microsoft.Graph.Authentication`, `Microsoft.Graph.Sites`, `Microsoft.Graph.Identity.DirectoryManagement`, `Microsoft.Graph.Identity.SignIns`, `Microsoft.Graph.Identity.Governance`, `Microsoft.Graph.Applications`, `Microsoft.Graph.Groups`, `Microsoft.Graph.Users`, `Microsoft.Graph.Calendar`).
Waar van toepassing worden ook compatibiliteitsmodules die alleen op Windows werken meegenomen (`WindowsAutopilotIntune`, `AzureAD`).

---

## Update-Modules.ps1

Werkt elke geïnstalleerde PowerShell-module bij naar de nieuwste versie. Draai het als administrator voor modules die systeembreed zijn geïnstalleerd.

Zorgt daarnaast voor een minimumversie van de specifieke Graph-submodules waar deze repo van afhangt (`Microsoft.Graph.Authentication`, `Microsoft.Graph.Sites`, `Identity.SignIns`, `Identity.Governance`, `Applications`, `Groups`), voordat al het andere wat op de machine is geïnstalleerd wordt bijgewerkt.

```powershell
.\scripts\Startup\Update-Modules.ps1
```

> Geen parameters. Loopt elke module langs die `Get-InstalledModule` teruggeeft, dus op een machine met veel geïnstalleerde modules kan het even duren.

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
