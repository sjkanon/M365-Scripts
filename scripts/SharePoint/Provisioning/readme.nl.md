[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [SharePoint](../readme.nl.md) › **Provisioning**

# SharePoint-structuur inrichten

Richt een SharePoint-structuur in en houd die bij — metadatamodel, bibliotheken,
inhoudstypen en groepsmachtigingen — vanuit één configuratiebestand, met PnP PowerShell
en Microsoft Graph.

Niets in de scripts is klantspecifiek: het model staat in de JSON, dus een tweede klant
is een tweede configuratiebestand, geen tweede fork. De voorbeelden hieronder gebruiken
een fictief Contoso NV met de merken Northwind en Fabrikam.

---

## Inhoud

- [Het idee](#het-idee)
- [Scripts](#scripts)
- [Voordat je begint](#voordat-je-begint)
- [Volgorde van werken](#volgorde-van-werken)
- [Het configuratiebestand](#het-configuratiebestand)
- [Standaardkanalen en machtigingen — lees dit](#standaardkanalen-en-machtigingen--lees-dit)
- [De audit inplannen](#de-audit-inplannen)
- [Wat de scripts nooit doen](#wat-de-scripts-nooit-doen)

---

## Het idee

Eén Microsoft 365-team, één kanaal per pijler, en het verschil tussen de pijlers
gedragen door **metadata en groepsmachtigingen** in plaats van door een wildgroei aan sites.

| | |
|---|---|
| **Pijlers** | MGMT (privékanaal), Leveranciers, Verkopers, Klanten, Marketing, TD |
| **Extra bibliotheek** | Beeldmateriaal voor klanten — alleen-lezen voor externe klanten |
| **Merk** | Northwind / Fabrikam / Beide — een label op elk bestand, nooit een aparte site of groep |
| **Groepen** | één Entra ID-beveiligingsgroep per pijler per toegangsniveau (`SG-CONTOSO-<Pijler>-RW` / `-RO`), plus `SG-CONTOSO-Klanten-Extern` |

Het metadatamodel, herbruikbaar in elke bibliotheek:

| Kolom | Interne naam | Type | Waarden |
|---|---|---|---|
| Merk | `PsMerk` | Keuze | Northwind / Fabrikam / Beide |
| Pijler | `PsPijler` | Keuze | MGMT / Leveranciers / Verkopers / Klanten / Marketing / TD |
| Regio | `PsRegio` | Keuze | Benelux / Duitsland / Frankrijk / Export — alleen verplicht op Verkoopdocument |
| Leverancier | `PsLeverancier` | Beheerde metadata | termenset, uit te breiden vanuit het termenarchief |
| Taal | `PsTaal` | Meerkeuze | NL / FR / DE / EN / Geen taal |
| Contenttype | `PsContenttype` | Keuze | Catalogus / Prijslijst / Schrijfrichtlijn / Afbeelding+certificaat / Marketingslag |
| Vertrouwelijkheid | `PsVertrouwelijkheid` | Keuze | Intern / Deelbaar met klant / Vertrouwelijk |
| Deelstatus | `PsDeelstatus` | Keuze | bijgehouden door het auditscript — nooit met de hand ingevuld |
| Status | `PsStatus` | Keuze | Actief / Te archiveren / Verouderd |

> **Waarom het voorvoegsel `Ps`.** "Contenttype" en "Status" zijn weergavenamen die
> SharePoint al voor iets anders gebruikt. Door de *interne* namen een voorvoegsel te geven
> blijven de kolommen eenduidig in CAML, in weergaven en in de driftcontrole, terwijl
> gebruikers gewoon Nederlandse labels zien.

Een inhoudstype per pijler bepaalt welke daarvan verplicht zijn — een Marketingdocument
kan niet worden opgeslagen zonder Taal en Contenttype, een Verkoopdocument niet zonder
Regio.

---

## Scripts

| Script | Wat het doet | Schrijft? |
|---|---|---|
| [`Install-SharePointStructure.ps1`](Install-SharePointStructure.ps1) ([docs](#install-sharepointstructureps1)) | **Begin hier.** Vraagt hoe alles moet heten en bouwt dan het geheel: app-registratie, team, kanalen, metadata, bibliotheken, machtigingen, verificatie | ja |
| [`New-StructureConfig.ps1`](New-StructureConfig.ps1) ([docs](#je-krijgt-vragen-geen-json-bestand)) | De vragen. Draait vanzelf vanuit het installatiescript; draai het los om vooraf een configuratie voor te bereiden | schrijft de configuratie |
| [`New-SharePointTeam.ps1`](New-SharePointTeam.ps1) ([docs](#install-sharepointstructureps1)) | Het Microsoft 365-team en zijn kanalen, privékanalen inbegrepen, en de site-URL's teruggeschreven | ja |
| [`New-SharePointMetadata.ps1`](New-SharePointMetadata.ps1) ([docs](#new-sharepointmetadataps1)) | Termenset, sitekolommen, inhoudstypen — op elke site in de configuratie | ja |
| [`Set-SharePointLibraries.ps1`](Set-SharePointLibraries.ps1) ([docs](#set-sharepointlibrariesps1)) | Bibliotheken, kanaalmappen, koppeling van inhoudstypen, standaardmetadata, weergaven, groepsmachtigingen | ja |
| [`Update-SharePointShareStatus.ps1`](Update-SharePointShareStatus.ps1) ([docs](#update-sharepointsharestatusps1)) | Leidt Deelstatus af uit de werkelijke machtigingen en markeert bestanden die breder gedeeld zijn dan hun label toestaat | één kolom |
| [`Test-SharePointStructure.ps1`](Test-SharePointStructure.ps1) ([docs](#test-sharepointstructureps1)) | Vergelijkt de tenant met de configuratie en rapporteert elk verschil | nooit |
| [`Sync-SharePointChannelMember.ps1`](Sync-SharePointChannelMember.ps1) ([docs](#privékanalen-en-groepen)) | Maakt een beveiligingsgroep de bron van waarheid voor wie er in een privékanaal zit | kanaalledenlijst |
| [`Add-SharePointHelpPage.ps1`](Add-SharePointHelpPage.ps1) ([docs](#overdracht-aan-de-klant)) | Zet de uitleg voor eindgebruikers op de teamsite, gegenereerd uit de configuratie | ja |
| [`Remove-SharePointStructure.ps1`](Remove-SharePointStructure.ps1) ([docs](#terugdraaien)) | Verwijdert wat gebouwd is — rapporteert alleen, tenzij je `-Apply` meegeeft | ja, met opzet |
| [`SharePointStructure.Common.ps1`](SharePointStructure.Common.ps1) | Gedeelde hulpfuncties — wordt gedot-sourcet, niet los gedraaid | — |
| [`SharePoint-Handleiding.md`](SharePoint-Handleiding.md) | **Handleiding voor eindgebruikers, in het Nederlands** — geef deze aan de klant: uploaden, labelen, dingen terugvinden | — |
| [`example.config.json`](example.config.json) | Het model, als voorbeeld om te kopiëren — nog op `CHANGEME`. Klantconfigs (`<klant>.config.json`) staan ernaast en worden door git genegeerd | — |

Elk schrijvend script ondersteunt `-WhatIf` en is idempotent: een tweede run meldt overal
`[ OK ]` en verandert niets.

---

## Voordat je begint

### 1. Modules

```powershell
Install-Module PnP.PowerShell         -Scope CurrentUser
Install-Module Microsoft.Graph.Groups -Scope CurrentUser   # alleen voor -EnsureGroups / -IncludeGroups
```

### 2. Een app-registratie

PnP.PowerShell levert geen gedeelde multitenant-app meer mee, dus je hebt een eigen app
nodig. Hergebruik de app die [`Find-SiteContent.ps1`](../Find-SiteContent.ps1) aanmaakt en
in `pnp.appid.json` bewaart, of registreer er een:

| Aanmelding | Vereist | Gebruik je voor |
|---|---|---|
| **Interactief** (`-Interactive -ClientId <app-id>`) | gedelegeerde `AllSites.FullControl`, aangemeld als beheerder die de site mag bewerken | inrichten, eenmalige runs |
| **App-only** (`-ClientId <app-id> -Thumbprint <thumb>`) | applicatiemachtiging `Sites.FullControl.All`, certificaat geüpload naar de app | de ingeplande deelstatus-audit |

Twee extra vereisten die je makkelijk over het hoofd ziet:

- **Termenarchief.** Voor het aanmaken van de termengroep en termenset is een
  termenarchiefbeheerder nodig. App-only kan dat alleen als de service principal van de
  app als zodanig is toegevoegd in het SharePoint-beheercentrum. Is dat een drempel, draai
  dan `New-SharePointMetadata.ps1 -Only TermSet` één keer interactief en laat de rest aan
  app-only over.
- **Groepen.** `-EnsureGroups` heeft `Group.ReadWrite.All` nodig; voor `-IncludeGroups` van
  de driftcontrole volstaat `Group.Read.All`.

### 3. Vul de configuratie in

`example.config.json` wordt geleverd met `CHANGEME` in de tenant- en site-URL's. Kopieer
het naar `<klant>.config.json` en vul het in, of laat `New-StructureConfig.ps1` er een
schrijven. Zonder `-ConfigPath` neemt elk script de ene `*.config.json` hier waarin geen
`CHANGEME` meer staat, en weigert het op het voorbeeld zelf — liever een duidelijke fout dan een
aanmelding die na vijf minuten in een run mislukt.

```jsonc
"tenant": "contoso.onmicrosoft.com",
"sites": {
  "team": "https://contoso.sharepoint.com/sites/Contoso",
  "mgmt": "https://contoso.sharepoint.com/sites/Contoso-MGMT"
}
```

> De **mgmt**-URL is de *eigen* siteverzameling van het privékanaal, geen map in de
> teamsite. Je vindt hem in het SharePoint-beheercentrum, of open het tabblad Bestanden
> van het MGMT-kanaal en klik op *Openen in SharePoint*. Een privékanaal heeft een eigen
> site, en daarom moeten de kolommen en inhoudstypen daar apart worden ingericht — een
> sitekolom reikt niet over een siteverzameling heen.

---

## Volgorde van werken

### De korte versie

```powershell
.\Install-SharePointStructure.ps1
```

Dat is alles. Zonder configuratie voor deze tenant vraagt het hoe alles moet heten,
schrijft het zelf de configuratie en maakt het daarna het team, de kanalen (het
privékanaal inbegrepen), het metadatamodel, de bibliotheken, de groepen, de weergaven en
de machtigingen aan — en controleert het resultaat.

Daarna: zet mensen in de beveiligingsgroepen. De uitleg voor hen staat al op de site, in
de navigatie links — de build heeft hem daar neergezet.

### Je krijgt vragen, geen JSON-bestand

`New-StructureConfig.ps1` draait de eerste keer vanzelf. Enter neemt de suggestie tussen
haakjes over, dus een standaardbuild is vooral Enter drukken plus twee echte antwoorden:

| Gevraagd | Suggestie |
|---|---|
| Klant, tenant | — |
| Teamnaam, alias (bepaalt de site-URL), eigenaar | afgeleid van de klantnaam |
| Merken, en hoe "hoort bij allemaal" heet | Northwind, Fabrikam, Beide |
| Pijlers, en welke een privékanaal zijn | MGMT, Leveranciers, Verkopers, Klanten, Marketing, TD — MGMT privé |
| Welke pijler leveranciers / verkoop behandelt | bepaalt waar Leverancier en Regio verplicht worden |
| Klantbibliotheek | Beeldmateriaal voor klanten |
| Groepsvoorvoegsel en de achtervoegsels voor bewerken/lezen | `SG-<CLIENT>` · RW · RO |
| Talen, regio's, documentsoorten, vertrouwelijkheidsniveaus, statussen, startleveranciers | de Nederlandse standaardwaarden |
| **Deelstatus-kolom bijhouden?** | nee — dit is het enige antwoord dat je een nachtelijk script kost |
| **Rechten per pijler afdwingen op kanaalmappen?** | nee — zie de opmerking over standaardkanalen hieronder |

Al het andere wordt afgeleid: per pijler een kanaal, een inhoudstype, twee
beveiligingsgroepen en een gegroepeerde weergave; per merk een weergave over alle pijlers.

Optionele vragen tonen `(of "geen")` in de hint — typ dat om de suggestie af te wijzen,
want Enter betekent "neem hem over".

#### `-All` — bepaal zelf elke naam

`New-StructureConfig.ps1 -All` vraagt ook naar de namen die anders worden afgeleid, elk
nog steeds met de afleiding als suggestie:

| Gevraagd met `-All` | Suggestie |
|---|---|
| URL van de teamsite | `https://<tenant>.sharepoint.com/sites/<alias>` |
| De bibliotheek achter de kanalen | `Documents` — vraag hiernaar op een niet-Engelstalige tenant |
| Kolomgroep en inhoudstypegroep | de teamnaam |
| Naam van de termenset | `Leveranciers` |
| Het label van elke kolom, zoals gebruikers het zien | Merk, Pijler, Regio, Leverancier, Taal, Contenttype, … |
| Per pijler: kanaalnaam, map, inhoudstype, beide groepsnamen, weergavetitel | afgeleid van de pijlernaam |
| Klantbibliotheek: naam van groep en inhoudstype | `<prefix>-Klanten-Extern`, `Klantmedia` |

De *interne* kolomnamen (`PsMerk`, `PsTaal`, …) liggen in beide gevallen vast. Ze worden
aan niemand getoond, en als je er een wijzigt nadat documenten hem dragen, verlies je de
metadata op die documenten.

**Twee dingen worden eenmalig gegenereerd en liggen daarna vast**, omdat SharePoint er
gegevens aan koppelt: de interne kolomnamen (`PsMerk`, `PsTaal`, …) en de
inhoudstype-ID's. Weergavenamen, kanaalnamen en groepsnamen kun je achteraf allemaal
wijzigen; die twee niet zonder de metadata te verliezen op documenten die ze al dragen.
Daarom weigert de wizard een bestaande configuratie te overschrijven zonder `-Force`.

### Install-SharePointStructure.ps1

Eén run, vijf stappen, en stoppen bij de eerste fout in plaats van verder te bouwen op
iets dat stuk is:

| Stap | Wat |
|---|---|
| 0 | App-registratie — aangemaakt en met beheerderstoestemming, of hergebruikt uit `pnp.appid.json` |
| 1 | `New-SharePointTeam.ps1` — het Microsoft 365-team, de kanalen inclusief het privékanaal, en de site-URL's teruggeschreven in de configuratie |
| 2 | `New-SharePointMetadata.ps1` — termenset, kolommen, inhoudstypen, op elke site |
| 3 | `Set-SharePointLibraries.ps1 -EnsureGroups` — groepen, bibliotheken, mappen, inhoudstypen, standaardwaarden, weergaven, machtigingen |
| 4 | `Add-SharePointHelpPage.ps1` — de uitleg, op de site, voor de mensen die ermee gaan werken |
| 5 | `Test-SharePointStructure.ps1` — alleen-lezen verificatie van wat er net is neergezet |
| 6 | `Update-SharePointShareStatus.ps1` met `-RunAudit` — de eerste deelstatus-ronde |

**Stap 4 hoort bij het bouwen, het is geen klusje voor later.** Een structuur waar niemand
over is verteld, is een structuur die niemand gebruikt, en de pagina wordt gegenereerd uit
dezelfde configuratie, dus hij beschrijft wat de run zojuist heeft gemaakt. `-SkipHelpPage`
laat hem weg; `-HelpContact` geeft aan bij wie mensen terechtkunnen.

**Stap 1 is de reden dat dit vanaf een lege tenant werkt.** De siteverzameling van een
privékanaal wordt asynchroon ingericht en de URL is niet vooraf te kennen — SharePoint
verzint hem op basis van de team- en kanaalnaam. Het script wacht erop (een paar minuten
is normaal) en schrijft hem in de configuratie, zodat de volgende stappen ergens mee
kunnen verbinden. Geef `-SkipTeam` mee als het team al bestaat.

```powershell
# Eenmalige build op een tenant die je niet dagelijks beheert: laat niets achter
.\Install-SharePointStructure.ps1 -TemporaryApp -RunAudit

# Voorzichtig: alles behalve de niet-ondersteunde machtigingen op kanaalmappen
.\Install-SharePointStructure.ps1 -SkipChannelFolderPermissions

# Gebruik een app-registratie die je al hebt
.\Install-SharePointStructure.ps1 -ClientId <app-id>
```

| Exitcode | Betekenis |
|---|---|
| 0 | gebouwd en gecontroleerd |
| 1 | een stap is mislukt |
| 2 | gebouwd, maar de verificatie vond verschillen |

**Over `-TemporaryApp`.** Het verwijdert de app-registratie aan het eind — maar alleen een
app die *deze run heeft aangemaakt*. Een app die al in de cache stond, bestond al vóór de
run en is aan iemand anders om te verwijderen, dus het script meldt dat in plaats van hem
stilletjes te verwijderen. Zonder de switch blijft de app staan en wordt de client-ID
bewaard, en dat is wat de ingeplande audit en latere driftcontroles nodig hebben.

**Een `-WhatIf`-run heeft een app nodig om mee aan te melden.** Zonder app in de cache voor
de tenant is er niets om als te verbinden, dus de proefdraai valideert de configuratie en
stopt daar. Draai hem één keer echt, of geef `-ClientId` van een bestaande app mee, om stap
voor stap een proefdraai te doen.

### Of stap voor stap

```powershell
# 1. Eerst kijken, dan springen - geen van beide verandert iets
.\Test-SharePointStructure.ps1 -Interactive -ClientId <app-id>
.\New-SharePointMetadata.ps1   -Interactive -ClientId <app-id> -WhatIf

# 2. Eerst het metadatamodel: de bibliotheken kunnen niet koppelen wat niet bestaat
.\New-SharePointMetadata.ps1 -Interactive -ClientId <app-id>

# 3. Bibliotheken, koppeling van inhoudstypen en machtigingen (maakt onderweg de groepen aan)
.\Set-SharePointLibraries.ps1 -Interactive -ClientId <app-id> -EnsureGroups -WhatIf
.\Set-SharePointLibraries.ps1 -Interactive -ClientId <app-id> -EnsureGroups

# 4. Vul de groepen met mensen (Entra ID-portal, of je eigen onboardingscript)

# 5. Eerste audit, en vanaf dan elke nacht
.\Update-SharePointShareStatus.ps1 -Interactive -ClientId <app-id> -ReportOnly

# 6. Bevestig het resultaat
.\Test-SharePointStructure.ps1 -Interactive -ClientId <app-id> -IncludeGroups
```

### New-SharePointMetadata.ps1

Termengroep, termenset en termen; de negen sitekolommen; de zeven inhoudstypen met de
juiste kolommen als verplicht gemarkeerd. Draait tegen **elke** site in de configuratie,
dus de site van het privékanaal krijgt een eigen kopie.

```powershell
# Alleen de site van het privékanaal, nadat MGMT naar een eigen kanaal is verhuisd
.\New-SharePointMetadata.ps1 -Interactive -ClientId <app-id> -Site mgmt

# Alleen de termenset, door een termenarchiefbeheerder
.\New-SharePointMetadata.ps1 -Interactive -ClientId <app-id> -Only TermSet
```

Een kolom achteraf verplicht maken werkt: de `Required`-vlag op een bestaande
veldkoppeling wordt ter plekke bijgewerkt en doorgezet naar de lijsten die het
inhoudstype al gebruiken.

### Set-SharePointLibraries.ps1

Per container: de bibliotheek of kanaalmap, de inhoudstypen gekoppeld aan de bibliotheek,
de eigen volgorde van inhoudstypen van de map (zodat het menu *Nieuw* in het kanaal
Leveranciers Leveranciersdocument aanbiedt en niet de vijf typen van andere pijlers), de
standaardkolomwaarden, een gegroepeerde weergave en de roltoewijzingen.

```powershell
# De externe bibliotheek, en verwijder alles wat de configuratie niet noemt
.\Set-SharePointLibraries.ps1 -Interactive -ClientId <app-id> `
    -Container KlantBibliotheek -RemoveOtherPermissions

# Alles behalve de niet-ondersteunde machtigingen op mappen van standaardkanalen
.\Set-SharePointLibraries.ps1 -Interactive -ClientId <app-id> `
    -SkipChannelFolderPermissions
```

| Switch | Effect |
|---|---|
| `-EnsureGroups` | maak de Entra ID-beveiligingsgroepen aan die nog niet bestaan |
| `-RemoveOtherPermissions` | verwijder roltoewijzingen die de configuratie niet noemt (eigenaren, deellinks en de everyone-claim worden nooit verwijderd) |
| `-RemoveStockContentType` | verwijder het ingebouwde type *Document* zodat niemand iets zonder metadata kan opslaan |
| `-SkipChannelFolderPermissions` | laat kanaalmappen overerven — zie hieronder |

Er wordt nooit een weergave de standaard gemaakt. De standaardweergave van een
Teams-bibliotheek is wat elk lid van het kanaal ziet op het moment dat hij Bestanden opent.

#### Doorsnijdende weergaven — wat "merk als label" echt maakt

Weergaven per pijler tonen altijd maar één kanaalmap. De sectie `libraryViews` voegt
weergaven toe op de gedeelde bibliotheek zelf met `Scope = RecursiveAll`, zodat ze **alle**
pijlermappen in één platte lijst omvatten:

| Weergave | Toont |
|---|---|
| `Alles - Northwind` | elk bestand met label Northwind **of Beide**, over alle pijlers, gegroepeerd per pijler |
| `Alles - Fabrikam` | hetzelfde voor Fabrikam |
| `Nog te taggen` | bestanden zonder Merk — wat slepen-en-neerzetten en OneDrive-synchronisatie achterlaten |
| `Extern gedeeld` | alles wat de audit buiten de organisatie aantrof |
| `Te archiveren` | Status is Te archiveren of Verouderd |

Dit is het antwoord op "één bestand, twee merken": een bestand met label `Beide` wordt één
keer opgeslagen en verschijnt in beide merkweergaven. Geen kopieën die uit elkaar lopen.

Het filter is kale CAML in de configuratie in plaats van een eigen minizoektaal die dit
script zelf heeft verzonnen:

```jsonc
{
  "title": "Alles - Northwind",
  "recursive": true,
  "groupBy": "PsPijler",
  "where": "<Or><Eq><FieldRef Name='PsMerk' /><Value Type='Text'>Northwind</Value></Eq><Eq><FieldRef Name='PsMerk' /><Value Type='Text'>Beide</Value></Eq></Or>",
  "fields": [ "DocIcon", "LinkFilename", "PsPijler", "PsContenttype", "PsTaal", "..." ]
}
```

Groepeer een weergave nooit op `PsTaal` — SharePoint weigert te groeperen op een kolom met
meerdere waarden. Filteren erop werkt prima.

### Update-SharePointShareStatus.ps1

Zoekt uit hoe elk document werkelijk gedeeld is en schrijft dat naar Deelstatus:

| Oordeel | Wanneer |
|---|---|
| `Extern - bewerken` | een Iedereen-link met bewerkrecht, een gast met Bijdragen of meer, of een bewerklink voor specifieke personen waar een gast in zit |
| `Extern - alleen bekijken` | hetzelfde, alleen bekijken |
| `Intern gedeeld` | een link of rechtstreekse toekenning die binnen de tenant blijft |
| `Niet gedeeld` | het bestand erft gewoon over van zijn map of bibliotheek |

Het beantwoordt ook de vraag waar het model eigenlijk voor is: **staat er iets met label
Intern of Vertrouwelijk achter een externe link?** Die komen in het rapport en de run
eindigt met exitcode 2.

De kolom wordt geschreven met `SystemUpdate`, dus Gewijzigd en Gewijzigd door blijven
staan en er wordt geen nieuwe versie aangemaakt — een nachtelijke run duwt niet de hele
bibliotheek bovenaan *recent gewijzigd*.

| Exitcode | Betekenis |
|---|---|
| 0 | klaar, niets breder gedeeld dan het label toestaat |
| 1 | fout |
| 2 | minstens één schending van Vertrouwelijkheid |

### Test-SharePointStructure.ps1

Alleen-lezen. Eén rij per verschil, in drie smaken:

| Soort | Betekenis | Opgelost door |
|---|---|---|
| `Missing` | in de configuratie, niet op de tenant | de inrichtingsscripts |
| `Different` | aanwezig, maar niet zoals geconfigureerd | de inrichtingsscripts |
| `Extra` | op de tenant, niet in de configuratie | niemand — dat is jouw beslissing |

Exitcode 2 betekent drift, dus het past direct in een monitor. De controle die zich het
vaakst terugbetaalt is de **verplicht-vlag op een veld van een inhoudstype** — iemand
vinkt hem uit in de browser en niets lijkt mis tot de halve bibliotheek geen Taal heeft.

---

## Het configuratiebestand

| Sectie | Bevat |
|---|---|
| `tenant`, `sites` | waar alles staat; naar sleutels in `sites` wordt door elke container verwezen |
| `termStore` | termengroep, termenset en de starttermen voor Leverancier |
| `columns` | de sitekolommen: interne naam, weergavenaam, type, keuzes, standaardwaarde |
| `contentTypes` | één per pijler, waarbij `fields[].required` bepaalt wat verplicht is |
| `groups` | de Entra ID-beveiligingsgroepen, op weergavenaam |
| `containers` | de bibliotheken en kanaalmappen, en wie welke rol erop krijgt |

Een container:

```jsonc
{
  "key": "Leveranciers",
  "kind": "ChannelFolder",          // ChannelFolder = een Teams-kanaal; Library = een eigen bibliotheek
  "site": "team",                   // sleutel uit de sectie sites
  "list": "Documents",              // de bibliotheek van het kanaal
  "folder": "Leveranciers",
  "contentTypes": [ "Leveranciersdocument" ],
  "defaultContentType": "Leveranciersdocument",
  "defaultColumnValues": { "PsPijler": "Leveranciers", "PsStatus": "Actief" },
  "uniquePermissions": true,
  "keepExistingPermissions": true,  // kopieer de overgeërfde rechten bij het verbreken van de overname
  "view": { "title": "Op leverancier", "fields": [ ... ], "groupBy": "PsLeverancier" },
  "permissions": [
    { "group": "SG-CONTOSO-Leveranciers-RW", "role": "Contribute" },
    { "group": "SG-CONTOSO-Leveranciers-RO", "role": "Read" }
  ]
}
```

De configuratie wordt gecontroleerd voordat er iets verbinding maakt: een inhoudstype dat
verwijst naar een kolom die niet gedefinieerd is, of een container die rechten geeft aan
een groep die niet in het model staat, faalt bij het laden in plaats van halverwege het
inrichten.

**Een leverancier toevoegen** vraagt geen wijziging in de configuratie — voeg de term toe
in het termenarchief en hij verschijnt in de kolom. De lijst `terms` is alleen de
startset; termen die buiten de configuratie zijn toegevoegd, worden door de driftcontrole
gemeld maar nooit verwijderd.

---

## Standaardkanalen en machtigingen — lees dit

De meegeleverde configuratie geeft elke map van een standaardkanaal unieke machtigingen.
**Microsoft ondersteunt die combinatie niet.**

Een standaardkanaal is per ontwerp zichtbaar voor elk lid van het team. De
SharePoint-machtigingen op de map van het kanaal aanscherpen verbergt de bestanden wel,
maar Teams blijft het kanaal tonen: een lid dat geen toegang meer heeft, krijgt een fout
op het tabblad Bestanden in plaats van een dichte deur. Het werkt, het is niet mooi, en
het is niet gezegend.

De ondersteunde manieren om een pijler af te schermen:

| Optie | Afweging |
|---|---|
| **Privékanaal** | eigen siteverzameling, eigen leden — wat MGMT al gebruikt. Het netst, maar het kanaal verschijnt helemaal niet voor niet-leden |
| **Gedeeld kanaal** | eigen siteverzameling, eigen leden, kan mensen buiten het team bevatten |
| **Eigen bibliotheek** (`kind: Library`) | buiten de kanaalstructuur, unieke machtigingen worden volledig ondersteund — wat de klantenbibliotheek gebruikt |

Verhuis je een pijler naar een privé- of gedeeld kanaal, zet dan `uniquePermissions` op
`false` voor zijn container en voeg zijn nieuwe site toe aan de sectie `sites`.

`Set-SharePointLibraries.ps1` waarschuwt bij elke map van een standaardkanaal waarop het
de overname verbreekt, en `-SkipChannelFolderPermissions` laat ze overerven terwijl de
inhoudstypen, standaardwaarden en weergaven wel worden ingericht.

---

## De audit inplannen

App-only, op een server of een RMM, elke nacht:

```powershell
pwsh -NoProfile -File .\Update-SharePointShareStatus.ps1 `
    -ClientId <app-id> -Thumbprint <thumbprint> -Quiet
```

Met `-Quiet` toont de run alleen de samenvatting, zodat hij in de activiteitenfeed
verschijnt met iets zinnigs te melden. Exitcode 2 betekent dat een bestand met label
Intern of Vertrouwelijk achter een externe link staat — een waarschuwing waard.

Een wekelijkse driftcontrole ernaast:

```powershell
pwsh -NoProfile -File .\Test-SharePointStructure.ps1 `
    -ClientId <app-id> -Thumbprint <thumbprint> -Quiet -IncludeGroups
```

---

## Een kanaal waar maar één groep in komt

Dit is de vorm om naar te grijpen als een pijler afgeschermd moet worden **en** een echte
alleen-lezen-rol moet houden. De wizard vraagt ernaar als `bibliotheek`:

```
Welke pijlers moeten afgeschermd worden [MGMT]:
In welke vorm [bibliotheek]:
```

Wat het bouwt:

| Onderdeel | Waar |
|---|---|
| Een normaal kanaal in het team | iedereen ziet het in Teams, zoals een kanaal hoort |
| Een eigen documentbibliotheek | op de teamsite, geen map in de gedeelde bibliotheek |
| **Unieke machtigingen, overname verbroken zonder te kopiëren** | zodat teamleden er **niet** als bewerker bij komen |
| `-RW` → Contribute, `-RO` → **Read** | een echte alleen-lezen-rol, wat een privékanaal niet kan bieden |
| Een tabblad in het kanaal dat naar de bibliotheek wijst | het eigen tabblad Bestanden van het kanaal kan niet worden omgeleid, dus de bibliotheek staat ernaast |

De machtigingen die overblijven zijn de eigen eigenaren van de site plus die twee groepen.
Dat is wat `keepExistingPermissions: false` op de container betekent, en het is het
verschil tussen "afgeschermd" en "in theorie afgeschermd".

> **Het ingebouwde tabblad Bestanden van het kanaal wijst nog steeds naar de
> teambibliotheek.** Daar komt een map in te staan die niemand gebruikt. Zeg mensen dat ze
> het benoemde tabblad moeten gebruiken, of verwijder het tabblad Bestanden één keer met
> de hand uit het kanaal.

Vergeleken met de alternatieven:

| Vorm | Alleen-lezen-rol | Kanaal zichtbaar voor niet-leden | Ondersteund |
|---|---|---|---|
| **Eigen bibliotheek + kanaal** | ja | ja (de bestanden niet) | ja |
| Privékanaal | nee — leden mogen bewerken, punt | nee | ja |
| Map in standaardkanaal met unieke rechten | ja | ja, maar het tabblad Bestanden geeft fouten | nee |

---

## Privékanalen en groepen

**Een privékanaal kan geen rechten krijgen via een groep.** Teams houdt het lidmaatschap
per persoon bij, en Graph accepteert daar alleen individuele gebruikers. Daar is geen weg
omheen, en de voor de hand liggende omweg is een valkuil:

| Aanpak | Oordeel |
|---|---|
| De groep als kanaallid toevoegen | Niet mogelijk — Graph neemt alleen gebruikers aan |
| De groep toevoegen aan de SharePoint-machtigingen van de kanaalsite | Werkt ongeveer een dag. Teams synchroniseert de kanaalledenlijst eroverheen terug, en in de tussentijd komen die mensen bij de bestanden terwijl het kanaal voor hen onzichtbaar blijft in Teams. Niet ondersteund |
| **Laat de groep de ledenlijst voeden** | Wat `Sync-SharePointChannelMember.ps1` doet |

```powershell
.\Sync-SharePointChannelMember.ps1 -WhatIf     # wie zou worden toegevoegd
.\Sync-SharePointChannelMember.ps1             # voeg ze toe
.\Sync-SharePointChannelMember.ps1 -Prune      # en verwijder wie niet meer in de groepen staat
```

Jij beheert de groep; het script zet de mensen erin in het kanaal. Geneste groepen worden
gevolgd, niet-gebruikers vallen af, en iedereen wordt eerst lid gemaakt van het
bovenliggende team — Teams weigert een lid van een privékanaal dat niet in het team zit,
en de foutmelding die het geeft, zegt dat niet.

Welke groepen welk kanaal voeden, komt uit `channelMembers` van de container. Een
configuratie die geschreven is voordat die sleutel bestond, valt terug op de
geconfigureerde groepen die naar de container zijn vernoemd, en meldt dat het dat deed.

> ### Er is geen alleen-lezen-rol in een privékanaal
>
> Een privékanaal heeft eigenaren en leden, en leden mogen berichten plaatsen en bestanden
> bewerken en verwijderen. Een groep met de naam `-RO` kan daar dus niet "mag kijken"
> betekenen — iedereen die dit script toevoegt, kan schrijven. De run meldt per groep
> hoeveel mensen hij heeft binnengehaald, zodat dat zichtbaar is in plaats van aangenomen.
>
> Als alleen-lezen voor een pijler echt belangrijk is, is een privékanaal daar de
> verkeerde vorm voor. Gebruik een documentbibliotheek met eigen machtigingen, waar Read
> een echte rol is — de klantbibliotheek werkt al zo.

**Het alternatief dat je moet kennen:** een *gedeeld* kanaal ondersteunt wel lidmaatschap
op basis van groepen. Een pijler daarheen verhuizen is de ondersteunde manier om groepen
de toegang tot een kanaal te laten bepalen. Het gedrag hangt af van de instellingen voor
extern delen en B2B direct connect van de tenant, dus probeer één kanaal uit voordat je
iets belangrijks verhuist.

---

## Overdracht aan de klant

De uitleg hoort op de site, niet in deze repo. `Add-SharePointHelpPage.ps1` zet hem daar
als SharePoint-pagina, gegenereerd uit dezelfde configuratie waaruit de structuur is
gebouwd:

```powershell
.\Add-SharePointHelpPage.ps1 -WhatIf                      # wat er zou staan
.\Add-SharePointHelpPage.ps1 -Interactive -ClientId <app-id>
.\Add-SharePointHelpPage.ps1 -Interactive -ClientId <app-id> -Force   # na een wijziging
```

Omdat hij gegenereerd wordt, kan hij niet uit de pas lopen: de kanalen die hij noemt, zijn
de kanalen die bestaan, de labels die hij uitlegt, dragen dezelfde helptekst die gebruikers
onder elk veld zien, en de verplichte velden per documenttype worden van de inhoudstypen
afgelezen. Hernoem een kanaal, draai hem opnieuw, en de pagina zegt het nieuwe.

Hij is geschreven voor degene die een catalogus uploadt. Geen groepsnamen, geen interne
kolomnamen, geen inhoudstypen of sitekolommen. Twee dingen uit de configuratie worden er
bewust buiten gehouden: de `note` op een container, die beveiligingsgroepen noemt, en de
`description` op een weergave, die over pijlers en gesynchroniseerde mappen gaat — hun
titels zijn op zichzelf duidelijk genoeg. Machtigingen staan er helemaal niet op: wie wat
mag zien is niet iets waar een gebruiker iets mee kan, en het uitleggen roept alleen de
vraag op waarom ze iets niet kunnen.


[`SharePoint-Handleiding.md`](SharePoint-Handleiding.md) is
geschreven voor de mensen die daadwerkelijk bestanden gaan uploaden — in het Nederlands,
zonder jargon, vijf minuten leestijd. Hij behandelt de drie manieren om een bestand toe te
voegen en waarom die zich anders gedragen, wat elk label betekent, en wat er gebeurt op het
moment dat je iets labelt.

Twee dingen daarin zijn als beheerder goed om te weten, omdat het de vragen zijn die
terugkomen:

- **Slepen-en-neerzetten en OneDrive-synchronisatie vragen niets.** Verplichte kolommen
  worden afgedwongen door het uploadformulier, niet door de bibliotheek. Bestanden die in
  bulk worden neergezet, landen met lege labels en een melding "Vereiste info" — ze worden
  niet tegengehouden. De weergave `Nog te taggen` is de opruimlijst, en de handleiding
  zegt gebruikers die af te werken met een meervoudige selectie en het detailvenster.
- **Een label is geen slot.** Vertrouwelijkheid op Vertrouwelijk zetten sluit niemand
  buiten; het is een afspraak, plus het signaal dat de nachtelijke audit gebruikt om
  overdeling te markeren. Toegang komt van de beveiligingsgroepen. De handleiding zegt dit
  in een kader, omdat gebruikers anders het tegendeel aannemen.

---

## Terugdraaien

[`Remove-SharePointStructure.ps1`](Remove-SharePointStructure.ps1) haalt dezelfde
configuratie uit elkaar, het diepste eerst. **Het heeft de omgekeerde standaard van al het
andere hier: zonder `-Apply` verandert het niets.** `-WhatIf` vergeten bij een destructief
script is de gevaarlijke kant op, dus de veilige toestand is degene die je gratis krijgt.

```powershell
.\Remove-SharePointStructure.ps1                                   # wat er zou verdwijnen
.\Remove-SharePointStructure.ps1 -Scope Channels,Groups -Apply     # een deel ervan
.\Remove-SharePointStructure.ps1 -Scope All,Team -IncludeContent -Apply   # opnieuw beginnen
```

| `-Scope` | Verwijdert |
|---|---|
| `Tabs` | de bibliotheektabbladen die aan kanalen zijn toegevoegd |
| `Channels` | de geconfigureerde kanalen, en de bestanden in hun mappen |
| `Libraries` | de bibliotheken die een container bezit (`kind: Library`) |
| `ContentTypes` | eerst losgekoppeld van de lijsten, daarna verwijderd |
| `Columns` | de sitekolommen, op elke site in de configuratie |
| `TermSet` | de termenset, zijn groep en zijn termen |
| `Groups` | de Entra ID-beveiligingsgroepen |
| `Team` | de Microsoft 365-groep — de site, elke bibliotheek, elk bestand, elke chat |
| `All` | alles hierboven **behalve** `Team` |

`All` omvat nooit het team. Het hele team van een klant verwijderen is niet iets wat je
zou moeten krijgen door om "alles" te vragen — je moet het noemen, en daarna de naam van
het team typen om te bevestigen.

**Wat het weigert te doen:**

- Een bibliotheek of kanaalmap waar nog bestanden in staan, wordt overgeslagen tenzij
  `-IncludeContent`. Het aantal items wordt hoe dan ook gemeld.
- Het kanaal Algemeen en de eigen bibliotheek Documenten van het team worden nooit
  verwijderd.
- Een inhoudstype dat nog in gebruik is, wordt gemeld, niet geforceerd.

**Wat geen prullenbak terugbrengt:** het verwijderen van de termenset maakt de waarde
Leverancier wees op elk document dat er een had — het veld houdt een GUID over die nergens
naar verwijst. Een kolom verwijderen neemt de gegevens mee. Beide worden met die prijs
gemeld voordat ze draaien. Een verwijderde groep of verwijderd team wordt 30 dagen
voorlopig verwijderd bewaard; een verwijderd kanaal heeft een eigen prullenbak van 30
dagen; bestanden uit een verwijderde bibliotheek gaan naar de prullenbak van de site.

---

## Wat de scripts nooit doen

Bewuste weglatingen, elk met een reden:

| Nooit | Waarom |
|---|---|
| Een sitekolom, inhoudstype of veldkoppeling verwijderen | dat zou de metadata op bestaande documenten meenemen |
| Een term uit het termenarchief verwijderen | gelabelde documenten verwijzen naar term-GUID's; de termenset is bedoeld om vanuit de UI uit te breiden |
| Een roltoewijzing verwijderen die de configuratie niet noemt | tenzij je `-RemoveOtherPermissions` meegeeft — een uitzondering die iemand met opzet heeft gemaakt, is geen drift |
| Site-eigenaren of deellinkgroepen verwijderen | `-RemoveOtherPermissions` slaat die over; ze weghalen sluit de klant buiten zijn eigen bibliotheek |
| Een deellink intrekken | de audit meldt overdeling. Intrekken is een beslissing, en een beslissing hoort bij een mens |
| Een weergave de standaard maken | elk lid van het kanaal zou het meteen merken |

---

## Opmerkingen

- Auteur: Sjoerd Kanon
- `SharePointStructure.Common.ps1` wordt door alle vier de scripts gedot-sourcet. Dat is
  een bewuste uitzondering op de regel "elk script staat op zichzelf" elders in deze repo:
  deze vier delen één configuratieschema, en drie kopieën van de machtigingscode zouden
  binnen een maand uit elkaar lopen.
- Test eerst tegen een niet-productietenant. De inrichtingsscripts wijzigen machtigingen
  op een live team.
