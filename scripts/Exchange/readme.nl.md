[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../readme.nl.md) › [scripts](../readme.nl.md) › **Exchange**

# Exchange-scripts

Scripts voor het beheer van agenda's, mailboxen en distributiegroepen in Exchange Online.

---

## Aanmelden

Elk script meldt aan via [`Connect-M365.ps1`](../Startup/readme.nl.md#connect-m365ps1): **standaard gedelegeerd** — je meldt je aan als jezelf, met een apparaatcode als `load.config.ps1` `useDeviceCodeAuth` zet, en onder GDAP wordt de klanttenant bereikt met `-DelegatedOrganization` (`-Organization` geldt alleen voor app-only aanmelden). App-only op verzoek met `-ClientId` + `-CertificateThumbprint`, of `-AppOnly` om beide uit `graph.appid.json` te halen. Een bestaande sessie voor dezelfde tenant wordt hergebruikt, en een script verbreekt alleen wat het zelf heeft geopend. Alle scripts hier vereisen PowerShell 7.

Graph is de standaard. Waar een script nog Exchange Online PowerShell gebruikt, of nog standaard app-only werkt, is dit de reden:

| Script | Gebruikt | Standaard aanmelding | Waarom |
|--------|----------|----------------------|--------|
| `Test-MailboxPermissions`, `Test-DistributionGroupPermissions`, `Set-Distributionlist-dynamic-static` | Exchange Online | Gedelegeerd | Graph heeft geen API voor Full Access, Send As, Send on Behalf, `ManagedBy` of dynamische distributiegroepen |
| `Get-DistributionGroupMembers` | Exchange Online | Gedelegeerd | Graph toont groepsleden, maar geen dynamische distributiegroepen, `ManagedBy`-eigenaars of e-mailcontactpersonen zoals het rapport ze toont |
| `Test-CalendarPermissions`, `Set-Calendar-rights` | Exchange Online | Gedelegeerd | Graph's `calendarPermissions` bereikt de agenda van een andere gebruiker alleen app-only of als je er al rechten op hebt, en kent geen rollen als `PublishingEditor`; Exchange accepteert je beheerdersrol voor elke mailbox |
| `Get-ExternalForwards`, `Test-DkimConfig`, `Get-MessageTraceReport` | Exchange Online | Gedelegeerd | Geen Graph-API voor mailboxdoorsturing, DKIM of message trace |
| `Get-MailboxSizes` | Exchange Online | Gedelegeerd | Het Graph-rapport `getMailboxUsageDetail` is geaggregeerd, loopt een dag of meer achter en toont verborgen namen als de tenant gebruikersgegevens afschermt |
| `Move-InboxToArchive`, `Restore-MailboxMessages`, `Remove-PhishingMessage` (Graph-engine) | Graph (+ Exchange) | **App-only** (tijdelijke app) | Een gedelegeerd token bereikt de mailbox van een andere gebruiker alleen met Full Access erop (`Mail.ReadWrite.Shared`). `-Delegated` neemt die route als je Full Access hebt (of, in `Move-InboxToArchive`, krijgt) — niet onder GDAP |
| `Get-CalendarMappings`, `Convert-SharedCalendarToResource`, `Move-SharedCalendar` | Graph (+ Exchange) | **App-only** (tijdelijke app) | Kan niet anders: ze lezen de agendalijst van elke gebruiker of schrijven in een resourcemailbox, en daar komt geen gedelegeerd token bij. Geen `-Delegated` |
| `Migrate-Calendar` | Graph + Exchange | Gedelegeerd om te lezen, **app-only** om te schrijven | Graph leest groepsagenda's alleen gedelegeerd, en schrijft alleen app-only in de nieuwe mailbox |

De tijdelijke app wordt aangemaakt met een gedelegeerde aanmelding (Global Administrator of Privileged Role Administrator) en aan het eind van de run weer verwijderd. In `Get-CalendarMappings`, `Convert-SharedCalendarToResource`, `Move-SharedCalendar`, `Remove-PhishingMessage` en `Restore-MailboxMessages` is die aanmelding altijd een **apparaatcode via gewone REST**, wat `load.config.ps1` ook zegt: die scripts praten ook met Exchange Online, waarvan de MSAL botst met die van de Graph SDK, dus laden ze de SDK daar niet voor. `Move-InboxToArchive` en `Migrate-Calendar` gebruiken er `Connect-M365Graph` voor en volgen `load.config.ps1`.

---

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`Migrate-Calendar.ps1`](Migrate-Calendar.ps1) ([docs](#migrate-calendarps1)) | Een gedeelde M365-groepsagenda migreren naar een Room Mailbox |
| [`Move-SharedCalendar.ps1`](Move-SharedCalendar.ps1) ([docs](#move-sharedcalendarps1)) | **Alles in één**: een agenda op trefwoord zoeken, naar een resourcemailbox verplaatsen en laten zien wie moet overstappen — met één aanmelding |
| [`Convert-SharedCalendarToResource.ps1`](Convert-SharedCalendarToResource.ps1) ([docs](#convert-sharedcalendartoresourceps1)) | Een gedeelde agenda uit de mailbox van een gebruiker naar een eigen room-/equipmentmailbox verplaatsen — inclusief items, reeksen, bijlagen en rechten |
| [`Set-Calendar-rights.ps1`](Set-Calendar-rights.ps1) ([docs](#set-calendar-rightsps1)) | Een gebruiker rechten geven op een agendamap |
| [`Set-Distributionlist-dynamic-static.ps1`](Set-Distributionlist-dynamic-static.ps1) ([docs](#set-distributionlist-dynamic-staticps1)) | De leden van een dynamische distributiegroep omzetten naar een gewone (statische) groep |
| [`Move-InboxToArchive.ps1`](Move-InboxToArchive.ps1) ([docs](#move-inboxtoarchiveps1)) | Alle (of op datum gefilterde) berichten uit het Postvak IN van een mailbox naar de map Archief verplaatsen |
| [`Test-CalendarPermissions.ps1`](Test-CalendarPermissions.ps1) ([docs](#test-calendarpermissionsps1)) | Rechten op agendamappen auditen |
| [`Get-CalendarMappings.ps1`](Get-CalendarMappings.ps1) ([docs](#get-calendarmappingsps1)) | Waar elke agenda daadwerkelijk in Outlook gekoppeld is, naast de rechten erachter — of één agenda op trefwoord zoeken (`-Search balie`) |
| [`Test-MailboxPermissions.ps1`](Test-MailboxPermissions.ps1) ([docs](#test-mailboxpermissionsps1)) | Delegatie via Full Access, Send As en Send on Behalf auditen |
| [`Test-DistributionGroupPermissions.ps1`](Test-DistributionGroupPermissions.ps1) ([docs](#test-distributiongrouppermissionsps1)) | Beheerders, Send As, Send on Behalf en ledenaantallen van distributiegroepen auditen |
| [`Test-DkimConfig.ps1`](Test-DkimConfig.ps1) ([docs](#test-dkimconfigps1)) | DKIM-ondertekeningsconfiguratie en DNS-records valideren |
| [`Get-ExternalForwards.ps1`](Get-ExternalForwards.ps1) ([docs](#get-externalforwardsps1)) | Mailboxen met externe doorsturing auditen |
| [`Get-MailboxSizes.ps1`](Get-MailboxSizes.ps1) ([docs](#get-mailboxsizesps1)) | Rapport van mailboxgroottes en aantallen items |
| [`Get-DistributionGroupMembers.ps1`](Get-DistributionGroupMembers.ps1) ([docs](#get-distributiongroupmembersps1)) | Wie op welke distributielijst staat, als Excel-werkmap die de klant kan lezen — of alleen de lijsten met één adres (`-Member jan@contoso.com`) of een heel domein (`-Member @be.verizon.com`) |
| [`Get-MessageTraceReport.ps1`](Get-MessageTraceReport.ps1) ([docs](#get-messagetracereportps1)) | Traceren wie wat ontving, op welk exact tijdstip, en waarheen het werd doorgestuurd |
| [`Remove-PhishingMessage.ps1`](Remove-PhishingMessage.ps1) ([docs](#remove-phishingmessageps1)) | Een phishingbericht verwijderen uit één, meerdere of alle mailboxen — standaard als proefdraai |
| [`Restore-MailboxMessages.ps1`](Restore-MailboxMessages.ps1) ([docs](#restore-mailboxmessagesps1)) | Berichten die op een bepaalde dag verplaatst of verwijderd zijn terugzetten in hun oorspronkelijke map, en rapporteren **wie** ze verplaatste of verwijderde — standaard als voorbeeld |

---

### Migrate-Calendar.ps1

Migreert een gedeelde M365-groepsagenda naar een Room Mailbox. Lost het probleem op dat groepsleden bij elke agenda-afspraak een e-mailmelding krijgen — een Room Mailbox gebruikt hetzelfde boekingsmechanisme als een vergaderruimte: geen meldingen, automatisch accepteren, zichtbaar voor iedereen.

**Hoe het werkt**

1. Meldt je bij Graph aan als jezelf (gedelegeerd) en bij Exchange Online
2. Leest de afspraken uit de M365-groepsagenda — gedelegeerd, de enige manier die Graph toestaat
3. Maakt een **tijdelijke** App Registration aan met `Calendars.ReadWrite` (plus `Group.ReadWrite.All` met `-DeleteSourceGroup`), of gebruikt je eigen app
4. Maakt een Room Mailbox aan als doelagenda, configureert AutoAccept en zet de Default-rechten op Reviewer
5. Kopieert de afspraken naar de Room Mailbox met het app-only-token
6. Verwijdert optioneel de bron-M365-groep
7. Verwijdert de tijdelijke App Registration — ook als de run mislukt

> Twee soorten toegang omdat Microsoft dat zo bepaalt: groepsagenda's kunnen alleen gedelegeerd gelezen worden, en een mailbox die niet van jou is alleen app-only beschreven.

Het secret van de tijdelijke app leeft twee uur, wordt nooit getoond, en de app wordt aan het eind verwijderd. (Vóór oktober 2026 maakte het script een permanente app `HolidaysCalendarMigration` aan en toonde het het secret.) De gedelegeerde aanmelding volgt `load.config.ps1`: apparaatcode met `useDeviceCodeAuth`, de GDAP-klant onder GDAP.

**Parameters**

| Parameter | Verplicht | Standaard | Omschrijving |
|-----------|----------|---------|-------------|
| `-TenantId` | Nee | GDAP-klant / je aanmelding | Tenant-ID of domein |
| `-AdminUPN` | Nee | — | Beheerder die het uitvoert (moet lid zijn van de brongroep). Alleen gebruikt om te waarschuwen als je met een ander account bent aangemeld |
| `-ClientId` | Nee | — | Je eigen App Registration om te schrijven, met `-ClientSecret` of `-CertificateThumbprint`. Zonder deze parameter wordt een tijdelijke aangemaakt en weer verwijderd |
| `-ClientSecret` | Nee | — | Client secret voor `-ClientId` |
| `-CertificateThumbprint` | Nee | — | Certificaatvingerafdruk voor `-ClientId` |
| `-AppOnly` | Nee | uit | Je eigen app met `ClientId` en `CertificateThumbprint` uit `graph.appid.json` |
| `-AppName` | Nee | `CalendarMigration-Temp` | Naamvoorvoegsel van de tijdelijke App Registration |
| `-SourceGroupMail` | Nee | — | E-mailadres van de bron-M365-groep |
| `-SourceGroupDisplayName` | Nee | — | Weergavenaam van de brongroep (gebruikt als alternatieve zoekmethode) |
| `-DestinationType` | Nee | `Room` | Type doelmailbox: `Room` of `Shared` |
| `-DestinationDisplayName` | Nee | `Holidays Calendar` | Weergavenaam van de doelmailbox |
| `-DestinationAlias` | Nee | `holidays-calendar` | Alias van de doelmailbox |
| `-DestinationEmail` | Nee | — | SMTP-adres van de doelmailbox |
| `-DaysBack` | Nee | `365` | Aantal dagen terug voor het ophalen van afspraken |
| `-DaysForward` | Nee | `730` | Aantal dagen vooruit voor het ophalen van afspraken |
| `-DeleteSourceGroup` | Nee | `$false` | De M365-groep na de migratie verwijderen |

**Voorbeelden**

```powershell
# Tijdelijke App Registration, aan het eind weer verwijderd
.\Migrate-Calendar.ps1 `
    -TenantId     "contoso.onmicrosoft.com" `
    -AdminUPN     "admin@contoso.com" `
    -SourceGroupMail "holidays@contoso.com"

# Je eigen App Registration (applicatiemachtiging Calendars.ReadWrite)
.\Migrate-Calendar.ps1 `
    -TenantId     "contoso.onmicrosoft.com" `
    -AppOnly `
    -SourceGroupMail "holidays@contoso.com"

# Proefdraai — geen wijzigingen
.\Migrate-Calendar.ps1 `
    -TenantId     "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -AdminUPN     "admin@contoso.com" `
    -SourceGroupMail "holidays@contoso.com" `
    -WhatIf
```

**Vereiste rechten**

| Recht | Doel |
|-----------|---------|
| Exchange Admin of Global Admin | Room Mailbox aanmaken |
| Global Administrator of Privileged Role Administrator | De tijdelijke App Registration aanmaken en haar machtiging toekennen (niet nodig met je eigen app) |
| Lid van de bron-M365-groep | Groepsagenda lezen via gedelegeerde toegang |

**Vereiste modules:** `ExchangeOnlineManagement`, `Microsoft.Graph.Authentication`, `.Applications`, `.Calendar`, `.Groups` — geïnstalleerd door `scripts\Startup\Install-Modules.ps1`.

---

### Set-Calendar-rights.ps1

Geeft een gebruiker toegangsrechten op de agenda van een andere gebruiker in Exchange Online, of wijzigt de rechten die de gebruiker daar al heeft. De agenda wordt op maptype gevonden, dus de taal van de mailbox maakt niet uit (`\Agenda`, `\Calendrier`, `\Calendar`, ...). Maakt zelf verbinding met Exchange Online (standaard gedelegeerd).

Blijft op Exchange Online: Graph's `calendarPermissions` kan de agenda van een andere gebruiker alleen app-only wijzigen of als je er al rechten op hebt, en kent geen rollen als `PublishingEditor` of `Contributor`. Het standaarddomein wordt nog steeds met `Get-AcceptedDomain` gelezen — de Exchange-sessie is toch al open, een Graph-aanmelding alleen voor `/domains` zou niets toevoegen.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-User` | Ja | Gebruiker die de rechten krijgt: UPN, of naam zonder domein (het standaarddomein wordt toegevoegd) |
| `-TargetMailbox` | Ja | Mailbox waarvan de agenda gedeeld wordt: UPN, of naam zonder domein |
| `-AccessRights` | Ja | Toegangsniveau (zie de tabel hieronder) |
| `-TenantId` | Nee | Tenant-ID of domein. Standaard: de GDAP-klant (`authMode = 'GDAP'` in `load.config.ps1`), anders je eigen tenant |
| `-ClientId` / `-CertificateThumbprint` | Nee | App-only aanmelden met je eigen app (vereist `Exchange.ManageAsApp` en een Exchange-rol). Zonder: gedelegeerd, als jezelf |
| `-AppOnly` | Nee | App-only met de `ClientId` en `CertificateThumbprint` van de tenant uit `graph.appid.json` |

**Toegangsniveaus**

| Waarde | Omschrijving |
|-------|-------------|
| `Owner` | Volledig beheer, inclusief verwijderen en mapbeheer |
| `PublishingEditor` | Lezen, aanmaken, wijzigen, verwijderen en submappen aanmaken |
| `Editor` | Lezen, aanmaken, wijzigen en verwijderen |
| `Author` | Lezen en aanmaken, eigen items wijzigen/verwijderen |
| `Reviewer` | Alleen lezen |
| `AvailabilityOnly` | Alleen vrij/bezet |
| `LimitedDetails` | Vrij/bezet met beperkte details |

**Voorbeelden**

```powershell
# Reviewer-rechten toekennen
.\Set-Calendar-rights.ps1 -User j.doe -TargetMailbox a.smith -AccessRights Reviewer

# Proefdraai, met volledige adressen
.\Set-Calendar-rights.ps1 -User j.doe@contoso.com -TargetMailbox a.smith@contoso.com -AccessRights Editor -WhatIf
```

---

### Set-Distributionlist-dynamic-static.ps1

Bepaalt welke leden op dit moment aan het filter van een dynamische distributiegroep voldoen en kopieert ze naar een gewone (statische) distributiegroep — de doelgroep wordt aangemaakt als die nog niet bestaat. Exporteert de gevonden ledenlijst ook naar CSV. Maakt zelf verbinding met Exchange Online (standaard gedelegeerd); een bestaande sessie voor dezelfde tenant wordt hergebruikt.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-DynamicGroupIdentity` | Ja | Dynamische brondistributiegroep (naam, alias, DN of SMTP-adres) |
| `-TargetGroupIdentity` | Ja | Gewone doeldistributiegroep — wordt aangemaakt als die niet bestaat |
| `-TargetDisplayName` | Nee | Weergavenaam voor een nieuwe doelgroep (standaard: `<DynamicDisplayName> Static`) |
| `-TargetAlias` | Nee | Alias voor een nieuwe doelgroep (standaard: `<DynamicAlias>-static`) |
| `-TargetPrimarySmtpAddress` | Nee | SMTP-adres voor een nieuwe doelgroep (standaard: het huidige primaire SMTP-adres van de dynamische groep) |
| `-CopyManagersFromDynamic` | Nee | `ManagedBy`-eigenaren van de dynamische groep naar de doelgroep kopiëren (standaard: aan) |
| `-DisableCopyManagersFromDynamic` | Nee | Het kopiëren van `ManagedBy`-eigenaren uitschakelen |
| `-MakeDynamicAddressTemporary` | Nee | De dynamische groep eerst een tijdelijk primair SMTP-adres geven, zodat het adres vrijkomt voor de doelgroep (standaard: aan) |
| `-DisableMakeDynamicAddressTemporary` | Nee | De automatische tijdelijke SMTP-wijziging uitschakelen |
| `-ClearTargetMembers` | Nee | Bestaande leden van de doelgroep verwijderen voordat de gevonden dynamische leden worden toegevoegd |
| `-ExportCsvPath` | Nee | CSV-exportpad voor de gevonden leden (standaard: `C:\Temp\DynamicGroupMembers_<timestamp>.csv` op Windows, `~/Downloads` op Linux/macOS) |
| `-SkipMemberAdd` | Nee | Alleen leden bepalen en exporteren, de doelgroep niet wijzigen |
| `-RenameDynamicGroupTo` | Nee | De dynamische brondistributiegroep na verwerking hernoemen |
| `-TenantId` | Nee | Tenant-ID of domein. Standaard: de GDAP-klant (`authMode = 'GDAP'` in `load.config.ps1`), anders je eigen tenant |
| `-ClientId` / `-CertificateThumbprint` | Nee | App-only aanmelden met je eigen app (vereist `Exchange.ManageAsApp` en een Exchange-rol). Zonder: gedelegeerd, als jezelf |
| `-AppOnly` | Nee | App-only met de `ClientId` en `CertificateThumbprint` van de tenant uit `graph.appid.json` |

**Voorbeelden**

```powershell
# Dynamische groep uitlezen en gewone groep vullen
.\Set-Distributionlist-dynamic-static.ps1 -DynamicGroupIdentity "All Sales" -TargetGroupIdentity "All Sales Static"

# Leden van de doelgroep volledig vernieuwen
.\Set-Distributionlist-dynamic-static.ps1 -DynamicGroupIdentity sales@contoso.com -TargetGroupIdentity sales-static@contoso.com -ClearTargetMembers

# Proefdraai
.\Set-Distributionlist-dynamic-static.ps1 -DynamicGroupIdentity "All Staff" -TargetGroupIdentity "All Staff Static" -WhatIf

# Omzetten en de oorspronkelijke dynamische groep hernoemen
.\Set-Distributionlist-dynamic-static.ps1 -DynamicGroupIdentity "All Sales" -TargetGroupIdentity "All Sales Static" -RenameDynamicGroupTo "All Sales (Legacy Dynamic)"

# Het SMTP-adres van de dynamische groep vrijmaken voor de nieuwe statische groep
.\Set-Distributionlist-dynamic-static.ps1 -DynamicGroupIdentity "All Sales" -TargetGroupIdentity "All Sales Static" -MakeDynamicAddressTemporary
```

**Opmerkingen**
- Vereist de module ExchangeOnlineManagement; het script maakt zelf verbinding
- Dynamische distributiegroepen zijn Exchange-objecten zonder Graph-API; dit script gebruikt Exchange-cmdlets, geen Graph

---

### Move-InboxToArchive.ps1

Verplaatst elk bericht in het Postvak IN van een mailbox naar de map Archief — dezelfde map waar de knop "Archiveren" in Outlook naartoe verplaatst. Optioneel kun je het bereik beperken tot een datumbereik (`-After` / `-Before`). Gebruikt het `$batch`-endpoint van Microsoft Graph om berichten in batches van 20 te verplaatsen, met retry/backoff bij throttling (429/503). Standaard draait het veilig in voorbeeldmodus — geef `-Apply` mee om de berichten echt te verplaatsen.

**Authenticatie (standaard: automatisch, geen Full Access nodig)**

App-only is hier de standaard omdat een gedelegeerd Graph-token de mailbox van een andere gebruiker alleen bereikt met Full Access erop, en die krijg je niet met een Exchange-beheerdersrol.

Standaard archiveert het script elke mailbox in de tenant zonder dat je er Full Access op nodig hebt. Het maakt gedelegeerd verbinding via `Connect-M365Graph` (`Application.ReadWrite.All` + `AppRoleAssignment.ReadWrite.All`; apparaatcode en GDAP-klant volgens `load.config.ps1`, een bestaande sessie met die scopes wordt hergebruikt), maakt een kortlevende tijdelijke App Registration aan, kent die zelf de applicatiemachtiging `Mail.ReadWrite` toe (geen apart admin-consentscherm — de gedelegeerde rol regelt dat), gebruikt die voor de mailboxbewerkingen en verwijdert haar weer als het script klaar is. Dit volgt hetzelfde patroon met een tijdelijke app als `Get-SharePointStorageReport.ps1` / `Remove-SharePointFileVersionsByDate.ps1`. Vereist Global Administrator of Privileged Role Administrator voor die eenmalige setup, en de module `Microsoft.Graph.Applications`.

- `-Delegated` slaat dat allemaal over en gebruikt in plaats daarvan een gewone gedelegeerde `Mail.ReadWrite` + `Mail.ReadWrite.Shared`-sessie (`Mail.ReadWrite.Shared` is wat de mailbox van een andere gebruiker bereikt) — daarvoor heb je Exchange Admin nodig, geen rechten om Entra-apps aan te maken. **Niet onder GDAP**: een partneraccount staat niet in de directory van de klant en kan dus geen Full Access krijgen; het script stopt met die melding. Voor een andere mailbox dan die van de aangemelde gebruiker zelf maakt het script verbinding met Exchange Online (`Connect-M365Exchange`), geeft dat account tijdelijk Full Access, pollt `Get-MailboxPermission` tot het recht echt zichtbaar is (tot ~3 minuten — rechtenwijzigingen in Exchange Online worden niet direct doorgevoerd), archiveert en trekt het recht daarna weer in (met een paar nieuwe pogingen, omdat de intrekking evengoed op een domain controller kan uitkomen die nog niet bij is).
  > **Bekende beperking:** `Get-MailboxPermission` toont de eigen status van Exchange vrijwel direct, maar de autorisatiecache van Microsoft Graph voor gedelegeerde mailboxtoegang kan daar tot **~60 minuten** op achterlopen — dit is een beperking aan de kant van Microsoft. Als het lezen van het Postvak IN na het pollen van Full Access nog steeds een 403 geeft, blijft het script het opnieuw proberen (met 60 s ertussen) tot een deadline van **`-MaxWaitMinutes`** (standaard 65, wat het door Microsoft gedocumenteerde worstcasescenario dekt) — het Full Access-recht blijft de hele wachttijd staan, omdat intrekken en opnieuw toekennen tussen de pogingen de doorvoerklok zou resetten. Verhoog `-MaxWaitMinutes` als 65 niet genoeg is, of laat `-Delegated` weg om de standaard app-only-modus te gebruiken, die zo'n vertraging niet heeft.
- `-ClientId` + `-ClientSecret`/`-CertificateThumbprint`, of `-AppOnly` (uit `graph.appid.json`), hergebruikt je eigen bestaande App Registration in plaats van een tijdelijke aan te maken — die app moet al de applicatiemachtiging `Mail.ReadWrite` hebben (met admin consent).

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|--------------|
| `-Mailbox` | Ja | UPN of object-ID van de mailbox waarvan het Postvak IN gearchiveerd moet worden |
| `-After` | Nee | Alleen berichten archiveren die op of na deze datum zijn ontvangen |
| `-Before` | Nee | Alleen berichten archiveren die vóór deze datum zijn ontvangen |
| `-TenantId` | Nee | Entra ID-tenant-ID (GUID) **of** een geverifieerd domein van de tenant (bijv. `contoso.com`) — beide werken. Optioneel als je al verbonden bent of als het af te leiden is uit een GDAP-klanttenantcontext; vereist voor app-only-authenticatie als het niet af te leiden is |
| `-ClientId` | Nee | Client-ID van een bestaande App Registration voor app-only-authenticatie — slaat de automatische tijdelijke app over. Gebruik samen met `-TenantId` en `-ClientSecret` of `-CertificateThumbprint` |
| `-ClientSecret` | Nee | Client secret voor de App Registration in `-ClientId` |
| `-CertificateThumbprint` | Nee | Certificaatvingerafdruk voor de App Registration in `-ClientId` |
| `-AppOnly` | Nee | Je eigen app met `ClientId` en `CertificateThumbprint` uit `graph.appid.json` |
| `-Delegated` | Nee | De automatische tijdelijke app-only-setup overslaan en gedelegeerd (`Mail.ReadWrite` + `Mail.ReadWrite.Shared`) verbinden. Voor andere mailboxen wordt tijdelijke Full Access via Exchange Online automatisch toegekend, gepolld en ingetrokken (vereist Exchange Admin). Niet onder GDAP |
| `-MaxWaitMinutes` | Nee | Alleen bij `-Delegated`. Hoe lang het script blijft proberen terwijl het wacht tot Graph het Full Access-recht honoreert, voordat het opgeeft en het recht intrekt. Standaard `65` |
| `-Apply` | Nee | De berichten echt verplaatsen. Zonder deze switch meldt het script alleen hoeveel berichten gearchiveerd zouden worden |

**Voorbeelden**

```powershell
# Voorbeeld — automatische app-only-setup, meldt het aantal, wijzigt niets
.\Move-InboxToArchive.ps1 -Mailbox "user@contoso.com"

# Alles in het Postvak IN archiveren
.\Move-InboxToArchive.ps1 -Mailbox "user@contoso.com" -Apply

# Alleen berichten ontvangen vóór 2025
.\Move-InboxToArchive.ps1 -Mailbox "user@contoso.com" -Before (Get-Date "2025-01-01") -Apply

# Alleen berichten ontvangen in 2024
.\Move-InboxToArchive.ps1 -Mailbox "user@contoso.com" -After (Get-Date "2024-01-01") -Before (Get-Date "2025-01-01") -Apply

# Gedelegeerd — tijdelijke Full Access via Exchange Online automatisch toekennen, pollen en intrekken
# in plaats van de Entra app-only-setup (vereist Exchange Admin, geen Global Admin)
.\Move-InboxToArchive.ps1 -Mailbox "user@contoso.com" -Delegated -Apply

# Gedelegeerd, bereid om Microsofts volledige worstcasescenario van ~90 minuten af te wachten
# tot Graph het Full Access-recht honoreert, in plaats van de standaard 65 minuten
.\Move-InboxToArchive.ps1 -Mailbox "user@contoso.com" -Delegated -MaxWaitMinutes 90 -Apply

# Een bestaande App Registration hergebruiken in plaats van een tijdelijke aan te maken.
# -TenantId accepteert het domein van de tenant in plaats van de GUID.
.\Move-InboxToArchive.ps1 -Mailbox "user@contoso.com" -TenantId "contoso.com" `
    -ClientId "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" -ClientSecret "your-client-secret" -Apply
```

**Opmerkingen**
- Mailboxen worden altijd via Microsoft Graph gelezen en berichten altijd via Graph verplaatst, niet via Exchange Online-cmdlets — vereist `Microsoft.Graph.Authentication` (en `Microsoft.Graph.Applications` voor de standaard automatische modus met tijdelijke app, of `ExchangeOnlineManagement` voor het tijdelijke Full Access-recht van `-Delegated`). Aan het eind worden alleen de sessies verbroken die het script zelf heeft geopend
- Toont voortgang met tijdstempel tijdens het pagineren door het Postvak IN, tijdens het verplaatsen van batches (`[HH:mm:ss] N / total moved (...%)`) en tijdens het pollen op de doorvoering van Full Access in `-Delegated`-modus
- GDAP-bewust: onder een GDAP-sessie (`$global:authMode -eq 'GDAP'`, ingesteld via `Connect-Tenant` / `load.ps1`) wordt `-TenantId`, als die ontbreekt, automatisch afgeleid uit de geselecteerde klanttenant (`$global:cid`) — via `Resolve-M365TenantId` in `Connect-M365.ps1`. Ook `$env:M365_CUSTOMER_TENANTID` / `$env:M365_AUTH_MODE` worden gerespecteerd

**Vereiste modules:** `Microsoft.Graph.Authentication`, `Microsoft.Graph.Applications`, `ExchangeOnlineManagement` — geïnstalleerd door `scripts\Startup\Install-Modules.ps1`.

---

### Move-SharedCalendar.ps1

**Alles in één:** van "waar staat de Balie-agenda?" naar "die heeft een eigen resourcemailbox" in één run, met één aanmelding. Het script koppelt [`Get-CalendarMappings.ps1`](Get-CalendarMappings.ps1) ([docs](#get-calendarmappingsps1)) en [`Convert-SharedCalendarToResource.ps1`](Convert-SharedCalendarToResource.ps1) ([docs](#convert-sharedcalendartoresourceps1)) aan elkaar — beide moeten in dezelfde map staan.

| Stap | |
|------|--|
| 1. Zoeken | `Get-CalendarMappings.ps1 -Search <keyword>`: waar de agenda staat, wie hem in Outlook heeft en wie er rechten op heeft |
| 2. Kiezen | De gevonden agenda die verplaatst kan worden. Meerdere treffers: kies er één uit een genummerde lijst, of beperk de keuze met `-Owner`. Een niet-interactieve run gokt nooit — die toont de kandidaten en stopt |
| 3. Verplaatsen | `Convert-SharedCalendarToResource.ps1`: voorbeeld, daarna drie vragen — doorgaan? uitnodigingen versturen? origineel verwijderen? `-Apply` slaat de voorbeeldronde over |
| 4. Melden | Wie de oude agenda in Outlook had en wie alleen rechten had: de mensen die moeten overstappen |

**Eén aanmelding.** Een tijdelijke App Registration met alles wat beide scripts nodig hebben (`Calendars.ReadWrite`, `User.Read.All`, `Group.Read.All`, `MailboxSettings.ReadWrite`) wordt één keer aangemaakt, aan beide doorgegeven en aan het eind verwijderd — ook als er iets misgaat. Ook met Exchange Online wordt maar één keer verbinding gemaakt — gedelegeerd via `Connect-M365Exchange` (apparaatcode en GDAP-klant volgens `load.config.ps1`), en beide scripts hergebruiken die sessie. Een bestaande app-only Graph-sessie of `-ClientId` / `-ClientSecret` wordt in plaats daarvan gebruikt als die is opgegeven.

App-only Graph is de standaard omdat het niet anders kan: de agenda vinden leest de agendalijst van elke mailbox en de verhuizing schrijft in een nieuwe resourcemailbox, en daar komt geen gedelegeerd token bij. De aanmelding voor de tijdelijke app is altijd een apparaatcode via gewone REST, omdat Exchange dan al verbonden is en de MSAL van de Graph SDK botst met die van Exchange — om dezelfde reden werkt je eigen app hier alleen met `-ClientSecret`, niet met een certificaat of `-AppOnly`.

Een mailbox waarvan de **hoofd**agenda overeenkomt (een `balie@`-account dat zelf de gedeelde agenda is) kan niet worden verplaatst; het script meldt dat en noemt het alternatief ter plaatse, `Set-Mailbox -Type Room`.

Een run die iets wijzigt, wordt gelogd naar `SharedCalendarMove_<timestamp>.log`, naast het koppelingsrapport en de back-up.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-Search` | Ja | Trefwoord: naam/adres van de eigenaar of naam van de agenda. Alias `-Keyword` |
| `-Owner` | Nee | Beperkt meerdere treffers tot één eigenaar (deel van de naam of het adres) |
| `-ResourceType` | Nee | `Room` (standaard) of `Equipment` |
| `-ResourceName` / `-ResourceAddress` | Nee | Naam en adres van de nieuwe mailbox (standaard: de naam van de agenda, op het domein van de eigenaar) |
| `-SourceOwnerRights` | Nee | Rechten voor de oorspronkelijke eigenaar: `Owner` (standaard) … `None` — gebruik `None` voor een gearchiveerde mailbox |
| `-SendSharingInvitation` | Nee | Gebruikers uitnodigen voor de nieuwe agenda (wordt gevraagd als het niet is opgegeven) |
| `-Apply` | Nee | Doorgaan zonder de voorbeeldronde |
| `-RemoveSourceCalendar` | Nee | Het origineel verwijderen na een foutloze verificatie (wordt gevraagd als het niet is opgegeven) |
| `-Force` | Nee | De getypte bevestiging overslaan — vereist om onbeheerd te verwijderen |
| `-OutputPath` | Nee | Map voor rapport, back-up en log (standaard `C:\Temp`) |
| `-TenantId` / `-ClientId` / `-ClientSecret` | Nee | Tenant, of je eigen App Registration in plaats van een tijdelijke |

**Voorbeelden**

```powershell
# Zoeken, voorbeeld bekijken, de vragen beantwoorden
.\Move-SharedCalendar.ps1 -Search balie

# "Balie planning" in een gearchiveerde mailbox: Equipment-mailbox, gebruikers uitgenodigd, origineel voorlopig behouden
.\Move-SharedCalendar.ps1 -Search "balie planning" -ResourceType Equipment -SourceOwnerRights None `
    -SendSharingInvitation -Apply

# Zodra iedereen is overgestapt: hetzelfde commando verwijdert het origineel
.\Move-SharedCalendar.ps1 -Search "balie planning" -ResourceType Equipment -SourceOwnerRights None `
    -Apply -RemoveSourceCalendar
```

De tweede run slaat elk item over dat al gekopieerd is, kopieert wat er intussen is bijgekomen en verwijdert pas daarna het origineel.

---

### Convert-SharedCalendarToResource.ps1

Verplaatst een gedeelde agenda uit de mailbox van een gebruiker naar een **eigen resourcemailbox** (Room of Equipment), met elk item en elk recht, en verwijdert daarna — op verzoek — het origineel. Gebouwd voor de typische "Balie"-agenda: een extra agenda in de mailbox van één persoon die de hele receptie gebruikt, en die met die persoon vertrekt.

**Standaard een voorbeeld.** Zonder `-Apply` leest en rapporteert het script alleen: hoeveel items, reeksen en uitzonderingen, welke rechten het zou overnemen en welke mailbox het zou aanmaken. Het origineel wordt alleen verwijderd met `-RemoveSourceCalendar`, alleen nadat elk item een geverifieerde kopie heeft, en alleen nadat je de naam van de agenda hebt getypt (over te slaan met `-Force`).

**Wat er gebeurt bij `-Apply`**

| Stap | |
|------|--|
| Back-up | Elk item (inclusief inhoud), elk voorkomen van een reeks en elk recht naar `calendar-backup.json`, voordat er iets wordt aangemaakt |
| Mailbox | Room- (standaard) of Equipment-mailbox met de taal en tijdzone van de bronmailbox. Agendaverwerking voor een gedeelde agenda: automatisch accepteren, overlappende items toegestaan, niets in een item herschreven, boekingsvenster 1080 dagen (het maximum van de dienst) |
| Rechten | Elk recht met de **exacte** Exchange-toegangsrechten (ook aangepaste rechten), `Default` en `Anonymous` zoals ze waren, de oorspronkelijke eigenaar als `-SourceOwnerRights` (standaard `Owner`). `-SendSharingInvitation` verstuurt de gebruikelijke mail "heeft een agenda met je gedeeld" |
| Categorieën | De gebruikte categorieën worden met hun kleur aangemaakt in de nieuwe mailbox |
| Items | Elk item gekopieerd — zie hieronder |
| Verificatie | Elk bronitem moet een volledige kopie hebben |
| Verwijderen | Alleen met `-RemoveSourceCalendar` en een foutloze verificatie |

**Hoe items worden gekopieerd**

- **Reeksen blijven reeksen.** Verplaatste of bewerkte voorkomens worden op de kopie toegepast en geannuleerde voorkomens worden erin geannuleerd, door beide reeksen voorkomen voor voorkomen naast elkaar te leggen. Voor een reeks zonder einddatum gebeurt dat tot `-SeriesHorizonDays` (1095) vooruit. Als de twee reeksen niet overeenkomen, laat het script die reeks met rust en meldt dat, in plaats van de verkeerde voorkomens te annuleren
- **Tijden behouden hun tijdzone.** Graph geeft UTC terug; elk item wordt teruggeschreven in de zone waarin het is aangemaakt, zodat een wekelijks item om 9:00 na de overgang naar zomer- of wintertijd nog steeds om 9:00 staat
- **Niemand wordt uitgenodigd.** Een vergadering kopiëren met de deelnemers erbij zou elke deelnemer een nieuwe uitnodiging vanuit de resourcemailbox sturen. Organisator en deelnemers worden in plaats daarvan onderaan de inhoud vermeld; de resourcemailbox is de organisator van elke kopie
- **Bijlagen** tot 3 MB worden gekopieerd, inline afbeeldingen inbegrepen. Grotere bestanden en bijgevoegde Outlook-items worden in de back-upmap opgeslagen en aan het eind vermeld
- **Opnieuw uit te voeren.** Elke kopie draagt de ID van het bronitem in een verborgen eigenschap. Een run die halverwege stopt, zet je voort door hetzelfde commando opnieuw uit te voeren: volledige kopieën worden overgeslagen, een half gekopieerde reeks wordt verwijderd en opnieuw gekopieerd

**Niet overgenomen, wel gerapporteerd:** gemachtigdevlaggen (een resourcemailbox heeft geen gemachtigden), mensen buiten de tenant (handmatig opnieuw delen), rechten van verwijderde accounts, bijlagen bij individuele uitzonderingen in een reeks. Een gepubliceerde agenda (`Anonymous`) krijgt een nieuwe link.

> **De hoofdagenda kan niet op deze manier worden omgezet.** Als de hele mailbox de gedeelde agenda *is* (een `balie@`-gebruikersaccount), zet je hem in plaats daarvan ter plaatse om — dan blijft alles behouden: `Set-Mailbox balie@contoso.com -Type Room`. Het script weigert de hoofdagenda en verwijst daarnaar.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-Mailbox` | Ja | Gebruikersmailbox waarin de agenda staat |
| `-Calendar` | Ja | Naam van de agenda zoals weergegeven in Outlook (bijv. `Balie`). Moet eigendom zijn van die gebruiker en mag niet diens hoofdagenda zijn |
| `-ResourceName` | Nee | Weergavenaam van de nieuwe mailbox (standaard: de naam van de agenda) |
| `-ResourceAddress` | Nee | SMTP-adres (standaard: de naam als alias op het domein van de gebruiker). Een bestaande room-/equipmentmailbox op dit adres wordt hergebruikt |
| `-ResourceType` | Nee | `Room` (standaard) of `Equipment` |
| `-SourceOwnerRights` | Nee | Rechten voor de oorspronkelijke eigenaar: `Owner` (standaard), `PublishingEditor`, `Editor`, `Reviewer`, `None` |
| `-SendSharingInvitation` | Nee | Gebruikers een deeluitnodiging sturen (alleen mogelijk voor Reviewer, Editor, LimitedDetails, AvailabilityOnly) |
| `-Apply` | Nee | Echt aanmaken, rechten toekennen en kopiëren. Zonder deze switch: voorbeeld |
| `-RemoveSourceCalendar` | Nee | Het origineel verwijderen na een foutloze verificatie. Vereist `-Apply` |
| `-Force` | Nee | De getypte bevestiging vóór het verwijderen overslaan. Een niet-interactieve sessie (scheduler, RMM) kan die niet typen, dus daar wordt het origineel alleen met `-Force` verwijderd |
| `-PassThru` | Nee | Een resultaatobject teruggeven (`ResourceAddress`, `Items`, `Verified`, `SourceRemoved`, `BackupPath`) voor een aanroepend script |
| `-SeriesHorizonDays` | Nee | Hoe ver vooruit uitzonderingen van open reeksen worden vergeleken (standaard 1095) |
| `-BackupPath` | Nee | Back-upmap (standaard `C:\Temp\CalendarConvert_<calendar>_<timestamp>`) |
| `-TenantId` | Nee | Tenant-ID of domein (standaard: de GDAP-klant, anders de tenant van de Exchange-sessie) |
| `-ClientId` / `-ClientSecret` / `-CertificateThumbprint` | Nee | Je eigen App Registration voor app-only Graph-toegang |
| `-AppOnly` | Nee | Je eigen app met `ClientId` en `CertificateThumbprint` uit `graph.appid.json` |

**Voorbeelden**

```powershell
# 1. Voorbeeld
.\Convert-SharedCalendarToResource.ps1 -Mailbox jan@contoso.com -Calendar Balie

# 2. De room mailbox aanmaken, alles kopiëren, de gebruikers uitnodigen - het origineel blijft staan
.\Convert-SharedCalendarToResource.ps1 -Mailbox jan@contoso.com -Calendar Balie -Apply -SendSharingInvitation

# 3. Zodra gebruikers zijn overgestapt: zoeken wie de oude agenda nog heeft, daarna verwijderen
.\Get-CalendarMappings.ps1 -Search Balie
.\Convert-SharedCalendarToResource.ps1 -Mailbox jan@contoso.com -Calendar Balie -Apply -RemoveSourceCalendar
```

Stap 3 voert eerst de kopie opnieuw uit: alles wat al gekopieerd is wordt overgeslagen, wat intussen aan het origineel is toegevoegd wordt gekopieerd, en pas daarna wordt het origineel verwijderd.

**Toegang**

| | |
|--|--|
| Exchange Online | Exchange Administrator (`New-Mailbox`, maprechten). Gedelegeerd via `Connect-M365Exchange` (apparaatcode en GDAP-klant volgens `load.config.ps1`); een bestaande sessie wordt hergebruikt |
| Graph | Applicatiemachtiging `Calendars.ReadWrite`, plus `MailboxSettings.ReadWrite` voor categoriekleuren (optioneel). Dezelfde drie routes als [`Remove-PhishingMessage.ps1`](Remove-PhishingMessage.ps1) ([docs](#remove-phishingmessageps1)): een bestaande app-only-sessie, een eigen App Registration (`-ClientId`, of `-AppOnly`), of een tijdelijke die wordt verwijderd als de run eindigt. Gewone REST, dus geen MSAL-conflict tussen Exchange en Graph — en daarom altijd een apparaatcode voor de aanmelding van de tijdelijke app |
| Waarom app-only | Kan niet anders: het script leest de agenda van één gebruiker en schrijft in een mailbox die het net heeft aangemaakt, en daar komt geen gedelegeerd token bij zonder expliciete rechten op beide — en Graph kan de categoriekleuren van een andere mailbox gedelegeerd helemaal niet lezen. Geen `-Delegated` |

**Opmerkingen**

- De kopie is een momentopname. Voer hem uit als het rustig is in de agenda en laat gebruikers direct daarna overstappen; stap 3 hierboven pakt op wat er tussendoor is bijgekomen
- Gebruikers die de oorspronkelijke agenda in hun lijst hadden, houden een vermelding over die niet meer werkt zodra het origineel is verwijderd — zoek ze eerst op met `Get-CalendarMappings.ps1 -Search`
- Verschilt van `Migrate-Calendar.ps1` (groepsagenda → room): dat script kopieert items zonder inhoud, reeksen of tijdzone. Dit script is bedoeld als getrouwe verhuizing

---

## Auditscripts

Maken automatisch verbinding met Exchange Online (standaard gedelegeerd, zie [Aanmelden](#aanmelden)); een bestaande sessie voor dezelfde tenant wordt hergebruikt en blijft open.

---

### Test-CalendarPermissions.ps1

Haalt de rechten op agendamappen op voor één of alle mailboxen. Gebruikt `FolderType` om de agendamap te vinden, ongeacht de taal van de mailbox (NL/FR/EN). Exporteert de resultaten naar CSV.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-Mailbox` | Nee | UPN van één mailbox. Zonder deze parameter worden alle gebruikers- en gedeelde mailboxen gecontroleerd |
| `-OutputPath` | Nee | Pad van het CSV-rapport (standaard: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Nee | Tenant-ID of domein. Standaard: de GDAP-klant (`authMode = 'GDAP'` in `load.config.ps1`), anders je eigen tenant |
| `-ClientId` / `-CertificateThumbprint` | Nee | App-only aanmelden met je eigen app (vereist `Exchange.ManageAsApp` en een Exchange-rol). Zonder: gedelegeerd, als jezelf |
| `-AppOnly` | Nee | App-only met de `ClientId` en `CertificateThumbprint` van de tenant uit `graph.appid.json` |

**Voorbeelden**

```powershell
# Alle mailboxen auditen
.\Test-CalendarPermissions.ps1

# Eén mailbox
.\Test-CalendarPermissions.ps1 -Mailbox "user@contoso.com"

# Eigen uitvoerpad
.\Test-CalendarPermissions.ps1 -OutputPath "C:\Reports\calendar.csv"
```

---

### Get-CalendarMappings.ps1

Toont **waar elke agenda gekoppeld is**: de agenda's die echt in de agendalijst van een gebruiker in Outlook staan, naast de rechten erachter. `Test-CalendarPermissions.ps1` beantwoordt "wie *mag* deze agenda openen"; dit script beantwoordt "waar *staat* hij" en markeert waar die twee niet overeenkomen. Alleen-lezen.

Op zoek naar één agenda? `-Search balie` vindt hem op trefwoord — zie *Zoeken op trefwoord* hieronder.

Voor elke mailbox leest het via Microsoft Graph:

- de **agendalijst** (`/users/{id}/calendars`). Elke agenda daarin die van iemand anders is, is een koppeling: een collega, een gedeelde mailbox, een ruimte, een Microsoft 365-groep of iemand buiten de organisatie
- de **rechten op de eigen hoofdagenda** (`/users/{id}/calendar/calendarPermissions`)

en voegt beide samen tot één rij per combinatie van agenda-eigenaar + gebruiker:

| Status | Betekenis |
|--------|---------|
| `Source` | Alleen bij `-Search`: de gevonden agenda staat in deze mailbox (de hoofdagenda of een secundaire agenda) |
| `Mapped` | In de agendalijst van de gebruiker, en de gebruiker heeft een expliciet recht |
| `MappedWithoutRight` | In de lijst, maar geen expliciet recht op de hoofdagenda van de eigenaar. De toegang komt dan van de organisatiebrede standaard, een groep of een secundaire agenda van de eigenaar — of het recht is ingetrokken en de vermelding is blijven staan (de gebruiker krijgt een foutmelding bij het openen) |
| `MappedGroupCalendar` | Een Microsoft 365-groepsagenda — toegang volgt het groepslidmaatschap |
| `MappedOwnerMissing` | De eigenaar bestaat niet meer in de tenant — een verouderde vermelding in de lijst van de gebruiker |
| `MappedExternal` | De eigenaar zit buiten de tenant |
| `NotMapped` | Expliciet recht, maar de agenda staat niet in de lijst van de gebruiker — kandidaat om op te schonen |
| `NotChecked` | Expliciet recht, maar de agendalijst van de gebruiker kon niet worden gelezen |
| `GrantedToGroup` | Een recht dat aan een groep is toegekend; de leden worden niet uitgevouwen |
| `GrantedToMissing` | Een recht voor een adres of account dat niet meer bestaat — kandidaat om op te schonen |
| `SharedExternally` | Een recht voor een adres buiten de tenant |
| `OrgWideDefault` | *My Organization* krijgt meer dan vrij/bezet — elke interne gebruiker kan de agenda openen |

> **Waarom Graph en geen Exchange Online PowerShell:** de Exchange-cmdlets zien mappen en de rechten daarop, niet de vermeldingen die een gebruiker aan zijn eigen agendalijst heeft toegevoegd. Die zijn alleen via Graph te lezen.

**Niet zichtbaar in dit rapport**

- **Full Access met AutoMapping** voegt een hele mailbox aan Outlook toe, inclusief de agenda. Dat is een mailboxrecht, geen agendavermelding — zie [`Test-MailboxPermissions.ps1`](Test-MailboxPermissions.ps1) ([docs](#test-mailboxpermissionsps1))
- Een agenda die geopend is in klassiek Outlook met *shared calendar improvements* uitgeschakeld, staat mogelijk alleen in dat Outlook-profiel en niet in de lijst die Graph teruggeeft
- Zonder `-Search` worden rechten vergeleken met de **hoofd**agenda van de eigenaar. Een secundaire agenda die de eigenaar heeft gedeeld, verschijnt als `MappedWithoutRight`; `-Search` leest de eigen rechten van een gevonden secundaire agenda

**Zoeken op trefwoord**

`-Search balie` (alias `-Keyword`) beantwoordt "waar staat de Balie-agenda?". Het trefwoord wordt hoofdletterongevoelig en op elke plek in de tekst vergeleken met:

- de **naam en elk adres** van de eigenaar — de gedeelde mailbox `balie@`, een ruimte, een groep met de naam *Balie-team*
- de **eigen naam** van de agenda — een secundaire agenda *Balie* in iemands mailbox

Jokertekens (`*`, `?`) worden gebruikt zoals opgegeven. Voor elke treffer toont het rapport waar de agenda staat (`Source`), elke mailbox die hem in de agendalijst heeft, en iedereen met een expliciet recht erop — voor een secundaire agenda de eigen rechten, die een normale run niet leest. De kolom `Calendar` geeft aan over welke agenda van de eigenaar een rij gaat.

```
Owner      Calendar       User Status              Rights MappedAs
-----      --------       ---- ------              ------ --------
Anna       Balie Planning      Source                     Balie Planning
Anna       Balie Planning Lisa Mapped              read   Balie Planning
Anna       Balie Planning Jan  NotMapped           write
Balie      Main                Source                     Agenda
Balie      Main           Kees Mapped              read   Balie
Balie      Main           Piet Mapped              write  Balie
Balie      Main           Lisa NotMapped           read
Balie-team                Kees MappedGroupCalendar        Balie-team
```

Elke agendalijst wordt nog steeds gelezen — een koppeling kan in elke mailbox staan — dus een zoekopdracht duurt voor de lijsten ongeveer even lang als een volledige scan, maar leest alleen de rechten van de gevonden agenda's.

> Een vermelding in de agendalijst bevat geen verwijzing terug naar de agenda waar ze vandaan komt. Een gedeelde secundaire agenda wordt daarom herkend aan een naam die overeenkomt met het trefwoord. Als een gebruiker hem onder een andere naam heeft staan, verschijnt hij als `NotMapped` met een opmerking die de vermelding noemt die het waarschijnlijk is (*"Has a calendar of this owner as 'Planning Anna' - probably this one"*).

**Bereik**

Zonder `-Mailbox` wordt elke mailbox in de tenant gescand — de enige manier om koppelingen te vinden die op de organisatiebrede standaard of op een groep berusten. Met `-Mailbox` wordt het rapport beperkt tot rijen waarin een van die mailboxen de **eigenaar of de gebruiker** is: hun eigen agendalijsten worden gelezen, plus de lijsten van iedereen met een expliciet recht op hun agenda. Voor een volledig antwoord op "waar is de agenda van X gekoppeld" gebruik je in plaats daarvan `-Search`. `-Mailbox` en `-Search` kunnen niet worden gecombineerd.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-Search` | Nee | Trefwoord om één agenda mee te zoeken (naam/adres van de eigenaar of naam van de agenda). Alias `-Keyword`. Niet te combineren met `-Mailbox` |
| `-Mailbox` | Nee | Een of meer mailboxadressen. Beperkt het rapport tot rijen waarin ze eigenaar of gebruiker zijn. Zonder deze parameter wordt elke mailbox gescand |
| `-OutputPath` | Nee | Pad van het CSV-rapport (standaard: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Nee | Tenant-ID of domein. Optioneel voor de route met de tijdelijke app — de aanmelding bepaalt dan de tenant, en die wordt getoond |
| `-ClientId` | Nee | Je eigen App Registration voor app-only Graph-toegang |
| `-ClientSecret` | Nee | Client secret voor `-ClientId` (gewone REST, geen Graph SDK) |
| `-CertificateThumbprint` | Nee | Certificaatvingerafdruk voor `-ClientId` (via `Connect-M365Graph`, dat een passende sessie hergebruikt en alleen verbreekt wat het zelf opende) |
| `-AppOnly` | Nee | Je eigen app met `ClientId` en `CertificateThumbprint` uit `graph.appid.json` |

**Voorbeelden**

```powershell
# Waar staat de Balie-agenda, en wie heeft hem gekoppeld?
.\Get-CalendarMappings.ps1 -Search balie

# Waar is elke agenda in de tenant gekoppeld?
.\Get-CalendarMappings.ps1 -TenantId contoso.com

# Waar is de agenda van Jan gekoppeld, en welke agenda's heeft Jan gekoppeld?
.\Get-CalendarMappings.ps1 -Mailbox jan@contoso.com

# Eigen App Registration
.\Get-CalendarMappings.ps1 -TenantId contoso.com -ClientId <appId> -ClientSecret <secret>
```

**Graph-toegang**

Vereist de applicatiemachtigingen `Calendars.Read` en `User.Read.All`, plus `Group.Read.All` om een groepsagenda te onderscheiden van een verwijderde mailbox (zonder die machtiging verschijnen groepsagenda's als `MappedOwnerMissing` met een opmerking die dat vermeldt). Verkregen op dezelfde drie manieren als bij [`Remove-PhishingMessage.ps1`](Remove-PhishingMessage.ps1) ([docs](#remove-phishingmessageps1)):

| # | Route | Wat er nodig is |
|---|-------|---------------|
| 1 | Een app-only Graph-sessie die je al had opgezet | Niets — wordt gebruikt zoals ze is |
| 2 | `-ClientId` + `-ClientSecret` of `-CertificateThumbprint`, of `-AppOnly` | Je eigen app met de machtigingen hierboven, met admin consent. Een ruimere machtiging (`Calendars.ReadWrite`, `Directory.Read.All`) wordt ook geaccepteerd |
| 3 | **Automatisch** — aanmelding met apparaatcode, een kortlevende App Registration die zichzelf de drie leesmachtigingen toekent en weer wordt verwijderd als de run eindigt (ook bij een fout) | Global Administrator of Privileged Role Administrator voor die aanmelding. Geen extra modules |

Er wordt geen verbinding met Exchange Online gemaakt, dus het MSAL-conflict tussen Exchange en Graph dat onder `Remove-PhishingMessage.ps1` wordt beschreven, speelt hier niet. GDAP-bewust zoals de andere Graph-scripts: onder een GDAP-sessie wordt `-TenantId` afgeleid uit de geselecteerde klanttenant.

**Waarom app-only, en geen `-Delegated`:** het rapport leest de agendalijst van elke mailbox. Een gedelegeerd token (`Calendars.Read.Shared`) ziet alleen de agenda's die met *jou* gedeeld zijn, niet wat andere gebruikers in hun eigen lijst zetten, dus een gedelegeerde run zou niets bruikbaars opleveren. De aanmelding van route 3 is altijd een apparaatcode via gewone REST, omdat `Move-SharedCalendar.ps1` dit script aanroept nadat het met Exchange verbonden is.

**Opmerkingen**

- Verzoeken gaan via Graph `$batch`, 20 mailboxen per aanroep. Items die gethrottled worden, worden opnieuw geprobeerd na de `Retry-After` die de dienst vraagt
- Gebruikers met een adres maar zonder Exchange Online-mailbox (404) worden overgeslagen en geteld. Een mailbox die niet gelezen kan worden, wordt apart vermeld en nooit gerapporteerd als "niets gekoppeld". Een 403 betekent daar meestal dat een Application Access Policy of RBAC for Applications de app beperkt
- Gedeelde mailboxen en ruimtes zijn uitgeschakelde accounts in Entra ID en worden dus meegenomen — geen filter op `accountEnabled`

---

### Test-MailboxPermissions.ps1

Auditeert alle drie de delegatietypen voor één of alle mailboxen:

- **Full Access** — gebruikers die de mailbox kunnen openen
- **Send As** — gebruikers die kunnen verzenden als de identiteit van de mailbox
- **Send on Behalf** — gebruikers die in `GrantSendOnBehalfTo` staan

Overgenomen vermeldingen en SELF-vermeldingen worden automatisch weggefilterd.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-Mailbox` | Nee | UPN van één mailbox. Zonder deze parameter worden alle gebruikers- en gedeelde mailboxen gecontroleerd |
| `-OutputPath` | Nee | Pad van het CSV-rapport (standaard: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Nee | Tenant-ID of domein. Standaard: de GDAP-klant (`authMode = 'GDAP'` in `load.config.ps1`), anders je eigen tenant |
| `-ClientId` / `-CertificateThumbprint` | Nee | App-only aanmelden met je eigen app (vereist `Exchange.ManageAsApp` en een Exchange-rol). Zonder: gedelegeerd, als jezelf |
| `-AppOnly` | Nee | App-only met de `ClientId` en `CertificateThumbprint` van de tenant uit `graph.appid.json` |

**Voorbeelden**

```powershell
# Alle mailboxen auditen
.\Test-MailboxPermissions.ps1

# Alleen een gedeelde mailbox
.\Test-MailboxPermissions.ps1 -Mailbox "shared@contoso.com"
```

---

### Test-DistributionGroupPermissions.ps1

Auditeert distributiegroepen en mail-enabled beveiligingsgroepen:

- **Instellingen** — aantal leden, beperkingen voor lid worden/uittreden, beleid voor externe afzenders
- **ManagedBy** — eigenaren/beheerders van de groep
- **Send As** — wie kan verzenden als de groep
- **Send on Behalf** — gemachtigden in `GrantSendOnBehalfTo`
- **Leden** (optioneel, gebruik `-IncludeMembers`)

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-Group` | Nee | Naam, alias of e-mailadres van één groep. Zonder deze parameter worden alle distributiegroepen geaudit |
| `-IncludeMembers` | Nee | Ook de individuele groepsleden in het rapport opnemen |
| `-OutputPath` | Nee | Pad van het CSV-rapport (standaard: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Nee | Tenant-ID of domein. Standaard: de GDAP-klant (`authMode = 'GDAP'` in `load.config.ps1`), anders je eigen tenant |
| `-ClientId` / `-CertificateThumbprint` | Nee | App-only aanmelden met je eigen app (vereist `Exchange.ManageAsApp` en een Exchange-rol). Zonder: gedelegeerd, als jezelf |
| `-AppOnly` | Nee | App-only met de `ClientId` en `CertificateThumbprint` van de tenant uit `graph.appid.json` |

**Voorbeelden**

```powershell
# Alle distributiegroepen auditen
.\Test-DistributionGroupPermissions.ps1

# Eén groep met ledenlijst
.\Test-DistributionGroupPermissions.ps1 -Group "helpdesk@contoso.com" -IncludeMembers
```

**Vereiste module**
```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

---

### Test-DkimConfig.ps1

Valideert de DKIM-ondertekeningsconfiguratie voor één of alle geaccepteerde domeinen:

- Controleert of DKIM-ondertekening is ingeschakeld
- Zoekt de CNAME-records `selector1/2._domainkey.<domain>` op en vergelijkt ze met de Exchange-configuratie
- Zoekt de TXT-records met de publieke sleutel van Microsoft op en controleert of de sleutel overeenkomt
- Somt de vereiste acties op voor gevonden problemen

> DNS-lookups gebruiken `Resolve-DnsName` (alleen Windows). Op macOS/Linux wordt de Exchange-configuratie getoond zonder DNS-validatie.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-Domain` | Nee | Te valideren domein. Zonder deze parameter worden alle domeinen met een ondertekeningsconfiguratie gecontroleerd |
| `-ShowAll` | Nee | Het volledige ondertekeningsconfiguratie-object tonen in plaats van de samengevatte weergave |
| `-TenantId` | Nee | Tenant-ID of domein. Standaard: de GDAP-klant (`authMode = 'GDAP'` in `load.config.ps1`), anders je eigen tenant |
| `-ClientId` / `-CertificateThumbprint` | Nee | App-only aanmelden met je eigen app (vereist `Exchange.ManageAsApp` en een Exchange-rol). Zonder: gedelegeerd, als jezelf |
| `-AppOnly` | Nee | App-only met de `ClientId` en `CertificateThumbprint` van de tenant uit `graph.appid.json` |

**Voorbeelden**

```powershell
# Alle domeinen controleren
.\Test-DkimConfig.ps1

# Eén domein
.\Test-DkimConfig.ps1 -Domain "contoso.com"
```

---

### Get-ExternalForwards.ps1

Controleert alle mailboxen op doorstuurregels die naar externe domeinen (buiten de tenant) wijzen. Externe doorsturing is een veelvoorkomend beveiligings- en complianceprobleem en moet regelmatig worden nagekeken. Exporteert naar CSV.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-Mailbox` | Nee | UPN van één mailbox. Zonder deze parameter worden alle mailboxen gecontroleerd |
| `-OutputPath` | Nee | Pad van het CSV-rapport (standaard: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Nee | Tenant-ID of domein. Standaard: de GDAP-klant (`authMode = 'GDAP'` in `load.config.ps1`), anders je eigen tenant |
| `-ClientId` / `-CertificateThumbprint` | Nee | App-only aanmelden met je eigen app (vereist `Exchange.ManageAsApp` en een Exchange-rol). Zonder: gedelegeerd, als jezelf |
| `-AppOnly` | Nee | App-only met de `ClientId` en `CertificateThumbprint` van de tenant uit `graph.appid.json` |

**Voorbeelden**

```powershell
# Alle mailboxen controleren
.\Get-ExternalForwards.ps1

# Eén mailbox
.\Get-ExternalForwards.ps1 -Mailbox "user@contoso.com"
```

---

### Get-MailboxSizes.ps1

Rapporteert mailboxgroottes (MB/GB), aantallen items en quotumstatus. Aflopend gesorteerd op grootte. Exporteert naar CSV.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-Mailbox` | Nee | UPN van één mailbox. Zonder deze parameter worden alle gebruikers- en gedeelde mailboxen gerapporteerd |
| `-OutputPath` | Nee | Pad van het CSV-rapport (standaard: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Nee | Tenant-ID of domein. Standaard: de GDAP-klant (`authMode = 'GDAP'` in `load.config.ps1`), anders je eigen tenant |
| `-ClientId` / `-CertificateThumbprint` | Nee | App-only aanmelden met je eigen app (vereist `Exchange.ManageAsApp` en een Exchange-rol). Zonder: gedelegeerd, als jezelf |
| `-AppOnly` | Nee | App-only met de `ClientId` en `CertificateThumbprint` van de tenant uit `graph.appid.json` |

**Voorbeelden**

```powershell
# Rapport van alle mailboxen
.\Get-MailboxSizes.ps1

# Eén mailbox
.\Get-MailboxSizes.ps1 -Mailbox "user@contoso.com"
```

---

### Get-DistributionGroupMembers.ps1

Exporteert elke distributielijst met haar leden naar één Excel-werkmap, bedoeld om ongewijzigd naar de klant te sturen.

De werkmap heeft twee bladen, allebei filterbare tabellen met een vastgezette koprij:

| Blad | Eén rij per | Kolommen |
|-------|-------------|---------|
| `Overzicht` | lijst | Lijst, E-mailadres, Type, Aantal leden, Eigenaar(s), Alias, Verborgen in adresboek, Alleen interne afzenders, Aangemaakt op |
| `Leden` | lid | Lijst, E-mailadres lijst, Type lijst, Lid, E-mailadres lid, Extern adres, Type lid |

Met `-Recurse` krijgen de bladen er `Aantal personen` en `Via groep` bij; met `-Member` krijgen ze er `Treffers` en `Treffer op` bij.

`Extern adres` wordt ingevuld voor e-mailcontacten en e-mailgebruikers. Hun primaire SMTP-adres is een interne placeholder — het adres dat de mail echt ontvangt is het externe, en voor een rapport over externe leden is dat de kolom die ertoe doet.

De koppen van de bladen en de ontvangertypen zijn Nederlands — `MailUniversalSecurityGroup` zegt niets tegen degene die het rapport leest, `Beveiligingsgroep (mail-enabled)` wel. Het script zelf blijft Engels, net als de rest van de repository.

**Geneste lijsten — lees dit voordat je op de uitvoer vertrouwt**

Exchange geeft altijd alleen **directe** leden terug. Een lijst die een andere lijst bevat, rapporteert die lijst als *één lid* en nooit de mensen daarin. Standaard geldt dus:

- iemand die alleen via een geneste groep mail ontvangt, verschijnt niet in het rapport;
- `-Member` meldt **geen treffers** op een lijst die in werkelijkheid wel bij die persoon aflevert.

Dat tweede punt is de gevaarlijke helft: een fout antwoord dat eruitziet als een zeker antwoord. `-Recurse` vouwt geneste groepen uit, zodat het rapport de mensen toont die de mail echt ontvangen:

```powershell
# "Komt er nog iets bij dat domein aan?" - de vorm om voor die vraag te gebruiken
.\Get-DistributionGroupMembers.ps1 -Member "@be.verizon.com" -Recurse
```

| | Zonder `-Recurse` | Met `-Recurse` |
|---|---|---|
| Geneste lijst | één ledenrij, niemand erachter | één ledenrij **plus** de mensen erin |
| Kolom `Via groep` | — | noemt de groep waarlangs een persoon binnenkwam, leeg voor een direct lid |
| `Aantal leden` | directe leden (wat Exchange en het EAC tonen) | ongewijzigd |
| `Aantal personen` | — | de werkelijke ontvangers die de lijst bereikt |

Het kost één extra query per geneste groep. Een groep die al is uitgevouwen, wordt niet opnieuw uitgevouwen; dat voorkomt ook dat een lidmaatschapscyclus (A bevat B, B bevat A) eindeloos doorloopt. Nesting dieper dan 20 niveaus wordt gemeld en met rust gelaten. Iemand die langs meerdere routes bereikbaar is, krijgt één rij met de routes samengevoegd, niet een rij per route.

De geneste groep zelf blijft als eigen rij in het rapport staan, zodat de structuur zichtbaar blijft.

**Filteren op een adres of een domein**

`-Member` accepteert één adres, één domein, of een domein met alles eronder:

```powershell
-Member "jan@contoso.com"     # op welke lijsten staat Jan?
-Member "@be.verizon.com"     # leden op precies dat domein
-Member "*.verizon.com"       # verizon.com EN elk subdomein ervan
```

| Geschreven als | Komt overeen met | Komt niet overeen met |
|---|---|---|
| `@be.verizon.com`, `be.verizon.com`, `*@be.verizon.com` | `jan@be.verizon.com` | `jan@verizon.com`, `jan@us.verizon.com`, `jan@notbe.verizon.com` |
| `*.verizon.com`, `.verizon.com`, `*@*.verizon.com` | `jan@verizon.com`, `jan@be.verizon.com`, `jan@us.verizon.com` | `jan@notverizon.com`, `jan@verizon.com.evil.test` |

Zonder de voorloop `*.` wordt er op dat **ene** domein gematcht — `@be.verizon.com` bereikt bewust geen zusterdomein als `@us.verizon.com`. Met de voorloop vallen het hoofddomein en elk subdomein binnen het bereik. De run toont welke van de twee hij doet (`...for members on verizon.com and its subdomains`), zodat het bereik nooit giswerk is.

Er wordt in beide gevallen op het volledige domeinlabel gematcht, en dat houdt `@notverizon.com` en de suffixtruc `@verizon.com.evil.test` buiten een `*.verizon.com`-run. Een jokerteken op een andere plek dan vooraan wordt niet ondersteund en wordt behandeld als letterlijk teken, in plaats van het filter stilletjes te verruimen.

De twee zijn niet even goedkoop. **Een adres** wordt omgezet naar zijn DN en door Exchange zelf gematcht (`Get-Recipient -Filter "Members -eq '<DN>'"`), dus niet elke groep in de tenant wordt doorlopen. **Een domein** kan dat niet: er is geen serverfilter voor *"heeft een lid waarvan het adres eindigt op @x"*, dus elke lijst wordt gelezen en daarna gefilterd. In een grote tenant is dat één `Get-DistributionGroupMember`-aanroep per lijst — trager, en goed om te weten voordat je het op duizenden groepen loslaat.

Er wordt gematcht op het primaire adres, **elke alias** en — voor e-mailcontacten en e-mailgebruikers — `ExternalEmailAddress`. Om dat laatste gaat het: een Verizon-contact in een distributielijst is doorgaans een e-mailcontact met een primair SMTP-adres als `marc.dubois@contoso.onmicrosoft.com`, met `@be.verizon.com` alleen in het externe adres. Alleen op het primaire adres matchen zou niets vinden.

Met een actief filter krijgen beide bladen een kolom erbij:

| Kolom | Blad | Betekenis |
|--------|-------|---------|
| `Treffers` | `Overzicht` | hoeveel leden van deze lijst overeenkwamen |
| `Treffer op` | `Leden` | het adres waarop dit lid overeenkwam, leeg als het niet overeenkwam |

`Treffer op` bevat het **adres**, geen Ja/Nee — iemand kan overeenkomen op een alias die nergens anders in het rapport voorkomt, en `Ja` naast `sara@contoso.com` roept alleen de vraag op waarom. `sara.willems@be.verizon.com` beantwoordt die.

Lijsten met een treffer worden **volledig** geëxporteerd, zodat de klant ziet wie er nog meer op staat; sorteer of filter op `Treffer op` om alleen de treffers te zien.

Alleen direct lidmaatschap — iemand in een geneste groep is geen treffer. De geneste groep zelf verschijnt wel als ledenrij, met `Distributielijst` als lidtype.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-Group` | Nee | Eén lijst (naam, alias of e-mailadres). Zonder deze parameter wordt elke lijst gerapporteerd |
| `-Member` | Nee | Alleen de lijsten met dit adres (`jan@contoso.com`), dit domein (`@be.verizon.com`), of dit domein en de subdomeinen ervan (`*.verizon.com`) |
| `-Recurse` | Nee | Geneste groepen uitvouwen, zodat ook de mensen achter een geneste lijst worden gerapporteerd |
| `-IncludeDynamic` | Nee | Ook dynamische distributiegroepen rapporteren (live geëvalueerd, één query per groep) |
| `-IncludeM365Groups` | Nee | Ook Microsoft 365-groepen rapporteren, inclusief groepen achter een Team |
| `-OutputPath` | Nee | Pad van de `.xlsx` (standaard: `C:\Temp\Distributielijsten_<timestamp>.xlsx`) |
| `-Csv` | Nee | Twee CSV-bestanden schrijven in plaats van Excel |
| `-TenantId` | Nee | Tenant-ID of domein. Standaard: de GDAP-klant (`authMode = 'GDAP'` in `load.config.ps1`), anders je eigen tenant |
| `-ClientId` / `-CertificateThumbprint` | Nee | App-only aanmelden met je eigen app (vereist `Exchange.ManageAsApp` en een Exchange-rol). Zonder: gedelegeerd, als jezelf |
| `-AppOnly` | Nee | App-only met de `ClientId` en `CertificateThumbprint` van de tenant uit `graph.appid.json` |

**Voorbeelden**

```powershell
# Elke distributielijst met al haar leden
.\Get-DistributionGroupMembers.ps1

# Op welke lijsten staat Jan? (en wie staan er nog meer op)
.\Get-DistributionGroupMembers.ps1 -Member "jan@contoso.com"

# Welke lijsten bevatten nog adressen op een partnerdomein?
.\Get-DistributionGroupMembers.ps1 -Member "@be.verizon.com"

# Hetzelfde, maar ook mensen vinden die in een geneste lijst zitten
.\Get-DistributionGroupMembers.ps1 -Member "@be.verizon.com" -Recurse

# Alles van Verizon: het hoofddomein en elk subdomein, geneste lijsten uitgevouwen
.\Get-DistributionGroupMembers.ps1 -Member "*.verizon.com" -Recurse

# Eén lijst, naar een vast pad
.\Get-DistributionGroupMembers.ps1 -Group "helpdesk@contoso.com" -OutputPath "C:\Reports\helpdesk.xlsx"

# Alles wat als groep mail kan ontvangen
.\Get-DistributionGroupMembers.ps1 -IncludeDynamic -IncludeM365Groups
```

**Opmerkingen**
- Vereist [ImportExcel](https://github.com/dfinke/ImportExcel) voor de `.xlsx`. Als de module ontbreekt, meldt het script dat, verwijst het naar `scripts\Startup\Install-Modules.ps1` en schrijft het in plaats daarvan twee CSV-bestanden (`*-overzicht.csv`, `*-leden.csv`) — een ontbrekende module kost je nooit het rapport. Het script installeert zelf geen modules meer
- Blijft op Exchange Online: Graph's `transitiveMembers` dekt distributielijsten en mail-enabled beveiligingsgroepen, maar geen dynamische distributiegroepen, `ManagedBy`-eigenaars of e-mailcontactpersonen zoals dit rapport ze toont
- Een lijst zonder leden krijgt een rij `(geen leden)` in het blad `Leden` in plaats van er stilletjes in te ontbreken — een lege lijst is precies wat een klant wil opmerken
- Lidmaatschap wordt per groep gelezen, dus iemand die op geen enkele lijst staat, verschijnt nergens — het rapport gaat over groepslidmaatschap, niet over de gebruikersdirectory
- Een domeinfilter zonder treffers meldt *"No distribution list has a member on @x"* en schrijft geen bestand — een lege werkmap leest als een mislukt rapport in plaats van als het antwoord dat het is
- CSV-uitvoer gebruikt `-UseCulture`, zodat een Nederlandstalige Excel het als kolommen opent in plaats van als één muur kommagescheiden tekst
- Een bestaande werkmap op `-OutputPath` wordt vervangen, niet aangevuld — `Export-Excel` zou anders een tweede run boven op de eerste stapelen

---

### Get-MessageTraceReport.ps1

Beantwoordt "wie heeft dit ontvangen, wanneer precies, en waar ging het daarna heen?". Voert een message trace uit en rapporteert per bericht het exacte tijdstip in **zowel lokale tijd als UTC** (Exchange slaat de tijdstempels van message traces op in UTC), afzender, ontvanger, onderwerp, status, grootte, het verzendende/afleverende IP-adres, `MessageId` en `MessageTraceId`.

**Detectie van doorsturing** — de kolom `ForwardedTo` wordt gevuld vanuit drie onafhankelijke signalen, en `ForwardDetection` noemt welk signaal afging:

| Methode | Detecteert |
|--------|---------|
| `SameMessageId` | Andere ontvangers die dezelfde `MessageId` ontvingen — SMTP-doorsturing, omleidingsregels, uitvouwen van distributiegroepen |
| `RedirectHop` | Omleidings-/transportregelhops uit `Get-MessageTraceDetailV2` (vereist `-IncludeDetails`) |
| `ClientForward(subject match)` | Een later bericht dat **door** de ontvanger is verzonden met hetzelfde genormaliseerde onderwerp — een "Doorsturen" in Outlook, dat een gloednieuwe `MessageId` krijgt. Heuristisch; antwoorden terug naar de oorspronkelijke afzender worden uitgesloten |

> **Waarom het doorstuurdoel apart wordt getraceerd:** een mailboxdoorsturing of een omleidingsregel houdt de **oorspronkelijke afzender** op de doorgestuurde kopie. De trace-rij voor de aflevering bij het doorstuurdoel noemt de getraceerde mailbox dus *niet als afzender en niet als ontvanger* — filteren op alleen de mailbox zou die rij nooit opleveren. Het script bepaalt de ingestelde doorstuurdoelen **vóór** het traceren en voegt ze toe als extra ontvangerfilters, zodat de eigenlijke overdracht met een eigen exact tijdstip zichtbaar wordt. `-ResolveSiblings` gaat verder en traceert elke gevonden `MessageId` opnieuw zonder filter, waarmee ook doorstuurdoelen worden gevonden die niet meer zijn ingesteld (een regel die is verwijderd nadat hij zijn werk had gedaan, laat zijn afleveringen nog steeds in de trace achter).

Daarbovenop rapporteert het script de **ingestelde** doorsturing van elke interne mailbox die in de trace voorkomt — `ForwardingSMTPAddress` / `ForwardingAddress` plus elke inboxregel met `ForwardTo` / `RedirectTo` / `ForwardAsAttachmentTo` — zodat een doorsturing die binnen het getraceerde venster niet is afgegaan, toch zichtbaar is.

Gebruikt alleen `Get-MessageTraceV2` en `Get-MessageTraceDetailV2` — de oude `Get-MessageTrace` / `Get-MessageTraceDetail` zijn uitgefaseerd, en het script stopt met een hint als de V2-cmdlets ontbreken. Bereiken die langer zijn dan de V2-limiet worden automatisch in blokken van 10 dagen opgesplitst, en elk blok wordt met `-StartingRecipientAddress` gepagineerd tot het uitgeput is. Exchange Online PowerShell omdat Graph geen message-trace-API heeft; de doorstuurconfiguratie wordt daar ook gelezen.

**Parameters**

| Parameter | Verplicht | Standaard | Omschrijving |
|-----------|----------|---------|-------------|
| `-Mailbox` | Nee | — | Beide richtingen traceren voor dit adres (verzonden **en** ontvangen) en de doorstuurconfiguratie ophalen |
| `-SenderAddress` | Nee | — | Filteren op afzenderadres. Alias `-Sender` (`$Sender` is een automatische variabele van PowerShell en kan dus niet de echte naam van de parameter zijn) |
| `-Recipient` | Nee | — | Filteren op ontvangeradres |
| `-ForwardAddress` | Nee | — | Bekende doorstuur-/exfiltratieadres(sen) om als ontvangers te traceren, bovenop de doorstuurconfiguratie die wordt gevonden. Gebruik dit als de doorsturing **al is verwijderd** — er is dan geen configuratie meer te vinden, maar de eerdere afleveringen staan nog in de trace |
| `-Subject` | Nee | — | Onderwerpfilter aan clientzijde, jokertekens toegestaan (message trace kan aan serverzijde niet op onderwerp filteren) |
| `-MessageId` | Nee | — | Te traceren internet-MessageId, met of zonder punthaken |
| `-Days` | Nee | `2` | Aantal dagen terug vanaf `-EndDate`. Genegeerd als `-StartDate` is opgegeven |
| `-StartDate` | Nee | — | Expliciet begin van het venster (lokale tijd) |
| `-EndDate` | Nee | nu | Expliciet einde van het venster (lokale tijd) |
| `-Status` | Nee | — | `Delivered`, `Failed`, `Pending`, `Expanded`, `Quarantined`, `FilteredAsSpam`, `GettingStatus`, `None` |
| `-IncludeDetails` | Nee | uit | Afleverdetails per hop ophalen — dit is wat omleidings-/transportregeldoelen zichtbaar maakt. Traag en onderhevig aan throttling |
| `-MaxDetailLookups` | Nee | `50` | Maximum aantal hopdetail-opvragingen; afkapping wordt expliciet gemeld |
| `-ResolveSiblings` | Nee | uit | Elke gevonden `MessageId` opnieuw traceren zonder afzender-/ontvangerfilter om **alle** ontvangers te tonen — vindt doorstuurdoelen die niet meer zijn ingesteld. Eén extra aanroep per MessageId |
| `-MaxSiblingLookups` | Nee | `100` | Maximum aantal sibling-opvragingen |
| `-SkipForwardingConfig` | Nee | uit | De controle van mailboxdoorsturing / inboxregels overslaan |
| `-OutputPath` | Nee | `C:\Temp\` / `~/Downloads` | Pad van de hoofd-CSV. Detail- en doorstuurrapporten worden ernaast geschreven met de achtervoegsels `_Details` / `_ForwardingConfig` |
| `-TenantId` | Nee | GDAP-klant / eigen tenant | Tenant-ID of domein |
| `-ClientId` / `-CertificateThumbprint` | Nee | — | App-only aanmelden met je eigen app. Zonder: gedelegeerd, als jezelf |
| `-AppOnly` | Nee | uit | App-only met de `ClientId` en `CertificateThumbprint` van de tenant uit `graph.appid.json` |

**Voorbeelden**

```powershell
# Alles wat één mailbox de afgelopen 2 dagen verzond en ontving, incl. doorsturingen
.\Get-MessageTraceReport.ps1 -Mailbox "user@contoso.com"

# Eén specifieke stroom over de afgelopen 30 dagen, met details per hop
.\Get-MessageTraceReport.ps1 -Sender "boss@contoso.com" -Recipient "user@contoso.com" -Days 30 -IncludeDetails

# Waar is dit specifieke bericht terechtgekomen?
.\Get-MessageTraceReport.ps1 -MessageId "<abc123@contoso.com>" -Days 10 -IncludeDetails

# Vermoeden van externe doorsturing — volledig beeld, incl. doelen die niet meer zijn ingesteld
.\Get-MessageTraceReport.ps1 -Mailbox "facturen@contoso.com" -Days 10 -ResolveSiblings -IncludeDetails

# De doorsturing is al verwijderd, maar het adres is bekend — toch traceren
.\Get-MessageTraceReport.ps1 -Mailbox "facturen@contoso.com" -Days 10 -ResolveSiblings `
    -ForwardAddress "exfil@lookalike-domain.nl"

# Alle mislukte mail van één afzender in een expliciet venster
.\Get-MessageTraceReport.ps1 -Sender "noreply@contoso.com" -Status Failed `
    -StartDate (Get-Date "2026-08-01") -EndDate (Get-Date "2026-08-08")
```

**Opmerkingen**
- Message trace bewaart **90 dagen**; het script waarschuwt als het gevraagde venster daar voorbij reikt
- Voor het lezen van inboxregels zijn rechten op de mailbox nodig — mailboxen die niet gelezen kunnen worden, worden stil overgeslagen (gebruik `-Verbose` om te zien welke)
- `-IncludeDetails` doet één API-aanroep per bericht en is onderhevig aan throttling van Exchange Online; verhoog `-MaxDetailLookups` bewust

**Vereiste module:** `ExchangeOnlineManagement` — geïnstalleerd door `scripts\Startup\Install-Modules.ps1`.

---

### Remove-PhishingMessage.ps1

Incident-response-tegenhanger van `Get-MessageTraceReport.ps1`: de trace vertelt je **wie** de phish **heeft ontvangen**, dit script **haalt hem weer weg**. Standaard een proefdraai — er wordt niets verwijderd zonder `-Apply`.

**Twee engines**

| Engine | Hoe berichten worden gevonden | Gebruik het als |
|--------|----------------------|-------------|
| `Purview` | Eén KQL Content Search over de hele tenant, daarna `New-ComplianceSearchAction -Purge` | Je de ontvangers **niet** kent, of een **HardDelete** nodig hebt. Rapporteert aantallen **per mailbox**, niet afzonderlijke berichten |
| `Graph` | Doorloopt elke doelmailbox via de Graph-mail-API en verwijdert bericht voor bericht | Je de ontvangers **wel** kent (uit de trace) en het **nu** weg wilt hebben, met een rapport per bericht |

De engine is standaard `Graph` als `-Mailbox` is opgegeven en anders `Purview`. Overschrijf dat met `-Engine`.

> **Waarom twee engines.** Purview leest de **zoekindex**, die zo'n 15–30 minuten achterloopt op de aflevering — een purge die direct na het binnenkomen van de phish wordt uitgevoerd, kan eerlijk *0 hits* melden en het bericht toch in elk Postvak IN laten staan. Graph bevraagt de mailbox rechtstreeks en heeft die vertraging niet, maar heeft de lijst met ontvangers nodig en kan niet naar `Recoverable Items\Purges` schrijven, dus kan niet hard verwijderen. Tijdens een lopende campagne is de gebruikelijke volgorde: trace → **Graph** direct op de bekende ontvangers → **Purview**-sweep over de hele tenant een halfuur later om de rest te vangen.

**Verwijdertypen**

| Waarde | Komt terecht in | Kan de gebruiker het terughalen? | Engines |
|-------|----------|-------------------|---------|
| `Recycle` | Verwijderde items | Ja, eenvoudig | Graph |
| `SoftDelete` *(standaard)* | `Recoverable Items\Deletions` | Ja, via "Verwijderde items herstellen" | Beide |
| `HardDelete` | `Recoverable Items\Purges` | Nee — alleen bewaard als de mailbox onder een hold staat | Purview |

**Parameters**

| Parameter | Verplicht | Standaard | Omschrijving |
|-----------|----------|---------|-------------|
| `-Mailbox` | Nee | — | Doelmailboxadres(sen). Vereist voor Graph, tenzij `-AllMailboxes`. Weggelaten bij Purview = elke mailbox in de tenant |
| `-AllMailboxes` | Nee | uit | Elke mailbox doorzoeken. Impliciet bij Purview; bij Graph is dit één query per mailbox en traag |
| `-MessageId` | Nee | — | Internet-MessageId, met of zonder punthaken. **De precieze selector** — matcht dat ene bericht en niets anders |
| `-SenderAddress` | Nee | — | Afzenderadres. Alias `-Sender` (`$Sender` is een automatische variabele van PowerShell) |
| `-Subject` | Nee | — | Purview matcht dit als geïndexeerde woordgroep; Graph matcht aan clientzijde en accepteert jokertekens |
| `-AttachmentName` | Nee | — | Bestandsnaam van de bijlage, jokertekens toegestaan (bijv. `*.html`) |
| `-BodyContains` | Nee | — | Woord of woordgroep in de inhoud. **Alleen Purview** — Graph zou elke berichtinhoud moeten downloaden |
| `-ReceivedAfter` | Nee | — | Alleen berichten ontvangen op of na dit moment (lokale tijd) |
| `-ReceivedBefore` | Nee | — | Alleen berichten ontvangen op of vóór dit moment (lokale tijd) |
| `-Engine` | Nee | zie hierboven | `Purview` of `Graph` |
| `-DeleteType` | Nee | `SoftDelete` | `Recycle`, `SoftDelete` of `HardDelete` (zie de tabel hierboven) |
| `-Apply` | Nee | uit | **Echt verwijderen.** Zonder deze switch meldt het script alleen wat het gevonden heeft |
| `-SearchName` | Nee | `Phish_<timestamp>` | Naam van de aan te maken Content Search. Purview vereist unieke namen |
| `-KeepSearch` | Nee | uit | De Content Search achteraf bewaren, zodat je hem in de Purview-portal kunt bekijken |
| `-IncludeCalendar` | Nee | uit | Ook overeenkomende **agenda-items** verwijderen, niet alleen mail. Werkt met **beide engines**; vereist `-Subject` of `-SenderAddress` |
| `-CalendarDaysBack` | Nee | `30` | Hoe ver terug de agenda wordt doorzocht |
| `-CalendarDaysForward` | Nee | `365` | Hoe ver vooruit de agenda wordt doorzocht |
| `-VerifyWithGraph` | Nee | uit | Na een Purview-purge de betrokken mailboxen via Graph controleren om te bevestigen dat de berichten echt weg zijn. Vereist dezelfde Graph-toegang als `-Engine Graph` |
| `-MaxPurgeRounds` | Nee | `10` | Purview purget maximaal 10 items per mailbox per actie, dus het script werkt in rondes. 10 rondes = tot 100 items per mailbox |
| `-MaxMessagesPerMailbox` | Nee | `500` | Veiligheidslimiet per mailbox voor Graph; het bereiken ervan wordt expliciet gemeld |
| `-TimeoutMinutes` | Nee | `30` | Hoe lang er wordt gewacht tot een zoek- of purge-actie klaar is |
| `-OutputPath` | Nee | `C:\Temp\` / `~/Downloads` | Pad van het CSV-rapport |
| `-TenantId` | Nee | GDAP-klant | Tenant-ID of domein, gebruikt als het script zelf verbinding moet maken |
| `-ClientId` | Nee | — | Je eigen App Registration voor app-only Graph-authenticatie — slaat de automatische tijdelijke app over |
| `-ClientSecret` | Nee | — | Client secret voor `-ClientId` |
| `-CertificateThumbprint` | Nee | — | Certificaatvingerafdruk voor `-ClientId` |
| `-AppOnly` | Nee | uit | Je eigen app met `ClientId` en `CertificateThumbprint` uit `graph.appid.json` |
| `-Delegated` | Nee | uit | Graph-engine als jezelf (`Mail.ReadWrite.Shared`, plus `Calendars.ReadWrite.Shared` met `-IncludeCalendar`) — helemaal geen app. Bereikt **alleen mailboxen waarop je al Full Access hebt**, dus geschikt voor een paar bekende ontvangers, niet voor `-AllMailboxes`. Niet onder GDAP |

**Voorbeelden**

```powershell
# 1. Wat zou er in de hele tenant worden verwijderd? (geen -Apply = er wordt niets verwijderd)
.\Remove-PhishingMessage.ps1 -MessageId "<abc123@evil.example>"

# 1b. Purgen en daarna via Graph bevestigen dat het echt weg is
.\Remove-PhishingMessage.ps1 -MessageId "<abc123@evil.example>" `
    -DeleteType HardDelete -Apply -VerifyWithGraph

# 2. Hetzelfde, nu echt purgen zodat de gebruiker het niet kan terughalen
.\Remove-PhishingMessage.ps1 -MessageId "<abc123@evil.example>" `
    -DeleteType HardDelete -Apply

# 3. Bekende ontvangers uit de message trace — direct, zonder indexvertraging
.\Remove-PhishingMessage.ps1 `
    -Mailbox "a@contoso.com","b@contoso.com" `
    -Sender  "no-reply@evil.example" `
    -Subject "*password expires*" `
    -Apply

# 4. Campagnesweep: alles van één afzender in een venster, over de hele tenant
.\Remove-PhishingMessage.ps1 -Sender "no-reply@evil.example" `
    -ReceivedAfter (Get-Date "2026-08-30") -DeleteType HardDelete -Apply

# 5. Phishing-VERGADERUITNODIGING, hele tenant: de mail EN de afspraken hard verwijderen
.\Remove-PhishingMessage.ps1 -Sender "no-reply@evil.example" -Subject "kick-off" `
    -DeleteType HardDelete -IncludeCalendar -Apply -VerifyWithGraph

# 6. Campagne met HTML-bijlage
.\Remove-PhishingMessage.ps1 -AttachmentName "*.html" `
    -Sender "billing@evil.example" -Apply

# 7. Twee bekende ontvangers waarop je Full Access hebt — Graph als jezelf, geen app
.\Remove-PhishingMessage.ps1 -Mailbox "a@contoso.com","b@contoso.com" `
    -Sender "no-reply@evil.example" -Delegated -Apply
```

**Typisch verloop van een incident**

```powershell
# Wie heeft het ontvangen, en waar ging het heen?
.\Get-MessageTraceReport.ps1 -Sender "no-reply@evil.example" -Days 2 -ResolveSiblings

# Het nu direct bij de bekende ontvangers weghalen (geen indexvertraging)
.\Remove-PhishingMessage.ps1 -Mailbox $recipients -MessageId "<abc123@evil.example>" -Apply

# ~30 minuten later de tenant doorzoeken op alles wat de trace heeft gemist
.\Remove-PhishingMessage.ps1 -MessageId "<abc123@evil.example>" -DeleteType HardDelete -Apply
```

**Vereiste rechten**

| Engine | Recht |
|--------|-----------|
| `Purview` | Lidmaatschap van de rol **Search And Purge** — in de praktijk de rolgroep *Organization Management* of *eDiscovery Manager* in de Purview-complianceportal. Maakt gedelegeerd verbinding via `Connect-IPPSSession -EnableSearchOnlySession`, en bereikt een GDAP-klant met `-DelegatedOrganization` (een eigen aanroep van het script, omdat `Connect-M365Exchange -IncludeCompliance` `-EnableSearchOnlySession` niet doorgeeft) |
| `Graph` | Standaard app-only `Mail.ReadWrite` — **dat hoef je niet zelf te regelen**, zie de drie routes hieronder. Of `-Delegated` met Full Access op de doelmailboxen |

**Hoe de Graph-engine (en `-VerifyWithGraph`) aan toegang komt**

Hetzelfde drieledige patroon als [`Move-InboxToArchive.ps1`](Move-InboxToArchive.ps1) ([docs](#move-inboxtoarchiveps1)) en de SharePoint-rapportagescripts, in deze volgorde geprobeerd:

| # | Route | Wat er nodig is |
|---|-------|---------------|
| 1 | Een app-only Graph-sessie die je al had opgezet | Niets — wordt gebruikt zoals ze is |
| 2 | `-ClientId` + `-TenantId` + (`-ClientSecret` of `-CertificateThumbprint`), of `-AppOnly` | Je eigen app met de applicatiemachtiging `Mail.ReadWrite`, met admin consent. **Met `-ClientSecret` is dit de robuustste route** — die haalt haar token via gewone REST en laadt nooit de Graph SDK |
| 3 | **Automatisch** — aanmelding met apparaatcode, daarna een kortlevende App Registration die zichzelf `Mail.ReadWrite` toekent, een app-only-token afgeeft en **weer wordt verwijderd als de run klaar is** | Global Administrator of Privileged Role Administrator voor die eenmalige aanmelding. Geen extra modules |

Route 3 is wat er gebeurt als je niets meegeeft, dus `-VerifyWithGraph` werkt direct. De gedelegeerde rol verleent de consent, dus er is geen apart admin-consentscherm. Als de setup halverwege mislukt, wordt de half aangemaakte app verwijderd voordat de fout wordt gemeld — er blijven geen wezen achter in Entra ID.

Routes 2 (met `-ClientSecret`) en 3 zijn allebei gebouwd op gewone REST — de device code flow voor de aanmelding, de Graph REST API voor het aanmaken en verwijderen van de App Registration. **Geen van beide laadt de Graph SDK**, en daardoor werken ze in dezelfde sessie die al met Exchange verbonden is. Route 3 toont een code om in te voeren op `microsoft.com/devicelogin`:

```
  ------------------------------------------------------------
   To sign in, use a web browser to open https://microsoft.com/devicelogin
   and enter the code ABCD-EFGH to authenticate.
  ------------------------------------------------------------
```

> **Exchange en Graph vechten om MSAL.** `ExchangeOnlineManagement` en `Microsoft.Graph.Authentication` bundelen elk hun eigen `Microsoft.Identity.Client`, en .NET laadt alleen de eerste die een proces aanraakt. Een Purview-purge (die met Exchange verbindt) gevolgd door `-VerifyWithGraph` in hetzelfde venster laat de Graph SDK dus aanroepen doen naar een MSAL waarvan de API niet overeenkomt, en dat mislukt met `Method not found: ... WithLogging(...)` — wat in niets lijkt op het versieconflict dat het is.
>
> Routes 2 (`-ClientSecret`) en 3 omzeilen dit volledig door de SDK nooit te laden. Alleen de variant met `-CertificateThumbprint` en het hergebruiken van een bestaande `Connect-MgGraph`-sessie gaan er nog doorheen, en beide melden het conflict voor wat het is, in plaats van je de stacktrace te laten ontcijferen.
>
> Als `-VerifyWithGraph` met de Purview-engine wordt gebruikt, wordt de Graph-toegang **vóór** de purge opgezet, zodat een verificatie die niet kan draaien vooraf wordt gemeld in plaats van nadat de berichten weg zijn. De purge wordt hoe dan ook uitgevoerd — een mislukte verificatie betekent nooit een mislukte purge.

> Gedelegeerde `Mail.ReadWrite` bereikt alleen ooit *je eigen* mailbox, dus een bestaande gedelegeerde `Connect-MgGraph`-sessie wordt bewust **niet** geaccepteerd voor de Graph-engine; het script valt in plaats daarvan door naar route 2 of 3. Daarom is app-only de standaard. `-Delegated` is het expliciete alternatief: een aanmelding met apparaatcode via dezelfde gewone REST met `Mail.ReadWrite.Shared`, ververst voor lange runs, die alleen werkt op mailboxen waarop je al Full Access hebt. Let op: `Mail.ReadWrite` (applicatie) geeft toegang tot **elke** mailbox in de tenant; beperk de app met `New-ApplicationAccessPolicy` als dat ruimer is dan je wilt.

> GDAP-bewust: onder een GDAP-sessie (`$global:authMode -eq 'GDAP'`, ingesteld door `Connect-Tenant` / `load.ps1`) wordt `-TenantId` afgeleid uit de geselecteerde klanttenant, net als bij de SharePoint-scripts.

**Opmerkingen**
- De agenda vereist `Calendars.ReadWrite`, wat niet onder `Mail.ReadWrite` valt. De automatische tijdelijke app kent die machtiging **alleen toe als `-IncludeCalendar` wordt gebruikt**, en het script controleert de roles-claim in het uitgegeven token voordat het iets doet — een app-only-token wordt uitgegeven, of de toekenning nu is doorgevoerd of niet, en een te vroeg aangemaakt token wordt een uur gecachet, wat anders uitdraait op een 403 op elke afzonderlijke mailbox voor de hele run
- **Een phishing-vergaderuitnodiging is pas half weg als de mail is verwijderd.** De uitnodiging laat een afspraak in de agenda achter, en `/messages` en `/events` zijn aparte verzamelingen — geen van beide engines raakt standaard de agenda aan. `-IncludeCalendar` neemt die ook mee, met een match op `-Subject` of `-SenderAddress` (als organisator). Zonder selector weigert het, in plaats van de hele agenda te doorlopen
- **`-IncludeCalendar` werkt ook met de Purview-engine**, en die combinatie is de volledige opruiming van een phish via een vergaderuitnodiging: Purview verwijdert de uitnodiging hard in de hele tenant, daarna verwijdert een Graph-ronde de afspraken die zijn achtergebleven in precies de mailboxen die de zoekopdracht vond. De agendaronde heeft dezelfde Graph-toegang nodig als `-VerifyWithGraph`, en een mailbox die ze niet kan bereiken, wordt gemeld in plaats van als schoon geteld
- De verificatie volgt daarin: met `-IncludeCalendar` controleert ze ook de agenda, en zonder toont ze *"mail only — calendar items are not checked"* in plaats van een mailbox op basis van onvolledig bewijs schoon te verklaren
- **Niets in Purview kan een purge bevestigen.** De purge-actie meldt wat de dienst denkt te hebben gedaan, en de zoekindex blijft gepurgede items tot ~30 minuten lang tonen — het script opnieuw uitvoeren is dus geen controle. `-VerifyWithGraph` is de enige verificatie zonder vertraging: die stelt *dezelfde* query waarop de Graph-engine verwijdert opnieuw, rechtstreeks op de mailboxen die de zoekopdracht vond. Soft- en hard verwijderde items staan in Recoverable Items, die Graph niet toont, dus een gepurged bericht leest terecht als weg
- De verificatie onderscheidt **"kon niet controleren"** van **"schoon"**. Een mailbox die 403 teruggeeft, wordt gemeld als niet-geverifieerd, nooit als bevestigd. Ze laat de run ook nooit mislukken — een purge die al heeft plaatsgevonden, wordt niet als mislukt gemeld omdat de controle niet kon draaien
- **Purview kan je de afzonderlijke berichten niet laten zien.** Content Search rapporteert aantallen items per mailbox; de preview-actie die vroeger afzender en onderwerp per bericht teruggaf, is sinds de eDiscovery-wijzigingen van mei 2025 [gedocumenteerd als alleen on-premises](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/new-compliancesearchaction?view=exchange-ps). Voor details per bericht neem je de lijst met mailboxen uit de Purview-run en voer je die adressen opnieuw uit via `-Engine Graph`
- Purview voert zoekopdrachten en purges aan serverzijde uit, en die duren geregeld minuten. Het script meldt tijdens het wachten ongeveer elke 15 seconden de jobstatus en de verstreken tijd, zodat een trage stap zichtbaar traag is in plaats van vastgelopen te lijken, en `-TimeoutMinutes` (standaard 30) begrenst het
- **Rondes worden gepland, niet gepolld.** Een purge verwijdert maximaal 10 items per mailbox per actie, dus het script berekent `ceil(max items per mailbox / 10)` uit de eerste zoekopdracht. Het blijft bewust *niet* herhalen tot de index stil wordt: de index loopt tot ~30 minuten achter op een purge, dus dat zou dezelfde items opnieuw purgen en daarna een onterechte afkapping melden. Wat elke ronde echt heeft verwijderd, wordt uit de purge-actie zelf teruggelezen
- Eén content search purget maximaal **50.000 mailboxen**; daarboven waarschuwt het script en moet je in batches werken met `-Mailbox`. Microsoft verwijst voor bulkwerk naar de Graph-API `ediscoverySearch: purgeData` (100 items per locatie)
- **Vereist ExchangeOnlineManagement 3.9.0+** voor de Purview-engine. Content Search draait op een backend die een gewone IPPS-verbinding niet meer bereikt: zonder `-EnableSearchOnlySession` zijn de cmdlets aanwezig, maar mislukt `Start-ComplianceSearch` bij de initialisatie. Het script geeft de switch mee als het zelf verbinding maakt. **Als je al zonder die switch verbonden was, is de sessie vanuit het proces niet te herstellen** — open een nieuw PowerShell-venster en laat het script de verbinding maken
- **Je hebt niet het hele onderwerp nodig.** `-Subject` matcht een fragment: bij Purview wordt het een KQL-woordgroep, die overal in het onderwerp matcht, dus `-Subject "kick-off meeting"` vindt elke mail die die woorden in die volgorde bevat. Een lang onderwerp vol leestekens is juist de *kwetsbare* keuze — komma's, apostroffen en tijden als `14:09` worden slecht in woorden opgesplitst. Kort en onderscheidend wint
- **Matchen binnen een woord** is het enige wat KQL niet kan. Een `*` vooraan wordt weggelaten (het script waarschuwt in plaats van het stilletjes te negeren) en een `*` achteraan blijft alleen werken bij één woord, omdat een jokerteken binnen een woordgroep tussen aanhalingstekens geen effect heeft. Voor echte substring-matching gebruik je `-Engine Graph` met `-Mailbox`, dat aan clientzijde matcht en jokertekens neemt zoals ze geschreven zijn
- **Selectors worden gecombineerd met AND.** Het gebruikelijke phishingpaar is de afzender plus een onderwerpfragment: `-Sender "no-reply@evil.example" -Subject "kick-off meeting"`. Als de afzender van adres wisselt, is `-BodyContains "a distinctive sentence"` (alleen Purview) vaak de duurzaamste selector
- Een run die halverwege mislukt, laat zijn Content Search niet meer achter — het opruimen gebeurt in een `finally`. Zoekopdrachten die door oudere runs zijn achtergelaten, kun je opvragen met `Get-ComplianceSearch | Where-Object Name -like 'Phish_*'` en verwijderen met `Remove-ComplianceSearch`
- Het script **weigert te draaien** zonder minstens één van `-MessageId`, `-SenderAddress`, `-Subject`, `-AttachmentName` of `-BodyContains` — een datumbereik op zich zou elk bericht in elke mailbox matchen
- **Indexvertraging** (alleen Purview): een bericht dat in de laatste ~30 minuten is afgeleverd, is mogelijk nog niet doorzoekbaar. Een resultaat van `0 hits` direct na de aflevering bewijst niet dat de phish weg is — wacht en voer het opnieuw uit, of gebruik de Graph-engine
- Purge dekt **alleen de primaire mailbox** — geen van beide engines bereikt de archiefmailbox
- KQL ondersteunt geen jokertekens binnen een woordgroep, dus bij `-Subject "*invoice*"` worden de jokertekens bij de Purview-engine weggehaald en wordt als woordgroep gematcht; bij Graph werken de jokertekens zoals geschreven
- Elke run schrijft een CSV-rapport van wat er is gevonden en wat er is verwijderd

**Vereiste modules:** `ExchangeOnlineManagement`, en optioneel `Microsoft.Graph.Authentication` — beide geïnstalleerd door `scripts\Startup\Install-Modules.ps1`.

Microsoft.Graph.Authentication is alleen nodig om een bestaande `Connect-MgGraph`-sessie te hergebruiken of om `-CertificateThumbprint` te gebruiken. De routes met `-ClientSecret` en met de automatische tijdelijke app draaien op gewone REST en hebben niets nodig naast ExchangeOnlineManagement.

---

### Restore-MailboxMessages.ps1

Draait een slechte dag in één mailbox terug: berichten die op een bepaalde datum zijn **verplaatst** of **verwijderd**, gaan terug naar de map waar ze vandaan kwamen, en de run meldt **wie het deed** — account, als eigenaar / gemachtigde / beheerder, met welke client, vanaf welk IP-adres. Standaard een voorbeeld — er wordt niets verplaatst zonder `-Apply`.

**Drie bronnen, omdat geen enkele elk geval dekt**

| Onderdeel | Waar het kijkt | Hoe het de oorspronkelijke map kent | Wie het deed |
|------|----------------|----------------------------------|------------|
| Auditlog | `Search-UnifiedAuditLog` — `Move`, `MoveToDeletedItems`, `SoftDelete`, `HardDelete` op deze mailbox | Legt de bronmap van elke verplaatsing vast | **Ja** — account, aanmeldingstype, client, IP, app-ID |
| `Deleted` | `Get-/Restore-RecoverableItems` over Verwijderde items, Recoverable Items en Purges (alleen bewaard onder een hold), gefilterd op het moment van verwijderen | Exchange houdt het zelf bij (`LastParentPath`) | Opgezocht in het auditlog op onderwerp en tijd |
| `Deleted` zonder de rol | Graph: **alles** in Verwijderde items en Recoverable Items\Deletions dat in het venster is gewijzigd, plus elke geauditeerde `MoveToDeletedItems` / `SoftDelete` | Het auditlog voor geauditeerde verwijderingen; anders — en voor alles wat uit Verwijderde items zelf is verwijderd — het **Postvak IN** | Uit de auditrecord (exact, op MessageId) |
| `Moved` | Elke geauditeerde `Move` op die dag, teruggevolgd naar de **eerste** map die het bericht verliet, via Graph gevonden op internet-MessageId en teruggezet | Alleen uit het auditlog | Uit de auditrecord |

> **Waarom het auditlog dubbel telt.** Een gewone verplaatsing laat geen spoor na van waar een bericht vandaan kwam — niet in Graph, niet in Exchange. De auditrecord is de enige plek die het weet, en dezelfde record noemt de persoon. Verplaatsingen **uit** Verwijderde items of Recoverable Items worden met rust gelaten: dat waren herstelacties, en die terugdraaien zou het bericht opnieuw verwijderen.

Berichten die zonder auditrecord in **Archief** zijn beland — Exchange auditeert de eigen `Move` van de eigenaar standaard niet, en archiveren door een app (bijv. [`Move-InboxToArchive.ps1`](Move-InboxToArchive.ps1)) kan ook ontbreken — worden uit de map Archief **opgesomd** op basis van wijzigingstijd. Ze worden alleen naar het Postvak IN verplaatst met `-UnauditedArchiveToInbox`, omdat lezen of markeren van een bericht die tijd ook verandert.

**Parameters**

| Parameter | Verplicht | Standaard | Omschrijving |
|-----------|----------|---------|-------------|
| `-Mailbox` | Ja | — | UPN of primair SMTP-adres |
| `-Date` | Een van | — | De dag waarop de berichten zijn verplaatst of verwijderd (lokale tijd, hele dag) |
| `-After` | deze | — | Begin van een venster (lokale tijd) in plaats van `-Date`. Zonder `-Before`: **alles vanaf deze datum tot nu**. Aliassen `-From`, `-Since` |
| `-Before` | Nee | nu | Einde van dat venster |
| `-Include` | Nee | `Deleted`, `Moved` | Welk onderdeel moet draaien. Het rapport over wie wat deed wordt altijd gemaakt |
| `-UnauditedArchiveToInbox` | Nee | uit | Ook niet-geauditeerde Archief-items die in het venster zijn gewijzigd naar het Postvak IN verplaatsen |
| `-Apply` | Nee | uit | **Echt terugzetten.** Zonder deze switch rapporteert de run alleen |
| `-OutputPath` | Nee | `C:\Temp\` / `~/Downloads` | Pad van het CSV-rapport; het audittrail wordt ernaast geschreven als `*_Audit.csv` |
| `-TenantId` | Nee | GDAP-klant | Tenant-ID of domein; nodig voor app-only Graph, tenzij af te leiden uit GDAP |
| `-ClientId` | Nee | — | Je eigen App Registration (Mail.ReadWrite-applicatiemachtiging) — slaat de tijdelijke app over |
| `-ClientSecret` | Nee | — | Client secret voor `-ClientId` |
| `-CertificateThumbprint` | Nee | — | Certificaatvingerafdruk voor `-ClientId` |
| `-AppOnly` | Nee | uit | Je eigen app met `ClientId` en `CertificateThumbprint` uit `graph.appid.json` |
| `-Delegated` | Nee | uit | Graph als jezelf (`Mail.ReadWrite.Shared`), helemaal geen app. Werkt alleen als je **al Full Access hebt** op `-Mailbox`; een vers toegekend recht kan tot een uur nodig hebben om Graph te bereiken. Niet onder GDAP |

**Voorbeelden**

```powershell
# 1. Wat is er op 25 september verplaatst of verwijderd, door wie, en wat zou er teruggaan?
.\Restore-MailboxMessages.ps1 -Mailbox "user@contoso.com" -Date 2026-09-25

# 2. Alles terugzetten
.\Restore-MailboxMessages.ps1 -Mailbox "user@contoso.com" -Date 2026-09-25 -Apply

# 2b. Alles wat vanaf 20 september tot nu is verplaatst of verwijderd
.\Restore-MailboxMessages.ps1 -Mailbox "user@contoso.com" -Since 2026-09-20 -Apply

# 3. Alleen de verwijderingen, in een precies venster — helemaal geen Graph-toegang nodig
.\Restore-MailboxMessages.ps1 -Mailbox "user@contoso.com" -Include Deleted `
    -After "2026-09-25 14:00" -Before "2026-09-25 16:00" -Apply

# 3b. Graph als jezelf, omdat je al Full Access op de mailbox hebt
.\Restore-MailboxMessages.ps1 -Mailbox "user@contoso.com" -Date 2026-09-25 -Delegated

# 4. Een run van Move-InboxToArchive.ps1 terugdraaien, inclusief de niet-geauditeerde Archief-items
.\Restore-MailboxMessages.ps1 -Mailbox "user@contoso.com" -Date 2026-09-25 `
    -UnauditedArchiveToInbox -Apply
```

**Uitvoer**

- Op het scherm: elk bericht met `[Status] time | current folder -> target folder | who | subject`, daarna een tabel **wie verplaatste / verwijderde wat**, gegroepeerd op account, aanmeldingstype, client en bewerking, met de tijdspanne en IP-adressen.
- `MailRestore_<mailbox>_<timestamp>.csv` — één rij per bericht: fase, tijdstip van de actie, uitvoerder, aanmeldingstype, client, IP, onderwerp, huidige en doelmap, status.
- `..._Audit.csv` — het ruwe audittrail voor de mailbox in het venster, inclusief bron- en doelmap.

| Status | Betekenis |
|--------|---------|
| `WouldRestore` / `Restored` | Voorbeeld / uitgevoerd |
| `AlreadyInPlace` | Staat al terug in de oorspronkelijke map — niets te doen |
| `NotFound` | Het verplaatste bericht staat niet meer in de mailbox (inmiddels verwijderd — het onderdeel `Deleted` dekt dat) |
| `Ambiguous` | Meerdere kopieën met dezelfde MessageId; met rust gelaten |
| `NotAudited` | Archief-item zonder auditrecord, alleen opgesomd |
| `NotRestored` | Restore-RecoverableItems meldde niets, maar het item staat daarna nog in Recoverable Items |
| `Unreachable` | Hard verwijderd (Purges) terwijl de run via de Graph-route loopt — alleen `Restore-RecoverableItems` met de rol kan erbij |

**Vereiste rechten**

| Onderdeel | Recht |
|------|-----------|
| Auditlog | **View-Only Audit Logs** of **Audit Logs** (Organization Management / Compliance Management) |
| `Deleted` | **Mailbox Import Export** — standaard in geen enkele rolgroep: `New-ManagementRoleAssignment -Role "Mailbox Import Export" -User admin@contoso.com`, daarna opnieuw verbinden. **Optioneel:** zonder deze rol zet de run via Graph terug — alles behalve hard verwijderde items |
| `Moved` (en `Deleted` zonder de rol) | App-only `Mail.ReadWrite`, via dezelfde drie routes als [`Remove-PhishingMessage.ps1`](#remove-phishingmessageps1): een bestaande app-only-sessie, `-ClientId` / `-AppOnly`, of een tijdelijke app die aan het eind wordt verwijderd (de aanmelding daarvoor is altijd een apparaatcode via gewone REST, om uit de buurt te blijven van de MSAL van de Exchange-module). App-only is de standaard omdat een gedelegeerd token de mailbox van een andere gebruiker alleen met Full Access erop bereikt; `-Delegated` neemt die route als je die hebt |
| Aanmelden | Exchange Online maakt gedelegeerd verbinding via `Connect-M365Exchange` (apparaatcode en GDAP-klant volgens `load.config.ps1`); een bestaande sessie wordt hergebruikt |

**Opmerkingen**

- Het auditlog loopt **30–90 minuten** (soms 24 uur) achter. Een run op dezelfde dag kan de laatste acties missen.
- Langere vensters (`-Since`) worden in het auditlog **dag voor dag** doorzocht, omdat één zoekopdracht stopt bij 50.000 records voor de hele tenant. De run waarschuwt als het venster verder terug reikt dan wat nog bewaard wordt: Recoverable Items bewaart verwijderde items gedurende `RetainDeletedItemsFor` (standaard 14 dagen, maximaal 30), tenzij de mailbox onder een hold staat, en het auditlog bewaart doorgaans 180 dagen. Voor Verwijderde items zelf geldt die limiet niet.
- De run toont welke van `Move`, `MoveToDeletedItems`, `SoftDelete`, `HardDelete` **niet geauditeerd** worden voor eigenaar, gemachtigde en beheerder op deze mailbox. Standaard wordt de eigen `Move` van de eigenaar niet geauditeerd — die verplaatsingen zijn niet terug te volgen of toe te schrijven. Schakel het in voor de volgende keer: `Set-Mailbox user@contoso.com -AuditOwner @{Add='Move'}`.
- Verplaatsingen door een inboxregel draaien als de mailbox zelf en worden vaak niet geauditeerd. Verplaatsingen naar de **online archiefmailbox** door een bewaarbeleid gaan naar een andere mailbox en vallen buiten het bereik.
- Als de oorspronkelijke map van een verplaatsing niet meer bestaat, gaat het bericht naar het Postvak IN en vermeldt de CSV dat.
