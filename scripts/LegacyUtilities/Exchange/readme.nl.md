[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [LegacyUtilities](../readme.nl.md) › **Exchange**

# Legacy Utilities — Exchange

Scripts voor mailbox- en contactbeheer, gemoderniseerd vanuit een reeks oude ad-hocscripts.
Ze maken automatisch verbinding met Exchange Online / Microsoft Graph als er geen sessie
actief is, en hergebruiken een bestaande sessie als je al verbonden bent.

---

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`Set-MailboxFolderPermission.ps1`](Set-MailboxFolderPermission.ps1) ([docs](#set-mailboxfolderpermissionps1)) | Ken een maprecht toe op elke map in een mailbox |
| [`Add-MailboxDelegateAccess.ps1`](Add-MailboxDelegateAccess.ps1) ([docs](#add-mailboxdelegateaccessps1)) | Ken Full Access / Send As toe op één mailbox, een CSV-lijst of elke mailbox |
| [`New-BulkSharedMailboxes.ps1`](New-BulkSharedMailboxes.ps1) ([docs](#new-bulksharedmailboxesps1)) | Maak gedeelde mailboxen in bulk aan vanuit een CSV |
| [`New-BulkMailContacts.ps1`](New-BulkMailContacts.ps1) ([docs](#new-bulkmailcontactsps1)) | Maak Mail Contacts in bulk aan vanuit een CSV en voeg ze optioneel toe aan een distributiegroep |
| [`Sync-UserContacts.ps1`](Sync-UserContacts.ps1) ([docs](#sync-usercontactsps1)) | Zet een gedeelde contactenlijst in de persoonlijke Outlook-contacten van gebruikers |
| [`Start-MailboxMessageTraceReport.ps1`](Start-MailboxMessageTraceReport.ps1) ([docs](#start-mailboxmessagetracereportps1)) | Dien aanvragen in voor historische message trace-rapporten |
| [`Remove-DuplicateMailItems.ps1`](Remove-DuplicateMailItems.ps1) ([docs](#remove-duplicatemailitemsps1)) | Vind/verwijder dubbele berichten in een mailboxmap via Graph |

---

### Set-MailboxFolderPermission.ps1

Past één maprecht toe op elke map in een mailbox (met uitzondering van Sync Issues,
Recoverable Items, Purges, Versions en Deletions). Bedoeld voor gedelegeerde toegang tot de
volledige mailbox in gevallen waarin `Add-MailboxPermission -AccessRights FullAccess` alleen
niet de juiste vorm is. Standaard een proefdraai.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-Mailbox` | Ja | Doelmailbox |
| `-User` | Ja | Gebruiker die toegang krijgt |
| `-AccessRights` | Ja | Niveau van het maprecht (Owner, Editor, Reviewer, ...) |
| `-Apply` | Nee | Ken het recht echt toe (standaard: voorbeeldweergave) |
| `-TenantId` | Nee | Tenant-ID of domein van Entra ID |

```powershell
.\Set-MailboxFolderPermission.ps1 -Mailbox "shared@contoso.com" -User "j.doe@contoso.com" -AccessRights Editor -Apply
```

---

### Add-MailboxDelegateAccess.ps1

Voegt drie oude, bijna identieke scripts samen tot één script (Full Access + Send As toekennen
op één mailbox; hetzelfde voor een CSV-lijst met gedeelde mailboxen; een algemene toekenning op
elke mailbox in de organisatie). Standaard een proefdraai.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-User` | Ja | Gebruiker die gedelegeerde toegang krijgt |
| `-Mailbox` | * | Eén doelmailbox |
| `-CsvPath` | * | CSV/TXT-lijst met doelmailboxen |
| `-AllMailboxes` | * | Elke gebruikers-/gedeelde mailbox in de tenant |
| `-AccessRights` | Nee | `FullAccess`, `SendAs` of `Both` (standaard) |
| `-AutoMapping` | Nee | Schakel automatische toewijzing in Outlook in (standaard: uit) |
| `-Apply` | Nee | Ken de toegang echt toe (standaard: voorbeeldweergave) |
| `-OutputPath` | Nee | Pad voor het CSV-rapport |
| `-TenantId` | Nee | Tenant-ID of domein van Entra ID |

*Precies één van `-Mailbox` / `-CsvPath` / `-AllMailboxes` bepaalt het bereik.

```powershell
.\Add-MailboxDelegateAccess.ps1 -Mailbox "sales@contoso.com" -User "j.doe@contoso.com" -Apply
.\Add-MailboxDelegateAccess.ps1 -AllMailboxes -User "helpdesk@contoso.com"   # bekijk eerst het bereik
```

---

### New-BulkSharedMailboxes.ps1

Maakt gedeelde mailboxen aan vanuit een CSV (`Name`, `PrimarySmtpAddress`, optioneel
`Alias`). Standaard een proefdraai.

```powershell
.\New-BulkSharedMailboxes.ps1 -CsvPath .\sharedmailboxes.csv -Apply
```

---

### New-BulkMailContacts.ps1

Maakt Mail Contacts aan vanuit een CSV (`Name`, `ExternalEmailAddress`) en voegt elk nieuw
contact optioneel toe aan een distributiegroep. Standaard een proefdraai.

```powershell
.\New-BulkMailContacts.ps1 -CsvPath .\contacts.csv -DistributionGroup "everyone@contoso.com" -Apply
```

---

### Sync-UserContacts.ps1

Zet via Microsoft Graph een CSV-contactenlijst in de persoonlijke map Contactpersonen van een
opgegeven lijst gebruikers of van elk lid van een groep. Elk contact dat het script aanmaakt,
krijgt een markering in `PersonalNotes`, zodat een latere run met `-RemoveExisting` kan
opschonen en verversen zonder de eigen contacten van de gebruiker aan te raken. Standaard een
proefdraai.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-CsvPath` | Ja | Contactenlijst (DisplayName en EmailAddress verplicht; GivenName/Surname/CompanyName/BusinessPhone/MobilePhone optioneel) |
| `-UserList` / `-GroupId` | * | Doelbereik |
| `-Tag` | Nee | Markering die in PersonalNotes wordt geschreven (standaard: `Synced-by-Sync-UserContacts`) |
| `-RemoveExisting` | Nee | Verwijder eerder gesynchroniseerde contacten vóór het opnieuw importeren |
| `-Apply` | Nee | Schrijf de contacten echt weg (standaard: voorbeeldweergave) |

```powershell
.\Sync-UserContacts.ps1 -CsvPath .\companycontacts.csv -GroupId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -RemoveExisting -Apply
```

---

### Start-MailboxMessageTraceReport.ps1

Dient per mailbox één `Start-HistoricalSearch`-aanvraag in (message trace, filter op
afzender), die als gecomprimeerde export per e-mail wordt afgeleverd. Zo haal je message
trace-gegevens op die ouder zijn dan het venster van 10 dagen dat `Get-MessageTrace` dekt.
Gemoderniseerde vervanging van een oud script dat de uitgefaseerde MSOnline-module gebruikte
om de lijst met mailboxen op te bouwen.

```powershell
.\Start-MailboxMessageTraceReport.ps1 -NotifyAddress "admin@contoso.com" -AllMailboxes -Apply
```

---

### Remove-DuplicateMailItems.ps1

Op Graph gebaseerde vervanging van een EWS-tool van derden voor het verwijderen van dubbele
items, die niet is meegenomen: Microsoft faseert de EWS-API voor Exchange Online uit.
Groepeert de berichten in een map op `internetMessageId`, houdt per groep het oudste bericht
en verwijdert (met `-Apply`) de rest. Standaard een proefdraai.

```powershell
.\Remove-DuplicateMailItems.ps1 -Mailbox "user@contoso.com" -IncludeSubfolders -Apply
```

---

**Vereiste modules**

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Install-Module Microsoft.Graph -Scope CurrentUser
```
