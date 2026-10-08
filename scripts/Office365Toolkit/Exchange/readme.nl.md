[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [Office365Toolkit](../readme.nl.md) › **Exchange**

# Office365Toolkit / Exchange

Hygiënebasislijn voor mailboxen, doorstuurrisico via inboxregels, mailbox-invoegtoepassingen en
zoeken in het Unified Audit Log.

**Aanmelden.** Elk script meldt aan via
[`Connect-M365.ps1`](../../Startup/readme.nl.md#connect-m365ps1): standaard delegated (je meldt
je aan als de beheerder), met device code en de GDAP-klant uit `load.config.ps1` — onder GDAP
wordt Exchange Online bereikt met `-DelegatedOrganization`. App-only met `-ClientId` +
`-CertificateThumbprint`, of `-AppOnly` (ClientId en vingerafdruk uit `graph.appid.json`). Een
sessie die al past wordt hergebruikt en blijft verbonden; de scripts verbreken alleen wat ze zelf
hebben geopend. `Search-MailboxAuditLog.ps1` gebruikt Microsoft Graph; de andere drie blijven op
Exchange Online omdat Graph geen API heeft voor wat ze lezen (zie de Opmerkingen per script).

> Deze scripts vullen de bestaande Exchange-auditscripts in
> [`scripts/Exchange/`](../../Exchange/readme.nl.md) aan en dupliceren ze niet — zie de sectie Opmerkingen
> van elk script hieronder voor de precieze afbakening.

---

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`Test-MailboxSecurityBaseline.ps1`](Test-MailboxSecurityBaseline.ps1) ([docs](#test-mailboxsecuritybaselineps1)) | Mailboxen toetsen aan een hygiënebasislijn (auditlogging, bewaring, litigation hold, archief, verouderde protocollen) |
| [`Test-MailboxForwardingRisk.ps1`](Test-MailboxForwardingRisk.ps1) ([docs](#test-mailboxforwardingriskps1)) | Inboxregels en Sweep-regels controleren op patronen van doorsturen/exfiltratie (BEC-indicator) |
| [`Get-MailboxAddIns.ps1`](Get-MailboxAddIns.ps1) ([docs](#get-mailboxaddinsps1)) | Rapport van de Outlook-invoegtoepassingen die per mailbox zijn geïnstalleerd |
| [`Search-MailboxAuditLog.ps1`](Search-MailboxAuditLog.ps1) ([docs](#search-mailboxauditlogps1)) | Het Unified Audit Log doorzoeken (Graph Audit Log Query API) op aanmeldingen en mailboxaanmeldingen |

> Message trace-rapportage staat in [`scripts/PatronToolkit/Exchange/Get-MessageTraceReport.ps1`](../../PatronToolkit/Exchange/readme.nl.md#get-messagetracereportps1) — voor beide toolkits is onafhankelijk een gelijkwaardig script gebouwd, dus er is er maar één behouden (met de ondersteuning voor `-IncludeDetail` uit deze versie erin samengevoegd).

---

### Test-MailboxSecurityBaseline.ps1

Toetst elke mailbox (of één enkele) aan een hygiënebasislijn: auditlogging
ingeschakeld + maximale leeftijd van het log, bewaring van verwijderde items, litigation hold,
archiefstatus, geen doorsturen op mailboxniveau, POP3/IMAP uitgeschakeld. Elke mailbox
krijgt per controle een Pass/Fail en een algehele `Status`. Alleen-lezen.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-Mailbox` | Nee | UPN van één mailbox. Zonder deze parameter worden alle gebruikers- en gedeelde mailboxen gecontroleerd |
| `-MinAuditLogAgeLimitDays` | Nee | Minimaal aanvaardbare maximale leeftijd van het auditlog in dagen (standaard `90`) |
| `-MinRetainDeletedItemsDays` | Nee | Minimaal aanvaardbare bewaartermijn voor verwijderde items in dagen (standaard `30`) |
| `-OutputPath` | Nee | Pad voor het CSV-rapport (standaard: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Nee | Tenant-ID of domein; standaard de GDAP-klant uit `load.config.ps1` (app-only vraagt de domeinvorm) |
| `-ClientId` | Nee | App-registratie voor app-only aanmelding (met `-CertificateThumbprint`) |
| `-CertificateThumbprint` | Nee | Certificaatvingerafdruk voor app-only aanmelding |
| `-AppOnly` | Nee | App-only met ClientId en vingerafdruk uit `graph.appid.json` |

**Voorbeelden**

```powershell
.\Test-MailboxSecurityBaseline.ps1

.\Test-MailboxSecurityBaseline.ps1 -Mailbox "user@contoso.com"

.\Test-MailboxSecurityBaseline.ps1 -MinAuditLogAgeLimitDays 180 -MinRetainDeletedItemsDays 30
```

**Opmerkingen**
- Blijft op Exchange Online: instellingen voor audit, bewaring, hold, archief, doorsturen en
  POP/IMAP (`Get-Mailbox` / `Get-CASMailbox`) hebben geen Graph-API.
- Extern doorsturen op mailboxniveau is hier een eenvoudige aanwezig/afwezig-controle; voor
  een domeinbewuste uitsplitsing extern-vs-intern gebruik je
  [`Get-ExternalForwards.ps1`](../../Exchange/readme.nl.md#get-externalforwardsps1).
- Voor doorsturen via inboxregels/Sweep-regels gebruik je in plaats daarvan `Test-MailboxForwardingRisk.ps1`
  hieronder.

**Vereiste module:** `ExchangeOnlineManagement`

---

### Test-MailboxForwardingRisk.ps1

Controleert de inboxregels en Sweep-regels van elke mailbox op patronen van doorsturen/omleiden/
exfiltratie — een klassieke indicator van business email compromise (BEC)
die niet zichtbaar is op het mailboxobject zelf. Markeert ontvangers van regels als
External/Internal/Unknown op basis van de geaccepteerde domeinen van de tenant. Alleen-lezen.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-Mailbox` | Nee | UPN van één mailbox. Zonder deze parameter worden alle mailboxen gecontroleerd |
| `-IncludeDisabledRules` | Nee | Ook uitgeschakelde regels rapporteren die aan de risicovolle patronen voldoen |
| `-OutputPath` | Nee | Pad voor het CSV-rapport |
| `-TenantId` | Nee | Tenant-ID of domein; standaard de GDAP-klant uit `load.config.ps1` (app-only vraagt de domeinvorm) |
| `-ClientId` | Nee | App-registratie voor app-only aanmelding (met `-CertificateThumbprint`) |
| `-CertificateThumbprint` | Nee | Certificaatvingerafdruk voor app-only aanmelding |
| `-AppOnly` | Nee | App-only met ClientId en vingerafdruk uit `graph.appid.json` |

**Voorbeelden**

```powershell
.\Test-MailboxForwardingRisk.ps1

.\Test-MailboxForwardingRisk.ps1 -Mailbox "user@contoso.com" -IncludeDisabledRules
```

**Opmerkingen**
- Blijft op Exchange Online: Graph leest de inboxregels (`messageRules`) van een andere gebruiker
  alleen met een app-only Mail-machtiging, en heeft geen API voor Sweep-regels.
- Een ontvanger binnen de organisatie (`[EX:/o=...]`) telt nu als Internal; die werd als External
  gemarkeerd omdat er geen domein was om te vergelijken.
- Vult [`Get-ExternalForwards.ps1`](../../Exchange/readme.nl.md#get-externalforwardsps1)
  (`ForwardingSmtpAddress` op mailboxniveau) en `Test-MailboxSecurityBaseline.ps1`
  hierboven (dezelfde controle) aan — dit script dekt alleen de laag van inboxregels/Sweep-regels.

**Vereiste module:** `ExchangeOnlineManagement`

---

### Get-MailboxAddIns.ps1

Toont de Outlook-invoegtoepassingen (`Get-App`) op elke mailbox — zowel centraal uitgerold
als door de gebruiker geïnstalleerd/gesideload. Handig om niet-goedgekeurde invoegtoepassingen op te sporen, een bekende
vector voor phishing en consent-grants. Alleen-lezen.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-Mailbox` | Nee | UPN van één mailbox. Zonder deze parameter worden alle gebruikers- en gedeelde mailboxen gecontroleerd |
| `-OutputPath` | Nee | Pad voor het CSV-rapport |
| `-TenantId` | Nee | Tenant-ID of domein; standaard de GDAP-klant uit `load.config.ps1` (app-only vraagt de domeinvorm) |
| `-ClientId` | Nee | App-registratie voor app-only aanmelding (met `-CertificateThumbprint`) |
| `-CertificateThumbprint` | Nee | Certificaatvingerafdruk voor app-only aanmelding |
| `-AppOnly` | Nee | App-only met ClientId en vingerafdruk uit `graph.appid.json` |

**Voorbeelden**

```powershell
.\Get-MailboxAddIns.ps1

.\Get-MailboxAddIns.ps1 -Mailbox "user@contoso.com"
```

**Opmerkingen**
- Blijft op Exchange Online: Outlook-invoegtoepassingen per mailbox (`Get-App`) hebben geen
  Graph-API.

**Vereiste module:** `ExchangeOnlineManagement`

---

### Search-MailboxAuditLog.ps1

Generieke zoekfunctie voor het Unified Audit Log. Gebruikt standaard de Audit Log Query API van
Microsoft Graph: het script maakt een query aan (`POST /security/auditLog/queries`), controleert
elke 30 seconden of die klaar is en bladert daarna door de records. Zoekt standaard in de laatste
2 dagen naar interactieve aanmeldingen (geslaagd/mislukt) en mailboxaanmeldingen. Volledig te
parametriseren voor andere recordtypen/bewerkingen/datumbereiken/gebruikers. `-UseExchange` voert
dezelfde zoekopdracht uit met `Search-UnifiedAuditLog` in Exchange Online. Alleen-lezen.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-Days` | Nee | Aantal dagen terug vanaf nu om te zoeken (standaard `2`); genegeerd als `-StartDate` is opgegeven |
| `-StartDate` | Nee | Expliciet begin van het venster (heeft voorrang op `-Days`) |
| `-EndDate` | Nee | Expliciet einde van het venster (standaard: nu) |
| `-RecordType` | Nee | Recordtype(n) van het auditlog (standaard `AzureActiveDirectoryStsLogon`, `ExchangeItem`) |
| `-Operations` | Nee | Naam/namen van bewerkingen (standaard `UserLoggedIn`, `UserLoginFailed`, `MailboxLogin`) |
| `-UserIds` | Nee | Beperken tot specifieke gebruiker(s) (UPN) |
| `-UseExchange` | Nee | Zoeken met `Search-UnifiedAuditLog` in Exchange Online in plaats van Graph |
| `-TimeoutMinutes` | Nee | Hoe lang op de Graph-query wordt gewacht (standaard `60`) |
| `-OutputPath` | Nee | Pad voor het CSV-rapport |
| `-TenantId` | Nee | Tenant-ID of domein; standaard de GDAP-klant uit `load.config.ps1` |
| `-ClientId` | Nee | App-registratie voor app-only aanmelding (met `-CertificateThumbprint`) |
| `-CertificateThumbprint` | Nee | Certificaatvingerafdruk voor app-only aanmelding |
| `-AppOnly` | Nee | App-only met ClientId en vingerafdruk uit `graph.appid.json` |

**Voorbeelden**

```powershell
# Laatste 2 dagen, aanmeldingen + mailboxaanmeldingen
.\Search-MailboxAuditLog.ps1

# Laatste 30 dagen, alleen mislukte aanmeldingen, één gebruiker
.\Search-MailboxAuditLog.ps1 -Days 30 -RecordType AzureActiveDirectoryStsLogon -Operations UserLoginFailed -UserIds "user@contoso.com"

# Expliciet venster
.\Search-MailboxAuditLog.ps1 -StartDate (Get-Date "2026-07-01") -EndDate (Get-Date "2026-07-15")

# Dezelfde zoekopdracht via Exchange Online
.\Search-MailboxAuditLog.ps1 -UseExchange
```

**Opmerkingen**
- De Graph-query draait asynchroon in de service en kan enkele minuten duren. Na
  `-TimeoutMinutes` stopt het script met wachten en toont het de query-ID; de query loopt door en
  kan later worden gelezen.
- Recordtypen geef je in de auditlogvorm (`ExchangeItem`); het script zet ze om naar de
  camelCase van de API (`exchangeItem`).
- `Search-UnifiedAuditLog` accepteert één recordtype per aanroep. Eerdere versies gaven beide
  standaardrecordtypen in één aanroep mee; `-UseExchange` voert nu per recordtype een eigen
  gepagineerde zoekopdracht uit.
- Het Unified Audit Log is niet direct bijgewerkt — reken op 30-60 minuten voordat recente
  activiteit verschijnt.

**Vereiste scope:** `AuditLogsQuery.Read.All` (plus een Purview-rol die het auditlog mag doorzoeken, bijv. Audit Logs)
**Vereiste modules:** `Microsoft.Graph.Authentication`; `ExchangeOnlineManagement` voor `-UseExchange`
