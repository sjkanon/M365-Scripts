[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [Office365Toolkit](../readme.nl.md) › **Exchange**

# Office365Toolkit / Exchange

Hygiënebasislijn voor mailboxen, doorstuurrisico via inboxregels, mailbox-invoegtoepassingen, zoeken in het Unified
Audit Log en message trace-rapportage. Maakt automatisch verbinding met Exchange Online
als er geen sessie actief is; hergebruikt een bestaande sessie als er al
verbinding is.

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
| [`Search-MailboxAuditLog.ps1`](Search-MailboxAuditLog.ps1) ([docs](#search-mailboxauditlogps1)) | Het Unified Audit Log doorzoeken op aanmeldingen en mailboxaanmeldingen |

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
| `-TenantId` | Nee | Tenant-ID of domein van Entra ID |

**Voorbeelden**

```powershell
.\Test-MailboxSecurityBaseline.ps1

.\Test-MailboxSecurityBaseline.ps1 -Mailbox "user@contoso.com"

.\Test-MailboxSecurityBaseline.ps1 -MinAuditLogAgeLimitDays 180 -MinRetainDeletedItemsDays 30
```

**Opmerkingen**
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
| `-TenantId` | Nee | Tenant-ID of domein van Entra ID |

**Voorbeelden**

```powershell
.\Test-MailboxForwardingRisk.ps1

.\Test-MailboxForwardingRisk.ps1 -Mailbox "user@contoso.com" -IncludeDisabledRules
```

**Opmerkingen**
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
| `-TenantId` | Nee | Tenant-ID of domein van Entra ID |

**Voorbeelden**

```powershell
.\Get-MailboxAddIns.ps1

.\Get-MailboxAddIns.ps1 -Mailbox "user@contoso.com"
```

**Vereiste module:** `ExchangeOnlineManagement`

---

### Search-MailboxAuditLog.ps1

Generieke wrapper voor zoeken in het Unified Audit Log (`Search-UnifiedAuditLog`), die
alle overeenkomende resultaten doorpagineert. Zoekt standaard in de laatste 2 dagen naar interactieve
aanmeldingen (geslaagd/mislukt) en mailboxaanmeldingen. Volledig te parametriseren voor andere
recordtypen/bewerkingen/datumbereiken/gebruikers. Alleen-lezen.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-Days` | Nee | Aantal dagen terug vanaf nu om te zoeken (standaard `2`); genegeerd als `-StartDate` is opgegeven |
| `-StartDate` | Nee | Expliciet begin van het venster (heeft voorrang op `-Days`) |
| `-EndDate` | Nee | Expliciet einde van het venster (standaard: nu) |
| `-RecordType` | Nee | Recordtype(n) van het auditlog (standaard `AzureActiveDirectoryStsLogon`, `ExchangeItem`) |
| `-Operations` | Nee | Naam/namen van bewerkingen (standaard `UserLoggedIn`, `UserLoginFailed`, `MailboxLogin`) |
| `-UserIds` | Nee | Beperken tot specifieke gebruiker(s) |
| `-OutputPath` | Nee | Pad voor het CSV-rapport |
| `-TenantId` | Nee | Tenant-ID of domein van Entra ID |

**Voorbeelden**

```powershell
# Laatste 2 dagen, aanmeldingen + mailboxaanmeldingen
.\Search-MailboxAuditLog.ps1

# Laatste 30 dagen, alleen mislukte aanmeldingen, één gebruiker
.\Search-MailboxAuditLog.ps1 -Days 30 -RecordType AzureActiveDirectoryStsLogon -Operations UserLoginFailed -UserIds "user@contoso.com"

# Expliciet venster
.\Search-MailboxAuditLog.ps1 -StartDate (Get-Date "2026-07-01") -EndDate (Get-Date "2026-07-15")
```

**Opmerkingen**
- Het Unified Audit Log is niet direct bijgewerkt — reken op 30-60 minuten voordat recente
  activiteit verschijnt.

**Vereiste module:** `ExchangeOnlineManagement`

---

### Get-MessageTraceReport.ps1

Rapporteert de mailflow over een recent venster (message trace gaat maar ~10 dagen terug),
met optionele filters op afzender/ontvanger/status. Gebruikt bij voorkeur de nieuwere
cmdlet `Get-MessageTraceV2` en valt op oudere moduleversies terug op de klassieke `Get-MessageTrace`.
Alleen-lezen.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-Hours` | Nee | Aantal uren terug vanaf nu om te zoeken (standaard `48`); genegeerd als `-StartDate` is opgegeven |
| `-StartDate` | Nee | Expliciet begin van het venster |
| `-EndDate` | Nee | Expliciet einde van het venster (standaard: nu) |
| `-SenderAddress` | Nee | Filteren op een specifieke afzender |
| `-RecipientAddress` | Nee | Filteren op een specifieke ontvanger |
| `-Status` | Nee | Filteren op een bezorgstatus (bijv. `Delivered`, `Failed`, `Quarantined`) |
| `-OutputPath` | Nee | Pad voor het CSV-rapport |
| `-TenantId` | Nee | Tenant-ID of domein van Entra ID |

**Voorbeelden**

```powershell
.\Get-MessageTraceReport.ps1

.\Get-MessageTraceReport.ps1 -Hours 24 -RecipientAddress "user@contoso.com"

.\Get-MessageTraceReport.ps1 -SenderAddress "billing@vendor.com" -Status Failed
```

**Vereiste module:** `ExchangeOnlineManagement`
