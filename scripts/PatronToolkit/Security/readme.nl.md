[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [PatronToolkit](../readme.nl.md) › **Security**

# Patron Toolkit — Security

Rapportage van de beveiligingsstatus van de tenant: risico van app-toestemmingen, BEC-detectie via mailboxregels, beveiligings-
waarschuwingen, configuratie van e-mailbeveiliging, auditlogging en validatie van SPF/DMARC.

---

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`Get-EntraAppConsents.ps1`](Get-EntraAppConsents.ps1) ([docs](#get-entraappconsentsps1)) | Tenantbreed gedelegeerde + applicatiemachtigingen via OAuth controleren en risicovolle scopes markeren |
| [`Get-SuspiciousInboxRules.ps1`](Get-SuspiciousInboxRules.ps1) ([docs](#get-suspiciousinboxrulesps1)) | Inboxregels in BEC-stijl opsporen (extern doorsturen, stil verwijderen, verborgen map + trefwoord) |
| [`Get-SecurityAlerts.ps1`](Get-SecurityAlerts.ps1) ([docs](#get-securityalertsps1)) | Rapport van de uniforme beveiligingswaarschuwingen van Defender/Entra |
| [`Test-EmailSecurityPosture.ps1`](Test-EmailSecurityPosture.ps1) ([docs](#test-emailsecuritypostureps1)) | Samengevoegd beveiligingsrapport voor Defender for O365 / antispam / DLP / mailflow |
| [`Test-MailboxAuditingConfig.ps1`](Test-MailboxAuditingConfig.ps1) ([docs](#test-mailboxauditingconfigps1)) | Hiaten in het Unified Audit Log + auditing per mailbox rapporteren en optioneel herstellen |
| [`Test-EmailAuthenticationRecords.ps1`](Test-EmailAuthenticationRecords.ps1) ([docs](#test-emailauthenticationrecordsps1)) | SPF- en DMARC-DNS-records valideren |

---

### Get-EntraAppConsents.ps1

Controleert elke enterprise-applicatie (service principal) met gedelegeerde machtigingen
(`OAuth2PermissionGrants`) of applicatiemachtigingen
(`AppRoleAssignments`) en markeert toekenningen die overeenkomen met een lijst van vaak misbruikte
scopes met hoge bevoegdheden. Dit is de klassieke controle op "illicit consent grants" / het risico van apps van
derden.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-RiskyOnly` | Nee | Alleen toekenningen opnemen die als High risk zijn gemarkeerd |
| `-OutputPath` | Nee | Pad voor het CSV-rapport (standaard: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Nee | Tenant-ID of domein. Standaard de GDAP-klant (`load.config.ps1`) of je eigen tenant; verplicht voor app-only |
| `-ClientId` | Nee | App-registratie voor app-only aanmelden (met `-CertificateThumbprint`). Zonder meldt het script gedelegeerd aan, als jezelf |
| `-CertificateThumbprint` | Nee | Certificaatvingerafdruk voor app-only aanmelden met `-ClientId` |
| `-AppOnly` | Nee | App-only aanmelden met de ClientId en vingerafdruk voor de tenant uit `graph.appid.json` |

**Voorbeelden**

```powershell
.\Get-EntraAppConsents.ps1

.\Get-EntraAppConsents.ps1 -RiskyOnly
```

**Opmerkingen**
- De risicomarkering is een heuristiek (tekstvergelijking met een lijst van bekende risicovolle scopes) — beoordeel
  gemarkeerde items handmatig en zie "Normal" niet als garantie dat iets veilig is
- Toekenningen met hoog risico staan nu bovenaan in de consoletabel (sorteren op de tekst van
  de markering zette `Normal` vóór `High`)
- Aanmelden via [`Connect-M365.ps1`](../../Startup/readme.nl.md#connect-m365ps1): Microsoft
  Graph, standaard gedelegeerd (scopes `Application.Read.All`, `Directory.Read.All`);
  app-only met `-ClientId` + `-CertificateThumbprint` of `-AppOnly` (dezelfde
  toepassingsmachtigingen)

---

### Get-SuspiciousInboxRules.ps1

Doorzoekt de inboxregels van elke mailbox op patronen die een gecompromitteerd account
vaak achterlaat: extern doorsturen/omleiden, stil verwijderen, of berichten met bepaalde trefwoorden
(invoice, payment, wire, password, ...) verplaatsen naar een map die zelden wordt bekeken. Standaard is het een
rapport dat alleen leest; `-Apply` schakelt overeenkomsten met hoge zekerheid uit (verwijdert ze niet).

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-Mailbox` | Nee | UPN van één mailbox. Zonder deze parameter worden alle mailboxen gecontroleerd |
| `-Apply` | Nee | Gemarkeerde regels met hoge zekerheid uitschakelen (standaard: alleen rapport) |
| `-OutputPath` | Nee | Pad voor het CSV-rapport (standaard: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Nee | Tenant-ID of domein. Standaard de GDAP-klant (`load.config.ps1`) of je eigen tenant; verplicht voor app-only |
| `-ClientId` | Nee | App-registratie voor app-only aanmelden (met `-CertificateThumbprint`). Zonder meldt het script gedelegeerd aan, als jezelf |
| `-CertificateThumbprint` | Nee | Certificaatvingerafdruk voor app-only aanmelden met `-ClientId` |
| `-AppOnly` | Nee | App-only aanmelden met de ClientId en vingerafdruk voor de tenant uit `graph.appid.json` |

**Voorbeelden**

```powershell
.\Get-SuspiciousInboxRules.ps1

.\Get-SuspiciousInboxRules.ps1 -Mailbox "user@contoso.com"

# Voorbeeldweergave van wat uitgeschakeld zou worden
.\Get-SuspiciousInboxRules.ps1 -Apply -WhatIf

.\Get-SuspiciousInboxRules.ps1 -Apply
```

Ondersteunt `-WhatIf` (`SupportsShouldProcess`).

**Opmerkingen**
- "Extern" wordt bepaald aan de hand van de geaccepteerde domeinen van de tenant (`Get-AcceptedDomain`)
- Doorstuurdoelen naar interne ontvangers (`EX:/o=ExchangeLabs/...`) tellen niet langer als
  extern — alleen SMTP-adressen buiten de geaccepteerde domeinen
- Uitschakelen (niet verwijderen) is bewust — het is terug te draaien en de regel blijft bewaard voor
  onderzoek bij incidentrespons
- Blijft bewust op Exchange Online. Graph heeft `messageRules`, maar gedelegeerd bereikt dat
  alleen mailboxen waartoe de aangemelde beheerder toegang heeft gekregen (een Exchange-
  beheerdersrol opent daar niet de regels van andere gebruikers), en het geeft nooit verborgen
  regels terug — `Get-InboxRule -IncludeHidden` wel, en een regel verbergen is een bekende
  aanvalstechniek
- Aanmelden via [`Connect-M365.ps1`](../../Startup/readme.nl.md#connect-m365ps1): Exchange
  Online, standaard gedelegeerd (device code en de GDAP-klant via `-DelegatedOrganization`);
  app-only met `-ClientId` + `-CertificateThumbprint` of `-AppOnly` (`Exchange.ManageAsApp`
  plus een Exchange-rol; `-TenantId` als domein)

---

### Get-SecurityAlerts.ps1

Rapporteert de uniforme feed met beveiligingswaarschuwingen van Microsoft Graph (`security/alerts_v2`), die
Defender for Office 365, Defender for Endpoint, Defender for Identity, Defender for
Cloud Apps en Entra ID Protection omvat.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-Days` | Nee | Terugkijkvenster in dagen (standaard: `30`) |
| `-Severity` | Nee | Filter: `informational`, `low`, `medium`, `high` |
| `-Status` | Nee | Filter: `new`, `inProgress`, `resolved` |
| `-OutputPath` | Nee | Pad voor het CSV-rapport (standaard: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Nee | Tenant-ID of domein. Standaard de GDAP-klant (`load.config.ps1`) of je eigen tenant; verplicht voor app-only |
| `-ClientId` | Nee | App-registratie voor app-only aanmelden (met `-CertificateThumbprint`). Zonder meldt het script gedelegeerd aan, als jezelf |
| `-CertificateThumbprint` | Nee | Certificaatvingerafdruk voor app-only aanmelden met `-ClientId` |
| `-AppOnly` | Nee | App-only aanmelden met de ClientId en vingerafdruk voor de tenant uit `graph.appid.json` |

**Voorbeelden**

```powershell
.\Get-SecurityAlerts.ps1

.\Get-SecurityAlerts.ps1 -Days 7 -Severity high,medium -Status new,inProgress
```

**Opmerkingen**
- Waarschuwingen worden gesorteerd high → medium → low → informational, nieuwste eerst (de
  eerdere alfabetische sortering zette `medium` eerst en `high` als laatste)
- Aanmelden via [`Connect-M365.ps1`](../../Startup/readme.nl.md#connect-m365ps1): Microsoft
  Graph, standaard gedelegeerd (scope `SecurityAlert.Read.All` plus een rol Security
  Reader); app-only met `-ClientId` + `-CertificateThumbprint` of `-AppOnly`
  (toepassingsmachtiging `SecurityAlert.Read.All`)

---

### Test-EmailSecurityPosture.ps1

Samengevoegd rapport (alleen-lezen) over de configuratie van e-mailbeveiliging in Exchange Online / Defender for Office 365:
Safe Links, Safe Attachments, antimalware, antispam
(inkomend/uitgaand), verbindingsfilter, externe domeinen (extern automatisch doorsturen), DLP-
beleid en een overzicht van waarschuwingsbeleid. Controleert optioneel (`-IncludeMailboxDetail`) ook
per mailbox of POP/IMAP is ingeschakeld en of er een litigation hold is.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-IncludeMailboxDetail` | Nee | Ook per mailbox verouderde protocollen + litigation hold controleren (trager) |
| `-OutputPath` | Nee | Pad voor het CSV-rapport (standaard: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Nee | Tenant-ID of domein. Standaard de GDAP-klant (`load.config.ps1`) of je eigen tenant; verplicht voor app-only |
| `-ClientId` | Nee | App-registratie voor app-only aanmelden (met `-CertificateThumbprint`). Zonder meldt het script gedelegeerd aan, als jezelf |
| `-CertificateThumbprint` | Nee | Certificaatvingerafdruk voor app-only aanmelden met `-ClientId` |
| `-AppOnly` | Nee | App-only aanmelden met de ClientId en vingerafdruk voor de tenant uit `graph.appid.json` |

**Voorbeelden**

```powershell
.\Test-EmailSecurityPosture.ps1

.\Test-EmailSecurityPosture.ps1 -IncludeMailboxDetail
```

**Opmerkingen**
- Dit is de belangrijkste samenvoeging, van ~18 upstream-scripts met één doel — zie
  `.NOTES` van het script voor de volledige lijst
- De wijzigende tegenhangers `*-set.ps1` / `*-del.ps1` uit het bronproject zijn
  bewust niet overgezet — elk ervan hardcodeerde de specifieke "aanbevolen" waarden van één MSP, zonder
  mogelijkheid om per tenant af te wijken
- Security & Compliance PowerShell (DLP, waarschuwingsbeleid) wordt nu geopend voor dezelfde
  tenant en aanmelding als Exchange (`Connect-M365Exchange -IncludeCompliance`); de eerdere
  kale `Connect-IPPSSession` negeerde `-TenantId` en las onder GDAP dus je eigen tenant
- `-IncludeMailboxDetail` leest POP/IMAP met één bulkaanroep `Get-EXOCASMailbox` in plaats
  van één `Get-CASMailbox` per mailbox
- Aanmelden via [`Connect-M365.ps1`](../../Startup/readme.nl.md#connect-m365ps1), standaard
  gedelegeerd (device code en de GDAP-klant via `-DelegatedOrganization`); app-only met
  `-ClientId` + `-CertificateThumbprint` of `-AppOnly` (`-TenantId` als domein). Blijft op
  Exchange Online / Security & Compliance: Graph heeft geen API voor EOP/Defender for Office
  365-beleid, externe domeinen, CAS-protocollen, DLP of waarschuwingsbeleid

**Vereiste modules**
```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

---

### Test-MailboxAuditingConfig.ps1

Rapporteert (en herstelt met `-Apply`) hiaten in het Unified Audit Log en in de auditlogging per mailbox:
of het organisatiebrede Unified Audit Log is ingeschakeld, of mailboxauditing voor de hele
organisatie is uitgezet (`AuditDisabled`, alleen gerapporteerd), en of bij elke mailbox
`AuditEnabled` aan staat met een voldoende hoge `AuditLogAgeLimit`.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-MinimumAuditLogAgeDays` | Nee | Minimaal aanvaardbare bewaartermijn in dagen (standaard: `180`) |
| `-Apply` | Nee | Het Unified Audit Log inschakelen en gemarkeerde mailboxen herstellen (standaard: alleen rapport) |
| `-OutputPath` | Nee | Pad voor het CSV-rapport (standaard: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Nee | Tenant-ID of domein. Standaard de GDAP-klant (`load.config.ps1`) of je eigen tenant; verplicht voor app-only |
| `-ClientId` | Nee | App-registratie voor app-only aanmelden (met `-CertificateThumbprint`). Zonder meldt het script gedelegeerd aan, als jezelf |
| `-CertificateThumbprint` | Nee | Certificaatvingerafdruk voor app-only aanmelden met `-ClientId` |
| `-AppOnly` | Nee | App-only aanmelden met de ClientId en vingerafdruk voor de tenant uit `graph.appid.json` |

**Voorbeelden**

```powershell
.\Test-MailboxAuditingConfig.ps1

.\Test-MailboxAuditingConfig.ps1 -Apply -WhatIf

.\Test-MailboxAuditingConfig.ps1 -MinimumAuditLogAgeDays 365 -Apply
```

Ondersteunt `-WhatIf` (`SupportsShouldProcess`).

**Opmerkingen**
- Met `AuditDisabled = True` (`Get-OrganizationConfig`) wordt geen enkele mailbox
  geaudit, wat de eigen `AuditEnabled` ook zegt. Het script rapporteert dit maar wijzigt het
  niet — dat is een bewuste organisatiekeuze (`Set-OrganizationConfig -AuditDisabled $false`)
- Aanmelden via [`Connect-M365.ps1`](../../Startup/readme.nl.md#connect-m365ps1): Exchange
  Online, standaard gedelegeerd (device code en de GDAP-klant via `-DelegatedOrganization`);
  app-only met `-ClientId` + `-CertificateThumbprint` of `-AppOnly` (`-TenantId` als
  domein). Blijft op Exchange Online: Graph heeft geen API voor de auditlogschakelaar of de
  auditinstellingen per mailbox

---

### Test-EmailAuthenticationRecords.ps1

Valideert SPF- en DMARC-DNS-records voor een of meer domeinen — controleert of SPF bestaat,
waar verwacht `spf.protection.outlook.com` bevat en niet te ruim `+all` gebruikt; controleert of DMARC
bestaat, rapporteert het beleid (`none`/`quarantine`/`reject`) en of er een adres voor
aggregate-rapporten is ingesteld. DKIM valt bewust buiten scope — gebruik daarvoor
[`Test-DkimConfig.ps1`](../../Exchange/readme.nl.md).

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-Domain` | Nee | Een of meer domeinen. Zonder deze parameter worden ze automatisch opgehaald via Microsoft Graph |
| `-OutputPath` | Nee | Pad voor het CSV-rapport (standaard: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Nee | Alleen gebruikt om domeinen via Graph op te halen als `-Domain` ontbreekt. Tenant-ID of domein. Standaard de GDAP-klant (`load.config.ps1`) of je eigen tenant; verplicht voor app-only |
| `-ClientId` | Nee | App-registratie voor app-only aanmelden (met `-CertificateThumbprint`). Zonder meldt het script gedelegeerd aan, als jezelf |
| `-CertificateThumbprint` | Nee | Certificaatvingerafdruk voor app-only aanmelden met `-ClientId` |
| `-AppOnly` | Nee | App-only aanmelden met de ClientId en vingerafdruk voor de tenant uit `graph.appid.json` |

**Voorbeelden**

```powershell
.\Test-EmailAuthenticationRecords.ps1 -Domain "contoso.com"

.\Test-EmailAuthenticationRecords.ps1
```

**Opmerkingen**
- Alleen voor Windows (gebruikt `Resolve-DnsName`)
- Het DMARC-beleid wordt uit de tag `p=` zelf gelezen; een record met `sp=` vóór `p=`
  rapporteerde voorheen het subdomeinbeleid
- Meldt alleen aan als `-Domain` ontbreekt, via
  [`Connect-M365.ps1`](../../Startup/readme.nl.md#connect-m365ps1): Microsoft Graph,
  standaard gedelegeerd (scope `Domain.Read.All`); app-only met `-ClientId` +
  `-CertificateThumbprint` of `-AppOnly`
