[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [Office365Toolkit](../readme.nl.md) › **Security**

# Office365Toolkit / Security

Scripts voor rapportage en hardening van de beveiligingsstatus: Secure Score-trend, beoordeling van toestemmingen voor enterprise-
apps, aanmelding op gedeelde mailboxen dichtzetten, en een EOP-basislijn voor antispam/
antimalware.

---

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`Get-SecureScoreReport.ps1`](Get-SecureScoreReport.ps1) ([docs](#get-securescorereportps1)) | Rapport van de Secure Score-trend en de zwakste controls |
| [`Remove-EnterpriseAppConsent.ps1`](Remove-EnterpriseAppConsent.ps1) ([docs](#remove-enterpriseappconsentps1)) | De OAuth-toestemmingen van een enterprise-app controleren en eventueel intrekken |
| [`Test-SharedMailboxSignIn.ps1`](Test-SharedMailboxSignIn.ps1) ([docs](#test-sharedmailboxsigninps1)) | Directe aanmelding op gedeelde mailboxen rapporteren/uitschakelen |
| [`New-EOPProtectionBaseline.ps1`](New-EOPProtectionBaseline.ps1) ([docs](#new-eopprotectionbaselineps1)) | Basislijnbeleid voor antispam + antimalware in EOP aanmaken/bijwerken |

---

### Get-SecureScoreReport.ps1

Rapporteert de Microsoft Secure Score-trend van de tenant door de tijd heen en splitst de
controlscores van de laatste momentopname uit, gesorteerd van zwakst naar sterkst. Alleen-lezen.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-HistoryCount` | Nee | Aantal historische momentopnamen om mee te nemen (standaard `30`) |
| `-OutputPath` | Nee | Pad voor het CSV-rapport (standaard: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Nee | Tenant-ID of domein van Entra ID |

**Voorbeelden**

```powershell
.\Get-SecureScoreReport.ps1

.\Get-SecureScoreReport.ps1 -HistoryCount 90 -OutputPath C:\Reports
```

**Vereiste scope:** `SecurityEvents.Read.All`
**Vereiste module:** `Microsoft.Graph.Security`

---

### Remove-EnterpriseAppConsent.ps1

Controleert (en trekt met `-Apply` in) de gedelegeerde en applicatiemachtigingen
die één enterprise-applicatie heeft gekregen — handig om onrechtmatige
consent-grants te beoordelen of op te ruimen voordat je een app verwijdert. Precies één van `-AppId` /
`-AppDisplayName` is verplicht, zodat één run niet per ongeluk elke
app in de tenant kan meenemen.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-AppId` | * | Applicatie-ID (client-ID) of object-ID van de service principal |
| `-AppDisplayName` | * | Exacte weergavenaam van de enterprise-app |
| `-IncludeUserConsent` | Nee | Ook gedelegeerde toestemming per gebruiker rapporteren/intrekken (niet alleen tenantbrede beheerderstoestemming) |
| `-Apply` | Nee | De gerapporteerde toestemmingen echt intrekken (standaard: alleen voorbeeldweergave) |
| `-OutputPath` | Nee | Pad voor het CSV-rapport |
| `-TenantId` | Nee | Tenant-ID of domein van Entra ID |

*Precies één van `-AppId` / `-AppDisplayName` is verplicht.

**Voorbeelden**

```powershell
# Alleen voorbeeldweergave
.\Remove-EnterpriseAppConsent.ps1 -AppDisplayName "Suspicious Reporting Tool"

# Tenantbrede beheerderstoestemming + applicatiemachtigingen intrekken
.\Remove-EnterpriseAppConsent.ps1 -AppDisplayName "Suspicious Reporting Tool" -Apply

# Ook de gedelegeerde toestemming van individuele gebruikers intrekken
.\Remove-EnterpriseAppConsent.ps1 -AppId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -IncludeUserConsent -Apply
```

Ondersteunt `-WhatIf` (`SupportsShouldProcess`).

**Vereiste scopes:** `Application.Read.All`, `DelegatedPermissionGrant.ReadWrite.All`, `AppRoleAssignment.ReadWrite.All`
**Vereiste module:** `Microsoft.Graph.Applications`

---

### Test-SharedMailboxSignIn.ps1

Legt de gedeelde mailboxen in Exchange Online naast de accountstatus in Entra ID
en rapporteert welke nog directe interactieve aanmelding toestaan — een veelvoorkomend
zwak punt, omdat gedeelde mailboxen zelden een licentie of MFA-beveiliging hebben. Met
`-Apply` wordt aanmelding uitgeschakeld (`AccountEnabled = $false`) voor elk gevonden ingeschakeld account.
Gedelegeerde toegang (Full Access / Send As) blijft ongewijzigd.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-Mailbox` | Nee | UPN van één gedeelde mailbox. Zonder deze parameter worden alle gedeelde mailboxen gecontroleerd |
| `-Apply` | Nee | Aanmelding echt uitschakelen voor gevonden ingeschakelde accounts (standaard: alleen voorbeeldweergave) |
| `-OutputPath` | Nee | Pad voor het CSV-rapport |
| `-TenantId` | Nee | Tenant-ID of domein van Entra ID |

**Voorbeelden**

```powershell
.\Test-SharedMailboxSignIn.ps1

.\Test-SharedMailboxSignIn.ps1 -Apply

.\Test-SharedMailboxSignIn.ps1 -Mailbox "helpdesk@contoso.com" -Apply
```

Ondersteunt `-WhatIf` (`SupportsShouldProcess`).

**Vereiste scope:** `User.ReadWrite.All`
**Vereiste modules:** `ExchangeOnlineManagement`, `Microsoft.Graph.Users`

---

### New-EOPProtectionBaseline.ps1

Rapporteert het huidige antispam-/antimalwarebeleid van de tenant ten opzichte van een
aanbevolen basislijn, en maakt met `-Apply` (of werkt met `-UpdateExisting` bij)
een basislijnbeleid voor hosted content filter (antispam) en/of een malware
filter-beleid + regel aan, beperkt tot opgegeven ontvangerdomeinen (standaard alle
geaccepteerde domeinen).

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-Domains` | Nee | Ontvangerdomein(en) waartoe de regel(s) worden beperkt (standaard: alle geaccepteerde domeinen) |
| `-Protection` | Nee | `Spam`, `Malware` of `Both` (standaard) |
| `-SpamPolicyName` | Nee | Naam voor het antispambeleid/de antispamregel (standaard `MSP Baseline Anti-Spam`) |
| `-MalwarePolicyName` | Nee | Naam voor het antimalwarebeleid/de antimalwareregel (standaard `MSP Baseline Anti-Malware`) |
| `-UpdateExisting` | Nee | Het beleid ter plekke bijwerken als er al een met dezelfde naam bestaat |
| `-Apply` | Nee | Het beleid echt aanmaken/bijwerken (standaard: alleen voorbeeldweergave) |
| `-TenantId` | Nee | Tenant-ID of domein van Entra ID |

**Voorbeelden**

```powershell
# Alleen voorbeeldweergave
.\New-EOPProtectionBaseline.ps1

# Beide basislijnbeleidsregels aanmaken voor alle geaccepteerde domeinen
.\New-EOPProtectionBaseline.ps1 -Apply

# Alleen antispam, specifieke domeinen, ter plekke bijwerken als het al bestaat
.\New-EOPProtectionBaseline.ps1 -Protection Spam -Domains "contoso.com" -UpdateExisting -Apply
```

**Opmerkingen**
- Dit zijn aanbevelingen voor een basislijn, geen volledige hardening — vergelijk ze
  met de vooraf ingestelde Standard/Strict-beveiligingsbeleidsregels van je tenant voordat je ze
  breed toepast.

Ondersteunt `-WhatIf` (`SupportsShouldProcess`).

**Vereiste module:** `ExchangeOnlineManagement`
