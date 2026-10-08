[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [TenantOnboarding](../readme.nl.md) › **UserManagement**

# UserManagement

Kleine hulpscripts voor gebruikers- en groepsbeheer in Entra ID / Exchange Online, gebruikt bij het onboarden van tenants en bij het doorlopend uitrollen van functies.

---

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`New-DynamicDistributionGroupByFilter.ps1`](New-DynamicDistributionGroupByFilter.ps1) ([docs](#new-dynamicdistributiongroupbyfilterps1)) | Maak een dynamische distributiegroep aan op basis van een functietitel of een eigen ontvangersfilter |
| [`Add-UserToFeatureGroup.ps1`](Add-UserToFeatureGroup.ps1) ([docs](#add-usertofeaturegroupps1)) | Voeg een gebruiker toe aan of verwijder hem uit een opgegeven Entra ID-groep om een aanvullende functie vrij te geven |

---

### New-DynamicDistributionGroupByFilter.ps1

Toont eerst de ontvangers die voldoen aan een Exchange-filter op functietitel (of een eigen filter), en maakt na bevestiging een dynamische distributiegroep met dat filter aan. Alleen Exchange Online: Microsoft Graph heeft geen API voor dynamische distributiegroepen of voor het vooraf bekijken van een ontvangersfilter. Maakt verbinding via `Connect-M365Exchange` ([`Connect-M365.ps1`](../../Startup/Connect-M365.ps1)): standaard gedelegeerd (Exchange-beheerder; apparaatcode volgens `$global:useDeviceCodeAuth`; de GDAP-klant via `-DelegatedOrganization`), app-only met `-ClientId` + `-CertificateThumbprint` en `-TenantId` als het `*.onmicrosoft.com`-domein. Een bestaande Exchange-sessie voor de tenant wordt hergebruikt en blijft open.

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-JobTitle` | * | Functietitel waarop gefilterd wordt |
| `-RecipientFilter` | * | Eigen filtertekenreeks voor Exchange-ontvangers |
| `-Name` | Nee | Groepsnaam (standaard: de functietitel; verplicht met `-RecipientFilter`) |
| `-PrimarySmtpAddress` | Nee | SMTP-adres voor de nieuwe groep (standaard: afgeleid door Exchange) |
| `-TenantId` | Nee | Tenantdomein of -ID (standaard: GDAP-klant, anders de tenant waarbij je je aanmeldt; app-only vereist het domein) |
| `-ClientId` / `-CertificateThumbprint` | Nee | App-only aanmelden |
| `-AppOnly` | Nee | App-only met ClientId en vingerafdruk uit `graph.appid.json` |
| `-Apply` | Nee | Maak de groep echt aan (standaard: alleen voorbeeldweergave) |

*Een van `-JobTitle` / `-RecipientFilter` is verplicht.

```powershell
.\New-DynamicDistributionGroupByFilter.ps1 -JobTitle "Sales Manager"
.\New-DynamicDistributionGroupByFilter.ps1 -JobTitle "Sales Manager" -Apply

.\New-DynamicDistributionGroupByFilter.ps1 -Name "Finance Dept" `
    -RecipientFilter "((Department -eq 'Finance') -and (ExchangeUserAccountControl -ne 'AccountDisabled'))" -Apply
```

**Vereiste module**
```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

---

### Add-UserToFeatureGroup.ps1

Voegt een gebruiker toe als direct lid van een Entra ID-beveiligingsgroep, of verwijdert hem eruit: de ondersteunde manier om een Conditional Access-beleid, Intune-toewijzing of licentiegroep te beperken tot een deel van de gebruikers. Vervangt de legacy-techniek waarbij een gebruiker werd gemarkeerd via een aangepast mailboxkenmerk (alleen voor rapportage, bakende in werkelijkheid niets af).

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-UserId` | Ja | UPN of object-ID |
| `-GroupId` | * | Object-ID van de doelgroep |
| `-GroupName` | * | Weergavenaam van de doelgroep (wordt automatisch opgezocht) |
| `-Remove` | Nee | Verwijder in plaats van toevoegen |
| `-TenantId` | Nee | Tenant-ID of domein van Entra ID (standaard: GDAP-klant, anders de tenant waarbij je je aanmeldt) |
| `-ClientId` / `-CertificateThumbprint` | Nee | App-only aanmelden |
| `-AppOnly` | Nee | App-only met ClientId en vingerafdruk uit `graph.appid.json` |
| `-Apply` | Nee | Wijzig het lidmaatschap echt (standaard: alleen voorbeeldweergave) |

*Een van `-GroupId` / `-GroupName` is verplicht.

```powershell
.\Add-UserToFeatureGroup.ps1 -UserId "user@contoso.com" -GroupName "SG - Enable Windows 365" -Apply
.\Add-UserToFeatureGroup.ps1 -UserId "user@contoso.com" -GroupName "SG - Enable Windows 365" -Remove -Apply
```

**Vereiste scopes:** `GroupMember.ReadWrite.All`, `User.Read.All`, `Group.Read.All`
**Vereiste modules:** `Microsoft.Graph.Users`, `Microsoft.Graph.Groups`

---

`Add-UserToFeatureGroup.ps1` meldt zich aan via [`Connect-M365.ps1`](../../Startup/Connect-M365.ps1): **standaard gedelegeerd** (je meldt je aan als beheerder; apparaatcode als `$global:useDeviceCodeAuth` is ingesteld; onder GDAP de klant uit `$global:cid`), **app-only** met `-ClientId` + `-CertificateThumbprint` of `-AppOnly`. Een bestaande Graph-sessie wordt alleen hergebruikt als die voor de juiste tenant is en de bovenstaande scopes heeft; het script verbreekt alleen een sessie die het zelf heeft geopend.
