[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [LegacyUtilities](../readme.nl.md) › **Entra**

# Legacy Utilities — Entra

Hulpscripts voor groepslidmaatschap en Conditional Access via Microsoft Graph. Ze melden aan via [`Connect-M365.ps1`](../../Startup/readme.nl.md): standaard delegated als beheerder (browser, of apparaatcode / GDAP-klant volgens `load.config.ps1`), app-only met `-ClientId` + `-CertificateThumbprint` of `-AppOnly` (app uit `graph.appid.json`). Een passende sessie voor de juiste tenant wordt hergebruikt en blijft verbonden; alleen een sessie die het script zelf opende, wordt verbroken. Elk script accepteert `-TenantId`, `-ClientId`, `-CertificateThumbprint` en `-AppOnly`.

---

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`Add-M365GroupMember.ps1`](Add-M365GroupMember.ps1) ([docs](#add-m365groupmemberps1)) | Groepsleden toevoegen of verwijderen, afzonderlijk of in bulk |
| [`Backup-ConditionalAccessPolicies.ps1`](Backup-ConditionalAccessPolicies.ps1) ([docs](#backup-conditionalaccesspoliciesps1)) | Exporteer elk CA-beleid naar een eigen JSON-bestand |

---

### Add-M365GroupMember.ps1

Voegt één lid of een CSV/TXT-lijst met leden toe aan een groep, of verwijdert ze eruit
(op object-ID of weergavenaam). Standaard een proefdraai.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-GroupId` | Ja | Object-ID of weergavenaam van de groep |
| `-Member` | * | Eén UPN/object-ID |
| `-CsvPath` | * | CSV/TXT-lijst met leden |
| `-Action` | Nee | `Add` (standaard) of `Remove` |
| `-Apply` | Nee | Wijzig het lidmaatschap echt (standaard: voorbeeldweergave) |
| `-TenantId` | Nee | Tenant-ID of domein van Entra ID (standaard: de GDAP-klant als `authMode` GDAP is) |
| `-ClientId` / `-CertificateThumbprint` | Nee | App-only aanmelden met deze app-registratie en dit certificaat |
| `-AppOnly` | Nee | App-only aanmelden met de app uit `graph.appid.json` |

```powershell
.\Add-M365GroupMember.ps1 -GroupId "Sales Team" -Member "j.doe@contoso.com" -Apply
.\Add-M365GroupMember.ps1 -GroupId "Sales Team" -CsvPath .\leavers.csv -Action Remove -Apply
```

**Opmerkingen**
- Aanhalingstekens in een groepsnaam worden ge-escaped voor het filter, en een naam die bij meer dan één groep past, stopt het script (voorheen werd de eerste treffer genomen)
- Gedelegeerde scopes: `GroupMember.ReadWrite.All`, `Group.Read.All`, `User.Read.All` (die laatste ontbrak voor het opzoeken van leden)

---

### Backup-ConditionalAccessPolicies.ps1

Exporteert elk Conditional Access-beleid naar één JSON-bestand per beleid (genoemd naar het
beleids-ID), plus een samenvatting in `_index.csv`. Alleen-lezen. Gemoderniseerde vervanging
van een oud script dat de uitgefaseerde AzureADPreview-module gebruikte.

```powershell
.\Backup-ConditionalAccessPolicies.ps1 -OutputPath "C:\Backups\CA-2026-07-24"
```

---

**Vereiste modules**

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

**Opmerkingen**
- Elk bestand bevat het beleid precies zoals Graph het teruggeeft (`GET /identity/conditionalAccess/policies`, gepagineerd); voorheen werden de SDK-objecten geserialiseerd met `-Depth 10`, wat SDK-wrappereigenschappen toevoegt en geneste voorwaarden kan afkappen
- Gedelegeerde scope: `Policy.Read.All`
