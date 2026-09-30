[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [PatronToolkit](../readme.nl.md) › **SharePoint**

# Patron Toolkit — SharePoint

Deelconfiguratie van de SharePoint Online-tenant en audit van externe gebruikers.

---

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`Test-SharePointSharingConfig.ps1`](Test-SharePointSharingConfig.ps1) ([docs](#test-sharepointsharingconfigps1)) | Rapport van de deelinstellingen van de tenant, afwijkingen per site en externe gebruikers |

---

### Test-SharePointSharingConfig.ps1

Rapporteert de tenantbrede instellingen voor extern delen die er bij een beveiligingsreview het meest toe doen
(deelmogelijkheden, standaard linktype, verloop/rechten van anonieme links, opnieuw delen door
externe gebruikers, verouderde authenticatie). Somt optioneel ook elke externe gebruiker (gast)
over alle sitecollecties op en markeert afzonderlijke sites waarvan de deelmogelijkheden
ruimer zijn dan de standaard van de tenant.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-TenantName` | * | Naam van de SharePoint-tenant, bijv. `contoso` voor `https://contoso-admin.sharepoint.com` |
| `-AdminUrl` | * | Volledige URL van het SharePoint-beheercentrum (alternatief voor `-TenantName`) |
| `-IncludeExternalUsers` | Nee | Ook externe gebruikers over alle sites opsommen (trager) |
| `-IncludeSiteOverrides` | Nee | Ook sites markeren die ruimer delen dan de standaard |
| `-OutputPath` | Nee | Map voor de CSV-rapport(en) (standaard: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Nee | Wordt doorgegeven aan `Connect-SPOService` als dat wordt ondersteund |

*Verplicht, tenzij er al verbinding is via `Connect-SPOService`.

**Voorbeelden**

```powershell
.\Test-SharePointSharingConfig.ps1 -TenantName "contoso"

.\Test-SharePointSharingConfig.ps1 -TenantName "contoso" -IncludeExternalUsers -IncludeSiteOverrides
```

**Opmerkingen**
- Voor rapportage over SharePoint-opslag/-versies, zie
  [`Reporting/Get-SharePointStorageReport.ps1`](../../Reporting/readme.nl.md) — een aparte,
  op Graph gebaseerde functie die al in deze repo zit

**Vereiste module**
```powershell
Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser
```
