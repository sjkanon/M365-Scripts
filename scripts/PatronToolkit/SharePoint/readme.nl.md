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

Rapporteert de tenantbrede instellingen voor extern delen die er bij een beveiligingsreview het meest
toe doen, gelezen uit Microsoft Graph (`GET /admin/sharepoint/settings`): deelmogelijkheden,
opnieuw delen door externe gebruikers, of het accepterende account moet overeenkomen met het
uitgenodigde, de modus voor toegestane/geblokkeerde domeinen, verouderde authenticatie en
afmelden bij inactiviteit. Optioneel, via PnP.PowerShell, ook het standaard linktype / de
linkrechten / het verloop van anonieme links, sites die ruimer delen dan de standaard van de
tenant, en elke externe gebruiker (gast) over alle sitecollecties.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-TenantName` | Nee | Naam van de SharePoint-tenant, bijv. `contoso` voor `https://contoso-admin.sharepoint.com`. Alleen voor het PnP-deel; zet ook `-IncludeLinkSettings` aan |
| `-AdminUrl` | Nee | Volledige URL van het SharePoint-beheercentrum (alternatief voor `-TenantName`). Zonder beide wordt de URL via Graph opgezocht (`/sites/root`) |
| `-IncludeLinkSettings` | Nee | Ook standaard linktype, standaard linkrechten en verloop van anonieme links rapporteren (PnP) |
| `-IncludeExternalUsers` | Nee | Ook externe gebruikers over alle sites opsommen (PnP, trager) |
| `-IncludeSiteOverrides` | Nee | Ook sites markeren die ruimer delen dan de standaard (PnP) |
| `-OutputPath` | Nee | Map voor de CSV-rapport(en) (standaard: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Nee | Tenant-ID of domein. Standaard de GDAP-klant (`load.config.ps1`) of je eigen tenant; verplicht voor app-only |
| `-ClientId` | Nee | App-registratie voor app-only aanmelden (met `-CertificateThumbprint`), voor Graph en PnP. Zonder meldt het script gedelegeerd aan, als jezelf |
| `-CertificateThumbprint` | Nee | Certificaatvingerafdruk voor app-only aanmelden met `-ClientId` |
| `-AppOnly` | Nee | App-only aanmelden (Graph en PnP) met de ClientId en vingerafdruk voor de tenant uit `graph.appid.json` |
| `-PnPClientId` | Nee | PnP-app-registratie voor het gedelegeerd aanmelden bij PnP (standaard: de vermelding van de tenant in `pnp.appid.json`) |

**Voorbeelden**

```powershell
# Alleen Graph: deelinstellingen van de tenant
.\Test-SharePointSharingConfig.ps1

# Zoals voorheen: tenantinstellingen plus linkinstellingen (PnP)
.\Test-SharePointSharingConfig.ps1 -TenantName "contoso"

.\Test-SharePointSharingConfig.ps1 -IncludeLinkSettings -IncludeExternalUsers -IncludeSiteOverrides

# App-only, Graph en PnP met de app uit graph.appid.json
.\Test-SharePointSharingConfig.ps1 -TenantId contoso.onmicrosoft.com -AppOnly -IncludeSiteOverrides
```

**Opmerkingen**
- De verouderde SharePoint Online Management Shell (`Connect-SPOService`, `Get-SPOTenant`,
  `Get-SPOSite`, `Get-SPOExternalUser`) wordt niet meer gebruikt. De eerdere `-TenantId` was
  gedocumenteerd maar werd nooit doorgegeven
- Graph heeft geen API voor de standaardinstellingen van links, de deelmogelijkheden per site
  of externe gebruikers per site, dus die blijven op PnP (`Get-PnPTenant`,
  `Get-PnPTenantSite`, `Get-PnPExternalUser`) en draaien alleen als erom gevraagd wordt.
  Mislukt het aanmelden bij PnP, dan worden die delen met een waarschuwing overgeslagen en
  wordt het Graph-deel toch gerapporteerd
- Graph geeft `sharingCapability` in camelCase (`externalUserAndGuestSharing`) en het
  omgekeerde van de oude instelling: `ResharingByExternalUsersEnabled` in plaats van
  `PreventExternalUsersFromResharing`
- Aanmelden via [`Connect-M365.ps1`](../../Startup/readme.nl.md#connect-m365ps1): Graph
  standaard gedelegeerd (scope `SharePointTenantSettings.Read.All`, plus `Sites.Read.All`
  alleen om de beheer-URL op te zoeken; rol SharePoint Administrator); app-only met
  `-ClientId` + `-CertificateThumbprint` of `-AppOnly`. PnP heeft sinds september 2024 een
  eigen app-registratie nodig: `-PnPClientId` of `pnp.appid.json` voor gedelegeerd, dezelfde
  app als Graph voor app-only (SharePoint `Sites.FullControl.All`)
- Voor rapportage over SharePoint-opslag/-versies, zie
  [`Reporting/Get-SharePointStorageReport.ps1`](../../Reporting/readme.nl.md) — een aparte,
  op Graph gebaseerde functie die al in deze repo zit

**Vereiste modules**
```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module PnP.PowerShell -Scope CurrentUser   # alleen voor het PnP-deel
```
