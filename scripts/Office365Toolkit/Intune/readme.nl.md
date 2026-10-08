[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [Office365Toolkit](../readme.nl.md) › **Intune**

# Office365Toolkit / Intune

Tenantbrede inventaris van Intune / Endpoint Manager-beleid.

---

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`Get-IntunePolicyInventory.ps1`](Get-IntunePolicyInventory.ps1) ([docs](#get-intunepolicyinventoryps1)) | Inventaris van al het beleid voor naleving/configuratie/app-beveiliging/Endpoint Security |

---

### Get-IntunePolicyInventory.ps1

Toont elk beleid op de belangrijkste Intune-beleidsonderdelen — nalevingsbeleid
voor apparaten, apparaatconfiguratieprofielen, configuratiebeleid uit de Settings Catalog,
app-beveiligingsbeleid en Endpoint Security-beleid ("intents") —
met het aantal toewijzingen. Een snelle momentopname als inventaris/checklist, geen
vergelijking met een basislijn. Alleen-lezen.

Meldt aan via [`Connect-M365.ps1`](../../Startup/readme.nl.md#connect-m365ps1): standaard
delegated (je meldt je aan als de beheerder), met device code en de GDAP-klant uit
`load.config.ps1`; app-only met `-ClientId` + `-CertificateThumbprint`, of `-AppOnly`
(`graph.appid.json`). Een Graph-sessie die al past wordt hergebruikt en blijft verbonden.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-OutputPath` | Nee | Bestandspad voor het CSV-rapport (standaard: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Nee | Tenant-ID of domein; standaard de GDAP-klant uit `load.config.ps1` |
| `-ClientId` | Nee | App-registratie voor app-only aanmelding (met `-CertificateThumbprint`) |
| `-CertificateThumbprint` | Nee | Certificaatvingerafdruk voor app-only aanmelding |
| `-AppOnly` | Nee | App-only met ClientId en vingerafdruk uit `graph.appid.json` |

**Voorbeelden**

```powershell
.\Get-IntunePolicyInventory.ps1

.\Get-IntunePolicyInventory.ps1 -OutputPath C:\Reports\intune.csv

.\Get-IntunePolicyInventory.ps1 -TenantId contoso.onmicrosoft.com -AppOnly
```

**Opmerkingen**
- Settings Catalog (`configurationPolicies`) en Endpoint Security (`intents`) bestaan alleen op
  het beta-endpoint van Graph. De v1.0-SDK heeft er geen cmdlets voor, dus eerdere versies vingen
  de fout af en lieten beide stilletjes weg. Alle vijf onderdelen worden nu gelezen met
  `Invoke-MgGraphRequest`, met paginering via `@odata.nextLink`.
- Beleid zonder toewijzingen toont nu `0` in plaats van een leeg aantal. App-beveiligingsbeleid
  heeft geen toewijzingen op het basistype `managedAppPolicy`, dus dat aantal blijft leeg.
- Om de Intune-configuratie van een klanttenant te vergelijken met een
  referentiebasislijn van de MSP (driftdetectie), gebruik je in plaats daarvan
  [`scripts/Intune/Compare-IntuneConfig.ps1`](../../Intune/readme.nl.md#compare-intuneconfigps1)
  — dat script maakt een volledige vergelijking op basis van een back-up; dit script is een snelle
  inventaris van wat er op dit moment bestaat.

**Vereiste scopes:** `DeviceManagementConfiguration.Read.All`, `DeviceManagementApps.Read.All`
**Vereiste module:** `Microsoft.Graph.Authentication`
