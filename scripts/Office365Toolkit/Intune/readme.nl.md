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

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-OutputPath` | Nee | Pad voor het CSV-rapport (standaard: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Nee | Tenant-ID of domein van Entra ID |

**Voorbeelden**

```powershell
.\Get-IntunePolicyInventory.ps1

.\Get-IntunePolicyInventory.ps1 -OutputPath C:\Reports
```

**Opmerkingen**
- Om de Intune-configuratie van een klanttenant te vergelijken met een
  referentiebasislijn van de MSP (driftdetectie), gebruik je in plaats daarvan
  [`scripts/Intune/Compare-IntuneConfig.ps1`](../../Intune/readme.nl.md#compare-intuneconfigps1)
  — dat script maakt een volledige vergelijking op basis van een back-up; dit script is een snelle
  inventaris van wat er op dit moment bestaat.

**Vereiste scopes:** `DeviceManagementConfiguration.Read.All`, `DeviceManagementApps.Read.All`
**Vereiste module:** `Microsoft.Graph.DeviceManagement`
