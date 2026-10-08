[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [PatronToolkit](../readme.nl.md) › **Intune**

# Patron Toolkit — Intune

Rapportage van Intune-beleidstoewijzingen en inventaris van Windows Autopilot-apparaten via Microsoft
Graph.

---

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`Get-IntunePolicyAssignments.ps1`](Get-IntunePolicyAssignments.ps1) ([docs](#get-intunepolicyassignmentsps1)) | Rapport van welke groepen aan welke Intune-profielen/-beleidsregels/-apps zijn toegewezen |
| [`Get-AutopilotDevices.ps1`](Get-AutopilotDevices.ps1) ([docs](#get-autopilotdevicesps1)) | Rapport van geregistreerde Windows Autopilot-apparaten en implementatieprofielen |

---

### Get-IntunePolicyAssignments.ps1

Somt de apparaatconfiguratieprofielen, het nalevingsbeleid, het beleid uit de Settings Catalog
en de mobiele apps in Intune op, zet de toewijzingen van elk object om naar leesbare groepsnamen
(of "All users"/"All devices") en geeft aan of de toewijzing Include of Exclude is; rijen voor
apps tonen ook de intentie (required/available/uninstall). Eén CSV-rij per combinatie van
beleid en toewijzing.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-PolicyType` | Nee | `DeviceConfiguration`, `CompliancePolicy`, `SettingsCatalog`, `MobileApp` (standaard: alle) |
| `-OutputPath` | Nee | Pad voor het CSV-rapport (standaard: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Nee | Tenant-ID of domein. Standaard de GDAP-klant (`load.config.ps1`) of je eigen tenant; verplicht voor app-only |
| `-ClientId` | Nee | App-registratie voor app-only aanmelden (met `-CertificateThumbprint`). Zonder meldt het script gedelegeerd aan, als jezelf |
| `-CertificateThumbprint` | Nee | Certificaatvingerafdruk voor app-only aanmelden met `-ClientId` |
| `-AppOnly` | Nee | App-only aanmelden met de ClientId en vingerafdruk voor de tenant uit `graph.appid.json` |

**Voorbeelden**

```powershell
.\Get-IntunePolicyAssignments.ps1

.\Get-IntunePolicyAssignments.ps1 -PolicyType CompliancePolicy,SettingsCatalog
```

**Opmerkingen**
- Alleen-lezen. Bulkwijzigingen van toewijzingen zijn bewust niet in een script gegoten — beoordeel de uitvoer
  van dit rapport en wijzig toewijzingen per beleid in het Intune-beheercentrum
- Alle aanroepen zijn `Invoke-MgGraphRequest` met `$expand=assignments` en paginering via
  `@odata.nextLink`. Beleid uit de Settings Catalog bestaat alleen in Graph **beta**
  (`/beta/deviceManagement/configurationPolicies`); de vorige versie riep
  `Get-MgDeviceManagementConfigurationPolicy` aan, die de Microsoft.Graph v2 SDK niet heeft,
  en meldde stilzwijgend geen enkel Settings Catalog-beleid. Een type dat niet gelezen kan
  worden, geeft nu een zichtbare waarschuwing in plaats van een leeg resultaat
- Aanmelden via [`Connect-M365.ps1`](../../Startup/readme.nl.md#connect-m365ps1): Microsoft
  Graph, standaard gedelegeerd (scopes `DeviceManagementConfiguration.Read.All`,
  `DeviceManagementApps.Read.All`, `Group.Read.All` plus een Intune-rol); app-only met
  `-ClientId` + `-CertificateThumbprint` of `-AppOnly` (dezelfde toepassingsmachtigingen)

---

### Get-AutopilotDevices.ps1

Toont elke Windows Autopilot-apparaatidentiteit die in de tenant is geregistreerd (serienummer,
model, fabrikant, group tag, inschrijvingsstatus, toewijzingsstatus van het implementatieprofiel),
plus een overzicht van de bestaande implementatieprofielen. Dit is een inventarisrapport
aan de tenantkant — iets anders dan
[`Get-Autopilot/Get-WindowsAutoPilotInfo.ps1`](../../Intune/readme.nl.md), dat de
hardwarehash van een fysiek apparaat verzamelt om het te registreren.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-GroupTag` | Nee | Alleen apparaten met deze group tag rapporteren |
| `-OutputPath` | Nee | Pad voor het CSV-rapport (standaard: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Nee | Tenant-ID of domein. Standaard de GDAP-klant (`load.config.ps1`) of je eigen tenant; verplicht voor app-only |
| `-ClientId` | Nee | App-registratie voor app-only aanmelden (met `-CertificateThumbprint`). Zonder meldt het script gedelegeerd aan, als jezelf |
| `-CertificateThumbprint` | Nee | Certificaatvingerafdruk voor app-only aanmelden met `-ClientId` |
| `-AppOnly` | Nee | App-only aanmelden met de ClientId en vingerafdruk voor de tenant uit `graph.appid.json` |

**Voorbeelden**

```powershell
.\Get-AutopilotDevices.ps1

.\Get-AutopilotDevices.ps1 -GroupTag "Finance-Laptops"
```

**Opmerkingen**
- Alleen-lezen. Bulkimport/-toewijzing/-verwijdering van apparaten is bewust niet in een script gegoten — gebruik
  `Get-WindowsAutoPilotInfo.ps1` (al aanwezig in deze repo) plus het Intune-beheercentrum voor
  registratie, en beoordeel dit rapport vóór een eventuele bulkhertoewijzing
- Implementatieprofielen bestaan alleen in Graph **beta**
  (`/beta/deviceManagement/windowsAutopilotDeploymentProfiles`) en worden gelezen met
  `Invoke-MgGraphRequest`; de vorige versie gebruikte een cmdlet die de Microsoft.Graph v2
  SDK niet heeft en toonde altijd nul profielen. Apparaten komen uit Graph v1.0
- De telling "zonder toegewezen implementatieprofiel" rekent `assignedInSync`,
  `assignedOutOfSync` en `assignedUnkownSyncState` nu als toegewezen — Graph kent geen
  kale waarde `assigned`, dus elk apparaat telde voorheen als niet-toegewezen
- Aanmelden via [`Connect-M365.ps1`](../../Startup/readme.nl.md#connect-m365ps1): Microsoft
  Graph, standaard gedelegeerd (scope `DeviceManagementServiceConfig.Read.All` plus een
  Intune-rol); app-only met `-ClientId` + `-CertificateThumbprint` of `-AppOnly`

**Vereiste modules**
```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```
