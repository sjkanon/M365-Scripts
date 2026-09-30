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
(of "All users"/"All devices") en geeft aan of de toewijzing Include of Exclude is. Eén CSV-
rij per combinatie van beleid en toewijzing.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-PolicyType` | Nee | `DeviceConfiguration`, `CompliancePolicy`, `SettingsCatalog`, `MobileApp` (standaard: alle) |
| `-OutputPath` | Nee | Pad voor het CSV-rapport (standaard: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Nee | Tenant-ID of domein van Entra ID |

**Voorbeelden**

```powershell
.\Get-IntunePolicyAssignments.ps1

.\Get-IntunePolicyAssignments.ps1 -PolicyType CompliancePolicy,SettingsCatalog
```

**Opmerkingen**
- Alleen-lezen. Bulkwijzigingen van toewijzingen zijn bewust niet in een script gegoten — beoordeel de uitvoer
  van dit rapport en wijzig toewijzingen per beleid in het Intune-beheercentrum
- Vereiste scopes: `DeviceManagementConfiguration.Read.All`,
  `DeviceManagementApps.Read.All`, `Group.Read.All`

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
| `-TenantId` | Nee | Tenant-ID of domein van Entra ID |

**Voorbeelden**

```powershell
.\Get-AutopilotDevices.ps1

.\Get-AutopilotDevices.ps1 -GroupTag "Finance-Laptops"
```

**Opmerkingen**
- Alleen-lezen. Bulkimport/-toewijzing/-verwijdering van apparaten is bewust niet in een script gegoten — gebruik
  `Get-WindowsAutoPilotInfo.ps1` (al aanwezig in deze repo) plus het Intune-beheercentrum voor
  registratie, en beoordeel dit rapport vóór een eventuele bulkhertoewijzing
- Vereiste scope: `DeviceManagementServiceConfig.Read.All`

**Vereiste modules**
```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```
