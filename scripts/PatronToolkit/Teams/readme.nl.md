[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [PatronToolkit](../readme.nl.md) › **Teams**

# Patron Toolkit — Teams

Governance van de Microsoft Teams-tenant en inventarisrapportage.

---

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`Get-TeamsConfigReport.ps1`](Get-TeamsConfigReport.ps1) ([docs](#get-teamsconfigreportps1)) | Rapport van de governance-instellingen van de Teams-tenant en een inventaris van teams |

---

### Get-TeamsConfigReport.ps1

Rapporteert het tenantbrede Teams-beleid dat er bij een governance-/beveiligings-
review het meest toe doet: externe toegang (federatie), gasttoegang, globaal vergaderbeleid (anoniem
deelnemen, opnemen, presentatorrol), globaal berichtenbeleid en app-instellingsbeleid
(sideloading). Toont ook elk team met zichtbaarheid, archiefstatus en aantallen eigenaren/leden
— en markeert elk team zonder eigenaren.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-IncludeTeamsInventory` | Nee | Ook elk team tonen met het aantal leden/eigenaren (standaard: aan) |
| `-OutputPath` | Nee | Map voor de CSV-rapport(en) (standaard: `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Nee | Tenant-ID of domein van Entra ID |

**Voorbeelden**

```powershell
.\Get-TeamsConfigReport.ps1

.\Get-TeamsConfigReport.ps1 -IncludeTeamsInventory:$false
```

**Opmerkingen**
- Alleen-lezen

**Vereiste module**
```powershell
Install-Module MicrosoftTeams -Scope CurrentUser
```
