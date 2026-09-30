[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [LegacyUtilities](../readme.nl.md) › **Workspace365**

# Legacy Utilities — Workspace 365

Provisioningscripts voor digitale werkplekomgevingen van [Workspace 365](https://workspace365.net).
Gemoderniseerde herschrijving van een oud interactief script van meer dan 600 regels: de
device-code-aanmelding via MSAL.PS en de kale Graph-aanroepen met bearer-token zijn vervangen
door `Connect-MgGraph` / `Invoke-MgGraphRequest` (met hergebruik van een bestaande sessie als je
al verbonden bent), en de altijd-interactieve vragen zijn omgezet in parameters volgens het
patroon van deze repository: standaard een proefdraai, `-Apply` om echt uit te voeren. De
provisioningsleutel en hostnaam geef je altijd zelf mee; het origineel sloeg ze op in een lokaal
`.cfg`-bestand in platte tekst naast het script, en dat opslaan is bewust geschrapt.

---

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`New-Workspace365Environment.ps1`](New-Workspace365Environment.ps1) ([docs](#new-workspace365environmentps1)) | Richt een nieuwe omgeving in, met de bijbehorende SSO-app-registratie en standaardkoppelingen naar Exchange/SharePoint |
| [`Remove-Workspace365Environment.ps1`](Remove-Workspace365Environment.ps1) ([docs](#remove-workspace365environmentps1)) | Verwijder een omgeving via de Provisioning API |

---

### New-Workspace365Environment.ps1

Maakt een App Registration in Entra ID aan voor SSO met Workspace 365 (gedelegeerde scopes voor
Graph/Power BI + een client secret), richt de omgeving in via de Workspace 365 Provisioning API,
koppelt de app als SSO-identiteitsprovider van die omgeving, en laat de standaard-URL's voor
Exchange/SharePoint naar deze tenant wijzen. Standaard een proefdraai.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-WorkspaceHostname` | Ja | bijv. `https://yourcompany.workspace365.net` |
| `-ProvisioningKey` | Ja | Provisioningsleutel van Workspace 365 (GUID); nooit hardcoden, geef hem mee bij het aanroepen |
| `-EnvironmentName` | Ja | Omgevingsnaam in kleine letters en cijfers |
| `-RequestingUserUpn` | Nee | Standaard de aangemelde gebruiker |
| `-Apply` | Nee | Voer de provisioning echt uit (standaard: voorbeeldweergave) |
| `-TenantId` | Nee | Tenant-ID of domein van Entra ID |

```powershell
.\New-Workspace365Environment.ps1 -WorkspaceHostname "https://yourcompany.workspace365.net" -ProvisioningKey $key -EnvironmentName "contoso" -Apply
```

---

### Remove-Workspace365Environment.ps1

Verwijdert een omgeving via de Provisioning API. Verwijdert de bijbehorende App Registration
**niet**; ruim die apart op (Entra-beheercentrum of `Remove-MgApplication`) als je hem niet meer
nodig hebt. Standaard een proefdraai.

```powershell
.\Remove-Workspace365Environment.ps1 -WorkspaceHostname "https://yourcompany.workspace365.net" -ProvisioningKey $key -EnvironmentName "contoso" -Apply
```

---

**Vereiste modules**

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

**Opmerkingen**
- Door deze scripts te gebruiken ga je akkoord met de
  [algemene voorwaarden van Workspace 365](https://workspace365.net/en/term-and-conditions).
