[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../readme.nl.md) › [scripts](../readme.nl.md) › **Graph**

# Graph

Scripts voor het beheren van Microsoft Graph-applicatiemachtigingen en service principals.

---

## Scripts

| Script | Omschrijving |
|--------|--------------|
| [`logic-permissies.ps1`](logic-permissies.ps1) ([docs](#logic-permissiesps1)) | Een Microsoft Graph-applicatiemachtiging (app role) toekennen aan de managed identity van een Azure Logic App — idempotent |

---

### logic-permissies.ps1

Kent een Microsoft Graph-applicatiemachtiging (app role) toe aan de managed identity van een Azure Logic App — oftewel: wijst een app role met applicatiemachtiging toe aan de service principal van de Logic App. Idempotent: slaat over als de toewijzing al bestaat.

**Parameters**

| Parameter | Verplicht | Standaard | Omschrijving |
|-----------|----------|---------|-------------|
| `-TenantId` | Nee | GDAP-klant / je eigen tenant | Tenant-ID of -domein van Entra ID (was verplicht) |
| `-LogicAppName` | Ja | — | Weergavenaam van de managed identity / enterprise app van de Logic App |
| `-PermissionValue` | Nee | `AuditLog.Read.All` | Waarde van de Graph-app role die wordt toegewezen |
| `-ResourceAppId` | Nee | `00000003-0000-0000-c000-000000000000` (Microsoft Graph) | App-ID van de resource die de app role aanbiedt |
| `-ModuleHandling` | Nee | `Skip` | `Skip` laat de modules aan `load.ps1` over; `InstallIfMissing` installeert ontbrekende modules. `InstallOrUpdate` wordt nog geaccepteerd maar draait geen `Update-Module -Force` meer — het werkt als `InstallIfMissing` |
| `-ClientId` / `-CertificateThumbprint` | Nee | — | App-only aanmelden met deze app-registratie en dit certificaat |
| `-AppOnly` | Nee | — | App-only aanmelden met de app uit `graph.appid.json` |

**Voorbeelden**

```powershell
# Ken de standaardmachtiging AuditLog.Read.All toe
.\logic-permissies.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -LogicAppName "MyLogicApp"

# Ken een specifieke machtiging toe
.\logic-permissies.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -LogicAppName "MyLogicApp" -PermissionValue "User.Read.All"

# Proefdraai
.\logic-permissies.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -LogicAppName "MyLogicApp" -WhatIf

# Onder GDAP is de klanttenant de standaard, dus -TenantId kan weg
.\logic-permissies.ps1 -LogicAppName "MyLogicApp"
```

**Vereisten**

- Global Administrator of Privileged Role Administrator (verleent applicatiemachtigingen)
- De managed identity van de Logic App moet al als enterprise application in de tenant bestaan
- Modules: `Microsoft.Graph.Authentication`, `Microsoft.Graph.Applications` (geïnstalleerd en bijgewerkt door `load.ps1`; `-ModuleHandling InstallIfMissing` installeert ze als ze ontbreken)
- Vereiste gedelegeerde scopes: `Application.Read.All`, `AppRoleAssignment.ReadWrite.All`

**Aanmelden**

Via [`Connect-M365.ps1`](../Startup/readme.nl.md): standaard delegated als beheerder (browser, of apparaatcode / GDAP-klant volgens `load.config.ps1`), app-only met `-ClientId` + `-CertificateThumbprint` of `-AppOnly`. Een Graph-sessie voor de juiste tenant die de scopes al heeft, wordt hergebruikt en blijft verbonden; alleen een sessie die het script zelf opende, wordt verbroken.

Ondersteunt `-WhatIf` (`SupportsShouldProcess`).
