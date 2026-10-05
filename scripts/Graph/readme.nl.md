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
| `-TenantId` | Ja | — | Tenant-ID van Entra ID |
| `-LogicAppName` | Ja | — | Weergavenaam van de managed identity / enterprise app van de Logic App |
| `-PermissionValue` | Nee | `AuditLog.Read.All` | Waarde van de Graph-app role die wordt toegewezen |
| `-ResourceAppId` | Nee | `00000003-0000-0000-c000-000000000000` (Microsoft Graph) | App-ID van de resource die de app role aanbiedt |
| `-ModuleHandling` | Nee | `InstallOrUpdate` | `InstallOrUpdate`, `InstallIfMissing` of `Skip` — hoe met de vereiste Graph-modules wordt omgegaan |

**Voorbeelden**

```powershell
# Ken de standaardmachtiging AuditLog.Read.All toe
.\logic-permissies.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -LogicAppName "MyLogicApp"

# Ken een specifieke machtiging toe
.\logic-permissies.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -LogicAppName "MyLogicApp" -PermissionValue "User.Read.All"

# Proefdraai
.\logic-permissies.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -LogicAppName "MyLogicApp" -WhatIf
```

**Vereisten**

- Global Administrator of Privileged Role Administrator (verleent applicatiemachtigingen)
- De managed identity van de Logic App moet al als enterprise application in de tenant bestaan
- Modules: `Microsoft.Graph.Authentication`, `Microsoft.Graph.Applications` (automatisch geïnstalleerd volgens `-ModuleHandling`)
- Vereiste gedelegeerde scopes: `Application.Read.All`, `AppRoleAssignment.ReadWrite.All`

Ondersteunt `-WhatIf` (`SupportsShouldProcess`).
