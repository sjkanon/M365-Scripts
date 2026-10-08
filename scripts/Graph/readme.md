**English** · [Nederlands](readme.nl.md) · [Français](readme.fr.md)

[M365-Scripts](../../readme.md) › [scripts](../readme.md) › **Graph**

# Graph

Scripts for managing Microsoft Graph application permissions and service principals.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`logic-permissies.ps1`](logic-permissies.ps1) ([docs](#logic-permissiesps1)) | Grant a Microsoft Graph application permission (app role) to an Azure Logic App's managed identity — idempotent |

---

### logic-permissies.ps1

Grants a Microsoft Graph application permission (app role) to an Azure Logic App's managed identity — i.e. assigns an application-permission app role to the Logic App's service principal. Idempotent: skips if the assignment already exists.

**Parameters**

| Parameter | Required | Default | Description |
|-----------|----------|---------|-------------|
| `-TenantId` | No | GDAP customer / your tenant | Entra ID tenant ID or domain (was mandatory) |
| `-LogicAppName` | Yes | — | Display name of the Logic App's managed identity / enterprise app |
| `-PermissionValue` | No | `AuditLog.Read.All` | Graph app role value to assign |
| `-ResourceAppId` | No | `00000003-0000-0000-c000-000000000000` (Microsoft Graph) | App ID of the resource exposing the app role |
| `-ModuleHandling` | No | `Skip` | `Skip` leaves modules to `load.ps1`; `InstallIfMissing` installs absent modules. `InstallOrUpdate` is still accepted but no longer runs `Update-Module -Force` — it behaves like `InstallIfMissing` |
| `-ClientId` / `-CertificateThumbprint` | No | — | App-only sign-in with this app registration and certificate |
| `-AppOnly` | No | — | App-only sign-in with the app from `graph.appid.json` |

**Examples**

```powershell
# Grant the default AuditLog.Read.All permission
.\logic-permissies.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -LogicAppName "MyLogicApp"

# Grant a specific permission
.\logic-permissies.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -LogicAppName "MyLogicApp" -PermissionValue "User.Read.All"

# Dry run
.\logic-permissies.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -LogicAppName "MyLogicApp" -WhatIf

# Under GDAP the customer tenant is the default, so -TenantId can be left out
.\logic-permissies.ps1 -LogicAppName "MyLogicApp"
```

**Requirements**

- Global Administrator or Privileged Role Administrator (grants application permissions)
- The Logic App's managed identity must already exist as an enterprise application in the tenant
- Modules: `Microsoft.Graph.Authentication`, `Microsoft.Graph.Applications` (installed and updated by `load.ps1`; `-ModuleHandling InstallIfMissing` installs them when absent)
- Required delegated scopes: `Application.Read.All`, `AppRoleAssignment.ReadWrite.All`

**Sign-in**

Through [`Connect-M365.ps1`](../Startup/readme.md): delegated as the admin by default (browser, or device code / GDAP customer per `load.config.ps1`), app-only with `-ClientId` + `-CertificateThumbprint` or `-AppOnly`. A Graph session for the right tenant that already has the scopes is reused and left connected; only a session the script opened is disconnected.

Supports `-WhatIf` (`SupportsShouldProcess`).
