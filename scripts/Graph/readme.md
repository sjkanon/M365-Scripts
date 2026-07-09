# Graph

Scripts for managing Microsoft Graph application permissions and service principals.

---

## Scripts

### logic-permissies.ps1

Grants a Microsoft Graph application permission (app role) to an Azure Logic App's managed identity — i.e. assigns an application-permission app role to the Logic App's service principal. Idempotent: skips if the assignment already exists.

**Parameters**

| Parameter | Required | Default | Description |
|-----------|----------|---------|-------------|
| `-TenantId` | Yes | — | Entra ID tenant ID |
| `-LogicAppName` | Yes | — | Display name of the Logic App's managed identity / enterprise app |
| `-PermissionValue` | No | `AuditLog.Read.All` | Graph app role value to assign |
| `-ResourceAppId` | No | `00000003-0000-0000-c000-000000000000` (Microsoft Graph) | App ID of the resource exposing the app role |
| `-ModuleHandling` | No | `InstallOrUpdate` | `InstallOrUpdate`, `InstallIfMissing`, or `Skip` — how to handle required Graph modules |

**Examples**

```powershell
# Grant the default AuditLog.Read.All permission
.\logic-permissies.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -LogicAppName "MyLogicApp"

# Grant a specific permission
.\logic-permissies.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -LogicAppName "MyLogicApp" -PermissionValue "User.Read.All"

# Dry run
.\logic-permissies.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -LogicAppName "MyLogicApp" -WhatIf
```

**Requirements**

- Global Administrator or Privileged Role Administrator (grants application permissions)
- The Logic App's managed identity must already exist as an enterprise application in the tenant
- Modules: `Microsoft.Graph.Authentication`, `Microsoft.Graph.Applications` (auto-installed per `-ModuleHandling`)
- Required delegated scopes: `Application.Read.All`, `AppRoleAssignment.ReadWrite.All`

Supports `-WhatIf` (`SupportsShouldProcess`).
