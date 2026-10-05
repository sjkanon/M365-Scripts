[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../readme.fr.md) › [scripts](../readme.fr.md) › **Graph**

# Graph

Scripts de gestion des autorisations d'application Microsoft Graph et des principaux de service.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`logic-permissies.ps1`](logic-permissies.ps1) ([docs](#logic-permissiesps1)) | Accorder une autorisation d'application Microsoft Graph (rôle d'application) à l'identité managée d'une Azure Logic App — idempotent |

---

### logic-permissies.ps1

Accorde une autorisation d'application Microsoft Graph (rôle d'application) à l'identité managée d'une Azure Logic App — autrement dit, attribue un rôle d'application de type autorisation d'application au principal de service de la Logic App. Idempotent : ignore l'opération si l'attribution existe déjà.

**Paramètres**

| Paramètre | Obligatoire | Valeur par défaut | Description |
|-----------|----------|---------|-------------|
| `-TenantId` | Oui | — | ID du tenant Entra ID |
| `-LogicAppName` | Oui | — | Nom d'affichage de l'identité managée / de l'application d'entreprise de la Logic App |
| `-PermissionValue` | Non | `AuditLog.Read.All` | Valeur du rôle d'application Graph à attribuer |
| `-ResourceAppId` | Non | `00000003-0000-0000-c000-000000000000` (Microsoft Graph) | App ID de la ressource qui expose le rôle d'application |
| `-ModuleHandling` | Non | `InstallOrUpdate` | `InstallOrUpdate`, `InstallIfMissing` ou `Skip` — comment gérer les modules Graph requis |

**Exemples**

```powershell
# Accorder l'autorisation par défaut AuditLog.Read.All
.\logic-permissies.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -LogicAppName "MyLogicApp"

# Accorder une autorisation précise
.\logic-permissies.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -LogicAppName "MyLogicApp" -PermissionValue "User.Read.All"

# Essai à blanc
.\logic-permissies.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -LogicAppName "MyLogicApp" -WhatIf
```

**Prérequis**

- Global Administrator ou Privileged Role Administrator (accorde des autorisations d'application)
- L'identité managée de la Logic App doit déjà exister en tant qu'application d'entreprise dans le tenant
- Modules : `Microsoft.Graph.Authentication`, `Microsoft.Graph.Applications` (installés automatiquement selon `-ModuleHandling`)
- Étendues déléguées requises : `Application.Read.All`, `AppRoleAssignment.ReadWrite.All`

Prend en charge `-WhatIf` (`SupportsShouldProcess`).
