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
| `-TenantId` | Non | client GDAP / votre tenant | ID ou domaine du tenant Entra ID (était obligatoire) |
| `-LogicAppName` | Oui | — | Nom d'affichage de l'identité managée / de l'application d'entreprise de la Logic App |
| `-PermissionValue` | Non | `AuditLog.Read.All` | Valeur du rôle d'application Graph à attribuer |
| `-ResourceAppId` | Non | `00000003-0000-0000-c000-000000000000` (Microsoft Graph) | App ID de la ressource qui expose le rôle d'application |
| `-ModuleHandling` | Non | `Skip` | `Skip` laisse les modules à `load.ps1` ; `InstallIfMissing` installe les modules absents. `InstallOrUpdate` est encore accepté mais n'exécute plus `Update-Module -Force` — il se comporte comme `InstallIfMissing` |
| `-ClientId` / `-CertificateThumbprint` | Non | — | Connexion app-only avec cette inscription d'application et ce certificat |
| `-AppOnly` | Non | — | Connexion app-only avec l'application de `graph.appid.json` |

**Exemples**

```powershell
# Accorder l'autorisation par défaut AuditLog.Read.All
.\logic-permissies.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -LogicAppName "MyLogicApp"

# Accorder une autorisation précise
.\logic-permissies.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -LogicAppName "MyLogicApp" -PermissionValue "User.Read.All"

# Essai à blanc
.\logic-permissies.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -LogicAppName "MyLogicApp" -WhatIf

# Sous GDAP, le tenant client est la valeur par défaut, -TenantId peut donc être omis
.\logic-permissies.ps1 -LogicAppName "MyLogicApp"
```

**Prérequis**

- Global Administrator ou Privileged Role Administrator (accorde des autorisations d'application)
- L'identité managée de la Logic App doit déjà exister en tant qu'application d'entreprise dans le tenant
- Modules : `Microsoft.Graph.Authentication`, `Microsoft.Graph.Applications` (installés et mis à jour par `load.ps1` ; `-ModuleHandling InstallIfMissing` les installe s'ils sont absents)
- Étendues déléguées requises : `Application.Read.All`, `AppRoleAssignment.ReadWrite.All`

**Connexion**

Via [`Connect-M365.ps1`](../Startup/readme.fr.md) : déléguée en tant qu'administrateur par défaut (navigateur, ou code d'appareil / client GDAP selon `load.config.ps1`), app-only avec `-ClientId` + `-CertificateThumbprint` ou `-AppOnly`. Une session Graph pour le bon tenant qui possède déjà les étendues est réutilisée et reste connectée ; seule une session ouverte par le script est fermée.

Prend en charge `-WhatIf` (`SupportsShouldProcess`).
