[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [TenantOnboarding](../readme.fr.md) › **Provisioning**

# Provisioning

Scripts d'amorçage d'un tenant unique nouvellement intégré : compte administrateur break-glass, groupes de sécurité de base et attribution de la stratégie de base Intune. Remplace un ancien script d'installation interactif piloté par menu par des scripts unitaires et paramétrés, cohérents avec le reste de ce dépôt. Exécutez-les dans l'ordre ci-dessous, dans le cadre d'une liste de contrôle pour nouveau tenant.

**La connexion** (pour les trois scripts) passe par [`Connect-M365.ps1`](../../Startup/Connect-M365.ps1) : **déléguée par défaut**, vous vous connectez en tant qu'administrateur du tenant (code d'appareil si `$global:useDeviceCodeAuth` est défini ; en GDAP, le tenant client vient de `$global:cid`, sauf si `-TenantId` en désigne un). **L'application seule** est une option avec `-ClientId` + `-CertificateThumbprint`, ou `-AppOnly` pour les lire dans `graph.appid.json`. Une session Graph existante n'est réutilisée que si elle porte sur le bon tenant et dispose déjà des scopes nécessaires au script ; le script ne ferme que la session qu'il a lui-même ouverte.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`New-BreakGlassAdminAccount.ps1`](New-BreakGlassAdminAccount.ps1) ([docs](#new-breakglassadminaccountps1)) | Créer un compte Global Administrator d'accès d'urgence, uniquement dans le cloud |
| [`New-TenantBaselineGroups.ps1`](New-TenantBaselineGroups.ps1) ([docs](#new-tenantbaselinegroupsps1)) | Créer le groupe d'exclusion CA et les groupes d'activation de fonctionnalités dont un nouveau tenant a besoin |
| [`Set-IntuneBaselinePolicyAssignment.ps1`](Set-IntuneBaselinePolicyAssignment.ps1) ([docs](#set-intunebaselinepolicyassignmentps1)) | Attribuer en masse à un groupe cible les stratégies de base Intune filtrées par nom |

**Ordre recommandé pour un nouveau tenant :**
1. `New-BreakGlassAdminAccount.ps1` : créer le compte d'accès d'urgence
2. `New-TenantBaselineGroups.ps1` : créer le groupe « tous les utilisateurs sauf break glass » (sa règle dynamique écarte le modèle d'UPN du compte break-glass) et les éventuels groupes d'activation de fonctionnalités
3. Uniquement si vos stratégies Conditional Access excluent un groupe *statique* : ajoutez-y le compte break-glass avec `New-BreakGlassAdminAccount.ps1 -ExcludeFromGroupId <group id>` lors de la création, ou plus tard avec `Add-UserToFeatureGroup.ps1` depuis `../UserManagement/`. Le groupe dynamique de l'étape 2 n'a pas besoin de membres ajoutés (et n'en accepte pas).
4. Importer/configurer vos stratégies de base Conditional Access et Intune (par ex. `Import-ConditionalAccessBaseline.ps1` dans `scripts/Entra/`)
5. `Set-IntuneBaselinePolicyAssignment.ps1` : attribuer les stratégies de base Intune au groupe « tous les utilisateurs sauf break glass »

---

### New-BreakGlassAdminAccount.ps1

Crée un utilisateur Entra ID uniquement cloud comme compte d'accès d'urgence (« break-glass »), conformément aux recommandations documentées de Microsoft : compte dédié uniquement cloud, long mot de passe aléatoire (affiché une seule fois, jamais enregistré sur disque), rôle Global Administrator attribué directement. Essai à blanc par défaut.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-UserPrincipalName` | Oui | UPN du nouveau compte ; utilisez le domaine `*.onmicrosoft.com` du tenant |
| `-DisplayName` | Non | Nom d'affichage (par défaut : `Break Glass Admin`) |
| `-PasswordLength` | Non | Longueur du mot de passe généré (par défaut : `24`) |
| `-AssignGlobalAdmin` | Non | Attribuer Global Administrator (par défaut : activé) |
| `-ExcludeFromGroupId` | Non | Un groupe d'exclusion CA statique auquel ajouter le compte (pas le groupe dynamique de `New-TenantBaselineGroups.ps1`) |
| `-TenantId` | Non | ID ou domaine du tenant Entra ID (par défaut : client GDAP, sinon le tenant de connexion) |
| `-ClientId` / `-CertificateThumbprint` | Non | Connexion en application seule |
| `-AppOnly` | Non | Application seule avec le ClientId et l'empreinte de `graph.appid.json` |
| `-Apply` | Non | Créer réellement le compte (par défaut : aperçu uniquement) |

**Exemples**

```powershell
.\New-BreakGlassAdminAccount.ps1 -UserPrincipalName "breakglass-admin@contoso.onmicrosoft.com"
.\New-BreakGlassAdminAccount.ps1 -UserPrincipalName "breakglass-admin@contoso.onmicrosoft.com" -Apply
```

**Remarques**
- Scopes : `User.ReadWrite.All`, `RoleManagement.ReadWrite.Directory`, `GroupMember.ReadWrite.All`.
- Global Administrator est attribué par une attribution de rôle unifiée (`roleManagement/directory/roleAssignments`, rôle `62e90394-69f5-4237-9190-012177145e10`). L'ancien code tentait d'activer le rôle avec `New-MgDirectoryRoleTemplate -RoleTemplateId` ; cette cmdlet crée un modèle de rôle et n'a pas de paramètre `-RoleTemplateId` : un tenant dans lequel le rôle n'avait jamais été activé échouait à cette étape.

**Modules requis**
```powershell
Install-Module Microsoft.Graph.Users -Scope CurrentUser
Install-Module Microsoft.Graph.Identity.Governance -Scope CurrentUser
Install-Module Microsoft.Graph.Groups -Scope CurrentUser
```

---

### New-TenantBaselineGroups.ps1

Crée un groupe d'exclusion dynamique « tous les utilisateurs sauf les comptes break-glass », ainsi qu'autant de groupes statiques d'activation de fonctionnalités que vous en nommez, pour délimiter les attributions Conditional Access et Intune dans un tenant nouvellement intégré.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-BreakGlassUpnPattern` | * | Modèle comparé aux UPN pour construire la règle dynamique du groupe d'exclusion |
| `-ExclusionGroupName` | Non | Nom d'affichage du groupe d'exclusion |
| `-SkipExclusionGroup` | Non | Ignorer entièrement le groupe d'exclusion |
| `-AdditionalGroupNames` | Non | Noms des groupes statiques supplémentaires à créer |
| `-TenantId` | Non | ID ou domaine du tenant Entra ID (par défaut : client GDAP, sinon le tenant de connexion) |
| `-ClientId` / `-CertificateThumbprint` | Non | Connexion en application seule |
| `-AppOnly` | Non | Application seule avec le ClientId et l'empreinte de `graph.appid.json` |
| `-Apply` | Non | Créer réellement les groupes (par défaut : aperçu uniquement) |

*Obligatoire, sauf si `-SkipExclusionGroup` est utilisé.

**Exemples**

```powershell
.\New-TenantBaselineGroups.ps1 -BreakGlassUpnPattern "breakglass-admin" `
    -AdditionalGroupNames "SG - Enable Password Manager","SG - Enable Windows 365"

.\New-TenantBaselineGroups.ps1 -BreakGlassUpnPattern "breakglass-admin" -Apply
```

**Remarques**
- Scope : `Group.ReadWrite.All`. L'appartenance dynamique nécessite Entra ID P1.

**Module requis**
```powershell
Install-Module Microsoft.Graph.Groups -Scope CurrentUser
```

---

### Set-IntuneBaselinePolicyAssignment.ps1

Attribue en masse à un seul groupe cible les profils de configuration, stratégies de conformité, modèles d'administration, scripts de plateforme et références de sécurité Intune dont le nom d'affichage correspond à un filtre. Le groupe est **ajouté** aux attributions existantes de chaque stratégie.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-TargetGroupId` | Oui | Groupe auquel attribuer les stratégies correspondantes |
| `-NameFilter` | Non | Filtre générique sur le nom d'affichage de la stratégie (par défaut : `*Default*`) |
| `-PolicyTypes` | Non | Types d'objets à inclure (par défaut : tous) |
| `-TenantId` | Non | ID ou domaine du tenant Entra ID (par défaut : client GDAP, sinon le tenant de connexion) |
| `-ClientId` / `-CertificateThumbprint` | Non | Connexion en application seule |
| `-AppOnly` | Non | Application seule avec le ClientId et l'empreinte de `graph.appid.json` |
| `-Apply` | Non | Créer réellement les attributions (par défaut : aperçu uniquement) |

**Exemples**

```powershell
.\Set-IntuneBaselinePolicyAssignment.ps1 -TargetGroupId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
.\Set-IntuneBaselinePolicyAssignment.ps1 -TargetGroupId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -Apply
```

**Module requis**
```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
```

**Remarques**
- Utilise le point de terminaison beta de Microsoft Graph : l'attribution des stratégies Intune n'est pas entièrement exposée en v1.0 pour tous les types de stratégies.
- L'action `/assign` d'Intune remplace toute la liste d'attributions d'une stratégie. Le script lit désormais les attributions actuelles et les renvoie avec le nouveau groupe ; auparavant, chaque stratégie traitée perdait ses autres attributions. Les stratégies déjà attribuées au groupe sont ignorées.
- Les listes de stratégies suivent `@odata.nextLink`, de sorte que les tenants comptant plus de stratégies qu'une page sont entièrement couverts.
- Les stratégies du catalogue de paramètres (`configurationPolicies`) ne sont pas incluses.
- Scopes : `DeviceManagementConfiguration.ReadWrite.All`, `DeviceManagementServiceConfig.ReadWrite.All`.
