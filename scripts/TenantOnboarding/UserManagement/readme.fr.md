[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [TenantOnboarding](../readme.fr.md) › **UserManagement**

# UserManagement

Petits utilitaires de gestion des utilisateurs et des groupes Entra ID / Exchange Online, utilisés lors de l'intégration des tenants et du déploiement continu de fonctionnalités.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`New-DynamicDistributionGroupByFilter.ps1`](New-DynamicDistributionGroupByFilter.ps1) ([docs](#new-dynamicdistributiongroupbyfilterps1)) | Créer un groupe de distribution dynamique à partir d'un intitulé de poste ou d'un filtre de destinataires personnalisé |
| [`Add-UserToFeatureGroup.ps1`](Add-UserToFeatureGroup.ps1) ([docs](#add-usertofeaturegroupps1)) | Ajouter un utilisateur à un groupe Entra ID nommé, ou l'en retirer, pour contrôler l'accès à une fonctionnalité complémentaire |

---

### New-DynamicDistributionGroupByFilter.ps1

Affiche d'abord les destinataires correspondant à un filtre Exchange sur l'intitulé de poste (ou personnalisé), puis crée après confirmation un groupe de distribution dynamique avec ce filtre. Nécessite une session Exchange Online active : les groupes de distribution dynamiques sont une fonctionnalité Exchange, et non Graph.

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-JobTitle` | * | Intitulé de poste sur lequel filtrer |
| `-RecipientFilter` | * | Chaîne de filtre de destinataires Exchange personnalisée |
| `-Name` | Non | Nom du groupe (par défaut : l'intitulé de poste ; obligatoire avec `-RecipientFilter`) |
| `-PrimarySmtpAddress` | Non | Adresse SMTP du nouveau groupe |
| `-Apply` | Non | Créer réellement le groupe (par défaut : aperçu uniquement) |

*L'un des paramètres `-JobTitle` / `-RecipientFilter` est obligatoire.

```powershell
.\New-DynamicDistributionGroupByFilter.ps1 -JobTitle "Sales Manager"
.\New-DynamicDistributionGroupByFilter.ps1 -JobTitle "Sales Manager" -Apply

.\New-DynamicDistributionGroupByFilter.ps1 -Name "Finance Dept" `
    -RecipientFilter "((Department -eq 'Finance') -and (ExchangeUserAccountControl -ne 'AccountDisabled'))" -Apply
```

**Module requis**
```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

---

### Add-UserToFeatureGroup.ps1

Ajoute un utilisateur comme membre direct d'un groupe de sécurité Entra ID, ou l'en retire : la méthode prise en charge pour limiter une stratégie Conditional Access, une attribution Intune ou un groupe de licences à un sous-ensemble d'utilisateurs. Remplace la technique héritée consistant à marquer un utilisateur via un attribut personnalisé de boîte aux lettres (à des fins de rapport uniquement, sans délimiter réellement quoi que ce soit).

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-UserId` | Oui | UPN ou ID d'objet |
| `-GroupId` | * | ID d'objet du groupe cible |
| `-GroupName` | * | Nom d'affichage du groupe cible (résolu automatiquement) |
| `-Remove` | Non | Retirer au lieu d'ajouter |
| `-TenantId` | Non | ID ou domaine du tenant Entra ID |
| `-Apply` | Non | Modifier réellement l'appartenance (par défaut : aperçu uniquement) |

*L'un des paramètres `-GroupId` / `-GroupName` est obligatoire.

```powershell
.\Add-UserToFeatureGroup.ps1 -UserId "user@contoso.com" -GroupName "SG - Enable Windows 365" -Apply
.\Add-UserToFeatureGroup.ps1 -UserId "user@contoso.com" -GroupName "SG - Enable Windows 365" -Remove -Apply
```

**Étendues requises :** `GroupMember.ReadWrite.All`, `User.Read.All`, `Group.Read.All`
**Modules requis :** `Microsoft.Graph.Users`, `Microsoft.Graph.Groups`

---

Se connecte automatiquement à Microsoft Graph si aucune session n'est active, et réutilise la session existante si vous êtes déjà connecté, comme `Test-M365GroupMembership.ps1` dans `scripts/Entra/`.
