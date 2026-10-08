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

Affiche d'abord les destinataires correspondant à un filtre Exchange sur l'intitulé de poste (ou personnalisé), puis crée après confirmation un groupe de distribution dynamique avec ce filtre. Exchange Online uniquement : Microsoft Graph n'a pas d'API pour les groupes de distribution dynamiques ni pour l'aperçu d'un filtre de destinataires. Se connecte via `Connect-M365Exchange` ([`Connect-M365.ps1`](../../Startup/Connect-M365.ps1)) : délégué par défaut (administrateur Exchange ; code d'appareil selon `$global:useDeviceCodeAuth` ; le client GDAP via `-DelegatedOrganization`), application seule avec `-ClientId` + `-CertificateThumbprint` et `-TenantId` sous forme de domaine `*.onmicrosoft.com`. Une session Exchange existante pour le tenant est réutilisée et reste ouverte.

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-JobTitle` | * | Intitulé de poste sur lequel filtrer |
| `-RecipientFilter` | * | Chaîne de filtre de destinataires Exchange personnalisée |
| `-Name` | Non | Nom du groupe (par défaut : l'intitulé de poste ; obligatoire avec `-RecipientFilter`) |
| `-PrimarySmtpAddress` | Non | Adresse SMTP du nouveau groupe (par défaut : déduite par Exchange) |
| `-TenantId` | Non | Domaine ou ID du tenant (par défaut : client GDAP, sinon le tenant de connexion ; l'application seule exige le domaine) |
| `-ClientId` / `-CertificateThumbprint` | Non | Connexion en application seule |
| `-AppOnly` | Non | Application seule avec le ClientId et l'empreinte de `graph.appid.json` |
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
| `-TenantId` | Non | ID ou domaine du tenant Entra ID (par défaut : client GDAP, sinon le tenant de connexion) |
| `-ClientId` / `-CertificateThumbprint` | Non | Connexion en application seule |
| `-AppOnly` | Non | Application seule avec le ClientId et l'empreinte de `graph.appid.json` |
| `-Apply` | Non | Modifier réellement l'appartenance (par défaut : aperçu uniquement) |

*L'un des paramètres `-GroupId` / `-GroupName` est obligatoire.

```powershell
.\Add-UserToFeatureGroup.ps1 -UserId "user@contoso.com" -GroupName "SG - Enable Windows 365" -Apply
.\Add-UserToFeatureGroup.ps1 -UserId "user@contoso.com" -GroupName "SG - Enable Windows 365" -Remove -Apply
```

**Étendues requises :** `GroupMember.ReadWrite.All`, `User.Read.All`, `Group.Read.All`
**Modules requis :** `Microsoft.Graph.Users`, `Microsoft.Graph.Groups`

---

`Add-UserToFeatureGroup.ps1` se connecte via [`Connect-M365.ps1`](../../Startup/Connect-M365.ps1) : **délégué par défaut** (vous vous connectez en tant qu'administrateur ; code d'appareil si `$global:useDeviceCodeAuth` est défini ; en GDAP, le client de `$global:cid`), **application seule** avec `-ClientId` + `-CertificateThumbprint` ou `-AppOnly`. Une session Graph existante n'est réutilisée que si elle porte sur le bon tenant et dispose des scopes ci-dessus ; le script ne ferme que la session qu'il a lui-même ouverte.
