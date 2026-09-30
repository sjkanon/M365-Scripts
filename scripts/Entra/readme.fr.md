[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../readme.fr.md) › [scripts](../readme.fr.md) › **Entra**

# Scripts Entra ID

Scripts de gestion des utilisateurs et des ressources dans Microsoft Entra ID (anciennement Azure AD) via Microsoft Graph.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Set-UserManager.ps1`](Set-UserManager.ps1) ([docs](#set-usermanagerps1)) | Signaler et, au besoin, définir en masse le responsable d'un ensemble d'utilisateurs Entra ID |
| [`Remove-M365Users.ps1`](Remove-M365Users.ps1) ([docs](#remove-m365usersps1)) | Supprimer en masse des comptes utilisateurs M365 (essai à blanc par défaut) |
| [`New-M365User.ps1`](New-M365User.ps1) ([docs](#new-m365userps1)) | Créer un utilisateur M365, licence facultative |
| [`Import-M365Users.ps1`](Import-M365Users.ps1) ([docs](#import-m365usersps1)) | Créer en masse des utilisateurs M365 depuis un CSV (essai à blanc par défaut) |
| [`Get-M365UserLicenses.ps1`](Get-M365UserLicenses.ps1) ([docs](#get-m365userlicensesps1)) | Signaler les licences attribuées pour une liste d'utilisateurs |
| [`Import-ConditionalAccessBaseline.ps1`](Import-ConditionalAccessBaseline.ps1) ([docs](#import-conditionalaccessbaselineps1)) | Importer la baseline Conditional Access de la communauté |
| [`Test-M365GroupMembership.ps1`](Test-M365GroupMembership.ps1) ([docs](#test-m365groupmembershipps1)) | Auditer les propriétaires et membres des groupes M365 / Teams |
| [`Copy-GroupMember.ps1`](Copy-GroupMember.ps1) ([docs](#copy-groupmemberps1)) | Copier les membres d'un groupe Entra ID vers un autre (essai à blanc par défaut) |
| [`New-TemporaryConditionalAccessPolicy.ps1`](New-TemporaryConditionalAccessPolicy.ps1) ([docs](#new-temporaryconditionalaccesspolicyps1)) | Créer une stratégie CA temporaire pour un utilisateur ou un groupe |
| [`Remove-TemporaryConditionalAccessPolicies.ps1`](Remove-TemporaryConditionalAccessPolicies.ps1) ([docs](#remove-temporaryconditionalaccesspoliciesps1)) | Supprimer les stratégies CA temporaires expirées ou toutes |
| [`New-UserTemporaryAccessPass.ps1`](New-UserTemporaryAccessPass.ps1) ([docs](#new-usertemporaryaccesspassps1)) | Créer un code TAP pour un utilisateur |
| [`Set-EntraPasskeyMigrationOptOut.ps1`](Set-EntraPasskeyMigrationOptOut.ps1) ([docs](#set-entrapasskeymigrationoptoutps1)) | Reporter l'activation automatique des passkeys du 1er septembre 2026 (un seul tenant ou une liste GDAP) |
| [`Phising-rollout.ps1`](Phising-rollout.ps1) ([docs](#phising-rolloutps1)) | Garder synchronisés, dans les deux sens, un groupe de déploiement de MFA résistante au hameçonnage et un groupe d'utilisateurs enregistrés |

> La conversion de groupes de distribution dynamiques en statiques (`Set-Distributionlist-dynamic-static.ps1`) se trouve dans [`scripts/Exchange/`](../Exchange/readme.fr.md) — elle utilise des cmdlets Exchange Online, pas Graph.

---

### Set-UserManager.ps1

Signale et, au besoin, définit en masse le responsable d'un ensemble d'utilisateurs Entra ID. Les utilisateurs peuvent être sélectionnés par groupe, service, responsable actuel ou liste explicite d'UPN ; lorsque `-NewManager` est omis, le script se contente de signaler le responsable actuel de chaque utilisateur.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-GroupId` | * | ID d'objet d'un groupe Entra ID dont les membres doivent être traités |
| `-GroupName` | * | Nom d'affichage d'un groupe Entra ID (converti automatiquement en ID ; erreur s'il est ambigu) |
| `-Department` | * | Chaîne de service servant à filtrer les utilisateurs (correspondance exacte via OData) |
| `-CurrentManager` | * | UPN ou ID d'objet d'un responsable — traite tous ses subordonnés directs dans l'ensemble du tenant |
| `-UserList` | * | Tableau explicite d'UPN ou d'ID d'objet |
| `-NewManager` | Non | UPN ou ID d'objet à définir comme responsable pour tous les utilisateurs trouvés. Omettez-le pour seulement signaler les responsables actuels |
| `-OutputPath` | Non | Chemin d'export CSV |
| `-TenantId` | Non | ID ou domaine du tenant Entra ID pour `Connect-MgGraph` |

*Un seul paramètre parmi `-GroupId` / `-GroupName` / `-Department` / `-CurrentManager` / `-UserList` sélectionne la source des utilisateurs.

**Exemples**

```powershell
# Afficher les responsables de tous les membres d'un groupe
.\Set-UserManager.ps1 -GroupName "Sales Team"

# Définir en masse le responsable pour un groupe
.\Set-UserManager.ps1 -GroupName "Sales Team" -NewManager "jane.doe@contoso.com"

# Définir en masse le responsable pour un service, exporter le rapport
.\Set-UserManager.ps1 -Department "Logistics" -NewManager "jane.doe@contoso.com" -OutputPath C:\Temp\ManagerReport.csv

# Trouver tous les subordonnés directs d'un responsable
.\Set-UserManager.ps1 -CurrentManager "old.boss@contoso.com"

# Réaffecter tous les subordonnés directs d'un responsable à un autre
.\Set-UserManager.ps1 -CurrentManager "old.boss@contoso.com" -NewManager "new.boss@contoso.com"
```

Prend en charge `-WhatIf` (`SupportsShouldProcess`).

---

### Remove-M365Users.ps1

Supprime en masse des comptes utilisateurs M365 d'un tenant. Révoque les sessions et retire les licences avant la suppression. Fonctionne par défaut en essai à blanc — passez `-Apply` pour effectuer réellement les suppressions.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-UserList` | * | Tableau d'UPN à supprimer |
| `-CsvPath` | * | Chemin vers un CSV (colonne `UserPrincipalName`) ou un TXT (un UPN par ligne) |
| `-Apply` | Non | Effectuer réellement les suppressions (par défaut : essai à blanc) |
| `-SkipLicenseRemoval` | Non | Ne pas retirer les licences avant la suppression |
| `-SkipSessionRevoke` | Non | Ne pas révoquer les sessions actives |
| `-OutputPath` | Non | Chemin du rapport CSV (par défaut : `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Non | ID ou domaine du tenant Entra ID |

*`-UserList` ou `-CsvPath` est obligatoire.

**Exemples**

```powershell
# Essai à blanc — aucune modification
.\Remove-M365Users.ps1 -UserList "user1@contoso.com","user2@contoso.com"

# Suppression réelle depuis un CSV
.\Remove-M365Users.ps1 -CsvPath .\users.csv -Apply

# Entrée par le pipeline
"user1@contoso.com","user2@contoso.com" | .\Remove-M365Users.ps1 -Apply
```

**Remarques**
- La suppression est réversible (soft delete) — les comptes arrivent dans Utilisateurs supprimés et sont récupérables pendant 30 jours
- Un rapport CSV est toujours écrit, même en essai à blanc
- Si une connexion Graph existe déjà (par ex. via `functies.ps1`), elle est réutilisée

**Module requis**
```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

---

### New-M365User.ps1

Crée un utilisateur M365 via Microsoft Graph. Génère un mot de passe aléatoire de 16 caractères si aucun n'est fourni. Attribue éventuellement une licence après la création.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-UserPrincipalName` | Oui | UPN du nouvel utilisateur |
| `-DisplayName` | Oui | Nom d'affichage |
| `-GivenName` | Non | Prénom |
| `-Surname` | Non | Nom de famille |
| `-Password` | Non | Mot de passe initial (généré automatiquement s'il est omis) |
| `-UsageLocation` | Non | Code pays ISO à deux lettres (par défaut : `NL`) |
| `-Department` | Non | Service |
| `-JobTitle` | Non | Fonction |
| `-MobilePhone` | Non | Numéro de téléphone mobile |
| `-LicenseSkuId` | Non | Référence de la SKU de licence à attribuer (par ex. `ENTERPRISEPACK`) |
| `-NoPasswordReset` | Non | Ne pas imposer de changement de mot de passe à la première connexion |
| `-TenantId` | Non | ID ou domaine du tenant Entra ID |

**Exemples**

```powershell
# Minimal — mot de passe généré automatiquement
.\New-M365User.ps1 -UserPrincipalName "j.doe@contoso.com" -DisplayName "Jane Doe"

# Informations complètes avec licence
.\New-M365User.ps1 -UserPrincipalName "j.doe@contoso.com" -DisplayName "Jane Doe" `
    -GivenName "Jane" -Surname "Doe" -Department "Finance" -LicenseSkuId "ENTERPRISEPACK"
```

---

### Import-M365Users.ps1

Crée en masse des utilisateurs M365 à partir d'un fichier CSV via Microsoft Graph. Fonctionne par défaut en essai à blanc — passez `-Apply` pour créer les comptes. Génère un mot de passe unique par utilisateur si aucune colonne Password n'est présente. Les mots de passe sont écrits dans le CSV de résultats.

**Colonnes CSV obligatoires :** `UserPrincipalName`, `DisplayName`

**Colonnes CSV facultatives :** `GivenName`, `Surname` (ou `LastName`), `Department`, `JobTitle`, `MobilePhone`, `UsageLocation`, `Password`, `LicenseSkuId`

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-CsvPath` | Oui | Chemin du CSV d'entrée |
| `-Apply` | Non | Créer réellement les comptes (par défaut : essai à blanc) |
| `-UsageLocation` | Non | Code pays par défaut pour tous les utilisateurs (par défaut : `NL`) |
| `-LicenseSkuId` | Non | Attribuer cette licence à tous les utilisateurs (remplace la colonne CSV) |
| `-NoPasswordReset` | Non | Ne pas imposer de changement de mot de passe à la première connexion |
| `-OutputPath` | Non | Chemin du rapport CSV (par défaut : `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Non | ID ou domaine du tenant Entra ID |

**Exemples**

```powershell
# D'abord un essai à blanc
.\Import-M365Users.ps1 -CsvPath .\users.csv

# Créer les comptes
.\Import-M365Users.ps1 -CsvPath .\users.csv -Apply

# Créer avec une licence pour tout le monde
.\Import-M365Users.ps1 -CsvPath .\users.csv -LicenseSkuId "ENTERPRISEPACK" -Apply
```

**Remarques**
- L'essai à blanc écrit toujours un CSV de résultats — vérifiez-le avant de lancer avec `-Apply`
- Les mots de passe générés figurent dans le CSV de résultats — transmettez-les de manière sécurisée
- Une licence nécessite que `UsageLocation` soit défini ; le script s'en charge automatiquement

---

### Get-M365UserLicenses.ps1

Vérifie les licences attribuées pour une liste d'utilisateurs via Microsoft Graph et exporte un rapport CSV. Accepte une liste en ligne, un CSV, un TXT et une entrée par le pipeline.

**Options d'entrée**
- `-UserList` avec des UPN/adresses e-mail
- `-CsvPath` vers un `.csv` (colonne : `UserPrincipalName`, `UPN` ou `Mail`)
- `-CsvPath` vers un `.txt` (un UPN/une adresse e-mail par ligne)

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-UserList` | * | Tableau d'UPN/adresses e-mail |
| `-CsvPath` | * | Chemin vers un CSV/TXT contenant les utilisateurs |
| `-OutputPath` | Non | Chemin du rapport CSV (par défaut : `C:\Temp\UserLicenseReport_<timestamp>.csv`) |
| `-TenantId` | Non | ID ou domaine du tenant Entra ID |

*`-UserList` ou `-CsvPath` est obligatoire.

**Exemples**

```powershell
# Vérifier des utilisateurs précis
.\Get-M365UserLicenses.ps1 -UserList "user1@contoso.com","user2@contoso.com"

# Vérifier des utilisateurs depuis un CSV
.\Get-M365UserLicenses.ps1 -CsvPath .\users.csv

# Entrée par le pipeline
"user1@contoso.com","user2@contoso.com" | .\Get-M365UserLicenses.ps1
```

**Sortie du rapport**
- Une ligne par attribution de licence à un utilisateur
- Les utilisateurs sans licence sont inclus avec `LicenseStatus = Unlicensed`
- Les utilisateurs introuvables sont inclus avec `LicenseStatus = NotFound`

---

### Import-ConditionalAccessBaseline.ps1

Importe la dernière version de [ConditionalAccessBaseline](https://github.com/j0eyv/ConditionalAccessBaseline) dans votre tenant à l'aide de Microsoft Graph.

Ce qu'il fait :
- Télécharge la dernière baseline (ou utilise `-SourcePath`)
- Crée ou réutilise les groupes d'exclusion CA requis
- Crée ou réutilise les emplacements nommés
- Fait correspondre les anciens ID de la baseline aux ID de votre tenant
- Importe les stratégies Conditional Access pour tous les personas/plateformes
- Importe les stratégies à l'état **Désactivé** par défaut (`state = disabled`)

Il prend aussi en charge une action ultérieure pour passer les stratégies importées en mode rapport uniquement ou les activer.

**Paramètres (les plus utilisés)**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-Action` | Non | `Import` (par défaut) ou `SetState` |
| `-PolicyStateOnImport` | Non | `disabled` (par défaut) ou `enabledForReportingButNotEnforced` |
| `-TargetState` | Non | Pour `SetState` : `disabled`, `enabledForReportingButNotEnforced` ou `enabled` |
| `-SourcePath` | Non | Dossier local de la baseline contenant `Config\...` |
| `-TenantId` | Non | ID ou domaine du tenant |
| `-UpdateExisting` | Non | Mettre à jour les stratégies existantes dont le nom d'affichage correspond |

**Exemples**

```powershell
# Importer la dernière baseline et laisser toutes les stratégies CA DÉSACTIVÉES
.\Import-ConditionalAccessBaseline.ps1

# Importer la baseline en mode rapport uniquement
.\Import-ConditionalAccessBaseline.ps1 -PolicyStateOnImport enabledForReportingButNotEnforced

# Plus tard : activer les stratégies de la baseline importée
.\Import-ConditionalAccessBaseline.ps1 -Action SetState -TargetState enabled
```

**Remarques**
- Gardez au moins un compte d'urgence (break-glass) exclu avant d'activer les stratégies
- Vérifiez les groupes d'exclusion et les emplacements nommés après l'import

---

### New-TemporaryConditionalAccessPolicy.ps1

Crée une stratégie Conditional Access temporaire pour un utilisateur ou un groupe.

- Marque le nom de la stratégie avec le préfixe `TEMP-CA -`
- Écrit un horodatage d'expiration dans la description de la stratégie (`Expires=<UTC timestamp>`)
- Prend en charge des fenêtres basées sur une durée ou sur des dates/heures locales exactes de début et de fin
- Par défaut, garde la session du script ouverte et supprime la stratégie dès que l'expiration est atteinte

Important :
- Le nettoyage immédiat à l'expiration exige que la session du script reste ouverte
- Si vous fermez la session plus tôt, lancez le script de nettoyage ultérieurement
- Ce script ne crée pas de tâche planifiée Windows ; la boucle d'attente/de nettoyage s'exécute dans la session courante

**Exemples**

```powershell
# Exigence MFA temporaire pendant 2 heures, suppression automatique à l'expiration
.\New-TemporaryConditionalAccessPolicy.ps1 -TargetType User -TargetId "<object-id>" -DisplayName "Temporary MFA" -DurationHours 2

# Stratégie temporaire avec date/heure locale exacte de début et de fin
.\New-TemporaryConditionalAccessPolicy.ps1 -TargetType User -TargetId "<object-id>" -DisplayName "Install Window" -StartDateTimeLocal "2026-07-23 19:00" -EndDateTimeLocal "2026-07-23 22:00"

# Stratégie de blocage temporaire, sans boucle d'attente de nettoyage automatique
.\New-TemporaryConditionalAccessPolicy.ps1 -TargetType Group -TargetId "<object-id>" -DisplayName "Temporary Block" -Action Block -NoAutoCleanup
```

---

### Remove-TemporaryConditionalAccessPolicies.ps1

Supprime les stratégies temporaires créées avec le préfixe `TEMP-CA -`.

Modes :
- par défaut : supprimer uniquement les stratégies TEMP-CA expirées
- `-PolicyId` : supprimer une stratégie précise
- `-RemoveAllTempPolicies` : supprimer toutes les stratégies TEMP-CA

**Exemples**

```powershell
# Supprimer uniquement les stratégies temporaires expirées
.\Remove-TemporaryConditionalAccessPolicies.ps1

# Supprimer une stratégie précise
.\Remove-TemporaryConditionalAccessPolicies.ps1 -PolicyId "<policy-id>"
```

---

### New-UserTemporaryAccessPass.ps1

Crée un Temporary Access Pass (TAP) pour un utilisateur.

**Exemple**

```powershell
.\New-UserTemporaryAccessPass.ps1 -UserId "user@contoso.com" -LifetimeMinutes 60 -IsUsableOnce
```

**Remarques**
- Privilégiez `-IsUsableOnce` pour les scénarios de support/d'installation
- Transmettez le code TAP par un canal sécurisé et faites-le expirer rapidement

---

### Test-M365GroupMembership.ps1

Liste tous les propriétaires et membres des groupes Microsoft 365 (y compris les groupes associés à Teams). Les résultats sont exportés en CSV avec une ligne par propriétaire/membre. Se connecte automatiquement à Graph si aucune session n'est active ; réutilise une session existante si une connexion est déjà établie.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-Group` | Non | Nom d'affichage ou ID d'objet d'un seul groupe. S'il est omis, tous les groupes M365 sont audités |
| `-OutputPath` | Non | Chemin du rapport CSV (par défaut : `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Non | ID ou domaine du tenant Entra ID |

**Exemples**

```powershell
# Auditer tous les groupes M365
.\Test-M365GroupMembership.ps1

# Un seul groupe par nom d'affichage
.\Test-M365GroupMembership.ps1 -Group "Team Finance"

# Un seul groupe par ID d'objet
.\Test-M365GroupMembership.ps1 -Group "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
```

**Étendues requises**
- `Group.Read.All`
- `Directory.Read.All`

**Module requis**
```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

---

### Copy-GroupMember.ps1

Copie les membres d'un groupe Entra ID dans un autre groupe. Les membres déjà présents dans la cible sont ignorés, le script peut donc être relancé sans risque. Fonctionne par défaut en essai à blanc — passez `-Apply` pour écrire les modifications.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-SourceGroup` | Oui | Nom d'affichage ou ID d'objet du groupe SOURCE de la copie |
| `-TargetGroup` | Oui | Nom d'affichage ou ID d'objet du groupe CIBLE de la copie |
| `-MemberType` | Non | `All` (par défaut), `User`, `Group`, `Device` ou `ServicePrincipal` |
| `-Flatten` | Non | Développer les groupes imbriqués et copier leurs membres effectifs au lieu de l'objet groupe imbriqué |
| `-Mirror` | Non | Retirer aussi de la cible les membres absents de la source (copie exacte au lieu d'une union) |
| `-Apply` | Non | Ajouter/retirer réellement les membres (par défaut : essai à blanc) |
| `-Disconnect` | Non | Se déconnecter de Graph à la fin (désactivé par défaut — la déconnexion vide le cache de jetons et impose une nouvelle invite dans le navigateur à l'exécution suivante) |
| `-OutputPath` | Non | Chemin du rapport CSV (par défaut : `C:\Temp\GroupMemberCopy_<timestamp>.csv`) |
| `-TenantId` | Non | ID ou domaine du tenant Entra ID |

**Exemples**

```powershell
# Essai à blanc — afficher ce qui serait copié
.\Copy-GroupMember.ps1 -SourceGroup "All Staff" -TargetGroup "MFA Rollout"

# Copier réellement les membres
.\Copy-GroupMember.ps1 -SourceGroup "All Staff" -TargetGroup "MFA Rollout" -Apply

# Copier uniquement les utilisateurs, en développant les groupes imbriqués
.\Copy-GroupMember.ps1 -SourceGroup "Sales" -TargetGroup "Sales Mail" -MemberType User -Flatten -Apply

# Faire de la cible une copie exacte de la source (ajouts et retraits)
.\Copy-GroupMember.ps1 -SourceGroup "Pilot" -TargetGroup "Pilot Copy" -Mirror -Apply
```

**Remarques**
- Les noms d'affichage sont résolus via Graph ; un nom ambigu provoque une erreur bloquante — utilisez alors l'ID d'objet
- Un groupe cible à appartenance dynamique est refusé : son appartenance est pilotée par des règles et ne peut pas être modifiée
- Les groupes de sécurité à extension messagerie et les groupes de distribution ne sont pas modifiables via Graph — utilisez les cmdlets Exchange Online pour ceux-ci
- Prend en charge `-WhatIf` (`SupportsShouldProcess`)

**Étendues requises**
- `Group.Read.All`
- `GroupMember.ReadWrite.All`

**Module requis**
```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

---

### Set-EntraPasskeyMigrationOptOut.ps1

Définit (ou lève) la désactivation temporaire (opt-out) de l'activation automatique des passkeys et du déploiement de la Registration Campaign d'Entra ID. Il modifie la stratégie des méthodes d'authentification du tenant :

```http
PATCH https://graph.microsoft.com/beta/policies/authenticationmethodspolicy
{ "optOutSettings": { "passkeyDynamicMigration": true } }
```

Accepte un tableau de tenants, pour qu'un partenaire GDAP puisse parcourir tous les tenants clients en une seule exécution. Chaque tenant est traité indépendamment — un tenant pour lequel la connexion, la lecture ou la modification échoue est signalé, et l'exécution continue avec le suivant.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-TenantId` | Non | Un ou plusieurs ID de tenant ou noms de domaine. Omettez-le pour utiliser la connexion actuelle / le tenant par défaut |
| `-Revert` | Non | Remettre `passkeyDynamicMigration` à `false` (réinscrire le tenant DANS la migration automatique) |
| `-ReportOnly` | Non | Lire et afficher la valeur actuelle sans rien modifier |

**Exemples**

```powershell
# Afficher le paramètre actuel pour le tenant connecté
.\Set-EntraPasskeyMigrationOptOut.ps1 -ReportOnly

# Prévisualiser la modification sans l'écrire
.\Set-EntraPasskeyMigrationOptOut.ps1 -TenantId contoso.onmicrosoft.com -WhatIf

# Désinscrire chaque tenant listé dans un fichier texte, sans confirmation par tenant
.\Set-EntraPasskeyMigrationOptOut.ps1 -TenantId (Get-Content .\tenants.txt) -Confirm:$false

# Réinscrire un tenant DANS la migration automatique
.\Set-EntraPasskeyMigrationOptOut.ps1 -TenantId contoso.onmicrosoft.com -Revert
```

**Calendrier**

| Date | Ce qui se passe |
|------|--------------|
| 1er septembre 2026 | Les utilisateurs activés pour les SMS ou la voix sont automatiquement activés pour les passkeys et incités par une campagne d'inscription gérée par Microsoft |
| 1er février 2027 | L'envoi SMS/voix fourni par Microsoft est retiré |
| Après le 1er février 2027 | Les utilisateurs dont la seule méthode MFA est le SMS ou la voix reçoivent à la connexion une invite **bloquante** d'inscription d'une passkey |

**Remarques**
- La désactivation reporte **uniquement** le comportement du 1er septembre 2026 → 1er février 2027. Il n'existe aucune désactivation pour l'application à partir du 1er février 2027 — elle s'applique à tous les tenants
- `optOutSettings` n'existe **qu'en beta** en août 2026 et n'est pas exposé dans le centre d'administration Entra
- Les écritures sont confirmées par tenant (`ConfirmImpact = 'High'`) ; passez `-Confirm:$false` pour les exécutions multi-tenants sans surveillance
- Le paramètre est relu environ 2 secondes après le PATCH ; un écart est signalé comme `PatchedUnverified` au lieu d'être traité comme un succès
- Renvoie un objet par tenant (`Tenant`, `Before`, `After`, `Status`, `Message`), de sorte qu'une exécution peut être envoyée vers `Export-Csv`
- Prend en charge `-WhatIf` (`SupportsShouldProcess`)
- Si vous avez besoin des SMS/de la voix après le 1er février 2027, configurez un fournisseur télécom géré par le client via le Microsoft Security Store (sélectionnable à partir du 30 octobre 2026)

**Étendue requise**
- `Policy.ReadWrite.AuthenticationMethod` (`Policy.Read.All` suffit pour `-ReportOnly`)

**Rôle requis**
- Authentication Policy Administrator (ou Global Administrator)

**Module requis**
```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
```

**Référence** — [Passkeys by default and retirement of Microsoft-provided SMS and voice authentication](https://learn.microsoft.com/en-us/entra/identity/authentication/concept-sms-voice-retirement)

---

### Phising-rollout.ps1

Maintient deux groupes statiques mutuellement exclusifs qui pilotent le déploiement d'une
MFA résistante au hameçonnage : un groupe **Rollout** pour les utilisateurs qui doivent
encore s'inscrire, et un groupe **Registered** pour ceux qui l'ont fait. Un utilisateur
n'est jamais dans les deux.

Pour chaque utilisateur figurant actuellement dans l'un des deux groupes :

| Situation | Ce qui se passe |
|-----------|--------------|
| Possède une méthode acceptée et se trouve dans Rollout | Déplacé vers Registered — il a franchi l'étape |
| N'a plus de méthode acceptée et se trouve dans Registered | Renvoyé vers Rollout — la méthode a été supprimée ou a expiré, il doit donc s'inscrire à nouveau |
| Tout autre cas | Rien ; le statut est déjà correct |

Le retour en arrière compte autant que l'avancée : sans lui, un utilisateur qui supprime sa
passkey conserve discrètement l'étiquette « conforme ».

**Ce qui compte comme inscrit**

| `-AcceptedMethod` | Signification |
|-------------------|---------|
| `AuthenticatorPasskey` (par défaut) | Uniquement une passkey dans Microsoft Authenticator — une `fido2AuthenticationMethod` dont l'AAGUID figure dans `-AllowedAaGuids`. Une YubiKey, Windows Hello for Business ou CBA ne compte **pas** et laisse l'utilisateur dans Rollout |
| `AnyPhishingResistant` | Toute méthode résistante au hameçonnage compte (clés FIDO2, WHfB, Platform SSO, CBA) |

> Le portail Entra appelle simplement une telle méthode « Passkey », avec un détail comme
> « MS Authenticator iOS ». Dans Graph, ce n'est pas un type de méthode distinct mais une
> méthode fido2 — seul l'AAGUID révèle qu'il s'agit d'Authenticator, c'est pourquoi la
> valeur par défaut est un filtre sur l'AAGUID plutôt qu'un nom de méthode.

**Paramètres**

| Paramètre | Description |
|-----------|-------------|
| `-RolloutGroupId` | ID d'objet du groupe statique Rollout (doit encore s'inscrire) |
| `-RegisteredGroupId` | ID d'objet du groupe statique Registered (déjà conforme) |
| `-AcceptedMethod` | `AuthenticatorPasskey` (par défaut) ou `AnyPhishingResistant` |
| `-AllowedAaGuids` | AAGUID qui comptent comme passkey dans Authenticator (par défaut : les AAGUID iOS et Android de Microsoft Authenticator). Ignoré pour `AnyPhishingResistant` |
| `-Interactive` | Se connecter via le navigateur au lieu d'une identité managée — pour l'exécuter depuis votre propre poste |
| `-UseAppRegistration` | Utiliser une inscription d'application existante (certificat ou secret) |
| `-UseTemporaryApp` | Créer une inscription d'application jetable, s'exécuter en app-only, puis la supprimer ensuite |

**Authentification**

Conçu comme **runbook Azure Automation** sur une identité managée affectée par le système,
à exécuter toutes les une à quatre heures. Sans `-Interactive` ni inscription
d'application, il tente l'identité managée, ce qui échoue toujours en dehors d'Azure.

Pour une exécution portant sur de nombreux utilisateurs, `-UseTemporaryApp` (ou
`-UseAppRegistration` avec un certificat) est le choix fiable : un jeton app-only est émis
à neuf à chaque appel et ne peut jamais déclencher une invite dans le navigateur au milieu
de l'exécution. Une session interactive le peut, et le fait — précisément lorsque le jeton
expire en cours d'exécution.

**Autorisations d'application Graph requises**

```
User.Read.All
UserAuthenticationMethod.Read.All
GroupMember.ReadWrite.All      (ou plus large : Group.ReadWrite.All)
```

`-UseTemporaryApp` nécessite en outre Application Administrator ou Global Administrator,
car il crée une application et lui attribue des rôles d'application.

> Le fichier s'écrit `Phising-rollout.ps1` dans le dépôt. Le renommer est une modification
> distincte — le runbook Automation qui l'appelle fait référence à ce nom.
