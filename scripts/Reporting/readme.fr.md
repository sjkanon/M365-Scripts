[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../readme.fr.md) › [scripts](../readme.fr.md) › **Reporting**

# Reporting Scripts

Scripts qui génèrent des rapports sur Active Directory, SharePoint Online et les licences.

---

## Dossiers

| Dossier | Description |
|--------|-------------|
| [`Licensing/`](Licensing/readme.fr.md) | Rapport mensuel des licences et des coûts Azure à partir des données Pax8 et Ingram |

## Scripts

| Script | Description |
|--------|-------------|
| [`Get-ComputerLastLogon.ps1`](Get-ComputerLastLogon.ps1) ([docs](#get-computerlastlogonps1)) | Date de dernière connexion des objets ordinateur d'une ou plusieurs OU, avec export CSV |
| [`Get-SharePointStorageReport.ps1`](Get-SharePointStorageReport.ps1) ([docs](#get-sharepointstoragereportps1)) | Rapport de stockage à l'échelle du tenant : sites, bibliothèques, historique des versions et corbeille |
| [`Get-SharePointPermissionsReport.ps1`](Get-SharePointPermissionsReport.ps1) ([docs](#get-sharepointpermissionsreportps1)) | Qui a accès à quoi, via quel groupe et à quel niveau — chaque site, liste, dossier et fichier doté de ses propres autorisations. Lecture seule, vers CSV et un unique classeur Excel |
| [`Remove-SharePointFileVersionsByDate.ps1`](Remove-SharePointFileVersionsByDate.ps1) ([docs](#remove-sharepointfileversionsbydateps1)) | Supprime les versions de fichiers antérieures à une date ; la version actuelle est toujours conservée. Rapport uniquement par défaut |

---

## Get-ComputerLastLogon.ps1

Indique la **date de dernière connexion** des objets ordinateur d'une ou plusieurs OU, avec export CSV.

### Fonctionnement

Deux modes de précision :

| Mode | Attribut | Décalage | Vitesse |
|---|---|---|---|
| Par défaut | `LastLogonTimestamp` (répliqué) | jusqu'à 14 jours | Rapide |
| `-AllDCs` | `LastLogon` par DC, meilleure valeur | Aucun | Plus lent |

Utilisez le mode par défaut pour les rapports sur les appareils inactifs (stale). Utilisez `-AllDCs` lorsqu'une précision absolue est requise.

### Paramètres

| Paramètre | Type | Défaut | Description |
|---|---|---|---|
| `-SearchBase` | `string[]` | _(domaine entier)_ | Un ou plusieurs distinguished names d'OU |
| `-AllDCs` | switch | désactivé | Interroge tous les DC pour obtenir le `LastLogon` le plus précis |
| `-InactiveDays` | int | `90` | Seuil en jours au-delà duquel un ordinateur est considéré comme _Stale_ |
| `-IncludeDisabled` | switch | désactivé | Inclut également les objets ordinateur désactivés |
| `-ExportPath` | string | `C:\Temp\` | Dossier du fichier CSV |

### Prérequis

- Module PowerShell ActiveDirectory (RSAT)
- Droits de lecture sur les OU indiquées

### Exemples

```powershell
# OU Laptops et Computers
.\Get-ComputerLastLogon.ps1 `
    -SearchBase "OU=Laptops,OU=Computers,DC=bedrijf,DC=local",
               "OU=Computers,DC=bedrijf,DC=local"

# Mode le plus précis — interroge tous les DC
.\Get-ComputerLastLogon.ps1 -SearchBase "OU=Computers,DC=bedrijf,DC=local" -AllDCs

# Avec les ordinateurs désactivés, seuil à 60 jours
.\Get-ComputerLastLogon.ps1 -SearchBase "OU=Computers,DC=bedrijf,DC=local" `
    -IncludeDisabled -InactiveDays 60
```

### Valeurs de statut

| Statut | Signification |
|---|---|
| `Active` | LastLogon dans la limite du seuil `-InactiveDays` |
| `Active (pwd recent)` | LastLogon semble ancien à cause du délai de réplication, mais le mot de passe du compte ordinateur a été renouvelé il y a moins de 35 jours — l'appareil est en ligne |
| `Stale` | LastLogon et PasswordLastSet dépassent tous deux le seuil — probablement réellement inactif |
| `Never` | Jamais connecté et aucun mot de passe récent |
| `Disabled` | Compte désactivé dans AD |

> **Astuce :** `Active (pwd recent)` désigne des PC qui *sont* actifs mais qui apparaissent à tort comme inactifs en raison du délai de réplication de 9 à 14 jours de `LastLogonTimestamp`. Utilisez `-AllDCs` pour des données exactes si cette distinction est critique.

### Colonnes CSV

| Colonne | Description |
|---|---|
| `Name` | Nom de l'ordinateur |
| `Status` | Voir les valeurs de statut ci-dessus |
| `Enabled` | True/False |
| `LastLogon` | Date de dernière connexion (dd/MM/yyyy HH:mm) |
| `DaysSinceLogon` | Nombre de jours écoulés |
| `PasswordLastSet` | Date du dernier changement de mot de passe du compte ordinateur (dd/MM/yyyy) |
| `DaysSincePasswordSet` | Nombre de jours depuis le renouvellement du mot de passe |
| `OperatingSystem` | Nom de l'OS |
| `OperatingSystemVersion` | Version de l'OS |
| `IPv4Address` | Adresse IP (si disponible) |
| `OU` | Chemin de l'OU (format lisible) |
| `Created` | Date de création dans AD |
| `Description` | Description issue d'AD |
| `DistinguishedName` | Chemin AD complet |

---

## Licensing/

Voir [Licensing/](Licensing/) pour le rapport mensuel des licences.

---

## Get-SharePointStorageReport.ps1

Rend compte de l'utilisation du stockage dans SharePoint Online grâce à une analyse de tout le tenant. Par défaut, le script se connecte en mode délégué et crée temporairement une App Registration (`Sites.Read.All`) pour l'énumération des sites ; cette application est supprimée à la fin.


### Couverture

- Toutes les site collections SharePoint (les sites personnels OneDrive sont exclus)
- Les sous-sites à tous les niveaux
- Les emplacements SharePoint liés à Teams :
    - Canaux standard sous forme de bibliothèques/dossiers dans le site Teams parent
    - Canaux privés/partagés sous forme de site collections distinctes
- Les dossiers et fichiers des bibliothèques de documents (uniquement avec `-Apply`)
- La sortie détaillée contient à la fois les dossiers et les fichiers (`ItemType`), pour que vous voyiez la structure complète
- En mode `-Apply`, tout est regroupé dans 1 CSV classé (plus gros dossiers + fichiers, historique des versions compris)
- Le CSV contient aussi `Level` (profondeur) : racine = `0`, dossier de premier niveau = `1`, etc.
- Le CSV contient aussi `ParentPath` pour les analyses hiérarchiques (arborescence dans Excel/Power BI)

### Corbeille (recycle bin)

La corbeille (stage 1 + stage 2) compte dans le quota de stockage du tenant ; elle est donc récupérée **séparément** de l'analyse des bibliothèques, et uniquement pour les véritables site collections SharePoint (pas OneDrive) :

- Par défaut (`-Apply`, en Phase 2b) ou seule avec **`-RecycleBinOnly`** (ignore entièrement l'analyse des bibliothèques, corbeille uniquement)
- Seules les site collections racines ont leur propre corbeille (les sous-webs partagent celle de la racine)
- Sortie : une ligne supplémentaire par site dans le CSV de synthèse (`Library = "Recycle Bin (stage 1 + 2)"`) plus une ligne par élément supprimé dans le CSV détaillé

```powershell
# Corbeille uniquement
.\Get-SharePointStorageReport.ps1 -RecycleBinOnly

# Analyse complète + corbeille comme phase supplémentaire
.\Get-SharePointStorageReport.ps1 -Apply
```

### Reprise après interruption (checkpoints) et progression

Avec `-Apply` (ou `-RecycleBinOnly`), un checkpoint est écrit dans le dossier de sortie après chaque bibliothèque (ou corbeille de site) terminée : `SharePoint_StorageReport_<hash>.state.json` + `.summary.partial.csv` + `.detail.partial.csv`. Le `<hash>` est dérivé de tous les paramètres d'analyse (site, mode, dossier de sortie, etc.), donc :

- **Relancer avec les mêmes paramètres** reprend automatiquement à partir de la dernière bibliothèque terminée — les bibliothèques déjà traitées sont ignorées (`[SKIP] Already completed in a previous run.`).
- **`-Restart`** supprime un checkpoint existant et relance l'analyse depuis le début, même si les paramètres sont identiques.
- Les fichiers de checkpoint sont nettoyés automatiquement dès que l'analyse se termine avec succès — s'ils sont toujours là, l'exécution précédente a été interrompue.

Pendant une longue analyse, le script affiche à la fois des lignes de journal défilantes et (dans une console interactive) des barres de progression imbriquées par phase : inventaire des sites/bibliothèques, analyse des dossiers/fichiers par bibliothèque, récupération de l'historique des versions et corbeilles. Lorsque Microsoft Graph applique un throttling (par exemple `activityLimitReached` pendant les recherches d'historique des versions), un message `[WAIT] throttled by Microsoft Graph — waiting ...` apparaît avec le temps d'attente, au lieu que le script semble bloqué sans rien dire.

> **Attention (appels délégués/SDK) :** par défaut, les cmdlets du SDK Microsoft.Graph relancent *elles-mêmes* silencieusement les 429/503, avec leur propre backoff interne qui peut respecter un `Retry-After` conséquent en cas de `activityLimitReached` — ce qui pouvait provoquer plusieurs minutes de silence sans que le message `[WAIT]` du script n'apparaisse jamais. Le script exécute donc `Set-MgRequestContext -ClientTimeout <-GraphTimeoutSec> -MaxRetry 0` juste après la connexion, pour que chaque appel du SDK Graph ait un timeout strict et que toutes les nouvelles tentatives passent par la logique propre et visible du script.

```powershell
# Reprend automatiquement une analyse de tenant interrompue
.\Get-SharePointStorageReport.ps1 -Apply

# Ignore le checkpoint et recommence depuis le début
.\Get-SharePointStorageReport.ps1 -Apply -Restart
```

### Totaux par site collection (comparaison avec le portail d'administration)

Les sous-sites et les canaux Teams partagent le quota de stockage de leur site collection racine, mais l'analyse les signale comme des **lignes de site distinctes** ; la corbeille est elle aussi suivie dans une **ligne à part**. Pris séparément, ces chiffres ne sont donc pas comparables un à un avec le chiffre unique « storage used » que le portail d'administration SharePoint affiche par site collection.

Avec `-Apply` (Phase 2c), le script additionne donc automatiquement le tout par site collection racine : les totaux des bibliothèques de tous les sous-sites/canaux sous-jacents + la corbeille de cette site collection. Sortie : `SharePoint_SiteCollectionTotals_<timestamp>.csv`, avec par site collection `LibrariesMB`, `RecycleBinMB`, `GrandTotalMB`/`GrandTotalGB` et le nombre de sous-sites/canaux comptabilisés. La console affiche aussi le top 10.

Si `GrandTotalGB` pour un site diffère encore du chiffre du portail d'administration, la cause la plus probable est l'une des suivantes :
- **Décalage temporel** — le chiffre du portail d'administration peut avoir jusqu'à 24 h de retard sur une analyse en direct
- **Dossiers ignorés silencieusement** — un message `[ERROR] Cannot read folder` dans la console signifie que cette arborescence de sous-dossiers (problème d'autorisation) n'a pas été comptée
- **Recherches de versions échouées** — elles se rabattent sur « 0 version » après des erreurs Graph répétées (rare, seulement après 3 tentatives échouées)
- Comparez d'abord sans `-Apply` (mode rapide) — il utilise le même chiffre officiel `quota.used` que le portail d'administration ; si celui-ci diffère déjà, l'écart ne vient pas du comptage de `-Apply` lui-même

### Performances (recherches d'historique des versions)

L'historique des versions est l'étape la plus coûteuse : par nature, 1 appel Graph par fichier. Trois optimisations limitent ce coût :

- **Ignoré lorsque le contrôle de version est désactivé** — si l'on sait avec certitude que le contrôle de version est désactivé pour une bibliothèque, aucun appel de version n'est fait par fichier (0 version est de toute façon la réponse). Si l'état est inconnu (repli via `Get-MgSiteDrive`), la récupération a toujours lieu par prudence.
- **Regroupé via le point de terminaison `$batch` de Graph** — les recherches de versions des fichiers d'une bibliothèque sont désormais récupérées par groupes de 20 en 1 appel HTTP, au lieu d'un appel distinct par fichier.
- **Nouvelles tentatives plus courtes pour ces appels** — 3 tentatives au maximum avec un backoff court (au lieu du `-MaxGraphRetry`/backoff standard, qui pour les appels critiques peut monter jusqu'à ~2 minutes par tentative). Une recherche de versions échouée se rabat sur « 0 version » au lieu de bloquer toute l'analyse.

`-SkipVersions` reste l'option la plus rapide si l'historique des versions n'est pas nécessaire — aucun appel de version n'est alors effectué.

Par ailleurs, `-SiteUrl` (1 site précis) est optimisé : en mode normal, le script utilise directement des appels Graph délégués pour ce seul site, ce qui aligne le temps de démarrage sur celui des autres commandes.

Pour la fiabilité avec GDAP, le script bascule automatiquement vers un bootstrap app-only pour les analyses d'un seul site lorsque `authMode=GDAP` est détecté (depuis `load.config.ps1`/le contexte du lanceur). Pour toujours le forcer, utilisez `-ForceAppOnlySingleSite`.

Pour les analyses complètes sous GDAP, le script utilise le même contexte de tenant client (`$global:cid`/`-TenantId`) pour `Connect-MgGraph` et pour le bootstrap de l'application temporaire, afin que le consentement et l'énumération des sites aient toujours lieu dans le bon tenant.

### Paramètres

| Paramètre | Description |
|---|---|
| `-SiteUrl` | Analyse 1 site précis. Une URL racine de tenant (par ex. `https://contoso.sharepoint.com`) déclenche automatiquement une analyse de tout le tenant |
| `-SkipVersions` | N'inclut pas l'historique des versions (plus rapide) |
| `-OutputPath` | Remplace le dossier de sortie par défaut (`C:\Temp\` / `~/Downloads/`) |
| `-TenantId` | ID de tenant Entra ID — détecté automatiquement s'il n'est pas indiqué ; obligatoire avec `-ClientId` |
| `-ClientId` | Client ID d'une App Registration existante — évite la création automatique ; à utiliser avec `-TenantId` et `-ClientSecret` ou `-CertificateThumbprint` |
| `-ClientSecret` | Client secret d'une app registration existante |
| `-CertificateThumbprint` | Empreinte du certificat d'une app registration existante |
| `-Apply` | Analyse récursive complète des bibliothèques, dossiers et fichiers. Sans ce switch, synthèse des quotas uniquement |
| `-UseHighPrivilege` | Mode auto : accorde temporairement `Sites.FullControl.All` au lieu de `Sites.Read.All` lorsque les droits en lecture seule s'avèrent insuffisants |
| `-RecycleBinOnly` | Ignore l'analyse du stockage/des bibliothèques — lit uniquement les éléments de la corbeille (stage 1 + stage 2) par site collection |
| `-ForceAppOnlySingleSite` | Force le bootstrap de l'application temporaire pour les analyses `-SiteUrl` (utile face aux restrictions GDAP/déléguées) |
| `-GraphTimeoutSec` | Timeout en secondes par appel Graph (défaut : `120`) |
| `-MaxGraphRetry` | Nombre maximal de nouvelles tentatives en cas de throttling/timeouts Graph (défaut : `6`) |
| `-VersionBatchConcurrency` | Nombre de workers `$batch` parallèles pour récupérer l'historique des versions, 1-8 (défaut : `4`) |
| `-MaxVersionRetryPasses` | Nombre maximal de passes de nouvelles tentatives pour l'historique des versions en cas de throttling persistant. `0` (défaut) s'adapte automatiquement au nombre de fichiers — SharePoint applique un plafond d'activité strict d'environ 1500-2500 recherches de versions résolues par passe, de sorte que sur des tenants comptant des centaines de milliers de fichiers, une valeur basse fixe (auparavant codée en dur à 8) abandonnait prématurément pour la majeure partie de l'analyse. Indiquez explicitement une valeur plus haute/basse pour remplacer l'ajustement automatique |
| `-Restart` | Supprime un checkpoint existant pour cette combinaison de paramètres et relance l'analyse depuis le début |

### Exemples

```powershell
# Synthèse rapide — quotas des sites uniquement, pas d'analyse des fichiers
.\Get-SharePointStorageReport.ps1

# Analyse complète du tenant, sous-sites et fichiers compris (app registration automatique)
.\Get-SharePointStorageReport.ps1 -Apply

# Idem, mais avec des droits d'application temporaires plus élevés si nécessaire
.\Get-SharePointStorageReport.ps1 -Apply -UseHighPrivilege

# Un seul site Teams précis
.\Get-SharePointStorageReport.ps1 -SiteUrl "https://contoso.sharepoint.com/teams/Operations" -Apply

# Analyse complète avec une app registration existante
.\Get-SharePointStorageReport.ps1 -Apply -ClientId "..." -TenantId "..." -ClientSecret "..."

# Ignorer une exécution interrompue et recommencer depuis le début
.\Get-SharePointStorageReport.ps1 -Apply -Restart
```

---

## Get-SharePointPermissionsReport.ps1

Indique **qui a accès à quoi** dans SharePoint Online — chaque site, sous-site, liste/bibliothèque, dossier et fichier doté de ses propres autorisations, exporté en CSV. Lecture seule : le script n'effectue que des appels `GET` et ne modifie jamais une autorisation.

### Couverture

- Administrateurs de site collection
- Attributions de rôles au niveau web (site et sous-site), héritage rompu compris
- Groupes SharePoint (Owners/Members/Visitors et groupes personnalisés) avec leur liste complète de membres
- Attributions de rôles sur les listes et bibliothèques de documents
- Niveau dossier et fichier : chaque élément avec `HasUniqueRoleAssignments`
- Liens de partage (anonyme / organisation / personnes spécifiques) et les personnes avec qui le partage a eu lieu
- Utilisateurs externes et invités, plus `Everyone` et `Everyone except external users`
- Groupes Entra ID, résolus en leur liste de membres transitive

L'héritage est suivi tel que SharePoint le modélise lui-même : un élément n'apparaît comme étendue (scope) propre que s'il possède ses propres autorisations. Le reste hérite du parent le plus proche, qui n'est signalé qu'une fois. Le CSV reste ainsi une carte de la structure des autorisations plutôt qu'une ligne par fichier.

> Les sites sont volontairement récupérés deux fois : d'abord à l'échelle du tenant via Graph (`getAllSites`), puis interrogés à nouveau pour les sous-sites via Graph et SharePoint REST (`/_api/web/webs`), et dédoublonnés par URL. Les sous-webs classiques que Graph ignore sont ainsi tout de même inclus.

### Authentification

La lecture des attributions de rôles n'est **pas** possible via Microsoft Graph et n'est pas non plus couverte par les rôles Read/Write/Manage de SharePoint : elle nécessite le rôle d'application `Sites.FullControl.All`. Le script vous connecte donc une fois de manière interactive, puis crée lui-même une App Registration de courte durée avec :

| Ressource | Rôle | Usage |
|---|---|---|
| SharePoint | `Sites.FullControl.All` | Attributions de rôles, groupes de site, étendues des éléments |
| Graph | `Sites.Read.All` | Énumération des sites à l'échelle du tenant |
| Graph | `GroupMember.Read.All` | Résolution de l'appartenance aux groupes Entra |

Cette application est supprimée à la fin. Malgré le rôle Full Control, le script n'écrit jamais rien. Si vous ne voulez pas d'application temporaire, indiquez `-ClientId` + `-TenantId` + `-CertificateThumbprint` d'une inscription existante qui possède déjà ces rôles.

> **Certificat, pas de secret — et ce n'est pas une préférence.** SharePoint Online refuse tout jeton app-only obtenu avec un client secret : vous obtenez `401` avec `x-ms-diagnostics: ... Unsupported app only token`. Seule l'authentification app-only par certificat fonctionne contre `_api`. L'application temporaire reçoit donc un certificat que le script crée **en mémoire** et enregistre sur l'application ; il ne va ni dans le magasin de certificats ni sur le disque, il n'y a donc rien à nettoyer ensuite. Si vous passez `-ClientSecret` avec votre propre application, le script vous avertit : la partie Graph fonctionnera, la partie SharePoint non.

#### Pourquoi pas simplement Graph ?

Graph en fait une partie : sur un `driveItem`, `/permissions` renvoie les autorisations, les liens de partage (avec type et date d'expiration) et, via `inheritedFrom`, si l'héritage a été rompu. Mais le reste du tableau y manque tout simplement — il n'existe aucun point de terminaison Graph pour :

| Quoi | Graph | SharePoint REST |
|---|---|---|
| Attributions de rôles au niveau site/web | ❌ n'existe pas | ✅ `/_api/web/roleassignments` |
| Groupes SharePoint et leurs membres | ❌ n'existe pas | ✅ `/_api/web/sitegroups` |
| Administrateurs de site collection | ❌ n'existe pas | ✅ `/_api/web/siteusers` |
| Nom du niveau d'autorisation (Full Control, Modification, niveaux personnalisés) | ❌ uniquement `read`/`write`/`owner` | ✅ `RoleDefinitionBindings` |
| Listes sans `driveItem` (listes ordinaires) | ❌ | ✅ |
| Filtrage peu coûteux sur les autorisations uniques | ❌ un appel par élément | ✅ `HasUniqueRoleAssignments` en un seul passage |

Cette dernière ligne est aussi une différence de vitesse : via Graph, il faudrait un appel `/permissions` pour *chaque* fichier, alors que SharePoint indique en un seul passage par liste *quels* éléments ont leurs propres autorisations. Une variante uniquement Graph *pourrait* fonctionner avec un client secret et moins de droits (`Sites.Read.All`), mais produirait un rapport sans propriétaires de site, sans groupes et sans niveaux d'autorisation — précisément là où commence une revue des autorisations.

### Sortie

| Fichier | Contenu |
|---|---|
| `SharePoint_Permissions_Detail_<ts>.csv` | Une ligne par attribution : étendue, principal, niveaux d'autorisation, type de lien de partage, externe oui/non, nombre de membres. Les étendues illisibles apparaissent comme `ItemType = Error` avec la raison dans la colonne `Error` ; `UnitKey` relie une ligne à l'unité de checkpoint qui l'a écrite |
| `SharePoint_Permissions_Summary_<ts>.csv` | Par site : nombre d'attributions, étendues uniques, webs, listes, dossiers/fichiers avec autorisations propres, liens de partage, liens anonymes, principals externes, attributions `Everyone` |
| `SharePoint_Permissions_SiteAccess_<ts>.csv` | **Par site, une ligne par personne**, avec le groupe par lequel passe l'accès et le niveau. Voir ci-dessous |
| `SharePoint_Permissions_Groups_<ts>.csv` | Par groupe, une ligne par membre — groupes SharePoint, groupes Entra imbriqués dedans, *et* groupes Entra attribués directement sur une étendue. Le tout aplati en personnes |
| `SharePoint_Permissions_EffectiveAccess_<ts>.csv` | Uniquement avec `-IncludeEffectiveAccess` : une ligne par utilisateur et par étendue, avec le groupe par lequel passe cet accès |

> Sur un grand tenant, le fichier d'accès effectif peut être plus volumineux que le CSV détaillé de plusieurs ordres de grandeur. C'est pourquoi il est désactivé sauf si vous le demandez.

#### Qui a accès où, via quel groupe — dans un seul onglet

La question avec laquelle on ouvre généralement ce rapport n'est pas « quelles attributions existent » mais **« qui peut accéder à ce SharePoint, et comment y est-il arrivé »**. Auparavant, c'était dispersé : `Rechten` indiquait *qu*'un groupe avait des autorisations, `Groepen` indiquait qui en faisait partie, et vous deviez faire le lien vous-même. « Site Owners a Contrôle total » plus « Site Owners contient cinq personnes » n'est pas encore une réponse.

D'où `SharePoint_Permissions_SiteAccess_<ts>.csv` (onglet `Toegang`) : **une ligne par personne et par site**, avec le groupe par lequel passe l'accès et le niveau.

| Colonne | Contenu |
|---|---|
| `SiteTitle` / `SiteUrl` | Le site, par son nom — 130 URL ne se lisent pas « d'un coup d'œil » |
| `UserDisplayName` / `UserPrincipalName` / `UserEmail` | Qui |
| `IsExternal` / `AccountEnabled` | Invité ou interne, compte actif |
| `ViaType` | `Direct`, `SharePointGroup`, `SecurityGroup`, `M365Group`, `Everyone`, … |
| `ViaName` / `ViaId` | Quel groupe, ou `(direct toegekend)`. L'id est indiqué parce qu'un titre comme `Site Owners` existe sur chaque site |
| `PermissionLevels` | Le niveau de cette attribution |

C'est volontairement **consolidé par site collection** : une personne qui aboutit à trente dossiers du même site via le même groupe représente une ligne — pas trente. Un autre niveau ou un autre groupe constitue *bien* une ligne distincte, car c'est un accès différent. Pour le voir par dossier ou fichier individuel, utilisez `-IncludeEffectiveAccess` ; cet onglet (`Effectief`) est par étendue et donc beaucoup plus volumineux.

Trois éléments qui ne disparaissent volontairement pas ici :

- **Les personnes attribuées directement** apparaissent en tant que telles, avec `ViaType = Direct` et `ViaName = (direct toegekend)`.
- **`Everyone` et `Everyone except external users`** ne se résolvent en aucune personne, mais c'est justement ce que vous voulez voir. Ils obtiennent une ligne avec la revendication (claim) comme nom.
- Même avec `-SkipGroupExpansion`, les personnes attribuées directement restent visibles ; seuls les membres des groupes manquent alors.

> Comparé à [NovaPoint](https://github.com/Barbarur/NovaPoint/wiki/Solution-Report-PermissionsReport), qui répond à la même question avec `AccessType` + `GroupId` et une colonne `Users` contenant une liste d'utilisateurs : ici, chaque utilisateur a sa propre ligne. C'est moins compact à lire, mais c'est la différence entre pouvoir filtrer ou pivoter sur une personne, ou non.

#### Tout dans un seul fichier Excel

Avec `-Excel`, vous obtenez en plus des CSV un classeur unique, `SharePoint_Permissions_<ts>.xlsx`, avec un onglet par rapport :

| Onglet | Contenu |
|---|---|
| `Samenvatting` | Par site : attributions, étendues uniques, liens de partage, principals externes, attributions `Everyone`, erreurs |
| `Rechten` | Chaque attribution individuellement |
| `Toegang` | **Par site, par personne : quel droit et via quel groupe.** L'onglet par lequel commencer |
| `Groepen` | Chaque groupe avec ses membres — groupes SharePoint, groupes Entra imbriqués dedans, **et** groupes Entra attribués directement sur une étendue |
| `Effectief` | Uniquement avec `-IncludeEffectiveAccess` : une ligne par utilisateur et par étendue |

Chaque onglet est un vrai tableau Excel, avec boutons de filtre et ligne d'en-tête figée. Les nombres arrivent en tant que nombres, pas en tant que texte : additionner et trier fonctionne sans conversion préalable.

#### Tableaux croisés dynamiques

Cinq tableaux croisés dynamiques prêts à l'emploi sont ajoutés, chacun sur son propre onglet :

| Onglet | Lignes | Colonnes | Valeur | Filtres |
|---|---|---|---|---|
| `Pivot rechten` | Site | Niveau d'autorisation | Nombre d'attributions | Type de principal, type d'étendue |
| `Pivot principals` | Principal | Type d'étendue | Nombre d'étendues | Site, externe oui/non |
| `Pivot groepen` | Groupe | Membre externe oui/non | Nombre de membres | Site, type de groupe |
| `Pivot toegang` | **Site → groupe → personne** | Niveau d'autorisation | Nombre | Externe oui/non, type d'accès |
| `Pivot per persoon` | Personne → site → groupe | Niveau d'autorisation | Nombre | Externe oui/non, type d'accès |

`Pivot toegang` suit la façon dont SharePoint distribue réellement les autorisations : un site a des groupes, et les groupes ont des personnes. Replié, vous voyez quels groupes sont présents sur un site ; déplié, qui ces groupes laissent entrer. `Pivot per persoon` lit les mêmes données dans l'autre sens — qu'atteint *cette* personne et par quel biais — ce qui est la question lors d'un offboarding.

> Une personne attribuée directement n'a pas de groupe. `ViaName` contient alors `(direct toegekend)` au lieu de rien : un niveau vide dans la hiérarchie se lit comme une donnée manquante, pas comme « attribué sans groupe ».

> **`PermissionLevels` ne se prête pas au tableau croisé, `PrimaryPermission` si.** SharePoint attribue souvent plusieurs niveaux à la fois, et ceux-ci figurent dans une seule colonne sous la forme `Read; Limited Access`. Un tableau croisé en fait une valeur distincte, si bien que `Full Control` et `Full Control; Limited Access` se retrouvent sur des lignes différentes. C'est pourquoi les onglets `Rechten` et `Effectief` ont une colonne supplémentaire `PrimaryPermission` à côté du texte complet, avec le niveau le plus élevé de cette attribution. `Limited Access` perd toujours face à un vrai niveau — SharePoint le pose lui-même pour que quelqu'un puisse naviguer jusqu'à un élément attribué plus en profondeur. Un niveau d'autorisation personnalisé pèse plus que `Lezen` (Lecture) mais moins que `Volledig beheer` (Contrôle total) : il a été créé volontairement, il ne doit donc pas disparaître. Les noms de niveaux néerlandais et anglais sont tous deux reconnus.

Pour créer vous-même un tableau croisé dynamique : placez le curseur dans un onglet et choisissez **Insertion → Tableau croisé dynamique** ; le tableau est déjà nommé, la plage est donc correcte d'emblée et s'étend avec les données.

Les CSV sont toujours conservés ; le classeur vient en plus. C'est voulu : les CSV sont la destination du flux de l'analyse et ce que complète une exécution reprise, ils existent donc de toute façon — et si l'écriture du classeur échoue (module manquant, fichier ouvert dans Excel, mémoire insuffisante), cela ne vous coûte jamais le rapport lui-même.

> **Limite de lignes.** Une feuille Excel s'arrête à 1 048 576 lignes et abandonne le reste sans avertissement. Le script tronque donc volontairement à 1 000 000 et indique quel onglet a été raccourci et quel CSV contient les données complètes. Seul `Effectief` s'en approche réalistement sur un grand tenant.

`-Excel` nécessite le module `ImportExcel` (présent dans `Install-Modules.ps1`). S'il manque, le script le signale et les CSV sont conservés normalement.

### Reprise après interruption (checkpoints)

Après chaque liste terminée, un checkpoint est écrit dans le dossier de sortie : `SharePoint_Permissions_<hash>.state.json`, `.keys.partial.log` et `.detail/.groups/.effective.partial.csv`. Contrairement aux deux scripts ci-dessus, les lignes sont envoyées directement dans ces fichiers partiels au lieu d'être conservées en mémoire — une exécution à l'échelle du tenant au niveau des éléments produit des millions de lignes. Le CSV de synthèse n'existe donc pas comme checkpoint : il est construit à la fin à partir du fichier détaillé, de sorte qu'une exécution reprise résume tout ce qui a jamais été écrit pour ce `<hash>`, et pas seulement la partie de la dernière session. Le `<hash>` provient des paramètres d'analyse ; relancer avec les mêmes paramètres reprend donc à partir de la dernière liste terminée. `-Restart` supprime ce checkpoint et recommence. Les fichiers de checkpoint ne sont nettoyés qu'une fois les CSV définitifs sur disque — s'ils sont toujours là, l'exécution précédente a été interrompue.

Les unités déjà terminées sont consignées dans `.keys.partial.log`, une ligne par clé, en ajout seul. Ce n'est volontairement pas une liste dans le JSON : la réécrire triée après chaque liste représente un travail quadratique, et sur un tenant comptant des milliers de listes, le checkpoint prendrait alors plus de temps que l'analyse elle-même. Une dernière ligne à moitié écrite (processus tué brutalement pendant l'écriture) est ignorée — cette unité est simplement analysée à nouveau.

### Robustesse

Une exécution à l'échelle du tenant dure des heures et touche des milliers d'objets ; les incidents ci-dessous ne sont donc pas des cas limites, mais sont à prévoir. Voici comment le script les gère :

| Situation | Comportement |
|---|---|
| Rôle d'application pas encore répliqué | Avant l'analyse, le jeton doit prouver *qu*'il porte les rôles (revendication `roles`). Entra émet en effet sans problème un jeton sans le rôle qui vient d'être attribué, et Graph y répond par `401` — pas `403`. Un tel jeton n'est pas mis en cache ; un nouveau jeton est émis jusqu'à ce que le rôle y figure (jusqu'à ~2,5 min), puis une erreur claire est levée |
| Jeton refusé (`401`) | Nouvelle authentification une fois avec un jeton neuf ; s'il est toujours refusé, **l'exécution s'arrête** avec la raison issue de `x-ms-diagnostics`. Un 401 ne concerne jamais un seul site, il n'est donc pas signalé par site |
| Pas d'accès à un site (`403`) ou objet disparu (`404`) | Ce site/cet objet est ignoré, le reste continue |
| Throttling (`429`/`503`) | Nouvelle tentative avec `Retry-After`, sinon backoff exponentiel jusqu'à 3 minutes |
| Une liste échoue (seuil d'affichage, modèle inhabituel) | Ligne d'erreur dans le CSV détaillé, les autres listes de ce site continuent normalement. L'unité n'est **pas** marquée comme terminée, une exécution reprise la retente donc |
| Liste avec autorisations propres mais sans attributions de rôles | Ne produit aucune ligne, et c'est correct. Auparavant, cela plantait sur `Cannot bind argument to parameter 'RoleAssignments'` |
| Attributions de rôles illisibles (`403`) | **Ligne d'erreur, pas un résultat vide.** C'est le seul endroit où un 403 n'est pas ignoré : une liste vide d'attributions de rôles se lit comme « personne n'a d'autorisation sur cette étendue », et « je n'ai pas le droit de regarder » est un fait différent de « il n'y a rien à voir » |
| Une liste de galerie refuse la sélection de champs (`400`) | La requête est réduite progressivement (quatre variantes) jusqu'à ce que SharePoint l'accepte. `Galerie van thema's` et `Galerie met basispagina's` n'ont pas tous les champs ; mieux vaut perdre le nom de fichier que les étendues uniques de cette liste. La variante qui fonctionne relève du **modèle** de liste : elle est apprise une fois puis réutilisée sur chaque site suivant — sur un locataire de 131 sites, cela divise par deux les allers-retours, et le journal signale chaque particularité une fois au lieu d'une fois par site |
| `Lijst met gebruikersgegevens` (modèle 112) | L'analyse des éléments est ignorée. SharePoint refuse `/items` sur cette liste système masquée quelle que *soit* la largeur de champs, et les éléments sont des fiches utilisateur, pas du contenu — les autorisations d'élément n'y signifient rien. La liste elle-même est bien signalée. Cela a évité 121 fausses lignes d'erreur par analyse de tenant |
| Lignes d'erreur d'une tentative précédente interrompue | Sont omises à l'écriture dès que la même unité a *finalement* réussi. Sinon, le rapport compte des erreurs déjà résolues, et le nombre « N étendues illisibles » est faux — précisément le nombre sur lequel quelqu'un agit |
| Le balayage des éléments échoue | Ligne d'erreur distincte : sans elle, la liste semblerait ne contenir aucun élément avec des autorisations propres |
| Le CSV est ouvert dans Excel | Cinq nouvelles tentatives avec un délai croissant ; en cas d'échec persistant, l'exécution s'arrête au lieu de perdre silencieusement des lignes |
| Le jeton expire au milieu d'une grande bibliothèque | Chaque vague récupère à nouveau le jeton — les workers reçoivent une copie et ne verraient pas un rafraîchissement ultérieur |
| SharePoint répète un lien de pagination | Détecté et interrompu au lieu de tourner indéfiniment |
| Erreur inattendue, où que ce soit | Un `trap` nettoie l'App Registration temporaire avant l'arrêt du script — aucune application Full Control n'est jamais laissée derrière |

> À la fin, le script indique explicitement si *tout* a pu être lu. En cas de `[WARN]` concernant des étendues illisibles, filtrez le CSV détaillé sur `ItemType = Error` : un rapport lacunaire ne doit pas ressembler à un rapport sans constat.

### Paramètres

| Paramètre | Type | Défaut | Description |
|---|---|---|---|
| `-TenantUrl` | string | — | Racine du tenant, par ex. `https://contoso.sharepoint.com`. Obligatoire pour une exécution à l'échelle du tenant |
| `-SiteUrl` | string | — | Une site collection (sous-sites compris) au lieu de tout le tenant |
| `-Scope` | `Site`/`List`/`Item` | `Item` | Profondeur : webs uniquement, webs + listes, ou tout jusqu'au niveau dossier et fichier |
| `-TenantId` | string | _(de la session)_ | ID de tenant Entra. Obligatoire avec `-ClientId` |
| `-ClientId` | string | — | App Registration existante ; évite l'application temporaire |
| `-ClientSecret` | string | — | Secret pour `-ClientId`. **Ne fonctionne pas contre SharePoint** (voir Authentification) ; le script vous avertit |
| `-CertificateThumbprint` | string | — | Certificat pour `-ClientId`, depuis `Cert:\CurrentUser\My` ou `Cert:\LocalMachine\My`. C'est la variante qui fonctionne |
| `-OutputPath` | string | `C:\Temp` | Dossier de sortie |
| `-IncludeOneDriveSites` | switch | désactivé | Inclut aussi les sites OneDrive personnels (un site par utilisateur) |
| `-IncludeHiddenLists` | switch | désactivé | Inclut les listes masquées et système (Form Templates, Style Library, historique de workflow, …) |
| `-ListTitle` | string[] | _(tout)_ | Limite à un ou plusieurs titres de liste/bibliothèque |
| `-ExcludeLimitedAccess` | switch | désactivé | Omet les attributions `Limited Access`. SharePoint les pose lui-même pour que quelqu'un puisse naviguer vers un élément attribué plus en profondeur — du bruit dans la plupart des revues, mais elles expliquent *bien* pourquoi quelqu'un voit un chemin de dossier |
| `-SkipGroupExpansion` | switch | désactivé | Ne résout pas l'appartenance aux groupes. Plus rapide, mais vous savez alors seulement *quel* groupe a accès, pas qui en fait partie |
| `-IncludeEffectiveAccess` | switch | désactivé | Écrit en plus le CSV d'accès effectif |
| `-Excel` | switch | désactivé | Écrit en plus un unique `.xlsx` avec un onglet par rapport. Nécessite `ImportExcel` |
| `-GraphTimeoutSec` | int | `120` | Timeout par appel Graph/SharePoint |
| `-MaxGraphRetry` | int | `6` | Nombre de nouvelles tentatives en cas de throttling ou de timeouts |
| `-Concurrency` | int (1-8) | `4` | Workers parallèles pour les recherches par élément qui dominent une exécution `-Scope Item`. `1` désactive le parallélisme |
| `-Restart` | switch | désactivé | Ignore un checkpoint existant et recommence |

### Prérequis

- Module `Microsoft.Graph.Authentication` (et `Microsoft.Graph.Applications` tant que vous laissez le script créer l'application temporaire) — `.\scripts\Startup\Install-Modules.ps1`
- Un compte autorisé à créer une App Registration et à accorder le consentement administrateur, sauf si vous passez le `-ClientId` d'une application existante

### Exemples

```powershell
# Tout, à l'échelle du tenant : sites, sous-sites, bibliothèques, dossiers et fichiers avec autorisations propres
.\Get-SharePointPermissionsReport.ps1 -TenantUrl "https://contoso.sharepoint.com"

# Une site collection, plus un CSV d'accès effectif par utilisateur
.\Get-SharePointPermissionsReport.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/Finance" -IncludeEffectiveAccess

# Tout dans un classeur Excel : synthèse, autorisations, groupes avec membres et accès effectif
.\Get-SharePointPermissionsReport.ps1 -TenantUrl "https://contoso.sharepoint.com" -IncludeEffectiveAccess -Excel

# Aperçu plus rapide : s'arrêter au niveau liste/bibliothèque et masquer les attributions de traversée automatiques
.\Get-SharePointPermissionsReport.ps1 -TenantUrl "https://contoso.sharepoint.com" -Scope List -ExcludeLimitedAccess

# Ignorer une analyse de tenant interrompue et recommencer depuis le début
.\Get-SharePointPermissionsReport.ps1 -TenantUrl "https://contoso.sharepoint.com" -Restart
```

---

## Remove-SharePointFileVersionsByDate.ps1

Signale ou supprime les **anciennes versions de fichiers** dans les bibliothèques de documents SharePoint Online à partir d'une date limite, tout en **conservant la version actuelle**.

### Comportement

- Par défaut : aperçu/rapport uniquement
- Avec `-Apply` : supprime réellement les versions antérieures correspondantes
- Fonctionne sur un site ou sur tous les sites du tenant
- Par défaut, ni sites OneDrive ni bibliothèques masquées
- Basé sur Microsoft Graph (`Invoke-MgGraphRequest`) — **pas** de `PnP.PowerShell` et **pas** besoin de votre propre app registration Entra dans le cas par défaut

### Authentification

Par défaut, le script se connecte de manière interactive (déléguée) avec `Sites.ReadWrite.All` + `Files.ReadWrite.All` via `Connect-MgGraph` — qui utilise l'application pré-consentie de Microsoft, donc sans App Registration ni `-ClientId` propres. Seule une **analyse à l'échelle du tenant** (sans `-SiteUrl`) nécessite en plus une App Registration temporaire, en lecture seule et de courte durée (`Sites.Read.All`) pour l'énumération des sites/bibliothèques *et* la récupération de l'historique des versions — Microsoft ne prend pas en charge l'énumération des sites du tenant en mode délégué. Avec `-VersionBatchConcurrency` supérieur à `1` (le défaut), une **deuxième** App Registration temporaire est également créée, uniquement pour doubler le débit des recherches de versions : le throttling « activityLimitReached » de SharePoint s'applique par app registration, donc deux applications disposent chacune de leur propre budget (même approche que `Get-SharePointStorageReport.ps1`). Les deux applications temporaires sont supprimées à la fin. Les **suppressions** de versions passent toujours par vos propres autorisations déléguées, jamais par une application temporaire.

Vous voulez vous passer de la ou des applications temporaires et utiliser votre propre app registration existante ? Indiquez alors `-ClientId` + `-TenantId` + `-ClientSecret` (ou `-CertificateThumbprint`) ; cette application doit déjà disposer de l'autorisation d'application `Sites.ReadWrite.All`.

> **Attention :** la suppression d'une version précise (`DELETE .../versions/{id}`) ne figure pas dans la référence officielle de l'API Graph de Microsoft, mais c'est une opération largement utilisée et dont le fonctionnement est confirmé (pour les bibliothèques de documents OneDrive comme SharePoint). La version actuelle/la plus récente ne peut pas être supprimée ainsi — Graph le refuse, ce qui constitue précisément la garantie de conservation de la version actuelle.

### Reprise après interruption (checkpoints) et progression

Comme `Get-SharePointStorageReport.ps1`, ce script écrit un checkpoint dans le dossier de sortie après chaque bibliothèque terminée : `SharePoint_VersionCleanup_<hash>.state.json` + `.summary.partial.csv` + `.detail.partial.csv`. Le `<hash>` est dérivé de tous les paramètres d'analyse (date limite, site, mode, dossier de sortie, etc.) :

- **Relancer avec les mêmes paramètres** reprend automatiquement — les bibliothèques déjà traitées sont ignorées (`[SKIP] Already completed in a previous run.`).
- **`-Restart`** supprime le checkpoint et recommence depuis le début. Cela n'affecte que le suivi de la progression, pas ce qui (avec `-Apply`) a déjà été réellement supprimé dans SharePoint — les versions supprimées restent bien entendu supprimées.
- Les fichiers de checkpoint sont nettoyés automatiquement dès que l'analyse se termine avec succès.

Pendant l'analyse, le script affiche des barres de progression imbriquées (sites → bibliothèques → analyse des dossiers/fichiers / récupération de l'historique des versions) à côté des lignes de journal défilantes, ainsi qu'un message `[WAIT] throttled by Microsoft Graph — waiting ...` dès que Graph applique un throttling, pour qu'une longue pause ne donne pas l'impression d'un blocage.

> **Attention (appels délégués/SDK) :** comme pour `Get-SharePointStorageReport.ps1`, le script exécute `Set-MgRequestContext -ClientTimeout <-GraphTimeoutSec> -MaxRetry 0` juste après la connexion — sans ce réglage, les cmdlets du SDK Microsoft.Graph relancent elles-mêmes silencieusement les 429/503 avec leur propre backoff, ce qui en cas de `activityLimitReached` peut provoquer plusieurs minutes de silence sans que le message `[WAIT]` du script n'apparaisse.

### Paramètres

| Paramètre | Type | Description |
|---|---|---|
| `-BeforeDate` | `datetime` | Supprime les versions antérieures à cette date |
| `-SiteUrl` | `string` | Facultatif : analyse un seul site |
| `-TenantUrl` | `string` | Obligatoire pour une analyse de tous les sites, par ex. `https://contoso.sharepoint.com` |
| `-TenantId` | `string` | ID de tenant Entra ID — détecté automatiquement s'il n'est pas indiqué ; obligatoire avec `-ClientId` |
| `-ClientId` | `string` | Client ID d'une App Registration existante — évite l'application temporaire ; à utiliser avec `-TenantId` et `-ClientSecret` ou `-CertificateThumbprint` |
| `-ClientSecret` | `string` | Client secret d'une app registration existante |
| `-CertificateThumbprint` | `string` | Empreinte du certificat d'une app registration existante |
| `-Apply` | `switch` | Effectue réellement la suppression |
| `-IncludeOneDriveSites` | `switch` | Inclut les sites OneDrive dans l'analyse du tenant |
| `-IncludeHiddenLibraries` | `switch` | Inclut les bibliothèques de documents masquées |
| `-LibraryTitle` | `string[]` | Filtre facultatif sur le titre de la bibliothèque |
| `-GraphTimeoutSec` | `int` | Timeout en secondes par appel Graph (défaut : `120`) |
| `-MaxGraphRetry` | `int` | Nombre maximal de nouvelles tentatives en cas de throttling/timeouts Graph (défaut : `6`) |
| `-VersionBatchConcurrency` | `int` | Nombre de workers `$batch` parallèles pour récupérer l'historique des versions dans les analyses de tenant, 1-8 (défaut : `4`). Au-delà de `1`, la deuxième application temporaire est également créée (voir Authentification) |
| `-MaxVersionRetryPasses` | `int` | Nombre maximal de passes de nouvelles tentatives pour récupérer les listes de versions en cas de throttling persistant. `0` (défaut) s'adapte automatiquement au nombre de fichiers — même approche et même raison que pour `Get-SharePointStorageReport.ps1` ci-dessus |
| `-Restart` | `switch` | Supprime un checkpoint existant pour cette combinaison de paramètres et relance l'analyse depuis le début |

### Exemples

```powershell
# Aperçu à l'échelle du tenant : tout ce qui est antérieur au 1er janvier 2025
.\Remove-SharePointFileVersionsByDate.ps1 `
    -TenantUrl "https://contoso.sharepoint.com" `
    -BeforeDate "2025-01-01"

# Suppression réelle sur un seul site
.\Remove-SharePointFileVersionsByDate.ps1 `
    -SiteUrl "https://contoso.sharepoint.com/sites/Finance" `
    -BeforeDate "2025-01-01" `
    -Apply

# Ignorer une analyse de tenant interrompue et recommencer depuis le début
.\Remove-SharePointFileVersionsByDate.ps1 `
    -TenantUrl "https://contoso.sharepoint.com" `
    -BeforeDate "2025-01-01" `
    -Apply -Restart
```
