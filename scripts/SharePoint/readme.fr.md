[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../readme.fr.md) › [scripts](../readme.fr.md) › **SharePoint**

# SharePoint Scripts

Opérations sur le contenu SharePoint Online et OneDrive via PnP PowerShell, avec connexion
administrateur interactive sur n'importe quel tenant (client).

---

## Dossiers

| Dossier | Description |
|--------|-------------|
| [`Provisioning/`](Provisioning/readme.fr.md) | Provisionner et maintenir une structure complète — modèle de métadonnées, types de contenu, bibliothèques et autorisations de groupe — à partir d'un seul fichier de configuration, avec un audit du partage et un contrôle de dérive |

## Scripts

| Script | Description |
|--------|-------------|
| [`Find-SiteContent.ps1`](Find-SiteContent.ps1) ([docs](#find-sitecontentps1)) | Rechercher dans tout un site (nom, chemin, type, taille, date ou texte intégral) et indiquer les autorisations de chaque résultat — PnP/CSOM, se connecte avec votre compte |
| [`Search-SharePointContent.ps1`](Search-SharePointContent.ps1) ([docs](#search-sharepointcontentps1)) | La même question à l'échelle du tenant via Microsoft Graph, en app-only, sans connexion interactive — fichiers et dossiers |
| [`Restore-RecycleBinItems.ps1`](Restore-RecycleBinItems.ps1) ([docs](#restore-recyclebinitemsps1)) | Restaurer des fichiers/dossiers supprimés depuis la corbeille d'un site ou d'un OneDrive (essai à blanc par défaut) |
| [`Trace-SharePointFile.ps1`](Trace-SharePointFile.ps1) ([docs](#trace-sharepointfileps1)) | Où est passé un fichier ? Renommages, déplacements, copies et suppressions depuis le journal d'audit, y compris via un dossier — en heure de Bruxelles, sur une période au choix |
| [`Revoke-SharePointUserAccess.ps1`](Revoke-SharePointUserAccess.ps1) ([docs](#revoke-sharepointuseraccessps1)) | Retirer partout l'accès d'un utilisateur : administrateur de site collection, attributions directes à tous les niveaux, groupes SharePoint et liens de partage. Rapport par défaut, suppression avec `-Apply` |
| [`Test-SharePointAccessScripts.ps1`](Test-SharePointAccessScripts.ps1) ([docs](#test-sharepointaccessscriptsps1)) | Vérifier les deux scripts d'accès sans toucher à un tenant — bloc d'authentification partagé identique, et l'entonnoir de révocation se comporte correctement |

---

## Lequel des deux scripts de recherche ?

Les deux répondent à « où cela se trouve-t-il et qui peut y accéder », et les deux écrivent le
même type de CSV. Ils diffèrent par ce qu'ils peuvent voir et par ce qu'ils attendent de vous.

| | `Find-SiteContent.ps1` (PnP) | `Search-SharePointContent.ps1` (Graph) |
|---|---|---|
| Connexion | Interactive, avec votre compte | App-only, sans surveillance — convient à une tâche planifiée |
| Droits nécessaires | Accès au site, ou `-GrantSiteAdmin` par site | Un consentement administrateur, une seule fois, pour tout le tenant |
| Portée | Une site collection (+ sous-sites) | Un site, ou **chaque site du tenant**, OneDrive compris |
| Fichiers et dossiers | Oui | Oui, et plus vite — `delta` lit une bibliothèque par pages de mille et les autorisations arrivent par 20 dans un `$batch` |
| Éléments des listes ordinaires | Oui, avec leurs autorisations | **Non** — Graph n'expose les autorisations que pour les driveItems |
| Droits au niveau site et liste | Oui : « ceci hérite de la bibliothèque, qui accorde Modification à Site Members » | **Non** — Graph n'a pas d'API pour les attributions de rôles SharePoint ; il indique *si* un élément hérite et d'où, pas ce que le site accorde |
| Liens de partage | Oui, à partir des groupes `SharingLinks.*` | Oui, plus riches : portée du lien, modification/lecture, date d'expiration et l'URL du lien elle-même |

Règle générale : **Graph** pour « trouve-le n'importe où dans le tenant et montre-moi les liens
et les invités dessus », **PnP** lorsque vous avez besoin de l'histoire complète des autorisations
d'un site, y compris ses listes et ses groupes.

---

### Find-SiteContent.ps1

Répond aux deux questions que l'on se pose généralement en même temps : **où cela se trouve-t-il**
et **qui peut y accéder**. Lecture seule — le script ne modifie jamais rien.

**Deux moteurs**

| Moteur | Quand | Ce qu'il voit |
|--------|------|--------------|
| Crawl (par défaut) | Pas de `-Content` indiqué | Parcourt chaque liste et bibliothèque. `-Name` et `-ItemType` sont placés dans une requête CAML lorsque c'est possible, pour que SharePoint ne renvoie que les correspondances ; ce que CAML ne peut pas exprimer se rabat sur la lecture complète de cette liste. Voit tout, y compris ce que l'index de recherche n'a pas encore pris en compte |
| Search (`-Content`) | Requête en texte intégral | Une requête KQL sur l'index de recherche, limitée au chemin du site — c'est celle qui trouve du texte *à l'intérieur* des documents. Rapide, mais limitée à ce qui est indexé et à ce que le compte connecté peut voir. Couvre automatiquement les sous-sites |

Les deux moteurs alimentent les mêmes filtres : `-Name` (caractères génériques), `-Path`,
`-Extension`, `-ItemType`, `-ListName`, `-ModifiedBy`, `-ModifiedAfter` / `-ModifiedBefore`,
`-MinSizeMB`. `-Name` est comparé au nom du fichier, au titre de l'élément *et* au dernier
segment de l'URL, de sorte qu'un élément dont le titre diffère du nom de fichier est quand même
trouvé.

Les bibliothèques masquées et système sont ignorées sauf si vous passez `-IncludeHidden`, et
pendant le crawl seul le web de premier niveau est parcouru sauf si vous ajoutez
`-IncludeSubsites` — le script indique par web combien de listes il a ignorées et pourquoi.
**`-Everything` désactive tout cela d'un coup** : chaque sous-site, chaque liste masquée et
système, aucune limite sur les résultats ni sur les recherches d'autorisations. Un crawl ne
compare toujours que les noms et les métadonnées ; utilisez `-Content` pour chercher dans les
documents eux-mêmes.

**Autorisations par résultat**

Pour chaque correspondance, le script détermine d'où viennent réellement les autorisations :

| Source | Signification |
|--------|---------|
| `Item` | L'élément a rompu l'héritage et porte ses propres attributions de rôles |
| `List` | Il hérite d'une bibliothèque/liste qui a des autorisations uniques |
| `Site` | Il hérite jusqu'au (sous-)site |

Les attributions de rôles sont aplaties en une ligne CSV par principal — type de principal,
identifiant, e-mail et noms de rôles (`Full Control`, `Edit`, …). `Limited Access` est masqué
sauf si vous passez `-IncludeLimitedAccess` ; ces entrées n'existent que pour permettre à
quelqu'un d'atteindre un élément plus profond et n'accordent rien par elles-mêmes.

Trois éléments sont signalés à part, car ce sont eux qui surprennent :

- **Liens de partage** — les groupes `SharingLinks.*` derrière chaque « Copier le lien ».
  Toujours développés jusqu'aux personnes qu'ils contiennent et étiquetés *Anyone* /
  *Organization* / *Specific people*, pour qu'un lien anonyme ne puisse pas se cacher dans le bruit
- **Utilisateurs externes** — comptes invités (`#ext#`) dans n'importe quelle attribution
- **Everyone** — « Everyone » et « Everyone except external users »

Les autorisations de site et de liste sont lues une seule fois puis mises en cache, et les
autorisations d'élément uniquement pour les éléments qui ont réellement rompu l'héritage : une
recherche avec une poignée de résultats coûte donc une poignée d'appels supplémentaires.
`-Permissions Unique` est le moyen rapide de répondre à « qu'est-ce qui, dans ce site, est partagé
différemment du reste » ; `-Permissions None` ignore complètement les autorisations.
`-MaxPermissionLookups` (1000 par défaut) empêche une recherche trop large de tourner pendant des
heures.

**Connexion**

Comme pour `Restore-RecycleBinItems.ps1` : la première exécution sur un tenant enregistre une
application Entra public-client (`AllSites.FullControl` délégué, avec consentement
administrateur) et met en cache le client ID par tenant dans `pnp.appid.json` à la racine du dépôt
(gitignored). Un client ID déjà mis en cache par l'autre script est réutilisé, donc cela ne coûte
généralement rien. Passez `-ClientId` pour ignorer entièrement l'enregistrement de l'application.

La lecture des autorisations nécessite un accès au site. `-GrantSiteAdmin` fait de
l'administrateur connecté un administrateur de site collection pendant la durée de l'exécution,
puis retire ces droits (conservez-les avec `-KeepSiteAdmin`) — c'est ce qui permet de chercher
dans le OneDrive de quelqu'un d'autre ou dans un site dont vous n'êtes pas membre.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-SiteUrl` | Oui | Site collection à parcourir (site d'équipe, site de communication ou OneDrive) |
| `-IncludeSubsites` | Non | Parcourt aussi chaque sous-site en dessous |
| `-Content` | Non | Requête texte intégral/KQL — bascule vers l'index de recherche |
| `-Name` | Non | Filtre sur le nom, caractères génériques autorisés (`*offerte*`) — comparé au nom du fichier, au titre de l'élément *et* au dernier segment de l'URL |
| `-Path` | Non | Filtre sur le dossier, correspondance partielle sur l'URL relative au serveur |
| `-Extension` | Non | Une ou plusieurs extensions, avec ou sans le point (`xlsx`,`pdf`) |
| `-ItemType` | Non | `All` (défaut), `File`, `Folder` ou `ListItem` |
| `-ListName` | Non | Uniquement ces listes/bibliothèques par titre, caractères génériques autorisés |
| `-ModifiedBy` | Non | Auteur de la dernière modification — nom d'affichage ou e-mail, caractères génériques autorisés |
| `-ModifiedAfter` / `-ModifiedBefore` | Non | Limite à une période de modification |
| `-MinSizeMB` | Non | Uniquement les fichiers d'au moins cette taille |
| `-IncludeHidden` | Non | Parcourt aussi les listes masquées, les catalogues et les bibliothèques système |
| `-Everything` | Non | Ne rien laisser de côté : `-IncludeSubsites -IncludeHidden` et aucune limite |
| `-Permissions` | Non | `Effective` (défaut), `Unique` (héritage rompu uniquement) ou `None` |
| `-ExpandGroups` | Non | Liste aussi les membres des groupes SharePoint ordinaires (les groupes de liens sont toujours développés) |
| `-IncludeLimitedAccess` | Non | Conserve les attributions `Limited Access` dans le rapport |
| `-MaxItems` | Non | S'arrête après ce nombre de correspondances (5000 par défaut, `0` = aucune limite) |
| `-MaxPermissionLookups` | Non | Limite du nombre de résultats dont les autorisations sont résolues (1000 par défaut, `0` = aucune limite) |
| `-PageSize` | Non | Éléments par appel serveur pendant le crawl (500 par défaut) |
| `-GrantSiteAdmin` | Non | Vous rend temporairement administrateur de site collection (SharePoint Administrator requis) |
| `-KeepSiteAdmin` | Non | Conserve ces droits au lieu de les retirer à la fin |
| `-AdminUpn` | Non | UPN à qui accorder l'administration du site (défaut : le compte connecté) |
| `-TenantId` / `-ClientId` / `-AppName` | Non | Paramètres de connexion, comme dans `Restore-RecycleBinItems.ps1` |
| `-OutputPath` | Non | Chemin du rapport CSV (défaut : `C:\Temp\SharePointFind_<timestamp>.csv`) |
| `-Disconnect` | Non | Se déconnecte de PnP à la fin |

**Exemples**

```powershell
# Où se trouve tout ce qui contient « offerte » dans le nom, et qui peut le voir ?
.\Find-SiteContent.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Sales -Name "*offerte*"

# Texte intégral : quels documents mentionnent « salarisschaal », n'importe où dans l'arborescence du site ?
.\Find-SiteContent.ps1 -SiteUrl https://contoso.sharepoint.com/sites/HR `
    -Content "salarisschaal"

# Ne rien laisser de côté : tous les sous-sites, toutes les listes masquées/système, aucune limite
.\Find-SiteContent.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Finance `
    -Name "*veiligheid*" -Everything

# Tout ce qui, dans le site, est partagé différemment du reste
.\Find-SiteContent.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Finance `
    -Permissions Unique -IncludeSubsites

# Gros PDF dans une bibliothèque, avec les groupes derrière les autorisations développés
.\Find-SiteContent.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Finance `
    -ListName "Gedeelde documenten" -Extension pdf -MinSizeMB 10 -ExpandGroups

# Chercher dans un OneDrive sur lequel vous n'avez aucun droit
.\Find-SiteContent.ps1 `
    -SiteUrl https://contoso-my.sharepoint.com/personal/jane_doe_contoso_com `
    -Name "*.xlsx" -GrantSiteAdmin
```

**Remarques**
- Le CSV contient une ligne par résultat *et par principal*, il se filtre et se pivote donc
  bien : triez sur `SharingLink`, `External` ou `UniqueRights` pour aller directement aux lignes
  intéressantes.
- `-Content` ne trouve que ce que l'index de recherche connaît. Les documents tout juste
  téléversés ou récemment modifiés peuvent mettre de quelques minutes à quelques heures à
  apparaître — le mode crawl les voit toujours.
- Un crawl lit chaque élément de chaque bibliothèque. Sur un site comptant des centaines de
  milliers d'éléments, cela prend un moment ; restreignez avec `-ListName` ou utilisez plutôt
  `-Content`.

---

### Search-SharePointContent.ps1

L'équivalent Graph de `Find-SiteContent.ps1` : même question, en app-only, et il atteint chaque
site et chaque OneDrive du tenant sans que vous ayez de droits sur aucun d'entre eux.
Lecture seule.

**Deux moteurs**

| Moteur | Quand | Ce qu'il fait |
|--------|------|--------------|
| Delta (par défaut) | Pas de `-Content` | `/drives/{id}/root/delta` parcourt toute l'arborescence d'une bibliothèque par pages de mille éléments. Le filtrage se fait côté client, donc les caractères génériques `*contient*` fonctionnent |
| Search (`-Content`) | Requête en texte intégral | `/search/query` sur l'index de recherche — trouve du texte *à l'intérieur* des documents. KQL ne gère que les caractères génériques en fin de terme (`veiligheid*`), pas en début. Une recherche app-only doit indiquer une géographie ; le script la lit dans l'emplacement des données du site, ou la trouve par essais, et `-Region` la remplace |

**Autorisations**

`/drives/{id}/items/{id}/permissions`, récupéré par lots de 20 via `/$batch`. Ce seul appel
contient toute l'information par élément :

| Champ | Rapporté comme |
|-------|-------------|
| `roles` | `Read` / `Edit` / `Full Control` |
| `grantedToV2` / `grantedToIdentitiesV2` | L'utilisateur, le groupe Entra, le groupe SharePoint ou l'utilisateur de site — une ligne CSV chacun |
| `link` | `SharingLink` (Anyone / Organization / Specific people, lecture ou modification), `LinkUrl`, `LinkExpires` |
| `inheritedFrom` | Absent → `PermissionSource = Item` (droits uniques). Présent → `Inherited`, avec le dossier d'où cela provient |

Les autorisations « Tout le monde disposant du lien » sont comptées à part dans la synthèse et
affichées en rouge, car elles sont accessibles sans même se connecter. Les invités externes
(`#EXT#`) et « Everyone except external users » ont aussi leurs propres compteurs.

**Ce que Graph ne sait pas faire** — bon à savoir avant de choisir celui-ci :

- **Pas d'autorisations au niveau site ou liste.** Il n'existe pas d'API Graph pour les
  attributions de rôles SharePoint ; `/sites/{id}/permissions` ne renvoie que les autorisations
  d'application (Sites.Selected). Ce script vous indique si un élément hérite et de quel dossier,
  pas ce que le site lui-même accorde à quel groupe. Utilisez `Find-SiteContent.ps1` pour cela.
- **Bibliothèques uniquement.** Les autorisations existent pour les driveItems, pas pour les
  éléments des listes ordinaires.
- Pas de définitions de rôles, et pas de nuance « Limited Access ».

**Connexion — app-only, créée pour vous**

La première exécution sur un tenant met l'application en place :

1. Connexion à Graph en tant que Global Administrator (une fois)
2. Création ou réutilisation d'une application nommée d'après `-AppName`
3. Attribution, avec consentement administrateur, du rôle **d'application** `Sites.Read.All`
   (plus `Group.Read.All` si vous utilisez `-ExpandGroups`)
4. Création d'un certificat auto-signé dans `Cert:\CurrentUser\My` et téléversement de sa clé
   publique — **aucun secret n'est écrit sur le disque**
5. Mise en cache du client ID et de l'empreinte par tenant dans `graph.appid.json` à la racine
   du dépôt (gitignored)

Les exécutions suivantes se connectent en app-only sans aucune invite, ce qui rend ce script
utilisable depuis une tâche planifiée. Utilisez votre propre application avec `-ClientId` plus
`-CertificateThumbprint` ou `-ClientSecret`. Le certificat n'est pas exportable et se trouve dans
le magasin de l'utilisateur qui l'a créé : une tâche planifiée doit donc s'exécuter sous ce même
compte.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-SiteUrl` | * | Une site collection (ou un OneDrive) à parcourir |
| `-AllSites` | * | Parcourt plutôt chaque site du tenant |
| `-SiteFilter` | Non | Avec `-AllSites` : uniquement les sites dont l'URL correspond à ce caractère générique |
| `-MaxSites` | Non | Avec `-AllSites` : s'arrête après ce nombre de sites |
| `-IncludePersonalSites` | Non | Avec `-AllSites` : inclut le OneDrive de chacun |
| `-IncludeSubsites` | Non | Avec `-SiteUrl` : parcourt aussi ses sous-sites |
| `-Content` | Non | Requête texte intégral/KQL — bascule vers l'index de recherche |
| `-Region` | Non | Géographie pour `-Content` (`EUR`, `NAM`, `DEU`, …). La recherche app-only en exige une ; détectée automatiquement si possible |
| `-Name` | Non | Filtre sur le nom du fichier/dossier, caractères génériques autorisés |
| `-Path` | Non | Filtre sur le chemin du dossier, correspondance partielle |
| `-Extension` | Non | Une ou plusieurs extensions, avec ou sans le point |
| `-ItemType` | Non | `All` (défaut), `File` ou `Folder` |
| `-LibraryName` | Non | Uniquement ces bibliothèques de documents, caractères génériques autorisés |
| `-ModifiedBy` | Non | Auteur de la dernière modification — nom d'affichage ou e-mail |
| `-ModifiedAfter` / `-ModifiedBefore` | Non | Limite à une période de modification |
| `-MinSizeMB` | Non | Uniquement les fichiers d'au moins cette taille |
| `-Permissions` | Non | `Effective` (défaut), `Unique` ou `None` |
| `-ExpandGroups` | Non | Liste aussi les membres des groupes Entra (nécessite `Group.Read.All`) |
| `-Everything` | Non | Sous-sites, sites personnels et aucune limite |
| `-MaxItems` | Non | S'arrête après ce nombre de correspondances (5000 par défaut, `0` = aucune limite) |
| `-MaxPermissionLookups` | Non | Limite du nombre de recherches d'autorisations (2000 par défaut, `0` = aucune limite) |
| `-TenantId` | * | Obligatoire avec `-AllSites` ; sinon déduit de `-SiteUrl` |
| `-ClientId` / `-CertificateThumbprint` / `-ClientSecret` / `-AppName` | Non | Utiliser votre propre app registration |
| `-OutputPath` | Non | Chemin du CSV (défaut : `C:\Temp\GraphSharePointFind_<timestamp>.csv`) |
| `-MaxRetries` | Non | Nouvelles tentatives en cas de throttling (429), 5 par défaut |

**Exemples**

```powershell
# Un site et ses sous-sites, tout ce qui contient « veiligheid » dans le nom
.\Search-SharePointContent.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Finance `
    -Name "*veiligheid*" -IncludeSubsites

# Texte intégral sur tout le tenant : quels documents mentionnent « salarisschaal » ?
.\Search-SharePointContent.ps1 -AllSites -TenantId contoso.onmicrosoft.com `
    -Content "salarisschaal"

# Tout ce qui, dans le tenant, porte ses propres autorisations — commencer par 25 sites
.\Search-SharePointContent.ps1 -AllSites -TenantId contoso.onmicrosoft.com `
    -Permissions Unique -MaxSites 25

# Sans surveillance, avec une application que vous avez déjà
.\Search-SharePointContent.ps1 -AllSites -TenantId contoso.onmicrosoft.com `
    -ClientId 0000...-4444 -CertificateThumbprint A1B2C3 `
    -Name "*.pfx" -OutputPath C:\Reports\keys.csv
```

**Remarques**
- Un crawl delta à l'échelle du tenant lit chaque élément de chaque bibliothèque qu'il touche.
  Commencez avec `-MaxSites`, et n'ajoutez `-IncludePersonalSites` que si vous le voulez
  vraiment — cela multiplie le travail par le nombre d'utilisateurs.
- Le throttling (HTTP 429) est géré : les appels individuels et les sous-requêtes des lots sont
  retentés, en respectant `Retry-After`.
- `-Content` interroge tout l'index du tenant et les résultats sont ensuite refiltrés sur les
  sites concernés : une requête avec de très nombreux résultats passe donc un peu de temps sur des
  résultats qu'elle écarte.

**Modules requis**
```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser  # pour l'exécution
Install-Module Microsoft.Graph.Applications  -Scope CurrentUser   # uniquement pour l'enregistrement unique de l'application
```

---

### Restore-RecycleBinItems.ps1

Liste la corbeille de premier et/ou de second niveau, filtre les éléments (nom, chemin, auteur
de la suppression, date) et restaure les correspondances à leur emplacement d'origine. Essai à
blanc par défaut — rien n'est restauré tant que vous ne passez pas `-Apply`.

**Deux modes**

| Mode | Ce qu'il fait |
|------|--------------|
| `-SiteUrl <url>` | Une site collection — site d'équipe, site de communication ou le OneDrive d'un utilisateur (`https://contoso-my.sharepoint.com/personal/jane_doe_contoso_com`) |
| `-AllSites -TenantUrl <url>` | Chaque site SharePoint du tenant. **Les sites personnels OneDrive sont exclus** — restaurez-les un par un avec `-SiteUrl` |

Le balayage de tout le tenant ignore aussi l'hôte My Site, les sites de redirection et les sites
verrouillés en lecture seule ou sans accès (une restauration y échouerait de toute façon). Il
indique combien de sites de chaque type il a ignorés. Restreignez avec
`-SiteFilter "*/sites/Finance*"` et faites un essai avec `-MaxSites 5` avant de le lancer sur tout
le tenant.

Un site en erreur (pas d'accès, throttling, disparu) est signalé et le balayage continue — un
site défaillant n'interrompt pas l'exécution. Les résultats par site figurent dans le tableau de
synthèse et dans le CSV, qui comporte une colonne `Site` dans les deux modes.

C'est à l'échelle du tenant qu'un essai à blanc prouve son utilité : il répond à « quels sites
contiennent des fichiers que ce compte a supprimés mardi » sans rien toucher.

**Vitesse — des lots, pas des threads**

Les restaurations passent par `Restore-PnPRecycleBinItem -IdList`, qui transmet à SharePoint tout
un ensemble d'éléments en un seul appel serveur. 200 fichiers coûtent alors un aller-retour au lieu
de 200, et c'est de là que vient la vitesse — la restauration d'une bibliothèque complète passe de
plusieurs dizaines de minutes à quelques minutes.

Restaurer dans des runspaces parallèles n'est *pas* la voie la plus rapide ici, et le script ne le
fait volontairement pas :

- PnP PowerShell n'est pas thread-safe ; chaque runspace a besoin de son propre import de module
  (quelques secondes) et de sa propre connexion, et les objets de connexion ne franchissent pas
  proprement les frontières entre runspaces
- Tous les éléments atterrissent dans la même site collection, donc les appels parallèles se
  disputent les mêmes listes et déclenchent le throttling SharePoint (HTTP 429) — vous obtenez des
  nouvelles tentatives et des échecs partiels, pas de la vitesse
- Un seul appel par lot fait déjà côté serveur ce que les threads essayaient de paralléliser

Comportement des lots :

- `-BatchSize` (1-200, 200 par défaut) détermine la taille des lots ; `-BatchSize 1` restaure
  strictement élément par élément
- Les dossiers et les fichiers ne partagent jamais un lot, pour que l'ordre « dossiers d'abord »
  soit préservé
- Un lot est tout ou rien et son erreur n'indique pas quel élément l'a fait échouer ; un lot en
  échec est donc automatiquement retenté un élément à la fois — les éléments valides sont quand
  même restaurés et le CSV nomme ceux qui ne l'ont pas été

**Combien de temps cela prend-il ?**

Le script vous le dit, pour que vous sachiez s'il faut attendre ou aller chercher un café :

- La lecture de la corbeille est chronométrée et indiquée par site — sur un site comptant des
  dizaines de milliers d'éléments supprimés, cela seul peut prendre plusieurs minutes (utilisez
  `-RowLimit` pour la plafonner). À l'échelle du tenant, cette lecture représente généralement
  l'essentiel du temps d'exécution, pas la restauration
- Pendant `-Apply`, la barre de progression affiche le temps écoulé et une estimation en direct,
  recalculée à partir du rythme réellement mesuré sur ce tenant plutôt que d'une estimation
  initiale. Avec `-AllSites`, il y a deux barres : les sites, et les lots au sein du site en cours
- La synthèse indique la durée réelle et, à l'échelle du tenant, un tableau par site avec les
  éléments trouvés, correspondants, restaurés et en échec, ainsi que les secondes de
  lecture/restauration de chacun
- Le CSV comporte les colonnes `Site`, `Batch` et `DurationSeconds`, qui permettent de repérer
  les plus lents

**Connexion — l'app registration est créée pour vous**

PnP PowerShell ne fournit plus d'application multi-tenant partagée, une app registration Entra
est donc nécessaire. La première exécution sur un tenant en crée une automatiquement :

1. Connexion à Microsoft Graph en tant que Global Administrator du tenant cible
2. Le script crée (ou réutilise) une application public-client nommée d'après `-AppName`
3. Il attribue, avec consentement administrateur, l'étendue SharePoint déléguée
   `AllSites.FullControl`
4. Le client ID est mis en cache par tenant dans `pnp.appid.json` à la racine du dépôt (gitignored)

Les exécutions suivantes lisent le client ID en cache et passent directement à la connexion
SharePoint interactive — la connexion administrateur à Graph n'a lieu qu'une fois par tenant.
Passez un `-ClientId` existant pour ignorer entièrement la création de l'application.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-SiteUrl` | * | URL de la site collection dans laquelle restaurer (site d'équipe, site de communication ou OneDrive) |
| `-AllSites` | * | Parcourt plutôt chaque site SharePoint du tenant ; OneDrive est exclu |
| `-TenantUrl` | * | URL du tenant pour `-AllSites`, par ex. `https://contoso.sharepoint.com` |
| `-SiteFilter` | Non | Avec `-AllSites` : uniquement les sites dont l'URL correspond à ce caractère générique |
| `-MaxSites` | Non | Avec `-AllSites` : s'arrête après ce nombre de sites |
| `-TenantId` | Non | ID/domaine du tenant pour la connexion unique à Graph (défaut : déduit de `-SiteUrl`) |
| `-ClientId` | Non | Application Entra existante avec laquelle se connecter — ignore l'enregistrement de l'application |
| `-AppName` | Non | Nom d'affichage de l'application à créer/réutiliser (défaut : `M365-Scripts SharePoint Restore`) |
| `-Name` | Non | Filtre sur le nom de l'élément, caractères génériques autorisés (`*.xlsx`, `Budget*`) |
| `-Path` | Non | Filtre sur l'emplacement d'origine, correspondance partielle (`Shared Documents/Finance`) |
| `-DeletedBy` | Non | Filtre sur l'auteur de la suppression — nom d'affichage ou e-mail, caractères génériques autorisés |
| `-DeletedAfter` / `-DeletedBefore` | Non | Limite à une période de suppression |
| `-ItemType` | Non | `All` (défaut), `File`, `Folder` ou `ListItem` |
| `-Stage` | Non | `All` (défaut), `FirstStage` (corbeille de l'utilisateur) ou `SecondStage` (corbeille de l'administrateur de site collection) |
| `-RowLimit` | Non | Limite du nombre d'entrées de corbeille récupérées (à utiliser sur des corbeilles énormes) |
| `-BatchSize` | Non | Éléments par appel serveur, 1-200 (200 par défaut). `1` = strictement un par un |
| `-GrantSiteAdmin` | Non | Vous ajoute temporairement comme administrateur de site collection pour chaque site — nécessaire pour le OneDrive d'un autre utilisateur et en pratique requis pour `-AllSites` |
| `-KeepSiteAdmin` | Non | Conserve ces droits au lieu de les retirer à la fin |
| `-AdminUpn` | Non | UPN à qui accorder l'administration du site (défaut : le compte connecté) |
| `-Apply` | Non | Restaure réellement. Sans cela, le script se contente d'un rapport |
| `-OutputPath` | Non | Chemin du rapport CSV (défaut : `C:\Temp\RecycleBinRestore_<timestamp>.csv`) |
| `-Disconnect` | Non | Se déconnecte de PnP à la fin |

**Exemples**

```powershell
# Essai à blanc — afficher tout le contenu des deux corbeilles d'un site
.\Restore-RecycleBinItems.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Finance

# Restaurer tout ce qu'un utilisateur a supprimé hier soir
.\Restore-RecycleBinItems.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Finance `
    -DeletedBy "jane.doe@contoso.com" -DeletedAfter "2026-08-27 18:00" -Apply

# Restaurer les fichiers Excel d'un dossier de bibliothèque, corbeille de second niveau uniquement
.\Restore-RecycleBinItems.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Finance `
    -Name "*.xlsx" -Path "Shared Documents/Budget" -Stage SecondStage -Apply

# Restaurer le OneDrive d'un utilisateur, en vous accordant temporairement l'administration du site
.\Restore-RecycleBinItems.ps1 `
    -SiteUrl https://contoso-my.sharepoint.com/personal/jane_doe_contoso_com `
    -GrantSiteAdmin -Apply

# Essai à blanc sur tout le tenant : quels sites contiennent des fichiers que ce compte a supprimés aujourd'hui ?
.\Restore-RecycleBinItems.ps1 -AllSites -TenantUrl https://contoso.sharepoint.com `
    -DeletedBy "jane.doe@contoso.com" -DeletedAfter "2026-08-28" -GrantSiteAdmin

# Restauration sur tout le tenant après une suppression en masse — essayer d'abord 5 sites
.\Restore-RecycleBinItems.ps1 -AllSites -TenantUrl https://contoso.sharepoint.com `
    -DeletedBy "jane.doe@contoso.com" -DeletedAfter "2026-08-28" `
    -GrantSiteAdmin -MaxSites 5 -Apply
```

**Remarques**
- Les éléments sont restaurés dossiers d'abord et chemins courts d'abord : un fichier ne peut pas
  être restauré tant que le dossier dans lequel il se trouvait est lui-même encore supprimé.
- La restauration échoue lorsqu'un élément portant le même nom existe déjà à l'emplacement
  d'origine — ces éléments sont signalés dans le CSV avec l'erreur SharePoint, rien d'autre n'est
  interrompu.
- Les éléments de second niveau (site collection) sont restaurés directement à leur emplacement
  d'origine, pas dans la corbeille de premier niveau.
- La rétention par défaut est de 93 jours sur les deux niveaux. Les éléments plus anciens ont
  disparu et aucun script ne peut les récupérer — c'est une limite de la plateforme Microsoft, pas
  du script.
- Nécessite SharePoint Administrator (ou Global Administrator) pour `-GrantSiteAdmin` et pour
  l'enregistrement unique de l'application ; la restauration elle-même ne nécessite qu'un accès au
  site.

**Modules requis**
```powershell
Install-Module PnP.PowerShell -Scope CurrentUser              # PowerShell 7.4+
Install-Module Microsoft.Graph.Applications -Scope CurrentUser # uniquement pour l'enregistrement unique de l'application
```

---

### Trace-SharePointFile.ps1

Répond à « où est passé mon fichier ? » pour OneDrive et SharePoint : renommé, déplacé,
copié, supprimé ou restauré, par qui et quand. Chaque heure est affichée en **heure de
Bruxelles** (heure d'été et d'hiver prises en compte, avec le décalage UTC), et vous pouvez
indiquer une période. Lecture seule : rien ne change dans le tenant.

SharePoint ne garde pas l'ancien nom ni l'ancien emplacement d'un fichier ; le **Unified
Audit Log**, si — c'est donc la source. Le script le lit et reconstitue la trace du fichier :

| Étape | Ce qu'elle fait |
|-------|-----------------|
| Lecture | Chaque renommage, déplacement, copie, suppression, mise à la corbeille, restauration et chargement de fichiers **et de dossiers** dans la période. Lu par jour ; une tranche contenant plus des 50 000 enregistrements qu'une recherche peut renvoyer est scindée jusqu'à ce qu'elle tienne (jusqu'à 15 minutes), et une recherche en échec ou incohérente est relancée |
| Point de départ | Les enregistrements qui citent le fichier via `-Name`, `-Url` ou `-ItemId` |
| Suivi | Chaque enregistrement du même élément (`ListItemUniqueId`, qui survit aux renommages et déplacements) et chaque enregistrement qui part d'un chemin vers lequel le fichier a été renommé ou déplacé. Une chaîne `A → B → C` aboutit à C, même si C ne ressemble en rien au nom recherché. Un déplacement vers un autre site est suivi via le chemin de destination |
| Dossiers | Renommer, déplacer ou supprimer un dossier déplace chaque fichier qu'il contient **sans enregistrement par fichier**. Les enregistrements de dossier sont donc rejoués sur le chemin du fichier à ce moment, pour que « déplacé avec le dossier X » et « supprimé avec le dossier X » apparaissent aussi |

Pour chaque élément vous obtenez la chronologie, le **dernier emplacement connu** et un
statut : `Present`, `In recycle bin` (à restaurer avec
[`Restore-RecycleBinItems.ps1`](#restore-recyclebinitemsps1)), `In second-stage recycle bin`
ou `Permanently deleted`.

**Paramètres**

| Paramètre | Type | Défaut | Description |
|-----------|------|--------|-------------|
| `-Name` | string | — | Nom du fichier tel qu'il a été à un moment donné. Sans extension, toute extension correspond (`Offerte` trouve `Offerte.docx`) ; `*` et `?` sont des caractères génériques |
| `-Url` | string | — | URL complète du fichier telle qu'elle était. `?web=1` est ignoré et un lien de partage `/:w:/r/` est reconverti en chemin. Les liens de partage opaques (`/:w:/s/`, `/:w:/g/`) ne peuvent pas être suivis — ouvrez-les et copiez l'adresse sur laquelle ils aboutissent |
| `-ItemId` | guid | — | Le `ListItemUniqueId`, p. ex. tiré du CSV d'une exécution précédente |
| `-SiteUrl` | string | — | Uniquement les enregistrements de ce site ou de ce OneDrive. Beaucoup plus rapide dans un grand tenant |
| `-StartDate` | date ou string | `-Days` avant `-EndDate` | Heure locale dans `-TimeZone` : `15-09-2026`, `15/09/2026 08:30`, `2026-09-15 08:30` (jour d'abord, à la belge) |
| `-EndDate` | date ou string | maintenant | Même notation. Une date sans heure inclut toute cette journée |
| `-Days` | int | `30` | Durée de la période si `-StartDate` n'est pas indiqué |
| `-TimeZone` | string | `Europe/Brussels` | ID IANA ou Windows ; fonctionne dans Windows PowerShell 5.1 et PowerShell 7 |
| `-FollowCopies` | switch | désactivé | Suivre aussi les copies. Par défaut une copie est signalée mais pas suivie — l'original reste où il était |
| `-IncludeActivity` | switch | désactivé | Aussi les ouvertures, modifications, téléchargements, synchronisations et archivages/extractions : qui y a travaillé en dernier. Beaucoup plus d'enregistrements, donc plus lent |
| `-OutputPath` | string | `C:\Temp\FileTrail_<nom>_<ts>.csv` | Chemin du CSV ; les enregistrements d'audit bruts de la trace vont dans le même nom avec `.json` |
| `-TenantId` | string | — | Domaine du tenant pour `Connect-ExchangeOnline` ; inutile si vous êtes déjà connecté |
| `-PassThru` | switch | désactivé | Renvoie aussi les lignes de la chronologie sous forme d'objets |

**Exemples**

```powershell
# Où est passé « Offerte Janssens.docx » ces 30 derniers jours ?
.\Trace-SharePointFile.ps1 -Name "Offerte Janssens.docx" -TenantId contoso.onmicrosoft.com

# Une période en heure de Bruxelles, un seul OneDrive
.\Trace-SharePointFile.ps1 -Name "Budget*" -StartDate '01-09-2026' -EndDate '15-09-2026' `
    -SiteUrl https://contoso-my.sharepoint.com/personal/jan_contoso_com

# À partir du lien que quelqu'un a envoyé un jour, un après-midi
.\Trace-SharePointFile.ps1 -Url "https://contoso.sharepoint.com/sites/Sales/Shared Documents/2026/Prijslijst.xlsx" `
    -StartDate '2026-09-12 13:00' -EndDate '2026-09-12 18:00'
```

**Remarques**

- Nécessite le rôle **View-Only Audit Logs** ou **Audit Logs** dans Exchange Online, et le module `ExchangeOnlineManagement`. Fonctionne dans Windows PowerShell 5.1 et PowerShell 7
- Le journal d'audit a 30 à 90 minutes (parfois 24 heures) de retard. Audit Standard conserve **180 jours** ; le script avertit si la période commence plus tôt
- Seul ce qui s'est passé **dans** la période peut être suivi. Un dossier renommé avant `-StartDate` est invisible ; si la trace semble commencer en cours de route, élargissez la période
- Les renommages par le client de synchronisation OneDrive (dans l'Explorateur) sont audités comme ceux du navigateur ; la colonne `UserAgent` les distingue
- Le dernier emplacement connu est ce que dit le journal d'audit — le script ne vérifie pas que le fichier s'y trouve encore
- Colonnes du CSV : `Item, Time, TimeUtc, Action, Operation, User, From, To, ViaFolder, ItemId, ClientIP, UserAgent, RecordId`. `ViaFolder` est rempli quand l'étape provient d'une action sur un dossier

---

### Revoke-SharePointUserAccess.ps1

Le pendant de [`Get-SharePointPermissionsReport.ps1`](../Reporting/readme.fr.md#get-sharepointpermissionsreportps1) : ce script-là vous dit qui a accès à quoi, celui-ci retire cet accès. Il recherche chaque endroit où un utilisateur désigné a accès et le supprime :

- **Administrateur de site collection** — en premier, car ce rôle l'emporte sur toutes les attributions de rôles en dessous ; le laisser en place rendrait le reste purement cosmétique
- **Attributions de rôles directes** sur un site, un sous-site, une liste/bibliothèque, un dossier ou un fichier individuel
- **Groupes SharePoint** (Owners, Members, Visitors et groupes personnalisés)
- **Liens de partage** — les groupes `SharingLinks.*` dans lesquels un lien partagé place ses destinataires. C'est ainsi que « tout le monde disposant du lien » et « personnes spécifiques » donnent réellement accès à une personne

Le rapport est le comportement par défaut. Rien ne change sans `-Apply`, et chaque exécution écrit un CSV indiquant exactement ce qui a été trouvé et ce qui en a été fait.

#### Ce qu'il ne fait volontairement *pas*

| | Pourquoi |
|---|---|
| Modifier l'appartenance aux groupes Entra ID | **Sauf si `-RemoveFromEntraGroups` est fourni.** Par défaut, quiconque entre via un groupe de sécurité ou M365 conserve cet accès — le groupe *est* l'attribution — et ces chemins sont signalés explicitement, avec le nom du groupe, pour que vous ne croyiez pas que c'est fermé alors que c'est encore ouvert |
| Supprimer `Everyone` / `Everyone except external users` | Cela retire l'accès à tout le tenant, pas à cette personne. Signalé, pas modifié |
| Nettoyer la propriété et les métadonnées | Un utilisateur révoqué reste l'auteur de ce qu'il a créé |

> **L'offboarding se fait donc en deux étapes.** Exécutez ce script, puis traitez les groupes Entra qui figurent dans le CSV sous `Action = CannotRevoke`. Sans cette seconde étape, l'accès n'a pas disparu.

#### Robustesse

Ce script supprime des autorisations ; ses modes de défaillance diffèrent donc de ceux d'un rapport : toucher silencieusement la mauvaise personne, ou ne pas pouvoir retracer ce que vous avez retiré.

| Situation | Comportement |
|---|---|
| Identifier l'utilisateur | **Comparaisons exactes uniquement.** Sur l'UPN, l'e-mail, le suffixe de la revendication et le nom d'invité décodé (`jan_partner.com#ext#@tenant` devient `jan@partner.com`). Jamais sur une sous-chaîne : `an@contoso.com` est contenu dans `jan@contoso.com`, et c'est exactement ainsi qu'on révoque la mauvaise personne |
| Deux comptes avec la même adresse | Le site n'est **pas** touché ; le script s'arrête avec les deux noms de connexion dans l'erreur. C'est à vous de choisir, pas au script |
| CSV d'audit | Écrit ligne par ligne pendant l'exécution, pas à la fin. Une exécution qui révoque deux cents choses puis plante doit toujours pouvoir dire *ce qui* a disparu |
| Le CSV est ouvert dans Excel | Cinq tentatives avec un délai croissant, puis l'exécution s'arrête — mieux vaut une exécution interrompue que des autorisations supprimées sans trace |
| La suppression renvoie `404` | `AlreadyGone`, pas une erreur. Lors d'une deuxième exécution, c'est le résultat normal ; compté comme erreur, une exécution propre paraîtrait défaillante |
| `-WhatIf` | Même branche qu'un essai à blanc, donc `WouldRevoke` dans le CSV — pas `Skipped`, qui laisserait croire que quelqu'un a refusé une invite |
| Impossible de retirer l'administrateur de site collection | **Signalé haut et fort, et l'exécution le compte comme un échec.** Ce rôle atteint chaque étendue du site, donc toutes les autres suppressions y sont alors cosmétiques |
| Throttling (`429`/`503`) | Nouvelle tentative avec `Retry-After` ; un jeton refusé est renouvelé une fois avant que l'exécution ne s'arrête |
| Erreur inattendue | Un `trap` nettoie l'application Full Control temporaire avant l'arrêt du script |

#### Paramètres

| Paramètre | Type | Défaut | Description |
|---|---|---|---|
| `-UserPrincipalName` | string | — | **Obligatoire.** L'utilisateur, par ex. `jan@contoso.com`. Pour un invité, l'adresse réelle (`jan@partner.com`) fonctionne aussi — le script trouve lui-même la variante `#ext#` |
| `-TenantUrl` | string | — | Racine du tenant. Obligatoire pour une exécution à l'échelle du tenant |
| `-SiteUrl` | string | — | Une site collection au lieu de tout le tenant |
| `-Apply` | switch | désactivé | Révoque réellement. Sans cela, rapport uniquement |
| `-Scope` | `Site`/`List`/`Item` | `Item` | Profondeur de recherche des attributions directes |
| `-IncludeGroupAccess` | switch | désactivé | Signale aussi les sites que l'utilisateur atteint via des groupes Entra, *y compris* là où il n'a rien d'autre. Rapport uniquement |
| `-KeepSharingLinks` | switch | désactivé | Laisse les liens de partage intacts ; tous les autres chemins sont bien révoqués |
| `-FromReport` | string | — | Prend les sites à visiter dans une exécution du rapport de permissions au lieu de reparcourir le locataire. Le CSV de détail, un fichier frère de la même exécution, ou le dossier. Sans `-TenantUrl`, le locataire vient aussi du rapport. Voir ci-dessous |
| `-RemoveFromEntraGroups` | switch | désactivé | **Retire aussi l'utilisateur des groupes Entra ID vus en train d'accorder l'accès** — ceux-là uniquement, jamais tous les groupes dont il fait partie. Nécessite Graph `GroupMember.ReadWrite.All`, que l'application temporaire ne demande qu'avec ce commutateur. Voir ci-dessous |
| `-RemoveFromSite` | switch | désactivé | Retire ensuite aussi l'utilisateur de la liste des utilisateurs de chaque site collection. Rattrape ce que le passage étendue par étendue n'a pas vu, mais le nom s'affiche ensuite comme compte supprimé dans les métadonnées plus anciennes |
| `-IncludeOneDriveSites` | switch | désactivé | Parcourt aussi les sites OneDrive personnels |
| `-IncludeHiddenLists` | switch | désactivé | Inclut aussi les listes masquées et système |
| `-TenantId` / `-ClientId` / `-CertificateThumbprint` | string | — | Votre propre app registration au lieu de l'application temporaire. Elle a besoin de SharePoint `Sites.FullControl.All` ainsi que de Graph `Sites.Read.All`, `User.Read.All` et `GroupMember.Read.All`. Sans `User.Read.All`, la recherche de l'utilisateur renvoie `403` et l'exécution s'arrête, au lieu de prendre cela pour un compte inexistant |
| `-ClientSecret` | string | — | Fonctionne pour Graph mais **pas** pour SharePoint (voir l'authentification dans le rapport) |
| `-OutputPath` | string | `C:\Temp` | Dossier de sortie |
| `-GraphTimeoutSec` / `-MaxGraphRetry` | int | `120` / `6` | Timeout et nouvelles tentatives |

L'authentification est identique à celle du rapport : une app registration de courte durée, basée sur un certificat, avec SharePoint `Sites.FullControl.All`, supprimée à la fin.


#### Retirer aussi les groupes Entra ID

`-RemoveFromEntraGroups` termine la seconde moitié d'un départ au lieu de se contenter de la signaler. **Seuls les groupes que cette exécution a réellement vus détenir une attribution de rôle sur une portée dans le périmètre sont touchés** — jamais tous les groupes auxquels l'utilisateur appartient. Un partant peut être dans cinquante groupes ; ceux qui accordent l'accès SharePoint sont ceux concernés ici.

> **Cela dépasse SharePoint.** Un groupe Entra n'est pas un objet SharePoint. La même appartenance porte souvent une équipe Teams, une boîte aux lettres, des licences et des attributions d'applications, qu'aucun de ces rapports ne voit. Lisez d'abord le rapport d'une exécution sans `-Apply`, puis relancez avec le commutateur.

> **La liste nest complète que dans la mesure du scan.** Une exécution sur un seul site, un `-Scope` restreint, OneDrive ou les listes masquées exclues et les portées illisibles la réduisent — un groupe accordant laccès à un endroit jamais fouillé ny figure pas. Lexécution nomme chaque limite avant tout retrait et de nouveau dans le résumé. Le script de révocation effectue son propre scan : le rapport de permissions nest pas un prérequis.

Quatre cas sont signalés plutôt que forcés, car les forcer échouerait ou ferait la mauvaise chose :

| Cas | Pourquoi il est laissé tel quel |
|---|---|
| Groupe dynamique | L'appartenance suit une règle et n'est pas stockée : il n'y a rien à retirer. Modifiez la règle, ou les attributs de l'utilisateur qu'elle cible |
| Synchronisé depuis AD on-premises | En lecture seule dans le cloud. L'appartenance doit être retirée dans Active Directory |
| Membre via un groupe imbriqué | L'utilisateur n'est pas membre direct : le retrait échouerait ici. L'accès doit être coupé au groupe qui le contient réellement, que le rapport nomme |
| Utilisateur introuvable dans Entra | Rien dont le retirer ; le volet SharePoint s'exécute quand même |

Le retrait passe par le même entonnoir que toute autre modification : `-Apply`, `-WhatIf`, la confirmation et le CSV d'audit se comportent à l'identique. Nécessite Graph `GroupMember.ReadWrite.All`, que l'application temporaire ne demande **que** si le commutateur est fourni — une exécution en lecture seule ne détient aucune permission capable de modifier une appartenance.


#### Travailler à partir du rapport de permissions

`-FromReport` prend les sites à visiter dans une exécution de [`Get-SharePointPermissionsReport.ps1`](../Reporting/readme.fr.md#get-sharepointpermissionsreportps1) au lieu de reparcourir le locataire. C'est ainsi que les deux scripts s'articulent : le rapport dit qui peut atteindre quoi, vous le lisez et décidez, et la révocation agit exactement sur ce que vous regardiez.

```powershell
# 1. Tout lister, sur tout le locataire, en un seul classeur
.\..\Reporting\Get-SharePointPermissionsReport.ps1 -TenantUrl "https://contoso.sharepoint.com" -IncludeEffectiveAccess -Excel

# 2. Lire, décider, puis révoquer une personne à partir des mêmes données
.\Revoke-SharePointUserAccess.ps1 -UserPrincipalName jan@contoso.com -FromReport C:\Temp
```

Indiquez le CSV de détail, un autre fichier de la même exécution, ou simplement le dossier. Sans `-TenantUrl`, le locataire est également tiré du rapport.

Il privilégie le fichier d'accès par site (`..._SiteAccess_...csv`), où chaque groupe est déjà résolu en personnes : un site n'est alors visité que si cet utilisateur fait réellement partie du groupe qui accorde l'accès. Le repli sur la liste brute des attributions oblige à faire confiance à un groupe SharePoint et à visiter tous les sites qui en comportent un ; l'exécution indique la source utilisée.

> **Le rapport décide où chercher, jamais quoi supprimer.** Chaque site nommé est tout de même lu en direct : une attribution disparue entre-temps est signalée `AlreadyGone` au lieu d'échouer, et une suppression manuelle n'est pas annulée. L'inverse n'est pas vrai : tout ce qui a été accordé *après* l'écriture du rapport est invisible ici, de même que ce que le rapport lui-même n'a pas pu lire. Les deux figurent dans le résumé, et un rapport vieux de plus d'un jour le signale.
#### Sortie

`SharePoint_Revoke_<user>_<ts>.csv`, une ligne par attribution trouvée, avec une colonne `Action` :

| Action | Signification |
|---|---|
| `WouldRevoke` | Trouvée, et serait supprimée — c'est ce que vous obtenez sans `-Apply` |
| `Revoked` | Supprimée |
| `Failed` | Tentative échouée ; la raison figure dans `Detail` |
| `AlreadyGone` | Plus rien à supprimer — le résultat normal lors d'une deuxième exécution |
| `CannotRevoke` | Via un groupe Entra ou `Everyone` — doit être résolu ailleurs |
| `Kept` | Volontairement conservée par `-KeepSharingLinks` |
| `Skipped` | Refusée à l'invite de confirmation |

#### Exemples

```powershell
# Que peut atteindre Jan ? Ne modifie rien
.\Revoke-SharePointUserAccess.ps1 -UserPrincipalName jan@contoso.com -TenantUrl "https://contoso.sharepoint.com"

# La même chose, et cette fois révoquer réellement
.\Revoke-SharePointUserAccess.ps1 -UserPrincipalName jan@contoso.com -TenantUrl "https://contoso.sharepoint.com" -Apply

# Retirer un invité d'une site collection, liens de partage compris
.\Revoke-SharePointUserAccess.ps1 -UserPrincipalName gast@partner.com -SiteUrl "https://contoso.sharepoint.com/sites/Finance" -Apply

# Checklist d'offboarding : aussi les groupes Entra qui donnent accès
.\Revoke-SharePointUserAccess.ps1 -UserPrincipalName jan@contoso.com -TenantUrl "https://contoso.sharepoint.com" -IncludeGroupAccess
```

> Exécution sans surveillance ? Passez `-Confirm:$false`, sinon le script demande une confirmation pour chaque suppression (`ConfirmImpact = 'High'`).

---

### Test-SharePointAccessScripts.ps1

Vérifie `Revoke-SharePointUserAccess.ps1` et `Get-SharePointPermissionsReport.ps1` sans toucher à un tenant. Exécutez-le après chaque modification de l'un ou l'autre ; un code de sortie 0 signifie que les deux sont en ordre.

Deux points sont vérifiés, tous deux des erreurs qui passent inaperçues en production :

1. **Le bloc d'authentification partagé est identique à l'octet près.** Les deux scripts contiennent la même couche d'authentification app-only et SharePoint REST, délimitée par `SHARED BLOCK START/END`. Cette couche a nécessité quatre exécutions réelles sur un tenant pour être correcte — certificat au lieu de secret, jetons qui doivent prouver leurs rôles d'application avant d'être mis en cache, 401 fatal plutôt que par site, pagination qui ne peut pas rester bloquée. Une seconde copie qui dérive silencieusement est un risque d'exactitude précisément dans le script qui supprime des autorisations. En cas de différence, la première ligne divergente est affichée.
2. **L'entonnoir de révocation se comporte correctement.** Un essai à blanc doit consigner son intention et ne rien exécuter, `-Apply` doit exécuter *et* consigner, un échec doit aboutir dans la piste d'audit au lieu de disparaître, et les attributions que le script doit refuser de supprimer (via un groupe Entra, ou accordées à tout le monde) doivent rester refusées.

```powershell
.\Test-SharePointAccessScripts.ps1
```
