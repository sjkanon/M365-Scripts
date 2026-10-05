[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [SharePoint](../readme.fr.md) › **Provisioning**

# Provisionnement de la structure SharePoint

Provisionnez et maintenez une structure SharePoint — modèle de métadonnées, bibliothèques,
types de contenu et autorisations de groupe — à partir d'un seul fichier de configuration,
avec PnP PowerShell et Microsoft Graph.

Rien dans les scripts n'est propre à un client : le modèle réside dans le JSON, donc un
deuxième client correspond à un deuxième fichier de configuration, pas à un deuxième fork.
Les exemples ci-dessous utilisent une société fictive, Contoso NV, avec les marques
Northwind et Fabrikam.

---

## Sommaire

- [Le principe](#le-principe)
- [Scripts](#scripts)
- [Avant de commencer](#avant-de-commencer)
- [Ordre des opérations](#ordre-des-opérations)
- [Le fichier de configuration](#le-fichier-de-configuration)
- [Canaux standard et autorisations — à lire](#canaux-standard-et-autorisations--à-lire)
- [Planifier l'audit](#planifier-laudit)
- [Ce que les scripts ne feront jamais](#ce-que-les-scripts-ne-feront-jamais)

---

## Le principe

Une seule équipe Microsoft 365, un canal par pilier, et la différence entre les piliers
portée par les **métadonnées et les autorisations de groupe** plutôt que par une
prolifération de sites.

| | |
|---|---|
| **Piliers** | MGMT (canal privé), Leveranciers, Verkopers, Klanten, Marketing, TD |
| **Bibliothèque supplémentaire** | Beeldmateriaal voor klanten — en lecture seule pour les clients externes |
| **Marque** | Northwind / Fabrikam / Beide — une étiquette sur chaque fichier, jamais un site ou un groupe distinct |
| **Groupes** | un groupe de sécurité Entra ID par pilier et par niveau d'accès (`SG-CONTOSO-<Pijler>-RW` / `-RO`), plus `SG-CONTOSO-Klanten-Extern` |

Le modèle de métadonnées, réutilisable dans toutes les bibliothèques :

| Colonne | Nom interne | Type | Valeurs |
|---|---|---|---|
| Merk | `PsMerk` | Choix | Northwind / Fabrikam / Beide |
| Pijler | `PsPijler` | Choix | MGMT / Leveranciers / Verkopers / Klanten / Marketing / TD |
| Regio | `PsRegio` | Choix | Benelux / Duitsland / Frankrijk / Export — obligatoire uniquement sur Verkoopdocument |
| Leverancier | `PsLeverancier` | Métadonnées gérées | ensemble de termes, extensible depuis le magasin de termes |
| Taal | `PsTaal` | Choix multiple | NL / FR / DE / EN / Geen taal |
| Contenttype | `PsContenttype` | Choix | Catalogus / Prijslijst / Schrijfrichtlijn / Afbeelding+certificaat / Marketingslag |
| Vertrouwelijkheid | `PsVertrouwelijkheid` | Choix | Intern / Deelbaar met klant / Vertrouwelijk |
| Deelstatus | `PsDeelstatus` | Choix | tenue à jour par le script d'audit — jamais remplie à la main |
| Status | `PsStatus` | Choix | Actief / Te archiveren / Verouderd |

> **Pourquoi le préfixe `Ps`.** « Contenttype » et « Status » sont des noms d'affichage que
> SharePoint utilise déjà pour autre chose. Préfixer les noms *internes* garde les colonnes
> sans ambiguïté dans le CAML, dans les vues et dans le contrôle de dérive, tandis que les
> utilisateurs voient toujours de simples libellés en néerlandais.

Un type de contenu par pilier décide lesquelles sont obligatoires — un Marketingdocument
ne peut pas être enregistré sans Taal et Contenttype, un Verkoopdocument ne peut pas être
enregistré sans Regio.

---

## Scripts

| Script | Ce qu'il fait | Écrit ? |
|---|---|---|
| [`Install-SharePointStructure.ps1`](Install-SharePointStructure.ps1) ([docs](#install-sharepointstructureps1)) | **Commencez ici.** Demande comment tout doit s'appeler, puis construit l'ensemble : inscription d'application, équipe, canaux, métadonnées, bibliothèques, autorisations, vérification | oui |
| [`New-StructureConfig.ps1`](New-StructureConfig.ps1) ([docs](#des-questions-pas-un-fichier-json)) | Les questions. S'exécute tout seul depuis l'installateur ; lancez-le seul pour préparer une configuration à l'avance | écrit la configuration |
| [`New-SharePointTeam.ps1`](New-SharePointTeam.ps1) ([docs](#install-sharepointstructureps1)) | L'équipe Microsoft 365 et ses canaux, canaux privés compris, et les URL de site réécrites dans la configuration | oui |
| [`New-SharePointMetadata.ps1`](New-SharePointMetadata.ps1) ([docs](#new-sharepointmetadataps1)) | Ensemble de termes, colonnes de site, types de contenu — sur chaque site de la configuration | oui |
| [`Set-SharePointLibraries.ps1`](Set-SharePointLibraries.ps1) ([docs](#set-sharepointlibrariesps1)) | Bibliothèques, dossiers de canal, liaison des types de contenu, métadonnées par défaut, vues, autorisations de groupe | oui |
| [`Update-SharePointShareStatus.ps1`](Update-SharePointShareStatus.ps1) ([docs](#update-sharepointsharestatusps1)) | Déduit Deelstatus des autorisations réelles, signale les fichiers partagés plus largement que leur étiquette ne le permet | une colonne |
| [`Test-SharePointStructure.ps1`](Test-SharePointStructure.ps1) ([docs](#test-sharepointstructureps1)) | Compare le tenant avec la configuration et signale chaque différence | jamais |
| [`Sync-SharePointChannelMember.ps1`](Sync-SharePointChannelMember.ps1) ([docs](#canaux-privés-et-groupes)) | Fait d'un groupe de sécurité la source de vérité pour les membres d'un canal privé | liste des membres du canal |
| [`Add-SharePointHelpPage.ps1`](Add-SharePointHelpPage.ps1) ([docs](#remise-au-client)) | Écrit l'explication destinée aux utilisateurs finaux sur le site d'équipe, générée à partir de la configuration | oui |
| [`Remove-SharePointStructure.ps1`](Remove-SharePointStructure.ps1) ([docs](#revenir-en-arrière)) | Supprime ce qui a été construit — se contente d'un rapport tant que vous ne passez pas `-Apply` | oui, volontairement |
| [`SharePointStructure.Common.ps1`](SharePointStructure.Common.ps1) | Fonctions d'aide partagées — chargées par dot-sourcing, pas exécutées seules | — |
| [`SharePoint-Handleiding.md`](SharePoint-Handleiding.md) | **Guide utilisateur, en néerlandais** — à remettre au client : téléverser, étiqueter, retrouver les documents | — |
| [`example.config.json`](example.config.json) | Le modèle, comme exemple à copier — encore sur `CHANGEME`. Les configurations client (`<client>.config.json`) se trouvent à côté et sont ignorées par git | — |

Chaque script qui écrit prend en charge `-WhatIf` et est idempotent : une deuxième
exécution affiche `[ OK ]` partout et ne modifie rien.

---

## Avant de commencer

### 1. Modules

```powershell
Install-Module PnP.PowerShell         -Scope CurrentUser
Install-Module Microsoft.Graph.Groups -Scope CurrentUser   # uniquement pour -EnsureGroups / -IncludeGroups
```

### 2. Une inscription d'application

PnP.PowerShell ne fournit plus d'application multitenant partagée ; il vous en faut donc
une à vous. Réutilisez l'application que [`Find-SiteContent.ps1`](../Find-SiteContent.ps1)
crée et met en cache dans `pnp.appid.json`, ou inscrivez-en une :

| Connexion | Nécessite | À utiliser pour |
|---|---|---|
| **Interactive** (`-Interactive -ClientId <app-id>`) | `AllSites.FullControl` déléguée, connecté en tant qu'administrateur autorisé à modifier le site | le provisionnement, les exécutions ponctuelles |
| **App-only** (`-ClientId <app-id> -Thumbprint <thumb>`) | `Sites.FullControl.All` d'application, certificat téléversé sur l'application | l'audit planifié du statut de partage |

Deux exigences supplémentaires faciles à manquer :

- **Magasin de termes.** Créer le groupe de termes et l'ensemble de termes nécessite un
  administrateur du magasin de termes. En app-only, ce n'est possible que si le principal
  de service de l'application a été ajouté comme tel dans le centre d'administration
  SharePoint. Si c'est un obstacle, lancez une fois `New-SharePointMetadata.ps1 -Only TermSet`
  en interactif et laissez le reste à l'app-only.
- **Groupes.** `-EnsureGroups` nécessite `Group.ReadWrite.All` ; le `-IncludeGroups` du
  contrôle de dérive se contente de `Group.Read.All`.

### 3. Compléter la configuration

`example.config.json` est livré avec `CHANGEME` dans les URL du tenant et des sites.
Copiez-le en `<client>.config.json` et complétez-le, ou laissez `New-StructureConfig.ps1`
en écrire un. Sans `-ConfigPath`, chaque script prend l'unique `*.config.json` d'ici qui
ne contient plus `CHANGEME`, et refuse de s'exécuter sur l'exemple lui-même.
Un fichier encore sur `CHANGEME` est toujours refusé — mieux vaut une
erreur claire qu'une connexion qui échoue cinq minutes après le début d'une exécution.

```jsonc
"tenant": "contoso.onmicrosoft.com",
"sites": {
  "team": "https://contoso.sharepoint.com/sites/Contoso",
  "mgmt": "https://contoso.sharepoint.com/sites/Contoso-MGMT"
}
```

> L'URL **mgmt** est la collection de sites *propre* au canal privé, pas un dossier du site
> d'équipe. Vous la trouverez dans le centre d'administration SharePoint, ou ouvrez l'onglet
> Fichiers du canal MGMT et cliquez sur *Ouvrir dans SharePoint*. Un canal privé possède son
> propre site, c'est pourquoi les colonnes et les types de contenu doivent y être
> provisionnés séparément — une colonne de site ne traverse pas les collections de sites.

---

## Ordre des opérations

### La version courte

```powershell
.\Install-SharePointStructure.ps1
```

C'est tout. Sans configuration pour ce tenant, il demande comment tout doit s'appeler,
écrit lui-même la configuration, puis crée l'équipe, les canaux (y compris le canal
privé), le modèle de métadonnées, les bibliothèques, les groupes, les vues et les
autorisations — et vérifie le résultat.

Ensuite : placez les personnes dans les groupes de sécurité. L'explication qui leur est
destinée se trouve déjà sur le site, dans la navigation de gauche — la construction l'y a
placée.

### Des questions, pas un fichier JSON

`New-StructureConfig.ps1` s'exécute tout seul la première fois. Entrée accepte la
suggestion entre crochets, donc une construction standard se résume surtout à des Entrée
et à deux vraies réponses :

| Question | Suggestion |
|---|---|
| Client, tenant | — |
| Nom d'équipe, alias (détermine l'URL du site), propriétaire | dérivé du nom du client |
| Marques, et comment s'appelle « appartient à toutes » | Northwind, Fabrikam, Beide |
| Piliers, et lesquels sont un canal privé | MGMT, Leveranciers, Verkopers, Klanten, Marketing, TD — MGMT privé |
| Quel pilier gère les fournisseurs / les ventes | détermine où Leverancier et Regio deviennent obligatoires |
| Bibliothèque client | Beeldmateriaal voor klanten |
| Préfixe de groupe et suffixes modification/lecture | `SG-<CLIENT>` · RW · RO |
| Langues, régions, types de documents, niveaux de confidentialité, statuts, fournisseurs de départ | les valeurs par défaut néerlandaises |
| **Tenir à jour la colonne de statut de partage ?** | non — c'est la seule réponse qui vous coûte un script nocturne |
| **Appliquer des droits par pilier sur les dossiers de canal ?** | non — voir la remarque sur les canaux standard ci-dessous |

Tout le reste est dérivé : par pilier un canal, un type de contenu, deux groupes de
sécurité et une vue groupée ; par marque une vue couvrant tous les piliers.

Les questions facultatives affichent `(of "geen")` dans l'indication — tapez-le pour
refuser la suggestion, car Entrée signifie « la prendre ».

#### `-All` — choisir vous-même chaque nom

`New-StructureConfig.ps1 -All` demande aussi les noms qui sont sinon dérivés, chacun
toujours avec la dérivation comme suggestion :

| Question avec `-All` | Suggestion |
|---|---|
| URL du site d'équipe | `https://<tenant>.sharepoint.com/sites/<alias>` |
| La bibliothèque derrière les canaux | `Documents` — à demander sur un tenant non anglophone |
| Groupe de colonnes et groupe de types de contenu | le nom de l'équipe |
| Nom de l'ensemble de termes | `Leveranciers` |
| Le libellé de chaque colonne, tel que les utilisateurs le voient | Merk, Pijler, Regio, Leverancier, Taal, Contenttype, … |
| Par pilier : nom du canal, dossier, type de contenu, les deux noms de groupe, titre de la vue | dérivé du nom du pilier |
| Bibliothèque client : nom du groupe et du type de contenu | `<prefix>-Klanten-Extern`, `Klantmedia` |

Les noms *internes* des colonnes (`PsMerk`, `PsTaal`, …) restent fixes dans tous les cas.
Ils ne sont montrés à personne, et en modifier un alors que des documents le portent fait
perdre les métadonnées de ces documents.

**Deux éléments sont générés une fois puis figés**, car SharePoint y rattache des données :
les noms internes des colonnes (`PsMerk`, `PsTaal`, …) et les ID des types de contenu. Les
noms d'affichage, les noms de canaux et les noms de groupes peuvent tous être modifiés
ensuite ; ces deux-là non, sans perdre les métadonnées des documents qui les portent déjà.
C'est pourquoi l'assistant refuse d'écraser une configuration existante sans `-Force`.

### Install-SharePointStructure.ps1

Une exécution, cinq étapes, avec arrêt à la première erreur plutôt que de construire sur
une base cassée :

| Étape | Quoi |
|---|---|
| 0 | Inscription d'application — créée avec consentement administrateur, ou réutilisée depuis `pnp.appid.json` |
| 1 | `New-SharePointTeam.ps1` — l'équipe Microsoft 365, les canaux y compris le canal privé, et les URL de site réécrites dans la configuration |
| 2 | `New-SharePointMetadata.ps1` — ensemble de termes, colonnes, types de contenu, sur chaque site |
| 3 | `Set-SharePointLibraries.ps1 -EnsureGroups` — groupes, bibliothèques, dossiers, types de contenu, valeurs par défaut, vues, autorisations |
| 4 | `Add-SharePointHelpPage.ps1` — l'explication, sur le site, pour les personnes qui vont l'utiliser |
| 5 | `Test-SharePointStructure.ps1` — vérification en lecture seule de ce qui vient d'être mis en place |
| 6 | `Update-SharePointShareStatus.ps1` avec `-RunAudit` — le premier passage du statut de partage |

**L'étape 4 fait partie de la construction, ce n'est pas une corvée pour plus tard.** Une
structure dont personne n'a entendu parler est une structure que personne n'utilise, et la
page est générée à partir de la même configuration, elle décrit donc ce que l'exécution
vient de créer. `-SkipHelpPage` l'omet ; `-HelpContact` indique à qui les utilisateurs
doivent s'adresser.

**L'étape 1 est la raison pour laquelle cela fonctionne à partir d'un tenant vide.** La
collection de sites d'un canal privé est provisionnée de façon asynchrone et son URL ne
peut pas être connue à l'avance — SharePoint la construit à partir du nom de l'équipe et du
canal. Le script l'attend en l'interrogeant régulièrement (quelques minutes, c'est normal)
et l'écrit dans la configuration, afin que les étapes suivantes aient un endroit auquel se
connecter. Passez `-SkipTeam` lorsque l'équipe existe déjà.

```powershell
# Construction ponctuelle sur un tenant que vous ne gérez pas au quotidien : ne rien laisser derrière soi
.\Install-SharePointStructure.ps1 -TemporaryApp -RunAudit

# Prudent : tout sauf les autorisations non prises en charge sur les dossiers de canal
.\Install-SharePointStructure.ps1 -SkipChannelFolderPermissions

# Utiliser une inscription d'application que vous avez déjà
.\Install-SharePointStructure.ps1 -ClientId <app-id>
```

| Code de sortie | Signification |
|---|---|
| 0 | construit et vérifié |
| 1 | une étape a échoué |
| 2 | construit, mais la vérification a trouvé des différences |

**À propos de `-TemporaryApp`.** Il supprime l'inscription d'application à la fin — mais
uniquement une application *créée par cette exécution*. Une application déjà en cache est
antérieure à l'exécution et c'est à quelqu'un d'autre de la supprimer ; le script le
signale donc au lieu de la supprimer discrètement. Sans ce commutateur, l'application reste
en place et l'ID client est mis en cache, ce dont l'audit planifié et les contrôles de
dérive ultérieurs ont besoin.

**Une exécution `-WhatIf` a besoin d'une application pour se connecter.** Sans application
en cache pour le tenant, il n'y a rien avec quoi se connecter ; l'essai à blanc valide donc
la configuration et s'arrête là. Lancez-le une fois pour de vrai, ou passez le `-ClientId`
d'une application existante, pour faire un essai à blanc étape par étape.

### Ou étape par étape

```powershell
# 1. Regarder avant de sauter - aucune de ces deux commandes ne modifie quoi que ce soit
.\Test-SharePointStructure.ps1 -Interactive -ClientId <app-id>
.\New-SharePointMetadata.ps1   -Interactive -ClientId <app-id> -WhatIf

# 2. Le modèle de métadonnées d'abord : les bibliothèques ne peuvent pas lier ce qui n'existe pas
.\New-SharePointMetadata.ps1 -Interactive -ClientId <app-id>

# 3. Bibliothèques, liaison des types de contenu et autorisations (crée les groupes au passage)
.\Set-SharePointLibraries.ps1 -Interactive -ClientId <app-id> -EnsureGroups -WhatIf
.\Set-SharePointLibraries.ps1 -Interactive -ClientId <app-id> -EnsureGroups

# 4. Remplir les groupes avec des personnes (portail Entra ID, ou votre propre script d'intégration)

# 5. Premier audit, puis chaque nuit
.\Update-SharePointShareStatus.ps1 -Interactive -ClientId <app-id> -ReportOnly

# 6. Confirmer le résultat
.\Test-SharePointStructure.ps1 -Interactive -ClientId <app-id> -IncludeGroups
```

### New-SharePointMetadata.ps1

Groupe de termes, ensemble de termes et termes ; les neuf colonnes de site ; les sept types
de contenu avec les bonnes colonnes marquées obligatoires. S'exécute sur **chaque** site de
la configuration, de sorte que le site du canal privé reçoit sa propre copie.

```powershell
# Uniquement le site du canal privé, après le passage de MGMT dans son propre canal
.\New-SharePointMetadata.ps1 -Interactive -ClientId <app-id> -Site mgmt

# Uniquement l'ensemble de termes, par un administrateur du magasin de termes
.\New-SharePointMetadata.ps1 -Interactive -ClientId <app-id> -Only TermSet
```

Rendre une colonne obligatoire après coup fonctionne : l'indicateur `Required` d'un lien de
champ existant est mis à jour sur place et propagé aux listes qui utilisent déjà le type de
contenu.

### Set-SharePointLibraries.ps1

Par conteneur : la bibliothèque ou le dossier de canal, les types de contenu liés à la
bibliothèque, l'ordre propre des types de contenu du dossier (pour que le menu *Nouveau*
du canal Leveranciers propose Leveranciersdocument et non les cinq types appartenant aux
autres piliers), les valeurs de colonne par défaut, une vue groupée et les attributions de
rôles.

```powershell
# La bibliothèque externe, en retirant tout ce que la configuration ne mentionne pas
.\Set-SharePointLibraries.ps1 -Interactive -ClientId <app-id> `
    -Container KlantBibliotheek -RemoveOtherPermissions

# Tout sauf les autorisations non prises en charge sur les dossiers des canaux standard
.\Set-SharePointLibraries.ps1 -Interactive -ClientId <app-id> `
    -SkipChannelFolderPermissions
```

| Commutateur | Effet |
|---|---|
| `-EnsureGroups` | créer les groupes de sécurité Entra ID qui n'existent pas encore |
| `-RemoveOtherPermissions` | supprimer les attributions de rôles que la configuration ne mentionne pas (les propriétaires, les liens de partage et la revendication everyone ne sont jamais supprimés) |
| `-RemoveStockContentType` | retirer le type intégré *Document* pour que personne ne puisse classer sans métadonnées |
| `-SkipChannelFolderPermissions` | laisser les dossiers de canal hériter — voir ci-dessous |

Aucune vue n'est jamais définie par défaut. La vue par défaut d'une bibliothèque Teams est
ce que chaque membre du canal voit dès qu'il ouvre Fichiers.

#### Vues transversales — ce qui rend réelle « la marque comme étiquette »

Les vues par pilier n'affichent jamais qu'un seul dossier de canal. La section
`libraryViews` ajoute des vues sur la bibliothèque partagée elle-même avec
`Scope = RecursiveAll`, afin qu'elles couvrent **tous** les dossiers de pilier dans une
seule liste à plat :

| Vue | Affiche |
|---|---|
| `Alles - Northwind` | chaque fichier étiqueté Northwind **ou Beide**, sur tous les piliers, groupé par pilier |
| `Alles - Fabrikam` | la même chose pour Fabrikam |
| `Nog te taggen` | les fichiers sans Merk — ce que laissent le glisser-déposer et la synchronisation OneDrive |
| `Extern gedeeld` | tout ce que l'audit a trouvé exposé hors de l'organisation |
| `Te archiveren` | Status vaut Te archiveren ou Verouderd |

C'est la réponse à « un fichier, deux marques » : un fichier étiqueté `Beide` est stocké
une seule fois et apparaît dans les deux vues de marque. Pas de copies qui divergent.

Le filtre est du CAML brut dans la configuration plutôt qu'un mini-langage de requête
inventé par ce script :

```jsonc
{
  "title": "Alles - Northwind",
  "recursive": true,
  "groupBy": "PsPijler",
  "where": "<Or><Eq><FieldRef Name='PsMerk' /><Value Type='Text'>Northwind</Value></Eq><Eq><FieldRef Name='PsMerk' /><Value Type='Text'>Beide</Value></Eq></Or>",
  "fields": [ "DocIcon", "LinkFilename", "PsPijler", "PsContenttype", "PsTaal", "..." ]
}
```

Ne groupez jamais une vue sur `PsTaal` — SharePoint refuse de grouper sur une colonne à
valeurs multiples. Filtrer dessus fonctionne très bien.

### Update-SharePointShareStatus.ps1

Détermine comment chaque document est réellement partagé et l'écrit dans Deelstatus :

| Verdict | Quand |
|---|---|
| `Extern - bewerken` | un lien Tout le monde en modification, un invité avec Collaboration ou plus, ou un lien de modification pour des personnes spécifiques incluant un invité |
| `Extern - alleen bekijken` | la même chose, en lecture seule |
| `Intern gedeeld` | un lien ou une attribution directe qui reste à l'intérieur du tenant |
| `Niet gedeeld` | le fichier hérite simplement de son dossier ou de sa bibliothèque |

Il répond aussi à la question pour laquelle le modèle existe vraiment : **y a-t-il quelque
chose étiqueté Intern ou Vertrouwelijk derrière un lien externe ?** Ces éléments figurent
dans le rapport et l'exécution se termine avec le code 2.

La colonne est écrite avec `SystemUpdate`, donc Modifié et Modifié par restent inchangés et
aucune nouvelle version n'est créée — une exécution nocturne ne fait pas remonter toute la
bibliothèque en tête des *modifications récentes*.

| Code de sortie | Signification |
|---|---|
| 0 | terminé, rien n'est partagé plus largement que son étiquette ne le permet |
| 1 | échec |
| 2 | au moins une violation de Vertrouwelijkheid |

### Test-SharePointStructure.ps1

Lecture seule. Une ligne par différence, en trois variantes :

| Type | Signification | Corrigé par |
|---|---|---|
| `Missing` | dans la configuration, pas sur le tenant | les scripts de provisionnement |
| `Different` | présent, mais pas comme configuré | les scripts de provisionnement |
| `Extra` | sur le tenant, pas dans la configuration | personne — c'est à vous de décider |

Le code de sortie 2 signifie une dérive, il s'intègre donc directement dans une
supervision. Le contrôle qui se rentabilise le plus souvent est l'**indicateur obligatoire
sur un champ de type de contenu** — quelqu'un le décoche dans le navigateur et rien ne
semble anormal jusqu'à ce que la moitié d'une bibliothèque n'ait plus de Taal.

---

## Le fichier de configuration

| Section | Contient |
|---|---|
| `tenant`, `sites` | où tout se trouve ; chaque conteneur fait référence aux clés de `sites` |
| `termStore` | groupe de termes, ensemble de termes et termes de départ pour Leverancier |
| `columns` | les colonnes de site : nom interne, nom d'affichage, type, choix, valeur par défaut |
| `contentTypes` | un par pilier, `fields[].required` décidant de ce qui est obligatoire |
| `groups` | les groupes de sécurité Entra ID, par nom d'affichage |
| `containers` | les bibliothèques et dossiers de canal, et qui y obtient quel rôle |

Un conteneur :

```jsonc
{
  "key": "Leveranciers",
  "kind": "ChannelFolder",          // ChannelFolder = un canal Teams ; Library = sa propre bibliothèque
  "site": "team",                   // clé de la section sites
  "list": "Documents",              // la bibliothèque du canal
  "folder": "Leveranciers",
  "contentTypes": [ "Leveranciersdocument" ],
  "defaultContentType": "Leveranciersdocument",
  "defaultColumnValues": { "PsPijler": "Leveranciers", "PsStatus": "Actief" },
  "uniquePermissions": true,
  "keepExistingPermissions": true,  // copier les droits hérités lors de la rupture de l'héritage
  "view": { "title": "Op leverancier", "fields": [ ... ], "groupBy": "PsLeverancier" },
  "permissions": [
    { "group": "SG-CONTOSO-Leveranciers-RW", "role": "Contribute" },
    { "group": "SG-CONTOSO-Leveranciers-RO", "role": "Read" }
  ]
}
```

La configuration est contrôlée avant toute connexion : un type de contenu qui fait
référence à une colonne non définie, ou un conteneur qui accorde des droits à un groupe
absent du modèle, échoue au chargement plutôt qu'au milieu du provisionnement.

**Ajouter un fournisseur (leverancier)** ne nécessite aucune modification de la
configuration — ajoutez le terme dans le magasin de termes et il apparaît dans la colonne.
La liste `terms` n'est que l'ensemble de départ ; les termes ajoutés en dehors de la
configuration sont signalés par le contrôle de dérive mais jamais supprimés.

---

## Canaux standard et autorisations — à lire

La configuration fournie donne des autorisations uniques à chaque dossier de canal
standard. **Microsoft ne prend pas en charge cette combinaison.**

Un canal standard est, par conception, visible par tous les membres de l'équipe.
Restreindre les autorisations SharePoint sur le dossier du canal masque bien les fichiers,
mais Teams continue d'afficher le canal : un membre qui a perdu l'accès obtient une erreur
dans l'onglet Fichiers au lieu d'une porte fermée. Ça fonctionne, ce n'est pas joli, et ce
n'est pas approuvé.

Les manières prises en charge de fermer un pilier :

| Option | Compromis |
|---|---|
| **Canal privé** | collection de sites propre, appartenance propre — ce que MGMT utilise déjà. Le plus propre, mais le canal n'apparaît pas du tout pour les non-membres |
| **Canal partagé** | collection de sites propre, appartenance propre, peut inclure des personnes extérieures à l'équipe |
| **Bibliothèque propre** (`kind: Library`) | en dehors de la structure des canaux, les autorisations uniques sont entièrement prises en charge — ce qu'utilise la bibliothèque client |

Si vous déplacez un pilier vers un canal privé ou partagé, mettez `uniquePermissions` à
`false` pour son conteneur et ajoutez son nouveau site à la section `sites`.

`Set-SharePointLibraries.ps1` avertit pour chaque dossier de canal standard dont il rompt
l'héritage, et `-SkipChannelFolderPermissions` les laisse hériter tout en traitant les
types de contenu, les valeurs par défaut et les vues.

---

## Planifier l'audit

En app-only, sur un serveur ou un RMM, chaque nuit :

```powershell
pwsh -NoProfile -File .\Update-SharePointShareStatus.ps1 `
    -ClientId <app-id> -Thumbprint <thumbprint> -Quiet
```

Avec `-Quiet`, l'exécution n'affiche que le résumé, de sorte qu'elle apparaît dans le flux
d'activité avec quelque chose d'utile à dire. Le code de sortie 2 signifie qu'un fichier
étiqueté Intern ou Vertrouwelijk se trouve derrière un lien externe — cela mérite une
alerte.

Un contrôle de dérive hebdomadaire en complément :

```powershell
pwsh -NoProfile -File .\Test-SharePointStructure.ps1 `
    -ClientId <app-id> -Thumbprint <thumbprint> -Quiet -IncludeGroups
```

---

## Un canal réservé à un seul groupe

C'est la forme à choisir lorsqu'un pilier doit être fermé **et** conserver un vrai rôle en
lecture seule. L'assistant la propose sous le nom `bibliotheek` :

```
Welke pijlers moeten afgeschermd worden [MGMT]:
In welke vorm [bibliotheek]:
```

Ce qu'il construit :

| Élément | Où |
|---|---|
| Un canal normal dans l'équipe | tout le monde le voit dans Teams, comme il se doit pour un canal |
| Sa propre bibliothèque de documents | sur le site d'équipe, pas un dossier dans la bibliothèque partagée |
| **Autorisations uniques, héritage rompu sans copie** | pour que les membres de l'équipe n'y arrivent **pas** en tant qu'éditeurs |
| `-RW` → Contribute, `-RO` → **Read** | un vrai rôle en lecture seule, ce qu'un canal privé ne peut pas offrir |
| Un onglet dans le canal pointant vers la bibliothèque | l'onglet Fichiers propre au canal ne peut pas être redirigé, la bibliothèque se place donc à côté |

Les autorisations qui subsistent sont celles des propriétaires du site plus ces deux
groupes. C'est ce que signifie `keepExistingPermissions: false` sur le conteneur, et c'est
la différence entre « fermé » et « fermé en théorie ».

> **L'onglet Fichiers intégré du canal pointe toujours vers la bibliothèque de l'équipe.**
> Il contiendra un dossier que personne n'utilise. Dites aux utilisateurs d'utiliser
> l'onglet nommé, ou retirez une fois à la main l'onglet Fichiers du canal.

Comparé aux alternatives :

| Forme | Rôle en lecture seule | Canal visible par les non-membres | Pris en charge |
|---|---|---|---|
| **Bibliothèque propre + canal** | oui | oui (les fichiers non) | oui |
| Canal privé | non — les membres peuvent modifier, point | non | oui |
| Dossier de canal standard avec droits uniques | oui | oui, mais l'onglet Fichiers affiche des erreurs | non |

---

## Canaux privés et groupes

**On ne peut pas accorder de droits sur un canal privé par l'intermédiaire d'un groupe.**
Teams suit l'appartenance personne par personne, et Graph n'y accepte que des utilisateurs
individuels. Il n'y a pas de contournement, et le contournement évident est un piège :

| Approche | Verdict |
|---|---|
| Ajouter le groupe comme membre du canal | Impossible — Graph n'accepte que des utilisateurs |
| Ajouter le groupe aux autorisations SharePoint du site du canal | Fonctionne environ un jour. Teams resynchronise la liste des membres du canal par-dessus, et entre-temps ces personnes accèdent aux fichiers alors que le canal leur reste invisible dans Teams. Non pris en charge |
| **Laisser le groupe alimenter la liste des membres** | Ce que fait `Sync-SharePointChannelMember.ps1` |

```powershell
.\Sync-SharePointChannelMember.ps1 -WhatIf     # qui serait ajouté
.\Sync-SharePointChannelMember.ps1             # les ajouter
.\Sync-SharePointChannelMember.ps1 -Prune      # et retirer ceux que les groupes ne listent plus
```

Vous gérez le groupe ; le script place ses membres dans le canal. Les groupes imbriqués
sont suivis, les objets qui ne sont pas des utilisateurs sont écartés, et chacun est
d'abord ajouté comme membre de l'équipe parente — Teams refuse un membre de canal privé
qui ne fait pas partie de l'équipe, et l'erreur renvoyée ne le dit pas.

Les groupes qui alimentent chaque canal proviennent de `channelMembers` du conteneur. Une
configuration écrite avant l'existence de cette clé se rabat sur les groupes configurés
portant le nom du conteneur, et signale qu'elle l'a fait.

> ### Il n'existe pas de rôle en lecture seule dans un canal privé
>
> Un canal privé a des propriétaires et des membres, et les membres peuvent publier,
> modifier et supprimer des fichiers. Un groupe nommé `-RO` ne peut donc pas y signifier
> « peut consulter » — toute personne ajoutée par ce script peut écrire. L'exécution
> indique par groupe combien de personnes elle a ajoutées, pour que ce soit visible plutôt
> que supposé.
>
> Si la lecture seule compte vraiment pour un pilier, un canal privé n'est pas la bonne
> forme. Utilisez une bibliothèque de documents avec ses propres autorisations, où Read est
> un vrai rôle — la bibliothèque client fonctionne déjà ainsi.

**L'alternative à connaître :** un canal *partagé* prend en charge l'appartenance par
groupe. Y déplacer un pilier est la manière prise en charge de laisser des groupes décider
de l'accès à un canal. Le comportement dépend des paramètres de partage externe et de B2B
direct connect du tenant ; testez donc un canal avant de déplacer quoi que ce soit
d'important.

---

## Remise au client

L'explication a sa place sur le site, pas dans ce dépôt. `Add-SharePointHelpPage.ps1`
l'y écrit sous forme de page SharePoint, générée à partir de la même configuration que
celle qui a servi à construire la structure :

```powershell
.\Add-SharePointHelpPage.ps1 -WhatIf                      # ce qu'elle contiendrait
.\Add-SharePointHelpPage.ps1 -Interactive -ClientId <app-id>
.\Add-SharePointHelpPage.ps1 -Interactive -ClientId <app-id> -Force   # après une modification
```

Comme elle est générée, elle ne peut pas diverger : les canaux qu'elle liste sont ceux qui
existent, les étiquettes qu'elle explique portent le même texte d'aide que celui que les
utilisateurs voient sous chaque champ, et les champs obligatoires par type de document sont
lus dans les types de contenu. Renommez un canal, relancez-la, et la page dit la nouvelle
chose.

Elle est écrite pour la personne qui téléverse un catalogue. Pas de noms de groupes, pas de
noms internes de colonnes, pas de types de contenu ni de colonnes de site. Deux éléments de
la configuration en sont volontairement exclus : la `note` d'un conteneur, qui nomme des
groupes de sécurité, et la `description` d'une vue, qui parle de piliers et de dossiers
synchronisés — leurs titres sont suffisamment clairs par eux-mêmes. Les autorisations n'y
figurent pas du tout : savoir qui peut voir quoi n'est pas quelque chose sur lequel un
utilisateur peut agir, et l'expliquer ne fait qu'inviter la question de savoir pourquoi il
ne le peut pas.


[`SharePoint-Handleiding.md`](SharePoint-Handleiding.md) est
écrit pour les personnes qui vont réellement téléverser des fichiers — en néerlandais, sans
jargon, cinq minutes de lecture. Il couvre les trois façons d'ajouter un fichier et pourquoi
elles se comportent différemment, ce que signifie chaque étiquette, et ce qui se passe dès
que vous étiquetez quelque chose.

Deux points qu'il contient méritent d'être connus de l'administrateur, car ce sont les
questions qui reviennent :

- **Le glisser-déposer et la synchronisation OneDrive ne demandent rien.** Les colonnes
  obligatoires sont imposées par le formulaire de téléversement, pas par la bibliothèque.
  Les fichiers déposés en masse arrivent avec des étiquettes vides et une invite
  « Informations requises » — ils ne sont pas bloqués. La vue `Nog te taggen` est la liste
  de nettoyage, et le guide indique aux utilisateurs de la traiter avec une sélection
  multiple et le volet de détails.
- **Une étiquette n'est pas un verrou.** Mettre Vertrouwelijkheid sur Vertrouwelijk
  n'exclut personne ; c'est un accord, plus le signal que l'audit nocturne utilise pour
  signaler le surpartage. L'accès vient des groupes de sécurité. Le guide le dit dans un
  encadré, car sinon les utilisateurs supposeront le contraire.

---

## Revenir en arrière

[`Remove-SharePointStructure.ps1`](Remove-SharePointStructure.ps1) démonte la même
configuration, en commençant par le plus profond. **Il a le comportement par défaut inverse
de tout le reste ici : sans `-Apply`, il ne modifie rien.** Oublier `-WhatIf` sur un script
destructif est la direction dangereuse, donc l'état sûr est celui que vous obtenez
d'office.

```powershell
.\Remove-SharePointStructure.ps1                                   # ce qui disparaîtrait
.\Remove-SharePointStructure.ps1 -Scope Channels,Groups -Apply     # une partie
.\Remove-SharePointStructure.ps1 -Scope All,Team -IncludeContent -Apply   # tout recommencer
```

| `-Scope` | Supprime |
|---|---|
| `Tabs` | les onglets de bibliothèque ajoutés aux canaux |
| `Channels` | les canaux configurés, et les fichiers de leurs dossiers |
| `Libraries` | les bibliothèques qu'un conteneur possède (`kind: Library`) |
| `ContentTypes` | d'abord dissociés des listes, puis supprimés |
| `Columns` | les colonnes de site, sur chaque site de la configuration |
| `TermSet` | l'ensemble de termes, son groupe et ses termes |
| `Groups` | les groupes de sécurité Entra ID |
| `Team` | le groupe Microsoft 365 — le site, chaque bibliothèque, chaque fichier, chaque conversation |
| `All` | tout ce qui précède **sauf** `Team` |

`All` n'inclut jamais l'équipe. Supprimer toute l'équipe d'un client n'est pas quelque
chose que l'on devrait obtenir en demandant « tout » — vous devez la nommer, puis taper le
nom de l'équipe pour confirmer.

**Ce qu'il refuse de faire :**

- Une bibliothèque ou un dossier de canal qui contient encore des fichiers est ignoré sauf
  avec `-IncludeContent`. Le nombre d'éléments est signalé dans tous les cas.
- Le canal Général et la bibliothèque Documents propre à l'équipe ne sont jamais supprimés.
- Un type de contenu encore utilisé est signalé, pas forcé.

**Ce qu'aucune corbeille ne ramène :** supprimer l'ensemble de termes rend orpheline la
valeur Leverancier de chaque document qui en portait une — le champ conserve un GUID qui ne
pointe vers rien. Supprimer une colonne emporte ses données. Les deux sont signalés avec ce
coût avant de s'exécuter. Un groupe ou une équipe supprimé(e) est conservé(e) en
suppression réversible pendant 30 jours ; un canal supprimé a sa propre corbeille de 30
jours ; les fichiers d'une bibliothèque supprimée vont dans la corbeille du site.

---

## Ce que les scripts ne feront jamais

Des omissions délibérées, chacune pour une raison :

| Jamais | Pourquoi |
|---|---|
| Supprimer une colonne de site, un type de contenu ou un lien de champ | cela emporterait les métadonnées des documents existants |
| Retirer un terme du magasin de termes | les documents étiquetés référencent des GUID de termes ; l'ensemble de termes est destiné à être étendu depuis l'interface |
| Supprimer une attribution de rôle que la configuration ne mentionne pas | sauf si vous passez `-RemoveOtherPermissions` — une exception créée volontairement par quelqu'un n'est pas une dérive |
| Retirer les propriétaires du site ou les groupes de liens de partage | `-RemoveOtherPermissions` les ignore ; les retirer empêcherait le client d'accéder à sa propre bibliothèque |
| Révoquer un lien de partage | l'audit signale le surpartage. Révoquer est une décision, et une décision revient à une personne |
| Définir une vue par défaut | chaque membre du canal le remarquerait immédiatement |

---

## Remarques

- Auteur : Sjoerd Kanon
- `SharePointStructure.Common.ps1` est chargé par dot-sourcing par les quatre scripts.
  C'est une exception délibérée à la règle « chaque script est autonome » appliquée
  ailleurs dans ce dépôt : ces quatre scripts partagent un même schéma de configuration, et
  trois copies du code d'autorisations divergeraient en moins d'un mois.
- Testez d'abord sur un tenant hors production. Les scripts de provisionnement modifient
  les autorisations d'une équipe en production.
