[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../readme.fr.md) › [scripts](../readme.fr.md) › **Startup**

# Startup

Scripts de démarrage et bibliothèque de fonctions M365 centrale.

---

## Fichiers

| Fichier | Description |
|------|-------------|
| [`functies.ps1`](functies.ps1) ([docs](#functiesps1)) | Bibliothèque de fonctions M365 — chargée par dot-sourcing par `menu.ps1` à la première utilisation |
| [`RequiredModules.psd1`](RequiredModules.psd1) ([docs](#requiredmodulespsd1)) | La liste unique des modules dont ce dépôt a besoin — lue par `load.ps1`, `Install-Modules.ps1` et `Update-Modules.ps1` |
| [`Install-Modules.ps1`](Install-Modules.ps1) ([docs](#install-modulesps1)) | Script d'amorçage — installe et importe tous les modules PowerShell nécessaires |
| [`Update-Modules.ps1`](Update-Modules.ps1) ([docs](#update-modulesps1)) | Vérifie les modules requis (manquant, trop ancien, mise à jour disponible) et les installe/met à jour ; met aussi à jour, au choix, tous les autres modules installés |
| [`Test-PowerShellSyntax.ps1`](Test-PowerShellSyntax.ps1) ([docs](#test-powershellsyntaxps1)) | Vérifie par analyse syntaxique les fichiers `.ps1` du dépôt, sans les exécuter |
| [`Update-ScriptIndex.ps1`](Update-ScriptIndex.ps1) ([docs](#update-scriptindexps1)) | Régénère [`scripts/INDEX.md`](../INDEX.md) — la liste A–Z consultable de tous les scripts |
| [`Test-MarkdownLinks.ps1`](Test-MarkdownLinks.ps1) ([docs](#test-markdownlinksps1)) | Vérifie chaque lien de chaque readme — fichiers qui doivent exister, ancres qui doivent correspondre à un titre |
| [`Convert-MarkdownToHtml.ps1`](Convert-MarkdownToHtml.ps1) ([docs](#convert-markdowntohtmlps1)) | Génère une page HTML autonome et mise en forme à partir d'un document markdown — à coller dans IT Glue ou à imprimer |
| [`Update-ReadmeHeader.ps1`](Update-ReadmeHeader.ps1) ([docs](#update-readmeheaderps1)) | Écrit le sélecteur de langue et le fil d'Ariane en tête de chaque readme, en anglais, néerlandais et français |

---

## Démarrage GDAP délégué

`load.ps1` permet désormais d'enregistrer des valeurs par défaut déléguées dans `load.config.ps1` :

- `authMode` (`GDAP` ou `Direct`)
- `defaultCustomerDomain` (facultatif)
- `useDeviceCodeAuth` (`$true` / `$false`)

Lorsque `authMode` vaut `GDAP` et qu'un `defaultCustomerDomain` est configuré, le menu exécute automatiquement `Connect-Tenant` après le chargement de `functies.ps1`.

Pour enregistrer le lanceur à l'ouverture de session Windows :

```powershell
.\load.ps1 -SetupStartup
```

Pour le retirer :

```powershell
.\load.ps1 -RemoveStartup
```

À chaque démarrage, avant l'ouverture du menu, `load.ps1` vérifie les modules de [`RequiredModules.psd1`](#requiredmodulespsd1) avec [`Update-Modules.ps1`](#update-modulesps1) : il liste ce qui manque, ce qui est plus ancien que le minimum ou en retard sur la PowerShell Gallery, et demande `Deze n module(s) nu installeren/updaten? [J/n]`. La galerie est interrogée au plus une fois toutes les 24 heures, un démarrage normal coûte donc moins d'une seconde. Pour sauter la vérification une fois :

```powershell
.\load.ps1 -SkipModuleCheck
```

Vous pouvez aussi activer ou désactiver le démarrage automatique depuis le menu du lanceur :

- `F` = Enable-LauncherStartup
- `G` = Disable-LauncherStartup

---

## functies.ps1

Bibliothèque de fonctions centrale pour la gestion M365 multi-tenant via Microsoft Graph et Exchange Online. Chargée automatiquement par le menu à la première utilisation d'une option B–E.

### Configuration

Adaptez le bloc `#region Configuration` en haut du fichier à votre organisation :

```powershell
$script:MspAdminAlias       = 'msp-admin'
$script:MspAdminDisplayName = 'MSP - Admin Account'
```

### Sélectionner un tenant client

```powershell
Connect-Tenant -Domain "customer.com"
# Définit $global:cid et $global:connectmsoldomain
# Toutes les fonctions suivantes ciblent automatiquement le tenant sélectionné
```

### Fonctions

**Connexion**

| Fonction | Description |
|----------|-------------|
| `Connect-Tenant` | Sélectionne un client CSP par domaine, renseigne `$cid` et `$connectmsoldomain` |
| `Test-GdapConnection` | Valide le contrat GDAP/CSP délégué + tente une connexion Exchange déléguée |
| `Test-ExoConnection` | Vérifie / rétablit la connexion Exchange Online |

**Exchange Online**

| Fonction | Description |
|----------|-------------|
| `Enable-CopyOfSentItems` | Active la copie des éléments envoyés pour toutes les boîtes aux lettres |
| `Add-SharedMailboxAccess` | Accorde FullAccess + SendAs sur une boîte aux lettres partagée |
| `Set-MailboxLocale` | Définit la langue et le fuseau horaire de toutes les boîtes aux lettres (par défaut : NL / W. Europe) |
| `Add-MailboxAlias` | Ajoute un alias à une boîte aux lettres |
| `Get-MailboxAliases` | Liste tous les alias SMTP par boîte aux lettres |
| `Export-DistributionGroups` | Exporte toutes les listes de distribution en CSV (`C:\Temp\`) |
| `Set-AutoReply` | Configure une réponse d'absence du bureau |

**Entra ID / Graph**

| Fonction | Description |
|----------|-------------|
| `Get-TenantAdmins` | Liste tous les Global Administrators |
| `Add-TenantDomain` | Ajoute un domaine et guide la vérification |
| `Get-TenantLicenses` | Affiche une vue d'ensemble des licences avec leur utilisation et leur disponibilité |
| `Get-TenantUsers` | Liste tous les utilisateurs avec UPN, nom d'affichage et licences |
| `Add-TenantAdmin` | Accorde les droits Global Administrator à un utilisateur |
| `Get-EntraApplication` | Recherche une Enterprise App par nom |
| `Reset-UserPassword` | Réinitialise le mot de passe d'un utilisateur |
| `Export-SignInLogs` | Exporte les journaux de connexion en CSV dans `C:\Temp\` (par défaut : 30 derniers jours) |

**Compte administrateur MSP**

| Fonction | Description |
|----------|-------------|
| `New-MspAdmin` | Crée le compte administrateur MSP en tant que Global Admin dans le tenant client |
| `Set-MspAdminAsGroupOwner` | Définit le compte administrateur MSP comme propriétaire de groupe |
| `Reset-MspAdminPassword` | Réinitialise le mot de passe du compte administrateur MSP |

---

## RequiredModules.psd1

Les modules dont dépend ce dépôt, dans un seul fichier de données PowerShell. `load.ps1`,
`Install-Modules.ps1` et `Update-Modules.ps1` le lisent tous, ils ne peuvent donc plus se
contredire — auparavant chacun avait sa propre liste, et `load.ps1` n'en vérifiait que sept.

**Pour ajouter un module :** ajoutez une ligne ici. Au prochain démarrage de `load.ps1`,
chaque machine le voit comme manquant et propose de l'installer. Relever une
`MinimumVersion` fonctionne de la même façon.

| Clé | Signification |
|-----|---------------|
| `Name` | Nom du module sur la PowerShell Gallery |
| `MinimumVersion` | Plus ancien que cela compte comme *trop ancien* (pas seulement *mise à jour disponible*) |
| `WindowsOnly` | Ignoré sous macOS et Linux |
| `MinimumPSVersion` | Ignoré sur un PowerShell plus ancien — `PnP.PowerShell` 3 exige 7.4 |
| `ImportAtStartup` | Importé par `load.ps1` avant l'ouverture du menu |

Liste actuelle : `ExchangeOnlineManagement`, les sous-modules Graph `Authentication`, `Sites`,
`Identity.DirectoryManagement`, `Identity.SignIns`, `Identity.Governance`, `Applications`,
`Calendar`, `Groups`, `Users`, plus `PnP.PowerShell`, `MicrosoftTeams`, `ImportExcel`, et
sous Windows `WindowsAutopilotIntune` et `IntuneWin32App`.

---

## Install-Modules.ps1

Installe et importe chaque module de [`RequiredModules.psd1`](#requiredmodulespsd1). À exécuter une fois sur une nouvelle machine ou après une installation propre de PowerShell. Les modules déjà installés ne sont pas touchés — les mettre à jour est le rôle d'[`Update-Modules.ps1`](#update-modulesps1).

```powershell
.\scripts\Startup\Install-Modules.ps1
```

**Paramètres**

| Paramètre | Description |
|-----------|-------------|
| `-Force` | Réinstaller les modules même s'ils sont déjà présents |
| `-Scope` | `CurrentUser` (par défaut) ou `AllUsers` (en administrateur) |

---

## Update-Modules.ps1

Vérifie chaque module de [`RequiredModules.psd1`](#requiredmodulespsd1) et lui attribue un statut :

| Statut | Signification | Sans `-CheckOnly` |
|--------|---------------|-------------------|
| `Missing` | Non installé | Installé |
| `BelowMinimum` | Plus ancien que sa `MinimumVersion` | Mis à jour |
| `UpdateAvailable` | Une version plus récente est sur la PowerShell Gallery | Mis à jour |
| `OK` | À jour | — |
| `Unknown` | Galerie injoignable, ou le module n'y figure plus | — |
| `Skipped` | Pas pour cette plateforme ou cette version de PowerShell | — |

Ensuite, sauf avec `-RequiredOnly`, il met à jour comme avant tous les autres modules installés via PowerShellGet. `load.ps1` l'exécute au démarrage sous la forme `-RequiredOnly -Prompt -MaxAgeHours 24`.

**Paramètres**

| Paramètre | Description |
|-----------|-------------|
| `-CheckOnly` | Rapport uniquement ; n'installe ni ne met à jour rien |
| `-RequiredOnly` | Seulement les modules de `RequiredModules.psd1`, pas tout le reste de ce qui est installé |
| `-MaxAgeHours` | Réutiliser les versions de la galerie mises en cache depuis moins de ce nombre d'heures (par défaut `0` = toujours interroger la galerie) |
| `-Scope` | Portée des modules nouvellement installés : `CurrentUser` (par défaut) ou `AllUsers` |
| `-Quiet` | Pas de ligne par module, seulement les erreurs |
| `-PassThru` | Renvoie un objet de statut par module requis (`Name`, `Installed`, `Minimum`, `Latest`, `Status`, `Reason`) |
| `-Prompt` | Pour le démarrage : n'affiche que ce qui manque ou est obsolète, puis demande avant d'installer/mettre à jour. Si tout est en ordre, une seule ligne, `Modules OK` |

**Exemples**

```powershell
# Qu'est-ce qui manque ou est obsolète ? Ne modifie rien
.\scripts\Startup\Update-Modules.ps1 -RequiredOnly -CheckOnly

# Installer ce qui manque et mettre à jour ce qui est en retard, seulement les modules du dépôt
.\scripts\Startup\Update-Modules.ps1 -RequiredOnly

# La même chose, puis mettre aussi à jour tous les autres modules installés
.\scripts\Startup\Update-Modules.ps1
```

**Dans votre profil PowerShell.** Si vous démarrez PowerShell avec votre propre profil (`$PROFILE`) plutôt qu'avec `load.ps1`, ajoutez-y la ligne qu'utilise `load.ps1`, pour que la vérification s'exécute à chaque démarrage de PowerShell :

```powershell
& "C:\chemin\vers\M365-Scripts\scripts\Startup\Update-Modules.ps1" -RequiredOnly -Prompt -MaxAgeHours 24
```

Si tout est à jour, cela affiche une seule ligne et prend bien moins d'une seconde ; la galerie est interrogée au plus une fois par jour.

**Remarques**

- La version installée est lue avec `Get-Module -ListAvailable`, donc un module qui n'a pas été installé via PowerShellGet (copie manuelle, MSI) compte aussi comme installé. Un tel module reçoit la nouvelle version à côté via `Install-Module`, car `Update-Module` refuse les modules qu'il n'a pas installés lui-même.
- Les versions de la galerie viennent de `Find-PSResource` si PSResourceGet est présent (environ 3 s pour toute la liste), sinon de `Find-Module` (environ 9 s). Elles sont mises en cache dans `%LOCALAPPDATA%\M365-Scripts\module-gallery-cache.json` ; `-MaxAgeHours` détermine l'âge maximal de ce cache.
- Hors ligne, l'interrogation de la galerie échoue en silence : les modules manquants et trop anciens sont toujours signalés, *mise à jour disponible* ne l'est pas.
- PowerShell 7 et Windows PowerShell 5.1 ont des dossiers de modules distincts, ils peuvent donc donner un résultat différent sur la même machine. C'est correct, pas un bogue.
- Les anciennes versions ne sont pas supprimées. Exécutez en administrateur pour mettre à jour les modules installés pour tous les utilisateurs.

---

## Test-PowerShellSyntax.ps1

Vérifie par analyse syntaxique les fichiers `.ps1` (et, en option, `.psm1`) sans les exécuter — utilise `[System.Management.Automation.Language.Parser]::ParseFile()`.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-Path` | Non | Fichier ou dossier à vérifier (par défaut : racine du dépôt) |
| `-Recurse` | Non | Parcourt les sous-dossiers lorsque `-Path` est un dossier |
| `-IncludePsm1` | Non | Vérifie aussi les fichiers de module `.psm1` |

**Exemples**

```powershell
# Vérifier tout le dépôt
.\Test-PowerShellSyntax.ps1 -Recurse

# Vérifier un seul fichier
.\Test-PowerShellSyntax.ps1 -Path .\scripts\Entra\New-M365User.ps1
```

Codes de sortie : `0` = aucune erreur, `1` = erreurs de syntaxe trouvées, `2` = erreur de chemin/d'argument.

---

## Update-ScriptIndex.ps1

Génère [`scripts/INDEX.md`](../INDEX.md) : une page qui liste tous les scripts du dépôt
de A à Z, avec un lien vers le fichier, un lien vers le readme de son dossier et une description d'une ligne.

Il existe parce que, sinon, trouver un script sur GitHub revient à deviner dans quel dossier
de charge de travail il se trouve et à ouvrir des readmes jusqu'à ce qu'il apparaisse. Une page générée
se cherche avec Ctrl-F et se clique, et — puisqu'elle est générée — ne peut pas diverger des fichiers comme
le fait un tableau tenu à la main.

**D'où vient la description**

| Ordre | Source |
|-------|--------|
| 1 | Le bloc `.SYNOPSIS` du script, réuni sur toutes les lignes qu'il occupe |
| 2 | À défaut, la première vraie ligne d'un bloc de commentaires `#` en tête |
| 3 | À défaut, rien — et le script est listé sous *Scripts without a description*, pour que le manque soit visible au lieu d'être silencieusement vide |

Un commentaire `#` isolé placé directement au-dessus du code n'est volontairement **pas** utilisé : une ligne
comme `# URL van de theme` au-dessus d'une affectation `$ThemeUrl` décrit cette variable, pas le
script, et la lire comme description mettrait dans le tableau quelque chose de pire que rien.
Un bloc d'en-tête s'étend sur plusieurs lignes, ou est séparé du code par une ligne vide.

**Paramètres**

| Paramètre | Description |
|-----------|-------------|
| `-Root` | Racine du dépôt (par défaut : deux niveaux au-dessus de ce script) |
| `-Check` | N'écrit rien ; sortie `1` lorsque la page commitée ne correspond plus aux scripts sur le disque |
| `-WhatIf` | Indique ce qui changerait sans rien écrire |

**Exemples**

```powershell
# Reconstruire l'index après l'ajout, le renommage ou la suppression d'un script
pwsh -File scripts/Startup/Update-ScriptIndex.ps1

# Échouer lorsque l'index est périmé — pour un hook ou un pipeline
pwsh -File scripts/Startup/Update-ScriptIndex.ps1 -Check
```

> Relancez-le chaque fois qu'un script est ajouté, renommé, déplacé ou supprimé — au moment même où
> les [règles de travail](../../.claude/CLAUDE.md) vous demandent déjà de mettre à jour les readmes et
> `menu.ps1`. Il ne réécrit rien si la page est déjà à jour, vous pouvez donc l'exécuter sans risque
> à chaque commit.

Codes de sortie : `0` = écrite ou déjà à jour, `1` = `-Check` a trouvé la page périmée.

---

## Convert-MarkdownToHtml.ps1

Les documents du service desk de ce dépôt sont en markdown, mais IT Glue et la plupart des systèmes de tickets attendent du texte enrichi. Convertir à la main signifie que le HTML est périmé dès la première modification du markdown — et `Update-TeamsClient-ITGlue.md` a changé six fois en deux jours — c'est pourquoi la page est générée.

Pris en charge, parce que ces documents l'utilisent : titres, tableaux avec ligne d'en-tête, blocs de code délimités (y compris ceux indentés dans les étapes numérotées), citations, listes numérotées et à puces, lignes horizontales, ainsi que code en ligne, gras, italique et liens. Tout le reste passe tel quel comme texte, plutôt que d'être interprété au hasard.

Le CSS est intégré, la page est donc autonome — rien à héberger, et rien qui casse lorsque le fichier est copié ailleurs. Les éléments vides sont écrits auto-fermants (`<hr/>`, `<br/>`), de sorte que la sortie s'analyse aussi bien en XML qu'en HTML et peut être vérifiée structurellement plutôt qu'à l'œil.

**Paramètres**

| Paramètre | Description |
|-----------|-------------|
| `-Path` | Le fichier markdown à convertir (obligatoire) |
| `-Destination` | Où écrire le HTML (par défaut : même dossier et même nom, `.html`) |
| `-Title` | Titre de la page et du navigateur (par défaut : le premier titre `#` du document) |
| `-Check` | N'écrit rien ; sortie `1` lorsque le HTML sur le disque ne correspond plus au markdown |
| `-WhatIf` | Affiche ce qui serait écrit et ne modifie rien |

**Exemples**

```powershell
# Générer la page IT Glue à côté de sa source markdown
pwsh -File scripts/Startup/Convert-MarkdownToHtml.ps1 -Path scripts/Device/Update-TeamsClient-ITGlue.md

# La page commitée est-elle en retard ? Code de sortie 1 si c'est le cas
pwsh -File scripts/Startup/Convert-MarkdownToHtml.ps1 -Path scripts/Device/Update-TeamsClient-ITGlue.md -Check
```

Depuis le menu : `menu.ps1`, touche **M**. Il propose le document Teams pour IT Glue comme chemin par défaut et demande s'il faut seulement vérifier.

**Remarques**

- **Le mettre dans IT Glue :** ouvrez le `.html` dans un navigateur, sélectionnez tout, copiez et collez dans l'éditeur de documents d'IT Glue. L'éditeur conserve les titres, tableaux et blocs de code et abandonne le CSS — ce qui est souhaitable ici, puisqu'IT Glue applique le sien.
- La ligne de date de génération est exclue de la comparaison `-Check`, donc le relancer sur un document inchangé ne signale aucune différence.
- Relancez-le après avoir modifié le markdown. `-Check` est ce qu'appellerait un hook pre-commit ou un pipeline.

Codes de sortie : `0` = écrite ou déjà à jour, `1` = `-Check` a trouvé la page périmée, ou la page n'existe pas encore.

---

## Test-MarkdownLinks.ps1

Parcourt chaque fichier `.md` du dépôt et signale les liens qui ne mènent nulle part. Un lien
mort dans un readme est invisible jusqu'à ce que quelqu'un clique dessus, c'est-à-dire généralement au moment
précis où il en avait besoin.

Il vérifie deux types de liens :

| Type | Ce qui peut mal tourner |
|------|-------------------|
| Un lien vers un fichier ou un dossier | Le fichier a été renommé, déplacé ou supprimé et le readme pointe encore vers l'ancien chemin. Les espaces encodés en pourcentage (`Time%20sync/readme.md`) sont décodés avant le test du chemin, car c'est ainsi que GitHub les sert |
| Une ancre dans la page (`#set-usermanagerps1`) | Le titre visé a été renommé, ou l'ancre a été tapée à la main et n'a jamais correspondu. Ces liens se dégradent en silence — rien ne vous avertit |

Les ancres sont résolues comme GitHub les construit : le titre est mis en minuscules, la mise en forme
markdown est retirée, tout ce qui n'est pas une lettre, un chiffre, une espace, `_` ou `-` est
supprimé, et les espaces deviennent des tirets — ainsi `### Watch-RDSLive.ps1` donne `#watch-rdsliveps1`,
et non `#watch-rdslivesps1`. Les titres répétés reçoivent le suffixe `-1`, `-2` de GitHub. Les caractères
invisibles (sélecteurs de variante, zero-width joiners — les octets qui font d'un emoji un
emoji) sont retirés du titre comme du lien avant la comparaison, pour qu'un
titre à emoji dans la table des matières ne soit pas considéré comme cassé.

Les blocs de code délimités et le code en ligne sont ignorés, afin qu'un readme qui *documente* la syntaxe
des liens ne se signale pas lui-même comme cassé — l'exemple `([docs](#…))` deux paragraphes plus haut
est un texte sur les liens, pas un lien.

Les liens externes (`http`, `https`, `mailto`) sont comptés mais pas récupérés : il s'agit d'une
vérification structurelle, et elle doit fonctionner hors ligne.

**Paramètres**

| Paramètre | Description |
|-----------|-------------|
| `-Root` | Racine du dépôt (par défaut : deux niveaux au-dessus de ce script) |
| `-Path` | Vérifie un seul fichier ou dossier au lieu de tout le dépôt |

**Exemples**

```powershell
# Vérifier chaque readme du dépôt
pwsh -File scripts/Startup/Test-MarkdownLinks.ps1

# Un seul dossier de charge de travail
pwsh -File scripts/Startup/Test-MarkdownLinks.ps1 -Path scripts/Exchange
```

Codes de sortie : `0` = chaque lien interne est résolu, `1` = quelque chose est cassé (chaque cas
listé avec le fichier dans lequel il se trouve et la raison de l'échec).

---

## Update-ReadmeHeader.ps1

Chaque dossier possède son readme en trois exemplaires : `readme.md` (anglais), `readme.nl.md`
(néerlandais) et `readme.fr.md` (français). Chacun commence par les deux mêmes lignes, qui ne
diffèrent que par leurs chemins : un sélecteur de langue vers la même page dans les autres
langues, et un fil d'Ariane pour remonter l'arborescence, dont chaque niveau pointe vers son
propre readme **dans la langue courante**.

Tenus à la main, ce sont précisément ces chemins qui se dégradent — un `../` de trop peu après
le déplacement d'un dossier, ou une page néerlandaise qui pointe vers le parent anglais. Ils sont
donc générés à partir du dossier où se trouve le readme, et de rien d'autre. Tout ce qui précède
le premier titre et qui est une ligne de sélecteur ou de fil d'Ariane est remplacé ; le reste du
fichier n'est pas modifié.

Un dossier qui a un `readme.md` mais pas de version néerlandaise ou française est signalé : son
sélecteur de langue pointerait vers un fichier inexistant.

**Paramètres**

| Paramètre | Description |
|-----------|-------------|
| `-Root` | Racine du dépôt (par défaut : deux niveaux au-dessus de ce script) |
| `-Check` | N'écrit rien ; code de sortie `1` si un en-tête est obsolète ou si une version linguistique manque |
| `-WhatIf` | Indique quels en-têtes seraient réécrits, sans rien écrire |

**Exemples**

```powershell
# Après l'ajout ou le déplacement d'un readme de dossier (écrivez d'abord les versions .nl.md et .fr.md)
pwsh -File scripts/Startup/Update-ReadmeHeader.ps1

# Échouer si un en-tête est obsolète ou une traduction manquante — pour un hook ou un pipeline
pwsh -File scripts/Startup/Update-ReadmeHeader.ps1 -Check
```

> Exécutez ensuite [`Test-MarkdownLinks.ps1`](#test-markdownlinksps1) : ce script écrit les
> liens, celui-là prouve qu'ils aboutissent.
