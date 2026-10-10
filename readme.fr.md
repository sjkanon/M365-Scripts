[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

# M365-Scripts

> Une collection de scripts PowerShell et d'outils de gestion M365 pour les ingénieurs MSP, maintenue par Sjoerd Kanon.

---

## Sommaire

- [Dossiers](#dossiers)
- [Premiers pas](#premiers-pas)
- [Trouver un script](#trouver-un-script)
- [Lanceur rapide](#lanceur-rapide)
- [Prérequis](#prérequis)
- [Menu](#menu)
- [Catégories de scripts](#catégories-de-scripts)
  - [Gestion M365](#️-gestion-m365)
  - [Exchange](#-exchange)
  - [Entra ID / Graph](#-entra-id--graph)
  - [Intune et Autopilot](#-intune-et-autopilot)
  - [SharePoint et OneDrive](#-sharepoint-et-onedrive)
  - [Tests et diagnostics](#-tests-et-diagnostics)
  - [Rapports](#-rapports)
  - [Infrastructure et appareils](#️-infrastructure-et-appareils)
  - [Outils maison](#-outils-maison)
  - [Infrastructure Azure](#️-infrastructure-azure)
  - [Réécritures des anciennes boîtes à outils](#️-réécritures-des-anciennes-boîtes-à-outils)
- [Structure du dépôt](#structure-du-dépôt)
- [Contribuer](#contribuer)
- [Historique des versions](#historique-des-versions)

---

## Dossiers

Chaque charge de travail a son propre dossier sous [`scripts/`](scripts/readme.fr.md), et chaque dossier a un readme : à quoi sert chaque script, ses paramètres, des exemples et des remarques. Commencez ici et naviguez ; chaque readme comporte en haut un fil d'Ariane pour remonter.

| Dossier | Description |
|--------|-------------|
| [`ActiveDirectory/`](scripts/ActiveDirectory/readme.fr.md) | Surveillance d'AD DS sur site (surveillance des verrouillages de compte) — cible directement un DC/serveur de fichiers, pas Entra ID |
| [`Azure/`](scripts/Azure/readme.fr.md) | Gestion des VM Azure IaaS (conversion du contrôleur de disque) — cible directement Azure via `Az`, pas le tenant M365 |
| [`Entra/`](scripts/Entra/readme.fr.md) | Cycle de vie des utilisateurs, attribution des responsables, rapports de licences, référentiel Conditional Access, fenêtres CA temporaires, codes TAP, audit des groupes M365 (Microsoft Graph) |
| [`Exchange/`](scripts/Exchange/readme.fr.md) | Migration/droits des calendriers, groupes de distribution, audits des boîtes aux lettres/calendriers/DKIM/transferts |
| [`Graph/`](scripts/Graph/readme.fr.md) | Gestion des autorisations d'application Microsoft Graph |
| [`Intune/`](scripts/Intune/readme.fr.md) | Inscription Autopilot, mise à jour de la stratégie de conformité iOS, déploiement du fond d'écran/écran de verrouillage de l'entreprise |
| [`SharePoint/`](scripts/SharePoint/readme.fr.md) | Opérations sur le contenu SharePoint Online / OneDrive — restauration de la corbeille par site ou à l'échelle du tenant (PnP PowerShell, inscription d'application automatique), et où est passé un fichier : renommé, déplacé ou supprimé (journal d'audit) |
| [`Reporting/`](scripts/Reporting/readme.fr.md) | Rapport de dernière connexion des ordinateurs, rapport de stockage SharePoint, rapport mensuel des licences |
| [`Device/`](scripts/Device/readme.fr.md) | Maintenance des postes Windows — activation, nettoyage, fichiers temporaires, synchronisation de l'heure, audio, diagnostic OpenVPN, disque temporaire + fichier d'échange Azure/AVD, pilotes d'imprimante + imprimantes depuis un fichier JSON |
| [`Linux/`](scripts/Linux/readme.fr.md) | Serveurs Linux (Debian/Ubuntu, 3CX Phone System) — nettoyage du disque en bash : paquets, journal, journaux, fichiers temporaires, caches utilisateur, Docker, journaux et sauvegardes 3CX |
| [`Network/`](scripts/Network/readme.fr.md) | Vérification de ports TCP, diagnostic d'authentification/réseau, test de charge des E/S fichiers |
| [`RDS/`](scripts/RDS/readme.fr.md) | Diagnostic des connexions RDP / RD Web Access, surveillance des sessions en direct, diagnostic et réduction des disques de profil FSLogix, préparation de l'image des hôtes de session (Teams, Outlook, Copilot) |
| [`SMTP/`](scripts/SMTP/readme.fr.md) | Tests de connectivité d'un relais SMTP (ponctuels et récurrents) |
| [`Deployment/`](scripts/Deployment/readme.fr.md) | Boîte à outils USB pour l'installation de Windows et l'inscription Autopilot pendant l'OOBE |
| [`DNS/`](scripts/DNS/readme.fr.md) | Résoudre des enregistrements DNS et les importer dans des zones DNS intégrées à AD |
| [`SAS/`](scripts/SAS/readme.fr.md) | Surveillance des erreurs des traitements batch SAS avec intégration Zabbix |
| [`Teams/`](scripts/Teams/readme.fr.md) | Export et archivage Microsoft Teams / SharePoint |
| [`Startup/`](scripts/Startup/readme.fr.md) | Bibliothèque de fonctions M365 [`functies.ps1`](scripts/Startup/functies.ps1) + amorçage des modules + vérificateur de syntaxe, chargés en dot-source par le menu |
| [`Custom Scripts/`](scripts/Custom%20Scripts/readme.fr.md) | Scripts liés à leur chemin — déploiement du thème Office (son URL de téléchargement pointe en dur vers ce chemin du dépôt) |
| [`TenantOnboarding/`](scripts/TenantOnboarding/readme.fr.md) | Provisionnement de nouveaux tenants, rapports multi-tenant/GDAP, déploiement d'applications, configuration des appareils, gestion de OneDrive, gestion des utilisateurs — modernisé à partir d'une boîte à outils interne de mise en place de tenants, aujourd'hui retirée |
| [`Office365Toolkit/`](scripts/Office365Toolkit/readme.fr.md) | Réécritures Security/Exchange/Intune des fonctionnalités encore utiles de la boîte à outils retirée `directorcia/Office365` (CIAOPS) |
| [`PatronToolkit/`](scripts/PatronToolkit/readme.fr.md) | Réécritures Entra/Exchange/Intune/Security/SharePoint/Teams des fonctionnalités encore utiles de la boîte à outils retirée `directorcia/patron` |
| [`LegacyUtilities/`](scripts/LegacyUtilities/readme.fr.md) | Scripts divers modernisés (Exchange, Entra, Teams, Network, Device, Workspace 365) issus de petits outils variés de la boîte à outils interne retirée |

Vous connaissez le nom du script mais pas le dossier ? [`scripts/INDEX.md`](scripts/INDEX.md) liste tous les scripts de A à Z.

---

## Premiers pas

```powershell
.\load.ps1
```

Au premier lancement, [`load.ps1`](load.ps1) va :

1. Demander votre UPN d'administrateur et votre nom d'affichage — enregistrés dans un `load.config.ps1` ignoré par git
2. Demander si vous voulez le mode GDAP délégué par défaut, et enregistrer éventuellement un domaine client par défaut
3. Demander si Graph doit utiliser par défaut la connexion par code d'appareil
4. Vérifier les modules requis — manquants, plus anciens que leur minimum, ou avec une mise à jour sur la PowerShell Gallery — et les installer ou les mettre à jour
5. Importer les modules principaux
6. Ouvrir le menu interactif

Par la suite, il saute les questions. La vérification des modules (étape 4) s'exécute à chaque démarrage sans rien demander ; elle interroge la galerie au plus une fois toutes les 24 heures, et si tout est à jour elle affiche une seule ligne. `.\load.ps1 -SkipModuleCheck` la saute une fois.

Pour lancer le lanceur automatiquement à l'ouverture de session Windows :

```powershell
.\load.ps1 -SetupStartup
```

Pour supprimer plus tard le raccourci de démarrage :

```powershell
.\load.ps1 -RemoveStartup
```

> Vous pouvez aussi lancer [`.\menu.ps1`](menu.ps1) directement — il vous demandera alors votre UPN en solution de repli.
> Pour réinstaller ou mettre à jour les modules manuellement : [`.\scripts\Startup\Install-Modules.ps1`](scripts/Startup/Install-Modules.ps1) ou [`.\scripts\Startup\Update-Modules.ps1`](scripts/Startup/Update-Modules.ps1). La liste des modules est [`RequiredModules.psd1`](scripts/Startup/RequiredModules.psd1) — ajoutez-y un module et chaque machine l'installe au prochain démarrage. Le hook de docs signale un module chargé par un script mais absent de la liste.

---

## Trouver un script

| Où | Ce que vous y trouvez |
|-------|-------------------|
| [`scripts/INDEX.md`](scripts/INDEX.md) | Tous les scripts de A à Z sur une seule page — nom, dossier et ce qu'il fait. Faites Ctrl-F dessus quand vous savez à peu près ce que vous cherchez, mais pas où il se trouve |
| [`scripts/readme.md`](scripts/readme.fr.md) | Dans l'autre sens : à quoi sert chaque dossier de charge de travail |
| [`.\menu.ps1`](menu.ps1) | Le lanceur interactif sélectionné pour les tâches courantes |
| `f <term>` | Recherche approximative depuis votre shell, décrite plus bas sous [Lanceur rapide](#lanceur-rapide) |
| Tout readme de dossier | Chaque nom de script dans un tableau `Scripts` renvoie directement au fichier, avec un lien `docs` vers sa section sur la même page |

[`INDEX.md`](scripts/INDEX.md) est généré à partir des en-têtes `.SYNOPSIS` des scripts eux-mêmes par [`scripts/Startup/Update-ScriptIndex.ps1`](scripts/Startup/Update-ScriptIndex.ps1) — relancez-le (ou lancez-le avec `-Check`) chaque fois qu'un script est ajouté, renommé, déplacé ou supprimé.

Ces liens sont vérifiés, pas supposés : [`scripts/Startup/Test-MarkdownLinks.ps1`](scripts/Startup/Test-MarkdownLinks.ps1) parcourt chaque readme et échoue sur un fichier absent, ou sur une ancre sans titre correspondant.

---

## Lanceur rapide

[`f.ps1`](f.ps1) crée une commande courte par script dans [scripts/](scripts/), pour que vous n'ayez plus à naviguer d'abord vers un dossier. Chargez-le en dot-source depuis votre profil :

```powershell
notepad $PROFILE
. "C:\Users\<you>\Git\M365-Scripts\f.ps1"
```

Le point initial n'est pas facultatif — sans lui, le fichier s'exécute dans sa propre portée et les commandes disparaissent aussitôt. `.\f.ps1 -Install` écrit la ligne pour vous (et ne fait rien si elle est déjà présente).

Après avoir redémarré votre shell :

```powershell
f-test-dkimconfig -Domain contoso.com
f-dkimconfig -Domain contoso.com          # forme courte, quand le nom est unique
f-get-mailboxsizes
f-m365                                    # le catalogue complet
f-m365 mailbox                            # filtré
```

Les wrappers copient le bloc de paramètres du script cible depuis son AST, si bien que `-Dom<Tab>` se complète et qu'un `ValidateSet` est appliqué avant toute exécution. Les valeurs par défaut sont volontairement retirées du wrapper : seuls les paramètres liés sont transmis, de sorte que les valeurs par défaut du script continuent de s'appliquer.

### Recherche approximative

Quand vous savez à peu près comment s'appelle un script mais pas exactement, `f` cherche dans le nom, le dossier et le `.SYNOPSIS` :

```powershell
f dkim                    # une correspondance -> l'exécute
f entra group             # plusieurs correspondances -> sélecteur numéroté
f dkim -Domain contoso.com
f -List mailbox           # affiche les correspondances, n'exécute rien
f -Show trace             # chemin, synopsis et paramètres
f -Edit bloatware         # ouvre dans $env:EDITOR, VS Code ou notepad
```

### Remarques

| | |
|---|---|
| Nommage | `f-<nom-complet-du-script>` existe toujours ; `f-<noun>` n'est ajouté que là où il reste sans ambiguïté. Retirer le verbe provoque ici 11 collisions (`Detect-`, `Install-` et `Uninstall-ClaudeDesktop-Intune` donnent le même nom), donc c'est sur le nom complet que vous pouvez toujours compter. |
| Cache | Les wrappers générés sont stockés dans `.f-index.json` (ignoré par git). Analyser chaque script coûte ~500 ms, lire le cache ~30 ms, et c'est ce qui garde le démarrage du shell rapide. Il se rafraîchit de lui-même quand un script est ajouté, supprimé ou modifié ; `f-refresh` force le rafraîchissement. |
| Collisions | Les commandes existantes ne sont jamais écrasées. Un profil peut en charger plusieurs — `ScriptRunner.Profile.ps1` dans *itce-testing* possède `f-scripts`, c'est pourquoi ce catalogue s'appelle `f-m365`. Tout ce qui est ignoré est signalé au chargement. |
| Désinstallation | `.\f.ps1 -Uninstall` retire le bloc balisé de `$PROFILE`. Une ligne de dot-source écrite à la main est signalée, pas supprimée. |

---

## Prérequis

| Prérequis | Détails |
|-------------|---------|
| PowerShell | 7.0+ (multiplateforme) ; certains scripts prennent en charge PS 5.1 sous Windows |
| Autorisations | Droits d'administration Microsoft 365 pour la charge de travail visée |
| Connexion | Chaque script M365 se connecte via [`Connect-M365.ps1`](scripts/Startup/readme.fr.md#connect-m365ps1) : Microsoft Graph d'abord, **délégué par défaut** (code d'appareil et client GDAP depuis `load.config.ps1`), app-only avec `-ClientId`/`-CertificateThumbprint` ou `-AppOnly` (`graph.appid.json`). Exchange Online, Teams et PnP uniquement là où Graph n'a pas d'API |
| Execution Policy | Windows uniquement : `Set-ExecutionPolicy RemoteSigned -Scope CurrentUser` |

---

## Menu

Le lanceur ([`menu.ps1`](menu.ps1)) couvre tous les outils de ce dépôt. Appuyez sur une touche pour lancer :

| Touche | Catégorie | Outil |
|-----|----------|------|
| `1` / `F1` | Testing | [Test-Ports](scripts/Network/Test-Ports.ps1) — vérification de ports TCP |
| `2` / `F2` | Exchange | [Migrate-Calendar](scripts/Exchange/Migrate-Calendar.ps1) |
| `3` / `F3` | Exchange | [Set-Calendar-rights](scripts/Exchange/Set-Calendar-rights.ps1) |
| `4` / `F4` | Testing | [Test-SMTP (ponctuel)](scripts/SMTP/testsmtp.ps1) |
| `5` / `F5` | Testing | [Test-SMTP (toutes les 5 min)](scripts/SMTP/testsmtp_5min.ps1) |
| `6` / `F6` | Device | [Restart-Time-Sync](scripts/Device/Time%20sync/Restart-Time-Sync.ps1) |
| `7` / `F7` | Device | [Detect-AudioDevices](scripts/Device/audio/detect-audiodevices.ps1) |
| `8` / `F8` | Device | [Disable-InternalMic](scripts/Device/audio/Disable-internalmic.ps1) |
| `I` | Device | [Remove-OemBloatware](scripts/Device/Remove-OemBloatware.ps1) — supprimer les bloatwares OEM + génériques du Store |
| `T` | Device | [Update-TeamsClient](scripts/Device/Update-TeamsClient.ps1) — mettre à jour le nouveau Teams + le complément de réunion Outlook s'ils sont obsolètes |
| `R` | Device | [Repair-AppxPackageStore](scripts/Device/Repair-AppxPackageStore.ps1) — réparer les paquets AppX qui échouent avec 0x80070490 (Teams, nouvel Outlook, FSLogix) |
| `K` | Device | [FSLogix-Shrink](scripts/RDS/Invoke-FSLogixShrink.ps1) — réduire les disques de profil FSLogix d'un partage, ou vérifier la compaction à la déconnexion |
| `J` | Device | [Update-SessionHostImage](scripts/RDS/Update-SessionHostImage.ps1) — préparer une image multisession / des hôtes AVD pour Teams, Outlook et Copilot avec FSLogix |
| `Y` | Device | [Watch-M365Apps](scripts/RDS/Watch-M365Apps.ps1) — watchdog : tester Teams, Outlook et Copilot avec les comptes IT, réparer, signaler à n8n |
| `Q` | Device | [Get-M365AppsLog](scripts/RDS/Get-M365AppsLog.ps1) — collecter ce qui s'est passé avec Teams, Outlook et Copilot par utilisateur, et ce qu'a fait le watchdog |
| `N` | Device | [Install-Printer](scripts/Device/Printer/Install-Printer.ps1) — installer des pilotes d'imprimante (depuis GitHub) et des imprimantes à partir d'un fichier JSON |
| `9` / `F9` | Startup | [Install-Modules](scripts/Startup/Install-Modules.ps1) |
| `U` | Startup | [Update-Modules](scripts/Startup/Update-Modules.ps1) — vérifier/mettre à jour les modules requis, au choix aussi tous les autres modules installés |
| `Z` | Startup | [Test-RequiredModules](scripts/Startup/Test-RequiredModules.ps1) — modules chargés par des scripts mais absents de `RequiredModules.psd1` |
| `X` | Startup | [Update-ScriptIndex](scripts/Startup/Update-ScriptIndex.ps1) — reconstruire [`scripts/INDEX.md`](scripts/INDEX.md), la liste A–Z de tous les scripts |
| `L` | Startup | [Test-MarkdownLinks](scripts/Startup/Test-MarkdownLinks.ps1) — vérifier chaque lien des readmes : fichiers et ancres internes à la page |
| `M` | Startup | [Convert-MarkdownToHtml](scripts/Startup/Convert-MarkdownToHtml.ps1) — produire une page HTML mise en forme à partir d'un document markdown, pour IT Glue |
| `A` / `F10` | Reporting | [Licensing-Report](scripts/Reporting/Licensing/genereer_rapport.ps1) |
| `P` | Reporting | [SharePoint-Perms](scripts/Reporting/Get-SharePointPermissionsReport.ps1) — indiquer qui a accès à quoi, à chaque niveau |
| `S` | SharePoint | [SharePoint-Structure](scripts/SharePoint/Provisioning/readme.fr.md) — provisionner/vérifier les métadonnées, bibliothèques et droits |
| `F` | Startup | [Enable-LauncherStartup](menu.ps1) — ajouter le lanceur au démarrage de Windows |
| `G` | Startup | [Disable-LauncherStartup](menu.ps1) — retirer le lanceur du démarrage de Windows |
| `B` | M365 | [Connect-Tenant](scripts/Startup/readme.fr.md#functiesps1) |
| `H` | M365 | [Test-GdapConnection](scripts/Startup/readme.fr.md#functiesps1) — valider l'accès GDAP délégué |
| `C` | M365 | [Sous-menu Exchange Online](scripts/Startup/readme.fr.md#functiesps1) |
| `D` | M365 | [Sous-menu Entra ID / Graph](scripts/Startup/readme.fr.md#functiesps1) |
| `E` | M365 | [Sous-menu MSP Admin](scripts/Startup/readme.fr.md#functiesps1) |

Les options M365 (`B`, `C`, `D`, `E`, `H`) ne chargent [`functies.ps1`](scripts/Startup/functies.ps1) qu'à la première utilisation — l'authentification Graph n'est déclenchée qu'en cas de besoin.

**Sous-menu Exchange (`C`)**

| Touche | Outil |
|-----|------|
| `8` | [Test-CalendarPermissions](scripts/Exchange/Test-CalendarPermissions.ps1) — auditer les droits sur les dossiers de calendrier (toutes les boîtes aux lettres ou une seule) |
| `9` | [Test-MailboxPermissions](scripts/Exchange/Test-MailboxPermissions.ps1) — auditer Full Access, Send As, Send on Behalf |
| `A` | [Test-GroupPermissions](scripts/Exchange/Test-DistributionGroupPermissions.ps1) — auditer les gestionnaires des groupes de distribution, Send As, Send on Behalf, nombre de membres |
| `B` | [Test-DkimConfig](scripts/Exchange/Test-DkimConfig.ps1) — valider la configuration de signature DKIM et les enregistrements DNS CNAME/TXT |
| `C` | [Get-ExternalForwards](scripts/Exchange/Get-ExternalForwards.ps1) — auditer les boîtes aux lettres avec un transfert externe configuré |
| `D` | [Get-MailboxSizes](scripts/Exchange/Get-MailboxSizes.ps1) — rapport de taille des boîtes aux lettres trié par stockage utilisé |
| `E` | [Move-InboxToArchive](scripts/Exchange/Move-InboxToArchive.ps1) — archiver les messages de la boîte de réception dans le dossier Archive |
| `F` | [Set-DL-Dynamic-Static](scripts/Exchange/Set-Distributionlist-dynamic-static.ps1) — convertir un groupe de distribution dynamique en groupe statique |
| `H` | [Get-CalendarMappings](scripts/Exchange/Get-CalendarMappings.ps1) — où un calendrier est ajouté dans Outlook, à côté des droits (recherche par mot-clé, p. ex. `balie`, ou toutes/certaines boîtes aux lettres) |
| `I` | [Convert-SharedCalendar](scripts/Exchange/Convert-SharedCalendarToResource.ps1) — déplacer un calendrier partagé hors de la boîte aux lettres d'un utilisateur vers une boîte aux lettres de salle/d'équipement (toujours un aperçu d'abord) |
| `J` | [Move-SharedCalendar](scripts/Exchange/Move-SharedCalendar.ps1) — tout-en-un : trouver un calendrier par mot-clé, le déplacer vers une boîte aux lettres de ressource, lister qui doit basculer |
| `K` | [Get-DLMembers](scripts/Exchange/Get-DistributionGroupMembers.ps1) — exporter chaque liste de distribution avec ses membres vers Excel, ou seulement les listes contenant une adresse, un domaine ou une arborescence de domaines (`-Recurse` pour développer les listes imbriquées) |
| `L` | [Restore-MailboxMessages](scripts/Exchange/Restore-MailboxMessages.ps1) — remettre en place le courrier déplacé ou supprimé un jour donné, ou depuis une date jusqu'à maintenant, et montrer qui l'a fait (toujours un aperçu d'abord) |

**Sous-menu Entra ID (`D`)**

| Touche | Outil |
|-----|------|
| `A` | [Test-M365GroupMembership](scripts/Entra/Test-M365GroupMembership.ps1) — auditer les propriétaires et membres des groupes M365 / Teams |
| `B` | [New-M365User](scripts/Entra/New-M365User.ps1) — créer un nouvel utilisateur (mot de passe généré automatiquement, licence facultative) |
| `C` | [Import-M365Users](scripts/Entra/Import-M365Users.ps1) — créer des utilisateurs en masse depuis un CSV, essai à blanc par défaut |
| `D` | [New-TemporaryCA](scripts/Entra/New-TemporaryConditionalAccessPolicy.ps1) — créer une stratégie Conditional Access temporaire pour un utilisateur/groupe (durée ou date et heure de début/fin) |
| `E` | [Remove-TemporaryCA](scripts/Entra/Remove-TemporaryConditionalAccessPolicies.ps1) — supprimer les stratégies CA temporaires expirées ou toutes |
| `F` | [New-UserTAP](scripts/Entra/New-UserTemporaryAccessPass.ps1) — créer un Temporary Access Pass pour un utilisateur |
| `G` | [Get-M365UserLicenses](scripts/Entra/Get-M365UserLicenses.ps1) — rapporter les licences attribuées à un ensemble d'utilisateurs |
| `H` | [Import-CA-Baseline](scripts/Entra/Import-ConditionalAccessBaseline.ps1) — importer le référentiel Conditional Access de la communauté |
| `I` | [Set-UserManager](scripts/Entra/Set-UserManager.ps1) — rapporter/définir en masse le responsable d'un ensemble d'utilisateurs |

---

## Catégories de scripts

### ☁️ Gestion M365

📂 Dossier : [`Startup/`](scripts/Startup/readme.fr.md)

Fonctions interactives de gestion M365 via Microsoft Graph et Exchange Online. Chargées comme bibliothèque via le menu. Les exports CSV et journaux vont dans `C:\Temp\` sous Windows ou `~/Downloads/` sous macOS.

| Domaine | Fonctionnalités |
|------|----------|
| Exchange Online | Accès aux boîtes aux lettres partagées, paramètres régionaux, alias, groupes de distribution, réponse automatique, copie des éléments envoyés |
| Entra ID / Graph | Administrateurs du tenant, domaines, licences, utilisateurs, réinitialisation de mot de passe, journaux de connexion, création/suppression en masse, fenêtres CA temporaires, codes TAP |
| MSP Admin | Créer/gérer un compte d'administration MSP sur l'ensemble des tenants clients |

---

### 📧 Exchange

📂 Dossier : [`Exchange/`](scripts/Exchange/readme.fr.md)

Scripts de gestion des calendriers et des boîtes aux lettres.

- Migration de calendrier entre utilisateurs
- Définir les droits sur les dossiers de calendrier (prise en charge des paramètres régionaux NL/FR/EN)
- **[Get-DistributionGroupMembers.ps1](scripts/Exchange/Get-DistributionGroupMembers.ps1)** — qui figure dans quelle liste de distribution, sous forme d'un classeur Excel unique prêt à être transmis au client
  - Feuille `Overzicht` (une ligne par liste) et feuille `Leden` (une ligne par membre), toutes deux des tableaux filtrables avec une ligne d'en-tête figée, en-têtes en néerlandais
  - `-Member jan@contoso.com` répond à « dans quelles listes figure cette personne ? » ; `-Member @be.verizon.com` y répond pour un domaine et `-Member *.verizon.com` pour un domaine et tous ses sous-domaines, en tenant compte des alias et de `ExternalEmailAddress` pour que les contacts externes soient réellement trouvés
  - `-Recurse` développe les listes imbriquées — sans cette option, une personne qui ne reçoit le courrier que via un groupe imbriqué est invisible, et un filtre signale « aucun résultat » sur une liste qui lui distribue pourtant le courrier
- **Get-MessageTraceReport.ps1** — retrouver qui a reçu quoi, à quelle heure exacte, et vers où cela a été transféré
- **[Remove-PhishingMessage.ps1](scripts/Exchange/Remove-PhishingMessage.ps1)** — supprimer un message de phishing d'une, de plusieurs ou de toutes les boîtes aux lettres ; essai à blanc par défaut
  - Deux moteurs : **Purview** Content Search + purge (à l'échelle du tenant, le seul capable de HardDelete) et **Graph** (par boîte aux lettres, sans délai de l'index de recherche, rapport par message)
  - `Recycle` / `SoftDelete` / `HardDelete` ; refuse de s'exécuter sans sélecteur de contenu, pour qu'une simple plage de dates ne puisse jamais correspondre à tous les messages
  - Enchaîne automatiquement les passes de purge pour contourner la limite de Purview de 10 éléments par boîte aux lettres, et écrit un CSV de tout ce qui a été trouvé et supprimé
- **[Restore-MailboxMessages.ps1](scripts/Exchange/Restore-MailboxMessages.ps1)** — remettre en place les messages déplacés ou supprimés un jour donné, et indiquer qui l'a fait ; aperçu par défaut
  - Les messages supprimés reviennent via `Restore-RecoverableItems` (Deleted Items, Recoverable Items, Purges) ; les messages déplacés sont retracés jusqu'à leur dossier d'origine via le journal d'audit et remis en place via Graph
  - Nomme l'auteur à partir du Unified Audit Log — compte, propriétaire/délégué/administrateur, client, IP — et indique quelles actions ne sont pas auditées sur la boîte aux lettres

---

### 👤 Entra ID / Graph

📂 Dossier : [`Entra/`](scripts/Entra/readme.fr.md), [`Graph/`](scripts/Graph/readme.fr.md)

Scripts de gestion du cycle de vie des utilisateurs via Microsoft Graph.

- Créer un utilisateur M365 (mot de passe généré automatiquement, licence facultative)
- Créer des utilisateurs en masse depuis un CSV — essai à blanc par défaut, mots de passe dans le CSV de sortie
- Supprimer des utilisateurs en masse depuis un CSV — essai à blanc par défaut, rapport CSV

---

### 📱 Intune et Autopilot

📂 Dossier : [`Intune/`](scripts/Intune/readme.fr.md)

Scripts d'inscription des appareils, d'enregistrement Autopilot et de gestion des stratégies de conformité.

- Récupérer les informations matérielles Windows Autopilot
- Outil d'aide CMD pour l'inscription Autopilot
- **[Compare-IntuneConfig.ps1](scripts/Intune/Compare-IntuneConfig.ps1)** — comparer la configuration Intune d'un tenant client à une sauvegarde du référentiel MSP (détection de dérive), via le module `IntuneBackupAndRestore` — en lecture seule
- **iOS Compliance Updater** — maintient automatiquement à jour la version iOS minimale exigée dans Intune
  - Récupère la dernière version d'iOS depuis le flux RSS d'Apple (avec repli sur la page Apple Support)
  - La compare au minimum actuel de la stratégie et la met à jour via l'API Microsoft Graph
  - Configuration unique via [`Setup.ps1`](scripts/Intune/iOS-Compliance-Updater/Setup.ps1) (crée l'App Registration, attribue les autorisations, écrit `config.json`)
  - S'exécute chaque semaine comme tâche planifiée Windows (SYSTEM, tous les lundis à 07:00)
  - Mode essai à blanc (`-WhatIf`) — montre ce qui changerait sans l'appliquer
- **Desktop** — ajouter l'écran de verrouillage au menu Démarrer et au bureau ; définir le fond d'écran de l'entreprise via Intune :
  - [`Set-CorporateWallpaper.ps1`](scripts/Intune/Desktop/Background/Desktop/Set-CorporateWallpaper.ps1) — générique, réutilisable pour chaque client ; seul le bloc CONFIGURATION doit être adapté
  - [`Make-lockscreen.ps1`](scripts/Intune/Desktop/Background/Lockscreen/Make-lockscreen.ps1) — applique la même image d'entreprise comme écran de verrouillage Windows via PersonalizationCSP
  - Télécharge le fond d'écran depuis une URL publique ; compare le hachage SHA256 au fichier existant — l'ignore s'il est déjà à jour, l'applique s'il est nouveau ou modifié
  - Le flux de l'écran de verrouillage télécharge depuis Internet via `Invoke-WebRequest`, valide les en-têtes d'image (`jpg/png/bmp`), bloque les réponses HTML et normalise les URL GitHub blob/raw courantes
  - Applique via PersonalizationCSP (application MDM), WinAPI (immédiat), le registre HKCU (style) et le profil Default User (nouveaux comptes)
  - Journal : `C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\CorporateWallpaper-<CLIENTNAME>.log`
  - Journal de l'écran de verrouillage : `C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\CorporateLockscreen-<CLIENTNAME>.log`
  - Déploiement via Intune : **Run as SYSTEM**, PowerShell 64 bits

  | Variable | Description |
  |---|---|
  | `$ImageUrl` | URL publique de l'image de fond d'écran (PNG ou JPG) |
  | `$WallpaperStyle` | `10` = Remplir · `6` = Ajuster · `2` = Étirer · `0` = Mosaïque · `22` = Étendre |
  | `$ClientName` | Nom du client — utilisé dans le nom du fichier journal et le chemin local de l'image |

---

### 📁 SharePoint et OneDrive

📂 Dossier : [`SharePoint/`](scripts/SharePoint/readme.fr.md)

Opérations sur le contenu des sites SharePoint Online et de OneDrive via PnP PowerShell.

#### Retrouver un fichier

**[Trace-SharePointFile.ps1](scripts/SharePoint/Trace-SharePointFile.ps1)** — « où est passé mon fichier ? » pour OneDrive et SharePoint, depuis le Unified Audit Log (Exchange Online, pas PnP). Lecture seule.

- Suit les renommages, déplacements, copies, suppressions et restaurations d'un fichier par nom (caractères génériques), ancienne URL ou ID d'élément — une chaîne `A → B → C` aboutit à C
- Rejoue sur le fichier les renommages, déplacements et suppressions de dossiers, car ceux-ci déplacent chaque fichier qu'ils contiennent sans enregistrement par fichier
- Chaque heure en heure de Bruxelles avec le décalage UTC ; `-StartDate` / `-EndDate` en notation belge (`15-09-2026 08:30`), une date de fin sans heure inclut toute la journée
- Lit par jour et scinde une tranche de plus de 50 000 enregistrements, relance les recherches en échec ; dernier emplacement connu et statut par élément, CSV plus les enregistrements d'audit bruts en JSON

#### Révoquer l'accès d'un utilisateur

**[Revoke-SharePointUserAccess.ps1](scripts/SharePoint/Revoke-SharePointUserAccess.ps1)** — le pendant du rapport d'autorisations : celui-là dit qui peut accéder à quoi, celui-ci retire cet accès. Produit un rapport par défaut, supprime avec `-Apply`, et écrit un CSV de chaque autorisation trouvée et de ce qu'il en est advenu.

- D'abord l'administrateur de collection de sites, parce qu'il prime sur toutes les attributions de rôle en dessous
- Les attributions de rôle directes sur le site, un sous-site, une liste ou bibliothèque, un dossier ou un seul fichier
- L'appartenance aux groupes SharePoint, et les **liens de partage** — les groupes `SharingLinks.*` dans lesquels « Toute personne disposant du lien » et « Personnes spécifiques » placent réellement quelqu'un
- Il ne modifie volontairement jamais l'appartenance aux groupes Entra ID : un utilisateur qui accède via un groupe de sécurité ou un groupe Microsoft 365 conserve cet accès, et le retirer de SharePoint ne le lui enlève pas. Ces chemins sont signalés avec le nom du groupe, si bien que le départ d'un utilisateur se fait en deux étapes et que la seconde est visible
- Les autorisations accordées à `Everyone` sont laissées intactes pour la raison inverse — en supprimer une révoque l'accès pour tout le tenant, pas pour cette personne
- `ConfirmImpact = 'High'`, il demande donc confirmation pour chaque suppression, sauf avec `-Confirm:$false`

**[Test-SharePointAccessScripts.ps1](scripts/SharePoint/Test-SharePointAccessScripts.ps1)** vérifie ce script et le rapport d'autorisations sans toucher à un tenant : la couche d'authentification app-only qu'ils partagent doit rester identique à l'octet près, et l'entonnoir de révocation doit consigner un essai à blanc sans l'exécuter, exécuter et consigner sous `-Apply`, et continuer à refuser les autorisations qu'il ne doit pas supprimer.

#### Restauration de la corbeille

Restaurer des fichiers et dossiers supprimés depuis la corbeille d'un site ou de OneDrive — essai à blanc par défaut, `-Apply` pour restaurer réellement.

- Un site (`-SiteUrl`, fonctionne aussi pour OneDrive) ou tous les sites SharePoint du tenant (`-AllSites`) — le balayage du tenant ignore les sites OneDrive, système et verrouillés, et un site en échec n'interrompt pas l'exécution
- Restreignez le balayage avec `-SiteFilter` et testez-le d'abord sur quelques sites avec `-MaxSites`
- Filtrer par nom, dossier d'origine, auteur de la suppression et fenêtre temporelle de suppression
- Corbeille de premier niveau (utilisateur) et de second niveau (collection de sites), ou les deux
- Restaure d'abord les dossiers, les chemins les moins profonds en premier — un fichier ne peut pas être restauré dans un dossier lui-même encore supprimé
- Crée automatiquement l'inscription d'application Entra nécessaire lors de la première exécution sur un tenant, puis met l'ID client en cache dans `pnp.appid.json` (ignoré par git) — les exécutions suivantes passent directement à la connexion interactive
- `-GrantSiteAdmin` vous rend temporairement administrateur de la collection de sites, site par site, et retire ces droits ensuite — nécessaire pour le OneDrive d'un autre utilisateur et pratiquement indispensable pour `-AllSites`
- Restaure par lots de 200 au maximum via un seul appel serveur (`-BatchSize`) — un lot en échec repasse en élément par élément pour qu'un mauvais fichier n'entraîne pas les autres
- Mesure du temps tout au long : durée de lecture de la corbeille, estimation préalable, barre de progression avec ETA en direct, et durée réelle dans le résumé
- Rapport CSV de chaque élément, restauré ou en échec, avec le site, l'erreur SharePoint, le numéro de lot et la durée

#### Provisionnement de la structure

Provisionner et maintenir toute une structure SharePoint — modèle de métadonnées, types de contenu, bibliothèques et autorisations de groupes — à partir d'un seul fichier de configuration JSON. Voir [`scripts/SharePoint/Provisioning/`](scripts/SharePoint/Provisioning/readme.fr.md).

- [`Install-SharePointStructure.ps1`](scripts/SharePoint/Provisioning/Install-SharePointStructure.ps1) — **la construction en une commande** : inscrit elle-même l'application Entra, exécute les trois étapes de provisionnement dans le seul ordre qui fonctionne, vérifie le résultat et, avec `-TemporaryApp`, supprime à nouveau l'inscription d'application pour ne rien laisser dans un tenant que vous ne gérez pas au quotidien
- Le modèle vit dans la configuration, pas dans le code : un deuxième client MSP, c'est un deuxième fichier de configuration, pas un deuxième fork de quatre scripts
- [`New-SharePointMetadata.ps1`](scripts/SharePoint/Provisioning/New-SharePointMetadata.ps1) — ensemble de termes de métadonnées gérées, colonnes de site et types de contenu, sur **chaque** site de la configuration (un canal privé Teams est une collection de sites à part entière, et une colonne de site ne la traverse pas)
- [`Set-SharePointLibraries.ps1`](scripts/SharePoint/Provisioning/Set-SharePointLibraries.ps1) — bibliothèques, dossiers de canaux Teams, liaison des types de contenu, ordre des types de contenu par dossier, valeurs de colonne par défaut, affichages groupés, et un groupe de sécurité Entra ID par pilier et par niveau d'accès ; `-EnsureGroups` crée les groupes au passage
- [`Update-SharePointShareStatus.ps1`](scripts/SharePoint/Provisioning/Update-SharePointShareStatus.ps1) — déduit une colonne Deelstatus des autorisations réellement présentes sur chaque fichier (lien Anyone, invité, lien d'organisation ou rien) et signale tout élément étiqueté Intern/Vertrouwelijk exposé derrière un lien externe ; code de sortie 2 pour une tâche RMM planifiée
- [`Test-SharePointStructure.ps1`](scripts/SharePoint/Provisioning/Test-SharePointStructure.ps1) — contrôle de dérive en lecture seule qui classe chaque écart comme Missing / Different / Extra ; le code de sortie 2 signifie que quelqu'un a modifié quelque chose
- Les quatre sont idempotents et prennent en charge `-WhatIf` ; en interactif ou en app-only avec un certificat
- Les affichages transversaux par marque (`Scope = RecursiveAll`) rendent concrète l'idée de « la marque comme étiquette » : *Alles - Northwind* est une liste à plat couvrant tous les dossiers de piliers, y compris tout ce qui est étiqueté **Beide** — un fichier, deux marques, aucune copie. Plus *Nog te taggen*, *Extern gedeeld* et *Te archiveren*
- [`SharePoint-Handleiding.md`](scripts/SharePoint/Provisioning/SharePoint-Handleiding.md) — documentation utilisateur en néerlandais à remettre au client : les trois façons d'ajouter un fichier et pourquoi elles se comportent différemment, ce que signifie chaque étiquette, et ce qui se passe dès que vous étiquetez quelque chose
- Documenté plutôt que caché : des autorisations uniques sur un dossier de canal **standard**, c'est ce que ce modèle demande et ce que Microsoft ne prend pas en charge — les membres continuent de voir le canal et obtiennent une erreur dans l'onglet Fichiers. `-SkipChannelFolderPermissions` est l'alternative prudente

---

### 🧪 Tests et diagnostics

Scripts d'audit et de diagnostic, classés par charge de travail. Se connectent eux-mêmes le cas échéant — réutilisent une session existante ou se connectent automatiquement. Les exports CSV vont dans `C:\Temp\` sous Windows ou `~/Downloads/` sous macOS.

#### Exchange Online

📂 Dossier : [`Exchange/`](scripts/Exchange/readme.fr.md)

- Auditer les droits sur les dossiers de calendrier (indépendamment des paramètres régionaux, exporte un CSV)
- Auditer les délégations Full Access, Send As, Send on Behalf (exporte un CSV)
- Auditer les gestionnaires des groupes de distribution, Send As, Send on Behalf, nombre de membres (exporte un CSV)
- Valider la configuration de signature DKIM et les enregistrements DNS CNAME/TXT ; liste les actions requises
- Auditer les boîtes aux lettres qui transfèrent vers des domaines extérieurs au tenant (audit de sécurité, exporte un CSV)
- Rapporter la taille des boîtes aux lettres et le nombre d'éléments, triés par stockage utilisé (exporte un CSV)

#### Entra ID / Graph

📂 Dossier : [`Entra/`](scripts/Entra/readme.fr.md)

- Auditer les propriétaires et membres des groupes M365 (y compris Teams) — une ligne par entrée, exporte un CSV
- Créer des stratégies Conditional Access temporaires pour des fenêtres d'installation (durée ou début/fin exacts en heure locale)
- Nettoyage automatique de la stratégie CA temporaire à l'heure de fin (même session) et script de nettoyage pour les sessions manquées
- Créer des codes Temporary Access Pass (TAP) pour l'intégration et le support des utilisateurs

#### SharePoint Online

📂 Dossier : [`SharePoint/`](scripts/SharePoint/readme.fr.md)

- Rapporter l'utilisation du stockage sur tous les sites d'un tenant — tailles de fichiers actuelles + historique des versions par bibliothèque et par fichier
- En deux phases : énumère d'abord tous les sites et bibliothèques de documents, puis récupère les données de stockage
- Mode rapide (données de quota uniquement) ou analyse récursive complète avec `-Apply`

#### Réseau et connectivité

📂 Dossier : [`Network/`](scripts/Network/readme.fr.md)

- Tester la connectivité TCP vers n'importe quel hôte — ports uniques, plages (`1294:1494`), combinaisons (`80,443,1294:1494`)
- Test SMTP ponctuel avec saisie interactive des identifiants
- Test SMTP récurrent (toutes les 5 minutes) avec mot de passe chiffré enregistré
- Diagnostic d'authentification et réseau — Observateur d'événements (échecs de connexion, Kerberos, NTLM, disponibilité des DC), synchronisation de l'heure, DNS, TCP, partages UNC, analyse facultative des journaux ; exporte un rapport txt dans `C:\Temp\`
- Diagnostic des E/S fichiers — boucle écriture/ajout/lecture/suppression sur n'importe quel chemin ; classe les échecs en AUTH/NETWORK/TIMEOUT/DISK/PATH ; à chaque échec, capture les événements FileSystemWatcher, l'écart des autorisations NTFS par rapport à la référence, les handles de processus ouverts (Handle.exe téléchargé automatiquement depuis Sysinternals), un instantané des nouveaux processus, les tickets Kerberos et le journal de sécurité ; s'arrête après 3 échecs
- **UniFi** — rapport HTML de documentation réseau (appareils, firmware, uptime, par site) et outils de mise à niveau du firmware pour une console UniFi Controller/UniFi OS ; identifiants via `Get-Credential`, jamais codés en dur

#### Appareil

📂 Dossier : [`Device/`](scripts/Device/readme.fr.md)

- Diagnostic OpenVPN Connect — cartes PnP, services, routes, DNS, journal des événements, logiciels VPN en conflit ; exporte un rapport txt dans `C:\Temp\`

#### RDS

📂 Dossier : [`RDS/`](scripts/RDS/readme.fr.md)

- Diagnostic RDP + RD Web Access ([`Test-RDSDiagnostics.ps1`](scripts/RDS/Test-RDSDiagnostics.ps1)) — comprendre pourquoi les utilisateurs ne peuvent pas se connecter à un serveur RDP ou RDWeb :
  - Services (TermService, SessionEnv, UmRdpService), RDP activé/désactivé, NLA, limites de session, RD Licensing, règles de pare-feu, sessions actives
  - Validité et expiration du certificat HTTPS sur RDWeb ; état du pool d'applications IIS et de RD Gateway (en local uniquement)
  - Vérifications du compte utilisateur : activé, verrouillé, mot de passe expiré, appartenance à Remote Desktop Users
  - Analyse des journaux d'événements : échecs de connexion (4625), verrouillages (4740), échecs Kerberos (4771), raisons de déconnexion de session (20/40)
  - Fichier journal horodaté enregistré dans `C:\Temp\` ; `-IncludeEventLogs` pour l'analyse des événements

- Moniteur RDS en temps réel ([`Watch-RDSLive.ps1`](scripts/RDS/Watch-RDSLive.ps1)) — interroge les journaux d'événements toutes les N secondes et diffuse les nouveaux événements vers la console + un fichier journal :
  - Événements de session : connexion (21), reconnexion (22/25), fermeture de session (23), déconnexion (24), échec de connexion (20), raison de déconnexion (40) avec des codes de raison lisibles
  - Sécurité : échecs de connexion RDP (4625 type 10), verrouillages de compte (4740)
  - Licences : événements `TerminalServices-Licensing/Admin` + fournisseur `TermServLicensing` du journal Système
  - Ligne de pulsation à chaque interrogation, avec le nombre de sessions actives et de nouveaux événements
  - À exécuter directement sur chaque serveur RDS/RDWeb ; `-IntervalSeconds` (20 par défaut), `-NoLogFile` pour ne pas écrire de fichier

- Diagnostic des profils FSLogix ([`Get-FSlogix-errors.ps1`](scripts/RDS/Get-FSlogix-errors.ps1)) — rassemble la version, la configuration, les conteneurs attachés, l'état SMB/Azure Files et les événements FSLogix/disque d'un hôte de session AVD dans une seule transcription

- Réduction des disques FSLogix ([`Invoke-FSLogixShrink.ps1`](scripts/RDS/Invoke-FSLogixShrink.ps1)) — récupère l'espace que conservent les fichiers VHDX dynamiques de profil/ODFC :
  - Télécharge Invoke-FslShrinkDisk (équipe FSLogix) à un commit épinglé et vérifie son SHA-256
  - `-ReportOnly` liste chaque conteneur du partage, du plus grand au plus petit ; sinon les réduit et résume les Go récupérés et les disques non traités (en cours d'utilisation)
  - `-CheckHost` vérifie si la compaction intégrée de FSLogix à la déconnexion peut s'exécuter (version, `VHDCompactDisk`, `defragsvc`, disques dynamiques)

- Image d'hôte de session ([`Update-SessionHostImage.ps1`](scripts/RDS/Update-SessionHostImage.ps1)) — rend une image Windows 11 multisession ou un hôte AVD apte au nouveau Teams, au nouvel Outlook et à Copilot avec FSLogix, sans modifier FSLogix :
  - Vérifie WebView2, les frameworks AppX dont dépendent les applications, les builds provisionnées et les écarts par utilisateur, Teams sur AVD (SlimCore, le redirecteur WebRTC retiré le 1er octobre 2026), Shared Computer Activation et le broker de connexion
  - Met à jour Teams, Outlook, Copilot, le complément de réunion et WebView2 vers leur build la plus récente à chaque exécution (mise à jour automatique de Teams désactivée seulement si FSLogix l'exige) et corrige les applications via `Repair-AppxPackageStore.ps1` et `Update-TeamsClient.ps1` ; `-ComputerName` compare tout le pool, `-ForCapture` vérifie que Sysprep peut passer

- Watchdog des applications ([`Watch-M365Apps.ps1`](scripts/RDS/Watch-M365Apps.ps1)) — une tâche planifiée (System) qui teste le nouveau Teams, le nouvel Outlook et Copilot avec notre propre compte (`itceadmin`) sur un hôte de session :
  - Vérifie que l'hôte provisionne les applications, et pour chaque compte connecté que chacune est inscrite, intacte et démarre réellement dans cette session
  - Vérifie aussi chaque autre utilisateur connecté, et chaque application qu'un utilisateur n'a pas pu ouvrir (TWinUI 5961) ou dont l'inscription a échoué (AppX 401/404), et la réinscrit dans la session de cet utilisateur sans fenêtre - jamais pendant qu'elle tourne chez lui, jamais de réinitialisation ; chaque rapport nomme l'utilisateur
  - Répare l'hôte via `Repair-AppxPackageStore.ps1 -Provision` (avec un délai de carence) et notre propre compte en réinscrivant ou réinitialisant l'application, sans toucher aux sessions des clients ; signale `repaired` / `repair-failed` / `recovered` en JSON à un webhook n8n
  - Avant une réparation : collecte les preuves avec `Get-M365AppsLog.ps1` et signale `repairing` à n8n, en nommant chaque utilisateur ; ensuite, ouvre l'application pour un utilisateur qui avait tenté de l'ouvrir
  - Signale chaque plantage et blocage des trois applications sur l'hôte (Application Error 1000 / Application Hang 1002), y compris ceux des clients, regroupés par application, module et code d'exception

- Collecteur de journaux ([`Get-M365AppsLog.ps1`](scripts/RDS/Get-M365AppsLog.ps1)) — lecture seule, après une plainte : par utilisateur ce qui est inscrit et lancé, les événements AppX/AppReadiness/FSLogix, registre et journaux, et la tâche, l'état et les journaux du watchdog — confrontés à ce que conclurait le watchdog, avec ses angles morts (comme l'application Copilot unifiée) ; un seul zip

---

### 📊 Rapports

📂 Dossier : [`Reporting/`](scripts/Reporting/readme.fr.md)

#### Rapport de dernière connexion des ordinateurs

Rapporter la date de dernière connexion de tous les objets ordinateur d'une ou plusieurs OU et l'exporter en CSV.

- Interroge Active Directory sur les ordinateurs des OU indiquées (p. ex. `OU=Laptops`, `OU=Computers`)
- Deux modes de précision : `LastLogonTimestamp` (rapide, jusqu'à 14 jours de retard) ou `-AllDCs` (interroge chaque DC pour la valeur exacte de `LastLogon`)
- Quatre statuts : **Active** · **Active (pwd recent)** · **Stale** · **Never** · **Disabled**
- `Active (pwd recent)` : appareil marqué à tort comme obsolète à cause du délai de réplication de 14 jours — un `PasswordLastSet` de moins de 35 jours confirme que la machine est en ligne (les comptes d'ordinateur changent automatiquement de mot de passe environ tous les 30 jours)
- Colonnes du CSV : Name, Status, Enabled, LastLogon, DaysSinceLogon, PasswordLastSet, DaysSincePasswordSet, OS, IPv4, chemin de l'OU, Created, Description
- Prend en charge plusieurs OU en une seule exécution ; `-IncludeDisabled` pour inclure les objets désactivés

#### Rapport des autorisations SharePoint

**[Get-SharePointPermissionsReport.ps1](scripts/Reporting/Get-SharePointPermissionsReport.ps1)** — qui peut accéder à quel SharePoint, via quel groupe, à quel niveau. En lecture seule : chaque appel qu'il effectue est un GET.

- Part d'une vue consolidée — une ligne par personne et par site, avec le groupe par lequel passe son accès et le niveau qu'il accorde. Sinon, les autorisations et les appartenances figurent dans des rapports séparés, et « Site Owners a Full Control » plus « Site Owners contient cinq personnes » ne constitue pas encore une réponse
- En dessous : administrateurs de collection de sites, attributions de rôle sur web/liste/élément, ruptures d'héritage, groupes SharePoint avec leurs membres, groupes Entra résolus en appartenance transitive, liens de partage avec leur type, principaux externes et invités, autorisations accordées à `Everyone`
- Un élément n'est rapporté comme étendue propre que s'il a des autorisations uniques, de sorte que le rapport cartographie la structure des autorisations au lieu de répéter une ligne par fichier
- Les attributions de rôle ne sont pas lisibles via Graph et ne sont pas couvertes par les rôles Read/Write/Manage de SharePoint ; il crée donc une application éphémère avec certificat et `Sites.FullControl.All`, puis la supprime. Un secret client ne peut pas fonctionner — SharePoint Online refuse les jetons app-only basés sur un secret
- `-Excel` écrit un classeur avec une feuille par rapport plus des tableaux croisés dynamiques prêts à l'emploi ; les CSV sont toujours écrits et le classeur est construit à partir d'eux
- Reprend après une interruption à partir de la dernière liste terminée, et indique à la fin si chaque étendue a réellement pu être lue

#### Rapport de stockage SharePoint

**[Get-SharePointStorageReport.ps1](scripts/Reporting/Get-SharePointStorageReport.ps1)** — stockage à l'échelle du tenant par site, bibliothèque, historique des versions et corbeille, les chemins les plus longs comparés aux limites de SharePoint et de Windows, avec des totaux par collection de sites comparables à ceux du centre d'administration.

#### Nettoyage des versions SharePoint

**[Remove-SharePointFileVersionsByDate.ps1](scripts/Reporting/Remove-SharePointFileVersionsByDate.ps1)** — rapporter (et, avec `-Apply`, supprimer) les versions de fichiers antérieures à une date limite. La version actuelle est toujours conservée.

#### Rapport des licences

📂 Dossier : [`Reporting/Licensing/`](scripts/Reporting/Licensing/readme.fr.md)

Générateur du rapport mensuel des licences et des coûts Azure.

- Combine les données de facturation de Pax8 (CSV) et d'Ingram (Excel) en un seul fichier Excel mis en forme
- Un onglet par client avec la consommation Azure, les licences et le détail Acronis
- Onglet récapitulatif avec les totaux et la marge par client
- Lanceur PowerShell avec contrôles préalables
- Tâche planifiée Windows facultative (s'exécute le 6 de chaque mois)

---

### 🖥️ Infrastructure et appareils

#### Boîte à outils d'installation USB

📂 Dossier : [`Deployment/`](scripts/Deployment/readme.fr.md)

Boîte à outils USB pour l'installation de Windows et l'inscription Autopilot pendant l'OOBE.

- Menu interactif (Gestionnaire de périphériques, Autopilot, jonction AD, renommage de l'appareil, clé de produit, Windows Update, redémarrage)
- Navigateur d'installations client depuis le menu de la boîte à outils USB :
  - Dossier local `Install` par client (`D`)
  - Partage réseau de `INSTALL_SHARE` dans `start.local.cmd` par client (`E`)
- Pour l'option locale `D`, copiez à la fois [`Browse-InstallScripts.ps1`](scripts/Deployment/Browse-InstallScripts.ps1) et le dossier `Install` complet à côté de [`start.bat`](scripts/Deployment/start.bat)
- Avant les options `D` et `E`, la boîte à outils crée ou met à jour l'administrateur local `LocalAdmin` avec le mot de passe de `start.local.cmd` (demandé en saisie masquée s'il manque), l'ajoute à `Administrators` et définit les indicateurs de saut de l'OOBE
- S'élève automatiquement, compatible OOBE via Maj+F10
- « Tout faire » scindé : `A` = Intune (Renommer + Autopilot + Mise à jour), `C` = AD (Renommer + Jonction au domaine + Mise à jour)

#### Gestion audio

📂 Dossier : [`Device/audio/`](scripts/Device/audio/readme.fr.md)

Trois scripts qui fonctionnent ensemble pour détecter, désactiver et rétablir les microphones internes des postes — déployés via NinjaOne.

| Script | Objectif |
|---|---|
| [`detect-audiodevices.ps1`](scripts%5CDevice%5Caudio%5Cdetect-audiodevices.ps1) | Inventaire de tous les périphériques audio de l'appareil |
| [`Disable-internalmic.ps1`](scripts%5CDevice%5Caudio%5CDisable-internalmic.ps1) | Désactiver le ou les microphones internes, les casques sont ignorés |
| [`Rollback-InternalMic.ps1`](scripts%5CDevice%5Caudio%5CRollback-InternalMic.ps1) | Réactiver les microphones internes précédemment désactivés |

**Déploiement NinjaOne (les trois scripts) :**

| Paramètre | Valeur |
|---|---|
| Run as | **SYSTEM** |
| Paramètres du script | _(aucun)_ |
| Custom field requis | `AudioDeviceInventory` (device, champ texte/textarea) |
| Code de sortie | `0` = succès · `1` = erreur (script marqué en échec) |

> Le custom field `AudioDeviceInventory` doit être créé comme custom field au niveau de l'appareil dans NinjaOne avant de déployer les scripts. La sortie de chaque script y est écrite via `Ninja-Property-Set AudioDeviceInventory`.

#### Synchronisation de l'heure

📂 Dossier : [`Device/Time sync/`](scripts/Device/Time%20sync/readme.fr.md)

- Redémarrer le service Windows Time et forcer la synchronisation

#### Nettoyage de Windows

📂 Dossier : [`Device/`](scripts/Device/readme.fr.md)

Nettoyage complet de l'espace disque des postes Windows.

- Nettoie les dossiers temporaires utilisateur/système, le cache de téléchargement de Windows Update, le cache de Delivery Optimization, Prefetch, les vidages mémoire, les files d'attente WER, le cache des miniatures, le cache des shaders DirectX, la Corbeille, les caches des navigateurs (Edge, Chrome, Firefox) et les journaux d'événements
- Nettoyage du magasin de composants DISM (`/StartComponentCleanup /ResetBase`) après les mises à jour Windows
- Vidage du cache DNS
- Essai à blanc par défaut — affiche l'espace récupérable par catégorie sans rien supprimer
- Lancez avec `-Apply` pour effectuer réellement le nettoyage ; chaque catégorie peut être ignorée avec `-SkipBrowserCache`, `-SkipEventLogs`, `-SkipDism`, `-SkipRecycleBin`
- Exporte un rapport CSV des octets libérés par catégorie dans `C:\Temp\`

#### Suppression des bloatwares OEM

📂 Dossier : [`Device/`](scripts/Device/readme.fr.md)

- Détecte le fabricant de l'appareil (HP/Lenovo/Dell) et supprime les bloatwares OEM connus via `winget`, ainsi qu'une liste générique d'applications grand public du Microsoft Store (Xbox, Solitaire, Bing News/Weather, Cortana, Clipchamp)
- Essai à blanc par défaut ; `-Apply` pour supprimer réellement. Rapport CSV des applications trouvées/supprimées dans `C:\Temp\`

#### Mappage de lecteurs cloud

📂 Dossier : [`Device/DriveMapping/`](scripts/Device/DriveMapping/readme.fr.md)

- Mappe des bibliothèques de documents SharePoint/OneDrive sur des lettres de lecteur persistantes via WebDAV (`net use`), à utiliser comme script d'ouverture de session par utilisateur (application Win32 Intune ou tâche planifiée)
- Piloté par un CSV de mappages (`DriveLetter`, `Url`, `Label` facultatif) ; essai à blanc par défaut, `-Apply` pour mapper réellement
- Aucun identifiant stocké — s'appuie sur la session tenant existante de l'utilisateur connecté (comme l'accès WebDAV par le navigateur)

#### Disque temporaire et fichier d'échange (Azure / AVD)

📂 Dossier : [`Device/TempDisk/`](scripts/Device/TempDisk/readme.fr.md)

Deux scripts qui maintiennent en place le disque temporaire éphémère (`D:`) d'une VM Azure ou d'un hôte de session AVD, et y gardent le fichier d'échange.

| Script | Objectif |
|---|---|
| [`Init-TempDisk.ps1`](scripts/Device/TempDisk/Init-TempDisk.ps1) | Rétablir le disque temporaire en `D:` et y configurer le fichier d'échange |
| [`Register-InitTempDiskTask.ps1`](scripts/Device/TempDisk/Register-InitTempDiskTask.ps1) | Installer ce script sur l'appareil et l'exécuter à chaque démarrage en tant que SYSTEM |

- Le disque temporaire est effacé à chaque désallocation, redimensionnement ou changement d'hôte — et Windows lit la configuration du fichier d'échange au démarrage, donc un fichier d'échange sur une lettre de lecteur absente au démarrage n'est jamais créé et la machine pagine de nouveau sur `C:`
- Rétablit le volume (disques RAW uniquement — un disque qui porte encore des partitions est signalé, jamais formaté), déplace un lecteur optique hors de `D:` s'il gêne, puis fait pointer le fichier d'échange vers `D:\pagefile.sys` et supprime l'entrée de tous les autres lecteurs
- Windows ne lit cette configuration qu'au démarrage, donc `-RestartIfNeeded` (utilisé par la tâche de démarrage) redémarre la machine une fois lorsque c'est la seule chose qui reste - jamais après une exécution en échec, jamais tant que quelqu'un est connecté, et au plus une fois par heure. Le compte à rebours ne s'applique que si quelqu'un est connecté pour le voir ; au démarrage, le redémarrage a lieu en quelques secondes
- `-CheckOnly` produit un rapport sans rien modifier (code de sortie `2` = travail à faire) ; `-WhatIf` parcourt tout le flux ; `-Quiet` garde silencieux un démarrage sans problème

#### Imprimantes depuis JSON

📂 Dossier : [`Device/Printer/`](scripts/Device/Printer/readme.fr.md)

| Script | Doel |
|---|---|
| [`Install-Printer.ps1`](scripts/Device/Printer/Install-Printer.ps1) | Installer des pilotes d'imprimante (téléchargés depuis GitHub) et des imprimantes décrits dans un fichier JSON |

- Un seul JSON indique quels pilotes existent, d'où chacun est téléchargé — un asset de release GitHub, un dossier d'un dépôt GitHub, une URL https quelconque ou un partage — et quelles imprimantes TCP/IP les utilisent. `-Printer` en retient une partie
- Ne télécharge que ce qui manque ou est plus ancien que la `version` du JSON ; `sha256` facultatif, vérification de la signature du catalogue, puis `pnputil /add-driver /install` + `Add-PrinterDriver`, port, imprimante et paramètres d'impression par défaut (recto verso, couleur, format de papier). `"ensure": "absent"` supprime une imprimante
- Renforcé pour qu'une exécution ne puisse pas échouer à moitié : tout le JSON et chaque INF sont validés avant toute modification, une seule exécution à la fois par machine, un spouleur bloqué est redémarré et l'étape retentée, le JSON d'une URL est mis en cache comme solution de repli, et une exécution sans erreur écrit un hachage de la configuration sous `HKLM:\SOFTWARE\M365-Scripts\InstallPrinter` pour la détection
- Conçu pour le premier démarrage d'un serveur ou d'un hôte de session provisionné depuis une golden image : attend le spouleur d'impression et retente les téléchargements pendant que le réseau se met en place (`-WaitSeconds`), se relance en 64 bits, et est idempotent pour pouvoir aussi s'exécuter à chaque démarrage
- `-CheckOnly` produit un rapport sans rien modifier (code de sortie `2` = travail à faire) ; `-WhatIf` parcourt tout le flux ; `-Quiet` garde silencieux un hôte conforme

---

### 🔧 Outils maison

#### Surveillance des traitements batch SAS

📂 Dossier : [`SAS/`](scripts/SAS/readme.fr.md)

Surveiller les journaux des traitements batch SAS et l'Observateur d'événements Windows à la recherche d'erreurs, avec intégration Zabbix et alertes e-mail facultatives.

- Détecte les erreurs de spawn, les échecs d'authentification sur la bibliothèque WORK, les abandons, les erreurs disque et les lignes `ERROR:` générales
- Formats de sortie texte, JSON et Zabbix ; période d'analyse rétrospective configurable
- Script d'installation unique — installe dans `C:\Scripts\`, crée une tâche planifiée quotidienne
- Configuration UserParameter Zabbix facultative pour des alertes automatisées

#### Gestion des appareils Windows

📂 Dossier : [`Device/`](scripts/Device/readme.fr.md)

Scripts de gestion et de maintenance des appareils Windows.

**[Invoke-WindowsActivation.ps1](scripts/Device/Invoke-WindowsActivation.ps1)** — activer Windows ou gérer les paramètres de licence :
- Installer une clé de produit retail ou une clé générique KMS (`-ProductKey`)
- Configurer un serveur d'activation KMS d'entreprise (`-KmsServer`, `-KmsPort`)
- Déclencher l'activation en ligne ou via KMS (`-Activate`)
- Afficher l'état d'activation via WMI et `slmgr /dli` (`-Status`)
- Supprimer la clé de produit avant une réinstallation d'image ou un transfert de licence (`-RemoveKey`)
- Réinitialiser le compteur de la période de grâce (`-ReArm`, max. ~3-5 fois par installation)
- Demandes de confirmation par défaut ; utilisez `-Force` pour les ignorer

**Déploiement NinjaOne :**

| Paramètre | Valeur |
|---|---|
| Run as | **Administrator** |
| Custom fields | _(aucun — sortie via la console/le journal du script)_ |
| Code de sortie | `0` = succès · `1` = erreur |

> **Attention :** `-RemoveKey` et `-ReArm` demandent une confirmation interactive. Ajoutez toujours `-Force` lorsque vous les exécutez via NinjaOne, sinon le script reste bloqué.

Paramètres de script NinjaOne courants :

| Scénario | Paramètres |
|---|---|
| Vérifier l'état | `-Status` |
| Activation KMS | `-KmsServer kms.bedrijf.local -Activate -Status` |
| KMS avec un port différent | `-KmsServer kms.bedrijf.local -KmsPort 2500 -Activate` |
| Installer une clé retail + activer | `-ProductKey XXXXX-XXXXX-XXXXX-XXXXX-XXXXX -Activate -Status` |
| Supprimer la clé (avant réinstallation d'image) | `-RemoveKey -Force` |
| Réinitialiser la période de grâce | `-ReArm -Force` |

**[Invoke-WindowsCleanup.ps1](scripts/Device/Invoke-WindowsCleanup.ps1)** — analyser et, au besoin, libérer l'espace disque récupérable :
- Fichiers temporaires utilisateur + système, cache Windows Update, Delivery Optimization, Prefetch
- Vidages mémoire, files d'attente WER, cache des miniatures/shaders DirectX, cache des polices
- Corbeille, caches des navigateurs (Edge/Chrome multi-profils + Firefox)
- Journaux d'événements, magasin de composants DISM (`/StartComponentCleanup /ResetBase`)
- Journaux d'applications et système : analyse dynamique de tout C:\ à la recherche de dossiers `logs`/`log`/`logging`
- Essai à blanc par défaut ; utilisez `-Apply` pour supprimer. Résumé par catégorie avec l'espace libéré

**[Invoke-LinuxCleanup.sh](scripts/Linux/Invoke-LinuxCleanup.sh)** — la même chose pour un serveur Debian/Ubuntu, 3CX Phone System compris (bash, exécuté en root sur le serveur ; 📂 [`Linux/`](scripts/Linux/readme.fr.md)) :
- Cache APT, `autoremove` (anciens noyaux), configuration résiduelle des paquets, révisions snap désactivées
- Journal systemd, journaux ayant subi une rotation dans `/var/log`, vidages après plantage, `/tmp`, caches et corbeilles des utilisateurs, Docker en option (`--docker`)
- Avec 3CX installé : journaux 3CX et sauvegardes au-delà des N plus récentes (`--keep-backups`) ; les enregistrements sont seulement signalés, jamais supprimés
- Essai à blanc par défaut, `--apply` pour supprimer, `--check-only` pour la supervision (code de sortie `2` s'il y a du travail)

**[Repair-AppxPackageStore.ps1](scripts/Device/Repair-AppxPackageStore.ps1)** — réparer les paquets AppX (Teams, nouvel Outlook, ou tout autre) qui échouent avec `0x80070490` / "Deployment Register operation ... from:  (AppxManifest.xml)" :
- Diagnostique les inscriptions dont les fichiers ont disparu, les copies provisionnées sans fichiers et les entrées orphelines de `AppxAllUserStore` (pas de profil, pas de fichiers, pas de manifeste)
- Sur les hôtes FSLogix, lit les erreurs `Microsoft-FSLogix-Apps` : la version exacte demandée par les profils face à ce que cet hôte provisionne, le build FSLogix, `InstallAppxPackages`, ODFC `IncludeTeams` et les stratégies d'installation AppX
- Liste **chaque** application dont l'installation, la mise à jour ou l'inscription a échoué au cours des `-Days` derniers jours (journal de déploiement AppX + journal FSLogix), avec la signification de chaque code d'erreur
- Répare dans un ordre fixe — déprovisionner, réinscrire, supprimer, puis sauvegarder chaque clé de registre en `.reg` avant de la supprimer — et relit tout ; `-Provision` réinstalle Teams / le nouvel Outlook pour tous les utilisateurs avec le programme d'installation de Microsoft lui-même, ou via winget avec `-UseWinget` ; `-WingetId` fait de même pour toute autre application
- `-CheckOnly` ne modifie rien ; avec `-Name '*'`, les paquets système/framework et les marqueurs Deprovisioned ne sont jamais touchés

#### Gestion DNS

📂 Dossier : [`DNS/`](scripts/DNS/readme.fr.md)

Scripts de gestion des enregistrements DNS dans des zones DNS intégrées à Active Directory.

- Résoudre des enregistrements DNS publics via Google DNS (dig) et les importer comme enregistrements A ou CNAME dans le DNS AD
- Essai à blanc par défaut — montre ce qui serait créé avant de l'appliquer
- Idempotent — ignore les enregistrements qui existent déjà

---

### ☁️ Infrastructure Azure

📂 Dossier : [`Azure/`](scripts/Azure/readme.fr.md)

Scripts qui ciblent directement Azure IaaS via le module `Az` — pas le tenant M365, et non intégrés à [`menu.ps1`](menu.ps1).

- **[Azure-NVMe-Conversion.ps1](scripts/Azure/VM/Azure-NVMe-Conversion.ps1)** — script tiers embarqué (Microsoft, licence MIT, issu de `Azure/SAP-on-Azure-Scripts-and-Utilities`) qui convertit le type de contrôleur de disque d'une VM entre SCSI et NVMe, avec vérifications et corrections de l'état de préparation des pilotes dans l'invité, pour les invités Windows comme Linux
- **[Search-AADDSUserActivity.ps1](scripts/Azure/Search-AADDSUserActivity.ps1)** — interroge toutes les tables d'audit d'Azure AD Domain Services dans Log Analytics pour un seul utilisateur en une seule requête `union`, au lieu de deviner dans quelle table un événement a atterri

---

### 🗄️ Réécritures des anciennes boîtes à outils

Un dépôt PowerShell interne aujourd'hui retiré (et deux boîtes à outils GitHub tierces forkées qu'il embarquait) a été passé en revue script par script et modernisé selon les conventions de ce dépôt — Graph/Exchange Online au lieu des modules retirés `MSOnline`/`AzureAD`, essai à blanc par défaut avec `-Apply` pour tout ce qui modifie quelque chose, aucune donnée client ni aucun secret codé en dur. Aucun de ces scripts n'est intégré à [`menu.ps1`](menu.ps1) — ce sont des scripts d'audit, de rapport et de mise en place destinés à être exécutés directement, sur le même modèle que [`scripts/RDS/`](scripts/RDS/readme.fr.md), [`scripts/Azure/`](scripts/Azure/readme.fr.md) et [`scripts/Network/UniFi/`](scripts/Network/UniFi/readme.fr.md). Chaque dossier a son propre readme avec la documentation complète des paramètres et de l'utilisation.

| Dossier | Source | Couvre |
|--------|--------|--------|
| [`TenantOnboarding/`](scripts/TenantOnboarding/readme.fr.md) | Boîte à outils interne de mise en place de tenants | Provisionnement de nouveaux tenants (administrateur break-glass, groupes de référence/attribution Intune), rapports multi-tenant/GDAP des licences + mots de passe break-glass, déploiement d'applications Win32/Chocolatey, configuration des appareils (alimentation kiosque, désinstallation d'Office, disposition du menu Démarrer), gestion de OneDrive, gestion des utilisateurs par DG dynamiques/groupes de fonctionnalités |
| [`Office365Toolkit/`](scripts/Office365Toolkit/readme.fr.md) | Fork de [`directorcia/Office365`](https://github.com/directorcia/Office365) (CIAOPS) | Rapports Secure Score, nettoyage des consentements d'applications d'entreprise, verrouillage de la connexion aux boîtes aux lettres partagées, référentiel EOP, audits d'hygiène des boîtes aux lettres/de risque de transfert, recherche dans le Unified Audit Log, inventaire des stratégies Intune |
| [`PatronToolkit/`](scripts/PatronToolkit/readme.fr.md) | Fork de [`directorcia/patron`](https://github.com/directorcia/patron) | Inscription MFA + export des stratégies CA, audits des consentements d'applications d'entreprise + règles de boîte de réception suspectes + alertes de sécurité unifiées, posture consolidée de sécurité de la messagerie + contrôles de l'audit des boîtes aux lettres, validation SPF/DMARC, rapports d'attribution des stratégies Intune + appareils Autopilot, message trace, configuration du partage SharePoint, rapport de configuration Teams |
| [`LegacyUtilities/`](scripts/LegacyUtilities/readme.fr.md) | Boîte à outils interne (petits scripts divers) | Droits sur les dossiers de boîtes aux lettres/accès délégué, création en masse de boîtes aux lettres partagées/contacts, synchronisation des contacts, nettoyage des éléments de courrier en double, appartenance aux groupes M365, sauvegarde des stratégies CA, clonage Teams/Planner, mappage de lecteurs Azure Files, réglages d'appareil NumLock/verrouillage du poste, provisionnement d'environnements Workspace 365 |

Les deux forks GitHub ont été passés en revue fonctionnalité par fonctionnalité plutôt que portés 1:1 — des scripts de rapport quasi identiques à usage unique ont été regroupés en un nombre réduit de scripts bien paramétrés, et les fonctionnalités déjà couvertes ailleurs dans ce dépôt ont été ignorées plutôt que dupliquées (voir le readme de chaque dossier pour la liste complète des éléments ignorés et leur justification). Tout le code est une nouvelle implémentation dans le style de ce dépôt, et non une copie des projets sources.

---

## Structure du dépôt

Chaque dossier a son propre [`readme.md`](readme.md) — cette arborescence est une carte ; suivez les liens pour la documentation complète des paramètres et de l'utilisation.

<pre>
<a href="readme.fr.md">M365-Scripts/</a>
├── <a href=".github/workflows/ci.yml">.github/workflows/ci.yml</a>         ← CI : syntaxe, analyseur, liens, docs générées, ShellCheck
├── <a href=".gitignore">.gitignore</a>
├── <a href=".vscode">.vscode/</a>
│   └── <a href=".vscode/settings.json">settings.json</a>
├── <a href="load.ps1">load.ps1</a>                         ← Point d'entrée : configuration au premier lancement + lance le menu
├── <a href="menu.ps1">menu.ps1</a>                         ← Lanceur interactif (tous les scripts + fonctions M365)
├── <a href="readme.fr.md">readme.md</a>
└── <a href="scripts/readme.fr.md">scripts/</a>
    ├── <a href="scripts/readme.fr.md">readme.md</a>                    ← Index de toutes les catégories ci-dessous
    ├── <a href="scripts/INDEX.md">INDEX.md</a>                     ← Tous les scripts de A à Z avec leur dossier (généré)
    ├── <a href="scripts/Azure/readme.fr.md">Azure/</a>                        ← cible directement Azure IaaS via Az, pas le tenant M365
    │   ├── <a href="scripts/Azure/readme.fr.md">readme.md</a>
    │   └── <a href="scripts/Azure/VM/readme.fr.md">VM/</a>
    │       ├── <a href="scripts/Azure/VM/readme.fr.md">readme.md</a>
    │       └── <a href="scripts/Azure/VM/Azure-NVMe-Conversion.ps1">Azure-NVMe-Conversion.ps1</a>   ← embarqué (Microsoft, MIT) — conversion du contrôleur de disque SCSI/NVMe
    ├── <a href="scripts/Entra/readme.fr.md">Entra/</a>
    │   ├── <a href="scripts/Entra/readme.fr.md">readme.md</a>
    │   ├── <a href="scripts/Entra/Set-UserManager.ps1">Set-UserManager.ps1</a>
    │   ├── <a href="scripts/Entra/Remove-M365Users.ps1">Remove-M365Users.ps1</a>
    │   ├── <a href="scripts/Entra/New-M365User.ps1">New-M365User.ps1</a>
    │   ├── <a href="scripts/Entra/Import-M365Users.ps1">Import-M365Users.ps1</a>
    │   ├── <a href="scripts/Entra/Get-M365UserLicenses.ps1">Get-M365UserLicenses.ps1</a>
    │   ├── <a href="scripts/Entra/Import-ConditionalAccessBaseline.ps1">Import-ConditionalAccessBaseline.ps1</a>
    │   └── <a href="scripts/Entra/Test-M365GroupMembership.ps1">Test-M365GroupMembership.ps1</a>   ← auditer les propriétaires et membres des groupes M365 / Teams
    ├── <a href="scripts/Exchange/readme.fr.md">Exchange/</a>
    │   ├── <a href="scripts/Exchange/readme.fr.md">readme.md</a>
    │   ├── <a href="scripts/Exchange/Migrate-Calendar.ps1">Migrate-Calendar.ps1</a>
    │   ├── <a href="scripts/Exchange/Convert-SharedCalendarToResource.ps1">Convert-SharedCalendarToResource.ps1</a>  ← calendrier partagé dans la boîte aux lettres d'un utilisateur → sa propre boîte aux lettres de salle/d'équipement
    │   ├── <a href="scripts/Exchange/Move-SharedCalendar.ps1">Move-SharedCalendar.ps1</a>  ← tout-en-un : recherche par mot-clé + conversion + qui doit basculer
    │   ├── <a href="scripts/Exchange/Set-Calendar-rights.ps1">Set-Calendar-rights.ps1</a>
    │   ├── <a href="scripts/Exchange/Set-Distributionlist-dynamic-static.ps1">Set-Distributionlist-dynamic-static.ps1</a>
    │   ├── <a href="scripts/Exchange/Move-InboxToArchive.ps1">Move-InboxToArchive.ps1</a>
    │   ├── <a href="scripts/Exchange/Restore-MailboxMessages.ps1">Restore-MailboxMessages.ps1</a>  ← remettre en place le courrier déplacé/supprimé à une date, et qui l'a fait
    │   ├── <a href="scripts/Exchange/Test-CalendarPermissions.ps1">Test-CalendarPermissions.ps1</a>
    │   ├── <a href="scripts/Exchange/Get-CalendarMappings.ps1">Get-CalendarMappings.ps1</a>  ← où chaque calendrier est ajouté dans Outlook, à côté des droits correspondants
    │   ├── <a href="scripts/Exchange/Test-MailboxPermissions.ps1">Test-MailboxPermissions.ps1</a>
    │   ├── <a href="scripts/Exchange/Test-DistributionGroupPermissions.ps1">Test-DistributionGroupPermissions.ps1</a>
    │   ├── <a href="scripts/Exchange/Test-DkimConfig.ps1">Test-DkimConfig.ps1</a>
    │   ├── <a href="scripts/Exchange/Get-ExternalForwards.ps1">Get-ExternalForwards.ps1</a>
    │   ├── <a href="scripts/Exchange/Get-MailboxSizes.ps1">Get-MailboxSizes.ps1</a>
    │   └── <a href="scripts/Exchange/Get-DistributionGroupMembers.ps1">Get-DistributionGroupMembers.ps1</a>  ← qui figure dans quelle liste de distribution, en classeur Excel pour le client
    ├── <a href="scripts/Graph/readme.fr.md">Graph/</a>
    │   ├── <a href="scripts/Graph/readme.fr.md">readme.md</a>
    │   └── <a href="scripts/Graph/logic-permissies.ps1">logic-permissies.ps1</a>     ← accorder un rôle d'application Graph à l'identité managée d'une Logic App
    ├── <a href="scripts/Intune/readme.fr.md">Intune/</a>
    │   ├── <a href="scripts/Intune/readme.fr.md">readme.md</a>
    │   ├── <a href="scripts/Intune/Compare-IntuneConfig.ps1">Compare-IntuneConfig.ps1</a>  ← dérive de la config Intune par rapport à une sauvegarde du référentiel MSP (IntuneBackupAndRestore)
    │   ├── <a href="scripts/Intune/Get-Autopilot/readme.fr.md">Get-Autopilot/</a>
    │   │   ├── <a href="scripts/Intune/Get-Autopilot/readme.fr.md">readme.md</a>
    │   │   ├── <a href="scripts/Intune/Get-Autopilot/Get-WindowsAutoPilotInfo.ps1">Get-WindowsAutoPilotInfo.ps1</a>
    │   │   └── <a href="scripts/Intune/Get-Autopilot/GetAutoPilot.CMD">GetAutoPilot.CMD</a>
    │   ├── <a href="scripts/Intune/iOS-Compliance-Updater/readme.fr.md">iOS-Compliance-Updater/</a>
    │   │   ├── <a href="scripts/Intune/iOS-Compliance-Updater/readme.fr.md">readme.md</a>
    │   │   ├── <a href="scripts/Intune/iOS-Compliance-Updater/Update-iOSCompliancePolicy.ps1">Update-iOSCompliancePolicy.ps1</a>   ← script principal (manuel ou tâche planifiée)
    │   │   ├── <a href="scripts/Intune/iOS-Compliance-Updater/Setup.ps1">Setup.ps1</a>                        ← une seule fois : App Registration + config.json
    │   │   ├── <a href="scripts/Intune/iOS-Compliance-Updater/Install-ScheduledTask.ps1">Install-ScheduledTask.ps1</a>        ← enregistrer la tâche planifiée hebdomadaire
    │   │   └── <a href="scripts/Intune/iOS-Compliance-Updater/config.example.json">config.example.json</a>
    │   └── <a href="scripts/Intune/Desktop/readme.fr.md">Desktop/</a>                  ← fond d'écran/écran de verrouillage de l'entreprise (le thème Office se trouve dans Custom Scripts/, voir plus bas)
    │       ├── <a href="scripts/Intune/Desktop/readme.fr.md">readme.md</a>
    │       ├── <a href="scripts/Intune/Desktop/Add%20Lockscreen%20to%20start%20and%20desktop/readme.fr.md">Add Lockscreen to start and desktop/</a>
    │       │   ├── <a href="scripts/Intune/Desktop/Add%20Lockscreen%20to%20start%20and%20desktop/readme.fr.md">readme.md</a>
    │       │   ├── <a href="scripts/Intune/Desktop/Add%20Lockscreen%20to%20start%20and%20desktop/add-lock.ps1">add-lock.ps1</a>               ← raccourci « Lock Workstation » dans la barre des tâches
    │       │   └── <a href="scripts/Intune/Desktop/Add%20Lockscreen%20to%20start%20and%20desktop/add-shortcut-lock.ps1">add-shortcut-lock.ps1</a>
    │       └── <a href="scripts/Intune/Desktop/Background/readme.fr.md">Background/</a>
    │           ├── <a href="scripts/Intune/Desktop/Background/readme.fr.md">readme.md</a>
    │           ├── <a href="scripts/Intune/Desktop/Background/Desktop/readme.fr.md">Desktop/</a>
    │           │   ├── <a href="scripts/Intune/Desktop/Background/Desktop/readme.fr.md">readme.md</a>
    │           │   ├── <a href="scripts/Intune/Desktop/Background/Desktop/Set-CorporateWallpaper.ps1">Set-CorporateWallpaper.ps1</a>  ← fond d'écran de l'entreprise via Intune (contrôle de hachage, PersonalizationCSP)
    │           │   └── <a href="scripts/Intune/Desktop/Background/Desktop/Remove-CorporateWallpaper.ps1">Remove-CorporateWallpaper.ps1</a>
    │           └── <a href="scripts/Intune/Desktop/Background/Lockscreen/readme.fr.md">Lockscreen/</a>
    │               ├── <a href="scripts/Intune/Desktop/Background/Lockscreen/readme.fr.md">readme.md</a>
    │               └── <a href="scripts/Intune/Desktop/Background/Lockscreen/Make-lockscreen.ps1">Make-lockscreen.ps1</a>         ← écran de verrouillage de l'entreprise via Intune (téléchargement validé, PersonalizationCSP)
    ├── <a href="scripts/Device/readme.fr.md">Device/</a>
    │   ├── <a href="scripts/Device/readme.fr.md">readme.md</a>
    │   ├── <a href="scripts/Device/Invoke-WindowsActivation.ps1">Invoke-WindowsActivation.ps1</a> ← activer Windows, définir la clé de produit / le serveur KMS
    │   ├── <a href="scripts/Device/Invoke-WindowsCleanup.ps1">Invoke-WindowsCleanup.ps1</a>    ← temp, cache, WU, DISM, navigateur, journaux d'événements
    │   ├── <a href="scripts/Device/Clear-TempFiles.ps1">Clear-TempFiles.ps1</a>
    │   ├── <a href="scripts/Device/Remove-OemBloatware.ps1">Remove-OemBloatware.ps1</a>      ← suppression des bloatwares HP/Lenovo/Dell + génériques du Store
    │   ├── <a href="scripts/Device/Repair-AppxPackageStore.ps1">Repair-AppxPackageStore.ps1</a>  ← réparer AppX 0x80070490 (entrées orphelines du magasin, rejeu FSLogix)
    │   ├── <a href="scripts/Device/Test-OpenVpnDiagnostics.ps1">Test-OpenVpnDiagnostics.ps1</a>  ← diagnostic OpenVPN Connect
    │   ├── <a href="scripts/Device/Update-TeamsClient.ps1">Update-TeamsClient.ps1</a>       ← mettre à jour le nouveau Teams + le complément de réunion quand un build plus récent existe
    │   ├── <a href="scripts/Device/Update-TeamsClient.md">Update-TeamsClient.md</a>        ← comment ce script décide, étape par étape
    │   ├── <a href="scripts/Device/Update-TeamsClient-ITGlue.md">Update-TeamsClient-ITGlue.md</a> ← version service desk (NL) à coller dans IT Glue
    │   ├── <a href="scripts/Device/audio/readme.fr.md">audio/</a>
    │   │   ├── <a href="scripts/Device/audio/readme.fr.md">readme.md</a>
    │   │   ├── <a href="scripts/Device/audio/detect-audiodevices.ps1">detect-audiodevices.ps1</a>
    │   │   ├── <a href="scripts/Device/audio/Disable-internalmic.ps1">Disable-internalmic.ps1</a>
    │   │   └── <a href="scripts/Device/audio/Rollback-InternalMic.ps1">Rollback-InternalMic.ps1</a>
    │   ├── <a href="scripts/Device/DriveMapping/readme.fr.md">DriveMapping/</a>
    │   │   ├── <a href="scripts/Device/DriveMapping/readme.fr.md">readme.md</a>
    │   │   └── <a href="scripts/Device/DriveMapping/New-CloudDriveMapping.ps1">New-CloudDriveMapping.ps1</a>   ← mapper des bibliothèques SharePoint/OneDrive sur des lettres de lecteur (WebDAV)
    │   ├── <a href="scripts/Device/Printer/readme.fr.md">Printer/</a>
    │   │   ├── <a href="scripts/Device/Printer/readme.fr.md">readme.md</a>
    │   │   ├── <a href="scripts/Device/Printer/Install-Printer.ps1">Install-Printer.ps1</a>            ← pilotes d'imprimante (depuis GitHub) + imprimantes depuis un fichier JSON
    │   │   └── <a href="scripts/Device/Printer/printers.example.json">printers.example.json</a>          ← chaque champ et chaque source de pilote
    │   ├── <a href="scripts/Device/TempDisk/readme.fr.md">TempDisk/</a>
    │   │   ├── <a href="scripts/Device/TempDisk/readme.fr.md">readme.md</a>
    │   │   ├── <a href="scripts/Device/TempDisk/Init-TempDisk.ps1">Init-TempDisk.ps1</a>              ← rétablir le disque temporaire éphémère en D: et y placer le fichier d'échange
    │   │   └── <a href="scripts/Device/TempDisk/Register-InitTempDiskTask.ps1">Register-InitTempDiskTask.ps1</a>  ← installer ce script et l'exécuter à chaque démarrage en tant que SYSTEM
    │   └── <a href="scripts/Device/Time%20sync/readme.fr.md">Time sync/</a>
    │       ├── <a href="scripts/Device/Time%20sync/readme.fr.md">readme.md</a>
    │       └── <a href="scripts/Device/Time%20sync/Restart-Time-Sync.ps1">Restart-Time-Sync.ps1</a>
    ├── <a href="scripts/Linux/readme.fr.md">Linux/</a>
    │   ├── <a href="scripts/Linux/readme.fr.md">readme.md</a>
    │   └── <a href="scripts/Linux/Invoke-LinuxCleanup.sh">Invoke-LinuxCleanup.sh</a>        ← bash : nettoyage du disque pour Debian/Ubuntu, 3CX compris
    ├── <a href="scripts/Network/readme.fr.md">Network/</a>
    │   ├── <a href="scripts/Network/readme.fr.md">readme.md</a>
    │   ├── <a href="scripts/Network/Test-Ports.ps1">Test-Ports.ps1</a>
    │   ├── <a href="scripts/Network/Test-AuthNetworkDiagnostics.ps1">Test-AuthNetworkDiagnostics.ps1</a>   ← diagnostic des problèmes d'authentification/réseau
    │   ├── <a href="scripts/Network/Test-FileIODiagnostics.ps1">Test-FileIODiagnostics.ps1</a>        ← test d'E/S fichiers + surveillance de répertoire en temps réel
    │   └── <a href="scripts/Network/UniFi/readme.fr.md">UniFi/</a>
    │       ├── <a href="scripts/Network/UniFi/readme.fr.md">readme.md</a>
    │       ├── <a href="scripts/Network/UniFi/UnifiApi.ps1">UnifiApi.ps1</a>                  ← fonction d'aide partagée de connexion/session (chargée en dot-source)
    │       ├── <a href="scripts/Network/UniFi/Get-UnifiNetworkReport.ps1">Get-UnifiNetworkReport.ps1</a>     ← rapport HTML de documentation réseau
    │       └── <a href="scripts/Network/UniFi/Update-UnifiFirmware.ps1">Update-UnifiFirmware.ps1</a>       ← lister/déclencher les mises à niveau du firmware sur plusieurs sites
    ├── <a href="scripts/RDS/readme.fr.md">RDS/</a>
    │   ├── <a href="scripts/RDS/readme.fr.md">readme.md</a>
    │   ├── <a href="scripts/RDS/Get-FSlogix-errors.ps1">Get-FSlogix-errors.ps1</a>            ← diagnostic des profils FSLogix / Azure Files
    │   ├── <a href="scripts/RDS/Get-M365AppsLog.ps1">Get-M365AppsLog.ps1</a>               ← collecteur : applis par utilisateur vs. le watchdog
    │   ├── <a href="scripts/RDS/Invoke-FSLogixShrink.ps1">Invoke-FSLogixShrink.ps1</a>          ← réduire les disques FSLogix, vérifier la compaction
    │   ├── <a href="scripts/RDS/Test-RDSDiagnostics.ps1">Test-RDSDiagnostics.ps1</a>           ← diagnostic des échecs de connexion RDP/RDWeb
    │   ├── <a href="scripts/RDS/Update-SessionHostImage.ps1">Update-SessionHostImage.ps1</a>       ← Teams / Outlook / Copilot compatibles FSLogix sur l'image
    │   ├── <a href="scripts/RDS/Watch-M365Apps.ps1">Watch-M365Apps.ps1</a>                ← watchdog : Teams / Outlook / Copilot, réparation, n8n
    │   ├── <a href="scripts/RDS/Watch-M365Apps-ITGlue.md">Watch-M365Apps-ITGlue.md</a>    ← version service desk (NL) à coller dans IT Glue
    │   └── <a href="scripts/RDS/Watch-RDSLive.ps1">Watch-RDSLive.ps1</a>                 ← surveillance en temps réel des sessions + licences
    ├── <a href="scripts/SMTP/readme.fr.md">SMTP/</a>
    │   ├── <a href="scripts/SMTP/readme.fr.md">readme.md</a>
    │   ├── <a href="scripts/SMTP/testsmtp.ps1">testsmtp.ps1</a>
    │   └── <a href="scripts/SMTP/testsmtp_5min.ps1">testsmtp_5min.ps1</a>
    ├── <a href="scripts/Deployment/readme.fr.md">Deployment/</a>                   ← boîte à outils d'installation USB (OOBE / Autopilot)
    │   ├── <a href="scripts/Deployment/readme.fr.md">readme.md</a>
    │   ├── <a href="scripts/Deployment/start.bat">start.bat</a>
    │   ├── <a href="scripts/Deployment/autorun.inf">autorun.inf</a>
    │   └── <a href="scripts/Deployment/Browse-InstallScripts.ps1">Browse-InstallScripts.ps1</a>
    ├── <a href="scripts/DNS/readme.fr.md">DNS/</a>
    │   ├── <a href="scripts/DNS/readme.fr.md">readme.md</a>
    │   ├── <a href="scripts/DNS/Import-DnsRecords.ps1">Import-DnsRecords.ps1</a>   ← résolution via Google DNS + import dans le DNS AD
    │   └── <a href="scripts/DNS/example-records.csv">example-records.csv</a>
    ├── <a href="scripts/SAS/readme.fr.md">SAS/</a>
    │   ├── <a href="scripts/SAS/readme.fr.md">readme.md</a>
    │   ├── <a href="scripts/SAS/rca.md">rca.md</a>
    │   ├── <a href="scripts/SAS/Monitor-SASBatchErrors.ps1">Monitor-SASBatchErrors.ps1</a>   ← analyse des journaux + de l'Observateur d'événements à la recherche d'erreurs SAS
    │   ├── <a href="scripts/SAS/Setup-SASMonitoring.ps1">Setup-SASMonitoring.ps1</a>      ← script d'installation, tâche planifiée, config Zabbix
    │   ├── <a href="scripts/SAS/Test-SASWorkDirectory.ps1">Test-SASWorkDirectory.ps1</a>    ← valider l'état du répertoire WORK
    │   └── <a href="scripts/SAS/zabbix_sas_monitor.conf">zabbix_sas_monitor.conf</a>
    ├── <a href="scripts/SharePoint/readme.fr.md">SharePoint/</a>
    │   ├── <a href="scripts/SharePoint/readme.fr.md">readme.md</a>
    │   ├── <a href="scripts/SharePoint/Find-SiteContent.ps1">Find-SiteContent.ps1</a>         ← rechercher dans tout un site (nom/chemin/type/date ou texte intégral) + rapporter les autorisations de chaque résultat (PnP)
    │   ├── <a href="scripts/SharePoint/Search-SharePointContent.ps1">Search-SharePointContent.ps1</a> ← idem, à l'échelle du tenant via Graph app-only : delta + /permissions, liens de partage et invités (fichiers/dossiers)
    │   ├── <a href="scripts/SharePoint/Restore-RecycleBinItems.ps1">Restore-RecycleBinItems.ps1</a>  ← restaurer des fichiers supprimés depuis une corbeille : un site/OneDrive ou tout le tenant (PnP, inscription d'application automatique)
    │   ├── <a href="scripts/SharePoint/Trace-SharePointFile.ps1">Trace-SharePointFile.ps1</a>     ← où est passé un fichier : renommages, déplacements, copies, suppressions (y compris via un dossier) depuis le journal d'audit, en heure de Bruxelles
    │   ├── <a href="scripts/SharePoint/Revoke-SharePointUserAccess.ps1">Revoke-SharePointUserAccess.ps1</a> ← retirer l'accès d'un utilisateur à tous les niveaux, liens de partage compris (rapport seul sans -Apply)
    │   ├── <a href="scripts/SharePoint/Test-SharePointAccessScripts.ps1">Test-SharePointAccessScripts.ps1</a> ← vérifier les deux scripts d'accès sans tenant (bloc d'authentification partagé + entonnoir de révocation)
    │   └── <a href="scripts/SharePoint/Provisioning/readme.fr.md">Provisioning/</a>                ← provisionner toute une structure à partir d'une seule config JSON (PnP + Graph)
    │       ├── <a href="scripts/SharePoint/Provisioning/readme.fr.md">readme.md</a>
    │       ├── <a href="scripts/SharePoint/Provisioning/SharePoint-Handleiding.md">SharePoint-Handleiding.md</a>      ← guide utilisateur (NL) à remettre au client (gabarit)
    │       ├── <a href="scripts/SharePoint/Provisioning/example.config.json">example.config.json</a>          ← modèle d'exemple (CHANGEME) — les configurations client à côté sont ignorées par git
    │       ├── <a href="scripts/SharePoint/Provisioning/Install-SharePointStructure.ps1">Install-SharePointStructure.ps1</a> ← tout construire en une exécution, y compris l'inscription d'application (temporaire)
    │       ├── <a href="scripts/SharePoint/Provisioning/SharePointStructure.Common.ps1">SharePointStructure.Common.ps1</a> ← fonctions d'aide partagées (chargées en dot-source par chaque script d'ici)
    │       ├── <a href="scripts/SharePoint/Provisioning/New-SharePointMetadata.ps1">New-SharePointMetadata.ps1</a>   ← ensemble de termes, colonnes de site, types de contenu (chaque site de la config)
    │       ├── <a href="scripts/SharePoint/Provisioning/Set-SharePointLibraries.ps1">Set-SharePointLibraries.ps1</a>  ← bibliothèques/dossiers de canaux, types de contenu, valeurs par défaut, affichages, droits des groupes
    │       ├── <a href="scripts/SharePoint/Provisioning/Update-SharePointShareStatus.ps1">Update-SharePointShareStatus.ps1</a> ← déduire la colonne Deelstatus, signaler le partage excessif (exit 2)
    │       └── <a href="scripts/SharePoint/Provisioning/Test-SharePointStructure.ps1">Test-SharePointStructure.ps1</a> ← contrôle de dérive en lecture seule par rapport à la config (exit 2)
    ├── <a href="scripts/Teams/readme.fr.md">Teams/</a>
    │   ├── <a href="scripts/Teams/readme.fr.md">readme.md</a>
    │   └── <a href="scripts/Teams/Invoke-TeamsArchive.ps1">Invoke-TeamsArchive.ps1</a>  ← export + archivage Teams/SharePoint (Graph, PS7+, Global Admin)
    ├── <a href="scripts/Reporting/readme.fr.md">Reporting/</a>
    │   ├── <a href="scripts/Reporting/readme.fr.md">readme.md</a>
    │   ├── <a href="scripts/Reporting/Get-ComputerLastLogon.ps1">Get-ComputerLastLogon.ps1</a>        ← dernière connexion par ordinateur dans une ou plusieurs OU, export CSV
    │   ├── <a href="scripts/Reporting/Get-SharePointStorageReport.ps1">Get-SharePointStorageReport.ps1</a>  ← rapport de stockage SharePoint à l'échelle du tenant
    │   ├── <a href="scripts/Reporting/Get-SharePointPermissionsReport.ps1">Get-SharePointPermissionsReport.ps1</a> ← qui a accès à quoi et via quel groupe, en CSV + Excel
    │   ├── <a href="scripts/Reporting/Remove-SharePointFileVersionsByDate.ps1">Remove-SharePointFileVersionsByDate.ps1</a> ← supprimer les versions de fichiers antérieures à une date
    │   └── <a href="scripts/Reporting/Licensing/readme.fr.md">Licensing/</a>
    │       ├── <a href="scripts/Reporting/Licensing/readme.fr.md">readme.md</a>
    │       ├── <a href="scripts/Reporting/Licensing/genereer_licentie_overzicht.py">genereer_licentie_overzicht.py</a>
    │       ├── <a href="scripts/Reporting/Licensing/genereer_rapport.ps1">genereer_rapport.ps1</a>
    │       ├── <a href="scripts/Reporting/Licensing/genereer_rapport.bat">genereer_rapport.bat</a>
    │       └── <a href="scripts/Reporting/Licensing/create_scheduled_task.ps1">create_scheduled_task.ps1</a>
    ├── <a href="scripts/Startup/readme.fr.md">Startup/</a>
    │   ├── <a href="scripts/Startup/readme.fr.md">readme.md</a>
    │   ├── <a href="scripts/Startup/functies.ps1">functies.ps1</a>             ← bibliothèque de fonctions M365 (chargée en dot-source par le menu)
    │   ├── <a href="scripts/Startup/RequiredModules.psd1">RequiredModules.psd1</a>     ← La liste unique des modules requis
    │   ├── <a href="scripts/Startup/Test-RequiredModules.ps1">Test-RequiredModules.ps1</a> ← Signale les modules chargés par des scripts mais absents de la liste
    │   ├── <a href="scripts/Startup/Install-Modules.ps1">Install-Modules.ps1</a>      ← Amorçage : installer et importer tous les modules
    │   ├── <a href="scripts/Startup/Update-Modules.ps1">Update-Modules.ps1</a>       ← Vérifier/mettre à jour les modules requis (load.ps1 l'exécute au démarrage), puis le reste
    │   ├── <a href="scripts/Startup/Test-PowerShellSyntax.ps1">Test-PowerShellSyntax.ps1</a>
    │   ├── <a href="scripts/Startup/Update-ScriptIndex.ps1">Update-ScriptIndex.ps1</a>   ← Régénère scripts/INDEX.md à partir des en-têtes .SYNOPSIS
    │   ├── <a href="scripts/Startup/Test-MarkdownLinks.ps1">Test-MarkdownLinks.ps1</a>   ← Vérifie chaque lien des readmes : fichiers et ancres internes à la page
    │   └── <a href="scripts/Startup/Convert-MarkdownToHtml.ps1">Convert-MarkdownToHtml.ps1</a> ← Document markdown → une page HTML autonome mise en forme
    ├── <a href="scripts/Custom%20Scripts/readme.fr.md">Custom Scripts/</a>                 ← déploiement des thèmes/couleurs Office, URL du thème en paramètre
    │   ├── <a href="scripts/Custom%20Scripts/readme.fr.md">readme.md</a>
    │   └── <a href="scripts/Custom%20Scripts/Intune/readme.fr.md">Intune/</a>
    │       ├── <a href="scripts/Custom%20Scripts/Intune/readme.fr.md">readme.md</a>
    │       └── <a href="scripts/Custom%20Scripts/Intune/Desktop/readme.fr.md">Desktop/</a>
    │           ├── <a href="scripts/Custom%20Scripts/Intune/Desktop/readme.fr.md">readme.md</a>
    │           ├── <a href="scripts/Custom%20Scripts/Intune/Desktop/Deploy-OfficeTheme.ps1">Deploy-OfficeTheme.ps1</a>        ← installe un thème Office (.thmx) depuis une URL
    │           └── <a href="scripts/Custom%20Scripts/Intune/Desktop/Office%20Themes/readme.fr.md">Office Themes/</a>
    │               ├── <a href="scripts/Custom%20Scripts/Intune/Desktop/Office%20Themes/readme.fr.md">readme.md</a>
    │               └── <a href="scripts/Custom%20Scripts/Intune/Desktop/Office%20Themes/Deploy-Officecolors.ps1">Deploy-Officecolors.ps1</a>   ← installe uniquement un jeu de couleurs depuis une URL
    ├── <a href="scripts/TenantOnboarding/readme.fr.md">TenantOnboarding/</a>                ← modernisé à partir d'une boîte à outils interne de mise en place de tenants retirée, hors menu
    │   ├── <a href="scripts/TenantOnboarding/readme.fr.md">readme.md</a>
    │   ├── <a href="scripts/TenantOnboarding/Provisioning/readme.fr.md">Provisioning/</a>         (3 scripts)  ← administrateur break-glass, groupes de référence, attribution des stratégies Intune
    │   ├── <a href="scripts/TenantOnboarding/MultiTenant/readme.fr.md">MultiTenant/</a>          (3 scripts)  ← rapport de licences GDAP, rotation des mots de passe break-glass, index du portail client
    │   ├── <a href="scripts/TenantOnboarding/AppDeployment/readme.fr.md">AppDeployment/</a>        (7 scripts)  ← installation Win32/Chocolatey, raccourcis, associations de fichiers, connexions d'imprimantes
    │   ├── <a href="scripts/TenantOnboarding/DeviceConfig/readme.fr.md">DeviceConfig/</a>         (6 scripts)  ← alimentation kiosque, désinstallation d'Office, disposition du menu Démarrer, règle de pare-feu Teams
    │   ├── <a href="scripts/TenantOnboarding/OneDriveManagement/readme.fr.md">OneDriveManagement/</a>   (3 scripts)  ← surveillance de la synchronisation, arrêt de la synchronisation de bibliothèques, redirection des dossiers connus
    │   └── <a href="scripts/TenantOnboarding/UserManagement/readme.fr.md">UserManagement/</a>       (2 scripts)  ← DG dynamique par filtre, appartenance aux groupes de fonctionnalités
    ├── <a href="scripts/Office365Toolkit/readme.fr.md">Office365Toolkit/</a>                ← réécriture du fork retiré de directorcia/Office365 (CIAOPS), hors menu
    │   ├── <a href="scripts/Office365Toolkit/readme.fr.md">readme.md</a>
    │   ├── <a href="scripts/Office365Toolkit/Security/readme.fr.md">Security/</a>             (4 scripts)  ← Secure Score, nettoyage des consentements d'applications, verrouillage des boîtes aux lettres partagées, référentiel EOP
    │   ├── <a href="scripts/Office365Toolkit/Exchange/readme.fr.md">Exchange/</a>             (4 scripts)  ← référentiel d'hygiène des boîtes aux lettres, risque de transfert, compléments, recherche dans le journal d'audit
    │   └── <a href="scripts/Office365Toolkit/Intune/readme.fr.md">Intune/</a>               (1 script)   ← inventaire des stratégies à l'échelle du tenant
    ├── <a href="scripts/PatronToolkit/readme.fr.md">PatronToolkit/</a>                    ← réécriture du fork retiré de directorcia/patron, hors menu
    │   ├── <a href="scripts/PatronToolkit/readme.fr.md">readme.md</a>
    │   ├── <a href="scripts/PatronToolkit/Entra/readme.fr.md">Entra/</a>                (2 scripts)  ← rapport d'inscription MFA, export des stratégies CA
    │   ├── <a href="scripts/PatronToolkit/Security/readme.fr.md">Security/</a>             (5 scripts)  ← consentements d'applications, règles de boîte de réception suspectes, alertes de sécurité, posture de sécurité de la messagerie, audit des boîtes aux lettres
    │   ├── <a href="scripts/PatronToolkit/Exchange/readme.fr.md">Exchange/</a>             (1 script)   ← rapport message trace
    │   ├── <a href="scripts/PatronToolkit/Intune/readme.fr.md">Intune/</a>               (2 scripts)  ← attributions de stratégies, appareils Autopilot
    │   ├── <a href="scripts/PatronToolkit/SharePoint/readme.fr.md">SharePoint/</a>           (1 script)   ← audit de la configuration du partage
    │   └── <a href="scripts/PatronToolkit/Teams/readme.fr.md">Teams/</a>                (1 script)   ← rapport de configuration Teams
    └── <a href="scripts/LegacyUtilities/readme.fr.md">LegacyUtilities/</a>                  ← scripts divers modernisés issus de la boîte à outils interne retirée, hors menu
        ├── <a href="scripts/LegacyUtilities/readme.fr.md">readme.md</a>
        ├── <a href="scripts/LegacyUtilities/Exchange/readme.fr.md">Exchange/</a>             (7 scripts)  ← droits sur les dossiers, accès délégué, boîtes aux lettres/contacts en masse, synchronisation des contacts, dédoublonnage, message trace
        ├── <a href="scripts/LegacyUtilities/Entra/readme.fr.md">Entra/</a>                (2 scripts)  ← appartenance aux groupes, sauvegarde des stratégies CA
        ├── <a href="scripts/LegacyUtilities/Teams/readme.fr.md">Teams/</a>                (3 scripts)  ← clonage d'équipes/de plans, provisionnement d'équipes projet
        ├── <a href="scripts/LegacyUtilities/Network/readme.fr.md">Network/</a>              (1 script)   ← mappage de lecteurs Azure Files
        ├── <a href="scripts/LegacyUtilities/Device/readme.fr.md">Device/</a>               (2 scripts)  ← NumLock par défaut, raccourci de verrouillage du poste
        └── <a href="scripts/LegacyUtilities/Workspace365/readme.fr.md">Workspace365/</a>         (2 scripts)  ← provisionnement/suppression d'environnements
</pre>

[`Deploy-OfficeTheme.ps1`](scripts/Custom%20Scripts/Intune/Desktop/Deploy-OfficeTheme.ps1) et [`Deploy-Officecolors.ps1`](scripts/Custom%20Scripts/Intune/Desktop/Office%20Themes/Deploy-Officecolors.ps1) reçoivent l'URL de téléchargement du thème en paramètre ; les fichiers de thème ne sont pas conservés dans le dépôt.

---

## Contribuer

Lorsque vous ajoutez de nouveaux scripts :

1. Respectez la convention de nommage existante (`Verb-Noun.ps1`)
2. Ajoutez un bloc de commentaires d'en-tête avec Synopsis, Description, Parameters et Example
3. Testez sur un tenant hors production avant de commiter
4. Placez le script dans le dossier de charge de travail approprié
5. Ajoutez-le à [`menu.ps1`](menu.ps1) et mettez à jour ce readme
6. Mettez à jour l'`Historique des versions` de ce fichier pour chaque changement fonctionnel ou structurel (obligatoire), y compris les changements demandés ou appliqués via Copilot/un assistant IA
7. Rédigez la modification dans les trois langues des readmes ([`readme.md`](readme.md), [`readme.nl.md`](readme.nl.md), [`readme.fr.md`](readme.fr.md))

**Ce qui se met à jour tout seul.** [`scripts/INDEX.md`](scripts/INDEX.md), le sélecteur de langue et le fil d'Ariane en tête de chaque readme, et la vérification des liens restent à jour automatiquement — vous ne les lancez pas à la main :

| Quand | Ce qui s'exécute |
|-------|------------------|
| Vous commitez | Le hook git [`.githooks/pre-commit`](.githooks/pre-commit) régénère l'index et les en-têtes des readmes, les ajoute au commit, et arrête le commit en cas de lien cassé. Il avertit lorsqu'un readme anglais a changé sans sa version néerlandaise/française — demandez alors à Claude de traduire |
| Claude Code modifie un fichier | Un hook dans [`.claude/settings.json`](.claude/settings.json) fait de même en arrière-plan après chaque modification, et avant que Claude ne termine il vérifie que les readmes modifiés ont été traduits et les scripts modifiés documentés |
| Vous poussez vers `devel`/`main` ou ouvrez une PR | [GitHub Actions](.github/workflows/ci.yml) refait les mêmes vérifications, pour les commits faits sans le hook ou dans l'éditeur web : syntaxe PowerShell (7 et 5.1), erreurs PSScriptAnalyzer, vérification des liens, `INDEX.md` et en-têtes des readmes à jour et les trois langues présentes dans chaque dossier, et ShellCheck sur les scripts `.sh` |

Activez le hook git une fois par clone :

```powershell
git config core.hooksPath .githooks
```

---

## Avertissement

Ces scripts sont fournis en l'état. Testez toujours dans un environnement hors production avant de les exécuter sur des tenants en production. Le mainteneur décline toute responsabilité pour les modifications involontaires résultant d'une mauvaise utilisation ou d'une mauvaise configuration.

---

## Historique des versions

> Remarque : les entrées plus anciennes peuvent faire référence à d'anciens noms de dossiers tels que [`Custom Scripts/`](scripts/Custom%20Scripts/readme.fr.md) et `Testing Scripts/`. Ces noms de chemins reflètent la structure du dépôt au moment de la modification concernée.

### 2026-10-10 (9)
| Modification |
|--------|
| **Le nouvel Outlook tombait sans cesse sur LEM-AVD-5** (n8n, 10-10) : à chaque nouvelle connexion son inscription échouait 10 à 16 fois avec `0x80073CF9` / `0x80070490`, le watchdog réinscrivait chaque utilisateur jusqu'à 30 minutes plus tard, et l'unique réparation de l'hôte (13:02) n'a rien changé. [`Repair-AppxPackageStore.ps1`](scripts/Device/readme.fr.md#repair-appxpackagestoreps1) ne cherchait la build demandée par les utilisateurs que dans le journal FSLogix, qui n'en nommait aucune sur cet hôte ; il n'avait donc rien à provisionner et terminait avec 0. Il compte désormais aussi une inscription AppX échouée avec `0x80070490` pour une build dont les fichiers ne sont pas sur l'hôte, et `-Provision` provisionne cette build depuis le CDN de Microsoft |
| [`Watch-M365Apps.ps1`](scripts/RDS/Watch-M365Apps.ps1) indique la build demandée pour chaque échec d'inscription, et fait d'un `0x80070490` pour une build de Teams/Outlook plus récente que celle provisionnée sur l'hôte un constat d'hôte *BuildMissing* : la réparation de l'hôte s'exécute immédiatement (une fois par build, sans attendre la pause) au lieu de réinscrire chaque utilisateur à chaque connexion, et le constat est réparé dès que l'hôte provisionne cette build. L'échec que sa propre réinscription journalise en chemin (vu à l'exécution suivante comme « 1x ... last at » l'heure de la réparation, à chaque exécution) n'est plus compté. Le watchdog et [`Update-SessionHostImage.ps1`](scripts/RDS/Update-SessionHostImage.ps1) épinglent l'outil sur `abe8202` |
| Vérifié : syntaxe ; `Get-OpenFailure` avec des événements simulés - deux utilisateurs demandant Outlook 1.2026.1001.300 sur un hôte qui provisionne 929.200 donnent *BuildMissing* plus leurs propres constats avec la build, et l'événement de notre propre réparation est écarté ; le format du message AppX correspond à l'expression du package ; les hachages épinglés correspondent à ce que sert GitHub. **Non** vérifié : quelle build les utilisateurs de LEM-AVD-5 demandent réellement et si le CDN de Microsoft la sert ; une exécution sur un hôte de session ; la carte n8n affiche *BuildMissing* / *NoShortcut* sous leur nom brut |

### 2026-10-10 (8)
| Modification |
|--------|
| [`Watch-M365Apps.ps1`](scripts/RDS/Watch-M365Apps.ps1) surveille la **nouvelle application Microsoft Copilot unifiée** au lieu de l'ancienne. Jusqu'ici l'ancienne application Microsoft 365 Copilot (`MicrosoftOfficeHub`) ou le Copilot grand public comptait comme Copilot, et avec l'application unifiée sur l'hôte chaque utilisateur était considéré en ordre sans test - l'application n'était jamais lancée. Désormais seule l'application unifiée compte : sur l'hôte, Edge Update doit l'avoir installée, son `copilotapp.exe` doit être sur le disque (trouvé via l'entrée de désinstallation, App Paths ou le dossier de type Edge - Microsoft n'en documente aucun) et un raccourci du menu Démarrer pour tous les utilisateurs doit pointer vers lui (créé s'il manque). Par utilisateur, son package d'identité, si l'hôte en a un, doit être inscrit et `Ok` - réparé en réinscrivant cette famille dans sa session, comme pour Teams et Outlook - et pour notre propre compte il est lancé (par AUMID, sinon `copilotapp.exe`) et doit rester actif. Un `copilotapp.exe` d'une build plus ancienne est fermé une fois dans notre session après une mise à jour Edge Update, afin que la nouvelle build soit testée |
| [`Get-M365AppsLog.ps1`](scripts/RDS/Get-M365AppsLog.ps1) reconnaît un tel watchdog et ne signale plus pour lui les anciens angles morts de Copilot ; le watchdog épingle cette version (`c2255d0`) |
| Vérifié : syntaxe ; `Test-UserCopilot` avec des cmdlets simulées - pas de package d'identité et notre compte (lancé via `copilotapp.exe`, `WontStart` s'il ne reste pas actif), package d'identité manquant chez un client (`NotRegistered` avec sa propre famille), présent et `Ok` (rien), pas d'application sur l'hôte (laissé au contrôle de l'hôte). **Non** vérifié sur un hôte de session : où se trouve réellement `copilotapp.exe`, si l'application unifiée inscrit un package d'identité par utilisateur, et le raccourci du menu Démarrer qu'elle crée |

### 2026-10-10 (7)
| Modification |
|--------|
| [`Watch-M365Apps.ps1`](scripts/RDS/Watch-M365Apps.ps1) maintient les applications à jour et teste lui-même chaque nouvelle build (nouvelle étape 1a, `-UpdateHours`, par défaut 6). Jusqu'ici il ne vérifiait et ne réparait que la build en place : une nouvelle build de Teams, Outlook ou Copilot n'arrivait sur l'hôte que par la mise à jour hebdomadaire de l'image ou par l'application elle-même, et c'était souvent un client qui la lançait le premier. Désormais, toutes les `-UpdateHours`, l'hôte reçoit les dernières versions de Teams et d'Outlook ([`Repair-AppxPackageStore.ps1`](scripts/Device/readme.fr.md#repair-appxpackagestoreps1) `-Latest -Provision`, uniquement plus récentes, signature vérifiée) et Edge Update est invité à vérifier Copilot ; à chaque exécution notre propre compte (`itceadmin`) passe à la build provisionnée s'il est en retard - application fermée dans notre session et réinscrite - de sorte que le test de lancement démarre la nouvelle build en premier. Les clients ne sont pas touchés ; Windows leur donne la build à leur prochaine connexion. Une nouvelle build sur l'hôte est signalée une fois en tant que `updated` (ancienne -> nouvelle), avec `updates` et `versions` dans le payload |
| Vérifié : syntaxe ; la détection des changements de version seule (la première exécution ne fait qu'enregistrer, une build Teams modifiée donne une ligne, état enregistré en JSON). **Non** vérifié : une exécution sur un hôte de session (provisionnement de la dernière build, réinscription dans la session d'`itceadmin`, vérification Edge Update) ; le flux n8n n'a pas encore de carte pour `updated` ; le téléchargement du [readme RDS](scripts/RDS/readme.fr.md#watch-m365appsps1) est épinglé sur `82807f8` (SHA-256 vérifié auprès de GitHub) ; les hôtes installés gardent l'ancienne version jusqu'à un nouveau `-Install` |

### 2026-10-10 (6)
| Modification |
|--------|
| [`Watch-M365Apps.ps1`](scripts/RDS/Watch-M365Apps.ps1) répare le problème d'un seul client dans la session de ce client uniquement, et l'hôte seulement quand cela dépasse un utilisateur : l'hôte lui-même, notre propre compte (`itceadmin`, le test pour tout l'hôte), ou la même application chez deux clients ou plus. Depuis les contrôles par utilisateur, chaque constat d'un client envoyait aussi tout l'hôte vers `Repair-AppxPackageStore.ps1 -Provision`, pour quelque chose qui ne concernait peut-être qu'un profil. Une réparation chez un client qui échoue est signalée comme *NIET HERSTELD* ; le niveau 2 décide pour l'hôte. [`Watch-M365Apps-ITGlue.md`](scripts/RDS/Watch-M365Apps-ITGlue.md) et son [HTML](scripts/RDS/Watch-M365Apps-ITGlue.html) disent la même chose. Le téléchargement du [readme RDS](scripts/RDS/readme.fr.md#watch-m365appsps1) passe à cette version |
| Vérifié sur des copies de test avec les scripts auxiliaires remplacés par un bouchon : un client avec Outlook cassé - pas de réparation de l'hôte ; `itceadmin` seul - `Repair-AppxPackageStore -Name outlook` ; deux clients avec Outlook - idem. **Non** vérifié : sur un hôte de session |

### 2026-10-10 (5)
| Modification |
|--------|
| [`Watch-M365Apps.ps1`](scripts/RDS/Watch-M365Apps.ps1) répare l'hôte quand une application échoue pour notre propre compte. Avant, *non inscrit* ou *ne démarre pas* chez `itceadmin` ne faisait que réinscrire ou réinitialiser l'application dans la session d'`itceadmin`, alors que ce compte est là pour représenter tout le monde sur l'hôte - l'hôte restait tel quel jusqu'à ce qu'un client tombe sur le problème. Désormais chaque constat, y compris celui d'`itceadmin`, envoie l'hôte vers `Repair-AppxPackageStore.ps1 -Provision` pour cette application (dans le délai de carence). La réparation de l'hôte ne lance rien dans la session de personne ; les clients ne voient toujours une application ouverte qu'après leur propre tentative récente. Le téléchargement du [readme RDS](scripts/RDS/readme.fr.md#watch-m365appsps1) passe à cette version |
| Vérifié sur une copie de test avec les scripts auxiliaires remplacés par un bouchon : un Outlook en échec chez `itceadmin` seul lance désormais `Repair-AppxPackageStore -Name outlook -Provision`, et rien ne tourne dans une session. **Non** vérifié : une vraie réparation de l'hôte sur un hôte de session |

### 2026-10-10 (4)
| Modification |
|--------|
| [`Watch-M365Apps.ps1`](scripts/RDS/Watch-M365Apps.ps1) n'ouvre une application pour un utilisateur après une réparation que s'il l'a essayée **lui-même, récemment** : sa dernière ouverture refusée date d'au plus 30 minutes, et pas des 2 minutes suivant la connexion. Avant, toute ouverture refusée depuis l'exécution précédente comptait - y compris celle du démarrage automatique de Teams ou du nouvel Outlook à la connexion, que Windows journalise de la même façon, et celle d'un utilisateur qui avait abandonné depuis longtemps - si bien qu'une application pouvait apparaître à l'écran d'un client qui ne l'avait pas demandée. Sinon l'action dit pourquoi elle n'a pas été ouverte. Rien d'autre n'est jamais lancé pour un client. Le téléchargement du [readme RDS](scripts/RDS/readme.fr.md#watch-m365appsps1) passe à cette version |
| Vérifié : syntaxe ; le nouveau contrôle seul pour une tentative récente (ouverte), une tentative vieille de 50 minutes, une 30 secondes après la connexion (toutes deux non ouvertes, avec la raison) et une 2,5 minutes après la connexion (ouverte). **Non** vérifié : si le démarrage automatique écrit vraiment TWinUI `5961`, et une exécution sur un hôte de session |

### 2026-10-10 (3)
| Modification |
|--------|
| Le flux n8n *ITCE – M365 App Watchdog → Teams* a une carte pour le nouveau signalement `repairing` : 🔧 **WORDT HERSTELD**, avec les utilisateurs et applications trouvés et une ligne *Bewijs* portant le chemin du zip de preuves (sur toute carte qui en reçoit un). Jusqu'ici il retombait sur une carte *PROBLEEM*. [`Watch-M365Apps-ITGlue.md`](scripts/RDS/Watch-M365Apps-ITGlue.md) explique la carte et ce qu'en fait le niveau 2, qu'une application qu'un utilisateur n'a pas pu ouvrir s'ouvre désormais d'elle-même après la réparation, où se trouvent les preuves et comment les collecter avec `Get-M365AppsLog.ps1`, l'application Copilot unifiée comme quelque chose que le watchdog ne teste pas par utilisateur, et que `-Install` ne reprend pas l'ancien `config.json`. [HTML](scripts/RDS/Watch-M365Apps-ITGlue.html) régénéré |
| Vérifié : le flux a été enregistré et publié ; son code de carte exécuté localement sur un exemple `repairing` et `repaired` a donné le titre, le libellé, les utilisateurs et la ligne *Bewijs* attendus ; le HTML est à jour (`-Check`) et la vérification des liens du dépôt passe. **Non** vérifié : une vraie carte dans le canal Teams CIPP (aucune carte de test n'y a été envoyée), le collage dans IT Glue |

### 2026-10-10 (2)
| Modification |
|--------|
| **[`Watch-M365Apps.ps1`](scripts/RDS/Watch-M365Apps.ps1) annonce avant de réparer, garde les preuves, et ouvre l'application pour l'utilisateur qui a essayé.** Jusqu'ici n8n n'apprenait un problème qu'après la réparation, rien ne gardait l'état de l'hôte, et un utilisateur dont Windows avait refusé l'ouverture devait réessayer. Quand un nouveau problème apparaît - pour tout utilisateur connecté, pas seulement `itceadmin` - il lance désormais [`Get-M365AppsLog.ps1`](scripts/RDS/Get-M365AppsLog.ps1) pour ces utilisateurs et applications (un zip dans `C:\IT\AppWatchdog\Diag`, conservé 14 jours, chemin dans `diagnostics`), envoie ensuite `repairing` avec chaque constat et utilisateur, puis répare et signale `repaired` / `repair-failed` comme avant. Le même problème n'est ni réannoncé ni recollecté dans le délai `-RenotifyHours`. Après avoir réinscrit une application qu'un utilisateur n'a pas pu ouvrir (TWinUI `5961`), il l'ouvre dans sa session ; si elle ne démarre pas ou ne reste pas ouverte, le problème reste ouvert. `-NoUserLaunch` et `-NoDiagnostics` désactivent ces deux comportements. `copilotapp.exe`, l'application Copilot unifiée, compte désormais pour les plantages et pour « lancée ». `-Install` copie aussi le collecteur, depuis le dépôt ou depuis GitHub à `6a94769` avec son SHA-256 |
| Le téléchargement du [readme RDS](scripts/RDS/readme.fr.md#watch-m365appsps1) passe à `7c9d2f3` et au SHA-256 de ce fichier (vérifié par rapport à ce que sert GitHub), pour qu'un hôte installé depuis GitHub reçoive cette version |
| Vérifié sur ce poste Windows 11, sans élévation, sur une copie de test (contrôle admin et verrouillage du dossier désactivés, une ouverture de Copilot refusée simulée, la réparation remplacée par un bouchon, un écouteur local comme webhook) : preuves collectées d'abord (zip gardé, dossier supprimé), puis `repairing` avec l'utilisateur et le chemin du zip, puis `repaired` avec `Fixed` et le même chemin ; l'épinglage du collecteur correspond à ce que sert GitHub. **Non** vérifié : une exécution en System sur un hôte de session, l'ouverture d'une application dans la session d'un vrai client, et la carte n8n - le flux ne connaît pas encore `repairing` et l'affiche comme carte *PROBLEEM* |

### 2026-10-10
| Modification |
|--------|
| **Nouveau [`scripts/RDS/Get-M365AppsLog.ps1`](scripts/RDS/Get-M365AppsLog.ps1) : après une plainte, collecter ce qui s'est passé avec le nouveau Teams, le nouvel Outlook et Copilot sur un hôte de session, et pourquoi le watchdog l'a réparé ou non.** Copilot ne fonctionnait pas pour des utilisateurs sur un serveur alors qu'on pensait que le watchdog contrôlait, signalait et réparait chaque utilisateur ; rien ne permettait de voir après coup ce qu'il avait vu. En lecture seule, dans un dossier et un zip : les réglages du watchdog (webhook et jeton masqués), tâche, état et journaux ; builds provisionnés, l'application Copilot unifiée et WebView2 ; l'état des packages de chaque utilisateur (`Get-AppxPackage -AllUsers`) ; par utilisateur connecté et application ce qui est inscrit et lancé, avec le verdict auquel arriverait le watchdog et **BLIND SPOT** là où il le jugerait correct sans test ; les événements AppX, AppReadiness, AppModel-Runtime, TWinUI et de plantage avec utilisateur, code et si le watchdog les lit ; `PackageStatus`, `Deprovisioned`, registre et journaux de FSLogix et Edge Update ; avec `-IncludeAppLogs` les journaux Teams et Outlook par utilisateur. Chaque partie tourne seule, `Get-AppxPackage` dans un job avec un délai de 120 s, les canaux d'événements absents sont ignorés, les journaux ouverts sont lus en partage. Les canaux et fichiers suivent MSRD-Collect de Microsoft et la documentation FSLogix sur AppX. Entrée de menu `Q` |
| Angles morts du watchdog qu'il montre, et qui expliquent probablement le cas Copilot : avec l'application Copilot unifiée (Edge Update) sur l'hôte, le watchdog accepte Copilot pour chaque utilisateur sans le tester ; cette application tourne sous `copilotapp.exe`, que le contrôle des plantages du watchdog ne connaît pas ; le Copilot grand public compte comme Copilot ; et l'application d'un client inscrite mais qui ne fonctionne pas n'est vue que si Windows a refusé l'ouverture (TWinUI `5961`) |
| Vérifié sur ce poste Windows 11, sans élévation : syntaxe ; une exécution complète et une exécution filtrée sur un utilisateur et Copilot ; le verdict par utilisateur et les événements (Teams `401`/`404`/`419` avec `0x80073D02` marqué bénin, un plantage d'Outlook marqué comme lu) ; un canal FSLogix et une clé de registre absents ignorés au lieu d'arrêter la partie ; une partie nécessitant l'élévation nommée à la fin pendant que le reste était collecté ; journaux d'application copiés et zip écrit. **Non** vérifié : une exécution élevée sur un hôte de session avec le watchdog installé, FSLogix présent et l'application Copilot unifiée |

### 2026-10-09 (10)
| Modification |
|--------------|
| [`Watch-M365Apps-ITGlue.md`](scripts/RDS/Watch-M365Apps-ITGlue.md) faisait lire au niveau 1 les cartes du watchdog dans le canal Teams CIPP, mais seuls les niveaux 2 et 3 ont accès à ce canal. Le niveau 1 interroge désormais, apporte une première aide sans les cartes (fermer complètement l'application et la rouvrir, attendre un quart d'heure après la connexion, se déconnecter et se reconnecter) et escalade avec ce qui a été tenté ; le niveau 2 reçoit une première étape pour trouver la carte de l'utilisateur et agir en conséquence, puis les commandes sur l'hôte comme avant. Le tableau des cartes dit ce que fait le niveau 2, et les phrases pour l'utilisateur sont réparties par niveau. Les titres vers lesquels pointent des liens utilisent deux-points au lieu d'un tiret, pour que les ancres soient identiques pour GitHub et le convertisseur HTML. [HTML](scripts/RDS/Watch-M365Apps-ITGlue.html) régénéré |
| Vérifié : chaque lien interne du HTML pointe vers un id existant, et le contrôle des liens du dépôt passe pour le markdown. **Non** vérifié : le collage dans IT Glue |

### 2026-10-09 (9)
| Modification |
|--------------|
| Nouveau [`scripts/RDS/Watch-M365Apps-ITGlue.md`](scripts/RDS/Watch-M365Apps-ITGlue.md) - la procédure service desk du watchdog, en néerlandais, à coller dans IT Glue, avec une [version HTML](scripts/RDS/Watch-M365Apps-ITGlue.html) mise en forme par `Convert-MarkdownToHtml.ps1`. Elle explique quels scripts interviennent (le watchdog, Repair-AppxPackageStore, Update-SessionHostImage, Update-TeamsClient, le flux n8n) et comment ils s'articulent ; ce que signifie chaque carte Teams et quoi en faire ; pour le niveau 1, quoi demander à un utilisateur qui appelle, où regarder, quoi faire par plainte, quoi dire et quand escalader ; pour le niveau 2, les commandes sur l'hôte (lancer le watchdog maintenant, suivre le journal, regarder sans réparer, dernière réparation, examiner l'hôte) et ce qu'il ne faut pas faire (pas de réinitialisation chez un utilisateur) ; ce que le watchdog ne voit pas ; les codes d'erreur courants ; et pour le niveau 3 l'installation, les réglages, le flux n8n et la maintenance hebdomadaire. Elle suit la forme de la procédure Teams existante. L'URL du webhook et le jeton n'y figurent pas - le dépôt est public |
| Vérifié : le HTML a été généré par le convertisseur du dépôt et s'analyse comme du XML ; le lien interne vers *Als een gebruiker belt* aboutit ; aucune URL de webhook ni jeton dans les deux fichiers ; contrôle des liens du dépôt. **Non** vérifié : le collage dans IT Glue, et la procédure face à un vrai appel |

### 2026-10-09 (8)
| Modification |
|--------------|
| **[`Watch-M365Apps.ps1`](scripts/RDS/Watch-M365Apps.ps1) trouve et répare désormais ce que rencontrent les clients, et dit chez qui.** Il ne testait que notre propre compte, si bien qu'un client qui cliquait sur Teams sans résultat passait inaperçu. Chaque exécution vérifie maintenant chaque autre utilisateur connecté une fois connecté depuis 10 minutes (inscription, fichiers, état - rien n'est lancé pour lui), et lit depuis l'exécution précédente chaque ouverture refusée par Windows (TWinUI `5961`) et chaque inscription échouée des packages des applications (AppXDeploymentServer `401`/`404`, par champ plutôt que par le texte traduit, sans « fermez d'abord l'application » ni « déjà installé », les deux événements d'un même échec comptés une fois), chacun avec l'utilisateur de l'événement. Il répare l'hôte via `Repair-AppxPackageStore.ps1 -Provision` puis réinscrit le package par nom de famille dans la session de cet utilisateur, via une tâche ponctuelle lançant `conhost --headless` pour qu'aucune fenêtre n'apparaisse - seulement si l'application ne tourne pas chez lui, une fois par délai de carence par utilisateur et application, jamais de réinitialisation. `-NoUserRepair` reste en dehors des sessions des clients. Chaque constat porte l'utilisateur (`Account`, `Customer`) et s'il a été réparé pour lui (`Fixed`) ; la carte n8n affiche les utilisateurs, *hersteld bij deze gebruiker* par ligne et une section *Wel hersteld* après une réparation partielle |
| `Get-AppxPackage -User` reçoit désormais `DOMAINE\utilisateur` au lieu d'un SID : avec un SID Entra ID (`S-1-12-1-...`), il répond *No valid SID could be determined*, dans Windows PowerShell 5.1 aussi - la vérification existante d'itceadmin ne pouvait donc pas fonctionner sur un hôte joint à Entra. Le téléchargement du [readme RDS](scripts/RDS/readme.fr.md#watch-m365appsps1) passe à `abfad94` |
| Vérifié sur cette machine Windows 11, sans élévation : la liste des sessions (un utilisateur Entra, avec l'heure de connexion) ; sur ses vrais journaux, le lecteur n'a rien trouvé avec le filtre livré, cinq inscriptions Teams échouées (`0x80073D02`, 401 et 404 comptés une fois) quand ce code n'était pas filtré, et des refus TWinUI `5961` avec utilisateur et code pour un composant Windows tenant lieu d'application ; HRESULT signés et non signés ; la vérification d'inscription d'un « client » (silencieuse pour Teams, Outlook et Copilot, un constat avec nom et SID pour un package absent) ; `Get-AppxPackage -User` échouant par SID et fonctionnant par nom ; `conhost --headless` attendant sa commande mais renvoyant `0` pour `exit 7` ; dans n8n, une charge de réparation partielle a produit la carte attendue. **Non** vérifié : une exécution en tant que System sur un hôte de session, une réinscription dans une vraie session client, et si quelque chose clignote à l'écran |

### 2026-10-09 (7)
| Modification |
|--------------|
| [`Watch-M365Apps.ps1`](scripts/RDS/Watch-M365Apps.ps1) `-TestNotification` s'arrêtait sur le premier hôte de session où il a été lancé avec *The variable '$script:logFile' cannot be retrieved because it has not been set* : le rapport contient le chemin du journal, et cette variable n'était définie qu'à l'ouverture du journal d'une exécution - après que `-TestNotification` avait déjà pris son propre chemin. Sous StrictMode, une variable non définie est une erreur, pas `$null`. Elle est désormais définie en tête avec les autres constantes. Les exécutions planifiées n'étaient pas touchées. Le téléchargement du [readme RDS](scripts/RDS/readme.fr.md#watch-m365appsps1) passe à `960576b` et au SHA-256 de ce fichier |
| Vérifié : contrôle de syntaxe ; `Send-Notification` extrait du script et exécuté sous StrictMode vers une adresse injoignable a levé l'erreur signalée sans la ligne et, avec elle, est allé jusqu'à trois tentatives de connexion et a renvoyé `$false` ; l'URL raw à `960576b` sert le fichier avec le SHA-256 documenté. **Non** vérifié : un message de test de l'hôte de session vers n8n |

### 2026-10-09 (6)
| Modification |
|--------------|
| [`Watch-M365Apps.ps1`](scripts/RDS/Watch-M365Apps.ps1) ne surveille plus par défaut que `itceadmin` au lieu d'`itceadmin` et `itce.user` : un seul compte à garder connecté sur chaque hôte, et une session de moins où le test de lancement peut ouvrir une fenêtre. `-Account itceadmin,itce.user` teste toujours les deux. Le téléchargement du [readme RDS](scripts/RDS/readme.fr.md#watch-m365appsps1) est épinglé sur `1182c3a` avec le SHA-256 de ce fichier, et la question du menu affiche la nouvelle valeur par défaut |
| Vérifié : contrôle de syntaxe du script et de `menu.ps1` ; l'URL raw à `1182c3a` sert le fichier avec `$Account = @('itceadmin')` et le SHA-256 documenté. **Non** vérifié sur un hôte de session |

### 2026-10-09 (5)
| Modification |
|--------------|
| **[`Install-Printer.ps1`](scripts/Device/Printer/Install-Printer.ps1) peut réserver une file à des groupes de sécurité.** Nouveau champ d'imprimante `permissions` : une liste de `{ "principal", "access" }` avec access `print`, `manageDocuments` ou `manage`, et le principal sous forme `DOMAINE\Groupe`, d'un nom local ou d'un SID (seul moyen d'indiquer un groupe Entra ID). Les principaux listés reçoivent cet accès et tout autre utilisateur ou groupe perd le sien - mais Administrators, SYSTEM, CREATOR OWNER, les packages d'application et les identités de service gardent toujours le leur : une ACL d'imprimante par défaut accorde aussi l'accès à ALL APPLICATION PACKAGES et à un SID de capacité d'impression, et remplacer l'ACL en bloc casserait en silence l'impression depuis le nouvel Outlook et d'autres applications empaquetées. L'état actuel et l'état voulu sont comparés selon le sens des masques d'accès plutôt que le texte SDDL, si bien qu'une entrée que Windows enregistre sous une autre forme ne compte pas comme une modification à chaque exécution. Échec sécurisé : un groupe introuvable empêche la création de cette file (`-CheckOnly` le signale à l'avance), et une nouvelle file dont les autorisations ne peuvent être définies est supprimée plutôt que laissée ouverte à tous. Une file qui ne diffère que par ses autorisations, son bac ou ses paramètres d'impression ne reçoit plus de `Set-Printer` superflu. Documenté dans le [readme Printer](scripts/Device/Printer/readme.fr.md#install-printerps1) ; le bac par file a été ajouté plus tôt le même jour (`inputBin`) |
| Vérifié sous Windows PowerShell 5.1 et PowerShell 7 : résolution des noms (`Everyone`, `BUILTIN\Users`, un utilisateur Entra, un SID, un groupe inconnu avec et sans nouvelle tentative) ; le plan d'ACL sur la vraie `AnyDesk Printer` de cette machine (Everyone et l'utilisateur installateur retirés, Users ajouté, entrées des packages d'application, de la capacité, des services, d'Administrators et de CREATOR OWNER conservées) ; un second plan sur l'ACL planifiée ne signale aucune modification, même après réécriture des droits génériques comme Windows peut les enregistrer ; la validation d'entrées mal formées ; `-CheckOnly` et `-WhatIf` de bout en bout, y compris une file dont le groupe est introuvable. Trouvé et corrigé en chemin : `-band` sur l'`AceFlags` de la taille d'un octet lève « Specified cast is not valid » sous Windows PowerShell 5.1. Ces exécutions contournaient la vérification d'élévation, cette session n'étant pas élevée. **Non** vérifié : la définition effective des autorisations avec `Set-Printer -PermissionSDDL`, et ce que voit un utilisateur sans accès |

### 2026-10-09 (5)
| Modification |
|--------|
| [`Watch-M365Apps.ps1`](scripts/RDS/Watch-M365Apps.ps1) s'installe directement depuis GitHub : le [readme RDS](scripts/RDS/readme.fr.md#watch-m365appsps1) contient désormais un téléchargement pour un hôte et pour plusieurs via PowerShell remoting, épinglé sur un commit et le SHA-256 du fichier, parce que ce qu'il installe s'exécute en tant que System. `-Install` apporte lui-même le script auxiliaire (commit épinglé, empreinte vérifiée). Lancé depuis un tel téléchargement, le script cherchait `Repair-AppxPackageStore.ps1` à côté de lui et dans `..\Device` - alors `C:\IT\Device` - des dossiers où tout utilisateur peut éventuellement écrire, et `-Install` aurait copié ce qu'il y trouvait dans le dossier exécuté par System. Il ne prend désormais le script auxiliaire que dans son propre dossier verrouillé (copie installée), dans `..\Device` d'une extraction du dépôt (reconnue à `menu.ps1`), et sinon depuis GitHub ; et chaque exécution, pas seulement `-Install`, verrouille d'abord `C:\IT\AppWatchdog`. L'URL du webhook et le jeton sont des espaces réservés dans le readme : le dépôt est public |
| Vérifié : contrôle de syntaxe ; l'URL raw à `6c088f4` sert le fichier avec ce correctif et le SHA-256 documenté ; la recherche du script auxiliaire dans Windows PowerShell 5.1 dans trois cas avec des fichiers déposés à côté du script et dans `..\Device` - depuis le dépôt il a utilisé `scripts\Device`, depuis un téléchargement isolé il a ignoré les deux fichiers déposés et récupéré le script auxiliaire sur GitHub avec l'empreinte vérifiée, en copie installée il n'a utilisé que son propre dossier. **Non** vérifié : le téléchargement et `-Install` sur un hôte de session, et la boucle de remoting |

### 2026-10-09 (4)
| Modification |
|--------|
| **[`Watch-M365Apps.ps1`](scripts/RDS/Watch-M365Apps.ps1) signale aussi quand Teams, le nouvel Outlook ou Copilot plante ou se bloque.** Le watchdog ne voyait qu'une application incapable de démarrer ; une application qui démarrait puis plantait - chez un client, ou dans notre compte entre deux exécutions - passait inaperçue. Chaque exécution lit désormais Application Error `1000` et Application Hang `1002` dans le journal Application depuis l'exécution précédente (première exécution : la dernière heure, jamais plus d'un jour en arrière), garde les événements des trois applications - par nom de processus, ou par le nom de package que porte l'événement, pour qu'un exécutable Copilot absent de la liste compte aussi - et les regroupe par application, module et code d'exception. Ils accompagnent tout signalement de l'exécution, ou forment un signalement à part (`crashed`) ; `-CrashThreshold` (par défaut `1`, `0` = désactivé) fixe le nombre par application avant envoi. Les plantages sont signalés, pas réparés : l'application d'un client n'est jamais relancée à sa place, et celle de notre propre compte est relancée par le test de lancement existant. Le flux n8n *ITCE – M365 App Watchdog → Teams* les affiche dans une section *Crashes en hangs* et sur une carte orange `CRASH` |
| Vérifié : contrôle de syntaxe ; sur le journal Application de cette machine, le lecteur a trouvé le vrai blocage du nouvel Outlook du 1er octobre (`olk.exe` 1.2026.915.300, type de blocage `Quiesce`), y compris avec son nom de processus retiré de la liste (correspondance par nom de package), n'a rien renvoyé pour la dernière heure, et a regroupé trois plantages de PowerToys en une ligne avec le module `Microsoft.UI.Xaml.dll` et le code `0xc000027b` quand PowerToys était dans la liste pour le test ; dans n8n, une exécution de test avec une charge `crashed` (envoi Teams épinglé) a produit la carte attendue. **Non** vérifié : un plantage de Teams ou de Copilot lui-même, et une exécution en tant que System sur un hôte de session |

### 2026-10-09 (3)
| Modification |
|--------|
| **Nouveau [`Watch-M365Apps.ps1`](scripts/RDS/Watch-M365Apps.ps1) : un watchdog qui teste le nouveau Teams, le nouvel Outlook et Copilot avec nos propres comptes (`itceadmin`, `itce.user`) sur un hôte de session, les répare avant qu'un client ne s'en aperçoive, et le signale à n8n.** Jusqu'ici, une application cassée était découverte quand un client appelait ; Update-SessionHostImage et Repair-AppxPackageStore la réparent, mais seulement si quelqu'un les lance. Installé avec `-Install` comme tâche *M365 App Watchdog* (System, toutes les 30 minutes), il vérifie que l'hôte provisionne les trois applications et, pour chaque compte surveillé connecté à l'hôte, que le package est inscrit, intact et démarre dans cette session (une tâche ponctuelle avec le SID de l'utilisateur et un jeton interactif, donc sans mot de passe). Les problèmes de l'hôte passent par `Repair-AppxPackageStore.ps1 -Provision` - au plus une fois toutes les 4 heures, sans `-RemoveOld` / `-Latest`, pour que rien de ce qu'un client a ouvert ne soit supprimé - et notre propre compte est réinscrit ou réinitialisé. Il relit tout et envoie en POST `repaired`, `repair-failed`, `recovered` ou `error` en JSON à un webhook n8n (en-tête `X-Watchdog-Token`), uniquement lors d'un changement et à nouveau après 12 heures pour un problème qui persiste. Son dossier est restreint à System et Administrateurs et contient l'URL du webhook et le jeton ; le script auxiliaire vient de `..\Device` ou de GitHub au même épinglage et à la même empreinte qu'Update-SessionHostImage. Entrée de menu `Y` |
| Vérifié : contrôle de syntaxe dans PowerShell 7 et l'analyseur de Windows PowerShell 5.1 ; le SHA-256 du script auxiliaire épinglé par rapport au commit `048cf96` ; sans élévation sur une machine Windows 11, la session d'un compte Entra (`AzureAD\...`) trouvée via le propriétaire d'explorer.exe, et l'entrée de démarrage et le processus de Teams, du nouvel Outlook et de Microsoft 365 Copilot depuis leurs manifestes - ce qui a montré que Teams liste `MSTeamsRemoteModuleContainer` en premier, d'où l'utilisation de la première entrée visible dans Démarrer. **Non** vérifié : une exécution en tant que System ou en élevé, `-Install`, le lancement ou la réinscription d'une application via une tâche dans la session d'un autre utilisateur, les réparations, et un POST vers un vrai webhook n8n - aucun hôte de session ni webhook n'était disponible |

### 2026-10-09 (2)
| Modification |
|--------|
| [`Install-Printer.ps1`](scripts/Device/Printer/Install-Printer.ps1) définit un **bac par défaut** par file : le nouveau champ d'imprimante `inputBin`. Les bacs ne sont pas un paramètre de Set-PrintConfiguration ; ils se trouvent dans le Print Schema comme options de `psk:JobInputBin` / `DocumentInputBin` / `PageInputBin`, avec des noms du fabricant comme `ns0000:Tray2`. Le script lit les PrintCapabilities du pilote, compare le nom du JSON au nom affiché ou au nom Print Schema (espaces et casse ignorés, donc `Tray 2`, `tray2` et `ns0000:Tray2` sont un seul bac), et l'écrit dans le ticket d'impression par défaut de la file pour chaque fonction de bac qui le propose - dans un job avec délai, comme les autres paramètres d'impression. Un nom inconnu du pilote donne un avertissement listant ce qu'il propose, pas une modification à chaque exécution |
| Un appareil avec plusieurs files - `Office` sur le bac 1, `Office - letterhead` sur le bac 2 - est documenté dans le [readme des imprimantes](scripts/Device/Printer/readme.fr.md#install-printerps1) et montré dans `printers.example.json` : même adresse, port `IP_<address>` partagé, bac et paramètres propres |
| Vérifié : contrôle de syntaxe ; contrôle des liens ; dans Windows PowerShell 5.1, la lecture des bacs avec les pilotes de cette machine (AnyDesk : `Automatically Select`, `Upper Paper Tray` ; Fax : un bac ; PDF et OneNote : aucun, signalé comme tel), la correspondance des noms y compris un échec, et la réécriture du ticket avec Set-PrintConfiguration simulé (espace de noms du fabricant déclaré, sous-propriétés de l'ancienne option supprimées, PageInputBin ajouté) ; le JSON d'exemple passe la validation du script et un JSON erroné donne les deux erreurs. **Non** vérifié : définir un bac sur un vrai pilote d'imprimante (nécessite l'élévation et un appareil à plusieurs bacs) |

### 2026-10-09
| Modification |
|--------|
| [`Install-Printer.ps1`](scripts/Device/Printer/Install-Printer.ps1) démarre au déploiement comme documenté. La ligne de commande que le readme et l'aide donnaient pour l'extension Custom Script et une tâche de démarrage, `powershell.exe -File Install-Printer.ps1 ... -Confirm:$false`, s'arrêtait avant le début du script : `-File` transmet chaque argument comme texte, donc `-Confirm:$false` arrivait comme la chaîne `'$false'` (*Cannot convert 'System.String' to the type SwitchParameter*) et aucune imprimante n'était installée. Ces exécutions n'ont pas de console et ne demandent jamais rien, le commutateur est donc retiré des exemples. Le relancement du script lui-même - en 64 bits quand Intune ou un RMM le lance en 32 bits, et pour l'élévation - transmettait `-Switch:$false` de la même façon ; il passe désormais ses paramètres via `-Command` (élévation : `-EncodedCommand`, car Start-Process perd les guillemets) |
| Nouveau dans le [readme des imprimantes](scripts/Device/Printer/readme.fr.md#install-printerps1) : installer au déploiement plutôt que dans l'image - un exemple Bicep avec Custom Script Extension et un exemple `az vm run-command`, `fileUris` épinglé sur un commit, une seule extension Custom Script par VM (sinon un Run Command managé), le jeton GitHub dans `protectedSettings`, et le code de sortie `1` qui fait échouer l'extension |
| Vérifié : contrôle de syntaxe ; contrôle des liens ; depuis Windows PowerShell 5.1 32 bits, le relancement est arrivé en 64 bits avec un chemin UNC contenant `$`, deux noms d'imprimante, `-Quiet`, `-WaitSeconds` et `-Confirm:$false` intacts et le code de sortie renvoyé ; l'ancienne ligne `-File ... -Confirm:$false` a reproduit l'erreur. **Non** vérifié : le relancement pour l'élévation (il ouvre une invite UAC), et un vrai déploiement par Custom Script Extension ou Run Command |

### 2026-10-08 (21)
| Modification |
|--------|
| [`Update-SessionHostImage.ps1`](scripts/RDS/Update-SessionHostImage.ps1) récupère ses scripts auxiliaires sur `main` à `048cf96` (2026-10-08) au lieu de `746541e`. Entre les deux, Update-TeamsClient.ps1 a changé en un commit (`8f74ebc`) : il installe désormais exactement la build Teams que le service de configuration Teams désigne comme actuelle, via son `buildLink` avec contrôle de taille et de signature Microsoft, au lieu de ce que distribue le déploiement progressif de teamsbootstrapper - les deux divergeaient pendant des semaines, si bien qu'un hôte restait « obsolète ». Ce diff a été lu avant l'épinglage. Repair-AppxPackageStore.ps1 n'a pas changé ; son empreinte reste |
| Vérifié : les deux SHA-256 calculés avec `git show` à `048cf96` et comparés à ce que raw.githubusercontent.com fournit pour ce commit ; contrôle de syntaxe. **Non** vérifié : une exécution sur un hôte de session |

### 2026-10-08 (20)
| Modification |
|--------|
| **Les 23 scripts qui ne s'analysaient pas sous Windows PowerShell 5.1 passent maintenant, et le job 5.1 de la CI bloque à nouveau.** 22 étaient en UTF-8 sans BOM avec un caractère comme `—` ou `é` : 5.1 lit un tel fichier en ANSI, un octet de ce caractère devient un guillemet isolé, et le script s'arrête sur des erreurs comme « string is missing the terminator » - précisément sur les machines qui utilisent 5.1 : Intune, GPO, tâches planifiées. Chacun a reçu un BOM UTF-8 et rien d'autre ; le contenu est identique octet pour octet. Parmi eux [`load.ps1`](load.ps1), [`Remove-OemBloatware.ps1`](scripts/Device/Remove-OemBloatware.ps1), les deux scripts audio, [`New-CloudDriveMapping.ps1`](scripts/Device/DriveMapping/New-CloudDriveMapping.ps1), les scripts Intune CoworkPrerequisites, les scripts SAS et SMTP et [`Invoke-TeamsArchive.ps1`](scripts/Teams/Invoke-TeamsArchive.ps1). [`create_scheduled_task.ps1`](scripts/Reporting/Licensing/create_scheduled_task.ps1) indique `#Requires -Version 5.1` mais utilisait `?.Source`, une syntaxe PowerShell 7 ; c'est maintenant `Select-Object -ExpandProperty Source` |
| [`ci.yml`](.github/workflows/ci.yml) : le job Windows PowerShell 5.1 n'est plus `continue-on-error`, donc un nouveau script enregistré en UTF-8 sans BOM avec un caractère non ASCII fait échouer l'exécution |
| Vérifié : chaque `.ps1`/`.psm1` sans `#Requires -Version 7` s'analyse avec `powershell.exe` 5.1 sur cette machine (0 erreur, il y en avait 23), et [`Test-PowerShellSyntax.ps1`](scripts/Startup/Test-PowerShellSyntax.ps1) sous PowerShell 7 reste propre. **Non** vérifié : les scripts eux-mêmes n'ont pas été exécutés - c'est l'analyse qui a changé, pas leur logique |

### 2026-10-08 (19)
| Modification |
|--------|
| [`Update-SessionHostImage.ps1`](scripts/RDS/Update-SessionHostImage.ps1) ne pouvait pas exécuter ses scripts auxiliaires : il lançait Repair-AppxPackageStore.ps1 et Update-TeamsClient.ps1 avec `powershell.exe -File ... -Confirm:$false`, et `-File` transmet chaque argument comme texte, donc `-Confirm:$false` arrivait comme la chaîne `'$false'` et les deux s'arrêtaient aussitôt avec *Cannot convert 'System.String' to the type SwitchParameter* - aucune application n'était jamais mise à jour. Son propre relancement de PowerShell 7 vers Windows PowerShell avait le même défaut. Les deux passent désormais par `-Command` avec les paramètres écrits en clair (chaînes entre apostrophes, tableaux en liste, commutateurs en `-Nom:$true/$false`) et transmettent le code de sortie |
| Vérifié : contrôle de syntaxe ; la construction de la commande a été exécutée contre un script de test dans Windows PowerShell 5.1 avec un tableau (contenant une apostrophe), des commutateurs activés et désactivés et `-Confirm:$false` sur un script `ConfirmImpact = 'High'` - chaque valeur est arrivée correctement, sans invite, code de sortie transmis. **Non** vérifié : une exécution sur un hôte de session |

### 2026-10-08 (18)
| Modification |
|--------|
| **[`TenantOnboarding/`](scripts/TenantOnboarding/readme.fr.md), [`Intune/`](scripts/Intune/readme.fr.md) et [`Azure/`](scripts/Azure/readme.fr.md) sur [`Connect-M365.ps1`](scripts/Startup/readme.fr.md#connect-m365ps1), et le dernier code AzureAD supprimé.** [`Get-WindowsAutoPilotInfo.ps1`](scripts/Intune/Get-Autopilot/readme.fr.md) (v3.5.1) utilisait encore les modules retirés AzureAD et Microsoft.Graph.Intune, si bien que `-Online`, `-AddToGroup` et `-Assign` ne fonctionnaient plus ; cette partie dialogue maintenant directement avec Graph (il reste autonome et compatible Windows PowerShell 5.1 pour la boîte à outils USB), délégué par défaut avec un nouveau `-DeviceCode`, que [`start.bat`](scripts/Deployment/readme.fr.md) utilise désormais car une connexion par navigateur peut ne pas s'ouvrir pendant l'OOBE. La version de la galerie (3.9) utilise encore les anciens modules. Plus rien ne charge `WindowsAutopilotIntune`, il quitte donc [`RequiredModules.psd1`](scripts/Startup/RequiredModules.psd1). [`Get-MultiTenantLicenseReport.ps1`](scripts/TenantOnboarding/MultiTenant/readme.fr.md) et `New-CustomerPortalIndex.ps1` exigeaient une application consentie chez chaque client alors que leur documentation disait que GDAP suffisait ; ils fonctionnent maintenant aussi en délégué via GDAP, client par client. `Search-AADDSUserActivity` reste sur Az (Log Analytics n'est pas dans Graph) et reçoit `-TenantId`, `-SubscriptionId` et le code d'appareil |
| **Bugs.** [`New-BreakGlassAdminAccount.ps1`](scripts/TenantOnboarding/Provisioning/readme.fr.md) utilisait `New-MgDirectoryRoleTemplate`, qui ne peut pas activer un rôle ; Global Admin est maintenant une attribution de rôle. [`Set-IntuneBaselinePolicyAssignment.ps1`](scripts/TenantOnboarding/Provisioning/readme.fr.md) appelait `/assign`, qui remplace toute la liste, si bien que chaque stratégie touchée perdait ses autres attributions ; il fusionne maintenant, avec pagination. [`Compare-IntuneConfig.ps1`](scripts/Intune/readme.fr.md) demandait deux étendues de lecture alors que `Start-IntuneBackup` se reconnecte seul sans tenant quand cinq étendues d'écriture manquent : sous GDAP, il sauvegardait votre propre tenant. [`Update-iOSCompliancePolicy.ps1`](scripts/Intune/iOS-Compliance-Updater/readme.fr.md) déclarait `-WhatIf` deux fois et ne démarrait jamais, et lisait un chemin RSS toujours vide ; il tourne maintenant sur le SDK Graph avec un certificat (Setup en crée un ; un secret reste accepté) et la tâche planifiée lance `pwsh.exe`. `New-CustomerPortalIndex` affichait chaque lien en HTML littéral. Les scripts de déploiement Intune peuvent utiliser une application à certificat permanente au lieu d'une temporaire ; la première exécution suivante retéléverse chaque paquet une fois |
| Vérifié : contrôle de syntaxe sur les trois dossiers ; `Get-WindowsAutoPilotInfo.ps1` s'analyse sous Windows PowerShell 5.1 ; contrôle des liens ; le contrôle des modules ; chaque cmdlet utilisée existe (Graph 2.41.1, IntuneWin32App 1.5.0, Az) ; exécutions simulées de l'index du portail, du rapport de licences (délégué et app-only), de la rotation break-glass, de la fusion et de la pagination des attributions de base, de `Compare-IntuneConfig` avec des sauvegardes existantes, et du programme de mise à jour iOS contre le flux Apple réel. Pendant les tests, une exécution a laissé le vrai `Start-IntuneBackup` tenter une connexion interactive, qui a échoué avant toute connexion. **Non** vérifié : tout appel à un tenant - le flux GDAP client par client et ses demandes de consentement, la création du certificat dans Setup, la tâche SYSTEM trouvant un certificat `LocalMachine`, IntuneWin32App avec un certificat, et l'import Autopilot en ligne |

### 2026-10-08 (17)
| Modification |
|--------|
| [`Update-SessionHostImage.ps1`](scripts/RDS/Update-SessionHostImage.ps1) s'exécute seul : si Update-TeamsClient.ps1 et Repair-AppxPackageStore.ps1 ne sont pas à côté (le dépôt, ou le dossier où `-ComputerName` les copie), il les télécharge depuis ce dépôt sur GitHub à un commit épinglé de `main` (`746541e`, 2026-10-05) et les refuse si le SHA-256 ne correspond pas. Auparavant, une copie de ce seul fichier sur la VM d'image sautait toutes les corrections d'applications. Épinglé plutôt que le dernier `main`, pour qu'un script modifié ne s'exécute jamais en tant que System sans avoir été relu ; monter de version demande un nouveau commit et deux empreintes, après lecture du diff |
| **[`Update-SessionHostImage.ps1`](scripts/RDS/Update-SessionHostImage.ps1) maintient les applications à leur build la plus récente au lieu de les bloquer.** Il définissait Microsoft Store `AutoDownload = 2` pour empêcher Outlook de s'écarter, mais le nouvel Outlook se met à jour chaque semaine depuis le CDN Office, pas via le Store, et Microsoft ne documente aucun réglage pour l'en empêcher ([Manage updates in new Outlook](https://learn.microsoft.com/microsoft-365-apps/outlook/manage/manage-updates-new-outlook-windows)) — le réglage ne bloquait que les mises à jour Store des frameworks et de Copilot. Il est désormais affiché, pas défini. La mise à jour automatique de Teams n'est désactivée que si FSLogix est antérieur à 2210 HF4 et rejoue des versions exactes. Chaque exécution sans `-CheckOnly` met désormais à jour Teams, Outlook, Copilot, le complément de réunion et WebView2 vers leur build la plus récente, avec ou sans constat, et WebView2 est en retard dès qu'une version est inférieure à Edge Stable, et non plus seulement d'une version majeure. Les readmes conseillent une exécution hebdomadaire sur chaque hôte |
| Vérifié : contrôle de syntaxe ; les deux empreintes ont été calculées avec `git show` à ce commit et correspondent à ce que raw.githubusercontent.com fournit. **Non** vérifié : une exécution sur un hôte de session qui les télécharge réellement |
| Vérifié : contrôle de syntaxe ; contrôle des liens ; le canal de mise à jour d'Outlook vient de Microsoft Learn (Manage updates in new Outlook for Windows). **Non** vérifié : une exécution sur un hôte de session - le rythme hebdomadaire, Teams désactivé seulement sous 2210 HF4, et la reprise de la build Outlook la plus récente des utilisateurs |

### 2026-10-08 (16)
| Modification |
|--------|
| **[`functies.ps1`](scripts/Startup/readme.fr.md#functiesps1), [`SharePoint/`](scripts/SharePoint/readme.fr.md), [`Teams/`](scripts/Teams/readme.fr.md) et [`Reporting/`](scripts/Reporting/readme.fr.md) se connectent via [`Connect-M365.ps1`](scripts/Startup/readme.fr.md#connect-m365ps1).** Dans `functies.ps1`, les reconnexions par fonction ignoraient le réglage de code d'appareil et passaient `$cid` même en mode Direct ; `Connect-Tenant` et `Test-GdapConnection` cessaient de fonctionner dès qu'une fonction avait basculé la session vers un client (ils trouvent maintenant le contrat dans le tenant partenaire via les nouveaux `Connect-PartnerGraph` / `Get-CustomerContract`), et `Add-TenantDomain`, `Get-TenantLicenses`, `Add-TenantAdmin` et `Get-EntraApplication` ne se connectaient pas du tout et tournaient dans la session ouverte par hasard. [`Search-SharePointContent.ps1`](scripts/SharePoint/readme.fr.md#search-sharepointcontentps1) ne fonctionnait qu'en app-only et est maintenant délégué par défaut (il voit alors ce que votre compte peut atteindre ; `-AppOnly` garde la vue complète). `Find-SiteContent`, `Restore-RecycleBinItems` et l'ensemble de provisionnement utilisent la connexion par appareil quand `load.config.ps1` le demande et lisent le ClientId PnP dans `pnp.appid.json`. [`Trace-SharePointFile.ps1`](scripts/SharePoint/readme.fr.md#trace-sharepointfileps1) passait `-Organization` avec une connexion interactive et lisait le journal d'audit du partenaire sous GDAP. Revoke et le rapport d'autorisations gardent leur application à certificat pour SharePoint REST (qui refuse les jetons Graph délégués et à secret), avec la connexion d'amorçage via le helper et un nouveau `-AppOnly`. Menu : la question du ClientId PnP peut rester vide |
| **[`Invoke-TeamsArchive.ps1`](scripts/Teams/readme.fr.md#invoke-teamsarchiveps1) v9.0 n'a plus besoin du module MicrosoftTeams ni d'une application temporaire.** Équipes, membres et canaux viennent de Graph, et les fichiers sont téléchargés via Graph au lieu de PnP : l'ancienne analyse des chemins ne connaissait que « Shared Documents » et échouait sur un « Gedeelde documenten » néerlandais, et les sous-dossiers étaient aplatis, si bien que deux fichiers de même nom s'écrasaient. Seul le recours administrateur de site (`Set-PnPSite -Owners`) reste PnP. Le redémarrage en session propre transmet aussi les réglages de connexion, et une seconde exécution dans la même fenêtre ne saute plus le nettoyage des modules. [`Get-SharePointStorageReport.ps1`](scripts/Reporting/readme.fr.md#get-sharepointstoragereportps1) parcourait les fichiers via SharePoint REST avec un jeton à secret client que SharePoint refuse ; il utilise maintenant Graph et trouve les bibliothèques masquées via `/lists`. La taille de la corbeille utilise encore ce jeton et reste probablement vide - documenté comme point ouvert |
| Vérifié : contrôle de syntaxe sur les quatre ; contrôle des liens ; `Test-SharePointAccessScripts.ps1` (le bloc commun de Revoke et du rapport d'autorisations est toujours identique) ; `Trace-SharePointFile.ps1` s'analyse sous Windows PowerShell 5.1 ; chaque cmdlet et paramètre PnP 3.4.1 et Graph 2.41.1 utilisé existe ; exécutions simulées du flux de `functies.ps1` (partenaire, client, retour au partenaire, Exchange avec `-DelegatedOrganization`), de `Connect-Structure` en mode strict (appareil, navigateur, `pnp.appid.json`, app-only) et du filtre d'équipes, de la pagination des canaux, du parcours récursif et de la détection des 403 de l'archiveur. Pendant les tests, une simulation a atteint le vrai `Connect-ExchangeOnline` avec un tenant inventé et a été refusée (HTTP 400) ; aucune connexion n'a eu lieu. **Non** vérifié : toute exécution réelle - le téléchargement Graph, le nombre d'invites de connexion par appareil PnP sur de nombreux sites, la couverture de la recherche déléguée, les bibliothèques masquées via `/lists`, le recours `Set-PnPSite` |

### 2026-10-08 (15)
| Modification |
|--------|
| **[`Connect-M365.ps1`](scripts/Startup/readme.fr.md#connect-m365ps1) : recherche d'application sous GDAP, pas d'étendues du partenaire chez un client, et connexion déléguée via votre propre application.** `graph.appid.json` et `pnp.appid.json` sont indexés par le domaine onmicrosoft, alors que sous GDAP le tenant est le GUID du client : `-AppOnly` et la recherche du ClientId PnP échouaient ; ils essaient maintenant aussi le domaine client issu de `Connect-Tenant` et le tenant de l'URL SharePoint. Une reconnexion déléguée gardait les étendues de la session précédente même en changeant de tenant, si bien que des étendues du partenaire comme `Domain.ReadWrite.All` étaient demandées chez le client ; elles ne sont plus reprises qu'au sein du même tenant. Nouveau `-DelegatedClient` : `-ClientId` désigne alors votre propre client public pour une connexion déléguée au lieu de signifier app-only (nécessaire au provisionnement SharePoint) |
| Vérifié : contrôle de syntaxe ; sous `Set-StrictMode` : le GUID GDAP donne la clé de domaine du client, une URL d'administration SharePoint la clé de son tenant, un tenant inconnu rien, et `graph.appid.json` est trouvé à partir d'une URL de site ; avec des cmdlets Graph simulées, un passage du tenant partenaire au tenant client ne demande que les nouvelles étendues, et `-DelegatedClient` transmet le ClientId à une connexion déléguée. **Non** vérifié : sur un tenant |

### 2026-10-08 (14)
| Modification |
|--------|
| **[`Exchange/`](scripts/Exchange/readme.fr.md) se connecte via [`Connect-M365.ps1`](scripts/Startup/readme.fr.md#connect-m365ps1), avec un tableau de connexion par script dans le readme.** Les 17 scripts nécessitent maintenant PowerShell 7 (le helper l'exige). Les scripts Exchange sont délégués par défaut et atteignent un client GDAP avec `-DelegatedOrganization` ; ils passaient `-Organization` avec une connexion interactive et lisaient donc le tenant du partenaire. Ils restent sur Exchange car Graph n'a pas d'API pour ce qu'ils font (transfert, statistiques de boîte, Full Access/SendAs, DKIM, groupes de distribution, rôles de dossier comme `PublishingEditor`) ; chaque section explique pourquoi. Les scripts qui lisent ou écrivent les boîtes et calendriers d'autres utilisateurs (`Get-CalendarMappings`, `Convert-SharedCalendarToResource`, `Move-SharedCalendar`, `Remove-PhishingMessage`, `Restore-MailboxMessages`, `Move-InboxToArchive`) gardent l'app-only par défaut, car un jeton d'administrateur délégué n'atteint pas une autre boîte sans Full Access ; `-AppOnly` lit `graph.appid.json`, et `Remove-PhishingMessage` et `Restore-MailboxMessages` reçoivent un nouveau `-Delegated` pour un administrateur qui a Full Access (refusé sous GDAP). `Move-InboxToArchive -Delegated` demandait `Mail.ReadWrite`, qui ne permet pas d'accéder à la boîte d'un autre ; il demande maintenant `Mail.ReadWrite.Shared` |
| **Corrections.** [`Migrate-Calendar.ps1`](scripts/Exchange/readme.fr.md#migrate-calendarps1) créait une inscription d'application permanente et affichait son secret client à l'écran ; il lit maintenant en délégué et écrit avec une application temporaire dont le secret de 2 heures n'est jamais affiché et qui est supprimée ensuite (ou votre propre application avec `-ClientId`/`-AppOnly`), et `-TenantId`/`-AdminUPN` sont facultatifs. [`Get-MessageTraceReport.ps1`](scripts/Exchange/readme.fr.md#get-messagetracereportps1) se rabattait sur `Get-MessageTrace`/`Get-MessageTraceDetail`, retirés, et n'utilise plus que V2, avec le curseur de pagination en UTC (une heure sans indication était décalée du fuseau local). [`Test-DkimConfig.ps1`](scripts/Exchange/readme.fr.md#test-dkimconfigps1) écrivait dans `$IsWindows`, en lecture seule, et échouait sous PowerShell 7. `Restore-MailboxMessages` et `Remove-PhishingMessage` fermaient la session Graph app-only de l'appelant. [`Set-Calendar-rights.ps1`](scripts/Exchange/readme.fr.md#set-calendar-rightsps1) se connecte lui-même, est en anglais, trouve le calendrier par type de dossier au lieu de trois noms devinés, et modifie une entrée existante au lieu d'échouer ; `Set-Distributionlist-dynamic-static` se connecte lui-même ; `Get-DistributionGroupMembers` n'installe plus ImportExcel de lui-même. Menu : `E` ne propose plus `-Delegated` sous GDAP, `F2` laisse TenantId et AdminUPN facultatifs, `F3` accepte un UPN |
| Vérifié : contrôle de syntaxe ; contrôle des liens ; chaque bloc de paramètres se lie ; chaque cmdlet du SDK Graph et les paramètres de connexion Exchange utilisés (`-DelegatedOrganization`, `-EnableSearchOnlySession`) existent dans Microsoft.Graph 2.41.1 et ExchangeOnlineManagement 3.10.1 ; les nouvelles erreurs de validation (`-Delegated` avec `-AppOnly`, `-Delegated` sous GDAP, `-ClientId` sans secret ni certificat) se déclenchent hors ligne. Pendant les tests, un script s'est connecté par erreur en app-only au tenant du propriétaire et a fait un appel en lecture (`GET /users`, réponse 403) ; rien n'a été écrit. **Non** vérifié : toute connexion Exchange ou déléguée, `-DelegatedOrganization` sous GDAP, la pagination V2 du suivi, le nouveau flux de `Migrate-Calendar`, et le comportement des cmdlets chargées par la session (`Get-MessageTraceV2`, `*-MailboxFolderPermission`) |

### 2026-10-08 (13)
| Modification |
|--------|
| **[`Entra/`](scripts/Entra/readme.fr.md), [`Graph/`](scripts/Graph/readme.fr.md) et [`LegacyUtilities/`](scripts/LegacyUtilities/readme.fr.md) se connectent via [`Connect-M365.ps1`](scripts/Startup/readme.fr.md#connect-m365ps1).** Délégué par défaut avec code d'appareil et client GDAP depuis `load.config.ps1`, `-ClientId`/`-CertificateThumbprint`/`-AppOnly` pour l'app-only, et plus de `Disconnect-MgGraph` inconditionnel qui fermait la session de l'appelant (`Set-UserManager`, `Set-EntraPasskeyMigrationOptOut`, `Copy-GroupMember`). [`Remove-M365Users.ps1`](scripts/Entra/readme.fr.md#remove-m365usersps1) ne se connectait jamais quand aucune session n'existait (`Get-MgContext` ne lève pas d'erreur). [`Phising-rollout.ps1`](scripts/Entra/readme.fr.md#phising-rolloutps1) se connecte maintenant en délégué quand on le lance à la main et utilise toujours automatiquement son identité managée dans Azure Automation ; nouveaux `-UseManagedIdentity` et `-AppOnly`, et sans commutateur il ne tourne plus silencieusement sur la session ouverte par hasard (c'est `-UseExistingSession`). [`logic-permissies.ps1`](scripts/Graph/readme.fr.md#logic-permissiesps1) ne lance plus `Update-Module -Force` par défaut (`load.ps1` tient les modules à jour) et `-TenantId` est facultatif |
| **Graph là où Exchange ou Teams n'était pas nécessaire, et bugs trouvés en route.** [`New-ProjectTeam.ps1`](scripts/LegacyUtilities/Teams/readme.fr.md#new-projectteamps1) ouvrait une session Graph puis faisait tout via le module MicrosoftTeams ; il crée maintenant groupe, équipe et canaux via Graph et attend l'approvisionnement. [`Copy-PlannerPlan.ps1`](scripts/LegacyUtilities/Teams/readme.fr.md#copy-plannerplanps1) ne copiait jamais les listes de contrôle et ne paginait pas les tâches ; [`Copy-Team.ps1`](scripts/LegacyUtilities/Teams/readme.fr.md#copy-teamps1) attendait « un groupe portant le nouveau nom », ce que satisfaisait aussi un groupe existant, et suit maintenant l'opération de clonage. [`New-M365User.ps1`](scripts/Entra/readme.fr.md#new-m365userps1) et `Import-M365Users.ps1` n'envoyaient pas le `mailNickname` exigé par Graph, et le résumé de simulation de l'import indiquait toujours 0. [`Sync-UserContacts.ps1`](scripts/LegacyUtilities/Exchange/readme.fr.md#sync-usercontactsps1) demandait `Contacts.ReadWrite` en délégué, qui ne permet pas d'écrire les contacts d'autres utilisateurs : l'app-only est maintenant le défaut, `-Delegated` uniquement pour votre propre boîte. [`Remove-DuplicateMailItems.ps1`](scripts/LegacyUtilities/Exchange/readme.fr.md#remove-duplicatemailitemsps1) demande `Mail.ReadWrite.Shared` (Full Access requis) et ne s'arrête plus au premier dossier sans sous-dossiers. Les filtres OData échappent `'` (`Set-UserManager`, `Test-M365GroupMembership`, `Add-M365GroupMember`, `Copy-Team`), et les membres sont lus en un seul appel paginé au lieu d'un par utilisateur. Les cinq scripts Exchange de LegacyUtilities restent sur Exchange (Full Access/SendAs, boîtes partagées, contacts de messagerie, autorisations de dossiers et historical search n'ont pas d'API Graph) et atteignent maintenant un client GDAP |
| Vérifié : contrôle de syntaxe sur les trois dossiers ; contrôle des liens ; les 37 cmdlets Mg utilisées existent dans Microsoft.Graph 2.41.1 ; exécutions hors ligne avec des cmdlets Graph simulées de la pagination et de la nouvelle colonne de `Test-M365GroupMembership`, du parcours des sous-dossiers de `Remove-DuplicateMailItems`, de l'échappement de `O'Brien` dans `Set-UserManager`, et de `Phising-rollout` qui choisit le délégué en local et l'identité managée avec `AUTOMATION_ASSET_ACCOUNTID`. **Non** vérifié : sur un tenant - l'approvisionnement d'équipe et le suivi du clonage, la copie des listes de contrôle Planner, `Mail.ReadWrite.Shared` en délégué, l'exigence de `mailNickname` et le runbook dans une vraie Azure Automation |

### 2026-10-08 (12)
| Modification |
|--------|
| **[`Connect-M365.ps1`](scripts/Startup/readme.fr.md#connect-m365ps1) fonctionne sous `Set-StrictMode` et peut forcer une nouvelle connexion.** Il lisait directement `$global:authMode`, `cid`, `connectmsoldomain`, `useDeviceCodeAuth` et `upn` ; un script avec `Set-StrictMode -Version Latest` lancé sans `load.ps1` (tâche planifiée, runbook, nouvelle fenêtre) s'arrêtait sur « variable has not been set ». Les réglages sont maintenant lus via `Get-Variable`, et `Invoke-M365GraphPaged` ne lit plus de `value` ou `@odata.nextLink` absent. Nouveau `Connect-M365Graph -Force` pour se reconnecter quand une session semble correcte mais que son jeton n'est plus valide |
| Vérifié : contrôle de syntaxe ; sous `Set-StrictMode -Version Latest` sans aucun réglage, le choix du tenant, du code d'appareil et du domaine Exchange renvoie vide, et avec GDAP le client ; `Invoke-M365GraphPaged` sous mode strict avec une réponse simulée de deux pages, une vide et un objet seul. **Non** vérifié : sur un tenant |

### 2026-10-08 (11)
| Modification |
|--------|
| **[`PatronToolkit/`](scripts/PatronToolkit/readme.fr.md) se connecte via [`Connect-M365.ps1`](scripts/Startup/readme.fr.md#connect-m365ps1), et le dernier code du SharePoint Online Management Shell disparaît.** Les 13 scripts : délégué par défaut (code d'appareil et client GDAP depuis `load.config.ps1`), nouveaux `-ClientId`, `-CertificateThumbprint` et `-AppOnly`, et ils ne ferment que ce qu'ils ont ouvert. [`Test-SharePointSharingConfig.ps1`](scripts/PatronToolkit/SharePoint/readme.fr.md#test-sharepointsharingconfigps1) utilisait `Connect-SPOService` et ignorait le `-TenantId` documenté ; les paramètres de partage du tenant viennent maintenant de Graph `/admin/sharepoint/settings`, et seul ce que Graph ne couvre pas (liens par défaut avec le nouveau `-IncludeLinkSettings`, dérogations de sites, utilisateurs externes) passe par PnP, ignoré avec un avertissement si PnP ne peut pas se connecter. [`Test-EmailSecurityPosture.ps1`](scripts/PatronToolkit/Security/readme.fr.md#test-emailsecuritypostureps1) appelait `Connect-IPPSSession` sans paramètres et lisait votre propre tenant sous GDAP |
| **Des résultats faux ou vides.** [`Get-IntunePolicyAssignments.ps1`](scripts/PatronToolkit/Intune/readme.fr.md#get-intunepolicyassignmentsps1) et [`Get-AutopilotDevices.ps1`](scripts/PatronToolkit/Intune/readme.fr.md#get-autopilotdevicesps1) appelaient des cmdlets qui n'existent pas dans Microsoft.Graph v2, masquées par `SilentlyContinue` : les attributions Settings Catalog et les profils de déploiement manquaient toujours. Les deux utilisent maintenant Graph avec pagination, et Autopilot ne compte plus chaque appareil comme non attribué (Graph renvoie `assignedInSync`, pas `assigned`). [`Get-MessageTraceReport.ps1`](scripts/PatronToolkit/Exchange/readme.fr.md#get-messagetracereportps1) s'arrêtait à 5000 lignes et utilisait `Get-MessageTraceDetail`, retiré ; il pagine maintenant V2 par tranches de 10 jours jusqu'à 90 jours (nouveau `-MaxResults`). L'analyse DMARC de `Test-EmailAuthenticationRecords` pouvait lire `sp=` comme stratégie ; `Get-SuspiciousInboxRules` qualifiait d'externes les destinataires internes `EX:` ; les rapports d'alertes et de consentements triaient la gravité par ordre alphabétique ; `Get-TeamsConfigReport` lisait une propriété de participation anonyme que sa cmdlet n'a pas, et prend maintenant l'inventaire des équipes dans Graph. `Get-SuspiciousInboxRules` reste volontairement sur Exchange : Graph `messageRules` ne renvoie jamais les règles masquées, que le script recherche justement |
| Vérifié : contrôle de syntaxe ; contrôle des liens ; chaque cmdlet et paramètre Graph 2.41.1 et PnP 3.4.1 utilisé existe, et les trois cmdlets Graph supprimées non ; tests unitaires de l'expression DMARC, du contrôle interne/externe et du tri par gravité ; blocs d'aide alignés sur les paramètres. **Non** vérifié : toute exécution sur un tenant - `$expand=assignments` sur toutes les collections Intune, les valeurs de `/admin/sharepoint/settings`, la pagination V2 au-delà de 5000 lignes, la connexion PnP et les noms de propriétés de la configuration Teams |

### 2026-10-08 (10)
| Modification |
|--------|
| **[`Connect-M365.ps1`](scripts/Startup/readme.fr.md#connect-m365ps1) : sessions Exchange fermées par identifiant de connexion, et ce dont les premiers scripts avaient besoin.** `Disconnect-M365Exchange` lançait `Disconnect-ExchangeOnline` sans `-ConnectionId`, si bien qu'un script qui n'ajoutait qu'une session Security & Compliance - ou un script appelé par un autre - fermait aussi la session Exchange de l'appelant. `Connect-M365Exchange` renvoie maintenant les identifiants des sessions qu'il a ouvertes, et seules celles-ci sont fermées. Une connexion Exchange déléguée ne passe plus `-DelegatedOrganization` que sous GDAP ; hors GDAP, elle connectait un administrateur à son propre tenant comme « partenaire ». Nouveau : `-EnableSearchOnlySession` (purges Content Search), `Disconnect-M365Teams`, `Connect-M365PnP -AppOnly` (l'application à certificat de `graph.appid.json`) et `Invoke-M365GraphPaged` (suit `@odata.nextLink`), que deux scripts avaient chacun écrit pour eux-mêmes |
| Vérifié : contrôle de syntaxe ; avec des cmdlets Exchange simulées : Direct avec `-TenantId` se connecte sans `-DelegatedOrganization`, un second appel réutilise la session, un enfant qui n'ajoute qu'IPPS ne ferme que cette session et celle du parent reste ouverte, GDAP utilise le domaine du client et un autre client ouvre une nouvelle session ; `Invoke-M365GraphPaged` avec une collection simulée de deux pages et une vide. **Non** vérifié : sur un tenant |

### 2026-10-08 (9)
| Modification |
|--------|
| **[`Office365Toolkit/`](scripts/Office365Toolkit/readme.fr.md) se connecte via [`Connect-M365.ps1`](scripts/Startup/readme.fr.md#connect-m365ps1).** Les neuf scripts : délégué par défaut (code d'appareil et client GDAP depuis `load.config.ps1`), nouveaux `-ClientId`, `-CertificateThumbprint` et `-AppOnly` pour l'app-only, et ils ne déconnectent que ce qu'ils ont ouvert. Les quatre scripts Exchange et `Test-SharedMailboxSignIn` passaient `-Organization` avec une connexion interactive et aboutissaient donc, sous GDAP, dans le tenant du partenaire ; ils atteignent maintenant le client avec `-DelegatedOrganization` |
| **Des scripts qui renvoyaient silencieusement trop peu.** [`Get-IntunePolicyInventory.ps1`](scripts/Office365Toolkit/Intune/readme.fr.md#get-intunepolicyinventoryps1) appelait `Get-MgDeviceManagementConfigurationPolicy` et `Get-MgDeviceManagementIntent`, qui n'existent pas dans Microsoft.Graph v2 : Settings Catalog et Endpoint Security manquaient toujours. Il lit maintenant les cinq types de stratégies via Graph, avec pagination. [`Get-SecureScoreReport.ps1`](scripts/Office365Toolkit/Security/readme.fr.md#get-securescorereportps1) affichait toujours un tableau de contrôles vide (il lisait des propriétés typées dans `AdditionalProperties`) ; il relie maintenant chaque contrôle à son profil pour le maximum de points et le titre |
| **Plus de Graph, moins de bugs.** [`Search-MailboxAuditLog.ps1`](scripts/Office365Toolkit/Exchange/readme.fr.md#search-mailboxauditlogps1) utilise maintenant l'API Graph Audit Log Query (créer, attendre, paginer) avec le nouveau `-TimeoutMinutes` ; `-UseExchange` garde `Search-UnifiedAuditLog`, désormais une recherche paginée par type d'enregistrement. [`Remove-EnterpriseAppConsent.ps1`](scripts/Office365Toolkit/Security/readme.fr.md#remove-enterpriseappconsentps1) filtre les autorisations côté serveur au lieu de lire toutes celles du tenant, s'arrête sur une lecture échouée au lieu d'annoncer « aucune autorisation », échappe les apostrophes dans les filtres, affiche les noms de rôles au lieu des GUID et ne demande les étendues d'écriture qu'avec `-Apply`. [`Test-MailboxForwardingRisk.ps1`](scripts/Office365Toolkit/Exchange/readme.fr.md#test-mailboxforwardingriskps1) ne qualifie plus d'externe un destinataire `[EX:/o=...]`. Ce qui reste sur Exchange (compléments, règles de transfert et de balayage, paramètres CAS, stratégies EOP, la liste des boîtes partagées) n'a pas d'API Graph ; chaque readme le dit |
| Vérifié : contrôle de syntaxe ; contrôle des liens ; chaque cmdlet Mg utilisée existe dans Microsoft.Graph 2.41.1 et les deux supprimées non ; exécutions hors ligne avec des appels Graph simulés de la logique Secure Score, de l'inventaire Intune (pagination, nombre d'attributions), de la requête d'audit (corps, deux pages, statut en échec) et du consentement d'applications (filtres, échappement, recherche des rôles) ; le classement des destinataires de transfert testé unitairement. **Non** vérifié : toute exécution sur un tenant - les vraies valeurs de statut et champs d'enregistrement de la requête d'audit, `$expand=assignments` sur `configurationPolicies` en beta, la suffisance des étendues d'aperçu, et le chemin GDAP en pratique |

### 2026-10-08 (9)
| Modification |
|--------|
| **Nouveau [`Update-SessionHostImage.ps1`](scripts/RDS/Update-SessionHostImage.ps1) : une image Windows 11 multisession (ou un hôte AVD) sur laquelle le nouveau Teams, le nouvel Outlook et Copilot continuent de fonctionner avec FSLogix, sans mettre FSLogix à jour.** Trois hôtes mutualisés construits à partir d'une image de 2024 (24H2, 26100) cassaient toujours de la même façon : une application se met à jour par utilisateur sur un hôte, FSLogix rejoue cette version exacte sur un autre hôte qui ne l'a pas, et l'inscription échoue avec `0x80070490`. Repair-AppxPackageStore et Update-TeamsClient réparent les applications, mais rien ne vérifiait ce que l'image fournit autour d'elles, ni n'empêchait l'écart. Le script vérifie l'édition et un redémarrage en attente, la build FSLogix (lecture seule), le blocage des mises à jour du Store et de Teams, WebView2 par rapport à Edge Stable, les builds provisionnées et les utilisateurs qui en ont une plus récente, chaque framework dont dépendent les manifestes des applications, Teams sur AVD (IsWVDEnvironment, la build minimale pour SlimCore, le complément de réunion, le redirecteur WebRTC plus pris en charge depuis le 1er octobre 2026 et supprimé le 1er avril 2027), Shared Computer Activation, le broker de connexion et `redirections.xml` ; `-ForCapture` ajoute la vérification Sysprep. Il corrige lui-même les stratégies, SCA et WebView2, les applications via les deux scripts existants, relit tout et, avec `-ComputerName`, compare le pool. Entrée de menu `J` |
| Vérifié : contrôle de syntaxe ; la recherche d'Edge Stable, la lecture de WebView2 / Edge / du canal Office dans le registre et l'expression régulière des familles de packages ont été exécutées localement sous Windows 11 (pas multisession, pas en mode élevé). **Non** vérifié : une exécution complète, les corrections, `-ComputerName` ou `-ForCapture` sur un vrai hôte de session ou une VM d'image |

### 2026-10-08 (8)
| Modification |
|--------|
| **Nouveau [`Connect-M365.ps1`](scripts/Startup/Connect-M365.ps1) : une seule connexion pour chaque script, Graph d'abord et délégué par défaut.** Un audit de tous les scripts M365 a montré que chacun se connectait à sa façon : seul le démarrage de `functies.ps1` respectait `useDeviceCodeAuth` de `load.config.ps1`, aucun script Exchange n'atteignait un client GDAP (ils passaient `-Organization`, qu'Exchange n'applique qu'en connexion app-only ; un partenaire a besoin de `-DelegatedOrganization`), une douzaine de scripts ne fonctionnaient qu'en app-only et quelque 35 uniquement dans le navigateur. Le nouveau fichier donne à `Connect-M365Graph`, `Connect-M365Exchange` (avec `-IncludeCompliance`), `Connect-M365Teams` et `Connect-M365PnP` les mêmes règles : délégué par défaut, code d'appareil quand `load.config.ps1` le demande, le client GDAP depuis `$global:cid` / `Connect-Tenant`, app-only avec `-ClientId` + `-CertificateThumbprint` ou `-AppOnly` depuis `graph.appid.json` ; une session existante est réutilisée quand elle convient, et `Disconnect-M365Graph` / `Disconnect-M365Exchange` ne ferment que ce que le script a ouvert lui-même. Exchange, Teams et PnP uniquement pour le travail que Graph ne couvre pas. Les scripts y passent dans les commits qui suivent |
| Vérifié : contrôle de syntaxe ; le choix du tenant (aucun réglage, GDAP avec `cid`, `-TenantId` explicite, Direct), le domaine client Exchange sous GDAP, le choix du code d'appareil (`useDeviceCodeAuth`, `-Interactive` qui le remplace) et la recherche dans `graph.appid.json` (par tenant, l'unique entrée, un tenant inconnu avec une erreur claire) exécutés localement ; noms de paramètres vérifiés avec ExchangeOnlineManagement 3.10.1 et Microsoft.Graph.Authentication 2.41.1. **Non** vérifié : une connexion réelle à un tenant |

### 2026-10-08 (7)
| Modification |
|--------|
| **GitHub Actions : [`.github/workflows/ci.yml`](.github/workflows/ci.yml).** Jusqu'ici chaque vérification ne tournait qu'en local, via le hook pre-commit et les hooks Claude Code ; un commit depuis un clone sans `core.hooksPath`, ou depuis l'éditeur web de GitHub, arrivait sur `main` sans contrôle. Le workflow tourne à chaque push et PR vers `devel`/`main` et n'exécute rien - il analyse et vérifie seulement. Quatre jobs : **PowerShell 7** ([`Test-PowerShellSyntax.ps1`](scripts/Startup/Test-PowerShellSyntax.ps1) sur chaque `.ps1`/`.psm1`, et PSScriptAnalyzer au niveau erreur, sans `PSAvoidUsingConvertToSecureStringWithPlainText`, que les scripts break-glass déclenchent volontairement) ; **Windows PowerShell 5.1** (analyse de chaque script sans `#Requires -Version 7`) ; **readmes** ([`Test-MarkdownLinks.ps1`](scripts/Startup/Test-MarkdownLinks.ps1), [`Update-ReadmeHeader.ps1`](scripts/Startup/Update-ReadmeHeader.ps1) - qui échoue sur une langue manquante - et [`Update-ScriptIndex.ps1`](scripts/Startup/Update-ScriptIndex.ps1), suivis de `git diff --exit-code` pour que des fichiers générés périmés échouent) ; **shell** (`bash -n` et ShellCheck au niveau avertissement sur chaque `.sh`). Les constats apparaissent en annotations sur le fichier et la ligne |
| **Le job 5.1 a trouvé 66 scripts qui ne s'analysent pas sous Windows PowerShell 5.1.** Presque tous sont en UTF-8 sans BOM avec un caractère comme `—` ou `é` : 5.1 les lit en ANSI, les octets deviennent un guillemet isolé, et le script casse avec « string is missing the terminator ». PowerShell 7 n'a pas ce problème, d'où le fait que personne ne l'a remarqué - mais Intune, les GPO et les tâches planifiées utilisent 5.1. Quelques autres utilisent `??` ou `?.` sans `#Requires -Version 7` (`Import-M365Users.ps1`, `create_scheduled_task.ps1`, `Get-SharePointStorageReport.ps1`). Le job est donc pour l'instant `continue-on-error` : il signale, il ne bloque pas. Corriger les scripts est un changement séparé |
| Vérifié : sur un checkout propre de `HEAD`, en local, la syntaxe PowerShell 7 (194 fichiers), l'analyseur au niveau erreur (seule la règle exclue), la vérification des liens, la génération des en-têtes et de l'index (aucune différence) et ShellCheck 0.11 sur [`Invoke-LinuxCleanup.sh`](scripts/Linux/Invoke-LinuxCleanup.sh) sont tous propres ; l'analyse 5.1 a été lancée avec `powershell.exe` et donne la liste ci-dessus. Si la première exécution sur GitHub a demandé une correction, elle est dans le commit suivant |

### 2026-10-08 (6)
| Modification |
|--------|
| **Les modules sont installés et mis à jour au démarrage sans rien demander.** [`load.ps1`](load.ps1) exécute maintenant [`Update-Modules.ps1`](scripts/Startup/Update-Modules.ps1) avec le nouveau `-Auto` au lieu de `-Prompt` : ce qui manque est installé, ce qui est obsolète est mis à jour, et seul cela est affiché ; si tout est à jour, une seule ligne, `Modules OK`. `-Prompt` reste pour qui veut qu'on lui demande. La même ligne `-RequiredOnly -Auto -MaxAgeHours 24` est celle qui va dans un profil PowerShell |
| **Nouveau [`Test-RequiredModules.ps1`](scripts/Startup/Test-RequiredModules.ps1), exécuté par le hook de docs après chaque modification.** La vérification au démarrage n'installe que ce que liste [`RequiredModules.psd1`](scripts/Startup/RequiredModules.psd1), et un script qui se mettait à utiliser un nouveau module fonctionnait sur la machine de son auteur et échouait partout ailleurs jusqu'à ce que quelqu'un pense à l'ajouter. Le script parcourt chaque `.ps1`/`.psm1` à la recherche de `#Requires -Modules`, `Import-Module` et `Install-Module` avec un nom littéral et signale chaque nom absent de la liste. [`sync-docs.ps1`](.claude/hooks/sync-docs.ps1) l'exécute après chaque modification d'un `.ps1`, `.psd1` ou `.md` (en réveillant Claude) et avertit avec lui au pre-commit. Les modules volontairement non installés depuis la galerie vont dans la nouvelle section `NotManaged`, avec leur raison : `ActiveDirectory`, `WebAdministration`, `AzureAD`, `Microsoft.Graph` |
| La première exécution de cette vérification a trouvé cinq modules utilisés par des scripts mais absents de toute liste : `Microsoft.Graph.Reports` et `Az.OperationalInsights` (dans des lignes `#Requires`), `DCToolbox`, `IntuneBackupAndRestore` et `Az.Accounts`. Les cinq sont maintenant dans la liste. Touches de menu `U` (Update-Modules, qui n'en avait pas encore) et `Z` (Test-RequiredModules) |
| Vérifié : contrôle de syntaxe ; `Test-RequiredModules.ps1` sur le dépôt est propre sous PowerShell 7.6 et 5.1, et un fichier de test avec une liste `#Requires` comprenant une spécification de module et un `Import-Module -Name '...'` a donné exactement ces trois noms, tandis que `Import-Module $dynamic` était ignoré ; le hook PostEdit a signalé le module manquant pour un fichier de test et est resté silencieux pour le dépôt propre. `-Auto` a tourné pour de vrai sur cette machine : il a installé le `DCToolbox` 2.1.6 manquant sans rien demander (environ 30 s, une fois), et le démarrage suivant n'a affiché que `Modules OK (20 vereist, actueel)` en 2 s démarrage de pwsh compris. **Non** vérifié : la mise à jour d'un module installé pour tous les utilisateurs depuis une session non élevée (une erreur est attendue) |

### 2026-10-08 (5)
| Modification |
|--------|
| **[`Invoke-LinuxCleanup.sh`](scripts/Linux/Invoke-LinuxCleanup.sh) après sa première exécution sur un vrai serveur 3CX.** L'essai à blanc y a révélé trois problèmes. Il affichait `apt/dpkg is running` et ignorait toutes les étapes des paquets, alors qu'`apt-get` fonctionnait normalement : la vérification cherchait un processus nommé `unattended-upgr`, et `unattended-upgrades` en maintient un en permanence. Elle lit désormais le verrou de dpkg lui-même dans `/proc/locks`. Le journal (3,2 Go sur ce serveur) ne figurait pas dans l'estimation ; l'essai à blanc calcule désormais ce que `journalctl --vacuum-time/--vacuum-size` libérera, comme journald en décide : uniquement les fichiers archivés, les plus anciens d'abord, d'après l'heure inscrite dans le nom du fichier. Enfin, une sauvegarde de 1,2 Go datant de 2020 dans `/var/lib/3cxpbx/Data/Backups`, le dossier antérieur à `Instance1`, n'apparaissait pas comme sauvegarde. Ce dossier et son `Logs` sont désormais pris en compte, et chaque sauvegarde est listée avec sa date et sa taille, afin que `--keep-backups` soit un choix éclairé |
| Vérifié : `bash -n` ; dans un conteneur Debian 12 avec un `systemd-journald` en cours d'exécution et des journaux ayant subi une rotation, l'estimation de l'essai à blanc (74,2 Mo) correspondait à ce que `--apply` a libéré (74,2 Mo), et un faux fichier archivé dont le nom le date de 60 jours a été compté puis supprimé ; avec `--keep-backups 3`, les sauvegardes de 2020 (ancien dossier) et celle de la v18 ont été supprimées et les 3 plus récentes conservées ; un processus nommé `unattended-upgr` ne bloque plus les étapes des paquets, tandis qu'un verrou dpkg détenu (`fcntl`) les bloque toujours. **Non** vérifié : `--apply` sur le vrai serveur 3CX |

### 2026-10-08 (4)
| Modification |
|--------------|
| **[`Install-Printer.ps1`](scripts/Device/Printer/Install-Printer.ps1) renforcé pour les exécutions sans surveillance après une golden image.** Un vrai bogue : le contrôle préalable demandait au spouleur si un pilote était installé *avant* de l'attendre, donc au premier démarrage chaque pilote aurait été signalé manquant. Au-delà, une exécution ne peut plus s'arrêter à mi-chemin sur une entrée erronée : tout le JSON est validé d'abord et chaque problème listé d'un coup (champs inconnus comme `adress`, doublons, noms refusés par Windows, adresses, ports, valeurs d'énumération, `"true"` en texte, deux imprimantes sur un port avec des adresses différentes) ; l'INF est lu avant que pnputil le voie (classe imprimante, une section pour cette architecture, le nom exact du pilote parmi ses modèles — en citant les plus proches s'il ne correspond pas — et un catalogue signé pour cette architecture) ; seul le pilote de l'architecture native compte comme installé. Documenté dans le [readme Printer](scripts/Device/Printer/readme.fr.md#install-printerps1) |
| Également : un verrou à l'échelle de la machine pour qu'une tâche de démarrage et un job RMM n'installent pas tous les deux ; le spouleur attendu jusqu'à ce qu'il réponde, et redémarré avec une nouvelle tentative s'il se bloque pendant une installation (volontairement pas après chaque pilote, comme le font certains scripts publiés — cela interrompt l'impression sur un hôte de session en service) ; paramètres d'impression lus et écrits dans un job avec délai ; pages de proxy/connexion et fichiers non zip refusés comme téléchargements, 1 Go libre exigé ; le JSON d'une URL mis en cache comme repli ; `-Proxy` ; fichiers GitHub publics récupérés via le lien de téléchargement de la release et `raw.githubusercontent.com`, hors de la limite API de 60 requêtes par heure qu'un pool derrière un même NAT épuiserait ; pointeurs LFS et arborescences tronquées refusés ; une adresse modifiée déplace ou reconstruit le port ; un seul `Install-Printer.log` renouvelé ; et une exécution sans erreur écrit `ConfigSha256`/`LastSuccess` sous `HKLM:\SOFTWARE\M365-Scripts\InstallPrinter` pour la détection. Comparé aux scripts publiés par MSEndpointMgr, Konrad Brunner (AlyaKoni) et Printune |
| Vérifié sous Windows PowerShell 5.1 et PowerShell 7 : le lecteur d'INF sur six INF d'imprimante de Windows, noms exacts et approchants sur `prnms009.inf`, paquets synthétiques (x86 uniquement, `%token%`, catalogue manquant, programme d'installation seul, pas une imprimante) ; HTML et fichier quelconque enregistrés en `.zip` ; le délai d'attente ; le verrou avec trois processus parallèles (l'un attend et l'obtient, l'autre abandonne après son délai) ; le repli sur le cache après un 404 ; un téléchargement raw pendant que l'API était limitée (cette limite a réellement été atteinte pendant les tests) ; un JSON cassé donnant 19 problèmes ; des exécutions complètes `-CheckOnly`/`-WhatIf`, une faute de frappe, un fichier vide et un proxy invalide. Ces exécutions contournaient la vérification d'élévation, cette session n'étant pas élevée. **Non** vérifié : une exécution élevée qui installe un pilote et une imprimante, la reconstruction d'un port, et un premier démarrage depuis une golden image |

### 2026-10-08 (3)
| Modification |
|--------|
| **[`load.ps1`](load.ps1) vérifie les modules à chaque démarrage : manquants, trop anciens, ou avec une mise à jour.** Auparavant, il regardait seulement si sept modules codés en dur existaient — un module obsolète, ou un module dont un script plus récent avait besoin, passait inaperçu jusqu'à ce qu'une commande échoue. Il exécute maintenant [`Update-Modules.ps1`](scripts/Startup/Update-Modules.ps1) `-RequiredOnly -Prompt`, liste ce qui manque, est sous la version minimale ou en retard sur la PowerShell Gallery, et demande `Deze n module(s) nu installeren/updaten? [J/n]`. Les versions de la galerie sont mises en cache 24 heures (`%LOCALAPPDATA%\M365-Scripts\module-gallery-cache.json`), un démarrage normal coûte donc moins d'une seconde au lieu de six. `-SkipModuleCheck` la saute une fois. Le même appel, `-RequiredOnly -Prompt -MaxAgeHours 24`, peut aller dans un profil PowerShell pour qui démarre via `$PROFILE` plutôt que via `load.ps1` |
| **Une seule liste de modules : nouveau [`RequiredModules.psd1`](scripts/Startup/RequiredModules.psd1).** `load.ps1`, [`Install-Modules.ps1`](scripts/Startup/Install-Modules.ps1) et `Update-Modules.ps1` avaient chacun leur propre liste (7, 14 et 7 modules) et elles ne concordaient pas. Les trois lisent maintenant ce fichier ; y ajouter un module suffit pour que chaque machine se le voie proposer au prochain démarrage. Ajoutés : `PnP.PowerShell` (utilisé par 10 scripts SharePoint/Teams, ignoré sous PowerShell 7.4) et `MicrosoftTeams`, que des scripts utilisaient mais qu'aucune liste n'installait. Retiré : `AzureAD` — Microsoft l'a retiré de la PowerShell Gallery, `Install-Modules.ps1` échouait donc dessus à chaque exécution |
| `Update-Modules.ps1` réécrit : statut par module (`Missing`, `BelowMinimum`, `UpdateAvailable`, `OK`, `Unknown`, `Skipped`), nouveaux `-CheckOnly`, `-RequiredOnly`, `-MaxAgeHours`, `-Scope`, `-Quiet`, `-PassThru`, `-Prompt`. Il lit la version installée avec `Get-Module -ListAvailable` (voit aussi les modules non installés via PowerShellGet), interroge la galerie avec `Find-PSResource` si disponible (3 s au lieu de 9 s avec `Find-Module`), et installe côte à côte là où `Update-Module` refuserait. Sans paramètres, il met ensuite toujours à jour tous les autres modules installés, comme avant |
| Vérifié : contrôle de syntaxe ; sous PowerShell 7.6, `-RequiredOnly -CheckOnly` contre la galerie en ligne signale les 15 modules `OK` sur cette machine (6 s), et depuis le cache en 1,7 s démarrage de pwsh compris (0,8 s dans une session ouverte) ; `-Prompt` affiche une seule ligne `Modules OK`, et sur un hôte non interactif ne modifie rien ; avec une liste de test et un cache modifié, il signale correctement `Missing`, `BelowMinimum`, `UpdateAvailable`, `OK` et `Skipped`, dans l'ordre de la liste, et `-PassThru` renvoie les objets que lit `load.ps1` ; le même test sous Windows PowerShell 5.1 fonctionne et voit les dossiers de modules propres à ce runtime. **Non** vérifié : une véritable exécution d'installation ou de mise à jour (rien n'a été installé sur cette machine), et `load.ps1` en interactif du démarrage au menu |

### 2026-10-08 (2)
| Modification |
|--------|
| **Nouveau : [`Invoke-LinuxCleanup.sh`](scripts/Linux/Invoke-LinuxCleanup.sh) dans le nouveau dossier [`Linux/`](scripts/Linux/readme.fr.md)** — l'équivalent Linux de `Invoke-WindowsCleanup.ps1`, pour un serveur Debian/Ubuntu et en particulier un serveur qui exécute 3CX Phone System. Un script bash, car un tel serveur n'a pas PowerShell. Nettoie le cache APT, `autoremove` (anciens noyaux ; refusé lorsque la liste contient un paquet 3CX, ignoré pendant qu'apt/dpkg s'exécute), la configuration résiduelle des paquets, les révisions snap désactivées, le journal systemd, les journaux ayant subi une rotation, les vidages après plantage, `/tmp`, les caches et corbeilles des utilisateurs, Docker en option (`--docker`, jamais les volumes) et, lorsque 3CX est détecté, les journaux 3CX et les sauvegardes au-delà des N plus récentes (`--keep-backups`). Les enregistrements d'appels, la base de données et la configuration sont seulement signalés. Essai à blanc par défaut, `--apply` pour nettoyer, `--check-only` pour la supervision avec le code de sortie `2` s'il y a du travail |
| Nouveau : [`.gitattributes`](.gitattributes) : `*.sh` est extrait avec des fins de ligne LF, même sous Windows — avec CRLF, bash échoue sur le serveur dès la première ligne |
| Vérifié : `bash -n` et ShellCheck (aucun avertissement) ; dans un conteneur Debian 12 avec une arborescence 3CX reconstituée, des fichiers anciens, 7 sauvegardes, un enregistrement, un candidat à autoremove et un paquet avec configuration résiduelle : l'essai à blanc et `--check-only` ne modifient rien (sortie `0` / `2`), `--apply` supprime exactement les fichiers anciens (journaux actifs, fichiers récents, l'enregistrement, le fichier de verrou de PostgreSQL et `systemd-private-*` restent), conserve les 3 sauvegardes les plus récentes, un second `--apply` ne trouve rien, sans 3CX les sections 3CX sont omises, et les totaux au-delà de 2 Go sont justes (le `mawk` de Debian débordait avec `printf "%d"` à 2 Gio, désormais `%.0f`). **Non** vérifié : un vrai serveur 3CX (les chemins des journaux et sauvegardes proviennent de l'organisation Linux de 3CX), le journal (le conteneur n'a pas systemd), snap et `--docker` |

### 2026-10-08
| Modification |
|--------------|
| **Nouveau : [`Install-Printer.ps1`](scripts/Device/Printer/Install-Printer.ps1) dans le nouveau dossier [`Device/Printer/`](scripts/Device/Printer/readme.fr.md).** Installe des pilotes d'imprimante et des imprimantes TCP/IP décrits dans un seul fichier JSON ([`printers.example.json`](scripts/Device/Printer/printers.example.json)). Les pilotes proviennent d'un asset de release GitHub, d'un dossier d'un dépôt GitHub (via l'API, donc un dépôt privé fonctionne avec `-GitHubToken` / `GITHUB_TOKEN` ; le jeton n'est envoyé qu'aux hôtes de GitHub), d'une URL https quelconque ou d'un partage. Seul un pilote absent, ou plus ancien que la `version` du JSON, est téléchargé ; un `sha256` facultatif et la signature du catalogue sont vérifiés avant `pnputil /add-driver /install` et `Add-PrinterDriver`. Puis port, imprimante, emplacement/commentaire/partage et paramètres d'impression par défaut ; `"ensure": "absent"` supprime une imprimante. Le script existant [`Add-NetworkPrinterConnection.ps1`](scripts/TenantOnboarding/AppDeployment/Add-NetworkPrinterConnection.ps1) n'ajoute une imprimante que pour un pilote déjà présent |
| Conçu pour s'exécuter sans surveillance sur un serveur ou un hôte de session tout juste issu d'une golden image : un spouleur d'impression pas encore démarré est démarré et attendu, les téléchargements qui échouent sur le DNS ou un délai d'attente sont retentés dans la limite de `-WaitSeconds` (un 401/403/404 échoue aussitôt), il se relance en 64 bits car `pnputil` n'existe pas sous SysWOW64, et chaque exécution est idempotente pour pouvoir aussi s'exécuter à chaque démarrage. Pratiques reprises de guides publiés pour imprimantes via Intune/RMM : codes de sortie pnputil `259`/`3010`/`1641` comptés comme réussite avec un renvoi vers `setupapi.dev.log`, SNMP désactivé sur les nouveaux ports sauf `"snmp": true`, `Set-PrintConfiguration` dans un job avec un délai de 2 minutes car certains pilotes universels s'y bloquent. Touche de menu `N` |
| Vérifié : contrôle de syntaxe ; sous Windows PowerShell 5.1 et PowerShell 7, les fonctions seules — un vrai asset de release de `cli/cli` (et le refus quand un motif correspond à 5 assets), un dossier de 7 fichiers de `actions/checkout` et un fichier seul via l'API, un 404 avec l'indication sur le jeton, `sha256` correct et incorrect, recherche d'INF dans un paquet UTF-16 avec dossiers x86/x64, la signature du catalogue du `prnms009.inf` de Windows (Microsoft Print To PDF) et sa version décodée en `10.0.26100.8951`, nouvelle tentative sur une erreur passagère, aucune sur un 404, abandon dans le délai imparti ; le script complet avec `-CheckOnly` (sortie `2`), `-WhatIf`, `-Printer` avec une seule correspondance, `-Quiet` sur un appareil conforme (aucune sortie, sortie `0`), les variables RMM, un JSON incomplet et la relance 32 bits. Lors de ces exécutions, la vérification d'élévation était contournée, car cette session n'était pas élevée. **Non** vérifié : une exécution élevée qui installe réellement un pilote et une imprimante, et un premier démarrage depuis une golden image |

### 2026-10-07
| Modification |
|--------------|
| **[`Update-TeamsClient.ps1`](scripts/Device/Update-TeamsClient.ps1) installe le build auquel il s'est comparé.** La vérification de version demande au service de configuration Teams quel build est actuel, mais l'installation laissait `teamsbootstrapper.exe -p` télécharger ce que son propre déploiement progressif distribuait — et les deux divergeaient pendant des semaines. Une exécution s'est terminée par `Installed build 26246.1604.5133.838 is still older than the published 26260.1701.5139.3736`, trois semaines après la publication de ce build, et la prochaine exécution planifiée aurait de nouveau déclaré l'hôte obsolète et l'aurait réinstallé. Le service de configuration fournit aussi un `buildLink` vers ce MSIX précis, que le script lisait sans jamais l'utiliser. L'étape 5 télécharge désormais ce paquet et en vérifie la taille et la signature Microsoft avant toute désinstallation, et l'étape 7 le provisionne avec `-p -o`. Si le téléchargement ou la vérification échoue, l'exécution avertit et se rabat sur le choix du bootstrapper, puisque rien n'a encore été supprimé |
| Le nouveau `-UseBootstrapperBuild` (NinjaOne : `useBootstrapperBuild`) rétablit l'ancien comportement, où le déploiement progressif de Microsoft décide. Incompatible avec `-UseWinget`. Documenté dans le [readme Device](scripts/Device/readme.fr.md#update-teamsclientps1), [Update-TeamsClient.md](scripts/Device/Update-TeamsClient.md) et la version IT Glue (md + html régénéré) |
| Vérifié : contrôle de syntaxe ; le service de configuration publie `26260.1701.5139.3736` avec un `buildLink` sur `teamsinstaller.public.onecdn.static.microsoft` ; les fonctions du script `Get-LatestTeamsBuild` et `Save-VerifiedDownload`, avec la logique de l'étape 5 sous Windows PowerShell 5.1, ont téléchargé le paquet de 274 Mo et accepté sa signature (`Valid`, `O=Microsoft Corporation`, également sous PowerShell 7) ; un lien répondant 404 a donné l'avertissement de repli sans laisser de fichier ; `-UseWinget -UseBootstrapperBuild` ensemble est refusé avec le code `1` sur les deux environnements. **Non** vérifié : une exécution complète avec droits d'administrateur qui provisionne à partir du paquet téléchargé — aucun hôte n'a encore été mis à jour avec cette version |
### 2026-10-06
| Modification |
|--------------|
| **[`Get-SharePointStorageReport.ps1`](scripts/Reporting/Get-SharePointStorageReport.ps1) ne se contredit plus au sujet de l'historique des versions.** Ignorer l'historique des versions, c'est ce que fait `-SkipVersions` depuis toujours, et `-FastMode` l'implique — mais cette implication n'était appliquée qu'*après* l'affichage de la ligne de mode : une exécution avec `-FastMode` annonçait donc d'abord `Full scan including version history`, puis, deux lignes plus bas, `Fast scan (no version history, no detail rows)`. Sur une exécution de plusieurs heures, c'est la différence entre se fier au résultat ou non. `-FastMode` définit désormais `-SkipVersions` avant la ligne de mode et y dispose de sa propre ligne |
| Vérifié : contrôle de syntaxe ; la logique de la ligne de mode parcourue pour les quatre combinaisons — `-Apply` donne `Full scan including version history`, `-Apply -SkipVersions` donne `Full scan (version history skipped)`, `-Apply -FastMode` et `-Apply -FastMode -SkipVersions` donnent tous deux exactement une ligne, `Fast scan (no version history, no detail rows)`. Rien n'a changé dans ce qui est analysé |

### 2026-10-05 (15)
| Modification |
|--------|
| **[`Get-SharePointStorageReport.ps1`](scripts/Reporting/Get-SharePointStorageReport.ps1) compare désormais aussi les PDF à la limite d'Adobe.** Acrobat et Reader n'ouvrent pas un PDF dont le chemin dépasse 255 caractères — notamment depuis un dossier synchronisé ou réseau, où le correctif d'Adobe de 2021 n'aide pas toujours — de sorte qu'un PDF de 256 à 259 caractères locaux passait le contrôle Windows sans pouvoir être ouvert. Un tel fichier est maintenant signalé `Adobe (255)` dans le CSV des chemins longs, la console et le rapport Markdown. Documenté dans le [readme Reporting](scripts/Reporting/readme.fr.md#chemins-longs-limites-windows) |
| Vérifié : contrôle de syntaxe ; la fonction de mesure exécutée sur des PDF de 255, 256, 259 et 260 caractères locaux — rien, `Adobe (255)`, `Adobe (255)` et `Windows (260)`. Pas exécuté sur un tenant réel |

### 2026-10-05 (14)
| Modification |
|--------|
| **[`Get-SharePointStorageReport.ps1`](scripts/Reporting/Get-SharePointStorageReport.ps1) signale désormais les chemins les plus longs et les compare aux limites de Windows.** Une bibliothèque sans problème dans SharePoint peut quand même échouer une fois synchronisée avec OneDrive : le chemin local `C:\Users\<utilisateur>\<Organisation>\<Site> - <Bibliothèque>\...` est plus long que le chemin SharePoint, et rien dans le rapport ne le montrait. Avec `-Apply` (y compris `-FastMode`), chaque fichier et dossier est maintenant mesuré par rapport aux 400 caractères de SharePoint, au `MAX_PATH` de Windows (260) et, pour les classeurs, aux 218 d'Excel. Résultat dans `SharePoint_LongPaths_<timestamp>.csv`, du plus long au plus court, plus un top 10 dans la console et le rapport Markdown ; il survit à une reprise depuis un checkpoint |
| Le chemin local est mesuré pour le compte membre activé dont l'**UPN est le plus long** dans le tenant, car le dossier de profil porte le nom du préfixe de l'UPN — un chemin qui convient à cet utilisateur convient à tous. Cela ajoute `User.Read.All` à la connexion interactive. Les nouveaux paramètres `-SyncProfilePath`, `-OrganizationName` et `-LongPathThreshold` (défaut 200) remplacent l'estimation. Documenté dans le [readme Reporting](scripts/Reporting/readme.fr.md#chemins-longs-limites-windows) |
| Vérifié : contrôle de syntaxe ; la fonction de mesure exécutée sur des chemins d'exemple — un fichier Excel de 205 caractères locaux reste sous 218, un chemin profond de 282 est signalé `Windows (260)`, un chemin de 441 caractères SharePoint `SharePoint (400)`, et un site personnel OneDrive est mesuré comme `OneDrive - <Organisation>`. Pas encore exécuté sur un tenant réel, donc la recherche de l'UPN et de l'organisation via Graph n'est pas testée |

### 2026-10-05 (13)
| Modification |
|--------|
| **`Remove-CorporateWallpaper.ps1` supprimait aussi l'écran de verrouillage d'entreprise.** Il effaçait toute la clé `PersonalizationCSP` et tout le dossier `C:\ProgramData\Wallpapers`, alors que `Make-lockscreen.ps1` conserve ses valeurs `LockScreen*` dans cette clé et son image dans ce dossier. Il ne supprime désormais que les valeurs `Desktop*` et les fichiers `corporate-background-*`, et la clé ou le dossier seulement s'ils sont ensuite vides |
| Il oubliait aussi la moitié de ce qu'écrit `Set-CorporateWallpaper.ps1` : la stratégie de repli `Policies\System` et toutes les autres ruches utilisateur chargées restaient sur le fond d'entreprise, et en SYSTEM sa réinitialisation de `HKCU` touchait le profil de SYSTEM au lieu de celui d'un utilisateur. Il remet désormais les valeurs de stratégie, chaque ruche utilisateur chargée et le profil Default User — uniquement là où ils pointent vers un fichier d'entreprise — sur l'image Windows par défaut, vide le cache de fond d'écran transcodé de ces utilisateurs, et ne touche `HKCU` que hors exécution en SYSTEM. Journalise dans le dossier de journaux Intune ; `-WhatIf` montre ce qu'il ferait |
| Vérifié : contrôle de syntaxe ; les fonctions testées sur une clé de registre temporaire — une valeur d'entreprise est remise sur `img0.jpg` en style Remplir, un fichier d'écran de verrouillage ou un fond situé ailleurs n'est pas reconnu comme le nôtre, un fond étranger est conservé, et `-WhatIf` ne modifie rien ; et une exécution `-WhatIf` complète sur un poste, qui ignore la ruche Default User avec un message lorsqu'elle n'est pas élevée. Non exécuté en SYSTEM ni via Intune |

### 2026-10-05 (12)
| Modification |
|--------|
| **Le mot de passe `LocalAdmin` de la boîte à outils USB était codé en dur dans `scripts/Deployment/start.bat`**, affiché à l'écran après chaque exécution et écrit en clair dans les readmes anglais — avec l'adresse IP interne du partage d'installation. `start.bat` lit désormais les deux dans `start.local.cmd` à côté de lui (ignoré par git ; modèle `start.local.example.cmd`). Si ce fichier ou une valeur manque, il les demande : le mot de passe en saisie masquée, le chemin du partage au choix de l'option `E`. Sans mot de passe, les options `D`/`E` s'arrêtent au lieu de créer un compte, et le mot de passe n'est plus affiché |
| **Le mot de passe reste dans l'historique git et sur chaque appareil préparé par la boîte à outils — changez-le.** Copiez `start.local.example.cmd` en `start.local.cmd` sur la clé USB avec le nouveau mot de passe et le partage |
| Vérifié avec une copie de `start.bat` où `net`, `wmic` et `reg` affichent seulement ce qu'ils feraient : avec `start.local.cmd`, le compte est créé avec le mot de passe du fichier (y compris `$`, `&` et `!`) et le menu affiche le partage ; sans lui et avec une réponse vide, rien n'est créé et l'option `E` demande le chemin. La saisie masquée elle-même nécessite une console et a été vérifiée séparément (SecureString reconverti en texte dans Windows PowerShell). Non exécuté en OOBE |

### 2026-10-05 (11)
| Modification |
|--------|
| Le readme racine est cliquable partout. L'arborescence **Structure du dépôt** était un bloc de code, aucune de ses 215 entrées n'était cliquable ; c'est désormais un bloc `<pre>` où chaque dossier pointe vers son readme (dans la langue du lecteur) et chaque fichier vers le fichier |
| Les tables du **Menu** relient chaque outil à ce que la touche exécute réellement, d'après `menu.ps1` : 48 lignes, y compris `Test-GroupPermissions` → `Test-DistributionGroupPermissions.ps1` et `Get-DLMembers` → `Get-DistributionGroupMembers.ps1`, dont le nom diffère du libellé ; les fonctions M365 pointent vers la section de `functies.ps1`. Chaque catégorie des **Catégories de scripts** a reçu une ligne de dossier, et les noms de fichiers du texte courant pointent vers le fichier — 384 par langue, uniquement lorsque le nom désigne exactement un fichier suivi |
| Vérifié : chaque lien de l'arborescence pointe vers un fichier ou dossier suivi, le même nombre de liens a été ajouté dans chaque langue, et le contrôle des liens réussit. Deux notes obsolètes de l'arborescence corrigées au passage (Custom Scripts « liés à leur chemin », le module Provisioning partagé « par les quatre ») |

### 2026-10-05 (10)
| Modification |
|--------|
| Chaque readme de dossier suit désormais le modèle de `scripts/Intune/` : chaque script de la table `## Scripts` pointe vers le fichier **et** vers sa section (`([docs](#…))`). 19 dossiers ne le faisaient pas, dans les trois langues : Azure/VM, DNS, Deployment, Device/DriveMapping, Graph et Intune/Desktop/Background/Desktop et /Lockscreen n'avaient aucune table de scripts, et Reporting/Licensing seulement une arborescence ; Device/audio, ClaudeDesktop, CoworkPrerequisites, DiskCleanup, iOS-Compliance-Updater, Get-Autopilot (pas de lien vers `GetAutoPilot.CMD`), UniFi, SAS, SMTP, Startup et une ligne de SharePoint/Provisioning avaient des lignes sans lien docs |
| Lorsqu'un script était documenté sous un titre descriptif, ce titre porte désormais le nom du fichier, pour que l'ancre soit identique dans chaque langue (après avoir vérifié qu'aucun lien ne visait l'ancienne ancre). Les scripts sans section en ont reçu une courte, tirée de leur propre commentaire d'en-tête : `Remove-CorporateWallpaper.ps1`, `Invoke-`/`Detect-DiskCleanupIntune.ps1`, `UnifiApi.ps1`, `Test-SASWorkDirectory.ps1`, `SharePointStructure.Common.ps1`, `Browse-InstallScripts.ps1`, et une section commune pour les scripts install/uninstall/detect de ClaudeDesktop. Les remarques de `Remove-CorporateWallpaper.ps1` notent qu'il n'annule pas tout ce qu'écrit `Set-CorporateWallpaper.ps1` 2.3+ |
| Également corrigé : le readme Intune disait encore que les scripts de thème Office codent leur URL en dur, et les remarques de Provisioning indiquaient que le module partagé est chargé par « les quatre » scripts — les dix scripts du dossier l'utilisent |
| Vérifié : chaque script de chaque readme de dossier a un lien vers le fichier et un lien docs dont l'ancre existe (contrôle scripté sur les 57 readmes de dossier), et le contrôle des liens réussit sur 4562 liens internes. Rien n'a été exécuté |

### 2026-10-05 (9)
| Modification |
|--------|
| `Test-MarkdownLinks.ps1` supprimait chaque trait de soulignement en transformant un titre en ancre, si bien que `## testsmtp_5min.ps1` devenait `#testsmtp5minps1` alors que GitHub en fait `#testsmtp_5minps1`. Un lien correct était signalé comme cassé (et un lien cassé serait passé). Les traits de soulignement ne sont désormais retirés qu'en bord de mot, où ils marquent l'emphase ; à l'intérieur d'un mot ils sont conservés, comme sur GitHub |
| Vérifié : contrôle de syntaxe, et contrôle des liens sur les 204 fichiers markdown, le nouveau lien `[docs]` vers `#testsmtp_5minps1` étant résolu. `testsmtp_5min.ps1` est le seul titre du dépôt avec un trait de soulignement à l'intérieur d'un mot, aucun autre résultat ne change |

### 2026-10-05 (8)
| Modification |
|--------|
| [`Install-SharePointStructure.ps1`](scripts/SharePoint/Provisioning/Install-SharePointStructure.ps1) terminait une construction réussie en demandant de remplir les groupes de sécurité d'un client, par préfixe de nom, quelle que soit la configuration construite. Il indique désormais le nombre de groupes et la configuration dont ils proviennent, lus via `Get-ConfigValue` pour qu'une configuration sans groupes ne bute pas sur le mode strict. L'exemple de sortie de [`Update-TeamsClient.md`](scripts/Device/Update-TeamsClient.md) montrait un vrai compte de domaine ; il affiche désormais `CONTOSO\admin` |
| Vérifié : contrôle de syntaxe, et le message rendu avec la configuration d'exemple (13 groupes) et une configuration sans groupes (0) en mode strict. Une recherche dans tout le dépôt hors historique des versions ne trouve plus aucun nom de client |

### 2026-10-05 (7)
| Modification |
|--------|
| **`Invoke-TeamsArchive.ps1 -DryRun` ne faisait pas de simulation lors d'une première exécution.** Le script supprime les modules Graph en conflit et se relance dans une session `pwsh` propre, mais le redémarrage ne transmettait que le chemin du script — tous les paramètres étaient perdus. La session redémarrée tournait avec les valeurs par défaut : pas de `-DryRun`, pas de `-Step10Only`, pas de `-ChannelAction`, de sorte qu'une exécution prévue comme simulation effectuait le véritable export et l'étape 10 interactive. Seule une exécution dans une session où l'indicateur de redémarrage était déjà défini conservait ses paramètres |
| Le redémarrage transmet désormais chaque paramètre fourni (les switchs uniquement s'ils sont activés, les valeurs telles quelles) et se termine avec le code de sortie de l'exécution redémarrée au lieu de toujours 0 |
| Vérifié avec un script de substitution construit à partir du véritable bloc de paramètres et du code de transmission : `-DryRun -Step10Only -Step10Action undo` et des valeurs contenant des espaces arrivent intacts dans la session redémarrée, les valeurs par défaut restent par défaut, et le code de sortie de l'enfant est renvoyé. L'archiveur lui-même n'a pas été exécuté |

### 2026-10-05 (6)
| Modification |
|--------|
| `scripts/Teams/vias_archiver.ps1` renommé en [`Invoke-TeamsArchive.ps1`](scripts/Teams/Invoke-TeamsArchive.ps1), et débarrassé d'un client : les titres de l'assistant, les invites pour le tenant et le compte admin, l'URL SharePoint d'exemple, le nom de l'application temporaire, les fichiers temporaires, le nom du dossier eDiscovery et le nom du rapport citaient tous ce client, et le fichier Excel et le dossier d'archive par défaut pointaient vers son propre fichier et son lecteur réseau. Les valeurs par défaut sont désormais `C:\Temp\Teams_Channels.xlsx` et `C:\Temp\Teams_Archive` |
| La feuille Excel était lue via un nom de feuille codé en dur au nom du client. Nouveau `-WorksheetName` ; sans lui, la première feuille est lue. La variable d'environnement de redémarrage est renommée avec le script |
| Le readme Teams a reçu la table `## Scripts` qui lui manquait, le nouveau paramètre et les colonnes requises dans le fichier Excel. Absent de [`menu.ps1`](menu.ps1), donc rien à renommer là |
| Vérifié : contrôle de syntaxe et recherche des noms de client restants. Non exécuté sur un tenant |

### 2026-10-05 (5)
| Modification |
|--------|
| [`SharePoint/Provisioning/`](scripts/SharePoint/Provisioning/readme.fr.md) se disait neutre vis-à-vis des clients, mais livrait la configuration complète d'un client (tenant, propriétaire, URL de sites, groupes), un guide utilisateur écrit pour un autre, et cinq scripts dont le `-ConfigPath` par défaut était un fichier de configuration au nom d'un client, inexistant — sans `-ConfigPath`, ils échouaient donc sur un fichier manquant |
| La configuration client est remplacée par [`example.config.json`](scripts/SharePoint/Provisioning/example.config.json) : le même modèle avec des noms Contoso et `CHANGEME` dans le tenant, le propriétaire et les URL des sites. Les configurations client (`<client>.config.json`) restent à côté mais sont ignorées par git : une configuration existante continue de fonctionner localement et n'est plus publiée |
| Nouvelle fonction `Resolve-StructureConfigPath` dans [`SharePointStructure.Common.ps1`](scripts/SharePoint/Provisioning/SharePointStructure.Common.ps1) : sans `-ConfigPath`, les cinq scripts prennent désormais l'unique `*.config.json` qui ne contient plus `CHANGEME` — la règle que suivaient déjà [`Install-SharePointStructure.ps1`](scripts/SharePoint/Provisioning/Install-SharePointStructure.ps1), [`Remove-SharePointStructure.ps1`](scripts/SharePoint/Provisioning/Remove-SharePointStructure.ps1) et [`Sync-SharePointChannelMember.ps1`](scripts/SharePoint/Provisioning/Sync-SharePointChannelMember.ps1). [`menu.ps1`](menu.ps1) l'indique dans sa question au lieu de nommer un fichier |
| Le guide utilisateur renommé en [`SharePoint-Handleiding.md`](scripts/SharePoint/Provisioning/SharePoint-Handleiding.md) et transformé en gabarit (Contoso NV, marques Northwind et Fabrikam, avec une note invitant à les remplacer). [`New-StructureConfig.ps1`](scripts/SharePoint/Provisioning/New-StructureConfig.ps1) ne propose plus le nom et les marques d'un client comme valeurs par défaut ; exemples et readmes utilisent Contoso. Les noms internes des colonnes (`PsMerk`, …) sont inchangés — ils figurent dans chaque configuration et dans les sites déjà construits |
| Vérifié : contrôle de syntaxe du dossier et de [`menu.ps1`](menu.ps1) ; l'exemple se charge sans erreur une fois `CHANGEME` renseigné et est refusé tel que livré ; le résolveur choisit l'unique configuration renseignée et ignore l'exemple. Aucune exécution sur un tenant |

### 2026-10-05 (4)
| Modification |
|--------|
| [`Deploy-OfficeTheme.ps1`](scripts/Custom%20Scripts/Intune/Desktop/Deploy-OfficeTheme.ps1) et [`Deploy-Officecolors.ps1`](scripts/Custom%20Scripts/Intune/Desktop/Office%20Themes/Deploy-Officecolors.ps1) avaient été conçus pour un seul client : l'URL de téléchargement et le nom de fichier du thème étaient codés en dur et pointaient vers le `.thmx` et le XML de couleurs de ce client dans un dépôt GitHub. Ils reçoivent désormais `-ThemeUrl`/`-ThemeName` et `-ColorsUrl`/`-ColorsName` ; le nom est par défaut le dernier segment de l'URL et doit se terminer par `.thmx` ou `.xml`. Sans URL, ils s'arrêtent avec le code de sortie 1 au lieu de demander, car personne ne répond sous Intune |
| Suppression des fichiers de thème du client (`.thmx` et `.xml` du jeu de couleurs) du dépôt. Les readmes ne disent plus que les scripts sont liés à ce chemin, et expliquent comment passer des paramètres sous Intune (commande d'installation d'une application Win32, ou copie avec valeurs par défaut renseignées) |
| **Les déploiements Intune existants contiennent leur propre copie de l'ancien script et continuent de télécharger depuis l'URL qu'elle contient** — cette URL pointe vers un autre dépôt GitHub, et si celui-ci est synchronisé depuis ce dépôt, ces déploiements perdent le fichier dès que cette modification atteint `main` |
| Vérifié : contrôle de syntaxe, les chemins sans URL et avec mauvaise extension s'arrêtent avec le code 1 et leur message, et une URL encodée en pourcentage donne le nom de fichier attendu. Aucun téléchargement ni déploiement Intune n'a été exécuté |

### 2026-10-05 (3)
| Modification |
|--------|
| [`Init-TempDisk.ps1`](scripts/Device/TempDisk/Init-TempDisk.ps1) mémorisait son dernier redémarrage déclenché sous une clé de registre au nom d'une entreprise. La clé est désormais le paramètre `-RestartMarkerPath`, par défaut `HKLM:\SOFTWARE\M365-Scripts\InitTempDisk`, validé comme chemin `HKLM:` ou `HKCU:` |
| **Sur les machines où il a déjà tourné, l'ancienne clé n'est plus lue** : le délai de carence repart donc une fois de zéro — au plus un redémarrage supplémentaire par machine, et seulement si toutes les autres conditions sont réunies. Passez l'ancienne clé via `-RestartMarkerPath` pour la conserver |
| Vérifié : contrôle de syntaxe, et le motif du paramètre accepte la valeur par défaut et refuse un chemin de fichier. Non exécuté sur un appareil |

### 2026-10-05 (2)
| Modification |
|--------|
| Remplacement des données client dans les exemples par des valeurs Contoso : [`Set-UserManager.ps1`](scripts/Entra/Set-UserManager.ps1) (un domaine de messagerie client), [`Import-DnsRecords.ps1`](scripts/DNS/Import-DnsRecords.ps1) (zone DNS et DC d’un client), [`Get-ComputerLastLogon.ps1`](scripts/Reporting/Get-ComputerLastLogon.ps1) (le chemin d’OU d’un client) et [`Get-SharePointStorageReport.ps1`](scripts/Reporting/Get-SharePointStorageReport.ps1), dont l'aide citait la boîte aux lettres d'une personne réelle |
| Le rapport de licences avait un chemin OneDrive d'entreprise (`C:\OneDrive\<Company>\...`) codé en dur dans [`genereer_rapport.ps1`](scripts/Reporting/Licensing/genereer_rapport.ps1) et [`genereer_licentie_overzicht.py`](scripts/Reporting/Licensing/genereer_licentie_overzicht.py), et le readme demandait de modifier les scripts. Le dossier provient désormais de `-ExportDir` / `--export-dir`, sinon de la variable d'environnement `LICENSING_EXPORT_DIR` ; sans l'un ni l'autre, les deux s'arrêtent avec le code 2 au lieu de deviner. Le lanceur le vérifie avant `Join-Path`, qui échouerait sinon sur un chemin vide |
| [`create_scheduled_task.ps1`](scripts/Reporting/Licensing/create_scheduled_task.ps1) lisait ses réglages dans des variables à modifier, dont un compte de service fixe. `-ExportDir` et `-RunAsUser` sont désormais des paramètres obligatoires, `-RunDay`/`-RunTime` facultatifs, et la tâche transmet `--export-dir` à Python. `-RunTime` était en outre ignoré pour calculer la première exécution (toujours 08:00) ; il est maintenant utilisé. **Une tâche enregistrée auparavant s'exécute sans `--export-dir` et s'arrête désormais immédiatement — réenregistrez-la** |
| Vérifié : contrôle de syntaxe du PowerShell modifié, `py_compile` sur le moteur Python, et le lanceur sans dossier d'export s'arrête avec le code 2 et le message. Le moteur Python lui-même n'a pas été exécuté (pas de `pandas` sur cette machine) et la tâche planifiée n'a pas été réenregistrée |

### 2026-10-05
| Modification |
|--------|
| Suppression de `scripts/djm` et `scripts/djm.pub` — une clé privée OpenSSH (`djm-portaal`) et sa moitié publique, commitées dans le dépôt. Rien dans le dépôt ne les utilisait. [`.gitignore`](.gitignore) exclut désormais `id_*`, `*.pem`, `*.key` et `*.pub`. **La clé reste dans l'historique git et doit être considérée comme compromise : faites-la tourner sur chaque hôte qui lui fait confiance** |

### 2026-10-02 (5)
| Modification |
|--------------|
| Une exécution réelle de [`Revoke-SharePointUserAccess.ps1`](scripts/SharePoint/Revoke-SharePointUserAccess.ps1) indiquait `Sites with access: 15` et `Grants found: 0`, sans aucune erreur. Le rapport de permissions recensait 30 attributions pour le même utilisateur sur les mêmes sites : les deux se contredisaient, et cest la révocation qui avait tort |
| **Les attributions de rôle étaient lues à la mauvaise URL.** `Get-ScopeRoleAssignments` recevait la base de portée (`/_api/web`) et linterrogeait telle quelle au lieu de `/_api/web/roleassignments`. SharePoint répondait par lobjet web, dépourvu de tableau `value` : lassistant de pagination navait rien à parcourir et renvoyait une collection vide. Rien néchouait, rien nétait trouvé. La fonction ajoute désormais `/roleassignments` elle-même, ce qui correspond aussi à la base de lURL de suppression |
| **`\24384` est une variable automatique en lecture seule.** `\24384 = [int]\.PrincipalId` lève `Cannot overwrite variable PID`. Masquée par le bug dURL, elle nest jamais apparue, mais aurait transformé chaque évaluation de portée en exception rattrapée. Renommée, et un contrôle parcourt désormais les deux scripts à la recherche daffectations à une automatique en lecture seule |
| Ajout du garde-fou général pour la moitié silencieuse : un point de terminaison de collection répond toujours par `value` (ou `d.results`), même vide. Une réponse sans lun ni lautre est donc la mauvaise URL, pas un résultat vide. Lassistant de pagination lève désormais au lieu de ne rien renvoyer, dans les deux scripts |
| Vérifié par 126 contrôles (6 nouveaux) et une reproduction fondée sur les formes réellement renvoyées par ce locataire : une attribution directe `Beperkte toegang` est désormais trouvée, tout comme les liens de partage, les groupes Entra rejoints et `Everyone`, tandis quune autre personne et un groupe non rejoint ne correspondent pas. **La lecture corrigée nest pas encore vérifiée sur un locataire réel** |

### 2026-10-02 (4)
| Modification |
|--------------|
| Les deux scripts daccès SharePoint sont désormais couplés par la sortie même du rapport. [`Revoke-SharePointUserAccess.ps1`](scripts/SharePoint/Revoke-SharePointUserAccess.ps1) reçoit `-FromReport` : il prend les sites à visiter dans une exécution de [`Get-SharePointPermissionsReport.ps1`](scripts/Reporting/Get-SharePointPermissionsReport.ps1) au lieu de reparcourir le locataire. Le rapport dit qui peut atteindre quoi, vous le lisez et décidez, et la révocation agit exactement sur ce que vous regardiez — sur un locataire où un balayage complet prend un quart dheure, un utilisateur ayant accès à quelques sites est révoqué en quelques secondes |
| Il lit de préférence le fichier daccès par site plutôt que la liste brute des attributions. Ce fichier a déjà résolu chaque groupe en personnes : un site nest visité que si lutilisateur appartient réellement au groupe qui accorde laccès. La première version utilisait la liste des attributions, qui nomme le groupe mais pas ses membres — sur un locataire où la plupart des sites accordent via `Site Members`, cela revenait à visiter presque tous les sites, annulant tout lintérêt. Un test prouve désormais que la voie privilégiée visite moins de sites que le repli |
| Le rapport décide où chercher, jamais quoi supprimer : chaque site nommé est tout de même lu en direct, donc une attribution disparue entre-temps revient en `AlreadyGone` plutôt quen échec, et une suppression manuelle nest pas annulée. Linverse est signalé plutôt que supposé — tout ce qui a été accordé après le rapport, et tout ce que le rapport na pas pu lire, figure dans le résumé, et un rapport de plus dun jour le signale |
| `-FromReport` accepte le CSV de détail, un fichier frère de la même exécution, ou le dossier ; sans `-TenantUrl`, le locataire est aussi tiré du rapport, car répéter une URL déjà contenue dans le fichier est une façon de se tromper |
| Vérifié par 120 contrôles (22 nouveaux), dont 20 exécutent les vraies fonctions sur de vrais fichiers de rapport : résolution du CSV de détail depuis un dossier, un fichier frère ou le classeur ; les voies privilégiée et de repli ; un invité retrouvé par son adresse ou son UPN de locataire ; les lignes dune autre personne ignorées ; une ligne derreur najoutant aucun site ; et une sous-chaîne dun UPN réel ne correspondant à rien. **Pas encore vérifié sur un locataire réel** |

### 2026-10-02 (3)
| Modification |
|--------------|
| `-RemoveFromEntraGroups` ne pouvait agir que sur les groupes trouvés par le scan de l''exécution elle-même, et rien ne le disait. Une exécution sur un seul site, un `-Scope` restreint, OneDrive ou les listes masquées exclues, ou des portées illisibles réduisent cette liste — et « tous les groupes accordant l''accès ont été retirés » se lit alors comme complet alors que ce ne l''est pas. C''est ainsi qu''un départ est validé à moitié |
| L''exécution détermine désormais ce qu''elle n''a **pas** couvert et le dit deux fois : avant tout retrait, puis de nouveau dans le résumé, en nommant chaque limite. Un groupe accordant l''accès à un endroit jamais fouillé est explicitement signalé comme absent de la liste |
| À dire clairement, car la question était légitime : le script de révocation effectue son propre scan. [`Get-SharePointPermissionsReport.ps1`](scripts/Reporting/Get-SharePointPermissionsReport.ps1) n''est pas un prérequis — découverte, révocation et phase Entra se déroulent dans une seule exécution, dans cet ordre |
| `$scanLimits` est déclaré à côté de `$stats` plutôt que dans le scan, afin qu''une exécution interrompue tôt laisse au résumé une liste vide au lieu d''une variable inexistante |
| Vérifié par 83 contrôles (7 nouveaux) : chaque limite est collectée, l''avertissement apparaît avant les retraits et de nouveau à la fin, et la liste survit à une sortie précoce |

### 2026-10-02 (2)
| Modification |
|--------------|
| Ajout de `-RemoveFromEntraGroups` à [`scripts/SharePoint/Revoke-SharePointUserAccess.ps1`](scripts/SharePoint/Revoke-SharePointUserAccess.ps1), qui termine la seconde moitié d'un départ au lieu de se contenter de la signaler. Jusqu'ici le script retirait chaque attribution SharePoint puis vous renvoyait traiter les groupes Entra vous-même |
| **Seuls les groupes que cette exécution a réellement vus détenir une attribution de rôle sur une portée dans le périmètre sont touchés** — jamais tous les groupes auxquels l'utilisateur appartient. Un partant peut être dans cinquante groupes, et élargir reviendrait à confondre révoquer un accès et détacher quelqu'un de l'organisation |
| Un groupe Entra n'est pas un objet SharePoint : la même appartenance porte souvent une équipe Teams, une boîte aux lettres, des licences et des attributions d'applications que ce rapport ne voit pas. Le commutateur est désactivé par défaut, la bannière et le résumé disent ce qu'il atteint, et l'entrée de menu répond non par défaut |
| Quatre cas sont signalés plutôt que forcés, car les forcer échouerait ou ferait la mauvaise chose : un groupe dynamique (l'appartenance suit une règle, rien n'est stocké à retirer), un groupe synchronisé depuis AD on-premises (en lecture seule dans le cloud), une appartenance héritée d'un groupe imbriqué (l'utilisateur n'est pas membre direct, la coupure doit se faire au groupe qui le contient), et un utilisateur introuvable dans Entra |
| La permission d'écriture suit le commutateur : `GroupMember.ReadWrite.All` n'est demandée que si `-RemoveFromEntraGroups` est fourni, de sorte qu'une exécution en lecture seule ne détient rien qui puisse modifier une appartenance à l'échelle du locataire. Les retraits passent par le même entonnoir que toute autre modification : `-Apply`, `-WhatIf`, la confirmation et le CSV d'audit se comportent à l'identique |
| Vérifié par 76 contrôles (11 nouveaux) : le commutateur existe et conditionne le rôle d'écriture, la phase Entra parcourt les groupes vus accordant l'accès et jamais `$userGroupIds`, les quatre refus sont présents, le retrait passe par l'entonnoir, et la passe par portée se contente toujours d'enregistrer une attribution Entra sans agir dessus |

### 2026-10-02
| Modification |
|--------|
| Ajout de [`scripts/RDS/Invoke-FSLogixShrink.ps1`](scripts/RDS/Invoke-FSLogixShrink.ps1) : les fichiers VHDX dynamiques de profil/ODFC FSLogix grandissent mais ne rendent jamais d'espace, et la réduction se faisait à la main à partir de commandes collées qui téléchargeaient ce qui se trouvait à ce moment sur la branche master d'Invoke-FslShrinkDisk, écrivaient le journal dans un `C:\Temp` qui n'existe peut-être pas (son `Export-Csv` échoue alors) et nommaient le partage d'un client. Le script télécharge Invoke-FslShrinkDisk à un commit épinglé (`bfe0504`, 2025-06-19) et le refuse si le SHA-256 ne correspond pas, liste chaque conteneur du partage du plus grand au plus petit (`-ReportOnly`), réduit avec les mêmes valeurs par défaut (≥ 5 Go, ≥ 10 % libre, 4 à la fois), crée le dossier du journal, et résume les Go récupérés et les disques qui n'ont pas pu être traités — généralement attachés parce que l'utilisateur est connecté |
| Recherche sur GitHub d'une meilleure solution : Invoke-FslShrinkDisk est toujours maintenu par l'équipe FSLogix et reste l'outil ; les forks et ShrinkVHD font la même chose avec moins de garanties. La vraie amélioration est la VHD Disk Compaction propre à FSLogix à chaque déconnexion (2210 et ultérieur, activée par défaut) ; `-CheckHost` indique donc si elle peut s'exécuter sur un hôte : version, `VHDCompactDisk`, `defragsvc` non désactivé, disques dynamiques. S'il réussit, une réduction manuelle ne fait que rattraper |
| Ajout de l'entrée de menu `K` (FSLogix-Shrink), et de [`Get-FSlogix-errors.ps1`](scripts/RDS/Get-FSlogix-errors.ps1) dans la catégorie RDS et l'arborescence du readme racine, où il manquait |
| Vérifié dans Windows PowerShell 5.1 sur un poste de travail : le téléchargement du commit épinglé et le contrôle du hachage, la deuxième exécution qui le réutilise, une copie altérée refusée, `-ReportOnly` sur un dossier de test avec deux fichiers VHDX (6 Go et 1 Go, plus un fichier non VHD ignoré), `-CheckHost` signalant que FSLogix n'est pas installé (code 1), et le résumé CSV sur un journal d'exemple (4,75 Go récupérés, un disque en cours d'utilisation nommé, code 1). La réduction réelle n'a pas été exécutée sur un partage — cela demande une session élevée sur un hôte ayant accès au partage de profils |

### 2026-10-01 (4)
| Modification |
|--------|
| Pour un paquet qui continue d'échouer alors que la bonne build est provisionnée, [`Repair-AppxPackageStore.ps1`](scripts/Device/Repair-AppxPackageStore.ps1) répond maintenant à la question qui détermine si cela compte : chaque utilisateur connecté a-t-il l'application ? lem-avd-4 avait Outlook 915 provisionné et FSLogix journalisait encore cet après-midi-là `Deployment Register ... from:  (AppxManifest.xml) failed with error 0x80070490` — FSLogix enregistrant avec un chemin vide. L'erreur ne permettait pas de savoir si des utilisateurs étaient privés d'Outlook ou si seul le journal se remplissait |
| L'exécution compare les ruches utilisateur chargées avec les utilisateurs pour lesquels le paquet est enregistré (Installed), et nomme la build la plus récente de chacun. Tous couverts : les échecs sont la propre relecture de FSLogix, l'exécution le dit, indique `InstallAppxPackages = 0` comme méthode documentée par Microsoft pour le faire taire sans le modifier, et se termine par 0. S'il manque à quelqu'un : la personne est nommée et l'exécution échoue |
| Vérifié de bout en bout dans Windows PowerShell 5.1 avec un état AppX simulé : lem-avd-4 avec Outlook enregistré pour l'utilisateur connecté → « all 1 signed-in user(s) have it (1.2026.915.300) », code 0 ; lem-avd-5 avec l'application enregistrée pour quelqu'un d'autre → l'utilisateur est nommé, code 2 ; un hôte vide → code 0. Pas encore exécuté sur les hôtes |

### 2026-10-01 (3)
| Modification |
|--------|
| [`Repair-AppxPackageStore.ps1`](scripts/Device/Repair-AppxPackageStore.ps1) s'est arrêté lors de la première exécution réelle sur le pool avec `The property 'Name' cannot be found on this object` (lem-avd-4). `@($exactTargets.Name)` lève une erreur sous `Set-StrictMode` dans Windows PowerShell 5.1 lorsque la liste est vide — ce qui est le cas sur un hôte qui a déjà la build la plus récente. Désormais construit à partir des éléments eux-mêmes |
| Trouvé grâce à l'exécution de bout en bout qui aurait dû exister plus tôt : avec `-Name teams,outlook -Provision`, les programmes d'installation de Microsoft s'exécutaient aussi pour des paquets déjà provisionnés, et ils livrent une build plus ancienne (« last known good ») — Outlook 818 par-dessus une 915 provisionnée, une rétrogradation. Les programmes d'installation ne s'exécutent plus que pour un paquet qui n'est pas provisionné du tout ; les builds plus récentes passent par la voie build exacte / `-Latest` |
| Vérifié en exécutant le script **entier** dans Windows PowerShell 5.1 avec les cmdlets AppX, les journaux d'événements, les téléchargements et les signatures simulés, dans le scénario de lem-avd-4 (915 provisionnée, FSLogix échouant sur 902/915), celui de lem-avd-5 (profils demandant 922) et un hôte vide, avec `-Name teams,outlook -Latest -Provision -RemoveOld` et avec `-CheckOnly` : aucun arrêt, aucun programme d'installation par-dessus un paquet provisionné, 922 provisionnée exactement dans le scénario lem-avd-5, codes de sortie 0/1/2 comme attendu. Pas encore réexécuté sur les hôtes |

### 2026-10-01 (2)
| Modification |
|--------|
| `Repair-AppxPackageStore.ps1 -RemoveOld` supprime toute référence qu'un hôte garde à une build plus ancienne des paquets nommés, une fois la plus récente provisionnée : copies provisionnées plus anciennes, builds plus anciennes enregistrées pour un utilisateur (pour tous les utilisateurs, ou par utilisateur quand cela échoue), et ce que `AppxAllUserStore` en retient encore dans les entrées utilisateur, fin de vie, suppression différée et machine — chaque clé sauvegardée en `.reg` d'abord. Sur l'hôte qui continuait d'échouer, Outlook 818 était encore provisionné à côté de 915 et 902 encore enregistré, ce qui laissait d'anciennes builds à portée d'une ouverture de session |
| Volontairement limité : uniquement les paquets nommés un par un (`-Name teams,outlook` ; ignoré avec un caractère générique), et rien du tout si la build conservée n'est pas provisionnée, afin qu'aucun utilisateur ne se retrouve sans l'application. Les dossiers `WindowsApps` sont laissés à Windows, qui en est propriétaire et les supprime dès que plus rien n'y fait référence, et la liste de chaque conteneur de profil à FSLogix, qui la réécrit à la prochaine déconnexion — les deux sont signalés. Le menu `R` le propose après le provisionnement, avec `-Latest` |
| Vérifié dans PowerShell 5.1 avec les cmdlets simulées et un `AppxAllUserStore` de test : en conservant 915, la 818 provisionnée, la 902 enregistrée et les trois entrées du magasin pour 902/818 ont été supprimées (trois sauvegardes `.reg`), 915 est restée partout et les deux anciens dossiers `WindowsApps` ont été nommés ; sans build provisionnée, rien n'a été supprimé. Pas exécuté sur un hôte de session |

### 2026-10-01
| Modification |
|--------|
| `Repair-AppxPackageStore.ps1 -Latest` provisionne la build Teams / Outlook la plus récente qui existe, pas seulement celle sur laquelle FSLogix a échoué. Teams provient du service de configuration de Microsoft, le flux qu'utilise le client lui-même (26246 aujourd'hui, avec son lien MSIX). Outlook n'a pas de tel flux — le catalogue du Store indiquait 1.2026.818.0 alors que 915.300 était déjà sur le CDN et dans les profils des utilisateurs — on utilise donc la build la plus récente dont l'existence est prouvée (demandée par FSLogix, enregistrée pour un utilisateur sur l'hôte, ou présente dans `WindowsApps`), et l'exécution indique la source utilisée |
| L'exécution n'affiche plus « Nothing to repair » lorsqu'un paquet continue d'échouer alors que la bonne build est provisionnée. Une exécution réelle après le correctif de build exacte montrait Outlook 1.2026.915.300 provisionné, FSLogix 26.01, des profils demandant 902 et 915 — et FSLogix échouant encore le même après-midi, plus 55× `0x80073CF9`. L'étape 1b indique désormais qu'il ne s'agit pas d'un écart de version (là où elle supposait « une ancienne version enregistrée qui disparaît à la prochaine déconnexion »), l'exécution se termine par 1, et l'étape 1c affiche les preuves : la dernière erreur de déploiement AppX avec le texte d'erreur spécifique de Windows et son `Get-AppPackageLog -ActivityID`, ainsi que les lignes du paquet dans le journal de profil de FSLogix |
| Vérifié dans PowerShell 5.1 : `-Latest` contre le **vrai** service de configuration Teams (26246 trouvée plus récente qu'une 26225 provisionnée, avec la bonne URL MSIX) et avec Outlook déduit de ce qui a été vu (915 au-dessus de 818) ; les preuves contre le **vrai** journal AppX de cette machine, qui a affiché le texte d'erreur spécifique et un ActivityId. Pas encore exécuté sur les hôtes de session |

### 2026-09-30 (12)
| Modification |
|--------------|
| Correction de `Request_UnsupportedQuery: Unsupported or invalid query filter clause specified for property appId` au démarrage, introduite par le commit précédent. En sortant la liste des rôles du bloc partagé vers `\`, elle sest retrouvée **au-dessus** des identifiants dapplication dont elle dépend : chaque `ResourceAppId` était vide et le filtre devenait `appId eq ` |
| Les constantes figurent désormais au-dessus de la liste des rôles dans chaque script plutôt que dans le bloc partagé, ce qui est leur place dès lors que la liste les utilise |
| Lanalyseur ne détecte pas une variable utilisée avant son affectation : le test vérifie donc lordre directement — les identifiants avant la liste des rôles, la liste avant le bloc partagé. Également prouvé en exécutant le prologue de chaque script et en confirmant que chaque rôle correspond à un vrai GUID |

### 2026-09-30 (11)
| Modification |
|--------------|
| [`scripts/SharePoint/Revoke-SharePointUserAccess.ps1`](scripts/SharePoint/Revoke-SharePointUserAccess.ps1) signalait un compte réel et actif comme `Not found in Entra ID`. L''application temporaire n''avait jamais reçu Graph `User.Read.All` — `GroupMember.Read.All` ne permet pas de lire un objet utilisateur quelconque — donc `GET /users/{upn}` renvoyait `403`, et le bloc catch l''interprétait comme un utilisateur inexistant |
| La même erreur qu''auparavant, à un nouvel endroit : se voir refuser une lecture n''est pas le même fait que l''absence de la chose, et un seul des deux peut être ignoré. Le script s''arrête désormais en nommant la permission manquante. Cela va au-delà du message : sans utilisateur résolu, les groupes Entra qui accordent aussi l''accès ne sont jamais listés, et c''est précisément la moitié du rapport qui dit ce que ce script **ne peut pas** révoquer |
| `User.Read.All` ajouté aux rôles demandés par le script de révocation. Pour garder le rapport au privilège minimal, la liste des rôles n''est plus codée en dur dans le bloc partagé : chaque script définit `$RequiredAppRoles` avant, et la vérification des rôles du jeton valide ce que ce script a demandé plutôt qu''une paire fixe. Le rapport ne demande toujours pas `User.Read.All`, car il développe des groupes et ne lit jamais un objet utilisateur |
| Vérifié par 63 contrôles (10 nouveaux) : le script de révocation demande `User.Read.All` et le rapport non, les deux demandent toujours les trois rôles communs, la vérification du jeton suit la liste propre au script, un `403` sur la recherche est fatal et nomme la permission, et une absence réelle ne fait toujours qu''avertir |

### 2026-09-30 (10)
| Modification |
|--------------|
| Une exécution sur tout le locataire de [`scripts/Reporting/Get-SharePointPermissionsReport.ps1`](scripts/Reporting/Get-SharePointPermissionsReport.ps1) a montré que l''échelle de sélection redécouvrait la même réponse sur chaque site : trois listes système (`Galerie van thema''s`, `Galerie met basispagina''s`, `Bibliotheek met onderhoudslogboeken`) refusaient les mêmes champs sur les 131 sites, chacune au prix d''un aller-retour inutile et d''une ligne de journal |
| Les champs qu''une liste accepte relèvent de son **modèle**, pas du site : le résultat est désormais appris une fois par modèle puis réutilisé. Une simulation du schéma observé sur 131 sites donne la moitié des allers-retours (1048 à 528), chaque modèle atterrissant toujours exactement sur le barreau qu''il accepte — aucun champ n''est perdu par ce raccourci |
| Le journal le signale une fois par modèle au lieu d''une fois par site — y compris le `[SKIP]` de la User Information List, présent sur les 131. Environ 350 lignes répétées en moins, précisément ce qui masquait le reste |
| Vérifié en simulant l''échelle sur 131 sites avec et sans mémoire : moins d''appels, résolution identique par modèle, et un modèle qui n''accepte rien se termine toujours au lieu de boucler |

### 2026-09-30 (9)
| Modification |
|--------------|
| Correction de [`scripts/Reporting/Get-SharePointPermissionsReport.ps1`](scripts/Reporting/Get-SharePointPermissionsReport.ps1) qui échouait sur **chaque** site avec `Cannot validate argument on parameter 'Kind'. The argument "A" does not belong to the set "U,G"`. L'ajout de la vue d'accès consolidée a appris au point de contrôle à *lire* un troisième type de clé (`A`) sans jamais élargir le `ValidateSet` de la fonction qui en *écrit* une : le premier web levait une erreur et les 131 sites signalaient un échec |
| Introduit en même temps que l'onglet `Toegang` et non détecté parce que la suite de tests du rapport vivait dans un dossier temporaire de session, vidé entre deux sessions — exactement le coût signalé à ce moment-là |
| Un contrôle a été ajouté pour toute la classe plutôt que pour ce seul cas : chaque type écrit doit figurer dans le `ValidateSet` **et** être relu par le commutateur de reprise, et chaque type autorisé doit réellement être utilisé. Vérifié qu'il se déclenche en l'exécutant sur les deux variantes cassées — le type absent de l'ensemble, et un type écrit mais jamais lu, qui ne lève rien mais perd silencieusement son état de reprise |
| Aucune donnée perdue. Les webs en échec ont écrit des lignes d'erreur portant leur `UnitKey`, et la logique de remplacement les supprime dès que le web réussit : une simple ré-exécution se nettoie elle-même |


### 2026-09-30 (8)
| Modification |
|--------|
| Passe de durcissement sur [`scripts/SharePoint/Revoke-SharePointUserAccess.ps1`](scripts/SharePoint/Revoke-SharePointUserAccess.ps1), guidée par une lecture du code à la recherche des modes de défaillance propres à un script destructif plutôt que de ceux d'un rapport. Trois étaient réels |
| **La correspondance d'un invité aurait pu révoquer la mauvaise personne.** La recherche de repli comparait avec `-like "*needle*"`, et `an@contoso.com` est une sous-chaîne de `jan@contoso.com`. Remplacée par une comparaison exacte avec l'UPN, l'adresse mail, le suffixe de claim et une connexion invité correctement décodée (`jan_partner.com#ext#@tenant` redevient `jan@partner.com`, en coupant sur le dernier underscore afin qu'une partie locale puisse en contenir un). Lorsque deux comptes différents répondent à la même adresse, le site reste intact et l'exécution s'arrête en nommant les deux — choisir revient à l'opérateur, pas au script |
| **Le CSV d'audit était écrit une seule fois, à la fin.** Une exécution qui aurait révoqué deux cents éléments avant de planter n'aurait laissé aucune trace de ce qu'elle avait supprimé, ce qui est précisément la seule chose qu'un tel script ne doit jamais faire. Les lignes sont désormais ajoutées au fil de l'eau, via le helper partagé qui réessaie sur un fichier verrouillé et arrête l'exécution plutôt que de perdre une ligne |
| **Un `404` lors d'une suppression comptait comme un échec.** Il signifie que l'accès a déjà disparu, ce qui est le résultat normal lors d'un second passage — une ré-exécution propre aurait signalé des échecs. Enregistré désormais comme `AlreadyGone` |
| L'échec de la suppression d'un administrateur de collection de sites est désormais bien visible et compte comme un échec : ce rôle couvre chaque périmètre du site, donc toute autre suppression y est cosmétique tant qu'il subsiste. Le résumé le dit explicitement plutôt que de ressembler à un succès |
| `-WhatIf` emprunte désormais la même branche qu'un essai à blanc, et enregistre donc `WouldRevoke` au lieu de `Skipped`, qui laissait entendre que quelqu'un avait refusé une demande de confirmation |
| [`Test-SharePointAccessScripts.ps1`](scripts/SharePoint/Test-SharePointAccessScripts.ps1) est passé de 27 à 50 contrôles : décodage des connexions invité, y compris une partie locale contenant un underscore, correspondance exacte face aux quasi-correspondances qu'un test de sous-chaîne aurait acceptées (plus court, plus long, domaine suffixé, invité d'un autre tenant, vide), existence du CSV d'audit contenant chaque ligne en cours d'exécution, et un 404 interprété comme déjà disparu. **Toujours non vérifié sur un tenant réel** |

### 2026-09-30 (7)
| Modification |
|--------|
| Suppression d'un paramètre `-Restart` mort dans [`scripts/SharePoint/Revoke-SharePointUserAccess.ps1`](scripts/SharePoint/Revoke-SharePointUserAccess.ps1). Il était déclaré et son aide promettait qu'il allait `discard any existing checkpoint and start over instead of resuming` - mais le script n'a ni checkpoint ni reprise, donc le switch ne faisait rien et l'aide décrivait un comportement inexistant. Trouvé en comparant le bloc de paramètres avec l'aide basée sur les commentaires et le readme du dossier, plutôt qu'en supposant qu'ils concordaient |
| Remplacé par une entrée `.NOTES` expliquant pourquoi il n'y a délibérément pas de reprise : la révocation est idempotente, donc une seconde exécution ne trouve que ce que la première n'a pas supprimé. Relancer après une interruption constitue à la fois la récupération et la vérification, et c'est plus sûr que de reprendre une opération destructive partiellement appliquée à partir d'une position enregistrée |
| Vérifié que les 17 paramètres restants figurent dans l'aide basée sur les commentaires et dans le readme du dossier, que `Get-Help` ne mentionne plus `-Restart`, et que les 27 contrôles de [`Test-SharePointAccessScripts.ps1`](scripts/SharePoint/Test-SharePointAccessScripts.ps1) passent toujours |

### 2026-09-30 (6)
| Modification |
|--------|
| Nouveau [`scripts/SharePoint/Trace-SharePointFile.ps1`](scripts/SharePoint/Trace-SharePointFile.ps1) : retrouve où est passé un fichier OneDrive ou SharePoint — renommé, déplacé, copié, supprimé, restauré — par qui et quand, depuis le Unified Audit Log. Auparavant il fallait parcourir à la main la recherche d'audit de Purview, où une chaîne de renommages ou un dossier parent renommé passe facilement inaperçu |
| La trace est suivie par ID d'élément et par le chemin vers lequel un fichier a été renommé ou déplacé, de sorte que `A → B → C` aboutit à C. Les renommages, déplacements et suppressions de dossiers sont rejoués sur le chemin du fichier, car SharePoint n'écrit pas d'enregistrement par fichier pour ceux-ci |
| Les heures sont affichées en heure de Bruxelles (`Europe/Brussels`, avec le décalage UTC, heure d'été/d'hiver prise en compte) et la période s'indique en heure locale de Bruxelles, jour d'abord ; une date de fin sans heure inclut toute la journée |
| Robustesse : la période est lue par jour, une tranche de plus de 50 000 enregistrements est scindée (jusqu'à 15 minutes), une recherche en échec ou incohérente (`ResultIndex -1`) est relancée avec une attente croissante, et les enregistrements en double sont écartés. [`menu.ps1`](menu.ps1) le propose sous la touche `O` |
| Vérifié : contrôle de syntaxe ; fonctionne dans PowerShell 7 et Windows PowerShell 5.1 contre un `Search-UnifiedAuditLog` **simulé** — chaîne de renommages, copie signalée mais non suivie, renommage et mise à la corbeille d'un dossier rejoués sur le fichier, ancienne URL et lien de partage `/:w:/r/`, `-SiteUrl` qui n'inclut pas un site voisin au même préfixe, le passage à l'heure d'été le 29 mars 2026 (+01:00 → +02:00), et la scission à 50 000 enregistrements. **Pas encore exécuté contre un tenant réel** ; la structure des champs d'audit (`SourceRelativeUrl`, `DestinationFileName`, `ListItemUniqueId`) suit le schéma documenté par Microsoft |

### 2026-09-30 (5)
| Modification |
|--------|
| La documentation se tient désormais à jour d'elle-même. [`.claude/hooks/sync-docs.ps1`](.claude/hooks/sync-docs.ps1) régénère [`scripts/INDEX.md`](scripts/INDEX.md) et les en-têtes des readmes et lance la vérification des liens ; il est appelé par un hook git pre-commit ([`.githooks/pre-commit`](.githooks/pre-commit), pour les modifications faites à la main) et par des hooks Claude Code dans [`.claude/settings.json`](.claude/settings.json) (après chaque modification, en arrière-plan). Auparavant il fallait penser aux trois — et [`INDEX.md`](scripts/INDEX.md) avait déjà deux scripts de retard |
| Avant que Claude ne termine, un hook Stop vérifie qu'une modification d'un readme anglais a aussi été faite en néerlandais et en français, et qu'un script modifié a vu le readme de son dossier mis à jour ; le hook git avertit pour le premier cas, car un hook ne sait pas traduire. Un lien cassé arrête le commit |
| Vérifié : chaque mode exécuté sur ce dépôt — une modification propre reste silencieuse, un lien cassé injecté sort avec le code 2 en indiquant fichier et cible, un readme non traduit et un script non documenté bloquent chacun le Stop une fois (et pas une seconde), et le hook Claude a été vu se déclencher après une modification. Le hook pre-commit s'est exécuté sur ce commit |

### 2026-09-30 (4)
| Modification |
|--------|
| Chaque readme existe désormais en trois langues : [`readme.md`](readme.md) (anglais, toujours la version principale), [`readme.nl.md`](readme.nl.md) (néerlandais) et [`readme.fr.md`](readme.fr.md) (français) — 66 dossiers, le readme racine y compris tout l'Historique des versions. Le sélecteur en tête de chaque page mène à la même page dans l'autre langue, et le fil d'Ariane reste dans la langue que vous lisez |
| Les titres qui sont un nom de script (`### Set-UserManager.ps1`) ne sont pas traduits, si bien que chaque ancre `#…ps1` est identique dans les trois langues ; les autres titres le sont, avec leurs liens internes adaptés. Les noms de paramètres, commandes, chemins et les chaînes littérales qu'un script affiche ou écrit (noms d'onglets Excel néerlandais, messages d'erreur) restent tels quels dans chaque langue |
| `Reporting/readme.md` et une partie de `SharePoint/readme.md` étaient en néerlandais dans un ensemble anglais ; ils ont d'abord été passés en anglais, et les versions néerlandaises reprennent la formulation d'origine |
| Le mot de passe de l'administrateur local qui figure en clair dans [`scripts/Deployment/readme.md`](scripts/Deployment/readme.md) et dans cet historique n'est **pas** recopié dans les versions néerlandaise et française ; il y apparaît comme omis |
| [`.claude/CLAUDE.md`](.claude/CLAUDE.md) exige désormais qu'une modification de readme soit faite dans les trois langues, suivie de [`Update-ReadmeHeader.ps1`](scripts/Startup/Update-ReadmeHeader.ps1) |
| Vérifié avec [`Test-MarkdownLinks.ps1`](scripts/Startup/Test-MarkdownLinks.ps1) : 204 fichiers markdown, tous les liens internes aboutissent ; `Update-ReadmeHeader.ps1 -Check` signale tous les en-têtes à jour. Les traductions ont été contrôlées sur la structure (sections, tableaux, nombre de lignes par rapport à l'anglais), pas relues ligne à ligne par un locuteur natif |

### 2026-09-30 (3)
| Modification |
|--------|
| Le nouveau [`scripts/Startup/Update-ReadmeHeader.ps1`](scripts/Startup/Update-ReadmeHeader.ps1) écrit les deux lignes en tête de chaque readme : un sélecteur de langue (`English · Nederlands · Français`) et le fil d'Ariane pour remonter l'arborescence, chaque niveau pointant vers son readme dans la langue courante. Tenus à la main, ce sont ces chemins relatifs qui cassent quand un dossier est déplacé ; générés à partir du dossier du readme, ils ne le peuvent pas. `-Check` sort avec le code 1 si un en-tête est obsolète ou si une version linguistique manque |
| [`scripts/INDEX.md`](scripts/INDEX.md) régénéré : outre le nouveau script, il liste désormais [`Revoke-SharePointUserAccess.ps1`](scripts/SharePoint/Revoke-SharePointUserAccess.ps1) et [`Test-SharePointAccessScripts.ps1`](scripts/SharePoint/Test-SharePointAccessScripts.ps1), ajoutés sans relancer [`Update-ScriptIndex.ps1`](scripts/Startup/Update-ScriptIndex.ps1) |
| Vérifié : contrôle de syntaxe sans erreur ; exécuté sous PowerShell 7 et un passage `-Check` sous Windows PowerShell 5.1 sur ce dépôt — 66 dossiers, 198 readmes, tous les en-têtes à jour ensuite |

### 2026-09-30 (2)
| Modification |
|--------|
| Chaque readme commence désormais par un fil d'Ariane (`M365-Scripts › scripts › Intune › Desktop`) qui renvoie à chaque niveau supérieur. Auparavant, 40 des 66 readmes de dossier n'offraient aucun moyen de remonter au dossier parent, hormis le bouton Précédent du navigateur |
| Le readme racine s'ouvre sur un tableau `## Folders` qui renvoie vers chaque dossier de workload, de sorte que le dépôt se parcourt depuis la page d'accueil vers le bas, et plus uniquement via [`scripts/readme.md`](scripts/readme.md) |
| Les sous-dossiers sont listés partout sous un titre `## Folders`. [`Device/`](scripts/Device/readme.fr.md), [`Network/`](scripts/Network/readme.fr.md), [`Reporting/`](scripts/Reporting/readme.fr.md) et [`SharePoint/`](scripts/SharePoint/readme.fr.md) les mélangeaient au tableau Scripts, [`Intune/Desktop/`](scripts/Intune/Desktop/readme.fr.md) et [`Custom Scripts/Intune/Desktop/`](scripts/Custom%20Scripts/Intune/Desktop/readme.fr.md) utilisaient un tableau `Contents`, [`TenantOnboarding/`](scripts/TenantOnboarding/readme.fr.md) indiquait `Subfolders` et [`LegacyUtilities/`](scripts/LegacyUtilities/readme.fr.md) n'avait aucun titre |
| Vérifié avec [`Test-MarkdownLinks.ps1`](scripts/Startup/Test-MarkdownLinks.ps1) : les 1 072 liens internes répartis sur 72 fichiers markdown se résolvent. Documentation uniquement ; aucun script modifié |

### 2026-09-29 (11)
| Modification |
|--------|
| [`Repair-AppxPackageStore.ps1`](scripts/Device/Repair-AppxPackageStore.ps1) provisionne le build **exact** de Teams / Outlook sur lequel FSLogix échoue. Un hôte de production montrait Outlook en échec 196× en une semaine — dont 186× `0x80070490` — pour 1.2026.902 et 915, alors que l'hôte provisionnait 818 et que les fichiers des deux builds demandés étaient sur le disque. Le verdict précédent (« avec un FSLogix à jour, l'écart est sans conséquence, aucune action nécessaire ») était faux, tout comme le fait de le combler avec les installateurs, qui ne livrent qu'un build last-known-good plus ancien |
| Le MSIX d'un build précis se trouve sur le CDN de Microsoft à l'URL versionnée qu'utilisent les manifestes de winget (`res.cdn.office.net/.../v2/<version>/Microsoft.OutlookForWindows_x64.msix`, `teamsinstaller.public.onecdn.static.microsoft/production-windows-x64/<version>/MSTeams-x64.msix`) — vérifié comme répondant pour Outlook 812/818/902/915 et Teams 26198/26225/26246. `-Provision` prend désormais le build le plus récent sur lequel FSLogix a échoué, le télécharge, vérifie la signature Microsoft, le provisionne et relit la version provisionnée ; l'installateur est ignoré pour ce package |
| Les codes d'erreur sont décodés avec le message propre de Windows pour chaque code Win32, au lieu d'une courte liste écrite à la main, un code par ligne : `0x80073D19` s'est révélé être « An error occurred because a user was logged off » — sans conséquence — et est désormais étiqueté comme tel |
| Vérifié sous PowerShell 5.1 : le scénario de cet hôte (FSLogix demandant 902 et 915, hôte à 818) donne 915 avec la bonne URL, un **vrai** téléchargement de ce MSIX de 32 Mo avec une signature Microsoft valide, la version provisionnée relue (avec `Add-AppxProvisionedPackage` simulé), et aucune cible une fois que l'hôte a 915 ; scénarios précédents inchangés. **Non exécuté sur l'hôte lui-même** |

### 2026-09-29 (10)
| Modification |
|--------|
| `Repair-AppxPackageStore.ps1 -Copilot -Provision` installe désormais la **nouvelle** application Microsoft Copilot unifiée au lieu de l'ancienne application Microsoft 365 Copilot. L'installateur ajouté en (9) livre l'ancien package AppX, que l'unification doit ensuite faire migrer ; la nouvelle application est installée à l'échelle de la machine par Edge Update. La méthode documentée est utilisée : `Install{C50565E9-...}` = 5 (Force Installs), `UpdaterExperimentationAndConfigurationServiceControl` = 1 (requis par Force Installs) et `CopilotUnificationAllowed{...}` = 1 sous `HKLM\SOFTWARE\Policies\Microsoft\EdgeUpdate`, écrits après une sauvegarde `.reg` de cette clé ; ensuite la tâche machine d'Edge Update est lancée et l'exécution attend jusqu'à 10 minutes que l'application apparaisse sous `EdgeUpdate\Clients`. Si elle n'apparaît pas, l'ancien installateur sert de solution de repli. `Install` = 0 n'est jamais écrasé. Le diagnostic et la vérification ne considèrent désormais Copilot comme présent que lorsque la nouvelle application l'est |
| `-Name` comprend `teams`, `outlook` et `copilot`, de sorte que le nouvel Outlook et la nouvelle application Copilot sur le pool s'obtiennent avec `-ComputerName lem-avd-4,lem-avd-5,lem-avd-6 -Name outlook,copilot -Provision`. Le menu (`R`) accepte les mêmes mots |
| Trouvé pendant les tests, corrigé : sous Windows PowerShell 5.1 avec `ErrorActionPreference = Stop`, un `reg.exe` qui écrit sur stderr est une erreur bloquante, si bien qu'une seule sauvegarde `.reg` en échec aurait mis fin à toute l'exécution au lieu de simplement laisser cette clé intacte. La sauvegarde accepte aussi désormais les chemins `HKLM:`/`HKCU:` |
| Vérifié sous PowerShell 5.1 et 7 : les raccourcis (cinq combinaisons), l'installation via Edge Update contre une clé de stratégie et une tâche simulées — valeurs écrites, sauvegarde effectuée, « apparition » de l'application détectée — et `Install` = 0 laissé intact ; le scénario précédent du magasin inchangé. **Non exécuté sur un hôte de session** : on n'a pas testé si Edge Update y installe l'application dans les 10 minutes |

### 2026-09-29 (9)
| Modification |
|--------|
| `Repair-AppxPackageStore.ps1 -Copilot` diagnostique et rétablit Copilot. L'étape 1d indique si l'application Microsoft 365 Copilot (`Microsoft.MicrosoftOfficeHub`) et l'application Windows Copilot (`Microsoft.Copilot`) sont enregistrées et provisionnées, la présence de l'application Microsoft Copilot unifiée qu'Edge Update installe depuis l'unification de septembre 2026 (lue depuis `EdgeUpdate\Clients\{C50565E9-...}`), la version d'Edge Update par rapport au 1.3.253.25 dont elle a besoin, et chaque stratégie qui tient Copilot à l'écart — `Install` / `Uninstall` / `Update{C50565E9-...}` sous `Policies\Microsoft\EdgeUpdate` (Force Installs l'emportant sur Uninstall, comme documenté), la mise en pause de l'unification, et les stratégies Windows `WindowsCopilot` / `WindowsAI` à l'échelle de la machine et par utilisateur connecté. Les stratégies sont signalées avec leur chemin et leur valeur et font échouer l'exécution, mais ne sont jamais modifiées : elles proviennent de GPO ou d'Intune |
| Avec `-Provision`, Copilot est installé pour tous les utilisateurs avec la commande documentée par Microsoft `M365CopilotDesktopInstaller.exe --quiet --start -p` depuis `go.microsoft.com/fwlink/?linkid=2325486` — vérifié aujourd'hui comme livrant un `xpdBootstrapper` 16.0.19305 signé par Microsoft — et considéré comme réussi lorsque le package AppX est provisionné ou que l'application unifiée apparaît sous Edge Update. `-Copilot` inclut aussi dans le périmètre les marqueurs Deprovisioned des deux packages, ce que laisse derrière lui un outil de debloat. winget ne propose qu'un `.exe` pour cette application, donc `-UseWinget` se rabat sur l'installateur. Le tableau du pool a gagné une colonne Copilot ; le menu (`R`) accepte `copilot` comme réponse pour le package |
| L'entrée précédente a été renumérotée en (8) : elle et l'entrée SharePoint ci-dessous avaient toutes deux été commitées sous (7) le même après-midi |
| Vérifié sous PowerShell 5.1 et 7 : l'étape 1d contre le registre et les packages réels de cette machine, et contre des stratégies Edge Update simulées (Uninstall seul fait échouer l'exécution, Uninstall avec Force Installs non) ; le tableau du pool de l'orchestrateur avec la nouvelle colonne ; le téléchargement de l'installateur et sa signature. **Non exécuté sur un hôte de session** : le provisionnement `-p` de l'installateur et l'apparition ultérieure de l'application unifiée n'ont pas été testés |

### 2026-09-29 (8)
| Modification |
|--------|
| [`Repair-AppxPackageStore.ps1`](scripts/Device/Repair-AppxPackageStore.ps1) s'exécute sur tout un pool avec `-ComputerName lem-avd-4,lem-avd-5,lem-avd-6` (éventuellement `-Credential`) : il se copie dans `C:\IT\AppxRepair` sur chaque hôte via PowerShell remoting, s'y exécute avec les mêmes paramètres — la sortie propre de l'hôte est renvoyée en direct — et se termine par un tableau unique couvrant le pool (code de sortie, build FSLogix, Teams / Outlook provisionnés) qui signale toute différence entre les hôtes. Une réparation est confirmée une seule fois pour tout le pool, car une session distante ne peut pas répondre de manière fiable à une demande de confirmation. Intégré au menu (la touche `R` demande les hôtes) |
| Le conseil de l'étape de vérification était faux. Après une exécution réelle de `-Provision`, il indiquait que Teams (26225) et Outlook (1.2026.818) étaient « toujours plus anciens que les 26246 / 902 demandés par les profils - amenez les autres hôtes au même build ». Mais aucun hôte n'est en avance : les deux applications se mettent à jour elles-mêmes par utilisateur, et les installateurs de Microsoft provisionnent un build last-known-good qui est en retard sur celui-ci, de sorte que le profil sera toujours en avance sur chaque hôte et que provisionner plus récent ne tient que jusqu'à la prochaine mise à jour. Ce qui détermine si cela pose problème, c'est FSLogix : à partir de 2210 HF4 (Teams) / 25.06 (Outlook), il enregistre par nom de famille et l'écart est sans conséquence (désormais signalé comme OK, code de sortie 0) ; sur un build plus ancien, le conseil est de mettre à jour FSLogix. Un tel écart ne compte plus comme quelque chose à provisionner |
| Une transcription qui refuse de démarrer — comme dans certaines sessions distantes et RMM — n'interrompt plus la réparation ; c'est un avertissement |
| Vérifié sous PowerShell 5.1 et 7 : l'orchestrateur contre deux hôtes injoignables (chacun nommé avec son erreur WinRM, le tableau du pool, code de sortie 1), et l'écart de version avec un FSLogix simulé au-dessus et en dessous du minimum. **Non exécuté contre de vrais hôtes de session** : pas de WinRM vers lem-avd-4/5/6 depuis ici, donc la copie, l'exécution distante et le tableau du pool avec des valeurs réelles n'ont pas été testés |

### 2026-09-29 (7)
| Modification |
|--------|
| Ajout de [`scripts/SharePoint/Revoke-SharePointUserAccess.ps1`](scripts/SharePoint/Revoke-SharePointUserAccess.ps1) — le pendant du rapport d'autorisations. Il trouve chaque endroit où un utilisateur nommé dispose d'un accès et le supprime : d'abord administrateur de collection de sites (ce rôle l'emporte sur tout ce qui se trouve en dessous, le laisser rendrait le reste cosmétique), puis les attributions de rôle directes sur les sites, sous-sites, listes, dossiers et fichiers individuels, l'appartenance aux groupes SharePoint, et les groupes `SharingLinks.*` qui portent « Anyone with the link » et « Specific people ». Le rapport est le comportement par défaut ; rien ne change sans `-Apply`, et chaque exécution écrit un CSV de ce qui a été trouvé et de ce qu'il en est advenu |
| Il refuse délibérément deux choses et le dit clairement. Un accès via un groupe Entra ID n'est pas révoqué — le groupe *est* l'accès, et retirer l'utilisateur de SharePoint laisserait l'accès en place tout en donnant l'impression qu'il est fermé ; le groupe est nommé dans le CSV sous `Action = CannotRevoke`, afin que l'offboarding se fasse visiblement en deux étapes. Un accès accordé à `Everyone` est laissé tel quel pour la raison inverse : le supprimer révoque l'accès pour tout le tenant plutôt que pour cette seule personne |
| La couche d'authentification app-only et REST SharePoint est partagée avec le rapport d'autorisations, octet pour octet, délimitée par `SHARED BLOCK START/END`. Il a fallu quatre exécutions réelles contre un tenant pour la mettre au point, et une seconde copie qui dériverait discrètement représente un risque d'exactitude dans le script qui supprime des autorisations. Pour la rendre partageable, la bannière du rapport a été déplacée au-dessus du bloc et le nom de l'application temporaire provient désormais de `$TempAppNamePrefix` ; le comportement du rapport est inchangé |
| Ajout de [`scripts/SharePoint/Test-SharePointAccessScripts.ps1`](scripts/SharePoint/Test-SharePointAccessScripts.ps1), qui vérifie que les deux copies sont identiques (en affichant la première ligne différente lorsqu'elles ne le sont pas) et exerce réellement l'entonnoir de révocation : un essai à blanc enregistre son intention et n'exécute rien, `-Apply` exécute et enregistre, un échec aboutit dans la piste d'audit au lieu de disparaître, et les refus ci-dessus restent des refus. 27 contrôles, tous réussis, exécutables depuis n'importe quel répertoire |
| Ajout de l'entrée de menu (touche `W`) et documentation des deux scripts dans le readme du dossier SharePoint, l'arborescence du dépôt et cette liste de catégories. Vérifié que chaque ancre `docs` de ce readme se résout |

### 2026-09-29 (6)
| Modification |
|--------|
| [`Repair-AppxPackageStore.ps1`](scripts/Device/Repair-AppxPackageStore.ps1) n'essaie plus de réenregistrer un package qui a été remplacé. La première exécution réelle a tenté `aimgr_0.20.61.0` à l'étape 3 et a obtenu `0x80073D06` (« a higher version 0.20.62.0 of this package is already installed ») : une ancienne version dont le statut n'est pas Ok alors qu'une version plus récente du même package se trouve à côté n'est pas un dommage, mais Windows qui attend de la supprimer, et la réenregistrer ne peut jamais réussir. Le diagnostic compare désormais les versions par nom de package, architecture et resource id, signale ces cas comme *Superseded* sur une ligne grise, et les exclut du nombre de réparations ; l'étape 3 traite aussi un `0x80073D06` qui surviendrait malgré tout comme « laissé à Windows » plutôt que comme un avertissement |
| Vérifié hors ligne sous PowerShell 7 et 5.1 avec exactement cette paire (0.20.61.0 `Modified` à côté de 0.20.62.0 `Ok`) : signalé comme remplacé et non compté, tandis qu'un package réellement endommagé dans la même exécution est toujours retenu pour réenregistrement. Pas encore relancé sur l'hôte où cela s'est produit |

### 2026-09-29 (5)
| Modification |
|--------|
| [`Repair-AppxPackageStore.ps1`](scripts/Device/Repair-AppxPackageStore.ps1) peut récupérer Teams et le nouvel Outlook via winget (`-UseWinget`) : `winget download` de `Microsoft.Teams` / `Microsoft.Outlook`, dont les manifestes pointent vers le MSIX sur le CDN de Microsoft, puis `Add-AppxProvisionedPackage` avec les éventuelles dépendances apportées par winget, de sorte que le package arrive pour tous les utilisateurs et pas seulement pour celui qui a lancé `winget install`. Chaque fichier doit porter une signature Microsoft valide. Les manifestes de winget sont en retard sur les installateurs de Microsoft (vérifié aujourd'hui : Teams 26198 contre les 26246 demandés par les profils, Outlook 1.2026.812 contre 902), donc les installateurs restent la valeur par défaut et l'exécution avertit lorsque le build de winget est plus ancien que ce que demandent les profils. `-WingetId` provisionne n'importe quelle autre application de la même façon |
| Il liste désormais chaque application en échec, et pas seulement celles du magasin de packages : l'étape 1c lit le journal de déploiement AppX et le journal FSLogix Apps sur `-Days`, regroupés par package, avec les codes d'erreur nommés (`0x80073D02` en cours d'utilisation, `0x80073CF6` échec de l'enregistrement, ...), les versions demandées et si cet hôte possède leurs fichiers ; `0x80070490` en premier, top 15 |
| Trouvé pendant les tests, corrigé : `Get-WinEvent` lève une erreur bloquante pour un fournisseur non enregistré — toute machine sans FSLogix — que `-ErrorAction SilentlyContinue` n'intercepte pas, de sorte que le contrôle FSLogix aurait interrompu l'exécution à cet endroit ; toutes les lectures d'événements passent désormais par un seul wrapper. Et le filtre de progression de winget contenait deux caractères non ASCII, que Windows PowerShell 5.1 lit comme ANSI dans un fichier sans BOM avant d'échouer à l'analyser — le script entier n'aurait pas pu s'exécuter depuis NinjaOne. Le fichier est de nouveau en pur ASCII, vérifié |
| Vérifié sous PowerShell 7 et 5.1 : la vue d'ensemble des applications en échec contre le journal AppX **réel** de cette machine (20 packages, codes traduits, limité à 15) ; un **vrai** `winget download` de `Microsoft.Outlook` (MSIX de 32 Mo, hash vérifié par winget, signature par le script) jusqu'à un `Add-AppxProvisionedPackage` simulé ; et le scénario de magasin simulé précédent, inchangé. Le provisionnement lui-même et le téléchargement de Teams (271 Mo) n'ont pas été exécutés ici |

### 2026-09-29 (4)
| Modification |
|--------|
| Nouveau [`scripts/Device/Repair-AppxPackageStore.ps1`](scripts/Device/Repair-AppxPackageStore.ps1) : répare les packages AppX qui échouent avec `0x80070490` et un chemin vide (« Deployment Register operation ... from:  (AppxManifest.xml) »), pour n'importe quel package — la même erreur est revenue pour `Microsoft.OutlookForWindows` sur un hôte où seul Teams disposait d'une réparation, et `Update-TeamsClient.ps1 -RepairAppxStore` est limité à `MSTeams` par conception |
| Les erreurs sur cet hôte étaient journalisées par `Apps (Microsoft-FSLogix-Apps)`, ce qui est une autre cause qu'un magasin endommagé : FSLogix enregistre les packages de chaque utilisateur avec leur version exacte dans `AppxPackages.xml` et les rejoue à la connexion (`InstallAppxPackages`, activé par défaut), de sorte qu'un hôte qui provisionne un autre build — ou aucun — répond `0x80070490`. Le script lit ces événements et compare la version demandée par les profils à celle que l'hôte provisionne, vérifie le build FSLogix par rapport aux premières versions qui enregistrent Teams (2210 HF4) et Outlook (25.06) par nom de famille, et avec `-Provision` remet Teams / le nouvel Outlook en place pour tous les utilisateurs avec l'installateur propre de Microsoft (`teamsbootstrapper.exe -p`, Outlook `Setup.exe --provision true --quiet --start-` ; les deux liens vérifiés comme pointant vers le CDN de Microsoft, les deux signatures vérifiées avant l'exécution) |
| La réparation du magasin généralise celle de Teams et ajoute ce qui la rend sûre sur un magasin entier : chaque clé de registre est exportée vers une sauvegarde `.reg` avant d'être supprimée, et n'est pas supprimée si la sauvegarde échoue ; avec un `-Name` générique, les packages système et framework sont signalés mais jamais touchés, et les marqueurs Deprovisioned (la façon dont les suppressions de bloatware sont mémorisées) sont laissés intacts ; un enregistrement utilisateur compte aussi comme orphelin lorsque son package n'a de fichiers nulle part, et pas seulement lorsque son SID n'a pas de profil. La modification de `StateRepository-Machine.srd` ou de `AppxPackages.xml` a été étudiée et délibérément écartée — les deux ne sont pas pris en charge |
| Intégré à [`menu.ps1`](menu.ps1) sous la touche Device `R` (diagnostic, sauf si vous confirmez la réparation ; pose une question distincte pour le provisionnement), documenté dans le readme Device |
| Vérifié hors ligne sous PowerShell 7 et 5.1 avec des cmdlets AppX simulées, des événements FSLogix et un `AppxAllUserStore` de test dans HKCU : le package Teams à chemin vide est détecté comme fantôme, ses entrées utilisateur et machine comme orphelines, une entrée Outlook pour un SID sans profil comme orpheline, le Teams provisionné plus ancien et l'Outlook manquant comme nécessitant un provisionnement, un fantôme de framework et un marqueur Deprovisioned sont laissés intacts avec `*` et inclus lorsqu'ils sont nommés, et la sauvegarde `.reg` est écrite. Cette exécution a aussi révélé que `-Name A,B` arrivait comme une seule chaîne via `powershell.exe -File` (et via les relances du script lui-même), désormais scindée. **Non exécuté sur un hôte réel** : aucune élévation ni aucun hôte AVD n'était disponible ici, donc les suppressions, les modifications du registre et les deux installateurs n'ont pas été exercés pour de vrai |

### 2026-09-29 (3)
| Modification |
|--------|
| [`Restore-MailboxMessages.ps1`](scripts/Exchange/Restore-MailboxMessages.ps1) traite désormais « tout depuis cette date jusqu'à maintenant » comme un cas à part entière : `-After` a reçu les alias `-From` et `-Since`, et l'entrée de menu (`L`) demande s'il faut restaurer uniquement ce jour-là ou tout ce qui a suivi. La fenêtre le permettait déjà, mais la recherche d'audit s'exécutait comme une seule requête sur toute la fenêtre, et une session de recherche unique s'arrête à 50 000 enregistrements à l'échelle du tenant — sur quelques semaines, cela faisait disparaître en silence des actions sur la boîte aux lettres en cours de restauration |
| Le journal d'audit est désormais parcouru jour par jour, chaque jour dans sa propre session avec pagination, et seuls les enregistrements qui mentionnent cette boîte aux lettres sont conservés en mémoire. Un jour qui dépasse à lui seul 50 000 enregistrements est nommé dans un avertissement |
| Une longue fenêtre peut remonter au-delà de ce que la boîte aux lettres conserve encore, ce qui se lirait comme « rien n'a été supprimé ». L'exécution avertit désormais lorsque la fenêtre commence avant le `RetainDeletedItemsFor` de la boîte aux lettres (14 jours par défaut) et que la boîte aux lettres n'est pas en conservation, et lorsqu'elle commence il y a plus de 180 jours, au-delà de la rétention d'audit habituelle |
| Vérifié hors ligne sous PowerShell 7 et 5.1 : `-Since` se lie à `-After` ; un `Search-UnifiedAuditLog` simulé sur une fenêtre de 2,6 jours a été appelé par jour avec des bornes UTC, la dernière tranche se terminant à la fin de la fenêtre, le premier jour paginé sur deux appels dans une même session, et seuls les enregistrements de cette boîte aux lettres conservés. L'avertissement de rétention n'est **pas** exercé - il nécessite un `Get-Mailbox` réel |

### 2026-09-29 (2)
| Modification |
|--------|
| [`Restore-MailboxMessages.ps1`](scripts/Exchange/Restore-MailboxMessages.ps1) n'a plus besoin du rôle Mailbox Import Export pour récupérer le courrier supprimé. La première exécution réelle s'est arrêtée sur « Get-RecoverableItems is not available » et a ignoré chaque message supprimé, alors que ce rôle ne fait partie d'aucun groupe de rôles par défaut — sur la plupart des tenants, la partie suppression ne faisait donc tout simplement rien |
| Sans ce rôle, l'exécution bascule désormais sur Graph et restaure **tout** ce qui a été supprimé dans la fenêtre, pas seulement ce que le journal d'audit a vu : chaque message dans Éléments supprimés et Recoverable Items\Deletions dont l'heure de modification tombe dans la fenêtre est remis en place. Les suppressions auditées (`MoveToDeletedItems`, `SoftDelete`, qu'Exchange audite par défaut pour le propriétaire) retournent dans le dossier que l'enregistrement indique avoir été quitté, l'acteur étant associé exactement via le MessageId ; le reste va dans la Boîte de réception. Un message supprimé *depuis* Éléments supprimés va lui aussi dans la Boîte de réception, car le remettre dans Éléments supprimés ne revient pas à le récupérer. Les éléments supprimés définitivement (Purges) sont hors de portée de Graph et signalés comme `Unreachable` au lieu de manquer en silence |
| Les recherches de messages portent désormais aussi sur `recoverableitemsdeletions`, que `/messages` ne couvre pas, de sorte qu'un message déplacé puis supprimé est trouvé dans l'une ou l'autre partie. Les parties déplacement et suppression partagent un seul chemin de recherche / déplacement / rapport au lieu de deux copies |
| Vérifié hors ligne sous PowerShell 7 et 5.1 contre un Graph simulé : un message déplacé puis supprimé de façon réversible retourne dans le sous-dossier d'origine, un message supprimé depuis Éléments supprimés va dans la Boîte de réception, les éléments non audités des deux dossiers sont restaurés, un élément supprimé cinq jours plus tôt est laissé tel quel, une suppression définitive est signalée `Unreachable`, et l'aperçu comme `-Apply` émettent exactement les déplacements attendus. **Non exécuté sur un tenant réel** ; en particulier, on n'a pas testé si Graph autorise un déplacement hors de `recoverableitemsdeletions`, ni si un déplacement modifie `lastModifiedDateTime` |

### 2026-09-29
| Modification |
|--------|
| Nouveau [`scripts/Exchange/Restore-MailboxMessages.ps1`](scripts/Exchange/Restore-MailboxMessages.ps1) : remet en place les messages qui ont été déplacés ou supprimés dans une boîte aux lettres un jour donné, et indique qui l'a fait. Il n'existait aucun moyen de revenir en arrière après un archivage raté ou une suppression massive, à part une restauration à la main dans Outlook, et aucune réponse à « qui a fait ça » sans écrire une requête de journal d'audit de zéro |
| Les messages supprimés sont restaurés via `Get-/Restore-RecoverableItems` (Éléments supprimés, Recoverable Items, Purges), un `EntryID` à la fois, filtrés sur le moment de la suppression — Exchange connaît lui-même le dossier d'origine. Après un `-Apply`, les dossiers sont relus, et tout ce qui s'y trouve encore est signalé comme `NotRestored` au lieu de se fier au silence de la cmdlet |
| Les messages déplacés n'ont pas cette mémoire : ni Graph ni Exchange n'enregistrent d'où provient un message déplacé. Le Unified Audit Log le fait, donc chaque `Move` audité est retracé jusqu'au **premier** dossier que le message a quitté ce jour-là, localisé via Graph par Internet MessageId et remis en place via `$batch`. Les dossiers sont associés selon leur chemin tel que l'écrit le journal d'audit, c'est-à-dire dans la langue propre de la boîte aux lettres (`\Postvak IN\Projecten`). Les déplacements hors d'Éléments supprimés ou de Recoverable Items sont ignorés, car il s'agissait de restaurations et les annuler supprimerait de nouveau le message |
| Les mêmes enregistrements d'audit nomment l'acteur — compte, propriétaire/délégué/admin, client (Outlook, OWA, application Graph avec son ID d'application), IP — par message dans le CSV, sous forme de tableau groupé « qui a déplacé / supprimé quoi » à l'écran, et sous forme de `_Audit.csv` brut. Les suppressions sont attribuées par objet et par heure la plus proche, car les éléments récupérables ne portent pas de MessageId. L'exécution liste aussi lesquelles des quatre actions ne sont pas auditées sur la boîte aux lettres, car par défaut le propre `Move` du propriétaire ne l'est pas, et un enregistrement manquant se lirait sinon comme « personne ne l'a fait » |
| Les éléments d'archive sans enregistrement d'audit (par ex. après [`Move-InboxToArchive.ps1`](scripts/Exchange/Move-InboxToArchive.ps1)) sont listés par heure de modification et ne sont déplacés vers la Boîte de réception qu'avec `-UnauditedArchiveToInbox`, car la lecture ou le marquage modifient aussi cette heure. L'accès Graph réutilise le modèle à trois voies purement REST de [`Remove-PhishingMessage.ps1`](scripts/Exchange/Remove-PhishingMessage.ps1), de sorte qu'il s'exécute à côté de la session Exchange sans le conflit MSAL. Ajouté au sous-menu Exchange sous `L` (aperçu d'abord, puis `-Apply`) |
| Vérifié hors ligne uniquement, sous PowerShell 7 et Windows PowerShell 5.1 : vérification de syntaxe, analyse des enregistrements d'audit sur des enregistrements fabriqués (filtre de boîte aux lettres, UTC vers heure locale, types de connexion, libellés de client, attribution par objet/heure), et tout le chemin des messages déplacés contre un Graph simulé — une chaîne de déplacements revenant au premier dossier, des noms de dossiers en néerlandais, une restauration faite par un utilisateur laissée telle quelle, un message déjà revenu ignoré, un élément d'archive audité tenu hors de la liste non auditée, et les requêtes de déplacement qui en résultent. **Pas encore exécuté sur un tenant réel** : les propriétés de sortie exactes de `Get-RecoverableItems`, la façon dont il interprète les heures du filtre, et le fait que les enregistrements d'audit portent `InternetMessageId` pour chaque client n'ont pas été testés |

### 2026-09-28 (6)
| Modification |
|--------|
| Le redémarrage déclenché par `-RestartIfNeeded` attendait 60 secondes sur un hôte où personne ne pouvait regarder. Le compte à rebours sert à prévenir les gens, il ne s'applique donc désormais que lorsqu'il y a des gens : avec quelqu'un de connecté, c'est `-RestartDelaySeconds` et `shutdown /a` l'arrête ; sans personne de connecté - le cas normal au démarrage, et le cas garanti sur un hôte de session dont le pool est en mode drain - il redémarre en quelques secondes. Quelques secondes sont conservées pour que la propre ligne de journal de l'exécution soit écrite avant le début de l'arrêt |
| Envisagé de bloquer les connexions depuis l'invité (`change logon /drainuntilrestart`) pour refermer la fenêtre entre le démarrage de la tâche et le redémarrage, puis abandonné : un hôte redémarré de cette façon a déjà son pool en mode drain, donc le switch côté invité ne ferait que dupliquer ce que le pool garantit - et il laisserait les connexions bloquées pour toute exécution qui planterait avant de les réactiver |
| Vérifié que `change.exe` et `chglogon.exe` existent sur ce build de Windows 11 et que `change logon /query` indique « Session logins are currently ENABLED » tout en sortant avec 1, raison pour laquelle l'idée a été mesurée avant d'être abandonnée plutôt qu'après |

### 2026-09-28 (5)
| Modification |
|--------|
| [`Update-TeamsClient.ps1`](scripts/Device/Update-TeamsClient.ps1) peut désormais réparer l'hôte au lieu de seulement le diagnostiquer. Un hôte de session où chaque voie répondait `0x80070490` avait un `AppxAllUserStore` rempli d'entrées que Windows ne peut plus résoudre, et le conseil honnête à ce stade était « redéployer » — ce que personne ne veut entendre pour une machine qui, par ailleurs, fonctionne bien |
| `-RepairAppxStore` procède en deux étapes. Un package dont les fichiers sont encore sur le disque est réenregistré à partir de son propre manifeste (`Add-AppxPackage -Register`), ce qui reconstruit la connaissance qu'en a le magasin et permet généralement à la suppression ordinaire de fonctionner à nouveau. Ce qui résiste est supprimé clé par clé : les enregistrements sous un SID sans profil sur cet hôte (également sous `EndOfLife` et `DeferredRemoval`), une entrée `Applications` à l'échelle de la machine dont le manifeste a disparu, et le marqueur `Deprovisioned` qui refuse purement et simplement le provisionnement. Chaque clé est nommée avec son chemin de registre complet avant d'être supprimée, et rien en dehors de MSTeams n'est jamais touché |
| Le preflight signale ces orphelins que le switch soit fourni ou non, de sorte que `-CheckOnly` constitue le diagnostic et que la réparation est une décision distincte — la même logique que `-ClearOrphanedAddInRegistration` pour Windows Installer |
| `-UseWinget` attaque le problème par l'autre côté : winget télécharge le MSIX en le vérifiant contre le SHA256 de son propre manifeste, et le bootstrapper provisionne ce fichier avec `-p -o`. Le déploiement dispose alors d'une source explicite plutôt que d'une entrée de magasin qu'il doit résoudre, et l'exécution sait quel build elle a installé. Le manifeste de winget est en retard sur le service de configuration — mesuré à `26198.304.4946.9672` contre un build `26246` — et l'exécution le signale lorsque c'est le cas |
| Exécuté en tant que System, winget n'est pas du tout dans le `PATH` : son alias est un shim MSIX par utilisateur. Il est donc résolu depuis `Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*`, ce qui a été vérifié en vidant `PATH` et en observant la solution de repli le trouver |
| Le lecteur du magasin a été exécuté contre le `AppxAllUserStore` réel de ce poste de travail, où il a trouvé deux vrais orphelins (`S-1-0-0` et le SID d'un profil supprimé sous `EndOfLife`), n'en a signalé aucun pour les packages sains et a gardé chaque chemin à l'intérieur du magasin. La **suppression** a été exercée pour de vrai contre un magasin reconstruit sous `HKCU` : 9 entrées MSTeams, 6 orphelines, les 6 supprimées, tandis que les enregistrements sains, l'entrée `Staged`, l'orphelin d'un autre produit et la propre clé du SID mort ont tous survécu |
| `winget download` a été mesuré de bout en bout : 271 Mo en 23 secondes, sans compte Store, avec son propre contrôle de hash, une signature `O=Microsoft Corporation` valide, le dossier de préparation vidé au préalable et supprimé ensuite. Contre un hôte dont le magasin de packages est réellement endommagé, les deux switches ne sont **pas testés** — aucune machine ici n'en a un |
| La ligne de redémarrage ne cite plus de raison. Elle indiquait « MSI returned 3010 » alors que trois choses différentes la déclenchent, et orientait les lecteurs vers un installateur qui n'avait jamais été exécuté |

### 2026-09-28 (4)
| Modification |
|--------|
| [`Update-TeamsClient.ps1`](scripts/Device/Update-TeamsClient.ps1) donnait des conseils qui ne pouvaient pas aider. Un hôte de session de production répondait `0x80070490` (« Element not found ») pour **chaque** détenteur du package, `NT AUTHORITY\SYSTEM` compris, et le script disait encore de mettre l'hôte en drain et de déconnecter les utilisateurs — sur un hôte déjà en drain et sans personne dessus. Un enregistrement que le magasin de packages ne trouve pas n'est pas un utilisateur qui détient le package |
| Le preflight indique désormais si une suppression en attente a encore quelqu'un à attendre. `Installed(pending removal)` ne se termine qu'à une déconnexion, donc un SID sans profil sous `ProfileList` attend un événement qui ne peut jamais se produire ; ceux-là sont signalés séparément de ceux qui attendent réellement, et seuls ces derniers signalent un redémarrage |
| Il vérifie aussi le seul état dont rien ne se remet tout seul : un package que le magasin liste mais dont le `InstallLocation` a disparu, ou qui n'a aucun emplacement d'installation. Cette seule ligne explique tout l'échec — chaque suppression répond `0x80070490` parce qu'il n'y a rien à supprimer, et provisionner la même version y répond aussi |
| Lorsque la suppression par utilisateur répond ce code pour chaque détenteur, l'exécution le dit clairement, et l'échec final change de conseil en conséquence : non plus « mettre l'hôte en drain », mais le fait que `Remove-AppxPackage`, le bootstrapper et DISM lisent tous le même magasin incohérent, si bien qu'aucun d'eux ne peut le réparer — un hôte de session poolé est redéployé à partir de son image, un hôte personnel est réparé sur place |
| Exercé contre les chaînes que cet hôte a réellement affichées, plus un SID réel de cette machine comme cas de contraste : quatre profils orphelins sont lus comme orphelins, un vrai profil donne toujours « déconnectez-les », un package sans emplacement d'installation est signalé, et un package sain reste silencieux. La **remédiation** n'est pas testée — cette machine n'a pas de magasin de packages endommagé sur lequel l'essayer |

### 2026-09-28 (3)
| Modification |
|--------|
| [`Init-TempDisk.ps1`](scripts/Device/TempDisk/Init-TempDisk.ps1) réparait le disque temporaire et y configurait le fichier d'échange, puis laissait la machine tourner le reste de cette session sans - Windows lit la configuration du fichier d'échange au démarrage et ne la relit jamais, donc le démarrage qui a dû reconstruire `D:` est précisément celui où le fichier d'échange n'existe pas. Ajout de `-RestartIfNeeded`, qui comble cet écart au lieu d'attendre le prochain démarrage |
| Un script qui s'exécute à chaque démarrage et peut redémarrer la machine est une boucle de redémarrage en puissance, donc il ne se déclenche que si toutes ces conditions sont réunies : l'exécution s'est terminée proprement (une exécution en échec ne redémarre jamais - cela masquerait l'échec derrière un redémarrage), le disque est présent, le fichier d'échange y est configuré, et la seule chose qui manque est que cette session ne l'utilise pas |
| Personne ne doit être connecté, qu'il s'agisse d'une session active ou déconnectée. Les sessions sont comptées à raison d'un `explorer.exe` par bureau interactif plutôt qu'en analysant `query.exe`, dont les en-têtes de colonnes suivent la langue d'affichage et liraient une liste vide sur un hôte de session néerlandais. `-RestartEvenIfUsersSignedIn` passe outre là où le compte à rebours suffit comme avertissement |
| Au plus un redémarrage par `-RestartCooldownMinutes` (60 par défaut), mémorisé sous forme d'horodatage round-trip sous `HKLM:\SOFTWARE\ICTKanon\InitTempDisk` - un horodatage formaté selon les paramètres régionaux, écrit par une exécution et lu par une autre, est la façon dont un délai de carence cesse discrètement de fonctionner. Un second redémarrage pour la même raison signifie que le premier n'a pas aidé, et l'exécution le dit au lieu de le répéter |
| Le redémarrage passe par `shutdown.exe` avec un compte à rebours de 60 secondes et la raison planifiée « Operating System: Reconfiguration », de sorte que toute personne sur la machine le voit venir, que `shutdown /a` l'arrête, et qu'il n'est pas signalé comme un redémarrage inattendu |
| [`Register-InitTempDiskTask.ps1`](scripts/Device/TempDisk/Register-InitTempDiskTask.ps1) déploie désormais la tâche avec `-Quiet -RestartIfNeeded`, et indique au moment de l'enregistrement si la tâche peut redémarrer la machine. `-ScriptArguments '-Quiet'` exclut le redémarrage |
| Numérotation des deux entrées précédentes d'aujourd'hui : deux titres `### 2026-09-28` identiques s'étaient retrouvés dans l'historique, ce qui se lit comme une seule modification coupée en deux |
| Vérifié sur cette machine sous PowerShell 5.1 et 7 : la détection de session nomme le compte connecté (donc cette machine refuserait de redémarrer), un marqueur absent se lit comme `$null`, un marqueur écrit puis relu est analysé en `DateTime` et bloque un second redémarrage pendant le délai de carence, un marqueur vieux de 90 minutes en autorise un, et un marqueur corrompu se dégrade en « pas de marqueur » au lieu de lever une exception |
| Mesuré également la façon dont un `shutdown.exe` en échec se signale, car le code bifurque dessus : `shutdown /a` sans rien en attente répond 1116 et définit `$LASTEXITCODE` sous 5.1 comme sous 7.6 sans lever d'exception, donc c'est la branche du code de sortie qui s'exécute. Le `try` qui l'entoure reste pour `$PSNativeCommandUseErrorActionPreference`, qui peut en faire une erreur bloquante à partir de 7.4 |
| **Aucun redémarrage n'a été déclenché depuis cette session** et les garde-fous restent non vérifiés sur une VM Azure réelle |

### 2026-09-28 (2)
| Modification |
|--------|
| [`Update-TeamsClient.ps1`](scripts/Device/Update-TeamsClient.ps1) plantait dès le contrôle préalable sur toute machine où le complément de réunion n'est enregistré nulle part : `The property 'Count' cannot be found on this object`. `$x = if (...) { @() }` affecte `$null`, car un tableau vide écrit dans le pipeline correspond à zéro objet — le `@()` doit entourer le `if` entier, et non se trouver dans ses branches. Reproduit sur la version committée puis corrigé ; les cinq chemins de la fonction de rapport passent désormais, et il est démontré que les trois chemins vides levaient une exception auparavant |
| Un package que la pile AppX refuse de supprimer ne met plus fin à l'exécution. `Remove-AppxPackage -AllUsers` a répondu `Catastrophic failure` sur un hôte de session portant deux versions de `MSTeams`, et `-ErrorAction` ne couvre pas une erreur bloquante ; il fallait donc un `try`/`catch`. L'exécution se poursuit et le provisionnement met à niveau sur place ce qui a survécu — s'arrêter à ce moment laissait l'hôte avec le complément désinstallé et sans Teams réinstallé |
| La **désinstallation** du complément est passée de l'étape 6 à l'étape 8, à côté de l'installation qui la remplace. Le nettoyage y avait déjà été déplacé ; laisser la désinstallation en arrière signifiait que tout échec ultérieur produisait le même résultat par un autre chemin. Tout ce qui est destructif pour le complément se trouve désormais à côté de ce qui l'annule |
| Les codes de sortie sont lisibles. `teamsbootstrapper.exe` répond par un HRESULT, que PowerShell affiche sous forme de grand entier négatif : « exit code -2147023728 » ne dit rien, `0x80070490 - Element not found` indique où chercher. Les codes MSI restent des nombres simples, et un HRESULT hors de la facility Win32 retombe sur de l'hexadécimal brut plutôt que d'inventer une signification |
| Un provisionnement en échec tente désormais une fois la désinstallation à l'échelle de la machine documentée par Microsoft (`teamsbootstrapper.exe -x -m`) et provisionne à nouveau avant d'abandonner, et l'erreur levée nomme la cause habituelle sur un hôte de session : un package retenu par un utilisateur connecté. **Non testé** — cette récupération n'a pas encore été exécutée sur un hôte qui en avait besoin |
| Ce message d'erreur mangeait ensuite son propre code de sortie : `-f` a une priorité plus forte que `+`, si bien que le formatage d'une chaîne concaténée ne s'appliquait qu'au dernier morceau et affichait un `{0}` littéral. Signalé depuis un hôte réel, où cela masquait précisément le code dont le lecteur avait besoin |
| Le bootstrapper affiche son propre verdict, et l'exécuter en mode masqué le jetait. `Invoke-Installer` peut désormais capturer stdout et stderr, et les lignes du bootstrapper sont renvoyées en sortie sous la forme `bootstrapper:`. Vérifié avec un processus qui affiche un verdict JSON et se termine avec un code non nul ; sans le switch, rien n'est capturé et aucun fichier temporaire n'est laissé |
| La capture de cette sortie a ensuite cassé tous les verdicts, et uniquement sur le runtime qui compte : sous Windows PowerShell 5.1, un `Start-Process -PassThru` redirigé ne renvoie aucun code de sortie, sauf si le handle du processus est d'abord sollicité. Un `-x -m` qui affichait `{"success": true}` était signalé comme un échec. Mesuré sur les deux runtimes ; `$null = $proc.Handle` le corrige, et PowerShell 7 n'a jamais eu le problème |
| Le succès est désormais déterminé par le propre verdict JSON du bootstrapper lorsqu'il y en a un, et non par un code de sortie qui peut manquer. Une sortie non JSON, un JSON étranger ou une sortie malformée retombent tous sur le code de sortie plutôt que de deviner |
| Le contrôle préalable comptait les entrées `PackageUserInformation` et déclarait un package « installé pour 1 profil utilisateur » alors que sa seule entrée était `S-1-5-18` qui le **mettait en staging** — ni un utilisateur, ni une installation. Il lit désormais les états : un package uniquement en staging est signalé comme tel, et un package dont les entrées indiquent `Installed(pending removal)` est signalé comme déjà supprimé et en attente de la déconnexion de ces utilisateurs, nommés, avec le redémarrage signalé. Mesuré sur les deux packages réellement remontés par un hôte de session de production |
| `Remove-AppxPackage -AllUsers` fonctionne en tout ou rien, donc un seul profil inaccessible fait échouer tout l'appel. Dans ce cas, le package est désormais supprimé utilisateur par utilisateur, et celui qui le retient encore est nommé — « un utilisateur connecté le retient » n'est pas exploitable tant qu'on ne sait pas lequel. Le SID est extrait de la forme texte de `PackageUserInformation`, qui diffère selon les builds, testée sur sept formes dont celle d'Entra `S-1-12-1`. **Non testé** sur un package qui refuse réellement la suppression |
| Le journal AppX a ensuite répondu à la question que masquait le `0x80070490` du bootstrapper : il provisionne `MSTeams_26246...`, le build enregistré pour un seul profil, et Windows ne trouve pas les fichiers de ce package. Deux versions de MSTeams côte à côte dont l'une est cassée : c'est l'état à rechercher |
| Un provisionnement en échec affiche désormais aussi les erreurs de déploiement AppX journalisées par Windows, qui portent la raison que masque son `0x80070490` — « Unable to install because the following apps need to be closed &lt;package&gt; ». En lecture seule, 10 ms, et silencieux lorsque le journal ne contient rien de récent |
| Le contrôle préalable signale un package AppX dont le `Status` n'est pas `Ok`. Un package que Windows considère comme Modified ou Tampered est précisément ce qui fait répondre `Catastrophic failure` à `Remove-AppxPackage` et échouer le provisionnement ensuite, et c'était invisible jusqu'ici |

### 2026-09-28
| Modification |
|--------|
| Ajout de [`scripts/Device/TempDisk/Init-TempDisk.ps1`](scripts/Device/TempDisk/Init-TempDisk.ps1) : le disque temporaire éphémère d'une VM Azure est effacé à chaque désallocation, redimensionnement ou déplacement d'hôte, et revient en RAW, hors ligne ou sans sa lettre de lecteur. Windows lit la configuration du fichier d'échange au démarrage et ne la relit jamais, donc un fichier d'échange configuré sur `D:` absent au démarrage n'est tout simplement jamais créé, et la machine pagine à nouveau sur `C:` - ou tourne sans aucun fichier d'échange. Le script restaure le volume en `D:` et y repointe le fichier d'échange |
| Un disque temporaire qui a seulement perdu sa lettre de lecteur la récupère au lieu d'être reformaté ; il est reconnu à son nom de volume (`Temporary Storage`) ou au fichier `DataLoss_Warning_Readme.txt` qu'Azure écrit sur le disque de ressources. Seul un disque RAW, non boot et non système est jamais initialisé : un disque temporaire vide et un disque de données non formaté se ressemblent parfaitement vus de l'extérieur, donc un disque comportant des partitions est signalé et laissé tel quel, et s'il y a plus d'un candidat RAW, le script refuse de deviner et demande `-DiskNumber`. `-Force` combiné à `-DiskNumber` est le seul moyen de formater un disque qui contient encore des données |
| Un lecteur optique qui occupe `D:` est d'abord déplacé - Windows attribue `D:` au lecteur DVD sur une image sans disque temporaire et ne la rend jamais, ce qui est la deuxième façon dont le fichier d'échange se retrouve sur `C:` |
| L'exécution distingue le fichier d'échange *configuré* (registre) du fichier d'échange *en cours d'utilisation* (cette session) et indique lequel est lequel, au lieu de signaler un succès pour une modification qui ne prend effet qu'au prochain redémarrage. Configurer un fichier d'échange sur un lecteur qui n'a pas pu être restauré est un échec bloquant plutôt qu'un paramètre que Windows ignore en silence |
| Ajout de [`scripts/Device/TempDisk/Register-InitTempDiskTask.ps1`](scripts/Device/TempDisk/Register-InitTempDiskTask.ps1), construit à partir d'un brouillon qui comportait deux défauts : la valeur par défaut de son `-ScriptSourcePath` faisait référence à `$ScriptTargetDir`, un paramètre déclaré *après* lui, si bien que la valeur par défaut s'étendait en `\Init-TempDisk.ps1` et ne se résolvait jamais ; et la copie s'exécutait avec `-ErrorAction SilentlyContinue`, donc une source manquante enregistrait une tâche de démarrage pointant vers un fichier inexistant - qui échoue ensuite à chaque démarrage sans que personne ne le voie. La source est désormais par défaut la copie située à côté du script, une source manquante est une erreur bloquante, et l'existence de la tâche est vérifiée après l'enregistrement |
| Les deux sont documentés dans un nouveau [`scripts/Device/TempDisk/readme.md`](scripts/Device/TempDisk/readme.md) ; le dossier a été ajouté au readme de [`Device/`](scripts/Device/readme.fr.md) et à l'arborescence du dépôt, le readme racine a reçu une entrée **Temp Disk & Pagefile (Azure / AVD)**, et [`Init-TempDisk.ps1`](scripts/Device/TempDisk/Init-TempDisk.ps1) a été branché dans [`menu.ps1`](menu.ps1) sous la touche Device `V` (par défaut `-CheckOnly`, sauf si vous confirmez la réparation) |
| Vérifié : les deux fichiers passent sans erreur [`Test-PowerShellSyntax.ps1`](scripts/Startup/Test-PowerShellSyntax.ps1). Les fonctions auxiliaires en lecture seule ont été réellement exécutées sur cette machine Windows 11 sous PowerShell 5.1 et 7 avec `Set-StrictMode -Version Latest` - recherche de lettre libre, recherche du lecteur optique, détection du disque temporaire et des candidats RAW, et lecture de l'état du fichier d'échange (qui a correctement indiqué la gestion automatique activée et `C:\pagefile.sys` en cours d'utilisation) - et les objets déclencheur, principal et paramètres de la tâche planifiée ont été construits et leurs valeurs contrôlées. **Pas encore vérifié sur une VM Azure ou un hôte de session réels** : aucun disque n'a été initialisé, aucun fichier d'échange modifié et aucune tâche enregistrée depuis cette session |

### 2026-09-25 (20)
| Modification |
|--------|
| La feuille `Pivot toegang` a été retournée pour imbriquer **site → groupe → personne** au lieu de site → personne → groupe. C'est ainsi que SharePoint accorde réellement l'accès — un site a des groupes, les groupes ont des personnes — donc, repliée, elle liste les groupes d'un site et, dépliée, elle nomme tous ceux qu'ils laissent entrer |
| Ajout de `Pivot per persoon` (personne → site → groupe) afin que l'autre sens reste consultable depuis la même feuille : ce qu'une personne donnée atteint, et par quel biais. C'est la question du départ d'un collaborateur, et un pivot centré sur le site ne peut pas y répondre |
| Une personne à qui l'accès est accordé directement n'a pas de groupe, et dans une hiérarchie à trois niveaux, ce niveau intermédiaire vide se lit comme une donnée manquante plutôt que comme « accordé sans groupe ». `ViaName` indique désormais `(direct toegekend)` pour ces lignes au lieu d'être vide |
| Vérifié localement avec 240 contrôles répartis sur onze suites : les deux pivots sont relus depuis le classeur avec leurs champs de ligne dans l'ordre prévu, et un accès direct est nommé plutôt que vide |

### 2026-09-25 (19)
| Modification |
|--------|
| La documentation a été auditée par rapport aux règles du dépôt au lieu de la supposer complète, et trois lacunes ont été trouvées. Les paramètres étaient en ordre : les 19 paramètres de [`Get-SharePointPermissionsReport.ps1`](scripts/Reporting/Get-SharePointPermissionsReport.ps1) figurent dans l'aide basée sur les commentaires et dans le tableau des paramètres du readme du dossier, sans rien d'obsolète dans l'un ni dans l'autre |
| [`scripts/Reporting/readme.md`](scripts/Reporting/readme.md) n'avait aucun tableau `## Scripts`, alors que [`Exchange/`](scripts/Exchange/readme.fr.md), [`Entra/`](scripts/Entra/readme.fr.md), [`Device/`](scripts/Device/readme.fr.md) et [`SharePoint/`](scripts/SharePoint/readme.fr.md) en ont tous un. Il a été ajouté, couvrant les cinq entrées du dossier — pas seulement le nouveau script — afin que le tableau décrive le dossier et non sa dernière modification. Chaque lien et chaque ancre qu'il contient ont été vérifiés |
| La catégorie `### 📊 Reporting` du readme racine ne listait que les rapports Computer Last Logon et Licensing. Les trois scripts de reporting SharePoint y manquaient, dont deux antérieurs à ce travail. Ajout d'une entrée pour [`Get-SharePointPermissionsReport.ps1`](scripts/Reporting/Get-SharePointPermissionsReport.ps1) et d'entrées courtes pour [`Get-SharePointStorageReport.ps1`](scripts/Reporting/Get-SharePointStorageReport.ps1) et [`Remove-SharePointFileVersionsByDate.ps1`](scripts/Reporting/Remove-SharePointFileVersionsByDate.ps1) |
| L'arborescence du dépôt décrivait encore le rapport des permissions comme allant « to CSV », ce qui n'était plus vrai depuis l'ajout de `-Excel` ; elle indique désormais CSV + Excel. Le libellé de [`menu.ps1`](menu.ps1) disait « who has access to what, at every level », ce qui décrit l'ancienne forme du rapport plutôt que la vue consolidée par site qu'il présente désormais en premier |
| Confirmé qu'aucune action n'est nécessaire pour [`f.ps1`](f.ps1) : son index se reconstruit lorsque la date d'écriture d'un script change, donc `f-sharepointpermissionsreport -Excel` prend en compte les nouveaux paramètres sans `f-refresh` |

### 2026-09-25 (18)
| Modification |
|--------|
| [`scripts/Reporting/Get-SharePointPermissionsReport.ps1`](scripts/Reporting/Get-SharePointPermissionsReport.ps1) répondait à « quels accès existent » mais pas à la question avec laquelle on l'ouvre réellement : **qui peut atteindre ce SharePoint, et comment y est-il arrivé.** `Rechten` indiquait qu'un groupe avait des droits, `Groepen` indiquait qui en faisait partie, et rien ne reliait les deux — « Site Owners a Full Control » plus « Site Owners contient cinq personnes » n'est pas une réponse. Ajout de `SharePoint_Permissions_SiteAccess_<ts>.csv` (feuille `Toegang`) : une ligne par personne et par site, avec le groupe par lequel passe son accès, l'id de ce groupe et le niveau d'autorisation |
| Consolidé par collection de sites à dessein : une personne qui atteint trente dossiers d'un même site via le même groupe correspond à une ligne, pas à trente. Un niveau différent ou un groupe différent donne une ligne distincte, car il s'agit d'un accès différent. Le détail par étendue reste derrière `-IncludeEffectiveAccess` |
| Trois choses ne disparaissent délibérément pas de cette vue : une personne à qui l'accès est accordé directement apparaît en son nom propre avec `ViaType = Direct` ; `Everyone` et `Everyone except external users` ne se résolvent en aucune personne, mais obtiennent une ligne nommant la revendication (claim), puisque c'est exactement ce que cherche un relecteur ; et avec `-SkipGroupExpansion`, les accès directs restent visibles, seuls les membres des groupes manquent |
| Ajout de `SiteTitle` — une vue consolidée de 130 sites n'est pas lisible sous forme de 130 URL, et le titre du web racine n'est connu que pendant l'analyse de ce web ; il est donc capturé à ce moment-là et recherché pour chaque ligne. Ajout de `ViaId` à côté de `ViaName` après comparaison avec [NovaPoint](https://github.com/Barbarur/NovaPoint/wiki/Solution-Report-PermissionsReport), qui place `GroupId` à côté de `AccessType` pour la même raison : un titre comme `Site Owners` se répète sur chaque site du tenant |
| Ajout d'une feuille `Pivot toegang` imbriquant site → personne → groupe par niveau d'autorisation, filtrée par externe et type d'accès. NovaPoint place ses utilisateurs dans une seule colonne `Users` sous forme de liste ; ici, chaque utilisateur obtient sa propre ligne, ce qui est moins compact à lire mais fait toute la différence entre pouvoir filtrer ou pivoter sur une personne et ne pas le pouvoir |
| Vérifié localement avec 238 contrôles répartis sur onze suites (25 nouveaux) : un accès de groupe liste ses personnes avec le nom du groupe, son id et le niveau ; le même accès via le même groupe sur une étendue plus profonde n'est pas répété, alors qu'un niveau différent l'est ; les accès directs, les revendications `Everyone` et `-SkipGroupExpansion` se comportent tous comme décrit ; les clés de point de reprise sont uniques, une par ligne ; et la feuille et son pivot sont relus depuis le classeur. **Pas encore vérifié sur un tenant réel** |

### 2026-09-25 (17)
| Modification |
|--------|
| Le classeur `-Excel` de [`scripts/Reporting/Get-SharePointPermissionsReport.ps1`](scripts/Reporting/Get-SharePointPermissionsReport.ps1) est désormais réellement exploitable en tableau croisé dynamique. Vérifié d'abord plutôt que supposé : les colonnes numériques arrivent déjà dans Excel comme des nombres, pas comme du texte, donc l'agrégation n'a jamais été le problème — l'obstacle était `PermissionLevels`, que SharePoint remplit avec plusieurs niveaux à la fois (`Read; Limited Access`). Un pivot traite chaque combinaison comme une valeur à part, si bien que `Full Control` et `Full Control; Limited Access` se retrouvent sur des lignes distinctes |
| Les feuilles contenant `PermissionLevels` reçoivent désormais une colonne `PrimaryPermission` juste à côté, contenant le niveau unique le plus fort de cet accès. `Limited Access` perd toujours face à un vrai niveau — SharePoint l'ajoute automatiquement pour la traversée — et un niveau personnalisé se classe au-dessus de `Read` mais en dessous de `Full Control`, car il a été créé délibérément et ne doit pas disparaître derrière un niveau intégré. Les noms de niveaux en néerlandais et en anglais sont tous deux reconnus, ce qui compte sur un tenant en néerlandais |
| Ajout de trois feuilles de pivot prêtes à l'emploi : `Pivot rechten` (site × niveau d'autorisation, nombre d'accès, filtré par principal et type d'étendue), `Pivot principals` (principal × type d'étendue, nombre d'étendues, filtré par site et externe) et `Pivot groepen` (groupe × membre externe, nombre de membres, filtré par site et type de groupe). Chacune ne référence que les colonnes que sa feuille source possède réellement, et une source manquante ou réduite est ignorée plutôt que de produire un pivot cassé |
| La création des pivots se fait au mieux et de façon isolée : un échec émet un avertissement et laisse les feuilles de données intactes, selon le même principe qui interdit au classeur lui-même de coûter les CSV |
| Vérifié localement avec 213 contrôles répartis sur dix suites (33 nouveaux) : le classement des niveaux pour des entrées simples, jointes, inversées, néerlandaises, personnalisées, de casse variable et vides ; la colonne dérivée placée à côté de l'originale sur les bonnes feuilles et pas sur les autres ; et les pivots relus depuis le package avec les bons champs de ligne, de colonne, de données et de filtre, en ignorant les sources absentes et les feuilles sans colonnes. **Pas encore vérifié sur un tenant réel** |

### 2026-09-25 (16)
| Modification |
|--------|
| Ajout de `-Excel` à [`scripts/Reporting/Get-SharePointPermissionsReport.ps1`](scripts/Reporting/Get-SharePointPermissionsReport.ps1) : un seul `.xlsx` à côté des CSV, avec une feuille par rapport — `Samenvatting`, `Rechten`, `Groepen` et, avec `-IncludeEffectiveAccess`, `Effectief` — chacune étant un vrai tableau Excel avec listes déroulantes de filtre et en-tête figé, selon le même modèle `ImportExcel` que [`Get-DistributionGroupMembers.ps1`](scripts/Exchange/Get-DistributionGroupMembers.ps1) |
| Les CSV sont toujours écrits et le classeur est construit à partir d'eux, pas à leur place. C'est dans eux que l'analyse écrit en flux et à eux qu'une exécution reprise ajoute des données, donc ils existent quoi qu'il arrive — et un classeur dont l'écriture échoue (module manquant, fichier ouvert, mémoire insuffisante) ne coûte alors qu'une copie de confort, pas le rapport |
| Une feuille de calcul s'arrête à 1,048,576 lignes et abandonne le reste sans prévenir ; les feuilles sont donc plafonnées à 1,000,000 avec un avertissement nommant la feuille et le CSV qui contient toujours l'intégralité. Sur un grand tenant, seule `Effectief` s'approche réalistement de cette limite |
| Correction d'une vraie lacune dans l'appartenance aux groupes lors de ce câblage : un groupe Entra ID auquel l'accès est accordé **directement** sur un site, une liste ou un élément ne passe jamais par `/sitegroups`, c'était donc le seul type de groupe dont le rapport ne listait jamais les membres — seulement les dix premiers noms dans `MemberPreview`. Ces groupes obtiennent désormais leurs propres lignes dans la sortie Groups, résolus en personnes, enregistrés une fois par groupe plutôt qu'une fois par accès, et partageant le schéma déjà utilisé par les lignes des groupes SharePoint |
| Ajout de la question Excel à l'entrée de [`menu.ps1`](menu.ps1) |
| Vérifié localement avec 180 contrôles répartis sur neuf suites (20 nouveaux) : un classeur est écrit et relu avec ses quatre feuilles dans l'ordre et leurs lignes intactes, une nouvelle exécution remplace au lieu d'ajouter, les sources absentes/vides/manquantes sont ignorées, l'absence de données à écrire ne laisse aucun fichier, une feuille trop volumineuse est plafonnée plutôt que tronquée par Excel, un groupe Entra accordé directement est listé une seule fois avec ses vrais membres, les groupes SharePoint sont laissés à `/sitegroups`, et les deux sources de groupes partagent un même schéma. **Pas encore vérifié sur un tenant réel** |

### 2026-09-25 (15)
| Modification |
|--------|
| Première exécution réelle complète de [`scripts/Reporting/Get-SharePointPermissionsReport.ps1`](scripts/Reporting/Get-SharePointPermissionsReport.ps1) : 130 webs, 2066 listes, 453 étendues uniques, 1554 accès, 83 liens de partage, 13 accès externes, 111 accès `Everyone`. Trois défauts dans la sortie elle-même, tous trouvés en lisant les CSV produits plutôt que les journaux |
| `-IncludeEffectiveAccess` produisait un fichier vide sur un tenant avec 1554 accès et plus d'un millier de membres résolus. La garde était `if ($IncludeEffectiveAccess -and $EffectiveRows)`, et **une `List[object]` vide est évaluée à faux en PowerShell** — le test échouait donc dès la toute première ligne et la liste ne pouvait jamais se remplir, ce qui la maintenait vide, ce qui maintenait l'échec du test. C'est désormais un contrôle explicite `$null -ne` |
| 121 des 158 lignes « could not be read » concernaient une seule liste système masquée, `Lijst met gebruikersgegevens` (modèle 112, la User Information List), sur chaque site. SharePoint refuse `/items` sur cette liste avec `400` quelle que soit la largeur du `$select`, y compris le barreau le plus étroit de l'échelle. Ses éléments sont des enregistrements d'annuaire plutôt que du contenu, donc des étendues au niveau des éléments n'y ont aucun sens pour une revue des accès — le balayage des éléments ignore désormais le modèle 112 et l'indique, tout en rapportant toujours l'étendue propre de la liste |
| Les 37 restantes étaient obsolètes : des lignes d'erreur écrites par la tentative précédente interrompue, reportées dans le CSV final par la reprise alors que ces listes avaient réussi lors de la nouvelle tentative. Une unité en échec est délibérément laissée non marquée pour être retentée, mais rien ne supprimait ses anciennes lignes. Les lignes de détail portent désormais le `UnitKey` qui les a produites, et une ligne dont l'unité est marquée comme terminée est écartée lors de l'écriture du CSV final — ainsi, le décompte d'incomplétude décrit le fichier que le lecteur ouvre |
| Le résumé est désormais construit à partir du CSV de détail publié plutôt qu'à partir du fichier partiel, de sorte que ses décomptes et le fichier concordent |
| Vérifié localement avec 158 contrôles répartis sur huit suites (26 nouveaux) : les lignes effectives sont émises une par utilisateur résolu avec le groupe par lequel il est arrivé, une liste nulle est tolérée, les lignes d'erreur remplacées sont écartées tandis que celles toujours en échec et sans clé sont conservées, rien d'autre n'est perdu, et le résumé ne compte que ce qui a été publié. Deux assertions antérieures se sont révélées mal parenthésées — l'une d'elles réussissait à tort — et ont été corrigées. **Les corrections des constats de cette exécution ne sont elles-mêmes pas encore vérifiées sur un tenant réel** |

### 2026-09-25 (14)
| Modification |
|--------|
| La deuxième exécution réelle de [`scripts/Reporting/Get-SharePointPermissionsReport.ps1`](scripts/Reporting/Get-SharePointPermissionsReport.ps1) s'est authentifiée proprement — le contrôle des rôles du jeton a détecté le délai de réplication dès sa première tentative et l'a attendu, SharePoint et Graph ont tous deux accepté leurs jetons, et 130 webs ont été découverts et leur analyse a commencé. Deux défauts au niveau de l'analyse se sont ensuite répétés sur chaque site |
| `ConvertTo-PermissionRows` rejetait une collection d'attributions de rôles vide : un paramètre `Mandatory [object[]]` refuse `@()`, donc chaque liste système ayant des permissions uniques mais plus aucune attribution de rôle (`User Information List`, `Converted Forms`, `Bibliotheek met onderhoudslogboeken`) échouait avec `Cannot bind argument to parameter 'RoleAssignments'`. Corrigé avec `[AllowEmptyCollection()]` — une étendue sans attributions ne produit légitimement aucune ligne |
| Plus lourd de conséquences, une liste d'attributions de rôles illisible était impossible à distinguer d'une liste vide. `Invoke-SPGet` avale les `403`/`404` et renvoie `$null`, que `Get-SPCollection` transforme en collection vide — et une collection vide se lit comme « personne n'a de droits sur cette étendue ». Les lectures des attributions de rôles utilisent désormais `-ThrowOnDenied`, de sorte qu'un refus devient une ligne d'erreur indiquant que les permissions sont inconnues plutôt qu'une affirmation silencieuse qu'il n'y en a aucune. C'est le seul endroit où un 403 n'est pas ignoré, car c'est le seul endroit où « pas autorisé à regarder » serait interprété à tort comme un constat |
| Les listes de galerie (`Galerie van thema's`, `Galerie met basispagina's`) répondaient `400 Bad Request` au `$select` des éléments, car leur schéma ne porte pas tous les champs qu'il nomme, et un 400 n'est pas quelque chose qu'une nouvelle tentative corrige. Le balayage des éléments descend désormais une échelle de quatre barreaux de clauses `$select` de plus en plus étroites jusqu'à ce que SharePoint en accepte une ; chaque barreau conserve `Id` et `HasUniqueRoleAssignments`, donc dans le pire des cas on perd un nom de fichier plutôt que les étendues uniques de la liste. Seul un 400 déclenche le rétrécissement — un refus, une limitation (throttling) ou un seuil d'affichage répond de la même façon, si peu de champs que l'on demande |
| Vérifié localement avec 132 contrôles répartis sur sept suites (21 nouveaux) : un ensemble d'attributions vide ne provoque plus de plantage et ne produit aucune ligne, une lecture refusée lève une exception avec « unknown rather than empty », un vrai `200` vide reste vide de bout en bout, un balayage ordinaire ignore toujours un 403, chaque barreau de l'échelle conserve les champs dont dépend l'analyse, et seul un 400 déclenche le rétrécissement. **La phase d'analyse au-delà du web 11 n'est toujours pas vérifiée sur un tenant réel** |

### 2026-09-25 (13)
| Modification |
|--------|
| La première exécution réelle de [`scripts/Reporting/Get-SharePointPermissionsReport.ps1`](scripts/Reporting/Get-SharePointPermissionsReport.ps1) a franchi SharePoint — l'identifiant par certificat a fonctionné et le contrôle préalable a signalé `SharePoint accepted the token (root web: ...)` — puis a échoué sur Graph avec `401` lors de la récupération des sites. Cause : le jeton Graph a été émis immédiatement après l'octroi des rôles d'application, avant que l'octroi ne soit répliqué, et ne portait donc aucune revendication `roles`. Graph répond à un tel jeton par `401`, pas par `403`, et comme le jeton était mis en cache pour toute son heure de validité, chacune des six nouvelles tentatives recevait le même jeton mort |
| Les jetons doivent désormais faire leurs preuves : `Get-ResourceToken` accepte `-RequiredRoles`, décode le JWT émis et refuse de mettre en cache un jeton dont la revendication `roles` ne contient pas ce dont l'exécution a besoin. Il continue d'en émettre de nouveaux (jusqu'à 15 tentatives, backoff plafonné à 20s) jusqu'à ce que l'octroi apparaisse, puis échoue en nommant le rôle manquant. Les deux jetons sont validés d'emblée — Graph pour `Sites.Read.All` + `GroupMember.Read.All`, SharePoint pour `Sites.FullControl.All` — afin qu'un délai de réplication soit attendu avant le début de l'analyse plutôt que découvert au bout de 130 sites |
| `Invoke-GraphGet` supprime désormais le jeton en cache et en émet un nouveau une fois sur un `401`, la même récupération dont `Invoke-SPGet` disposait déjà. Retenter seulement la requête n'aurait jamais pu fonctionner face à une entrée de cache empoisonnée |
| Une application `-ClientId` fournie par l'utilisateur n'est délibérément **pas** soumise à la validation des rôles : une application fonctionnelle peut détenir des rôles plus larges (`Directory.Read.All` au lieu de `GroupMember.Read.All`), et la rejeter serait un faux échec. Le contrôle préalable SharePoint détecte toujours une application réellement sous-autorisée |
| `Get-JwtClaim` renvoie `$null` pour un jeton vide au lieu de lever une erreur de liaison de paramètre, et `Disconnect-MgGraph` ne laisse plus fuiter son objet de contexte sous forme d'un tableau parasite `ClientId`/`TenantId`/`Scopes` après le résumé |
| Vérifié localement avec 111 contrôles répartis sur six suites (dont 23 nouveaux, pilotant le vrai `Get-ResourceToken` contre un faux point de terminaison de jetons) : un rôle qui arrive en retard est attendu et seul le jeton qui le porte est mis en cache, un rôle qui n'arrive jamais échoue bruyamment sans rien mettre en cache, le backoff augmente et reste plafonné, et la mise en cache reste isolée par ressource. **Le chemin Graph corrigé n'est pas encore vérifié sur un tenant réel** |

### 2026-09-25 (12)
| Modification |
|--------|
| Renforcement de [`scripts/Reporting/Get-SharePointPermissionsReport.ps1`](scripts/Reporting/Get-SharePointPermissionsReport.ps1) pour les longues exécutions à l'échelle du tenant. Un `401` pouvait encore être récupéré sous forme de ligne d'erreur par site : le gestionnaire par liste le relançait, mais le gestionnaire par web l'interceptait à nouveau, si bien qu'un identifiant cessant de fonctionner en cours d'exécution aurait écrit une ligne d'erreur pour chaque site restant — exactement le schéma de défaillance que la correction du certificat venait d'éliminer. Les deux gestionnaires laissent désormais passer un 401 et l'exécution s'arrête |
| Les recherches d'éléments en parallèle lisaient le jeton bearer une fois par bibliothèque au lieu d'une fois par vague. Une bibliothèque comptant suffisamment de portées uniques survit à la durée de vie d'un jeton, et la fin de son traitement aurait échoué sans aucune indication de la cause. Le jeton est désormais relu avant chaque vague |
| L'énumération des éléments ne matérialise plus une bibliothèque entière avant de filtrer. `Invoke-SPCollectionPaged` transmet chaque page à un callback et seuls les éléments ayant réellement leur propre portée sont conservés — une bibliothèque d'un million d'éléments coûte désormais une page de mémoire au lieu d'un million d'objets actifs |
| Ajout d'une protection de pagination : un SharePoint renvoyant un `nextLink` identique provoquait auparavant une boucle infinie sur un tenant réel ; ce cas est désormais détecté et interrompu |
| Les clés de point de contrôle ont été déplacées du fichier d'état JSON vers un `.keys.partial.log` en ajout seul. Réécrire une liste triée de toutes les clés terminées après chaque liste a un coût quadratique ; sur un tenant comptant des milliers de listes, le point de contrôle coûtait plus cher que l'analyse elle-même. Une dernière ligne tronquée par un processus tué est tolérée — cette unité est simplement réanalysée |
| Une écriture CSV qui échoue (le fichier partiel ouvert dans Excel) est retentée cinq fois, puis l'exécution s'arrête. Auparavant, l'erreur était levée alors que l'unité était déjà marquée comme terminée, et ces lignes disparaissaient définitivement du rapport |
| Une liste en échec ne coûte plus que cette liste, et non le reste du site : la gestion d'erreurs par liste écrit une ligne d'erreur, conserve ce que la liste a déjà produit et laisse délibérément l'unité non marquée pour qu'une exécution reprise la retente. Un balayage d'éléments en échec a sa propre ligne, car sans elle la liste semble simplement ne rien contenir avec des autorisations uniques |
| Les workers par élément ne signalent plus `401`/`403` comme un ensemble d'autorisations vide — seul `404` (élément réellement supprimé pendant l'analyse) signifie « aucune autorisation ». Affirmer qu'un élément illisible n'a aucun droit est pire que de le signaler |
| Ajout d'un test d'écriture dans le dossier de sortie avant l'authentification, d'un abandon lorsque la découverte ne trouve aucun site, et d'un `trap` qui supprime l'inscription d'application temporaire Full Control en cas d'erreur non gérée |
| Suppression du tableau de contexte `Set-MgRequestContext` qui fuyait vers stdout ; l'exécution se termine désormais en indiquant si chaque portée ciblée a été lue ou combien ont été manquées |
| Vérifié localement avec 53 contrôles répartis sur quatre suites : analyse des principaux/claims, cohérence du schéma des lignes CSV (six formes de ligne, 26 colonnes chacune), génération du certificat et de l'assertion signée, et comportement HTTP piloté via un faux transport — 401 abandonne après une seule réauthentification, 403/404 restent par objet, 429 est retenté jusqu'au succès, les boucles de pagination sont rompues, les fichiers verrouillés sont retentés et le journal de points de contrôle survit à une ligne tronquée. **Toujours pas vérifié sur un tenant réel** |

### 2026-09-25 (11)
| Modification |
|--------|
| Correction de [`scripts/Reporting/Get-SharePointPermissionsReport.ps1`](scripts/Reporting/Get-SharePointPermissionsReport.ps1) qui produisait un rapport vide sur un tenant réel : les 130 webs revenaient tous avec `[SKIP] Web not accessible with the current permissions`. L'App Registration temporaire s'authentifiait avec un secret client, et **SharePoint Online refuse tout jeton app-only obtenu avec un secret** — `401` avec `x-ms-diagnostics: ... Unsupported app only token`. Graph acceptait ce même identifiant, si bien que l'énumération des sites fonctionnait et que seuls les appels `_api` échouaient, d'où l'apparence d'un problème d'autorisations par site |
| L'application temporaire reçoit désormais un certificat au lieu d'un secret. Il est généré en mémoire avec `CertificateRequest`, enregistré comme `keyCredential` et utilisé pour signer une assertion client RFC 7523 — il ne touche jamais le magasin de certificats ni le disque, de sorte qu'une exécution interrompue ne laisse rien derrière elle |
| `Invoke-SPGet` n'avale plus `401` en même temps que `403`/`404`. Un `401` n'est jamais propre à un site — c'est la même réponse pour tout le tenant — et le traiter comme « ce site-là n'est pas accessible » est précisément ce qui transformait un unique défaut d'identifiant en 130 lignes qui ressemblaient à des constats. Il lève désormais une erreur, avec le motif `x-ms-diagnostics` et, dans le cas du secret, ce qu'il faut faire |
| Ajout d'une vérification préalable SharePoint : un appel sur la racine du tenant après la connexion, avant toute énumération. Que SharePoint accepte ou non l'identifiant est un simple oui/non pour toute l'exécution ; le savoir coûte donc désormais une requête au lieu d'un balayage complet |
| `-ClientSecret` avertit désormais au démarrage que la partie SharePoint échouera, et la documentation dit la même chose. Des alternatives reposant uniquement sur Graph ont été envisagées puis écartées : Graph n'a aucun endpoint pour les attributions de rôles web, les groupes SharePoint, les administrateurs de collection de sites ou les niveaux d'autorisation nommés, et nécessiterait un appel `/permissions` par élément au lieu d'un seul balayage `HasUniqueRoleAssignments` par liste |
| Suppression d'un tableau parasite `ClientTimeout RetryDelay MaxRetry` que `Set-MgRequestContext` affichait sur stdout à la fin de chaque exécution |
| Vérifié localement : 21 contrôles sur la génération du certificat et l'assertion signée, notamment que la signature se vérifie avec la clé publique du certificat, que `x5t` correspond à son hachage SHA-1 et que rien n'est écrit dans `Cert:\CurrentUser\My`. Le chemin d'authentification corrigé lui-même n'est **pas encore vérifié sur un tenant réel** |

### 2026-09-25 (10)
| Modification |
|--------|
| [`Convert-MarkdownToHtml.ps1`](scripts/Startup/Convert-MarkdownToHtml.ps1) affichait chaque liste numérotée sous forme de puces vides. `$Matches` est une seule variable par portée : la branche liste capturait le texte de l'élément, puis exécutait un second `-match` pour déterminer si la liste était ordonnée, et ce second match écrasait la capture. Les listes à tirets ne s'en sortaient que parce que leur second match échouait et laissait `$Matches` intact. Les deux captures proviennent désormais d'un seul match, et le marqueur détermine le type sans nouveau match |
| Un bloc de code délimité, indenté pour s'aligner à l'intérieur d'une étape numérotée, conservait cette indentation, de sorte que quiconque copiait la commande depuis la page copiait aussi les espaces de tête. L'indentation propre du délimiteur est désormais retirée de son contenu — et uniquement celle-là : un bloc délimité en colonne 0 conserve tous ses espaces, ce dont l'exemple de sortie du script a besoin |
| Vérifié sur la page régénérée : 36 éléments de liste et **aucun** vide, la page s'analyse toujours comme du XML, 29 tableaux et 19 blocs de code intacts, la commande indentée ressort propre, et les sept blocs de code qui commencent légitimement par des espaces le font toujours |

### 2026-09-25 (9)
| Modification |
|--------|
| Vous pouvez désormais cliquer depuis un readme directement vers le script qu'il décrit. Chaque nom de script dans le tableau `Scripts` d'un readme de dossier renvoyait vers une section plus bas sur la même page, jamais vers le fichier — le readme vous disait donc ce que faisait un script sans vous donner le moyen de l'ouvrir. 172 liens répartis sur 47 readmes pointent désormais vers le fichier, avec à côté un lien `([docs](#…))` vers la section qui existait auparavant |
| 34 d'entre eux n'étaient même pas des liens : un nom de script dans une cellule de tableau, entre backticks, sans rien derrière. Ce sont désormais aussi des liens vers le fichier |
| Ajout de [`scripts/Startup/Test-MarkdownLinks.ps1`](scripts/Startup/Test-MarkdownLinks.ps1), car des liens jamais vérifiés sont des liens qui pourrissent en silence. Il parcourt chaque `.md` et échoue sur deux choses : un lien relatif vers un fichier inexistant (les espaces encodés en pourcentage étant d'abord décodés, comme GitHub les sert), et une ancre sans titre correspondant. Les ancres sont résolues comme GitHub les construit, y compris le suffixe `-1`/`-2` pour les titres répétés |
| Il a trouvé trois ancres qui n'avaient jamais fonctionné : `#watch-rdslivesps1` avait un `s` de trop pour `### Watch-RDSLive.ps1`, et deux liens du readme Intune utilisaient `#detect--remediate-…` alors que le titre `### Detect- / Remediate-StuckWin32AppEnforcement.ps1` produit `#detect---remediate-…` — trois tirets, car la barre oblique disparaît et les espaces qui l'entourent deviennent chacun un tiret. Personne ne l'aurait repéré à l'œil |
| Les caractères invisibles sont retirés du titre comme du lien avant leur comparaison. Sans cela, les quatre entrées avec emoji de la table des matières racine apparaissaient comme cassées : le titre et le lien portent tous deux un sélecteur de variante, qui n'est ni une lettre ni un chiffre. Reproduire octet par octet le slugger de GitHub sur des caractères que personne ne voit n'est pas l'objectif — établir qu'un lien et un titre se correspondent l'est |
| Vérifié : 854 liens internes répartis sur 71 fichiers markdown se résolvent tous, les 181 fichiers s'analysent tous, et l'index des scripts est à jour avec 178 scripts. `L` ajouté au menu pour la vérification des liens, à côté de `X` pour l'index |

### 2026-09-25 (8)
| Modification |
|--------|
| La procédure IT Glue pour Teams existe désormais aussi sous forme de page HTML mise en forme, [`scripts/Device/Update-TeamsClient-ITGlue.html`](scripts/Device/Update-TeamsClient-ITGlue.html), à coller dans IT Glue ou à imprimer. Elle est **générée** par le nouveau [`scripts/Startup/Convert-MarkdownToHtml.ps1`](scripts/Startup/Convert-MarkdownToHtml.ps1) plutôt qu'écrite à la main : ce document a changé six fois en deux jours, et une copie faite à la main aurait été fausse dès le lendemain matin |
| Le convertisseur couvre ce que ces documents utilisent réellement — titres, tableaux, blocs de code délimités, y compris ceux indentés dans les étapes numérotées, citations, les deux types de listes, séparateurs, ainsi que code en ligne, gras, italique et liens — et laisse passer tout le reste sous forme de texte au lieu de deviner. `-Check` n'écrit rien et se termine avec `1` lorsque la page commitée a pris du retard sur son markdown, ce qu'un hook ou un pipeline appellerait |
| Les éléments vides sont émis auto-fermants, de sorte que la page s'analyse comme du XML aussi bien que du HTML. C'est ainsi qu'elle a été vérifiée plutôt qu'en la regardant : la sortie s'analyse et contient 29 tableaux, 190 lignes, 19 blocs de code et 46 titres, **sans** aucun tableau contenant une ligne dont la largeur diffère de son en-tête. Également contrôlé : aucun `**`, backtick ou `](` ne subsiste dans le texte rendu, et les caractères ✅/❌ et accentués sont préservés |
| Les deux chemins d'échec de `-Check` ont été exercés : une page qui n'existe pas encore, et un fichier markdown qui a évolué — chacun se termine avec `1` et le motif. La ligne de date de génération est exclue de la comparaison, de sorte qu'un document inchangé ne signale aucune différence |
| Ajouté comme touche de menu **M**, documenté dans [`scripts/Startup/readme.md`](scripts/Startup/readme.fr.md), et [`scripts/INDEX.md`](scripts/INDEX.md) régénéré — l'ajout d'un script l'avait rendu obsolète, ce que `Update-ScriptIndex.ps1 -Check` a signalé |

### 2026-09-25 (7)
| Modification |
|--------|
| Fusion des parties lisibles d'une ancienne version IT Glue de cette même procédure dans [`Update-TeamsClient-ITGlue.md`](scripts/Device/Update-TeamsClient-ITGlue.md) : la phrase unique indiquant ce que couvre la procédure, et la forme ✅/❌ pour « ce script convient-il à ce ticket », y compris les deux plaintes d'utilisateurs que cette version mentionnait et la nôtre non — Teams qui se bloque au démarrage et Teams qui se ferme de manière inattendue |
| Les deux versions divergeaient sur qui peut exécuter la mise à jour — l'ancienne la place au niveau 1, la nôtre au niveau 2 — ce qui sert moins bien un service desk que l'une ou l'autre réponse prise isolément. Le niveau global est remplacé par un tableau « wie mag wat » qui attribue un niveau par action : la vérification reste au niveau 1, la mise à jour et les réparations sont au niveau 2, et les trois switches qui ignorent la vérification de signature ou modifient la base de données Windows Installer sont au niveau 3. Modifier la politique d'escalade se résume désormais à un tableau, et non à une relecture du document |
| Délibérément non fusionné, en se mesurant au script plutôt qu'en jugeant à l'œil : le tableau des paramètres de cette version couvre 8 des 16 paramètres, affirme que Teams classique n'est jamais touché (`-RemoveClassicTeams` fait exactement cela), et montre deux lignes d'exemple de sortie que le script ne produit pas — `Found MSTeams ...` et `Teams installation completed` |
| Correction de numérotation : deux entrées portaient toutes deux le libellé `(4)`. La synchronisation IT Glue est plus récente que l'entrée d'index des scripts au-dessus, elle est donc désormais `(6)` et se place dans l'ordre où le travail a été effectué |

### 2026-09-25 (6)
| Modification |
|--------|
| Le document IT Glue a été remis en cohérence avec le script, en comparant son texte au bloc de paramètres plutôt qu'en le lisant : il lui manquait entièrement `-BootstrapperUrl` et `-SkipSignatureCheck`, et son tableau de variables NinjaOne omettait `removeWebRtcRedirector`, `clearOrphanedAddInRegistration`, `skipSignatureCheck` et `webRtcUrl`. Les trois documents couvrent désormais les 16 paramètres, et le tableau de variables les 16 variables d'environnement |
| Cette même comparaison a révélé une vraie lacune dans le script : chaque autre champ texte pouvait être défini depuis une variable NinjaOne, sauf `bootstrapperUrl`, qui n'était tout simplement jamais lu. Un administrateur qui l'aurait défini l'aurait vu ignoré en silence. Il est désormais lu, à côté de `webRtcUrl` |
| Deux affirmations du document IT Glue n'étaient plus vraies. « Il ne touche pas à Teams classique » n'est vrai que sans `-RemoveClassicTeams`, et le résumé en langage courant promettait encore que chaque copie du complément est supprimée lors d'une mise à jour - ce balayage est désormais conditionné à la possibilité d'installer un remplacement |
| Ajout d'une procédure de niveau 3 pour le blocage `1612` + `1638` rencontré sur l'hôte de production : ce que signifie chaque code, la commande en lecture seule qui indique si Windows Installer dispose encore de son MSI en cache, lequel des deux résultats nécessite `-ClearOrphanedAddInRegistration`, et la remarque selon laquelle démarrer Teams et redémarrer Outlook rend entre-temps le bouton de réunion à l'utilisateur |

### 2026-09-25 (5)
| Modification |
|--------|
| Trouver un script sur GitHub revenait à deviner dans lequel des 56 dossiers de workload il se trouvait et à ouvrir des readmes jusqu'à tomber dessus. Il existe désormais une page qui répond à cette question : [`scripts/INDEX.md`](scripts/INDEX.md) liste les 176 scripts de A à Z avec un lien vers le fichier, un lien vers le readme de son dossier et ce qu'il fait — Ctrl-F au lieu d'une chasse au trésor |
| La page est **générée**, par le nouveau [`scripts/Startup/Update-ScriptIndex.ps1`](scripts/Startup/Update-ScriptIndex.ps1), et ne peut donc pas s'écarter des fichiers comme le fait un tableau tenu à la main. `-Check` signale un index obsolète sans écrire (code de sortie `1`), ce qu'un hook ou un pipeline appellerait ; une exécution qui trouve la page à jour n'écrit rien du tout |
| Les descriptions proviennent des scripts eux-mêmes : le bloc `.SYNOPSIS`, réassemblé sur toutes les lignes sur lesquelles il s'étend plutôt qu'en prenant la première ligne, ce qui laissait dans le tableau des demi-phrases comme « Grant Full Access and/or Send As delegate rights on one mailbox, a CSV list of ». Lorsqu'un synopsis commence par une phrase puis énumère ses cas, l'introduction est conservée et la liste n'est pas entraînée derrière elle |
| Pour les anciens scripts sans `.SYNOPSIS`, un bloc de commentaires `#` en tête est utilisé à la place — mais uniquement un véritable en-tête. Une seule ligne de commentaire posée directement au-dessus du code décrit cette ligne, pas le script : `# URL van de theme` au-dessus d'une affectation `$ThemeUrl` était lu comme une description, ce qui est pire dans un tableau qu'une cellule vide |
| Les huit scripts qui se retrouvaient encore sans rien ont reçu un véritable `.SYNOPSIS` au lieu d'une cellule vide : [`add-lock.ps1`](scripts/Intune/Desktop/Add%20Lockscreen%20to%20start%20and%20desktop/add-lock.ps1), [`add-shortcut-lock.ps1`](scripts/Intune/Desktop/Add%20Lockscreen%20to%20start%20and%20desktop/add-shortcut-lock.ps1), [`logic-permissies.ps1`](scripts/Graph/logic-permissies.ps1), [`Test-OpenVpnDiagnostics.ps1`](scripts/Device/Test-OpenVpnDiagnostics.ps1), [`Deploy-OfficeTheme.ps1`](scripts/Custom%20Scripts/Intune/Desktop/Deploy-OfficeTheme.ps1), [`Restart-Time-Sync.ps1`](scripts/Device/Time%20sync/Restart-Time-Sync.ps1), [`Test-PowerShellSyntax.ps1`](scripts/Startup/Test-PowerShellSyntax.ps1) et [`functies.ps1`](scripts/Startup/functies.ps1). Les 176 scripts se décrivent désormais eux-mêmes, et l'index n'a donc plus de section « sans description » |
| Documentation des deux scripts qu'aucun readme ne mentionnait : [`Phising-rollout.ps1`](scripts/Entra/Phising-rollout.ps1) dans [`scripts/Entra/readme.md`](scripts/Entra/readme.fr.md) (la synchronisation bidirectionnelle entre le déploiement de la MFA résistante au phishing et les groupes enregistrés, ce qui compte comme enregistré et pourquoi le défaut est un filtre AAGUID) et [`Get-FSlogix-errors.ps1`](scripts/RDS/Get-FSlogix-errors.ps1) dans [`scripts/RDS/readme.md`](scripts/RDS/readme.fr.md) (ce que collecte le diagnostic FSLogix et le fait qu'il doit s'exécuter sur l'hôte de session). Son en-tête pointait encore vers un nom de fichier qui n'existe plus et nommait un vrai client dans l'exemple ; les deux ont été corrigés |
| Le tableau `Menu` racine s'était écarté de [`menu.ps1`](menu.ps1) — `I`, `T`, `S` et `P` manquaient. Synchronisé, et `X` ajouté pour le générateur d'index, qui figure aussi dans le readme Startup et dans l'arborescence du dépôt |
| Vérifié : les 176 fichiers s'analysent ; le générateur est idempotent (une seconde exécution indique « already up to date » et n'écrit rien) ; `-Check` se termine avec `0` lorsque l'index est à jour ; chaque lien markdown du dépôt se résout, noms de dossiers encodés en pourcentage compris ; et [`f.ps1`](f.ps1) trouve toujours le nouveau script comme ceux nouvellement décrits |
| Correction de numérotation : deux entrées ci-dessous portaient toutes deux le libellé `(3)`. Renumérotées dans l'ordre où le travail a réellement été effectué |

### 2026-09-25 (4)
| Modification |
|--------|
| Correction de l'entrée précédente : le `1638` sur le complément de réunion n'était **pas** causé par `-Force` rétrogradant le client. Mesuré sur l'hôte lui-même, le complément enregistré était en `1.25.28902` et le MSI en cours d'installation en `1.26.21803` - plus récent, et pourtant refusé. Ce MSI refuse de s'installer tant qu'une autre copie du complément est enregistrée, quelle qu'en soit la version. La comparaison de versions ajoutée lors de la dernière modification n'aurait donc pas empêché l'échec ; la protection vérifie désormais si un enregistrement a survécu à la désinstallation, ce qui est réellement déterminant |
| `-ClearOrphanedAddInRegistration` est la porte de sortie de l'état dans lequel se trouve cet hôte. Une désinstallation qui répond `1612` signifie que Windows Installer a perdu le MSI en cache dont il a besoin et ne peut plus supprimer le produit par aucun moyen pris en charge, tandis que son enregistrement continue de refuser toute réinstallation. Le switch fait oublier ce seul produit au programme d'installation : ses clés sous `Installer\Products`, `Installer\Features` et `Installer\UserData\S-1-5-18\Products`, son entrée sous le code de mise à niveau, et l'entrée Programmes et fonctionnalités. Ce que faisait MsiZap, limité à un seul produit, uniquement après que msiexec a prouvé qu'il n'y parvient pas, désactivé par défaut, et chaque clé passant par `ShouldProcess` |
| Trouver ces clés nécessite le ProductCode sous forme du GUID « compacté » de 32 caractères de Windows Installer. Cette transformation a été validée avant que quoi que ce soit ne l'utilise pour désigner des clés à supprimer : sur 57 entrées de désinstallation nommées par GUID sur un poste de travail, les 32 disposant de données produit à l'échelle de la machine correspondaient toutes à une clé compactée existante avec un `DisplayName` identique, et les 25 autres sont des installations par utilisateur situées sous le SID de l'utilisateur |
| Vérifié en lecture seule sur trois produits réels : chacun produit trois clés produit plus exactement une entrée de code de mise à niveau, et le chemin construit est lisible tel qu'écrit. Un code produit fictif ne renvoie rien et un produit inconnu ne renvoie aucune clé, de sorte que le nettoyage ne peut pas se déclencher dans le vide. **Non testé :** la suppression elle-même, et la réinstallation qui devrait la suivre |

### 2026-09-25 (3)
| Modification |
|--------|
| Ajout de [`scripts/Reporting/Get-SharePointPermissionsReport.ps1`](scripts/Reporting/Get-SharePointPermissionsReport.ps1) — un rapport exhaustif en lecture seule des autorisations SharePoint Online : administrateurs de collection de sites, attributions de rôles web y compris les ruptures d'héritage, groupes SharePoint avec leur appartenance complète, attributions de rôles des listes et bibliothèques, chaque dossier et élément ayant une portée unique, liens de partage avec leur type, principaux externes/invités, octrois à `Everyone`, et octrois à des groupes Entra résolus en appartenance transitive. Quatre CSV : détail, synthèse par site, appartenance aux groupes et — derrière `-IncludeEffectiveAccess` — une ligne par utilisateur résolu et par portée, avec le groupe par lequel passe l'accès |
| L'héritage est suivi comme SharePoint le modélise : un élément n'est signalé comme portée propre que lorsque `HasUniqueRoleAssignments` est vrai, si bien que le CSV est une carte de la structure des autorisations plutôt qu'une ligne par fichier. La découverte des sites est volontairement redondante — Graph `getAllSites`, puis les sous-sites via Graph et via SharePoint REST (`/_api/web/webs`), dédoublonnés sur l'URL — car Graph omet les sous-webs classiques |
| L'authentification a dû passer en app-only : les attributions de rôles ne sont pas du tout lisibles via Graph, et ne sont pas non plus couvertes par les rôles d'application Read/Write/Manage de SharePoint — seul `Sites.FullControl.All` permet de les énumérer. Le script se connecte une fois de manière interactive, crée une App Registration de courte durée avec ce rôle plus Graph `Sites.Read.All` et `GroupMember.Read.All`, et la supprime à nouveau à la sortie. Malgré le rôle Full Control, il n'émet jamais que des `GET` : il n'écrit jamais et ne modifie jamais une autorisation. `-ClientId`/`-TenantId` avec un secret ou un certificat évite l'application temporaire |
| Reprise possible comme pour les autres longues analyses SharePoint : un point de contrôle par liste terminée, indexé sur un hachage des paramètres d'analyse, de sorte qu'une exécution interrompue sur le tenant reprend au lieu de recommencer ; `-Restart` le supprime. Les fichiers de point de contrôle ne sont nettoyés qu'une fois les CSV finaux écrits, si bien que leur présence est en soi le signal qu'une exécution a été interrompue |
| Documenté dans [`scripts/Reporting/readme.md`](scripts/Reporting/readme.md) (couverture, authentification, les quatre fichiers de sortie, points de contrôle, tableau complet des paramètres, exemples), ajouté à l'arborescence racine du dépôt et à [`menu.ps1`](menu.ps1) sous Reporting en tant que `P` — qui demande aussi s'il faut analyser le tenant ou un seul site, la portée, et s'il faut écrire le CSV d'accès effectif |
| L'arborescence du dépôt ne listait ni ce script ni [`Remove-SharePointFileVersionsByDate.ps1`](scripts/Reporting/Remove-SharePointFileVersionsByDate.ps1) ; les deux y figurent désormais |
| **Non testé par moi** : ce script n'a pas été exécuté sur un tenant réel lors de cette session — seule sa syntaxe a été vérifiée. Le chemin de l'App Registration temporaire, les nouvelles tentatives en cas de limitation et la reprise par point de contrôle ne sont pas vérifiés ici et doivent être éprouvés sur un tenant pilote, en commençant par `-SiteUrl` et `-Scope Site`, avant une exécution à l'échelle du tenant |

### 2026-09-25 (2)
| Modification |
|--------|
| Un complément enregistré mais dont les fichiers ont disparu ne compte plus comme installé. C'est exactement l'état laissé par l'exécution en échec ci-dessus, et le script aurait répondu « Teams is up to date - nothing to do » : la décision de travail ne regardait que l'entrée Programmes et fonctionnalités, tandis que le contrôle de DLL qui repère ce cas se contentait de le signaler |
| [`Update-TeamsClient.ps1`](scripts/Device/Update-TeamsClient.ps1) ne supprime plus un complément de réunion fonctionnel avant de savoir qu'il peut installer un remplacement. Une exécution `-Force` en production a supprimé toutes les copies à l'étape 6 puis a échoué à l'étape 8 avec `1638`, laissant l'hôte de session sans aucun complément. Le balayage a été déplacé à l'étape 8, derrière la comparaison de versions : un MSI plus ancien que le complément enregistré signifie désormais que le balayage et l'installation sont ignorés et que le complément fonctionnel est laissé exactement tel quel |
| Trois éléments devaient concorder pour cela, et les trois sont désormais gérés. `-Force` sur un hôte dont le build est plus récent que celui publié est une **rétrogradation**, que le contrôle de version signale désormais nommément. Une désinstallation du complément qui répond `1612` signifie que Windows Installer a perdu sa source ; elle est donc retentée avec son propre MSI en cache sous `C:\Windows\Installer` (via `Installer\UserData\S-1-5-18\Products\*\InstallProperties`, `LocalPackage`) ; lorsque celui-ci a aussi disparu, l'exécution indique que l'enregistrement ne peut pas être supprimé et ce que cela entraînera. Un `1638` sur le complément est désormais un avertissement plutôt qu'un abandon, de sorte que la vérification s'exécute quand même et rapporte ce dont Outlook dispose réellement |
| La vérification préalable affiche la version du complément enregistré au lieu d'un simple « is installed » - ce seul chiffre constituait tout le diagnostic de l'échec, et c'était la seule chose qui n'apparaissait pas à l'écran |
| La ligne AppLocker n'affiche plus de résumé vide lorsque `SrpV2` existe sans aucune collection de règles en dessous (comme sur l'hôte de production) : elle indique « no rule collections configured, so it blocks nothing » |
| Vérifié : la recherche du package en cache sur des produits réellement installés, et le fait qu'interroger un produit absent de cette machine ne renvoie rien sans lever d'erreur ; la protection de version dans les quatre combinaisons (plus ancien, plus récent, égal, non analysable) ; la ligne AppLocker à collection vide. **Non testé :** la nouvelle tentative `msiexec /x <cached msi>` et le chemin d'avertissement `1638` sur un hôte réel |

### 2026-09-25
| Modification |
|--------|
| [`Update-TeamsClient.ps1`](scripts/Device/Update-TeamsClient.ps1) cesse de crier au loup à propos d'AppLocker. Le contrôle des bloqueurs SlimCore avertissait dès que `HKLM:\SOFTWARE\Policies\Microsoft\Windows\SrpV2` existait — ce qui est le cas sur tout parc ayant un jour écrit une seule règle Exe — si bien que l'avertissement se déclenchait sur des machines où rien n'était bloqué. Un contrôle qui se déclenche toujours est un contrôle que personne ne lit |
| La stratégie est désormais lue au lieu d'être simplement détectée, sur les trois points qui déterminent si elle peut bloquer le MSIX : seule la collection des applications empaquetées (`Appx`) s'applique, car un MSIX ne rencontre jamais les règles `Exe`/`Msi`/`Script`/`Dll` ; une collection contenant des règles dont l'application est *non configurée* est appliquée malgré tout, selon Microsoft, et seul un `EnforcementMode = 0` explicite laisse tout passer ; et rien n'est appliqué tant que le service Identité de l'application (`AppIDSvc`) est arrêté, ce qui est désormais dit explicitement plutôt que supposé dans un sens ou dans l'autre |
| Le rapport est exploitable par un technicien : le chemin de registre, le mode par collection, l'état du service et les cinq premiers noms de règles `Appx` avec leur action. Une règle qui autorise déjà les packages par leur nom est signalée comme `[ OK ]` ; une règle autorisant tout ce qui est signé par `O=MICROSOFT CORPORATION` est signalée comme probablement suffisante, avec une note invitant à vérifier qu'elle n'a pas été restreinte à un seul nom de produit |
| Une stratégie trouvée sur un hôte de session est désormais mentionnée sous forme de ligne de référence `[SKIP]` au lieu d'être masquée : elle n'y bloque rien, car la préparation se fait sur le poste client, mais il s'agit généralement du même GPO — ce qui vaut la peine d'être vérifié, c'est donc s'il atteint aussi les postes clients |
| Chaque branche a été exercée sur une arborescence de stratégie simulée : une collection `Exe` appliquée sans collection `Appx` ne produit aucun bloqueur (l'ancien faux positif, disparu) ; une collection `Appx` appliquée sans règle d'autorisation correspondante avertit avec le chemin et le nombre de règles ; les règles d'autorisation SlimCore explicite et éditeur Microsoft produisent leurs deux notes distinctes ; `EnforcementMode = 0` se lit comme audit uniquement ; application non configurée avec règles se lit comme appliquée ; sept règles en affichent cinq et `... and 2 more` ; une clé `SrpV2` absente ne produit rien. Confirmé silencieux sur cette machine, qui n'a aucune stratégie AppLocker. **Non testé** sur une stratégie AppLocker réellement appliquée sur un poste client réel |
| Limitation connue, documentée plutôt que masquée : la correspondance des règles d'autorisation est une correspondance textuelle sur le XML de la règle, de sorte qu'une règle large ne nommant ni Microsoft ni les packages (`PublisherName="*"`) laisserait réellement passer SlimCore mais est tout de même signalée comme bloqueur. Les noms de règles affichés à côté permettent de trancher |

### 2026-09-24
| Modification |
|--------|
| [`Update-TeamsClient.ps1`](scripts/Device/Update-TeamsClient.ps1) répond à la question que l'inventaire ne pouvait pas trancher : le preflight lit désormais les événements `Microsoft Teams VDI` du journal Application sur tout hôte de session — pas seulement avec `-AvdOptimizations` — et traduit les codes d'après la table des erreurs de connexion de Microsoft, si bien qu'un simple `-CheckOnly` indique si les utilisateurs sont réellement optimisés, et plus seulement si les composants sont installés |
| `24002`/`24010` indiquent que l'utilisateur est sur SlimCore, `16002` qu'un poste client n'a toujours pas de plugin, `16389`/`10083`/`1951` qu'une stratégie sur le poste client bloque le MSIX. Un `errc` à zéro est volontairement absent de la table : il signifie que cette phase n'a levé aucune erreur, et afficher « OK » à côté d'un véritable échec dans l'autre phase serait un mensonge |
| La requête utilise `-FilterXPath`, car `Get-WinEvent -FilterHashtable @{ ProviderName = ... }` lève une exception lorsque le fournisseur n'a jamais écrit d'événement — ce qui est le cas normal sur une machine saine hors VDI. Mesuré : 357 ms et une erreur non bloquante en cas d'absence, 104 ms en cas de présence |
| Le nouveau `-RemoveWebRtcRedirector` supprime l'ancienne optimisation, retirée le 1er octobre 2026. Il est mutuellement exclusif avec `-AvdOptimizations` et refusé avant l'invite UAC, réutilise le chemin `msiexec /x` + entrée `1605` obsolète éprouvé pour Teams classique, et laisse `IsWVDEnvironment` en place car SlimCore a lui aussi besoin de ce flag. Désactivé par défaut : un poste client incapable d'utiliser SlimCore et qui ne trouve plus le redirector se rabat silencieusement sur le rendu des médias sur l'hôte de session |
| Les deux documentations corrigées là où elles demandaient encore à un technicien de chercher SlimCore sur l'hôte de session |

### 2026-09-20 (8)
| Modification |
|--------|
| « Le complément ne se charge toujours pas » obtient désormais une réponse plutôt qu'un simple état. Un enregistrement présent mais qui ne se charge pas est confronté aux trois causes qui ne laissent aucune trace dans `LoadBehavior` lui-même, chacune signalée par une ligne `why:` : une incompatibilité d'architecture (bitness) entre Outlook et le chargeur enregistré, un complément qu'Outlook a mis de côté dans ses listes de résilience `DisabledItems`/`CrashedAddins`, et une stratégie de groupe qui écrase le comportement de chargement choisi par l'utilisateur |
| La vérification de résilience décode les valeurs binaires dans la ruche de cet utilisateur et effectue la correspondance sur le chemin du complément ; elle signale ainsi la seule cause qu'un technicien ne peut absolument pas voir depuis `LoadBehavior` — Outlook désactive un complément qui a planté et le maintient désactivé, ce qui explique pourquoi recocher la case ne tient pas |
| Lorsque rien sur la machine ne le bloque, le script le dit aussi, ce qui est également une réponse : il reste un redémarrage complet d'Outlook et un utilisateur qui s'est connecté à Teams au moins une fois |
| Les deux détections ont été éprouvées : un chemin de chargeur x86 face à l'Office x64 de ce poste produit le motif d'architecture, et une valeur binaire `CrashedAddins` implantée est décodée et signalée. Un enregistrement sain ne produit aucune ligne `why:` |

### 2026-09-20 (7)
| Modification |
|--------|
| Une réinstallation complète supprime désormais **toutes** les copies du complément de réunion avant d'installer la nouvelle, et pas seulement celle que connaît le MSI : le dossier à l'échelle de la machine, les dossiers par profil sous `%LOCALAPPDATA%\Microsoft\TeamsMeetingAdd-in`, et les enregistrements COM par utilisateur dans chaque ruche chargée |
| Cela boucle la boucle sur le `LoadBehavior 2` que ce script traque depuis deux jours. Une copie laissée derrière lors d'une réinstallation est précisément ce qui devient un enregistrement par utilisateur qui masque le nouvel enregistrement à l'échelle de la machine tout en pointant vers des fichiers qui n'existent plus — c'est ainsi que `admin` s'est retrouvé enregistré sur le complément `1.24.19202` |
| `Get-TeamsAddInFolder` vérifié sur cet appareil : il trouve la véritable copie par profil. Le plan `-WhatIf` montre la suppression du dossier et des deux vues CLSID avant la réinstallation. La suppression elle-même réutilise des mécanismes déjà éprouvés en conditions réelles dans les tests de Teams classique et de réparation, mais le nettoyage dans son ensemble s'exécute pour la première fois sur un hôte de production |

### 2026-09-20 (6)
| Modification |
|--------|
| Correction d'une vérification qui cherchait au mauvais endroit : le script avertissait `SlimCore packages not found` sur les hôtes de session, mais Microsoft déploie SlimCore **sur le poste client**, pas sur la VM — *"Step 3: SlimCore MSIX staging and registration on the endpoint ... the plugin silently executes this step, without user or admin intervention"*. L'avertissement n'était que du bruit sur chaque hôte de session, et une vérification qui cherche au mauvais endroit n'échoue pas, elle ment |
| Le rapport tient désormais compte du contexte. Sur un hôte de session, il confirme le build de Teams par rapport au minimum documenté `24193.1805.3040.8975` et précise que SlimCore relève du poste client. Sur un poste client, il indique si les paquets sont déployés et vérifie les trois stratégies que Microsoft documente comme bloquant ce déploiement, chacune avec le code d'erreur Teams qu'elle fait apparaître : `BlockNonAdminUserInstall` (16389), `AllowAllTrustedApps` (15615) et AppLocker (10083) |
| La question de version de la semaine dernière est également tranchée : Windows App pour Windows `2.0.352.0` est le minimum documenté sur le poste client, et le client Bureau à distance classique n'est plus du tout pris en charge pour cela |
| Résilience : le code de sortie MSI `1641` (succès, redémarrage déjà lancé) était compté comme un échec et interrompait l'exécution. C'est désormais un succès avec redémarrage signalé, au même titre que `3010` |
| Vérifié des deux côtés : ce poste client signale `Microsoft.Teams.SlimCoreVdiHost.win-x64 2026.31.1.16` ; avec un `RDInfraAgent` simulé, c'est le libellé côté hôte de session qui apparaît ; les trois bloqueurs ont été éprouvés sur des lectures de registre simulées |

### 2026-09-20 (5)
| Modification |
|--------|
| Une exécution de production propre a confirmé trois correctifs antérieurs sur un véritable hôte de session : le redirector réparé sur place (`The download is the installed version (1.56.2603.20001)`, donc l'exécution précédente a bien fait la mise à niveau 1.54 → 1.56), le complément résolu à partir du paquet déployé après le provisionnement, et l'ensemble du flux terminé avec le code de sortie `0` |
| Elle a aussi permis de cerner le seul avertissement restant : `BAKKERPARTNERS\admin` possède un enregistrement pointant vers le complément `1.24.19202`, une copie par utilisateur disparue depuis longtemps, qui masque un `1.26.21803` à l'échelle de la machine parfaitement sain. `-RepairOutlookAddIn` (variable Ninja `repairOutlookAddIn`) supprime désormais cette clé obsolète `Classes\CLSID\{19A6E644-...}` et remet `LoadBehavior` à 3, de sorte que COM résout à nouveau vers l'enregistrement à l'échelle de la machine |
| Il n'agit que lorsque cet enregistrement à l'échelle de la machine est sain — supprimer l'enregistrement masquant sans rien derrière laisserait l'utilisateur dans une situation pire — et il est désactivé par défaut, car il écrit dans la ruche d'un autre utilisateur. Il compte comme du travail, donc `-CheckOnly` le signale et `-Quiet` le fait remonter |
| Testé sur des clés implantées dans les deux vues du registre : `-WhatIf` planifie les deux actions, une exécution appliquée supprime les clés CLSID, règle `LoadBehavior` sur 3 et se termine avec `0`. Non testé : si Outlook charge ensuite réellement le complément pour cet utilisateur — ce sera la prochaine exécution de production |

### 2026-09-20 (4)
| Modification |
|--------|
| « Je ne le vois pas encore chargé sur tous les profils » relevait d'un manque de visibilité, et pas seulement d'un problème Teams : un profil dont la ruche n'est pas montée ne peut pas être lu du tout, et le script l'omettait simplement — si bien qu'un profil illisible et un profil sain paraissaient identiques dans la sortie. Il liste désormais ces profils par nom, avec ce que cela implique pour eux : avec un enregistrement sain à l'échelle de la machine, ils récupèrent le complément au premier démarrage d'Outlook ; sans cet enregistrement, il n'y a rien sur quoi se rabattre |
| Cela mérite d'être dit clairement, car c'est ce qui détermine s'il y a quelque chose à corriger : un profil non connecté n'est pas cassé. L'enregistrement à l'échelle de la machine couvre les utilisateurs qui n'ont aucun état par utilisateur ; seul un utilisateur qui a déjà le sien (désactivé, ou pointant vers une DLL supprimée) continue de le masquer |
| Les deux messages vérifiés sur des listes de profils simulées. Non vérifié ici : monter une ruche non montée pour inspecter ou réparer un profil déconnecté — `reg load` requiert des privilèges dont ce poste de travail ne dispose pas, ce mécanisme n'est donc volontairement pas construit sur une hypothèse non testée |

### 2026-09-20 (3)
| Modification |
|--------|
| Réponse à une question que le script ne pouvait pas trancher : **pourquoi** un compte affiche `LoadBehavior 2`. Outlook résout le complément via `Classes\CLSID\{19A6E644-...}\InprocServer32`, et un enregistrement par utilisateur dans `HKCU\SOFTWARE\Classes` l'emporte sur celui à l'échelle de la machine — un utilisateur continue donc de charger la copie de son propre profil même après qu'une installation `ALLUSERS=1` a atterri dans `Program Files (x86)`. Mesuré sur un appareil : la classe se résout en `%LOCALAPPDATA%\Microsoft\TeamsMeetingAdd-in\<version>\x64\Microsoft.Teams.AddinLoader.dll` |
| Le script résout désormais ce chemin pour chaque utilisateur connecté et distingue les deux cas, car ils demandent des correctifs différents : le complément désactivé mais sa DLL présente (recocher la case) contre un enregistrement pointant vers une DLL disparue (recocher ne tiendra pas — il faut le réinstaller pour cet utilisateur). La recherche ciblée dans les ruches chargées coûte ~100 ms |
| Testé avec un enregistrement implanté pointant vers une DLL manquante, sans toucher au véritable enregistrement |

### 2026-09-20 (2)
| Modification |
|--------|
| Troisième échec en production sur le même hôte de session, troisième correctif : `Uninstall of Teams Machine-Wide Installer failed (exit code 1605)`. 1605 signifie « cette action n'est valide que pour les produits actuellement installés » — l'entrée dans Programmes et fonctionnalités a survécu au produit, ce qui est courant une fois que le nouveau bootstrapper Teams est passé sur une machine |
| Il n'y a rien à désinstaller dans ce cas, mais l'entrée obsolète ferait signaler Teams classique par le script à chaque exécution ; il supprime donc désormais l'entrée de registre à la place et poursuit. Testé en conditions réelles : un véritable `msiexec /x` sur un code produit inconnu renvoie 1605, l'exécution avertit, supprime une entrée obsolète implantée, vérifie que tout est propre et se termine avec `0` |
| Les objets d'entrée de désinstallation portent désormais leurs `RegistryPath` et `UninstallString`, ce qui rend ce nettoyage possible |

### 2026-09-20
| Modification |
|--------|
| Correction du deuxième échec en production sur un hôte de session : `WebRTC Redirector install failed (exit code 1638)`. Ce MSI conserve un même ProductCode d'une version à l'autre, de sorte que `msiexec /i` sur une installation existante refuse avec « une autre version de ce produit est déjà installée » au lieu de faire la mise à niveau — et `-Force` tombe droit dedans sur tout hôte où le redirector est déjà présent |
| Le script lit désormais le ProductVersion du fichier téléchargé et décide : même version → réparation sur place (`REINSTALL=ALL REINSTALLMODE=vomus`), version différente → désinstallation de l'ancienne d'abord, puis installation. Un `1638` qui passerait encore entre les mailles est signalé comme « l'existant est laissé en place » au lieu de faire échouer toute l'exécution |
| Mesuré lors de la correction : `aka.ms/msrdcwebrtcsvc/msi` sert désormais `1.56.2603.20001`, alors que cet hôte avait `1.54.2408.19001` installé — il s'agissait donc d'une mise à niveau refusée, et non d'une installation en double. Les deux chemins sont correctement planifiés sous `-WhatIf` ; aucun des deux appels msiexec n'a encore été exécuté pour de vrai, ce que la documentation dit explicitement |
| Également consigné noir sur blanc : un redirector installé n'est **pas** mis à niveau silencieusement par une exécution normale. Seul `-Force` le remplace. Avec la fin du support de WebRTC le 1er octobre 2026, garder ce choix délibéré vaut mieux que de mettre à niveau automatiquement un composant en voie de disparition |

### 2026-09-18 (4)
| Modification |
|--------|
| [`Get-DistributionGroupMembers.ps1`](scripts/Exchange/Get-DistributionGroupMembers.ps1) — `-Member "*.verizon.com"` correspond désormais à un domaine **et à tous ses sous-domaines** (`.verizon.com` et `*@*.verizon.com` sont équivalents). Sans le `*.` initial, le filtre reste limité à ce seul domaine, de sorte que `@be.verizon.com` n'atteint toujours pas, volontairement, `@us.verizon.com` |
| L'exécution indique lequel des deux elle applique — *"scanning N list(s) for members on verizon.com and its subdomains"* — car un filtre dont vous devez deviner la portée est un filtre auquel vous ne pouvez pas vous fier dans un rapport client |
| La correspondance porte sur le libellé de domaine complet, vérifié contre `@notverizon.com` et l'astuce du suffixe `@verizon.com.evil.test` ; aucun des deux ne correspond lors d'une exécution `*.verizon.com`. Un caractère générique placé ailleurs qu'en tête est échappé plutôt que d'élargir discrètement le filtre |

### 2026-09-18 (3)
| Modification |
|--------|
| [`Get-DistributionGroupMembers.ps1`](scripts/Exchange/Get-DistributionGroupMembers.ps1) — **`-Recurse`**, après avoir vérifié si le rapport couvrait vraiment tout le monde : ce n'était pas le cas. Exchange ne renvoie jamais que les membres *directs*, de sorte qu'une liste contenant une autre liste signalait cette liste comme un seul membre, jamais les personnes qui s'y trouvent. Une personne qui ne reçoit le courrier que via un groupe imbriqué était invisible, et `-Member` signalait « aucun résultat » sur une liste qui lui distribue bel et bien le courrier — une réponse fausse qui a l'air sûre d'elle |
| `Via groep` nomme le groupe par lequel une personne est arrivée (vide pour un membre direct), et une personne joignable par plusieurs chemins obtient une seule ligne avec les chemins réunis plutôt qu'une ligne par chemin |
| `Aantal leden` continue de compter les membres directs, car c'est le nombre qu'affichent Exchange et l'EAC ; le nouveau `Aantal personen` compte les destinataires réellement atteints |
| Un groupe déjà développé n'est pas développé à nouveau, ce qui empêche aussi un cycle d'appartenance (A contient B, B contient A) de boucler indéfiniment. Vérifié sur une paire de listes de test volontairement cycliques ; une imbrication au-delà de 20 niveaux est signalée et laissée telle quelle |
| Documenté ce que le rapport ne couvre *toujours pas* : il lit l'appartenance aux groupes, donc un utilisateur qui ne figure sur aucune liste n'apparaît nulle part |

### 2026-09-18 (2)
| Modification |
|--------|
| [`Get-DistributionGroupMembers.ps1`](scripts/Exchange/Get-DistributionGroupMembers.ps1) — `-Member` accepte désormais aussi un **domaine** : `-Member "@be.verizon.com"` signale toutes les listes qui contiennent encore une adresse de ce domaine (`be.verizon.com` et `*@be.verizon.com` sont équivalents). Une adresse est mise en correspondance par Exchange lui-même ; un domaine ne peut pas l'être, donc chaque liste est lue puis filtrée — plus lent, et documenté comme tel |
| La correspondance couvre l'adresse principale, chaque alias **et `ExternalEmailAddress`**. C'est tout l'intérêt pour un domaine partenaire : un tel membre est généralement un contact de messagerie dont l'adresse SMTP principale est `...@contoso.onmicrosoft.com`, le véritable `@be.verizon.com` ne figurant que dans son adresse externe. Une correspondance sur l'adresse principale n'aurait rien trouvé et aurait annoncé « aucun » sans sourciller |
| Nouvelle colonne `Extern adres` dans la feuille `Leden`, afin que l'adresse qui reçoit réellement le courrier soit visible pour les contacts, au lieu du seul espace réservé interne |
| Avec un filtre actif : `Treffers` par liste dans `Overzicht`, et `Treffer op` par membre dans `Leden`. `Treffer op` contient l'**adresse** correspondante, et non Ja/Nee — sinon, une correspondance sur un alias serait inexplicable dans un rapport qui n'affiche pas les alias |
| Un filtre de domaine qui ne correspond à rien le dit et n'écrit aucun fichier, plutôt que de livrer un classeur vide qui ressemble à un export raté |

### 2026-09-18
| Modification |
|--------|
| Ajout de [`Get-DistributionGroupMembers.ps1`](scripts/Exchange/Get-DistributionGroupMembers.ps1) — toutes les listes de distribution avec leurs membres dans un seul classeur Excel : une feuille `Overzicht` (une ligne par liste) et une feuille `Leden` (une ligne par membre), toutes deux sous forme de tableaux filtrables avec une ligne d'en-tête figée. Les en-têtes de feuille et les types de destinataires sont en néerlandais, car c'est le classeur que lit le client |
| `-Member user@domain` répond à « sur quelles listes figure cette personne ? » côté serveur via `Get-Recipient -Filter "Members -eq '<DN>'"` au lieu de parcourir chaque groupe, et exporte malgré tout les listes trouvées en entier pour que le client voie qui d'autre y figure |
| `-IncludeDynamic` et `-IncludeM365Groups` élargissent le rapport au-delà des simples groupes de distribution ; les groupes dynamiques sont évalués en direct, puisqu'ils ne stockent aucune appartenance interrogeable |
| Se rabat sur deux fichiers CSV lorsque `ImportExcel` est absent (et propose de l'installer d'abord), de sorte qu'un module manquant ne vous coûte jamais le rapport. `ImportExcel` ajouté à [`Install-Modules.ps1`](scripts/Startup/Install-Modules.ps1) — `vias_archiver.ps1` en avait déjà besoin |
| Sous-menu Exchange : `K` Get-DLMembers |

### 2026-09-17 (3)
| Modification |
|--------|
| Correction du bug révélé par une exécution de production sur un hôte de session AVD : `teamsbootstrapper.exe -p` **provisionne** le paquet pour les futures connexions, il ne l'installe pas pour la personne qui a lancé le script. Le bootstrapper signalait un succès, puis l'étape du complément échouait sur `New Teams package not found after install`, car `Get-AppxPackage -Name MSTeams` interroge l'utilisateur courant et l'administrateur qui exécutait le script n'avait pas Teams |
| Le MSI du complément est désormais trouvé en parcourant `%ProgramFiles%\WindowsApps\MSTeams_*_x64__8wekyb3d8bbwe\MicrosoftTeamsMeetingAddinInstaller.msi` et en prenant la version la plus récente, de sorte qu'il fonctionne qu'un utilisateur ait ou non le paquet installé. Retesté pour une installation par utilisateur, un hôte uniquement provisionné et une exécution `-Force` |
| Même angle mort par utilisateur dans la vérification SlimCore, qui signalait « introuvable » sur un hôte où le nouveau Teams est installé pour un profil : elle interroge désormais `-AllUsers` en premier. Et les deux vues de registre Outlook à l'échelle de la machine sont étiquetées 64 bits/32 bits pour les distinguer, car cette exécution affichait deux lignes identiques `all users (machine-wide)`, ce qui ressemble à un bug |
| Cette exécution a aussi prouvé l'utilité des nouvelles vérifications : elle a supprimé un véritable Teams classique par profil `1.4.00.11161`, et trouvé `LoadBehavior 2` pour un compte - Outlook avait désactivé le complément, ce qu'aucune réinstallation ne corrige |

### 2026-09-17 (2)
| Modification |
|--------|
| [`Update-TeamsClient.ps1`](scripts/Device/Update-TeamsClient.ps1) peut désormais aussi supprimer Teams classique, via `-RemoveClassicTeams` (variable Ninja `removeClassicTeams`). Désactivé par défaut : retirer une application aux utilisateurs n'est pas une décision qu'une tâche de mise à jour doit prendre seule. Il désinstalle le *Teams Machine-Wide Installer* via msiexec — celui qui compte, car tant qu'il est présent, Windows continue de déployer Teams classique dans chaque nouveau profil — et, par profil, supprime la racine d'installation, l'entrée de démarrage automatique `Run\com.squirrel.Teams.Teams` et la clé obsolète `Uninstall\Teams` |
| La désinstallation par utilisateur documentée (`Update.exe --uninstall -s`) doit s'exécuter en tant que propriétaire du profil, ce que System ne peut pas faire ; les fichiers sont donc supprimés à la place. Les données itinérantes dans `%APPDATA%\Microsoft\Teams` ne sont pas touchées |
| La gestion des échecs distingue volontairement les deux cas : un installateur à l'échelle de la machine qui survit à la désinstallation est un véritable échec (code de sortie 1), tandis qu'un dossier par profil qui survit est presque toujours dû à un verrou de fichier d'un Teams classique en cours d'exécution — un avertissement, levé à l'exécution suivante une fois l'utilisateur déconnecté |
| Testé avec une détection simulée, puisque l'appareil de test n'a aucune des deux variantes : `-WhatIf` planifie la désinstallation msiexec et la suppression du dossier et ignore les étapes 5 à 8, et une exécution appliquée a supprimé un dossier de profil simulé, la vérification le déclarant propre. Le chemin msiexec lui-même n'a **pas** été exécuté sur un véritable Machine-Wide Installer — consigné dans la documentation plutôt que sous-entendu |

### 2026-09-17
| Modification |
|--------|
| Correction d'un bug mis au jour par une véritable exécution sur un hôte de session : `Get-AppxPackage -AllUsers` ne trouve rien lorsque Teams est seulement *provisionné* et qu'aucun utilisateur ne l'a encore, si bien que la vérification de version n'avait rien à comparer, déclarait l'hôte obsolète et réinstallait ~275 MB à chaque exécution planifiée. La version installée se rabat désormais sur la version du paquet provisionné. Vérifié sur ce scénario précis : signale `Provisioned MSTeams <version>`, compare, ne fait rien |
| [`Update-TeamsClient.ps1`](scripts/Device/Update-TeamsClient.ps1) vérifie désormais si **Outlook lui-même** voit le complément de réunion, et pas seulement que le MSI s'est installé. Il lit `HKEY_USERS\<sid>\...\Outlook\Addins\TeamsAddin.FastConnect` pour chaque utilisateur connecté, plus la clé à l'échelle de la machine : `LoadBehavior 3` = chargé, `2`/`0` = Outlook l'a désactivé (le vrai cas « le bouton a disparu »). Signalé dans le preflight et la vérification, jamais comme un échec — un profil sur lequel personne n'est connecté ne peut pas être lu |
| Le preflight est devenu un inventaire complet de tous les endroits où Teams peut résider : AppX par utilisateur, le paquet provisionné, le *Teams Machine-Wide Installer* classique, les installations classiques par profil, le complément dans les deux ruches, l'enregistrement Outlook. Teams classique est signalé, pas supprimé — il partage la date de fin de support d'octobre 2026, et un installateur à l'échelle de la machine résiduel continue de le redéployer dans les nouveaux profils |
| Deux bugs trouvés en l'exécutant plutôt qu'en le lisant. Énumérer les profils avec une liste blanche `S-1-5-21-*` ignore **tous** les utilisateurs d'un appareil joint à Entra, où les SID sont de la forme `S-1-12-1-*` — la vérification affirmait que personne n'avait le complément enregistré alors que `LoadBehavior=3` était bien là. Et `New-PSDrive` respecte `ShouldProcess`, donc sous `-WhatIf` le lecteur `HKEY_USERS` n'était jamais créé et la même vérification en lecture seule mentait ; les ruches sont désormais adressées via `Registry::HKEY_USERS`, sans lecteur à créer |

### 2026-09-11 (5)
| Modification |
|--------|
| Ajout de [`scripts/Exchange/Move-SharedCalendar.ps1`](scripts/Exchange/Move-SharedCalendar.ps1) — tout-en-un : `-Search balie` trouve le calendrier, montre qui l'utilise, le déplace dans une boîte aux lettres de ressource avec [`Convert-SharedCalendarToResource.ps1`](scripts/Exchange/Convert-SharedCalendarToResource.ps1) et liste qui doit basculer. Une seule App Registration temporaire avec les autorisations des deux scripts, créée une fois et supprimée à la fin, pour une seule connexion au lieu de deux. Plusieurs correspondances se choisissent dans une liste ou se restreignent avec `-Owner` ; une exécution non interactive les liste et s'arrête plutôt que de deviner. Les deux scripts sont appelés, pas copiés, de sorte qu'il existe une seule implémentation de chaque étape |
| Correction de [`Convert-SharedCalendarToResource.ps1`](scripts/Exchange/Convert-SharedCalendarToResource.ps1) : dans une session non interactive, la confirmation à saisir était ignorée et le calendrier d'origine supprimé sans `-Force`. Il est désormais laissé en place avec un avertissement, sauf si `-Force` est indiqué |
| Les deux scripts de calendrier : un `-ClientId` explicite prime désormais sur une session Graph app-only existante, de sorte que l'application d'un script appelant est réellement utilisée. [`Convert-SharedCalendarToResource.ps1`](scripts/Exchange/Convert-SharedCalendarToResource.ps1) gagne `-PassThru` (objet résultat pour les appelants) |
| Option `J` ajoutée au sous-menu Exchange (`C`) |

### 2026-09-11 (4)
| Modification |
|--------|
| [`Get-CalendarMappings.ps1`](scripts/Exchange/Get-CalendarMappings.ps1) — la première exécution réelle (un tenant où « Balie planning » s'est avéré être un calendrier secondaire dans une boîte aux lettres archivée) a confirmé les hypothèses sur Graph : les lectures app-only renvoient les calendriers que les utilisateurs ont ajoutés, et un calendrier secondaire partagé apparaît dans leur liste sous son propre nom. Elle a aussi révélé une indication erronée : une ligne `NotMapped` désignait « Reservering vergaderzaal LBM » comme « probablement celui-ci », alors qu'il s'agit d'un *second* calendrier partagé par le même propriétaire. L'indication ignore désormais les entrées portant le nom d'un autre calendrier du propriétaire et - pour un calendrier secondaire - les entrées portant le nom du propriétaire, qui correspondent à son calendrier principal |

### 2026-09-11 (3)
| Modification |
|--------|
| Ajout de [`scripts/Exchange/Convert-SharedCalendarToResource.ps1`](scripts/Exchange/Convert-SharedCalendarToResource.ps1) — déplace un calendrier partagé (le calendrier « Balie » dans la boîte aux lettres d'une personne) vers une boîte aux lettres Room ou Equipment dédiée, avec chaque élément et chaque autorisation, puis supprime l'original sur demande. Aperçu par défaut ; `-Apply` crée et copie, `-RemoveSourceCalendar` ne supprime l'original qu'une fois que chaque élément a une copie vérifiée et que le nom du calendrier a été saisi en guise de confirmation |
| Les éléments sont copiés fidèlement plutôt qu'approximativement : les séries périodiques restent des séries, avec leurs occurrences déplacées et annulées appliquées (mises en correspondance occurrence par occurrence, et laissées telles quelles avec un avertissement si les deux séries ne concordent pas), les heures sont réécrites dans le fuseau horaire dans lequel elles ont été créées afin que les éléments hebdomadaires survivent à un changement d'heure, les catégories conservent leur couleur, les pièces jointes jusqu'à 3 MB sont copiées et les plus volumineuses enregistrées dans le dossier de sauvegarde. Les participants sont listés dans le corps au lieu d'être copiés, pour que personne ne reçoive une nouvelle invitation |
| Les autorisations conservent leurs droits d'accès Exchange exacts, droits personnalisés compris ; `-SendSharingInvitation` envoie aux utilisateurs l'invitation standard. Les personnes externes, les comptes supprimés et les indicateurs de délégué sont signalés, pas abandonnés en silence |
| Chaque copie porte l'identifiant de son élément source dans une propriété masquée, de sorte qu'une exécution interrompue à mi-chemin reprend là où elle s'était arrêtée ; une série à moitié terminée est refaite. Une sauvegarde JSON de tout ce qui a été lu est écrite avant toute création |
| Option `I` ajoutée au sous-menu Exchange (`C`) : toujours un aperçu d'abord, puis une seconde étape explicite |

### 2026-09-11 (2)
| Modification |
|--------|
| [`Get-CalendarMappings.ps1`](scripts/Exchange/Get-CalendarMappings.ps1) gagne `-Search` (alias `-Keyword`) : « où se trouve le calendrier Balie ? » en une seule exécution. Le mot-clé est comparé au nom du propriétaire et à chaque adresse (une boîte aux lettres partagée `balie@`, une salle, un groupe) ainsi qu'aux noms de calendrier (un calendrier secondaire *Balie* dans la boîte aux lettres de quelqu'un). Le rapport montre où réside le calendrier (nouvel état `Source`), qui l'a dans sa liste de calendriers et qui dispose de droits dessus |
| Un calendrier **secondaire** correspondant voit désormais ses propres autorisations lues, au lieu d'être comparé au calendrier principal du propriétaire et de finir en `MappedWithoutRight`. Une nouvelle colonne `Calendar` indique de quel calendrier du propriétaire traite une ligne |
| Une entrée de liste de calendriers ne contient aucun lien vers le dossier dont elle provient, donc un calendrier secondaire partagé est mis en correspondance par son nom. Lorsqu'un utilisateur l'a sous un autre nom, la ligne `NotMapped` désigne l'entrée qui correspond probablement plutôt que de laisser un faux négatif silencieux |
| L'option de menu `H` demande d'abord un mot-clé ; laissée vide, elle revient au rapport complet ou par boîte aux lettres |

### 2026-09-11
| Modification |
|--------|
| Ajout de [`scripts/Exchange/Get-CalendarMappings.ps1`](scripts/Exchange/Get-CalendarMappings.ps1) — montre où chaque calendrier est réellement mappé : pour chaque boîte aux lettres, il lit la liste de calendriers dans Outlook et les droits sur son propre calendrier principal, et regroupe les deux en une ligne par propriétaire + utilisateur avec un état (`Mapped`, `MappedWithoutRight`, `NotMapped`, `MappedOwnerMissing`, `SharedExternally`, …). [`Test-CalendarPermissions.ps1`](scripts/Exchange/Test-CalendarPermissions.ps1) dit qui *peut* ouvrir un calendrier ; celui-ci dit où il *est*, et où les deux divergent |
| Graph plutôt qu'Exchange Online PowerShell, car les entrées qu'un utilisateur a ajoutées à sa propre liste de calendriers ne sont visibles par aucune cmdlet Exchange. L'accès app-only suit les trois mêmes voies que [`Remove-PhishingMessage.ps1`](scripts/Exchange/Remove-PhishingMessage.ps1) (session existante, application propre, ou application temporaire supprimée dans un `finally`), avec les autorisations en lecture seule `Calendars.Read`, `User.Read.All` et `Group.Read.All`. Aucune connexion Exchange, donc aucun conflit MSAL |
| Les exécutions à l'échelle du tenant passent par `$batch` (20 boîtes aux lettres par appel), avec nouvelle tentative pour les éléments limités (throttling). Une boîte aux lettres illisible est signalée comme telle plutôt que comme « rien de mappé » |
| Consigné ce que le rapport ne peut pas voir : Full Access avec AutoMapping (une autorisation de boîte aux lettres — [`Test-MailboxPermissions.ps1`](scripts/Exchange/Test-MailboxPermissions.ps1)), les calendriers ouverts dans Outlook classique sans les améliorations des calendriers partagés, et les calendriers secondaires, qui apparaissent comme `MappedWithoutRight` |
| Option `H` ajoutée au sous-menu Exchange (`C`) |

### 2026-09-16 (5)
| Modification |
|--------|
| La page d'explication est désormais l'étape 4 de la construction en une seule commande, et non plus un script séparé dont quelqu'un se souvient une semaine plus tard. Une structure dont personne n'a été informé est une structure que personne n'utilise, et comme la page est générée à partir de la même configuration, elle décrit exactement ce que l'exécution vient de créer |
| `-SkipHelpPage` l'omet, `-HelpContact` indique à qui les utilisateurs doivent s'adresser. L'installateur passe `-Force`, car cette page lui appartient : relancer la construction réaligne l'explication sur ce que la construction a créé |
| La vérification et l'audit de la deelstatus passent aux étapes 5 et 6, et les avertissements obsolètes « l'étape 2 modifie les autorisations sur une équipe en production » citent désormais l'étape 3 |
### 2026-09-16 (4)
| Modification |
|--------|
| Ajout de [`scripts/SharePoint/Provisioning/Add-SharePointHelpPage.ps1`](scripts/SharePoint/Provisioning/Add-SharePointHelpPage.ps1) — place l'explication destinée aux utilisateurs finaux sur le site d'équipe sous forme de page SharePoint, avec un lien dans la navigation de gauche. Une handleiding dans un dépôt n'est lue par personne ; ce script l'écrit là où se trouvent déjà les personnes qui déposent des fichiers |
| La page est générée à partir de la configuration plutôt que rédigée à la main, elle ne peut donc pas s'écarter de ce que font réellement les bibliothèques : les canaux listés sont ceux qui existent, les étiquettes portent le même texte d'aide que celui affiché sous chaque champ du formulaire de dépôt, et les champs obligatoires par type de document sont lus depuis les types de contenu |
| Rédigée pour la personne qui dépose un catalogue. Deux éléments de la configuration en sont volontairement exclus : la `note` d'un conteneur, qui nomme des groupes de sécurité, et la `description` d'une vue, qui parle de piliers et de dossiers synchronisés. Les autorisations sont entièrement omises — qui peut voir quoi n'est pas quelque chose sur lequel un utilisateur peut agir |
| Corrigé au passage, découvert en affichant la page plutôt qu'en lisant le code : elle annonçait trois façons d'ajouter un fichier et n'en listait que deux, et l'assistant écrivait le nom de l'équipe là où devait figurer le nom de l'entreprise (« Intern blijft binnen Laseto-NewTeams ») |
### 2026-09-16 (3)
| Modification |
|--------|
| Ajout de [`scripts/SharePoint/Provisioning/Remove-SharePointStructure.ps1`](scripts/SharePoint/Provisioning/Remove-SharePointStructure.ps1) — démonte la même configuration, en commençant par le plus profond : onglets, canaux, bibliothèques, types de contenu (d'abord détachés de leurs listes), colonnes de site, ensemble de termes, groupes de sécurité, et l'équipe elle-même |
| **Volontairement l'inverse du comportement par défaut de tout le reste du dossier : sans `-Apply`, rien n'est modifié.** Oublier `-WhatIf` sur un script destructeur est la direction dangereuse, l'état sûr est donc celui que l'on obtient sans rien faire |
| `-Scope All` n'inclut jamais l'équipe. Supprimer toute l'équipe d'un client n'est pas quelque chose que l'on doit obtenir en demandant « tout » — elle doit être nommée explicitement, puis son nom saisi pour confirmer |
| Refuse par défaut plutôt que de demander pardon : une bibliothèque ou un dossier de canal qui contient encore des fichiers est ignoré sauf avec `-IncludeContent` (le nombre d'éléments est signalé dans tous les cas), le canal Général et la bibliothèque Documents propre à l'équipe ne sont jamais supprimés, et un type de contenu encore utilisé est signalé plutôt que forcé |
| Signalés avec leur coût avant d'être exécutés, car aucune corbeille ne les restaure : supprimer l'ensemble de termes rend orpheline la valeur Leverancier sur chaque document qui en portait une, et supprimer une colonne emporte ses données avec elle. Ce qui *est* récupérable est également indiqué — un groupe ou une équipe supprimé(e) est conservé(e) en suppression réversible pendant 30 jours, un canal dispose de sa propre corbeille de 30 jours, et les fichiers d'une bibliothèque supprimée atterrissent dans la corbeille du site |
| L'étape `6` du menu l'exécute ; le menu pose trois questions distinctes — sur l'application, sur les fichiers et sur l'équipe — plutôt qu'une seule |
### 2026-09-16 (2)
| Modification |
|--------|
| Un pilier restreint peut désormais prendre la forme d'**une bibliothèque propre derrière un canal ordinaire**, la seule configuration qui offre un vrai rôle en lecture seule tout en plaçant un canal dans Teams. L'assistant demande quels piliers sont restreints, puis sous quelle forme - `bibliotheek` (par défaut) ou `privekanaal` |
| La forme bibliothèque rompt l'héritage **sans le copier**, et c'est tout l'intérêt : la copie reporterait chaque membre de l'équipe en tant qu'éditeur, ce qui est exactement la porte que cette forme est censée fermer. Ce qui subsiste, ce sont les propriétaires du site lui-même plus les deux groupes du pilier - Contribute et Read |
| Un canal privé n'offre aucun rôle en lecture seule : propriétaires et membres, et les membres peuvent publier, modifier et supprimer. Un pilier qui a besoin de « peut consulter » ne peut donc pas être un canal privé, et l'assistant le dit désormais au moment où le choix est fait |
| Le canal est créé comme d'habitude mais ne reçoit pas de dossier dans la bibliothèque partagée, et la bibliothèque restreinte apparaît comme onglet dans ce canal - l'onglet Fichiers propre à un canal standard pointe toujours vers la bibliothèque de l'équipe et ne peut pas être redirigé, elle se place donc à côté |
| Signalé dans le readme car cela serait sinon rapporté comme un bug : cet onglet Fichiers intégré reste, pointant vers un dossier que personne n'utilise. Soit vous orientez les utilisateurs vers l'onglet nommé, soit vous supprimez une fois pour toutes l'onglet Fichiers du canal à la main |
### 2026-09-16
| Modification |
|--------|
| Ajout de [`scripts/SharePoint/Provisioning/Sync-SharePointChannelMember.ps1`](scripts/SharePoint/Provisioning/Sync-SharePointChannelMember.ps1) — fait d'un groupe de sécurité Entra ID la source de vérité pour les membres d'un canal Teams privé. Un canal privé ne peut pas du tout recevoir de droits via un groupe : Teams gère sa liste de membres personne par personne et Graph n'y accepte que des utilisateurs individuels, le groupe alimente donc la liste de membres à la place |
| Le contournement évident est un piège et est documenté comme tel : ajouter le groupe aux autorisations SharePoint du site du canal privé fonctionne jusqu'à ce que Teams resynchronise la liste de membres par-dessus, et entre-temps ces personnes accèdent aux fichiers alors que le canal leur reste invisible dans Teams. Non pris en charge par Microsoft |
| Les groupes imbriqués sont suivis, les objets qui ne sont pas des utilisateurs sont écartés, et chacun est d'abord ajouté à l'équipe parente — Teams refuse un membre de canal privé qui ne fait pas partie de l'équipe, et l'erreur renvoyée ne le mentionne pas. `-Prune` retire aussi les personnes que les groupes ne listent plus ; les propriétaires du canal ne sont jamais retirés |
| Consigné parce que cela change la conception, pas seulement le script : **un canal privé n'a pas de rôle en lecture seule.** Propriétaires et membres, et les membres peuvent publier, modifier et supprimer. Un groupe nommé `-RO` ne peut pas y signifier « peut consulter », l'exécution indique donc par groupe combien de personnes il a ajoutées plutôt que de laisser passer cela inaperçu. Là où la lecture seule compte vraiment, une bibliothèque de documents avec ses propres autorisations est la bonne forme |
| `-EnsureGroups` crée désormais chaque groupe du modèle, et plus seulement ceux auxquels une bibliothèque accorde des droits. Un canal privé n'accorde rien, la paire MGMT figurait donc dans la configuration sans jamais être créée — et c'est précisément le pilier dont on cherche les groupes en premier |
| L'assistant écrit une section `channelMembers` pour les piliers privés, nommant les groupes qui alimentent la liste de membres ; une configuration écrite avant l'existence de cette clé se rabat sur les groupes configurés portant le nom du conteneur, et signale ce repli |
| L'étape `5` du menu exécute la synchronisation ; elle se connecte elle-même à Graph et ne demande donc pas d'inscription d'application PnP |
### 2026-09-15 (3)
| Modification |
|--------|
| `New-StructureConfig.ps1 -All` demande les noms qui étaient encore dérivés à l'insu de l'opérateur : par pilier le nom du canal, le dossier, le type de contenu, les deux noms de groupe et le titre de la vue ; plus la bibliothèque derrière les canaux, les groupes de colonnes et de types de contenu, l'ensemble de termes, l'URL du site d'équipe et l'étiquette que chaque colonne affiche pour l'utilisateur. Chacun garde sa dérivation comme suggestion, donc `-All` se résume encore surtout à des Entrée |
| Seuls les noms *internes* des colonnes restent fixes. Ils ne sont jamais montrés à personne, et en changer un alors que des documents le portent fait perdre les métadonnées de ces documents |
| Corrigé : une question facultative ne pouvait jamais être refusée, car Entrée signifie « accepter la suggestion ». Les questions facultatives affichent désormais `(of "geen")` et acceptent geen/none/nee/- comme un vrai « aucun » — auparavant, ne rien répondre à « customer library » créait quand même FUTECH |
| Corrigé une liste à un élément revenant sous forme de simple chaîne : PowerShell déroule un tableau à un seul élément au retour, si bien qu'un client avec une seule marque faisait planter l'assistant sur `.Count`. Désormais renvoyé avec une virgule en tête |
| Corrigé deux affectations `$x = if (...) { @() }` qui donnent `$null` au lieu d'un tableau vide — une configuration sans pilier ventes ou sans fournisseurs échouait au récapitulatif |
| Les trois parcours vérifiés de bout en bout avec le validateur de configuration : la configuration par défaut complète à six piliers, une exécution `-All` avec des noms volontairement différents partout, et un tenant minimal à deux piliers sans canal privé, sans fournisseurs, sans régions et sans bibliothèque clients |
### 2026-09-15 (2)
| Modification |
|--------|
| [`New-StructureConfig.ps1`](scripts/SharePoint/Provisioning/New-StructureConfig.ps1) demande comment tout doit s'appeler et écrit lui-même la configuration — personne ne devrait avoir à ouvrir un fichier JSON pour nommer un canal. Entrée accepte la suggestion entre crochets, donc une construction standard se résume surtout à des Entrée plus le tenant et le propriétaire de l'équipe |
| Tout le reste est dérivé de ces réponses : par pilier un canal, un type de contenu, deux groupes de sécurité et une vue groupée ; par marque une vue transversale couvrant tous les dossiers de piliers. Le pilier qui gère les fournisseurs et celui qui gère les ventes déterminent où Leverancier et Regio deviennent des champs obligatoires |
| Deux des réponses sont celles qui coûtent quelque chose plus tard, elles sont donc posées en dernier avec non par défaut : la maintenance de la colonne de statut de partage (le seul script nocturne) et l'application de droits par pilier sur les dossiers de canaux standard (la partie que Microsoft ne prend pas en charge) |
| [`New-SharePointTeam.ps1`](scripts/SharePoint/Provisioning/New-SharePointTeam.ps1) crée l'équipe Microsoft 365 et ses canaux, y compris le canal privé MGMT, afin que la structure puisse être construite à partir d'un tenant vide. La collection de sites d'un canal privé est provisionnée de façon asynchrone et son URL ne peut pas être connue à l'avance — le script l'interroge à intervalles réguliers et la réécrit dans la configuration, ce qui permet aux étapes suivantes de se connecter à quelque chose |
| Ne renomme ni ne supprime jamais un canal : un canal dont le nom ne correspond pas à la configuration est signalé, pas corrigé, car renommer un canal déplace son dossier et casse chaque lien que quelqu'un a partagé |
| [`Install-SharePointStructure.ps1`](scripts/SharePoint/Provisioning/Install-SharePointStructure.ps1) lance lui-même l'assistant lorsqu'il ne trouve aucune configuration pour le tenant, et l'étape équipe est désormais l'étape 1 sur six. `-SkipTeam` pour une équipe qui existe déjà |
| Les noms internes des colonnes et les ID de types de contenu sont générés une seule fois puis figés — SharePoint indexe les métadonnées des documents sur les deux — c'est pourquoi l'assistant refuse d'écraser une configuration existante sans `-Force`. Les noms d'affichage, noms de canaux et noms de groupes restent modifiables |
| Suppression des derniers noms de colonnes codés en dur : l'audit du statut de partage lit quelle colonne est laquelle dans une nouvelle section `fieldRoles` au lieu de supposer `PsDeelstatus` et `PsVertrouwelijkheid` |
| L'inscription d'application consent désormais aussi `Channel.Create`, `ChannelSettings.ReadWrite.All` et `Team.Create`, dont l'étape équipe a besoin |
### 2026-09-15
| Modification |
|--------|
| Ajout de [`scripts/SharePoint/Provisioning/Install-SharePointStructure.ps1`](scripts/SharePoint/Provisioning/Install-SharePointStructure.ps1) — construit toute la structure en une seule exécution : inscrit lui-même l'application Entra et donne le consentement administrateur à ses étendues déléguées, puis métadonnées → bibliothèques/groupes/autorisations → vérification, et éventuellement le premier audit de la deelstatus. Chaque étape reste son propre script, un échec se relance donc isolément au lieu de tout recommencer |
| `-TemporaryApp` supprime à nouveau l'inscription d'application à la fin, pour une construction ponctuelle sur un tenant que vous ne gérez pas au quotidien. Il ne supprime jamais qu'une application **créée par cette exécution** — une application déjà en cache est antérieure à l'exécution et c'est à quelqu'un d'autre de la supprimer, le script le signale donc plutôt que de la supprimer discrètement. Sans le switch, l'application reste et l'ID client est mis en cache dans `pnp.appid.json`, partagé avec les autres scripts PnP de ce dépôt |
| Signalé dans la documentation car cela serait sinon rapporté comme un bug : une exécution `-WhatIf` a besoin d'une application pour se connecter. Sans application en cache pour le tenant, il n'y a rien avec quoi se connecter, l'essai à blanc valide donc la configuration et s'arrête là — exécutez-le une fois pour de vrai, ou passez `-ClientId`, pour faire un essai à blanc étape par étape |
| Comblé le manque qui rendait « merk als tag » à moitié vrai seulement : des vues transversales sur la bibliothèque partagée avec `Scope = RecursiveAll`, si bien que *Alles - Butterstone* est une seule liste à plat couvrant tous les dossiers de piliers, **y compris tout ce qui est étiqueté Beide** — un seul fichier, deux marques, pas de copies qui divergent. Plus *Nog te taggen* (ce que le glisser-déposer et la synchronisation OneDrive laissent derrière eux), *Extern gedeeld* et *Te archiveren*. Le filtre est du CAML brut dans la configuration plutôt qu'un mini-langage de requête inventé par le script |
| Ne jamais grouper une vue sur `PsTaal` : SharePoint refuse de grouper sur une colonne à valeurs multiples. Le filtrage dessus fonctionne très bien, et aucune vue livrée ne groupe dessus |
| Ajout de `Petsolutions-SharePoint-Handleiding.md` — documentation utilisateur final en néerlandais à remettre au client. Elle couvre les trois façons d'ajouter un fichier et pourquoi elles se comportent différemment, ce que signifie chaque étiquette, et ce qui se passe au moment où l'on étiquette quelque chose (le fichier ne bouge pas, les liens continuent de fonctionner, `Beide` apparaît dans les vues des deux marques, la recherche a quelques minutes de retard sur les vues) |
| Deux points que le guide énonce clairement parce que les utilisateurs supposent le contraire : **le glisser-déposer et la synchronisation OneDrive ne demandent rien** — les colonnes obligatoires sont imposées par le formulaire de dépôt, pas par la bibliothèque, donc les fichiers déposés en masse arrivent avec des étiquettes vides et une invite « Informations requises » au lieu d'être bloqués ; et **une étiquette n'est pas un verrou** — Vertrouwelijkheid n'exclut personne, c'est un accord plus le signal que l'audit nocturne utilise pour repérer un partage excessif |
| L'élément de menu `S` a reçu l'étape `0` pour la construction tout-en-un ; les options par étape sont inchangées |

### 2026-09-10 (4)
| Modification |
|--------|
| Ajout de [`scripts/SharePoint/Provisioning/`](scripts/SharePoint/Provisioning/readme.fr.md) — provisionner et maintenir toute une structure SharePoint (modèle de métadonnées, types de contenu, bibliothèques, autorisations par groupes Entra ID) pour un client MSP à partir d'une seule configuration JSON, avec un audit de partage et un contrôle de dérive en lecture seule. Conçu pour Petsolutions NV (marques Butterstone/Laseto), mais rien dans les scripts n'est propre à ce client |
| Le modèle se trouve dans `petsolutions.config.json`, contrôlé de manière croisée au chargement : un type de contenu faisant référence à une colonne non définie, ou un conteneur accordant des droits à un groupe absent du modèle, échoue avant toute connexion plutôt qu'au milieu du provisionnement. Les URL de tenant/site `CHANGEME` livrées sont refusées d'emblée |
| Les noms internes des colonnes portent un préfixe `Ps`. « Contenttype » et « Status » sont des noms d'affichage que SharePoint utilise déjà pour autre chose, et le préfixe les rend non ambigus dans le CAML, dans les vues et dans le contrôle de dérive, tandis que les utilisateurs voient toujours de simples étiquettes en néerlandais. Les ID de types de contenu sont fixes plutôt que générés, la même structure est donc reproductible d'un tenant à l'autre |
| [`New-SharePointMetadata.ps1`](scripts/SharePoint/Provisioning/New-SharePointMetadata.ps1) s'exécute sur **chaque** site de la configuration, pas seulement sur le site d'équipe : un canal privé Teams (ici MGMT) est sa propre collection de sites et une colonne de site ne la traverse pas. Rendre une colonne obligatoire après coup fonctionne — l'indicateur `Required` d'un lien de champ existant est mis à jour sur place et propagé aux listes qui utilisent déjà le type de contenu |
| [`Set-SharePointLibraries.ps1`](scripts/SharePoint/Provisioning/Set-SharePointLibraries.ps1) définit un ordre de types de contenu par dossier, si bien que le menu *Nouveau* dans le canal Leveranciers propose Leveranciersdocument et non les cinq types appartenant aux autres piliers — la bibliothèque partagée doit tous les porter, le dossier n'a pas à tous les afficher. Aucune vue n'est jamais définie par défaut : la vue par défaut d'une bibliothèque Teams est ce que chaque membre du canal voit à l'instant où il ouvre Fichiers |
| Consigné plutôt que caché : des autorisations uniques sur un dossier de canal **standard** sont ce que ce modèle demande et ce que Microsoft ne prend pas en charge. Les membres qui perdent l'accès continuent de voir le canal dans Teams et obtiennent une erreur sur l'onglet Fichiers au lieu d'une porte fermée. Le script le fait, avertit par dossier, et `-SkipChannelFolderPermissions` laisse ces dossiers hériter. Un canal privé, un canal partagé ou une bibliothèque propre (ce qu'utilise FUTECH) sont les moyens pris en charge pour isoler un pilier |
| [`Update-SharePointShareStatus.ps1`](scripts/SharePoint/Provisioning/Update-SharePointShareStatus.ps1) dérive la colonne Deelstatus des autorisations réellement présentes sur chaque fichier. Il pose d'abord la question la moins coûteuse — un fichier qui hérite n'est pas partagé — si bien qu'un aller-retour par centaine d'éléments règle presque toute la bibliothèque ; seuls les fichiers ayant rompu l'héritage voient leurs attributions de rôles lues, et parmi eux seuls les liens pour des personnes spécifiques doivent être développés (un lien Anyone ou Organization indique déjà dans son nom si un invité peut se trouver derrière). Écrit avec `SystemUpdate` afin que Modified/Modified By restent inchangés et qu'aucune version ne soit créée |
| L'audit ne révoque jamais un lien. Il signale les fichiers étiquetés Intern ou Vertrouwelijk qui se trouvent derrière un lien externe et se termine avec `2`, afin qu'une tâche RMM planifiée remonte exactement quand une personne a une décision à prendre. [`Test-SharePointStructure.ps1`](scripts/SharePoint/Provisioning/Test-SharePointStructure.ps1) fait de même pour la dérive structurelle, classée en Missing / Different / Extra — « Extra » n'est jamais corrigé automatiquement, car une colonne en trop contient des données et une attribution de rôle en trop est généralement l'exception délibérée de quelqu'un |
| [`SharePointStructure.Common.ps1`](scripts/SharePoint/Provisioning/SharePointStructure.Common.ps1) est chargé par dot-sourcing dans les quatre — une exception délibérée à la règle « chaque script est autonome » appliquée ailleurs dans ce dépôt, car ils partagent un même schéma de configuration et trois copies du code d'autorisations divergeraient en moins d'un mois |
| Ajout de l'élément de menu `S` pour l'ensemble (choisissez une étape, `-WhatIf` sauf si vous confirmez ; le contrôle de dérive saute la question car il n'écrit jamais) |
### 2026-09-10 (4)
| Modification |
|--------|
| Rattachement de la date d'expiration à la fonctionnalité `-AvdOptimizations` : Microsoft retire l'optimisation multimédia AVD basée sur WebRTC le **1er octobre 2026** (fin du support) et le **1er avril 2027** (fin de disponibilité), et Teams affiche déjà une bannière à ce sujet aux utilisateurs. Le switch continue d'installer le redirecteur car Microsoft le recommande toujours comme solution de repli — avec une note pour y revenir avant avril 2027 |
| Son remplaçant, SlimCore, ne nécessite rien sur l'hôte de session : il est intégré au nouveau Teams. Confirmé sur un appareil avec Teams `26225.1806.5074.1452`, qui embarque `Microsoft.Teams.SlimCoreVdiHost.win-x64` `2026.31.1.16` ainsi que plusieurs packages de framework. Le contrôle préalable signale désormais ce package sous `-AvdOptimizations` |
| Ce signalement est volontairement informatif et ne crée aucune tâche : le chemin multimédia utilisé dépend de la version de Windows App sur le poste depuis lequel l'utilisateur se connecte, ce qu'un script s'exécutant sur l'hôte de session ne peut pas voir. L'audit des versions client des postes est le véritable travail de migration |
| Ajout d'une section service desk à la doc IT Glue pour la bannière que les utilisateurs signalent : ce qu'elle signifie (une annonce, pas une panne), les deux dates, le fait que la correction se fait sur l'appareil local et non sur l'hôte de session, comment lire la ligne `AVD SlimCore Media Optimized` / `AVD Media Optimized` sous Teams > À propos, et un texte prêt à l'emploi pour l'utilisateur |

### 2026-09-10 (3)
| Modification |
|--------|
| Correction d'une affirmation erronée dans la documentation Teams et dans le commentaire du script : l'entrée de désinstallation du complément de réunion ne se trouve **pas** toujours dans `WOW6432Node`. Mesuré sur un poste Windows 11, le complément `1.26.21803` s'enregistre dans la ruche **64 bits**, avec `InstallSource` pointant vers un cache MSI par utilisateur. Parcourir les deux ruches (ce que le script faisait déjà) est correct — la raison invoquée ne l'était pas |
| Documenté comment le complément arrive réellement sur un appareil, mesuré plutôt que supposé : le script l'installe pour toute la machine (`ALLUSERS=1`, `Program Files (x86)`) pour les machines partagées et les hôtes de session, tandis que sur un poste ordinaire le client Teams l'installe et le met à jour **par utilisateur** depuis `%LOCALAPPDATA%\Microsoft\TeamsMeetingAddinMsis` vers `%LOCALAPPDATA%\Microsoft\TeamsMeetingAdd-in`, en s'enregistrant uniquement dans `HKCU\...\Office\Outlook\Addins` |
| Consigné en même temps : ce que l'étape 8 prouve réellement. Elle lit les clés de désinstallation HKLM, elle confirme donc que l'installation pour toute la machine a réussi — pas que l'Outlook d'un utilisateur donné affiche le bouton. Exécuté en tant que System, le script ne peut pas du tout voir le `HKCU` d'un utilisateur |

### 2026-09-10 (2)
| Modification |
|--------|
| [`scripts/Device/Update-TeamsClient.ps1`](scripts/Device/Update-TeamsClient.ps1) intègre les parties AVD/VDI de l'ancien installateur complémentaire derrière `-AvdOptimizations` : l'indicateur multimédia `IsWVDEnvironment` (défini à l'étape 3, avant le provisionnement du client, car Teams le lit au démarrage pour choisir son chemin multimédia) et le Remote Desktop WebRTC Redirector Service depuis `aka.ms/msrdcwebrtcsvc/msi` |
| Volontairement un switch et non une détection automatique : définir cet indicateur sur un poste normal indique à Teams de confier le multimédia à un redirecteur qui n'existe pas. Sans le switch, le script se contente de *signaler* qu'un appareil ressemble à un hôte de session (`HKLM:\SOFTWARE\Microsoft\RDInfraAgent`) |
| Les deux composants ne sont installés que s'ils sont absents (`-Force` réinstalle le redirecteur), si bien qu'une exécution planifiée sur un hôte de session configuré ne télécharge toujours rien et n'affiche rien sous `-Quiet`. Vérifié de bout en bout, y compris que le MSI du redirecteur (1.7 MB, `1.54.2408.19001`) passe la vérification de signature Microsoft |
| L'étape du complément se saute désormais aussi elle-même lorsque le complément est présent et que le client n'a pas été remplacé — auparavant, il réinstallait le complément lors d'une exécution qui n'avait pour but que de corriger les composants AVD |
| Le téléchargement et la vérification de signature ont été regroupés dans une seule fonction d'aide `Save-VerifiedDownload`, partagée par le bootstrapper et le redirecteur : https uniquement, taille minimale, Authenticode `Valid` et signé par `O=Microsoft Corporation`, sinon une exception est levée |
| Corrigé dans la documentation : les noms de variables de script Ninja ne sont **pas** sensibles à la casse. Les recherches de variables d'environnement Windows ne tiennent pas compte de la casse, donc des variables nommées `Quiet` ou `Force` fonctionnent exactement comme `quiet` et `force` |

### 2026-09-10
| Modification |
|--------|
| [`scripts/Device/Update-TeamsClient-ITGlue.md`](scripts/Device/Update-TeamsClient-ITGlue.md) a reçu une annexe de configuration NinjaOne : quelles valeurs choisir pour chaque champ lors de l'ajout du script (PowerShell 5.1 plutôt que 7, 64 bits, Run As System), les noms des variables de script avec un moyen de vérifier qu'elles arrivent bien, l'exécution de test sur un seul appareil, l'automatisation planifiée avec `-Quiet -Confirm:$false`, et la tâche de détection facultative |
| Consigné explicitement car cela serait sinon rapporté comme un bug : `-CheckOnly` se termine avec `2` lorsqu'une mise à jour est disponible, et NinjaOne affiche tout code de sortie non nul comme une tâche en échec. C'est voulu — ce sont les appareils qui demandent de l'attention — et c'est sur cela qu'une condition de résultat de script peut s'appuyer |
| Également signalé : le délai d'expiration du script Ninja doit dépasser `-TimeoutSeconds` (900 s) plus le téléchargement d'environ 275 MB, sinon Ninja interrompt la tâche en pleine installation |

### 2026-09-08 (5)
| Modification |
|--------|
| Ajout de [`scripts/Device/Update-TeamsClient-ITGlue.md`](scripts/Device/Update-TeamsClient-ITGlue.md) — la version service desk de cette documentation, en néerlandais, à coller dans IT Glue. Organisée par niveau de support : le N1 vérifie avec `-CheckOnly -Quiet` et lit la sortie étiquetée, le N2 lance la mise à jour depuis NinjaOne ou à la main et vérifie ensuite, le N3 dispose des paramètres, codes de sortie, chemins et garde-fous intégrés. Comprend un tableau d'erreurs avec le niveau d'escalade par message, une FAQ et un texte prêt à l'emploi pour l'utilisateur final |
| Elle commence par le point qui piège les gens : aucune sortie signifie que l'appareil est déjà à jour, ce qui est une exécution réussie et non un échec. Le volume de téléchargement par appareil (~275 MB : un bootstrapper de 1.9 MB qui récupère un package d'environ 273 MB) a été mesuré, pas estimé |

### 2026-09-08 (4)
| Modification |
|--------|
| Ajout de [`scripts/Device/Update-TeamsClient.md`](scripts/Device/Update-TeamsClient.md) — une référence pour ce script : l'arbre de décision (en retard → réinstallation complète, à jour mais complément manquant → complément uniquement, à jour → rien du tout), les sept étapes, la vérification de version via le service de configuration avec un exemple de réponse, les modes de sortie et codes de sortie, le tableau des variables de script NinjaOne, les choix de conception derrière l'ordre des opérations, un tableau de dépannage et ce qui a réellement été testé. Lié depuis le readme Device |

### 2026-09-08 (3)
| Modification |
|--------|
| [`scripts/Device/Update-TeamsClient.ps1`](scripts/Device/Update-TeamsClient.ps1) ne réinstalle plus inconditionnellement : il interroge le service de configuration du client Teams (`config.teams.microsoft.com/config/v1/MicrosoftTeams/...`, `BuildSettings.WebView2PreAuth.<arch>.latestVersion` — le même flux que le client utilise pour décider qu'il est obsolète) pour savoir quel build est publié pour cette architecture, et laisse un appareil à jour totalement tranquille |
| Le client est à jour mais le complément de réunion manque ? Alors seul le complément est installé — pas de téléchargement, pas de désinstallation, pas de reprovisionnement |
| `-Quiet` retient toute sortie jusqu'à ce qu'il y ait du nouveau, si bien qu'une exécution NinjaOne planifiée n'affiche rien sur un appareil à jour et n'apparaît dans le flux d'activité que lorsqu'elle a trouvé un build plus récent ou rencontré un problème. Vérifié : une exécution `-Quiet` sur un appareil à jour produit zéro octet de sortie et le code de sortie 0 |
| `-CheckOnly` fait un rapport sans rien modifier et se termine avec `2` lorsqu'un build plus récent est disponible, pour une utilisation comme tâche de détection/condition Ninja. `-Ring` sélectionne un anneau de mise à jour autre que celui par défaut |
| Lorsque le service de configuration est injoignable, l'exécution s'arrête au lieu de réinstaller à l'aveugle ; `-Force` signifie désormais « réinstaller même s'il est à jour » ainsi que « continuer sans Teams ni information de version » |
| Une transcription n'est écrite que lorsque l'exécution modifie réellement quelque chose, si bien qu'une vérification horaire ne laisse pas de journaux inutiles dans `C:\Temp` |

### 2026-09-08 (2)
| Modification |
|--------|
| [`scripts/Device/Update-TeamsClient.ps1`](scripts/Device/Update-TeamsClient.ps1) peut désormais être exécuté sans surveillance depuis un RMM (NinjaOne) *et* à la main. Il se relance lui-même en 64 bits via `SysNative` lorsque l'agent démarre PowerShell en 32 bits — sinon les lectures HKLM sont redirigées vers `WOW6432Node` et `$env:ProgramFiles` pointe vers le dossier x86, si bien que ni le package AppX ni le MSI du complément ne sont jamais trouvés |
| Les variables de script NinjaOne (`whatIf`, `force`, `skipMeetingAddIn`, `skipSignatureCheck`, `workingDir`, `logPath`) sont lues depuis l'environnement lorsque le paramètre correspondant n'est pas passé, si bien qu'une exécution d'aperçu peut se faire par une case à cocher au lieu d'une chaîne de paramètres |
| Lancé à la main sans élévation, il demande désormais l'UAC et continue dans une fenêtre élevée, au lieu d'échouer sur une ligne `#Requires -RunAsAdministrator`, et une exécution interactive en mode application demande une seule confirmation. `-Confirm:$false` le rend autonome ; le menu le passe car il a déjà posé la question |
| Réordonné de sorte que le bootstrapper soit téléchargé **et** que sa signature Authenticode Microsoft soit vérifiée avant la première désinstallation — un téléchargement échoué ou une URL bloquée ne peut plus laisser un appareil sans client Teams. TLS 1.2 est imposé pour le téléchargement, et un `-BootstrapperUrl` non https est refusé |
| `msiexec` et le bootstrapper passent désormais par une seule fonction d'aide avec un délai d'expiration (`-TimeoutSeconds`, 900 par défaut, processus tué à expiration), une nouvelle tentative sur 1618 (une autre installation en cours) et 3010 traité comme un succès avec une note de redémarrage en attente, si bien qu'une tâche RMM ne peut jamais bloquer l'agent |
| Le package AppX est également déprovisionné (`Remove-AppxProvisionedPackage`), sinon les nouveaux profils utilisateur continuent de recevoir l'ancienne version préparée depuis l'image |
| La version du MSI du complément provient désormais de la table des propriétés du MSI via l'objet COM `WindowsInstaller.Installer`. `Get-AppLockerFileInformation` — ce qu'utilise l'exemple de Microsoft lui-même — est absent de certaines éditions et, sous PowerShell 7, il fait intervenir la couche de compatibilité Windows PowerShell, qui échoue et inonde une exécution `-WhatIf` de sorties de copie de fichiers sans rapport |
| Les exécutions en mode application écrivent une transcription dans `C:\Temp\Update-TeamsClient_<timestamp>.log` ; les erreurs inattendues interrompent l'exécution au lieu de continuer à moitié ; le code de sortie est 0 en cas de succès (`-WhatIf` inclus) et 1 en cas d'échec |

### 2026-09-08
| Modification |
|--------|
| Ajout de [`scripts/Device/Update-TeamsClient.ps1`](scripts/Device/Update-TeamsClient.ps1) — réinstallation propre du nouveau Teams sur un poste ou un hôte de session AVD : désinstaller le Teams Meeting Add-in, supprimer le package AppX `MSTeams` pour tous les utilisateurs, télécharger `teamsbootstrapper.exe`, provisionner Teams (`-p`) et installer le MSI du complément de réunion fourni dans le nouveau package Teams |
| Chaque étape qui modifie l'état passe par `ShouldProcess`, si bien que `-WhatIf` parcourt tout le flux et affiche chaque désinstallation/téléchargement/installation sans toucher à la machine ; les étapes qui n'existent qu'après une vraie installation (nouvelle version de Teams, chemin du MSI du complément, vérification finale) sont signalées comme telles au lieu de faire échouer l'exécution |
| La recherche du complément lit à la fois la ruche de désinstallation 64 bits et `WOW6432Node` — le complément s'installe en 32 bits, la ruche 64 bits seule ne le trouve donc jamais (la désinstallation et la vérification le manquaient toutes deux auparavant) |
| Les codes de sortie et les codes de sortie de msiexec/du bootstrapper sont vérifiés au lieu d'être supposés ; `-SkipMeetingAddIn` ne remplace que le client, `-Force` installe sur un appareil sans aucun Teams. Intégré à [`menu.ps1`](menu.ps1) (touche T), qui utilise par défaut un aperçu `-WhatIf` |

### 2026-09-07 (2)
| Modification |
|--------|
| Ajout de [`scripts/SharePoint/Search-SharePointContent.ps1`](scripts/SharePoint/Search-SharePointContent.ps1) — l'équivalent Microsoft Graph de [`Find-SiteContent.ps1`](scripts/SharePoint/Find-SiteContent.ps1) : app-only, sans connexion interactive, et il recherche dans un site ou dans **chaque site et OneDrive du tenant** |
| La recherche de contenu utilise `/drives/{id}/root/delta` (toute l'arborescence d'une bibliothèque par pages de mille éléments, donc les caractères génériques `*contains*` fonctionnent) ou `/search/query` avec `-Content` pour le texte à l'intérieur des documents ; les filtres sont les mêmes que dans le script PnP |
| Les autorisations proviennent de `/drives/{id}/items/{id}/permissions`, 20 par appel `/$batch`. Un seul appel fournit les rôles, les identités bénéficiaires, le lien de partage avec sa portée (anyone/organization/specific people), modification ou lecture, expiration et URL, ainsi que `inheritedFrom` — ce qui détermine `PermissionSource = Item` (unique) ou `Inherited`. « Toute personne disposant du lien » a son propre compteur car ces liens ne nécessitent aucune connexion |
| Consigné explicitement, dans le script et dans le readme : Graph n'a pas d'API pour les attributions de rôles SharePoint, donc les droits au niveau du site et de la liste ainsi que les éléments des listes ordinaires (hors bibliothèques) restent du ressort de [`Find-SiteContent.ps1`](scripts/SharePoint/Find-SiteContent.ps1). Le readme contient un tableau comparatif pour choisir entre les deux |
| La connexion est app-only : la première exécution inscrit une application, donne le consentement au rôle d'application `Sites.Read.All`, crée un certificat auto-signé dans `CurrentUser\My` et en téléverse la clé publique — aucun secret sur disque — et met en cache l'ID client et l'empreinte par tenant dans `graph.appid.json` (ajouté à [`.gitignore`](.gitignore)). Les exécutions suivantes se connectent sans invite, cela fonctionne donc aussi depuis une tâche planifiée. La limitation (429) donne lieu à de nouvelles tentatives, en respectant `Retry-After`, aussi bien pour les appels uniques que pour les sous-requêtes de lot |

### 2026-09-07
| Modification |
|--------|
| Ajout de [`scripts/SharePoint/Find-SiteContent.ps1`](scripts/SharePoint/Find-SiteContent.ps1) — rechercher du contenu dans tout un site SharePoint ou OneDrive et indiquer quelles autorisations s'appliquent à chaque résultat. Lecture seule |
| Deux moteurs : une exploration de chaque liste et bibliothèque (voit tout, `-IncludeSubsites` pour les sous-sites) et une requête KQL sur l'index de recherche (`-Content`) qui trouve aussi du texte *à l'intérieur* des documents. Les deux partagent les filtres `-Name`, `-Path`, `-Extension`, `-ItemType`, `-ListName`, `-ModifiedBy`, `-ModifiedAfter`/`-ModifiedBefore` et `-MinSizeMB` |
| Pour chaque résultat, le script détermine d'où viennent les autorisations — l'élément lui-même (héritage rompu), sa liste ou le site — et aplatit les attributions de rôles en une ligne CSV par principal avec le type, l'identifiant, l'adresse e-mail et les noms de rôles. `Limited Access` est filtré sauf avec `-IncludeLimitedAccess` |
| Les liens de partage (les groupes `SharingLinks.*` derrière « Copier le lien ») sont toujours développés en la liste des personnes qu'ils contiennent et étiquetés Anyone/Organization/Specific people ; les invités externes (`#ext#`) et « Everyone (except external users) » sont signalés séparément dans le récapitulatif et dans le CSV |
| Les autorisations du site et des listes sont lues une seule fois et mises en cache, et les autorisations d'éléments uniquement pour les éléments ayant rompu l'héritage, si bien que le coût dépend du nombre de résultats et non de la taille du site ; `-Permissions Unique` ne signale que ce qui est partagé différemment, `-Permissions None` ignore les autorisations, et `-MaxPermissionLookups` plafonne une recherche trop large |
| Réutilise le flux d'inscription d'application et le cache `pnp.appid.json` par tenant de [`Restore-RecycleBinItems.ps1`](scripts/SharePoint/Restore-RecycleBinItems.ps1), et peut s'accorder temporairement les droits d'administrateur de collection de sites (`-GrantSiteAdmin`) pour rechercher dans un site ou un OneDrive sur lequel il n'a aucun droit |

### 2026-08-28
| Modification |
|--------|
| Ajout de [`scripts/SharePoint/`](scripts/SharePoint/readme.fr.md) avec [`Restore-RecycleBinItems.ps1`](scripts/SharePoint/Restore-RecycleBinItems.ps1) — restaure les fichiers/dossiers supprimés depuis la corbeille d'un site SharePoint ou d'un OneDrive, en essai à blanc par défaut, avec des filtres sur le nom, le dossier d'origine, l'auteur de la suppression et une fenêtre temporelle de suppression |
| Deux portées : `-SiteUrl` pour une seule collection de sites (OneDrive compris), ou `-AllSites -TenantUrl` pour parcourir tous les sites SharePoint du tenant. Le balayage du tenant exclut les sites personnels OneDrive, l'hôte My Site, les sites de redirection et les sites verrouillés, prend en charge `-SiteFilter`/`-MaxSites`, et continue lorsqu'un site isolé échoue — les résultats par site figurent dans un tableau récapitulatif et dans une colonne `Site` du CSV |
| Les restaurations s'effectuent par lots de 200 éléments maximum via `Restore-PnPRecycleBinItem -IdList` (un appel serveur par lot) au lieu d'un appel par élément ; dossiers et fichiers ne partagent jamais un lot, et un lot qui échoue dans son ensemble est relancé élément par élément afin que les erreurs par élément soient toujours signalées. Les runspaces parallèles ont été volontairement écartés — PnP PowerShell n'est pas thread-safe et des appels simultanés sur une même collection de sites déclenchent la limitation (throttling) de SharePoint |
| Le script indique la durée de chaque étape — le temps de lecture de la corbeille, une estimation préalable de la restauration, une barre de progression avec un temps restant estimé en direct à partir du débit mesuré, la durée réelle dans le récapitulatif, et une colonne `DurationSeconds` par élément dans le CSV |
| Le script enregistre sa propre application Entra lors de la première exécution sur un tenant (client public, `AllSites.FullControl` délégué, avec consentement administrateur), car PnP PowerShell ne fournit plus d'application multi-tenant partagée ; l'ID client est mis en cache par tenant dans `pnp.appid.json` (ajouté à [`.gitignore`](.gitignore)) |

### 2026-07-24 (3)
| Modification |
|--------|
| Mise hors service d'un dépôt PowerShell interne qui n'était plus maintenu (`Windows-Powershell`, dernier commit en mars 2023) : chaque script a été passé en revue et tout ce qui avait encore de la valeur a été modernisé dans ce dépôt — rien n'a été copié tel quel ; tout a été réécrit pour Microsoft Graph / Exchange Online (les scripts du dépôt source basés sur `MSOnline`/`AzureAD` ne fonctionnent plus du tout depuis que Microsoft a retiré ces points de terminaison) |
| Ajout de [`scripts/TenantOnboarding/`](scripts/TenantOnboarding/readme.fr.md) (24 scripts répartis entre Provisioning/MultiTenant/AppDeployment/DeviceConfig/OneDriveManagement/UserManagement) — modernisés à partir des scripts de configuration/onboarding de tenant du dépôt source |
| Ajout de [`scripts/Office365Toolkit/`](scripts/Office365Toolkit/readme.fr.md) (9 scripts répartis entre Security/Exchange/Intune) — modernisés à partir d'une copie forkée du projet GitHub retiré `directorcia/Office365` (CIAOPS) trouvée dans le dépôt source ; revus fonctionnalité par fonctionnalité et consolidés, pas portés 1:1 |
| Ajout de [`scripts/PatronToolkit/`](scripts/PatronToolkit/readme.fr.md) (13 scripts répartis entre Entra/Security/Exchange/Intune/SharePoint/Teams) — modernisés à partir d'une copie forkée du projet GitHub retiré `directorcia/patron` trouvée dans le dépôt source, selon la même approche de consolidation par fonctionnalité |
| Ajout de [`scripts/LegacyUtilities/`](scripts/LegacyUtilities/readme.fr.md) (17 scripts répartis entre Exchange/Entra/Teams/Network/Device/Workspace365) — modernisés à partir de divers petits outils du dépôt source non couverts par ce qui précède |
| Règle de traitement des données appliquée partout : les dossiers `Klanten`, `created-users`, `csv files` et `Archief` du dépôt source (vrais noms de clients/domaines de tenant/mots de passe générés) n'ont jamais été lus ni portés ; tout autre script qui codait en dur de vrais identifiants de client/tenant ou des secrets a été généralisé en paramètres, ou purement écarté — voir le readme de chaque nouveau dossier pour sa liste d'exclusions spécifique |
| Aucun des ~63 nouveaux scripts n'est relié à [`menu.ps1`](menu.ps1) — ce sont des scripts d'audit/de reporting/de configuration destinés à être exécutés directement, conformément au modèle existant pour [`scripts/RDS/`](scripts/RDS/readme.fr.md), [`scripts/Azure/`](scripts/Azure/readme.fr.md) et [`scripts/Network/UniFi/`](scripts/Network/UniFi/readme.fr.md) |

### 2026-07-24 (2)
| Modification |
|--------|
| Correction de [`scripts/Reporting/Get-SharePointStorageReport.ps1`](scripts/Reporting/Get-SharePointStorageReport.ps1) qui abandonnait silencieusement les recherches d'historique des versions sur les très grandes bibliothèques : le plafond de passes de relance de `Invoke-GraphBatchGet` était codé en dur à 8, alors que la limitation d'activité par application de SharePoint Online n'autorise qu'environ 1500 à 2500 recherches de versions résolues par passe avant qu'une pause de ~60-90 s ne se répète — sur un tenant de 200k fichiers, environ 90 % des fichiers étaient ainsi marqués « abandonnés » avant que l'analyse ne soit réellement terminée |
| Application de la même correction à la copie de cette même fonction de relance par lots dans [`scripts/Reporting/Remove-SharePointFileVersionsByDate.ps1`](scripts/Reporting/Remove-SharePointFileVersionsByDate.ps1) (`Get-FileVersionsBatch`), qui avait le même plafond de 8 passes codé en dur |
| Ajout du paramètre `-MaxVersionRetryPasses` aux deux scripts (par défaut `0` = le plafond de passes s'adapte automatiquement au volume de requêtes, limité à 500 passes) ; un `Write-Warning` explicite indique désormais exactement combien de fichiers ont été abandonnés et suggère le paramètre si le plafond est encore atteint |
| Mise à jour de [`scripts/Reporting/readme.md`](scripts/Reporting/readme.md) pour documenter `-VersionBatchConcurrency` et `-MaxVersionRetryPasses` pour les deux scripts |

### 2026-07-24 (1)
| Modification |
|--------|
| Documentation de [`scripts/Azure/VM/Azure-NVMe-Conversion.ps1`](scripts/Azure/VM/Azure-NVMe-Conversion.ps1) (auparavant absent de tout readme/arborescence) : ajout de [`scripts/Azure/readme.md`](scripts/Azure/readme.md) et [`scripts/Azure/VM/readme.md`](scripts/Azure/VM/readme.md), ainsi que d'une nouvelle catégorie « Azure Infrastructure » dans ce readme |
| Raccordement à [`menu.ps1`](menu.ps1) de 5 scripts entièrement documentés mais jusque-là absents du menu : [`Move-InboxToArchive.ps1`](scripts/Exchange/Move-InboxToArchive.ps1) et [`Set-Distributionlist-dynamic-static.ps1`](scripts/Exchange/Set-Distributionlist-dynamic-static.ps1) (sous-menu Exchange, touches E/F), [`Get-M365UserLicenses.ps1`](scripts/Entra/Get-M365UserLicenses.ps1), [`Import-ConditionalAccessBaseline.ps1`](scripts/Entra/Import-ConditionalAccessBaseline.ps1), [`Set-UserManager.ps1`](scripts/Entra/Set-UserManager.ps1) (sous-menu Entra, touches G/H/I) — et ajout de ceux-ci aux tableaux Menu de ce readme |
| Ajout de [`scripts/Device/Remove-OemBloatware.ps1`](scripts/Device/Remove-OemBloatware.ps1) — détecte HP/Lenovo/Dell et supprime les logiciels superflus du constructeur via winget, ainsi que les applications inutiles génériques du Microsoft Store via AppX ; essai à blanc par défaut ; relié à [`menu.ps1`](menu.ps1) (touche I) |
| Ajout de [`scripts/Device/DriveMapping/New-CloudDriveMapping.ps1`](scripts/Device/DriveMapping/New-CloudDriveMapping.ps1) — mappe des bibliothèques de documents SharePoint/OneDrive sur des lettres de lecteur via WebDAV, pour une utilisation comme script d'ouverture de session ; essai à blanc par défaut |
| Ajout de [`scripts/Network/UniFi/`](scripts/Network/UniFi/readme.fr.md) — [`UnifiApi.ps1`](scripts/Network/UniFi/UnifiApi.ps1), un assistant de connexion partagé (contrôleur classique + détection automatique d'UniFi OS), [`Get-UnifiNetworkReport.ps1`](scripts/Network/UniFi/Get-UnifiNetworkReport.ps1) (rapport de documentation HTML), [`Update-UnifiFirmware.ps1`](scripts/Network/UniFi/Update-UnifiFirmware.ps1) (outil de mise à niveau du firmware en essai à blanc) ; identifiants toujours via `Get-Credential`, jamais codés en dur |
| Ajout de [`scripts/Intune/Compare-IntuneConfig.ps1`](scripts/Intune/Compare-IntuneConfig.ps1) — détection des écarts de configuration Intune entre un tenant client et une sauvegarde de référence MSP, via le module `IntuneBackupAndRestore` ; lecture seule |

### 2026-07-22 (6)
| Modification |
|--------|
| Mise à jour de [`scripts/Reporting/Get-SharePointStorageReport.ps1`](scripts/Reporting/Get-SharePointStorageReport.ps1) pour rendre les analyses complètes compatibles GDAP, en résolvant et en figeant un seul contexte de tenant effectif (`-TenantId`, ou le contexte client GDAP issu de `$global:cid`) pour la connexion Graph, la création de l'application temporaire et l'émission du jeton app-only |
| Ajout de garde-fous pour les flux GDAP/app-only : erreurs plus claires lorsque le contexte du tenant client est manquant (exécutez d'abord `Connect-Tenant` ou passez `-TenantId`) et lorsque `-ClientId` est fourni sans ID de tenant résolvable |
| Mise à jour des entrées de la signature de point de reprise dans [`Get-SharePointStorageReport.ps1`](scripts/Reporting/Get-SharePointStorageReport.ps1) pour inclure `-ForceAppOnlySingleSite` et le contexte de tenant résolu, afin d'éviter des collisions de reprise entre contextes |
| Mise à jour de [`scripts/Reporting/readme.md`](scripts/Reporting/readme.md) pour documenter le rattachement au tenant GDAP lors des analyses complètes |

### 2026-07-22 (5)
| Modification |
|--------|
| Mise à jour de [`scripts/Reporting/Get-SharePointStorageReport.ps1`](scripts/Reporting/Get-SharePointStorageReport.ps1) pour rendre les analyses d'un seul site compatibles GDAP : lorsque `authMode=GDAP` est détecté, le script utilise désormais automatiquement le chemin d'amorçage application temporaire/app-only pour les analyses `-SiteUrl`, afin d'éviter les lacunes des autorisations déléguées |
| Ajout du paramètre `-ForceAppOnlySingleSite` pour forcer explicitement l'amorçage app-only lors des analyses d'un seul site, indépendamment du mode d'authentification détecté |
| Mise à jour de [`scripts/Reporting/readme.md`](scripts/Reporting/readme.md) pour documenter le nouveau comportement GDAP et le paramètre `-ForceAppOnlySingleSite` |

### 2026-07-22 (4)
| Modification |
|--------|
| Renforcement du chemin délégué de [`scripts/Reporting/Get-SharePointStorageReport.ps1`](scripts/Reporting/Get-SharePointStorageReport.ps1) pour supprimer la dépendance à des cmdlets Graph manquantes : remplacement de l'utilisation de `Get-MgDriveItemChild` par une pagination REST Graph via `Invoke-MgGraphRequest` pour le parcours des enfants d'un drive |
| Mise à jour des solutions de repli pour l'énumération des drives d'un site dans [`Get-SharePointStorageReport.ps1`](scripts/Reporting/Get-SharePointStorageReport.ps1), qui utilisent désormais l'API REST Graph (`/sites/{id}/drives`) au lieu de `Get-MgSiteDrive` dans les branches déléguées/non app-only |
| Mise à jour de la récupération de la corbeille dans [`Get-SharePointStorageReport.ps1`](scripts/Reporting/Get-SharePointStorageReport.ps1), qui utilise désormais la pagination REST Graph (`/sites/{id}/recycleBin/items`) comme chemin délégué principal/de repli au lieu de dépendre de la cmdlet `Get-MgSiteRecycleBinItem` |
| Suppression des avertissements MSAL non bloquants sur l'autorité lors de la déconnexion, en encapsulant `Disconnect-MgGraph` avec une suppression temporaire des avertissements lors du nettoyage |

### 2026-07-22 (3)
| Modification |
|--------|
| Mise à jour de la gestion du mode d'analyse de [`scripts/Reporting/Get-SharePointStorageReport.ps1`](scripts/Reporting/Get-SharePointStorageReport.ps1) afin que les exécutions sur un seul site (`-SiteUrl` avec `/sites/...` ou `/teams/...`) ne déclenchent plus l'enregistrement d'une application temporaire + l'amorçage app-only ; la configuration app-only n'est désormais utilisée que pour l'énumération à l'échelle du tenant |
| Amélioration des performances et de la fiabilité de la recherche d'un seul site en résolvant directement le site exact via le chemin d'URL Graph (`/sites/{hostname}:{path}`) au lieu d'un flux de recherche/filtrage |
| Mise à jour des notes de performance de [`scripts/Reporting/readme.md`](scripts/Reporting/readme.md) pour documenter le chemin optimisé pour un seul site et la vitesse de démarrage attendue |

### 2026-07-22 (2)
| Modification |
|--------|
| Correction d'un problème de dépendance du rapport SharePoint provoquant des erreurs de commande introuvable pour `Get-MgSite` : mise à jour de [`load.ps1`](load.ps1), [`scripts/Startup/Install-Modules.ps1`](scripts/Startup/Install-Modules.ps1) et [`scripts/Startup/Update-Modules.ps1`](scripts/Startup/Update-Modules.ps1) pour inclure `Microsoft.Graph.Sites` |
| Mise à jour de [`scripts/Reporting/Get-SharePointStorageReport.ps1`](scripts/Reporting/Get-SharePointStorageReport.ps1) avec une vérification préalable explicite des modules `Microsoft.Graph.Authentication` et `Microsoft.Graph.Sites`, avec une indication d'installation claire lorsque des modules manquent |
| Mise à jour de la documentation des dépendances de modules dans [`scripts/Startup/readme.md`](scripts/Startup/readme.md) pour inclure `Microsoft.Graph.Sites` dans les prérequis d'installation/de mise à jour |

### 2026-07-22 (1)
| Modification |
|--------|
| Mise à jour de [`load.ps1`](load.ps1) — ajout des switches de lancement au démarrage `-SetupStartup` / `-RemoveStartup` ; la configuration de première exécution enregistre désormais les valeurs par défaut de l'authentification déléguée (`authMode`, `defaultCustomerDomain` facultatif, `useDeviceCodeAuth`) dans `load.config.ps1` |
| Mise à jour de [`menu.ps1`](menu.ps1) — ajout des actions Startup `F` (Enable-LauncherStartup) et `G` (Disable-LauncherStartup) ; ajout de l'action M365 `H` (Test-GdapConnection) ; le sous-menu Entra comprend désormais les actions CA temporaire et TAP (`D`/`E`/`F`) |
| Mise à jour de [`scripts/Startup/functies.ps1`](scripts/Startup/functies.ps1) — la connexion Graph au démarrage prend désormais en charge la préférence pour l'authentification déléguée par code d'appareil + les scopes requis pour le flux GDAP ; ajout de l'assistant `Test-GdapConnection` pour les vérifications de contrat délégué/de connectivité |
| Mise à jour de la maintenance des modules au démarrage : [`scripts/Startup/Install-Modules.ps1`](scripts/Startup/Install-Modules.ps1) et [`scripts/Startup/Update-Modules.ps1`](scripts/Startup/Update-Modules.ps1) incluent désormais `Microsoft.Graph.Identity.DirectoryManagement` |
| Ajout de [`scripts/Entra/New-TemporaryConditionalAccessPolicy.ps1`](scripts/Entra/New-TemporaryConditionalAccessPolicy.ps1) — crée une stratégie CA temporaire pour un utilisateur/groupe, avec soit une fenêtre basée sur une durée, soit une date/heure locale exacte de début/fin ; nettoyage automatique facultatif dans la même session à l'heure de fin |
| Ajout de [`scripts/Entra/Remove-TemporaryConditionalAccessPolicies.ps1`](scripts/Entra/Remove-TemporaryConditionalAccessPolicies.ps1) — supprime une ou plusieurs stratégies CA temporaires (`TEMP-CA -`), y compris en mode expirées uniquement ou suppression totale |
| Ajout de [`scripts/Entra/New-UserTemporaryAccessPass.ps1`](scripts/Entra/New-UserTemporaryAccessPass.ps1) — crée un Temporary Access Pass (TAP) pour un utilisateur, avec une durée de vie configurable et une option d'usage unique |
| Mise à jour de la documentation de ce qui précède dans [`readme.md`](readme.md), [`scripts/readme.md`](scripts/readme.md), [`scripts/Startup/readme.md`](scripts/Startup/readme.md) et [`scripts/Entra/readme.md`](scripts/Entra/readme.md) ; précision que le nettoyage automatique des CA temporaires s'exécute dans la session en cours (aucune tâche planifiée n'est créée) |

### 2026-07-09 (8)
| Modification |
|--------|
| [`scripts/Reporting/Get-SharePointStorageReport.ps1`](scripts/Reporting/Get-SharePointStorageReport.ps1) — ajout d'une phase « totaux par collection de sites » (`-Apply` uniquement, Phase 2c) : les sous-sites/Teams-kanalen et la corbeille sont désormais automatiquement agrégés par collection de sites racine dans `SharePoint_SiteCollectionTotals_<timestamp>.csv`, de sorte que le total général est directement comparable au chiffre unique « stockage utilisé » affiché par site dans le centre d'administration SharePoint |
| Documentation dans [`scripts/Reporting/readme.md`](scripts/Reporting/readme.md) de la nouvelle sortie et des causes les plus probables d'un écart restant avec le chiffre du portail d'administration (décalage temporel, dossiers ignorés silencieusement en cas d'erreurs d'autorisation, échecs de recherche de versions) |

### 2026-07-09 (7)
| Modification |
|--------|
| [`scripts/Reporting/Get-SharePointStorageReport.ps1`](scripts/Reporting/Get-SharePointStorageReport.ps1) — extension brève des recherches dans la corbeille (Phase 2b de `-Apply` et `-RecycleBinOnly`) aux sites personnels OneDrive, puis annulation le jour même sur demande — la portée de la corbeille reste limitée aux collections de sites SharePoint, OneDrive reste entièrement exclu (analyse du stockage comme corbeille) |
| Ajout d'une section « Prullenbak (recycle bin) » à [`scripts/Reporting/readme.md`](scripts/Reporting/readme.md) documentant la portée de la corbeille (SharePoint uniquement) |

### 2026-07-09 (6)
| Modification |
|--------|
| Suppression complète du dossier conteneur `Testing Scripts/` — ses sous-dossiers dupliquaient les noms de catégories de premier niveau existants par verbe (préfixe `Test-`/`Get-`) plutôt que par domaine. Son contenu a été fusionné dans le dossier de domaine correspondant : `Testing Scripts/Entra/Test-M365GroupMembership.ps1` → [`Entra/`](scripts/Entra/readme.fr.md), `Testing Scripts/Exchange/*` (6 scripts) → [`Exchange/`](scripts/Exchange/readme.fr.md), `Testing Scripts/Device/Test-OpenVpnDiagnostics.ps1` → [`Device/`](scripts/Device/readme.fr.md). [`Network/`](scripts/Network/readme.fr.md), [`RDS/`](scripts/RDS/readme.fr.md) et [`SMTP/`](scripts/SMTP/readme.fr.md) (sans équivalent de premier niveau existant) ont été promus au rang de dossiers de catégorie de premier niveau à part entière |
| Fusion des readmes correspondants dans le readme.md existant de chaque dossier de destination, au lieu de conserver des documents « Testing — » séparés |
| Mise à jour de 10 chemins de scripts dans [`menu.ps1`](menu.ps1) (sous-menu d'audit Exchange, sous-menu d'audit Entra, Test-Ports, tests SMTP) pour les nouveaux emplacements |

### 2026-07-09 (5)
| Modification |
|--------|
| Rangement de `Testing Scripts/` et de la racine du dépôt : déplacement de `vias_archiver.ps1` de `Testing Scripts/Device/` vers une nouvelle catégorie [`scripts/Teams/`](scripts/Teams/readme.fr.md) — il s'agit d'un outil d'export et d'archivage Teams/SharePoint, pas d'un script de diagnostic, il n'avait donc pas sa place sous « Testing » |
| Déplacement de [`Update-modules.ps1`](scripts/Startup/Update-Modules.ps1), situé à la racine, vers [`scripts/Startup/`](scripts/Startup/readme.fr.md) (renommé [`Update-Modules.ps1`](scripts/Startup/Update-Modules.ps1) par cohérence de nommage) — c'est un script de maintenance des modules comme [`Install-Modules.ps1`](scripts/Startup/Install-Modules.ps1), pas un point d'entrée du dépôt comme [`load.ps1`](load.ps1)/[`menu.ps1`](menu.ps1) |
| Suppression du dossier `Testing Scripts/SharePoint/` (il ne contenait qu'un readme de renvoi, aucun script) — ce renvoi figure désormais directement dans `Testing Scripts/readme.md` |
| Documentation de [`Test-PowerShellSyntax.ps1`](scripts/Startup/Test-PowerShellSyntax.ps1) dans [`scripts/Startup/readme.md`](scripts/Startup/readme.md), qui n'avait auparavant aucune documentation |

### 2026-07-09 (4)
| Modification |
|--------|
| Optimisation des recherches d'historique des versions de [`scripts/Reporting/Get-SharePointStorageReport.ps1`](scripts/Reporting/Get-SharePointStorageReport.ps1), principale raison pour laquelle le script semblait bloqué sur les grandes bibliothèques (un appel Graph séquentiel par fichier, chacun pouvant faire l'objet de 6 relances avec un délai d'attente allant jusqu'à ~2 minutes en cas de limitation) : (1) la recherche est entièrement ignorée lorsqu'on sait avec certitude que le contrôle de version est désactivé sur une bibliothèque, (2) jusqu'à 20 recherches de versions de fichiers sont regroupées par appel HTTP via le point de terminaison `$batch` de Graph au lieu d'un appel par fichier, (3) une politique de relance courte et peu coûteuse à 3 tentatives est utilisée pour ces appels précis au lieu de la politique principale de relance/délai, puisqu'une recherche échouée se replie sans risque sur « 0 versions » |
| Suppression de la fonction `Get-VersionSize`, devenue inutile, remplacée par `Invoke-GraphBatchGet` + la résolution par lots dans `Get-AllDriveItems` |
| Mise à jour de [`scripts/Reporting/readme.md`](scripts/Reporting/readme.md) avec une section « Performance » documentant ce qui précède |

### 2026-07-09 (3)
| Modification |
|--------|
| Retour de [`Deploy-OfficeTheme.ps1`](scripts/Custom%20Scripts/Intune/Desktop/Deploy-OfficeTheme.ps1), [`Deploy-Officecolors.ps1`](scripts/Custom%20Scripts/Intune/Desktop/Office%20Themes/Deploy-Officecolors.ps1) et de leurs ressources de thème (`2026 Vias institute colours (2).thmx`, `Office Themes/`) dans [`Custom Scripts/Intune/Desktop/`](scripts/Custom%20Scripts/Intune/Desktop/readme.fr.md) — les deux scripts codent en dur leur URL de téléchargement vers ce chemin exact du dépôt, ce qui permet de garder l'URL valide au lieu d'exiger une mise à jour du script + un redéploiement Intune. [`Custom Scripts/`](scripts/Custom%20Scripts/readme.fr.md) et [`Custom Scripts/Intune/`](scripts/Custom%20Scripts/Intune/readme.fr.md) ont été recréés comme dossiers minimaux figés sur ce chemin (juste cet élément) plutôt que comme l'ancienne catégorie complète |
| [`scripts/Intune/Desktop/`](scripts/Intune/Desktop/readme.fr.md) ne contient désormais que le déploiement du fond d'écran/de l'écran de verrouillage/des raccourcis de la barre des tâches ; `Intune/readme.md` et `Intune/Desktop/readme.md` renvoient tous deux à [`Custom Scripts/Intune/Desktop/`](scripts/Custom%20Scripts/Intune/Desktop/readme.fr.md) pour les scripts de thème Office |

### 2026-07-09 (2)
| Modification |
|--------|
| Suppression du dossier conteneur [`Custom Scripts/`](scripts/Custom%20Scripts/readme.fr.md) — il mélangeait des outils génériques et des scripts propres à des clients sous une étiquette confuse, et faisait doublon avec la catégorie [`Intune/`](scripts/Intune/readme.fr.md). Le contenu a été redistribué dans les catégories de premier niveau appropriées : `Custom Scripts/device/` → [`Device/`](scripts/Device/readme.fr.md), `Custom Scripts/DNS/` → [`DNS/`](scripts/DNS/readme.fr.md), `Custom Scripts/SAS/` → [`SAS/`](scripts/SAS/readme.fr.md), `Custom Scripts/Save install time/` → [`Deployment/`](scripts/Deployment/readme.fr.md) (renommé), [`Custom Scripts/Intune/Desktop/`](scripts/Custom%20Scripts/Intune/Desktop/readme.fr.md) → fusionné dans [`Intune/Desktop/`](scripts/Intune/Desktop/readme.fr.md) |
| Mise à jour des chemins de scripts dans [`menu.ps1`](menu.ps1) pour [`Restart-Time-Sync.ps1`](scripts/Device/Time%20sync/Restart-Time-Sync.ps1), [`detect-audiodevices.ps1`](scripts/Device/audio/detect-audiodevices.ps1), [`Disable-internalmic.ps1`](scripts/Device/audio/Disable-internalmic.ps1) vers leur nouvel emplacement [`scripts/Device/`](scripts/Device/readme.fr.md) |
| Mise à jour des renvois dans [`scripts/Intune/readme.md`](scripts/Intune/readme.md), [`scripts/Intune/Get-Autopilot/readme.md`](scripts/Intune/Get-Autopilot/readme.md) et [`scripts/readme.md`](scripts/readme.md) pour les nouveaux emplacements de dossiers |
| ~~**Problème connu (volontaire) :** [`Deploy-OfficeTheme.ps1`](scripts/Custom%20Scripts/Intune/Desktop/Deploy-OfficeTheme.ps1) et [`Deploy-Officecolors.ps1`](scripts/Custom%20Scripts/Intune/Desktop/Office%20Themes/Deploy-Officecolors.ps1) codent toujours en dur leur URL de téléchargement vers l'ancien chemin — laissé inchangé sur demande.~~ **Résolu ci-dessus** — les scripts ont plutôt été remis à l'emplacement correspondant à leur URL codée en dur. |

### 2026-07-09
| Modification |
|--------|
| Ajout d'un [`readme.md`](readme.md) dans chaque dossier qui n'en avait pas : [`scripts/`](scripts/readme.fr.md), [`scripts/Custom Scripts/`](scripts/Custom%20Scripts/readme.fr.md), [`scripts/Custom Scripts/Intune/`](scripts/Custom%20Scripts/Intune/readme.fr.md) (+ `Desktop/`, `Office Themes/`, `Add Lockscreen to start and desktop/`, `Background/`), `scripts/Custom Scripts/device/Time sync/`, [`scripts/Graph/`](scripts/Graph/readme.fr.md), [`scripts/Intune/`](scripts/Intune/readme.fr.md) (+ `Get-Autopilot/`), `scripts/Testing Scripts/`, `scripts/Testing Scripts/Network/`, `scripts/Testing Scripts/RDS/` — chacun avec une liste de fichiers et une documentation des paramètres/de l'utilisation |
| Correction de `scripts/Custom Scripts/device/audio/Rollback-InternalMic` — il manquait l'extension `.ps1` au fichier |
| Renommage de `scripts/Entra/remove-m365users.ps1` → [`Remove-M365Users.ps1`](scripts/Entra/Remove-M365Users.ps1) par cohérence de nommage (menu.ps1 et les readmes faisaient déjà référence à la forme PascalCase) |
| Correction de [`scripts/Entra/readme.md`](scripts/Entra/readme.md) — suppression d'une entrée obsolète `Distributionlist.ps1` qui documentait en réalité [`scripts/Exchange/Set-Distributionlist-dynamic-static.ps1`](scripts/Exchange/Set-Distributionlist-dynamic-static.ps1) ; déplacement de la documentation exacte vers [`scripts/Exchange/readme.md`](scripts/Exchange/readme.md) ; ajout de la documentation manquante de [`Set-UserManager.ps1`](scripts/Entra/Set-UserManager.ps1) |
| Correction de `scripts/Testing Scripts/SharePoint/readme.md` — c'était un doublon obsolète de la documentation de [`Get-SharePointStorageReport.ps1`](scripts/Reporting/Get-SharePointStorageReport.ps1) (le script ne se trouve pas dans ce dossier) ; remplacé par un renvoi vers [`scripts/Reporting/readme.md`](scripts/Reporting/readme.md), qui documente désormais l'ensemble actuel des paramètres du script (`-ClientId`, `-ClientSecret`, `-CertificateThumbprint`, `-RecycleBinOnly`, `-GraphTimeoutSec`, `-MaxGraphRetry` n'étaient auparavant pas documentés) |
| Correction de `scripts/Custom Scripts/device/audio/readme.md` — casse des noms de scripts corrigée pour correspondre aux fichiers réels sur le disque |
| Suppression des fichiers `.DS_Store` suivis dans git et ajout de `.DS_Store` à [`.gitignore`](.gitignore) |

### 2026-04-17
| Modification |
|--------|
| Mise à jour de `scripts/Custom Scripts/Save install time/start.bat` — ajout de l'option `D` (scripts d'installation client depuis le dossier local `Install`) et de l'option `E` (scripts d'installation client depuis un partage réseau) ; avant le début du déploiement, crée/met à jour l'administrateur local `LocalAdmin` (`<mot de passe omis>`), l'ajoute à `Administrators` et définit les indicateurs de saut de l'OOBE |
| Ajout de `scripts/Custom Scripts/Save install time/Browse-InstallScripts.ps1` — navigateur axé sur les clients qui affiche les dossiers clients comme éléments de menu et lance les scripts `.ps1`, `.bat` et `.cmd` |
| Mise à jour de `scripts/Custom Scripts/Save install time/readme.md` — documentation des nouvelles options de menu `D`/`E`, en précisant que l'option `D` nécessite de copier à la fois [`Browse-InstallScripts.ps1`](scripts/Deployment/Browse-InstallScripts.ps1) et l'intégralité du dossier `Install`, et que les options de déploiement client préparent `LocalAdmin` ainsi que les indicateurs de saut de l'OOBE |

### 2026-04-16
| Modification |
|--------|
| Mise à jour de `scripts/Custom Scripts/Intune/Desktop/Background/Lockscreen/Make-lockscreen.ps1` en v2.0 — source de l'écran de verrouillage alignée sur la configuration du fond d'écran d'entreprise (`$ImageUrl`, `$ClientName`), remplacement du `WebClient` direct par un flux de téléchargement Internet validé (`Invoke-WebRequest`), ajout de la normalisation des URL GitHub blob/raw, de vérifications de signature d'image (`jpg/png/bmp`), d'une protection contre les réponses HTML, d'une journalisation Intune structurée et d'une gestion plus sûre des téléchargements temporaires |
| Ajout de `scripts/Custom Scripts/Intune/Desktop/Background/Lockscreen/readme.md` — documentation de la configuration, du déploiement, de la journalisation, du flux de travail et de l'historique des versions propre à l'écran de verrouillage |
| Mise à jour du [`readme.md`](readme.md) racine — documentation Intune Desktop/Background étendue et structure du dépôt complétée pour inclure le script d'écran de verrouillage et sa documentation |

### 2026-04-15
| Modification |
|--------|
| Mise à jour de `scripts/Custom Scripts/SAS/Test-SASWorkDirectory.ps1` — ajout de diagnostics du profil des lecteurs avec prise en compte explicite des lecteurs éphémères (`G:`/`U:`), journalisation plus détaillée des échecs d'E/S (type d'exception, exception interne, HResult) et seuil d'avertissement d'espace faible adaptatif pour les disques de travail éphémères |
| Mise à jour de `scripts/Custom Scripts/SAS/Test-SASWorkDirectory.ps1` — affinage de l'analyse des événements Application SAS : classification explicite des refus d'accès `hc_disk_delete*` (`Return code 5`) comme échecs nécessitant une action, dédoublonnage des événements SAS répétés, et traitement informatif du bruit de télémétrie `ARM Application data not available` |
| Mise à jour de `scripts/Custom Scripts/SAS/Test-SASWorkDirectory.ps1` — ajout de diagnostics AV/EDR : état et exclusions de Defender, analyse des événements Defender Operational, analyse des événements System FilterManager/WdFilter et instantané des minifiltres actifs (`fltmc`) pour corrélation avec les refus d'accès intermittents lors de la suppression de WORK |
| Mise à jour de `scripts/Custom Scripts/SAS/readme.md` — documentation du comportement des disques éphémères pour les `WORK`/`USERWORK` SAS, et explication de la raison pour laquelle basculer entre `G:` et `U:` n'est pas une stratégie de basculement à long terme lorsque les deux sont éphémères |
| Ajout de `scripts/Custom Scripts/SAS/rca.md` — analyse formelle des causes profondes des échecs intermittents de suppression de WORK SAS, comprenant la chronologie des éléments de preuve, les constats AV/ASR, l'évaluation de la cause profonde et le plan de remédiation |

### 2026-04-13
| Modification |
|--------|
| Mise à jour de `scripts/Testing Scripts/Device/vias_archiver.ps1` en v8.11 — l'étape 10 prend désormais en charge un mode non interactif via `-Step10Only -Step10Action undo|archive|skip` ; ajout d'un avertissement explicite indiquant que l'archivage/le désarchivage dans Teams est une action au niveau de l'équipe (pas par canal) |
| Mise à jour de `scripts/Testing Scripts/Device/vias_archiver.ps1` en v8.12 — ajout d'un mode d'archivage souple par canal à l'étape 10 (marqueur de renommage avec annulation), y compris les paramètres du mode rapide `-Step10Only -ChannelAction archive|undo` et l'option `-ChannelArchiveTag` |
| Mise à jour de `scripts/Testing Scripts/Device/vias_archiver.ps1` en v8.13 — le flux par canal utilise désormais la véritable API Graph d'archivage/désarchivage de canal, ajout du scope de canal `ChannelSettings.ReadWrite.All` et d'un repli facultatif par renommage via `-ChannelFallbackToRename` |
| Mise à jour de `scripts/Testing Scripts/Device/vias_archiver.ps1` en v8.14 — ajout d'un mode `-DryRun` pour l'étape 10 afin de simuler l'archivage/le désarchivage d'équipe/de canal (et le repli facultatif par renommage) sans rien modifier |
| Mise à jour de `scripts/Testing Scripts/Device/vias_archiver.ps1` en v8.15 — `-DryRun` étendu à l'ensemble du script : les actions de configuration/export/rapport/nettoyage qui modifient quelque chose sont ignorées, tandis que la vérification et la sortie simulée de l'étape 10 sont conservées |
| Mise à jour de `scripts/Testing Scripts/Device/vias_archiver.ps1` en v8.16 — `-DryRun` conserve désormais la création de l'application temporaire, l'amorçage des autorisations, la connexion complète et le flux d'export/de rapport ; seules les modifications d'archivage/de désarchivage de l'étape 10 restent simulées |
| Mise à jour de `scripts/Testing Scripts/Device/vias_archiver.ps1` en v8.17 — affinage de `-DryRun` pour les étapes 6 à 9 afin de valider l'existence/les comptages (Teams/SharePoint/Graph) sans écrire d'exports ; le rapport de l'étape 11 utilise désormais ces comptages de sondage |
| Mise à jour de `scripts/Testing Scripts/Device/vias_archiver.ps1` en v8.18 — fiabilité de la recherche de canaux corrigée (valeurs Team/Channel d'Excel nettoyées des espaces et réutilisation du résolveur de canaux mis en cache avec repli Graph à l'étape 9) afin de réduire les faux « Kanaal niet gevonden » en essai à blanc |
| Mise à jour de `scripts/Testing Scripts/Device/vias_archiver.ps1` en v8.19 — ajout d'une correspondance normalisée des noms de canaux (espaces de début/fin, espaces internes, casse) dans le cache des canaux + la recherche de repli Graph, pour mieux gérer les différences subtiles de noms lors d'un essai à blanc |

### 2026-04-09
| Modification |
|--------|
| Mise à jour de `scripts/Custom Scripts/Intune/Desktop/Background/Desktop/Set-CorporateWallpaper.ps1` v2.1 — ajout de `Set-ExecutionPolicy Bypass -Scope Process` en tête du script pour éviter le code de sortie 3 lorsque la stratégie d'exécution d'Intune bloque le script |

### 2026-04-08
| Modification |
|--------|
| Mise à jour de [`scripts/Reporting/Licensing/genereer_licentie_overzicht.py`](scripts/Reporting/Licensing/genereer_licentie_overzicht.py) — lorsqu'un même produit apparaît avec plusieurs périodes de facturation sur une même facture (Pax8 et Ingram), chaque période est désormais affichée sur une ligne distincte avec la plage de la période dans la colonne Category/Detail, au lieu d'être additionnée de façon incorrecte |
| Mise à jour de [`scripts/Reporting/Licensing/genereer_rapport.ps1`](scripts/Reporting/Licensing/genereer_rapport.ps1) — la fenêtre CMD se ferme désormais automatiquement lors d'une exécution en tâche planifiée ; les pauses `Read-Host` sont ignorées lorsque `[Environment]::UserInteractive` vaut false |
| Mise à jour de [`scripts/Reporting/Licensing/genereer_rapport.bat`](scripts/Reporting/Licensing/genereer_rapport.bat) — redirige stdin depuis `NUL` afin que la pause interactive de Python ne soit jamais déclenchée lors d'une exécution en tâche planifiée |

### 2026-04-01
| Modification |
|--------|
| Mise à jour de [`scripts/Exchange/Migrate-Calendar.ps1`](scripts/Exchange/Migrate-Calendar.ps1) — correction des problèmes de réservation des boîtes aux lettres de salle : délai d'attente de provisionnement porté de 15 s à 60 s ; ajout d'une boucle de relance (5×30 s) pour `Set-CalendarProcessing` avec gestion des erreurs et instructions de repli ; `BookingWindowInDays 0` remplacé par `1825` et ajout de `EnforceSchedulingHorizon $false` pour éviter les refus de réservation silencieux |

### 2026-03-30
| Modification |
|--------|
| Mise à jour de `scripts/Custom Scripts/Intune/Desktop/Background/Desktop/Set-CorporateWallpaper.ps1` — ajout d'une vérification d'idempotence : télécharge l'image dans un dossier temporaire et compare son hachage SHA256 à celui du fichier existant ; ignore l'opération si le hachage correspond et que PersonalizationCSP est correct ; applique (sans second téléchargement) si l'image est nouvelle ou modifiée |
| Mise à jour du readme — section Intune & Autopilot : documentation de [`Set-CorporateWallpaper.ps1`](scripts/Intune/Desktop/Background/Desktop/Set-CorporateWallpaper.ps1) étendue avec un tableau de configuration, les étapes de déploiement, le chemin du journal et les instructions de déploiement NinjaOne/Intune |

### 2026-03-27
| Modification |
|--------|
| Ajout de [`scripts/Reporting/Get-ComputerLastLogon.ps1`](scripts/Reporting/Get-ComputerLastLogon.ps1) — rapport de dernière ouverture de session pour les ordinateurs d'une ou plusieurs OU ; mode `LastLogonTimestamp` (rapide) ou `-AllDCs` (précis) ; marque les ordinateurs Active/Stale/Never/Disabled ; exporte un CSV horodaté vers `C:\Temp\` ; paramètres `-InactiveDays`, `-IncludeDisabled`, `-ExportPath` |
| Ajout de [`scripts/Reporting/readme.md`](scripts/Reporting/readme.md) — documente [`Get-ComputerLastLogon.ps1`](scripts/Reporting/Get-ComputerLastLogon.ps1) avec un tableau des paramètres, une référence des colonnes du CSV et des exemples d'utilisation |
| Mise à jour de [`scripts/Reporting/Get-ComputerLastLogon.ps1`](scripts/Reporting/Get-ComputerLastLogon.ps1) — ajout des colonnes `PasswordLastSet` / `DaysSincePasswordSet` ; nouveau statut `Active (pwd recent)` pour les appareils marqués à tort comme inactifs en raison du délai de réplication de 14 jours de `LastLogonTimestamp` |

### 2026-03-26
| Modification |
|--------|
| Mise à jour de `scripts/Custom Scripts/device/audio/Disable-internalmic.ps1` — motifs de micro interne étendus : ajout des variantes linguistiques Conexant (EN/FR/NL), Synaptics EN, des pilotes Intel SST, IDT, Cirrus Logic, et des noms de réseaux de microphones multilingues (FR/DE/ES/PT/IT) |
| Mise à jour du readme — section Audio Management : ajout d'un tableau de déploiement NinjaOne (Run as SYSTEM, sans paramètres, champ personnalisé `AudioDeviceInventory`, codes de sortie) pour les trois scripts audio |
| Mise à jour du readme — [`Invoke-WindowsActivation.ps1`](scripts/Device/Invoke-WindowsActivation.ps1) : ajout d'un tableau de déploiement NinjaOne avec Run as Administrator, codes de sortie et exemples de paramètres par scénario ; ajout d'un avertissement indiquant que `-RemoveKey`/`-ReArm` nécessitent `-Force` |
| Ajout de `scripts/Testing Scripts/RDS/Watch-RDSLive.ps1` — moniteur RDS en temps réel : interroge toutes les 20 s les événements de session (20/21/22/23/24/25/40), les échecs d'ouverture de session RDP (4625), les verrouillages (4740) et les événements de licence ; signal de vie à chaque interrogation avec le nombre de sessions ; à exécuter directement sur chaque serveur RDS |

### 2026-03-25
| Modification |
|--------|
| Mise à jour de `scripts/Testing Scripts/Network/Test-FileIODiagnostics.ps1` — intégration d'un moniteur en temps réel : FileSystemWatcher, comparaison des autorisations NTFS avec une référence, téléchargement automatique de Handle.exe de Sysinternals, comparaison d'instantanés de processus, tickets Kerberos au moment de l'échec, événements d'audit Security (4625/4740/4656/4663/4670) ; s'arrête après 3 échecs |
| Ajout de `scripts/Testing Scripts/Network/Test-FileIODiagnostics.ps1` — test de charge des E/S de fichiers sur n'importe quel chemin (local ou UNC/lecteur mappé) ; classe les échecs en AUTH / NETWORK / TIMEOUT / DISK / PATH ; collecte automatiquement les tickets Kerberos, net use, la vérification du port SMB et le journal des événements Security dès le premier échec ; paramètres `-Iterations`, `-StopOnFirstError`, `-DelayMs` |
| Mise à jour de `scripts/Testing Scripts/RDS/Test-RDSDiagnostics.ps1` — ajout de l'analyse de l'Observateur d'événements pour RDWeb : TerminalServices-WebAccess/Admin+Operational, TerminalServices-Gateway/Admin+Operational, erreurs IIS/ASP.NET du journal Application ; déclenché par -IncludeEventLogs |
| Ajout de `scripts/Testing Scripts/RDS/Test-RDSDiagnostics.ps1` — diagnostic des échecs de connexion RDP/RDWeb : services, registre, NLA, limites de session, licences, pare-feu, certificat HTTPS, pool d'applications IIS, compte utilisateur (activé/verrouillé/expiré/groupe), journaux d'événements (4625/4740/4771/20/40) ; journal horodaté dans C:\Temp\ |
| Ajout de `scripts/Custom Scripts/device/Invoke-WindowsActivation.ps1` — active Windows, installe une clé de produit, configure le serveur/port KMS, supprime la clé, réarme la période de grâce (ReArm) ; essai à blanc sûr avec demandes de confirmation ; -Force pour les ignorer |
| Mise à jour de `scripts/Custom Scripts/device/Invoke-WindowsCleanup.ps1` — ajout d'une analyse dynamique des dossiers de journaux (section 13) : parcourt récursivement tout le lecteur C:\ (profondeur max. 7) à la recherche de dossiers nommés logs/log/logging/diagnostics ; ignore les dossiers système Windows et les artefacts de développement (node_modules, .git, venv) ; chemins fixes des journaux système Windows (CBS archivés en .cab, DISM, WU, Panther, IIS) ; nouveau paramètre `-SkipAppLogs` |
| Mise à jour de `scripts/Custom Scripts/device/Invoke-WindowsCleanup.ps1` — analyse de tous les profils utilisateur dans C:\Users\ pour les fichiers temporaires, WER, le cache des miniatures/shaders et les caches des navigateurs (Edge multiprofil, Chrome multiprofil, Firefox) ; le récapitulatif indique l'espace récupérable par catégorie |
| Correction de `scripts/Custom Scripts/device/Invoke-WindowsCleanup.ps1` — correction des avertissements de l'analyseur PS : `$profile` renommé en `$ffProfile`, suppression de l'affectation inutilisée `$dismResult` |
| Correction de `scripts/Custom Scripts/device/Invoke-WindowsCleanup.ps1` — remplacement de l'opérateur de fusion null `??` par `-as [int64]` pour la compatibilité avec PowerShell 5.1 |
| Ajout de `scripts/Custom Scripts/device/Invoke-WindowsCleanup.ps1` — nettoyage complet du disque Windows : fichiers temporaires, cache WU, Delivery Optimization, Prefetch, vidages mémoire, WER, cache des miniatures/shaders, Corbeille, caches des navigateurs, journaux d'événements, magasin de composants DISM ; essai à blanc par défaut, `-Apply` pour exécuter |
| Mise à jour de [`Get-SharePointStorageReport.ps1`](scripts/Reporting/Get-SharePointStorageReport.ps1) — approche en deux phases (énumération préalable de tous les sites/bibliothèques, puis récupération des données de stockage) ; correction d'une erreur de syntaxe `if` en ligne |
| Ajout de [`Test-AuthNetworkDiagnostics.ps1`](scripts/Network/Test-AuthNetworkDiagnostics.ps1) — diagnostics d'authentification et réseau : Observateur d'événements (4625/4771/4776/4740/5719), synchronisation de l'heure, cache Kerberos, DNS, TCP, partages UNC, analyse facultative des journaux |
| Ajout de `scripts/Custom Scripts/SAS/` — surveillance des erreurs des traitements par lots SAS avec intégration Zabbix, alertes par e-mail et analyse de l'Observateur d'événements |

### 2026-03-24
| Modification |
|--------|
| Ajout de [`scripts/Intune/iOS-Compliance-Updater/`](scripts/Intune/iOS-Compliance-Updater/readme.fr.md) — met à jour automatiquement la version iOS minimale dans la stratégie de conformité Intune via l'API Graph ; tâche planifiée hebdomadaire, prise en charge de l'essai à blanc, Setup.ps1 à exécution unique |
| Ajout de [`Get-SharePointStorageReport.ps1`](scripts/Reporting/Get-SharePointStorageReport.ps1) — rapport de stockage SharePoint à l'échelle du tenant avec l'historique des versions par fichier ; mode rapide (quota) et analyse récursive complète |
| Ajout de [`Test-OpenVpnDiagnostics.ps1`](scripts/Device/Test-OpenVpnDiagnostics.ps1) — diagnostics d'OpenVPN Connect : adaptateurs PnP, services, routes, DNS, journal des événements, logiciels VPN en conflit ; export txt vers `C:\Temp\` |

### 2026-03-23
| Modification |
|--------|
| Tous les exports CSV vont désormais dans `C:\Temp\` (Windows) ou `~/Downloads/` (macOS/Linux) |
| Ajout de [`Import-DnsRecords.ps1`](scripts/DNS/Import-DnsRecords.ps1) — résout le DNS public via dig (Google 8.8.8.8) et importe les enregistrements A/CNAME dans le DNS AD, essai à blanc par défaut |
| Ajout de [`New-M365User.ps1`](scripts/Entra/New-M365User.ps1) — crée un seul utilisateur M365 via Graph, mot de passe généré automatiquement, licence facultative |
| Ajout de [`Import-M365Users.ps1`](scripts/Entra/Import-M365Users.ps1) — création d'utilisateurs en masse à partir d'un CSV via Graph, essai à blanc par défaut, mots de passe dans le CSV de sortie |
| Ajout de [`Remove-M365Users.ps1`](scripts/Entra/Remove-M365Users.ps1) — suppression en masse d'utilisateurs Entra ID, essai à blanc par défaut, rapport CSV |
| Ajout de [`Get-ExternalForwards.ps1`](scripts/Exchange/Get-ExternalForwards.ps1) — audit des règles de transfert externe sur toutes les boîtes aux lettres, export CSV |
| Ajout de [`Get-MailboxSizes.ps1`](scripts/Exchange/Get-MailboxSizes.ps1) — rapport de taille des boîtes aux lettres + nombre d'éléments, trié par stockage, export CSV |
| Ajout de [`Test-DkimConfig.ps1`](scripts/Exchange/Test-DkimConfig.ps1) — validation de la configuration de signature DKIM + des CNAME/TXT DNS, avec une sortie des actions requises |
| Ajout de [`Test-M365GroupMembership.ps1`](scripts/Entra/Test-M365GroupMembership.ps1) — audit des propriétaires et membres des groupes M365 / Teams via Graph, export CSV |
| Ajout de [`Test-DistributionGroupPermissions.ps1`](scripts/Exchange/Test-DistributionGroupPermissions.ps1) — gestionnaires de DG, Send As, Send on Behalf, nombre de membres, export CSV |
| Ajout de [`Test-CalendarPermissions.ps1`](scripts/Exchange/Test-CalendarPermissions.ps1) — audit des autorisations de calendrier indépendant de la langue, export CSV |
| Ajout de [`Test-MailboxPermissions.ps1`](scripts/Exchange/Test-MailboxPermissions.ps1) — audit Full Access / Send As / Send on Behalf, export CSV |
| Sous-menu Exchange (`C`) et sous-menu Entra (`D`) : ajout des outils d'audit des autorisations |
| Déplacement de [`Test-Ports.ps1`](scripts/Network/Test-Ports.ps1) vers `scripts/Testing Scripts/Network/` |
| [`load.ps1`](load.ps1) — importe automatiquement les modules au démarrage ; détecte les modules manquants et propose de les installer |
| Ajout de [`load.ps1`](load.ps1) — configuration de première exécution (UPN + nom), enregistrée dans `load.config.ps1` (ignoré par git), lance le menu |
| Extension de [`menu.ps1`](menu.ps1) avec une section M365 (B–E) : sous-menus Exchange, Entra ID, MSP Admin |
| Ajout de [`menu.ps1`](menu.ps1) — lanceur interactif, une seule touche, chiffres + touches F, multiplateforme |
| Ajout de [`scripts/Network/Test-Ports.ps1`](scripts/Network/Test-Ports.ps1) — vérificateur de ports TCP, syntaxe de plage/liste, plusieurs cibles |
| Scripts de licences : traduits en anglais, rendus génériques, dossier d'export configurable |
| Réécriture de [`create_scheduled_task.ps1`](scripts/Reporting/Licensing/create_scheduled_task.ps1) — vérification des droits administrateur, détection automatique de Python, date de déclenchement dynamique |
| Ajout de [`scripts/Reporting/Licensing/`](scripts/Reporting/Licensing/readme.fr.md) — boîte à outils Pax8 + Ingram → rapport Excel |

### 2026-03-20
| Modification |
|--------|
| [`start.bat`](scripts/Deployment/start.bat) v2.8 — scission de « Do it all » : A = Intune, C = AD ; ajout du renommage de l'appareil (B) et de la jonction à AD (8) |
| Réécriture de [`Migrate-Calendar.ps1`](scripts/Exchange/Migrate-Calendar.ps1) v2.0 — en anglais, générique, paramètres obligatoires |
| Ajout de scripts de test SMTP ; ajout de `Test-SmtpRelay` à [`functies.ps1`](scripts/Startup/functies.ps1) |
| Réécriture de [`functies.ps1`](scripts/Startup/functies.ps1) — remplacement de MSOnline/AzureAD par Microsoft Graph, multiplateforme |
| Ajout de [`Set-CorporateWallpaper.ps1`](scripts/Intune/Desktop/Background/Desktop/Set-CorporateWallpaper.ps1), [`Set-Calendar-rights.ps1`](scripts/Exchange/Set-Calendar-rights.ps1), [`Restart-Time-Sync.ps1`](scripts/Device/Time%20sync/Restart-Time-Sync.ps1) |
| Suppression de toutes les références propres à l'entreprise ; traduction de tous les readmes en anglais |

### 2026-03-19
| Modification |
|--------|
| Mise en ligne initiale du dépôt |

---

## Mainteneur

**Sjoerd Kanon** — Orienté sécurité | Esprit d'équipe & orienté projet | Microsoft 365 & infrastructure
