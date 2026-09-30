[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../readme.fr.md) › [scripts](../readme.fr.md) › **Exchange**

# Scripts Exchange

Scripts de gestion des calendriers, des boîtes aux lettres et des groupes de distribution Exchange Online.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Migrate-Calendar.ps1`](Migrate-Calendar.ps1) ([docs](#migrate-calendarps1)) | Migrer un calendrier de groupe M365 partagé vers une Room Mailbox |
| [`Move-SharedCalendar.ps1`](Move-SharedCalendar.ps1) ([docs](#move-sharedcalendarps1)) | **Tout-en-un** : trouver un calendrier par mot-clé, le déplacer dans une boîte aux lettres de ressource, lister qui doit basculer — une seule connexion |
| [`Convert-SharedCalendarToResource.ps1`](Convert-SharedCalendarToResource.ps1) ([docs](#convert-sharedcalendartoresourceps1)) | Sortir un calendrier partagé de la boîte aux lettres d'un utilisateur pour le placer dans sa propre boîte aux lettres de salle/d'équipement — éléments, séries, pièces jointes et droits compris |
| [`Set-Calendar-rights.ps1`](Set-Calendar-rights.ps1) ([docs](#set-calendar-rightsps1)) | Accorder à un utilisateur des autorisations sur un dossier de calendrier |
| [`Set-Distributionlist-dynamic-static.ps1`](Set-Distributionlist-dynamic-static.ps1) ([docs](#set-distributionlist-dynamic-staticps1)) | Convertir les membres d'un groupe de distribution dynamique en un groupe ordinaire (statique) |
| [`Move-InboxToArchive.ps1`](Move-InboxToArchive.ps1) ([docs](#move-inboxtoarchiveps1)) | Déplacer tous les messages (ou ceux filtrés par date) de la boîte de réception d'une boîte aux lettres vers son dossier Archive |
| [`Test-CalendarPermissions.ps1`](Test-CalendarPermissions.ps1) ([docs](#test-calendarpermissionsps1)) | Auditer les autorisations sur les dossiers de calendrier |
| [`Get-CalendarMappings.ps1`](Get-CalendarMappings.ps1) ([docs](#get-calendarmappingsps1)) | Où chaque calendrier est réellement ajouté dans Outlook, à côté des droits qui le sous-tendent — ou trouver un calendrier par mot-clé (`-Search balie`) |
| [`Test-MailboxPermissions.ps1`](Test-MailboxPermissions.ps1) ([docs](#test-mailboxpermissionsps1)) | Auditer les délégations Full Access, Send As et Send on Behalf |
| [`Test-DistributionGroupPermissions.ps1`](Test-DistributionGroupPermissions.ps1) ([docs](#test-distributiongrouppermissionsps1)) | Auditer les gestionnaires, Send As, Send on Behalf et le nombre de membres des groupes de distribution |
| [`Test-DkimConfig.ps1`](Test-DkimConfig.ps1) ([docs](#test-dkimconfigps1)) | Valider la configuration de signature DKIM et les enregistrements DNS |
| [`Get-ExternalForwards.ps1`](Get-ExternalForwards.ps1) ([docs](#get-externalforwardsps1)) | Auditer les boîtes aux lettres avec un transfert externe |
| [`Get-MailboxSizes.ps1`](Get-MailboxSizes.ps1) ([docs](#get-mailboxsizesps1)) | Rapport sur la taille des boîtes aux lettres et le nombre d'éléments |
| [`Get-DistributionGroupMembers.ps1`](Get-DistributionGroupMembers.ps1) ([docs](#get-distributiongroupmembersps1)) | Qui figure sur quelle liste de distribution, sous forme de classeur Excel lisible par le client — ou uniquement les listes contenant une adresse (`-Member jan@contoso.com`) ou tout un domaine (`-Member @be.verizon.com`) |
| [`Get-MessageTraceReport.ps1`](Get-MessageTraceReport.ps1) ([docs](#get-messagetracereportps1)) | Tracer qui a reçu quoi, à quelle heure exacte, et vers où cela a été transféré |
| [`Remove-PhishingMessage.ps1`](Remove-PhishingMessage.ps1) ([docs](#remove-phishingmessageps1)) | Supprimer un message d'hameçonnage d'une, de plusieurs ou de toutes les boîtes aux lettres — essai à blanc par défaut |
| [`Restore-MailboxMessages.ps1`](Restore-MailboxMessages.ps1) ([docs](#restore-mailboxmessagesps1)) | Remettre dans leur dossier d'origine les messages déplacés ou supprimés un jour donné, et indiquer **qui** les a déplacés ou supprimés — aperçu par défaut |

---

### Migrate-Calendar.ps1

Migre un calendrier de groupe M365 partagé vers une Room Mailbox. Résout le problème des membres du groupe qui reçoivent une notification par e-mail pour chaque événement du calendrier — une Room Mailbox utilise le même mécanisme de réservation qu'une salle de réunion : pas de notifications, acceptation automatique, visible par tous.

**Fonctionnement**

1. Crée une App Registration Entra ID (ou en réutilise une existante)
2. Crée une Room Mailbox comme calendrier de destination
3. Configure AutoAccept et définit les autorisations Default sur Reviewer
4. Lit les événements du calendrier du groupe M365 via un accès délégué
5. Copie les événements vers la Room Mailbox via une authentification d'application
6. Supprime éventuellement le groupe M365 source

> Le script utilise un double flux d'authentification, car Microsoft exige un accès délégué pour lire les calendriers de groupe, mais des autorisations d'application pour écrire dans d'autres boîtes aux lettres.

**Paramètres**

| Paramètre | Obligatoire | Par défaut | Description |
|-----------|----------|---------|-------------|
| `-TenantId` | Oui | — | ID de tenant Entra ID |
| `-AdminUPN` | Oui | — | UPN de l'administrateur qui exécute le script (doit être membre du groupe source) |
| `-ClientId` | Non | — | ID client de l'App Registration. S'il est omis, une nouvelle inscription est créée automatiquement |
| `-ClientSecret` | Non | — | Secret client. S'il est omis, il est créé automatiquement |
| `-AppName` | Non | `HolidaysCalendarMigration` | Nom de l'App Registration |
| `-SourceGroupMail` | Non | — | Adresse e-mail du groupe M365 source |
| `-SourceGroupDisplayName` | Non | — | Nom d'affichage du groupe source (utilisé comme recherche de secours) |
| `-DestinationType` | Non | `Room` | Type de boîte aux lettres de destination : `Room` ou `Shared` |
| `-DestinationDisplayName` | Non | `Holidays Calendar` | Nom d'affichage de la boîte aux lettres de destination |
| `-DestinationAlias` | Non | `holidays-calendar` | Alias de la boîte aux lettres de destination |
| `-DestinationEmail` | Non | — | Adresse SMTP de la boîte aux lettres de destination |
| `-DaysBack` | Non | `365` | Nombre de jours en arrière pour la récupération des événements |
| `-DaysForward` | Non | `730` | Nombre de jours en avant pour la récupération des événements |
| `-DeleteSourceGroup` | Non | `$false` | Supprimer le groupe M365 après la migration |

**Exemples**

```powershell
# Première exécution — créer automatiquement l'App Registration
.\Migrate-Calendar.ps1 `
    -TenantId     "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -AdminUPN     "admin@contoso.com" `
    -SourceGroupMail "holidays@contoso.com"

# Exécutions suivantes — réutiliser l'App Registration existante
.\Migrate-Calendar.ps1 `
    -TenantId     "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -AdminUPN     "admin@contoso.com" `
    -ClientId     "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -ClientSecret "your-client-secret" `
    -SourceGroupMail "holidays@contoso.com"

# Essai à blanc — aucune modification
.\Migrate-Calendar.ps1 `
    -TenantId     "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -AdminUPN     "admin@contoso.com" `
    -SourceGroupMail "holidays@contoso.com" `
    -WhatIf
```

**Autorisations requises**

| Autorisation | Objectif |
|-----------|---------|
| Exchange Admin ou Global Admin | Créer la Room Mailbox |
| Global Admin | Créer l'App Registration + accorder le consentement administrateur |
| Membre du groupe M365 source | Lire le calendrier du groupe via un accès délégué |

**Modules requis**

```powershell
Install-Module ExchangeOnlineManagement     -Scope CurrentUser
Install-Module Microsoft.Graph.Applications  -Scope CurrentUser
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module Microsoft.Graph.Calendar      -Scope CurrentUser
Install-Module Microsoft.Graph.Groups        -Scope CurrentUser
Install-Module Microsoft.Graph.Users         -Scope CurrentUser
```

---

### Set-Calendar-rights.ps1

Accorde à un utilisateur des droits d'accès sur le dossier de calendrier d'un autre utilisateur dans Exchange Online. Prend en charge les boîtes aux lettres en néerlandais, français et anglais.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-User` | Oui | Nom d'utilisateur (sans domaine) qui reçoit les autorisations |
| `-TargetMailbox` | Oui | Nom d'utilisateur (sans domaine) de la boîte aux lettres cible |
| `-AccessRights` | Oui | Niveau d'accès (voir le tableau ci-dessous) |

**Niveaux d'accès**

| Valeur | Description |
|-------|-------------|
| `Owner` | Contrôle total, y compris la suppression et la gestion des dossiers |
| `PublishingEditor` | Lire, créer, modifier, supprimer et créer des sous-dossiers |
| `Editor` | Lire, créer, modifier et supprimer |
| `Author` | Lire et créer, modifier/supprimer ses propres éléments |
| `Reviewer` | Lecture seule |
| `AvailabilityOnly` | Disponibilité (libre/occupé) uniquement |
| `LimitedDetails` | Disponibilité avec détails limités |

**Exemples**

```powershell
# Accorder les droits Reviewer
.\Set-Calendar-rights.ps1 -User j.doe -TargetMailbox a.smith -AccessRights Reviewer

# Essai à blanc
.\Set-Calendar-rights.ps1 -User j.doe -TargetMailbox a.smith -AccessRights Editor -WhatIf
```

**Module requis**

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

---

### Set-Distributionlist-dynamic-static.ps1

Détermine les membres qui correspondent actuellement au filtre d'un groupe de distribution dynamique et les copie dans un groupe de distribution ordinaire (statique) — le groupe cible est créé s'il n'existe pas. Exporte aussi la liste des membres obtenue au format CSV. Nécessite une session Exchange Online active (`Connect-ExchangeOnline`).

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-DynamicGroupIdentity` | Oui | Groupe de distribution dynamique source (nom, alias, DN ou adresse SMTP) |
| `-TargetGroupIdentity` | Oui | Groupe de distribution ordinaire cible — créé s'il n'existe pas |
| `-TargetDisplayName` | Non | Nom d'affichage d'un nouveau groupe cible (par défaut : `<DynamicDisplayName> Static`) |
| `-TargetAlias` | Non | Alias d'un nouveau groupe cible (par défaut : `<DynamicAlias>-static`) |
| `-TargetPrimarySmtpAddress` | Non | Adresse SMTP d'un nouveau groupe cible (par défaut : l'adresse SMTP principale actuelle du groupe dynamique) |
| `-CopyManagersFromDynamic` | Non | Copier les propriétaires `ManagedBy` du groupe dynamique vers le groupe cible (par défaut : activé) |
| `-DisableCopyManagersFromDynamic` | Non | Désactiver la copie des propriétaires `ManagedBy` |
| `-MakeDynamicAddressTemporary` | Non | Attribuer d'abord au groupe dynamique une adresse SMTP principale temporaire, afin de libérer son adresse pour le groupe cible (par défaut : activé) |
| `-DisableMakeDynamicAddressTemporary` | Non | Désactiver le changement automatique d'adresse SMTP temporaire |
| `-ClearTargetMembers` | Non | Retirer les membres existants du groupe cible avant d'ajouter les membres dynamiques obtenus |
| `-ExportCsvPath` | Non | Chemin d'export CSV des membres obtenus (par défaut : `C:\Temp\DynamicGroupMembers_<timestamp>.csv` sous Windows, `~/Downloads` sous Linux/macOS) |
| `-SkipMemberAdd` | Non | Uniquement déterminer et exporter les membres, sans modifier le groupe cible |
| `-RenameDynamicGroupTo` | Non | Renommer le groupe de distribution dynamique source après le traitement |

**Exemples**

```powershell
# Résoudre le groupe dynamique et remplir le groupe ordinaire
.\Set-Distributionlist-dynamic-static.ps1 -DynamicGroupIdentity "All Sales" -TargetGroupIdentity "All Sales Static"

# Actualisation complète des membres du groupe cible
.\Set-Distributionlist-dynamic-static.ps1 -DynamicGroupIdentity sales@contoso.com -TargetGroupIdentity sales-static@contoso.com -ClearTargetMembers

# Essai à blanc
.\Set-Distributionlist-dynamic-static.ps1 -DynamicGroupIdentity "All Staff" -TargetGroupIdentity "All Staff Static" -WhatIf

# Convertir et renommer le groupe dynamique d'origine
.\Set-Distributionlist-dynamic-static.ps1 -DynamicGroupIdentity "All Sales" -TargetGroupIdentity "All Sales Static" -RenameDynamicGroupTo "All Sales (Legacy Dynamic)"

# Libérer l'adresse SMTP du groupe dynamique pour le nouveau groupe statique
.\Set-Distributionlist-dynamic-static.ps1 -DynamicGroupIdentity "All Sales" -TargetGroupIdentity "All Sales Static" -MakeDynamicAddressTemporary
```

**Remarques**
- Nécessite le module PowerShell Exchange Online et une session EXO active (`Connect-ExchangeOnline`)
- Les groupes de distribution dynamiques sont des objets Exchange ; ce script utilise des cmdlets Exchange, pas Graph

---

### Move-InboxToArchive.ps1

Déplace chaque message de la boîte de réception d'une boîte aux lettres vers son dossier Archive — le même dossier que cible le bouton « Archiver » d'Outlook. Vous pouvez éventuellement limiter la portée à une plage de dates (`-After` / `-Before`). Utilise le point de terminaison `$batch` de Microsoft Graph pour déplacer les messages par lots de 20, avec nouvelles tentatives/backoff en cas de limitation (429/503). Par défaut, le script fonctionne en mode aperçu sans risque — passez `-Apply` pour déplacer réellement les messages.

**Authentification (par défaut : automatique, sans Full Access)**

Par défaut, le script archive n'importe quelle boîte aux lettres du tenant sans nécessiter de Full Access sur celle-ci. Il se connecte de manière interactive (déléguée, `Application.ReadWrite.All` + `AppRoleAssignment.ReadWrite.All`), crée une App Registration temporaire de courte durée, s'accorde lui-même l'autorisation d'application `Mail.ReadWrite` (pas d'écran de consentement administrateur séparé — le rôle délégué s'en charge), l'utilise pour les opérations sur la boîte aux lettres, puis la supprime à la fin du script. C'est le même modèle d'application temporaire que dans `Get-SharePointStorageReport.ps1` / `Remove-SharePointFileVersionsByDate.ps1`. Nécessite Global Administrator ou Privileged Role Administrator pour cette configuration ponctuelle, ainsi que le module `Microsoft.Graph.Applications`.

- `-Delegated` ignore tout cela et utilise à la place une simple session déléguée `Mail.ReadWrite` — il faut alors Exchange Admin, et non des droits de création d'applications Entra. Pour une boîte aux lettres autre que celle de l'utilisateur connecté, le script se connecte à Exchange Online, accorde temporairement Full Access à ce compte, interroge `Get-MailboxPermission` jusqu'à ce que le droit soit réellement visible (jusqu'à ~3 minutes — les modifications d'autorisations dans Exchange Online ne se propagent pas instantanément), archive, puis retire à nouveau le droit (avec quelques nouvelles tentatives, car le retrait peut lui aussi tomber sur un contrôleur de domaine qui n'est pas encore à jour).
  > **Limitation connue :** `Get-MailboxPermission` reflète presque immédiatement l'état propre d'Exchange, mais le cache d'autorisation de Microsoft Graph pour l'accès délégué aux boîtes aux lettres peut avoir jusqu'à **~60 minutes** de retard — c'est une limitation côté Microsoft. Si la lecture de la boîte de réception renvoie toujours 403 après l'interrogation du Full Access, le script continue d'essayer (toutes les 60 s) jusqu'à une échéance fixée par **`-MaxWaitMinutes`** (65 par défaut, ce qui couvre le pire cas documenté par Microsoft) — le droit Full Access reste en place pendant toute l'attente, car le retirer et le réaccorder entre deux tentatives remettrait à zéro le délai de propagation. Augmentez `-MaxWaitMinutes` si 65 ne suffit pas, ou retirez `-Delegated` pour utiliser le mode app-only par défaut, qui n'a pas ce délai.
- `-ClientId` + `-ClientSecret`/`-CertificateThumbprint` réutilise votre propre App Registration existante au lieu d'en créer une temporaire — cette application doit déjà disposer de l'autorisation d'application `Mail.ReadWrite` (consentement administrateur accordé).

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|--------------|
| `-Mailbox` | Oui | UPN ou ID d'objet de la boîte aux lettres dont la boîte de réception doit être archivée |
| `-After` | Non | Archiver uniquement les messages reçus à cette date ou après |
| `-Before` | Non | Archiver uniquement les messages reçus avant cette date |
| `-TenantId` | Non | ID de tenant Entra ID (GUID) **ou** un domaine vérifié du tenant (par ex. `contoso.com`) — les deux fonctionnent. Facultatif si vous êtes déjà connecté ou s'il peut être déduit d'un contexte de tenant client GDAP ; obligatoire pour l'authentification app-only s'il ne peut pas être déduit |
| `-ClientId` | Non | ID client d'une App Registration existante pour l'authentification app-only — ignore l'application temporaire automatique. À utiliser avec `-TenantId` et `-ClientSecret` ou `-CertificateThumbprint` |
| `-ClientSecret` | Non | Secret client de l'App Registration indiquée dans `-ClientId` |
| `-CertificateThumbprint` | Non | Empreinte du certificat de l'App Registration indiquée dans `-ClientId` |
| `-Delegated` | Non | Ignorer la configuration app-only temporaire automatique et se connecter en mode délégué. Pour les autres boîtes aux lettres, accorde, interroge et retire automatiquement un Full Access temporaire via Exchange Online (nécessite Exchange Admin) |
| `-MaxWaitMinutes` | Non | Uniquement avec `-Delegated`. Durée pendant laquelle le script continue d'essayer en attendant que Graph honore le droit Full Access, avant d'abandonner et de le retirer. `65` par défaut |
| `-Apply` | Non | Déplacer réellement les messages. Sans ce commutateur, le script indique seulement combien de messages seraient archivés |

**Exemples**

```powershell
# Aperçu — configuration app-only automatique, affiche le nombre, ne modifie rien
.\Move-InboxToArchive.ps1 -Mailbox "user@contoso.com"

# Archiver tout le contenu de la boîte de réception
.\Move-InboxToArchive.ps1 -Mailbox "user@contoso.com" -Apply

# Uniquement les messages reçus avant 2025
.\Move-InboxToArchive.ps1 -Mailbox "user@contoso.com" -Before (Get-Date "2025-01-01") -Apply

# Uniquement les messages reçus en 2024
.\Move-InboxToArchive.ps1 -Mailbox "user@contoso.com" -After (Get-Date "2024-01-01") -Before (Get-Date "2025-01-01") -Apply

# Délégué — accorde, interroge et retire automatiquement un Full Access temporaire via Exchange Online
# au lieu de la configuration app-only Entra (nécessite Exchange Admin, pas Global Admin)
.\Move-InboxToArchive.ps1 -Mailbox "user@contoso.com" -Delegated -Apply

# Délégué, en acceptant d'attendre le pire cas complet de ~90 minutes selon Microsoft pour que Graph
# honore le droit Full Access, au lieu des 65 minutes par défaut
.\Move-InboxToArchive.ps1 -Mailbox "user@contoso.com" -Delegated -MaxWaitMinutes 90 -Apply

# Réutiliser une App Registration existante au lieu d'en créer une temporaire.
# -TenantId accepte le domaine du tenant au lieu de son GUID.
.\Move-InboxToArchive.ps1 -Mailbox "user@contoso.com" -TenantId "contoso.com" `
    -ClientId "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" -ClientSecret "your-client-secret" -Apply
```

**Remarques**
- Les lectures et déplacements dans la boîte aux lettres passent toujours par Microsoft Graph, et non par les cmdlets Exchange Online — nécessite `Microsoft.Graph.Authentication` (ainsi que `Microsoft.Graph.Applications` pour le mode par défaut avec application temporaire automatique, ou `ExchangeOnlineManagement` pour le Full Access temporaire de `-Delegated`)
- Affiche une progression horodatée pendant la pagination des messages de la boîte de réception, pendant le déplacement des lots (`[HH:mm:ss] N / total moved (...%)`) et pendant l'interrogation de la propagation du Full Access en mode `-Delegated`
- Compatible GDAP : dans une session GDAP (`$global:authMode -eq 'GDAP'`, définie via `Connect-Tenant` / `load.ps1`), `-TenantId` est déduit automatiquement du tenant client sélectionné (`$global:cid`) s'il est omis — même mécanisme de secours que dans `Get-SharePointStorageReport.ps1` / `Remove-SharePointFileVersionsByDate.ps1`. `$env:M365_CUSTOMER_TENANTID` / `$env:M365_AUTH_MODE` sont également pris en compte

**Modules requis**

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module Microsoft.Graph.Applications    -Scope CurrentUser
Install-Module ExchangeOnlineManagement        -Scope CurrentUser
```

---

### Move-SharedCalendar.ps1

**Tout-en-un :** de « où se trouve le calendrier Balie ? » à « il a sa propre boîte aux lettres de ressource » en une seule exécution, avec une seule connexion. Le script enchaîne [`Get-CalendarMappings.ps1`](Get-CalendarMappings.ps1) ([docs](#get-calendarmappingsps1)) et [`Convert-SharedCalendarToResource.ps1`](Convert-SharedCalendarToResource.ps1) ([docs](#convert-sharedcalendartoresourceps1)) — les deux doivent se trouver dans le même dossier.

| Étape | |
|------|--|
| 1. Trouver | `Get-CalendarMappings.ps1 -Search <keyword>` : où se trouve le calendrier, qui l'a dans Outlook, qui a des droits dessus |
| 2. Choisir | Le calendrier correspondant qui peut être déplacé. Plusieurs résultats : choisissez-en un dans une liste numérotée, ou affinez avec `-Owner`. Une exécution non interactive ne devine jamais — elle liste les candidats et s'arrête |
| 3. Déplacer | `Convert-SharedCalendarToResource.ps1` : aperçu, puis trois questions — continuer ? envoyer les invitations ? supprimer l'original ? `-Apply` saute le tour d'aperçu |
| 4. Informer | Qui avait l'ancien calendrier dans Outlook et qui n'avait que des droits : les personnes qui doivent basculer |

**Une seule connexion.** Une App Registration temporaire avec tout ce dont les deux scripts ont besoin (`Calendars.ReadWrite`, `User.Read.All`, `Group.Read.All`, `MailboxSettings.ReadWrite`) est créée une fois, transmise aux deux, puis supprimée à la fin — y compris en cas d'échec. La connexion à Exchange Online n'est établie qu'une fois, elle aussi. Une session Graph app-only existante ou `-ClientId` / `-ClientSecret` est utilisée à la place si elle est fournie.

Une boîte aux lettres dont le calendrier **principal** correspond (un compte `balie@` qui est lui-même le calendrier partagé) ne peut pas être déplacée ; le script le signale et indique l'alternative sur place, `Set-Mailbox -Type Room`.

Une exécution qui modifie quelque chose est journalisée dans `SharedCalendarMove_<timestamp>.log`, à côté du rapport de mappage et de la sauvegarde.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-Search` | Oui | Mot-clé : nom/adresse du propriétaire ou nom du calendrier. Alias `-Keyword` |
| `-Owner` | Non | Réduit plusieurs résultats à un seul propriétaire (partie du nom ou de l'adresse) |
| `-ResourceType` | Non | `Room` (par défaut) ou `Equipment` |
| `-ResourceName` / `-ResourceAddress` | Non | Nom et adresse de la nouvelle boîte aux lettres (par défaut : le nom du calendrier, sur le domaine du propriétaire) |
| `-SourceOwnerRights` | Non | Droits du propriétaire d'origine : `Owner` (par défaut) … `None` — utilisez `None` pour une boîte aux lettres archivée |
| `-SendSharingInvitation` | Non | Inviter les utilisateurs au nouveau calendrier (demandé si non fourni) |
| `-Apply` | Non | Continuer sans le tour d'aperçu |
| `-RemoveSourceCalendar` | Non | Supprimer l'original après une vérification sans erreur (demandé si non fourni) |
| `-Force` | Non | Ignorer la confirmation à saisir — requis pour supprimer sans surveillance |
| `-OutputPath` | Non | Dossier du rapport, de la sauvegarde et du journal (par défaut `C:\Temp`) |
| `-TenantId` / `-ClientId` / `-ClientSecret` | Non | Tenant, ou votre propre App Registration au lieu d'une temporaire |

**Exemples**

```powershell
# Trouver, afficher l'aperçu, répondre aux questions
.\Move-SharedCalendar.ps1 -Search balie

# "Balie planning" dans une boîte aux lettres archivée : boîte Equipment, utilisateurs invités, original conservé pour l'instant
.\Move-SharedCalendar.ps1 -Search "balie planning" -ResourceType Equipment -SourceOwnerRights None `
    -SendSharingInvitation -Apply

# Une fois que tout le monde a basculé : la même commande supprime l'original
.\Move-SharedCalendar.ps1 -Search "balie planning" -ResourceType Equipment -SourceOwnerRights None `
    -Apply -RemoveSourceCalendar
```

La deuxième exécution ignore chaque élément déjà copié, copie ce qui a été ajouté entre-temps, et ne supprime l'original qu'ensuite.

---

### Convert-SharedCalendarToResource.ps1

Sort un calendrier partagé de la boîte aux lettres d'un utilisateur pour le placer dans **sa propre boîte aux lettres de ressource** (Room ou Equipment), avec chaque élément et chaque autorisation, puis — sur demande — supprime l'original. Conçu pour le calendrier « Balie » typique : un calendrier supplémentaire dans la boîte aux lettres d'une personne, utilisé par tout l'accueil, et qui part avec cette personne.

**Aperçu par défaut.** Sans `-Apply`, le script se contente de lire et de rendre compte : combien d'éléments, de séries et d'exceptions, quelles autorisations il reprendrait et quelle boîte aux lettres il créerait. L'original n'est supprimé qu'avec `-RemoveSourceCalendar`, uniquement après que chaque élément a une copie vérifiée, et uniquement après que vous avez saisi le nom du calendrier (à ignorer avec `-Force`).

**Ce qui se passe avec `-Apply`**

| Étape | |
|------|--|
| Sauvegarde | Chaque élément (corps compris), chaque occurrence de série et chaque autorisation dans `calendar-backup.json`, avant toute création |
| Boîte aux lettres | Boîte Room (par défaut) ou Equipment avec la langue et le fuseau horaire de la boîte source. Traitement du calendrier adapté à un calendrier partagé : acceptation automatique, éléments qui se chevauchent autorisés, rien n'est réécrit dans un élément, fenêtre de réservation de 1080 jours (le maximum du service) |
| Droits | Chaque autorisation avec ses droits d'accès Exchange **exacts** (droits personnalisés compris), `Default` et `Anonymous` tels qu'ils étaient, le propriétaire d'origine selon `-SourceOwnerRights` (par défaut `Owner`). `-SendSharingInvitation` envoie l'e-mail habituel « a partagé un calendrier avec vous » |
| Catégories | Les catégories utilisées sont créées dans la nouvelle boîte aux lettres avec leur couleur |
| Éléments | Chaque élément est copié — voir ci-dessous |
| Vérification | Chaque élément source doit avoir une copie complète |
| Suppression | Uniquement avec `-RemoveSourceCalendar` et une vérification sans erreur |

**Comment les éléments sont copiés**

- **Les séries restent des séries.** Les occurrences déplacées ou modifiées sont appliquées à la copie et les occurrences annulées y sont annulées, en comparant les deux séries occurrence par occurrence. Pour une série sans date de fin, cela se fait jusqu'à `-SeriesHorizonDays` (1095) jours à l'avance. Si les deux séries ne concordent pas, le script laisse cette série telle quelle et le signale plutôt que d'annuler les mauvaises occurrences
- **Les heures conservent leur fuseau horaire.** Graph renvoie de l'UTC ; chaque élément est réécrit dans le fuseau dans lequel il a été créé, de sorte qu'un élément hebdomadaire à 9:00 reste à 9:00 après le passage à l'heure d'été ou d'hiver
- **Personne n'est invité.** Copier une réunion avec ses participants enverrait à chaque participant une nouvelle invitation depuis la boîte aux lettres de ressource. L'organisateur et les participants sont plutôt listés en bas du corps ; la boîte aux lettres de ressource est l'organisatrice de chaque copie
- **Les pièces jointes** jusqu'à 3 Mo sont copiées, images incorporées comprises. Les fichiers plus volumineux et les éléments Outlook joints sont enregistrés dans le dossier de sauvegarde et listés à la fin
- **Réexécutable.** Chaque copie porte l'ID de son élément source dans une propriété masquée. Une exécution qui s'arrête à mi-chemin se poursuit en relançant la même commande : les copies complètes sont ignorées, une série à moitié copiée est supprimée puis recopiée

**Non repris, mais signalé :** les indicateurs de délégué (une boîte aux lettres de ressource n'a pas de délégués), les personnes extérieures au tenant (repartager manuellement), les autorisations de comptes supprimés, les pièces jointes des exceptions individuelles d'une série. Un calendrier publié (`Anonymous`) reçoit un nouveau lien.

> **Le calendrier principal ne peut pas être converti de cette manière.** Si toute la boîte aux lettres *est* le calendrier partagé (un compte utilisateur `balie@`), convertissez-la plutôt sur place — cela conserve tout : `Set-Mailbox balie@contoso.com -Type Room`. Le script refuse le calendrier principal et renvoie vers cette solution.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-Mailbox` | Oui | Boîte aux lettres utilisateur qui contient le calendrier |
| `-Calendar` | Oui | Nom du calendrier tel qu'affiché dans Outlook (par ex. `Balie`). Doit appartenir à cet utilisateur et ne pas être son calendrier principal |
| `-ResourceName` | Non | Nom d'affichage de la nouvelle boîte aux lettres (par défaut : le nom du calendrier) |
| `-ResourceAddress` | Non | Adresse SMTP (par défaut : le nom comme alias sur le domaine de l'utilisateur). Une boîte aux lettres de salle/d'équipement existante à cette adresse est réutilisée |
| `-ResourceType` | Non | `Room` (par défaut) ou `Equipment` |
| `-SourceOwnerRights` | Non | Droits du propriétaire d'origine : `Owner` (par défaut), `PublishingEditor`, `Editor`, `Reviewer`, `None` |
| `-SendSharingInvitation` | Non | Envoyer une invitation de partage aux utilisateurs (possible uniquement pour Reviewer, Editor, LimitedDetails, AvailabilityOnly) |
| `-Apply` | Non | Créer, accorder les droits et copier réellement. Sans ce commutateur : aperçu |
| `-RemoveSourceCalendar` | Non | Supprimer l'original après une vérification sans erreur. Nécessite `-Apply` |
| `-Force` | Non | Ignorer la confirmation à saisir avant la suppression. Une session non interactive (planificateur, RMM) ne peut pas la saisir ; l'original n'y est donc supprimé qu'avec `-Force` |
| `-PassThru` | Non | Renvoyer un objet résultat (`ResourceAddress`, `Items`, `Verified`, `SourceRemoved`, `BackupPath`) à un script appelant |
| `-SeriesHorizonDays` | Non | Jusqu'où à l'avance les exceptions des séries sans fin sont comparées (1095 par défaut) |
| `-BackupPath` | Non | Dossier de sauvegarde (par défaut `C:\Temp\CalendarConvert_<calendar>_<timestamp>`) |
| `-TenantId` | Non | ID de tenant ou domaine (par défaut : le tenant de la session Exchange) |
| `-ClientId` / `-ClientSecret` / `-CertificateThumbprint` | Non | Votre propre App Registration pour un accès Graph app-only |

**Exemples**

```powershell
# 1. Aperçu
.\Convert-SharedCalendarToResource.ps1 -Mailbox jan@contoso.com -Calendar Balie

# 2. Créer la boîte aux lettres de salle, tout copier, inviter les utilisateurs - l'original reste
.\Convert-SharedCalendarToResource.ps1 -Mailbox jan@contoso.com -Calendar Balie -Apply -SendSharingInvitation

# 3. Une fois que les utilisateurs ont basculé : trouver qui a encore l'ancien calendrier, puis le supprimer
.\Get-CalendarMappings.ps1 -Search Balie
.\Convert-SharedCalendarToResource.ps1 -Mailbox jan@contoso.com -Calendar Balie -Apply -RemoveSourceCalendar
```

L'étape 3 relance d'abord la copie : tout ce qui a déjà été copié est ignoré, ce qui a été ajouté à l'original entre-temps est copié, et ce n'est qu'ensuite que l'original est supprimé.

**Accès**

| | |
|--|--|
| Exchange Online | Exchange Administrator (`New-Mailbox`, autorisations sur les dossiers). Une session existante est réutilisée |
| Graph | Autorisation d'application `Calendars.ReadWrite`, plus `MailboxSettings.ReadWrite` pour les couleurs des catégories (facultatif). Les trois mêmes voies que [`Remove-PhishingMessage.ps1`](Remove-PhishingMessage.ps1) ([docs](#remove-phishingmessageps1)) : session app-only existante, votre propre App Registration, ou une App Registration temporaire supprimée à la fin de l'exécution. Du REST pur, donc pas de conflit MSAL entre Exchange et Graph |

**Remarques**

- La copie est un instantané. Lancez-la quand le calendrier est calme et demandez aux utilisateurs de basculer juste après ; l'étape 3 ci-dessus reprend ce qui a été ajouté entre-temps
- Les utilisateurs qui avaient le calendrier d'origine dans leur liste conservent une entrée qui cesse de fonctionner une fois l'original supprimé — identifiez-les d'abord avec `Get-CalendarMappings.ps1 -Search`
- Diffère de `Migrate-Calendar.ps1` (calendrier de groupe → salle) : ce script-là copie les éléments sans corps, sans séries ni fuseau horaire. Celui-ci se veut un déplacement fidèle

---

## Scripts d'audit

Se connectent automatiquement à Exchange Online si aucune session n'est active ; réutilisent une session existante si vous êtes déjà connecté.

---

### Test-CalendarPermissions.ps1

Récupère les autorisations des dossiers de calendrier pour une ou toutes les boîtes aux lettres. Utilise `FolderType` pour localiser le dossier de calendrier indépendamment de la langue de la boîte aux lettres (NL/FR/EN). Exporte les résultats au format CSV.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-Mailbox` | Non | UPN d'une seule boîte aux lettres. S'il est omis, toutes les boîtes aux lettres utilisateur et partagées sont vérifiées |
| `-OutputPath` | Non | Chemin du rapport CSV (par défaut : `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Non | ID de tenant Entra ID ou domaine |

**Exemples**

```powershell
# Auditer toutes les boîtes aux lettres
.\Test-CalendarPermissions.ps1

# Une seule boîte aux lettres
.\Test-CalendarPermissions.ps1 -Mailbox "user@contoso.com"

# Chemin de sortie personnalisé
.\Test-CalendarPermissions.ps1 -OutputPath "C:\Reports\calendar.csv"
```

---

### Get-CalendarMappings.ps1

Montre **où chaque calendrier est ajouté** : les calendriers qui figurent réellement dans la liste de calendriers d'un utilisateur dans Outlook, à côté des droits qui les sous-tendent. `Test-CalendarPermissions.ps1` répond à « qui *peut* ouvrir ce calendrier » ; ce script répond à « où *se trouve*-t-il » et signale les endroits où les deux divergent. Lecture seule.

Vous cherchez un seul calendrier ? `-Search balie` le trouve par mot-clé — voir *Recherche par mot-clé* ci-dessous.

Pour chaque boîte aux lettres, il lit via Microsoft Graph :

- la **liste de calendriers** (`/users/{id}/calendars`). Chaque calendrier de cette liste appartenant à quelqu'un d'autre est un mappage : un collègue, une boîte aux lettres partagée, une salle, un groupe Microsoft 365 ou une personne extérieure à l'organisation
- les **autorisations sur son propre calendrier principal** (`/users/{id}/calendar/calendarPermissions`)

et regroupe les deux en une ligne par couple propriétaire du calendrier + utilisateur :

| Statut | Signification |
|--------|---------|
| `Source` | Uniquement avec `-Search` : le calendrier correspondant se trouve dans cette boîte aux lettres (son calendrier principal ou un calendrier secondaire) |
| `Mapped` | Dans la liste de calendriers de l'utilisateur, et l'utilisateur a un droit explicite |
| `MappedWithoutRight` | Dans la liste, mais aucun droit explicite sur le calendrier principal du propriétaire. L'accès provient alors de la valeur par défaut à l'échelle de l'organisation, d'un groupe, d'un calendrier secondaire du propriétaire — ou le droit a été retiré et l'entrée est restée (l'utilisateur obtient une erreur en l'ouvrant) |
| `MappedGroupCalendar` | Un calendrier de groupe Microsoft 365 — l'accès suit l'appartenance au groupe |
| `MappedOwnerMissing` | Le propriétaire n'existe plus dans le tenant — une entrée obsolète dans la liste de l'utilisateur |
| `MappedExternal` | Le propriétaire est extérieur au tenant |
| `NotMapped` | Droit explicite, mais le calendrier ne figure pas dans la liste de l'utilisateur — candidat au nettoyage |
| `NotChecked` | Droit explicite, mais la liste de calendriers de l'utilisateur n'a pas pu être lue |
| `GrantedToGroup` | Un droit accordé à un groupe ; les membres ne sont pas développés |
| `GrantedToMissing` | Un droit pour une adresse ou un compte qui n'existe plus — candidat au nettoyage |
| `SharedExternally` | Un droit pour une adresse extérieure au tenant |
| `OrgWideDefault` | *My Organization* obtient plus que la disponibilité — tout utilisateur interne peut ouvrir le calendrier |

> **Pourquoi Graph et non Exchange Online PowerShell :** les cmdlets Exchange voient les dossiers et leurs autorisations, pas les entrées qu'un utilisateur a ajoutées à sa propre liste de calendriers. Celles-ci ne sont lisibles que via Graph.

**Non visible dans ce rapport**

- **Full Access avec AutoMapping** ajoute toute une boîte aux lettres à Outlook, calendrier compris. C'est une autorisation de boîte aux lettres, pas une entrée de calendrier — voir [`Test-MailboxPermissions.ps1`](Test-MailboxPermissions.ps1) ([docs](#test-mailboxpermissionsps1))
- Un calendrier ouvert dans Outlook classique avec les *shared calendar improvements* désactivées peut n'exister que dans ce profil Outlook et pas dans la liste renvoyée par Graph
- Sans `-Search`, les droits sont comparés au calendrier **principal** du propriétaire. Un calendrier secondaire partagé par le propriétaire apparaît comme `MappedWithoutRight` ; `-Search` lit les droits propres d'un calendrier secondaire correspondant

**Recherche par mot-clé**

`-Search balie` (alias `-Keyword`) répond à « où se trouve le calendrier Balie ? ». Le mot-clé est comparé sans tenir compte de la casse, n'importe où dans le texte, avec :

- le **nom et chaque adresse** du propriétaire — la boîte aux lettres partagée `balie@`, une salle, un groupe nommé *Balie-team*
- le **nom propre** du calendrier — un calendrier secondaire *Balie* dans la boîte aux lettres de quelqu'un

Les caractères génériques (`*`, `?`) sont utilisés tels quels. Pour chaque résultat, le rapport indique où se trouve le calendrier (`Source`), chaque boîte aux lettres qui l'a dans sa liste de calendriers, et toutes les personnes ayant un droit explicite dessus — pour un calendrier secondaire, ses propres droits, qu'une exécution normale ne lit pas. La colonne `Calendar` indique de quel calendrier du propriétaire traite une ligne.

```
Owner      Calendar       User Status              Rights MappedAs
-----      --------       ---- ------              ------ --------
Anna       Balie Planning      Source                     Balie Planning
Anna       Balie Planning Lisa Mapped              read   Balie Planning
Anna       Balie Planning Jan  NotMapped           write
Balie      Main                Source                     Agenda
Balie      Main           Kees Mapped              read   Balie
Balie      Main           Piet Mapped              write  Balie
Balie      Main           Lisa NotMapped           read
Balie-team                Kees MappedGroupCalendar        Balie-team
```

Chaque liste de calendriers est quand même lue — un mappage peut se trouver dans n'importe quelle boîte aux lettres — une recherche prend donc à peu près autant de temps qu'une analyse complète pour les listes, mais ne lit que les autorisations des calendriers correspondants.

> Une entrée de liste de calendriers ne contient aucun lien vers le calendrier dont elle provient. Un calendrier secondaire partagé est donc reconnu à son nom correspondant au mot-clé. Si un utilisateur l'a sous un autre nom, il apparaît comme `NotMapped` avec une note indiquant l'entrée qui est probablement la bonne (*"Has a calendar of this owner as 'Planning Anna' - probably this one"*).

**Portée**

Sans `-Mailbox`, chaque boîte aux lettres du tenant est analysée — le seul moyen de trouver les mappages qui reposent sur la valeur par défaut à l'échelle de l'organisation ou sur un groupe. Avec `-Mailbox`, le rapport se limite aux lignes où l'une de ces boîtes aux lettres est le **propriétaire ou l'utilisateur** : leurs propres listes de calendriers sont lues, ainsi que les listes de toutes les personnes ayant un droit explicite sur leur calendrier. Pour une réponse complète à « où le calendrier de X est-il ajouté », utilisez plutôt `-Search`. `-Mailbox` et `-Search` ne peuvent pas être combinés.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-Search` | Non | Mot-clé pour trouver un calendrier (nom/adresse du propriétaire ou nom du calendrier). Alias `-Keyword`. Ne peut pas être combiné avec `-Mailbox` |
| `-Mailbox` | Non | Une ou plusieurs adresses de boîtes aux lettres. Limite le rapport aux lignes où elles sont propriétaire ou utilisateur. S'il est omis, chaque boîte aux lettres est analysée |
| `-OutputPath` | Non | Chemin du rapport CSV (par défaut : `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Non | ID de tenant ou domaine. Facultatif pour la voie de l'application temporaire — la connexion détermine alors le tenant, qui est affiché |
| `-ClientId` | Non | Votre propre App Registration pour un accès Graph app-only |
| `-ClientSecret` | Non | Secret client pour `-ClientId` (REST pur, pas de SDK Graph) |
| `-CertificateThumbprint` | Non | Empreinte du certificat pour `-ClientId` (via `Connect-MgGraph`) |

**Exemples**

```powershell
# Où se trouve le calendrier Balie, et qui l'a ajouté ?
.\Get-CalendarMappings.ps1 -Search balie

# Où chaque calendrier du tenant est-il ajouté ?
.\Get-CalendarMappings.ps1 -TenantId contoso.com

# Où le calendrier de Jan est-il ajouté, et quels calendriers Jan a-t-il ajoutés ?
.\Get-CalendarMappings.ps1 -Mailbox jan@contoso.com

# Votre propre App Registration
.\Get-CalendarMappings.ps1 -TenantId contoso.com -ClientId <appId> -ClientSecret <secret>
```

**Accès Graph**

Nécessite les autorisations d'application `Calendars.Read` et `User.Read.All`, plus `Group.Read.All` pour distinguer un calendrier de groupe d'une boîte aux lettres supprimée (sans elle, les calendriers de groupe apparaissent comme `MappedOwnerMissing` avec une note le précisant). Obtenues des trois mêmes manières que pour [`Remove-PhishingMessage.ps1`](Remove-PhishingMessage.ps1) ([docs](#remove-phishingmessageps1)) :

| # | Voie | Ce qu'il faut |
|---|-------|---------------|
| 1 | Une session Graph app-only que vous avez déjà établie | Rien — elle est utilisée telle quelle |
| 2 | `-ClientId` + `-ClientSecret` ou `-CertificateThumbprint` | Votre propre application avec les autorisations ci-dessus, consentement administrateur accordé. Une autorisation plus large (`Calendars.ReadWrite`, `Directory.Read.All`) est également acceptée |
| 3 | **Automatique** — connexion par code d'appareil, une App Registration de courte durée qui s'accorde elle-même les trois autorisations de lecture, supprimée à la fin de l'exécution (y compris en cas d'échec) | Global Administrator ou Privileged Role Administrator pour cette connexion. Aucun module supplémentaire |

Aucune connexion à Exchange Online n'est établie, le conflit MSAL entre Exchange et Graph décrit sous `Remove-PhishingMessage.ps1` ne s'applique donc pas. Compatible GDAP comme les autres scripts Graph : dans une session GDAP, `-TenantId` est déduit du tenant client sélectionné.

**Remarques**

- Les requêtes passent par Graph `$batch`, 20 boîtes aux lettres par appel. Les éléments limités sont réessayés après le `Retry-After` demandé par le service
- Les utilisateurs ayant une adresse mais pas de boîte aux lettres Exchange Online (404) sont ignorés et comptés. Une boîte aux lettres illisible est listée séparément, jamais signalée comme « rien d'ajouté ». Un 403 à cet endroit signifie généralement qu'une Application Access Policy ou RBAC for Applications limite l'application
- Les boîtes aux lettres partagées et les salles sont des comptes désactivés dans Entra ID ; elles sont donc incluses — pas de filtre sur `accountEnabled`

---

### Test-MailboxPermissions.ps1

Audite les trois types de délégation pour une ou toutes les boîtes aux lettres :

- **Full Access** — utilisateurs qui peuvent ouvrir la boîte aux lettres
- **Send As** — utilisateurs qui peuvent envoyer en tant qu'identité de la boîte aux lettres
- **Send on Behalf** — utilisateurs listés dans `GrantSendOnBehalfTo`

Les entrées héritées et les entrées SELF sont automatiquement filtrées.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-Mailbox` | Non | UPN d'une seule boîte aux lettres. S'il est omis, toutes les boîtes aux lettres utilisateur et partagées sont vérifiées |
| `-OutputPath` | Non | Chemin du rapport CSV (par défaut : `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Non | ID de tenant Entra ID ou domaine |

**Exemples**

```powershell
# Auditer toutes les boîtes aux lettres
.\Test-MailboxPermissions.ps1

# Uniquement une boîte aux lettres partagée
.\Test-MailboxPermissions.ps1 -Mailbox "shared@contoso.com"
```

---

### Test-DistributionGroupPermissions.ps1

Audite les groupes de distribution et les groupes de sécurité à extension messagerie :

- **Paramètres** — nombre de membres, restrictions d'adhésion/de départ, stratégie pour les expéditeurs externes
- **ManagedBy** — propriétaires/gestionnaires du groupe
- **Send As** — qui peut envoyer en tant que groupe
- **Send on Behalf** — délégués dans `GrantSendOnBehalfTo`
- **Membres** (facultatif, utilisez `-IncludeMembers`)

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-Group` | Non | Nom, alias ou adresse e-mail d'un seul groupe. S'il est omis, tous les groupes de distribution sont audités |
| `-IncludeMembers` | Non | Lister aussi les membres individuels du groupe dans le rapport |
| `-OutputPath` | Non | Chemin du rapport CSV (par défaut : `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Non | ID de tenant Entra ID ou domaine |

**Exemples**

```powershell
# Auditer tous les groupes de distribution
.\Test-DistributionGroupPermissions.ps1

# Un seul groupe avec la liste des membres
.\Test-DistributionGroupPermissions.ps1 -Group "helpdesk@contoso.com" -IncludeMembers
```

**Module requis**
```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

---

### Test-DkimConfig.ps1

Valide la configuration de signature DKIM pour un ou tous les domaines acceptés :

- Vérifie si la signature DKIM est activée
- Résout les enregistrements CNAME `selector1/2._domainkey.<domain>` et les compare à la configuration Exchange
- Résout les enregistrements TXT de clé publique de Microsoft et vérifie que la clé correspond
- Liste les actions requises pour chaque problème détecté

> Les résolutions DNS utilisent `Resolve-DnsName` (Windows uniquement). Sous macOS/Linux, la configuration Exchange est affichée sans validation DNS.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-Domain` | Non | Domaine à valider. S'il est omis, tous les domaines ayant une configuration de signature sont vérifiés |
| `-ShowAll` | Non | Afficher l'objet complet de configuration de signature au lieu de la vue résumée |
| `-TenantId` | Non | ID de tenant Entra ID ou domaine |

**Exemples**

```powershell
# Vérifier tous les domaines
.\Test-DkimConfig.ps1

# Un seul domaine
.\Test-DkimConfig.ps1 -Domain "contoso.com"
```

---

### Get-ExternalForwards.ps1

Audite toutes les boîtes aux lettres à la recherche de règles de transfert pointant vers des domaines externes (hors tenant). Le transfert externe est un risque courant en matière de sécurité et de conformité et doit être examiné régulièrement. Exporte au format CSV.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-Mailbox` | Non | UPN d'une seule boîte aux lettres. S'il est omis, toutes les boîtes aux lettres sont vérifiées |
| `-OutputPath` | Non | Chemin du rapport CSV (par défaut : `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Non | ID de tenant Entra ID ou domaine |

**Exemples**

```powershell
# Vérifier toutes les boîtes aux lettres
.\Get-ExternalForwards.ps1

# Une seule boîte aux lettres
.\Get-ExternalForwards.ps1 -Mailbox "user@contoso.com"
```

---

### Get-MailboxSizes.ps1

Rend compte de la taille des boîtes aux lettres (Mo/Go), du nombre d'éléments et de l'état des quotas. Trié par taille décroissante. Exporte au format CSV.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-Mailbox` | Non | UPN d'une seule boîte aux lettres. S'il est omis, toutes les boîtes aux lettres utilisateur et partagées sont incluses dans le rapport |
| `-OutputPath` | Non | Chemin du rapport CSV (par défaut : `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Non | ID de tenant Entra ID ou domaine |

**Exemples**

```powershell
# Rapport sur toutes les boîtes aux lettres
.\Get-MailboxSizes.ps1

# Une seule boîte aux lettres
.\Get-MailboxSizes.ps1 -Mailbox "user@contoso.com"
```

---

### Get-DistributionGroupMembers.ps1

Exporte chaque liste de distribution avec ses membres dans un seul classeur Excel, destiné à être envoyé tel quel au client.

Le classeur comporte deux feuilles, toutes deux des tableaux filtrables avec une ligne d'en-tête figée :

| Feuille | Une ligne par | Colonnes |
|-------|-------------|---------|
| `Overzicht` | liste | Lijst, E-mailadres, Type, Aantal leden, Eigenaar(s), Alias, Verborgen in adresboek, Alleen interne afzenders, Aangemaakt op |
| `Leden` | membre | Lijst, E-mailadres lijst, Type lijst, Lid, E-mailadres lid, Extern adres, Type lid |

Avec `-Recurse`, les feuilles gagnent `Aantal personen` et `Via groep` ; avec `-Member`, elles gagnent `Treffers` et `Treffer op`.

`Extern adres` est rempli pour les contacts de messagerie et les utilisateurs de messagerie. Leur adresse SMTP principale est un espace réservé interne — l'adresse qui reçoit réellement le courrier est l'adresse externe, et pour un rapport sur les membres externes, c'est la colonne qui compte.

Les en-têtes des feuilles et les types de destinataires sont en néerlandais — `MailUniversalSecurityGroup` ne dit rien à la personne qui lit le rapport, `Beveiligingsgroep (mail-enabled)` si. Le script lui-même reste en anglais, comme le reste du dépôt.

**Listes imbriquées — à lire avant de se fier au résultat**

Exchange ne renvoie jamais que les membres **directs**. Une liste qui contient une autre liste signale cette liste comme *un seul membre* et jamais les personnes qu'elle contient. Par défaut, donc :

- une personne qui ne reçoit le courrier que par un groupe imbriqué n'apparaît pas dans le rapport ;
- `-Member` ne signale **aucun résultat** sur une liste qui distribue pourtant bien à cette personne.

Ce deuxième point est la moitié dangereuse : c'est une mauvaise réponse qui a l'air sûre d'elle. `-Recurse` développe les groupes imbriqués, de sorte que le rapport liste les personnes qui reçoivent réellement le courrier :

```powershell
# "Est-ce que quelque chose atteint encore ce domaine ?" - la forme à utiliser pour cette question
.\Get-DistributionGroupMembers.ps1 -Member "@be.verizon.com" -Recurse
```

| | Sans `-Recurse` | Avec `-Recurse` |
|---|---|---|
| Liste imbriquée | une ligne de membre, personne derrière | une ligne de membre **plus** les personnes qu'elle contient |
| Colonne `Via groep` | — | indique le groupe par lequel une personne est arrivée, vide pour un membre direct |
| `Aantal leden` | membres directs (ce qu'affichent Exchange et l'EAC) | inchangé |
| `Aantal personen` | — | les destinataires réels que la liste atteint |

Cela coûte une requête supplémentaire par groupe imbriqué. Un groupe déjà développé ne l'est pas une seconde fois, ce qui empêche aussi un cycle d'appartenance (A contient B, B contient A) de boucler indéfiniment ; une imbrication de plus de 20 niveaux est signalée et laissée telle quelle. Une personne joignable par plusieurs chemins obtient une seule ligne avec les chemins regroupés, et non une ligne par chemin.

Le groupe imbriqué lui-même reste dans le rapport sous forme de ligne distincte, afin que la structure reste visible.

**Filtrer sur une adresse ou un domaine**

`-Member` accepte une adresse, un domaine, ou un domaine et tout ce qui se trouve en dessous :

```powershell
-Member "jan@contoso.com"     # sur quelles listes Jan figure-t-il ?
-Member "@be.verizon.com"     # membres exactement sur ce domaine
-Member "*.verizon.com"       # verizon.com ET chacun de ses sous-domaines
```

| Écrit sous la forme | Correspond à | Ne correspond pas à |
|---|---|---|
| `@be.verizon.com`, `be.verizon.com`, `*@be.verizon.com` | `jan@be.verizon.com` | `jan@verizon.com`, `jan@us.verizon.com`, `jan@notbe.verizon.com` |
| `*.verizon.com`, `.verizon.com`, `*@*.verizon.com` | `jan@verizon.com`, `jan@be.verizon.com`, `jan@us.verizon.com` | `jan@notverizon.com`, `jan@verizon.com.evil.test` |

Sans le `*.` initial, la correspondance porte sur ce **seul** domaine — `@be.verizon.com` n'atteint volontairement pas un domaine frère comme `@us.verizon.com`. Avec lui, le domaine racine et chaque sous-domaine sont inclus. L'exécution affiche lequel des deux elle applique (`...for members on verizon.com and its subdomains`), de sorte que la portée n'est jamais laissée à l'interprétation.

Dans les deux cas, la correspondance porte sur le label de domaine complet, ce qui exclut `@notverizon.com` et l'astuce du suffixe `@verizon.com.evil.test` d'une exécution `*.verizon.com`. Un caractère générique ailleurs qu'en tête n'est pas pris en charge et est traité comme un caractère littéral plutôt que d'élargir discrètement le filtre.

Les deux n'ont pas le même coût. **Une adresse** est résolue en son DN et comparée par Exchange lui-même (`Get-Recipient -Filter "Members -eq '<DN>'"`), sans parcourir tous les groupes du tenant. **Un domaine** ne le peut pas : il n'existe aucun filtre côté serveur pour *« a un membre dont l'adresse se termine par @x »*, donc chaque liste est lue puis filtrée. Sur un grand tenant, cela représente un appel `Get-DistributionGroupMember` par liste — plus lent, et bon à savoir avant de le lancer sur des milliers de groupes.

La correspondance porte sur l'adresse principale, **chaque alias** et — pour les contacts et utilisateurs de messagerie — `ExternalEmailAddress`. C'est ce dernier point qui compte : un contact Verizon dans une liste de distribution est généralement un contact de messagerie dont l'adresse SMTP principale ressemble à `marc.dubois@contoso.onmicrosoft.com`, avec `@be.verizon.com` uniquement dans son adresse externe. Une correspondance sur la seule adresse principale ne trouverait rien.

Lorsqu'un filtre est actif, les deux feuilles gagnent une colonne :

| Colonne | Feuille | Signification |
|--------|-------|---------|
| `Treffers` | `Overzicht` | combien de membres de cette liste correspondent |
| `Treffer op` | `Leden` | l'adresse sur laquelle ce membre a correspondu, vide s'il n'a pas correspondu |

`Treffer op` contient l'**adresse**, pas un Ja/Nee — une personne peut correspondre sur un alias qui n'apparaît nulle part ailleurs dans le rapport, et `Ja` à côté de `sara@contoso.com` ne fait que soulever la question du pourquoi. `sara.willems@be.verizon.com` y répond.

Les listes correspondantes sont exportées **en entier**, afin que le client voie qui d'autre y figure ; triez ou filtrez sur `Treffer op` pour n'obtenir que les résultats.

Appartenance directe uniquement — une personne dans un groupe imbriqué n'est pas un résultat. Le groupe imbriqué lui-même apparaît bien comme ligne de membre, avec `Distributielijst` comme type de membre.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-Group` | Non | Une liste (nom, alias ou adresse e-mail). S'il est omis, chaque liste figure dans le rapport |
| `-Member` | Non | Uniquement les listes contenant cette adresse (`jan@contoso.com`), ce domaine (`@be.verizon.com`), ou ce domaine et ses sous-domaines (`*.verizon.com`) |
| `-Recurse` | Non | Développer les groupes imbriqués, afin que les personnes derrière une liste imbriquée figurent aussi dans le rapport |
| `-IncludeDynamic` | Non | Inclure aussi les groupes de distribution dynamiques (évalués en direct, une requête par groupe) |
| `-IncludeM365Groups` | Non | Inclure aussi les groupes Microsoft 365, y compris ceux associés à une équipe Teams |
| `-OutputPath` | Non | Chemin du `.xlsx` (par défaut : `C:\Temp\Distributielijsten_<timestamp>.xlsx`) |
| `-Csv` | Non | Écrire deux fichiers CSV au lieu d'Excel |
| `-TenantId` | Non | ID de tenant Entra ID ou domaine |

**Exemples**

```powershell
# Chaque liste de distribution avec tous ses membres
.\Get-DistributionGroupMembers.ps1

# Sur quelles listes Jan figure-t-il ? (et qui d'autre y figure)
.\Get-DistributionGroupMembers.ps1 -Member "jan@contoso.com"

# Quelles listes contiennent encore des adresses sur un domaine partenaire ?
.\Get-DistributionGroupMembers.ps1 -Member "@be.verizon.com"

# Idem, mais en trouvant aussi les personnes situées dans une liste imbriquée
.\Get-DistributionGroupMembers.ps1 -Member "@be.verizon.com" -Recurse

# Tout Verizon : le domaine racine et chaque sous-domaine, listes imbriquées développées
.\Get-DistributionGroupMembers.ps1 -Member "*.verizon.com" -Recurse

# Une seule liste, vers un chemin fixe
.\Get-DistributionGroupMembers.ps1 -Group "helpdesk@contoso.com" -OutputPath "C:\Reports\helpdesk.xlsx"

# Tout ce qui peut recevoir du courrier en tant que groupe
.\Get-DistributionGroupMembers.ps1 -IncludeDynamic -IncludeM365Groups
```

**Remarques**
- Nécessite [ImportExcel](https://github.com/dfinke/ImportExcel) pour le `.xlsx`. S'il manque, le script propose de l'installer, et écrit deux fichiers CSV (`*-overzicht.csv`, `*-leden.csv`) si vous refusez — un module manquant ne vous coûte jamais le rapport. `Install-Modules.ps1` l'installe
- Une liste sans membres obtient une ligne `(geen leden)` dans la feuille `Leden` au lieu d'en être discrètement absente — une liste vide est exactement ce qu'un client veut repérer
- L'appartenance est lue par groupe ; une personne qui ne figure sur aucune liste n'apparaît donc nulle part — le rapport couvre l'appartenance aux groupes, pas l'annuaire des utilisateurs
- Un filtre de domaine qui ne trouve rien affiche *"No distribution list has a member on @x"* et n'écrit aucun fichier — un classeur vide se lirait comme un rapport en échec plutôt que comme la réponse qu'il est
- La sortie CSV utilise `-UseCulture`, afin qu'un Excel néerlandais l'ouvre en colonnes plutôt qu'en un bloc de texte séparé par des virgules
- Un classeur existant à `-OutputPath` est remplacé, et non complété — sinon `Export-Excel` empilerait une deuxième exécution sur la première

---

### Get-MessageTraceReport.ps1

Répond à « qui a reçu ceci, quand exactement, et où est-ce allé ensuite ? ». Exécute un suivi des messages et indique pour chaque message l'horodatage exact en **heure locale et en UTC** (Exchange stocke les horodatages du suivi des messages en UTC), l'expéditeur, le destinataire, l'objet, le statut, la taille, l'IP d'origine/de remise, `MessageId` et `MessageTraceId`.

**Détection des transferts** — la colonne `ForwardedTo` est remplie à partir de trois signaux indépendants, et `ForwardDetection` indique lequel s'est déclenché :

| Méthode | Détecte |
|--------|---------|
| `SameMessageId` | D'autres destinataires ayant reçu le même `MessageId` — transfert SMTP, règles de redirection, développement de groupes de distribution |
| `RedirectHop` | Sauts de redirection / de règle de transport extraits de `Get-MessageTraceDetailV2` (nécessite `-IncludeDetails`) |
| `ClientForward(subject match)` | Un message ultérieur envoyé **par** le destinataire avec le même objet normalisé — un « Transférer » d'Outlook, qui reçoit un tout nouveau `MessageId`. Heuristique ; les réponses à l'expéditeur d'origine sont exclues |

> **Pourquoi la cible du transfert est tracée séparément :** un transfert de boîte aux lettres ou une règle de redirection conserve l'**expéditeur d'origine** sur la copie transférée. La ligne de suivi correspondant à la remise à la cible du transfert ne mentionne donc la boîte aux lettres tracée *ni comme expéditeur ni comme destinataire* — filtrer sur la seule boîte aux lettres ne la renverrait jamais. Le script résout les cibles de transfert configurées **avant** le suivi et les ajoute comme filtres de destinataire supplémentaires, de sorte que la remise effective apparaisse avec son propre horodatage exact. `-ResolveSiblings` va plus loin et retrace chaque `MessageId` trouvé sans aucun filtre, ce qui détecte aussi les cibles de transfert qui ne sont plus configurées (une règle supprimée après avoir fait son œuvre laisse toujours ses remises dans le suivi).

En plus de cela, le script indique le transfert **configuré** de chaque boîte aux lettres interne qui apparaît dans le suivi — `ForwardingSMTPAddress` / `ForwardingAddress` plus toute règle de boîte de réception avec `ForwardTo` / `RedirectTo` / `ForwardAsAttachmentTo` — de sorte qu'un transfert qui ne s'est pas déclenché dans la fenêtre tracée reste visible.

Utilise `Get-MessageTraceV2` lorsqu'il est disponible et se rabat sur l'ancien `Get-MessageTrace`, désormais retiré. Les plages plus longues que la limite V2 sont automatiquement découpées en tranches de 10 jours, et chaque tranche est paginée jusqu'à épuisement.

**Paramètres**

| Paramètre | Obligatoire | Par défaut | Description |
|-----------|----------|---------|-------------|
| `-Mailbox` | Non | — | Tracer les deux sens pour cette adresse (envoyé **et** reçu) et récupérer sa configuration de transfert |
| `-SenderAddress` | Non | — | Filtrer sur l'adresse de l'expéditeur. Alias `-Sender` (`$Sender` est une variable automatique PowerShell, elle ne peut donc pas être le vrai nom du paramètre) |
| `-Recipient` | Non | — | Filtrer sur l'adresse du destinataire |
| `-ForwardAddress` | Non | — | Adresse(s) de transfert/d'exfiltration connue(s) à tracer comme destinataires, en plus de la configuration de transfert découverte. À utiliser lorsque le transfert a **déjà été supprimé** — il ne reste alors plus de configuration à trouver, mais ses remises passées figurent toujours dans le suivi |
| `-Subject` | Non | — | Filtre d'objet côté client, caractères génériques autorisés (le suivi des messages ne peut pas filtrer sur l'objet côté serveur) |
| `-MessageId` | Non | — | MessageId Internet à tracer, avec ou sans chevrons |
| `-Days` | Non | `2` | Nombre de jours en arrière à partir de `-EndDate`. Ignoré si `-StartDate` est fourni |
| `-StartDate` | Non | — | Début explicite de la fenêtre (heure locale) |
| `-EndDate` | Non | maintenant | Fin explicite de la fenêtre (heure locale) |
| `-Status` | Non | — | `Delivered`, `Failed`, `Pending`, `Expanded`, `Quarantined`, `FilteredAsSpam`, `GettingStatus`, `None` |
| `-IncludeDetails` | Non | désactivé | Récupérer le détail de remise par saut — c'est ce qui révèle les cibles de redirection / de règle de transport. Lent et soumis à limitation |
| `-MaxDetailLookups` | Non | `50` | Plafond de recherches de détail par saut ; une troncature est signalée explicitement |
| `-ResolveSiblings` | Non | désactivé | Retracer chaque `MessageId` trouvé sans filtre expéditeur/destinataire pour révéler **tous** les destinataires — détecte les cibles de transfert qui ne sont plus configurées. Un appel supplémentaire par MessageId |
| `-MaxSiblingLookups` | Non | `100` | Plafond de recherches de messages apparentés |
| `-SkipForwardingConfig` | Non | désactivé | Ignorer l'inspection du transfert de boîte aux lettres / des règles de boîte de réception |
| `-OutputPath` | Non | `C:\Temp\` / `~/Downloads` | Chemin du CSV principal. Les rapports de détail et de transfert sont écrits à côté avec les suffixes `_Details` / `_ForwardingConfig` |
| `-TenantId` | Non | — | ID de tenant Entra ID ou domaine |

**Exemples**

```powershell
# Tout ce qu'une boîte aux lettres a envoyé et reçu ces 2 derniers jours, transferts compris
.\Get-MessageTraceReport.ps1 -Mailbox "user@contoso.com"

# Un flux précis sur les 30 derniers jours, avec le détail par saut
.\Get-MessageTraceReport.ps1 -Sender "boss@contoso.com" -Recipient "user@contoso.com" -Days 30 -IncludeDetails

# Où ce message précis a-t-il abouti ?
.\Get-MessageTraceReport.ps1 -MessageId "<abc123@contoso.com>" -Days 10 -IncludeDetails

# Transfert externe suspecté — vue complète, y compris les cibles qui ne sont plus configurées
.\Get-MessageTraceReport.ps1 -Mailbox "facturen@contoso.com" -Days 10 -ResolveSiblings -IncludeDetails

# Le transfert a déjà été supprimé, mais l'adresse est connue — la tracer quand même
.\Get-MessageTraceReport.ps1 -Mailbox "facturen@contoso.com" -Days 10 -ResolveSiblings `
    -ForwardAddress "exfil@lookalike-domain.nl"

# Tous les messages en échec d'un expéditeur dans une fenêtre explicite
.\Get-MessageTraceReport.ps1 -Sender "noreply@contoso.com" -Status Failed `
    -StartDate (Get-Date "2026-08-01") -EndDate (Get-Date "2026-08-08")
```

**Remarques**
- Le suivi des messages conserve **90 jours** ; le script avertit lorsque la fenêtre demandée remonte au-delà
- La lecture des règles de boîte de réception nécessite des autorisations sur la boîte aux lettres — les boîtes aux lettres illisibles sont ignorées sans message (utilisez `-Verbose` pour voir lesquelles)
- `-IncludeDetails` effectue un appel d'API par message et est soumis à la limitation d'Exchange Online ; augmentez `-MaxDetailLookups` en connaissance de cause

**Module requis**

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

---

### Remove-PhishingMessage.ps1

Complément de réponse aux incidents de `Get-MessageTraceReport.ps1` : le suivi vous indique **qui a reçu** l'hameçonnage, ce script **le retire**. Essai à blanc par défaut — rien n'est supprimé sans `-Apply`.

**Deux moteurs**

| Moteur | Comment il trouve les messages | À utiliser quand |
|--------|----------------------|-------------|
| `Purview` | Une Content Search KQL sur tout le tenant, puis `New-ComplianceSearchAction -Purge` | Vous ne connaissez **pas** les destinataires, ou vous avez besoin d'un **HardDelete**. Indique des nombres **par boîte aux lettres**, pas les messages individuels |
| `Graph` | Parcourt chaque boîte aux lettres cible via l'API de messagerie Graph et supprime message par message | Vous connaissez **bien** les destinataires (grâce au suivi) et voulez que le message disparaisse **maintenant**, avec un rapport par message |

Le moteur est `Graph` par défaut lorsque `-Mailbox` est fourni, et `Purview` sinon. Remplacez-le avec `-Engine`.

> **Pourquoi deux moteurs.** Purview lit l'**index de recherche**, qui a environ 15 à 30 minutes de retard sur la remise — une purge lancée juste après l'arrivée de l'hameçonnage peut honnêtement indiquer *0 résultat* tout en laissant le message dans chaque boîte de réception. Graph interroge directement la boîte aux lettres et n'a pas ce retard, mais il a besoin de la liste des destinataires et ne peut pas écrire dans `Recoverable Items\Purges`, il ne peut donc pas supprimer définitivement. Pendant une campagne en cours, l'enchaînement habituel est : suivi → **Graph** immédiatement sur les destinataires connus → balayage **Purview** sur tout le tenant une demi-heure plus tard pour attraper le reste.

**Types de suppression**

| Valeur | Aboutit dans | L'utilisateur peut-il récupérer ? | Moteurs |
|-------|----------|-------------------|---------|
| `Recycle` | Éléments supprimés | Oui, facilement | Graph |
| `SoftDelete` *(par défaut)* | `Recoverable Items\Deletions` | Oui, via « Récupérer les éléments supprimés » | Les deux |
| `HardDelete` | `Recoverable Items\Purges` | Non — conservé uniquement si la boîte aux lettres est sous conservation (hold) | Purview |

**Paramètres**

| Paramètre | Obligatoire | Par défaut | Description |
|-----------|----------|---------|-------------|
| `-Mailbox` | Non | — | Adresse(s) de la ou des boîtes aux lettres cibles. Obligatoire pour Graph sauf avec `-AllMailboxes`. Omis avec Purview = chaque boîte aux lettres du tenant |
| `-AllMailboxes` | Non | désactivé | Balayer chaque boîte aux lettres. Implicite pour Purview ; pour Graph, c'est une requête par boîte aux lettres, et c'est lent |
| `-MessageId` | Non | — | MessageId Internet, avec ou sans chevrons. **Le sélecteur précis** — correspond à ce seul message et à rien d'autre |
| `-SenderAddress` | Non | — | Adresse de l'expéditeur. Alias `-Sender` (`$Sender` est une variable automatique PowerShell) |
| `-Subject` | Non | — | Purview le compare comme une expression indexée ; Graph compare côté client et accepte les caractères génériques |
| `-AttachmentName` | Non | — | Nom de fichier de la pièce jointe, caractères génériques autorisés (par ex. `*.html`) |
| `-BodyContains` | Non | — | Mot ou expression dans le corps. **Purview uniquement** — Graph devrait télécharger chaque corps |
| `-ReceivedAfter` | Non | — | Uniquement les messages reçus à ce moment ou après (heure locale) |
| `-ReceivedBefore` | Non | — | Uniquement les messages reçus à ce moment ou avant (heure locale) |
| `-Engine` | Non | voir ci-dessus | `Purview` ou `Graph` |
| `-DeleteType` | Non | `SoftDelete` | `Recycle`, `SoftDelete` ou `HardDelete` (voir le tableau ci-dessus) |
| `-Apply` | Non | désactivé | **Supprimer réellement.** Sans ce commutateur, le script indique seulement ce qu'il a trouvé |
| `-SearchName` | Non | `Phish_<timestamp>` | Nom de la Content Search à créer. Purview exige des noms uniques |
| `-KeepSearch` | Non | désactivé | Conserver la Content Search ensuite pour pouvoir l'examiner dans le portail Purview |
| `-IncludeCalendar` | Non | désactivé | Supprimer aussi les **éléments de calendrier** correspondants, pas seulement le courrier. Fonctionne avec **les deux moteurs** ; nécessite `-Subject` ou `-SenderAddress` |
| `-CalendarDaysBack` | Non | `30` | Jusqu'où remonter dans le calendrier |
| `-CalendarDaysForward` | Non | `365` | Jusqu'où avancer dans le calendrier |
| `-VerifyWithGraph` | Non | désactivé | Après une purge Purview, vérifier les boîtes aux lettres concernées via Graph pour confirmer que les messages ont réellement disparu. Nécessite la même session Graph app-only que `-Engine Graph` |
| `-MaxPurgeRounds` | Non | `10` | Purview purge au maximum 10 éléments par boîte aux lettres et par action, le script procède donc par tours. 10 tours = jusqu'à 100 éléments par boîte aux lettres |
| `-MaxMessagesPerMailbox` | Non | `500` | Plafond de sécurité Graph par boîte aux lettres ; son atteinte est signalée explicitement |
| `-TimeoutMinutes` | Non | `30` | Durée d'attente de la fin d'une recherche ou d'une action de purge |
| `-OutputPath` | Non | `C:\Temp\` / `~/Downloads` | Chemin du rapport CSV |
| `-TenantId` | Non | — | ID de tenant ou domaine, utilisé lorsque le script doit se connecter lui-même |
| `-ClientId` | Non | — | Votre propre App Registration pour l'authentification Graph app-only — ignore l'application temporaire automatique |
| `-ClientSecret` | Non | — | Secret client pour `-ClientId` |
| `-CertificateThumbprint` | Non | — | Empreinte du certificat pour `-ClientId` |

**Exemples**

```powershell
# 1. Que serait-il supprimé sur tout le tenant ? (sans -Apply = rien n'est supprimé)
.\Remove-PhishingMessage.ps1 -MessageId "<abc123@evil.example>"

# 1b. Purger, puis confirmer via Graph que le message a réellement disparu
.\Remove-PhishingMessage.ps1 -MessageId "<abc123@evil.example>" `
    -DeleteType HardDelete -Apply -VerifyWithGraph

# 2. Idem, cette fois en purgeant réellement, hors de portée de la récupération par l'utilisateur
.\Remove-PhishingMessage.ps1 -MessageId "<abc123@evil.example>" `
    -DeleteType HardDelete -Apply

# 3. Destinataires connus grâce au suivi des messages — immédiat, sans retard d'index
.\Remove-PhishingMessage.ps1 `
    -Mailbox "a@contoso.com","b@contoso.com" `
    -Sender  "no-reply@evil.example" `
    -Subject "*password expires*" `
    -Apply

# 4. Balayage de campagne : tout ce qui vient d'un expéditeur dans une fenêtre, sur tout le tenant
.\Remove-PhishingMessage.ps1 -Sender "no-reply@evil.example" `
    -ReceivedAfter (Get-Date "2026-08-30") -DeleteType HardDelete -Apply

# 5. INVITATION DE RÉUNION d'hameçonnage, sur tout le tenant : suppression définitive du courrier ET des événements
.\Remove-PhishingMessage.ps1 -Sender "no-reply@evil.example" -Subject "kick-off" `
    -DeleteType HardDelete -IncludeCalendar -Apply -VerifyWithGraph

# 6. Campagne avec pièce jointe HTML
.\Remove-PhishingMessage.ps1 -AttachmentName "*.html" `
    -Sender "billing@evil.example" -Apply
```

**Déroulement type d'un incident**

```powershell
# Qui l'a reçu, et où est-il allé ?
.\Get-MessageTraceReport.ps1 -Sender "no-reply@evil.example" -Days 2 -ResolveSiblings

# Le retirer tout de suite chez les destinataires connus (sans retard d'index)
.\Remove-PhishingMessage.ps1 -Mailbox $recipients -MessageId "<abc123@evil.example>" -Apply

# ~30 minutes plus tard, balayer le tenant pour tout ce que le suivi a manqué
.\Remove-PhishingMessage.ps1 -MessageId "<abc123@evil.example>" -DeleteType HardDelete -Apply
```

**Autorisations requises**

| Moteur | Autorisation |
|--------|-----------|
| `Purview` | Appartenance au rôle **Search And Purge** — en pratique le groupe de rôles *Organization Management* ou *eDiscovery Manager* dans le portail de conformité Purview. Se connecte via `Connect-IPPSSession -EnableSearchOnlySession` |
| `Graph` | `Mail.ReadWrite` en app-only. **Vous n'avez pas à le mettre en place vous-même** — voir les trois voies ci-dessous |

**Comment le moteur Graph (et `-VerifyWithGraph`) obtient son accès**

Le même modèle en trois voies que [`Move-InboxToArchive.ps1`](Move-InboxToArchive.ps1) ([docs](#move-inboxtoarchiveps1)) et les scripts de rapport SharePoint, essayées dans l'ordre :

| # | Voie | Ce qu'il faut |
|---|-------|---------------|
| 1 | Une session Graph app-only que vous avez déjà établie | Rien — elle est utilisée telle quelle |
| 2 | `-ClientId` + `-TenantId` + (`-ClientSecret` ou `-CertificateThumbprint`) | Votre propre application avec l'autorisation d'application `Mail.ReadWrite`, consentement administrateur accordé. **Avec `-ClientSecret`, c'est la voie la plus robuste** — elle obtient son jeton en REST pur et ne charge jamais le SDK Graph |
| 3 | **Automatique** — connexion par code d'appareil, puis une App Registration de courte durée qui s'accorde elle-même `Mail.ReadWrite`, remet un jeton app-only et est **supprimée à la fin de l'exécution** | Global Administrator ou Privileged Role Administrator pour cette connexion ponctuelle. Aucun module supplémentaire |

La voie 3 est celle qui s'applique lorsque vous ne fournissez rien, de sorte que `-VerifyWithGraph` fonctionne d'emblée. Le rôle délégué accorde le consentement, il n'y a donc pas d'écran de consentement administrateur séparé. Si la configuration échoue à mi-chemin, l'application partiellement créée est supprimée avant que l'erreur ne soit signalée — aucun orphelin ne reste dans Entra ID.

Les voies 2 (avec `-ClientSecret`) et 3 reposent toutes deux sur du REST pur — le flux par code d'appareil pour la connexion, l'API REST Graph pour créer et supprimer l'App Registration. **Aucune ne charge le SDK Graph**, ce qui leur permet de fonctionner dans la même session déjà connectée à Exchange. La voie 3 affiche un code à saisir sur `microsoft.com/devicelogin` :

```
  ------------------------------------------------------------
   To sign in, use a web browser to open https://microsoft.com/devicelogin
   and enter the code ABCD-EFGH to authenticate.
  ------------------------------------------------------------
```

> **Exchange et Graph se disputent MSAL.** `ExchangeOnlineManagement` et `Microsoft.Graph.Authentication` embarquent chacun leur propre `Microsoft.Identity.Client`, et .NET ne charge que le premier qu'un processus touche. Une purge Purview (qui se connecte à Exchange) suivie de `-VerifyWithGraph` dans la même fenêtre amène donc le SDK Graph à appeler un MSAL dont l'API ne correspond pas, et cela échoue avec `Method not found: ... WithLogging(...)` — ce qui ne ressemble en rien au conflit de versions qu'il est en réalité.
>
> Les voies 2 (`-ClientSecret`) et 3 contournent entièrement ce problème en ne chargeant jamais le SDK. Seules la variante `-CertificateThumbprint` et la réutilisation d'une session `Connect-MgGraph` existante passent encore par lui, et toutes deux signalent le conflit pour ce qu'il est plutôt que de vous laisser déchiffrer la trace de pile.
>
> Lorsque `-VerifyWithGraph` est utilisé avec le moteur Purview, l'accès Graph est établi **avant** la purge, de sorte qu'une vérification impossible est signalée d'emblée plutôt qu'après la disparition des messages. La purge a lieu dans tous les cas — une vérification en échec ne signifie jamais une purge en échec.

> `Mail.ReadWrite` délégué n'atteint jamais que *votre propre* boîte aux lettres ; une session déléguée n'est donc volontairement **pas** acceptée pour le moteur Graph, et le script passe à la voie 2 ou 3. Notez que `Mail.ReadWrite` (application) donne accès à **toutes** les boîtes aux lettres du tenant ; limitez l'application avec `New-ApplicationAccessPolicy` si c'est plus large que souhaité.

> Compatible GDAP : dans une session GDAP (`$global:authMode -eq 'GDAP'`, définie par `Connect-Tenant` / `load.ps1`), `-TenantId` est déduit du tenant client sélectionné, comme pour les scripts SharePoint.

**Remarques**
- Le calendrier nécessite `Calendars.ReadWrite`, que `Mail.ReadWrite` ne couvre pas. L'application temporaire automatique l'accorde **uniquement lorsque `-IncludeCalendar` est utilisé**, et le script vérifie la revendication roles dans le jeton émis avant de faire quoi que ce soit — un jeton app-only est délivré que l'attribution soit effective ou non, et un jeton émis trop tôt est mis en cache pendant une heure, ce qui se traduirait sinon par un 403 sur chaque boîte aux lettres pendant toute l'exécution
- **Une invitation de réunion d'hameçonnage n'est qu'à moitié supprimée lorsque le courrier est supprimé.** L'invitation laisse un événement dans le calendrier, et `/messages` et `/events` sont des collections distinctes — aucun des deux moteurs ne touche au calendrier par défaut. `-IncludeCalendar` les balaie aussi, en comparant sur `-Subject` ou `-SenderAddress` (en tant qu'organisateur). Sans sélecteur, il refuse plutôt que de parcourir tout le calendrier
- **`-IncludeCalendar` fonctionne aussi avec le moteur Purview**, et cette combinaison constitue le nettoyage complet d'un hameçonnage par invitation de réunion : Purview supprime définitivement l'invitation sur tout le tenant, puis une passe Graph supprime les événements laissés dans exactement les boîtes aux lettres touchées par la recherche. La passe calendrier nécessite le même accès Graph que `-VerifyWithGraph`, et une boîte aux lettres qu'elle ne peut pas atteindre est signalée plutôt que comptée comme propre
- La vérification suit le même principe : avec `-IncludeCalendar`, elle contrôle aussi le calendrier, et sans lui, elle affiche *"mail only — calendar items are not checked"* plutôt que de déclarer une boîte aux lettres propre sur la base de preuves incomplètes
- **Rien dans Purview ne peut confirmer une purge.** L'action de purge indique ce que le service croit avoir fait, et l'index de recherche continue de lister les éléments purgés pendant jusqu'à ~30 minutes — relancer le script n'est donc pas une vérification. `-VerifyWithGraph` est la seule vérification sans retard : elle repose *la même* requête que celle sur laquelle le moteur Graph supprime, directement sur les boîtes aux lettres touchées par la recherche. Les éléments supprimés de manière réversible ou définitive se trouvent dans Recoverable Items, que Graph ne liste pas ; un message purgé apparaît donc à juste titre comme disparu
- La vérification distingue **« n'a pas pu vérifier »** de **« propre »**. Une boîte aux lettres qui renvoie 403 est signalée comme non vérifiée, jamais comme confirmée. Elle ne fait jamais échouer l'exécution non plus — une purge déjà effectuée n'est pas signalée comme échouée parce que la vérification n'a pas pu s'exécuter
- **Purview ne peut pas vous montrer les messages individuels.** Content Search indique le nombre d'éléments par boîte aux lettres ; l'action d'aperçu qui renvoyait autrefois l'expéditeur et l'objet de chaque message est [documentée comme réservée à l'environnement local](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/new-compliancesearchaction?view=exchange-ps) depuis les changements eDiscovery de mai 2025. Pour le détail par message, prenez la liste des boîtes aux lettres issue de l'exécution Purview et relancez ces adresses avec `-Engine Graph`
- Purview exécute les recherches et les purges côté serveur, et elles prennent régulièrement plusieurs minutes. Le script indique l'état de la tâche et le temps écoulé environ toutes les 15 secondes pendant l'attente, de sorte qu'une étape lente paraît visiblement lente plutôt que bloquée, et `-TimeoutMinutes` (30 par défaut) la borne
- **Les tours sont planifiés, pas interrogés.** Une purge supprime au maximum 10 éléments par boîte aux lettres et par action, le script calcule donc `ceil(max items per mailbox / 10)` à partir de la première recherche. Il ne boucle volontairement *pas* jusqu'à ce que l'index se calme : l'index a jusqu'à ~30 minutes de retard sur une purge, ce qui reviendrait à repurger les mêmes éléments puis à signaler une fausse troncature. Ce que chaque tour a réellement supprimé est relu à partir de l'action de purge elle-même
- Une seule content search purge au maximum **50 000 boîtes aux lettres** ; au-delà, le script avertit et vous devez procéder par lots avec `-Mailbox`. Pour le traitement en masse, Microsoft renvoie vers l'API Graph `ediscoverySearch: purgeData` (100 éléments par emplacement)
- **Nécessite ExchangeOnlineManagement 3.9.0+** pour le moteur Purview. Content Search s'exécute sur un backend qu'une simple connexion IPPS n'atteint plus : sans `-EnableSearchOnlySession`, les cmdlets sont présentes mais `Start-ComplianceSearch` échoue à l'initialisation. Le script passe ce commutateur lorsqu'il se connecte lui-même. **Si vous étiez déjà connecté sans lui, la session ne peut pas être réparée depuis le processus** — ouvrez une nouvelle fenêtre PowerShell et laissez le script se connecter
- **Vous n'avez pas besoin de l'objet complet.** `-Subject` correspond à un fragment : avec Purview, il devient une expression KQL, qui correspond n'importe où dans l'objet, donc `-Subject "kick-off meeting"` trouve chaque message contenant ces mots dans cet ordre. Un objet long et chargé de ponctuation est en fait le choix *fragile* — les virgules, apostrophes et heures comme `14:09` sont mal découpées en mots. Court et distinctif, c'est gagnant
- **La correspondance à l'intérieur d'un mot** est la seule chose que KQL ne sait pas faire. Un `*` initial est supprimé (le script avertit plutôt que de l'ignorer discrètement) et un `*` final ne fonctionne que sur un mot unique, car un caractère générique est inopérant dans une expression entre guillemets. Pour une vraie recherche de sous-chaîne, utilisez `-Engine Graph` avec `-Mailbox`, qui compare côté client et prend les caractères génériques tels qu'écrits
- **Les sélecteurs se combinent en ET (AND).** La paire habituelle en matière d'hameçonnage est l'expéditeur plus un fragment d'objet : `-Sender "no-reply@evil.example" -Subject "kick-off meeting"`. Lorsque l'expéditeur change d'adresse, `-BodyContains "a distinctive sentence"` (Purview uniquement) est souvent le sélecteur le plus durable
- Une exécution qui échoue en cours de route ne laisse plus sa Content Search derrière elle — le nettoyage s'exécute dans un `finally`. Les recherches laissées par des exécutions plus anciennes peuvent être listées avec `Get-ComplianceSearch | Where-Object Name -like 'Phish_*'` et supprimées avec `Remove-ComplianceSearch`
- Le script **refuse de s'exécuter** sans au moins l'un des paramètres `-MessageId`, `-SenderAddress`, `-Subject`, `-AttachmentName` ou `-BodyContains` — une plage de dates seule correspondrait à chaque message de chaque boîte aux lettres
- **Retard d'index** (Purview uniquement) : un message remis au cours des ~30 dernières minutes peut ne pas encore être indexé. Un résultat `0 hits` juste après la remise ne prouve pas que l'hameçonnage a disparu — attendez et relancez, ou utilisez le moteur Graph
- La purge ne couvre que la **boîte aux lettres principale** — aucun des deux moteurs n'atteint la boîte aux lettres d'archive
- KQL ne prend pas en charge les caractères génériques dans une expression ; avec `-Subject "*invoice*"`, les caractères génériques sont donc retirés sur le moteur Purview et la comparaison se fait sur l'expression ; sur Graph, les caractères génériques fonctionnent tels qu'écrits
- Chaque exécution écrit un rapport CSV de ce qui a été trouvé et de ce qui a été supprimé

**Modules requis**

```powershell
Install-Module ExchangeOnlineManagement       -Scope CurrentUser
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser   # facultatif, voir ci-dessous
```

Microsoft.Graph.Authentication n'est nécessaire que pour réutiliser une session `Connect-MgGraph` existante ou pour utiliser `-CertificateThumbprint`. Les voies `-ClientSecret` et application temporaire automatique fonctionnent en REST pur et n'ont besoin de rien d'autre qu'ExchangeOnlineManagement.

---

### Restore-MailboxMessages.ps1

Annule une mauvaise journée dans une boîte aux lettres : les messages **déplacés** ou **supprimés** à une date donnée retournent dans le dossier d'où ils venaient, et l'exécution indique **qui l'a fait** — compte, en tant que propriétaire / délégué / administrateur, avec quel client, depuis quelle IP. Aperçu par défaut — rien n'est déplacé sans `-Apply`.

**Trois sources, car aucune ne couvre tous les cas**

| Partie | Où elle cherche | Comment elle connaît le dossier d'origine | Qui l'a fait |
|------|----------------|----------------------------------|------------|
| Journal d'audit | `Search-UnifiedAuditLog` — `Move`, `MoveToDeletedItems`, `SoftDelete`, `HardDelete` sur cette boîte aux lettres | Enregistre le dossier source de chaque déplacement | **Oui** — compte, type d'ouverture de session, client, IP, ID d'application |
| `Deleted` | `Get-/Restore-RecoverableItems` sur Éléments supprimés, Recoverable Items et Purges (conservés uniquement sous conservation), filtré sur le moment de la suppression | Exchange le conserve lui-même (`LastParentPath`) | Recherché dans le journal d'audit par objet et heure |
| `Deleted` sans le rôle | Graph : **tout** ce qui, dans Éléments supprimés et Recoverable Items\Deletions, a changé dans la fenêtre, plus chaque `MoveToDeletedItems` / `SoftDelete` audité | Le journal d'audit pour les suppressions auditées ; sinon — et pour tout ce qui a été supprimé depuis Éléments supprimés lui-même — la **boîte de réception** | À partir de l'enregistrement d'audit (exact, par MessageId) |
| `Moved` | Chaque `Move` audité ce jour-là, remonté jusqu'au **premier** dossier quitté par le message, retrouvé via Graph par MessageId Internet et remis en place | Uniquement à partir du journal d'audit | À partir de l'enregistrement d'audit |

> **Pourquoi le journal d'audit compte doublement.** Un simple déplacement ne laisse aucune trace de l'endroit d'où venait un message — ni dans Graph, ni dans Exchange. L'enregistrement d'audit est le seul endroit qui le sait, et ce même enregistrement nomme la personne. Les déplacements **hors de** Éléments supprimés ou de Recoverable Items sont laissés de côté : il s'agissait de restaurations, et les annuler supprimerait à nouveau le message.

Les messages arrivés dans **Archive** sans enregistrement d'audit — Exchange n'audite pas par défaut le `Move` du propriétaire lui-même, et l'archivage piloté par une application (par ex. [`Move-InboxToArchive.ps1`](Move-InboxToArchive.ps1)) peut aussi manquer — sont **listés** à partir du dossier Archive selon leur date de modification. Ils ne sont déplacés vers la boîte de réception qu'avec `-UnauditedArchiveToInbox`, car lire ou marquer un message modifie aussi cette date.

**Paramètres**

| Paramètre | Obligatoire | Par défaut | Description |
|-----------|----------|---------|-------------|
| `-Mailbox` | Oui | — | UPN ou adresse SMTP principale |
| `-Date` | L'un de | — | Le jour où les messages ont été déplacés ou supprimés (heure locale, journée entière) |
| `-After` | ceux-ci | — | Début d'une fenêtre (heure locale) au lieu de `-Date`. Sans `-Before` : **tout depuis cette date jusqu'à maintenant**. Alias `-From`, `-Since` |
| `-Before` | Non | maintenant | Fin de cette fenêtre |
| `-Include` | Non | `Deleted`, `Moved` | Quelle partie exécuter. Le rapport « qui a fait quoi » est toujours produit |
| `-UnauditedArchiveToInbox` | Non | désactivé | Déplacer aussi vers la boîte de réception les éléments d'Archive non audités modifiés dans la fenêtre |
| `-Apply` | Non | désactivé | **Restaurer réellement.** Sans ce commutateur, l'exécution se contente de rendre compte |
| `-OutputPath` | Non | `C:\Temp\` / `~/Downloads` | Chemin du rapport CSV ; la piste d'audit est écrite à côté sous `*_Audit.csv` |
| `-TenantId` | Non | — | ID de tenant ou domaine ; nécessaire pour Graph app-only, sauf s'il peut être déduit de GDAP |
| `-ClientId` | Non | — | Votre propre App Registration (autorisation d'application Mail.ReadWrite) — ignore l'application temporaire |
| `-ClientSecret` | Non | — | Secret client pour `-ClientId` |
| `-CertificateThumbprint` | Non | — | Empreinte du certificat pour `-ClientId` |

**Exemples**

```powershell
# 1. Qu'est-ce qui a été déplacé ou supprimé le 25 septembre, par qui, et qu'est-ce qui serait remis en place ?
.\Restore-MailboxMessages.ps1 -Mailbox "user@contoso.com" -Date 2026-09-25

# 2. Tout remettre en place
.\Restore-MailboxMessages.ps1 -Mailbox "user@contoso.com" -Date 2026-09-25 -Apply

# 2b. Tout ce qui a été déplacé ou supprimé du 20 septembre à maintenant
.\Restore-MailboxMessages.ps1 -Mailbox "user@contoso.com" -Since 2026-09-20 -Apply

# 3. Uniquement les suppressions, dans une fenêtre précise — aucun accès Graph nécessaire
.\Restore-MailboxMessages.ps1 -Mailbox "user@contoso.com" -Include Deleted `
    -After "2026-09-25 14:00" -Before "2026-09-25 16:00" -Apply

# 4. Annuler une exécution de Move-InboxToArchive.ps1, y compris les éléments d'Archive non audités
.\Restore-MailboxMessages.ps1 -Mailbox "user@contoso.com" -Date 2026-09-25 `
    -UnauditedArchiveToInbox -Apply
```

**Sortie**

- À l'écran : chaque message avec `[Status] time | current folder -> target folder | who | subject`, puis un tableau **qui a déplacé / supprimé quoi** regroupé par compte, type d'ouverture de session, client et opération, avec la période et les IP.
- `MailRestore_<mailbox>_<timestamp>.csv` — une ligne par message : phase, heure de l'action, auteur, type d'ouverture de session, client, IP, objet, dossier actuel et dossier cible, statut.
- `..._Audit.csv` — la piste d'audit brute de la boîte aux lettres dans la fenêtre, dossiers source et destination compris.

| Statut | Signification |
|--------|---------|
| `WouldRestore` / `Restored` | Aperçu / effectué |
| `AlreadyInPlace` | Déjà revenu dans le dossier d'origine — rien à faire |
| `NotFound` | Le message déplacé n'est plus dans la boîte aux lettres (supprimé depuis — la partie `Deleted` couvre ce cas) |
| `Ambiguous` | Plusieurs copies avec le même MessageId ; laissé tel quel |
| `NotAudited` | Élément d'Archive sans enregistrement d'audit, seulement listé |
| `NotRestored` | Restore-RecoverableItems n'a rien signalé, mais l'élément se trouve toujours dans Recoverable Items ensuite |
| `Unreachable` | Supprimé définitivement (Purges) alors que l'exécution passe par la voie Graph — seul `Restore-RecoverableItems` avec le rôle peut l'atteindre |

**Autorisations requises**

| Partie | Autorisation |
|------|-----------|
| Journal d'audit | **View-Only Audit Logs** ou **Audit Logs** (Organization Management / Compliance Management) |
| `Deleted` | **Mailbox Import Export** — dans aucun groupe de rôles par défaut : `New-ManagementRoleAssignment -Role "Mailbox Import Export" -User admin@contoso.com`, puis reconnectez-vous. **Facultatif :** sans ce rôle, l'exécution restaure via Graph — tout sauf les éléments supprimés définitivement |
| `Moved` (et `Deleted` sans le rôle) | `Mail.ReadWrite` en app-only, par les trois mêmes voies que [`Remove-PhishingMessage.ps1`](#remove-phishingmessageps1) : une session app-only existante, `-ClientId`, ou une application temporaire supprimée à la fin |

**Remarques**

- Le journal d'audit a **30 à 90 minutes** (parfois 24 heures) de retard. Une exécution le jour même peut manquer les dernières actions.
- Les fenêtres plus longues (`-Since`) sont recherchées dans le journal d'audit **un jour à la fois**, car une seule recherche s'arrête à 50 000 enregistrements à l'échelle du tenant. L'exécution avertit lorsque la fenêtre remonte au-delà de ce qui est encore conservé : Recoverable Items conserve les éléments supprimés pendant `RetainDeletedItemsFor` (14 jours par défaut, 30 au maximum), sauf si la boîte aux lettres est sous conservation, et le journal d'audit conserve généralement 180 jours. Éléments supprimés lui-même n'est pas concerné par cette limite.
- L'exécution liste lesquelles des opérations `Move`, `MoveToDeletedItems`, `SoftDelete`, `HardDelete` ne sont **pas auditées** pour le propriétaire, le délégué et l'administrateur sur cette boîte aux lettres. Par défaut, le `Move` du propriétaire lui-même ne l'est pas — ces déplacements ne peuvent être ni retracés ni attribués. Activez-le pour la prochaine fois : `Set-Mailbox user@contoso.com -AuditOwner @{Add='Move'}`.
- Les déplacements effectués par une règle de boîte de réception s'exécutent en tant que la boîte aux lettres elle-même et ne sont souvent pas audités. Les déplacements vers la **boîte aux lettres d'archive en ligne** par une stratégie de rétention concernent une autre boîte aux lettres et sont hors du périmètre.
- Lorsque le dossier d'origine d'un déplacement n'existe plus, le message va dans la boîte de réception et le CSV l'indique.
