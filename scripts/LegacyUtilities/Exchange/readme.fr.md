[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [LegacyUtilities](../readme.fr.md) › **Exchange**

# Legacy Utilities — Exchange

Scripts de gestion des boîtes aux lettres et des contacts, modernisés à partir d'un ensemble d'anciens scripts ad hoc. Ils se connectent via [`Connect-M365.ps1`](../../Startup/readme.fr.md) : en délégué en tant qu'administrateur par défaut (navigateur, ou code d'appareil / client GDAP selon `load.config.ps1`), en app-only avec `-ClientId` + `-CertificateThumbprint` ou `-AppOnly` (application de `graph.appid.json`). Une session adaptée pour le bon tenant est réutilisée et reste connectée ; seule une session ouverte par le script est fermée. Chaque script accepte `-TenantId`, `-ClientId`, `-CertificateThumbprint` et `-AppOnly`.

Graph est la valeur par défaut. Cinq scripts restent sur Exchange Online PowerShell, car Graph n'a pas d'API pour ce qu'ils font (voir chaque section) ; sous GDAP, ils atteignent désormais le client avec `-DelegatedOrganization` — auparavant, `-TenantId` était passé comme `-Organization`, qui ne s'applique qu'à la connexion app-only. `Sync-UserContacts.ps1` est le seul script qui se connecte **en app-only par défaut**, car un jeton délégué ne peut pas écrire les contacts d'autres utilisateurs.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Set-MailboxFolderPermission.ps1`](Set-MailboxFolderPermission.ps1) ([docs](#set-mailboxfolderpermissionps1)) | Accorder une autorisation de dossier sur tous les dossiers d'une boîte aux lettres |
| [`Add-MailboxDelegateAccess.ps1`](Add-MailboxDelegateAccess.ps1) ([docs](#add-mailboxdelegateaccessps1)) | Accorder Full Access / Send As sur une boîte aux lettres, une liste CSV ou toutes les boîtes |
| [`New-BulkSharedMailboxes.ps1`](New-BulkSharedMailboxes.ps1) ([docs](#new-bulksharedmailboxesps1)) | Créer en masse des boîtes aux lettres partagées à partir d'un CSV |
| [`New-BulkMailContacts.ps1`](New-BulkMailContacts.ps1) ([docs](#new-bulkmailcontactsps1)) | Créer en masse des Mail Contacts à partir d'un CSV, avec ajout facultatif à un groupe de distribution |
| [`Sync-UserContacts.ps1`](Sync-UserContacts.ps1) ([docs](#sync-usercontactsps1)) | Pousser une liste de contacts partagée dans les contacts Outlook personnels des utilisateurs |
| [`Start-MailboxMessageTraceReport.ps1`](Start-MailboxMessageTraceReport.ps1) ([docs](#start-mailboxmessagetracereportps1)) | Soumettre des demandes de rapport de suivi des messages historique |
| [`Remove-DuplicateMailItems.ps1`](Remove-DuplicateMailItems.ps1) ([docs](#remove-duplicatemailitemsps1)) | Trouver/supprimer les messages en double d'un dossier de boîte aux lettres via Graph |

---

### Set-MailboxFolderPermission.ps1

Applique une même autorisation de dossier à tous les dossiers d'une boîte aux lettres (en
ignorant Sync Issues, Recoverable Items, Purges, Versions et Deletions). Destiné à l'accès
délégué à l'ensemble de la boîte lorsque `Add-MailboxPermission -AccessRights FullAccess` seul
ne convient pas. Essai à blanc par défaut.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-Mailbox` | Oui | Boîte aux lettres cible |
| `-User` | Oui | Utilisateur à qui accorder l'accès |
| `-AccessRights` | Oui | Niveau d'autorisation du dossier (Owner, Editor, Reviewer, ...) |
| `-Apply` | Non | Accorder réellement l'autorisation (par défaut : aperçu) |
| `-TenantId` | Non | Domaine ou ID du tenant (par défaut : le client GDAP si `authMode` vaut GDAP) ; l'app-only exige le domaine |
| `-ClientId` / `-CertificateThumbprint` | Non | Connexion app-only avec cette inscription d'application et ce certificat |
| `-AppOnly` | Non | Connexion app-only avec l'application de `graph.appid.json` |

```powershell
.\Set-MailboxFolderPermission.ps1 -Mailbox "shared@contoso.com" -User "j.doe@contoso.com" -AccessRights Editor -Apply
```

**Remarques**
- Reste sur Exchange Online : Graph n'a pas d'API pour les autorisations de dossiers de boîte aux lettres (seul le calendrier dispose de `calendarPermission`)

---

### Add-MailboxDelegateAccess.ps1

Regroupe en un seul script trois anciens scripts quasi identiques (accorder Full Access + Send
As sur une boîte aux lettres ; la même chose pour une liste CSV de boîtes partagées ; un octroi
global sur toutes les boîtes de l'organisation). Essai à blanc par défaut.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-User` | Oui | Utilisateur à qui accorder l'accès délégué |
| `-Mailbox` | * | Boîte aux lettres cible unique |
| `-CsvPath` | * | Liste CSV/TXT des boîtes aux lettres cibles |
| `-AllMailboxes` | * | Toutes les boîtes utilisateur/partagées du tenant |
| `-AccessRights` | Non | `FullAccess`, `SendAs` ou `Both` (par défaut) |
| `-AutoMapping` | Non | Activer le mappage automatique dans Outlook (par défaut : désactivé) |
| `-Apply` | Non | Accorder réellement l'accès (par défaut : aperçu) |
| `-OutputPath` | Non | Chemin du rapport CSV |
| `-TenantId` | Non | Domaine ou ID du tenant (par défaut : le client GDAP si `authMode` vaut GDAP) ; l'app-only exige le domaine |
| `-ClientId` / `-CertificateThumbprint` | Non | Connexion app-only avec cette inscription d'application et ce certificat |
| `-AppOnly` | Non | Connexion app-only avec l'application de `graph.appid.json` |

*Un seul paramètre parmi `-Mailbox` / `-CsvPath` / `-AllMailboxes` définit la portée.

```powershell
.\Add-MailboxDelegateAccess.ps1 -Mailbox "sales@contoso.com" -User "j.doe@contoso.com" -Apply
.\Add-MailboxDelegateAccess.ps1 -AllMailboxes -User "helpdesk@contoso.com"   # vérifier d'abord la portée
```

**Remarques**
- Reste sur Exchange Online : Full Access et Send As sont des autorisations Exchange sans API Graph

---

### New-BulkSharedMailboxes.ps1

Crée des boîtes aux lettres partagées à partir d'un CSV (`Name`, `PrimarySmtpAddress`,
`Alias` facultatif). Essai à blanc par défaut.

```powershell
.\New-BulkSharedMailboxes.ps1 -CsvPath .\sharedmailboxes.csv -Apply
```

**Remarques**
- Reste sur Exchange Online : Graph ne peut pas créer de boîtes aux lettres partagées

---

### New-BulkMailContacts.ps1

Crée des Mail Contacts à partir d'un CSV (`Name`, `ExternalEmailAddress`), en ajoutant
éventuellement chaque nouveau contact à un groupe de distribution. Essai à blanc par défaut.

```powershell
.\New-BulkMailContacts.ps1 -CsvPath .\contacts.csv -DistributionGroup "everyone@contoso.com" -Apply
```

**Remarques**
- Reste sur Exchange Online : les contacts de messagerie et l'appartenance aux groupes de distribution sont des objets Exchange ; `orgContact` dans Graph est en lecture seule et Graph ne peut pas modifier les membres des groupes de distribution

---

### Sync-UserContacts.ps1

Pousse, via Microsoft Graph, une liste de contacts CSV dans le dossier Contacts personnel d'une
liste explicite d'utilisateurs ou de chaque membre d'un groupe. Chaque contact créé est marqué
dans `PersonalNotes`, de sorte qu'une exécution ultérieure avec `-RemoveExisting` puisse
nettoyer et actualiser sans toucher aux contacts propres de l'utilisateur. Essai à blanc par
défaut.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-CsvPath` | Oui | Liste de contacts (DisplayName et EmailAddress obligatoires ; GivenName/Surname/CompanyName/BusinessPhone/MobilePhone facultatifs) |
| `-UserList` / `-GroupId` | * | Portée cible |
| `-Tag` | Non | Marqueur écrit dans PersonalNotes (par défaut : `Synced-by-Sync-UserContacts`) |
| `-RemoveExisting` | Non | Supprimer les contacts synchronisés précédemment avant de réimporter |
| `-Apply` | Non | Écrire réellement les contacts (par défaut : aperçu) |
| `-ClientId` / `-CertificateThumbprint` | Non | Inscription d'application avec l'autorisation d'application `Contacts.ReadWrite` (plus `GroupMember.Read.All` pour `-GroupId`) |
| `-AppOnly` | Non | L'application de `graph.appid.json` — déjà la valeur par défaut |
| `-Delegated` | Non | Se connecter en son propre nom ; seule votre propre boîte aux lettres peut être une cible |
| `-TenantId` | Non | Tenant (par défaut : le client GDAP) ; choisit aussi l'entrée dans `graph.appid.json` |

```powershell
.\Sync-UserContacts.ps1 -CsvPath .\companycontacts.csv -GroupId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -RemoveExisting -Apply
```

**Remarques**
- **App-only par défaut** : un jeton délégué, même celui d'un Global Administrator, ne peut écrire que les propres contacts de l'utilisateur connecté. Sans `-ClientId`/`-CertificateThumbprint`, l'application provient de `graph.appid.json` ; s'il n'y en a pas, le script s'arrête et explique pourquoi
- Avec `-Delegated`, chaque cible doit être vous-même ; le script s'arrête si un autre utilisateur figure dans la liste ou le groupe
- Seuls les utilisateurs membres de `-GroupId` sont ciblés (les groupes imbriqués et appareils sont ignorés)
- Un CSV sans `DisplayName` ou `EmailAddress` est refusé (ce contrôle ne se déclenchait jamais auparavant)

---

### Start-MailboxMessageTraceReport.ps1

Soumet une requête `Start-HistoricalSearch` par boîte aux lettres (suivi des messages, filtre
sur l'expéditeur), livrée par e-mail sous forme d'export compressé : c'est ainsi que l'on
récupère les données de suivi des messages antérieures à la fenêtre de 10 jours couverte par
`Get-MessageTrace`. Remplacement modernisé d'un ancien script qui utilisait le module retiré
MSOnline pour constituer la liste des boîtes aux lettres.

```powershell
.\Start-MailboxMessageTraceReport.ps1 -NotifyAddress "admin@contoso.com" -AllMailboxes -Apply
```

**Remarques**
- Reste sur Exchange Online : `Start-HistoricalSearch` (suivi historique des messages) n'existe que dans Exchange Online PowerShell

---

### Remove-DuplicateMailItems.ps1

Remplacement basé sur Graph d'un outil tiers EWS de suppression des éléments en double, qui n'a
pas été repris : Microsoft retire l'API EWS pour Exchange Online. Regroupe les messages d'un
dossier par `internetMessageId`, conserve le plus ancien de chaque groupe et (avec `-Apply`)
supprime les autres. Essai à blanc par défaut.

```powershell
.\Remove-DuplicateMailItems.ps1 -Mailbox "user@contoso.com" -IncludeSubfolders -Apply
```

---

**Modules requis**

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Install-Module Microsoft.Graph -Scope CurrentUser
```

**Remarques**
- En délégué, Graph n'atteint la boîte aux lettres d'un autre utilisateur que si elle est partagée avec vous : le script demande `Mail.ReadWrite` + `Mail.ReadWrite.Shared`, et l'administrateur connecté a besoin de **Full Access** sur la boîte cible (par ex. `Add-MailboxDelegateAccess.ps1 -AccessRights FullAccess`)
- Sans Full Access, exécutez en app-only (`-ClientId`/`-CertificateThumbprint` ou `-AppOnly`) avec l'autorisation d'application `Mail.ReadWrite`, de préférence limitée avec RBAC for Applications
- `-IncludeSubfolders` fonctionne à nouveau : le parcours des dossiers échouait sur le premier dossier sans sous-dossiers
