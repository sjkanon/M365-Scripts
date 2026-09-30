[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [LegacyUtilities](../readme.fr.md) › **Exchange**

# Legacy Utilities — Exchange

Scripts de gestion des boîtes aux lettres et des contacts, modernisés à partir d'un ensemble
d'anciens scripts ad hoc. Ils se connectent automatiquement à Exchange Online / Microsoft Graph
si aucune session n'est active, et réutilisent la session existante si vous êtes déjà connecté.

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
| `-TenantId` | Non | ID ou domaine du tenant Entra ID |

```powershell
.\Set-MailboxFolderPermission.ps1 -Mailbox "shared@contoso.com" -User "j.doe@contoso.com" -AccessRights Editor -Apply
```

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
| `-TenantId` | Non | ID ou domaine du tenant Entra ID |

*Un seul paramètre parmi `-Mailbox` / `-CsvPath` / `-AllMailboxes` définit la portée.

```powershell
.\Add-MailboxDelegateAccess.ps1 -Mailbox "sales@contoso.com" -User "j.doe@contoso.com" -Apply
.\Add-MailboxDelegateAccess.ps1 -AllMailboxes -User "helpdesk@contoso.com"   # vérifier d'abord la portée
```

---

### New-BulkSharedMailboxes.ps1

Crée des boîtes aux lettres partagées à partir d'un CSV (`Name`, `PrimarySmtpAddress`,
`Alias` facultatif). Essai à blanc par défaut.

```powershell
.\New-BulkSharedMailboxes.ps1 -CsvPath .\sharedmailboxes.csv -Apply
```

---

### New-BulkMailContacts.ps1

Crée des Mail Contacts à partir d'un CSV (`Name`, `ExternalEmailAddress`), en ajoutant
éventuellement chaque nouveau contact à un groupe de distribution. Essai à blanc par défaut.

```powershell
.\New-BulkMailContacts.ps1 -CsvPath .\contacts.csv -DistributionGroup "everyone@contoso.com" -Apply
```

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

```powershell
.\Sync-UserContacts.ps1 -CsvPath .\companycontacts.csv -GroupId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -RemoveExisting -Apply
```

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
