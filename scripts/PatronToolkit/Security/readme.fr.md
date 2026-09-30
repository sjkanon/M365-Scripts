[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [PatronToolkit](../readme.fr.md) › **Security**

# Patron Toolkit — Security

Rapports sur la posture de sécurité du tenant : risque lié aux consentements d'applications, détection de BEC via les règles de boîte aux lettres, alertes
de sécurité, configuration de la sécurité de la messagerie, journalisation d'audit et validation SPF/DMARC.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Get-EntraAppConsents.ps1`](Get-EntraAppConsents.ps1) ([docs](#get-entraappconsentsps1)) | Auditer à l'échelle du tenant les autorisations OAuth déléguées + d'application et signaler les étendues à haut risque |
| [`Get-SuspiciousInboxRules.ps1`](Get-SuspiciousInboxRules.ps1) ([docs](#get-suspiciousinboxrulesps1)) | Détecter les règles de boîte de réception de type BEC (transfert externe, suppression silencieuse, dossier masqué + mot-clé) |
| [`Get-SecurityAlerts.ps1`](Get-SecurityAlerts.ps1) ([docs](#get-securityalertsps1)) | Rapport des alertes de sécurité unifiées Defender/Entra |
| [`Test-EmailSecurityPosture.ps1`](Test-EmailSecurityPosture.ps1) ([docs](#test-emailsecuritypostureps1)) | Rapport de sécurité consolidé Defender for O365 / anti-spam / DLP / flux de messagerie |
| [`Test-MailboxAuditingConfig.ps1`](Test-MailboxAuditingConfig.ps1) ([docs](#test-mailboxauditingconfigps1)) | Signaler et éventuellement corriger les lacunes de l'Unified Audit Log + de l'audit par boîte aux lettres |
| [`Test-EmailAuthenticationRecords.ps1`](Test-EmailAuthenticationRecords.ps1) ([docs](#test-emailauthenticationrecordsps1)) | Valider les enregistrements DNS SPF et DMARC |

---

### Get-EntraAppConsents.ps1

Audite chaque application d'entreprise (principal de service) disposant d'autorisations déléguées
(`OAuth2PermissionGrants`) ou d'autorisations d'application
(`AppRoleAssignments`), et signale celles qui correspondent à une liste d'étendues à privilèges élevés
couramment détournées. C'est le contrôle classique des « consentements illicites » / du risque lié aux applications
tierces.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-RiskyOnly` | Non | N'inclure que les autorisations signalées High risk |
| `-OutputPath` | Non | Chemin du rapport CSV (par défaut : `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Non | ID de tenant ou domaine Entra ID |

**Exemples**

```powershell
.\Get-EntraAppConsents.ps1

.\Get-EntraAppConsents.ps1 -RiskyOnly
```

**Remarques**
- Le signalement du risque est une heuristique (correspondance de chaînes avec une liste d'étendues connues pour être risquées) — examinez
  manuellement les entrées signalées et ne considérez pas « Normal » comme une garantie de sécurité
- Étendues requises : `Application.Read.All`, `Directory.Read.All`

---

### Get-SuspiciousInboxRules.ps1

Analyse les règles de boîte de réception de chaque boîte aux lettres à la recherche de schémas souvent laissés par un compte
compromis : transfert/redirection externe, suppression silencieuse, ou déplacement des messages contenant certains mots-clés
(invoice, payment, wire, password, ...) vers un dossier rarement consulté. Par défaut, il s'agit d'un
rapport en lecture seule ; `-Apply` désactive (sans supprimer) les correspondances à forte certitude.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-Mailbox` | Non | UPN d'une seule boîte aux lettres. Si omis, toutes les boîtes aux lettres sont vérifiées |
| `-Apply` | Non | Désactiver les règles signalées à forte certitude (par défaut : rapport uniquement) |
| `-OutputPath` | Non | Chemin du rapport CSV (par défaut : `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Non | ID de tenant ou domaine Entra ID |

**Exemples**

```powershell
.\Get-SuspiciousInboxRules.ps1

.\Get-SuspiciousInboxRules.ps1 -Mailbox "user@contoso.com"

# Aperçu de ce qui serait désactivé
.\Get-SuspiciousInboxRules.ps1 -Apply -WhatIf

.\Get-SuspiciousInboxRules.ps1 -Apply
```

Prend en charge `-WhatIf` (`SupportsShouldProcess`).

**Remarques**
- « Externe » est déterminé à partir des domaines acceptés du tenant (`Get-AcceptedDomain`)
- La désactivation (et non la suppression) est volontaire — elle est réversible et conserve la règle pour
  l'analyse en réponse à incident

---

### Get-SecurityAlerts.ps1

Rapporte le flux unifié d'alertes de sécurité de Microsoft Graph (`security/alerts_v2`), qui
couvre Defender for Office 365, Defender for Endpoint, Defender for Identity, Defender for
Cloud Apps et Entra ID Protection.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-Days` | Non | Fenêtre d'historique en jours (par défaut : `30`) |
| `-Severity` | Non | Filtre : `informational`, `low`, `medium`, `high` |
| `-Status` | Non | Filtre : `new`, `inProgress`, `resolved` |
| `-OutputPath` | Non | Chemin du rapport CSV (par défaut : `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Non | ID de tenant ou domaine Entra ID |

**Exemples**

```powershell
.\Get-SecurityAlerts.ps1

.\Get-SecurityAlerts.ps1 -Days 7 -Severity high,medium -Status new,inProgress
```

**Remarques**
- Étendue requise : `SecurityAlert.Read.All`

---

### Test-EmailSecurityPosture.ps1

Rapport consolidé en lecture seule sur la configuration de la sécurité de la messagerie Exchange Online / Defender for Office 365 :
Safe Links, Safe Attachments, anti-malware, anti-spam
(entrant/sortant), filtre de connexion, domaines distants (transfert automatique externe), stratégies
DLP et récapitulatif des stratégies d'alerte. En option (`-IncludeMailboxDetail`), vérifie aussi
l'activation de POP/IMAP et la litigation hold par boîte aux lettres.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-IncludeMailboxDetail` | Non | Vérifier aussi, par boîte aux lettres, les protocoles hérités + la litigation hold (plus lent) |
| `-OutputPath` | Non | Chemin du rapport CSV (par défaut : `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Non | ID de tenant ou domaine Entra ID |

**Exemples**

```powershell
.\Test-EmailSecurityPosture.ps1

.\Test-EmailSecurityPosture.ps1 -IncludeMailboxDetail
```

**Remarques**
- C'est la consolidation phare d'environ 18 scripts d'origine à usage unique — voir la section
  `.NOTES` du script pour la liste complète
- Les équivalents `*-set.ps1` / `*-del.ps1` du projet source, qui modifient la configuration, n'ont
  volontairement pas été portés — chacun codait en dur les valeurs « recommandées » d'un MSP particulier, sans
  possibilité de les adapter par tenant

**Modules requis**
```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

---

### Test-MailboxAuditingConfig.ps1

Signale (et, avec `-Apply`, corrige) les lacunes de l'Unified Audit Log et de la journalisation d'audit par boîte aux lettres :
l'Unified Audit Log est-il activé pour toute l'organisation, et chaque boîte aux lettres a-t-elle
`AuditEnabled` activé avec une `AuditLogAgeLimit` suffisante.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-MinimumAuditLogAgeDays` | Non | Durée de conservation minimale acceptable, en jours (par défaut : `180`) |
| `-Apply` | Non | Activer l'Unified Audit Log et corriger les boîtes aux lettres signalées (par défaut : rapport uniquement) |
| `-OutputPath` | Non | Chemin du rapport CSV (par défaut : `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Non | ID de tenant ou domaine Entra ID |

**Exemples**

```powershell
.\Test-MailboxAuditingConfig.ps1

.\Test-MailboxAuditingConfig.ps1 -Apply -WhatIf

.\Test-MailboxAuditingConfig.ps1 -MinimumAuditLogAgeDays 365 -Apply
```

Prend en charge `-WhatIf` (`SupportsShouldProcess`).

---

### Test-EmailAuthenticationRecords.ps1

Valide les enregistrements DNS SPF et DMARC d'un ou plusieurs domaines — vérifie que SPF existe, inclut
`spf.protection.outlook.com` là où c'est attendu et n'est pas permissif (`+all`) ; vérifie que DMARC
existe, indique sa stratégie (`none`/`quarantine`/`reject`) et si une adresse de rapports agrégés
est configurée. DKIM est volontairement hors périmètre — utilisez
[`Test-DkimConfig.ps1`](../../Exchange/readme.fr.md) pour cela.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-Domain` | Non | Un ou plusieurs domaines. Si omis, ils sont découverts automatiquement via Microsoft Graph |
| `-OutputPath` | Non | Chemin du rapport CSV (par défaut : `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Non | Utilisé uniquement pour la découverte automatique des domaines via Graph |

**Exemples**

```powershell
.\Test-EmailAuthenticationRecords.ps1 -Domain "contoso.com"

.\Test-EmailAuthenticationRecords.ps1
```

**Remarques**
- Windows uniquement (utilise `Resolve-DnsName`)
