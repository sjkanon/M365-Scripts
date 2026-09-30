[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [Office365Toolkit](../readme.fr.md) › **Security**

# Office365Toolkit / Security

Scripts de rapport et de durcissement de la posture de sécurité : tendance du Secure Score, revue des consentements
d'applications d'entreprise, verrouillage de la connexion aux boîtes aux lettres partagées et base de référence EOP anti-spam/
anti-malware.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Get-SecureScoreReport.ps1`](Get-SecureScoreReport.ps1) ([docs](#get-securescorereportps1)) | Rapport sur la tendance du Secure Score et les contrôles les plus faibles |
| [`Remove-EnterpriseAppConsent.ps1`](Remove-EnterpriseAppConsent.ps1) ([docs](#remove-enterpriseappconsentps1)) | Auditer et éventuellement révoquer les consentements OAuth d'une application d'entreprise |
| [`Test-SharedMailboxSignIn.ps1`](Test-SharedMailboxSignIn.ps1) ([docs](#test-sharedmailboxsigninps1)) | Signaler/désactiver la connexion directe aux boîtes aux lettres partagées |
| [`New-EOPProtectionBaseline.ps1`](New-EOPProtectionBaseline.ps1) ([docs](#new-eopprotectionbaselineps1)) | Créer/mettre à jour des stratégies EOP de référence anti-spam + anti-malware |

---

### Get-SecureScoreReport.ps1

Rapporte l'évolution du Microsoft Secure Score du tenant dans le temps et détaille les
scores par contrôle du dernier instantané, triés du plus faible au plus fort. Lecture seule.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-HistoryCount` | Non | Nombre d'instantanés historiques à inclure (par défaut `30`) |
| `-OutputPath` | Non | Chemin du rapport CSV (par défaut : `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Non | ID de tenant ou domaine Entra ID |

**Exemples**

```powershell
.\Get-SecureScoreReport.ps1

.\Get-SecureScoreReport.ps1 -HistoryCount 90 -OutputPath C:\Reports
```

**Étendue requise :** `SecurityEvents.Read.All`
**Module requis :** `Microsoft.Graph.Security`

---

### Remove-EnterpriseAppConsent.ps1

Audite (et, avec `-Apply`, révoque) les autorisations déléguées et d'application
accordées à une seule application d'entreprise — utile pour examiner des consentements
illicites ou faire le ménage avant de supprimer une application. Exactement un des paramètres `-AppId` /
`-AppDisplayName` est obligatoire, afin qu'une exécution ne puisse pas balayer par accident toutes
les applications du tenant.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-AppId` | * | ID d'application (client) ou ID d'objet du principal de service |
| `-AppDisplayName` | * | Nom d'affichage exact de l'application d'entreprise |
| `-IncludeUserConsent` | Non | Signaler/révoquer aussi le consentement délégué par utilisateur (pas uniquement le consentement administrateur à l'échelle du tenant) |
| `-Apply` | Non | Révoquer réellement les autorisations signalées (par défaut : aperçu uniquement) |
| `-OutputPath` | Non | Chemin du rapport CSV |
| `-TenantId` | Non | ID de tenant ou domaine Entra ID |

*Exactement un des paramètres `-AppId` / `-AppDisplayName` est obligatoire.

**Exemples**

```powershell
# Aperçu uniquement
.\Remove-EnterpriseAppConsent.ps1 -AppDisplayName "Suspicious Reporting Tool"

# Révoquer le consentement administrateur à l'échelle du tenant + les autorisations d'application
.\Remove-EnterpriseAppConsent.ps1 -AppDisplayName "Suspicious Reporting Tool" -Apply

# Révoquer aussi le consentement délégué des utilisateurs individuels
.\Remove-EnterpriseAppConsent.ps1 -AppId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -IncludeUserConsent -Apply
```

Prend en charge `-WhatIf` (`SupportsShouldProcess`).

**Étendues requises :** `Application.Read.All`, `DelegatedPermissionGrant.ReadWrite.All`, `AppRoleAssignment.ReadWrite.All`
**Module requis :** `Microsoft.Graph.Applications`

---

### Test-SharedMailboxSignIn.ps1

Croise les boîtes aux lettres partagées Exchange Online avec l'état de leur compte Entra ID
et signale celles qui autorisent encore la connexion interactive directe — une cible facile
fréquente, car les boîtes aux lettres partagées sont rarement licenciées ou protégées par MFA. Avec
`-Apply`, désactive la connexion (`AccountEnabled = $false`) pour chaque compte trouvé activé.
L'accès délégué (Full Access / Send As) n'est pas affecté.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-Mailbox` | Non | UPN d'une seule boîte aux lettres partagée. Si omis, toutes les boîtes aux lettres partagées sont vérifiées |
| `-Apply` | Non | Désactiver réellement la connexion pour les comptes activés trouvés (par défaut : aperçu uniquement) |
| `-OutputPath` | Non | Chemin du rapport CSV |
| `-TenantId` | Non | ID de tenant ou domaine Entra ID |

**Exemples**

```powershell
.\Test-SharedMailboxSignIn.ps1

.\Test-SharedMailboxSignIn.ps1 -Apply

.\Test-SharedMailboxSignIn.ps1 -Mailbox "helpdesk@contoso.com" -Apply
```

Prend en charge `-WhatIf` (`SupportsShouldProcess`).

**Étendue requise :** `User.ReadWrite.All`
**Modules requis :** `ExchangeOnlineManagement`, `Microsoft.Graph.Users`

---

### New-EOPProtectionBaseline.ps1

Compare les stratégies anti-spam/anti-malware actuelles du tenant à une base de référence
recommandée et, avec `-Apply`, crée (ou, avec `-UpdateExisting`, met à jour)
une stratégie de référence de filtrage du contenu hébergé (anti-spam) et/ou une stratégie de filtrage
des programmes malveillants + règle, limitées aux domaines destinataires indiqués (par défaut, tous les
domaines acceptés).

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-Domains` | Non | Domaine(s) destinataire(s) auxquels limiter la ou les règles (par défaut : tous les domaines acceptés) |
| `-Protection` | Non | `Spam`, `Malware` ou `Both` (par défaut) |
| `-SpamPolicyName` | Non | Nom de la stratégie/règle anti-spam (par défaut `MSP Baseline Anti-Spam`) |
| `-MalwarePolicyName` | Non | Nom de la stratégie/règle anti-malware (par défaut `MSP Baseline Anti-Malware`) |
| `-UpdateExisting` | Non | Mettre à jour la stratégie sur place si une stratégie du même nom existe déjà |
| `-Apply` | Non | Créer/mettre à jour réellement les stratégies (par défaut : aperçu uniquement) |
| `-TenantId` | Non | ID de tenant ou domaine Entra ID |

**Exemples**

```powershell
# Aperçu uniquement
.\New-EOPProtectionBaseline.ps1

# Créer les deux stratégies de référence pour tous les domaines acceptés
.\New-EOPProtectionBaseline.ps1 -Apply

# Anti-spam uniquement, domaines précis, mise à jour sur place si elle existe
.\New-EOPProtectionBaseline.ps1 -Protection Spam -Domains "contoso.com" -UpdateExisting -Apply
```

**Remarques**
- Il s'agit de recommandations de base, pas d'un durcissement complet — comparez-les
  aux stratégies de sécurité prédéfinies Standard/Strict de votre tenant avant de les
  appliquer à grande échelle.

Prend en charge `-WhatIf` (`SupportsShouldProcess`).

**Module requis :** `ExchangeOnlineManagement`
