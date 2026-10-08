[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [Office365Toolkit](../readme.fr.md) › **Exchange**

# Office365Toolkit / Exchange

Base de référence d'hygiène des boîtes aux lettres, risque de transfert via les règles de boîte de réception,
compléments de boîte aux lettres et recherche dans l'Unified Audit Log.

**Connexion.** Chaque script se connecte via
[`Connect-M365.ps1`](../../Startup/readme.fr.md#connect-m365ps1) : délégué par défaut (vous vous
connectez en tant qu'administrateur), avec le code d'appareil et le client GDAP issus de
`load.config.ps1` — sous GDAP, Exchange Online est atteint avec `-DelegatedOrganization`. App-only
avec `-ClientId` + `-CertificateThumbprint`, ou `-AppOnly` (ClientId et empreinte issus de
`graph.appid.json`). Une session déjà adaptée est réutilisée et reste connectée ; les scripts ne
déconnectent que ce qu'ils ont ouvert eux-mêmes. `Search-MailboxAuditLog.ps1` utilise Microsoft
Graph ; les trois autres restent sur Exchange Online, car Graph n'a pas d'API pour ce qu'ils lisent
(voir les Remarques de chaque script).

> Ces scripts complètent, sans les dupliquer, les scripts d'audit Exchange existants dans
> [`scripts/Exchange/`](../../Exchange/readme.fr.md) — voir la section Remarques de chaque script
> ci-dessous pour la délimitation exacte.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Test-MailboxSecurityBaseline.ps1`](Test-MailboxSecurityBaseline.ps1) ([docs](#test-mailboxsecuritybaselineps1)) | Auditer les boîtes aux lettres par rapport à une base de référence d'hygiène (journalisation d'audit, rétention, litigation hold, archive, protocoles hérités) |
| [`Test-MailboxForwardingRisk.ps1`](Test-MailboxForwardingRisk.ps1) ([docs](#test-mailboxforwardingriskps1)) | Auditer les règles de boîte de réception et les règles Sweep à la recherche de schémas de transfert/exfiltration (indicateur de BEC) |
| [`Get-MailboxAddIns.ps1`](Get-MailboxAddIns.ps1) ([docs](#get-mailboxaddinsps1)) | Rapport des compléments Outlook installés par boîte aux lettres |
| [`Search-MailboxAuditLog.ps1`](Search-MailboxAuditLog.ps1) ([docs](#search-mailboxauditlogps1)) | Rechercher dans l'Unified Audit Log (API Audit Log Query de Graph) les événements de connexion et de connexion aux boîtes aux lettres |

> Les rapports de suivi des messages se trouvent dans [`scripts/PatronToolkit/Exchange/Get-MessageTraceReport.ps1`](../../PatronToolkit/Exchange/readme.fr.md#get-messagetracereportps1) — un script équivalent a été développé indépendamment pour les deux boîtes à outils, un seul a donc été conservé (avec la prise en charge de `-IncludeDetail` reprise de celui-ci).

---

### Test-MailboxSecurityBaseline.ps1

Vérifie chaque boîte aux lettres (ou une seule) par rapport à une base de référence d'hygiène : journalisation
d'audit activée + durée de conservation du journal, rétention des éléments supprimés, litigation hold,
état de l'archive, absence de transfert au niveau de la boîte aux lettres, POP3/IMAP désactivés. Chaque boîte aux lettres
obtient un Pass/Fail par contrôle et un `Status` global. Lecture seule.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-Mailbox` | Non | UPN d'une seule boîte aux lettres. Si omis, toutes les boîtes aux lettres utilisateur et partagées sont vérifiées |
| `-MinAuditLogAgeLimitDays` | Non | Durée minimale acceptable de conservation du journal d'audit, en jours (par défaut `90`) |
| `-MinRetainDeletedItemsDays` | Non | Durée minimale acceptable de rétention des éléments supprimés, en jours (par défaut `30`) |
| `-OutputPath` | Non | Chemin du rapport CSV (par défaut : `C:\Temp\` / `~/Downloads\`) |
| `-TenantId` | Non | ID de tenant ou domaine ; par défaut le client GDAP de `load.config.ps1` (l'app-only exige la forme domaine) |
| `-ClientId` | Non | Inscription d'application pour la connexion app-only (avec `-CertificateThumbprint`) |
| `-CertificateThumbprint` | Non | Empreinte du certificat pour la connexion app-only |
| `-AppOnly` | Non | App-only avec ClientId et empreinte issus de `graph.appid.json` |

**Exemples**

```powershell
.\Test-MailboxSecurityBaseline.ps1

.\Test-MailboxSecurityBaseline.ps1 -Mailbox "user@contoso.com"

.\Test-MailboxSecurityBaseline.ps1 -MinAuditLogAgeLimitDays 180 -MinRetainDeletedItemsDays 30
```

**Remarques**
- Reste sur Exchange Online : les paramètres d'audit, de rétention, de conservation, d'archive, de
  transfert et POP/IMAP (`Get-Mailbox` / `Get-CASMailbox`) n'ont pas d'API Graph.
- Le transfert externe au niveau de la boîte aux lettres est ici un simple contrôle présent/absent ; pour
  une ventilation externe/interne tenant compte des domaines, utilisez
  [`Get-ExternalForwards.ps1`](../../Exchange/readme.fr.md#get-externalforwardsps1).
- Pour le transfert via les règles de boîte de réception/règles Sweep, utilisez plutôt `Test-MailboxForwardingRisk.ps1`
  ci-dessous.

**Module requis :** `ExchangeOnlineManagement`

---

### Test-MailboxForwardingRisk.ps1

Audite les règles de boîte de réception et les règles Sweep de chaque boîte aux lettres à la recherche de schémas de transfert/redirection/
exfiltration — un indicateur classique de compromission de messagerie professionnelle (BEC)
qui n'apparaît pas sur l'objet boîte aux lettres lui-même. Classe les destinataires des règles en
External/Internal/Unknown en fonction des domaines acceptés du tenant. Lecture seule.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-Mailbox` | Non | UPN d'une seule boîte aux lettres. Si omis, toutes les boîtes aux lettres sont vérifiées |
| `-IncludeDisabledRules` | Non | Signaler aussi les règles désactivées qui correspondent aux schémas à risque |
| `-OutputPath` | Non | Chemin du rapport CSV |
| `-TenantId` | Non | ID de tenant ou domaine ; par défaut le client GDAP de `load.config.ps1` (l'app-only exige la forme domaine) |
| `-ClientId` | Non | Inscription d'application pour la connexion app-only (avec `-CertificateThumbprint`) |
| `-CertificateThumbprint` | Non | Empreinte du certificat pour la connexion app-only |
| `-AppOnly` | Non | App-only avec ClientId et empreinte issus de `graph.appid.json` |

**Exemples**

```powershell
.\Test-MailboxForwardingRisk.ps1

.\Test-MailboxForwardingRisk.ps1 -Mailbox "user@contoso.com" -IncludeDisabledRules
```

**Remarques**
- Reste sur Exchange Online : Graph ne lit les règles de boîte de réception (`messageRules`) d'un
  autre utilisateur qu'avec une autorisation Mail app-only, et n'a pas d'API pour les règles Sweep.
- Un destinataire interne à l'organisation (`[EX:/o=...]`) compte désormais comme Internal ; il
  était signalé External faute de domaine à comparer.
- Complète [`Get-ExternalForwards.ps1`](../../Exchange/readme.fr.md#get-externalforwardsps1)
  (`ForwardingSmtpAddress` au niveau de la boîte aux lettres) et `Test-MailboxSecurityBaseline.ps1`
  ci-dessus (même contrôle) — ce script couvre uniquement la couche des règles de boîte de réception/règles Sweep.

**Module requis :** `ExchangeOnlineManagement`

---

### Get-MailboxAddIns.ps1

Liste les compléments Outlook (`Get-App`) présents sur chaque boîte aux lettres — déployés de manière centralisée
ou installés par l'utilisateur/chargés à part (sideload). Utile pour repérer les compléments non approuvés, un vecteur
connu d'hameçonnage et d'octroi de consentement. Lecture seule.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-Mailbox` | Non | UPN d'une seule boîte aux lettres. Si omis, toutes les boîtes aux lettres utilisateur et partagées sont vérifiées |
| `-OutputPath` | Non | Chemin du rapport CSV |
| `-TenantId` | Non | ID de tenant ou domaine ; par défaut le client GDAP de `load.config.ps1` (l'app-only exige la forme domaine) |
| `-ClientId` | Non | Inscription d'application pour la connexion app-only (avec `-CertificateThumbprint`) |
| `-CertificateThumbprint` | Non | Empreinte du certificat pour la connexion app-only |
| `-AppOnly` | Non | App-only avec ClientId et empreinte issus de `graph.appid.json` |

**Exemples**

```powershell
.\Get-MailboxAddIns.ps1

.\Get-MailboxAddIns.ps1 -Mailbox "user@contoso.com"
```

**Remarques**
- Reste sur Exchange Online : les compléments Outlook par boîte aux lettres (`Get-App`) n'ont pas
  d'API Graph.

**Module requis :** `ExchangeOnlineManagement`

---

### Search-MailboxAuditLog.ps1

Recherche générique dans l'Unified Audit Log. Par défaut, le script utilise l'API Audit Log Query
de Microsoft Graph : il crée une requête (`POST /security/auditLog/queries`), vérifie toutes les
30 secondes si elle est terminée, puis parcourt les enregistrements page par page. Par défaut,
couvre les 2 derniers jours pour les connexions interactives (réussies/échouées) et les connexions
aux boîtes aux lettres. Entièrement paramétrable pour d'autres types d'enregistrements/opérations/
plages de dates/utilisateurs. `-UseExchange` exécute la même recherche avec
`Search-UnifiedAuditLog` dans Exchange Online. Lecture seule.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-Days` | Non | Nombre de jours de recherche en arrière à partir de maintenant (par défaut `2`) ; ignoré si `-StartDate` est fourni |
| `-StartDate` | Non | Début explicite de la fenêtre (prioritaire sur `-Days`) |
| `-EndDate` | Non | Fin explicite de la fenêtre (par défaut : maintenant) |
| `-RecordType` | Non | Type(s) d'enregistrement du journal d'audit (par défaut `AzureActiveDirectoryStsLogon`, `ExchangeItem`) |
| `-Operations` | Non | Nom(s) d'opération (par défaut `UserLoggedIn`, `UserLoginFailed`, `MailboxLogin`) |
| `-UserIds` | Non | Limiter à un ou plusieurs utilisateurs précis (UPN) |
| `-UseExchange` | Non | Rechercher avec `Search-UnifiedAuditLog` dans Exchange Online au lieu de Graph |
| `-TimeoutMinutes` | Non | Durée d'attente maximale de la requête Graph (par défaut `60`) |
| `-OutputPath` | Non | Chemin du rapport CSV |
| `-TenantId` | Non | ID de tenant ou domaine ; par défaut le client GDAP de `load.config.ps1` |
| `-ClientId` | Non | Inscription d'application pour la connexion app-only (avec `-CertificateThumbprint`) |
| `-CertificateThumbprint` | Non | Empreinte du certificat pour la connexion app-only |
| `-AppOnly` | Non | App-only avec ClientId et empreinte issus de `graph.appid.json` |

**Exemples**

```powershell
# 2 derniers jours, connexions + connexions aux boîtes aux lettres
.\Search-MailboxAuditLog.ps1

# 30 derniers jours, connexions échouées uniquement, un seul utilisateur
.\Search-MailboxAuditLog.ps1 -Days 30 -RecordType AzureActiveDirectoryStsLogon -Operations UserLoginFailed -UserIds "user@contoso.com"

# Fenêtre explicite
.\Search-MailboxAuditLog.ps1 -StartDate (Get-Date "2026-07-01") -EndDate (Get-Date "2026-07-15")

# Même recherche via Exchange Online
.\Search-MailboxAuditLog.ps1 -UseExchange
```

**Remarques**
- La requête Graph s'exécute de façon asynchrone dans le service et peut prendre plusieurs
  minutes. Après `-TimeoutMinutes`, le script cesse d'attendre et affiche l'ID de la requête ;
  celle-ci continue et peut être lue plus tard.
- Les types d'enregistrement s'indiquent sous leur forme du journal d'audit (`ExchangeItem`) ; le
  script les convertit en camelCase pour l'API (`exchangeItem`).
- `Search-UnifiedAuditLog` n'accepte qu'un type d'enregistrement par appel. Les versions
  précédentes passaient les deux types par défaut en un seul appel ; `-UseExchange` exécute
  désormais une recherche paginée par type d'enregistrement.
- L'Unified Audit Log n'est pas immédiat — comptez jusqu'à 30-60 minutes avant que l'activité
  récente n'apparaisse.

**Étendue requise :** `AuditLogsQuery.Read.All` (plus un rôle Purview autorisé à rechercher dans le journal d'audit, p. ex. Audit Logs)
**Modules requis :** `Microsoft.Graph.Authentication` ; `ExchangeOnlineManagement` pour `-UseExchange`
