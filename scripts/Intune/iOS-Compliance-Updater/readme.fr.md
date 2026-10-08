[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [Intune](../readme.fr.md) › **iOS-Compliance-Updater**

# Intune — iOS Compliance Updater

Maintient automatiquement à jour l'exigence de version iOS minimale d'une stratégie de conformité Intune via l'API Microsoft Graph. S'exécute comme tâche planifiée sur un serveur Windows — aucune intervention manuelle nécessaire.

**Fonctionnement**

1. Récupère la dernière version publiée d'iOS depuis le flux RSS d'Apple (avec repli sur la page Apple Support) ; les entrées beta et RC sont ignorées et la version la plus élevée l'emporte (le flux liste aussi les mises à jour des versions majeures plus anciennes)
2. La compare à la version minimale actuelle de la stratégie de conformité Intune
3. Relève la stratégie si une version plus récente est disponible (sans jamais l'abaisser)
4. Journalise toutes les actions dans un fichier journal mensuel

**La connexion** passe par [`Connect-M365.ps1`](../../Startup/Connect-M365.ps1) : gardez donc ce dossier dans la structure du dépôt. Comme l'outil s'exécute sans surveillance, **l'application seule est son mode par défaut** : un certificat (empreinte dans `config.json`, clé privée dans le magasin de certificats Windows). Un `ClientSecret` dans un ancien `config.json` fonctionne toujours. `Setup.ps1` vous connecte en **délégué** pour créer l'application. Tous les scripts exigent PowerShell 7.

---

## Fichiers

| Fichier | Description |
|------|-------------|
| [`Update-iOSCompliancePolicy.ps1`](Update-iOSCompliancePolicy.ps1) ([docs](#utilisation)) | Script principal — à exécuter manuellement ou via une tâche planifiée |
| [`Setup.ps1`](Setup.ps1) ([docs](#option-a--automatique-recommandée)) | Configuration initiale unique — crée l'App Registration et le certificat, écrit config.json |
| [`Install-ScheduledTask.ps1`](Install-ScheduledTask.ps1) ([docs](#enregistrer-la-tâche-planifiée)) | Enregistre la tâche planifiée Windows (pwsh.exe) |
| `config.example.json` | Exemple de fichier de configuration |

---

## Configuration initiale

### Option A — Automatique (recommandée)

Exécutez `Setup.ps1` une seule fois, **en tant qu'administrateur** (pour que le certificat soit placé dans `LocalMachine\My`, où la tâche SYSTEM peut l'utiliser). Il s'occupe de tout :

```powershell
.\Setup.ps1
```

Le script va :
- Installer les modules PowerShell requis
- Vous connecter (délégué, navigateur ou code d'appareil selon `load.config.ps1`)
- Créer (ou réutiliser) l'App Registration dans Entra ID
- Attribuer l'autorisation d'API requise et accorder le consentement administrateur
- Créer un certificat auto-signé et téléverser sa clé publique dans l'application (les clés existantes d'une application réutilisée sont conservées)
- Afficher les stratégies de conformité iOS disponibles pour que vous en choisissiez une (toutes les pages)
- Écrire `config.json` automatiquement

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-TenantId` | Non | Tenant à configurer (par défaut : client GDAP, sinon le tenant de connexion) |
| `-CompliancePolicyName` | Non | Stratégie à utiliser ; sans ce paramètre, vous choisissez dans une liste |
| `-AppName` | Non | Nom de l'App Registration (par défaut : `Intune iOS Compliance Updater`) |
| `-CredentialType` | Non | `Certificate` (par défaut) ou `Secret` |
| `-CertificateStoreLocation` | Non | `LocalMachine` (par défaut en administrateur) ou `CurrentUser` |
| `-CertificateExpiryYears` | Non | Validité du certificat (par défaut : `2`) |
| `-SecretExpiryYears` | Non | Validité du secret avec `-CredentialType Secret` (par défaut : `2`) |
| `-ConfigPath` | Non | Emplacement de `config.json` (par défaut : dossier du script) |

```powershell
# Indiquer directement le tenant et le nom de la stratégie pour sauter l'invite de sélection
.\Setup.ps1 -TenantId "contoso.onmicrosoft.com" -CompliancePolicyName "iOS - Minimum version compliance"

# Ancien comportement : un secret client dans config.json
.\Setup.ps1 -CredentialType Secret
```

> Rôle requis : **Global Administrator**, ou **Application Administrator + Privileged Role Administrator** (consentement administrateur) **+ Intune Administrator**

---

### Option B — Manuelle

**Étape 1 — Créer l'App Registration**

1. Entra ID → **App registrations** → **New registration**
2. Nom : `Intune iOS Compliance Updater`
3. Après la création → **API permissions** → **Add a permission**
4. Choisissez **Microsoft Graph** → **Application permissions**
5. Ajoutez : `DeviceManagementConfiguration.ReadWrite.All`
6. Cliquez sur **Grant admin consent**
7. **Certificates & secrets** → **Certificates** → téléversez le `.cer` d'un certificat dont la clé privée se trouve dans `LocalMachine\My` sur le serveur (ou, moins sûr, créez un secret client)

**Étape 2 — Trouver l'ID de la stratégie de conformité**

1. Intune Admin Center → **Devices** → **Compliance** → ouvrez la stratégie iOS
2. Copiez le GUID depuis l'URL : `.../deviceCompliancePolicies/xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`

**Étape 3 — Créer config.json**

Copiez `config.example.json` vers `config.json` et renseignez les valeurs :

```json
{
    "TenantId":              "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
    "ClientId":              "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
    "CertificateThumbprint": "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "CompliancePolicyId":    "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
}
```

Avec un secret client, utilisez `"ClientSecret": "..."` à la place de `CertificateThumbprint`.

> Ne commitez jamais `config.json` dans Git. Il figure dans `.gitignore`.

---

### Enregistrer la tâche planifiée

Après la configuration (quelle que soit l'option), enregistrez la tâche :

```powershell
# Exécuter en tant qu'administrateur
.\Install-ScheduledTask.ps1
```

La tâche s'exécute chaque **lundi à 07:00** en tant que SYSTEM, avec `pwsh.exe` (PowerShell 7 doit être installé ; la tâche utilisait auparavant `powershell.exe`, qui ne peut pas exécuter l'outil en PowerShell 7). Relancez-le après la mise à jour vers cette version.

---

## Utilisation

```powershell
# Exécution manuelle
.\Update-iOSCompliancePolicy.ps1

# Essai à blanc — aucune modification effectuée
.\Update-iOSCompliancePolicy.ps1 -WhatIf

# Chemin de configuration ou de journal personnalisé
.\Update-iOSCompliancePolicy.ps1 -ConfigPath "D:\configs\intune.json" -LogPath "D:\logs"

# Exécution manuelle déléguée sur une stratégie (pas de ClientId dans config.json)
.\Update-iOSCompliancePolicy.ps1 -CompliancePolicyId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -WhatIf
```

**Paramètres** (`Update-iOSCompliancePolicy.ps1`)

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-ConfigPath` | Non | Fichier de configuration (par défaut : `config.json` à côté du script) |
| `-LogPath` | Non | Dossier des journaux (par défaut : `logs\`) |
| `-CompliancePolicyId` | Non | Remplace `CompliancePolicyId` de `config.json` |
| `-TenantId` / `-ClientId` / `-CertificateThumbprint` | Non | Remplacent les valeurs de `config.json` |
| `-AppOnly` | Non | Application seule avec le ClientId et l'empreinte de `graph.appid.json` |
| `-WhatIf` | Non | Essai à blanc |

**Remarques**
- `-WhatIf` était déclaré à la fois comme paramètre du script et via `SupportsShouldProcess`, si bien que PowerShell refusait de démarrer le script (« A parameter with the name 'WhatIf' was defined multiple times »). C'est désormais uniquement le paramètre commun.
- Le PATCH envoie désormais `@odata.type` (`#microsoft.graph.iosCompliancePolicy`), et le script refuse un ID de stratégie qui n'est pas une stratégie de conformité iOS.
- La recherche de version n'a jamais fonctionné : `Invoke-RestMethod` renvoie directement les éléments RSS, si bien que `$rss.channel.item.title` était toujours vide, et la page de repli (`111900`) ne contenait plus de version. Le script lit désormais les éléments directement (vérifié sur le flux en direct : 27.0.1, alors que le flux liste aussi 26.6.2 et 18.7.10) et se replie sur « À propos des mises à jour d'iOS » (`support.apple.com/100100`, « The latest version of iOS and iPadOS is … »).

---

## Journalisation

Les journaux sont écrits dans `logs\compliance-updater-<yyyy-MM>.log`, à côté du script :

```
[2026-03-24 07:00:01] [INFO   ] Script started
[2026-03-24 07:00:03] [SUCCESS] Latest iOS version: 18.3.2
[2026-03-24 07:00:04] [INFO   ] Policy 'iOS - Minimum version' — current minimum: 18.3.1
[2026-03-24 07:00:05] [SUCCESS] Compliance policy updated to iOS 18.3.2.
```

---

## Sécurité

- Privilégiez le certificat : `config.json` ne contient alors aucun secret et la clé privée n'est pas exportable depuis le magasin de certificats du serveur.
- Un secret client dans `config.json` est en clair : toute personne pouvant lire le fichier peut modifier la configuration Intune. Limitez l'ACL du fichier à SYSTEM et aux administrateurs, ou passez à un certificat en relançant `Setup.ps1`.
- Ne commitez jamais `config.json` dans Git (déjà présent dans `.gitignore`)
- L'App Registration n'utilise que l'autorisation minimale requise : `DeviceManagementConfiguration.ReadWrite.All`
