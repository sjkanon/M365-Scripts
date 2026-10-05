[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [Intune](../readme.fr.md) › **iOS-Compliance-Updater**

# Intune — iOS Compliance Updater

Maintient automatiquement à jour l'exigence de version iOS minimale d'une stratégie de conformité Intune via l'API Microsoft Graph. S'exécute comme tâche planifiée sur un serveur Windows — aucune intervention manuelle nécessaire.

**Fonctionnement**

1. Récupère la dernière version d'iOS depuis le flux RSS d'Apple (avec repli sur la page Apple Support)
2. La compare à la version minimale actuelle de la stratégie de conformité Intune
3. Met à jour la stratégie si une version plus récente est disponible
4. Journalise toutes les actions dans un fichier journal mensuel

---

## Fichiers

| Fichier | Description |
|------|-------------|
| [`Update-iOSCompliancePolicy.ps1`](Update-iOSCompliancePolicy.ps1) ([docs](#utilisation)) | Script principal — à exécuter manuellement ou via une tâche planifiée |
| [`Setup.ps1`](Setup.ps1) ([docs](#option-a--automatique-recommandée)) | Configuration initiale unique — crée l'App Registration et écrit config.json |
| [`Install-ScheduledTask.ps1`](Install-ScheduledTask.ps1) ([docs](#enregistrer-la-tâche-planifiée)) | Enregistre la tâche planifiée Windows |
| `config.example.json` | Exemple de fichier de configuration |

---

## Configuration initiale

### Option A — Automatique (recommandée)

Exécutez `Setup.ps1` une seule fois. Il s'occupe de tout :

```powershell
.\Setup.ps1
```

Le script va :
- Installer les modules PowerShell requis
- Ouvrir un navigateur pour l'authentification
- Créer l'App Registration dans Entra ID
- Attribuer l'autorisation d'API requise et accorder le consentement administrateur
- Créer un Client Secret
- Afficher les stratégies de conformité iOS disponibles pour que vous en choisissiez une
- Écrire `config.json` automatiquement

```powershell
# Indiquer directement le nom de la stratégie pour sauter l'invite de sélection
.\Setup.ps1 -CompliancePolicyName "iOS - Minimum version compliance"
```

> Rôle requis : **Global Administrator** ou **Application Administrator + Intune Administrator**

---

### Option B — Manuelle

**Étape 1 — Créer l'App Registration**

1. Entra ID → **App registrations** → **New registration**
2. Nom : `Intune iOS Compliance Updater`
3. Après la création → **API permissions** → **Add a permission**
4. Choisissez **Microsoft Graph** → **Application permissions**
5. Ajoutez : `DeviceManagementConfiguration.ReadWrite.All`
6. Cliquez sur **Grant admin consent**
7. **Certificates & secrets** → **New client secret** — notez la valeur

**Étape 2 — Trouver l'ID de la stratégie de conformité**

1. Intune Admin Center → **Devices** → **Compliance** → ouvrez la stratégie iOS
2. Copiez le GUID depuis l'URL : `.../deviceCompliancePolicies/xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`

**Étape 3 — Créer config.json**

Copiez `config.example.json` vers `config.json` et renseignez les valeurs :

```json
{
    "TenantId":           "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
    "ClientId":           "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
    "ClientSecret":       "your-client-secret",
    "CompliancePolicyId": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
}
```

> Ne commitez jamais `config.json` dans Git — il contient des secrets. Il figure dans `.gitignore`.

---

### Enregistrer la tâche planifiée

Après la configuration (quelle que soit l'option), enregistrez la tâche :

```powershell
# Exécuter en tant qu'administrateur
.\Install-ScheduledTask.ps1
```

La tâche s'exécute chaque **lundi à 07:00** en tant que SYSTEM.

---

## Utilisation

```powershell
# Exécution manuelle
.\Update-iOSCompliancePolicy.ps1

# Essai à blanc — aucune modification effectuée
.\Update-iOSCompliancePolicy.ps1 -WhatIf

# Chemin de configuration ou de journal personnalisé
.\Update-iOSCompliancePolicy.ps1 -ConfigPath "D:\configs\intune.json" -LogPath "D:\logs"
```

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

- Ne commitez jamais `config.json` dans Git (déjà présent dans `.gitignore`)
- Pour les environnements de production, envisagez de stocker le Client Secret dans le **Gestionnaire d'identification Windows** (Windows Credential Manager) ou dans **Azure Key Vault**
- L'App Registration n'utilise que l'autorisation minimale requise : `DeviceManagementConfiguration.ReadWrite.All`
