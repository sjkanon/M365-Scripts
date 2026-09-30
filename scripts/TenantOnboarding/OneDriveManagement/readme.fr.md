[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [TenantOnboarding](../readme.fr.md) › **OneDriveManagement**

# OneDriveManagement

Scripts de maintenance de OneDrive for Business côté appareil : une surveillance de redémarrage/réinitialisation, l'arrêt de la synchronisation par bibliothèque et la redirection Known Folder Move. À distinguer de [`scripts/Device/DriveMapping/`](../../Device/DriveMapping/readme.fr.md) (mappage de SharePoint/OneDrive sur une lettre de lecteur via WebDAV), qui est une fonctionnalité différente.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Register-OneDriveWatchdog.ps1`](Register-OneDriveWatchdog.ps1) ([docs](#register-onedrivewatchdogps1)) | Maintenir OneDrive en cours d'exécution via une tâche planifiée ; prend aussi en charge un `/reset` ponctuel |
| [`Stop-OneDriveLibrarySync.ps1`](Stop-OneDriveLibrarySync.ps1) ([docs](#stop-onedrivelibrarysyncps1)) | Arrêter la synchronisation d'une bibliothèque précise sans toucher aux autres |
| [`Set-OneDriveKnownFolderRedirect.ps1`](Set-OneDriveKnownFolderRedirect.ps1) ([docs](#set-onedriveknownfolderredirectps1)) | Rediriger Bureau/Documents/Images/Téléchargements vers OneDrive |

---

### Register-OneDriveWatchdog.ps1

Enregistre une tâche planifiée (exécutée sous l'utilisateur actuel, toutes les 59 minutes par défaut) qui relance OneDrive s'il n'est pas en cours d'exécution, ou le redémarre de force à chaque exécution en option. `-Reset` effectue à la place un `onedrive.exe /reset` ponctuel pour un profil de synchronisation bloqué.

| Paramètre | Description |
|-----------|-------------|
| `-IntervalMinutes` | Intervalle de surveillance (par défaut : `59`) |
| `-ForceRestartOnEachRun` | Arrêter et relancer OneDrive à chaque déclenchement, pas seulement lorsqu'il ne tourne pas |
| `-Reset` | Exécuter un `/reset` ponctuel au lieu d'enregistrer la surveillance |
| `-Apply` | Enregistrer/réinitialiser réellement (par défaut : aperçu uniquement) |

```powershell
.\Register-OneDriveWatchdog.ps1 -Apply
.\Register-OneDriveWatchdog.ps1 -ForceRestartOnEachRun -IntervalMinutes 30 -Apply
.\Register-OneDriveWatchdog.ps1 -Reset -Apply
```

---

### Stop-OneDriveLibrarySync.ps1

Arrête proprement la synchronisation d'une bibliothèque SharePoint/OneDrive précise : ferme OneDrive, supprime le fichier de stratégie et l'entrée du cache de point de montage de cette bibliothèque, modifie l'ini de la base de synchronisation, supprime le dossier local et relance OneDrive, le tout sans perturber les autres bibliothèques synchronisées.

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-MatchPattern` | Oui | Chaîne distinctive identifiant le fichier de stratégie de la bibliothèque |
| `-LocalFolderPath` | Oui | Dossier local synchronisé à supprimer |
| `-Apply` | Non | Effectuer réellement l'arrêt (par défaut : aperçu uniquement) |

```powershell
.\Stop-OneDriveLibrarySync.ps1 -MatchPattern "ProjectX" -LocalFolderPath "$env:USERPROFILE\Contoso\ProjectX - Documents" -Apply
```

> Agit sur des fichiers d'état internes non documentés de OneDrive. Testez sur une seule machine avant un déploiement plus large.

---

### Set-OneDriveKnownFolderRedirect.ps1

Redirige les dossiers connus vers la racine de synchronisation de OneDrive for Business à l'aide de `SHSetKnownFolderPath`, migre le contenu existant avec Robocopy et masque (sans le supprimer) le dossier d'origine. Peut en option activer l'indicateur OneDrive `Timerautomount` pour que les bibliothèques déjà synchronisées se montent automatiquement sur un nouvel appareil.

| Paramètre | Description |
|-----------|-------------|
| `-Folders` | Dossiers à rediriger (par défaut : Desktop, Documents, Pictures, Downloads) |
| `-EnableAutoMountSharedLibraries` | Activer également `Timerautomount` |
| `-Apply` | Rediriger réellement (par défaut : aperçu uniquement) |

```powershell
.\Set-OneDriveKnownFolderRedirect.ps1 -Apply
.\Set-OneDriveKnownFolderRedirect.ps1 -Folders Desktop,Documents -EnableAutoMountSharedLibraries -Apply
```

> À exécuter dans le contexte de l'utilisateur connecté, et non en tant que SYSTEM : le script lit la clé de registre du compte OneDrive de cet utilisateur.

---

Tous les scripts effectuent un essai à blanc par défaut ; passez `-Apply` pour appliquer les modifications, conformément au style maison.
