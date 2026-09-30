[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [LegacyUtilities](../readme.fr.md) › **Device**

# Legacy Utilities — Device

Petits utilitaires de configuration des postes de travail.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Set-NumLockDefault.ps1`](Set-NumLockDefault.ps1) ([docs](#set-numlockdefaultps1)) | Définir l'état par défaut du Verr Num pour les nouveaux profils et l'écran de connexion |
| [`New-LockWorkstationShortcut.ps1`](New-LockWorkstationShortcut.ps1) ([docs](#new-lockworkstationshortcutps1)) | Créer un raccourci de verrouillage de la station de travail |

---

### Set-NumLockDefault.ps1

Écrit `InitialKeyboardIndicators` sous `HKU\.DEFAULT` (le modèle sur lequel reposent les
nouveaux profils, également utilisé sur l'écran de connexion) ainsi que dans la ruche propre à
l'utilisateur actuel. Essai à blanc par défaut.

```powershell
.\Set-NumLockDefault.ps1 -State On -Apply
```

---

### New-LockWorkstationShortcut.ps1

Crée un raccourci `.lnk` qui exécute la commande standard
`rundll32.exe user32.dll,LockWorkStation` : aucun téléchargement externe, aucune astuce de
registre pour l'épingler. Remplacement généralisé d'un script qui téléchargeait une icône et un
fichier batch personnalisés depuis un point de terminaison interne et les épinglait à la barre
des tâches via une clé de registre non documentée ; cette version se contente de créer le
raccourci et laisse l'utilisateur l'épingler lui-même s'il le souhaite. Essai à blanc par défaut.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-TargetFolder` | Non | Emplacement de création (par défaut : Public Desktop) |
| `-ShortcutName` | Non | Par défaut : "Lock Workstation" |
| `-IconPath` | Non | Fichier `.ico` local facultatif |
| `-Apply` | Non | Créer réellement le raccourci (par défaut : aperçu) |

```powershell
.\New-LockWorkstationShortcut.ps1 -Apply
```
