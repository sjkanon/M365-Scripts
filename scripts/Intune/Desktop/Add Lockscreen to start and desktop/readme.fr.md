[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../../readme.fr.md) › [scripts](../../../readme.fr.md) › [Intune](../../readme.fr.md) › [Desktop](../readme.fr.md) › **Add Lockscreen to start and desktop**

# Ajouter le verrouillage au menu Démarrer et au bureau

Épingle un raccourci « Lock Workstation » au menu Démarrer / au bureau à l'aide d'un `.bat` + `.ico` téléchargés, en utilisant le verbe Explorer non documenté `Windows.taskbarpin` pour épingler sans intervention de l'utilisateur.

---

## Fichiers

| Fichier | Description |
|------|-------------|
| [`add-lock.ps1`](add-lock.ps1) ([docs](#add-lockps1)) | Télécharge le script de verrouillage + l'icône et crée le raccourci |
| [`add-shortcut-lock.ps1`](add-shortcut-lock.ps1) ([docs](#add-shortcut-lockps1)) | Épingle un raccourci quelconque via le verbe Explorer `Windows.taskbarpin` |

---

### add-lock.ps1

Crée `C:\Program Files\EOO\lockworkstation\` (ignoré s'il existe déjà), télécharge `lock.txt` + `lock.ico` depuis `https://endpoint.eoo.cloud/lock/`, renomme `lock.txt` en `lock.bat` et crée un raccourci du menu Démarrer (`lock.lnk`) sous `C:\ProgramData\...\Start Menu\Programs\EOO\lockworkstation\` qui lance `explorer.exe` avec le `.bat` en argument.

```powershell
.\add-lock.ps1
```

> Aucun paramètre. Prévu pour s'exécuter une seule fois par appareil (Intune, contexte SYSTEM) — une réexécution ne fait plus rien dès que le dossier `EOO` existe.

---

### add-shortcut-lock.ps1

Épingle un fichier cible à la barre des tâches/au menu Démarrer à l'aide du verbe de registre `Windows.taskbarpin` `ExplorerCommandHandler`, puis supprime les clés de registre temporaires qu'il a créées.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-Target` | Oui | Chemin du fichier/raccourci à épingler |

```powershell
.\add-shortcut-lock.ps1 -Target "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\EOO\lockworkstation\lock.lnk"
```

> Dépend du raccourci `lock.lnk` créé par `add-lock.ps1`.
