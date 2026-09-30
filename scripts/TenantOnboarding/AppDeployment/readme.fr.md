[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [TenantOnboarding](../readme.fr.md) › **AppDeployment**

# AppDeployment

Scripts génériques et paramétrés de déploiement d'applications côté appareil, qui remplacent une grande famille d'anciens scripts codant chacun en dur l'URL de téléchargement d'un fournisseur ou les détails d'un raccourci ou d'une imprimante. Faites pointer chaque script vers votre propre dépôt de paquets / URL ; aucun ne code en dur un point de terminaison interne.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Install-Win32AppPackage.ps1`](Install-Win32AppPackage.ps1) ([docs](#install-win32apppackageps1)) | Télécharger un paquet PSAppDeployToolkit zippé et lancer son installation silencieuse |
| [`Install-ChocolateyPackage.ps1`](Install-ChocolateyPackage.ps1) ([docs](#install-chocolateypackageps1)) | Installer, mettre à niveau ou désinstaller un paquet via Chocolatey |
| [`Set-DefaultFileAssociation.ps1`](Set-DefaultFileAssociation.ps1) ([docs](#set-defaultfileassociationps1)) | Définir l'application par défaut d'une extension de fichier (wrapper PS-SFTA) |
| [`New-DesktopShortcutsFromStartMenu.ps1`](New-DesktopShortcutsFromStartMenu.ps1) ([docs](#new-desktopshortcutsfromstartmenups1)) | Copier un ensemble de raccourcis du menu Démarrer vers le Public Desktop |
| [`New-DesktopUrlShortcut.ps1`](New-DesktopUrlShortcut.ps1) ([docs](#new-desktopurlshortcutps1)) | Créer un raccourci `.url` sur le Public Desktop |
| [`Remove-DesktopShortcut.ps1`](Remove-DesktopShortcut.ps1) ([docs](#remove-desktopshortcutps1)) | Supprimer les raccourcis du bureau correspondant à un modèle de nom |
| [`Add-NetworkPrinterConnection.ps1`](Add-NetworkPrinterConnection.ps1) ([docs](#add-networkprinterconnectionps1)) | Ajouter une imprimante réseau via un port IP et un nom de pilote |

---

### Install-Win32AppPackage.ps1

Télécharge un paquet `.zip`, le décompresse, exécute son point d'entrée PSAppDeployToolkit `Deploy-<App>.ps1 -DeploymentType Install -DeployMode NonInteractive`, puis fait le ménage. Peut en option enregistrer une tâche planifiée déclenchée à l'ouverture de session pour revérifier les mises à jour à chaque connexion (pour les clients de bureau à mise à jour automatique).

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-AppName` | Oui | Nom convivial de l'application (dossier de travail + nom par défaut du script de déploiement) |
| `-SourceUri` | Oui | URL du paquet `.zip` |
| `-DeployScriptName` | Non | Nom du script de déploiement dans le paquet (par défaut : `Deploy-<AppName>.ps1`) |
| `-DeploymentType` | Non | Type de déploiement PSADT (par défaut : `Install`) |
| `-DeployMode` | Non | Mode de déploiement PSADT (par défaut : `NonInteractive`) |
| `-WorkingRoot` | Non | Dossier de travail local (par défaut : `C:\Install`) |
| `-RegisterLogonUpdateTask` | Non | Enregistrer également une tâche planifiée à l'ouverture de session qui relance cette installation |
| `-Apply` | Non | Installer réellement (par défaut : aperçu uniquement) |

**Exemples**
```powershell
.\Install-Win32AppPackage.ps1 -AppName "AdobeReaderDC" -SourceUri "https://packages.contoso.com/Installers/AdobeReaderDC.zip" -Apply

.\Install-Win32AppPackage.ps1 -AppName "LineOfBusinessDesktop" `
    -SourceUri "https://packages.contoso.com/Installers/LineOfBusinessDesktop.zip" -RegisterLogonUpdateTask -Apply
```

**Remarques :** remplace toute une famille d'anciens scripts quasi identiques (Adobe Reader, AnyDesk, Citrix Workspace, Google Drive, Jabra Direct, TeamViewer, ainsi qu'un client de bureau métier à mise à jour automatique) qui codaient chacun en dur l'URL d'un fournisseur : le même schéma en quatre étapes, un seul script.

---

### Install-ChocolateyPackage.ps1

Installe Chocolatey s'il est absent, puis installe/met à niveau ou désinstalle le paquet indiqué.

| Paramètre | Description |
|-----------|-------------|
| `-PackageName` | ID du paquet Chocolatey (obligatoire) |
| `-Uninstall` | Désinstaller au lieu d'installer/mettre à niveau |
| `-Apply` | Exécuter réellement (par défaut : aperçu uniquement) |

```powershell
.\Install-ChocolateyPackage.ps1 -PackageName git -Apply
.\Install-ChocolateyPackage.ps1 -PackageName git -Uninstall -Apply
```

---

### Set-DefaultFileAssociation.ps1

Définit une association de fichiers par défaut via l'outil communautaire [PS-SFTA](https://github.com/DanysysTeam/PS-SFTA), qui calcule le hachage exigé par Windows 10 1803+ pour que les modifications de `UserChoice` soient conservées.

| Paramètre | Description |
|-----------|-------------|
| `-ProgId` | ProgId ou `Applications\<exe>` à définir par défaut (obligatoire) |
| `-Extension` | Une ou plusieurs extensions, par ex. `.pdf` (obligatoire) |
| `-Apply` | Modifier réellement (par défaut : aperçu uniquement) |

```powershell
.\Set-DefaultFileAssociation.ps1 -ProgId "Applications\7zFM.exe" -Extension ".zip",".rar" -Apply
.\Set-DefaultFileAssociation.ps1 -ProgId "Acrobat.Document.DC" -Extension ".pdf" -Apply
```

> Télécharge PS-SFTA depuis GitHub à l'exécution. Embarquez-le d'abord localement si votre politique exige des sources de scripts préapprouvées.

---

### New-DesktopShortcutsFromStartMenu.ps1

Copie les raccourcis indiqués du menu Démarrer de tous les utilisateurs vers le Public Desktop.

```powershell
.\New-DesktopShortcutsFromStartMenu.ps1 -AppNames "Excel","Word","Outlook" -Apply
```

---

### New-DesktopUrlShortcut.ps1

Crée sur le bureau un raccourci `.url` vers n'importe quelle adresse web.

```powershell
.\New-DesktopUrlShortcut.ps1 -Name "Company Portal" -Url "https://portal.contoso.com/" -Apply
```

---

### Remove-DesktopShortcut.ps1

Supprime les raccourcis du bureau correspondant à un modèle générique, sur le Public Desktop ou sur le bureau de l'utilisateur actuel.

```powershell
.\Remove-DesktopShortcut.ps1 -NamePattern "*.rdp" -Apply
```

---

### Add-NetworkPrinterConnection.ps1

Ajoute un port d'imprimante TCP/IP + une connexion d'imprimante à l'aide d'un pilote déjà installé.

| Paramètre | Description |
|-----------|-------------|
| `-PrinterName` | Nom d'affichage (obligatoire) |
| `-PortAddress` | IP/nom d'hôte de l'imprimante (obligatoire) |
| `-DriverName` | Nom du pilote déjà installé (obligatoire) |
| `-PortName` | Nom de l'objet port (par défaut : `IP_<PortAddress>`) |
| `-Apply` | Créer réellement (par défaut : aperçu uniquement) |

```powershell
.\Add-NetworkPrinterConnection.ps1 -PrinterName "Label Printer - Warehouse" -PortAddress "10.0.5.50" -DriverName "Dymo LabelWriter 450 Turbo" -Apply
```

> Installez d'abord le pilote d'imprimante : ce script crée uniquement le port et la connexion.

---

Tous les scripts effectuent un essai à blanc par défaut ; passez `-Apply` pour appliquer les modifications, conformément au style maison.
