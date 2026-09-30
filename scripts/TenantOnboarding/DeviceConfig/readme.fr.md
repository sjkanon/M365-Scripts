[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [TenantOnboarding](../readme.fr.md) › **DeviceConfig**

# DeviceConfig

Scripts de configuration et de durcissement des appareils Windows, utilisés lors de l'intégration des tenants/appareils : auto-élévation via groupes locaux, durcissement du stockage des identifiants, paramètres d'alimentation kiosque, suppression d'Office, disposition du menu Démarrer et règle de pare-feu Teams pour le LAN.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Set-LocalGroupSelfElevation.ps1`](Set-LocalGroupSelfElevation.ps1) ([docs](#set-localgroupselfelevationps1)) | Accorder/révoquer à la prochaine ouverture de session l'appartenance de l'utilisateur connecté à un groupe local |
| [`Disable-CredentialManagerVault.ps1`](Disable-CredentialManagerVault.ps1) ([docs](#disable-credentialmanagervaultps1)) | Désactiver le stockage local des mots de passe du Gestionnaire d'identification Windows |
| [`Set-KioskPowerSettings.ps1`](Set-KioskPowerSettings.ps1) ([docs](#set-kioskpowersettingsps1)) | Désactiver la veille/le démarrage rapide pour les appareils kiosque ou toujours allumés |
| [`Uninstall-MicrosoftOffice.ps1`](Uninstall-MicrosoftOffice.ps1) ([docs](#uninstall-microsoftofficeps1)) | Désinstaller silencieusement Microsoft Office / Microsoft 365 Apps |
| [`Import-StartMenuLayout.ps1`](Import-StartMenuLayout.ps1) ([docs](#import-startmenulayoutps1)) | Appliquer un fichier XML de disposition du menu Démarrer |
| [`Set-TeamsFirewallRule.ps1`](Set-TeamsFirewallRule.ps1) ([docs](#set-teamsfirewallruleps1)) | Créer la règle de pare-feu entrante nécessaire à Teams pour le partage d'écran sur le LAN |

---

### Set-LocalGroupSelfElevation.ps1

Regroupe quatre anciens scripts en un seul (accorder/révoquer Administrators local, accorder/révoquer « Network Configuration Operators » local). Enregistre une tâche planifiée SYSTEM, déclenchée à l'ouverture de session, qui ajoute l'utilisateur connecté, quel qu'il soit, au groupe local indiqué ou l'en retire, et supprime toute tâche opposée encore en attente.

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-GroupName` | Oui | Nom du groupe local |
| `-Action` | Oui | `Grant` ou `Revoke` |
| `-TaskName` | Non | Nom de la tâche planifiée (par défaut : `<Action>-<GroupName>`) |
| `-Apply` | Non | Enregistrer réellement la tâche (par défaut : aperçu uniquement) |

```powershell
.\Set-LocalGroupSelfElevation.ps1 -GroupName "Administrators" -Action Grant -Apply
.\Set-LocalGroupSelfElevation.ps1 -GroupName "Administrators" -Action Revoke -Apply
.\Set-LocalGroupSelfElevation.ps1 -GroupName "Network Configuration Operators" -Action Grant -Apply
```

---

### Disable-CredentialManagerVault.ps1

Arrête et désactive le service `VaultSvc`, ce qui empêche la conservation locale des identifiants enregistrés.

```powershell
.\Disable-CredentialManagerVault.ps1 -Apply
```

---

### Set-KioskPowerSettings.ps1

Règle les délais d'extinction de l'écran et de mise en veille sur jamais (secteur et batterie) et désactive le démarrage rapide, pour les appareils kiosque, d'accueil ou toujours allumés.

```powershell
.\Set-KioskPowerSettings.ps1 -Apply
```

---

### Uninstall-MicrosoftOffice.ps1

Recherche les entrées Microsoft Office / Microsoft 365 Apps et les désinstalle silencieusement à l'aide de leurs chaînes de désinstallation enregistrées.

```powershell
.\Uninstall-MicrosoftOffice.ps1
.\Uninstall-MicrosoftOffice.ps1 -Apply
```

---

### Import-StartMenuLayout.ps1

Applique un fichier XML de disposition du menu Démarrer, depuis un fichier local ou une URL.

| Paramètre | Description |
|-----------|-------------|
| `-LayoutXmlPath` | Chemin local du XML de disposition |
| `-LayoutUrl` | URL depuis laquelle télécharger le XML de disposition |
| `-Apply` | Appliquer réellement (par défaut : aperçu uniquement) |

```powershell
.\Import-StartMenuLayout.ps1 -LayoutXmlPath "C:\Deploy\startmenu.xml" -Apply
.\Import-StartMenuLayout.ps1 -LayoutUrl "https://packages.contoso.com/config/startmenu.xml" -Apply
```

---

### Set-TeamsFirewallRule.ps1

Crée la règle de pare-feu entrante autorisant le partage d'écran pair à pair de Teams sur le LAN pour le profil Domaine (bloqué pour les profils Public/Privé), limitée à l'utilisateur actuellement connecté. À exécuter en tant que SYSTEM (application Win32 Intune ou tâche d'ouverture de session).

```powershell
.\Set-TeamsFirewallRule.ps1 -Apply
```

> Concept original (c) Microsoft Corporation 2018 et Michael Mardahl (msendpointmgr.com), fourni en l'état ; ceci est une réécriture au style maison.

---

Tous les scripts effectuent un essai à blanc par défaut ; passez `-Apply` pour appliquer les modifications, conformément au style maison.
