[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../readme.fr.md) › [scripts](../readme.fr.md) › **Network**

# Network

Scripts de diagnostic réseau et de connectivité. Multiplateformes lorsque c'est indiqué ; les autres nécessitent Windows + Administrateur.

---

## Dossiers

| Dossier | Description |
|--------|-------------|
| [`UniFi/`](UniFi/readme.fr.md) | Rapport de documentation réseau du UniFi Controller + outillage de mise à niveau du firmware |

## Scripts

| Script | Description |
|--------|-------------|
| [`Test-Ports.ps1`](Test-Ports.ps1) ([docs](#test-portsps1)) | Vérification de la connectivité des ports TCP (multiplateforme) |
| [`Test-AuthNetworkDiagnostics.ps1`](Test-AuthNetworkDiagnostics.ps1) ([docs](#test-authnetworkdiagnosticsps1)) | Diagnostic des problèmes d'authentification/réseau (Event Viewer, Kerberos, DNS, partages) |
| [`Test-FileIODiagnostics.ps1`](Test-FileIODiagnostics.ps1) ([docs](#test-fileiodiagnosticsps1)) | Test de charge des E/S fichiers avec diagnostic des échecs en direct |

---

### Test-Ports.ps1

Teste la connectivité TCP sur un ou plusieurs ports vers un ou plusieurs hôtes. Prend en charge les ports isolés, les plages (`1294:1494`) et les combinaisons séparées par des virgules. Fonctionne sous Windows (PS 5.1+), macOS et Linux (PS 7+).

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-Target` | Oui | Adresse IP ou nom(s) d'hôte à tester |
| `-Ports` | Oui | Spécification des ports : isolé (`80`), plage (`1294:1494`), liste (`80,443,3389`) ou mixte (`80,443,1294:1494`) |
| `-TimeoutMs` | Non | Délai d'expiration de la connexion TCP en ms (par défaut : `500`) |
| `-ShowClosed` | Non | Inclure les ports fermés dans la sortie (par défaut, seuls les ports ouverts sont affichés) |

**Exemples**

```powershell
.\Test-Ports.ps1 -Target 192.168.1.1 -Ports 1294:1494
.\Test-Ports.ps1 -Target 10.0.0.1 -Ports 80,443,3389,8080:8090
.\Test-Ports.ps1 -Target server01.contoso.local -Ports 22,3389 -TimeoutMs 1000 -ShowClosed
.\Test-Ports.ps1 -Target 10.0.0.1,10.0.0.2 -Ports 80,443
```

---

### Test-AuthNetworkDiagnostics.ps1

Diagnostique les problèmes d'authentification et de réseau sur un serveur/poste Windows : analyse l'Event Viewer à la recherche d'échecs d'ouverture de session (4625), d'erreurs Kerberos (4771/4768), d'échecs NTLM (4776) et d'erreurs SMB/Netlogon/client DNS ; vérifie la synchronisation de l'heure, le cache des tickets Kerberos, la résolution DNS, la connectivité TCP, l'accès aux partages UNC et l'état des cartes réseau ; analyse éventuellement des fichiers journaux à la recherche de motifs d'erreur. Écrit un rapport txt dans `C:\Temp\`.

**Paramètres**

| Paramètre | Description |
|-----------|-------------|
| `-TestHosts` | Noms d'hôte/IP sur lesquels tester la résolution DNS et la connectivité TCP |
| `-TestPorts` | Ports TCP à tester sur chaque entrée de `-TestHosts` (par défaut : `445, 88, 389, 636`) |
| `-TestShares` | Chemins UNC dont tester l'accès en lecture (p. ex. `\\server\share`) |
| `-LogDirectory` | Dossier facultatif à analyser pour y chercher des motifs d'erreur d'authentification/réseau |
| `-LogDaysBack` | Nombre de jours en arrière pour inclure les fichiers journaux (par défaut : `1`) |
| `-EventLogHours` | Nombre d'heures en arrière à vérifier dans l'Event Viewer (par défaut : `24`) |
| `-OutputPath` | Remplacer le dossier de sortie par défaut (`C:\Temp\`) |

**Exemples**

```powershell
# Exécution de base — Event Viewer uniquement
.\Test-AuthNetworkDiagnostics.ps1

# Tester la connectivité + les partages UNC
.\Test-AuthNetworkDiagnostics.ps1 -TestHosts "dc01","fileserver" -TestShares "\\fileserver\data"

# Analyse complète, fichiers journaux SAS compris
.\Test-AuthNetworkDiagnostics.ps1 -TestHosts "sasserver" -LogDirectory "E:\SAS\Logs"

# Vérifier les événements des 48 dernières heures
.\Test-AuthNetworkDiagnostics.ps1 -EventLogHours 48
```

Nécessite les droits Administrateur.

---

### Test-FileIODiagnostics.ps1

Exécute une boucle écriture/ajout/lecture/suppression sur un chemin cible et, à chaque échec, classe l'erreur (AUTH/NETWORK/TIMEOUT/DISK/PATH/IO) et capture le contexte de diagnostic : événements FileSystemWatcher, un diff des autorisations NTFS (`icacls`) par rapport à la référence prise au démarrage, les handles de fichiers ouverts via Sysinternals `Handle.exe` (téléchargé automatiquement dans `C:\Temp\handle\`), un instantané des nouveaux processus, les tickets Kerberos et les entrées du journal d'événements Security. S'arrête après 3 échecs.

**Paramètres**

| Paramètre | Description |
|-----------|-------------|
| `-TestPath` | Dossier à tester — local ou UNC (p. ex. `G:\sas\work`, `\\server\share\folder`) |
| `-Iterations` | Nombre de cycles écriture/suppression (par défaut : `5000`) |
| `-DelayMs` | Millisecondes entre les itérations (par défaut : `50`) |
| `-StopOnFirstError` | S'arrêter après le premier échec au lieu de continuer jusqu'à 3 |
| `-HandleExe` | Chemin vers un `Handle.exe` existant — évite le téléchargement automatique |
| `-SkipHandleDownload` | Ne pas tenter de télécharger `Handle.exe` (p. ex. serveur isolé du réseau) |
| `-LogPath` | Dossier de sortie du fichier journal (par défaut : `C:\Temp\`) |

**Exemples**

```powershell
# Exécution de base — télécharge Handle.exe automatiquement
.\Test-FileIODiagnostics.ps1 -TestPath G:\sas\work

# Test de charge rapide, arrêt au premier échec
.\Test-FileIODiagnostics.ps1 -TestPath G:\sas\work -Iterations 10000 -DelayMs 0 -StopOnFirstError

# Utiliser un Handle.exe existant, sans téléchargement
.\Test-FileIODiagnostics.ps1 -TestPath G:\sas\work -HandleExe C:\Tools\handle.exe

# Serveur isolé du réseau — ignorer le téléchargement
.\Test-FileIODiagnostics.ps1 -TestPath G:\sas\work -SkipHandleDownload
```
