[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../readme.fr.md) › [scripts](../readme.fr.md) › **RDS**

# RDS

Scripts de diagnostic et de supervision pour l'infrastructure RDP / RD Web Access. Exécutez-les directement sur le serveur RDS/RDWeb pour obtenir des résultats complets — les cibles distantes ne bénéficient que de vérifications au niveau de la connectivité.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Test-RDSDiagnostics.ps1`](Test-RDSDiagnostics.ps1) ([docs](#test-rdsdiagnosticsps1)) | Contrôle d'état ponctuel — services, configuration, certificats, compte utilisateur, journaux d'événements |
| [`Watch-RDSLive.ps1`](Watch-RDSLive.ps1) ([docs](#watch-rdsliveps1)) | Surveillance en temps réel des événements de session et de licences |
| [`Get-FSlogix-errors.ps1`](Get-FSlogix-errors.ps1) ([docs](#get-fslogix-errorsps1)) | Diagnostic des profils FSLogix / Azure Files sur un hôte de session AVD |

---

### Test-RDSDiagnostics.ps1

Détermine pourquoi les utilisateurs ne peuvent pas se connecter à un serveur RDP ou RD Web Access.

**Vérifications effectuées**

| Domaine | Détails |
|------|---------|
| Serveur RDP | État de `TermService`/`SessionEnv`/`UmRdpService`, RDP activé/désactivé, NLA, limites de session, mode RD Licensing, groupe Remote Desktop Users, règles de pare-feu, sessions actives (`quser`) |
| Serveur RDWeb | Accessibilité du port 443, validité/expiration du certificat HTTPS, IIS + pool d'applications RDWeb (en local uniquement), service RD Gateway (en local uniquement) |
| Compte utilisateur (facultatif, `-Username`) | Activé/verrouillé/expiré, appartenance aux groupes, restrictions de postes d'ouverture de session, dernière connexion, âge du mot de passe |
| Journaux d'événements (facultatif, `-IncludeEventLogs`) | Security 4625 (échec d'ouverture de session RDP), 4740 (verrouillage), `TerminalServices-LocalSessionManager` 20/40 (échec de session/motif de déconnexion) |

**Paramètres**

| Paramètre | Description |
|-----------|-------------|
| `-RdpServer` | Nom d'hôte/IP du serveur RDP/Terminal Server. À omettre pour exécuter les vérifications locales |
| `-RdWebServer` | Nom d'hôte/IP du serveur RD Web Access |
| `-Username` | Vérifier ce qui bloque la connexion d'un compte précis (nécessite le module ActiveDirectory ou le repli ADSI) |
| `-LogPath` | Dossier du fichier journal de sortie (par défaut : `C:\Temp\`) |
| `-IncludeEventLogs` | Inclure l'analyse des journaux d'événements des `-Hours` dernières heures |
| `-Hours` | Nombre d'heures d'historique des journaux d'événements à analyser (par défaut : `24`) |

**Exemples**

```powershell
# Vérification complète — RDP + RDWeb + compte utilisateur
.\Test-RDSDiagnostics.ps1 -RdpServer rdp01.company.local -RdWebServer rdweb.company.local -Username jdoe -IncludeEventLogs

# Exécution locale sur l'hôte RDS, avec vérification des journaux d'événements
.\Test-RDSDiagnostics.ps1 -IncludeEventLogs -Hours 48
```

Les résultats sont écrits dans la console et dans un fichier journal horodaté dans `C:\Temp\`.

---

### Watch-RDSLive.ps1

Interroge les journaux d'événements Windows toutes les N secondes et diffuse les nouveaux événements dans la console + un fichier journal. À exécuter directement sur chaque serveur RDS/RDWeb.

**Événements surveillés**

| Source | Événements |
|--------|--------|
| `TerminalServices-LocalSessionManager` | Ouverture de session (21), reconnexion (22/25), fermeture de session (23), déconnexion (24), échec d'ouverture de session (20), motif de déconnexion (40) — codes de motif lisibles |
| Security | Échec d'ouverture de session RDP (4625, type 10), verrouillage de compte (4740) |
| `TerminalServices-Licensing` | Événements de licence accordée/refusée/avertissement |
| System | Fournisseur `TermServLicensing` (période de grâce, erreurs du serveur de licences) |

Affiche une ligne de pulsation (heartbeat) à chaque interrogation, avec le nombre de sessions actives.

**Paramètres**

| Paramètre | Description |
|-----------|-------------|
| `-IntervalSeconds` | Intervalle d'interrogation en secondes (par défaut : `20`) |
| `-LogPath` | Dossier du fichier journal de sortie (par défaut : `C:\Temp\`) |
| `-NoLogFile` | Sortie console uniquement, sans fichier journal |

**Exemples**

```powershell
# Exécuter sur le serveur RDS
.\Watch-RDSLive.ps1

# Interrogation plus rapide, sans fichier journal
.\Watch-RDSLive.ps1 -IntervalSeconds 10 -NoLogFile
```

> Appuyez sur `Ctrl+C` pour arrêter. Exécutez en tant qu'Administrateur pour accéder au journal Security.

---

### Get-FSlogix-errors.ps1

Rassemble, en une seule exécution, tout ce qu'il faut pour comprendre pourquoi un profil
FSLogix ne se monte pas sur un hôte de session AVD — erreurs de montage, VHDX verrouillé,
problème SMB/Azure Files ou erreurs de disque. Lecture seule : il collecte et rend compte, il ne répare rien.

**Ce qu'il collecte**

| Domaine | Détails |
|------|---------|
| Système | Nom d'hôte, build de l'OS, uptime |
| FSLogix | Version installée, la configuration complète `Profiles`/`Containers` et l'état du service |
| Conteneurs | Fichiers VHD(X) attachés, le registre des sessions FSLogix et les chemins de profil issus de `ProfileList` |
| Stockage | Connexions SMB vers Azure Files, et si le partage VHD est seulement joignable |
| Événements | Événements FSLogix des `-Days` derniers jours, comparés aux motifs d'échec critiques connus, plus les erreurs de disque/NTFS et les événements du User Profile Service |
| Restes | Profils locaux sous `C:\Users` et fichiers journaux FSLogix |

**Paramètres**

| Paramètre | Description |
|-----------|-------------|
| `-User` | Produire en plus une section filtrée sur un seul utilisateur — le compte dont le profil échoue |
| `-Days` | Nombre de jours d'historique d'événements à analyser (par défaut : `7`) |
| `-OutputPath` | Dossier de la transcription (par défaut : `%SystemDrive%\Temp\FSLogixDiag`) |

**Exemples**

```powershell
# Tout ce qui concerne la dernière semaine
.\Get-FSlogix-errors.ps1

# Un seul utilisateur, deux semaines en arrière, rapport ailleurs
.\Get-FSlogix-errors.ps1 -User jdoe -Days 14 -OutputPath C:\Temp
```

> Exécutez-le dans une session élevée **sur l'hôte de session lui-même** — les données de
> conteneur, SMB et événements n'existent que là. L'ensemble de l'exécution est écrit dans
> `FSLogixDiag_<host>_<timestamp>.log` dans le dossier de sortie ; c'est ce fichier qu'il
> faut joindre à un ticket.
