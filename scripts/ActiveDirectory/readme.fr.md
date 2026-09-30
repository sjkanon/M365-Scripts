[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../readme.fr.md) › [scripts](../readme.fr.md) › **ActiveDirectory**

# ActiveDirectory

Supervision d'Active Directory Domain Services on-premise — par opposition à [`Entra/`](../Entra/readme.fr.md), qui cible l'annuaire cloud via Microsoft Graph.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Watch-ADAccountLockouts.ps1`](Watch-ADAccountLockouts.ps1) ([docs](#watch-adaccountlockoutsps1)) | Surveiller AD toutes les 20 minutes pour détecter les comptes utilisateurs verrouillés, en ne journalisant que les nouveaux verrouillages |

---

### Watch-ADAccountLockouts.ps1

Interroge l'**émulateur PDC** (PDC Emulator) du domaine pour obtenir les comptes utilisateurs actuellement verrouillés — volontairement pas un DC aléatoire/local, car les compteurs de verrouillage et `lockoutTime` sont suivis par DC et ne font autorité de manière garantie que sur l'émulateur PDC, jusqu'à ce que la réplication rattrape son retard ailleurs.

Conserve un petit fichier d'état (`state.json`, à côté du journal) afin que seul un verrouillage réellement **nouveau** — vu pour la première fois, ou un déverrouillage suivi d'un nouveau verrouillage du même compte — soit écrit dans le journal. Un compte qui reste verrouillé sur de nombreuses exécutions n'est pas rejournalisé à chaque exécution.

**Paramètres**

| Paramètre | Description |
|-----------|-------------|
| `-SearchBase` | Facultatif. Limiter à une seule OU (distinguished name). Par défaut : tout le domaine |
| `-LogPath` | Remplacer le fichier journal (par défaut : `C:\ProgramData\ADLockoutMonitor\lockouts.log`) |
| `-RegisterTask` | Enregistrer ce script comme tâche planifiée récurrente au lieu d'exécuter une seule vérification |
| `-TaskIntervalMinutes` | Intervalle de répétition pour `-RegisterTask` (par défaut : `20`) |

**Exemples**

```powershell
# Vérification ponctuelle — la même chose que ce que la tâche planifiée exécute toutes les 20 minutes
.\Watch-ADAccountLockouts.ps1

# Enregistrer la tâche récurrente une seule fois, sur le DC/serveur de fichiers (s'exécute en tant que SYSTEM)
.\Watch-ADAccountLockouts.ps1 -RegisterTask

# Ne surveiller qu'une seule OU, et vérifier plutôt toutes les 5 minutes
.\Watch-ADAccountLockouts.ps1 -RegisterTask -TaskIntervalMinutes 5 -SearchBase "OU=Sales,DC=contoso,DC=com"
```

**Remarques**
- Pas d'alerte par e-mail/Teams — le script écrit uniquement dans le fichier journal. Pointez votre outil de supervision/RMM sur `lockouts.log`, ou suivez-le avec tail, pour être alerté.
- Exécutez `-RegisterTask` **une seule fois** — cela enregistre une tâche planifiée nommée `AD Account Lockout Monitor` qui rappelle ensuite ce même script (avec vos `-LogPath`/`-SearchBase`) toutes les `-TaskIntervalMinutes`, en tant que SYSTEM.
- Nécessite le module PowerShell `ActiveDirectory` (présent sur les contrôleurs de domaine ; sur un serveur de fichiers, installez la fonctionnalité RSAT `RSAT-AD-PowerShell` — `Install-WindowsFeature RSAT-AD-PowerShell`).
