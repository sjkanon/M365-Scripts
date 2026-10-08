[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../readme.fr.md) › [scripts](../readme.fr.md) › **SAS**

# Supervision des erreurs des jobs batch SAS

Surveille les journaux des jobs batch SAS et l'Event Viewer de Windows à la recherche d'erreurs, avec intégration Zabbix et alertes par e-mail facultatives.

---

## Fichiers

| Fichier | Description |
|------|-------------|
| [`Monitor-SASBatchErrors.ps1`](Monitor-SASBatchErrors.ps1) ([docs](#monitor-sasbatcherrorsps1)) | Script principal — analyse les fichiers journaux et l'Event Viewer |
| [`Setup-SASMonitoring.ps1`](Setup-SASMonitoring.ps1) ([docs](#setup-sasmonitoringps1)) | Configuration initiale unique — installe le script, la tâche planifiée et la configuration Zabbix |
| [`Test-SASWorkDirectory.ps1`](Test-SASWorkDirectory.ps1) ([docs](#test-sasworkdirectoryps1)) | Vérifie l'état et les autorisations du répertoire SAS WORK |
| `rca.md` | Analyse des causes profondes des échecs intermittents access-denied lors des suppressions dans SAS WORK |
| `zabbix_sas_monitor.conf` | Exemple de configuration Zabbix UserParameter |

---

## Erreurs détectées

| Type | Gravité | Motif |
|------|----------|---------|
| `SpawnError` | Critical | `Can't spawn "sas.bat"` |
| `WorkLibAuth` | Critical | Erreur d'autorisation sur la bibliothèque WORK |
| `SASAbort` | Critical | `SAS has ABORTED processing` |
| `DiskError` | Critical | Erreurs de disque/système de fichiers (Event Viewer) |
| `SQLViewError` | High | Échec de la définition d'une vue SQL |
| `GeneralError` | Medium | Toute ligne `^ERROR:` |

---

## Setup-SASMonitoring.ps1

```powershell
# À exécuter en tant qu'Administrateur
.\Setup-SASMonitoring.ps1

# Avec intégration Zabbix
.\Setup-SASMonitoring.ps1 -InstallZabbix

# Avec alertes par e-mail
.\Setup-SASMonitoring.ps1 -EmailAlerts -SmtpServer "smtp.contoso.com" -EmailTo "admin@contoso.com" -EmailFrom "sas-monitor@contoso.com"
```

La configuration installe `Monitor-SASBatchErrors.ps1` dans `C:\Scripts\` et crée une tâche planifiée qui s'exécute tous les jours à 08:00.

---

## Monitor-SASBatchErrors.ps1

```powershell
# Analyser les 7 derniers jours, sortie texte
.\Monitor-SASBatchErrors.ps1 -LogDirectory "E:\SAS\Logs"

# Analyser les dernières 24 heures
.\Monitor-SASBatchErrors.ps1 -LogDirectory "E:\SAS\Logs" -DaysToCheck 1

# Sortie JSON (pour l'automatisation)
.\Monitor-SASBatchErrors.ps1 -LogDirectory "E:\SAS\Logs" -OutputFormat JSON

# Sortie Zabbix (renvoie le nombre d'erreurs)
.\Monitor-SASBatchErrors.ps1 -LogDirectory "E:\SAS\Logs" -OutputFormat Zabbix

# Inclure l'Event Viewer
.\Monitor-SASBatchErrors.ps1 -LogDirectory "E:\SAS\Logs" -IncludeEventLog

# Enregistrer dans un fichier
.\Monitor-SASBatchErrors.ps1 -LogDirectory "E:\SAS\Logs" -OutputFile "C:\Temp\report.txt"
```

**Codes de sortie :** `0` = aucune erreur critical/high · `1` = gravité high · `2` = critical · `-1` = erreur du script

---

## Test-SASWorkDirectory.ps1

Vérifie que les répertoires SAS WORK (`G:\sas\work`) et USERWORK (`U:\sas\userwork`) sont accessibles et fonctionnent, avec des tests d'E/S simples qui ne perturbent pas les jobs SAS en cours (`-Iterations`, 10 par défaut), et contrôle le profil des disques de travail éphémères.

```powershell
# Contrôle de l'état de WORK/USERWORK avec diagnostic antivirus/filtres
.\Test-SASWorkDirectory.ps1 -Iterations 1000 -EventLogHours 2 -IncludeAVDiagnostics $true -AVLogHours 2
```

Le diagnostic antivirus de `Test-SASWorkDirectory.ps1` comprend :

- Un instantané de l'état de Defender (protection en temps réel, surveillance du comportement, antivirus activé)
- Une vérification des exclusions Defender par rapport aux chemins WORK/USERWORK surveillés
- Les événements Defender Operational (blocked/quarantine/denied/CFA/tamper et messages liés aux chemins)
- Les événements des pilotes de filtre du journal System (FilterManager/WdFilter)
- Un instantané des minifiltres actifs (`fltmc filters`)

---

## Remarque sur les disques éphémères (G: / U:)

Si `G:` et `U:` sont des disques temporaires éphémères/locaux dans votre environnement :

- Ils conviennent aux données temporaires SAS `WORK`/`USERWORK`, mais pas comme stockage persistant.
- Déplacer la charge de `G:` vers `U:` ne constitue pas un basculement structurel si les deux sont éphémères.
- Les `Access is denied` intermittents pendant des boucles d'E/S intensives sont souvent causés par des verrouillages/une activité de filtres transitoires (antivirus, sauvegarde, indexation, EDR) plutôt que par l'expiration de Kerberos.

`Test-SASWorkDirectory.ps1` journalise désormais le contexte du profil de disque et signale explicitement les lettres de lecteur éphémères configurées, ce qui rend plus claire l'analyse des échecs dans les journaux redirigés.

Il classe aussi plus précisément les événements SAS du journal Application :

- `hc_disk_delete*` + `Access is denied` / code retour `5` sont signalés comme des échecs de suppression WORK nécessitant une action.
- `ARM Application data not available` est journalisé comme bruit de télémétrie informatif, sauf si d'autres erreurs SAS sont présentes.
- Les entrées d'événements SAS en double ayant le même horodatage/message sont dédoublonnées dans la sortie.

---

## Intégration Zabbix

Copiez `zabbix_sas_monitor.conf` dans `C:\Program Files\Zabbix Agent 2\zabbix_agent2.d\`, puis redémarrez le service Zabbix Agent. Ou utilisez `Setup-SASMonitoring.ps1 -InstallZabbix` pour le faire automatiquement.

Éléments Zabbix :

| Clé | Description |
|-----|-------------|
| `sas.batch.errors.critical` | Erreurs critiques des 7 derniers jours |
| `sas.batch.errors.critical.24h` | Erreurs critiques des dernières 24 heures |
| `sas.batch.errors.json` | Rapport JSON complet |

---

## Ajouter des motifs d'erreur

Modifiez `Monitor-SASBatchErrors.ps1` et ajoutez à `$ErrorPatterns` :

```powershell
'MyNewError' = @{
    Pattern     = 'your regex here'
    Severity    = 'Critical'   # Critical / High / Medium
    Description = 'What this error means'
}
```
