[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [Device](../readme.fr.md) › **Time sync**

# Synchronisation de l'heure

Corrige la dérive de la synchronisation de l'heure Windows en faisant pointer `W32time` vers les serveurs du pool NTP néerlandais et en maintenant ensuite l'heure synchronisée.

---

## Fichiers

| Fichier | Description |
|------|-------------|
| [`Restart-Time-Sync.ps1`](Restart-Time-Sync.ps1) ([docs](#restart-time-syncps1)) | Force une resynchronisation immédiate et enregistre une tâche planifiée récurrente |

---

### Restart-Time-Sync.ps1

1. Passe le service `W32time` en démarrage automatique et le démarre
2. Configure la liste des pairs NTP sur `0.nl.pool.ntp.org` / `1.nl.pool.ntp.org` et force une resynchronisation immédiate
3. Écrit cette même logique dans `%ProgramFiles%\EOO\Restart-NTP.ps1`
4. Enregistre une tâche planifiée (**"Restart NTP"**, exécutée en tant que `SYSTEM`) qui relance la resynchronisation toutes les 59 minutes, à partir de 8 h, pendant ~27 ans (`RepetitionDuration` 9999 jours)

```powershell
.\Restart-Time-Sync.ps1
```

> Aucun paramètre. À exécuter une fois par appareil — la tâche planifiée maintient ensuite l'heure synchronisée.
