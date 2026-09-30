[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [Device](../readme.nl.md) › **Time sync**

# Tijdsynchronisatie

Verhelpt afwijkingen in de Windows-tijdsynchronisatie door `W32time` naar Nederlandse NTP-poolservers te laten wijzen en de tijd daarna gesynchroniseerd te houden.

---

## Bestanden

| Bestand | Omschrijving |
|------|-------------|
| [`Restart-Time-Sync.ps1`](Restart-Time-Sync.ps1) ([docs](#restart-time-syncps1)) | Forceert direct een hersynchronisatie en registreert een terugkerende geplande taak |

---

### Restart-Time-Sync.ps1

1. Zet de service `W32time` op automatisch opstarten en start hem
2. Stelt de NTP-peerlijst in op `0.nl.pool.ntp.org` / `1.nl.pool.ntp.org` en forceert direct een hersynchronisatie
3. Schrijft diezelfde logica naar `%ProgramFiles%\EOO\Restart-NTP.ps1`
4. Registreert een geplande taak (**"Restart NTP"**, draait als `SYSTEM`) die de hersynchronisatie elke 59 minuten opnieuw uitvoert, vanaf 8 uur 's ochtends, gedurende ~27 jaar (`RepetitionDuration` 9999 dagen)

```powershell
.\Restart-Time-Sync.ps1
```

> Geen parameters. Draai het één keer per apparaat — daarna houdt de geplande taak de tijd gesynchroniseerd.
