# Time Sync

Fixes Windows time synchronization drift by pointing `W32time` at Dutch NTP pool servers and keeping it in sync going forward.

---

## Files

| File | Description |
|------|-------------|
| [`Restart-Time-Sync.ps1`](#restart-time-syncps1) | Force an immediate resync and register a recurring scheduled task |

---

### Restart-Time-Sync.ps1

1. Sets the `W32time` service to Automatic startup and starts it
2. Configures the NTP peer list to `0.nl.pool.ntp.org` / `1.nl.pool.ntp.org` and forces an immediate resync
3. Writes that same logic to `%ProgramFiles%\EOO\Restart-NTP.ps1`
4. Registers a scheduled task (**"Restart NTP"**, runs as `SYSTEM`) that reruns the resync every 59 minutes, starting at 8am, for ~27 years (`RepetitionDuration` 9999 days)

```powershell
.\Restart-Time-Sync.ps1
```

> No parameters. Run once per device — the scheduled task takes care of keeping time in sync afterward.
