# ActiveDirectory

On-prem Active Directory Domain Services monitoring — as opposed to [`Entra/`](../Entra/readme.md), which targets the cloud directory via Microsoft Graph.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Watch-ADAccountLockouts.ps1`](#watch-adaccountlockoutsps1) | Monitor AD for locked-out user accounts, log only new lockouts, every 20 minutes |

---

### Watch-ADAccountLockouts.ps1

Queries the domain's **PDC Emulator** for currently locked-out user accounts — deliberately not a random/local DC, since lockout counters and `lockoutTime` are tracked per-DC and only guaranteed authoritative on the PDC Emulator until replication catches up elsewhere.

Keeps a small state file (`state.json`, next to the log) so only a genuinely **new** lockout — first time seen, or an unlock-then-relock of the same account — gets written to the log. An account that stays locked across many runs is not re-logged every run.

**Parameters**

| Parameter | Description |
|-----------|-------------|
| `-SearchBase` | Optional. Limit to one OU (distinguished name). Default: whole domain |
| `-LogPath` | Override the log file (default: `C:\ProgramData\ADLockoutMonitor\lockouts.log`) |
| `-RegisterTask` | Register this script as a recurring Scheduled Task instead of running a single check |
| `-TaskIntervalMinutes` | Repetition interval for `-RegisterTask` (default: `20`) |

**Examples**

```powershell
# One-off check — same thing the scheduled task runs every 20 minutes
.\Watch-ADAccountLockouts.ps1

# Register the recurring task once, on the DC/file server (runs as SYSTEM)
.\Watch-ADAccountLockouts.ps1 -RegisterTask

# Only monitor one OU, check every 5 minutes instead
.\Watch-ADAccountLockouts.ps1 -RegisterTask -TaskIntervalMinutes 5 -SearchBase "OU=Sales,DC=contoso,DC=com"
```

**Notes**
- No email/Teams alerting — this only writes to the log file. Point your monitoring/RMM tool at `lockouts.log`, or tail it, to get alerted.
- Run `-RegisterTask` **once** — it registers a Scheduled Task named `AD Account Lockout Monitor` that then re-invokes this same script (with your `-LogPath`/`-SearchBase`) every `-TaskIntervalMinutes`, as SYSTEM.
- Requires the `ActiveDirectory` PowerShell module (present on Domain Controllers; on a file server, install the RSAT `RSAT-AD-PowerShell` feature — `Install-WindowsFeature RSAT-AD-PowerShell`).
