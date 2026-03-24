# SAS Batch Error Monitoring

Monitors SAS batch job logs and Windows Event Viewer for errors, with optional Zabbix integration and email alerts.

---

## Files

| File | Description |
|------|-------------|
| `Monitor-SASBatchErrors.ps1` | Main script — scans log files and Event Viewer |
| `Setup-SASMonitoring.ps1` | One-time setup — installs script, scheduled task, Zabbix config |
| `Test-SASWorkDirectory.ps1` | Validates SAS WORK directory health and permissions |
| `zabbix_sas_monitor.conf` | Example Zabbix UserParameter config |

---

## Detected Errors

| Type | Severity | Pattern |
|------|----------|---------|
| `SpawnError` | Critical | `Can't spawn "sas.bat"` |
| `WorkLibAuth` | Critical | WORK library authorization error |
| `SASAbort` | Critical | `SAS has ABORTED processing` |
| `DiskError` | Critical | Disk/filesystem errors (Event Viewer) |
| `SQLViewError` | High | SQL view definition failure |
| `GeneralError` | Medium | Any `^ERROR:` line |

---

## Setup

```powershell
# Run as Administrator
.\Setup-SASMonitoring.ps1

# With Zabbix integration
.\Setup-SASMonitoring.ps1 -InstallZabbix

# With email alerts
.\Setup-SASMonitoring.ps1 -EmailAlerts -SmtpServer "smtp.contoso.com" -EmailTo "admin@contoso.com" -EmailFrom "sas-monitor@contoso.com"
```

Setup installs `Monitor-SASBatchErrors.ps1` to `C:\Scripts\` and creates a scheduled task that runs daily at 08:00.

---

## Usage

```powershell
# Scan last 7 days, text output
.\Monitor-SASBatchErrors.ps1 -LogDirectory "E:\SAS\Logs"

# Scan last 24 hours
.\Monitor-SASBatchErrors.ps1 -LogDirectory "E:\SAS\Logs" -DaysToCheck 1

# JSON output (for automation)
.\Monitor-SASBatchErrors.ps1 -LogDirectory "E:\SAS\Logs" -OutputFormat JSON

# Zabbix output (returns error count)
.\Monitor-SASBatchErrors.ps1 -LogDirectory "E:\SAS\Logs" -OutputFormat Zabbix

# Include Event Viewer
.\Monitor-SASBatchErrors.ps1 -LogDirectory "E:\SAS\Logs" -IncludeEventLog

# Save to file
.\Monitor-SASBatchErrors.ps1 -LogDirectory "E:\SAS\Logs" -OutputFile "C:\Temp\report.txt"
```

**Exit codes:** `0` = no critical/high errors · `1` = high severity · `2` = critical · `-1` = script error

---

## Zabbix Integration

Copy `zabbix_sas_monitor.conf` to `C:\Program Files\Zabbix Agent 2\zabbix_agent2.d\`, then restart the Zabbix Agent service. Or use `Setup-SASMonitoring.ps1 -InstallZabbix` to do this automatically.

Zabbix items:

| Key | Description |
|-----|-------------|
| `sas.batch.errors.critical` | Critical errors in last 7 days |
| `sas.batch.errors.critical.24h` | Critical errors in last 24 hours |
| `sas.batch.errors.json` | Full JSON report |

---

## Adding Error Patterns

Edit `Monitor-SASBatchErrors.ps1` and add to `$ErrorPatterns`:

```powershell
'MyNewError' = @{
    Pattern     = 'your regex here'
    Severity    = 'Critical'   # Critical / High / Medium
    Description = 'What this error means'
}
```
