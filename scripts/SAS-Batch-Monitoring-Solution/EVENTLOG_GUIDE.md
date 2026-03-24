# Event Viewer Integration Guide

## Overview

The enhanced monitoring scripts now check **both** SAS log files **and** Windows Event Viewer to catch issues at the system level that may cause SAS batch failures.

## Why Check Event Viewer?

Your SAS errors might have root causes that only show up in Windows logs:

| SAS Error | Possible Event Log Cause |
|-----------|-------------------------|
| `Can't spawn sas.bat` | Disk I/O errors, permission issues, antivirus blocking |
| `WORK library authorization error` | NTFS permission changes, security policy changes |
| General batch failures | Disk full, hardware errors, service crashes |

## What Gets Checked

### System Event Log
- **Disk errors** (Event IDs: 7, 9, 11, 15, 51, 52, 153, 154, 155)
- **Storage errors** (Volume shadow copy, storage driver issues)
- **NTFS filesystem errors**
- **I/O failures**

### Application Event Log
- **SAS application errors**
- **Stargate scheduler errors**
- **SAS service crashes**
- **Launch failures**

### Security Event Log (requires admin)
- **Access denied events** (Event IDs: 4656, 4663)
- **Permission failures** on G:\sas\work and U:\sas\userwork

## Usage Examples

### Test-SASWorkDirectory.ps1

```powershell
# Check disk health + last 24 hours of Event Viewer
.\Test-SASWorkDirectory.ps1

# Check disk health + last 48 hours of Event Viewer
.\Test-SASWorkDirectory.ps1 -EventLogHours 48

# More thorough performance test + event logs
.\Test-SASWorkDirectory.ps1 -Iterations 100 -EventLogHours 24
```

**Sample Output:**
```
=== SAS Work Directory Health Check ===

--- Checking Windows Event Logs ---
[2024-12-09 14:30:45] [INFO] Checking Event Viewer logs (last 24 hours)...
[2024-12-09 14:30:45] [INFO] Scanning System log for disk errors...
[2024-12-09 14:30:45] [WARNING]   Found 2 disk-related events:
[2024-12-09 14:30:45] [WARNING]   [2024-12-09 06:30:04] Disk - Warning
[2024-12-09 14:30:45] [WARNING]     The device, \Device\Harddisk1\DR1, has a bad block.
[2024-12-09 14:30:45] [INFO] Scanning Application log for SAS errors...
[2024-12-09 14:30:46] [SUCCESS]   ✓ No SAS errors found
```

### Monitor-SASBatchErrors-Enhanced.ps1

```powershell
# Standard log file scanning only
.\Monitor-SASBatchErrors-Enhanced.ps1 -LogDirectory "E:\SAS\Logs"

# Include Event Viewer (recommended!)
.\Monitor-SASBatchErrors-Enhanced.ps1 -LogDirectory "E:\SAS\Logs" -IncludeEventLog

# Comprehensive check: 7 days of logs + 24 hours of events
.\Monitor-SASBatchErrors-Enhanced.ps1 -LogDirectory "E:\SAS\Logs" `
    -DaysToCheck 7 `
    -IncludeEventLog `
    -EventLogHours 24 `
    -OutputFormat Text

# For Zabbix with Event Viewer
.\Monitor-SASBatchErrors-Enhanced.ps1 -LogDirectory "E:\SAS\Logs" `
    -IncludeEventLog `
    -OutputFormat Zabbix
```

**Enhanced Output Example:**
```
================================================================================
SAS Batch Error Report (Enhanced) - 2024-12-09 08:00:00
================================================================================

SUMMARY:
----------------------------------------
Total Errors: 5
Critical: 3
High: 2
Medium: 0

Sources:
  Log Files: 3
  Event Viewer: 2

ERRORS BY TYPE:
----------------------------------------

[SpawnError] - 1 occurrence(s)
  Description: SAS.BAT spawning failure

  Source: LogFile
  File/Log: DWH051.13792.20251129063004.log
  Time: 2024-11-29 06:30:04
  Message: Can't spawn "sas.bat": No error at E:\Tools\Stargate\bin\\launchsas.pl
  ---

[DiskError] - 2 occurrence(s)
  Description: Disk/Storage system error

  Source: EventLog
  File/Log: EventLog:System
  Time: 2024-11-29 06:28:15
  Message: The device, \Device\Harddisk1\DR1, has a bad block.
  ---

  Source: EventLog
  File/Log: EventLog:System
  Time: 2024-11-29 06:25:30
  Message: The IO operation at logical block address 0x1234567 for Disk 1 failed
  ---
```

## Correlation Examples

### Example 1: Spawn Error + Disk Error

**Log file shows:**
```
Can't spawn "sas.bat": No error at E:\Tools\Stargate\bin\\launchsas.pl line 192
```

**Event Viewer shows (30 minutes earlier):**
```
Event ID: 51 (Disk Error)
The device, \Device\Harddisk1\DR1, has a bad block.
```

**Root cause:** Disk hardware failure → SAS can't write temp files → spawn fails

---

### Example 2: WORK Library Auth + Security Event

**Log file shows:**
```
ERROR: User does not have appropriate authorization level for library WORK.
```

**Event Viewer shows:**
```
Event ID: 4663 (Security)
An attempt was made to access an object: G:\sas\work
Access: DENIED
```

**Root cause:** NTFS permissions changed (group policy update, admin error)

---

### Example 3: Multiple Batch Failures + Storage Event

**Multiple SAS jobs fail with:**
```
ERROR: SAS has ABORTED processing!
```

**Event Viewer shows:**
```
Event ID: 7 (NTFS)
The file system structure on the disk is corrupt and unusable.
Volume: G:
```

**Root cause:** Filesystem corruption → all G:\sas\work access fails

## Zabbix Integration

### Updated Zabbix Configuration

Add this to your Zabbix template:

```bash
# Check with Event Viewer integration (recommended for critical monitoring)
UserParameter=sas.batch.errors.enhanced,powershell -NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\Monitor-SASBatchErrors-Enhanced.ps1" -LogDirectory "E:\SAS\Logs" -IncludeEventLog -OutputFormat Zabbix

# Separate check for Event Log only (faster, runs more frequently)
UserParameter=sas.eventlog.errors,powershell -NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\Monitor-SASBatchErrors-Enhanced.ps1" -LogDirectory "E:\SAS\Logs" -IncludeEventLog -DaysToCheck 0 -OutputFormat Zabbix
```

### Recommended Monitoring Strategy

| Check | Frequency | Purpose |
|-------|-----------|---------|
| Log files only | Every 10 minutes | Quick error detection |
| Log files + Event Viewer | Every 30 minutes | Comprehensive monitoring |
| Disk health test | Daily at 7:00 AM | Proactive hardware monitoring |

## Troubleshooting

### "Could not read Security log"

**Issue:** Script needs admin privileges to read Security log

**Solution:**
```powershell
# Run as Administrator
Start-Process powershell -Verb RunAs -ArgumentList "-File Test-SASWorkDirectory.ps1"

# Or configure scheduled task to run as SYSTEM
```

### "No Event Viewer data"

**Possible causes:**
1. Event logs are full/rotated
2. Logging is disabled
3. Time range too narrow

**Check Event Log settings:**
```powershell
# Check log size and retention
Get-WinEvent -ListLog System, Application, Security | 
    Select-Object LogName, RecordCount, MaximumSizeInBytes

# Check if logging is enabled
wevtutil gl System
wevtutil gl Application
```

### Performance Impact

Event log queries can be slow on busy systems:

- **System log:** ~100ms for 24 hours
- **Application log:** ~150ms for 24 hours  
- **Security log:** ~500ms for 24 hours (requires admin)

**Optimization tips:**
- Use `-EventLogHours 24` instead of checking full log
- Run comprehensive checks less frequently
- Exclude Security log if not needed

## Best Practices

1. **Always use `-IncludeEventLog` for root cause analysis**
   - Don't just fix the symptom in the SAS log
   - Find the underlying system issue

2. **Check Event Viewer BEFORE restarting failed jobs**
   - If there's a disk error, restarting won't help
   - Fix the root cause first

3. **Set up alerts for specific Event IDs**
   - Event 51, 7, 153 = disk hardware problems (URGENT)
   - Event 4663 = permission issues (investigate security policies)

4. **Correlate timestamps**
   - Event log errors often appear 5-30 minutes before SAS errors
   - Look for patterns (same time each day = scheduled task conflict)

5. **Run health checks BEFORE weekly batch**
   - Saturday morning: `.\Test-SASWorkDirectory.ps1 -EventLogHours 168`
   - Fix any issues before batch starts

## Event ID Quick Reference

### Critical Disk Errors (Requires immediate attention)

- **Event 7**: Disk has bad blocks
- **Event 9**: Disk device error
- **Event 51**: Page file errors / disk timeout
- **Event 153**: Volume shadow copy failure
- **Event 55**: NTFS filesystem corruption

### SAS-Related Application Events

- **Event 1000**: Application crash (SAS.EXE)
- **Event 1001**: Windows Error Reporting
- **Event 5617**: SAS service start failure

### Permission Issues

- **Event 4656**: Object access attempt (includes denials)
- **Event 4663**: Object access (detailed access info)

## Support

For issues or questions about Event Viewer integration:
- Review the correlation examples above
- Check Event Viewer manually: `eventvwr.msc`
- Export relevant events: `wevtutil epl System C:\temp\system_events.evtx`

---

**Pro tip:** When opening a ticket with SAS support about batch failures, ALWAYS include:
1. The SAS log file
2. Relevant Event Viewer exports
3. Output from `Test-SASWorkDirectory.ps1 -IncludeEventLog`
