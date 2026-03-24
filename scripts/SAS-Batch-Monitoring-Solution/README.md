# SAS Batch Error Monitoring Solution

Automated monitoring system for detecting and alerting on SAS batch job errors.

## Overview

This solution monitors SAS batch job logs (both sequence and job logs) for critical errors and provides:
- Real-time error detection
- Daily automated reports
- Zabbix integration for enterprise monitoring
- Email alerts for critical issues
- Historical error tracking

## Detected Error Types

| Error Type | Severity | Description |
|------------|----------|-------------|
| **SpawnError** | Critical | `Can't spawn "sas.bat"` - SAS process spawning failures |
| **WorkLibAuth** | Critical | Authorization level errors for library WORK |
| **SASAbort** | Critical | `SAS has ABORTED processing` - Complete batch failures |
| **SQLViewError** | High | SQL view definition failures |
| **GeneralError** | Medium | Other SAS errors |

## Components

### 1. Monitor-SASBatchErrors.ps1
Main monitoring script that scans log files and detects errors.

**Usage:**
```powershell
# Basic usage - scan last 7 days
.\Monitor-SASBatchErrors.ps1 -LogDirectory "E:\SAS\Logs"

# Scan specific time period
.\Monitor-SASBatchErrors.ps1 -LogDirectory "E:\SAS\Logs" -DaysToCheck 1

# Output formats
.\Monitor-SASBatchErrors.ps1 -LogDirectory "E:\SAS\Logs" -OutputFormat Text
.\Monitor-SASBatchErrors.ps1 -LogDirectory "E:\SAS\Logs" -OutputFormat JSON
.\Monitor-SASBatchErrors.ps1 -LogDirectory "E:\SAS\Logs" -OutputFormat Zabbix

# Save to file
.\Monitor-SASBatchErrors.ps1 -LogDirectory "E:\SAS\Logs" -OutputFile "C:\Reports\SAS_Errors.txt"
```

**Exit Codes:**
- `0` - No critical/high severity errors found
- `1` - High severity errors found
- `2` - Critical errors found
- `-1` - Script execution error

### 2. Setup-SASMonitoring.ps1
Automated installation and configuration script.

**Usage:**
```powershell
# Basic installation
.\Setup-SASMonitoring.ps1

# Install with Zabbix integration
.\Setup-SASMonitoring.ps1 -InstallZabbix

# Install with email alerts
.\Setup-SASMonitoring.ps1 -EmailAlerts -EmailTo "admin@bravehub.be"

# Complete setup
.\Setup-SASMonitoring.ps1 -InstallZabbix -EmailAlerts `
    -SmtpServer "smtp.bravehub.be" `
    -EmailTo "sjoerd@bravehub.be" `
    -EmailFrom "sas-monitor@bravehub.be" `
    -SASLogPath "E:\SAS\Logs"
```

## Installation

### Quick Start

1. **Download the scripts:**
   - Monitor-SASBatchErrors.ps1
   - Setup-SASMonitoring.ps1

2. **Run setup as Administrator:**
```powershell
# Basic installation
.\Setup-SASMonitoring.ps1 -InstallZabbix
```

3. **Verify installation:**
```powershell
# Test the monitoring script
C:\Scripts\Monitor-SASBatchErrors.ps1 -LogDirectory "E:\SAS\Logs" -DaysToCheck 1
```

### Manual Installation

1. **Create directories:**
```powershell
New-Item -Path "C:\Scripts" -ItemType Directory -Force
New-Item -Path "C:\Scripts\Logs" -ItemType Directory -Force
```

2. **Copy monitoring script:**
```powershell
Copy-Item ".\Monitor-SASBatchErrors.ps1" -Destination "C:\Scripts\" -Force
```

3. **Configure paths in the script** (if SAS logs are not in E:\SAS\Logs)

## Zabbix Integration

### Setup

1. **Install Zabbix configuration:**
```powershell
.\Setup-SASMonitoring.ps1 -InstallZabbix
```

2. **Restart Zabbix Agent:**
```powershell
Restart-Service "Zabbix Agent 2"
```

3. **Test from Zabbix server:**
```bash
zabbix_get -s <sas-server-hostname> -k sas.batch.errors.critical
```

### Zabbix Items

Add these items to your Zabbix template:

| Key | Type | Description | Trigger Condition |
|-----|------|-------------|-------------------|
| `sas.batch.errors.critical` | Numeric | Critical errors (7 days) | `last()>0` |
| `sas.batch.errors.critical.24h` | Numeric | Critical errors (24h) | `last()>0` |
| `sas.batch.errors.json` | Text | Full error report (JSON) | N/A |

### Sample Zabbix Triggers

```
# Critical SAS batch errors detected
Expression: {Template SAS Batch:sas.batch.errors.critical.last()}>0
Severity: High
Description: {ITEM.LASTVALUE} critical SAS batch errors detected in the last 7 days

# Multiple critical errors in 24h
Expression: {Template SAS Batch:sas.batch.errors.critical.24h.last()}>2
Severity: Disaster
Description: {ITEM.LASTVALUE} critical errors in the last 24 hours - immediate attention required
```

## Email Alerts

Email alerts are sent automatically when critical errors are detected.

**Configuration:**
```powershell
.\Setup-SASMonitoring.ps1 -EmailAlerts `
    -SmtpServer "smtp.yourdomain.com" `
    -EmailTo "admin@yourdomain.com" `
    -EmailFrom "sas-monitor@yourdomain.com"
```

**Alert Schedule:**
- Checks every hour
- Only sends email if critical errors are found
- Includes full error report in email body

## Scheduled Tasks

The setup creates two scheduled tasks:

### 1. Daily Error Report
- **Name:** SAS Batch Error Monitoring
- **Schedule:** Daily at 08:00
- **Action:** Generates text report of last 24 hours
- **Output:** `C:\Scripts\Logs\SAS_Error_Report_YYYYMMDD.txt`

### 2. Email Alerts (Optional)
- **Name:** SAS Batch Error Email Alerts
- **Schedule:** Hourly
- **Action:** Sends email if critical errors found

## Output Examples

### Text Format
```
================================================================================
SAS Batch Error Report - 2024-12-09 08:00:00
================================================================================

SUMMARY:
----------------------------------------
Total Errors: 3
Critical: 2
High: 1
Medium: 0

ERRORS BY TYPE:
----------------------------------------

[SpawnError] - 1 occurrence(s)
  Description: SAS.BAT spawning failure

  File: DWH051.13792.20251129063004.log
  Time: 2024-11-29 06:30:04
  Line: Can't spawn "sas.bat": No error at E:\Tools\Stargate\bin\\launchsas.pl line 192
  ---

[WorkLibAuth] - 1 occurrence(s)
  Description: WORK library authorization error

  File: J_DM_TRAN_SP_DIM_RISK_TYPE.13980.20251130225322.log
  Time: 2024-11-30 22:53:13
  Line: ERROR: User does not have appropriate authorization level for library WORK.
  ---

================================================================================
```

### JSON Format
```json
{
  "Timestamp": "2024-12-09T08:00:00.000Z",
  "TotalErrors": 3,
  "CriticalErrors": 2,
  "HighErrors": 1,
  "MediumErrors": 0,
  "Errors": [
    {
      "File": "DWH051.13792.20251129063004.log",
      "FullPath": "E:\\SAS\\Logs\\DWH051.13792.20251129063004.log",
      "LineNumber": 192,
      "Timestamp": "2024-11-29T06:30:04",
      "ErrorType": "SpawnError",
      "Severity": "Critical",
      "Description": "SAS.BAT spawning failure",
      "ErrorLine": "Can't spawn \"sas.bat\": No error"
    }
  ]
}
```

## Troubleshooting

### Script Doesn't Find Logs

1. **Verify log directory exists:**
```powershell
Test-Path "E:\SAS\Logs"
```

2. **Check for .log files:**
```powershell
Get-ChildItem "E:\SAS\Logs" -Recurse -Filter "*.log" | Select-Object -First 5
```

3. **Adjust time range:**
```powershell
# Check last 30 days instead of 7
.\Monitor-SASBatchErrors.ps1 -LogDirectory "E:\SAS\Logs" -DaysToCheck 30
```

### Zabbix Returns Error

1. **Test manually on SAS server:**
```powershell
C:\Scripts\Monitor-SASBatchErrors.ps1 -LogDirectory "E:\SAS\Logs" -OutputFormat Zabbix
```

2. **Check Zabbix Agent log:**
```powershell
Get-Content "C:\Program Files\Zabbix Agent 2\zabbix_agent2.log" -Tail 50
```

3. **Verify Zabbix Agent can run PowerShell:**
```bash
zabbix_get -s <hostname> -k system.run["powershell.exe -Command Get-Date"]
```

### Email Alerts Not Working

1. **Test SMTP connectivity:**
```powershell
Send-MailMessage -SmtpServer "smtp.bravehub.be" `
    -From "test@bravehub.be" `
    -To "sjoerd@bravehub.be" `
    -Subject "Test" `
    -Body "Test message"
```

2. **Check scheduled task:**
```powershell
Get-ScheduledTask -TaskName "SAS Batch Error Email Alerts" | Get-ScheduledTaskInfo
```

3. **Run email script manually:**
```powershell
C:\Scripts\Send-SASErrorAlert.ps1
```

## Maintenance

### Log Cleanup

Old reports should be cleaned up periodically:

```powershell
# Delete reports older than 90 days
Get-ChildItem "C:\Scripts\Logs\SAS_Error_Report_*.txt" | 
    Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-90) } | 
    Remove-Item
```

### Adding New Error Patterns

Edit `Monitor-SASBatchErrors.ps1` and add to `$ErrorPatterns`:

```powershell
$ErrorPatterns = @{
    # ... existing patterns ...
    'YourNewError' = @{
        Pattern = 'your regex pattern here'
        Severity = 'Critical|High|Medium'
        Description = 'Description of the error'
    }
}
```

## Support & Contact

For issues or questions:
- Email: sjoerd@bravehub.be
- Company: BraveHub
- Documentation: https://github.com/bravehub/sas-monitoring (if applicable)

## Version History

- **v1.0** (2024-12-09)
  - Initial release
  - Support for spawn errors and WORK library authorization errors
  - Zabbix integration
  - Email alerts
  - Scheduled task automation

## License

Internal use - BraveHub proprietary

---

**Note:** This solution is specifically designed for monitoring SAS batch jobs based on the errors encountered in November 2024 (29.11.2024 and 30.11.2024). Additional error patterns can be added as needed.
