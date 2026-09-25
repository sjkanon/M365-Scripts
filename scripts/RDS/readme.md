# RDS

Diagnostic and monitoring scripts for RDP / RD Web Access infrastructure. Run directly on the RDS/RDWeb server for full results — remote targets only get connectivity-level checks.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Test-RDSDiagnostics.ps1`](Test-RDSDiagnostics.ps1) ([docs](#test-rdsdiagnosticsps1)) | One-shot health check — services, config, certs, user account, event logs |
| [`Watch-RDSLive.ps1`](Watch-RDSLive.ps1) ([docs](#watch-rdsliveps1)) | Real-time session + licensing event monitor |
| [`Get-FSlogix-errors.ps1`](Get-FSlogix-errors.ps1) ([docs](#get-fslogix-errorsps1)) | FSLogix / Azure Files profile diagnostics on an AVD session host |

---

### Test-RDSDiagnostics.ps1

Diagnoses why users cannot log in to an RDP or RD Web Access server.

**Checks performed**

| Area | Details |
|------|---------|
| RDP server | `TermService`/`SessionEnv`/`UmRdpService` status, RDP enabled/disabled, NLA, session limits, RD Licensing mode, Remote Desktop Users group, firewall rules, active sessions (`quser`) |
| RDWeb server | Port 443 reachability, HTTPS certificate validity/expiry, IIS + RDWeb app pool (local only), RD Gateway service (local only) |
| User account (optional, `-Username`) | Enabled/locked/expired, group membership, logon workstation restrictions, last logon, password age |
| Event logs (optional, `-IncludeEventLogs`) | Security 4625 (failed RDP logon), 4740 (lockout), `TerminalServices-LocalSessionManager` 20/40 (session failure/disconnect reason) |

**Parameters**

| Parameter | Description |
|-----------|-------------|
| `-RdpServer` | Hostname/IP of the RDP/Terminal Server. Omit to run local checks |
| `-RdWebServer` | Hostname/IP of the RD Web Access server |
| `-Username` | Check a specific account for login blockers (requires ActiveDirectory module or ADSI fallback) |
| `-LogPath` | Folder for the output log file (default: `C:\Temp\`) |
| `-IncludeEventLogs` | Include event log analysis for the last `-Hours` hours |
| `-Hours` | Hours of event log history to analyse (default: `24`) |

**Examples**

```powershell
# Full check — RDP + RDWeb + user account
.\Test-RDSDiagnostics.ps1 -RdpServer rdp01.company.local -RdWebServer rdweb.company.local -Username jdoe -IncludeEventLogs

# Run locally on the RDS host, check event logs
.\Test-RDSDiagnostics.ps1 -IncludeEventLogs -Hours 48
```

Results are written to the console and a timestamped log file in `C:\Temp\`.

---

### Watch-RDSLive.ps1

Polls Windows event logs every N seconds and streams new events to the console + log file. Run directly on each RDS/RDWeb server.

**Events monitored**

| Source | Events |
|--------|--------|
| `TerminalServices-LocalSessionManager` | Logon (21), reconnect (22/25), logoff (23), disconnect (24), logon failed (20), disconnect reason (40) — human-readable reason codes |
| Security | Failed RDP logon (4625, type 10), account lockout (4740) |
| `TerminalServices-Licensing` | License granted/denied/warning events |
| System | `TermServLicensing` provider (grace period, license server errors) |

Prints a heartbeat line per poll with the active session count.

**Parameters**

| Parameter | Description |
|-----------|-------------|
| `-IntervalSeconds` | Polling interval in seconds (default: `20`) |
| `-LogPath` | Folder for the output log file (default: `C:\Temp\`) |
| `-NoLogFile` | Console output only, skip the log file |

**Examples**

```powershell
# Run on the RDS server
.\Watch-RDSLive.ps1

# Faster polling, no log file
.\Watch-RDSLive.ps1 -IntervalSeconds 10 -NoLogFile
```

> Press `Ctrl+C` to stop. Run as Administrator for Security log access.

---

### Get-FSlogix-errors.ps1

Collects, in one run, everything needed to work out why an FSLogix profile will not
mount on an AVD session host — mount errors, a locked VHDX, SMB/Azure Files trouble or
disk errors. Read-only: it gathers and reports, it repairs nothing.

**What it collects**

| Area | Details |
|------|---------|
| System | Host name, OS build, uptime |
| FSLogix | Installed version, the full `Profiles`/`Containers` configuration, and the service state |
| Containers | Attached VHD(X) files, the FSLogix session registry, and the profile paths from `ProfileList` |
| Storage | SMB connections to Azure Files, and whether the VHD share is reachable at all |
| Events | FSLogix events over the last `-Days` days, matched against the known critical failure patterns, plus disk/NTFS errors and User Profile Service events |
| Leftovers | Local profiles under `C:\Users` and the FSLogix log files |

**Parameters**

| Parameter | Description |
|-----------|-------------|
| `-User` | Also produce a section filtered to one user — the account whose profile is failing |
| `-Days` | Days of event history to analyse (default: `7`) |
| `-OutputPath` | Folder for the transcript (default: `%SystemDrive%\Temp\FSLogixDiag`) |

**Examples**

```powershell
# Everything from the last week
.\Get-FSlogix-errors.ps1

# One user, two weeks back, report somewhere else
.\Get-FSlogix-errors.ps1 -User jdoe -Days 14 -OutputPath C:\Temp
```

> Run in an elevated session **on the session host itself** — the container, SMB and
> event data only exist there. The whole run is written to
> `FSLogixDiag_<host>_<timestamp>.log` in the output folder, which is the file to
> attach to a ticket.
