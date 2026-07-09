# Network

Network and connectivity diagnostic scripts. Cross-platform where noted; the rest require Windows + Administrator.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Test-Ports.ps1`](#test-portsps1) | TCP port connectivity checker (cross-platform) |
| [`Test-AuthNetworkDiagnostics.ps1`](#test-authnetworkdiagnosticsps1) | Auth/network issue diagnostics (Event Viewer, Kerberos, DNS, shares) |
| [`Test-FileIODiagnostics.ps1`](#test-fileiodiagnosticsps1) | File I/O stress test with live failure diagnostics |

---

### Test-Ports.ps1

Tests TCP connectivity on one or more ports against one or more hosts. Supports single ports, ranges (`1294:1494`), and comma-separated combinations. Works on Windows (PS 5.1+), macOS, and Linux (PS 7+).

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-Target` | Yes | IP address or hostname(s) to test |
| `-Ports` | Yes | Port spec: single (`80`), range (`1294:1494`), list (`80,443,3389`), or mixed (`80,443,1294:1494`) |
| `-TimeoutMs` | No | TCP connect timeout in ms (default: `500`) |
| `-ShowClosed` | No | Include closed ports in output (default: only open ports shown) |

**Examples**

```powershell
.\Test-Ports.ps1 -Target 192.168.1.1 -Ports 1294:1494
.\Test-Ports.ps1 -Target 10.0.0.1 -Ports 80,443,3389,8080:8090
.\Test-Ports.ps1 -Target server01.contoso.local -Ports 22,3389 -TimeoutMs 1000 -ShowClosed
.\Test-Ports.ps1 -Target 10.0.0.1,10.0.0.2 -Ports 80,443
```

---

### Test-AuthNetworkDiagnostics.ps1

Diagnoses authentication and network issues on a Windows server/endpoint: scans Event Viewer for logon failures (4625), Kerberos errors (4771/4768), NTLM failures (4776), SMB/Netlogon/DNS client errors; checks time sync, Kerberos ticket cache, DNS resolution, TCP connectivity, UNC share access, and network adapter status; optionally scans log files for error patterns. Writes a txt report to `C:\Temp\`.

**Parameters**

| Parameter | Description |
|-----------|-------------|
| `-TestHosts` | Hostnames/IPs to test DNS resolution and TCP connectivity against |
| `-TestPorts` | TCP ports to test on each `-TestHosts` entry (default: `445, 88, 389, 636`) |
| `-TestShares` | UNC paths to test read access (e.g. `\\server\share`) |
| `-LogDirectory` | Optional directory to scan for auth/network error patterns |
| `-LogDaysBack` | Days back to include log files (default: `1`) |
| `-EventLogHours` | Hours back to check Event Viewer (default: `24`) |
| `-OutputPath` | Override the default output folder (`C:\Temp\`) |

**Examples**

```powershell
# Basic run — Event Viewer only
.\Test-AuthNetworkDiagnostics.ps1

# Test connectivity + UNC shares
.\Test-AuthNetworkDiagnostics.ps1 -TestHosts "dc01","fileserver" -TestShares "\\fileserver\data"

# Full scan including SAS log files
.\Test-AuthNetworkDiagnostics.ps1 -TestHosts "sasserver" -LogDirectory "E:\SAS\Logs"

# Check last 48 hours of events
.\Test-AuthNetworkDiagnostics.ps1 -EventLogHours 48
```

Requires Administrator.

---

### Test-FileIODiagnostics.ps1

Runs a write/append/read/delete loop against a target path and, on every failure, classifies the error (AUTH/NETWORK/TIMEOUT/DISK/PATH/IO) and captures diagnostic context: FileSystemWatcher events, an NTFS permission diff (`icacls`) against the startup baseline, open file handles via Sysinternals `Handle.exe` (auto-downloaded to `C:\Temp\handle\`), a new-process snapshot, Kerberos tickets, and Security event log entries. Stops after 3 failures.

**Parameters**

| Parameter | Description |
|-----------|-------------|
| `-TestPath` | Folder to test — local or UNC (e.g. `G:\sas\work`, `\\server\share\folder`) |
| `-Iterations` | Number of write/delete cycles (default: `5000`) |
| `-DelayMs` | Milliseconds between iterations (default: `50`) |
| `-StopOnFirstError` | Stop after the first failure instead of continuing to 3 |
| `-HandleExe` | Path to an existing `Handle.exe` — skips auto-download |
| `-SkipHandleDownload` | Do not attempt to download `Handle.exe` (e.g. air-gapped server) |
| `-LogPath` | Output folder for the log file (default: `C:\Temp\`) |

**Examples**

```powershell
# Basic run — downloads Handle.exe automatically
.\Test-FileIODiagnostics.ps1 -TestPath G:\sas\work

# Fast stress test, stop on first failure
.\Test-FileIODiagnostics.ps1 -TestPath G:\sas\work -Iterations 10000 -DelayMs 0 -StopOnFirstError

# Use an existing Handle.exe, no download
.\Test-FileIODiagnostics.ps1 -TestPath G:\sas\work -HandleExe C:\Tools\handle.exe

# Air-gapped server — skip the download
.\Test-FileIODiagnostics.ps1 -TestPath G:\sas\work -SkipHandleDownload
```
