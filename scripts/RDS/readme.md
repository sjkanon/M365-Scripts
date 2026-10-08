**English** · [Nederlands](readme.nl.md) · [Français](readme.fr.md)

[M365-Scripts](../../readme.md) › [scripts](../readme.md) › **RDS**

# RDS

Diagnostic, monitoring and preparation scripts for RDP / RD Web Access and AVD session hosts. Run directly on the RDS/RDWeb server for full results — remote targets only get connectivity-level checks.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Test-RDSDiagnostics.ps1`](Test-RDSDiagnostics.ps1) ([docs](#test-rdsdiagnosticsps1)) | One-shot health check — services, config, certs, user account, event logs |
| [`Watch-RDSLive.ps1`](Watch-RDSLive.ps1) ([docs](#watch-rdsliveps1)) | Real-time session + licensing event monitor |
| [`Get-FSlogix-errors.ps1`](Get-FSlogix-errors.ps1) ([docs](#get-fslogix-errorsps1)) | FSLogix / Azure Files profile diagnostics on an AVD session host |
| [`Invoke-FSLogixShrink.ps1`](Invoke-FSLogixShrink.ps1) ([docs](#invoke-fslogixshrinkps1)) | Shrink FSLogix profile disks on a share (Invoke-FslShrinkDisk), or check whether FSLogix compacts them itself at sign-out |
| [`Update-SessionHostImage.ps1`](Update-SessionHostImage.ps1) ([docs](#update-sessionhostimageps1)) | Check and prepare a Windows 11 multi-session image or AVD session host so new Teams, new Outlook and Copilot keep working with FSLogix — FSLogix itself is left alone |

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

---

### Invoke-FSLogixShrink.ps1

Gives the space back that FSLogix profile and ODFC containers keep after data inside
them was deleted: a dynamic VHDX grows but never shrinks on its own. It is a wrapper
around [Invoke-FslShrinkDisk](https://github.com/FSLogix/Invoke-FslShrinkDisk), the
FSLogix team's own shrink script — still the best tool for this; the forks and
alternatives on GitHub do the same with less behind them.

**What it does**

| Step | Details |
|------|---------|
| Download | Fetches Invoke-FslShrinkDisk at a **pinned commit** to `C:\Scripts\Invoke-FslShrinkDisk`, unblocks it and checks the SHA-256 of the script. A changed upstream version is never run unseen; a tampered local copy is refused |
| Report | Every `.vhd`/`.vhdx` on the share, largest first, with folder, size and last write, and the total |
| Shrink | Runs Invoke-FslShrinkDisk recursively, then summarises its CSV log: disks shrunk, GB recovered, and the disks it could not process |
| `-CheckHost` | On a session host: can FSLogix's **built-in compaction at sign-out** run? FSLogix version (2210 / 2.9.8361 or later), `VHDCompactDisk`, the Optimize Drives service (`defragsvc` not Disabled) and dynamic disks |

**Parameters**

| Parameter | Description |
|-----------|-------------|
| `-Path` | The share holding the containers, e.g. `\\<storageaccount>.file.core.windows.net\<share>\Profiles`. Searched recursively |
| `-ReportOnly` | Only list the disks and their size; shrink nothing |
| `-IgnoreLessThanGB` | Skip disks smaller than this (default: `5`) |
| `-RatioFreeSpace` | Only shrink a disk with at least this fraction free inside (default: `0.1` = 10%) |
| `-ThrottleLimit` | Disks processed at the same time (default: `4`; at most twice the CPU cores) |
| `-LogFilePath` | CSV log (default: `C:\Temp\FslShrink_<timestamp>.csv`); the folder is created when missing |
| `-ToolPath` | Where Invoke-FslShrinkDisk is kept (default: `C:\Scripts\Invoke-FslShrinkDisk`) |
| `-Force` | Download Invoke-FslShrinkDisk again |
| `-CheckHost` | Check FSLogix's own compaction on this host instead; needs no `-Path` |

**Examples**

```powershell
# Look first: every container on the share, largest first
.\Invoke-FSLogixShrink.ps1 -Path \\sa.file.core.windows.net\profiles\Profiles -ReportOnly

# Shrink everything of 5 GB or more with at least 10% free inside
.\Invoke-FSLogixShrink.ps1 -Path \\sa.file.core.windows.net\profiles\Profiles

# Does FSLogix compact the disks itself on this host?
.\Invoke-FSLogixShrink.ps1 -CheckHost
```

**Notes**

- FSLogix 2210 and later compacts a container itself at every sign-out, when the disk is
  over 1 GB and at least 20% can be won ([Microsoft Learn](https://learn.microsoft.com/en-us/fslogix/concepts-vhd-disk-compaction)).
  Run `-CheckHost` first: when that passes, a manual shrink only catches up disks of users
  who rarely sign out, or below the 20% threshold.
- A disk that is attached — the user is signed in — cannot be shrunk; it shows up in the
  summary as not processed and the run ends with exit code 1. Run it out of hours or with
  the hosts drained.
- Run elevated (each disk is mounted), with access to the share: on Azure Files through
  Kerberos or the storage account key. Hyper-V is not needed.
- To move to a newer Invoke-FslShrinkDisk: put the new commit and the SHA-256 of its
  `Invoke-FslShrinkDisk.ps1` in `$ToolCommit` / `$ToolHash` at the top of the script,
  after reading the diff.

---

### Update-SessionHostImage.ps1

Teams, new Outlook and Copilot are MSIX apps, and on a pooled AVD host with FSLogix they
break in one recurring way: a user's app updates itself on host A, FSLogix saves that
exact version in the profile at sign-out, and at the next sign-in on host B — which does
not have that version — registering it fails with `0x80070490`. FSLogix 2210 HF4 (Teams)
and 25.06 (Outlook) register by package family instead, but this script deliberately
leaves FSLogix alone. It makes the image carry everything the apps need at one build,
and stops the apps from drifting away from it per user.

**What it checks**

| Step | Details |
|------|---------|
| Windows | Edition (Enterprise multi-session), build, pending reboot |
| FSLogix | Build and `InstallAppxPackages` — read only. Below 25.06 Outlook is re-registered at its exact saved version, so the hold-back below is what keeps it working |
| Hold-back | Microsoft Store `AutoDownload = 2` and Teams `disableAutoUpdate = 1`, so the apps only change with the image. Edge Update policies that block WebView2 or Edge are reported |
| WebView2 | The Evergreen runtime all three apps render with, against the current Edge Stable build (`edgeupdates.microsoft.com`) |
| Apps | Teams, new Outlook, the Microsoft 365 Copilot app and the unified Copilot app: provisioned build, and users holding a newer build than the image provisions |
| Frameworks | Every `PackageDependency` in those apps' manifests (VCLibs, UI.Xaml, WindowsAppRuntime, …) must be on the machine at the `MinVersion` the manifest asks for — typically too old on an image from 2024 |
| Teams on AVD | `IsWVDEnvironment`, a Teams build new enough for SlimCore (`24193.1805.3040.8975`), the Teams Meeting add-in, and the WebRTC redirector: out of support since **1 October 2026**, stops working **1 April 2027**, kept only as a fallback for endpoints that cannot do SlimCore yet |
| Office | Shared Computer Activation (required on multi-session) and the update channel |
| Sign-in | `Microsoft.AAD.BrokerPlugin` present, and no FSLogix `redirections.xml` exclusion of `AppData\Local\Packages` or the app folders |
| Capture | With `-ForCapture`: packages installed for a user but not provisioned (Sysprep stops on them) and a pending reboot fail the check |

**What it fixes** (without `-CheckOnly`, in this order): the two hold-back policies,
Shared Computer Activation, WebView2 (Evergreen Standalone installer, signature-checked),
then the apps through the existing scripts —
[`Repair-AppxPackageStore.ps1`](../Device/readme.md#repair-appxpackagestoreps1)
`-Name teams,outlook -Latest -Provision -RemoveOld` and `-Name copilot -Provision`, and
[`Update-TeamsClient.ps1`](../Device/readme.md#update-teamsclientps1) `-AvdOptimizations`.
Everything is read back afterwards.

**Parameters**

| Parameter | Description |
|-----------|-------------|
| `-CheckOnly` | Report only, change nothing. Exit code `2` means there is work to do |
| `-ForCapture` | The machine is the image VM about to be sysprepped: also fail on a pending reboot and on per-user packages that are not provisioned |
| `-SkipApps` | Do not call Repair-AppxPackageStore / Update-TeamsClient; only policies, Shared Computer Activation and WebView2 |
| `-ComputerName` | Session hosts to run on over PowerShell remoting; ends with one table across the pool and names every column that differs between hosts |
| `-Credential` | Credential for `-ComputerName` |
| `-WorkingDir` | Folder for downloads (default: `C:\IT\SessionHostImage`) |
| `-LogPath` | Folder for the transcript of a run that changes something (default: `C:\Temp`) |

**Examples**

```powershell
# What does this host or image need? Changes nothing.
.\Update-SessionHostImage.ps1 -CheckOnly

# All three session hosts side by side, read only
.\Update-SessionHostImage.ps1 -ComputerName avd-0,avd-1,avd-2 -CheckOnly

# Prepare the image VM, then check it is fit for capture
.\Update-SessionHostImage.ps1 -Confirm:$false
.\Update-SessionHostImage.ps1 -CheckOnly -ForCapture

# Bring all three hosts level (drain them first)
.\Update-SessionHostImage.ps1 -ComputerName avd-0,avd-1,avd-2 -Confirm:$false
```

**Notes**

- Run elevated or as System. It relaunches itself in 64-bit Windows PowerShell, because
  the AppX cmdlets need it.
- With the hold-back on, Teams and Outlook only update when this script (or a new image)
  updates them: run it monthly, on every host at once, after the Windows update.
- SlimCore also needs the tenant side: Teams VDI policy `VDI2Optimization` enabled, and
  Windows App 2.0.352.0 or newer on the endpoints. Neither can be checked from the host;
  `Update-TeamsClient.ps1 -CheckOnly` reads the session host's Teams VDI events, which do
  show whether users are really on SlimCore.
- `-ComputerName` copies this script and the two it calls to `C:\IT\SessionHostImage`
  on each host.
