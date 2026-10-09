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
| [`Watch-M365Apps.ps1`](Watch-M365Apps.ps1) ([docs](#watch-m365appsps1)) | Watchdog (scheduled task) — tests new Teams, new Outlook and Copilot with our own account (`itceadmin`) on a session host, repairs what is broken before a customer runs into it, and reports to n8n |

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
leaves FSLogix alone. It keeps every host at the newest build of the apps and of
everything they need, the same on every host.

**What it checks**

| Step | Details |
|------|---------|
| Windows | Edition (Enterprise multi-session), build, pending reboot |
| FSLogix | Build and `InstallAppxPackages` — read only |
| Updates | New Outlook updates itself **weekly from the Office CDN**, not through the Store, and has no switch to stop it: below FSLogix 25.06 the cure is running this script weekly on every host. Teams: below 2210 HF4 its self-update is turned off (`disableAutoUpdate = 1`) and every run updates it centrally; on a newer FSLogix it may update itself. The Store setting is shown, not changed. Edge Update policies that block WebView2 or Edge are reported |
| WebView2 | The Evergreen runtime all three apps render with, against the current Edge Stable build (`edgeupdates.microsoft.com`) |
| Apps | Teams, new Outlook, the Microsoft 365 Copilot app and the unified Copilot app: provisioned build, and users holding a newer build than the image provisions |
| Frameworks | Every `PackageDependency` in those apps' manifests (VCLibs, UI.Xaml, WindowsAppRuntime, …) must be on the machine at the `MinVersion` the manifest asks for — typically too old on an image from 2024 |
| Teams on AVD | `IsWVDEnvironment`, a Teams build new enough for SlimCore (`24193.1805.3040.8975`), the Teams Meeting add-in, and the WebRTC redirector: out of support since **1 October 2026**, stops working **1 April 2027**, kept only as a fallback for endpoints that cannot do SlimCore yet |
| Office | Shared Computer Activation (required on multi-session) and the update channel |
| Sign-in | `Microsoft.AAD.BrokerPlugin` present, and no FSLogix `redirections.xml` exclusion of `AppData\Local\Packages` or the app folders |
| Capture | With `-ForCapture`: packages installed for a user but not provisioned (Sysprep stops on them) and a pending reboot fail the check |

**What it updates** (every run without `-CheckOnly`, finding or not, in this order):
Teams' self-update setting where FSLogix needs it, Shared Computer Activation, WebView2
(Evergreen Standalone installer, signature-checked), then the apps to their newest build
through the existing scripts —
[`Repair-AppxPackageStore.ps1`](../Device/readme.md#repair-appxpackagestoreps1)
`-Name teams,outlook -Latest -Provision -RemoveOld` and `-Name copilot -Provision`, and
[`Update-TeamsClient.ps1`](../Device/readme.md#update-teamsclientps1) `-AvdOptimizations`
(newest Teams, meeting add-in, IsWVDEnvironment, WebRTC redirector). They only change what
is behind. Everything is read back afterwards.

The newest Outlook: Microsoft publishes no version feed for it, so `-Latest` takes the newest
build users on the host already updated to (Microsoft's installer otherwise). On a pool every
host then provisions what the most recent user got — exactly the build FSLogix will ask the
other hosts for.

**Parameters**

| Parameter | Description |
|-----------|-------------|
| `-CheckOnly` | Report only, change nothing. Exit code `2` means there is work to do |
| `-ForCapture` | The machine is the image VM about to be sysprepped: also fail on a pending reboot and on per-user packages that are not provisioned |
| `-SkipApps` | Do not call Repair-AppxPackageStore / Update-TeamsClient; only policies, Shared Computer Activation and WebView2 |
| `-ComputerName` | Session hosts to run on over PowerShell remoting; ends with one table across the pool and names every column that differs between hosts |
| `-Credential` | Credential for `-ComputerName` |
| `-WorkingDir` | Folder for downloads — WebView2 and, when they are not next to this script, the two helper scripts (default: `C:\IT\SessionHostImage`) |
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

- It never restarts the machine, and neither do the scripts and installers it calls (every
  msiexec runs with `/norestart`). A pending reboot is reported and left to you.
- Run elevated or as System. It relaunches itself in 64-bit Windows PowerShell, because
  the AppX cmdlets need it.
- Run it **weekly** on every host at once (a scheduled task as System works), and after each
  Windows update: Outlook moves weekly, and every host has to keep up with it.
- SlimCore also needs the tenant side: Teams VDI policy `VDI2Optimization` enabled, and
  Windows App 2.0.352.0 or newer on the endpoints. Neither can be checked from the host;
  `Update-TeamsClient.ps1 -CheckOnly` reads the session host's Teams VDI events, which do
  show whether users are really on SlimCore.
- The two scripts it calls come from the repo when it runs from there, or from the folder
  `-ComputerName` copies them to. Run on its own — only this file on the image VM — it fetches
  them from GitHub at a **pinned commit** of `main` and runs them only when their SHA-256
  matches, the same way [`Invoke-FSLogixShrink.ps1`](#invoke-fslogixshrinkps1) handles Invoke-FslShrinkDisk. To
  move up: put the new commit and the hashes of both files in `$HelperCommit` /
  `$HelperHashes` at the top of the script, after reading the diff.
- `-ComputerName` copies this script and the two it calls to `C:\IT\SessionHostImage`
  on each host.

---

### Watch-M365Apps.ps1

A watchdog for new Teams, new Outlook and Copilot on a session host. It runs as a
scheduled task under System and uses **our own account** — `itceadmin` by default — as a
canary: when an app does not start for it, it will not start for a customer either. The watchdog repairs the host before a customer notices, and reports
to an n8n webhook.

**Each run**

| Step | What happens |
|------|--------------|
| 0. Crashes | Every crash (Application Error `1000`) and every hang that ended in a close (Application Hang `1002`) of Teams, new Outlook or Copilot since the previous run, from the Application log — for **every user on the host**, customers included — grouped per app, module and exception code. Reported, not repaired: a customer's app is never restarted for them |
| 1. Host | Teams and new Outlook are provisioned for all users; Copilot is there (provisioned MicrosoftOfficeHub / Copilot, or the unified app Edge Update installs) |
| 2. Accounts | For each watched account signed in on this host: the package is registered for that user, its files are there and its status is `Ok`. Then the app has to run in that session — if it does not, it is started there (`shell:AppsFolder\<AUMID>`, through a one-off task in that user's own session) and has to still be running 15 seconds later |
| 3. Repair | Host problems and packages whose files are gone: [`Repair-AppxPackageStore.ps1`](../Device/readme.md#repair-appxpackagestoreps1) `-Provision` (Microsoft's installers, signature-checked), at most once per `-RepairCooldownHours`. Then, in **our own account only**: a package that is not registered is registered by family name, an app that does not start is reset (`Reset-AppxPackage`) |
| 4. Read back | Steps 1 and 2 again |
| 5. Report | A JSON POST to the webhook when something is wrong, was repaired, recovered on its own, or crashed at least `-CrashThreshold` times — not on every healthy run. A problem that stays is reported again after `-RenotifyHours` |

Nothing is closed, removed or reset for any other user: no `-RemoveOld`, no `-Latest`,
no stopping of customers' processes.

**Parameters**

| Parameter | Description |
|-----------|-------------|
| `-Account` | Accounts to test with — user name, UPN or `DOMAIN\user` (default: `itceadmin`; more than one tests each, e.g. `itceadmin,itce.user`) |
| `-App` | `Teams`, `Outlook`, `Copilot` (default: all three) |
| `-WebhookUrl` | n8n webhook (production URL) the report is POSTed to. Without it the run only logs |
| `-WebhookToken` | Sent as header `X-Watchdog-Token`; match it with Header Auth on the n8n Webhook node |
| `-IntervalMinutes` | How often the task runs (default: `30`) |
| `-RepairCooldownHours` | Minimum time between two host repairs, so a problem it cannot fix is not retried every run (default: `4`) |
| `-RenotifyHours` | Report a problem that stays the same again after this many hours (default: `12`) |
| `-CrashThreshold` | Report the crashes and hangs of one app once there are this many since the previous run (default: `1`, every crash; `0` turns crash reporting off) |
| `-NoRepair` | Test and report only, change nothing |
| `-SkipLaunchTest` | Do not start an app that is not running; only check its registration |
| `-Install` | Copy the watchdog to `-WorkingDir` and register the task **M365 App Watchdog**, with the other parameters as its settings |
| `-Uninstall` | Remove the task and `-WorkingDir` |
| `-TestNotification` | Send one test message to the webhook and stop |
| `-WorkingDir` | Watchdog, settings, state and logs (default: `C:\IT\AppWatchdog`) |

**Examples**

```powershell
# One run now in this console, report only
.\Watch-M365Apps.ps1 -NoRepair

# Install on a session host, reporting to n8n
.\Watch-M365Apps.ps1 -Install -WebhookUrl 'https://n8n.example.com/webhook/m365-apps' -WebhookToken '<token>' -Confirm:$false

# Check the webhook end to end with the installed settings
.\Watch-M365Apps.ps1 -TestNotification

# Remove it again
.\Watch-M365Apps.ps1 -Uninstall -Confirm:$false
```

**What n8n receives**

```json
{
  "source": "Watch-M365Apps",
  "event": "repaired",
  "host": "AVD-0",
  "time": "2026-10-09T14:30:02.1234567+02:00",
  "summary": "AVD-0: 1 problem(s) found and repaired - itceadmin Outlook NotRegistered",
  "accounts": ["itceadmin"],
  "findings": [],
  "before": [{ "Account": "itceadmin", "App": "Outlook", "Problem": "NotRegistered", "Detail": "not registered for this user" }],
  "actions": ["itceadmin Outlook: re-registered as the user - result 0"],
  "crashes": [{ "App": "Teams", "Kind": "Crash", "Count": 2, "Last": "2026-10-09T14:12:40.0000000+02:00", "Exe": "ms-teams.exe", "Version": "26260.1704.5188.5238", "Module": "msedgewebview2.dll", "Code": "0xc0000005" }],
  "log": "C:\IT\AppWatchdog\Logs\Watch-M365Apps_20261009.log"
}
```

`event` is `repaired`, `repair-failed`, `failing` (with `-NoRepair`), `crashed` (only
crashes this run), `recovered`, `error` (the run itself failed) or `test`; `crashes` rides
along with any of them. `Problem` is `NotProvisioned` (host), `NotRegistered`,
`Broken` (files gone or status not `Ok`) or `WontStart`. In n8n: a **Webhook** node (POST,
Header Auth on `X-Watchdog-Token`), then route on `{{$json.body.event}}` to Teams, mail or
a ticket.

**Installing from GitHub**

A session host needs no copy of the repo: download this one file and run `-Install`.
The helper comes along by itself — `-Install` fetches `Repair-AppxPackageStore.ps1` from
GitHub at the pinned commit and checks its SHA-256. The download below is pinned the same
way, to a commit and a hash, because what it installs runs as System. Run it in an
elevated PowerShell on the host:

```powershell
# Watch-M365Apps.ps1 at a fixed commit - move both lines together
$commit = '960576b2119a4ac147366f136a6bd1fcbff5726c'
$sha256 = '38EECD24C9EDD6BC3D52128619CDA2A40E8FEA8ADB80BB4C18CAB02E725D7F7B'
$file   = Join-Path $env:TEMP 'Watch-M365Apps.ps1'
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
Invoke-WebRequest "https://raw.githubusercontent.com/sjkanon/M365-Scripts/$commit/scripts/RDS/Watch-M365Apps.ps1" -OutFile $file -UseBasicParsing
if ((Get-FileHash $file -Algorithm SHA256).Hash -ne $sha256) { Remove-Item $file; throw 'SHA-256 does not match - not run' }
# then, as in the examples: first -NoRepair to look, then install
& $file -Install -WebhookUrl '<n8n webhook URL>' -WebhookToken '<token>' -Confirm:$false
```

On several hosts at once, from your own machine over PowerShell remoting:

```powershell
# Same download on every host, in parallel
Invoke-Command -ComputerName avd-0, avd-1, avd-2 -ScriptBlock {
    $commit = '960576b2119a4ac147366f136a6bd1fcbff5726c'
    $sha256 = '38EECD24C9EDD6BC3D52128619CDA2A40E8FEA8ADB80BB4C18CAB02E725D7F7B'
    $file   = Join-Path $env:TEMP 'Watch-M365Apps.ps1'
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest "https://raw.githubusercontent.com/sjkanon/M365-Scripts/$commit/scripts/RDS/Watch-M365Apps.ps1" -OutFile $file -UseBasicParsing
    if ((Get-FileHash $file -Algorithm SHA256).Hash -ne $sha256) { Remove-Item $file; throw "SHA-256 mismatch on $env:COMPUTERNAME" }
    & $file -Install -WebhookUrl '<n8n webhook URL>' -WebhookToken '<token>' -Confirm:$false
}
```

Afterwards check one host with `-TestNotification`. To move to a newer version: put that
commit and the file's SHA-256 at that commit
(`(Get-FileHash .\scripts\RDS\Watch-M365Apps.ps1).Hash` in a checkout of it) in the two
lines, after reading the diff, and run `-Install` again — the task keeps running its own
copy in `C:\IT\AppWatchdog` until then. The webhook URL and token never go in this repo:
it is public.

**Notes**

- **Keep a session of `itceadmin` (and any other watched account) open on every host** (disconnected is fine). An
  account that is not signed in is skipped: its packages live in its FSLogix container
  and cannot be tested without it. Without any session the host check (step 1) still runs.
- Crash events name no user, so a crash can be a customer's or ours. A crash of our own
  account's app is followed up in step 2: the app is no longer running, so it is started
  again. On a busy pool where the odd Teams crash is noise, raise `-CrashThreshold`.
- The launch test starts an app that is not running in our own session. A window can
  appear there, and a reset Teams asks our account to sign in again. `-SkipLaunchTest`
  turns it off.
- `-Install` locks `-WorkingDir` to System and Administrators (the task runs what is in it
  as System), copies this script and `Repair-AppxPackageStore.ps1` there — from
  `..\Device` in a checkout of the repo, otherwise always from GitHub (never from beside a
  downloaded copy, a folder any user may have written to) at the same pinned commit and SHA-256 as
  [`Update-SessionHostImage.ps1`](#update-sessionhostimageps1) — and keeps the webhook URL and
  token only in `config.json` there, not in the task's command line. Change a setting by
  running `-Install` again with all parameters.
- Logs: `C:\IT\AppWatchdog\Logs`, one file per day, kept 14 days. Repair transcripts and
  `.reg` backups: `C:\IT\AppWatchdog\Repair`.
- Run elevated or as System; it relaunches itself in 64-bit Windows PowerShell for the
  AppX cmdlets. Exit code `0` healthy or repaired, `1` something still broken.
