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
| [`Get-M365AppsLog.ps1`](Get-M365AppsLog.ps1) ([docs](#get-m365appslogps1)) | Collects, read only, what happened with new Teams, new Outlook and Copilot on a session host — per user what is registered and running, the events, and what the watchdog did and would conclude, with its blind spots — into one zip |

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
canary: when an app does not start for it, it will not start for a customer either. It also checks every other signed-in user and every app a user could not open,
repairs it on the host and in that user's own session before they call, and reports to
an n8n webhook — naming the user each finding and each repair belongs to.

> Service desk version for IT Glue (Dutch, per support level — what to do when a user calls, how to read the Teams cards): [Watch-M365Apps-ITGlue.md](Watch-M365Apps-ITGlue.md), or the styled [Watch-M365Apps-ITGlue.html](Watch-M365Apps-ITGlue.html) to paste in.

**Each run**

| Step | What happens |
|------|--------------|
| 0. Crashes | Every crash (Application Error `1000`) and every hang that ended in a close (Application Hang `1002`) of Teams, new Outlook or Copilot since the previous run, from the Application log — for **every user on the host**, customers included — grouped per app, module and exception code. Reported, not repaired: a customer's app is never restarted for them |
| 1. Host | Teams and new Outlook are provisioned for all users; Copilot is the **new, unified Microsoft Copilot app** only (`copilotapp.exe`, installed machine-wide by Edge Update) — the old Microsoft 365 Copilot app (`MicrosoftOfficeHub`) no longer counts. Its `copilotapp.exe` has to be on disk and a Start menu entry for all users has to point at it; a missing entry is created, a missing app installed by the host repair (`Repair-AppxPackageStore.ps1 -Name copilot -Provision`, through Edge Update) |
| 1a. Update | Every `-UpdateHours` (default 6): the **newest** Teams and new Outlook are provisioned on the host — [`Repair-AppxPackageStore.ps1`](../Device/readme.md#repair-appxpackagestoreps1) `-Latest -Provision` (Teams from Microsoft's config service, Outlook the newest build seen on this host or in a profile; only ever newer, signature-checked, no `-RemoveOld`) — and Edge Update is asked to check now for the unified Copilot app. Every run: when **our own account** has an older build than the host provisions, the app is closed in our session and registered again from the provisioned build, so step 2 starts the new build — a build that does not run is found by us, not by a customer at their next sign-in. **Customers are never updated by the watchdog**: Windows gives them the new build at their next sign-in. A new build on the host — from this step or on its own — is reported once (`updated`, with the old and new version) |
| 2. Accounts | For each watched account signed in on this host: the package is registered for that user, its files are there and its status is `Ok` — for Copilot its identity package, when the host has one (its files are machine-wide). Then the app has to run in that session — if it does not, it is started there (`shell:AppsFolder\<AUMID>`, or `copilotapp.exe` itself, through a one-off task in that user's own session) and has to still be running 15 seconds later |
| 2b. Users | Every other signed-in user (customers), once signed in for 10 minutes: the same registration check, **without starting anything**. Plus every attempt since the previous run, by any user, to open one of the apps that Windows refused (TWinUI `5961`), and every failed registration of their packages (AppXDeploymentServer `401`/`404`; "close the app first" and "already installed" are left out), with the user it happened to |
| 2c. Announce | Something new is wrong: [`Get-M365AppsLog.ps1`](#get-m365appslogps1) collects the evidence for the users and apps concerned into `Diag\` (one zip, kept 14 days) while it is still broken, and a `repairing` report goes to n8n naming every user — before anything is changed. The same problem again is not announced or collected again within `-RenotifyHours` |
| 3. Repair | A problem of **one customer is fixed in their session only**; the host is left alone. The host is repaired for an app only when it is more than that one user — the host itself, **our own account** (`itceadmin` is the test for the whole host), or the same app at two or more customers — without starting anything for anyone: [`Repair-AppxPackageStore.ps1`](../Device/readme.md#repair-appxpackagestoreps1) `-Provision` (Microsoft's installers, signature-checked), at most once per `-RepairCooldownHours`. Then **per user, in their own session**: the package is registered again by family name (`Add-AppxPackage -RegisterByFamilyName`) through a one-off task running a headless console, so no window appears. For a customer only while the app is not running for them, at most once per `-RepairCooldownHours` per user and app, and never a reset. In our own account an app that does not start is reset (`Reset-AppxPackage`). A user who tried to open the app **themselves** (TWinUI `5961`) in the last 30 minutes gets it **opened for them** in their session once it is registered again — it has to start and stay up, otherwise the problem stays open. An attempt in the first 2 minutes after sign-in is the app's autostart and does not count, and nothing else is ever started for a customer (`-NoUserLaunch` turns this off) |
| 4. Read back | Steps 1, 2 and the registration check of 2b again; a user's problem counts as repaired when the package is registered and `Ok` for them afterwards (or the app is running for them) |
| 5. Report | A JSON POST to the webhook when something is wrong, was repaired, recovered on its own, or crashed at least `-CrashThreshold` times — not on every healthy run. A problem that stays is reported again after `-RenotifyHours` |

Nothing is closed or removed for a customer, and a customer's app is never reset: no
`-RemoveOld`, `-Latest` only for the host in step 1a, no stopping of customers' processes. Every report lists, per
finding, the user (`Account`, with `Customer` true for a customer) and, under `before`,
whether it was repaired for them (`Fixed`); the Teams card shows *hersteld bij deze
gebruiker* next to each one.

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
| `-UpdateHours` | How often the newest Teams / Outlook are provisioned and Edge Update checks for Copilot (default: `6`; `0` turns updating off, for our own account as well; `-NoRepair` also updates nothing) |
| `-CrashThreshold` | Report the crashes and hangs of one app once there are this many since the previous run (default: `1`, every crash; `0` turns crash reporting off) |
| `-NoRepair` | Test and report only, change nothing |
| `-NoUserRepair` | Repair the host and our own account, but never run anything in a customer's session — their problems are still reported, with their name |
| `-SkipLaunchTest` | Do not start an app that is not running; only check its registration |
| `-NoUserLaunch` | After repairing an app a user could not open, do not open it for them |
| `-NoDiagnostics` | Do not run `Get-M365AppsLog.ps1` before a repair |
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
  "log": "C:\IT\AppWatchdog\Logs\Watch-M365Apps_20261009.log",
  "diagnostics": "C:\IT\AppWatchdog\Diag\M365AppsLog_AVD-0_20261009-1430.zip"
}
```

`event` is `repairing` (found, being repaired — sent before the repair, with `diagnostics`), `repaired`, `repair-failed`, `failing` (with `-NoRepair`), `crashed` (only
crashes this run), `recovered`, `error` (the run itself failed) or `test`; `crashes` rides
along with any of them. `Problem` is `NotProvisioned` (host), `NotRegistered`,
`Broken` (files gone or status not `Ok`) or `WontStart` (our account),
`WontOpen` (Windows refused to open it for a user) or `RegisterFailed`. In n8n: a **Webhook** node (POST,
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
$commit = '82807f82f2e5bbc108c24ed967070dc2345bea8d'
$sha256 = 'FDC21FC632BC5D06EDFCD730919A688D2899788A6C064B4B1B4F42E13C364E8E'
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
    $commit = '82807f82f2e5bbc108c24ed967070dc2345bea8d'
    $sha256 = 'FDC21FC632BC5D06EDFCD730919A688D2899788A6C064B4B1B4F42E13C364E8E'
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
- Repairing in a customer's session uses `conhost --headless`, so no console or Windows
  Terminal window should appear; that was checked to wait for the command and to drop its
  exit code (hence the read-back), but **not** yet looked at in a real customer session.
- `Get-AppxPackage -User` is given `DOMAIN\user`, not a SID: with an Entra ID SID
  (`S-1-12-1-…`) it answers *No valid SID could be determined*.
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
- Evidence: `C:\IT\AppWatchdog\Diag`, one zip per new problem from `Get-M365AppsLog.ps1`, kept 14 days; its path is in `diagnostics` of the report. `-Install` copies the collector next to the watchdog (from the repo, or from GitHub at a pinned commit and SHA-256); without it the watchdog repairs and reports as before.
- Logs: `C:\IT\AppWatchdog\Logs`, one file per day, kept 14 days. Repair transcripts and
  `.reg` backups: `C:\IT\AppWatchdog\Repair`.
- Run elevated or as System; it relaunches itself in 64-bit Windows PowerShell for the
  AppX cmdlets. Exit code `0` healthy or repaired, `1` something still broken.

---

### Get-M365AppsLog.ps1

Collects everything about new Teams, new Outlook and Copilot on a session host into one
folder and zip, to answer two questions after a complaint: *why did the app not work for
this user*, and *why did the watchdog ([`Watch-M365Apps.ps1`](#watch-m365appsps1)) not
catch or repair it*. Read only — nothing is started, registered, repaired or closed.

**What it collects**

| Part | What |
|------|------|
| Watchdog | `config.json` (webhook and token masked), whether `NoRepair` / `NoUserRepair` is on and the app is watched, the installed version and whether it has the per-user checks at all; the task's state, last run, result and history, probe tasks left behind; `state.json` (last run, last host repair, open problems, per-user re-registrations); its logs and repair logs from the window, with the lines about the apps and users in question pulled out |
| Host | Provisioned builds, the Copilot unified app (Edge Update) and the WebView2 runtime |
| All users | `Get-AppxPackage -AllUsers`: every user each package is known for and its install state — signed-off users included |
| Signed-in users | Per user and app: the packages registered for them, status, files present, running in their session, and the verdict the watchdog would reach — marked **BLIND SPOT** where it would call it fine without testing it |
| Events | From the last `-Hours`: TWinUI `5960`/`5961`, AppXDeploymentServer (errors and `401`/`404`), AppXDeployment, AppxPackaging, AppReadiness, AppModel-Runtime, Application Error `1000` / Hang `1002` of the apps — each with the user, the error code, and whether the watchdog reads that event at all |
| Registry | A non-zero `PackageStatus` (Windows marked the package bad), packages in `Deprovisioned`, FSLogix `InstallAppxPackages`, the Edge Update Copilot policies; exports of the FSLogix, Edge Update policy, Teams and Deprovisioned keys |
| FSLogix and Edge Update | Their log files from the window (FSLogix honours `Logging\LogDir`), the Profile log lines about the apps that carry an error, FSLogix error and warning events |
| Copilot | Installed Copilot apps with their folder, and which Copilot process runs for whom — `copilotapp.exe` is the unified app, `M365Copilot.exe` the packaged one |
| App logs (`-IncludeAppLogs`) | Per signed-in user in scope: Teams logs (LocalCache — gone at sign-out with FSLogix), the Teams diagnostics bundle in Downloads, new Outlook logs, FSLogix `AppxPackages.xml` |

Output in `-OutputPath` (falls back to `%TEMP%`): `summary.txt` (the screen), `sessions.csv`,
`packages.csv`, `allusers.csv`, `events.csv`, `fslogix-events.csv`, `watchdog\`,
`registry\`, `fslogix\`, `edgeupdate\` and `applogs\`, zipped.

**Blind spots it points out**

- With a watchdog from before 2026-10-10 (8): Copilot not registered for a user while the
  unified app is on the host is accepted for everyone without a test, the unified app
  (`copilotapp.exe`) is never started, older ones do not see its crashes either, and
  consumer Copilot (`Microsoft.Copilot`) counts as much as Microsoft 365 Copilot. From
  that version on the watchdog counts only the unified app, checks its identity package
  per user and starts it for our own account — the collector says which one it found.
- A customer's app registered and `Ok` but not working: the watchdog starts nothing for
  customers, so it only sees an open that Windows refused (TWinUI `5961`).
- A user signed in less than 10 minutes, a user who is signed off, the watched account
  not signed in on the host, a task that has not run, an older watchdog without the
  per-user checks, events the watchdog does not read.

**Parameters**

| Parameter | Description |
|-----------|-------------|
| `-User` | Only these users — user name, UPN or `DOMAIN\user`, also when signed off (default: everyone) |
| `-App` | `Teams`, `Outlook`, `Copilot` (default: all three) |
| `-Hours` | How far back to read events and logs (default: `24`, at most `336`) |
| `-OutputPath` | Where the folder and zip are written (default: `C:\Temp`) |
| `-WorkingDir` | The watchdog's folder (default: `C:\IT\AppWatchdog`) |
| `-IncludeAppLogs` | Also copy the apps' own logs per signed-in user — they hold names and mail addresses, so only on request |
| `-NoZip` | Leave the folder, do not zip it |

**Examples**

```powershell
# A customer says Copilot did not open this morning
.\Get-M365AppsLog.ps1 -User jansen -App Copilot -Hours 12

# Everything on this host from the last two days, with the Teams and Outlook logs
.\Get-M365AppsLog.ps1 -Hours 48 -IncludeAppLogs
```

**Notes**

- Run elevated on the host, soon after the complaint; it relaunches itself in 64-bit
  Windows PowerShell for the AppX cmdlets. Menu item `Q`.
- Every part runs on its own: one that fails is named at the end and the rest is still
  collected. `Get-AppxPackage` runs in a separate process and is given up after 120
  seconds, because it can hang on exactly the host this is for. An event channel that
  does not exist on the host (no FSLogix, an older build) is skipped.
- Logs that are still open are read shared; a log over 50 MB keeps its last 50 MB.
- The webhook URL and token are masked in the copy of `config.json`.
- Comparable to the AppX and FSLogix parts of Microsoft's MSRD-Collect, but limited to
  these three apps and set against what the watchdog does.
