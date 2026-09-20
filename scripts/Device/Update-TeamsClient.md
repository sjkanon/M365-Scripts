# Update-TeamsClient.ps1 — how it works

Reference for [`Update-TeamsClient.ps1`](Update-TeamsClient.ps1): what it decides, in which order, and why it is built the way it is. For the short version (parameters and examples) see the [Device readme](readme.md#update-teamsclientps1). For the service desk there is a Dutch, support-level-oriented version to paste into IT Glue: [Update-TeamsClient-ITGlue.md](Update-TeamsClient-ITGlue.md).

---

## Table of Contents

- [What it is for](#what-it-is-for)
- [The decision](#the-decision)
- [The nine steps](#the-nine-steps)
- [AVD / VDI](#avd--vdi)
- [Classic Teams](#classic-teams)
- [The meeting add-in: machine-wide versus per user](#the-meeting-add-in-machine-wide-versus-per-user)
- [The version check](#the-version-check)
- [Output modes](#output-modes)
- [Exit codes](#exit-codes)
- [Parameters](#parameters)
- [Running it by hand](#running-it-by-hand)
- [Running it from NinjaOne](#running-it-from-ninjaone)
- [Design decisions](#design-decisions)
- [Troubleshooting](#troubleshooting)
- [What has been tested](#what-has-been-tested)

---

## What it is for

Keeping the **new Teams client** (the `MSTeams` MSIX/AppX package) and the **Teams Meeting Add-in for Outlook** current on a Windows endpoint or an AVD session host, without reinstalling devices that are already current.

It replaces the common "uninstall Teams, download the bootstrapper, reinstall, install the add-in" one-off script with something that can be scheduled: it first asks Microsoft which build is published, and only acts when the device is behind.

Three ways to run it:

| Use | Command |
|-----|---------|
| Look before you leap | `.\Update-TeamsClient.ps1 -WhatIf` |
| Update when needed | `.\Update-TeamsClient.ps1` |
| Scheduled, unattended | `.\Update-TeamsClient.ps1 -Quiet -Confirm:$false` |

---

## The decision

Everything the script does hangs off facts gathered read-only in steps 1 and 2: **is the client behind**, **is the add-in present**, and — only with `-AvdOptimizations` — **are the AVD components in place**.

```
   is the installed build behind the published one?
   is the meeting add-in installed?
   are the AVD components in place?   (-AvdOptimizations only)
                       |
   +-------------------+-------------------+
   |                   |                   |
 behind            current,            everything
 (or -Force)       something missing   in place
   |                   |                   |
 full path         only that part      nothing at all
 steps 3-8         (3 and/or 7)        exit 0, no output
                                       in -Quiet mode
```

`-Force` short-circuits the comparison and always takes the full path — that is the repair mode. `-CheckOnly` stops right after the decision and reports it (exit code `2` when there is work to do).

---

## The nine steps

| # | Step | What happens | `-WhatIf` |
|---|------|--------------|-----------|
| 1 | Preflight | Inventory every place Teams can live: `*MSTEAMS*` AppX packages per user, the provisioned package, classic Teams (machine-wide installer and per-profile installs), the meeting add-in in both uninstall hives, whether Outlook has it registered, the AVD components, running Teams/Outlook | read-only |
| 2 | Version check | Ask the Teams config service for the published build, compare, decide | read-only |
| 3 | AVD | Only with `-AvdOptimizations`: set `IsWVDEnvironment`, install the WebRTC redirector | guarded |
| 4 | Classic | Only with `-RemoveClassicTeams`: uninstall the Teams Machine-Wide Installer, clear the per-profile installs | guarded |
| 5 | Download | Create the working folder, download `teamsbootstrapper.exe`, check size and Authenticode signature | guarded |
| 6 | Uninstall | `msiexec /x` the add-in **and every other copy of it** (machine-wide folder, per-profile folders, per-user COM registrations), `Remove-AppxPackage -AllUsers`, `Remove-AppxProvisionedPackage` | guarded |
| 7 | Install | `teamsbootstrapper.exe -p` (provision for all users) | guarded |
| 8 | Add-in | Locate `MicrosoftTeamsMeetingAddinInstaller.msi` inside the installed package, install it with `ALLUSERS=1` | guarded |
| 8b | Repair | Only with `-RepairOutlookAddIn`: clear a per-user registration that points at a removed DLL, put `LoadBehavior` back to 3 | guarded |
| 9 | Verify | Re-read the add-in registration machine-wide **and** per signed-in user in Outlook, the provisioned package, the classic removal and the AVD components, compare against the published build | reported as skipped |

"Guarded" means the step is wrapped in `$PSCmdlet.ShouldProcess(...)`, so under `-WhatIf` it prints what it would do and changes nothing.

Only what is actually missing gets done: steps 5–7 are skipped when the client is current, step 8 when the add-in is already there and the client was not replaced, step 3 unless `-AvdOptimizations` is given and step 4 unless `-RemoveClassicTeams` is given.

---

## AVD / VDI

`-AvdOptimizations` adds the two things a session host needs for Teams media optimization, both from the original VDI gap-fill script this replaces:

| Component | What it is | When it is installed |
|-----------|------------|----------------------|
| `HKLM:\SOFTWARE\Microsoft\Teams\IsWVDEnvironment` (DWORD `1`) | Tells Teams to hand media to the redirector instead of rendering it in the session | When not already `1` |
| Remote Desktop WebRTC Redirector Service | MSI from `https://aka.ms/msrdcwebrtcsvc/msi` (~1.7 MB, signed by Microsoft) | When not already installed, or with `-Force`. An installed redirector is **not** silently upgraded — `-Force` is what replaces it |

The flag is step 3, before the client is provisioned, because Teams reads it at startup to decide which media path to use.

Both are only touched when missing, so a scheduled run on a fully configured session host still downloads nothing and, with `-Quiet`, prints nothing. Without the switch the script does not change any of this — it only points out that the device looks like a session host (`HKLM:\SOFTWARE\Microsoft\RDInfraAgent` exists).

### WebRTC is on its way out

Microsoft is retiring the WebRTC-based media optimization: **end of support 1 October 2026, end of availability 1 April 2027**. Teams shows users a banner about it. Its replacement, **SlimCore**, needs nothing installed on the session host — it ships inside new Teams and inside Windows App on the endpoint. Verified on a device with Teams `26225.1806.5074.1452`:

```
Microsoft.Teams.SlimCoreVdiHost.win-x64          2026.31.1.16
Microsoft.Teams.SlimCoreVdiFwk.win-x64.2026.31   2026.31.1.16   (plus older framework versions)
```

Preflight reports that package when `-AvdOptimizations` is used, but it is **informational only**: which path is actually taken depends on the Windows App version on the endpoint the user connects from, and a script on the session host cannot see that. Auditing endpoint client versions is the real migration work.

`IsWVDEnvironment` stays required either way, and Microsoft's current guidance is to keep the redirector installed as a fallback for endpoints that cannot do SlimCore — so `-AvdOptimizations` keeps installing it. Revisit before April 2027.

---

## Classic Teams

`-RemoveClassicTeams` removes the old client as well. Off by default: taking an application away from users is not a decision an update job should make on its own, and the switch keeps that decision with whoever schedules the job.

| What | How | Why it matters |
|------|-----|----------------|
| Teams Machine-Wide Installer | `msiexec /x <ProductCode> /qn /norestart` | The important one — while it is installed, Windows keeps staging classic Teams into every new user profile |
| Per-profile install root | Remove `%LOCALAPPDATA%\Microsoft\Teams` | Where the per-user copy actually lives |
| Autostart entry | Remove `Run\com.squirrel.Teams.Teams` from that user's hive | Otherwise Windows tries to start something that is gone |
| Stale uninstall key | Remove `Uninstall\Teams` from that user's hive | Keeps Programs and Features honest |

The documented per-user uninstall is `Update.exe --uninstall -s`, but that has to run as the profile owner, which System cannot do — so the files are removed instead. Roaming data in `%APPDATA%\Microsoft\Teams` is left alone; it is inert once the client is gone.

One case gets its own treatment. `msiexec /x` answering **1605** ("this action is only valid for products that are currently installed") means the entry in Programs and Features outlived the product — common once the new Teams bootstrapper has been over the machine. There is nothing to uninstall, but the stale entry would keep this script reporting classic Teams forever, so the registry entry is removed instead and the run carries on.

Failure handling splits the rest deliberately:

- A **machine-wide installer that survives** the uninstall is a real failure → `[FAIL]`, exit `1`.
- A **per-profile folder that survives** is almost always a file lock from a running classic Teams → `[WARN]`, and the next run clears it once the user has signed out. The script warns up front when it sees `Teams.exe` running.

Classic Teams reached end of support on **1 October 2026** (end of availability 1 April 2027), the same date as the WebRTC optimisation, which is why both show up in the same inventory.

---

## The meeting add-in: machine-wide versus per user

The script installs the add-in MSI with `ALLUSERS=1` into `%ProgramFiles(x86)%\Microsoft\TeamsMeetingAddin\<version>` — a **per-machine** install, which is what Microsoft prescribes for a shared machine or an AVD/RDS session host so that every user who logs on has it.

That is not the only way the add-in gets onto a device. On an ordinary endpoint the Teams client installs and updates it **per user**, by itself. Measured on a Windows 11 endpoint with Teams `26225.1806.5074.1452`:

| What | Where |
|------|-------|
| Add-in files | `%LOCALAPPDATA%\Microsoft\TeamsMeetingAdd-in\1.26.21803` (note the hyphen) |
| MSI the client installed it from | `%LOCALAPPDATA%\Microsoft\TeamsMeetingAddinMsis\1.26.21803\` — seven cached versions, one per Teams update since April |
| Outlook COM registration | `HKCU\SOFTWARE\Microsoft\Office\Outlook\Addins\TeamsAddin.FastConnect`, `LoadBehavior=3` |
| The DLL Outlook actually loads | `HKCU\SOFTWARE\Classes\CLSID\{19A6E644-14E6-4A60-B8D7-DD20610A871D}\InprocServer32` → `%LOCALAPPDATA%\Microsoft\TeamsMeetingAdd-in\<version>\x64\Microsoft.Teams.AddinLoader.dll` |
| Machine-wide Outlook registration | none — no Teams entry under `HKLM\...\Office\Outlook\Addins` at all |
| Uninstall entry | `HKLM` **64-bit** hive, product code `{A7AB73A3-...}`, `InstallSource` pointing at that per-user MSI cache |

Three consequences worth knowing:

- **Outlook loads add-ins per user.** A machine-wide install makes the add-in *available* to everyone; each user's Outlook still picks it up on its next start. That is why the script warns when Outlook is running.
- **What the verification proves.** Step 8 reads the `HKLM` uninstall keys *and* asks whether Outlook itself has the add-in registered, for every signed-in user: `HKEY_USERS\<sid>\SOFTWARE\Microsoft\Office\Outlook\Addins\TeamsAddin.FastConnect`, plus the machine-wide equivalent under `HKLM`. `LoadBehavior 3` means Outlook loads it at startup; `2` or `0` means Outlook switched it off, which is the real "the button is gone" case and needs a human. A profile nobody is signed into cannot be read at all — not broken, just unseen — so none of this is treated as a failure.
- **A per-user registration outranks the machine-wide one.** `HKCU\SOFTWARE\Classes` wins over `HKLM\SOFTWARE\Classes` for that user, so a user who already has the client-installed copy keeps loading the DLL from their own profile even after a machine-wide install lands in `Program Files (x86)`. Which also means the two are not interchangeable: the machine-wide install does not repair a broken per-user one.
- **A profile that is not signed in is not broken, just unreadable.** Its hive is not mounted, so nothing can be said about it. With a healthy machine-wide registration in place that profile picks the add-in up the first time that user starts Outlook — which is why "not loaded on every profile yet" is usually a matter of waiting rather than a fault. The script lists those profiles by name instead of leaving them out, because an absent profile and a healthy one look identical in the output otherwise.
- **A dangling registration is what `LoadBehavior 2` usually means.** If that `InprocServer32` path no longer exists — the profile copy was removed while the registration stayed — Outlook tries, fails and switches the add-in off again. Ticking the box back on does not survive that; the add-in has to be installed again for that user. The script resolves the CLSID per signed-in user and reports the two cases apart, because they need different fixes.
- **A reinstall sweeps every copy first.** The MSI clears one. The machine-wide folder, the per-profile folders under `%LOCALAPPDATA%\Microsoft\TeamsMeetingAdd-in` and the per-user COM registrations are removed with it, because a copy left behind is exactly what becomes a shadowing registration pointing at files that no longer exist. `-RepairOutlookAddIn` exists for the devices where that already happened.
- **Repairing it means removing the shadow, not adding another install.** `-RepairOutlookAddIn` deletes the user''s stale `Classes\CLSID\{19A6E644-...}` key so COM resolves to the machine-wide registration again, and puts `LoadBehavior` back to 3. It only acts when that machine-wide registration is healthy — clearing the shadow with nothing behind it would leave the user worse off. Off by default: it writes into another user''s hive.
- **`-SkipMeetingAddIn` is defensible on normal endpoints.** There the client keeps the add-in current on its own; the machine-wide install is what session hosts and shared machines need.

> The uninstall entry landed in the **64-bit** hive here, not in `WOW6432Node`. It is not fixed which one it is, which is exactly why both are scanned.

---

## The version check

The source is the same feed the Teams client itself uses to decide it is out of date:

```
https://config.teams.microsoft.com/config/v1/MicrosoftTeams/0.0.0.0
    ?environment=prod&audienceGroup=general&teamsRing=general&agent=TeamsBuilds
```

The relevant part of the response:

```json
{
  "BuildSettings": {
    "WebView2PreAuth": {
      "x64": {
        "latestVersion": "26225.1806.5074.1452",
        "buildLink": "https://teamsinstaller.public.onecdn.static.microsoft/production-windows-x64/26225.1806.5074.1452/MSTeams-x64.msix"
      },
      "x86":   { "latestVersion": "..." },
      "arm64": { "latestVersion": "..." }
    }
  }
}
```

- The architecture is derived from `PROCESSOR_ARCHITECTURE` (with a `PROCESSOR_ARCHITEW6432` fallback for a WOW64 process) and maps to `x64`, `x86` or `arm64`.
- `WebView2PreAuth` holds the Windows builds; the older `WebView2` key is checked as a fallback (it still carries macOS).
- `-Ring` replaces both `audienceGroup` and `teamsRing`, for tenants on a non-default update ring.
- The comparison is a `[version]` comparison against the installed AppX package version, falling back to the **provisioned** package version when no user has Teams installed yet. That fallback matters: on a pooled session host or a fresh image, Teams is often only provisioned, and without it the version check has nothing to compare, declares the host outdated and reinstalls ~275 MB on every scheduled run.
- **Equal or newer means nothing to do** — a device on an insider build is left alone rather than downgraded.
- If the service cannot be reached, the run **stops** instead of reinstalling blindly. `-Force` overrides that.

---

## Output modes

| Mode | Behaviour |
|------|-----------|
| default | Every step prints as it happens |
| `-Quiet` | Output is held in memory and only flushed when there is news: a newer build, an action about to be taken, or a failure. An up-to-date device produces **zero bytes** |
| `-WhatIf` | Prints the plan, changes nothing, exit `0` |
| `-CheckOnly` | Prints the version comparison and stops, exit `2` if there is work |

Line prefixes: `[ OK ]` fact confirmed, `[NEW ]` news (an update is available), `[SKIP]` deliberately not done, `[WARN]` worth knowing, not fatal, `[FAIL]` the run failed.

A transcript (`C:\Temp\Update-TeamsClient_<timestamp>.log`, override with `-LogPath`) is started **only** once the run has decided to change something — a scheduled check on an up-to-date fleet leaves no log litter.

---

## Exit codes

| Code | Meaning |
|------|---------|
| `0` | Success, already up to date, or cancelled at the confirmation prompt |
| `1` | Failure: no rights, no version info, download/signature refused, installer failed or timed out, verification failed |
| `2` | `-CheckOnly` only: there is work to do (newer build, or missing add-in) |

---

## Parameters

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-WhatIf` | — | Walk the flow, change nothing |
| `-Quiet` | — | Print nothing unless there is news |
| `-CheckOnly` | — | Report the decision and stop (exit `2` when work is due) |
| `-Confirm:$false` | — | Never ask the interactive confirmation |
| `-Ring` | `general` | Update ring queried at the config service |
| `-WorkingDir` | `C:\IT\AVD\Teams` | Where the bootstrapper is downloaded |
| `-LogPath` | `C:\Temp` | Transcript folder |
| `-BootstrapperUrl` | Microsoft fwlink | Download URL, https only |
| `-WebRtcUrl` | `aka.ms/msrdcwebrtcsvc/msi` | WebRTC redirector MSI, https only |
| `-SkipMeetingAddIn` | — | Leave the add-in alone; a missing add-in is then not "work" |
| `-SkipSignatureCheck` | — | Accept an installer not signed by Microsoft (internal mirror) |
| `-TimeoutSeconds` | `900` | Per-process timeout for msiexec and the bootstrapper |
| `-Force` | — | Reinstall even when current; also allows running without Teams or without version info |

---

## Running it by hand

```powershell
powershell -ExecutionPolicy Bypass -File .\Update-TeamsClient.ps1 -WhatIf
```

- **Elevation**: not elevated and interactive, the script asks for UAC and continues in a new elevated window that stays open (`-NoExit`), forwarding the same parameters. Not elevated and non-interactive, it fails with exit `1` instead of hanging on a prompt.
- **Confirmation**: an interactive run that is about to change something asks once (`Continue? [y/N]`). Answering anything else exits `0` with nothing changed. `-Confirm:$false` skips the question.
- **From the menu**: `menu.ps1` key `T`. It asks whether to install or only check, and passes `-WhatIf` or `-Confirm:$false` accordingly, so the question is never asked twice.

---

## Running it from NinjaOne

> Setting it up for the first time (script fields, script variables, scheduled automation, detection job) is written out step by step in [Update-TeamsClient-ITGlue.md](Update-TeamsClient-ITGlue.md#bijlage--het-script-in-ninjaone-zetten-eenmalig-level-3).

1. Add the script — Language **PowerShell**, OS **Windows**, Architecture **All**, Run As **System**.
2. **Preview on one device first**: Parameters `-WhatIf -Confirm:$false`. The job output shows the version comparison and every step an update would take; the device is untouched.
3. **Schedule the real run**: Parameters `-Quiet -Confirm:$false`. Up-to-date devices print nothing and exit `0`, so the activity feed only shows devices where something actually happened.
4. **Detection / condition job**: Parameters `-CheckOnly -Quiet`. Silent and `0` when current, output and exit `2` when a newer build is published.

Script variables are read from the environment when the matching parameter is not passed on the command line, so a technician can tick a checkbox instead of typing parameters:

| Variable | Type | Maps to |
|----------|------|---------|
| `whatIf` | checkbox | `-WhatIf` |
| `quiet` | checkbox | `-Quiet` |
| `checkOnly` | checkbox | `-CheckOnly` |
| `force` | checkbox | `-Force` |
| `avdOptimizations` | checkbox | `-AvdOptimizations` |
| `removeClassicTeams` | checkbox | `-RemoveClassicTeams` |
| `repairOutlookAddIn` | checkbox | `-RepairOutlookAddIn` |
| `skipMeetingAddIn` | checkbox | `-SkipMeetingAddIn` |
| `skipSignatureCheck` | checkbox | `-SkipSignatureCheck` |
| `workingDir` | text | `-WorkingDir` |
| `logPath` | text | `-LogPath` |
| `ring` | text | `-Ring` |
| `webRtcUrl` | text | `-WebRtcUrl` |

Accepted truthy values: `true`, `1`, `yes`. Capitalisation of the variable name does not matter — environment lookups are case-insensitive on Windows, so a Ninja variable named `Quiet` or `QUIET` is picked up just as well. A parameter passed on the command line always wins over the environment.

> Always add `-Confirm:$false` for unattended runs. The confirmation is gated on `[Environment]::UserInteractive`, which should be false under the agent, but the flag makes it impossible for the question to appear at all.

---

## Design decisions

Why the script looks the way it does — most of these are scars from a real failure mode.

**Download before uninstall.** The obvious order (remove Teams, then fetch the installer) leaves a device with no Teams client at all when the download fails, the URL is blocked by a proxy, or the CDN is having a day. The installer is fetched *and* signature-checked first; only then is anything removed.

**Both uninstall hives.** The meeting add-in's uninstall entry does not always land in the same place — on a Windows 11 endpoint with add-in 1.26.21803 it was the 64-bit hive, other builds put it under `HKLM:\SOFTWARE\WOW6432Node\...`. A lookup in one hive alone silently finds nothing, and then the uninstall is skipped and the final verification reports failure on a perfectly good install.

**Relaunch 64-bit.** An RMM agent may start PowerShell 32-bit. Under WOW64 the HKLM reads are redirected to `WOW6432Node` and `$env:ProgramFiles` points at `Program Files (x86)`, so neither the AppX package nor the add-in MSI is found. The script re-executes itself through `%WINDIR%\SysNative\WindowsPowerShell\v1.0\powershell.exe` with the same parameters before doing any work.

**MSI version over COM, not AppLocker.** Microsoft's own sample reads the add-in version with `Get-AppLockerFileInformation`. That module is absent on some editions, and under PowerShell 7 it loads through the Windows PowerShell compatibility layer — in testing it failed outright and flooded a `-WhatIf` run with unrelated "Copy File" lines. The script reads `ProductVersion` straight from the MSI property table with the `WindowsInstaller.Installer` COM object.

**Timeouts, 1618, 3010.** Every `msiexec`/bootstrapper call goes through one helper that waits with a timeout and kills the process when it expires, so an RMM job can never block the agent. Exit code `1618` (another installation in progress) is retried twice with a 30-second pause; `3010` counts as success and flags a pending reboot in the summary.

**Deprovisioning as well as removing.** Removing the AppX package for all users leaves the *provisioned* copy in the image, so every new user profile keeps getting the old version staged. `Remove-AppxProvisionedPackage` is part of the uninstall step.

**Errors abort, they do not continue.** `$ErrorActionPreference = 'Stop'` plus one `try/catch/finally` around the whole flow: a failure stops the run rather than pressing on with a half-removed client. Deliberate stops (up to date, check-only, cancelled) travel the same path but carry a planned exit code, so they are reported as `[SKIP]`, not as failures.

**AVD behind a switch, not autodetected.** A session host needs the `IsWVDEnvironment` flag and the WebRTC redirector; a normal endpoint needs neither, and setting the flag there tells Teams to hand media to a redirector that is not present. Autodetecting AVD would make that a silent, machine-dependent side effect, so it is an explicit switch — the script only *mentions* that a device looks like a session host.

**The redirector MSI does not upgrade itself.** It keeps one ProductCode across versions, so `msiexec /i` over an existing install answers `1638` ("another version of this product is already installed") instead of upgrading — which is exactly what a `-Force` run on a configured session host hit. The script now compares the downloaded ProductVersion with what is registered: same version means a repair (`REINSTALL=ALL REINSTALLMODE=vomus`), a different version means the old one is uninstalled first, and a `1638` that still slips through is reported as "leaving the existing one in place" instead of failing the run.

**SlimCore is not ours to install.** An earlier version of this script warned "SlimCore packages not found" on session hosts, which was wrong twice over: the packages are staged on the endpoint by the plugin, not on the VM, and no script can install them there on the endpoint''s behalf. The check now reports whichever side it is standing on, and on an endpoint it looks for the policies that actually stop the staging. That mistake is worth remembering: a check that looks in the wrong place does not fail, it lies.

**Provisioning is not installing.** `teamsbootstrapper.exe -p` stages the package for future sign-ins; it does not install it for whoever ran the script. A first production run on a session host proved the point: the bootstrapper reported success and the add-in step then died on `New Teams package not found after install`, because `Get-AppxPackage -Name MSTeams` looks at the current user and the admin running the script had no Teams. The add-in MSI is therefore located by globbing `%ProgramFiles%\WindowsApps\MSTeams_*_x64__8wekyb3d8bbwe\MicrosoftTeamsMeetingAddinInstaller.msi` and taking the newest version, which works whether or not any user has the package installed. The same blind spot applied to the SlimCore check, which now asks `-AllUsers` first.

**No SID whitelist.** Enumerating user profiles by matching `S-1-5-21-*` looks right and silently breaks on Entra-joined devices, where user SIDs are `S-1-12-1-*`. Measured on this machine: the whitelist skipped the only real user, and the Outlook check reported "nobody has it registered" while `LoadBehavior=3` was sitting right there. The filter excludes the service SIDs (`S-1-5-18/19/20`), `.DEFAULT` and the `_Classes` hives instead.

**Read-only checks ignore -WhatIf.** `New-PSDrive` supports `ShouldProcess`, so mounting `HKEY_USERS` as a drive meant that under `-WhatIf` the drive was never created and the Outlook check falsely reported nothing registered. An inspection must give the same answer in a dry run as in a real one, so the hives are addressed through `Registry::HKEY_USERS` directly, with no drive to create.

**Output buffering.** `-Quiet` exists because a fleet-wide scheduled job that prints on every device makes the one device that needs attention invisible. Lines are held in a list and flushed at the first sign of news; if the run ends with nothing to report, they are simply dropped.

---

## Troubleshooting

| Symptom | Cause and fix |
|---------|---------------|
| `Administrator rights are required` | Run as System (RMM) or from an elevated session. Interactive sessions get a UAC prompt instead |
| `Could not enumerate AppX packages` | Same thing: `Get-AppxPackage -AllUsers` needs elevation |
| `Could not determine the latest published build` | `config.teams.microsoft.com` is unreachable (proxy, firewall, no DNS). Use `-Force` to reinstall without the check |
| `Bootstrapper signature is NotSigned/HashMismatch` | The download was intercepted or is a proxy error page. Check `-BootstrapperUrl` and the proxy; `-SkipSignatureCheck` only for a deliberate internal mirror |
| `Downloaded file is only N bytes` | Same cause — a captive portal or error page instead of the installer |
| `Uninstall of Teams Machine-Wide Installer failed (exit code 1605)` | Should no longer happen: 1605 means Windows Installer does not know the product, so the leftover Programs and Features entry is removed instead |
| `WebRTC Redirector install failed (exit code 1638)` | Should no longer happen: the old version is uninstalled first. If it does, the redirector is registered under a version msiexec disagrees with - remove it by hand from Programs and Features and re-run |
| `timed out after 900 seconds and was killed` | A hung msiexec or a slow image. Raise `-TimeoutSeconds`; check whether another installation is running |
| `No add-in MSI found under ...\WindowsApps` | The bootstrapper did not stage a package. Check the step 7 output and `C:\Program Files\WindowsApps` for an `MSTeams_*` folder |
| Add-in installed but not visible in Outlook | Outlook has to restart. The script warns when Outlook is running during the install |
| Nothing happens at all | That is the point: the device is current. Run with `-WhatIf` (without `-Quiet`) to see the comparison |

---

## What has been tested

Verified on a Windows 11 device with Teams `26225.1806.5074.1452` and add-in `1.26.21803` installed, in a non-elevated session with the elevation check and the admin-only AppX reads stubbed:

| Scenario | Result |
|----------|--------|
| Up to date, default | Reports current, `[SKIP] Teams is up to date`, exit `0` |
| Up to date, `-Quiet` | 0 bytes of output, exit `0` |
| Up to date, `-CheckOnly -Quiet` | 0 bytes of output, exit `0` |
| Newer build published, `-CheckOnly -Quiet` | Version comparison + `[NEW ]`, exit `2` |
| Newer build published, `-WhatIf` | Full plan for steps 3–7, exit `0` |
| Current build, add-in missing, `-WhatIf` | Steps 3–5 skipped, add-in install planned, exit `0` |
| `-Force` on a current device, `-WhatIf` | Full reinstall planned, exit `0` |
| `-AvdOptimizations` on a device without either component, `-WhatIf` | Only step 3 planned (flag + redirector download + install); steps 4-7 skipped, exit `0` |
| `-AvdOptimizations -CheckOnly -Quiet`, components missing | `Work is due: the AVD optimizations are incomplete`, exit `2` |
| `-AvdOptimizations -Quiet`, everything already in place | 0 bytes of output, exit `0` |
| WebRTC redirector MSI from `aka.ms/msrdcwebrtcsvc/msi` | 1.7 MB, Authenticode `Valid`, signed by `O=Microsoft Corporation` - passes the signature check |
| `-BootstrapperUrl http://...` | Refused before any change, exit `1` |
| Version check via the live config service | Returned `26225.1806.5074.1452`, matching the installed build |
| Only provisioned, no per-user install (pooled session host) | Reports `Provisioned MSTeams <version>`, compares correctly, does nothing, exit `0` |
| Outlook add-in registration on an Entra-joined device | Found `AzureAD\<user>` with `LoadBehavior 3`; the earlier `S-1-5-21` whitelist found nobody |
| Outlook check under `-WhatIf` | Same answer as a real run, after dropping `New-PSDrive` |
| Classic Teams detection | Neither variant is present on the test device, so detection's positive path is **untested** against a real install |
| Classic removal, `-WhatIf` | With detection faked, plans the msiexec uninstall and the per-profile folder removal, skips steps 5-8, exit `0` |
| Classic removal, applied | Against a faked profile folder in a scratch directory: folder removed, verification reports `No per-user classic Teams left`, exit `0` |
| Classic removal on a real session host | `-Force -RemoveClassicTeams -AvdOptimizations` on an AVD host removed a real per-profile classic Teams `1.4.00.11161`. The machine-wide msiexec path is still **untested** - that host had no Machine-Wide Installer |
| Add-in step after provisioning | Broke on that same production run (`Get-AppxPackage` is per user, provisioning is not) and now resolves the MSI from the staged `WindowsApps` package instead. Re-tested here for a per-user install, a provisioned-only host and a `-Force` run |
| Outlook add-in switched off | That production run found `LoadBehavior 2` for one account - the check earns its keep on the first real device it saw |
| Profiles that cannot be read | Listed by name with what it means for them: with a machine-wide registration they pick it up at the next Outlook start, without one there is nothing to fall back on. Both messages verified against stubbed profile lists |
| Add-in sweep before reinstall | `Get-TeamsAddInFolder` verified against this device (finds the real per-profile copy), and the `-WhatIf` plan shows the folder plus both CLSID views being removed before the reinstall. The removal itself reuses mechanics proven live in the classic-Teams and repair tests; the sweep as a whole runs for the first time on a production host |
| Per-user registration repair | Planted a stale CLSID (both registry views) plus `LoadBehavior 2` against a healthy machine-wide registration: `-WhatIf` planned both actions, an applied run cleared the keys and set `LoadBehavior` to 3, exit `0`. The real fix on a production host is still to be confirmed |
| Dangling add-in registration | Measured here: the COM class resolves to `%LOCALAPPDATA%\Microsoft\TeamsMeetingAdd-in\1.26.21803\x64\Microsoft.Teams.AddinLoader.dll`, a targeted lookup across loaded hives costs ~100 ms, and a planted registration pointing at a missing DLL is reported as "re-enabling will not stick" |
| SlimCore reporting | Verified on both sides: on this endpoint it finds `Microsoft.Teams.SlimCoreVdiHost.win-x64 2026.31.1.16`; with `RDInfraAgent` faked it reports the session-host wording instead. The three blocker policies were exercised against stubbed registry reads |
| Redirector replacement | A second session host failed on `1638` with redirector `1.54.2408.19001` installed while `aka.ms` now serves `1.56.2603.20001`. Both paths are now planned correctly under `-WhatIf` (replace: `/x` then `/i`; same version: `REINSTALL=ALL`), but **neither msiexec call has been run for real** |
| Stale machine-wide classic entry | That same host then failed on `1605` from the classic uninstall. Tested live: a real `msiexec /x` against an unknown product code returns 1605, the run warns, removes the stale registry entry, continues and verifies clean, exit `0`. A classic uninstall that actually removes a registered product is still **untested** |

Not yet exercised: a real apply run (uninstall + install) and the UAC self-elevation. Run `-WhatIf -Confirm:$false` on one pilot device before rolling out.



















