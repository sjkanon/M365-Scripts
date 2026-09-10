# Update-TeamsClient.ps1 — how it works

Reference for [`Update-TeamsClient.ps1`](Update-TeamsClient.ps1): what it decides, in which order, and why it is built the way it is. For the short version (parameters and examples) see the [Device readme](readme.md#update-teamsclientps1). For the service desk there is a Dutch, support-level-oriented version to paste into IT Glue: [Update-TeamsClient-ITGlue.md](Update-TeamsClient-ITGlue.md).

---

## Table of Contents

- [What it is for](#what-it-is-for)
- [The decision](#the-decision)
- [The eight steps](#the-eight-steps)
- [AVD / VDI](#avd--vdi)
- [The meeting add-in: machine-wide versus per user](#the-meeting-add-in-machine-wide-versus-per-user)
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

## The eight steps

| # | Step | What happens | `-WhatIf` |
|---|------|--------------|-----------|
| 1 | Preflight | Enumerate `*MSTEAMS*` AppX packages for all users, find the meeting add-in in both uninstall hives, check the AVD components, note running Teams/Outlook | read-only |
| 2 | Version check | Ask the Teams config service for the published build, compare, decide | read-only |
| 3 | AVD | Only with `-AvdOptimizations`: set `IsWVDEnvironment`, install the WebRTC redirector | guarded |
| 4 | Download | Create the working folder, download `teamsbootstrapper.exe`, check size and Authenticode signature | guarded |
| 5 | Uninstall | `msiexec /x` the add-in, `Remove-AppxPackage -AllUsers`, `Remove-AppxProvisionedPackage` | guarded |
| 6 | Install | `teamsbootstrapper.exe -p` (provision for all users) | guarded |
| 7 | Add-in | Locate `MicrosoftTeamsMeetingAddinInstaller.msi` inside the installed package, install it with `ALLUSERS=1` | guarded |
| 8 | Verify | Re-read the add-in registration, the provisioned package and the AVD components, compare against the published build | reported as skipped |

"Guarded" means the step is wrapped in `$PSCmdlet.ShouldProcess(...)`, so under `-WhatIf` it prints what it would do and changes nothing.

Only what is actually missing gets done: steps 4–6 are skipped when the client is current, step 7 when the add-in is already there and the client was not replaced, and step 3 unless `-AvdOptimizations` is given.

---

## AVD / VDI

`-AvdOptimizations` adds the two things a session host needs for Teams media optimization, both from the original VDI gap-fill script this replaces:

| Component | What it is | When it is installed |
|-----------|------------|----------------------|
| `HKLM:\SOFTWARE\Microsoft\Teams\IsWVDEnvironment` (DWORD `1`) | Tells Teams to hand media to the redirector instead of rendering it in the session | When not already `1` |
| Remote Desktop WebRTC Redirector Service | MSI from `https://aka.ms/msrdcwebrtcsvc/msi` (~1.7 MB, signed by Microsoft) | When not already installed, or with `-Force` |

The flag is step 3, before the client is provisioned, because Teams reads it at startup to decide which media path to use.

Both are only touched when missing, so a scheduled run on a fully configured session host still downloads nothing and, with `-Quiet`, prints nothing. Without the switch the script does not change any of this — it only points out that the device looks like a session host (`HKLM:\SOFTWARE\Microsoft\RDInfraAgent` exists).

---

## The meeting add-in: machine-wide versus per user

The script installs the add-in MSI with `ALLUSERS=1` into `%ProgramFiles(x86)%\Microsoft\TeamsMeetingAddin\<version>` — a **per-machine** install, which is what Microsoft prescribes for a shared machine or an AVD/RDS session host so that every user who logs on has it.

That is not the only way the add-in gets onto a device. On an ordinary endpoint the Teams client installs and updates it **per user**, by itself. Measured on a Windows 11 endpoint with Teams `26225.1806.5074.1452`:

| What | Where |
|------|-------|
| Add-in files | `%LOCALAPPDATA%\Microsoft\TeamsMeetingAdd-in\1.26.21803` (note the hyphen) |
| MSI the client installed it from | `%LOCALAPPDATA%\Microsoft\TeamsMeetingAddinMsis\1.26.21803\` — seven cached versions, one per Teams update since April |
| Outlook COM registration | `HKCU\SOFTWARE\Microsoft\Office\Outlook\Addins\TeamsAddin.FastConnect`, `LoadBehavior=3` |
| Machine-wide Outlook registration | none — no Teams entry under `HKLM\...\Office\Outlook\Addins` at all |
| Uninstall entry | `HKLM` **64-bit** hive, product code `{A7AB73A3-...}`, `InstallSource` pointing at that per-user MSI cache |

Three consequences worth knowing:

- **Outlook loads add-ins per user.** A machine-wide install makes the add-in *available* to everyone; each user's Outlook still picks it up on its next start. That is why the script warns when Outlook is running.
- **What the verification proves.** Step 8 reads the `HKLM` uninstall keys. That confirms the machine-wide install succeeded — it does not prove that a particular user's Outlook has the button. Run as System (the normal RMM case) the script sees `HKCU` of the System account, so a per-user-only install is invisible to it.
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
- The comparison is a `[version]` comparison against the installed AppX package version. **Equal or newer means nothing to do** — a device on an insider build is left alone rather than downgraded.
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
| `timed out after 900 seconds and was killed` | A hung msiexec or a slow image. Raise `-TimeoutSeconds`; check whether another installation is running |
| `Teams Meeting Add-in installer not found` | The installed package has no add-in MSI (an unexpected build). Re-run; if it persists, install the client first and check the `WindowsApps` folder |
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

Not yet exercised: a real apply run (uninstall + install) and the UAC self-elevation. Run `-WhatIf -Confirm:$false` on one pilot device before rolling out.





