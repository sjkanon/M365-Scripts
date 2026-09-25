# Update-TeamsClient.ps1 — how it works

Reference for [`Update-TeamsClient.ps1`](Update-TeamsClient.ps1): what it decides, in which order, and why it is built the way it is. For the short version (parameters and examples) see the [Device readme](readme.md#update-teamsclientps1). For the service desk there is a Dutch, support-level-oriented version to paste into IT Glue: [Update-TeamsClient-ITGlue.md](Update-TeamsClient-ITGlue.md).

---

## Table of Contents

- [What it is for](#what-it-is-for)
- [The decision](#the-decision)
- [The nine steps](#the-nine-steps)
- [AVD / VDI](#avd--vdi)
- [Proving it works: the VDI event log](#proving-it-works-the-vdi-event-log)
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

"Is the add-in installed?" means the files are there, not that a registry key says so. A machine-wide registration pointing at a loader DLL that no longer exists counts as **missing**, and the reason says so — `the machine-wide add-in registration points at files that are gone`. Reading the key alone is how a device whose add-in was deleted underneath it gets told there is nothing to do.

---

## The nine steps

| # | Step | What happens | `-WhatIf` |
|---|------|--------------|-----------|
| 1 | Preflight | Inventory every place Teams can live: `*MSTEAMS*` AppX packages per user, the provisioned package, classic Teams (machine-wide installer and per-profile installs), the meeting add-in in both uninstall hives, whether Outlook has it registered, the AVD components, what Teams logged about the media optimization, running Teams/Outlook | read-only |
| 2 | Version check | Ask the Teams config service for the published build, compare, decide | read-only |
| 3 | AVD | Only with `-AvdOptimizations`: set `IsWVDEnvironment`, install the WebRTC redirector. Only with `-RemoveWebRtcRedirector`: uninstall that redirector | guarded |
| 4 | Classic | Only with `-RemoveClassicTeams`: uninstall the Teams Machine-Wide Installer, clear the per-profile installs | guarded |
| 5 | Download | Create the working folder, download `teamsbootstrapper.exe`, check size and Authenticode signature | guarded |
| 6 | Uninstall | `msiexec /x` the add-in MSI (`1612` is retried against Windows Installer's cached copy), `Remove-AppxPackage -AllUsers`, `Remove-AppxProvisionedPackage` | guarded |
| 7 | Install | `teamsbootstrapper.exe -p` (provision for all users) | guarded |
| 8 | Add-in | Locate `MicrosoftTeamsMeetingAddinInstaller.msi` inside the installed package, compare its version with what is still registered, then clear **every other copy of it** (machine-wide folder, per-profile folders, per-user COM registrations) and install with `ALLUSERS=1` | guarded |
| 8b | Repair | Only with `-RepairOutlookAddIn`: clear a per-user registration that points at a removed DLL, put `LoadBehavior` back to 3 | guarded |
| 9 | Verify | Re-read the add-in registration machine-wide **and** per signed-in user in Outlook, the provisioned package, the classic removal and the AVD components, compare against the published build | reported as skipped |

"Guarded" means the step is wrapped in `$PSCmdlet.ShouldProcess(...)`, so under `-WhatIf` it prints what it would do and changes nothing.

Only what is actually missing gets done: steps 5–7 are skipped when the client is current, step 8 when the add-in is already there and the client was not replaced, step 3 unless `-AvdOptimizations` or `-RemoveWebRtcRedirector` is given, and step 4 unless `-RemoveClassicTeams` is given.

---

## AVD / VDI

`-AvdOptimizations` adds the two things a session host needs for Teams media optimization, both from the original VDI gap-fill script this replaces:

| Component | What it is | When it is installed |
|-----------|------------|----------------------|
| `HKLM:\SOFTWARE\Microsoft\Teams\IsWVDEnvironment` (DWORD `1`) | Tells Teams to hand media to the redirector instead of rendering it in the session | When not already `1` |
| Remote Desktop WebRTC Redirector Service | MSI from `https://aka.ms/msrdcwebrtcsvc/msi` (~1.7 MB, signed by Microsoft) | When not already installed, or with `-Force`. An installed redirector is **not** silently upgraded — `-Force` is what replaces it |

The flag is step 3, before the client is provisioned, because Teams reads it at startup to decide which media path to use.

Both are only touched when missing, so a scheduled run on a fully configured session host still downloads nothing and, with `-Quiet`, prints nothing. Without the switch the script does not change any of this — but it does **report** it: the flag, the redirector and the event log are read on every run on a session host, because that state is half the answer to "is Teams healthy here". A session host is recognised by `HKLM:\SOFTWARE\Microsoft\RDInfraAgent`.

### WebRTC is on its way out

Microsoft is retiring the WebRTC-based media optimization: **end of support 1 October 2026, end of availability 1 April 2027**. Teams shows users a banner about it. Its replacement, **SlimCore**, needs nothing installed on the session host — it ships inside new Teams and inside Windows App on the endpoint. Verified on a device with Teams `26225.1806.5074.1452`:

```
Microsoft.Teams.SlimCoreVdiHost.win-x64          2026.31.1.16
Microsoft.Teams.SlimCoreVdiFwk.win-x64.2026.31   2026.31.1.16   (plus older framework versions)
```

Preflight reports that package when it is standing on an endpoint. On a session host there is nothing to report, because nothing is staged there — which is why the script no longer looks for it in that case. A check that looks in the wrong place does not fail, it lies.

Which path is actually taken depends on the Windows App version on the endpoint the user connects from:

| Side | Minimum |
|------|---------|
| Session host | Teams `24193.1805.3040.8975` |
| Endpoint, Windows | Windows App `2.0.352.0` (the classic Remote Desktop client is no longer supported) |
| Endpoint, macOS | Windows App `11.3.4`, the non-Store `.pkg` build |

`IsWVDEnvironment` stays required either way, and Microsoft's current guidance is to keep the redirector installed as a fallback for endpoints that cannot do SlimCore — so `-AvdOptimizations` keeps installing it.

### What actually blocks the staging

When SlimCore does not arrive on an endpoint, it is usually a policy that stopped it. Preflight checks the three Microsoft documents, each with the code Teams logs for it:

| Policy | Effect | Teams code |
|--------|--------|------------|
| `BlockNonAdminUserInstall = 1` | A non-admin cannot register the package | `16389` |
| `AllowAllTrustedApps = 0` | Sideloading is off, the MSIX cannot install | `15615` |
| AppLocker on packaged apps, with no rule that lets the SlimCoreVdi packages through | The packages are blocked | `10083` |

AppLocker is **read**, not merely detected. An earlier version warned on the mere existence of `HKLM:\SOFTWARE\Policies\Microsoft\Windows\SrpV2`, which cried wolf on every fleet that has ever written an Exe rule. Three things decide whether it can stop the MSIX, and each is easy to get wrong:

- **Only the packaged-app collection applies.** An MSIX never meets the `Exe`, `Msi`, `Script` or `Dll` collections, so an enforced Exe rule set is not a SlimCore blocker and is no longer reported as one.
- **"Not configured" is not harmless.** Microsoft: a collection that holds at least one rule and has enforcement not configured *is enforced*. Only an explicit `EnforcementMode = 0` (audit only) lets everything through. The report says so in as many words — `Appx enforced (enforcement not configured, which still enforces)`.
- **Nothing is enforced while the service is stopped.** AppLocker does nothing at all unless the Application Identity service (`AppIDSvc`) is running. The service state is printed, and when it is not running the blocker line says the policy is not being enforced *at this moment* — which is a different problem from the policy being wrong, and a service that starts on the next reboot turns it into one.

What gets printed is the registry path, the mode per collection, the service state and the rule names (the first five, then a count of the rest) — because "AppLocker policy is present" is not something a technician can act on, and a path plus a rule name is. When a rule already allows the packages by name it is reported as `[ OK ]`; when a rule allows anything signed by `O=MICROSOFT CORPORATION` the report says it *should* cover SlimCore and to check it has not been narrowed to one product name.

The matching is a text match against the rule XML, so it errs on the side of warning: a broad allow rule that names neither Microsoft nor the packages — `PublisherName="*"`, say — would in reality let SlimCore through, but is still reported as a blocker. The rule names printed alongside it are there so a technician can settle that in a second.

On a **session host** these policies block nothing, because the staging happens on the endpoint. They are still named there, as `[SKIP]` reference lines rather than warnings, because it is usually the same GPO that is linked to both — so the thing worth checking is whether it also reaches the endpoints.

### Removing the old optimization

`-RemoveWebRtcRedirector` is the other direction, for a fleet that has finished the migration. It uninstalls the redirector (`msiexec /x`, or clears the leftover Programs and Features entry when msiexec answers `1605`) and touches nothing else. It is **mutually exclusive** with `-AvdOptimizations`, which installs the same component; giving both is refused before the UAC prompt.

`IsWVDEnvironment` is deliberately **not** cleared: SlimCore needs that flag just as much as WebRTC did.

Do not run this until every endpoint really can do SlimCore. An endpoint below the minimum above that no longer finds the redirector does not fail — it falls back to rendering media on the session host, which is exactly the CPU load both optimizations exist to avoid. Check the event log below first: `code 16002` means at least one endpoint still has no plugin.

---

## Proving it works: the VDI event log

Everything above only proves the parts are in place. Whether users are actually optimized is a question about their endpoints, and the one place a session host can answer it is the Application event log. Teams (build `24123` or newer) writes a **`Microsoft Teams VDI`** event, ID `0`, on every connect and disconnect, and the description carries the codes from Microsoft's connection error table.

Preflight reads the last seven days, reports the count and the newest timestamp, shows up to three errors with their message, and translates every code it recognises. The codes worth knowing:

| Code | Meaning |
|------|---------|
| `24002`, `24010` (`3000`, `3001`) | Not errors — the user is on SlimCore. This is what success looks like |
| `16002` (`2000`) | The endpoint has no plugin: Windows App is missing, too old, or did not load |
| `16026` (`2003`) | A Citrix policy blocks the `MSTEAMS`/`MSTEAM1`/`MSTEAM2` virtual channels |
| `16043` (`2005`) | Teams runs as a published app or RemoteApp — those sessions stay on WebRTC by design |
| `16066` (`2008`) | A Mac endpoint on the Store build of Windows App; only the non-Store build is supported |
| `16389` | `BlockNonAdminUserInstall` on the endpoint stopped the MSIX registration |
| `10083` (`1260`) | AppLocker or WDAC blocked the SlimCore packages |
| `1951` (`15615`) | Sideloading is off on the endpoint (`AllowAllTrustedApps` is `0`) |
| `24018` (`3002`), `24043` (`3005`), `24058` (`3007`) | SlimCore never downloaded, or timed out doing so |
| `24170` (`3021`) | Packages present but no usable set — the Host MSIX is missing |
| `1722`, `15616`, `24035` (`3004`) | Transient. Only worth chasing if users stay unoptimized |
| `4390` | Thin client with a write filter or RAM disk — point `MSTEAMSVDI_BITS_TMP_PATH` at a real disk |

Two details that cost time to find out:

- The query has to use `-FilterXPath`, not `-FilterHashtable`. `Get-WinEvent -FilterHashtable @{ ProviderName = 'Microsoft Teams VDI' }` **throws** when the provider has never written an event — which is the normal case on a healthy non-VDI machine. The XPath form returns nothing and a soft error instead. `[System.Diagnostics.EventLog]::SourceExists()` is no use either: it walks the Security log and throws for anything but SYSTEM.
- A `0` in `loadErrc` or `deployErrc` means *that phase* raised no error, not that the session is fine. It is deliberately absent from the table, because printing "OK" next to a real failure in the other phase would be a lie.

To filter on the source by hand in Event Viewer, register it once on the session host:

```powershell
New-EventLog -LogName Application -Source 'Microsoft Teams VDI'
```

The script does not need this — it filters on the provider name in the event itself — so it does not do it for you.

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

## Why the sweep waits for the MSI

A production `-Force` run on a session host ended like this:

```
6. Uninstall current Teams
[WARN] Uninstall of Microsoft Teams Meeting Add-in for Microsoft Office failed (exit code 1612)
[ OK ] Removed the add-in copy for all users (machine-wide)
[ OK ] Removed the add-in copy for ETNO\itceadmin
8. Teams Meeting Add-in (install)
[FAIL] Aborted: Teams Meeting Add-in install failed (exit code 1638)
```

Every working copy of the add-in was deleted and the replacement then refused to install, so the device was left with no add-in at all. Three things had to line up, and all three are now handled:

| What happened | Why | Now |
|---------------|-----|-----|
| The host ran Teams `26246.1604.5133.838`, newer than the published `26225.1806.5074.1452`, and `-Force` reinstalled anyway | `-Force` means "reinstall regardless of the version check", which on a newer-than-published build is a **downgrade** | The version check warns that `-Force` is about to downgrade, and names the add-in consequence |
| `msiexec /x` on the add-in answered `1612` | Windows Installer could not find the source it needs to uninstall, so the old product stayed registered | `1612` is retried against Windows Installer's own cached MSI under `C:\Windows\Installer` (found through `...\Installer\UserData\S-1-5-18\Products\*\InstallProperties`, `LocalPackage`). When that file is gone too, the run says the registration cannot be removed and what that will cause |
| Installing add-in `1.26.21803` over the newer registered one answered `1638` | Windows Installer refuses an older version over a newer one | The versions are compared **before** anything is deleted. An older MSI means the sweep and the install are skipped and the working add-in is left alone; a `1638` that still happens is a warning, not an abort, so verification runs and reports what Outlook actually has |

The general rule this follows is the one already applied to the bootstrapper: nothing that works is destroyed until its replacement is in hand and known to be installable.

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
- **"It still does not load" gets a reason, not a shrug.** A registration that is present but not loading is checked against the three causes that leave no trace in `LoadBehavior` itself, and each is reported as a `why:` line under the warning:

  | Cause | What the script looks at | The fix |
  |-------|--------------------------|---------|
  | Bitness mismatch | Office platform from the Click-to-Run configuration versus the `\x64\` or `\x86\` loader the registration points at | Register the loader matching Outlook''s bitness |
  | Outlook parked it | `Resiliency\DisabledItems` and `CrashedAddins` in that user''s hive, decoded from the binary values | File > Options > Add-ins > Manage: Disabled Items, then the `DoNotDisableAddinList` policy to keep it out |
  | Policy overrides the user | `HKLM\SOFTWARE\Policies\Microsoft\Office\<ver>\Outlook\Addins\TeamsAddin.FastConnect` | Change or remove that policy |

  When none of them applies it says so, which is also an answer: nothing on the machine is blocking it, so what is left is a full Outlook restart and a user who has signed in to Teams at least once.

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
| `-RemoveWebRtcRedirector` | — | Uninstall the old WebRTC media optimization. Cannot be combined with `-AvdOptimizations` |
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

**Nothing that works is deleted before its replacement is in hand.** The add-in sweep used to run in step 6, before the new package was staged - so the MSI version that would replace it was not knowable yet. On a host where `-Force` downgraded the client, step 6 deleted every working copy and step 8 then discovered the package carried an older add-in than the one still registered, which Windows Installer refuses with `1638`. The run aborted with the device holding no add-in at all. The sweep now happens in step 8, after the version comparison, which makes the destructive part conditional on the constructive part being possible - the same rule that already put the signature-checked download before the first uninstall.

**The redirector MSI does not upgrade itself.** It keeps one ProductCode across versions, so `msiexec /i` over an existing install answers `1638` ("another version of this product is already installed") instead of upgrading — which is exactly what a `-Force` run on a configured session host hit. The script now compares the downloaded ProductVersion with what is registered: same version means a repair (`REINSTALL=ALL REINSTALLMODE=vomus`), a different version means the old one is uninstalled first, and a `1638` that still slips through is reported as "leaving the existing one in place" instead of failing the run.

**A policy is worth reading, not just detecting.** The AppLocker check used to warn whenever the `SrpV2` key existed. That key exists on any fleet with a single Exe rule, so the warning fired constantly on machines where nothing was blocked — and a check that always fires is a check nobody reads. It now looks at the one collection an MSIX can meet, at whether that collection is really enforced (including the "not configured with rules" case, which is), at whether the Application Identity service makes any of it real, and at whether an existing rule already allows the packages. Warning on the presence of a thing rather than on what it says is the same mistake as looking in the wrong place: it does not fail, it just stops meaning anything.

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
| `Use either -AvdOptimizations ... or -RemoveWebRtcRedirector` | The two switches install and uninstall the same component. Pick one |
| `Could not remove the WebRTC Redirector` | msiexec refused the uninstall. Check `C:\Temp` transcript for the exit code; `1605` is handled (stale entry), anything else usually means another installation is running |
| Users see `Azure Virtual Desktop Media Optimized` instead of SlimCore | Expected while the redirector is still installed and the endpoint has no plugin. Check the event log section for `16002` |
| `Uninstall of ... Add-in ... failed (exit code 1612)` | Windows Installer has lost the source MSI. It is retried from the cached copy under `C:\Windows\Installer`; if that is gone the product cannot be uninstalled by msiexec at all and keeps refusing other versions with `1638` |
| `Teams Meeting Add-in install failed (exit code 1638)` | An older add-in is being installed over a newer registered one. No longer fatal, and normally prevented: the versions are compared first. Usually caused by `-Force` downgrading the client on a host whose build was newer than the published one |
| `The staged package carries add-in X, older than the Y still registered` | Working as intended - the add-in was left alone. Get a Teams build whose add-in is at least Y, or clear the orphaned product registration by hand |
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
| `Microsoft Teams VDI` event query on a machine without the provider | `[SKIP] No 'Microsoft Teams VDI' events in the last 7 days`, 357 ms, no terminating error. Exercised by running the real functions out of the script; preflight calling them has **not** been run on a live session host |
| Same query against a provider that does exist | Counted 42 events and reported the newest timestamp in 104 ms; the error branch formats, truncates at 200 characters and decodes nothing from version strings |
| Code decoder against Microsoft's documented log lines | `deployErrc=24002` → on SlimCore; `"val":4390` → write-filter error; a line of nothing but version numbers → decodes nothing |
| `-AvdOptimizations -RemoveWebRtcRedirector` together | Refused with exit `1` before the UAC prompt |
| `-RemoveWebRtcRedirector` end to end | **Untested** — no session host with the redirector installed was available. The uninstall reuses the `msiexec /x` + `1605` path that is proven for classic Teams |
| Cached-package lookup for a `1612` uninstall | Verified against real products on this machine: `...\Installer\UserData\S-1-5-18\Products\*\InstallProperties` yields `DisplayName`, `DisplayVersion` and a `LocalPackage` path that exists. Asking for the add-in on a machine that does not have it machine-wide returns nothing and does not throw. The `msiexec /x <cached msi>` retry itself is **untested** |
| Add-in version guard | Exercised directly: an MSI older than the registered add-in skips, newer installs, equal installs, and an unparsable registered version does not block the install |
| `1638` on the add-in install | Now a warning that lets verification run, instead of an abort. **Untested** live - the production host that produced it has not been re-run |
| Classic removal, applied | Against a faked profile folder in a scratch directory: folder removed, verification reports `No per-user classic Teams left`, exit `0` |
| Classic removal on a real session host | `-Force -RemoveClassicTeams -AvdOptimizations` on an AVD host removed a real per-profile classic Teams `1.4.00.11161`. The machine-wide msiexec path is still **untested** - that host had no Machine-Wide Installer |
| Add-in step after provisioning | Broke on that same production run (`Get-AppxPackage` is per user, provisioning is not) and now resolves the MSI from the staged `WindowsApps` package instead. Re-tested here for a per-user install, a provisioned-only host and a `-Force` run |
| Outlook add-in switched off | That production run found `LoadBehavior 2` for one account - the check earns its keep on the first real device it saw |
| Profiles that cannot be read | Listed by name with what it means for them: with a machine-wide registration they pick it up at the next Outlook start, without one there is nothing to fall back on. Both messages verified against stubbed profile lists |
| Add-in load diagnostics | Both detections exercised: an x86 loader path against x64 Office produces the bitness reason, and a planted binary `CrashedAddins` value is decoded and reported. A healthy registration produces no `why:` line at all |
| Add-in sweep before reinstall | `Get-TeamsAddInFolder` verified against this device (finds the real per-profile copy), and the `-WhatIf` plan shows the folder plus both CLSID views being removed before the reinstall. The removal itself reuses mechanics proven live in the classic-Teams and repair tests; the sweep as a whole runs for the first time on a production host |
| Per-user registration repair | Planted a stale CLSID (both registry views) plus `LoadBehavior 2` against a healthy machine-wide registration: `-WhatIf` planned both actions, an applied run cleared the keys and set `LoadBehavior` to 3, exit `0`. The real fix on a production host is still to be confirmed |
| Dangling add-in registration | Measured here: the COM class resolves to `%LOCALAPPDATA%\Microsoft\TeamsMeetingAdd-in\1.26.21803\x64\Microsoft.Teams.AddinLoader.dll`, a targeted lookup across loaded hives costs ~100 ms, and a planted registration pointing at a missing DLL is reported as "re-enabling will not stick" |
| AppLocker reading | Every branch exercised against a stubbed policy tree: an enforced `Exe` collection with no `Appx` collection produces **no** blocker (the old false positive); an enforced `Appx` collection with no matching allow rule warns and names the path and rule count; an explicit SlimCore allow rule reports `[ OK ]`; an `O=MICROSOFT CORPORATION` rule reports the "check it is not narrowed" note; `EnforcementMode = 0` is reported as audit only and blocks nothing; rules with enforcement not configured are reported as enforced; seven rules print five and `... and 2 more`; a missing `SrpV2` key returns nothing at all. Verified on a machine with no AppLocker policy: output is silent, as it should be. **Untested** against a live enforced AppLocker policy on a real endpoint |
| SlimCore reporting | Verified on both sides: on this endpoint it finds `Microsoft.Teams.SlimCoreVdiHost.win-x64 2026.31.1.16`; with `RDInfraAgent` faked it reports the session-host wording instead. The three blocker policies were exercised against stubbed registry reads |
| Redirector replacement | A second session host failed on `1638` with redirector `1.54.2408.19001` installed while `aka.ms` now serves `1.56.2603.20001`. Both paths are now planned correctly under `-WhatIf` (replace: `/x` then `/i`; same version: `REINSTALL=ALL`), but **neither msiexec call has been run for real** |
| Stale machine-wide classic entry | That same host then failed on `1605` from the classic uninstall. Tested live: a real `msiexec /x` against an unknown product code returns 1605, the run warns, removes the stale registry entry, continues and verifies clean, exit `0`. A classic uninstall that actually removes a registered product is still **untested** |

Not yet exercised: a real apply run (uninstall + install) and the UAC self-elevation. Run `-WhatIf -Confirm:$false` on one pilot device before rolling out.




















