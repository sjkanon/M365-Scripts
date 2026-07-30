# CoworkPrerequisites

Windows-side prerequisites for [Claude Cowork](https://support.claude.com/en/articles/12622667-enterprise-configuration-for-claude-desktop) — the `VirtualMachinePlatform` optional feature and disabling Windows Fast Startup — deployed as an Intune **Proactive Remediation** (Detect + Remediate script pair), entirely separate from [`../ClaudeDesktop/`](../ClaudeDesktop/readme.md).

## Why a remediation instead of a Win32 app

This used to be a Win32 app, wrapped around the same `Add-/Set-IntuneWin32App` machinery as Claude Desktop. That turned out to be the wrong tool for the job:

- There's no actual binary here — just two OS config changes (enable a feature, flip a registry value). All the Win32-app scaffolding (detection rule, requirement rule, `@odata.type`-tagged objects, return-code mapping for the 3010 "soft reboot" exit code) was overhead that added failure surface without adding value, and it *did* fail in practice — see the `IntuneWin32App`-module bug documented in [`../ClaudeDesktop/readme.md`](../ClaudeDesktop/readme.md#known-issue-0x80070001-install-failures--and-a-much-deeper-intunewin32app-module-bug-behind-it), which affected this app too.
- A Win32 app only runs its install command **once** per device (until content/version changes force a re-push). If a device fails once — a transient DISM lock, a user who never restarts — nothing re-tries it.
- A Proactive Remediation re-runs **on its own schedule** (daily by default): it re-checks compliance and re-applies the fix automatically, without you needing to redeploy anything. That fits "enable this feature and nag the user to restart until they do" much better than a one-shot install.

## Contents

| Script | Role |
|---|---|
| `Deploy-CoworkPrerequisitesRemediation.ps1` | **The one you run.** Creates/updates the remediation in Intune and assigns it — see "Deploy" below. |
| `Detect-CoworkPrerequisites.ps1` | Detection half — exit 0 if `VirtualMachinePlatform` is enabled **and** the HCS services (`vmcompute`, `HNS`, `vfpext`) exist, exit 1 otherwise |
| `Remediate-CoworkPrerequisites.ps1` | Remediation half — enables `VirtualMachinePlatform`, disables Fast Startup, notifies the logged-on user if a restart is now required |

## What it checks / fixes

1. **VirtualMachinePlatform**: `Enable-WindowsOptionalFeature`, with retry against transient DISM failures. No-op if already enabled.
2. **HCS services present** (`vmcompute`, `HNS`, `vfpext`): `VirtualMachinePlatform` showing `Enabled` in DISM doesn't guarantee these already exist (see Anthropic's own Cowork troubleshooting: *"Missing HCS services: HNS, vmcompute, vfpext"*) — particularly right after enabling the feature but before the required restart. Detection deliberately does **not** check service `Status` (e.g. `Running`): `vmcompute` is a trigger-start service and is expected to be `Stopped` whenever no Cowork session is active — checking for `Running` would flag a perfectly healthy device as non-compliant. Only the service being completely absent is a reliable "not there yet" signal.
3. **Fast Startup**: sets `HiberbootEnabled = 0` under `HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power`, on every remediation run. Anthropic's own docs warn: *"Restart the machine using Restart, not shut down and power on. With Windows Fast Startup enabled, a shutdown cycle can leave the virtualization services uninitialized."*

## Restart: the user does it, the script doesn't

If `VirtualMachinePlatform` was just enabled, the remediation script sends an English `msg.exe` notification to the active console session (matched via the untranslated `SESSIONNAME` value `console`, not the OS-language-dependent `STATE` text "Active" — a plain string match on "Active" would silently never fire on a non-English Windows display language) telling the user to restart. **It does not restart the device itself.**

This is deliberate, not a missing feature: the user restarts on their own schedule, not one forced by the script. Because detection keeps reporting "not compliant" until the restart actually happens, and remediations re-run daily by default, the reminder simply repeats on the next cycle instead of being a one-shot notification that's easy to miss and never followed up on.

## Deploy

```powershell
.\Deploy-CoworkPrerequisitesRemediation.ps1 -AssignmentGroupName "SG-Apps-ClaudeDesktop"
```

Reads `Detect-CoworkPrerequisites.ps1` + `Remediate-CoworkPrerequisites.ps1`, base64-encodes them, and creates or updates the remediation via the Microsoft Graph `deviceManagement/deviceHealthScripts` API (`beta` only — Proactive Remediations have no `v1.0` endpoint, so this doesn't use the `IntuneWin32App` module at all, just `Invoke-MgGraphRequest` on an ordinary delegated `Connect-MgGraph` session). First run creates and assigns it; later runs always re-PATCH the content (cheap, idempotent) but leave the existing assignment alone unless you pass `-ReassignGroup`.

Assigned to the group on a **daily schedule** by default (`-IntervalDays 1`, `-RunTimeLocal "03:00"`) — unlike [`Repair-StuckWin32AppEnforcement.ps1`'s remediation pair](../../readme.md#detect--remediate-stuckwin32appenforcementps1) (deliberately left unassigned/on-demand, since silently auto-clearing a retry lockout can mask a genuinely broken deployment), this one is meant to proactively reach and self-heal the whole target fleet, not just react to one flagged device.

| Parameter | Default | Description |
|---|---|---|
| `-AssignmentGroupName` | *(required)* | Entra ID group. Used at first creation, and with `-ReassignGroup` |
| `-DisplayName` | `Cowork Windows Prerequisites` | Used to find the existing remediation on later runs |
| `-IntervalDays` | `1` | Recurrence in days for the schedule |
| `-RunTimeLocal` | `03:00` | Local time the daily run fires |
| `-ReassignGroup` | off | Re-apply the assignment (group/schedule) on an existing remediation — needed to change `-AssignmentGroupName`/`-IntervalDays`/`-RunTimeLocal` later |
| `-TenantId` | auto-detected | Entra ID tenant ID |
| `-Force` | off | Skip the confirmation prompt(s) |

Required role: one that can grant `DeviceManagementScripts.ReadWrite.All` consent (Intune Administrator or Global Administrator).

## Independence from Claude Desktop

Cowork is optional: Claude Desktop works fine without it. Keeping this as a separate remediation (rather than folding it into Claude Desktop's own install) means a Windows-feature hiccup unrelated to Claude never blocks or gets confused with a Claude Desktop install problem — each has its own, separately visible status in Intune. There's no Intune "Dependency" between the two (remediations can't be a Win32-app dependency target anyway) — assign both to the same group and let them run independently.

## Reverting

`VirtualMachinePlatform` is **not** disabled by this remediation pair, and there's no "uninstall" concept for a remediation — if you need to turn the feature back off on a device, do it manually (`Disable-WindowsOptionalFeature -Online -FeatureName VirtualMachinePlatform`) after confirming nothing else (WSL, Hyper-V-based tools, other Cowork-like apps) depends on it. Fast Startup is left disabled either way — it's a harmless, device-wide setting, not something specific to Cowork.

## Prerequisites

- `Microsoft.Graph.Authentication`, `Microsoft.Graph.Groups` PowerShell modules — install with `.\scripts\Startup\Install-Modules.ps1`
- Run from Windows or any platform with PowerShell 5.1+ — nothing here is Windows-only (unlike Claude Desktop's MSIX/AppX packaging step), it's just Graph calls
