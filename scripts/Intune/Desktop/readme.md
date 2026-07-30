# Desktop

Intune-deployed desktop customization: corporate wallpaper + lockscreen, and a taskbar lock-workstation shortcut.

> Office theme/color deployment (`Deploy-OfficeTheme.ps1`, `Deploy-Officecolors.ps1`) lives in [`Custom Scripts/Intune/Desktop/`](../../Custom%20Scripts/Intune/Desktop/readme.md) — those scripts hardcode their download URL to that path, so they stay put.

---

## Contents

| Item | Description |
|------|-------------|
| [`Background/`](Background/readme.md) | Corporate wallpaper (`Desktop/`) and lockscreen (`Lockscreen/`) |
| [`Add Lockscreen to start and desktop/`](Add%20Lockscreen%20to%20start%20and%20desktop/readme.md) | Pins a "Lock Workstation" shortcut to Start |
| [`ClaudeDesktop/`](ClaudeDesktop/readme.md) | Machine-wide Claude Desktop deployment, one script run monthly to stay current |
| [`CoworkPrerequisites/`](CoworkPrerequisites/readme.md) | Windows-side Cowork prerequisites (`VirtualMachinePlatform`, Fast Startup) — an Intune Proactive Remediation, not bundled into Claude Desktop |
| [`Deploy-AllIntune.ps1`](#deploy-allintuneps1) | Runs both of the above deploy scripts in one call |

---

### Deploy-AllIntune.ps1

Thin orchestrator with no Intune/Graph logic of its own — runs `CoworkPrerequisites/Deploy-CoworkPrerequisitesRemediation.ps1` then `ClaudeDesktop/Deploy-ClaudeDesktopIntune.ps1`, passing through `-AssignmentGroupName`/`-TenantId`/`-Force` to both. Each deploy still manages its own Graph session independently (the Cowork remediation deploy connects delegated directly, no temporary App Registration needed there — see its readme) — this just saves running two commands by hand.

```powershell
# Both apps, one command
.\Deploy-AllIntune.ps1 -AssignmentGroupName "SG-Apps-ClaudeDesktop"

# Unattended (e.g. scheduled task)
.\Deploy-AllIntune.ps1 -AssignmentGroupName "SG-Apps-ClaudeDesktop" -Force

# Skip Cowork Prerequisites, only Claude Desktop
.\Deploy-AllIntune.ps1 -AssignmentGroupName "SG-Apps-ClaudeDesktop" -SkipCoworkPrerequisites
```

Stops before running Claude Desktop if the Cowork Prerequisites deploy fails (pass `-ContinueOnError` to run it anyway).
