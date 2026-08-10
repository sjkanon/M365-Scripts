# Legacy Utilities — Device

Small workstation configuration utilities.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Set-NumLockDefault.ps1`](#set-numlockdefaultps1) | Set the default Num Lock state for new profiles and the sign-in screen |
| [`New-LockWorkstationShortcut.ps1`](#new-lockworkstationshortcutps1) | Create a Lock Workstation shortcut |

---

### Set-NumLockDefault.ps1

Writes `InitialKeyboardIndicators` under `HKU\.DEFAULT` (the template new
profiles are based on, also used at the sign-in screen) and the current user's
own hive. Dry-run by default.

```powershell
.\Set-NumLockDefault.ps1 -State On -Apply
```

---

### New-LockWorkstationShortcut.ps1

Creates a `.lnk` shortcut that runs the standard
`rundll32.exe user32.dll,LockWorkStation` command — no external download, no
registry pinning hacks. Generalized replacement for a script that downloaded a
custom icon/batch file from an internal endpoint and pinned it to the taskbar
via an undocumented registry key; this version just creates the shortcut and
lets the user pin it themselves if wanted. Dry-run by default.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-TargetFolder` | No | Where to create it (default: Public Desktop) |
| `-ShortcutName` | No | Default: "Lock Workstation" |
| `-IconPath` | No | Optional local `.ico` file |
| `-Apply` | No | Actually create the shortcut (default: preview) |

```powershell
.\New-LockWorkstationShortcut.ps1 -Apply
```
