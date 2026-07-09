# Add Lockscreen to Start and Desktop

Pins a "Lock Workstation" shortcut to Start / Desktop via a downloaded `.bat` + `.ico`, using the undocumented `Windows.taskbarpin` Explorer verb to pin without user interaction.

---

## Files

| File | Description |
|------|-------------|
| [`add-lock.ps1`](#add-lockps1) | Downloads the lock script + icon and creates the shortcut |
| [`add-shortcut-lock.ps1`](#add-shortcut-lockps1) | Pins an arbitrary shortcut via the `Windows.taskbarpin` Explorer verb |

---

### add-lock.ps1

Creates `C:\Program Files\EOO\lockworkstation\` (skips if it already exists), downloads `lock.txt` + `lock.ico` from `https://endpoint.eoo.cloud/lock/`, renames `lock.txt` to `lock.bat`, and creates a Start Menu shortcut (`lock.lnk`) under `C:\ProgramData\...\Start Menu\Programs\EOO\lockworkstation\` that runs `explorer.exe` with the `.bat` as an argument.

```powershell
.\add-lock.ps1
```

> No parameters. Intended to run once per device (Intune, SYSTEM context) — re-running is a no-op once the `EOO` folder exists.

---

### add-shortcut-lock.ps1

Pins a target file to the taskbar/Start using the `Windows.taskbarpin` `ExplorerCommandHandler` registry verb, then cleans up the temporary registry keys it created.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-Target` | Yes | Path to the file/shortcut to pin |

```powershell
.\add-shortcut-lock.ps1 -Target "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\EOO\lockworkstation\lock.lnk"
```

> Depends on the `lock.lnk` shortcut created by `add-lock.ps1`.
