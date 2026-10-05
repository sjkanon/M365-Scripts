**English** · [Nederlands](readme.nl.md) · [Français](readme.fr.md)

[M365-Scripts](../../../readme.md) › [scripts](../../readme.md) › [Device](../readme.md) › **TempDisk**

# Temp Disk

Keeps the ephemeral temp disk (`D:`) of an Azure VM or AVD session host in place, and keeps the pagefile on it.

An Azure VM's temp disk — the resource disk, or the local NVMe disk on the newer sizes — is wiped whenever the VM is deallocated, resized or moved to another host. It comes back empty, sometimes RAW, sometimes offline, sometimes without its drive letter. Windows reads the pagefile configuration at boot, so **a pagefile configured on a drive letter that is not there at boot is simply not created**: the machine ends up paging on `C:` again, or running with no pagefile at all. That is what these two scripts exist to prevent.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Init-TempDisk.ps1`](Init-TempDisk.ps1) ([docs](#init-tempdiskps1)) | Restore the temp disk as `D:` and configure the pagefile on it |
| [`Register-InitTempDiskTask.ps1`](Register-InitTempDiskTask.ps1) ([docs](#register-inittempdisktaskps1)) | Install that script on the device and run it at every boot as SYSTEM |

---

### Init-TempDisk.ps1

Per run:

1. **Preflight** — every disk, what holds `D:`, the pagefile as configured (registry) and the pagefile actually in use in this session
2. **Letter** — an optical drive sitting on `D:` is moved out of the way; Windows hands `D:` to the DVD on an image with no temp disk and never gives it back
3. **Disk** — a volume that already *is* the temp disk gets its drive letter back; otherwise a RAW, non-boot, non-system disk is brought online, initialised GPT, partitioned and formatted NTFS
4. **Pagefile** — automatic management off, the pagefile pointed at `D:\pagefile.sys`, the entry for every other drive removed
5. **Verify** — all of it read back, saying what is in effect now and what waits for the next restart
6. **Restart** — only with `-RestartIfNeeded`: restart the machine when that is the one thing left between the configuration and a pagefile that is actually in use

A temp disk that only lost its drive letter is given the letter back, never reformatted — the volume is recognised by its label (`Temporary Storage`) or by the `DataLoss_Warning_Readme.txt` that Azure writes on the resource disk.

**Parameters**

| Parameter | Description |
|-----------|-------------|
| `-DriveLetter` | Drive letter for the temp disk (default: `D`) |
| `-Label` | Label written when formatting, and the label an existing temp volume is recognised by (default: `Temporary Storage`) |
| `-DiskNumber` | Format this disk instead of letting the script pick one — required when several RAW disks are present |
| `-InitialSizeMB` / `-MaximumSizeMB` | Pagefile size in MB. `0` (the default) on both means system managed |
| `-KeepSystemDrivePagefile` | Leave an existing pagefile on `C:` in place instead of removing it |
| `-SkipPagefile` | Only restore the disk; do not touch the pagefile configuration |
| `-OpticalDriveLetter` | Letter an optical drive is moved to when it holds `D:` (default: `Z`) |
| `-RestartIfNeeded` | Restart the machine when that is the only thing left between the configuration and a pagefile in use. Off by default |
| `-RestartDelaySeconds` | Countdown before a restart that happens while someone is signed in (default: `60`) — `shutdown /a` cancels it. With nobody signed in it restarts within seconds |
| `-RestartCooldownMinutes` | Shortest interval between two restarts triggered by this script (default: `60`) |
| `-RestartMarkerPath` | Registry key where the last self-triggered restart is remembered (default: `HKLM:\SOFTWARE\M365-Scripts\InitTempDisk`) |
| `-RestartEvenIfUsersSignedIn` | Restart even when someone is signed in. On a session host, drain it instead |
| `-CheckOnly` | Report only, change nothing. Exit code `2` means work is due |
| `-Quiet` | Print nothing unless there is news — the log file always gets the full story |
| `-LogPath` | Folder for `Init-TempDisk.log` (default: `C:\Temp`), appended and rotated past 1 MB |
| `-Force` | With `-DiskNumber`: format that disk even though it still carries partitions |

**Examples**

```powershell
# Read-only health report: where the temp disk is and what the pagefile really does
.\Init-TempDisk.ps1 -CheckOnly

# Walk the whole flow without touching the machine
.\Init-TempDisk.ps1 -WhatIf

# How the scheduled task runs it — silent while everything is in order,
# and one restart when that is what the pagefile is waiting for
.\Init-TempDisk.ps1 -Quiet -RestartIfNeeded

# Fixed 16 GB pagefile instead of a system managed one
.\Init-TempDisk.ps1 -InitialSizeMB 16384 -MaximumSizeMB 16384

# Say which disk is the temp disk, even though it still carries partitions
.\Init-TempDisk.ps1 -DiskNumber 2 -Force
```

**Exit codes**

| Code | Meaning |
|------|---------|
| `0` | The temp disk and the pagefile are as they should be |
| `1` | Failure |
| `2` | `-CheckOnly` only: work is due |

**Notes**
- **Nothing that is not RAW is ever initialised.** An empty temp disk and an unformatted data disk look identical from the outside, so a disk that already carries partitions is reported and left alone. With more than one RAW candidate the script refuses to guess and asks for `-DiskNumber`; `-Force` plus `-DiskNumber` is the only way to format a disk that still has partitions
- A local NVMe disk wins over other RAW disks when several are present — on the newer VM sizes that *is* the temp disk
- **A pagefile written by a run appears at the next restart.** Windows reads the configuration at boot and never re-reads it, so the run says so rather than claiming success. Because the task runs at every boot the device heals itself even without `-RestartIfNeeded`: the boot that recreates `D:` configures the pagefile, the next boot puts it in use
- `-RestartIfNeeded` closes that gap instead of waiting for it, and a script that runs at every boot and may restart the machine is a reboot loop waiting to happen — so it only fires when **all** of this holds: the run finished clean, the disk is there, the pagefile is configured on it and only this session is not using it; nobody is signed in (connected *or* disconnected — one `explorer.exe` per desktop, which is language independent where parsing `query.exe` is not); and no restart was triggered in the last `-RestartCooldownMinutes`, remembered under `-RestartMarkerPath`. A run that failed never restarts — that would hide the failure behind a reboot
- The restart goes through `shutdown.exe` with the planned "Operating System: Reconfiguration" reason, so it shows up as intended rather than unexpected. The countdown exists to warn people, so it only applies when there are people: signed in it is `-RestartDelaySeconds` and `shutdown /a` stops it; nobody signed in — the normal case at boot, and the guaranteed one on a session host whose pool is drained — and it restarts within seconds, because waiting a minute for an audience of nobody only costs availability
- Configuring a pagefile on a drive that is not there would write a setting Windows ignores, so the run fails instead if `D:` could not be restored
- Needs administrator rights; started by hand from an ordinary window it asks for elevation itself

---

### Register-InitTempDiskTask.ps1

Run once per device — by hand, from Tactical RMM / NinjaOne, or as an Intune platform script. Copies `Init-TempDisk.ps1` to a local folder (the repository is not there at boot time) and registers a task that runs it at every startup, as SYSTEM, with the highest privileges and without anyone signing in.

Re-running it is safe: a task with the same name is replaced, so this is also how the arguments the task runs with are changed.

The task runs `Init-TempDisk.ps1 -Quiet -RestartIfNeeded` by default. Windows reads the pagefile configuration at boot, so a boot on which the temp disk had to be rebuilt runs that whole session without a pagefile on `D:` unless the machine restarts once — and the guards above are what make that safe to hand to a boot task. Pass `-ScriptArguments '-Quiet'` to leave the restart out.

**Parameters**

| Parameter | Description |
|-----------|-------------|
| `-ScriptSourcePath` | `Init-TempDisk.ps1` to install (default: the copy next to this script) |
| `-ScriptTargetDir` | Folder it is copied to on the device (default: `C:\Scripts`) |
| `-TaskName` | Name of the scheduled task (default: `InitTempDisk`) |
| `-ScriptArguments` | Arguments for `Init-TempDisk.ps1` (default: `-Quiet -RestartIfNeeded`) |
| `-DelaySeconds` | Delay between boot and the task starting (default: `30`) |
| `-RunNow` | Also start the task once immediately, instead of waiting for a reboot |
| `-Unregister` | Remove the task and the installed copy of the script |

**Examples**

```powershell
# Show what would be installed and registered
.\Register-InitTempDiskTask.ps1 -WhatIf

# Install, register the boot task and run it once now
.\Register-InitTempDiskTask.ps1 -RunNow

# Same, but with a fixed 16 GB pagefile
.\Register-InitTempDiskTask.ps1 -ScriptArguments '-Quiet -RestartIfNeeded -InitialSizeMB 16384 -MaximumSizeMB 16384'

# Without the restart — the pagefile lands at whatever restart happens next
.\Register-InitTempDiskTask.ps1 -ScriptArguments '-Quiet'

# Take it off the device again
.\Register-InitTempDiskTask.ps1 -Unregister
```

**Notes**
- The task runs the *copy* in `C:\Scripts`, never the source, so the repository or the RMM staging folder can go away afterwards
- A missing source file is a hard error rather than a silent skip — a task registered against a file that is not there runs and fails at every boot without anyone noticing, until the pagefile is gone
- The startup trigger is delayed 30 seconds by default: the storage stack does not always have the temp disk enumerated the moment the task engine is up
- `-Unregister` deliberately leaves the pagefile configuration alone — removing the task must not take a machine's pagefile with it
- Registering says out loud whether the task may restart the machine; running it with `-RunNow` on a device nobody is signed in to can therefore start a 60-second countdown
