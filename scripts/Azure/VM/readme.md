# Azure-NVMe-Conversion.ps1

> **Vendored third-party script.** This is Microsoft's own tool from [`Azure/SAP-on-Azure-Scripts-and-Utilities`](https://github.com/Azure/SAP-on-Azure-Scripts-and-Utilities) (MIT licensed) — kept as-is rather than rewritten, since it's already maintained upstream. Check the `.LINK` in the script header for the latest version before relying on it for something critical.

Converts an Azure VM's disk controller type between SCSI and NVMe, including in-guest driver preparation. Changing controller type changes how disks are presented inside the OS, so switching an already-provisioned VM to an NVMe-only size (e.g. `Standard_E*bds_v5`/`v6`) without preparing the guest first can cause `INACCESSIBLE_BOOT_DEVICE` on boot.

**What it does**

1. Validates the target VM (exists, running, Gen2 image, current controller type, target SKU supports the requested controller type and is available in the VM's zone)
2. For Windows VMs converting to NVMe: runs a **read-only check** inside the guest via `Invoke-AzVMRunCommand` (validates the `stornvme` driver is present and boot-started across *every* `ControlSet`, not just the current one) — pass `-FixOperatingSystemSettings` to also run the **fix** (removes the `StartOverride` registry key that blocks the driver from loading at boot, with an explicit registry flush so the change survives an immediate deallocate)
3. For Linux VMs: checks/fixes the `nvme` driver in `initrd`/`initramfs` depending on distro (Ubuntu/Debian: `update-initramfs`; RHEL family/SUSE: `dracut`)
4. Updates the OS disk's supported capabilities and the VM's disk controller type / size
5. Optionally restarts the VM (`-StartVM`) and writes a timestamped log file (`-WriteLogfile`)

**Parameters**

| Parameter | Required | Default | Description |
|-----------|----------|---------|-------------|
| `-ResourceGroupName` | Yes | — | Resource group containing the VM |
| `-VMName` | Yes | — | VM to convert |
| `-VMSize` | Yes | — | Target VM size (must support the target controller type) |
| `-NewControllerType` | No | `NVMe` | `NVMe` or `SCSI` |
| `-StartVM` | No | off | Start the VM after conversion |
| `-WriteLogfile` | No | off | Write `Azure-NVMe-Conversion-<VMName>-<timestamp>.log` to the current directory |
| `-FixOperatingSystemSettings` | No | off | Actually apply the in-guest driver fix (requires the VM to be running and the VM Agent ready) |
| `-IgnoreOSCheck` | No | off | Skip the in-guest readiness check entirely |
| `-IgnoreSKUCheck` | No | off | Skip target-SKU availability/capability validation |
| `-IgnoreWindowsVersionCheck` | No | off | Skip the Windows Server 2019+ / Windows 10 1809+ requirement check |
| `-IgnoreAzureModuleCheck` | No | off | Skip `Az.Compute`/`Az.Accounts`/`Az.Resources` version checks |
| `-SleepSeconds` | No | `15` | Delay after the VM size/controller update before continuing |

**Example**

```powershell
Connect-AzAccount
.\Azure-NVMe-Conversion.ps1 -ResourceGroupName "myResourceGroup" -VMName "myVM" `
    -NewControllerType NVMe -VMSize "Standard_E4bds_v5" -FixOperatingSystemSettings -StartVM -WriteLogfile
```

**Required modules**

```powershell
Install-Module Az.Accounts  -MinimumVersion 4.0 -Scope CurrentUser
Install-Module Az.Compute   -MinimumVersion 9.0 -Scope CurrentUser
Install-Module Az.Resources -MinimumVersion 7.0 -Scope CurrentUser
```

**Notes**
- Not part of the interactive `menu.ps1` launcher — that menu targets the M365 tenant (Graph/Exchange), this script targets Azure IaaS via the `Az` module and a separate `Connect-AzAccount` session.
- Requires `Virtual Machine Contributor` (or equivalent) on the target resource group.
- Each SCSI boot re-creates the blocking `StartOverride` registry key — don't boot the VM on SCSI between running the fix and converting to NVMe.
