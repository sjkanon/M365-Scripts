**English** · [Nederlands](readme.nl.md) · [Français](readme.fr.md)

[M365-Scripts](../../../readme.md) › [scripts](../../readme.md) › [Device](../readme.md) › **Printer**

# Printer

Installs printer drivers and TCP/IP printers from one JSON file. The drivers are downloaded from a GitHub repository (release asset or folder), any https URL or a share.

Built to run unattended after a golden image has provisioned a new server or session host. On first boot it waits for the Print Spooler and the network instead of failing, and because every run is idempotent the same command can also run at every startup.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Install-Printer.ps1`](Install-Printer.ps1) ([docs](#install-printerps1)) | Install printer drivers (downloaded from GitHub) and printers as described in a JSON file |

## Files

| File | Description |
|------|-------------|
| [`printers.example.json`](printers.example.json) | Example configuration showing every field and every driver source |

---

### Install-Printer.ps1

Per run:

1. **Config** — read the JSON (local, UNC or https URL) and pick the printers asked for with `-Printer`. The whole file is validated before anything is downloaded
2. **Preflight** — start the Print Spooler if needed and wait for it, then compare every driver (installed? which version?) and every printer (port, driver, settings) with the JSON
3. **Download** — only for a driver that is missing or older than the JSON's `version`. An optional `sha256` is checked, and `.zip`/`.cab` files are extracted
4. **Driver** — find the INF (named in the JSON, or the one that declares the driver name), check the catalog signature, then run `pnputil /add-driver /install` + `Add-PrinterDriver`
5. **Printer** — create the TCP/IP port, add the printer or correct its driver/port, and set location, comment, sharing and print defaults (duplex, color, paper size)
6. **Verify** — read everything back

A device that is already in order costs one read of the JSON. Nothing is downloaded and nothing changes.

**The JSON**

```json
{
  "drivers": [
    {
      "name": "HP Universal Printing PCL 6",
      "version": "7.2.0.25780",
      "inf": "hpcu270u.inf",
      "source": { "type": "githubRelease", "repository": "contoso/printer-drivers",
                  "tag": "latest", "asset": "hp-upd-pcl6-x64-*.zip" }
    }
  ],
  "printers": [
    { "name": "Office 1st floor", "driver": "HP Universal Printing PCL 6",
      "address": "10.0.5.20", "location": "1st floor", "duplex": "TwoSidedLongEdge" }
  ]
}
```

| Driver field | Description |
|--------------|-------------|
| `name` | Exact driver name from the INF (as `Get-PrinterDriver` shows it). Printers refer to it |
| `version` | Optional: an installed driver older than this is updated. Without it, an installed driver is left alone (unless `-Force` is given) |
| `inf` | Optional: INF name or path within the package (wildcards allowed). Without it, the script takes the INF that mentions `name`, preferring the x64/arm64 folder |
| `install` | Optional `true`: install even when none of the selected printers uses it |
| `source` | Where the files come from — see below |

| Source `type` | Fields |
|---------------|--------|
| `githubRelease` | `repository` (`owner/name`), `tag` (default `latest`), `asset` (name, wildcards allowed — must match exactly one) |
| `github` | `repository`, `path` (a `.zip` or a folder holding the INF), `ref` (branch/tag/commit, default: the default branch) |
| `url` | `url` (https) — a `.zip`, `.cab` or single file |
| `path` | `path` to a folder or `.zip`; relative paths resolve against the folder of the JSON |
| *(all)* | `sha256` — optional, checked before extracting (single files only) |

| Printer field | Description |
|---------------|-------------|
| `name`, `driver`, `address` | Required (`address` = IP or hostname) |
| `portName` | Default `IP_<address>` |
| `portNumber` | RAW port, default `9100` |
| `lprQueue` | Use LPR with this queue instead of RAW |
| `snmp` | `true` turns SNMP status on. Off by default, so a printer that does not answer SNMP is not shown as Offline |
| `location`, `comment` | Shown to users |
| `shared`, `shareName` | Share the printer (share name defaults to `name`) |
| `duplex` | `OneSided`, `TwoSidedLongEdge` or `TwoSidedShortEdge` |
| `color` | `true` / `false` |
| `paperSize` | e.g. `A4`, `Letter` |
| `ensure` | `absent` removes the printer (and its port when no other printer uses it). The driver always stays |

**Parameters**

| Parameter | Description |
|-----------|-------------|
| `-ConfigPath` | The JSON: local/UNC path or https URL (e.g. a raw link in the same GitHub repository) |
| `-Printer` | Only these printers from the JSON (names, wildcards allowed). Default: all |
| `-GitHubToken` | Token for a private repository (falls back to `$env:GITHUB_TOKEN`). Only sent to GitHub's own hosts, never to a `url` source |
| `-Proxy` | Proxy for every download, e.g. `http://proxy.contoso.local:8080`. As SYSTEM there is no user proxy to inherit; uses the computer account's credentials |
| `-WorkingDir` | Download/extract folder (default: `C:\IT\Printers`) |
| `-LogPath` | Folder for `Install-Printer.log` (default: `C:\Temp`), appended by every run that may change something and rotated past 1 MB |
| `-WaitSeconds` | How long a freshly provisioned machine gets for the spooler, the network and another run of this script to finish (default: `300`, `0` = no waiting) |
| `-CheckOnly` | Report only, change nothing. Exit code `2` means work is due |
| `-Quiet` | Print nothing unless there is work or a failure |
| `-Force` | Reinstall drivers even when the version matches |
| `-SkipSignatureCheck` | Skip the script's own catalog signature check (Windows still refuses unsigned package drivers) |

**Examples**

```powershell
# What would be done - nothing is changed
.\Install-Printer.ps1 -ConfigPath .\printers.json -CheckOnly

# Only the Office printers, as a dry run, JSON straight from GitHub
.\Install-Printer.ps1 -ConfigPath https://raw.githubusercontent.com/contoso/printer-drivers/main/printers.json -Printer 'Office*' -WhatIf

# First boot or deployment (Custom Script Extension, Run Command, startup task as SYSTEM)
powershell.exe -ExecutionPolicy Bypass -File C:\IT\Install-Printer.ps1 -ConfigPath https://raw.githubusercontent.com/contoso/printer-drivers/main/printers.json -Quiet

# Private repository
$env:GITHUB_TOKEN = '<fine-grained token, Contents: read>'
.\Install-Printer.ps1 -ConfigPath \\fs01\it$\printers.json -Confirm:$false
```

**Installing at deployment instead of in the image**

The image stays free of drivers and printers; every new session host gets them when it is deployed, so a changed printer is a change to the JSON, not a new image. Run the script once as SYSTEM through the Azure Custom Script Extension in the VM's Bicep (every host the pool adds runs it automatically), or through Run Command on a host that already exists:

```bicep
// Custom Script Extension: runs once, as SYSTEM, when the VM is deployed
resource installPrinters 'Microsoft.Compute/virtualMachines/extensions@2024-07-01' = {
  parent: vm
  name: 'InstallPrinters'
  location: location
  properties: {
    publisher: 'Microsoft.Compute'
    type: 'CustomScriptExtension'
    typeHandlerVersion: '1.10'
    autoUpgradeMinorVersion: true
    settings: {
      fileUris: [
        'https://raw.githubusercontent.com/sjkanon/M365-Scripts/<commit>/scripts/Device/Printer/Install-Printer.ps1'
      ]
    }
    protectedSettings: {
      commandToExecute: 'powershell.exe -ExecutionPolicy Bypass -File Install-Printer.ps1 -ConfigPath https://raw.githubusercontent.com/contoso/printer-drivers/main/printers.json -Quiet'
    }
  }
}
```

```powershell
# Run Command: afterwards, on a VM that is already running
az vm run-command invoke -g <resource-group> -n <vm> --command-id RunPowerShellScript `
  --scripts '@Install-Printer.ps1' `
  --parameters 'ConfigPath=https://raw.githubusercontent.com/contoso/printer-drivers/main/printers.json'
```

- Point `fileUris` at a **commit** rather than `main`, so a host deployed next month runs the script you tested. The JSON may stay on a branch: it is data, validated in full before anything changes.
- **No `-Confirm:$false`** on these command lines: a run without a console never asks, and through `-File` the switch would arrive as the text `'$false'` and stop the script before it starts.
- A VM holds **one** Custom Script Extension. If the deployment already uses it for something else, use a managed Run Command (`Microsoft.Compute/virtualMachines/runCommands`) for this instead.
- A GitHub token for a private driver repository belongs in `protectedSettings` (`-GitHubToken`), never in `settings`, which is readable on the VM.
- Exit code `1` fails the extension, so a deployment with a printer that could not be installed is visible in the portal; `Install-Printer.log` in `-LogPath` says which one.

**Exit codes**

| Code | Meaning |
|------|---------|
| `0` | Everything as configured, or installed successfully |
| `1` | Failure (a driver or printer that could not be installed, or an invalid JSON) |
| `2` | `-CheckOnly` only: work is due |

**What prevents a failed run**
- **The whole JSON is checked before anything happens**: required fields, unknown fields (a typo such as `adress` or `loaction` is an error, not silently ignored), duplicate names, characters Windows refuses in a printer name, IP/host names, port numbers, `duplex`/`ensure` values, `true`/`false` written as text, and two printers on one port with different addresses. Every problem is listed at once, and nothing is changed
- **The INF is checked before pnputil sees it**: printer class, a section for this architecture (an x86-only package on an x64 server is named as such), the exact driver name among the models it declares — with the names it *does* declare in the error, closest first — and a signed catalog for this architecture
- **Downloads are checked for what they are**: a proxy block page, login portal or GitHub error saved as `.zip` is refused, as is a file that is not a zip or cab; there has to be 1 GB free before anything is fetched
- **One run at a time**: a startup task and an RMM job that overlap wait for each other on a machine-wide lock (up to `-WaitSeconds`) instead of installing the same driver twice. A run that was killed half-way is detected and the next one continues from what it left
- **The spooler**: on first boot it is started and waited for until it actually answers, not just until it says Running. When it stops or stalls during an install — common right after a vendor driver lands — it is restarted and that one step is tried once more. It is deliberately *not* restarted after every driver, as some published scripts do: on a session host in use that interrupts everyone's printing
- **The JSON from a URL is cached** in `-WorkingDir`. When the URL cannot be reached, the last good copy is used with a warning, so a host that boots while GitHub is down keeps its printers
- **Driver vendor code is fenced off**: reading and writing the print defaults run in a job with a timeout, because some universal drivers hang there
- **One failing driver does not stop the rest**: its printers are skipped and reported, all others are installed, and the exit code is `1`. The downloaded files of the failed driver are kept for inspection; after a successful install they are removed (the driver store keeps its own copy)
- **A changed address** is applied: with default port names the printer moves to the new `IP_<address>` port and the old one is removed once unused; a named port used only by this printer is rebuilt in place. A port shared with other printers is not changed behind their back
- **Detection**: a clean run writes `ConfigSha256`, `ConfigPath`, `LastSuccess` and `Printers` under `HKLM:\SOFTWARE\M365-Scripts\InstallPrinter`. An Intune detection rule or RMM condition can compare the hash with the JSON's, which tells "installed with the current configuration" apart from "installed with last month's"
- **Everything is logged** to `Install-Printer.log` in `-LogPath`, also when the run fails early, so an unattended first boot can be read back

**Notes**
- **After a golden image:** run it as SYSTEM from the Azure Custom Script Extension, a startup scheduled task baked into the image, `SetupComplete.cmd`, Intune or an RMM. If the Print Spooler has not started yet, the script starts it and waits. Downloads that fail on DNS or a timeout are retried with backoff, both within `-WaitSeconds`. A 401/403/404 fails at once, because waiting does not change it. A spooler that the image *disabled* (PrintNightmare hardening) is reported, not silently enabled
- Printers are created machine-wide, so on an RDS/AVD session host every user sees them. No per-user Point and Print step is involved, so the `RestrictDriverInstallationToAdministrators` restriction (KB5005652) does not apply — the driver is installed by an administrator/SYSTEM
- **GitHub:** the release metadata and the folder listing go through the API (one or two requests per driver). The files themselves come from the release's download link and `raw.githubusercontent.com` without a token — neither counts toward the 60 API requests an hour GitHub allows per public IP, so a pool of hosts booting behind one NAT does not run out. With a token everything goes through the API, which is what a private repository needs. A private repository answers 404 without a token, and the error says so. A repository too large for one tree listing, or a file stored through Git LFS, is refused rather than half-downloaded — publish the driver as a release asset instead
- Only INF-based driver packages are supported. A vendor `.exe` setup is not — extract the package (most vendors offer a "driver only" zip) and put that in the repository. The catalog (`.cat`) must carry a valid signature
- `pnputil` exit codes `0`, `259` (no device waiting — normal for printers), `3010` and `1641` (reboot) count as success. Anything else points at `C:\Windows\INF\setupapi.dev.log`
- There is no `Set-PrinterPort`: a port that already exists for another address is reported, not re-created, because other printers may use it. Give the printer its own `portName`
- `Set-PrintConfiguration` runs the vendor driver's own code and hangs on some universal drivers, so it runs in a job with a 2-minute timeout. A failure there is a warning — the printer is installed all the same
- Default printer is a per-user setting and is deliberately not handled: as SYSTEM there is no user to set it for
- Started 32-bit (Intune Management Extension, some RMM agents), the script relaunches itself 64-bit, because `pnputil` does not exist under SysWOW64. Started by hand without elevation, it asks for it
- NinjaOne script variables: `configPath`, `printer`, `githubToken`, `workingDir`, `logPath`, `waitSeconds`, `proxy`, and the checkboxes `checkOnly`, `whatIf`, `quiet`, `force`, `skipSignatureCheck`
- `-CheckOnly` is the health check for an RMM condition or a scheduled detection job: exit `2` means the device does not match the JSON
