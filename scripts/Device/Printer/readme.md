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
| `-WorkingDir` | Download/extract folder (default: `C:\IT\Printers`) |
| `-LogPath` | Folder for the transcript of a run that changes something (default: `C:\Temp`) |
| `-WaitSeconds` | How long a freshly provisioned machine gets for the spooler and the network (default: `300`, `0` = no waiting) |
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

# First boot of a server from a golden image (Custom Script Extension / startup task as SYSTEM)
powershell.exe -ExecutionPolicy Bypass -File C:\IT\Install-Printer.ps1 -ConfigPath https://raw.githubusercontent.com/contoso/printer-drivers/main/printers.json -Quiet -Confirm:$false

# Private repository
$env:GITHUB_TOKEN = '<fine-grained token, Contents: read>'
.\Install-Printer.ps1 -ConfigPath \\fs01\it$\printers.json -Confirm:$false
```

**Exit codes**

| Code | Meaning |
|------|---------|
| `0` | Everything as configured, or installed successfully |
| `1` | Failure (a driver or printer that could not be installed, or an invalid JSON) |
| `2` | `-CheckOnly` only: work is due |

**Notes**
- **After a golden image:** run it as SYSTEM from the Azure Custom Script Extension, a startup scheduled task baked into the image, `SetupComplete.cmd`, Intune or an RMM. If the Print Spooler has not started yet, the script starts it and waits. Downloads that fail on DNS or a timeout are retried with backoff, both within `-WaitSeconds`. A 401/403/404 fails at once, because waiting does not change it. A spooler that the image *disabled* (PrintNightmare hardening) is reported, not silently enabled
- Printers are created machine-wide, so on an RDS/AVD session host every user sees them. No per-user Point and Print step is involved, so the `RestrictDriverInstallationToAdministrators` restriction (KB5005652) does not apply — the driver is installed by an administrator/SYSTEM
- **GitHub:** a release asset is downloaded through the API asset URL, so the same code works with and without a token. A folder is listed once through the git tree API and every file under it is fetched. Unauthenticated, GitHub allows 60 API requests an hour per IP, and a folder costs one request per file, so for many hosts behind one NAT use a release `.zip` or a token. A private repository answers 404 without a token, and the error says so
- Only INF-based driver packages are supported. A vendor `.exe` setup is not — extract the package (most vendors offer a "driver only" zip) and put that in the repository. The catalog (`.cat`) must carry a valid signature
- `pnputil` exit codes `0`, `259` (no device waiting — normal for printers), `3010` and `1641` (reboot) count as success. Anything else points at `C:\Windows\INF\setupapi.dev.log`
- There is no `Set-PrinterPort`: a port that already exists for another address is reported, not re-created, because other printers may use it. Give the printer its own `portName`
- `Set-PrintConfiguration` runs the vendor driver's own code and hangs on some universal drivers, so it runs in a job with a 2-minute timeout. A failure there is a warning — the printer is installed all the same
- Default printer is a per-user setting and is deliberately not handled: as SYSTEM there is no user to set it for
- Started 32-bit (Intune Management Extension, some RMM agents), the script relaunches itself 64-bit, because `pnputil` does not exist under SysWOW64. Started by hand without elevation, it asks for it
- NinjaOne script variables: `configPath`, `printer`, `githubToken`, `workingDir`, `logPath`, `waitSeconds`, and the checkboxes `checkOnly`, `whatIf`, `quiet`, `force`, `skipSignatureCheck`
- `-CheckOnly` is the health check for an RMM condition or a scheduled detection job: exit `2` means the device does not match the JSON
