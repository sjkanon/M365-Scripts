**English** · [Nederlands](readme.nl.md) · [Français](readme.fr.md)

[M365-Scripts](../../../readme.md) › [scripts](../../readme.md) › [Device](../readme.md) › **Printer**

# Printer

Installs printer drivers and TCP/IP printers from one JSON file. The drivers are downloaded from a GitHub repository (release asset or folder), any https URL or a share.

[`Export-Printer.ps1`](#export-printerps1) works the other way round: it reads the printers of an existing machine and writes that same JSON, with the drivers beside it, into one zip.

Built to run unattended after a golden image has provisioned a new server or session host. On first boot it waits for the Print Spooler and the network instead of failing, and because every run is idempotent the same command can also run at every startup.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Export-Printer.ps1`](Export-Printer.ps1) ([docs](#export-printerps1)) | Export the printers of a machine - settings as a JSON file, drivers, full print tickets - to one zip in `C:\Temp`, ready for `Install-Printer.ps1` on another one |
| [`Install-Printer.ps1`](Install-Printer.ps1) ([docs](#install-printerps1)) | Install printer drivers (downloaded from GitHub) and printers as described in a JSON file |
| [`New-PrinterGroup.ps1`](New-PrinterGroup.ps1) ([docs](#new-printergroupps1)) | One AD or Entra ID security group per printer (`SG-U-PRINTM-<printer>`) from a printer x user matrix, members added, and the group SIDs written into `printers.json` as permissions |

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
| `inputBin` | Default tray, by the name the driver shows (`Tray 2`, `Manual Feed`) or its Print Schema name (`ns0000:Tray2`); spaces and case do not matter. A name the driver does not have is a warning that lists the trays it does offer |
| `permissions` | Who may use this queue: a list of `{ "principal": "CONTOSO\\Finance", "access": "print" }` — see *Security groups per queue* below. Without the field the permissions are left alone |
| `ensure` | `absent` removes the printer (and its port when no other printer uses it). The driver always stays |

**One device, several queues (trays)**

Users pick a queue, not a tray. Give each queue its own entry with the **same address**: they share the `IP_<address>` port, and each gets its own default tray, duplex or colour:

```json
"printers": [
  { "name": "Office",            "driver": "HP Universal Printing PCL 6", "address": "10.0.5.20", "inputBin": "Tray 1" },
  { "name": "Office - letterhead", "driver": "HP Universal Printing PCL 6", "address": "10.0.5.20", "inputBin": "Tray 2", "duplex": "OneSided" },
  { "name": "Office - envelopes", "driver": "HP Universal Printing PCL 6", "address": "10.0.5.20", "inputBin": "Manual Feed" }
]
```

The tray is set in the queue's default print ticket, so it is what every user of that queue starts from; a user can still pick another tray in a single print dialog. Tray names differ per driver - run the script once with a name you expect, and a wrong one is answered with the list the driver offers. Universal drivers only show the trays they know the device has once the device's options are known; if only `Auto Select` comes back, set the installable options in the printer's properties first, or use the vendor's model-specific driver.

**Security groups per queue**

A queue can be limited to one or more groups, for example a cheque tray only Finance may print to:

```json
{ "name": "Office - cheques", "driver": "HP Universal Printing PCL 6", "address": "10.0.5.20", "inputBin": "Tray 3",
  "permissions": [
    { "principal": "CONTOSO\\Finance",          "access": "print" },
    { "principal": "CONTOSO\\Office Managers",  "access": "manageDocuments" },
    { "principal": "S-1-12-1-1111-2222-3333-4444", "access": "print" }
  ] }
```

| `access` | Means |
|----------|-------|
| `print` | Print and manage one's own jobs — what Everyone has by default |
| `manageDocuments` | Also pause, restart and delete other people's jobs |
| `manage` | Also change the printer itself (settings, sharing, permissions) |

- `principal` is `DOMAIN\Group`, a local user or group, or a **SID**. An Entra ID group can only be given by its SID (`S-1-12-1-…`, from Entra ID or `whoami /groups` of a member). `Everyone`, `Authenticated Users`, `Users`, `Interactive`, `Administrators` and `SYSTEM` are understood in English on a Windows of any language
- Listed principals get that access; **Everyone and every other user or group lose theirs**, including the account that happened to install the printer. Administrators, SYSTEM, CREATOR OWNER, app packages and service identities always keep their access — printing from packaged apps such as new Outlook depends on those entries, so they are never removed. `[]` means administrators only. A `print` principal that does not resolve stops that queue (fail closed); a `manage` or `manageDocuments` one - such as the admin account an exported queue was created with - is left out with a warning and the queue goes in
- **Fails closed:** a group that does not resolve stops that queue from being created (`-CheckOnly` reports it in advance), and a new queue whose permissions cannot be set is removed again rather than left open to everyone. An existing queue keeps its old permissions and the run fails. On a first boot an unresolved name is retried for up to 2 minutes, for a domain controller that is not reachable yet
- Compared by meaning, not by text: Windows may store an entry differently from how it was given, and that does not count as a change. The permissions are checked on every run, so a group added by hand in the printer's properties is taken away again — the JSON is the source of truth
- The permissions decide who may print to the queue. Whether a user without access still *sees* it in the list depends on the application and is not something this script controls

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
| `-SkipPermissions` | Ignore every `permissions` field in the JSON: install without groups and leave each queue's ACL alone (Everyone may print) |

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
- NinjaOne script variables: `configPath`, `printer`, `githubToken`, `workingDir`, `logPath`, `waitSeconds`, `proxy`, and the checkboxes `checkOnly`, `whatIf`, `quiet`, `force`, `skipSignatureCheck`, `skipPermissions`
- `-CheckOnly` is the health check for an RMM condition or a scheduled detection job: exit `2` means the device does not match the JSON

---

### Export-Printer.ps1

The reverse of `Install-Printer.ps1`: reads the printers of this machine and writes one zip, by default `C:\Temp\PrinterExport_<computer>_<yyyyMMdd-HHmmss>.zip`. Nothing on the machine is changed.

| In the zip | Contents |
|------------|----------|
| `printers.json` | The printers and their settings in Install-Printer's format: driver, address, port name/number, LPR queue, SNMP, location, comment, sharing, duplex, colour, paper size, default tray (`inputBin`) and permissions. Each driver points at its own folder in the zip (`"type": "path"`) |
| `drivers\<driver name>\` | The complete driver package as the driver store holds it: INF, signed catalog and binaries - what `pnputil` installs from. Drivers that ship with Windows (`ntprint.inf`, `prnms*.inf`) are not copied |
| `settings\<printer>.xml` | The full default print ticket per printer, including every vendor setting the JSON fields do not cover |
| `inventory.json` | Everything that was read, also for printers the JSON cannot describe (WSD, USB, PDF, OneNote, fax, print server connections), with the reason they are not in `printers.json` and the raw `PermissionSDDL` |
| `<computer>.printerExport` | Only with `-PrintBrm`: Windows' own Printer Migration backup |

On the target machine: unpack the zip and run `.\Install-Printer.ps1 -ConfigPath <folder>\printers.json -CheckOnly`, then without `-CheckOnly`. The relative driver paths resolve against the folder of the JSON.

**Parameters**

| Parameter | Description |
|-----------|-------------|
| `-Printer` | Only these printers (names, wildcards allowed). Default: all |
| `-OutputPath` | Folder for the zip (default: `C:\Temp`), created when missing |
| `-NoDriver` | Settings only, no driver files - `printers.json` then lists no drivers, so the target needs them already |
| `-PrintBrm` | Also make a `PrintBrm.exe -b` backup and put it in the zip. Needs administrator rights |

**Examples**

```powershell
# Every printer, its settings and drivers to C:\Temp
.\Export-Printer.ps1

# Only the Office printers, settings only
.\Export-Printer.ps1 -Printer 'Office*' -NoDriver

# Also a PrintBrm backup, the zip on a share
.\Export-Printer.ps1 -PrintBrm -OutputPath \\fs01\it$\exports
```

**Notes**
- Only TCP/IP printers (Standard TCP/IP Port) go into `printers.json`, as that is what Install-Printer creates. A printer it would refuse - a name with `\ , !`, an address that is neither an IP address nor a host name - is reported and kept in `inventory.json` only
- A printer on a vendor port monitor (EpsonNet Print, IPP) whose port name holds its address - `10.1.50.235:`, `http://10.0.5.20:631/ipp` - is exported as a Standard TCP/IP port (RAW 9100) to that address, with its driver and settings, a warning and a note in its `description`: that port monitor only exists where the vendor's software is installed. Print a test page on the target
- Fields are written only where they differ from Install-Printer's defaults (port `IP_<address>`, port 9100, SNMP off), so the file reads like one written by hand
- Permissions are exported only when they differ from the Windows default (Everyone may print). Administrators, SYSTEM, CREATOR OWNER and app identities are left out, as Install-Printer always keeps them. So is "manage" for the account the queue was created with: Windows grants it to the creator, so it is left out when that account is in the local Administrators group (directly or through a domain group), or is a local account that no longer exists (the setup account of an image). A queue with nothing else on it than Everyone printing then gets no `permissions` at all. Well-known groups are written by their English name, an Entra ID user or group by its SID (an `AzureAD\name` resolves nowhere else), a local account of this machine with a warning
- Print defaults are read in a job with a 60-second timeout, as some universal drivers hang in `Get-PrintConfiguration`; a failure there is a warning, and that printer is exported without them
- Admin rights are not needed, except for `-PrintBrm`. Started 32-bit, the script relaunches itself 64-bit, because the driver store is not visible from SysWOW64
- NinjaOne script variables: `printer`, `outputPath`, and the checkboxes `noDriver`, `printBrm`

---

### New-PrinterGroup.ps1

Turns a printer x user matrix - as copied out of Excel - into one security group per printer, in **Active Directory** or **Entra ID**, named `<Prefix><printer name>` with spaces removed (`Office - Duplex` → `SG-U-PRINTM-Office-Duplex`), and adds the ticked users. With `-ConfigPath` it writes each group into that printer's entry in `printers.json` as `"permissions": [ { "principal": <the group>, "access": "print" } ]`, so `Install-Printer.ps1` lets that group print and takes printing away from Everyone; `manage` entries are kept. Preview by default; `-Apply` changes the directory and the file.

**AD or Entra ID** - a printer ACL is checked against the user's Windows sign-in token, so the group must be one that token carries:

| | Users sign in with | Groups | In `printers.json` |
|---|---|---|---|
| `-Directory AD` | an AD account: domain-joined or hybrid-joined server or session host, classic print server | created in `-OU` - pick one Entra Connect synchronises and they appear in Entra ID too | `DOMAIN\SG-U-PRINTM-...` |
| `-Directory Entra` | an Entra ID account on an Entra-joined device | cloud groups | their SID `S-1-12-1-...`, computed from the object id |
| `-Directory Entra -AdWriteback` | an AD account (hybrid), but the groups are managed in Entra ID | cloud groups that Entra Cloud Sync provisions to AD ("Microsoft Entra ID to AD sync") | the AD copy, `DOMAIN\<its sAMAccountName>` |

In a hybrid environment that is `-Directory AD`: a cloud group never reaches the token of an AD sign-in, and the queue would refuse everyone. `-Directory Entra` warns when the matrix users are synced from AD. Either way a new member gets access at their next sign-in.

The matrix: first column the printer names, header row the users (display name, UPN, or `DOMAIN\account` for AD), any non-empty cell (`x`) means access. Tab, semicolon or comma separated; a header without the empty top-left cell, as Excel copies it, is fine. `Name (DEPARTMENT)` also matches the display name `Name`. A name that matches no user, or several, is reported and left out.

**Parameters**

| Parameter | Description |
|-----------|-------------|
| `-MatrixPath` | The matrix file |
| `-Directory` | `AD` or `Entra` (default) |
| `-OU` | AD: distinguished name of the OU for new groups. Required as soon as a group has to be created |
| `-AdWriteback` | Entra: the groups reach AD through Cloud Sync; put their AD copies in `printers.json` and check that the members arrived. Needs Graph (PowerShell 7) and the ActiveDirectory module on the same domain-joined machine |
| `-Prefix` | Group name prefix (default: `SG-U-PRINTM-`). AD caps a group name at 64 characters; longer ones stop the run before anything changes |
| `-ConfigPath` | `printers.json` to write the permissions into. `manage`/`manageDocuments` entries are kept, `Everyone` and other `print` entries replaced; a `.bak` is written first. A group is only written once it exists - Install-Printer refuses to create a queue whose group does not resolve |
| `-RemoveExtraMembers` | Also remove group members the matrix does not tick (default: reported, left alone) |
| `-TenantId` | Entra: target tenant; `-TenantId` defaults to the GDAP customer tenant (`$global:cid` / `$env:M365_CUSTOMER_TENANTID`) |
| `-Apply` | Create groups, change members, write the JSON. Without it: report only |

**Examples**

```powershell
# Hybrid / on-premises: what would be created in AD
.\New-PrinterGroup.ps1 -MatrixPath .\printer-matrix.tsv -Directory AD

# Create the AD groups and make them the queues' print permission
.\New-PrinterGroup.ps1 -MatrixPath .\printer-matrix.tsv -Directory AD -OU 'OU=Printers,OU=Groups,DC=contoso,DC=local' -ConfigPath .\printers.json -Apply

# Entra-joined devices: cloud groups
.\New-PrinterGroup.ps1 -MatrixPath .\printer-matrix.tsv -TenantId contoso.com -ConfigPath .\printers.json -Apply
```

**Notes**
- **Cloud Sync for `-AdWriteback`**: put the groups in the configuration's scope (selected groups, or all groups with a filter such as `displayName` REGEX MATCH `^SG-U-PRINTM-`); turn on **Provision membership to on-premises users** - it is off by default, and without it the AD copies stay empty; and add a `sAMAccountName` mapping (e.g. `[displayName]`), or the AD copy gets a random name - which works, but reads badly in an ACL. The AD copy is found by `displayName` and its `msDS-ExternalDirectoryObjectId`; a group not written back yet is reported and left out of the JSON - run again after the next provisioning cycle
- **Where**: the run names the target first - the AD domain and the domain controller it uses (the PDC emulator, so a group created a moment ago is there when members are added), or the Entra tenant - and `-Apply` asks before creating anything (`-Confirm:$false` skips that). Without `-TenantId` Entra uses the GDAP customer, else the tenant you sign in to, or one an existing Graph session in that window is already in
- `-Directory AD` needs the ActiveDirectory module (RSAT) on a domain-joined machine and runs in Windows PowerShell 5.1; `-Directory Entra` needs PowerShell 7 and Microsoft Graph (`Group.ReadWrite.All`, `User.Read.All`)
- `-Directory Entra` does not use `Connect-M365.ps1`: sign-in (GDAP or a normal admin account) is done beforehand. An existing Graph session for the target tenant is reused; otherwise the script signs in interactively with `Connect-MgGraph` itself (for example as `admin@tenant.onmicrosoft.com`) and signs out again at the end
- Keep the matrix out of this repository: it names real people. The gitignored `local/` folder is a fine place for it
