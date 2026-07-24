# AppDeployment

Generic, parameterized device-side app deployment scripts — replacing a large family of old scripts that each hardcoded one vendor's download URL or one shortcut/printer's details. Point every script here at your own package repository / URL; none of them hardcode an internal endpoint.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Install-Win32AppPackage.ps1`](#install-win32apppackageps1) | Download a zipped PSAppDeployToolkit package and run its silent install |
| [`Install-ChocolateyPackage.ps1`](#install-chocolateypackageps1) | Install, upgrade, or uninstall a package via Chocolatey |
| [`Set-DefaultFileAssociation.ps1`](#set-defaultfileassociationps1) | Set the default app for a file extension (PS-SFTA wrapper) |
| [`New-DesktopShortcutsFromStartMenu.ps1`](#new-desktopshortcutsfromstartmenups1) | Copy a set of Start Menu shortcuts to the Public Desktop |
| [`New-DesktopUrlShortcut.ps1`](#new-desktopurlshortcutps1) | Create a `.url` shortcut on the Public Desktop |
| [`Remove-DesktopShortcut.ps1`](#remove-desktopshortcutps1) | Remove desktop shortcuts matching a name pattern |
| [`Add-NetworkPrinterConnection.ps1`](#add-networkprinterconnectionps1) | Add a network printer by IP-based port and driver name |

---

### Install-Win32AppPackage.ps1

Downloads a `.zip` package, expands it, and runs its PSAppDeployToolkit `Deploy-<App>.ps1 -DeploymentType Install -DeployMode NonInteractive` entry point, then cleans up. Optionally registers a logon-triggered scheduled task to re-check for updates on every sign-in (for self-updating desktop clients).

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-AppName` | Yes | Friendly app name (working folder + default deploy script name) |
| `-SourceUri` | Yes | URL to the `.zip` package |
| `-DeployScriptName` | No | Deploy script name inside the package (default: `Deploy-<AppName>.ps1`) |
| `-DeploymentType` | No | PSADT deployment type (default: `Install`) |
| `-DeployMode` | No | PSADT deploy mode (default: `NonInteractive`) |
| `-WorkingRoot` | No | Local working folder (default: `C:\Install`) |
| `-RegisterLogonUpdateTask` | No | Also register a logon scheduled task to re-run this install |
| `-Apply` | No | Actually install (default: preview only) |

**Examples**
```powershell
.\Install-Win32AppPackage.ps1 -AppName "AdobeReaderDC" -SourceUri "https://packages.contoso.com/Installers/AdobeReaderDC.zip" -Apply

.\Install-Win32AppPackage.ps1 -AppName "LineOfBusinessDesktop" `
    -SourceUri "https://packages.contoso.com/Installers/LineOfBusinessDesktop.zip" -RegisterLogonUpdateTask -Apply
```

**Notes:** replaces a whole family of near-identical old scripts (Adobe Reader, AnyDesk, Citrix Workspace, Google Drive, Jabra Direct, TeamViewer, plus a self-updating line-of-business desktop client) that each hardcoded one vendor's URL — same four-step pattern, one script.

---

### Install-ChocolateyPackage.ps1

Installs Chocolatey if missing, then installs/upgrades or uninstalls a named package.

| Parameter | Description |
|-----------|-------------|
| `-PackageName` | Chocolatey package ID (required) |
| `-Uninstall` | Uninstall instead of install/upgrade |
| `-Apply` | Actually run (default: preview only) |

```powershell
.\Install-ChocolateyPackage.ps1 -PackageName git -Apply
.\Install-ChocolateyPackage.ps1 -PackageName git -Uninstall -Apply
```

---

### Set-DefaultFileAssociation.ps1

Sets a default file association via the community [PS-SFTA](https://github.com/DanysysTeam/PS-SFTA) tool, which computes the hash Windows 10 1803+ requires for `UserChoice` changes to stick.

| Parameter | Description |
|-----------|-------------|
| `-ProgId` | ProgId or `Applications\<exe>` to set as default (required) |
| `-Extension` | One or more extensions, e.g. `.pdf` (required) |
| `-Apply` | Actually change (default: preview only) |

```powershell
.\Set-DefaultFileAssociation.ps1 -ProgId "Applications\7zFM.exe" -Extension ".zip",".rar" -Apply
.\Set-DefaultFileAssociation.ps1 -ProgId "Acrobat.Document.DC" -Extension ".pdf" -Apply
```

> Downloads PS-SFTA from GitHub at runtime — vendor it locally first if your policy requires pre-approved script sources.

---

### New-DesktopShortcutsFromStartMenu.ps1

Copies named shortcuts from the all-users Start Menu to the Public Desktop.

```powershell
.\New-DesktopShortcutsFromStartMenu.ps1 -AppNames "Excel","Word","Outlook" -Apply
```

---

### New-DesktopUrlShortcut.ps1

Creates a `.url` desktop shortcut to any web address.

```powershell
.\New-DesktopUrlShortcut.ps1 -Name "Company Portal" -Url "https://portal.contoso.com/" -Apply
```

---

### Remove-DesktopShortcut.ps1

Removes desktop shortcuts matching a wildcard pattern, from the Public Desktop or the current user's desktop.

```powershell
.\Remove-DesktopShortcut.ps1 -NamePattern "*.rdp" -Apply
```

---

### Add-NetworkPrinterConnection.ps1

Adds a TCP/IP printer port + printer connection using an already-installed driver.

| Parameter | Description |
|-----------|-------------|
| `-PrinterName` | Display name (required) |
| `-PortAddress` | Printer IP/hostname (required) |
| `-DriverName` | Already-installed driver name (required) |
| `-PortName` | Port object name (default: `IP_<PortAddress>`) |
| `-Apply` | Actually create (default: preview only) |

```powershell
.\Add-NetworkPrinterConnection.ps1 -PrinterName "Label Printer - Warehouse" -PortAddress "10.0.5.50" -DriverName "Dymo LabelWriter 450 Turbo" -Apply
```

> Install the printer driver first — this script only creates the port and connection.

---

All scripts default to a dry run; pass `-Apply` to make changes, per house style.
