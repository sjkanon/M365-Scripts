# UniFi

Tooling for a UniFi Network Controller or UniFi OS console (UDM/UDM-Pro/UDR). Talks directly to the controller's API — not part of [`menu.ps1`](../../../menu.ps1), since each run needs a controller URL and credentials. Credentials are always requested via `-Credential`/`Get-Credential`, never hardcoded.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Get-UnifiNetworkReport.ps1`](#get-unifinetworkreportps1) | Generate an HTML network documentation report (devices, firmware, uptime) |
| [`Update-UnifiFirmware.ps1`](#update-unififirmwareps1) | List and optionally trigger firmware upgrades across sites |
| `UnifiApi.ps1` | Shared login/session helper, dot-sourced automatically by the two scripts above — not meant to be run directly |

---

### Get-UnifiNetworkReport.ps1

Read-only. Logs in, enumerates every site (or one with `-Site`), lists adopted devices per site, and writes an HTML report (name, model, MAC, IP, firmware version, state, uptime) to `C:\Temp\`.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-Controller` | Yes | Controller base URL, e.g. `https://192.168.1.1` or `https://unifi.contoso.local:8443` |
| `-Credential` | No | Admin credentials — prompted via `Get-Credential` if omitted |
| `-Site` | No | Limit to one site by name. If omitted, all sites are included |
| `-SkipCertificateCheck` | No | Accept self-signed/untrusted certificates (common for on-prem controllers) |
| `-OutputPath` | No | Report folder (default: `C:\Temp\` / `~/Downloads`) |

**Examples**

```powershell
.\Get-UnifiNetworkReport.ps1 -Controller "https://192.168.1.1" -SkipCertificateCheck
.\Get-UnifiNetworkReport.ps1 -Controller "https://unifi.contoso.local:8443" -Site "Head Office"
```

---

### Update-UnifiFirmware.ps1

Lists devices per site with a firmware upgrade available (per the controller's own `upgradable` flag). Runs as a dry-run by default — no upgrade is triggered without `-Apply`. Upgrading reboots the device, causing a brief outage for anything connected through it — per-device confirmation is required unless `-Force` is also passed.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-Controller` | Yes | Controller base URL |
| `-Credential` | No | Admin credentials — prompted via `Get-Credential` if omitted |
| `-Site` | No | Limit to one site by name. If omitted, all sites are processed |
| `-SkipCertificateCheck` | No | Accept self-signed/untrusted certificates |
| `-Apply` | No | Actually trigger upgrades (default: dry run) |
| `-Force` | No | Skip per-device confirmation when used with `-Apply` |
| `-OutputPath` | No | Report folder (default: `C:\Temp\` / `~/Downloads`) |

**Examples**

```powershell
# Dry run across all sites
.\Update-UnifiFirmware.ps1 -Controller "https://192.168.1.1" -SkipCertificateCheck

# Upgrade one site, confirm per device
.\Update-UnifiFirmware.ps1 -Controller "https://192.168.1.1" -Site "Head Office" -Apply

# Unattended, all sites — use with care, devices reboot
.\Update-UnifiFirmware.ps1 -Controller "https://192.168.1.1" -Apply -Force
```

**Notes**
- Both scripts support the classic self-hosted UniFi Network Controller (`/api/login`) and UniFi OS consoles (`/api/auth/login` + `/proxy/network/...`) — login auto-detects which one is in front of it.
- `-SkipCertificateCheck` is implemented for both PowerShell 7+ (native parameter) and Windows PowerShell 5.1 (temporary certificate validation callback, reset immediately after the request).
- A CSV/HTML report is always written after each run, even for a dry run.
