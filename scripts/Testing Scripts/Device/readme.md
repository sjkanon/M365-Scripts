# Testing — Device

Diagnostic scripts for Windows endpoints.
All scripts require administrator privileges.

---

## Scripts

### vias_archiver.ps1

Teams archiver with Graph, Teams and SharePoint export flow.

Current behavior (v8.6):
- Creates a unique temporary Entra app registration for the run.
- Grants only required delegated setup permissions during bootstrap.
- Applies Graph/SharePoint delegated consent to that temporary app.
- Removes the temporary app and service principal at cleanup (and on key setup failures).
- Registers an exit cleanup hook so the temporary app is also removed on PowerShell exit/Ctrl+C.
- Checks Teams/SharePoint folder access first during file export, and conditionally grants higher Graph rights when access is denied.
- Resolves channel file locations via Graph filesFolder for all channel types (standard/private/shared), with channel caching and fallback lookup.
- Handles SharePoint NotFound during export as a controlled skip instead of noisy hard failures.
- Downloads files with per-file retries, reconnect fallback, and post-download count validation to ensure completeness.
- Does not archive Teams by default; archiving now requires explicit confirmation during Step 10.

### Test-OpenVpnDiagnostics.ps1

Collects and evaluates diagnostic information for OpenVPN Connect issues on a Windows machine. Checks each relevant layer from driver to network and reports any problems found.

**Checks performed**

| Section | What is checked |
|---------|----------------|
| Wintun / TAP adapters | PnP device status — flags anything not `OK` |
| Virtual network adapters | Adapter visibility — warns if none found while VPN should be active |
| Network profiles | Flags VPN adapters set to `Public` (should be `Private`) |
| Installed VPN software | Lists all VPN-related apps — warns about potential conflicts with OpenVPN Connect |
| Hyper-V / WSL / virtualisation | Lists enabled features — warns if Hyper-V is active (can conflict with Wintun) |
| OpenVPN service | Service status and start type — flags if not running |
| Active routes | Routes via VPN adapter — warns if adapter exists but no routes are present |
| DNS configuration | DNS servers per active adapter |
| Event Log | Last 20 OpenVPN entries from the Application log |

Results are printed to screen with a summary of all issues at the end.

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-ExportTxt` | No | Save the full report to a txt file |
| `-OutputPath` | No | Custom path for the report (implies `-ExportTxt`). Default: `C:\Temp\OpenVpnDiagnostics_<timestamp>.txt` |

**Examples**

```powershell
# Run diagnostics, output to screen only
.\Test-OpenVpnDiagnostics.ps1

# Run and save report to C:\Temp\
.\Test-OpenVpnDiagnostics.ps1 -ExportTxt

# Save to a custom path
.\Test-OpenVpnDiagnostics.ps1 -OutputPath "C:\Support\vpn-report.txt"
```
