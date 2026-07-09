# DNS Scripts

Scripts for resolving and importing DNS records into Active Directory-integrated DNS zones.

---

## Scripts

### Import-DnsRecords.ps1

Reads a list of FQDNs from a CSV, resolves each one via Google DNS (8.8.8.8) using `dig`, and optionally imports the results into an Active Directory DNS zone.

**Resolution logic per FQDN**

1. Check for CNAME → if found, record type is CNAME
2. Check for A → if found, record type is A
3. No answer → reported as unresolvable, skipped

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-CsvPath` | Yes | Path to CSV file with a `FQDN` column |
| `-ExportCsv` | No | Save resolved records to CSV for manual review |
| `-ExportPath` | No | Custom path for the export CSV (implies `-ExportCsv`). Default: `C:\Temp\` / `~/Downloads\` |
| `-Apply` | No | Write resolved records into AD DNS (requires Windows + DnsServer module) |
| `-ZoneName` | Only with `-Apply` | AD DNS zone to add records to (e.g. `contoso.com`) |
| `-DnsServer` | No | DNS server to write to (default: `localhost`) |
| `-Ttl` | No | TTL in seconds (default: `3600`) |

**CSV format**

```csv
FQDN
mail.contoso.com
webmail.contoso.com
portal.contoso.com
www.contoso.com
```

**Examples**

```powershell
# Resolve and show on screen — works on macOS/Linux
.\Import-DnsRecords.ps1 -CsvPath .\records.csv

# Resolve and export to CSV for manual import
.\Import-DnsRecords.ps1 -CsvPath .\records.csv -ExportCsv

# Resolve and import directly into AD DNS
.\Import-DnsRecords.ps1 -CsvPath .\records.csv -ZoneName contoso.com -Apply

# Remote DNS server
.\Import-DnsRecords.ps1 -CsvPath .\records.csv -ZoneName contoso.com -DnsServer dc01.contoso.com -Apply
```

**Requirements**

- `dig` — install via `choco install bind-toolsonly` or [isc.org/download](https://www.isc.org/download/)
- `DnsServer` module — only required with `-Apply` (RSAT or Windows DNS Server role)

**Notes**

- Skips records that already exist — safe to re-run
- FQDNs not matching `-ZoneName` are skipped with a warning when using `-Apply`
- CNAME is detected first; A record is used as fallback
