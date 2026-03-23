# DNS Scripts

Scripts for managing DNS records in Active Directory-integrated DNS zones.

---

## Scripts

### Add-DnsRecords.ps1

Resolves a list of FQDNs via Google DNS (8.8.8.8) and imports the results as A or CNAME records into an Active Directory DNS zone. Defaults to dry-run — pass `-Apply` to write records.

**How it works:**

1. Reads a CSV with a `FQDN` column
2. Resolves each FQDN via `dig @8.8.8.8` (CNAME checked first, then A)
3. Shows what was found and what would be created
4. With `-Apply`: writes the records to AD DNS (skips records that already exist)

**Parameters**

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-CsvPath` | Yes | Path to CSV file with a `FQDN` column |
| `-ZoneName` | Yes | AD DNS zone to add records to (e.g. `vias.be`) |
| `-DnsServer` | No | DNS server to write to (default: `localhost`) |
| `-Ttl` | No | TTL in seconds (default: `3600`) |
| `-Apply` | No | Actually create the records (default: dry run) |

**CSV format**

The CSV must have a single `FQDN` column with fully qualified domain names:

```csv
FQDN
briefings.vias.be
helpdesk.vias.be
meetweek.vias.be
mobisafetyscan.vias.be
semaindecomptage.vias.be
www.briefings.vias.be
www.meetweek.vias.be
www.mobisafetyscan.vias.be
www.semaindecomptage.vias.be
```

**Examples**

```powershell
# Dry run — resolve and show what would be added
.\Add-DnsRecords.ps1 -CsvPath .\records.csv -ZoneName vias.be

# Resolve and import into local AD DNS
.\Add-DnsRecords.ps1 -CsvPath .\records.csv -ZoneName vias.be -Apply

# Remote DNS server
.\Add-DnsRecords.ps1 -CsvPath .\records.csv -ZoneName vias.be -DnsServer dc01.vias.be -Apply
```

**Notes**

- Requires `dig` — install via `choco install bind-toolsonly` or [isc.org/download](https://www.isc.org/download/)
- Requires the `DnsServer` PowerShell module (RSAT or Windows DNS Server role)
- Skips records that already exist — safe to re-run
- FQDNs not matching the `-ZoneName` are skipped with a warning
- CNAME records are detected first; A records are used as fallback
