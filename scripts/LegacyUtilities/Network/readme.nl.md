[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [LegacyUtilities](../readme.nl.md) › **Network**

# Legacy Utilities — Network

---

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`Connect-AzureFileShareDrive.ps1`](Connect-AzureFileShareDrive.ps1) ([docs](#connect-azurefilesharedriveps1)) | Koppel een Azure Files SMB-share als blijvende stationsletter |

---

### Connect-AzureFileShareDrive.ps1

Test de SMB-verbinding (poort 445) met het opslagaccount, slaat de toegangssleutel op via
`cmdkey` en koppelt de share met `New-PSDrive`. Gegeneraliseerde vervanging van een script
waarin een specifieke naam en toegangssleutel van een opslagaccount hardcoded stonden; deze
versie krijgt ze als parameters mee en bewaart de sleutel nergens anders dan waar `cmdkey`
hem zelf opslaat. Standaard een proefdraai (alleen verbindingstest, geen koppeling).

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-StorageAccountName` | Ja | Naam van het Azure Storage-account |
| `-ShareName` | Ja | Naam van de bestandsshare |
| `-StorageAccountKey` | Ja | Toegangssleutel van het opslagaccount |
| `-DriveLetter` | Nee | Stationsletter (standaard: `S`) |
| `-Persist` | Nee | Laat de koppeling een herstart overleven |
| `-Apply` | Nee | Sla de referentie echt op en koppel de share (standaard: alleen voorbeeldweergave/test) |

```powershell
.\Connect-AzureFileShareDrive.ps1 -StorageAccountName "contosofiles" -ShareName "documents" -StorageAccountKey $key -DriveLetter S -Persist -Apply
```

**Opmerkingen**
- Vereist uitgaand TCP 445 naar `*.file.core.windows.net`, wat door sommige
  providers/firewalls wordt geblokkeerd. Gebruik een Azure P2S/S2S-VPN of ExpressRoute om
  SMB-verkeer over een andere poort te tunnelen als 445 niet beschikbaar is.
- Voor het koppelen van SharePoint/OneDrive-documentbibliotheken (geen kale Azure Files),
  zie in plaats daarvan [`scripts/Device/DriveMapping/`](../../Device/DriveMapping/readme.nl.md).
