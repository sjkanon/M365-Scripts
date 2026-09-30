[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [LegacyUtilities](../readme.fr.md) › **Network**

# Legacy Utilities — Network

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Connect-AzureFileShareDrive.ps1`](Connect-AzureFileShareDrive.ps1) ([docs](#connect-azurefilesharedriveps1)) | Monter un partage SMB Azure Files sous une lettre de lecteur persistante |

---

### Connect-AzureFileShareDrive.ps1

Teste la connectivité SMB (port 445) vers le compte de stockage, enregistre sa clé d'accès via
`cmdkey` et mappe le partage avec `New-PSDrive`. Remplacement généralisé d'un script qui
codait en dur le nom et la clé d'accès d'un compte de stockage précis ; cette version les reçoit
en paramètres et ne conserve jamais la clé au-delà de ce que `cmdkey` stocke lui-même. Essai à
blanc par défaut (test de connectivité uniquement, sans montage).

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-StorageAccountName` | Oui | Nom du compte Azure Storage |
| `-ShareName` | Oui | Nom du partage de fichiers |
| `-StorageAccountKey` | Oui | Clé d'accès du compte de stockage |
| `-DriveLetter` | Non | Lettre de lecteur (par défaut : `S`) |
| `-Persist` | Non | Conserver le mappage après un redémarrage |
| `-Apply` | Non | Enregistrer réellement les identifiants et monter le partage (par défaut : aperçu/test uniquement) |

```powershell
.\Connect-AzureFileShareDrive.ps1 -StorageAccountName "contosofiles" -ShareName "documents" -StorageAccountKey $key -DriveLetter S -Persist -Apply
```

**Remarques**
- Nécessite le TCP 445 sortant vers `*.file.core.windows.net`, bloqué par certains
  FAI/pare-feu. Utilisez un VPN Azure P2S/S2S ou ExpressRoute pour faire transiter le trafic
  SMB par un autre port si le 445 n'est pas disponible.
- Pour le mappage de bibliothèques de documents SharePoint/OneDrive (et non d'Azure Files
  brut), consultez plutôt [`scripts/Device/DriveMapping/`](../../Device/DriveMapping/readme.fr.md).
