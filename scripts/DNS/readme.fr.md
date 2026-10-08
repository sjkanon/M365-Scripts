[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../readme.fr.md) › [scripts](../readme.fr.md) › **DNS**

# Scripts DNS

Scripts de résolution et d'import d'enregistrements DNS dans des zones DNS intégrées à Active Directory.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Import-DnsRecords.ps1`](Import-DnsRecords.ps1) ([docs](#import-dnsrecordsps1)) | Résoudre les FQDN d'un CSV via Google DNS et importer éventuellement les enregistrements dans une zone DNS intégrée à AD (exemple d'entrée : [`example-records.csv`](example-records.csv)) |

---

### Import-DnsRecords.ps1

Lit une liste de FQDN dans un CSV, résout chacun d'eux via Google DNS (8.8.8.8) à l'aide de `dig`, et importe éventuellement les résultats dans une zone DNS Active Directory.

**Logique de résolution par FQDN**

1. Recherche d'un CNAME → s'il est trouvé, le type d'enregistrement est CNAME
2. Recherche d'un A → s'il est trouvé, le type d'enregistrement est A
3. Aucune réponse → signalé comme non résolvable, ignoré

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-CsvPath` | Oui | Chemin du fichier CSV contenant une colonne `FQDN` |
| `-ExportCsv` | Non | Enregistrer les enregistrements résolus dans un CSV pour vérification manuelle |
| `-ExportPath` | Non | Chemin personnalisé pour le CSV d'export (implique `-ExportCsv`). Par défaut : `C:\Temp\` / `~/Downloads\` |
| `-Apply` | Non | Écrire les enregistrements résolus dans le DNS AD (nécessite Windows + le module DnsServer) |
| `-ZoneName` | Uniquement avec `-Apply` | Zone DNS AD à laquelle ajouter les enregistrements (p. ex. `contoso.com`) |
| `-DnsServer` | Non | Serveur DNS dans lequel écrire (par défaut : `localhost`) |
| `-Ttl` | Non | TTL en secondes (par défaut : `3600`) |

**Format CSV**

```csv
FQDN
mail.contoso.com
webmail.contoso.com
portal.contoso.com
www.contoso.com
```

**Exemples**

```powershell
# Résoudre et afficher à l'écran — fonctionne sous macOS/Linux
.\Import-DnsRecords.ps1 -CsvPath .\records.csv

# Résoudre et exporter en CSV pour un import manuel
.\Import-DnsRecords.ps1 -CsvPath .\records.csv -ExportCsv

# Résoudre et importer directement dans le DNS AD
.\Import-DnsRecords.ps1 -CsvPath .\records.csv -ZoneName contoso.com -Apply

# Serveur DNS distant
.\Import-DnsRecords.ps1 -CsvPath .\records.csv -ZoneName contoso.com -DnsServer dc01.contoso.com -Apply
```

**Prérequis**

- `dig` — à installer via `choco install bind-toolsonly` ou [isc.org/download](https://www.isc.org/download/)
- Module `DnsServer` — nécessaire uniquement avec `-Apply` (RSAT ou rôle Windows DNS Server)

**Remarques**

- Ignore les enregistrements qui existent déjà — peut être relancé sans risque
- Les FQDN qui ne correspondent pas à `-ZoneName` sont ignorés avec un avertissement lors de l'utilisation de `-Apply`
- Le CNAME est détecté en premier ; l'enregistrement A sert de repli
