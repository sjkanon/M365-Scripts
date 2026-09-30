[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [Network](../readme.fr.md) › **UniFi**

# UniFi

Outillage pour un UniFi Network Controller ou une console UniFi OS (UDM/UDM-Pro/UDR). Communique directement avec l'API du contrôleur — ne fait pas partie de [`menu.ps1`](../../../menu.ps1), car chaque exécution nécessite une URL de contrôleur et des identifiants. Les identifiants sont toujours demandés via `-Credential`/`Get-Credential`, jamais codés en dur.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Get-UnifiNetworkReport.ps1`](Get-UnifiNetworkReport.ps1) ([docs](#get-unifinetworkreportps1)) | Générer un rapport HTML de documentation réseau (équipements, firmware, uptime) |
| [`Update-UnifiFirmware.ps1`](Update-UnifiFirmware.ps1) ([docs](#update-unififirmwareps1)) | Lister et éventuellement déclencher les mises à niveau du firmware sur l'ensemble des sites |
| [`UnifiApi.ps1`](UnifiApi.ps1) | Utilitaire partagé de connexion/session, chargé automatiquement (dot-sourced) par les deux scripts ci-dessus — n'est pas destiné à être exécuté directement |

---

### Get-UnifiNetworkReport.ps1

Lecture seule. Se connecte, énumère chaque site (ou un seul avec `-Site`), liste les équipements adoptés par site et écrit un rapport HTML (nom, modèle, MAC, IP, version du firmware, état, uptime) dans `C:\Temp\`.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-Controller` | Oui | URL de base du contrôleur, p. ex. `https://192.168.1.1` ou `https://unifi.contoso.local:8443` |
| `-Credential` | Non | Identifiants administrateur — demandés via `Get-Credential` s'ils sont omis |
| `-Site` | Non | Limiter à un seul site par son nom. S'il est omis, tous les sites sont inclus |
| `-SkipCertificateCheck` | Non | Accepter les certificats auto-signés/non approuvés (courant pour les contrôleurs on-premise) |
| `-OutputPath` | Non | Dossier du rapport (par défaut : `C:\Temp\` / `~/Downloads`) |

**Exemples**

```powershell
.\Get-UnifiNetworkReport.ps1 -Controller "https://192.168.1.1" -SkipCertificateCheck
.\Get-UnifiNetworkReport.ps1 -Controller "https://unifi.contoso.local:8443" -Site "Head Office"
```

---

### Update-UnifiFirmware.ps1

Liste par site les équipements pour lesquels une mise à niveau du firmware est disponible (d'après l'indicateur `upgradable` du contrôleur lui-même). S'exécute par défaut en essai à blanc — aucune mise à niveau n'est déclenchée sans `-Apply`. La mise à niveau redémarre l'équipement, ce qui provoque une brève coupure pour tout ce qui y est connecté — une confirmation par équipement est demandée, sauf si `-Force` est également passé.

**Paramètres**

| Paramètre | Obligatoire | Description |
|-----------|----------|-------------|
| `-Controller` | Oui | URL de base du contrôleur |
| `-Credential` | Non | Identifiants administrateur — demandés via `Get-Credential` s'ils sont omis |
| `-Site` | Non | Limiter à un seul site par son nom. S'il est omis, tous les sites sont traités |
| `-SkipCertificateCheck` | Non | Accepter les certificats auto-signés/non approuvés |
| `-Apply` | Non | Déclencher réellement les mises à niveau (par défaut : essai à blanc) |
| `-Force` | Non | Ignorer la confirmation par équipement lorsqu'utilisé avec `-Apply` |
| `-OutputPath` | Non | Dossier du rapport (par défaut : `C:\Temp\` / `~/Downloads`) |

**Exemples**

```powershell
# Essai à blanc sur tous les sites
.\Update-UnifiFirmware.ps1 -Controller "https://192.168.1.1" -SkipCertificateCheck

# Mettre à niveau un site, avec confirmation par équipement
.\Update-UnifiFirmware.ps1 -Controller "https://192.168.1.1" -Site "Head Office" -Apply

# Sans surveillance, tous les sites — à utiliser avec prudence, les équipements redémarrent
.\Update-UnifiFirmware.ps1 -Controller "https://192.168.1.1" -Apply -Force
```

**Remarques**
- Les deux scripts prennent en charge le UniFi Network Controller classique auto-hébergé (`/api/login`) et les consoles UniFi OS (`/api/auth/login` + `/proxy/network/...`) — la connexion détecte automatiquement lequel des deux se trouve en face.
- `-SkipCertificateCheck` est implémenté à la fois pour PowerShell 7+ (paramètre natif) et Windows PowerShell 5.1 (callback temporaire de validation des certificats, réinitialisé immédiatement après la requête).
- Un rapport CSV/HTML est toujours écrit après chaque exécution, même en essai à blanc.
