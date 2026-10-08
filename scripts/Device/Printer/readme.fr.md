[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [Device](../readme.fr.md) › **Printer**

# Imprimantes

Installe des pilotes d'imprimante et des imprimantes TCP/IP à partir d'un seul fichier JSON. Les pilotes sont téléchargés depuis un dépôt GitHub (asset de release ou dossier), une URL https quelconque ou un partage.

Conçu pour s'exécuter sans surveillance après qu'une image de référence (golden image) a provisionné un nouveau serveur ou hôte de session. Au premier démarrage, il attend le spouleur d'impression et le réseau au lieu d'échouer, et comme chaque exécution est idempotente, la même commande peut aussi s'exécuter à chaque démarrage.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Install-Printer.ps1`](Install-Printer.ps1) ([docs](#install-printerps1)) | Installer des pilotes d'imprimante (téléchargés depuis GitHub) et des imprimantes décrits dans un fichier JSON |

## Fichiers

| Fichier | Description |
|---------|-------------|
| [`printers.example.json`](printers.example.json) | Exemple de configuration avec chaque champ et chaque source de pilote |

---

### Install-Printer.ps1

Par exécution :

1. **Config** — lire le JSON (local, UNC ou URL https) et retenir les imprimantes demandées avec `-Printer`. Tout le fichier est validé avant le moindre téléchargement
2. **Contrôle préalable** — démarrer le spouleur d'impression si nécessaire et l'attendre, puis comparer chaque pilote (installé ? quelle version ?) et chaque imprimante (port, pilote, paramètres) au JSON
3. **Téléchargement** — uniquement pour un pilote absent ou plus ancien que la `version` du JSON. Un `sha256` facultatif est vérifié, et les fichiers `.zip`/`.cab` sont extraits
4. **Pilote** — trouver l'INF (nommé dans le JSON, ou celui qui déclare le nom du pilote), vérifier la signature du catalogue, puis exécuter `pnputil /add-driver /install` + `Add-PrinterDriver`
5. **Imprimante** — créer le port TCP/IP, ajouter l'imprimante ou corriger son pilote/port, et définir l'emplacement, le commentaire, le partage et les paramètres d'impression par défaut (recto verso, couleur, format de papier)
6. **Vérification** — tout relire

Un appareil déjà conforme ne coûte qu'une lecture du JSON. Rien n'est téléchargé et rien ne change.

**Le JSON**

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

| Champ pilote | Description |
|--------------|-------------|
| `name` | Nom exact du pilote dans l'INF (tel que `Get-PrinterDriver` l'affiche). Les imprimantes y font référence |
| `version` | Facultatif : un pilote installé plus ancien est mis à jour. Sans ce champ, un pilote installé n'est pas touché (sauf avec `-Force`) |
| `inf` | Facultatif : nom ou chemin de l'INF dans le paquet (caractères génériques autorisés). Sans ce champ, le script prend l'INF qui mentionne `name`, en préférant le dossier x64/arm64 |
| `install` | Facultatif `true` : installer même si aucune des imprimantes retenues ne l'utilise |
| `source` | D'où viennent les fichiers — voir ci-dessous |

| `type` de source | Champs |
|------------------|--------|
| `githubRelease` | `repository` (`propriétaire/nom`), `tag` (par défaut `latest`), `asset` (nom, caractères génériques autorisés — doit correspondre à un seul asset) |
| `github` | `repository`, `path` (un `.zip` ou un dossier contenant l'INF), `ref` (branche/tag/commit, par défaut : la branche par défaut) |
| `url` | `url` (https) — un `.zip`, un `.cab` ou un fichier seul |
| `path` | `path` vers un dossier ou un `.zip` ; les chemins relatifs partent du dossier du JSON |
| *(toutes)* | `sha256` — facultatif, vérifié avant l'extraction (fichiers seuls uniquement) |

| Champ imprimante | Description |
|------------------|-------------|
| `name`, `driver`, `address` | Obligatoires (`address` = IP ou nom d'hôte) |
| `portName` | Par défaut `IP_<address>` |
| `portNumber` | Port RAW, par défaut `9100` |
| `lprQueue` | Utiliser LPR avec cette file au lieu de RAW |
| `snmp` | `true` active l'état SNMP. Désactivé par défaut, pour qu'une imprimante qui ne répond pas en SNMP ne s'affiche pas Hors connexion |
| `location`, `comment` | Visibles par les utilisateurs |
| `shared`, `shareName` | Partager l'imprimante (nom de partage par défaut : `name`) |
| `duplex` | `OneSided`, `TwoSidedLongEdge` ou `TwoSidedShortEdge` |
| `color` | `true` / `false` |
| `paperSize` | p. ex. `A4`, `Letter` |
| `ensure` | `absent` supprime l'imprimante (et son port si aucune autre imprimante ne l'utilise). Le pilote reste toujours |

**Paramètres**

| Paramètre | Description |
|-----------|-------------|
| `-ConfigPath` | Le JSON : chemin local/UNC ou URL https (p. ex. un lien raw dans le même dépôt GitHub) |
| `-Printer` | Uniquement ces imprimantes du JSON (noms, caractères génériques autorisés). Par défaut : toutes |
| `-GitHubToken` | Jeton pour un dépôt privé (à défaut, `$env:GITHUB_TOKEN`). Envoyé uniquement aux hôtes de GitHub, jamais à une source `url` |
| `-WorkingDir` | Dossier de téléchargement/extraction (par défaut : `C:\IT\Printers`) |
| `-LogPath` | Dossier de la transcription d'une exécution qui modifie quelque chose (par défaut : `C:\Temp`) |
| `-WaitSeconds` | Délai accordé à une machine fraîchement provisionnée pour le spouleur et le réseau (par défaut : `300`, `0` = pas d'attente) |
| `-CheckOnly` | Rapport uniquement, aucune modification. Le code de sortie `2` signifie qu'il y a du travail |
| `-Quiet` | N'afficher rien sauf s'il y a du travail ou un échec |
| `-Force` | Réinstaller les pilotes même si la version correspond |
| `-SkipSignatureCheck` | Ignorer la vérification de signature du catalogue faite par le script (Windows refuse toujours les pilotes de paquet non signés) |

**Exemples**

```powershell
# Ce qui serait fait - rien n'est modifié
.\Install-Printer.ps1 -ConfigPath .\printers.json -CheckOnly

# Uniquement les imprimantes Office, en simulation, JSON directement depuis GitHub
.\Install-Printer.ps1 -ConfigPath https://raw.githubusercontent.com/contoso/printer-drivers/main/printers.json -Printer 'Office*' -WhatIf

# Premier démarrage d'un serveur issu d'une golden image (Custom Script Extension / tâche de démarrage en SYSTEM)
powershell.exe -ExecutionPolicy Bypass -File C:\IT\Install-Printer.ps1 -ConfigPath https://raw.githubusercontent.com/contoso/printer-drivers/main/printers.json -Quiet -Confirm:$false

# Dépôt privé
$env:GITHUB_TOKEN = '<fine-grained token, Contents: read>'
.\Install-Printer.ps1 -ConfigPath \\fs01\it$\printers.json -Confirm:$false
```

**Codes de sortie**

| Code | Signification |
|------|---------|
| `0` | Tout est conforme, ou l'installation a réussi |
| `1` | Échec (un pilote ou une imprimante qui n'a pas pu être installé, ou un JSON invalide) |
| `2` | Uniquement avec `-CheckOnly` : il y a du travail |

**Remarques**
- **Après une golden image :** exécutez-le en SYSTEM depuis l'extension Azure Custom Script, une tâche planifiée au démarrage intégrée à l'image, `SetupComplete.cmd`, Intune ou un RMM. Si le spouleur d'impression n'a pas encore démarré, le script le démarre et l'attend. Les téléchargements qui échouent sur le DNS ou un délai d'attente sont retentés avec un délai croissant, dans la limite de `-WaitSeconds`. Un 401/403/404 échoue aussitôt, car attendre n'y change rien. Un spouleur *désactivé* par l'image (durcissement PrintNightmare) est signalé, pas réactivé en silence
- Les imprimantes sont créées pour toute la machine, donc sur un hôte de session RDS/AVD chaque utilisateur les voit. Aucune étape Point and Print par utilisateur n'intervient, donc la restriction `RestrictDriverInstallationToAdministrators` (KB5005652) ne s'applique pas — le pilote est installé par un administrateur/SYSTEM
- **GitHub :** un asset de release est téléchargé via l'URL d'asset de l'API, si bien que le même code fonctionne avec et sans jeton. Un dossier est listé une fois via l'API git tree, puis chaque fichier qu'il contient est récupéré. Sans authentification, GitHub autorise 60 requêtes API par heure et par IP, et un dossier coûte une requête par fichier ; pour de nombreux hôtes derrière un même NAT, utilisez un `.zip` de release ou un jeton. Un dépôt privé répond 404 sans jeton, et le message d'erreur le dit
- Seuls les paquets de pilotes à base d'INF sont pris en charge. Un programme d'installation `.exe` du fabricant ne l'est pas — extrayez le paquet (la plupart des fabricants proposent un zip « driver only ») et placez-le dans le dépôt. Le catalogue (`.cat`) doit porter une signature valide
- Les codes de sortie `pnputil` `0`, `259` (aucun périphérique en attente — normal pour les imprimantes), `3010` et `1641` (redémarrage) comptent comme une réussite. Pour tout autre code, le message renvoie à `C:\Windows\INF\setupapi.dev.log`
- Il n'existe pas de `Set-PrinterPort` : un port qui existe déjà pour une autre adresse est signalé, pas recréé, car d'autres imprimantes peuvent l'utiliser. Donnez à l'imprimante son propre `portName`
- `Set-PrintConfiguration` exécute le code propre au pilote du fabricant et se bloque avec certains pilotes universels ; il s'exécute donc dans un job avec un délai de 2 minutes. Un échec à cet endroit est un avertissement — l'imprimante est installée malgré tout
- L'imprimante par défaut est un réglage par utilisateur et n'est volontairement pas gérée : en SYSTEM, il n'y a pas d'utilisateur pour qui la définir
- Lancé en 32 bits (Intune Management Extension, certains agents RMM), le script se relance en 64 bits, car `pnputil` n'existe pas sous SysWOW64. Lancé à la main sans élévation, il la demande
- Variables de script NinjaOne : `configPath`, `printer`, `githubToken`, `workingDir`, `logPath`, `waitSeconds`, et les cases à cocher `checkOnly`, `whatIf`, `quiet`, `force`, `skipSignatureCheck`
- `-CheckOnly` sert de contrôle de santé pour une condition RMM ou une tâche de détection planifiée : le code de sortie `2` signifie que l'appareil ne correspond pas au JSON
