[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [Device](../readme.fr.md) › **Printer**

# Imprimantes

Installe des pilotes d'imprimante et des imprimantes TCP/IP à partir d'un seul fichier JSON. Les pilotes sont téléchargés depuis un dépôt GitHub (asset de release ou dossier), une URL https quelconque ou un partage.

[`Export-Printer.ps1`](#export-printerps1) fait l'inverse : il lit les imprimantes d'une machine existante et écrit ce même JSON, avec les pilotes à côté, dans un seul zip.

Conçu pour s'exécuter sans surveillance après qu'une image de référence (golden image) a provisionné un nouveau serveur ou hôte de session. Au premier démarrage, il attend le spouleur d'impression et le réseau au lieu d'échouer, et comme chaque exécution est idempotente, la même commande peut aussi s'exécuter à chaque démarrage.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Export-Printer.ps1`](Export-Printer.ps1) ([docs](#export-printerps1)) | Exporter les imprimantes d'une machine - paramètres en fichier JSON, pilotes, tickets d'impression complets - dans un zip dans `C:\Temp`, prêt pour `Install-Printer.ps1` sur une autre |
| [`Install-Printer.ps1`](Install-Printer.ps1) ([docs](#install-printerps1)) | Installer des pilotes d'imprimante (téléchargés depuis GitHub) et des imprimantes décrits dans un fichier JSON |
| [`New-PrinterGroup.ps1`](New-PrinterGroup.ps1) ([docs](#new-printergroupps1)) | Un groupe de sécurité Entra ID par imprimante (`SG-U-PRINTM-<imprimante>`) à partir d'une matrice imprimantes x utilisateurs, membres ajoutés, et les SID des groupes écrits comme autorisations dans `printers.json` |

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
| `inputBin` | Bac par défaut, par le nom qu'affiche le pilote (`Tray 2`, `Manual Feed`) ou son nom Print Schema (`ns0000:Tray2`) ; espaces et casse sans importance. Un nom inconnu du pilote donne un avertissement qui liste les bacs qu'il propose |
| `permissions` | Qui peut utiliser cette file : une liste de `{ "principal": "CONTOSO\\Finance", "access": "print" }` — voir *Groupes de sécurité par file* ci-dessous. Sans ce champ, les autorisations ne sont pas touchées |
| `ensure` | `absent` supprime l'imprimante (et son port si aucune autre imprimante ne l'utilise). Le pilote reste toujours |

**Un appareil, plusieurs files (bacs)**

Les utilisateurs choisissent une file, pas un bac. Donnez à chaque file sa propre entrée avec **la même adresse** : elles partagent le port `IP_<address>`, et chacune a son propre bac par défaut, recto verso ou couleur :

```json
"printers": [
  { "name": "Office",            "driver": "HP Universal Printing PCL 6", "address": "10.0.5.20", "inputBin": "Tray 1" },
  { "name": "Office - letterhead", "driver": "HP Universal Printing PCL 6", "address": "10.0.5.20", "inputBin": "Tray 2", "duplex": "OneSided" },
  { "name": "Office - envelopes", "driver": "HP Universal Printing PCL 6", "address": "10.0.5.20", "inputBin": "Manual Feed" }
]
```

Le bac est défini dans le ticket d'impression par défaut de la file : c'est le point de départ de chaque utilisateur de cette file, qui peut toujours choisir un autre bac pour une impression. Les noms de bacs varient selon le pilote - exécutez le script une fois avec le nom attendu, et un nom erroné reçoit en réponse la liste que propose le pilote. Les pilotes universels n'affichent les bacs réels de l'appareil qu'une fois ses options connues ; si seul `Auto Select` revient, configurez d'abord les options installables dans les propriétés de l'imprimante, ou utilisez le pilote spécifique au modèle du fabricant.

**Groupes de sécurité par file**

Une file peut être réservée à un ou plusieurs groupes, par exemple un bac à chèques sur lequel seule la Finance peut imprimer :

```json
{ "name": "Office - cheques", "driver": "HP Universal Printing PCL 6", "address": "10.0.5.20", "inputBin": "Tray 3",
  "permissions": [
    { "principal": "CONTOSO\\Finance",          "access": "print" },
    { "principal": "CONTOSO\\Office Managers",  "access": "manageDocuments" },
    { "principal": "S-1-12-1-1111-2222-3333-4444", "access": "print" }
  ] }
```

| `access` | Signification |
|----------|---------------|
| `print` | Imprimer et gérer ses propres travaux — ce qu'a Everyone par défaut |
| `manageDocuments` | Aussi suspendre, relancer et supprimer les travaux des autres |
| `manage` | Aussi modifier l'imprimante elle-même (paramètres, partage, autorisations) |

- `principal` est `DOMAINE\Groupe`, un utilisateur ou groupe local, ou un **SID**. Un groupe Entra ID ne peut être indiqué que par son SID (`S-1-12-1-…`, depuis Entra ID ou `whoami /groups` d'un membre). `Everyone`, `Authenticated Users`, `Users`, `Interactive`, `Administrators` et `SYSTEM` sont compris en anglais, sur un Windows de n'importe quelle langue
- Les principaux listés reçoivent cet accès ; **Everyone et tout autre utilisateur ou groupe perdent le leur**, y compris le compte qui a installé l'imprimante. Administrators, SYSTEM, CREATOR OWNER, les packages d'application et les identités de service gardent toujours leur accès — l'impression depuis des applications empaquetées comme le nouvel Outlook dépend de ces entrées, elles ne sont donc jamais supprimées. `[]` signifie administrateurs uniquement
- **Échec sécurisé :** un groupe introuvable empêche la création de cette file (`-CheckOnly` le signale à l'avance), et une nouvelle file dont les autorisations ne peuvent être définies est supprimée plutôt que laissée ouverte à tous. Une file existante garde ses anciennes autorisations et l'exécution échoue. Au premier démarrage, un nom introuvable est retenté jusqu'à 2 minutes, le temps qu'un contrôleur de domaine soit joignable
- Comparaison par le sens, pas par le texte : Windows peut enregistrer une entrée autrement qu'elle a été donnée, et cela ne compte pas comme une modification. Les autorisations sont vérifiées à chaque exécution, donc un groupe ajouté à la main dans les propriétés de l'imprimante est retiré — le JSON fait foi
- Les autorisations décident qui peut imprimer sur la file. Qu'un utilisateur sans accès la *voie* encore dans la liste dépend de l'application et n'est pas contrôlé par ce script

**Paramètres**

| Paramètre | Description |
|-----------|-------------|
| `-ConfigPath` | Le JSON : chemin local/UNC ou URL https (p. ex. un lien raw dans le même dépôt GitHub) |
| `-Printer` | Uniquement ces imprimantes du JSON (noms, caractères génériques autorisés). Par défaut : toutes |
| `-GitHubToken` | Jeton pour un dépôt privé (à défaut, `$env:GITHUB_TOKEN`). Envoyé uniquement aux hôtes de GitHub, jamais à une source `url` |
| `-Proxy` | Proxy pour chaque téléchargement, p. ex. `http://proxy.contoso.local:8080`. En SYSTEM, il n'y a pas de proxy utilisateur à reprendre ; utilise les informations d'identification du compte ordinateur |
| `-WorkingDir` | Dossier de téléchargement/extraction (par défaut : `C:\IT\Printers`) |
| `-LogPath` | Dossier de `Install-Printer.log` (par défaut : `C:\Temp`), complété par chaque exécution susceptible de modifier quelque chose et renouvelé au-delà de 1 Mo |
| `-WaitSeconds` | Délai accordé à une machine fraîchement provisionnée pour le spouleur, le réseau et la fin d'une autre exécution de ce script (par défaut : `300`, `0` = pas d'attente) |
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

# Premier démarrage ou déploiement (Custom Script Extension, Run Command, tâche de démarrage en SYSTEM)
powershell.exe -ExecutionPolicy Bypass -File C:\IT\Install-Printer.ps1 -ConfigPath https://raw.githubusercontent.com/contoso/printer-drivers/main/printers.json -Quiet

# Dépôt privé
$env:GITHUB_TOKEN = '<fine-grained token, Contents: read>'
.\Install-Printer.ps1 -ConfigPath \\fs01\it$\printers.json -Confirm:$false
```

**Installer au déploiement plutôt que dans l'image**

L'image reste sans pilotes ni imprimantes ; chaque nouvel hôte de session les reçoit lors de son déploiement, de sorte qu'une imprimante modifiée est une modification du JSON et non une nouvelle image. Exécutez le script une fois en SYSTEM via l'extension Azure Custom Script dans le Bicep de la VM (chaque hôte ajouté au pool l'exécute automatiquement), ou via Run Command sur un hôte qui existe déjà :

```bicep
// Custom Script Extension: runs once, as SYSTEM, when the VM is deployed
resource installPrinters 'Microsoft.Compute/virtualMachines/extensions@2024-07-01' = {
  parent: vm
  name: 'InstallPrinters'
  location: location
  properties: {
    publisher: 'Microsoft.Compute'
    type: 'CustomScriptExtension'
    typeHandlerVersion: '1.10'
    autoUpgradeMinorVersion: true
    settings: {
      fileUris: [
        'https://raw.githubusercontent.com/sjkanon/M365-Scripts/<commit>/scripts/Device/Printer/Install-Printer.ps1'
      ]
    }
    protectedSettings: {
      commandToExecute: 'powershell.exe -ExecutionPolicy Bypass -File Install-Printer.ps1 -ConfigPath https://raw.githubusercontent.com/contoso/printer-drivers/main/printers.json -Quiet'
    }
  }
}
```

```powershell
# Run Command: afterwards, on a VM that is already running
az vm run-command invoke -g <resource-group> -n <vm> --command-id RunPowerShellScript `
  --scripts '@Install-Printer.ps1' `
  --parameters 'ConfigPath=https://raw.githubusercontent.com/contoso/printer-drivers/main/printers.json'
```

- Faites pointer `fileUris` vers un **commit** plutôt que `main`, pour qu'un hôte déployé le mois prochain exécute le script que vous avez testé. Le JSON peut rester sur une branche : ce sont des données, entièrement validées avant tout changement.
- **Pas de `-Confirm:$false`** sur ces lignes de commande : une exécution sans console ne demande jamais rien, et via `-File` le commutateur arriverait comme le texte `'$false'` et arrêterait le script avant son démarrage.
- Une VM n'a qu'**une** extension Custom Script. Si le déploiement l'utilise déjà pour autre chose, utilisez plutôt un Run Command managé (`Microsoft.Compute/virtualMachines/runCommands`).
- Un jeton GitHub pour un dépôt de pilotes privé va dans `protectedSettings` (`-GitHubToken`), jamais dans `settings`, lisible sur la VM.
- Le code de sortie `1` fait échouer l'extension, si bien qu'un déploiement avec une imprimante non installée est visible dans le portail ; `Install-Printer.log` dans `-LogPath` indique laquelle.

**Codes de sortie**

| Code | Signification |
|------|---------|
| `0` | Tout est conforme, ou l'installation a réussi |
| `1` | Échec (un pilote ou une imprimante qui n'a pas pu être installé, ou un JSON invalide) |
| `2` | Uniquement avec `-CheckOnly` : il y a du travail |

**Ce qui évite une exécution en échec**
- **Tout le JSON est vérifié avant toute action** : champs obligatoires, champs inconnus (une faute de frappe comme `adress` ou `loaction` est une erreur, pas ignorée en silence), noms en double, caractères refusés par Windows dans un nom d'imprimante, adresses IP/noms d'hôte, numéros de port, valeurs de `duplex`/`ensure`, `true`/`false` écrits comme du texte, et deux imprimantes sur un même port avec des adresses différentes. Tous les problèmes sont listés d'un coup, et rien n'est modifié
- **L'INF est vérifié avant que pnputil le voie** : classe imprimante, une section pour cette architecture (un paquet uniquement x86 sur un serveur x64 est signalé comme tel), le nom exact du pilote parmi les modèles déclarés — avec, dans l'erreur, les noms qu'il déclare *réellement*, les plus proches d'abord — et un catalogue signé pour cette architecture
- **Les téléchargements sont vérifiés** : une page de blocage de proxy, un portail de connexion ou une erreur GitHub enregistrés en `.zip` sont refusés, de même qu'un fichier qui n'est ni un zip ni un cab ; 1 Go libre est exigé avant tout téléchargement
- **Une seule exécution à la fois** : une tâche de démarrage et un job RMM qui se chevauchent s'attendent via un verrou à l'échelle de la machine (jusqu'à `-WaitSeconds`) au lieu d'installer deux fois le même pilote. Une exécution interrompue à mi-chemin est détectée et la suivante reprend à partir de ce qu'elle a laissé
- **Le spouleur** : au premier démarrage, il est démarré et attendu jusqu'à ce qu'il réponde vraiment, pas seulement jusqu'à ce qu'il indique Running. S'il s'arrête ou se bloque pendant une installation — fréquent juste après l'arrivée d'un pilote de fabricant —, il est redémarré et cette étape est retentée une fois. Il n'est volontairement *pas* redémarré après chaque pilote, comme le font certains scripts publiés : sur un hôte de session en service, cela interrompt l'impression de tout le monde
- **Le JSON d'une URL est mis en cache** dans `-WorkingDir`. Si l'URL est injoignable, la dernière copie valide est utilisée avec un avertissement, pour qu'un hôte qui démarre pendant une panne de GitHub garde ses imprimantes
- **Le code du fabricant est isolé** : la lecture et l'écriture des paramètres d'impression par défaut s'exécutent dans un job avec délai, car certains pilotes universels s'y bloquent
- **Un pilote en échec n'arrête pas les autres** : ses imprimantes sont ignorées et signalées, toutes les autres sont installées, et le code de sortie est `1`. Les fichiers téléchargés du pilote en échec sont conservés pour analyse ; après une installation réussie, ils sont supprimés (le magasin de pilotes garde sa propre copie)
- **Une adresse modifiée** est appliquée : avec les noms de port par défaut, l'imprimante passe au nouveau port `IP_<address>` et l'ancien est supprimé dès qu'il n'est plus utilisé ; un port nommé utilisé uniquement par cette imprimante est reconstruit sur place. Un port partagé avec d'autres imprimantes n'est pas modifié à leur insu
- **Détection** : une exécution sans erreur écrit `ConfigSha256`, `ConfigPath`, `LastSuccess` et `Printers` sous `HKLM:\SOFTWARE\M365-Scripts\InstallPrinter`. Une règle de détection Intune ou une condition RMM peut comparer ce hachage à celui du JSON, et distinguer « installé avec la configuration actuelle » de « installé avec celle du mois dernier »
- **Tout est journalisé** dans `Install-Printer.log` dans `-LogPath`, même quand l'exécution échoue tôt, pour pouvoir relire après coup un premier démarrage sans surveillance

**Remarques**
- **Après une golden image :** exécutez-le en SYSTEM depuis l'extension Azure Custom Script, une tâche planifiée au démarrage intégrée à l'image, `SetupComplete.cmd`, Intune ou un RMM. Si le spouleur d'impression n'a pas encore démarré, le script le démarre et l'attend. Les téléchargements qui échouent sur le DNS ou un délai d'attente sont retentés avec un délai croissant, dans la limite de `-WaitSeconds`. Un 401/403/404 échoue aussitôt, car attendre n'y change rien. Un spouleur *désactivé* par l'image (durcissement PrintNightmare) est signalé, pas réactivé en silence
- Les imprimantes sont créées pour toute la machine, donc sur un hôte de session RDS/AVD chaque utilisateur les voit. Aucune étape Point and Print par utilisateur n'intervient, donc la restriction `RestrictDriverInstallationToAdministrators` (KB5005652) ne s'applique pas — le pilote est installé par un administrateur/SYSTEM
- **GitHub :** les métadonnées de la release et la liste du dossier passent par l'API (une ou deux requêtes par pilote). Sans jeton, les fichiers eux-mêmes viennent du lien de téléchargement de la release et de `raw.githubusercontent.com` — aucun des deux ne compte dans les 60 requêtes API par heure que GitHub autorise par IP publique, donc un pool d'hôtes démarrant derrière un même NAT n'épuise pas la limite. Avec un jeton, tout passe par l'API, ce qu'exige un dépôt privé. Un dépôt privé répond 404 sans jeton, et le message d'erreur le dit. Un dépôt trop grand pour une seule liste d'arborescence, ou un fichier stocké via Git LFS, est refusé plutôt que téléchargé à moitié — publiez alors le pilote comme asset de release
- Seuls les paquets de pilotes à base d'INF sont pris en charge. Un programme d'installation `.exe` du fabricant ne l'est pas — extrayez le paquet (la plupart des fabricants proposent un zip « driver only ») et placez-le dans le dépôt. Le catalogue (`.cat`) doit porter une signature valide
- Les codes de sortie `pnputil` `0`, `259` (aucun périphérique en attente — normal pour les imprimantes), `3010` et `1641` (redémarrage) comptent comme une réussite. Pour tout autre code, le message renvoie à `C:\Windows\INF\setupapi.dev.log`
- Il n'existe pas de `Set-PrinterPort` : un port qui existe déjà pour une autre adresse est signalé, pas recréé, car d'autres imprimantes peuvent l'utiliser. Donnez à l'imprimante son propre `portName`
- `Set-PrintConfiguration` exécute le code propre au pilote du fabricant et se bloque avec certains pilotes universels ; il s'exécute donc dans un job avec un délai de 2 minutes. Un échec à cet endroit est un avertissement — l'imprimante est installée malgré tout
- L'imprimante par défaut est un réglage par utilisateur et n'est volontairement pas gérée : en SYSTEM, il n'y a pas d'utilisateur pour qui la définir
- Lancé en 32 bits (Intune Management Extension, certains agents RMM), le script se relance en 64 bits, car `pnputil` n'existe pas sous SysWOW64. Lancé à la main sans élévation, il la demande
- Variables de script NinjaOne : `configPath`, `printer`, `githubToken`, `workingDir`, `logPath`, `waitSeconds`, `proxy`, et les cases à cocher `checkOnly`, `whatIf`, `quiet`, `force`, `skipSignatureCheck`
- `-CheckOnly` sert de contrôle de santé pour une condition RMM ou une tâche de détection planifiée : le code de sortie `2` signifie que l'appareil ne correspond pas au JSON

---

### Export-Printer.ps1

L'inverse de `Install-Printer.ps1` : lit les imprimantes de cette machine et écrit un zip, par défaut `C:\Temp\PrinterExport_<ordinateur>_<yyyyMMdd-HHmmss>.zip`. Rien n'est modifié sur la machine.

| Dans le zip | Contenu |
|-------------|---------|
| `printers.json` | Les imprimantes et leurs paramètres au format d'Install-Printer : pilote, adresse, nom/numéro de port, file LPR, SNMP, emplacement, commentaire, partage, recto verso, couleur, format de papier, bac par défaut (`inputBin`) et autorisations. Chaque pilote pointe vers son propre dossier dans le zip (`"type": "path"`) |
| `drivers\<nom du pilote>\` | Le package de pilote complet tel que le magasin de pilotes le conserve : INF, catalogue signé et binaires - ce à partir de quoi `pnputil` installe. Les pilotes fournis avec Windows (`ntprint.inf`, `prnms*.inf`) ne sont pas copiés |
| `settings\<imprimante>.xml` | Le ticket d'impression par défaut complet de chaque imprimante, y compris chaque paramètre du fabricant que les champs JSON ne couvrent pas |
| `inventory.json` | Tout ce qui a été lu, aussi pour les imprimantes que le JSON ne peut pas décrire (WSD, USB, PDF, OneNote, fax, connexions à un serveur d'impression), avec la raison de leur absence de `printers.json` et le `PermissionSDDL` brut |
| `<ordinateur>.printerExport` | Uniquement avec `-PrintBrm` : la sauvegarde Printer Migration propre à Windows |

Sur la machine cible : décompresser le zip et lancer `.\Install-Printer.ps1 -ConfigPath <dossier>\printers.json -CheckOnly`, puis sans `-CheckOnly`. Les chemins relatifs des pilotes sont résolus par rapport au dossier du JSON.

**Paramètres**

| Paramètre | Description |
|-----------|-------------|
| `-Printer` | Uniquement ces imprimantes (noms, caractères génériques autorisés). Par défaut : toutes |
| `-OutputPath` | Dossier du zip (par défaut : `C:\Temp`), créé s'il manque |
| `-NoDriver` | Paramètres seulement, sans fichiers de pilote - `printers.json` ne liste alors aucun pilote, la cible doit donc déjà les avoir |
| `-PrintBrm` | Faire aussi une sauvegarde `PrintBrm.exe -b` et la mettre dans le zip. Nécessite les droits administrateur |

**Exemples**

```powershell
# Toutes les imprimantes, leurs paramètres et pilotes vers C:\Temp
.\Export-Printer.ps1

# Uniquement les imprimantes Office, paramètres seulement
.\Export-Printer.ps1 -Printer 'Office*' -NoDriver

# Aussi une sauvegarde PrintBrm, le zip sur un partage
.\Export-Printer.ps1 -PrintBrm -OutputPath \\fs01\it$\exports
```

**Remarques**
- Seules les imprimantes TCP/IP (Standard TCP/IP Port) vont dans `printers.json`, car c'est ce qu'Install-Printer crée. Une imprimante qu'il refuserait - un nom avec `\ , !`, une adresse qui n'est ni une adresse IP ni un nom d'hôte - est signalée et reste uniquement dans `inventory.json`
- Une imprimante sur un moniteur de port du fabricant (EpsonNet Print, IPP) dont le nom de port contient l'adresse - `10.1.50.235:`, `http://10.0.5.20:631/ipp` - est exportée comme port Standard TCP/IP (RAW 9100) vers cette adresse, avec son pilote et ses paramètres, un avertissement et une note dans sa `description` : ce moniteur de port n'existe que là où le logiciel du fabricant est installé. Imprimez une page de test sur la cible
- Les champs ne sont écrits que lorsqu'ils diffèrent des valeurs par défaut d'Install-Printer (port `IP_<adresse>`, port 9100, SNMP désactivé), si bien que le fichier se lit comme écrit à la main
- Les autorisations ne sont exportées que si elles diffèrent du défaut Windows (Everyone peut imprimer). Administrators, SYSTEM, CREATOR OWNER et les identités d'application sont omis, car Install-Printer les garde toujours. Les groupes connus sont écrits sous leur nom anglais, un utilisateur ou groupe Entra ID par son SID (un `AzureAD\nom` ne se résout nulle part ailleurs), un compte local de cette machine avec un avertissement
- Les paramètres d'impression sont lus dans un job avec un délai de 60 secondes, car certains pilotes universels se bloquent dans `Get-PrintConfiguration` ; un échec y est un avertissement et l'imprimante est exportée sans eux
- Les droits administrateur ne sont pas nécessaires, sauf pour `-PrintBrm`. Lancé en 32 bits, le script se relance en 64 bits, car le magasin de pilotes n'est pas visible depuis SysWOW64
- Variables de script NinjaOne : `printer`, `outputPath`, et les cases à cocher `noDriver`, `printBrm`

---

### New-PrinterGroup.ps1

Transforme une matrice imprimantes x utilisateurs - copiée depuis Excel - en un groupe de sécurité Entra ID par imprimante, nommé `<Prefix><nom de l'imprimante>` sans espaces (`Office - Duplex` → `SG-U-PRINTM-Office-Duplex`), et y ajoute les utilisateurs cochés. Avec `-ConfigPath`, il écrit chaque groupe dans l'entrée de cette imprimante dans `printers.json` sous la forme `"permissions": [ { "principal": "<SID du groupe>", "access": "print" } ]`, prêt pour `Install-Printer.ps1`. Aperçu par défaut ; `-Apply` modifie le tenant et le fichier.

La matrice : première colonne les noms d'imprimante, ligne d'en-tête les utilisateurs (nom d'affichage ou UPN), toute cellule non vide (`x`) signifie un accès. Séparée par tabulation, point-virgule ou virgule ; un en-tête sans la cellule vide en haut à gauche, comme Excel le copie, convient. `Nom (SERVICE)` correspond aussi au nom d'affichage `Nom`. Un nom qui ne correspond à aucun utilisateur, ou à plusieurs, est signalé et ignoré - mettez alors l'UPN dans l'en-tête.

**Paramètres**

| Paramètre | Description |
|-----------|-------------|
| `-MatrixPath` | Le fichier de matrice |
| `-Prefix` | Préfixe des noms de groupe (par défaut : `SG-U-PRINTM-`) |
| `-ConfigPath` | `printers.json` dans lequel écrire les autorisations. Les entrées `manage`/`manageDocuments` existantes sont conservées, `Everyone` et les autres entrées `print` remplacées ; un `.bak` est écrit d'abord |
| `-RemoveExtraMembers` | Retirer aussi les membres que la matrice ne coche pas (par défaut : signalés, laissés en place) |
| `-TenantId`, `-ClientId`, `-CertificateThumbprint`, `-AppOnly` | Connexion, via `Connect-M365.ps1` comme dans les autres scripts Entra |
| `-Apply` | Créer les groupes, modifier les membres, écrire le JSON. Sans : rapport uniquement |

**Exemples**

```powershell
# Ce qui serait créé et ajouté
.\New-PrinterGroup.ps1 -MatrixPath C:\Temp\printer-matrix.tsv

# Créer les groupes et écrire leurs SID dans un printers.json exporté
.\New-PrinterGroup.ps1 -MatrixPath C:\Temp\printer-matrix.tsv -ConfigPath C:\Temp\export\printers.json -Apply
```

**Remarques**
- **Un groupe Entra ID ne fonctionne dans une ACL d'imprimante que si le jeton Windows de l'utilisateur le contient** : un compte Entra ID qui se connecte à un appareil joint à Entra (hôtes de session AVD joints à Entra inclus), à partir de la connexion suivant l'ajout. Un utilisateur qui se connecte avec un compte Active Directory (serveur joint au domaine ou hybride, serveur d'impression classique) a des groupes AD dans son jeton, jamais des groupes cloud - la file refuse alors tout le monde sauf les administrateurs. Utilisez-y des groupes AD (`DOMAINE\Groupe`)
- Le SID est calculé à partir de l'ID d'objet du groupe (`S-1-12-1-` + le GUID en quatre nombres de 32 bits), la forme que Windows utilise pour les groupes Entra ID
- **Tenant** : sans `-TenantId`, l'exécution utilise le client GDAP, sinon le tenant auquel vous vous connectez - ou celui d'une session Graph déjà ouverte dans cette fenêtre. Elle nomme d'abord le tenant, et `-Apply` demande confirmation avant de créer quoi que ce soit (`-Confirm:$false` saute cette question). Passez `-TenantId <domaine>` pour en être sûr
- **Hybride** : les utilisateurs synchronisés depuis Active Directory se connectent aux machines jointes au domaine ou hybrides avec leur compte AD, et ce jeton ne contient jamais de groupe cloud. L'exécution avertit quand les utilisateurs trouvés sont synchronisés ; pour ces machines, créez les groupes dans AD (synchronisés vers Entra ID) et utilisez `DOMAINE\Groupe` comme principal
- Gardez la matrice hors de ce dépôt : elle nomme de vraies personnes. `C:\Temp` convient très bien
- Scopes requis : `Group.ReadWrite.All`, `User.Read.All`
