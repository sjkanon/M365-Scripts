[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [Azure](../readme.fr.md) › **VM**

# VM

Maintenance des machines virtuelles Azure.

## Scripts

| Script | Description |
|--------|-------------|
| [`Azure-NVMe-Conversion.ps1`](Azure-NVMe-Conversion.ps1) ([docs](#azure-nvme-conversionps1)) | Convertir le type de contrôleur de disque d'une VM Azure entre SCSI et NVMe, y compris la préparation des pilotes dans le système invité (script Microsoft intégré tel quel) |

---

### Azure-NVMe-Conversion.ps1

> **Script tiers intégré tel quel.** Il s'agit de l'outil de Microsoft lui-même, issu de [`Azure/SAP-on-Azure-Scripts-and-Utilities`](https://github.com/Azure/SAP-on-Azure-Scripts-and-Utilities) (licence MIT) — conservé en l'état plutôt que réécrit, puisqu'il est déjà maintenu en amont. Vérifiez le `.LINK` dans l'en-tête du script pour obtenir la dernière version avant de vous y fier pour une opération critique.

Convertit le type de contrôleur de disque d'une VM Azure entre SCSI et NVMe, y compris la préparation des pilotes dans le système invité. Changer de type de contrôleur modifie la façon dont les disques sont présentés au sein du système d'exploitation ; faire passer une VM déjà provisionnée à une taille exclusivement NVMe (par ex. `Standard_E*bds_v5`/`v6`) sans préparer d'abord le système invité peut provoquer un `INACCESSIBLE_BOOT_DEVICE` au démarrage.

**Ce qu'il fait**

1. Valide la VM cible (existante, en cours d'exécution, image Gen2, type de contrôleur actuel, SKU cible prenant en charge le type de contrôleur demandé et disponible dans la zone de la VM)
2. Pour les VM Windows converties en NVMe : exécute une **vérification en lecture seule** dans le système invité via `Invoke-AzVMRunCommand` (vérifie que le pilote `stornvme` est présent et démarré au boot dans *chaque* `ControlSet`, pas seulement le jeu actuel) — passez `-FixOperatingSystemSettings` pour exécuter aussi la **correction** (supprime la clé de registre `StartOverride` qui empêche le pilote de se charger au démarrage, avec un vidage explicite du registre pour que la modification survive à une désallocation immédiate)
3. Pour les VM Linux : vérifie/corrige le pilote `nvme` dans `initrd`/`initramfs` selon la distribution (Ubuntu/Debian : `update-initramfs` ; famille RHEL/SUSE : `dracut`)
4. Met à jour les capacités prises en charge du disque du système d'exploitation ainsi que le type de contrôleur de disque / la taille de la VM
5. Redémarre éventuellement la VM (`-StartVM`) et écrit un fichier journal horodaté (`-WriteLogfile`)

**Paramètres**

| Paramètre | Obligatoire | Valeur par défaut | Description |
|-----------|----------|---------|-------------|
| `-ResourceGroupName` | Oui | — | Groupe de ressources contenant la VM |
| `-VMName` | Oui | — | VM à convertir |
| `-VMSize` | Oui | — | Taille cible de la VM (doit prendre en charge le type de contrôleur cible) |
| `-NewControllerType` | Non | `NVMe` | `NVMe` ou `SCSI` |
| `-StartVM` | Non | désactivé | Démarrer la VM après la conversion |
| `-WriteLogfile` | Non | désactivé | Écrire `Azure-NVMe-Conversion-<VMName>-<timestamp>.log` dans le répertoire courant |
| `-FixOperatingSystemSettings` | Non | désactivé | Appliquer réellement la correction des pilotes dans le système invité (la VM doit être en cours d'exécution et le VM Agent prêt) |
| `-IgnoreOSCheck` | Non | désactivé | Ignorer entièrement la vérification de préparation dans le système invité |
| `-IgnoreSKUCheck` | Non | désactivé | Ignorer la validation de disponibilité/capacités du SKU cible |
| `-IgnoreWindowsVersionCheck` | Non | désactivé | Ignorer la vérification de l'exigence Windows Server 2019+ / Windows 10 1809+ |
| `-IgnoreAzureModuleCheck` | Non | désactivé | Ignorer les vérifications de version de `Az.Compute`/`Az.Accounts`/`Az.Resources` |
| `-SleepSeconds` | Non | `15` | Délai après la mise à jour de la taille/du contrôleur de la VM avant de poursuivre |

**Exemple**

```powershell
Connect-AzAccount
.\Azure-NVMe-Conversion.ps1 -ResourceGroupName "myResourceGroup" -VMName "myVM" `
    -NewControllerType NVMe -VMSize "Standard_E4bds_v5" -FixOperatingSystemSettings -StartVM -WriteLogfile
```

**Modules requis**

```powershell
Install-Module Az.Accounts  -MinimumVersion 4.0 -Scope CurrentUser
Install-Module Az.Compute   -MinimumVersion 9.0 -Scope CurrentUser
Install-Module Az.Resources -MinimumVersion 7.0 -Scope CurrentUser
```

**Remarques**
- Ne fait pas partie du lanceur interactif `menu.ps1` — ce menu cible le tenant M365 (Graph/Exchange), alors que ce script cible Azure IaaS via le module `Az` et une session `Connect-AzAccount` distincte.
- Nécessite `Virtual Machine Contributor` (ou équivalent) sur le groupe de ressources cible.
- Chaque démarrage en SCSI recrée la clé de registre bloquante `StartOverride` — ne démarrez pas la VM en SCSI entre l'application de la correction et la conversion en NVMe.
