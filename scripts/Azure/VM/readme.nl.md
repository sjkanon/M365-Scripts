[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [Azure](../readme.nl.md) › **VM**

# VM

Onderhoud van virtuele Azure-machines.

## Scripts

| Script | Omschrijving |
|--------|--------------|
| [`Azure-NVMe-Conversion.ps1`](Azure-NVMe-Conversion.ps1) ([docs](#azure-nvme-conversionps1)) | Het type schijfcontroller van een Azure-VM omzetten tussen SCSI en NVMe, inclusief voorbereiding van de drivers in de gast (overgenomen script van Microsoft) |

---

### Azure-NVMe-Conversion.ps1

> **Overgenomen script van derden.** Dit is de eigen tool van Microsoft uit [`Azure/SAP-on-Azure-Scripts-and-Utilities`](https://github.com/Azure/SAP-on-Azure-Scripts-and-Utilities) (MIT-licentie) — ongewijzigd overgenomen in plaats van herschreven, omdat het upstream al wordt onderhouden. Controleer de `.LINK` in de scriptheader op de nieuwste versie voordat je er voor iets kritieks op vertrouwt.

Zet het type schijfcontroller van een Azure-VM om tussen SCSI en NVMe, inclusief voorbereiding van de drivers in het gastbesturingssysteem. Een ander controllertype verandert hoe schijven binnen het OS worden aangeboden, dus een al ingerichte VM omzetten naar een grootte die alleen NVMe ondersteunt (bijv. `Standard_E*bds_v5`/`v6`) zonder de gast eerst voor te bereiden, kan bij het opstarten `INACCESSIBLE_BOOT_DEVICE` veroorzaken.

**Wat het doet**

1. Valideert de doel-VM (bestaat, draait, Gen2-image, huidig controllertype, doel-SKU ondersteunt het gevraagde controllertype en is beschikbaar in de zone van de VM)
2. Voor Windows-VM's die naar NVMe gaan: voert via `Invoke-AzVMRunCommand` een **alleen-lezen controle** in de gast uit (controleert of de driver `stornvme` aanwezig is en bij het opstarten wordt gestart in *elke* `ControlSet`, niet alleen de huidige) — geef `-FixOperatingSystemSettings` mee om ook de **reparatie** uit te voeren (verwijdert de registersleutel `StartOverride` die voorkomt dat de driver bij het opstarten laadt, met een expliciete registerflush zodat de wijziging een direct daaropvolgende deallocate overleeft)
3. Voor Linux-VM's: controleert/repareert de driver `nvme` in `initrd`/`initramfs`, afhankelijk van de distributie (Ubuntu/Debian: `update-initramfs`; RHEL-familie/SUSE: `dracut`)
4. Werkt de ondersteunde mogelijkheden van de OS-schijf en het schijfcontrollertype / de grootte van de VM bij
5. Start de VM optioneel opnieuw (`-StartVM`) en schrijft een logbestand met tijdstempel (`-WriteLogfile`)

**Parameters**

| Parameter | Verplicht | Standaard | Omschrijving |
|-----------|----------|---------|-------------|
| `-ResourceGroupName` | Ja | — | Resourcegroep die de VM bevat |
| `-VMName` | Ja | — | VM die moet worden omgezet |
| `-VMSize` | Ja | — | Doelgrootte van de VM (moet het doelcontrollertype ondersteunen) |
| `-NewControllerType` | Nee | `NVMe` | `NVMe` of `SCSI` |
| `-StartVM` | Nee | uit | Start de VM na de omzetting |
| `-WriteLogfile` | Nee | uit | Schrijf `Azure-NVMe-Conversion-<VMName>-<timestamp>.log` naar de huidige map |
| `-FixOperatingSystemSettings` | Nee | uit | Voer de driverreparatie in de gast echt uit (vereist dat de VM draait en de VM Agent gereed is) |
| `-IgnoreOSCheck` | Nee | uit | Sla de gereedheidscontrole in de gast volledig over |
| `-IgnoreSKUCheck` | Nee | uit | Sla de validatie van beschikbaarheid/mogelijkheden van de doel-SKU over |
| `-IgnoreWindowsVersionCheck` | Nee | uit | Sla de controle op de vereiste Windows Server 2019+ / Windows 10 1809+ over |
| `-IgnoreAzureModuleCheck` | Nee | uit | Sla de versiecontroles van `Az.Compute`/`Az.Accounts`/`Az.Resources` over |
| `-SleepSeconds` | Nee | `15` | Wachttijd na het bijwerken van VM-grootte/controller voordat het script verdergaat |

**Voorbeeld**

```powershell
Connect-AzAccount
.\Azure-NVMe-Conversion.ps1 -ResourceGroupName "myResourceGroup" -VMName "myVM" `
    -NewControllerType NVMe -VMSize "Standard_E4bds_v5" -FixOperatingSystemSettings -StartVM -WriteLogfile
```

**Vereiste modules**

```powershell
Install-Module Az.Accounts  -MinimumVersion 4.0 -Scope CurrentUser
Install-Module Az.Compute   -MinimumVersion 9.0 -Scope CurrentUser
Install-Module Az.Resources -MinimumVersion 7.0 -Scope CurrentUser
```

**Opmerkingen**
- Geen onderdeel van de interactieve starter `menu.ps1` — dat menu richt zich op de M365-tenant (Graph/Exchange), dit script richt zich op Azure IaaS via de module `Az` en een aparte `Connect-AzAccount`-sessie.
- Vereist `Virtual Machine Contributor` (of gelijkwaardig) op de doelresourcegroep.
- Elke keer dat de VM op SCSI opstart, wordt de blokkerende registersleutel `StartOverride` opnieuw aangemaakt — start de VM dus niet op SCSI tussen het uitvoeren van de reparatie en de omzetting naar NVMe.
