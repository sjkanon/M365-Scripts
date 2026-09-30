[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [LegacyUtilities](../readme.nl.md) › **Device**

# Legacy Utilities — Device

Kleine hulpscripts voor het configureren van werkstations.

---

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`Set-NumLockDefault.ps1`](Set-NumLockDefault.ps1) ([docs](#set-numlockdefaultps1)) | Stel de standaardstatus van Num Lock in voor nieuwe profielen en het aanmeldscherm |
| [`New-LockWorkstationShortcut.ps1`](New-LockWorkstationShortcut.ps1) ([docs](#new-lockworkstationshortcutps1)) | Maak een snelkoppeling om het werkstation te vergrendelen |

---

### Set-NumLockDefault.ps1

Schrijft `InitialKeyboardIndicators` onder `HKU\.DEFAULT` (de sjabloon waarop nieuwe
profielen worden gebaseerd, en die ook op het aanmeldscherm wordt gebruikt) en in de eigen
hive van de huidige gebruiker. Standaard een proefdraai.

```powershell
.\Set-NumLockDefault.ps1 -State On -Apply
```

---

### New-LockWorkstationShortcut.ps1

Maakt een `.lnk`-snelkoppeling die het standaardcommando
`rundll32.exe user32.dll,LockWorkStation` uitvoert: geen externe download, geen
registertrucs om iets vast te maken. Gegeneraliseerde vervanging van een script dat een
eigen pictogram/batchbestand van een intern endpoint downloadde en dat via een
ongedocumenteerde registersleutel aan de taakbalk vastmaakte; deze versie maakt alleen de
snelkoppeling aan en laat de gebruiker die zelf vastmaken als hij dat wil. Standaard een proefdraai.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-TargetFolder` | Nee | Waar de snelkoppeling wordt aangemaakt (standaard: Public Desktop) |
| `-ShortcutName` | Nee | Standaard: "Lock Workstation" |
| `-IconPath` | Nee | Optioneel lokaal `.ico`-bestand |
| `-Apply` | Nee | Maak de snelkoppeling echt aan (standaard: voorbeeldweergave) |

```powershell
.\New-LockWorkstationShortcut.ps1 -Apply
```
