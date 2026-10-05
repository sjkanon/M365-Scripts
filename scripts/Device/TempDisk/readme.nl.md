[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [Device](../readme.nl.md) › **TempDisk**

# Tijdelijke schijf

Houdt de tijdelijke (ephemeral) schijf (`D:`) van een Azure-VM of AVD-sessiehost op zijn plek, en houdt de pagefile erop.

De tijdelijke schijf van een Azure-VM — de resource disk, of op de nieuwere groottes de lokale NVMe-schijf — wordt gewist zodra de VM wordt gedealloceerd, van grootte verandert of naar een andere host verhuist. Hij komt leeg terug, soms RAW, soms offline, soms zonder stationsletter. Windows leest de pagefileconfiguratie bij het opstarten, dus **een pagefile die is ingesteld op een stationsletter die er bij het opstarten niet is, wordt gewoon niet aangemaakt**: de machine pagineert dan weer op `C:`, of draait helemaal zonder pagefile. Dat is wat deze twee scripts moeten voorkomen.

---

## Scripts

| Script | Omschrijving |
|--------|-------------|
| [`Init-TempDisk.ps1`](Init-TempDisk.ps1) ([docs](#init-tempdiskps1)) | Herstelt de tijdelijke schijf als `D:` en stelt de pagefile daarop in |
| [`Register-InitTempDiskTask.ps1`](Register-InitTempDiskTask.ps1) ([docs](#register-inittempdisktaskps1)) | Installeert dat script op het apparaat en draait het bij elke start als SYSTEM |

---

### Init-TempDisk.ps1

Per run:

1. **Preflight** — elke schijf, wat `D:` bezet, de pagefile zoals geconfigureerd (register) en de pagefile die in deze sessie echt in gebruik is
2. **Letter** — een optisch station dat op `D:` zit, wordt opzij gezet; op een image zonder tijdelijke schijf geeft Windows `D:` aan de dvd-speler en geeft hem nooit meer terug
3. **Schijf** — een volume dat al de tijdelijke schijf *is*, krijgt zijn stationsletter terug; anders wordt een RAW-schijf die geen boot- of systeemschijf is online gebracht, als GPT geïnitialiseerd, gepartitioneerd en als NTFS geformatteerd
4. **Pagefile** — automatisch beheer uit, de pagefile naar `D:\pagefile.sys`, de vermelding voor elk ander station verwijderd
5. **Verificatie** — alles wordt teruggelezen, met wat nu van kracht is en wat op de volgende herstart wacht
6. **Herstart** — alleen met `-RestartIfNeeded`: herstart de machine als dat het enige is wat nog tussen de configuratie en een pagefile die echt in gebruik is staat

Een tijdelijke schijf die alleen zijn stationsletter kwijt is, krijgt die letter terug en wordt nooit opnieuw geformatteerd — het volume wordt herkend aan zijn label (`Temporary Storage`) of aan de `DataLoss_Warning_Readme.txt` die Azure op de resource disk schrijft.

**Parameters**

| Parameter | Omschrijving |
|-----------|-------------|
| `-DriveLetter` | Stationsletter voor de tijdelijke schijf (standaard: `D`) |
| `-Label` | Label dat bij het formatteren wordt geschreven, en het label waaraan een bestaand tijdelijk volume wordt herkend (standaard: `Temporary Storage`) |
| `-DiskNumber` | Formatteer deze schijf in plaats van het script er een te laten kiezen — verplicht als er meerdere RAW-schijven zijn |
| `-InitialSizeMB` / `-MaximumSizeMB` | Grootte van de pagefile in MB. `0` (de standaard) op beide betekent door het systeem beheerd |
| `-KeepSystemDrivePagefile` | Laat een bestaande pagefile op `C:` staan in plaats van hem te verwijderen |
| `-SkipPagefile` | Herstel alleen de schijf; laat de pagefileconfiguratie ongemoeid |
| `-OpticalDriveLetter` | Letter waarnaar een optisch station wordt verplaatst als het `D:` bezet (standaard: `Z`) |
| `-RestartIfNeeded` | Herstart de machine als dat het enige is wat nog tussen de configuratie en een pagefile in gebruik staat. Standaard uit |
| `-RestartDelaySeconds` | Aftelling vóór een herstart terwijl er iemand is aangemeld (standaard: `60`) — `shutdown /a` breekt hem af. Is er niemand aangemeld, dan herstart hij binnen enkele seconden |
| `-RestartCooldownMinutes` | Kortste tijd tussen twee herstarts die dit script heeft veroorzaakt (standaard: `60`) |
| `-RestartMarkerPath` | Registersleutel waar de laatste door het script veroorzaakte herstart wordt bijgehouden (standaard: `HKLM:\SOFTWARE\M365-Scripts\InitTempDisk`) |
| `-RestartEvenIfUsersSignedIn` | Herstart ook als er iemand is aangemeld. Op een sessiehost: zet hem liever in drain mode |
| `-CheckOnly` | Alleen rapporteren, niets wijzigen. Exitcode `2` betekent dat er werk te doen is |
| `-Quiet` | Toon niets tenzij er nieuws is — het logbestand krijgt altijd het volledige verhaal |
| `-LogPath` | Map voor `Init-TempDisk.log` (standaard: `C:\Temp`), aangevuld en geroteerd boven 1 MB |
| `-Force` | Met `-DiskNumber`: formatteer die schijf ook al heeft hij nog partities |

**Voorbeelden**

```powershell
# Alleen-lezen gezondheidsrapport: waar de tijdelijke schijf is en wat de pagefile echt doet
.\Init-TempDisk.ps1 -CheckOnly

# De hele flow doorlopen zonder de machine aan te raken
.\Init-TempDisk.ps1 -WhatIf

# Zoals de geplande taak het draait — stil zolang alles in orde is,
# en één herstart als de pagefile daarop wacht
.\Init-TempDisk.ps1 -Quiet -RestartIfNeeded

# Vaste pagefile van 16 GB in plaats van een door het systeem beheerde
.\Init-TempDisk.ps1 -InitialSizeMB 16384 -MaximumSizeMB 16384

# Aangeven welke schijf de tijdelijke schijf is, ook al heeft hij nog partities
.\Init-TempDisk.ps1 -DiskNumber 2 -Force
```

**Exitcodes**

| Code | Betekenis |
|------|---------|
| `0` | De tijdelijke schijf en de pagefile zijn zoals ze moeten zijn |
| `1` | Fout |
| `2` | Alleen bij `-CheckOnly`: er is werk te doen |

**Opmerkingen**
- **Niets wat niet RAW is, wordt ooit geïnitialiseerd.** Een lege tijdelijke schijf en een ongeformatteerde dataschijf zien er van buitenaf hetzelfde uit, dus een schijf die al partities heeft, wordt gerapporteerd en met rust gelaten. Bij meer dan één RAW-kandidaat weigert het script te gokken en vraagt het om `-DiskNumber`; `-Force` plus `-DiskNumber` is de enige manier om een schijf te formatteren die nog partities heeft
- Een lokale NVMe-schijf gaat voor op andere RAW-schijven als er meerdere zijn — op de nieuwere VM-groottes *is* dat de tijdelijke schijf
- **Een pagefile die een run instelt, verschijnt bij de volgende herstart.** Windows leest de configuratie bij het opstarten en leest hem daarna nooit opnieuw, dus de run zegt dat ook in plaats van succes te claimen. Omdat de taak bij elke start draait, herstelt het apparaat zichzelf ook zonder `-RestartIfNeeded`: de start die `D:` opnieuw aanmaakt, stelt de pagefile in, de start daarna neemt hem in gebruik
- `-RestartIfNeeded` dicht dat gat in plaats van erop te wachten, en een script dat bij elke start draait en de machine mag herstarten, is een herstartlus die op zijn kans wacht — daarom gaat hij alleen af als **al** het volgende klopt: de run is foutloos afgerond, de schijf is er, de pagefile is erop ingesteld en alleen deze sessie gebruikt hem niet; er is niemand aangemeld (verbonden *of* verbroken — één `explorer.exe` per bureaublad, wat taalonafhankelijk is, waar het parsen van `query.exe` dat niet is); en er is in de laatste `-RestartCooldownMinutes` geen herstart veroorzaakt, bijgehouden onder `-RestartMarkerPath`. Een mislukte run herstart nooit — dan zou de fout achter een reboot verdwijnen
- De herstart loopt via `shutdown.exe` met de geplande reden "Operating System: Reconfiguration", zodat hij als bedoeld verschijnt en niet als onverwacht. De aftelling is er om mensen te waarschuwen, dus geldt alleen als er mensen zijn: met iemand aangemeld is het `-RestartDelaySeconds` en stopt `shutdown /a` hem; is er niemand aangemeld — het normale geval bij het opstarten, en het gegarandeerde op een sessiehost waarvan de pool in drain mode staat — dan herstart hij binnen enkele seconden, want een minuut wachten op een publiek van niemand kost alleen beschikbaarheid
- Een pagefile instellen op een station dat er niet is, zou een instelling schrijven die Windows negeert, dus de run faalt in plaats daarvan als `D:` niet kon worden hersteld
- Vereist administratorrechten; handmatig gestart vanuit een gewoon venster vraagt het zelf om verhoging

---

### Register-InitTempDiskTask.ps1

Draai het één keer per apparaat — handmatig, vanuit Tactical RMM / NinjaOne, of als Intune-platformscript. Kopieert `Init-TempDisk.ps1` naar een lokale map (de repository is er bij het opstarten niet) en registreert een taak die het bij elke start draait, als SYSTEM, met de hoogste rechten en zonder dat iemand zich aanmeldt.

Opnieuw draaien is veilig: een taak met dezelfde naam wordt vervangen, dus zo wijzig je ook de argumenten waarmee de taak draait.

De taak draait standaard `Init-TempDisk.ps1 -Quiet -RestartIfNeeded`. Windows leest de pagefileconfiguratie bij het opstarten, dus een start waarbij de tijdelijke schijf opnieuw moest worden opgebouwd, draait die hele sessie zonder pagefile op `D:`, tenzij de machine één keer herstart — en de waarborgen hierboven maken het veilig om dat aan een opstarttaak over te laten. Geef `-ScriptArguments '-Quiet'` mee om de herstart weg te laten.

**Parameters**

| Parameter | Omschrijving |
|-----------|-------------|
| `-ScriptSourcePath` | `Init-TempDisk.ps1` dat wordt geïnstalleerd (standaard: de kopie naast dit script) |
| `-ScriptTargetDir` | Map op het apparaat waarnaar het wordt gekopieerd (standaard: `C:\Scripts`) |
| `-TaskName` | Naam van de geplande taak (standaard: `InitTempDisk`) |
| `-ScriptArguments` | Argumenten voor `Init-TempDisk.ps1` (standaard: `-Quiet -RestartIfNeeded`) |
| `-DelaySeconds` | Vertraging tussen het opstarten en de start van de taak (standaard: `30`) |
| `-RunNow` | Start de taak ook direct één keer, in plaats van op een herstart te wachten |
| `-Unregister` | Verwijder de taak en de geïnstalleerde kopie van het script |

**Voorbeelden**

```powershell
# Tonen wat er zou worden geïnstalleerd en geregistreerd
.\Register-InitTempDiskTask.ps1 -WhatIf

# Installeren, de opstarttaak registreren en hem nu één keer draaien
.\Register-InitTempDiskTask.ps1 -RunNow

# Hetzelfde, maar met een vaste pagefile van 16 GB
.\Register-InitTempDiskTask.ps1 -ScriptArguments '-Quiet -RestartIfNeeded -InitialSizeMB 16384 -MaximumSizeMB 16384'

# Zonder de herstart — de pagefile komt er bij de eerstvolgende herstart
.\Register-InitTempDiskTask.ps1 -ScriptArguments '-Quiet'

# Weer van het apparaat halen
.\Register-InitTempDiskTask.ps1 -Unregister
```

**Opmerkingen**
- De taak draait de *kopie* in `C:\Scripts`, nooit de bron, dus de repository of de stagingmap van de RMM mag daarna verdwijnen
- Een ontbrekend bronbestand is een harde fout, geen stille overslag — een taak die is geregistreerd op een bestand dat er niet is, draait en faalt bij elke start zonder dat iemand het merkt, tot de pagefile weg is
- De opstarttrigger is standaard 30 seconden vertraagd: de opslagstack heeft de tijdelijke schijf niet altijd al opgesomd op het moment dat de taakengine draait
- `-Unregister` laat de pagefileconfiguratie bewust met rust — het verwijderen van de taak mag de pagefile van een machine niet meenemen
- Bij het registreren wordt hardop gemeld of de taak de machine mag herstarten; `-RunNow` op een apparaat waar niemand is aangemeld, kan daardoor een aftelling van 60 seconden starten
