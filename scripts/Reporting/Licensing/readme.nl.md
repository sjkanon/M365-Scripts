[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [Reporting](../readme.nl.md) › **Licensing**

# Toolkit voor licentierapportage

> Auteur: Sjoerd Kanon

Genereert per klant een maandelijks overzicht van licenties en Azure-kosten als opgemaakt Excel-bestand. Combineert factuurgegevens van **Pax8** (CSV) en **Ingram** (Excel) in één rapport met een overzichtstabblad en een eigen tabblad per klant.

---

## Mapstructuur

```
Licensing/
├── genereer_licentie_overzicht.py   ← Python-engine — leest Pax8 + Ingram, schrijft Excel
├── genereer_rapport.ps1             ← PowerShell-launcher — controleert de omgeving, roept Python aan
├── genereer_rapport.bat             ← Eenvoudige launcher om op te dubbelklikken (geen controles)
└── create_scheduled_task.ps1        ← Registreert een maandelijkse geplande taak (eenmalig uitvoeren)
```

---

## Vereisten

| Vereiste | Versie |
|---|---|
| Python | 3.8 of hoger |
| pandas | `pip install pandas` |
| openpyxl | `pip install openpyxl` |
| OneDrive | Gesynchroniseerd en aangemeld |

Afhankelijkheden installeren:

```bash
pip install pandas openpyxl
```

---

## Invoerbestanden

Zet de invoerbestanden vóór het uitvoeren in de juiste submappen onder de exportmap:

```
C:\OneDrive\<Company>\<Company> - Finance - Licenses_facturatie_upload\
├── Import\
│   ├── Ingram\     ← zet hier precies 1 .xlsx-bestand (factuurexport van Ingram)
│   └── Pax8\       ← zet hier precies 1 .csv-bestand (factuurexport van Pax8)
├── Archive\        ← invoerbestanden worden hier na verwerking automatisch naartoe verplaatst
└── Licentie_Overzicht_YYYY-MM.xlsx   ← uitvoer wordt hier geschreven
```

> Het script stopt met een duidelijke foutmelding als: de OneDrive-map niet bereikbaar is, er geen bestand gevonden wordt, of er meer dan één bestand in een map staat.

---

## Gebruik

### Optie 1 — PowerShell-launcher (aanbevolen)

Controleert de omgeving voordat het wordt uitgevoerd. Toont duidelijke foutmeldingen als er iets ontbreekt.

```powershell
.\genereer_rapport.ps1
```

### Optie 2 — Dubbelklikken

Voer `genereer_rapport.bat` rechtstreeks uit. Geen controles vooraf — vertrouwt op de eigen foutafhandeling van het Python-script.

### Optie 3 — CLI met expliciete paden

```bash
python genereer_licentie_overzicht.py --ingram "path\to\ingram.xlsx" --pax8 "path\to\pax8.csv"
python genereer_licentie_overzicht.py --ingram "path\to\ingram.xlsx" --pax8 "path\to\pax8.csv" --output "C:\output\rapport.xlsx"
```

---

## Uitvoer

Het gegenereerde Excel-bestand bevat:

| Tabblad | Inhoud |
|---|---|
| `Overzicht` | Samenvatting van alle klanten — totaal inkoop, totaal verkoop, marge |
| `<Customer name>` | Eén tabblad per klant met alle secties |

### Secties per klanttabblad

| Sectie | Bron | Inhoud |
|---|---|---|
| ☁ Azure (via Pax8) | Pax8 | Azure-verbruik gegroepeerd per abonnement → categorie |
| ☁ Azure (via Ingram) | Ingram | Azure-verbruik gegroepeerd per abonnement → categorie |
| 📋 Licenties (via Pax8) | Pax8 | M365- en andere licenties gegroepeerd per product |
| 📋 Licenties (via Ingram) | Ingram | Licenties gegroepeerd per product |
| 📋 Acronis Backup | Pax8 | Getoond op het tabblad van de eindklant als Acronis via een reseller wordt gefactureerd |
| **TOTAAL** | — | Eindtotaal inkoop + verkoop voor deze klant |

Elke rij toont: omschrijving, categorie, aantal, inkoopprijs per stuk, verkoopprijs per stuk, totaal inkoop, totaal verkoop.

---

## Geplande taak

`create_scheduled_task.ps1` registreert een geplande Windows-taak die het Python-script automatisch uitvoert op de **6e van elke maand om 08:00**.

### Configuratie (aanpassen vóór het uitvoeren)

| Variabele | Standaard | Omschrijving |
|---|---|---|
| `$TaskName` | `"Licensing Report Generator"` | Naam van de taak in Taakplanner |
| `$RunAsUser` | `"$env:USERDOMAIN\sa-halo"` | Serviceaccount dat de taak uitvoert — **aanpassen voor jouw domein** |
| `$RunDay` | `6` | Dag van de maand waarop de taak draait |
| `$RunTime` | `"08:00"` | Tijdstip |

Het script detecteert `python.exe` automatisch via PATH en gangbare installatielocaties. `$ScriptPath` wordt automatisch bepaald ten opzichte van de scriptmap.

### Eenmalig uitvoeren om te registreren:

```powershell
# Uitvoeren als Administrator
.\create_scheduled_task.ps1
```

### Na registratie handmatig testen:

```powershell
Start-ScheduledTask -TaskName "... Licentie Overzicht Generator"
```

> **Let op:** Het serviceaccount (`$RunAsUser`) moet leestoegang hebben tot de importmappen van Ingram/Pax8 en schrijftoegang tot de exportmap in OneDrive.

---

## Logbestand

Er wordt een logbestand geschreven naar `Log\licentie_rapport.log` in de scriptmap. Elke run voegt een blok met tijdstempel toe met `[INFO]`- en `[ERROR]`-regels. Controleer dit bestand als het rapport via de geplande taak stilletjes mislukt.

---

## Aanpassen

### Labels voor Pax8-abonnementen

Bewerk `PAX8_SUB_LABELS` in `genereer_licentie_overzicht.py` om ruwe abonnements-ID's aan leesbare namen te koppelen:

```python
PAX8_SUB_LABELS = {
    "MY-SUB-001": "My Subscription Label",
}
```

### Aliassen voor Acronis-eindklanten

Bewerk `ACRONIS_ENDCUSTOMER_ALIASES` om namen van eindklanten in Pax8 te koppelen aan de namen die elders in het rapport worden gebruikt:

```python
ACRONIS_ENDCUSTOMER_ALIASES = {
    "Customer Name in Pax8": "Customer Name in Report",
}
```

### OneDrive-pad

Het basispad is hardcoded in zowel `genereer_rapport.ps1` als `genereer_licentie_overzicht.py`:

```
C:\OneDrive\<Company>\<Company> - Finance - Licenses_facturatie_upload
```

Pas `$ExportDir` in `genereer_rapport.ps1` en `EXPORT_DIR` in `genereer_licentie_overzicht.py` aan zodat ze overeenkomen met de naam van jouw OneDrive-map.

---

## Probleemoplossing

| Fout | Oorzaak | Oplossing |
|---|---|---|
| `OneDrive map niet bereikbaar` | OneDrive niet gesynchroniseerd of niet aangemeld | Meld je aan bij OneDrive en wacht tot de synchronisatie klaar is |
| `Geen Excel bestand gevonden in Import\Ingram\` | Ingram-bestand niet geplaatst | Zet precies 1 `.xlsx` in de map `Ingram` |
| `Meerdere bestanden gevonden` | Meer dan 1 bestand in een map | Verwijder of archiveer het extra bestand |
| `Python niet gevonden` | Python staat niet in PATH | Installeer Python en voeg het toe aan PATH, of gebruik het volledige pad in de taak |
| Lege secties in de uitvoer | Klantnaam verschilt tussen Pax8 en Ingram | Controleer de bronbestanden op afsluitende spaties of verschillend hoofdlettergebruik |
