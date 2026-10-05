[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../readme.nl.md) › [scripts](../../readme.nl.md) › [Reporting](../readme.nl.md) › **Licensing**

# Toolkit voor licentierapportage

> Auteur: Sjoerd Kanon

Genereert per klant een maandelijks overzicht van licenties en Azure-kosten als opgemaakt Excel-bestand. Combineert factuurgegevens van **Pax8** (CSV) en **Ingram** (Excel) in één rapport met een overzichtstabblad en een eigen tabblad per klant.

---

## Scripts

| Script | Omschrijving |
|--------|--------------|
| [`genereer_rapport.ps1`](genereer_rapport.ps1) ([docs](#genereer_rapportps1)) | PowerShell-launcher — controleert de omgeving en roept daarna de Python-engine aan (aanbevolen) |
| [`genereer_rapport.bat`](genereer_rapport.bat) ([docs](#genereer_rapportbat)) | Eenvoudige launcher om te dubbelklikken, zonder controles — vereist `LICENSING_EXPORT_DIR` |
| [`genereer_licentie_overzicht.py`](genereer_licentie_overzicht.py) ([docs](#genereer_licentie_overzichtpy)) | Python-engine — leest de Pax8-CSV en de Ingram-Excel, schrijft het Excel-rapport per klant |
| [`create_scheduled_task.ps1`](create_scheduled_task.ps1) ([docs](#create_scheduled_taskps1)) | Registreert een maandelijkse geplande taak die de Python-engine uitvoert (eenmalig, als Administrator) |

---

## Vereisten

| Vereiste | Versie |
|---|---|
| Python | 3.8 of hoger |
| pandas | `pip install pandas` |
| openpyxl | `pip install openpyxl` |
| Exportmap | Bereikbaar voor het account dat het rapport draait (een gesynchroniseerde OneDrive-map volstaat) |

Afhankelijkheden installeren:

```bash
pip install pandas openpyxl
```

---

## Invoerbestanden

Zet de invoerbestanden vóór het uitvoeren in de juiste submappen onder de exportmap:

```
<ExportDir>\
├── Import\
│   ├── Ingram\     ← zet hier precies 1 .xlsx-bestand (factuurexport van Ingram)
│   └── Pax8\       ← zet hier precies 1 .csv-bestand (factuurexport van Pax8)
├── Archive\        ← invoerbestanden worden hier na verwerking automatisch naartoe verplaatst
└── Licentie_Overzicht_YYYY-MM.xlsx   ← uitvoer wordt hier geschreven
```

> Het script stopt met een duidelijke foutmelding als: er geen exportmap is opgegeven, de exportmap niet bereikbaar is, er geen bestand gevonden wordt, of er meer dan één bestand in een map staat.

---

## Gebruik

### genereer_rapport.ps1

Optie 1 — PowerShell-launcher (aanbevolen). Controleert de omgeving voordat het wordt uitgevoerd. Toont duidelijke foutmeldingen als er iets ontbreekt.

```powershell
.\genereer_rapport.ps1 -ExportDir "D:\Finance\Licensing"
```

### genereer_rapport.bat

Optie 2 — dubbelklikken. Voer `genereer_rapport.bat` rechtstreeks uit — vereist de omgevingsvariabele `LICENSING_EXPORT_DIR`. Geen controles vooraf — vertrouwt op de eigen foutafhandeling van het Python-script.

### genereer_licentie_overzicht.py

Optie 3 — de Python-engine rechtstreeks aanroepen met expliciete paden:

```bash
python genereer_licentie_overzicht.py --export-dir "D:\Finance\Licensing" --ingram "path\to\ingram.xlsx" --pax8 "path\to\pax8.csv"
python genereer_licentie_overzicht.py --ingram "path\to\ingram.xlsx" --pax8 "path\to\pax8.csv" --export-dir "D:\Finance\Licensing" --output "C:\output\rapport.xlsx"
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

## create_scheduled_task.ps1

Geplande taak. `create_scheduled_task.ps1` registreert een geplande Windows-taak die het Python-script automatisch uitvoert standaard op de **6e van elke maand om 08:00**.

### Parameters

| Parameter | Standaard | Omschrijving |
|---|---|---|
| `-ExportDir` | *(verplicht)* | Map met `Import\` en `Archive\`, waar het rapport wordt weggeschreven; wordt als `--export-dir` aan het Python-script doorgegeven |
| `-RunAsUser` | *(verplicht)* | Serviceaccount dat de taak uitvoert, bv. `CONTOSO\svc-licensing` |
| `-RunDay` | `6` | Dag van de maand waarop de taak draait (1–28) |
| `-RunTime` | `"08:00"` | Tijdstip |

Het script detecteert `python.exe` automatisch via PATH en gangbare installatielocaties. `$ScriptPath` wordt automatisch bepaald ten opzichte van de scriptmap.

### Eenmalig uitvoeren om te registreren:

```powershell
# Uitvoeren als Administrator
.\create_scheduled_task.ps1 -ExportDir "D:\Finance\Licensing" -RunAsUser "CONTOSO\svc-licensing"
```

### Na registratie handmatig testen:

```powershell
Start-ScheduledTask -TaskName "... Licentie Overzicht Generator"
```

> **Let op:** Het serviceaccount (`-RunAsUser`) moet leestoegang hebben tot de importmappen van Ingram/Pax8 en schrijftoegang tot de exportmap. Een taak die vóór 2026-10-05 is geregistreerd, start het Python-script zonder `--export-dir` en stopt nu meteen — voer `create_scheduled_task.ps1` opnieuw uit met `-ExportDir`.

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

### Exportmap

Er is niets hardcoded. De exportmap komt, in deze volgorde, uit:

1. `-ExportDir` (`genereer_rapport.ps1`) of `--export-dir` (`genereer_licentie_overzicht.py`)
2. de omgevingsvariabele `LICENSING_EXPORT_DIR`

Zonder een van beide stoppen beide scripts met exitcode 2. Een gesynchroniseerde OneDrive- of SharePoint-map werkt prima als exportmap.

---

## Probleemoplossing

| Fout | Oorzaak | Oplossing |
|---|---|---|
| `No export directory given` | `-ExportDir`/`--export-dir` en `LICENSING_EXPORT_DIR` zijn geen van beide gezet | Geef de map mee, of zet de omgevingsvariabele |
| `Export directory not found` | Verkeerd pad, of een gesynchroniseerde map die nog niet gesynchroniseerd is | Controleer het pad; meld je bij OneDrive aan en wacht tot de synchronisatie klaar is |
| `Geen Excel bestand gevonden in Import\Ingram\` | Ingram-bestand niet geplaatst | Zet precies 1 `.xlsx` in de map `Ingram` |
| `Meerdere bestanden gevonden` | Meer dan 1 bestand in een map | Verwijder of archiveer het extra bestand |
| `Python niet gevonden` | Python staat niet in PATH | Installeer Python en voeg het toe aan PATH, of gebruik het volledige pad in de taak |
| Lege secties in de uitvoer | Klantnaam verschilt tussen Pax8 en Ingram | Controleer de bronbestanden op afsluitende spaties of verschillend hoofdlettergebruik |
