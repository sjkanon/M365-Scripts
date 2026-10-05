**English** · [Nederlands](readme.nl.md) · [Français](readme.fr.md)

[M365-Scripts](../../../readme.md) › [scripts](../../readme.md) › [Reporting](../readme.md) › **Licensing**

# Licensing Report Toolkit

> Author: Sjoerd Kanon

Generates a monthly licensing and Azure cost overview per customer as a formatted Excel file. Combines billing data from **Pax8** (CSV) and **Ingram** (Excel) into one report with a summary tab and a dedicated tab per customer.

---

## Scripts

| Script | Description |
|--------|-------------|
| [`genereer_rapport.ps1`](genereer_rapport.ps1) ([docs](#genereer_rapportps1)) | PowerShell launcher — validates the environment, then calls the Python engine (recommended) |
| [`genereer_rapport.bat`](genereer_rapport.bat) ([docs](#genereer_rapportbat)) | Simple double-click launcher, no validation — needs `LICENSING_EXPORT_DIR` |
| [`genereer_licentie_overzicht.py`](genereer_licentie_overzicht.py) ([docs](#genereer_licentie_overzichtpy)) | Python engine — reads the Pax8 CSV and Ingram Excel, writes the per-customer Excel report |
| [`create_scheduled_task.ps1`](create_scheduled_task.ps1) ([docs](#create_scheduled_taskps1)) | Registers a monthly scheduled task that runs the Python engine (run once, as Administrator) |

---

## Prerequisites

| Requirement | Version |
|---|---|
| Python | 3.8 or later |
| pandas | `pip install pandas` |
| openpyxl | `pip install openpyxl` |
| Export directory | Reachable for the account that runs the report (a synced OneDrive folder is fine) |

Install dependencies:

```bash
pip install pandas openpyxl
```

---

## Input files

Place input files in the correct subfolders under the export directory before running:

```
<ExportDir>\
├── Import\
│   ├── Ingram\     ← place exactly 1 .xlsx file here (Ingram billing export)
│   └── Pax8\       ← place exactly 1 .csv file here (Pax8 invoice export)
├── Archive\        ← input files are moved here automatically after processing
└── Licentie_Overzicht_YYYY-MM.xlsx   ← output written here
```

> The script fails with a clear error message if: no export directory is given, the export directory is not reachable, no file is found, or more than one file is present in a folder.

---

## Usage

### genereer_rapport.ps1

Option 1 — PowerShell launcher (recommended). Validates the environment before running. Shows clear error messages if something is missing.

```powershell
.\genereer_rapport.ps1 -ExportDir "D:\Finance\Licensing"
```

### genereer_rapport.bat

Option 2 — double-click. Run `genereer_rapport.bat` directly — requires the `LICENSING_EXPORT_DIR` environment variable. No pre-flight checks — relies on the Python script's own error handling.

### genereer_licentie_overzicht.py

Option 3 — call the Python engine directly with explicit paths:

```bash
python genereer_licentie_overzicht.py --export-dir "D:\Finance\Licensing" --ingram "path\to\ingram.xlsx" --pax8 "path\to\pax8.csv"
python genereer_licentie_overzicht.py --ingram "path\to\ingram.xlsx" --pax8 "path\to\pax8.csv" --export-dir "D:\Finance\Licensing" --output "C:\output\rapport.xlsx"
```

---

## Output

The generated Excel contains:

| Tab | Content |
|---|---|
| `Overzicht` | Summary of all customers — purchase total, sales total, margin |
| `<Customer name>` | One tab per customer with all sections |

### Per-customer tab sections

| Section | Source | Content |
|---|---|---|
| ☁ Azure (via Pax8) | Pax8 | Azure consumption grouped by subscription → category |
| ☁ Azure (via Ingram) | Ingram | Azure consumption grouped by subscription → category |
| 📋 Licenties (via Pax8) | Pax8 | M365 and other licenses grouped by product |
| 📋 Licenties (via Ingram) | Ingram | Licenses grouped by product |
| 📋 Acronis Backup | Pax8 | Shown on the end-customer tab when Acronis is billed via a reseller |
| **TOTAAL** | — | Grand total purchase + sales for this customer |

Each row shows: description, category, quantity, unit purchase price, unit sales price, total purchase, total sales.

---

## create_scheduled_task.ps1

Scheduled task. `create_scheduled_task.ps1` registers a Windows scheduled task that runs the Python script automatically on the **6th of every month at 08:00** by default.

### Parameters

| Parameter | Default | Description |
|---|---|---|
| `-ExportDir` | *(required)* | Folder holding `Import\` and `Archive\`, where the report is written; passed to the Python script as `--export-dir` |
| `-RunAsUser` | *(required)* | Service account that runs the task, e.g. `CONTOSO\svc-licensing` |
| `-RunDay` | `6` | Day of month to run (1–28) |
| `-RunTime` | `"08:00"` | Time of day |

The script auto-detects `python.exe` from PATH and common install locations. `$ScriptPath` is resolved automatically relative to the script folder.

### Run once to register:

```powershell
# Run as Administrator
.\create_scheduled_task.ps1 -ExportDir "D:\Finance\Licensing" -RunAsUser "CONTOSO\svc-licensing"
```

### Test manually after registration:

```powershell
Start-ScheduledTask -TaskName "... Licentie Overzicht Generator"
```

> **Note:** The service account (`-RunAsUser`) must have read access to the Ingram/Pax8 import folders and write access to the export folder. A task registered before 2026-10-05 runs the Python script without `--export-dir` and now stops at once — re-run `create_scheduled_task.ps1` with `-ExportDir`.

---

## Log file

A log file is written to `Log\licentie_rapport.log` in the script folder. Each run appends a timestamped block with `[INFO]` and `[ERROR]` entries. Check this file if the report fails silently via the scheduled task.

---

## Customization

### Pax8 subscription labels

Edit `PAX8_SUB_LABELS` in `genereer_licentie_overzicht.py` to map raw subscription IDs to human-readable names:

```python
PAX8_SUB_LABELS = {
    "MY-SUB-001": "My Subscription Label",
}
```

### Acronis end-customer aliases

Edit `ACRONIS_ENDCUSTOMER_ALIASES` to map Pax8 end-customer names to the names used elsewhere in the report:

```python
ACRONIS_ENDCUSTOMER_ALIASES = {
    "Customer Name in Pax8": "Customer Name in Report",
}
```

### Export directory

Nothing is hardcoded. The export directory is taken from, in order:

1. `-ExportDir` (`genereer_rapport.ps1`) or `--export-dir` (`genereer_licentie_overzicht.py`)
2. the `LICENSING_EXPORT_DIR` environment variable

With neither, both scripts stop with exit code 2. A synced OneDrive or SharePoint folder works fine as the export directory.

---

## Troubleshooting

| Error | Cause | Fix |
|---|---|---|
| `No export directory given` | Neither `-ExportDir`/`--export-dir` nor `LICENSING_EXPORT_DIR` is set | Pass the folder, or set the environment variable |
| `Export directory not found` | Wrong path, or a synced folder that is not synced yet | Check the path; for OneDrive, sign in and wait for sync |
| `Geen Excel bestand gevonden in Import\Ingram\` | Ingram file not placed | Place exactly 1 `.xlsx` in the `Ingram` folder |
| `Meerdere bestanden gevonden` | More than 1 file in a folder | Remove or archive the extra file |
| `Python niet gevonden` | Python not in PATH | Install Python and add to PATH, or use full path in the task |
| Empty sections in output | Customer name mismatch between Pax8 and Ingram | Check for trailing spaces or different capitalisation in source files |
