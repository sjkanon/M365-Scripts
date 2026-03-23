# Licensing Report Toolkit

> Author: Sjoerd Kanon

Generates a monthly licensing and Azure cost overview per customer as a formatted Excel file. Combines billing data from **Pax8** (CSV) and **Ingram** (Excel) into one report with a summary tab and a dedicated tab per customer.

---

## Folder structure

```
Licensing/
├── genereer_licentie_overzicht.py   ← Python engine — reads Pax8 + Ingram, writes Excel
├── genereer_rapport.ps1             ← PowerShell launcher — validates env, calls Python
├── genereer_rapport.bat             ← Simple double-click launcher (no validation)
└── create_scheduled_task.ps1        ← Registers a monthly scheduled task (run once)
```

---

## Prerequisites

| Requirement | Version |
|---|---|
| Python | 3.8 or later |
| pandas | `pip install pandas` |
| openpyxl | `pip install openpyxl` |
| OneDrive | Synced and signed in |

Install dependencies:

```bash
pip install pandas openpyxl
```

---

## Input files

Place input files in the correct subfolders under the export directory before running:

```
C:\OneDrive\<Company>\<Company> - Finance - Licenses_facturatie_upload\
├── Import\
│   ├── Ingram\     ← place exactly 1 .xlsx file here (Ingram billing export)
│   └── Pax8\       ← place exactly 1 .csv file here (Pax8 invoice export)
├── Archive\        ← input files are moved here automatically after processing
└── Licentie_Overzicht_YYYY-MM.xlsx   ← output written here
```

> The script fails with a clear error message if: the OneDrive folder is not reachable, no file is found, or more than one file is present in a folder.

---

## Usage

### Option 1 — PowerShell launcher (recommended)

Validates the environment before running. Shows clear error messages if something is missing.

```powershell
.\genereer_rapport.ps1
```

### Option 2 — Double-click

Run `genereer_rapport.bat` directly. No pre-flight checks — relies on the Python script's own error handling.

### Option 3 — CLI with explicit paths

```bash
python genereer_licentie_overzicht.py --ingram "path\to\ingram.xlsx" --pax8 "path\to\pax8.csv"
python genereer_licentie_overzicht.py --ingram "path\to\ingram.xlsx" --pax8 "path\to\pax8.csv" --output "C:\output\rapport.xlsx"
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

## Scheduled task

`create_scheduled_task.ps1` registers a Windows scheduled task that runs the Python script automatically on the **6th of every month at 08:00**.

### Configuration (edit before running)

| Variable | Default | Description |
|---|---|---|
| `$TaskName` | `"... Licentie Overzicht Generator"` | Task name in Task Scheduler |
| `$ScriptPath` | auto (same folder) | Path to `genereer_licentie_overzicht.py` |
| `$RunAsUser` | `"DOMAIN\sa-halo"` | Service account that runs the task — **update for your domain** |

### Run once to register:

```powershell
# Run as Administrator
.\create_scheduled_task.ps1
```

### Test manually after registration:

```powershell
Start-ScheduledTask -TaskName "... Licentie Overzicht Generator"
```

> **Note:** The service account (`$RunAsUser`) must have read access to the Ingram/Pax8 import folders and write access to the OneDrive export folder.

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

### OneDrive path

The base path is hardcoded in both `genereer_rapport.ps1` and `genereer_licentie_overzicht.py`:

```
C:\OneDrive\<Company>\<Company> - Finance - Licenses_facturatie_upload
```

Update `$ExportDir` in `genereer_rapport.ps1` and `EXPORT_DIR` in `genereer_licentie_overzicht.py` to match your OneDrive folder name.

---

## Troubleshooting

| Error | Cause | Fix |
|---|---|---|
| `OneDrive map niet bereikbaar` | OneDrive not synced or not signed in | Sign in to OneDrive and wait for sync |
| `Geen Excel bestand gevonden in Import\Ingram\` | Ingram file not placed | Place exactly 1 `.xlsx` in the `Ingram` folder |
| `Meerdere bestanden gevonden` | More than 1 file in a folder | Remove or archive the extra file |
| `Python niet gevonden` | Python not in PATH | Install Python and add to PATH, or use full path in the task |
| Empty sections in output | Customer name mismatch between Pax8 and Ingram | Check for trailing spaces or different capitalisation in source files |
