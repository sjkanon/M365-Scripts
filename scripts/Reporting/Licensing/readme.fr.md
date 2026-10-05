[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../../readme.fr.md) › [scripts](../../readme.fr.md) › [Reporting](../readme.fr.md) › **Licensing**

# Kit de rapport des licences

> Auteur : Sjoerd Kanon

Génère chaque mois, pour chaque client, un récapitulatif des licences et des coûts Azure sous forme de fichier Excel mis en forme. Combine les données de facturation de **Pax8** (CSV) et d'**Ingram** (Excel) dans un seul rapport, avec un onglet de synthèse et un onglet dédié par client.

---

## Arborescence du dossier

```
Licensing/
├── genereer_licentie_overzicht.py   ← Moteur Python — lit Pax8 + Ingram, écrit l'Excel
├── genereer_rapport.ps1             ← Lanceur PowerShell — vérifie l'environnement, appelle Python
├── genereer_rapport.bat             ← Lanceur simple par double-clic (sans vérification)
└── create_scheduled_task.ps1        ← Enregistre une tâche planifiée mensuelle (à exécuter une fois)
```

---

## Prérequis

| Prérequis | Version |
|---|---|
| Python | 3.8 ou ultérieure |
| pandas | `pip install pandas` |
| openpyxl | `pip install openpyxl` |
| Dossier d'export | Accessible au compte qui exécute le rapport (un dossier OneDrive synchronisé convient) |

Installer les dépendances :

```bash
pip install pandas openpyxl
```

---

## Fichiers d'entrée

Avant l'exécution, placez les fichiers d'entrée dans les bons sous-dossiers du répertoire d'export :

```
<ExportDir>\
├── Import\
│   ├── Ingram\     ← placez ici exactement 1 fichier .xlsx (export de facturation Ingram)
│   └── Pax8\       ← placez ici exactement 1 fichier .csv (export de factures Pax8)
├── Archive\        ← les fichiers d'entrée sont déplacés ici automatiquement après traitement
└── Licentie_Overzicht_YYYY-MM.xlsx   ← la sortie est écrite ici
```

> Le script s'arrête avec un message d'erreur clair si : aucun dossier d'export n'est indiqué, le dossier d'export n'est pas accessible, aucun fichier n'est trouvé, ou plus d'un fichier se trouve dans un dossier.

---

## Utilisation

### Option 1 — Lanceur PowerShell (recommandé)

Vérifie l'environnement avant l'exécution. Affiche des messages d'erreur clairs s'il manque quelque chose.

```powershell
.\genereer_rapport.ps1 -ExportDir "D:\Finance\Licensing"
```

### Option 2 — Double-clic

Exécutez directement `genereer_rapport.bat` — nécessite la variable d'environnement `LICENSING_EXPORT_DIR`. Aucune vérification préalable — s'appuie sur la gestion des erreurs propre au script Python.

### Option 3 — Ligne de commande avec chemins explicites

```bash
python genereer_licentie_overzicht.py --export-dir "D:\Finance\Licensing" --ingram "path\to\ingram.xlsx" --pax8 "path\to\pax8.csv"
python genereer_licentie_overzicht.py --ingram "path\to\ingram.xlsx" --pax8 "path\to\pax8.csv" --export-dir "D:\Finance\Licensing" --output "C:\output\rapport.xlsx"
```

---

## Sortie

Le fichier Excel généré contient :

| Onglet | Contenu |
|---|---|
| `Overzicht` | Synthèse de tous les clients — total achats, total ventes, marge |
| `<Customer name>` | Un onglet par client avec toutes les sections |

### Sections de l'onglet de chaque client

| Section | Source | Contenu |
|---|---|---|
| ☁ Azure (via Pax8) | Pax8 | Consommation Azure regroupée par abonnement → catégorie |
| ☁ Azure (via Ingram) | Ingram | Consommation Azure regroupée par abonnement → catégorie |
| 📋 Licenties (via Pax8) | Pax8 | Licences M365 et autres, regroupées par produit |
| 📋 Licenties (via Ingram) | Ingram | Licences regroupées par produit |
| 📋 Acronis Backup | Pax8 | Affiché sur l'onglet du client final lorsqu'Acronis est facturé via un revendeur |
| **TOTAAL** | — | Total général achats + ventes pour ce client |

Chaque ligne affiche : description, catégorie, quantité, prix d'achat unitaire, prix de vente unitaire, total achats, total ventes.

---

## Tâche planifiée

`create_scheduled_task.ps1` enregistre une tâche planifiée Windows qui exécute automatiquement le script Python par défaut **le 6 de chaque mois à 08:00**.

### Paramètres

| Paramètre | Par défaut | Description |
|---|---|---|
| `-ExportDir` | *(obligatoire)* | Dossier contenant `Import\` et `Archive\`, où le rapport est écrit ; transmis au script Python sous la forme `--export-dir` |
| `-RunAsUser` | *(obligatoire)* | Compte de service qui exécute la tâche, p. ex. `CONTOSO\svc-licensing` |
| `-RunDay` | `6` | Jour du mois de l'exécution (1–28) |
| `-RunTime` | `"08:00"` | Heure de la journée |

Le script détecte automatiquement `python.exe` via le PATH et les emplacements d'installation courants. `$ScriptPath` est résolu automatiquement par rapport au dossier du script.

### À exécuter une fois pour l'enregistrer :

```powershell
# À exécuter en tant qu'Administrateur
.\create_scheduled_task.ps1 -ExportDir "D:\Finance\Licensing" -RunAsUser "CONTOSO\svc-licensing"
```

### Tester manuellement après l'enregistrement :

```powershell
Start-ScheduledTask -TaskName "... Licentie Overzicht Generator"
```

> **Remarque :** le compte de service (`-RunAsUser`) doit disposer d'un accès en lecture aux dossiers d'import Ingram/Pax8 et d'un accès en écriture au dossier d'export. Une tâche enregistrée avant le 2026-10-05 lance le script Python sans `--export-dir` et s'arrête désormais immédiatement — relancez `create_scheduled_task.ps1` avec `-ExportDir`.

---

## Fichier journal

Un fichier journal est écrit dans `Log\licentie_rapport.log`, dans le dossier du script. Chaque exécution ajoute un bloc horodaté avec des entrées `[INFO]` et `[ERROR]`. Consultez ce fichier si le rapport échoue silencieusement via la tâche planifiée.

---

## Personnalisation

### Libellés des abonnements Pax8

Modifiez `PAX8_SUB_LABELS` dans `genereer_licentie_overzicht.py` pour associer les identifiants bruts des abonnements à des noms lisibles :

```python
PAX8_SUB_LABELS = {
    "MY-SUB-001": "My Subscription Label",
}
```

### Alias des clients finaux Acronis

Modifiez `ACRONIS_ENDCUSTOMER_ALIASES` pour associer les noms de clients finaux de Pax8 aux noms utilisés ailleurs dans le rapport :

```python
ACRONIS_ENDCUSTOMER_ALIASES = {
    "Customer Name in Pax8": "Customer Name in Report",
}
```

### Dossier d'export

Rien n'est codé en dur. Le dossier d'export provient, dans cet ordre :

1. de `-ExportDir` (`genereer_rapport.ps1`) ou `--export-dir` (`genereer_licentie_overzicht.py`)
2. de la variable d'environnement `LICENSING_EXPORT_DIR`

Sans l'un ni l'autre, les deux scripts s'arrêtent avec le code de sortie 2. Un dossier OneDrive ou SharePoint synchronisé convient parfaitement comme dossier d'export.

---

## Dépannage

| Erreur | Cause | Solution |
|---|---|---|
| `No export directory given` | Ni `-ExportDir`/`--export-dir` ni `LICENSING_EXPORT_DIR` ne sont définis | Indiquez le dossier, ou définissez la variable d'environnement |
| `Export directory not found` | Chemin erroné, ou dossier synchronisé pas encore à jour | Vérifiez le chemin ; pour OneDrive, connectez-vous et attendez la synchronisation |
| `Geen Excel bestand gevonden in Import\Ingram\` | Fichier Ingram non déposé | Placez exactement 1 `.xlsx` dans le dossier `Ingram` |
| `Meerdere bestanden gevonden` | Plus d'1 fichier dans un dossier | Supprimez ou archivez le fichier en trop |
| `Python niet gevonden` | Python absent du PATH | Installez Python et ajoutez-le au PATH, ou utilisez le chemin complet dans la tâche |
| Sections vides dans la sortie | Nom de client différent entre Pax8 et Ingram | Vérifiez les espaces en fin de chaîne ou les différences de majuscules dans les fichiers sources |
