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
| OneDrive | Synchronisé et connecté |

Installer les dépendances :

```bash
pip install pandas openpyxl
```

---

## Fichiers d'entrée

Avant l'exécution, placez les fichiers d'entrée dans les bons sous-dossiers du répertoire d'export :

```
C:\OneDrive\<Company>\<Company> - Finance - Licenses_facturatie_upload\
├── Import\
│   ├── Ingram\     ← placez ici exactement 1 fichier .xlsx (export de facturation Ingram)
│   └── Pax8\       ← placez ici exactement 1 fichier .csv (export de factures Pax8)
├── Archive\        ← les fichiers d'entrée sont déplacés ici automatiquement après traitement
└── Licentie_Overzicht_YYYY-MM.xlsx   ← la sortie est écrite ici
```

> Le script s'arrête avec un message d'erreur clair si : le dossier OneDrive n'est pas accessible, aucun fichier n'est trouvé, ou plus d'un fichier se trouve dans un dossier.

---

## Utilisation

### Option 1 — Lanceur PowerShell (recommandé)

Vérifie l'environnement avant l'exécution. Affiche des messages d'erreur clairs s'il manque quelque chose.

```powershell
.\genereer_rapport.ps1
```

### Option 2 — Double-clic

Exécutez directement `genereer_rapport.bat`. Aucune vérification préalable — s'appuie sur la gestion des erreurs propre au script Python.

### Option 3 — Ligne de commande avec chemins explicites

```bash
python genereer_licentie_overzicht.py --ingram "path\to\ingram.xlsx" --pax8 "path\to\pax8.csv"
python genereer_licentie_overzicht.py --ingram "path\to\ingram.xlsx" --pax8 "path\to\pax8.csv" --output "C:\output\rapport.xlsx"
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

`create_scheduled_task.ps1` enregistre une tâche planifiée Windows qui exécute automatiquement le script Python **le 6 de chaque mois à 08:00**.

### Configuration (à modifier avant l'exécution)

| Variable | Par défaut | Description |
|---|---|---|
| `$TaskName` | `"Licensing Report Generator"` | Nom de la tâche dans le Planificateur de tâches |
| `$RunAsUser` | `"$env:USERDOMAIN\sa-halo"` | Compte de service qui exécute la tâche — **à adapter à votre domaine** |
| `$RunDay` | `6` | Jour du mois de l'exécution |
| `$RunTime` | `"08:00"` | Heure de la journée |

Le script détecte automatiquement `python.exe` via le PATH et les emplacements d'installation courants. `$ScriptPath` est résolu automatiquement par rapport au dossier du script.

### À exécuter une fois pour l'enregistrer :

```powershell
# À exécuter en tant qu'Administrateur
.\create_scheduled_task.ps1
```

### Tester manuellement après l'enregistrement :

```powershell
Start-ScheduledTask -TaskName "... Licentie Overzicht Generator"
```

> **Remarque :** le compte de service (`$RunAsUser`) doit disposer d'un accès en lecture aux dossiers d'import Ingram/Pax8 et d'un accès en écriture au dossier d'export OneDrive.

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

### Chemin OneDrive

Le chemin de base est codé en dur à la fois dans `genereer_rapport.ps1` et dans `genereer_licentie_overzicht.py` :

```
C:\OneDrive\<Company>\<Company> - Finance - Licenses_facturatie_upload
```

Modifiez `$ExportDir` dans `genereer_rapport.ps1` et `EXPORT_DIR` dans `genereer_licentie_overzicht.py` pour qu'ils correspondent au nom de votre dossier OneDrive.

---

## Dépannage

| Erreur | Cause | Solution |
|---|---|---|
| `OneDrive map niet bereikbaar` | OneDrive non synchronisé ou non connecté | Connectez-vous à OneDrive et attendez la fin de la synchronisation |
| `Geen Excel bestand gevonden in Import\Ingram\` | Fichier Ingram non déposé | Placez exactement 1 `.xlsx` dans le dossier `Ingram` |
| `Meerdere bestanden gevonden` | Plus d'1 fichier dans un dossier | Supprimez ou archivez le fichier en trop |
| `Python niet gevonden` | Python absent du PATH | Installez Python et ajoutez-le au PATH, ou utilisez le chemin complet dans la tâche |
| Sections vides dans la sortie | Nom de client différent entre Pax8 et Ingram | Vérifiez les espaces en fin de chaîne ou les différences de majuscules dans les fichiers sources |
