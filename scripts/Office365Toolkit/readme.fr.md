[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../readme.fr.md) › [scripts](../readme.fr.md) › **Office365Toolkit**

# Office365Toolkit

Réécritures modernes, basées sur Microsoft Graph / Exchange Online, des fonctionnalités encore utiles de
la boîte à outils PowerShell retirée [`directorcia/Office365`](https://github.com/directorcia/Office365)
(CIAOPS) — une collection bien connue de scripts d'administration Office 365,
construite à l'origine sur les modules `MSOnline` / `AzureAD`, aujourd'hui retirés, et sur divers
scripts propres à un tenant avec des valeurs codées en dur.

Le projet source (~58 scripts) a été examiné fonctionnalité par fonctionnalité plutôt que
porté à l'identique : les scripts de rapport à usage unique en double ont été regroupés,
les scripts d'aide à la connexion ont été supprimés (ce dépôt se connecte déjà automatiquement et réutilise les sessions),
tout ce qui nécessitait les modules retirés `MSOnline`/`AzureAD` a été réécrit avec
`Microsoft.Graph.*` ou `ExchangeOnlineManagement`, et les fonctionnalités déjà couvertes
ailleurs dans ce dépôt (audits de taille/permissions des boîtes aux lettres, transfert externe,
stockage SharePoint, rapports de licences, Conditional Access, etc. — voir
[`scripts/Entra/`](../Entra/readme.fr.md), [`scripts/Exchange/`](../Exchange/readme.fr.md),
[`scripts/Reporting/`](../Reporting/readme.fr.md)) ont été ignorées. Tout le code ici est une
implémentation nouvelle, dans le style propre à ce dépôt, et non une copie du projet source.

---

## Dossiers

| Dossier | Description |
|--------|-------------|
| [`Security/`](Security/readme.fr.md) | Rapports Secure Score, nettoyage des consentements d'applications d'entreprise, verrouillage de la connexion aux boîtes aux lettres partagées, base de référence EOP anti-spam/anti-malware |
| [`Exchange/`](Exchange/readme.fr.md) | Base de référence d'hygiène des boîtes aux lettres, risque de transfert via les règles de boîte de réception, compléments de boîte aux lettres, recherche dans l'Unified Audit Log (Graph) |
| [`Intune/`](Intune/readme.fr.md) | Inventaire à l'échelle du tenant des stratégies Intune/Endpoint Manager |

---

## Fonctionnalités ignorées (et pourquoi)

**Déjà couvertes ailleurs dans ce dépôt :**
- Audits de taille/permissions/calendrier/DKIM des boîtes aux lettres, transfert externe via le paramètre
  de boîte aux lettres, utilisation du stockage SharePoint, rapports de licences, import de la
  base de référence Conditional Access, codes TAP — tous disposent déjà de scripts Graph/EXO sous `scripts/Entra/`,
  `scripts/Exchange/` et `scripts/Reporting/`.
- Recherche du nom convivial des SKU de licence (`o365-skus.ps1`) — une table de hachage statique de
  noms de SKU obsolètes ; remplacée par `scripts/Entra/Get-M365UserLicenses.ps1`.

**Aides de connexion/infrastructure, pas des fonctionnalités** (`*-connect*.ps1`,
`graph-connect.ps1`, `msgraph-connect.ps1`, `Intune-connect.ps1`, `az-connect*.ps1`,
`o365-setup.ps1`, `o365-update.ps1`, `o365-getrepo.ps1`, `save-cred-file.ps1`,
`c.ps1`, `r.ps1`, `sc-config.ps1`, `text-colour.ps1`) : ce dépôt se connecte de la même façon dans chaque script via
[`scripts/Startup/Connect-M365.ps1`](../Startup/readme.fr.md#connect-m365ps1) (Graph d'abord,
délégué par défaut, app-only sur demande, compatible GDAP, avec réutilisation des sessions
existantes), les scripts de connexion autonomes n'apportent donc rien. `o365-setup.ps1` codait en outre en dur le chemin OneDrive d'un vrai client et
installe les modules retirés `MSOnline`/`AzureAD` ; `save-cred-file.ps1` stocke
des identifiants dans un fichier XML local — tous deux hors périmètre selon les règles de sécurité de ce projet.

**Hors du périmètre du tenant M365** (diagnostics locaux de l'appareil/du réseau, pas Graph/EXO) :
`win10-asr-get.ps1`, `win10-audit-get.ps1`, `win10-def-get.ps1` (vérifications locales de Windows
Defender/de la stratégie d'audit — au niveau de l'appareil, pas du tenant), `Cleanup AzureAD
device registration.ps1` (nettoyage local du registre), `sec-test.ps1` (menu local de simulation EICAR/
malware), `ipget.ps1`/`ipinf.ps1` (recherches génériques de géolocalisation IP,
dont une avec une clé d'API personnelle codée en dur — sans rapport avec M365).

**Propres à un fournisseur/tenant, non réutilisables de façon générique :** `sc-config.ps1` (codé en dur
pour la configuration Teams Direct Routing d'un opérateur RTC australien).

**Obsolètes ou trop spécifiques, peu rentables à reconstruire :**
- `o365-addin-deploy.ps1` (Centralized Deployment des compléments Outlook) — Microsoft
  oriente les administrateurs vers l'interface Integrated Apps du centre d'administration ; il n'existe
  pas encore d'applet de commande Graph/EXO équivalente.
- `o365-mcas-api.ps1` (API Cloud App Security) et `endpoint-api-svbm.ps1`
  (API de vulnérabilités de Defender for Endpoint) — tous deux codent en dur des variables fictives
  d'URI/jeton/secret et ciblent des API largement remplacées par le portail
  Defender unifié ; ne s'intègrent pas proprement dans Graph/EXO.
- `o365-atp-timer.ps1` (mesure ponctuelle de la latence d'analyse ATP) — diagnostic
  de niche, pas une fonctionnalité de gestion continue.
- `az-sentinel-ruleget.ps1` (rapport des règles d'analyse Azure Sentinel) — périmètre Azure
  Resource Manager (`Az.SecurityInsights`), pas une fonctionnalité M365/Graph/EXO.
- Administrateurs de collections de sites SharePoint, liste des utilisateurs externes et paramètres
  de partage (`o365-spo-admins.ps1`, `o365-spo-extusr.ps1`, `o365-spo-getsharing.ps1`)
  — pas exposés proprement via Microsoft Graph sans SharePoint Online Management
  Shell ou PnP.PowerShell, qui sortent du périmètre des modules Graph/EXO de ce projet ;
  `o365-spo-getusage.ps1` (utilisation du stockage) est déjà couvert par
  `scripts/Reporting/Get-SharePointStorageReport.ps1`.
