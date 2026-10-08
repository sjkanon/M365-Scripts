[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../readme.fr.md) › [scripts](../readme.fr.md) › **PatronToolkit**

# Patron Toolkit

Réécritures modernes, basées sur Microsoft Graph / Exchange Online, des fonctionnalités encore utiles de
la boîte à outils retirée [`directorcia/patron`](https://github.com/directorcia/patron) — une
vaste collection d'environ 183 scripts de gestion et de rapport pour les tenants Microsoft 365, construite en partie sur
les modules PowerShell `MSOnline` et `AzureAD`, aujourd'hui retirés.

Aucun code ci-dessous n'est copié de ce projet — il a inspiré la *liste des fonctionnalités*
(quoi vérifier, à quoi ressemble un bon rapport), mais chaque script ici est une
implémentation nouvelle dans le style maison de ce dépôt : essai à blanc par défaut pour toute modification,
export CSV, réutilisation de la connexion compatible GDAP/`-TenantId`, et aucun identifiant ni
donnée de tenant codés en dur. Comme le projet source fournissait un petit script par contrôle avec beaucoup
de recoupements, cette boîte à outils regroupe environ 70 de ces scripts à usage unique en 13
scripts bien paramétrés, organisés par fonctionnalité plutôt que portés à l'identique. Voir la section
`.NOTES` de chaque script pour savoir exactement quels scripts d'origine il remplace.

Chaque script se connecte via [`Connect-M365.ps1`](../Startup/readme.fr.md#connect-m365ps1) :
Microsoft Graph partout où Graph a une API, **en délégué par défaut** (vous vous connectez en
tant qu'administrateur ; code d'appareil et client GDAP depuis `load.config.ps1`), app-only
avec `-ClientId` + `-CertificateThumbprint` ou `-AppOnly`. Exchange Online, Teams PowerShell et
PnP ne servent qu'au travail que Graph ne sait pas faire — les remarques de chaque script
précisent lequel et pourquoi.

---

## Dossiers

| Dossier | Description |
|--------|-------------|
| [`Entra/`](Entra/readme.fr.md) | Rapports d'inscription MFA/SSPR, sauvegarde des stratégies Conditional Access |
| [`Security/`](Security/readme.fr.md) | Audit des consentements d'applications, règles de boîte de réception suspectes, alertes de sécurité, posture de sécurité de la messagerie, journalisation d'audit, validation SPF/DMARC |
| [`Intune/`](Intune/readme.fr.md) | Rapports d'affectation des stratégies Intune, inventaire des appareils Autopilot |
| [`Exchange/`](Exchange/readme.fr.md) | Suivi des messages / diagnostics du flux de messagerie |
| [`SharePoint/`](SharePoint/readme.fr.md) | Configuration du partage SharePoint Online et audit des utilisateurs externes |
| [`Teams/`](Teams/readme.fr.md) | Gouvernance du tenant Teams et rapports d'inventaire |

---

## Ce qui a été ignoré, et pourquoi

Les quelque 183 scripts du projet source ont été réduits à ces 13 pour plusieurs raisons :

- **Couverture native de la plateforme** — la grande famille de scripts `endpoint-*-set.ps1` /
  `intune-*comp-set.ps1` / `intune-*ap-set.ps1` (bases de sécurité Intune,
  stratégies de conformité, stratégies de protection des applications pour Windows/iOS/Android/macOS) code en dur
  les paramètres très tranchés d'un MSP particulier. Les modèles Intune Security Baseline et de
  stratégies de conformité de Microsoft dans le centre d'administration couvrent désormais cela nativement et sont tenus
  à jour par Microsoft — scripter une base de référence figée de 2020 n'apporte aucun avantage.
- **Déjà couvert dans ce dépôt** — les rapports de base sur les licences/utilisateurs/groupes
  (`o365-sku-audit-csv.ps1`, `graph-sku-get.ps1`, `o365-NoSPO-ADAcct*`), l'import de la base de référence CA
  (`ca-policy-import.ps1` — voir `scripts/Entra/Import-ConditionalAccessBaseline.ps1`) et
  la détection de dérive de configuration Intune (`endpoint-policy-get.ps1` — voir
  `scripts/Intune/Compare-IntuneConfig.ps1`) existent déjà et sont activement maintenus.
- **Modules retirés/hérités sans équivalent moderne sûr pour cette fonctionnalité précise
  et limitée** — `o365-add-domain.ps1` (MSOnline + configuration d'une zone Azure DNS codée en dur),
  le flux d'identifiants d'origine de `graph-usrreg-read.ps1` via un XML local, `mcas-*.ps1` (le modèle
  d'authentification distinct par jeton de portail de Defender for Cloud Apps, de plus en plus intégré au portail
  Defender unifié) — la *fonctionnalité* sous-jacente des deux derniers a été conservée mais réécrite
  (rapport MFA ; audit des consentements d'applications), le reste a été abandonné.
- **Non portable entre tenants / peu utile pour un MSP qui gère de nombreux tenants clients** —
  rapports de rapprochement avec l'AD local (`o365-NoSPO-ADAcct*`, `o365-SPO-NoADAcct*`), recherche
  WHOIS (`o365-whois-get.ps1`), export générique d'enregistrements DNS (`o365-dns-get.ps1` — réduit
  au contrôle SPF/DMARC réellement exploitable), et la plomberie locale Hyper-V/menu/cache d'identifiants
  (`hyperv-*.ps1`, `start.ps1`, `*-connect.ps1`, `*-creds-save.ps1`).
- **Hors sujet** — `reclaimwin10.ps1` et `win10-bp-get.ps1` sont de gros scripts de registre/GPO
  sur la machine locale (l'un est un script tiers embarqué, même pas du code propre à patron)
  qui agissent sur un seul poste Windows 10, pas sur la gestion M365 à l'échelle du tenant.
- **Modifications en masse à haut risque volontairement laissées hors périmètre** — l'import/la suppression en masse
  de stratégies Conditional Access, l'ajout/le retrait en masse d'affectations de stratégies/applications Intune, et
  l'import/la suppression/la réaffectation en masse d'appareils Autopilot avaient tous des fonctionnalités de *rapport* que nous avons conservées (voir
  `Export-ConditionalAccessPolicies.ps1`, `Get-IntunePolicyAssignments.ps1`,
  `Get-AutopilotDevices.ps1`), mais nous n'avons pas porté leurs équivalents qui modifient — ce sont
  des opérations propres à chaque tenant, au rayon d'impact élevé, qu'il vaut mieux examiner une par une dans le centre
  d'administration plutôt que de les scripter de façon générique.

Le détail exact des scripts source que chaque rapport remplace se trouve dans la section
`.NOTES` de ce script.
