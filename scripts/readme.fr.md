[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../readme.fr.md) › **scripts**

# scripts/

Tout l'outillage PowerShell de ce dépôt, regroupé par workload. Lancez tout depuis la racine via [`.\menu.ps1`](../menu.ps1) (ou [`.\load.ps1`](../load.ps1) lors de la première exécution) — consultez le [readme racine](../readme.fr.md) pour la référence complète du menu et les étapes de prise en main.

---

## Vous cherchez un script précis ?

[**INDEX.md**](INDEX.md) les répertorie tous de A à Z sur une seule page — script, dossier et fonction — pour que vous puissiez faire Ctrl-F au lieu de deviner dans quel dossier il se trouve. Cette liste est générée à partir des scripts eux-mêmes par [`Startup/Update-ScriptIndex.ps1`](Startup/Update-ScriptIndex.ps1) ; relancez-le après avoir ajouté, renommé ou supprimé un script.

Le tableau ci-dessous fonctionne dans l'autre sens : à quoi *sert* chaque catégorie.

---

## Dossiers

| Dossier | Description |
|--------|-------------|
| [`ActiveDirectory/`](ActiveDirectory/readme.fr.md) | Supervision d'AD DS on-premise (surveillance des verrouillages de comptes) — cible directement un DC/serveur de fichiers, pas Entra ID |
| [`Azure/`](Azure/readme.fr.md) | Gestion des VM Azure IaaS (conversion du contrôleur de disque) — cible directement Azure via `Az`, pas le tenant M365 |
| [`Entra/`](Entra/readme.fr.md) | Cycle de vie des utilisateurs, attribution des responsables, rapports de licences, baseline Conditional Access, fenêtres CA temporaires, codes TAP, audit des groupes M365 (Microsoft Graph) |
| [`Exchange/`](Exchange/readme.fr.md) | Migration/autorisations de calendriers, listes de distribution, audits des boîtes aux lettres/calendriers/DKIM/transferts |
| [`Graph/`](Graph/readme.fr.md) | Gestion des autorisations d'application Microsoft Graph |
| [`Intune/`](Intune/readme.fr.md) | Inscription Autopilot, mise à jour de la stratégie de conformité iOS, déploiement du fond d'écran/écran de verrouillage de l'entreprise |
| [`SharePoint/`](SharePoint/readme.fr.md) | Opérations sur le contenu SharePoint Online / OneDrive — restauration de la corbeille par site ou à l'échelle du tenant (PnP PowerShell, inscription d'application automatique), et où est passé un fichier : renommé, déplacé ou supprimé (journal d'audit) |
| [`Reporting/`](Reporting/readme.fr.md) | Rapport de dernière connexion des ordinateurs, rapport de stockage SharePoint, rapport mensuel des licences |
| [`Device/`](Device/readme.fr.md) | Maintenance des postes Windows — activation, nettoyage, fichiers temporaires, synchronisation de l'heure, audio, diagnostic OpenVPN, disque temporaire + fichier d'échange Azure/AVD |
| [`Linux/`](Linux/readme.fr.md) | Serveurs Linux (Debian/Ubuntu, 3CX Phone System) — nettoyage du disque en bash : paquets, journal, journaux, fichiers temporaires, caches utilisateur, Docker, journaux et sauvegardes 3CX |
| [`Network/`](Network/readme.fr.md) | Vérification de ports TCP, diagnostic d'authentification/réseau, tests de charge des E/S fichiers |
| [`RDS/`](RDS/readme.fr.md) | Diagnostic des connexions RDP / RD Web Access, supervision des sessions en direct, diagnostic et réduction des disques de profil FSLogix |
| [`SMTP/`](SMTP/readme.fr.md) | Tests de connectivité du relais SMTP (ponctuels et récurrents) |
| [`Deployment/`](Deployment/readme.fr.md) | Kit USB pour l'installation de Windows et l'inscription Autopilot pendant l'OOBE |
| [`DNS/`](DNS/readme.fr.md) | Résolution et import d'enregistrements DNS dans des zones DNS intégrées à AD |
| [`SAS/`](SAS/readme.fr.md) | Supervision des erreurs des jobs batch SAS avec intégration Zabbix |
| [`Teams/`](Teams/readme.fr.md) | Export et archivage Microsoft Teams / SharePoint |
| [`Startup/`](Startup/readme.fr.md) | Bibliothèque de fonctions M365 `functies.ps1` + amorçage des modules + vérificateur de syntaxe, chargés (dot-sourced) par le menu |
| [`Custom Scripts/`](Custom%20Scripts/readme.fr.md) | Scripts liés à leur chemin — déploiement du thème Office (l'URL de téléchargement vers ce chemin du dépôt est codée en dur) |
| [`TenantOnboarding/`](TenantOnboarding/readme.fr.md) | Provisionnement de nouveaux tenants, rapports multi-tenant/GDAP, déploiement d'applications, configuration des appareils, gestion OneDrive, gestion des utilisateurs — modernisé à partir d'un kit interne de configuration de tenants retiré |
| [`Office365Toolkit/`](Office365Toolkit/readme.fr.md) | Réécritures Security/Exchange/Intune des fonctionnalités encore utiles du kit retiré `directorcia/Office365` (CIAOPS) |
| [`PatronToolkit/`](PatronToolkit/readme.fr.md) | Réécritures Entra/Exchange/Intune/Security/SharePoint/Teams des fonctionnalités encore utiles du kit retiré `directorcia/patron` |
| [`LegacyUtilities/`](LegacyUtilities/readme.fr.md) | Scripts divers modernisés (Exchange, Entra, Teams, Network, Device, Workspace 365) issus de petits outils variés du kit interne retiré |
