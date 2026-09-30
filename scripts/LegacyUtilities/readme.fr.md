[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../readme.fr.md) › [scripts](../readme.fr.md) › **LegacyUtilities**

# Legacy Utilities

Équivalents modernisés, dans le style maison de ce dépôt, d'un ensemble de scripts issus d'un
dépôt interne retiré. Aucun d'eux n'est une copie conforme de l'original : plusieurs anciens
scripts ponctuels qui réalisaient de légères variantes de la même tâche ont été regroupés en un
seul script bien paramétré, les modules obsolètes (MSOnline/AzureAD) ont été remplacés par leurs
équivalents Microsoft Graph / Exchange Online, et chaque nom de client, domaine de tenant, nom
d'hôte, mot de passe ou secret codé en dur dans les originaux a été transformé en paramètre.
Aucune de ces données n'a été reprise.

## Dossiers

Ce dossier est organisé par thème, avec un sous-dossier par domaine :

| Dossier | Description |
|--------|----------|
| [`Exchange/`](Exchange/readme.fr.md) | Accès délégué aux boîtes aux lettres, création en masse de boîtes partagées/contacts, synchronisation de contacts, suivi des messages, dédoublonnage de boîtes aux lettres |
| [`Entra/`](Entra/readme.fr.md) | Modifications d'appartenance aux groupes, sauvegarde des stratégies Conditional Access |
| [`Teams/`](Teams/readme.fr.md) | Clonage d'équipes, copie de plans Planner, création en masse d'équipes projet |
| [`Network/`](Network/readme.fr.md) | Montage de partages SMB Azure Files |
| [`Device/`](Device/readme.fr.md) | État par défaut du Verr Num, raccourci Verrouiller la station de travail |
| [`Workspace365/`](Workspace365/readme.fr.md) | Provisionnement/déprovisionnement d'environnements Workspace 365 |

Aucun de ces scripts n'est encore intégré à [`menu.ps1`](../../menu.ps1) ni à la
documentation de premier niveau ; cette intégration fera l'objet d'une passe distincte.

---

## Ce qui a été volontairement laissé de côté

Une grande partie du matériel source a été écartée plutôt que portée : soit parce qu'elle était
déjà entièrement couverte ailleurs dans ce dépôt, soit parce qu'il s'agissait d'une véritable
impasse, soit parce qu'elle n'existait que sous forme de données réelles de clients/tenants qui ne
doivent jamais être reproduites. Consultez le rapport de portage pour le détail complet ; en
résumé :

- **Déjà couvert ailleurs dans ce dépôt** : collecte d'informations Windows Autopilot
  (`scripts/Intune/Get-Autopilot/`), nettoyage général du disque et des journaux
  (`scripts/Device/Invoke-WindowsCleanup.ps1`), suppression des logiciels OEM superflus
  (`scripts/Device/Remove-OemBloatware.ps1`), rapports réseau et mises à jour de firmware
  UniFi (`scripts/Network/UniFi/`), le modèle de connexion aux tenants CSP/GDAP
  (`scripts/Startup/functies.ps1`), les autorisations de dossiers de calendrier
  (`scripts/Exchange/Set-Calendar-rights.ps1`), ainsi qu'un ensemble de scripts ponctuels
  poste de travail/AppDeployment (raccourcis URL sur le bureau, copie de raccourcis du menu
  Démarrer, désinstallation d'Office, règle de pare-feu Teams pour le LAN) déjà généralisés sous
  `scripts/TenantOnboarding/`.
- **Impasses** : un kit complet d'atténuation de PrintNightmare (la CVE du spouleur de 2021)
  construit autour de l'outil obsolète `subinacl.exe`. La vulnérabilité est corrigée depuis
  des années et le contournement n'a plus aucune utilité.
- **Données réelles de clients/tenants, pas du code** : plusieurs anciens scripts et exports
  JSON contenaient des ID de tenant réels, des GUID d'objets groupe/utilisateur, des clés
  d'accès de comptes de stockage, des secrets client d'applications ou des adresses e-mail/domaines
  de clients. Conformément aux règles de traitement des données du projet, rien de tout cela n'a
  été reproduit nulle part (pas même dans les rapports) ; seule la *fonctionnalité* générique
  sous-jacente (par ex. « sauvegarder les stratégies Conditional Access », « monter un partage
  Azure Files ») a été portée, toutes les valeurs identifiantes étant transformées en paramètres.
