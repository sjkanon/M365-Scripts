[English](readme.md) · [Nederlands](readme.nl.md) · **Français**

[M365-Scripts](../../readme.fr.md) › [scripts](../readme.fr.md) › **Linux**

# Scripts Linux

Scripts pour serveurs Linux : Debian/Ubuntu, y compris les serveurs qui exécutent 3CX Phone System. Ce sont des scripts **bash**, pas PowerShell : un serveur Linux comme une appliance 3CX n'a généralement pas `pwsh`. Ils s'exécutent en root sur le serveur lui-même et ne font pas partie de [`menu.ps1`](../../menu.ps1).

---

## Scripts

| Script | Description |
|--------|-------------|
| [`Invoke-LinuxCleanup.sh`](Invoke-LinuxCleanup.sh) ([docs](#invoke-linuxcleanupsh)) | Analyser et libérer l'espace disque d'un serveur Debian/Ubuntu : paquets, journal, journaux, fichiers temporaires, caches utilisateur, Docker en option, ainsi que les journaux et sauvegardes 3CX lorsque 3CX est installé |

---

### Invoke-LinuxCleanup.sh

L'équivalent Linux de [`Invoke-WindowsCleanup.ps1`](../Device/readme.fr.md#invoke-windowscleanupps1). Sans argument, c'est un essai à blanc : il indique par catégorie combien d'espace peut être libéré. `--apply` effectue le nettoyage, et `--check-only` est la même analyse pour la supervision, avec un code de sortie. Tout ce qui est supprimé selon l'âge doit être plus ancien que `--days` (14 par défaut). Rien n'est mis à jour : sur un serveur 3CX, 3CX met à jour lui-même son logiciel et le système via sa propre console.

**Ce qui est nettoyé**

| Catégorie | Détails |
|-----------|---------|
| Paquets | Cache APT (`apt-get clean`) ; paquets devenus inutiles, anciens noyaux compris (`apt-get autoremove --purge`), la liste étant affichée lors de l'essai à blanc ; configuration résiduelle des paquets supprimés (état dpkg `rc`) |
| Snap | Révisions snap désactivées, lorsque snap est installé |
| Journal | Journal systemd, réduit à `--days` et `--journal-size` |
| Journaux système | Journaux ayant subi une rotation dans `/var/log` (`*.1`, `*.gz`, `*.xz`, `*.old`, ...), PostgreSQL et nginx compris — les journaux actifs sont conservés |
| Vidages après plantage | `/var/lib/systemd/coredump`, `/var/crash` |
| Temp | `/tmp`, `/var/tmp` |
| Cache utilisateur | `~/.cache`, `~/.npm/_cacache` et `~/.local/share/Trash` de root et de chaque dossier personnel sous `/home` |
| Docker | Uniquement avec `--docker` : `docker system prune -f` (conteneurs arrêtés, réseaux inutilisés, images orphelines, cache de build). Sans cette option, l'espace récupérable indiqué par Docker est affiché |
| Journaux 3CX | `<data-dir>/Logs` et les journaux nginx de 3CX (`/var/lib/3cxpbx/Bin/nginx/logs`) — uniquement lorsque 3CX est installé |
| Sauvegardes 3CX | `*.zip` dans `<data-dir>/Backups` au-delà des N plus récentes — uniquement avec `--keep-backups N` |

Signalé, jamais supprimé : les enregistrements d'appels (avec leur taille), les sauvegardes 3CX sans `--keep-backups`, les plus gros dossiers du dossier de données 3CX, et les 10 plus gros fichiers de plus de 500 Mo sur le serveur.

**Paramètres**

| Paramètre | Description |
|-----------|-------------|
| `--apply` | Effectuer le nettoyage (par défaut : essai à blanc uniquement) |
| `--check-only` | Analyser seulement, ne rien modifier ; code de sortie `2` lorsqu'il y a de l'espace à libérer, `0` sinon. Incompatible avec `--apply` |
| `--days N` | Ne supprimer que les fichiers de plus de N jours (par défaut : `14`). Le journal, `/tmp` et les caches utilisateur conservent au moins un jour |
| `--journal-size SIZE` | Réduire le journal systemd à SIZE au maximum, p. ex. `200M` ou `1G` (par défaut : `200M`) |
| `--docker` | Exécuter aussi `docker system prune -f`. Les volumes ne sont jamais supprimés |
| `--keep-backups N` | Supprimer les sauvegardes 3CX au-delà des N plus récentes (N ≥ 1) |
| `--data-dir PATH` | Dossier de données 3CX (par défaut : `/var/lib/3cxpbx/Instance1/Data`) |
| `--skip-apt` | Ignorer tout ce qui concerne APT/dpkg : cache, autoremove, configuration résiduelle |
| `--skip-autoremove` | Ignorer uniquement `apt-get autoremove` |
| `--skip-journal` | Ignorer le journal systemd |
| `--skip-user-cache` | Ignorer les caches et corbeilles des utilisateurs |
| `--skip-3cx` | Ignorer les sections 3CX |
| `-h`, `--help` | Afficher l'aide |

**Exemples**

```bash
# Essai à blanc — ce qui peut être libéré
sudo ./Invoke-LinuxCleanup.sh

# Supervision / RMM : code de sortie 2 s'il y a du travail, 0 si tout est propre
sudo ./Invoke-LinuxCleanup.sh --check-only

# Nettoyer
sudo ./Invoke-LinuxCleanup.sh --apply

# Nettoyer davantage : fichiers de plus de 7 jours, conserver les 5 sauvegardes 3CX les plus récentes
sudo ./Invoke-LinuxCleanup.sh --apply --days 7 --keep-backups 5

# Nettoyer aussi Docker
sudo ./Invoke-LinuxCleanup.sh --apply --docker

# Directement depuis GitHub, sans copier le fichier (dépôt public uniquement)
curl -fsSL https://raw.githubusercontent.com/sjkanon/M365-Scripts/main/scripts/Linux/Invoke-LinuxCleanup.sh | sudo bash -s -- --check-only
```

**Codes de sortie**

| Code | Signification |
|------|---------------|
| `0` | Terminé, ou rien à faire |
| `1` | Argument incorrect, ou pas exécuté en root |
| `2` | `--check-only` uniquement : de l'espace peut être libéré |

**Remarques**

- Nécessite root (`sudo`). Fonctionne sur Debian et Ubuntu ; sur une distribution sans `apt-get`, la section des paquets est ignorée et le reste s'exécute normalement.
- 3CX est détecté par le paquet `3cxpbx` ou par le dossier de données ; sans 3CX, les sections 3CX sont omises. `--skip-3cx` les omet aussi sur un serveur 3CX.
- **Sécurité :** si `apt`/`dpkg` est déjà en cours (une mise à jour du système ou de 3CX), la section des paquets est ignorée pour cette exécution. `autoremove` est refusé lorsque sa liste contient un paquet 3CX. Dans `/tmp`, `systemd-private-*` (appartenant aux services en cours) et le fichier de verrou du socket PostgreSQL `.s.PGSQL.*` ne sont pas touchés.
- Les enregistrements d'appels ne sont jamais supprimés : 3CX dispose de son propre paramètre de conservation dans la console d'administration. Il en va de même pour la base de données, la messagerie vocale, les messages d'accueil et la configuration.
- « Total freed (measured) » compare l'espace utilisé de `/`, `/var`, `/tmp`, `/home` et du dossier de données 3CX avant et après ; « reported » additionne les catégories. Le journal et Docker ne figurent pas dans l'estimation de l'essai à blanc, car on ne peut pas savoir à l'avance ce qu'un vacuum ou un prune libérera.
- Un fichier journal des dossiers 3CX qu'un service garde encore ouvert, mais dans lequel rien n'a été écrit depuis `--days` jours, est également supprimé. L'espace n'est libéré qu'au redémarrage de ce service.
- Les chemins 3CX (`/var/lib/3cxpbx/Instance1/Data`, `Logs`, `Backups`, `Recordings`, `Bin/nginx/logs`) suivent l'organisation Linux de 3CX ; utilisez `--data-dir` si une installation diffère.
- Pas encore vérifié sur un vrai serveur 3CX. Testé sur Debian 12 (bookworm) avec une arborescence 3CX reconstituée : essai à blanc, `--check-only`, `--apply` et un second `--apply`.
