[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../readme.nl.md) › [scripts](../readme.nl.md) › **Linux**

# Linux-scripts

Scripts voor Linux-servers: Debian/Ubuntu, ook servers waarop 3CX Phone System draait. Dit zijn **bash**-scripts, geen PowerShell: een Linux-server zoals een 3CX-appliance heeft meestal geen `pwsh`. Ze draaien als root op de server zelf en maken geen deel uit van [`menu.ps1`](../../menu.ps1).

---

## Scripts

| Script | Omschrijving |
|--------|--------------|
| [`Invoke-LinuxCleanup.sh`](Invoke-LinuxCleanup.sh) ([docs](#invoke-linuxcleanupsh)) | Schijfruimte scannen en vrijmaken op een Debian/Ubuntu-server: pakketten, journal, logs, temp, gebruikerscaches, optioneel Docker, plus 3CX-logs en -back-ups als 3CX geïnstalleerd is |

---

### Invoke-LinuxCleanup.sh

De Linux-tegenhanger van [`Invoke-WindowsCleanup.ps1`](../Device/readme.nl.md#invoke-windowscleanupps1). Zonder argumenten is het een proefdraai: per categorie ziet u hoeveel er vrij kan komen. `--apply` voert de opschoning uit, en `--check-only` is dezelfde scan voor monitoring, met een exitcode. Alles wat op leeftijd wordt verwijderd, moet ouder zijn dan `--days` (standaard 14). Er wordt niets bijgewerkt: op een 3CX-server werkt 3CX zichzelf en het OS bij via de eigen console.

**Wat het opruimt**

| Categorie | Details |
|-----------|---------|
| Pakketten | APT-cache (`apt-get clean`); pakketten die niemand meer nodig heeft, oude kernels inbegrepen (`apt-get autoremove --purge`), met de lijst in de proefdraai; achtergebleven configuratie van verwijderde pakketten (dpkg-status `rc`) |
| Snap | Uitgeschakelde snap-revisies, als snap geïnstalleerd is |
| Journal | systemd-journal, teruggebracht tot `--days` en `--journal-size`. De proefdraai rekent uit wat die vacuum vrijmaakt, zoals journald het beslist: alleen gearchiveerde bestanden, de oudste eerst, op de tijd in de bestandsnaam |
| Systeemlogs | Geroteerde logs in `/var/log` (`*.1`, `*.gz`, `*.xz`, `*.old`, ...), ook die van PostgreSQL en nginx — actieve logs blijven staan |
| Crashdumps | `/var/lib/systemd/coredump`, `/var/crash` |
| Temp | `/tmp`, `/var/tmp` |
| Gebruikerscache | `~/.cache`, `~/.npm/_cacache` en `~/.local/share/Trash` van root en elke home onder `/home` |
| Docker | Alleen met `--docker`: `docker system prune -f` (gestopte containers, ongebruikte netwerken, dangling images, buildcache). Zonder die optie wordt Dockers eigen terug te winnen ruimte getoond |
| 3CX-logs | `<data-dir>/Logs`, de nginx-logs van 3CX (`/var/lib/3cxpbx/Bin/nginx/logs`) en `/var/lib/3cxpbx/Data/Logs`, achtergebleven op servers die zijn bijgewerkt vanaf de indeling van vóór `Instance1` — alleen als 3CX geïnstalleerd is |
| 3CX-back-ups | Elke `*.zip` in `<data-dir>/Backups` en de oude `/var/lib/3cxpbx/Data/Backups` wordt met datum en grootte getoond, nieuwste eerst, als één lijst. Met `--keep-backups N` wordt alles voorbij de nieuwste N verwijderd |

Gerapporteerd, nooit verwijderd: gespreksopnames (met hun grootte), 3CX-back-ups zonder `--keep-backups`, de grootste mappen in de 3CX-datamap, en de 10 grootste bestanden boven 500 MB op de server.

**Parameters**

| Parameter | Omschrijving |
|-----------|--------------|
| `--apply` | Voer de opschoning uit (standaard: alleen proefdraai) |
| `--check-only` | Alleen scannen, niets wijzigen; exitcode `2` als er iets vrij kan komen, `0` als dat niet zo is. Niet te combineren met `--apply` |
| `--days N` | Alleen bestanden verwijderen die ouder zijn dan N dagen (standaard: `14`). Journal, `/tmp` en gebruikerscaches houden minstens één dag |
| `--journal-size SIZE` | De systemd-journal verkleinen tot hoogstens SIZE, bijv. `200M` of `1G` (standaard: `200M`) |
| `--docker` | Ook `docker system prune -f` uitvoeren. Volumes worden nooit opgeruimd |
| `--keep-backups N` | 3CX-back-ups voorbij de nieuwste N verwijderen (N ≥ 1) |
| `--data-dir PATH` | 3CX-datamap (standaard: `/var/lib/3cxpbx/Instance1/Data`) |
| `--skip-apt` | Alles van APT/dpkg overslaan: cache, autoremove, achtergebleven configuratie |
| `--skip-autoremove` | Alleen `apt-get autoremove` overslaan |
| `--skip-journal` | De systemd-journal overslaan |
| `--skip-user-cache` | De gebruikerscaches en prullenbakken overslaan |
| `--skip-3cx` | De 3CX-onderdelen overslaan |
| `-h`, `--help` | De help tonen |

**Voorbeelden**

```bash
# Proefdraai — wat kan er vrij komen
sudo ./Invoke-LinuxCleanup.sh

# Monitoring / RMM: exitcode 2 als er werk is, 0 als alles schoon is
sudo ./Invoke-LinuxCleanup.sh --check-only

# Opruimen
sudo ./Invoke-LinuxCleanup.sh --apply

# Grondiger opruimen: bestanden ouder dan 7 dagen, de 5 nieuwste 3CX-back-ups houden
sudo ./Invoke-LinuxCleanup.sh --apply --days 7 --keep-backups 5

# Ook Docker opruimen
sudo ./Invoke-LinuxCleanup.sh --apply --docker

# Rechtstreeks van GitHub, zonder het bestand te kopiëren (alleen bij een openbare repository)
curl -fsSL https://raw.githubusercontent.com/sjkanon/M365-Scripts/main/scripts/Linux/Invoke-LinuxCleanup.sh | sudo bash -s -- --check-only
```

**Exitcodes**

| Code | Betekenis |
|------|-----------|
| `0` | Klaar, of niets te doen |
| `1` | Verkeerd argument, of niet als root uitgevoerd |
| `2` | Alleen `--check-only`: er kan iets vrij komen |

**Opmerkingen**

- Vereist root (`sudo`). Werkt op Debian en Ubuntu; op een distributie zonder `apt-get` wordt het pakketonderdeel overgeslagen en draait de rest gewoon.
- 3CX wordt herkend aan het pakket `3cxpbx` of aan de datamap; zonder 3CX vallen de 3CX-onderdelen weg. `--skip-3cx` laat ze ook op een 3CX-server weg.
- **Veiligheid:** draait `apt`/`dpkg` al (een OS- of 3CX-update), dan wordt het pakketonderdeel die run overgeslagen. Dat wordt afgelezen aan de eigen lock van dpkg (`/proc/locks`), niet aan procesnamen: `unattended-upgrades` houdt altijd een proces `unattended-upgr` draaiende, waardoor de eerste versie de pakketten bij elke run oversloeg. `autoremove` wordt geweigerd als de lijst een 3CX-pakket bevat. In `/tmp` blijven `systemd-private-*` (van draaiende services) en het socket-lockbestand van PostgreSQL `.s.PGSQL.*` met rust gelaten.
- Gespreksopnames worden nooit verwijderd: daarvoor heeft 3CX een eigen bewaarinstelling in de beheerconsole. Hetzelfde geldt voor de database, voicemail, prompts en configuratie.
- "Total freed (measured)" vergelijkt de gebruikte ruimte van `/`, `/var`, `/tmp`, `/home` en de 3CX-datamap voor en na; "reported" telt de categorieën op. Docker zit niet in de schatting van de proefdraai, omdat vooraf niet te zeggen is wat een prune vrijmaakt.
- Een logbestand in de 3CX-logmappen dat een service nog open heeft, maar waar al `--days` dagen niet in geschreven is, wordt ook verwijderd. De ruimte komt pas vrij als die service herstart.
- De 3CX-paden (`/var/lib/3cxpbx/Instance1/Data`, `Logs`, `Backups`, `Recordings`, `Bin/nginx/logs`) volgen de Linux-indeling van 3CX; gebruik `--data-dir` als een installatie afwijkt.
- Als proefdraai uitgevoerd op een echte 3CX-server (Debian 12, 3CX uit `repo.3cx.com`); `--apply` is daar nog niet gedraaid. Getest op Debian 12 (bookworm) met een nagebootste 3CX-mapindeling en een draaiende `systemd-journald`: proefdraai, `--check-only`, `--apply` en een tweede `--apply`.
