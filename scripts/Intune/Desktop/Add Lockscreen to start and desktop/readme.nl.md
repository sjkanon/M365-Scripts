[English](readme.md) · **Nederlands** · [Français](readme.fr.md)

[M365-Scripts](../../../../readme.nl.md) › [scripts](../../../readme.nl.md) › [Intune](../../readme.nl.md) › [Desktop](../readme.nl.md) › **Add Lockscreen to start and desktop**

# Vergrendelen toevoegen aan Start en bureaublad

Maakt een snelkoppeling "Lock Workstation" vast aan Start / bureaublad via een gedownloade `.bat` + `.ico`, met het ongedocumenteerde Explorer-werkwoord `Windows.taskbarpin` om vast te maken zonder tussenkomst van de gebruiker.

---

## Bestanden

| Bestand | Omschrijving |
|------|-------------|
| [`add-lock.ps1`](add-lock.ps1) ([docs](#add-lockps1)) | Downloadt het vergrendelscript + pictogram en maakt de snelkoppeling aan |
| [`add-shortcut-lock.ps1`](add-shortcut-lock.ps1) ([docs](#add-shortcut-lockps1)) | Maakt een willekeurige snelkoppeling vast via het Explorer-werkwoord `Windows.taskbarpin` |

---

### add-lock.ps1

Maakt `C:\Program Files\EOO\lockworkstation\` aan (slaat over als die al bestaat), downloadt `lock.txt` + `lock.ico` van `https://endpoint.eoo.cloud/lock/`, hernoemt `lock.txt` naar `lock.bat` en maakt onder `C:\ProgramData\...\Start Menu\Programs\EOO\lockworkstation\` een snelkoppeling in het Startmenu (`lock.lnk`) aan die `explorer.exe` start met de `.bat` als argument.

```powershell
.\add-lock.ps1
```

> Geen parameters. Bedoeld om één keer per apparaat te draaien (Intune, SYSTEM-context) — opnieuw uitvoeren doet niets meer zodra de map `EOO` bestaat.

---

### add-shortcut-lock.ps1

Maakt een doelbestand vast aan de taakbalk/Start met het registerwerkwoord `Windows.taskbarpin` `ExplorerCommandHandler`, en ruimt daarna de tijdelijke registersleutels op die het heeft aangemaakt.

**Parameters**

| Parameter | Verplicht | Omschrijving |
|-----------|----------|-------------|
| `-Target` | Ja | Pad naar het bestand/de snelkoppeling om vast te maken |

```powershell
.\add-shortcut-lock.ps1 -Target "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\EOO\lockworkstation\lock.lnk"
```

> Afhankelijk van de snelkoppeling `lock.lnk` die door `add-lock.ps1` wordt aangemaakt.
