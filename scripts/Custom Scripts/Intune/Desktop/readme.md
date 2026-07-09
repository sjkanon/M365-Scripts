# Desktop

Intune-deployed desktop customization: Office theme/colors, corporate wallpaper + lockscreen, and a taskbar lock-workstation shortcut.

---

## Contents

| Item | Description |
|------|-------------|
| [`Deploy-OfficeTheme.ps1`](#deploy-officethemeps1) | Installs the full VIAS Institute `.thmx` Office theme |
| [`Office Themes/`](Office%20Themes/readme.md) | `Deploy-Officecolors.ps1` — installs just the color scheme |
| [`Background/`](Background/readme.md) | Corporate wallpaper (`Desktop/`) and lockscreen (`Lockscreen/`) |
| [`Add Lockscreen to start and desktop/`](Add%20Lockscreen%20to%20start%20and%20desktop/readme.md) | Pins a "Lock Workstation" shortcut to Start |
| `2026 Vias institute colours (2).thmx` | The Office theme file downloaded by `Deploy-OfficeTheme.ps1` |

---

### Deploy-OfficeTheme.ps1

Downloads `2026 Vias institute colours (2).thmx` from this repo's `main` branch on GitHub and copies it into `%APPDATA%\Microsoft\Templates\Document Themes\`, so it appears under Office's **Design > Themes** picker.

```powershell
.\Deploy-OfficeTheme.ps1
```

> No parameters — source URL and filename are hardcoded at the top of the script. Deploy via Intune as the logged-on user (writes to `%APPDATA%`).
