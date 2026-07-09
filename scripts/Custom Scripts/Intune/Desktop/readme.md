# Desktop (Office theme)

Office theme and color palette deployment via Intune. Kept at this path deliberately — both scripts hardcode their download URL to this exact repo location (`main` branch), so moving them would break the download until the scripts are updated and redeployed to Intune.

For wallpaper/lockscreen/taskbar-shortcut deployment, see [`scripts/Intune/Desktop/`](../../../Intune/Desktop/readme.md).

---

## Contents

| Item | Description |
|------|-------------|
| [`Deploy-OfficeTheme.ps1`](#deploy-officethemeps1) | Installs the full VIAS Institute `.thmx` Office theme |
| [`Office Themes/`](Office%20Themes/readme.md) | `Deploy-Officecolors.ps1` — installs just the color scheme |
| `2026 Vias institute colours (2).thmx` | The Office theme file downloaded by `Deploy-OfficeTheme.ps1` |

---

### Deploy-OfficeTheme.ps1

Downloads `2026 Vias institute colours (2).thmx` from this repo's `main` branch on GitHub and copies it into `%APPDATA%\Microsoft\Templates\Document Themes\`, so it appears under Office's **Design > Themes** picker.

```powershell
.\Deploy-OfficeTheme.ps1
```

> No parameters — source URL and filename are hardcoded at the top of the script. Deploy via Intune as the logged-on user (writes to `%APPDATA%`).
