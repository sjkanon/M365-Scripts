**English** · [Nederlands](readme.nl.md) · [Français](readme.fr.md)

[M365-Scripts](../../../../readme.md) › [scripts](../../../readme.md) › [Custom Scripts](../../readme.md) › [Intune](../readme.md) › **Desktop**

# Desktop (Office theme)

Office theme and color palette deployment via Intune. Both scripts take the download URL of the theme as a parameter, so one script serves every customer; the theme files themselves are not kept in this repo.

For wallpaper/lockscreen/taskbar-shortcut deployment, see [`scripts/Intune/Desktop/`](../../../Intune/Desktop/readme.md).

---

## Folders

| Folder | Description |
|--------|-------------|
| [`Office Themes/`](Office%20Themes/readme.md) | `Deploy-Officecolors.ps1` — installs just a color scheme |

## Scripts

| Script | Description |
|--------|-------------|
| [`Deploy-OfficeTheme.ps1`](Deploy-OfficeTheme.ps1) ([docs](#deploy-officethemeps1)) | Downloads an Office `.thmx` theme from a URL and installs it for the signed-in user |

---

### Deploy-OfficeTheme.ps1

Downloads the `.thmx` from `-ThemeUrl` into `%ProgramData%\OfficeThemes` and copies it into `%APPDATA%\Microsoft\Templates\Document Themes\`, so it appears under Office's **Design > Themes** picker.

**Parameters**

| Parameter | Description |
|-----------|-------------|
| `-ThemeUrl` | Direct download URL of the `.thmx` file. Required |
| `-ThemeName` | File name to save it as, ending in `.thmx` — the name Office shows. Default: the last segment of the URL |

**Examples**

```powershell
.\Deploy-OfficeTheme.ps1 -ThemeUrl 'https://contoso.blob.core.windows.net/branding/Contoso.thmx'

# Win32 app install command
powershell.exe -ExecutionPolicy Bypass -File .\Deploy-OfficeTheme.ps1 -ThemeUrl 'https://example.com/theme.thmx' -ThemeName 'Contoso 2026.thmx'
```

**Notes**

- Run as the logged-on user (writes to `%APPDATA%`).
- Intune platform scripts cannot pass parameters: deploy it as a Win32 app with the parameters on the install command line, or upload a copy with the defaults filled in.
- Without `-ThemeUrl`, or with a name not ending in `.thmx`, it stops with exit code 1 instead of prompting — under Intune nobody would answer the prompt.
