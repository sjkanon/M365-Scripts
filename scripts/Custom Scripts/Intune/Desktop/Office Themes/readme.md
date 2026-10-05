**English** · [Nederlands](readme.nl.md) · [Français](readme.fr.md)

[M365-Scripts](../../../../../readme.md) › [scripts](../../../../readme.md) › [Custom Scripts](../../../readme.md) › [Intune](../../readme.md) › [Desktop](../readme.md) › **Office Themes**

# Office Theme Colors

Deploys an Office color palette (theme colors only, not a full `.thmx` theme — see [`Deploy-OfficeTheme.ps1`](../readme.md#deploy-officethemeps1) in the parent folder for that).

---

## Files

| File | Description |
|------|-------------|
| [`Deploy-Officecolors.ps1`](Deploy-Officecolors.ps1) ([docs](#deploy-officecolorsps1)) | Downloads a color scheme XML from a URL and installs it into Office's Theme Colors folder |

---

### Deploy-Officecolors.ps1

Downloads the color scheme definition (`<a:clrScheme>`) from `-ColorsUrl` to `%APPDATA%\Microsoft\Templates\Document Themes\Theme Colors\`, creating the folder if needed. Once deployed, the scheme appears under Office's **Design > Colors** picker.

**Parameters**

| Parameter | Description |
|-----------|-------------|
| `-ColorsUrl` | Direct download URL of the color scheme `.xml`. Required |
| `-ColorsName` | File name to save it as, ending in `.xml`. Default: the last segment of the URL |

**Examples**

```powershell
.\Deploy-Officecolors.ps1 -ColorsUrl 'https://contoso.blob.core.windows.net/branding/Contoso.xml'
```

**Notes**

- Run as the logged-on user (writes to `%APPDATA%`). Under Intune, deploy as a Win32 app with the parameters on the install command line, or upload a copy with the defaults filled in.
- Without `-ColorsUrl`, or with a name not ending in `.xml`, it stops with exit code 1.
