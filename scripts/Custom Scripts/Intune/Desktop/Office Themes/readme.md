# Office Theme Colors

Deploys the VIAS Institute Office color palette (theme colors only, not the full `.thmx` theme — see [`Deploy-OfficeTheme.ps1`](../readme.md#deploy-officethemeps1) in the parent folder for that).

---

## Files

| File | Description |
|------|-------------|
| [`Deploy-Officecolors.ps1`](#deploy-officecolorsps1) | Downloads and installs the color scheme XML into Office's Theme Colors folder |
| `Test VIAS.xml` | The color scheme definition (`<a:clrScheme>`) — dark/light/accent colors for the VIAS Institute theme |

---

### Deploy-Officecolors.ps1

Downloads `Test VIAS.xml` from this repo's `main` branch on GitHub and copies it to `%APPDATA%\Microsoft\Templates\Document Themes\Theme Colors\`, creating the folder if needed. Once deployed, "Test VIAS" appears as a selectable color scheme under Office's **Design > Colors** picker.

```powershell
.\Deploy-Officecolors.ps1
```

> No parameters — source URL and filename are hardcoded at the top of the script. Deploy via Intune as the logged-on user (writes to `%APPDATA%`).
