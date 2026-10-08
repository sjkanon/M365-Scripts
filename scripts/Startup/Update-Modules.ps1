#Requires -Version 5.1
<#
.SYNOPSIS
    Controleert en update de PowerShell modules van deze repo, en daarna desgewenst alle andere.

.DESCRIPTION
    Leest de vereiste modules uit RequiredModules.psd1 (dezelfde lijst als Install-Modules.ps1
    en load.ps1) en deelt elke module in:

      Missing          niet geinstalleerd                 -> wordt geinstalleerd
      BelowMinimum     ouder dan MinimumVersion           -> wordt geupdatet
      UpdateAvailable  nieuwere versie op de PSGallery    -> wordt geupdatet
      OK               actueel
      Unknown          PSGallery niet bereikbaar, of de module staat er niet (meer) op
      Skipped          niet voor dit platform / deze PowerShell-versie

    Zonder -RequiredOnly worden daarna ook alle andere via PowerShellGet geinstalleerde
    modules bijgewerkt. Met -CheckOnly verandert er niets: alleen het rapport.

    De PSGallery-versies kunnen worden hergebruikt uit een cache
    (%LOCALAPPDATA%\M365-Scripts\module-gallery-cache.json), zodat load.ps1 de gallery
    maar eens per -MaxAgeHours bevraagt in plaats van bij elke start.

.PARAMETER CheckOnly
    Alleen rapporteren, niets installeren of updaten.

.PARAMETER RequiredOnly
    Alleen de modules uit RequiredModules.psd1, niet alle andere geinstalleerde modules.

.PARAMETER MaxAgeHours
    Hergebruik PSGallery-versies uit de cache als die jonger is dan dit aantal uur.
    0 (standaard) = altijd de gallery bevragen.

.PARAMETER Scope
    Scope voor nieuw te installeren modules: CurrentUser (standaard) of AllUsers (als administrator).

.PARAMETER Quiet
    Geen regel per module; alleen fouten. Bedoeld voor load.ps1, dat zelf samenvat.

.PARAMETER PassThru
    Geef per vereiste module een statusobject terug (Name, Installed, Minimum, Latest, Status, Reason).

.PARAMETER Prompt
    Voor bij het starten (load.ps1, je PowerShell-profiel): toon alleen wat ontbreekt of een
    update heeft, en vraag dan of het nu geinstalleerd/geupdatet moet worden. Alles in orde
    geeft een enkele regel.

.EXAMPLE
    .\Update-Modules.ps1 -RequiredOnly -CheckOnly
    Laat zien welke vereiste modules ontbreken of een update hebben, zonder iets te wijzigen.

.EXAMPLE
    .\Update-Modules.ps1 -RequiredOnly
    Installeert ontbrekende en updatet verouderde vereiste modules.

.EXAMPLE
    .\Update-Modules.ps1 -RequiredOnly -Prompt -MaxAgeHours 24
    Wat load.ps1 bij elke start draait, en wat in je PowerShell-profiel kan: gallery hooguit
    eens per dag, en alleen een vraag als er iets te doen is.

.EXAMPLE
    .\Update-Modules.ps1
    Zoals -RequiredOnly, en werkt daarna alle andere geinstalleerde modules bij.

.NOTES
    Author: Sjoerd Kanon
#>

[CmdletBinding()]
param(
    [switch]$CheckOnly,
    [switch]$RequiredOnly,
    [ValidateRange(0, 8760)]
    [int]$MaxAgeHours = 0,
    [ValidateSet('CurrentUser', 'AllUsers')]
    [string]$Scope = 'CurrentUser',
    [switch]$Quiet,
    [switch]$PassThru,
    [switch]$Prompt
)

# -Prompt shows only what needs doing, then asks; the per-module list would be noise there
if ($Prompt) { $Quiet = $true }

$onWindows = ($PSVersionTable.PSEdition -eq 'Desktop') -or $IsWindows

function Write-Line {
    param([string]$Message, [string]$Color = 'Gray')
    if (-not $Quiet) { Write-Host $Message -ForegroundColor $Color }
}

function ConvertTo-Version {
    param([object]$Value)
    if ($null -eq $Value) { return $null }
    # Strip a prerelease or build suffix ("3.0.0-preview1") before parsing
    $text = ([string]$Value) -replace '[-+].*$', ''
    $parsed = $null
    if ([version]::TryParse($text, [ref]$parsed)) { return $parsed }
    return $null
}

#region Gallery lookup, with a cache

$cacheDir  = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'M365-Scripts'
$cacheFile = Join-Path $cacheDir 'module-gallery-cache.json'

function Get-GalleryVersions {
    # Returns @{ Versions = @{name = [version]}; Reachable = bool; CheckedAt = datetime }
    param([string[]]$Names)

    if ($MaxAgeHours -gt 0 -and (Test-Path $cacheFile)) {
        try {
            $cache = Get-Content $cacheFile -Raw | ConvertFrom-Json
            $checkedAt = [datetime]::Parse($cache.CheckedAt, [Globalization.CultureInfo]::InvariantCulture,
                                           [Globalization.DateTimeStyles]::RoundtripKind)
            $cachedNames = @($cache.Names)
            $covered = -not ($Names | Where-Object { $cachedNames -notcontains $_ })
            if ($covered -and ((Get-Date) - $checkedAt).TotalHours -lt $MaxAgeHours) {
                $versions = @{}
                foreach ($p in $cache.Versions.PSObject.Properties) { $versions[$p.Name] = ConvertTo-Version $p.Value }
                return @{ Versions = $versions; Reachable = $true; CheckedAt = $checkedAt; FromCache = $true }
            }
        } catch {
            # A damaged cache is just a cache miss
        }
    }

    $versions = @{}
    $found = @()
    try {
        # Find-PSResource (PSResourceGet) is roughly three times faster than Find-Module
        if (Get-Command Find-PSResource -ErrorAction SilentlyContinue) {
            $found = @(Find-PSResource -Name $Names -Repository PSGallery -ErrorAction SilentlyContinue)
        } else {
            $found = @(Find-Module -Name $Names -Repository PSGallery -ErrorAction SilentlyContinue)
        }
    } catch {
        $found = @()
    }
    foreach ($f in $found) { $versions[$f.Name] = ConvertTo-Version $f.Version }

    # Nothing back for any name means the gallery was not reachable, not that every module vanished
    $reachable = $versions.Count -gt 0
    $now = Get-Date
    if ($reachable) {
        try {
            if (-not (Test-Path $cacheDir)) { New-Item -ItemType Directory -Path $cacheDir -Force | Out-Null }
            $plain = [ordered]@{}
            foreach ($k in $versions.Keys) { $plain[$k] = [string]$versions[$k] }
            [ordered]@{
                CheckedAt = $now.ToString('o')
                Names     = $Names
                Versions  = $plain
            } | ConvertTo-Json -Depth 4 | Set-Content -Path $cacheFile -Encoding UTF8
        } catch {
            # Without a cache the next start simply asks the gallery again
        }
    }
    return @{ Versions = $versions; Reachable = $reachable; CheckedAt = $now; FromCache = $false }
}

#endregion

#region Required modules: status

$listFile = Join-Path $PSScriptRoot 'RequiredModules.psd1'
$required = @((Import-PowerShellDataFile -Path $listFile).Modules)

$applicable = @()
$skipReason = @{}
$status = New-Object System.Collections.Generic.List[object]

foreach ($req in $required) {
    if ($req.WindowsOnly -and -not $onWindows) {
        $skipReason[$req.Name] = 'alleen Windows'
    } elseif ($req.MinimumPSVersion -and $PSVersionTable.PSVersion -lt [version]$req.MinimumPSVersion) {
        $skipReason[$req.Name] = "vereist PowerShell $($req.MinimumPSVersion)+"
    } else {
        $applicable += $req
    }
}

Write-Line "Vereiste modules controleren ($($applicable.Count))..." 'Cyan'

$names = @($applicable | ForEach-Object { $_.Name })
# One module-path scan for all names; Get-Module -ListAvailable also sees modules that
# were not installed through PowerShellGet, unlike Get-InstalledModule
$installedAll = @()
if ($names.Count) { $installedAll = @(Get-Module -ListAvailable -Name $names -ErrorAction SilentlyContinue) }
$gallery = Get-GalleryVersions -Names $names

if (-not $gallery.Reachable) {
    Write-Host "  PSGallery niet bereikbaar - alleen ontbrekende en te oude modules worden gemeld." -ForegroundColor DarkYellow
} elseif ($gallery.FromCache) {
    Write-Line "  PSGallery-versies uit de cache van $($gallery.CheckedAt.ToString('yyyy-MM-dd HH:mm'))" 'DarkGray'
}

foreach ($req in $required) {
    if ($skipReason.ContainsKey($req.Name)) {
        $status.Add([pscustomobject]@{
            Name = $req.Name; Installed = $null; Minimum = ConvertTo-Version $req.MinimumVersion
            Latest = $null; Status = 'Skipped'; Reason = $skipReason[$req.Name]
        })
        continue
    }
    $installed = $installedAll | Where-Object { $_.Name -eq $req.Name } |
        Sort-Object Version -Descending | Select-Object -First 1
    $installedVersion = if ($installed) { $installed.Version } else { $null }
    $minimum = ConvertTo-Version $req.MinimumVersion
    $latest  = $null
    foreach ($k in $gallery.Versions.Keys) { if ($k -eq $req.Name) { $latest = $gallery.Versions[$k] } }

    $state = if (-not $installedVersion) { 'Missing' }
             elseif ($minimum -and $installedVersion -lt $minimum) { 'BelowMinimum' }
             elseif ($latest -and $installedVersion -lt $latest) { 'UpdateAvailable' }
             elseif (-not $latest) { 'Unknown' }
             else { 'OK' }
    $reason = if ($state -eq 'Unknown' -and $gallery.Reachable) { 'niet gevonden op PSGallery' }
              elseif ($state -eq 'Unknown') { 'PSGallery niet bereikbaar' }
              else { $null }

    $status.Add([pscustomobject]@{
        Name = $req.Name; Installed = $installedVersion; Minimum = $minimum
        Latest = $latest; Status = $state; Reason = $reason
    })
}

foreach ($s in $status) {
    switch ($s.Status) {
        'Missing'         { Write-Line "  [ontbreekt] $($s.Name)$(if ($s.Latest) { " (nieuwste $($s.Latest))" })" 'Yellow' }
        'BelowMinimum'    { Write-Line "  [te oud]    $($s.Name) $($s.Installed) < minimum $($s.Minimum)" 'Yellow' }
        'UpdateAvailable' { Write-Line "  [update]    $($s.Name) $($s.Installed) -> $($s.Latest)" 'Yellow' }
        'OK'              { Write-Line "  [ok]        $($s.Name) $($s.Installed)" 'Gray' }
        'Unknown'         { Write-Line "  [?]         $($s.Name) $($s.Installed) ($($s.Reason))" 'DarkGray' }
        'Skipped'         { Write-Line "  [--]        $($s.Name) ($($s.Reason))" 'DarkGray' }
    }
}

#endregion

#region Required modules: install / update

$todo = @($status | Where-Object { $_.Status -in 'Missing', 'BelowMinimum', 'UpdateAvailable' })

if ($Prompt -and -not $CheckOnly) {
    if ($todo.Count) {
        foreach ($s in $todo) {
            $line = switch ($s.Status) {
                'Missing'         { "ontbreekt         $($s.Name)" }
                'BelowMinimum'    { "te oud            $($s.Name) $($s.Installed) (minimum $($s.Minimum))" }
                'UpdateAvailable' { "update            $($s.Name) $($s.Installed) -> $($s.Latest)" }
            }
            Write-Host "  $line" -ForegroundColor Yellow
        }
        try {
            $answer = Read-Host "  Deze $($todo.Count) module(s) nu installeren/updaten? [J/n]"
        } catch {
            # A -NonInteractive host cannot ask; leave the modules as they are
            $answer = 'n'
        }
        if ($answer -match '^[Nn]') {
            $CheckOnly = $true
            if ($todo | Where-Object { $_.Status -ne 'UpdateAvailable' }) {
                Write-Host "  Scripts die deze modules nodig hebben werken pas na installatie." -ForegroundColor DarkYellow
            }
        }
        $Quiet = $false
    } else {
        Write-Host "  Modules OK ($(@($status | Where-Object Status -eq 'OK').Count) vereist, actueel)" -ForegroundColor DarkGray
    }
}

if (-not $CheckOnly -and $todo.Count) {
    Write-Line "`nVereiste modules installeren/updaten ($($todo.Count))..." 'Cyan'
    foreach ($s in $todo) {
        try {
            # Update-Module only works on modules PowerShellGet installed itself; anything
            # else (a manual copy, an MSI) gets the new version side by side via Install-Module
            $viaGet = Get-InstalledModule -Name $s.Name -ErrorAction SilentlyContinue
            if ($s.Status -ne 'Missing' -and $viaGet) {
                Write-Line "  Updaten: $($s.Name) $($s.Installed) -> $(if ($s.Latest) { $s.Latest } else { 'nieuwste' })" 'Yellow'
                Update-Module -Name $s.Name -Force -ErrorAction Stop
            } else {
                Write-Line "  Installeren: $($s.Name)" 'Yellow'
                Install-Module -Name $s.Name -Scope $Scope -AllowClobber -Force -SkipPublisherCheck -ErrorAction Stop
            }
            $new = Get-Module -ListAvailable -Name $s.Name | Sort-Object Version -Descending | Select-Object -First 1
            $s.Installed = $new.Version
            $s.Status = if ($s.Latest -and $new.Version -lt $s.Latest) { 'UpdateAvailable' } else { 'OK' }
            Write-Line "    OK ($($new.Version))" 'Green'
        } catch {
            Write-Host "    FOUT bij $($s.Name): $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

#endregion

#region All other installed modules

if (-not $RequiredOnly) {
    $handled = @($status | ForEach-Object { $_.Name })
    $modules = @(Get-InstalledModule | Where-Object { $handled -notcontains $_.Name })
    $total = $modules.Count
    $i = 0
    Write-Line "`nOverige geinstalleerde modules ($total)..." 'Cyan'

    foreach ($module in $modules) {
        $i++
        Write-Progress -Activity "Modules controleren" -Status "$($module.Name) ($i/$total)" -PercentComplete (($i / [math]::Max($total, 1)) * 100)
        try {
            $latest = Find-Module -Name $module.Name -ErrorAction Stop
            if ((ConvertTo-Version $latest.Version) -gt (ConvertTo-Version $module.Version)) {
                if ($CheckOnly) {
                    Write-Line "[$i/$total] Update beschikbaar: $($module.Name) $($module.Version) -> $($latest.Version)" 'Yellow'
                } else {
                    Write-Line "[$i/$total] Updaten: $($module.Name) $($module.Version) -> $($latest.Version)" 'Yellow'
                    Update-Module -Name $module.Name -Force -ErrorAction Stop
                    Write-Line "  OK" 'Green'
                }
            } else {
                Write-Line "[$i/$total] Up-to-date: $($module.Name) $($module.Version)" 'Gray'
            }
        } catch {
            Write-Host "  FOUT bij $($module.Name): $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    Write-Progress -Activity "Modules controleren" -Completed
}

#endregion

$open = @($status | Where-Object { $_.Status -in 'Missing', 'BelowMinimum', 'UpdateAvailable' })
if ($CheckOnly) {
    Write-Line "`nKlaar (alleen gecontroleerd): $($open.Count) vereiste module(s) ontbreken of hebben een update." 'Cyan'
} else {
    Write-Line "`nKlaar! $($open.Count) vereiste module(s) nog niet actueel." 'Cyan'
}

if ($PassThru) { $status }
