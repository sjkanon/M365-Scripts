## updaten van modules

# Update alle geinstalleerde PowerShell modules
# Als administrator uitvoeren voor system-wide modules

Write-Host "Geinstalleerde modules ophalen..." -ForegroundColor Cyan

$requiredModules = @(
    @{ Name = 'Microsoft.Graph.Authentication'; MinimumVersion = '2.0.0' }
    @{ Name = 'Microsoft.Graph.Sites'; MinimumVersion = '2.0.0' }
    @{ Name = 'Microsoft.Graph.Identity.DirectoryManagement'; MinimumVersion = '2.0.0' }
    @{ Name = 'Microsoft.Graph.Identity.SignIns'; MinimumVersion = '2.0.0' }
    @{ Name = 'Microsoft.Graph.Identity.Governance'; MinimumVersion = '2.0.0' }
    @{ Name = 'Microsoft.Graph.Applications'; MinimumVersion = '2.0.0' }
    @{ Name = 'Microsoft.Graph.Groups'; MinimumVersion = '2.0.0' }
)

foreach ($req in $requiredModules) {
    try {
        $installed = Get-InstalledModule -Name $req.Name -ErrorAction SilentlyContinue

        if (-not $installed) {
            Write-Host "Installeren ontbrekende module: $($req.Name) (min $($req.MinimumVersion))" -ForegroundColor Yellow
            Install-Module -Name $req.Name -MinimumVersion $req.MinimumVersion -Scope CurrentUser -AllowClobber -Force -ErrorAction Stop
            Write-Host "  OK" -ForegroundColor Green
            continue
        }

        if ([Version]$installed.Version -lt [Version]$req.MinimumVersion) {
            Write-Host "Updaten vereiste module: $($req.Name) $($installed.Version) -> min $($req.MinimumVersion)" -ForegroundColor Yellow
            Update-Module -Name $req.Name -Force -ErrorAction Stop
            Write-Host "  OK" -ForegroundColor Green
        } else {
            Write-Host "Vereiste module OK: $($req.Name) $($installed.Version)" -ForegroundColor Gray
        }
    } catch {
        Write-Host "  FOUT bij vereiste module $($req.Name): $($_.Exception.Message)" -ForegroundColor Red
    }
}

$modules = Get-InstalledModule
$total = $modules.Count
$i = 0

foreach ($module in $modules) {
    $i++
    Write-Progress -Activity "Modules updaten" -Status "$($module.Name) ($i/$total)" -PercentComplete (($i / $total) * 100)

    try {
        $latest = Find-Module -Name $module.Name -ErrorAction Stop
        if ($latest.Version -gt $module.Version) {
            Write-Host "[$i/$total] Updaten: $($module.Name) $($module.Version) -> $($latest.Version)" -ForegroundColor Yellow
            Update-Module -Name $module.Name -Force -ErrorAction Stop
            Write-Host "  OK" -ForegroundColor Green
        } else {
            Write-Host "[$i/$total] Up-to-date: $($module.Name) $($module.Version)" -ForegroundColor Gray
        }
    } catch {
        Write-Host "  FOUT bij $($module.Name): $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host "`nKlaar! $total modules gecontroleerd." -ForegroundColor Cyan