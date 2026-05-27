##updaten van modules

# Update alle geïnstalleerde PowerShell modules
# Als administrator uitvoeren voor system-wide modules

Write-Host "Geïnstalleerde modules ophalen..." -ForegroundColor Cyan

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