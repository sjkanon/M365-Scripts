# Stel de werkdirectory in
$workDirectory = "C:\Users\sjoerd.kanon\BraveHub\C - Grant Thornton & First Partnership - General\Security Meetings"
Set-Location -Path $workDirectory

# Krijg de huidige datum en tijd
$now = Get-Date

# Maak de jaar- en maandnamen
$year = $now.Year
$month = $now.Month.ToString("00")

# Maak de paden voor de jaar- en maandmappen
$yearFolderPath = Join-Path -Path $workDirectory -ChildPath $year
$monthFolderPath = Join-Path -Path $yearFolderPath -ChildPath "$year $month"

# Controleer of de jaar- en maandmappen al bestaan
if (-not (Test-Path -Path $yearFolderPath)) {
    # Maak de jaarmap als deze nog niet bestaat
    New-Item -ItemType Directory -Path $yearFolderPath | Out-Null
}

if (-not (Test-Path -Path $monthFolderPath)) {
    # Maak de maandmap als deze nog niet bestaat
    New-Item -ItemType Directory -Path $monthFolderPath | Out-Null

    # Maak de submappen binnen de maandmap
    $subfolders = @("Defender", "Rapid 7", "Mimecast")
    foreach ($subfolder in $subfolders) {
        $subfolderPath = Join-Path -Path $monthFolderPath -ChildPath $subfolder
        New-Item -ItemType Directory -Path $subfolderPath | Out-Null
    }

    Write-Host "Maandmap en submappen zijn succesvol gemaakt."
} else {
    Write-Host "De maandmap bestaat al."
}
