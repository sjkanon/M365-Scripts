# ================================================
# Scheduled Task aanmaken voor Licentie Rapport
# Uitvoeren als Administrator op fir-app-001
# ================================================

$TaskName   = "BraveHub - Licentie Overzicht Generator"
$TaskDesc   = "Genereert maandelijks het licentie- en Azure-kostenrapport via Pax8 en Ingram data."
$ScriptPath = (Join-Path $PSScriptRoot "genereer_licentie_overzicht.py")
$RunAsUser  = "BRAVEHUB\sa-halo"  # Aanpassen indien ander domein

# ── Trigger via XML: elke 6e van de maand om 08:00 ──
$TriggerXml = @"
<CalendarTrigger xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <StartBoundary>2026-03-06T08:00:00</StartBoundary>
  <Enabled>true</Enabled>
  <ScheduleByMonth>
    <DaysOfMonth>
      <Day>6</Day>
    </DaysOfMonth>
    <Months>
      <January/>
      <February/>
      <March/>
      <April/>
      <May/>
      <June/>
      <July/>
      <August/>
      <September/>
      <October/>
      <November/>
      <December/>
    </Months>
  </ScheduleByMonth>
</CalendarTrigger>
"@

$Action = New-ScheduledTaskAction `
    -Execute "python.exe" `
    -Argument "`"$ScriptPath`""

$Settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Hours 2) `
    -RestartCount 1 `
    -RestartInterval (New-TimeSpan -Minutes 30) `
    -StartWhenAvailable `
    -RunOnlyIfNetworkAvailable

# Registreren met een dagelijkse trigger als placeholder
$PlaceholderTrigger = New-ScheduledTaskTrigger -Daily -At "08:00"

Register-ScheduledTask `
    -TaskName    $TaskName `
    -Description $TaskDesc `
    -Trigger     $PlaceholderTrigger `
    -Action      $Action `
    -Settings    $Settings `
    -RunLevel    Highest `
    -User        $RunAsUser `
    -Force | Out-Null

# Trigger vervangen door maandelijks schema via XML
$Task    = Get-ScheduledTask -TaskName $TaskName
$TaskXml = [xml]($Task | Export-ScheduledTask)

# Verwijder bestaande triggers en vervang door maandelijkse trigger
$ns  = "http://schemas.microsoft.com/windows/2004/02/mit/task"
$triggersNode = $TaskXml.Task.Triggers

# Verwijder alle bestaande triggers
while ($triggersNode.HasChildNodes) {
    $triggersNode.RemoveChild($triggersNode.FirstChild) | Out-Null
}

# Voeg maandelijkse trigger in
$newTrigger = $TaskXml.CreateDocumentFragment()
$newTrigger.InnerXml = $TriggerXml
$triggersNode.AppendChild($newTrigger) | Out-Null

# Sla het nieuwe XML op en herregistreer
$UpdatedXml = $TaskXml.OuterXml
Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
Register-ScheduledTask -TaskName $TaskName -Xml $UpdatedXml -User $RunAsUser -Force | Out-Null

Write-Host ""
Write-Host "Scheduled task aangemaakt:" -ForegroundColor Green
Write-Host "  Naam    : $TaskName"
Write-Host "  Account : $RunAsUser"
Write-Host "  Schema  : Elke 6e van de maand om 08:00"
Write-Host "  Script  : $ScriptPath"
Write-Host ""
Write-Host "Handmatig testen:" -ForegroundColor Yellow
Write-Host "  Start-ScheduledTask -TaskName '$TaskName'"