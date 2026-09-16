#Requires -Version 7.0
<#
.SYNOPSIS
    Read-only: shows exactly what the team lookup gets back from Graph.

.DESCRIPTION
    Temporary. Removes nothing, changes nothing - it only asks the same question
    Remove-SharePointStructure asks and prints the shape of the answer, because that
    shape is the one thing the failure so far has not revealed.
#>
[CmdletBinding()]
param(
    [string] $ConfigPath,
    [string] $Tenant,
    [string] $ClientId
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

. (Join-Path $PSScriptRoot 'SharePointStructure.Common.ps1')

if (-not $ConfigPath) {
    $ConfigPath = @(Get-ChildItem -Path $PSScriptRoot -Filter '*.config.json' |
                    Where-Object { (Get-Content $_.FullName -Raw) -notmatch 'CHANGEME' })[0].FullName
}
$config = Import-StructureConfig -Path $ConfigPath -AllowPlaceholders
if (-not $Tenant) { $Tenant = $config.tenant }

function Show-Shape {
    param([string] $Label, $Value)
    Write-Host "  $Label" -ForegroundColor Cyan
    if ($null -eq $Value) { Write-Host '    <null>' -ForegroundColor Yellow; return }
    Write-Host "    type  : $($Value.GetType().FullName)"
    Write-Host "    count : $(@($Value).Count)"
    if ($Value -is [System.Collections.IDictionary]) {
        Write-Host "    keys  : $(($Value.Keys | ForEach-Object { [string] $_ }) -join ', ')"
    } else {
        $names = @($Value.PSObject.Properties | ForEach-Object { $_.Name })
        Write-Host "    props : $($names -join ', ')"
    }
    Write-Host "    text  : $Value"
}

Connect-StructureGraph -Tenant $Tenant -ClientId $ClientId -AuthenticationOnly `
    -Scopes @('Group.ReadWrite.All', 'Directory.Read.All', 'Channel.ReadBasic.All')

$team    = Get-ConfigValue $config 'team'
$alias   = $team.mailNickname
$escaped = $alias -replace "'", "''"
Write-Host ''
Write-Host "  Looking for mailNickname '$alias'" -ForegroundColor Cyan
Write-Host ''

$found = Invoke-StructureGraph -Url "v1.0/groups?`$filter=mailNickname eq '$escaped'&`$select=id,displayName"
Show-Shape 'raw response' $found

$value = Get-ConfigValue $found 'value' @()
Show-Shape "response 'value'" $value

$group = @($value) | Select-Object -First 1
Show-Shape 'first group' $group
Show-Shape "its 'id' via Get-ConfigValue" (Get-ConfigValue $group 'id')

# The same question without $select, in case the projection is what comes back short.
Write-Host ''
Write-Host '  Same lookup without $select:' -ForegroundColor Cyan
$plain  = Invoke-StructureGraph -Url "v1.0/groups?`$filter=mailNickname eq '$escaped'"
$pgroup = @(Get-ConfigValue $plain 'value' @()) | Select-Object -First 1
Show-Shape 'first group (no $select)' $pgroup
Show-Shape "its 'id'" (Get-ConfigValue $pgroup 'id')

# And what the whole function hands its caller - the thing that ends up in $teamGroup.
Write-Host ''
Write-Host '  What lands in $teamGroup:' -ForegroundColor Cyan
function Test-Return {
    $g = @(Get-ConfigValue (Invoke-StructureGraph -Url "v1.0/groups?`$filter=mailNickname eq '$escaped'&`$select=id,displayName") 'value' @()) |
         Select-Object -First 1
    if (-not $g) { return $null }
    $id = [string] (Get-ConfigValue $g 'id')
    $guid = [guid]::Empty
    if (-not [guid]::TryParse($id, [ref] $guid)) { Write-Warn "id '$id' is not a GUID"; return $null }
    return [pscustomobject]@{ Id = $id; DisplayName = [string] (Get-ConfigValue $g 'displayName' $alias) }
}
$returned = Test-Return
Show-Shape 'returned' $returned
Write-Host ''
Write-Host "  passes 'if (-not `$teamGroup)' : $([bool]$returned)" -ForegroundColor Yellow
try   { Write-Host "  .Id -> '$($returned.Id)'" -ForegroundColor Green }
catch { Write-Host "  .Id THROWS: $($_.Exception.Message)" -ForegroundColor Red }
Write-Host ''
