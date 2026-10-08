#Requires -Version 5.1
<#
.SYNOPSIS
    Reports modules that scripts load but RequiredModules.psd1 does not list.

.DESCRIPTION
    RequiredModules.psd1 is what load.ps1 installs and keeps up to date on every machine.
    A script that starts using a new module works on the machine it was written on and
    fails everywhere else until someone adds the module to that list - which is easy to
    forget. This script finds those modules, so the docs hook can say so the moment the
    script is saved.

    Every .ps1/.psm1 in the repository (outside .claude) is searched for module names in:

      #Requires -Modules A, B, @{ ModuleName = 'C'; ... }
      Import-Module [-Name] A
      Install-Module [-Name] A

    Only literal names count; a name in a variable ($mod) cannot be checked. A name found
    in neither Modules nor NotManaged in RequiredModules.psd1 is reported with the files
    that use it.

.PARAMETER Root
    Repository root (default: two levels above this script).

.EXAMPLE
    pwsh -File scripts/Startup/Test-RequiredModules.ps1
    Lists modules missing from RequiredModules.psd1; exit code 1 when there are any.

.NOTES
    Exit codes: 0 = every module a script loads is listed, 1 = something is missing.
#>
[CmdletBinding()]
param(
    [string]$Root
)

# $PSScriptRoot is not yet set in a parameter default under Windows PowerShell 5.1
if (-not $Root) { $Root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot) }

$list    = Import-PowerShellDataFile -Path (Join-Path $PSScriptRoot 'RequiredModules.psd1')
$known   = @($list.Modules | ForEach-Object { $_.Name })
$ignored = @(if ($list.NotManaged) { $list.NotManaged.Keys })

$name = "[A-Za-z][\w]*(?:\.[\w]+)*"
# A #Requires statement starts at column 0, so an indented example in help text is not one
$rxRequires = [regex]"(?im)^#Requires\s+-Modules\s+(?<list>.+)$"
$rxCmd      = [regex]"(?i)\b(?:Import|Install)-Module\s+(?:-Name\s+)?['`"]?(?<n>$name)"
$rxSpecName = [regex]"(?i)ModuleName\s*=\s*['`"](?<n>$name)"

$found = @{}
$files = Get-ChildItem -Path $Root -Recurse -File -Include *.ps1, *.psm1 |
    Where-Object { $_.FullName -notmatch '[\\/]\.(claude|git)[\\/]' }

foreach ($file in $files) {
    $text = Get-Content -LiteralPath $file.FullName -Raw
    if (-not $text) { continue }
    $names = @()
    foreach ($m in $rxRequires.Matches($text)) {
        $line = $m.Groups['list'].Value
        $names += $rxSpecName.Matches($line) | ForEach-Object { $_.Groups['n'].Value }
        # Strip module specifications, then the rest is a comma-separated list of names
        $plain = $line -replace '@\{[^}]*\}', ''
        $names += $plain -split ',' | ForEach-Object { $_.Trim().Trim("'", '"') } | Where-Object { $_ -match "^$name$" }
    }
    $names += $rxCmd.Matches($text) | ForEach-Object { $_.Groups['n'].Value }

    $relative = $file.FullName.Substring($Root.Length).TrimStart('\', '/') -replace '\\', '/'
    foreach ($n in $names | Select-Object -Unique) {
        if (-not $found.ContainsKey($n)) { $found[$n] = New-Object System.Collections.Generic.List[string] }
        $found[$n].Add($relative)
    }
}

$missing = @($found.Keys | Where-Object { $known -notcontains $_ -and $ignored -notcontains $_ } | Sort-Object)

Write-Host "  $($files.Count) script(s) searched, $($found.Count) module name(s) found, $($known.Count) listed in RequiredModules.psd1."
if (-not $missing.Count) {
    Write-Host "  Every module a script loads is in RequiredModules.psd1." -ForegroundColor Green
    exit 0
}

Write-Host ""
foreach ($n in $missing) {
    Write-Host "  MISSING  $n" -ForegroundColor Red
    foreach ($f in $found[$n] | Select-Object -First 5) { Write-Host "           used in $f" }
}
Write-Host ""
Write-Host "  Add each to Modules in scripts/Startup/RequiredModules.psd1 (so load.ps1 installs it)," -ForegroundColor Yellow
Write-Host "  or to NotManaged with the reason it is not installed from the gallery." -ForegroundColor Yellow
exit 1
