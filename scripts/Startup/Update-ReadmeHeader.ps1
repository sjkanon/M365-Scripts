#Requires -Version 5.1
<#
.SYNOPSIS
    Put the language switcher and the breadcrumb at the top of every readme, in all
    three languages. Supports -WhatIf.

.DESCRIPTION
    Every folder has its readme three times: readme.md (English), readme.nl.md
    (Dutch) and readme.fr.md (French). Each one opens with two lines that are the
    same everywhere except for the paths in them:

      English · [Nederlands](readme.nl.md) · [Français](readme.fr.md)

      [M365-Scripts](../../readme.md) › [scripts](../readme.md) › Intune

    The first switches to the same page in another language, the second climbs back
    up the tree - each level links to its own readme in the current language. Kept
    by hand, those paths are exactly what goes wrong: one ../ too few after a folder
    moves, or a Dutch page that suddenly links to the English parent. So they are
    generated, from the folder the readme is in and nothing else.

    Everything above the first heading that is a switcher or breadcrumb line is
    replaced; the rest of the file is left alone. A readme without a sibling in one
    of the languages is reported, because its switcher would link to nothing.

.PARAMETER Root
    Repository root (default: the folder two levels above this script).

.PARAMETER Check
    Do not write. Exit 1 when a readme's header is not what it should be, or a
    language version is missing - for a pre-commit hook or a pipeline.

.EXAMPLE
    .\Update-ReadmeHeader.ps1

    Rewrite the header of every readme, readme.nl.md and readme.fr.md.

.EXAMPLE
    .\Update-ReadmeHeader.ps1 -Check

    Exit 1 when a header is stale or a readme has no Dutch or French version.

.NOTES
    Author  : Sjoerd Kanon
    Runs on Windows, macOS and Linux - it only reads and rewrites markdown.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string] $Root,
    [switch] $Check
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not $Root) { $Root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot) }
$Root = (Resolve-Path -LiteralPath $Root).Path

$languages = @(
    [pscustomobject]@{ Suffix = '';    Name = 'English' }
    [pscustomobject]@{ Suffix = '.nl'; Name = 'Nederlands' }
    [pscustomobject]@{ Suffix = '.fr'; Name = 'Français' }
)
$utf8 = New-Object System.Text.UTF8Encoding($false)

function Get-ReadmeHeader {
    <#
        The lines that belong above the first heading of one readme: the language
        switcher, and - below the root - the breadcrumb. $Parts is the folder path
        from the repository root, one element per level.
    #>
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [string[]] $Parts,
        [Parameter(Mandatory)] [AllowEmptyString()] [string] $Suffix
    )

    $switch = foreach ($language in $languages) {
        if ($language.Suffix -eq $Suffix) { "**$($language.Name)**" }
        else { "[$($language.Name)](readme$($language.Suffix).md)" }
    }
    $header = @(($switch -join ' · '), '')

    if ($Parts.Count -gt 0) {
        $depth  = $Parts.Count
        $crumbs = @("[M365-Scripts](" + ('../' * $depth) + "readme$Suffix.md)")
        for ($i = 0; $i -lt $depth - 1; $i++) {
            $crumbs += "[$($Parts[$i])](" + ('../' * ($depth - 1 - $i)) + "readme$Suffix.md)"
        }
        $crumbs += "**$($Parts[-1])**"
        $header += @(($crumbs -join ' › '), '')
    }
    return , $header
}

# Every folder that has an English readme is expected to have all three.
$scriptsDir = Join-Path $Root 'scripts'
$folders = @(Get-Item -LiteralPath $Root) + @(Get-Item -LiteralPath $scriptsDir) +
           @(Get-ChildItem -LiteralPath $scriptsDir -Recurse -Directory)
$folders = @($folders | Where-Object { Test-Path -LiteralPath (Join-Path $_.FullName 'readme.md') })

$stale   = New-Object System.Collections.Generic.List[string]
$missing = New-Object System.Collections.Generic.List[string]
$written = 0

foreach ($folder in $folders) {
    $relative = $folder.FullName.Substring($Root.Length).TrimStart('\', '/') -replace '\\', '/'
    $parts    = @(if ($relative) { $relative -split '/' })

    foreach ($language in $languages) {
        $file    = Join-Path $folder.FullName "readme$($language.Suffix).md"
        $display = (@($relative, "readme$($language.Suffix).md") | Where-Object { $_ }) -join '/'
        if (-not (Test-Path -LiteralPath $file)) { $missing.Add($display); continue }

        $raw     = [System.IO.File]::ReadAllText($file)
        $newline = if ($raw.Contains("`r`n")) { "`r`n" } else { "`n" }
        $lines   = @($raw -split "`r?`n")
        if ($lines.Count -gt 1 -and $lines[-1] -eq '') { $lines = $lines[0..($lines.Count - 2)] }

        # Drop the old header: blank lines, switcher lines and breadcrumb lines before the body.
        $start = 0
        while ($start -lt $lines.Count -and (
               $lines[$start] -eq '' -or
               $lines[$start] -match '^(\*\*|\[)(English|Nederlands|Français)(\*\*|\])' -or
               $lines[$start] -match '^\[M365-Scripts\]\(')) { $start++ }
        $body = if ($start -lt $lines.Count) { $lines[$start..($lines.Count - 1)] } else { @() }

        $header  = Get-ReadmeHeader -Parts $parts -Suffix $language.Suffix
        $content = ((@($header) + @($body)) -join $newline) + $newline
        if ($content -ceq $raw) { continue }

        $stale.Add($display)
        if (-not $Check -and $PSCmdlet.ShouldProcess($display, 'Rewrite readme header')) {
            [System.IO.File]::WriteAllText($file, $content, $utf8)
            $written++
        }
    }
}

Write-Host ''
Write-Host "  $($folders.Count) folder(s), $($folders.Count * $languages.Count) readme(s) expected." -ForegroundColor Cyan
if ($missing.Count) {
    Write-Host "  Missing language version(s):" -ForegroundColor Yellow
    $missing | ForEach-Object { Write-Host "    $_" -ForegroundColor Yellow }
}

if ($Check) {
    if ($stale.Count) {
        Write-Host "  Header out of date:" -ForegroundColor Yellow
        $stale | ForEach-Object { Write-Host "    $_" -ForegroundColor Yellow }
        Write-Host '  Rerun without -Check to rewrite them.' -ForegroundColor Yellow
    }
    if ($stale.Count -or $missing.Count) { exit 1 }
    Write-Host '  Every readme header is current.' -ForegroundColor Green
    return
}

Write-Host "  $written header(s) rewritten." -ForegroundColor Green
if ($missing.Count) { exit 1 }
