#Requires -Version 5.1
<#
.SYNOPSIS
    Build scripts/INDEX.md - one searchable A-Z table of every script in the repo,
    with a link to the file, its folder readme and what it does. Supports -WhatIf.

.DESCRIPTION
    Finding a script on GitHub used to mean guessing which workload folder it lived
    in and opening readmes until it turned up. This walks scripts/ once and writes a
    single page you can Ctrl-F: script name, folder, and the one-line description
    taken from the script itself.

    The description comes from the script's own header, in this order:

      1. the first line of its .SYNOPSIS block - the documented answer;
      2. failing that, the first real line of a leading # comment block, which is
         what the older scripts here have instead;
      3. failing that, nothing at all, and the script is listed under
         "Scripts without a description" so it is visible rather than silently
         blank. That list is the point: it says which headers still need writing.

    Because the page is generated, it cannot drift from the files the way a
    hand-kept table does - rerun it after adding, renaming or removing a script.
    -Check reports whether the committed page is still current without writing
    anything, which is what a pre-commit hook or a pipeline would call.

.PARAMETER Root
    Repository root (default: the folder two levels above this script).

.PARAMETER Check
    Do not write. Compare the generated page with the one on disk and exit 1 when
    they differ, so a stale index fails instead of going unnoticed.

.EXAMPLE
    .\Update-ScriptIndex.ps1

    Rewrite scripts/INDEX.md from the scripts currently in the repo.

.EXAMPLE
    .\Update-ScriptIndex.ps1 -WhatIf

    Report what would change - how many scripts, and which have no description yet.

.EXAMPLE
    .\Update-ScriptIndex.ps1 -Check

    Exit 1 when scripts/INDEX.md no longer matches the scripts on disk.

.NOTES
    Author  : Sjoerd Kanon
    Runs on Windows, macOS and Linux - it only reads files and writes one page.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string] $Root,
    [switch] $Check
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not $Root) { $Root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot) }
$scriptsDir = Join-Path $Root 'scripts'
$indexPath  = Join-Path $scriptsDir 'INDEX.md'

if (-not (Test-Path $scriptsDir)) { throw "No scripts folder under $Root" }

function Get-ScriptDescription {
    <#
        The one-line description for a script, from its own header. .SYNOPSIS is the
        documented place; a leading # comment block is what the older scripts here
        have instead, and is better than an empty cell. Only the first 60 lines are
        read - a header that has not started by then is not a header.
    #>
    param([Parameter(Mandatory)] [string] $Path)

    $lines = @()
    try { $lines = @(Get-Content -LiteralPath $Path -TotalCount 60 -ErrorAction Stop) } catch { return '' }

    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -notmatch '^\s*\.SYNOPSIS\s*$') { continue }

        # The whole paragraph, not its first line: these synopses wrap at ~80
        # columns, and taking one line leaves half a sentence in the table.
        $parts   = [System.Collections.Generic.List[string]]::new()
        $started = $false
        for ($j = $i + 1; $j -lt $lines.Count; $j++) {
            $text = $lines[$j].Trim()
            # Leading blanks are padding; a blank after the text ends the paragraph.
            if ($text -eq '') { if ($started) { break } else { continue } }
            # An empty .SYNOPSIS runs straight into the next keyword or the block end.
            if ($text -match '^\.[A-Za-z]' -or $text -match '^#>') { break }
            # A synopsis that opens with a sentence and then lists its cases belongs
            # in the script, not in a table cell: the lead-in is the description, and
            # dragging half of the first bullet in behind it only reads as truncated.
            if ($started -and $text -match '^([-*•]|\d+[.)])\s') { break }
            $parts.Add($text)
            $started = $true
        }
        return ($parts -join ' ')
    }

    # No .SYNOPSIS: fall back to a leading # comment block, skipping the directives
    # and separator rules that carry no meaning for a reader.
    $comments = [System.Collections.Generic.List[string]]::new()
    $followedByBlank = $false
    foreach ($line in $lines) {
        $text = $line.Trim()
        if ($text -eq '') {
            if ($comments.Count -gt 0) { $followedByBlank = $true; break }
            continue
        }
        if ($text -match '^#Requires') { continue }
        if ($text -match '^<#') { return '' }          # a comment block with no .SYNOPSIS
        if ($text -notmatch '^#') { break }            # code started, so the block ends here
        $text = ($text -replace '^#+', '').Trim()
        if ($text -eq '' -or $text -match '^[-=_~*]{3,}$') { continue }
        $comments.Add($text)
    }

    # One comment line sitting straight on top of code is a comment about that line -
    # "# URL van de theme" above a $ThemeUrl assignment - not a description of the
    # script. Reading it as one puts something worse than nothing in the table. A real
    # header runs to several lines, or is set off from the code by a blank line.
    if ($comments.Count -eq 0) { return '' }
    if ($comments.Count -eq 1 -and -not $followedByBlank) { return '' }
    return $comments[0]
}

function Format-Cell {
    <# Keep a description on one table row: no pipes, no newlines, not endless. #>
    param([string] $Text, [int] $MaxLength = 160)

    if (-not $Text) { return '' }
    $clean = ($Text -replace '\s+', ' ').Trim() -replace '\|', '\|'
    # A lead-in whose list was dropped ends on a colon or a dash, which reads as a
    # sentence that got cut off. It did not - the list simply belongs in the script.
    $clean = $clean -replace '\s*[:,;-]$', ''
    if ($clean.Length -gt $MaxLength) { $clean = $clean.Substring(0, $MaxLength - 1).TrimEnd() + '…' }
    return $clean
}

function ConvertTo-MarkdownPath {
    <# A relative path GitHub will follow: forward slashes, spaces percent-encoded. #>
    param([Parameter(Mandatory)] [string] $RelativePath)

    $parts = $RelativePath -split '[\\/]'
    return (($parts | ForEach-Object { [Uri]::EscapeDataString($_) }) -join '/')
}

# -- Collect -------------------------------------------------------------------
$items = foreach ($file in (Get-ChildItem -LiteralPath $scriptsDir -Recurse -Filter '*.ps1' -File)) {
    $relative = $file.FullName.Substring($scriptsDir.Length).TrimStart('\', '/')
    $folder   = Split-Path -Parent $relative
    if (-not $folder) { $folder = '.' }

    [PSCustomObject]@{
        Name        = $file.Name
        Folder      = ($folder -replace '\\', '/')
        LinkPath    = ConvertTo-MarkdownPath $relative
        Description = Format-Cell (Get-ScriptDescription -Path $file.FullName)
    }
}
$items = @($items | Sort-Object Name, Folder)

$described   = @($items | Where-Object { $_.Description })
$undescribed = @($items | Where-Object { -not $_.Description })

# A folder is only worth linking to when it actually has a readme to land on.
$folders = @($items | Select-Object -ExpandProperty Folder -Unique | Sort-Object)
$folderLink = @{}
foreach ($folder in $folders) {
    $readme = if ($folder -eq '.') { Join-Path $scriptsDir 'readme.md' }
              else { Join-Path $scriptsDir (Join-Path $folder 'readme.md') }
    $folderLink[$folder] = (Test-Path -LiteralPath $readme)
}

function Get-FolderCell {
    param([string] $Folder)
    if ($Folder -eq '.') { return '`scripts/`' }
    if ($folderLink[$Folder]) {
        return '[`{0}/`]({1}/readme.md)' -f $Folder, (ConvertTo-MarkdownPath $Folder)
    }
    return '`{0}/`' -f $Folder
}

# -- Render --------------------------------------------------------------------
$lines = [System.Collections.Generic.List[string]]::new()
$lines.Add('# Script index')
$lines.Add('')
$lines.Add('<!-- Generated by scripts/Startup/Update-ScriptIndex.ps1 - do not edit by hand. -->')
$lines.Add('')
$lines.Add(('Every script in this repository, A-Z, with the folder it lives in. Use your browser''s find (Ctrl+F) — this page exists so you do not have to guess which workload folder a script is under.'))
$lines.Add('')
$lines.Add(('{0} scripts across {1} folders. Regenerate with `pwsh -File scripts/Startup/Update-ScriptIndex.ps1` after adding, renaming or removing one.' -f $items.Count, $folders.Count))
$lines.Add('')
$lines.Add('> Looking for what a category contains rather than a specific script? Start at the [folder overview](readme.md).')
$lines.Add('')
$lines.Add('---')
$lines.Add('')
$lines.Add('## All scripts')
$lines.Add('')
$lines.Add('| Script | Folder | What it does |')
$lines.Add('|--------|--------|--------------|')
foreach ($item in $described) {
    $lines.Add(('| [`{0}`]({1}) | {2} | {3} |' -f $item.Name, $item.LinkPath, (Get-FolderCell $item.Folder), $item.Description))
}

if ($undescribed.Count -gt 0) {
    $lines.Add('')
    $lines.Add('---')
    $lines.Add('')
    $lines.Add('## Scripts without a description')
    $lines.Add('')
    $lines.Add(('These {0} have no `.SYNOPSIS` and no usable leading comment, so this page cannot say what they do. Their folder readme can — and adding a `.SYNOPSIS` to the script itself removes it from this list.' -f $undescribed.Count))
    $lines.Add('')
    $lines.Add('| Script | Folder |')
    $lines.Add('|--------|--------|')
    foreach ($item in $undescribed) {
        $lines.Add(('| [`{0}`]({1}) | {2} |' -f $item.Name, $item.LinkPath, (Get-FolderCell $item.Folder)))
    }
}

$lines.Add('')
$rendered = ($lines -join "`n") + "`n"

# -- Write or compare ----------------------------------------------------------
$existing = if (Test-Path -LiteralPath $indexPath) {
    (Get-Content -LiteralPath $indexPath -Raw) -replace "`r`n", "`n"
} else { $null }

$current = ($existing -eq $rendered)

Write-Host ''
Write-Host ('  {0} script(s) indexed, {1} with a description, {2} without.' -f $items.Count, $described.Count, $undescribed.Count) -ForegroundColor Cyan
foreach ($item in $undescribed) {
    Write-Host ('    no description: {0}/{1}' -f $item.Folder, $item.Name) -ForegroundColor DarkGray
}

if ($Check) {
    if ($current) {
        Write-Host '  scripts/INDEX.md is up to date.' -ForegroundColor Green
        exit 0
    }
    Write-Host '  scripts/INDEX.md is out of date - rerun without -Check to rebuild it.' -ForegroundColor Yellow
    exit 1
}

if ($current) {
    Write-Host '  scripts/INDEX.md is already up to date - nothing written.' -ForegroundColor DarkGray
    exit 0
}

if ($PSCmdlet.ShouldProcess($indexPath, 'Write the generated script index')) {
    Set-Content -LiteralPath $indexPath -Value $rendered -Encoding UTF8 -NoNewline
    Write-Host ('  Wrote {0}' -f $indexPath) -ForegroundColor Green
}

Write-Host ''
exit 0
