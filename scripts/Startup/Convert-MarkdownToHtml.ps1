#Requires -Version 5.1
<#
.SYNOPSIS
    Turn a repository markdown document into one self-contained, styled HTML page -
    made to paste into IT Glue or to print. Supports -WhatIf.

.DESCRIPTION
    The service desk documents in this repo are written in markdown, but IT Glue and
    most ticket systems want rich text. Converting by hand means the HTML is stale the
    first time the markdown changes, and these documents change often - so this
    generates it instead, and -Check tells you when the committed HTML has fallen
    behind its source.

    Everything the documents here actually use is supported: headings, tables with a
    header row, fenced code blocks (including the indented ones inside numbered
    steps), blockquotes, ordered and unordered lists, horizontal rules, and inline
    code, bold, italic and links. Anything else is passed through as text rather than
    guessed at.

    The page is standalone: the CSS is embedded, so there is nothing to host and
    nothing to break when the file is copied somewhere else.

    To get it into IT Glue: open the .html in a browser, select all, copy, and paste
    into the IT Glue document editor. The editor keeps the headings, tables and code
    blocks; it drops the CSS, which is what you want there - IT Glue applies its own.

    Void elements are written self-closing (<hr/>, <br/>), so the page parses as XML
    as well as HTML - which is how the output is checked structurally rather than by
    looking at it.

.PARAMETER Path
    The markdown file to convert.

.PARAMETER Destination
    Where to write the HTML (default: the same folder and name, with .html).

.PARAMETER Title
    Browser title and heading of the page (default: the document's first # heading).

.PARAMETER Check
    Do not write. Compare with the HTML already on disk and exit 1 when they differ,
    so a stale page fails visibly instead of going unnoticed.

.EXAMPLE
    .\Convert-MarkdownToHtml.ps1 -Path ..\Device\Update-TeamsClient-ITGlue.md

    Writes ..\Device\Update-TeamsClient-ITGlue.html next to the source.

.EXAMPLE
    .\Convert-MarkdownToHtml.ps1 -Path ..\Device\Update-TeamsClient-ITGlue.md -Check

    Reports whether the committed HTML still matches the markdown. Exit code 1 when
    it does not.

.NOTES
    Author  : Sjoerd Kanon
    Platform: Windows PowerShell 5.1 and PowerShell 7+
#>
[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(Mandatory)] [string] $Path,
    [string] $Destination,
    [string] $Title,
    [switch] $Check
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $Path)) { throw "Markdown file not found: $Path" }
$source = Get-Item -LiteralPath $Path
if (-not $Destination) { $Destination = [IO.Path]::ChangeExtension($source.FullName, '.html') }

function Convert-Inline {
    <#
        Inline markup for one line of text. HTML is escaped first, then code spans are
        lifted out before bold, italic and links run - otherwise a `-Confirm:$false`
        or an *.md in a code span gets mangled into markup.
    #>
    param([string] $Text)

    $escaped = $Text -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;'

    $spans = [System.Collections.Generic.List[string]]::new()
    $escaped = [regex]::Replace($escaped, '`([^`]+)`', {
        param($match)
        $spans.Add($match.Groups[1].Value)
        return [char] 0xE000 + ($spans.Count - 1).ToString() + [char] 0xE001
    })

    $escaped = $escaped -replace '\*\*([^*]+)\*\*', '<strong>$1</strong>'
    $escaped = $escaped -replace '(?<![\*\w])\*([^*\n]+)\*(?!\*)', '<em>$1</em>'
    $escaped = [regex]::Replace($escaped, '\[([^\]]+)\]\(([^)]+)\)', {
        param($match)
        '<a href="{0}">{1}</a>' -f $match.Groups[2].Value, $match.Groups[1].Value
    })

    return [regex]::Replace($escaped, "$([char] 0xE000)(\d+)$([char] 0xE001)", {
        param($match)
        '<code>{0}</code>' -f $spans[[int] $match.Groups[1].Value]
    })
}

function Convert-MarkdownBody {
    <#
        Walk the document once, tracking the few block states these documents use.
        A list stays open across an indented code block so the numbering of a
        step-by-step procedure survives the command it tells you to run.
    #>
    param([string[]] $Lines)

    $html       = [System.Collections.Generic.List[string]]::new()
    $inCode     = $false
    $codeBuffer = [System.Collections.Generic.List[string]]::new()
    $tableRows  = [System.Collections.Generic.List[string[]]]::new()
    $quote      = [System.Collections.Generic.List[string]]::new()
    $listType   = $null      # 'ul' or 'ol' while a list is open
    $itemOpen   = $false

    function Close-Table {
        if ($tableRows.Count -eq 0) { return }
        $html.Add('<div class="table-scroll"><table>')
        $html.Add('<thead><tr>' + (($tableRows[0] | ForEach-Object { '<th>' + (Convert-Inline $_) + '</th>' }) -join '') + '</tr></thead>')
        if ($tableRows.Count -gt 1) {
            $html.Add('<tbody>')
            foreach ($row in $tableRows[1..($tableRows.Count - 1)]) {
                $html.Add('<tr>' + (($row | ForEach-Object { '<td>' + (Convert-Inline $_) + '</td>' }) -join '') + '</tr>')
            }
            $html.Add('</tbody>')
        }
        $html.Add('</table></div>')
        $tableRows.Clear()
    }
    function Close-Quote {
        if ($quote.Count -eq 0) { return }
        $html.Add('<blockquote>' + (($quote | ForEach-Object { Convert-Inline $_ }) -join '<br/>') + '</blockquote>')
        $quote.Clear()
    }
    function Close-List {
        # The caller clears $itemOpen: an assignment in here would only make a local
        # copy, which is the kind of thing that reads as working and is not.
        if (-not $listType) { return }
        if ($itemOpen) { $html.Add('</li>') }
        $html.Add("</$listType>")
    }

    for ($i = 0; $i -lt $Lines.Count; $i++) {
        $line    = $Lines[$i]
        $trimmed = $line.Trim()

        # -- fenced code ------------------------------------------------------
        if ($trimmed -match '^```') {
            if ($inCode) {
                $body = ($codeBuffer | ForEach-Object {
                    $_ -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;'
                }) -join "`n"
                $html.Add('<pre><code>' + $body + '</code></pre>')
                $codeBuffer.Clear()
                $inCode = $false
            } else {
                Close-Table; Close-Quote
                $inCode = $true
            }
            continue
        }
        if ($inCode) { $codeBuffer.Add($line); continue }

        # -- table ------------------------------------------------------------
        if ($trimmed.StartsWith('|')) {
            Close-Quote
            if ($trimmed -match '^\|[\s:|-]+\|$') { continue }   # the |---|---| separator
            $cells = $trimmed.Trim('|').Split('|') | ForEach-Object { $_.Trim() }
            $tableRows.Add(@($cells))
            continue
        }
        Close-Table

        # -- blockquote -------------------------------------------------------
        if ($trimmed -match '^>\s?(.*)$') { $quote.Add($Matches[1]); continue }
        Close-Quote

        # -- headings and rules ----------------------------------------------
        if ($trimmed -match '^(#{1,6})\s+(.*)$') {
            Close-List; $listType = $null; $itemOpen = $false
            $level = $Matches[1].Length
            $text  = Convert-Inline $Matches[2]
            $anchor = ($Matches[2].ToLower() -replace '[^a-z0-9\s-]', '' -replace '\s+', '-')
            $html.Add("<h$level id=`"$anchor`">$text</h$level>")
            continue
        }
        if ($trimmed -match '^(---+|\*\*\*+)$') {
            Close-List; $listType = $null; $itemOpen = $false
            $html.Add('<hr/>')
            continue
        }

        # -- lists ------------------------------------------------------------
        if ($trimmed -match '^[-*]\s+(.*)$' -or $trimmed -match '^\d+\.\s+(.*)$') {
            $wanted = if ($trimmed -match '^\d+\.') { 'ol' } else { 'ul' }
            $text   = Convert-Inline $Matches[1]
            if ($listType -ne $wanted) {
                Close-List
                $listType = $wanted
                $itemOpen = $false
                $html.Add("<$listType>")
            }
            if ($itemOpen) { $html.Add('</li>') }
            $html.Add("<li>$text")
            $itemOpen = $true
            continue
        }

        # -- blank and plain text --------------------------------------------
        if ($trimmed -eq '') {
            # A blank line inside a list can be followed by indented content that
            # belongs to the open item, so the list is not closed here.
            continue
        }
        if ($listType -and $line -match '^\s{2,}') {
            $html.Add('<p>' + (Convert-Inline $trimmed) + '</p>')
            continue
        }
        Close-List; $listType = $null; $itemOpen = $false
        $html.Add('<p>' + (Convert-Inline $trimmed) + '</p>')
    }

    Close-Table; Close-Quote; Close-List
    return $html
}

$lines = Get-Content -LiteralPath $source.FullName
if (-not $Title) {
    $heading = $lines | Where-Object { $_ -match '^#\s+(.+)$' } | Select-Object -First 1
    $Title = if ($heading) { ($heading -replace '^#\s+', '').Trim() } else { $source.BaseName }
}

$body = Convert-MarkdownBody -Lines $lines
$generated = Get-Date -Format 'yyyy-MM-dd HH:mm'

$css = @'
:root { color-scheme: light; }
* { box-sizing: border-box; }
body { margin: 0 auto; padding: 2.5rem 1.5rem 4rem; max-width: 62rem;
       font: 16px/1.65 -apple-system, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
       color: #1c2430; background: #ffffff; }
h1, h2, h3, h4 { line-height: 1.25; color: #10243c; margin: 2.2rem 0 .8rem; }
h1 { font-size: 1.9rem; margin-top: 0; padding-bottom: .5rem; border-bottom: 3px solid #0b5fa5; }
h2 { font-size: 1.4rem; margin-top: 2.6rem; padding-bottom: .3rem; border-bottom: 1px solid #d7dee7; }
h3 { font-size: 1.15rem; }
h4 { font-size: 1rem; text-transform: uppercase; letter-spacing: .04em; color: #45556a; }
p { margin: .7rem 0; }
a { color: #0b5fa5; }
hr { height: 0; margin: 2.4rem 0; border: 0; border-top: 1px solid #e3e8ef; }
code { padding: .12em .35em; border-radius: 3px; background: #eef2f7; color: #0f3255;
       font-family: Consolas, "SF Mono", Menlo, monospace; font-size: .89em; }
pre { margin: .9rem 0; padding: .9rem 1.1rem; overflow-x: auto; border-radius: 6px;
      background: #17222f; color: #e7eef6; }
pre code { padding: 0; background: none; color: inherit; font-size: .86em; line-height: 1.5; }
.table-scroll { overflow-x: auto; margin: 1rem 0; }
table { width: 100%; border-collapse: collapse; font-size: .94rem; }
th, td { padding: .55rem .7rem; text-align: left; vertical-align: top;
         border: 1px solid #dde3ea; }
th { background: #f2f6fa; font-weight: 600; color: #10243c; }
tbody tr:nth-child(even) { background: #fafcfe; }
blockquote { margin: 1.1rem 0; padding: .8rem 1.1rem; border-left: 4px solid #0b5fa5;
             border-radius: 0 4px 4px 0; background: #f3f8fc; }
blockquote code { background: #e2ecf5; }
ul, ol { margin: .7rem 0; padding-left: 1.6rem; }
li { margin: .35rem 0; }
li > p { margin: .4rem 0; }
.meta { margin: 0 0 2rem; padding: .6rem .9rem; border-radius: 4px;
        background: #f2f6fa; color: #45556a; font-size: .86rem; }
@media print {
  body { max-width: none; padding: 0; font-size: 11pt; }
  pre { background: #f4f6f9; color: #17222f; border: 1px solid #d7dee7; }
  h2 { page-break-after: avoid; }
  table, pre, blockquote { page-break-inside: avoid; }
}
'@

$html = @()
$html += '<!DOCTYPE html>'
$html += '<html lang="nl">'
$html += '<head>'
$html += '<meta charset="utf-8"/>'
$html += '<meta name="viewport" content="width=device-width, initial-scale=1"/>'
$html += "<title>$([System.Net.WebUtility]::HtmlEncode($Title))</title>"
$html += '<style>'
$html += $css
$html += '</style>'
$html += '</head>'
$html += '<body>'
$html += "<p class=`"meta`">Gegenereerd uit <code>$($source.Name)</code> op $generated. Wijzig de markdown, niet deze pagina: <code>Convert-MarkdownToHtml.ps1</code> maakt hem opnieuw.</p>"
$html += $body
$html += '</body>'
$html += '</html>'

$rendered = ($html -join "`r`n") + "`r`n"

if ($Check) {
    if (-not (Test-Path -LiteralPath $Destination)) {
        Write-Host "  [FAIL] $Destination does not exist yet" -ForegroundColor Red
        exit 1
    }
    # The generated-on line moves every run, so it is excluded from the comparison.
    $strip  = { param($Text) ($Text -split "`r?`n" | Where-Object { $_ -notmatch '^<p class="meta">' }) -join "`n" }
    $onDisk = & $strip (Get-Content -LiteralPath $Destination -Raw)
    if ((& $strip $rendered) -eq $onDisk) {
        Write-Host "  [ OK ] $([IO.Path]::GetFileName($Destination)) is up to date" -ForegroundColor Green
        exit 0
    }
    Write-Host "  [FAIL] $([IO.Path]::GetFileName($Destination)) is out of date - rerun without -Check" -ForegroundColor Red
    exit 1
}

if ($PSCmdlet.ShouldProcess($Destination, 'Write the generated HTML page')) {
    [IO.File]::WriteAllText($Destination, $rendered, [Text.UTF8Encoding]::new($false))
    $size = [math]::Round((Get-Item -LiteralPath $Destination).Length / 1KB, 1)
    Write-Host "  [ OK ] Wrote $Destination ($size KB)" -ForegroundColor Green
}
