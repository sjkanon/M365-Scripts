#Requires -Version 5.1
<#
.SYNOPSIS
    Check every link in the repository's markdown: files that must exist, and
    in-page anchors that must match a real heading.

.DESCRIPTION
    A dead link in a readme is invisible until someone clicks it, which is usually
    the moment they needed it. This walks every .md file and reports two kinds of
    breakage:

      - a relative link to a file or folder that is not there. Spaces are
        percent-encoded on GitHub, so "Time%20sync/readme.md" is decoded before the
        path is tested;
      - an anchor - "#set-usermanagerps1" - with no heading behind it. These rot
        silently: renaming a heading breaks every link to it, and nothing complains.
        Anchors are resolved the way GitHub builds them, so what passes here is what
        works on github.com.

    External links (http, https, mailto) are listed but not fetched: this is a
    structural check, not a crawler, and it has to work offline.

.PARAMETER Root
    Repository root (default: the folder two levels above this script).

.PARAMETER Path
    Check one file or folder instead of the whole repository.

.EXAMPLE
    .\Test-MarkdownLinks.ps1

    Check every markdown file in the repository. Exit 1 when anything is broken.

.EXAMPLE
    .\Test-MarkdownLinks.ps1 -Path scripts/Exchange

    Check only the readmes under one workload folder.

.NOTES
    Author  : Sjoerd Kanon
    Runs on Windows, macOS and Linux - it only reads files.
#>
[CmdletBinding()]
param(
    [string] $Root,
    [string] $Path
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not $Root) { $Root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot) }
$target = if ($Path) { if ([System.IO.Path]::IsPathRooted($Path)) { $Path } else { Join-Path $Root $Path } } else { $Root }
if (-not (Test-Path -LiteralPath $target)) { throw "No such path: $target" }

function Remove-InvisibleCharacter {
    <#
        Variation selectors and zero-width joiners - the bytes that make an emoji an
        emoji. A heading like "### ☁️ Azure Infrastructure" carries one, and so does
        the anchor a hand-written table of contents points at. They are invisible on
        both sides, so stripping them from both is what stops four working links from
        being reported as broken; it is not an attempt to reproduce GitHub's slugger
        byte for byte on characters nobody can see.
    #>
    param([Parameter(Mandatory)] [AllowEmptyString()] [string] $Text)
    return ($Text -replace '[︀-️​-‍⁠﻿]', '')
}

function ConvertTo-GitHubAnchor {
    <#
        GitHub's own rule for turning a heading into an anchor: lower-case it, drop
        everything that is not a letter, digit, space or hyphen, then turn spaces into
        hyphens. Markdown formatting inside the heading is stripped first, because
        "### `Name.ps1`" anchors as "nameps1" - the backticks and the dot are gone,
        which is exactly the kind of detail that gets written wrong by hand.
    #>
    param([Parameter(Mandatory)] [string] $Heading)

    $text = $Heading
    $text = Remove-InvisibleCharacter -Text $text
    $text = [regex]::Replace($text, '!?\[([^\]]*)\]\([^)]*\)', '$1')   # links -> their text
    $text = $text -replace '[`*_~]', ''                                # code, bold, italic
    $text = $text.Trim().ToLowerInvariant()
    $text = [regex]::Replace($text, '[^\p{L}\p{Nd} _-]', '')           # keep letters/digits/space/_/-
    return ($text -replace ' ', '-')
}

function Get-HeadingAnchor {
    <# Every anchor a file offers, with GitHub's -1/-2 suffix for repeated headings. #>
    # Not mandatory: an empty .md file legitimately has no lines, and a mandatory
    # [string[]] rejects that instead of simply offering no anchors.
    param([string[]] $Lines = @())

    $seen    = @{}
    $anchors = [System.Collections.Generic.HashSet[string]]::new()
    $fenced  = $false

    foreach ($line in $Lines) {
        # A "#" inside a fenced code block is a comment, not a heading.
        if ($line -match '^\s*(```|~~~)') { $fenced = -not $fenced; continue }
        if ($fenced) { continue }
        if ($line -notmatch '^(#{1,6})\s+(.*)$') { continue }

        $anchor = ConvertTo-GitHubAnchor -Heading $Matches[2]
        if ($anchor -eq '') { continue }
        if ($seen.ContainsKey($anchor)) {
            $seen[$anchor]++
            $null = $anchors.Add("$anchor-$($seen[$anchor])")
        } else {
            $seen[$anchor] = 0
            $null = $anchors.Add($anchor)
        }
    }
    return $anchors
}

$files = @(Get-ChildItem -LiteralPath $target -Recurse -Filter '*.md' -File |
           Where-Object { $_.FullName -notlike '*\.git\*' -and $_.FullName -notlike '*/.git/*' })

$anchorCache = @{}
function Get-AnchorsFor {
    param([Parameter(Mandatory)] [string] $FilePath)
    if (-not $anchorCache.ContainsKey($FilePath)) {
        $lines = @()
        try { $lines = @(Get-Content -LiteralPath $FilePath -ErrorAction Stop) } catch { }
        $lines = @($lines | Where-Object { $null -ne $_ })
        $anchorCache[$FilePath] = Get-HeadingAnchor -Lines $lines
    }
    return $anchorCache[$FilePath]
}

$broken  = [System.Collections.Generic.List[object]]::new()
$checked = 0
$external = 0

function Hide-CodeSpan {
    <#
        Blank out fenced blocks and inline code, keeping the length so nothing shifts.
        A readme that documents link syntax contains things that look exactly like
        links - "([docs](#...))" in a sentence about how these tables are built - and
        checking those reports the documentation as broken. Fenced blocks go first:
        they can hold backticks of their own.
    #>
    param([Parameter(Mandatory)] [AllowEmptyString()] [string] $Text)

    $blank = { param($m) ' ' * $m.Value.Length }
    $out = [regex]::Replace($Text, '(?ms)^\s*(```|~~~).*?^\s*\1[^\r\n]*$', $blank)
    return [regex]::Replace($out, '`[^`\r\n]*`', $blank)
}

foreach ($file in $files) {
    $text = Hide-CodeSpan -Text (Get-Content -LiteralPath $file.FullName -Raw)
    $relFile = $file.FullName.Substring($Root.Length).TrimStart('\', '/')

    foreach ($match in [regex]::Matches($text, '\[(?<text>[^\]]*)\]\((?<url>[^)\s]+)(?:\s+"[^"]*")?\)')) {
        $url = $match.Groups['url'].Value
        if ($url -match '^(https?:|mailto:|tel:)') { $external++; continue }

        $checked++
        $pathPart   = ($url -split '#', 2)[0]
        $anchorPart = if ($url -match '#') { ($url -split '#', 2)[1] } else { '' }

        # Which file does this link land in - this one, or another?
        if ($pathPart -eq '') {
            $targetFile = $file.FullName
        } else {
            $decoded = [Uri]::UnescapeDataString($pathPart)
            $targetFile = Join-Path $file.DirectoryName $decoded
            if (-not (Test-Path -LiteralPath $targetFile)) {
                $broken.Add([PSCustomObject]@{ File = $relFile; Link = $url; Reason = 'no such file or folder' })
                continue
            }
            $targetFile = (Resolve-Path -LiteralPath $targetFile).Path
        }

        if ($anchorPart -eq '') { continue }
        # Only markdown carries headings; an anchor on anything else is not ours to judge.
        if ([System.IO.Path]::GetExtension($targetFile) -ne '.md') { continue }

        $anchor = (Remove-InvisibleCharacter -Text ([Uri]::UnescapeDataString($anchorPart))).ToLowerInvariant()
        if (-not (Get-AnchorsFor -FilePath $targetFile).Contains($anchor)) {
            $where = if ($pathPart -eq '') { 'this file' } else { $pathPart }
            $broken.Add([PSCustomObject]@{ File = $relFile; Link = $url; Reason = "no heading '$anchor' in $where" })
        }
    }
}

Write-Host ''
Write-Host ("  {0} markdown file(s), {1} internal link(s) checked, {2} external link(s) skipped." -f
            $files.Count, $checked, $external) -ForegroundColor Cyan

if ($broken.Count -eq 0) {
    Write-Host '  Every internal link resolves.' -ForegroundColor Green
    Write-Host ''
    exit 0
}

Write-Host ''
foreach ($item in ($broken | Sort-Object File, Link)) {
    Write-Host ("  BROKEN  {0}" -f $item.File) -ForegroundColor Yellow
    Write-Host ("          -> {0}   ({1})" -f $item.Link, $item.Reason) -ForegroundColor DarkGray
}
Write-Host ''
Write-Host ("  {0} broken link(s)." -f $broken.Count) -ForegroundColor Red
Write-Host ''
exit 1
