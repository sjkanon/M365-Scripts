#Requires -Version 5.1
<#
.SYNOPSIS
    Quick fuzzy launcher for every script in this repo.

.DESCRIPTION
    Instead of navigating to a folder and typing .\Some-Script.ps1, type a few
    characters of the script name from anywhere:

        f dkim              -> runs scripts\Exchange\Test-DkimConfig.ps1
        f mailboxsizes      -> runs scripts\Exchange\Get-MailboxSizes.ps1
        f entra passkey     -> every term must match (name, folder or synopsis)

    Arguments after the search terms are passed straight to the script:

        f dkim -Domain contoso.com
        f copygroup -SourceGroup "Grp A" -TargetGroup "Grp B" -WhatIf

    When more than one script matches you get a numbered picker. With no search
    terms you get an interactive prompt.

    Run ".\f.ps1 -Install" once to register a global "f" command in your
    PowerShell profile so it works from any directory.

.PARAMETER Arguments
    Search terms followed by any arguments for the target script. Everything up
    to the first argument starting with "-" is treated as a search term.

.PARAMETER List
    Show matches only, do not run anything.

.PARAMETER Show
    Show path, synopsis and parameters of the match instead of running it.

.PARAMETER Edit
    Open the match in your editor ($env:EDITOR, VS Code, or notepad).

.PARAMETER Refresh
    Rebuild the cached script index.

.PARAMETER Install
    Add an "f" function to your PowerShell profile (also updates an existing one).

.PARAMETER Uninstall
    Remove the "f" function from your PowerShell profile.

.EXAMPLE
    .\f.ps1 -Install
    Registers "f" globally. After restarting the shell: f dkim

.EXAMPLE
    f -List mailbox
    Lists every script matching "mailbox" without running one.

.EXAMPLE
    f -Show trace
    Shows the parameters of Get-MessageTraceReport before you run it.

.NOTES
    The launcher's own switches (-List, -Show, -Edit, -Refresh, -Install,
    -Uninstall) are consumed here and never forwarded to the target script.
    Run such a script directly if it needs one of those parameter names.
#>
[CmdletBinding()]
param(
    [Parameter(ValueFromRemainingArguments = $true, Position = 0)]
    [object[]]$Arguments,

    [switch]$List,
    [switch]$Show,
    [switch]$Edit,
    [switch]$Refresh,
    [switch]$Install,
    [switch]$Uninstall
)

$ErrorActionPreference = 'Stop'
$ROOT       = $PSScriptRoot
$ScriptsDir = Join-Path $ROOT 'scripts'
$CacheFile  = Join-Path $ROOT '.f-index.json'

# -- Profile integration ------------------------------------------------------

$MarkerStart = '# >>> M365-Scripts f launcher >>>'
$MarkerEnd   = '# <<< M365-Scripts f launcher <<<'

function Set-FLauncher {
    param([bool]$Enable)

    $profilePath = $PROFILE.CurrentUserAllHosts
    $profileDir  = Split-Path $profilePath -Parent
    if (-not (Test-Path $profileDir))  { New-Item -ItemType Directory -Path $profileDir -Force | Out-Null }
    if (-not (Test-Path $profilePath)) { New-Item -ItemType File -Path $profilePath -Force | Out-Null }

    $existing = Get-Content -Path $profilePath -Raw -ErrorAction SilentlyContinue
    if ($null -eq $existing) { $existing = '' }

    # Strip any previous block first, so -Install doubles as an update.
    $pattern = [regex]::Escape($MarkerStart) + '.*?' + [regex]::Escape($MarkerEnd)
    $cleaned = [regex]::Replace($existing, $pattern, '', 'Singleline').TrimEnd()

    if (-not $Enable) {
        if ($cleaned -eq $existing.TrimEnd()) {
            Write-Host "  No f launcher found in $profilePath" -ForegroundColor DarkGray
        } else {
            Set-Content -Path $profilePath -Value $cleaned -Encoding UTF8
            Write-Host "  Removed f launcher from $profilePath" -ForegroundColor Green
            Write-Host "  Restart your shell to apply." -ForegroundColor DarkGray
        }
        return
    }

    $launcher = Join-Path $ROOT 'f.ps1'
    $block    = $MarkerStart + [Environment]::NewLine +
                'function f { & "' + $launcher + '" @args }' + [Environment]::NewLine +
                $MarkerEnd

    $newContent = ($cleaned + [Environment]::NewLine + [Environment]::NewLine + $block).TrimStart()
    Set-Content -Path $profilePath -Value $newContent -Encoding UTF8

    Write-Host ""
    Write-Host "  Installed 'f' in $profilePath" -ForegroundColor Green
    Write-Host '  Restart your shell (or run: . $PROFILE.CurrentUserAllHosts) and try: f dkim' -ForegroundColor DarkGray
    Write-Host ""
}

if ($Install)   { Set-FLauncher -Enable $true;  return }
if ($Uninstall) { Set-FLauncher -Enable $false; return }

# -- Index --------------------------------------------------------------------

function Get-ScriptSynopsis {
    param([string]$Path)

    try { $lines = Get-Content -Path $Path -TotalCount 60 -ErrorAction Stop }
    catch { return '' }

    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match '^\s*\.SYNOPSIS\s*$') {
            for ($j = $i + 1; $j -lt $lines.Count; $j++) {
                $text = $lines[$j].Trim()
                if ($text -eq '') { continue }
                if ($text -match '^\.[A-Z]' -or $text -match '^#>') { return '' }
                return $text
            }
        }
    }
    return ''
}

function New-ScriptIndex {
    $files = Get-ChildItem -Path $ScriptsDir -Recurse -Filter '*.ps1' -File | Sort-Object FullName

    $items = foreach ($file in $files) {
        $rel   = $file.FullName.Substring($ROOT.Length).TrimStart('\', '/')
        $parts = $rel -split '[\\/]'
        $category = if ($parts.Count -gt 2) { ($parts[1..($parts.Count - 2)]) -join '/' } else { 'root' }

        [pscustomobject]@{
            BaseName = $file.BaseName
            RelPath  = $rel
            FullName = $file.FullName
            Category = $category
            Synopsis = Get-ScriptSynopsis -Path $file.FullName
        }
    }

    return @($items)
}

function Get-IndexSignature {
    $files = Get-ChildItem -Path $ScriptsDir -Recurse -Filter '*.ps1' -File
    if (-not $files) { return '0|0' }
    $maxTicks = ($files | Measure-Object -Property LastWriteTimeUtc -Maximum).Maximum.Ticks
    return "$($files.Count)|$maxTicks"
}

function Get-ScriptIndex {
    param([switch]$Force)

    $signature = Get-IndexSignature

    if (-not $Force -and (Test-Path $CacheFile)) {
        try {
            $cache = Get-Content -Path $CacheFile -Raw | ConvertFrom-Json
            if ($cache.Signature -eq $signature) { return @($cache.Items) }
        } catch {
            # Corrupt or unreadable cache - fall through and rebuild.
        }
    }

    $items = New-ScriptIndex
    try {
        [pscustomobject]@{ Signature = $signature; Items = $items } |
            ConvertTo-Json -Depth 4 | Set-Content -Path $CacheFile -Encoding UTF8
    } catch {
        Write-Verbose "Could not write index cache: $($_.Exception.Message)"
    }

    return $items
}

# -- Matching -----------------------------------------------------------------

function Test-Subsequence {
    # True when every character of $Pattern occurs in $Text in order, and the
    # match is compact enough to be a plausible abbreviation rather than a
    # coincidence spread across a long name.
    param([string]$Text, [string]$Pattern)

    if ($Pattern.Length -lt 3) { return $false }

    $ti    = 0
    $first = -1
    $last  = -1

    foreach ($char in $Pattern.ToCharArray()) {
        $found = $false
        while ($ti -lt $Text.Length) {
            if ($Text[$ti] -eq $char) {
                if ($first -lt 0) { $first = $ti }
                $last  = $ti
                $found = $true
                $ti++
                break
            }
            $ti++
        }
        if (-not $found) { return $false }
    }

    $span = $last - $first + 1
    return ($span -le [Math]::Max($Pattern.Length * 2, $Pattern.Length + 4))
}

function Get-MatchScore {
    param($Item, [string[]]$Terms)

    $base = $Item.BaseName.ToLowerInvariant()
    $rel  = ($Item.RelPath.ToLowerInvariant()) -replace '[\\/]', ' '
    $syn  = "$($Item.Synopsis)".ToLowerInvariant()
    $flat = $base -replace '[-_ ]', ''

    $total = 0
    foreach ($term in $Terms) {
        $t     = $term.ToLowerInvariant()
        $tFlat = $t -replace '[-_ ]', ''
        $score = 0

        if     ($base -eq $t)                { $score = 1000 }
        elseif ($flat -eq $tFlat)            { $score = 900 }
        elseif ($base.StartsWith($t))        { $score = 600 }
        elseif ($base -match ('-' + [regex]::Escape($t))) { $score = 500 }
        elseif ($flat.Contains($tFlat))      { $score = 400 }
        elseif ($base.Contains($t))          { $score = 350 }
        elseif ($rel.Contains($t))           { $score = 200 }
        elseif ($syn.Contains($t))           { $score = 120 }
        elseif (Test-Subsequence -Text $flat -Pattern $tFlat) { $score = 60 }

        if ($score -eq 0) { return -1 }
        $total += $score
    }

    # Nudge shorter, shallower names to the top.
    $total -= $Item.BaseName.Length
    $total -= (($Item.RelPath -split '[\\/]').Count * 2)
    return $total
}

function Find-ScriptMatch {
    param($Index, [string[]]$Terms)

    if (-not $Terms -or $Terms.Count -eq 0) { return @($Index) }

    $scored = foreach ($item in $Index) {
        $score = Get-MatchScore -Item $item -Terms $Terms
        if ($score -ge 0) {
            [pscustomobject]@{ Item = $item; Score = $score }
        }
    }

    return @($scored | Sort-Object -Property Score -Descending | ForEach-Object { $_.Item })
}

# -- Display ------------------------------------------------------------------

function Write-ScriptList {
    param($Items, [switch]$Numbered)

    $n = 1
    foreach ($item in $Items) {
        $prefix = if ($Numbered) { '{0,3}. ' -f $n } else { '     ' }
        Write-Host ("  {0}{1,-42}" -f $prefix, $item.BaseName) -ForegroundColor White -NoNewline
        Write-Host (" {0}" -f $item.Category) -ForegroundColor DarkCyan
        if ($item.Synopsis) {
            Write-Host ("       {0}" -f $item.Synopsis) -ForegroundColor DarkGray
        }
        $n++
    }
}

function Show-ScriptDetail {
    param($Item)

    Write-Host ""
    Write-Host "  $($Item.BaseName)" -ForegroundColor Cyan
    Write-Host "  $($Item.RelPath)" -ForegroundColor DarkGray
    if ($Item.Synopsis) {
        Write-Host ""
        Write-Host "  $($Item.Synopsis)"
    }

    try {
        $command = Get-Command -Name $Item.FullName -ErrorAction Stop
        $common  = [System.Management.Automation.PSCmdlet]::CommonParameters +
                   [System.Management.Automation.PSCmdlet]::OptionalCommonParameters
        $params  = $command.Parameters.Keys | Where-Object { $common -notcontains $_ }

        if ($params) {
            Write-Host ""
            Write-Host "  Parameters:" -ForegroundColor Cyan
            foreach ($name in $params) {
                $p = $command.Parameters[$name]
                $isMandatory = $p.Attributes | Where-Object {
                    $_ -is [System.Management.Automation.ParameterAttribute] -and $_.Mandatory
                }
                $flag = if ($isMandatory) { ' (required)' } else { '' }
                Write-Host ("    -{0,-28} [{1}]{2}" -f $name, $p.ParameterType.Name, $flag) -ForegroundColor DarkGray
            }
        }
    } catch {
        Write-Host "  (could not read parameters: $($_.Exception.Message))" -ForegroundColor DarkYellow
    }
    Write-Host ""
}

function Open-InEditor {
    param($Item)

    $editors = @()
    if ($env:EDITOR) { $editors += $env:EDITOR }
    $editors += 'code'
    $editors += 'notepad'

    foreach ($editor in $editors) {
        $cmd = Get-Command $editor -ErrorAction SilentlyContinue
        if ($cmd) {
            Write-Host "  Opening $($Item.RelPath) in $editor" -ForegroundColor DarkGray
            & $cmd.Source $Item.FullName
            return
        }
    }
    Write-Host "  No editor found. Path: $($Item.FullName)" -ForegroundColor Yellow
}

function Resolve-ParameterName {
    # Map a typed parameter name (possibly abbreviated or an alias) onto the
    # script's real parameter name. Returns $null when it cannot be resolved.
    param($Command, [string]$Name)

    if (-not $Command) { return $null }

    foreach ($key in $Command.Parameters.Keys) {
        if ($key -eq $Name) { return $key }
    }
    foreach ($key in $Command.Parameters.Keys) {
        if ($Command.Parameters[$key].Aliases -contains $Name) { return $key }
    }

    $prefixed = @($Command.Parameters.Keys | Where-Object { $_ -like "$Name*" })
    if ($prefixed.Count -eq 1) { return $prefixed[0] }

    return $null
}

function ConvertTo-ArgumentSplat {
    # Turn a flat token list ("-Domain", "contoso.com", "-WhatIf") into a
    # hashtable of named parameters plus a list of positional values. Splatting
    # a plain array would pass "-Domain" as a positional value instead.
    param([object[]]$Tokens, [string]$ScriptPath)

    $named      = @{}
    $positional = @()

    $command = $null
    try { $command = Get-Command -Name $ScriptPath -ErrorAction Stop } catch { }

    $isName = {
        param($token)
        $t = "$token"
        return ($t.StartsWith('-') -and $t.Length -gt 1 -and $t -notmatch '^-[\d.]')
    }

    for ($i = 0; $i -lt $Tokens.Count; $i++) {
        $token = "$($Tokens[$i])"

        if (-not (& $isName $token)) {
            $positional += $Tokens[$i]
            continue
        }

        $name  = $token.TrimStart('-')
        $value = $null

        # -Name:value form (PowerShell sometimes pre-splits this into two tokens,
        # leaving a trailing colon behind - handled below).
        if ($name -match '^([^:]+):(.*)$') {
            $name  = $Matches[1]
            $value = $Matches[2]
            if ($value -eq '') { $value = $null }
        }

        $resolved = Resolve-ParameterName -Command $command -Name $name
        if ($resolved) { $name = $resolved }

        if ($null -ne $value) {
            $named[$name] = if ($value -eq 'true') { $true } elseif ($value -eq 'false') { $false } else { $value }
            continue
        }

        $isSwitch = $false
        if ($resolved -and $command.Parameters[$resolved].SwitchParameter) { $isSwitch = $true }

        $next    = if ($i + 1 -lt $Tokens.Count) { $Tokens[$i + 1] } else { $null }
        $hasNext = ($null -ne $next) -and -not (& $isName $next)

        if ($isSwitch) {
            # A switch only swallows an explicit boolean, so "-Apply false" and
            # "-Apply:false" both disable it instead of leaking a stray argument.
            if ($hasNext -and "$next" -match '^\$?(true|false)$') {
                $named[$name] = ("$next" -match 'true')
                $i++
            } else {
                $named[$name] = $true
            }
        } elseif ($hasNext) {
            $named[$name] = $next
            $i++
        } else {
            $named[$name] = $true
        }
    }

    return @{ Named = $named; Positional = $positional }
}

function Select-ScriptFromList {
    param($Items)

    Write-Host ""
    Write-Host "  $($Items.Count) matches:" -ForegroundColor Cyan
    Write-Host ""
    Write-ScriptList -Items $Items -Numbered
    Write-Host ""

    $answer = Read-Host "  Run which one? [1-$($Items.Count)], 0 to cancel"
    if ($answer -notmatch '^\d+$') { return $null }

    $choice = [int]$answer
    if ($choice -lt 1 -or $choice -gt $Items.Count) { return $null }
    return $Items[$choice - 1]
}

# -- Split search terms from pass-through arguments ---------------------------

$terms    = @()
$passThru = @()
$inArgs   = $false

foreach ($arg in @($Arguments)) {
    $text = "$arg"
    if (-not $inArgs -and $text.StartsWith('-') -and $text.Length -gt 1) { $inArgs = $true }
    if ($inArgs) { $passThru += $arg } else { $terms += $text }
}

# -- Main ---------------------------------------------------------------------

$index = Get-ScriptIndex -Force:$Refresh

if ($Refresh -and $terms.Count -eq 0 -and -not $List) {
    Write-Host "  Indexed $($index.Count) scripts." -ForegroundColor Green
    return
}

if ($terms.Count -eq 0 -and -not $List) {
    Write-Host ""
    Write-Host "  M365-Scripts launcher - $($index.Count) scripts indexed" -ForegroundColor Cyan
    Write-Host "  Type part of a script name, or leave blank to cancel." -ForegroundColor DarkGray
    Write-Host ""
    $searchInput = Read-Host "  Search"
    if ([string]::IsNullOrWhiteSpace($searchInput)) { return }
    $terms = @($searchInput -split '\s+' | Where-Object { $_ })
}

$found = Find-ScriptMatch -Index $index -Terms $terms

if ($found.Count -eq 0) {
    Write-Host ""
    Write-Host "  No script matches: $($terms -join ' ')" -ForegroundColor Yellow
    Write-Host "  Try fewer characters, or run 'f -List' to see everything." -ForegroundColor DarkGray
    Write-Host ""
    return
}

if ($List) {
    Write-Host ""
    Write-Host "  $($found.Count) script(s):" -ForegroundColor Cyan
    Write-Host ""
    Write-ScriptList -Items $found
    Write-Host ""
    return
}

$target = if ($found.Count -eq 1) { $found[0] } else { Select-ScriptFromList -Items $found }
if (-not $target) { return }

if ($Show) { Show-ScriptDetail -Item $target; return }
if ($Edit) { Open-InEditor     -Item $target; return }

Write-Host ""
Write-Host "  > $($target.RelPath) $($passThru -join ' ')" -ForegroundColor DarkGray
Write-Host ""

$splat = ConvertTo-ArgumentSplat -Tokens $passThru -ScriptPath $target.FullName
$named = $splat.Named
$rest  = $splat.Positional

& $target.FullName @named @rest
