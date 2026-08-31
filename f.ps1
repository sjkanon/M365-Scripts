#Requires -Version 5.1
<#
.SYNOPSIS
    Short commands for every script in this repo. Load from your PowerShell profile.

.DESCRIPTION
    Creates one function per script in scripts\, so you no longer have to navigate
    to a folder and type .\Some-Script.ps1:

        f-test-dkimconfig -Domain contoso.com
        f-dkimconfig -Domain contoso.com          # short form when the noun is unique
        f-get-mailboxsizes

    The functions are GENERATED from scripts\, not maintained by hand. Add a script
    and there is a command for it in the next shell. A hand-kept list would be a
    second place that can drift.

    Next to the per-script commands there is 'f', a fuzzy search for when you know
    roughly what a script is called but not exactly:

        f dkim                    # one match -> runs it
        f entra group             # several matches -> numbered picker
        f -List mailbox           # show matches, run nothing
        f -Show trace             # path, synopsis and parameters
        f-m365                    # the whole catalogue

    Install - once, dot-sourced:

        notepad $PROFILE
        . "C:\Users\<you>\Git\M365-Scripts\f.ps1"

    Note the leading dot. Without dot-sourcing this file runs in its own scope and
    the commands are gone again immediately - without an error.

    Or let it write that line for you:  .\f.ps1 -Install

    The generated commands are cached in .f-index.json (gitignored). Parsing all
    scripts costs about half a second, reading the cache about 30 ms, so the cache
    is what keeps your shell start quick. It refreshes itself as soon as a script
    is added, removed or changed; 'f-refresh' forces it.

.PARAMETER Arguments
    Only used when the file is run instead of dot-sourced: search terms followed by
    arguments for the target script, same as the 'f' function.

.PARAMETER Install
    Add the dot-source line to $PROFILE.

.PARAMETER Uninstall
    Remove the dot-source line from $PROFILE.

.EXAMPLE
    .\f.ps1 -Install
    Wires this file into your profile. After restarting your shell: f-m365

.EXAMPLE
    f -Show trace
    Shows the parameters of Get-MessageTraceReport before you run it.

.NOTES
    WHY THE WRAPPERS COPY THE PARAMETER BLOCK. A function with only @args has no
    parameter metadata, so PowerShell cannot complete -Domain and you only hear
    about a typo when the script runs. The wrappers therefore take the parameter
    block of the target script from its AST. Completion is then correct by
    construction, also for a parameter added tomorrow.

    Default values are deliberately dropped from the wrapper: they are never
    forwarded (only bound parameters are), and a default that calls out to
    something would otherwise be evaluated on every invocation of the wrapper.
#>
[CmdletBinding()]
param(
    [Parameter(ValueFromRemainingArguments = $true, Position = 0)]
    [object[]]$Arguments,

    [switch]$List,
    [switch]$Show,
    [switch]$Edit,
    [switch]$Install,
    [switch]$Uninstall
)

# Global, not script-scoped: the f- commands are global, so the state they read
# has to outlive this file's own scope when it is run instead of dot-sourced.
$global:FLauncherRoot = $PSScriptRoot
$global:FLauncherPath = Join-Path $PSScriptRoot 'f.ps1'
$global:FScriptsDir   = Join-Path $PSScriptRoot 'scripts'
$global:FCacheFile    = Join-Path $PSScriptRoot '.f-index.json'

# ── Index ─────────────────────────────────────────────────────────────────────

function global:Get-FScriptSynopsis {
    param([string]$Path, [string[]]$Lines)

    for ($i = 0; $i -lt $Lines.Count; $i++) {
        if ($Lines[$i] -match '^\s*\.SYNOPSIS\s*$') {
            for ($j = $i + 1; $j -lt $Lines.Count; $j++) {
                $text = $Lines[$j].Trim()
                if ($text -eq '') { continue }
                if ($text -match '^\.[A-Z]' -or $text -match '^#>') { return '' }
                return $text
            }
        }
    }
    return ''
}

function global:Get-FProxyBody {
    <#
        Builds the body of the wrapper function: the parameter block of the target
        script, without default values, forwarding only what the caller actually
        bound so the script's own defaults still apply.
    #>
    param([string]$Path)

    $quoted = "'" + ($Path -replace "'", "''") + "'"

    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$errors)

    if ($errors -or -not $ast -or -not $ast.ParamBlock) {
        # No parameter block (or unparseable): a plain pass-through still beats
        # typing the path, it just cannot complete parameters.
        return "& $quoted @args"
    }

    $attributes = ($ast.ParamBlock.Attributes | ForEach-Object { $_.Extent.Text }) -join [Environment]::NewLine

    $parameters = $ast.ParamBlock.Parameters | ForEach-Object {
        $attrs = ($_.Attributes | ForEach-Object { $_.Extent.Text }) -join ' '
        $name  = '$' + $_.Name.VariablePath.UserPath
        if ($attrs) { "$attrs $name" } else { $name }
    }

    if (-not $parameters) { return "& $quoted @args" }

    return @(
        $attributes
        'param('
        ($parameters -join ("," + [Environment]::NewLine))
        ')'
        ''
        "& $quoted @PSBoundParameters"
    ) -join [Environment]::NewLine
}

function global:New-FScriptIndex {
    $files = Get-ChildItem -Path $global:FScriptsDir -Recurse -Filter '*.ps1' -File |
             Sort-Object { ($_.FullName -split '[\\/]').Count }, FullName

    $items = New-Object System.Collections.ArrayList
    $taken = @{}

    foreach ($file in $files) {
        $rel      = $file.FullName.Substring($global:FLauncherRoot.Length).TrimStart('\', '/')
        $parts    = $rel -split '[\\/]'
        $category = if ($parts.Count -gt 2) { ($parts[1..($parts.Count - 2)]) -join '/' } else { 'root' }

        $lines = @()
        try { $lines = @(Get-Content -Path $file.FullName -TotalCount 60 -ErrorAction Stop) } catch { }

        # Full name first: stripping the verb collides 11 times here (Detect-,
        # Install- and Uninstall-ClaudeDesktop-Intune all become the same noun),
        # and those are genuinely different actions.
        $command = 'f-' + $file.BaseName.ToLowerInvariant()
        if ($taken.ContainsKey($command)) {
            $command = 'f-' + $parts[1].ToLowerInvariant() + '-' + $file.BaseName.ToLowerInvariant()
        }
        $taken[$command] = $true

        $null = $items.Add([pscustomobject]@{
            BaseName = $file.BaseName
            RelPath  = $rel
            FullName = $file.FullName
            Category = $category
            Synopsis = Get-FScriptSynopsis -Path $file.FullName -Lines $lines
            Command  = $command
            Short    = ''
            Body     = Get-FProxyBody -Path $file.FullName
        })
    }

    # Short form (verb dropped) only where it stays unambiguous, so f-dkimconfig
    # works but f-claudedesktop-intune - which would be three scripts - does not.
    $shorts = @{}
    foreach ($item in $items) {
        $noun = $item.BaseName -replace '^[A-Za-z]+-', ''
        if ($noun -eq $item.BaseName) { continue }
        $short = 'f-' + $noun.ToLowerInvariant()
        if ($taken.ContainsKey($short)) { continue }
        if ($shorts.ContainsKey($short)) { $shorts[$short] = $null } else { $shorts[$short] = $item }
    }
    foreach ($short in $shorts.Keys) {
        if ($shorts[$short]) { $shorts[$short].Short = $short }
    }

    return @($items)
}

function global:Get-FIndexSignature {
    $files = @(Get-ChildItem -Path $global:FScriptsDir -Recurse -Filter '*.ps1' -File)
    if ($files.Count -eq 0) { return '0|0' }
    $maxTicks = ($files | Measure-Object -Property LastWriteTimeUtc -Maximum).Maximum.Ticks
    return "$($files.Count)|$maxTicks"
}

function global:Get-FScriptIndex {
    param([switch]$Force)

    if (-not (Test-Path $global:FScriptsDir)) { return @() }

    $signature = Get-FIndexSignature

    if (-not $Force -and (Test-Path $global:FCacheFile)) {
        try {
            $cache = Get-Content -Path $global:FCacheFile -Raw | ConvertFrom-Json
            if ($cache.Signature -eq $signature) { return @($cache.Items) }
        } catch {
            # Corrupt or unreadable cache - fall through and rebuild.
        }
    }

    $items = New-FScriptIndex
    try {
        [pscustomobject]@{ Signature = $signature; Items = $items } |
            ConvertTo-Json -Depth 4 | Set-Content -Path $global:FCacheFile -Encoding UTF8
    } catch {
        Write-Verbose "Could not write index cache: $($_.Exception.Message)"
    }

    return $items
}

# ── Matching ──────────────────────────────────────────────────────────────────

function global:Test-FSubsequence {
    # True when every character of $Pattern occurs in $Text in order, and the match
    # is compact enough to be a plausible abbreviation rather than a coincidence
    # spread across a long name.
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

function global:Get-FMatchScore {
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
        elseif (Test-FSubsequence -Text $flat -Pattern $tFlat) { $score = 60 }

        if ($score -eq 0) { return -1 }
        $total += $score
    }

    # Nudge shorter, shallower names to the top.
    $total -= $Item.BaseName.Length
    $total -= (($Item.RelPath -split '[\\/]').Count * 2)
    return $total
}

function global:Find-FScriptMatch {
    param($Index, [string[]]$Terms)

    if (-not $Terms -or $Terms.Count -eq 0) { return @($Index) }

    $scored = foreach ($item in $Index) {
        $score = Get-FMatchScore -Item $item -Terms $Terms
        if ($score -ge 0) {
            [pscustomobject]@{ Item = $item; Score = $score }
        }
    }

    return @($scored | Sort-Object -Property Score -Descending | ForEach-Object { $_.Item })
}

# ── Argument handling ─────────────────────────────────────────────────────────

function global:Resolve-FParameterName {
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

function global:ConvertTo-FArgumentSplat {
    # Turn a flat token list ("-Domain", "contoso.com", "-WhatIf") into a hashtable
    # of named parameters plus a list of positional values. Splatting a plain array
    # would pass "-Domain" as a positional value instead.
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

        $resolved = Resolve-FParameterName -Command $command -Name $name
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

# ── Display ───────────────────────────────────────────────────────────────────

function global:Write-FScriptList {
    param($Items, [switch]$Numbered)

    $n = 1
    foreach ($item in $Items) {
        $prefix = if ($Numbered) { '{0,3}. ' -f $n } else { '     ' }
        $name   = if ($item.Short) { $item.Short } else { $item.Command }
        Write-Host ("  {0}{1,-40}" -f $prefix, $name) -ForegroundColor White -NoNewline
        Write-Host (" {0}" -f $item.Category) -ForegroundColor DarkCyan
        if ($item.Synopsis) {
            Write-Host ("       {0}" -f $item.Synopsis) -ForegroundColor DarkGray
        }
        $n++
    }
}

function global:Show-FScriptDetail {
    param($Item)

    Write-Host ""
    Write-Host "  $($Item.BaseName)" -ForegroundColor Cyan
    Write-Host "  $($Item.RelPath)" -ForegroundColor DarkGray
    Write-Host "  command: $($Item.Command)$(if ($Item.Short) { "  /  $($Item.Short)" })" -ForegroundColor DarkGray
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

function global:Open-FInEditor {
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

function global:Select-FScriptFromList {
    param($Items)

    Write-Host ""
    Write-Host "  $($Items.Count) matches:" -ForegroundColor Cyan
    Write-Host ""
    Write-FScriptList -Items $Items -Numbered
    Write-Host ""

    $answer = Read-Host "  Run which one? [1-$($Items.Count)], 0 to cancel"
    if ($answer -notmatch '^\d+$') { return $null }

    $choice = [int]$answer
    if ($choice -lt 1 -or $choice -gt $Items.Count) { return $null }
    return $Items[$choice - 1]
}

# ── Registration ──────────────────────────────────────────────────────────────

function global:Register-FCommands {
    <#
        Never overwrites a command that belongs to something else. Another tool
        loaded from the same profile can own f- names too (ScriptRunner does), and
        silently replacing one of those would break it without a word. Names we
        registered ourselves are fair game, so reloading stays possible.
    #>
    param($Index)

    if (-not $global:FRegisteredCommands) { $global:FRegisteredCommands = @{} }

    # One enumeration instead of a Get-Command per name: this runs 280 times.
    $existing = New-Object 'System.Collections.Generic.HashSet[string]' (
        [string[]]@(Get-ChildItem -Path function: | ForEach-Object { $_.Name }),
        [StringComparer]::OrdinalIgnoreCase)

    $skipped = New-Object System.Collections.ArrayList

    foreach ($item in $Index) {
        $block = [scriptblock]::Create($item.Body)

        foreach ($name in @($item.Command, $item.Short)) {
            if (-not $name) { continue }
            if ($existing.Contains($name) -and -not $global:FRegisteredCommands.ContainsKey($name)) {
                $null = $skipped.Add($name)
                continue
            }
            Set-Item -Path "function:global:$name" -Value $block
            $global:FRegisteredCommands[$name] = $true
        }
    }

    return $skipped
}

function global:f-refresh {
    <#
    .SYNOPSIS
        Rebuilds the script index and reloads all f- commands.
    #>
    [CmdletBinding()]
    param()

    $global:FScriptIndex = Get-FScriptIndex -Force
    Register-FCommands -Index $global:FScriptIndex
    Write-Host "  Indexed $($global:FScriptIndex.Count) scripts." -ForegroundColor Green
}

function global:f-m365 {
    <#
    .SYNOPSIS
        Shows the catalogue and the f- command that belongs to each script.

    .DESCRIPTION
        Named f-m365 and not f-scripts because a profile can load more than one of
        these: ScriptRunner in itce-testing already owns f-scripts, and taking that
        name would break it.

    .EXAMPLE
        f-m365
        f-m365 mailbox
    #>
    [CmdletBinding()]
    param([Parameter(Position = 0)][string]$Filter)

    $items = $global:FScriptIndex
    if ($Filter) {
        $items = @($items | Where-Object {
            $_.Command -like "*$Filter*" -or $_.Short -like "*$Filter*" -or
            $_.RelPath -like "*$Filter*" -or $_.Synopsis -like "*$Filter*"
        })
    }

    $items |
        Sort-Object Category, BaseName |
        Select-Object @{ n = 'Command'; e = { if ($_.Short) { $_.Short } else { $_.Command } } },
                      @{ n = 'Full';    e = { $_.Command } },
                      @{ n = 'Script';  e = { $_.BaseName } },
                      Category
}

# Declared unconditionally - 'f' is the headline command - but say so if it was
# already someone else's, rather than shadowing it in silence.
if (-not $global:FRegisteredCommands) { $global:FRegisteredCommands = @{} }
if ((Test-Path -LiteralPath 'function:f') -and -not $global:FRegisteredCommands.ContainsKey('f')) {
    Write-Warning "M365-Scripts replaced an existing 'f' command."
}
$global:FRegisteredCommands['f'] = $true

function global:f {
    <#
    .SYNOPSIS
        Fuzzy search across every script in M365-Scripts, then run it.

    .DESCRIPTION
        For when you know roughly what a script is called but not exactly. Matches
        on name, folder and .SYNOPSIS. One match runs straight away, several give a
        numbered picker. Arguments after the search terms go to the script.

        Use the generated f-<name> commands when you do know the name - those have
        real parameter completion.

    .EXAMPLE
        f dkim -Domain contoso.com

    .EXAMPLE
        f -List mailbox
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromRemainingArguments = $true, Position = 0)]
        [object[]]$Arguments,

        [switch]$List,
        [switch]$Show,
        [switch]$Edit
    )

    $terms    = @()
    $passThru = @()
    $inArgs   = $false

    foreach ($arg in @($Arguments)) {
        $text = "$arg"
        if (-not $inArgs -and $text.StartsWith('-') -and $text.Length -gt 1) { $inArgs = $true }
        if ($inArgs) { $passThru += $arg } else { $terms += $text }
    }

    $index = $global:FScriptIndex

    if ($terms.Count -eq 0 -and -not $List) {
        Write-Host ""
        Write-Host "  M365-Scripts - $($index.Count) scripts indexed" -ForegroundColor Cyan
        Write-Host "  Type part of a script name, or leave blank to cancel." -ForegroundColor DarkGray
        Write-Host ""
        $searchInput = Read-Host "  Search"
        if ([string]::IsNullOrWhiteSpace($searchInput)) { return }
        $terms = @($searchInput -split '\s+' | Where-Object { $_ })
    }

    $found = Find-FScriptMatch -Index $index -Terms $terms

    if ($found.Count -eq 0) {
        Write-Host ""
        Write-Host "  No script matches: $($terms -join ' ')" -ForegroundColor Yellow
        Write-Host "  Try fewer characters, or run 'f-m365' to see everything." -ForegroundColor DarkGray
        Write-Host ""
        return
    }

    if ($List) {
        Write-Host ""
        Write-Host "  $($found.Count) script(s):" -ForegroundColor Cyan
        Write-Host ""
        Write-FScriptList -Items $found
        Write-Host ""
        return
    }

    $target = if ($found.Count -eq 1) { $found[0] } else { Select-FScriptFromList -Items $found }
    if (-not $target) { return }

    if ($Show) { Show-FScriptDetail -Item $target; return }
    if ($Edit) { Open-FInEditor     -Item $target; return }

    Write-Host ""
    Write-Host "  > $($target.RelPath) $($passThru -join ' ')" -ForegroundColor DarkGray
    Write-Host ""

    $splat = ConvertTo-FArgumentSplat -Tokens $passThru -ScriptPath $target.FullName
    $named = $splat.Named
    $rest  = $splat.Positional

    & $target.FullName @named @rest
}

# ── Profile integration ───────────────────────────────────────────────────────

function global:Set-FLauncher {
    param([bool]$Enable)

    $markerStart = '# >>> M365-Scripts f commands >>>'
    $markerEnd   = '# <<< M365-Scripts f commands <<<'

    $profilePath = $PROFILE.CurrentUserCurrentHost
    $profileDir  = Split-Path $profilePath -Parent
    if (-not (Test-Path $profileDir))  { New-Item -ItemType Directory -Path $profileDir -Force | Out-Null }
    if (-not (Test-Path $profilePath)) { New-Item -ItemType File -Path $profilePath -Force | Out-Null }

    $existing = Get-Content -Path $profilePath -Raw -ErrorAction SilentlyContinue
    if ($null -eq $existing) { $existing = '' }

    # Strip any previous block first, so -Install doubles as an update.
    $pattern = [regex]::Escape($markerStart) + '.*?' + [regex]::Escape($markerEnd)
    $cleaned = [regex]::Replace($existing, $pattern, '', 'Singleline').TrimEnd()

    if (-not $Enable) {
        if ($cleaned -eq $existing.TrimEnd()) {
            Write-Host "  No marked f block found in $profilePath" -ForegroundColor DarkGray
            if ($cleaned -match [regex]::Escape($global:FLauncherPath)) {
                Write-Host "  There is a hand-written dot-source line for f.ps1 - remove that yourself." -ForegroundColor Yellow
            }
        } else {
            Set-Content -Path $profilePath -Value $cleaned -Encoding UTF8
            Write-Host "  Removed f commands from $profilePath" -ForegroundColor Green
            Write-Host "  Restart your shell to apply." -ForegroundColor DarkGray
        }
        return
    }

    if ($cleaned -match [regex]::Escape($global:FLauncherPath)) {
        Write-Host ""
        Write-Host "  Already dot-sourced from $profilePath - nothing to do." -ForegroundColor DarkGray
        Write-Host ""
        return
    }

    $block = $markerStart + [Environment]::NewLine +
             '. "' + $global:FLauncherPath + '"' + [Environment]::NewLine +
             $markerEnd

    $newContent = ($cleaned + [Environment]::NewLine + [Environment]::NewLine + $block).TrimStart()
    Set-Content -Path $profilePath -Value $newContent -Encoding UTF8

    Write-Host ""
    Write-Host "  Added f commands to $profilePath" -ForegroundColor Green
    Write-Host '  Restart your shell (or run: . $PROFILE) and try: f-m365' -ForegroundColor DarkGray
    Write-Host ""
}

# ── Load ──────────────────────────────────────────────────────────────────────

$FDotSourced = ($MyInvocation.InvocationName -eq '.')

if ($Install -or $Uninstall) {
    Set-FLauncher -Enable ([bool]$Install)
    return
}

$global:FScriptIndex = Get-FScriptIndex
$FSkipped = Register-FCommands -Index $global:FScriptIndex

if ($FDotSourced) {
    Write-Host "M365-Scripts loaded - $($global:FScriptIndex.Count) commands. 'f-m365' lists them, 'f <term>' searches." -ForegroundColor DarkGray
    if ($FSkipped.Count -gt 0) {
        Write-Host "  $($FSkipped.Count) name(s) already taken by another tool, left alone: $($FSkipped -join ', ')" -ForegroundColor DarkYellow
    }
} else {
    # Run instead of dot-sourced: behave like the f function so .\f.ps1 dkim works.
    f @Arguments -List:$List -Show:$Show -Edit:$Edit
}
