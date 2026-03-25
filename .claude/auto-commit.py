#!/usr/bin/env python3
"""
Auto-commit hook: analyseert Edit/Write tool input + git diff
en schrijft een beschrijvend commit bericht voor PowerShell scripts.
"""
import sys
import json
import subprocess
import os
import re

REPO = '/Users/sjoerd/Github2.0/M365-Scripts'

def run(cmd):
    r = subprocess.run(cmd, shell=True, capture_output=True, text=True, cwd=REPO)
    return r.stdout.strip(), r.returncode

def extract_lines(text, prefix):
    return [l[1:] for l in (text or '').splitlines()
            if l.startswith(prefix) and not l.startswith(prefix * 3)]

def first_match(pattern, text, flags=re.I):
    m = re.search(pattern, text, flags)
    return m.group(1) if m else None

def nearest_section(content, changed_text):
    """Zoek de dichtstbijzijnde ## of ### header boven de gewijzigde tekst."""
    pos = content.find(changed_text)
    if pos == -1:
        return None
    before = content[:pos]
    headers = re.findall(r'^#{1,3} (.+)', before, re.M)
    return headers[-1].strip() if headers else None

def short_diff_summary(old, new):
    """Geef een korte samenvatting van wat er inhoudelijk veranderd is."""
    old_lines = set(l.strip() for l in old.splitlines() if l.strip())
    new_lines = set(l.strip() for l in new.splitlines() if l.strip())
    added   = [l for l in (new_lines - old_lines) if len(l) > 4]
    removed = [l for l in (old_lines - new_lines) if len(l) > 4]

    # Vermijd interne commentaar/technische regels
    def is_meaningful(l):
        return not re.match(r'^[#<{}\[\]()/*-]', l) and len(l) < 80

    added   = [l for l in added   if is_meaningful(l)]
    removed = [l for l in removed if is_meaningful(l)]

    if added and removed:
        return f"'{removed[0][:40]}' → '{added[0][:40]}'"
    if added:
        return f"'{added[0][:60]}' toegevoegd"
    if removed:
        return f"'{removed[0][:60]}' verwijderd"
    return None

# ---------- detectie op basis van old/new string (Edit tool) ----------

def detect_from_edit(old, new, basename):
    old = old or ''
    new = new or ''

    # New PowerShell function added or removed
    old_funcs = set(re.findall(r'function\s+([\w-]+)\s*[{(]', old, re.I))
    new_funcs = set(re.findall(r'function\s+([\w-]+)\s*[{(]', new, re.I))
    added = new_funcs - old_funcs
    removed = old_funcs - new_funcs
    if added and not removed:
        return f"{basename}: add function '{list(added)[0]}'"
    if removed and not added:
        return f"{basename}: remove function '{list(removed)[0]}'"
    if added and removed:
        return f"{basename}: rename function '{list(removed)[0]}' to '{list(added)[0]}'"

    # New parameter added
    old_params = set(re.findall(r'\[(?:string|switch|int|bool|array|object)\]\s*\$([\w]+)', old, re.I))
    new_params = set(re.findall(r'\[(?:string|switch|int|bool|array|object)\]\s*\$([\w]+)', new, re.I))
    added_params = new_params - old_params
    removed_params = old_params - new_params
    if added_params and not removed_params:
        return f"{basename}: add parameter '-{list(added_params)[0]}'"
    if removed_params and not added_params:
        return f"{basename}: remove parameter '-{list(removed_params)[0]}'"

    # Two-phase structure
    if 'Phase 1' in new and 'Phase 1' not in old:
        return f"{basename}: add two-phase approach (enumerate first, then retrieve data)"
    if 'Phase 2' in new and 'Phase 2' not in old:
        return f"{basename}: add phase 2"

    # New Graph API endpoint
    old_uris = set(re.findall(r'graph\.microsoft\.com/v[\d.]+/([\w/]+)', old))
    new_uris = set(re.findall(r'graph\.microsoft\.com/v[\d.]+/([\w/]+)', new))
    added_uris = new_uris - old_uris
    if added_uris:
        endpoint = list(added_uris)[0].split('/')[0]
        return f"{basename}: add Graph API call to '{endpoint}'"

    # Connect-* or module import
    old_conn = set(re.findall(r'(Connect-\w+|Import-Module\s+[\w.]+)', old, re.I))
    new_conn = set(re.findall(r'(Connect-\w+|Import-Module\s+[\w.]+)', new, re.I))
    added_conn = new_conn - old_conn
    if added_conn:
        return f"{basename}: add '{list(added_conn)[0]}'"

    # CSV export added
    if 'Export-Csv' in new and 'Export-Csv' not in old:
        return f"{basename}: add CSV export"

    # New section header
    new_headers = re.findall(r'Write-Host\s+".*?={3,}.*?"', new)
    old_headers = re.findall(r'Write-Host\s+".*?={3,}.*?"', old)
    if len(new_headers) > len(old_headers):
        return f"{basename}: add new output section"

    # Error handling added
    if 'try {' in new and 'try {' not in old:
        return f"{basename}: add error handling"

    # Readme changes
    if basename.lower() == 'readme.md':
        # New version history entry
        new_ver = re.findall(r'\|\s*(202\d-\d{2}-\d{2})\s*\|(.+?)(?:\||$)', new)
        old_ver = re.findall(r'\|\s*(202\d-\d{2}-\d{2})\s*\|(.+?)(?:\||$)', old)
        added_ver = [v for v in new_ver if v not in old_ver]
        if added_ver:
            entry = added_ver[0][1].strip()[:80]
            return f"readme: update version history — {entry}"

        # Find which section changed
        try:
            with open(f'{REPO}/readme.md', 'r') as f:
                full = f.read()
            section = nearest_section(full, new[:60]) if new else None
            if section:
                summary = short_diff_summary(old, new)
                if summary:
                    return f"readme '{section}': {summary}"
                return f"readme: update section '{section}'"
        except Exception:
            pass

        summary = short_diff_summary(old, new)
        if summary:
            return f"readme: {summary}"
        return "readme: general update"

    # Generic PowerShell: short diff summary
    summary = short_diff_summary(old, new)
    if summary:
        return f"{basename}: {summary}"
    return None  # fall through to diff-based detection

# ---------- detectie op basis van git diff ----------

def detect_from_diff(diff, basename):
    added   = extract_lines(diff, '+')
    removed = extract_lines(diff, '-')
    a = ' '.join(added)
    r = ' '.join(removed)

    # Functions
    new_funcs = set(re.findall(r'function\s+([\w-]+)\s*[{(]', a, re.I))
    old_funcs = set(re.findall(r'function\s+([\w-]+)\s*[{(]', r, re.I))
    if new_funcs - old_funcs:
        return f"{basename}: add function '{list(new_funcs - old_funcs)[0]}'"

    # Parameters
    new_params = set(re.findall(r'\[(?:string|switch|int)\]\s*\$([\w]+)', a, re.I))
    old_params = set(re.findall(r'\[(?:string|switch|int)\]\s*\$([\w]+)', r, re.I))
    if new_params - old_params:
        return f"{basename}: add parameter '-{list(new_params - old_params)[0]}'"

    # Graph endpoints
    new_uris = set(re.findall(r'graph\.microsoft\.com/v[\d.]+/([\w/]+)', a))
    old_uris = set(re.findall(r'graph\.microsoft\.com/v[\d.]+/([\w/]+)', r))
    if new_uris - old_uris:
        endpoint = list(new_uris - old_uris)[0].split('/')[0]
        return f"{basename}: add Graph API call to '{endpoint}'"

    # CSV export
    if 'Export-Csv' in a and 'Export-Csv' not in r:
        return f"{basename}: add CSV export"

    if basename.lower() == 'readme.md':
        meaningful = [l.strip() for l in added if l.strip() and not l.strip().startswith('#') and not l.strip().startswith('|') and len(l.strip()) > 5]
        if meaningful:
            return f"readme: add '{meaningful[0][:60]}'"
        return "readme: general update"

    # Fallback: line count + first meaningful addition
    n_add = len([l for l in added   if l.strip()])
    n_rem = len([l for l in removed if l.strip()])
    meaningful_add = [l.strip() for l in added if l.strip() and not re.match(r'^[#<{}\[\]()/*]', l.strip()) and len(l.strip()) > 5]
    meaningful_rem = [l.strip() for l in removed if l.strip() and not re.match(r'^[#<{}\[\]()/*]', l.strip()) and len(l.strip()) > 5]

    if n_rem > 0 and n_add == 0:
        snippet = f" — removed '{meaningful_rem[0][:50]}'" if meaningful_rem else ""
        return f"{basename}: remove code{snippet}"
    if n_rem > n_add * 2:
        return f"{basename}: clean up code"
    if n_add > 15:
        snippet = f" — add '{meaningful_add[0][:50]}'" if meaningful_add else ""
        return f"{basename}: add new functionality{snippet}"
    if n_add > 0:
        if meaningful_add and meaningful_rem:
            return f"{basename}: change '{meaningful_rem[0][:40]}' to '{meaningful_add[0][:40]}'"
        snippet = f" — '{meaningful_add[0][:50]}'" if meaningful_add else ""
        return f"{basename}: small update{snippet}"
    return f"{basename}: update"


# ---------- main ----------

def main():
    raw = sys.stdin.read()
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        return

    tool_input = data.get('tool_input', {})
    file_path  = tool_input.get('file_path', '')

    if not file_path:
        return
    if '/.claude/' in file_path:
        return
    if not file_path.startswith(REPO + '/'):
        return

    basename = os.path.basename(file_path)

    run(f'git add "{file_path}"')

    _, rc = run('git diff --cached --quiet')
    if rc == 0:
        return  # niets gestaged

    # Probeer eerst old_string/new_string analyse (Edit tool)
    old = tool_input.get('old_string', '')
    new = tool_input.get('new_string', '')
    message = None
    if old or new:
        message = detect_from_edit(old, new, basename)

    # Fallback: git diff analyse
    if not message:
        diff, _ = run('git diff --cached')
        message = detect_from_diff(diff, basename)

    _, commit_rc = run(f'git commit -m "{message}"')
    if commit_rc == 0:
        run('git push')


if __name__ == '__main__':
    main()
