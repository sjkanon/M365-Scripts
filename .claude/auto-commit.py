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

# ---------- detectie op basis van old/new string (Edit tool) ----------

def detect_from_edit(old, new, basename):
    old = old or ''
    new = new or ''

    # Nieuwe PowerShell functie toegevoegd
    old_funcs = set(re.findall(r'function\s+([\w-]+)\s*[{(]', old, re.I))
    new_funcs = set(re.findall(r'function\s+([\w-]+)\s*[{(]', new, re.I))
    added = new_funcs - old_funcs
    removed = old_funcs - new_funcs
    if added and not removed:
        return f"Functie '{list(added)[0]}' toegevoegd aan {basename}"
    if removed and not added:
        return f"Functie '{list(removed)[0]}' verwijderd uit {basename}"

    # Nieuwe parameter toegevoegd
    old_params = set(re.findall(r'\[(?:string|switch|int|bool|array|object)\]\s*\$([\w]+)', old, re.I))
    new_params = set(re.findall(r'\[(?:string|switch|int|bool|array|object)\]\s*\$([\w]+)', new, re.I))
    added_params = new_params - old_params
    if added_params:
        param = list(added_params)[0]
        return f"Parameter '-{param}' toegevoegd aan {basename}"

    # Twee-fasen structuur toegevoegd
    if 'Phase 1' in new and 'Phase 1' not in old:
        return f"Twee-fasen aanpak toegevoegd aan {basename}"
    if 'Phase 2' in new and 'Phase 2' not in old:
        return f"Fase 2 toegevoegd aan {basename}"

    # Nieuwe Graph API endpoint
    old_uris = set(re.findall(r'graph\.microsoft\.com/v[\d.]+/([\w/]+)', old))
    new_uris = set(re.findall(r'graph\.microsoft\.com/v[\d.]+/([\w/]+)', new))
    added_uris = new_uris - old_uris
    if added_uris:
        endpoint = list(added_uris)[0].split('/')[0]
        return f"Graph API endpoint '{endpoint}' toegevoegd aan {basename}"

    # Connect-* of module import
    old_conn = set(re.findall(r'(Connect-\w+|Import-Module\s+[\w.]+)', old, re.I))
    new_conn = set(re.findall(r'(Connect-\w+|Import-Module\s+[\w.]+)', new, re.I))
    added_conn = new_conn - old_conn
    if added_conn:
        return f"'{list(added_conn)[0]}' toegevoegd aan {basename}"

    # Export / CSV output toegevoegd
    if 'Export-Csv' in new and 'Export-Csv' not in old:
        return f"CSV export toegevoegd aan {basename}"

    # Write-Host sectie / fase header toegevoegd
    new_headers = re.findall(r'Write-Host\s+".*?={3,}.*?"', new)
    old_headers = re.findall(r'Write-Host\s+".*?={3,}.*?"', old)
    if len(new_headers) > len(old_headers):
        return f"Nieuwe sectie toegevoegd aan {basename}"

    # Foutafhandeling toegevoegd
    if 'try {' in new and 'try {' not in old:
        return f"Foutafhandeling toegevoegd aan {basename}"

    # Versiegeschiedenis bijgewerkt in readme
    if basename.lower() == 'readme.md':
        m = re.search(r'\|\s*(202\d-\d{2}-\d{2})\s*\|(.+?)(?:\||$)', new)
        if m:
            entry = m.group(2).strip()[:80]
            return f"Versiegeschiedenis bijgewerkt: {entry}"
        return "Readme bijgewerkt"

    return None

# ---------- detectie op basis van git diff ----------

def detect_from_diff(diff, basename):
    added   = extract_lines(diff, '+')
    removed = extract_lines(diff, '-')
    a = ' '.join(added)
    r = ' '.join(removed)

    # Functies
    new_funcs = set(re.findall(r'function\s+([\w-]+)\s*[{(]', a, re.I))
    old_funcs = set(re.findall(r'function\s+([\w-]+)\s*[{(]', r, re.I))
    if new_funcs - old_funcs:
        return f"Functie '{list(new_funcs - old_funcs)[0]}' toegevoegd aan {basename}"

    # Parameters
    new_params = set(re.findall(r'\[(?:string|switch|int)\]\s*\$([\w]+)', a, re.I))
    old_params = set(re.findall(r'\[(?:string|switch|int)\]\s*\$([\w]+)', r, re.I))
    if new_params - old_params:
        return f"Parameter '-{list(new_params - old_params)[0]}' toegevoegd aan {basename}"

    # Graph endpoints
    new_uris = set(re.findall(r'graph\.microsoft\.com/v[\d.]+/([\w/]+)', a))
    old_uris = set(re.findall(r'graph\.microsoft\.com/v[\d.]+/([\w/]+)', r))
    if new_uris - old_uris:
        endpoint = list(new_uris - old_uris)[0].split('/')[0]
        return f"Graph API endpoint '{endpoint}' toegevoegd aan {basename}"

    # Export
    if 'Export-Csv' in a and 'Export-Csv' not in r:
        return f"CSV export toegevoegd aan {basename}"

    if basename.lower() == 'readme.md':
        return "Readme bijgewerkt"

    # Fallback op aantal regels
    n_add = len([l for l in added   if l.strip()])
    n_rem = len([l for l in removed if l.strip()])
    if n_rem > 0 and n_add == 0:
        return f"Code verwijderd uit {basename}"
    if n_rem > n_add * 2:
        return f"Code opgeschoond in {basename}"
    if n_add > 15:
        return f"Nieuwe functionaliteit toegevoegd aan {basename}"
    if n_add > 0:
        return f"Kleine wijziging in {basename}"
    return f"Update {basename}"


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

    run(f'git commit -m "{message}"')


if __name__ == '__main__':
    main()
