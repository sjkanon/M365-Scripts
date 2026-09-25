# Working rules for this repository

These rules apply to every change in this repo, without being asked each time.

---

## 1. Commit every change

- Commit each logical change **as it is made**, not in one lump at the end of a session.
- One commit per coherent change: a fix, a new script, a readme update that belongs with it.
- Write a descriptive, human message in the imperative or narrative style already used in
  `git log` — say *what changed and why*, never `update`, `wip`, `fix` or `changes`.
- Never use `--no-verify` or skip signing.
- Never amend or rewrite a commit that has already been pushed.

## 2. Push after committing

- Push to `origin` after each commit (or after a short burst of related commits) —
  do not leave the branch sitting ahead for a whole session.
- Work on `devel`. Never commit straight to `main`; `main` is updated through a PR from `devel`.
- If the push is rejected, fetch and rebase onto `origin/devel`, then push again.
  Never force-push a shared branch unless explicitly asked.

## 3. Update the readmes with the change

A code change is not finished until the documentation matches it. For every functional or
structural change, update **all** of the following that apply:

1. **The folder readme** (`scripts/<Workload>/readme.md`)
   - Add or update the row in the `## Scripts` table (script link + one-line description).
   - Add or update the `### <Script>.ps1` section below it: purpose, the `**Parameters**`
     table, `**Examples**` block, and `**Notes**`.
   - If a folder was added, add it to the `## Folders` table of the parent readme, and give
     the new folder its own `readme.md`.
2. **Parent folder readmes** up the chain, if the entry there no longer describes reality.
3. **The root `readme.md`**
   - Update the relevant `Script Categories` entry and the `Repository Structure` tree.
   - Add an entry to `## Version History` — **required for every functional or structural
     change**. Use the existing shape: a `### YYYY-MM-DD` heading (append ` (n)` for the
     n-th change that day) with a `| Change |` table, one row per point. Describe what
     changed, why it was wrong before, and what was actually verified.
4. **`menu.ps1`** — add, rename, or remove the entry when a script is added, renamed, moved
   or deleted, and check that no other script or readme still points at the old path.

Rename or move a file with `git mv`, and grep the repo for the old path afterwards
(readmes, `menu.ps1`, `f.ps1`, other scripts) before committing.

## 4. Before committing PowerShell

- Run the syntax check on what you touched:
  `pwsh -NoProfile -File scripts/Startup/Test-PowerShellSyntax.ps1 -Path <path> -Recurse`
- Keep the existing header comment block (Synopsis, Description, Parameters, Example) in sync
  with the actual parameters.
- Report honestly what was verified on a live tenant/device and what was not — the version
  history distinguishes the two, and it should stay that way.
