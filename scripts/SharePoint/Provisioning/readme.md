# SharePoint Structure Provisioning

Provision and maintain a SharePoint structure — metadata model, libraries, content
types and group permissions — from one configuration file, with PnP PowerShell and
Microsoft Graph.

Built for Petsolutions NV (brands Butterstone and Laseto), but nothing in the scripts
is client-specific: the model lives in the JSON, so a second client is a second config
file, not a second fork.

---

## Contents

- [The idea](#the-idea)
- [Scripts](#scripts)
- [Before you start](#before-you-start)
- [Order of operations](#order-of-operations)
- [The configuration file](#the-configuration-file)
- [Standard channels and permissions — read this](#standard-channels-and-permissions--read-this)
- [Scheduling the audit](#scheduling-the-audit)
- [What the scripts will never do](#what-the-scripts-will-never-do)

---

## The idea

One Microsoft 365 Team, one channel per pillar, and the difference between the pillars
carried by **metadata and group permissions** rather than by a sprawl of sites.

| | |
|---|---|
| **Pillars** | MGMT (private channel), Leveranciers, Verkopers, Klanten, Marketing, TD |
| **Extra library** | FUTECH Images and videos — read-only for external customers |
| **Brand** | Butterstone / Laseto / Beide — a tag on every file, never a separate site or group |
| **Groups** | one Entra ID security group per pillar per access level (`SG-PETSOL-<Pijler>-RW` / `-RO`), plus `SG-PETSOL-FUTECH-Klanten` |

The metadata model, reusable across every library:

| Column | Internal name | Type | Values |
|---|---|---|---|
| Merk | `PsMerk` | Choice | Butterstone / Laseto / Beide |
| Pijler | `PsPijler` | Choice | MGMT / Leveranciers / Verkopers / Klanten / Marketing / TD |
| Regio | `PsRegio` | Choice | Benelux / Duitsland / Frankrijk / Export — only required on Verkoopdocument |
| Leverancier | `PsLeverancier` | Managed metadata | term set, extendable from the term store |
| Taal | `PsTaal` | MultiChoice | NL / FR / DE / EN / Geen taal |
| Contenttype | `PsContenttype` | Choice | Catalogus / Prijslijst / Schrijfrichtlijn / Afbeelding+certificaat / Marketingslag |
| Vertrouwelijkheid | `PsVertrouwelijkheid` | Choice | Intern / Deelbaar met klant / Vertrouwelijk |
| Deelstatus | `PsDeelstatus` | Choice | maintained by the audit script — never filled in by hand |
| Status | `PsStatus` | Choice | Actief / Te archiveren / Verouderd |

> **Why the `Ps` prefix.** "Contenttype" and "Status" are display names SharePoint
> already uses for something else. Prefixing the *internal* names keeps the columns
> unambiguous in CAML, in views and in the drift check, while users still see plain
> Dutch labels.

A content type per pillar decides which of those are mandatory — a Marketingdocument
cannot be saved without Taal and Contenttype, a Verkoopdocument cannot be saved
without Regio.

---

## Scripts

| Script | What it does | Writes? |
|---|---|---|
| [`New-SharePointMetadata.ps1`](#new-sharepointmetadataps1) | Term set, site columns, content types — on every site in the config | yes |
| [`Set-SharePointLibraries.ps1`](#set-sharepointlibrariesps1) | Libraries, channel folders, content type binding, default metadata, views, group permissions | yes |
| [`Update-SharePointShareStatus.ps1`](#update-sharepointsharestatusps1) | Derives Deelstatus from the real permissions, flags files shared wider than their tag allows | one column |
| [`Test-SharePointStructure.ps1`](#test-sharepointstructureps1) | Compares the tenant with the config and reports every difference | never |
| `SharePointStructure.Common.ps1` | Shared helpers — dot-sourced, not run on its own | — |
| `petsolutions.config.json` | The model | — |

Every writing script supports `-WhatIf` and is idempotent: a second run reports `[ OK ]`
across the board and changes nothing.

---

## Before you start

### 1. Modules

```powershell
Install-Module PnP.PowerShell         -Scope CurrentUser
Install-Module Microsoft.Graph.Groups -Scope CurrentUser   # only for -EnsureGroups / -IncludeGroups
```

### 2. An app registration

PnP.PowerShell no longer ships a shared multi-tenant app, so you need one of your own.
Reuse the app that [`Find-SiteContent.ps1`](../Find-SiteContent.ps1) creates and caches
in `pnp.appid.json`, or register one:

| Sign-in | Needs | Use it for |
|---|---|---|
| **Interactive** (`-Interactive -ClientId <app-id>`) | delegated `AllSites.FullControl`, signed in as an admin who may edit the site | provisioning, one-off runs |
| **App-only** (`-ClientId <app-id> -Thumbprint <thumb>`) | application `Sites.FullControl.All`, certificate uploaded to the app | the scheduled share-status audit |

Two extra requirements that are easy to miss:

- **Term store.** Creating the term group and term set needs a term store administrator.
  App-only can only do it when the app's service principal is added as one in the
  SharePoint admin centre. If that is a hurdle, run
  `New-SharePointMetadata.ps1 -Only TermSet` interactively once and leave the rest to
  app-only.
- **Groups.** `-EnsureGroups` needs `Group.ReadWrite.All`; the drift check's
  `-IncludeGroups` gets by with `Group.Read.All`.

### 3. Fill in the config

`petsolutions.config.json` ships with `CHANGEME` in the tenant and site URLs. Every
script refuses to run until those are replaced — better a clear error than a sign-in
that fails five minutes into a run.

```jsonc
"tenant": "petsolutions.onmicrosoft.com",
"sites": {
  "team": "https://petsolutions.sharepoint.com/sites/Petsolutions",
  "mgmt": "https://petsolutions.sharepoint.com/sites/Petsolutions-MGMT"
}
```

> The **mgmt** URL is the private channel's *own* site collection, not a folder in the
> team site. Find it in the SharePoint admin centre, or open the MGMT channel's Files
> tab and click *Open in SharePoint*. A private channel has its own site, which is why
> the columns and content types have to be provisioned there separately — a site column
> does not reach across a site collection.

---

## Order of operations

```powershell
# 1. Look before you leap - both of these change nothing
.\Test-SharePointStructure.ps1 -Interactive -ClientId <app-id>
.\New-SharePointMetadata.ps1   -Interactive -ClientId <app-id> -WhatIf

# 2. Metadata model first: the libraries cannot bind what does not exist
.\New-SharePointMetadata.ps1 -Interactive -ClientId <app-id>

# 3. Libraries, content type binding and permissions (creates the groups on the way)
.\Set-SharePointLibraries.ps1 -Interactive -ClientId <app-id> -EnsureGroups -WhatIf
.\Set-SharePointLibraries.ps1 -Interactive -ClientId <app-id> -EnsureGroups

# 4. Fill the groups with people (Entra ID portal, or your own onboarding script)

# 5. First audit, and from then on nightly
.\Update-SharePointShareStatus.ps1 -Interactive -ClientId <app-id> -ReportOnly

# 6. Confirm the result
.\Test-SharePointStructure.ps1 -Interactive -ClientId <app-id> -IncludeGroups
```

### New-SharePointMetadata.ps1

Term group, term set and terms; the nine site columns; the seven content types with
the right columns marked required. Runs against **every** site in the config, so the
private channel site gets its own copy.

```powershell
# Only the private channel site, after MGMT moved to its own channel
.\New-SharePointMetadata.ps1 -Interactive -ClientId <app-id> -Site mgmt

# Only the term set, by a term store admin
.\New-SharePointMetadata.ps1 -Interactive -ClientId <app-id> -Only TermSet
```

Making a column required after the fact works: the `Required` flag on an existing field
link is updated in place and pushed down to the lists already using the content type.

### Set-SharePointLibraries.ps1

Per container: the library or channel folder, the content types bound to the library,
the folder's own content type order (so the *New* menu in the Leveranciers channel
offers Leveranciersdocument and not the five types belonging to other pillars), the
default column values, a grouped view, and the role assignments.

```powershell
# The external library, and strip anything the config does not list
.\Set-SharePointLibraries.ps1 -Interactive -ClientId <app-id> `
    -Container FUTECH -RemoveOtherPermissions

# Everything except the unsupported standard-channel folder permissions
.\Set-SharePointLibraries.ps1 -Interactive -ClientId <app-id> `
    -SkipChannelFolderPermissions
```

| Switch | Effect |
|---|---|
| `-EnsureGroups` | create the Entra ID security groups that do not exist yet |
| `-RemoveOtherPermissions` | remove role assignments the config does not list (owners, sharing links and the everyone-claim are never removed) |
| `-RemoveStockContentType` | drop the built-in *Document* type so nobody can file without metadata |
| `-SkipChannelFolderPermissions` | leave channel folders inheriting — see below |

No view is ever made the default. The default view of a Teams library is what every
member of the channel sees the second they open Files.

### Update-SharePointShareStatus.ps1

Works out how each document is really shared and writes it to Deelstatus:

| Verdict | When |
|---|---|
| `Extern - bewerken` | an Anyone-edit link, a guest with Contribute or more, or a specific-people edit link containing a guest |
| `Extern - alleen bekijken` | the same, view-only |
| `Intern gedeeld` | a link or a direct grant that stays inside the tenant |
| `Niet gedeeld` | the file just inherits from its folder or library |

It also answers the question the model is really for: **is anything tagged Intern or
Vertrouwelijk sitting behind an external link?** Those go in the report and the run
exits with code 2.

The column is written with `SystemUpdate`, so Modified and Modified By stay put and no
new version is created — a nightly run does not push the whole library to the top of
*recently changed*.

| Exit code | Meaning |
|---|---|
| 0 | done, nothing shared wider than its tag allows |
| 1 | failure |
| 2 | at least one Vertrouwelijkheid violation |

### Test-SharePointStructure.ps1

Read-only. One row per difference, in three flavours:

| Kind | Meaning | Fixed by |
|---|---|---|
| `Missing` | in the config, not on the tenant | the provisioning scripts |
| `Different` | present, but not as configured | the provisioning scripts |
| `Extra` | on the tenant, not in the config | nobody — that is your call |

Exit code 2 means drift, so it slots straight into a monitor. The check that earns its
keep most often is the **required flag on a content type field** — somebody unticks it
in the browser and nothing looks wrong until half a library has no Taal on it.

---

## The configuration file

| Section | Holds |
|---|---|
| `tenant`, `sites` | where everything lives; `sites` keys are referenced by each container |
| `termStore` | term group, term set and the starting terms for Leverancier |
| `columns` | the site columns: internal name, display name, type, choices, default |
| `contentTypes` | one per pillar, with `fields[].required` deciding what is mandatory |
| `groups` | the Entra ID security groups, by display name |
| `containers` | the libraries and channel folders, and who gets which role on them |

A container:

```jsonc
{
  "key": "Leveranciers",
  "kind": "ChannelFolder",          // ChannelFolder = a Teams channel; Library = its own library
  "site": "team",                   // key from the sites section
  "list": "Documents",              // the channel's library
  "folder": "Leveranciers",
  "contentTypes": [ "Leveranciersdocument" ],
  "defaultContentType": "Leveranciersdocument",
  "defaultColumnValues": { "PsPijler": "Leveranciers", "PsStatus": "Actief" },
  "uniquePermissions": true,
  "keepExistingPermissions": true,  // copy the inherited rights when breaking inheritance
  "view": { "title": "Op leverancier", "fields": [ ... ], "groupBy": "PsLeverancier" },
  "permissions": [
    { "group": "SG-PETSOL-Leveranciers-RW", "role": "Contribute" },
    { "group": "SG-PETSOL-Leveranciers-RO", "role": "Read" }
  ]
}
```

The config is cross-checked before anything connects: a content type referring to a
column that is not defined, or a container granting a group that is not in the model,
fails at load time rather than halfway through provisioning.

**Adding a leverancier** does not need a config change — add the term in the term store
and it appears in the column. The `terms` list is only the starting set; terms added
outside the config are reported by the drift check but never removed.

---

## Standard channels and permissions — read this

The shipped configuration gives each standard-channel folder unique permissions.
**Microsoft does not support that combination.**

A standard channel is visible to every member of the team by design. Tightening the
SharePoint permissions on the channel's folder does hide the files, but Teams keeps
showing the channel: a member who lost access gets an error on the Files tab rather
than a closed door. It works, it is not pretty, and it is not blessed.

The supported ways to close a pillar off:

| Option | Trade-off |
|---|---|
| **Private channel** | own site collection, own membership — what MGMT already uses. Cleanest, but the channel does not appear for non-members at all |
| **Shared channel** | own site collection, own membership, can include people outside the team |
| **Own library** (`kind: Library`) | outside the channel structure, unique permissions are entirely supported — what FUTECH uses |

If you move a pillar to a private or shared channel, set `uniquePermissions` to `false`
for its container and add its new site to the `sites` section.

`Set-SharePointLibraries.ps1` warns on every standard-channel folder it breaks
inheritance on, and `-SkipChannelFolderPermissions` leaves them inheriting while still
doing the content types, defaults and views.

---

## Scheduling the audit

App-only, on a server or an RMM, nightly:

```powershell
pwsh -NoProfile -File .\Update-SharePointShareStatus.ps1 `
    -ClientId <app-id> -Thumbprint <thumbprint> -Quiet
```

With `-Quiet` the run prints only the summary, so it turns up in the activity feed with
something to say. Exit code 2 means a file tagged Intern or Vertrouwelijk is behind an
external link — worth an alert.

A weekly drift check alongside it:

```powershell
pwsh -NoProfile -File .\Test-SharePointStructure.ps1 `
    -ClientId <app-id> -Thumbprint <thumbprint> -Quiet -IncludeGroups
```

---

## What the scripts will never do

Deliberate omissions, each for a reason:

| Never | Why |
|---|---|
| Delete a site column, content type or field link | it would take the metadata on existing documents with it |
| Remove a term from the term store | tagged documents reference term GUIDs; the term set is meant to be extended from the UI |
| Remove a role assignment the config does not list | unless you pass `-RemoveOtherPermissions` — an exception someone made on purpose is not drift |
| Remove site owners or sharing-link groups | `-RemoveOtherPermissions` skips those; stripping them locks the client out of their own library |
| Revoke a sharing link | the audit reports over-sharing. Revoking is a decision, and a decision belongs with a person |
| Make a view the default | every member of the channel would notice at once |

---

## Notes

- Author: Sjoerd Kanon
- `SharePointStructure.Common.ps1` is dot-sourced by all four scripts. That is a
  deliberate exception to the "every script stands alone" rule elsewhere in this repo:
  these four share one config schema, and three copies of the permission code would
  drift apart within a month.
- Test against a non-production tenant first. The provisioning scripts change
  permissions on a live team.
