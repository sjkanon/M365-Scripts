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
| [`Install-SharePointStructure.ps1`](Install-SharePointStructure.ps1) ([docs](#install-sharepointstructureps1)) | **Start here.** Asks what everything should be called, then builds the lot: app registration, team, channels, metadata, libraries, permissions, verification | yes |
| [`New-StructureConfig.ps1`](New-StructureConfig.ps1) ([docs](#you-are-asked-not-handed-a-json-file)) | The questions. Runs by itself from the installer; run it alone to prepare a configuration up front | writes the config |
| [`New-SharePointTeam.ps1`](New-SharePointTeam.ps1) ([docs](#install-sharepointstructureps1)) | The Microsoft 365 team and its channels, private ones included, and the site URLs written back | yes |
| [`New-SharePointMetadata.ps1`](New-SharePointMetadata.ps1) ([docs](#new-sharepointmetadataps1)) | Term set, site columns, content types — on every site in the config | yes |
| [`Set-SharePointLibraries.ps1`](Set-SharePointLibraries.ps1) ([docs](#set-sharepointlibrariesps1)) | Libraries, channel folders, content type binding, default metadata, views, group permissions | yes |
| [`Update-SharePointShareStatus.ps1`](Update-SharePointShareStatus.ps1) ([docs](#update-sharepointsharestatusps1)) | Derives Deelstatus from the real permissions, flags files shared wider than their tag allows | one column |
| [`Test-SharePointStructure.ps1`](Test-SharePointStructure.ps1) ([docs](#test-sharepointstructureps1)) | Compares the tenant with the config and reports every difference | never |
| [`Sync-SharePointChannelMember.ps1`](Sync-SharePointChannelMember.ps1) ([docs](#private-channels-and-groups)) | Makes a security group the source of truth for who is in a private channel | channel roster |
| [`Add-SharePointHelpPage.ps1`](Add-SharePointHelpPage.ps1) ([docs](#handing-it-over-to-the-customer)) | Writes the end-user explanation onto the team site, generated from the config | yes |
| [`Remove-SharePointStructure.ps1`](Remove-SharePointStructure.ps1) ([docs](#undoing-it)) | Removes what was built — reports only unless you pass `-Apply` | yes, on purpose |
| [`SharePointStructure.Common.ps1`](SharePointStructure.Common.ps1) | Shared helpers — dot-sourced, not run on its own | — |
| [`Petsolutions-SharePoint-Handleiding.md`](Petsolutions-SharePoint-Handleiding.md) | **End-user guide, in Dutch** — hand this to the customer: uploading, tagging, finding things back | — |
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

### The short version

```powershell
.\Install-SharePointStructure.ps1
```

That is the whole thing. With no configuration for this tenant it asks what everything
should be called, writes the configuration itself, then creates the team, the channels
(the private one included), the metadata model, the libraries, the groups, the views and
the permissions — and verifies the result.

Afterwards: put people in the security groups. The explanation for them is already
on the site, in the left-hand navigation — the build put it there.

### You are asked, not handed a JSON file

`New-StructureConfig.ps1` runs by itself the first time. Enter accepts the suggestion in
brackets, so a standard build is mostly Enters and two real answers:

| Asked | Suggestion |
|---|---|
| Client, tenant | — |
| Team name, alias (decides the site URL), owner | derived from the client name |
| Brands, and what "belongs to all of them" is called | Butterstone, Laseto, Beide |
| Pillars, and which are a private channel | MGMT, Leveranciers, Verkopers, Klanten, Marketing, TD — MGMT private |
| Which pillar handles suppliers / sales | decides where Leverancier and Regio become required |
| Customer library | FUTECH Images and videos |
| Group prefix and the edit/read suffixes | `SG-<CLIENT>` · RW · RO |
| Languages, regions, document kinds, confidentiality levels, statuses, starting suppliers | the Dutch defaults |
| **Maintain the share-status column?** | no — this is the only answer that costs you a nightly script |
| **Enforce per-pillar rights on channel folders?** | no — see the note on standard channels below |

Everything else is derived: per pillar a channel, a content type, two security groups and
a grouped view; per brand a view spanning every pillar.

Optional questions say `(of "geen")` in the hint — type that to turn the suggestion down,
because Enter means "take it".

#### `-All` — decide every name yourself

`New-StructureConfig.ps1 -All` also asks for the names that are otherwise derived, each
still with the derivation as its suggestion:

| Asked with `-All` | Suggestion |
|---|---|
| Team site URL | `https://<tenant>.sharepoint.com/sites/<alias>` |
| The library behind the channels | `Documents` — ask this on a non-English tenant |
| Column group and content type group | the team name |
| Term set name | `Leveranciers` |
| The label of every column, as users see it | Merk, Pijler, Regio, Leverancier, Taal, Contenttype, … |
| Per pillar: channel name, folder, content type, both group names, view title | derived from the pillar name |
| Customer library: group and content type name | `<prefix>-Klanten-Extern`, `Klantmedia` |

The column *internal* names (`PsMerk`, `PsTaal`, …) stay fixed either way. They are never
shown to anyone, and changing one after documents carry it loses the metadata on those
documents.

**Two things are generated once and then fixed**, because SharePoint keys data to them:
the column internal names (`PsMerk`, `PsTaal`, …) and the content type IDs. Display names,
channel names and group names can all be changed afterwards; those two cannot without
losing the metadata on documents that already carry them. That is why the wizard refuses
to overwrite an existing configuration without `-Force`.

### Install-SharePointStructure.ps1

One run, five steps, stopping at the first failure rather than building on a broken one:

| Step | What |
|---|---|
| 0 | App registration — created and admin-consented, or reused from `pnp.appid.json` |
| 1 | `New-SharePointTeam.ps1` — the Microsoft 365 team, the channels including the private one, and the site URLs written back into the configuration |
| 2 | `New-SharePointMetadata.ps1` — term set, columns, content types, on every site |
| 3 | `Set-SharePointLibraries.ps1 -EnsureGroups` — groups, libraries, folders, content types, defaults, views, permissions |
| 4 | `Add-SharePointHelpPage.ps1` — the explanation, on the site, for the people who will use it |
| 5 | `Test-SharePointStructure.ps1` — read-only verification of what just landed |
| 6 | `Update-SharePointShareStatus.ps1` with `-RunAudit` — the first deelstatus pass |

**Step 4 is part of building it, not an errand for later.** A structure nobody was told
about is a structure nobody uses, and the page is generated from the same configuration,
so it describes what the run just made. `-SkipHelpPage` leaves it out; `-HelpContact`
says who people should ask.

**Step 1 is why this works from an empty tenant.** A private channel's site collection is
provisioned asynchronously and its URL cannot be known in advance — SharePoint invents it
from the team and channel name. The script polls for it (a couple of minutes is normal)
and writes it into the configuration, so the steps below have somewhere to connect to.
Pass `-SkipTeam` when the team already exists.

```powershell
# One-off build on a tenant you do not manage day to day: leave nothing behind
.\Install-SharePointStructure.ps1 -TemporaryApp -RunAudit

# Conservative: everything except the unsupported channel-folder permissions
.\Install-SharePointStructure.ps1 -SkipChannelFolderPermissions

# Use an app registration you already have
.\Install-SharePointStructure.ps1 -ClientId <app-id>
```

| Exit code | Meaning |
|---|---|
| 0 | built and verified |
| 1 | a step failed |
| 2 | built, but the verification found differences |

**About `-TemporaryApp`.** It deletes the app registration at the end — but only one
*this run created*. An app that was already cached predates the run and is somebody
else's to remove, so the script says so instead of quietly deleting it. Without the
switch the app stays and the client ID is cached, which is what the scheduled audit and
later drift checks need.

**A `-WhatIf` run needs an app to sign in with.** With no cached app for the tenant there
is nothing to connect as, so the dry run validates the configuration and stops there. Run
it once for real, or pass `-ClientId` of an existing app, to dry-run step by step.

### Or step by step

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

#### Cross-cutting views — what makes "brand as a tag" real

Per-pillar views only ever show one channel folder. The `libraryViews` section adds views
on the shared library itself with `Scope = RecursiveAll`, so they span **every** pillar
folder in one flat list:

| View | Shows |
|---|---|
| `Alles - Butterstone` | every file tagged Butterstone **or Beide**, across all pillars, grouped by pillar |
| `Alles - Laseto` | the same for Laseto |
| `Nog te taggen` | files with no Merk — what drag-and-drop and OneDrive sync leave behind |
| `Extern gedeeld` | everything the audit found sitting outside the organisation |
| `Te archiveren` | Status is Te archiveren or Verouderd |

This is the answer to "one file, two brands": a file tagged `Beide` is stored once and
appears in both brand views. No copies to drift apart.

The filter is raw CAML in the config rather than a mini query language of this script's
own invention:

```jsonc
{
  "title": "Alles - Butterstone",
  "recursive": true,
  "groupBy": "PsPijler",
  "where": "<Or><Eq><FieldRef Name='PsMerk' /><Value Type='Text'>Butterstone</Value></Eq><Eq><FieldRef Name='PsMerk' /><Value Type='Text'>Beide</Value></Eq></Or>",
  "fields": [ "DocIcon", "LinkFilename", "PsPijler", "PsContenttype", "PsTaal", "..." ]
}
```

Never group a view on `PsTaal` — SharePoint refuses to group on a multi-value column.
Filtering on it works fine.

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

## A channel where only one group gets in

This is the shape to reach for when a pillar has to be closed off **and** keep a real
read-only role. The wizard asks for it as `bibliotheek`:

```
Welke pijlers moeten afgeschermd worden [MGMT]:
In welke vorm [bibliotheek]:
```

What it builds:

| Piece | Where |
|---|---|
| A normal channel in the team | everyone sees it in Teams, as a channel should be |
| Its own document library | on the team site, not a folder in the shared one |
| **Unique permissions, inheritance broken without copying** | so team members do **not** come across as editors |
| `-RW` → Contribute, `-RO` → **Read** | a real read-only role, which a private channel cannot offer |
| A tab in the channel pointing at the library | the channel's own Files tab cannot be repointed, so the library sits beside it |

The permissions that survive are the site's own owners plus those two groups. That is
what `keepExistingPermissions: false` on the container means, and it is the difference
between "closed off" and "closed off in theory".

> **The channel's built-in Files tab still points at the team library.** It will hold a
> folder that nobody uses. Tell people to use the named tab, or remove the Files tab
> from the channel by hand once.

Compared with the alternatives:

| Shape | Read-only role | Channel visible to non-members | Supported |
|---|---|---|---|
| **Own library + channel** | yes | yes (files are not) | yes |
| Private channel | no — members may edit, full stop | no | yes |
| Standard-channel folder with unique rights | yes | yes, but the Files tab errors | no |

---

## Private channels and groups

**A private channel cannot be given rights through a group.** Teams tracks its
membership one person at a time, and Graph only accepts individual users there. There
is no way around that, and the obvious workaround is a trap:

| Approach | Verdict |
|---|---|
| Add the group as a channel member | Not possible — Graph takes users only |
| Add the group to the channel site's SharePoint permissions | Works for about a day. Teams syncs the channel roster back over it, and in the meantime those people reach the files while the channel stays invisible to them in Teams. Unsupported |
| **Let the group feed the roster** | What `Sync-SharePointChannelMember.ps1` does |

```powershell
.\Sync-SharePointChannelMember.ps1 -WhatIf     # who would be added
.\Sync-SharePointChannelMember.ps1             # add them
.\Sync-SharePointChannelMember.ps1 -Prune      # and remove who the groups no longer list
```

You manage the group; the script puts its people in the channel. Nested groups are
followed, non-users are dropped, and everyone is made a member of the parent team first
— Teams refuses a private-channel member who is not on the team, and the error it gives
does not say so.

Which groups feed which channel comes from the container's `channelMembers`. A
configuration written before that key existed falls back to the configured groups named
after the container, and says that it did.

> ### There is no read-only role in a private channel
>
> A private channel has owners and members, and members may post, edit and delete
> files. A group named `-RO` therefore cannot mean "may look" there — everyone this
> script adds can write. The run reports per group how many people it brought in, so
> that is visible rather than assumed.
>
> If read-only genuinely matters for a pillar, a private channel is the wrong shape for
> it. Use a document library with its own permissions, where Read is a real role — the
> customer library already works that way.

**The alternative worth knowing about:** a *shared* channel does support group-based
membership. Moving a pillar there is the supported way to have groups decide access to
a channel. Behaviour varies with the tenant's external-sharing and B2B direct connect
settings, so try one channel before moving anything that matters.

---

## Handing it over to the customer

The explanation belongs on the site, not in this repo. `Add-SharePointHelpPage.ps1`
writes it there as a SharePoint page, generated from the same configuration the
structure was built from:

```powershell
.\Add-SharePointHelpPage.ps1 -WhatIf                      # what would it say
.\Add-SharePointHelpPage.ps1 -Interactive -ClientId <app-id>
.\Add-SharePointHelpPage.ps1 -Interactive -ClientId <app-id> -Force   # after a change
```

Because it is generated, it cannot drift: the channels it lists are the channels that
exist, the labels it explains carry the same help text users see under each field, and
the required fields per document type are read off the content types. Rename a channel,
rerun it, and the page says the new thing.

It is written for the person uploading a catalogue. No group names, no internal column
names, no content types or site columns. Two things from the configuration are
deliberately kept off it: the `note` on a container, which names security groups, and
the `description` on a view, which talks about pillars and synced folders — their titles
are plain enough on their own. Permissions are not on it at all: who may see what is not
something a user can act on, and explaining it only invites the question of why they
cannot.


[`Petsolutions-SharePoint-Handleiding.md`](Petsolutions-SharePoint-Handleiding.md) is
written for the people who will actually upload files — in Dutch, no jargon, five minutes
to read. It covers the three ways of adding a file and why they behave differently, what
each label means, and what happens the moment you tag something.

Two things in there are worth knowing about as the administrator, because they are the
questions that come back:

- **Drag-and-drop and OneDrive sync ask nothing.** Required columns are enforced by the
  upload form, not by the library. Files dropped in bulk land with empty labels and a
  "Required info" prompt — they are not blocked. The `Nog te taggen` view is the cleanup
  list, and the guide tells users to work it with a multi-select and the details pane.
- **A label is not a lock.** Setting Vertrouwelijkheid to Vertrouwelijk shuts nobody out;
  it is an agreement, plus the signal the nightly audit uses to flag over-sharing. Access
  comes from the security groups. The guide says this in a call-out box, because users
  will otherwise assume the opposite.

---

## Undoing it

[`Remove-SharePointStructure.ps1`](Remove-SharePointStructure.ps1) takes the same
configuration apart, deepest first. **It has the reverse default of everything else
here: without `-Apply` it changes nothing.** Forgetting `-WhatIf` on a destructive
script is the dangerous direction, so the safe state is the one you get for free.

```powershell
.\Remove-SharePointStructure.ps1                                   # what would go
.\Remove-SharePointStructure.ps1 -Scope Channels,Groups -Apply     # part of it
.\Remove-SharePointStructure.ps1 -Scope All,Team -IncludeContent -Apply   # start over
```

| `-Scope` | Removes |
|---|---|
| `Tabs` | the library tabs added to channels |
| `Channels` | the configured channels, and the files in their folders |
| `Libraries` | the libraries a container owns (`kind: Library`) |
| `ContentTypes` | unbound from the lists first, then removed |
| `Columns` | the site columns, on every site in the config |
| `TermSet` | the term set, its group and its terms |
| `Groups` | the Entra ID security groups |
| `Team` | the Microsoft 365 group — the site, every library, every file, every chat |
| `All` | everything above **except** `Team` |

`All` never includes the team. Deleting a client's whole team is not something you
should get by asking for "all" — you have to name it, and then type the team's name to
confirm.

**What it refuses to do:**

- A library or channel folder that still holds files is skipped unless `-IncludeContent`.
  The item count is reported either way.
- The General channel and the team's own Documents library are never removed.
- A content type still in use is reported, not forced.

**What no recycle bin brings back:** removing the term set orphans the Leverancier value
on every document that carried one — the field keeps a GUID that resolves to nothing.
Removing a column takes its data with it. Both are reported with that cost before they
run. A deleted group or team is soft-deleted for 30 days; a deleted channel has its own
30-day recycle; files from a removed library go to the site recycle bin.

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
