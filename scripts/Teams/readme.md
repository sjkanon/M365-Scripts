# Teams

Microsoft Teams / SharePoint export and archiving tooling.

---

## Scripts

### vias_archiver.ps1

Teams archiver with Graph, Teams and SharePoint export flow. PowerShell 7+ required, run as Global Admin.

**Parameters**

| Parameter | Description |
|-----------|-------------|
| `-Step10Action` | `interactive` (default), `archive`, `undo`, or `skip` |
| `-Step10Only` | Run only Step 10 (archive/unarchive), non-interactively |
| `-ChannelAction` | `none` (default), `archive`, or `undo` — per-channel quick mode |
| `-ChannelArchiveTag` | Marker text used for the rename fallback (default: `[ARCHIEF]`) |
| `-ChannelFallbackToRename` | Fall back to a rename marker if the Graph archive/unarchive API call fails |
| `-DryRun` | Simulate — keeps full auth/bootstrap and validates Steps 6-9 by probing counts, without writing exports or mutating archive state |

Current behavior (v8.19):
- Creates a unique temporary Entra app registration for the run.
- Grants only required delegated setup permissions during bootstrap.
- Applies Graph/SharePoint delegated consent to that temporary app.
- Removes the temporary app and service principal at cleanup (and on key setup failures).
- Registers an exit cleanup hook so the temporary app is also removed on PowerShell exit/Ctrl+C.
- Checks Teams/SharePoint folder access first during file export, and conditionally grants higher Graph rights when access is denied.
- Resolves channel file locations via Graph filesFolder for all channel types (standard/private/shared), with channel caching and fallback lookup.
- Normalizes TeamName/ChannelName values from Excel (trim) to avoid lookup misses caused by trailing spaces.
- Reuses the same cached channel resolver with Graph fallback in chat export, improving consistency for channel detection in dry-run and normal runs.
- Applies normalized channel-name matching (trim + whitespace collapse + lowercase) in cached and Graph fallback lookup to reduce false "Kanaal niet gevonden" cases.
- Handles SharePoint NotFound during export as a controlled skip instead of noisy hard failures.
- Downloads files with per-file retries, reconnect fallback, and post-download count validation to ensure completeness.
- Stores output in channel-based structure: `Teams > Team > Channel > Files, Chat, Members`.
- Does not archive Teams by default; archiving now requires explicit confirmation during Step 10.
- Step 10 also supports undo archiving (`unarchive`) with retry logic.
- Step 10 supports non-interactive quick mode: `-Step10Only -Step10Action undo|archive|skip`.
- Important: archive/unarchive is a team-level action in Microsoft Teams, not channel-level.
- Step 10 supports real per-channel archive/unarchive via Microsoft Graph (`/channels/{id}/archive|unarchive`).
- Quick per-channel mode: `-Step10Only -ChannelAction archive|undo`.
- Optional fallback to rename marker on API failure: `-ChannelFallbackToRename` (marker via `-ChannelArchiveTag`).
- Dry-run mode keeps full authentication/bootstrap and validates Steps 6-9 by probing Teams/SharePoint/Graph existence and counts, without writing member/chat/file exports to disk.
- Step 11 report in dry-run uses probe counts (detected files/messages) instead of local exported files.
- Step 10 archive/unarchive mutations remain simulated with `[DRYRUN]` output.
