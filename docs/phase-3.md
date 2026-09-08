# Phase 3 — UI-assisted development and testing foundation

Status: **In progress 🚧 — current read/alignment increment is preparing the next focused PR**

Started: 2026-09-08

## Purpose

Phase 1 and Phase 2 proved the Classic Outlook COM boundary and genuine native meeting forwarding, but continued development became inefficient when manual testing required copying `GlobalAppointmentID`, `EntryID`, source SMTP addresses, and target addresses into PowerShell commands.

Phase 3 brings a usable WinUI application forward so identity, authority and safety behavior can be inspected through the application before write features are introduced.

Phase 2 is complete and merged in PR #5. Phase 3 is intentionally being delivered as smaller independently testable PR increments rather than waiting for the whole roadmap phase to finish.

The current Phase 3 application remains deliberately conservative: Calendar and Alignment can discover and reason about Outlook state, native Forward can be checked/prepared-and-discarded, and Move Selected can be previewed, but there is still no Move write command and the UI cannot send a meeting.

## Current increment scope

This increment is intended to stand on its own as a review/test unit. It contains:

- real WinUI shell and navigation;
- GUI-driven Outlook account/calendar discovery;
- safe native Forward capability/prepare testing;
- preliminary non-recurring correlation and alignment states;
- read-only managed-copy metadata detection;
- authority reasoning;
- preview-only Move Selected planning;
- packaged Windows development artifact.

It deliberately excludes:

- GUI Forward send;
- Copy Full/Busy writes;
- Move writes;
- recurring identity/write support;
- persistent authority/history;
- production named-pipe IPC;
- final FullCalendar/WebView2 integration.

## Current UI

The unpackaged WinUI 3 application now has four functional destinations:

- **Calendar** — reads the bounded Outlook calendar window, selects a meeting, derives its source account, offers the other discovered accounts as targets, and exposes the safe native-Forward test actions.
- **Alignment** — groups safely correlated non-recurring meetings, shows discrepancy states, resolves or lets the user select authority, identifies validated Outlook Aligner-managed copies, and builds a read-only Move Selected preview.
- **Settings** — owns the current scan horizon and shows discovered Classic Outlook accounts/calendar availability.
- **Diagnostics** — contains raw host output, Outlook IDs, managed-copy metadata and partial-failure details that should not appear in normal workflows.

Internal Outlook identifiers are therefore no longer normal user input.

## OutlookHost protocol v2

Phase 3 introduces OutlookHost protocol version **2** for managed-copy metadata.

The UI validates the protocol before consuming scan output. An older protocol-v1 OutlookHost cannot silently feed the new UI a calendar result that lacks the ownership information needed by Copy/Move safety logic.

## Managed-copy metadata — read only

The calendar reader looks for these existing custom appointment properties:

- `OutlookAligner.SyncGroupId`
- `OutlookAligner.SourceGlobalAppointmentId`
- `OutlookAligner.SourceAccountId`
- `OutlookAligner.CopyType`
- `OutlookAligner.SchemaVersion`

The Phase 3 reader only calls Outlook property lookup APIs. It does not add properties, set values, call `Save()`, or otherwise modify an appointment.

Metadata is classified as:

- `None`
- `Valid`
- `Incomplete`
- `UnsupportedSchema`
- `Unreadable`

Only `Valid` metadata is trusted as managed-copy ownership/provenance.

`Incomplete`, `UnsupportedSchema`, and `Unreadable` metadata now **fail closed in the read model as well as the future write model**:

- the item is not silently downgraded to an ordinary unmanaged appointment;
- it is excluded from normal native-GlobalAppointmentID correlation;
- it appears as an explicit `Conflict` row;
- authority selection/inference is blocked for that unresolved item;
- it cannot become a Move candidate.

This prevents suspicious Aligner metadata from making the Alignment view look more certain than the evidence supports.

Current managed-copy schema version is `1`. Supported copy-type markers are `Full` and `Busy`.

## Correlation semantics

Ordinary non-recurring Outlook items correlate by native `GlobalAppointmentID`.

A validated Outlook Aligner-managed local copy is different: a locally created copy may have its own native Outlook `GlobalAppointmentID`, so its logical correlation key is the validated stored `OutlookAligner.SourceGlobalAppointmentId`.

This means:

- a source appointment and a valid managed copy can form one logical row even when the copy's own native GlobalAppointmentID differs;
- an unmanaged appointment never gets to influence correlation merely because some similarly named custom field is present;
- a managed item without validated source identity remains unresolved rather than falling back to a guess;
- suspicious/incomplete/unsupported/unreadable managed metadata is surfaced as `Conflict`, not normal unmanaged correlation;
- recurring items remain deliberately unresolved until recurrence identity is proven.

The current discrepancy states are: `Aligned`, `Missing`, `Moved`, `DetailsDifferent`, `Duplicate`, `Conflict`, `RecurrenceIdentityUnresolved`, and `Uncorrelated`.

## Authority

Alignment supports two authority paths in this increment.

### Known origin

Validated Outlook Aligner copy metadata can establish `KnownOrigin` only when:

1. all validated managed copies in the logical group agree on their recorded source account and sync-group identity;
2. that recorded source account resolves to exactly one observed member in the logical group;
3. the observed source member is not itself a managed copy.

If those conditions hold, the UI selects that source account as authority and explains that it came from validated managed-copy provenance.

### User selected

When origin cannot be proven but the logical row is otherwise safe, the user can select authority explicitly for the running session.

Authority remains unavailable for recurrence-unresolved, uncorrelated, duplicate, and conflict rows.

The app never infers authority from `LastModificationTime` and does not pretend pre-existing copies reveal their origin.

## Move Selected preview

Phase 3 contains a pure Core `MovePreviewPlanner` plus an Alignment UI preview.

The planner can propose a Move only when all current safety prerequisites hold. It:

1. requires exactly one observed member in the selected authority account;
2. never proposes the authority item as a target;
3. blocks duplicate members;
4. blocks conflicting/incomplete managed provenance;
5. blocks recurring writes;
6. considers only validated Outlook Aligner-managed non-authority copies eligible for change;
7. skips pre-existing/unmanaged Outlook appointments;
8. skips copies already aligned with the authority;
9. records the current and proposed Start/End values for every candidate action.

The Alignment UI renders the resulting plan as plain-language `MOVE`, `SKIP`, or `BLOCKED` lines. It does not show locator IDs in this normal workflow.

**There is no Move write command in this increment.** The preview is intentionally useful before mutation exists.

## Forward testing from the UI

For a selected Calendar event the UI offers:

### Check Forward

Runs the proven Calendar native-Forward capability probe. No Forward is executed.

### Prepare Forward (discard)

This safe test action:

1. re-runs native Forward capability immediately before preparation;
2. uses the selected event's internal source SMTP / `GlobalAppointmentID` / `EntryID`;
3. uses the selected target Outlook account's SMTP address;
4. asks OutlookHost to create the native Calendar Forward;
5. resolves exactly that target recipient;
6. pins `SendUsingAccount` to the source account;
7. discards the native forwarded MeetingItem unsent.

The Phase 3 UI still cannot send a meeting. Real send remains behind the already-proven Phase 2 mechanism until a deliberate GUI confirmation/error flow is implemented in a later focused PR.

## Diagnostics and partial failures

Normal workflows describe meetings, accounts, states and actions. Technical identifiers live in Diagnostics.

Diagnostics currently includes:

- selected item native `GlobalAppointmentID`, `EntryID`, and `StoreID`;
- managed-copy classification and property values;
- per-account scan result summaries;
- count of valid managed copies and invalid/unreadable managed metadata;
- OutlookHost warnings and safe failure output.

Calendar/Alignment surface a user-facing warning when a scan is partial while directing technical investigation to Diagnostics.

## Temporary UI/host transport

The application currently launches `OutlookAligner.OutlookHost.exe` as a child process and exchanges JSON/text over redirected standard streams.

This remains a transitional local transport. The important architecture boundary is already enforced:

- all Outlook COM remains inside OutlookHost;
- WinUI receives only plain DTOs/results;
- OutlookHost remains an interactive STA process;
- no COM RCW crosses into WinUI;
- the UI/host contract is explicitly versioned.

Before production write workflows are considered complete, transport should converge on the planned local IPC/named-pipe host without changing the view-model/Core contracts.

## Distribution for development testing

CI publishes a self-contained unpackaged win-x64 bundle containing the WinUI application and sibling OutlookHost.

Green CI has already proved that the application can be built and assembled through GitHub Actions. Development builds may override the host executable with `OUTLOOK_ALIGNER_HOST_PATH`.

## Acceptance gate for this increment

Before this increment is merged, verify through hosted CI and the real Windows/Classic Outlook profile that:

1. the packaged development artifact launches the WinUI app;
2. Outlook account/calendar refresh works without manual IDs;
3. Calendar event selection and Diagnostics update correctly;
4. Check Forward still behaves as a capability probe only;
5. Prepare Forward creates/discards the native forwarded meeting without sending;
6. Alignment loads without crashing and recurring rows remain explicitly unresolved;
7. authority selection behaves sensibly for ordinary rows;
8. no Move/Copy write control exists;
9. existing Outlook appointments are not modified by refresh/alignment/preview activity;
10. partial failures remain visible rather than blanking the whole UI.

Managed-copy-specific UI behavior can only be manually exercised once real managed copies exist. Its Core/read-side safety is therefore primarily protected by hosted tests in this increment.

## Subsequent Phase 3 increments

After this PR-sized increment is tested/merged, keep subsequent work narrow:

1. real-machine repair findings from this UI increment, if any;
2. richer Calendar/FullCalendar integration only after a real Node 24 lockfile can be generated and CI switched to `npm ci`;
3. deliberate GUI confirmation/send flow for native Forward;
4. further loading/degraded-state/accessibility hardening;
5. persistence and production IPC in focused slices when required by the next write workflow.

Copy Full, Copy Busy, Move writes, and recurrence writes remain behind their dedicated safety/implementation gates rather than being smuggled into Phase 3.

## Safety invariants

- No automatic PR merge.
- Phases may be split into smaller PRs.
- No vCalendar fallback.
- No Graph dependency.
- No Move write command in the current Phase 3 UI.
- No real Forward send from the current UI.
- Managed-copy metadata is read without creating or modifying properties.
- Invalid/incomplete/unsupported/unreadable managed metadata fails closed and appears unresolved/conflicting.
- Only validated managed copies can appear as Move candidates.
- Authority is explicit or proven by consistent managed provenance and never inferred from latest modification time.
- Recurring Move remains blocked.
- Technical Outlook identifiers remain in Diagnostics rather than normal workflows.
