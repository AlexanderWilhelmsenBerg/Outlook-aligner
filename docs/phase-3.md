# Phase 3 — UI-assisted development and testing foundation

Status: **In progress 🚧 — current branch is based on merged Phase 2**

Started: 2026-09-08

## Purpose

Phase 1 and Phase 2 proved the Classic Outlook COM boundary and genuine native meeting forwarding, but continued development became inefficient when manual testing required copying `GlobalAppointmentID`, `EntryID`, source SMTP addresses, and target addresses into PowerShell commands.

Phase 3 brings a usable WinUI application forward so identity, authority and safety behavior can be inspected through the application before write features are introduced.

The current Phase 3 application remains deliberately conservative: Calendar and Alignment can discover and reason about Outlook state, native Forward can be checked/prepared-and-discarded, and Move Selected can be previewed, but there is still no Move write command and the UI cannot send a meeting.

## Current UI

The unpackaged WinUI 3 application now has four functional destinations:

- **Calendar** — reads the bounded Outlook calendar window, selects a meeting, derives its source account, offers the other discovered accounts as targets, and exposes the safe native-Forward test actions.
- **Alignment** — groups safely correlated non-recurring meetings, shows discrepancy states, lets the user explicitly select authority, identifies validated Outlook Aligner-managed copies, and builds a read-only Move Selected preview.
- **Settings** — owns the current scan horizon and shows discovered Classic Outlook accounts/calendar availability.
- **Diagnostics** — contains raw host output, Outlook IDs, managed-copy metadata and partial-failure details that should not appear in normal workflows.

Internal Outlook identifiers are therefore no longer normal user input.

## OutlookHost protocol v2

Phase 3 introduces OutlookHost protocol version **2** for managed-copy metadata.

The UI validates the protocol before consuming scan output. An older protocol-v1 OutlookHost cannot silently feed the new UI a calendar result that lacks the ownership information needed by Copy/Move safety logic.

## Managed-copy metadata — read only

The calendar reader now looks for these existing custom appointment properties:

- `OutlookAligner.SyncGroupId`
- `OutlookAligner.SourceGlobalAppointmentId`
- `OutlookAligner.SourceAccountId`
- `OutlookAligner.CopyType`
- `OutlookAligner.SchemaVersion`

The Phase 3 reader only calls Outlook's property lookup APIs. It does not add properties, set values, call `Save()`, or otherwise modify an appointment.

Metadata is classified as:

- `None`
- `Valid`
- `Incomplete`
- `UnsupportedSchema`
- `Unreadable`

Only `Valid` metadata is allowed to affect managed-copy identity or Move preview eligibility. Incomplete, unknown-schema and unreadable metadata fail closed and are treated as non-managed for mutation safety.

Current managed-copy schema version is `1`. Supported copy-type markers are `Full` and `Busy`.

## Correlation semantics

Ordinary non-recurring Outlook items continue to correlate by native `GlobalAppointmentID`.

A validated Outlook Aligner-managed local copy is different: a locally created copy may have its own native Outlook `GlobalAppointmentID`, so its logical correlation key is the validated stored `OutlookAligner.SourceGlobalAppointmentId`.

This means:

- a source appointment and a valid managed copy can form one logical row even when the copy's own native GlobalAppointmentID differs;
- an unmanaged appointment never gets to influence correlation merely because some similarly named custom field is present;
- a managed item without a validated source GlobalAppointmentID remains `Uncorrelated` rather than falling back to a potentially unsafe guess;
- recurring items remain deliberately unresolved until recurrence identity is proven.

The existing discrepancy states remain conservative: `Aligned`, `Missing`, `Moved`, `DetailsDifferent`, `Duplicate`, `Conflict`, `RecurrenceIdentityUnresolved`, and `Uncorrelated`.

## Authority

Alignment lets the user explicitly select the account that should control a logical meeting.

Authority selection is session-only in the current slice and does not modify Outlook. The app does not infer authority from `LastModificationTime`, and it does not pretend that pre-existing copies reveal their origin.

Authority cannot be chosen for recurrence-unresolved, uncorrelated, or duplicate rows.

## Move Selected preview

Phase 3 now contains a pure Core `MovePreviewPlanner` plus an Alignment UI preview.

The planner can propose a Move only when all safety prerequisites hold. It:

1. requires exactly one observed member in the selected authority account;
2. never proposes the authority item as a target;
3. blocks duplicate members;
4. blocks recurring writes;
5. considers only validated Outlook Aligner-managed non-authority copies eligible for change;
6. skips pre-existing/unmanaged Outlook appointments;
7. skips copies already aligned with the authority;
8. records the current and proposed Start/End values for every candidate action.

The Alignment UI renders the resulting plan as plain-language `MOVE`, `SKIP`, or `BLOCKED` lines. It does not show locator IDs in this normal workflow.

**There is no Move write command in this Phase 3 slice.** The preview is intentionally useful before mutation exists.

## Forward testing from the UI

For a selected Calendar event the UI currently offers:

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

The Phase 3 UI still cannot send a meeting. Real send remains behind the explicit Phase 2 CLI confirmation gate until a deliberate GUI confirmation/error model is implemented.

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

A green Phase 3 bundle has already proved that the application can be built and assembled through GitHub Actions. Development builds may override the host executable with `OUTLOOK_ALIGNER_HOST_PATH`.

## Remaining Phase 3 work

The next UI-foundation work should remain narrow and testable:

1. keep the managed-copy/Move preview UI green through the packaged WinUI build;
2. improve user-facing partial-failure/error presentation where raw Diagnostics is still doing too much work;
3. decide how intentional `Copy Busy` privacy differences should be represented in Alignment before Copy Busy itself is implemented;
4. add deliberate GUI confirmation only when real native Forward send is ready to leave the Phase 2 CLI gate;
5. integrate the richer FullCalendar/WebView2 calendar only when the frontend can also satisfy the project's real npm lockfile/`npm ci` requirement — never fabricate a lockfile;
6. leave Copy Full, Copy Busy, Move writes and recurrence writes behind their dedicated implementation/safety gates rather than using Phase 3 to smuggle in mutation behavior.

## Safety invariants

- No automatic PR merge.
- No vCalendar fallback.
- No Graph dependency.
- No Move write command in the current Phase 3 UI.
- No real Forward send from the current UI.
- Managed-copy metadata is read without creating or modifying properties.
- Invalid/incomplete managed metadata fails closed.
- Only validated managed copies can appear as Move candidates.
- Authority is explicit and never inferred from latest modification time.
- Recurring Move remains blocked.
- Technical Outlook identifiers remain in Diagnostics rather than normal workflows.
