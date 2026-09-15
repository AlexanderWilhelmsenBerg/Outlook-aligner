# Phase 3 — UI-assisted read/alignment foundation

Status: **Foundation increment complete pending PR #6 merge. Future product work follows the Segment A/B roadmap in `plan.md`.**

Started: 2026-09-08

Roadmap decision updated: 2026-09-11

## Purpose

Phase 1 and Phase 2 proved the Classic Outlook COM boundary and genuine native meeting forwarding, but continued development was inefficient while manual testing required copying Outlook IDs and addresses into PowerShell commands.

This increment brought forward a usable WinUI application so Outlook state, identity, correlation, authority, warnings and diagnostics could be inspected through the product itself.

The owner has now validated that the packaged application launches and is useful for the initial read-only use case. The project therefore moves on to a complete read-only product experience before reconciliation/write behavior is tested or productized.

## Delivered foundation

PR #6 contains:

- real WinUI 3 `App.xaml` / `MainWindow` application;
- NavigationView shell;
- GUI-driven Classic Outlook account/calendar discovery;
- bounded Outlook calendar reads without manual IDs;
- Calendar event selection and detail display;
- preliminary non-recurring logical correlation;
- explicit `Aligned`, `Missing`, `Moved`, `DetailsDifferent`, `Duplicate`, `Conflict`, recurrence-unresolved and uncorrelated states;
- read-only Outlook Aligner managed-copy metadata inspection;
- fail-closed behavior for suspicious/incomplete/unreadable managed metadata;
- session-only authority selection and safe known-origin inference from validated provenance;
- read-only Move Selected planning with no Move execution;
- native Forward capability/prepare-and-discard diagnostics with no GUI send;
- structured persistent diagnostics log;
- self-contained win-x64 UI + OutlookHost CI artifact;
- local updater with portable GitHub CLI support;
- optional local Authenticode signing using an already-trusted current-user test certificate.

## OutlookHost protocol

The current UI/host contract is protocol **v3**.

The contract carries the read metadata needed by the UI, including managed-copy state and meeting-status classification. The UI validates the protocol before consuming host results.

The current child-process/stdout JSON transport remains transitional. The architectural invariant is what matters:

- Outlook COM stays inside OutlookHost;
- WinUI receives plain DTOs/results only;
- OutlookHost remains an interactive STA process;
- COM objects never cross into UI/Core.

## Managed-copy metadata — read only

Existing custom properties inspected by the read model:

- `OutlookAligner.SyncGroupId`
- `OutlookAligner.SourceGlobalAppointmentId`
- `OutlookAligner.SourceAccountId`
- `OutlookAligner.CopyType`
- `OutlookAligner.SchemaVersion`

The reader does not create properties or call `Save()`.

Metadata classifications:

- `None`
- `Valid`
- `Incomplete`
- `UnsupportedSchema`
- `Unreadable`

Invalid/suspicious metadata fails closed into conflict/unresolved behavior rather than being silently treated as trusted ordinary Outlook data.

## Correlation and authority foundation

Ordinary non-recurring items can correlate by native `GlobalAppointmentID`.

Validated managed copies correlate through their recorded source GlobalAppointmentID because their own Outlook GlobalAppointmentID may differ.

Recurring identity remains deliberately incomplete in this foundation; Segment A Phase 3 now owns the full recurring read-model work.

Authority is event-specific. The current implementation can use validated managed provenance or explicit session selection. It never uses `LastModificationTime` as latest-wins.

The revised roadmap further clarifies that scan/enumeration order is not origin evidence.

## Forward foundation

Phase 2 already proved real native forwarding. This branch retains two safe UI diagnostics:

- **Check Forward** — capability probe only.
- **Prepare Forward (discard)** — creates/prepares the native forwarded MeetingItem and discards it unsent.

There is no GUI send action.

Following the 2026-09-11 roadmap decision, Forward reliability/product testing is **deferred to Segment B**. It is not a merge gate for PR #6 and must not distract from completing the read-only product.

## Move preview foundation

Alignment can build a read-only Move preview for safe managed-copy cases. There is no Move command and no Outlook mutation.

This planner remains useful groundwork for Segment B, but it is not part of Segment A acceptance.

## Diagnostics

Diagnostics includes:

- selected event StoreID / EntryID / GlobalAppointmentID;
- meeting/recurrence classification;
- managed-copy metadata/classification;
- partial scan warnings;
- OutlookHost diagnostics;
- persistent structured event log with Debug/Info/Warning/Error filtering.

Normal UI keeps these technical details out of the primary workflow.

## Development distribution

CI publishes a self-contained unpackaged win-x64 bundle containing the WinUI application and sibling OutlookHost.

`scripts/Update-OutlookAlignerTestApp.cmd` / `.ps1` can fetch the newest successful artifact into `%LOCALAPPDATA%\OutlookAligner\TestApp`, retain one previous build, and launch it.

When a trusted `CN=Outlook Aligner Local Test` current-user code-signing certificate already exists, the updater locally signs Outlook Aligner EXE/DLL files and verifies `Valid` Authenticode status before launch. Trust creation is intentionally not automatic.

## Merge gate for PR #6

The owner has already confirmed the key read-only foundation behavior needed to merge this increment:

1. packaged app launches on the real work computer;
2. account/calendar reading works for the initial overview use case;
3. no GUI Forward send exists;
4. no Copy/Move execution exists;
5. local update/install workflow works;
6. read-only usage does not intentionally mutate source appointments.

Forward/Prepare behavior, recurring reconciliation, and Move-preview depth are no longer required manual merge tests for this foundation PR. They are deferred to the appropriate later segment.

The merge itself remains owner-controlled; never merge automatically.

## Next work

After PR #6 is merged, stop using the old technical Phase 3/4/5 sequence for product planning. Follow `plan.md`:

### Segment A — Read-only overview

1. Month calendar view.
2. Account dots + green/yellow/red alignment coloring.
3. Recurring meeting/occurrence identity + lower-right recurrence symbol.
4. Read-only reconciliation report for Missing/Moved/Duplicate/Conflict.
5. Calendar/Report view modes and filters.

### Segment B — Reconciliation

Only after Segment A is manually accepted: action previews, Forward, Copy Full/Busy, Move Selected, bulk reconciliation, then recurring writes.

## Safety invariants

- No automatic PR merge.
- Segment A remains persistently read-only.
- No Graph dependency.
- No vCalendar fallback disguised as native Forward.
- No Move/Copy execution in this foundation.
- No GUI Forward send in this foundation.
- Invalid managed metadata fails closed.
- Recurrence writes remain blocked until recurring identity is proven.
- Technical Outlook identifiers stay in Diagnostics rather than normal workflows.
