# Agent 10 domain note — A1.3

## Role and mode

**Agent 10 — Outlook Domain / Identity & Calendar Semantics**, A1.3 planning-readiness review only. No source implementation was performed.

## Repository state reviewed

- Repository: `AlexanderWilhelmsenBerg/Outlook-aligner`
- `main` HEAD used as the implementation baseline: `7dbd437c1ee86bf6e41d8ad939a54cc81ee7255e`
- Planning PR: `#11`
- Planning branch: `planning/a1.2-a1.3`
- Initial planning HEAD inspected: `34f52541a4cf7a5a6e0752e068e7e3b169b2a650`
- Refreshed planning HEAD before this note was written: `a037b7a1761b258a6563e52de0cc677e887653b5`

The branch moved during this review. The refreshed state was used before editing the existing placeholder `slice planning/A1.3/10-domain-note.md`.

## Evidence inspected

- `AGENTS.md`
- `plan.md`, especially Slice A1.3 and the A2 authority boundary
- `slice planning/A1.3/00-brief.md`
- `slice planning/A1.3/10-domain-prompt.md`
- `src/OutlookAligner.Core/Alignment/EventCorrelation.cs`
- `src/OutlookAligner.Outlook.Contracts/ProbeDtos.cs`
- `src/OutlookAligner.OutlookHost/Calendars/OutlookCalendarReader.cs`
- `src/OutlookAligner.OutlookHost/Calendars/ManagedCopyMetadataReader.cs`
- `src/OutlookAligner.App/ViewModels/MainViewModel.cs`, specifically the Outlook DTO -> `ObservedCalendarEvent` projection and partial-result handling
- `tests/OutlookAligner.Core.Tests/EventCorrelationTests.cs`
- `tests/OutlookAligner.Core.Tests/ManagedCorrelationTests.cs`
- `tests/OutlookAligner.OutlookHost.Tests/ManagedCopyMetadataReaderTests.cs`
- `docs/phase-1.md`
- `docs/phase-2.md`

## A1.3 domain boundary

A1.3 may standardize and expose the existing Core correlation behavior as the shared logical-event read model used by Calendar and later Report, but it must not broaden the identity evidence beyond what is already proven for supported non-recurring items.

A1.3 owns only the question: **which observed non-recurring calendar items may be represented as one logical event?**

A1.3 does **not** decide:

- which account is authoritative/origin for an ambiguous logical event;
- whether a time or detail difference is healthy/unhealthy for A2 presentation;
- any latest-wins rule;
- recurrence occurrence identity;
- reconciliation/write eligibility.

Those remain later contracts. In particular, correlation must not implicitly become authority.

## Exact supported non-recurring correlation contract

### 1. Native positive identity evidence

For an ordinary, non-recurring observation with no Outlook Aligner managed-copy metadata issue:

- trim `GlobalAppointmentId` for correlation purposes;
- a non-empty `GlobalAppointmentId` is accepted as the positive native correlation key for A1.3;
- observations with the same normalized value may belong to one logical-event group;
- comparison is ordinal; do not case-fold or otherwise rewrite the ID;
- `StoreID`, `EntryID`, scan order, subject, location, start/end time, organizer-like display text, or modification time must not be used as substitutes for logical identity.

`EntryID` remains a locator for one Outlook item. It may be used to keep an unresolved observation individually selectable, but it must never make two items the same logical event.

### 2. Managed-copy positive identity evidence

A managed copy is eligible for A1.3 correlation only when its metadata was classified `ManagedCopyState.Valid` by the OutlookHost metadata boundary.

For such an item:

- correlate by the validated `SourceGlobalAppointmentId`, not by the copy's own native `GlobalAppointmentId`;
- require the validated managed-copy fields already enforced by `ManagedCopyMetadataReader`: parseable `SyncGroupId`, non-empty `SourceGlobalAppointmentId`, non-empty `SourceAccountId`, supported `CopyType`, and current schema version;
- do not trust a `ManagedSourceGlobalAppointmentId` field on an item that was not classified as a valid managed copy.

This permits an ordinary source item and its valid managed copy to produce one logical-event group when the source native `GlobalAppointmentId` equals the managed copy's validated `SourceGlobalAppointmentId`.

### 3. One representation does not mean silent collapse

For a positively correlated identity key, A1.3 may expose **one logical-event presentation item** containing all member observations, but it must preserve every member in the read model.

If more than one member with the same correlation key exists for the same account, the group is `Duplicate` and all duplicate members remain inspectable. The presentation must not pick one duplicate and discard the others.

A duplicate group may still be rendered as one logical-event row/item because the logical identity is known; its unresolved duplicate condition must remain explicit in the shared logical-event model/detail surface.

### 4. Managed-provenance conflicts remain one explicit conflict group only when positive identity agrees

If valid managed copies correlate to the same positive source identity but disagree about validated managed provenance fields that are already semantically checked by Core (for example source account or sync group), the result is `Conflict`.

The conflict must not be repaired by choosing one metadata value, using scan order, or falling back to the copies' native IDs.

Conversely, an item whose Outlook Aligner metadata is incomplete, unsupported, or unreadable does **not** contribute to a correlated group merely because its native `GlobalAppointmentId` happens to match another item. It remains an isolated `Conflict` observation.

### 5. No positive identity -> no correlation

A non-recurring item remains `Uncorrelated` when:

- an ordinary observation has no non-blank native `GlobalAppointmentId`; or
- a valid managed copy lacks the validated source identity needed for correlation.

A1.3 must not guess correlation from matching subject, location, time, duration, attendee-visible text, account position, or any combination of those fields.

Two visually identical uncorrelated items therefore remain two logical-event presentation items.

### 6. Suspicious Outlook Aligner metadata fails closed

Any observation whose metadata state is `Incomplete`, `UnsupportedSchema`, or `Unreadable` is suspicious for correlation purposes.

Such an observation must:

- be surfaced independently as `Conflict`;
- retain its own locator-backed presentation identity;
- not participate in native-`GlobalAppointmentId` grouping;
- not be reclassified as an ordinary unmanaged item;
- not authorize any mutation or later reconciliation action.

The existing fail-closed behavior is correct and is part of the A1.3 contract.

## Recurrence boundary — explicitly unresolved in A1.3

Every observation with `IsRecurring == true` remains outside supported A1.3 correlation, regardless of whether Outlook exposes the same `GlobalAppointmentId` on observations that appear related.

Each recurring observation must remain independently represented with `RecurrenceIdentityUnresolved` until a later recurrence slice defines an occurrence identity that distinguishes at minimum:

- series identity from occurrence identity;
- normal generated occurrences;
- modified exceptions;
- moved exceptions while preserving them as the same logical occurrence;
- deleted occurrences/absence evidence;
- recurrence behavior across accounts.

A1.3 must not group recurring items by series `GlobalAppointmentId`, start time, ordinal occurrence number, subject, or any other heuristic. `GlobalAppointmentId` alone is not accepted as universal recurrence-occurrence identity.

## Partial-read and incomplete-scan behavior

Positive correlation may be built from observations that were successfully read, even when the overall refresh reports partial results. Partial-read state must not cause two positively identified observed items to be split apart.

However, **absence is not positive evidence**. Therefore A1.3 must not strengthen a partial scan into a claim that an unread item/account is definitely missing.

The current host/app behavior has two relevant failure levels:

1. **Account/calendar failure** — accounts whose calendar is unavailable are already excluded from `expectedAccountKeys`. This is fail-closed for missing-account classification and must be preserved.
2. **Per-item read warning within an otherwise available calendar** — `OutlookCalendarReader` currently skips an unreadable appointment and emits only a warning. Because the app cannot identify which logical event that skipped item belonged to, A1.3 must treat any absence-derived classification from that scan as potentially incomplete rather than as proven negative evidence.

Required A1.3 behavior for partial scans:

- keep successful positive correlation deterministic;
- preserve a scan-level partial/incomplete indicator from the existing account errors/warnings so Calendar/detail can remain truthful;
- do not invent an unseen member;
- do not merge an unresolved item into a group because another account had a read error;
- do not interpret missing observation data as deletion evidence or reconciliation permission;
- if the shared logical-event DTO exposes `MissingAccountKeys` or a derived `Missing` state while the scan contains a per-item read warning capable of hiding an appointment, it must also expose enough scan-completeness context for the consumer to avoid presenting that absence as certain.

This does **not** require A1.3 to design A2 health colors or authority wording. It does require the A1.3 read model not to erase the fact that the underlying scan was partial.

## Determinism requirements

For the same set of observed DTOs and scan-completeness input, A1.3 grouping must be deterministic and independent of enumeration order.

In particular:

- group membership is decided only by the positive identity rules above;
- account-key comparisons may remain case-insensitive as currently implemented;
- duplicate detection is per normalized account key within an already-established logical identity;
- display subject selection and presentation ordering must not alter group identity;
- no first-seen item becomes a master/origin by virtue of being first.

The existing `gaid:{identity}` and per-item unresolved keys are suitable internal grouping keys for A1.3 provided consumers treat them as opaque read-model identifiers rather than Outlook-native authority evidence.

## Required implementation tests for Agent 40

At minimum, A1.3 implementation must preserve/add deterministic tests proving:

1. same native non-recurring `GlobalAppointmentId` across two or more accounts -> one logical event with all members;
2. valid managed copy `SourceGlobalAppointmentId` matching the source native ID -> one logical event;
3. a managed copy's own different native ID does not prevent correlation to its validated source ID;
4. an unmanaged item's stray managed-source field is ignored;
5. missing/blank native ID on ordinary non-recurring items -> separate `Uncorrelated` logical items;
6. visually identical uncorrelated items are never guessed together;
7. duplicate same-identity members in one account -> one group marked `Duplicate`, with all members preserved;
8. valid managed copies with conflicting source-account metadata -> `Conflict`;
9. valid managed copies with conflicting sync-group metadata -> `Conflict`;
10. incomplete managed metadata -> isolated `Conflict`, even if native ID matches a healthy source item;
11. unsupported-schema metadata -> isolated `Conflict`;
12. unreadable metadata -> isolated `Conflict`;
13. recurring observations sharing an apparent identity -> remain separate `RecurrenceIdentityUnresolved` items;
14. correlation output/group identity is stable under input enumeration reordering;
15. an unavailable/error account is not treated as proven missing membership;
16. a per-item read warning preserves partial-scan truth alongside any absence-derived result, rather than turning skipped data into certain absence;
17. technical `EntryID`/StoreID locators do not influence cross-account correlation;
18. subject/time/location equality alone never correlates observations.

Existing Core and OutlookHost tests already cover important portions of items 1-13. A1.3 implementation should extend rather than replace those tests and add the shared-presentation/partial-scan coverage introduced by this slice.

## Findings and blockers

### Domain decision: positive identity evidence is sufficient for supported A1.3 non-recurring correlation

The repository already contains a coherent fail-closed correlation core for native `GlobalAppointmentId` and validated managed-copy source identity. No new heuristic identity rule is required for A1.3.

### Domain requirement: preserve partial-scan uncertainty

The current app visibly reports partial results, but `EventCorrelation.BuildGroups` itself receives only observed items plus expected account keys. A1.3's shared logical-event presentation contract must not discard scan-completeness/error context when it starts driving the Calendar surface. Agent 30 may choose the layer/DTO shape; the domain requirement is simply that absence derived from an incomplete scan cannot be presented as certain.

This is a requirement for implementation, not a blocker requiring a new identity rule.

### Deferred by design

- recurrence occurrence identity;
- A2 authority/origin semantics;
- A2 account-status/color semantics;
- any write/reconciliation eligibility.

These deferrals are explicit and must not be filled in by Agent 40 during A1.3.

## Readiness

The supported non-recurring identity rules, fail-closed unresolved cases, duplicate/conflict behavior, managed-copy handling, partial-read constraint, and recurrence boundary are defined precisely enough that Agent 40 does not need to invent a domain rule for A1.3.

**DOMAIN READY**
