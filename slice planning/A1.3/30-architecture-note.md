# Agent 30 architecture note — A1.3

Status: **ARCH READY, provisional until A1.2 is accepted/merged and Agent 00 revalidates this note against the implemented A1.2 seam**

## Role and mode

Agent 30 — Architecture / Integration & Data Contracts, A1.3 readiness review only. No source implementation was performed.

## Exact state reviewed

- Repository: `AlexanderWilhelmsenBerg/Outlook-aligner`
- `main` HEAD reviewed: `7dbd437c1ee86bf6e41d8ad939a54cc81ee7255e`
- Planning PR: #11
- Planning branch: `planning/a1.2-a1.3`
- Planning HEAD initially read: `34f52541a4cf7a5a6e0752e068e7e3b169b2a650`
- Planning HEAD refreshed immediately before this note was committed: `34f52541a4cf7a5a6e0752e068e7e3b169b2a650`
- Agent 10 note at that refresh: only the placeholder `PENDING AGENT 10 INPUT` was present, so there was no completed Agent-10 result to incorporate. This architecture deliberately treats Agent 10's identity/correlation decisions as an input contract rather than duplicating or inventing them.
- A1.2 implementation branch existed at the same `7dbd437c1ee86bf6e41d8ad939a54cc81ee7255e` head at refresh time; therefore this review uses the accepted A1.2 planning contract, not an unaccepted implementation assumption.

## Evidence inspected

- current `AGENTS.md`;
- `plan.md`, especially Slice A1.3 and the Segment-A ownership rules;
- `slice planning/A1.3/00-brief.md`;
- `slice planning/A1.3/30-architecture-prompt.md`;
- `slice planning/A1.3/10-domain-note.md` at final refresh (placeholder only);
- accepted A1.2 reconciliation contract in `slice planning/A1.2/00-reconciliation.md`;
- A1.2 architecture note in `slice planning/A1.2/30-architecture-note.md`;
- `src/OutlookAligner.Core/Alignment/EventCorrelation.cs`;
- `src/OutlookAligner.App/ViewModels/MainViewModel.cs`;
- `src/OutlookAligner.App/ViewModels/CalendarEventRowViewModel.cs`;
- `src/OutlookAligner.App/ViewModels/AlignmentGroupViewModel.cs`;
- `src/OutlookAligner.Outlook.Contracts/ProbeDtos.cs`;
- `tests/OutlookAligner.Core.Tests/EventCorrelationTests.cs` and the current Core-test layout.

## Current architecture facts relevant to A1.3

1. OutlookHost already returns plain transport DTOs. COM objects do not need to cross into Core, App read models, WebView messages, or tests.
2. `CalendarEventDto` still contains transport/domain inputs that are not presentation-safe as a wholesale WebView payload, including `EntryId`, `GlobalAppointmentId`, and managed-copy metadata.
3. Core already contains the correlation seam: App maps host observations to `ObservedCalendarEvent`, then `EventCorrelation.BuildGroups(...)` produces `LogicalEventGroup` values with explicit `AlignmentState`, members, missing-account keys, difference flags, and explanation.
4. `MainViewModel` currently builds those groups separately from its per-observation `Events` collection and exposes them through `AlignmentGroups`. That means the codebase already has one domain truth source but two UI projections.
5. A1.2's accepted contract intentionally renders per-observation items through an opaque, session-scoped presentation ID and keeps FullCalendar presentation-only. A1.3 should replace only the render/selection projection, not move correlation into TypeScript.
6. Existing `LogicalEventGroup`/`ObservedCalendarEvent` records are Core/domain records, not the final WebView contract. In particular, members contain `LocatorKey` and native/domain identity inputs that should remain outside normal Calendar/Report presentation contracts.

## Architecture contract for A1.3

### 1. Keep correlation/domain identity in Core

The A1.3 pipeline must remain:

`OutlookHost plain DTOs -> App host-to-Core mapping -> Core EventCorrelation/domain result -> App shared logical-event read model -> Calendar/Report presentation`

A1.3 must **not**:

- correlate in TypeScript/FullCalendar;
- regroup records in the WebView based on subject, time, account, or IDs;
- introduce a second correlation implementation in App;
- infer identity from `StoreID + EntryID`, scan order, or modification time;
- reinterpret unresolved Core states as correlated merely to obtain one visual row.

Agent 10 owns the exact rules that determine which observations Core may place in the same `LogicalEventGroup`. Agent 40 must consume that accepted domain behavior rather than add architecture-side identity heuristics.

### 2. Introduce one App-owned shared logical-event read model

Create one immutable/read-only App presentation read model projected from each Core `LogicalEventGroup`. Naming is implementation detail, but the contract should be equivalent to:

```text
LogicalEventReadModel
- presentationId: string
- displaySubject: string
- state: explicit enum/value copied from the Core result
- explanation: string
- startLocal: DateTime
- endLocal: DateTime
- isAllDay: bool
- isRecurring: bool
- timesDiffer: bool
- detailsDiffer: bool
- members: LogicalEventMemberReadModel[]
- missingAccountKeys: string[]

LogicalEventMemberReadModel
- accountKey/display account identity needed by the UI
- startLocal: DateTime
- endLocal: DateTime
- isAllDay: bool
- subject: string?
- location: string?
- busyStatus/sensitivity only if required by the existing comparison/detail surface
- recurrence/presentation facts required by that surface
```

The precise field set may be smaller if existing UI proves fields unnecessary. The important boundary is that this read model contains **display/comparison facts, not raw Outlook locators or native correlation keys**.

Specifically exclude from the shared Calendar/Report read model and ordinary WebView payload:

- StoreID;
- EntryID / `ObservedCalendarEvent.LocatorKey`;
- GlobalAppointmentID and managed source GlobalAppointmentID;
- SyncGroupId and other managed-copy technical identifiers;
- Outlook COM objects/RCWs.

If Diagnostics still needs those technical values, keep its mapping/reference separate from the normal shared presentation model.

### 3. Projection rules are structural, not a new domain engine

The App projector may choose deterministic display values only where no identity/authority inference is involved. It must primarily copy already-decided Core truth.

For timing used to place the one calendar item:

- if the Core group exposes one agreed time (`TimesDiffer == false`), use that common member time;
- if the Core/domain result is explicitly unresolved/different, the projector must preserve that state and must not silently choose a member as authoritative merely because it is first in a list.

If Agent 10's accepted note requires a specific calendar-placement rule for a multi-member unresolved/different-time group, that rule must be consumed here. Until then, architecture does not define one.

This is the only domain-dependent gap in the concrete field mapping; it is intentionally delegated to Agent 10 rather than solved by scan order.

### 4. Explicit unresolved cases remain first-class read-model items

Every Core `LogicalEventGroup` result maps to exactly one shared read-model item. The App must preserve its explicit state/explanation.

Therefore groups produced as `Duplicate`, `Conflict`, `RecurrenceIdentityUnresolved`, or `Uncorrelated` remain visible and selectable as unresolved truth. A1.3 must not merge two separate unresolved Core groups together in App or JavaScript.

If Core returns one unresolved group per observation, Calendar/Report preserve that one-for-one result. If Agent 10 later changes the accepted Core grouping contract, the shared projector follows the Core result without requiring a new frontend correlation algorithm.

### 5. Calendar and Report consume the same collection

`MainViewModel` (or a narrow App read-model service used by it) should own one refreshed collection of `LogicalEventReadModel` values for the scan result.

Calendar and the later Report must bind/project from this same collection. Do not retain a Calendar-only logical-event model and later create a second Report correlation/mapping path.

The current `AlignmentGroups` projection can be migrated to wrap/consume the shared read model rather than independently representing the same Core group. The exact class name is not important; the invariant is one Core correlation result and one App shared read-model projection per refresh.

The existing per-observation collection may remain temporarily for Diagnostics or bounded migration if needed, but it must stop being the source used to decide Calendar logical items once A1.3 is active.

### 6. WebView contract remains presentation-only and smaller than the shared read model

A1.3 should evolve A1.2's versioned render payload from an observation item to a logical-calendar item, for example:

```text
CalendarLogicalEvent
- presentationId: string
- subject: string
- startLocal: string
- endLocal: string
- isAllDay: bool
- isRecurring: bool
- state/display-status token only if A1.3 visibly requires it
```

Do not send the entire `LogicalEventReadModel` into JavaScript merely because it exists in App. Account members, explanations, technical comparison fields, and diagnostics stay native unless the WebView genuinely needs them to render the month item.

The A1.2 bridge invariants remain:

- versioned structured JSON;
- opaque App-owned session `presentationId`;
- refresh atomically replaces the current ID map;
- unknown/stale IDs fail closed;
- no executable JavaScript string interpolation;
- no raw Outlook IDs in browser messages.

If A1.3 does not require a protocol-version bump under the implementation's chosen envelope/versioning policy, the message type/payload may evolve compatibly; otherwise bump explicitly and test both rejection and supported parsing. Do not silently reinterpret an old message shape.

### 7. Selection/detail wiring

The App-side `presentationId` map for Calendar should now resolve to the shared `LogicalEventReadModel` (or its stable in-memory object/reference for the current refresh), not to a single `CalendarEventRowViewModel`.

On `logicalEventSelected`/equivalent client selection:

1. validate version/type/presentation ID through the existing A1.2 bridge seam;
2. resolve the ID only against the current refresh map;
3. update the existing logical comparison/detail selection (`SelectedAlignmentGroup` or its A1.3 replacement backed by the shared read model);
4. present the members/state/explanation from that same selected logical event.

A Calendar selection must not arbitrarily set one member as authority or use the first member as the selected Outlook item. If a diagnostics-only member inspection remains, it is a secondary explicit action/state, not the logical-event selection contract.

This gives Calendar now and Report later one selection identity inside the App while keeping Outlook locators private to native diagnostic/host code.

### 8. Refresh/lifecycle behavior

A scan refresh should construct the whole logical-event snapshot off to the side, then replace the shared collection and WebView `presentationId` map coherently. Avoid incrementally mutating the browser map while the underlying correlation result is being rebuilt.

Selection from the previous snapshot becomes stale after replacement unless the implementation deliberately reselects by an App-owned read-model key proven safe for that purpose. A1.3 does not require persistence of selection across refresh, and raw Outlook IDs must not be promoted to presentation identity to achieve it.

Month navigation remains local presentation behavior from A1.2 and must not recalculate correlation or trigger an Outlook scan.

### 9. Layer placement

Smallest accepted placement:

- **Outlook.Contracts / OutlookHost:** unchanged unless Agent 10 proves a missing observation fact is required. No A1.3 presentation concerns belong here.
- **Core:** existing correlation/domain records and rules; domain changes only as required by Agent 10's accepted contract and covered by Core tests.
- **App:** host-to-Core observation mapping, one shared logical-event projector/read model, current-refresh presentation-ID map, Calendar bridge mapping, selection/detail state.
- **Web frontend:** render/navigate/select only. No account correlation, identity, authority, or reconciliation logic.
- **Persistence:** no A1.3 requirement. Do not persist locators, presentation IDs, or authority choices merely to implement this slice.

A narrow projector/service outside `MainViewModel` is preferred if it keeps correlation/mapping/test logic independently testable; moving the whole domain pipeline into a new service framework is unnecessary churn.

## Migration from A1.2 observation rendering

Agent 40 should make the migration narrowly:

1. keep A1.2 WebView lifecycle, asset loading, native navigation, range reporting, failure handling, and defensive message parser;
2. after host refresh, build Core observations once and call the accepted Core correlation seam once;
3. project the returned groups once into the shared App logical-event collection;
4. assign one fresh opaque `presentationId` to each projected logical item;
5. send a Calendar-sized render projection of those logical items instead of per-account observations;
6. change WebView selection messages/mapping from observation selection to logical-event selection;
7. wire that selection to the existing comparison/detail state;
8. leave Report implementation for its later slice, but make its future input the same shared collection.

Do not refactor unrelated Forward/Move/Segment-B paths as part of this migration.

## Required automated evidence

A1.3 implementation should add or preserve focused tests for these seams:

### Core/domain tests

- Agent-10-owned supported non-recurring correlation cases produce the expected number of `LogicalEventGroup` results;
- duplicate/conflict/uncorrelated/suspicious-metadata/recurrence-unresolved cases remain explicit;
- no architecture-side test should encode a new identity heuristic that Agent 10 did not approve.

### App projection tests

- one supported correlated Core group -> one `LogicalEventReadModel`;
- N-account members remain N members of that one read model;
- explicit unresolved Core states/explanations survive projection unchanged;
- projection strips locator/native correlation/managed technical IDs from the shared normal presentation model;
- deterministic time/display mapping follows the accepted Agent 10 rule and never depends on scan/member order when that would imply authority;
- refresh replaces the presentation-ID map atomically and stale/unknown IDs do not select anything;
- WebView selection resolves to logical comparison/detail selection, not an arbitrary member;
- Calendar render serialization excludes `EntryId`, StoreID, GlobalAppointmentID, SyncGroupId, source-account metadata used only for correlation, and COM objects.

### Shared-model seam tests

- Calendar mapping is derived from the shared logical-event read model;
- existing/future Report adapter can consume the same read-model object/contract without rerunning correlation;
- no frontend test contains subject/time/ID grouping logic; synthetic frontend tests verify only render/update/navigation/selection behavior for already-projected logical items.

### Regression gates

Preserve the repository's existing Core/OutlookHost/.NET/frontend build, test, audit, publish, and packaged-asset verification established by A1.1/A1.2.

## Manual acceptance implications

After A1.2 is accepted/merged and A1.3 is implemented, the real-profile acceptance gate must verify at minimum:

- ordinary supported copies of one non-recurring logical event across accounts appear as one calendar item;
- selecting that item opens/updates one logical comparison/detail view containing the corresponding account observations;
- unresolved/duplicate/conflict/recurrence cases remain visibly truthful and are not guessed into a healthy correlated event;
- refresh does not leave stale selectable browser IDs;
- no raw Outlook technical IDs appear in the normal month-calendar surface or browser payload;
- existing A1.2 navigation, empty/degraded/error handling, accessibility, and read-only behavior remain intact.

## Blockers / revalidation gates

No unresolved **architecture** rule remains for Agent 40 once the domain contract is supplied by Agent 10.

Two external gates still apply before Agent 00 may issue A1.3 `SLICE READY`:

1. **Agent 10 domain readiness:** this note intentionally does not define supported identity/correlation rules or the authoritative placement rule when members disagree on time. Agent 10's accepted note must supply any domain rule needed there.
2. **A1.2 acceptance/merge revalidation:** this review is against the accepted A1.2 architecture/reconciliation contract because the A1.2 implementation had not advanced from current `main` at the final PR refresh. Agent 00 must revalidate the actual merged A1.2 bridge/read-model seam before implementation handoff for A1.3.

These are slice sequencing/domain gates, not unresolved architecture ownership.

**ARCH READY**
