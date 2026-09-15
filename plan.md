# Outlook Aligner — Slice Implementation Plan

Status: **Segment A — read-only overview is the active product milestone. Segment B reconciliation is gated behind explicit Segment A acceptance.**

Last restructured: 2026-09-15

## 1. North star

Outlook Aligner must first become a dependable read-only desktop calendar overview for the accounts configured in one Classic Outlook profile. A user should be able to open it and immediately understand:

- what is on the calendar;
- which accounts contain each logical event;
- which events are fully aligned;
- which events are missing or moved;
- which account is authoritative when that can be established;
- which items are recurring;
- what discrepancies need attention.

Only after that read model and UI are trusted should the product begin changing Outlook state.

The application remains Classic-Outlook-COM-only. Graph/MSAL/Azure app registration are out of scope.

## 2. Baseline and execution rule

The App Director must refresh current `main` before starting the first slice and record the exact baseline HEAD. If the WinUI/read-alignment foundation PR is not yet merged, finish that merge decision first; do not build the new roadmap on an ambiguous base.

Each slice below is a product contract, not automatically one PR. Split a slice further when that produces a safer independently testable change. Never combine later slices merely to reduce PR count.

Every implementation PR follows:

`specialist readiness -> Agent 00 SLICE READY -> Agent 40 implementation -> exact-head CI -> Agent 50 independent acceptance -> owner manual gate -> owner merge`

No agent merges automatically.

## 3. Segment A — Read-only overview

Segment A ends with a genuinely useful read-only product. Normal Segment A workflows must not Forward, Copy, Move, delete, save, or otherwise persistently mutate Outlook calendar items.

### Phase A1 — Month calendar view

#### Slice A1.1 — Deterministic calendar frontend toolchain

**Goal:** make the existing TypeScript/FullCalendar frontend reproducible before expanding it.

**Owners:** Agent 30 architecture; Agent 40 implementation; Agent 50 review.

**Contract:**

- Refresh the pinned Node/npm baseline already documented by the repo.
- Generate a real dependency lockfile from the declared frontend dependencies; never fabricate one.
- CI uses deterministic installation (`npm ci`) once the lockfile is committed.
- Preserve existing typecheck/lint/format/build/audit gates.
- Do not change product behavior beyond what is required to make the frontend build deterministic.

**Acceptance:** clean checkout can install/build the frontend deterministically in CI; no unrelated dependency upgrade sweep.

#### Slice A1.2 — WebView2 month-calendar host

**Goal:** replace the list-first Calendar experience with a real month calendar surface while preserving the WinUI shell and OutlookHost boundary.

**Owners:** Agent 20 UX + Agent 30 architecture readiness; Agent 40 implementation.

**Contract:**

- Calendar opens in month view.
- Today and previous/next month navigation work.
- WinUI sends plain read-model DTOs into WebView2; the WebView never receives COM objects.
- FullCalendar is presentation only; correlation/status rules stay in Core/App read models.
- Loading, empty, host-error, and partial-result states remain visible outside/around the WebView.
- Existing Forward/Prepare diagnostic code is not expanded and is removed from the normal Calendar workflow for Segment A.

**Acceptance:** on the real profile, the month grid opens and renders the bounded calendar data without PowerShell or raw Outlook IDs.

#### Slice A1.3 — One logical event, one calendar item

**Goal:** the month view represents a logical event once rather than drawing one copy per account for supported non-recurring cases.

**Owners:** Agent 10 domain + Agent 30 architecture readiness; Agent 40 implementation.

**Contract:**

- Create/standardize a presentation DTO for one logical event.
- Supported non-recurring items correlate through the accepted identity rules.
- Duplicate, conflict, uncorrelated, suspicious-metadata, and recurrence-unresolved cases remain explicit rather than being guessed together.
- Selecting a calendar item opens/updates the existing read-only comparison/detail model.
- Calendar and later Report must consume the same logical-event model.

**Acceptance:** ordinary copies of the same non-recurring logical event across accounts produce one selectable calendar item; unresolved cases remain truthful.

### Phase A2 — Account markers and alignment color coding

#### Slice A2.1 — Event-specific authority/origin read contract

**Goal:** make authority semantics precise enough that `Moved` and health colors are truthful.

**Owners:** Agent 10 domain; Agent 20 presentation; Agent 50 review.

**Contract:**

- Authority is per logical event, never a globally preferred A/B/C account.
- Reliable native/provenance evidence may establish `KnownOrigin`.
- Safe deterministic inference must be explicitly defined and tested before using `Inferred`.
- Pre-existing ambiguous copies remain `Unknown` unless the user selects authority.
- User-selected authority may be retained as read-only app state/persistence; it does not itself authorize a future write.
- Scan order and `LastModificationTime` are forbidden authority heuristics.
- Detail UI explains authority reason.

**Acceptance:** Event 1 can truthfully be authoritative from A while Event 2 is authoritative from B; ambiguous data does not silently pick a master.

#### Slice A2.2 — Stable account visual identity and presence markers

**Goal:** show where each event exists independently from its alignment health.

**Owners:** Agent 20 UX + Agent 30 data contract; Agent 40 implementation.

**Contract:**

- Each active account has a stable display label and visual marker identity for the session, with persistence introduced only if needed.
- Each logical event renders one small marker/dot per account where it is observed.
- Markers occupy the top-left event region.
- Marker layout scales to N accounts; do not hard-code exactly three positions.
- Accessible text/tooltips name present/missing accounts so color is not required to understand membership.
- Account legend uses the same identities as event markers.

**Acceptance:** a user can tell which account(s) contain an event without opening Diagnostics.

#### Slice A2.3 — Pure health classifier and event status presentation

**Goal:** make green/yellow/red derive from a tested domain classifier, not ad-hoc UI conditions.

**Owners:** Agent 10 status truth + Agent 20 visual contract; Agent 40 implementation.

**Classifier contract for N active accounts:**

- `Red`: the logical event is present in exactly one active account.
- `Yellow`: present in at least two but fewer than all active accounts, or a relevant observed copy has a Start/End disagreement from the authority.
- `Green`: present in all active accounts and relevant times agree with the authoritative occurrence.
- `Moved` overrides an otherwise-green presence count to yellow.
- Duplicate/conflict/unresolved/unknown-authority cases that cannot be honestly reduced to the three health states receive an explicit secondary icon/badge/text state; do not paint them deceptively green.

**UI contract:**

- event background/body carries health color;
- account markers remain separate from health;
- status accessible text and non-color indicator communicates `Aligned`, `Missing`, `Moved`, `Duplicate`, `Conflict`, or unresolved state.

**Acceptance:** classifier unit tests cover 1..N membership, moved copies, unknown authority, duplicate/conflict, and three-account examples; UI matches classifier output exactly.

### Phase A3 — Recurring meetings

#### Slice A3.1 — Recurrence identity contract and fixtures

**Goal:** define recurring identity before changing correlation code.

**Owners:** Agent 10 domain with Agent 30 architecture review.

**Contract:**

- Define the plain DTO fields required to distinguish series master, normal occurrence, modified/moved exception, and deleted occurrence evidence.
- Define a stable logical occurrence key that does not rely on expanded-item EntryID stability or series GlobalAppointmentID alone.
- A moved exception remains the same logical occurrence as its original occurrence identity.
- Bound all recurrence expansion to the requested scan horizon.
- Define fail-closed behavior for recurrence information Outlook cannot resolve reliably.
- Add representative pure fixtures/tests before UI integration.

**Acceptance:** Agent 10 returns `DOMAIN READY`; tests demonstrate two occurrences in the same series cannot collapse into one and a moved exception remains associated with the correct occurrence.

#### Slice A3.2 — OutlookHost recurrence observation support

**Goal:** emit the recurrence evidence required by A3.1 without leaking COM or mutating Outlook.

**Owners:** Agent 10 + Agent 30 readiness; Agent 40 implementation.

**Contract:**

- Read only the Outlook recurrence properties required by the accepted domain contract.
- Reacquire/release recurrence COM objects safely.
- Preserve bounded enumeration and partial-item failure behavior.
- Version DTO/protocol deliberately if the host/app contract changes.
- No recurrence writes, saves, forwarding, or repair operations.

**Acceptance:** hosted tests cover DTO parsing/classification where possible; manual scan verifies recurring observations arrive without altering appointments.

#### Slice A3.3 — Recurring correlation and calendar presentation

**Goal:** move supported recurring cases out of the generic unresolved bucket and render them truthfully.

**Owners:** Agents 10 + 20 + 30 readiness; Agent 40 implementation.

**Contract:**

- Correlate the same logical occurrence across accounts.
- Preserve moved-exception identity.
- Apply account markers and health classifier per occurrence.
- Render recurrence symbol at the lower-right of recurring calendar items.
- Keep unsupported recurrence cases explicit as unresolved/conflict.

**Acceptance:** ordinary recurring occurrences, moved exception, and at least one unsupported/degraded recurrence case are covered; manual month view shows correct recurrence indicator and grouping.

### Phase A4 — Read-only reconciliation report

#### Slice A4.1 — Shared discrepancy query model

**Goal:** derive report rows from the same logical-event model as Calendar.

**Owners:** Agent 10 truth + Agent 30 architecture; Agent 40 implementation.

**Contract:**

- Report candidates include at least `Missing`, `Moved`, `Duplicate`, `Conflict`, and unresolved states.
- A report row contains subject, authority/reason when known, authoritative time, account presence, per-account differing time, missing accounts, recurrence state, and explanation.
- No independent report-only correlation engine.
- Pure query/model tests prove Calendar and Report classification agree.

**Acceptance:** given one logical-event set, Calendar status and report inclusion cannot contradict one another.

#### Slice A4.2 — Reconciliation report UI

**Goal:** answer “what needs attention?” without manual calendar comparison.

**Owners:** Agent 20 UX; Agent 40 implementation.

**Contract:**

- Provide a readable sortable/list-style discrepancy view.
- Missing accounts and moved times are visible without raw IDs.
- Account markers and status language match Calendar.
- Selecting a report row opens/selects the same logical-event detail model used by Calendar.
- No report button mutates Outlook.

**Acceptance:** user can identify moved and missing meetings from the report alone on the real profile.

### Phase A5 — Filters, view modes, persistence and read-only closeout

#### Slice A5.1 — Shared filters

**Goal:** one filter state drives Calendar and Report.

**Owners:** Agent 20 UX + Agent 30 state architecture.

**Required filters:**

- All;
- Aligned;
- Missing;
- Moved;
- Duplicate;
- Conflict / unresolved;
- account include/exclude where useful.

**Contract:** filtering changes presentation only. It must not rescan Outlook unnecessarily, rewrite identity, or change persisted Outlook data.

**Acceptance:** “Moved only” produces the same logical result set in Calendar and Report.

#### Slice A5.2 — Calendar / Report / combined view modes

**Goal:** support broad overview and discrepancy-focused work without losing selection/filter context.

**Owners:** Agent 20 UX; Agent 30 layout/state; Agent 40 implementation.

**Modes:**

- Calendar only;
- Report only;
- Calendar + Report split/combined view where practical.

**Acceptance:** switching modes preserves logical selection/filter state and remains usable at practical desktop window sizes/high DPI.

#### Slice A5.3 — Persist read-only user preferences

**Goal:** retain the read experience without creating write-workflow persistence early.

**Owners:** Agent 30 architecture with Agent 20 UX.

**Candidate persisted state:**

- scan horizon;
- account display identities/order if user-configurable;
- current view mode;
- filter selections if appropriate;
- user-selected authority overrides.

**Contract:** no Outlook passwords/tokens; no Segment B operation history required yet; schema/migrations are explicit if SQLite is introduced/expanded.

**Acceptance:** restart restores the agreed read-only preferences without changing Outlook state.

#### Slice A5.4 — Segment A hardening and acceptance gate

**Goal:** close Segment A as an independently useful product milestone.

**Owners:** Agent 50 independent acceptance; Agent 00 final reconciliation; owner manual acceptance.

**Required gates:**

- all hosted build/test/audit/publish checks green on exact head;
- month calendar is primary and usable;
- logical events render once for supported ordinary and recurring cases;
- account markers scale beyond three accounts;
- health classifier is tested and presentation is accessible without color;
- event-specific authority reason is visible and truthful;
- recurring supported cases and moved exceptions behave correctly;
- report identifies Missing/Moved/Duplicate/Conflict/unresolved;
- filters and view modes are consistent;
- partial failures/degraded states remain visible;
- Diagnostics remains available without leaking raw IDs into normal UI;
- updater/signing workflow remains functional;
- normal Segment A workflows perform no persistent Outlook calendar mutation;
- owner manually accepts the real-profile read-only workflow.

Only after this gate may Agent 00 mark Segment A complete and schedule Segment B implementation.

## 4. Segment B — Reconciliation (blocked until Segment A acceptance)

Segment B reuses the exact same logical-event, authority, recurrence, report, and filter model. It must not create a second reconciliation engine.

### Slice B1 — Action planner and immutable preview contract

Turn discrepancies into explicit candidate plans without executing them. Define source, authority, target, current state, intended state, unsupported reasons, stale-data revalidation, and immutable confirmation semantics.

### Slice B2 — Native Forward selected

Productize the already-proven Classic Outlook native Forward path with exact-event capability revalidation, explicit target, confirmation, failure reporting, and no ICS/vCalendar masquerade.

### Slice B3 — Copy Full

Create Outlook Aligner-managed full-detail copies with durable provenance, preview, explicit semantics, and operation history.

### Slice B4 — Copy Busy

Create privacy-preserving busy placeholders with separate semantics from Copy Full and no accidental detail leakage.

### Slice B5 — Move Selected

Align only eligible non-authoritative managed copies to authoritative Start/End. Never mutate the authoritative original as part of Move.

### Slice B6 — Bulk reconciliation

Build preview-first multi-event reconciliation with immutable plans, explicit confirmation, conflict/unsupported exclusions, and per-item result/history.

### Slice B7 — Recurring writes

Add only the recurrence mutations supported by the proven Segment A occurrence identity. Require explicit occurrence-vs-series semantics and block unsupported cases.

### Slice B8 — Integrated reconciliation hardening/release

Run the broad real-profile matrix for Forward/Copy/Move/recurrence, crash recovery, stale locators, partial COM failures, persistence/history, accessibility, packaging/signing/update, performance, and migration. Deletion synchronization remains out of scope unless explicitly added later.

## 5. Global quality rules

- Repository/current code outranks remembered discussion.
- Do not implement hidden product decisions in Agent 40.
- One logical read model feeds Calendar, Report, filters, and future action planning.
- Color communicates health but never stands alone.
- Account identity and health state remain separate visual channels.
- All N-account rules must work beyond the initial three-account profile unless a documented Outlook limitation prevents it.
- Segment A is read-only in normal use.
- Segment B writes require explicit intent, capability/safety checks, and operation history.
- Bulk writes require preview + confirmation.
- Never merge automatically.
