# Outlook Aligner — Implementation Plan

Status: **Technical foundation established. The product roadmap now prioritizes a complete read-only overview experience before any reconciliation/write workflows. Segment A is the active roadmap.**

Last reviewed: 2026-09-11

## 1. Product goal

Build a Windows desktop application that uses the **Classic Outlook COM/Object Model** as its Outlook integration boundary, discovers the calendar-capable accounts already configured in one Outlook profile, reads their calendars, correlates the same logical event across accounts, and gives the user a clear visual overview of what is aligned, missing, moved, or otherwise inconsistent.

The product is intentionally split into two major stages:

- **Segment A — Read-only overview:** make the application genuinely useful without changing Outlook data.
- **Segment B — Reconciliation:** add deliberate Forward/Copy/Move actions only after the read model and visual comparison experience are trustworthy.

Microsoft Graph, MSAL, Azure app registration, tenant consent, and separate Outlook credentials are out of scope.

### PR sizing rule

Roadmap phases are product slices, not mandatory PR sizes. Use the smallest coherent PR that can be independently reviewed and tested. Never merge a PR automatically; merge only after explicit owner approval.

## 2. Technical foundation already completed

The following work predates the Segment A/B product roadmap and remains valid foundation rather than being discarded.

### Foundation 0 — repository/toolchain bootstrap ✅

- .NET 10 / C# 14 solution and CI.
- WinUI 3 / Windows App SDK.
- Outlook interop packaging.
- TypeScript/Vite calendar frontend foundation.
- tests, coverage, vulnerability audit, artifacts.

### Foundation 1 — Classic Outlook read boundary ✅

Merged in PR #4.

Proved:

- account/store/default Calendar discovery;
- bounded calendar scanning;
- recurrence-safe enumeration inside a time horizon;
- StoreID + EntryID locator handling;
- GlobalAppointmentID reads;
- partial-failure behavior;
- packaged real-machine execution.

### Foundation 2 — native Forward technical spike ✅

Merged in PR #5.

A genuine Outlook native Forward route was proven end-to-end on a real accepted meeting. That code remains useful for Segment B, but **Forward is no longer a prerequisite for completing Segment A** and broad Forward testing is deliberately deferred until the read-only product is complete.

### Foundation 3 — WinUI/read-alignment shell ✅ pending PR #6 merge

PR #6 provides the first usable desktop shell:

- WinUI app and NavigationView;
- Calendar, Alignment, Settings and Diagnostics destinations;
- GUI-driven Outlook account/calendar discovery;
- structured persistent diagnostics;
- preliminary non-recurring correlation;
- read-only managed-copy metadata inspection;
- authority selection/preview scaffolding;
- safe Forward capability/prepare diagnostics with no GUI send;
- self-contained CI test bundle and local updater.

The owner has validated that this build launches and satisfies the initial read use case. Forward/reconciliation behavior is **not** a merge gate for this foundation PR.

## 3. Segment A — Read-only overview

Segment A is complete only when Outlook Aligner is useful as a read-only cross-account calendar overview application. No Forward, Copy, Move, delete, or other persistent Outlook mutation is required to close Segment A.

### Segment A Phase 1 — Month calendar view

Create the main production calendar surface.

Requirements:

- Calendar defaults to a conventional **month view**.
- Show one logical event once, rather than rendering a separate duplicate row/card for every account copy.
- Date navigation and **Today** control.
- Keep the bounded Outlook refresh model and current scan horizon.
- Selecting a calendar event opens its read-only comparison/detail information.
- Calendar must remain usable with more than three configured accounts even though the initial real profile has three.
- FullCalendar/WebView2 may be used as planned; if used, generate a real Node 24 lockfile and switch CI to `npm ci`. Never fabricate a lockfile.

Acceptance:

- A user can open Outlook Aligner and understand the month schedule without using raw IDs, PowerShell, or Diagnostics.
- A logical event occupies one visual calendar item even when copies exist in multiple accounts.

### Segment A Phase 2 — Account markers and alignment color coding

Add visual cross-account state directly to every calendar event.

#### Account-presence markers

Each logical event shows small account markers in the **top-left corner** of the event.

- One dot/marker per account in which that logical event is present.
- Each account has a stable visual identity in the account legend/settings.
- Marker design must scale beyond A/B/C; do not hard-code exactly three visual slots.
- Color alone must not be the only accessible representation; marker tooltip/accessible text names the account(s).

#### Event status color

The event body/background represents the logical alignment state across the active accounts:

- **Green:** present in all active accounts and times agree with the authoritative occurrence.
- **Yellow:** present in at least two accounts but not all active accounts, **or** at least one observed copy is moved/time-shifted from the authoritative occurrence.
- **Red:** present in only one active account.

For N accounts this generalizes as:

- 1 of N = red;
- 2 through N-1 of N = yellow;
- N of N = green only when the relevant times are aligned;
- any moved/time disagreement overrides green to yellow.

Conflict/duplicate/unresolved states must have an additional symbol/text treatment so they are not misleadingly reduced to green/yellow/red.

Acceptance:

- The user can scan the month view and immediately see which events are healthy, incomplete, or isolated.
- Account membership and alignment state are represented independently: account dots answer **where is it?**, event color answers **how healthy is the cross-account state?**

### Segment A Phase 3 — Recurring meetings in the read model

Make recurring events first-class in the read-only experience.

Requirements:

- Correctly represent series masters, ordinary occurrences, and modified/moved exceptions inside the bounded scan horizon.
- Preserve logical occurrence identity when a recurring occurrence moves.
- Do not correlate unrelated occurrences merely because they share a series GlobalAppointmentID.
- Calendar events that belong to a recurring series show a **recurrence symbol in the lower-right corner**.
- The symbol is informational/read-only in Segment A.
- Deleted occurrence diagnosis may be displayed, but Segment A never deletes or recreates anything.

Acceptance:

- Recurring meetings no longer fall into a generic unresolved bucket during normal supported cases.
- A moved exception stays associated with the correct logical occurrence.
- The calendar clearly marks recurring items without cluttering account/status markers.

### Segment A Phase 4 — Reconciliation report view (read-only)

Add a new report focused on discrepancies. Despite the name, this phase **does not perform reconciliation**; it reports what would need attention.

The report must include at least:

- **Moved:** one or more accounts have a different Start/End from the event authority.
- **Missing:** event is absent from one or more active accounts.
- **Single-account:** useful red-state subset of Missing.
- **Duplicate:** multiple candidate copies exist in the same account.
- **Conflict / unresolved:** identity or authority cannot be determined safely.

Each report row shows:

- subject;
- authoritative/origin account when known;
- account-presence markers;
- authoritative time;
- per-account observed time where different;
- missing account names;
- recurrence indicator;
- clear reason/status.

Selecting a report row selects/opens the same logical event detail model used by Calendar. Calendar and Report must not invent separate correlation logic.

Acceptance:

- The user can answer “what needs attention?” without manually comparing calendars.
- No report action mutates Outlook in Segment A.

### Segment A Phase 5 — Filters and view modes

Finish the read-only overview workflow by allowing the user to focus the Calendar and Report surfaces.

Required view modes:

- **Calendar only**
- **Report only**
- **Calendar + Report** combined/split view where practical

Required status filters include:

- All
- Aligned
- Missing
- Moved
- Duplicate
- Conflict / unresolved

Also support account filters so the user can include/exclude configured calendars from the current comparison where useful.

Filter behavior must be shared between Calendar and Report: if the user asks to show only Moved, both surfaces should reflect that same logical result set unless a view explicitly documents otherwise.

Acceptance:

- User can switch between broad overview and discrepancy-focused inspection without rescanning Outlook.
- Filter state changes presentation only; it does not change Outlook data or silently redefine stored identity.

## 4. Segment A authority/origin model

Authority is **event-specific**, never globally tied to Account A/B/C.

Examples:

- If logical Event 1 originates/arrives first through Account A and that origin can be established reliably, Account A is authoritative for Event 1.
- If Event 2 originates/arrives through Account B, Account B is authoritative for Event 2.
- Account C may be authoritative for a different event.

There is no permanent “master calendar” unless the user explicitly introduces such a preference later.

Authority reasons remain explicit:

- `KnownOrigin`
- `UserSelected`
- `Inferred`
- `Unknown`

Important safety rule: **scan order is not origin evidence**. Outlook Aligner must not call whichever account happened to be enumerated first authoritative. Reliable native/provenance evidence may establish origin; otherwise the event remains Unknown or the user selects authority.

`LastModificationTime` must never become a latest-wins authority rule.

In Segment A, authority is used only to explain comparison state and determine what counts as “moved”; it does not authorize a write.

## 5. Logical event/read-state model

Primary states remain:

- `Aligned`
- `Missing`
- `Moved`
- `DetailsDifferent`
- `Duplicate`
- `Conflict`
- `Ignored`
- `DeletedOrMissing` for diagnosis only

Temporary implementation safety states such as `RecurrenceIdentityUnresolved` and `Uncorrelated` may remain while Segment A phases are being built, but supported ordinary and recurring cases should progressively leave those buckets.

### Identity rules

- `StoreID + EntryID` is a locator cache, not stable logical identity.
- `GlobalAppointmentID` is the primary native correlation candidate for ordinary non-recurring Outlook meetings.
- Managed copies may use validated Outlook Aligner provenance metadata.
- Recurrence requires occurrence/exception identity; series GlobalAppointmentID alone is insufficient.
- Suspicious or unreadable managed metadata fails closed into Conflict/unresolved behavior.

## 6. Segment A exit gate

Do not begin Segment B implementation merely because underlying Forward code already exists. Segment A closes when the owner can use the application as a reliable read-only daily overview.

Required exit criteria:

- app launches normally from the test/release package;
- configured Classic Outlook accounts are discovered automatically;
- month calendar is the primary usable view;
- logical events render once with account-presence markers;
- green/yellow/red status behaves correctly for N active accounts;
- event-specific authority is visible and explainable;
- recurring events and moved exceptions are represented correctly for supported cases;
- recurrence symbol appears in the lower-right of recurring events;
- read-only reconciliation report shows Missing/Moved/Duplicate/Conflict cases;
- Calendar only / Report only / combined view modes work;
- status/account filters work consistently;
- partial account/item failures remain visible;
- Diagnostics remains available without raw technical IDs leaking into normal UI;
- normal use requires no PowerShell/manual Outlook IDs;
- no Segment A action persistently changes Outlook calendar data;
- user has manually accepted the read-only workflow on the real Classic Outlook profile.

## 7. Segment B — Reconciliation

Only after Segment A is accepted, add controlled mutations. Segment B reuses the exact same logical-event, authority, recurrence, report, and filter model rather than creating a second reconciliation engine.

Proposed order:

### Segment B Phase 1 — action planning and immutable previews

- Turn read-only discrepancy rows into explicit candidate action plans.
- Source, authority, target, old state and intended new state visible before execution.
- Unsupported/conflict/unknown cases remain blocked.

### Segment B Phase 2 — Forward selected

- Productize the already-proven native Outlook Forward mechanism.
- Revalidate exact event capability before send.
- Explicit target and confirmation.
- No vCalendar/ICS fallback masquerading as native Forward.

### Segment B Phase 3 — Copy Full / Copy Busy

- Create Aligner-managed copies with durable provenance.
- Full and privacy-preserving Busy modes remain semantically distinct.

### Segment B Phase 4 — Move Selected

- Align non-authoritative managed copies to authoritative Start/End.
- Never mutate the authoritative original as part of Move.

### Segment B Phase 5 — bulk reconciliation

- Preview-first multi-event actions.
- Immutable plan, explicit confirmation, per-item result/history.
- Conflicts and unsupported recurrence excluded safely.

### Segment B Phase 6 — recurring writes

- Only after recurring read identity is proven by Segment A Phase 3.
- Explicit occurrence-vs-series semantics.

Deletion synchronization remains out of scope for v1 unless deliberately added in a later roadmap decision.

## 8. Architecture

Chosen stack remains:

- C# 14 / .NET 10 LTS
- WinUI 3 / Windows App SDK
- WebView2 + FullCalendar Vanilla
- TypeScript + Vite
- SQLite via Microsoft.Data.Sqlite + Dapper
- CommunityToolkit.Mvvm
- Microsoft.Extensions hosting/config/logging + Serilog
- xUnit v3 / Microsoft Testing Platform
- BenchmarkDotNet
- GitHub Actions + Dependabot

Process boundary:

```text
OutlookAligner.App.exe
  WinUI / presentation / correlation read model / local persistence
        |
        | versioned local transport
        v
OutlookAligner.OutlookHost.exe
  interactive STA process
  Classic Outlook COM/Object Model
```

COM never enters UI/Core DTOs. OutlookHost remains an interactive user process, never a Windows Service. The current child-process/stdout transport is transitional; production write workflows should move to the planned versioned local IPC/named-pipe boundary without coupling the read model to transport details.

## 9. Local persistence

Planned tables remain:

- `Accounts`
- `SyncGroups`
- `EventMembers`
- `UserOverrides`
- `OperationHistory`
- `SchemaMigrations`

Segment A may introduce persistence needed for display preferences, filters, account identities/order, and user-selected authority. It must not require write-operation history until Segment B.

No Outlook passwords or tokens are stored.

## 10. Global safety rules

- Never merge a PR automatically.
- Segment A is read-only: do not introduce Outlook mutation merely to make testing easier.
- Never infer deletion permission from a missing event.
- Never use scan order or `LastModificationTime` as automatic authority.
- Invalid/unreadable managed metadata fails closed.
- Never leak body/attendees/online links into ordinary diagnostics.
- Keep account identity and alignment status visually distinct.
- Do not rely on color alone for accessibility.
- Segment B writes require explicit user intent, capability checks, preview where relevant, and operation history.
- Bulk writes require preview + confirmation.
- Never silently fall back from native Forward to vCalendar/ICS.

## 11. V1 direction

The immediate release milestone is **Segment A complete: useful read-only Outlook Aligner**.

A later reconciliation-capable v1 can add Segment B incrementally after the read-only product is trusted on the real profile.
