# Outlook Aligner — Production UI Contract

Status: **Segment A read-only overview is the active product target. Reconciliation/write actions are deferred to Segment B.**

This document defines the user-facing UI contract. The current WinUI shell from PR #6 is the foundation; the next work turns it into a complete read-only cross-account calendar overview before Forward/Copy/Move workflows are productized.

## 1. Product surface

Outlook Aligner is a native WinUI 3 Windows desktop application for Classic Outlook.

Primary destinations:

1. **Calendar** — default/home page and month overview
2. **Report** — read-only discrepancy/reconciliation report
3. **Settings** — horizon, account identity/order, display preferences
4. **Diagnostics** — Outlook/host health and technical evidence

The existing Alignment destination may evolve into or be renamed to Report as Segment A Phase 4 is implemented. The underlying logical-event model must be shared rather than duplicated.

Normal use must not require PowerShell, command-line arguments, SMTP entry, StoreID, EntryID, or GlobalAppointmentID copying.

## 2. Calendar — primary month view

### Purpose

Give the user a conventional calendar view while simultaneously showing cross-account health.

### Layout

Top command area:

- Refresh
- Today
- previous/next period navigation
- current month title
- scan horizon where appropriate
- view mode selector
- status filters
- account filters/legend

Main area:

- month calendar is the default view;
- one visual event represents one logical event, even if copies exist in several accounts;
- selecting an event opens a read-only detail/comparison pane;
- calendar remains usable when the number of configured accounts grows beyond three.

FullCalendar/WebView2 remains the intended rendering stack. Before expanding it, generate a real Node 24 lockfile and move CI to `npm ci`; never fabricate a lockfile.

## 3. Event visual language

Account presence and logical alignment are different concepts and must be rendered independently.

### Account-presence dots — top-left

Each event shows small account markers in its **top-left corner**.

- One marker for each active account in which the logical event is present.
- Each account has a stable visual identity from Settings/account legend.
- The model must scale beyond A/B/C rather than assuming exactly three slots.
- Accessible text/tooltips identify the represented accounts; color is not the sole signal.

Example interpretation:

- dots A+B+C: observed in all three;
- dots A+C: missing from B;
- dot B only: only Account B currently contains the logical event.

### Alignment color — event body

For N active accounts:

- **Green:** event is present in all N and observed times agree with the authoritative occurrence.
- **Yellow:** event exists in at least two accounts but fewer than N, or any observed copy is moved/time-shifted from authority.
- **Red:** event exists in exactly one active account.

A moved event stays yellow even if all accounts contain it, because presence alone does not mean alignment.

Duplicate, Conflict, unresolved identity, or other exceptional states must also have an icon/text/accessibility treatment so they are not misleadingly flattened into red/yellow/green.

### Recurrence symbol — lower-right

Recurring events show a recurrence symbol in the **lower-right corner**.

The recurrence symbol is visually separate from account markers and status color. Modified recurring exceptions remain marked as recurring and must stay associated with the correct logical occurrence.

## 4. Event detail/comparison pane

Selecting an event exposes:

- subject;
- logical status;
- authoritative/origin account and authority reason;
- account membership;
- authoritative Start/End;
- per-account Start/End where different;
- missing accounts;
- recurrence state;
- duplicate/conflict/unresolved explanation where applicable;
- relevant normal details such as location when allowed by privacy settings.

Raw technical identifiers belong in Diagnostics, not this pane.

During Segment A the pane is **read-only**. Existing Forward capability/prepare diagnostics from the foundation branch are not part of Segment A acceptance and should not drive the roadmap.

## 5. Event-specific authority

Authority is chosen per logical event, not per application/account.

Examples:

- Event 1 may have Account A as authority.
- Event 2 may have Account B as authority.
- Event 3 may have Account C as authority.

Authority reasons:

- `KnownOrigin`
- `UserSelected`
- `Inferred`
- `Unknown`

Reliable provenance/native evidence may establish origin. **Scan/enumeration order is never sufficient evidence.** When origin cannot be established safely, show Unknown or use an explicit user selection.

`LastModificationTime` is never a latest-wins authority rule.

In Segment A, authority only explains comparison/moved state; it does not authorize a write.

## 6. Recurring read behavior

Supported recurring cases must be represented as logical occurrences, not merely as a shared series ID.

The read model distinguishes:

- series master;
- ordinary occurrence;
- modified/moved exception;
- deleted occurrence diagnosis where available.

A moved occurrence stays logically tied to its original occurrence. Recurrence expansion remains bounded to the configured scan horizon.

Segment A never creates, moves, deletes, or repairs a recurring item.

## 7. Report page — read-only reconciliation report

### Purpose

Answer “what needs attention?” without changing Outlook.

Required report categories:

- Missing
- Moved
- Single-account
- Duplicate
- Conflict / unresolved

Rows show at least:

- subject;
- authoritative/origin account when known;
- account-presence markers;
- authoritative time;
- differing per-account times;
- missing account names;
- recurrence indicator;
- concise reason/status.

Selecting a row opens/selects the same logical-event detail model used by Calendar.

The report may be called “Reconciliation Report,” but **Segment A performs no reconciliation writes**.

## 8. View modes and filters

Required view modes:

- Calendar only
- Report only
- Calendar + Report combined/split where practical

Required logical-state filters:

- All
- Aligned
- Missing
- Moved
- Duplicate
- Conflict / unresolved

Account include/exclude filters are also required.

Filters must operate on the shared logical-event result set. Choosing “Moved” should produce consistent results in Calendar and Report without re-reading Outlook merely to change presentation.

## 9. Settings

At minimum:

- future scan horizon, default 90 days;
- discovered account display names/order;
- stable account visual identifiers used by top-left markers;
- account include/exclude defaults where useful;
- privacy/display preferences;
- theme behavior.

Segment A may persist view/filter preferences and explicit authority selections locally. It must not request Outlook passwords, Graph credentials, Azure registration, or tenant consent.

## 10. Diagnostics

Diagnostics may expose technical information hidden from normal UI:

- Classic Outlook availability;
- profile/account/store discovery;
- OutlookHost protocol/health;
- refresh duration and partial failures;
- StoreID / EntryID / GlobalAppointmentID;
- recurrence identity evidence;
- managed-copy metadata/classification;
- structured event log;
- app/build information.

Meeting bodies, attendee lists, attachments, and online-meeting URLs must not be dumped into diagnostics by default.

## 11. Loading/error/degraded states

The read-only product deliberately handles:

- Outlook unavailable;
- profile unavailable;
- one account/store unavailable;
- partial item/calendar failures;
- refresh in progress;
- no events in range;
- OutlookHost crash/restart;
- stale locators;
- unresolved recurrence identity;
- conflicting/unreadable managed metadata.

One failed account/event must not blank healthy results from the others.

## 12. Accessibility and Windows behavior

- Keyboard navigable.
- Visible focus states.
- Account/status meaning not dependent on color alone.
- Accessible names for account markers, recurrence state and status.
- High-DPI/display scaling.
- Light/dark theme support/follow Windows by default.
- Resizable window with sensible minimum dimensions.
- Long subjects/account names truncate gracefully with accessible full text.

## 13. UI/OutlookHost boundary

The UI never receives COM objects.

`OutlookAligner.App` exchanges versioned plain DTOs/results with `OutlookAligner.OutlookHost`; OutlookHost owns all Classic Outlook COM access.

The current child-process/stdout JSON transport is transitional. Segment A must keep the logical read model transport-independent. Production write workflows in Segment B should move to the planned versioned local IPC/named-pipe model before they are considered complete.

## 14. Segment A UI acceptance gate

Segment A is complete when:

- app launches normally;
- accounts/calendars discover automatically;
- month calendar is the primary usable view;
- each logical event renders once;
- account dots appear top-left and scale beyond three accounts;
- green/yellow/red behavior matches shared logical state;
- recurrence symbol appears lower-right;
- ordinary recurring occurrences and moved exceptions correlate correctly for supported cases;
- event-specific authority is visible/explainable;
- Report exposes Missing/Moved/Duplicate/Conflict cases;
- Calendar only / Report only / combined modes work;
- status and account filters work consistently;
- partial failures remain visible;
- no normal workflow requires raw Outlook identifiers or PowerShell;
- no Segment A interaction persistently mutates Outlook;
- keyboard/color/accessibility basics are verified on the real machine.

Only after this gate is accepted does Segment B reconciliation become the active UI target.
