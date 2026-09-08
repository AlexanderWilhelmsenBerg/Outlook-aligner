# Outlook Aligner — Production UI Contract

Status: **Required for v1. Implementation is underway in Phase 3; the current WinUI/read-alignment increment is not yet the complete production UI.**

This document defines the production interaction contract. The CLI tools used in Phases 1–2 are diagnostic/acceptance harnesses. Phase 3 moves normal testing into the WinUI application shell while preserving the OutlookHost COM boundary.

## 1. Product surface

Outlook Aligner is a native Windows desktop application built with WinUI 3. The production experience must not require PowerShell or command-line arguments for normal operation.

The application shell uses a `NavigationView` with four primary destinations:

1. **Calendar** — default/home page
2. **Alignment** — discrepancies and corrective actions
3. **Settings** — user preferences and scan configuration
4. **Diagnostics** — Outlook health, capabilities, operation history, and technical details

The current Phase 3 branch already implements this shell and basic functional versions of all four destinations. Production completeness still requires richer calendar visualization, persistence, final IPC, write confirmation/history, accessibility hardening, and the later Copy/Move write phases.

## 2. Calendar page

### Purpose

Give the user one understandable view of the configured calendars and make it obvious where events are aligned or differ.

### Current Phase 3 subset

The current increment:

- reads the bounded Outlook calendar window from the GUI;
- lists events across discovered accounts;
- derives source account and technical identity internally;
- offers other discovered accounts as Forward test targets;
- shows selected-event details;
- exposes native Forward capability check and safe **Prepare Forward (discard)**;
- keeps raw Outlook IDs in Diagnostics.

The current increment does not send a meeting from the GUI.

### Production layout

Top command area:

- Refresh
- scan horizon selector (default 90 days)
- date navigation / Today
- view selector (day / week / month where supported)
- account legend

Main area:

- FullCalendar hosted in WebView2;
- events visually associated with their source account;
- alignment state represented separately from account identity so color is not the only status signal;
- selecting an event opens a comparison/detail pane rather than requiring navigation away from the calendar.

FullCalendar/WebView2 integration should not be expanded until the repository can generate and commit a real Node 24 lockfile and use `npm ci`; never fabricate a lockfile.

Comparison/detail pane:

- logical event status;
- source/account membership;
- Start/End comparison;
- authority/origin and confidence;
- relevant detail differences;
- supported actions for the selected event;
- concise explanation for disabled/unsupported actions.

### Forward meeting action

Phase 2 proved a genuine Classic Outlook native Forward path for an accepted Calendar meeting and merged that mechanism in PR #5. The production UI presents this as one normal user action, **Forward meeting**, not as a choice between diagnostic COM mechanisms.

When the user selects Forward meeting:

1. the UI shows the intended source event/account and target account;
2. OutlookHost revalidates the selected Calendar item's identity and meeting state;
3. OutlookHost checks that Classic Outlook's built-in `Forward` command is valid, visible and enabled for that exact event;
4. only then may OutlookHost invoke the native Forward command and capture the native `MeetingItem` supplied by `AppointmentItem.Forward`;
5. no unexpected pre-existing recipients are allowed;
6. only the user-selected target is added/resolved;
7. `SendUsingAccount` is pinned to the selected source account;
8. send occurs only after explicit user intent/confirmation for that action.

If the event cannot be natively forwarded, **Forward meeting is disabled** and the detail pane explains why in user-facing language. The UI must never silently substitute `ForwardAsVcal()`/ICS and pretend it is equivalent native forwarding.

The retained-request recovery route investigated in Phase 2 is an implementation detail/diagnostic path and must not appear as a separate product action.

The current Phase 3 increment stops one step earlier: it can check capability and prepare/discard the native forwarded item, but it intentionally exposes no GUI Send action.

## 3. Alignment page

### Purpose

Provide a task-oriented queue of items that need attention.

### Current Phase 3 subset

The current increment already provides:

- preliminary non-recurring logical grouping;
- `Aligned`, `Missing`, `Moved`, `DetailsDifferent`, `Duplicate`, `Conflict`, recurrence-unresolved and uncorrelated states;
- read-only detection/classification of Outlook Aligner managed-copy metadata;
- fail-closed `Conflict` handling when managed metadata is incomplete, unsupported, unreadable, or internally inconsistent;
- `KnownOrigin` authority inference only from validated consistent managed-copy provenance;
- session-only manual authority selection where origin is unknown;
- a read-only Move Selected preview with `MOVE`, `SKIP`, and `BLOCKED` lines;
- no Move execution command.

Recurring identity remains deliberately unresolved.

### Production page

The page includes:

- status filters (`Missing`, `Moved`, `DetailsDifferent`, `Duplicate`, `Conflict`, `Ignored`);
- account/source filters;
- date range;
- authority/confidence filter where useful;
- sortable list/table of logical events;
- selection and detail pane;
- safe single-event actions;
- later batch selection/preview for Move All.

The page must clearly distinguish:

- original/authoritative events;
- Aligner-managed copies;
- true forwarded meetings;
- unmanaged/pre-existing possible matches;
- unresolved/suspicious metadata that must not be treated as trusted ownership evidence.

## 4. Settings page

At minimum:

- future scan horizon, default 90 days;
- account display names/order where appropriate;
- account visual identifiers;
- privacy preference for normal diagnostics/details;
- theme preference if the app does not simply follow Windows;
- confirmation preferences only where they do not weaken mandatory safety gates.

The current Phase 3 Settings page already owns the in-session scan horizon and displays discovered accounts/calendar availability. Persistence is not yet implemented.

Settings must not expose or request Outlook passwords, Graph credentials, Azure app registration, or tenant consent.

## 5. Diagnostics page

Diagnostics can expose technical information hidden from normal UI:

- Classic Outlook availability;
- profile/account/store discovery;
- OutlookHost/IPC status and protocol version;
- last refresh duration/results/errors;
- per-event native Forward capability where useful;
- `StoreID`, `EntryID`, `GlobalAppointmentID` for troubleshooting;
- managed-copy classification/property values;
- HRESULT/type context;
- operation history;
- app/package/version information.

The current Phase 3 Diagnostics page already shows selected-event IDs, managed-copy metadata/classification, scan summaries, warnings, and host output.

Meeting body, attendee lists and online-meeting URLs must not be dumped into diagnostics by default.

## 6. Write interaction rules

- Every write action identifies source and target before execution.
- Unsupported actions are disabled with an explanation rather than failing late where capability can be known in advance.
- Single-event destructive/significant actions require deliberate user invocation.
- Bulk writes always show an immutable preview first.
- The preview states exactly which events will be changed/skipped and why.
- Once confirmed, execution uses the previewed plan rather than silently recalculating a materially different plan.
- Partial failures are shown per item; already successful operations remain visible in history.
- The authoritative original is never mutated by Move Selected/Move All.
- Invalid/unreadable managed-copy metadata blocks managed-copy write eligibility.
- V1 has no deletion synchronization.

The current Phase 3 Move preview is intentionally preview-only and therefore exercises safety planning without introducing Outlook Move writes.

## 7. Loading, empty, error, and degraded states

The production UI must deliberately handle:

- Outlook not running/available;
- Classic Outlook not installed or profile unavailable;
- one account/store unavailable;
- partial calendar read failure;
- refresh in progress;
- no events in range;
- OutlookHost crash/restart;
- native Forward unavailable for a particular event;
- stale event locator that requires refresh;
- managed metadata that is unreadable/incomplete/unsupported;
- write action that fails after preview.

A failure in one account or event must not unnecessarily blank the entire application.

The current increment includes a user-facing warning for partial scans and directs technical detail to Diagnostics; later focused PRs should continue hardening degraded-state presentation.

## 8. Accessibility and Windows behavior

- Keyboard navigable.
- Visible focus states.
- Accessible names/automation properties for controls and status indicators.
- Account/status meaning must not rely on color alone.
- High-DPI and display scaling supported.
- Light/dark theme supported/follows Windows by default.
- Resizable desktop window with sensible minimum size.
- Long subjects/account names truncate gracefully and remain discoverable through accessible/tool-tip text.

These remain production acceptance requirements even where the current development shell has not yet been manually verified for every item.

## 9. UI/OutlookHost boundary

The UI never receives COM objects.

`OutlookAligner.App` exchanges versioned plain DTOs/results with `OutlookAligner.OutlookHost`. OutlookHost owns all COM interaction, including capability checks and writes.

The current Phase 3 development transport launches the sibling OutlookHost process and uses redirected standard streams for versioned JSON/text. This is a temporary implementation aid, not the final production IPC design.

Before production write workflows are complete, the app must move to the planned versioned local IPC/named-pipe model. View models/Core DTOs must not depend on the temporary transport.

## 10. Production UI acceptance gate

The production GUI is not complete until:

- a packaged WinUI 3 application launches normally;
- Calendar, Alignment, Settings and Diagnostics are functional;
- production OutlookHost IPC is wired with versioned DTOs;
- account discovery and refresh work from the GUI;
- FullCalendar renders the read model;
- selecting an event exposes comparison/status/authority/action state;
- Forward meeting capability and confirmation/send behavior are represented correctly;
- loading/empty/error/partial-failure states exist;
- normal usage needs no CLI arguments;
- persistence/history required by write workflows exists;
- keyboard/high-DPI/theme/accessibility basics are verified.

These acceptance items may be delivered through multiple focused PRs. A single PR is not required to complete an entire roadmap phase.
