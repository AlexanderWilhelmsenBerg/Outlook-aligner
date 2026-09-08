# Outlook Aligner — Production UI Contract

Status: **Required for v1. Implementation begins in Phase 4.**

This document turns the UI from an architectural intention into a product requirement. The CLI tools used in Phases 1–3 are diagnostic harnesses only.

## 1. Product surface

Outlook Aligner is a native Windows desktop application built with WinUI 3. The production experience must not require PowerShell or command-line arguments for normal operation.

The application shell uses a `NavigationView` with four primary destinations:

1. **Calendar** — default/home page
2. **Alignment** — discrepancies and corrective actions
3. **Settings** — user preferences and scan configuration
4. **Diagnostics** — Outlook health, capabilities, operation history, and technical details

The title area should expose Outlook connection/profile state and refresh activity without turning technical status into the visual focus of the application.

## 2. Calendar page

### Purpose

Give the user one understandable view of the three calendars and make it obvious where events are aligned or differ.

### Layout

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

Comparison/detail pane:

- logical event status;
- source/account membership;
- Start/End comparison;
- authority/origin and confidence;
- relevant detail differences;
- supported actions for the selected event;
- concise explanation for disabled/unsupported actions.

### Forward meeting action

Phase 2 has proven a genuine Classic Outlook native Forward path for an accepted Calendar meeting. The production UI must present this as one normal user action, **Forward meeting**, not as a choice between diagnostic COM mechanisms.

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

## 3. Alignment page

### Purpose

Provide a task-oriented queue of items that need attention.

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
- unmanaged/pre-existing possible matches.

## 4. Settings page

At minimum:

- future scan horizon, default 90 days;
- account display names/order where appropriate;
- account visual identifiers;
- privacy preference for normal diagnostics/details;
- theme preference if the app does not simply follow Windows;
- confirmation preferences only where they do not weaken mandatory safety gates.

Settings must not expose or request Outlook passwords, Graph credentials, Azure app registration, or tenant consent.

## 5. Diagnostics page

Diagnostics can expose technical information hidden from normal UI:

- Classic Outlook availability;
- profile/account/store discovery;
- OutlookHost/IPC status and protocol version;
- last refresh duration/results/errors;
- per-event native Forward capability where useful;
- `StoreID`, `EntryID`, `GlobalAppointmentID` for troubleshooting;
- HRESULT/type context;
- operation history;
- app/package/version information.

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
- V1 has no deletion synchronization.

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
- write action that fails after preview.

A failure in one account or event must not unnecessarily blank the entire application.

## 8. Accessibility and Windows behavior

- Keyboard navigable.
- Visible focus states.
- Accessible names/automation properties for controls and status indicators.
- Account/status meaning must not rely on color alone.
- High-DPI and display scaling supported.
- Light/dark theme supported/follows Windows by default.
- Resizable desktop window with sensible minimum size.
- Long subjects/account names truncate gracefully and remain discoverable through accessible/tool-tip text.

## 9. UI/OutlookHost boundary

The UI never receives COM objects.

`OutlookAligner.App` exchanges versioned plain DTOs with `OutlookAligner.OutlookHost` through local IPC. OutlookHost owns all COM interaction, including capability checks and writes.

The production app must not shell out to the diagnostic CLI and scrape console text as its application protocol.

## 10. Phase 4 acceptance gate

Phase 4 is not complete until:

- a real packaged WinUI 3 application launches;
- `App.xaml` / `MainWindow` / NavigationView exist;
- Calendar, Alignment, Settings and Diagnostics pages are navigable;
- OutlookHost IPC is wired with versioned DTOs;
- account discovery and refresh work from the GUI;
- FullCalendar renders the read model;
- selecting an event exposes comparison/status/authority/action state;
- Forward meeting capability can be represented correctly in the UI even before all production write phases are enabled;
- loading/empty/error/partial-failure states exist;
- normal usage needs no CLI arguments;
- keyboard/high-DPI/theme/accessibility basics are verified.
