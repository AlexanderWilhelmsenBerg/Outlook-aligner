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
- loading, empty, partial/degraded, and error states are explicit.

Details/comparison pane:

- opens when an event/logical group is selected;
- subject and time at top;
- one row/card per account;
- presence/missing state;
- start/end and meaningful differences;
- authority account + reason;
- supported action buttons;
- no raw IDs in the normal pane.

### Action availability

Buttons are capability-driven.

Examples:

- `Forward meeting` appears/enables only when genuine native forwarding is supported for that event;
- `Copy Full` and `Copy Busy` appear when the destination/safety rules permit them;
- unsupported recurring mutations are disabled with a short explanation;
- a disabled capability must never silently degrade into a different operation.

## 3. Alignment page

### Purpose

Provide a work queue for events that need attention rather than requiring the user to visually inspect the calendar.

### Required controls

Filters:

- Missing
- Moved
- DetailsDifferent
- Duplicate
- Conflict
- Ignored

Each row/group shows:

- event title/time;
- status;
- account membership;
- authority;
- concise reason for the discrepancy;
- available action(s).

Selection supports a comparison pane equivalent to Calendar.

### Authority

The user can select/override authority where allowed. The UI displays why the current authority exists:

- KnownOrigin
- UserSelected
- Inferred
- Unknown

`Unknown` must be visually obvious and must block unsafe bulk movement.

## 4. Write-action UX

No write action should behave like a casual toolbar toggle.

### Single event

Before a write, show enough context to answer:

- what item is changing;
- which account is the source/authority;
- which account is the target;
- what fields/action will change;
- what will **not** change.

### Bulk actions

`Move All` and any future bulk operation require an immutable preview before execution.

Preview groups items into:

- Will execute
- Excluded
- Unsafe / requires user decision

The user confirms the resulting plan, not merely the button label.

After execution, show per-item results and allow failures to be retried or inspected independently.

## 5. Settings page

Initial v1 settings:

- future scan horizon;
- account display order/naming preferences;
- privacy defaults for copied events;
- theme/appearance options where appropriate;
- diagnostics verbosity preference if later required.

There are **no credential fields** because Outlook authentication remains owned by the configured Classic Outlook profile.

## 6. Diagnostics page

Diagnostics is for technical state that should not clutter normal use.

Show:

- Classic Outlook availability;
- current profile;
- discovered accounts/stores/calendars;
- last scan time/duration/counts;
- OutlookHost/IPC health;
- per-event capability details when useful, including native Forward availability;
- privacy-safe errors/HRESULT categories;
- operation history;
- application/build/schema versions.

Allow explicit copying/exporting of diagnostic IDs such as StoreID, EntryID, and GlobalAppointmentID. Do not expose those IDs by default on Calendar or Alignment pages.

## 7. Visual behavior

### Account identity

The three accounts need stable visual identity, but color cannot be the only differentiator. Pair color/accent with label/icon/pattern/state text where relevant.

### Alignment state

Use consistent badges/icons/text for Aligned, Missing, Moved, DetailsDifferent, Duplicate, Conflict, and Ignored.

### Theme

Support Windows light/dark mode. Avoid hard-coded assumptions in the WebView calendar; the FullCalendar bridge must receive theme tokens from WinUI.

### Density

Desktop-first. Optimize for useful calendar/comparison density rather than oversized mobile-style controls.

## 8. Accessibility and input

Minimum v1 expectations:

- keyboard navigation through navigation, filters, event list, panes, and confirmation dialogs;
- visible focus state;
- AutomationProperties/accessible labels for custom controls;
- useful status text in addition to color;
- high-DPI scaling;
- no pointer-only action that cannot be performed by keyboard.

## 9. Technical boundaries

### WinUI app

`OutlookAligner.App` owns:

- window/application lifecycle;
- navigation;
- MVVM view models;
- user intent and confirmation;
- presentation state;
- WebView2/FullCalendar hosting;
- local persisted UI preferences.

### OutlookHost

`OutlookAligner.OutlookHost` owns:

- all Classic Outlook COM calls;
- STA execution;
- COM object lifetime;
- capability checks;
- read/write operations explicitly requested by the app.

COM RCWs never enter the UI process/contracts.

### IPC

IPC messages are versioned plain DTOs. The UI must be able to display:

- connected;
- Outlook unavailable;
- partial account/calendar failure;
- host restarted/reconnecting;
- operation failed with privacy-safe reason.

## 10. Phase 4 acceptance

Phase 4 is complete only when all of the following are true:

- [ ] `OutlookAligner.App` contains real WinUI app/window code rather than the current compile-only marker shell;
- [ ] app launches into Calendar;
- [ ] NavigationView switches between Calendar, Alignment, Settings, Diagnostics;
- [ ] the real three-account profile can be refreshed from the GUI;
- [ ] scan horizon can be changed from the GUI;
- [ ] FullCalendar displays the read-only scan results;
- [ ] selecting an event/group opens comparison details;
- [ ] Alignment filters work;
- [ ] authority can be selected locally without writing Outlook data;
- [ ] loading/empty/error/partial states are implemented;
- [ ] Diagnostics exposes OutlookHost/account health and IDs deliberately;
- [ ] normal GUI workflow does not require the CLI;
- [ ] no Outlook writes exist in Phase 4;
- [ ] WinUI build/package is part of CI;
- [ ] keyboard, light/dark, and high-DPI smoke tests pass.

## 11. Later write phases

Phases 5–9 add write capabilities into this UI rather than creating separate command-line workflows. Every new write capability must define:

- where its button/action lives;
- enable/disable capability rules;
- preview/confirmation behavior;
- progress state;
- per-item success/failure presentation;
- recovery/retry behavior;
- privacy impact.

The GUI is therefore part of the architecture and acceptance model, not decoration applied after the synchronization engine is complete.