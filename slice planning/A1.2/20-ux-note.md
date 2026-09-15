# Agent 20 UX note — A1.2

## Review basis

- Current `main` HEAD reviewed: `e8bc60cd6e2acc1bc32e538d661db88262198226`.
- Planning branch HEAD refreshed before writing: `d4c29dac5fc2c1362cda903943e8ad46d1370354`.
- This review is provisional under the one-slice-ahead rule because A1.1 is still the current implementation candidate. Agent 00 must refresh `main` after A1.1 is accepted/merged and revalidate this note before issuing A1.2 `SLICE READY`.

Files/contracts inspected:

- `AGENTS.md`
- `plan.md`, especially Slice A1.2
- `slice planning/Slice-planning.md`
- `slice planning/A1.2/00-brief.md`
- `slice planning/A1.2/20-ux-prompt.md`
- `docs/ui.md`
- `docs/phase-1.md`
- `docs/phase-3.md`
- `src/OutlookAligner.App/MainWindow.xaml`
- `src/OutlookAligner.App/MainWindow.xaml.cs`
- `src/OutlookAligner.App/ViewModels/MainViewModel.cs`
- `web/calendar/src/main.ts`
- `web/calendar/src/styles.css`
- current test layout under `tests/` (no dedicated App/WebView visible-state test project is present at this baseline)

## Current-state UX findings

The current Calendar destination is list-first. It has a WinUI command/status shell, an observation list, a selected-observation detail area, and visible Forward capability/prepare-and-discard diagnostic controls. The frontend already constructs a FullCalendar `dayGridMonth` with its built-in header disabled, but it is not yet the product Calendar host.

The existing WinUI loading/status/notice treatment is useful and should remain the durable user-facing fallback around WebView2. The present Forward/Prepare controls, target-account chooser, and Forward status are Segment-B diagnostic remnants and should no longer appear in the normal Segment-A Calendar workflow.

## Required A1.2 UX contract

### 1. Primary Calendar surface

The Calendar destination opens directly to a real **month grid**. The month grid replaces the current list as the primary browsing surface; a list must not remain as a co-equal default Calendar presentation in this slice.

A1.2 renders the bounded calendar observations available from the current read path. It must not visually claim that multiple copies have already been correlated into one logical event. Until A1.3 supplies that model, separate observations may therefore appear separately when the current data contains separate copies.

### 2. Month navigation and title

There is one coherent period-navigation group in the WinUI Calendar command area, visually associated with the month title:

- **Today**
- **Previous month**
- **Next month**
- the currently displayed month/range title, e.g. `September 2026`

The user must not see a second competing FullCalendar toolbar. From a UX ownership perspective, WinUI owns these visible controls and the surrounding application chrome; the WebView month surface responds to those commands and reports its displayed range back as needed. The exact message plumbing belongs to Agent 30.

The previous/next controls need accessible names that include their action (`Previous month`, `Next month`), not icon-only semantics. `Today` returns the grid to the month containing today's date. The title changes immediately when the displayed month changes.

### 3. WinUI shell that remains around the WebView

Keep the existing native shell responsibilities around the month grid:

- application NavigationView/destinations;
- `Calendar` page heading;
- scan horizon control where currently applicable;
- `Refresh` command;
- refresh/busy indication;
- user-readable status text;
- warning/degraded `InfoBar` or equivalent native status surface;
- Diagnostics destination.

The month navigation/title group belongs in this same command region. The WebView should consume the remaining Calendar content area rather than introducing another application shell inside HTML.

The current explanatory line under `Calendar` may remain only if rewritten to describe the read-only calendar truthfully. It must not imply identity/correlation guarantees that A1.3 has not implemented.

### 4. Loading, empty, error and degraded states

These states remain visible in native WinUI so they still communicate if WebView2 fails to initialize or its content is unavailable.

**Initial/loading state**

- Keep the Calendar page/chrome visible.
- Show the existing busy/progress indication and a concise status such as `Reading Outlook accounts and calendars…`.
- Do not replace the entire window with a spinner.
- Period navigation may be disabled until the calendar surface is ready, but Refresh/status/Diagnostics remain reachable.

**Successful empty range**

- Keep the month grid visible so the user can understand which month is empty and can navigate elsewhere.
- Show an explicit native empty message such as `No calendar events were found in this range.`
- Do not present empty data as an Outlook failure.

**OutlookHost/read failure**

- Keep the native Calendar shell visible even if no events can be rendered.
- Show a clear failure status and native notice with the existing direction to Diagnostics for technical details.
- No raw Outlook IDs belong in this normal error surface.

**Partial/degraded result**

- Render healthy returned observations in the month grid.
- Keep the native warning/degraded notice visible at the same time; one failed account/item must not make healthy results disappear.
- Wording should describe incomplete results, not classify events as Missing/Moved/Conflict. Those cross-account semantics are later slices.

**WebView2/calendar-host failure**

- The native shell must communicate that the calendar view could not be displayed and keep Refresh/Diagnostics usable.
- Do not leave a blank white/black rectangle as the only indication of failure.

### 5. Selection behavior in A1.2

Calendar event selection remains **read-only** and means only: select the rendered observation represented by that calendar item.

At this slice the selected-item detail may continue to show truthful single-observation fields already available today, such as:

- subject;
- source account/display identity;
- observed time;
- location;
- recurrence fact already present on that observation.

Selection must not:

- merge sibling observations into one logical event;
- infer authority/origin;
- claim alignment health;
- claim an account is missing;
- expose A1.3/A2 comparison semantics;
- expose write actions.

If retaining a detail pane would materially crowd the month grid, it may be reduced to a simple read-only selected-observation region for this slice. The essential requirement is that clicking/keyboard-selecting a calendar item has a deterministic, truthful result without inventing logical-event semantics.

### 6. Forward/Prepare removal from normal Calendar workflow

Remove or hide from the normal Calendar destination:

- the `Forward` capability row;
- test target-account picker;
- `Check Forward`;
- `Prepare Forward (discard)`;
- explanatory Forward diagnostic text.

A1.2 does **not** require deletion of dormant Segment-B spike code, diagnostic services, commands, or host capability code where retaining them avoids unrelated churn. They simply must not be presented as normal Segment-A Calendar actions.

No Copy, Move, bulk reconcile, recurring write, send, or other mutation affordance is introduced.

### 7. Accessibility and resilient presentation

A1.2 must preserve basic Windows desktop accessibility rather than treating the WebView as an inaccessible canvas:

- All native commands and calendar items are keyboard reachable with visible focus.
- Tab order moves predictably through month controls, horizon/Refresh, calendar content, and any retained read-only detail region.
- Previous/next icon buttons have accessible names.
- Event subjects must remain readable at high DPI/text scaling; long subjects truncate/wrap without breaking the grid, with the full accessible name available to assistive technology/tooltips where practical.
- Long account/source text in any retained detail area wraps or truncates safely rather than forcing horizontal overflow.
- Light/dark presentation follows the Windows/app theme closely enough that the WebView does not look like a permanently light island in a dark shell (or vice versa), with readable contrast in both.
- Month navigation and current-range information are not conveyed by position/icon alone.
- Native loading/error/degraded/fallback communication remains available independently of the WebView DOM.

A1.2 does not yet need account-marker accessibility, alignment-status iconography, or recurrence-correlation accessibility beyond the truthful per-observation recurrence information already exposed; those belong to later slices.

## A1.2 must not visually imply later semantics

Before A1.3/A2, the month view must stay deliberately neutral. Do **not** introduce:

- one-card-per-logical-event claims for observations not yet correlated by the accepted A1.3 model;
- top-left per-account presence dots;
- green/yellow/red alignment health;
- `Aligned`, `Missing`, `Moved`, `Duplicate`, `Conflict`, or equivalent status badges inferred by the WebView;
- authority/origin labels derived from calendar rendering order;
- reconciliation/report action affordances;
- selectable account-status legends that imply A2 behavior;
- recurrence grouping/exception claims beyond the current observation's known recurrence fact;
- raw StoreID, EntryID, or GlobalAppointmentID in normal Calendar presentation.

FullCalendar is a presentation surface in this slice, not a second correlation/status engine.

## Acceptance / manual UX checks

On the real Windows/Classic Outlook profile, verify at minimum:

1. App opens Calendar with a usable month grid as the primary content.
2. `Today`, `Previous month`, and `Next month` are keyboard reachable and move the grid correctly; the visible month title tracks the shown range.
3. Refresh/horizon/status remain in the native shell and Outlook data appears without PowerShell or raw Outlook identifiers.
4. Selecting an event gives a truthful read-only single-observation selection/detail result and exposes no write action.
5. The normal Calendar workflow contains no Forward/Prepare/test-target controls.
6. A successful zero-event month leaves the grid visible and clearly reports that there are no events in the range.
7. Simulated/real host failure shows a native failure state and leaves Diagnostics reachable rather than presenting a blank WebView.
8. A partial-result case keeps healthy events visible and shows a native degraded warning without labelling events Missing/Moved/Conflict.
9. Keyboard navigation and visible focus work across native controls and rendered event items.
10. At practical resized desktop dimensions and Windows display/text scaling, the month grid, title, controls, and long subjects remain usable without essential controls disappearing.
11. Light and dark Windows/app themes keep the shell and calendar readable and visually coherent.
12. No A1.3/A2 account markers, health colors, logical-event claims, or reconciliation actions appear.

## Blockers / dependencies

No unresolved **UX** decision blocks Agent 40 once Agent 00 revalidates this note against the post-A1.1 repository state.

A1.2 implementation remains procedurally blocked until A1.1 is accepted/merged and Agent 00 performs the required refresh/reconciliation. Any A1.1 change that materially alters the frontend host/toolchain assumptions should trigger a focused recheck, but it does not currently create a UX blocker.

UX READY
