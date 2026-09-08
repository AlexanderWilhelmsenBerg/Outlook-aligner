# Phase 3 — UI-assisted development and testing foundation

Status: **In progress 🚧 — stacked on Phase 2**

Started: 2026-09-08

## Why this phase moved forward

The Phase 1/2 command-line harnesses successfully proved the Classic Outlook COM boundary and genuine native meeting forwarding, but continued manual testing became inefficient because the user had to copy `GlobalAppointmentID`, `EntryID`, source SMTP addresses, and target addresses into PowerShell commands.

Phase 3 therefore brings a usable WinUI application shell forward **before** the deeper identity prototype. Future identity, recurrence, Forward, Copy and Move work should increasingly be exercised through the application rather than by assembling diagnostic commands manually.

## First usable UI slice

The first slice provides a real unpackaged WinUI 3 desktop application with the production navigation shape:

- Calendar;
- Alignment;
- Settings;
- Diagnostics.

Calendar is functional first. The other destinations are scaffolded so later slices can grow without replacing the shell.

On launch/refresh the app:

1. starts the sibling `OutlookAligner.OutlookHost.exe`;
2. runs the bounded JSON calendar probe with explicit meeting details enabled for the local UI;
3. validates the versioned OutlookHost protocol;
4. displays events from all discovered Outlook accounts;
5. keeps `StoreID`, `EntryID`, and `GlobalAppointmentID` internal to the selected event;
6. derives the source account from the event instead of asking the user to type it;
7. offers the other discovered SMTP-capable accounts as target choices.

## Forward testing from the UI

For a selected event the UI currently offers:

### Check Forward

Runs the proven Calendar native-Forward capability probe. No Forward is executed.

### Prepare Forward (discard)

This is the safe UI test action. It:

1. re-runs the native Forward capability check immediately before preparation;
2. uses the selected event's hidden source SMTP / `GlobalAppointmentID` / `EntryID`;
3. uses the selected target Outlook account's SMTP address;
4. asks OutlookHost to create the native Calendar Forward;
5. resolves exactly that one target recipient;
6. pins `SendUsingAccount` to the source account;
7. discards the native forwarded MeetingItem unsent.

The first Phase 3 UI slice **cannot send** a meeting. Real send remains behind the explicit Phase 2 CLI confirmation gate until the UI confirmation/error model is implemented deliberately.

## Diagnostics

Raw host output, warnings and technical identifiers are available in Diagnostics/technical expanders. They are not required for normal selection or action workflows.

This preserves the product rule that internal Outlook identifiers belong in Diagnostics rather than being part of ordinary user interaction.

## Temporary UI/host transport

The first slice launches OutlookHost as a child process and exchanges JSON/text over redirected standard streams.

This is intentionally a transitional local transport because it lets the usable UI arrive without duplicating COM logic. The boundary invariants already hold:

- all Outlook COM stays inside OutlookHost;
- UI receives only plain DTOs/results;
- OutlookHost remains an interactive STA process;
- no COM RCW crosses into WinUI.

Before production write workflows are considered complete, this transport should converge on the planned versioned local IPC/named-pipe host model. View models and DTO contracts should not depend on the transport choice.

## Distribution for development testing

CI publishes a self-contained, unpackaged win-x64 test bundle containing:

- `OutlookAligner.App.exe`;
- `OutlookAligner.OutlookHost.exe` beside it;
- the self-contained Windows App SDK/.NET payload required by the WinUI executable.

The UI uses the sibling OutlookHost automatically. Development builds can override the host path with `OUTLOOK_ALIGNER_HOST_PATH`.

## Next development slices before broad manual testing

1. make the Calendar page richer and integrate the existing FullCalendar/WebView2 assets;
2. introduce correlation/identity view models so logical meetings can be grouped across accounts;
3. implement Alignment discrepancy states and authority selection;
4. add user-facing error/partial-failure presentation instead of relying on raw Diagnostics;
5. add deliberate GUI confirmation for a real native Forward send;
6. implement managed-copy identity foundations required by Copy/Move;
7. implement Move Selected preview/action only after authority and managed-copy safety are established.

The remaining Phase 2 reliability matrix can then be executed through this UI instead of manual identifier lookup.

## Safety

- No automatic PR merge.
- No vCalendar fallback.
- The UI does not perform a real send in this first slice.
- Prepare always revalidates Forward capability.
- Source and target accounts are derived from discovered Outlook data.
- Technical identifiers remain available for diagnosis but are not user input.
- Move/copy writes are not introduced until their identity/authority prerequisites exist.
