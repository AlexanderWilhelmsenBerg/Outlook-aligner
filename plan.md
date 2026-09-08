# Outlook Aligner — Implementation Plan

Status: **Phase 0 complete ✅. Phase 1 complete ✅ and merged. Phase 2 native-forwarding mechanism proven end-to-end; broad reliability testing deferred into the UI-assisted test phase. Phase 3 UI-assisted development in progress 🚧.**

Last reviewed: 2026-09-08

## 1. Goal

Build a **Windows desktop application with a real graphical user interface** that uses the Classic Outlook COM/Object Model as its only Outlook integration boundary. It discovers the calendar-capable accounts already configured in one Outlook profile, reads their calendars for a configurable future horizon, correlates the same logical meeting across accounts, shows differences clearly, and lets the user deliberately align them.

Microsoft Graph is explicitly out of scope. The design must not require Azure app registration, Graph delegated permissions, tenant administrator consent, or separate Outlook credentials.

The command-line programs used in early phases are **diagnostic and acceptance-test harnesses only**. They are not the shipped product and do not satisfy the v1 UI requirement.

## 2. Confirmed product requirements

### User interface — required for v1

- Outlook Aligner ships as a **WinUI 3 desktop application**.
- A functional GUI is a v1 acceptance criterion, not an optional later enhancement.
- The main app uses a persistent navigation shell with four primary views:
  1. **Calendar**
  2. **Alignment**
  3. **Settings**
  4. **Diagnostics**
- Calendar is the primary/home view.
- The UI must expose account discovery, refresh, scan horizon, event comparison, authority, status, and supported alignment actions without requiring CLI use.
- The UI must discover and carry Outlook technical identity itself. Normal workflows must not require the user to copy SMTP addresses, StoreIDs, EntryIDs, or GlobalAppointmentIDs.
- All write actions must clearly identify source/target account and intended effect before execution.
- Bulk writes require an immutable preview and explicit confirmation.
- The UI must show partial failures instead of hiding them or aborting the whole operation.
- Normal UI must not expose raw StoreIDs, EntryIDs, GlobalAppointmentIDs, HRESULTs, or internal metadata; those belong in Diagnostics.
- Support keyboard navigation, high-DPI scaling, light/dark theme, and accessible labels/automation properties.
- UI code never receives COM objects. It communicates with `OutlookAligner.OutlookHost` through plain DTOs/local IPC.
- Detailed UI contract: [`docs/ui.md`](docs/ui.md).

### Outlook environment

- Windows desktop application.
- Classic Outlook for Windows is required.
- The three target accounts are configured in the same Outlook profile.
- Accounts/calendars are discovered automatically through Outlook COM.
- New Outlook is out of scope because it does not expose the required Classic Outlook Object Model automation surface.
- Outlook Aligner stores no Outlook passwords or access tokens.

### Calendar scan

- Scan starts at today.
- Future horizon is configurable; default 90 days.
- Recurrences must be expanded only inside the bounded horizon.
- Refresh is user-driven in v1.
- One failed calendar/item must not abort the entire scan.
- The overlap window is half-open: `eventEnd > windowStart && eventStart < windowEnd`.

### Logical event states

- `Aligned`
- `Missing`
- `Moved`
- `DetailsDifferent`
- `Duplicate`
- `Conflict`
- `Ignored`
- `DeletedOrMissing` for diagnosis only; v1 never automatically deletes.

### Transfer actions

#### Forward meeting

Phase 2 has proven a genuine Classic Outlook forwarding route for an accepted Calendar meeting on the user's real profile.

Primary production mechanism:

1. reopen the exact accepted Calendar `AppointmentItem` in the selected source Store;
2. verify event identity (`StoreID + EntryID` locator plus `GlobalAppointmentID`) and meeting state;
3. query Outlook's built-in `Forward` command and require it to be valid, visible, and enabled;
4. invoke Outlook's native Calendar Forward command;
5. capture the native `MeetingItem` supplied by `AppointmentItem.Forward`;
6. require zero pre-existing recipients;
7. add/resolve only the explicitly intended target account(s) permitted by the product action;
8. pin `SendUsingAccount` to the selected source Outlook account;
9. send only after the GUI's explicit user confirmation and all capability/safety checks pass.

The real-machine Phase 2 spike proved this path end-to-end on an accepted Kverneland meeting forwarded to a user-controlled Knowit account. The target received the forwarded meeting.

`AppointmentItem.ForwardAsVcal()` is not equivalent and must never be silently presented as native forwarding.

The older retained-`MeetingItem` recovery path is not a prerequisite for production Forward because real testing showed accepted meetings may no longer have a recoverable request in Inbox/Deleted Items. It may remain as a secondary path only if later testing finds a clear benefit worth the extra complexity.

The GUI treats Forward as a **per-event capability**. If Outlook reports native Forward unavailable, the action must be disabled/hidden with a clear reason and must not fall back to fake ICS forwarding.

#### Copy Full

Create an Outlook Aligner-managed calendar copy containing the supported full details required by the product: subject, time, all-day state, location, body, online-meeting link where available, reminder settings, busy state, sensitivity, and supported recurrence data.

A copied item is a managed local copy; it must not be described as making the target account an attendee of the organizer's original meeting.

#### Copy Busy

Create a privacy-preserving managed placeholder containing time/all-day/busy state and optionally a generic `Busy` subject. By default it must not copy body, attendees, Teams link, or location.

### Move actions and authority

Each logical event has an authoritative account. Normal rule: `AuthorityAccount = OriginAccount`.

Authority reason is recorded as:

- `KnownOrigin`
- `UserSelected`
- `Inferred`
- `Unknown`

`Move selected` updates only non-authoritative managed/local copies to authoritative Start/End. The authoritative original is untouched.

`Move all` builds a preview first and excludes conflicts, unknown authority, ignored events, unsupported recurrence mutations, and failed safety checks.

`LastModificationTime` must never become an implicit latest-wins authority rule.

### Deletion

Deletion synchronization is explicitly excluded from v1. Missing data is never interpreted as permission to delete another calendar's item.

## 3. Event identity model

### EntryID is a locator, not identity

Store `StoreID + EntryID` only as the current Outlook locator cache. EntryID may change after moves or other Outlook operations.

### Primary native correlation

Use Outlook `GlobalAppointmentID` as the primary native meeting-correlation candidate.

### Outlook Aligner-managed identity

Managed copies use custom properties such as:

- `OutlookAligner.SyncGroupId`
- `OutlookAligner.SourceGlobalAppointmentId`
- `OutlookAligner.SourceAccountId`
- `OutlookAligner.CopyType`
- `OutlookAligner.SchemaVersion`

Do not assume custom properties propagate through true meeting forwarding.

### Recurrence

Series masters, normal occurrences, modified exceptions, and deleted occurrences require separate identity handling. A moved occurrence must remain associated with its original occurrence rather than becoming an unrelated standalone event.

## 4. Architecture

### Chosen stack

- C# 14 / .NET 10 LTS.
- WinUI 3 / Windows App SDK for the production desktop UI.
- WebView2 + FullCalendar Vanilla for calendar visualization.
- TypeScript + Vite for calendar assets.
- SQLite via Microsoft.Data.Sqlite + Dapper.
- CommunityToolkit.Mvvm.
- Microsoft.Extensions.Hosting/configuration/logging + Serilog.
- xUnit v3 on Microsoft Testing Platform.
- `Microsoft.Testing.Extensions.CodeCoverage` for MTP-native coverage.
- NSubstitute only when a handwritten fake is not clearer.
- BenchmarkDotNet.
- GitHub Actions + Dependabot.

Microsoft Graph and MSAL are not dependencies.

### Process model

```text
OutlookAligner.App.exe
  WinUI 3 / NavigationView / MVVM / alignment engine / SQLite / WebView2
        |
        | local IPC
        v
OutlookAligner.OutlookHost.exe
  interactive user process
  STA entry thread
  Classic Outlook COM/Object Model
```

The Outlook host is not a Windows service.

During the first Phase 3 UI-assisted slice, the app may launch OutlookHost as a child process and exchange versioned JSON/text through redirected standard streams. This is a **transitional local transport**, not a change to the architecture boundary. Before production write workflows are considered complete, the transport converges on the planned versioned local IPC/named-pipe host. UI/view-model contracts must remain transport-independent.

### UI/host boundary invariants

- `OutlookAligner.App` owns presentation, view state, confirmation flows, and user interaction.
- `OutlookAligner.OutlookHost` owns all Outlook COM access.
- COM objects never cross IPC or enter Core/UI DTOs.
- IPC contracts are versioned plain DTOs.
- A UI crash must not leave a hidden Outlook write operation running.
- The app must remain usable for read-only inspection when a write capability is unavailable.

### COM boundary invariants

- Convert Outlook objects immediately to plain DTOs.
- Keep Outlook automation on an STA thread.
- Never pass COM RCWs over IPC.
- Release COM objects deterministically, especially recurrence objects and collection enumerators.
- Do not call `Application.Quit()` merely because Outlook Aligner attached to or started Outlook.
- Catch/report one item or folder failure without aborting the whole scan where safe.

## 5. Persistence

Planned local SQLite tables:

- `Accounts`
- `SyncGroups`
- `EventMembers`
- `UserOverrides`
- `OperationHistory`
- `SchemaMigrations`

No Outlook passwords or tokens are stored.

## 6. Phase roadmap

### Phase 0 — repository/toolchain bootstrap ✅

Complete.

### Phase 1 — Outlook COM discovery/read probe ✅

Complete and merged in PR #4. Proved real Classic Outlook account/store/calendar discovery, bounded recurrence-safe reads, privacy defaults, packaged Outlook interop, and production-machine execution.

### Phase 2 — native meeting-forwarding technical spike 🚧

End-to-end Calendar-command mechanism is proven on a real accepted meeting:

- native Forward capability detected on Calendar `AppointmentItem`;
- native `MeetingItem` produced by Outlook's Forward event;
- cancellation/prepare behavior proven;
- exact recipient preparation and `SendUsingAccount` pinning proven;
- explicitly confirmed `MeetingItem.Send()` completed;
- forwarded meeting arrived in another user-controlled Outlook account;
- no vCalendar fallback used.

The broad reliability matrix is intentionally deferred until the UI-assisted test workflow can select meetings/accounts itself. PR #5 remains the forwarding foundation and is not merged without explicit user approval.

### Phase 3 — UI-assisted development and testing foundation 🚧

Bring a usable application shell forward so subsequent development can be tested efficiently without manual identifier lookup.

First slice:

- real `App.xaml` / `MainWindow` WinUI application;
- NavigationView with Calendar, Alignment, Settings and Diagnostics;
- automatic Outlook account/calendar discovery through OutlookHost;
- event list/detail selection across discovered accounts;
- source account derived from the selected event;
- target account chosen from other discovered Outlook accounts;
- technical IDs hidden from the normal workflow and available only for Diagnostics;
- per-event native Forward capability check;
- safe **Prepare Forward (discard)** action that revalidates capability and never sends;
- packaged self-contained Windows test bundle containing UI + OutlookHost.

Next Phase 3 slices before broad manual testing:

- correlation/grouping by native meeting identity;
- preliminary Alignment states and comparison view;
- authority/origin model presentation;
- FullCalendar/WebView2 event visualization;
- InfoBar-style user-facing error/partial-failure handling;
- deliberate GUI confirmation for real native Forward send;
- managed-copy identity foundations required by Copy/Move;
- diagnostics that make recurrence/source-account portability testing easy.

See [`docs/phase-3.md`](docs/phase-3.md).

### Phase 4 — identity, correlation and read-only alignment model

Use the Phase 3 UI as the working test surface while proving:

- GlobalAppointmentID behavior across the three accounts;
- series master vs occurrence vs exception identity;
- moved occurrence handling;
- managed-copy custom-property schema;
- authority/origin confidence model;
- `Aligned`, `Missing`, `Moved`, `DetailsDifferent`, `Duplicate`, and `Conflict` classification;
- comparison/detail UI driven by logical groups rather than raw calendar rows.

The production Calendar/Alignment read experience should mature during this phase instead of waiting until after the identity prototype.

### Phase 5 — Copy Full

Implement managed full-detail copies with identity properties, preview, user-facing errors, and operation history.

### Phase 6 — Copy Busy

Implement privacy placeholders and associated UI/preview.

### Phase 7 — Move Selected

Allow deliberate time alignment of managed non-authoritative copies only. This begins only after managed-copy identity and authority are proven. The UI must show source, authority, target, old time, new time, and unsupported reasons before execution.

### Phase 8 — Move All

Preview-first batch movement with conflict/unsupported exclusions and immutable execution plan.

### Phase 9 — recurring writes

Only after recurrence identity is proven, add supported recurring Copy/Move operations with explicit unsupported-case handling.

### Phase 10 — integrated testing, hardening and release

Use the production-like UI for the deferred broad test matrix:

- native Forward across recurring/Teams meetings and different source accounts;
- unavailable/disabled Forward behavior;
- Copy Full / Copy Busy safety and fidelity;
- Move Selected / Move All with authority/conflict/error cases;
- recurrence and exception cases;
- partial Outlook/COM failures and recovery;
- packaging/installer;
- upgrade/migration tests;
- crash-safe operation history;
- diagnostics/export;
- accessibility and DPI verification;
- performance limits;
- documentation and release checklist.

## 7. Global safety rules

- Never merge a PR automatically; merge only after explicit user approval.
- Never silently fall back from native Forward to vCalendar/ICS.
- Never infer deletion permission from a missing event.
- Never mutate the authoritative original as part of Move selected/all.
- Never use `LastModificationTime` as automatic authority.
- Never leak meeting body/attendees/online links into ordinary diagnostics.
- All production write actions require explicit user intent, capability checks, and operation history.
- Bulk writes require preview + confirmation.
- Do not introduce Move writes merely to make the UI testable; identity, managed-copy ownership, and authority prerequisites remain mandatory.

## 8. V1 acceptance gate

V1 is not complete until all of the following are true:

- packaged WinUI 3 desktop GUI exists and launches normally;
- Calendar/Alignment/Settings/Diagnostics are functional;
- no normal workflow requires CLI/PowerShell or manual Outlook IDs;
- all three configured Outlook accounts are discoverable;
- bounded calendar scanning/correlation works;
- Forward/Copy/Move actions truthfully represent their semantics;
- Forward is enabled only when native Outlook capability checks pass;
- authority/conflict behavior is visible before writes;
- no deletion sync exists;
- operation history/diagnostics exist;
- tests, packaging and migration checks pass;
- user has manually tested the release candidate on the real Classic Outlook profile.
