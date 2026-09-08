# Outlook Aligner — Implementation Plan

Status: **Phase 0 complete ✅. Phase 1 complete ✅ and merged. Phase 2 native forwarding complete ✅ and merged in PR #5. Phase 3 UI-assisted read/alignment development in progress 🚧.**

Last reviewed: 2026-09-08

## 1. Goal

Build a **Windows desktop application with a real graphical user interface** that uses the Classic Outlook COM/Object Model as its only Outlook integration boundary. It discovers the calendar-capable accounts already configured in one Outlook profile, reads their calendars for a configurable future horizon, correlates the same logical meeting across accounts, shows differences clearly, and lets the user deliberately align them.

Microsoft Graph is explicitly out of scope. The design must not require Azure app registration, Graph delegated permissions, tenant administrator consent, or separate Outlook credentials.

The command-line programs used in early phases are **diagnostic and acceptance-test harnesses only**. They are not the shipped product and do not satisfy the v1 UI requirement.

### PR sizing rule

Phases are roadmap/architecture buckets, **not mandatory PR-sized units**.

Implementation should be split into the smallest coherent, independently testable PRs that make review and real-machine validation easier. A PR may deliver only one slice of a phase. Do not hold a safe, useful increment open merely to finish an entire phase, and do not broaden a PR just to match a phase heading.

Every implementation PR remains unmerged until the user explicitly approves merge.

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

Current read-model-only implementation also uses explicit `RecurrenceIdentityUnresolved` and `Uncorrelated` safety states while identity work remains incomplete.

### Transfer actions

#### Forward meeting

Phase 2 proved a genuine Classic Outlook forwarding route for an accepted Calendar meeting on the user's real profile and merged that foundation in PR #5.

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

Validated Outlook Aligner-managed copy metadata may establish `KnownOrigin` only when the recorded source account resolves to exactly one observed non-managed source appointment and the managed copies have internally consistent provenance.

Conflicting, incomplete, unsupported-schema, or unreadable managed metadata must fail closed and block automatic authority/action planning.

`Move selected` updates only non-authoritative managed/local copies to authoritative Start/End. The authoritative original is untouched.

`Move all` builds a preview first and excludes conflicts, unknown authority, ignored events, unsupported recurrence mutations, and failed safety checks.

`LastModificationTime` must never become an implicit latest-wins authority rule.

### Deletion

Deletion synchronization is explicitly excluded from v1. Missing data is never interpreted as permission to delete another calendar's item.

## 3. Event identity model

### EntryID is a locator, not identity

Store `StoreID + EntryID` only as the current Outlook locator cache. EntryID may change after moves or other Outlook operations.

### Primary native correlation

Use Outlook `GlobalAppointmentID` as the primary native meeting-correlation candidate for non-recurring ordinary Outlook events.

### Outlook Aligner-managed identity

Managed copies use custom properties:

- `OutlookAligner.SyncGroupId`
- `OutlookAligner.SourceGlobalAppointmentId`
- `OutlookAligner.SourceAccountId`
- `OutlookAligner.CopyType`
- `OutlookAligner.SchemaVersion`

Current read-side schema version is `1`; supported copy-type markers are `Full` and `Busy`.

A validated managed local copy correlates using its stored `SourceGlobalAppointmentId`, not its own native GlobalAppointmentID. This is necessary because a locally created Outlook copy may receive a different native ID.

Do not assume custom properties propagate through true meeting forwarding.

### Recurrence

Series masters, normal occurrences, modified exceptions, and deleted occurrences require separate identity handling. A moved occurrence must remain associated with its original occurrence rather than becoming an unrelated standalone event.

Current UI correlation intentionally leaves recurring events unresolved until this identity is proven.

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

The current Phase 3 development UI launches OutlookHost as a child process and exchanges versioned JSON/text through redirected standard streams. This is a **transitional local transport**, not the final production application protocol. Before production write workflows are considered complete, the transport converges on versioned local IPC/named pipes. UI/view-model/Core contracts must remain transport-independent.

### UI/host boundary invariants

- `OutlookAligner.App` owns presentation, view state, confirmation flows, and user interaction.
- `OutlookAligner.OutlookHost` owns all Outlook COM access.
- COM objects never cross IPC or enter Core/UI DTOs.
- IPC/contracts are versioned plain DTOs.
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

The current Phase 3 authority override is session-only; persistence belongs in a later focused PR.

## 6. Phase roadmap

### Phase 0 — repository/toolchain bootstrap ✅

Complete.

### Phase 1 — Outlook COM discovery/read probe ✅

Complete and merged in PR #4. Proved real Classic Outlook account/store/calendar discovery, bounded recurrence-safe reads, privacy defaults, packaged Outlook interop, and production-machine execution.

### Phase 2 — native meeting-forwarding technical spike ✅

Complete and merged in PR #5.

Proved on a real accepted Calendar meeting:

- native Forward capability detection on Calendar `AppointmentItem`;
- native `MeetingItem` production through Outlook's Forward event;
- cancellation/prepare behavior;
- exact recipient preparation and `SendUsingAccount` pinning;
- explicitly confirmed `MeetingItem.Send()`;
- real delivery to another user-controlled Outlook account;
- no vCalendar fallback.

Additional reliability sampling can continue through later GUI-assisted test PRs without reopening or enlarging PR #5.

### Phase 3 — UI-assisted development and testing foundation 🚧

Phase 3 is intentionally split into smaller PR-sized increments.

Current read/alignment increment:

- real `App.xaml` / `MainWindow` WinUI application;
- NavigationView with Calendar, Alignment, Settings and Diagnostics;
- automatic Outlook account/calendar discovery through OutlookHost;
- event list/detail selection across discovered accounts;
- source and target account selection without manual Outlook IDs;
- per-event native Forward capability check;
- safe **Prepare Forward (discard)** action that never sends;
- preliminary non-recurring correlation and discrepancy states;
- protocol-v2 read-only managed-copy metadata;
- fail-closed handling of suspicious managed metadata;
- `KnownOrigin` authority resolution from validated provenance plus session-only manual authority;
- read-only Move Selected preview targeting validated managed copies only;
- packaged self-contained Windows development bundle containing UI + OutlookHost;
- user-facing partial-result warning plus Diagnostics for technical evidence.

This increment deliberately does **not** add Copy writes, Move writes, recurring correlation writes, GUI Forward send, persistence, or production named-pipe IPC.

Likely subsequent Phase 3 PRs:

1. focused real-machine UI acceptance/repair of the current read/alignment increment;
2. richer Calendar/FullCalendar integration once a real Node 24 lockfile can be generated and CI can switch to `npm ci`;
3. deliberate GUI confirmation/send flow for native Forward;
4. further error/degraded-state and accessibility hardening;
5. persistence/transport slices only when needed by the next production behavior.

See [`docs/phase-3.md`](docs/phase-3.md).

### Phase 4 — deeper identity and recurrence model

Use the Phase 3 UI as the working test surface while proving the identity cases that remain intentionally unsupported:

- GlobalAppointmentID behavior across the three accounts under more scenarios;
- series master vs occurrence vs exception identity;
- moved occurrence handling;
- deleted occurrence diagnosis without deletion sync;
- persistence of authority/origin decisions;
- robust comparison/detail behavior driven by logical groups rather than raw rows.

Some non-recurring correlation and managed-copy foundations have already been pulled forward into Phase 3 because they are required to make the UI safe and useful. Phase 4 should deepen those foundations rather than reimplement them.

### Phase 5 — Copy Full

Implement managed full-detail copies with identity properties, preview, user-facing errors, and operation history.

### Phase 6 — Copy Busy

Implement privacy placeholders and associated UI/preview.

### Phase 7 — Move Selected

Allow deliberate time alignment of managed non-authoritative copies only. This begins only after managed-copy ownership and authority are proven for the write path. The UI must show source, authority, target, old time, new time, and unsupported reasons before execution.

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
- Prefer small coherent PRs over phase-sized PRs when that improves review/testability.
- Never silently fall back from native Forward to vCalendar/ICS.
- Never infer deletion permission from a missing event.
- Never mutate the authoritative original as part of Move selected/all.
- Never use `LastModificationTime` as automatic authority.
- Never leak meeting body/attendees/online links into ordinary diagnostics.
- All production write actions require explicit user intent, capability checks, and operation history.
- Bulk writes require preview + confirmation.
- Do not introduce Move writes merely to make the UI testable; identity, managed-copy ownership, and authority prerequisites remain mandatory.
- Invalid or unreadable managed-copy metadata must fail closed rather than being silently treated as ordinary trusted data.

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
