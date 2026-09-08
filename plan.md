# Outlook Aligner — Implementation Plan

Status: **Phase 0 complete ✅. Phase 1 complete ✅ and merged. Phase 2 native-forwarding spike in progress 🚧 on PR #5.**

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

Expose a real **Forward meeting** action only if Phase 2 proves genuine Classic Outlook forwarding for accepted meetings.

- Preferred direct COM mechanism: recover a native `MeetingItem` and call `MeetingItem.Forward()`.
- Accepted meetings normally exist in Calendar as `AppointmentItem`; Phase 2 must also determine whether Outlook's own Calendar Forward UI command can be safely automated when the original request is no longer retained.
- `AppointmentItem.ForwardAsVcal()` is not equivalent and must never be silently presented as native forwarding.
- If native forwarding is only conditionally available, the GUI must disable/hide Forward for unsupported events and explain why.

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
- Explicitly release locally owned COM references.
- Avoid COM `foreach` patterns where hidden enumerator RCWs matter.
- Never use `EntryID` as cross-account logical identity.
- Do not call `Outlook.Application.Quit()` merely because OutlookHost activated Outlook.
- Normal diagnostics contain HRESULT/operation context, not meeting body/attendees/Teams URLs.

## 5. Persistence plan

SQLite tables planned:

- `Accounts`
- `SyncGroups`
- `EventMembers`
- `UserOverrides`
- `OperationHistory`
- `SchemaMigrations`

No credentials or cloud tokens are stored.

## 6. UI plan

The production UI begins after the COM/forwarding/identity spikes, but the UX contract is defined now so earlier phases produce the DTOs and capability flags it needs.

### Application shell

Use WinUI `NavigationView` with Calendar as the default page, plus Alignment, Settings, and Diagnostics. The shell includes Outlook connection/profile status and refresh progress without blocking navigation.

### Calendar

- configurable horizon and Refresh command;
- account legend/cards for the three discovered accounts;
- FullCalendar-based day/week/month visualization;
- visually distinguish accounts and logical alignment states;
- selecting an event opens a details/comparison pane rather than navigating away;
- pane shows per-account presence, time, status, authority, and available actions;
- Forward / Copy Full / Copy Busy availability comes from explicit capability state, not optimistic UI assumptions.

### Alignment

- filterable list/grid of logical events requiring attention;
- filters for Missing, Moved, DetailsDifferent, Duplicate, Conflict, Ignored;
- per-event authority selector with reason/confidence shown;
- single-event action preview;
- bulk `Move all` builds a preview containing included, excluded, and unsafe items before confirmation;
- execution results remain visible per event.

### Settings

- scan horizon;
- account display/order preferences;
- privacy defaults for copied events;
- theme/appearance where appropriate;
- no credential entry fields.

### Diagnostics

- Outlook/profile/account/store health;
- last scan time and counts;
- capability information such as whether genuine Forward is available for a selected event;
- operation history and privacy-safe errors;
- optional copy/export of diagnostic IDs/details;
- raw identifiers remain here rather than in normal Calendar/Alignment views.

Every bulk write operation requires preview + explicit confirmation. See [`docs/ui.md`](docs/ui.md) for the detailed interaction and acceptance contract.

## 7. Verified implementation baseline

The detailed latest-stable version matrix is maintained in [`docs/phase-0.md`](docs/phase-0.md). Current baseline remains .NET 10/C# 14, Windows App SDK/WinUI 3, WebView2, Outlook PIA, SQLite/Dapper, CommunityToolkit.Mvvm, Serilog, xUnit v3/MTP, BenchmarkDotNet, Node 24 LTS, TypeScript/Vite/FullCalendar, and pinned GitHub Actions as recorded there.

### Phase 1 packaging lesson

Real-machine testing proved that a successful single-file build plus `--help` smoke test was not enough to validate Office interop packaging. Outlook PIA metadata is now embedded from the resolved reference, and CI runs `--interop-check` against the actual published EXE before upload.

That check is a permanent regression gate.

## 8. Engineering quality and privacy

### C# gates

- nullable reference types;
- implicit usings;
- C# 14;
- `AnalysisLevel=latest-Recommended`;
- warnings as errors;
- repository `.editorconfig`;
- `dotnet format --verify-no-changes`.

### Frontend gates

- TypeScript strict mode;
- ESLint for supported sources;
- Prettier;
- Vite production build.

The first PR that materially changes `web/calendar` must generate a real `package-lock.json` under Node 24 and switch CI/local verification to `npm ci`.

### Privacy

Normal logs/output may include counts, sync IDs/global IDs, operation type, HRESULT/error category, and timings. They must not include subject, body, attendees, Teams links, or meeting location by default.

Diagnostic details are explicit opt-in.

## 9. Testing strategy

### Hosted CI

Pure Core/Persistence/contract tests run without Outlook. Windows CI builds/publishes OutlookHost but cannot perform live Outlook integration because hosted runners do not have the user's Outlook profile.

### Manual Outlook integration

A downloadable diagnostic executable is produced for Outlook-specific phases. Manual tests are merge gates whenever behavior depends on real Classic Outlook, Exchange/Teams state, or Outlook profile contents.

### UI testing

From Phase 4 onward, CI/manual gates also cover:

- app launch/navigation;
- loading/empty/error states;
- view-model behavior independent of Outlook COM;
- WebView2 ↔ host message contract;
- keyboard navigation and basic accessibility;
- light/dark and high-DPI smoke tests;
- write previews cannot be bypassed by UI routing/state restoration.

### Side-effect rule

Any test path that sends, saves, moves, creates, forwards, or deletes Outlook data must require explicit CLI/UI intent. Read-only inspection is the default.

## 10. CI/CD

Every pull request targeting `main` must pass the applicable .NET restore, formatter/analyzer, Release build, xUnit/MTP coverage, benchmark compile, NuGet audit, frontend install/typecheck/lint/format/build/npm audit, self-contained OutlookHost publish, and phase-specific published-EXE smoke tests.

From Phase 4 onward, the WinUI application must also build/package successfully in CI; a green OutlookHost alone is no longer sufficient.

Dependabot checks NuGet, npm, and GitHub Actions weekly. Dependency PRs are not auto-merged.

## 11. Implementation phases

### Phase 0 — Repository/toolchain bootstrap — **Complete ✅**

See [`docs/phase-0.md`](docs/phase-0.md).

### Phase 1 — Outlook COM discovery/read probe — **Complete ✅ / merged PR #4**

Delivered STA Classic Outlook activation/profile discovery, account/store/default Calendar enumeration, bounded recurrence-safe scanning, half-open overlap filtering, privacy-safe DTO output, deterministic COM release, self-contained artifact, and the interop packaging regression gate.

Real-machine testing found and fixed the single-file interop packaging issue. The repaired live probe worked and PR #4 was merged on 2026-09-08.

See [`docs/phase-1.md`](docs/phase-1.md).

### Phase 2 — Native meeting-forwarding technical spike — **In progress 🚧 / PR #5**

Goal: determine whether genuine Outlook meeting forwarding is reliable enough to ship.

Current direct-COM path:

- select source account explicitly by SMTP;
- select meeting by `GlobalAppointmentID`;
- search retained `IPM.Schedule.Meeting.Request` items;
- correlate using `GetAssociatedAppointment(false)`;
- Inspect is read-only;
- Prepare may call `MeetingItem.Forward()` and discard unsent;
- Send requires an exact confirmation token;
- never fall back to `ForwardAsVcal()` under the Forward label.

**Real-machine result #1:** an accepted Kverneland meeting with a valid calendar `GlobalAppointmentID` produced **zero retained matching `MeetingItem` objects in Inbox/Deleted Items**. Therefore retained-request recovery is already proven not to be universally available after acceptance.

Remaining Phase 2 investigation:

1. verify a positive control where a retained request is visibly present;
2. probe Outlook's built-in Calendar **Forward** command against the accepted `AppointmentItem` without executing it;
3. only if the command is valid/enabled, test an explicit prepare/discard path;
4. only after prepare succeeds, test one explicit send to an account controlled by the user;
5. classify Forward as reliable, conditional, or unsuitable.

See [`docs/phase-2.md`](docs/phase-2.md).

### Phase 3 — Identity model prototype

Implement GlobalAppointmentID correlation, StoreID+EntryID locators, Aligner metadata, origin/authority, and recurrence occurrence/exception identity. A moved event must remain the same logical event.

### Phase 4 — Read-only production UI

Turn the current compile-only `OutlookAligner.App` shell into the first real application:

- `App.xaml` / activation/lifetime;
- `MainWindow` with NavigationView;
- Calendar, Alignment, Settings, Diagnostics pages;
- MVVM view models and DI/hosting;
- IPC client to OutlookHost using plain DTOs;
- account cards/horizon controls/refresh;
- FullCalendar WebView2 integration;
- event selection + comparison pane;
- status filters and local authority selection;
- loading, empty, degraded, and error states;
- **no calendar writes yet**.

Phase 4 is not complete until the user can launch and operate the GUI without using the diagnostic CLI for normal read-only workflow.

### Phase 5 — Copy Full

Create safe Aligner-managed full copies with sync metadata and restart-safe correlation, surfaced through the GUI with confirmation.

### Phase 6 — Copy Busy

Create privacy-preserving managed busy placeholders without detail leakage, surfaced through the GUI.

### Phase 7 — Move Selected

Apply authority-driven Start/End updates only to selected non-authoritative managed copies with independent failure handling and operation history.

### Phase 8 — Move All

Build an immutable action plan, exclude unsafe/conflicted items, show GUI preview, require confirmation, execute independently, rescan, and summarize.

### Phase 9 — Recurring writes

Enable recurring series/exception copy and movement only after recurrence identity/read behavior is proven.

### Phase 10 — Hardening and release

Handle Outlook unavailable/busy/restart states, MAPI/profile/store failures, permissions, large calendars, all-day/timezone/DST/private items, malformed identities, crash recovery, database migration, application packaging, installer/update path, and user documentation.

## 12. V1 acceptance criteria

V1 is complete when:

1. a packaged **WinUI Outlook Aligner GUI** launches reliably on supported Windows + Classic Outlook;
2. the normal workflow can be completed without CLI commands;
3. configured accounts are discovered automatically and displayed in the GUI;
4. scan horizon is configurable and recurrence-bounded;
5. Calendar and Alignment views present the three accounts and logical event groups clearly;
6. logical meetings are grouped across accounts;
7. EntryID changes alone do not break identity;
8. Missing/Moved/Different/Conflict states are clear;
9. authority can be chosen and remembered;
10. Copy Full and Copy Busy work safely through explicit GUI actions;
11. Forward is offered only if Phase 2 proves genuine native forwarding for that event/situation;
12. Move Selected and Move All obey authority/safety rules and bulk preview requirements;
13. originals are not automatically deleted;
14. state survives restart;
15. one item failure does not abort the batch;
16. normal UI/logs remain privacy-safe;
17. Diagnostics provides useful health/error information without being required for normal use;
18. keyboard/high-DPI/light-dark smoke tests pass;
19. build/test/quality/security gates remain green.

## 13. Known risks

- Native MeetingItem recovery after accepted requests are deleted/archived.
- Whether Outlook's built-in Calendar Forward command can be safely invoked from an external automation host.
- True forwarding behavior under Exchange/Teams organizer policies such as Allow Forwarding.
- Outlook recurrence and moved exceptions.
- COM lifetime/rejected-call behavior.
- Cross-calendar privacy when making full copies.
- Timezone/DST semantics when moving/correlating appointments.
- WinUI/WebView2 lifecycle and IPC recovery when Outlook restarts.

## 14. Branch and merge policy

Phase 0 was the only direct-to-main exception.

Phase 1 and all later implementation work is PR-only. Never merge a PR unless the user explicitly instructs it after the relevant hosted/manual gates have been reviewed.

## 15. Next action

Finish Phase 2 on PR #5 by testing both retained-request native forwarding and the accepted-AppointmentItem Calendar Forward command path. Do not merge until the real Outlook results support a truthful product decision. Then Phase 3 establishes identity contracts needed by the explicitly required Phase 4 GUI.