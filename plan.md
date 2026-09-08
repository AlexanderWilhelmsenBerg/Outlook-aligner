# Outlook Aligner — Implementation Plan

Status: **Phase 0 complete ✅. Phase 1 complete ✅ and merged. Phase 2 native-forwarding spike in progress 🚧.**

Last reviewed: 2026-09-08

## 1. Goal

Build a Windows desktop application that uses the Classic Outlook COM/Object Model as its only Outlook integration boundary. It discovers the calendar-capable accounts already configured in one Outlook profile, reads their calendars for a configurable future horizon, correlates the same logical meeting across accounts, shows differences, and lets the user deliberately align them.

Microsoft Graph is explicitly out of scope. The design must not require Azure app registration, Graph delegated permissions, tenant administrator consent, or separate Outlook credentials.

## 2. Confirmed product requirements

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

Expose a real `Forward meeting` action only if Phase 2 proves genuine Classic Outlook forwarding for accepted meetings.

- Preferred mechanism: recover a native `MeetingItem` and call `MeetingItem.Forward()`.
- `AppointmentItem.ForwardAsVcal()` is not equivalent and must never be silently presented as native forwarding.
- If native recovery is only conditionally reliable, the product must expose Forward only when that capability is actually available.

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

Use Outlook `GlobalAppointmentID` as the primary native meeting-correlation candidate. Outlook documents it as the Global Object ID used to correlate meeting updates/responses and retained across copies.

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
  WinUI 3 / MVVM / alignment engine / SQLite / WebView2
        |
        | local IPC (planned)
        v
OutlookAligner.OutlookHost.exe
  interactive user process
  STA entry thread
  Classic Outlook COM/Object Model
```

The Outlook host is not a Windows service.

### COM boundary invariants

- COM objects never cross IPC or enter Core/UI DTOs.
- Convert Outlook objects immediately to plain DTOs.
- Keep Outlook automation on an STA thread.
- Explicitly release locally owned COM references.
- Avoid COM `foreach` patterns where hidden enumerator RCWs matter.
- Never use `EntryID` as cross-account logical identity.
- Do not call `Outlook.Application.Quit()` merely because OutlookHost activated Outlook.
- Normal diagnostics contain HRESULT/operation context, not meeting body/attendees/Teams URLs.

## 5. Persistence plan

SQLite tables planned for later phases:

- `Accounts`
- `SyncGroups`
- `EventMembers`
- `UserOverrides`
- `OperationHistory`
- `SchemaMigrations`

No credentials or cloud tokens are stored.

## 6. UI plan

Production UI begins after the COM/forwarding/identity spikes.

Primary views:

1. Calendar
2. Alignment
3. Settings
4. Diagnostics

Calendar view will show three synchronized account calendars. Alignment view will provide discrepancy filters, authority selection, and explicit single/bulk actions. Every bulk write operation requires a preview.

## 7. Verified implementation baseline

The detailed latest-stable version matrix is maintained in [`docs/phase-0.md`](docs/phase-0.md). Current baseline remains:

- .NET SDK 10.0.400 / runtime 10.0.11 / C# 14;
- Windows App SDK 2.4.0;
- WebView2 1.0.4191.47;
- Microsoft.Office.Interop.Outlook 15.0.4797.1004;
- Microsoft.Data.Sqlite 10.0.11 / Dapper 2.1.79;
- CommunityToolkit.Mvvm 8.4.2;
- Serilog stack as recorded in Phase 0;
- xUnit v3 4.0.0 / CodeCoverage 18.11.0 / NSubstitute 6.2.0;
- BenchmarkDotNet 0.15.8;
- Node.js 24.20.0 LTS / npm 11.19.0;
- TypeScript 7.0.2 / Vite 8.2.2 / FullCalendar 7.1.0;
- actions/checkout 7.0.1 / setup-dotnet 6.0.0 / setup-node 7.0.0 / upload-artifact 7.0.1.

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

Direct npm versions are exact. The first PR that materially changes `web/calendar` must generate a real `package-lock.json` under Node 24 and switch CI/local verification to `npm ci`.

### Privacy

Normal logs/output may include counts, sync IDs/global IDs, operation type, HRESULT/error category, and timings. They must not include subject, body, attendees, Teams links, or meeting location by default.

Diagnostic details are explicit opt-in.

## 9. Testing strategy

### Hosted CI

Pure Core/Persistence/contract tests run without Outlook. Windows CI builds/publishes OutlookHost but cannot perform live Outlook integration because hosted runners do not have the user's Outlook profile.

### Manual Outlook integration

A downloadable diagnostic executable is produced for Outlook-specific phases. Manual tests are merge gates whenever behavior depends on real Classic Outlook, Exchange/Teams state, or Outlook profile contents.

### Side-effect rule

Any test path that sends, saves, moves, creates, forwards, or deletes Outlook data must require explicit CLI/UI intent. Read-only inspection is the default.

## 10. CI/CD

Every pull request targeting `main` must pass:

1. checkout;
2. .NET restore;
3. formatter/analyzer verification;
4. Release build;
5. xUnit/MTP tests with coverage;
6. benchmark project compilation;
7. NuGet vulnerability audit;
8. Node 24 LTS setup;
9. frontend install;
10. TypeScript typecheck;
11. ESLint;
12. Prettier;
13. Vite production build;
14. npm production vulnerability audit;
15. self-contained Windows OutlookHost publish for Outlook-specific phases;
16. published-EXE `--help` and `--interop-check` smoke tests;
17. phase-specific no-Outlook CLI smoke tests before artifact upload.

Dependabot checks NuGet, npm, and GitHub Actions weekly. Dependency PRs are not auto-merged.

## 11. Implementation phases

### Phase 0 — Repository/toolchain bootstrap — **Complete ✅**

Completed directly on `main` as the one-time bootstrap exception.

See [`docs/phase-0.md`](docs/phase-0.md) for implementation and version-verification evidence.

### Phase 1 — Outlook COM discovery/read probe — **Complete ✅ / merged PR #4**

Delivered:

- STA Classic Outlook activation/profile discovery;
- account/store/default Calendar enumeration;
- configurable bounded calendar scan;
- recurrence-safe `Sort -> IncludeRecurrences -> Restrict -> GetFirst/GetNext` sequence;
- half-open overlap filter;
- privacy-safe console/JSON DTO output;
- deterministic COM release boundary;
- self-contained Windows artifact;
- embedded Outlook interop metadata + published-EXE `--interop-check` regression gate.

Real-machine testing found and fixed the interop packaging issue. The user confirmed a repaired live 14-day probe worked and merged PR #4 on 2026-09-08.

See [`docs/phase-1.md`](docs/phase-1.md).

### Phase 2 — Native meeting-forwarding technical spike — **In progress 🚧 / PR required**

Goal: determine whether genuine Outlook meeting forwarding is reliable enough to ship.

Implementation rules:

- select a source account explicitly by SMTP;
- select a real calendar meeting by `GlobalAppointmentID`;
- search the source account's retained `IPM.Schedule.Meeting.Request` items in Inbox and Deleted Items;
- call `GetAssociatedAppointment(false)` to correlate request -> appointment without adding calendar data;
- default Inspect mode is read-only;
- Prepare mode may invoke `MeetingItem.Forward()`, resolve one explicit recipient, set `SendUsingAccount`, then discard unsent;
- Send mode requires an exact confirmation token before calling `MeetingItem.Send()`;
- refuse zero/multiple/ambiguous native request matches;
- never fall back to `ForwardAsVcal()` under the Forward label;
- keep body/attendees/Teams URL out of diagnostic output.

Decision outcomes:

1. reliable native forwarding -> ship Forward later;
2. conditionally recoverable -> show Forward only when capability exists;
3. unreliable -> omit Forward and rely on Copy Full / Copy Busy later.

See [`docs/phase-2.md`](docs/phase-2.md) for the manual Teams/Outlook test matrix.

### Phase 3 — Identity model prototype

Implement GlobalAppointmentID correlation, StoreID+EntryID locators, Aligner metadata, origin/authority, and recurrence occurrence/exception identity. A moved event must remain the same logical event.

### Phase 4 — Read-only production UI

Implement account cards, horizon controls, synchronized Calendar view, Alignment view, filters, comparison pane, status calculation, and local authority selection. Still no calendar writes.

### Phase 5 — Copy Full

Create safe Aligner-managed full copies with sync metadata and restart-safe correlation.

### Phase 6 — Copy Busy

Create privacy-preserving managed busy placeholders without detail leakage.

### Phase 7 — Move Selected

Apply authority-driven Start/End updates only to selected non-authoritative managed copies with independent failure handling and operation history.

### Phase 8 — Move All

Build an immutable action plan, exclude unsafe/conflicted items, show preview, require confirmation, execute independently, rescan, and summarize.

### Phase 9 — Recurring writes

Enable recurring series/exception copy and movement only after recurrence identity/read behavior is proven.

### Phase 10 — Hardening and release

Handle Outlook unavailable/busy/restart states, MAPI/profile/store failures, permissions, large calendars, all-day/timezone/DST/private items, malformed identities, crash recovery, database migration, packaging, and manual documentation.

## 12. V1 acceptance criteria

V1 is complete when:

1. supported Windows + Classic Outlook launches reliably;
2. configured accounts are discovered automatically;
3. scan horizon is configurable and recurrence-bounded;
4. logical meetings are grouped across accounts;
5. EntryID changes alone do not break identity;
6. Missing/Moved/Different/Conflict states are clear;
7. authority can be chosen/remembered;
8. Copy Full and Copy Busy work safely;
9. Forward is offered only if Phase 2 proves genuine native forwarding for that situation;
10. Move Selected and Move All obey authority/safety rules;
11. originals are not automatically deleted;
12. state survives restart;
13. one item failure does not abort the batch;
14. normal logs remain privacy-safe;
15. build/test/quality/security gates remain green.

## 13. Known risks

- Native MeetingItem recovery after accepted requests are deleted/archived.
- True forwarding behavior under Exchange/Teams organizer policies.
- Outlook recurrence and moved exceptions.
- COM lifetime/rejected-call behavior.
- Cross-calendar privacy when making full copies.
- Timezone/DST semantics when moving/correlating appointments.

## 14. Branch and merge policy

Phase 0 was the only direct-to-main exception.

Phase 1 and all later implementation work is PR-only. Never merge a PR unless the user explicitly instructs it after the relevant hosted/manual gates have been reviewed.

## 15. Next action

Complete Phase 2 on a feature branch/PR, publish the forwarding-spike executable, and use a real accepted Teams meeting plus another account controlled by the user to decide whether native Outlook forwarding is reliable, conditional, or unsuitable for the product.
