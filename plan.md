# Outlook Aligner — Implementation Plan

Status: **Phase 0 complete ✅. Phase 1 is in progress 🚧 on PR #4.**

Last reviewed: 2026-09-07

## 1. Goal

Build a Windows desktop application that uses the Classic Outlook COM/Object Model as its only Outlook integration boundary. It discovers the calendar-capable accounts already configured in one Outlook profile, reads their calendars for a configurable future horizon, correlates the same logical meeting across accounts, shows differences, and lets the user deliberately align them.

Microsoft Graph is explicitly out of scope. The design must not require Azure app registration, Graph delegated permissions, tenant administrator consent, or separate Outlook credentials.

## 2. Confirmed product requirements

### Outlook environment

- Windows desktop application.
- Classic Outlook for Windows is required.
- The three target accounts are configured in the same Outlook profile.
- Accounts/calendars are discovered automatically through Outlook COM.
- New Outlook is out of scope because it does not expose the Classic Outlook Object Model automation surface required by this project.
- Outlook Aligner stores no Outlook passwords or access tokens.

### Calendar scan

- Scan starts at today.
- Future horizon is configurable; default is 90 days.
- Recurrences must be expanded only inside the bounded horizon.
- Refresh is user-driven in v1.
- One failed calendar/item must not abort the entire scan.

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

A real `Forward meeting` action may be exposed only if the Phase 2 technical spike proves that Classic Outlook COM can reliably reproduce Outlook's genuine forwarding behavior for accepted meetings. `AppointmentItem.ForwardAsVcal()` is not equivalent and must never be silently presented as true forwarding.

#### Copy full

Create an Outlook Aligner-managed calendar copy containing the available full details required by the product: subject, time, all-day state, location, body, online-meeting link where available, reminder settings, busy state, sensitivity, and supported recurrence data.

#### Copy busy

Create a privacy-preserving managed placeholder containing time/all-day/busy state and optionally a generic `Busy` subject. By default it must not copy body, attendees, Teams link, or location.

### Move actions and authority

Each logical event has an authoritative account. Normal rule: `AuthorityAccount = OriginAccount`.

Authority reason is recorded as:

- `KnownOrigin`
- `UserSelected`
- `Inferred`
- `Unknown`

`Move selected` updates only non-authoritative managed/local copies to the authoritative Start/End. The authoritative original is untouched.

`Move all` builds a preview first and excludes conflicts, unknown authority, ignored events, unsupported recurrence mutations, and failed safety checks.

### Deletion

Deletion synchronization is explicitly excluded from v1. Missing data is never interpreted as permission to delete another calendar's item.

## 3. Event identity model

### EntryID is a locator, not identity

Store `StoreID + EntryID` only as a current Outlook locator cache. EntryID may change after moves or other Outlook operations.

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
- WebView2 + FullCalendar Vanilla for rich calendar visualization.
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
- Explicitly release short-lived COM references, especially recurrence objects.
- Avoid COM `foreach` patterns where they hide enumerator RCWs.
- Never use `EntryID` as cross-account logical identity.
- Do not call `Outlook.Application.Quit()` merely because the probe started/attached to Outlook.
- Normal diagnostics contain HRESULT/operation context, not meeting subject/body/attendees/Teams URLs.

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

## 7. Verified stable implementation baseline

Verification date: **2026-09-07**. `✅` means rechecked against the latest non-preview/non-RC stable release. No Phase 0 dependency upgrade was required.

### Platform/toolchain

| Component | Version | Latest stable verified |
| --- | ---: | :---: |
| .NET SDK | 10.0.400 | ✅ |
| .NET Runtime/Desktop Runtime | 10.0.11 | ✅ |
| C# | 14.0 | ✅ |
| Visual Studio 2026 (recommended IDE) | 18.9.2 | ✅ |
| Node.js LTS | 24.20.0 | ✅ |
| npm with Node 24.20.0 | 11.19.0 | ✅ |

Node 24 is intentionally the latest **LTS** line used by the repository; Node 26 is the newer Current line.

### NuGet application packages

| Package | Version | Verified |
| --- | ---: | :---: |
| Microsoft.WindowsAppSDK | 2.4.0 | ✅ |
| CommunityToolkit.Mvvm | 8.4.2 | ✅ |
| Microsoft.Web.WebView2 | 1.0.4191.47 | ✅ |
| Microsoft.Office.Interop.Outlook | 15.0.4797.1004 | ✅ |
| Microsoft.Data.Sqlite | 10.0.11 | ✅ |
| Dapper | 2.1.79 | ✅ |
| Microsoft.Extensions.Hosting | 10.0.11 | ✅ |
| Microsoft.Extensions.Configuration.Json | 10.0.11 | ✅ |
| Microsoft.Extensions.Logging | 10.0.11 | ✅ |
| Serilog | 4.4.0 | ✅ |
| Serilog.Extensions.Hosting | 10.0.0 | ✅ |
| Serilog.Settings.Configuration | 10.0.1 | ✅ |
| Serilog.Sinks.File | 7.0.0 | ✅ |

The Outlook interop package has an old-looking version because the PIA package version does not track current Outlook product builds; `15.0.4797.1004` remains the current stable package.

### Test/quality packages

| Package | Version | Verified |
| --- | ---: | :---: |
| xunit.v3 | 4.0.0 | ✅ |
| Microsoft.Testing.Extensions.CodeCoverage | 18.11.0 | ✅ |
| NSubstitute | 6.2.0 | ✅ |
| BenchmarkDotNet | 0.15.8 | ✅ |

The repository does **not** use the earlier planning-only `Microsoft.NET.Test.Sdk` + `coverlet.collector` combination.

### Frontend

| Package/tool | Version | Verified |
| --- | ---: | :---: |
| fullcalendar | 7.1.0 | ✅ |
| temporal-polyfill | 1.0.4 | ✅ |
| TypeScript | 7.0.2 | ✅ |
| Vite | 8.2.2 | ✅ |
| ESLint | 10.10.0 | ✅ |
| Prettier | 3.9.6 | ✅ |

The current stable `typescript-eslint` line does not support TypeScript 7, so Phase 0 intentionally does not force an unsupported peer dependency. Strict `tsc --noEmit` is the TypeScript static-analysis gate until compatible stable tooling is available.

### GitHub Actions

| Action | Version | Verified |
| --- | ---: | :---: |
| actions/checkout | 7.0.1 | ✅ |
| actions/setup-dotnet | 6.0.0 | ✅ |
| actions/setup-node | 7.0.0 | ✅ |

Phase 1 adds `actions/upload-artifact@7.0.1`, also verified latest stable before use.

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
- ESLint for currently supported sources;
- Prettier;
- Vite production build.

Direct npm versions are exact, but Phase 0 does not fabricate a lockfile. CI currently uses `npm install --ignore-scripts`. The first PR that materially changes `web/calendar` must generate a real `package-lock.json` and move CI/local verification to `npm ci`.

### Logging

Normal logs may include counts, sync IDs, operation type, HRESULT/error category, and timings. They must not include subject, body, attendees, Teams links, or meeting location by default.

## 9. Testing strategy

### Hosted CI

Pure Core/Persistence/contract tests run without Outlook. Windows CI also builds the OutlookHost but cannot perform live Outlook integration because hosted runners do not have the user's Outlook profile.

### Manual Outlook integration

A downloadable diagnostic executable is produced from successful Phase 1 CI onward. Manual scenarios include account discovery, bounded recurrence scans, Teams meetings, all-day events, DST boundaries, Outlook cold start/restart, unavailable stores, and repeated scans. Passing the real three-account Classic Outlook suite is a Phase 1 merge gate.

A dedicated self-hosted Outlook integration runner may be added later but is not required initially.

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
14. npm production vulnerability audit.

From Phase 1, successful CI additionally publishes a self-contained Windows x64 Outlook probe executable and uploads it as a GitHub Actions artifact for manual testing.

Dependabot checks NuGet, npm, and GitHub Actions weekly. Dependency PRs are not auto-merged.

## 11. Implementation phases

### Phase 0 — Repository/toolchain bootstrap — **Complete ✅**

Completed directly on `main` as the one-time bootstrap exception.

Delivered:

- solution/project structure;
- .NET 10.0.400 / C# 14 baseline;
- Central Package Management;
- formatting/analyzer policy;
- xUnit v3 + Microsoft Testing Platform + coverage;
- BenchmarkDotNet infrastructure;
- compile-only WinUI and OutlookHost boundaries;
- TypeScript/Vite/FullCalendar shell;
- local verification script;
- GitHub Actions CI;
- Dependabot;
- latest-stable version verification.

Acceptance evidence:

- [x] clean hosted checkout builds;
- [x] formatter/analyzer gates pass;
- [x] tests and coverage pass;
- [x] benchmark project compiles;
- [x] frontend typecheck/lint/format/build passes;
- [x] NuGet/npm audits pass;
- [x] GitHub Actions run `34101604197` is green;
- [x] direct dependencies and Actions reverified latest stable;
- [x] documentation records actual implementation versions.

**Branch policy:** Phase 1 and every later implementation phase is PR-only. Do not merge without explicit user instruction.

### Phase 1 — Outlook COM discovery/read probe — **In progress 🚧 / PR #4**

Implement only read-only behavior:

- activate/attach to `Outlook.Application` automation on an STA thread;
- access the current MAPI namespace/profile;
- enumerate accounts and stores;
- locate each account's default Calendar;
- read a configurable bounded date range (default 90 days);
- use half-open overlap semantics (`eventEnd > windowStart` and `eventStart < windowEnd`);
- expand recurrences only inside the bound;
- extract plain account/store/event DTOs;
- expose privacy-safe console/JSON diagnostics with defensive output redaction;
- release COM references deterministically;
- never save, send, forward, delete, or modify Outlook data.

Phase 1 build artifact:

- successful CI publishes `OutlookAligner.OutlookHost` for `win-x64` as a self-contained single-file executable;
- artifact name: `OutlookAligner-Phase1-Probe-win-x64`;
- latest stable `actions/upload-artifact@v7.0.1` is used;
- the artifact is intended for manual testing on a Windows machine with Classic Outlook and the three-account profile configured.

Acceptance:

- [ ] PR CI passes;
- [ ] downloadable executable artifact is produced;
- [ ] executable discovers the expected three accounts;
- [ ] default Calendar is located for each usable account;
- [ ] bounded calendar items can be read repeatedly;
- [ ] recurring events remain bounded to the requested horizon;
- [ ] half-open lower/upper window boundaries behave correctly;
- [ ] default console and JSON output do not disclose event subject/body/location;
- [ ] Outlook already-running and cold-start paths both work;
- [ ] no Outlook data is written;
- [ ] repeated manual scans do not destabilize Outlook;
- [ ] the full manual merge-gate checklist in `docs/phase-1.md` passes on the real Classic Outlook profile.

### Phase 2 — Forwarding technical spike

Test true native forwarding with a real received/accepted Teams meeting. Prove whether a corresponding `MeetingItem` can be recovered and `MeetingItem.Forward()` can reproduce native Outlook forwarding. If not reliable, do not expose a misleading `Forward meeting` action.

### Phase 3 — Identity model prototype

Implement GlobalAppointmentID correlation, StoreID+EntryID locators, Aligner metadata, origin/authority, and recurrence occurrence/exception identity. A moved event must remain the same logical event.

### Phase 4 — Read-only production UI

Implement account cards, horizon controls, synchronized Calendar view, Alignment view, filters, comparison pane, status calculation, and local authority selection. Still no calendar writes.

### Phase 5 — Copy full

Create safe Aligner-managed full copies with sync metadata and restart-safe correlation.

### Phase 6 — Copy busy

Create privacy-preserving managed busy placeholders without detail leakage.

### Phase 7 — Move selected

Apply authority-driven Start/End updates only to selected non-authoritative members with independent failure handling and operation history.

### Phase 8 — Move all

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
9. Forward is offered only if Phase 2 proves genuine native forwarding;
10. Move selected and Move All obey authority/safety rules;
11. originals are not automatically deleted;
12. state survives restart;
13. one item failure does not abort the batch;
14. normal logs remain privacy-safe;
15. build/test/quality/security gates remain green.

## 13. Known risks

- True COM meeting forwarding for accepted meetings.
- Outlook recurrence and moved exceptions.
- COM lifetime/rejected-call behavior.
- Exchange/Teams tenant or organizer policies.
- Cross-calendar privacy when making full copies.

## 14. Reference sources

Primary sources are Microsoft Learn, .NET release/download pages, NuGet package pages, npm package pages, Node release information, and the official GitHub Action repositories. Stable versions were rechecked on 2026-09-07 before closing Phase 0.

## 15. Next action

Finish **Phase 1 on PR #4**: get hosted CI fully green, produce the downloadable read-only probe executable, run the documented merge-gate suite against the user's real three-account Classic Outlook profile, and merge only when the user explicitly requests it. Phase 2 follows only after Phase 1 is tested and merged.
