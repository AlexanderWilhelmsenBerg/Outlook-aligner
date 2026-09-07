# Outlook Aligner — Implementation Plan

Status: planning only; no application code has been implemented yet.

Last reviewed: 2026-09-07

## 1. Goal

Build a Windows desktop application that uses the Classic Outlook COM/Object Model as its only Outlook integration boundary. The application will discover the three calendar-capable accounts already configured in the user's current Outlook profile, read their calendar events for a configurable future horizon, correlate the same logical meeting across accounts, show differences, and let the user deliberately align them.

The core user problem is not merely showing three calendars. The application must know that the same meeting exists in more than one account, detect when a copy is missing or has moved, remember which account is authoritative, and give the user safe single-event and bulk actions.

Microsoft Graph is explicitly out of scope. The design must not require Azure app registration, Graph delegated permissions, tenant administrator consent, or any other additional cloud API access.

## 2. Confirmed requirements

### 2.1 Outlook environment

- Windows desktop application.
- Classic Outlook for Windows is required.
- All three accounts are already configured inside the same Outlook profile.
- The application must discover and display the accounts automatically through Outlook COM.
- No separate account credentials are stored by Outlook Aligner.
- New Outlook support is out of scope because it does not expose the classic COM/Object Model automation surface required by this project.

### 2.2 Calendar scan

The user can configure how far into the future the program scans.

Initial defaults:

- Start: today.
- Future horizon: configurable number of days.
- Suggested default: 90 days.
- Recurring meetings must be expanded only inside that bounded horizon.
- Refresh can be initiated manually.
- Automatic refresh may be added later, but v1 is deliberately user-driven.

### 2.3 Three-account display

The main program must visibly show the three Outlook accounts and their calendar state.

Each account card should show at minimum:

- Account display name.
- SMTP/email address where available.
- Outlook store/calendar name.
- Enabled/disabled status for alignment.
- Connection/read status.

The UI must support both:

1. A visual three-calendar comparison view.
2. An alignment/discrepancy table for efficient review and bulk actions.

### 2.4 Event states

Each logical meeting/event is grouped across accounts and classified as one of:

- `Aligned` — present and matching in all selected accounts.
- `Missing` — absent from one or more selected accounts.
- `Moved` — same logical meeting but Start/End differs.
- `DetailsDifferent` — same meeting/time but selected metadata differs.
- `Duplicate` — more than one candidate copy exists in one account.
- `Conflict` — identity or authority cannot be resolved safely.
- `Ignored` — explicitly excluded by the user.
- `DeletedOrMissing` — reserved for later deletion support; v1 must not automatically delete anything.

### 2.5 Transfer actions

The application must distinguish three user intents.

#### A. Forward meeting

Goal: use Outlook's actual meeting-forwarding behavior where it can be driven reliably through the Outlook Object Model so the target account participates in the meeting rather than merely owning a visual copy.

This is especially important because the user wants:

- normal Outlook reminder behavior;
- Teams meeting notifications where the target account is genuinely associated with the meeting, including meeting-start/join-related behavior where Outlook/Teams/tenant policy supports it.

This requires a dedicated technical spike before the rest of the write path is considered complete.

Relevant COM capabilities:

- `MeetingItem.Forward()` performs Outlook's Forward action on a MeetingItem and returns the resulting MeetingItem.
- `AppointmentItem.ForwardAsVcal()` creates a MailItem with calendar information attached as a vCal. This is not equivalent to a true Outlook meeting forward and must not be presented as such.

Open technical question for the spike:

- Accepted calendar meetings are normally represented as `AppointmentItem` objects in Calendar, while the original meeting request is a `MeetingItem`. Outlook does not guarantee that the original request remains conveniently available after acceptance. We must validate a reliable COM-only route for forwarding an accepted calendar meeting.

Candidates to test, in this order:

1. Locate/retain the corresponding `MeetingItem` and use `MeetingItem.Forward()`.
2. Validate whether Outlook's native UI Forward command can be invoked safely through the Outlook Object Model/Office command surface for a selected AppointmentItem.
3. Use `ForwardAsVcal()` only as an explicitly labelled vCalendar-mail fallback, never as a silent substitute for true meeting forwarding.

Creating a brand-new meeting and making the user the organizer is not an acceptable substitute for forwarding the original meeting because it changes meeting identity and organizer semantics.

#### B. Copy full

Create an Outlook Aligner-managed calendar copy in the target calendar containing full available details, including:

- subject;
- start/end;
- all-day flag;
- location;
- body/description;
- Teams/online-meeting URL where present in the Outlook item;
- reminder settings where permitted;
- busy status;
- sensitivity/private status;
- recurrence information where supported.

This copy is not treated as proof that the target account is a real meeting attendee.

#### C. Copy as busy

Create a privacy-preserving Outlook Aligner-managed placeholder containing only the information necessary to block time.

Default busy-copy contents:

- start;
- end;
- all-day state;
- busy status;
- optional generic subject such as `Busy`.

Do not copy body, attendees, Teams link, location, or sensitive meeting details into a busy copy unless a later setting explicitly allows it.

### 2.6 Move actions

A logical event has an authoritative account.

The normal rule is:

- `AuthorityAccount = OriginAccount`.

When Outlook Aligner first sees an event in exactly one selected account and later creates/forwards copies to the other accounts, the originating account becomes authoritative automatically.

For pre-existing events already present in multiple accounts before Outlook Aligner has history, authority may be inferred, but inference must not be silently treated as certainty.

Authority metadata must track how the decision was made:

- `KnownOrigin` — created/forwarded through Outlook Aligner or otherwise established with high confidence.
- `UserSelected` — explicitly selected by the user.
- `Inferred` — best-effort inference awaiting confirmation if the event later conflicts.
- `Unknown` — no safe authority available.

The user can always select a logical event and choose a specific account as authoritative.

#### Move selected

When an event differs across accounts:

1. User selects the logical event.
2. User selects/accepts the authoritative account.
3. User presses Move.
4. Outlook Aligner updates the non-authoritative local copies to the authoritative Start/End.
5. The authoritative original is not modified by this operation.

#### Move all

After exceptions/conflicts have been handled manually, Move All aligns every non-conflicted event using each event's own authoritative account.

Move All must show a preview before writing any changes.

It must not operate on:

- unresolved conflicts;
- unknown-authority events;
- ignored events;
- unsupported recurrence mutations;
- items that fail a safety check.

### 2.7 Deletion

Deletion synchronization is planned but explicitly excluded from v1.

V1 rules:

- Never automatically delete an event.
- Never interpret a missing event as permission to delete another account's copy.
- A missing previously-known item may be shown as `DeletedOrMissing` for diagnosis.

A future deletion phase must introduce tombstones/history and a separate confirmation model before any destructive synchronization is implemented.

## 3. Event identity model

### 3.1 Do not use EntryID as logical identity

Outlook `EntryID` is a locator, not the cross-account identity key. It may change after moves between stores/folders and other Outlook operations.

Use it only together with the store as a current locator cache.

Stored locator fields:

- Outlook StoreID.
- Outlook EntryID.
- Last successful resolution timestamp.

### 3.2 Primary correlation key

Primary identity candidate:

- Outlook `GlobalAppointmentID`.

It is intended to correlate copies of the same meeting and is substantially more appropriate than EntryID for cross-calendar matching.

### 3.3 Outlook Aligner sync identity

For Aligner-created copies, add custom user properties so identity does not rely solely on Outlook heuristics.

Proposed properties:

- `OutlookAligner.SyncGroupId` — application-generated GUID.
- `OutlookAligner.SourceGlobalAppointmentId` — source GlobalAppointmentID where available.
- `OutlookAligner.SourceAccountId` — stable local account record ID.
- `OutlookAligner.CopyType` — `Full`, `Busy`, or other future type.
- `OutlookAligner.SchemaVersion` — custom-property schema version.

For a forwarded real meeting, do not assume these custom properties will propagate to the received meeting. Correlation must therefore still understand Outlook's native meeting identity.

### 3.4 Recurrence identity

Recurring items require separate treatment for:

- series master;
- normal occurrence;
- modified occurrence/exception;
- deleted occurrence.

A logical recurring occurrence must be identifiable using the series/global identity plus occurrence information, not only the visible current start time. This is essential for cases where one occurrence of a weekly meeting was moved to another date/time.

The recurrence identity model must be proven by tests before recurring write operations are enabled.

## 4. Architecture

### 4.1 Chosen stack

- Language: C# 14.
- Runtime: .NET 10 LTS.
- Native UI: WinUI 3 / Windows App SDK.
- Calendar visualization: WebView2 hosting FullCalendar Vanilla.
- Frontend calendar code: TypeScript.
- Database: SQLite through Microsoft.Data.Sqlite + Dapper.
- MVVM: CommunityToolkit.Mvvm.
- App hosting/configuration/DI: Microsoft.Extensions.Hosting and configuration packages.
- Logging: Microsoft.Extensions.Logging + Serilog.
- Unit tests: xUnit v3.
- Test doubles: NSubstitute only where a simple handwritten fake is not clearer.
- Coverage: coverlet.collector.
- Benchmarks: BenchmarkDotNet.
- CI: GitHub Actions on Windows.
- Dependency updates: Dependabot.

Microsoft Graph and MSAL are not dependencies.

### 4.2 Process model

Outlook COM automation should run in an interactive user process with STA semantics. It must not be implemented as a Windows service.

Recommended process split:

```text
OutlookAligner.App.exe
  WinUI 3
  MVVM
  Alignment engine
  SQLite
  WebView2 / FullCalendar
        |
        | named pipe / local IPC
        v
OutlookAligner.OutlookHost.exe
  hidden interactive user process
  STA entry thread
  Classic Outlook COM/Object Model
  no UI business logic
```

Benefits:

- COM runtime-callable wrappers never leak into the UI/domain layers.
- Outlook COM failures can be isolated from the main UI process.
- The core matching engine remains fully unit-testable without Outlook installed.
- The helper can be restarted if Outlook automation becomes unhealthy.
- A future non-COM backend could be added without rewriting the UI/domain model, although no Graph backend is planned at present.

### 4.3 COM boundary

All COM access lives under `OutlookAligner.Outlook` / `OutlookAligner.OutlookHost`.

Rules:

- Never pass COM objects over IPC.
- Never retain Outlook COM objects in view models or core models.
- Convert Outlook objects immediately into plain DTOs.
- Explicitly release short-lived COM references, particularly during recurrence enumeration.
- Avoid COM `foreach` patterns that hide enumerator RCWs when deterministic release is important.
- Reacquire recurring items before editing and release them immediately after save.
- Log COM error codes and operation context, but do not log meeting bodies/subjects by default.

### 4.4 Domain projects

Proposed solution layout:

```text
Outlook-aligner/
├── OutlookAligner.sln
├── global.json
├── Directory.Build.props
├── Directory.Packages.props
├── .editorconfig
├── .gitignore
├── README.md
├── plan.md
│
├── src/
│   ├── OutlookAligner.App/
│   │   ├── Views/
│   │   ├── ViewModels/
│   │   ├── Services/
│   │   └── Assets/
│   │
│   ├── OutlookAligner.Core/
│   │   ├── Models/
│   │   ├── Identity/
│   │   ├── Matching/
│   │   ├── Alignment/
│   │   └── Interfaces/
│   │
│   ├── OutlookAligner.Persistence/
│   │   ├── Database/
│   │   ├── Migrations/
│   │   └── Repositories/
│   │
│   ├── OutlookAligner.Outlook.Contracts/
│   │   └── Dtos/
│   │
│   └── OutlookAligner.OutlookHost/
│       ├── Com/
│       ├── Accounts/
│       ├── Calendars/
│       ├── Forwarding/
│       └── Ipc/
│
├── web/
│   └── calendar/
│       ├── src/
│       ├── package.json
│       ├── package-lock.json
│       ├── tsconfig.json
│       └── vite.config.ts
│
├── tests/
│   ├── OutlookAligner.Core.Tests/
│   ├── OutlookAligner.Persistence.Tests/
│   └── OutlookAligner.OutlookHost.Tests/
│
├── benchmarks/
│   └── OutlookAligner.Benchmarks/
│
├── docs/
│   ├── architecture.md
│   ├── event-identity.md
│   ├── forwarding-spike.md
│   └── manual-test-plan.md
│
└── .github/
    ├── workflows/
    │   ├── ci.yml
    │   └── benchmark.yml
    └── dependabot.yml
```

## 5. Database plan

Use SQLite as a local embedded database.

Do not store Outlook passwords, access tokens, or account credentials.

Proposed tables:

### Accounts

- Id.
- Outlook store identifier.
- SMTP/email where available.
- Display name.
- Calendar folder identifier.
- Enabled flag.
- LastSeenUtc.

### SyncGroups

One row per logical meeting/event.

- Id (GUID).
- GlobalAppointmentID if available.
- OriginAccountId.
- AuthorityAccountId.
- AuthorityReason.
- IsIgnored.
- FirstSeenUtc.
- LastSeenUtc.

### EventMembers

One row for each account-side member of a SyncGroup.

- SyncGroupId.
- AccountId.
- StoreID.
- EntryID.
- StartUtc/local-time metadata.
- EndUtc/local-time metadata.
- Recurrence identity.
- CopyType.
- IsAlignerManaged.
- LastModifiedUtc where available.
- LastSeenUtc.

### UserOverrides

- Authority overrides.
- Ignore decisions.
- Per-event copy policy where needed.

### OperationHistory

Auditable local history of actions without storing unnecessary sensitive content.

- Timestamp.
- Operation type.
- SyncGroupId.
- Source account.
- Target account.
- Outcome.
- Error category/code if failed.

### SchemaMigrations

Simple ordered application migrations.

Dapper is preferred over a full ORM because the schema is small, explicit, and relationship-oriented.

## 6. UI plan

### 6.1 Main navigation

Primary views:

1. `Calendar`
2. `Alignment`
3. `Settings`
4. `Diagnostics` (may be hidden behind an advanced toggle in v1)

### 6.2 Calendar view

Use three synchronized FullCalendar instances, one per selected Outlook account.

They should share:

- visible date;
- day/week/month view;
- scroll position where practical;
- working-hour range;
- zoom/time-slot settings.

A useful default desktop layout is three columns:

```text
+----------------+----------------+----------------+
| Account A      | Account B      | Account C      |
+----------------+----------------+----------------+
| 09:00 Project  | 09:00 Project  | 09:00 Project  |
| 10:30 Supplier |   MISSING      | 10:30 Supplier |
| 13:00 Workshop | 13:00 Workshop | 14:00 Workshop |
+----------------+----------------+----------------+
```

Selecting a logical event opens a native WinUI comparison pane showing each account's copy side-by-side and the available actions.

### 6.3 Alignment view

Optimized for discrepancy handling rather than visual scheduling.

Suggested columns:

- selection checkbox;
- subject/label;
- date/time;
- Account A state/time;
- Account B state/time;
- Account C state/time;
- authority;
- status;
- action menu.

Filters:

- All.
- Missing.
- Moved.
- Different.
- Conflict.
- Ignored.

Bulk actions:

- Forward selected missing meetings.
- Copy selected full.
- Copy selected as busy.
- Move selected.
- Move all safe events.

Every bulk write action requires a preview dialog summarizing exactly what will change.

### 6.4 Settings

Initial settings:

- Scan horizon in days.
- Enabled accounts.
- Default missing-event action: none / ask every time.
- Busy-copy subject style.
- Reminder policy for Aligner-managed copies.
- Default calendar view.
- Log level.
- Privacy-safe logging always on by default.

No automatic destructive-sync setting exists in v1.

## 7. Latest stable tool and package baseline

Versions below were verified as stable on 2026-09-07. Previews, RCs, experimental builds, dev builds, and .NET 11 previews are intentionally excluded.

### 7.1 Platform/runtime

| Component | Stable version | Planned use |
|---|---:|---|
| .NET SDK | 10.0.400 | pinned by `global.json` |
| .NET Runtime/Desktop Runtime | 10.0.11 | runtime baseline |
| C# | 14.0 | language version |
| Visual Studio | 2026 18.9 | recommended Windows IDE |
| Node.js LTS | 24.20.0 | calendar frontend build toolchain |

### 7.2 .NET application packages

| Package | Stable version | Planned use |
|---|---:|---|
| Microsoft.WindowsAppSDK | 2.4.0 | WinUI 3 / Windows App SDK |
| CommunityToolkit.Mvvm | 8.4.2 | MVVM source generators/helpers |
| Microsoft.Web.WebView2 | 1.0.4191.47 | embedded calendar visualization |
| Microsoft.Office.Interop.Outlook | 15.0.4797.1004 | Outlook PIA/COM compile-time interop |
| Microsoft.Data.Sqlite | 10.0.11 | SQLite provider |
| Dapper | 2.1.79 | lightweight DB mapping |
| Microsoft.Extensions.Hosting | 10.0.11 | DI, lifecycle, configuration host |
| Microsoft.Extensions.Configuration.Json | 10.0.11 | JSON configuration |
| Microsoft.Extensions.Logging | 10.0.11 | logging abstraction |
| Serilog | 4.4.0 | structured logging implementation |
| Serilog.Extensions.Hosting | 10.0.0 | generic-host integration |
| Serilog.Settings.Configuration | 10.0.1 | logger configuration through appsettings |
| Serilog.Sinks.File | 7.0.0 | rolling local diagnostic log |

Notes:

- `Microsoft.Office.Interop.Outlook` has an old-looking version because the Outlook PIA package version has not tracked modern Outlook release numbering. It is still the current stable NuGet package.
- Runtime behavior must always be tested against the installed Classic Outlook version; the PIA package is not the Outlook application itself.
- Use NuGet Central Package Management (`Directory.Packages.props`) so versions are defined once.

### 7.3 .NET test/quality packages

| Package | Stable version | Planned use |
|---|---:|---|
| xunit.v3 | 4.0.0 | unit/integration test framework |
| Microsoft.NET.Test.Sdk | 18.9.0 | test platform integration |
| NSubstitute | 6.2.0 | focused mocking only where useful |
| coverlet.collector | 10.0.1 | code coverage |
| BenchmarkDotNet | 0.15.8 | alignment-engine benchmarks |

Do not add FluentAssertions by default. xUnit assertions are sufficient and avoid introducing a dependency whose licensing model is unnecessary for this project.

### 7.4 Calendar/web packages

FullCalendar v7 moved the Vanilla JavaScript calendar to the `fullcalendar` package. Use the v7 package rather than assembling the older v6 `@fullcalendar/*` plugin package pattern.

| Package/tool | Stable version | Planned use |
|---|---:|---|
| fullcalendar | 7.1.0 | Vanilla calendar UI including day/time/list plugins |
| temporal-polyfill | 1.0.4 | Temporal support required/recommended by current FullCalendar package |
| TypeScript | 7.0.2 | frontend source language |
| Vite | 8.2.2 | frontend build/bundling |
| ESLint | 10.10.0 | TypeScript/JavaScript linting baseline; verify package ecosystem compatibility during bootstrap |
| Prettier | 3.9.6 | frontend formatting |

Before implementation begins, bootstrap must run a clean install and resolve any TypeScript-eslint compatibility needed for TypeScript 7/ESLint 10. If the relevant TypeScript ESLint packages have not yet released compatible stable versions, prefer a temporary reduced ESLint rule set over using prerelease tooling.

## 8. Formatting, linting, and code-quality rules

### 8.1 C# formatting

Repository-wide `.editorconfig` is authoritative.

Recommended project defaults:

```xml
<Nullable>enable</Nullable>
<ImplicitUsings>enable</ImplicitUsings>
<AnalysisLevel>latest-Recommended</AnalysisLevel>
<TreatWarningsAsErrors>true</TreatWarningsAsErrors>
```

Use:

- file-scoped namespaces;
- nullable reference types;
- Roslyn analyzers built into the .NET SDK;
- `dotnet format` for formatting/analyzer verification;
- no disabled warning without a documented reason.

CI gate:

```text
dotnet format --verify-no-changes
```

### 8.2 TypeScript formatting/linting

Use:

- TypeScript strict mode.
- ESLint stable configuration compatible with the chosen TypeScript version.
- Prettier.
- No generated/bundled files manually edited.

Suggested scripts:

```text
npm run typecheck
npm run lint
npm run format:check
npm run build
```

## 9. Logging and privacy

Default logs must be useful without leaking calendar contents.

Allowed normal log examples:

- account count;
- number of occurrences scanned;
- number of logical groups;
- number of missing/moved/conflicted items;
- SyncGroup GUID;
- operation type;
- Outlook/COM HRESULT/error category;
- timings.

Do not log by default:

- subject;
- body;
- attendees;
- Teams URLs;
- meeting location;
- email content.

If a future diagnostic mode allows PII logging, it must be explicit, temporary, and visibly indicated in the UI.

Rolling log retention should be bounded.

## 10. Testing strategy

### 10.1 Pure core tests

These run on every CI build and do not require Outlook.

Test:

- event correlation;
- GlobalAppointmentID grouping;
- Aligner SyncGroup metadata grouping;
- authoritative-account decisions;
- missing detection;
- move detection;
- detail-difference detection;
- conflict detection;
- copy/move action planning;
- Move All safety exclusion;
- recurrence occurrence identity;
- time-zone and DST comparisons.

### 10.2 Persistence tests

Run against temporary SQLite databases.

Test:

- migrations;
- account upsert;
- SyncGroup persistence;
- history;
- locator replacement when EntryID changes;
- restart/reload behavior;
- schema upgrade.

### 10.3 Outlook integration tests

Outlook COM tests cannot be meaningfully executed on ordinary GitHub-hosted runners because they lack the user's Outlook profile/accounts.

Create a manual/optional integration harness for a real Windows machine with Classic Outlook configured.

Required scenarios:

1. Discover all three configured accounts.
2. Read default calendar from each store.
3. Bounded scan with recurrence enabled.
4. One-time external meeting.
5. User-created appointment.
6. Teams meeting.
7. Copy full.
8. Copy busy.
9. Meeting forwarding spike.
10. Meeting moved by organizer.
11. Local copy moved by user.
12. Recurring series.
13. Modified recurrence occurrence.
14. Deleted recurrence occurrence.
15. All-day event.
16. DST boundary.
17. Outlook restart while Aligner is open.
18. One account/store temporarily unavailable.
19. EntryID changes while GlobalAppointmentID remains stable.

A self-hosted GitHub Actions runner with a dedicated test Outlook profile may be added later, but it is not required for the initial project.

## 11. Benchmarking

Do not benchmark COM performance as a PR quality gate. Outlook state, cache state, profile/network condition, and machine variance make that unsuitable for deterministic CI thresholds.

Benchmark the pure alignment engine using generated datasets:

- 100 events;
- 1,000 events;
- 10,000 events;
- 50,000 expanded occurrences.

Benchmark:

- identity matching;
- grouping;
- difference calculation;
- action-plan generation;
- SQLite lookup/load;
- Move All planning.

BenchmarkDotNet benchmark projects must compile in CI.

Actual benchmark runs should be manual or scheduled and used to detect major regressions, not small noisy percentage changes.

## 12. CI/CD plan

### 12.1 Pull-request/main CI

GitHub Actions Windows runner.

Steps:

1. Checkout.
2. Install/pin .NET SDK from `global.json`.
3. Install Node 24 LTS.
4. `dotnet restore`.
5. `dotnet format --verify-no-changes`.
6. `dotnet build -c Release` with warnings as errors.
7. `dotnet test -c Release`.
8. Collect coverage.
9. `npm ci` in `web/calendar`.
10. TypeScript typecheck.
11. ESLint.
12. Prettier check.
13. Vite production build.
14. Ensure benchmark project compiles.
15. `dotnet list package --vulnerable --include-transitive` review/gate for actionable vulnerabilities.
16. `npm audit` for high/critical production-impacting issues, with exceptions documented rather than silently ignored.
17. Produce an unsigned development artifact after the application exists.

### 12.2 Dependency maintenance

Dependabot:

- NuGet — weekly.
- npm — weekly.
- GitHub Actions — weekly.

Do not auto-merge dependency PRs initially. Build/test first because Outlook/WinUI dependencies can have platform-specific regressions.

### 12.3 Releases

Initial release strategy:

- Windows x64 first.
- Self-contained publish or packaged WinUI deployment selected during bootstrap spike.
- MSIX is a later packaging option after normal development deployment works reliably.
- Release artifacts should never contain user database/log files.

## 13. Implementation phases

### Phase 0 — Repository/toolchain bootstrap

Create:

- solution/project structure;
- `global.json` pinning .NET 10.0.400;
- central NuGet package management;
- `.editorconfig`;
- frontend package setup;
- format/lint/test scripts;
- GitHub Actions CI;
- Dependabot.

Acceptance:

- clean checkout builds without Outlook-specific runtime tests;
- formatter/linter gates pass;
- empty test suite executes;
- frontend builds into assets usable by WebView2.

### Phase 1 — Outlook COM discovery/read probe

Implement only enough to prove:

- attach/create `Outlook.Application` COM automation;
- access current MAPI namespace/profile;
- enumerate accounts/stores;
- locate each account's default Calendar;
- read a bounded date range;
- safely expand recurrences;
- extract identity/time/detail DTOs;
- release COM references safely.

No writes.

Acceptance:

- all three accounts appear in a diagnostic UI/harness;
- calendar items for the configured horizon can be read repeatedly without Outlook crashes/leaks;
- recurring events are bounded correctly.

### Phase 2 — Forwarding technical spike

This phase is deliberately early because true meeting forwarding is a core user requirement and COM behavior must be proven empirically.

Test with a real received/accepted Teams meeting.

Questions to answer:

1. Can the related MeetingItem be recovered reliably after acceptance?
2. Does `MeetingItem.Forward()` reproduce the same native Outlook forwarding behavior as using Forward in the Outlook UI?
3. If not, can Outlook's native Forward command be safely invoked for the selected Calendar AppointmentItem through supported Office/Outlook automation?
4. What happens to GlobalAppointmentID on the forwarded target account?
5. Does the target account receive a genuine meeting request/calendar item?
6. Does Outlook reminder behavior work normally?
7. Does Teams recognize the target account sufficiently to provide the meeting notification behavior the user expects?
8. How do Exchange/tenant policies affect forwarding?
9. What happens for recurring meetings?
10. What happens if meeting forwarding is disabled by the organizer or tenant?

Document results in `docs/forwarding-spike.md`.

Blocking rule:

- Do not label any action `Forward meeting` unless this spike proves it is a true native meeting-forward operation.
- If COM cannot perform it reliably, expose only the capabilities proven to work (`Copy full`, `Copy busy`, and explicitly-labelled vCal mail if useful) and document the limitation. Do not silently fake a forward.

### Phase 3 — Identity model prototype

Implement correlation for:

- GlobalAppointmentID;
- StoreID + EntryID locator;
- Aligner SyncGroup custom properties;
- origin/authority;
- recurring occurrences/exceptions.

Use a read-only comparison harness first.

Acceptance:

- moving an appointment does not cause a new unrelated logical event;
- an Aligner-created copy is re-associated after restart;
- recurring moved occurrence remains associated with its original occurrence.

### Phase 4 — Read-only production UI

Implement:

- account cards;
- horizon control;
- Calendar view;
- Alignment view;
- filters;
- event comparison drawer;
- status calculation;
- authority selection stored locally.

No calendar writes yet.

### Phase 5 — Copy full

Implement safe Aligner-managed copies with sync metadata.

Acceptance:

- copy appears in chosen target calendar;
- reminder/details copy as specified;
- source is unchanged;
- rescanning correlates source and target into one SyncGroup;
- restart preserves correlation.

### Phase 6 — Copy as busy

Implement privacy-preserving placeholders.

Acceptance:

- blocks correct time;
- no source body, Teams URL, attendee list, or location leaks by default;
- clearly identifiable internally as an Aligner-managed busy copy.

### Phase 7 — Move selected

Implement authority-driven Start/End update of non-authoritative members.

Acceptance:

- chosen authority is persisted;
- only target member(s) are modified;
- source/authority event is untouched;
- failed target updates do not corrupt other members;
- operation history records success/failure.

### Phase 8 — Move All

Implement safe bulk alignment.

Required flow:

1. Scan current state.
2. Build immutable action plan.
3. Exclude conflicts/unknowns/unsupported items.
4. Show preview counts and affected account directions.
5. User confirms.
6. Execute each operation independently.
7. Rescan.
8. Show result summary.

No silent automatic bulk writes.

### Phase 9 — Recurring writes

Enable copy/move of recurring series and individual exceptions only after recurrence identity/read behavior is proven.

### Phase 10 — Hardening and release

Handle:

- Outlook not installed/running;
- Outlook startup prompts;
- MAPI/profile unavailable;
- COM busy/rejected-call behavior;
- unavailable account/store;
- permission failures;
- Outlook restart;
- large calendars;
- all-day events;
- time zones/DST;
- private items;
- malformed/duplicate identities;
- crash recovery;
- database migration/backup.

Prepare release packaging and manual user documentation.

## 14. V1 acceptance criteria

V1 is complete when all of the following are true:

1. Application runs on supported Windows with Classic Outlook.
2. It automatically discovers and visibly displays the three accounts in the current Outlook profile.
3. User can configure the future scan horizon.
4. Scan is bounded and supports recurring appointments inside the horizon.
5. Same logical meeting is grouped across accounts instead of shown as unrelated copies.
6. EntryID changes do not by themselves break logical identity.
7. Missing events are clearly identified per account.
8. Moved events are clearly identified.
9. User can set/override the authoritative account for an event.
10. Known-origin events default to their origin account as authority.
11. Copy Full works and persists correlation.
12. Copy Busy works without leaking details by default.
13. True Forward is available only if the COM forwarding spike proves native forwarding reliably; otherwise the limitation is explicit.
14. Move selected aligns non-authoritative local copies to the authoritative Start/End.
15. Move All previews and applies only safe, non-conflicted changes.
16. Original authoritative items are never deleted.
17. No automatic deletion exists in v1.
18. Database/state survives application and Outlook restart.
19. One failed event does not abort the whole scan/bulk operation.
20. Normal logs do not contain sensitive calendar content.
21. Core/persistence tests run in CI.
22. Formatting, linting, build, tests, coverage collection, and frontend build are CI-gated.

## 15. Planned post-v1 features

- Deletion/tombstone synchronization with explicit safety rules.
- Optional background/periodic refresh while the interactive user session is active.
- Tray mode.
- More than three accounts.
- Account grouping/profiles.
- Per-account copy privacy policy.
- One-click `Align all missing` with policy preview.
- Rich operation history/undo where Outlook semantics make undo safe.
- Additional calendar views/search.
- Optional self-hosted Outlook integration CI.

## 16. Known risks

### True COM meeting forwarding

Highest technical risk. Outlook exposes `MeetingItem.Forward()`, but reliably obtaining/operating on the correct MeetingItem for an already-accepted Calendar AppointmentItem must be proven. This is why the forwarding spike happens before the write-heavy phases.

### Recurrence

Outlook recurrence is stateful and COM-reference-sensitive. A moved exception must not be mistaken for a new standalone event.

### COM lifecycle

Poor COM reference handling can leak Outlook objects or make automation unstable. Keep the COM surface tiny and isolated.

### Tenant/organizer policy

Even when native Outlook forwarding is technically possible, Exchange/Teams policy or organizer settings may reject forwarding. The UI must report policy rejection rather than trying to work around it.

### Calendar privacy

Copy Full can duplicate work/private content into another mailbox. Copy Busy must be prominently available and privacy-safe by default.

## 17. Source references used for this plan

Primary references checked on 2026-09-07:

- .NET 10 downloads/releases: https://dotnet.microsoft.com/en-us/download/dotnet/10.0
- Windows App SDK NuGet: https://www.nuget.org/packages/Microsoft.WindowsAppSDK
- CommunityToolkit.Mvvm NuGet: https://www.nuget.org/packages/CommunityToolkit.Mvvm
- WebView2 NuGet: https://www.nuget.org/packages/Microsoft.Web.WebView2
- Outlook Interop NuGet: https://www.nuget.org/packages/Microsoft.Office.Interop.Outlook
- Microsoft.Data.Sqlite NuGet: https://www.nuget.org/packages/Microsoft.Data.Sqlite
- Dapper NuGet: https://www.nuget.org/packages/Dapper
- xUnit v3 NuGet: https://www.nuget.org/packages/xunit.v3
- BenchmarkDotNet NuGet: https://www.nuget.org/packages/BenchmarkDotNet
- FullCalendar npm: https://www.npmjs.com/package/fullcalendar
- TypeScript npm: https://www.npmjs.com/package/typescript
- Vite npm: https://www.npmjs.com/package/vite
- ESLint npm: https://www.npmjs.com/package/eslint
- Prettier npm: https://www.npmjs.com/package/prettier
- Node.js release index: https://nodejs.org/en/blog/release
- Outlook MeetingItem.Forward: https://learn.microsoft.com/en-us/office/vba/api/outlook.meetingitem.forward%28method%29
- Outlook AppointmentItem.ForwardAsVcal: https://learn.microsoft.com/en-us/office/vba/api/outlook.appointmentitem.forwardasvcal

## 18. Next action

The next repository change should be **Phase 0 only**: bootstrap the solution/tooling/CI without implementing Outlook behavior. After that passes cleanly, implement **Phase 1 Outlook read probe**, followed immediately by the **Phase 2 COM forwarding spike** before building mutation features.
