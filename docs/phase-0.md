# Phase 0 — Repository and Toolchain Bootstrap

Status: **Complete ✅**

Completed: 2026-09-07

Verification date: 2026-09-07

## Purpose

Phase 0 establishes a reproducible engineering baseline before Outlook behavior is implemented. It deliberately contains no calendar writes, forwarding, matching, synchronization, or destructive behavior.

This is the only implementation phase authorized to be committed directly to `main`. Phase 1 and every later implementation phase must be developed on a branch and delivered through a pull request.

## Completion evidence

The final Phase 0 implementation baseline was validated by GitHub Actions run `34101604197` on Windows and Ubuntu hosted runners. Both jobs passed all configured gates: restore/install, formatting, analyzers, release build, tests with coverage, benchmark compilation, frontend typecheck/lint/format/build, and NuGet/npm vulnerability audits.

No direct dependency required an upgrade during final stable-version verification.

## Verified stable implementation baseline

A checkmark means the pinned version was rechecked against the latest non-preview/non-RC stable release on 2026-09-07.

### Platform and language

| Component | Repository baseline | Verification |
| --- | ---: | :---: |
| .NET SDK | 10.0.400 | ✅ Latest stable .NET 10 SDK |
| .NET Runtime/Desktop Runtime | 10.0.11 | ✅ Latest stable .NET 10 patch |
| C# | 14.0 | ✅ Current stable language for .NET 10 |
| Visual Studio 2026 | 18.9.2 | ✅ Latest stable IDE release (recommended, not required by CI) |
| Node.js | 24.20.0 LTS | ✅ Latest Node 24 LTS patch |
| npm | 11.19.0 | ✅ Version shipped with Node 24.20.0 |

Node 26 is the newer **Current** release line, but Outlook Aligner intentionally uses the latest Node 24 **LTS** line for build infrastructure.

### NuGet application packages

| Package | Pinned version | Verified latest stable |
| --- | ---: | :---: |
| `Microsoft.WindowsAppSDK` | 2.4.0 | ✅ |
| `CommunityToolkit.Mvvm` | 8.4.2 | ✅ |
| `Microsoft.Web.WebView2` | 1.0.4191.47 | ✅ |
| `Microsoft.Office.Interop.Outlook` | 15.0.4797.1004 | ✅ |
| `Microsoft.Data.Sqlite` | 10.0.11 | ✅ |
| `Dapper` | 2.1.79 | ✅ |
| `Microsoft.Extensions.Hosting` | 10.0.11 | ✅ |
| `Microsoft.Extensions.Configuration.Json` | 10.0.11 | ✅ |
| `Microsoft.Extensions.Logging` | 10.0.11 | ✅ |
| `Serilog` | 4.4.0 | ✅ |
| `Serilog.Extensions.Hosting` | 10.0.0 | ✅ |
| `Serilog.Settings.Configuration` | 10.0.1 | ✅ |
| `Serilog.Sinks.File` | 7.0.0 | ✅ |

`Microsoft.Office.Interop.Outlook` has an old-looking version because the PIA package version does not track modern Outlook application build numbers. `15.0.4797.1004` is still the current stable NuGet package. Runtime compatibility is validated against installed Classic Outlook separately.

### NuGet test and quality packages

| Package | Pinned version | Verified latest stable |
| --- | ---: | :---: |
| `xunit.v3` | 4.0.0 | ✅ |
| `Microsoft.Testing.Extensions.CodeCoverage` | 18.11.0 | ✅ |
| `NSubstitute` | 6.2.0 | ✅ |
| `BenchmarkDotNet` | 0.15.8 | ✅ |

The repository uses xUnit v3 through **Microsoft Testing Platform (MTP)**. The earlier planning-only `Microsoft.NET.Test.Sdk` + `coverlet.collector` combination is not used.

### Calendar/frontend packages

| Package/tool | Pinned version | Verified latest stable |
| --- | ---: | :---: |
| `fullcalendar` | 7.1.0 | ✅ |
| `temporal-polyfill` | 1.0.4 | ✅ |
| TypeScript | 7.0.2 | ✅ |
| Vite | 8.2.2 | ✅ |
| ESLint | 10.10.0 | ✅ |
| Prettier | 3.9.6 | ✅ |

FullCalendar 7 uses the `fullcalendar` Vanilla package rather than the older v6 `@fullcalendar/*` package layout.

TypeScript 7 is intentionally gated by strict `tsc --noEmit`. The stable `typescript-eslint` line available during Phase 0 does not yet support TypeScript 7, so the repository does not force an unsupported peer dependency. ESLint currently gates the JavaScript configuration layer. TypeScript-aware ESLint should be restored when a compatible stable parser is available.

### GitHub Actions

| Action | Pinned version | Verified latest stable |
| --- | ---: | :---: |
| `actions/checkout` | 7.0.1 | ✅ |
| `actions/setup-dotnet` | 6.0.0 | ✅ |
| `actions/setup-node` | 7.0.0 | ✅ |

Phase 1 adds `actions/upload-artifact` for the test executable; its version is verified separately in the Phase 1 PR.

## Repository baseline

Phase 0 created:

- `.NET 10` solution and project boundaries;
- `global.json` SDK pinning and Microsoft Testing Platform selection;
- NuGet Central Package Management;
- C# 14, nullable reference types, latest recommended analyzers, warnings-as-errors, and deterministic builds;
- `.editorconfig`, `.gitattributes`, and repository formatting rules;
- xUnit v3 test projects with MTP-native coverage;
- BenchmarkDotNet harness;
- compile-only OutlookHost STA process boundary with no COM behavior in Phase 0;
- compile-only WinUI/WebView2 application shell;
- TypeScript/Vite/FullCalendar frontend shell;
- local verification script;
- GitHub Actions CI;
- Dependabot for NuGet, npm, and GitHub Actions.

## Frontend lockfile note

Direct npm dependencies are pinned to exact versions. Phase 0 intentionally does not contain a fabricated `package-lock.json`; CI therefore uses `npm install --ignore-scripts` rather than `npm ci`.

The first PR that materially changes `web/calendar` must generate a real lockfile using the pinned Node 24 LTS toolchain and switch frontend verification to `npm ci`. Phase 1 does not modify the frontend.

## CI gates

Every push to `main` and every pull request targeting `main` must pass:

1. .NET restore.
2. `dotnet format --verify-no-changes`.
3. Release build with warnings as errors.
4. xUnit v3 tests through Microsoft Testing Platform.
5. MTP code-coverage collection.
6. Benchmark project compilation.
7. NuGet transitive vulnerability audit.
8. Frontend dependency installation.
9. TypeScript strict typecheck.
10. ESLint for currently supported sources.
11. Prettier check.
12. Vite production build.
13. Production npm vulnerability audit.

## Acceptance checklist

- [x] Solution/project structure created.
- [x] .NET SDK 10.0.400 pinned.
- [x] C# 14 and analyzer policy configured.
- [x] Central package management configured.
- [x] Formatting and line-ending policy configured.
- [x] xUnit v3 / Microsoft Testing Platform test baseline created.
- [x] MTP code coverage configured.
- [x] Benchmark infrastructure created.
- [x] Frontend project created with exact direct dependency pins.
- [x] Local verification script created.
- [x] CI workflow created.
- [x] Dependabot configured.
- [x] Every direct implementation dependency reverified against latest stable releases.
- [x] GitHub Actions dependencies reverified against latest stable releases.
- [x] Hosted CI passes cleanly from a fresh checkout (`34101604197`).
- [x] `plan.md` updated to record Phase 0 completion and the PR-only policy from Phase 1 onward.

## Phase 1 handoff

Phase 1 is the first PR-only implementation phase. It introduces a **read-only Classic Outlook COM discovery/calendar probe** and a CI-produced self-contained Windows x64 test executable. It must not save, modify, forward, delete, or otherwise write calendar data.
