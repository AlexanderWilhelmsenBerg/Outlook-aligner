# Phase 0 — Repository and Toolchain Bootstrap

Status: **validation pending**

Date: 2026-09-07

## Purpose

Phase 0 creates a reproducible engineering baseline before Outlook behavior is implemented. It deliberately contains no calendar reads, writes, forwarding, matching, or synchronization logic.

This is the only implementation phase authorized to be committed directly to `main`. Phase 1 and every later phase must be developed on a branch and delivered through a pull request.

## Scope

Phase 0 establishes:

- the .NET solution and project boundaries described in `plan.md`;
- .NET 10 SDK pinning and C# 14 defaults;
- central NuGet package version management;
- formatting and analyzer policy;
- xUnit v3 using Microsoft Testing Platform;
- MTP-native code-coverage support;
- a compile-only Outlook host boundary with STA entry point and no COM behavior;
- a compile-only application shell with the chosen WinUI/WebView2 dependencies;
- the TypeScript/Vite/FullCalendar calendar frontend shell;
- BenchmarkDotNet infrastructure;
- local verification script;
- GitHub Actions CI;
- Dependabot for NuGet, npm, and GitHub Actions.

## Important Phase 0 decisions

### Microsoft Testing Platform

The repository uses xUnit v3 through Microsoft Testing Platform, selected in `global.json`. This replaces the older planned `Microsoft.NET.Test.Sdk` + `coverlet.collector` combination. Coverage uses `Microsoft.Testing.Extensions.CodeCoverage`.

### Outlook COM isolation

`OutlookAligner.OutlookHost` is an interactive Windows executable with an STA entry point, but Phase 0 does not instantiate `Outlook.Application` or access MAPI. That boundary is exercised for the first time in Phase 1.

### WinUI shell

`OutlookAligner.App` is intentionally a compile-only project shell in Phase 0. It references the selected Windows App SDK, CommunityToolkit, WebView2, hosting, and logging packages, but no XAML/application behavior is implemented. Production WinUI activation belongs to the read-only UI phase after the Outlook/identity spikes.

### Frontend lockfile

Direct npm dependencies are pinned to exact stable versions. A `package-lock.json` has intentionally not been fabricated because the current execution environment cannot resolve npm packages to generate a trustworthy lockfile. Bootstrap CI therefore uses `npm install` rather than `npm ci`.

The first pull request that materially touches `web/calendar` must generate and commit a real npm lockfile using Node 24, then change CI and `scripts/verify.ps1` to `npm ci`. Until then, top-level package versions are exact but transitive npm versions are not fully frozen.

## Repository map

```text
src/
  OutlookAligner.App/
  OutlookAligner.Core/
  OutlookAligner.Persistence/
  OutlookAligner.Outlook.Contracts/
  OutlookAligner.OutlookHost/
tests/
  OutlookAligner.Core.Tests/
  OutlookAligner.Persistence.Tests/
  OutlookAligner.OutlookHost.Tests/
benchmarks/
  OutlookAligner.Benchmarks/
web/calendar/
docs/
scripts/
.github/workflows/
```

## CI gates

Every push to `main` and every future pull request must pass:

1. .NET restore.
2. `dotnet format --verify-no-changes`.
3. Release build with warnings as errors.
4. xUnit v3 tests on Microsoft Testing Platform.
5. Coverage collection.
6. Benchmark project compilation.
7. NuGet vulnerability audit.
8. Frontend dependency install.
9. TypeScript typecheck.
10. ESLint.
11. Prettier check.
12. Vite production build.
13. Production npm vulnerability audit.

Benchmarks themselves run only through the manual `Benchmarks` workflow; performance numbers are not a noisy PR gate.

## Acceptance checklist

- [x] Solution/project structure created.
- [x] .NET 10.0.400 pinned.
- [x] Central package management configured.
- [x] Formatting/analyzer policy configured.
- [x] Test and benchmark projects created.
- [x] Frontend project created with exact direct dependency pins.
- [x] Local verification script created.
- [x] CI workflow created.
- [x] Dependabot configured.
- [ ] Hosted CI passes cleanly from a fresh checkout.
- [ ] `plan.md` updated to record Phase 0 completion and Phase 1 PR policy.

Phase 0 is not marked complete until the hosted CI result has been inspected and any bootstrap failures corrected.

## Phase 1 handoff

Phase 1 is a read-only Outlook COM discovery probe. It must be delivered by pull request and must not write or forward calendar data.
