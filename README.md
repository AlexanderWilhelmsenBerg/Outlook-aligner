# Outlook Aligner

Outlook Aligner is a **WinUI 3 Windows desktop application** for comparing and deliberately aligning calendar events across multiple accounts in one Classic Outlook profile.

The command-line tools used in the early technical phases are diagnostic/test harnesses only. Normal product use is moving into the WinUI application shell; see [`docs/ui.md`](docs/ui.md).

## Integration boundary

The project uses the Classic Outlook COM/Object Model. Microsoft Graph and MSAL are deliberately out of scope; no Azure app registration, Graph permission, or tenant administrator consent is required by the design.

## Project status

- **Phase 0 — Repository/toolchain bootstrap: Complete ✅**
- **Phase 1 — Outlook COM discovery/read probe: Complete ✅ / merged in PR #4**
- **Phase 2 — Native meeting forwarding: Complete ✅ / merged in PR #5**
- **Phase 3 — WinUI read/alignment foundation: In progress 🚧**

Phase 0 was the one-time bootstrap authorized directly on `main`. Every implementation change from Phase 1 onward is delivered through a pull request and is not merged automatically.

Phase 1 proved the packaged Classic Outlook COM read boundary on the user's real Windows/Outlook profile, including the self-contained interop packaging fix discovered during manual testing. See [`docs/phase-1.md`](docs/phase-1.md).

Phase 2 proved genuine Classic Outlook native forwarding for accepted Calendar meetings through Outlook's built-in Forward command and was merged in PR #5. The product does not silently substitute `ForwardAsVcal()`/ICS. See [`docs/phase-2.md`](docs/phase-2.md).

## Current Phase 3 increment

The current `phase-3/ui-assisted-testing` branch intentionally stops at a **read-only alignment foundation plus safe Forward preparation** rather than trying to finish the entire phase in one PR.

It currently provides:

- a real unpackaged WinUI 3 `App.xaml` / `MainWindow` application;
- NavigationView destinations for Calendar, Alignment, Settings and Diagnostics;
- automatic OutlookHost launch and protocol-version validation;
- bounded account/calendar discovery from the GUI;
- selected-event native Forward capability checks;
- **Prepare Forward (discard)** using the proven native Outlook forwarding path, with no send action in this increment;
- preliminary logical-event correlation for non-recurring events;
- explicit `Aligned`, `Missing`, `Moved`, `DetailsDifferent`, `Duplicate`, `Conflict`, recurrence-unresolved and uncorrelated states;
- read-only Outlook Aligner managed-copy metadata detection;
- fail-closed handling of incomplete, unsupported or unreadable managed-copy metadata;
- `KnownOrigin` authority inference only from internally consistent validated managed-copy provenance;
- session-only user authority selection when origin is not known;
- a **preview-only Move Selected plan** that can target validated managed copies but cannot execute Outlook writes;
- a self-contained Windows CI artifact containing the UI and OutlookHost.

Recurring events remain deliberately unresolved for identity/correlation. Copy writes, Move writes, real Forward send from the GUI, persistence, operation history and production IPC are not part of this increment.

## Production UI direction

The v1 GUI uses a WinUI 3 `NavigationView` shell with:

- **Calendar** — primary multi-account calendar and event selection;
- **Alignment** — discrepancy queue, authority and safe previews/actions;
- **Settings** — horizon, privacy and UI preferences;
- **Diagnostics** — Outlook/host health, operation history and technical identifiers.

Normal v1 operation must not require PowerShell or CLI arguments. Write actions are capability-driven, and bulk writes require preview + explicit confirmation. The detailed UI/UX contract is in [`docs/ui.md`](docs/ui.md).

For **Forward meeting**, the production GUI treats native forwarding as a capability of the selected Calendar event: enable the action only when Classic Outlook reports the built-in Forward command as available and all identity/safety checks pass. Unsupported events fail closed with a clear explanation. The retained-request diagnostic path is not exposed as a separate user-facing action.

## Update the current test app

During development, use the updater instead of manually opening GitHub Actions and replacing extracted artifacts:

```powershell
./scripts/Update-OutlookAlignerTestApp.ps1
```

One-time prerequisite: install and authenticate GitHub CLI (`gh`). The updater finds the latest **successful push CI** build from `phase-3/ui-assisted-testing`, downloads `OutlookAligner-Phase3-UI-TestHarness-win-x64`, verifies that the UI and OutlookHost executables are present, and installs it under `%LOCALAPPDATA%\OutlookAligner\TestApp`. The previous build is retained as `%LOCALAPPDATA%\OutlookAligner\TestApp.previous`, and the updated app launches automatically.

Useful options:

```powershell
# Update without launching.
./scripts/Update-OutlookAlignerTestApp.ps1 -NoLaunch

# Redownload even if the latest CI run is already installed.
./scripts/Update-OutlookAlignerTestApp.ps1 -Force

# Also remove downloaded-file zone markers from the verified artifact files.
./scripts/Update-OutlookAlignerTestApp.ps1 -UnblockFiles
```

The persistent diagnostics log is stored separately under `%LOCALAPPDATA%\OutlookAligner\events.jsonl`, so updating the test app does not erase test history.

## Verified toolchain baseline

- .NET SDK 10.0.400 / .NET 10.0.11 / C# 14
- Windows App SDK 2.4.0 / WinUI 3
- WebView2 1.0.4191.47
- Outlook Interop 15.0.4797.1004
- Node.js 24.20.0 LTS / npm 11.19.0
- TypeScript 7.0.2 / Vite 8.2.2 / FullCalendar 7.1.0
- xUnit v3 4.0.0 on Microsoft Testing Platform with CodeCoverage 18.11.0
- BenchmarkDotNet 0.15.8
- GitHub Actions + Dependabot

## Local verification

Prerequisites:

- .NET SDK 10.0.400
- Node.js 24 LTS
- PowerShell 7 recommended

Run:

```powershell
./scripts/verify.ps1
```

Classic Outlook is required for live COM/manual acceptance testing, not for normal hosted build/unit-test gates.

## Documentation

- [`plan.md`](plan.md) — product, architecture, safety rules, phases, and acceptance criteria.
- [`docs/ui.md`](docs/ui.md) — production GUI/interaction contract and current implementation boundary.
- [`docs/phase-0.md`](docs/phase-0.md) — completed bootstrap and verified stable-version matrix.
- [`docs/phase-1.md`](docs/phase-1.md) — completed read-only Outlook probe and packaging lessons.
- [`docs/phase-2.md`](docs/phase-2.md) — completed native meeting-forwarding spike and real-machine evidence.
- [`docs/phase-3.md`](docs/phase-3.md) — current UI-assisted development/read-alignment increment.
