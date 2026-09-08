# Outlook Aligner

Outlook Aligner is a Windows desktop project for comparing and deliberately aligning calendar events across multiple accounts in one Classic Outlook profile.

## Integration boundary

The project uses the Classic Outlook COM/Object Model. Microsoft Graph and MSAL are deliberately out of scope; no Azure app registration, Graph permission, or tenant administrator consent is required by the design.

## Project status

- **Phase 0 — Repository/toolchain bootstrap: Complete ✅**
- **Phase 1 — Outlook COM discovery/read probe: Complete ✅ / merged in PR #4**
- **Phase 2 — Native meeting-forwarding technical spike: In progress 🚧 / PR-only**

Phase 0 was the one-time bootstrap authorized directly on `main`. Every implementation change from Phase 1 onward is delivered through a pull request and is not merged automatically.

Phase 1 proved the packaged Classic Outlook COM read boundary on the user's real Windows/Outlook profile, including the self-contained interop packaging fix discovered during manual testing. See [`docs/phase-1.md`](docs/phase-1.md).

## Phase 2 test executable

Phase 2 CI publishes the OutlookHost as a self-contained Windows x64 diagnostic executable. The existing read probe remains available, while the forwarding spike adds explicit inspect/prepare/send modes.

First identify a real accepted meeting and its `GlobalAppointmentId` from the read probe:

```powershell
.\OutlookAligner.OutlookHost.exe --days 90 --json --include-details
```

Then inspect whether Outlook still retains the native meeting request for a chosen source account:

```powershell
.\OutlookAligner.OutlookHost.exe --forward-spike --source-smtp source@example.com --global-id GLOBAL_ID
```

Use `--forward-spike --help` for the deliberately gated prepare/send commands. No vCalendar fallback is used. See [`docs/phase-2.md`](docs/phase-2.md).

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

Classic Outlook is required for live Phase 1/2 COM testing, not for normal hosted build/unit-test gates.

## Documentation

- [`plan.md`](plan.md) — product, architecture, safety rules, phases, and acceptance criteria.
- [`docs/phase-0.md`](docs/phase-0.md) — completed bootstrap and verified stable-version matrix.
- [`docs/phase-1.md`](docs/phase-1.md) — completed read-only Outlook probe and packaging lessons.
- [`docs/phase-2.md`](docs/phase-2.md) — native meeting-forwarding spike design and manual acceptance procedure.
