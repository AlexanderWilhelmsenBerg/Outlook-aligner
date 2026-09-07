# Outlook Aligner

Outlook Aligner is a Windows desktop project for comparing and deliberately aligning calendar events across multiple accounts in one Classic Outlook profile.

## Integration boundary

The project uses the Classic Outlook COM/Object Model. Microsoft Graph and MSAL are deliberately out of scope; no Azure app registration, Graph permission, or tenant administrator consent is required by the design.

## Project status

- **Phase 0 — Repository/toolchain bootstrap: Complete ✅**
- **Phase 1 — Outlook COM discovery/read probe: In progress 🚧 (PR-only)**

Phase 0 was the one-time bootstrap authorized directly on `main`. Every implementation change from Phase 1 onward is delivered through a pull request and is not merged automatically.

The Phase 0 implementation baseline passed hosted CI on 2026-09-07 and every direct dependency was reverified against the latest stable release. See [`docs/phase-0.md`](docs/phase-0.md) for the version matrix and verification checkmarks.

## Phase 1 test executable

Phase 1 CI publishes a self-contained Windows x64 diagnostic executable as the Actions artifact:

`OutlookAligner-Phase1-Probe-win-x64`

After downloading/extracting it on the Windows PC with Classic Outlook configured, run:

```powershell
.\OutlookAligner.OutlookHost.exe --days 90
```

The probe is read-only and withholds event subjects/locations by default. See [`docs/phase-1.md`](docs/phase-1.md) for download instructions, CLI options, and the manual acceptance checklist.

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

Classic Outlook is required only for the Phase 1 manual COM integration probe, not for normal hosted build/unit-test gates.

## Documentation

- [`plan.md`](plan.md) — product, architecture, safety rules, phases, and acceptance criteria.
- [`docs/phase-0.md`](docs/phase-0.md) — completed bootstrap and verified stable-version matrix.
- [`docs/phase-1.md`](docs/phase-1.md) — read-only Outlook probe design and test procedure.
