# Outlook Aligner

Outlook Aligner is a Windows desktop project for comparing and deliberately aligning calendar events across multiple accounts in one Classic Outlook profile.

## Integration boundary

The project uses the Classic Outlook COM/Object Model. Microsoft Graph and MSAL are deliberately out of scope; no Azure app registration, Graph permission, or tenant administrator consent is required by the design.

## Project status

Phase 0 bootstraps the repository and toolchain. Outlook COM behavior begins in Phase 1.

After Phase 0, all implementation changes are delivered through pull requests rather than directly to `main`.

## Toolchain

- .NET SDK 10.0.400 / C# 14
- Windows App SDK / WinUI 3 dependencies
- Node.js 24 LTS
- TypeScript + Vite + FullCalendar
- xUnit v3 on Microsoft Testing Platform
- BenchmarkDotNet
- GitHub Actions + Dependabot

See [`plan.md`](plan.md) for the product/architecture roadmap and [`docs/phase-0.md`](docs/phase-0.md) for bootstrap details.

## Local verification

Prerequisites:

- .NET SDK 10.0.400
- Node.js 24 or newer in the Node 24 LTS line
- PowerShell 7 recommended

Run:

```powershell
./scripts/verify.ps1
```

Classic Outlook is not needed for Phase 0 build/tests. It becomes necessary for the Phase 1 manual COM integration probe.
