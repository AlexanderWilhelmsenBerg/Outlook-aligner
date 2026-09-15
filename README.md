# Outlook Aligner

Outlook Aligner is a **WinUI 3 Windows desktop application** for comparing calendar events across multiple accounts in one Classic Outlook profile, first as a complete read-only overview and later as a deliberate reconciliation tool.

The integration boundary is the **Classic Outlook COM/Object Model**. Microsoft Graph/MSAL/Azure app registration are deliberately out of scope.

## Current product direction

Technical foundations are already in place for Outlook discovery, bounded calendar reads, native Forward experimentation, a WinUI shell, preliminary correlation, diagnostics, and packaged test builds.

The active roadmap is now split into two segments:

- **Segment A — Read-only overview**: finish a genuinely useful calendar/report product before testing or productizing reconciliation writes.
- **Segment B — Reconciliation**: add Forward/Copy/Move only after Segment A is accepted.

See [`plan.md`](plan.md) for the full roadmap and [`docs/ui.md`](docs/ui.md) for the UI contract.

### Segment A phases

1. **Month calendar view** — one visual event per logical event.
2. **Account markers + status colors** — account dots top-left; green/all aligned, yellow/partial or moved, red/single-account.
3. **Recurring meetings** — occurrence/exception identity plus recurrence symbol lower-right.
4. **Read-only reconciliation report** — Missing, Moved, Duplicate, Conflict/unresolved.
5. **Filters and view modes** — Calendar only, Report only, combined, plus status/account filters.

Authority is event-specific rather than tied to one permanent master account. If one event reliably originates through Account A, A can be authoritative for that event; another event can be authoritative from B or C. Scan order and LastModificationTime are not authority evidence.

## Current PR #6 foundation

The current `phase-3/ui-assisted-testing` branch provides:

- real WinUI 3 application shell;
- Calendar, Alignment, Settings and Diagnostics destinations;
- GUI-driven account/calendar discovery;
- bounded Outlook reads;
- preliminary non-recurring logical grouping;
- read-only managed-copy metadata classification;
- authority/Move-preview scaffolding with no Move execution;
- structured persistent diagnostics log;
- safe Forward capability/prepare diagnostics with no GUI send;
- self-contained Windows CI test bundle;
- one-command local updater.

This foundation is intended to be merged before the new Segment A calendar/report work begins. Forward/reconciliation testing is no longer a merge gate for this PR.

## Update the current test app

Clone the repository and use:

```text
scripts\Update-OutlookAlignerTestApp.cmd
```

or from PowerShell:

```powershell
./scripts/Update-OutlookAlignerTestApp.ps1
```

Administrator rights are **not required**. If GitHub CLI is not installed, the updater downloads the official portable Windows `gh.exe` into:

```text
%LOCALAPPDATA%\OutlookAligner\Tools\GitHubCLI\gh.exe
```

The updater then:

1. finds the latest successful CI build for the configured branch;
2. downloads `OutlookAligner-Phase3-UI-TestHarness-win-x64`;
3. verifies the UI and OutlookHost executables exist;
4. keeps the previous local build as `TestApp.previous`;
5. installs the new build under `%LOCALAPPDATA%\OutlookAligner\TestApp`;
6. unblocks/signs Outlook Aligner binaries when a trusted local test certificate is available;
7. verifies Authenticode signatures are `Valid` before launch;
8. launches the app unless `-NoLaunch` is supplied.

### Local test signing

The updater looks for a current-user code-signing certificate with subject:

```text
CN=Outlook Aligner Local Test
```

It **does not create or trust a root certificate automatically**. Trust setup is intentionally a separate one-time user/admin-policy decision, especially on managed work computers.

Once such a certificate already exists in `Cert:\CurrentUser\My` with its public certificate trusted in the current user's Root and TrustedPublisher stores, normal updater runs sign these local test files automatically:

- `OutlookAligner.*.exe`
- `OutlookAligner.*.dll`

Useful updater options:

```powershell
# Update but do not launch.
./scripts/Update-OutlookAlignerTestApp.ps1 -NoLaunch

# Force a fresh download of the current successful build.
./scripts/Update-OutlookAlignerTestApp.ps1 -Force

# Remove downloaded-file zone markers from all files in the installed bundle.
./scripts/Update-OutlookAlignerTestApp.ps1 -UnblockFiles

# Intentionally skip local Authenticode signing.
./scripts/Update-OutlookAlignerTestApp.ps1 -SkipLocalSigning

# Select a specific local signing certificate if multiple matching certs exist.
./scripts/Update-OutlookAlignerTestApp.ps1 -SigningCertificateThumbprint "<thumbprint>"
```

The persistent diagnostics log lives separately at:

```text
%LOCALAPPDATA%\OutlookAligner\events.jsonl
```

so replacing the test app does not erase test history.

## Verified toolchain baseline

- .NET SDK 10.0.400 / .NET 10 / C# 14
- Windows App SDK 2.4.0 / WinUI 3
- WebView2
- Outlook Interop
- Node.js 24 / npm
- TypeScript / Vite / FullCalendar
- xUnit v3 on Microsoft Testing Platform
- BenchmarkDotNet
- GitHub Actions + Dependabot

## Local repository verification

```powershell
./scripts/verify.ps1
```

Classic Outlook is required for live COM/manual acceptance testing, not for normal hosted build/unit-test gates.

## Documentation

- [`plan.md`](plan.md) — Segment A/Segment B roadmap, architecture, safety rules and acceptance gates.
- [`docs/ui.md`](docs/ui.md) — read-only-first production UI contract.
- [`docs/phase-0.md`](docs/phase-0.md) — repository/toolchain bootstrap history.
- [`docs/phase-1.md`](docs/phase-1.md) — Classic Outlook read-boundary proof.
- [`docs/phase-2.md`](docs/phase-2.md) — native Forward technical proof.
- [`docs/phase-3.md`](docs/phase-3.md) — current WinUI/read-alignment foundation history.
