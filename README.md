# Outlook Aligner

Outlook Aligner is a **WinUI 3 Windows desktop application** for comparing and deliberately aligning calendar events across multiple accounts in one Classic Outlook profile.

The command-line tools used in the early technical phases are diagnostic/test harnesses only. The shipped v1 product requires a graphical interface; see [`docs/ui.md`](docs/ui.md).

## Integration boundary

The project uses the Classic Outlook COM/Object Model. Microsoft Graph and MSAL are deliberately out of scope; no Azure app registration, Graph permission, or tenant administrator consent is required by the design.

## Project status

- **Phase 0 — Repository/toolchain bootstrap: Complete ✅**
- **Phase 1 — Outlook COM discovery/read probe: Complete ✅ / merged in PR #4**
- **Phase 2 — Native meeting-forwarding technical spike: end-to-end mechanism proven; reliability matrix in progress 🚧 / PR #5**
- **Phase 4 — Production WinUI GUI: required for v1; specified now, implemented after identity work**

Phase 0 was the one-time bootstrap authorized directly on `main`. Every implementation change from Phase 1 onward is delivered through a pull request and is not merged automatically.

Phase 1 proved the packaged Classic Outlook COM read boundary on the user's real Windows/Outlook profile, including the self-contained interop packaging fix discovered during manual testing. See [`docs/phase-1.md`](docs/phase-1.md).

## Phase 2 result so far

Phase 2 has proven a genuine native forwarding path on a real accepted Classic Outlook meeting even when the original meeting request is no longer retained in Inbox/Deleted Items.

The proven path is:

```text
accepted Calendar AppointmentItem
  -> verify Outlook built-in Forward capability
  -> ExecuteMso("Forward")
  -> AppointmentItem.Forward event
  -> native MeetingItem (IPM.Schedule.Meeting.Request)
  -> exactly one resolved recipient
  -> SendUsingAccount = selected source account
  -> MeetingItem.Send()
```

A real Kverneland meeting was forwarded to another user-controlled Outlook account and arrived there as a forwarded meeting. No `AppointmentItem.ForwardAsVcal()` or fake ICS fallback is used.

The remaining Phase 2 work is a deliberately small reliability matrix covering recurrence/Teams behavior, a different source account, and forwarding-disabled behavior if a suitable event is readily available. See [`docs/phase-2.md`](docs/phase-2.md).

## Production UI

The v1 GUI is not optional. `OutlookAligner.App` will provide a WinUI 3 `NavigationView` shell with:

- **Calendar** — primary synchronized three-account calendar, event selection and comparison;
- **Alignment** — discrepancy queue, filters, authority and safe actions;
- **Settings** — horizon, privacy and UI preferences;
- **Diagnostics** — Outlook/IPC health, operation history and technical identifiers.

Normal v1 operation must not require PowerShell or CLI arguments. Write actions will be capability-driven, and bulk writes require preview + explicit confirmation. The detailed UI/UX and Phase 4 acceptance contract is in [`docs/ui.md`](docs/ui.md).

For **Forward meeting**, the production GUI treats native forwarding as a capability of the selected Calendar event: enable the action only when Classic Outlook reports the built-in Forward command as available and all identity/safety checks pass. Unsupported events fail closed with a clear explanation. The retained-request diagnostic path is not exposed as a separate user-facing action.

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
- [`docs/ui.md`](docs/ui.md) — required production GUI/interaction contract.
- [`docs/phase-0.md`](docs/phase-0.md) — completed bootstrap and verified stable-version matrix.
- [`docs/phase-1.md`](docs/phase-1.md) — completed read-only Outlook probe and packaging lessons.
- [`docs/phase-2.md`](docs/phase-2.md) — native meeting-forwarding spike design, real-machine evidence, and remaining reliability matrix.
