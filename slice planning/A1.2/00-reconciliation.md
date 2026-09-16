# Agent 00 reconciliation — A1.2

Status: **SLICE READY**

## Authoritative state

- Current `main`: `7dbd437c1ee86bf6e41d8ad939a54cc81ee7255e`
- Exact-head `main` CI: **success**, run `34961414443` (#212)
- A1.1 merged in the current `main` head and was accepted post-merge by Agent 50.
- A1.2 UX readiness: `slice planning/A1.2/20-ux-note.md`, final token `UX READY`.
- A1.2 architecture readiness: `slice planning/A1.2/30-architecture-note.md`, final token `ARCH READY`.

Agent 00 revalidated both specialist notes against the actual merged A1.1 result. A1.1 changed only the deterministic frontend toolchain surfaces: `web/calendar/package-lock.json`, `.github/workflows/ci.yml`, and `scripts/verify.ps1`. It preserved the existing frontend source, FullCalendar/Vite versions, WinUI/App/Core/OutlookHost behavior, WebView2 assumptions, and Segment A product semantics. Therefore neither A1.2 specialist contract requires refresh.

The UX and architecture notes are compatible. No unresolved cross-discipline conflict remains.

## Reconciled A1.2 implementation contract

Agent 40 may implement **A1.2 — WebView2 month-calendar host** with these fixed decisions:

1. **Primary surface:** Calendar opens to a real FullCalendar month grid. The current list is no longer the co-equal/default Calendar presentation.
2. **Native shell ownership:** WinUI retains application navigation, Calendar heading, scan horizon, Refresh, busy/progress state, status/warning/error communication, Diagnostics, and the visible Today / Previous month / Next month controls plus current month title.
3. **WebView ownership:** WebView2 hosts presentation only. FullCalendar keeps its own toolbar disabled and responds to typed navigation/render commands from WinUI.
4. **Observation semantics only:** A1.2 renders current per-account observations. It must not claim one-logical-event correlation, authority, alignment health, Missing/Moved/Conflict, account markers, or other A1.3/A2 semantics.
5. **Process boundary:** Classic Outlook COM remains confined to OutlookHost. App receives plain host DTOs and maps only presentation-safe fields into a small versioned WebView contract. Never send COM objects or raw Outlook locators/identity fields to JavaScript.
6. **Presentation identity:** The App creates an opaque session-scoped `presentationId` for each rendered observation and owns the in-memory mapping back to the current observation/view model. Selection from WebView returns only that opaque ID. Unknown/stale IDs fail safely.
7. **Versioned messaging:** Use structured JSON with a small versioned host/client message contract. Required directions are render/update observations, Today/Previous/Next navigation, ready/rangeChanged, and observationSelected. Malformed/unknown version/type messages are ignored safely and may be logged diagnostically.
8. **Asset lifecycle:** Build Vite output reproducibly using the merged A1.1 Node/npm/lockfile contract. Package built assets beside/in the WinUI publish output under a stable application-owned directory. Do not commit generated `dist` and do not depend on repository/developer-machine paths.
9. **Runtime loading:** Load bundled assets through an app-owned WebView2 virtual-host mapping (or current equivalent) with a stable local HTTPS-style origin. Block unexpected top-level navigation; the WebView is not a general browser.
10. **CI/publish integration:** The Windows publish path must deterministically obtain/build the frontend before publishing the App, and publishing must fail if required calendar assets are missing. Existing separate frontend gates stay intact.
11. **Refresh/navigation behavior:** Create one calendar instance per WebView lifetime. Refresh replaces the bounded render payload without reloading the browser. Month navigation is local and must not trigger a new OutlookHost scan; navigating beyond currently loaded observations may truthfully show an empty month.
12. **Selection:** Selecting a rendered item updates the existing read-only single-observation selection/detail model. No write action, account authority, correlation decision, or Segment-B command is exposed.
13. **Visible states:** Loading, successful-empty, OutlookHost failure, partial/degraded results, and WebView/content failure remain visibly communicated by native WinUI. Healthy observations remain visible on partial results. A blank WebView is not an accepted error state.
14. **Forward diagnostics:** Remove/hide Forward capability, test-target, Check Forward, and Prepare Forward controls from the normal Segment A Calendar workflow. Dormant Segment-B spike/service code may remain if deleting it would cause unrelated churn.
15. **Accessibility/resilience:** Today/Previous/Next and calendar items are keyboard reachable with visible focus; icon controls have accessible names; long text/high DPI/text scaling remain usable; light/dark presentation is coherent; native fallback/status communication remains independent of the WebView DOM.

## Required automated evidence

Agent 40 must add or preserve focused evidence for:

- App-side mapping from plain observation/read DTOs to the presentation DTO;
- serialized WebView payloads excluding raw Outlook technical IDs/managed metadata;
- opaque presentation-ID mapping and safe stale/unknown selection handling;
- valid and invalid client-message parsing/version/type validation;
- synthetic frontend render/update, navigation, range reporting, and selection behavior without Outlook COM;
- deterministic A1.1 install/build gates remaining green;
- published WinUI output containing the built calendar `index.html` and assets;
- publish/CI failing rather than silently succeeding when required frontend payload is absent;
- existing .NET/Core/OutlookHost tests remaining green.

Do not introduce a large browser-automation framework merely to test pure message-shaping logic if smaller TypeScript/.NET seams suffice.

## Required manual acceptance evidence

On the exact implementation head and a real Windows/Classic Outlook profile, verify at minimum:

- Calendar opens with the month grid as the primary surface;
- Today/Previous/Next work and the native visible month title follows the displayed range;
- Refresh/horizon/status remain native and bounded Outlook observations render without PowerShell or raw Outlook IDs;
- selecting an event produces truthful read-only single-observation detail;
- zero-result month leaves a usable grid plus explicit native empty state;
- OutlookHost failure and WebView/asset failure remain visible natively with Diagnostics reachable;
- partial results keep healthy events visible with a degraded warning;
- Forward/Prepare/test-target controls are absent from the normal Calendar workflow;
- keyboard/focus, practical resize/high DPI/text scaling, and light/dark presentation remain usable;
- no A1.3/A2 logical-event, account-marker, authority, or health semantics appear.

## Expected implementation boundary

A conforming implementation is expected to stay primarily within:

- `src/OutlookAligner.App/**` for Calendar host layout, WebView lifecycle/bridge, presentation messages/mapping, and selection wiring;
- `web/calendar/src/**` for rendering/navigation/message handling;
- focused App/frontend tests;
- App project/build packaging and `.github/workflows/ci.yml` / `scripts/verify.ps1` only as needed to consume the deterministic frontend output.

A1.2 should not require OutlookHost COM enumeration changes, Core correlation/status logic, persistence, host protocol-version changes, or Segment-B action implementation. If a render-critical host fact is genuinely missing, stop for architecture re-review rather than broadening the boundary ad hoc.

## Explicit exclusions

Do not implement A1.3 logical-event correlation, A2 account markers/health colors, authority/origin classification, reconciliation actions, Copy/Move/Forward product workflows, recurring writes, deletion synchronization, or any other Outlook mutation.

## Blockers

None in current authoritative state.

**SLICE READY**
