# Agent 30 architecture note — A1.2

Status: **READY, provisional until A1.1 is merged and revalidated by Agent 00**

## Exact state reviewed

- `main` HEAD reviewed: `e8bc60cd6e2acc1bc32e538d661db88262198226`
- `planning/segment-a-slices` HEAD read before writing: `211e6c0f90e1d5524e6564f8b89b6a02a0e03455`
- Planning PR: #9, documentation/planning only.

This review is intentionally one slice ahead. A1.2 implementation remains blocked until A1.1 is accepted and merged. Agent 00 must refresh `main` after that merge and revalidate this note before issuing `SLICE READY`.

## Files and contracts inspected

- `AGENTS.md`, especially Agent 30 ownership, process-boundary rules, and Segment A read-only constraints.
- `plan.md`, especially **Slice A1.2 — WebView2 month-calendar host** and the A1.3/A2 boundaries.
- `slice planning/Slice-planning.md`.
- `slice planning/A1.2/00-brief.md`.
- `slice planning/A1.2/30-architecture-prompt.md`.
- `slice planning/A1.1/30-architecture-note.md` as the pending prerequisite/toolchain contract.
- `slice planning/A1.2/20-ux-note.md` to align architecture with the accepted visible-shell/navigation expectations.
- `docs/ui.md`.
- `docs/phase-3.md`.
- `.github/workflows/ci.yml`.
- `src/OutlookAligner.App/MainWindow.xaml`.
- `src/OutlookAligner.App/MainWindow.xaml.cs`.
- `src/OutlookAligner.App/OutlookAligner.App.csproj`.
- `src/OutlookAligner.App/ViewModels/MainViewModel.cs`.
- `src/OutlookAligner.App/ViewModels/CalendarEventRowViewModel.cs`.
- `src/OutlookAligner.Outlook.Contracts/ProbeDtos.cs`.
- current `tests/` project layout.
- `web/calendar/package.json`.
- `web/calendar/src/main.ts` and the current FullCalendar scaffold.

## Current architecture facts

1. The WinUI shell already owns navigation, refresh, scan horizon, busy state, status text, degraded/error notice presentation, Diagnostics, and selected-observation detail.
2. `MainViewModel` already consumes plain `OutlookProbeResult` / `OutlookAccountDto` / `CalendarEventDto` records from the host client. Outlook COM RCWs therefore do not need to move any closer to the UI to implement A1.2.
3. The current app displays one `CalendarEventRowViewModel` per observed account event. That is acceptable for A1.2. One-logical-event correlation belongs to A1.3 and must not be invented here.
4. The current frontend already creates a FullCalendar `dayGridMonth` with `headerToolbar: false`; it is presently a standalone scaffold and is not hosted by the WinUI Calendar page.
5. `Microsoft.Web.WebView2` is already referenced by the App project, but there is no current WebView2 lifecycle/message implementation in `MainWindow`.
6. The current Calendar UI exposes Forward capability/prepare diagnostics. A1.2 removes those controls from the normal Segment-A Calendar surface but does not require deleting dormant Segment-B host/service code.
7. Current hosted CI builds the calendar frontend independently on Ubuntu and publishes the WinUI app independently on Windows. Today, the published WinUI artifact does not consume `web/calendar/dist`.
8. A1.1's pending contract establishes a real npm lockfile, Node `24.20.0` / npm `11.19.0`, and `npm ci --ignore-scripts`. A1.2 must consume that deterministic output rather than introduce a second frontend dependency/install policy.

## Architecture contract for A1.2

### 1. Ownership and lifecycle

The existing **WinUI `MainWindow`/Calendar page owns the WebView2 control and its lifecycle**.

The WebView is a child presentation surface inside the Calendar content area. It does not become an application shell and does not own Outlook refresh, host process lifecycle, diagnostics, scan horizon, application navigation, or product state.

A narrow App-side bridge/controller should own WebView2-specific concerns instead of allowing `MainWindow.xaml.cs` to accumulate protocol parsing and JavaScript strings. The exact class name is implementation detail, but its responsibilities are constrained to:

- initialize the WebView2 environment/control;
- map/load the built frontend assets;
- attach/detach WebView message/navigation handlers;
- serialize/version messages sent to the frontend;
- validate messages received from the frontend;
- expose a small typed API to the WinUI layer such as initialize, set observations, navigate today/previous/next, and dispose.

`MainWindow` remains responsible for binding these bridge events to the existing `MainViewModel` selection/status state.

Initialization must be one-shot per WebView instance and safe against repeated Loaded/navigation callbacks. Event handlers must be detached/disposed when the window/control is closed so a WebView recreation does not accumulate duplicate handlers.

### 2. Frontend asset build and location

A1.2 consumes **built Vite output**, not TypeScript source at runtime.

After A1.1 is merged, `web/calendar` is installed/built with its deterministic Node/npm/lockfile contract. The resulting `web/calendar/dist/**` files are copied into the App output/publish tree under one stable application-owned directory, for example `CalendarWeb/**`. The exact folder name may differ, but it must be stable and identical for local build, CI publish, and the portable test bundle.

Do not commit generated `dist` as a substitute for a reproducible build.

The App build/publish must fail if the expected built `index.html`/assets are missing. A successful .NET publish must therefore imply that the frontend payload packaged beside the app is complete.

Because the current Windows .NET CI job does not install/build the frontend, A1.2 must update hosted build orchestration so the **Windows job that publishes the app** obtains the deterministic frontend output before `dotnet publish`. The simplest accepted shape is to set up the exact A1.1 Node/npm baseline in that job and run `npm ci --ignore-scripts` + `npm run build` before the app publish/copy step. Reusing an artifact produced by another job is also valid if exact-head provenance is explicit, but A1.2 must not depend on an undeclared local `dist` directory.

The existing separate frontend quality-gate job remains useful and should not be weakened.

### 3. Runtime loading mechanism

Load application-owned frontend assets through a **WebView2 virtual-host mapping** rooted at the packaged calendar asset directory, then navigate to a stable HTTPS-style local origin such as:

`https://calendar.outlook-aligner.local/index.html`

The exact host string is implementation detail, but it must be app-owned and used only for bundled content.

Prefer `SetVirtualHostNameToFolderMapping` (or the equivalent current WebView2 API) over `file://` navigation. This gives the frontend a stable origin and avoids brittle absolute filesystem URLs.

The mapping must resolve from the actual executable/output location in both ordinary development runs and the self-contained published app. Do not hard-code repository paths or developer-machine paths.

Navigation outside the mapped app origin is not needed for A1.2. The host should cancel/block unexpected top-level navigation attempts rather than turning the calendar WebView into a general browser.

### 4. Small versioned WinUI -> WebView render contract

Do **not** send `OutlookProbeResult`, `OutlookAccountDto`, or `CalendarEventDto` wholesale to JavaScript. Those transport DTOs contain technical locators/identity fields that the normal Calendar surface does not need.

Introduce a small **presentation-only versioned envelope**. A concrete contract may use records/classes with different names, but must be equivalent to:

```text
CalendarHostMessage
- version: 1
- type: "render"
- payload:
  - observations: CalendarObservation[]

CalendarObservation
- presentationId: string
- subject: string
- startLocal: string (ISO-8601 local/date-time representation with explicit agreed semantics)
- endLocal: string
- isAllDay: bool
- sourceDisplayName: string
- location: string?   // only if current privacy/UI contract allows it
- isRecurring: bool
```

The contract is intentionally an **observation renderer**, not a logical-event model.

Rules:

- `presentationId` is an opaque session/UI key generated/owned by the App layer. It is not StoreID, EntryID, GlobalAppointmentID, SMTP address, SyncGroupId, or any other Outlook locator/domain identity.
- Raw technical IDs and managed-copy metadata are excluded from WebView messages.
- A1.2 may include only fields needed to render/select the current observation truthfully.
- No `Aligned`, `Missing`, `Moved`, `Duplicate`, `Conflict`, authority, origin, cross-account membership, or health fields are added. Those semantics belong to A1.3/A2 and later.
- The App keeps the mapping from opaque `presentationId` back to the existing selected observation/view-model object in memory for the current refresh.
- A refresh replaces that mapping atomically. A selection message carrying an unknown/stale presentation ID is ignored/rejected safely rather than being interpreted as an Outlook identifier.

Serialization must use normal structured JSON serialization. Do not build executable JavaScript with interpolated subjects/locations.

### 5. WebView -> WinUI contract

A1.2 needs only two event families from the frontend:

1. **Ready/range state** after the calendar has initialized or changed date range.
2. **Observation selection** when a rendered event is activated.

Equivalent versioned messages:

```text
CalendarClientMessage
- version: 1
- type: "ready" | "rangeChanged" | "observationSelected"
- payload: type-specific data
```

For `rangeChanged`, return presentation data such as displayed title and visible start/end. These are display/navigation facts, not an instruction to query Outlook directly.

For `observationSelected`, return only the opaque `presentationId`.

Every received message must be parsed defensively and validated by version/type before use. Unknown versions/types or malformed payloads are logged diagnostically and ignored; they must not crash the App or trigger host/write actions.

### 6. Navigation ownership

Per the A1.2 UX contract, **WinUI owns the visible Today / Previous month / Next month controls and current-period title**. FullCalendar's own header remains disabled.

WinUI commands invoke typed bridge methods that issue small commands to the frontend:

```text
navigate: today
navigate: previous
navigate: next
```

The frontend performs the FullCalendar navigation and emits `rangeChanged` containing the resulting title/range. WinUI updates its visible month title from that validated response.

A1.2 does **not** make range changes trigger a new Outlook scan automatically. The current bounded refresh/horizon remains the source dataset; navigating to a month outside the loaded observations may truthfully show an empty grid until the user changes horizon/refreshes. A future slice may choose smarter range-driven scanning, but Agent 40 must not invent it here.

### 7. Loading, empty, host error, partial results, and WebView failure

The native WinUI shell remains the authoritative state/fallback channel.

- **Loading:** the existing `IsBusy`/status/progress UI remains visible independently of the WebView. The calendar may remain hidden/disabled until bridge readiness, but Refresh and Diagnostics remain native and reachable.
- **Successful empty result:** send an empty observation array, keep the month grid alive, and expose the agreed native empty-state text. Empty is not a host failure.
- **OutlookHost/read failure:** retain the current native failure status/InfoBar behavior. Do not replace it with an HTML error page.
- **Partial result:** send/render healthy returned observations and keep the current native degraded warning visible. Do not derive Missing/Moved/Conflict from partiality in TypeScript.
- **WebView2 initialization/content failure:** keep the native Calendar shell alive, surface a native calendar-host failure message, log technical detail to Diagnostics, and leave Refresh/Diagnostics usable. A blank WebView is not an accepted failure mode.

The App should distinguish at least **data/read failure** from **calendar-renderer failure** so the user is not told Outlook failed merely because WebView2/assets failed.

### 8. Selection and existing detail model

A1.2 selection is single-observation selection only.

When the WebView emits a valid `presentationId`, App resolves it to the current `CalendarEventRowViewModel` (or equivalent observation view model) and assigns the existing `SelectedEvent`. The retained native read-only detail fields then continue to be sourced from App/view-model state, not reconstructed in TypeScript.

Do not pass raw IDs back from JavaScript. Do not let the frontend choose an account, authority, correlation group, Forward target, Move candidate, or any other Segment-B/domain action.

The existing Forward/Prepare controls are removed/hidden from the normal Calendar UI per the slice contract, but dormant command/service implementation may remain untouched if not needed for the visible path.

### 9. Explicit COM and Segment-B boundary

The allowed data path is:

`Classic Outlook COM -> OutlookHost STA -> plain host protocol DTOs -> App/MainViewModel -> small calendar presentation DTO -> WebView2/FullCalendar`

The reverse path is limited to presentation events:

`WebView2 -> validated presentation message -> App selection/navigation state`

Forbidden in A1.2:

- COM RCWs or dynamic COM objects in App/Core/WebView payloads;
- StoreID/EntryID/GlobalAppointmentID in ordinary WebView messages;
- JavaScript directly invoking OutlookHost or a child-process command;
- Forward/Prepare/Copy/Move/send/delete/action messages from WebView;
- persistence of WebView selection/navigation as a new product contract;
- a second correlation/status engine in TypeScript;
- production IPC redesign merely to host the month calendar.

No OutlookHost protocol-version increment is required solely for A1.2 because the current host DTO already contains the bounded observation fields needed by the App. If implementation discovers that the existing host DTO genuinely lacks a render-critical observation fact, that is an architecture re-review point rather than permission to expose COM/raw objects.

### 10. Frontend responsibilities

TypeScript/FullCalendar may:

- render the passed observations;
- format/present month cells and observation text;
- perform local FullCalendar date navigation;
- emit range and selection events;
- follow the app light/dark presentation contract;
- provide accessible labels/focus behavior required by Agent 20.

TypeScript/FullCalendar may not:

- correlate observations across accounts;
- infer origin/authority;
- classify alignment health;
- inspect or reason over Outlook technical IDs;
- decide that a partial result implies a missing copy;
- initiate Outlook writes or host actions.

### 11. Performance/reload expectation

A1.2 is bounded-month presentation, not virtualization infrastructure.

- Create/render one FullCalendar instance for the WebView lifetime rather than recreating the WebView or reloading HTML on every Outlook refresh.
- After bridge readiness, refresh data by replacing/updating the frontend event source from the serialized observation array.
- Avoid one WebMessage call per observation; send a single bounded render payload per successful refresh (or another small bounded batch strategy only if concrete payload evidence requires it).
- Navigation should be local/immediate and must not spawn an OutlookHost scan.
- A normal refresh must not require a full browser navigation/reinitialization.

No arbitrary event-count threshold is introduced in this planning slice. Existing scan horizon provides the bound. If real-profile evidence shows message-size/render latency is problematic, record measurements and bring the threshold/batching decision back to architecture rather than inventing hidden pagination.

## Test seams and required evidence

Agent 40 should implement/test A1.2 with seams that do not require Outlook COM for presentation validation.

At minimum:

1. **App-side presentation mapping tests**
   - Given plain account/event DTOs or observation view models, mapping produces the expected calendar presentation DTO.
   - Technical IDs/managed metadata are absent from serialized WebView payloads.
   - Empty input serializes as a valid empty render message.
   - Opaque presentation IDs resolve to the current observation and stale/unknown IDs fail safely.

2. **Message parsing/validation tests**
   - valid ready/range/selection messages are accepted;
   - malformed JSON, unknown message types, unknown protocol versions, and unknown presentation IDs do not crash or trigger actions.

3. **Frontend tests/seams without COM**
   - frontend logic accepts synthetic render messages and creates/removes FullCalendar events correctly;
   - today/previous/next commands call the corresponding calendar navigation and report range changes;
   - selection emits only the opaque presentation ID.

   Agent 40 may use the existing Node toolchain and add the smallest suitable frontend test harness if needed. Do not introduce a browser-automation stack merely to prove message-shaping functions that can be tested as pure TypeScript.

4. **Build/publish evidence**
   - clean checkout uses the post-A1.1 `npm ci --ignore-scripts` path;
   - frontend `typecheck`, `lint`, `format:check`, `build`, and audit gates stay green;
   - the Windows published App bundle contains the built `CalendarWeb`/equivalent `index.html` and hashed assets;
   - CI fails if the asset payload is absent rather than silently publishing a blank host.

5. **Manual Windows/WebView2 evidence on the exact implementation head**
   - app launches with Calendar as the default destination;
   - month grid loads from packaged assets in the published bundle;
   - Today/Previous/Next work and title tracks the displayed month;
   - Outlook refresh populates the grid with current bounded observations;
   - selecting an event updates the truthful read-only observation detail;
   - zero results, OutlookHost failure, partial result, and WebView/asset failure each preserve a native visible state;
   - normal Calendar workflow exposes no Forward/Prepare target/action controls and no raw technical IDs.

Existing .NET/Core/OutlookHost tests remain COM-independent where they already are. A1.2 does not require live Outlook COM in unit tests.

## Expected implementation scope

The exact file split is Agent 40's implementation choice, but an architecture-conforming A1.2 diff is expected to touch narrowly:

- `src/OutlookAligner.App/MainWindow.xaml` and/or its Calendar page host layout;
- `src/OutlookAligner.App/MainWindow.xaml.cs` for shell wiring only;
- one or more App-side WebView bridge/message/presentation classes;
- `src/OutlookAligner.App/OutlookAligner.App.csproj` or equivalent build-copy target for packaged frontend assets;
- `web/calendar/src/**` for the bridge/event rendering/navigation logic;
- focused presentation/message tests;
- `.github/workflows/ci.yml` and, if necessary, `scripts/verify.ps1` only to ensure the published App consumes the deterministic frontend output established by A1.1.

A1.2 should **not** need changes to OutlookHost COM enumeration, host protocol DTO shape/version, Core correlation/authority/status logic, persistence, or Segment-B action planners.

## Blockers / dependency

There is no unresolved A1.2 architecture decision in the repository state reviewed.

The only procedural dependency is A1.1: A1.2 must not be handed to Agent 40 until A1.1 is accepted/merged and Agent 00 verifies that the actual merged lockfile/build changes still satisfy the assumptions above. If A1.1 materially changes the frontend output directory, Node/npm baseline, install command, or CI build arrangement, this note needs a focused refresh before `SLICE READY`.

ARCH READY
