# Agent 20 prompt — A1.2 UX readiness

Repository: `AlexanderWilhelmsenBerg/Outlook-aligner`

Act as **Agent 20 — UX / UI / Visual Design** in **narrow Slice-A1.2 implementation-readiness mode** under current merged `AGENTS.md`.

Use the repository to communicate. Do not implement source code.

## Start

1. Refresh current `main` and record its exact HEAD.
2. Refresh `planning/segment-a-slices` and read `Slice-planning.md`, the A1.2 `00-brief.md`, and this prompt.
3. Read current `AGENTS.md`, `plan.md` A1.2, `docs/ui.md`, relevant phase docs, current WinUI `MainWindow`/view-model Calendar presentation, current `web/calendar` frontend, and visible-state tests where present.
4. Inspect the real current UI rather than relying on remembered discussion.

## UX decisions to make implementation-safe

Define the minimum truthful A1.2 UX contract for:

- month view as the primary Calendar surface;
- today / previous / next month controls and where ownership of those controls should live from a user-experience perspective;
- how the month title/current range is presented;
- what remains in the WinUI shell around the WebView;
- loading, empty, OutlookHost error, and partial-result/degraded states;
- calendar selection behavior at this slice, without inventing A1.3 logical-event correlation semantics;
- removal/hiding of Forward/Prepare diagnostic controls from the normal Segment A Calendar workflow without deleting dormant Segment-B spike code unnecessarily;
- accessibility: keyboard reachability/focus, high DPI/text scaling, light/dark theme, long subject/account text, and non-WebView fallback/status communication;
- what A1.2 must **not** visually imply before A1.3/A2 status/account semantics exist.

Do not design A1.3 one-logical-event projection, A2 account markers/health colors, reconciliation actions, or any write workflow.

## Repository reply

Write only:

`slice planning/A1.2/20-ux-note.md`

on branch `planning/segment-a-slices`.

Record exact `main` HEAD, exact planning-branch HEAD read, files inspected, required UX behavior, acceptance/manual checks, blockers, and final token `UX READY` only if Agent 40 can implement without inventing UX behavior.
