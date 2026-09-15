# Agent 30 prompt — A1.2 architecture readiness

Repository: `AlexanderWilhelmsenBerg/Outlook-aligner`

Act as **Agent 30 — Architecture, Integration & Data Contracts** in **narrow Slice-A1.2 implementation-readiness mode** under current merged `AGENTS.md`.

Use the repository to communicate. Do not implement source code.

## Start

1. Refresh current `main` and record its exact HEAD.
2. Refresh `planning/segment-a-slices` and read `Slice-planning.md`, the A1.2 `00-brief.md`, and this prompt.
3. Read current `AGENTS.md`, `plan.md` A1.2, `docs/ui.md`, relevant phase/architecture docs, App/Core/OutlookHost boundaries, `MainWindow`, current WebView2/FullCalendar scaffold, frontend build integration, and relevant tests.
4. Treat A1.1 as a prerequisite that may still be under implementation. Do not assume unmerged A1.1 code exists.

## Architecture decisions to make implementation-safe

Define the A1.2 contract for:

- ownership and lifecycle of the WebView2 month-calendar surface inside the existing WinUI shell;
- how built frontend assets are located/loaded in dev/test/published app flows;
- the smallest versioned/plain message or DTO contract needed to render current bounded calendar observations, without moving correlation/status business rules into TypeScript;
- initialization order and handling when WebView2 or frontend assets fail;
- preserving visible loading, empty, OutlookHost error, and partial-result states outside/around the WebView;
- navigation events (today/prev/next/range change), selection, and any WinUI↔WebView message direction required by A1.2 only;
- test seams that do not require Outlook COM for presentation tests;
- packaging/build/CI implications, including how A1.1's deterministic frontend output should be consumed once merged;
- the exact boundary preventing COM RCWs or Segment-B action semantics from crossing into WebView2;
- performance/reload expectations sufficient for the bounded month-calendar use case.

Do not invent A1.3 correlation semantics, A2 status logic, persistence not required by A1.2, or production IPC work beyond what this slice demonstrably needs.

## Repository reply

Write only:

`slice planning/A1.2/30-architecture-note.md`

on branch `planning/segment-a-slices`.

Record exact `main` HEAD, exact planning-branch HEAD read, files/contracts inspected, precise architecture/data-contract decisions, tests/evidence, blockers, and final token `ARCH READY` only if Agent 40 can implement without making an architecture decision.
