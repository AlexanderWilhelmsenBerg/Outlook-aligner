# A1.1 — Agent 00 slice brief

## Slice

**A1.1 — Deterministic calendar frontend toolchain**

Goal from `plan.md`: make the existing TypeScript/FullCalendar frontend reproducible before expanding it.

Required readiness owner: **Agent 30 — Architecture, Integration & Data Contracts**.

Implementation/review owners after readiness: Agent 40 / Agent 50.

## Authoritative baseline at planning-workspace creation

`main`: `e8bc60cd6e2acc1bc32e538d661db88262198226`

Open implementation PRs: none at creation time.

The WinUI/read-alignment foundation is merged through PR #6. PR #7 subsequently updated Vite to `8.3.0`. PR #8 merged `AGENTS.md` and the Segment A slice plan.

Observed frontend state at that baseline:

- `web/calendar/package.json` exists;
- no `package-lock.json` is committed;
- `package.json` pins `fullcalendar` 7.1.0, `temporal-polyfill` 1.0.4, ESLint 10.10.0, Prettier 3.9.6, TypeScript 7.0.2, and Vite 8.3.0;
- `package.json` currently declares `node >=24.0.0`;
- CI uses `actions/setup-node@v7.0.0` with `node-version: "24"`;
- CI currently runs `npm install --ignore-scripts`;
- CI preserves typecheck, lint, format check, production build, and production dependency audit;
- `docs/phase-0.md` records Node 24.20.0 LTS / npm 11.19.0 and explicitly requires the first material frontend PR to generate a real lockfile and switch to `npm ci`;
- the phase-0 Vite table still says 8.2.2, but merged repository state now pins Vite 8.3.0. A1.1 must not opportunistically downgrade or upgrade the stack.

## Scope boundary

A1.1 is infrastructure only. It must not implement A1.2 month-calendar hosting, alter product UI behavior, introduce WebView2 message contracts, change Outlook DTO/domain semantics, or touch Segment B write workflows.

## Readiness gate

Agent 30 must resolve the exact lockfile/toolchain/CI contract in `30-architecture-note.md`.

Agent 00 will then re-read current `main`, reconcile the note in `00-reconciliation.md`, and only then create the Agent-40 handoff.
