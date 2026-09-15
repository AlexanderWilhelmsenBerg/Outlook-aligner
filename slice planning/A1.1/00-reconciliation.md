# Agent 00 reconciliation — A1.1

Status: **SLICE READY**

## Authoritative state

- `main` HEAD refreshed and reconciled: `e8bc60cd6e2acc1bc32e538d661db88262198226`
- Exact-head `main` CI: **success**, run `34943491782`
- Planning PR: #9
- Planning branch HEAD before this reconciliation: `ba9ec5447476809dd42a5bdea3ad41e88c42bea2`
- Required specialist input: Agent 30 architecture note, committed in `d4c29dac5fc2c1362cda903943e8ad46d1370354`, final token `ARCH READY`
- No Agent 10 or Agent 20 readiness is required for A1.1 because this slice changes deterministic frontend tooling only and introduces no domain or visible UX behavior.

The repository has not moved since Agent 30 reviewed `main`, so its architecture findings remain current. The merged WinUI/read-alignment foundation is present and A1.1 is the first incomplete Segment A slice.

## Reconciled implementation contract

Agent 40 may implement **A1.1 — Deterministic calendar frontend toolchain** with the following fixed contract:

1. Generate a real `web/calendar/package-lock.json` using **Node 24.20.0 / npm 11.19.0** from the current `web/calendar/package.json`. Never synthesize or manually fabricate lock data.
2. Preserve the current declared dependency versions. In particular keep Vite `8.3.0`; do not perform an upgrade/downgrade sweep.
3. Pin the frontend CI runtime to Node `24.20.0`, expose/verify the effective Node/npm versions, and treat a mismatch from npm `11.19.0` as a blocker rather than silently changing the baseline.
4. Replace frontend verification installation with `npm ci --ignore-scripts` in hosted CI.
5. Update `scripts/verify.ps1` narrowly from `npm install --ignore-scripts` to `npm ci --ignore-scripts`.
6. Preserve the existing frontend quality gates: typecheck, lint, Prettier check, production build, and production dependency audit. Preserve existing .NET/publish/smoke verification.
7. No Dependabot policy change is required for this slice.
8. Expected implementation diff is normally limited to:
   - new `web/calendar/package-lock.json`;
   - `.github/workflows/ci.yml`;
   - `scripts/verify.ps1`.
   Any additional file requires a concrete A1.1 reproducibility reason.

## Required implementation evidence

Agent 40 must record and verify on the exact implementation head:

- `node --version` = `v24.20.0` and `npm --version` = `11.19.0` for lockfile generation;
- the real npm command used to generate the lockfile;
- no opportunistic declared dependency-version changes;
- clean `npm ci --ignore-scripts` from `web/calendar`;
- a repeated clean install does not modify `package.json` or `package-lock.json`;
- `npm run typecheck`, `npm run lint`, `npm run format:check`, `npm run build`, and `npm audit --omit=dev --audit-level=high` succeed;
- `scripts/verify.ps1` succeeds under the agreed baseline;
- clean-checkout verification succeeds without first running `npm install`;
- exact-head hosted CI is green, including the existing .NET job.

If the exact documented Node distribution does not provide npm `11.19.0`, if `npm ci --ignore-scripts` proves incompatible with the current dependency graph, or if locking the current graph introduces a high/critical production audit failure, Agent 40 must stop and report the concrete evidence rather than changing the architecture contract ad hoc.

## Explicit exclusions

A1.1 must not implement A1.2 month-calendar hosting, alter FullCalendar product behavior, add WebView2 contracts, change App/Core/OutlookHost domain behavior, expose technical Forward functionality as a product workflow, or introduce any Segment B write capability.

## Blockers

None in current authoritative state.

**SLICE READY**
