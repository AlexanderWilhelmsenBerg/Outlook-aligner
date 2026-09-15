# Agent 40 implementation prompt — A1.1

Repository: `AlexanderWilhelmsenBerg/Outlook-aligner`

Act as **Agent 40 — Coder / Implementation Owner** in build mode under current `AGENTS.md`.

Implement **A1.1 — Deterministic calendar frontend toolchain** only.

## Authoritative planning contract

The accepted Agent-00 reconciliation is immutable planning commit:

`e518fa85b771e69f6be2f8d9d4ff05ef685b3fab`

Read from that commit:

- `slice planning/A1.1/00-reconciliation.md`
- `slice planning/A1.1/30-architecture-note.md`

Also read current merged `AGENTS.md`, `plan.md` A1.1, and the actual current toolchain/CI files.

## Start

Refresh `main`, record its exact HEAD, and inspect open PRs before editing. Repository state outranks remembered discussion. If `main` materially changed after the reconciliation, stop and report the conflict to Agent 00 rather than silently adapting the contract.

Implement on a **separate branch and implementation PR from current `main`**. Do not implement source changes on planning PR #9.

Follow the reconciled contract exactly: generate the real npm lockfile with the agreed Node/npm baseline, move deterministic verification to `npm ci --ignore-scripts`, preserve all quality gates, and avoid dependency/version or product/UI changes outside A1.1.

Run and record all evidence required by `00-reconciliation.md`, including clean-checkout/repeated-install proof and exact-head CI.

Do not implement A1.2 and do not merge the PR.

When finished, return:

- exact base `main` HEAD;
- implementation PR number;
- exact implementation PR HEAD;
- files changed;
- verification results;
- exact-head CI/check state;
- any blocker or deviation from the immutable reconciliation.
