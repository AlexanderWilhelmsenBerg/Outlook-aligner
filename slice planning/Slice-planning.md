# Segment A slice planning control

This folder is the repository-backed coordination space for Segment A planning.

The planning PR is deliberately separate from implementation PRs. It may remain open while individual slice implementation PRs are reviewed and merged.

## Current coordination state

Planning branch: `planning/segment-a-slices`

Planning PR: **#9 — Coordinate Segment A slice planning**

Planning base when this workspace was created: `main` at `e8bc60cd6e2acc1bc32e538d661db88262198226`.

| Slice | Planning state | Required readiness owners | Implementation state |
| --- | --- | --- | --- |
| A1.1 — Deterministic calendar frontend toolchain | **SLICE READY** | Agent 30 — complete | **Agent 40 handoff issued** |
| A1.2 — WebView2 month-calendar host | **PARALLEL READINESS COMPLETE** | Agents 20 + 30 — complete | Blocked until A1.1 accepted/merged and revalidated |
| A1.3 — One logical event, one calendar item | QUEUED | Agents 10 + 30 | Blocked |

**Current implementation candidate:** A1.1. The accepted reconciliation is immutable planning commit `e518fa85b771e69f6be2f8d9d4ff05ef685b3fab`; the Agent-40 handoff is in `A1.1/40-implementation-prompt.md`.

**Current parallel planning candidate:** A1.2. Both specialist notes are present with `UX READY` / `ARCH READY`. They remain provisional until A1.1 is accepted/merged; Agent 00 must then refresh `main`, inspect the merged A1.1 diff, and revalidate or refresh the A1.2 notes before final reconciliation.

## Communication protocol

Repository state and merged `AGENTS.md` / `plan.md` remain authoritative. This planning PR records specialist evidence and Agent-00 reconciliation; it does not override merged contracts by itself.

For each slice:

1. Agent 00 refreshes `main`, records the exact HEAD, creates/updates the slice brief and specialist prompts.
2. Each required specialist reads the current repository plus its prompt and writes only its assigned `*-note.md` file on this planning branch.
3. Each specialist note records:
   - exact `main` HEAD reviewed;
   - exact planning-branch HEAD read before writing;
   - files/contracts inspected;
   - decisions and blockers;
   - the role-specific readiness token required by `AGENTS.md` when ready.
4. Specialists do not implement source code during readiness review and do not edit another agent's note.
5. Agent 00 reads all required notes, refreshes authoritative state, resolves conflicts, and writes `00-reconciliation.md`.
6. Agent 00 issues `SLICE READY` only if Agent 40 can implement without inventing domain, UX, product, or architecture rules.
7. Agent 00 then writes the exact `40-implementation-prompt.md`, including the immutable planning commit containing the accepted reconciliation.
8. Agent 40 implements on a separate branch/PR based on current `main`. Agent 40 does not implement on this planning branch.
9. As soon as Agent 40 begins the current slice, specialists may plan the next slice in this same planning PR. This is readiness planning only; implementation remains sequential unless the owner explicitly changes that rule.
10. After implementation, Agent 00 writes an exact-head `50-acceptance-prompt.md`. Agent 50 independently reviews the implementation PR and exact-head CI.
11. The owner retains the manual merge gate. No agent merges automatically.

## Parallel-planning rule

Planning may run **one slice ahead** of implementation to keep the pipeline moving. The next slice cannot receive final `SLICE READY` merely from an older specialist note. After the previous slice is merged, Agent 00 must:

- refresh `main` and record its new exact HEAD;
- inspect the merged diff for effects on the next slice;
- determine whether existing specialist notes remain valid;
- request a specialist refresh if the merged change materially affects that specialist's contract;
- only then write/finalize the next reconciliation and Agent-40 handoff.

This keeps planning parallel without letting implementation contracts drift away from the repository.

## Branch ownership

- **Agent 00:** this control file, slice briefs, `00-reconciliation.md`, Agent-40 and Agent-50 prompts.
- **Agent 10/20/30:** only their assigned specialist note files, unless Agent 00 explicitly asks for a prompt correction.
- **Agent 40:** implementation branch/PR only; may read this branch but should not use it for source implementation.
- **Agent 50:** implementation review; acceptance result may be recorded in the slice folder after review.

## Segment A safety boundary

Segment A remains independently useful and read-only. Planning and implementation must not smuggle in Forward productization/testing as a user workflow, Copy Full / Copy Busy, Move execution, bulk reconciliation writes, recurring writes, or deletion synchronization. Existing technical Forward spike code may remain dormant for Segment B.

## Status-token rule

A specialist token (`DOMAIN READY`, `UX READY`, `ARCH READY`) is evidence, not the implementation gate. Only Agent 00 writes `SLICE READY`, and only after all required readiness inputs are reconciled against current authoritative state.
