# Segment A slice planning control

This folder is the repository-backed coordination space for Segment A planning.

Planning PRs are separate from implementation PRs. Planning may run one slice ahead, but implementation remains sequential unless the owner explicitly changes that rule.

## Current coordination state

Planning branch: `planning/a1.2-a1.3`

Planning PR: **pending creation**

Current authoritative `main`: `7dbd437c1ee86bf6e41d8ad939a54cc81ee7255e`.

Exact-head `main` CI: **success**, run `34961414443` (#212).

A1.1 was accepted post-merge by Agent 50. Its implementation is therefore the completed prerequisite for A1.2.

| Slice | Planning state | Required readiness owners | Implementation state |
| --- | --- | --- | --- |
| A1.1 — Deterministic calendar frontend toolchain | COMPLETE / ACCEPTED | Agent 30 + Agent 50 | Merged in `7dbd437c1ee86bf6e41d8ad939a54cc81ee7255e` |
| A1.2 — WebView2 month-calendar host | **SLICE READY** | Agents 20 + 30 — complete and revalidated | Agent 40 handoff next |
| A1.3 — One logical event, one calendar item | **PARALLEL READINESS ALLOWED** | Agents 10 + 30 | Blocked until A1.2 accepted/merged |

**Current implementation candidate:** A1.2.

**Current parallel planning candidate:** A1.3. Agents 10 and 30 may prepare A1.3 readiness while Agent 40 implements A1.2. Agent 00 must refresh `main` after A1.2 acceptance/merge and revalidate those notes before issuing A1.3 `SLICE READY`.

## Communication protocol

Repository state and merged `AGENTS.md` / `plan.md` are authoritative.

For each slice:

1. Agent 00 refreshes `main`, records the exact HEAD, and defines the slice boundary.
2. Required specialists review implementation readiness and write only their assigned `*-note.md` file.
3. Agent 00 reconciles all required inputs against current `main` and writes `00-reconciliation.md`.
4. Agent 00 issues `SLICE READY` only when Agent 40 can implement without inventing a domain, UX, product, or architecture rule.
5. Agent 40 implements the exact accepted contract on a separate branch/PR from current `main`.
6. While Agent 40 implements, specialists may plan the next slice in this planning PR.
7. Agent 50 independently reviews each implementation on its exact head and exact-head CI.
8. Visible/runtime slices retain the owner's manual acceptance gate.
9. The owner merges. No agent merges automatically.

## One-slice-ahead rule

Parallel specialist notes are provisional until the preceding implementation is accepted/merged. Agent 00 must then:

- refresh `main` and exact-head CI;
- inspect the merged diff for effects on the next slice;
- decide whether existing specialist notes remain valid;
- request a focused specialist refresh only if a material assumption changed;
- then reconcile and issue the next Agent-40 handoff.

## Branch ownership

- **Agent 00:** this control file, slice briefs, reconciliations, and Agent-40/50 prompts.
- **Agents 10/20/30:** only their assigned readiness note files unless Agent 00 explicitly requests otherwise.
- **Agent 40:** implementation branch/PR only.
- **Agent 50:** independent implementation acceptance.

## Segment A safety boundary

Segment A remains independently useful and read-only. Do not productize Forward testing, Copy Full / Copy Busy, Move execution, bulk reconciliation writes, recurring writes, deletion synchronization, or other Outlook mutations. Existing technical Forward spike code may remain dormant for Segment B.

## Status-token rule

Specialist tokens (`DOMAIN READY`, `UX READY`, `ARCH READY`) are evidence. Only Agent 00 writes `SLICE READY` after reconciliation against current authoritative state.
