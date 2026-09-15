# A1.3 — Agent 00 slice brief

## Slice

**A1.3 — One logical event, one calendar item**

Goal from `plan.md`: represent a supported non-recurring logical event once in Calendar rather than drawing one observation per account.

Required readiness owners: **Agent 10 — Outlook Domain, Identity & Calendar Semantics** and **Agent 30 — Architecture, Integration & Data Contracts**.

A1.3 implementation remains blocked until A1.2 is accepted/merged. Readiness planning may proceed in parallel while Agent 40 implements A1.2.

## Fixed contract from merged plan

- Create/standardize a presentation DTO for one logical event.
- Supported non-recurring items correlate through accepted identity rules.
- Duplicate, conflict, uncorrelated, suspicious-metadata, and recurrence-unresolved cases remain explicit instead of being guessed together.
- Selecting one calendar item updates the existing read-only comparison/detail model.
- Calendar and later Report consume the same logical-event model.
- Segment A remains read-only.

## Parallel-planning condition

Agent 10 and Agent 30 may complete readiness notes against current `main`. Their notes remain provisional until A1.2 is accepted/merged and Agent 00 verifies that the actual A1.2 implementation did not materially change their assumptions.

Agent 00 will not issue A1.3 `SLICE READY` before that boundary.
