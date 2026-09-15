# Agent 30 prompt — A1.3 architecture readiness

Repository: `AlexanderWilhelmsenBerg/Outlook-aligner`

Act as **Agent 30 — Architecture / Integration & Data Contracts**.

Work on the current Segment-A planning PR only. Follow current `AGENTS.md` completely.

Read `plan.md` A1.3, `slice planning/A1.3/00-brief.md`, Agent 10's note when available, current App/Core logical/read models, A1.2 presentation contracts, OutlookHost DTOs, Calendar/Report seams, and relevant tests.

Define the smallest shared logical-event read-model/data-contract architecture that lets Calendar now and Report later consume the same truth without moving correlation into TypeScript or leaking COM/raw Outlook locators into presentation.

Cover ownership/layer placement, mapping from observations to logical-event projections, explicit unresolved cases, selection/detail wiring, test seams, and migration from A1.2 per-observation rendering.

Do not implement source code and do not invent domain identity rules; Agent 10 owns those.

Write only:

`slice planning/A1.3/30-architecture-note.md`

Record exact `main` HEAD, planning HEAD read, evidence inspected, precise architecture/tests/blockers, and end with `ARCH READY` only if implementation needs no invented architecture rule.
