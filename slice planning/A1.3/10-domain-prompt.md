# Agent 10 prompt — A1.3 domain readiness

Repository: `AlexanderWilhelmsenBerg/Outlook-aligner`

Act as **Agent 10 — Outlook Domain, Identity & Calendar Semantics**.

Work on the current Segment-A planning PR only. Follow current `AGENTS.md` completely.

Read `plan.md` A1.3, `slice planning/A1.3/00-brief.md`, current identity/correlation docs and code, OutlookHost calendar DTOs, managed metadata handling, recurrence handling, and relevant tests.

Define the exact domain contract Agent 40 will need for supported **non-recurring** logical-event correlation, including positive identity evidence, fail-closed duplicate/conflict/uncorrelated/suspicious cases, partial-read behavior, and what recurrence-related cases must remain unresolved in A1.3.

Do not implement source code and do not design A2 authority/status rules.

Write only:

`slice planning/A1.3/10-domain-note.md`

Record exact `main` HEAD, planning HEAD read, evidence inspected, precise rules/tests/blockers, and end with `DOMAIN READY` only if implementation needs no invented domain rule.
