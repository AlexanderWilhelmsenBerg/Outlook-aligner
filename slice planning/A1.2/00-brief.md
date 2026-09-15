# A1.2 — Agent 00 slice brief

## Slice

**A1.2 — WebView2 month-calendar host**

Goal from `plan.md`: replace the list-first Calendar experience with a real month calendar surface while preserving the WinUI shell and OutlookHost boundary.

Required readiness owners: **Agent 20 — UX/UI** and **Agent 30 — Architecture**.

Implementation remains blocked until A1.1 is accepted and merged. Readiness planning may proceed in parallel.

## Fixed product contract from merged plan

- Calendar opens in month view.
- Today and previous/next month navigation work.
- WinUI sends plain read-model DTOs into WebView2; the WebView never receives COM objects.
- FullCalendar is presentation only; correlation/status rules stay in Core/App read models.
- Loading, empty, host-error, and partial-result states remain visible outside/around the WebView.
- Existing Forward/Prepare diagnostic code is not expanded and is removed from the normal Calendar workflow for Segment A.
- Acceptance is a real-profile month grid rendering bounded calendar data without PowerShell or raw Outlook IDs.

## Parallel-planning condition

Agent 20 and Agent 30 may complete their readiness notes while A1.1 is under implementation. Their notes are provisional evidence until Agent 00 refreshes `main` after A1.1 merge and checks whether the merged toolchain work materially changes A1.2 assumptions.

Agent 00 will not issue A1.2 `SLICE READY` before the A1.1 acceptance/merge boundary.
