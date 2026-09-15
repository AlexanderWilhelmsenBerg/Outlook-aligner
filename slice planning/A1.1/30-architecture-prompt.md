# Agent 30 prompt — A1.1 architecture readiness

Repository: `AlexanderWilhelmsenBerg/Outlook-aligner`

Act as **Agent 30 — Architecture, Integration & Data Contracts** in **narrow Slice-A1.1 implementation-readiness mode** under the current merged `AGENTS.md`.

Use the repository to communicate. Do not implement source code and do not create the implementation PR.

## Start

1. Refresh current `main` and record its exact HEAD.
2. Refresh branch `planning/segment-a-slices` and read:
   - `slice planning/Slice-planning.md`;
   - `slice planning/A1.1/00-brief.md`;
   - this prompt.
3. Read current `AGENTS.md`, `plan.md` A1.1, `docs/phase-0.md`, `.github/workflows/ci.yml`, `.github/dependabot.yml`, the complete `web/calendar` tree, `web/calendar/package.json`, and relevant verification scripts.
4. Inspect actual current repository state rather than relying on the seed summary.

## Required architecture decisions

Define a precise A1.1 contract for Agent 40 covering:

- the Node/npm baseline that must generate the real lockfile;
- whether CI should pin an exact Node patch/npm version or another repository-consistent mechanism is sufficient;
- the real npm procedure that generates `package-lock.json` from the current manifest (never manually synthesize/fabricate lock data);
- the deterministic CI install command, including whether `--ignore-scripts` remains required with `npm ci`;
- preservation of typecheck, lint, formatting, build, and production audit gates;
- any narrow corresponding adjustment needed in local verification or Dependabot configuration;
- exact clean-checkout/local/CI evidence Agent 40 must provide;
- expected files to change where architecture can constrain them without dictating implementation mechanics.

## Hard scope boundaries

Do not:

- implement A1.2;
- change FullCalendar product behavior;
- introduce WebView2 message contracts;
- alter Core/App/OutlookHost DTO/domain behavior;
- productize Forward/Copy/Move or any write workflow;
- perform a dependency upgrade sweep;
- downgrade Vite because old phase documentation records 8.2.2;
- fabricate a lockfile.

Repository state outranks stale version documentation where an intentional merged dependency PR has changed the manifest.

## Repository reply

Write your result to:

`slice planning/A1.1/30-architecture-note.md`

Edit only that assigned note file on `planning/segment-a-slices`.

The note must include:

- exact `main` HEAD reviewed;
- exact planning-branch HEAD read before writing;
- files/contracts inspected;
- current frontend/toolchain facts;
- precise implementation contract;
- verification requirements;
- blockers, if any;
- final token `ARCH READY` only if Agent 40 can implement without making an architecture/toolchain decision.

Do not edit `Slice-planning.md`, Agent-00 files, or source code.
