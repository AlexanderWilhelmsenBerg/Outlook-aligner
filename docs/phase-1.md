# Phase 1 — Outlook COM Discovery / Read Probe

Status: **Complete ✅ — merged as PR #4 on 2026-09-08**

Started: 2026-09-07  
Merged: 2026-09-08  
Merge commit: `3175f1964b3980c63419b8bc03fa38eb3cd9f99f`

## Completion evidence

Phase 1 was validated in two layers:

1. Hosted CI passed the build/test/publish gates.
2. The first real-machine run exposed a missing Office interop packaging dependency that hosted `--help` testing could not detect. The PR was repaired to embed the resolved Outlook interop metadata and add a published-EXE `--interop-check`. The user then confirmed the real 14-day Classic Outlook probe worked and explicitly merged PR #4.

Final pre-merge CI evidence:

- run: `34121178052`;
- head: `e0acfc95bba4af2b766f92d063f43eae97440090`;
- artifact: `OutlookAligner-Phase1-Probe-win-x64`;
- artifact ID: `10018308723`;
- SHA-256: `1b955e2d25fef3451a571a4a8f3ba181363f6380ca5303ca2e703d4811a13bd3`.

The user did not separately report every scenario in the longer manual checklist below, so those boxes remain historical test guidance rather than fabricated completion evidence.

## Purpose

Phase 1 proves the real Classic Outlook COM boundary before any synchronization or forwarding logic is added. The output is intentionally a diagnostic command-line executable rather than the production UI.

This phase is **strictly read-only**. It contains no Outlook `Save`, `Send`, `Delete`, forwarding, item creation, custom-property mutation, or calendar movement operations.

## What the probe does

- activates/attaches to Classic Outlook through `Outlook.Application` COM automation;
- gets the current MAPI namespace/profile;
- enumerates Outlook stores;
- enumerates configured accounts;
- resolves each account's delivery store and default Calendar folder;
- scans from today for a bounded configurable number of days;
- sorts Calendar items by Start, enables recurrence expansion, then applies a bounded restriction;
- uses a half-open overlap filter: event end must be strictly after the window start and event start must be before the window end;
- enumerates the restricted results without relying on recurrence-expanded `Items.Count`;
- extracts plain DTOs containing Outlook locator/identity/time/state information;
- explicitly releases short-lived COM references;
- releases its Outlook Application RCW without calling `Application.Quit()`.

## Privacy behavior

Default console and JSON output withholds event subject and location. Body, attendees, email contents, attachments, and Teams URLs are not extracted by the Phase 1 probe.

Event subjects/locations are shown only when `--include-details` is explicitly supplied. JSON output also performs defensive redaction at the output boundary when details are not enabled, even if a future producer accidentally supplies Subject or Location values in a DTO.

Unhandled non-COM failures report the exception type only; raw exception messages are not printed by default because future exception text could contain sensitive data.

## Packaging lesson from real-machine testing

The original self-contained single-file artifact built and passed `--help` in hosted CI but failed on the user's machine at the first live Outlook path with `FileNotFoundException`.

Root cause: the Outlook PIA was consumed through `PackageReference`, so interop metadata was not embedded automatically. The repaired build marks the resolved `Microsoft.Office.Interop.Outlook` reference with `EmbedInteropTypes=true` after `ResolveReferences`.

The final CI smoke test runs both:

```powershell
.\OutlookAligner.OutlookHost.exe --help
.\OutlookAligner.OutlookHost.exe --interop-check
```

`--interop-check` resolves Outlook interop metadata without opening Outlook, covering the packaging failure that the original help-only smoke test missed.

## Usage

Default 90-day privacy-safe scan:

```powershell
.\OutlookAligner.OutlookHost.exe
```

Shorter scan:

```powershell
.\OutlookAligner.OutlookHost.exe --days 14
```

Privacy-safe JSON:

```powershell
.\OutlookAligner.OutlookHost.exe --days 90 --json
```

Explicitly include subjects and locations:

```powershell
.\OutlookAligner.OutlookHost.exe --days 14 --include-details
```

Validate packaged interop metadata without opening Outlook:

```powershell
.\OutlookAligner.OutlookHost.exe --interop-check
```

## Expected result

For each configured account, the console shows:

- display name;
- SMTP address where Outlook exposes one;
- Outlook account type;
- delivery store;
- whether the default Calendar was available;
- number of calendar occurrences/items inside the requested horizon.

When `--include-details` is enabled, console event rows show both Subject and Location.

The JSON mode additionally exposes the extracted read-only DTO fields, including StoreID/EntryID locator data, GlobalAppointmentID where available, Start/End, recurrence state, busy state, and sensitivity.

## Historical/manual acceptance checklist

These scenarios remain useful regression tests even though only the successful live probe itself was explicitly reported before merge:

- [ ] `--help` runs without starting/touching Outlook.
- [x] `--interop-check` succeeds in hosted CI and the repaired package runs on the user's machine.
- [x] A real 14-day Classic Outlook probe completes successfully on the user's machine.
- [ ] With Classic Outlook already running, the default probe completes successfully.
- [ ] With Classic Outlook completely exited, the probe cold-starts/initializes the default profile and completes successfully.
- [ ] All three expected accounts are listed.
- [ ] Each expected delivery store/default Calendar is found.
- [ ] A 14-day scan contains only events intersecting that bounded range.
- [ ] An event ending exactly at the scan-window start is excluded.
- [ ] A long-running event that starts before the window but ends inside/after it is included.
- [ ] An all-day event exactly on a horizon boundary is handled correctly.
- [ ] A 90-day scan contains expected recurring meetings but does not run unbounded.
- [ ] A recurrence with no end date reports only occurrences inside the requested horizon.
- [ ] A modified recurrence exception is read without becoming an unbounded scan.
- [ ] An all-day event is represented with `IsAllDay=true` in JSON.
- [ ] At least one meeting exposes `GlobalAppointmentId` where expected.
- [ ] A scan spanning a DST transition completes with plausible local Start/End values.
- [ ] A Teams meeting can be scanned while body, attendees, attachments, and Teams URL remain absent from the Phase 1 DTO/output.
- [ ] Default console output does not print meeting subjects or locations.
- [ ] Default JSON output does not print meeting subjects or locations.
- [ ] `--include-details` prints both subjects and locations only when explicitly requested.
- [ ] Run the probe 10 times consecutively; Outlook remains responsive and event counts remain plausible.
- [ ] Close/reopen Classic Outlook and rerun the probe successfully.
- [ ] An unavailable/inaccessible store is reported as an error/warning without aborting all other accounts.
- [ ] No calendar item is created, modified, forwarded, moved, saved, or deleted.

## Hosted CI acceptance

Final Phase 1 head completed all hosted gates successfully:

- [x] formatter/analyzer gate passes;
- [x] Release solution build passes;
- [x] xUnit/MTP tests and coverage pass;
- [x] benchmark project compiles;
- [x] NuGet/npm audits pass;
- [x] self-contained `win-x64` publish succeeds;
- [x] published executable passes `--help`;
- [x] published executable passes `--interop-check`;
- [x] executable artifact is uploaded successfully.

Hosted tests cover CLI horizon bounds/errors, half-open filter semantics, `en-US` and `nb-NO` filter formatting, defensive JSON redaction, explicit detail output, an Outlook interop metadata check, and a Contracts-layer architecture guard against Outlook interop references.

## Known limitations carried forward

- Hosted GitHub runners cannot test live Outlook COM because they do not have the user's Outlook profile.
- Phase 1 does not correlate events across accounts; that belongs to the identity phase.
- Phase 1 does not test true meeting forwarding; that is Phase 2.
- It does not write custom Aligner metadata.
- `DateTime` local Start/End values are acceptable for this diagnostic phase, but later identity/alignment design must preserve enough timezone/offset semantics for DST and cross-account comparisons.

## Exit codes

- `0` — success/help.
- `2` — invalid command-line arguments.
- `3` — Outlook COM activation/read failure.
- `4` — unexpected/dependency failure.

## Handoff to Phase 2

Phase 1 is closed. Phase 2 may now test genuine `MeetingItem.Forward()` behavior in a separate PR. The vCalendar path remains explicitly non-equivalent and must not be mislabeled as native forwarding.
