# Phase 1 — Outlook COM Discovery / Read Probe

Status: **In progress 🚧 — pull request only**

Started: 2026-09-07

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

## Build artifact

A **successful** Phase 1 pull-request CI run publishes a self-contained Windows x64 single-file executable:

`OutlookAligner-Phase1-Probe-win-x64`

The artifact contains:

`OutlookAligner.OutlookHost.exe`

Failed upstream CI gates correctly prevent artifact publication.

The workflow uses `actions/upload-artifact@v7.0.1`, verified as the latest stable upload-artifact action on 2026-09-07.

## Download and run

1. Open the Phase 1 pull request on GitHub.
2. Open the successful `CI` workflow run.
3. Download the `OutlookAligner-Phase1-Probe-win-x64` artifact.
4. Extract the ZIP on the Windows PC that has Classic Outlook and the target Outlook profile configured.
5. Open PowerShell in the extracted folder.

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

Explicitly include subjects and locations for manual inspection:

```powershell
.\OutlookAligner.OutlookHost.exe --days 14 --include-details
```

Show usage without opening Outlook:

```powershell
.\OutlookAligner.OutlookHost.exe --help
```

## Expected result

For each configured account, the console should show:

- display name;
- SMTP address where Outlook exposes one;
- Outlook account type;
- delivery store;
- whether the default Calendar was available;
- number of calendar occurrences/items inside the requested horizon.

When `--include-details` is enabled, console event rows show both Subject and Location.

The JSON mode additionally exposes the extracted read-only DTO fields, including StoreID/EntryID locator data, GlobalAppointmentID where available, Start/End, recurrence state, busy state, and sensitivity.

## Manual acceptance test — merge gate

Hosted CI cannot validate the user's real Outlook profile. **Phase 1 must not be merged until this real Classic Outlook acceptance suite has passed.**

Run these on the Windows PC with the real three-account Classic Outlook profile:

- [ ] `--help` runs without starting/touching Outlook.
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

- [ ] formatter/analyzer gate passes;
- [ ] Release solution build passes;
- [ ] xUnit/MTP tests and coverage pass;
- [ ] benchmark project compiles;
- [ ] NuGet/npm audits pass;
- [ ] self-contained `win-x64` publish succeeds;
- [ ] published executable passes a `--help` smoke test on the hosted Windows runner;
- [ ] executable artifact is uploaded successfully.

Hosted tests cover CLI horizon bounds/errors, half-open filter semantics, `en-US` and `nb-NO` filter formatting, defensive JSON redaction, explicit detail output, and a Contracts-layer architecture guard against Outlook interop references.

## Known limitations in this phase

- Hosted GitHub runners cannot test live Outlook COM because they do not have the user's Outlook profile.
- Phase 1 does not yet correlate events across accounts; that belongs to the identity phase.
- Phase 1 does not test true meeting forwarding; that is Phase 2.
- It does not write custom Aligner metadata.
- `DateTime` local Start/End values are acceptable for this diagnostic phase, but the later identity/alignment design must preserve enough timezone/offset semantics for DST and cross-account comparisons.
- Subject/location output is diagnostic only and opt-in.

## Exit codes

- `0` — success/help.
- `2` — invalid command-line arguments.
- `3` — Outlook COM activation/read failure.
- `4` — unexpected probe failure.

## Handoff to Phase 2

Do not begin the forwarding spike until this PR has green hosted CI, the real three-account manual acceptance suite above has passed, and the user has explicitly merged Phase 1.
