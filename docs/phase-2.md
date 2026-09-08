# Phase 2 — Native Meeting Forwarding Technical Spike

Status: **In progress 🚧 — PR #5 only**

Started: 2026-09-08

## Purpose

Determine whether Outlook Aligner can honestly provide **genuine Classic Outlook meeting forwarding** for accepted meetings.

The product must never label `AppointmentItem.ForwardAsVcal()` as equivalent native forwarding.

## Outlook facts driving the spike

Microsoft documents that:

- `MeetingItem.Forward()` executes Outlook's native Forward action and returns a new `MeetingItem`;
- `MeetingItem.GetAssociatedAppointment(false)` maps a retained meeting request back to its calendar appointment without adding an item;
- an accepted calendar meeting is represented as `AppointmentItem`;
- `AppointmentItem` exposes a `Forward` event when the user invokes Outlook's Forward action, and that event supplies the new forwarded object plus a cancellable `Cancel` flag;
- setting `Cancel = true` in `AppointmentItem.Forward` prevents the Forward operation from completing and prevents the new item from being displayed;
- Classic Outlook itself can forward an accepted meeting directly from Calendar;
- `CommandBars.ExecuteMso(idMso)` can invoke a built-in Office command when the object model has no direct method.

Primary references:

- https://learn.microsoft.com/en-us/office/vba/api/outlook.meetingitem.forward%28method%29
- https://learn.microsoft.com/en-us/office/vba/api/outlook.meetingitem.getassociatedappointment
- https://learn.microsoft.com/en-us/office/vba/api/outlook.appointmentitem.forward
- https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.itemevents_10_forwardeventhandler
- https://learn.microsoft.com/en-us/office/vba/api/outlook.appointmentitem.forwardasvcal
- https://learn.microsoft.com/en-us/office/vba/api/office.commandbars.executemso
- https://support.microsoft.com/en-us/outlook/calendar/forward-a-meeting-in-outlook
- https://github.com/OfficeDev/office-fluent-ui-command-identifiers

## Path A — retained MeetingItem recovery

The first implementation selects a source account by SMTP address, then searches that delivery store's default Inbox and Deleted Items for `IPM.Schedule.Meeting.Request` items.

Each candidate is correlated to its appointment with `GetAssociatedAppointment(false)` and compared using `GlobalAppointmentID`.

### Safety modes

#### Inspect — default/read-only

```powershell
.\OutlookAligner.OutlookHost.exe --forward-spike `
  --source-smtp source@example.com `
  --global-id GLOBAL_ID
```

Does not call `Forward()`, create, save, send, move, or delete anything.

#### Prepare — native forward, discarded unsent

```powershell
.\OutlookAligner.OutlookHost.exe --forward-spike `
  --source-smtp source@example.com `
  --global-id GLOBAL_ID `
  --prepare `
  --to target@example.com
```

Requires exactly one native retained request. Calls `MeetingItem.Forward()`, resolves the target, selects the source Outlook account, then closes/discards the unsent forward.

#### Send — explicit side effect

```powershell
.\OutlookAligner.OutlookHost.exe --forward-spike `
  --source-smtp source@example.com `
  --global-id GLOBAL_ID `
  --send `
  --to target@example.com `
  --confirm-send SEND-NATIVE-MEETING
```

Only this exact gated shape reaches `MeetingItem.Send()`.

## Real-machine result #1 — retained request unavailable

Tested 2026-09-08 against a real accepted Kverneland Group meeting on the user's Classic Outlook profile.

Source account:

`Alexander.Berg@kvernelandgroup.com`

Calendar `GlobalAppointmentID`:

`040000008200E00074C5B7101A82E00800000000D05887327B3BDD01000000000000000010000000EB2313460E921449BF83A27C3A21C827`

Observed result:

- source Outlook account selected successfully;
- Inbox searched;
- Deleted Items searched;
- matching retained meeting requests: **0**;
- no vCalendar fallback attempted;
- no Outlook data modified.

### Conclusion from result #1

The hypothesis that Outlook Aligner can **always** recover the original/retained `MeetingItem` after a meeting has been accepted is disproven.

This does not prove genuine forwarding is impossible. Classic Outlook's Calendar UI still exposes Forward for accepted `AppointmentItem` objects, so Phase 2 has a second technical path.

## Path B — accepted AppointmentItem / Outlook Calendar Forward command

### B1. Capability probe — no Forward execution

Command:

```powershell
.\OutlookAligner.OutlookHost.exe --forward-spike `
  --source-smtp source@example.com `
  --global-id GLOBAL_ID `
  --probe-calendar-command `
  --entry-id ENTRY_ID
```

For the exact accepted Calendar `AppointmentItem`, the probe:

1. reopens it by EntryID in the selected source Store;
2. verifies the requested `GlobalAppointmentID`;
3. verifies it is a meeting;
4. opens an Outlook Inspector context;
5. queries Microsoft's built-in Fluent command identifier `Forward`;
6. reports identifier validity, label, visibility and enabled state;
7. closes the Inspector without executing Forward or modifying the appointment.

Late-bound Office command access is intentional so OutlookHost does not reintroduce a runtime dependency on `office.dll`.

## Real-machine result #2 — Calendar native Forward capability present

Tested 2026-09-08 against the **same accepted meeting** used for result #1.

Observed:

- `idMso`: `Forward`;
- identifier valid: **True**;
- label: `Forward`;
- visible: **True**;
- enabled: **True**;
- command not executed;
- verified subject with explicit diagnostic opt-in: `List of application`.

### Conclusion from result #2

Classic Outlook exposes its built-in native Forward action for this accepted Calendar appointment even though no retained matching `MeetingItem` can be recovered from Inbox or Deleted Items.

This makes the Calendar-command path a serious candidate for the product's general Forward mechanism. It still requires a controlled prepare/cancel proof before any recipient or send test.

### B2. Prepare/cancel — command executes, Forward event cancels before completion

New diagnostic mode:

```powershell
.\OutlookAligner.OutlookHost.exe --forward-spike `
  --source-smtp source@example.com `
  --global-id GLOBAL_ID `
  --prepare-calendar-command `
  --entry-id ENTRY_ID
```

Safety contract:

1. reopen exact `AppointmentItem` by EntryID in the selected source Store;
2. verify `GlobalAppointmentID` and meeting state;
3. verify `Forward` remains valid, visible and enabled;
4. subscribe to `AppointmentItem.Forward`;
5. invoke only the built-in `Forward` command with `ExecuteMso("Forward")`;
6. when Outlook raises `AppointmentItem.Forward`, immediately set `Cancel = true`;
7. inspect only the transient new object's type/message class to determine whether Outlook supplied a native `MeetingItem`;
8. do **not** add a recipient;
9. do **not** call `Send()`, `Save()`, or create a replacement appointment;
10. detach the event handler and close the source Inspector.

The CLI deliberately rejects `--to` in this mode. If the event is not raised, cancellation is not observed, or the new item is not a native `MeetingItem`, the experiment fails closed.

### B3. Recipient/send — only after B2 succeeds

Only after B2 proves a safely capturable native `MeetingItem` should the spike add a separate recipient/send experiment.

That later step must use another account controlled by the user, require an explicit confirmation token, and verify genuine meeting behavior at the target rather than a generic email/ICS attachment.

## Positive-control test for Path A

Path A should still be sampled against one recent accepted meeting whose original invitation is visibly retained in Inbox.

Possible outcomes:

- **1 native match:** Path A remains conditionally useful/direct;
- **0 matches despite the visible request:** the current request-correlation implementation needs repair before judging Path A.

This positive control is useful evidence but no longer blocks testing Path B because result #2 independently proves the Calendar native Forward capability exists.

## Privacy

The spike does not read/print meeting body, attendees, attachments, Teams URLs, or location by default.

Subject remains explicit `--include-details` diagnostic opt-in.

Errors use operation/HRESULT/type context rather than raw message contents.

## Current decision model

Phase 2 may end in one of four truthful outcomes:

1. **Reliable native forwarding** — safe native route works broadly; expose Forward normally.
2. **Conditional native forwarding** — expose Forward only when per-event capability proves it is available.
3. **Outlook-UI-command forwarding** — retained request is unreliable but the supported Calendar command/event route is stable enough to automate; expose it with explicit capability checks.
4. **Not reliable enough** — omit Forward from the product and use Copy Full / Copy Busy instead.

At no point does `ForwardAsVcal()` become a silent fallback for the Forward action.

## Manual merge gate

Before PR #5 may merge:

### Retained-request path

- [x] zero-match real-world result recorded for an accepted meeting;
- [ ] positive control with invitation visibly retained in Inbox;
- [ ] if exactly one match exists, retained-request Prepare succeeds and discards unsent;
- [ ] retained-request send is optional if the Calendar-command route proves superior and reliable.

### Calendar-command path

- [x] exact Microsoft Fluent command identifier `Forward` confirmed;
- [x] capability probe reports valid/visible/enabled state without executing Forward;
- [ ] prepare/cancel invocation succeeds;
- [ ] `AppointmentItem.Forward` produces a native `MeetingItem`;
- [ ] `Cancel = true` prevents the forward from completing/displaying;
- [ ] no draft remains after the experiment;
- [ ] source appointment remains unchanged;
- [ ] one explicitly confirmed recipient/send test occurs only after prepare/cancel passes;
- [ ] target receives genuine meeting behavior and Teams join remains usable where applicable.

### General

- [ ] repeat against several meeting types/accounts;
- [ ] organizer-disabled forwarding is handled as unavailable, not an error-prone workaround;
- [ ] no unintended recipients or calendar writes occur;
- [ ] hosted CI is green on final PR head.

## Hosted CI

Hosted runners cannot prove live Outlook/Exchange forwarding but must continue to prove:

- formatter/analyzers;
- Release solution build;
- xUnit/MTP tests + coverage;
- benchmark compile;
- NuGet/npm audits;
- self-contained win-x64 publish;
- published EXE `--help`;
- published EXE `--interop-check`;
- published EXE `--forward-spike --help` without opening Outlook;
- Phase 2 artifact upload.

## Exit codes

- `0` — requested operation completed successfully.
- `2` — invalid command line.
- `3` — Outlook COM failure.
- `4` — unexpected/dependency failure.
- `5` — source/native selection, command capability, event capture, or correlation could not be resolved safely.
- `6` — recipient could not be resolved; unsent forward discarded.

## Explicit non-goals

Phase 2 does not:

- create a replacement organizer-owned meeting;
- use `AppointmentItem.ForwardAsVcal()` as native Forward;
- copy Teams/body content manually to fake forwarding;
- implement the production GUI;
- implement Copy Full / Copy Busy;
- merge automatically.

The production GUI is nevertheless a required v1 surface and is specified separately in [`docs/ui.md`](ui.md).
