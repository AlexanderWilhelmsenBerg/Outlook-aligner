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
- `AppointmentItem` exposes a `Forward` **event** when the user invokes Outlook's Forward action, but does not expose an equivalent native Forward method;
- Classic Outlook itself can forward an accepted meeting directly from Calendar;
- `CommandBars.ExecuteMso(idMso)` can invoke a built-in Office command when the object model has no direct method.

Primary references:

- https://learn.microsoft.com/en-us/office/vba/api/outlook.meetingitem.forward%28method%29
- https://learn.microsoft.com/en-us/office/vba/api/outlook.meetingitem.getassociatedappointment
- https://learn.microsoft.com/en-us/office/vba/api/outlook.appointmentitem.forward
- https://learn.microsoft.com/en-us/office/vba/api/outlook.appointmentitem.forwardasvcal
- https://learn.microsoft.com/en-us/office/vba/api/office.commandbars.executemso
- https://support.microsoft.com/en-us/outlook/calendar/forward-a-meeting-in-outlook

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

Command:

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

## Real-machine result #3 — Calendar Forward produces a native MeetingItem and cancels cleanly

Tested 2026-09-08 against the **same accepted meeting** used for results #1 and #2.

Observed:

- `idMso`: `Forward`;
- identifier valid: **True**;
- visible: **True**;
- enabled: **True**;
- `AppointmentItem.Forward` event raised;
- forwarded object type: **native `MeetingItem`**;
- forwarded message class: **`IPM.Schedule.Meeting.Request`**;
- `Cancel=True` set inside the event;
- forward operation did not complete or display the new item;
- verified subject: `List of application`.

### Conclusion from result #3

The Calendar-command route has now proved all of the important pre-send mechanics on a real accepted meeting whose original request is no longer recoverable:

1. Outlook exposes native Calendar Forward;
2. the command can be invoked programmatically;
3. Outlook raises the documented `AppointmentItem.Forward` event;
4. the event supplies a genuine native `MeetingItem`, not a MailItem/vCalendar substitute;
5. cancellation works before the forward is presented to the user.

This is substantially stronger than Path A for the tested meeting and justifies progressing to recipient handling. It does **not** yet prove that the forwarded meeting can be addressed and sent safely from the selected source account.

### B3a. Recipient prepare/discard — next gate

The next experiment should deliberately avoid sending. It will allow Outlook's Forward operation to complete far enough to obtain the native forwarded `MeetingItem`, then:

1. add exactly one explicit recipient supplied with `--to`;
2. require `Recipients.ResolveAll()`;
3. pin `SendUsingAccount` to the selected source Outlook account;
4. verify no unexpected pre-existing recipients are present;
5. close/discard the unsent forwarded item;
6. verify no draft remains and the source appointment is unchanged.

Only if B3a succeeds should Phase 2 expose a separately confirmed Calendar-command send mode.

### B3b. Native Calendar-command send — only after B3a succeeds

The final send experiment must:

- target another account controlled by the user;
- require an explicit send mode and exact confirmation token;
- revalidate meeting identity and Forward capability immediately before executing;
- add exactly one intended recipient;
- resolve the recipient;
- set `SendUsingAccount` to the selected source account;
- call `MeetingItem.Send()` only after all gates pass;
- verify the target receives a genuine meeting request with normal Accept/Tentative/Decline behavior and working Teams join where applicable.

## Positive-control test for Path A

Path A should still be sampled against one recent accepted meeting whose original invitation is visibly retained in Inbox.

Possible outcomes:

- **1 native match:** Path A remains conditionally useful/direct;
- **0 matches despite the visible request:** the current request-correlation implementation needs repair before judging Path A.

This positive control is useful evidence but no longer blocks testing Path B because results #2 and #3 independently prove the Calendar native Forward route exists and produces a native MeetingItem.

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
- [ ] if exactly one match exists, Prepare succeeds and discards unsent;
- [ ] if Prepare succeeds, one explicitly confirmed native send is verified.

### Calendar-command path

- [x] exact Microsoft Fluent command identifier confirmed;
- [x] capability probe reports valid/visible/enabled state without executing Forward;
- [x] prepare/cancel invocation tested only after capability probe succeeds;
- [x] appointment Forward event produces the expected native `MeetingItem`;
- [x] cancellation prevents the forward from completing/displaying;
- [ ] recipient prepare/discard succeeds with exactly one intended recipient;
- [ ] no draft remains after recipient prepare/discard;
- [ ] source appointment remains unchanged;
- [ ] one explicitly confirmed Calendar-command send is tested only after recipient prepare passes;
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
- `5` — source/native selection or correlation could not be resolved safely.
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
