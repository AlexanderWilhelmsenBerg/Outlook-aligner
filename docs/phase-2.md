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

This does **not** yet prove that genuine forwarding is impossible. Classic Outlook's Calendar UI can still expose Forward for an accepted `AppointmentItem`, so Phase 2 now has a second technical path.

## Path B — accepted AppointmentItem / Outlook Calendar Forward command

The next experiment is deliberately staged.

### B1. Capability probe — no Forward execution

For the exact accepted Calendar `AppointmentItem`:

1. reopen it by StoreID + EntryID;
2. obtain/create its Outlook Inspector context;
3. query Outlook's built-in Fluent command state for the native Calendar Forward control;
4. report whether the command identifier is valid, visible, and enabled;
5. close the Inspector without modifying the appointment.

This probe must **not** call `ExecuteMso()` yet.

The command identifier must come from Microsoft's Office Fluent UI command-identifier catalogue rather than being treated as an undocumented guess.

### B2. Prepare/discard — only if B1 proves capability

If the native Calendar Forward command is valid/enabled:

1. subscribe to the appointment's `Forward` event;
2. invoke only the proven built-in Forward command;
3. capture/verify the new object supplied by Outlook's Forward event;
4. do not add a recipient or send;
5. close/discard the unsent forward;
6. verify the original calendar appointment is unchanged.

### B3. Recipient/send — only after prepare succeeds

Only after B2 succeeds should the spike resolve a recipient and perform one explicitly confirmed send to another account controlled by the user.

Verify that the received item behaves like a real forwarded meeting rather than a generic email/ICS attachment.

## Positive-control test for Path A

Path A should still be tested against one recent accepted meeting whose original invitation is visibly retained in Inbox.

Possible outcomes:

- **1 native match:** Path A is conditionally useful.
- **0 matches despite the visible request:** the current request-correlation implementation needs repair before judging Path A.

Do not run `--prepare` or `--send` for the positive control until Inspect returns exactly one match.

## Privacy

The spike does not read/print meeting body, attendees, attachments, Teams URLs, or location by default.

Subject remains explicit `--include-details` diagnostic opt-in.

Errors use operation/HRESULT/type context rather than raw message contents.

## Current decision model

Phase 2 may end in one of four truthful outcomes:

1. **Reliable native forwarding** — safe native route works broadly; expose Forward normally.
2. **Conditional native forwarding** — expose Forward only when per-event capability proves it is available.
3. **Outlook-UI-command-only forwarding** — retained request is unreliable but a supported Calendar command route is stable enough to automate; expose it with explicit capability checks.
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

- [ ] exact Microsoft Fluent command identifier confirmed;
- [ ] capability probe reports valid/visible/enabled state without executing Forward;
- [ ] prepare/discard invocation tested only after capability probe succeeds;
- [ ] appointment Forward event produces the expected native object;
- [ ] no draft remains after discard;
- [ ] source appointment remains unchanged;
- [ ] one explicitly confirmed send is tested only after prepare passes;
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