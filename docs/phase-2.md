# Phase 2 — Native Meeting Forwarding Technical Spike

Status: **In progress 🚧 — end-to-end mechanism proven; reliability matrix remains**

Started: 2026-09-08

## Purpose

Determine whether Outlook Aligner can honestly provide **genuine Classic Outlook meeting forwarding** for accepted meetings.

The product must never label `AppointmentItem.ForwardAsVcal()` as equivalent native forwarding.

## Outlook facts driving the spike

Microsoft documents that:

- `MeetingItem.Forward()` executes Outlook's native Forward action and returns a new `MeetingItem`;
- `MeetingItem.GetAssociatedAppointment(false)` maps a retained meeting request back to its calendar appointment without adding an item;
- an accepted calendar meeting is represented as `AppointmentItem`;
- `AppointmentItem` exposes a `Forward` event when Outlook's Forward action is invoked;
- Classic Outlook can forward an accepted meeting directly from Calendar;
- `CommandBars.ExecuteMso(idMso)` can invoke a built-in Office command where no direct object-model method exists.

Primary references:

- https://learn.microsoft.com/en-us/office/vba/api/outlook.meetingitem.forward%28method%29
- https://learn.microsoft.com/en-us/office/vba/api/outlook.meetingitem.getassociatedappointment
- https://learn.microsoft.com/en-us/office/vba/api/outlook.appointmentitem.forward
- https://learn.microsoft.com/en-us/office/vba/api/outlook.appointmentitem.forwardasvcal
- https://learn.microsoft.com/en-us/office/vba/api/office.commandbars.executemso
- https://support.microsoft.com/en-us/outlook/calendar/forward-a-meeting-in-outlook

## Path A — retained MeetingItem recovery

The first implementation selects a source account by SMTP address, then searches that delivery store's default Inbox and Deleted Items for retained `IPM.Schedule.Meeting.Request` items. Candidates are correlated to Calendar with `GetAssociatedAppointment(false)` and `GlobalAppointmentID`.

### Real-machine result #1 — retained request unavailable

Tested 2026-09-08 against an accepted Kverneland Group meeting.

Source account:

`Alexander.Berg@kvernelandgroup.com`

Calendar `GlobalAppointmentID`:

`040000008200E00074C5B7101A82E00800000000D05887327B3BDD01000000000000000010000000EB2313460E921449BF83A27C3A21C827`

Observed:

- source account resolved;
- Inbox searched;
- Deleted Items searched;
- matching retained meeting requests: **0**;
- no vCalendar fallback;
- no Outlook data modified.

Conclusion: retained-request recovery is **not universal after acceptance** and cannot be the product's primary Forward prerequisite.

Path A remains useful diagnostic evidence and may remain as an optional direct route if later testing justifies the complexity, but the product must not depend on it.

## Path B — accepted Calendar AppointmentItem / built-in Forward command

Path B operates directly from the accepted Calendar `AppointmentItem` and Outlook's built-in Fluent command `Forward`.

### B1 — capability probe

```powershell
.\OutlookAligner.OutlookHost.exe --forward-spike `
  --source-smtp source@example.com `
  --global-id GLOBAL_ID `
  --probe-calendar-command `
  --entry-id ENTRY_ID
```

The probe reopens the exact appointment by EntryID in the selected Store, verifies `GlobalAppointmentID`, verifies meeting state, opens an Inspector, queries the built-in `Forward` command state, and closes without executing Forward.

### Real-machine result #2 — native Forward capability present

On the same accepted meeting used for result #1:

- `idMso`: `Forward`;
- identifier valid: **True**;
- label: `Forward`;
- visible: **True**;
- enabled: **True**;
- command not executed;
- diagnostic subject: `List of application`.

Conclusion: Classic Outlook exposes native Calendar Forward even though the retained request is unavailable.

### B2 — prepare/cancel native-object proof

```powershell
.\OutlookAligner.OutlookHost.exe --forward-spike `
  --source-smtp source@example.com `
  --global-id GLOBAL_ID `
  --prepare-calendar-command `
  --entry-id ENTRY_ID
```

The experiment subscribes to `AppointmentItem.Forward`, invokes only the verified `Forward` command, inspects the object Outlook supplies, sets `Cancel=True` inside the event, and never adds a recipient or calls `Send()`/`Save()`.

### Real-machine result #3 — native MeetingItem produced and cancelled

On the same meeting:

- `Forward` remained valid/visible/enabled;
- `AppointmentItem.Forward` fired;
- forwarded object type: **native `MeetingItem`**;
- message class: **`IPM.Schedule.Meeting.Request`**;
- `Cancel=True` prevented the forward from completing/displaying;
- source meeting remained the selected accepted appointment.

Conclusion: Outlook's Calendar Forward command produces a genuine native meeting request object, not a MailItem/vCalendar substitute.

### B3a — recipient prepare/discard

```powershell
.\OutlookAligner.OutlookHost.exe --forward-spike `
  --source-smtp source@example.com `
  --global-id GLOBAL_ID `
  --prepare-calendar-recipient `
  --entry-id ENTRY_ID `
  --to target@example.com
```

Safety gates:

1. exact source account/store;
2. exact EntryID + GlobalAppointmentID identity;
3. meeting state;
4. `Forward` valid/visible/enabled;
5. Forward event supplies native `MeetingItem`;
6. forwarded item starts with zero recipients;
7. exactly one explicit recipient is added and `ResolveAll()` succeeds;
8. recipient count remains exactly one;
9. `SendUsingAccount` is pinned to the selected source account;
10. no `Send()` or `Save()`; the item is discarded with `olDiscard`.

### Real-machine result #4 — recipient prepare succeeds

Tested with:

- source: `Alexander.Berg@kvernelandgroup.com`;
- target: `alexander.berg@knowit.no`;
- subject: `List of application`.

Observed:

- Calendar Forward produced native `MeetingItem`;
- forwarded item started without unintended recipients;
- exactly one target recipient resolved;
- `SendUsingAccount` was pinned to Kverneland;
- no send occurred;
- transient forward was discarded.

Conclusion: the Calendar-command path remains usable after the Forward event completes and can be safely addressed from the intended Outlook account.

### B3b — explicitly confirmed Calendar-command send

```powershell
.\OutlookAligner.OutlookHost.exe --forward-spike `
  --source-smtp source@example.com `
  --global-id GLOBAL_ID `
  --send-calendar-command `
  --entry-id ENTRY_ID `
  --to target@example.com `
  --confirm-calendar-send SEND-CALENDAR-FORWARD
```

The Calendar send confirmation token is deliberately separate from the older retained-request send token.

Before `MeetingItem.Send()` the implementation repeats every identity, capability, native-object, zero-recipient, recipient-resolution, exact-recipient-count and `SendUsingAccount` gate proven in B1-B3a.

### Real-machine result #5 — end-to-end native Calendar Forward succeeds

Tested 2026-09-08 with the same accepted meeting:

- source: `Alexander.Berg@kvernelandgroup.com`;
- target: `alexander.berg@knowit.no`;
- subject: `List of application`;
- `Forward` valid/visible/enabled immediately before execution;
- native `MeetingItem` created;
- exactly one recipient resolved;
- `SendUsingAccount` pinned to Kverneland;
- `MeetingItem.Send()` completed;
- the forwarded meeting was present in the Knowit account.

Console result:

> Outlook's Calendar Forward created a native MeetingItem, exactly one recipient resolved, SendUsingAccount was pinned to the selected source account, and MeetingItem.Send() completed.

### Product decision after result #5

The **Calendar-command route is now the primary candidate for the production Forward action**.

For the tested accepted meeting, the complete supported path is:

```text
accepted Calendar AppointmentItem
  -> capability check for built-in Forward
  -> ExecuteMso("Forward")
  -> AppointmentItem.Forward event
  -> native MeetingItem
  -> exactly one intended recipient
  -> SendUsingAccount = source account
  -> MeetingItem.Send()
```

This route does not require the original meeting request to still exist in Inbox/Deleted Items and does not fake forwarding through ICS/vCalendar.

The production GUI should therefore treat Forward as an **event capability**: enable the action when the selected Calendar meeting passes the native Forward capability checks; otherwise disable it with a clear reason.

## Remaining reliability matrix

End-to-end mechanism is proven. Before Phase 2 closes, sample the smallest useful matrix rather than repeating identical sends:

| Case | Minimum test | Purpose |
| --- | --- | --- |
| Normal accepted meeting | **PASS** through real send | Baseline end-to-end behavior |
| Recurring/Teams meeting or occurrence | capability + prepare; send only if useful | Recurrence/online-meeting behavior |
| Different source Outlook account | capability + prepare, preferably one real send | Account/store selection and `SendUsingAccount` portability |
| Forwarding-disabled meeting, if readily available | capability probe only | Must report unavailable/fail closed |

For an additional real send, use only an account controlled by the user.

## Path A positive control

A retained-request positive control is now **non-blocking**. It may still be run against a recent meeting whose invitation visibly remains in Inbox to decide whether Path A is worth retaining as a secondary direct COM route.

If Path B remains reliable across the matrix, Phase 2 closeout should prefer one production mechanism and remove/simplify unnecessary retained-request complexity rather than maintaining two equivalent write paths without benefit.

## Privacy and safety

- No meeting body, attendees, attachments, Teams URLs, or location are printed by default.
- Subject requires explicit `--include-details` diagnostic opt-in.
- Errors use operation/HRESULT/type context rather than raw meeting content.
- No `ForwardAsVcal()` fallback.
- No replacement organizer-owned meeting is created.
- Calendar send requires an exact dedicated confirmation token.
- Source account, event identity, Forward capability and recipient count are revalidated before send.

## Manual merge gate

Before PR #5 may merge:

### Calendar-command path

- [x] exact Fluent command identifier confirmed;
- [x] capability probe succeeds without executing Forward;
- [x] prepare/cancel produces native `MeetingItem`;
- [x] cancellation prevents completion/display;
- [x] recipient prepare/discard succeeds with exactly one intended recipient;
- [x] source account can be pinned with `SendUsingAccount`;
- [x] one explicitly confirmed Calendar-command send completes;
- [x] forwarded meeting is present in the controlled target account;
- [ ] recurring/Teams case sampled;
- [ ] different source account sampled;
- [ ] forwarding-disabled behavior sampled if a suitable event is readily available.

### General

- [ ] confirm no unintended recipients or source-calendar mutation were observed during the completed send test;
- [ ] decide whether Path A stays as a secondary route or is removed/simplified;
- [ ] hosted CI is green on the final documentation/code head;
- [ ] user explicitly approves merge.

## Hosted verification

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

Final send implementation checkpoint before this documentation update:

- commit `c32d83ae96603a5bbe9f3c1325ec1887ca7a5954`;
- CI run `34209900331` green;
- artifact `OutlookAligner-Phase2-Forwarding-Spike-win-x64`;
- artifact ID `10049400062`;
- SHA-256 `13b8aad80a682dd20869f8eb86bc4b28e93f5f7821a7d91b6ab6ddea20949201`.

## Exit codes

- `0` — requested operation completed successfully.
- `2` — invalid command line.
- `3` — Outlook COM failure.
- `4` — unexpected/dependency failure.
- `5` — source/native selection or correlation could not be resolved safely.
- `6` — recipient could not be resolved; unsent forward discarded where applicable.

## Explicit non-goals

Phase 2 does not:

- create a replacement organizer-owned meeting;
- use `AppointmentItem.ForwardAsVcal()` as native Forward;
- copy Teams/body content manually to fake forwarding;
- implement the production GUI;
- implement Copy Full / Copy Busy;
- merge automatically.

The production GUI is nevertheless a required v1 surface and is specified separately in [`docs/ui.md`](ui.md).