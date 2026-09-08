# Phase 2 — Native Meeting Forwarding Technical Spike

Status: **In progress 🚧 — pull request only**

Started: 2026-09-08

## Purpose

Phase 2 proves whether Outlook Aligner can reproduce **genuine Classic Outlook meeting forwarding** for an accepted meeting. This phase deliberately does not treat a vCalendar attachment as equivalent.

Microsoft's Outlook Object Model exposes three facts that define the spike:

- `MeetingItem.Forward()` executes Outlook's native Forward action and returns a new `MeetingItem`.
- `MeetingItem.GetAssociatedAppointment(false)` returns the appointment associated with a meeting request without adding a new calendar item.
- `AppointmentItem.ForwardAsVcal()` only creates a mail item with a vCalendar attachment, so it is not an acceptable substitute for native meeting forwarding.

Primary references:

- https://learn.microsoft.com/en-us/office/vba/api/outlook.meetingitem.forward%28method%29
- https://learn.microsoft.com/en-us/office/vba/api/outlook.meetingitem.getassociatedappointment
- https://learn.microsoft.com/en-us/office/vba/api/outlook.appointmentitem.forwardasvcal
- https://learn.microsoft.com/en-us/office/vba/api/outlook.meetingitem.sendusingaccount

## Technical hypothesis

Accepted meetings normally appear in Calendar as `AppointmentItem` objects, while native forwarding is exposed on `MeetingItem`. The spike therefore asks:

> Can the original/retained `MeetingItem` still be recovered reliably enough from the source account to call `MeetingItem.Forward()`?

The first implementation searches only the selected account's default **Inbox** and **Deleted Items** for retained `IPM.Schedule.Meeting.Request` messages. Each request is correlated back to an appointment with `GetAssociatedAppointment(false)` and compared against the target `GlobalAppointmentID`.

This is intentionally conservative. If the request has already been purged, archived elsewhere, or multiple requests match ambiguously, the spike reports that result and stops. It does not silently switch to vCalendar or create a new organizer-owned meeting.

## Safety model

Three modes are deliberately separated.

### 1. Inspect — default / read-only

```powershell
.\OutlookAligner.OutlookHost.exe --forward-spike `
  --source-smtp source@example.com `
  --global-id GLOBAL_ID
```

Behavior:

- opens Classic Outlook COM on the STA thread;
- selects exactly one Outlook account by SMTP address;
- searches that account's Inbox and Deleted Items;
- correlates retained meeting requests through `GetAssociatedAppointment(false)`;
- reports zero, one, or multiple native matches;
- does **not** call `MeetingItem.Forward()`;
- does **not** create, save, send, move, delete, or modify any Outlook item.

### 2. Prepare — native Forward proof, but no send

```powershell
.\OutlookAligner.OutlookHost.exe --forward-spike `
  --source-smtp source@example.com `
  --global-id GLOBAL_ID `
  --prepare `
  --to target@example.com
```

Behavior:

- requires exactly one retained native meeting request;
- calls the real `MeetingItem.Forward()` method;
- adds and resolves the requested recipient;
- sets `SendUsingAccount` to the explicitly selected source Outlook account;
- discards the unsent forwarded `MeetingItem` with `olDiscard`;
- never calls `Send()`.

This is the preferred first proof that Outlook can construct a native forward for the chosen accepted meeting.

### 3. Send — explicit network side effect

```powershell
.\OutlookAligner.OutlookHost.exe --forward-spike `
  --source-smtp source@example.com `
  --global-id GLOBAL_ID `
  --send `
  --to target@example.com `
  --confirm-send SEND-NATIVE-MEETING
```

Behavior:

- does everything in Prepare mode;
- requires the exact confirmation token `SEND-NATIVE-MEETING`;
- calls `MeetingItem.Send()` only after the recipient resolves;
- uses the selected source Outlook account through `SendUsingAccount`.

The tool refuses `--send` without both `--to` and the exact confirmation token.

## Privacy behavior

The spike does not read or print body content, attendees, attachments, or Teams URLs.

By default it reports only opaque identity, folder, time range, message class, and result state. Add `--include-details` only when you explicitly want the matched appointment subject printed for validation.

Errors use exception type/HRESULT context rather than raw calendar/message bodies.

## Bounded search

The spike searches only two default folders:

1. Inbox;
2. Deleted Items.

Only `IPM.Schedule.Meeting.Request` items are considered. Each folder is capped at 5,000 retained meeting requests to prevent a pathological unbounded diagnostic scan.

A zero-match result is **not automatically a bug**. It is evidence about whether the native `MeetingItem` recovery strategy is reliable enough for the product.

## Getting the GlobalAppointmentID

Use the existing Phase 1 read probe:

```powershell
.\OutlookAligner.OutlookHost.exe --days 90 --json --include-details > probe.json
```

Find a real future accepted meeting that you are **not** the organizer of. Record:

- the source account SMTP address;
- the appointment's `GlobalAppointmentId`;
- subject/start time for your own verification.

Prefer a normal Teams meeting for the main acceptance test.

## Manual acceptance sequence — merge gate

Use a meeting that is safe to forward to one of your own other accounts.

### A. Inspect

- [ ] `--forward-spike --help` runs without opening Outlook.
- [ ] Inspect mode finds the selected source account.
- [ ] The correct Inbox/Deleted Items folders are searched.
- [ ] The chosen meeting produces exactly one native retained `MeetingItem` match, **or** a zero-match result is recorded as evidence that this recovery strategy is insufficient.
- [ ] Inspect mode creates/sends nothing.
- [ ] With `--include-details`, the reported subject matches the intended appointment.

### B. Prepare

Run Prepare mode only if Inspect produced exactly one match.

- [ ] `MeetingItem.Forward()` succeeds.
- [ ] The target recipient resolves.
- [ ] The forward uses the selected source Outlook account.
- [ ] No message arrives at the target account.
- [ ] No draft or unsent meeting remains after the command exits.
- [ ] The source calendar appointment remains unchanged.

### C. Native send

Run Send mode only after Prepare succeeds, using another account you control as the target.

- [ ] The command refuses to send without the exact confirmation token.
- [ ] The forwarded item arrives at the target as a genuine Outlook meeting invitation/update, not a generic email with `.ics` attachment.
- [ ] The target can Accept/Tentative/Decline using normal Outlook meeting controls.
- [ ] The forwarded meeting keeps the expected original organizer identity/meeting relationship.
- [ ] For a Teams meeting, the Teams join capability remains usable at the target.
- [ ] Normal reminder/meeting behavior at the target is plausible after acceptance.
- [ ] The selected source account is the forwarding sender.
- [ ] The source appointment is not recreated or replaced.
- [ ] No unintended attendees or accounts receive the forward.

### D. Reliability sampling

Repeat Inspect/Prepare against several meetings:

- [ ] recent accepted meeting request still in Inbox;
- [ ] accepted request moved to Deleted Items;
- [ ] older accepted meeting;
- [ ] recurring Teams meeting;
- [ ] meeting from each of the three configured accounts where applicable.

Record how often a retained native `MeetingItem` can actually be recovered.

## Decision rule

Phase 2 is successful only if the native path is reliable enough to expose honestly in Outlook Aligner.

Possible outcomes:

1. **Reliable native forwarding** — continue with `MeetingItem.Forward()` as the product's Forward action.
2. **Conditionally available** — expose Forward only when a recoverable native `MeetingItem` exists; otherwise offer Copy Full / Copy Busy later.
3. **Not reliable enough** — do not ship a Forward action. Do not substitute `ForwardAsVcal()` under the same label.

## Hosted CI

Hosted runners cannot prove live Outlook forwarding because they do not have the user's Outlook profile.

CI must still prove:

- [ ] formatter/analyzer gate passes;
- [ ] Release solution build passes;
- [ ] xUnit/MTP tests and coverage pass;
- [ ] benchmark project compiles;
- [ ] NuGet/npm audits pass;
- [ ] self-contained `win-x64` publish succeeds;
- [ ] published executable passes `--help`;
- [ ] published executable passes `--interop-check`;
- [ ] published executable passes `--forward-spike --help` without opening Outlook;
- [ ] Phase 2 artifact uploads successfully.

Hosted tests cover command gating, read-only default mode, prepare/send separation, confirmation-token enforcement, and malformed forwarding arguments.

## Exit codes

- `0` — requested inspect/prepare/send operation completed successfully.
- `2` — invalid command-line arguments.
- `3` — Outlook COM failure.
- `4` — unexpected/dependency failure.
- `5` — source account/native request selection or correlation could not be resolved safely.
- `6` — forwarding recipient could not be resolved; unsent forward discarded.

## Explicit non-goals

Phase 2 does not:

- create a replacement organizer-owned meeting;
- use `AppointmentItem.ForwardAsVcal()` as fallback;
- copy meeting bodies or Teams URLs manually;
- correlate meetings across all three calendars beyond `GlobalAppointmentID` selection;
- persist forwarding state;
- implement production UI;
- implement Copy Full / Copy Busy;
- merge automatically.

Do not merge Phase 2 until hosted CI is green and the user has completed enough real-Outlook tests to decide which native-forwarding outcome is truthful.
