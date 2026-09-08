using OutlookAligner.Outlook.Contracts;

namespace OutlookAligner.App.ViewModels;

public sealed class CalendarEventRowViewModel
{
    public CalendarEventRowViewModel(OutlookAccountDto account, CalendarEventDto calendarEvent)
    {
        Account = account;
        CalendarEvent = calendarEvent;
    }

    public OutlookAccountDto Account { get; }

    public CalendarEventDto CalendarEvent { get; }

    public string Subject => string.IsNullOrWhiteSpace(CalendarEvent.Subject) ? "(no subject)" : CalendarEvent.Subject;

    public string AccountName => Account.DisplayName;

    public string SourceSmtp => Account.SmtpAddress ?? "(SMTP unavailable)";

    public string TimeDisplay => CalendarEvent.IsAllDay
        ? $"{CalendarEvent.StartLocal:d} · all day"
        : $"{CalendarEvent.StartLocal:g} – {CalendarEvent.EndLocal:g}";

    public string RecurrenceDisplay => CalendarEvent.IsRecurring
        ? $"Recurring · {MeetingStatusDisplay} · {CalendarEvent.RecurrenceState}"
        : $"Single event · {MeetingStatusDisplay}";

    public string Location => string.IsNullOrWhiteSpace(CalendarEvent.Location) ? "—" : CalendarEvent.Location;

    public string MeetingStatusDisplay => CalendarEvent.MeetingStatus switch
    {
        "olNonMeeting" => "Appointment (not a meeting)",
        "olMeeting" => "Meeting",
        "olMeetingReceived" => "Received meeting",
        "olMeetingCanceled" => "Canceled meeting",
        "olMeetingReceivedAndCanceled" => "Received canceled meeting",
        _ => CalendarEvent.MeetingStatus,
    };

    public bool HasNativeIdentity
        => !string.IsNullOrWhiteSpace(Account.SmtpAddress)
            && !string.IsNullOrWhiteSpace(CalendarEvent.EntryId)
            && !string.IsNullOrWhiteSpace(CalendarEvent.GlobalAppointmentId);

    public bool CanUseNativeForward
        => HasNativeIdentity
            && !CalendarEvent.IsRecurring
            && CalendarEvent.MeetingStatus is not "olNonMeeting"
            && CalendarEvent.MeetingStatus is not "olMeetingCanceled"
            && CalendarEvent.MeetingStatus is not "olMeetingReceivedAndCanceled";

    // MainViewModel's current command gate uses this property. In this increment,
    // "forward identity" means identity plus a meeting type whose Forward semantics are proven.
    public bool HasForwardIdentity => CanUseNativeForward;

    public string ForwardEligibilityLabel => CanUseNativeForward ? "Ready to check" : "Not available for this item";

    public string ForwardEligibilityMessage
    {
        get
        {
            if (!HasNativeIdentity)
            {
                return "Outlook did not provide enough native identity for safe meeting Forward.";
            }

            if (CalendarEvent.MeetingStatus == "olNonMeeting")
            {
                return "This is an Outlook appointment without meeting attendees. Meeting Forward does not apply.";
            }

            if (CalendarEvent.MeetingStatus is "olMeetingCanceled" or "olMeetingReceivedAndCanceled")
            {
                return "This meeting is canceled, so Outlook Aligner will not prepare a Forward action.";
            }

            if (CalendarEvent.IsRecurring)
            {
                return "Recurring meeting Forward is intentionally disabled in this increment until occurrence-versus-series identity is implemented safely.";
            }

            return "This one-off meeting can be checked against Outlook's native Forward capability.";
        }
    }
}

public sealed class OutlookAccountChoice
{
    public OutlookAccountChoice(OutlookAccountDto account)
    {
        Account = account;
    }

    public OutlookAccountDto Account { get; }

    public string DisplayName => string.IsNullOrWhiteSpace(Account.SmtpAddress)
        ? Account.DisplayName
        : $"{Account.DisplayName} ({Account.SmtpAddress})";

    public string? SmtpAddress => Account.SmtpAddress;

    public string AccountType => Account.AccountType;

    public int EventCount => Account.EventCount;

    public string CalendarStatus => Account.CalendarAvailable
        ? $"Calendar available · {Account.EventCount} events in current horizon"
        : "Calendar unavailable";

    public string ErrorSummary => string.IsNullOrWhiteSpace(Account.Error)
        ? string.Empty
        : "Outlook reported an account error; see Diagnostics.";
}
