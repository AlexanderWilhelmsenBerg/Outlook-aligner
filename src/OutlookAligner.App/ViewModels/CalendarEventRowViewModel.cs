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
        ? $"Recurring · {CalendarEvent.RecurrenceState}"
        : "Single event";

    public string Location => string.IsNullOrWhiteSpace(CalendarEvent.Location) ? "—" : CalendarEvent.Location;

    public bool HasForwardIdentity
        => !string.IsNullOrWhiteSpace(Account.SmtpAddress)
            && !string.IsNullOrWhiteSpace(CalendarEvent.EntryId)
            && !string.IsNullOrWhiteSpace(CalendarEvent.GlobalAppointmentId);
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
}
