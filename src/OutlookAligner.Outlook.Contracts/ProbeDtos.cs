namespace OutlookAligner.Outlook.Contracts;

public sealed record OutlookProbeResult(
    int ProtocolVersion,
    DateTime CapturedAtUtc,
    DateTime WindowStartLocal,
    DateTime WindowEndLocal,
    string? ProfileName,
    IReadOnlyList<OutlookStoreDto> Stores,
    IReadOnlyList<OutlookAccountDto> Accounts,
    IReadOnlyList<string> Warnings);

public sealed record OutlookStoreDto(
    string DisplayName,
    string StoreId);

public sealed record OutlookAccountDto(
    string DisplayName,
    string? SmtpAddress,
    string AccountType,
    string? StoreDisplayName,
    string? StoreId,
    string? CalendarEntryId,
    bool CalendarAvailable,
    int EventCount,
    IReadOnlyList<CalendarEventDto> Events,
    string? Error);

public sealed record CalendarEventDto(
    string? EntryId,
    string? GlobalAppointmentId,
    DateTime StartLocal,
    DateTime EndLocal,
    bool IsAllDay,
    bool IsRecurring,
    string RecurrenceState,
    string BusyStatus,
    string Sensitivity,
    string? Subject,
    string? Location);
