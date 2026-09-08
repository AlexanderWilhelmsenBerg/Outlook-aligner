namespace OutlookAligner.Outlook.Contracts;

public static class OutlookAlignerManagedProperties
{
    public const string SyncGroupId = "OutlookAligner.SyncGroupId";
    public const string SourceGlobalAppointmentId = "OutlookAligner.SourceGlobalAppointmentId";
    public const string SourceAccountId = "OutlookAligner.SourceAccountId";
    public const string CopyType = "OutlookAligner.CopyType";
    public const string SchemaVersion = "OutlookAligner.SchemaVersion";

    public const string CurrentSchemaVersion = "1";
    public const string CopyTypeFull = "Full";
    public const string CopyTypeBusy = "Busy";
}

public enum ManagedCopyState
{
    None,
    Valid,
    Incomplete,
    UnsupportedSchema,
    Unreadable,
}

public sealed record ManagedCopyMetadataDto(
    ManagedCopyState State,
    string? SyncGroupId,
    string? SourceGlobalAppointmentId,
    string? SourceAccountId,
    string? CopyType,
    string? SchemaVersion);

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
    string? Location,
    ManagedCopyMetadataDto? ManagedCopy = null);
