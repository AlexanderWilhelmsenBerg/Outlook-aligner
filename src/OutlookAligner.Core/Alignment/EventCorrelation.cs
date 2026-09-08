namespace OutlookAligner.Core.Alignment;

public enum AlignmentState
{
    Aligned,
    Missing,
    Moved,
    DetailsDifferent,
    Duplicate,
    Conflict,
    RecurrenceIdentityUnresolved,
    Uncorrelated,
}

public sealed record ObservedCalendarEvent(
    string AccountKey,
    string LocatorKey,
    string? GlobalAppointmentId,
    DateTime StartLocal,
    DateTime EndLocal,
    bool IsAllDay,
    bool IsRecurring,
    string BusyStatus,
    string Sensitivity,
    string? Subject,
    string? Location,
    bool IsManagedCopy = false,
    string? ManagedSourceGlobalAppointmentId = null,
    string? ManagedSyncGroupId = null,
    string? ManagedSourceAccountId = null,
    string? ManagedCopyType = null);

public sealed record LogicalEventGroup(
    string GroupKey,
    string DisplaySubject,
    AlignmentState State,
    IReadOnlyList<ObservedCalendarEvent> Members,
    IReadOnlyList<string> MissingAccountKeys,
    bool TimesDiffer,
    bool DetailsDiffer,
    string Explanation);

public static class EventCorrelation
{
    public static IReadOnlyList<LogicalEventGroup> BuildGroups(
        IEnumerable<ObservedCalendarEvent> events,
        IEnumerable<string> expectedAccountKeys)
    {
        ArgumentNullException.ThrowIfNull(events);
        ArgumentNullException.ThrowIfNull(expectedAccountKeys);

        var expectedAccounts = expectedAccountKeys
            .Where(account => !string.IsNullOrWhiteSpace(account))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(account => account, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        var observations = events.ToArray();
        var results = new List<LogicalEventGroup>();

        foreach (var observation in observations.Where(item => item.IsRecurring))
        {
            results.Add(CreateUnresolvedGroup(
                observation,
                AlignmentState.RecurrenceIdentityUnresolved,
                "Recurring occurrence identity has not been proven yet, so this item is deliberately not correlated by meeting identity alone."));
        }

        foreach (var observation in observations.Where(item =>
                     !item.IsRecurring && string.IsNullOrWhiteSpace(GetCorrelationGlobalAppointmentId(item))))
        {
            results.Add(CreateUnresolvedGroup(
                observation,
                AlignmentState.Uncorrelated,
                observation.IsManagedCopy
                    ? "The managed copy is missing its validated source GlobalAppointmentID and cannot be safely correlated."
                    : "The item has no GlobalAppointmentID and cannot be safely correlated yet."));
        }

        var correlated = observations
            .Where(item => !item.IsRecurring && !string.IsNullOrWhiteSpace(GetCorrelationGlobalAppointmentId(item)))
            .GroupBy(item => GetCorrelationGlobalAppointmentId(item)!, StringComparer.Ordinal)
            .Select(group => BuildCorrelatedGroup(group.Key, group.ToArray(), expectedAccounts));

        results.AddRange(correlated);

        return results
            .OrderBy(group => group.Members.Min(member => member.StartLocal))
            .ThenBy(group => group.DisplaySubject, StringComparer.CurrentCultureIgnoreCase)
            .ToArray();
    }

    private static LogicalEventGroup BuildCorrelatedGroup(
        string correlationGlobalAppointmentId,
        ObservedCalendarEvent[] members,
        string[] expectedAccounts)
    {
        var duplicateAccounts = members
            .GroupBy(member => member.AccountKey, StringComparer.OrdinalIgnoreCase)
            .Where(group => group.Count() > 1)
            .Select(group => group.Key)
            .OrderBy(account => account, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        var representedAccounts = members
            .Select(member => member.AccountKey)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var missingAccounts = expectedAccounts
            .Where(account => !representedAccounts.Contains(account))
            .ToArray();

        var first = members[0];
        var timesDiffer = members.Any(member =>
            member.StartLocal != first.StartLocal
            || member.EndLocal != first.EndLocal
            || member.IsAllDay != first.IsAllDay);

        var detailsDiffer = members.Any(member =>
            !string.Equals(member.Subject, first.Subject, StringComparison.Ordinal)
            || !string.Equals(member.Location, first.Location, StringComparison.Ordinal)
            || !string.Equals(member.BusyStatus, first.BusyStatus, StringComparison.Ordinal)
            || !string.Equals(member.Sensitivity, first.Sensitivity, StringComparison.Ordinal));

        AlignmentState state;
        string explanation;

        if (duplicateAccounts.Length > 0)
        {
            state = AlignmentState.Duplicate;
            explanation = $"More than one correlated item exists in: {string.Join(", ", duplicateAccounts)}.";
        }
        else if (missingAccounts.Length > 0)
        {
            state = AlignmentState.Missing;
            var suffix = BuildDifferenceSuffix(timesDiffer, detailsDiffer);
            explanation = $"Missing from: {string.Join(", ", missingAccounts)}.{suffix}";
        }
        else if (timesDiffer)
        {
            state = AlignmentState.Moved;
            explanation = detailsDiffer
                ? "The correlated copies have different times and other details."
                : "The correlated copies have different start/end times.";
        }
        else if (detailsDiffer)
        {
            state = AlignmentState.DetailsDifferent;
            explanation = "The correlated copies have matching times but differing visible details/state.";
        }
        else
        {
            state = AlignmentState.Aligned;
            explanation = "All expected accounts contain one correlated item with matching time and visible details/state.";
        }

        return new LogicalEventGroup(
            $"gaid:{correlationGlobalAppointmentId}",
            PickDisplaySubject(members),
            state,
            members,
            missingAccounts,
            timesDiffer,
            detailsDiffer,
            explanation);
    }

    private static LogicalEventGroup CreateUnresolvedGroup(
        ObservedCalendarEvent observation,
        AlignmentState state,
        string explanation)
        => new(
            $"item:{observation.AccountKey}:{observation.LocatorKey}",
            string.IsNullOrWhiteSpace(observation.Subject) ? "(no subject)" : observation.Subject,
            state,
            [observation],
            Array.Empty<string>(),
            TimesDiffer: false,
            DetailsDiffer: false,
            explanation);

    private static string? GetCorrelationGlobalAppointmentId(ObservedCalendarEvent item)
    {
        if (item.IsManagedCopy)
        {
            return Normalize(item.ManagedSourceGlobalAppointmentId);
        }

        return Normalize(item.GlobalAppointmentId);
    }

    private static string? Normalize(string? value)
        => string.IsNullOrWhiteSpace(value) ? null : value.Trim();

    private static string PickDisplaySubject(ObservedCalendarEvent[] members)
        => members
            .Select(member => member.Subject)
            .FirstOrDefault(subject => !string.IsNullOrWhiteSpace(subject))
            ?? "(no subject)";

    private static string BuildDifferenceSuffix(bool timesDiffer, bool detailsDiffer)
    {
        if (timesDiffer && detailsDiffer)
        {
            return " Existing copies also differ in time and details";
        }

        if (timesDiffer)
        {
            return " Existing copies also differ in time";
        }

        if (detailsDiffer)
        {
            return " Existing copies also differ in details";
        }

        return string.Empty;
    }
}
