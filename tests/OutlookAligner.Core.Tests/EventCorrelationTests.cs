using OutlookAligner.Core.Alignment;
using Xunit;

namespace OutlookAligner.Core.Tests;

public sealed class EventCorrelationTests
{
    private static readonly string[] Accounts = ["a@example.invalid", "b@example.invalid", "c@example.invalid"];

    [Fact]
    public void MatchingNonRecurringCopiesAcrossAllAccountsAreAligned()
    {
        var groups = EventCorrelation.BuildGroups(
            Accounts.Select(account => Event(account, "global", 9, 10)),
            Accounts);

        var group = Assert.Single(groups);
        Assert.Equal(AlignmentState.Aligned, group.State);
        Assert.False(group.TimesDiffer);
        Assert.False(group.DetailsDiffer);
        Assert.Empty(group.MissingAccountKeys);
    }

    [Fact]
    public void MissingAccountIsReportedEvenWhenRemainingCopiesDiffer()
    {
        var groups = EventCorrelation.BuildGroups(
            [
                Event(Accounts[0], "global", 9, 10),
                Event(Accounts[1], "global", 10, 11) with { Location = "Other room" },
            ],
            Accounts);

        var group = Assert.Single(groups);
        Assert.Equal(AlignmentState.Missing, group.State);
        Assert.Equal([Accounts[2]], group.MissingAccountKeys);
        Assert.True(group.TimesDiffer);
        Assert.True(group.DetailsDiffer);
        Assert.Contains("also differ", group.Explanation, StringComparison.Ordinal);
    }

    [Fact]
    public void TimeDifferenceIsMovedWhenAllAccountsAreRepresented()
    {
        var groups = EventCorrelation.BuildGroups(
            [
                Event(Accounts[0], "global", 9, 10),
                Event(Accounts[1], "global", 9, 10),
                Event(Accounts[2], "global", 11, 12),
            ],
            Accounts);

        var group = Assert.Single(groups);
        Assert.Equal(AlignmentState.Moved, group.State);
        Assert.True(group.TimesDiffer);
    }

    [Fact]
    public void DetailDifferenceIsReportedWhenTimesMatch()
    {
        var events = Accounts.Select(account => Event(account, "global", 9, 10)).ToArray();
        events[2] = events[2] with { Location = "Different room" };

        var group = Assert.Single(EventCorrelation.BuildGroups(events, Accounts));
        Assert.Equal(AlignmentState.DetailsDifferent, group.State);
        Assert.False(group.TimesDiffer);
        Assert.True(group.DetailsDiffer);
    }

    [Fact]
    public void DuplicateItemsInOneAccountAreNotSilentlyCollapsed()
    {
        var groups = EventCorrelation.BuildGroups(
            [
                Event(Accounts[0], "global", 9, 10, "one"),
                Event(Accounts[0], "global", 9, 10, "two"),
                Event(Accounts[1], "global", 9, 10),
                Event(Accounts[2], "global", 9, 10),
            ],
            Accounts);

        var group = Assert.Single(groups);
        Assert.Equal(AlignmentState.Duplicate, group.State);
        Assert.Contains(Accounts[0], group.Explanation, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void RecurringItemsAreNotGroupedByGlobalAppointmentIdYet()
    {
        var groups = EventCorrelation.BuildGroups(
            [
                Event(Accounts[0], "series", 9, 10, "one") with { IsRecurring = true },
                Event(Accounts[1], "series", 9, 10, "two") with { IsRecurring = true },
            ],
            Accounts);

        Assert.Equal(2, groups.Count);
        Assert.All(groups, group => Assert.Equal(AlignmentState.RecurrenceIdentityUnresolved, group.State));
    }

    [Fact]
    public void ItemWithoutGlobalAppointmentIdRemainsUncorrelated()
    {
        var group = Assert.Single(EventCorrelation.BuildGroups(
            [Event(Accounts[0], null, 9, 10)],
            Accounts));

        Assert.Equal(AlignmentState.Uncorrelated, group.State);
    }

    private static ObservedCalendarEvent Event(
        string account,
        string? globalId,
        int startHour,
        int endHour,
        string locator = "locator")
        => new(
            account,
            locator,
            globalId,
            new DateTime(2026, 9, 8, startHour, 0, 0),
            new DateTime(2026, 9, 8, endHour, 0, 0),
            IsAllDay: false,
            IsRecurring: false,
            BusyStatus: "Busy",
            Sensitivity: "Normal",
            Subject: "Meeting",
            Location: "Room");
}
