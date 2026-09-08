using OutlookAligner.Core.Alignment;
using Xunit;

namespace OutlookAligner.Core.Tests;

public sealed class ManagedCorrelationTests
{
    private static readonly string[] Accounts = ["source@example.invalid", "copy@example.invalid"];

    [Fact]
    public void ManagedCopyCorrelatesByStoredSourceGlobalAppointmentId()
    {
        var source = Event(
            Accounts[0],
            locator: "source",
            globalAppointmentId: "source-global",
            isManagedCopy: false,
            managedSourceGlobalAppointmentId: null);
        var copy = Event(
            Accounts[1],
            locator: "copy",
            globalAppointmentId: "copy-native-global",
            isManagedCopy: true,
            managedSourceGlobalAppointmentId: "source-global");

        var group = Assert.Single(EventCorrelation.BuildGroups([source, copy], Accounts));

        Assert.Equal(AlignmentState.Aligned, group.State);
        Assert.Equal(2, group.Members.Count);
        Assert.Equal("gaid:source-global", group.GroupKey);
    }

    [Fact]
    public void UnmanagedItemDoesNotTrustManagedSourceField()
    {
        var source = Event(
            Accounts[0],
            locator: "source",
            globalAppointmentId: "source-global",
            isManagedCopy: false,
            managedSourceGlobalAppointmentId: null);
        var unrelated = Event(
            Accounts[1],
            locator: "other",
            globalAppointmentId: "other-global",
            isManagedCopy: false,
            managedSourceGlobalAppointmentId: "source-global");

        var groups = EventCorrelation.BuildGroups([source, unrelated], Accounts);

        Assert.Equal(2, groups.Count);
        Assert.All(groups, group => Assert.Equal(AlignmentState.Missing, group.State));
    }

    [Fact]
    public void ManagedCopyWithoutValidatedSourceIdentityFailsClosed()
    {
        var copy = Event(
            Accounts[1],
            locator: "copy",
            globalAppointmentId: "copy-native-global",
            isManagedCopy: true,
            managedSourceGlobalAppointmentId: null);

        var group = Assert.Single(EventCorrelation.BuildGroups([copy], Accounts));

        Assert.Equal(AlignmentState.Uncorrelated, group.State);
        Assert.Contains("managed copy", group.Explanation, StringComparison.OrdinalIgnoreCase);
    }

    private static ObservedCalendarEvent Event(
        string account,
        string locator,
        string globalAppointmentId,
        bool isManagedCopy,
        string? managedSourceGlobalAppointmentId)
        => new(
            account,
            locator,
            globalAppointmentId,
            new DateTime(2026, 9, 8, 9, 0, 0),
            new DateTime(2026, 9, 8, 10, 0, 0),
            IsAllDay: false,
            IsRecurring: false,
            BusyStatus: "Busy",
            Sensitivity: "Normal",
            Subject: "Meeting",
            Location: "Room",
            IsManagedCopy: isManagedCopy,
            ManagedSourceGlobalAppointmentId: managedSourceGlobalAppointmentId,
            ManagedSyncGroupId: isManagedCopy ? "11111111-1111-1111-1111-111111111111" : null,
            ManagedSourceAccountId: isManagedCopy ? Accounts[0] : null,
            ManagedCopyType: isManagedCopy ? "Full" : null);
}
