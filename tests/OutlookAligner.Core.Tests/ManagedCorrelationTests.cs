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

    [Fact]
    public void SuspiciousManagedMetadataIsNotDowngradedToUnmanagedCorrelation()
    {
        var source = Event(Accounts[0], "source", "shared-global", false, null);
        var suspicious = Event(Accounts[1], "copy", "shared-global", false, null) with
        {
            HasManagedMetadataIssue = true,
            ManagedMetadataIssue = "Outlook Aligner metadata is incomplete; correlation and alignment actions are blocked.",
        };

        var groups = EventCorrelation.BuildGroups([source, suspicious], Accounts);

        Assert.Equal(2, groups.Count);
        var conflict = Assert.Single(groups.Where(group => group.State == AlignmentState.Conflict));
        Assert.Equal("copy", Assert.Single(conflict.Members).LocatorKey);
        Assert.Contains("incomplete", conflict.Explanation, StringComparison.OrdinalIgnoreCase);

        var sourceGroup = Assert.Single(groups.Where(group => group.State == AlignmentState.Missing));
        Assert.Equal("source", Assert.Single(sourceGroup.Members).LocatorKey);
    }

    [Fact]
    public void ManagedCopiesWithConflictingOriginsBecomeConflict()
    {
        var source = Event(Accounts[0], "source", "source-global", false, null);
        var copyOne = Event(Accounts[1], "copy-one", "copy-one-global", true, "source-global");
        var copyTwo = Event("third@example.invalid", "copy-two", "copy-two-global", true, "source-global") with
        {
            ManagedSourceAccountId = "other-source@example.invalid",
        };

        var group = Assert.Single(EventCorrelation.BuildGroups(
            [source, copyOne, copyTwo],
            [Accounts[0], Accounts[1], "third@example.invalid"]));

        Assert.Equal(AlignmentState.Conflict, group.State);
        Assert.Contains("source account", group.Explanation, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ManagedCopiesWithConflictingSyncGroupsBecomeConflict()
    {
        var source = Event(Accounts[0], "source", "source-global", false, null);
        var copyOne = Event(Accounts[1], "copy-one", "copy-one-global", true, "source-global");
        var copyTwo = Event("third@example.invalid", "copy-two", "copy-two-global", true, "source-global") with
        {
            ManagedSyncGroupId = "22222222-2222-2222-2222-222222222222",
        };

        var group = Assert.Single(EventCorrelation.BuildGroups(
            [source, copyOne, copyTwo],
            [Accounts[0], Accounts[1], "third@example.invalid"]));

        Assert.Equal(AlignmentState.Conflict, group.State);
        Assert.Contains("sync group", group.Explanation, StringComparison.OrdinalIgnoreCase);
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
