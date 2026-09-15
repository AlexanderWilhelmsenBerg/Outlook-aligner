using OutlookAligner.Core.Alignment;
using Xunit;

namespace OutlookAligner.Core.Tests;

public sealed class AuthorityResolverTests
{
    [Fact]
    public void ValidManagedCopyProvenanceResolvesKnownOrigin()
    {
        var group = Group(
            Source("source@example.invalid"),
            ManagedCopy("copy@example.invalid", "source@example.invalid", "group-one"));

        var resolution = AuthorityResolver.Resolve(group);

        Assert.True(resolution.HasAuthority);
        Assert.Equal("source@example.invalid", resolution.AuthorityAccountKey);
        Assert.Equal(AuthorityReason.KnownOrigin, resolution.Reason);
    }

    [Fact]
    public void PreExistingCopiesDoNotInventAuthority()
    {
        var group = Group(
            Source("one@example.invalid"),
            Source("two@example.invalid"));

        var resolution = AuthorityResolver.Resolve(group);

        Assert.False(resolution.HasAuthority);
        Assert.Equal(AuthorityReason.Unknown, resolution.Reason);
    }

    [Fact]
    public void ConflictingOriginMetadataDoesNotInferAuthority()
    {
        var group = Group(
            Source("source@example.invalid"),
            ManagedCopy("copy-one@example.invalid", "source@example.invalid", "group-one"),
            ManagedCopy("copy-two@example.invalid", "other@example.invalid", "group-one"));

        var resolution = AuthorityResolver.Resolve(group);

        Assert.False(resolution.HasAuthority);
        Assert.Contains("disagree", resolution.Explanation, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void MissingObservedSourceDoesNotInferAuthority()
    {
        var group = Group(
            ManagedCopy("copy@example.invalid", "source@example.invalid", "group-one"));

        var resolution = AuthorityResolver.Resolve(group);

        Assert.False(resolution.HasAuthority);
        Assert.Contains("observed source", resolution.Explanation, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ConflictingGroupIdsDoNotInferAuthority()
    {
        var group = Group(
            Source("source@example.invalid"),
            ManagedCopy("copy-one@example.invalid", "source@example.invalid", "group-one"),
            ManagedCopy("copy-two@example.invalid", "source@example.invalid", "group-two"));

        var resolution = AuthorityResolver.Resolve(group);

        Assert.False(resolution.HasAuthority);
        Assert.Contains("sync group", resolution.Explanation, StringComparison.OrdinalIgnoreCase);
    }

    private static LogicalEventGroup Group(params ObservedCalendarEvent[] members)
        => new(
            "gaid:source-global",
            "Meeting",
            AlignmentState.Aligned,
            members,
            Array.Empty<string>(),
            TimesDiffer: false,
            DetailsDiffer: false,
            "test");

    private static ObservedCalendarEvent Source(string account)
        => new(
            account,
            account + ":source",
            "source-global",
            new DateTime(2026, 9, 8, 9, 0, 0),
            new DateTime(2026, 9, 8, 10, 0, 0),
            IsAllDay: false,
            IsRecurring: false,
            BusyStatus: "Busy",
            Sensitivity: "Normal",
            Subject: "Meeting",
            Location: "Room");

    private static ObservedCalendarEvent ManagedCopy(string account, string sourceAccount, string syncGroupId)
        => new(
            account,
            account + ":copy",
            "copy-native-global",
            new DateTime(2026, 9, 8, 9, 0, 0),
            new DateTime(2026, 9, 8, 10, 0, 0),
            IsAllDay: false,
            IsRecurring: false,
            BusyStatus: "Busy",
            Sensitivity: "Normal",
            Subject: "Meeting",
            Location: "Room",
            IsManagedCopy: true,
            ManagedSourceGlobalAppointmentId: "source-global",
            ManagedSyncGroupId: syncGroupId,
            ManagedSourceAccountId: sourceAccount,
            ManagedCopyType: "Full");
}
