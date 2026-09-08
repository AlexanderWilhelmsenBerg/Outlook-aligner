using OutlookAligner.Core.Alignment;
using Xunit;

namespace OutlookAligner.Core.Tests;

public sealed class MovePreviewPlannerTests
{
    [Fact]
    public void ManagedNonAuthorityCopyProducesMoveAction()
    {
        var preview = MovePreviewPlanner.Build(
            "group",
            "authority@example.invalid",
            [
                Member("authority@example.invalid", "authority", 9, 10, managed: false),
                Member("copy@example.invalid", "copy", 11, 12, managed: true),
            ]);

        Assert.True(preview.CanExecute);
        var action = Assert.Single(preview.Actions);
        Assert.Equal("copy@example.invalid", action.AccountKey);
        Assert.Equal(new DateTime(2026, 9, 8, 9, 0, 0), action.TargetStartLocal);
        Assert.Equal(new DateTime(2026, 9, 8, 10, 0, 0), action.TargetEndLocal);
        Assert.Empty(preview.BlockingReasons);
    }

    [Fact]
    public void AuthorityIsNeverProposedAsMoveTarget()
    {
        var preview = MovePreviewPlanner.Build(
            "group",
            "authority@example.invalid",
            [
                Member("authority@example.invalid", "authority", 9, 10, managed: true),
                Member("copy@example.invalid", "copy", 11, 12, managed: true),
            ]);

        Assert.DoesNotContain(preview.Actions, action => action.AccountKey == "authority@example.invalid");
    }

    [Fact]
    public void UnmanagedNonAuthorityMemberIsSkipped()
    {
        var preview = MovePreviewPlanner.Build(
            "group",
            "authority@example.invalid",
            [
                Member("authority@example.invalid", "authority", 9, 10, managed: false),
                Member("other@example.invalid", "other", 11, 12, managed: false),
            ]);

        Assert.False(preview.CanExecute);
        Assert.Empty(preview.Actions);
        var skip = Assert.Single(preview.Skips);
        Assert.Contains("not an Outlook Aligner-managed copy", skip.Reason, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AlreadyAlignedManagedCopyIsSkipped()
    {
        var preview = MovePreviewPlanner.Build(
            "group",
            "authority@example.invalid",
            [
                Member("authority@example.invalid", "authority", 9, 10, managed: false),
                Member("copy@example.invalid", "copy", 9, 10, managed: true),
            ]);

        Assert.False(preview.CanExecute);
        Assert.Empty(preview.Actions);
        Assert.Contains(preview.Skips, skip => skip.Reason.Contains("Already aligned", StringComparison.Ordinal));
    }

    [Fact]
    public void RecurringMemberBlocksMovePreview()
    {
        var preview = MovePreviewPlanner.Build(
            "group",
            "authority@example.invalid",
            [
                Member("authority@example.invalid", "authority", 9, 10, managed: false) with { IsRecurring = true },
                Member("copy@example.invalid", "copy", 11, 12, managed: true) with { IsRecurring = true },
            ]);

        Assert.False(preview.CanExecute);
        Assert.Empty(preview.Actions);
        Assert.Contains(preview.BlockingReasons, reason => reason.Contains("Recurring Move", StringComparison.Ordinal));
    }

    [Fact]
    public void DuplicateAccountBlocksMovePreview()
    {
        var preview = MovePreviewPlanner.Build(
            "group",
            "authority@example.invalid",
            [
                Member("authority@example.invalid", "authority", 9, 10, managed: false),
                Member("copy@example.invalid", "copy-one", 11, 12, managed: true),
                Member("copy@example.invalid", "copy-two", 11, 12, managed: true),
            ]);

        Assert.False(preview.CanExecute);
        Assert.Contains(preview.BlockingReasons, reason => reason.Contains("Duplicate members", StringComparison.Ordinal));
    }

    [Fact]
    public void MissingAuthorityBlocksMovePreview()
    {
        var preview = MovePreviewPlanner.Build(
            "group",
            "authority@example.invalid",
            [Member("copy@example.invalid", "copy", 11, 12, managed: true)]);

        Assert.False(preview.CanExecute);
        Assert.Empty(preview.Actions);
        Assert.Contains(preview.BlockingReasons, reason => reason.Contains("authority account has no observed member", StringComparison.OrdinalIgnoreCase));
    }

    private static MovePreviewMember Member(
        string account,
        string locator,
        int startHour,
        int endHour,
        bool managed)
        => new(
            account,
            locator,
            new DateTime(2026, 9, 8, startHour, 0, 0),
            new DateTime(2026, 9, 8, endHour, 0, 0),
            IsAllDay: false,
            IsRecurring: false,
            IsManagedCopy: managed);
}
