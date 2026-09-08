namespace OutlookAligner.Core.Alignment;

public sealed record MovePreviewMember(
    string AccountKey,
    string LocatorKey,
    DateTime StartLocal,
    DateTime EndLocal,
    bool IsAllDay,
    bool IsRecurring,
    bool IsManagedCopy,
    string? ManagedSyncGroupId = null,
    string? ManagedSourceAccountId = null);

public sealed record MovePreviewAction(
    string AccountKey,
    string LocatorKey,
    DateTime CurrentStartLocal,
    DateTime CurrentEndLocal,
    DateTime TargetStartLocal,
    DateTime TargetEndLocal,
    bool TargetIsAllDay);

public sealed record MovePreviewSkip(
    string AccountKey,
    string LocatorKey,
    string Reason);

public sealed record MoveSelectedPreview(
    string GroupKey,
    string AuthorityAccountKey,
    IReadOnlyList<MovePreviewAction> Actions,
    IReadOnlyList<MovePreviewSkip> Skips,
    IReadOnlyList<string> BlockingReasons)
{
    public bool CanExecute => BlockingReasons.Count == 0 && Actions.Count > 0;
}

public static class MovePreviewPlanner
{
    public static MoveSelectedPreview Build(
        string groupKey,
        string authorityAccountKey,
        IEnumerable<MovePreviewMember> members)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(groupKey);
        ArgumentException.ThrowIfNullOrWhiteSpace(authorityAccountKey);
        ArgumentNullException.ThrowIfNull(members);

        var memberArray = members.ToArray();
        var blockingReasons = new List<string>();
        var skips = new List<MovePreviewSkip>();
        var actions = new List<MovePreviewAction>();

        if (memberArray.Length == 0)
        {
            blockingReasons.Add("The logical event has no observed calendar members.");
            return CreateResult();
        }

        var duplicateAccounts = memberArray
            .GroupBy(member => member.AccountKey, StringComparer.OrdinalIgnoreCase)
            .Where(group => group.Count() > 1)
            .Select(group => group.Key)
            .OrderBy(account => account, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        if (duplicateAccounts.Length > 0)
        {
            blockingReasons.Add($"Duplicate members must be resolved first: {string.Join(", ", duplicateAccounts)}.");
        }

        AddManagedMetadataBlockingReasons(memberArray, blockingReasons);

        var authorityMembers = memberArray
            .Where(member => string.Equals(member.AccountKey, authorityAccountKey, StringComparison.OrdinalIgnoreCase))
            .ToArray();

        if (authorityMembers.Length != 1)
        {
            blockingReasons.Add(authorityMembers.Length == 0
                ? "The selected authority account has no observed member."
                : "The selected authority account has more than one observed member.");
            return CreateResult();
        }

        if (blockingReasons.Count > 0)
        {
            return CreateResult();
        }

        var authority = authorityMembers[0];
        if (authority.IsRecurring || memberArray.Any(member => member.IsRecurring))
        {
            blockingReasons.Add("Recurring Move is disabled until recurrence write identity is proven.");
            return CreateResult();
        }

        foreach (var member in memberArray.Where(member =>
                     !string.Equals(member.AccountKey, authorityAccountKey, StringComparison.OrdinalIgnoreCase)))
        {
            if (!member.IsManagedCopy)
            {
                skips.Add(new MovePreviewSkip(
                    member.AccountKey,
                    member.LocatorKey,
                    "Not an Outlook Aligner-managed copy; it will not be modified."));
                continue;
            }

            if (member.StartLocal == authority.StartLocal
                && member.EndLocal == authority.EndLocal
                && member.IsAllDay == authority.IsAllDay)
            {
                skips.Add(new MovePreviewSkip(
                    member.AccountKey,
                    member.LocatorKey,
                    "Already aligned with the authoritative time."));
                continue;
            }

            actions.Add(new MovePreviewAction(
                member.AccountKey,
                member.LocatorKey,
                member.StartLocal,
                member.EndLocal,
                authority.StartLocal,
                authority.EndLocal,
                authority.IsAllDay));
        }

        return CreateResult();

        MoveSelectedPreview CreateResult()
            => new(
                groupKey,
                authorityAccountKey,
                actions.ToArray(),
                skips.ToArray(),
                blockingReasons.ToArray());
    }

    private static void AddManagedMetadataBlockingReasons(
        MovePreviewMember[] members,
        ICollection<string> blockingReasons)
    {
        var managedCopies = members.Where(member => member.IsManagedCopy).ToArray();
        if (managedCopies.Length == 0)
        {
            return;
        }

        if (managedCopies.Any(member =>
                string.IsNullOrWhiteSpace(member.ManagedSyncGroupId)
                || string.IsNullOrWhiteSpace(member.ManagedSourceAccountId)))
        {
            blockingReasons.Add("Managed-copy identity is incomplete; Move is blocked.");
            return;
        }

        var syncGroupCount = managedCopies
            .Select(member => member.ManagedSyncGroupId!.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Count();
        if (syncGroupCount > 1)
        {
            blockingReasons.Add("Managed copies disagree about their sync group; Move is blocked.");
        }

        var sourceAccountCount = managedCopies
            .Select(member => member.ManagedSourceAccountId!.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Count();
        if (sourceAccountCount > 1)
        {
            blockingReasons.Add("Managed copies disagree about their recorded source account; Move is blocked.");
        }
    }
}
