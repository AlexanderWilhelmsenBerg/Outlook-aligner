using OutlookAligner.Core.Alignment;

namespace OutlookAligner.App.ViewModels;

public sealed class AlignmentGroupViewModel
{
    public AlignmentGroupViewModel(LogicalEventGroup group)
    {
        Group = group;
    }

    public LogicalEventGroup Group { get; }

    public string GroupKey => Group.GroupKey;

    public string Subject => Group.DisplaySubject;

    public string State => Group.State switch
    {
        AlignmentState.RecurrenceIdentityUnresolved => "Recurrence unresolved",
        AlignmentState.DetailsDifferent => "Details differ",
        _ => Group.State.ToString(),
    };

    public string Explanation => Group.Explanation;

    public string AccountSummary
        => string.Join(" · ", Group.Members
            .Select(member => member.AccountKey)
            .Distinct(StringComparer.OrdinalIgnoreCase));

    public string MissingSummary => Group.MissingAccountKeys.Count == 0
        ? "None"
        : string.Join(", ", Group.MissingAccountKeys);

    public string TimeSummary
    {
        get
        {
            var ordered = Group.Members.OrderBy(member => member.StartLocal).ToArray();
            if (ordered.Length == 0)
            {
                return "—";
            }

            var first = ordered[0];
            return Group.TimesDiffer
                ? $"{first.StartLocal:g} · time differs across copies"
                : $"{first.StartLocal:g} – {first.EndLocal:g}";
        }
    }

    public string MemberDetails => string.Join(
        Environment.NewLine,
        Group.Members
            .OrderBy(member => member.AccountKey, StringComparer.OrdinalIgnoreCase)
            .ThenBy(member => member.StartLocal)
            .Select(member =>
            {
                var managed = member.IsManagedCopy
                    ? $" · managed {member.ManagedCopyType ?? "copy"}"
                    : string.Empty;
                return $"{member.AccountKey}: {member.StartLocal:g} – {member.EndLocal:g}{managed}";
            }));

    public bool NeedsAttention => Group.State != AlignmentState.Aligned;

    public bool CanChooseAuthority
        => Group.State is not AlignmentState.RecurrenceIdentityUnresolved
            and not AlignmentState.Uncorrelated
            and not AlignmentState.Duplicate
            && Group.Members.Count > 0;
}

public sealed record AuthorityChoice(string AccountKey)
{
    public string DisplayName => AccountKey;
}
