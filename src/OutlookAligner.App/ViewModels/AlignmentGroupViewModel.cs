using OutlookAligner.Core.Alignment;

namespace OutlookAligner.App.ViewModels;

public sealed class AlignmentGroupViewModel
{
    public AlignmentGroupViewModel(LogicalEventGroup group)
    {
        Group = group;
    }

    public LogicalEventGroup Group { get; }

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

    public bool NeedsAttention => Group.State != AlignmentState.Aligned;
}
