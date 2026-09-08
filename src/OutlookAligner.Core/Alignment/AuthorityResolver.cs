namespace OutlookAligner.Core.Alignment;

public enum AuthorityReason
{
    KnownOrigin,
    UserSelected,
    Inferred,
    Unknown,
}

public sealed record AuthorityResolution(
    string? AuthorityAccountKey,
    AuthorityReason Reason,
    string Explanation)
{
    public bool HasAuthority => !string.IsNullOrWhiteSpace(AuthorityAccountKey);
}

public static class AuthorityResolver
{
    public static AuthorityResolution Resolve(LogicalEventGroup group)
    {
        ArgumentNullException.ThrowIfNull(group);

        if (group.State is AlignmentState.RecurrenceIdentityUnresolved
            or AlignmentState.Uncorrelated
            or AlignmentState.Duplicate
            or AlignmentState.Conflict)
        {
            return Unknown("Authority cannot be inferred while the logical event identity is unresolved or conflicting.");
        }

        var managedCopies = group.Members.Where(member => member.IsManagedCopy).ToArray();
        if (managedCopies.Length == 0)
        {
            return Unknown("No validated Outlook Aligner-managed copy records the origin account.");
        }

        var sourceAccounts = managedCopies
            .Select(member => Normalize(member.ManagedSourceAccountId))
            .Where(source => source is not null)
            .Select(source => source!)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        if (sourceAccounts.Length != 1)
        {
            return Unknown(sourceAccounts.Length == 0
                ? "Managed copies do not expose a validated origin account."
                : "Managed copies disagree about the origin account.");
        }

        var syncGroupIds = managedCopies
            .Select(member => Normalize(member.ManagedSyncGroupId))
            .Where(syncGroupId => syncGroupId is not null)
            .Select(syncGroupId => syncGroupId!)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        if (syncGroupIds.Length > 1)
        {
            return Unknown("Managed copies disagree about their sync group.");
        }

        var sourceAccount = sourceAccounts[0];
        var sourceMembers = group.Members
            .Where(member => string.Equals(member.AccountKey, sourceAccount, StringComparison.OrdinalIgnoreCase))
            .ToArray();

        if (sourceMembers.Length != 1 || sourceMembers[0].IsManagedCopy)
        {
            return Unknown("The recorded origin account does not resolve to exactly one observed source appointment.");
        }

        return new AuthorityResolution(
            sourceMembers[0].AccountKey,
            AuthorityReason.KnownOrigin,
            "Validated Outlook Aligner copy metadata identifies the observed source appointment.");
    }

    private static AuthorityResolution Unknown(string explanation)
        => new(null, AuthorityReason.Unknown, explanation);

    private static string? Normalize(string? value)
        => string.IsNullOrWhiteSpace(value) ? null : value.Trim();
}
