namespace OutlookAligner.OutlookHost.Probe;

internal sealed record ProbeOptions(
    int Days,
    bool Json,
    bool IncludeDetails,
    bool ShowHelp)
{
    internal const int DefaultDays = 90;
    internal const int MaximumDays = 3650;

    internal static bool TryParse(
        IReadOnlyList<string> args,
        out ProbeOptions options,
        out string? error)
    {
        var days = DefaultDays;
        var json = false;
        var includeDetails = false;
        var showHelp = false;

        for (var index = 0; index < args.Count; index++)
        {
            var argument = args[index];

            if (string.Equals(argument, "--json", StringComparison.OrdinalIgnoreCase))
            {
                json = true;
                continue;
            }

            if (string.Equals(argument, "--include-details", StringComparison.OrdinalIgnoreCase))
            {
                includeDetails = true;
                continue;
            }

            if (string.Equals(argument, "--help", StringComparison.OrdinalIgnoreCase)
                || string.Equals(argument, "-h", StringComparison.OrdinalIgnoreCase))
            {
                showHelp = true;
                continue;
            }

            if (argument.StartsWith("--days=", StringComparison.OrdinalIgnoreCase))
            {
                var value = argument["--days=".Length..];
                if (!TryParseDays(value, out days, out error))
                {
                    options = new ProbeOptions(DefaultDays, false, false, false);
                    return false;
                }

                continue;
            }

            if (string.Equals(argument, "--days", StringComparison.OrdinalIgnoreCase))
            {
                if (++index >= args.Count)
                {
                    options = new ProbeOptions(DefaultDays, false, false, false);
                    error = "--days requires an integer value.";
                    return false;
                }

                if (!TryParseDays(args[index], out days, out error))
                {
                    options = new ProbeOptions(DefaultDays, false, false, false);
                    return false;
                }

                continue;
            }

            options = new ProbeOptions(DefaultDays, false, false, false);
            error = $"Unknown argument: {argument}";
            return false;
        }

        options = new ProbeOptions(days, json, includeDetails, showHelp);
        error = null;
        return true;
    }

    private static bool TryParseDays(string value, out int days, out string? error)
    {
        if (!int.TryParse(value, out days) || days < 1 || days > MaximumDays)
        {
            error = $"Scan horizon must be an integer from 1 to {MaximumDays}.";
            return false;
        }

        error = null;
        return true;
    }
}
