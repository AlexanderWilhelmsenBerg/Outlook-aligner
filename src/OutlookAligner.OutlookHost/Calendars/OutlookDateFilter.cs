using System.Globalization;

namespace OutlookAligner.OutlookHost.Calendars;

internal static class OutlookDateFilter
{
    internal static string Build(DateTime startLocal, DateTime endLocal)
        => Build(startLocal, endLocal, CultureInfo.CurrentCulture);

    internal static string Build(DateTime startLocal, DateTime endLocal, CultureInfo culture)
    {
        ArgumentNullException.ThrowIfNull(culture);

        if (endLocal <= startLocal)
        {
            throw new ArgumentOutOfRangeException(nameof(endLocal), "The scan end must be after the scan start.");
        }

        var start = Escape(startLocal.ToString("g", culture));
        var end = Escape(endLocal.ToString("g", culture));
        return $"[End] >= '{start}' AND [Start] < '{end}'";
    }

    private static string Escape(string value)
        => value.Replace("'", "''", StringComparison.Ordinal);
}
