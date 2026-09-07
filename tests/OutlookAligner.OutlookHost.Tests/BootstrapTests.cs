using System.Globalization;
using OutlookAligner.Outlook.Contracts;
using OutlookAligner.OutlookHost.Calendars;
using OutlookAligner.OutlookHost.Probe;
using Xunit;

namespace OutlookAligner.OutlookHost.Tests;

public sealed class BootstrapTests
{
    [Fact]
    public void ProtocolVersionStartsAtOne()
    {
        Assert.Equal(1, OutlookHostProtocol.Version);
    }

    [Fact]
    public void ProbeOptionsDefaultToNinetyDaysAndPrivacySafeOutput()
    {
        var parsed = ProbeOptions.TryParse([], out var options, out var error);

        Assert.True(parsed);
        Assert.Null(error);
        Assert.Equal(90, options.Days);
        Assert.False(options.Json);
        Assert.False(options.IncludeDetails);
        Assert.False(options.ShowHelp);
    }

    [Fact]
    public void ProbeOptionsAcceptExplicitDaysAndJson()
    {
        var parsed = ProbeOptions.TryParse(
            ["--days", "14", "--json", "--include-details"],
            out var options,
            out var error);

        Assert.True(parsed);
        Assert.Null(error);
        Assert.Equal(14, options.Days);
        Assert.True(options.Json);
        Assert.True(options.IncludeDetails);
    }

    [Fact]
    public void ProbeOptionsRejectUnboundedDays()
    {
        var parsed = ProbeOptions.TryParse(["--days=0"], out _, out var error);

        Assert.False(parsed);
        Assert.NotNull(error);
    }

    [Fact]
    public void DateFilterUsesHalfOpenBoundedWindow()
    {
        var start = new DateTime(2026, 9, 7, 0, 0, 0, DateTimeKind.Local);
        var end = start.AddDays(90);

        var filter = OutlookDateFilter.Build(start, end, CultureInfo.InvariantCulture);

        Assert.Contains("[End] >=", filter, StringComparison.Ordinal);
        Assert.Contains("[Start] <", filter, StringComparison.Ordinal);
        Assert.DoesNotContain("<=", filter, StringComparison.Ordinal);
    }
}
