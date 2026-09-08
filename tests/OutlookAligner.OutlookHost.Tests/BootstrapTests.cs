using System.Globalization;
using OutlookAligner.Outlook.Contracts;
using OutlookAligner.OutlookHost.Calendars;
using OutlookAligner.OutlookHost.Com;
using OutlookAligner.OutlookHost.Probe;
using Xunit;

namespace OutlookAligner.OutlookHost.Tests;

public sealed class BootstrapTests
{
    [Fact]
    public void ProtocolVersionReflectsMeetingStatusContract()
    {
        Assert.Equal(3, OutlookHostProtocol.Version);
    }

    [Fact]
    public void ContractsAssemblyDoesNotReferenceOutlookInterop()
    {
        var references = typeof(OutlookHostProtocol).Assembly.GetReferencedAssemblies();

        Assert.DoesNotContain(
            references,
            reference => string.Equals(
                reference.Name,
                "Microsoft.Office.Interop.Outlook",
                StringComparison.Ordinal));
    }

    [Fact]
    public void InteropDependencyCheckResolvesMetadataWithoutOpeningOutlook()
    {
        var result = InteropDependencyCheck.Run();

        Assert.Equal("Outlook interop metadata resolved successfully.", result);
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
    public void ProbeOptionsAcceptMaximumHorizon()
    {
        var parsed = ProbeOptions.TryParse(["--days=3650"], out var options, out var error);

        Assert.True(parsed);
        Assert.Null(error);
        Assert.Equal(3650, options.Days);
    }

    [Theory]
    [InlineData("--days=0")]
    [InlineData("--days=-1")]
    [InlineData("--days=3651")]
    public void ProbeOptionsRejectOutOfRangeHorizons(string argument)
    {
        var parsed = ProbeOptions.TryParse([argument], out _, out var error);

        Assert.False(parsed);
        Assert.NotNull(error);
    }

    [Fact]
    public void ProbeOptionsRejectMissingDaysValue()
    {
        var parsed = ProbeOptions.TryParse(["--days"], out _, out var error);

        Assert.False(parsed);
        Assert.Equal("--days requires an integer value.", error);
    }

    [Fact]
    public void ProbeOptionsRejectUnknownArgument()
    {
        var parsed = ProbeOptions.TryParse(["--unexpected"], out _, out var error);

        Assert.False(parsed);
        Assert.Equal("Unknown argument: --unexpected", error);
    }

    [Fact]
    public void DateFilterUsesHalfOpenBoundedWindow()
    {
        var start = new DateTime(2026, 9, 7, 0, 0, 0, DateTimeKind.Local);
        var end = start.AddDays(90);

        var filter = OutlookDateFilter.Build(start, end, CultureInfo.InvariantCulture);

        Assert.Contains("[End] >", filter, StringComparison.Ordinal);
        Assert.Contains("[Start] <", filter, StringComparison.Ordinal);
        Assert.DoesNotContain("[End] >=", filter, StringComparison.Ordinal);
        Assert.DoesNotContain("<=", filter, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("en-US")]
    [InlineData("nb-NO")]
    public void DateFilterUsesRequestedCultureForOutlookDateFormatting(string cultureName)
    {
        var culture = CultureInfo.GetCultureInfo(cultureName);
        var start = new DateTime(2026, 9, 7, 13, 30, 0, DateTimeKind.Local);
        var end = start.AddDays(14);

        var filter = OutlookDateFilter.Build(start, end, culture);

        Assert.Contains(start.ToString("g", culture), filter, StringComparison.Ordinal);
        Assert.Contains(end.ToString("g", culture), filter, StringComparison.Ordinal);
        Assert.Contains("[End] >", filter, StringComparison.Ordinal);
        Assert.Contains("[Start] <", filter, StringComparison.Ordinal);
    }

    [Fact]
    public void JsonOutputRedactsEventDetailsUnlessExplicitlyEnabled()
    {
        var result = CreateProbeResult("Sensitive subject", "Sensitive location");

        var json = ProbeOutput.SerializeJson(result, includeDetails: false);

        Assert.DoesNotContain("Sensitive subject", json, StringComparison.Ordinal);
        Assert.DoesNotContain("Sensitive location", json, StringComparison.Ordinal);
        Assert.Contains("\"Subject\": null", json, StringComparison.Ordinal);
        Assert.Contains("\"Location\": null", json, StringComparison.Ordinal);
        Assert.Contains("\"MeetingStatus\": \"olMeetingReceived\"", json, StringComparison.Ordinal);
    }

    [Fact]
    public void JsonOutputIncludesEventDetailsWhenExplicitlyEnabled()
    {
        var result = CreateProbeResult("Expected subject", "Expected location");

        var json = ProbeOutput.SerializeJson(result, includeDetails: true);

        Assert.Contains("Expected subject", json, StringComparison.Ordinal);
        Assert.Contains("Expected location", json, StringComparison.Ordinal);
        Assert.Contains("\"MeetingStatus\": \"olMeetingReceived\"", json, StringComparison.Ordinal);
    }

    [Fact]
    public void ConsoleDetailFormattingIncludesSubjectAndLocation()
    {
        var calendarEvent = CreateProbeResult("Expected subject", "Expected location")
            .Accounts[0]
            .Events[0];

        var output = ProbeOutput.FormatEventDetails(calendarEvent);

        Assert.Contains("Expected subject", output, StringComparison.Ordinal);
        Assert.Contains("Location: Expected location", output, StringComparison.Ordinal);
    }

    private static OutlookProbeResult CreateProbeResult(string subject, string location)
    {
        var start = new DateTime(2026, 9, 7, 9, 0, 0, DateTimeKind.Local);
        var calendarEvent = new CalendarEventDto(
            EntryId: "entry-id",
            GlobalAppointmentId: "global-id",
            StartLocal: start,
            EndLocal: start.AddHours(1),
            IsAllDay: false,
            IsRecurring: false,
            RecurrenceState: "olApptNotRecurring",
            MeetingStatus: "olMeetingReceived",
            BusyStatus: "olBusy",
            Sensitivity: "olNormal",
            Subject: subject,
            Location: location);

        var account = new OutlookAccountDto(
            DisplayName: "Test Account",
            SmtpAddress: "test@example.invalid",
            AccountType: "olExchange",
            StoreDisplayName: "Test Store",
            StoreId: "store-id",
            CalendarEntryId: "calendar-id",
            CalendarAvailable: true,
            EventCount: 1,
            Events: [calendarEvent],
            Error: null);

        return new OutlookProbeResult(
            ProtocolVersion: OutlookHostProtocol.Version,
            CapturedAtUtc: new DateTime(2026, 9, 7, 7, 0, 0, DateTimeKind.Utc),
            WindowStartLocal: start.Date,
            WindowEndLocal: start.Date.AddDays(90),
            ProfileName: "Test Profile",
            Stores: [new OutlookStoreDto("Test Store", "store-id")],
            Accounts: [account],
            Warnings: []);
    }
}
