using OutlookAligner.App.Presentation;
using OutlookAligner.App.ViewModels;
using OutlookAligner.Outlook.Contracts;
using Xunit;

namespace OutlookAligner.OutlookHost.Tests;

public sealed class CalendarWebBridgeTests
{
    [Fact]
    public void RenderPayloadUsesOpaquePresentationIdAndExcludesOutlookTechnicalIds()
    {
        var row = CreateRow();
        var session = new CalendarPresentationSession();
        var observations = session.Reset([row]);
        var json = CalendarWebProtocol.SerializeRender(observations);

        Assert.Single(observations);
        Assert.Equal(32, observations[0].PresentationId.Length);
        Assert.Contains("Example meeting", json, StringComparison.Ordinal);
        Assert.DoesNotContain("ENTRY-SECRET", json, StringComparison.Ordinal);
        Assert.DoesNotContain("GLOBAL-SECRET", json, StringComparison.Ordinal);
        Assert.DoesNotContain("STORE-SECRET", json, StringComparison.Ordinal);
        Assert.DoesNotContain("SYNC-SECRET", json, StringComparison.Ordinal);
        Assert.True(session.TryResolve(observations[0].PresentationId, out var resolved));
        Assert.Same(row, resolved);
    }

    [Fact]
    public void ResetInvalidatesStalePresentationIds()
    {
        var session = new CalendarPresentationSession();
        var oldId = session.Reset([CreateRow()])[0].PresentationId;

        session.Reset([]);

        Assert.False(session.TryResolve(oldId, out _));
    }

    [Theory]
    [InlineData("{\"version\":1,\"type\":\"ready\"}", CalendarClientMessageKind.Ready)]
    [InlineData("{\"version\":1,\"type\":\"rangeChanged\",\"title\":\"September 2026\",\"start\":\"2026-09-01\",\"end\":\"2026-10-01\",\"visibleCount\":3}", CalendarClientMessageKind.RangeChanged)]
    [InlineData("{\"version\":1,\"type\":\"observationSelected\",\"presentationId\":\"opaque\"}", CalendarClientMessageKind.ObservationSelected)]
    public void ParsesKnownClientMessages(string json, CalendarClientMessageKind expectedKind)
    {
        Assert.True(CalendarWebProtocol.TryParseClientMessage(json, out var message));
        Assert.Equal(expectedKind, message!.Kind);
    }

    [Theory]
    [InlineData("not-json")]
    [InlineData("{\"version\":2,\"type\":\"ready\"}")]
    [InlineData("{\"version\":1,\"type\":\"unknown\"}")]
    [InlineData("{\"version\":1,\"type\":\"observationSelected\"}")]
    public void RejectsMalformedUnknownOrUnsupportedClientMessages(string json)
        => Assert.False(CalendarWebProtocol.TryParseClientMessage(json, out _));

    private static CalendarEventRowViewModel CreateRow()
    {
        var metadata = new ManagedCopyMetadataDto(
            ManagedCopyState.Valid,
            "SYNC-SECRET",
            "GLOBAL-SECRET",
            "source@example.com",
            OutlookAlignerManagedProperties.CopyTypeFull,
            OutlookAlignerManagedProperties.CurrentSchemaVersion);
        var calendarEvent = new CalendarEventDto(
            "ENTRY-SECRET",
            "GLOBAL-SECRET",
            new DateTime(2026, 9, 15, 10, 0, 0),
            new DateTime(2026, 9, 15, 11, 0, 0),
            false,
            false,
            "olApptNotRecurring",
            "olMeetingReceived",
            "olBusy",
            "olNormal",
            "Example meeting",
            "Room 1",
            metadata);
        var account = new OutlookAccountDto(
            "Work",
            "work@example.com",
            "Exchange",
            "Mailbox",
            "STORE-SECRET",
            "CALENDAR-SECRET",
            true,
            1,
            [calendarEvent],
            null);
        return new CalendarEventRowViewModel(account, calendarEvent);
    }
}
