using OutlookAligner.OutlookHost.Forwarding;
using Xunit;

namespace OutlookAligner.OutlookHost.Tests;

public sealed class ForwardSpikeOptionsTests
{
    [Fact]
    public void ForwardSpikeDefaultsToReadOnlyInspect()
    {
        var parsed = ForwardSpikeOptions.TryParse(
            ["--forward-spike", "--source-smtp", "source@example.invalid", "--global-id", "global-id"],
            out var options,
            out var error);

        Assert.True(parsed);
        Assert.Null(error);
        Assert.Equal(ForwardSpikeAction.Inspect, options.Action);
        Assert.Null(options.CalendarEntryId);
        Assert.Null(options.Recipient);
        Assert.False(options.IncludeDetails);
    }

    [Fact]
    public void CalendarCommandProbeRequiresEntryIdAndRemainsSeparateFromWrites()
    {
        var parsed = ForwardSpikeOptions.TryParse(
            [
                "--forward-spike",
                "--source-smtp", "source@example.invalid",
                "--global-id", "global-id",
                "--probe-calendar-command",
                "--entry-id", "calendar-entry-id",
            ],
            out var options,
            out var error);

        Assert.True(parsed);
        Assert.Null(error);
        Assert.Equal(ForwardSpikeAction.ProbeCalendarCommand, options.Action);
        Assert.Equal("calendar-entry-id", options.CalendarEntryId);
        Assert.Null(options.Recipient);
    }

    [Fact]
    public void CancelledCalendarCommandPrepareRequiresEntryIdAndRejectsRecipient()
    {
        var parsed = ForwardSpikeOptions.TryParse(
            [
                "--forward-spike",
                "--source-smtp", "source@example.invalid",
                "--global-id", "global-id",
                "--prepare-calendar-command",
                "--entry-id", "calendar-entry-id",
            ],
            out var options,
            out var error);

        Assert.True(parsed);
        Assert.Null(error);
        Assert.Equal(ForwardSpikeAction.PrepareCalendarCommand, options.Action);
        Assert.Equal("calendar-entry-id", options.CalendarEntryId);
        Assert.Null(options.Recipient);

        var withRecipient = ForwardSpikeOptions.TryParse(
            [
                "--forward-spike",
                "--source-smtp", "source@example.invalid",
                "--global-id", "global-id",
                "--prepare-calendar-command",
                "--entry-id", "calendar-entry-id",
                "--to", "target@example.invalid",
            ],
            out _,
            out var recipientError);

        Assert.False(withRecipient);
        Assert.Contains("--to is not valid", recipientError, StringComparison.Ordinal);
    }

    [Fact]
    public void CalendarRecipientPrepareRequiresEntryIdAndRecipient()
    {
        var parsed = ForwardSpikeOptions.TryParse(
            [
                "--forward-spike",
                "--source-smtp", "source@example.invalid",
                "--global-id", "global-id",
                "--prepare-calendar-recipient",
                "--entry-id", "calendar-entry-id",
                "--to", "target@example.invalid",
            ],
            out var options,
            out var error);

        Assert.True(parsed);
        Assert.Null(error);
        Assert.Equal(ForwardSpikeAction.PrepareCalendarRecipient, options.Action);
        Assert.Equal("calendar-entry-id", options.CalendarEntryId);
        Assert.Equal("target@example.invalid", options.Recipient);
    }

    [Fact]
    public void CalendarRecipientPrepareWithoutRecipientIsRejected()
    {
        var parsed = ForwardSpikeOptions.TryParse(
            [
                "--forward-spike",
                "--source-smtp", "source@example.invalid",
                "--global-id", "global-id",
                "--prepare-calendar-recipient",
                "--entry-id", "calendar-entry-id",
            ],
            out _,
            out var error);

        Assert.False(parsed);
        Assert.Contains("--to is required", error, StringComparison.Ordinal);
    }

    [Fact]
    public void CalendarCommandModeWithoutEntryIdIsRejected()
    {
        var parsed = ForwardSpikeOptions.TryParse(
            [
                "--forward-spike",
                "--source-smtp", "source@example.invalid",
                "--global-id", "global-id",
                "--probe-calendar-command",
            ],
            out _,
            out var error);

        Assert.False(parsed);
        Assert.Contains("--entry-id is required", error, StringComparison.Ordinal);
    }

    [Fact]
    public void EntryIdWithoutCalendarCommandModeIsRejected()
    {
        var parsed = ForwardSpikeOptions.TryParse(
            [
                "--forward-spike",
                "--source-smtp", "source@example.invalid",
                "--global-id", "global-id",
                "--entry-id", "calendar-entry-id",
            ],
            out _,
            out var error);

        Assert.False(parsed);
        Assert.Equal("--entry-id is valid only with a Calendar command mode.", error);
    }

    [Fact]
    public void CalendarCommandModesAreMutuallyExclusiveWithRetainedPrepare()
    {
        var parsed = ForwardSpikeOptions.TryParse(
            [
                "--forward-spike",
                "--source-smtp", "source@example.invalid",
                "--global-id", "global-id",
                "--prepare-calendar-recipient",
                "--entry-id", "calendar-entry-id",
                "--prepare",
                "--to", "target@example.invalid",
            ],
            out _,
            out var error);

        Assert.False(parsed);
        Assert.Contains("mutually exclusive", error, StringComparison.Ordinal);
    }

    [Fact]
    public void PrepareRequiresRecipientButNotSendConfirmation()
    {
        var parsed = ForwardSpikeOptions.TryParse(
            [
                "--forward-spike",
                "--source-smtp=source@example.invalid",
                "--global-id=global-id",
                "--prepare",
                "--to=target@example.invalid",
            ],
            out var options,
            out var error);

        Assert.True(parsed);
        Assert.Null(error);
        Assert.Equal(ForwardSpikeAction.Prepare, options.Action);
        Assert.Equal("target@example.invalid", options.Recipient);
    }

    [Fact]
    public void SendRequiresExactConfirmationToken()
    {
        var parsed = ForwardSpikeOptions.TryParse(
            [
                "--forward-spike",
                "--source-smtp", "source@example.invalid",
                "--global-id", "global-id",
                "--send",
                "--to", "target@example.invalid",
                "--confirm-send", ForwardSpikeOptions.ConfirmationToken,
            ],
            out var options,
            out var error);

        Assert.True(parsed);
        Assert.Null(error);
        Assert.Equal(ForwardSpikeAction.Send, options.Action);
    }

    [Fact]
    public void SendWithoutConfirmationIsRejected()
    {
        var parsed = ForwardSpikeOptions.TryParse(
            [
                "--forward-spike",
                "--source-smtp", "source@example.invalid",
                "--global-id", "global-id",
                "--send",
                "--to", "target@example.invalid",
            ],
            out _,
            out var error);

        Assert.False(parsed);
        Assert.Contains("--confirm-send", error, StringComparison.Ordinal);
    }

    [Fact]
    public void PrepareAndSendAreMutuallyExclusive()
    {
        var parsed = ForwardSpikeOptions.TryParse(
            [
                "--forward-spike",
                "--source-smtp", "source@example.invalid",
                "--global-id", "global-id",
                "--prepare",
                "--send",
                "--to", "target@example.invalid",
                "--confirm-send", ForwardSpikeOptions.ConfirmationToken,
            ],
            out _,
            out var error);

        Assert.False(parsed);
        Assert.Contains("mutually exclusive", error, StringComparison.Ordinal);
    }

    [Fact]
    public void HelpDoesNotRequireOutlookSelectionArguments()
    {
        var parsed = ForwardSpikeOptions.TryParse(
            ["--forward-spike", "--help"],
            out var options,
            out var error);

        Assert.True(parsed);
        Assert.Null(error);
        Assert.True(options.ShowHelp);
    }

    [Fact]
    public void MissingSourceValueIsRejectedInsteadOfConsumingNextOption()
    {
        var parsed = ForwardSpikeOptions.TryParse(
            ["--forward-spike", "--source-smtp", "--global-id", "global-id"],
            out _,
            out var error);

        Assert.False(parsed);
        Assert.Equal("--source-smtp requires a value.", error);
    }
}
