namespace OutlookAligner.OutlookHost.Forwarding;

internal enum ForwardSpikeAction
{
    Inspect,
    ProbeCalendarCommand,
    Prepare,
    Send,
}

internal sealed record ForwardSpikeOptions(
    string SourceSmtp,
    string GlobalAppointmentId,
    ForwardSpikeAction Action,
    string? CalendarEntryId,
    string? Recipient,
    bool IncludeDetails,
    bool ShowHelp)
{
    internal const string ConfirmationToken = "SEND-NATIVE-MEETING";

    internal static bool IsRequested(IReadOnlyList<string> args)
        => args.Any(argument => string.Equals(
            argument,
            "--forward-spike",
            StringComparison.OrdinalIgnoreCase));

    internal static bool TryParse(
        IReadOnlyList<string> args,
        out ForwardSpikeOptions options,
        out string? error)
    {
        string? sourceSmtp = null;
        string? globalAppointmentId = null;
        string? calendarEntryId = null;
        string? recipient = null;
        string? confirmation = null;
        var includeDetails = false;
        var showHelp = false;
        var probeCalendarCommand = false;
        var prepare = false;
        var send = false;

        for (var index = 0; index < args.Count; index++)
        {
            var argument = args[index];

            if (string.Equals(argument, "--forward-spike", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (string.Equals(argument, "--probe-calendar-command", StringComparison.OrdinalIgnoreCase))
            {
                probeCalendarCommand = true;
                continue;
            }

            if (string.Equals(argument, "--prepare", StringComparison.OrdinalIgnoreCase))
            {
                prepare = true;
                continue;
            }

            if (string.Equals(argument, "--send", StringComparison.OrdinalIgnoreCase))
            {
                send = true;
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

            if (MatchesOption(argument, "--source-smtp"))
            {
                if (!TryReadValue(args, ref index, "--source-smtp", argument, out sourceSmtp, out error))
                {
                    options = Empty();
                    return false;
                }

                continue;
            }

            if (MatchesOption(argument, "--global-id"))
            {
                if (!TryReadValue(args, ref index, "--global-id", argument, out globalAppointmentId, out error))
                {
                    options = Empty();
                    return false;
                }

                continue;
            }

            if (MatchesOption(argument, "--entry-id"))
            {
                if (!TryReadValue(args, ref index, "--entry-id", argument, out calendarEntryId, out error))
                {
                    options = Empty();
                    return false;
                }

                continue;
            }

            if (MatchesOption(argument, "--to"))
            {
                if (!TryReadValue(args, ref index, "--to", argument, out recipient, out error))
                {
                    options = Empty();
                    return false;
                }

                continue;
            }

            if (MatchesOption(argument, "--confirm-send"))
            {
                if (!TryReadValue(args, ref index, "--confirm-send", argument, out confirmation, out error))
                {
                    options = Empty();
                    return false;
                }

                continue;
            }

            options = Empty();
            error = $"Unknown forwarding-spike argument: {argument}";
            return false;
        }

        if (showHelp)
        {
            options = new ForwardSpikeOptions(
                string.Empty,
                string.Empty,
                ForwardSpikeAction.Inspect,
                null,
                null,
                includeDetails,
                ShowHelp: true);
            error = null;
            return true;
        }

        var requestedActions = (probeCalendarCommand ? 1 : 0) + (prepare ? 1 : 0) + (send ? 1 : 0);
        if (requestedActions > 1)
        {
            options = Empty();
            error = "--probe-calendar-command, --prepare, and --send are mutually exclusive.";
            return false;
        }

        if (string.IsNullOrWhiteSpace(sourceSmtp))
        {
            options = Empty();
            error = "--source-smtp is required for the forwarding spike.";
            return false;
        }

        if (string.IsNullOrWhiteSpace(globalAppointmentId))
        {
            options = Empty();
            error = "--global-id is required for the forwarding spike.";
            return false;
        }

        var action = send
            ? ForwardSpikeAction.Send
            : prepare
                ? ForwardSpikeAction.Prepare
                : probeCalendarCommand
                    ? ForwardSpikeAction.ProbeCalendarCommand
                    : ForwardSpikeAction.Inspect;

        if (action == ForwardSpikeAction.ProbeCalendarCommand
            && string.IsNullOrWhiteSpace(calendarEntryId))
        {
            options = Empty();
            error = "--entry-id is required with --probe-calendar-command.";
            return false;
        }

        if (action != ForwardSpikeAction.ProbeCalendarCommand && calendarEntryId is not null)
        {
            options = Empty();
            error = "--entry-id is valid only together with --probe-calendar-command.";
            return false;
        }

        if (action is ForwardSpikeAction.Prepare or ForwardSpikeAction.Send
            && string.IsNullOrWhiteSpace(recipient))
        {
            options = Empty();
            error = "--to is required with --prepare or --send.";
            return false;
        }

        if (action == ForwardSpikeAction.Send
            && !string.Equals(confirmation, ConfirmationToken, StringComparison.Ordinal))
        {
            options = Empty();
            error = $"--send requires --confirm-send {ConfirmationToken}.";
            return false;
        }

        if (action != ForwardSpikeAction.Send && confirmation is not null)
        {
            options = Empty();
            error = "--confirm-send is valid only together with --send.";
            return false;
        }

        options = new ForwardSpikeOptions(
            sourceSmtp,
            globalAppointmentId,
            action,
            calendarEntryId,
            recipient,
            includeDetails,
            ShowHelp: false);
        error = null;
        return true;
    }

    private static bool TryReadValue(
        IReadOnlyList<string> args,
        ref int index,
        string optionName,
        string argument,
        out string? value,
        out string? error)
    {
        value = null;
        error = null;

        if (argument.StartsWith(optionName + "=", StringComparison.OrdinalIgnoreCase))
        {
            value = argument[(optionName.Length + 1)..];
            if (string.IsNullOrWhiteSpace(value))
            {
                error = $"{optionName} requires a value.";
                return false;
            }

            return true;
        }

        if (++index >= args.Count
            || string.IsNullOrWhiteSpace(args[index])
            || args[index].StartsWith("--", StringComparison.Ordinal))
        {
            error = $"{optionName} requires a value.";
            return false;
        }

        value = args[index];
        return true;
    }

    private static bool MatchesOption(string argument, string optionName)
        => string.Equals(argument, optionName, StringComparison.OrdinalIgnoreCase)
            || argument.StartsWith(optionName + "=", StringComparison.OrdinalIgnoreCase);

    private static ForwardSpikeOptions Empty()
        => new(
            string.Empty,
            string.Empty,
            ForwardSpikeAction.Inspect,
            null,
            null,
            IncludeDetails: false,
            ShowHelp: false);
}
