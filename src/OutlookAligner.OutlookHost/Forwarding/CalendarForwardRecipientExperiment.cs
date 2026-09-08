using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using OutlookAligner.OutlookHost.Com;
using OutlookInterop = Microsoft.Office.Interop.Outlook;

namespace OutlookAligner.OutlookHost.Forwarding;

internal static class CalendarForwardRecipientExperiment
{
    private const string ForwardCommandId = "Forward";

    internal static ForwardSpikeResult Run(ForwardSpikeOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);

        var warnings = new List<string>();
        OutlookInterop.Application? application = null;
        OutlookInterop.NameSpace? session = null;
        OutlookInterop.Accounts? accounts = null;
        OutlookInterop.Account? sourceAccount = null;
        OutlookInterop.Store? deliveryStore = null;
        object? sourceObject = null;
        OutlookInterop.AppointmentItem? appointment = null;
        OutlookInterop.Inspector? sourceInspector = null;
        object? commandBars = null;
        OutlookInterop.MeetingItem? forwardedMeeting = null;
        OutlookInterop.Recipients? recipients = null;
        OutlookInterop.Recipient? recipient = null;
        OutlookInterop.ItemEvents_10_ForwardEventHandler? forwardHandler = null;

        var eventRaised = false;
        var sent = false;

        try
        {
            application = new OutlookInterop.Application();
            session = application.GetNamespace("MAPI");
            accounts = session.Accounts;
            sourceAccount = FindSourceAccount(accounts, options.SourceSmtp, warnings);
            if (sourceAccount is null)
            {
                return Failure(5, $"No unique Outlook account with SMTP address '{options.SourceSmtp}' was found.", options, warnings);
            }

            deliveryStore = sourceAccount.DeliveryStore;
            if (deliveryStore is null)
            {
                return Failure(5, "The selected source account has no accessible delivery store.", options, warnings);
            }

            sourceObject = session.GetItemFromID(options.CalendarEntryId!, deliveryStore.StoreID);
            appointment = sourceObject as OutlookInterop.AppointmentItem;
            if (appointment is null)
            {
                return Failure(5, "The requested Calendar EntryID did not reopen as an AppointmentItem in the selected source store.", options, warnings);
            }

            if (!string.Equals(appointment.GlobalAppointmentID, options.GlobalAppointmentId, StringComparison.Ordinal))
            {
                return Failure(5, "The reopened AppointmentItem GlobalAppointmentID did not match the requested meeting. No Outlook command was executed.", options, warnings);
            }

            if (appointment.MeetingStatus == OutlookInterop.OlMeetingStatus.olNonMeeting)
            {
                return Failure(5, "The requested Calendar item is not a meeting, so native meeting Forward is not applicable.", options, warnings);
            }

            sourceInspector = appointment.GetInspector;
            sourceInspector.Display(false);
            commandBars = GetComProperty(sourceInspector, "CommandBars");
            if (commandBars is null)
            {
                return Failure(5, "The Outlook appointment Inspector did not expose its built-in command surface. No command was executed.", options, warnings);
            }

            var commandProbe = ProbeCommand(commandBars, warnings);
            if (!commandProbe.IdentifierValid || !commandProbe.Visible || !commandProbe.Enabled)
            {
                return Failure(
                    5,
                    $"The built-in Outlook Forward command is not safely executable for this appointment. IdentifierValid={commandProbe.IdentifierValid}; Visible={commandProbe.Visible}; Enabled={commandProbe.Enabled}.",
                    options,
                    warnings,
                    commandProbe);
            }

            void OnForward(object forward, ref bool cancel)
            {
                eventRaised = true;
                cancel = false;

                if (forward is OutlookInterop.MeetingItem meetingItem)
                {
                    forwardedMeeting = meetingItem;
                }
                else
                {
                    warnings.Add($"Outlook's Forward event supplied '{forward.GetType().FullName ?? forward.GetType().Name}' instead of a native MeetingItem.");
                }
            }

            forwardHandler = OnForward;
            ((OutlookInterop.ItemEvents_10_Event)appointment).Forward += forwardHandler;

            InvokeComMethod(commandBars, "ExecuteMso", ForwardCommandId);

            if (!eventRaised || forwardedMeeting is null)
            {
                return Failure(
                    5,
                    "Outlook executed the built-in Forward command but no native MeetingItem was captured. Any transient forward will be discarded if accessible.",
                    options,
                    warnings,
                    commandProbe);
            }

            recipients = forwardedMeeting.Recipients;
            if (recipients.Count != 0)
            {
                return Failure(
                    5,
                    $"The newly forwarded MeetingItem unexpectedly already contained {recipients.Count} recipient(s). The experiment refuses to modify or send it.",
                    options,
                    warnings,
                    commandProbe);
            }

            recipient = recipients.Add(options.Recipient!);
            if (!recipients.ResolveAll())
            {
                return Failure(
                    6,
                    "The requested forwarding recipient could not be resolved. The unsent forwarded meeting will be discarded.",
                    options,
                    warnings,
                    commandProbe);
            }

            if (recipients.Count != 1)
            {
                return Failure(
                    5,
                    $"Recipient preparation produced {recipients.Count} recipients instead of exactly one. The unsent forwarded meeting will be discarded.",
                    options,
                    warnings,
                    commandProbe);
            }

            forwardedMeeting.SendUsingAccount = sourceAccount;

            var subjectSuffix = options.IncludeDetails
                ? $" Subject: {appointment.Subject ?? "(no subject)"}."
                : string.Empty;

            if (options.Action == ForwardSpikeAction.SendCalendarCommand)
            {
                ((OutlookInterop._MeetingItem)forwardedMeeting).Send();
                sent = true;

                return new ForwardSpikeResult(
                    0,
                    $"Outlook's Calendar Forward created a native MeetingItem, exactly one recipient resolved, SendUsingAccount was pinned to the selected source account, and MeetingItem.Send() completed.{subjectSuffix}",
                    options.SourceSmtp,
                    options.GlobalAppointmentId,
                    options.Action,
                    options.Recipient,
                    Array.Empty<string>(),
                    Array.Empty<ForwardCandidateInfo>(),
                    commandProbe,
                    warnings);
            }

            return new ForwardSpikeResult(
                0,
                $"Outlook's Calendar Forward created a native MeetingItem, exactly one recipient resolved, and SendUsingAccount was pinned to the selected source account. The forwarded meeting was not sent and will be discarded.{subjectSuffix}",
                options.SourceSmtp,
                options.GlobalAppointmentId,
                options.Action,
                options.Recipient,
                Array.Empty<string>(),
                Array.Empty<ForwardCandidateInfo>(),
                commandProbe,
                warnings);
        }
        finally
        {
            if (appointment is not null && forwardHandler is not null)
            {
                try
                {
                    ((OutlookInterop.ItemEvents_10_Event)appointment).Forward -= forwardHandler;
                }
                catch (COMException exception)
                {
                    warnings.Add($"The temporary AppointmentItem.Forward handler could not be detached cleanly (HRESULT 0x{exception.ErrorCode:X8}).");
                }
            }

            if (forwardedMeeting is not null && !sent)
            {
                try
                {
                    ((OutlookInterop._MeetingItem)forwardedMeeting).Close(OutlookInterop.OlInspectorClose.olDiscard);
                }
                catch (COMException exception)
                {
                    warnings.Add($"The unsent forwarded MeetingItem could not be discarded cleanly (HRESULT 0x{exception.ErrorCode:X8}).");
                }
            }

            ComRelease.Release(recipient);
            ComRelease.Release(recipients);
            ComRelease.Release(forwardedMeeting);
            ComRelease.Release(commandBars);

            if (sourceInspector is not null)
            {
                try
                {
                    sourceInspector.Close(OutlookInterop.OlInspectorClose.olDiscard);
                }
                catch (COMException exception)
                {
                    warnings.Add($"The source appointment Inspector could not be closed cleanly (HRESULT 0x{exception.ErrorCode:X8}).");
                }
            }

            ComRelease.Release(sourceInspector);

            if (!ReferenceEquals(sourceObject, appointment))
            {
                ComRelease.Release(sourceObject);
            }

            ComRelease.Release(appointment);
            ComRelease.Release(deliveryStore);
            ComRelease.Release(sourceAccount);
            ComRelease.Release(accounts);
            ComRelease.Release(session);
            ComRelease.Release(application);
        }
    }

    private static OutlookInterop.Account? FindSourceAccount(
        OutlookInterop.Accounts accounts,
        string sourceSmtp,
        List<string> warnings)
    {
        OutlookInterop.Account? match = null;

        for (var index = 1; index <= accounts.Count; index++)
        {
            OutlookInterop.Account? account = null;
            try
            {
                account = accounts[index];
                if (account is null)
                {
                    continue;
                }

                string? smtpAddress;
                try
                {
                    smtpAddress = account.SmtpAddress;
                }
                catch (COMException exception)
                {
                    warnings.Add($"One account SMTP address could not be read (HRESULT 0x{exception.ErrorCode:X8}).");
                    continue;
                }

                if (!string.Equals(smtpAddress, sourceSmtp, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (match is not null)
                {
                    ComRelease.Release(match);
                    warnings.Add("More than one Outlook account matched the requested SMTP address; the experiment refuses the ambiguous selection.");
                    return null;
                }

                match = account;
                account = null;
            }
            finally
            {
                ComRelease.Release(account);
            }
        }

        return match;
    }

    private static ForwardCommandProbeInfo ProbeCommand(object commandBars, List<string> warnings)
    {
        string? label;
        try
        {
            label = InvokeComMethod(commandBars, "GetLabelMso", ForwardCommandId) as string;
        }
        catch (Exception exception) when (TryGetComException(exception, out var comException))
        {
            warnings.Add($"Built-in command '{ForwardCommandId}' could not be resolved (HRESULT 0x{comException.ErrorCode:X8}).");
            return new ForwardCommandProbeInfo(ForwardCommandId, false, null, false, false);
        }

        var visible = ReadMsoBoolean(commandBars, "GetVisibleMso", warnings);
        var enabled = ReadMsoBoolean(commandBars, "GetEnabledMso", warnings);
        return new ForwardCommandProbeInfo(ForwardCommandId, true, label, visible, enabled);
    }

    private static bool ReadMsoBoolean(object commandBars, string methodName, List<string> warnings)
    {
        try
        {
            return InvokeComMethod(commandBars, methodName, ForwardCommandId) is true;
        }
        catch (Exception exception) when (TryGetComException(exception, out var comException))
        {
            warnings.Add($"{methodName} failed for built-in command '{ForwardCommandId}' (HRESULT 0x{comException.ErrorCode:X8}).");
            return false;
        }
    }

    private static object? GetComProperty(object target, string propertyName)
        => target.GetType().InvokeMember(
            propertyName,
            BindingFlags.GetProperty,
            binder: null,
            target,
            args: null,
            CultureInfo.InvariantCulture);

    private static object? InvokeComMethod(object target, string methodName, params object?[] arguments)
        => target.GetType().InvokeMember(
            methodName,
            BindingFlags.InvokeMethod,
            binder: null,
            target,
            arguments,
            CultureInfo.InvariantCulture);

    private static bool TryGetComException(Exception exception, out COMException comException)
    {
        if (exception is COMException direct)
        {
            comException = direct;
            return true;
        }

        if (exception is TargetInvocationException { InnerException: COMException inner })
        {
            comException = inner;
            return true;
        }

        comException = null!;
        return false;
    }

    private static ForwardSpikeResult Failure(
        int exitCode,
        string summary,
        ForwardSpikeOptions options,
        List<string> warnings,
        ForwardCommandProbeInfo? commandProbe = null)
        => new(
            exitCode,
            summary,
            options.SourceSmtp,
            options.GlobalAppointmentId,
            options.Action,
            options.Recipient,
            Array.Empty<string>(),
            Array.Empty<ForwardCandidateInfo>(),
            commandProbe,
            warnings);
}
