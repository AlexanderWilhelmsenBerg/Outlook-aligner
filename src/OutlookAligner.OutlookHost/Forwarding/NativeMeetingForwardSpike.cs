using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using OutlookAligner.OutlookHost.Com;
using OutlookInterop = Microsoft.Office.Interop.Outlook;

namespace OutlookAligner.OutlookHost.Forwarding;

internal static class NativeMeetingForwardSpike
{
    private const int MaximumMeetingRequestsPerFolder = 5000;
    private const string CalendarForwardCommandId = "Forward";

    internal static ForwardSpikeResult Run(ForwardSpikeOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);

        var warnings = new List<string>();
        var searchedFolders = new List<string>();
        var candidates = new List<ForwardCandidateInfo>();

        OutlookInterop.Application? application = null;
        OutlookInterop.NameSpace? session = null;
        OutlookInterop.Accounts? accounts = null;
        OutlookInterop.Account? sourceAccount = null;
        OutlookInterop.Store? deliveryStore = null;

        try
        {
            application = new OutlookInterop.Application();
            session = application.GetNamespace("MAPI");
            accounts = session.Accounts;

            sourceAccount = FindSourceAccount(accounts, options.SourceSmtp, warnings);
            if (sourceAccount is null)
            {
                return Failure(
                    5,
                    $"No unique Outlook account with SMTP address '{options.SourceSmtp}' was found.",
                    options,
                    searchedFolders,
                    candidates,
                    warnings);
            }

            deliveryStore = sourceAccount.DeliveryStore;
            if (deliveryStore is null)
            {
                return Failure(
                    5,
                    "The selected source account has no accessible delivery store.",
                    options,
                    searchedFolders,
                    candidates,
                    warnings);
            }

            if (options.Action == ForwardSpikeAction.ProbeCalendarCommand)
            {
                return ProbeCalendarForwardCommand(
                    session,
                    deliveryStore,
                    options,
                    searchedFolders,
                    candidates,
                    warnings);
            }

            SearchFolder(
                deliveryStore,
                OutlookInterop.OlDefaultFolders.olFolderInbox,
                options,
                searchedFolders,
                candidates,
                warnings);
            SearchFolder(
                deliveryStore,
                OutlookInterop.OlDefaultFolders.olFolderDeletedItems,
                options,
                searchedFolders,
                candidates,
                warnings);

            if (candidates.Count == 0)
            {
                return Failure(
                    5,
                    "No retained MeetingItem matching this GlobalAppointmentID was found in Inbox or Deleted Items. This is a valid Phase 2 reliability result; no vCalendar fallback was attempted.",
                    options,
                    searchedFolders,
                    candidates,
                    warnings);
            }

            if (candidates.Count > 1)
            {
                return Failure(
                    5,
                    "Multiple retained MeetingItems matched. The spike refuses to choose one automatically.",
                    options,
                    searchedFolders,
                    candidates,
                    warnings);
            }

            if (options.Action == ForwardSpikeAction.Inspect)
            {
                return Success(
                    "A retained native MeetingItem was recovered and correlated to the appointment. No forward object was created or sent.",
                    options,
                    searchedFolders,
                    candidates,
                    warnings);
            }

            return ExecuteNativeForward(
                session,
                sourceAccount,
                deliveryStore,
                candidates[0],
                options,
                searchedFolders,
                candidates,
                warnings);
        }
        finally
        {
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
                    warnings.Add("More than one Outlook account matched the requested SMTP address; the forwarding spike refuses the ambiguous selection.");
                    ComRelease.Release(match);
                    match = null;
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

    private static ForwardSpikeResult ProbeCalendarForwardCommand(
        OutlookInterop.NameSpace session,
        OutlookInterop.Store deliveryStore,
        ForwardSpikeOptions options,
        List<string> searchedFolders,
        List<ForwardCandidateInfo> candidates,
        List<string> warnings)
    {
        object? sourceObject = null;
        OutlookInterop.AppointmentItem? appointment = null;
        OutlookInterop.Inspector? inspector = null;
        object? commandBars = null;

        try
        {
            sourceObject = session.GetItemFromID(options.CalendarEntryId!, deliveryStore.StoreID);
            appointment = sourceObject as OutlookInterop.AppointmentItem;
            if (appointment is null)
            {
                return Failure(
                    5,
                    "The requested Calendar EntryID did not reopen as an AppointmentItem in the selected source store.",
                    options,
                    searchedFolders,
                    candidates,
                    warnings);
            }

            if (!string.Equals(
                    appointment.GlobalAppointmentID,
                    options.GlobalAppointmentId,
                    StringComparison.Ordinal))
            {
                return Failure(
                    5,
                    "The reopened AppointmentItem GlobalAppointmentID did not match the requested meeting. The command probe stopped without invoking any Outlook command.",
                    options,
                    searchedFolders,
                    candidates,
                    warnings);
            }

            if (appointment.MeetingStatus == OutlookInterop.OlMeetingStatus.olNonMeeting)
            {
                return Failure(
                    5,
                    "The requested Calendar item is not a meeting, so native meeting Forward capability is not applicable.",
                    options,
                    searchedFolders,
                    candidates,
                    warnings);
            }

            inspector = appointment.GetInspector;
            inspector.Display(false);

            commandBars = GetComProperty(inspector, "CommandBars");
            if (commandBars is null)
            {
                return Failure(
                    5,
                    "The Outlook appointment Inspector did not expose its built-in command surface. No command was executed.",
                    options,
                    searchedFolders,
                    candidates,
                    warnings);
            }

            var commandProbe = ProbeMsoCommand(commandBars, CalendarForwardCommandId, warnings);
            var subjectSuffix = options.IncludeDetails
                ? $" Subject: {appointment.Subject ?? "(no subject)"}."
                : string.Empty;

            var summary = commandProbe.IdentifierValid
                ? $"The built-in Outlook command '{CalendarForwardCommandId}' resolved for the accepted Calendar appointment. Visible={commandProbe.Visible}; Enabled={commandProbe.Enabled}. The command was not executed.{subjectSuffix}"
                : $"The built-in Outlook command identifier '{CalendarForwardCommandId}' did not resolve in this appointment Inspector. The command was not executed.{subjectSuffix}";

            return Success(
                summary,
                options,
                searchedFolders,
                candidates,
                warnings,
                commandProbe);
        }
        finally
        {
            ComRelease.Release(commandBars);

            if (inspector is not null)
            {
                try
                {
                    inspector.Close(OutlookInterop.OlInspectorClose.olDiscard);
                }
                catch (COMException exception)
                {
                    warnings.Add($"The temporary appointment Inspector could not be closed cleanly (HRESULT 0x{exception.ErrorCode:X8}).");
                }
            }

            ComRelease.Release(inspector);

            if (!ReferenceEquals(sourceObject, appointment))
            {
                ComRelease.Release(sourceObject);
            }

            ComRelease.Release(appointment);
        }
    }

    private static ForwardCommandProbeInfo ProbeMsoCommand(
        object commandBars,
        string commandId,
        List<string> warnings)
    {
        string? label;
        try
        {
            label = InvokeComMethod(commandBars, "GetLabelMso", commandId) as string;
        }
        catch (Exception exception) when (TryGetComException(exception, out var comException))
        {
            warnings.Add($"Built-in command '{commandId}' could not be resolved (HRESULT 0x{comException.ErrorCode:X8}).");
            return new ForwardCommandProbeInfo(commandId, false, null, false, false);
        }

        var visible = ReadMsoBoolean(commandBars, "GetVisibleMso", commandId, warnings);
        var enabled = ReadMsoBoolean(commandBars, "GetEnabledMso", commandId, warnings);

        return new ForwardCommandProbeInfo(commandId, true, label, visible, enabled);
    }

    private static bool ReadMsoBoolean(
        object commandBars,
        string methodName,
        string commandId,
        List<string> warnings)
    {
        try
        {
            var value = InvokeComMethod(commandBars, methodName, commandId);
            return value is bool boolean && boolean;
        }
        catch (Exception exception) when (TryGetComException(exception, out var comException))
        {
            warnings.Add($"{methodName} failed for built-in command '{commandId}' (HRESULT 0x{comException.ErrorCode:X8}).");
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

    private static void SearchFolder(
        OutlookInterop.Store deliveryStore,
        OutlookInterop.OlDefaultFolders folderKind,
        ForwardSpikeOptions options,
        List<string> searchedFolders,
        List<ForwardCandidateInfo> candidates,
        List<string> warnings)
    {
        OutlookInterop.MAPIFolder? folder = null;
        OutlookInterop.Items? items = null;
        OutlookInterop.Items? meetingRequests = null;
        object? current = null;

        try
        {
            folder = deliveryStore.GetDefaultFolder(folderKind);
            if (folder is null)
            {
                warnings.Add($"Default folder {folderKind} was unavailable.");
                return;
            }

            searchedFolders.Add(folder.Name);
            items = folder.Items;
            meetingRequests = items.Restrict("[MessageClass] = 'IPM.Schedule.Meeting.Request'");
            current = meetingRequests.GetFirst();
            var inspected = 0;

            while (current is not null && inspected < MaximumMeetingRequestsPerFolder)
            {
                inspected++;
                OutlookInterop.AppointmentItem? associatedAppointment = null;

                try
                {
                    if (current is OutlookInterop.MeetingItem meetingRequest)
                    {
                        associatedAppointment = meetingRequest.GetAssociatedAppointment(false);
                        if (associatedAppointment is not null
                            && string.Equals(
                                associatedAppointment.GlobalAppointmentID,
                                options.GlobalAppointmentId,
                                StringComparison.Ordinal))
                        {
                            candidates.Add(new ForwardCandidateInfo(
                                folder.Name,
                                meetingRequest.EntryID,
                                meetingRequest.MessageClass,
                                associatedAppointment.Start,
                                associatedAppointment.End,
                                options.IncludeDetails ? associatedAppointment.Subject : null));
                        }
                    }
                }
                catch (COMException exception)
                {
                    warnings.Add($"One retained meeting request could not be correlated (HRESULT 0x{exception.ErrorCode:X8}).");
                }
                finally
                {
                    ComRelease.Release(associatedAppointment);
                }

                var next = meetingRequests.GetNext();
                ComRelease.Release(current);
                current = next;
            }

            if (current is not null)
            {
                warnings.Add($"Stopped scanning {folder.Name} after {MaximumMeetingRequestsPerFolder} meeting requests to keep the spike bounded.");
            }
        }
        catch (COMException exception)
        {
            warnings.Add($"Could not inspect default folder {folderKind} (HRESULT 0x{exception.ErrorCode:X8}).");
        }
        finally
        {
            ComRelease.Release(current);
            ComRelease.Release(meetingRequests);
            ComRelease.Release(items);
            ComRelease.Release(folder);
        }
    }

    private static ForwardSpikeResult ExecuteNativeForward(
        OutlookInterop.NameSpace session,
        OutlookInterop.Account sourceAccount,
        OutlookInterop.Store deliveryStore,
        ForwardCandidateInfo candidate,
        ForwardSpikeOptions options,
        List<string> searchedFolders,
        List<ForwardCandidateInfo> candidates,
        List<string> warnings)
    {
        object? sourceObject = null;
        OutlookInterop.MeetingItem? sourceMeeting = null;
        OutlookInterop.MeetingItem? forwardedMeeting = null;
        OutlookInterop.Recipients? recipients = null;
        OutlookInterop.Recipient? recipient = null;

        try
        {
            sourceObject = session.GetItemFromID(candidate.EntryId, deliveryStore.StoreID);
            sourceMeeting = sourceObject as OutlookInterop.MeetingItem;
            if (sourceMeeting is null)
            {
                return Failure(
                    5,
                    "The correlated item could not be reopened as a MeetingItem.",
                    options,
                    searchedFolders,
                    candidates,
                    warnings);
            }

            forwardedMeeting = ((OutlookInterop._MeetingItem)sourceMeeting).Forward();
            if (forwardedMeeting is null)
            {
                return Failure(
                    5,
                    "MeetingItem.Forward() returned no forwarded meeting object.",
                    options,
                    searchedFolders,
                    candidates,
                    warnings);
            }

            recipients = forwardedMeeting.Recipients;
            recipient = recipients.Add(options.Recipient!);
            if (!recipients.ResolveAll())
            {
                ((OutlookInterop._MeetingItem)forwardedMeeting).Close(OutlookInterop.OlInspectorClose.olDiscard);
                return Failure(
                    6,
                    "The forwarding recipient could not be resolved. The unsent forwarded meeting was discarded.",
                    options,
                    searchedFolders,
                    candidates,
                    warnings);
            }

            forwardedMeeting.SendUsingAccount = sourceAccount;

            if (options.Action == ForwardSpikeAction.Prepare)
            {
                ((OutlookInterop._MeetingItem)forwardedMeeting).Close(OutlookInterop.OlInspectorClose.olDiscard);
                return Success(
                    "MeetingItem.Forward() succeeded, the recipient resolved, and the unsent native forward was discarded.",
                    options,
                    searchedFolders,
                    candidates,
                    warnings);
            }

            ((OutlookInterop._MeetingItem)forwardedMeeting).Send();
            return Success(
                "MeetingItem.Forward() succeeded and the native forwarded meeting was sent using the selected source Outlook account.",
                options,
                searchedFolders,
                candidates,
                warnings);
        }
        finally
        {
            ComRelease.Release(recipient);
            ComRelease.Release(recipients);
            ComRelease.Release(forwardedMeeting);

            if (!ReferenceEquals(sourceObject, sourceMeeting))
            {
                ComRelease.Release(sourceObject);
            }

            ComRelease.Release(sourceMeeting);
        }
    }

    private static ForwardSpikeResult Success(
        string summary,
        ForwardSpikeOptions options,
        List<string> searchedFolders,
        List<ForwardCandidateInfo> candidates,
        List<string> warnings,
        ForwardCommandProbeInfo? commandProbe = null)
        => new(
            0,
            summary,
            options.SourceSmtp,
            options.GlobalAppointmentId,
            options.Action,
            options.Recipient,
            searchedFolders,
            candidates,
            commandProbe,
            warnings);

    private static ForwardSpikeResult Failure(
        int exitCode,
        string summary,
        ForwardSpikeOptions options,
        List<string> searchedFolders,
        List<ForwardCandidateInfo> candidates,
        List<string> warnings)
        => new(
            exitCode,
            summary,
            options.SourceSmtp,
            options.GlobalAppointmentId,
            options.Action,
            options.Recipient,
            searchedFolders,
            candidates,
            null,
            warnings);
}
