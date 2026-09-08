namespace OutlookAligner.OutlookHost.Forwarding;

internal static class ForwardSpikeOutput
{
    internal static void Write(ForwardSpikeResult result, bool includeDetails)
    {
        ArgumentNullException.ThrowIfNull(result);

        var writer = result.Success ? Console.Out : Console.Error;
        writer.WriteLine("Outlook Aligner — Phase 2 native forwarding spike");
        writer.WriteLine($"Mode: {result.Action.ToString().ToLowerInvariant()}");
        writer.WriteLine($"Source account: {result.SourceSmtp}");
        writer.WriteLine($"GlobalAppointmentId: {result.GlobalAppointmentId}");

        if (result.Action is ForwardSpikeAction.ProbeCalendarCommand
            or ForwardSpikeAction.PrepareCalendarCommand
            or ForwardSpikeAction.PrepareCalendarRecipient)
        {
            if (result.CommandProbe is not null)
            {
                writer.WriteLine("Calendar Forward command:");
                writer.WriteLine($"  idMso: {result.CommandProbe.CommandId}");
                writer.WriteLine($"  Identifier valid: {result.CommandProbe.IdentifierValid}");
                writer.WriteLine($"  Label: {result.CommandProbe.Label ?? "(unavailable)"}");
                writer.WriteLine($"  Visible: {result.CommandProbe.Visible}");
                writer.WriteLine($"  Enabled: {result.CommandProbe.Enabled}");
            }
        }
        else
        {
            writer.WriteLine($"Searched folders: {string.Join(", ", result.SearchedFolders)}");
            writer.WriteLine($"Matching retained meeting requests: {result.Candidates.Count}");

            foreach (var candidate in result.Candidates)
            {
                writer.WriteLine(
                    $"- {candidate.FolderName} | {candidate.StartLocal:g} -> {candidate.EndLocal:g} | {candidate.MessageClass}");

                if (includeDetails)
                {
                    writer.WriteLine($"  Subject: {candidate.Subject ?? "(no subject)"}");
                }
            }
        }

        if (result.Recipient is not null)
        {
            writer.WriteLine($"Recipient: {result.Recipient}");
        }

        writer.WriteLine($"Result: {result.Summary}");

        if (result.Warnings.Count > 0)
        {
            writer.WriteLine("Warnings:");
            foreach (var warning in result.Warnings)
            {
                writer.WriteLine($"- {warning}");
            }
        }
    }

    internal static void WriteUsage()
    {
        Console.WriteLine("Phase 2 native meeting-forwarding spike");
        Console.WriteLine();
        Console.WriteLine("Inspect whether a retained native MeetingItem can be recovered:");
        Console.WriteLine("  OutlookAligner.OutlookHost.exe --forward-spike --source-smtp ADDRESS --global-id ID");
        Console.WriteLine();
        Console.WriteLine("Probe the accepted Calendar item's built-in Forward command without executing it:");
        Console.WriteLine("  OutlookAligner.OutlookHost.exe --forward-spike --source-smtp ADDRESS --global-id ID --probe-calendar-command --entry-id ENTRY_ID");
        Console.WriteLine();
        Console.WriteLine("Execute Calendar Forward, observe the native MeetingItem, and cancel before completion:");
        Console.WriteLine("  OutlookAligner.OutlookHost.exe --forward-spike --source-smtp ADDRESS --global-id ID --prepare-calendar-command --entry-id ENTRY_ID");
        Console.WriteLine();
        Console.WriteLine("Create the Calendar Forward, resolve exactly one recipient, pin the source account, then discard unsent:");
        Console.WriteLine("  OutlookAligner.OutlookHost.exe --forward-spike --source-smtp ADDRESS --global-id ID --prepare-calendar-recipient --entry-id ENTRY_ID --to ADDRESS");
        Console.WriteLine();
        Console.WriteLine("Invoke retained MeetingItem.Forward(), resolve a recipient, then discard without sending:");
        Console.WriteLine("  OutlookAligner.OutlookHost.exe --forward-spike --source-smtp ADDRESS --global-id ID --prepare --to ADDRESS");
        Console.WriteLine();
        Console.WriteLine("Actually send the retained-request native forwarded meeting:");
        Console.WriteLine(
            $"  OutlookAligner.OutlookHost.exe --forward-spike --source-smtp ADDRESS --global-id ID --send --to ADDRESS --confirm-send {ForwardSpikeOptions.ConfirmationToken}");
        Console.WriteLine();
        Console.WriteLine("Options:");
        Console.WriteLine("  --source-smtp ADDRESS          Select the Outlook account/store that received the meeting.");
        Console.WriteLine("  --global-id ID                 GlobalAppointmentID copied from the read probe.");
        Console.WriteLine("  --probe-calendar-command       Query Outlook's built-in Forward command state without executing it.");
        Console.WriteLine("  --prepare-calendar-command     Execute Calendar Forward and cancel inside AppointmentItem.Forward before completion.");
        Console.WriteLine("  --prepare-calendar-recipient   Allow Calendar Forward to create its native MeetingItem, resolve one recipient, then discard unsent.");
        Console.WriteLine("  --entry-id ENTRY_ID            Calendar EntryID required by Calendar command modes.");
        Console.WriteLine("  --prepare                      Prepare/discard through a retained native MeetingItem.");
        Console.WriteLine("  --send                         Send through the retained native MeetingItem path. Requires exact confirmation token.");
        Console.WriteLine("  --to ADDRESS                   Recipient for retained prepare/send and Calendar recipient-prepare modes.");
        Console.WriteLine("  --include-details              Show the matched/verified meeting subject. Off by default.");
        Console.WriteLine("  --help, -h                     Show this help without opening Outlook.");
        Console.WriteLine();
        Console.WriteLine("No vCalendar fallback is used. Calendar recipient-prepare resolves one recipient but never calls Send() or Save().");
    }
}
