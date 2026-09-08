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

        if (result.Action == ForwardSpikeAction.ProbeCalendarCommand)
        {
            if (result.CommandProbe is not null)
            {
                writer.WriteLine("Calendar Forward command probe:");
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
        Console.WriteLine("Invoke MeetingItem.Forward(), resolve a recipient, then discard without sending:");
        Console.WriteLine("  OutlookAligner.OutlookHost.exe --forward-spike --source-smtp ADDRESS --global-id ID --prepare --to ADDRESS");
        Console.WriteLine();
        Console.WriteLine("Actually send the native forwarded meeting:");
        Console.WriteLine(
            $"  OutlookAligner.OutlookHost.exe --forward-spike --source-smtp ADDRESS --global-id ID --send --to ADDRESS --confirm-send {ForwardSpikeOptions.ConfirmationToken}");
        Console.WriteLine();
        Console.WriteLine("Options:");
        Console.WriteLine("  --source-smtp ADDRESS       Select the Outlook account/store that received the meeting.");
        Console.WriteLine("  --global-id ID              GlobalAppointmentID copied from the read probe.");
        Console.WriteLine("  --probe-calendar-command    Open the exact accepted appointment and query Outlook's built-in Forward command state without executing it.");
        Console.WriteLine("  --entry-id ENTRY_ID         Calendar EntryID required by --probe-calendar-command.");
        Console.WriteLine("  --prepare                   Create a native retained-request forward, resolve the recipient, then discard it unsent.");
        Console.WriteLine("  --send                      Send the native retained-request forward. Requires the exact confirmation token.");
        Console.WriteLine("  --to ADDRESS                Recipient for prepare/send modes.");
        Console.WriteLine("  --include-details           Show the matched/verified meeting subject. Off by default.");
        Console.WriteLine("  --help, -h                  Show this help without opening Outlook.");
        Console.WriteLine();
        Console.WriteLine("No vCalendar fallback is used. The Calendar command probe never executes Forward; prepare mode discards the unsent retained-request forward.");
    }
}
