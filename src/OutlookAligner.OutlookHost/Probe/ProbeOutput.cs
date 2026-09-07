using System.Text.Json;
using OutlookAligner.Outlook.Contracts;

namespace OutlookAligner.OutlookHost.Probe;

internal static class ProbeOutput
{
    internal static void Write(OutlookProbeResult result, ProbeOptions options)
    {
        ArgumentNullException.ThrowIfNull(result);
        ArgumentNullException.ThrowIfNull(options);

        if (options.Json)
        {
            Console.WriteLine(JsonSerializer.Serialize(
                result,
                new JsonSerializerOptions { WriteIndented = true }));
            return;
        }

        Console.WriteLine("Outlook Aligner — Phase 1 read-only probe");
        Console.WriteLine($"Protocol: {result.ProtocolVersion}");
        Console.WriteLine($"Profile: {result.ProfileName ?? "(unavailable)"}");
        Console.WriteLine($"Window: {result.WindowStartLocal:g} -> {result.WindowEndLocal:g}");
        Console.WriteLine($"Stores: {result.Stores.Count}");
        Console.WriteLine($"Accounts: {result.Accounts.Count}");
        Console.WriteLine();

        foreach (var account in result.Accounts)
        {
            Console.WriteLine($"Account: {account.DisplayName}");
            Console.WriteLine($"  SMTP: {account.SmtpAddress ?? "(unavailable)"}");
            Console.WriteLine($"  Type: {account.AccountType}");
            Console.WriteLine($"  Store: {account.StoreDisplayName ?? "(unavailable)"}");
            Console.WriteLine($"  Calendar: {(account.CalendarAvailable ? "available" : "unavailable")}");
            Console.WriteLine($"  Events in window: {account.EventCount}");

            if (account.Error is not null)
            {
                Console.WriteLine($"  Error: {account.Error}");
            }

            if (options.IncludeDetails)
            {
                foreach (var calendarEvent in account.Events)
                {
                    Console.WriteLine(
                        $"    {calendarEvent.StartLocal:g} -> {calendarEvent.EndLocal:g} | "
                        + $"{calendarEvent.Subject ?? "(no subject)"}");
                }
            }

            Console.WriteLine();
        }

        if (result.Warnings.Count > 0)
        {
            Console.WriteLine("Warnings:");
            foreach (var warning in result.Warnings)
            {
                Console.WriteLine($"- {warning}");
            }
        }

        if (!options.IncludeDetails)
        {
            Console.WriteLine("Event subjects and locations were withheld. Use --include-details explicitly to display them.");
        }
    }

    internal static void WriteUsage()
    {
        Console.WriteLine("Usage: OutlookAligner.OutlookHost.exe [--days N] [--json] [--include-details]");
        Console.WriteLine();
        Console.WriteLine($"  --days N           Scan 1-{ProbeOptions.MaximumDays} days from today (default {ProbeOptions.DefaultDays}).");
        Console.WriteLine("  --json             Emit the probe result as JSON.");
        Console.WriteLine("  --include-details  Include event subject and location in output. Off by default.");
        Console.WriteLine("  --help, -h         Show this help text without opening Outlook.");
        Console.WriteLine();
        Console.WriteLine("Phase 1 is read-only: the probe does not save, send, forward, move, or delete Outlook items.");
    }
}
