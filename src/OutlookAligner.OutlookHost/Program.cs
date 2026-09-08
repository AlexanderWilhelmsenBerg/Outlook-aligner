using System.Runtime.InteropServices;
using OutlookAligner.OutlookHost.Com;
using OutlookAligner.OutlookHost.Forwarding;
using OutlookAligner.OutlookHost.Probe;

namespace OutlookAligner.OutlookHost;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        if (args.Length == 1
            && string.Equals(args[0], "--interop-check", StringComparison.OrdinalIgnoreCase))
        {
            try
            {
                Console.WriteLine(InteropDependencyCheck.Run());
                return 0;
            }
            catch (FileNotFoundException exception)
            {
                WriteMissingDependency(exception);
                return 4;
            }
            catch (Exception exception)
            {
                Console.Error.WriteLine($"Outlook interop check failed: {exception.GetType().Name}.");
                return 4;
            }
        }

        if (ForwardSpikeOptions.IsRequested(args))
        {
            return RunForwardingSpike(args);
        }

        if (!ProbeOptions.TryParse(args, out var options, out var error))
        {
            Console.Error.WriteLine(error);
            Console.Error.WriteLine();
            ProbeOutput.WriteUsage();
            return 2;
        }

        if (options.ShowHelp)
        {
            ProbeOutput.WriteUsage();
            return 0;
        }

        try
        {
            var result = OutlookProbeRunner.Run(options);
            ProbeOutput.Write(result, options);
            return 0;
        }
        catch (FileNotFoundException exception)
        {
            WriteMissingDependency(exception);
            return 4;
        }
        catch (COMException exception)
        {
            Console.Error.WriteLine(
                $"Outlook COM probe failed (HRESULT 0x{exception.ErrorCode:X8}). "
                + "Confirm that Classic Outlook is installed and a profile is configured.");
            return 3;
        }
        catch (Exception exception)
        {
            Console.Error.WriteLine($"Outlook probe failed: {exception.GetType().Name}.");
            return 4;
        }
    }

    private static int RunForwardingSpike(string[] args)
    {
        if (!ForwardSpikeOptions.TryParse(args, out var options, out var error))
        {
            Console.Error.WriteLine(error);
            Console.Error.WriteLine();
            ForwardSpikeOutput.WriteUsage();
            return 2;
        }

        if (options.ShowHelp)
        {
            ForwardSpikeOutput.WriteUsage();
            return 0;
        }

        try
        {
            var result = options.Action == ForwardSpikeAction.PrepareCalendarCommand
                ? CalendarForwardCommandExperiment.Run(options)
                : NativeMeetingForwardSpike.Run(options);
            ForwardSpikeOutput.Write(result, options.IncludeDetails);
            return result.ExitCode;
        }
        catch (FileNotFoundException exception)
        {
            WriteMissingDependency(exception);
            return 4;
        }
        catch (COMException exception)
        {
            Console.Error.WriteLine(
                $"Outlook native-forwarding spike failed (HRESULT 0x{exception.ErrorCode:X8}). "
                + "No vCalendar fallback was attempted.");
            return 3;
        }
        catch (Exception exception)
        {
            Console.Error.WriteLine($"Outlook native-forwarding spike failed: {exception.GetType().Name}.");
            return 4;
        }
    }

    private static void WriteMissingDependency(FileNotFoundException exception)
    {
        var missingDependency = string.IsNullOrWhiteSpace(exception.FileName)
            ? "(unknown assembly)"
            : Path.GetFileName(exception.FileName);

        Console.Error.WriteLine($"Outlook probe dependency missing: {missingDependency}.");
    }
}
