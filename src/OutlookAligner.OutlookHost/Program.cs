using System.Runtime.InteropServices;
using OutlookAligner.OutlookHost.Probe;

namespace OutlookAligner.OutlookHost;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
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
        catch (COMException exception)
        {
            Console.Error.WriteLine(
                $"Outlook COM probe failed (HRESULT 0x{exception.ErrorCode:X8}). "
                + "Confirm that Classic Outlook is installed and a profile is configured.");
            return 3;
        }
        catch (Exception exception)
        {
            Console.Error.WriteLine($"Outlook probe failed: {exception.GetType().Name}: {exception.Message}");
            return 4;
        }
    }
}
