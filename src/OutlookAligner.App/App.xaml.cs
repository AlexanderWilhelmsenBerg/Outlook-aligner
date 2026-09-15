using Microsoft.UI.Xaml;

namespace OutlookAligner.App;

public partial class App : Application
{
    private Window? _window;

    public App()
    {
        StartupDiagnostics.Reset();
        StartupDiagnostics.Write("App constructor entered.");
        UnhandledException += OnUnhandledException;

        try
        {
            InitializeComponent();
            StartupDiagnostics.Write("Application resources initialized.");
        }
        catch (Exception exception)
        {
            StartupDiagnostics.WriteException("Application resource initialization failed.", exception);
            throw;
        }
    }

    protected override void OnLaunched(LaunchActivatedEventArgs args)
    {
        _ = args;
        StartupDiagnostics.Write("OnLaunched entered.");

        try
        {
            _window = new MainWindow();
            StartupDiagnostics.Write("MainWindow constructed.");
            _window.Activate();
            StartupDiagnostics.Write("MainWindow activated.");
        }
        catch (Exception exception)
        {
            StartupDiagnostics.WriteException("MainWindow startup failed.", exception);
            throw;
        }
    }

    private static void OnUnhandledException(object sender, Microsoft.UI.Xaml.UnhandledExceptionEventArgs args)
    {
        _ = sender;
        StartupDiagnostics.WriteException("Unhandled WinUI exception.", args.Exception);
    }
}

internal static class StartupDiagnostics
{
    private static readonly object Sync = new();

    private static string LogPath
        => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "OutlookAligner",
            "startup.log");

    internal static void Reset()
    {
        try
        {
            var directory = Path.GetDirectoryName(LogPath)!;
            Directory.CreateDirectory(directory);
            File.WriteAllText(LogPath, string.Empty);
        }
        catch
        {
            // Startup diagnostics must never become a startup dependency themselves.
        }
    }

    internal static void Write(string message)
    {
        try
        {
            lock (Sync)
            {
                var directory = Path.GetDirectoryName(LogPath)!;
                Directory.CreateDirectory(directory);
                File.AppendAllText(
                    LogPath,
                    $"{DateTimeOffset.Now:O} {message}{Environment.NewLine}");
            }
        }
        catch
        {
            // Best-effort diagnostics only.
        }
    }

    internal static void WriteException(string message, Exception? exception)
    {
        Write(message);
        if (exception is not null)
        {
            Write(exception.ToString());
        }
    }
}
