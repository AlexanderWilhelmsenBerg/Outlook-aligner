using System.Globalization;
using System.Text.Json;

namespace OutlookAligner.App.Diagnostics;

public enum AppLogLevel
{
    Debug,
    Info,
    Warning,
    Error,
}

public sealed class AppLogEntry
{
    public AppLogEntry()
    {
    }

    public AppLogEntry(
        DateTimeOffset timestamp,
        AppLogLevel level,
        string source,
        string message,
        string? eventContext,
        string? details)
    {
        Timestamp = timestamp;
        Level = level;
        Source = source;
        Message = message;
        EventContext = eventContext;
        Details = details;
    }

    public DateTimeOffset Timestamp { get; set; }

    public AppLogLevel Level { get; set; }

    public string Source { get; set; } = string.Empty;

    public string Message { get; set; } = string.Empty;

    public string? EventContext { get; set; }

    public string? Details { get; set; }

    public string TimestampDisplay => Timestamp.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.CurrentCulture);

    public string LevelDisplay => Level.ToString().ToUpperInvariant();
}

public sealed class LogLevelChoice
{
    public LogLevelChoice()
    {
    }

    public LogLevelChoice(string name, AppLogLevel? level)
    {
        Name = name;
        Level = level;
    }

    public string Name { get; set; } = string.Empty;

    public AppLogLevel? Level { get; set; }
}

internal sealed class AppEventLog
{
    private const int MaximumInMemoryEntries = 1000;
    private const long MaximumLogFileBytes = 5 * 1024 * 1024;

    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web);
    private static readonly Lazy<AppEventLog> CurrentInstance = new(() => new AppEventLog());

    private readonly object _gate = new();
    private readonly List<AppLogEntry> _entries = [];

    internal AppEventLog(string? filePathOverride = null)
    {
        FilePath = filePathOverride ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "OutlookAligner",
            "events.jsonl");

        var directory = Path.GetDirectoryName(FilePath);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            try
            {
                Directory.CreateDirectory(directory);
            }
            catch (IOException)
            {
            }
            catch (UnauthorizedAccessException)
            {
            }
        }

        LoadExistingEntries();
    }

    internal static AppEventLog Current => CurrentInstance.Value;

    internal string FilePath { get; }

    internal event Action<AppLogEntry>? EntryAdded;

    internal IReadOnlyList<AppLogEntry> Snapshot()
    {
        lock (_gate)
        {
            return _entries.ToArray();
        }
    }

    internal void Debug(string source, string message, string? eventContext = null, string? details = null)
        => Write(AppLogLevel.Debug, source, message, eventContext, details);

    internal void Info(string source, string message, string? eventContext = null, string? details = null)
        => Write(AppLogLevel.Info, source, message, eventContext, details);

    internal void Warning(string source, string message, string? eventContext = null, string? details = null)
        => Write(AppLogLevel.Warning, source, message, eventContext, details);

    internal void Error(string source, string message, string? eventContext = null, string? details = null)
        => Write(AppLogLevel.Error, source, message, eventContext, details);

    private void Write(
        AppLogLevel level,
        string source,
        string message,
        string? eventContext,
        string? details)
    {
        var entry = new AppLogEntry(
            DateTimeOffset.Now,
            level,
            source,
            message,
            Normalize(eventContext),
            Normalize(details));

        lock (_gate)
        {
            _entries.Add(entry);
            if (_entries.Count > MaximumInMemoryEntries)
            {
                _entries.RemoveRange(0, _entries.Count - MaximumInMemoryEntries);
            }

            TryPersist(entry);
        }

        EntryAdded?.Invoke(entry);
    }

    private void TryPersist(AppLogEntry entry)
    {
        try
        {
            var directory = Path.GetDirectoryName(FilePath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            TryRollLogFile();
            var json = JsonSerializer.Serialize(entry, JsonOptions);
            File.AppendAllText(FilePath, json + Environment.NewLine);
        }
        catch (IOException)
        {
            // In-memory diagnostics remain available if persistence temporarily fails.
        }
        catch (UnauthorizedAccessException)
        {
            // In-memory diagnostics remain available if persistence is unavailable.
        }
    }

    private void LoadExistingEntries()
    {
        if (!File.Exists(FilePath))
        {
            return;
        }

        try
        {
            foreach (var line in File.ReadLines(FilePath).TakeLast(MaximumInMemoryEntries))
            {
                if (string.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                try
                {
                    var entry = JsonSerializer.Deserialize<AppLogEntry>(line, JsonOptions);
                    if (entry is not null)
                    {
                        _entries.Add(entry);
                    }
                }
                catch (JsonException)
                {
                    // Ignore one malformed historical line so diagnostics remain usable.
                }
            }
        }
        catch (IOException)
        {
            // The in-app log can still start fresh if an old file is temporarily unreadable.
        }
        catch (UnauthorizedAccessException)
        {
            // The in-app log can still start fresh if the old file cannot be read.
        }
    }

    private void TryRollLogFile()
    {
        try
        {
            if (!File.Exists(FilePath) || new FileInfo(FilePath).Length < MaximumLogFileBytes)
            {
                return;
            }

            var previousPath = FilePath + ".previous";
            File.Move(FilePath, previousPath, overwrite: true);
        }
        catch (IOException)
        {
            // Logging should never break the application because rotation failed.
        }
        catch (UnauthorizedAccessException)
        {
            // Logging should never break the application because rotation failed.
        }
    }

    private static string? Normalize(string? value)
        => string.IsNullOrWhiteSpace(value) ? null : value.Trim();
}
