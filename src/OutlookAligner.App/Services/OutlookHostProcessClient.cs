using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using OutlookAligner.App.Diagnostics;
using OutlookAligner.Outlook.Contracts;

namespace OutlookAligner.App.Services;

internal sealed class OutlookHostProcessClient
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
    };

    private readonly AppEventLog _eventLog;
    private readonly string? _hostPathOverride;

    internal OutlookHostProcessClient(AppEventLog? eventLog = null, string? hostPathOverride = null)
    {
        _eventLog = eventLog ?? AppEventLog.Current;
        _hostPathOverride = hostPathOverride;
    }

    internal async Task<OutlookProbeResult> ScanAsync(int days, CancellationToken cancellationToken = default)
    {
        var result = await RunAsync(
            ["--days", days.ToString(CultureInfo.InvariantCulture), "--json", "--include-details"],
            "Calendar scan",
            cancellationToken).ConfigureAwait(false);

        result.ThrowIfFailed("Outlook calendar scan");

        var probe = JsonSerializer.Deserialize<OutlookProbeResult>(result.StandardOutput, JsonOptions)
            ?? throw new InvalidOperationException("OutlookHost returned an empty calendar result.");

        if (probe.ProtocolVersion != OutlookHostProtocol.Version)
        {
            throw new InvalidOperationException(
                $"OutlookHost protocol {probe.ProtocolVersion} is incompatible with UI protocol {OutlookHostProtocol.Version}.");
        }

        return probe;
    }

    internal Task<HostCommandResult> ProbeForwardAsync(
        string sourceSmtp,
        string globalAppointmentId,
        string entryId,
        CancellationToken cancellationToken = default)
        => RunAsync(
            [
                "--forward-spike",
                "--source-smtp", sourceSmtp,
                "--global-id", globalAppointmentId,
                "--probe-calendar-command",
                "--entry-id", entryId,
                "--include-details",
            ],
            "Forward capability probe",
            cancellationToken);

    internal Task<HostCommandResult> PrepareForwardAsync(
        string sourceSmtp,
        string globalAppointmentId,
        string entryId,
        string recipient,
        CancellationToken cancellationToken = default)
        => RunAsync(
            [
                "--forward-spike",
                "--source-smtp", sourceSmtp,
                "--global-id", globalAppointmentId,
                "--prepare-calendar-recipient",
                "--entry-id", entryId,
                "--to", recipient,
                "--include-details",
            ],
            "Forward prepare/discard",
            cancellationToken);

    private async Task<HostCommandResult> RunAsync(
        IReadOnlyList<string> arguments,
        string operation,
        CancellationToken cancellationToken)
    {
        var hostPath = ResolveHostPath();
        var startInfo = new ProcessStartInfo
        {
            FileName = hostPath,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
        };

        foreach (var argument in arguments)
        {
            startInfo.ArgumentList.Add(argument);
        }

        _eventLog.Debug(
            "OutlookHost",
            $"Starting {operation}.",
            details: $"Executable: {Path.GetFileName(hostPath)}; options: {DescribeOptions(arguments)}");

        var stopwatch = Stopwatch.StartNew();
        using var process = new Process { StartInfo = startInfo };
        if (!process.Start())
        {
            _eventLog.Error("OutlookHost", $"Could not start {operation}.");
            throw new InvalidOperationException("OutlookHost could not be started.");
        }

        var standardOutputTask = process.StandardOutput.ReadToEndAsync(cancellationToken);
        var standardErrorTask = process.StandardError.ReadToEndAsync(cancellationToken);

        try
        {
            await process.WaitForExitAsync(cancellationToken).ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            TryKill(process);
            _eventLog.Warning(
                "OutlookHost",
                $"{operation} was canceled after {stopwatch.ElapsedMilliseconds} ms.");
            throw;
        }

        var standardOutput = await standardOutputTask.ConfigureAwait(false);
        var standardError = await standardErrorTask.ConfigureAwait(false);
        stopwatch.Stop();

        var result = new HostCommandResult(process.ExitCode, standardOutput.Trim(), standardError.Trim());
        var details = $"ExitCode={result.ExitCode}; ElapsedMs={stopwatch.ElapsedMilliseconds}"
            + (string.IsNullOrWhiteSpace(result.DiagnosticText)
                ? string.Empty
                : Environment.NewLine + result.DiagnosticText);

        if (result.Success)
        {
            _eventLog.Debug("OutlookHost", $"{operation} completed successfully.", details: details);
        }
        else
        {
            _eventLog.Warning("OutlookHost", $"{operation} returned exit code {result.ExitCode}.", details: details);
        }

        return result;
    }

    private string ResolveHostPath()
    {
        if (!string.IsNullOrWhiteSpace(_hostPathOverride) && File.Exists(_hostPathOverride))
        {
            return _hostPathOverride;
        }

        var environmentPath = Environment.GetEnvironmentVariable("OUTLOOK_ALIGNER_HOST_PATH");
        if (!string.IsNullOrWhiteSpace(environmentPath) && File.Exists(environmentPath))
        {
            return environmentPath;
        }

        var siblingPath = Path.Combine(AppContext.BaseDirectory, "OutlookAligner.OutlookHost.exe");
        if (File.Exists(siblingPath))
        {
            return siblingPath;
        }

        throw new FileNotFoundException(
            "OutlookAligner.OutlookHost.exe was not found beside the UI. Set OUTLOOK_ALIGNER_HOST_PATH for development builds.",
            siblingPath);
    }

    private static string DescribeOptions(IReadOnlyList<string> arguments)
        => string.Join(
            ' ',
            arguments.Where(argument => argument.StartsWith("--", StringComparison.Ordinal)));

    private static void TryKill(Process process)
    {
        try
        {
            if (!process.HasExited)
            {
                process.Kill(entireProcessTree: true);
            }
        }
        catch (InvalidOperationException)
        {
        }
    }
}

internal sealed record HostCommandResult(int ExitCode, string StandardOutput, string StandardError)
{
    internal bool Success => ExitCode == 0;

    internal string DiagnosticText
        => string.IsNullOrWhiteSpace(StandardError)
            ? StandardOutput
            : string.IsNullOrWhiteSpace(StandardOutput)
                ? StandardError
                : StandardOutput + Environment.NewLine + StandardError;

    internal string? ResultSummary
        => DiagnosticText
            .Split(["\r\n", "\n"], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .FirstOrDefault(line => line.StartsWith("Result: ", StringComparison.Ordinal))?["Result: ".Length..];

    internal void ThrowIfFailed(string operation)
    {
        if (Success)
        {
            return;
        }

        var detail = string.IsNullOrWhiteSpace(DiagnosticText)
            ? $"OutlookHost exit code {ExitCode}."
            : DiagnosticText;
        throw new InvalidOperationException($"{operation} failed: {detail}");
    }
}
